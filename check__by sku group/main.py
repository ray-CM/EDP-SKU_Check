import requests
import re
import yaml
import pandas as pd
import os
from urllib.parse import urlparse

def find_patterns(text):
    #Finds all matches of the pattern ????-????-???? in the text.- looking for SKU IDs within > and <
    pattern1 = r">.{4}-.{4}-.{4}<"
    matches1 = re.findall(pattern1, text)
    matches1 = [match.split('<')[0].split('>')[1].strip() for match in matches1]
    
    # Alternative pattern - directly looking for SKU ID format
    pattern2 = r"[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}"
    matches2 = re.findall(pattern2, text)
    matches2 = [match.strip() for match in matches2]
    
    # Combine results, remove duplicates by converting to set then back to list
    all_matches = list(set(matches1 + matches2))
    
    return all_matches

def make_request(url, params=None, headers=None):
    # Makes a GET request to the specified URL.
    try:
        response = requests.get(url, params=params, headers=headers, timeout=10)
        response.raise_for_status()  # Raise an error for HTTP errors
        return response
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        return None

def get_sku_group(url):
    #Extracts the last segment of the URL path.

    parsed_url = urlparse(url)
    path_segments = parsed_url.path.strip('/').split('/')
    return path_segments[-1] if path_segments else ""

def ensure_output_directory(directory):
    #Ensures that the output directory exists, creates it if it doesn't.

    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"Created directory: {directory}")
    return directory

def export_to_excel(data_dict, filename="gcp_sku_report.xlsx", output_dir="SKU_Report"):
    #Exports the SKU data to an Excel file in the specified output directory.

    ensure_output_directory(output_dir)
    
    # full file path
    file_path = os.path.join(output_dir, filename)
    
    # Create a dataframe from the dictionary
    rows = []
    for sku_id, info in data_dict.items():
        row = {
            "Service ID": sku_id,
            "SKU description": info[0],
            "Cost ($)": info[1]
        }
        # Add SKU Group if available
        if len(info) > 2:
            row["SKU Group"] = info[2]
            
        rows.append(row)
    
    df = pd.DataFrame(rows)
    
    # When mode is True, sort by SKU Group so similar groups are together
    if "SKU Group" in df.columns and df["SKU Group"].notna().any():
        df = df.sort_values(by="SKU Group")
    
    # Export to Excel
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='SKU Report')
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['SKU Report']
    
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })
    
    # Format the currency column
    money_format = workbook.add_format({'num_format': '#,##0.00'})
    
    # column headers 
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    worksheet.set_column('C:C', 15, money_format)
    
    # Auto-adjust columns' width
    for i, col in enumerate(df.columns):
        column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, column_width)
    
    writer.close()
    
    print(f"Report saved as {file_path}")
    print(f"Finished! \n")

with open('config.yaml', 'r') as file:
    data = yaml.safe_load(file)

urls = data["config"]["urls"]

# 檢查配置中的文件格式
filename = data["config"]["filename"]
if filename.endswith('.csv'):
    # 如果文件名是 CSV，提示需要修改 config.yaml
    print(f"Warning: filename in config.yaml is still set to CSV format: {filename}")
    print("Consider updating config.yaml to use an Excel file instead.")
    
    # 讀取 CSV 
    df = pd.read_csv(filename)
    df = df.dropna()
    
    # 轉換為 Excel 
    excel_filename = filename.replace('.csv', '.xlsx')
    print(f"Converting CSV to Excel format: {excel_filename}")
    df.to_excel(excel_filename, index=False)
    print("Excel file created. Please update config.yaml to use this file instead.")
else:
    # 讀取 Excel 資料
    df = pd.read_excel(filename)
    df = df.dropna()

# 確保 SKU ID 的格式一致
if 'SKU ID' in df.columns:
    df['SKU ID'] = df['SKU ID'].astype(str).str.strip()

sku_dict = df.set_index('SKU ID')[['SKU description', 'Subtotal ($)']].apply(list, axis=1).to_dict()

# 印出 SKU ID 總數
print(f"\nTotal SKUs in original file: {len(sku_dict)}")

sku_to_group = {}

for url in urls:
    sku_group = get_sku_group(url)
    print(f" \nProcessing URL: {url} (SKU Group: {sku_group})")
    response = make_request(url)
    if response:
        matches = find_patterns(response.text)
        print(f"  Found {len(matches)} SKUs in {sku_group} ")
        for match in matches:
            # If SKU already exists in another group, print a warning
            if match in sku_to_group and sku_to_group[match] != sku_group:
                print(f"  Warning: SKU {match} found in both {sku_to_group[match]} and {sku_group}")
            sku_to_group[match] = sku_group
    else:
        print(f"  Error: Could not retrieve data from {url}")

ansT = {}
ansF = {}

# 所有從網頁找到的 SKU
print(f"\nMatching SKUs from webpages with input file...")
for sku_id in sku_to_group.keys():
    # 找到的 SKU ID 標準化
    normalized_sku = sku_id.strip()
    
    # 檢查是否在輸入檔案中找到此 SKU
    found_in_input = False
    matching_key = None
    
    for input_key in sku_dict.keys():
        # 輸入檔案中的 SKU ID 標準化
        normalized_input = input_key.strip()
        
        # 不區分大小寫
        if normalized_sku.upper() == normalized_input.upper():
            found_in_input = True
            matching_key = input_key
            break
    
    if found_in_input:
        # 如果找到匹配，加入 ansT
        ansT[matching_key] = sku_dict[matching_key] + [sku_to_group[sku_id]]

# 在輸入檔案中但不在網頁上的 SKU
for query in sku_dict.keys():
    normalized_query = query.strip().upper()
    
    if query in ansT:
        continue
    
    found = False
    for web_sku in sku_to_group.keys():
        if web_sku.strip().upper() == normalized_query:
            ansT[query] = sku_dict[query] + [sku_to_group[web_sku]]
            found = True
            break

    if not found:
        ansF[query] = sku_dict[query]

# to folder
output_directory = "SKU_Report"

if data["config"]["mode"]:
    print("True:")
    print(f"Total SKUs found: {len(ansT)} \n")
    export_to_excel(ansT, "found_skus_report.xlsx", output_directory)
else:
    print("\nFalse:")
    print(f"Total SKUs not found: {len(ansF)} \n ")
    export_to_excel(ansF, "not_found_skus_report.xlsx", output_directory)