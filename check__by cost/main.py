import requests
import re
import yaml
import pandas as pd
import os

def find_patterns(text):
    #Finds all matches of the pattern ????-????-???? in the text.

    pattern = r">.{4}-.{4}-.{4}<"
    matches = re.findall(pattern, text)
    matches = [match.split('<')[0].split('>')[1] for match in matches]
    return matches

def make_request(url, params=None, headers=None):
    #Makes a GET request to the specified URL.
    
    try:
        response = requests.get(url, params=params, headers=headers, timeout=10)
        response.raise_for_status()  # Raise an error for HTTP errors
        return response
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        return None

def export_to_excel(data_dict, filename="gcp_sku_report.xlsx"):
    # Exports the SKU data to an Excel file in the 'SKU_Report' folder.
    # Create 'SKU_Report' directory if it doesn't exist
    
    report_dir = "SKU_Report"
    if not os.path.exists(report_dir):
        os.makedirs(report_dir)
        print(f"Created directory: {report_dir}")
    
    # Create the complete file path
    file_path = os.path.join(report_dir, filename)
    
    # Create a dataframe from the dictionary
    rows = []
    for sku_id, info in data_dict.items():
        rows.append({
            "Service ID": sku_id,
            "SKU description": info[0],
            "Cost ($)": info[1]
        })
    
    df = pd.DataFrame(rows)
    
    # Sort by Cost in descending order
    df = df.sort_values(by="Cost ($)", ascending=False)
    
    try:
        # Try using openpyxl engine
        writer = pd.ExcelWriter(file_path, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='SKU Report')
        
        # Auto-adjust columns' width
        worksheet = writer.sheets['SKU Report']
        for idx, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            ) + 2
            # openpyxl column widths are measured in characters
            worksheet.column_dimensions[chr(65 + idx)].width = max_len
        
        # Close the Pandas Excel writer and output the Excel file
        writer.close()
        print(f"\n Excel report saved as {file_path}")
        
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        print("Attempting to save with basic Excel export...")
        df.to_excel(file_path, index=False)
        print(f"Basic Excel report saved as {file_path}")

# Read Excel file
def read_excel_file(filename):
    """
    Reads an Excel file.
    
    :param filename: Excel filename to read
    :return: Pandas DataFrame
    """
    try:
        print(f"Reading Excel file {filename}...")
        return pd.read_excel(filename)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        raise ValueError(f"Could not read the Excel file {filename}: {e}")

with open('config.yaml', 'r') as file:
    data = yaml.safe_load(file)

urls = data["config"]["urls"]
excel_filename = data["config"]["filename"]

# Check if the file exists
if not os.path.isfile(excel_filename):
    print(f"Error: Excel file '{excel_filename}' not found.")
else:
    try:
        # Read the Excel file
        df = read_excel_file(excel_filename)
        df = df.dropna()
        
        # Check if required columns exist
        required_columns = ['SKU ID', 'SKU description', 'Subtotal ($)']
        for col in required_columns:
            if col not in df.columns:
                print(f"Error: Column '{col}' not found in Excel file. Available columns: {df.columns.tolist()}")
                raise ValueError(f"Required column '{col}' not found in Excel file")
        
        sku_dict = df.set_index('SKU ID')[['SKU description', 'Subtotal ($)']].apply(list, axis=1).to_dict()

        # Store matches for each URL separately to track duplicates
        url_matches = {}
        all_matches = set()
        
        for url in urls:
            response = make_request(url).text
            matches = find_patterns(response)
            url_group = url.split('/')[-1]  # Extract the group name from URL
            
            print(f"\n Processing URL: {url} (SKU Group: {url_group})")
            url_matches[url_group] = set(matches)
            print(f"  Found {len(url_matches[url_group])} SKUs in {url_group}")
            
            # Check for duplicates with previously processed URLs
            for prev_group, prev_matches in url_matches.items():
                if prev_group != url_group:
                    duplicates = url_matches[url_group].intersection(prev_matches)
                    if duplicates:
                        for dup in duplicates:
                            print(f"Warning: SKU {dup} found in both {prev_group} and {url_group}")
            
            # Add to all matches
            all_matches.update(matches)
        
        # Create dictionaries for found and not found SKUs
        ansT = {}
        ansF = {}
        for query in sku_dict.keys():
            if query in all_matches:
                ansT[query] = sku_dict[query]
            else:
                ansF[query] = sku_dict[query]

        # Export based on mode
        if data["config"]["mode"]:
            export_to_excel(ansT, "found_skus_report.xlsx")
        else:
            export_to_excel(ansF, "not_found_skus_report.xlsx")
    except ValueError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")