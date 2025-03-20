# EDP 檢查 - SKU_Check - by sku group

## Use step
1. 確認 config.yaml ：
- urls: 要搜索的 URL 清單
- filename: the source file
- mode: 
-- True:  the sku id in the source data that matches --> Output: found_skus_report.xlsx
-- False: does not match the sku id from the url --> OUtput: not_found_skus_report.xlsx

## 執行
``` bash
pip install -r requirements.txt
python main.py
```