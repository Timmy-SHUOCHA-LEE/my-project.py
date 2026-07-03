import pandas as pd
import requests
from FinMind.data import DataLoader
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
import os  # 用於路徑檢查

# 定義 API 的 URL 和 Headers
url = 'https://www.fundclear.com.tw/api/etf/product/query'
headers = {
    'accept': 'application/json',
    'content-type': 'application/json',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36'
}

# 定義 API 請求資料
data = {
    "_pageSize": 500,
    "_pageNum": 1,
    "column": "",
    "asc": True,
    "searchName": "",
    "etfType": "",
    "listingDateEnd": "",
    "listingDateStart": "",
    "maxAmount": "10000",
    "maxBeneficiary": "",
    "maxClosingPrice": "1000",
    "minAmount": "0",
    "minBeneficiary": "",
    "minClosingPrice": "0",
    "etfCate": [
        "國內成分股ETF",
        "國外成分股ETF",
        "債券及固定收益ETF",
        "原型期貨ETF",
        "槓桿型及反向型ETF",
        "槓桿型及反向型期貨ETF"
    ]
}

# 發送 POST 請求
response = requests.post(url, headers=headers, json=data)

# 檢查是否請求成功
if response.status_code == 200:
    print("API 請求成功")
    response_data = response.json()
    
    # 提取 ETF 資料
    if 'list' in response_data:
        etf_list = response_data['list']
        
        if len(etf_list) > 0:
            print("成功提取 ETF 資料")
            
            # 讀取 Excel 檔案
            excel_path = "C:\\Users\\asus\\Desktop\\定期定額清單 (12).xlsx"
            if not os.path.exists(excel_path):
                print("找不到指定的 Excel 檔案，請檢查檔案路徑！")
            else:
                try:
                    df = pd.read_excel(excel_path, sheet_name="工作表1")
                    print("成功讀取 Excel 檔案")
                    
                    # 確認是否有 "代號" 欄位
                    if '代號' in df.columns:
                        print("開始篩選無法比對的項目")
                        unmatched_stock_numbers = []  # 存儲無法匹配的 stock_no
                        
                        # 迭代 API 回傳的資料，檢查是否能在 Excel 中找到匹配
                        for etf in etf_list:
                            stock_no = etf.get('stockNo', 'N/A')  # API 的 stockNo
                            name = etf.get('name', 'N/A')        # API 的 name
                            
                            # 檢查是否能匹配 Excel 中的 "代號" 欄位
                            matched_row = df[df['代號'] == stock_no]  # 篩選匹配的列
                            if matched_row.empty:  # 如果無匹配
                                unmatched_stock_numbers.append(stock_no)
                        
                        # 輸出結果
                        if unmatched_stock_numbers:
                            stock_symbols = unmatched_stock_numbers  # 更新 stock_symbols
                            output = ", ".join([f"'{stock_no}'" for stock_no in unmatched_stock_numbers])
                            print(f"無法匹配的股票代碼: {output}")
                        else:
                            stock_symbols = []  # 如果所有代碼匹配，設為空列表
                            print("所有股票代碼均能匹配")
                    else:
                        print("Excel 中未找到 '代號' 欄位，請確認欄位名稱！")
                        stock_symbols = []  # 防止變數未定義
                except Exception as e:
                    print(f"讀取 Excel 檔案時發生錯誤：{e}")
        else:
            print("ETF 列表為空，無數據可處理")
            stock_symbols = []  # ETF 資料為空時預設為空列表
    else:
        print("返回數據中未找到 'list'")
        stock_symbols = []  # 保險起見設為空列表
else:
    print(f"API 請求失敗，狀態碼: {response.status_code}")
    stock_symbols = []  # 確保程式不會中斷

# 初始化 FinMind API
api = DataLoader()
api.login_by_token(api_token="eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNS0wNi0xNiAwODo0NDoxNSIsInVzZXJfaWQiOiI4OTEwMDciLCJpcCI6IjQ5LjIxOC4yMDguNTYiLCJleHAiOjE3NTA2Mzk0NTV9.lUrXuHpdt48RnU1T1xOyKs8qnD91JUM7GHYRKHtarOg")  # 替換為有效的 token

# 計算日期範圍
current_date = datetime.now()
start_date = (current_date - relativedelta(months=1) + timedelta(days=1)).strftime("%Y-%m-%d")
end_date = current_date.strftime("%Y-%m-%d")

# 從 FinMind 查詢股票名稱函數
def get_stock_names():
    stock_info = api.taiwan_stock_info()
    stock_name_dict = {row["stock_id"]: row["stock_name"] for _, row in stock_info.iterrows()}
    return stock_name_dict

# 獲取股票數據函數
def get_stock_data(symbol):
    stock_data = api.taiwan_stock_daily(
        stock_id=symbol,
        start_date=start_date,
        end_date=end_date
    )
    return stock_data[['Trading_money', 'Trading_Volume']] if not stock_data.empty else pd.DataFrame()

# 計算指標
def calculate_metrics(symbol_list, stock_names):
    metrics = []
    for symbol in symbol_list:
        if any(char in symbol for char in ['C', 'R', 'L']):  # 跳過包含 'C', 'R', 'U' 的股票代碼
            print(f"Skipping {symbol} as it contains 'C', 'R' , or 'L'")
            continue
        data = get_stock_data(symbol)
        if not data.empty:
            total_money = data['Trading_money'].sum() / 1000
            total_volume = data['Trading_Volume'].sum() / 1000
            avg_volume = total_volume / len(data)
            avg_price = total_money / total_volume if total_volume > 0 else 0
            stock_name = stock_names.get(symbol, "N/A")
            metrics.append((symbol, stock_name, total_money, total_volume, avg_volume, avg_price))
    return metrics

# 打印指標
def print_metrics(metrics):
    print("Stock Metrics")
    print("-" * 30)
    for symbol, stock_name, total_money, total_volume, avg_volume, avg_price in metrics:
        print(f"股票代碼: {symbol}")
        print(f"股票名稱: {stock_name}")
        print(f"成交值 (千): {total_money}")
        print(f"成交量 (千): {total_volume}")
        print(f"日均量: {avg_volume}")
        print(f"均價: {avg_price}")
        print("-" * 30)

# 保存指標到指定路徑的 Excel
def save_metrics_to_excel(metrics, file_path):
    if not os.path.exists(os.path.dirname(file_path)):
        os.makedirs(os.path.dirname(file_path))  # 確保目錄存在
    wb = Workbook()
    ws = wb.active
    ws.title = "Metrics"
    headers = ["股票代碼", "股票名稱", "成交值 (千)", "成交量 (千)", "日均量", "均價"]
    ws.append(headers)
    for metric in metrics:
        ws.append(metric)
    wb.save(file_path)
    print(f"Metrics have been successfully saved to {file_path}")

# 執行邏輯
output_file_path = "C:\\python training\\ETF專案\\final data\\final data.xlsx"  # 替換為你的實際路徑
stock_names = get_stock_names()  # 從 FinMind 獲取股票名稱
metrics = calculate_metrics(stock_symbols, stock_names)
print_metrics(metrics)
save_metrics_to_excel(metrics, output_file_path)