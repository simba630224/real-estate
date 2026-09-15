import os
import io
import json
import zipfile
import datetime
import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ==============================================================================
# 1. 內政部實價登錄 Open Data (下載最新一期壓縮檔)
# ==============================================================================
def scrape_moi_real_estate():
    print("開始下載內政部實價登錄資料...")
    url = "https://plvr.land.moi.gov.tw/DownloadSeason?season=current&type=zip&fileName=lvr_landcsv.zip"
    
    # 加上完整的 Headers 模擬真實瀏覽器，避免被擋
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    }
    
    try:
        # 使用 requests 並設定 timeout
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status() # 若不是 HTTP 200 OK，會拋出例外
        
        # 檢查下載的檔案內容類型是否為 zip
        if 'application/zip' not in response.headers.get('Content-Type', '') and 'text/html' in response.headers.get('Content-Type', ''):
             print("警告：下載的內容似乎是 HTML 而非 ZIP 壓縮檔。內政部可能暫時阻擋存取。")
             return pd.DataFrame()
             
        zip_data = response.content
        
    except Exception as e:
        print(f"內政部下載連線失敗: {e}")
        return pd.DataFrame()

    try:
        with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
            with z.open('H_lvr_land_A.csv') as f:
                # 略過第一行的英文欄位名
                df = pd.read_csv(f, encoding='utf-8', skiprows=[1])
    except zipfile.BadZipFile:
        print("錯誤：下載成功，但檔案不是有效的 ZIP 格式 (BadZipFile)。可能抓到錯誤網頁。")
        return pd.DataFrame()
    except Exception as e:
        print(f"ZIP 解壓縮或讀取 CSV 失敗: {e}")
        return pd.DataFrame()

    return df

# ==============================================================================
# 2. 核心商業邏輯 (維持不變)
# ==============================================================================
def process_real_estate_data(df):
    if df.empty:
        return pd.DataFrame()
    
    cols_to_keep = ['鄉鎮市區', '土地區段位置建物門牌', '交易年月日', '建築完成年月', '建物格局-房', '單價元平方公尺']
    cols_exist = [c for c in cols_to_keep if c in df.columns]
    df = df[cols_exist].copy()
    
    df.dropna(subset=['土地區段位置建物門牌', '單價元平方公尺', '建物格局-房'], inplace=True)
    df['建物格局-房'] = pd.to_numeric(df['建物格局-房'], errors='coerce')
    df = df[df['建物格局-房'].isin([2, 3])]
    
    def assign_mrt_station(address):
        addr = str(address)
        if any(kw in addr for kw in ['青昇', '橫山', '大成路', '大智路', '高鐵北路二段', '領航北路四段']):
            return 'A17 領航站'
        elif any(kw in addr for kw in ['青平', '青溪', '高鐵南路一段', '高鐵北路一段', '青埔路', '領航南路三段', '領航北路二段', '領航北路三段']):
            return 'A18 高鐵桃園站'
        elif any(kw in addr for kw in ['青峰', '青芝', '洽溪', '高鐵南路二段', '領航南路一段', '領航南路二段', '文德路', '文智路']):
            return 'A19 桃園體育園區站'
        return '其他'

    df['捷運站點'] = df['土地區段位置建物門牌'].apply(assign_mrt_station)
    df = df[df['捷運站點'] != '其他']
    
    current_minguo_year = datetime.datetime.now().year - 1911
    
    def calculate_age_group(build_date):
        try:
            build_year = int(str(build_date)[:-4])
            age = current_minguo_year - build_year
            
            if age <= 2: return '2年內'
            elif 2 < age <= 5: return '2-5年'
            elif 5 < age <= 10: return '5-10年'
            else: return '超過10年'
        except:
            return '未知'
            
    df['房齡級距'] = df['建築完成年月'].apply(calculate_age_group)
    df = df[df['房齡級距'].isin(['2年內', '2-5年', '5-10年'])]
    
    df['單價元平方公尺'] = pd.to_numeric(df['單價元平方公尺'], errors='coerce')
    df['單價(萬/坪)'] = (df['單價元平方公尺'] * 3.305785) / 10000
    
    summary = df.groupby(['捷運站點', '房齡級距'])['單價(萬/坪)'].agg(['max', 'min', 'mean']).reset_index()
    summary.columns = ['捷運站點', '房齡級距', '每坪最高價(萬)', '每坪最低價(萬)', '每坪平均價(萬)']
    
    summary['每坪最高價(萬)'] = summary['每坪最高價(萬)'].round(2)
    summary['每坪最低價(萬)'] = summary['每坪最低價(萬)'].round(2)
    summary['每坪平均價(萬)'] = summary['每坪平均價(萬)'].round(2)
    
    return summary

# ==============================================================================
# 3. Google Sheets 寫入邏輯 (已修復 update 語法)
# ==============================================================================
def update_google_sheet(site_name, summary_df, spreadsheet_url):
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        print("未設定 GOOGLE_CREDENTIALS，本地測試結果如下：")
        print(summary_df)
        return

    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    sheet = client.open_by_url(spreadsheet_url)
    
    try:
        worksheet = sheet.worksheet(site_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title=site_name, rows="100", cols="20")
    
    worksheet.clear()
    update_time_str = f"更新日期: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    
    if not summary_df.empty:
        header = summary_df.columns.values.tolist()
        data_rows = summary_df.values.tolist()
        write_data = [[update_time_str], [], header] + data_rows
    else:
        write_data = [[update_time_str], [], ["本期無符合條件之交易資料或下載失敗"]]
        
    # [修復] 使用具名參數 values=與 range_name= 消除 DeprecationWarning
    worksheet.update(values=write_data, range_name='A1')
    print(f"✅ [{site_name}] 資料已更新至 Google Sheet。")


if __name__ == "__main__":
    TARGET_SHEET_URL = "https://docs.google.com/spreadsheets/d/1PBRTc9i1AYTzXBNXxXL9u0oZixlNqxA8Htsu0gUJO18/edit"
    
    print("--- 開始執行內政部實價登錄資料爬取與分析 ---")
    raw_df = scrape_moi_real_estate()
    summary_df = process_real_estate_data(raw_df)
    update_google_sheet("內政部實價登錄", summary_df, TARGET_SHEET_URL)
