import os
import io
import json
import zipfile
import datetime
import urllib.request
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ==============================================================================
# 1. 內政部實價登錄 Open Data (下載最新一期壓縮檔)
# ==============================================================================
def scrape_moi_real_estate():
    """下載內政部實價登錄 Open Data (當期)，並解析桃園市買賣資料"""
    print("開始下載內政部實價登錄資料...")
    url = "https://plvr.land.moi.gov.tw/DownloadSeason?season=current&type=zip&fileName=lvr_landcsv.zip"
    
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        response = urllib.request.urlopen(req)
        zip_data = response.read()
    except Exception as e:
        print(f"內政部下載失敗: {e}")
        return pd.DataFrame()

    try:
        with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
            # H_lvr_land_A.csv 代表: H(桃園市)_lvr_land(實價登錄)_A(買賣)
            with z.open('H_lvr_land_A.csv') as f:
                # 略過第一行的英文欄位名，直接讀取第二行的中文欄位
                df = pd.read_csv(f, encoding='utf-8', skiprows=[1])
    except Exception as e:
        print(f"ZIP 解壓縮或讀取失敗: {e}")
        return pd.DataFrame()

    return df

# ==============================================================================
# 2. 核心商業邏輯 (A17/A18/A19 分類、屋齡計算、房數篩選、價格計算)
# ==============================================================================
def process_real_estate_data(df):
    if df.empty:
        return pd.DataFrame()
    
    # 1. 保留必要欄位
    cols_to_keep = ['鄉鎮市區', '土地區段位置建物門牌', '交易年月日', '建築完成年月', '建物格局-房', '單價元平方公尺']
    cols_exist = [c for c in cols_to_keep if c in df.columns]
    df = df[cols_exist].copy()
    
    # 清除關鍵欄位的空值
    df.dropna(subset=['土地區段位置建物門牌', '單價元平方公尺', '建物格局-房'], inplace=True)
    
    # 2. 條件篩選：僅保留 2~3 房
    df['建物格局-房'] = pd.to_numeric(df['建物格局-房'], errors='coerce')
    df = df[df['建物格局-房'].isin([2, 3])]
    
    # 3. 站點分類 (利用門牌、路段關鍵字判斷青埔三大站周邊)
    def assign_mrt_station(address):
        addr = str(address)
        # A17 領航站 (大園區為主)
        if any(kw in addr for kw in ['青昇', '橫山', '大成路', '大智路', '高鐵北路二段', '領航北路四段']):
            return 'A17 領航站'
        # A18 高鐵桃園站 (青埔核心)
        elif any(kw in addr for kw in ['青平', '青溪', '高鐵南路一段', '高鐵北路一段', '青埔路', '領航南路三段', '領航北路二段', '領航北路三段']):
            return 'A18 高鐵桃園站'
        # A19 桃園體育園區站 (中壢區為主)
        elif any(kw in addr for kw in ['青峰', '青芝', '洽溪', '高鐵南路二段', '領航南路一段', '領航南路二段', '文德路', '文智路']):
            return 'A19 桃園體育園區站'
        return '其他'

    df['捷運站點'] = df['土地區段位置建物門牌'].apply(assign_mrt_station)
    df = df[df['捷運站點'] != '其他'] # 剔除不屬於這三站周邊的交易
    
    # 4. 計算屋齡並區分級距 (建築完成年月至迄今)
    current_minguo_year = datetime.datetime.now().year - 1911 # 取得目前民國年 (如2026年為115)
    
    def calculate_age_group(build_date):
        try:
            # 建築完成年月 (例如: 1080101 -> 提取前三碼 108)
            build_year = int(str(build_date)[:-4])
            age = current_minguo_year - build_year
            
            if age <= 2: return '2年內'
            elif 2 < age <= 5: return '2-5年'
            elif 5 < age <= 10: return '5-10年'
            else: return '超過10年'
        except:
            # 若無建築完成日期，歸類為未知並排除
            return '未知'
            
    df['房齡級距'] = df['建築完成年月'].apply(calculate_age_group)
    df = df[df['房齡級距'].isin(['2年內', '2-5年', '5-10年'])] # 過濾掉超過10年或未知的物件
    
    # 5. 轉換單價 (原本為 元/平方公尺，轉為 萬/坪)
    # 1 平方公尺 = 0.3025 坪 => 1坪 = 3.305785 平方公尺
    df['單價元平方公尺'] = pd.to_numeric(df['單價元平方公尺'], errors='coerce')
    df['單價(萬/坪)'] = (df['單價元平方公尺'] * 3.305785) / 10000
    
    # 6. 分群計算最大值、最小值與平均值
    summary = df.groupby(['捷運站點', '房齡級距'])['單價(萬/坪)'].agg(['max', 'min', 'mean']).reset_index()
    summary.columns = ['捷運站點', '房齡級距', '每坪最高價(萬)', '每坪最低價(萬)', '每坪平均價(萬)']
    
    summary['每坪最高價(萬)'] = summary['每坪最高價(萬)'].round(2)
    summary['每坪最低價(萬)'] = summary['每坪最低價(萬)'].round(2)
    summary['每坪平均價(萬)'] = summary['每坪平均價(萬)'].round(2)
    
    return summary

# ==============================================================================
# 3. Google Sheets 寫入邏輯
# ==============================================================================
def update_google_sheet(site_name, summary_df, spreadsheet_url):
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        print("未設定 GOOGLE_CREDENTIALS 環境變數。本地測試預覽如下：")
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
        write_data = [[update_time_str], [], ["本期無符合條件之交易資料"]]
        
    worksheet.update('A1', write_data)
    print(f"✅ [{site_name}] 資料已更新至 Google Sheet。")

# ==============================================================================
# 主程式執行區
# ==============================================================================
if __name__ == "__main__":
    # 目標 Google Sheet 網址
    TARGET_SHEET_URL = "https://docs.google.com/spreadsheets/d/1ffi9H6GdzzlH0p0-06oAsm5_D-e_tVrZYMzrnRGMDBQ/edit?gid=0#gid=0"
    
    print("--- 開始執行內政部實價登錄資料爬取與分析 ---")
    
    # 1. 下載並取得原始資料
    raw_df = scrape_moi_real_estate()
    
    # 2. 資料清洗與運算
    summary_df = process_real_estate_data(raw_df)
    
    # 3. 寫入 Google Sheet (工作表名稱設為 "內政部實價登錄")
    update_google_sheet("內政部實價登錄", summary_df, TARGET_SHEET_URL)
