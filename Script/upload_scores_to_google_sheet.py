# -*- coding: utf-8 -*-
"""
將「學生資料夾」內的「{名稱}_評分表.csv」直接寫入 Google 試算表對應的工作表。
需先完成一次性設定：服務帳號 JSON、試算表分享給服務帳號。見「Google_試算表_上傳設定說明.txt」
"""
import csv
import os
import sys

# 試算表 ID（從網址 .../d/【這裡】/edit 取得）
SPREADSHEET_ID = "1T5AhtfP1guJXVZQzi-lgBvJi-9GGrd3Y6H9or8Rq6-s"

try:
    from paths_config import BASE_DIR, iter_student_folders
except ImportError:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    def iter_student_folders():
        skip = {"Script", "Data", "Doc", "Gallery", "__pycache__"}
        for name in sorted(os.listdir(BASE_DIR)):
            path = os.path.join(BASE_DIR, name)
            if os.path.isdir(path) and name not in skip:
                yield name

# 服務帳號金鑰 JSON 路徑（可改為環境變數 GOOGLE_APPLICATION_CREDENTIALS 或此預設路徑）
CREDENTIALS_PATH = os.environ.get(
    "GOOGLE_APPLICATION_CREDENTIALS",
    os.path.join(os.path.dirname(__file__), "google_服務帳號金鑰.json"),
)


def load_csv_rows(csv_path):
    """用標準 csv 模組讀取，支援欄位內逗號與引號。"""
    with open(csv_path, "r", encoding="utf-8") as f:
        return list(csv.reader(f))


def upload_one(student_folder_name, csv_path=None, spreadsheet_id=None, credentials_path=None):
    """
    將單一學生的評分表 CSV 寫入 Google 試算表。
    - student_folder_name: 工作表名稱，例如「周昱翔_開心」
    - csv_path: CSV 檔路徑；若為 None，則用「學生作業目錄/student_folder_name/{名稱}_評分表.csv」
    - spreadsheet_id / credentials_path: 若為 None 則用上方常數
    """
    spreadsheet_id = spreadsheet_id or SPREADSHEET_ID
    credentials_path = credentials_path or CREDENTIALS_PATH

    if csv_path is None:
        csv_path = os.path.join(BASE_DIR, student_folder_name, f"{student_folder_name}_評分表.csv")
    csv_path = os.path.normpath(csv_path)

    if not os.path.isfile(csv_path):
        raise FileNotFoundError(f"找不到評分表: {csv_path}")

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("請先安裝： pip install gspread google-auth")
        sys.exit(1)

    if not os.path.isfile(credentials_path):
        print(f"找不到服務帳號金鑰: {credentials_path}")
        print("請依「Google_試算表_上傳設定說明.txt」建立金鑰並將檔案放在此路徑，或設定環境變數 GOOGLE_APPLICATION_CREDENTIALS")
        sys.exit(1)

    rows = load_csv_rows(csv_path)
    if not rows:
        print("CSV 為空，略過寫入。")
        return

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id)

    try:
        ws = sh.worksheet(student_folder_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=student_folder_name, rows=100, cols=10)

    ws.clear()
    ws.update(range_name="A1", values=rows)
    print(f"已寫入試算表：工作表「{student_folder_name}」，共 {len(rows)} 列。")


def upload_all(base_dir=None):
    """掃描「學生作業」目錄下各資料夾（排除 Script/Data/Doc/Gallery/__pycache__），若有評分表就上傳。"""
    base_dir = base_dir or BASE_DIR
    uploaded = 0
    for name in iter_student_folders(base_dir):
        folder = os.path.join(base_dir, name)
        csv_path = os.path.join(folder, f"{name}_評分表.csv")
        if not os.path.isfile(csv_path):
            continue
        try:
            upload_one(name, csv_path=csv_path)
            uploaded += 1
        except Exception as e:
            print(f"上傳「{name}」失敗: {e}")
    print(f"共上傳 {uploaded} 位學生。")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # 未給參數：批次上傳所有有評分表的學生
        upload_all()
    else:
        # 單一學生：python upload_scores_to_google_sheet.py 周昱翔_開心
        upload_one(sys.argv[1])
