import os
import time
import json
import requests
import io
import base64
import re
from datetime import datetime
from io import BytesIO
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfReader

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# 引入自訂的轉換模組
from ai_converter_v2 import convert_ai_to_pdf, cleanup_temp_file

# ==========================================
# 1. 配置區
# ==========================================
OPENROUTER_API_KEY = "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"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
VISION_MODEL_NAME = "openai/gpt-4o-mini" 

INBOX_FOLDER_ID = "18CbxAwraE0zbWJ9wI1rSo2_DEG1Nlpoy"
STORAGE_ROOT_ID = INBOX_FOLDER_ID 
MANUAL_REVIEW_FOLDER_ID = "1qw3UcZx8RBXjMJjaDHgz-atp0oyNBXsu"
SCOPES = ['https://www.googleapis.com/auth/drive']

PRESET_YEAR = "2026"
PRESET_DEVICE = "電腦"

# 用途清單 (用於 AI 選項)
PURPOSE_LIST = [
    "1a(機器人)", "banner(機器人)", "六宮格選單(機器人)", "名片小卡", "店面", 
    "社群平台", "公告", "立架", "布條", "酷卡", "貼紙", "桌面大圖/柱子", 
    "網站", "包裝/禮盒", "兌換券/折價券", "茶包", "dm菜單", "三折"
]

# ==========================================
# 2. 強力攔截與物理偵測模組
# ==========================================
# ==========================================
# 2. Google Drive 核心功能
# ==========================================

import re # 記得在檔案最上方 import re
def move_to_manual_review_drive(service, file_id, file_name, reason):
    """將出錯或超大的檔案移至人工確認資料夾，並加上備註"""
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents', []))
    
    body = {
        'name': f"[人工確認]_{file_name}",
        'description': f"分流原因: {reason}"
    }
    
    service.files().update(
        fileId=file_id,
        addParents=MANUAL_REVIEW_FOLDER_ID,
        removeParents=previous_parents,
        body=body
    ).execute()
    print(f"⚠️  [分流系統] 檔案已移至人工確認區：{file_name} (原因: {reason})")
    
def get_year_from_filename(filename, fallback_year):
    """從原始檔名中尋找 20xx 的年份，找不到就用預設值"""
    match = re.search(r'(20\d{2})', filename)
    if match:
        return match.group(1)
    return fallback_year

def get_drive_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def download_drive_file(service, file_id, file_name):
    """將雲端檔案暫存到本地以供視覺處理"""
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    
    fh.close() # 🌟 【關鍵修復】下載完畢後必須關閉檔案，否則 Windows 不給刪除！
    return file_name

def move_and_update_file(service, file_id, new_name, year_folder, activity_folder, description=""):
    """移動檔案至 [年份] -> [活動] 雙層資料夾，並更新檔名與備註"""
    
    # --- 第一層：找尋或創建「年份」資料夾 ---
    query_year = f"name = '{year_folder}' and mimeType = 'application/vnd.google-apps.folder' and '{STORAGE_ROOT_ID}' in parents and trashed = false"
    results_year = service.files().list(q=query_year).execute().get('files', [])
    
    if not results_year:
        year_metadata = {'name': year_folder, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [STORAGE_ROOT_ID]}
        folder_y = service.files().create(body=year_metadata, fields='id').execute()
        target_year_id = folder_y.get('id')
    else:
        target_year_id = results_year[0]['id']

    # --- 第二層：在「年份」資料夾底下，找尋或創建「活動」子資料夾 ---
    query_activity = f"name = '{activity_folder}' and mimeType = 'application/vnd.google-apps.folder' and '{target_year_id}' in parents and trashed = false"
    results_activity = service.files().list(q=query_activity).execute().get('files', [])
    
    if not results_activity:
        act_metadata = {'name': activity_folder, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [target_year_id]}
        folder_a = service.files().create(body=act_metadata, fields='id').execute()
        target_activity_id = folder_a.get('id')
    else:
        target_activity_id = results_activity[0]['id']

    # --- 執行檔案移動與更新 ---
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents', []))
    
    body = {'name': new_name}
    if description:
        body['description'] = description

    service.files().update(
        fileId=file_id, 
        addParents=target_activity_id, # 將檔案放入第二層的活動資料夾中
        removeParents=previous_parents, 
        body=body
    ).execute()
    
def get_info_from_filename(filename):
    """
    【霸道攔截器】從原始檔名精準擷取尺寸與用途。
    優先權最高，直接覆蓋後續 AI 與幾何判定的結果。
    """
    info = {"size": None, "purpose": None}
    fn = filename.lower()

    # --- 1. 偵測用途關鍵字 ---
    purpose_map = {
        "1a": "1a(機器人)",
        "banner": "banner(機器人)",
        "六宮格": "六宮格選單(機器人)",
        "名片": "名片小卡",
        "店面": "店面",
        "ig": "社群平台",
        "fb": "社群平台",
        "社群": "社群平台",
        "公告": "公告",
        "立架": "立架",
        "布條": "布條",
        "酷卡": "酷卡",
        "貼紙": "貼紙",
        "桌面": "桌面大圖/柱子",
        "柱子": "桌面大圖/柱子",
        "茶包": "茶包", # 🌟 獨立分類，不再誤判為包裝
        "包裝": "包裝/禮盒",
        "禮盒": "包裝/禮盒",
        "兌換券": "兌換券/折價券",
        "折價券": "兌換券/折價券",
        "菜單": "dm菜單",
        "dm": "dm菜單",
        "三折": "三折"
    }
    for kw, val in purpose_map.items():
        if kw in fn:
            info["purpose"] = val
            break

    # --- 2. 偵測尺寸關鍵字 ---
    # 支援 A4, 4K, 或是 108x180mm 等格式
    size_match = re.search(r'(\d+[xX]\d+(?:mm|cm|px)?|[aA][345]|4[kK])', fn)
    if size_match:
        info["size"] = size_match.group(1).upper()
    
    return info

def get_pdf_physical_size(pdf_path):
    """
    【物理偵測】利用 PDF MediaBox 點數精準區分 A4, 50x70, 4K。
    1 pt = 1/72 inch = 0.3527 mm
    """
    try:
        reader = PdfReader(pdf_path)
        mb = reader.pages[0].mediabox
        w_pt, h_pt = float(mb.width), float(mb.height)
        short, long = sorted([round((w_pt/72)*25.4), round((h_pt/72)*25.4)])
        
        # 精準比對 mm (允許小誤差)
        if 205 <= short <= 215 and 292 <= long <= 302: return "A4"
        if 495 <= short <= 505 and 695 <= long <= 705: return "50x70cm"
        if 385 <= short <= 395 and 535 <= long <= 545: return "4K"
        return f"{short}x{long}mm"
    except:
        return None

# ==========================================
# 3. 視覺處理邏輯 (已整合攔截器)
# ==========================================

def get_images_and_dimensions(file_path, filename_info):
    """讀取設計檔，若檔名已有尺寸則優先使用"""
    ext = os.path.splitext(file_path)[1].lower()
    images = []
    
    # 物理偵測優先 (針對 AI 轉出的 PDF)
    physical_size = None
    if ext == ".pdf":
        physical_size = get_pdf_physical_size(file_path)
        images = convert_from_path(file_path, first_page=1, last_page=5)
    else:
        with Image.open(file_path) as img:
            images = [img.copy()]

    # 尺寸決定權重：1. 檔名寫的 > 2. PDF 物理偵測 > 3. 像素計算
    width, height = images[0].size
    if filename_info["size"]:
        size_str = filename_info["size"]
    elif physical_size:
        size_str = physical_size
    else:
        size_str = f"{width}x{height}px"

    # 用途決定權重：1. 檔名寫的 > 2. 比例計算
    purpose_str = filename_info["purpose"]
    if not purpose_str:
        ratio = max(width, height) / min(width, height)
        # (保留原本的比例判斷作為 fallback)
        if 3.20 <= ratio <= 3.26: purpose_str = "布條"
        elif 2.64 <= ratio <= 2.69: purpose_str = "立架"
        elif 1.64 <= ratio <= 1.69: purpose_str = "酷卡"
        elif 1.23 <= ratio <= 1.27: purpose_str = "店面"

    base64_images = []
    for img in images:
        buffered = BytesIO()
        img.convert("RGB").save(buffered, format="JPEG")
        base64_images.append(base64.b64encode(buffered.getvalue()).decode("utf-8"))
    
    return base64_images, size_str, purpose_str

def ask_vision_ai(base64_images, size_str, purpose_str, original_filename):
    """處理視覺辨識，並加入 1a 與 茶包 的特殊說明"""
    
    # 針對 1a 的備註
    one_a_note = "【注意：1a 是指點餐機大螢幕的圖示。】" if "1a" in original_filename.lower() else ""

    purpose_prompt = f"目前初步判定用途為：{purpose_str}，請確認或校正。" if purpose_str else f"請從清單挑選最適合用途：{PURPOSE_LIST}"

    prompt = f"""你是一個精準的設計檔案分類系統。{one_a_note}
任務：觀察圖片並參考原始檔名【{original_filename}】進行歸類。

對應規則：
- 若畫面或檔名包含「茶包」 -> 輸出活動：其他，用途：茶包 (請勿選包裝/禮盒)
- 若檔名有「1a」 -> 輸出用途：1a(機器人)
- 活動清單：[一年免費喝, 年節禮, 母親節, 父親節, 中秋, 端午, 周年慶, 其他]

請嚴格輸出 JSON：
{{"畫面描述": "簡述", "活動名稱": "填入選項", "用途": "填入選項"}}"""

    headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}", "Content-Type": "application/json"}
    content_list = [{"type": "text", "text": prompt}]
    for img_b64 in base64_images:
        content_list.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}})

    payload = {"model": VISION_MODEL_NAME, "messages": [{"role": "user", "content": content_list}], "temperature": 0.1}
    
    try:
        response = requests.post(OPENROUTER_URL, headers=headers, json=payload)
        res_json = response.json()['choices'][0]['message']['content'].replace("```json", "").replace("```", "").strip()
        return json.loads(res_json)
    except:
        return {"畫面描述": "辨識失敗", "活動名稱": "其他", "用途": purpose_str or "未分類"}

# ==========================================
# 4. 主程式邏輯
# ==========================================

# (其他輔助函數如 get_drive_service, download_drive_file, move_and_update_file 保持不變)

def main():
    service = get_drive_service()
    print("🚀 AI 檔案管理員 V4 (防崩潰分流版) 啟動中...")
    
    while True:
        try:
            # 這是最外層的 try，對應最下方的 except Exception as e
            query = f"'{INBOX_FOLDER_ID}' in parents and trashed = false"
            results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
            files = results.get('files', [])

            for file in files:
                if file['mimeType'] == 'application/vnd.google-apps.folder': continue 
                
                original_name = file['name']
                file_id = file['id']
                print(f"\n📥 處理檔案: {original_name}")

                # 變數初始化（放在 try 外面，確保 finally 抓得到）
                local_path = None
                target_file = None
                is_temp_pdf = False

                try:
                    # --- 這裡開始是內層 try，處理單一檔案 ---
                    
                    # [步驟 1] 檔名強力攔截
                    filename_info = get_info_from_filename(original_name)
                    file_year = get_year_from_filename(original_name, PRESET_YEAR)

                    # 下載檔案
                    local_path = download_drive_file(service, file_id, original_name)
                    target_file = local_path

                    # [步驟 2] AI 檔案轉檔
                    if '.ai' in original_name.lower():
                        # 使用方案一：Inkscape 轉檔
                        target_file = convert_ai_to_pdf(local_path)
                        is_temp_pdf = True

                    # [步驟 3] 獲取圖片 (此處最容易發生像素爆炸)
                    # 傳入 filename_info 優先使用檔名標註的尺寸
                    base64_images, size_str, purpose_str = get_images_and_dimensions(target_file, filename_info)

                    # [步驟 4] AI 視覺辨識
                    ai_result = ask_vision_ai(base64_images, size_str, purpose_str, original_name)
                    
                    # 最終合成邏輯
                    activity = ai_result.get("活動名稱", "其他")
                    final_purpose = filename_info["purpose"] or ai_result.get("用途", "未分類")
                    date_str = datetime.now().strftime("%Y%m%d")
                    
                    new_name = f"{file_year}_{activity}_{size_str.replace(' ', '')}_{final_purpose}_{PRESET_DEVICE}_{date_str}{os.path.splitext(original_name)[1]}"

                    # 移動檔案至成功資料夾
                    move_and_update_file(service, file_id, new_name, file_year, activity, f"AI 視覺觀察：{ai_result.get('畫面描述')}")
                    print(f"🎯 歸檔成功: {new_name}")

                except Exception as e:
                    # --- 內層的 except：處理單一檔案失敗的情況 ---
                    error_msg = str(e)
                    print(f"❌ 單一檔案處理失敗: {error_msg}")
                    
                    # 判定是否為大圖導致的像素爆炸，或是其他導致卡住的錯誤
                    if "exceeds limit" in error_msg or "decompression bomb" in error_msg.lower():
                        move_to_manual_review_drive(service, file_id, original_name, "檔案像素過大 (大圖輸出)")
                    else:
                        move_to_manual_review_drive(service, file_id, original_name, f"系統報錯: {error_msg[:50]}")

                finally:
                    # --- 內層的 finally：無論成功失敗都要清理電腦暫存檔 ---
                    if local_path and os.path.exists(local_path): 
                        os.remove(local_path)
                    if is_temp_pdf and target_file and os.path.exists(target_file): 
                        cleanup_temp_file(target_file)

            # 迴圈結束後休息，準備下一次掃描
            print("\n😴 掃描完畢，等待 30 秒後進行下次檢查...")
            time.sleep(30)

        except Exception as e:
            # --- 最外層的 except：處理 Google API 連線或重大邏輯錯誤 ---
            print(f"💥 [核心錯誤] 系統運行異常: {e}")
            time.sleep(10)
        finally:
            # 確保無論成功失敗都會清理暫存檔
            if local_path and os.path.exists(local_path): os.remove(local_path)
            if is_temp_pdf and target_file and os.path.exists(target_file): cleanup_temp_file(target_file)

if __name__ == "__main__":
    main()