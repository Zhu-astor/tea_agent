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

# 引入自訂轉換模組
try:
    from ai_converter import convert_ai_to_pdf, cleanup_temp_file
except ImportError:
    print("❌ 錯誤：找不到 ai_converter.py，請確認該檔案存在。")

# ==========================================
# 1. 配置與安全區
# ==========================================
Image.MAX_IMAGE_PIXELS = None 


OPENROUTER_API_KEY = "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"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
VISION_MODEL_NAME = "openai/gpt-4o-mini" 

INBOX_FOLDER_ID = "18CbxAwraE0zbWJ9wI1rSo2_DEG1Nlpoy"
STORAGE_ROOT_ID = INBOX_FOLDER_ID 
PROCESSED_FOLDER_ID  = "1Ofe5aAL5w4gdQ3Y6nKEY4Y8llca329BC"
SCOPES = ['https://www.googleapis.com/auth/drive']

LARGE_FILE_THRESHOLD_MB = 25 
POLLING_INTERVAL = 30 

PURPOSE_LIST = [
    "1a(機器人)", "banner(機器人)", "六宮格選單(機器人)", "名片小卡", "店面", 
    "社群平台", "公告", "立架", "布條", "酷卡", "貼紙", "桌面大圖/柱子", 
    "網站", "包裝/禮盒", "兌換券/折價券", "茶包", "dm菜單", "三折"
]

# ==========================================
# 2. 基礎提取器 (保留年份、設備、1a)
# ==========================================

def get_essential_meta(filename):
    """
    僅提取必要背景資訊，不攔截『用途』。
    """
    meta = {"year": "2026", "device": "電腦", "is_1a": False}
    fn = filename.lower()
    
    # 提取年份
    year_match = re.search(r'(20\d{2})', filename)
    if year_match: meta["year"] = year_match.group(1)
    
    # 提取設備
    for d in ["電腦", "iPad", "手機", "DJI", "iPhone"]:
        if d in filename:
            meta["device"] = d
            break
            
    # 提取 1a 標籤 (這是你說不需要關掉的關鍵代號)
    if "1a" in fn:
        meta["is_1a"] = True
        
    return meta

def get_pdf_physical_size(pdf_path):
    """【物理偵測】這屬於硬體掃描，不屬於檔名攔截，保留以增加精準度"""
    try:
        reader = PdfReader(pdf_path)
        mb = reader.pages[0].mediabox
        w_pt, h_pt = float(mb.width), float(mb.height)
        short, long = sorted([round((w_pt/72)*25.4), round((h_pt/72)*25.4)])
        if 205 <= short <= 215 and 292 <= long <= 302: return "A4"
        if 495 <= short <= 505 and 695 <= long <= 705: return "50x70cm"
        if 385 <= short <= 395 and 535 <= long <= 545: return "4K"
        return f"{short}x{long}mm"
    except: return None

# ==========================================
# 3. 盲測去識別化 (核心)
# ==========================================

def v4_hide_filename(meta, ext):
    """生成完全沒線索的假檔名"""
    tag = "_1a" if meta["is_1a"] else ""
    timestamp = datetime.now().strftime("%H%M%S")
    # 生成範例：2026_電腦_1a_TESTFILE_143005.ai
    return f"{meta['year']}_{meta['device']}{tag}_TESTFILE_{timestamp}{ext}"

# ==========================================
# 4. 影像與 AI 處理
# ==========================================

def get_images_and_dimensions(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    images, physical_size = [], None
    
    if ext == ".pdf":
        physical_size = get_pdf_physical_size(file_path)
        # DPI 100 記憶體優化
        images = convert_from_path(file_path, dpi=100, first_page=1, last_page=3)
    else:
        with Image.open(file_path) as img:
            images = [img.copy()]

    width, height = images[0].size
    size_str = physical_size or f"{width}x{height}px"

    base64_images = []
    for img in images:
        buf = BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=85)
        base64_images.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
        
    return base64_images, size_str

def ask_vision_ai_blind(base64_images, size_str, hidden_name):
    prompt = f"""你正在進行【視覺辨識盲測】。
提供的檔名：【{hidden_name}】是隨機生成的，請勿參考。
偵測尺寸：【{size_str}】。

任務：
1. 觀察圖片內容。
2. 判斷【活動名稱】：[一年免費喝, 年節禮, 中秋, 周年慶, 端午, 父親節, 其他]。
3. 判斷【用途】：請從 {PURPOSE_LIST} 中選擇。

💡 盲測準則：
- 若畫面出現茶包袋，請務必選「茶包」。
- 若檔名含「1a」或畫面是 APP 介面，請選 1a(機器人)。
- 忽視檔名中的 TESTFILE 字眼。

請回傳 JSON：
{{"視覺描述": "描述", "活動判定": "選項", "用途判定": "選項"}}"""

    headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}", "Content-Type": "application/json"}
    content_list = [{"type": "text", "text": prompt}]
    for b64 in base64_images:
        content_list.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})
    
    try:
        res = requests.post(OPENROUTER_URL, headers=headers, json={"model": VISION_MODEL_NAME, "messages": [{"role": "user", "content": content_list}], "temperature": 0.1}, timeout=60)
        return json.loads(res.json()['choices'][0]['message']['content'].replace("```json", "").replace("```", "").strip())
    except:
        return {"活動判定": "其他", "用途判定": "未分類"}

# ==========================================
# 5. 持續監聽與主流程
# ==========================================

def process_file(service, file_info, label=""):
    original_name = file_info['name']
    file_id = file_info['id']
    ext = os.path.splitext(original_name)[1].lower()
    
    # 變數初始化，防止 UnboundLocalError
    local_path = None
    target_file = None
    is_temp_pdf = False

    print(f"\n{label} 偵測到新檔案：{original_name}")

    try:
        # 1. 提取必要 Meta 並生成盲測檔名
        meta = get_essential_meta(original_name)
        hidden_name = v4_hide_filename(meta, ext)
        print(f"🕵️ 盲測開始：偽裝為 {hidden_name}")

        # 2. 下載檔案
        local_path = download_drive_file(service, file_id, original_name)
        target_file = local_path

        # 3. 處理 .ai 轉 PDF
        if ext == '.ai':
            target_file = convert_ai_to_pdf(local_path)
            is_temp_pdf = True

        # 4. 影像提取
        base64_imgs, size_str = get_images_and_dimensions(target_file)
        
        # 5. AI 盲測辨識 (完全不提供原始檔名)
        ai_res = ask_vision_ai_blind(base64_imgs, size_str, hidden_name)
        
        # 5.5 生成最終命名建議
        final_suggested_name = generate_final_name(meta, ai_res, size_str, ext)
        print(f"✨ 最終命名建議：{final_suggested_name}") 
        
        print(f"📊 視覺描述：{ai_res.get('視覺描述')}")
        print(f"🎯 AI 判定結果 -> 活動：{ai_res.get('活動判定')} | 用途：{ai_res.get('用途判定')}")

        # 6. 歸檔
        move_file_to_processed(service, file_id)
        print(f"📦 處理完畢，已移至歸檔區。")

    except Exception as e:
        print(f"❌ 處理失敗: {e}")
    finally:
        # 7. 安全清理
        if local_path and os.path.exists(local_path): os.remove(local_path)
        if is_temp_pdf and target_file and os.path.exists(target_file): cleanup_temp_file(target_file)

def main():
    service = get_drive_service()
    print(f"🚀 AI 視覺盲測監聽系統啟動 (Polling: {POLLING_INTERVAL}s)...")
    
    while True:
        try:
            query = f"'{INBOX_FOLDER_ID}' in parents and trashed = false"
            items = service.files().list(q=query, fields="files(id, name, mimeType, size)").execute().get('files', [])

            if items:
                normal_q, heavy_q = [], []
                for f in items:
                    if f['mimeType'] == 'application/vnd.google-apps.folder': continue
                    size_mb = int(f.get('size', 0)) / (1024 * 1024)
                    if size_mb > LARGE_FILE_THRESHOLD_MB:
                        heavy_q.append(f)
                    else:
                        normal_q.append(f)
                
                for f in normal_q: process_file(service, f, "[一般]")
                for f in heavy_q: process_file(service, f, "[大型]")
            else:
                print(".", end="", flush=True)

            time.sleep(POLLING_INTERVAL)
        except Exception as e:
            print(f"\n⚠️ 監聽中斷: {e}")
            time.sleep(60)
# ==========================================
# 4. Google Drive 底層工具包 (補回缺失的功能)
# ==========================================

def get_drive_service():
    """驗證並啟動 Google Drive 服務"""
    creds = None
    # token.json 儲存使用者的存取與重新整理權杖
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # 如果沒有可用的憑證，請讓使用者登入
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # 儲存憑證供下次使用
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return build('drive', 'v3', credentials=creds)

def download_drive_file(service, file_id, file_name):
    """將雲端檔案下載到本地端"""
    print(f"⏳ 正在下載檔案：{file_name}...")
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.close()
    return file_name

def move_file_to_processed(service, file_id):
    """處理完畢後，將檔案從 Inbox 移至 Processed 資料夾，防止重複讀取"""
    try:
        # 取得檔案目前的父資料夾
        file = service.files().get(fileId=file_id, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))
        
        # 移動檔案：新增目標資料夾，移除舊資料夾
        service.files().update(
            fileId=file_id,
            addParents=PROCESSED_FOLDER_ID,
            removeParents=previous_parents,
            fields='id, parents'
        ).execute()
    except Exception as e:
        print(f"⚠️ 移動檔案時發生錯誤 (請確認 PROCESSED_FOLDER_ID 是否正確): {e}")
        
        
        
def generate_final_name(meta, ai_res, size_str, original_ext):
    """
    根據 AI 辨識結果，組合出最後的歸檔檔名。
    格式範例：2026_店面_一年免費喝_布條_352x260cm.ai
    """
    year = meta.get("year", "2026")
    device = meta.get("device", "電腦")
    activity = ai_res.get("活動判定", "其他")
    purpose = ai_res.get("用途判定", "未分類")
    
    # 移除「其他」或「未分類」字眼，讓檔名更乾淨
    activity_str = f"_{activity}" if activity != "其他" else ""
    
    # 組合最終檔名
    final_name = f"{year}_{device}{activity_str}_{purpose}_{size_str}{original_ext}"
    return final_name


if __name__ == "__main__":
    main()