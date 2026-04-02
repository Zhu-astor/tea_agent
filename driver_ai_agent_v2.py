import os
import time
import json
import requests
import io
import base64
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
from ai_converter import convert_ai_to_pdf, cleanup_temp_file

# ==========================================
# 1. 配置區
# ==========================================
OPENROUTER_API_KEY = "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"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
# 建議統一使用支援視覺與結構化輸出的模型
VISION_MODEL_NAME = "openai/gpt-4o-mini" 
TEXT_MODEL_NAME = "openai/gpt-4o-mini" 

# Google Drive 設定
INBOX_FOLDER_ID = "18CbxAwraE0zbWJ9wI1rSo2_DEG1Nlpoy"
STORAGE_ROOT_ID = INBOX_FOLDER_ID 
SCOPES = ['https://www.googleapis.com/auth/drive']

# 設計檔命名常數
PRESET_YEAR = "2026"
PRESET_DEVICE = "電腦"
PURPOSE_LIST = [
    "1a(機器人)", "banner(機器人)", "六宮格選單(機器人)", "名片小卡", "店面", 
    "社群平台", "公告", "立架", "布條", "酷卡", "貼紙", "桌面大圖/柱子", 
    "網站", "包裝/禮盒", "兌換券/折價券", "茶包", "dm菜單", "三折"
]

# ==========================================
# 2. Google Drive 核心功能
# ==========================================
import re # 記得在檔案最上方 import re

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

# ==========================================
# 3. AI 文檔與視覺處理邏輯
# ==========================================
def extract_text(file_path, mime_type):
    """讀取本地文本或 PDF 內容"""
    if 'pdf' in mime_type:
        reader = PdfReader(file_path)
        return " ".join([page.extract_text() for page in reader.pages[:3] if page.extract_text()])
    else:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()

def get_ai_naming_text(text_content, original_name):
    """處理純文檔的 AI 命名"""
    prompt = f"""你是一位檔案管理專家。請閱讀以下文件內容，並根據規則產出檔名與分類。
原始檔名：{original_name}
規則：日期_類別_內容簡述_上傳者 (日期請用 {datetime.now().strftime("%Y%m%d")}，上傳者寫 系統)
格式請嚴格輸出 JSON：{{"new_name": "檔名", "category": "分類名稱", "description": "文件重點摘要"}}
文件內容：{text_content[:2000]}"""
    
    headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}", "Content-Type": "application/json"}
    data = {"model": TEXT_MODEL_NAME, "messages": [{"role": "user", "content": prompt}]}
    
    try:
        response = requests.post(OPENROUTER_URL, headers=headers, json=data)
        res_json = response.json()['choices'][0]['message']['content'].replace("```json", "").replace("```", "").strip()
        return json.loads(res_json)
    except:
        return None

def get_images_and_dimensions(file_path):
    """讀取設計檔轉 Base64 列表與尺寸"""
    ext = os.path.splitext(file_path)[1].lower()
    images = []
    
    if ext == ".pdf":
        pages = convert_from_path(file_path, first_page=1, last_page=5)
        images = pages
    else:
        # 🌟 【關鍵修復】使用 with 語法，確保圖片讀取後記憶體會自動釋放檔案佔用
        with Image.open(file_path) as img:
            images = [img.copy()]

    width, height = images[0].size
    size_str = f"{width}x{height}px"
    purpose_str = None

    dims = {width, height}
    if dims == {1024, 768}: size_str, purpose_str = "1024x768px", "1a(機器人)"
    elif dims == {770, 250}: size_str, purpose_str = "770x250px", "banner(機器人)"
    elif dims == {2500, 1686}: size_str, purpose_str = "2500x1686px", "六宮格選單(機器人)"
    elif dims == {1080, 1350}: size_str, purpose_str = "1080x1350px", "社群平台"
    elif dims == {1080, 1920}: size_str, purpose_str = "1080x1920px", "社群平台"
    else:
        ratio = max(width, height) / min(width, height)
        if 3.20 <= ratio <= 3.26: size_str, purpose_str = "197x61cm", "布條"
        elif 2.64 <= ratio <= 2.69: size_str, purpose_str = "60x160cm", "立架"
        elif 3.30 <= ratio <= 3.36: size_str, purpose_str = "54x180mm", "名片小卡"
        elif 1.64 <= ratio <= 1.69: size_str, purpose_str = "108x180mm", "酷卡"
        elif 1.23 <= ratio <= 1.27: size_str, purpose_str = "40x50cm", "店面"
        elif 1.38 <= ratio <= 1.43: size_str, purpose_str = "A4_或_50x70_或_4k", "請AI判斷"

    base64_images = []
    for img in images:
        buffered = BytesIO()
        img.convert("RGB").save(buffered, format="JPEG")
        base64_images.append(base64.b64encode(buffered.getvalue()).decode("utf-8"))
    
    return base64_images, size_str, purpose_str

def ask_vision_ai(base64_images, size_str, purpose_str, original_filename):
    """處理設計圖的 AI 視覺辨識（包含完整的活動清單與對照規則）"""
    purpose_prompt = f"這張圖的用途已被確認為：【{purpose_str}】，請直接輸出「{purpose_str}」。" if purpose_str and purpose_str != "請AI判斷" else f"系統無法自動判斷用途。請觀察排版，從清單選出最適合的：{PURPOSE_LIST}"

    # 🌟 【重點修復】把完整的活動對照規則放回 Prompt 中
    prompt = f"""你是一個精準的檔案分類系統。請觀察提供的圖片，並「強烈參考」【原始檔案名稱】：{original_filename}

任務 1：畫面觀察與描述
請詳細記錄畫面中的文字、品牌、產品、價格，以及檔名給予的線索。

任務 2：判斷【活動名稱】
請仔細閱讀圖片內容與檔名，並「嚴格遵守」以下對應規則進行單選輸出：
- 若畫面或檔名包含「跨年」、「共享尾牙」、「一年免費喝」 -> 輸出：一年免費喝
- 若畫面或檔名包含「蟬吃茶年節禮」、「春節禮」、「蟬茶禮組」、「年節禮」 -> 輸出：年節禮
- 若畫面或檔名包含「母親節特惠組」、「母親節」 -> 輸出：母親節
- 若畫面或檔名包含「父親節」、「父親節禮組」、「88節」 -> 輸出：父親節
- 若畫面或檔名包含「中秋好禮」、「慶團圓」、「中秋節禮」、「中秋」 -> 輸出：中秋
- 若畫面或檔名包含「端午特惠」、「端午」 -> 輸出：端午
- 若畫面或檔名包含「周年慶」 -> 輸出：周年慶
- 若以上皆非，或無法辨識 -> 輸出：其他
💡 重要提示：如果圖片上沒有明確寫出活動，但檔名有暗示，請優先採納檔名線索！

任務 3：判斷【用途】
{purpose_prompt}

請嚴格輸出 JSON，格式如下：
{{"畫面描述": "簡述內容與檔名線索", "活動名稱": "填入選項", "用途": "填入選項"}}"""

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
        return {"畫面描述": "辨識失敗", "活動名稱": "其他", "用途": "未分類"}

# ==========================================
# 4. 主程式：24小時監聽
# ==========================================
def main():
    service = get_drive_service()
    print("🚀 AI 全能檔案管理機器人啟動中... (監聽 Drive Inbox)")
    
    while True:
        try:
            query = f"'{INBOX_FOLDER_ID}' in parents and trashed = false"
            results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
            files = results.get('files', [])

            for file in files:
                if file['mimeType'] == 'application/vnd.google-apps.folder':
                    continue 
                
                original_name = file['name']
                file_id = file['id']
                ext = os.path.splitext(original_name)[1].lower()
                print(f"\n📥 抓取到新檔案: {original_name} | 下載至本地分析中...")
                
                # 下載檔案到本地暫存
                local_path = download_drive_file(service, file_id, original_name)
                is_temp_pdf = False
                target_file = local_path

                try: # 🌟 新增的內部 try，專門保護單一檔案的處理過程
                    # 【路線 A：設計圖檔與工作檔】
                    if ext in ['.jpg', '.jpeg', '.png', '.pdf', '.ai']:
                        if ext == '.ai':
                            target_file = convert_ai_to_pdf(local_path)
                            is_temp_pdf = True
                            
                        base64_images, size_str, purpose_str = get_images_and_dimensions(target_file)
                        ai_result = ask_vision_ai(base64_images, size_str, purpose_str, original_name)
                        
                        activity = ai_result.get("活動名稱", "其他")
                        final_purpose = ai_result.get("用途", "未知用途")
                        description = f"👀 AI 視覺觀察：{ai_result.get('畫面描述', '無紀錄')}\n📏 尺寸判定：{size_str}"
                        
                        # 🌟 【新邏輯】從原始檔名判斷年份，若無則用 PRESET_YEAR
                        file_year = get_year_from_filename(original_name, PRESET_YEAR)
                        
                        date_str = datetime.now().strftime("%Y%m%d")
                        safe_size = size_str.replace(" ", "")
                        
                        # 檔名最前面的年份，改用我們剛剛抓出來的 file_year
                        new_name = f"{file_year}_{activity}_{safe_size}_{final_purpose}_{PRESET_DEVICE}_{date_str}{ext}"
                        
                        print(f"✅ 視覺分析完成! 年份: {file_year} | 活動: {activity} | 用途: {final_purpose}")

                    # 【路線 B：純文字或文件】
                    else:
                        text = extract_text(local_path, file['mimeType'])
                        ai_result = get_ai_naming_text(text, original_name)
                        
                        if ai_result:
                            # 🌟 【新邏輯】文本也要抓年份
                            file_year = get_year_from_filename(original_name, PRESET_YEAR)
                            activity = ai_result.get('category', '未分類')
                            
                            new_name = ai_result.get('new_name', original_name)
                            # 確保新檔名帶有年份前綴 (可依你的規則調整)
                            if not new_name.startswith(file_year):
                                new_name = f"{file_year}_{new_name}"

                            description = f"📝 AI 文檔摘要：{ai_result.get('description', '無紀錄')}"
                            print(f"✅ 文檔分析完成! 年份: {file_year} | 分類: {activity}")
                        else:
                            raise ValueError("AI 回傳空白結果")

                    # 🌟 【新邏輯】執行 Drive 的重新命名、移動與更新備註 (傳入 file_year 和 activity)
                    move_and_update_file(service, file_id, new_name, file_year, activity, description)
                    print(f"🎯 歸檔成功: {new_name} -> [{file_year}] / [{activity}] 資料夾")

                except Exception as file_error:
                    # 🌟 【關鍵修復】攔截加密或損壞的檔案，將它們移至隔離區
                    error_msg = str(file_error).lower()
                    if "password" in error_msg or "encrypt" in error_msg:
                        print(f"⚠️ 警告: {original_name} 受到密碼保護，AI 無法讀取！")
                        move_and_update_file(service, file_id, f"[加密無法讀取]_{original_name}", "需人工確認", "此檔案受密碼保護，系統已略過。")
                    else:
                        print(f"❌ 檔案解析失敗: {file_error}")
                        move_and_update_file(service, file_id, f"[解析失敗]_{original_name}", "需人工確認", f"錯誤資訊: {file_error}")
                    print("🎯 已將問題檔案移至 [需人工確認] 資料夾，避免卡死。")

                finally:
                    # 徹底清理本地暫存檔
                    try:
                        if os.path.exists(local_path):
                            os.remove(local_path)
                        if is_temp_pdf and os.path.exists(target_file):
                            cleanup_temp_file(target_file)
                    except Exception as cleanup_err:
                        print(f"🧹 暫存檔清理失敗 (可能仍在佔用中): {cleanup_err}")
                        
            time.sleep(30) 
        except Exception as e:
            print(f"❌ 發生錯誤: {e}")
            time.sleep(10)

if __name__ == "__main__":
    main()