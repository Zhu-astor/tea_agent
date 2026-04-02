import os
import time
import json
import requests
import io
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from PyPDF2 import PdfReader

# --- 配置區 ---
OPENROUTER_API_KEY = "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"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL_NAME = "nvidia/nemotron-3-nano-30b-a3b:free"

# 您的 Google Drive 收件匣 ID (從連結提取)
INBOX_FOLDER_ID = "18CbxAwraE0zbWJ9wI1rSo2_DEG1Nlpoy"
# 歸檔的根目錄 ID (建議在雲端另外設一個，或直接放在 Inbox 同層)
STORAGE_ROOT_ID = INBOX_FOLDER_ID 

SCOPES = ['https://www.googleapis.com/auth/drive']

# --- Google Drive 授權 ---
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

# --- 讀取文件內容 ---
def extract_text(service, file_id, mime_type):
    if 'pdf' in mime_type:
        request = service.files().get_media(fileId=file_id)
        f = io.BytesIO()
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        f.seek(0)
        reader = PdfReader(f)
        return " ".join([page.extract_text() for page in reader.pages[:3]]) # 讀前3頁
    elif 'text' in mime_type or 'plain' in mime_type:
        content = service.files().get_media(fileId=file_id).execute()
        return content.decode('utf-8')
    return "無法讀取此類型文件"

# --- 呼叫 OpenRouter AI ---
def get_ai_naming(text_content):
    prompt = f"""
    你是一位蟬吃茶的檔案管理專家。請閱讀以下文件內容，並根據規則產出檔名與分類。
    規則：日期_類別_內容簡述_上傳者 (日期請用今天 20260224，上傳者若不知則寫 系統)
    格式請嚴格輸出 JSON：{{"new_name": "檔名.pdf", "category": "分類名稱"}}
    文件內容：{text_content[:2000]}
    """
    headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}"}
    data = {
        "model": MODEL_NAME,
        "messages": [{"role": "user", "content": prompt}]
    }
    response = requests.post(OPENROUTER_URL, headers=headers, json=data)
    try:
        res_json = response.json()['choices'][0]['message']['content']
        return json.loads(res_json)
    except:
        return None

# --- 資料夾管理與移動 ---
def move_file(service, file_id, new_name, category):
    # 1. 找尋或創建分類資料夾
    query = f"name = '{category}' and mimeType = 'application/vnd.google-apps.folder' and '{STORAGE_ROOT_ID}' in parents"
    results = service.files().list(q=query).execute().get('files', [])
    
    if not results:
        file_metadata = {'name': category, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [STORAGE_ROOT_ID]}
        folder = service.files().create(body=file_metadata, fields='id').execute()
        target_folder_id = folder.get('id')
    else:
        target_folder_id = results[0]['id']

    # 2. 更改檔名並移動
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    service.files().update(fileId=file_id, addParents=target_folder_id, 
                           removeParents=previous_parents, body={'name': new_name}).execute()

# --- 主程式：24小時監聽 ---
def main():
    service = get_drive_service()
    print("AI 自動命名機器人啟動中...")
    
    while True:
        try:
            # 搜尋 Inbox 資料夾下的所有檔案
            query = f"'{INBOX_FOLDER_ID}' in parents and trashed = false"
            results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
            files = results.get('files', [])

            for file in files:
                if file['mimeType'] == 'application/vnd.google-apps.folder':
                    continue # 跳過資料夾
                
                print(f"處理中: {file['name']}")
                text = extract_text(service, file['id'], file['mimeType'])
                ai_decision = get_ai_naming(text)
                
                if ai_decision:
                    move_file(service, file['id'], ai_decision['new_name'], ai_decision['category'])
                    print(f"完成歸檔: {ai_decision['new_name']} -> {ai_decision['category']}")
            
            time.sleep(30) # 每分鐘檢查一次
        except Exception as e:
            print(f"發生錯誤: {e}")
            time.sleep(10)

if __name__ == "__main__":
    main()