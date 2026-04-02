import os, time, json, requests, io, base64, re
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

# 處理 .ai 檔案的自訂模組 (請確保 ai_converter.py 在同目錄)
try:
    from ai_converter import convert_ai_to_pdf, cleanup_temp_file
except ImportError:
    print("❌ 警告：找不到 ai_converter.py")

# ==========================================
# 1. 配置區 (IDs 與 API Key)
# ==========================================
Image.MAX_IMAGE_PIXELS = None 
OPENROUTER_API_KEY = "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"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
VISION_MODEL_NAME = "openai/gpt-4o-mini" 

INBOX_FOLDER_ID = "18CbxAwraE0zbWJ9wI1rSo2_DEG1Nlpoy"
PROCESSED_ROOT_ID = "1Ofe5aAL5w4gdQ3Y6nKEY4Y8llca329BC" # 歸檔總目錄
SCOPES = ['https://www.googleapis.com/auth/drive']

LARGE_FILE_THRESHOLD_MB = 25 
POLLING_INTERVAL_SECONDS = 30 

PURPOSE_LIST = [
    "1a(機器人)", "banner(機器人)", "六宮格選單(機器人)", "名片小卡", "店面", 
    "社群平台", "公告", "立架", "布條", "酷卡", "貼紙", "桌面大圖/柱子", 
    "網站", "包裝/禮盒", "兌換券/折價券", "茶包", "dm菜單", "三折"
]
SIZE_RULE_MAP = {
    "197X61CM": "布條",
    "60X160CM": "立架",
    "1024X768PX": "1a(機器人)",
    "770X250PX": "banner(機器人)",
    "2500X1686PX": "六宮格選單(機器人)",
    "1080X1350PX": "社群平台",
    "1080X1920PX": "社群平台",
    "54X180MM": "名片小卡",
    "108X180MM": "酷卡",
    "A4": "dm菜單"
}
# ==========================================
# 1.1 活動名稱對應表 (括號內為主)
# ==========================================
ACTIVITY_MAP = {
    "一年免費喝": ["一年免費喝", "跨年共享尾牙", "跨年一年免費喝"],
    "年節禮": ["年節禮", "蟬吃茶年節禮", "蟬吃春節禮", "蟬茶禮組"],
    "母親節": ["母親節", "感恩母親節特惠組"],
    "父親節": ["父親節", "感恩父親節", "父親節禮組", "88節"],
    "中秋": ["中秋", "蟬吃中秋好禮", "中秋慶團圓", "中秋節禮"],
    "端午": ["端午", "端午特惠組"],
    "周年慶": ["周年慶", "蟬吃xx周年慶", "周年慶"],
}

def map_activity_name(ai_activity):
    """將 AI 判斷的活動名稱映射到標準命名"""
    # 如果 AI 沒給出判定，先預設為其他
    if not ai_activity or ai_activity == "其他":
        return "其他"

    for standard_name, keywords in ACTIVITY_MAP.items():
        for kw in keywords:
            if kw in ai_activity:
                return standard_name
                
    # 🌟 修正：若沒對到清單中的關鍵字 (如火龍果、蜂蜜等)，一律回傳 "其他"
    return "其他"
# ==========================================
# 2. 物理偵測與 Meta 提取
# ==========================================

def get_essential_meta(original_name):
    """提取年份、設備與特殊標記，不提取用途"""
    # 🌟 修正：年份預設改為 "無年份"
    meta = {"year": "無年份", "device": "電腦", "is_1a": False}
    
    # 提取年份 (支援 2020-2029)
    year_match = re.search(r'(20\d{2})', original_name)
    if year_match: 
        meta["year"] = year_match.group(1)
    
    # 提取設備
    for d in ["電腦", "iPad", "手機", "DJI", "iPhone"]:
        if d in original_name: 
            meta["device"] = d
            break
            
    if "1a" in original_name.lower(): 
        meta["is_1a"] = True
        
    return meta

def analyze_physical_size(file_path, original_name):
    """【整合偵測】確保橫直比例也能正確換算"""
    info = {"size": "未知尺寸", "is_narrow": False}
    ext = os.path.splitext(file_path)[1].lower()
    
    # A. 檔名優先
    size_match = re.search(r'(\d+[xX]\d+(?:cm|mm|px)?)', original_name)
    if size_match:
        info["size"] = size_match.group(1).upper()
    
    # B. PDF 偵測
    elif ext == '.pdf':
        try:
            reader = PdfReader(file_path)
            mb = reader.pages[0].mediabox
            w_mm, h_mm = round(float(mb.width)/72*25.4), round(float(mb.height)/72*25.4)
            short, long = sorted([w_mm, h_mm])
            # A4 判定
            if 205 <= short <= 215 and 292 <= long <= 302: info["size"] = "A4"
            else: info["size"] = f"{w_mm}x{h_mm}mm"
        except: pass
        
    # C. JPG/PNG DPI 解析 (增加長寬比計算)
    elif ext in ['.jpg', '.jpeg', '.png']:
        try:
            with Image.open(file_path) as img:
                w_px, h_px = img.size
                ratio = max(w_px, h_px) / min(w_px, h_px)
                if ratio >= 2.5: info["is_narrow"] = True # 強制修正狹長判定
                
                if info["size"] == "未知尺寸":
                    dpi = img.info.get('dpi')
                    if dpi and dpi[0] > 0:
                        w_cm = round((w_px / dpi[0]) * 2.54)
                        h_cm = round((h_px / dpi[1]) * 2.54)
                        info["size"] = f"{w_cm}x{h_cm}cm".upper()
                    else:
                        info["size"] = f"{w_px}x{h_px}px".upper()
        except: pass
    return info

# ==========================================
# 3. 核心流程
# ==========================================

    
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
        with open('token.json', 'w') as token: token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def get_or_create_year_folder(service, parent_id, year):
    """自動建立年份子資料夾"""
    query = f"name = '{year}' and '{parent_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    folders = service.files().list(q=query).execute().get('files', [])
    if folders: return folders[0]['id']
    folder_meta = {'name': year, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    return service.files().create(body=folder_meta, fields='id').execute().get('id')

def archive_and_rename(service, file_id, new_name, year):
    """【整合歸檔】重新命名 + 移至年份資料夾"""
    try:
        target_folder_id = get_or_create_year_folder(service, PROCESSED_ROOT_ID, year)
        file = service.files().get(fileId=file_id, fields='parents').execute()
        prev_parents = ",".join(file.get('parents'))
        service.files().update(
            fileId=file_id,
            body={'name': new_name},
            addParents=target_folder_id,
            removeParents=prev_parents
        ).execute()
        return True
    except Exception as e:
        print(f"❌ 歸檔失敗: {e}"); return False

def download_drive_file(service, file_id, file_name):
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done: _, done = downloader.next_chunk()
    fh.close()
    return file_name

# ==========================================
# 4. AI 盲測影像處理
# ==========================================

def get_images_for_ai(file_path):
    """優化 AI 預覽圖 (DPI 100)"""
    ext = os.path.splitext(file_path)[1].lower()
    imgs = []
    if ext == ".pdf":
        imgs = convert_from_path(file_path, dpi=100, first_page=1, last_page=3)
    else:
        with Image.open(file_path) as img: imgs = [img.copy()]
    
    b64_list = []
    for i in imgs:
        buf = BytesIO()
        i.convert("RGB").save(buf, format="JPEG", quality=85)
        b64_list.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
    return b64_list

def ask_vision_ai_blind(base64_images, size_str, hidden_name, narrow_hint=""):
    """盲測專用 Prompt"""
    prompt = f"""你正在進行【視覺盲測】。檔名【{hidden_name}】不具參考價值。
物理尺寸：【{size_str}】。{narrow_hint}
任務：觀察圖片判定活動與用途。用途請從 {PURPOSE_LIST} 選擇。
注意：畫面有茶包袋請選「茶包」；比例狹長請優先考慮「布條」。
輸出 JSON：{{"視覺描述": "...", "活動判定": "...", "用途判定": "..."}}"""

    payload = {
        "model": VISION_MODEL_NAME,
        "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}] + 
                    [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img}"}} for img in base64_images]}],
        "temperature": 0.1
    }
    headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}"}
    try:
        res = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=60)
        return json.loads(res.json()['choices'][0]['message']['content'].replace("```json", "").replace("```", "").strip())
    except: return {"活動判定": "其他", "用途判定": "未分類"}

# ==========================================
# 5. 主流程 (監聽與處理)
# ==========================================
def process_file(service, file_item):
    orig_name = file_item['name']
    f_id = file_item['id']
    ext = os.path.splitext(orig_name)[1]
    
    print(f"\n" + "="*50)
    print(f"📄 【原始檔名】: {orig_name}")

    local_path, target_file, is_temp_pdf = None, None, False

    try:
        # 1. 提取基礎 Meta
        meta = get_essential_meta(orig_name)
        local_path = download_drive_file(service, f_id, orig_name)
        target_file = local_path
        
        if orig_name.lower().endswith('.ai'):
            target_file = convert_ai_to_pdf(local_path)
            is_temp_pdf = True

        # 2. 物理尺寸偵測
        phys = analyze_physical_size(target_file, orig_name)
        size_label = phys['size']
        print(f"📏 【物理尺寸】: {size_label} (狹長判定: {phys['is_narrow']})")

        # 3. AI 盲測
        b64_imgs = get_images_for_ai(target_file)
        hint = "此圖比例狹長，請優先考慮布條。" if phys["is_narrow"] else ""
        ai_res = ask_vision_ai_blind(b64_imgs, size_label, f"HIDE_{int(time.time())}", hint)

        # 4. 邏輯修正與映射
        # (A) 用途校正
        final_purpose = ai_res.get('用途判定', '未分類')
        for rule_size, rule_purpose in SIZE_RULE_MAP.items():
            if rule_size in size_label:
                final_purpose = rule_purpose
                break
        
        # (B) 活動名稱校正 (依照括號清單)
        raw_activity = ai_res.get('活動判定', '其他')
        final_activity = map_activity_name(raw_activity)

        print(f"🤖 【AI 判斷內容】:")
        print(f"   ├─ 視覺描述: {ai_res.get('視覺描述', 'None')}")
        print(f"   └─ 最終判定: {final_activity} / {final_purpose}")

        # 5. 🌟 按照指定順序產生新名稱: [年份]_[活動名稱]_[尺寸]_[用途]_[設備]
        # 格式：2026_一年免費喝_197X61CM_布條_電腦.jpg
        # 🌟 命名順序：[年份]_[活動名稱]_[尺寸]_[用途]_[設備]
        new_name = f"{meta['year']}_{final_activity}_{size_label}_{final_purpose}_{meta['device']}{ext}"
        
        print(f"📝 【最終改名】: {new_name}")
        
        if archive_and_rename(service, f_id, new_name, meta['year']):
            print(f"✅ 【歸檔成功】: 已移至 /{meta['year']} 資料夾")
        
    except Exception as e:
        print(f"❌ 【處理失敗】: {e}")
    finally:
        if local_path and os.path.exists(local_path): os.remove(local_path)
        if is_temp_pdf: cleanup_temp_file(target_file)
    print("="*50)

def main():
    service = get_drive_service()
    print(f"🚀 Tea Agent V5 啟動 (監聽中)...")
    while True:
        try:
            query = f"'{INBOX_FOLDER_ID}' in parents and trashed = false"
            items = service.files().list(q=query, fields="files(id, name, mimeType, size)").execute().get('files', [])
            if items:
                # 分流並處理 (雙隊列邏輯可在這擴充)
                for f in items:
                    if f['mimeType'] == 'application/vnd.google-apps.folder': continue
                    process_file(service, f)
            else: print(".", end="", flush=True)
            time.sleep(POLLING_INTERVAL_SECONDS)
        except Exception as e: time.sleep(60)

if __name__ == "__main__":
    main()