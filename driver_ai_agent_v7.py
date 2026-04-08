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
# 人工確認資料夾 ID (用來存放捷徑)
MANUAL_REVIEW_FOLDER_ID = "1zxySiRN8R8Q7tBxP0hrxks195OxQCScc"
INBOX_FOLDER_ID = "1riy-9Uxwkm6OGG3dWiXfWUByxBSNCw7v"
PROCESSED_ROOT_ID = "1riy-9Uxwkm6OGG3dWiXfWUByxBSNCw7v" # 歸檔總目錄
SCOPES = ['https://www.googleapis.com/auth/drive']

LARGE_FILE_THRESHOLD_MB = 25 
POLLING_INTERVAL_SECONDS = 30 

PURPOSE_LIST = [
    "1a(機器人)", "banner(機器人)", "六宮格選單(機器人)", "名片小卡","名片dm小卡","dm小卡", "店面", 
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
    "C045X045MM": "貼紙",
    "C050X050MM": "貼紙",
    "C055X055MM": "貼紙",
    "C060X060MM": "貼紙",
    "B060X060MM": "貼紙",
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
def map_activity_name(original_name, ai_activity):
    """活動名稱判定：檔名優先，沒對到清單一律回傳「其他」"""
    combined_text = (original_name + "_" + (ai_activity if ai_activity else "")).lower()

    for standard_name, keywords in ACTIVITY_MAP.items():
        for kw in keywords:
            if kw.lower() in combined_text:
                return standard_name
    return "其他"

def map_purpose_name(original_name, ai_purpose):
    """
    用途判定邏輯：
    1. 掃描檔名是否有 PURPOSE_LIST 的關鍵字 (最高優先)
    2. 檢查 AI 判斷是否有標籤關鍵字
    3. 若都沒對到，回傳「不確定」
    """
    fn = original_name.lower()
    ai_p = (ai_purpose if ai_purpose else "").lower()
    
    for standard_label in PURPOSE_LIST:
        # 去掉括號內容進行匹配，例如從 "1a(機器人)" 提取出 "1a"
        keyword = re.sub(r'\(.*\)', '', standard_label).lower()
        # 處理特殊標籤如 "桌面大圖/柱子"
        keywords = keyword.split('/')
        
        for kw in keywords:
            if kw and (kw in fn or kw in ai_p):
                return standard_label
                
    return "不確定"
# ==========================================
# 2. 物理偵測與 Meta 提取
# ==========================================

def get_essential_meta(service, file_id, original_name):
    """提取年份、設備，並判斷是否為重複上傳(更新)"""
    # 獲取檔案在雲端的詳細資訊 (version 代表修改次數)
    file_info = service.files().get(
        fileId=file_id, 
        fields='version, modifiedTime',
        supportsAllDrives=True 
    ).execute()
    version = int(file_info.get('version', 1))
    
    meta = {
        "year": "無年份", 
        "device": "電腦", 
        "is_update": version > 1, # 版本大於 1 代表被修改過或重複上傳
        "modify_date": ""
    }

    # 如果是更新，則獲取當前日期
    if meta["is_update"]:
        meta["modify_date"] = datetime.now().strftime("%Y%m%d")

    # 檢查年份
    year_match = re.search(r'(20\d{2})', original_name)
    if year_match: 
        meta["year"] = year_match.group(1)
    
    # 檢查設備
    for d in ["電腦", "iPad", "手機", "DJI", "iPhone"]:
        if d in original_name: 
            meta["device"] = d
            break
            
    return meta
def analyze_physical_size(file_path, original_name):
    info = {"size": "不確定", "is_narrow": False}
    ext = os.path.splitext(file_path)[1].lower()
    
    # A. 檔名優先 (修正：支援抓取 C 或 B 開頭的尺寸)
    # [CB]? 代表可選的 C 或 B；\d+ 代表數字
    regex = r'(([CB]?\d+[xX]\d+)|([CB]\d+))(?:mm|cm|px)?'
    match = re.search(regex, original_name, re.IGNORECASE)
    
    if match:
        raw_size = match.group(1).upper() # 取得匹配到的部分，如 C055 或 54X180
        
        # 🌟 自動擴展邏輯：如果抓到的是 C055 或 B060 這種縮寫
        if re.match(r'^[CB]\d+$', raw_size):
            prefix = raw_size[0]      # 'C' 或 'B'
            num = raw_size[1:]        # '055'
            standardized_size = f"{prefix}{num}X{num}MM" # 變成 C055X055MM
        else:
            # 如果已經有 X 了，就補上 MM 單位方便比對
            standardized_size = raw_size if "MM" in raw_size else raw_size + "MM"

        info["size"] = standardized_size
        info["clean_size"] = standardized_size.replace(" ", "")
        return info

    # B. PDF 偵測 (MediaBox)
    if ext == '.pdf':
        try:
            reader = PdfReader(file_path)
            mb = reader.pages[0].mediabox
            w_mm, h_mm = round(float(mb.width)/72*25.4), round(float(mb.height)/72*25.4)
            short, long = sorted([w_mm, h_mm])
            if 205 <= short <= 215 and 292 <= long <= 302: info["size"] = "A4"
            else: info["size"] = f"{w_mm}x{h_mm}mm"
        except: pass
        
    # C. JPG/PNG DPI 解析
    elif ext in ['.jpg', '.jpeg', '.png']:
        try:
            with Image.open(file_path) as img:
                w_px, h_px = img.size
                ratio = max(w_px, h_px) / min(w_px, h_px)
                if ratio >= 2.5: info["is_narrow"] = True
                
                dpi = img.info.get('dpi')
                if dpi and dpi[0] > 0:
                    w_cm = round((w_px / dpi[0]) * 2.54)
                    h_cm = round((h_px / dpi[1]) * 2.54)
                    info["size"] = f"{w_cm}x{h_cm}cm".upper()
        except: pass
    
    # 移除 PX, CM, MM 與空格，轉大寫
    clean_size = info["size"].upper().replace(" ", "")
    info["clean_size"] = clean_size
    return info
def create_drive_shortcut(service, target_id, target_name):
    """在人工確認資料夾建立檔案捷徑"""
    try:
        shortcut_metadata = {
            'name': f"[需確認]_{target_name}",
            'mimeType': 'application/vnd.google-apps.shortcut',
            'shortcutDetails': {'targetId': target_id},
            'parents': [MANUAL_REVIEW_FOLDER_ID]
        }
        service.files().create(body=shortcut_metadata, supportsAllDrives=True).execute()
        print(f"🔗 已在人工確認夾建立捷徑")
    except Exception as e:
        print(f"⚠️ 建立捷徑失敗: {e}")
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
    """自動建立年份子資料夾 (支援共用硬碟版)"""
    query = f"name = '{year}' and '{parent_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    
    # 這裡也要加上支援開關
    folders = service.files().list(
        q=query, 
        supportsAllDrives=True, 
        includeItemsFromAllDrives=True
    ).execute().get('files', [])
    
    if folders: return folders[0]['id']
    
    folder_meta = {'name': year, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    
    # 建立時也要加上 supportsAllDrives
    return service.files().create(
        body=folder_meta, 
        fields='id', 
        supportsAllDrives=True 
    ).execute().get('id')
    
def archive_and_rename(service, file_id, new_name, year):
    """【整合歸檔】重新命名 + 移至年份資料夾"""
    try:
        target_folder_id = get_or_create_year_folder(service, PROCESSED_ROOT_ID, year)
        file = service.files().get(
            fileId=file_id, 
            fields='parents',
            supportsAllDrives=True 
        ).execute()
        prev_parents = ",".join(file.get('parents'))
        service.files().update(
            fileId=file_id,
            body={'name': new_name},
            addParents=target_folder_id,
            removeParents=prev_parents,
            supportsAllDrives=True 
        ).execute()
        return True
    except Exception as e:
        print(f"❌ 歸檔失敗: {e}"); return False

def download_drive_file(service, file_id, file_name):
    request = service.files().get_media(
        fileId=file_id,
        supportsAllDrives=True 
    )
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
        # 1. 提取 Meta (年份、設備、更新日期)
        meta = get_essential_meta(service, f_id, orig_name)
        local_path = download_drive_file(service, f_id, orig_name)
        target_file = local_path
        if orig_name.lower().endswith('.ai'):
            target_file = convert_ai_to_pdf(local_path)
            is_temp_pdf = True

        # 2. 尺寸偵測
        phys = analyze_physical_size(target_file, orig_name)
        # 🌟 統一轉大寫去空格來比對
        current_size = phys.get("clean_size", "").upper().replace(" ", "")

        # 2. 硬性規則攔截
        fixed_purpose = None
        # 這裡用一個小迴圈來確保 "C055X055" 也能對到 "C055X055MM"
        for rule_size, rule_purpose in SIZE_RULE_MAP.items():
            if rule_size.upper().replace(" ", "") == current_size:
                fixed_purpose = rule_purpose
                print(f"⚖️ 【硬性規則命中】: 尺寸 {current_size} 自動判定為 {fixed_purpose}")
                break

        # 3. AI 視覺描述 (即便命中硬性規則，AI 還是可以幫忙判斷「活動名稱」)
        b64_imgs = get_images_for_ai(target_file)
        ai_res = ask_vision_ai_blind(b64_imgs, phys["size"], "HIDDEN")

        # 4. 決策
        final_activity = map_activity_name(orig_name, ai_res.get('活動判定'))
        
        # 如果硬性規則有抓到，優先使用；否則才用 map_purpose_name (含 AI 判定)
        if fixed_purpose:
            final_purpose = fixed_purpose
        else:
            final_purpose = map_purpose_name(orig_name, ai_res.get('用途判定'))
        # 5. 組合最終名稱
        base_name = f"{meta['year']}_{final_activity}_{phys['size']}_{final_purpose}_{meta['device']}"
        if meta["is_update"]:
            base_name += f"_{meta['modify_date']}"
        new_name = base_name + ext

        print(f"🤖 AI 視覺描述: {ai_res.get('視覺描述', 'None')}")
        print(f"📝 最終改名建議: {new_name}")

        # 6. 執行歸檔 (Google Drive 移動與改名)
        if archive_and_rename(service, f_id, new_name, meta['year']):
            print(f"✅ 歸檔成功：/{meta['year']}/{new_name}")
            if "不確定" in new_name:
                create_drive_shortcut(service, f_id, new_name)

            # ==========================================
            # 🌟 V7 新增：戰情室縮圖與 Log 同步
            # ==========================================
            thumb_filename = f"thumb_{int(time.time())}.jpg"
            thumb_save_path = f"dashboard/thumbnails/{thumb_filename}"
            os.makedirs(os.path.dirname(thumb_save_path), exist_ok=True)
            
            try:
                source_img = None
                if target_file.lower().endswith(".pdf"):
                    from pdf2image import convert_from_path
                    pdf_thumbs = convert_from_path(target_file, dpi=72, first_page=1, last_page=1)
                    if pdf_thumbs: source_img = pdf_thumbs[0]
                else:
                    source_img = Image.open(target_file)
                
                if source_img:
                    if source_img.mode != "RGB": source_img = source_img.convert("RGB")
                    source_img.thumbnail((300, 300))
                    source_img.save(thumb_save_path, "JPEG", quality=80, optimize=True)
                    source_img.close()
                    print(f"🖼️ 縮圖已生成：{thumb_save_path}")
            except Exception as e:
                print(f"❌ 縮圖生成失敗: {e}")
                thumb_filename = "default_error.jpg" 

            # 寫入日誌並推送到 GitHub
            log_to_github(orig_name, new_name, final_activity, final_purpose, phys['size'], f"thumbnails/{thumb_filename}")
            sync_to_github()

    except Exception as e:
        print(f"❌ 處理失敗: {e}")
        # 🌟 修正：發生錯誤時，也要把檔案移走，否則會無限循環
        archive_and_rename(service, f_id, f"[錯誤發生]_{orig_name}", "手動處理") 
        create_drive_shortcut(service, f_id, f"ERROR_{orig_name}")
    finally:
        if local_path and os.path.exists(local_path): os.remove(local_path)
        if is_temp_pdf: cleanup_temp_file(target_file)
    print("="*50)
    
import subprocess

def log_to_github(orig_name, new_name, activity, purpose, size, thumb_path):
    """將處理結果寫入 JSON 並準備同步到 GitHub"""
    log_entry = {
        "date": datetime.now().strftime("%Y-%m-%d"),
        "timestamp": datetime.now().strftime("%H:%M:%S"),
        "original_name": orig_name,
        "final_name": new_name,
        "activity": activity,
        "purpose": purpose,
        "size": size,
        "thumbnail": thumb_path
    }
    
    # 讀取並更新 log.json
    log_file = "dashboard/log.json"
    data = []
    if os.path.exists(log_file):
        with open(log_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    
    data.insert(0, log_entry) # 最新紀錄放在最前面
    
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def sync_to_github():
    """自動執行 Git 指令將 Log 推送到 GitHub"""
    try:
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", f"Log Update: {datetime.now().strftime('%Y-%m-%d %H:%M')}"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("🚀 GitHub 報表已更新！")
    except Exception as e:
        print(f"⚠️ Git 同步失敗: {e}")
        
# --- 找到 main() 裡面的這段並替換 ---
def main():
    service = get_drive_service()
    print(f"🚀 Tea Agent V6 (捷徑與版本追蹤版) 啟動...")
    while True:
        try:
            # 加上 supportsAllDrives 與 includeItemsFromAllDrives
            query = f"'{INBOX_FOLDER_ID}' in parents and trashed = false"
            fields = "files(id, name, mimeType, version, modifiedTime)"
            
            items = service.files().list(
                q=query, 
                fields=fields,
                supportsAllDrives=True,         # 🌟 關鍵：支援共用雲端硬碟
                includeItemsFromAllDrives=True  # 🌟 關鍵：包含共用雲端硬碟的項目
            ).execute().get('files', [])
            
            if items:
                for f in items:
                    if f['mimeType'] == 'application/vnd.google-apps.folder': continue
                    process_file(service, f)
            else:
                # 沒檔案時會印點點
                print(".", end="", flush=True)
            
            time.sleep(POLLING_INTERVAL_SECONDS) # 建議用變數，你設 30 秒
        except Exception as e:
            print(f"⚠️ 監聽中斷: {e}")
            time.sleep(60)

if __name__ == "__main__":
    main()