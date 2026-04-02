import subprocess
import os
import json
import base64
from datetime import datetime
import time

# --- 配置區 ---
# 確保系統 PATH 裡有 ffmpeg 可執行檔
FFMPEG_PATH = r"C:\Users\bubbl\anaconda3\envs\video_tool\Library\bin\ffmpeg.exe"
# 暫存截圖的資料夾
TEMP_FRAME_DIR = "temp_video_frames"

if not os.path.exists(TEMP_FRAME_DIR):
    os.makedirs(TEMP_FRAME_DIR)

# ==========================================
#         第一部分：影音工程核心 (零下載抽幀)
# ==========================================

def get_video_duration_gdrive(service, file_id):
    """
    透過 Google Drive API 直接獲取影片總長度，完全不需讀取檔案本體。
    成本：0。耗時：極低。
    """
    try:
        file_metadata = service.files().get(
            fileId=file_id, 
            fields='videoMediaMetadata'
        ).execute()
        
        # GDrive 回傳的時長單位是毫秒 (milliseconds)
        duration_ms = int(file_metadata.get('videoMediaMetadata', {}).get('durationMillis', 0))
        duration_sec = duration_ms / 1000.0
        
        if duration_sec == 0:
            print(f"⚠️ 無法從 API 取得影片長度，可能非標準影片格式或 GDrive 尚未處理完成。")
            return None
            
        print(f"📹 偵測到影片總長度: {duration_sec:.2f} 秒")
        return duration_sec
        
    except Exception as e:
        print(f"❌ 取得影片 Metadata 失敗: {e}")
        return None

def extract_keyframes_cloud_seek(file_id, access_token, duration_sec):
    """
    【核心黑科技】利用 FFmpeg 的 HTTP Range Request 特性。
    透過網址跳躍式讀取數據，只下載頭、中、尾三個時間點的定格畫面。
    """
    if not duration_sec or duration_sec < 2:
        print("❌ 影片過短或無長度資訊，跳過抽幀。")
        return []

    # 建構 Google Drive 原始下載網址 (FFmpeg 需要此網址)
    video_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"

    # 建構帶著授權 Token 的 Header (FFmpeg 支援 HTTP Header)
    # 這是為了讓 FFmpeg 能代表你向 GDrive 請求數據
    ffmpeg_headers = f"Authorization: Bearer {access_token}"

    # 計算三個關鍵時間點 (t1, t2, t3)
    # 規則：跳過開頭 1 秒(防黑屏)，中間，結尾前 2 秒(防謝幕黑屏)
    t1 = max(1.0, duration_sec * 0.1) # 開頭 10% 或 1 秒
    t2 = duration_sec * 0.5            # 中間 50%
    t3 = max(duration_sec - 2.0, duration_sec * 0.9) # 結尾前 2 秒或 90%
    
    seek_times = [t1, t2, t3]
    time_labels = ["開頭", "中間", "結尾"]
    extracted_frames_b64 = []

    print(f"🚀 啟動「雲端跳躍抽幀」技術，目標時間點: {[f'{t:.2f}s' for t in seek_times]}")
    start_time = time.time()

    for i, seek_t in enumerate(seek_times):
        label = time_labels[i]
        timestamp_str = datetime.now().strftime("%H%M%S")
        out_filename = os.path.join(TEMP_FRAME_DIR, f"{file_id}_frame_{i}_{timestamp_str}.jpg")
        
        # --- 核心 FFmpeg 指令辯論與實作 ---
        # 為什麼高效？
        # 1. -ss 放在 -i 之前：這是「輸入 Seek」，FFmpeg 會利用 Range Request 
        #    告訴 GDrive 伺服器：「我要從這個 Byte 開始讀」，直接跳過前面的數據。
        # 2. -headers: 傳遞 Token 進行認證。
        # 3. -frames:v 1: 只抓一張圖。
        cmd = [
            FFMPEG_PATH,
            '-y',               # 覆蓋輸出檔案
            '-ss', str(seek_t),  # 【關鍵】在輸入前 Seek，啟用 Range Request
            '-headers', ffmpeg_headers,
            '-i', video_url,    # 輸入網址
            '-frames:v', '1',   # 影片流只取 1 幀
            '-q:v', '2',        # 圖片質量 (2-31, 2很好)
            '-f', 'image2',     # 輸出格式為圖片
            '-update', '1',
            out_filename
        ]
        
        try:
            # 執行 FFmpeg，隱藏標準輸出，只顯示錯誤
            print(f"-> 正在擷取 [{label}] 畫面 ({seek_t:.2f}s)...")
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
            
            # 檢查檔案是否存在且大於 0 (防撞)
            if os.path.exists(out_filename) and os.path.getsize(out_filename) > 0:
                # 將圖片轉為 Base64 字符串
                with open(out_filename, "rb") as img_f:
                    b64_str = base64.b64encode(img_f.read()).decode('utf-8')
                    extracted_frames_b64.append(b64_str)
                
                # 【節省硬碟】轉完 Base64 後立刻刪除本地圖片
                os.remove(out_filename) 
            else:
                print(f"⚠️ 擷取 [{label}] 失敗，產出空檔案。")
                
        except subprocess.CalledProcessError as e:
            # FFmpeg 經常會報一些關於 network packet 的無關緊要錯誤，只要有產出圖就好
            # 這裡我們選擇記錄錯誤，但不中斷程式
            # print(f"❌ FFmpeg 執行出錯 (Seek: {seek_t}s): {e.stderr.decode('utf-8')}")
            # 有時 Seek 太靠後會失敗，我們補救：如果檔案存在就認了
            if os.path.exists(out_filename) and os.path.getsize(out_filename) > 0:
                 with open(out_filename, "rb") as img_f:
                    b64_str = base64.b64encode(img_f.read()).decode('utf-8')
                    extracted_frames_b64.append(b64_str)
                 os.remove(out_filename)
            else:
                 print(f"❌ 擷取 [{label}] 徹底失敗。")

    end_time = time.time()
    print(f"✅ 抽幀完成。共取得 {len(extracted_frames_b64)} 張關鍵畫面。耗時: {end_time - start_time:.2f} 秒 (「零下載」達成)")
    return extracted_frames_b64

# ==========================================
#         第二部分：AI 決策核心 (視覺判斷)
# ==========================================

import requests # 記得在最上方 import requests

# --- 請確保這些配置與你主程式一致 ---
OPENROUTER_API_KEY = "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"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
# 建議統一使用支援視覺與結構化輸出的模型
VISION_MODEL_NAME = "openai/gpt-4o-mini" 
TEXT_MODEL_NAME = "openai/gpt-4o-mini" 

def ask_vision_ai_to_analyze_video(base64_frames, original_filename):
    """
    將三張關鍵截圖打包送往 AI，分析影片的故事線、活動與用途
    """
    if not base64_frames:
        return {"畫面描述": "無法取得影片截圖", "活動名稱": "其他", "用途": "未分類素材"}

    print(f"🤖 正在呼叫 Vision AI 分析影片故事線 (共 {len(base64_frames)} 張截圖)...")

    prompt = f"""你是一個專業的影片素材分析師。
我們從一支影片中擷取了【開頭】、【中間】、【結尾】三張定格截圖。
原始檔案名稱為：【{original_filename}】

任務 1：內容觀察 (分析故事線)
請綜合這三張截圖的順序變化，描述影片的內容(務必精確，如果沒有則填寫「無明確內容」)。
（例如：起初是產品靜態特寫，隨後轉為人員倒茶動作，最後出現品牌 Logo 與價格字卡）。

任務 2：判定【活動名稱】
請根據檔名與視覺文字，從清單選一輸出：[一年免費喝, 年節禮, 中秋, 周年慶, 其他]。
💡 提示：若檔名已有明確暗示，請優先採納。

任務 3：判定【用途】
請結合畫面排版與檔名做出判定：
- 行銷短影音：畫面有經過剪輯、帶有字卡、特效、背景音樂感強、或是排版精美的直式/橫式影片。
- 無人機素材：鏡頭平穩、高空俯瞰風景、無字卡。
- 原始相機素材：生活紀錄、無字卡、畫面較為隨性。
- 1a(機器人)：若檔名明確出現 '1a' 或機器人相關字眼。

請嚴格輸出 JSON 格式：
{{
  "畫面描述": "簡述故事線",
  "活動名稱": "上述清單選項",
  "用途": "上述清單選項"
}}"""

    # 封裝多張圖片到消息體中
    content_list = [{"type": "text", "text": prompt}]
    for i, b64_img in enumerate(base64_frames):
        content_list.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{b64_img}"}
        })

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "model": VISION_MODEL_NAME,
        "messages": [{"role": "user", "content": content_list}],
        "temperature": 0.1 # 降低隨機性，確保分類穩定
    }

    try:
        response = requests.post(OPENROUTER_URL, headers=headers, json=payload)
        response.raise_for_status()
        
        # 解析 AI 回傳的 JSON 字串
        raw_content = response.json()['choices'][0]['message']['content']
        # 清理可能夾雜的 Markdown 標籤
        clean_json = raw_content.replace("```json", "").replace("```", "").strip()
        result = json.loads(clean_json)
        
        print(f"✅ AI 判定成功！用途：{result.get('用途')}")
        return result
    except Exception as e:
        print(f"❌ AI 辨識出錯: {e}")
        return {"畫面描述": f"辨識出錯: {str(e)}", "活動名稱": "其他", "用途": "辨識失敗"}

# 請確保在檔案最上方有 import 這些套件
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os

if __name__ == "__main__":
    print("=== 影片雲端Seek抽幀模組 測試啟動 ===")
    
    # 🔑 1. 自動讀取你資料夾裡的 token.json 來取得 Access Token
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
        TEST_ACCESS_TOKEN = creds.token
        print("✅ 成功從 token.json 讀取授權金鑰！")
    else:
        print("❌ 找不到 token.json，請確認你跟主程式放在同一個資料夾。")
        exit()

    # 🔑 2. 建立真實的 Google Drive 服務連線
    service = build('drive', 'v3', credentials=creds)

    # 🎯 3. 填入你剛剛複製出來的 File ID
    TEST_FILE_ID = "15wHy-oKsOO0-HH9Eev84RHVLDyLIfgNK" 
    TEST_FILENAME = "測試用影片.mp4"

    # --- 開始真實測試 ---
    # 步驟一：取得時長
    duration = get_video_duration_gdrive(service, TEST_FILE_ID)
    
    if duration:
        # 步驟二：雲端跳躍抽幀
        b64_frames = extract_keyframes_cloud_seek(TEST_FILE_ID, TEST_ACCESS_TOKEN, duration)
        
        if b64_frames:
            # 🌟 步驟三：呼叫 AI 進行真正的內容分析
            ai_result = ask_vision_ai_to_analyze_video(b64_frames, TEST_FILENAME)
            
            print("\n" + "="*50)
            print("🏁 【影片歸檔建議報告】")
            print(f"原始檔名: {TEST_FILENAME}")
            print(f"AI 描述: {ai_result['畫面描述']}")
            print(f"判定活動: {ai_result['活動名稱']}")
            print(f"判定用途: {ai_result['用途']}")
            print("="*50)