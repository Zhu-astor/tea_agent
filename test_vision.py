import os
import base64
import requests
import json
from datetime import datetime
from io import BytesIO
from PIL import Image
from pdf2image import convert_from_path

# 引入剛剛寫好的轉換模組
from ai_converter import convert_ai_to_pdf, cleanup_temp_file

# ==========================================
# 1. 專案環境設定
# ==========================================
OPENROUTER_API_KEY = "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"
MODEL_NAME = "openai/gpt-4o-mini"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# 事先設定好的常數
PRESET_YEAR = "2026"
PRESET_DEVICE = "電腦"

# 這次請換成你要測試的 .ai 檔案路徑！
TEST_FILE_PATH = r"C:\Users\bubbl\Desktop\Intership\Lovecloud\Work6_Tea_agent\material\test8.ai"

# 用途完整清單
PURPOSE_LIST = [
    "1a(機器人)", "banner(機器人)", "六宮格選單(機器人)", "名片小卡", "店面", 
    "社群平台", "公告", "立架", "布條", "酷卡", "貼紙", "桌面大圖/柱子", 
    "網站", "包裝/禮盒", "兌換券/折價券", "茶包", "dm菜單", "三折"
]

# ==========================================
# 2. 核心邏輯功能
# ==========================================

def get_image_and_dimensions(file_path):
    """讀取檔案 (PDF/圖片)、轉成 Base64 並計算尺寸"""
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".pdf":
        print("-> 讀取 PDF 中，正在抽取第一頁...")
        pages = convert_from_path(file_path, first_page=1, last_page=1)
        img = pages[0]
    elif ext in [".jpg", ".jpeg", ".png"]:
        img = Image.open(file_path)
    else:
        raise ValueError(f"不支援的檔案格式: {ext}")

    width, height = img.size
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
        elif 1.38 <= ratio <= 1.43: 
            size_str = "A4_或_50x70_或_4k"
            purpose_str = "請AI判斷" 

    buffered = BytesIO()
    img.convert("RGB").save(buffered, format="JPEG")
    img_base64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
    
    return img_base64, size_str, purpose_str

# def ask_vision_ai(img_base64, size_str, purpose_str):
#     """呼叫 OpenRouter 模型進行辨識"""
    
#     purpose_prompt = f"這張圖的用途已經被系統確認為：【{purpose_str}】，請直接在 JSON 的用途欄位輸出「{purpose_str}」。"
#     if purpose_str == "請AI判斷" or purpose_str is None:
#         purpose_prompt = f"系統無法自動判斷用途 (目前的尺寸特徵為 {size_str})。請觀察圖片的排版與內容，從以下清單選出一個最適合的用途：{PURPOSE_LIST}"

#     prompt = f"""你是一個精準的檔案分類系統。請觀察這張圖片，完成以下兩項任務：

# 任務 1：判斷【活動名稱】
# 請仔細閱讀圖片上的文字，並「嚴格遵守」以下對應規則進行輸出：
# - 若畫面包含「跨年」、「共享尾牙」、「一年免費喝」 -> 輸出：一年免費喝
# - 若畫面包含「蟬吃茶年節禮」、「春節禮」、「蟬茶禮組」 -> 輸出：年節禮
# - 若畫面包含「母親節特惠組」 -> 輸出：母親節
# - 若畫面包含「父親節」、「父親節禮組」、「88節」 -> 輸出：父親節
# - 若畫面包含「中秋好禮」、「慶團圓」、「中秋節禮」 -> 輸出：中秋
# - 若畫面包含「端午特惠」 -> 輸出：端午
# - 若畫面包含「周年慶」 -> 輸出：周年慶
# - 若以上皆非，或無法辨識 -> 輸出：日常活動

# 任務 2：判斷【用途】
# {purpose_prompt}

# 請嚴格以 JSON 格式輸出，不要包含任何 Markdown 標記或其他說明文字。格式如下：
# {{"活動名稱": "填入選項", "用途": "填入選項"}}
# """

#     headers = {
#         "Authorization": f"Bearer {OPENROUTER_API_KEY}",
#         "Content-Type": "application/json"
#     }

#     payload = {
#         "model": MODEL_NAME, # 暫時改用 Llama 確保連線正常
#         "messages": [
#             {
#                 "role": "user",
#                 "content": [
#                     {"type": "text", "text": prompt},
#                     {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}}
#                 ]
#             }
#         ],
#         "temperature": 0.1 
#     }

#     print("-> 正在呼叫 AI 視覺模型分析圖片內容...")
#     response = requests.post(OPENROUTER_URL, headers=headers, json=payload)
    
#     if response.status_code != 200:
#         print(f"❌ API 呼叫失敗！HTTP 狀態碼：{response.status_code}")
#         print(f"❌ 詳細錯誤訊息：{response.text}")
#     response.raise_for_status()
    
#     result_text = response.json()['choices'][0]['message']['content']
#     result_text = result_text.replace("```json", "").replace("```", "").strip()
    
#     try:
#         return json.loads(result_text)
#     except json.JSONDecodeError:
#         print(f"解析錯誤！模型原始輸出為：{result_text}")
#         return {"活動名稱": "辨識失敗", "用途": "辨識失敗"}

def ask_vision_ai(img_base64, size_str, purpose_str):
    """呼叫 OpenRouter 模型進行辨識"""
    
    purpose_prompt = f"這張圖的用途已經被系統確認為：【{purpose_str}】，請直接在 JSON 的用途欄位輸出「{purpose_str}」。"
    if purpose_str == "請AI判斷" or purpose_str is None:
        purpose_prompt = f"系統無法自動判斷用途 (目前的尺寸特徵為 {size_str})。請觀察圖片的排版與內容，從以下清單選出一個最適合的用途：{PURPOSE_LIST}"

    # 【修改重點 1】在 Prompt 中加入「畫面觀察」的任務
    prompt = f"""你是一個精準的檔案分類系統。請觀察這張圖片，完成以下三項任務：

任務 1：畫面觀察與描述
請先仔細閱讀圖片上的「文字內容」，並觀察主要的「視覺元素與排版特徵」，將這些觀察記錄下來，文字也請記錄下來(品牌、產品、價格等)。

任務 2：判斷【活動名稱】
請仔細閱讀圖片上的文字，並「嚴格遵守」以下對應規則進行輸出：
- 若畫面包含「跨年」、「共享尾牙」、「一年免費喝」 -> 輸出：一年免費喝
- 若畫面包含「蟬吃茶年節禮」、「春節禮」、「蟬茶禮組」 -> 輸出：年節禮
- 若畫面包含「母親節特惠組」 -> 輸出：母親節
- 若畫面包含「父親節」、「父親節禮組」、「88節」 -> 輸出：父親節
- 若畫面包含「中秋好禮」、「慶團圓」、「中秋節禮」 -> 輸出：中秋
- 若畫面包含「端午特惠」 -> 輸出：端午
- 若畫面包含「周年慶」 -> 輸出：周年慶
- 若以上皆非，或無法辨識 -> 輸出：日常活動

任務 3：判斷【用途】
{purpose_prompt}

請嚴格以 JSON 格式輸出，不要包含任何 Markdown 標記或其他說明文字。格式如下：
{{
  "畫面描述": "簡述你看到的關鍵文字與視覺特徵",
  "活動名稱": "填入選項",
  "用途": "填入選項"
}}
"""

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": MODEL_NAME, # 或替換成你正在測試的模型
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}}
                ]
            }
        ],
        "temperature": 0.1 
    }

    print("-> 正在呼叫 AI 視覺模型分析圖片內容...")
    response = requests.post(OPENROUTER_URL, headers=headers, json=payload)
    
    if response.status_code != 200:
        print(f"❌ API 呼叫失敗！HTTP 狀態碼：{response.status_code}")
        print(f"❌ 詳細錯誤訊息：{response.text}")
    response.raise_for_status()
    
    result_text = response.json()['choices'][0]['message']['content']
    result_text = result_text.replace("```json", "").replace("```", "").strip()
    
    try:
        return json.loads(result_text)
    except json.JSONDecodeError:
        print(f"解析錯誤！模型原始輸出為：{result_text}")
        return {"畫面描述": "解析失敗", "活動名稱": "辨識失敗", "用途": "辨識失敗"}


def main():
    if not os.path.exists(TEST_FILE_PATH):
        print(f"找不到測試檔案：{TEST_FILE_PATH}，請確認路徑與檔名。")
        return

    print(f"開始處理檔案: {TEST_FILE_PATH}")

    # ==========================================
    # 【新增邏輯】攔截 .ai 檔案並進行預先轉換
    # ==========================================
    target_file = TEST_FILE_PATH
    is_temp_pdf = False
    
    if TEST_FILE_PATH.lower().endswith('.ai'):
        target_file = convert_ai_to_pdf(TEST_FILE_PATH)
        is_temp_pdf = True

    try:
        # 1. 取得尺寸與初步用途 (Python 負責，讀取轉換後的 target_file)
        img_base64, size_str, purpose_str = get_image_and_dimensions(target_file)
        print(f"-> Python 幾何分析結果: 尺寸=[{size_str}], 用途初步判定=[{purpose_str or '交給AI判斷'}]")

        # 2. 呼叫 AI 進行語意辨識 (AI 負責)
        ai_result = ask_vision_ai(img_base64, size_str, purpose_str)
        
        # 【新增這行】將 AI 的觀察內容印出來
        print(f"👀 AI 觀察到的畫面內容: {ai_result.get('畫面描述', '無紀錄')}")
        
        activity = ai_result.get("活動名稱", "日常活動")
        final_purpose = ai_result.get("用途", "未知用途")
        print(f"-> AI 視覺分析結果: 活動=[{activity}], 用途=[{final_purpose}]")

        # 3. 取得原始檔案最後修改日期 (讀取原始的 TEST_FILE_PATH)
        mtime = os.path.getmtime(TEST_FILE_PATH)
        date_str = datetime.fromtimestamp(mtime).strftime("%Y%m%d")

        # 4. 組合最終檔名：[年份]_[活動名稱]_[尺寸]_[用途]_[設備]_[日期]
        safe_size = size_str.replace(" ", "")
        final_filename = f"{PRESET_YEAR}_{activity}_{safe_size}_{final_purpose}_{PRESET_DEVICE}_{date_str}"
        
        # 副檔名依然要保留原本的 .ai
        original_ext = os.path.splitext(TEST_FILE_PATH)[1]
        
        print("\n" + "="*60)
        print(f"✅ 原始檔名: {os.path.basename(TEST_FILE_PATH)}")
        print(f"🎯 建議重新命名: {final_filename}{original_ext}")
        print("="*60)

    finally:
        # 確保無論程式成功或報錯，都會把暫存的 PDF 刪掉，不留垃圾
        if is_temp_pdf:
            cleanup_temp_file(target_file)

if __name__ == "__main__":
    main()