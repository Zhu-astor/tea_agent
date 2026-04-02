import fitz  # PyMuPDF
import os

def deep_scan_ai(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"❌ 找不到檔案: {pdf_path}")
        return

    print(f"🚀 開始深層掃描: {os.path.basename(pdf_path)}")
    doc = fitz.open(pdf_path)
    page = doc[0]
    
    # 1. 取得畫布範圍
    view_rect = page.cropbox
    print(f"📦 畫布範圍 (CropBox): {view_rect}")
    print(f"📐 物理範圍 (MediaBox): {page.mediabox}")

    # 2. 掃描各類物件數量
    text_blocks = page.get_text("blocks")
    paths = page.get_drawings()
    images = page.get_image_info()
    
    print(f"--- 內容統計 ---")
    print(f"📝 文字區塊數: {len(text_blocks)}")
    print(f"🎨 向量路徑數: {len(paths)}")
    print(f"🖼️ 圖片物件數: {len(images)}")

    # 3. 尋找「躲在畫布外」的任何東西
    # 我們改用 get_bboxlog()，這能抓到所有渲染物件的邊界
    print(f"\n--- 邊界分析 ---")
    try:
        # 取得這一頁所有物件產生的總邊界
        full_bbox = page.get_bbox()
        print(f"📍 所有物件總邊界: {full_bbox}")
        
        if full_bbox.x0 < view_rect.x0 or full_bbox.y0 < view_rect.y0 or \
           full_bbox.x1 > view_rect.x1 or full_bbox.y1 > view_rect.y1:
            print("🚨 偵測到：確實有物件超出畫布範圍！")
        else:
            print("✅ 掃描：所有物件都在畫布內（或 PDF 層根本沒東西）。")
            
    except Exception as e:
        print(f"⚠️ 無法計算總邊界: {e}")

    doc.close()

# 🌟 確保這段在檔案最下方，否則 python 執行時會沒反應
if __name__ == "__main__":
    # 請替換成你實際的路徑
    test_path = r"D:\Download\【用途】-20260311T054848Z-1-001\【用途】\2025_茶王_54x180mm_名片小卡_電腦＿設計檔.ai"
    deep_scan_ai(test_path)