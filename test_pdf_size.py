import os
from PyPDF2 import PdfReader

def get_pdf_physical_size(pdf_path):
    """讀取 PDF 底層的 MediaBox 來獲取絕對物理尺寸"""
    if not os.path.exists(pdf_path):
        print(f"❌ 找不到檔案: {pdf_path}")
        return

    try:
        reader = PdfReader(pdf_path)
        page = reader.pages[0] # 取第一頁
        
        # 獲取 MediaBox 的寬與高 (單位: point)
        # MediaBox 通常是一個陣列 [x0, y0, x1, y1]，寬度是 x1-x0，高度是 y1-y0
        mb = page.mediabox
        width_pt = float(mb.width)
        height_pt = float(mb.height)
        
        # 將 Point 轉換為公釐 (mm)
        # 公式: mm = (pt / 72) * 25.4
        width_mm = round((width_pt / 72) * 25.4, 1)
        height_mm = round((height_pt / 72) * 25.4, 1)
        
        print("="*50)
        print(f"📄 測試檔案: {os.path.basename(pdf_path)}")
        print(f"🔍 底層原始數據 (Point): {width_pt} x {height_pt} pt")
        print(f"📏 換算物理尺寸 (公釐): {width_mm} x {height_mm} mm")
        
        # 進行絕對物理尺寸比對 (允許 5mm 的出血/裁切誤差)
        short_edge = min(width_mm, height_mm)
        long_edge = max(width_mm, height_mm)
        
        print("\n🎯 系統精準判定結果:")
        if 205 <= short_edge <= 215 and 292 <= long_edge <= 302:
            print("👉 這絕對是【A4】(標準 210x297mm)")
        elif 495 <= short_edge <= 505 and 695 <= long_edge <= 705:
            print("👉 這絕對是【50x70cm 海報】(標準 500x700mm)")
        elif 385 <= short_edge <= 395 and 535 <= long_edge <= 545:
            print("👉 這絕對是【4K 海報】(標準 390x540mm)")
        else:
            print(f"👉 其他尺寸，長寬約為 {short_edge} x {long_edge} mm")
            
        print("="*50)

    except Exception as e:
        print(f"讀取 PDF 發生錯誤: {e}")

if __name__ == "__main__":
    # 請將這裡換成你要測試的 PDF 檔案路徑
    # (如果是 .ai 檔，請先手動複製一份並將副檔名改為 .pdf 來測試)
    TEST_FILE = "test.pdf" 
    get_pdf_physical_size(TEST_FILE)