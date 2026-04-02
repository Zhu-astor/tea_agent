import os
import subprocess

def convert_ai_to_pdf(ai_file_path):
    """
    [方案一：Inkscape 引擎版]
    將 .ai 檔案轉換為 .pdf 格式，並強制包含工作區域外的所有物件。
    """
    if not os.path.exists(ai_file_path):
        raise FileNotFoundError(f"找不到檔案: {ai_file_path}")
        
    base_name = os.path.splitext(ai_file_path)[0]
    # 輸出的暫存檔路徑
    temp_pdf_path = f"{base_name}_temp_full_content.pdf"
    
    # Inkscape 指令說明：
    # --export-type=pdf: 指定輸出格式為 PDF
    # --export-area-drawing: 重要！這會抓取「所有物件」的範圍，忽略原本畫布的裁切框
    # --export-filename: 指定輸出的路徑
    inkscape_path = r"D:\Inkscape\bin\inkscape.exe"
    
    inkscape_cmd = [
        inkscape_path,
        ai_file_path, 
        "--export-type=pdf", 
        "--export-area-drawing", 
        "--export-filename", temp_pdf_path
    ]
    
    try:
        print(f"🚀 [轉換模組] 正在透過 Inkscape 解析 {os.path.basename(ai_file_path)} (包含工作區域外內容)...")
        
        # 執行指令
        result = subprocess.run(inkscape_cmd, capture_output=True, text=True, check=True)
        
        print(f"✅ [轉換模組] 解析成功！已生成完整的 PDF 以供 AI 辨識")
        return temp_pdf_path

    except subprocess.CalledProcessError as e:
        print(f"❌ [錯誤] Inkscape 轉換失敗: {e.stderr}")
        raise RuntimeError("請確保系統已安裝 Inkscape 並且已加入環境變數 (Path)")
    except FileNotFoundError:
        print(f"❌ [錯誤] 找不到 Inkscape 執行檔。請確認電腦已安裝 Inkscape。")
        raise

def cleanup_temp_file(temp_file_path):
    """解析完成後，刪除暫存的 PDF 檔案"""
    if os.path.exists(temp_file_path):
        os.remove(temp_file_path)
        print(f"🧹 [清理模組] 已刪除暫存檔: {os.path.basename(temp_file_path)}")