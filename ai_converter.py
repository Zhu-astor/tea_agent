import os
import shutil

def convert_ai_to_pdf(ai_file_path):
    """
    將 .ai 檔案轉換為 .pdf 格式
    原理：利用 Illustrator 預設的 PDF 相容層，複製並更改副檔名即可讓 pdf2image 讀取
    """
    if not os.path.exists(ai_file_path):
        raise FileNotFoundError(f"找不到檔案: {ai_file_path}")
        
    # 建立暫存的 PDF 檔名 (例如: test_temp_converted.pdf)
    base_name = os.path.splitext(ai_file_path)[0]
    temp_pdf_path = f"{base_name}_temp_converted.pdf"
    
    # 複製檔案並重新命名為 .pdf
    shutil.copy2(ai_file_path, temp_pdf_path)
    print(f"🔧 [轉換模組] 偵測到 .ai 工作檔！已暫時轉存為 {os.path.basename(temp_pdf_path)} 以供解析")
    
    return temp_pdf_path

def cleanup_temp_file(temp_file_path):
    """解析完成後，刪除暫存的 PDF 檔案"""
    if os.path.exists(temp_file_path):
        os.remove(temp_file_path)
        print(f"🧹 [清理模組] 已刪除暫存檔: {os.path.basename(temp_file_path)}")