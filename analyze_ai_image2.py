import win32com.client
import os
import time

def export_ai_to_png_safe(file_path):
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
    except:
        print("無法啟動 Illustrator")
        return

    # 強制關閉 AI 的彈出視窗（如：缺字體警告），避免腳本卡死
    # -1 代表 aiNoUserInteraction
    app.UserInteractionLevel = -1 

    abs_path = os.path.abspath(file_path)
    doc = app.Open(abs_path)
    time.sleep(2) # 給 AI 充足時間處理檔案內容

    try:
        # 1. 確保所有圖層都解鎖並顯示 (防止無法選取)
        for i in range(1, doc.Layers.Count + 1):
            layer = doc.Layers.Item(i)
            layer.Locked = False
            layer.Visible = True

        # 2. 執行全選
        app.ExecuteMenuCommand("selectall")
        
        # 檢查是否有東西被選取
        if doc.Selection:
            print(f"成功選取 {len(doc.Selection)} 個物件")
            # 嘗試調整工作區域，如果失敗則跳過，不影響導出
            try:
                app.ExecuteMenuCommand("fitArtboardToSelectedArt")
                print("已調整工作區域")
            except:
                print("⚠️ 無法調整工作區域，將維持原尺寸導出")
        else:
            print("⚠️ 警告：全選後未發現任何物件，可能內容都在剪裁遮罩或隱藏層內")

        # 3. 設定 PNG 導出選項
        export_options = win32com.client.Dispatch("Illustrator.ExportOptionsPNG24")
        export_options.AntiAliasing = True
        export_options.Transparency = False
        export_options.HorizontalScale = 400.0 # 提高到 400% 確保 OCR 準確度
        export_options.VerticalScale = 400.0

        # 4. 執行導出
        png_path = abs_path.replace(".ai", "_final_export.png")
        doc.Export(png_path, 5, export_options)
        print(f"✅ 導出成功：{png_path}")

    except Exception as e:
        print(f"❌ 執行中發生錯誤: {e}")

    finally:
        # 恢復互動層級，不然以後你手動用 AI 時它不會跳警告
        app.UserInteractionLevel = 1 
        doc.Close(2)
        print("檔案已關閉")

# --- 執行 ---
target_ai = r"D:\Download\【用途】-20260311T054848Z-1-001\【用途】\2025_茶王_54x180mm_名片小卡_電腦＿設計檔.ai"
export_ai_to_png_safe(target_ai)