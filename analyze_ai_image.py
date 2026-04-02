import win32com.client
import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def auto_restore_and_visualize(file_path):
    # 1. 初始化 Illustrator
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
    except:
        print("請確保 Illustrator 已開啟")
        return

    abs_path = os.path.abspath(file_path)
    doc = app.Open(abs_path)
    print(f"成功開啟檔案：{doc.Name}")

    # 2. 準備繪圖畫布
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # 3. 抓取所有路徑座標並繪圖
    print("正在處理路徑物件，請稍候...")
    all_items = doc.PathItems
    count = all_items.Count
    
    for i in range(1, count + 1):
        try:
            item = all_items.Item(i)
            bounds = item.GeometricBounds # (左, 上, 右, 下)
            
            x1, y1, x2, y2 = bounds
            width = x2 - x1
            height = y2 - y1 # 注意：AI 的 Y 軸通常向下為負
            
            # 建立矩形方塊代表該路徑
            rect = patches.Rectangle((x1, y1), width, height, linewidth=0, facecolor='black', alpha=0.7)
            ax.add_patch(rect)
            
            # 為了效能，每 500 個物件印一次進度
            if i % 500 == 0:
                print(f"已處理 {i}/{count} 個物件...")
        except:
            continue

    # 4. 設定畫布格式
    ax.set_aspect('equal')
    ax.autoscale_view()
    plt.axis('off') # 隱藏座標軸以便 OCR 辨識
    
    # 5. 儲存還原後的圖像
    output_img = "restored_view.png"
    plt.savefig(output_img, dpi=300, bbox_inches='tight')
    print(f"\n✅ 視覺還原完成！請查看：{os.path.abspath(output_img)}")
    
    # 6. 關閉文件 (不儲存)
    doc.Close(2) 

# --- 執行處 ---
target_ai = r"D:\Download\【用途】-20260311T054848Z-1-001\【用途】\2025_茶王_54x180mm_名片小卡_電腦＿設計檔.ai"
auto_restore_and_visualize(target_ai)