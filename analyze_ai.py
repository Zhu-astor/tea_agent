# import win32com.client
# import os

# def analyze_illustrator_file(file_path):
#     if not os.path.exists(file_path):
#         print(f"錯誤：找不到檔案 {file_path}")
#         return

#     # 初始化 Illustrator
#     try:
#         ai_app = win32com.client.Dispatch("Illustrator.Application")
#     except Exception as e:
#         print(f"無法啟動 Illustrator: {e}")
#         return

#     # 隱藏 UI 開啟檔案 (可視需求調整)
#     # 註：Illustrator 有時仍會跳出視窗，但在腳本執行下通常很快
#     doc = ai_app.Open(file_path)
    
#     print(f"--- 正在分析檔案: {doc.Name} ---")
    
#     # 1. 取得所有工作區域的邊界
#     # Illustrator 的座標格式為 [left, top, right, bottom]
#     artboards_bounds = []
#     for i in range(1, doc.Artboards.Count + 1):
#         ab = doc.Artboards.Item(i)
#         artboards_bounds.append(ab.ArtboardRect)
#         print(f"工作區域 {i} 範圍: {ab.ArtboardRect}")

#     # 2. 定義判斷物件是否在工作區域內的函數
#     def is_inside_any_artboard(item_bounds, ab_rects):
#         i_left, i_top, i_right, i_bottom = item_bounds
#         for ab_rect in ab_rects:
#             a_left, a_top, a_right, a_bottom = ab_rect
            
#             # 判斷是否「完全不在」該工作區域內 (碰撞偵測反向邏輯)
#             if not (i_right < a_left or i_left > a_right or i_top < a_bottom or i_bottom > a_top):
#                 return True
#         return False

#     # 3. 遍歷所有頁面物件 (PageItems)
#     all_items = doc.PageItems
#     inside_count = 0
#     outside_items = []

#     print(f"檔案內共有 {all_items.Count} 個物件。")

#     for j in range(1, all_items.Count + 1):
#         item = all_items.Item(j)
#         try:
#             bounds = item.GeometricBounds # [left, top, right, bottom]
#             item_info = {
#                 "name": item.Name if item.Name else "未命名物件",
#                 "type": item.typename,
#                 "bounds": bounds
#             }

#             if is_inside_any_artboard(bounds, artboards_bounds):
#                 inside_count += 1
#             else:
#                 outside_items.append(item_info)
#         except:
#             continue

#     # 4. 輸出結果
#     print(f"\n分析結果：")
#     print(f"✅ 在工作區域內的物件數量: {inside_count}")
#     print(f"❌ 在工作區域外的物件數量: {len(outside_items)}")
    
#     if outside_items:
#         print("\n--- 區域外物件清單 ---")
#         for idx, item in enumerate(outside_items, 1):
#             print(f"{idx}. [{item['type']}] {item['name']} - 座標: {item['bounds']}")

#     # 關閉檔案 (不儲存)
#     # doc.Close(2) # 2 代表 aiDoNotSaveChanges

# if __name__ == "__main__":
#     path = input("請貼上 .ai 檔案的完整路徑: ").strip('"')
#     analyze_illustrator_file(path)

import win32com.client
import os

def extract_data_from_ai_file(file_path):
    # 1. 建立與 Illustrator 的連線
    # 注意：電腦必須安裝 Illustrator 且此腳本目前僅支援 Windows
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
    except Exception as e:
        print(f"無法啟動 Illustrator: {e}")
        return

    # 2. 轉換為絕對路徑 (Illustrator 需要完整路徑)
    abs_path = os.path.abspath(file_path)
    
    if not os.path.exists(abs_path):
        print(f"找不到檔案: {abs_path}")
        return

    # 3. 開啟檔案
    print(f"正在開啟檔案: {abs_path} ...")
    doc = app.Open(abs_path)

    # 4. 提取文字內容 (TextFrames)
    print("\n--- 文字內容 ---")
    if doc.TextFrames.Count > 0:
        for i in range(1, doc.TextFrames.Count + 1):
            tf = doc.TextFrames.Item(i)
            # tf.Contents 是文字內容
            # tf.GeometricBounds 是座標 (左, 上, 右, 下)
            print(f"{i}. 內容: {tf.Contents}")
            print(f"   座標: {tf.GeometricBounds}")
    else:
        print("檔案中沒有可讀取的文字物件（可能已被轉外框）。")

    # 5. 提取所有路徑物件 (PathItems) 的座標
    # 如果文字被「建立外框」，它們會出現在這裡
    print("\n--- 路徑物件座標 (前 10 個範例) ---")
    limit = min(doc.PathItems.Count, 10) 
    for j in range(1, limit + 1):
        pi = doc.PathItems.Item(j)
        print(f"{j}. [PathItem] 座標: {pi.GeometricBounds}")

    # 6. (選用) 關閉檔案而不儲存
    # doc.Close(2) # 2 代表 aiDoNotSaveChanges
    # print("\n檔案已關閉。")

# --- 執行處 ---
# 請將下方的路徑換成你實際的 .ai 檔案路徑
my_file = r"D:\Download\【用途】-20260311T054848Z-1-001\【用途】\2025_茶王_54x180mm_名片小卡_電腦＿設計檔.ai"
extract_data_from_ai_file(my_file)