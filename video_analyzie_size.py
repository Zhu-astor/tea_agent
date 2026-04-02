import cv2
import os

def get_video_specs_cv2(video_path):
    try:
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            return "無法開啟影片檔案"

        # 取得寬高
        width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        
        # 取得幀率
        fps = round(cap.get(cv2.CAP_PROP_FPS))

        # 判斷方向
        orientation = "直式" if height > width else "橫式"
        max_side = max(width, height)
        
        # 解析度分類
        res_label = "4K" if max_side >= 3800 else \
                    "3K" if 2800 <= max_side <= 3200 else \
                    "1080p" if 1800 <= max_side <= 2000 else f"{max_side}p"
        
        fps_label = "60" if fps > 45 else "30"
        final_spec = f"{res_label}{fps_label}_{orientation}"
        
        cap.release()
        return {
            "檔案名稱": os.path.basename(video_path),
            "辨識結果": final_spec,
            "原始數據": f"{width}x{height} @ {fps}fps"
        }
    except Exception as e:
        return f"辨識失敗: {e}"

if __name__ == "__main__":
    # 測試
    res = get_video_specs_cv2(r"D:\Download\4762374-uhd_2160_4096_24fps.mp4")
    print(res)