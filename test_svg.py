import os
import subprocess
import re

def brute_force_reveal_svg(svg_path):
    """暴力解鎖 SVG：移除遮罩並擴張畫布"""
    with open(svg_path, 'r', encoding='utf-8') as f:
        content = f.read()

    print("🔓 正在執行暴力解鎖：移除剪裁遮罩...")
    
    # 1. 移除所有 clip-path 屬性
    content = re.sub(r'clip-path="url\(#.*?\)"', '', content)
    # 2. 移除所有 clip-rule
    content = re.sub(r'clip-rule=".*?"', '', content)
    
    # 3. 強行把 viewBox 改成一個極大的範圍 (例如 -5000 到 10000)
    # 這樣不管東西在哪裡，理論上都能呈現在畫面上
    content = re.sub(r'viewBox=".*?"', 'viewBox="-5000 -5000 15000 15000"', content)
    
    unmasked_path = svg_path.replace(".svg", "_revealed.svg")
    with open(unmasked_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✨ 暴力解鎖完成！請查看：{os.path.basename(unmasked_path)}")

def test_ai_to_svg_v2(ai_input_path):
    inkscape_path = r"D:\Inkscape\bin\inkscape.exe"
    base_name = os.path.splitext(ai_input_path)[0]
    output_svg = f"{base_name}_v2.svg"

    # 使用 --export-plain-svg 減少 Inkscape 自己的干擾
    inkscape_cmd = [
        inkscape_path,
        ai_input_path,
        "--export-type=svg",
        "--export-plain-svg",
        "--export-area-drawing",
        "--export-filename", output_svg
    ]

    try:
        print(f"🚀 正在使用 Inkscape 進行深度轉檔...")
        subprocess.run(inkscape_cmd, check=True, capture_output=True)
        
        if os.path.exists(output_svg):
            # 轉換成功後，立刻進行暴力解鎖
            brute_force_reveal_svg(output_svg)
            
    except Exception as e:
        print(f"❌ 錯誤：{e}")

if __name__ == "__main__":
    target = r"D:\Download\【用途】-20260311T054848Z-1-001\【用途】\2025_茶王_54x180mm_名片小卡_電腦＿設計檔.ai"
    test_ai_to_svg_v2(target)