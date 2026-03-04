# -*- coding: utf-8 -*-
"""
依「學生名稱 + 五維分數」產出雷達圖 SVG，直接以 UTF-8 寫入學生姓名，避免錯字或簡體。
使用：python gen_radar_svg.py "周昱翔_開心" 5 4 4 4 5
不帶參數：掃描學生作業根目錄下所有學生資料夾（排除 Script/Data/Doc/Gallery/__pycache__），依評分表產出雷達圖。
"""
import csv
import os
import re
import sys
import math

try:
    from paths_config import BASE_DIR, iter_student_folders
except ImportError:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    def iter_student_folders():
        skip = {"Script", "Data", "Doc", "Gallery", "__pycache__"}
        for name in sorted(os.listdir(BASE_DIR)):
            path = os.path.join(BASE_DIR, name)
            if os.path.isdir(path) and name not in skip:
                yield name

CX, CY = 200, 200
R_MAX = 150  # score 5 = 150
# 五軸角度（從上方順時針，弧度）：90°, 18°, -54°, -126°, 162°
ANGLES = [math.pi / 2, math.pi / 2 - 2 * math.pi / 5, math.pi / 2 - 4 * math.pi / 5,
          math.pi / 2 - 6 * math.pi / 5, math.pi / 2 - 8 * math.pi / 5]


def scores_to_polygon(scores):
    """五個分數 (0~5) 轉成五個頂點座標。"""
    points = []
    for i, s in enumerate(scores):
        r = (s * R_MAX / 5) if s >= 0 else 0
        r = max(0, min(R_MAX, r))
        x = CX + r * math.cos(ANGLES[i])
        y = CY - r * math.sin(ANGLES[i])
        points.append(f"{x:.0f},{y:.0f}")
    return " ".join(points)


def scores_to_circles(scores):
    """五個分數對應的圓點標籤 (cx cy)。"""
    lines = []
    for i, s in enumerate(scores):
        r = (s * R_MAX / 5) if s >= 0 else 0
        r = max(0, min(R_MAX, r))
        x = CX + r * math.cos(ANGLES[i])
        y = CY - r * math.sin(ANGLES[i])
        lines.append(f'  <circle cx="{x:.0f}" cy="{y:.0f}" r="3" fill="#4169e1"/>')
    return "\n".join(lines)


def build_svg(student_name, scores):
    """student_name: 直接使用資料夾名稱（如 周昱翔_開心），以 UTF-8 寫入，不轉實體。"""
    a, b, c, d, e = scores[0], scores[1], scores[2], scores[3], scores[4]
    scores_str = f"({a}-{b}-{c}-{d}-{e})"
    polygon_pts = scores_to_polygon(scores)
    circle_tags = scores_to_circles(scores)

    # 姓名與分數直接嵌入，存檔時用 UTF-8 即可正確顯示繁體
    footer_text = f"{student_name} {scores_str}"

    svg = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="-28 0 472 400" width="400" height="400">
  <rect x="-28" y="0" width="500" height="400" fill="#ffffff"/>
  <title>5 Dimensions Radar</title>
  <g fill="none" stroke="#d0d0d0" stroke-width="0.8">
    <circle cx="{CX}" cy="{CY}" r="30"/>
    <circle cx="{CX}" cy="{CY}" r="60"/>
    <circle cx="{CX}" cy="{CY}" r="90"/>
    <circle cx="{CX}" cy="{CY}" r="120"/>
  </g>
  <circle cx="{CX}" cy="{CY}" r="150" fill="none" stroke="#000000" stroke-width="2"/>
  <g stroke="#d0d0d0" stroke-width="0.8">
    <line x1="{CX}" y1="{CY}" x2="200" y2="50"/>
    <line x1="{CX}" y1="{CY}" x2="314" y2="162"/>
    <line x1="{CX}" y1="{CY}" x2="271" y2="297"/>
    <line x1="{CX}" y1="{CY}" x2="129" y2="297"/>
    <line x1="{CX}" y1="{CY}" x2="58" y2="154"/>
  </g>
  <text x="200" y="202" text-anchor="middle" font-size="11" fill="#000">0</text>
  <text x="200" y="170" text-anchor="middle" font-size="11" fill="#000">1</text>
  <text x="200" y="140" text-anchor="middle" font-size="11" fill="#000">2</text>
  <text x="200" y="110" text-anchor="middle" font-size="11" fill="#000">3</text>
  <text x="200" y="80" text-anchor="middle" font-size="11" fill="#000">4</text>
  <text x="200" y="50" text-anchor="middle" font-size="11" fill="#000">5</text>
  <polygon points="{polygon_pts}" fill="#87ceeb" fill-opacity="0.45" stroke="#4169e1" stroke-width="2"/>
{circle_tags}
  <rect x="98" y="19" width="204" height="16" fill="#ffffff" stroke="#e0e0e0" stroke-width="0.5" rx="2"/>
  <text x="200" y="30" text-anchor="middle" font-size="11" fill="#000">意圖解析與拆解</text>
  <rect x="318" y="149" width="132" height="16" fill="#ffffff" stroke="#e0e0e0" stroke-width="0.5" rx="2"/>
  <text x="384" y="160" text-anchor="middle" font-size="11" fill="#000">審美轉譯與內涵</text>
  <rect x="218" y="309" width="132" height="16" fill="#ffffff" stroke="#e0e0e0" stroke-width="0.5" rx="2"/>
  <text x="284" y="320" text-anchor="middle" font-size="11" fill="#000">溝通效率與迭代</text>
  <rect x="50" y="309" width="132" height="16" fill="#ffffff" stroke="#e0e0e0" stroke-width="0.5" rx="2"/>
  <text x="116" y="320" text-anchor="middle" font-size="11" fill="#000">環境預見與判斷</text>
  <rect x="-22" y="149" width="132" height="16" fill="#ffffff" stroke="#e0e0e0" stroke-width="0.5" rx="2"/>
  <text x="44" y="160" text-anchor="middle" font-size="11" fill="#000">生字內化與演化</text>
  <rect x="108" y="373" width="184" height="20" fill="#ffffff" stroke="#e0e0e0" stroke-width="0.5" rx="2"/>
  <text x="200" y="387" text-anchor="middle" font-size="12" font-weight="bold" fill="#000">{footer_text}</text>
</svg>'''
    return svg


def star_string_to_score(s):
    """將「得分 (5★)」欄的 ★★★★★ / ★★★★ 等轉成 1~5。"""
    s = (s or "").strip()
    n = len(re.findall(r"★", s))
    return max(1, min(5, n)) if n else 0


def load_scores_from_csv(csv_path):
    """從評分表 CSV 讀取五維分數（依列順序：意圖、審美、溝通、環境、生字）。回傳 [int]*5 或 None。"""
    try:
        with open(csv_path, "r", encoding="utf-8") as f:
            rows = list(csv.reader(f))
    except Exception:
        return None
    if len(rows) < 2:
        return None
    headers = [h.strip() for h in rows[0]]
    col = None
    for i, h in enumerate(headers):
        if re.match(r"得分\s*\(.*★", h):
            col = i
            break
    if col is None:
        return None
    scores = []
    for row in rows[1:]:
        if len(row) <= col:
            return None
        scores.append(star_string_to_score(row[col]))
    if len(scores) != 5:
        return None
    return scores


def run_all(base_dir):
    """掃描所有學生資料夾，有評分表 CSV 就產出雷達圖。"""
    ok, skip = [], []
    for name in iter_student_folders():
        folder = os.path.join(base_dir, name)
        if not os.path.isdir(folder):
            continue
        csv_path = os.path.join(folder, f"{name}_評分表.csv")
        if not os.path.isfile(csv_path):
            skip.append(name)
            continue
        scores = load_scores_from_csv(csv_path)
        if not scores:
            skip.append(name)
            continue
        out_path = os.path.join(folder, f"{name}_雷達圖.svg")
        svg_content = build_svg(name, scores)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(svg_content)
        print("已儲存:", out_path)
        ok.append(name)
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    def safe_print(msg):
        try:
            print(msg)
        except UnicodeEncodeError:
            print(msg.encode(enc, errors="replace").decode(enc))
    print()
    print("---------- 雷達圖產出結果 ----------")
    print(f"成功（共 {len(ok)} 位）：")
    for s in ok:
        safe_print(f"  · {s}")
    if skip:
        print(f"跳過（無評分表或格式不符，共 {len(skip)} 位）：")
        for s in skip:
            safe_print(f"  · {s}")


def main():
    if len(sys.argv) >= 7:
        student_name = sys.argv[1]
        scores = [int(x) for x in sys.argv[2:7]]
        out_dir = os.path.join(BASE_DIR, student_name)
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, f"{student_name}_雷達圖.svg")
        svg_content = build_svg(student_name, scores)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(svg_content)
        print("已儲存:", out_path)
    else:
        run_all(BASE_DIR)


if __name__ == "__main__":
    main()
