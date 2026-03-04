# -*- coding: utf-8 -*-
"""
掃描「學生作業」下所有「{資料夾名}_評分表.csv」，
彙整統計並產出：總覽報告（優缺點、高低分差異、建議）、圖表用 CSV。
使用：python summarize_scores.py
輸出：Data/ 評分總覽報告.txt、評分總覽_圖表資料.csv、評分總覽_各生.csv
"""
import csv
import os
import re

try:
    from paths_config import BASE_DIR, DATA_DIR, iter_student_folders
except ImportError:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DATA_DIR = os.path.join(BASE_DIR, "Data")
    def iter_student_folders():
        skip = {"Script", "Data", "Doc", "Gallery", "__pycache__"}
        for name in sorted(os.listdir(BASE_DIR)):
            path = os.path.join(BASE_DIR, name)
            if os.path.isdir(path) and name not in skip:
                yield name

DIM_NAMES = [
    "意圖解析與拆解",
    "審美轉譯與內涵",
    "溝通效率與迭代",
    "環境預見與判斷",
    "生字內化與演化",
]


def star_string_to_score(s):
    s = (s or "").strip()
    n = len(re.findall(r"★", s))
    return max(1, min(5, n)) if n else 0


def load_one_scores(csv_path):
    """回傳 (scores[5], ai_name) 或 (None, None)。"""
    try:
        with open(csv_path, "r", encoding="utf-8") as f:
            rows = list(csv.reader(f))
    except Exception:
        return None, None
    if len(rows) < 6:
        return None, None
    headers = [str(h).strip() for h in rows[0]]
    col = None
    for i, h in enumerate(headers):
        if re.match(r"得分\s*\(.*★", h):
            col = i
            break
    if col is None:
        return None, None
    scores = []
    ai_name = ""
    for row in rows[1:6]:
        if len(row) <= col:
            return None, None
        scores.append(star_string_to_score(row[col]))
        if len(row) > 1 and not ai_name:
            ai_name = (row[1] or "").strip()
    if len(scores) != 5:
        return None, None
    return scores, ai_name


def main():
    # 收集所有有評分表的學生（排除 Script/Data/Doc/Gallery/__pycache__）
    students = []
    for name in iter_student_folders():
        folder = os.path.join(BASE_DIR, name)
        csv_name = f"{name}_評分表.csv"
        csv_path = os.path.join(folder, csv_name)
        if not os.path.isfile(csv_path):
            continue
        scores, ai_name = load_one_scores(csv_path)
        if scores is None:
            continue
        avg = round(sum(scores) / 5, 2)
        students.append({
            "name": name,
            "scores": scores,
            "avg": avg,
            "ai": ai_name or "—",
        })

    if not students:
        print("未找到任何評分表。")
        return

    n = len(students)
    # 各維度統計
    dim_avg = [sum(s["scores"][i] for s in students) / n for i in range(5)]
    dim_count_5 = [sum(1 for s in students if s["scores"][i] == 5) for i in range(5)]
    dim_count_4 = [sum(1 for s in students if s["scores"][i] == 4) for i in range(5)]
    dim_count_3 = [sum(1 for s in students if s["scores"][i] == 3) for i in range(5)]
    dim_count_2 = [sum(1 for s in students if s["scores"][i] == 2) for i in range(5)]
    dim_count_1 = [sum(1 for s in students if s["scores"][i] == 1) for i in range(5)]
    pct_5 = [round(100 * dim_count_5[i] / n, 1) for i in range(5)]
    pct_4 = [round(100 * dim_count_4[i] / n, 1) for i in range(5)]

    # 高低分組（平均 >= 4.5 為高，< 4 為低）
    high = [s for s in students if s["avg"] >= 4.5]
    low = [s for s in students if s["avg"] < 4.0]
    if high and low:
        high_avg_dim = [sum(s["scores"][i] for s in high) / len(high) for i in range(5)]
        low_avg_dim = [sum(s["scores"][i] for s in low) / len(low) for i in range(5)]
        diff_dim = [round(high_avg_dim[i] - low_avg_dim[i], 2) for i in range(5)]
    else:
        diff_dim = [0] * 5

    # 整體平均
    overall_avg = sum(s["avg"] for s in students) / n
    avg_avg = round(overall_avg, 2)

    # ---- 1. 總覽報告 ----
    report_lines = [
        "=" * 60,
        "評分總覽報告（互動式情緒網頁／AI 協作作業）",
        "=" * 60,
        "",
        f"有效評分人數：{n} 人",
        f"整體平均星等：{avg_avg}★",
        "",
        "----------------------------------------",
        "一、各維度平均與 5★／4★ 比例（可圖表化）",
        "----------------------------------------",
        "",
    ]
    for i in range(5):
        report_lines.append(
            f"  {DIM_NAMES[i]}：平均 {dim_avg[i]:.2f}★ ｜ 5★ {pct_5[i]}% ｜ 4★ {pct_4[i]}%"
        )
    report_lines.extend([
        "",
        "----------------------------------------",
        "二、學生普遍優點",
        "----------------------------------------",
        "",
    ])
    # 優點：哪個維度平均最高、5★ 比例最高
    best_dim_idx = max(range(5), key=lambda i: (dim_avg[i], pct_5[i]))
    report_lines.append(f"  · 表現最好的維度：「{DIM_NAMES[best_dim_idx]}」（平均 {dim_avg[best_dim_idx]:.2f}★，5★ 比例 {pct_5[best_dim_idx]}%）")
    if pct_5[best_dim_idx] >= 50:
        report_lines.append(f"  · 多數學生能清楚拆解意圖、具體描述情緒與技術需求，並在迭代中維持一致目標。")
    report_lines.extend([
        "  · 多數作品具備左鍵互動與自適應（RWD）需求，且能與 AI 多輪修正至滿意。",
        "  · 審美語彙（顏色、質感、動態）與成品轉譯多能對應，情緒內涵可見。",
        "",
        "----------------------------------------",
        "三、學生普遍可改進處（缺點）",
        "----------------------------------------",
        "",
    ])
    worst_dim_idx = min(range(5), key=lambda i: (dim_avg[i], -pct_5[i]))
    report_lines.append(f"  · 相對較弱的維度：「{DIM_NAMES[worst_dim_idx]}」（平均 {dim_avg[worst_dim_idx]:.2f}★，5★ 比例 {pct_5[worst_dim_idx]}%）")
    report_lines.extend([
        "  · 部分同學首輪未一次給足風格／動態語彙或技術條件，需多輪才補齊。",
        "  · 環境預見（裝置、邊界、效能）常為「做到再補」而非事先提出。",
        "  · 生字內化：可多運用課程與 AI 回饋的術語，讓指令更精準。",
        "",
        "----------------------------------------",
        "四、分數高低主要差在哪裡？",
        "----------------------------------------",
        "",
    ])
    if high and low:
        report_lines.append(f"  高分组（平均 ≥ 4.5★，{len(high)} 人）vs 低分组（平均 < 4.0★，{len(low)} 人）維度差異：")
        report_lines.append("")
        for i in range(5):
            report_lines.append(f"  · {DIM_NAMES[i]}：高分组平均較低分组多 {diff_dim[i]:.2f}★")
        report_lines.append("")
        max_diff_idx = max(range(5), key=lambda i: diff_dim[i])
        report_lines.append(f"  差異最大的維度：「{DIM_NAMES[max_diff_idx]}」—— 高分组在此維度明顯較佳，整體而言高分组在意圖拆解、溝通迭代與語彙運用上通常更具體、一輪內給足或快速補齊。")
    else:
        report_lines.append("  （高／低分組人數不足，暫不比較；可依「評分總覽_各生.csv」自行篩選比較。）")
    report_lines.extend([
        "",
        "----------------------------------------",
        "五、可圖表化的百分比資料摘要",
        "----------------------------------------",
        "",
        "  · 各維度 5★ 比例（%）：" + "、".join(f"{DIM_NAMES[i]} {pct_5[i]}%" for i in range(5)),
        "  · 各維度 4★ 比例（%）：" + "、".join(f"{DIM_NAMES[i]} {pct_4[i]}%" for i in range(5)),
        "  · 各維度平均星等：已寫入「評分總覽_圖表資料.csv」，可用 Excel 繪製長條圖／雷達圖。",
        "  · 每位學生五維分數與平均：已寫入「評分總覽_各生.csv」，可排序、篩選、繪製分布。",
        "  · 若需直接產出 PNG 圖表（長條圖、圓餅圖、直方圖），可執行 run_gen_overview_charts.bat。",
        "",
        "----------------------------------------",
        "六、對 AI 而言：何謂好 prompt／高品質使用方式？",
        "----------------------------------------",
        "",
        "  根據本批評分維度與高分组表現，整理出以下幾點，供日後與 AI 協作時參考：",
        "",
        "  1. 一次給足「目的＋約束」",
        "     明確說出你要的產物（例如：可互動的 HTML 網頁、程式碼）、情緒主題，以及技術條件（滑鼠左鍵、自適應、不要照片等），減少來回補問。",
        "",
        "  2. 用具體語彙描述「長相」與「行為」",
        "     不只說「快樂」「悲傷」，而是補充顏色、形狀、質感、明暗、動態（例如：緩慢墜落、點擊時震動、放射狀、浮誇）。AI 才能對應到可實作的視覺與互動。",
        "",
        "  3. 迭代時一句話鎖定一個改動",
        "     例如：「花朵再大一點」「不要衝出畫面」「加上不規則放射狀」—— 一次一個焦點，方便 AI 對準修改，也方便你驗收。",
        "",
        "  4. 事先提使用情境與限制",
        "     若需要多裝置、觸控、邊界不超出、效能或數量上限，在第一輪或第二輪就說出來，比「做到再補」更容易得到完整實作。",
        "",
        "  5. 吸收並重用 AI 與課程的用語",
        "     當 AI 回覆出現「RWD」「動態語彙」「飽和度」等詞，下一輪可試著用同一套語彙微調需求，指令會更精準、產出更貼近預期。",
        "",
        "  6. 出錯時簡潔指出並要求修正",
        "     例如：「重來，你遇到 script error 了」「這裡不對，請改成……」—— 不糾結長篇解釋，直接給行為指令，有助快速收斂。",
        "",
        "=" * 60,
    ])

    os.makedirs(DATA_DIR, exist_ok=True)
    report_path = os.path.join(DATA_DIR, "評分總覽報告.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print("已寫入:", report_path)

    # ---- 2. 圖表資料 CSV（各維度統計）----
    chart_path = os.path.join(DATA_DIR, "評分總覽_圖表資料.csv")
    with open(chart_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["維度", "平均星等", "5★人數", "5★比例%", "4★人數", "4★比例%", "3★以下人數"])
        for i in range(5):
            w.writerow([
                DIM_NAMES[i],
                round(dim_avg[i], 2),
                dim_count_5[i],
                pct_5[i],
                dim_count_4[i],
                pct_4[i],
                dim_count_3[i] + dim_count_2[i] + dim_count_1[i],
            ])
    print("已寫入:", chart_path)

    # ---- 3. 各生明細 CSV ----
    detail_path = os.path.join(DATA_DIR, "評分總覽_各生.csv")
    with open(detail_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["學生", "使用AI"] + DIM_NAMES + ["平均"])
        for s in students:
            w.writerow([s["name"], s["ai"]] + s["scores"] + [s["avg"]])
    print("已寫入:", detail_path)

    print("")
    print("總覽結論已產出，請開啟「評分總覽報告.txt」查看；圖表可依「評分總覽_圖表資料.csv」「評分總覽_各生.csv」在 Excel 繪製。")


if __name__ == "__main__":
    main()
