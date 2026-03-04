# -*- coding: utf-8 -*-
"""
讀取 Data/ 的「評分總覽_圖表資料.csv」與「評分總覽_各生.csv」，
產出長條圖、堆疊長條圖、圓餅圖等 PNG 圖檔到 Data/。
使用：python gen_overview_charts.py（需先執行 run_summarize_scores.bat 產出 CSV）
依賴：pip install matplotlib
"""
import csv
import os
import sys

try:
    from paths_config import DATA_DIR, GALLERY_DIR
except ImportError:
    _base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DATA_DIR = os.path.join(_base, "Data")
    GALLERY_DIR = os.path.join(_base, "Gallery")

def main():
    chart_csv = os.path.join(DATA_DIR, "評分總覽_圖表資料.csv")
    detail_csv = os.path.join(DATA_DIR, "評分總覽_各生.csv")

    if not os.path.isfile(chart_csv):
        print("找不到 評分總覽_圖表資料.csv，請先執行 run_summarize_scores.bat")
        sys.exit(1)

    try:
        import matplotlib
        matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei", "SimHei", "sans-serif"]
        matplotlib.rcParams["axes.unicode_minus"] = False
        import matplotlib.pyplot as plt
        import numpy as np
    except ImportError:
        print("請先安裝： pip install matplotlib")
        sys.exit(1)

    # 讀取圖表資料
    dims, avgs, n5, p5, n4, p4, n3below = [], [], [], [], [], [], []
    with open(chart_csv, "r", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        for row in r:
            dims.append(row["維度"])
            avgs.append(float(row["平均星等"]))
            n5.append(int(row["5★人數"]))
            p5.append(float(row["5★比例%"]))
            n4.append(int(row["4★人數"]))
            p4.append(float(row["4★比例%"]))
            n3below.append(int(row["3★以下人數"]))

    n_students = sum(n5) // 5 if n5 else 0  # 約等於人數
    total = n5[0] + n4[0] + n3below[0] if dims else 0
    p3below = [100 - p5[i] - p4[i] for i in range(len(dims))]

    # 短標籤（圖上較不擠）
    dim_short = ["意圖解析", "審美轉譯", "溝通迭代", "環境預見", "生字內化"]

    # ---- 圖1：各維度平均星等（橫向長條圖）----
    fig, ax = plt.subplots(figsize=(8, 4))
    y_pos = np.arange(len(dims))
    bars = ax.barh(y_pos, avgs, color="#6b9bd1", edgecolor="#2e5a87", height=0.6)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(dim_short, fontsize=11)
    ax.set_xlim(0, 5.5)
    ax.set_xlabel("平均星等", fontsize=11)
    ax.set_title("各維度平均星等", fontsize=13)
    for i, v in enumerate(avgs):
        ax.text(v + 0.08, i, f"{v:.2f}★", va="center", fontsize=10)
    plt.tight_layout()
    os.makedirs(DATA_DIR, exist_ok=True)
    out1 = os.path.join(DATA_DIR, "評分總覽_圖1_各維度平均星等.png")
    plt.savefig(out1, dpi=120, bbox_inches="tight")
    plt.close()
    print("已儲存:", out1)

    # ---- 圖2：各維度 5★ / 4★ / 3★以下 比例（堆疊橫向長條圖）----
    fig, ax = plt.subplots(figsize=(9, 4))
    y_pos = np.arange(len(dims))
    w = 0.6
    left = [0] * len(dims)
    bar5 = ax.barh(y_pos, p5, w, label="5★", color="#2e7d32", left=left)
    left = [left[i] + p5[i] for i in range(len(dims))]
    bar4 = ax.barh(y_pos, p4, w, label="4★", color="#f9a825", left=left)
    left = [left[i] + p4[i] for i in range(len(dims))]
    bar3 = ax.barh(y_pos, p3below, w, label="3★以下", color="#c62828", left=left)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(dim_short, fontsize=11)
    ax.set_xlim(0, 100)
    ax.set_xlabel("比例（%）", fontsize=11)
    ax.set_title("各維度得分分布（5★ / 4★ / 3★以下）", fontsize=13)
    ax.legend(loc="lower right", fontsize=9)
    plt.tight_layout()
    out2 = os.path.join(DATA_DIR, "評分總覽_圖2_各維度得分分布.png")
    plt.savefig(out2, dpi=120, bbox_inches="tight")
    plt.close()
    print("已儲存:", out2)

    # ---- 圖3：各維度 5★ 比例（圓餅圖）----
    fig, ax = plt.subplots(figsize=(7, 7))
    colors_pie = ["#2e7d32", "#1976d2", "#7b1fa2", "#e65100", "#00838f"]
    wedges, texts, autotexts = ax.pie(
        p5, labels=dim_short, autopct="%1.1f%%", colors=colors_pie,
        startangle=90, textprops={"fontsize": 11}
    )
    for t in autotexts:
        t.set_fontsize(10)
    ax.set_title("各維度 5★ 比例（%）", fontsize=13)
    plt.tight_layout()
    out3 = os.path.join(DATA_DIR, "評分總覽_圖3_各維度5星比例圓餅.png")
    plt.savefig(out3, dpi=120, bbox_inches="tight")
    plt.close()
    print("已儲存:", out3)

    # ---- 圖4：使用 AI 分布（圓餅圖）----
    # 優先用 Gallery/gallery_data.json（與藝廊網頁一致，會含「其他」）；若無則用評分總覽_各生.csv
    ai_count = {}
    gallery_json = os.path.join(GALLERY_DIR, "gallery_data.json")
    if os.path.isfile(gallery_json):
        try:
            import json
            with open(gallery_json, "r", encoding="utf-8") as f:
                data = json.load(f)
            for ent in data.get("entries") or []:
                ai_used = ent.get("aiUsed") or []
                label = ai_used[0] if ai_used else "未提供"
                ai_count[label] = ai_count.get(label, 0) + 1
        except Exception:
            pass
    if not ai_count and os.path.isfile(detail_csv):
        with open(detail_csv, "r", encoding="utf-8-sig") as f:
            r = csv.DictReader(f)
            for row in r:
                ai = (row.get("使用AI") or "").strip() or "—"
                ai_count[ai] = ai_count.get(ai, 0) + 1
    if ai_count:
        fig, ax = plt.subplots(figsize=(6, 6))
        labels = list(ai_count.keys())
        sizes = list(ai_count.values())
        total_n = sum(sizes)
        def autopct_count(pct):
            return f"{int(round(pct / 100 * total_n))} 人"
        colors_ai = ["#4285f4", "#0d9d57", "#ea4335", "#9e9e9e", "#ff9800"][: len(labels)]
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, autopct=autopct_count, colors=colors_ai[: len(labels)],
            startangle=90, textprops={"fontsize": 11}
        )
        ax.set_title("使用 AI 分布（人數）", fontsize=13)
        plt.tight_layout()
        out4 = os.path.join(DATA_DIR, "評分總覽_圖4_使用AI分布圓餅.png")
        plt.savefig(out4, dpi=120, bbox_inches="tight")
        plt.close()
        print("已儲存:", out4)

    # ---- 圖5：學生平均分數分布（依實際出現的分數分類）----
    # 五維平均必為 0.2 的倍數，依「實際出現的平均值」計數畫長條圖，看出各級距人數差異
    if os.path.isfile(detail_csv):
        avgs_student = []
        with open(detail_csv, "r", encoding="utf-8-sig") as f:
            r = csv.DictReader(f)
            for row in r:
                try:
                    avgs_student.append(float(row.get("平均", 0)))
                except (ValueError, TypeError):
                    pass
        if avgs_student:
            from collections import Counter
            # 四捨五入到 0.2，避免浮點誤差
            rounded = [round(a * 5) / 5 for a in avgs_student]
            count_by_score = Counter(rounded)
            # 依分數排序（3.8, 4.0, 4.2, ... 5.0）
            scores_sorted = sorted(count_by_score.keys())
            counts = [count_by_score[s] for s in scores_sorted]
            labels = [f"{s:.1f}★" for s in scores_sorted]
            fig, ax = plt.subplots(figsize=(7, 4))
            x_pos = np.arange(len(scores_sorted))
            bars = ax.bar(x_pos, counts, color="#5c6bc0", edgecolor="#283593", width=0.7)
            ax.set_xticks(x_pos)
            ax.set_xticklabels(labels, fontsize=11)
            ax.set_ylabel("人數", fontsize=11)
            ax.set_xlabel("平均星等", fontsize=11)
            ax.set_title("學生平均分數分布（依實際分數級距）", fontsize=13)
            for i, c in enumerate(counts):
                ax.text(i, c + 0.15, str(c), ha="center", fontsize=11, fontweight="bold")
            ax.set_ylim(0, max(counts) * 1.2 + 0.5 if counts else 1)
            plt.tight_layout()
            out5 = os.path.join(DATA_DIR, "評分總覽_圖5_平均分數分布.png")
            plt.savefig(out5, dpi=120, bbox_inches="tight")
            plt.close()
            print("已儲存:", out5)

    print("")
    print("圖表產出完成。請至學生作業目錄查看 PNG 檔。")


if __name__ == "__main__":
    main()
