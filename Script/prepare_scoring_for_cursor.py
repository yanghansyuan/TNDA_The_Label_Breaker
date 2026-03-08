# -*- coding: utf-8 -*-
"""
掃描「有 AI對話內容 但尚無評分表」的學生，產出 Doc/請AI評分_待辦.md，
內容為已填好待評分名單的完整提示詞，供使用者在 Cursor 中開啟後一次送給 Composer 請 AI 評分。

執行：python prepare_scoring_for_cursor.py
建議：在 run_all_scripts.bat 步驟 3 中與 analyze_prompt_and_feedback.py 一併執行。
"""
import os
import re

try:
    from paths_config import BASE_DIR, DOC_DIR, iter_student_folders
except ImportError:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DOC_DIR = os.path.join(BASE_DIR, "Doc")
    def iter_student_folders():
        skip = {"Script", "Data", "Doc", "Gallery", "__pycache__"}
        for name in sorted(os.listdir(BASE_DIR)):
            path = os.path.join(BASE_DIR, name)
            if os.path.isdir(path) and name not in skip:
                yield name


def has_ai_chat_file(folder_path):
    """資料夾內是否有 AI對話內容 相關文字檔。"""
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.endswith(".txt") and "AI對話" in f:
                return True
    return False


def has_score_csv(folder_path, folder_name):
    """是否已有 {資料夾名}_評分表.csv。"""
    csv_name = f"{folder_name}_評分表.csv"
    path = os.path.join(folder_path, csv_name)
    return os.path.isfile(path)


def get_prompt_template():
    """從 Doc/評分用AI提示詞.txt 讀取【請從這裡開始複製】～【複製到這裡為止】之間的內容。"""
    path = os.path.join(DOC_DIR, "評分用AI提示詞.txt")
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception:
        return None
    start_m = re.search(r"【請從這裡開始複製】\s*\n+", text)
    end_m = re.search(r"\n\s*【複製到這裡為止】", text)
    if not start_m or not end_m:
        return None
    return text[start_m.end() : end_m.start()].strip()


def main():
    os.makedirs(DOC_DIR, exist_ok=True)
    out_path = os.path.join(DOC_DIR, "請AI評分_待辦.md")

    need_scoring = []
    for name in iter_student_folders():
        folder = os.path.join(BASE_DIR, name)
        if not os.path.isdir(folder):
            continue
        if not has_ai_chat_file(folder):
            continue
        if has_score_csv(folder, name):
            continue
        need_scoring.append(name)

    prompt_body = get_prompt_template()
    if not prompt_body:
        # 若讀不到提示詞檔，仍可產出待辦，但提示詞區塊用簡短說明
        prompt_body = (
            "請依 Doc/評分用AI提示詞.txt 的評分標準與 CSV 格式，"
            "為下方「■ 要評分的學生」所列每位學生產出 {資料夾名}_評分表.csv，存到該生資料夾。"
        )

    # 替換「■ 要評分的學生」底下的佔位為實際名單
    student_list = "、".join(need_scoring) if need_scoring else "（無）"
    if "■ 要評分的學生" in prompt_body and "（請在此列出" in prompt_body:
        prompt_body = re.sub(
            r"（請在此列出[^）]+。）",
            f"以下 {len(need_scoring)} 位：{student_list}。請在 Cursor 對話中 @ 各生的 AI對話內容.txt 後，依對話內容評分。",
            prompt_body,
            count=1,
        )
    else:
        prompt_body = f"■ 要評分的學生\n{student_list}\n\n" + prompt_body

    # 列出可 @ 的檔案路徑（相對於專案根，方便 Cursor @）
    attach_hint = ""
    if need_scoring:
        attach_hint = "請在 Composer 對話中 @ 以下檔案（每位學生一個 AI對話內容.txt），再送出下方評分指示：\n\n"
        for name in need_scoring:
            attach_hint += f"- `{name}/AI對話內容.txt`\n"
        attach_hint += "\n---\n\n"

    content = f"""# 請 AI 評分待辦（由 prepare_scoring_for_cursor.py 自動產生）

## 使用方式（Cursor）

1. 在 Cursor 中開啟本檔案。
2. 開啟 Composer（Ctrl+I 或 Cmd+I）。
3. 在輸入框用 **@** 提及下方「請 @ 的檔案」中列出的每個 `AI對話內容.txt`（可一次 @ 多個）。
4. 將本頁 **「以下為評分指示」** 區塊內的整段文字複製貼到 Composer，送給 AI。
5. AI 會依標準產出每位學生的 `{{資料夾名}}_評分表.csv`，請確認已存到該生資料夾。
6. 完成後執行 `run_all_scripts.bat 4 8` 併入總表與雷達圖。

---

## 待評分學生數：{len(need_scoring)}

{attach_hint}

## 以下為評分指示（整段複製貼到 Composer）

{prompt_body}
"""

    if not need_scoring:
        content = """# 請 AI 評分待辦

目前**沒有**需要評分的學生（每位有 AI對話內容 的學生都已有 評分表.csv）。

若之後有新同學加入並有 AI對話內容.txt，請重新執行步驟 3（或執行 `python Script/prepare_scoring_for_cursor.py`），本檔案會更新為待評分名單與完整提示詞。
"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)

    if need_scoring:
        print(f"  已產出 {out_path}，共 {len(need_scoring)} 位待評分：{student_list}")
    else:
        print("  無待評分學生，已更新 請AI評分_待辦.md 為「無待辦」說明。")


if __name__ == "__main__":
    main()
