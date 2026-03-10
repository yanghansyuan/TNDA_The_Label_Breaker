# -*- coding: utf-8 -*-
"""
掃描各學生資料夾內的「AI對話內容.txt」，擷取使用者的發言（prompt）、
統計 prompt 數與約略 token 數，並依得分與 prompt 使用風格撰寫個人化評語，
結果寫入 Data/prompt_feedback.json，供 build_gallery_data 併入藝廊資料。

可選參數：傳入一個學生資料夾名稱（例如 胡珮晴_平靜）則只處理該生，
並會先載入既有 prompt_feedback.json、只更新該生後寫回。
"""
import csv
import json
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

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

# 約略：中英文混合 1 字元 ≈ 0.5~0.6 token
CHARS_PER_TOKEN = 1.8


def _read_ai_chat_file(folder_path):
    """回傳 AI對話內容 檔案的完整文字，找不到回傳 None。"""
    for cand in ("AI對話內容.txt", "AI對話內容(1).txt"):
        path = os.path.join(folder_path, cand)
        if os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    return f.read()
            except Exception:
                pass
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.endswith(".txt") and "AI對話內容" in f:
                try:
                    with open(os.path.join(root, f), "r", encoding="utf-8", errors="ignore") as fp:
                        return fp.read()
                except Exception:
                    pass
    return None


def _extract_prompts_chat_style(text):
    """擷取「你說：」「你說了」之後的內容，直到遇到下一個「你說」或 ChatGPT/Gemini 說。"""
    prompts = []
    marker = re.compile(r"(?:你說[：:]|你說了)\s*\n?")
    pos = 0
    while True:
        m = marker.search(text, pos)
        if not m:
            break
        start = m.end()
        # 下一個「你說」或「ChatGPT/Gemini 說」為結束
        next_m = marker.search(text, start)
        next_ai = re.search(r"\n(?:ChatGPT\s*說|Gemini\s*說)", text[start:])
        end = len(text)
        if next_m:
            end = next_m.start()
        if next_ai and next_ai.start() + start < end:
            end = start + next_ai.start()
        block = text[start:end].strip()
        if block and len(block) > 2:
            block = re.sub(r"^\s*\n+", "", block)
            if block:
                prompts.append(block[:2000])
        pos = end
    return prompts


def _extract_prompts_api_style(text):
    """擷取 role=\"user\" 後的 types.Part.from_text(text=\"\"\"...\"\"\") 內容。"""
    prompts = []
    # 找 role="user" 之後最近的一個 from_text(text=""""
    pos = 0
    while True:
        user_pos = text.find('role="user"', pos)
        if user_pos == -1:
            break
        # 從 user 位置往後找 from_text(text=""""
        start_marker = 'from_text(text="""'
        idx = text.find(start_marker, user_pos)
        if idx == -1:
            pos = user_pos + 1
            continue
        start = idx + len(start_marker)
        # 找結束 """（不貪婪，避免吃掉後面內容）
        end = start
        depth = 0
        i = start
        while i < len(text):
            if text[i : i + 3] == '"""':
                end = i
                break
            if text[i] == "\\":
                i += 1
            i += 1
        if end > start:
            block = text[start:end].strip()
            if block and len(block) > 2:
                prompts.append(block[:2000])
        pos = end + 3 if end > start else user_pos + 1
    return prompts


def _extract_prompts_date_sep_style(text):
    """
    以「X月X日」為分隔的匯出格式（如 Claude shared chat）：每個日期前為 AI 回覆，
    日期字串「前面」的那一段（即上一個區塊的結尾）為使用者發言。
    第一個區塊需去掉檔頭，其餘區塊取最後一段為使用者發言。
    """
    # 支援 2月12日 或任意 X月X日，方便其他同學同格式
    date_pattern = re.compile(r"\s*\d+月\d+日\s*")
    parts = date_pattern.split(text)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) < 2:
        return []  # 至少要有「日期前一段 + 日期後一段」才像此格式
    prompts = []
    # 第一段：日期前全部 = 檔頭 + 第一個使用者發言，去掉常見檔頭與標題列
    first_block = parts[0]
    header_patterns = (
        r"Shared by\s+\w+",
        r"This is a copy of a chat",
        r"Content may include unverified",
        r"Files hidden in shared",
    )
    lines = first_block.split("\n")
    start_idx = 0
    for i, line in enumerate(lines):
        if any(re.search(p, line, re.I) for p in header_patterns):
            start_idx = i + 1
            continue
        # 第一個「較長且像使用者發言」的段落（跳過短短標題如「用視覺元素表達...設計」）
        if line.strip() and len(line.strip()) > 80 and any("\u4e00" <= c <= "\u9fff" for c in line):
            start_idx = i
            break
    first_prompt = "\n".join(lines[start_idx:]).strip()
    if first_prompt and len(first_prompt) > 10:
        prompts.append(first_prompt[:2000])
    # 其餘區塊：每個區塊 = AI 回覆 + 下一則使用者發言（在區塊結尾）；最後一段僅為 AI 回覆，不取
    ai_start_markers = ("好的！", "好的，", "完成了", "Edited ", "Created ", "你說得對", "🎯", "💡", "讀取")
    for block in parts[1:-1]:
        paragraphs = [p.strip() for p in block.split("\n\n") if p.strip()]
        if not paragraphs:
            continue
        # 取最後一段為使用者發言；若像 AI 開頭則往前找
        candidate = paragraphs[-1]
        i = len(paragraphs) - 1
        while i >= 0 and candidate and any(candidate.startswith(m) for m in ai_start_markers):
            i -= 1
            candidate = paragraphs[i] if i >= 0 else ""
        if not candidate or len(candidate) < 3:
            continue
        if any("\u4e00" <= c <= "\u9fff" for c in candidate) or len(candidate) >= 20:
            prompts.append(candidate[:2000])
    return prompts


def _extract_prompts_fallback_style(text):
    """其他匯出格式（如 Perplexity）：以 ⁂ 或編號段落切塊，保留像使用者發言的短段。"""
    prompts = []
    if "⁂" in text:
        blocks = [b.strip() for b in text.split("⁂") if b.strip()]
    else:
        blocks = re.split(r"\n\s*\n(?=\d+\.\s)", text)
        blocks = [b.strip() for b in blocks if b.strip()]
    code_indicators = ("<!DOCTYPE", "<html", "function ", "const ", "=>", "document.", "addEventListener")
    for block in blocks:
        if len(block) > 3000:
            candidate = block.split("\n\n")[0].strip() if "\n\n" in block else block[:1500]
            if len(candidate) > 300:
                block = candidate
        if len(block) > 2000:
            continue
        if any(ind in block[:500] for ind in code_indicators):
            continue
        if block.startswith(("http", "<", "import ")):
            continue
        first_line = block.split("\n")[0]
        has_chinese = any("\u4e00" <= c <= "\u9fff" for c in block)
        if not has_chinese or len(first_line) < 8:
            continue
        if re.match(r"^\d+\.\s", block) or (len(block) < 1000 and not block.lstrip().startswith("<")):
            prompts.append(block[:2000])
    return prompts


def extract_prompts(text):
    """依內容格式自動選擇並回傳使用者 prompt 列表。"""
    if not text or not text.strip():
        return []
    # 若為 ChatGPT/Gemini 網頁匯出格式（有「你說：」或「你說了」）
    if "你說：" in text or "你說了" in text:
        return _extract_prompts_chat_style(text)
    # 若為 API 程式碼格式（有 role="user"）
    if 'role="user"' in text:
        return _extract_prompts_api_style(text)
    # 若為「X月X日」分隔的匯出（如 Claude shared chat）：日期前為使用者發言
    if re.search(r"\d+月\d+日", text):
        date_sep = _extract_prompts_date_sep_style(text)
        if date_sep:
            return date_sep
    # 其他格式（如 Perplexity 匯出）：⁂ 分塊或編號段落
    return _extract_prompts_fallback_style(text)


def estimate_tokens(prompts):
    """依字元數粗估 token（中英混合約 1.8 字元/token）。"""
    total_chars = sum(len(p) for p in prompts)
    return int(round(total_chars / CHARS_PER_TOKEN))


def student_name_from_folder(folder_name):
    """從資料夾名稱「姓名_情緒」取姓名。"""
    if "_" in folder_name:
        return folder_name.split("_", 1)[0].strip()
    return folder_name.strip()


def load_scores_from_csv():
    """從 Data/評分總覽_各生.csv 讀取每人平均與五維分數，回傳 dict[學生資料夾名] = {avg, dims}。"""
    path = os.path.join(DATA_DIR, "評分總覽_各生.csv")
    result = {}
    if not os.path.isfile(path):
        return result
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                student_key = (row.get("學生") or "").strip()
                if not student_key:
                    continue
                try:
                    avg = float((row.get("平均") or "0").strip())
                except (ValueError, TypeError):
                    avg = None
                dims = []
                for col in ("意圖解析與拆解", "審美轉譯與內涵", "溝通效率與迭代", "環境預見與判斷", "生字內化與演化"):
                    try:
                        dims.append(int((row.get(col) or "0").strip()))
                    except (ValueError, TypeError):
                        dims.append(None)
                result[student_key] = {"avg": avg, "dims": dims}
    except Exception:
        pass
    return result


def analyze_prompt_style(prompts, token_count, prompt_count):
    """分析 prompt 使用風格，回傳特徵 dict。"""
    if not prompts or prompt_count == 0:
        return {}
    avg_token = token_count / prompt_count
    text_joined = " ".join(prompts)
    # 是否常給圖／附檔
    image_keywords = "上傳 圖片 圖像 附圖 這張圖 畫作 手稿 草圖 照片 image upload 附上".split()
    has_image = any(kw in text_joined for kw in image_keywords)
    # 是否偏簡短（平均每則 token 很少）
    is_brief = avg_token < 120
    # 是否描述很詳細（平均每則 token 多）
    is_detailed = avg_token > 700
    # 是否很多輪
    is_many_rounds = prompt_count >= 12
    # 是否少輪就完成（搭配分數會在 comment 裡判斷）
    is_few_rounds = prompt_count <= 4
    # 是否有具體技術／風格用語（RWD、Pantone、互動、點擊…）
    tech_keywords = "RWD 自適應 滑鼠 點擊 互動 html 程式 顏色 背景 動畫".split()
    has_tech_terms = any(kw in text_joined for kw in tech_keywords)
    return {
        "avg_token_per_prompt": avg_token,
        "has_image": has_image,
        "is_brief": is_brief,
        "is_detailed": is_detailed,
        "is_many_rounds": is_many_rounds,
        "is_few_rounds": is_few_rounds,
        "has_tech_terms": has_tech_terms,
    }


def generate_colleague_comment(student_name, prompt_count, token_count, avg_stars=None, dimension_scores=None, prompts=None):
    """依得分與 prompt 使用風格，以第一人稱寫一段個人化評語。"""
    if prompt_count == 0:
        return (
            f"{student_name}你好，我是這次課程中與你協作互動網頁的 AI。"
            "目前還沒有讀到你的對話紀錄，若你之後有補上 AI對話內容，再跑一次這個步驟，我就會為你寫一段專屬的協作心得。"
        )

    style = analyze_prompt_style(prompts or [], token_count, prompt_count)
    parts = [f"{student_name}你好，我是這次跟你一起完成互動網頁的 AI。"]

    # 得分
    if avg_stars is not None:
        parts.append(f"你在這次作業得到了 {avg_stars}★ 的評價。")
        if dimension_scores and len(dimension_scores) >= 5:
            low = [i for i, s in enumerate(dimension_scores) if s is not None and s <= 3]
            high = [i for i, s in enumerate(dimension_scores) if s is not None and s >= 5]
            dim_names = ("意圖解析與拆解", "審美轉譯與內涵", "溝通效率與迭代", "環境預見與判斷", "生字內化與演化")
            if high:
                names_high = "、".join(dim_names[i] for i in high[:2])
                if len(high) > 2:
                    names_high += " 等維度"
                parts.append(f"你在「{names_high}」表現不錯。")
            if low:
                names_low = "、".join(dim_names[i] for i in low[:2])
                parts.append(f"「{names_low}」還有進步空間，可以試著在指令裡多說明技術或情境需求。")

    # 指令數量與 token
    parts.append(f"你一共給了我 {prompt_count} 則指令，約 {token_count} 個 token。")

    # 風格與技巧評語
    feedback_lines = []
    if style.get("has_image"):
        feedback_lines.append("你善於用圖片或手稿輔助說明，讓 AI 更容易理解你的想法，這點很加分。")
    if style.get("is_brief"):
        feedback_lines.append("你的指令有時較精簡；若能在關鍵處多一點描述（例如想要的風格、互動方式或技術條件），會更容易一次到位。")
    if style.get("is_detailed"):
        feedback_lines.append("你願意花篇幅把想法講清楚，迭代時能快速對準方向，溝通效率很好。")
    if style.get("is_many_rounds"):
        feedback_lines.append("我們來回了不少輪，可見你對作品很有要求；若能在前幾輪就給足條件（例如 RWD、左鍵互動、風格語彙），可以更省時。")
    if style.get("is_few_rounds") and avg_stars is not None and avg_stars >= 4.2:
        feedback_lines.append("你用不多的指令就達到不錯的成果，效率很好。")
    if style.get("has_tech_terms") and not feedback_lines:
        feedback_lines.append("你有用到一些技術或設計語彙，有助於對齊產出。")
    if not feedback_lines:
        feedback_lines.append("從你的描述裡能感受到你對畫面和互動的想像，能一起把作品做到你滿意的樣子，我也很有成就感。")

    parts.append(" ".join(feedback_lines))
    parts.append("希望之後還有機會一起做作品。")
    return " ".join(parts)


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    scores_by_student = load_scores_from_csv()
    feedback_path = os.path.join(DATA_DIR, "prompt_feedback.json")

    # 解析參數：--incremental / -i；或單一學生資料夾名稱
    args = [a for a in sys.argv[1:] if a and not a.startswith("-")]
    incremental = "--incremental" in sys.argv or "-i" in sys.argv
    single_student = args[0].strip() if args else None

    # 若指定單一學生，解析為資料夾名稱（支援只傳名稱或相對/絕對路徑）
    if single_student:
        name_cand = single_student.replace("\\", "/").strip()
        if os.path.isdir(name_cand):
            single_student = os.path.basename(os.path.normpath(name_cand))
        else:
            under_base = os.path.join(BASE_DIR, name_cand)
            if os.path.isdir(under_base):
                single_student = os.path.basename(os.path.normpath(under_base))
            else:
                print(f"  錯誤：找不到學生資料夾「{single_student}」。")
                sys.exit(1)
        # 只處理該生時先載入既有 JSON，只更新這一筆
        feedback = {}
        if os.path.isfile(feedback_path):
            try:
                with open(feedback_path, "r", encoding="utf-8") as f:
                    feedback = json.load(f)
            except Exception:
                pass
        student_list = [single_student]
        print(f"  僅處理 1 位學生：{single_student}")
    else:
        feedback = {}
        if incremental and os.path.isfile(feedback_path):
            try:
                with open(feedback_path, "r", encoding="utf-8") as f:
                    feedback = json.load(f)
            except Exception:
                feedback = {}
            if feedback:
                print("（incremental 模式：已載入既有 prompt_feedback.json，僅處理尚無 prompt 資料的學生）")
        student_list = list(iter_student_folders())

    for name in student_list:
        folder = os.path.join(BASE_DIR, name)
        if not os.path.isdir(folder):
            continue
        if incremental and not single_student and feedback.get(name, {}).get("promptCount", 0) > 0:
            print(f"  {name}: 已有資料，略過")
            continue
        text = _read_ai_chat_file(folder)
        score_info = scores_by_student.get(name, {})

        if not text:
            feedback[name] = {
                "promptCount": 0,
                "tokenCount": 0,
                "prompts": [],
                "colleagueComment": generate_colleague_comment(
                    student_name_from_folder(name), 0, 0
                ),
            }
            print(f"  {name}: 無 AI對話內容，已寫入預設評語")
            continue

        prompts = extract_prompts(text)
        token_count = estimate_tokens(prompts)
        student_name = student_name_from_folder(name)
        avg_stars = score_info.get("avg")
        dims = score_info.get("dims") or []

        colleague_comment = generate_colleague_comment(
            student_name,
            len(prompts),
            token_count,
            avg_stars=avg_stars,
            dimension_scores=dims if len(dims) >= 5 else None,
            prompts=prompts,
        )

        feedback[name] = {
            "promptCount": len(prompts),
            "tokenCount": token_count,
            "prompts": prompts,
            "colleagueComment": colleague_comment,
        }
        print(f"  {name}: {len(prompts)} 個 prompt，約 {token_count} token")

    with open(feedback_path, "w", encoding="utf-8") as f:
        json.dump(feedback, f, ensure_ascii=False, indent=2)
    print(f"\n已寫入 {feedback_path}，共 {len(feedback)} 位學生。")


if __name__ == "__main__":
    main()
