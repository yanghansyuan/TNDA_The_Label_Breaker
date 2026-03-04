# -*- coding: utf-8 -*-
"""
一次處理「學生作業」下所有學生：從各人的 AI 連結檔（標準檔名或任意 docx/txt/rtf/pdf）讀取連結，
用瀏覽器開啟並擷取對話內容，寫入該生資料夾的 AI對話內容.txt。
需先安裝：pip install playwright  →  playwright install chromium；掃 PDF 需：pip install pypdf
"""
import os
import re
import sys

# 不處理的學生資料夾（依檢查結果排除）
SKIP_STUDENTS = ["楊漢軒_興奮"]

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

# 從現有腳本複用「從 docx / pdf 擷取 AI 連結」的邏輯
try:
    from extract_ai_chat_from_docx import extract_urls_from_docx, extract_urls_from_pdf, is_ai_chat_url
except ImportError:
    import zipfile
    from xml.etree import ElementTree as ET
    AI_CHAT_PATTERNS = ("chatgpt.com/share", "gemini.google.com/share", "g.co/gemini", "perplexity.ai", "claude.ai", "copilot.microsoft.com")
    def is_ai_chat_url(url):
        return any(p in url for p in AI_CHAT_PATTERNS)
    def extract_urls_from_docx(docx_path):
        urls, seen = [], set()
        try:
            with zipfile.ZipFile(docx_path, "r") as z:
                rels_path = "word/_rels/document.xml.rels"
                if rels_path in z.namelist():
                    rels = ET.fromstring(z.read(rels_path))
                    for rel in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                        t = rel.get("Target")
                        if t and (t.startswith("http://") or t.startswith("https://")) and is_ai_chat_url(t) and t not in seen:
                            urls.append(t); seen.add(t)
                doc = z.read("word/document.xml").decode("utf-8", errors="ignore")
                for u in re.findall(r"https?://[^\s<>\"']+", doc):
                    u = re.sub(r"[)\],\.{}\s]+$", "", u)
                    if is_ai_chat_url(u) and u not in seen:
                        urls.append(u); seen.add(u)
        except Exception as e:
            print("  讀取 docx 錯誤:", e)
        return urls
    def extract_urls_from_pdf(pdf_path):
        return []


def extract_urls_from_text_file(file_path):
    """從 .txt / .rtf 等純文字檔擷取 AI 連結。"""
    urls = []
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    except Exception:
        return urls
    for u in re.findall(r"https?://[^\s<>\"'\]]+", text):
        u = re.sub(r"[)\],\.{}\s]+$", "", u)
        if is_ai_chat_url(u):
            urls.append(u)
    return list(dict.fromkeys(urls))


STANDARD_NAMES_LOWER = ("ai對話串網址.docx", "ai對話串網址.txt", "ai對話串網址.pdf")
# 擷取結果檔名，不可當成連結來源
OUTPUT_FILENAME_LOWER = "ai對話內容.txt"
# 可從中擷取連結的副檔名（含非標準檔名）
LINK_FILE_EXTENSIONS = (".docx", ".txt", ".rtf", ".pdf")


def get_urls_from_student_folder(folder_path):
    """取得該生資料夾內的 AI 對話連結（含子資料夾；標準檔名優先，其次任意 docx/txt/rtf/pdf）。"""
    # 1) 優先：標準檔名 .docx / .txt（不區分大小寫，如 ai對話串網址.docx）
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER or f.lower() not in STANDARD_NAMES_LOWER:
                continue
            path = os.path.join(root, f)
            if f.lower().endswith(".docx"):
                return extract_urls_from_docx(path)
            if f.lower().endswith(".txt"):
                urls = []
                with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                    for line in fh:
                        line = line.strip()
                        if line and (line.startswith("http://") or line.startswith("https://")) and is_ai_chat_url(line):
                            urls.append(line)
                return urls
    # 2) 從任一 .pdf 擷取連結（如 對話紀錄.pdf、FIREMOOD 對話網址.pdf）
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER or not f.lower().endswith(".pdf"):
                continue
            path = os.path.join(root, f)
            urls = extract_urls_from_pdf(path)
            if urls:
                return urls
    # 3) 從任意 .docx / .txt / .rtf 擷取（需修改命名的檔案，如 對話連結.docx、企劃\\AI協作.docx）
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER:
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext not in (".docx", ".txt", ".rtf"):
                continue
            path = os.path.join(root, f)
            urls = []
            if ext == ".docx":
                urls = extract_urls_from_docx(path)
            else:
                urls = extract_urls_from_text_file(path)
            if urls:
                return urls
    return []


def has_ai_link_file_in_folder(folder_path):
    """該生資料夾內是否有可擷取連結的檔案（標準檔名、.pdf、或任意 .docx/.txt/.rtf）。"""
    for _root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER:
                continue
            if f.lower() in STANDARD_NAMES_LOWER:
                return True
            ext = os.path.splitext(f)[1].lower()
            if ext in LINK_FILE_EXTENSIONS:
                return True
    return False


def fetch_with_browser(url, wait_seconds=10):
    """用 Playwright 開網頁、等 JS 載入後擷取可見文字。Claude 分享頁延長等待並使用較像真實瀏覽器的設定。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None, "請先安裝: pip install playwright  然後執行: playwright install chromium"

    is_claude = "claude.ai" in url
    wait = 25 if is_claude else wait_seconds
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-dev-shm-usage",
                    "--no-sandbox",
                ],
            )
            context = browser.new_context(
                viewport={"width": 1280, "height": 720},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                locale="zh-TW",
            )
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(wait * 1000)
            text = page.evaluate("""() => {
                const body = document.body;
                if (!body) return '';
                return body.innerText || body.textContent || '';
            }""")
            browser.close()
            content = (text.strip() or "(頁面無擷取到文字)")
            if is_claude and (
                "security verification" in content.lower()
                or "ray id:" in content.lower()
                or "malicious bots" in content.lower()
            ):
                content += "\n\n[說明] 此為 Claude 分享連結，目前無法自動擷取對話（Cloudflare 阻擋）。請手動開啟連結、複製對話內容貼到本檔，或請學生改用 Gemini/ChatGPT 分享連結。"
            return content, None
    except Exception as e:
        return None, str(e)


def main():
    base = sys.argv[1] if len(sys.argv) > 1 else BASE_DIR

    # 找出所有「有 AI 連結檔」的學生資料夾（排除 SKIP_STUDENTS 與 Script/Data/Doc/Gallery/__pycache__）
    student_folders = []
    for name in iter_student_folders(base):
        if name in SKIP_STUDENTS:
            print(f"  略過（不處理）: {name}")
            continue
        path = os.path.join(base, name)
        if has_ai_link_file_in_folder(path):
            student_folders.append((name, path))

    if not student_folders:
        print("在底下沒有找到可處理的學生資料夾（或已全部排除）。")
        return

    print(f"找到 {len(student_folders)} 個學生資料夾（已排除：{', '.join(SKIP_STUDENTS)}），開始擷取對話…")

    total_urls = 0
    for name, folder in student_folders:
        urls = get_urls_from_student_folder(folder)
        if not urls:
            print(f"  略過 {name}（無 AI 連結）")
            continue
        total_urls += len(urls)
        out_path = os.path.join(folder, "AI對話內容.txt")
        lines = [
            "=" * 60,
            f"學生／資料夾: {name}",
            f"擷取之連結數: {len(urls)}",
            "=" * 60,
        ]
        for i, url in enumerate(urls, 1):
            lines.append("")
            lines.append("-" * 40)
            lines.append(f"[連結 {i}] {url}")
            lines.append("-" * 40)
            print(f"  {name} – 連結 {i}/{len(urls)}: {url[:50]}…")
            content, err = fetch_with_browser(url)
            if err:
                lines.append(f"[擷取失敗] {err}")
            else:
                lines.append(content if content else "（無擷取到內容）")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"  已寫入: {out_path}")

    print(f"\n完成。共處理 {len(student_folders)} 個學生、{total_urls} 個連結。")


if __name__ == "__main__":
    main()
