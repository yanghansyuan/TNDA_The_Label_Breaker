# -*- coding: utf-8 -*-
"""
從「AI對話串網址.docx」讀取網址，抓取 AI 對話內容並另存為文字檔。
支援 ChatGPT share、Gemini share 等公開分享連結。
未指定路徑時，在學生作業根目錄掃描學生資料夾（排除 Script/Data/Doc/Gallery/__pycache__）。
"""
import zipfile
import re
import sys
import os
from xml.etree import ElementTree as ET

try:
    from paths_config import BASE_DIR, iter_student_folders
except ImportError:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    def iter_student_folders(base=None):
        base = base or BASE_DIR
        skip = {"Script", "Data", "Doc", "Gallery", "__pycache__"}
        for name in sorted(os.listdir(base)):
            path = os.path.join(base, name)
            if os.path.isdir(path) and name not in skip:
                yield name
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# Word 命名空間
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


# 只保留常見的 AI 對話分享連結（排除 docx 內部的 schema 等）
AI_CHAT_URL_PATTERNS = (
    "chatgpt.com/share",
    "gemini.google.com/share",
    "aistudio.google.com",  # Google AI Studio 分享／prompts 連結（常為 PDF 超連結）
    "g.co/gemini",  # 短網址，如 https://g.co/gemini/share/xxx
    "perplexity.ai",
    "claude.ai",
    "copilot.microsoft.com",
)


def is_ai_chat_url(url):
    return any(p in url for p in AI_CHAT_URL_PATTERNS)


def _get_uri_from_annot(annot):
    """從單一註解取得 URI（支援直接 /URI 或 /A 字典內的 /URI，PDF 超連結常用後者）。"""
    try:
        obj = annot.get_object() if hasattr(annot, "get_object") else annot
        # 直接放在註解上的 /URI（少見）
        if "/URI" in obj:
            u = obj["/URI"]
            if isinstance(u, bytes):
                u = u.decode("utf-8", errors="ignore")
            return u if u else None
        # 超連結通常放在 /A (Action) 字典內：/A -> /URI
        if "/A" in obj:
            action = obj["/A"]
            if hasattr(action, "get_object"):
                action = action.get_object()
            if isinstance(action, dict) and "/URI" in action:
                u = action["/URI"]
                if isinstance(u, bytes):
                    u = u.decode("utf-8", errors="ignore")
                return u if u else None
    except Exception:
        pass
    return None


def extract_urls_from_pdf(pdf_path):
    """從 .pdf 擷取文字與可點擊連結（/Annots 的 /URI 或 /A /URI）並找出 AI 對話連結。需安裝 pypdf：pip install pypdf"""
    urls = []
    try:
        from pypdf import PdfReader
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
            try:
                if "/Annots" in page:
                    for annot in page["/Annots"]:
                        u = _get_uri_from_annot(annot)
                        if u and is_ai_chat_url(u):
                            urls.append(u)
            except Exception:
                pass
        for u in re.findall(r"https?://[^\s<>\"'\]]+", text):
            u = re.sub(r"[)\],\.{}\s]+$", "", u)
            if is_ai_chat_url(u):
                urls.append(u)
        urls = list(dict.fromkeys(urls))
    except ImportError:
        pass  # 未安裝 pypdf 則略過 PDF
    except Exception:
        pass
    return urls


def extract_urls_from_txt(txt_path):
    """從 .txt 讀取內容並找出 AI 對話連結。"""
    urls = []
    try:
        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    except Exception:
        return urls
    for u in re.findall(r"https?://[^\s<>\"'\]]+", text):
        u = re.sub(r"[)\],\.{}\s]+$", "", u)
        if is_ai_chat_url(u):
            urls.append(u)
    return list(dict.fromkeys(urls))


def extract_urls_from_docx(docx_path):
    """從 .docx 中取出「AI 對話分享」類的超連結（不含 docx 內部 schema）。"""
    urls = []
    seen = set()
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            # 讀取關係檔以取得 hyperlink 的實際 URL
            rels_path = "word/_rels/document.xml.rels"
            if rels_path in z.namelist():
                rels = ET.fromstring(z.read(rels_path))
                for rel in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    target = rel.get("Target")
                    if target and (target.startswith("http://") or target.startswith("https://")):
                        if is_ai_chat_url(target) and target not in seen:
                            urls.append(target)
                            seen.add(target)

            # 讀取 document.xml 中的文字，並抓出內文裡的 AI 對話 URL
            doc = z.read("word/document.xml")
            text = doc.decode("utf-8", errors="ignore")
            found = re.findall(r"https?://[^\s<>\"']+", text)
            for u in found:
                u = re.sub(r"[)\],\.{}\s]+$", "", u)
                if is_ai_chat_url(u) and u not in seen:
                    urls.append(u)
                    seen.add(u)
    except Exception as e:
        print("讀取 docx 時發生錯誤:", e)
    return urls


def fetch_url_as_text(url, max_size=500000):
    """抓取 URL 內容並盡量抽出可讀文字（去除 HTML 標籤）。動態載入的頁面（如 Gemini/ChatGPT 分享）無法用此法取得對話。"""
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})
        with urlopen(req, timeout=15) as resp:
            raw = resp.read()
            if len(raw) > max_size:
                raw = raw[:max_size]
            html = raw.decode("utf-8", errors="replace")
    except (URLError, HTTPError, OSError) as e:
        return f"[無法取得此連結: {url}]\n錯誤: {e}\n"
    html = re.sub(r"<script[^>]*>[\s\S]*?</script>", " ", html, flags=re.I)
    html = re.sub(r"<style[^>]*>[\s\S]*?</style>", " ", html, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    return text[:80000] if len(text) > 80000 else text


def fetch_with_browser(url, wait_seconds=10):
    """用 Playwright 開網頁、等 JS 載入後擷取可見文字（才能拿到 Gemini/ChatGPT 分享頁的對話）。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None, "請先安裝: pip install playwright  然後執行: playwright install chromium"
    # Claude 分享頁常被 Cloudflare 阻擋，延長等待並使用較像真實瀏覽器的設定
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
            # 若仍為 Cloudflare 驗證頁，附加說明
            if is_claude and (
                "security verification" in content.lower()
                or "ray id:" in content.lower()
                or "malicious bots" in content.lower()
            ):
                content += "\n\n[說明] 此為 Claude 分享連結，目前無法自動擷取對話（Cloudflare 阻擋）。請手動開啟連結、複製對話內容貼到本檔，或請學生改用 Gemini/ChatGPT 分享連結。"
            return content, None
    except Exception as e:
        return None, str(e)


# 擷取結果檔名，不可當成「連結來源」否則會重複讀到 Cloudflare 驗證頁
OUTPUT_FILENAME_LOWER = "ai對話內容.txt"

# 模糊比對：檔名「包含」ai、對話、網址且不含「內容」即視為標準連結檔（與 check_missing_ai_links 一致）
def is_standard_ai_link_filename(filename):
    """檔名含 ai、對話、網址，且不含「內容」→ 視為標準 AI 連結檔。"""
    lower = filename.lower()
    if "內容" in filename:
        return False
    return "ai" in lower and "對話" in filename and "網址" in filename


def get_urls_from_student_folder(folder_path):
    """從該生資料夾（含子資料夾）取得 AI 連結檔與 URL 列表。回傳 (urls, source_file_path)，找不到則 ([], None)。"""
    folder_path = os.path.abspath(folder_path)
    if not os.path.isdir(folder_path):
        return [], None
    # 1) 優先：檔名模糊符合「AI＋對話＋網址」的 .docx / .txt / .pdf（含子資料夾）
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER:
                continue  # 跳過擷取結果檔
            if not is_standard_ai_link_filename(f):
                continue
            path = os.path.join(root, f)
            if f.lower().endswith(".docx"):
                urls = extract_urls_from_docx(path)
                return (urls, path)
            if f.lower().endswith(".txt"):
                urls = extract_urls_from_txt(path)
                return (urls, path)
            if f.lower().endswith(".pdf"):
                urls = extract_urls_from_pdf(path)
                if urls:
                    return (urls, path)
    # 2) 任一 .pdf（非標準檔名者）
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER or not f.lower().endswith(".pdf"):
                continue
            path = os.path.join(root, f)
            urls = extract_urls_from_pdf(path)
            if urls:
                return (urls, path)
    # 3) 任意 .docx / .txt / .rtf（含子資料夾，如 企劃\\AI Chat history web link.rtf）
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower() == OUTPUT_FILENAME_LOWER:
                continue  # 跳過擷取結果檔
            ext = os.path.splitext(f)[1].lower()
            if ext not in (".docx", ".txt", ".rtf"):
                continue
            path = os.path.join(root, f)
            if ext == ".docx":
                urls = extract_urls_from_docx(path)
            else:
                urls = extract_urls_from_txt(path)  # .txt / .rtf 皆用文字擷取
            if urls:
                return (urls, path)
    return [], None


# 支援的副檔名（標準連結檔）
LINK_FILE_EXTENSIONS = (".docx", ".txt", ".pdf")


def find_docx_in_folder(base_dir):
    """在學生作業根目錄下搜尋「檔名含 AI、對話、網址」的連結檔（.docx / .txt / .pdf），優先王欣怡_悲傷。"""
    def first_matching_file(dir_path):
        if not os.path.isdir(dir_path):
            return None
        for f in sorted(os.listdir(dir_path)):
            ext = os.path.splitext(f)[1].lower()
            if ext not in LINK_FILE_EXTENSIONS:
                continue
            if not is_standard_ai_link_filename(f):
                continue
            return os.path.join(dir_path, f)
        return None
    # 王欣怡_悲傷 優先
    for name in iter_student_folders(base_dir):
        if "王欣怡" in name or "悲傷" in name:
            path = first_matching_file(os.path.join(base_dir, name))
            if path:
                return path
    for name in iter_student_folders(base_dir):
        path = first_matching_file(os.path.join(base_dir, name))
        if path:
            return path
    return None


def main():
    arg_path = sys.argv[1] if len(sys.argv) > 1 else None
    docx_path = None
    urls_from_folder = None  # 由 get_urls_from_student_folder 取得時使用，輸出目錄為學生資料夾
    if arg_path:
        arg_path = os.path.normpath(os.path.abspath(arg_path))
        # 若從 Script 目錄執行，傳入「陳子怡_平靜」會變成 Script\陳子怡_平靜（不存在），改在學生作業根目錄下找
        if not os.path.isfile(arg_path) and not os.path.isdir(arg_path):
            try_below_base = os.path.join(BASE_DIR, sys.argv[1].strip())
            if os.path.isdir(try_below_base):
                arg_path = os.path.normpath(os.path.abspath(try_below_base))
        if os.path.isfile(arg_path):
            docx_path = arg_path
        elif os.path.isdir(arg_path):
            urls_found, source_path = get_urls_from_student_folder(arg_path)
            if urls_found and source_path:
                docx_path = source_path
                urls_from_folder = urls_found
            else:
                docx_path = None
        else:
            docx_path = None
        if not docx_path or not os.path.isfile(docx_path):
            print("指定的路徑不存在或該資料夾內沒有可用的連結檔（AI對話串網址.docx / .txt / .pdf，或子資料夾內之 .docx / .txt / .rtf / .pdf）")
            print("請傳「學生資料夾」路徑，例如：run_extract_ai_chat.bat \"顏士傑_憤怒\" 或完整路徑")
            sys.exit(1)
    else:
        docx_path = find_docx_in_folder(BASE_DIR)
        if not docx_path or not os.path.isfile(docx_path):
            print("找不到標準連結檔（AI對話串網址.docx / .txt / AI對話網址.pdf 等），請指定路徑: python extract_ai_chat_from_docx.py <資料夾或檔案路徑>")
            sys.exit(1)
    # 傳入學生資料夾時，結果寫在該資料夾根目錄；傳入檔案時寫在該檔案所在目錄
    output_dir = arg_path if (arg_path and os.path.isdir(arg_path) and urls_from_folder is not None) else os.path.dirname(docx_path)
    output_txt = os.path.join(output_dir, "AI對話內容.txt")

    if urls_from_folder is not None:
        urls = urls_from_folder
    elif docx_path.lower().endswith(".txt") or docx_path.lower().endswith(".rtf"):
        urls = extract_urls_from_txt(docx_path)
    elif docx_path.lower().endswith(".pdf"):
        urls = extract_urls_from_pdf(docx_path)
    else:
        urls = extract_urls_from_docx(docx_path)
    if not urls:
        print("未在文件中找到任何網址。")
        with open(output_txt, "w", encoding="utf-8") as f:
            f.write("（未找到任何 AI 對話連結）\n")
        print("已寫入:", output_txt)
        return

    lines = []
    lines.append("=" * 60)
    source_label = os.path.relpath(docx_path, output_dir) if output_dir != os.path.dirname(docx_path) else os.path.basename(docx_path)
    lines.append("來源檔案: " + source_label)
    lines.append("擷取之連結數: " + str(len(urls)))
    lines.append("=" * 60)

    for i, url in enumerate(urls, 1):
        lines.append("")
        lines.append("-" * 40)
        lines.append(f"[連結 {i}] {url}")
        lines.append("-" * 40)
        content, err = fetch_with_browser(url)
        if err:
            content = fetch_url_as_text(url)
            if not content.strip() or len(content.strip()) < 200:
                lines.append(f"[建議用瀏覽器擷取] {err}")
                lines.append("（此連結為動態載入，請改執行 run_extract_ai_chats_all.bat 或安裝 Playwright）")
            else:
                lines.append(content)
        else:
            lines.append(content if content else "（無擷取到內容）")

    with open(output_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("已擷取", len(urls), "個連結的內容，另存為:")
    print(output_txt)


if __name__ == "__main__":
    main()
