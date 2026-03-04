# -*- coding: utf-8 -*-
"""
詳細檢查「學生作業」底下各學生資料夾：
1. 命名正確：檔名含「AI」「對話」「網址」且內容含 AI 連結（模糊比對，如 AI對話串網址.docx、AI對話網址.pdf）
2. 需修改命名：有 AI 連結但檔名不符合上述（會列出錯誤檔名，方便要求學生改）
3. 完全沒有：任何非 HTML 文字檔裡都找不到 AI 聊天串連結
會掃描資料夾內「含子資料夾」的所有非 HTML 文字類檔案並搜尋連結。
若標準檔在子資料夾（如 企劃\\AI對話串網址.docx）會特別標示，方便要求改放根目錄。
掃描時排除：Script、Data、Doc、Gallery、__pycache__。
"""
import os
import re
import sys
import zipfile
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

# AI 對話分享連結的關鍵字（與 extract_ai_chat_from_docx 一致）
AI_CHAT_PATTERNS = (
    "chatgpt.com/share",
    "gemini.google.com/share",
    "g.co/gemini",
    "perplexity.ai",
    "claude.ai",
    "copilot.microsoft.com",
)
STANDARD_DOCX = "AI對話串網址.docx"
STANDARD_TXT = "AI對話串網址.txt"
STANDARD_PDF = "AI對話網址.pdf"

# 模糊比對：檔名「包含」這些字串即視為標準連結檔（排除含「內容」的擷取結果檔）
def is_standard_ai_link_filename(filename):
    """檔名含 ai、對話、網址，且不含「內容」→ 視為標準 AI 連結檔（如 AI對話串網址.docx、AI對話網址.pdf）。"""
    lower = filename.lower()
    if "內容" in filename:
        return False
    return "ai" in lower and "對話" in filename and "網址" in filename

# 要掃描的副檔名（非 HTML 的文字類，含 PDF）
TEXT_EXTENSIONS = {".docx", ".txt", ".rtf", ".doc", ".md", ".odt", ".pdf"}
SKIP_EXTENSIONS = {".html", ".htm", ".xhtml", ".css", ".js", ".json", ".lock"}  # 略過


def is_ai_chat_url(url):
    return any(p in url for p in AI_CHAT_PATTERNS)


def extract_urls_from_docx(docx_path):
    """從 .docx 取出 AI 對話連結。"""
    urls = []
    seen = set()
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            rels_path = "word/_rels/document.xml.rels"
            if rels_path in z.namelist():
                rels = ET.fromstring(z.read(rels_path))
                for rel in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    target = rel.get("Target")
                    if target and (target.startswith("http://") or target.startswith("https://")) and is_ai_chat_url(target) and target not in seen:
                        urls.append(target)
                        seen.add(target)
            doc = z.read("word/document.xml").decode("utf-8", errors="ignore")
            for u in re.findall(r"https?://[^\s<>\"']+", doc):
                u = re.sub(r"[)\],\.{}\s]+$", "", u)
                if is_ai_chat_url(u) and u not in seen:
                    urls.append(u)
                    seen.add(u)
    except Exception:
        pass
    return urls


def extract_urls_from_text_file(file_path):
    """從純文字檔（.txt, .rtf, .md 等）讀取內容並找出 AI 連結。"""
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
    return list(dict.fromkeys(urls))  # 去重保留順序


def extract_urls_from_doc_legacy(file_path):
    """舊版 .doc 為二進位，僅嘗試在原始位元組中搜尋 URL 字樣。"""
    urls = []
    try:
        with open(file_path, "rb") as f:
            raw = f.read(2 * 1024 * 1024)  # 最多 2MB
        text = raw.decode("utf-8", errors="ignore") + raw.decode("latin-1", errors="ignore")
        for u in re.findall(r"https?://[^\s<>\"'\]]+", text):
            u = re.sub(r"[)\],\.{}\s]+$", "", u)
            if is_ai_chat_url(u):
                urls.append(u)
    except Exception:
        pass
    return list(dict.fromkeys(urls))


def _pypdf_available():
    try:
        from pypdf import PdfReader
        return True
    except ImportError:
        return False


def extract_urls_from_pdf(file_path):
    """從 .pdf 擷取文字與「可點擊連結」並找出 AI 連結。需安裝 pypdf：pip install pypdf"""
    urls = []
    try:
        from pypdf import PdfReader
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
            # 擷取頁面中的超連結註解（可點擊的連結常存在這裡）
            try:
                if "/Annots" in page:
                    for annot in page["/Annots"]:
                        obj = annot.get_object()
                        if "/URI" in obj:
                            u = obj["/URI"]
                            if isinstance(u, bytes):
                                u = u.decode("utf-8", errors="ignore")
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
        pass
    except Exception:
        pass
    return urls


def get_extension(path):
    return os.path.splitext(path)[1].lower()


def scan_folder_for_ai_links(folder_path):
    """
    遞迴掃描該生資料夾（含子資料夾如 企劃）內所有可讀文字檔，回傳：
    - standard_in_root: bool  標準檔在「根目錄」且內容有 AI 連結
    - standard_in_subfolder: list of 相對路徑  標準檔在子資料夾（如 企劃\\AI對話串網址.docx）
    - wrong_name_files: list of (相對路徑, 連結數)  有連結但檔名不正確
    - all_found_urls: set
    """
    standard_in_root = False
    standard_in_subfolder = []
    wrong_name_files = []
    all_found_urls = set()

    for root, _dirs, files in os.walk(folder_path):
        for filename in files:
            path = os.path.join(root, filename)
            ext = get_extension(filename)
            if ext in SKIP_EXTENSIONS or ext not in TEXT_EXTENSIONS:
                continue

            rel = os.path.relpath(path, folder_path)  # 如 "企劃\\AI對話串網址.docx" 或 "ai對話串網址.docx"

            urls = []
            if ext == ".docx":
                urls = extract_urls_from_docx(path)
            elif ext == ".doc":
                urls = extract_urls_from_doc_legacy(path)
            elif ext == ".pdf":
                urls = extract_urls_from_pdf(path)
            else:
                urls = extract_urls_from_text_file(path)

            # 模糊比對：檔名含 ai、對話、網址且不含「內容」即視為標準連結檔
            is_standard = is_standard_ai_link_filename(filename)
            if urls:
                for u in urls:
                    all_found_urls.add(u)
                if not is_standard:
                    wrong_name_files.append((rel, len(urls)))
            if is_standard:
                if root == folder_path:
                    standard_in_root = True
                else:
                    standard_in_subfolder.append(rel)

    return standard_in_root, standard_in_subfolder, wrong_name_files, all_found_urls


def _students_with_pdf(base):
    """回傳「資料夾內有 .pdf 檔」的學生名稱列表（僅掃描學生資料夾，排除 Script/Data/Doc/Gallery/__pycache__）。"""
    out = []
    for name in iter_student_folders(base):
        path = os.path.join(base, name)
        for _root, _dirs, files in os.walk(path):
            if any(f.lower().endswith(".pdf") for f in files):
                out.append(name)
                break
    return out


def main():
    base = sys.argv[1] if len(sys.argv) > 1 else BASE_DIR

    # 若有人交 PDF 但未安裝 pypdf，先提示
    has_pypdf = _pypdf_available()
    students_with_pdf = _students_with_pdf(base)
    if not has_pypdf and students_with_pdf:
        print("!" * 70)
        print("【重要】偵測到以下學生資料夾內有 .pdf 檔，但尚未安裝 pypdf，無法讀取 PDF 內的連結：")
        for n in students_with_pdf:
            print("  -", n)
        print("請在命令列執行：  pip install pypdf")
        print("安裝完成後重新執行本檢查，即可掃描 PDF（如 FIREMOOD 對話網址.pdf）。")
        print("!" * 70)
        print()

    correct = []              # 命名正確且在「根目錄」
    in_subfolder = []         # 命名正確但在子資料夾 → (學生名, [相對路徑, ...])
    wrong_name = []           # 有 AI 連結但檔名不正確
    missing = []              # 完全沒有找到 AI 連結

    for name in iter_student_folders(base):
        path = os.path.join(base, name)

        standard_root, standard_sub, wrong_files, all_urls = scan_folder_for_ai_links(path)

        if standard_root:
            correct.append(name)
        elif standard_sub:
            in_subfolder.append((name, standard_sub))
            if wrong_files:
                wrong_name.append((name, wrong_files))
        elif wrong_files:
            wrong_name.append((name, wrong_files))
        else:
            missing.append(name)

    # 輸出
    print("=" * 70)
    print("檢查資料夾:", os.path.abspath(base))
    print("標準檔名（模糊比對）：檔名含「AI」「對話」「網址」即可（例：AI對話串網址.docx、AI對話網址.pdf）")
    print("會掃描「含子資料夾」內所有文字檔。")
    print("=" * 70)

    print(f"\n【✓ 命名正確且在根目錄】{len(correct)} 人")
    for name in correct:
        print(f"  ✓ {name}")
    if not correct and not in_subfolder and not wrong_name and not missing:
        print("  （無）")

    print(f"\n【✓ 命名正確但在子資料夾】請要求改放到該生資料夾「根目錄」：{len(in_subfolder)} 人")
    if in_subfolder:
        for name, rel_paths in in_subfolder:
            print(f"  ✓ {name}")
            for rel in rel_paths:
                print(f"      目前位置：「{rel}」")
                print(f"      請將檔案移到：{name}\\ 根目錄（與 企劃、HTML 等資料夾同一層）")
    else:
        print("  （無）")

    print(f"\n【⚠ 需修改命名】有 AI 連結但檔名不正確：{len(wrong_name)} 人")
    if wrong_name:
        for name, files in wrong_name:
            if not files:
                continue
            print(f"  ⚠ {name}")
            for rel, count in files:
                print(f"      目前：「{rel}」 （內含 {count} 個 AI 連結）")
                print(f"      請改檔名為：檔名含「AI」「對話」「網址」（例：AI對話串網址.docx、AI對話網址.pdf）")
    else:
        print("  （無）")

    print(f"\n【✗ 完全沒有】未在任何文字檔（含子資料夾）中找到 AI 聊天串連結：{len(missing)} 人")
    if missing:
        for name in missing:
            print(f"  ✗ {name}")
        print("  以上學生請補交 AI 對話連結（ChatGPT / Gemini 等分享連結）。")
        if not has_pypdf and students_with_pdf:
            in_both = [n for n in missing if n in students_with_pdf]
            if in_both:
                print()
                print("  其中以下學生資料夾內有 .pdf 檔，請先執行：pip install pypdf  後重新執行本檢查。")
                for n in in_both:
                    print("    -", n)
    else:
        print("  （無）")

    print()


if __name__ == "__main__":
    main()
