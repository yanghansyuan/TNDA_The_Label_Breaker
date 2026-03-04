# -*- coding: utf-8 -*-
"""
檢查「學生作業」下每位學生資料夾（含子資料夾）是否至少有一個檔名含 "final" 的 .html/.htm 檔。
列出：有 final.html 的學生、檔名含 final 但副檔名錯誤的、沒有 final 的學生。
掃描時排除：Script、Data、Doc、Gallery、__pycache__。
"""
import os

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


def find_final_in_folder(folder_path):
    """
    在該資料夾與所有子資料夾中尋找檔名含 "final" 的檔案。
    回傳 (html_list, wrong_ext_list)
    - html_list: [(相對路徑, 絕對路徑), ...]  副檔名為 .html 或 .htm
    - wrong_ext_list: [(相對路徑, 檔名), ...]  檔名含 final 但副檔名不是 .html/.htm
    """
    html_list = []
    wrong_ext_list = []
    folder_path = os.path.normpath(folder_path)
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            name_lower = f.lower()
            if "final" not in name_lower:
                continue
            abs_path = os.path.join(root, f)
            rel_path = os.path.relpath(abs_path, folder_path)
            if name_lower.endswith(".html") or name_lower.endswith(".htm"):
                html_list.append((rel_path, abs_path))
            else:
                wrong_ext_list.append((rel_path, f))
    return html_list, wrong_ext_list


def main():
    has_final = []
    wrong_extension = []
    no_final = []

    for name in iter_student_folders():
        folder = os.path.join(BASE_DIR, name)
        html_list, wrong_ext_list = find_final_in_folder(folder)
        if html_list:
            has_final.append((name, html_list))
        if wrong_ext_list:
            wrong_extension.append((name, wrong_ext_list))
        if not html_list and not wrong_ext_list:
            no_final.append(name)

    print("=" * 60)
    print("檢查：各學生資料夾（含子資料夾）是否含有「final」的 .html/.htm 檔")
    print("=" * 60)
    print()
    print(f"【有 final 的學生】共 {len(has_final)} 位（檔名含 final 且為 .html/.htm）")
    for name, paths in has_final:
        print(f"  ✓ {name}")
        for rel, _ in paths:
            print(f"      → {rel}")
    print()
    print(f"【檔名含 final 但副檔名不是 .html】共 {len(wrong_extension)} 位（需請學生改為 .html 或重新匯出）")
    for name, paths in wrong_extension:
        print(f"  ⚠ {name}")
        for rel, fname in paths:
            print(f"      → {rel}")
    if not wrong_extension:
        print("  （無）")
    print()
    print(f"【沒有 final 的學生】共 {len(no_final)} 位（沒有任何檔名含 final 的檔案）")
    for name in no_final:
        print(f"  ✗ {name}")
    if not no_final:
        print("  （無）")
    print()


if __name__ == "__main__":
    main()
