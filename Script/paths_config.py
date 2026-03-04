# -*- coding: utf-8 -*-
"""
集中定義「學生作業」目錄結構，供各腳本使用。
腳本位於 Script/ 下，學生作業根目錄為上一層。
輸出：資料 → Data/，說明／LOG → Doc/，藝廊網頁與資料 → Gallery/。
掃描學生資料夾時排除：Script, Data, Doc, Gallery, __pycache__
"""
import os

# 學生作業根目錄（Script 的上一層）
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(_SCRIPT_DIR)

DATA_DIR = os.path.join(BASE_DIR, "Data")
DOC_DIR = os.path.join(BASE_DIR, "Doc")
GALLERY_DIR = os.path.join(BASE_DIR, "Gallery")

# 掃描學生資料夾時要排除的目錄（非學生作業）
SKIP_FOLDERS = {"Script", "Data", "Doc", "Gallery", "__pycache__"}


def is_student_folder(name):
    """若為排除目錄或非目錄則不當成學生資料夾。"""
    return name not in SKIP_FOLDERS and not name.startswith(".")


def iter_student_folders(base=None):
    """依序 yield 學生資料夾名稱（已排序，且排除 SKIP_FOLDERS）。"""
    base = base or BASE_DIR
    for name in sorted(os.listdir(base)):
        path = os.path.join(base, name)
        if not os.path.isdir(path):
            continue
        if not is_student_folder(name):
            continue
        yield name
