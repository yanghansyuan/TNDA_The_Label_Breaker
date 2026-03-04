# -*- coding: utf-8 -*-
"""
掃描「學生作業」資料夾，依學生姓名排序後印出所有人。
執行：python list_students_sorted.py
結果寫入 Doc/list_students_sorted.txt（UTF-8）。掃描時排除 Script/Data/Doc/Gallery/__pycache__。
"""
import os

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


def main():
    students = []

    for name in iter_student_folders():
        if "_" not in name:
            continue
        student_name, emotion = name.split("_", 1)
        if not student_name.strip() or not emotion.strip():
            continue
        students.append((student_name, emotion, name))

    students.sort(key=lambda x: x[0])  # 依姓名排序

    os.makedirs(DOC_DIR, exist_ok=True)
    out_path = os.path.join(DOC_DIR, "list_students_sorted.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("學生姓名（依姓名排序），共 {} 人：\n".format(len(students)))
        f.write("-" * 40 + "\n")
        for student_name, emotion, folder_name in students:
            f.write("{}  （{}）\n".format(student_name, emotion))
        f.write("-" * 40 + "\n")
        f.write("共 {} 人\n".format(len(students)))

    print("已寫入 " + out_path)
    print("共 {} 人".format(len(students)))


if __name__ == "__main__":
    main()
