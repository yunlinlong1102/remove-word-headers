import os
import sys
from docx import Document

def remove_header_from_docx(file_path):
    try:
        doc = Document(file_path)

        for section in doc.sections:
            header = section.header
            header_element = header._element
            for child in list(header_element):
                header_element.remove(child)

        doc.save(file_path)
        print(f"已处理：{file_path}")

    except Exception as e:
        print(f"处理失败：{file_path}，原因：{e}")

def process_all_docx(root_folder):
    for root, dirs, files in os.walk(root_folder):
        for file in files:

            # ✅ 跳过 Word 临时文件
            if file.startswith("~$"):
                continue

            if file.lower().endswith(".docx"):
                full_path = os.path.join(root, file)
                remove_header_from_docx(full_path)

if __name__ == "__main__":

    if getattr(sys, 'frozen', False):
        ROOT_FOLDER = os.path.dirname(sys.executable)
    else:
        ROOT_FOLDER = os.path.dirname(os.path.abspath(__file__))

    print(f"开始处理目录：{ROOT_FOLDER}")
    process_all_docx(ROOT_FOLDER)
    print("全部处理完成 ✅")
    input("按回车键退出...")