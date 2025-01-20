# -*- coding: utf-8 -*-
import os
import shutil


def compare_and_copy_files(src_folder, dest_folder):
    # 获取第一个文件夹的文件名（去掉扩展名）
    src_files = {os.path.splitext(f)[0] for f in os.listdir(src_folder) if os.path.isfile(os.path.join(src_folder, f))}

    # 获取第二个文件夹的文件名（去掉扩展名）
    dest_files = {os.path.splitext(f)[0] for f in os.listdir(dest_folder) if
                  os.path.isfile(os.path.join(dest_folder, f))}

    # 找到在源文件夹中但不在目标文件夹中的文件
    files_to_copy = src_files - dest_files

    # 复制文件
    for filename in files_to_copy:
        # 为了复制，构建源和目标文件的完整路径
        src_file_path = os.path.join(src_folder, filename)
        # 复制源文件到目标文件夹，确保保持相同的后缀名
        for ext in ['.txt', '.docx', '.xlsx', '.pdf']: # 你可以添加更多的文件扩展名
            src_file_with_ext = f"{src_file_path}{ext}"
            if os.path.isfile(src_file_with_ext):
                shutil.copy(src_file_with_ext, dest_folder)
                print(f"Copied: {src_file_with_ext} to {dest_folder}")



