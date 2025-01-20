# -*- coding: utf-8 -*-
#doc批量转pdf
import os
from docx2pdf import convert


def batch_convert_doc_to_pdf(doc_folder, pdf_folder):
    if not os.path.exists(pdf_folder):#如果不存在pdf文件夹，则创建
        os.makedirs(pdf_folder)#创建pdf文件夹

    for filename in os.listdir(doc_folder):#遍历doc_folder文件夹下的所有文件
        if filename.endswith('.doc') or filename.endswith('.docx'):#判断文件是否为doc或docx文件
            doc_path = os.path.abspath(os.path.join(doc_folder, filename))#获取文件的绝对路径
            pdf_path = os.path.abspath(os.path.join(pdf_folder, f"{os.path.splitext(filename)[0]}.pdf"))#获取pdf文件的绝对路径

            print(f"Attempting to convert: {doc_path} to {pdf_path}")#打印转换信息

            try:
                convert(doc_path, pdf_path)#调用convert函数进行转换
                print(f"Successfully converted: {filename}")#打印转换成功信息
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")#抛出异常

            # 使用示例

# 使用示例
doc_folder = r'***'  # 替换为你的 Word 文档路径
pdf_folder = r'***'  # 替换为你想保存 PDF 的路径
batch_convert_doc_to_pdf(doc_folder, pdf_folder)