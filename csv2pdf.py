import os
import win32com.client


def convert_xlsx_to_pdf(folder_path):
    # 初始化 Excel 应用
    excel = win32com.client.Dispatch("Excel.Application")

    # 遍历文件夹中的所有 .xlsx 文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            xlsx_file_path = os.path.join(folder_path, filename)
            pdf_file_path = os.path.splitext(xlsx_file_path)[0] + '.pdf'

            try:
                # 打开 Excel 文件
                workbook = excel.Workbooks.Open(xlsx_file_path)
                # 导出为 PDF
                workbook.ExportAsFixedFormat(0, pdf_file_path)
                print(f"成功转换: '{xlsx_file_path}' -> '{pdf_file_path}'")
            except Exception as e:
                print(f"转换 '{xlsx_file_path}' 时出错: {e}")
            finally:
                # 确保 workbook 被关闭
                if 'workbook' in locals():
                    workbook.Close(False)  # 不保存更改

    # 退出 Excel 应用
    excel.Quit()

#使用处理的文件夹路径
folder_path = r''  # 文件夹路径
convert_xlsx_to_pdf(folder_path)