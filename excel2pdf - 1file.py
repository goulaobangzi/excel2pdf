import glob
import os
import win32com.client as win32

def get_file_paths():
	folder_path = os.getcwd()  # 获取当前工作目录
	# 搜索当前目录及所有子目录的Excel文件
	excel_files = glob.glob(os.path.join(folder_path, '**/*.xls*'), recursive=True)
	return excel_files

def convert_sheets_to_pdf(file_paths):
	# 初始化Excel应用程序
	excel = win32.Dispatch("Excel.Application")
	excel.DisplayAlerts = False
	excel.Visible = True
	for file_path in file_paths:
		print("正在处理: " + file_path)
		workbook = excel.Workbooks.Open(file_path)
		pdf_path = file_path[:-5]
		workbook.Sheets(['W1', 'W2', 'W3', 'W4', 'W5']).Select()
		print("正在输出: " + pdf_path)
		workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
            # 关闭工作簿
		workbook.Close(SaveChanges=False)
    # 退出Excel应用程序
	excel.Quit()

# 获取脚本文件的完整路径
script_path = os.path.abspath(__file__)

# 获取脚本文件所在的目录
script_dir = os.path.dirname(script_path)

# 改变当前工作目录到脚本文件所在目录
os.chdir(script_dir)

# 安装必要库： pip install comtypes
file_paths = get_file_paths()
convert_sheets_to_pdf(file_paths)
