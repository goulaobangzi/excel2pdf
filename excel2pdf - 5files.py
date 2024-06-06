import glob
import os
import comtypes.client
from PyPDF2 import PdfWriter

def get_file_paths():
	folder_path = os.getcwd()  # 获取当前工作目录
	# 搜索当前目录及所有子目录的Excel文件
	excel_files = glob.glob(os.path.join(folder_path, '**/*.xls*'), recursive=True)
	return excel_files



def convert_sheets_to_pdf(file_paths):
	# 初始化Excel应用程序
	excel = comtypes.client.CreateObject("Excel.Application")
	excel.DisplayAlerts = False
	excel.Visible = True
	for file_path in file_paths:
		print("正在处理: " + file_path)
		pdf_files = []
		wb = excel.Workbooks.Open(file_path)
		for sheet_name in ['W1', 'W2', 'W3', 'W4', 'W5']:
			try:
				ws = wb.Worksheets(sheet_name)
				ws.PageSetup.PrintArea = 'A1:N58'
				ws.Activate()
				# 导出为PDF
				output_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_{sheet_name}.pdf"
				output_path = os.path.join(os.path.dirname(file_path), output_filename)
				pdf_files.append(output_path)
				print("正在保存： " + output_path)
				if os.path.exists(output_path):
					print("文件已存在，覆盖保存: " + output_path)
					os.remove(output_path)
				ws.ExportAsFixedFormat(0, output_path)
			except Exception as e:
				print(f"Error processing {sheet_name} in {file_path}: {str(e)}")
        # 合并PDF文件
		for pdf_file in pdf_files:
			with open(pdf_file, "rb") as f:
				pdf_reader = PyPDF2.PdfReader(f)
				for page in range(len(pdf_reader.pages)):
					pdf_writer.add_page(pdf_reader.pages[page])
        # 删除单页PDF文件
		for pdf_file in pdf_files:
			os.remove(pdf_file)
            # 关闭工作簿
		wb.Close(SaveChanges=False)
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
