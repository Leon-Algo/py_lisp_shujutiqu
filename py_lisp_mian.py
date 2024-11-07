from pyautocad import Autocad, APoint
import pandas as pd
import os
import win32com.client
import time

def open_target_drawing():
	"""
	打开目标CAD图纸
	"""
	# target_dwg = "P-21040185陕西榆林体育中心及会展中心项目—能耗管理系统.dwg"
	target_dwg = "P-23040167深圳平湖智造园项目 工厂区电能管理系统图_t3（系统图）.dwg"
	
	try:
		# 获取当前目录的完整路径
		current_dir = os.path.abspath(os.path.dirname(__file__))
		dwg_path = os.path.join(current_dir, target_dwg)
		
		if not os.path.exists(dwg_path):
			print(f"错误：找不到目标图纸 {target_dwg}")
			return False
			
		# 使用win32com获取或创建AutoCAD实例
		try:
			acad = win32com.client.GetActiveObject("AutoCAD.Application")
		except:
			print("正在启动AutoCAD...")
			acad = win32com.client.Dispatch("AutoCAD.Application")
			acad.Visible = True
			time.sleep(2)  # 等待AutoCAD启动
		
		# 检查当前是否已经打开了目标图纸
		try:
			current_dwg = acad.ActiveDocument.Name
			if target_dwg in current_dwg:
				print("目标图纸已经打开")
				return True
		except:
			pass
		
		# 打开目标图纸
		print(f"正在打开图纸: {target_dwg}")
		acad.Documents.Open(dwg_path)
		time.sleep(2)  # 等待图纸打开
		print("图纸打开成功")
		
		# 确保图纸被激活
		for doc in acad.Documents:
			if target_dwg in doc.Name:
				doc.Activate()
				break
				
		return True
		
	except Exception as e:
		print(f"打开图纸时出错: {str(e)}")
		return False

def run_lisp_function(function_name):
	"""
	在AutoCAD中运行指定的LISP函数。
	"""
	try:
		# 使用pyautocad获取当前的AutoCAD实例
		acad = Autocad(create_if_not_exists=False)  # 不要创建新实例
		acad.prompt(f"\n运行LISP函数: {function_name}\n")
		# 使用绝对路径加载LISP文件
		current_dir = os.path.abspath(os.path.dirname(__file__))
		lisp_path = os.path.join(current_dir, "format(2).lsp")
		acad.doc.SendCommand(f"(load \"{lisp_path}\") (c:{function_name}) ")
		time.sleep(1)  # 给LISP函数执行时间
	except Exception as e:
		print(f"运行LISP函数时出错: {str(e)}")
		return False
	return True

def collect_data():
	"""
	收集CAD中的数据并生成CSV文件
	"""
	# 获取文本坐标
	if not run_lisp_function("getAllTextWithCoords"):
		return False
		
	# 获取设备块信息
	if not run_lisp_function("getAllDeviceBlocks"):
		return False
	
	# 检查文件是否生成成功
	files_exist = True
	csv_path = "e:\\temp\\"  # 新增CSV文件路径
	if not os.path.exists(os.path.join(csv_path, "numbers.csv")):
		print("错误：未能生成numbers.csv文件")
		files_exist = False
	if not os.path.exists(os.path.join(csv_path, "devices.csv")):
		print("错误：未能生成devices.csv文件")
		files_exist = False
	
	return files_exist

def main():
	"""
	主函数：执行完整的数据收集和处理流程
	"""
	print("\n开始处理CAD图纸...")
	
	# 首先确保目标图纸已打开
	if not open_target_drawing():
		print("无法打开目标图纸，程序终止")
		return
	
	print("\n开始从CAD中收集数据...")
	if collect_data():
		print("\nCAD数据收集完成，开始处理数据...")
		# 调用数据处理模块
		from py_lisp_dataProcess import process_data
		process_data()
	else:
		print("\n数据收集失败，请检查CAD环境和LISP函数。")

if __name__ == "__main__":
	main()

