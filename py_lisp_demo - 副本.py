from pyautocad import Autocad, APoint
import pandas as pd
import os

def run_lisp_function(function_name):
	"""
	在AutoCAD中运行指定的LISP函数。
	"""
	acad = Autocad(create_if_not_exists=True)
	acad.prompt(f"\n运行LISP函数: {function_name}\n")
	# 注意：请将路径改为你的实际路径
	acad.doc.SendCommand(f"(load \"./format(2).lsp\") (c:{function_name}) ")

def read_text_coordinates():
	"""
	调用getAllTextWithCoords函数并读取生成的CSV文件
	"""
	run_lisp_function("getAllTextWithCoords")
	# 等待文件生成
	if os.path.exists("numbers.csv"):
		df = pd.read_csv("numbers.csv", encoding='gbk')
		return df
	return None

def read_device_blocks():
	"""
	调用getAllDeviceBlocks函数并读取生成的CSV文件
	"""
	run_lisp_function("getAllDeviceBlocks")
	# 等待文件生成
	if os.path.exists("devices.csv"):
		df = pd.read_csv("devices.csv", encoding='gbk')
		return df
	return None

if __name__ == "__main__":
	# 获取文本坐标
	text_df = read_text_coordinates()
	if text_df is not None:
		print("\n文本坐标数据：")
		print(text_df)
	
	# 获取设备块信息
	devices_df = read_device_blocks()
	if devices_df is not None:
		print("\n设备块数据：")
		print(devices_df)

