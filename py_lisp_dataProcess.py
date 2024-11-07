import pandas as pd
import numpy as np
from ast import literal_eval

def process_coordinates(coord_str):
    """处理坐标字符串，转换为元组"""
    try:
        # 移除多余的引号和括号，转换为元组
        coord_str = coord_str.strip('"')
        return literal_eval(coord_str)
    except:
        return None

def calculate_distance(coord1, coord2):
    """计算两个三维坐标点之间的距离"""
    return np.sqrt(sum((a - b) ** 2 for a, b in zip(coord1, coord2)))

def match_numbers_to_devices(devices_df, numbers_df):
    """将数量文本匹配到最近的设备"""
    # 首先打印列名，帮助调试
    print("设备表列名:", devices_df.columns.tolist())
    print("数量表列名:", numbers_df.columns.tolist())
    
    # 添加设备数量列和数量文本坐标列，默认值为1和空值
    devices_df['设备数量'] = 1
    devices_df['数量文本坐标'] = ''
    
    # 处理坐标字符串
    devices_df['coordinates'] = devices_df.iloc[:, 2].apply(process_coordinates)  # 使用第三列作为坐标
    numbers_df['coordinates'] = numbers_df.iloc[:, 1].apply(process_coordinates)  # 使用第二列作为坐标
    
    def extract_number(text):
        try:
            # 移除'X'或'x'前缀，并尝试转换为整数
            text = str(text).strip().lower()
            if text.startswith('x'):
                text = text[1:]
            return int(text)
        except:
            return 1

    numbers_df['数值'] = numbers_df.iloc[:, 0].apply(extract_number)  # 使用第一列作为数量文本
    
    # 为每个数量文本找到最近的设备
    for idx, number_row in numbers_df.iterrows():
        number_coord = number_row['coordinates']
        number_value = number_row['数值']
        
        if number_coord is None:
            continue
            
        min_distance = float('inf')
        closest_device_idx = None
        
        # 计算到每个设备的距离
        for device_idx, device_row in devices_df.iterrows():
            device_coord = device_row['coordinates']
            if device_coord is None:
                continue
                
            distance = calculate_distance(number_coord, device_coord)
            if distance < min_distance:
                min_distance = distance
                closest_device_idx = device_idx
        
        # 如果找到最近的设备，更新其数量和对应的数量文本坐标
        if closest_device_idx is not None and min_distance < 100:
            devices_df.at[closest_device_idx, '设备数量'] = number_value
            devices_df.at[closest_device_idx, '数量文本坐标'] = str(number_coord)
    
    # 删除临时坐标列
    devices_df = devices_df.drop('coordinates', axis=1)
    
    return devices_df

def main():
    try:
        # 读取CSV文件并显示前几行数据
        devices_df = pd.read_csv('devices.csv', encoding='gbk')
        numbers_df = pd.read_csv('numbers.csv', encoding='gbk')
        
        print("\n设备表前几行:")
        print(devices_df.head())
        print("\n数量表前几行:")
        print(numbers_df.head())
        
        # 处理数据
        result_df = match_numbers_to_devices(devices_df, numbers_df)
        
        # 保存结果
        result_df.to_csv('devices_with_quantities.csv', encoding='gbk', index=False)
        print("\n处理完成，结果已保存到 devices_with_quantities.csv")
        
        # 显示结果统计
        print("\n设备数量统计：")
        print(result_df.groupby(result_df.columns[0])['设备数量'].sum())  # 使用第一列作为设备名称，分组求数量和
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
