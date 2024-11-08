import pandas as pd
import numpy as np
from ast import literal_eval
import os

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
    print("\n=== 数据处理过程 ===")
    print("设备表列名:", devices_df.columns.tolist())
    print("数量表列名:", numbers_df.columns.tolist())
    
    # 清理列名中的空格
    devices_df.columns = devices_df.columns.str.strip()
    numbers_df.columns = numbers_df.columns.str.strip()
    
    # 添加设备数量列、数量文本坐标列和匹配距离列，默认值为1、空值和0
    devices_df['设备数量'] = 1
    devices_df['数量文本坐标'] = ''
    devices_df['匹配距离'] = 0.0  # 新增匹配距离列
    
    # 处理坐标字符串并打印调试信息
    print("\n处理坐标中...")
    devices_df['coordinates'] = devices_df['坐标'].apply(process_coordinates)
    numbers_df['coordinates'] = numbers_df['坐标'].apply(process_coordinates)   
  
    def extract_number(text):
        """提取数量文本中的数值"""
        # 移除'X'或'x'或'*'前缀，并尝试转换为整数
        try:
            text = str(text).strip().lower()
            if text.startswith('x') or text.startswith('*'):
                text = text[1:]
            num = int(text)
            return num
        except:
            print(f"无法提取数值: {text}, 使用默认值1")
            return 1

    numbers_df['数值'] = numbers_df.iloc[:, 0].apply(extract_number)
    
    # 打印匹配过程
    print("\n开始匹配设备和数量...")
    match_count = 0
    
    # 创建已匹配设备的集合和未匹配数量的列表
    matched_devices = set()
    unmatched_numbers = []
    
    # 计算坐标的量纲
    device_coords = [coord for coord in devices_df['coordinates'] if coord is not None]
    if device_coords:
        x_coords = [coord[0] for coord in device_coords]
        y_coords = [coord[1] for coord in device_coords]
        x_range = max(x_coords) - min(x_coords)
        y_range = max(y_coords) - min(y_coords)
        
        # 计算平均设备间距
        avg_x_spacing = x_range / (len(set(x_coords)) + 1)
        avg_y_spacing = y_range / (len(set(y_coords)) + 1)
        
        # 使用较大的间距作为基础，避免y方向间距过小的影响
        base_threshold = max(avg_x_spacing, avg_y_spacing) * 0.5
        
        # 根据坐标量纲调整阈值
        magnitude = max(abs(max(x_coords)), abs(min(x_coords)), 
                       abs(max(y_coords)), abs(min(y_coords)))
        
        # 增加量纲调整的影响
        if magnitude > 1000:
            # 使用更大的调整系数
            distance_threshold = base_threshold * (2 + np.log10(magnitude/100))
        else:
            distance_threshold = base_threshold
            
        print(f"坐标量纲: {magnitude:.2f}")
        print(f"平均设备间距: X={avg_x_spacing:.2f}, Y={avg_y_spacing:.2f}")
        print(f"基础阈值: {base_threshold:.2f}")
        print(f"最终距离阈值: {distance_threshold:.2f}")
    else:
        distance_threshold = 500
    
    # 为了调试，先打印一些示例距离
    print("\n示例距离计算:")
    for i, number_row in numbers_df.head(3).iterrows():
        number_coord = number_row['coordinates']
        if number_coord is None:
            continue
        for j, device_row in devices_df.head(3).iterrows():
            device_coord = device_row['coordinates']
            if device_coord is None:
                continue
            distance = calculate_distance(number_coord, device_coord)
            print(f"数量{number_row['数值']}({number_coord}) 到 设备({device_coord}) 的距离: {distance:.2f}")
    
    # 为每个数量文本找到最近的未匹配设备
    for idx, number_row in numbers_df.iterrows():
        number_coord = number_row['coordinates']
        number_value = number_row['数值']
        
        if number_coord is None:
            continue
            
        min_distance = float('inf')
        closest_device_idx = None
        
        # 计算到每个未匹配设备的距离
        for device_idx, device_row in devices_df.iterrows():
            if device_idx in matched_devices:
                continue
                
            device_coord = device_row['coordinates']
            if device_coord is None:
                continue
                
            distance = calculate_distance(number_coord, device_coord)
            
            # 只有当距离小于阈值时才考虑匹配
            if distance < distance_threshold and distance < min_distance:
                min_distance = distance
                closest_device_idx = device_idx
        
        # 如果找到最近的未匹配设备，更新其数量和对应的数量文本坐标
        if closest_device_idx is not None:
            devices_df.at[closest_device_idx, '设备数量'] = number_value
            devices_df.at[closest_device_idx, '数量文本坐标'] = str(number_coord)
            devices_df.at[closest_device_idx, '匹配距离'] = min_distance  # 保存匹配距离
            matched_devices.add(closest_device_idx)
            match_count += 1
            print(f"匹配成功: 数量{number_value}在{number_coord}匹配到设备{devices_df.at[closest_device_idx, '设备类型']}在{device_coord},距离{min_distance:.2f}")
        else:
            unmatched_numbers.append({
                '数值': number_value,
                '坐标': number_coord
            })

    print(f"\n成功匹配 {match_count} 个数量到设备")
    print(f"未匹配数量文本: {len(unmatched_numbers)} 个")
    if unmatched_numbers:
        print("\n未匹配的数量文本详情:")
        for item in unmatched_numbers:
            print(f"数值: {item['数值']}, 坐标: {item['坐标']}")
    
    # 删除临时坐标列
    devices_df = devices_df.drop('coordinates', axis=1)
    
    return devices_df

def process_data():
    """
    处理CSV数据的主函数
    """
    try:
        # 修改CSV文件读取路径
        csv_path = "e:\\temp\\"
        devices_df = pd.read_csv(os.path.join(csv_path, 'devices.csv'), encoding='gbk')
        numbers_df = pd.read_csv(os.path.join(csv_path, 'numbers.csv'), encoding='gbk')
        
        # 打印CSV文件的前3行内容
        print("\n=== devices.csv 前3行 ===")
        print(devices_df.head(3))
        print("\n=== numbers.csv 前3行 ===")
        print(numbers_df.head(3))
        
        print("\n读取到的数据概况：")
        print(f"设备表：{len(devices_df)}行")
        print(f"数量表：{len(numbers_df)}行")
        
        # 处理数据
        result_df = match_numbers_to_devices(devices_df, numbers_df)
        
        # 打印处理后的结果前3行
        print("\n=== devices_with_quantities.csv 前3行 ===")
        print(result_df.head(3))
        
        # 计算设备数量统计
        quantity_stats = result_df.groupby('设备类型')['设备数量'].agg(['count', 'sum']).reset_index()
        quantity_stats.columns = ['设备类型', '设备图元(个)', '设备总数']
        
        # 使用ExcelWriter保存多个sheet
        excel_path = 'devices_with_quantities.xlsx'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name='设备明细', index=False)
            quantity_stats.to_excel(writer, sheet_name='设备统计', index=False)
        
        print("\n处理完成，结果已保存到", excel_path)
        
        # 显示结果统计
        print("\n设备数量统计：")
        print(quantity_stats)
        
        # 调用格式化函数
        try:
            from format_excel import format_excel
            format_excel(excel_path)
        except Exception as e:
            print(f"格式化Excel时出错: {str(e)}")
        
        return True
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return False

if __name__ == "__main__":
    process_data()
