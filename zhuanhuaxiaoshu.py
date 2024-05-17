import pandas as pd
import re

def dms_to_decimal(dms_str, default_direction):
    """
    将DMS（度、分、秒）格式转换为十进制度数格式。
    
    参数:
    dms_str (str): 以DMS格式表示的经纬度字符串，例如 "123° 45' 6\""
    default_direction (str): 默认方向，可以是"N", "S", "E", "W"
    
    返回:
    float: 十进制度数格式的经纬度
    """
    dms_str = dms_str.strip()
    
    # 如果输入字符串为空，返回None
    if not dms_str:
        return None

    # 替换可能出现的非标准符号
    dms_str = dms_str.replace('′', "'").replace('″', '"').replace(' ', '')
    
    # 使用正则表达式匹配度、分、秒部分
    pattern = re.compile(r'(\d+)°(\d+)\'(\d+(?:\.\d+)?)\"')
    match = pattern.match(dms_str)

    if not match:
        raise ValueError(f"Invalid DMS format: {dms_str}")

    degrees = float(match.group(1))
    minutes = float(match.group(2))
    seconds = float(match.group(3))

    # 计算十进制度数
    decimal = degrees + minutes / 60 + seconds / 3600

    # 根据默认方向调整符号，南纬和西经为负
    if default_direction in ['S', 'W']:
        decimal = -decimal

    return decimal

# 输入文件路径
input_file = 'C:/Users/Zhangyl/Desktop/1.xlsx'

try:
    # 读取Excel文件到DataFrame
    df = pd.read_excel(input_file)
    
    # 打印列名以确认正确的列名
    print("列名:")
    print(df.columns)
    
    # 假设经度和纬度分别为经度和纬度的列名，请根据实际情况修改
    longitude_col = '经度'
    latitude_col = '纬度'
    
    # 检查列名是否在DataFrame中
    if longitude_col not in df.columns or latitude_col not in df.columns:
        print(f"列名 '{longitude_col}' 或 '{latitude_col}' 不存在，请确认列名是否正确。")
    else:
        # 将经度和纬度列的DMS格式转换为十进制度数格式，假设经度为东经，纬度为北纬
        df['经度_decimal'] = df[longitude_col].apply(lambda x: dms_to_decimal(x, 'E'))
        df['纬度_decimal'] = df[latitude_col].apply(lambda x: dms_to_decimal(x, 'N'))

        # 打印转换后的数据（可选）
        print("转换后的数据:")
        print(df[[longitude_col, '经度_decimal', latitude_col, '纬度_decimal']])
        
        # 输出文件路径
        output_file = 'output.xlsx'

        # 将转换后的DataFrame保存到新的Excel文件
        df.to_excel(output_file, index=False)
        
        # 打印确认信息
        print(f"转换后的数据已保存到 {output_file}")
except FileNotFoundError:
    print(f"文件 {input_file} 未找到，请确认文件路径是否正确。")
except Exception as e:
    print(f"发生错误: {e}")
