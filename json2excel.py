import json
import pandas as pd

# 函数：将RGB值从0-1范围转换为HEX值
def rgb_to_hex(rgb):
    return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))

# 读取JSON文件
with open('colors.json') as f:
    data = json.load(f)

# 处理数据
index = 1
colors_data = []
for item in data:
    name = item['name']  # Unicode名称
    color_type = item['type']
    mode = item['data']['mode']
    values = item['data']['values']
    hex_value = rgb_to_hex(values)  # 转换RGB到HEX
    colors_data.append({'index': index, 'name': name, 'hex_value': hex_value, 'type': color_type, 'mode': mode, 'rgb_value': values})
    index += 1 

# 将处理后的数据转换为DataFrame
df = pd.DataFrame(colors_data)

# 写入Excel文件
df.to_excel('colors_converted.xlsx', index=False)