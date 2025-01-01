import openpyxl

def calculate_average(file_path, sheet_name, cell_range):
    # 打开Excel文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    
    # 获取指定范围的单元格
    cells = sheet[cell_range]
    
    # 提取数值并计算平均值
    values = []
    for row in cells:
        for cell in row:
            if cell.value is not None and isinstance(cell.value, (int, float)):
                values.append(cell.value)
    
    if values:
        average = sum(values) / len(values)
        return average
    else:
        return None

#使用
file_path = "D:\heatmap_data.xlsx"  # 替换为你的Excel文件路径
sheet_name = "Sheet1"            # 替换为你的工作表名称
cell_range = "B2:F6"            # 替换为你的数据范围

average = calculate_average(file_path, sheet_name, cell_range)
if average is not None:
    print(f"指定范围内的平均值是: {average}")
else:
    print("指定范围内没有有效数据")