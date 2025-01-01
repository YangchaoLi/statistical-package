import openpyxl
import statistics

def calculate_median(file_path, sheet_name, cell_range):
    # 打开Excel文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    
    # 获取指定范围的单元格
    cells = sheet[cell_range]
    
    # 提取数值并计算中位数
    values = []
    for row in cells:
        for cell in row:
            if cell.value is not None and isinstance(cell.value, (int, float)):
                values.append(cell.value)
    
    if values:
        median = statistics.median(values)
        return median
    else:
        return None

# 示例使用
file_path = "D:\heatmap_data.xlsx"  # 替换为你的Excel文件路径
sheet_name = "Sheet1"            # 替换为你的工作表名称
cell_range = "B2:F6"            # 替换为你的数据范围

median = calculate_median(file_path, sheet_name, cell_range)
if median is not None:
    print(f"指定范围内的中位数是: {median}")
else:
    print("指定范围内没有有效数据")