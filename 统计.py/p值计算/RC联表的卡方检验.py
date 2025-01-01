#R*C列联表资料的卡方检验
import pandas as pd
from scipy.stats import chi2_contingency
from decimal import Decimal
def chi_square_test_from_excel(file_path, sheet_name, start_row, start_col, end_row, end_col):
    """
    从 Excel 文件中指定范围读取 R×C 列联表，计算卡方检验统计量和 p 值
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param start_row: 列联表起始行（从 1 开始计数）
    :param start_col: 列联表起始列（从 1 开始计数）
    :param end_row: 列联表结束行（从 1 开始计数）
    :param end_col: 列联表结束列（从 1 开始计数）
    """
    try:
        # 读取 Excel 数据
        data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # 提取列联表数据
        contingency_table = data.iloc[start_row - 1:end_row, start_col - 1:end_col]

        # 检查数据是否包含非数值
        if not contingency_table.applymap(lambda x: isinstance(x, (int, float))).all().all():
            raise ValueError("列联表数据包含非数值，请检查 Excel 表格内容。")

        # 将数据转换为 NumPy 数组
        contingency_table = contingency_table.values.astype(float)

        # 打印列联表
        print("输入列联表数据:")
        print(pd.DataFrame(contingency_table))

        # 进行卡方检验
        chi2_stat, p_value, dof, expected = chi2_contingency(contingency_table)
        p_value_decimal = Decimal(p_value)

        # 输出结果
        print("\n卡方统计量 (χ²):", round(chi2_stat, 4))
        print("p 值:", round(p_value_decimal, 10))
        print("自由度:", dof)
        print("\n期望频数表:")
        print(pd.DataFrame(expected).round(2))
        
        # 显著性结论
        if p_value < 0.05:
            print("\n结论: 差异具有统计学显著性 (p < 0.05)")
        else:
            print("\n结论: 差异不具有统计学显著性 (p ≥ 0.05)")

    except Exception as e:
        print("运行错误:", str(e))

# 替换为文件路径、表名和数据位置
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为实际文件路径
sheet_name = "联表的卡方检验"  # 替换为实际工作表名称
start_row = 2  # 替换为列联表起始行（1-based）
start_col = 3  # 替换为列联表起始列（1-based）
end_row = 4   # 替换为列联表结束行（1-based）
end_col = 4   # 替换为列联表结束列（1-based）

chi_square_test_from_excel(file_path, sheet_name, start_row, start_col, end_row, end_col)