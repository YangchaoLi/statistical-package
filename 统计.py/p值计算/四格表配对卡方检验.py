# 配对四格表的卡方检验
import pandas as pd
from scipy.stats import chi2

def paired_chi_square_test_full_table(file_path, sheet_name, start_row, start_col):
    """
    从 Excel 文件中读取完整四格表，提取 b 和 c 数据，计算配对卡方检验
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param start_row: 四格表起始行（从 1 开始计数）
    :param start_col: 四格表起始列（从 1 开始计数）
    """
    # 读取 Excel 数据
    data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    # 提取四格表数据
    a = data.iloc[start_row - 1, start_col - 1]
    b = data.iloc[start_row - 1, start_col]
    c = data.iloc[start_row, start_col - 1]
    d = data.iloc[start_row, start_col]

    # 检查四格表数据是否合理
    if any(x < 0 for x in [a, b, c, d]):
        print("四格表数据无效：所有值必须为非负整数。")
        return

    # 打印完整四格表
    print("完整四格表:")
    print(pd.DataFrame([[a, b], [c, d]], columns=["+", "-"], index=["+", "-"]))

    # 分母
    total = b + c

    # 检查总样本量
    if total == 0:
        print("b 和 c 的总和为 0，无法计算。")
        return

    # 分子根据总数进行调整
    if total >= 40:
        chi2_stat = (abs(b - c) ** 2) / total
    else:
        chi2_stat = ((abs(b - c) - 1) ** 2) / total

    # 自由度
    dof = 1

    # p 值
    p_value = 1 - chi2.cdf(chi2_stat, df=dof)

    # 打印结果
    print("\n卡方统计量 (χ²):", round(chi2_stat, 4))
    print("p 值:", round(p_value, 4))
    print("自由度:", dof)

    # 显著性结论
    if p_value < 0.05:
        print("\n结论: 差异具有统计学显著性 (p < 0.05)")
    else:
        print("\n结论: 差异不具有统计学显著性 (p ≥ 0.05)")

# 替换为Excel 文件路径、表名和数据位置
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为实际文件路径
sheet_name = "配对四格表"  # 替换为实际工作表名称
start_row = 3  # 替换为配对四格表起始行（1-based）
start_col = 2  # 替换为配对四格表起始列（1-based）

paired_chi_square_test_full_table(file_path, sheet_name, start_row, start_col)