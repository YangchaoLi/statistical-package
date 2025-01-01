#n>=40,1<=T<=5
import pandas as pd
from scipy.stats import chi2

def chi_square_test_with_yates_from_excel(file_path, sheet_name, start_row, start_col):
    """
    从 Excel 文件中指定位置读取四格表，计算校正后的卡方统计量和 p 值
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param start_row: 数据起始行（从 1 开始计数）
    :param start_col: 数据起始列（从 1 开始计数）
    """
    # 读取 Excel 数据
    data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # 获取四格表数据 (a, b, c, d)
    a = data.iloc[start_row - 1, start_col - 1]
    b = data.iloc[start_row - 1, start_col]
    c = data.iloc[start_row, start_col - 1]
    d = data.iloc[start_row, start_col]

    # 观察值四格表
    observed = [[a, b], [c, d]]

    # 总计和期望频数
    total = a + b + c + d
    expected = [[(sum(row) * sum(col)) / total for col in zip(*observed)] for row in observed]

    # Yates 校正的卡方统计量
    chi2_corrected = sum(
        (abs(observed[i][j] - expected[i][j]) - 0.5) ** 2 / expected[i][j]
        for i in range(2) for j in range(2)
    )

    # 自由度
    dof = 1

    # 矫正后的 p 值
    p_value_corrected = 1 - chi2.cdf(chi2_corrected, df=dof)

    # 打印结果
    print("观察值四格表:")
    print(pd.DataFrame(observed, columns=["事件发生", "事件未发生"], index=["组1", "组2"]))
    print("\n期望频数表:")
    print(pd.DataFrame(expected, columns=["事件发生", "事件未发生"], index=["组1", "组2"]))
    print("\n矫正后的卡方统计量 (χ²):", round(chi2_corrected, 4))
    print("矫正后的 p 值:", round(p_value_corrected, 4))
    print("自由度:", dof)

    # 显著性结论
    if p_value_corrected < 0.05:
        print("\n结论: 差异具有统计学显著性 (p < 0.05)")
    else:
        print("\n结论: 差异不具有统计学显著性 (p ≥ 0.05)")


# 替换为 Excel 文件路径、表名和数据位置
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为实际文件路径
sheet_name = "四格表校正"  # 替换为实际工作表名称
start_row =2  # 替换为四格表起始行（从1开始数）
start_col = 2  # 替换为四格表起始列（从1开始数）

chi_square_test_with_yates_from_excel(file_path, sheet_name, start_row, start_col)