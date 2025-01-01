# #n<40 or T<1

import pandas as pd
from scipy.stats import fisher_exact

def fisher_exact_test_from_excel(file_path, sheet_name, start_row, start_col):
    """
    从 Excel 文件中指定位置读取四格表，计算 Fisher 确切概率法的双侧和单侧检验 p 值
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

    # 构造观察值表
    observed = [[a, b], [c, d]]

    # Fisher 确切概率检验
    _, p_two_sided = fisher_exact(observed, alternative='two-sided')  # 双侧检验
    _, p_less = fisher_exact(observed, alternative='less')            # 单侧检验：less
    _, p_greater = fisher_exact(observed, alternative='greater')      # 单侧检验：greater

    # 打印结果
    print("观察值四格表:")
    print(pd.DataFrame(observed, columns=["事件发生", "事件未发生"], index=["组1", "组2"]))

    print("\nFisher 确切概率法的 p 值:")
    print(f"双侧检验 p 值: {round(p_two_sided, 4)}")
    print(f"单侧检验 p 值（组1 < 组2）: {round(p_less, 4)}")
    print(f"单侧检验 p 值（组1 > 组2）: {round(p_greater, 4)}")

    # 显著性结论
    print("\n统计学结论:")
    if p_two_sided < 0.05:
        print("双侧检验：差异具有统计学显著性 (p < 0.05)")
    else:
        print("双侧检验：差异不具有统计学显著性 (p ≥ 0.05)")

    if p_less < 0.05:
        print("单侧检验（组1 < 组2）：差异具有统计学显著性 (p < 0.05)")
    else:
        print("单侧检验（组1 < 组2）：差异不具有统计学显著性 (p ≥ 0.05)")

    if p_greater < 0.05:
        print("单侧检验（组1 > 组2）：差异具有统计学显著性 (p < 0.05)")
    else:
        print("单侧检验（组1 > 组2）：差异不具有统计学显著性 (p ≥ 0.05)")


# 替换为Excel 文件路径、表名和数据位置
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为实际文件路径
sheet_name = "四格表确切概率"  # 替换为实际工作表名称
start_row = 2  # 替换为四格表起始行（从1开始数）
start_col = 2  # 替换为四格表起始列（从1开始数）

fisher_exact_test_from_excel(file_path, sheet_name, start_row, start_col)