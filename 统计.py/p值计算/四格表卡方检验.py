#n>=40,且所有T>=5时
import pandas as pd
from scipy.stats import chi2_contingency

def chi_square_test_from_excel(file_path, sheet_name, range_start, range_end):
    """
    从 Excel 文件读取四格表数据并进行卡方检验
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param range_start: 数据范围的起始单元格，例如 'A1'
    :param range_end: 数据范围的结束单元格，例如 'B2'
    :return: 打印观察值表、期望频数表、卡方统计量、p 值和统计结论
    """
    # 读取指定范围的数据
    data = pd.read_excel(file_path, sheet_name=sheet_name, usecols=range_start[0] + ':' + range_end[0], nrows=int(range_end[1]) - int(range_start[1]) + 1).values.tolist()

    # 创建观察值表
    observed_data = pd.DataFrame(data, columns=["事件发生", "事件未发生"], index=["组1", "组2"])
    print("观察值表格:")
    print(observed_data)

    # 卡方检验
    chi2, p, dof, expected = chi2_contingency(data)

    # 手动计算每格的贡献值
    chi2_contributions = [
        (data[0][0] - expected[0][0])**2 / expected[0][0],  # (1,1)
        (data[0][1] - expected[0][1])**2 / expected[0][1],  # (1,2)
        (data[1][0] - expected[1][0])**2 / expected[1][0],  # (2,1)
        (data[1][1] - expected[1][1])**2 / expected[1][1]   # (2,2)
    ]
    
    # 手动计算卡方统计量
    chi2_corrected = sum(chi2_contributions)

    # 输出结果
    print("\n期望频数表:")
    print(pd.DataFrame(expected, columns=["事件发生", "事件未发生"], index=["组1", "组2"]))
    print("\n每格的贡献值:", chi2_contributions)
    print("\n卡方统计量 (χ²):", chi2_corrected)
    print("p 值:", p)
    print("自由度:", dof)

    # 显示统计结论
    if p < 0.05:
        print("\n结论: 差异具有统计学显著性 (p < 0.05)")
    else:
        print("\n结论: 差异不具有统计学显著性 (p ≥ 0.05)")

# 调用函数
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为您的 Excel 文件路径
sheet_name = '四格表校正'  # 替换为目标工作表名称
range_start = "B2"  # 替换为数据起始单元格
range_end = "C3"  # 替换为数据结束单元格

chi_square_test_from_excel(file_path, sheet_name, range_start, range_end)