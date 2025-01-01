# import pandas as pd
# import numpy as np

# def livak_analysis(file_path, sheet_name, group_column, value_column):
#     """
#     使用 Livak 法分析组间差异性。
    
#     :param file_path: str，Excel 文件路径
#     :param sheet_name: str，工作表名称
#     :param group_column: str，表示组别的列名
#     :param value_column: str，表示数据值的列名
#     :return: pd.DataFrame，包含组均值、标准差和差异分析结果
#     """
#     # 读取 Excel 文件
#     data = pd.read_excel(file_path, sheet_name=sheet_name)
    
#     # 检查必要的列是否存在
#     if group_column not in data.columns or value_column not in data.columns:
#         raise ValueError("指定的列不存在，请检查输入的列名。")
    
#     # 按组进行分组
#     grouped = data.groupby(group_column)[value_column]
    
#     # 计算均值和标准差
#     results = grouped.agg(['mean', 'std', 'count']).reset_index()
#     results.columns = [group_column, 'mean', 'std', 'count']
    
#     # 对两组进行差异性分析（假设只有两组）
#     if len(results) != 2:
#         raise ValueError("Livak 法仅支持两组数据的分析，请检查组数。")
    
#     mean_diff = results['mean'].diff().iloc[-1]
#     pooled_std = np.sqrt(((results.loc[0, 'std'] ** 2 + results.loc[1, 'std'] ** 2) / 2))
#     fold_change = np.exp(mean_diff / pooled_std)
    
#     # 将差异性结果添加到 DataFrame
#     results['mean_diff'] = [None, mean_diff]#均值差异
#     results['fold_change'] = [None, fold_change]#倍数变化
    
#     return results

# # 使用
# file_path = r"C:\Users\huawei\OneDrive\桌面\工作簿1.xlsx"  # Excel 文件路径
# sheet_name = "Sheet1"       # 工作表名称
# group_column = "Group"      # 分组列名
# value_column = "Value"      # 数据列名

# try:
#     result = livak_analysis(file_path, sheet_name, group_column, value_column)
#     print("分析结果：")
#     print(result)
# except Exception as e:
#     print(f"分析失败：{e}")

import pandas as pd
import numpy as np
from scipy.stats import ttest_ind

def livak_analysis_with_stats(file_path, sheet_name, group_column, value_column):
    """
    使用 Livak 法分析组间差异性，计算 δT 和 p 值。
    
    :param file_path: str，Excel 文件路径
    :param sheet_name: str，工作表名称
    :param group_column: str，表示组别的列名
    :param value_column: str，表示数据值的列名
    :return: pd.DataFrame，包含组均值、标准差、δT 和 p 值等结果
    """
    # 读取 Excel 文件
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 检查必要的列是否存在
    if group_column not in data.columns or value_column not in data.columns:
        raise ValueError("指定的列不存在，请检查输入的列名。")
    
    # 按组进行分组
    grouped = data.groupby(group_column)[value_column]
    
    # 计算均值、标准差和样本数
    results = grouped.agg(['mean', 'std', 'count']).reset_index()
    results.columns = [group_column, 'mean', 'std', 'count']
    
    print("分组后的统计结果：")
    print(results)
    
    # 对两组进行差异性分析（假设只有两组）
    if len(results) != 2:
        raise ValueError("Livak 法仅支持两组数据的分析，请检查组数。")
    
    # 提取两组数据
    group_A = data[data[group_column] == results.loc[0, group_column]][value_column]
    group_B = data[data[group_column] == results.loc[1, group_column]][value_column]
    
    # 计算均值差异 (mean_diff)
    mean_diff = results['mean'].diff().iloc[-1]
    print(f"\n两组均值的差异 (mean_diff): {mean_diff:.6f}")
    
    # 计算合并标准差 (pooled_std)
    pooled_std = np.sqrt(((results.loc[0, 'std'] ** 2 + results.loc[1, 'std'] ** 2) / 2))
    print(f"合并标准差 (pooled_std): {pooled_std:.6f}")
    
    # 计算 δT
    delta_T = mean_diff / pooled_std
    print(f"标准化均值差异 (δT): {delta_T:.6f}")
    
    # 计算倍数变化 (fold_change)
    fold_change = np.exp(delta_T)
    print(f"倍数变化 (fold_change): {fold_change:.6f}")
    
    # 计算 p 值
    t_stat, p_value = ttest_ind(group_A, group_B, equal_var=True)
    print(f"t 检验统计量 (t_stat): {t_stat:.6f}")
    print(f"p 值 (p_value): {p_value:.6f}")
    
    # 将差异性结果添加到 DataFrame
    results['mean_diff'] = [None, mean_diff]
    results['fold_change'] = [None, fold_change]
    results['delta_T'] = [None, delta_T]
    results['p_value'] = [None, p_value]
    
    # 禁用科学计数法显示
    pd.set_option('display.float_format', '{:.6f}'.format)
    
    return results

# 示例使用
file_path =r"C:\Users\huawei\OneDrive\桌面\工作簿1.xlsx"  # Excel 文件路径
sheet_name = "Sheet1"       # 工作表名称
group_column = "Group"      # 分组列名
value_column = "Value"      # 数据列名

try:
    result = livak_analysis_with_stats(file_path, sheet_name, group_column, value_column)
    print("\n最终结果：")
    print(result)
except Exception as e:
    print(f"分析失败：{e}")