import pandas as pd
from scipy.stats import ttest_ind

def independent_t_test(file_path, sheet_name, column1, column2):
    """
    计算独立样本 t 检验
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param column1: 第一组数据列名
    :param column2: 第二组数据列名
    :return: t 值和 p 值
    """
    # 读取 Excel 数据
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 检查列是否存在
    if column1 not in df.columns or column2 not in df.columns:
        raise ValueError("指定的列名不存在于表格中")
    
    # 获取数据并删除缺失值
    data1 = df[column1].dropna()
    data2 = df[column2].dropna()
    
    # 计算独立样本 t 检验
    t_stat, p_value = ttest_ind(data1, data2, equal_var=False)  # Welch's t-test
    return t_stat, p_value

# 示例使用
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为你的 Excel 文件路径
sheet_name = "Sheet1"           # 替换为你的工作表名称
column1 = "Group1"              # 替换为第一组变量名
column2 = "Group2"              # 替换为第二组变量名

t_stat, p_value = independent_t_test(file_path, sheet_name, column1, column2)
print(f"独立样本 t 检验结果: t 值 = {t_stat:.3f}, p 值 = {p_value:.3f}")
if p_value < 0.05:
    print("结果具有统计显著性 (p < 0.05)")
else:
    print("结果不具有统计显著性 (p >= 0.05)")