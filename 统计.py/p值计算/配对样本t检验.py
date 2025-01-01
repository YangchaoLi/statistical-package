import pandas as pd
from scipy.stats import ttest_rel
from decimal import Decimal

def paired_t_test_excel(file_path, sheet_name, column1, column2):
    """
    计算 Excel 表中两列数据的配对样本 t 检验
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param column1: 第一组数据的列名
    :param column2: 第二组数据的列名
    :return: t 值和 p 值
    """
    # 读取 Excel 数据
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 检查列是否存在
    if column1 not in df.columns or column2 not in df.columns:
        raise ValueError("指定的列名不存在于表格中")
    
    # 获取两列数据并删除缺失值
    data1 = df[column1].dropna()
    data2 = df[column2].dropna()
    
    # 检查数据长度是否匹配
    if len(data1) != len(data2):
        raise ValueError("两列数据长度不一致，无法进行配对样本 t 检验")
    
    # 计算配对样本 t 检验
    t_stat, p_value = ttest_rel(data1, data2)
    return t_stat, p_value

# 示例使用
file_path = r"D:\VS code project\t_test_data.xlsx"  # 替换为你的 Excel 文件路径
sheet_name = "配对样本t检验"           # 替换为你的工作表名称
column1 = "饮用前"              # 替换为第一列变量名
column2 = "饮用后"              # 替换为第二列变量名

t_stat, p_value = paired_t_test_excel(file_path, sheet_name, column1, column2)
t_stat_decimal = Decimal(t_stat)
p_value_decimal = Decimal(p_value)
print(f"配对样本 t 检验结果: t 值 =  {t_stat_decimal:.4f}, p 值 = {p_value_decimal:.3f}")
if p_value < 0.05:
    print("结果具有统计显著性 (p < 0.05)")
else:
    print("结果不具有统计显著性 (p >= 0.05)")