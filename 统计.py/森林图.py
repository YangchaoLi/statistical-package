import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import matplotlib
matplotlib.rcParams['font.family'] = 'SimHei' 
matplotlib.rcParams['axes.unicode_minus'] = False
def calculate_or_and_generate_forest_plot(input_path, sheet_name, start_row, end_row, start_col, end_col, output_path, plot_path):
    """
    从Excel表格中指定区域提取数据，计算每一行的OR值并绘制森林图，支持处理某些数值为0的情况。
    
    参数说明：
    - input_path: 输入的Excel文件路径
    - sheet_name: 指定工作表名称
    - start_row, end_row: 数据的起始行和结束行（从1开始计数）
    - start_col, end_col: 数据的起始列和结束列（从1开始计数）
    - output_path: 编辑后的Excel文件保存路径
    - plot_path: 森林图保存路径
    """
    # 检查文件路径是否存在
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"输入的Excel文件不存在: {input_path}")
    
    # 读取Excel文件
    df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    # 提取指定区域的数据
    try:
        selected_data = df.iloc[start_row-1:end_row, start_col-1:end_col].copy()
    except IndexError:
        raise ValueError("指定的行列范围超出Excel文件的数据范围，请检查输入参数。")

    # 设置列名（默认假设为四格表数据）
    selected_data.columns = ["a", "b", "c", "d"]

    # 处理数值为0的情况（加0.5校正法）
    selected_data[["a", "b", "c", "d"]] = selected_data[["a", "b", "c", "d"]].replace(0, 0.0000000000001)

    # 计算OR值和95%置信区间
    selected_data["OR"] = (selected_data["a"] * selected_data["d"]) / (selected_data["b"] * selected_data["c"])
    selected_data["Lower_CI"] = np.exp(np.log(selected_data["OR"]) - 1.96 * 
                                       np.sqrt(1/selected_data["a"] + 1/selected_data["b"] + 
                                               1/selected_data["c"] + 1/selected_data["d"]))
    selected_data["Upper_CI"] = np.exp(np.log(selected_data["OR"]) + 1.96 * 
                                       np.sqrt(1/selected_data["a"] + 1/selected_data["b"] + 
                                               1/selected_data["c"] + 1/selected_data["d"]))

    # 合并原始数据和计算结果
    result = selected_data.copy()

    # 保存编辑后的Excel文件
    try:
        result.to_excel(output_path, index=False)
        print(f"计算结果已保存到: {output_path}")
    except Exception as e:
        raise IOError(f"保存编辑后的Excel文件失败: {e}")

    # 绘制森林图
    create_forest_plot(result, plot_path)

def create_forest_plot(df, plot_path):
    """
    根据输入数据绘制森林图并保存到指定路径。
    
    参数说明：
    - df: 包含OR值和置信区间的数据框
    - plot_path: 森林图保存路径
    """
    studies = df.index + 1  # 使用行号作为研究名称
    or_values = df["OR"]
    lower_cis = df["Lower_CI"]
    upper_cis = df["Upper_CI"]

    # 处理置信区间中的无穷值
    finite_upper = upper_cis[np.isfinite(upper_cis)]
    if len(finite_upper) > 0:
        upper_cis = np.where(
            np.isinf(upper_cis),
            np.nanmax(finite_upper) * 1.5,
            upper_cis
        )

    # 计算误差条
    xerr = [or_values - lower_cis, upper_cis - or_values]

    # 绘制森林图
    plt.figure(figsize=(10,6))
    plt.errorbar(
        or_values,
        range(len(or_values)),
        xerr=xerr,
        fmt='o',
        capsize=3,
        color='blue',
        label='OR with 95% CI'
    )
    plt.axvline(x=1, color='red', linestyle='--', label='参考线 (OR=1)')
    plt.yticks(range(len(or_values)), [f"Study {i}" for i in studies])
    plt.xlabel("Odds Ratio (OR)")
    plt.title("森林图", fontproperties="SimHei")
    plt.gca().invert_yaxis()  # 反转y轴便于阅读
    plt.legend()

    # 自动调整x轴范围
    
    max_or = max(upper_cis)
    plt.xlim(-30, 30 )
    
    # 保存森林图
    try:
        plt.tight_layout()
        plt.savefig(plot_path)
        plt.close()
        print(f"森林图已保存到: {plot_path}")
    except Exception as e:
        raise
calculate_or_and_generate_forest_plot(
    input_path=r"D:\VS code project\t_test_data.xlsx",       # 输入Excel文件路径
    sheet_name="森林图",           # 工作表名称
    start_row=2,                   # 数据起始行
    end_row=12,                    # 数据结束行
    start_col=2,                   # 数据起始列
    end_col=5,                     # 数据结束列
    output_path="D:\VS code project\output.xlsx",     # 计算结果保存路径
    plot_path=r"D:\VS code project\forest_plot.png"    # 森林图保存路径
)