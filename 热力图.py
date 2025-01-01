#pip install pandas seaborn matplotlib openpyxl安装pandas、seaborn、matplotlib、openpyxl
import pandas as pd#数据处理和分析
import seaborn as sns#也是matplotlib的模块，绘图
import matplotlib.pyplot as plt#创建绘图的框架

# 读取 Excel 文件
file_path = "D:\heatmap_data.xlsx"  # 替换为实际文件路径
data = pd.read_excel(file_path, index_col=0)

# 设置绘图样式
plt.figure(figsize=(8, 6))#每个格子的宽，高
sns.heatmap(data, annot=True, fmt="d", cmap="RdYlBu", linewidths=0.5)

# 添加标题
plt.title("Heatmap of Metrics", fontsize=16)

# 显示图形
plt.show()