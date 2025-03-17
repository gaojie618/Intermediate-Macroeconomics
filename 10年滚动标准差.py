import pandas as pd
import matplotlib.pyplot as plt  # 确保导入 matplotlib.pyplot
import numpy as np
from scipy.stats import linregress

# 读取Excel文件中的数据，假设数据在第一个sheet中
df = pd.read_excel('Macro_Data.xlsx')  

# 打印列名以确保它们是正确的
print("数据列名：", df.columns)

# 确保列名为 'Year' 和 'China_Real_GDP_Hundred_Million_Yuan'
# 如果列名不同，修改成实际的列名
df = df[['Year', 'China_Real_GDP_Hundred_Million_Yuan']]  # 选择只包含Year和GDP的列

# 计算10年滚动标准差（计算所有年份的滚动标准差）
df['Rolling_STD'] = df['China_Real_GDP_Hundred_Million_Yuan'].rolling(window=10).std()

# 删除包含NaN的行（跳过无法计算滚动标准差的部分）
df = df.dropna(subset=['Rolling_STD'])

# 显示结果
print(df[['Year', 'China_Real_GDP_Hundred_Million_Yuan', 'Rolling_STD']])

# 可选：将结果保存为新的Excel文件
df.to_excel('china_gdp_with_rolling_std_all_years.xlsx', index=False)

# 只选择1987-2023年之间的数据
df_1987_2023 = df[(df['Year'] >= 1987) & (df['Year'] <= 2023)]

# 绘制1987-2023年之间的滚动标准差图形
plt.figure(figsize=(10, 6))
plt.plot(df_1987_2023['Year'], df_1987_2023['Rolling_STD'], label='10-year Rolling Std Dev', color='b')
plt.title('10-Year Rolling Standard Deviation of GDP (1987-2023)', fontsize=14)
plt.xlabel('Year', fontsize=12)
plt.ylabel('Rolling Standard Deviation', fontsize=12)
plt.grid(True)
plt.legend()

# 确保图表设置无误
plt.show()

# 进行线性回归分析，评估经济波动是否趋于稳定
slope, intercept, r_value, p_value, std_err = linregress(df_1987_2023['Year'], df_1987_2023['Rolling_STD'])

# 输出回归分析结果
print(f'回归分析结果:')
print(f'回归斜率（Slope）：{slope:.4f}')
print(f'R-squared：{r_value**2:.4f}')
print(f'p-value：{p_value:.4f}')

# 如果p-value小于0.05，表明回归结果显著，可以判断波动是否存在趋势。
if p_value < 0.05:
    print("经济波动趋势显著，存在明显的变化趋势,经济波动没有趋于稳定")
else:
    print("经济波动趋势不显著，可能没有明显的变化趋势")
