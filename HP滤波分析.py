import pandas as pd
import matplotlib.pyplot as plt
from statsmodels.tsa.filters.hp_filter import hpfilter
import numpy as np

# 读取数据并处理
df = pd.read_excel('Macro_Data.xlsx')
gdp_data = df[['Year', 'US_Real_GDP_Hundred_Million_Yuan', 
               'China_Real_GDP_Hundred_Million_Yuan']].copy()
gdp_data.columns = ['Year', 'US_GDP', 'CN_GDP']
gdp_data.set_index('Year', inplace=True)

# 定义HP滤波函数（lambda=100用于年度数据）
def hp_decomposition(series, lamb=100):
    cycle, trend = hpfilter(series, lamb=lamb)
    return pd.DataFrame({'Trend': trend, 'Cycle': cycle})

# 分解并标准化周期成分（百分比偏离趋势）
us_decomp = hp_decomposition(gdp_data['US_GDP'])
cn_decomp = hp_decomposition(gdp_data['CN_GDP'])

# 转换为百分比波动
us_decomp['Cycle_pct'] = (us_decomp['Cycle'] / us_decomp['Trend']) * 100
cn_decomp['Cycle_pct'] = (cn_decomp['Cycle'] / cn_decomp['Trend']) * 100

# 设置专业绘图样式（使用Matplotlib内置样式）
plt.style.use('ggplot')  # 替换为'classic'或'default'均可
plt.rcParams.update({
    'font.sans-serif': 'SimHei',  # 支持中文显示
    'axes.unicode_minus': False,
    'figure.dpi': 150,
    'axes.titlesize': 10,
    'axes.labelsize': 9,
    'legend.fontsize': 7,
    'xtick.labelsize': 8,
    'ytick.labelsize': 8
})

# 绘制趋势分解图
fig, axes = plt.subplots(2, 2, figsize=(12, 10), 
                         gridspec_kw={'hspace': 0.5, 'wspace': 0.3})  # 增大 hspace

# 美国趋势组件
axes[0,0].plot(us_decomp['Trend'], color='#1f77b4', lw=1.5, label='趋势成分')
axes[0,0].plot(gdp_data['US_GDP'], '--', color='#aec7e8', lw=1, label='实际GDP')
axes[0,0].set_title('(a) 美国GDP趋势成分', pad=10)
axes[0,0].set_xlabel('年份')
axes[0,0].set_ylabel('GDP（亿元人民币）')
axes[0,0].legend(loc='upper left')
axes[0,0].grid(True, alpha=0.3)

# 添加横坐标年份标注
axes[0,0].set_xticks(range(1980, 2021, 10))  # 设置横坐标刻度
axes[0,0].set_xticklabels([str(year) for year in range(1980, 2021, 10)])  # 设置横坐标标签

# 美国周期组件（百分比形式）
axes[0,1].plot(us_decomp['Cycle_pct'], color='#d62728', lw=1.2)
axes[0,1].fill_between(us_decomp.index, us_decomp['Cycle_pct'], color='#ff9896', alpha=0.3)
axes[0,1].set_title('(b) 美国GDP周期波动', pad=10)
axes[0,1].set_xlabel('年份')
axes[0,1].set_ylabel('偏离趋势百分比 (%)')
axes[0,1].axhline(0, color='k', linestyle=':', lw=0.8)
axes[0,1].set_ylim(-20, 20)

# 添加横坐标年份标注
axes[0,1].set_xticks(range(1980, 2021, 10))
axes[0,1].set_xticklabels([str(year) for year in range(1980, 2021, 10)])

# 中国趋势组件
axes[1,0].plot(cn_decomp['Trend'], color='#2ca02c', lw=1.5, label='趋势成分')
axes[1,0].plot(gdp_data['CN_GDP'], '--', color='#98df8a', lw=1, label='实际GDP')
axes[1,0].set_title('(c) 中国GDP趋势成分', pad=10)
axes[1,0].set_xlabel('年份')
axes[1,0].set_ylabel('GDP（亿元人民币）')
axes[1,0].legend(loc='upper left')
axes[1,0].grid(True, alpha=0.3)
axes[1,0].ticklabel_format(style='scientific', axis='y', scilimits=(0, 0))  # 启用科学计数法

# 添加横坐标年份标注
axes[1,0].set_xticks(range(1980, 2021, 10))
axes[1,0].set_xticklabels([str(year) for year in range(1980, 2021, 10)])

# 中国周期组件（百分比形式）
axes[1,1].plot(cn_decomp['Cycle_pct'], color='#9467bd', lw=1.2)
axes[1,1].fill_between(cn_decomp.index, cn_decomp['Cycle_pct'], color='#c5b0d5', alpha=0.3)
axes[1,1].set_title('(d) 中国GDP周期波动', pad=10)
axes[1,1].set_xlabel('年份')
axes[1,1].set_ylabel('偏离趋势百分比 (%)')
axes[1,1].axhline(0, color='k', linestyle=':', lw=0.8)
axes[1,1].set_ylim(-20, 20)

# 添加横坐标年份标注
axes[1,1].set_xticks(range(1980, 2021, 10))
axes[1,1].set_xticklabels([str(year) for year in range(1980, 2021, 10)])

plt.tight_layout()
plt.show()

# 周期同步性分析图
fig, ax = plt.subplots(figsize=(10, 5))

# 标准化周期波动
# 为了避免实际数值差距过大而影响比较，我们对周期波动进行了标准化处理
us_norm = (us_decomp['Cycle_pct'] - us_decomp['Cycle_pct'].mean()) / us_decomp['Cycle_pct'].std()
cn_norm = (cn_decomp['Cycle_pct'] - cn_decomp['Cycle_pct'].mean()) / cn_decomp['Cycle_pct'].std()

ax.plot(us_norm, label='美国周期波动', color='#1f77b4', lw=1.2)
ax.plot(cn_norm, label='中国周期波动', color='#ff7f0e', lw=1.2)
ax.set_title('中美经济周期同步性分析（1978-2023）', pad=10)
ax.set_xlabel('年份')
ax.set_ylabel('标准化波动值')
ax.grid(True, alpha=0.3)

# 添加关键事件标记
events = {
    1997: 'Asian Financial Crisis',
    2008: 'Global Financial Crisis',
    2015: 'China Stock Market Crash',
    2020: 'COVID-19 Pandemic'
}
for year, label in events.items():
    ax.axvline(x=year, color='grey', linestyle='--', alpha=0.5)
    ax.text(year, 2.5, label, rotation=90, va='top', ha='center', fontsize=8)

# 计算滚动相关系数
rolling_corr = us_norm.rolling(window=5, min_periods=3).corr(cn_norm)
ax2 = ax.twinx()
ax2.plot(rolling_corr, color='#2ca02c', lw=1.5, alpha=0.7, label='5年滚动相关系数')
ax2.set_ylabel('相关系数', color='#2ca02c')
ax2.set_ylim(-1, 1)
ax2.tick_params(axis='y', colors='#2ca02c')

# 合并图例
lines, labels = ax.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax.legend(lines + lines2, labels + labels2, loc='upper left')

plt.tight_layout()
plt.show()

# 计算总体相关系数
# 对齐数据
us_norm, cn_norm = us_norm.align(cn_norm, join='inner')

# 填充缺失值（使用推荐的方法）
us_norm = us_norm.ffill().bfill()  # 前向填充 + 后向填充
cn_norm = cn_norm.ffill().bfill()  # 前向填充 + 后向填充

# 确保数据是数值类型
us_norm = pd.to_numeric(us_norm, errors='coerce')
cn_norm = pd.to_numeric(cn_norm, errors='coerce')

# 计算相关系数
corr_coef = us_norm.corr(cn_norm)
print(f"总体周期成分相关系数: {corr_coef:.3f}")

# 手动计算相关系数以验证
corr_matrix = np.corrcoef(us_norm, cn_norm)
print(f"手动计算的相关系数: {corr_matrix[0, 1]:.3f}")

# 绘制散点图以直观检查相关性
plt.scatter(us_norm, cn_norm)
plt.title('Scatter Plot of US vs China Business Cycle')
plt.xlabel('US Cycle')
plt.ylabel('China Cycle')
plt.grid(True)
plt.show()

# 阶段相关系数计算（保持原逻辑）
periods = {'1978-1990': (1978, 1990), '1991-2001': (1991, 2001),
           '2002-2008': (2002, 2008), '2009-2019': (2009, 2019),
           '2020-2023': (2020, 2023)}

corr_results = []
for name, (start, end) in periods.items():
    mask = (gdp_data.index >= start) & (gdp_data.index <= end)
    corr = us_norm[mask].corr(cn_norm[mask])
    corr_results.append([name, corr])

corr_df = pd.DataFrame(corr_results, columns=['时间段', '相关系数'])
print("分阶段相关系数：")
print(corr_df)