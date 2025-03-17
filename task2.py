import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# 读取数据
data = pd.read_excel('Macro_Data.xlsx')

# 计算实际GDP增长率
data['US_GDP_Growth_Rate'] = data['US_Real_GDP_Billion_Yuan'].pct_change() * 100
data['China_GDP_Growth_Rate'] = data['China_Real_GDP_Billion_Yuan'].pct_change() * 100

# 1. 双轴折线图：比较中国和美国的实际GDP增长率
fig1 = go.Figure()

fig1.add_trace(go.Scatter(x=data['Year'], y=data['US_GDP_Growth_Rate'], name='美国GDP增长率', line=dict(color='blue')))
fig1.add_trace(go.Scatter(x=data['Year'], y=data['China_GDP_Growth_Rate'], name='中国GDP增长率', line=dict(color='red'), yaxis='y2'))

# 添加关键事件标注
fig1.add_vline(x=2008, line_dash="dash", line_color="green", annotation_text="2008金融危机", annotation_position="top left")
fig1.add_vline(x=2020, line_dash="dash", line_color="purple", annotation_text="COVID-19大流行", annotation_position="top left")

fig1.update_layout(
    title='中国和美国的实际GDP增长率比较 (1978-2023)',
    xaxis_title='年份',
    yaxis_title='美国GDP增长率 (%)',
    yaxis2=dict(title='中国GDP增长率 (%)', overlaying='y', side='right'),
    legend=dict(x=0.02, y=0.98),
    hovermode="x unified",
    annotations=[dict(
        x=1, y=-0.05,
        xref='paper', yref='paper',
        text='数据来源：同花顺iFinD',
        showarrow=False,
        font=dict(size=12, color="black"),
        xanchor='right', yanchor='auto'
    )]
)

# 2. 堆叠面积图：中国的消费率和资本形成率
data['Consumption_Rate'] = data['China_Final_Consumption_Billion_Yuan'] / data['China_Real_GDP_Billion_Yuan'] * 100
data['Capital_Formation_Rate'] = data['China_Capital_Formation_Billion_Yuan'] / data['China_Real_GDP_Billion_Yuan'] * 100

fig2 = go.Figure()

fig2.add_trace(go.Scatter(x=data['Year'], y=data['Consumption_Rate'], fill='tozeroy', name='消费率', line=dict(color='blue')))
fig2.add_trace(go.Scatter(x=data['Year'], y=data['Capital_Formation_Rate'], fill='tonexty', name='资本形成率', line=dict(color='orange')))

# 标注投资驱动增长占主导的阶段
start_year = 2000
end_year = 2023
fig2.add_vrect(
    x0=start_year,
    x1=end_year,
    fillcolor="green",
    opacity=0.1,
    line_width=0,
    annotation=dict(
        text="投资驱动增长占主导 (2000年后基础设施繁荣)",
        x=(start_year + end_year) / 2,
        y=0.1,  # 进一步调整垂直位置
        xref='x',
        yref='paper',
        showarrow=False,
        font=dict(size=12, color="red"),  # 更改字体颜色为红色，使其更明显
        align="center"
    )
)

fig2.update_layout(
    title='中国的消费率和资本形成率 (1978-2023)',
    xaxis_title='年份',
    yaxis_title='百分比 (%)',
    legend=dict(x=0.02, y=0.95),  # 调整图例位置，避免遮挡标注
    hovermode="x unified",
    annotations=[dict(
        x=1, y=-0.05,
        xref='paper', yref='paper',
        text='数据来源：同花顺iFinD',
        showarrow=False,
        font=dict(size=12, color="black"),
        xanchor='right', yanchor='auto'
    )]
)

# 3. 散点图：青年失业率与GDP增长的关系 (2020-2023)
recent_data = data[data['Year'] >= 2020]

fig3 = px.scatter(recent_data, x='China_Real_GDP_Billion_Yuan', y='China_Youth_Unemployment_Rate', trendline="ols",
                  title='青年失业率与GDP增长的关系 (2020-2023)',
                  labels={'China_Real_GDP_Billion_Yuan': '中国实际GDP (亿人民币)', 'China_Youth_Unemployment_Rate': '青年失业率 (%)'})

fig3.update_layout(
    legend=dict(x=0.02, y=0.98),
    hovermode="x unified",
    annotations=[dict(
        x=1, y=-0.05,
        xref='paper', yref='paper',
        text='数据来源：同花顺iFinD',
        showarrow=False,
        font=dict(size=12, color="black"),
        xanchor='right', yanchor='auto'
    )]
)

# 将图表保存为HTML文件
fig1.write_html("gdp_growth_comparison.html")
fig2.write_html("consumption_capital_formation.html")
fig3.write_html("youth_unemployment_gdp.html")

print("图表已保存为HTML文件，请手动打开以下文件查看：")
print("1. gdp_growth_comparison.html")
print("2. consumption_capital_formation.html")
print("3. youth_unemployment_gdp.html")