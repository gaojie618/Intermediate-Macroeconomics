import pandas as pd
import scipy.stats as stats

# 定义 Excel 文件名和要分析的变量列表
file_name = "Macro_Data.xlsx"  # 请替换为实际的 Excel 文件名
variables = ["China_Real_GDP_Hundred_Million_Yuan", "China_Youth_Unemployment_Rate", "China_Fixed_Asset_Investment_Hundred_Million_Yuan", 
             "China_Population_Billion", "China_Final_Consumption_Hundred_Million_Yuan", "China_Capital_Formation_Hundred_Million_Yuan"]

# 读取 Excel 文件
try:
    df = pd.read_excel(file_name, engine="openpyxl")

    # 去除列名空格，防止匹配失败
    df.columns = df.columns.str.strip()

    # 存储统计结果
    summary_stats = []

    for variable in variables:
        if variable in df.columns:
            df[variable] = pd.to_numeric(df[variable], errors='coerce')  # 转换为数值型
            data = df[variable].dropna()

            if len(data) > 0:
                mean_val = data.mean()
                median_val = data.median()
                std_val = data.std()
                min_val = data.min()
                max_val = data.max()
                skew_val = stats.skew(data)
                kurtosis_val = stats.kurtosis(data)

                summary_stats.append([variable, mean_val, median_val, std_val, min_val, max_val, skew_val, kurtosis_val])
            else:
                print(f"Warning: {variable} has no valid data.")

        else:
            print(f"Warning: Column {variable} not found in {file_name}")

    # 生成描述性统计表格
    columns = ["Variable", "Mean", "Median", "Std Dev", "Min", "Max", "Skewness", "Kurtosis"]
    summary_df = pd.DataFrame(summary_stats, columns=columns)

    # 输出到 Excel
    output_file = "descriptive_statistics_summary.xlsx"
    summary_df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Descriptive statistics saved to {output_file}")

except Exception as e:
    print(f"Error processing {file_name}: {e}")
