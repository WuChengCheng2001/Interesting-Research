import pandas as pd
import numpy as np

# ==============================
# Step 1: 加载语言-国家映射文件
# ==============================
print("正在加载 language_country.csv...")
lang_country_df = pd.read_csv('language_country.csv', header=None)
# 根据你多次上传的内容，列顺序为：Language, Country_EN, Country_ZH, Country_Code
lang_country_df.columns = ['Language', 'Country_EN', 'Country_ZH', 'Country_Code']

# 定义十大语言（按你指定的顺序）
target_languages = [
    "English", "Mandarin", "Spanish", "Hindi", "Arabic",
    "French", "Bengali", "Portuguese", "Russian", "Indonesian"
]

# 构建语言 -> 国家代码集合的字典
lang_to_countries = {}
for lang in target_languages:
    codes = lang_country_df[lang_country_df['Language'] == lang]['Country_Code'].unique()
    lang_to_countries[lang] = set(codes)

print("十大语言及其国家数量：")
for lang, codes in lang_to_countries.items():
    print(f"  {lang}: {len(codes)} 个国家")

# ==============================
# Step 2: 加载 World Bank GDP 数据
# ==============================
print("\n正在加载 API_NY.GDP.MKTP.CD_DS2_en_excel_v2_129.xls...")
# World Bank 的 GDP 文件通常前3行为元数据，第4行开始是列名
gdp_df = pd.read_excel(
    'API_NY.GDP.MKTP.CD_DS2_en_excel_v2_129.xls',
    skiprows=3,  # 跳过前3行说明
    na_values=['..', '']  # 将缺失值标记为 NaN
)

# 检查必要列是否存在
required_cols = ['Country Name', 'Country Code']
if not all(col in gdp_df.columns for col in required_cols):
    raise ValueError("GDP数据文件缺少 'Country Name' 或 'Country Code' 列！")

# 提取年份列（1960 到 2024）
year_columns = [str(year) for year in range(1960, 2025)]
available_years = [col for col in year_columns if col in gdp_df.columns]
missing_years = [y for y in year_columns if y not in gdp_df.columns]
if missing_years:
    print(f"⚠️ 警告：GDP文件中缺失年份：{missing_years}")

# 只保留 Country Code 和年份列
gdp_clean = gdp_df[['Country Code'] + available_years].copy()

# 设置 Country Code 为索引，便于查询
gdp_clean.set_index('Country Code', inplace=True)

# ==============================
# Step 3: 汇总每种语言的年度GDP
# ==============================
print("\n正在汇总各语言年度GDP...")

# 初始化结果 DataFrame：行是年份，列是语言
result_df = pd.DataFrame(index=available_years, columns=target_languages, dtype=float)

# 对每种语言进行聚合
for lang in target_languages:
    country_codes = lang_to_countries[lang]
    # 过滤出在GDP数据中存在的国家
    valid_codes = [code for code in country_codes if code in gdp_clean.index]
    if not valid_codes:
        print(f"⚠️ 警告：语言 '{lang}' 无匹配的国家GDP数据！")
        result_df[lang] = np.nan
        continue
    
    # 对这些国家的GDP按年求和（单位：当前美元）
    lang_gdp_sum = gdp_clean.loc[valid_codes, available_years].sum(axis=0)
    result_df[lang] = lang_gdp_sum.values

# 转换索引为整数年份（方便后续使用）
result_df.index = result_df.index.astype(int)

# ==============================
# Step 4: 导出结果到 Excel
# ==============================
output_file = 'Global_Language_GDP_1960_2024.xlsx'
result_df.to_excel(output_file, sheet_name='Language GDP (Current USD)')
print(f"\n✅ 成功导出结果至: {output_file}")

# 可选：打印前5行预览
print("\n结果预览（前5年）:")
print(result_df.head())