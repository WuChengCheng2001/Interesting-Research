import pandas as pd
import numpy as np

# ==============================
# Step 1: 加载语言-国家映射文件
# ==============================
print("正在加载 language_country.csv...")
lang_country_df = pd.read_csv('language_country.csv', header=None)
# 根据你提供的内容，列顺序为：Language, Country_EN, Country_ZH, Country_Code
lang_country_df.columns = ['Language', 'Country_EN', 'Country_ZH', 'Country_Code']

# 获取十大语言列表（按你指定的顺序）
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
# Step 2: 加载 World Bank 人口数据
# ==============================
print("\n正在加载 API_SP.POP.TOTL_DS2_en_excel_v2_20.xls...")
# World Bank 数据通常有多个 sheet，我们读取第一个（或包含数据的 sheet）
# 注意：该文件可能包含元数据行，需跳过前几行
pop_df = pd.read_excel(
    'API_SP.POP.TOTL_DS2_en_excel_v2_20.xls',
    skiprows=3,  # 跳过前3行说明文字（World Bank 标准格式）
    na_values='..'  # 将 '..' 视为 NaN
)

# 检查列名：World Bank 文件通常包含 ['Country Name', 'Country Code', 'Indicator Name', 'Indicator Code', '1960', '1961', ..., '2024']
required_cols = ['Country Name', 'Country Code']
if not all(col in pop_df.columns for col in required_cols):
    raise ValueError("人口数据文件缺少必要列：'Country Name' 或 'Country Code'")

# 提取年份列（1960 到 2024）
year_columns = [str(year) for year in range(1960, 2025)]
missing_years = [y for y in year_columns if y not in pop_df.columns]
if missing_years:
    print(f"警告：缺失年份列：{missing_years}")

# 只保留需要的列：Country Code + 年份
pop_clean = pop_df[['Country Code'] + year_columns].copy()

# 将 Country Code 设为索引，方便后续查询
pop_clean.set_index('Country Code', inplace=True)

# ==============================
# Step 3: 汇总每种语言的年度人口
# ==============================
print("\n正在汇总各语言年度人口...")

# 初始化结果 DataFrame：行是年份，列是语言
result_df = pd.DataFrame(index=year_columns, columns=target_languages, dtype=float)

# 对每种语言进行聚合
for lang in target_languages:
    country_codes = lang_to_countries[lang]
    # 过滤出在人口数据中存在的国家
    valid_codes = [code for code in country_codes if code in pop_clean.index]
    if not valid_codes:
        print(f"警告：语言 '{lang}' 无匹配的国家人口数据！")
        result_df[lang] = np.nan
        continue
    
    # 对这些国家的人口按年求和
    lang_pop_sum = pop_clean.loc[valid_codes, year_columns].sum(axis=0)
    result_df[lang] = lang_pop_sum.values

# 转换索引为整数年份
result_df.index = result_df.index.astype(int)

# ==============================
# Step 4: 导出结果到 Excel
# ==============================
output_file = 'Global_Language_Population_1960_2024.xlsx'
result_df.to_excel(output_file, sheet_name='Language Population')
print(f"\n✅ 成功导出结果至: {output_file}")

# 可选：打印前5行预览
print("\n结果预览（前5年）:")
print(result_df.head())