import pandas as pd

# 读取 Excel 文件
excel_file = 'kkk/流水.xlsx'  # 替换为你的 Excel 文件路径
xls = pd.ExcelFile(excel_file)

# 获取所有 sheet 名称
sheet_names = xls.sheet_names

# 初始化变量
total_rows = 0  # 总行数
filtered_rows = 0  # 筛选后包含 '李云' 的行数
all_columns = []  # 存储所有列名

# 遍历所有 sheet
for sheet in sheet_names:
    # 读取当前 sheet 的数据
    df = pd.read_excel(excel_file, sheet_name=sheet)

    # 计算总行数
    total_rows += len(df)

    # 获取列名
    columns = df.columns.tolist()  # 获取列名并转换为列表
    all_columns.extend(columns)  # 将所有列名加入列表

    # 打印每个 sheet 的列名
    print(f"Sheet {sheet} 列名: {columns}")

    # 确保所有列的内容都被转换为字符串类型，以避免出现非字符串列导致的错误
    df = df.applymap(str)

    # 筛选包含 '李云' 的行，假设备注列在最后一列
    filtered_df = df[df.iloc[:, -1].str.contains('李云', na=False)]  # 这里假设最后一列是备注列

    # 计算筛选后的行数
    filtered_rows += len(filtered_df)

# 去重所有列名并计算总列数
all_columns = list(set(all_columns))  # 去重列名
num_columns = len(all_columns)  # 计算列数

# 输出结果
print(f"总行数: {total_rows}")
print(f"筛选出包含 '李云' 的行数: {filtered_rows}")
print(f"文件中一共存在 {num_columns} 个列名：")
print(all_columns)
