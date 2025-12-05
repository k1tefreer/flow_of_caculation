import pandas as pd
import os

# 自动化脚本：计算指定目录下所有xlsx文件中的总金额
def calculate_total_amount(path, direction, keyword):
    total_amount = 0

    if os.path.isfile(path):  # 如果是单个文件
        files = [path]
    elif os.path.isdir(path):  # 如果是目录
        files = [os.path.join(path, f) for f in os.listdir(path) if f.endswith('.xlsx')]
    else:
        print("路径无效，请检查")
        return

    for file_path in files:
        print(f"正在读取：{file_path}")
        try:
            df = pd.read_excel(file_path)
            print("列名：", df.columns)

            # 修改列名，确保完全匹配
            if '借贷类型' in df.columns and '交易用途类型' in df.columns and '交易金额(分)' in df.columns and '对手侧账户名称' in df.columns:
                if direction == '提现':
                    filtered_df = df[(df['交易用途类型'].str.contains('提现', na=False)) & (df['对手侧账户名称'].str.contains(keyword, na=False))]
                else:
                    filtered_df = df[(df['借贷类型'] == direction) & (df['对手侧账户名称'].str.contains(keyword, na=False))]

                total_amount += filtered_df['交易金额(分)'].sum()
                print(f"{file_path}: {filtered_df['交易金额(分)'].sum()} 已计算")
            else:
                print(f"{file_path}: 缺少必要列，跳过")

        except Exception as e:
            print(f"{file_path} 加载失败: {str(e)}")

    print(f"\n总金额（分）：{total_amount}")

# 修改为直接在代码中输入
if __name__ == '__main__':
    path = 'tc/tc.xlsx'  # 修改为你的绝对路径
    direction = '出'  # 可改为 '入'、'出' 或 '提现'
    keyword = '黄一十'  # 修改为你的关键词
    calculate_total_amount(path, direction, keyword)