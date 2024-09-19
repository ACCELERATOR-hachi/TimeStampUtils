import pandas as pd

# 指定 Excel 文件路径
excel_path = './招聘表.xlsx'
backup_excel_path = './备案表.xlsx'

# 读取 Excel 文件中的"状态跟踪表"sheet
try:
    data = pd.read_excel(excel_path, sheet_name='状态跟踪表')

    # 确保状态跟踪表中有"姓名"、"部门"和"状态明细"列
    required_columns = ['姓名', '部门', '状态明细']
    for col in required_columns:
        if col not in data.columns:
            raise ValueError(f"状态跟踪表中缺少'{col}'列")

    # 读取现有的备案表数据
    try:
        backup_data = pd.read_excel(backup_excel_path)
    except FileNotFoundError:
        # 如果备案表不存在，则创建新的备案表
        print("备案表不存在，将创建新的备案表。")
        backup_data = pd.DataFrame(columns=['部门', '姓名', '状态明细'])
    except Exception as e:
        print(f"读取备案表时发生错误：{e}")
        backup_data = pd.DataFrame(columns=['部门', '姓名', '状态明细'])

    # 确保备案表中有"部门"、"姓名"和"状态明细"列
    for col in ['部门', '姓名', '状态明细']:
        if col not in backup_data.columns:
            raise ValueError(f"备案表中缺少'{col}'列")

    # 遍历状态跟踪表中的每一行
    for index, row in data.iterrows():
        # 检查备案表中是否存在该姓名
        if row["姓名"] in backup_data["姓名"].values:
            # 获取备案表中对应姓名的行索引
            backup_index = backup_data[backup_data["姓名"] == row["姓名"]].index[0]
            # 更新备案表中的"部门"和"状态明细"列
            backup_data.at[backup_index, '部门'] = row['部门']
            backup_data.at[backup_index, '状态明细'] = row['状态明细']
        else:
            # 如果备案表中不存在该姓名，则添加新记录
            new_row = pd.DataFrame([{
                '部门': row['部门'],
                '姓名': row['姓名'],
                '状态明细': row['状态明细']
            }], index=[0])  # 提供索引
            backup_data = pd.concat([backup_data, new_row], ignore_index=True)

    # 确保列顺序为"部门"、"姓名"、"状态明细"
    backup_data = backup_data[['部门', '姓名', '状态明细']+ [col for col in backup_data.columns if col not in ['部门', '姓名', '状态明细']]]

    # 将更新后的备案表数据写入文件
    try:
        backup_data.to_excel(backup_excel_path, index=False)
        print(f"备案表已更新，新的信息已写入到文件：{backup_excel_path}")
    except Exception as e:
        print(f"写入备案表时发生错误：{e}")

except FileNotFoundError:
    print("文件未找到，请检查文件路径是否正确。")
except ValueError as ve:
    print(f"列名错误：{ve}")
except Exception as e:
    print(f"发生错误：{e}")


def update_status(data):
    """根据状态明细的值更新相应行的值"""
    for index, row in data.iterrows():
        status = row['状态明细']

        # 根据状态明细值更新后面的单元格
        if status == '开始':
            data.at[index, '状态1'] = 1  # 假设要更新的列名是'对应列名'
        elif status == '结束':
            data.at[index, '状态2'] = 1
        elif status == '进行':
            data.at[index, '状态3'] = 1

# 使用示例
try:
    data = pd.read_excel(backup_excel_path)
    update_status(data)
    # 将更新后的数据写回文件
    data.to_excel(backup_excel_path, index=False)
    print("状态已更新。")
except FileNotFoundError:
    print("文件未找到，请检查文件路径是否正确。")
except Exception as e:
    print(f"发生错误：{e}")