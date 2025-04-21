import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory



def merge_excel_files(folder_path, key_column, output_file):
    """
    合并文件夹内多个 Excel 文件，按照指定的列进行合并，保存到新的 Excel 文件。

    :param folder_path: 文件夹路径，包含所有待合并的 Excel 文件
    :param key_column: 用于合并的列名
    :param output_file: 输出合并结果的 Excel 文件路径
    """
    # 用于存储所有 Excel 数据的列表
    dataframes = []
    
    # 遍历文件夹内的 Excel 文件
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):  # 只处理 .xlsx 文件
            file_path = os.path.join(folder_path, file_name)
            print(f"正在处理文件: {file_path}")
            # 读取 Excel 文件
            df = pd.read_excel(file_path)
            # 确保关键列存在
            if key_column not in df.columns:
                raise ValueError(f"文件 {file_name} 中缺少关键列 '{key_column}'")
            dataframes.append(df)
    
    # 按关键列合并所有数据
    print("正在合并数据...")
    merged_data = dataframes[0]
    for df in dataframes[1:]:
        merged_data = pd.merge(merged_data, df, on=key_column, how='outer')  # 使用 outer 保留所有数据

    # 保存合并后的数据到新的 Excel 文件
    print(f"正在将合并结果保存到: {output_file}")
    merged_data.to_excel(output_file, index=False)
    print("合并完成！")

# 使用示例
folder_path = 'C:\\Users\\MrGreed\\Desktop\\新建文件夹'  # 替换为包含 Excel 文件的文件夹路径
key_column = 'SN'  # 替换为合并所依据的列名
output_file = folder_path + "\\" + "合并表格.xlsx"  # 合并结果保存的文件名



def select_folder_and_list_files():
    """
    打开文件夹选择对话框，让用户选择一个文件夹，并列出该文件夹中的所有文件。
    """
    # 创建 Tkinter 根窗口并隐藏它
    Tk().withdraw()

    # 弹出文件夹选择对话框
    folder_path = askdirectory(title="选择文件夹")
    
    # 如果用户未选择文件夹，则退出程序
    if not folder_path:
        print("未选择任何文件夹，程序退出。")
        return

    # 列出文件夹中的所有文件
    print(f"选择的文件夹路径: {folder_path}")
    print("文件列表:")
    for file_name in os.listdir(folder_path):
        print(f"- {file_name}")



# 调用函数
select_folder_and_list_files()
merge_excel_files(folder_path, key_column, output_file)