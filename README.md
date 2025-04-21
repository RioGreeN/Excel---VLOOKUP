# Excel---VLOOKUP
将Excel的VLOOKUP功能通过Panda实现


程序说明
参数说明：

folder_path：包含待合并 Excel 文件的文件夹路径。
key_column：用于合并的列名，假设该列在所有 Excel 文件中都有且数据存在一一对应关系。
output_file：合并结果保存的文件路径，例如 merged_data.xlsx。
逻辑流程：

遍历指定文件夹内的所有 .xlsx 文件。
读取每个文件的数据，并确保 key_column 存在于数据中。
使用 Pandas 的 merge 方法按 key_column 对所有数据框进行合并，how='outer' 确保不会丢失任何数据。
结果文件：

合并后的数据保存为新的 Excel 文件 merged_data.xlsx，其中包含所有原始数据。
注意事项：

确保所有 Excel 文件中用于合并的列名一致。
安装依赖库：
bash
pip install pandas openpyxl
示例数据： 假设文件夹内有以下两个 Excel 文件：

file1.xlsx：

Key	ColumnA	ColumnB
1	A1	B1
2	A2	B2
file2.xlsx：

Key	ColumnC	ColumnD
2	C2	D2
3	C3	D3
合并结果 merged_data.xlsx：

Key	ColumnA	ColumnB	ColumnC	ColumnD
1	A1	B1	NaN	NaN
2	A2	B2	C2	D2
3	NaN	NaN	C3	D3
