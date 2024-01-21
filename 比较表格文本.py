import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def compare_excel_files(file1_path, file2_path, output_path, text_column):
    try:
        # 读取两个Excel文件
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        # 使用merge函数将两个数据框合并，根据用户选择的文本列进行合并
        merged_df = pd.merge(df1, df2, how='inner', on=text_column)

        # 确保输出目录存在，如果不存在则创建
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 将结果保存到新的Excel文件
        merged_df.to_excel(output_path, index=False)
        print(f"相同文本已提取并保存到 {output_path} 中。")
    except Exception as e:
        print(f"出现错误: {e}")

def select_file(entry_widget):
    file_path = filedialog.askopenfilename()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def on_save_as_button_click(entry_widget):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def on_compare_button_click():
    file1_path = entry_file1.get()
    file2_path = entry_file2.get()
    output_path = entry_output.get()
    text_column = entry_text_column.get()  # 获取用户输入的文本列名

    compare_excel_files(file1_path, file2_path, output_path, text_column)

# 创建主窗口
root = tk.Tk()
root.title("Excel文件比较工具")

# 创建文件选择和输入框
label_file1 = tk.Label(root, text="选择文件1:")
entry_file1 = tk.Entry(root, width=50)
button_browse1 = tk.Button(root, text="浏览", command=lambda: select_file(entry_file1))

label_file2 = tk.Label(root, text="选择文件2:")
entry_file2 = tk.Entry(root, width=50)
button_browse2 = tk.Button(root, text="浏览", command=lambda: select_file(entry_file2))

label_output = tk.Label(root, text="输出文件:")
entry_output = tk.Entry(root, width=50)
button_output_browse = tk.Button(root, text="浏览", command=lambda: on_save_as_button_click(entry_output))

# 添加文本列输入框
label_text_column = tk.Label(root, text="文本列名:")
entry_text_column = tk.Entry(root, width=20)

# 创建比较按钮
button_compare = tk.Button(root, text="比较文件", command=on_compare_button_click)

# 布局
label_file1.grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
entry_file1.grid(row=0, column=1, columnspan=2, pady=5)
button_browse1.grid(row=0, column=3, pady=5)

label_file2.grid(row=1, column=0, padx=10, pady=5, sticky=tk.E)
entry_file2.grid(row=1, column=1, columnspan=2, pady=5)
button_browse2.grid(row=1, column=3, pady=5)

label_output.grid(row=2, column=0, padx=10, pady=5, sticky=tk.E)
entry_output.grid(row=2, column=1, columnspan=2, pady=5)
button_output_browse.grid(row=2, column=3, pady=5)

# 添加文本列输入框
label_text_column.grid(row=3, column=0, padx=10, pady=5, sticky=tk.E)
entry_text_column.grid(row=3, column=1, pady=5)

button_compare.grid(row=4, column=1, pady=10)

# 运行主循环
root.mainloop()
