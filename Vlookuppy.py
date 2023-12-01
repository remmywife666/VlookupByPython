import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import csv
import pandas as pd

def select_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(title="选择文件1")
    if file1_path:
        try:
            file1_extension = file1_path.split(".")[-1]

            if file1_extension == 'csv':
                df = pd.read_csv(file1_path, encoding=encoding_var.get(), dtype=str)  # 读取文件并将所有列设为字符串类型
            elif file1_extension in ['xls', 'xlsx']:
                df = pd.read_excel(file1_path, dtype=str)  # 读取文件并将所有列设为字符串类型
            else:
                messagebox.showinfo("提示", "文件1文件类型不受支持")
                return

            column_names = df.columns.tolist()
            # 清空现有内容
            listbox['menu'].delete(0, 'end')
            for col_name in column_names:
                listbox['menu'].add_command(label=col_name, command=tk._setit(listbox_var, col_name))

        except Exception as e:
            messagebox.showerror("错误", str(e))

        messagebox.showinfo("提示", "文件1列名读取完毕")

def select_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(title="选择文件2")
    if file2_path:
        try:
            file2_extension = file2_path.split(".")[-1]

            if file2_extension == 'csv':
                df = pd.read_csv(file2_path, dtype=str)  # 读取文件并将所有列设为字符串类型
            elif file2_extension in ['xls', 'xlsx']:
                df = pd.read_excel(file2_path, dtype=str)  # 读取文件并将所有列设为字符串类型
            else:
                messagebox.showinfo("提示", "文件2文件类型不受支持")
                return

            column_names = df.columns.tolist()
            # 清空现有内容
            listbox2['menu'].delete(0, 'end')
            for col_name in column_names:
                listbox2['menu'].add_command(label=col_name, command=tk._setit(listbox2_var, col_name))
        except Exception as e:
            messagebox.showerror("错误", str(e))

        messagebox.showinfo("提示", "文件2列名读取完毕")

def match_columns():
    global file1_path, file2_path

    if file1_path and file2_path:
        try:
            result_label.config(text="正在读取文件...")

            file1_extension = file1_path.split(".")[-1]
            file2_extension = file2_path.split(".")[-1]

            # 读取文件1和文件2
            if file1_extension == 'csv':
                df_file1 = pd.read_csv(file1_path, encoding=encoding_var.get(), dtype=str)
            elif file1_extension in ['xls', 'xlsx']:
                df_file1 = pd.read_excel(file1_path, dtype=str)
            else:
                messagebox.showinfo("提示", "文件1文件类型不受支持")
                return

            if file2_extension == 'csv':
                df_file2 = pd.read_csv(file2_path, dtype=str)
            elif file2_extension in ['xls', 'xlsx']:
                df_file2 = pd.read_excel(file2_path, dtype=str)
            else:
                messagebox.showinfo("提示", "文件2文件类型不受支持")
                return

            result_label.config(text="正在匹配并写入新文件...")

            # 进行数据匹配
            merged_df = pd.merge(df_file1, df_file2, how='left', left_on=listbox_var.get(), right_on=listbox2_var.get())

            # 保存结果到新文件
            new_file_name = f"{file1_path.split('/')[-1].split('.')[0]}_V过的文件.csv"  # 新文件名
            merged_df.to_csv(new_file_name, index=False)

            result_label.config(text=f"新文件 '{new_file_name}' 已生成")
        except Exception as e:
            result_label.config(text=f"发生错误：{str(e)}")
    else:
        result_label.config(text="请选择文件1和文件2")


# 创建主窗口
root = tk.Tk()
root.title("文件匹配工具")
root.geometry("600x300")

# 全局文件路径
file1_path = None
file2_path = None

# 编码选择控件1
encoding_label = tk.Label(root, text="选择文件1编码:")
encoding_label.grid(row=0, column=0, sticky="w")

encoding_var = tk.StringVar(root)
encoding_var.set("GBK")

encoding_options = ["UTF-8", "GBK", "latin1"]
encoding_menu = tk.OptionMenu(root, encoding_var, *encoding_options)
encoding_menu.grid(row=0, column=1)

# 文件1选择按钮
select_file1_button = tk.Button(root, text="选择文件1", command=select_file1)
select_file1_button.grid(row=1, column=0)

column_label = tk.Label(root, text="文件1需要v的列名:")
column_label.grid(row=1, column=1, sticky="w")

# 替换 listbox
listbox_var = tk.StringVar(root)
listbox = tk.OptionMenu(root, listbox_var, '')
listbox.grid(row=1, column=2)

encoding_label2 = tk.Label(root, text="选择文件2编码:")
encoding_label2.grid(row=2, column=0, sticky="w")

encoding_var2 = tk.StringVar(root)
encoding_var2.set("GBK")

encoding_menu2 = tk.OptionMenu(root, encoding_var2, *encoding_options)
encoding_menu2.grid(row=2, column=1)

select_file2_button = tk.Button(root, text="选择文件2", command=select_file2)
select_file2_button.grid(row=3, column=0)

lookup_lable = tk.Label(root, text="文件2需要v的列名:")
lookup_lable.grid(row=3, column=1)

# 替换 listbox2
listbox2_var = tk.StringVar(root)
listbox2 = tk.OptionMenu(root, listbox2_var, '')
listbox2.grid(row=3, column=2)

match_button = tk.Button(root, text="确认匹配并生成新文件", command=match_columns)
match_button.grid(row=5, column=1)

# 结果显示标签
result_label = tk.Label(root, text="")
result_label.grid(row=6, columnspan=2)

root.mainloop()
