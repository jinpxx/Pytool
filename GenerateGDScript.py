import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import ast
import json
import shutil

# Godot项目中Excel表的路径
DataTablePath = "DataTable"

# Godot项目中脚本中的位置
GenerateGDScriptPath = "Script/Auto"

# 配置文件路径
config_file = "config.json"

def preprocess_value(value):
    if isinstance(value, str):
        # 替换数组中的中文分隔符为英文分隔符
        value = value.replace('，', ',')
        # 替换字典中的中文键值对分隔符为英文分隔符
        value = value.replace('：', ':')
    return value

def excel_to_dict(excel_file):
    df = pd.read_excel(excel_file)
    table_name = os.path.splitext(os.path.basename(excel_file))[0]
    return {table_name: df.to_dict(orient="records")}

def get_file_times(file_path):
    creation_time = os.path.getctime(file_path)
    modification_time = os.path.getmtime(file_path)
    creation_time_str = datetime.fromtimestamp(creation_time).strftime('%Y年%m月%d日 %H:%M:%S')
    modification_time_str = datetime.fromtimestamp(modification_time).strftime('%Y年%m月%d日 %H:%M:%S')
    return creation_time_str, modification_time_str

def generate_main_gdscript(data_dicts, gdscript_file, file_times):
    with open(gdscript_file, "w", encoding="utf-8") as f:
        f.write("# Auto Dictionary\n")
        f.write(f"# {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')} 通过工具生成\n")
        f.write(f"# 该脚本根据Excel文件生成，请勿手动修改\n")
        f.write(f"# 定义静态类 DataTable\n")
        f.write(f"class_name DataTable\n\n")

        for table_name, records in data_dicts.items():
            creation_time, modification_time = file_times[table_name]
            f.write(f"# 配置文件名为: {table_name}\n")
            f.write(f"# 文件创建时间: {creation_time}\n")
            f.write(f"# 最后修改时间: {modification_time}\n")
            table_name_upper = table_name.upper()  # 将表名转换为大写
            f.write(f"const {table_name_upper} : Dictionary = {{\n")
            
            # 计算每列字段值的最大长度
            field_max_lengths = {field: max(len(field), max(len(str(record.get(field, ""))) for record in records)) for field in records[0].keys()}

            for idx, record in enumerate(records, start=1):
                f.write(f"\t{idx}: ")
                f.write("{ ")
                for field, value in record.items():
                    value = preprocess_value(value)
                    
                    # 解析字符串中的数组和字典
                    if isinstance(value, str) and (value.startswith("[") and value.endswith("]") or value.startswith("{") and value.endswith("}")):
                        try:
                            value = ast.literal_eval(value)
                        except (ValueError, SyntaxError):
                            value = f'"{value}"'
                    
                    # 处理字符串值
                    if isinstance(value, str):
                        value = f'"{value}"'
                    
                    # 处理列表值，确保格式正确
                    elif isinstance(value, list):
                        value = str(value).replace('[', '[ ').replace(']', ' ]')
                    
                    # 计算适当的填充空格以对齐字段和值
                    padding = ' ' * (field_max_lengths[field] - len(str(value)) + 2)  # 加2是为了与字段名后的空格对齐
                    f.write(f'"{field}" : {value},{padding} ')

                f.write("},\n")
            f.write("}\n\n")



def create_directories_if_not_exist(project_path):
    excel_dir = os.path.join(project_path, DataTablePath)
    auto_dir = os.path.join(project_path, GenerateGDScriptPath)

    if not os.path.exists(excel_dir):
        os.makedirs(excel_dir)
        print(f"创建目录: {excel_dir}")

    if not os.path.exists(auto_dir):
        os.makedirs(auto_dir)
        print(f"创建目录: {auto_dir}")

def process_files(project_path):
    create_directories_if_not_exist(project_path)

    excel_dir = os.path.join(project_path, DataTablePath)

    excel_files = [
        filename
        for filename in os.listdir(excel_dir)
        if filename.endswith(".xlsx") or filename.endswith(".xls")
    ]

    if not excel_files:
        messagebox.showerror("错误", "未找到任何 Excel 文件。")
        return

    data_dicts = {}
    file_times = {}
    for excel_file in excel_files:
        try:
            table_name = os.path.splitext(excel_file)[0]
            file_path = os.path.join(excel_dir, excel_file)
            data_dict = excel_to_dict(file_path)
            data_dicts.update(data_dict)
            file_times[table_name] = get_file_times(file_path)
        except PermissionError as e:
            messagebox.showerror("错误", f"无法读取文件 {excel_file}。请确保文件未在其他程序中打开，然后重试。")
            return

    gdscript_file = os.path.join(project_path, GenerateGDScriptPath, "DataTable.gd")
    generate_main_gdscript(data_dicts, gdscript_file, file_times)
    print(f"生成 {gdscript_file}")

    messagebox.showinfo("完成", "Excel 文件已成功转换并生成 GDScript 文件。")

def delete_generated_scripts(project_path):
    auto_dir = os.path.join(project_path, GenerateGDScriptPath)
    if os.path.isdir(auto_dir):
        for filename in os.listdir(auto_dir):
            file_path = os.path.join(auto_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"删除 {file_path}")
        messagebox.showinfo("完成", "已成功删除生成的 GDScript 文件。")
    else:
        messagebox.showerror("错误", f"未找到生成的 GDScript 文件目录: {auto_dir}")

def save_config(project_path):
    config = {"last_project_path": project_path}
    with open(config_file, "w") as f:
        json.dump(config, f)

def load_config():
    if os.path.isfile(config_file):
        with open(config_file, "r") as f:
            config = json.load(f)
            return config.get("last_project_path")
    return None

def select_project_path():
    project_path = filedialog.askdirectory(title="选择 Godot 项目路径")
    if project_path:
        project_godot_path = os.path.join(project_path, "project.godot")
        if os.path.isfile(project_godot_path):
            entry_path.delete(0, tk.END)
            entry_path.insert(0, project_path)
            save_config(project_path)
            print("有效的 Godot 项目路径已选择。")
        else:
            messagebox.showerror("错误", "所选文件夹中没有找到 'project.godot' 文件。")
    else:
        print("未选择路径。")

def start_processing():
    project_path = entry_path.get()
    if not project_path or not os.path.isdir(project_path):
        messagebox.showerror("错误", "请选择有效的 Godot 项目路径。")
    else:
        save_config(project_path)
        process_files(project_path)

def start_deleting():
    project_path = entry_path.get()
    if not project_path or not os.path.isdir(project_path):
        messagebox.showerror("错误", "请选择有效的 Godot 项目路径。")
    else:
        delete_generated_scripts(project_path)

def start_cleaning():
    project_path = entry_path.get()
    if not project_path or not os.path.isdir(project_path):
        messagebox.showerror("错误", "请选择有效的 Godot 项目路径。")
    else:
        remove_macosx_and_files(project_path)
        messagebox.showinfo("完成", "已成功删除指定文件和文件夹。")

def remove_macosx_and_files(project_path):
    # 遍历指定文件夹
    for root, dirs, files in os.walk(project_path, topdown=False):
        # 删除 __MACOSX 文件夹
        if '__MACOSX' in dirs:
            macosx_path = os.path.join(root, '__MACOSX')
            try:
                shutil.rmtree(macosx_path)
                print(f"Deleted folder and its contents: {macosx_path}")
            except OSError as e:
                print(f"Failed to delete folder {macosx_path}: {e}")
        
        # 删除 .DS_Store 和 ._.DS_Store 文件
        for file in files:
            if file.endswith('.DS_Store') or file.endswith('._.DS_Store'):
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
                except OSError as e:
                    print(f"Error deleting file {file_path}: {e}")


# 创建主窗口
root = tk.Tk()
root.title("Excel To GDScript")

# 创建路径选择组件
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="选择 Godot 项目路径:")
label.pack(side=tk.LEFT)

entry_path = tk.Entry(frame, width=50)
entry_path.pack(side=tk.LEFT, padx=5)

button_browse = tk.Button(frame, text="浏览", command=select_project_path)
button_browse.pack(side=tk.LEFT)

# 创建开始按钮
button_start = tk.Button(root, text="开始转换", command=start_processing)
button_start.pack(pady=5)

# 创建删除按钮
button_delete = tk.Button(root, text="删除生成的脚本", command=start_deleting)
button_delete.pack(pady=5)

# 创建清理项目素材按钮
button_delete = tk.Button(root, text="清理 MACOSX 和 DS_Store 文件", command=start_cleaning)
button_delete.pack(pady=5)

# 获取屏幕宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 计算窗口宽度和高度
window_width = 600  # 自行调整宽度，确保足够展示按钮
window_height = 180

# 计算窗口居中时左上角的坐标
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# 设置窗口的初始尺寸和位置
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 加载上次选择的路径
last_project_path = load_config()
if last_project_path:
    entry_path.insert(0, last_project_path)

# 运行主循环
root.mainloop()
