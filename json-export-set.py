import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import os

# 提取 JSON 数据的通用方法
def extract_value(data, path):
    # 如果data不是字典或者路径为空，直接返回空字符串
    if not isinstance(data, dict) or not path:
        return ""

    # 分割路径
    keys = path.split(".")
    # 初始化当前数据指针
    current_data = data
    # 遍历路径中的每个key
    for key in keys:
        # 处理数组的情况
        if key.endswith("[]"):
            key_base = key[:-2]
            # 如果当前数据是指定名字的列表
            if isinstance(current_data.get(key_base), list):
                # 提取列表中每个元素对应路径的值，组成列表返回
                # 这里使用递归调用extract_value来处理列表中的每个元素
                result = []
                for item in current_data[key_base]:
                    # 递归调用时，需要更新路径，去掉已经处理的数组部分
                    next_path = ".".join(keys[keys.index(key) + 1:])
                    # 注意：这里我们假设如果数组之后的路径为空，就返回元素本身
                    # 如果需要其他处理，可以修改这里的逻辑
                    result.append(extract_value(item, next_path) if next_path else item)
                return result
            else:
                # 如果不是列表，返回空字符串
                return ""
        # 处理普通字典键的情况
        elif key in current_data and isinstance(current_data, dict):
            # 移动到下一个数据层级
            current_data = current_data[key]
        else:
            return ""
    return current_data

# def extract_value(data, path):
#     # 如果data不是字典，或者路径不存在于数据中，直接返回空字符串
#     if not isinstance(data, dict) or not path:
#         return ""
#
#     # 分割路径
#     keys = path.split(".")
#     # 初始化当前数据指针
#     current_data = data
#
#     # 遍历路径中的每个key
#     for key in keys:
#         # 处理数组
#         if key.endswith("[]"):
#             key = key[:-2]
#             # 如果当前数据是指定名字的列表
#             if isinstance(current_data.get(key), list):
#                 # 提取列表中每个元素的最后一个key对应的值，组成列表返回
#                 return [item.get(keys[-1] if keys[-1] != key else None, "") for item in current_data[key]]
#             else:
#                 # 如果不是列表，返回空字符串
#                 return ""
#         # 如果当前数据是字典，并且包含当前key
#         elif key in current_data and isinstance(current_data, dict):
#             # 移动到下一个数据层级
#             current_data = current_data[key]
#         else:
#             # 如果不包含当前key，返回空字符串
#             return ""
#
#     # 如果成功遍历完路径，返回当前数据（通常应该是基础数据类型）
#     return current_data



# 从 JSON 文件读取数据
def load_json(file_path):
    with open(file_path, "r", encoding="utf-8") as file:
        return json.load(file)


def extract_values_from_dict(data, field_mapping):
    record = {}
    for column_name, json_path in field_mapping.items():
        record[column_name] = extract_value(data, json_path)
    return record

CONFIG_FILE = "field_mapping_config.json"  # 配置文件名

class JsonToExcelCsvApp:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON to Excel/CSV Converter_V0.0.5  by tygb")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f0f0")  # 设置背景颜色

        # 初始化变量
        self.json_data = None
        self.field_mapping = {}
        self.dragged_item = None  # 用于记录被拖动的项

        # 设置主窗口在屏幕中间显示
        self.center_window()

        # 界面布局
        self.create_widgets()

        # 加载配置文件
        self.load_config()


    def save_config(self):
        """增量保存字段配置到文件"""
        # 初始化一个空的字典用于存储现有配置
        existing_mapping = {}

        # 检查配置文件是否存在
        if os.path.exists(CONFIG_FILE):
            # 如果存在，读取现有配置
            with open(CONFIG_FILE, "r", encoding="utf-8") as file:
                existing_mapping = json.load(file)

        # 更新配置
        for item in self.mapping_tree.get_children():
            header, path = self.mapping_tree.item(item, "values")
            existing_mapping[header] = path  # 更新或添加字段

        # 保存到文件
        with open(CONFIG_FILE, "w", encoding="utf-8") as file:
            json.dump(existing_mapping, file, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", "字段配置已保存！")

    def load_config(self):
        """从文件加载字段配置"""
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as file:
                field_mapping = json.load(file)
                for header, path in field_mapping.items():
                    self.mapping_tree.insert("", tk.END, values=(header, path))
            messagebox.showinfo("成功", "字段配置已加载！")
        else:
            messagebox.showwarning("警告", "没有找到配置文件！")

    def delete_config(self):
        """删除配置文件"""
        if os.path.exists(CONFIG_FILE):
            os.remove(CONFIG_FILE)
            messagebox.showinfo("成功", "字段配置文件已删除！")
        else:
            messagebox.showwarning("警告", "没有找到配置文件！")


    def center_window(self):
        # 获取屏幕宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 计算窗口的宽度和高度
        window_width = 900
        window_height = 700

        # 计算窗口的 x 和 y 坐标
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        # 设置窗口的位置
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="JSON 转 Excel/CSV 转换器", font=("Arial", 16), bg="#f0f0f0")
        title_label.pack(pady=10)

        # 文件选择
        file_frame = tk.Frame(self.root, bg="#f0f0f0")
        file_frame.pack(fill=tk.X, pady=10, padx=10)

        tk.Label(file_frame, text="选择 JSON 文件:", bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        self.file_entry = tk.Entry(file_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT, padx=5)

        # JSON 预览
        preview_frame = tk.LabelFrame(self.root, text="JSON 文件预览", bg="#f0f0f0")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)

        self.json_preview = tk.Text(preview_frame, wrap=tk.NONE, height=10)
        self.json_preview.pack(fill=tk.BOTH, expand=True)

        # 表头配置
        mapping_frame = tk.LabelFrame(self.root, text="自定义表头配置", bg="#f0f0f0")
        mapping_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)

        # 使用 grid 布局将表格和按钮放在同一行
        mapping_frame.grid_rowconfigure(0, weight=1)  # 让表格占用更多空间
        mapping_frame.grid_columnconfigure(0, weight=9)  # 表格占用9份
        mapping_frame.grid_columnconfigure(1, weight=1)  # 按钮占用1份

        # 表格显示
        self.mapping_tree = ttk.Treeview(mapping_frame, columns=("表头", "JSON路径"), show="headings")
        self.mapping_tree.heading("表头", text="表头")
        self.mapping_tree.heading("JSON路径", text="JSON 路径")
        self.mapping_tree.grid(row=0, column=0, sticky="nsew")  # 表格放在第一列

        # 绑定拖动事件
        self.mapping_tree.bind("<ButtonPress-1>", self.on_item_press)
        self.mapping_tree.bind("<B1-Motion>", self.on_item_drag)
        self.mapping_tree.bind("<ButtonRelease-1>", self.on_item_release)

        # 按钮框架
        button_frame = tk.Frame(mapping_frame, bg="#f0f0f0")
        button_frame.grid(row=0, column=1, sticky="ns")  # 按钮放在第二列

        # 设置按钮框架的宽度
        button_frame.grid_propagate(True)  # 不允许自动调整大小
        button_frame.config(width=150)  # 设置固定宽度

        tk.Button(button_frame, text="添加字段", command=self.add_mapping).pack(side=tk.TOP, padx=5, pady=5)
        tk.Button(button_frame, text="编辑字段", command=self.edit_mapping).pack(side=tk.TOP, padx=5, pady=5)
        tk.Button(button_frame, text="删除字段", command=self.delete_mapping).pack(side=tk.TOP, padx=5, pady=5)
        tk.Button(button_frame, text="清空字段", command=self.clear_fields).pack(side=tk.TOP, padx=5, pady=5)
        tk.Button(button_frame, text="上  移", command=self.move_up).pack(side=tk.TOP, padx=5, pady=5)
        tk.Button(button_frame, text="下  移", command=self.move_down).pack(side=tk.TOP, padx=5, pady=5)

        # 操作按钮
        operation_frame = tk.Frame(self.root, bg="#f0f0f0")
        operation_frame.pack(fill=tk.X, pady=10, padx=10)

        tk.Button(operation_frame, text="预览数据", command=self.preview_data).pack(side=tk.LEFT, padx=5)
        tk.Button(operation_frame, text="导出 Excel", command=lambda: self.export_file("excel")).pack(side=tk.RIGHT, padx=5)
        tk.Button(operation_frame, text="导出 CSV", command=lambda: self.export_file("csv")).pack(side=tk.RIGHT, padx=5)


        tk.Button(operation_frame, text="保存配置", command=self.save_config).pack(side=tk.RIGHT, padx=5)
        tk.Button(operation_frame, text="删除配置", command=self.delete_config).pack(side=tk.RIGHT, padx=5)

        # 状态栏
        self.status_label = tk.Label(self.root, text="欢迎使用 JSON 转换器！", bg="#f0f0f0")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def clear_fields(self):
        # 清空表格中的所有字段
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)

    def on_item_press(self, event):
        # 获取被点击的项
        self.dragged_item = self.mapping_tree.selection()[0]

    def on_item_drag(self, event):
        # 获取鼠标位置
        x, y = event.x, event.y
        # 获取鼠标位置对应的项
        item = self.mapping_tree.identify_row(y)
        if item and item != self.dragged_item:
            # 在新位置插入项
            self.mapping_tree.move(self.dragged_item, self.mapping_tree.parent(item), self.mapping_tree.index(item))

    def on_item_release(self, event):
        # 清空拖动的项
        self.dragged_item = None


    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.load_json_data(file_path)

    def load_json_data(self, file_path):
        try:
            self.json_data = load_json(file_path)
            self.json_preview.delete("1.0", tk.END)
            self.json_preview.insert(tk.END, json.dumps(self.json_data, indent=4, ensure_ascii=False))
            messagebox.showinfo("成功", "JSON 文件已成功加载！")
        except Exception as e:
            messagebox.showerror("错误", f"加载 JSON 文件失败: {e}")
            self.status_label.config(text="加载 JSON 文件失败！")

    def add_mapping(self):
        self.open_mapping_window("添加字段")

    def edit_mapping(self):
        selected_item = self.mapping_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请先选择要编辑的字段！")
            return
        header, path = self.mapping_tree.item(selected_item, "values")
        self.open_mapping_window("编辑字段", header, path)

    def open_mapping_window(self, title, header="", path=""):
        if not self.json_data:
            messagebox.showerror("错误提示", "请先加载 JSON 文件！")
            return

        # 提取所有 JSON 路径
        all_paths = sorted(self.extract_all_paths(self.json_data))

        # 创建弹窗
        mapping_window = tk.Toplevel(self.root)
        mapping_window.title(title)
        mapping_window.geometry("500x200")  # 设置窗口大小

        # **计算屏幕中心位置**
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 500
        window_height = 200
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        mapping_window.geometry(f"{window_width}x{window_height}+{x}+{y}")  # 设置弹窗居中

        # 使用 Frame 来组织布局
        frame = tk.Frame(mapping_window, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # 表头输入
        tk.Label(frame, text="请输入表头:").grid(row=0, column=0, pady=5, sticky=tk.W)
        header_entry = tk.Entry(frame, width=40)
        header_entry.grid(row=0, column=1, pady=5)
        header_entry.insert(0, header)

        # JSON 路径选择
        tk.Label(frame, text="请选择 JSON 路径:").grid(row=1, column=0, pady=5, sticky=tk.W)
        path_combo = ttk.Combobox(frame, values=all_paths, width=50)
        path_combo.grid(row=1, column=1, pady=5)
        path_combo.set(path)

        # 自动填充表头（根据路径的最后一部分）
        def autofill_header(event):
            selected_path = path_combo.get()
            if selected_path and not header_entry.get():
                last_part = selected_path.split(".")[-1].replace("[]", "").replace("[", "").replace("]", "")  # 去除数组索引
                header_entry.insert(0, last_part)

        path_combo.bind("<<ComboboxSelected>>", autofill_header)

        # 添加一些帮助提示
        tk.Label(frame, text="(提示: 表头不能为空，JSON 路径请选择有效的路径)").grid(row=2, column=0, columnspan=2, pady=5, sticky=tk.W)

        # 保存映射
        def save_mapping():
            new_header = header_entry.get().strip()  # 去除空格
            new_path = path_combo.get()
            if not new_header or not new_path:
                messagebox.showerror("错误提示", "表头和 JSON 路径都不能为空！")
                return
            if title == "添加字段":
                self.mapping_tree.insert("", tk.END, values=(new_header, new_path))
            else:
                selected_item = self.mapping_tree.selection()
                if selected_item:
                    self.mapping_tree.item(selected_item, values=(new_header, new_path))
            mapping_window.destroy()

        tk.Button(frame, text="保存并关闭", command=save_mapping).grid(row=3, column=0, columnspan=2, pady=10, sticky=tk.W + tk.E)

    def delete_mapping(self):
        selected_item = self.mapping_tree.selection()
        if selected_item:
            self.mapping_tree.delete(selected_item)
        else:
            messagebox.showerror("错误", "请先选择要删除的字段！")

    def extract_all_paths(self, data, parent_key=""):
        """递归提取 JSON 数据的所有路径"""
        paths = []
        if isinstance(data, dict):
            for key, value in data.items():
                full_key = f"{parent_key}.{key}" if parent_key else key
                paths.extend(self.extract_all_paths(value, full_key))
        elif isinstance(data, list):
            for i, item in enumerate(data):
                array_key = f"{parent_key}[]" if parent_key else ""
                paths.extend(self.extract_all_paths(item, array_key))
                return list(set(paths))  # 去重
        else:
            paths.append(parent_key)
        return sorted(paths)  # 按字母排序路径

    def move_up(self):
        selected_item = self.mapping_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请先选择要移动的字段！")
            return
        index = self.mapping_tree.index(selected_item)
        if index > 0:
            self.mapping_tree.move(selected_item, self.mapping_tree.parent(selected_item), index - 1)

    def move_down(self):
        selected_item = self.mapping_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请先选择要移动的字段！")
            return
        index = self.mapping_tree.index(selected_item)
        self.mapping_tree.move(selected_item, self.mapping_tree.parent(selected_item), index + 1)



    def extract_records_from_json(self, json_data, field_mapping):
        # 判断 JSON 数据是对象还是数组
        if isinstance(json_data, dict):
            # 初始化记录列表
            records = []
            # 初始化一个字典来存储每个字段提取出的值列表（针对数组字段）
            field_values = {}

            # 遍历field_mapping中的每个字段和对应的路径
            for column_name, json_path in field_mapping.items():
                # 提取该路径下的所有值
                extracted_values = extract_value(json_data, json_path)
                # 将提取出的值存储到field_values字典中
                field_values[column_name] = extracted_values

            # 确定记录的数量（基于数组字段的长度）
            # 这里假设所有数组字段提取出的列表长度相同（或至少有一个非空列表作为基准）
            # 如果实际情况不是这样，可能需要额外的逻辑来处理不同长度的列表
            num_records = max(len(values) for values in field_values.values() if isinstance(values, list) and values)

            # 构建记录列表
            for i in range(num_records):
                record = {}
                # 遍历field_values字典中的每个字段和对应的值列表
                for column_name, values in field_values.items():
                    # 如果值是列表，则取当前索引对应的值；否则，如果值不是列表（可能是单个值或None），则直接使用该值（但这里需要处理None的情况）
                    if isinstance(values, list):
                        # 对于数组字段，取当前索引对应的值（如果存在）
                        if i < len(values):
                            record[column_name] = values[i]
                        else:
                            # 如果索引超出列表长度，可以使用None或其他默认值（根据需求决定）
                            record[column_name] = None
                    else:
                        # 对于非数组字段，直接使用提取出的值（但需要注意处理None或单个值的情况）
                        # 这里假设非数组字段只有一个值（或None），并且该值适用于所有记录
                        # 如果实际情况不是这样，可能需要额外的逻辑来处理非数组字段的值
                        record[column_name] = values if values is not None else ""  # 或者使用其他默认值

                # 将构建好的记录添加到记录列表中
                records.append(record)

            # 返回最终构建的记录列表
            return records

        elif isinstance(json_data, list):
            # 如果是数组，就按原来的逻辑处理
            records = []
            for item in json_data:
                record = extract_values_from_dict(item, field_mapping)
                records.append(record)
        else:
            # 如果 JSON 数据既不是对象也不是数组，那就抛出异常或者给出错误提示
            raise ValueError("JSON 数据格式不正确，应该是对象或数组")

        return records


    def preview_data(self):
        if not self.json_data:
            messagebox.showerror("错误", "请先加载 JSON 文件！")
            return

        field_mapping = {}
        for item in self.mapping_tree.get_children():
            header, path = self.mapping_tree.item(item, "values")
            field_mapping[header] = path

        records = self.extract_records_from_json(self.json_data, field_mapping)

        preview_window = tk.Toplevel(self.root)
        preview_window.title("数据预览")
        preview_window.geometry("800x400")

        data_frame = pd.DataFrame(records)

            # 创建一个Frame容器来放置Treeview和Scrollbar
        frame = ttk.Frame(preview_window)
        frame.pack(fill=tk.BOTH, expand=True)

        # 创建Treeview，但不直接pack，而是放在Frame中
        tree = ttk.Treeview(frame, columns=list(data_frame.columns), show="headings")
        for col in data_frame.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, stretch=tk.NO)  # 禁止列自动拉伸

        # 创建垂直滚动条，并将其与Treeview的yscrollcommand绑定
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=vsb.set)

        # 创建水平滚动条（如果需要的话），并将其与Treeview的xscrollcommand绑定
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(xscrollcommand=hsb.set)

        # 将Treeview pack到Frame中（在滚动条创建之后）
        tree.pack(fill=tk.BOTH, expand=True)

        for _, row in data_frame.iterrows():
            tree.insert("", tk.END, values=list(row))

    def export_file(self, file_type):
        if not self.json_data:
            messagebox.showerror("错误", "请先加载 JSON 文件！")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=f".{file_type}",
                                                 filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
        if not file_path:
            return

        field_mapping = {}
        for item in self.mapping_tree.get_children():
            header, path = self.mapping_tree.item(item, "values")
            field_mapping[header] = path

        data_frame = pd.DataFrame(self.extract_records_from_json(self.json_data, field_mapping))

        try:
            if file_type == "excel":
                data_frame.to_excel(file_path, index=False, engine="openpyxl")
            elif file_type == "csv":
                data_frame.to_csv(file_path, index=False)
            messagebox.showinfo("成功", f"数据已成功导出到 {file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出文件失败: {e}")
            # 启动应用程序
if __name__ == "__main__":
    root = tk.Tk()
    app = JsonToExcelCsvApp(root)
    root.mainloop()