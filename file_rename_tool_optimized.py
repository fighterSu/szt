import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import pandas as pd
import shutil
import logging
from datetime import datetime
from collections import defaultdict
import re
import subprocess


class FileRenameTool:
    # --- 状态常量 ---
    STATUS_PENDING = "待处理"
    STATUS_SUCCESS = "成功"
    STATUS_SKIPPED = "已跳过"
    STATUS_UNMATCHED = "未匹配"
    STATUS_PREVIEW_CONFLICT = "预览冲突"
    STATUS_TARGET_EXISTS = "目标已存在"
    STATUS_ERROR_PREFIX = "错误: "

    # --- 其他常量 ---
    MAX_CSV_SIZE = 10 * 1024 * 1024  # 10MB
    MAX_FILENAME_LENGTH = 255

    def __init__(self, root):
        self.root = root
        self.root.title("文件/夹批量重命名工具 v3.0 (优化版)")
        self.root.geometry("1200x900")  # 稍微增加宽度以适应紧凑布局

        # 提前初始化UI组件变量
        self.log_text = None
        self.load_button = None
        self.preview_button = None
        self.execute_button = None
        self.reset_button = None
        self.compress_checkbutton = None
        self.compress_format_combo = None

        self.init_logging()
        self.set_app_icon()

        # 变量初始化
        self.csv_file_path = ""
        self.source_folder_path = ""
        self.target_folder_path = ""
        self.name_column = ""
        self.filename_column = ""
        self.file_mapping = {}
        self.matched_files = defaultdict(list)
        self.operation_log = []

        self.create_widgets()
        self.update_button_states()
        self._update_folder_options_state()

        self.center_window()

    def center_window(self):
        """将窗口置于屏幕中央"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def init_logging(self):
        """初始化日志系统"""
        log_dir = os.path.join(os.path.expanduser("~"), "FileRenameToolLogs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"rename_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        logging.basicConfig(
            filename=log_file, level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8'
        )
        self.log_file_path = log_file

    def log_operation(self, message, is_error=False, is_success=False):
        """记录并显示日志"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        full_message = f"{timestamp} - {message}"
        self.operation_log.append(full_message)
        if len(self.operation_log) > 1000: self.operation_log.pop(0)

        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            tag = "info"
            if is_error:
                tag = "error"
            elif is_success:
                tag = "success"
            self.log_text.insert(tk.END, full_message + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        if is_error:
            logging.error(message)
        else:
            logging.info(message)

    def set_app_icon(self):
        """设置程序图标"""
        try:
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_path, "app_icon.ico")
            if os.path.exists(icon_path): self.root.iconbitmap(icon_path)
        except Exception as e:
            self.log_operation(f"加载图标失败: {e}", is_error=True)

    def create_widgets(self):
        # 主框架，减少内边距
        main_frame = ttk.Frame(self.root, padding="5")
        main_frame.pack(fill="both", expand=True)

        # --- 步骤 1: 路径选择 (紧凑布局) ---
        path_frame = ttk.LabelFrame(main_frame, text="步骤 1: 选择路径", padding="5")
        path_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=2)
        path_frame.columnconfigure(1, weight=1)

        # 使用更紧凑的行间距
        ttk.Label(path_frame, text="对应表:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.csv_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.csv_path_var).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_frame, text="浏览", command=self.select_csv_file, width=8).grid(row=0, column=2, padx=5, pady=2)

        ttk.Label(path_frame, text="源文件夹:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.source_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.source_path_var).grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_frame, text="浏览", command=self.select_source_folder, width=8).grid(row=1, column=2, padx=5, pady=2)

        ttk.Label(path_frame, text="目标文件夹:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.target_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.target_path_var).grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_frame, text="浏览", command=self.select_target_folder, width=8).grid(row=2, column=2, padx=5, pady=2)

        # --- 步骤 2: 规则配置 (更紧凑的布局) ---
        rules_frame = ttk.LabelFrame(main_frame, text="步骤 2: 配置规则", padding="5")
        rules_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)
        
        # 使用单行布局而不是两列
        rules_container = ttk.Frame(rules_frame)
        rules_container.pack(fill="x", expand=True)

        # 对应关系
        map_frame = ttk.Frame(rules_container)
        map_frame.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Label(map_frame, text="姓名列:").pack(side="left", padx=2)
        self.name_column_var = tk.StringVar()
        self.name_column_combo = ttk.Combobox(map_frame, textvariable=self.name_column_var, state="readonly", width=15)
        self.name_column_combo.pack(side="left", padx=2)
        ttk.Label(map_frame, text="目标列:").pack(side="left", padx=(10, 2))
        self.filename_column_var = tk.StringVar()
        self.filename_column_combo = ttk.Combobox(map_frame, textvariable=self.filename_column_var, state="readonly", width=15)
        self.filename_column_combo.pack(side="left", padx=2)

        # 匹配规则
        match_frame = ttk.Frame(rules_container)
        match_frame.pack(side="left", padx=10, fill="x")
        ttk.Label(match_frame, text="对象:").pack(side="left", padx=2)
        self.rename_mode_var = tk.StringVar(value="文件")
        ttk.Radiobutton(match_frame, text="文件", variable=self.rename_mode_var, value="文件",
                        command=self._update_folder_options_state).pack(side="left", padx=2)
        ttk.Radiobutton(match_frame, text="文件夹", variable=self.rename_mode_var, value="文件夹",
                        command=self._update_folder_options_state).pack(side="left", padx=2)
        
        ttk.Label(match_frame, text="匹配:").pack(side="left", padx=(10, 2))
        self.match_type_var = tk.StringVar(value="包含")
        ttk.Radiobutton(match_frame, text="包含", variable=self.match_type_var, value="包含").pack(side="left", padx=2)
        ttk.Radiobutton(match_frame, text="精确", variable=self.match_type_var, value="精确").pack(side="left", padx=2)

        # 输出选项
        output_frame = ttk.Frame(rules_container)
        output_frame.pack(side="left", padx=10, fill="x")
        ttk.Label(output_frame, text="重名:").pack(side="left", padx=2)
        self.duplicate_handling_var = tk.StringVar(value="序号")
        ttk.Radiobutton(output_frame, text="序号", variable=self.duplicate_handling_var, value="序号").pack(side="left", padx=2)
        ttk.Radiobutton(output_frame, text="跳过", variable=self.duplicate_handling_var, value="跳过").pack(side="left", padx=2)

        self.compress_folder_var = tk.BooleanVar(value=False)
        self.compress_checkbutton = ttk.Checkbutton(output_frame, text="压缩", variable=self.compress_folder_var)
        self.compress_checkbutton.pack(side="left", padx=(10, 2))
        self.compress_format_var = tk.StringVar(value="zip")
        self.compress_format_combo = ttk.Combobox(output_frame, textvariable=self.compress_format_var,
                                                  values=["zip", "7z", "rar"], state="readonly", width=6)
        self.compress_format_combo.pack(side="left", padx=2)

        # --- 步骤 3: 操作按钮 (紧凑布局) ---
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=2)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(action_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=5, pady=2)

        # 按钮容器
        button_frame = ttk.Frame(action_frame)
        button_frame.pack(fill="x", padx=5, pady=2)
        
        # 安全提示和按钮在同一行
        safety_label = ttk.Label(button_frame, text="▲安全提示: 本工具会复制文件/夹到目标位置",
                                 foreground="blue", font=("SimHei", 9))
        safety_label.pack(side="left", padx=5)

        # 操作按钮
        self.load_button = ttk.Button(button_frame, text="①加载", command=self.load_files, width=12)
        self.load_button.pack(side="left", padx=2)
        self.preview_button = ttk.Button(button_frame, text="②预览", command=self.generate_preview, width=12)
        self.preview_button.pack(side="left", padx=2)
        self.execute_button = ttk.Button(button_frame, text="③执行", command=self.execute_rename, width=12)
        self.execute_button.pack(side="left", padx=2)
        self.reset_button = ttk.Button(button_frame, text="重置", command=self.reset_all, width=8)
        self.reset_button.pack(side="left", padx=2)
        ttk.Button(button_frame, text="日志", command=self.show_log, width=8).pack(side="left", padx=2)
        ttk.Button(button_frame, text="退出", command=self.root.quit, width=8).pack(side="right", padx=5)

        # --- 预览与日志 (调整比例) ---
        paned_window = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned_window.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=2)

        # 预览结果框架
        result_frame = ttk.LabelFrame(paned_window, text="预览与执行结果", padding="2")
        paned_window.add(result_frame, weight=5)  # 增加权重，占更多空间
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        # 创建Treeview
        columns = ("序号", "匹配的姓名", "原始名称", "目标名称", "状态")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=20)  # 设置初始高度
        for col in columns: 
            self.tree.heading(col, text=col)
        self.tree.column("序号", width=50, anchor="center")
        self.tree.column("匹配的姓名", width=150)
        self.tree.column("原始名称", width=350)
        self.tree.column("目标名称", width=350)
        self.tree.column("状态", width=100)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 布局
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        # 日志框架（减小高度）
        log_frame = ttk.LabelFrame(paned_window, text="操作日志", padding="2")
        paned_window.add(log_frame, weight=1)  # 减少权重，占较少空间
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, state=tk.DISABLED, wrap=tk.WORD, height=5)  # 减小初始高度
        self.log_text.tag_config("info", foreground="black")
        self.log_text.tag_config("success", foreground="#008000")
        self.log_text.tag_config("error", foreground="#FF0000")
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.grid(row=4, column=0, columnspan=2, sticky="ew")

        # 配置网格权重
        main_frame.rowconfigure(3, weight=1)  # PanedWindow 所在行可伸缩
        main_frame.columnconfigure(0, weight=1)

    def reset_all(self):
        """重置所有输入、设置和结果"""
        if not messagebox.askyesno("确认", "确定要清空所有设置和预览结果吗？"):
            return

        # 清空路径变量
        self.csv_path_var.set("")
        self.source_path_var.set("")
        self.target_path_var.set("")
        self.csv_file_path = ""
        self.source_folder_path = ""
        self.target_folder_path = ""

        # 重置下拉框
        self.name_column_var.set("")
        self.filename_column_var.set("")
        self.name_column_combo['values'] = []
        self.filename_column_combo['values'] = []

        # 重置单选和复选框
        self.rename_mode_var.set("文件")
        self.match_type_var.set("包含")
        self.duplicate_handling_var.set("序号")
        self.compress_folder_var.set(False)
        self.compress_format_var.set("zip")

        # 清空内部数据
        self.file_mapping.clear()
        self.matched_files.clear()
        self.operation_log.clear()

        # 清空UI结果
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        # 重置状态
        self.status_var.set("已重置，请重新开始。")
        self.progress_var.set(0)
        self.log_operation("=== 用户重置了所有操作 ===")
        self.update_button_states()
        self._update_folder_options_state()

    def _update_folder_options_state(self):
        """根据是否选择"文件夹"模式，启用或禁用压缩选项"""
        if self.rename_mode_var.get() == "文件夹":
            self.compress_checkbutton.config(state=tk.NORMAL)
            self.compress_format_combo.config(state="readonly")
        else:
            self.compress_checkbutton.config(state=tk.DISABLED)
            self.compress_format_combo.config(state=tk.DISABLED)
            self.compress_folder_var.set(False)

    def _sanitize_filename(self, filename):
        """移除或替换掉文件名中的非法字符，并限制长度。"""
        if not filename: return "未命名"
        sanitized = re.sub(r'[<>:"/\\|?*]', '', str(filename))
        sanitized = re.sub(r'[\x00-\x1F\x7F]', '', sanitized)
        sanitized = re.sub(r'\.{2,}', '.', sanitized)
        sanitized = re.sub(r'\s+', ' ', sanitized).strip()
        name, ext = os.path.splitext(sanitized)
        if len(sanitized) > self.MAX_FILENAME_LENGTH:
            max_name_length = self.MAX_FILENAME_LENGTH - len(ext)
            if max_name_length > 0:
                name = name[:max_name_length]
                sanitized = f"{name}{ext}"
            else:
                sanitized = sanitized[:self.MAX_FILENAME_LENGTH]
        if not sanitized: sanitized = "未命名"
        return sanitized

    def update_button_states(self):
        """根据程序当前状态更新按钮的可用性"""
        has_paths = self.csv_file_path and self.source_folder_path and self.target_folder_path
        has_columns = self.name_column_var.get() and self.filename_column_var.get()
        self.load_button.config(state=tk.NORMAL if has_paths and has_columns else tk.DISABLED)

        self.preview_button.config(state=tk.NORMAL if self.matched_files else tk.DISABLED)

        can_execute = False
        if self.tree.get_children():
            for item in self.tree.get_children():
                item_values = self.tree.item(item, "values")
                if item_values and item_values[-1] == self.STATUS_PENDING:
                    can_execute = True
                    break
        self.execute_button.config(state=tk.NORMAL if can_execute else tk.DISABLED)

    def select_csv_file(self):
        file_path = filedialog.askopenfilename(title="选择姓名与文件名对应表",
                                               filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"),
                                                          ("所有文件", "*.*")])
        if not file_path: return
        try:
            if os.path.getsize(file_path) > self.MAX_CSV_SIZE:
                messagebox.showerror("文件过大", f"选择的文件超过了最大限制 {self.MAX_CSV_SIZE / 1024 / 1024:.1f}MB。")
                return
        except Exception as e:
            messagebox.showerror("错误", f"检查文件大小时出错: {e}")
            return
        self.csv_path_var.set(file_path)
        self.csv_file_path = file_path
        self.load_csv_columns()
        self.log_operation(f"已选择对应表: {file_path}")
        self.update_button_states()

    def select_source_folder(self):
        folder_path = filedialog.askdirectory(title="选择源文件夹")
        if folder_path:
            self.source_path_var.set(folder_path)
            self.source_folder_path = folder_path
            self.log_operation(f"已选择源文件夹: {folder_path}")
            self.update_button_states()

    def select_target_folder(self):
        folder_path = filedialog.askdirectory(title="选择重命名后存放文件夹")
        if folder_path:
            self.target_path_var.set(folder_path)
            self.target_folder_path = folder_path
            self.log_operation(f"已选择目标文件夹: {folder_path}")
            self.update_button_states()

    def load_csv_columns(self):
        try:
            df = pd.read_excel(self.csv_file_path, nrows=5) if self.csv_file_path.lower().endswith(
                ('.xlsx', '.xls')) else pd.read_csv(self.csv_file_path, nrows=5, encoding='utf-8-sig')
            columns = df.columns.tolist()
            self.name_column_combo['values'] = columns
            self.filename_column_combo['values'] = columns
            if len(columns) >= 2:
                self.name_column_var.set(columns[0])
                self.filename_column_var.set(columns[1])
            self.status_var.set(f"对应表加载成功，找到 {len(columns)} 列。请确认列选择是否正确。")
            self.log_operation(f"对应表列加载成功: {', '.join(columns)}")
        except Exception as e:
            messagebox.showerror("错误", f"加载对应表列失败: {e}")
            self.log_operation(f"加载对应表列失败: {e}", is_error=True)
            self.status_var.set("加载对应表列失败")
        finally:
            self.update_button_states()

    def load_files(self):
        if not all([self.csv_file_path, self.source_folder_path, self.target_folder_path, self.name_column_var.get(),
                    self.filename_column_var.get()]):
            messagebox.showwarning("警告", "请先完成步骤1和步骤2中的所有选择。")
            return

        self.matched_files.clear()
        for item in self.tree.get_children(): self.tree.delete(item)
        self.update_button_states()

        try:
            self.status_var.set("正在加载和校验对应表...")
            self.root.update_idletasks()

            df = pd.read_excel(self.csv_file_path) if self.csv_file_path.lower().endswith(
                ('.xlsx', '.xls')) else pd.read_csv(self.csv_file_path, encoding='utf-8-sig')
            name_col, filename_col = self.name_column_var.get(), self.filename_column_var.get()

            if name_col not in df.columns or filename_col not in df.columns:
                messagebox.showerror("列名错误", f"指定的列 '{name_col}' 或 '{filename_col}' 在文件中不存在。")
                return

            df.dropna(subset=[name_col, filename_col], inplace=True)
            df[name_col] = df[name_col].astype(str).str.strip()
            df[filename_col] = df[filename_col].astype(str).str.strip()

            is_duplicate_series = df[name_col].str.lower().duplicated(keep=False)
            name_duplicates = df[is_duplicate_series]
            if not name_duplicates.empty:
                duplicate_names = sorted(name_duplicates[name_col].unique())
                messagebox.showerror("发现重名错误",
                                     f"在对应表的"{name_col}"列中发现重复的姓名（已忽略大小写）：\n\n{', '.join(duplicate_names)}\n\n请修正源文件后重试。")
                self.status_var.set("错误：姓名列存在重名。")
                return

            filename_duplicates = df[df[filename_col].duplicated(keep=False)]
            if not filename_duplicates.empty:
                error_details_list = "\n".join([f"- {row[filename_col]} (来自: {row[name_col]})" for _, row in
                                                filename_duplicates.sort_values(by=filename_col).iterrows()])
                self._show_scrollable_error_dialog(title="发现目标文件名重复错误",
                                                   message=f"在"{filename_col}"列中发现重复的目标文件名，请修正后重试。",
                                                   details_text=error_details_list)
                self.status_var.set("错误：目标名称列存在重名。")
                return

            self.file_mapping = {row[name_col]: row[filename_col] for _, row in df.iterrows()}
            self.status_var.set("正在扫描源文件夹...")
            self.root.update_idletasks()

            rename_mode = self.rename_mode_var.get()
            if rename_mode == "文件":
                source_items = [f for f in os.listdir(self.source_folder_path) if
                                os.path.isfile(os.path.join(self.source_folder_path, f))]
            else:
                source_items = [d for d in os.listdir(self.source_folder_path) if
                                os.path.isdir(os.path.join(self.source_folder_path, d))]

            unmatched_items, processed_items = [], set()
            sorted_names = sorted(self.file_mapping.keys(), key=len, reverse=True)

            for item_name in source_items:
                if item_name in processed_items: continue
                is_matched = False
                for name in sorted_names:
                    item_base_name = os.path.splitext(item_name)[0]
                    match_found = (self.match_type_var.get() == "包含" and name.lower() in item_name.lower()) or \
                                  (self.match_type_var.get() != "包含" and item_base_name.lower() == name.lower())
                    if match_found:
                        self.matched_files[name].append(item_name)
                        processed_items.add(item_name)
                        is_matched = True
                        break
                if not is_matched: unmatched_items.append(item_name)

            for i, item_name in enumerate(unmatched_items):
                self.tree.insert("", "end", values=(i + 1, "N/A", item_name, "", self.STATUS_UNMATCHED))

            match_count = sum(len(files) for files in self.matched_files.values())
            status_text = f"扫描完成：找到 {len(source_items)} 个{rename_mode}，匹配 {match_count} 个，未匹配 {len(unmatched_items)} 个。"
            self.status_var.set(status_text)
            self.log_operation(status_text)
            self.update_button_states()

        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {e}")
            self.log_operation(f"加载文件失败: {e}", is_error=True)
            self.status_var.set("加载文件失败")
            self.update_button_states()

    def generate_preview(self):
        if not self.matched_files:
            messagebox.showwarning("警告", "请先加载并匹配，或没有可匹配项。")
            return

        items_to_keep = [self.tree.item(item_id, "values") for item_id in self.tree.get_children() if
                         self.tree.item(item_id, "values")[4] == self.STATUS_UNMATCHED]
        for item_id in self.tree.get_children(): self.tree.delete(item_id)

        target_filenames_in_preview = set()
        item_index = 1
        rename_mode = self.rename_mode_var.get()
        is_compressing = self.compress_folder_var.get() and rename_mode == "文件夹"

        for name, original_items in self.matched_files.items():
            base_target_name_from_csv = self._sanitize_filename(self.file_mapping.get(name, "未命名"))
            base_name_for_ops, _ = os.path.splitext(base_target_name_from_csv)

            if rename_mode == "文件":
                base_target_name = base_target_name_from_csv
            elif is_compressing:
                compress_format = self.compress_format_var.get()
                base_target_name = f"{base_name_for_ops}.{compress_format}"
            else:  # 文件夹非压缩
                base_target_name = base_target_name_from_csv

            if len(original_items) > 1 and self.duplicate_handling_var.get() == "跳过":
                for original_item in original_items:
                    self.tree.insert("", "end",
                                     values=(item_index, name, original_item, base_target_name, self.STATUS_SKIPPED))
                    item_index += 1
                continue

            for i, original_item in enumerate(original_items):
                final_target_name = base_target_name
                # 如果一个姓名匹配了多个文件/夹，且处理方式是"添加序号"
                if len(original_items) > 1 and self.duplicate_handling_var.get() == "序号":
                    target_main, target_ext = os.path.splitext(base_target_name)
                    final_target_name = f"{target_main}_{i + 1}{target_ext}"

                final_target_name = self._sanitize_filename(os.path.basename(final_target_name))

                status = self.STATUS_PENDING
                if final_target_name.lower() in (fn.lower() for fn in target_filenames_in_preview):
                    status = self.STATUS_PREVIEW_CONFLICT
                elif os.path.exists(os.path.join(self.target_folder_path, final_target_name)):
                    status = self.STATUS_TARGET_EXISTS

                if status == self.STATUS_PENDING:
                    target_filenames_in_preview.add(final_target_name)

                self.tree.insert("", "end", values=(item_index, name, original_item, final_target_name, status))
                item_index += 1

        for values in items_to_keep: self.tree.insert("", "end", values=values)
        self.status_var.set(f"生成预览完成，共 {len(self.tree.get_children())} 个项目。")
        self.log_operation(f"生成预览完成，共 {len(self.tree.get_children())} 个项目。")
        self.update_button_states()

    def execute_rename(self):
        pending_items = [item for item in self.tree.get_children() if
                         self.tree.item(item, "values")[-1] == self.STATUS_PENDING]
        if not pending_items:
            messagebox.showwarning("警告", "没有状态为"待处理"的可执行项目。")
            return
        if not messagebox.askyesno("确认操作", f"即将开始处理 {len(pending_items)} 个项目。\n确定要继续吗？"):
            return

        rename_mode = self.rename_mode_var.get()
        is_compressing = self.compress_folder_var.get() and rename_mode == "文件夹"
        compress_format = self.compress_format_var.get()

        if is_compressing and compress_format in ["7z", "rar"]:
            tool = "7z.exe" if compress_format == "7z" else "Rar.exe"
            if not shutil.which(tool):
                messagebox.showerror("依赖缺失",
                                     f"未找到 '{tool}' 命令。\n请确认您已安装相应压缩软件(WinRAR或7-Zip)并将其程序目录添加至系统环境变量PATH。")
                self.log_operation(f"依赖缺失: {tool}", is_error=True)
                return

        try:
            os.makedirs(self.target_folder_path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建目标文件夹: {e}")
            self.log_operation(f"无法创建目标文件夹: {e}", is_error=True)
            return

        success_count, skip_count, error_count = 0, 0, 0
        total_items = len(self.tree.get_children())
        processed_count = 0

        for item in self.tree.get_children():
            processed_count += 1
            self.progress_var.set(int(100 * processed_count / total_items))
            self.root.update_idletasks()

            values = list(self.tree.item(item, "values"))

            if len(values) < 5 or values[4] != self.STATUS_PENDING:
                if len(values) >= 5 and values[4] != self.STATUS_SUCCESS and not values[4].startswith(
                        self.STATUS_ERROR_PREFIX):
                    skip_count += 1
                continue

            original_name, target_name = values[2], values[3]
            src_path = os.path.join(self.source_folder_path, original_name)
            dst_path = os.path.join(self.target_folder_path, target_name)

            try:
                if os.path.exists(dst_path): raise FileExistsError("目标已存在")
                if not os.path.exists(src_path): raise FileNotFoundError("源不存在")

                if is_compressing:
                    self.status_var.set(f"正在压缩: {original_name}")
                    self.root.update_idletasks()
                    dst_path_base = os.path.splitext(dst_path)[0]

                    if compress_format == 'zip':
                        shutil.make_archive(base_name=dst_path_base, format='zip', root_dir=self.source_folder_path,
                                            base_dir=original_name)
                    else:  # 7z 或 rar
                        abs_src_path = os.path.abspath(src_path)
                        abs_dst_path = os.path.abspath(dst_path)
                        cmd = []
                        if compress_format == '7z':
                            cmd = ['7z.exe', 'a', '-t7z', abs_dst_path, abs_src_path]
                        elif compress_format == 'rar':
                            cmd = ['Rar.exe', 'a', '-r', abs_dst_path, abs_src_path]

                        subprocess.run(cmd, check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)

                elif rename_mode == "文件":
                    shutil.copy2(src_path, dst_path)
                else:
                    shutil.copytree(src_path, dst_path)

                values[4] = self.STATUS_SUCCESS
                success_count += 1
                self.log_operation(f"'{original_name}' 成功处理为 '{target_name}'", is_success=True)

            except Exception as e:
                error_msg = f"{type(e).__name__}: {str(e)}"
                if isinstance(e, subprocess.CalledProcessError):
                    error_output = e.stderr.decode(sys.getdefaultencoding(), errors='ignore') if e.stderr else '未知压缩错误'
                    error_msg = f"压缩命令失败: {error_output}"
                values[4] = f"{self.STATUS_ERROR_PREFIX}{error_msg[:60]}"
                error_count += 1
                self.log_operation(f"处理 '{original_name}' 失败: {error_msg}", is_error=True)

            self.tree.item(item, values=tuple(values))

        summary_msg = f"操作完成！\n\n成功: {success_count} 个\n跳过 (含未匹配): {skip_count} 个\n错误: {error_count} 个"
        self.status_var.set(f"操作完成: 成功 {success_count}, 跳过 {skip_count}, 错误 {error_count}")
        self.log_operation(summary_msg)
        messagebox.showinfo("完成", summary_msg)
        self.update_button_states()
        self.progress_var.set(0)

    def show_log(self):
        if not hasattr(self, 'log_window') or not self.log_window.winfo_exists():
            self.log_window = tk.Toplevel(self.root)
            self.log_window.title("操作日志")
            self.log_window.geometry("800x600")
            log_text_widget = scrolledtext.ScrolledText(self.log_window, wrap=tk.WORD)
            log_text_widget.pack(fill="both", expand=True, padx=10, pady=10)
            log_text_widget.tag_config("info", foreground="black")
            log_text_widget.tag_config("success", foreground="#008000")
            log_text_widget.tag_config("error", foreground="#FF0000")
            try:
                with open(self.log_file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if " - ERROR - " in line:
                            log_text_widget.insert(tk.END, line, "error")
                        elif "成功" in line or "完成" in line:
                            log_text_widget.insert(tk.END, line, "success")
                        else:
                            log_text_widget.insert(tk.END, line, "info")
            except Exception as e:
                log_text_widget.insert(tk.END, f"无法读取日志文件: {e}\n", "error")
            log_text_widget.config(state=tk.DISABLED)
            ttk.Button(self.log_window, text="关闭", command=self.log_window.destroy).pack(pady=10)
        else:
            self.log_window.lift()

    def _show_scrollable_error_dialog(self, title, message, details_text):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill="both", expand=True)
        message_label = ttk.Label(main_frame, text=message, wraplength=550)
        message_label.pack(side="top", fill="x", pady=(0, 10))
        text_frame = ttk.Frame(main_frame, borderwidth=1, relief="sunken")
        text_frame.pack(side="top", fill="both", expand=True)
        details_widget = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, state=tk.NORMAL, width=70, height=15)
        details_widget.pack(fill="both", expand=True, padx=2, pady=2)
        details_widget.insert(tk.END, details_text)
        details_widget.config(state=tk.DISABLED)
        ok_button = ttk.Button(main_frame, text="确定", command=dialog.destroy)
        ok_button.pack(side="bottom", pady=(15, 0))
        dialog.update_idletasks()
        screen_width, screen_height = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
        window_width, window_height = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
        max_width, max_height = int(screen_width * 0.9), int(screen_height * 0.9)
        if window_width > max_width: window_width = max_width
        if window_height > max_height: window_height = max_height
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        dialog.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        dialog.minsize(min(400, window_width), min(300, window_height))
        self.root.wait_window(dialog)


if __name__ == "__main__":
    try:
        root = tk.Tk()
        # 尝试设置更美观的样式
        style = ttk.Style(root)
        try:
            # 如果系统支持，'clam' 或 'alt' 等主题可能比默认的更好看
            available_themes = style.theme_names()
            if 'vista' in available_themes:
                style.theme_use('vista')
            elif 'clam' in available_themes:
                style.theme_use('clam')
        except tk.TclError:
            pass  # 使用默认主题

        try:
            default_font = ('Microsoft YaHei UI', 10)  # 优先使用雅黑
            root.option_add("*Font", default_font)
        except tk.TclError:
            print("警告：未找到 'Microsoft YaHei UI' 字体，将使用系统默认字体。")
            pass

        app = FileRenameTool(root)
        root.mainloop()
    except Exception as e:
        import traceback

        error_msg = f"程序启动时发生严重错误：\n{e}\n\n{traceback.format_exc()}"
        try:
            error_root = tk.Tk()
            error_root.withdraw()
            messagebox.showerror("致命错误", error_msg)
        except:
            print(error_msg)

        try:
            log_dir = os.path.join(os.path.expanduser("~"), "FileRenameToolLogs")
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, f"startup_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(error_msg)
        except:
            pass