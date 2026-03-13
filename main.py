import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import platform
import threading
import re
from datetime import datetime
import pandas as pd
import io

# 只导入 Excel 解密库（纯 Python，M4 完全兼容）
try:
    import msoffcrypto
    from openpyxl import load_workbook
except ImportError:
    msoffcrypto = None
    load_workbook = None

OS_TYPE = platform.system()
if OS_TYPE == 'Darwin':
    CONFIG_FILE = os.path.join(os.path.expanduser("~"), "Documents", "海关工具配置.json")
    DEFAULT_OUTPUT = os.path.join(os.path.expanduser("~"), "Documents", "海关解密输出")
else:
    CONFIG_FILE = "config.json"
    DEFAULT_OUTPUT = "output"

class PortConfigManager:
    def __init__(self):
        self.config_file = CONFIG_FILE
        self.config = self.load_config()
        if OS_TYPE == 'Darwin':
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            os.makedirs(DEFAULT_OUTPUT, exist_ok=True)

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {"ports": {}}
        return {"ports": {}}

    def save_config(self):
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def get_ports(self):
        return list(self.config.get("ports", {}).keys())

    def get_password(self, port):
        return self.config.get("ports", {}).get(port, "")

    def add_port(self, port, password):
        self.config.setdefault("ports", {})[port] = password
        self.save_config()

    def delete_port(self, port):
        if port in self.config.get("ports", {}):
            del self.config["ports"][port]
            self.save_config()

    def batch_add_ports(self, port_list):
        """批量添加口岸，port_list是[(port, password), ...]"""
        for port, pwd in port_list:
            if port and pwd:
                self.config.setdefault("ports", {})[port] = pwd
        self.save_config()

class FileProcessor:
    def __init__(self, password):
        self.password = password

    def process_file(self, input_path, output_dir, houses_to_delete=None):
        filename = os.path.basename(input_path)
        output_path = os.path.join(output_dir, filename)
        ext = os.path.splitext(filename)[1].lower()

        try:
            if ext in ['.xlsx', '.xls', '.xlsm']:
                if not msoffcrypto:
                    return False, "缺少 msoffcrypto，请运行: pip3 install msoffcrypto-tool"
                return self.process_excel(input_path, output_path, houses_to_delete)
            elif ext == '.pdf':
                return False, "本版本暂不支持 PDF（请使用 Excel 文件）"
            else:
                return False, f"不支持的格式: {ext}"
        except Exception as e:
            return False, str(e)

    def process_excel(self, input_path, output_path, houses_to_delete=None):
        try:
            # 解密到内存
            with open(input_path, "rb") as f:
                file = msoffcrypto.OfficeFile(f)
                decrypted = io.BytesIO()
                file.load_key(password=self.password)
                file.decrypt(decrypted)
                decrypted.seek(0)

            ext = os.path.splitext(input_path)[1].lower()

            if ext in ['.xlsx', '.xlsm']:
                # 先保存解密后的文件（保留格式）
                wb = load_workbook(decrypted, data_only=False)
                wb.save(output_path)

                # 如果需要删除 House Numbers
                if houses_to_delete and len(houses_to_delete) > 0:
                    delete_result = self.delete_house_numbers(output_path, houses_to_delete)
                    if not delete_result[0]:
                        return delete_result

                return True, "处理成功（已保留原格式）"
            else:
                # .xls 旧格式
                decrypted.seek(0)
                xl = pd.ExcelFile(decrypted)
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet in xl.sheet_names:
                        df = pd.read_excel(decrypted, sheet_name=sheet, engine='openpyxl')
                        df.to_excel(writer, sheet_name=sheet, index=False)
                return True, "处理成功（警告：.xls 旧格式无法保留样式）"

        except Exception as e:
            return False, f"Excel错误: {str(e)}"

    def delete_house_numbers(self, output_path, houses_to_delete):
        """删除指定的 House Numbers 对应的行"""
        try:
            wb = load_workbook(output_path)
            ws = wb.active  # 默认处理第一个工作表

            # AG列是第33列 (A=1, G=7, 26+7=33)
            ag_col = 33

            # 验证第9行表头
            header_cell = ws.cell(row=9, column=ag_col).value
            if not header_cell or "house" not in str(header_cell).lower():
                return False, f"AG列第9行表头检查失败，当前值为：{header_cell}，应为 House Number 相关字段"

            # 收集要删除的行（从第10行开始检查）
            rows_to_delete = []
            found_houses = set()

            for row_idx in range(10, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=ag_col).value
                if cell_value and str(cell_value).strip() in houses_to_delete:
                    rows_to_delete.append(row_idx)
                    found_houses.add(str(cell_value).strip())

            # 检查是否有未找到的 House Number
            not_found = set(houses_to_delete) - found_houses
            if not_found:
                return False, f"找不到对应的 house number: {', '.join(not_found)}"

            # 从后往前删除（避免行号变化影响）
            if rows_to_delete:
                for row_idx in reversed(sorted(rows_to_delete)):
                    ws.delete_rows(row_idx)
                wb.save(output_path)

            return True, f"成功删除 {len(rows_to_delete)} 行数据"

        except Exception as e:
            return False, f"删除 House Number 时出错: {str(e)}"

class FileListItem:
    """文件列表项，每行一个Frame包含所有控件"""
    def __init__(self, parent_frame, filename, file_path, app_ref):
        self.filename = filename
        self.file_path = file_path
        self.app_ref = app_ref
        self.houses = []

        # 创建行Frame
        self.frame = tk.Frame(parent_frame, relief=tk.RIDGE, borderwidth=1, bg="#f5f5f5")
        self.frame.pack(fill=tk.X, pady=2, padx=5)

        # 文件名标签（左对齐，固定宽度）
        self.lbl_name = tk.Label(self.frame, text=filename, width=35, anchor="w", 
                                bg="#f5f5f5", font=('Arial', 10))
        self.lbl_name.pack(side=tk.LEFT, padx=5)

        # House Numbers 显示（中间，可变宽度）
        self.lbl_houses = tk.Label(self.frame, text="待删除: 无", width=35, anchor="w",
                                  bg="#f5f5f5", fg="#666666", font=('Arial', 9))
        self.lbl_houses.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 状态标签
        self.lbl_status = tk.Label(self.frame, text="待处理", width=10, anchor="center",
                                  bg="#FFF3CD", fg="#856404", font=('Arial', 9, 'bold'))
        self.lbl_status.pack(side=tk.LEFT, padx=5)

        # 编辑按钮
        self.btn_edit = tk.Button(self.frame, text="编辑", width=5, command=self.edit_houses,
                                 relief=tk.FLAT, bg="#e3f2fd", cursor="hand2")
        self.btn_edit.pack(side=tk.LEFT, padx=2)

        # 删除按钮
        self.btn_delete = tk.Button(self.frame, text="删除", width=5, command=self.delete_self,
                                   relief=tk.FLAT, bg="#ffebee", fg="#d32f2f", cursor="hand2")
        self.btn_delete.pack(side=tk.LEFT, padx=2)

    def edit_houses(self):
        """编辑 House Numbers"""
        self.app_ref.edit_houses_for_file(self)

    def delete_self(self):
        """删除自己"""
        msg = "确定从列表中移除？" + "\n" + self.filename
        if messagebox.askyesno("确认删除", msg):
            self.app_ref.remove_file_item(self)

    def update_houses(self, houses):
        """更新 House Numbers 显示"""
        self.houses = houses
        if houses:
            display = ", ".join(houses[:2]) + ("..." if len(houses) > 2 else "")
            self.lbl_houses.config(text="待删除: " + display, fg="#d32f2f")
            self.lbl_status.config(text="已配置", bg="#D4EDDA", fg="#155724")
        else:
            self.lbl_houses.config(text="待删除: 无", fg="#666666")
            self.lbl_status.config(text="待处理", bg="#FFF3CD", fg="#856404")

    def set_status(self, status, success=True):
        """设置处理状态"""
        self.lbl_status.config(text=status)
        if status == "成功":
            self.lbl_status.config(bg="#D4EDDA", fg="#155724")
        elif status == "失败":
            self.lbl_status.config(bg="#F8D7DA", fg="#721C24")
        else:
            self.lbl_status.config(bg="#FFF3CD", fg="#856404")

    def destroy(self):
        """销毁控件"""
        self.frame.destroy()

class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("海关文件解密工具 v2.4 (密码显示隐藏版)")
        self.root.geometry("900x700")

        if OS_TYPE == 'Darwin':
            try:
                from tkinter import font
                font.nametofont("TkDefaultFont").configure(size=12)
            except:
                pass

        self.config_mgr = PortConfigManager()
        self.file_items = []  # FileListItem 对象列表
        self.port_entries = []  # 配置窗口中的口岸行引用
        self.setup_ui()
        self.refresh_ports()

        if not self.config_mgr.get_ports():
            self.show_first_time_help()

    def show_first_time_help(self):
        help_text = """欢迎使用海关文件解密工具 v2.4！

🆕 v2.4 新特性：
• 每行文件直接显示删除按钮，一键移除
• 批量添加口岸密码（支持复制粘贴）
• 支持按 House Number 精准删除行
• 密码配置支持显示/隐藏切换

操作步骤：
1. 点击【配置口岸密码】批量添加口岸，支持显示/隐藏密码
2. 添加 Excel 文件，每行显示：文件名 | 待删除House | 状态 | 编辑 | 删除
3. 点击 编辑 配置 House Numbers，点击 删除 直接删除该行
4. 点击【开始处理】

⚠️ 注意：若 House Number 找不到，该文件会标记为处理失败"""
        messagebox.showinfo("使用指南 v2.4", help_text)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 口岸选择
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=10)

        ttk.Label(select_frame, text="选择口岸：", font=('Arial', 12)).pack(side=tk.LEFT)
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(select_frame, textvariable=self.port_var, state="readonly", width=30, font=('Arial', 12))
        self.port_combo.pack(side=tk.LEFT, padx=10)
        ttk.Button(select_frame, text="配置口岸密码", command=self.open_config).pack(side=tk.LEFT)

        # 文件列表区域（使用Canvas实现滚动）
        ttk.Label(main_frame, text="待处理文件列表（每行独立操作）：", font=('Arial', 11)).pack(anchor=tk.W, pady=5)

        # 创建带滚动条的Frame容器
        list_container = ttk.Frame(main_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=5)

        # Canvas + Scrollbar
        self.canvas = tk.Canvas(list_container, bg="white", highlightthickness=1, highlightbackground="#ddd")
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.canvas.yview)

        self.file_frame = tk.Frame(self.canvas, bg="white")
        self.file_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas.create_window((0, 0), window=self.file_frame, anchor="nw", width=850)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮支持
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="添加 Excel 文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="清空所有", command=self.clear_all_files).pack(side=tk.LEFT, padx=5)

        # 输出目录
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=10)
        ttk.Label(out_frame, text="输出目录：").pack(side=tk.LEFT)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        ttk.Entry(out_frame, textvariable=self.output_var, font=('Arial', 11)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(out_frame, text="浏览", command=self.browse_output).pack(side=tk.LEFT)

        # 处理按钮 - 高级蓝色
        process_btn = tk.Button(main_frame, text="开始处理", command=self.start_process,
                               bg="#1565C0", fg="white", font=('Arial', 14, 'bold'), height=2,
                               activebackground="#0D47A1", activeforeground="white",
                               cursor="hand2")
        process_btn.pack(fill=tk.X, pady=15)

        # 日志
        ttk.Label(main_frame, text="处理日志：").pack(anchor=tk.W)
        self.log_area = scrolledtext.ScrolledText(main_frame, height=10, font=('Monaco', 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=5)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪 - v2.4 密码显示隐藏")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)

    def refresh_ports(self):
        ports = self.config_mgr.get_ports()
        self.port_combo['values'] = ports
        if ports:
            self.port_var.set(ports[0])
            self.status_var.set(f"就绪 - 当前口岸: {ports[0]}")

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择加密的 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls *.xlsm"),
                ("Excel 2007+", "*.xlsx"),
                ("Excel 宏文件", "*.xlsm"),
                ("Excel 97-2003", "*.xls")
            ]
        )
        for f in files:
            filename = os.path.basename(f)
            item = FileListItem(self.file_frame, filename, f, self)
            self.file_items.append(item)

        if files:
            self.log(f"已添加 {len(files)} 个文件")
            # 更新 Canvas 滚动区域
            self.file_frame.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def edit_houses_for_file(self, file_item):
        """为指定文件编辑 House Numbers"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑删除项 - {file_item.filename}")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="输入要删除的 House Numbers：", font=('Arial', 11, 'bold')).pack(pady=10)
        ttk.Label(dialog, text="（多个号码可用逗号、空格或换行分隔）", foreground="gray").pack()

        # 文本输入框
        text_frame = ttk.Frame(dialog)
        text_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        text_input = tk.Text(text_frame, width=40, height=10, font=('Arial', 11))
        text_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(text_frame, command=text_input.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_input.config(yscrollcommand=scrollbar.set)

        # 填入现有值
        if file_item.houses:
            text_input.insert("1.0", "\n".join(file_item.houses))

        # 按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        def save_houses():
            content = text_input.get("1.0", tk.END).strip()
            # 支持逗号、空格、换行分隔
            houses = [h.strip() for h in re.split(r'[,\s]+', content) if h.strip()]
            file_item.update_houses(houses)
            self.log(f"【{file_item.filename}】已配置 {len(houses)} 个待删除 House Numbers")
            dialog.destroy()

        ttk.Button(btn_frame, text="保存", command=save_houses).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def remove_file_item(self, file_item):
        """从列表移除指定项"""
        file_item.destroy()
        self.file_items.remove(file_item)
        self.log(f"已移除: {file_item.filename}")
        # 更新滚动区域
        self.file_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def clear_all_files(self):
        """清空所有文件"""
        if not self.file_items:
            return
        count = len(self.file_items)
        msg = f"确定清空所有 {count} 个文件？"
        if messagebox.askyesno("确认", msg):
            for item in self.file_items:
                item.destroy()
            self.file_items.clear()
            self.log("已清空所有文件")
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def browse_output(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_var.set(dir_path)

    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_area.see(tk.END)
        self.root.update()

    def start_process(self):
        thread = threading.Thread(target=self.process)
        thread.daemon = True
        thread.start()

    def process(self):
        port = self.port_var.get()
        if not port:
            messagebox.showerror("错误", "请先配置口岸密码")
            return

        password = self.config_mgr.get_password(port)
        if not password:
            messagebox.showerror("错误", f"口岸 [{port}] 未设置密码")
            return

        if not self.file_items:
            messagebox.showerror("错误", "请先添加文件")
            return

        output_dir = self.output_var.get()
        os.makedirs(output_dir, exist_ok=True)

        self.log(f"\n开始处理 {len(self.file_items)} 个文件...")
        processor = FileProcessor(password)
        success = failed = 0

        for i, item in enumerate(self.file_items, 1):
            houses = item.houses
            house_info = f"[删除{len(houses)}个House]" if houses else "[仅解密]"
            self.log(f"[{i}/{len(self.file_items)}] 处理: {item.filename} {house_info}")

            result, msg = processor.process_file(item.file_path, output_dir, houses)
            if result:
                self.log(f"  成功: {msg}")
                item.set_status("成功", True)
                success += 1
            else:
                self.log(f"  失败: {msg}")
                item.set_status("失败", False)
                failed += 1

        self.log(f"\n完成！成功: {success}, 失败: {failed}")
        self.status_var.set(f"处理完成 - 成功:{success} 失败:{failed}")

        if success > 0 and messagebox.askyesno("完成", f"成功处理 {success} 个文件！\n\n打开输出文件夹？"):
            if OS_TYPE == 'Darwin':
                os.system(f'open "{output_dir}"')
            elif OS_TYPE == 'Windows':
                os.system(f'start "" "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')

    def open_config(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("配置口岸密码")
        dialog.geometry("550x550")
        dialog.transient(self.root)

        # 批量添加区域
        ttk.Label(dialog, text="批量添加（格式：口岸名 密码，每行一个）：", 
                 font=('Arial', 11, 'bold')).pack(pady=(10,5))

        batch_frame = ttk.Frame(dialog)
        batch_frame.pack(padx=10, pady=5, fill=tk.X)

        batch_text = tk.Text(batch_frame, width=50, height=6, font=('Arial', 11))
        batch_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(batch_frame, command=batch_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        batch_text.config(yscrollcommand=scrollbar.set)

        ttk.Label(dialog, text="示例：上海港 password123", foreground="gray", justify=tk.LEFT).pack(anchor=tk.W, padx=10)

        # 保存按钮（单独一行，显眼）
        save_btn = ttk.Button(dialog, text="保存配置", width=15)
        save_btn.pack(pady=10)

        # 现有配置列表（使用 Frame 代替 Treeview，支持显示/隐藏）
        ttk.Label(dialog, text="现有配置（点击显示/隐藏按钮切换密码可见性）：", 
                 font=('Arial', 11, 'bold')).pack(pady=(10,5), anchor=tk.W, padx=10)

        # 创建带滚动条的容器
        list_container = ttk.Frame(dialog)
        list_container.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_container, bg="white", highlightthickness=1, highlightbackground="#ddd")
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)

        port_list_frame = tk.Frame(canvas, bg="white")
        port_list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=port_list_frame, anchor="nw", width=480)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 配置保存按钮的命令（需要在 port_list_frame 创建后）
        save_btn.config(command=lambda: self.save_batch_config(batch_text, port_list_frame))

        # 加载现有数据
        self.port_entries = []  # 存储每行的引用
        for p in self.config_mgr.get_ports():
            self.add_port_row(port_list_frame, p, self.config_mgr.get_password(p))

        # 关闭按钮
        ttk.Button(dialog, text="关闭", command=dialog.destroy, width=12).pack(pady=10)

    def add_port_row(self, parent, port, password):
        """添加一行口岸配置（带显示/隐藏密码功能）"""
        row = tk.Frame(parent, relief=tk.RIDGE, borderwidth=1, bg="#f8f9fa")
        row.pack(fill=tk.X, pady=2, padx=5)

        # 口岸名
        lbl_port = tk.Label(row, text=port, width=15, anchor="w", bg="#f8f9fa", font=('Arial', 10, 'bold'))
        lbl_port.pack(side=tk.LEFT, padx=5)

        # 密码显示区域（使用 Entry 可以方便切换 show 属性）
        pwd_var = tk.StringVar(value=password)
        entry_pwd = tk.Entry(row, textvariable=pwd_var, width=25, show="*", font=('Arial', 10))
        entry_pwd.pack(side=tk.LEFT, padx=5)

        # 显示/隐藏按钮 - 使用文字避免emoji问题
        is_visible = [False]  # 使用列表包装以便在内部函数修改

        def toggle_visibility():
            if is_visible[0]:
                entry_pwd.config(show="*")
                btn_toggle.config(text="显示", bg="#e3f2fd")
                is_visible[0] = False
            else:
                entry_pwd.config(show="")
                btn_toggle.config(text="隐藏", bg="#fff3e0")
                is_visible[0] = True

        btn_toggle = tk.Button(row, text="显示", width=5, command=toggle_visibility,
                              relief=tk.FLAT, bg="#e3f2fd", cursor="hand2", font=('Arial', 9))
        btn_toggle.pack(side=tk.LEFT, padx=2)

        # 删除按钮
        def delete_this():
            if messagebox.askyesno("确认", f"确定删除口岸 [{port}] 的配置？"):
                self.config_mgr.delete_port(port)
                row.destroy()
                # 从列表中移除引用
                self.port_entries = [e for e in self.port_entries if e['port'] != port]
                self.refresh_ports()

        btn_delete = tk.Button(row, text="删除", width=5, command=delete_this,
                              relief=tk.FLAT, bg="#ffebee", fg="#d32f2f", cursor="hand2")
        btn_delete.pack(side=tk.LEFT, padx=5)

        # 保存引用
        self.port_entries.append({
            'port': port,
            'row': row,
            'pwd_var': pwd_var,
            'entry': entry_pwd
        })

    def save_batch_config(self, batch_text, port_list_frame):
        """保存批量添加的配置"""
        content = batch_text.get("1.0", tk.END).strip()
        added_count = 0
        if content:
            lines = content.splitlines()
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                parts = line.split(maxsplit=1)
                if len(parts) == 2:
                    port_name, port_pwd = parts
                    if port_name and port_pwd:
                        is_new = port_name not in self.config_mgr.get_ports()
                        self.config_mgr.add_port(port_name, port_pwd)
                        if is_new:
                            self.add_port_row(port_list_frame, port_name, port_pwd)
                        else:
                            # 更新现有显示
                            for entry in self.port_entries:
                                if entry['port'] == port_name:
                                    entry['pwd_var'].set(port_pwd)
                                    break
                        added_count += 1

        batch_text.delete("1.0", tk.END)
        self.refresh_ports()

        if added_count > 0:
            messagebox.showinfo("成功", f"已保存 {added_count} 个口岸配置！")
        else:
            messagebox.showinfo("提示", "没有新的配置需要保存")

if __name__ == "__main__":
    if not msoffcrypto:
        print("首次使用，请先安装依赖：")
        print("pip3 install msoffcrypto-tool pandas openpyxl")
        exit()

    root = tk.Tk()
    app = CustomsApp(root)
    root.mainloop()
