import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import platform
import threading
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

class FileProcessor:
    def __init__(self, password):
        self.password = password

    def process_file(self, input_path, output_dir):
        filename = os.path.basename(input_path)
        output_path = os.path.join(output_dir, filename)
        ext = os.path.splitext(filename)[1].lower()

        try:
            if ext in ['.xlsx', '.xls', '.xlsm']:
                if not msoffcrypto:
                    return False, "缺少 msoffcrypto，请运行: pip3 install msoffcrypto-tool"
                return self.process_excel(input_path, output_path)
            elif ext == '.pdf':
                return False, "本版本暂不支持 PDF（请使用 Excel 文件）"
            else:
                return False, f"不支持的格式: {ext}"
        except Exception as e:
            return False, str(e)

    def process_excel(self, input_path, output_path):
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
                # 使用 openpyxl 保留完整格式（包括列宽、行高、合并单元格、样式、公式等）
                wb = load_workbook(decrypted, data_only=False)  # data_only=False 保留公式而非计算值
                wb.save(output_path)

                # 验证文件完整性
                if self.verify_excel_openpyxl(output_path):
                    return True, "处理成功（已保留原格式：列宽、行高、合并单元格、样式等）"
                else:
                    os.remove(output_path)
                    return False, "文件校验失败"
            else:
                # .xls 旧格式使用 pandas 方式（格式会丢失，但内容保留）
                decrypted.seek(0)
                xl = pd.ExcelFile(decrypted)
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet in xl.sheet_names:
                        df = pd.read_excel(decrypted, sheet_name=sheet, engine='openpyxl')
                        df.to_excel(writer, sheet_name=sheet, index=False)
                return True, "处理成功（警告：.xls 旧格式无法保留样式，仅保留数据内容）"

        except Exception as e:
            return False, f"Excel错误: {str(e)}"

    def verify_excel_openpyxl(self, output_path):
        """验证解密后的文件是否能正常打开"""
        try:
            wb = load_workbook(output_path)
            # 尝试读取每个工作表的一个单元格，确保文件未损坏
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if ws.max_row > 0 and ws.max_column > 0:
                    _ = ws.cell(1, 1).value
            return True
        except Exception as e:
            print(f"验证失败: {e}")
            return False

    def verify_excel(self, original, output_path):
        """旧版验证方法（使用 pandas 对比数据内容）"""
        try:
            xl_orig = pd.ExcelFile(original)
            xl_new = pd.ExcelFile(output_path)
            if xl_orig.sheet_names != xl_new.sheet_names:
                return False
            for sheet in xl_orig.sheet_names:
                df_orig = pd.read_excel(xl_orig, sheet_name=sheet, engine='openpyxl')
                df_new = pd.read_excel(xl_new, sheet_name=sheet, engine='openpyxl')
                if not df_orig.equals(df_new):
                    return False
            return True
        except:
            return False

class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("海关文件解密工具 v1.2 (格式保留版)")
        self.root.geometry("800x600")

        if OS_TYPE == 'Darwin':
            try:
                from tkinter import font
                font.nametofont("TkDefaultFont").configure(size=12)
            except:
                pass

        self.config_mgr = PortConfigManager()
        self.setup_ui()
        self.refresh_ports()

        if not self.config_mgr.get_ports():
            self.show_first_time_help()

    def show_first_time_help(self):
        help_text = """欢迎使用海关文件解密工具（格式保留版）！

✨ 新版本特性：
• 完美保留原文件格式（列宽、行高、合并单元格、字体颜色、边框等）
• 保留 Excel 公式（不是只保留计算结果）
• 支持 .xlsx 和 .xlsm 格式

首次使用步骤：
1. 点击【配置口岸密码】添加口岸和密码
2. 选择要解密的 Excel 文件
3. 点击【开始处理】

注意：本版本专门优化支持 Apple M4 芯片
处理后的文件保存在 Documents/海关解密输出/"""
        messagebox.showinfo("使用指南", help_text)

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
        ttk.Button(select_frame, text="⚙️ 配置口岸密码", command=self.open_config).pack(side=tk.LEFT)

        # 文件列表
        ttk.Label(main_frame, text="待处理文件（支持 .xlsx / .xlsm / .xls）：", font=('Arial', 11)).pack(anchor=tk.W, pady=5)

        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.file_list = tk.Listbox(list_frame, height=6, font=('Arial', 11))
        self.file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, command=self.file_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_list.config(yscrollcommand=scrollbar.set)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="➕ 添加 Excel 文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 清空", command=self.clear_files).pack(side=tk.LEFT)

        # 输出目录
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=10)
        ttk.Label(out_frame, text="输出目录：").pack(side=tk.LEFT)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        ttk.Entry(out_frame, textvariable=self.output_var, font=('Arial', 11)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(out_frame, text="浏览", command=self.browse_output).pack(side=tk.LEFT)

        # 处理按钮
        process_btn = tk.Button(main_frame, text="🚀 开始处理", command=self.start_process,
                               bg="#4CAF50", fg="white", font=('Arial', 14, 'bold'), height=2)
        process_btn.pack(fill=tk.X, pady=15)

        # 日志
        ttk.Label(main_frame, text="处理日志（格式保留版）：").pack(anchor=tk.W)
        self.log_area = scrolledtext.ScrolledText(main_frame, height=10, font=('Monaco', 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=5)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪 - 格式保留模式")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)

    def refresh_ports(self):
        ports = self.config_mgr.get_ports()
        self.port_combo['values'] = ports
        if ports:
            self.port_var.set(ports[0])
            self.status_var.set(f"就绪 - 当前口岸: {ports[0]} (保留格式模式)")

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
            self.file_list.insert(tk.END, f)
        if files:
            self.log(f"已添加 {len(files)} 个文件（将保留原格式）")

    def clear_files(self):
        self.file_list.delete(0, tk.END)

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

        files = list(self.file_list.get(0, tk.END))
        if not files:
            messagebox.showerror("错误", "请先添加文件")
            return

        output_dir = self.output_var.get()
        os.makedirs(output_dir, exist_ok=True)

        self.log(f"\n🚀 开始处理 {len(files)} 个文件（格式保留模式）...")
        self.log("💡 提示：.xlsx/.xlsm 将保留完整格式，.xls 旧格式仅保留数据")
        processor = FileProcessor(password)
        success = failed = 0

        for i, file_path in enumerate(files, 1):
            filename = os.path.basename(file_path)
            self.log(f"[{i}/{len(files)}] 处理: {filename}")

            result, msg = processor.process_file(file_path, output_dir)
            if result:
                self.log(f"  ✅ {msg}")
                success += 1
            else:
                self.log(f"  ❌ {msg}")
                failed += 1

        self.log(f"\n🎉 完成！成功: {success}, 失败: {failed}")
        self.status_var.set(f"处理完成 - 成功:{success}")

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
        dialog.geometry("400x300")
        dialog.transient(self.root)

        tree = ttk.Treeview(dialog, columns=("口岸", "密码"), show="headings")
        tree.heading("口岸", text="口岸")
        tree.heading("密码", text="密码")
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for p in self.config_mgr.get_ports():
            tree.insert("", tk.END, values=(p, "*" * len(self.config_mgr.get_password(p))))

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)

        def add():
            win = tk.Toplevel(dialog)
            win.title("添加口岸")
            win.geometry("300x150")

            ttk.Label(win, text="口岸名称：").pack(pady=5)
            port_entry = ttk.Entry(win)
            port_entry.pack()

            ttk.Label(win, text="密码：").pack(pady=5)
            pwd_entry = ttk.Entry(win, show="*")
            pwd_entry.pack()

            def save():
                if port_entry.get() and pwd_entry.get():
                    self.config_mgr.add_port(port_entry.get(), pwd_entry.get())
                    tree.insert("", tk.END, values=(port_entry.get(), "*" * len(pwd_entry.get())))
                    self.refresh_ports()
                    win.destroy()

            tk.Button(win, text="保存", command=save).pack(pady=10)

        def edit():
            sel = tree.selection()
            if not sel:
                return
            port = tree.item(sel[0])["values"][0]

            win = tk.Toplevel(dialog)
            win.title(f"修改 {port}")
            win.geometry("300x120")

            ttk.Label(win, text="新密码：").pack(pady=5)
            pwd_entry = ttk.Entry(win, show="*")
            pwd_entry.pack()

            def update():
                if pwd_entry.get():
                    self.config_mgr.add_port(port, pwd_entry.get())
                    self.log(f"更新 {port} 密码")
                    win.destroy()

            tk.Button(win, text="更新", command=update).pack(pady=5)

        def delete():
            sel = tree.selection()
            if sel and messagebox.askyesno("确认", "确定删除？"):
                port = tree.item(sel[0])["values"][0]
                self.config_mgr.delete_port(port)
                tree.delete(sel[0])
                self.refresh_ports()

        ttk.Button(btn_frame, text="添加", command=add).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="修改", command=edit).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除", command=delete).pack(side=tk.LEFT, padx=5)
        ttk.Button(dialog, text="关闭", command=dialog.destroy).pack(pady=10)

if __name__ == "__main__":
    if not msoffcrypto:
        print("首次使用，请先安装依赖：")
        print("pip3 install msoffcrypto-tool pandas openpyxl")
        exit()

    root = tk.Tk()
    app = CustomsApp(root)
    root.mainloop()
