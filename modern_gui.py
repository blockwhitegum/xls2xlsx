import customtkinter as ctk
import threading
from tkinter import filedialog, messagebox
import os
import sys

# 导入核心逻辑
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    import batch_converter
except ImportError:
    batch_converter = None

# 设置外观模式和默认颜色主题
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 窗口配置
        self.title("Excel 批量转换器 (xls -> xlsx)")
        self.geometry("700x500")
        
        # 检查依赖
        if batch_converter is None:
            messagebox.showerror("错误", "无法导入 batch_converter 模块")
            self.destroy()
            return

        # 布局配置
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)  # 日志区域自适应

        # 1. 顶部标题与说明
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        self.label_title = ctk.CTkLabel(self.header_frame, text="批量 Excel 格式转换", 
                                      font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=5)
        
        self.label_desc = ctk.CTkLabel(self.header_frame, 
                                     text="将选定文件夹内的所有 .xls 转换为 .xlsx，保留格式/公式/图表，并删除原文件。\n注意：需要本机安装 Microsoft Excel。",
                                     text_color="gray")
        self.label_desc.pack(pady=(0, 10))

        # 2. 文件夹选择区域
        self.folder_frame = ctk.CTkFrame(self)
        self.folder_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.folder_frame.grid_columnconfigure(1, weight=1)

        self.label_folder = ctk.CTkLabel(self.folder_frame, text="目标文件夹:")
        self.label_folder.grid(row=0, column=0, padx=10, pady=10)

        self.entry_folder = ctk.CTkEntry(self.folder_frame, placeholder_text="请选择包含 .xls 文件的文件夹...")
        self.entry_folder.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.btn_browse = ctk.CTkButton(self.folder_frame, text="浏览...", command=self.browse_folder)
        self.btn_browse.grid(row=0, column=2, padx=10, pady=10)

        # 3. 日志区域
        self.textbox_log = ctk.CTkTextbox(self, width=250)
        self.textbox_log.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.log_message("准备就绪。请选择文件夹并点击开始转换。")

        # 4. 底部控制区域
        self.footer_frame = ctk.CTkFrame(self)
        self.footer_frame.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.footer_frame.grid_columnconfigure(0, weight=1)

        self.progressbar = ctk.CTkProgressBar(self.footer_frame)
        self.progressbar.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        self.progressbar.set(0)

        self.btn_start = ctk.CTkButton(self.footer_frame, text="开始转换", command=self.start_conversion_thread,
                                     font=ctk.CTkFont(size=16, weight="bold"), height=40)
        self.btn_start.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.is_running = False

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, folder_path)
            self.log_message(f"已选择文件夹: {folder_path}")

    def log_message(self, msg):
        self.textbox_log.insert("end", msg + "\n")
        self.textbox_log.see("end")

    def update_progress(self, current, total):
        # 在主线程更新进度条
        progress = current / total if total > 0 else 0
        self.progressbar.set(progress)
        self.title(f"正在转换... ({current}/{total})")

    def start_conversion_thread(self):
        if self.is_running:
            return

        folder_path = self.entry_folder.get().strip()
        if not folder_path:
            messagebox.showwarning("提示", "请先选择一个文件夹")
            return
        
        if not os.path.exists(folder_path):
            messagebox.showerror("错误", "文件夹不存在")
            return

        # 锁定 UI
        self.is_running = True
        self.btn_start.configure(state="disabled", text="正在转换中...")
        self.btn_browse.configure(state="disabled")
        self.entry_folder.configure(state="disabled")
        self.progressbar.set(0)
        self.textbox_log.delete("0.0", "end") # 清空日志
        self.log_message("--- 任务开始 ---")

        # 启动线程
        thread = threading.Thread(target=self.run_conversion, args=(folder_path,))
        thread.start()

    def run_conversion(self, folder_path):
        try:
            batch_converter.batch_convert_xls_to_xlsx(
                folder_path,
                progress_callback=lambda c, t: self.after(0, self.update_progress, c, t),
                log_callback=lambda msg: self.after(0, self.log_message, msg)
            )
            self.after(0, lambda: messagebox.showinfo("完成", "转换任务已完成！"))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("错误", f"发生异常:\n{str(e)}"))
        finally:
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.is_running = False
        self.btn_start.configure(state="normal", text="开始转换")
        self.btn_browse.configure(state="normal")
        self.entry_folder.configure(state="normal")
        self.title("Excel 批量转换器 (xls -> xlsx)")
        self.log_message("--- 任务结束 ---")

if __name__ == "__main__":
    app = App()
    app.mainloop()
