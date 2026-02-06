#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import sys
import os

# 尝试导入转换器核心逻辑
# 确保当前目录在 sys.path 中，以便能导入同目录下的模块
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    import xls_xlsx_converter
except ImportError:
    # 如果作为独立脚本运行且找不到转换器模块，提示错误
    # 这里主要防止用户只拷贝了 gui 脚本而没有拷贝核心脚本
    xls_xlsx_converter = None

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 格式转换器")
        self.root.geometry("500x220")
        self.root.resizable(False, False)

        if xls_xlsx_converter is None:
            messagebox.showerror("错误", "未找到 xls_xlsx_converter.py 模块，请确保它与本脚本在同一目录下。")
            root.destroy()
            return

        # 变量
        self.input_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="准备就绪")
        self.overwrite_var = tk.BooleanVar(value=False)

        # 构建界面
        self.create_widgets()

    def create_widgets(self):
        # 容器 padding
        padding_opts = {'padx': 10, 'pady': 5}

        # 1. 文件选择区域
        file_frame = tk.LabelFrame(self.root, text="文件选择", padx=10, pady=10)
        file_frame.pack(fill="x", **padding_opts)

        tk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky="w")
        
        entry = tk.Entry(file_frame, textvariable=self.input_path_var, width=40)
        entry.grid(row=0, column=1, padx=5)
        
        btn_browse = tk.Button(file_frame, text="浏览...", command=self.browse_file)
        btn_browse.grid(row=0, column=2)

        # 2. 选项区域
        opt_frame = tk.Frame(self.root)
        opt_frame.pack(fill="x", **padding_opts)
        
        tk.Checkbutton(opt_frame, text="覆盖已存在的文件", variable=self.overwrite_var).pack(side="left")

        # 3. 操作区域
        action_frame = tk.Frame(self.root, pady=10)
        action_frame.pack(fill="x", padx=10)

        self.btn_convert = tk.Button(action_frame, text="开始转换", command=self.run_conversion, 
                                   bg="#007bff", fg="black", font=("Arial", 12, "bold"), height=2)
        self.btn_convert.pack(fill="x")

        # 4. 状态栏
        status_label = tk.Label(self.root, textvariable=self.status_var, fg="gray", anchor="w")
        status_label.pack(side="bottom", fill="x", padx=10, pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_path_var.set(file_path)
            self.status_var.set("已选择文件，点击开始转换")

    def run_conversion(self):
        input_str = self.input_path_var.get().strip()
        if not input_str:
            messagebox.showwarning("提示", "请先选择一个输入文件")
            return

        input_path = Path(input_str)
        if not input_path.exists():
            messagebox.showerror("错误", "输入文件不存在")
            return

        # 禁用按钮防止重复点击
        self.btn_convert.config(state="disabled")
        self.status_var.set("正在转换中...")
        self.root.update() # 强制刷新界面

        try:
            # 推断输出路径
            output_path = xls_xlsx_converter.derive_output_path(input_path, None, None)
            
            # 检查输出是否存在
            if output_path.exists() and not self.overwrite_var.get():
                # 询问是否覆盖
                if not messagebox.askyesno("文件已存在", f"输出文件已存在:\n{output_path}\n\n是否覆盖？"):
                    self.status_var.set("已取消")
                    self.btn_convert.config(state="normal")
                    return

            # 执行转换
            xls_xlsx_converter.convert_file(input_path, output_path)
            
            self.status_var.set(f"成功: 已保存为 {output_path.name}")
            messagebox.showinfo("成功", f"转换完成！\n保存在: {output_path}")

        except RuntimeError as re:
            # 依赖缺失等运行时错误
            messagebox.showerror("环境错误", str(re))
            self.status_var.set("错误: 缺少依赖")
        except Exception as e:
            messagebox.showerror("转换失败", f"发生错误:\n{e}")
            self.status_var.set("转换失败")
        finally:
            self.btn_convert.config(state="normal")

def main():
    root = tk.Tk()
    # 尝试设置图标（如果有的话），这里略过
    app = ConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
