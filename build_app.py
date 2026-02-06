import PyInstaller.__main__
import customtkinter
import os
import sys

# 获取 customtkinter 的库路径
ctk_path = os.path.dirname(customtkinter.__file__)

# 确定分隔符（Windows使用;，macOS/Linux使用:）
sep = ';' if sys.platform.startswith("win") else ':'

# 目标脚本
target_script = "modern_gui.py"

# 打包参数
args = [
    target_script,
    '--name=ExcelBatchConverter',  # 生成的可执行文件名称
    '--noconsole',                 # 不显示控制台窗口（GUI应用）
    '--onefile',                   # 打包成单文件
    '--clean',                     # 清理缓存
    f'--add-data={ctk_path}{sep}customtkinter/', # 添加 customtkinter 的资源文件
]

# 提示开始
print(f"开始打包 {target_script} ...")
print(f"CustomTkinter 路径: {ctk_path}")
print("参数:", args)

# 执行打包
PyInstaller.__main__.run(args)

print("\n打包完成！请在 'dist' 文件夹中查看生成的可执行文件。")
