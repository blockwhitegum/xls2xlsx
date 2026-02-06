import os
import time
from pathlib import Path
import xlwings as xw

class ConversionError(Exception):
    pass

def batch_convert_xls_to_xlsx(folder_path, progress_callback=None, log_callback=None):
    """
    批量将文件夹内的 xls 文件转换为 xlsx，并删除源文件。
    保留格式、公式、图表等（依赖本地 Excel）。
    
    :param folder_path: 目标文件夹路径
    :param progress_callback: 进度回调函数，接收 (current, total)
    :param log_callback: 日志回调函数，接收 (message)
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ConversionError("文件夹不存在或不是目录")

    # 查找所有 xls 文件（不递归，仅当前目录，除非用户需要递归？用户说“文件夹里面的”，通常指一层，也可以做递归。这里暂定一层以简化，或者询问。为了稳妥，做成非递归，简单明了）
    # 注意排除临时文件 (~$ 开头的)
    xls_files = [f for f in folder.glob("*.xls") if not f.name.startswith("~$")]
    
    total = len(xls_files)
    if total == 0:
        if log_callback:
            log_callback("未找到 .xls 文件")
        return

    if log_callback:
        log_callback(f"找到 {total} 个 .xls 文件，准备开始转换...")

    # 启动 Excel 实例（隐藏模式）
    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        for i, xls_file in enumerate(xls_files):
            try:
                xlsx_file = xls_file.with_suffix(".xlsx")
                
                if log_callback:
                    log_callback(f"正在转换: {xls_file.name} ...")
                
                # 打开工作簿
                book = app.books.open(str(xls_file))
                
                # 另存为 xlsx
                # FileFormat 51 = xlOpenXMLWorkbook (xlsx)
                book.save(str(xlsx_file))
                book.close()
                
                # 删除源文件
                os.remove(xls_file)
                
                if log_callback:
                    log_callback(f"成功: {xls_file.name} -> {xlsx_file.name} (源文件已删除)")
                
            except Exception as e:
                if log_callback:
                    log_callback(f"失败: {xls_file.name} - {str(e)}")
            
            # 更新进度
            if progress_callback:
                progress_callback(i + 1, total)
                
    except Exception as e:
        if log_callback:
            log_callback(f"Excel 进程错误: {str(e)}")
        raise e
    finally:
        if app:
            try:
                app.quit()
            except:
                pass
        if log_callback:
            log_callback("所有任务完成。")
