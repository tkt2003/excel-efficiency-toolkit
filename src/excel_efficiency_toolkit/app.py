import tkinter as tk
from tkinter import scrolledtext
from .logging_utils import setup_logger
from .sheet_ops import list_sheet_names_to_active_sheet

class ExcelToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 效率工具台")
        self.root.geometry("600x400")

        # 顶部按钮区域
        self.frame_top = tk.Frame(root)
        self.frame_top.pack(pady=20, fill=tk.X)

        self.btn_list_sheets = tk.Button(
            self.frame_top,
            text="列出当前工作簿所有工作表",
            font=("Microsoft YaHei", 12),
            command=self.run_list_sheets,
            bg="#f0f0f0"
        )
        self.btn_list_sheets.pack()

        # 底部日志区域
        self.frame_bottom = tk.Frame(root)
        self.frame_bottom.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        # 日志标签
        tk.Label(self.frame_bottom, text="运行日志：").pack(anchor="w")

        # 日志输出文本框
        self.log_text = scrolledtext.ScrolledText(
            self.frame_bottom, 
            state='disabled', 
            height=15, 
            font=("Consolas", 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 初始化自定义 logger
        self.logger = setup_logger(self.log_text)
        self.logger.info("欢迎使用 Excel 效率工具台。程序已就绪。")

    def run_list_sheets(self):
        """按钮回调函数，执行获取工作表的操作"""
        self.btn_list_sheets.config(state="disabled")
        try:
            list_sheet_names_to_active_sheet(self.logger)
        finally:
            self.btn_list_sheets.config(state="normal")

def main():
    root = tk.Tk()
    app = ExcelToolkitApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()