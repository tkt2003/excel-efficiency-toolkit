import logging
import tkinter as tk

def setup_logger(gui_text_widget=None):
    """
    配置并返回 logger。
    如果提供了 gui_text_widget，则将日志同时输出到该 Tkinter Text 控件中。
    """
    logger = logging.getLogger("ExcelToolkit")
    logger.setLevel(logging.INFO)
    
    # 避免重复添加 Handler
    if logger.hasHandlers():
        logger.handlers.clear()

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # 控制台输出 Handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # GUI 文本框输出 Handler
    if gui_text_widget:
        class GUIHandler(logging.Handler):
            def __init__(self, widget):
                super().__init__()
                self.widget = widget

            def emit(self, record):
                msg = self.format(record)
                def append_text():
                    self.widget.configure(state='normal')
                    self.widget.insert(tk.END, msg + '\n')
                    self.widget.see(tk.END)
                    self.widget.configure(state='disabled')
                # 确保在主线程执行 UI 更新
                self.widget.after(0, append_text)
                
        gui_handler = GUIHandler(gui_text_widget)
        gui_handler.setFormatter(formatter)
        logger.addHandler(gui_handler)

    return logger