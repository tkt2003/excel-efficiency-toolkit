import win32com.client
import pythoncom

def get_active_excel():
    """
    获取当前活动的 Excel 实例。
    返回 win32com Excel.Application 对象，如果未运行则返回 None。
    """
    try:
        # 初始化 COM 环境
        pythoncom.CoInitialize()
        # 尝试获取当前运行的 Excel 实例
        excel = win32com.client.GetActiveObject("Excel.Application")
        return excel
    except Exception:
        return None