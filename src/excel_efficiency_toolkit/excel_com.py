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


def connect_active_excel():
    """连接正在运行的 Excel 实例；未运行时抛出带中文提示的 RuntimeError。

    不负责 COM 会话初始化，调用方需已 CoInitialize（或使用 run_in_excel_com_session）。
    延迟导入以便测试注入 fake win32com。
    """
    try:
        import win32com.client as win32com_client
    except Exception as e:
        raise RuntimeError("无法加载 Excel COM 组件，请确认已安装 pywin32 并在 Windows + Excel 环境运行。") from e

    try:
        return win32com_client.GetActiveObject("Excel.Application")
    except Exception as e:
        raise RuntimeError("未检测到正在运行的 Excel，请先打开目标/合并工作簿后重试。") from e


def run_in_excel_com_session(runner):
    """在一次 COM 会话中连接活动 Excel 并执行 runner(excel)，返回其结果。"""
    try:
        import pythoncom as pythoncom_module
    except Exception as e:
        raise RuntimeError("无法加载 Excel COM 组件，请确认已安装 pywin32 并在 Windows + Excel 环境运行。") from e

    pythoncom_module.CoInitialize()
    try:
        excel = connect_active_excel()
        return runner(excel)
    finally:
        pythoncom_module.CoUninitialize()
