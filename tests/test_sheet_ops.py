import pytest
from unittest.mock import Mock, patch
import logging
from src.excel_efficiency_toolkit.sheet_ops import list_sheet_names_to_active_sheet

@pytest.fixture
def mock_logger():
    """提供一个 Mock 的 Logger 对象用于测试验证日志输出"""
    return Mock(spec=logging.Logger)

@patch('src.excel_efficiency_toolkit.sheet_ops.get_active_excel')
def test_list_sheet_names_no_excel(mock_get_excel, mock_logger):
    """测试场景：没有活动的 Excel 实例"""
    mock_get_excel.return_value = None
    
    result = list_sheet_names_to_active_sheet(mock_logger)
    
    assert result is False
    mock_logger.error.assert_called_with("操作失败：未检测到正在运行的 Excel。请先打开 Excel。")

@patch('src.excel_efficiency_toolkit.sheet_ops.get_active_excel')
def test_list_sheet_names_no_workbook(mock_get_excel, mock_logger):
    """测试场景：Excel 实例存在，但没有打开的工作簿"""
    mock_excel = Mock()
    mock_excel.ActiveWorkbook = None
    mock_get_excel.return_value = mock_excel

    result = list_sheet_names_to_active_sheet(mock_logger)
    
    assert result is False
    mock_logger.error.assert_called_with("操作失败：没有打开的工作簿。请先打开或新建一个 Excel 文件。")

@patch('src.excel_efficiency_toolkit.sheet_ops.get_active_excel')
def test_list_sheet_names_success(mock_get_excel, mock_logger):
    """测试场景：正常获取并写入工作表名称"""
    # 构建 Mock 对象层级结构
    mock_excel = Mock()
    mock_workbook = Mock()
    mock_workbook.Name = "TestBook.xlsx"
    
    mock_sheet1 = Mock()
    mock_sheet1.Name = "Sheet1"
    mock_sheet2 = Mock()
    mock_sheet2.Name = "Sheet2"
    mock_workbook.Sheets = [mock_sheet1, mock_sheet2]
    
    mock_active_sheet = Mock()
    # 模拟 Cells 返回一个能够设置 Value 属性的对象
    mock_cell = Mock()
    mock_active_sheet.Cells.return_value = mock_cell

    # 组装到 excel 实例
    mock_excel.ActiveWorkbook = mock_workbook
    mock_excel.ActiveSheet = mock_active_sheet
    mock_get_excel.return_value = mock_excel

    # 执行被测函数
    result = list_sheet_names_to_active_sheet(mock_logger)
    
    # 验证逻辑
    assert result is True
    mock_logger.info.assert_any_call("写入完成！请在 Excel 中查看。")
    
    # 验证单元格调用次数 (表头1次 + 2个Sheet名称)
    assert mock_active_sheet.Cells.call_count == 3
    # 验证 Cells 被正确索引调用
    mock_active_sheet.Cells.assert_any_call(1, 1)
    mock_active_sheet.Cells.assert_any_call(2, 1)
    mock_active_sheet.Cells.assert_any_call(3, 1)