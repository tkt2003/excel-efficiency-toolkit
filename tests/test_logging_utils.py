import logging

from src.excel_efficiency_toolkit.logging_utils import setup_logger


def test_logger_timestamp_is_precise_to_seconds_without_milliseconds():
    logger = setup_logger()
    record = logging.LogRecord("ExcelToolkit", logging.INFO, "", 0, "任务已就绪", (), None)
    record.created = 1783916907.661
    record.msecs = 661

    message = logger.handlers[0].formatter.format(record)

    assert message.endswith(" - INFO - 任务已就绪")
    assert ",661" not in message
    assert message.count(":") == 2
