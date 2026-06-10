import pytest
from src.excel_efficiency_toolkit.name_utils import (
    clean_sheet_name,
    get_safe_sheet_name,
    get_unique_sheet_name
)

def test_clean_sheet_name():
    # 1. 非法字符清洗
    assert clean_sheet_name("a\\b/c?d*e[f]g:h") == "abcdefgh"
    
    # 2. None 和空字符串
    assert clean_sheet_name(None) == ""
    assert clean_sheet_name("") == ""
    assert clean_sheet_name("   ") == ""
    
    # 3. 31 字符截断
    long_name = "a" * 40
    assert clean_sheet_name(long_name) == "a" * 31
    assert len(clean_sheet_name(long_name)) == 31

def test_get_safe_sheet_name():
    # 4. fallback
    assert get_safe_sheet_name(None) == "Sheet"
    assert get_safe_sheet_name("") == "Sheet"
    assert get_safe_sheet_name("   ", fallback="Backup") == "Backup"
    
    # 5. fallback 也非法时回退到 Sheet
    assert get_safe_sheet_name(None, fallback="[Invalid]") == "Invalid"
    assert get_safe_sheet_name(None, fallback="*?/") == "Sheet"

    # 31 字符截断
    long_fallback = "b" * 40
    assert len(get_safe_sheet_name(None, fallback=long_fallback)) == 31

def test_get_unique_sheet_name():
    existing_names = set()
    
    # 不重复的直接返回
    name1 = get_unique_sheet_name("Data", existing_names)
    assert name1 == "Data"
    assert "Data" in existing_names
    
    # 6. 重名追加 _2、_3
    name2 = get_unique_sheet_name("Data", existing_names)
    assert name2 == "Data_2"
    assert "Data_2" in existing_names
    
    name3 = get_unique_sheet_name("Data", existing_names)
    assert name3 == "Data_3"
    assert "Data_3" in existing_names
    
    # 7. 大小写重名
    name4 = get_unique_sheet_name("data", existing_names)
    assert name4 == "data_4"
    assert "data_4" in existing_names
    
    # 8. 追加序号后仍不超过 31 字符
    long_base = "a" * 31
    name5 = get_unique_sheet_name(long_base, existing_names)
    assert name5 == "a" * 31
    assert name5 in existing_names
    
    name6 = get_unique_sheet_name(long_base, existing_names)
    # 应当为 "a" * 29 + "_2"，总长 31
    assert name6 == "a" * 29 + "_2"
    assert len(name6) == 31
    assert name6 in existing_names

    name7 = get_unique_sheet_name(long_base, existing_names)
    assert name7 == "a" * 29 + "_3"
    assert len(name7) == 31
    assert name7 in existing_names
    
    # 超过 9 后，后缀变为 3 字符，例如 "_10"
    for _ in range(4, 10):
        get_unique_sheet_name(long_base, existing_names)
        
    name10 = get_unique_sheet_name(long_base, existing_names)
    assert name10 == "a" * 28 + "_10"
    assert len(name10) == 31
    assert name10 in existing_names
