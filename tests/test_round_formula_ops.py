from src.excel_efficiency_toolkit.round_formula_ops import (
    build_round_formula_from_formula,
    build_round_formula_from_value,
    is_already_round_formula,
    should_skip_cell_value,
)


def test_build_round_formula_from_value():
    assert build_round_formula_from_value(123.456) == "=ROUND(123.456,2)"
    assert build_round_formula_from_value(123) == "=ROUND(123,2)"


def test_build_round_formula_from_formula():
    assert build_round_formula_from_formula("=A1+B1") == "=ROUND(A1+B1,2)"


def test_build_round_formula_from_external_link_formula():
    formula = r"='D:\xx\[TB.xlsx]Sheet1'!A1"

    assert build_round_formula_from_formula(formula) == r"=ROUND('D:\xx\[TB.xlsx]Sheet1'!A1,2)"


def test_already_round_formula_detection():
    assert is_already_round_formula("=ROUND(A1+B1,2)")
    assert is_already_round_formula("= round ( A1+B1 , 2 )")
    assert not is_already_round_formula("=SUM(ROUND(A1,2),B1)")


def test_build_round_formula_keeps_already_round_formula():
    formula = "=ROUND(A1+B1,2)"

    assert build_round_formula_from_formula(formula) == formula


def test_should_skip_cell_value_for_empty_text_bool_and_non_number():
    assert should_skip_cell_value(None)
    assert should_skip_cell_value("")
    assert should_skip_cell_value("123")
    assert should_skip_cell_value(True)
    assert not should_skip_cell_value(123)
    assert not should_skip_cell_value(123.456)
