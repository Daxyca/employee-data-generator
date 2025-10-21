import pandas as pd
from src.generator import create_employee_data


def test_generate_correct_length():
    df = create_employee_data(10)
    assert len(df) == 10
    assert isinstance(df, pd.DataFrame)


def test_generate_has_expected_columns():
    df = create_employee_data(5)
    expected = {"Employee ID", "Full Name", "Department", "Salary", "Hire Date"}
    assert expected.issubset(df.columns)


def test_salary_range():
    df = create_employee_data(50)
    assert (df["Salary"] >= 25000).all()
    assert (df["Salary"] <= 120000).all()
