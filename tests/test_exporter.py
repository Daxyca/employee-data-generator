import pandas as pd
from pathlib import Path
from src.exporter import write_employees_excel_file


def test_export_creates_file(tmp_path):
    # Create fake employee data
    df = pd.DataFrame(
        [
            {
                "Employee ID": 1,
                "Full Name": "John Doe",
                "Department": "IT",
                "Salary": 50000,
                "Hire Date": "2021-01-01",
            },
            {
                "Employee ID": 2,
                "Full Name": "Jane Smith",
                "Department": "HR",
                "Salary": 60000,
                "Hire Date": "2022-03-05",
            },
        ]
    )

    # Export to a temporary folder
    file_path = tmp_path / "employees.xlsx"
    write_employees_excel_file(df, file_path)

    # Check file exists
    assert Path(file_path).exists()
    assert file_path.name == "employees.xlsx"
