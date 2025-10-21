from datetime import datetime
from pathlib import Path
import pandas as pd


def write_employees_excel_file(employee_data: pd.DataFrame, path: Path) -> None:
    """Write the main data and summary to an Excel file."""
    summary = (
        employee_data.groupby("Department")["Salary"]
        .mean()
        .reset_index()
        .rename(columns={"Salary": "Average Salary"})
    )
    summary["Average Salary"] = summary["Average Salary"].round(2)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        employee_data.to_excel(writer, sheet_name="Employees", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        summary_sheet = writer.sheets["Summary"]
        row = len(summary) + 3
        summary_sheet.cell(row=row, column=1, value="Export Timestamp:")
        summary_sheet.cell(row=row, column=2, value=timestamp)
