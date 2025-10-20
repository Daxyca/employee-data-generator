import sys
import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
)
from PySide6.QtCore import Qt


# --- Constants ---
FIRST_NAMES = [
    "James",
    "Mary",
    "John",
    "Patricia",
    "Robert",
    "Jennifer",
    "Michael",
    "Linda",
    "William",
    "Barbara",
    "David",
    "Elizabeth",
    "Richard",
    "Susan",
    "Joseph",
    "Jessica",
    "Thomas",
    "Sarah",
    "Charles",
    "Karen",
    "Christopher",
    "Nancy",
    "Daniel",
    "Lisa",
    "Matthew",
    "Betty",
    "Anthony",
    "Margaret",
    "Mark",
    "Sandra",
    "Donald",
    "Ashley",
    "Steven",
    "Kimberly",
    "Paul",
    "Emily",
    "Andrew",
    "Donna",
    "Joshua",
    "Michelle",
]

LAST_NAMES = [
    "Smith",
    "Johnson",
    "Williams",
    "Brown",
    "Jones",
    "Garcia",
    "Miller",
    "Davis",
    "Rodriguez",
    "Martinez",
    "Hernandez",
    "Lopez",
    "Gonzalez",
    "Wilson",
    "Anderson",
    "Thomas",
    "Taylor",
    "Moore",
    "Jackson",
    "Martin",
    "Lee",
    "Thompson",
    "White",
    "Harris",
    "Sanchez",
    "Clark",
    "Ramirez",
    "Lewis",
    "Robinson",
    "Walker",
    "Young",
]

DEPARTMENTS = ["IT", "HR", "Operations", "Administration", "Finance"]


class EmployeeGeneratorApp(QMainWindow):
    """A simple GUI application to generate and export random employee data to Excel."""

    def __init__(self):
        super().__init__()
        self.selected_folder: str | None = None
        self.employee_data: pd.DataFrame | None = None
        self._setup_ui()

    # -------------------------------------------------------------------------
    # UI SETUP
    # -------------------------------------------------------------------------
    def _setup_ui(self) -> None:
        self.setWindowTitle("Employee Data Generator")
        self.setMinimumSize(500, 300)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        layout.addLayout(self._build_employee_input())
        layout.addLayout(self._build_folder_selector())
        layout.addWidget(self._build_button("Generate Data", self.generate_data))
        self.export_button = self._build_button(
            "Export to Excel", self.export_to_excel, enabled=False
        )
        layout.addWidget(self.export_button)

        self.message_label = QLabel("")
        self._style_message_label()
        layout.addWidget(self.message_label)

        layout.addStretch()

    def _build_employee_input(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        label = QLabel("Number of Employees:")
        self.emp_input = QLineEdit()
        self.emp_input.setPlaceholderText("Enter number (e.g., 100)")
        layout.addWidget(label)
        layout.addWidget(self.emp_input)
        return layout

    def _build_folder_selector(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        self.selected_folder_label = QLabel("No folder selected")
        button = QPushButton("Select Folder")
        button.clicked.connect(self.select_folder)
        layout.addWidget(self.selected_folder_label)
        layout.addWidget(button)
        return layout

    def _build_button(self, text: str, handler, enabled: bool = True) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(handler)
        button.setEnabled(enabled)
        return button

    def _style_message_label(self, color: str = "green") -> None:
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setStyleSheet(
            f"font-weight: bold; color: {color}; padding: 10px;"
        )

    # -------------------------------------------------------------------------
    # EVENT HANDLERS
    # -------------------------------------------------------------------------
    def select_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder to Save Excel File"
        )
        if folder:
            self.selected_folder = folder
            self.selected_folder_label.setText(f"Selected: {folder}")
            self.message_label.clear()

    def generate_data(self) -> None:
        """Generate random employee data based on user input."""
        try:
            num_employees = int(self.emp_input.text())
            if num_employees <= 0:
                raise ValueError

            self.employee_data = self._create_employee_data(num_employees)
            self.export_button.setEnabled(True)
            self._update_message(
                f"Generated {num_employees} employee records", color="blue"
            )

        except ValueError:
            QMessageBox.warning(
                self, "Invalid Input", "Please enter a valid positive number."
            )
            self.message_label.clear()

    def export_to_excel(self) -> None:
        """Export employee data to an Excel file with a summary sheet."""
        if self.employee_data is None:
            QMessageBox.warning(self, "No Data", "Please generate data first.")
            return

        if not self.selected_folder:
            QMessageBox.warning(self, "No Folder", "Please select a folder first.")
            return

        file_path = Path(self.selected_folder) / "employees.xlsx"

        try:
            self._write_excel_file(file_path)
            self._update_message("File generated successfully!", color="green")
            QMessageBox.information(self, "Success", f"File saved at:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to export file:\n{str(e).replace('\\\\', '\\')}"
            )
            self.message_label.clear()

    # -------------------------------------------------------------------------
    # CORE LOGIC
    # -------------------------------------------------------------------------
    def _create_employee_data(self, count: int) -> pd.DataFrame:
        """Generate a DataFrame with random employee data."""
        start_date = datetime(2020, 1, 1)
        end_date = datetime.now()
        date_range = (end_date - start_date).days

        employees = [
            {
                "Employee ID": i + 1,
                "Full Name": f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}",
                "Department": random.choice(DEPARTMENTS),
                "Salary": random.randint(25_000, 120_000),
                "Hire Date": (
                    start_date + timedelta(days=random.randint(0, date_range))
                ).strftime("%Y/%m/%d"),
            }
            for i in range(count)
        ]

        return pd.DataFrame(employees)

    def _write_excel_file(self, path: Path) -> None:
        """Write the main data and summary to an Excel file."""
        summary = (
            self.employee_data.groupby("Department")["Salary"]
            .mean()
            .reset_index()
            .rename(columns={"Salary": "Average Salary"})
        )
        summary["Average Salary"] = summary["Average Salary"].round(2)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            self.employee_data.to_excel(writer, sheet_name="Employees", index=False)
            summary.to_excel(writer, sheet_name="Summary", index=False)

            summary_sheet = writer.sheets["Summary"]
            row = len(summary) + 3
            summary_sheet.cell(row=row, column=1, value="Export Timestamp:")
            summary_sheet.cell(row=row, column=2, value=timestamp)

    def _update_message(self, text: str, color: str = "green") -> None:
        """Update the message label with styled feedback."""
        self.message_label.setText(text)
        self._style_message_label(color)


def apply_preset_styles(app):
    app.setStyle("Fusion")
    app.setStyleSheet("""
        QMainWindow {
            background-color: #1e1e1e;
            color: #ffffff;
        }

        QLabel {
            color: #e0e0e0;
            font-size: 14px;
        }

        QLineEdit {
            background-color: #2d2d2d;
            color: #ffffff;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 6px;
        }

        QPushButton {
            background-color: #0078d7;
            color: white;
            border-radius: 8px;
            padding: 8px 14px;
            font-weight: bold;
        }

        QPushButton:hover {
            background-color: #1a8cff;
        }

        QPushButton:disabled {
            background-color: #3a3a3a;
            color: #777;
        }

        QMessageBox {
            background-color: #2c2c2c;
            color: #ffffff;
        }
    """)


# -------------------------------------------------------------------------
# ENTRY POINT
# -------------------------------------------------------------------------
def main():
    app = QApplication(sys.argv)
    apply_preset_styles(app)
    window = EmployeeGeneratorApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
