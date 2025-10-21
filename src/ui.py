from pathlib import Path
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
import pandas as pd

from .exporter import write_employees_excel_file
from .generator import create_employee_data


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

            self.employee_data = create_employee_data(num_employees)
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
            write_employees_excel_file(self.employee_data, file_path)
            self._update_message("File generated successfully!", color="green")
            QMessageBox.information(self, "Success", f"File saved at:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to export file:\n{str(e).replace('\\\\', '\\')}"
            )
            self.message_label.clear()

    # -------------------------------------------------------------------------
    # UPDATE LOGIC
    # -------------------------------------------------------------------------
    def _update_message(self, text: str, color: str = "green") -> None:
        """Update the message label with styled feedback."""
        self.message_label.setText(text)
        self._style_message_label(color)


def apply_preset_styles(app: QApplication) -> None:
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
