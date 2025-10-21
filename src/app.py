import sys
from PySide6.QtWidgets import QApplication

from .ui import apply_preset_styles, EmployeeGeneratorApp


def main():
    app = QApplication(sys.argv)
    apply_preset_styles(app)
    window = EmployeeGeneratorApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
