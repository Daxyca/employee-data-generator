# Employee Data Generator (PySide6 + Pandas)

A simple desktop application built with **PySide6** that generates random employee data and exports it to **Excel**.

## Features

- Generate any number of random employee records
- Assigns random departments and realistic salaries
- Includes random hire dates between 2020 and today
- Exports to an Excel file with:
  - detailed `Employees` sheet
  - `Summary` sheet showing average salary per department
- Choose where to save the Excel file
- Simple, modern PySide6 GUI — no command line needed

### Employee Data Generator UI

<img width="502" height="332" alt="image" src="https://github.com/user-attachments/assets/8f995f41-dd8e-4124-a1dc-f8cd74f62eb5" />

## Tech Stack

| Component                   | Description                      |
| --------------------------- | -------------------------------- |
| **Python 3.12+**            | Programming language             |
| **PySide6 (Qt for Python)** | GUI framework                    |
| **Pandas**                  | Data manipulation & Excel export |
| **OpenPyXL**                | Excel writing backend for Pandas |

## Installation

### 1. Clone the Repository

```bash
git clone git@github.com:Daxyca/employee-data-generator.git
cd employee-data-generator
```

### 2. Create a Virtual Environment (optional but recommended)

```bash
python -m venv .venv
source .venv/bin/activate     # On Windows: .venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

Run the application with:

```bash
python -m src.app
```

Then:

1. Enter how many employees you want to generate.
2. Choose a folder where the Excel file will be saved.
3. Click **“Generate Data”**.
4. Click **“Export to Excel”** — and your file will be ready!

The exported file will contain:

- `employees.xlsx`

  - **Employees** – full employee dataset
  - **Summary** – department-level salary averages with timestamp

## Example Output

**Employees Sheet**

| Employee ID | Full Name    | Department | Salary | Hire Date  |
| ----------- | ------------ | ---------- | ------ | ---------- |
| 1           | Mary Johnson | IT         | 72000  | 2022/04/15 |
| 2           | Robert Smith | HR         | 51000  | 2021/08/23 |
| 3           | Linda Brown  | Finance    | 84000  | 2020/11/02 |

**Summary Sheet**

| Department        | Average Salary      |
| ----------------- | ------------------- |
| IT                | 68500.23            |
| HR                | 59830.75            |
| Finance           | 74510.89            |
|                   |                     |
| Export Timestamp: | 2025-10-21 03:51:10 |

## File Structure

```
employee-data-generator/
├─ src
│  ├─ app.py             # Launch entry point
│  ├─ exporter.py        # Handles Excel writing (pure logic)
│  ├─ generator.py       # Generates synthetic data (pure logic)
│  └─ ui.py              # Handles PySide6 widgets and layout
├─ tests
│  ├─ test_exporter.py
│  └─ test_generator.py
├─ LICENSE               # MIT License
├─ README.md             # Project documentation
├─ requirements.txt      # Dependencies
└─ employees.xlsx        # Generated after export
```

## Customization Tips

You can easily:

- Add more **departments** or **names** in the constant lists.
- Adjust salary ranges in `create_employee_data()`.
- Change the export filename or Excel sheet names.
- Add filters, sorting.

## Possible Enhancements

- Progress bar for large datasets
- Dark/light theme toggle
- Department and salary filters
- Option to export to CSV
- Add random job titles or office locations

## License

This project is open-source and available under the [MIT License](LICENSE).
