import random
from datetime import datetime, timedelta
import pandas as pd
from faker import Faker

DEPARTMENTS = ["IT", "HR", "Operations", "Administration", "Finance"]


def create_employee_data(count: int) -> pd.DataFrame:
    """Generate a DataFrame with random employee data."""
    start_date = datetime(2020, 1, 1)
    end_date = datetime.now()
    date_range = (end_date - start_date).days

    fake = Faker()
    employees: list[dict[str, str | int]] = [
        {
            "Employee ID": i + 1,
            "Full Name": f"{fake.first_name()} {fake.last_name()}",
            "Department": random.choice(DEPARTMENTS),
            "Salary": random.randint(25_000, 120_000),
            "Hire Date": (
                start_date + timedelta(days=random.randint(0, date_range))
            ).strftime("%Y/%m/%d"),
        }
        for i in range(count)
    ]

    return pd.DataFrame(employees)
