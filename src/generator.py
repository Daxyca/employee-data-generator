import random
from datetime import datetime, timedelta
import pandas as pd

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


def create_employee_data(count: int) -> pd.DataFrame:
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
