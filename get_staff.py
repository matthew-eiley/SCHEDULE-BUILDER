import pandas as pd
from random import randint

"""
RANDINT RANGES
- BEER DELIVERY : 1-5
- SET-UP : 6-16
- 5A6 : 17-30
- 6A7 : 31-44
- 7A8 : 45-58
- CLEAN-UP : 59-70
"""

EXCEL_PATH = "staff.xlsx"
STAFF_SHEET_NAME = "STAFF"

def read_staff_excel(file_path):
    df = pd.read_excel(io=file_path, sheet_name=STAFF_SHEET_NAME)
    return df


def main():
    df = read_staff_excel(EXCEL_PATH)
    print(df)

