import pandas as pd

STAFF_EXCEL_PATH = "staff.xlsx"

def read_staff_excel(file_path):
    df = pd.read_excel(file_path)
    return df

print(read_staff_excel(STAFF_EXCEL_PATH))