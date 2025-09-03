import pandas as pd
from random import randint

"""
NOTES

RANDINT RANGES
- BEER DELIVERY : 1-5
- SET-UP : 6-16
- 5A6 : 17-30
- 6A7 : 31-44
- 7A8 : 45-58
- CLEAN-UP : 59-70

IF DID BEER DELIVERY, INELIGIBLE FOR DOUBLE-SHIFT
IF DID SET-UP, INELIGIBLE FOR CLEAN-UP
"""

EXCEL_PATH = "staff.xlsx"
STAFF_SHEET_NAME = "STAFF"

def read_staff_excel(file_path):
    df = pd.read_excel(io=file_path, sheet_name=STAFF_SHEET_NAME)
    return df

def add_beer_delivery(df):
    inds_used = []
    df['BEER_DELIVERY'] = False
    while len(inds_used) < 5:
        rand_index = randint(0, len(df)-1)
        if rand_index in inds_used:
            continue
        inds_used.append(rand_index)
        df.at[rand_index, 'BEER_DELIVERY'] = True
        print(df.iloc[rand_index])

def main():
    df = read_staff_excel(EXCEL_PATH)
    add_beer_delivery(df)
    print(df)

if __name__ == "__main__":
    main()