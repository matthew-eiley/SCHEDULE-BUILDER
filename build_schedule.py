import pandas as pd
from random import randint, shuffle
import openpyxl
from datetime import date

EXCEL_PATH = "staff.xlsx"
STAFF_SHEET_NAME = "STAFF"
SCHEDULE_SHEET_NAME = "SCHEDULE"
BEER_DELIVERY_CELLS = [
    'J9', 'J10', 'J11', 'J12', 'J13'
]
SET_UP_CELLS = [
    'C18', 'C19', 'C20',
    'F18', 'F19',
    'I18', 'I19',
    'L18', 'L19',
    'O18', 'O19'
]
CELLS_5A6 =[
    'D25', 'D26', 'D27',
    'F25', 'F26',
    'H25', 'H26',
    'J25', 'J26',
    'L25', 'L26',
    'N25', 'N26',
    'P25'
]
CELLS_6A7 =[
    'D28', 'D29', 'D30',
    'F28', 'F29',
    'H28', 'H29',
    'J28', 'J29',
    'L28', 'L29',
    'N28', 'N29',
    'P28'
]
CELLS_7A8 =[
    'D31', 'D32', 'D33',
    'F31', 'F32',
    'H31', 'H32',
    'J31', 'J32',
    'L31', 'L32',
    'N31', 'N32',
    'P31'
]
CLEAN_UP_CELLS =[
    'C38', 'C39', 'C40', 'C41',
    'F38', 'F39',
    'I38', 'I39',
    'L38', 'L39',
    'O38', 'O39'
]

def read_staff_excel(file_path):
    df = pd.read_excel(io=file_path, sheet_name=STAFF_SHEET_NAME)
    return df

def add_beer_delivery(df):
    inds_used = []
    df['BEER_DELIVERY'] = False
    while len(inds_used) < 5:
        rand_index = randint(0, len(df)-1)
        if (rand_index in inds_used) or (df.loc[rand_index, 'NUM_SHIFTS'] >= 2):
            continue
        inds_used.append(rand_index)
        df.loc[rand_index, 'BEER_DELIVERY'] = True
        df.loc[rand_index, 'NUM_SHIFTS'] += 1
    return df

def add_set_up(df):
    inds_used = []
    df['SET_UP'] = False
    while len(inds_used) < 11:
        rand_index = randint(0, len(df)-1)
        if (rand_index in inds_used) or (df.loc[rand_index, 'NUM_SHIFTS'] >= 2):
            continue
        inds_used.append(rand_index)
        df.loc[rand_index, 'SET_UP'] = True
        df.loc[rand_index, 'NUM_SHIFTS'] += 1
    return df

def add_5a6(df):
    inds_used = []
    df['5A6'] = False
    while len(inds_used) < 14:
        rand_index = randint(0, len(df)-1)
        if (rand_index in inds_used) or (df.loc[rand_index, 'NUM_SHIFTS'] >= 2):
            continue
        inds_used.append(rand_index)
        df.loc[rand_index, '5A6'] = True
        df.loc[rand_index, 'NUM_SHIFTS'] += 1
    return df

def add_6a7(df):
    inds_used = []
    df['6A7'] = False
    while len(inds_used) < 14:
        rand_index = randint(0, len(df)-1)
        if (rand_index in inds_used) or (df.loc[rand_index, 'NUM_SHIFTS'] >= 2):
            continue
        inds_used.append(rand_index)
        df.loc[rand_index, '6A7'] = True
        df.loc[rand_index, 'NUM_SHIFTS'] += 1
    return df

def add_7a8(df):
    inds_used = []
    df['7A8'] = False
    while len(inds_used) < 14:
        rand_index = randint(0, len(df)-1)
        if (rand_index in inds_used) or (df.loc[rand_index, 'NUM_SHIFTS'] >= 2):
            continue
        inds_used.append(rand_index)
        df.loc[rand_index, '7A8'] = True
        df.loc[rand_index, 'NUM_SHIFTS'] += 1
    return df

def add_clean_up(df):
    inds_used = []
    df['CLEAN_UP'] = False
    while len(inds_used) < 12:
        rand_index = randint(0, len(df)-1)
        if (rand_index in inds_used) or (df.loc[rand_index, 'NUM_SHIFTS'] >= 2):
            continue
        inds_used.append(rand_index)
        df.loc[rand_index, 'CLEAN_UP'] = True
        df.loc[rand_index, 'NUM_SHIFTS'] += 1
    return df
    
def shift_staff(df):
    df['NUM_SHIFTS'] = 0
    df = add_beer_delivery(df)
    df = add_set_up(df)
    df = add_5a6(df)
    df = add_6a7(df)
    df = add_7a8(df)
    df = add_clean_up(df)
    return df

def get_and_shuffle(df, col_name):
    lst = df.loc[df[col_name], 'NAME'].tolist()
    shuffle(lst)
    return lst

def fill_template(df):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SCHEDULE_SHEET_NAME]

    ws['B3'] = f"MUS 4À7 SCHEDULE ({date.today()})"

    beer_delivery_names = get_and_shuffle(df, 'BEER_DELIVERY')
    for i, cell in enumerate(BEER_DELIVERY_CELLS):
        ws[cell] = beer_delivery_names[i]

    set_up_names = get_and_shuffle(df, 'SET_UP')
    for i, cell in enumerate(SET_UP_CELLS):
        ws[cell] = set_up_names[i]

    names_5a6 = get_and_shuffle(df, '5A6')
    for i, cell in enumerate(CELLS_5A6):
        ws[cell] = names_5a6[i]

    names_6a7 = get_and_shuffle(df, '6A7')
    for i, cell in enumerate(CELLS_6A7):
        ws[cell] = names_6a7[i]

    names_7a8 = get_and_shuffle(df, '7A8')
    for i, cell in enumerate(CELLS_7A8):
        ws[cell] = names_7a8[i]

    clean_up_names = get_and_shuffle(df, 'CLEAN_UP')
    for i, cell in enumerate(CLEAN_UP_CELLS):
        ws[cell] = clean_up_names[i]

    wb.save(EXCEL_PATH)

def main():
    df = read_staff_excel(EXCEL_PATH)
    df = shift_staff(df)
    fill_template(df)
    print(df)

if __name__ == "__main__":
    main()