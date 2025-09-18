import pandas as pd
from random import randint, shuffle
import openpyxl
from datetime import date

EXCEL_PATH = "chief_of_staff.xlsx"
STAFF_SHEET_NAME = "STAFF"

def read_staff_excel(file_path):
    df = pd.read_excel(io=file_path, sheet_name=STAFF_SHEET_NAME)
    df['STATUS'] = df['STATUS'].map({'Available': True, 'Unavailable': False})
    return df

def do_beer_delivery(df):
    num_needed = 5
    num_staff = len(df)
    possible_inds = [i for i in range(num_staff)]
    inds_chosen = []

    for _ in range(num_needed):
        rand_ind = randint(0, len(possible_inds) - 1)
        inds_chosen.append(possible_inds[rand_ind])
        possible_inds.pop(rand_ind)

    df['BEER_DELIVERY'] = False
    df.loc[inds_chosen, 'BEER_DELIVERY'] = True

    return df

def do_set_and_clean_up(df):
    set_up_num_needed = 11
    clean_up_num_needed = 12
    num_staff = len(df)
    possible_inds = [i for i in range(num_staff) if df.loc[i, 'STATUS']]
    set_up_inds_chosen = []
    clean_up_inds_chosen = []

    for _ in range(set_up_num_needed):
        rand_ind = randint(0, len(possible_inds) - 1)
        set_up_inds_chosen.append(possible_inds[rand_ind])
        possible_inds.pop(rand_ind)

    for _ in range(clean_up_num_needed):
        rand_ind = randint(0, len(possible_inds) - 1)
        clean_up_inds_chosen.append(possible_inds[rand_ind])
        possible_inds.pop(rand_ind)

    df['SET_UP'] = False
    df.loc[set_up_inds_chosen, 'SET_UP'] = True

    df['CLEAN_UP'] = False
    df.loc[clean_up_inds_chosen, 'CLEAN_UP'] = True

    return df

def do_56_and_78(df):
    # assumes num_needed >= num_needed for setup and cleanup
    num_needed = 14
    num_spots_56 = num_needed - df['SET_UP'].sum()
    num_spots_78 = num_needed - df['CLEAN_UP'].sum()
    possible_inds = [i for i in range(len(df)) if df.loc[i, 'STATUS'] and not df.loc[i, 'SET_UP'] and not df.loc[i, 'CLEAN_UP']]
    inds_chosen_56 = []
    inds_chosen_78 = []

    for _ in range(num_spots_56):
        rand_ind = randint(0, len(possible_inds) - 1)
        inds_chosen_56.append(possible_inds[rand_ind])
        possible_inds.pop(rand_ind)

    for _ in range(num_spots_78):
        rand_ind = randint(0, len(possible_inds) - 1)
        inds_chosen_78.append(possible_inds[rand_ind])
        possible_inds.pop(rand_ind)

    df['5-6'] = False
    df.loc[inds_chosen_56, '5-6'] = True
    df.loc[df['SET_UP'], '5-6'] = True

    df['7-8'] = False
    df.loc[inds_chosen_78, '7-8'] = True
    df.loc[df['CLEAN_UP'], '7-8'] = True

    return df

def main():
    df = read_staff_excel(EXCEL_PATH)
    df = do_beer_delivery(df)
    df = do_set_and_clean_up(df)
    df = do_56_and_78(df)
    print(df)

if __name__ == "__main__":
    main()