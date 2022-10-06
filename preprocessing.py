# link to google sheet
# https://docs.google.com/spreadsheets/d/11VbtcJpKkxl-c88S-cNVWrBVq89cQEYNo6Y760j4I6I/edit#gid=0

import numpy as np
import pandas as pd
import os
import shutil


def save_excel(file_name, df, standard_type):
    with pd.ExcelWriter(f'{os.getcwd()}\\D. Query tool\\{standard_type}\\{file_name}.xlsx',
                        engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Service Packages')


def rearrange_columns(df):
    return df.loc[:,
           ['Store Name',  # to group by
            'Store Name (read only)',
            'Package ID (read only)',
            'Package Name',
            'Items In Package',
            'Addons Names',
            'Addons Prices',
            'Gender (read only)',
            'Categories',
            'Description',
            'Prerequisite',
            'Price (RM)',
            'Status']]


def save_df_as_excel(df, standard_type):
    # this method takes a giant df
    # group by all the store name
    # save each store as an excel file

    # create directory "D. Query tool" if not exist
    d_query_tool_dir = f'{os.getcwd()}\\D. Query tool'
    if not os.path.exists(d_query_tool_dir):
        os.mkdir(d_query_tool_dir)
    # create subdirectory "{standard_type}" if not exist
    standard_type_path = f'{os.getcwd()}\\D. Query tool\\{standard_type}'
    if not os.path.exists(standard_type_path):
        os.mkdir(standard_type_path)

    groups = df.groupby('Store Name')

    for store_name, group_df in groups:
        group_df.drop('Store Name', axis=1, inplace=True)
        # display(x)

        try:
            save_excel(store_name, group_df, standard_type)
        # except OSError:
        #     # this file has naming problem
        #     # Klinik Malaysia Masai\n.xlsx
        #     print(store_name)
        #     store_name = store_name.replace( '\n', '')
        #     save_excel(store_name, group_df, standard_type)

        except OSError:
            print(store_name)
            store_name = store_name.replace('\n', '')


def main(uploaded_file):
    #################################
    # STEP 1, read and preprocess the xlsx file
    #################################
    xls = pd.ExcelFile(uploaded_file)
    # xls.sheet_names
    # ['Sheet1', 'C', 'B', 'A1', 'daniel testing']

    A = pd.read_excel(xls, 'A1')

    # drop( 'No', axis=1).\
    B = pd.read_excel(xls, 'B'). \
        rename(columns={'Name of Clinic (Store Name)': 'Store Name'}) \
        [['Store Name', 'Name of Services', 'Selling Price (RM)\n(Display on DOC Platform)',
          'Standard Category']]  # extract only the required columns

    C = pd.read_excel(xls, 'C'). \
        rename(columns={'Package Name': 'Name of Services'})
    C = C[C['Name of Services'].notnull()] # drop rows where 'Name of Services' is null





    #################################
    # STEP 2, inner join all 3 sheets
    #################################
    A_B = A.merge(B, on='Store Name', how='inner')            # inner join, KEEP only stores with services
    A_B_C = A_B.merge(C, on='Name of Services', how='inner')  # inner join, KEEP only service currently in our service list

    # rename columns
    A_B_C.rename(
        columns={
            'Name of Services': 'Package Name',
            'Selling Price (RM)\n(Display on DOC Platform)': 'Price (RM)'},
        inplace=True)

    # as per output format
    A_B_C['Store Name (read only)'] = np.nan
    A_B_C['Package ID (read only)'] = np.nan
    A_B_C['Status'] = np.nan

    # separated into 'Standard' and 'Non Standard' as per requirements
    ABC_standard = A_B_C[A_B_C['Standard Category'] == 'Standard']
    ABC_nonstandard = A_B_C[A_B_C['Standard Category'] == 'Non Standard']

    # rearrange columns to fit that of the google sheet
    ABC_standard = rearrange_columns(ABC_standard)
    ABC_nonstandard = rearrange_columns(ABC_nonstandard)





    ##############################################
    # STEP 3, save them as xlsx file, then zip it
    ##############################################
    save_df_as_excel(ABC_standard, 'Standard')
    save_df_as_excel(ABC_nonstandard, 'Non Standard')

    # save folder as zip
    #                     output zip              input folder
    shutil.make_archive('output', 'zip', 'D. Query tool')