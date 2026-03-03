import os
import ast
import pandas as pd

from vissim.utils import get_row_num
from vissim.config_vissim import *


def adjusts():
    df_time_interval_am = pd.read_excel(intersection_list_project, sheet_name='AM', header=0)
    df_time_interval_am = df_time_interval_am.dropna()
    df_time_interval_am['Time_Interval'] = df_time_interval_am['Time_Interval'].apply(lambda x: ast.literal_eval(str(x)))

    df_time_interval_pm = pd.read_excel(intersection_list_project, sheet_name='PM', header=0)
    df_time_interval_pm = df_time_interval_pm.dropna()
    df_time_interval_pm['Time_Interval'] = df_time_interval_pm['Time_Interval'].apply(lambda x: ast.literal_eval(str(x)))

    #####
    # intersection 58: Bel-Red Rd & 20th NE 
    #####
    # 2024 AM
    time_interval_2024_am = df_time_interval_am[df_time_interval_am['Id']==58]['Time_Interval'].values[0]
    fname = os.path.join(dir_vehicle_count[2024], '058AMOct2024.xlsx')
    count_58_2024 = pd.read_excel(fname, header=0, sheet_name='Vehicle')
    count_58_2024 = count_58_2024.set_index('Interval Start')
    count_58_2024 = count_58_2024.loc[time_interval_2024_am, :].copy(deep=True)

    # 2025 PM
    time_interval_2025_pm = df_time_interval_pm[df_time_interval_pm['Id']==58]['Time_Interval'].values[0]
    fname = os.path.join(dir_vehicle_count[2025], '58_Bel-Red Rd_NE 20th St_PM.xlsx')
    df = pd.read_excel(fname, header=10)
    row_start = get_row_num(df, string='Count Summaries - All Vehicles')
    count_58_2025 = df.iloc[row_start+4:row_start+4+8, 1:18]
    count_58_2025.columns = COUNT_COLUMNS
    count_58_2025['Interval Start'] = pd.to_datetime(count_58_2025['Interval Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
    count_58_2025 = count_58_2025.set_index('Interval Start')
    count_58_2025 = count_58_2025.loc[time_interval_2025_pm, :].copy(deep=True)

    # use 2025 ratio to get 2024 movements
    count_58_2025['Total'] = count_58_2025.sum(axis=1).astype(float)
    count_58_2024['15-min Total'] = count_58_2024['15-min Total'].str.replace(',', '').astype(float)
    count_58_2024['idx'] = range(len(count_58_2024))
    count_58_2025['idx'] = range(len(count_58_2025))
    count_58_2024 = count_58_2024.set_index('idx')
    count_58_2025 = count_58_2025.set_index('idx')

    for i in COUNT_COLUMNS[1:]:
        count_58_2024[i] = count_58_2024[i].str.replace(',', '').astype(float)
        count_58_2025[i] = count_58_2025[i].astype(float)
        count_58_2024.loc[:, i] = (count_58_2024.loc[:, '15-min Total'] * count_58_2025.loc[:, i] / count_58_2025['Total']).round(0).astype(int)

    count_58_2024['Interval Start'] = time_interval_2024_am
    count_58_2024 = count_58_2024.set_index('Interval Start')
    count_58_2024_am = count_58_2024[COUNT_COLUMNS[1:]]

    #####
    # intersection 337: 156th Ave NE & NE 10th St
    #####
    # 2025 AM
    time_interval_2025_am = df_time_interval_am[df_time_interval_am['Id']==337]['Time_Interval'].values[0]
    folder = dir_vehicle_count['2025Special']
    fname = os.path.join(folder, '337_156th Ave NE_NE 10th St_AM.xlsx')
    df = pd.read_excel(fname, header=20)
    row_start = get_row_num(df, string='Count Summaries - All Vehicles')
    count_337_am = df.iloc[row_start+4:row_start+4+8, 1:22]
    count_337_am.columns = ['Interval Start', 'EB-UT', 'EB-LT', 'EB-TH', 'EB-RT', 'EB-HR',
                                           'WB-UT', 'WB-LT', 'WB-BL', 'WB-TH', 'WB-RT',
                                           'NB-UT', 'NB-HL', 'NB-LT', 'NB-TH', 'NB-RT',
                                           'SB-UT', 'SB-LT', 'SB-TH', 'SB-BR', 'SB-RT']
    count_337_am['Interval Start'] = pd.to_datetime(count_337_am['Interval Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
    count_337_am = count_337_am.set_index('Interval Start')
    count_337_am.loc[:, 'SB-UT'] = [4, 19, 11, 16, 10, 5, 14, 29]  # from Derq (09/23/2025) 
    count_337_am = count_337_am.loc[time_interval_2025_am, COUNT_COLUMNS[1:]].copy(deep=True)
    
    # 2025 PM
    time_interval_2025_pm = df_time_interval_pm[df_time_interval_pm['Id']==337]['Time_Interval'].values[0]
    fname = os.path.join(folder, '337_156th Ave NE_NE 10th St_PM.xlsx')
    df = pd.read_excel(fname, header=20)
    row_start = get_row_num(df, string='Count Summaries - All Vehicles')
    count_337_pm = df.iloc[row_start+4:row_start+4+8, 1:22]
    count_337_pm.columns = ['Interval Start', 'EB-UT', 'EB-LT', 'EB-TH', 'EB-RT', 'EB-HR',
                                           'WB-UT', 'WB-LT', 'WB-BL', 'WB-TH', 'WB-RT',
                                           'NB-UT', 'NB-HL', 'NB-LT', 'NB-TH', 'NB-RT',
                                           'SB-UT', 'SB-LT', 'SB-TH', 'SB-BR', 'SB-RT']
    count_337_pm['Interval Start'] = pd.to_datetime(count_337_pm['Interval Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
    count_337_pm = count_337_pm.set_index('Interval Start')
    count_337_pm.loc[:, 'SB-UT'] = [7, 14, 12, 11, 16, 13, 19, 15]  # from Derq (09/23/2025) 
    count_337_pm = count_337_pm.loc[time_interval_2025_pm, COUNT_COLUMNS[1:]].copy(deep=True)

    # return the adjusted numbers
    return {'58_AM': count_58_2024_am, 
            '337_AM': count_337_am, 
            '337_PM': count_337_pm}


if __name__ == '__main__':
    adjusts()