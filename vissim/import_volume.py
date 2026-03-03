import os
import ast
import logging
import pandas as pd

from logging import config as logger_config
from vissim.config_vissim import *
from vissim.adjust_volume import adjusts
from vissim.utils import get_row_num, start_vissim

logger_config.fileConfig(os.path.join(os.getcwd(), 'vissim', 'logger.conf'))
# create logger
logger = logging.getLogger('importVolume')


def get_count():
    logger.info('Loading traffic volumes...')
    # read intersection list
    df_intersection_project = pd.read_excel(intersection_list_project, sheet_name='Counts', header=3)
    df_intersection_project = df_intersection_project.dropna(subset=['Id', f'{PROJECT_PERIOD}_Counts'])
    df_time_interval = pd.read_excel(intersection_list_project, sheet_name=f'{PROJECT_PERIOD}', header=0)
    df_time_interval = df_time_interval.dropna()
    df_time_interval['Time_Interval'] = df_time_interval['Time_Interval'].apply(lambda x: ast.literal_eval(str(x)))
    book_veh = {}
    book_ped = {}
    for _, row in df_intersection_project.iterrows():
        idx = int(row['Id'])
        ns_street = row['NS_Street'].rstrip(' ')
        ew_street = row['EW_Street'].lstrip(' ')
        time_interval = df_time_interval[df_time_interval['Id']==idx]['Time_Interval'].values[0]
        if isinstance(row[f'{PROJECT_PERIOD}_Counts'], int):
            year = int(row[f'{PROJECT_PERIOD}_Counts'])
        else:
            year = row[f'{PROJECT_PERIOD}_Counts']
        # skip Redmond 2023 count until we know the location
        if year == 'NaN': continue
        # tailor the reading according to the year
        if year == 2023:
            folder = dir_vehicle_count[year]
            for month in ['May', 'July', 'Oct']:
                fname = os.path.join(folder, f'{idx:03d}{PROJECT_PERIOD.lower()}{month}{year}.xlsx')
                if os.path.exists(fname):
                    break
            df = pd.read_excel(fname, header=29)
            count = df.iloc[:8, 1:36].dropna(axis=1)
            count.columns = COUNT_COLUMNS
            count['Interval Start'] = pd.to_datetime(count['Interval Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
            count = count.set_index('Interval Start')
            ped = df.iloc[16:24, :].dropna(axis=1, how='all')
            ped.columns = PED_COLUMNS
            ped = ped[['Start', 'Ped-EB', 'Ped-WB', 'Ped-NB', 'Ped-SB', 'Ped-Total']]
            ped = ped.rename(columns={i: i.split('-')[1] for i in ped.columns if i != 'Start'})
            ped['Start'] = pd.to_datetime(ped['Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
            ped = ped.set_index('Start')
            # by movement
            for col, (first, second) in PED_SPLITS.items():
                ped[first] = (ped[col] / 2).round().astype(int)
                ped[second] = ped[col] - ped[first]
            # by input
            for pos, split in PED_POSITIONS.items():
                ped[pos] = ped[split].sum(axis=1)
            book_veh[idx] = count.loc[time_interval, :]
            book_ped[idx] = ped.loc[time_interval, :]
        elif year == 2024:
            folder = dir_vehicle_count[year]
            for month in ['May', 'July', 'Oct']:
                fname = os.path.join(folder, f'{idx:03d}{PROJECT_PERIOD.lower()}{month}{year}.xlsx')
                if os.path.exists(fname):
                    break
            count = pd.read_excel(fname, header=0, sheet_name='Vehicle')
            count = count.set_index('Interval Start')
            book_veh[idx] = count.loc[time_interval, :]
            # split ped count by directions
            ped = pd.read_excel(fname, header=0, sheet_name='Other')[['Start', 'Ped-EB', 'Ped-WB', 'Ped-NB', 'Ped-SB', 'Ped-Total']]
            ped = ped.rename(columns={i: i.split('-')[1] for i in ped.columns if i != 'Start'})
            ped = ped.set_index('Start')
            # by movement
            for col, (first, second) in PED_SPLITS.items():
                ped[first] = (ped[col] / 2).round().astype(int)
                ped[second] = ped[col] - ped[first]
            # by input
            for pos, split in PED_POSITIONS.items():
                ped[pos] = ped[split].sum(axis=1)
            book_ped[idx] = ped.loc[time_interval, :]
        elif year == 2025 or year == '2025Special':
            folder = dir_vehicle_count[year]
            fname = os.path.join(folder, f'{idx:.0f}_{ns_street}_{ew_street}_{PROJECT_PERIOD}.xlsx')
            df = pd.read_excel(fname, header=10)
            row_start = get_row_num(df, string='Count Summaries - All Vehicles')
            ped = df.iloc[row_start+20:row_start+20+8, :].dropna(axis=1)
            if idx == 337:
                count = df.iloc[row_start+4:row_start+4+8, 1:22]
                count.columns = ['Interval Start', 'EB-UT', 'EB-LT', 'EB-TH', 'EB-RT', 'EB-HR',
                                                   'WB-UT', 'WB-LT', 'WB-BL', 'WB-TH', 'WB-RT',
                                                   'NB-UT', 'NB-HL', 'NB-LT', 'NB-TH', 'NB-RT',
                                                   'SB-UT', 'SB-LT', 'SB-TH', 'SB-BR', 'SB-RT']
                ped.columns = ['Start', 'HV-EB', 'HV-WB', 'HV-NB', 'HV-SB', 'HV-NEB', 'HV-Total', 
                               'Bike-EB', 'Bike-WB', 'Bike-NB', 'Bike-SB', 'Bike-NEB', 'Bike-Total',
                               'Ped-EB', 'Ped-WB', 'Ped-NB', 'Ped-SB', 'Ped-SW', 'Ped-Total',]
                ped = ped[PED_COLUMNS].copy(deep=True)
            else:
                count = df.iloc[row_start+4:row_start+4+8, 1:18]
                count.columns = COUNT_COLUMNS
                ped.columns = PED_COLUMNS
            count['Interval Start'] = pd.to_datetime(count['Interval Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
            count = count.set_index('Interval Start')
            ped = ped[['Start', 'Ped-EB', 'Ped-WB', 'Ped-NB', 'Ped-SB', 'Ped-Total']]
            ped = ped.rename(columns={i: i.split('-')[1] for i in ped.columns if i != 'Start'})
            ped['Start'] = pd.to_datetime(ped['Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
            ped = ped.set_index('Start')
            # by movement
            for col, (first, second) in PED_SPLITS.items():
                ped[first] = (ped[col] / 2).round().astype(int)
                ped[second] = ped[col] - ped[first]
            # by input
            for pos, split in PED_POSITIONS.items():
                ped[pos] = ped[split].sum(axis=1)
            book_veh[idx] = count.loc[time_interval, :]
            book_ped[idx] = ped.loc[time_interval, :]
        elif year == '2023Redmond':
            folder = dir_vehicle_count[year]
            fname = os.path.join(folder, f'{idx:.0f}_{ns_street}_{ew_street}_{PROJECT_PERIOD}.xlsx')
            df = pd.read_excel(fname, header=30)
            count = df.iloc[:8, 1:36].dropna(axis=1)
            count.columns = COUNT_COLUMNS
            count['Interval Start'] = pd.to_datetime(count['Interval Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
            count = count.set_index('Interval Start')
            ped = df.iloc[14:22].dropna(axis=1)
            ped.columns = PED_COLUMNS
            ped = ped[['Start', 'Ped-EB', 'Ped-WB', 'Ped-NB', 'Ped-SB', 'Ped-Total']]
            ped = ped.rename(columns={i: i.split('-')[1] for i in ped.columns if i != 'Start'})
            ped['Start'] = pd.to_datetime(ped['Start'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M %p').str.lstrip('0')
            ped = ped.set_index('Start')
            # by movement
            for col, (first, second) in PED_SPLITS.items():
                ped[first] = (ped[col] / 2).round().astype(int)
                ped[second] = ped[col] - ped[first]
            # by input
            for pos, split in PED_POSITIONS.items():
                ped[pos] = ped[split].sum(axis=1)
            book_veh[idx] = count.loc[time_interval, :]
            book_ped[idx] = ped.loc[time_interval, :]
        
    # get the adjusted traffic volumes 
    logger.info('Adjusting traffic volumes...')
    adjusted = adjusts()
    for k, v in adjusted.items():
        idx, period = k.split('_')
        logger.info(f'    Int {idx} {period} period')
        idx = int(idx)
        if period == PROJECT_PERIOD:
            book_veh[idx] = v

    return book_veh, book_ped, df_intersection_project

def main():
    # get input volumes from actual vehicle counts
    book_veh, book_ped, intersections = get_count()
    intersection_list = intersections['Id'].values

    # start vissim
    vissim = start_vissim()
    # load a Vissim project:
    net_name = os.path.join(WORKING_PATH, f'{PROJECT_NAME}.inpx')
    
    # load network
    flag_read_additionally = True # you can read network(elements) additionally, in this case set flag_read_additionally" to true
    vissim.LoadNet(net_name, flag_read_additionally)
    
    # load layout
    layout_name = os.path.join(WORKING_PATH, f'{PROJECT_NAME}.layx')
    vissim.LoadLayout(layout_name)

    # vehicle routes
    logger.info('Importing vehicle movement splits...')
    vnet = vissim.Net
    veh_routes = vnet.VehicleRoutingDecisionsStatic
    for veh_route in veh_routes:
        route_name = veh_route.AttValue('Name')
        int_idx = int(route_name.split('_')[0])
        if int_idx not in intersection_list : continue
        count = book_veh[int_idx]
        route_decisions = veh_route.VehRoutSta
        for route_decision in route_decisions:
            movement = route_decision.AttValue('Name')
            # configure volumnes at each time interval
            # time interval: 1: 0-900, 2: 900-1800, 3:1800-2700, 4: 2700-3600, 5: 3600-4500, 6: 4500-MAX
            for i, (_, row) in enumerate(count.iterrows()):
                mvmnt_count = int(row[movement]) * 4
                route_decision.SetAttValue(f'RelFlow({i+1})', mvmnt_count)

    # vehicle inputs
    logger.info('Importing vehicle volume inputs...')
    veh_inputs = vnet.VehicleInputs
    for veh_input in veh_inputs:
        input_name = veh_input.AttValue('Name')
        int_no = veh_input.AttValue('No')  # intersection no. in vissim
        int_idx = int(input_name.split('_')[0])
        if int_idx not in intersection_list: continue
        count = book_veh[int_idx]
        veh_route = veh_routes.ItemByKey(int_no)
        route_decisions = veh_route.VehRoutSta
        movements = []
        for route_decision in route_decisions:
            movements.append(route_decision.AttValue('Name'))
        count_int = count.loc[:, movements].astype(int).sum(axis=1).to_frame()
        count_int.columns = ['Count']
        # configure volumes
        for i, (_, row) in enumerate(count_int.iterrows()):
            veh_input.SetAttValue(f'Volume({i+1})', row['Count'] * 4)   # the input is volume per hour

    # ped inputs
    logger.info('Importing pedestrian movement splits and volume inputs...')
    ped_inputs = vnet.PedestrianInputs
    ped_routes = vnet.PedestrianRoutingDecisionsStatic
    for ped_input in ped_inputs:
        input_name = ped_input.AttValue('Name')
        int_no = ped_input.AttValue('No')  # intersection no. in vissim
        int_idx = int(input_name.split('_')[0])
        int_pos = input_name.split('_')[2]
        if int_idx not in intersection_list: continue
        count = book_ped[int_idx]
        for i, (_, row) in enumerate(count.iterrows()):
            ped_input.SetAttValue(f'Volume({i+1})', row[int_pos])
        # ped routes
        ped_route = ped_routes.ItemByKey(int_no)
        route_decisions = ped_route.PedRoutSta
        for route_decision in route_decisions:
            movement = route_decision.AttValue('Name')
            # configure volumnes at each time interval
            # time interval: 1: 0-900, 2: 900-1800, 3:1800-2700, 4: 2700-3600, 5: 3600-4500, 6: 4500-MAX
            for i, (_, row) in enumerate(count.iterrows()):
                mvmnt_count = int(row[f'{int_pos}-{movement}'])
                route_decision.SetAttValue(f'RelFlow({i+1})', mvmnt_count)
    
    vissim.SaveNetAs(net_name)
    logger.info(f'Vehicle and pedestrian volumes imported!')
    

if __name__ == '__main__':
    main()