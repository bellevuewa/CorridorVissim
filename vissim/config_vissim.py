# Vissim project directory
WORKING_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM\2024_PM_V1'
PROJECT_NAME = '2024_PM_V1'
PROJECT_PERIOD = 'PM'

# simulation settings
PERIOD_TIME = 5400  # simulation second [s]
START_TIME = '16:30:00'
STEP_TIME = 7  # Number of steps in a second
RANDOM_SEED_INCR = 1  # random seed increment
NUM_RUNS = 10  # desired total simulation runs
USE_MAX_SPEED = True  # if to use max speed
# if not USE_MAX_SPEED, set SIM_SPEED_FACTOR a value less than 1
SIM_SPEED_FACTOR = 1.0  # simulated speed factor
# if not USE_MAX_SPEED, set DYNA_ASSIGN_VOL_INCREMENT a value greater than 0
DYNA_ASSIGN_VOL_INCREMENT = 0.00  # dynamic assignment volume increment
USE_ALL_CORES = True  # if to use all cores

# evaluation time settings (in seconds)
EVAL_FROM_TIME = 0
EVAL_TO_TIME = 5400
EVAL_INTERVAL = 900

# evaluation measures
MEASURE_SYNTAX = {
    'Queue length'          : 'QLen(Current, Current)',
    'Queue length max'      : 'QLenMax(Current, Current)',
    'Vehicle stopped delay' : 'StopDelay(Current, Current, All)',
    'Vehicle delay'         : 'VehDelay(Current, Current, All)',
    'Stops'                 : 'Stops(Current, Current, All)',
    'Vehicles'              : 'Vehs(Current, Current, All)',
    'Ped delay'             : 'PersDelay(Current, Current, All)',
    'Peds'                  : 'Pers(Current, Current, All)'
}

# evaluation base settings
INTERSECTIONS = {
    1: {'Name': '110thAveNE-NE6thSt',  # key is node ID
        'Approach': {'NB': 1,
                     'SB': 2,},  # vehicle travel time measurement ID in Vissim
        'Movement':  {'NBT': [3, 10],
                      'NBR': [3, 7],
                      'NBL': [3, 1],}  # from link id and to link id
    }
}



# SCREELINES = {'Line_148th_In': 2}

# vehicle counts
COUNT_COLUMNS = ['Interval Start', 'EB-UT', 'EB-LT', 'EB-TH', 'EB-RT',
                                   'WB-UT', 'WB-LT', 'WB-TH', 'WB-RT',
                                   'NB-UT', 'NB-LT', 'NB-TH', 'NB-RT',
                                   'SB-UT', 'SB-LT', 'SB-TH', 'SB-RT']

# ped counts
PED_COLUMNS = ['Start', 'HV-EB', 'HV-WB', 'HV-NB', 'HV-SB', 'HV-Total',
                        'Bike-EB', 'Bike-WB', 'Bike-NB', 'Bike-SB', 'Bike-Total',
                        'Ped-EB', 'Ped-WB', 'Ped-NB', 'Ped-SB', 'Ped-Total',]

PED_POSITIONS = {'NW': ['NW-EB', 'NW-SB'], 
                 'NE': ['NE-WB', 'NE-SB'], 
                 'SE': ['SE-WB', 'SE-NB'], 
                 'SW': ['SW-NB', 'SW-EB']}
PED_SPLITS = {
    'EB': ('NW-EB', 'SW-EB'),
    'WB': ('NE-WB', 'SE-WB'),
    'NB': ('SW-NB', 'SE-NB'),
    'SB': ('NW-SB', 'NE-SB')
}

intersection_list_project = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\Analysis\Study_Intersection.xlsx'
intersection_list_all = r'J:\Traffic Data Program\2_Turning Movement Counts (TMC)\Intersection Numbering List.xls'
dir_vehicle_count = {
    2023: r'J:\Traffic Data Program\2_Turning Movement Counts (TMC)\2023',
    2024: r'J:\Traffic Data Program\2_Turning Movement Counts (TMC)\2024',
    2025: r'J:\Traffic Data Program\2_Turning Movement Counts (TMC)\2025\Fall_2025',
    '2025Special': r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\Data\TMC',
    '2023Redmond': r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\Data\TMC\Redmond'
}



