# Vissim project directory
PROJECT_NAME = '2024_PM_V8'
WORKING_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM' + f'\{PROJECT_NAME}'
PUA_FILE_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM' + f'\{PROJECT_NAME}' + r'\2024_Signals'
PROJECT_PERIOD = 'PM'

# simulation settings
EVAL_FROM_TIME = 0
EVAL_TO_TIME = 5400
EVAL_INTERVAL = 900
PERIOD_TIME = 5400  # simulation second [s]
STEP_TIME = 1
RANDOM_SEEDS = [30, 32, 34, 39, 42, 47, 49, 55, 56, 57]

# coordination parameters with offsets
COORD_SIGNAL_OFFSET = {
    47:  {48: 5},
    335: {60: 6},
    60:  {61: 5},
    61:  {62: 20},
    63:  {337: 3},  # 63, 70
    337: {336: 3},  # 70, 67
    336: {338: 3},  # 66, 70
}

CROSSING_NAMES = ['SG102', 'SG104', 'SG106', 'SG108']

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



