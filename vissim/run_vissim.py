import os

from vissim.utils import start_vissim

PROJECT_NAME = '2024_PM_V8'
WORKING_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM' + f'\{PROJECT_NAME}'
PUA_FILE_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM' + f'\{PROJECT_NAME}' + r'\2024_Signals'
RANDOM_SEEDS = [30, 32, 34, 39, 42, 47, 49, 55, 56, 57]
# evaluation time settings (in seconds)
EVAL_FROM_TIME = 0
EVAL_TO_TIME = 5400
EVAL_INTERVAL = 900
PERIOD_TIME = 5400  # simulation second [s]
STEP_TIME = 1

# coordination signals
COORD_SIGNALS = [[60, 335],
                 [60, 61],
                 [61, 62],
                 [63, 337, 338], # 67, 66
                 [63, 336], # 70,  # all NT/ST
                 ]

CROSSING_NAMES = ['SG102', 'SG104', 'SG106', 'SG108']


def get_stage_signal_groups_from_controller(sc_id):
    """
    Reads signal group information from a PUA file and organizes it by signal control stages.
    This function opens a PUA (VISSIM signal control) file associated with a given signal 
    controller ID and parses lines that define stages. Each stage line contains a stage name 
    followed by signal group identifiers. The function extracts and returns a mapping of 
    stage names to their corresponding signal groups.
    Args:
        sc_id (int or str): The signal controller ID used to locate the corresponding PUA file
                           (formatted as 'sig_{sc_id}.pua').
    Returns:
        dict: A dictionary where keys are stage names (e.g., 'stage_1', 'stage_2') and 
              values are lists of signal group identifiers for each stage. Returns an empty 
              dictionary if the PUA file does not exist.
    Note:
        - PUA files are expected to be located in the PUA_FILE_PATH directory.
        - Duplicate stage names are skipped (only the first occurrence is retained).
        - Empty strings resulting from split operations are filtered out from signal groups.
    """

    pua_file = os.path.join(PUA_FILE_PATH, f'sig_{sc_id}.pua')
    
    if not os.path.exists(pua_file):
        return {}
    
    stages = {}
    with open(pua_file, 'r') as f:
        for line in f:
            line = line.strip()
            if line.startswith('stage_'):
                stage_phases = line.split(' ')
                stage_name = stage_phases[0]
                if stage_name in stages: continue
                # skip the empty string and take the rest as signal groups
                signal_groups = [item for item in stage_phases[2:] if item]
                stages[stage_name] = signal_groups    
    return stages


def get_signal_group_no_by_name(sc, group_name):
    """
    Get the 'No' (ID) of a signal group by its name.
    
    Args:
        sc: The signal controller object (e.g., from vnet.SignalControllers.ItemByKey(id))
        group_name (str): The name of the signal group (e.g., 'NBTR', 'SBTR')
    
    Returns:
        int or None: The 'No' of the signal group if found, else None
    """
    signal_groups = sc.SignalGroups
    for sg in signal_groups:
        if sg.AttValue('Name') == group_name:
            return sg.AttValue('No')
    return None  # If not found


def get_active_signal_groups(sc):
    """
    Get a list of signal group names that are currently GREEN for a signal controller.
    
    Args:
        sc: The signal controller object (e.g., from vnet.SignalControllers.ItemByKey(id))
    
    Returns:
        list: List of signal group names (strings) that are GREEN
    """
    signal_groups = sc.SGs
    green_groups = []
    for sg in signal_groups:
        if sg.AttValue('SigState') == 'GREEN' or sg.AttValue('SigState') == 'AMBER':
            green_groups.append(sg.AttValue('Name'))
    return green_groups


def get_yellow_signal_groups(sc):
    """
    Get a list of signal group names that are currently GREEN for a signal controller.
    
    Args:
        sc: The signal controller object (e.g., from vnet.SignalControllers.ItemByKey(id))
    
    Returns:
        list: List of signal group names (strings) that are GREEN
    """
    signal_groups = sc.SGs
    yellow_groups = []
    for sg in signal_groups:
        if sg.AttValue('SigState') == 'AMBER':
            yellow_groups.append(sg.AttValue('Name'))
    return yellow_groups


def get_red_signal_groups(sc):
    """
    Get a list of signal group names that are currently GREEN for a signal controller.
    
    Args:
        sc: The signal controller object (e.g., from vnet.SignalControllers.ItemByKey(id))
    
    Returns:
        list: List of signal group names (strings) that are GREEN
    """
    signal_groups = sc.SGs
    yellow_groups = []
    for sg in signal_groups:
        if sg.AttValue('SigState') == 'RED':
            yellow_groups.append(sg.AttValue('Name'))
    return yellow_groups


def get_active_stage(sc_id, stages_all, green_groups_sc):
    # Get the current stage based on green groups
    current_stage = None
    if sc_id in stages_all:
        for stage, groups in stages_all[sc_id].items():
            if set(groups) == set(green_groups_sc):
                current_stage = stage
                return current_stage
    return None


def whether_stage_transition(yellow_groups_lead, active_groups_lead):
    yellow_signal_names = ['NBT', 'NBTR', 'NBTRL',
                           'SBT', 'SBTR', 'SBTRL',
                           'EBT', 'EBTR', 'EBTRL',
                           'WBT', 'WBTR', 'WBTRL']

    whether_yellow_crossing = any(group in CROSSING_NAMES for group in yellow_groups_lead)
    whether_yellow_signals = any(group in yellow_signal_names for group in yellow_groups_lead)

    if whether_yellow_crossing and whether_yellow_signals:
        return 'Crossing and Signal'
    elif whether_yellow_crossing and not whether_yellow_signals:
        return 'Crossing'
    elif not whether_yellow_crossing and not whether_yellow_signals:
        if len(active_groups_lead) > 0:
            return 'Not Transition'
        else:
            return 'All Red'


def coordinate_signal_stages(scs_coordinated, all_signal_controls, all_stages, stage_lead, whether_lead_transition):
    for coordinated in scs_coordinated:  # let the same stage operate
        sc_coord = all_signal_controls.ItemByKey(coordinated)
        signal_groups_coord = sc_coord.SGs
        try:
            # if it's all red
            if whether_lead_transition == 'All Red':
                for sg in signal_groups_coord:
                    sg_name = sg.AttValue('Name')
                    sg.SetAttValue('SigState', 'RED')
            else:
                stage_groups = all_stages[coordinated][stage_lead]
                # control signal phasing
                if whether_lead_transition == 'Crossing and Signal':
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups:
                            sg.SetAttValue('SigState', 'AMBER')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                elif whether_lead_transition == 'Crossing':
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups and sg_name in CROSSING_NAMES:
                            sg.SetAttValue('SigState', 'AMBER')
                        elif sg_name in stage_groups and not sg_name in CROSSING_NAMES:
                            sg.SetAttValue('SigState', 'GREEN')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                elif whether_lead_transition == 'Not Transition':
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups:
                            sg.SetAttValue('SigState', 'GREEN')
                        else:
                            sg.SetAttValue('SigState', 'RED')
        except:
            continue  # user's opening features during simualtion


def load_project():
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
    vissim.SaveNetAs(net_name)
    return vissim


def main():
    vissim = load_project()
    sim = vissim.Simulation
    vnet = vissim.Net # create Net COM-interface
    signal_controls = vnet.SignalControllers
    # get all stages
    signal_control_ids = [sc.AttValue('No') for sc in signal_controls]
    stages = {}
    for sc_id in signal_control_ids:
        stages[sc_id] = get_stage_signal_groups_from_controller(sc_id)

    for i, seed in enumerate(RANDOM_SEEDS):
        print(f'Running {i+1}th random seed {seed}...')
        # set the random seed for Vissim (not Python's random)
        sim.SetAttValue('RandSeed', seed)
        for sim_step in range(EVAL_FROM_TIME, EVAL_FROM_TIME+PERIOD_TIME*STEP_TIME+1):
            sim.RunSingleStep()
            # let signals coordinated
            for scs in COORD_SIGNALS:
                lead = scs[0]
                sc_lead = signal_controls.ItemByKey(lead)
                active_groups_lead = get_active_signal_groups(sc_lead)  # check what stage it is now
                yellow_groups_lead = get_yellow_signal_groups(sc_lead)  # check what stage it is now
                stage_lead = get_active_stage(lead, stages, active_groups_lead)
                whether_lead_transition = whether_stage_transition(yellow_groups_lead, active_groups_lead)
                if stage_lead == None and whether_lead_transition == 'Not Transition':
                    continue
                if stage_lead == None and whether_lead_transition == 'Crossing':
                    continue  # vissim configuration is not correct
                # run the same stage of the coordinated signal controllers
                if lead == 60 in scs[:1]:
                    if stage_lead == 'stage_6':
                        stage_lead = 'stage_2'
                elif lead == 63 and 336 in scs[1:]:
                    if stage_lead == 'stage_6':
                        stage_lead = 'stage_4'
                elif lead == 63 and 338 in scs[1:]:
                    if stage_lead == 'stage_5':
                        stage_lead = 'stage_3'
                    elif stage_lead == 'stage_6':
                        stage_lead = 'stage_4'
                coordinate_signal_stages(scs[1:], signal_controls, stages, stage_lead, whether_lead_transition)            

    vissim = None


if __name__ == '__main__':
    main()