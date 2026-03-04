import os

from vissim.utils import start_vissim

PROJECT_NAME = '2024_PM_V8'
WORKING_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM' + f'\{PROJECT_NAME}'
PUA_FILE_PATH = r'I:\Modeling and Analysis Group\02_Model Applications\021ProjectModeling\021_Studies\156th_corridor_study\VISSIM' + f'\{PROJECT_NAME}' + r'\2024_Signals'
# simulation settings
EVAL_FROM_TIME = 0
EVAL_TO_TIME = 5400
EVAL_INTERVAL = 900
PERIOD_TIME = 5400  # simulation second [s]
STEP_TIME = 1
RANDOM_SEEDS = [30, 32, 34]#, 39, 42, 47, 49, 55, 56, 57]

# coordination parameters
COORD_OFFSET = [9, 
                5, 
                9, 
                7,
                7,
                8,
                ]  # offset in seconds for each coordinated signal group
# coordination signals
COORD_SIGNALS = [[335, 60],
                 [60, 61],
                 [61, 62],                 
                 [63, 337],
                 [337, 336],  # 70, 67
                 [336, 338],  # 66, 70
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


def enforce_adaptive_stage_timing(signal_controls, stages, stage_transition_time, previous_stage, lead_sc_id, coordinated_sc_ids, lead_stage_tracking):
    """
    Adaptive stage timing using PREVIOUS cycle's lead stage_C duration.
    
    - When lead completes stage_C, store its total duration
    - In NEXT cycle, use that stored duration to calculate half-duration
    - Coordinated signals run 2 stages, each for (previous_stage_C_duration / 2)
    
    Example:
    - Cycle 1: Lead runs stage_C for 20s → store as previous_duration
    - Cycle 2: Coordinated runs stage_C(10s) then stage_F(10s), based on previous 20s
    """
    try:
        # Get lead signal's current stage
        sc_lead = signal_controls.ItemByKey(lead_sc_id)
        active_groups_lead = get_active_signal_groups(sc_lead)
        current_lead_stage = get_active_stage(lead_sc_id, stages, active_groups_lead)
        
        # Detect when lead EXITS stage_C to capture and store final duration
        if lead_sc_id in previous_stage and previous_stage[lead_sc_id] == 'stage_C' and current_lead_stage != 'stage_C':
            # Lead just exited stage_C, store the duration for next cycle
            if lead_sc_id in stage_transition_time and 'stage_C' in stage_transition_time[lead_sc_id]:
                final_duration = stage_transition_time[lead_sc_id]['stage_C']
                # ONLY store the previous stage_C duration
                lead_stage_tracking[lead_sc_id] = final_duration
            return
        
        # Only enforce when lead is in stage_C
        if current_lead_stage != 'stage_C':
            return
        
        # Get the PREVIOUS stage_C duration from last cycle
        if lead_sc_id not in lead_stage_tracking:
            # First cycle, use current duration as fallback
            if lead_sc_id in stage_transition_time and 'stage_C' in stage_transition_time[lead_sc_id]:
                previous_duration = stage_transition_time[lead_sc_id]['stage_C']
            else:
                return
        else:
            # Use stored previous duration
            previous_duration = lead_stage_tracking[lead_sc_id]
        
        half_duration = previous_duration / 2.0
        
        # Enforce timing for coordinated signals based on PREVIOUS lead cycle's duration
        for coord_sc_id in coordinated_sc_ids:
            sc_coord = signal_controls.ItemByKey(coord_sc_id)
            active_groups_coord = get_active_signal_groups(sc_coord)
            current_coord_stage = get_active_stage(coord_sc_id, stages, active_groups_coord)
            
            if coord_sc_id not in stage_transition_time:
                continue
            
            if current_coord_stage not in stage_transition_time[coord_sc_id]:
                continue
            
            time_in_coord_stage = stage_transition_time[coord_sc_id][current_coord_stage]
            
            # Check if coordinated signal should transition at half-duration
            if time_in_coord_stage >= half_duration:
                # Get available stages for this signal
                available_stages = list(stages[coord_sc_id].keys())
                
                # Find next stage (prefer the one after current, or loop)
                current_idx = available_stages.index(current_coord_stage) if current_coord_stage in available_stages else 0
                next_stage_idx = (current_idx + 1) % len(available_stages)
                next_stage = available_stages[next_stage_idx]
                
                # Force transition to next stage
                stage_groups = stages[coord_sc_id].get(next_stage, [])
                signal_groups = sc_coord.SGs
                for sg in signal_groups:
                    sg_name = sg.AttValue('Name')
                    if sg_name in stage_groups:
                        sg.SetAttValue('SigState', 'GREEN')
                    else:
                        sg.SetAttValue('SigState', 'RED')
                
                # Reset timing for next stage
                previous_stage[coord_sc_id] = next_stage
                stage_transition_time[coord_sc_id][next_stage] = 0
    except:
        pass


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


def coordinate_signal_stages_with_offset(scs_coordinated, all_signal_controls, all_stages, stage_lead, 
                                         whether_lead_transition, offset_seconds, stage_transition_time, lead_signal_id, coordinated_stage_start_time):
    """
    Coordinate signal stages with a time offset.
    
    Args:
        scs_coordinated: List of coordinated signal controller IDs
        all_signal_controls: Signal controls object
        all_stages: Dictionary of stages for each controller
        stage_lead: Current stage of lead signal
        whether_lead_transition: Type of transition
        offset_seconds: Offset time in seconds
        stage_transition_time: Dictionary tracking when each stage transition occurred for each controller
        lead_signal_id: The ID of the lead signal
        coordinated_stage_start_time: Dictionary tracking when coordinated signal started transitioning to new stage
    """
    # Allow coordination even during transitions to ensure coordinated signals actually change
    # Only skip if stage_lead is None (undetectable state)
    if stage_lead is None:
        return
    
    for coordinated in scs_coordinated:
        sc_coord = all_signal_controls.ItemByKey(coordinated)
        signal_groups_coord = sc_coord.SGs
        
        # Check if offset time has elapsed since lead signal changed to this stage
        if lead_signal_id in stage_transition_time and stage_lead and stage_lead in stage_transition_time[lead_signal_id]:
            time_since_lead_transition = stage_transition_time[lead_signal_id][stage_lead]
            if time_since_lead_transition < offset_seconds:
                # During offset wait period, maintain the current stage (do nothing - let VISSIM control it)
                continue
        
        # Offset elapsed and lead signal is stable, now apply coordination
        try:
            stage_groups = all_stages[coordinated][stage_lead]
            
            # Check if this coordinated signal just started transitioning to the new stage
            coord_key = f"{coordinated}_{stage_lead}"
            if coord_key not in coordinated_stage_start_time:
                # First time transitioning to this stage after offset - start amber phase
                coordinated_stage_start_time[coord_key] = 0  # Track amber duration
                # Apply amber to indicate transition
                for sg in signal_groups_coord:
                    sg_name = sg.AttValue('Name')
                    if sg_name in stage_groups:
                        sg.SetAttValue('SigState', 'AMBER')
                    else:
                        sg.SetAttValue('SigState', 'RED')
            else:
                # Already in transition, track amber duration
                amber_duration = coordinated_stage_start_time[coord_key]
                amber_duration += STEP_TIME
                coordinated_stage_start_time[coord_key] = amber_duration
                
                # Maintain amber for 4 seconds, then all red for 1 second, then GREEN
                if amber_duration < 4:  # 4 second amber phase
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups:
                            sg.SetAttValue('SigState', 'AMBER')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                elif amber_duration < 5:  # 1 second all red phase
                    for sg in signal_groups_coord:
                        sg.SetAttValue('SigState', 'RED')
                else:
                    # All red phase complete, apply the actual stage (GREEN/RED)
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
        
        # Track stage transitions with offset
        # stage_transition_time[controller_id][stage_name] = elapsed_time_in_current_stage
        stage_transition_time = {sc_id: {} for sc_id in signal_control_ids}
        previous_stage = {sc_id: None for sc_id in signal_control_ids}
        # Track when coordinated signals start transitioning to new stage
        coordinated_stage_start_time = {}
        # Track lead signal stage durations for adaptive timing
        lead_stage_tracking = {}
        
        for sim_step in range(EVAL_FROM_TIME, EVAL_FROM_TIME+PERIOD_TIME*STEP_TIME+1):
            sim.RunSingleStep()
            # let signals coordinated
            for coord_idx, scs in enumerate(COORD_SIGNALS):
                lead = scs[0]
                sc_lead = signal_controls.ItemByKey(lead)
                active_groups_lead = get_active_signal_groups(sc_lead)  # check what stage it is now
                yellow_groups_lead = get_yellow_signal_groups(sc_lead)  # check what stage it is now
                stage_lead = get_active_stage(lead, stages, active_groups_lead)
                whether_lead_transition = whether_stage_transition(yellow_groups_lead, active_groups_lead)
                
                # Update stage transition tracking for LEAD signal
                if stage_lead != previous_stage[lead]:
                    # New stage detected
                    previous_stage[lead] = stage_lead
                    if stage_lead:
                        stage_transition_time[lead][stage_lead] = 0
                    # Clear coordinated_stage_start_time for all coordinated signals when lead stage changes
                    # This allows them to start fresh with the new lead stage (possibly modified)
                    coordinated_scs = scs[1:]
                    coords_to_remove = [key for key in coordinated_stage_start_time if any(key.startswith(f"{coord}_") for coord in coordinated_scs)]
                    for key in coords_to_remove:
                        del coordinated_stage_start_time[key]
                else:
                    # Same stage, increment elapsed time
                    if stage_lead and stage_lead in stage_transition_time[lead]:
                        stage_transition_time[lead][stage_lead] += STEP_TIME
                
                # Update stage transition tracking for all COORDINATED signals
                for coordinated in scs[1:]:
                    sc_coord = signal_controls.ItemByKey(coordinated)
                    active_groups_coord = get_active_signal_groups(sc_coord)
                    stage_coord = get_active_stage(coordinated, stages, active_groups_coord)
                    
                    if stage_coord != previous_stage[coordinated]:
                        # New stage detected for coordinated signal
                        previous_stage[coordinated] = stage_coord
                        if stage_coord:
                            stage_transition_time[coordinated][stage_coord] = 0
                        # Clear coordinated_stage_start_time entries for this signal when stage changes
                        # This ensures we track only the current stage transition
                        coords_to_remove = [key for key in coordinated_stage_start_time if key.startswith(f"{coordinated}_")]
                        for key in coords_to_remove:
                            del coordinated_stage_start_time[key]
                    else:
                        # Same stage, increment elapsed time
                        if stage_coord and stage_coord in stage_transition_time[coordinated]:
                            stage_transition_time[coordinated][stage_coord] += STEP_TIME
                
                if stage_lead == None and whether_lead_transition == 'Not Transition':
                    continue
                if stage_lead == None and whether_lead_transition == 'Crossing':
                    continue  # vissim configuration is not correct
                # run the same stage of the coordinated signal controllers
                if lead == 60:
                    if stage_lead == 'stage_6':
                        stage_lead = 'stage_4'
                elif lead == 62:
                    if stage_lead == 'stage_4':
                        stage_lead = 'stage_3'
                elif lead == 63:
                    if stage_lead == 'stage_4' or stage_lead == 'stage_6':
                        stage_lead = 'stage_3'
                    elif stage_lead == 'stage_5':
                        stage_lead = 'stage_4'
                elif lead == 337:
                    if stage_lead == 'stage_3':
                        stage_lead = 'stage_4'
                elif lead == 336:
                    if stage_lead == 'stage_4':
                        stage_lead = 'stage_3'
                    elif stage_lead == 'stage_5':
                        stage_lead = 'stage_4'
                
                # Apply coordination with offset for this signal group
                coordinate_signal_stages_with_offset(scs[1:], signal_controls, stages, stage_lead, 
                                                     whether_lead_transition, COORD_OFFSET[coord_idx], stage_transition_time, lead, coordinated_stage_start_time)
        vissim.SaveNet()
    
    vissim.Exit()
    vissim = None


if __name__ == '__main__':
    main()