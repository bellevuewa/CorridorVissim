import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from queue import Queue

from vissim.utils import load_project
from config_vissim import *


# Thread-safe locks for shared data
state_lock = Lock()


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
    Get a list of signal group names that are currently RED for a signal controller.
    
    Args:
        sc: The signal controller object (e.g., from vnet.SignalControllers.ItemByKey(id))
    
    Returns:
        list: List of signal group names (strings) that are RED
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


def get_signal_state_for_controller(sc):
    """Read signal state data from a controller (called on main thread only)."""
    try:
        active_groups = []
        yellow_groups = []
        
        for sg in sc.SGs:
            sig_state = sg.AttValue('SigState')
            sg_name = sg.AttValue('Name')
            
            if sig_state in ['GREEN', 'AMBER']:
                active_groups.append(sg_name)
            if sig_state == 'AMBER':
                yellow_groups.append(sg_name)
        
        return {
            'active_groups': active_groups,
            'yellow_groups': yellow_groups,
        }
    except:
        return {'active_groups': [], 'yellow_groups': []}


def compute_coordination_decision(lead_id, coordinated_id, lead_data, coord_data, stages,
                                 stage_transition_time, previous_stage, 
                                 coordinated_stage_start_time, offset, previous_stage_lead):
    """
    Compute coordination decision from pre-read signal data (can be done in parallel).
    No VISSIM COM access - only pure Python logic.
    
    Args:
        lead_id, coordinated_id: Controller IDs
        lead_data: Dict with 'active_groups' and 'yellow_groups' for lead signal
        coord_data: Dict with 'active_groups' and 'yellow_groups' for coordinated signal
        stages: Dictionary of all stages
        stage_transition_time: Dictionary tracking stage transition times
        previous_stage: Dictionary tracking previous stages
        coordinated_stage_start_time: Dictionary tracking coordination timing
        offset: Offset time in seconds
        previous_stage_lead: Dictionary tracking previous valid stage_lead for each signal
        
    Returns:
        dict: Actions to apply (signal group states to set)
    """
    try:
        active_groups_lead = lead_data['active_groups']
        yellow_groups_lead = lead_data['yellow_groups']
        active_groups_coord = coord_data['active_groups']
        
        stage_lead = get_active_stage(lead_id, stages, active_groups_lead)
        # Use previous stage_lead if current stage_lead is None
        if stage_lead is None:
            stage_lead = previous_stage_lead.get(lead_id)
        whether_lead_transition = whether_stage_transition(yellow_groups_lead, active_groups_lead)
        stage_coord = get_active_stage(coordinated_id, stages, active_groups_coord)
        
        # Check for stage transitions
        stage_changed_lead = (stage_lead != previous_stage.get(lead_id))
        stage_changed_coord = (stage_coord != previous_stage.get(coordinated_id))
        
        # Return computed decision without applying it yet
        return {
            'lead_id': lead_id,
            'coordinated_id': coordinated_id,
            'stage_lead': stage_lead,
            'stage_coord': stage_coord,
            'whether_lead_transition': whether_lead_transition,
            'stage_changed_lead': stage_changed_lead,
            'stage_changed_coord': stage_changed_coord,
            'offset': offset,
            'original_stage_lead': get_active_stage(lead_id, stages, active_groups_lead),  # Track original before fallback
        }
    except Exception as e:
        print(f"Error computing coordination decision for {lead_id}->{coordinated_id}: {e}")
        return None


def apply_coordination_decision(signal_controls, stages, decision, 
                                stage_transition_time, previous_stage, coordinated_stage_start_time, previous_stage_lead):
    """
    Apply coordination decisions on the main thread (BISSIM COM operations must happen here).
    
    Args:
        signal_controls: VISSIM signal controls object
        stages: Dictionary of all stages
        decision: Computed coordination decision from parallel thread
        stage_transition_time: Dictionary tracking stage transition times
        previous_stage: Dictionary tracking previous stages
        coordinated_stage_start_time: Dictionary tracking coordination timing
        previous_stage_lead: Dictionary tracking previous valid stage_lead for each signal
    """
    if decision is None:
        return
    
    try:
        lead_id = decision['lead_id']
        coordinated_id = decision['coordinated_id']
        stage_lead = decision['stage_lead']
        original_stage_lead = decision['original_stage_lead']
        stage_coord = decision['stage_coord']
        whether_lead_transition = decision['whether_lead_transition']
        
        # Update previous_stage_lead only if original_stage_lead was valid (not None)
        if original_stage_lead is not None:
            previous_stage_lead[lead_id] = original_stage_lead
        stage_changed_lead = decision['stage_changed_lead']
        stage_changed_coord = decision['stage_changed_coord']
        offset = decision['offset']
        
        # Update tracking for lead signal
        if stage_changed_lead:
            previous_stage[lead_id] = stage_lead
            if stage_lead:
                stage_transition_time[lead_id][stage_lead] = 0
            # Clear coordinated tracking when lead stage changes
            coords_to_remove = [key for key in coordinated_stage_start_time if key.startswith(f"{coordinated_id}_")]
            for key in coords_to_remove:
                del coordinated_stage_start_time[key]
        else:
            if stage_lead and stage_lead in stage_transition_time[lead_id]:
                stage_transition_time[lead_id][stage_lead] += STEP_TIME
        
        # Update tracking for coordinated signal
        if stage_changed_coord:
            previous_stage[coordinated_id] = stage_coord
            if stage_coord:
                stage_transition_time[coordinated_id][stage_coord] = 0
            has_active_coordination = any(
                key.startswith(f"{coordinated_id}_") and 
                not (isinstance(coordinated_stage_start_time[key], tuple) and 
                     coordinated_stage_start_time[key][0] == "COMPLETED")
                for key in coordinated_stage_start_time
            )
            if not has_active_coordination:
                coords_to_remove = [key for key in coordinated_stage_start_time if key.startswith(f"{coordinated_id}_")]
                for key in coords_to_remove:
                    del coordinated_stage_start_time[key]
        else:
            if stage_coord and stage_coord in stage_transition_time[coordinated_id]:
                stage_transition_time[coordinated_id][stage_coord] += STEP_TIME
        
        # Skip coordination in invalid states, but allow current state to persist (let VISSIM resume control)
        # Only completely skip if we're not forcing all-red right now
        if stage_lead is None:
            # When lead signal becomes undetectable, stop active coordination
            # to allow VISSIM's normal signal controller to resume

            return
        
        # Check if offset time has elapsed since lead signal changed to this stage
        stage_groups = stages[coordinated_id].get(stage_lead, [])
        
        if lead_id in stage_transition_time and stage_lead and stage_lead in stage_transition_time[lead_id]:
            time_since_lead_transition = stage_transition_time[lead_id][stage_lead]
            if time_since_lead_transition < offset:
                # During offset wait period
                if stage_coord == stage_lead:
                    # Already in target stage - keep GREEN during offset
                    sc_coord = signal_controls.ItemByKey(coordinated_id)
                    for sg in sc_coord.SGs:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups:
                            sg.SetAttValue('SigState', 'GREEN')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                    return
                # If not in target stage, fall through to start 4+1 transition
        
        # Offset elapsed or signal not in target stage - apply signal state changes
        # Apply signal state changes
        sc_coord = signal_controls.ItemByKey(coordinated_id)
        signal_groups_coord = sc_coord.SGs
        
        try:
            if whether_lead_transition == 'All Red':
                # Only apply all red when lead signal is in a valid detectable state
                for sg in signal_groups_coord:
                    sg.SetAttValue('SigState', 'RED')
            else:
                stage_groups = stages[coordinated_id].get(stage_lead, [])
                
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
                        elif sg_name in stage_groups and sg_name not in CROSSING_NAMES:
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
            pass
    except Exception as e:
        print(f"Error applying coordination decision: {e}")


def main():
    vissim = load_project(working_path=WORKING_PATH, project_name=PROJECT_NAME)
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
        stage_transition_time = {sc_id: {} for sc_id in signal_control_ids}
        previous_stage = {sc_id: None for sc_id in signal_control_ids}
        # Track previous valid stage_lead for each signal controller
        previous_stage_lead = {sc_id: None for sc_id in signal_control_ids}
        # Track when coordinated signals start transitioning to new stage
        coordinated_stage_start_time = {}
        
        # Thread pool for parallel state reading and decision computing
        # Use max_workers=len(COORD_SIGNAL_OFFSET) to parallelize across all coordination pairs
        max_workers = min(4, len(COORD_SIGNAL_OFFSET))  # Limit to 4 threads
        
        for sim_step in range(EVAL_FROM_TIME, EVAL_FROM_TIME+PERIOD_TIME*STEP_TIME+1):
            sim.RunSingleStep()
            
            # Read all signal states on main thread (VISSIM COM access must be here)
            signal_states = {}
            for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                sc_lead = signal_controls.ItemByKey(lead)
                signal_states[lead] = get_signal_state_for_controller(sc_lead)
                
                for coordinated, offset in coordinated_dict.items():
                    sc_coord = signal_controls.ItemByKey(coordinated)
                    signal_states[coordinated] = get_signal_state_for_controller(sc_coord)
            
            # Parallelize decision computing for all coordination pairs
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_pair = {}
                
                # Submit all coordination decision tasks to thread pool (pure data only)
                for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                    for coordinated, offset in coordinated_dict.items():
                        lead_data = signal_states[lead]
                        coord_data = signal_states[coordinated]
                        
                        future_decision = executor.submit(
                            compute_coordination_decision,
                            lead, coordinated, lead_data, coord_data, stages,
                            stage_transition_time, previous_stage,
                            coordinated_stage_start_time, offset, previous_stage_lead
                        )
                        future_to_pair[future_decision] = (lead, coordinated)
                
                # Collect results
                decisions = []
                for future in as_completed(future_to_pair):
                    lead, coordinated = future_to_pair[future]
                    decision = future.result()
                    if decision:
                        decisions.append((lead, decision))
            
            # Apply all decisions sequentially on main thread (required for VISSIM COM)
            with state_lock:
                for lead, decision in decisions:
                    apply_coordination_decision(
                        signal_controls, stages, decision,
                        stage_transition_time, previous_stage, coordinated_stage_start_time, previous_stage_lead
                    )
        
        vissim.SaveNet()
        # Stop simulation before setting new seed
        sim.Stop()
        time.sleep(2)  # Wait for simulation to fully stop
    
    vissim.Exit()
    vissim = None


if __name__ == '__main__':
    main()
