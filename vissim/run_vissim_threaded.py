import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

from vissim.utils import load_project
from config_vissim import *


# Thread-safe locks for shared data
state_lock = Lock()


def get_stage_signal_groups_from_pua(sc_id):
    """
    Reads signal group information from a PUA file and organizes it by signal control stages.
    This function opens a PUA (VISSIM signal control) file associated with a given signal 
    controller ID and parses stages, interstage transitions, and active phases.
    
    Args:
        sc_id (int or str): The signal controller ID used to locate the corresponding PUA file
                           (formatted as 'sig_{sc_id}.pua').
    
    Returns:
        dict: A dictionary with the following structure:
        {
            'stages': {
                'stage_name': {
                    'active_phases': [list of active signal groups],
                    'red_phases': [list of red signal groups]
                },
                ...
            },
            'interstages': [
                {
                    'number': int,
                    'length': float (seconds),
                    'from_stage': 'stage_name',
                    'to_stage': 'stage_name',
                    'active_phases': {signal_group: (start_time, end_time), ...}
                },
                ...
            ]
        }
        Returns an empty dict if the PUA file does not exist.
    
    Note:
        - PUA files are expected to be located in the PUA_FILE_PATH directory.
        - Duplicate stage names are skipped (only the first occurrence is retained).
        - Interstage info includes transitions with timing details for each signal group.
    """

    pua_file = os.path.join(PUA_FILE_PATH, f'sig_{sc_id}.pua')
    
    if not os.path.exists(pua_file):
        return {}
    
    result = {
        'stages': {},
        'interstages': []
    }
    
    with open(pua_file, 'r') as f:
        lines = f.readlines()
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Parse STAGES section
        if line.startswith('stage_') and not line.startswith('$'):
            parts = line.split()
            stage_name = parts[0]
            
            # Skip if already parsed
            if stage_name in result['stages']:
                i += 1
                continue
            
            # Active phases are in parts[1:]
            active_phases = [p for p in parts[1:] if p]
            
            # Next line should be "red" with red phases
            red_phases = []
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if next_line.startswith('red'):
                    red_parts = next_line.split()
                    red_phases = [p for p in red_parts[1:] if p]
                    i += 1
            
            result['stages'][stage_name] = {
                'active_phases': active_phases,
                'red_phases': red_phases
            }
        
        # Parse INTERSTAGE section
        elif line.startswith('$INTERSTAGE'):
            interstage_info = {
                'number': None,
                'length': None,
                'from_stage': None,
                'to_stage': None,
                'active_phases': {}
            }
            
            i += 1
            # Parse interstage metadata
            while i < len(lines):
                line = lines[i].strip()
                
                if line.startswith('INTERSTAGE_number'):
                    parts = line.split(':')
                    interstage_info['number'] = int(parts[1].strip())
                elif line.startswith('length [s]'):
                    parts = line.split(':')
                    interstage_info['length'] = float(parts[1].strip())
                elif line.startswith('from stage'):
                    parts = line.split(':')
                    interstage_info['from_stage'] = f"stage_{parts[1].strip()}"
                elif line.startswith('to stage'):
                    parts = line.split(':')
                    interstage_info['to_stage'] = f"stage_{parts[1].strip()}"
                elif line.startswith('$'):
                    # End of this interstage section
                    break
                elif line and not any(line.startswith(x) for x in ['INTERSTAGE_number', 'length', 'from', 'to']):
                    # This is an active phase line with timing info
                    parts = line.split()
                    if len(parts) >= 3 and parts[0] not in ['INTERSTAGE_number', 'length', 'from', 'to']:
                        signal_group = parts[0]
                        start_time = int(parts[1])
                        end_time = int(parts[2])
                        interstage_info['active_phases'][signal_group] = (start_time, end_time)
                
                i += 1
            
            if interstage_info['number'] is not None:
                result['interstages'].append(interstage_info)
            continue
        
        i += 1
    
    return result


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


def get_active_stage(sc_id, stages_all, active_groups_sc, yellow_groups_sc):
    """
    Determine current stage and interstage state from signal groups.
    
    Args:
        sc_id: Signal controller ID
        stages_all: Dictionary of all stages and interstages
        active_groups_sc: List of active (GREEN or AMBER) signal groups
        yellow_groups_sc: List of yellow (AMBER) signal groups
    
    Returns:
        dict: {
            'stage': 'stage_name' or None,
            'interstage': {
                'number': int,
                'from_stage': 'stage_name',
                'to_stage': 'stage_name',
                'length': float
            } or None
        }
    """
    result = {
        'stage': None,
        'interstage': None
    }
    
    if sc_id not in stages_all:
        return result
    
    stage_info = stages_all[sc_id]
    
    # Try to find current stage from active groups (treating yellow/AMBER as green)
    if 'stages' in stage_info:
        # Combine active groups and yellow groups for stage matching
        all_active_groups = set(active_groups_sc + yellow_groups_sc)
        for stage_name, stage_data in stage_info['stages'].items():
            stage_active_phases = set(stage_data.get('active_phases', []))
            if stage_active_phases == all_active_groups:
                result['stage'] = stage_name
                return result
    
    # If no stage match, we're in an interstage transition
    if 'interstages' in stage_info and yellow_groups_sc:
        # Find which interstage we're in based on yellow/amber signals
        for interstage in stage_info['interstages']:
            interstage_signals = set(interstage.get('active_phases', {}).keys())
            if interstage_signals & set(yellow_groups_sc):  # If there's overlap
                result['interstage'] = {
                    'number': interstage['number'],
                    'from_stage': interstage['from_stage'],
                    'to_stage': interstage['to_stage'],
                    'length': interstage['length']
                }
                return result
    
    return result


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


def get_next_stage_with_vehicles(current_stage, stages_info, vnet, sc_id):
    """
    Find the next stage with vehicles by checking detector occupancy.
    Iterates through stages sequentially (stage_i+1, stage_i+2, ...) until finding one with vehicles.
    If no stages have vehicles, returns stage_1. If a stage doesn't exist, skips to the next.
    
    Args:
        current_stage: Current stage name (e.g., 'stage_1')
        stages_info: Dictionary of stage information from the PUA file
        vnet: VISSIM network object for detector access
        sc_id: Signal controller ID
    
    Returns:
        str: Next stage name with vehicles, or 'stage_1' as fallback
    """
    if 'stages' not in stages_info:
        return 'stage_1'
    
    # Extract stage number from current stage name
    try:
        stage_num = int(current_stage.split('_')[1])
    except (IndexError, ValueError):
        return 'stage_1'
    
    available_stages = sorted([s for s in stages_info['stages'].keys() if s.startswith('stage_')],
                             key=lambda x: int(x.split('_')[1]))
    
    if not available_stages:
        return 'stage_1'
    
    max_stage_num = int(available_stages[-1].split('_')[1])
    
    # Try stages starting from stage_{i+1}
    for next_stage_num in range(stage_num + 1, max_stage_num + 2):
        next_stage = f'stage_{next_stage_num}'
        
        # If this stage doesn't exist, try the next one
        if next_stage not in stages_info['stages']:
            # If we've exceeded available stages, wrap to stage_1
            if next_stage_num > max_stage_num:
                return 'stage_1'
            continue
        
        # Check if any detectors for this stage have vehicles
        try:
            # Try to find detectors associated with this stage
            # Detector naming convention: might be "{sc_id}_{stage_name}" or similar
            detectors = vnet.Detectors
            for det in detectors:
                det_name = det.AttValue('Name')
                # Check if detector name contains stage number
                if str(next_stage_num) in det_name or next_stage in det_name:
                    det_count = det.AttValue('Count')
                    if det_count and det_count > 0:
                        return next_stage
        except:
            # If we can't access detectors, just return this stage as a fallback
            return next_stage
    
    # If no stages have vehicles, return stage_1
    return 'stage_1'


def get_all_stage_occupancy(signal_control_ids, detector_stage_map, vnet):
    """
    Pre-calculate stage occupancy for all signal controllers on the main thread.
    
    Args:
        signal_control_ids: List of signal controller IDs
        detector_stage_map: Dictionary mapping {sc_id: {stage_name: [detector_names]}}
        vnet: VISSIM network object for detector access
    
    Returns:
        dict: {sc_id: {stage_name: occupancy_value}}
    """
    stage_occupancy = {}
    detectors = vnet.Detectors
    
    try:
        for sc_id in signal_control_ids:
            stage_occupancy[sc_id] = {}
            
            if sc_id not in detector_stage_map:
                continue
            
            # For each stage, sum the occupancy of all its detectors
            for stage_name, detector_nos_list in detector_stage_map[sc_id].items():
                total_occupancy = 0
                detector_count = 0
                
                # detector_nos_list is a list of detector No lists
                for detector_nos in detector_nos_list:
                    if isinstance(detector_nos, list):
                        for det_no in detector_nos:
                            try:
                                # Access detector by its No (ID) using ItemByKey
                                det = detectors.ItemByKey(det_no)
                                occupancy = det.AttValue('Occ')
                                if occupancy:
                                    total_occupancy += occupancy
                                    detector_count += 1
                            except:
                                pass
                    else:
                        try:
                            # Access detector by its No (ID) using ItemByKey
                            det = detectors.ItemByKey(detector_nos)
                            occupancy = det.AttValue('Occ')
                            if occupancy:
                                total_occupancy += occupancy
                                detector_count += 1
                        except:
                            pass
                
                # Store average occupancy for this stage
                if detector_count > 0:
                    stage_occupancy[sc_id][stage_name] = total_occupancy / detector_count
                else:
                    stage_occupancy[sc_id][stage_name] = 0
        
        return stage_occupancy
    
    except Exception as e:
        print(f"Error calculating stage occupancy: {e}")
        return {}


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


def get_all_detectors(vnet):
    """
    Retrieve all detectors from VISSIM network organized by signal controller ID and port number.
    Uses detector 'No' (ID) as the identifier.
    
    Args:
        vnet: VISSIM network object (from vissim.Net)
    
    Returns:
        dict: Dictionary with structure {sc_id: {port_number: [detector_nos]}}
              where sc_id is the signal controller ID, port_number is the detector's lane number,
              and detector_nos is a list of detector No (ID) values on that lane for that controller.
    """
    detector_dict = {}
    
    try:
        detectors = vnet.Detectors
        
        # Get all detectors
        for detector in detectors:
            det_no = detector.AttValue('No')  # Get detector No (ID)
            det_sc = detector.AttValue('SC')
            det_port = detector.AttValue('PortNo')
            
            sc_id = int(det_sc)
            
            # Initialize nested dict if needed
            if sc_id not in detector_dict:
                detector_dict[sc_id] = {}
            
            # Add detector No with port number to the list
            if det_port in detector_dict[sc_id]:
                detector_dict[sc_id][det_port].append(det_no)
            else:
                detector_dict[sc_id][det_port] = [det_no]
        
        return detector_dict
    
    except Exception as e:
        print(f"Error retrieving detectors: {e}")
        return {}


def get_best_available_stage(coordinated_id, preferred_stage, available_stages, stage_occupancy):
    """
    Select the best available stage based on occupancy data.
    Prioritizes preferred_stage if available, otherwise selects stage with highest occupancy.
    
    Args:
        coordinated_id: The coordinated signal controller ID
        preferred_stage: The preferred stage (e.g., from remapping logic)
        available_stages: Collection of available stage names for this signal
        stage_occupancy: Dict of occupancy per stage {sc_id: {stage_name: occupancy_value}}
    
    Returns:
        str: Best stage name, or 'stage_1' as fallback
    """
    if not available_stages:
        return 'stage_1'
    
    # If preferred stage exists and has occupancy data, use it
    if preferred_stage and preferred_stage in available_stages:
        if coordinated_id in stage_occupancy:
            if preferred_stage in stage_occupancy[coordinated_id]:
                return preferred_stage
    
    # Otherwise, find the stage with highest occupancy
    if coordinated_id in stage_occupancy:
        coord_occupancy = stage_occupancy[coordinated_id]
        best_stage = None
        best_occupancy = -1
        
        for stage in available_stages:
            if stage in coord_occupancy:
                occ = coord_occupancy[stage]
                if occ > best_occupancy:
                    best_occupancy = occ
                    best_stage = stage
        
        if best_stage:
            return best_stage
    
    # Fallback: return first available stage or stage_1
    if 'stage_1' in available_stages:
        return 'stage_1'
    return list(available_stages)[0] if available_stages else 'stage_1'


def get_coordinated_stage(lead_id, stage):
    """
    Apply intelligent stage remapping based on lead signal ID.
    Maps detected stages from lead signal to coordinated stages.
    
    Args:
        lead_id: The lead signal controller ID
        stage: The stage name from the lead signal (e.g., 'stage_1')
    
    Returns:
        str: The remapped stage for the coordinated signal
    """
    target_stage = stage
    
    if lead_id == 47:
        if stage == 'stage_2':
            target_stage = 'stage_1'
        elif stage == 'stage_5':
            target_stage = 'stage_6'
    elif lead_id == 60:
        if stage == 'stage_6':
            target_stage = 'stage_4'
    elif lead_id == 63:
        if stage in ['stage_4', 'stage_6']:
            target_stage = 'stage_3'
        elif stage == 'stage_5':
            target_stage = 'stage_4'
    elif lead_id == 337:
        if stage == 'stage_3':
            target_stage = 'stage_4'
        elif stage == 'stage_4':
            target_stage = 'stage_5'
    elif lead_id == 336:
        if stage == 'stage_4':
            target_stage = 'stage_3'
        elif stage == 'stage_5':
            target_stage = 'stage_4'
    
    return target_stage


def compute_coordination_decision(lead_id, coordinated_id, lead_data, coord_data,
                                 previous_stage, offset, stage_occupancy):
    """
    Compute coordination decision from pre-read signal data (can be done in parallel).
    No VISSIM COM access - only pure Python logic.

    lead_data and coord_data must include a 'stage_info' key pre-computed by the
    main thread (via get_active_stage) alongside 'active_groups' and 'yellow_groups'.
    
    Args:
        lead_id, coordinated_id: Controller IDs
        lead_data: Signal state dict for lead (active_groups, yellow_groups, stage_info)
        coord_data: Signal state dict for coordinated (active_groups, yellow_groups, stage_info)
        previous_stage: Dictionary tracking previous stages
        offset: Offset time in seconds
        stage_occupancy: Dict of precalculated occupancy per stage {sc_id: {stage_name: occupancy_value}}
        
    Returns:
        dict: Actions to apply (signal group states to set)
    """
    try:
        stage_info_lead = lead_data['stage_info']
        stage_lead = stage_info_lead['stage']
        interstage_lead = stage_info_lead['interstage']
        original_stage_lead = stage_lead  # save before fallback may change stage_lead
        
        # If stage_lead is None and interstage is None, use detector occupancy to select stage
        if stage_lead is None and interstage_lead is None:
            if coordinated_id in stage_occupancy:
                # Find stage with highest occupancy
                coord_occupancy = stage_occupancy[coordinated_id]
                if coord_occupancy:
                    stage_lead = max(coord_occupancy, key=coord_occupancy.get)
                else:
                    stage_lead = 'stage_1'
            else:
                stage_lead = 'stage_1'
        
        stage_coord = coord_data['stage_info']['stage']
        
        # Check for stage transitions
        stage_changed_lead = (stage_lead != previous_stage.get(lead_id))
        stage_changed_coord = (stage_coord != previous_stage.get(coordinated_id))
        
        # Return computed decision without applying it yet
        return {
            'lead_id': lead_id,
            'coordinated_id': coordinated_id,
            'stage_lead': stage_lead,
            'stage_coord': stage_coord,
            'interstage_lead': interstage_lead,
            'stage_changed_lead': stage_changed_lead,
            'stage_changed_coord': stage_changed_coord,
            'offset': offset,
            'original_stage_lead': original_stage_lead,
        }
    except Exception as e:
        print(f"Error computing coordination decision for {lead_id}->{coordinated_id}: {e}")
        return None


def apply_coordination_decision(signal_controls, stages, decision, 
                                stage_transition_time, previous_stage, coordinated_stage_start_time, previous_stage_lead, stage_occupancy=None, vnet=None):
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
        stage_occupancy: Dict of precalculated occupancy per stage {sc_id: {stage_name: occupancy_value}}
        vnet: VISSIM network object for detector access
    """
    if decision is None:
        return
    
    try:
        lead_id = decision['lead_id']
        coordinated_id = decision['coordinated_id']
        stage_lead = decision['stage_lead']
        interstage_lead = decision['interstage_lead']
        original_stage_lead = decision['original_stage_lead']
        stage_coord = decision['stage_coord']
        
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
        
        # Check if any coordinated signals have in-progress transitions (in amber/all-red phase)
        has_in_progress_transition = any(
            key for key in coordinated_stage_start_time 
            if not (isinstance(coordinated_stage_start_time[key], tuple) and 
                    coordinated_stage_start_time[key][0] == "COMPLETED")
        )
        
        # Only skip coordination if stage_lead is None AND no in-progress transitions need to complete
        # This allows ongoing amber+all-red phases to finish even when lead signal becomes undetectable
        if stage_lead is None and not has_in_progress_transition:
            return
        
        sc_coord = signal_controls.ItemByKey(coordinated_id)
        signal_groups_coord = sc_coord.SGs
        
        try:
            # Case 0: Lead signal has both current stage and interstage (transitioning with stage info)
            if stage_lead is not None and interstage_lead is not None:
                # Get the next stage with vehicles, starting from the interstage's target stage
                target_stage = interstage_lead['to_stage']
                
                # Use detector-based selection to find the next stage with vehicles
                try:
                    if vnet is not None and stage_lead in stages[lead_id].get('stages', {}):
                        # Extract stage number from stage_lead
                        try:
                            stage_num = int(stage_lead.split('_')[1])
                            target_stage = get_next_stage_with_vehicles(
                                stage_lead, stages[coordinated_id], vnet, coordinated_id
                            )
                        except (IndexError, ValueError):
                            # If parsing fails, check if target stage exists, otherwise use stage_1
                            if target_stage not in stages[coordinated_id].get('stages', {}):
                                target_stage = 'stage_1'
                    else:
                        # If target stage doesn't exist in coordinated signal, find next with vehicles or use stage_1
                        if target_stage not in stages[coordinated_id].get('stages', {}):
                            target_stage = 'stage_1'
                except:
                    # Fallback: use target stage or stage_1
                    if target_stage not in stages[coordinated_id].get('stages', {}):
                        target_stage = 'stage_1'
                
                # Only apply if coordinated signal is in transition or not in target stage
                if stage_coord != target_stage:
                    stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage, {}).get('active_phases', [])
                    
                    # Get the current stage groups for transition
                    current_stage_groups = []
                    if stage_coord and stage_coord in stages[coordinated_id].get('stages', {}):
                        current_stage_groups = stages[coordinated_id]['stages'][stage_coord].get('active_phases', [])
                    
                    # Start 4+1 transition to target stage
                    coord_key = f"{coordinated_id}_{target_stage}"
                    if coord_key not in coordinated_stage_start_time:
                        # If coordinated signal is already in all-red, skip amber phase (nothing to clear)
                        initial_timer = 4 if stage_coord is None else 0
                        coordinated_stage_start_time[coord_key] = (initial_timer, current_stage_groups)
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if sg_name in current_stage_groups:
                                sg.SetAttValue('SigState', 'AMBER')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                    else:
                        tracking_data = coordinated_stage_start_time[coord_key]
                        
                        if isinstance(tracking_data, tuple) and tracking_data[0] == "COMPLETED":
                            # Maintain GREEN on target stage
                            for sg in signal_groups_coord:
                                sg_name = sg.AttValue('Name')
                                if sg_name in stage_groups:
                                    sg.SetAttValue('SigState', 'GREEN')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                        else:
                            # Continue transition timing
                            if isinstance(tracking_data, tuple):
                                amber_duration = tracking_data[0]
                                stored_stage_groups = tracking_data[1]
                            else:
                                amber_duration = tracking_data
                                stored_stage_groups = current_stage_groups
                            
                            amber_duration += STEP_TIME
                            coordinated_stage_start_time[coord_key] = (amber_duration, stored_stage_groups)
                            
                            if amber_duration < 4:  # AMBER phase
                                for sg in signal_groups_coord:
                                    sg_name = sg.AttValue('Name')
                                    if sg_name in stored_stage_groups:
                                        sg.SetAttValue('SigState', 'AMBER')
                                    else:
                                        sg.SetAttValue('SigState', 'RED')
                            elif amber_duration < 5:  # All-red phase
                                for sg in signal_groups_coord:
                                    sg.SetAttValue('SigState', 'RED')
                            else:
                                # Transition complete, switch to GREEN on target stage
                                for sg in signal_groups_coord:
                                    sg_name = sg.AttValue('Name')
                                    if sg_name in stage_groups:
                                        sg.SetAttValue('SigState', 'GREEN')
                                    else:
                                        sg.SetAttValue('SigState', 'RED')
                                coordinated_stage_start_time[coord_key] = ("COMPLETED", 0)
                else:
                    # Already in target stage
                    stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage, {}).get('active_phases', [])
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups:
                            sg.SetAttValue('SigState', 'GREEN')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                return
            
            # Case 1: Lead signal is in a defined stage (not transitioning)
            if stage_lead is not None:
                # If original_stage_lead was None (lead undetectable), use occupancy-selected stage without remapping
                # This allows responsive adaptation to traffic changes instead of rigid 4+1 timing
                if original_stage_lead is None:
                    # Check if signal is stuck in all-red - if so, force it out immediately
                    if stage_coord is None or stage_coord == '':
                        # Signal in all-red - use intelligent occupancy selection to escape, bypass minimum duration
                        available_stages = stages[coordinated_id].get('stages', {}).keys()
                        if stage_occupancy and coordinated_id in stage_occupancy:
                            target_stage_lead = get_best_available_stage(coordinated_id, stage_lead, available_stages, stage_occupancy)
                        else:
                            target_stage_lead = stage_lead if stage_lead else 'stage_1'
                        
                        stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage_lead, {}).get('active_phases', [])
                        
                        # Force transition out of all-red - set signals to GREEN on target stage
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if sg_name in stage_groups:
                                sg.SetAttValue('SigState', 'GREEN')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                        
                        # Initialize 10-second duration tracking for the escaped stage
                        stage_duration_key = f"{coordinated_id}_occupancy_stage_duration"
                        coordinated_stage_start_time[stage_duration_key] = (0, target_stage_lead)
                        return
                    
                    # Signal is in a normal stage - enforce minimum stage duration (10 seconds)
                    min_stage_duration = 10.0  # seconds
                    stage_duration_key = f"{coordinated_id}_occupancy_stage_duration"
                    
                    # Track when current stage started
                    if stage_duration_key not in coordinated_stage_start_time:
                        # First time or stage just changed - record start time and stage
                        coordinated_stage_start_time[stage_duration_key] = (0, stage_coord)
                    
                    tracking_data = coordinated_stage_start_time[stage_duration_key]
                    if isinstance(tracking_data, tuple):
                        elapsed_time, current_tracked_stage = tracking_data
                    else:
                        elapsed_time = tracking_data
                        current_tracked_stage = stage_coord
                    
                    elapsed_time += STEP_TIME
                    
                    # If current stage has not run for minimum duration, stick with it
                    if elapsed_time < min_stage_duration and stage_coord == current_tracked_stage:
                        # Keep current stage - don't switch yet
                        target_stage_lead = stage_coord
                        coordinated_stage_start_time[stage_duration_key] = (elapsed_time, current_tracked_stage)
                    else:
                        # 10-second transition point: use intelligent occupancy selection, not random stage_lead
                        # This avoids oscillating between stages if occupancy keeps changing
                        available_stages = stages[coordinated_id].get('stages', {}).keys()
                        if stage_occupancy and coordinated_id in stage_occupancy:
                            target_stage_lead = get_best_available_stage(coordinated_id, stage_lead, available_stages, stage_occupancy)
                        else:
                            target_stage_lead = stage_lead if stage_lead else 'stage_1'
                        # Reset duration tracking for next stage
                        coordinated_stage_start_time[stage_duration_key] = (0, target_stage_lead)
                    
                    stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage_lead, {}).get('active_phases', [])
                    
                    # Direct transition to selected stage - no offset, no amber phase
                    if stage_coord != target_stage_lead:
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if sg_name in stage_groups:
                                sg.SetAttValue('SigState', 'GREEN')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                    else:
                        # Already in target stage - maintain GREEN
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if sg_name in stage_groups:
                                sg.SetAttValue('SigState', 'GREEN')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                    return
                
                # Original lead stage was detected - apply remapping and use 4+1 offset-based coordination
                # Apply intelligent stage remapping based on lead signal ID
                target_stage_lead = get_coordinated_stage(lead_id, stage_lead)
                
                # Validate that target stage exists in coordinated signal, fallback to stage_lead if not
                if target_stage_lead not in stages[coordinated_id].get('stages', {}):
                    target_stage_lead = stage_lead
                
                # If still invalid, use stage_1
                if target_stage_lead not in stages[coordinated_id].get('stages', {}):
                    target_stage_lead = 'stage_1'
                
                stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage_lead, {}).get('active_phases', [])
                
                # Check if offset time has elapsed
                if lead_id in stage_transition_time and stage_lead in stage_transition_time[lead_id]:
                    time_since_lead_transition = stage_transition_time[lead_id][stage_lead]
                    if time_since_lead_transition < offset:
                        # During offset wait period - keep coordinated signal in current stage
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if stage_coord == target_stage_lead:
                                # Already in target stage - keep GREEN
                                if sg_name in stage_groups:
                                    sg.SetAttValue('SigState', 'GREEN')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                            else:
                                # Not yet in target stage - stay GREEN on current stage during offset
                                # so that stage_coord is valid when the 4+1 transition starts
                                current_stage_coord_groups = []
                                if stage_coord and stage_coord in stages[coordinated_id].get('stages', {}):
                                    current_stage_coord_groups = stages[coordinated_id]['stages'][stage_coord].get('active_phases', [])
                                if sg_name in current_stage_coord_groups:
                                    sg.SetAttValue('SigState', 'GREEN')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                        return
                
                # Offset elapsed - transition coordinated signal to target_stage_lead
                if stage_coord == target_stage_lead:
                    # Already in target stage, just maintain it
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in stage_groups:
                            sg.SetAttValue('SigState', 'GREEN')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                    # Mark as completed
                    coord_key = f"{coordinated_id}_{target_stage_lead}"
                    coordinated_stage_start_time[coord_key] = ("COMPLETED", 0)
                    return
                
                # Get the current stage groups for transition
                current_stage_groups = []
                if stage_coord and stage_coord in stages[coordinated_id].get('stages', {}):
                    current_stage_groups = stages[coordinated_id]['stages'][stage_coord].get('active_phases', [])
                
                # Start 4+1 transition to target stage
                coord_key = f"{coordinated_id}_{target_stage_lead}"
                if coord_key not in coordinated_stage_start_time:
                    # If coordinated signal is already in all-red, skip amber phase (nothing to clear)
                    initial_timer = 4 if stage_coord is None else 0
                    coordinated_stage_start_time[coord_key] = (initial_timer, current_stage_groups)
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in current_stage_groups:
                            sg.SetAttValue('SigState', 'AMBER')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                else:
                    tracking_data = coordinated_stage_start_time[coord_key]
                    
                    if isinstance(tracking_data, tuple) and tracking_data[0] == "COMPLETED":
                        # Maintain GREEN on target stage
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if sg_name in stage_groups:
                                sg.SetAttValue('SigState', 'GREEN')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                    else:
                        # Continue transition timing
                        if isinstance(tracking_data, tuple):
                            amber_duration = tracking_data[0]
                            stored_stage_groups = tracking_data[1]
                        else:
                            amber_duration = tracking_data
                            stored_stage_groups = current_stage_groups
                        
                        amber_duration += STEP_TIME
                        coordinated_stage_start_time[coord_key] = (amber_duration, stored_stage_groups)
                        
                        if amber_duration < 4:  # AMBER phase
                            for sg in signal_groups_coord:
                                sg_name = sg.AttValue('Name')
                                if sg_name in stored_stage_groups:
                                    sg.SetAttValue('SigState', 'AMBER')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                        elif amber_duration < 5:  # All-red phase
                            for sg in signal_groups_coord:
                                sg.SetAttValue('SigState', 'RED')
                        else:
                            # Transition complete, switch to GREEN on target stage
                            for sg in signal_groups_coord:
                                sg_name = sg.AttValue('Name')
                                if sg_name in stage_groups:
                                    sg.SetAttValue('SigState', 'GREEN')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                            coordinated_stage_start_time[coord_key] = ("COMPLETED", 0)
            
            # Case 2: Lead signal is in an interstage transition
            elif interstage_lead is not None:
                from_stage = interstage_lead['from_stage']
                to_stage = interstage_lead['to_stage']
                interstage_length = interstage_lead['length']
                
                # Skip if from_stage and to_stage are the same (no real transition)
                if from_stage == to_stage:
                    return
                
                # Apply intelligent stage remapping based on lead signal ID
                target_stage = get_coordinated_stage(lead_id, to_stage)
                
                # Find matching interstage in coordinated signal by target_stage
                matching_interstage = None
                if 'interstages' in stages[coordinated_id]:
                    for interstage in stages[coordinated_id]['interstages']:
                        if interstage['to_stage'] == target_stage:
                            matching_interstage = interstage
                            break
                
                # If remapped target stage not found, try original to_stage as fallback
                if matching_interstage is None and target_stage != to_stage:
                    if 'interstages' in stages[coordinated_id]:
                        for interstage in stages[coordinated_id]['interstages']:
                            if interstage['to_stage'] == to_stage:
                                matching_interstage = interstage
                                target_stage = to_stage
                                break
                
                if matching_interstage is None:
                    return
                
                # Check if offset time has elapsed since lead entered this interstage
                coord_key = f"{coordinated_id}_interstage_{from_stage}_{target_stage}"
                
                if coord_key not in coordinated_stage_start_time:
                    # First time - record start but wait for offset
                    coordinated_stage_start_time[coord_key] = ('WAITING_FOR_OFFSET', matching_interstage.get('active_phases', {}))
                    # Keep coordinated signal in its current stage during offset wait
                    if stage_coord and stage_coord in stages[coordinated_id].get('stages', {}):
                        stage_data = stages[coordinated_id]['stages'][stage_coord]
                        stage_groups = stage_data.get('active_phases', [])
                        for sg in signal_groups_coord:
                            sg_name = sg.AttValue('Name')
                            if sg_name in stage_groups:
                                sg.SetAttValue('SigState', 'GREEN')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                    return
                
                tracking_data = coordinated_stage_start_time[coord_key]
                
                # Extract interstage info
                if isinstance(tracking_data, tuple):
                    interstage_state, interstage_signals = tracking_data
                else:
                    interstage_state = 'TRANSITIONING'
                    interstage_signals = matching_interstage.get('active_phases', {})
                
                # Check if offset has elapsed
                if interstage_state == 'WAITING_FOR_OFFSET':
                    if lead_id in stage_transition_time:
                        # Need to track offset time for interstage
                        # For now, assume offset has passed and start transition
                        coordinated_stage_start_time[coord_key] = (0, interstage_signals)
                        interstage_state = 'TRANSITIONING'
                    else:
                        return
                
                # Execute interstage transition timing
                if isinstance(tracking_data, tuple):
                    interstage_duration = tracking_data[0]
                    interstage_signals = tracking_data[1]
                else:
                    interstage_duration = 0
                    interstage_signals = matching_interstage.get('active_phases', {})
                
                interstage_duration += STEP_TIME
                coordinated_stage_start_time[coord_key] = (interstage_duration, interstage_signals)
                
                # Apply interstage signal timings
                for sg in signal_groups_coord:
                    sg_name = sg.AttValue('Name')
                    if sg_name in interstage_signals:
                        timing = interstage_signals[sg_name]
                        if isinstance(timing, tuple) and len(timing) == 2:
                            start_time, end_time = timing
                            # Check if signal should be AMBER or RED based on timing
                            if start_time <= interstage_duration <= end_time:
                                sg.SetAttValue('SigState', 'AMBER')
                            elif interstage_duration > end_time:
                                sg.SetAttValue('SigState', 'RED')
                            else:
                                sg.SetAttValue('SigState', 'RED')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                    else:
                        sg.SetAttValue('SigState', 'RED')
                
                # Check if interstage is complete
                if interstage_duration >= interstage_length:
                    # Interstage transition complete - transition to remapped target stage
                    target_stage_data = stages[coordinated_id].get('stages', {}).get(target_stage, {})
                    target_stage_groups = target_stage_data.get('active_phases', [])
                    for sg in signal_groups_coord:
                        sg_name = sg.AttValue('Name')
                        if sg_name in target_stage_groups:
                            sg.SetAttValue('SigState', 'GREEN')
                        else:
                            sg.SetAttValue('SigState', 'RED')
                    coordinated_stage_start_time[coord_key] = ("COMPLETED", 0)
            
            # Case 3: Lead signal has neither stage nor interstage (position unknown)
            # Stage was already selected by detectors in compute_coordination_decision
            else:
                # Check if there's an in-progress transition first
                existing_transition_key = None
                for key in coordinated_stage_start_time:
                    if key.startswith(f"{coordinated_id}_occupancy_"):
                        tracking_data = coordinated_stage_start_time[key]
                        if not (isinstance(tracking_data, tuple) and tracking_data[0] == "COMPLETED"):
                            # Found in-progress transition, extract target stage
                            existing_transition_key = key
                            target_stage = key.replace(f"{coordinated_id}_occupancy_", "")
                            break
                
                # If no in-progress transition, use occupancy-based intelligent stage selection
                if existing_transition_key is None:
                    if stage_occupancy and coordinated_id in stage_occupancy:
                        available_stages = stages[coordinated_id].get('stages', {}).keys()
                        target_stage = get_best_available_stage(coordinated_id, stage_lead, available_stages, stage_occupancy)
                    else:
                        target_stage = stage_lead if stage_lead else 'stage_1'
                
                # Check if coordinated signal needs to transition
                if target_stage is not None and stage_coord != target_stage:
                        # Start transition to target stage based on occupancy
                        coord_key = f"{coordinated_id}_occupancy_{target_stage}"
                        
                        current_stage_groups = []
                        if stage_coord and stage_coord in stages[coordinated_id].get('stages', {}):
                            current_stage_groups = stages[coordinated_id]['stages'][stage_coord].get('active_phases', [])
                        
                        if coord_key not in coordinated_stage_start_time:
                            # If coordinated signal is already in all-red, skip amber phase (nothing to clear)
                            initial_timer = 4 if stage_coord is None else 0
                            coordinated_stage_start_time[coord_key] = (initial_timer, current_stage_groups)
                            for sg in signal_groups_coord:
                                sg_name = sg.AttValue('Name')
                                if sg_name in current_stage_groups:
                                    sg.SetAttValue('SigState', 'AMBER')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                        else:
                            tracking_data = coordinated_stage_start_time[coord_key]
                            
                            if isinstance(tracking_data, tuple) and tracking_data[0] == "COMPLETED":
                                # Maintain GREEN on target stage
                                target_stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage, {}).get('active_phases', [])
                                for sg in signal_groups_coord:
                                    sg_name = sg.AttValue('Name')
                                    if sg_name in target_stage_groups:
                                        sg.SetAttValue('SigState', 'GREEN')
                                    else:
                                        sg.SetAttValue('SigState', 'RED')
                            else:
                                # Continue transition timing
                                if isinstance(tracking_data, tuple):
                                    amber_duration = tracking_data[0]
                                    stored_stage_groups = tracking_data[1]
                                else:
                                    amber_duration = tracking_data
                                    stored_stage_groups = current_stage_groups
                                
                                amber_duration += STEP_TIME
                                coordinated_stage_start_time[coord_key] = (amber_duration, stored_stage_groups)
                                
                                if amber_duration < 4:  # AMBER phase
                                    for sg in signal_groups_coord:
                                        sg_name = sg.AttValue('Name')
                                        if sg_name in stored_stage_groups:
                                            sg.SetAttValue('SigState', 'AMBER')
                                        else:
                                            sg.SetAttValue('SigState', 'RED')
                                elif amber_duration < 5:  # All-red phase
                                    for sg in signal_groups_coord:
                                        sg.SetAttValue('SigState', 'RED')
                                else:
                                    # Transition complete, switch to GREEN on target stage
                                    target_stage_groups = stages[coordinated_id].get('stages', {}).get(target_stage, {}).get('active_phases', [])
                                    for sg in signal_groups_coord:
                                        sg_name = sg.AttValue('Name')
                                        if sg_name in target_stage_groups:
                                            sg.SetAttValue('SigState', 'GREEN')
                                        else:
                                            sg.SetAttValue('SigState', 'RED')
                                    coordinated_stage_start_time[coord_key] = ("COMPLETED", 0)
        except Exception as e:
            pass  # user's opening features during simulation
    except Exception as e:
        print(f"Error applying coordination decision: {e}")


def get_signal_vaps(signal_controls):
    """
    Retrieve SupplyFile1 attribute from all signal controllers.
    
    Args:
        signal_controls: VISSIM signal controllers collection
    
    Returns:
        dict: Mapping of signal controller ID to SupplyFile1 path
              Example: {335: 'I:\\path\\sig_335.vap', 336: 'I:\\path\\sig_336_TSP.vap', ...}
    """
    supply_files = {}
    
    try:
        for sc in signal_controls:
            sc_id = sc.AttValue('No')
            supply_file = sc.AttValue('SupplyFile1')
            if supply_file:
                supply_files[sc_id] = supply_file
    except Exception as e:
        print(f"Error retrieving SupplyFile1: {e}")
    
    return supply_files


def get_stage_with_detectors(detectors, vap_files):
    """
    Parse VAP files and organize detector information by signal controller and port number.
    Extracts stage-to-detector mappings from VAP file EXPRESSIONS section for reference.
    
    Args:
        signal_controls: VISSIM signal controllers collection
        detectors: Dictionary of detectors organized by signal controller ID and port number
                   Structure: {sc_id: {port_number: [detector_names]}}
        vap_files: Dictionary mapping signal controller ID to SupplyFile1 path
                   Structure: {sc_id: supply_file_path}
    
    Returns:
        dict: Detectors organized by signal controller ID and port number
              Example: {
                  335: {2: ['det_335_2_1', 'det_335_2_2'], 6: ['det_335_6_1']},
                  336: {2: ['det_336_2_1'], 5: ['det_336_5_1']},
                  ...
              }
    """
    import re
    from pathlib import Path
    
    stage_detector_map = {}
    
    # Parse VAP files to extract stage-to-detector mappings
    for sc_id, supply_file in vap_files.items():
        if not supply_file or not Path(supply_file).exists():
            print(f"Warning: VAP file not found for signal controller {sc_id}: {supply_file}")
            continue
        
        # Initialize sc_id entry
        stage_detector_map[sc_id] = {}
        
        try:
            with open(supply_file, 'r') as f:
                content = f.read()
            
            # Extract the EXPRESSIONS section
            expressions_match = re.search(r'/\*\s*EXPRESSIONS\s*\*/(.*?)(?=/\*|$)', content, re.DOTALL)
            if not expressions_match:
                continue
            
            expressions_section = expressions_match.group(1)
            
            # Dictionary to map stage letters to stage_x names
            stage_mapping = {}
            
            # Extract stage mappings (A := stage_1, B := stage_2, etc.)
            stage_pattern = r'(\w+)\s*:=\s*(stage_\d+)\s*;'
            for match in re.finditer(stage_pattern, expressions_section):
                letter = match.group(1)
                stage_name = match.group(2)
                stage_mapping[letter] = stage_name
            
            # Extract detector port numbers from NoDetect_X expressions (ignore Stop_X)
            # Pattern: NoDetect_X := (Detection/Occupancy( port_number ) = ...)
            nodetect_pattern = r'NoDetect_(\w+)\s*:=\s*(.+?);'
            
            for match in re.finditer(nodetect_pattern, expressions_section):
                letter = match.group(1)
                expression = match.group(2)
                
                # Extract port numbers from Detection() and Occupancy() calls
                port_pattern = r'(?:Detection|Occupancy)\s*\(\s*(\w+)\s*\)'
                ports = []
                
                for port_match in re.finditer(port_pattern, expression):
                    port = port_match.group(1)
                    # Try to convert to int, otherwise keep as string (for SG102, etc.)
                    try:
                        ports.append(int(port))
                    except ValueError:
                        ports.append(port)
                
                # Map stage_x to ports
                if letter in stage_mapping:
                    stage_name = stage_mapping[letter]
                    if stage_name not in stage_detector_map[sc_id]:
                        stage_detector_map[sc_id][stage_name] = []
                    # Add ports, port number with detector names, avoiding duplicates
                    for port in ports:
                        if port not in stage_detector_map[sc_id][stage_name]:
                            if isinstance(port, str):
                                port_no = int(port.replace('SG', ''))
                            else:
                                port_no = port
                            
                            # Check if this port has detectors before accessing
                            if sc_id in detectors and port_no in detectors[sc_id]:
                                detector_names = detectors[sc_id][port_no]
                                stage_detector_map[sc_id][stage_name].append(detector_names)
                            else:
                                # Port doesn't have detectors, skip it
                                pass
        
        except Exception as e:
            print(f"Error parsing VAP file for signal controller {sc_id}: {e}")
    
    return stage_detector_map


def main():
    vissim = load_project(working_path=WORKING_PATH, project_name=PROJECT_NAME)
    sim = vissim.Simulation
    vnet = vissim.Net # create Net COM-interface
    signal_controls = vnet.SignalControllers
    # get all stages
    signal_control_ids = [sc.AttValue('No') for sc in signal_controls]
    stages = {}
    for sc_id in signal_control_ids:
        stages[sc_id] = get_stage_signal_groups_from_pua(sc_id)
    # Get all detectors by stage
    detectors = get_all_detectors(vnet)
    # Retrieve SupplyFile1 from all signal controllers
    vap_files = get_signal_vaps(signal_controls)
    # Build a map for detector port numbers and stages using VAP files
    detector_stage_map = get_stage_with_detectors(detectors, vap_files)

    for i, seed in enumerate(RANDOM_SEEDS):
        print(f'Running {i+1}th random seed {seed}...')
        # Set the random seed for Vissim (not Python's random)
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
            
            # During warm-up period, track stage duration for each coordinated signal
            # so that when coordination starts the elapsed time is already accurate
            if sim_step < 120:
                for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                    sc_lead = signal_controls.ItemByKey(lead)
                    lead_state = get_signal_state_for_controller(sc_lead)
                    stage_info_lead = get_active_stage(lead, stages, lead_state['active_groups'], lead_state['yellow_groups'])
                    if stage_info_lead['stage'] is None and stage_info_lead['interstage'] is None:
                        for coordinated in coordinated_dict.keys():
                            sc_coord = signal_controls.ItemByKey(coordinated)
                            coord_state = get_signal_state_for_controller(sc_coord)
                            coord_stage_info = get_active_stage(coordinated, stages, coord_state['active_groups'], coord_state['yellow_groups'])
                            stage_duration_key = f"{coordinated}_occupancy_stage_duration"
                            current_stage = coord_stage_info['stage']
                            if stage_duration_key not in coordinated_stage_start_time:
                                coordinated_stage_start_time[stage_duration_key] = (0, current_stage)
                            else:
                                elapsed_time, tracked_stage = coordinated_stage_start_time[stage_duration_key]
                                if current_stage == tracked_stage:
                                    coordinated_stage_start_time[stage_duration_key] = (elapsed_time + STEP_TIME, tracked_stage)
                                else:
                                    # Stage changed - reset tracker
                                    coordinated_stage_start_time[stage_duration_key] = (0, current_stage)
                continue  # Starting coordination after 2 minutes
            
            # Read all signal states on main thread (VISSIM COM access must be here).
            # stage_info is computed once here so downstream code never repeats the call.
            signal_states = {}
            for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                sc_lead = signal_controls.ItemByKey(lead)
                state = get_signal_state_for_controller(sc_lead)
                state['stage_info'] = get_active_stage(lead, stages, state['active_groups'], state['yellow_groups'])
                signal_states[lead] = state
                
                for coordinated, offset in coordinated_dict.items():
                    if coordinated not in signal_states:
                        sc_coord = signal_controls.ItemByKey(coordinated)
                        state = get_signal_state_for_controller(sc_coord)
                        state['stage_info'] = get_active_stage(coordinated, stages, state['active_groups'], state['yellow_groups'])
                        signal_states[coordinated] = state
            
            # Determine which coordinated signals need occupancy calculated
            # (when their lead signal has no valid stage/interstage)
            coordinated_ids_needing_occupancy = set()
            min_stage_duration = 6.0  # seconds - must match the duration in apply_coordination_decision
            
            for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                stage_info_lead = signal_states[lead]['stage_info']
                
                # If lead signal is undetectable (no stage, no interstage), check if occupancy calculation is needed
                if stage_info_lead['stage'] is None and stage_info_lead['interstage'] is None:
                    for coordinated in coordinated_dict.keys():
                        # Always calculate occupancy if signal is in all-red (stage_coord is None)
                        # so we can select best stage to escape to
                        coord_stage_info = signal_states[coordinated]['stage_info']
                        if coord_stage_info['stage'] is None and coord_stage_info['interstage'] is None:
                            # Signal is in all-red - must calculate occupancy to escape
                            coordinated_ids_needing_occupancy.add(coordinated)
                            continue
                        
                        # Skip occupancy calculation if signal is still in minimum stage duration window and it's not in all-red
                        stage_duration_key = f"{coordinated}_occupancy_stage_duration"
                        if stage_duration_key in coordinated_stage_start_time:
                            tracking_data = coordinated_stage_start_time[stage_duration_key]
                            if isinstance(tracking_data, tuple):
                                elapsed_time, _ = tracking_data
                                # Still in minimum duration window - skip occupancy calculation
                                if elapsed_time < min_stage_duration:
                                    continue
                        
                        # Outside minimum window or no tracking - calculate occupancy
                        coordinated_ids_needing_occupancy.add(coordinated)
            
            # Calculate stage occupancy only for signal controllers that need it
            stage_occupancy = {}
            if coordinated_ids_needing_occupancy:
                occupancy_for_needed = get_all_stage_occupancy(list(coordinated_ids_needing_occupancy), detector_stage_map, vnet)
                stage_occupancy.update(occupancy_for_needed)
            
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
                            lead, coordinated, lead_data, coord_data,
                            previous_stage, offset, stage_occupancy
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
                        stage_transition_time, previous_stage, coordinated_stage_start_time, previous_stage_lead, stage_occupancy, vnet
                    )
        
        vissim.SaveNet()
        # Stop simulation before setting new seed
        sim.Stop()
        time.sleep(2)  # Wait for simulation to fully stop
    
    vissim.Exit()
    vissim = None


if __name__ == '__main__':
    main()
