"""
Alternative Threading Implementation using Work Queue Pattern

This approach uses a producer-consumer pattern with a shared work queue.
May be more efficient than ThreadPoolExecutor for continuous simulation steps.

Use this if run_vissim_threaded.py doesn't provide expected speedup.
"""

import os
import time
from queue import Queue, Empty
from threading import Thread, Event
from typing import Dict

from vissim.utils import load_project
from config_vissim import *


class CoordinationWorker(Thread):
    """Worker thread that computes coordination decisions from pre-read signal data."""
    
    def __init__(self, worker_id: int, task_queue: Queue, result_queue: Queue, stop_event: Event):
        """
        Args:
            worker_id: Unique ID for this worker
            task_queue: Queue containing pure data coordination tasks (no VISSIM objects)
            result_queue: Queue to put computed decisions
            stop_event: Event to signal worker to stop
        """
        super().__init__(daemon=True)
        self.worker_id = worker_id
        self.task_queue = task_queue
        self.result_queue = result_queue
        self.stop_event = stop_event
    
    def run(self):
        """Worker main loop - process tasks until stop_event is set."""
        while not self.stop_event.is_set():
            try:
                # Get task with timeout to check stop_event
                task = self.task_queue.get(timeout=0.1)
                if task is None:  # Sentinel value to stop
                    break
                
                # Process coordination task (pure data, no VISSIM COM access)
                lead_id, coordinated_id, lead_data, coord_data, stages, previous_stage, lead_signal_id, offset, previous_stage_lead = task
                
                decision = self._compute_coordination(
                    lead_id, coordinated_id, lead_data, coord_data, stages,
                    previous_stage, lead_signal_id, offset, previous_stage_lead
                )
                
                # Put result back in queue
                self.result_queue.put((lead_id, coordinated_id, decision))
                self.task_queue.task_done()
                
            except Empty:
                # Queue timeout - normal operation, just continue waiting
                continue
            except Exception as e:
                # Only print actual errors, not queue timeouts
                print(f"Worker {self.worker_id} error: {e}")
    
    def _compute_coordination(self, lead_id: int, coordinated_id: int, 
                             lead_data: Dict, coord_data: Dict,
                             stages: Dict, previous_stage: Dict,
                             lead_signal_id: int, offset: int, previous_stage_lead: Dict) -> Dict:
        """
        Compute coordination decision from pre-read signal data.
        No VISSIM COM access - only pure Python logic.
        
        Args:
            lead_data: Dict with 'active_groups' and 'yellow_groups' for lead signal
            coord_data: Dict with 'active_groups' and 'yellow_groups' for coordinated signal
            lead_signal_id: The ID of the lead signal
            offset: Offset time in seconds
            previous_stage_lead: Dictionary tracking previous valid stage_lead for each signal
        """
        try:
            active_groups_lead = lead_data['active_groups']
            yellow_groups_lead = lead_data['yellow_groups']
            
            active_groups_coord = coord_data['active_groups']
            
            stage_lead = self._get_active_stage(lead_id, stages, active_groups_lead)
            # Use previous stage_lead if current stage_lead is None
            original_stage_lead = stage_lead
            if stage_lead is None:
                stage_lead = previous_stage_lead.get(lead_id)
            whether_lead_transition = self._get_transition_type(yellow_groups_lead, active_groups_lead)
            stage_coord = self._get_active_stage(coordinated_id, stages, active_groups_coord)
            
            return {
                'lead_id': lead_id,
                'coordinated_id': coordinated_id,
                'stage_lead': stage_lead,
                'original_stage_lead': original_stage_lead,
                'stage_coord': stage_coord,
                'whether_lead_transition': whether_lead_transition,
                'active_groups_coord': active_groups_coord,
                'stage_changed_lead': (stage_lead != previous_stage.get(lead_id)),
                'stage_changed_coord': (stage_coord != previous_stage.get(coordinated_id)),
                'lead_signal_id': lead_signal_id,
                'offset': offset,
            }
        except Exception as e:
            print(f"Error in coordination for {lead_id}->{coordinated_id}: {e}")
            return None
    
    def _get_active_stage(self, sc_id, stages, green_groups):
        """Get current stage based on active signal groups."""
        if sc_id not in stages:
            return None
        
        for stage, groups in stages[sc_id].items():
            if set(groups) == set(green_groups):
                return stage
        return None
    
    def _get_transition_type(self, yellow_groups, active_groups):
        """Determine stage transition type."""
        yellow_signal_names = ['NBT', 'NBTR', 'NBTRL',
                               'SBT', 'SBTR', 'SBTRL',
                               'EBT', 'EBTR', 'EBTRL',
                               'WBT', 'WBTR', 'WBTRL']
        
        has_yellow_crossing = any(g in CROSSING_NAMES for g in yellow_groups)
        has_yellow_signals = any(g in yellow_signal_names for g in yellow_groups)
        
        if has_yellow_crossing and has_yellow_signals:
            return 'Crossing and Signal'
        elif has_yellow_crossing:
            return 'Crossing'
        elif not has_yellow_signals and len(active_groups) > 0:
            return 'Not Transition'
        else:
            return 'All Red'


def get_stage_signal_groups_from_controller(sc_id):
    """Read stage definitions from PUA file."""
    pua_file = os.path.join(PUA_FILE_PATH, f'sig_{sc_id}.pua')
    
    if not os.path.exists(pua_file):
        return {}
    
    stages = {}
    try:
        with open(pua_file, 'r') as f:
            for line in f:
                line = line.strip()
                if line.startswith('stage_'):
                    stage_phases = line.split(' ')
                    stage_name = stage_phases[0]
                    if stage_name in stages:
                        continue
                    signal_groups = [item for item in stage_phases[2:] if item]
                    stages[stage_name] = signal_groups
    except:
        pass
    return stages


def apply_coordination_decision(signal_controls, stages, lead_id, coordinated_id,
                               decision, stage_transition_time, previous_stage):
    """Apply coordination decision on main thread (VISSIM COM operations)."""
    if decision is None:
        return
    
    try:
        stage_lead = decision['stage_lead']
        whether_lead_transition = decision['whether_lead_transition']
        
        # Update tracking
        if decision['stage_changed_lead']:
            previous_stage[lead_id] = stage_lead
            if stage_lead:
                stage_transition_time[lead_id][stage_lead] = 0
        else:
            if stage_lead and stage_lead in stage_transition_time[lead_id]:
                stage_transition_time[lead_id][stage_lead] += STEP_TIME
        
        if decision['stage_changed_coord']:
            previous_stage[coordinated_id] = decision['stage_coord']
            if decision['stage_coord']:
                stage_transition_time[coordinated_id][decision['stage_coord']] = 0
        else:
            if decision['stage_coord'] and decision['stage_coord'] in stage_transition_time[coordinated_id]:
                stage_transition_time[coordinated_id][decision['stage_coord']] += STEP_TIME
        
        # Skip in invalid states - when lead signal is undetectable, stop forcing coordination
        # This allows VISSIM to resume normal signal control
        if stage_lead is None:
            return
        
        # Apply signal changes
        sc_coord = signal_controls.ItemByKey(coordinated_id)
        stage_groups = stages[coordinated_id].get(stage_lead, [])
        
        if whether_lead_transition == 'All Red':
            for sg in sc_coord.SGs:
                sg.SetAttValue('SigState', 'RED')
        elif whether_lead_transition == 'Crossing and Signal':
            for sg in sc_coord.SGs:
                sg_name = sg.AttValue('Name')
                sg.SetAttValue('SigState', 'AMBER' if sg_name in stage_groups else 'RED')
        elif whether_lead_transition == 'Crossing':
            for sg in sc_coord.SGs:
                sg_name = sg.AttValue('Name')
                if sg_name in stage_groups and sg_name in CROSSING_NAMES:
                    sg.SetAttValue('SigState', 'AMBER')
                elif sg_name in stage_groups:
                    sg.SetAttValue('SigState', 'GREEN')
                else:
                    sg.SetAttValue('SigState', 'RED')
        elif whether_lead_transition == 'Not Transition':
            for sg in sc_coord.SGs:
                sg_name = sg.AttValue('Name')
                sg.SetAttValue('SigState', 'GREEN' if sg_name in stage_groups else 'RED')
    except:
        pass


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


def main():
    """Main simulation loop with queue-based worker threads."""
    vissim = load_project(working_path=WORKING_PATH, project_name=PROJECT_NAME)
    sim = vissim.Simulation
    vnet = vissim.Net
    signal_controls = vnet.SignalControllers
    
    # Get all stages
    signal_control_ids = [sc.AttValue('No') for sc in signal_controls]
    stages = {}
    for sc_id in signal_control_ids:
        stages[sc_id] = get_stage_signal_groups_from_controller(sc_id)
    
    # Setup coordination queue
    num_workers = min(3, len(COORD_SIGNAL_OFFSET))
    task_queue = Queue(maxsize=100)
    result_queue = Queue()
    stop_event = Event()
    
    # Start worker threads
    workers = []
    for i in range(num_workers):
        worker = CoordinationWorker(i, task_queue, result_queue, stop_event)
        worker.start()
        workers.append(worker)
    
    print(f"Started {num_workers} coordination worker threads")
    
    for i, seed in enumerate(RANDOM_SEEDS):
        print(f'Running {i+1}th random seed {seed}...')
        sim.SetAttValue('RandSeed', seed)
        
        # Initialize tracking
        stage_transition_time = {sc_id: {} for sc_id in signal_control_ids}
        previous_stage = {sc_id: None for sc_id in signal_control_ids}
        # Track previous valid stage_lead for each signal controller
        previous_stage_lead = {sc_id: None for sc_id in signal_control_ids}
        
        for sim_step in range(EVAL_FROM_TIME, EVAL_FROM_TIME+PERIOD_TIME*STEP_TIME+1):
            sim.RunSingleStep()
            
            # Read signal states on main thread (VISSIM COM access must be here)
            signal_states = {}
            for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                sc_lead = signal_controls.ItemByKey(lead)
                signal_states[lead] = get_signal_state_for_controller(sc_lead)
                
                for coordinated, offset in coordinated_dict.items():
                    sc_coord = signal_controls.ItemByKey(coordinated)
                    signal_states[coordinated] = get_signal_state_for_controller(sc_coord)
            
            # Submit all coordination tasks to worker queue (pure data only)
            for lead, coordinated_dict in COORD_SIGNAL_OFFSET.items():
                for coordinated, offset in coordinated_dict.items():
                    lead_data = signal_states[lead]
                    coord_data = signal_states[coordinated]
                    
                    task = (lead, coordinated, lead_data, coord_data, stages, previous_stage, lead, offset, previous_stage_lead)
                    task_queue.put(task)
            
            # Collect all results for this step and apply them on main thread
            num_tasks = sum(len(v) for v in COORD_SIGNAL_OFFSET.values())
            for _ in range(num_tasks):
                try:
                    lead_id, coordinated_id, decision = result_queue.get(timeout=5)
                    if decision:
                        # Skip coordination for invalid lead signal states
                        stage_lead = decision.get('stage_lead')
                        original_stage_lead = decision.get('original_stage_lead')
                        whether_lead_transition = decision.get('whether_lead_transition')
                        stage_coord = decision.get('stage_coord')
                        lead_signal_id = decision.get('lead_signal_id')
                        offset = decision.get('offset')
                        
                        # Update previous_stage_lead only if original_stage_lead was valid (not None)
                        if original_stage_lead is not None:
                            previous_stage_lead[lead_id] = original_stage_lead
                        
                        # When lead signal is undetectable, stop forcing coordination
                        # This allows VISSIM to resume normal signal control
                        if stage_lead is None:
                            continue
                        
                        # Check if offset time has elapsed since lead signal changed to this stage
                        stage_groups = stages[coordinated_id].get(stage_lead, [])
                        
                        if lead_signal_id in stage_transition_time and stage_lead and stage_lead in stage_transition_time[lead_signal_id]:
                            time_since_lead_transition = stage_transition_time[lead_signal_id][stage_lead]
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
                                    continue
                                # If not in target stage, fall through to apply coordination
                        
                        # Extract data from decision and apply
                        sc_coord = signal_controls.ItemByKey(coordinated_id)
                        
                        if whether_lead_transition == 'All Red':
                            for sg in sc_coord.SGs:
                                sg.SetAttValue('SigState', 'RED')
                        elif whether_lead_transition == 'Crossing and Signal':
                            for sg in sc_coord.SGs:
                                sg_name = sg.AttValue('Name')
                                sg.SetAttValue('SigState', 'AMBER' if sg_name in stage_groups else 'RED')
                        elif whether_lead_transition == 'Crossing':
                            for sg in sc_coord.SGs:
                                sg_name = sg.AttValue('Name')
                                if sg_name in stage_groups and sg_name in CROSSING_NAMES:
                                    sg.SetAttValue('SigState', 'AMBER')
                                elif sg_name in stage_groups:
                                    sg.SetAttValue('SigState', 'GREEN')
                                else:
                                    sg.SetAttValue('SigState', 'RED')
                        elif whether_lead_transition == 'Not Transition':
                            for sg in sc_coord.SGs:
                                sg_name = sg.AttValue('Name')
                                sg.SetAttValue('SigState', 'GREEN' if sg_name in stage_groups else 'RED')
                except Exception as ex:
                    if 'Empty' not in str(ex):
                        print(f"Error applying coordination at step {sim_step}: {ex}")
        
        vissim.SaveNet()
        sim.Stop()
        time.sleep(2)
    
    # Stop workers
    stop_event.set()
    for _ in range(num_workers):
        task_queue.put(None)  # Sentinel
    
    for worker in workers:
        worker.join(timeout=5)
    
    vissim.Exit()
    vissim = None


if __name__ == '__main__':
    main()
