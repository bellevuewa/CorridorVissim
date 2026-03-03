import os
import logging
import pandas as pd

from logging import config as logger_config
from vissim.config_vissim import *

logger_config.fileConfig(os.path.join(os.getcwd(), 'vissim', 'logger.conf'))
# create logger
logger = logging.getLogger('importVehicle')


def start_vissim():
    # COM-Server
    import win32com.client as com
    ## Connecting the COM Server => Open a new Vissim Window:
    Vissim = com.Dispatch("Vissim.Vissim")
    return Vissim


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
    signal_control = signal_controls.ItemByKey(1)
    
    for sim_step in range(EVAL_FROM_TIME, EVAL_FROM_TIME+PERIOD_TIME*STEP_TIME+1):
        sim.RunSingleStep()
        if sim_step == 0: continue
        if (sim_step / STEP_TIME) % EVAL_INTERVAL*STEP_TIME == 0:
            if signal_control.AttValue('SigState') == 'RED':
                signal_control.SetAttValue('SigState', 'GREEN')
                for i in [2, 3, 4, 5, 9]:
                    sig = signal_controls.ItemByKey(i)
                    sig.SetAttValue('SigState', 'RED')


if __name__ == '__main__':
    main()
