#pylint:disable = all
import threading as th 
import os
import yaml
# Define global variables and events
macro_done : th.Event = th.Event()
get_user_and_path : th.Event = th.Event()
is_run_macro : th.Event = th.Event()
done : th.Event = th.Event()
is_event : th.Event = th.Event()
progress : int = 0
yaml_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.yml')
count_get_sap : int = 5