#pylint:disable = all
import os
# Fix TERM error cho Windows
if os.name == "nt":
    os.environ.setdefault("TERM", "xterm")
from blessed import Terminal
import threading as th 
import platform


# Define global variables and events
macro_done : th.Event = th.Event()
get_user_and_path : th.Event = th.Event()
is_run_macro : th.Event = th.Event()
done : th.Event = th.Event()
is_event : th.Event = th.Event()
progress : int = 0
term = Terminal()
