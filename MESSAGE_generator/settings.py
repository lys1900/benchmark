#YL, IAEA, 2026 September
"""
Where the folders this project works in are.

Every path lives here rather than at the top of each script, so that the
workbench can change one and the other scripts pick it up. The values below
are the defaults; whatever the workbench saves is kept in settings.json next
to these scripts and wins over them. Delete that file to go back to the
defaults.

Running any script by hand works the same either way: it reads settings.json
if there is one, otherwise it uses what is written here.
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
saved_fp = os.path.join(HERE, 'settings.json')

DEFAULTS = {
    #the reference case copied to seed every case that is built
    'orig_base_fd': 'E:/Work/benchmark/MESSAGE_generator/MESSAGE_orig/CountryA',
    #one row per scenario: main, input_fn, input_fd, make, add_storage
    'create_cases_fp': 'E:/Work/benchmark/MESSAGE_generator/create_cases.csv',
    #where MESSAGE itself is installed, the models tree hangs off this
    'MESSAGE_fd': 'C:/Programs/MESSAGE_INT/',
    #cases are written here first, then copied into the MESSAGE tree
    'output_base_fd': 'E:/Work/benchmark/MESSAGE_generator/MESSAGE_out/',
    #one row per processed run, read by process_results.py
    'process_results_fp': 'E:/Work/benchmark/MESSAGE_generator/process_results.csv',
    #a folder of csv files per run
    'results_fd': 'E:/Work/benchmark/results/',
    #the workbooks a new scenario can be built from
    'input_fd': 'E:/Work/benchmark/input_China/',
}

#what the workbench is allowed to change, in the order it shows them
EDITABLE = ['orig_base_fd', 'create_cases_fp', 'MESSAGE_fd', 'output_base_fd',
            'process_results_fp', 'results_fd', 'input_fd']

LABELS = {
    'orig_base_fd': 'Reference case folder',
    'create_cases_fp': 'create_cases.csv',
    'MESSAGE_fd': 'MESSAGE installation',
    'output_base_fd': 'Staging folder',
    'process_results_fp': 'process_results.csv',
    'results_fd': 'Results folder',
    'input_fd': 'Workbook folder',
}


def load():
    """the defaults, with anything the workbench saved written over them"""
    values = dict(DEFAULTS)
    if os.path.exists(saved_fp):
        try:
            with open(saved_fp, encoding='utf-8') as f:
                saved = json.load(f)
            values.update({k: str(v) for k, v in saved.items()
                           if k in DEFAULTS and str(v).strip()})
        except (OSError, ValueError) as err:
            #a broken settings file should not stop a build
            print(f"warning: {saved_fp} could not be read ({err}), using the defaults")

    return values


def save(values):
    """keep the paths for the next session, ignoring anything unrecognised"""
    keep = {k: str(v).strip() for k, v in values.items()
            if k in DEFAULTS and str(v).strip()}
    with open(saved_fp, 'w', encoding='utf-8') as f:
        json.dump(keep, f, indent=2)

    return load()
