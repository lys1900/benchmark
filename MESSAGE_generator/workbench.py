#YL, IAEA, 2026 September
"""
Drive the workflow from one page, locally.

    python workbench.py

opens http://localhost:8764. From there you can point the scripts at the
folders you keep things in, add a scenario, and launch the three steps:

    build    python MESSAGE_trans.py     for the ticked scenarios
    solve    run_<solver>_<case>_adb.bat in each case folder
    process  python process_results.py   for the ticked scenarios

Each launch opens its own command window, so you watch it there and this
server never holds a long solve open. Nothing leaves the machine.

Scenarios are create_cases.csv: the page shows what is in it and writes back
to it as soon as you add or remove one. The paths are settings.json, which
every script reads, so setting one here sets it everywhere. Every launch is
appended to actions.csv, which is what the history table shows.
"""
import csv
import datetime
import glob
import http.server
import json
import os
import re
import socketserver
import subprocess
import sys
import threading
import webbrowser

import settings

PORT = 8764
HERE = os.path.dirname(os.path.abspath(__file__))
page_fp = os.path.join(HERE, 'workbench.html')
actions_fp = os.path.join(HERE, 'actions.csv')
runner_fd = os.path.join(HERE, '.workbench')  #the little batch files each window runs
SOLVERS = ['HiGHS', 'cplex']
VIEWER_PORT = 8765          #results_viz.py


# --------------------------------------------------------------------------
# what is on disk
# --------------------------------------------------------------------------

def interpreter_ok():
    """
    whether the python that would run the steps can import what they need.

    the workbench itself needs nothing but the standard library, so it starts
    happily under a python that MESSAGE_trans would fail under. checked once,
    since it costs a process
    """
    try:
        r = subprocess.run([sys.executable, '-c', 'import pandas, openpyxl, numpy'],
                           capture_output=True, text=True, timeout=60)
        if r.returncode == 0:
            return True, ''
        last = (r.stderr or '').strip().splitlines()
        return False, last[-1] if last else 'could not import pandas'
    except Exception as err:
        return False, f"{type(err).__name__}: {err}"


PYTHON_OK, PYTHON_WHY = interpreter_ok()


def stamp(path):
    """when a file was last written, or None if it is not there"""
    if not os.path.exists(path):
        return None

    return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d %b %Y %H:%M')


def newer_than(path, than):
    """
    whether path was written after than.

    a case is seeded by copying the reference case, which carries a res folder
    of its own, so the mere presence of a result file does not mean this model
    was solved. the same goes for results processed before the last solve
    """
    if not os.path.exists(path):
        return None
    if not than or not os.path.exists(than):
        return True

    return os.path.getmtime(path) > os.path.getmtime(than)


def case_fd(paths, case):
    """the case folder in the MESSAGE tree"""
    return f"{paths['MESSAGE_fd']}models/{case}/{case}"


def read_scenarios(paths):
    """create_cases.csv as a list, or empty if it is not there yet"""
    fp = paths['create_cases_fp']
    if not os.path.exists(fp):
        return []
    rows = []
    with open(fp, newline='') as f:
        for r in csv.DictReader(f):
            if not r.get('main'):
                continue
            rows.append({'main': str(r['main']).strip(),
                         'input_fn': str(r['input_fn']).strip(),
                         #a folder joined straight onto a workbook name
                         'input_fd': settings.as_folder(r['input_fd']),
                         'make': int(float(r.get('make') or 0)),
                         'add_storage': int(float(r.get('add_storage') or 0)),
                         #blank means work it out from the Years sheet
                         'ntrun': str(r.get('ntrun') or '').strip()})

    return rows


def write_scenarios(paths, rows):
    """
    put the list back. the make column is what MESSAGE_trans reads to decide
    what to build, so the build step sets it from whatever was ticked
    """
    fp = paths['create_cases_fp']
    os.makedirs(os.path.dirname(fp) or '.', exist_ok=True)
    with open(fp, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['main', 'input_fn', 'input_fd', 'make', 'add_storage', 'ntrun'])
        for r in rows:
            w.writerow([r['main'], r['input_fn'], settings.as_folder(r['input_fd']),
                        int(r.get('make', 0)), int(r.get('add_storage', 0)),
                        str(r.get('ntrun') or '').strip()])


def write_process_control(paths, cases):
    """process_results.csv, so the results step knows where each run lives"""
    fp = paths['process_results_fp']
    os.makedirs(os.path.dirname(fp) or '.', exist_ok=True)
    with open(fp, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['run', 'path', 'input_fd', 'input_fn'])
        for c in cases:
            w.writerow([c['case'], case_fd(paths, c['case']), c['input_fd'], c['input_fn']])


def workbooks(input_fd):
    """
    the workbooks a scenario could be built from. every scenario carries its
    own folder in create_cases.csv, so two of them can be built from two
    different sets of province data
    """
    if not input_fd or not os.path.isdir(input_fd):
        return []
    names = [os.path.basename(p) for p in glob.glob(os.path.join(input_fd, '*.xlsx'))]

    return sorted(n for n in names if not n.startswith('~$'))


def viewer_running():
    """whether results_viz.py is already listening, so we do not start a second"""
    import socket
    with socket.socket() as s:
        s.settimeout(0.25)
        return s.connect_ex(('127.0.0.1', VIEWER_PORT)) == 0


def record(step, scenario, detail):
    """append what the user just launched, so the history outlives the page"""
    new = not os.path.exists(actions_fp)
    with open(actions_fp, 'a', newline='') as f:
        w = csv.writer(f)
        if new:
            w.writerow(['when', 'step', 'scenario', 'detail'])
        w.writerow([datetime.datetime.now().strftime('%d %b %Y %H:%M:%S'),
                    step, scenario, detail])


def history(limit=40):
    """the most recent actions, newest first"""
    if not os.path.exists(actions_fp):
        return []
    with open(actions_fp, newline='') as f:
        rows = list(csv.DictReader(f))

    return rows[-limit:][::-1]


def state():
    """everything the page draws, read fresh off disk"""
    paths = settings.load()
    scenarios = []
    for s in read_scenarios(paths):
        case = f"{s['main']}_{s['input_fn']}"
        fd = case_fd(paths, case)
        scenarios.append({
            **s,
            'case': case,
            'workbook': f"{s['input_fn']}.xlsx",
            'built': stamp(f"{fd}/data/{case}.adb"),
            'solved': stamp(f"{fd}/res/{case}_adb.txt"),
            'processed': stamp(f"{paths['results_fd']}{case}/prod_all.csv"),
            #None not there, False there but older than the step before it
            'solvedFresh': newer_than(f"{fd}/res/{case}_adb.txt", f"{fd}/data/{case}.adb"),
            'processedFresh': newer_than(f"{paths['results_fd']}{case}/prod_all.csv",
                                         f"{fd}/res/{case}_adb.txt"),
            'solvers_ready': [v for v in SOLVERS
                              if os.path.exists(f"{fd}/run_{v}_{case}_adb.bat")],
        })

    return {'paths': paths, 'editable': settings.EDITABLE, 'labels': settings.LABELS,
            'defaults': settings.DEFAULTS, 'solvers': SOLVERS,
            'workbooks': workbooks(paths['input_fd']),
            'scenarios': scenarios, 'history': history(),
            'viewerUrl': f"http://localhost:{VIEWER_PORT}",
            'python': sys.executable, 'pythonOk': PYTHON_OK, 'pythonWhy': PYTHON_WHY,
            'refreshed': datetime.datetime.now().strftime('%H:%M:%S')}


# --------------------------------------------------------------------------
# launching the three steps, each in its own window
# --------------------------------------------------------------------------

def open_window(title, cmdline, cwd, hold=True):
    """
    run a command in a command window of its own.

    the command is written to a small batch file first and that is what gets
    started. passing it to "cmd /k" directly does not survive the quoting: a
    python path with a space in it comes back as \\"C:\\Program Files\\... and
    cmd reports it is not a recognised command

    :param hold: keep the window open when the command finishes, so a failure
                 can be read. off for the solver, whose own batch file already
                 ends on a pause
    """
    if os.name != 'nt':
        subprocess.Popen(cmdline, cwd=cwd, shell=True,
                         stdin=subprocess.DEVNULL, close_fds=True)
        return

    os.makedirs(runner_fd, exist_ok=True)
    safe = re.sub(r'[^A-Za-z0-9]+', '_', title).strip('_').lower() or 'run'
    wrapper = os.path.join(runner_fd, f"{safe}.bat")
    with open(wrapper, 'w', newline='\r\n') as f:
        f.write('@echo off\n')
        f.write(f'title {title}\n')
        f.write(f'cd /d "{cwd}"\n')
        f.write(f'{cmdline}\n')
        if hold:
            f.write('echo.\n')
            f.write('echo === finished, exit code %errorlevel% ===\n')
            f.write('pause\n')
    #start opens the new window; the empty pair is the window title it expects
    subprocess.Popen(f'start "" "{wrapper}"', cwd=cwd, shell=True, close_fds=True)

    return wrapper


def launch(step, cases, solver):
    """turn a request from the page into command windows"""
    paths = settings.load()
    names = ', '.join(c['case'] for c in cases) or '-'
    #the viewer is the one step that acts on the results folder as a whole
    if not cases and step != 'view':
        raise ValueError("nothing selected")

    if step == 'build':
        #MESSAGE_trans builds whatever has make set, so the ticks are written
        #into that column first
        wanted = {c['case'] for c in cases}
        rows = read_scenarios(paths)
        for r in rows:
            r['make'] = 1 if f"{r['main']}_{r['input_fn']}" in wanted else 0
        write_scenarios(paths, rows)
        open_window("MESSAGE build", f'"{sys.executable}" MESSAGE_trans.py', HERE)
        record('build', names, 'MESSAGE_trans.py')
        return f"building {names} in a new window"

    if step == 'solve':
        missing = [c['case'] for c in cases
                   if not os.path.exists(
                       f"{case_fd(paths, c['case'])}/run_{solver}_{c['case']}_adb.bat")]
        if missing:
            raise FileNotFoundError(
                f"no {solver} batch file for {', '.join(missing)} — build it first")
        for c in cases:
            fd = case_fd(paths, c['case'])
            open_window(f"solve {c['case']}", f"run_{solver}_{c['case']}_adb.bat", fd, hold=False)
        record('solve', names, solver)
        return f"solving {names} with {solver}, one window each"

    if step == 'process':
        write_process_control(paths, cases)
        open_window("MESSAGE results", f'"{sys.executable}" process_results.py', HERE)
        record('process', names, 'process_results.py')
        return f"processing {names} in a new window"

    if step == 'view':
        if viewer_running():
            record('view', '-', f"opened the viewer already on port {VIEWER_PORT}")
            return 'viewer already running'
        #--no-browser because the page opens the tab, and the two of them
        #together were opening one each
        open_window("MESSAGE viewer",
                    f'"{sys.executable}" results_viz.py --no-browser', HERE)
        record('view', '-', f"started results_viz.py on port {VIEWER_PORT}")
        return 'started'

    raise ValueError(f"unknown step {step!r}")


# --------------------------------------------------------------------------
# the server
# --------------------------------------------------------------------------

class Handler(http.server.BaseHTTPRequestHandler):

    def reply(self, obj, code=200):
        body = json.dumps(obj).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split('?')[0]
        if path == '/api/state':
            self.reply(state())
        elif path == '/api/workbooks':
            #which workbooks sit in a folder the user has just typed in
            import urllib.parse
            q = urllib.parse.parse_qs(self.path.partition('?')[2])
            fd = settings.as_folder((q.get('fd') or [''])[0])
            self.reply({'folder': fd, 'exists': bool(fd) and os.path.isdir(fd),
                        'workbooks': workbooks(fd)})
        elif path in ('/', '/index.html'):
            body = open(page_fp, encoding='utf-8').read().encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_error(404)

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        try:
            body = json.loads(self.rfile.read(length) or b'{}')
        except ValueError as err:
            self.reply({'error': f"bad request: {err}"}, 400)
            return
        try:
            if self.path == '/api/paths':
                settings.save(body.get('paths', {}))
                self.reply({'ok': True, 'said': 'paths saved', 'state': state()})
            elif self.path == '/api/scenarios':
                write_scenarios(settings.load(), body['scenarios'])
                self.reply({'ok': True, 'said': 'create_cases.csv written',
                            'state': state()})
            elif self.path == '/api/run':
                said = launch(body['step'], body.get('scenarios', []),
                              body.get('solver', 'HiGHS'))
                self.reply({'ok': True, 'said': said, 'state': state()})
            else:
                self.send_error(404)
        except Exception as err:
            self.reply({'error': f"{type(err).__name__}: {err}"}, 400)

    def log_message(self, fmt, *args):
        pass                                          #the console stays readable


if __name__ == '__main__':
    s = state()
    print(f"create_cases.csv  {s['paths']['create_cases_fp']}")
    print(f"{len(s['scenarios'])} scenario(s)")
    for sc in s['scenarios']:
        print(f"  {sc['case']:12} built {sc['built'] or '-':>17}   "
              f"solved {sc['solved'] or '-':>17}   results {sc['processed'] or '-':>17}")
    url = f"http://localhost:{PORT}"
    print(f"\nserving {url}   (ctrl-c to stop)")
    print("each step opens its own command window and keeps running if you close this")
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nstopped")
