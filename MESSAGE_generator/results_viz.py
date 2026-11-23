#YL, IAEA, 2026 September
"""
Serve the results of every processed run as a page of charts, locally.

    python results_viz.py

opens http://localhost:8765 in the browser. The page reads its data from this
server rather than carrying a copy, so Refresh picks up any run that has been
processed since it was opened. Nothing leaves the machine.

A run appears here once process_results.py has written its folder under
results/. Which workbook a run came from is read from process_results.csv,
which is also where the transmission links come from.

Nothing is read until the page asks for it. /api/index lists the runs off the
folder names alone, /api/run reads one run, and its dispatch comes in a second
request, so a folder of thirty sensitivities opens as fast as a folder of one.
/api/compare reads the small summary files of every run in one folder, which
is what the comparison page plots.

A run folder may sit one level down, so a folder of sensitivities kept together
under the results folder is found as well; such a run is named
"<folder>/<run>". A "Scenario List.csv" next to those runs is read for what
each one varies, and labels the comparison page.
"""
import csv
import datetime
import http.server
import json
import os
import socketserver
import threading
import urllib.parse
import webbrowser

import pandas as pd

import process_results as pr

PORT = 8765
page_fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'results_viz.html')
#a run carries its dispatch, which is the one big thing here: a handful of them
#at a time is comfortable, every run at once is not
RUN_CACHE = 3

#a fuel with capture is checked before the plain one, and the two storages are
#kept apart because the line charts want them separately
GROUP = [('coal', 'ccs', 'Coal CCS'), ('coal', None, 'Coal'),
         ('gas', 'ccs', 'Gas CCS'), ('gas', None, 'Gas'),
         ('nuclear', None, 'Nuclear'),
         ('biomass', 'ccs', 'Biomass CCS'), ('biomass', None, 'Biomass'),
         ('hydro', None, 'Hydro'),
         ('wind', None, 'Wind'), ('solar', None, 'Solar'),
         ('bat', None, 'Battery'), ('turbpum', None, 'Pumped')]
#the stack order the palette was validated on, with each textured variant next
#to the hue it belongs to
ORDER = ['Coal', 'Coal CCS', 'Gas', 'Gas CCS', 'Nuclear', 'Biomass',
         'Biomass CCS', 'Hydro', 'Wind', 'Solar', 'Battery', 'Pumped']
RENEW = ['Wind', 'Hydro', 'Solar']
COST = [('Table_14_investmentcost', 'Investment'), ('Table_15_fixedcost', 'Fixed O&M'),
        ('Table_16_variablecost', 'Variable O&M')]
FORMS = ['ElectricityVRE', 'ElectricityNonVRE']

#a schematic tile grid, west to east and north to south. not a geographic map:
#no boundary data is used, each province is one square in roughly its place
TILES = {
    'Heilongjiang': (7, 0), 'InnerMongolia': (4, 0),
    'Xinjiang': (0, 1), 'Beijing': (5, 1), 'Tianjin': (6, 1), 'Jilin': (7, 1),
    'Gansu': (2, 2), 'Ningxia': (3, 2), 'Shanxi': (4, 2), 'Hebei': (5, 2),
    'Shandong': (6, 2), 'Liaoning': (7, 2),
    'Qinghai': (1, 3), 'Shaanxi': (3, 3), 'Henan': (4, 3), 'Anhui': (5, 3), 'Jiangsu': (6, 3),
    'Tibet': (0, 4), 'Sichuan': (2, 4), 'Chongqing': (3, 4), 'Hubei': (4, 4),
    'Zhejiang': (6, 4), 'Shanghai': (7, 4),
    'Guizhou': (3, 5), 'Hunan': (4, 5), 'Jiangxi': (5, 5), 'Fujian': (6, 5),
    'Yunnan': (2, 6), 'Guangxi': (3, 6), 'Guangdong': (4, 6),
    'Hainan': (3, 7),
    #not a place: the synthetic region the storage lives in, parked clear of
    #the others so it cannot be mistaken for one
    'Dummy': (0, 7),
}


def group_of(tech):
    """which family a technology is drawn as"""
    tech = str(tech)
    for prefix, needs, name in GROUP:
        if tech.startswith(prefix) and (needs is None or needs in tech):
            return name

    return 'Other'


def season_fractions():
    """
    what share of the year each time slice stands for. the model reports
    energy in MWyr, so a slice value divided by its share is average power
    """
    y0 = 2024
    dates = [datetime.date(y0, m, d) for m, d in pr.mt.season_bounds]
    dates.append(datetime.date(y0 + 1, *pr.mt.season_bounds[0]))
    days = [(dates[i + 1] - dates[i]).days for i in range(len(pr.mt.season_bounds))]

    return [round(d / (sum(days) * 24), 6) for d in days]


def by_group(fd, name, years=None):
    """
    one summary csv rolled up from technologies to families. a file a half
    written run never got to is read as zeros against the years it would have
    had, which the caller passes in
    """
    d = table_of(fd, name)
    if d is None:
        years = years or []
        return years, {g: [0.0] * len(years) for g in ORDER}
    d = d[~d.index.isin(['total'])]
    d.index = [group_of(i) for i in d.index]
    d = d.groupby(level=0).sum()
    years = [int(c) for c in d.columns]

    return years, {g: [round(num(x), 1) for x in d.loc[g]] if g in d.index
                   else [0.0] * len(years) for g in ORDER}


def interconnections(input_fd, input_fn):
    """the transmission links, from the from-to matrix of the workbook"""
    wb = pr.mt.read_workbook(f"{input_fd}{input_fn}.xlsx")
    m = wb['Interconnection'].set_index('from-to')

    return [[str(s).replace(' ', ''), str(d).replace(' ', ''), round(float(v))]
            for s, row in m.iterrows() for d, v in row.items() if pd.notna(v) and v]


def table_of(fd, name):
    """
    one csv of a run, or None if it is not there.

    a run that was interrupted part way through has some of its files and not
    others, and one of those is worth drawing what there is of rather than
    nothing at all
    """
    fp = f"{fd}/{name}.csv"
    if not os.path.exists(fp):
        return None

    return pd.read_csv(fp, index_col=0)


def num(value):
    """
    a number json can carry. a blank or ragged cell reads back as NaN, which
    python writes into json as a bare NaN and the browser then refuses to
    parse: one such cell would otherwise lose the whole run
    """
    try:
        value = float(value)
    except (TypeError, ValueError):
        return 0.0

    return value if value == value and abs(value) != float('inf') else 0.0


def finite(obj):
    """the same payload with every NaN and infinity turned into a zero"""
    if isinstance(obj, float):
        return num(obj)
    if isinstance(obj, dict):
        return {k: finite(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [finite(v) for v in obj]

    return obj


def dumps(obj, **kw):
    """
    json a browser can parse. the fast path is tried first, since a payload
    with nothing odd in it is the usual case and these are large
    """
    try:
        return json.dumps(obj, allow_nan=False, **kw)
    except ValueError:
        return json.dumps(finite(obj), allow_nan=False, **kw)


def add_row(acc, values, n):
    """
    add one row of a table into a per-year accumulator.

    a table of a half written run can be short of years, so only the years it
    has are added and the rest are left at whatever they were. the row is read
    by position, which a repeated row label would otherwise turn into a frame
    """
    for j in range(min(n, len(values))):
        acc[j] += num(values.iloc[j])


def province_groups(fd, name, years, short):
    """
    one of the _province csv files rolled up to {province: {family: [years]}}.
    the page sums these itself when every province is selected, so nothing is
    stored twice
    """
    table = table_of(fd, name)
    if table is None:
        return {}
    rows = {}
    for n, i in enumerate(table.index):
        prov, tech, _ = pr.split_name(pr.SECOND_ACTIVITY.sub('', str(i)))
        if prov not in short:
            continue
        fam = rows.setdefault(short[prov], {}).setdefault(group_of(tech), [0.0] * len(years))
        add_row(fam, table.iloc[n], len(years))

    return {p: {g: [round(v, 1) for v in vs] for g, vs in d.items()}
            for p, d in rows.items()}


def run_core(fd, input_fd, input_fn):
    """
    everything the page plots for a single run bar its dispatch, province by
    province. all of it comes from the summary csv files, which are small
    """
    years, _ = by_group(fd, 'prod_all')
    short = {name[:6]: name for name in TILES}
    out = {'years': years}

    #everything with a province dimension is kept per province. the page adds
    #them up for "all provinces", so a filter needs no round trip
    out['prod'] = province_groups(fd, 'prod_all_province', years, short)
    out['cap'] = province_groups(fd, 'cap_all_province', years, short)
    out['emission'] = {p: [round(sum(v[j] for v in d.values()) / 1000, 3)
                           for j in range(len(years))]
                       for p, d in province_groups(fd, 'emission_province',
                                                   years, short).items()}

    #cost split by kind, which the single cost_all rolls together
    out['cost'] = {}
    for fn, label in COST:
        table = table_of(fd, fn)
        per = {}
        for n, i in enumerate(table.index if table is not None else []):
            prov, _tech, _st = pr.split_name(pr.SECOND_ACTIVITY.sub('', str(i)))
            if prov not in short:
                continue
            acc = per.setdefault(short[prov], [0.0] * len(years))
            add_row(acc, table.iloc[n], len(years))
        out['cost'][label] = {p: [round(v, 1) for v in vs] for p, vs in per.items()}

    #the T&D technologies are named <province>_Transmission_<form>. their
    #capacity is the transmission capacity, and what they produce is what the
    #province is actually delivered
    t12 = table_of(fd, 'Table_12_capall')
    t11 = table_of(fd, 'Table_11_prodall')
    out['trans'] = {}
    for form in FORMS:
        per = {}
        for n, i in enumerate(t12.index if t12 is not None else []):
            if not str(i).endswith('_Transmission_' + form):
                continue
            prov = short.get(str(i).split('_')[0])
            if prov:
                acc = per.setdefault(prov, [0.0] * len(years))
                add_row(acc, t12.iloc[n], len(years))
        out['trans'][form] = {p: [round(v, 1) for v in vs] for p, vs in per.items()}

    out['demand'] = {}
    for n, i in enumerate(t11.index if t11 is not None else []):
        prov = short.get(str(i).split('_')[0])
        if prov and '_Transmission_' in str(i):
            add_row(out['demand'].setdefault(prov, [0.0] * len(years)),
                    t11.iloc[n], len(years))
    out['demand'] = {p: [round(v, 1) for v in vs] for p, vs in out['demand'].items()}

    cp = table_of(fd, 'cap_all_province')
    out['nuclear'] = {}
    for n, i in enumerate(cp.index if cp is not None else []):
        prov = short.get(str(i).split('_')[0])
        if prov and '_nuclear' in str(i):
            add_row(out['nuclear'].setdefault(prov, [0.0] * len(years)),
                    cp.iloc[n], len(years))
    out['nuclear'] = {p: [round(v) for v in vs] for p, vs in out['nuclear'].items()}

    out['provinces'] = sorted(set(out['prod']) | set(out['cap']))
    out['links'] = interconnections(input_fd, input_fn) if input_fn else []
    out['built'] = datetime.datetime.fromtimestamp(
        os.path.getmtime(f"{fd}/prod_all.csv")).strftime('%d %b %Y %H:%M')

    return out


def run_dispatch(fd, links):
    """
    the hourly half of a run: what each technology was dispatched at, the load
    it was dispatched against, and the flow on each interconnector.

    this is the slow part, since profiles.csv is the one large file a run
    writes, which is why the page asks for it on its own
    """
    short = {name[:6]: name for name in TILES}
    out = {'profiles': {}, 'demandProfile': {}, 'flows': {}}
    if not os.path.exists(f"{fd}/profiles.csv"):
        return out                                    #never written, or still being

    #dispatch, province by province so the same filter applies
    p = pd.read_csv(f"{fd}/profiles.csv").copy()
    #a storage runs two activities: one fills the store, the other gives back
    #to the grid. both are kept, with the filling one carried negative so the
    #chart can show charging below the axis. a battery fills on its first
    #activity, a pumped scheme on its second
    #todo: these are the technology names of the current workbook. the type in
    #TechData (BAT, PUM) is what really identifies them, but profiles.csv does
    #not carry it, so a workbook that renames them needs this changing too
    BATTERY, PUMPED = ['bat'], ['turbpum']
    second = p.name.str.endswith('[2.]')
    charging = ((p.technology.isin(BATTERY) & ~second)
                | (p.technology.isin(PUMPED) & second))
    p.loc[charging, 'value'] = -p.loc[charging, 'value']
    p['g'] = p.technology.map(group_of)
    p['prov'] = p.province.map(short)
    prof = {}
    for (prov, y, day), block in p.dropna(subset=['prov']).groupby(['prov', 'year', 'day']):
        s = block.groupby(['g', 'hour']).value.sum().unstack(fill_value=0)
        prof.setdefault(prov, {})[f"{y}-{day}"] = {
            g: [round(num(v), 1) for v in s.loc[g]] if g in s.index else [0.0] * 24
            for g in ORDER if g in s.index}
    out['profiles'] = prof

    #the load those profiles are dispatched against. an input rather than a
    #result, written next to them by process_results.py
    dem = {}
    demand_fp = f"{fd}/demand_profile.csv"
    if os.path.exists(demand_fp):
        dp = pd.read_csv(demand_fp)
        #this file names a province in full, taken from its case folder, where
        #profiles.csv carries the six character prefix the technology names use
        dp['prov'] = dp.province.where(dp.province.isin(TILES))
        for (prov, y, day), block in dp.dropna(subset=['prov']).groupby(['prov', 'year', 'day']):
            hours = block.groupby('hour').value.sum()
            dem.setdefault(prov, {})[f"{y}-{day}"] = [round(num(hours.get(h, 0.0)), 1)
                                                      for h in range(1, 25)]
    out['demandProfile'] = dem

    #hourly flow on each interconnector, seen from each end: a line A_B sends
    #from A, so it is an export for A and an import for B. these rows only
    #appear once MESSAGE_trans has written the interconnector codes into
    #techcodes.csv, so a run built before that simply has none
    named = {f"{a}_{b}": (a, b) for a, b, _mw in links}
    flows = {}
    lines = p[p.technology.isin(named)]
    for (tech, y, day), block in lines.groupby(['technology', 'year', 'day']):
        src, dst = named[tech]
        hours = block.groupby('hour').value.sum()
        series = [round(num(hours.get(h, 0.0)), 1) for h in range(1, 25)]
        key = f"{y}-{day}"
        if src in TILES:
            into = flows.setdefault(src, {}).setdefault(key, {})
            into[dst] = [round(v - x, 1) for v, x in
                         zip(into.get(dst, [0.0] * 24), series)]      #export
        if dst in TILES:
            into = flows.setdefault(dst, {}).setdefault(key, {})
            into[src] = [round(v + x, 1) for v, x in
                         zip(into.get(src, [0.0] * 24), series)]      #import
    out['flows'] = flows

    return out


# --------------------------------------------------------------------------
# which runs are on disk
# --------------------------------------------------------------------------

def is_run(fd):
    """a folder counts as a run once process_results.py has summarised into it"""
    return os.path.isdir(fd) and os.path.exists(f"{fd}/prod_all.csv")


def control_rows():
    """which workbook each run came from, which is only needed for the links"""
    control = {}
    if os.path.exists(pr.control_fp):
        for _, r in pd.read_csv(pr.control_fp).iterrows():
            control[str(r['run'])] = (str(r['input_fd']), str(r['input_fn']))

    return control


def run_index():
    """
    every run under the results folder, named but not read.

    a run usually sits directly under it. a set of sensitivities is normally
    kept together in a folder of its own, so one level further down is looked
    at as well and such a run is named "<folder>/<run>"
    """
    out = []
    if not os.path.isdir(pr.results_fd):
        return out

    for name in sorted(os.listdir(pr.results_fd)):
        fd = os.path.join(pr.results_fd, name)
        if not os.path.isdir(fd):
            continue
        if is_run(fd):
            out.append({'name': name, 'group': '', 'fd': fd})
            continue
        try:
            inside = sorted(os.listdir(fd))
        except OSError:                               #not ours to look in
            continue
        for sub in inside:
            sfd = os.path.join(fd, sub)
            if is_run(sfd):
                out.append({'name': f"{name}/{sub}", 'group': name, 'fd': sfd})

    return out


def folder_of(name):
    """the folder of a named run, refusing anything not in the index"""
    for r in run_index():
        if r['name'] == name:
            return r['fd']

    raise KeyError(f"no run named {name!r}")


def input_of(control, name):
    """
    the workbook a run was built from. a run in a subfolder is looked up by
    its own name as well as the full one, since process_results.csv carries
    the case name alone
    """
    for key in (name, name.split('/')[-1]):
        if key in control:
            return control[key]

    return '', ''


def scenario_table(fd):
    """
    the sensitivity table a folder of runs may carry next to them, as
    {'columns': [...], 'rows': {sensitivity: {column: value}}}.

    the first column names the sensitivity and the rest are what it varies. a
    column with no heading is the note at the end of the row
    """
    found = None
    try:
        inside = sorted(os.listdir(fd))
    except OSError:
        return None
    for fn in inside:
        low = fn.lower()
        if low.endswith('.csv') and 'scenario' in low and 'list' in low:
            found = os.path.join(fd, fn)
            break
    if not found:
        return None

    with open(found, newline='', encoding='utf-8-sig') as f:
        rows = [r for r in csv.reader(f) if any(str(c).strip() for c in r)]
    if len(rows) < 2:
        return None

    head = [str(c).strip() for c in rows[0]]
    #a trailing comma gives a nameless last column, which is where the note is
    names = [h or 'Note' for h in head]
    cols, table = [], {}
    for r in rows[1:]:
        key = str(r[0]).strip()
        if not key:
            continue
        row = {}
        for i, name in enumerate(names[1:], start=1):
            value = str(r[i]).strip() if i < len(r) else ''
            if value:
                row[name] = value
                if name not in cols:
                    cols.append(name)
        table[key] = row

    return {'columns': [c for c in names[1:] if c in cols], 'rows': table,
            'key': names[0], 'file': os.path.basename(found)}


def scenario_tables():
    """
    the scenario list of the results folder and of every folder of runs inside
    it, as one table.

    which folder the viewer works in is the one setting the workbench keeps, so
    the page has no folder of its own to choose: everything under it is read
    together, and a list beside either level labels the runs it names
    """
    found = [scenario_table(pr.results_fd)]
    for group in sorted({r['group'] for r in run_index() if r['group']}):
        found.append(scenario_table(os.path.join(pr.results_fd, group)))

    columns, rows, files = [], {}, []
    for t in [t for t in found if t]:
        for c in t['columns']:
            if c not in columns:
                columns.append(c)
        #first list wins the order; a name in two lists keeps its first place
        for k, v in t['rows'].items():
            rows[k] = v
        files.append(t['file'])

    return {'columns': columns, 'rows': rows,
            'file': ', '.join(dict.fromkeys(files))}


def sensitivity_of(run, keys):
    """
    which row of the scenario list a run is. a run folder is named for the
    case it came from, sn_NW10 for NW10, so the longest name the folder ends
    on is the one meant: NW1 does not claim sn_NW10
    """
    base = run.split('/')[-1]
    best = ''
    for k in keys:
        if (base == k or base.endswith(k)) and len(k) > len(best):
            best = k

    return best


# --------------------------------------------------------------------------
# reading one run, and the small summaries of a whole folder of them
# --------------------------------------------------------------------------

#name -> (mtime, payload). the whole point is not to read a run twice, and
#not to hold more of them than a session needs at once
_core, _dispatch, _summary = {}, {}, {}
_links = {}


def cached(store, key, fd, build, keep=None):
    """
    build once per version of a run. prod_all.csv is written last of the
    summaries, so its stamp says whether what is cached is still the run
    """
    stamp = os.path.getmtime(f"{fd}/prod_all.csv")
    hit = store.get(key)
    if hit and hit[0] == stamp:
        return hit[1]
    value = build()
    store.pop(key, None)                              #so the freshest is last
    store[key] = (stamp, value)
    while keep and len(store) > keep:
        store.pop(next(iter(store)))

    return value


def links_of(input_fd, input_fn):
    """the transmission links of a workbook, read once per workbook"""
    if not input_fn:
        return []
    key = f"{input_fd}|{input_fn}"
    if key not in _links:
        try:
            _links[key] = interconnections(input_fd, input_fn)
        except Exception as err:                      #no workbook, no arrows
            print(f"  no links for {input_fn}: {err}")
            _links[key] = []

    return _links[key]


def run_payload(name, part):
    """one run, either its summaries or its dispatch"""
    fd = folder_of(name)
    control = control_rows()
    input_fd, input_fn = input_of(control, name)
    if part == 'dispatch':
        links = links_of(input_fd, input_fn)
        return cached(_dispatch, name, fd,
                      lambda: run_dispatch(fd, links), RUN_CACHE)

    return cached(_core, name, fd,
                  lambda: run_core(fd, input_fd, input_fn), RUN_CACHE * 4)


def one_run(run, fd, input_fd, input_fn):
    """a whole run, summaries and dispatch together, as a snapshot wants it"""
    out = run_core(fd, input_fd, input_fn)
    out.update(run_dispatch(fd, out['links']))

    return out


def totals(fd, name, mul, years):
    """one summary csv added up over its technologies, year by year"""
    d = table_of(fd, name)
    if d is None:
        return [0.0] * len(years)
    d = d[~d.index.isin(['total'])]
    out = [round(num(v) * mul, 4) for v in d.sum()]

    return (out + [0.0] * len(years))[:len(years)]


def run_light(fd):
    """
    what the comparison page plots for one run: the national summaries only,
    rolled up to families. no province file and no dispatch, so this is a few
    milliseconds a run rather than a few seconds
    """
    years, prod = by_group(fd, 'prod_all')
    _, cap = by_group(fd, 'cap_all', years)

    return {'years': years, 'prod': prod, 'cap': cap,
            #emission.csv is in kt and cost_all.csv in US$
            'emission': totals(fd, 'emission', 1 / 1000, years),
            'cost': totals(fd, 'cost_all', 1.0, years)}


def compare_payload():
    """
    every processed run under the results folder, side by side. the folder is
    whichever one the workbench is set to, so there is nothing to choose here
    """
    rows, skipped = [], []
    table = scenario_tables()
    keys = list(table['rows'])

    for r in run_index():
        sens = sensitivity_of(r['name'], keys)
        try:
            one = cached(_summary, r['name'], r['fd'],
                         lambda fd=r['fd']: run_light(fd))
        except Exception as err:                      #a half written run
            skipped.append(f"{r['name']}: {type(err).__name__} {err}")
            continue
        rows.append({'name': r['name'], 'folder': r['group'],
                     'label': sens or r['name'].split('/')[-1],
                     'sensitivity': sens, 'attrs': table['rows'].get(sens, {}),
                     **one})

    #the scenario list is written in the order the sensitivities belong in, one
    #sweep after another, which reads far better than NW10 before NW2. a run
    #the list does not mention goes after the ones it does
    rows.sort(key=lambda r: (keys.index(r['sensitivity']) if r['sensitivity'] in keys
                             else len(keys), r['name']))

    return {'folder': pr.results_fd, 'runs': rows, 'skipped': skipped,
            'columns': table['columns'], 'listFile': table['file'],
            'refreshed': datetime.datetime.now().strftime('%H:%M:%S')}


def constants():
    """what every page needs to draw, and never changes between runs"""
    return {'order': ORDER, 'renew': RENEW, 'costOrder': [c[1] for c in COST],
            'forms': FORMS, 'tiles': TILES,
            #MWyr is the model's energy unit: one MW held for a year
            'twhPerMWyr': 0.00876, 'sliceFrac': season_fractions()}


def index_payload():
    """
    the runs on disk, off their folder names alone. nothing in a run folder is
    opened here, so this answers at once however many runs there are
    """
    runs = []
    for r in run_index():
        runs.append({'name': r['name'], 'group': r['group'],
                     'built': datetime.datetime.fromtimestamp(
                         os.path.getmtime(f"{r['fd']}/prod_all.csv")
                     ).strftime('%d %b %Y %H:%M')})

    return {**constants(), 'runs': runs, 'resultsFd': pr.results_fd,
            'refreshed': datetime.datetime.now().strftime('%H:%M:%S')}


def payload():
    """
    every run, read in full. this is what a snapshot carries; the page itself
    asks for one run at a time instead
    """
    control = control_rows()
    data = {**constants(), 'refreshed': datetime.datetime.now().strftime('%H:%M:%S'),
            'runs': {}, 'skipped': [], 'resultsFd': pr.results_fd}

    for r in run_index():
        input_fd, input_fn = input_of(control, r['name'])
        try:
            data['runs'][r['name']] = one_run(r['name'], r['fd'], input_fd, input_fn)
        except Exception as err:                      # a half written run
            data['skipped'].append(f"{r['name']}: {type(err).__name__} {err}")
            print(f"  skipping {r['name']}: {err}")

    data['compare'] = compare_payload()
    print(f"  {len(data['runs'])} run(s): {', '.join(data['runs']) or 'none'}")

    return data


class Handler(http.server.BaseHTTPRequestHandler):
    """the page itself, and the data behind it"""

    def send_json(self, build):
        """
        answer with what build() returns, or with why it could not. the page
        asks for one thing at a time, so a run it cannot read leaves the rest
        of the page standing
        """
        try:
            #allow_nan off because a bare NaN is not json and the browser
            #throws the whole answer away; anything that got past num() is
            #written as a zero rather than losing the run over it
            body = dumps(build()).encode('utf-8')
        except Exception as err:
            self.send_error(500, f"{type(err).__name__}: {err}")
            return
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path, _, query = self.path.partition('?')
        q = urllib.parse.parse_qs(query)
        one = lambda k: (q.get(k) or [''])[0]

        #the run list, off the folder names alone
        if path == '/api/index':
            self.send_json(index_payload)
            return

        #one run: its summaries, then its dispatch in a second request
        if path == '/api/run':
            name, part = one('name'), one('part') or 'core'
            self.send_json(lambda: run_payload(name, part))
            return

        #every run under the results folder, summaries only
        if path == '/api/compare':
            self.send_json(compare_payload)
            return

        #kept for a page that has not been reloaded since this was split up
        if path.startswith('/api/data'):
            self.send_json(payload)
            return

        if path in ('/', '/index.html'):
            body = open(page_fp, encoding='utf-8').read().replace('__DATA__', 'null')
            body = body.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        self.send_error(404)

    def log_message(self, fmt, *args):
        pass                                          #the console stays readable


def snapshot(out_fp):
    """
    the same page with the data written into it, for sending to someone who
    cannot run the server
    """
    html = open(page_fp, encoding='utf-8').read()
    html = html.replace('__DATA__', dumps(payload(), separators=(',', ':')))
    open(out_fp, 'w', encoding='utf-8').write(html)
    print(f"wrote {out_fp}")


if __name__ == '__main__':
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == 'snapshot':
        snapshot(sys.argv[2] if len(sys.argv) > 2 else 'results_snapshot.html')
        raise SystemExit

    #the workbench opens the tab itself, so it starts this with --no-browser.
    #without that the two of them open one tab each
    open_tab = '--no-browser' not in sys.argv

    print(f"reading {pr.results_fd}")
    #only the folder names, so the server is listening at once however many
    #runs there are. a run is read when the page asks for it
    found = index_payload()['runs']
    print(f"  {len(found)} run(s): {', '.join(r['name'] for r in found) or 'none'}")
    url = f"http://localhost:{PORT}"
    print(f"serving {url}   (ctrl-c to stop)")
    #threaded, so asking for the next run does not wait on the dispatch of the
    #one before it
    socketserver.ThreadingTCPServer.allow_reuse_address = True
    socketserver.ThreadingTCPServer.daemon_threads = True
    with socketserver.ThreadingTCPServer(("127.0.0.1", PORT), Handler) as httpd:
        if open_tab:
            threading.Timer(0.5, lambda: webbrowser.open(url)).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nstopped")
