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
"""
import datetime
import http.server
import json
import os
import socketserver
import threading
import webbrowser

import pandas as pd

import process_results as pr

PORT = 8765
page_fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'results_viz.html')

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


def by_group(fd, name):
    """one summary csv rolled up from technologies to families"""
    d = pd.read_csv(f"{fd}/{name}.csv", index_col=0)
    d = d[~d.index.isin(['total'])]
    d.index = [group_of(i) for i in d.index]
    d = d.groupby(level=0).sum()
    years = [int(c) for c in d.columns]

    return years, {g: [round(float(x), 1) for x in d.loc[g]] if g in d.index
                   else [0.0] * len(years) for g in ORDER}


def interconnections(input_fd, input_fn):
    """the transmission links, from the from-to matrix of the workbook"""
    wb = pr.mt.read_workbook(f"{input_fd}{input_fn}.xlsx")
    m = wb['Interconnection'].set_index('from-to')

    return [[str(s).replace(' ', ''), str(d).replace(' ', ''), round(float(v))]
            for s, row in m.iterrows() for d, v in row.items() if pd.notna(v) and v]


def province_groups(fd, name, years, short):
    """
    one of the _province csv files rolled up to {province: {family: [years]}}.
    the page sums these itself when every province is selected, so nothing is
    stored twice
    """
    table = pd.read_csv(f"{fd}/{name}.csv", index_col=0)
    rows = {}
    for i in table.index:
        prov, tech, _ = pr.split_name(pr.SECOND_ACTIVITY.sub('', str(i)))
        if prov not in short:
            continue
        fam = rows.setdefault(short[prov], {}).setdefault(group_of(tech), [0.0] * len(years))
        values = table.loc[i]
        for j in range(len(years)):
            fam[j] += float(values.iloc[j])

    return {p: {g: [round(v, 1) for v in vs] for g, vs in d.items()}
            for p, d in rows.items()}


def one_run(run, fd, input_fd, input_fn):
    """everything the page plots for a single run, province by province"""
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
        table = pd.read_csv(f"{fd}/{fn}.csv", index_col=0)
        per = {}
        for i in table.index:
            prov, _tech, _st = pr.split_name(pr.SECOND_ACTIVITY.sub('', str(i)))
            if prov not in short:
                continue
            acc = per.setdefault(short[prov], [0.0] * len(years))
            values = table.loc[i]
            for j in range(len(years)):
                acc[j] += float(values.iloc[j])
        out['cost'][label] = {p: [round(v, 1) for v in vs] for p, vs in per.items()}

    #the T&D technologies are named <province>_Transmission_<form>. their
    #capacity is the transmission capacity, and what they produce is what the
    #province is actually delivered
    t12 = pd.read_csv(f"{fd}/Table_12_capall.csv", index_col=0)
    t11 = pd.read_csv(f"{fd}/Table_11_prodall.csv", index_col=0)
    out['trans'] = {}
    for form in FORMS:
        rows = [i for i in t12.index if str(i).endswith('_Transmission_' + form)]
        per = {}
        for i in rows:
            prov = short.get(str(i).split('_')[0])
            if prov:
                per[prov] = [round(float(v), 1) for v in t12.loc[i]]
        out['trans'][form] = per

    delivered = t11.loc[[i for i in t11.index if '_Transmission_' in str(i)]].copy()
    delivered.index = [str(i).split('_')[0] for i in delivered.index]
    delivered = delivered.groupby(level=0).sum()
    out['demand'] = {short[k]: [round(float(v), 1) for v in delivered.loc[k]]
                     for k in delivered.index if k in short}

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
            g: [round(float(v), 1) for v in s.loc[g]] if g in s.index else [0.0] * 24
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
            dem.setdefault(prov, {})[f"{y}-{day}"] = [round(float(hours.get(h, 0.0)), 1)
                                                      for h in range(1, 25)]
    out['demandProfile'] = dem

    cp = pd.read_csv(f"{fd}/cap_all_province.csv", index_col=0)
    nuc = cp.loc[[i for i in cp.index if '_nuclear' in str(i)]].copy()
    nuc.index = [str(i).split('_')[0] for i in nuc.index]
    nuc = nuc.groupby(level=0).sum()
    out['nuclear'] = {short[k]: [round(float(v)) for v in nuc.loc[k]]
                      for k in nuc.index if k in short}

    out['provinces'] = sorted(set(out['prod']) | set(out['cap']))
    out['links'] = interconnections(input_fd, input_fn) if input_fn else []

    #hourly flow on each interconnector, seen from each end: a line A_B sends
    #from A, so it is an export for A and an import for B. these rows only
    #appear once MESSAGE_trans has written the interconnector codes into
    #techcodes.csv, so a run built before that simply has none
    named = {f"{a}_{b}": (a, b) for a, b, _mw in out['links']}
    flows = {}
    lines = p[p.technology.isin(named)]
    for (tech, y, day), block in lines.groupby(['technology', 'year', 'day']):
        src, dst = named[tech]
        hours = block.groupby('hour').value.sum()
        series = [round(float(hours.get(h, 0.0)), 1) for h in range(1, 25)]
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
    out['built'] = datetime.datetime.fromtimestamp(
        os.path.getmtime(f"{fd}/prod_all.csv")).strftime('%d %b %Y %H:%M')

    return out


def payload():
    """
    every run that has been processed, read fresh off disk.

    a folder counts as a run once it has the summaries in it. process_results.csv
    says which workbook each came from, which is only needed for the links
    """
    control = {}
    if os.path.exists(pr.control_fp):
        for _, r in pd.read_csv(pr.control_fp).iterrows():
            control[str(r['run'])] = (str(r['input_fd']), str(r['input_fn']))

    data = {'order': ORDER, 'renew': RENEW, 'costOrder': [c[1] for c in COST],
            'forms': FORMS, 'tiles': TILES,
            #MWyr is the model's energy unit: one MW held for a year
            'twhPerMWyr': 0.00876, 'sliceFrac': season_fractions(),
            'refreshed': datetime.datetime.now().strftime('%H:%M:%S'),
            'runs': {}, 'skipped': []}

    if not os.path.isdir(pr.results_fd):
        return data

    for run in sorted(os.listdir(pr.results_fd)):
        fd = os.path.join(pr.results_fd, run)
        if not os.path.isdir(fd) or not os.path.exists(f"{fd}/prod_all.csv"):
            continue
        input_fd, input_fn = control.get(run, ('', ''))
        try:
            data['runs'][run] = one_run(run, fd, input_fd, input_fn)
        except Exception as err:                      # a half written run
            data['skipped'].append(f"{run}: {type(err).__name__} {err}")
            print(f"  skipping {run}: {err}")

    print(f"  {len(data['runs'])} run(s): {', '.join(data['runs']) or 'none'}")

    return data


class Handler(http.server.BaseHTTPRequestHandler):
    """the page itself, and the data behind it"""

    def do_GET(self):
        if self.path.startswith('/api/data'):
            try:
                body = json.dumps(payload()).encode('utf-8')
            except Exception as err:
                self.send_error(500, f"{type(err).__name__}: {err}")
                return
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if self.path in ('/', '/index.html'):
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
    html = html.replace('__DATA__', json.dumps(payload(), separators=(',', ':')))
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
    payload()                                         #so problems show at startup
    url = f"http://localhost:{PORT}"
    print(f"serving {url}   (ctrl-c to stop)")
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
        if open_tab:
            threading.Timer(0.5, lambda: webbrowser.open(url)).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nstopped")
