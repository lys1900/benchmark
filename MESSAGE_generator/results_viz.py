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

#coal with capture is checked before plain coal, and the two storages are kept
#apart because the line charts want them separately
GROUP = [('coal', 'ccs', 'Coal CCS'), ('coal', None, 'Coal'),
         ('gas', None, 'Gas'), ('nuclear', None, 'Nuclear'),
         ('biomass', None, 'Biomass'), ('hydro', None, 'Hydro'),
         ('wind', None, 'Wind'), ('solar', None, 'Solar'),
         ('bat', None, 'Battery'), ('turbpum', None, 'Pumped')]
#the stack order the palette was validated on, with the two textured variants
#next to the hue they belong to
ORDER = ['Coal', 'Coal CCS', 'Gas', 'Nuclear', 'Biomass', 'Hydro', 'Wind',
         'Solar', 'Battery', 'Pumped']
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


def one_run(run, fd, input_fd, input_fn):
    """everything the page plots for a single run"""
    years, prod = by_group(fd, 'prod_all')
    _, cap = by_group(fd, 'cap_all')
    out = {'years': years, 'prod_all': prod, 'cap_all': cap}

    out['cost'] = {}
    for fn, label in COST:
        t = pd.read_csv(f"{fd}/{fn}.csv", index_col=0)
        keep = [i for i in t.index if pr.split_name(str(i))[0] is not None]
        out['cost'][label] = [round(float(v), 1) for v in t.loc[keep].sum()]

    #the T&D technologies are named <province>_Transmission_<form>. what they
    #produce is what reaches the distribution level, so their capacity is the
    #transmission capacity and their output is the electricity a province is
    #actually delivered
    t12 = pd.read_csv(f"{fd}/Table_12_capall.csv", index_col=0)
    t11 = pd.read_csv(f"{fd}/Table_11_prodall.csv", index_col=0)
    out['trans'] = {f: [round(float(v), 1) for v in
                        t12.loc[[i for i in t12.index
                                 if str(i).endswith('_Transmission_' + f)]].sum()]
                    for f in FORMS}

    short = {name[:6]: name for name in TILES}
    delivered = t11.loc[[i for i in t11.index if '_Transmission_' in str(i)]].copy()
    delivered.index = [str(i).split('_')[0] for i in delivered.index]
    delivered = delivered.groupby(level=0).sum()
    out['provDemand'] = {short[k]: [round(float(v), 1) for v in delivered.loc[k]]
                         for k in delivered.index if k in short}

    e = pd.read_csv(f"{fd}/emission.csv", index_col=0)
    out['emission'] = [round(float(x) / 1000, 1) for x in e.loc['total']]

    p = pd.read_csv(f"{fd}/profiles.csv")
    second = p.name.str.endswith('[2.]')
    p = p[~((p.technology.eq('bat') & ~second) | (p.technology.eq('turbpum') & second))].copy()
    p['g'] = p.technology.map(group_of)
    prof = {}
    for y in sorted(p.year.unique()):
        for day in sorted(p.day.unique()):
            s = (p[(p.year == y) & (p.day == day)]
                 .groupby(['g', 'hour']).value.sum().unstack(fill_value=0))
            prof[f"{y}-{day}"] = {g: [round(float(v), 1) for v in s.loc[g]] if g in s.index
                                  else [0.0] * 24 for g in ORDER}
    out['profiles'] = prof

    cp = pd.read_csv(f"{fd}/cap_all_province.csv", index_col=0)
    nuc = cp.loc[[i for i in cp.index if '_nuclear' in str(i)]].copy()
    nuc.index = [str(i).split('_')[0] for i in nuc.index]
    nuc = nuc.groupby(level=0).sum()
    out['nuclear'] = {short[k]: [round(float(v)) for v in nuc.loc[k]]
                      for k in nuc.index if k in short}

    pp = pd.read_csv(f"{fd}/prod_all_province.csv", index_col=0)
    rows = {}
    for i in pp.index:
        prov, tech, _ = pr.split_name(pr.SECOND_ACTIVITY.sub('', str(i)))
        if prov not in short:
            continue
        fam = rows.setdefault(short[prov], {}).setdefault(group_of(tech), [0.0] * len(years))
        for j in range(len(years)):
            fam[j] += float(pp.iloc[:, j][i])
    out['provProd'] = {p_: {g: [round(v, 1) for v in vs] for g, vs in d.items()}
                       for p_, d in rows.items()}

    out['links'] = interconnections(input_fd, input_fn) if input_fn else []
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
