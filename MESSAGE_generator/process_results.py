#YL, IAEA, 2026 September
"""
Turn the yearly result tables of a solved model into csv files, doing the work
the results excel template used to do by hand.

Run it after solving. process_results.csv says which runs to read: a run name
and the folder of the multiregional parent case. Everything for one run is
written to results/<run name>/.

Per run it writes:
    Table_<n>_<name>.csv        the raw tables, as they appear in the res file
    prod_all.csv                production by technology and year
    cap_all.csv                 capacity by technology and year
    capadd_all.csv              capacity additions by technology and year
    cost_all.csv                investment plus fixed plus variable cost
    emission.csv                production times the emission factor
    <name>_long.csv             the same, one row per technology and year
    <name>_province.csv         the rows those were summed from
    profiles.csv                production in every time slice, long

Capacity is not taken from the model's own capall table, which is wrong.
It is rebuilt from the capacity additions, the historic capacity in
hist_tab.csv and the lifetimes in lt_tab.csv, both written by MESSAGE_trans.py
into the parent case folder. See capacity_by_year for how.

The profiles come from the solution file rather than the res file, which only
carries yearly totals. See write_profiles for how a technology is found in it.
"""
import csv
import os
import re

import pandas as pd

import MESSAGE_trans as mt
import settings

#one row per run, columns: run, path, input_fd, input_fn. see settings.py for
#where these come from and how the workbench changes them
_paths = settings.load()
control_fp = _paths['process_results_fp']
results_fd = _paths['results_fd']

TABLE_START = re.compile(r'^\s*"?Table\s+(\d+):')
#the res tables name a technology as <province>_<technology>_<status>
MSG_NAME = re.compile(r'^([^_]+)_(.+)_([^_]+)$')
#a technology with a second activity is reported once per activity, the
#second one marked like Dummy_bat_generic[2.]
SECOND_ACTIVITY = re.compile(r'\[2\.\]$')
#batteries take from the grid on their first activity, see drop_other_activity
BAT_TYPE = 'BAT'


def read_tables(path):
    """
    read the Table blocks of a res file
    :return: {table number: {'name': str, 'years': [...], 'columns': [...],
                             'values': {column: {year: float}}}}
    """
    lines = open(path, encoding='utf-8', errors='replace').read().splitlines()
    starts = [i for i, l in enumerate(lines) if TABLE_START.match(l)]
    tables = {}
    for k, i in enumerate(starts):
        number = int(TABLE_START.match(lines[i]).group(1))
        end = starts[k + 1] if k + 1 < len(starts) else len(lines)
        block = [l for l in lines[i + 1:end] if l.strip()]
        if len(block) < 2:
            continue

        #first line carries the date and the table name, second the columns
        name = strip_cells(block[0])[-1]
        columns = unique_columns(strip_cells(block[1])[1:])
        years, values = [], {c: {} for c in columns}
        for row in block[2:]:
            cells = strip_cells(row)
            #a data row starts with the year, anything else is a stray line
            if not cells or not re.match(r'^-?\d+(\.\d+)?$', cells[0]):
                continue
            year = int(float(cells[0]))
            years.append(year)
            for c, cell in zip(columns, cells[1:]):
                values[c][year] = float(cell) if cell else 0.0
        tables[number] = {'name': name, 'years': years,
                          'columns': columns, 'values': values}

    return tables


def unique_columns(columns):
    """
    a res table names an interconnector once from each end, so the same column
    name turns up twice with different values. keeping them apart as
    Hebei_Beijing and Hebei_Beijing.2 means neither is silently dropped
    """
    seen, out = {}, []
    for c in columns:
        seen[c] = seen.get(c, 0) + 1
        out.append(c if seen[c] == 1 else f"{c}.{seen[c]}")

    return out


def strip_cells(line):
    """split a tab separated line and take the quotes and spaces off each cell"""
    return [c.strip().strip('"').strip() for c in line.replace('\xa0', ' ').split('\t')]


def table_frame(table):
    """one table as a dataframe of columns by year"""
    df = pd.DataFrame(table['values']).T
    df.index.name = 'name'

    return df.reindex(columns=table['years'])


def split_name(name):
    """
    Anhui_solarpv_Exist -> (Anhui, solarpv, Exist). names which are not in
    that shape, the interconnectors for instance, come back as (None, ...)
    """
    m = MSG_NAME.match(str(name))
    if not m:
        return None, str(name), None

    return m.group(1), m.group(2), m.group(3)


def by_technology(df, types):
    """sum a table of technology rows by technology, dropping province and status"""
    if df.empty:
        return df
    out = df.copy()
    out.insert(0, 'tech', [split_name(SECOND_ACTIVITY.sub('', str(n)))[1]
                           for n in df.index])

    return out.groupby('tech', sort=True).sum()


def drop_other_activity(keep, types):
    """
    a storage technology runs two activities and a table which reports both
    would count it twice, once into the grid and once into its own store.

    only the activity facing the grid is kept: a battery discharges on its
    second activity, a pumped scheme turbines on its first. where a table
    carries only one of the two, capacity for instance, that one is kept
    """
    both = {}
    for name in keep:
        both.setdefault(SECOND_ACTIVITY.sub('', str(name)), []).append(name)
    out = {}
    for base, names in both.items():
        if len(names) < 2:
            out[names[0]] = keep[names[0]]
            continue
        second = [n for n in names if SECOND_ACTIVITY.search(str(n))]
        first = [n for n in names if not SECOND_ACTIVITY.search(str(n))]
        wanted = second if types[keep[names[0]]] == BAT_TYPE else first
        for n in wanted:
            out[n] = keep[n]

    return out


def number(value):
    """a cell as a float, an empty one as zero"""
    return 0.0 if pd.isna(value) else float(value)


def year_weights(years):
    """
    how many years each milestone stands for. the first counts once, the rest
    take half the gap before and half the gap after.

    this has to be given the whole horizon of the Years sheet, not the years a
    res file happens to report. a run which stops early still has the gap after
    its last reported year, so 2025 in a 2024, 2025, 2030 horizon stands for
    (1 + 5) * 0.5 = 3 years whether or not 2030 was solved.

    the last year of the sheet has no gap after it and its own gap before is
    used instead. the template gets the same answer by carrying one year past
    the end of the horizon, 2070 for results which stop at 2065
    """
    weights = {}
    for i, y in enumerate(years):
        if i == 0:
            weights[y] = 1
            continue
        before = y - years[i - 1]
        after = years[i + 1] - y if i + 1 < len(years) else before
        weights[y] = (before + after) * 0.5

    return weights


def capacity_by_year(capadd, hist_tab, lt_tab, years, model_years):
    """
    capacity in each milestone year, rebuilt rather than taken from the model.

    additions in a milestone year are spread over the years that milestone
    stands for, anything before the first milestone comes from hist_tab, and a
    unit retires once it has been running for its lifetime. the running total
    of additions less retirements is the capacity.

    :param years: the milestones the res file reports, which are the columns
    :param model_years: the whole horizon of the Years sheet, which is what
                        says how many years each milestone stands for
    """
    weights = year_weights(sorted(set(model_years) | set(years)))
    first_year = years[0]
    hist_years = [int(c) for c in hist_tab.columns]
    calendar = range(min(hist_years + [first_year]), max(years) + 1)

    capacity = pd.DataFrame(0.0, index=capadd.index, columns=years)
    for name in capadd.index:
        #additions per calendar year
        added = {y: 0.0 for y in calendar}
        for y in years:
            if y in added:
                added[y] += number(capadd.loc[name, y]) * weights[y]
        if name in hist_tab.index:
            for y in hist_years:
                if y < first_year and y in added:
                    #an empty cell of hist_tab is no capacity, not a missing
                    #one. left as a nan it would carry through the whole
                    #running total below and every year would come out blank
                    added[y] += number(hist_tab.loc[name, y])

        lifetime = lt_tab.get(name)
        running = 0.0
        for y in calendar:
            running += added[y]
            if lifetime and (y - lifetime) in added:
                running -= added[y - lifetime]     #retires after its lifetime
            if y in capacity.columns:
                capacity.loc[name, y] = running

    return capacity


def read_hist_lt(case_fd):
    """
    hist_tab.csv and lt_tab.csv as MESSAGE_trans.py writes them: hist_tab has
    a row per year and a column per technology, lt_tab a single row
    """
    hist = pd.read_csv(f"{case_fd}/hist_tab.csv", index_col=0)
    hist = hist.T                                  #technologies as rows
    hist.columns = [int(float(c)) for c in hist.columns]

    lt = pd.read_csv(f"{case_fd}/lt_tab.csv", index_col=0)
    lifetimes = {c: float(lt[c].iloc[0]) for c in lt.columns if pd.notna(lt[c].iloc[0])}

    return hist, lifetimes


def workbook_data(input_fd, input_fn):
    """
    what the summaries need from the workbook the model was built from
    :return: (whole horizon, type per technology name,
              emission factor per technology name, efficiency per technology name)
    """
    wb = mt.read_workbook(f"{input_fd}{input_fn}.xlsx")
    years = [int(y) for y in wb["Years"]['years']]
    td = mt.strstrip(wb["TechData"], ['Technology name']).set_index('Technology name')
    types = {name: str(v).strip() for name, v in td['type'].items()}
    factors = {name: float(v) for name, v in td['emissionfactor'].items() if pd.notna(v)}
    efficiency = {name: float(v) for name, v in td['efficiency'].items() if pd.notna(v)}

    return years, types, factors, efficiency


def write_wide(path, df, label='technology'):
    """a table of rows by year"""
    out = df.copy()
    out.index.name = label
    out.to_csv(path)
    print(f"   wrote {os.path.basename(path)}  ({len(out)} rows)")


def write_long(path, df, label='technology'):
    """the same table with one row per cell"""
    out = df.copy()
    out.index.name = label
    long = out.reset_index().melt(id_vars=label, var_name='year', value_name='value')
    long.to_csv(path, index=False)
    print(f"   wrote {os.path.basename(path)}  ({len(long)} rows)")


# --------------------------------------------------------------------------
# one run
# --------------------------------------------------------------------------

#the res tables this script uses, by the name the table carries. these are the
#per subregion ones, which the .cin asks for with an !aggr line per case. the
#parent also reports its own 'investment cost' and so on, but those have no
#technology columns
PROD_TABLE = 'prodall'
CAPADD_TABLE = 'capadd'
COST_TABLES = ['investmentcost', 'fixedcost', 'variablecost']


def find_table(tables, name):
    """the table with this name, or None"""
    for number, t in tables.items():
        if t['name'].strip().lower() == name.lower():
            return t

    return None


def technology_rows(df, types):
    """
    keep the rows which are a technology of the workbook, drop the rest.

    the transmission links, the interconnectors and the storage dummies are
    all named like province_technology_status too, so the shape of the name is
    not enough to tell them apart. the results template did the same thing by
    listing the technologies it wanted by hand
    """
    keep = {}                          #row name -> technology it counts under
    for name in df.index:
        province, tech, status = split_name(SECOND_ACTIVITY.sub('', str(name)))
        if province is not None and tech in types:
            keep[name] = tech

    return df.loc[list(drop_other_activity(keep, types))]


# --------------------------------------------------------------------------
# the hourly profiles, which are in the solution file and not in the res file
# --------------------------------------------------------------------------

#a column of the solution file, the name on its own line when it is too long
#for the column the solver lays out, which every technology name is
SOL_NAME = re.compile(r'^\s*\d+\s+(\S+)\s*$')
SOL_BOTH = re.compile(r'^\s*\d+\s+(\S+)\s+[A-Z*]+\s+(\S+)')
#<code>.<time slice><year>, as in fmag....a.aaa024
SOL_COLUMN = re.compile(r'^(\S+)\.([a-z]{3})(\d{3})$')


def read_tech_codes(case_fd):
    """
    what each technology is called in the matrix, from the techcodes.csv
    MESSAGE_trans.py writes
    :return: {code: technology}
    """
    path = f"{case_fd}/techcodes.csv"
    if not os.path.exists(path):
        return {}
    codes = pd.read_csv(path)

    return dict(zip(codes['code'], codes['technology']))


def sol_path(case_fd):
    """the solution file of this case, whatever the solver named it"""
    intm = f"{case_fd}/intm"
    if not os.path.isdir(intm):
        return None
    sols = sorted(f for f in os.listdir(intm) if f.lower().endswith('.sol'))

    return f"{intm}/{sols[0]}" if sols else None


def read_sol_activity(path, codes, year_of):
    """
    the activity of every column belonging to one of the technologies.

    the solution file is laid out for reading, so a name too long for its
    column pushes the numbers onto the next line. every technology column is
    too long, so nearly all of them are split in two

    :return: {(technology, year): {time slice: activity}}
    """
    series = {}

    def take(full, value):
        m = SOL_COLUMN.match(full)
        if not m:
            return
        code, slice_, digits = m.groups()
        if code not in codes or digits not in year_of:
            return
        series.setdefault((codes[code], year_of[digits]), {})[slice_] = value

    lines = open(path, encoding='utf-8', errors='replace').read().splitlines()
    #the rows come first and use the same tokens with a letter in front, so
    #only the columns after this header are read
    start = next((i for i, l in enumerate(lines) if 'Column name' in l), 0)
    pending = None
    for line in lines[start + 2:]:
        if pending is not None:
            cells = line.split()
            #status then activity, the activity being the first number
            if len(cells) > 1:
                try:
                    take(pending, float(cells[1]))
                except ValueError:
                    pass
            pending = None
            continue
        m = SOL_BOTH.match(line)
        if m:
            try:
                take(m.group(1), float(m.group(2)))
            except ValueError:
                pass
            continue
        m = SOL_NAME.match(line)
        if m:
            pending = m.group(1)

    return series


def write_profiles(out_fd, case_fd, model_years, efficiency):
    """
    production in every time slice, one row per technology, year and slice.

    the solution file holds the activity of a technology, which is what it
    burns. production is that times the efficiency, which was checked against
    the yearly totals of the res file and agrees exactly for every technology

    the three letters after the code say when it is: the first is the day, the
    third the hour, the middle one the day type, of which there is only one
    """
    codes = read_tech_codes(case_fd)
    path = sol_path(case_fd)
    if not codes or path is None:
        print("   no techcodes.csv or no solution file, no profiles written")
        return

    year_of = {f"{y % 1000:03d}": y for y in model_years}
    series = read_sol_activity(path, codes, year_of)
    if not series:
        print(f"   nothing in {os.path.basename(path)} matched a technology")
        return

    rows = []
    for (name, year), values in series.items():
        province, tech, status = split_name(name)
        factor = float(efficiency.get(tech, 1) or 1)
        for slice_, activity in sorted(values.items()):
            rows.append((name, province, tech, status, year, slice_,
                         ord(slice_[0]) - ord('a') + 1, ord(slice_[2]) - ord('a') + 1,
                         activity * factor))

    out = pd.DataFrame(rows, columns=['name', 'province', 'technology', 'status',
                                      'year', 'slice', 'day', 'hour', 'value'])
    out.to_csv(f"{out_fd}/profiles.csv", index=False)
    print(f"   wrote profiles.csv  ({len(out)} rows, "
          f"{out['name'].nunique()} technologies, from {os.path.basename(path)})")


def process_run(run, case_fd, input_fd, input_fn):
    """read one solved case and write its tables and summaries"""
    case_name = os.path.basename(os.path.normpath(case_fd))
    res_fp = f"{case_fd}/res/{case_name}_adb.txt"
    if not os.path.exists(res_fp):
        print(f"   no result file at {res_fp}, skipping")
        return

    out_fd = f"{results_fd}{run}"
    os.makedirs(out_fd, exist_ok=True)
    tables = read_tables(res_fp)
    print(f"   {len(tables)} tables, years {tables[1]['years'] if tables else '?'}")

    #the tables as they come
    for number, t in sorted(tables.items()):
        safe = re.sub(r'[^A-Za-z0-9]+', '_', t['name']).strip('_').lower()
        write_wide(f"{out_fd}/Table_{number}_{safe}.csv", table_frame(t), label='name')

    #production, additions and cost by technology
    prod = find_table(tables, PROD_TABLE)
    capadd = find_table(tables, CAPADD_TABLE)
    if prod is None or capadd is None:
        print(f"   no {PROD_TABLE} or {CAPADD_TABLE} table, no summaries written")
        return
    years = prod['years']
    model_years, types, factors, efficiency = workbook_data(input_fd, input_fn)
    if [y for y in years if y not in model_years]:
        print(f"   warning: the res file reports {years}, which the Years sheet "
              f"{model_years} does not cover")
    elif years != model_years:
        print(f"   the run stops at {years[-1]}, the horizon runs to {model_years[-1]}")

    prod_tech = technology_rows(table_frame(prod), types)
    capadd_tech = technology_rows(table_frame(capadd), types)
    write_summary(out_fd, 'prod_all', prod_tech, types)
    write_summary(out_fd, 'capadd_all', capadd_tech, types)

    cost_tech = None
    for name in COST_TABLES:
        t = find_table(tables, name)
        if t is None:
            print(f"   warning: no '{name}' table, it is left out of cost_all")
            continue
        part = technology_rows(table_frame(t), types)
        cost_tech = part if cost_tech is None else cost_tech.add(part, fill_value=0)
    if cost_tech is not None:
        write_summary(out_fd, 'cost_all', cost_tech, types)

    #capacity, rebuilt from the additions rather than read from the model
    hist_tab, lt_tab = read_hist_lt(case_fd)
    lt_by_name = {n: lt_tab.get(n) for n in capadd_tech.index}
    missing_lt = [n for n, v in lt_by_name.items() if not v]
    if missing_lt:
        print(f"   warning: no lifetime for {len(missing_lt)} technologies, "
              f"eg {missing_lt[:3]}, they never retire")
    capacity = capacity_by_year(capadd_tech, hist_tab, lt_by_name, years, model_years)
    write_summary(out_fd, 'cap_all', capacity, types)

    #emissions, production times the factor of that technology
    emis_tech = emission_rows(prod_tech, factors)
    if emis_tech.empty:
        print("   no technology has an emission factor")
    else:
        emis = by_technology(emis_tech, types)
        emis.loc['total'] = emis.sum()
        write_summary(out_fd, 'emission', emis_tech, types, aggregated=emis)

    #the profiles, which the res file does not carry
    write_profiles(out_fd, case_fd, model_years, efficiency)


def emission_rows(prod_tech, factors):
    """production times the factor of its technology, still province by province"""
    keep = [n for n in prod_tech.index
            if split_name(SECOND_ACTIVITY.sub('', str(n)))[1] in factors]
    out = prod_tech.loc[keep].copy()
    for name in keep:
        tech = split_name(SECOND_ACTIVITY.sub('', str(name)))[1]
        out.loc[name] = out.loc[name] * factors[tech]

    return out


def write_summary(out_fd, name, per_province, types, aggregated=None):
    """
    one summary, three ways: by technology wide and long, and the province by
    province rows it was summed from, in the same shape as the Table files
    """
    if aggregated is None:
        aggregated = by_technology(per_province, types)
    write_wide(f"{out_fd}/{name}.csv", aggregated)
    write_long(f"{out_fd}/{name}_long.csv", aggregated)
    write_wide(f"{out_fd}/{name}_province.csv", per_province, label='name')


if __name__ == '__main__':
    runs = pd.read_csv(control_fp)
    for _, r in runs.iterrows():
        print(f"=== {r['run']} ===")
        process_run(r['run'], r['path'], r['input_fd'], r['input_fn'])
