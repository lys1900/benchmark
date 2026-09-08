#YL, IAEA, 2026 August
"""
Add batteries and pumped storage to cases which MESSAGE_trans.py has just
built. MESSAGE_trans.py calls this itself for every scenario whose add_storage
column is 1, so there is only one script to run. Setting that column to 0
leaves the model without storage, and this file can also be run on its own to
add storage to cases which are already built.

The storage technologies are the rows of TechData whose type is BAT or PUM.
Their shared parameters (lifetime, capex, fom, vom) come from the usual
sheets, the storage specific ones from StorageData and HydroData, and which
province may build them from the generics sheet, exactly like any other
generic technology.

MESSAGE writes storage with syntax this generator does not use elsewhere:
a second activity per technology, consa coupling to a storage relation, and
per day dummy technologies which carry the state from one time slice to the
next. That structure is reproduced here from a working model.
"""
import os
import re

import pandas as pd

import MESSAGE_trans as mt

#the storage blocks reference relations by a 4 character code. the first
#character is the index of the storage, so several can live in one province
BAT_TYPE = 'BAT'
PUM_TYPE = 'PUM'


def storage_of_province(cfg, province_long):
    """
    which storage technologies this province may build
    :return: list of (tech_name, type) in the order they appear in generics
    """
    out = []
    for t_ in cfg['generics_dict'].get(province_long, {}):
        tech_name = re.sub(r'[^A-Za-z0-9]', '', t_).lower()
        if cfg['tech_param'][tech_name]['type'] in (BAT_TYPE, PUM_TYPE):
            out.append((tech_name, cfg['tech_param'][tech_name]['type']))

    return out


def free_form_letter(adb_lines, used):
    """
    a letter for a new energy form which no level or form already uses
    """
    taken = set(used)
    for line in adb_lines:
        parts = line.split()
        if len(parts) in (2, 3) and len(parts[1]) == 1 and parts[1].isalpha():
            taken.add(parts[1])
    for letter in 'pqrstuvwxyz':
        if letter not in taken:
            return letter
    raise SystemExit("no free energy form letter left")


def matrix_code(out_pair, in_pair, activity):
    """
    what MESSAGE calls a technology in the matrix, which is how the results
    script finds its time slices in the solution file. it is the level the
    technology produces on, the form it consumes or a dot when it consumes
    nothing, its activity letter and the form it produces, padded to eight
    characters. the subregion letter goes on the end once the case is known,
    then the time slice and the year, as in fecq....F.aaa025

    MESSAGE_trans.py builds the same thing inline for the ordinary
    technologies. checked against the solution file for both activities of a
    battery and of a pumped scheme
    """
    out_form, out_level = out_pair.split('-')
    in_form = in_pair.split('-')[0] if in_pair else '.'

    return f"{out_level}{in_form}{activity}{out_form}...."


def capfac_block(name, activity, year0, days_year, per_day):
    """
    a capacity factor curve in the shape MESSAGE_trans already writes
    :param per_day: one list of slice values per representative day
    """
    s = (f"systems.{name}.{activity}.capfac\n"
         f"{year0}\n")
    s += f"{' '.join(str(i) for i in [1] * days_year)}\n"
    for day in per_day:
        s += ("1.000000\n"
              f"{' '.join(f'{v:.6f}' for v in day)}\n")

    return s


def capfac_line(name, activity, per_day):
    """the same curve as one line, which is how the .adb carries it"""
    flat = [v for day in per_day for v in day]

    return f"systems.{name}.{activity}.capfac {' '.join(f'{v:.6f}' for v in flat)}\n"


def battery_blocks(cfg, province, tech_name, idx, form_letter, letters):
    """
    every piece a battery adds. see the module docstring for why the dummies
    exist: one per representative day, holding the state of charge across it
    :param letters: activity letters to draw from. MESSAGE names a technology
                    after its input form, activity letter and output form, so
                    these all need a letter of their own or it reports
                    "Technology id already used"
    """
    row = cfg['StorageData'].loc[tech_name]
    tp = cfg['tech_param'][tech_name]
    days, steps = cfg['days_year'], cfg['timesteps_day']
    act_in, act_out = next(letters), next(letters)
    act_helper = next(letters)
    act_dummy = [(next(letters), next(letters)) for _ in range(days)]
    name = f"{province}_{tech_name}_generic"
    form = f"{province}{row['form suffix']}"
    lvl_in = cfg['ef_pair'][row['charge from']]
    lvl_sto = f"{form_letter}-{cfg['ef']['Transmission']}"

    parts = {'energyforms': f"    {form} {form_letter} l \n    #\n"}

    parts['relationss'] = (
        f"StoBat{idx} {idx}StB o 0\n"
        f"    units\ttype: energy, cost:US$'00/kWyr, inv:US$'00/kW, fom:US$'00/kW/yr, pll:yr, cmix:MW, hisccap:MW, ctime:yr\n"
        f"    for_ldr\tall\n"
        f"    upper\tc 9999999\n"
        f"    stortype\tcontinuous\n"
        f"    type\tNone\n"
        f"    con1a\t{idx}Rel c 1\n"
        f"*\n")

    #each block ends on its own newline and none starts with one, so that the
    #blocks join without leaving the blank lines MESSAGE strips out
    rel1 = (f"StoTech{idx} {idx}Sto o\n"
            f"    units\tgroup: activity, type: energy, cost:US$'00/kWyr, upper:MWyr, lower:MWyr\n"
            f"    for_ldr\tnone\n"
            f"    upper\tc 0\n"
            f"    lower\tc 0\n"
            f"    type\tNone\n"
            f"*\n"
            f"RelBat{idx} {idx}Rel o 0\n"
            f"    units\tgroup: activity, type: energy, cost:US$'00/kWyr, upper:MWyr, lower:MWyr\n"
            f"    for_ldr\tall\n"
            f"    upper\tc 0\n"
            f"    type\tNone\n"
            f"*\n")
    for d in range(1, days + 1):
        rel1 += (f"Dummy{idx}{d} {idx}Dm{d} o\n"
                 f"    units\tgroup: activity, type: energy, cost:US$'00/kWyr, upper:MWyr, lower:MWyr\n"
                 f"    for_ldr\tnone\n"
                 f"    upper\tc 0\n"
                 f"    lower\tc 0\n"
                 f"    type\tNone\n"
                 f"*\n")
    parts['relations1'] = rel1

    #the technology itself: charge on activity a, discharge on the second one
    inv = ' '.join(str(v) for v in tp['capex'])
    fom = ' '.join(str(v) for v in tp['fom'])
    systems = (f"{name} {act_in}\n"
               f"    minp\t{lvl_in} 1.\n"
               f"    moutp\t{lvl_sto} c 1\n"
               f"    pll\tc {tp['lifetime']}\n"
               f"    inv\tts {inv}\n"
               f"    fom\tts {fom}\n"
               f"    hisc\t0.\thc {row['hisc']}\n"
               f"    con1c {idx}Sto:tin\tc 1\n"
               f"    consa {idx}StB\tc {row['efficiency']}\n"
               f"    abda lo\tc {row['lower bound']}\n"
               "#\n"
               f"2. activity {act_out}\n"
               f"    moutp\t{cfg['ef_pair'][row['discharge to']]} c 1\n"
               f"    consa {idx}StB\tc -1\n"
               "*\n")

    #the helper which lets the state of charge cross a time slice
    helper = f"{province}_StoTech{idx}_generic"
    systems += (f"{helper} {act_helper}\n"
                f"    moutp\t{lvl_sto} c 1\n"
                f"    inv\tc 1.0\n"
                f"    con1c {idx}Sto:tin\tc -1\n"
                f"    con1a {idx}Rel\tc -1\n"
                "#\n"
                "*\n")

    #one dummy per representative day, on for its first and last slice
    dummies = []
    for d in range(1, days + 1):
        dn = f"{province}Dum{idx}{d}"
        first, last = act_dummy[d - 1]
        systems += (f"{dn} {first}\n"
                    f"    moutp\t{lvl_sto} c 1\n"
                    f"    inv\tc 1.0\n"
                    f"    con1a {idx}Dm{d}\tc 1\n"
                    f"    consa {idx}StB\tc 1\n"
                    "#\n"
                    f"2. activity {last}\n"
                    f"    moutp\t{lvl_sto} c 1\n"
                    f"    con1a {idx}Dm{d}\tc -1\n"
                    f"    consa {idx}StB\tc -1\n"
                    "*\n")
        dummies.append((dn, d, first, last))
    parts['systems'] = systems

    #the charge takes from the grid and fills the store, the discharge gives
    #back to the grid and takes nothing
    lvl_out = cfg['ef_pair'][row['discharge to']]
    parts['codes'] = {name: matrix_code(lvl_sto, lvl_in, act_in),
                      f"{name}[2.]": matrix_code(lvl_out, None, act_out)}

    parts['ldb'] = (f"{name} {act_in}\n2. activity {act_out}\n*\n"
                    f"{helper} {act_helper}\n*\n")
    for dn, _, first, last in dummies:
        parts['ldb'] += f"{dn} {first}\n2. activity {last}\n*\n"

    #the .ldb has to declare every relation the .adb does, as the header line
    #and its closing *
    parts['ldb_relationss'] = f"StoBat{idx} {idx}StB o 0\n*\n"
    parts['ldb_relations1'] = (f"StoTech{idx} {idx}Sto o\n*\n"
                               f"RelBat{idx} {idx}Rel o 0\n*\n")
    for d in range(1, days + 1):
        parts['ldb_relations1'] += f"Dummy{idx}{d} {idx}Dm{d} o\n*\n"
    parts['ldb_relations2'] = ""

    #capacity factors: the helper carries every slice but the last of a day,
    #each dummy is on for the first slice of its day and its second activity
    #for the last slice of the same day
    year0, ldr, adb_lc = cfg['year0'], "", ""
    helper_days = [[round(1 / steps, 6)] * (steps - 1) + [0.0] for _ in range(days)]
    ldr += capfac_block(helper, act_helper, year0, days, helper_days)
    adb_lc += capfac_line(helper, act_helper, helper_days)
    for dn, d, act_a, act_b in dummies:
        on_first = [[1.0 if (s == 0 and day == d) else 0.0 for s in range(steps)]
                    for day in range(1, days + 1)]
        on_last = [[1.0 if (s == steps - 1 and day == d) else 0.0 for s in range(steps)]
                   for day in range(1, days + 1)]
        ldr += capfac_block(dn, act_a, year0, days, on_first)
        ldr += capfac_block(dn, act_b, year0, days, on_last)
        adb_lc += capfac_line(dn, act_a, on_first) + capfac_line(dn, act_b, on_last)
    parts['ldr'] = ldr
    parts['loadcurve'] = adb_lc

    return parts


def hydro_blocks(cfg, province, tech_name, idx, form_letter, letters):
    """every piece a pumped storage scheme adds"""
    row = cfg['HydroData'].loc[tech_name]
    tp = cfg['tech_param'][tech_name]
    act_turb, act_pump = next(letters), next(letters)
    name = f"{province}_{tech_name}_generic"
    form = f"{province}{row['form suffix']}"
    lvl_res = f"{form_letter}-{cfg['ef']['Transmission']}"
    lvl_grid = cfg['ef_pair'][row['turbine to']]

    parts = {'energyforms': f"    {form} {form_letter} l \n    #\n"}

    #the full unit list a storage relation carries, as the working model has
    #it. an earlier version of this stopped at ctime:yr, which left out the
    #units of transfac among others
    units = ("    units\ttype: energy, cost:US$'00/kWyr, inv:US$'00/kW, "
             "fom:US$'00/kW/yr, pll:yr, cmix:MW, hisccap:MW, ctime:yr, reten:yr, "
             "retenhist:MWyr, upper:MWyr, lower:MWyr, transfac:%\n")
    #the reservoirs draw from the river, referenced as p followed by that
    #relation's code, so the two must agree
    river_code = f"{idx}riv"
    parts['relationss'] = ""
    #only the upper reservoir spills, the lower one has nowhere to spill to
    for code, label, spills in ((f"{idx}upr", f"upres{idx}", True),
                                (f"{idx}lor", f"lores{idx}", False)):
        parts['relationss'] += (
            f"{label} {code} o 0\n"
            f"{units}"
            f"    fyear\t{cfg['year0']}\n"
            f"    for_ldr\tall\n"
            f"    upper\tc {row['reservoir upper']}\n"
            f"    initval\t{float(row['initial level'])}\n"
            f"    stortype\tcontinuous\n"
            f"    transfac\tc {float(row['transfac']):.4f}\n"
            f"    type\thydro\n"
            f"    inflow\tp{river_code}\tc {row['reservoir inflow']}\n")
        if spills:
            parts['relationss'] += "    overflow\tpnone\tc 1.\n"
        parts['relationss'] += "*\n"

    parts['relations2'] = (f"river{idx} {river_code} o 0\n"
                           f"    units\tgroup: activity, type: energy, cost:US$'00/kWyr, upper:MWyr, lower:MWyr\n"
                           f"    for_ldr\tall\n"
                           f"    lower\tc 0\n"
                           f"    type\triver\n"
                           f"    inflow\tpnatural\tc {row['river inflow']}\n"
                           f"*\n")
    parts['relations1'] = f"DummyHyd{idx} {idx}DHy o\n*\n"

    inv = ' '.join(str(v) for v in tp['capex'])
    fom = ' '.join(str(v) for v in tp['fom'])
    vom = ' '.join(str(v) for v in tp['vom'])
    parts['systems'] = (f"{name} {act_turb}\n"
                        f"    moutp\t{lvl_grid} c 1\n"
                        f"    pll\tc {tp['lifetime']}\n"
                        f"    inv\tts {inv}\n"
                        f"    fom\tts {fom}\n"
                        f"    vom\tts {vom}\n"
                        f"    hisc\t0.\thc {row['hisc']}\n"
                        f"    bdi up\tc {row['upper build']}\n"
                        f"    consa {idx}lor\tc 1\n"
                        f"    consa {idx}upr\tc -1\n"
                        "#\n"
                        f"2. activity {act_pump}\n"
                        f"    minp\t{cfg['ef_pair'][row['pump from']]} 1.\n"
                        f"    moutp\t{lvl_res} c 1\n"
                        f"    vom\tts {vom}\n"
                        f"    consa {idx}lor\tc -1\n"
                        f"    consa {idx}upr\tc 1\n"
                        "*\n")
    #the turbine gives to the grid and takes nothing, the pump takes from the
    #grid and fills the upper reservoir
    parts['codes'] = {name: matrix_code(lvl_grid, None, act_turb),
                      f"{name}[2.]": matrix_code(lvl_res,
                                                 cfg['ef_pair'][row['pump from']],
                                                 act_pump)}

    parts['ldb'] = f"{name} {act_turb}\n2. activity {act_pump}\n*\n"
    #the .ldb has to declare the reservoirs, the river and the hydro dummy too
    parts['ldb_relationss'] = (f"upres{idx} {idx}upr o 0\n*\n"
                               f"lores{idx} {idx}lor o 0\n*\n")
    parts['ldb_relations1'] = f"DummyHyd{idx} {idx}DHy o\n*\n"
    parts['ldb_relations2'] = f"river{idx} {river_code} o 0\n*\n"
    parts['ldr'] = ""
    #the river needs no load curve entries, the model resolves it without them
    parts['loadcurve'] = ""

    return parts


# --------------------------------------------------------------------------
# putting the blocks into the files MESSAGE_trans.py wrote
# --------------------------------------------------------------------------

def insert_at(lines, marker, text, before=True):
    """
    put text next to the line which starts with marker
    :param before: insert above that line, otherwise below it
    """
    for i, line in enumerate(lines):
        if line.lower().startswith(marker.lower()):
            at = i if before else i + 1
            return lines[:at] + text.splitlines(keepends=True) + lines[at:]
    raise SystemExit(f"{marker!r} not found, the file is not what was expected")


def close_of_level(lines, level):
    """the index of the * which closes an energyforms level"""
    for i, line in enumerate(lines):
        if line.split()[:1] == [level]:
            for j in range(i + 1, len(lines)):
                if lines[j].strip() == '*':
                    return j
    raise SystemExit(f"level {level!r} not found in energyforms")


def add_to_adb(path, parts_list):
    """splice every block into its section of the .adb"""
    lines = open(path, encoding='utf-8').read().splitlines(keepends=True)

    for parts in parts_list:
        at = close_of_level(lines, 'Transmission')
        lines = lines[:at] + parts['energyforms'].splitlines(keepends=True) + lines[at:]

    #the relation blocks end on a * with no newline, matching how
    #MESSAGE_trans.py builds them, so close them off before splicing or the
    #next section header ends up glued to the last *
    joined = {}
    for k in ('loadcurve', 'relationss', 'relations1', 'relations2', 'systems'):
        text = ''.join(p.get(k, '') for p in parts_list)
        joined[k] = text if not text or text.endswith('\n') else text + '\n'

    #loadcurve, relationss and relations2 are written as bare headers, the
    #content goes straight after them. relations1 already has the ramp
    #constraints, so its blocks go at the end, before relations2
    #the capfac lines go at the END of the load curve block. putting them
    #straight after the loadcurve: header would come before its year line,
    #which MESSAGE reports as "Keyword year in load curve not found"
    lines = insert_at(lines, 'relationsc:', joined['loadcurve'], before=True)
    lines = insert_at(lines, 'relationss:', joined['relationss'], before=False)
    lines = insert_at(lines, 'relations2:', joined['relations1'], before=True)
    lines = insert_at(lines, 'relations2:', joined['relations2'], before=False)
    lines = insert_at(lines, 'resources:', joined['systems'], before=True)

    open(path, 'w', encoding='utf-8').write(''.join(lines))


def add_to_ldb(path, parts_list):
    """the .ldb only carries the names, and the same load curves"""
    lines = open(path, encoding='utf-8').read().splitlines(keepends=True)
    for parts in parts_list:
        at = close_of_level(lines, 'Transmission')
        lines = lines[:at] + parts['energyforms'].splitlines(keepends=True) + lines[at:]
    lines = insert_at(lines, 'relationsc:',
                      ''.join(p['loadcurve'] for p in parts_list), before=True)
    #a relation declared in the .adb has to be declared here too, or the
    #matrix generator fails on the ones it cannot find
    lines = insert_at(lines, 'relationss:',
                      ''.join(p['ldb_relationss'] for p in parts_list), before=False)
    lines = insert_at(lines, 'relations2:',
                      ''.join(p['ldb_relations1'] for p in parts_list), before=True)
    lines = insert_at(lines, 'relations2:',
                      ''.join(p['ldb_relations2'] for p in parts_list), before=False)
    lines = insert_at(lines, 'resources:',
                      ''.join(p['ldb'] for p in parts_list), before=True)
    open(path, 'w', encoding='utf-8').write(''.join(lines))


def add_to_ldr(path, parts_list):
    """the .ldr just gains the capacity factor curves at the end"""
    with open(path, 'a', encoding='utf-8') as f:
        f.write(''.join(p['ldr'] for p in parts_list))


def region_letters(path):
    """
    the subregion letter of each case, from the regid file MESSAGE_trans.py
    writes. every technology of that case carries the letter in its code
    """
    letters = {}
    if not os.path.exists(path):
        return letters
    for line in open(path, encoding='utf-8'):
        parts = line.split()
        if len(parts) == 2 and len(parts[1]) == 1 and parts[1].isalpha():
            letters[parts[0]] = parts[1]

    return letters


def add_to_techcodes(path, codes):
    """
    the storage technologies, which MESSAGE_trans.py does not know about, added
    to the table of matrix names so that their profiles can be read out too

    :param codes: {technology: code including the subregion letter}
    """
    if not codes:
        return
    if not os.path.exists(path):
        print(f"warning: {path} not found, the storage profiles cannot be read")
        return

    table = pd.read_csv(path)
    table = table[~table['technology'].isin(codes)]     #so a rerun replaces
    table = pd.concat([table, pd.DataFrame(sorted(codes.items()),
                                           columns=['technology', 'code'])])
    table.sort_values('technology').to_csv(path, index=False)
    print(f"  {os.path.basename(path)}: added {len(codes)} storage activities")


def add_to_hist_tab(path, entries):
    """
    the capacity the storage technologies already have.

    MESSAGE_trans.py builds hist_tab from the TechCapacity sheet, which has no
    row for a storage: its existing capacity is the hisc of StorageData or
    HydroData instead, one cell holding the year and the capacity. without
    this the results script rebuilds their capacity from nothing and reports
    zero, while the model itself runs on the hisc in the .adb

    :param entries: {technology name: 'year capacity'}
    """
    if not entries:
        return
    if not os.path.exists(path):
        print(f"warning: {path} not found, the storage capacity is not recorded")
        return

    hist = pd.read_csv(path, index_col=0)
    for name, hisc in entries.items():
        year, capacity = str(hisc).split()
        hist.loc[int(year), name] = float(capacity)
    hist.sort_index().to_csv(path)
    print(f"  {os.path.basename(path)}: recorded the capacity of "
          f"{len(entries)} storage technologies")


def add_to_cin(path, storage_forms):
    """
    let the parent's report definitions cover the storage energy forms too,
    or the res file has no columns for them.

    :param storage_forms: {case name: [(form, type), ...]} for the cases which
                          got storage
    """
    if not os.path.exists(path):
        print(f"warning: {path} not found, the storage forms are not reported")
        return

    #which quantity each per subregion block asks for. the cost blocks only
    #cover the battery forms, the reservoir of a pumped scheme carries no cost
    #of its own
    verbs = {'prodall': ('production for', True),
             'capall': ('total capacity for', True),
             'capadd': ('new capacity for', True),
             'investmentcost': ('inv for', False),
             'fixedcost': ('fom for', False),
             'variablecost': ('vom for', False)}
    lines = open(path, encoding='utf-8').read().splitlines(keepends=True)
    present = set(lines)     #so that running this twice does not double them
    out, title, case_name = [], None, None
    for line in lines:
        out.append(line)
        if line.startswith('title: '):
            title = line[len('title: '):].strip()
        elif line.startswith('!aggr '):
            case_name = line.split(None, 1)[1].strip()
        #the Distribution line is the last table of an !aggr block, so the
        #storage forms go straight after it
        elif (line.startswith('!table:') and 'on level Distribution' in line
                and title in verbs and case_name in storage_forms):
            verb, with_hydro = verbs[title]
            for form, tech_type in storage_forms[case_name]:
                if tech_type == PUM_TYPE and not with_hydro:
                    continue
                new = f"!table: {verb} energyform {form} on level Transmission\n"
                if new not in present:
                    out.append(new)

    open(path, 'w', encoding='utf-8').write(''.join(out))
    print(f"  {os.path.basename(path)}: reported the storage forms of "
          f"{len(storage_forms)} cases")


# --------------------------------------------------------------------------
# reading what the sheets say
# --------------------------------------------------------------------------

def load_config(input_fd, input_fn):
    """
    the parts of the workbook this script needs, read with the same helpers
    MESSAGE_trans.py uses so that names normalise identically
    """
    wb = mt.read_workbook(f"{input_fd}{input_fn}.xlsx")
    general = wb["General"].set_index('Parameter')['Value'].to_dict()
    years = list(wb["Years"]['years'])

    tech_param = mt.strstrip(wb["TechData"], ['Technology name']).set_index(
        'Technology name').T.to_dict()
    for sheet, key in (("TechCapex", 'capex'), ("fom", 'fom'), ("vom", 'vom')):
        df = mt.strstrip(wb[sheet], ['Technology name'])
        for name in tech_param:
            rows = df[df['Technology name'] == name]
            tech_param[name][key] = ([0] * len(years) if len(rows) == 0
                                     else rows[years].fillna(0).to_dict('tight')['data'][0])

    generics = wb["generics"].set_index('region').fillna(0)
    generics_dict = {p: {t: v for t, v in row.items() if v != 0}
                     for p, row in generics.to_dict('index').items()}

    #<form>-<level> for the levels a storage can connect to
    ef_pair = {}
    for level, forms in mt.energyforms.items():
        for form_name, _flag in forms:
            ef_pair[form_name] = f"{mt.ef[form_name]}-{mt.ef[level]}"

    return {
        'years': years,
        'year0': int(years[0]),
        'days_year': int(general["Days per year"]),
        'timesteps_day': int(general["Timesteps per day"]),
        'tech_param': tech_param,
        'generics_dict': generics_dict,
        'ef': mt.ef,
        'ef_pair': ef_pair,
        #a scenario with no storage need not carry the sheets at all
        'StorageData': storage_sheet(wb, "StorageData"),
        'HydroData': storage_sheet(wb, "HydroData"),
    }


def storage_sheet(wb, name):
    """the sheet indexed by storage name, or empty if the workbook has none"""
    if name not in wb:
        return pd.DataFrame(columns=['Storage name']).set_index('Storage name')

    return mt.strstrip(wb[name], ['Storage name']).set_index('Storage name')


# --------------------------------------------------------------------------
# driver
# --------------------------------------------------------------------------

def add_storage_to_scenario(input_fn, input_fd, main_name):
    """add the storage technologies to every case of one scenario"""
    cfg = load_config(input_fd, input_fn)

    if not any(p['type'] in (BAT_TYPE, PUM_TYPE) for p in cfg['tech_param'].values()):
        print(f"  {input_fn} has no {BAT_TYPE} or {PUM_TYPE} technologies, nothing to add")
        return

    province_list = []
    with open(f"{input_fd}runprovince.csv", newline='') as f:
        import csv
        for row in csv.reader(f):
            province_list.append(row[0])

    storage_forms = {}   #case name -> the energy forms its storage introduced
    hisc = {}            #technology name -> the capacity it starts with
    codes = {}           #technology name -> what it is called in the matrix
    main_fd = f"{mt.MESSAGE_root_fd}/{main_name}/{main_name}"
    letters = region_letters(f"{main_fd}/regid")
    for province_long in province_list:
        storages = storage_of_province(cfg, province_long)
        if not storages:
            continue
        province = province_long[0:6].replace(" ", "")
        #the same name MESSAGE_trans gives the subregion: the scenario, not
        #the workbook, so two scenarios from one workbook stay apart
        case_name = f"{province_long}_{main_name}".replace(" ", "")
        #the storage goes into the staging copy, which MESSAGE_trans.py then
        #copies to the MESSAGE tree once every case is complete
        case_dir = f"{mt.output_base_fd}{case_name}/data"
        if not os.path.isdir(case_dir):
            print(f"warning: {case_dir} not found, skipping {province_long}")
            continue

        #what this case's storage adds to the parent's own files. taken from
        #the sheets rather than from what was written, so that the two below
        #are still right for a case which is skipped as already done
        for tech_name, tech_type in storages:
            sheet = cfg['StorageData'] if tech_type == BAT_TYPE else cfg['HydroData']
            storage_forms.setdefault(case_name, []).append(
                (f"{province}{sheet.loc[tech_name, 'form suffix']}", tech_type))
            hisc[f"{province}_{tech_name}_generic"] = sheet.loc[tech_name, 'hisc']

        adb = f"{case_dir}/{case_name}.adb"
        #appending twice would double every block, so only touch a case which
        #MESSAGE_trans.py has freshly written
        already = [n for n, _ in storages
                   if f"{province}_{n}_generic" in open(adb, encoding='utf-8').read()]
        if already:
            print(f"  {case_name}: {', '.join(already)} already present, skipping")
            continue

        used_letters = []
        parts_list = []
        #one run of activity letters across every storage of this case, so that
        #no two of them end up with the same technology id
        activity_letters = iter(mt.ascii_all)
        for idx, (tech_name, tech_type) in enumerate(storages, start=1):
            letter = free_form_letter(open(adb, encoding='utf-8').read().splitlines(),
                                      used_letters)
            used_letters.append(letter)
            build = battery_blocks if tech_type == BAT_TYPE else hydro_blocks
            parts_list.append(build(cfg, province, tech_name, idx, letter,
                                    activity_letters))

        #the subregion letter finishes the matrix code of each activity. it is
        #only known once the case is built, which is why it is added here
        if case_name in letters:
            for tech, code in ((t, c) for p in parts_list
                               for t, c in p['codes'].items()):
                codes[tech] = f"{code}{letters[case_name]}"
        else:
            print(f"  {case_name}: not in regid, its profiles cannot be read")

        add_to_adb(adb, parts_list)
        add_to_ldb(f"{case_dir}/{case_name}_adb.ldb", parts_list)
        add_to_ldr(f"{case_dir}/{case_name}_adb.ldr", parts_list)
        print(f"  {case_name}: added {', '.join(n for n, _ in storages)}")

    #the parent reports every case, so its .cin has to know the new forms
    add_to_cin(f"{mt.output_base_fd}{main_name}/data/{main_name}.cin", storage_forms)
    #these two are written straight to the MESSAGE tree, not to the staging copy
    add_to_hist_tab(f"{main_fd}/hist_tab.csv", hisc)
    add_to_techcodes(f"{main_fd}/techcodes.csv", codes)


if __name__ == '__main__':
    for _, scenario in mt.read_create_cases(mt.create_cases_fp).iterrows():
        if scenario['add_storage'] != 1:
            print(f"skipping {scenario['input_fn']}, add_storage is off")
            continue
        main_name = f"{scenario['main']}_{scenario['input_fn']}"
        print(f"=== {scenario['input_fn']} -> {main_name} ===")
        add_storage_to_scenario(scenario['input_fn'], scenario['input_fd'], main_name)
