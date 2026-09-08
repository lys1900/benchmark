#YL, IAEA, 2025 June
import collections
import math

import pandas as pd
import openpyxl
import os
import shutil
from string import ascii_lowercase, ascii_uppercase
import re
import csv
import datetime
import numpy as np

import add_storage
import settings

#the folders this runs in. settings.py holds the defaults and the workbench
#writes settings.json over them, so a path only has to be set in one place
paths = settings.load()

orig_name = 'CountryA'
orig_base_fd = paths['orig_base_fd']  #reference case copied to seed every case
grf_fd = "/tmp/grf/iaea"  #graphics folder written into the .dir files
#one row per scenario, with the columns input_fn, input_fd, make and create_reg
create_cases_fp = paths['create_cases_fp']
MESSAGE_fd = paths['MESSAGE_fd']
generate_main = True  #if running multiregional model
output_base_fd = paths['output_base_fd']
#MESSAGE_root_fd = "E:/Work/China_MESSAGE/temp/"
MESSAGE_root_fd = f"{MESSAGE_fd}models/"
MESSAGE_bat_all_path = f"{MESSAGE_root_fd}/run_all_adb.bat"
MESSAGE_mms_fils = f"{MESSAGE_root_fd}/mms_fils/"
#dummy seasons only to represent days
seasons = ["Season1", "Season2", "Season3", "Season4"]
#the (month, day) a season starts on. the days in each season and the length
#of a time slice are both derived from these, so that the .ldr and the .adb
#cannot disagree about how long a season is
season_bounds = [(3, 25), (6, 25), (9, 25), (12, 25)]
doyinquarter = {i: [(i * 90) + 85] for i in range(0, 4)}
hours_year = [i for i in range(1, 8761)]  #hour numbers used to select the profile rows
solvers = ['HiGHS', 'cplex']
rampdir = {'u':1, 'd':-1}
#technology types which add_storage.py writes out, not this script. they are
#left out of the model altogether when a scenario has add_storage off
storage_types = ['BAT', 'PUM']
#Energy forms of the model: each level, and the forms carried on it as
#(name, flag) pairs. The .adb gives every level and form a letter in the order
#they appear here, and refers to a form as <form>-<level>, so g-f below is
#ElectricityNonVRE on the Transmission level. Adding a form here re-letters
#everything after it and shifts the fuel letters with it. The Fuel level is
#filled per province with the fuels that are actually used.
energyforms = {
    'Final':        [('ElectricityDemand', 'l'), ('HeatDemand', 'l')],
    'Distribution': [('ElectricityDistribution', 'l')],
    'Transmission': [('ElectricityNonVRE', 'l'), ('ElectricityVRE', 'l'), ('Heat', '')],
    'Fuel':         [],
}
ef = {}  #energy form or level -> its letter in the .adb
for ef_level, ef_forms in energyforms.items():
    ef[ef_level] = ascii_lowercase[len(ef)]
    for ef_name, ef_flag in ef_forms:
        ef[ef_name] = ascii_lowercase[len(ef)]
fuel_letter0 = len(ef) + 1  #fuels are lettered from here, the +1 leaves 'k' free as before

#shorthand for the <form>-<level> pairs the .adb text refers to
lvl_elecdemand = f"{ef['ElectricityDemand']}-{ef['Final']}"                #b-a
lvl_heatdemand = f"{ef['HeatDemand']}-{ef['Final']}"                       #c-a
lvl_elecdist = f"{ef['ElectricityDistribution']}-{ef['Distribution']}"     #e-d
lvl_nonvre = f"{ef['ElectricityNonVRE']}-{ef['Transmission']}"             #g-f
lvl_vre = f"{ef['ElectricityVRE']}-{ef['Transmission']}"                   #h-f


def strstrip(df, cols):
    """
    function to strip columns and change to lower case
    """
    df[cols] = df[cols].replace(r'[^A-Za-z0-9]', '', regex=True)
    for c in cols:
        df[c] = df[c].str.lower()
    return df


def numstr(v):
    """
    write a number without a trailing .0 when it is whole, so an efficiency of
    1 stays "1" and a lifetime of 40 stays "40"
    """
    v = float(v)
    return str(int(v)) if v.is_integer() else str(v)


def read_create_cases(path):
    """
    read the list of scenarios, keeping only those flagged to be made. a blank
    make reads as NaN and is dropped with the rest
    """
    scenarios = pd.read_csv(path)

    return scenarios[scenarios['make'] == 1]


def registered_cases(MESSAGE_mms_fils):
    """
    which cases are already written into mms.pro. create_reg_func appends, so
    a case must only be registered the first time it is built
    """
    mms_pro = f"{MESSAGE_mms_fils}mms.pro"
    if not os.path.exists(mms_pro):
        return set()
    with open(mms_pro) as f:
        return {line.split('.dirfile')[0].strip() for line in f if '.dirfile' in line}


def read_workbook(path):
    """
    read every sheet of an excel file into a dict of dataframes
    """
    xl = pd.ExcelFile(path)

    return {sheetname: xl.parse(sheetname) for sheetname in xl.sheet_names}


def ldb_stub(name, activity):
    """
    the .ldb only needs the name and activity letter of each technology
    """
    return (f"{name} {activity}\n"
            "*\n")


def check_tech_keys(input_df_all, years):
    """
    check the sheets which are keyed by technology name against TechMap, and
    collect everything at once rather than stopping at the first problem
    :return: (problems, warnings). a name in a sheet which TechMap does not
             have is a problem and ends the script, a name in TechMap which a
             sheet does not have is only a warning
    """
    problems = []
    warnings = []

    names = list(strstrip(input_df_all["TechMap"], ['Technology name'])['Technology name'])
    known = set(names)
    dup_map = sorted({n for n in names if names.count(n) > 1})
    if dup_map:
        problems.append(f"TechMap repeats the technology names {dup_map}")

    #tm keeps one technology name per technology and technology type, so where
    #several share a pair only the last of them gets the TechCapacity rows
    pairs = strstrip(input_df_all["TechMap"], ['Technology', 'Technology Type'])
    shared = pairs.groupby(['Technology', 'Technology Type'], sort=False)['Technology name'].apply(list)
    for (tech, ttype), maps_to in shared.items():
        if len(maps_to) > 1:
            warnings.append(f"TechMap maps ('{tech}', '{ttype}') to {maps_to}, "
                            f"TechCapacity rows for it will use '{maps_to[-1]}'")

    for sheet in ['TechData', 'TechCapex', 'fom', 'vom', 'TechConstraints']:
        df = input_df_all[sheet]
        sheet_names = list(strstrip(df, ['Technology name'])['Technology name'])

        #the same name twice makes the lookup ambiguous, one row silently wins
        for d in sorted({n for n in sheet_names if sheet_names.count(n) > 1}):
            rows = [i + 2 for i, n in enumerate(sheet_names) if n == d]  # +2 gives the excel row
            problems.append(f"{sheet} repeats '{d}' on excel rows {rows}")

        if set(sheet_names) - known:
            problems.append(f"{sheet} has names which are not in TechMap: "
                            f"{sorted(set(sheet_names) - known)}")
        if known - set(sheet_names):
            warnings.append(f"{sheet} has no row for {sorted(known - set(sheet_names))}, "
                            f"so they are not used")

        if sheet not in ['TechData', 'TechConstraints']:
            no_column = [y for y in years if y not in df.columns]
            if no_column:
                problems.append(f"{sheet} has no column for the model years {no_column}")
            for _, row in df.iterrows():
                blank = [y for y in years if y in df.columns and pd.isna(row[y])]
                if blank:
                    problems.append(f"{sheet} '{row['Technology name']}' is blank in {blank}")

    generics_cols = [re.sub(r'[^A-Za-z0-9]', '', str(c)).lower()
                     for c in input_df_all["generics"].columns if c != 'region']
    if set(generics_cols) - known:
        problems.append(f"generics has columns which are not in TechMap: "
                        f"{sorted(set(generics_cols) - known)}")
    if known - set(generics_cols):
        warnings.append(f"generics has no column for {sorted(known - set(generics_cols))}, "
                        f"so they are not used")

    return problems, warnings


def check_tech_capacity(tech_p, tm, province_long):
    """
    check the province TechCapacity rows against TechMap. rows whose
    technology and technology type have no TechMap entry are dropped, so that
    the run continues rather than failing on the tech_name lookup
    :return: (the rows which do map, warnings about the ones which do not)
    """
    warnings = []
    maps = [t in tm and ty in tm[t]
            for t, ty in zip(tech_p['Technology'], tech_p['Technology Type'])]
    unmatched = tech_p[[not m for m in maps]]
    for t, ty in sorted({(t, ty) for t, ty in
                         zip(unmatched['Technology'], unmatched['Technology Type'])}):
        n = len(unmatched[(unmatched['Technology'] == t) & (unmatched['Technology Type'] == ty)])
        warnings.append(f"{province_long}: TechCapacity has ('{t}', '{ty}') with no TechMap "
                        f"entry, {n} row(s) dropped")

    return tech_p[maps], warnings


def cin_block(title, body):
    """one report definition in the .cin, a title and its table lines"""
    return (f"title: {title}\n"
            "unit: , 1.\n"
            "@\n"
            f"{body}"
            "@\n")


def cin_string(case_names_prov):
    """
    the .cin, which says which tables cap writes into the res file.

    the first blocks report the parent case itself, the ones after them ask
    for the same thing per subregion, which is where the per technology
    columns of prodall, capall and capadd come from. without those the res
    file only has the interconnectors.
    """
    #the parent's own tables
    s = cin_block('Objective', "obj        = func:act\n")
    s += cin_block('emissions', "emit       =          nCO2L:act\n")
    for title, verb in (('production', 'production of'),
                        ('total capacity', 'total capacity for'),
                        ('new capacity', 'new capacity for')):
        s += cin_block(title, f"!table: {verb} energyform ElectricityDistribution "
                              f"on level Distribution\n")
    s += cin_block('shadow price', "!table: shp of energyforms on level Final\n")
    for title, key in (('investment cost', 'inv'), ('variable cost', 'vom'),
                       ('fixed cost', 'fom')):
        body = ""
        for form, level in (('ElectricityNonVRE', 'Transmission'),
                            ('ElectricityVRE', 'Transmission'),
                            ('ElectricityDistribution', 'Distribution'),
                            ('ElectricityDemand', 'Final')):
            body += f"!table: {key} for technologies producing energyform {form} on level {level}\n"
        s += cin_block(title, body)

    #demands are named by the letters of the form and its level
    s += ("title: Demands\n"
          "@\n"
          f"ElectricityDemand =                {ef['Final']}{ef['ElectricityDemand']}:act\n"
          f"HeatDemand =                {ef['Final']}{ef['HeatDemand']}:act\n"
          "@\n")

    #the same quantities again, but per subregion
    for title, verb in (('prodall', 'production of'),
                        ('capall', 'total capacity for'),
                        ('capadd', 'new capacity for'),
                        ('investmentcost', 'inv for'),
                        ('fixedcost', 'fom for'),
                        ('variablecost', 'vom for')):
        body = ""
        for case_name in case_names_prov:
            body += f"!aggr {case_name}\n"
            for form, level in (('ElectricityVRE', 'Transmission'),
                                ('ElectricityNonVRE', 'Transmission'),
                                ('ElectricityDistribution', 'Distribution')):
                body += f"!table: {verb} energyform {form} on level {level}\n"
        s += cin_block(title, body)

    return s


def set_ntrun(gen_path, ntrun):
    """
    the last period the matrix is built for, in the .gen copied from the
    reference case. the reference says 5, which is only right for a five
    period model, so it is rewritten to suit the horizon of this scenario
    """
    lines = open(gen_path, encoding='utf-8').read().splitlines(keepends=True)
    for i, line in enumerate(lines):
        if line.split(':')[0].strip() == 'ntrun':
            comment = line.partition(';')[2]
            lines[i] = f"ntrun:\t{ntrun}\t;{comment}" if comment else f"ntrun:\t{ntrun}\n"
            break
    else:
        raise SystemExit(f"no ntrun line in {gen_path}, the reference case is not "
                         f"what was expected")
    open(gen_path, 'w', encoding='utf-8').write(''.join(lines))


def read_constraint_bounds(constraints_properties, constraints_names):
    """
    which side of each constraint the values from the Constraints sheet go on.

    the boundtype column of ConstraintsTypes says "upper" or "lower". a
    constraint which says neither is taken as an upper one, which is how they
    all behaved before the column existed, and is reported so that a blank
    cell or a typo is not mistaken for a deliberate choice

    :return: {constraint: 'upper' or 'lower'}
    """
    bounds, assumed = {}, []
    for con in constraints_names:
        given = str(constraints_properties.get(con, {}).get('boundtype', '')).strip().lower()
        if given in ('upper', 'lower'):
            bounds[con] = given
        else:
            bounds[con] = 'upper'
            assumed.append(con if not given or given == 'nan' else f"{con} (says {given!r})")
    if assumed:
        print(f"warning: no boundtype for {', '.join(assumed)}, "
              f"assuming upper")

    return bounds


def energyforms_block(levels):
    """
    write the energyforms section of the .adb from the energyforms structure
    :param levels: the levels to write, as {level: [(form, flag), ...]}
    """
    s = "energyforms: \n"
    for level, forms in levels.items():
        s += (f"{level} {ef[level]}\n"
              "#\n")
        for name, flag in forms:
            if flag:
                s += f"    {name} {ef[name]} {flag} \n"
            else:
                s += f"    {name} {ef[name]}\n"
            s += "    #\n"
        s += "*\n"

    return s


def custom_reader(input, sheetname='Demand'):
    """
    Custom reader to read some excel sheets into dict
    :param sheetname: the sheet name to read
    :return: a dictionary in dictionary, with keys country and parameter
    """
    cols = []
    for i in input[sheetname].columns:
        cols.append(i.split('.', 1)[0])

    new_columns = list(zip(input[sheetname].loc[0], cols))
    input[sheetname].columns = pd.MultiIndex.from_tuples(new_columns)
    input_new = input[sheetname][1:]

    dict_out = collections.defaultdict(dict)
    for column in input_new:
        dict_out[column[1]][column[0]] = list(input_new[column])

    return dict_out


def custom_reader_2(input, sheetname='Demand', idlist=[1]):
    """
    Custom reader to read some excel sheets into dict
    :param sheetname: the sheet name to read
    :return: a dictionary of dictionary, with keys parameter
    """
    df_out = input[sheetname].set_index(input[sheetname].columns[0]).dropna(axis=1, how='all')
    dict_out = df_out[df_out.index.isin(idlist)].to_dict('list')

    return dict_out


def shift_profile(profi, gmt):
    """
    Shift profile by gmt hours to local time zone
    """
    profi_shifted = profi[-gmt:] + profi[:-gmt]

    return profi_shifted


def sample_days(profi, hoyinday):
    """
    take the hours of each representative day out of an 8760 hour profile
    :return: a dictionary of day of year, with the hours of that day
    """
    return {day: [profi[i] for i in hours] for day, hours in hoyinday.items()}


def create_reg_func(MESSAGE_mms_fils, case_name, hasmain, mn=""):
    if hasmain:
        case_path = f'{mn}/{case_name}'
    else:
        case_path = case_name

    dir_s = (f"#call              answer\n"
             f"supply             $MMS_HOME/{case_path}      \n"
             f"grfdir             {grf_fd}           \n"
             f"cin                $MMS_HOME/{case_path}/data \n"
             f"tdb                $MMS_HOME/tdb           \n"
             f"adb                $MMS_HOME/{case_path}/data \n"
             f"ldb                $MMS_HOME/{case_path}/data \n"
             f"upd                $MMS_HOME/{case_path}/data \n"
             f"gen                $MMS_HOME/{case_path}/data \n"
             f"data               $MMS_HOME/{case_path}/data \n"
             )
    if hasmain and case_name == mn:
        dir_s += f"regid	$MMS_HOME/{mn}/{mn}/regid \n"
    #the registry folder may not exist yet on a fresh MESSAGE tree
    os.makedirs(MESSAGE_mms_fils, exist_ok=True)
    # Add line to mms.pro
    with open(f"{MESSAGE_mms_fils}mms.pro", "a") as f1:
        f1.write(f"{case_name}.dirfile   $MMS_HOME/mms_fils/{case_name}.dir\n")
    with open(f"{MESSAGE_mms_fils}glob.reg", "a") as f2:
        f2.write(f"{case_name}	{case_name}	{case_name}	+	empty\n")
    dir_f = f"{MESSAGE_mms_fils}{case_name}.dir"
    with open(dir_f, "w") as f3:
        f3.write(dir_s)


ascii_all = ascii_lowercase + ascii_uppercase

def generate_scenario(input_fn, input_fd, main_name, with_storage, bat_all_s,
                      ntrun=None):
    """
    build every case of one scenario: the provinces named in
    runprovince.csv and the multiregional parent which joins them
    :param input_fn: scenario name, also the name of the national workbook
    :param input_fd: folder holding the national and provincial workbooks
    :param main_name: name of the multiregional parent case
    :param with_storage: whether add_storage writes the storage technologies
                         at the end, they are left out entirely when it is off
    :param bat_all_s: run_all_adb.bat so far, scenarios add to it
    :return: run_all_adb.bat with this scenario's cases appended
    """
    province_fn = f"{input_fd}runprovince.csv"
    MESSAGE_main_fd = f"{MESSAGE_root_fd}/{main_name}"
    main_case_fd = f"{MESSAGE_root_fd}/{main_name}/{main_name}"
    output_main_fd = f"{output_base_fd}/{main_name}/"
    #cases already in mms.pro must not be registered a second time,
    #create_reg_func appends
    already_registered = registered_cases(MESSAGE_mms_fils)

    province_list = []
    with open(province_fn, newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            province_list.append(row[0])

    hoyinday = {}
    for i, j in doyinquarter.items():
        for k in j:
            hoyinday[k] = list(range((k - 1) * 24, k * 24))

    # Nationwide
    input_fp = f'{input_fd}{input_fn}.xlsx'
    input_df_all = read_workbook(input_fp)

    # Read all parameters
    #General
    general = input_df_all["General"].set_index('Parameter')['Value'].to_dict()
    drate = general["Discount rate"]
    days_year = int(general["Days per year"])
    timesteps_day = int(general["Timesteps per day"])
    years = list(input_df_all["Years"]['years'])
    #the last period the matrix is built for. the Years sheet carries one year
    #past the horizon, so the periods to solve are one fewer than its rows.
    #create_cases.csv can override it per scenario
    if not ntrun:
        ntrun = len(years) - 1
        print(f"  ntrun {ntrun}, from {len(years)} years in the sheet")
    else:
        ntrun = int(ntrun)
        print(f"  ntrun {ntrun}, as set in create_cases.csv")
    year0 = int(years[0])  #first milestone year, taken from the Years sheet
    baseyear = year0 - 1
    yearx = years[-1]
    years_continuous = [i for i in range(year0, yearx + 1)]
    years_intervalsafterprev = {years[i]: years[i] - years[i - 1] for i in range(1, len(years))}
    years_intervalsafterprev[year0] = 1
    years_inyearscontinuous = {years[i]: [j for j in years_continuous if j <= years[i] and j > years[i - 1]] for i in
                               range(1, len(years))}
    years_inyearscontinuous[year0] = [year0]
    years_incl_by = [baseyear] + years

    #Seasons. the boundary dates give the days in each season, one season per
    #representative day
    season_dates = [datetime.date(year0, m, d) for m, d in season_bounds]
    season_dates.append(datetime.date(year0 + 1, *season_bounds[0]))
    daysinseason = [(season_dates[i + 1] - season_dates[i]).days
                    for i in range(len(season_bounds))]
    days_in_year = sum(daysinseason)
    if len(daysinseason) != days_year:
        raise SystemExit(f"there are {len(daysinseason)} seasons but the General sheet "
                         f"says {days_year} days per year, they have to match")

    #Time slices. a slice is as long as its share of its own season, so the
    #lengths differ between seasons of different length
    ts = []
    lengths = []

    ts_count = days_year * timesteps_day

    for season, i in enumerate(ascii_lowercase[0:days_year]):
        ts_length = round(daysinseason[season] / (days_in_year * timesteps_day), 6)
        for j in ascii_lowercase[0:timesteps_day]:
            ts.append(f"{i}a{j}")
            lengths.append(ts_length)

    #Emissions
    emissions = input_df_all["Emissions"].set_index('years')
    emissions = emissions[emissions.index.isin(years)]

    #constraints
    constraints = input_df_all["Constraints"].set_index('years').fillna(0)
    constraints = constraints[constraints.index.isin(years)]
    constraints_names = list(constraints.columns)
    constraints_types = input_df_all["ConstraintsTypes"].groupby('type')['constraint'].apply(list).to_dict()
    constraints_properties = input_df_all["ConstraintsTypes"].set_index('constraint').to_dict('index')
    constraint_bounds = read_constraint_bounds(constraints_properties, constraints_names)
    #TechMap
    tech_map = input_df_all["TechMap"]
    tech_map = strstrip(tech_map, ['Technology name', 'Technology', 'Technology Type'])
    tm = {}
    for k, g in tech_map.groupby('Technology'):
        tm[k] = dict(zip(g['Technology Type'], g['Technology name']))

    #FuelPrice
    #fuel_param = input_df_all["FuelPrice"].set_index('Fuel').T
    #fuel = list(fuel_param.index)
    #fuel_price = fuel_param.to_dict()

    #TechnologyData, TechnologyCapex, fom, vom

    #check the technology names before anything is read with them
    key_problems, key_warnings = check_tech_keys(input_df_all, years)
    for w in key_warnings:
        print(f"warning: {w}")
    if key_problems:
        for p in key_problems:
            print(f"problem: {p}")
        raise SystemExit(f"{len(key_problems)} problem(s) in {input_fp}, fix these before running")

    tech_param = strstrip(input_df_all["TechData"], ['Technology name']).set_index(
        'Technology name').T.to_dict()
    tech_capex = strstrip(input_df_all["TechCapex"], ['Technology name'])
    tech_fom = strstrip(input_df_all["fom"], ['Technology name'])
    tech_vom = strstrip(input_df_all["vom"], ['Technology name'])
    tech_constraints = strstrip(input_df_all["TechConstraints"], ['Technology name'])

    #cost sheets which all have one row per technology name and one column per year
    tech_cost_sheets = {'capex': tech_capex, 'fom': tech_fom, 'vom': tech_vom}

    for key in tech_param:
        for param, cost_df in tech_cost_sheets.items():
            cost_rows = cost_df[cost_df['Technology name'] == key]
            if len(cost_rows) == 0:  #warned about by check_tech_keys
                tech_param[key][param] = [0] * len(years)
            else:
                tech_param[key][param] = cost_rows[years].to_dict('tight')['data'][0]

        con_rows = tech_constraints[tech_constraints['Technology name'] == key]
        if len(con_rows) == 0:  #warned about by check_tech_keys
            tech_param[key]['constraints'] = {con: 0 for con in constraints_names}
        else:
            tech_param[key]['constraints'] = con_rows.fillna(0).to_dict('records')[0]

    #TandDData
    td_param = input_df_all["TandDData"]
    #a blank cost means no cost, minp and moutp are left alone so that a blank
    #there can leave the line out altogether
    td_param[['inv', 'fom', 'vom']] = td_param[['inv', 'fom', 'vom']].fillna(0)
    #InterconnectionData
    ic_param = input_df_all["InterconnectionData"].set_index('Parameter')['Value'].to_dict()
    interconnection = input_df_all["Interconnection"]
    interconnection_long = pd.melt(interconnection, id_vars=interconnection.columns[0],
                                   value_vars=interconnection.columns[1:])
    interconnection_long = interconnection_long.rename(columns={'from-to': 'from', 'variable': 'to'})
    interconnection_long = interconnection_long.dropna(subset=['value'])
    interconnection_long = interconnection_long[interconnection_long['value'] != 0]
    interconnection_long['line_name'] = interconnection_long['from'] + '_' + interconnection_long['to']
    interconnection_main = interconnection_long[
        interconnection_long['from'].isin(province_list) & interconnection_long['to'].isin(province_list)]
    #FuelPrice
    fuel_y = custom_reader_2(input_df_all, 'FuelPrice', years)
    #GenericsTech
    generics = input_df_all["generics"].set_index('region').fillna(0)
    generics = generics[generics.index.isin(province_list)]
    generics_dict = generics.to_dict('index')
    generics_dict_ = dict(generics_dict.items())

    # Create a copy of the dictionary to iterate over
    generics_dict_copy = {p_: dict(pt) for p_, pt in generics_dict_.items()}
    for p_, pt in generics_dict_copy.items():
        for t_ in list(pt.keys()):  # Use list to avoid changing size during iteration
            if pt[t_] == 0:
                del generics_dict[p_][t_]

    if not os.path.exists(output_base_fd):
        os.makedirs(output_base_fd)

    #subregions carry the scenario name, not the workbook name. MESSAGE looks
    #a subregion up in mms_fils/mms.pro, which is shared by every scenario, so
    #naming them after the workbook made two scenarios built from the same
    #workbook read each other's subregions
    case_names_prov = [f'{p}_{main_name}'.replace(" ", "") for p in province_list]
    if generate_main:
        output_main_fd = f"{output_base_fd}/{main_name}/"
        if os.path.exists(output_main_fd) and os.path.isdir(output_main_fd):
            shutil.rmtree(output_main_fd)
        os.makedirs(output_main_fd)

        MESSAGE_main_fd = f"{MESSAGE_root_fd}/{main_name}"
        os.makedirs(MESSAGE_main_fd, exist_ok=True)  #so a scenario can be rebuilt

        cases_all = [main_name] + province_list
        case_names_all = [main_name] + case_names_prov
        province_dict = {main_name: '.'}  #dictionary of province, and their ids, initialise with main_name
        province_dict_1 = {}
        regid_string = f"{main_name} . \n"


    else:
        cases_all = province_list
        case_names_all = case_names_prov

    regid_counter = 0
    case_folders = []  #(staging folder, MESSAGE folder) for each case built
    hist_tab = {} #historic capacity table
    lt_tab = {} #life time table
    #what each technology is called in the matrix, so that the results script
    #can find it in the solution file. see where it is filled in below
    tech_codes = {}

    #energy forms shared by every case, province cases add the Fuel level to this.
    #built from the energyforms structure at the top, which writes out as:
    #                      "energyforms: \n"
    #                      "Final a\n"
    #                      "#\n"
    #                      "    ElectricityDemand b l \n"
    #                      "    #\n"
    #                      "    HeatDemand c l \n"
    #                      "    #\n"
    #                      "*\n"
    #                      "Distribution d\n"
    #                      "#\n"
    #                      "    ElectricityDistribution e l \n"
    #                      "    #\n"
    #                      "*\n"
    #                      "Transmission f\n"
    #                      "#\n"
    #                      "    ElectricityNonVRE g l \n"
    #                      "    #\n"
    #                      "    ElectricityVRE h l \n"
    #                      "    #\n"
    #                      "    Heat i\n"
    #                      "    #\n"
    #                      "*\n"
    energyforms_base_s = energyforms_block({level: forms for level, forms in energyforms.items()
                                            if level != 'Fuel'})


    for cid, c in enumerate(cases_all):

        if c in province_list:
            print(c)
            province_long = c
            #first 6 characters only, so that the technology names built from it,
            #<province>_<tech_name>_<status>, stay within the MESSAGE name length
            province = province_long[0:6].replace(" ", "")
            case_name = case_names_all[cid]  #todo:strip spaces
            if generate_main:
                province_dict_1[province[0]] = province_dict_1.get(province[0],
                                                                   0) + 1  #count number of times the first letter appears
                province_dict[province] = province[0] + str(province_dict_1[province[
                    0]])  #create region id with first letter of province and number of times it occurs
                regid_string += f"{case_name}  {ascii_all[regid_counter]} \n"
                regid_counter += 1
            input_fn_p = province_long
            input_fp_p = f'{input_fd}{input_fn_p}.xlsx'

            # Read excel file
            # Provincial
            input_df_p = read_workbook(input_fp_p)

            #Demand
            demand_y = custom_reader_2(input_df_p, 'Demand', years)
            demand_y['heat'] = [0] * len(demand_y['electricity'])
            #convert demand_y from 0.1GWh to MWy
            for e in demand_y:
                demand_y[e] = [i / 8760 * 100000 for i in demand_y[e]]

            #DemandProfile
            #reprocess demand profile
            input_df_p['DemandProfile'] = input_df_p['DemandProfile'].drop(input_df_p['DemandProfile'].columns[[1, 2]],
                                                                           axis=1)
            input_df_p['DemandProfile'].loc[-1] = input_df_p['DemandProfile'].columns  # add headers as a new row
            input_df_p['DemandProfile'].index = input_df_p['DemandProfile'].index + 1  # shift index
            input_df_p['DemandProfile'] = input_df_p['DemandProfile'].sort_index()  # re-order index
            input_df_p['DemandProfile'].columns = ['Carrier', 'electricity']
            demand_ts = custom_reader_2(input_df_p, 'DemandProfile', hours_year)
            demand_ts['heat'] = [1] * len(demand_ts['electricity'])  #hardcoded for now, will be removed when heat is added

            #REProfile
            re_ts = custom_reader_2(input_df_p, 'REProfile', hours_year)
            for re_ in re_ts:
                re_ts[re_] = shift_profile(re_ts[re_], 8)
            #remove profiles which have no values

            #TechCapacity
            #read TechCapacity, create candidate technologies
            tech_p = strstrip(input_df_p['TechCapacity'], ['Technology', 'Technology Type'])  #technology in province
            #rows whose Technology Type is labelled 'unknown' are dropped
            tech_p = tech_p[tech_p['Technology Type'] != 'unknown']
            if "capacity addition" in tech_p.columns:
                tech_p = tech_p.dropna(axis=0, subset=['capacity addition'])
            else:
                tech_p = tech_p.dropna(axis=0, subset=['sum of capacity'])

            #drop rows whose technology and technology type are not in TechMap,
            #they cannot be given a technology name
            tech_p, cap_warnings = check_tech_capacity(tech_p, tm, province_long)
            for w in cap_warnings:
                print(f"warning: {w}")

            dict_status = {'existing': 'Exist', 'exogenous': 'Constr', 'endogenous': 'Plan', 'generic': 'Endo'}
            tech_p['tech_name'] = [tm[i][j] for i, j in zip(tech_p['Technology'], tech_p['Technology Type'])]
            tech_p['msg_name'] = province + '_' + tech_p['tech_name'] + '_' + tech_p['status'].replace(dict_status)

            tech_p_dict = {}
            activity_count = {}

            for _, row in tech_p.iterrows():
                msg_name = row['msg_name']
                tech_name = row['tech_name']
                status = row['status']
                year = row['start year']
                value = row['capacity addition']

                if msg_name not in tech_p_dict:
                    tech_p_dict[msg_name] = {
                        'tech_name': tech_name,
                        'status': status,
                        'existing': {},
                        'exogenous': {},
                        'endogenous': {}
                    }

                if year >= 0:
                    tech_p_dict[msg_name][status][int(year)] = value
                else:
                    tech_p_dict[msg_name][status]['someyear'] = value
            #add also those generic techs to the dict
            for t_ in generics_dict[c]:
                tech_name = re.sub(r'[^A-Za-z0-9]', '', t_).lower()
                #storage is written by add_storage, leave it out when off
                if not with_storage and tech_param[tech_name]['type'] in storage_types:
                    continue
                msg_name = province + '_' + tech_name + '_' + 'generic'
                status = 'generic'
                value = 0
                if msg_name not in tech_p_dict:
                    tech_p_dict[msg_name] = {
                        'tech_name': tech_name,
                        'status': status,
                        'existing': {},
                        'exogenous': {},
                        'endogenous': {},
                    }
            #            tech_p_dict[msg_name][status]['someyear'] = value

            for k_, t_ in tech_p_dict.items():
                act_n = activity_count.get(tech_param[t_['tech_name']]['fuel'], 0)  #get index of activity based on fuel
                activity_count[tech_param[t_['tech_name']]['fuel']] = act_n + 1
                tech_p_dict[k_]['activity'] = ascii_all[act_n]

            # load regions (timeslices), capfac, demand curves
            # use 1 day per quarter to create profiles
            # get values of demand for each day
            demand_ts_inday = {}
            demand_tot_inday = {}
            demand_frac_ts_day = {}
            demand_frac_ts_year = {}
            demand_frac_day_year = {}
            demand_tot_inyear = {}

            re_ts_inday = {}

            #demand: get ldr: sum(demand of day i.e. season)/sum(demand of year i.e. all days), demand/sum(demand of day), adb: demand/sum(demand of year)
            #capfac: get fraction of max production (1)

            for eform in demand_ts.keys():
                demand_ts_inday[eform] = sample_days(demand_ts[eform], hoyinday)
                demand_tot_inday[eform] = {day: sum(v) for day, v in demand_ts_inday[eform].items()}
                demand_tot_inyear[eform] = sum(demand_tot_inday[eform].values())

                demand_frac_ts_day[eform] = {}
                demand_frac_ts_year[eform] = {}
                demand_frac_day_year[eform] = {}
                tot_year = demand_tot_inyear[eform]
                for day in hoyinday:
                    tot_day = demand_tot_inday[eform][day]
                    demand_frac_ts_day[eform][day] = [i / tot_day if tot_day != 0 else 0
                                                      for i in demand_ts_inday[eform][day]]
                    demand_frac_ts_year[eform][day] = [i / tot_year if tot_year != 0 else 0
                                                       for i in demand_ts_inday[eform][day]]
                    demand_frac_day_year[eform][day] = tot_day / tot_year if tot_year != 0 else 0

            for res in re_ts.keys():
                res_ = re.sub(r'[^a-zA-Z0-9 \n\.]', '', res).lower()  #change to lower case
                re_ts_inday[res_] = sample_days(re_ts[res], hoyinday)
                for msg_name, tp in tech_p_dict.items():
                    if tp['tech_name'] == res_:
                        tech_p_dict[msg_name]['capfac'] = re_ts_inday[res_]

            #todo: note - unknown technology types are dropped

            ##### MESSAGE #####

            # File .adb

            #the fuels used in this province become the forms of the Fuel level
            energyforms_fuel_s = ""
            fuel_c = {}
            counter = fuel_letter0
            for key in fuel_y.keys():
                if activity_count.get(key, 0) >= 1:  # if fuel is used for any activity
                    fuel_c[key] = ascii_lowercase[counter]  # generate a dict which reads fuel
                    counter += 1
                    #fuels are indented 3 spaces here, unlike the other forms
                    energyforms_fuel_s += (f"   {key} {fuel_c[key]}\n"
                                           f"   #\n")

            energyforms_s = energyforms_base_s + (f"Fuel {ef['Fuel']}\n"
                                                  "#\n"
                                                  f"{energyforms_fuel_s}"
                                                  "*\n")
            demand_s = ("demand:\n"
                        f"{lvl_elecdemand} ts {' '.join([str(round(i, 3)) for i in demand_y['electricity']])}\n"
                        f"{lvl_heatdemand} ts {' '.join([str(round(i, 3)) for i in demand_y['heat']])}\n"
                        )
            #load curve values are written to 6 decimals, the way MESSAGE
            #writes them itself. a bare str() gives the full float repr, eg
            #0.010416666666666666, which is far wider than any value MESSAGE
            #puts in this block
            loadcurve_s = ("loadcurve:\n"
                           f"year {year0}\n"
                           f"{lvl_elecdemand} {' '.join(f'{k:.6f}' for k in [j for i in demand_frac_ts_year['electricity'].values() for j in i])}\n"
                           f"{lvl_heatdemand} {' '.join(f'{0:.6f}' for k in [j for i in demand_frac_ts_year['heat'].values() for j in i])}\n"
                           #only 0s
                           )

            loadcurve_systems_s = ""

            systems_fuel_s = "systems:\n"
            for key in fuel_c.keys():
                fuel_s = (f"Fuel_{key} a\n"
                          f"    moutp	{fuel_c[key]}-{ef['Fuel']} c 1\n"
                          f"    vom	ts {' '.join(str(i * 8760) for i in fuel_y[key])}\n"
                          "#\n"
                          "*\n")
                systems_fuel_s += fuel_s

            systems_pp_s = ""
            ldb_systems_pp_s = ""
            ldr_loadcurve_systems_s = ""

            tech_count = 0
            relations1_ramp_s = ""
            ldb_relations1_ramp_s = ""
            for msg_name in tech_p_dict:

                val = tech_p_dict[msg_name]
                tech_name = val['tech_name']
                hist_tab[msg_name] = val['existing']
                lt_tab[msg_name] = tech_param[tech_name]['lifetime']
                #all capacity additions after year0 should be included as bdc constraint
                bdc = []
                for year in years:
                    bdc_val = 0
                    for y in years_inyearscontinuous[year]:
                        bdc_val += (val['exogenous'].get(y, 0) + val['existing'].get(y, 0)) / years_intervalsafterprev[
                            year]
                    bdc.append(bdc_val)
                bdc_all_s = f"    bdc fx ts {' '.join(str(v) for v in bdc)}\n"

                if val['status'] == 'existing':
                    fyear_s = ""
                    inv_s = f"    inv  c 0\n"
                    hist_s = f"    hisc 0. hc {' '.join(f'{k} {v}' for k, v in val['existing'].items() if (k < year0 and k >= year0 - tech_param[tech_name]['lifetime']))}\n"
                    bdc_s = bdc_all_s
                elif val['status'] == 'exogenous':
                    fyear_s = ""
                    inv_s = f"    inv	ts {' '.join(str(i) for i in tech_param[tech_name]['capex'])}\n"
                    hist_s = ""
                    bdc_s = bdc_all_s
                elif val['status'] == 'endogenous':
                    continue
                    #Removing all endogenous so only considering generics
                    #inv_s = f"    inv	ts {' '.join(str(i) for i in tech_param[tech_name]['capex'])}\n"
                    #hist_s = ""
                    #bdc_s = ""
                elif val['status'] == 'generic':
                    if np.isnan(tech_param[tech_name]['genericfirstyear']):
                        fyear_s = ""
                    else:
                        fyear_s = f"    fyear   {int(tech_param[tech_name]['genericfirstyear'])} \n"
                    inv_s = f"    inv	ts {' '.join(str(i) for i in tech_param[tech_name]['capex'])}\n"
                    hist_s = ""
                    bdc_s = ""

                #create ramp constraints here
                ramp_s = ""
                con1c_s = ""
                con1a_s = ""

                if tech_param[tech_name]['type'] == "PP" and tech_param[tech_name]['ramprate']>0:
                    for id, t in enumerate(ts[:-1]): #for each time slice until second last timeslice
                        for rd in rampdir: #for each direction
                            # create ramp constraint names. MESSAGE relation codes
                            # are 4 characters, so rampname[:4] is the code and
                            # rampname the full name
                            rampname = f"{ascii_all[tech_count]}{id}{rd}"
                            relations1_ramp_s += (f"\n"
                                             f"{rampname} {rampname[:4]} o\n"
                                             f"    units  group: activity, type: energy, cost:US$'00/kWyr, upper:MWyr, lower:MWyr\n"
                                             f"    for_ldr	none\n"
                                             f"    upper	c 0\n"
                                             f"    type	None\n"
                                             f"*"
                                             )
                            con1c_s += f"    con1c {rampname[:4]}:tin	c {tech_param[tech_name]['ramprate'] * -1}\n"
                            #ts_count is the number of time slices, was written as 24*4
                            con1a_s += f"    con1a {rampname[:4]}:{ts[id]}	c {rampdir[rd]*-1*ts_count}\n"
                            con1a_s += f"    con1a {rampname[:4]}:{ts[id+1]}	c {rampdir[rd]*ts_count}\n"
                            ldb_relations1_ramp_s += ("\n"
                                f"{rampname} {rampname[:4]} o\n"
                                "*"
                                )



                for con in constraints_types['con1c']:
                    if tech_param[tech_name]['constraints'][con] == 1 and constraints_properties[con]['except'] != c:
                        con1c_ldr_s = ''
                        if not (isinstance(constraints_properties[con]['ldr'], float) and math.isnan(constraints_properties[con]['ldr'])):
                            con1c_ldr_s += f':{constraints_properties[con]['ldr']}'
                        #con[:4] because MESSAGE relation codes are 4 characters
                        con1c_s += f"    con1c {con[:4]}{con1c_ldr_s}	c 1\n"

                if tech_param[tech_name]['type'] in ["PP", "RE"]:

                    #the name MESSAGE gives this technology in the matrix, and
                    #so in the solution file: the fuel it burns, its activity
                    #letter, the form it produces and the subregion it is in.
                    #the time slice and the year follow it there
                    if tech_param[tech_name]['type'] == "PP":
                        minp_s = f"    minp	{fuel_c[tech_param[tech_name]['fuel']]}-{ef['Fuel']} 1.\n"
                        moutp_lvl = lvl_nonvre
                        tech_codes[msg_name] = (f"f{fuel_c[tech_param[tech_name]['fuel']]}"
                                                f"{val['activity']}g....{ascii_all[regid_counter-1]}")
                    elif tech_param[tech_name]['type'] == "RE":
                        minp_s = ""  #if RE, then no input fuel
                        moutp_lvl = lvl_vre
                        tech_codes[msg_name] = (f"f.{val['activity']}h...."
                                                f"{ascii_all[regid_counter-1]}")
                        try:
                            #6 decimals throughout, the width MESSAGE uses in
                            #the load curve blocks
                            capfac_list = [j for i in val['capfac'].values() for j in i]
                            loadcurve_systems_s += (
                                f"systems.{msg_name}.{val['activity']}.capfac {' '.join(f'{i:.6f}' for i in capfac_list)}\n")
                            ldr_loadcurve_systems_s += (f"systems.{msg_name}.{val['activity']}.capfac\n"
                                                        f"{year0}\n")
                            ldr_loadcurve_systems_s += f"{' '.join(f'{i:.6f}' for i in [1] * days_year)}\n"  #change here for day
                            for i in val['capfac'].values():
                                ldr_loadcurve_systems_s += (f"1.000000\n"
                                                            f"{' '.join(f'{j:.6f}' for j in i)}\n")
                        except:
                            print(f"Issue with capfac for technology {msg_name} so not loaded")

                    tech_s = (f"{msg_name} {val['activity']}\n"  #@
                              f"{minp_s}"
                              f"    moutp	{moutp_lvl} c {tech_param[tech_name]['efficiency']}\n"
                              f"{fyear_s} "
                              f"    optm	c {tech_param[tech_name]['availability']}\n"
                              f"    pll	c {tech_param[tech_name]['lifetime']}\n"
                              f"{inv_s}"
                              f"    fom	ts {' '.join(str(i) for i in tech_param[tech_name]['fom'])}\n"
                              f"    vom	ts {' '.join(str(i * 8760) for i in tech_param[tech_name]['vom'])}\n"
                              f"{hist_s}"
                              f"    ctime	c {tech_param[tech_name]['construction time']}\n"
                              f"{bdc_s}"
                              f"{con1c_s}"
                              f"{con1a_s}"
                              f"    con1a CO2L	c {tech_param[tech_name]['emissionfactor']}\n"
                              "#\n"
                              "*\n")

                    ldb_tech_s = ldb_stub(msg_name, val['activity'])

                    systems_pp_s += tech_s
                    ldb_systems_pp_s += ldb_tech_s

                tech_count += 1

            #transmission and distribution, read from the TandDData sheet
            systems_trans_s = ""
            for _, td in td_param.iterrows():
                if td['province prefix'] == 1:
                    td_name = f"{province}_{td['tech']}"
                else:
                    td_name = td['tech']
                systems_trans_s += f"{td_name} {td['activity']}\n"
                #heat_dummy has no input, so a blank minp leaves the line out
                if not pd.isna(td['minp']):
                    systems_trans_s += f"    minp	{td['minp']} 1.\n"
                if not pd.isna(td['moutp']):
                    systems_trans_s += f"    moutp	{td['moutp']} c {numstr(td['efficiency'])}\n"
                systems_trans_s += f"    inv	c {float(td['inv'])}\n"
                systems_trans_s += f"    fom	c {float(td['fom'])}\n"
                systems_trans_s += f"    vom	c {float(td['vom'])}\n"
                systems_trans_s += ("#\n"
                                    "*\n")

            # File _adb.ldr

            ldr_loadcurve_season_s = {}
            for k, v in demand_frac_ts_day.items():
                for i in v:
                    if k not in ldr_loadcurve_season_s:  #start with the day of year fractions
                        ldr_loadcurve_season_s[
                            k] = f"{' '.join(f'{l:.6f}' for l in demand_frac_day_year[k].values())}\n"
                    ldr_loadcurve_season_s[k] += (f"1.000000\n"
                                                  f"{' '.join(f'{j:.6f}' for j in v[i])}\n")
            ldr_loadcurves_s = (f"loadcurves: \n"
                                f"{lvl_elecdemand}\n"
                                f"{year0}\n"
                                f"{ldr_loadcurve_season_s['electricity']}"
                                f"{lvl_heatdemand}\n"
                                f"{year0}\n"
                                f"{ldr_loadcurve_season_s['heat']}"
                                )
            relationsp_s = ""
            ldb_relationsp_s = ldb_relations1_ramp_s
            relations1_s = relations1_ramp_s
            ldb_relations1_s = ("\n"
                                "CO2Limit CO2L o\n"
                                "*"
                                )

        else:  #else if not province, so is main
            case_name = main_name
            energyforms_s = energyforms_base_s
            demand_s = "demand:\n"
            loadcurve_s = "loadcurve:\n"
            loadcurve_systems_s = ""
            systems_fuel_s = "systems:\n"
            systems_pp_s = ""
            ldb_systems_pp_s = ''
            counter_line = 0
            for id, line in interconnection_main.iterrows():
                tech_s = (f"{line['line_name']} {ascii_all[counter_line]}\n"
                          f"    minp  {lvl_elecdist}-{line['from']}_{main_name} 1.\n"
                          f"    moutp {lvl_elecdist}-{line['to']}_{main_name} c {numstr(ic_param['efficiency'])}\n"
                          f"    pll	c {numstr(ic_param['lifetime'])}\n"
                          f"    inv	c {float(ic_param['inv'])}\n"
                          f"    fom	c {float(ic_param['fom'])}\n"
                          f"    vom	c {float(ic_param['vom'])}\n"
                          f"    hisc 0. hc {numstr(ic_param['historic capacity year'])} {line['value']}\n"
                          "#\n"
                          "*\n")

                ldb_tech_s = ldb_stub(line['line_name'], ascii_all[counter_line])

                systems_pp_s += tech_s
                ldb_systems_pp_s += ldb_tech_s
                counter_line += 1

            relations1_s = (f"\n"
                            f"CO2Limit CO2L o\n"
                            f"    units	group: activity, type: weight, cost:US$'00/ton, upper:kton, lower:kton\n"
                            f"    for_ldr	none\n"
                            f"    upper	ts {' '.join(str(i) for i in emissions['emissions'])} \n"
                            f"    lower	c 0\n"
                            f"    type	None\n"
                            f"*"
                            )
            for con in constraints_names:
                #the values go on whichever side ConstraintsTypes asks for
                values = ' '.join(str(i) for i in constraints[con])
                if constraint_bounds[con] == 'lower':
                    bounds_s = f"    upper	c 0\n    lower	ts {values} \n"
                else:
                    bounds_s = f"    upper	ts {values} \n    lower	c 0\n"
                relations1_s += (f"\n"
                                 f"{con} {con[:4]} o\n"
                                 f"    units	group: capacity, type: power, cost:US$'00/kWyr, upper:MWyr, lower:MWyr\n"
                                 f"    for_ldr	none\n"
                                 f"{bounds_s}"
                                 f"    type	None\n"
                                 f"*"
                                 )

            ldb_relations1_s = ("\n"
                                "CO2Limit CO2L o\n"
                                "*"
                                )
            for con in constraints_names:
                ldb_relations1_s += (
                                    "\n"
                                    f"{con} {con[:4]} o\n"
                                    "*"
                )

            systems_trans_s = ""
            ldr_loadcurves_s = f"loadcurves: \n"
            ldr_loadcurve_systems_s = ""
            subregion_s = f"subregions:     {' '.join(str(k) for k in case_names_prov)} \n"

        # File .adb
        tdb_s = "tdb: empty\n"
        adb_s = f"adb: {case_name}\n"
        problem_s = f"problem: {case_name}\n"
        description_s = "description:\n"
        drate_s = f"drate: {drate * 100}\n"
        timesteps_s = f"timesteps: {' '.join(str(x) for x in years_incl_by)}\n" 


        loadregions_s = (f"loadregions: \n"
                         f"ltype  ordered seasonal 1 0 \n"
                         f"year   {year0} 1 {int(ts_count)} \n"
                         f"name   {' '.join(str(i) for i in ts)} \n"
                         f"length {' '.join(str(i) for i in lengths)} \n"
                         )

        relations_s = ("relationsc:\n"
                       "relationsp:\n"
                       "relationss:\n"
                       f"relations1:{relations1_s}\n"
                       "relations2:\n"
                       "variables:\n")

        ldb_relations_s = ("relationsc:\n"
                           "relationsp:\n"
                           "relationss:\n"
                           f"relations1:{ldb_relations1_s}\n"
                           "relations2:\n"
                           "variables:\n")

        end_s = ("resources: \n"
                 "endata")

        # File _adb.ldr
        ldr_loadregions_s = (f"loadregions: \n"
                             f"ltype    seasonal \n"
                             f"year  {year0} {' '.join(seasons)} \n"
                             f"range {' '.join(d.isoformat() for d in season_dates)}\n"
                             )
        ldr_season_s = ""
        ts_perday_length = 1 / timesteps_day
        for idx, s in enumerate(seasons):
            s_s = (
                f"season    {s} anyday \n"
                f"day   anyday {daysinseason[idx]} \n"
                f"name  {' '.join(str(i) for i in ts[idx * 24:(idx + 1) * 24])}\n"
                f"length    {' '.join(str(i) for i in [ts_perday_length] * timesteps_day)}\n"
            )

            ldr_season_s += s_s

        # File .adb
        adb_string = (tdb_s + adb_s + problem_s + description_s + drate_s + timesteps_s + loadregions_s + energyforms_s
                      + demand_s + loadcurve_s + loadcurve_systems_s + relations_s + systems_fuel_s + systems_pp_s + systems_trans_s + end_s)

        # File _adb.ldb
        ldb_string = (tdb_s + adb_s + problem_s + description_s + drate_s + timesteps_s + loadregions_s + energyforms_s
                      + demand_s + loadcurve_s + loadcurve_systems_s + ldb_relations_s + systems_fuel_s + ldb_systems_pp_s + systems_trans_s + end_s)
        # File .ldr
        ldr_string = ldr_loadregions_s + ldr_season_s + ldr_loadcurves_s + ldr_loadcurve_systems_s


        # Create batch file to run MESSAGE

        # Create new folder
        if generate_main:
            output_fd = f"{output_base_fd}{case_name}"
        else:
            output_fd = f"{output_base_fd}/{main_name}/{case_name}"
        #remove folder if exists
        if os.path.exists(output_fd) and os.path.isdir(output_fd):
            shutil.rmtree(output_fd)
        os.makedirs(output_fd)

        # Copy all files over from orig folder to new folder
        shutil.copytree(orig_base_fd, output_fd, dirs_exist_ok=True)

        # Replace all file names with new case name

        for root, dirs, filenames in os.walk(output_fd):
            for fn in filenames:
                if orig_name in fn:
                    path = os.path.join(root, fn)
                    newpath = os.path.join(root, fn.replace(orig_name, case_name))
                    os.rename(path, newpath)

        #write out adb file
        adb_new_path = f'{output_fd}/data/{case_name}.adb'
        with open(adb_new_path, 'w') as file:
            file.write(adb_string)

        #write out ldb file
        ldb_new_path = f'{output_fd}/data/{case_name}_adb.ldb'
        with open(ldb_new_path, 'w') as file:
            file.write(ldb_string)

        #write out ldr file
        ldr_new_path = f'{output_fd}/data/{case_name}_adb.ldr'
        with open(ldr_new_path, 'w') as file:
            file.write(ldr_string)

        # .gen file
        #every case is solved over the same horizon, so every .gen gets ntrun
        gen_new_path = f'{output_fd}/data/{case_name}.gen'
        set_ntrun(gen_new_path, ntrun)

        #add a line to gen file
        if case_name == main_name and generate_main:
            with open(gen_new_path, 'a') as file:
                file.write(subregion_s)

            #.cin file, which says which tables cap writes. the one copied
            #from the reference case only reports the parent itself, so the
            #res file would have no per technology columns
            cin_new_path = f'{output_fd}/data/{case_name}.cin'
            with open(cin_new_path, 'w') as file:
                file.write(cin_string(case_names_prov))

        #Edit MESSAGE directories
        # the folder is copied to the MESSAGE directory at the end of the
        # scenario, after add_storage has written into the staging copy, so
        # that what runs and what is staged are the same model
        if generate_main:
            MESSAGE_mod_fd = f"{MESSAGE_main_fd}/{case_name}"
        else:
            MESSAGE_mod_fd = f"{MESSAGE_root_fd}{case_name}"
        os.makedirs(MESSAGE_mod_fd, exist_ok=True)  #so a case can be rebuilt
        case_folders.append((output_fd, MESSAGE_mod_fd))

        # Create batch file

        for sol in solvers:
            if sol == 'cplex':
                sol_s = (
                    f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/csol -v -s adb {case_name}_adb | {MESSAGE_fd}message_bin/tee {case_name}_adb.itl" \n'
                    f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/sol2dbm -s adb -o cplex  {case_name}" \n')

            elif sol == 'HiGHS':
                sol_s = (f'cd {MESSAGE_mod_fd}\\intm \n'
                         f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/highs --options_file={MESSAGE_fd}message_bin/highs_settings.txt {case_name}_adb.mps --solution_file {case_name}_adb_lin.sol | {MESSAGE_fd}message_bin/tee {case_name}_adb.itl" \n'
                         f'cd {MESSAGE_mod_fd} \n'
                         f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/sol2dbm -s adb -o glpk  {case_name}" \n'
                         )
            else:
                sol_s = ""

            MESSAGE_bat_path = f"{MESSAGE_mod_fd}/run_{sol}_{case_name}_adb.bat"

            if sol == 'cplex':
                # currently bat all only runs cplex
                bat_all_s += (f'cd {MESSAGE_mod_fd} \n'
                              f'start {MESSAGE_bat_path}\n'
                              )

            bat_s = (f'set MMS_HOME={MESSAGE_root_fd}\n'
                     f'set MSG_HOME={MESSAGE_root_fd}\n'
                     f'set MMS_PRO={MESSAGE_mms_fils}mms.pro\n'
                     f'set MSG_ROOT={MESSAGE_fd}\n'
                     f'set MSG_BIN={MESSAGE_fd}message_bin\n'
                     f'set LANGUAGE=english \n'
                     f'set USER=unknown \n'
                     f'set LS_COLORS= \n'
                     f'C: \n'
                     f'cd {MESSAGE_mod_fd}/intm\n'
                     f'del {case_name}_adb.* \n'
                     f'del {case_name}_adb_lin.sol \n'
                     f'cd {MESSAGE_mod_fd} \n'
                     f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/mxg -f mxgerr -o cplex -v -n nbd -s adb   -x intm/powerchs.mps -W IAEA   {case_name}" \n'
                     f'{sol_s} \n'
                     f'copy sdbvars.txt sdbvars_{case_name}_adb.txt \n'
                     f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/cap -s adb -c {case_name} -t {case_name}  -T \'{case_name}, adb\'  -g spr -o {case_name}_adb -p \'MESSAGE Int_V2\'  {case_name}" \n'
                     f'pause \n')
            with open(MESSAGE_bat_path, 'w') as file:
                file.write(bat_s)

        with open(MESSAGE_bat_all_path, 'w') as file:
            file.write(bat_all_s)

        #register the case the first time it is built, mms.pro says which
        #cases are already known so that the entries are not duplicated
        if case_name not in already_registered:
            create_reg_func(MESSAGE_mms_fils, case_name, generate_main, mn=main_name)
            already_registered.add(case_name)

    #todo: essential files to change: ldr, ldb, adb, dic, chkunits, chn

    if generate_main:
        #the matrix name of every technology, which is how the results script
        #picks its time slices out of the solution file. the time slice and
        #the year follow the code there, as in fmag....a.aaa024
        codes_df = pd.DataFrame(sorted(tech_codes.items()),
                                columns=['technology', 'code'])
        codes_df.to_csv(f"{main_case_fd}/techcodes.csv", index=False)

        regid_new_path = f'{main_case_fd}/regid'  #write out regid_string
        with open(regid_new_path, 'w') as file:
            file.write(regid_string)

        #write out historic installation data
        hist_df = pd.DataFrame(hist_tab)
        hist_df.to_csv(f"{main_case_fd}/hist_tab.csv")

        lt_df = pd.DataFrame([lt_tab])
        lt_df.to_csv(f"{main_case_fd}/lt_tab.csv")

        #todo: cin file, hourly
        #todo: convert from cap file to excel

    #every case of this scenario has been written by now, so the storage
    #blocks can be added to the staging copies
    if with_storage:
        add_storage.add_storage_to_scenario(input_fn, input_fd, main_name)

    #and only then are the finished cases copied to the MESSAGE directory
    for output_fd, MESSAGE_mod_fd in case_folders:
        shutil.copytree(output_fd, MESSAGE_mod_fd, dirs_exist_ok=True)

    return bat_all_s


#build every scenario which is flagged to be made. run_all_adb.bat collects
#the cases of every scenario, so it launches the whole set. guarded so that
#add_storage can import the helpers without building anything
if __name__ == '__main__':
    bat_all_s = ""
    for _, scenario in read_create_cases(create_cases_fp).iterrows():
        main_name = f"{scenario['main']}_{scenario['input_fn']}"
        with_storage = scenario['add_storage'] == 1
        #a blank ntrun means work it out from the Years sheet
        ntrun = scenario.get('ntrun')
        ntrun = None if pd.isna(ntrun) or not ntrun else int(ntrun)
        print(f"=== {scenario['input_fn']} -> {main_name}"
              f"{' with storage' if with_storage else ''} ===")
        bat_all_s = generate_scenario(scenario['input_fn'], scenario['input_fd'],
                                      main_name, with_storage, bat_all_s, ntrun)
