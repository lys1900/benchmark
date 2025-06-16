#YL, IAEA, 2025 June
import collections
import pandas as pd
import openpyxl
import os
import shutil
from string import ascii_lowercase, ascii_uppercase
import re
import csv


orig_name = 'CountryA'
input_fd = 'E:/Work/benchmark/input/'
input_fn = 'NW' #needs to be short
province_fn = 'E:/Work/benchmark/input/runprovince.csv'
create_reg = 0
MESSAGE_fd = "C:/Programs/MESSAGE_INT/"
MESSAGE_root_fd = f"{MESSAGE_fd}models/"
MESSAGE_bat_all_path = f"{MESSAGE_root_fd}/run_all_adb.bat"
MESSAGE_mms_fils = f"{MESSAGE_root_fd}/mms_fils/"
#dummy seasons only to represent days
seasons = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November","December"]
daysinseason = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
doyinquarter = {i:[(i*91)+45,(i*91)+46,(i*91)+47] for i in range(0,4)}
solvers = ['HiGHS', 'cplex']
bat_all_s = ""


def strstrip(df, cols):
    """
    function to strip columns and change to lower case
    """
    df[cols] = df[cols].replace(r'[^A-Za-z0-9]', '', regex=True)
    for c in cols:
        df[c] = df[c].str.lower()
    return df

def custom_reader(input, sheetname = 'Demand'):
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

def custom_reader_2 (input, sheetname = 'Demand'):
    """
    Custom reader to read some excel sheets into dict
    :param sheetname: the sheet name to read
    :return: a dictionary of dictionary, with keys parameter
    """
    dict_out = input[sheetname].set_index(input[sheetname].columns[0]).dropna(axis=1, how='all').to_dict('list')

    return dict_out

def shift_profile(profi, gmt):
    """
    Shift profile by gmt hours to local time zone
    """
    profi_shifted = profi[-gmt:] + profi[:-gmt]

    return profi_shifted


province_list =[]
with open(province_fn, newline='') as f:
    reader = csv.reader(f)
    for row in reader:
        province_list.append(row[0])

hoyinday = {}
for i,j in doyinquarter.items():
    for k in j:
        hoyinday[k] = list(range((k-1)*24, k*24))

# Nationwide
input_fp = f'{input_fd}{input_fn}.xlsx'
input_xl = pd.ExcelFile(input_fp)
input_sh = input_xl.sheet_names

input_df_all = {}
for i in input_sh:
    input_df_all[i] = input_xl.parse(i)

# Read all parameters
#General
drate = input_df_all["General"].loc[input_df_all["General"]["Parameter"] == "Discount rate", "Value"].item()
days_year = int(input_df_all["General"].loc[input_df_all["General"]["Parameter"] == "Days per year", "Value"].item())
timesteps_day = int(input_df_all["General"].loc[input_df_all["General"]["Parameter"] == "Timesteps per day", "Value"].item())
year0 = int(input_df_all["General"].loc[input_df_all["General"]["Parameter"] == "First year", "Value"].item())
yearx = int(input_df_all["General"].loc[input_df_all["General"]["Parameter"] == "Last year", "Value"].item())
years = list(input_df_all["Years"]['years'])
baseyear = year0-1


#TechMap
tech_map = input_df_all["TechMap"]
tech_map = strstrip(tech_map, ['Technology name', 'Technology', 'Technology Type'])
tm = {}
for k, g in tech_map.groupby('Technology'):
    tm[k] = dict(zip(g['Technology Type'], g['Technology name']))

#FuelPrice
fuel_param = input_df_all["FuelPrice"].set_index('Fuel').T
fuel = list(fuel_param.index)
fuel_price = fuel_param.to_dict()

#TechnologyData, TechnologyCapex, fom, vom

tech_param = strstrip(input_df_all["TechData"], ['Technology name', 'Technology', 'Technology Type']).set_index('Technology name').T.to_dict()
tech_capex = strstrip(input_df_all["TechCapex"], ['Technology', 'Technology Type'])
tech_fom = strstrip(input_df_all["fom"], ['Technology', 'Technology Type'])
tech_vom = strstrip(input_df_all["vom"], ['Technology', 'Technology Type'])

for key in tech_param:
    try:
        tech_param[key]['capex'] = tech_capex[(tech_capex['Technology'] == tech_param[key]['Technology']) & (tech_capex['Technology Type'] == tech_param[key]['Technology Type'])][years].to_dict('tight')['data'][0]
    except:
        print(f"issue with '{key}', will not have capex")
        tech_param[key]['capex'] = 0

    try:
        tech_param[key]['fom'] = tech_fom[(tech_fom['Technology'] == tech_param[key]['Technology']) & (tech_fom['Technology Type'] == tech_param[key]['Technology Type'])][years].to_dict('tight')['data'][0]
    except:
        print(f"issue with '{key}', will not have fom")
        tech_param[key]['fom'] = 0

    try:
        tech_param[key]['vom'] = tech_vom[(tech_vom['Technology'] == tech_param[key]['Technology']) & (tech_vom['Technology Type'] == tech_param[key]['Technology Type'])][years].to_dict('tight')['data'][0]
    except:
        print(f"'issue with '{key}', will not have vom")
        tech_param[key]['vom'] = 0

#TandDData
#td_param = input_df_all["TandDData"].set_index('tech').T.to_dict()
interconnection = input_df_all["Interconnection"]

#FuelPrice
fuel_y = custom_reader_2(input_df_all, 'FuelPrice')


for province_long in province_list:
    province = province_long[0:6].replace(" ","") #needs to be short        p = row[0].

    input_fn_p = province_long
    input_fp_p = f'{input_fd}{input_fn_p}.xlsx'
    case_name = f'{province_long}_{input_fn}' #todo:strip spaces

    # Read excel file
    # Provincial
    input_xl_p = pd.ExcelFile(input_fp_p)
    input_sh_p = input_xl_p.sheet_names

    input_df_p = {}
    for i in input_sh_p:
        input_df_p[i] = input_xl_p.parse(i)



    #Demand
    demand_y = custom_reader_2(input_df_p, 'Demand')
    demand_y['heat'] = [0] * len(demand_y['electricity'])
    #convert demand_y from 0.1GWh to MWy
    for e in demand_y:
        demand_y[e] = [i/ 8760 * 100 for i in demand_y[e]]

    #DemandProfile
    #reprocess demand profile
    input_df_p['DemandProfile'] = input_df_p['DemandProfile'].drop(input_df_p['DemandProfile'].columns[[1,2]],axis=1)
    input_df_p['DemandProfile'].loc[-1] = input_df_p['DemandProfile'].columns  # add headers as a new row
    input_df_p['DemandProfile'].index = input_df_p['DemandProfile'].index + 1  # shift index
    input_df_p['DemandProfile'] = input_df_p['DemandProfile'].sort_index()     # re-order index
    input_df_p['DemandProfile'].columns = ['Carrier', 'electricity']
    demand_ts = custom_reader_2(input_df_p, 'DemandProfile')
    demand_ts['heat'] = [1] * len(demand_ts['electricity']) #hardcoded for now, will be removed when heat is added

    #REProfile
    re_ts = custom_reader_2(input_df_p, 'REProfile')
    for re_ in re_ts:
        re_ts[re_] = shift_profile(re_ts[re_], 8)
    #remove profiles which have no values



    #TechCapacity
    #todo: read TechCapacity, create candidate technologies
    tech_p = strstrip(input_df_p['TechCapacity'], ['Technology', 'Technology Type']) #technology in province
    tech_p = tech_p[tech_p['Technology Type'] != 'unknown']
    tech_p = tech_p.dropna(axis=0, subset=['capacity addition'])


    dict_status = {'existing':'Exist','exogenous':'Constr', 'endogenous':'Plan' }
    tech_p['mapname'] = [tm[i][j] for i, j in zip(tech_p['Technology'], tech_p['Technology Type'])]
    tech_p['messagename'] = province + '_' + tech_p['mapname'] + '_' + tech_p['status'].replace(dict_status)

    tech_p_dict = {}
    activity_count = {}
    ascii_all = ascii_lowercase + ascii_uppercase

    for _, row in tech_p.iterrows():
        messagename = row['messagename']
        mapname = row['mapname']
        status = row['status']
        year = row['start year']
        value = row['capacity addition']

        if messagename not in tech_p_dict:
            tech_p_dict[messagename] = {
                'mapname':mapname,
                'status':status,
                'existing':{},
                'exogenous':{},
                'endogenous':{}
            }

        if year >= 0:
            tech_p_dict[messagename][status][int(year)] = value
        else:
            tech_p_dict[messagename][status]['someyear'] = value

        act_n = activity_count.get(tech_param[mapname]['fuel'], 0) #get index of activity based on fuel
        activity_count[tech_param[mapname]['fuel']] = act_n + 1
        tech_p_dict[messagename]['activity'] = ascii_all[act_n]

    # load regions (timeslices), capfac, demand curves
    # use 3 days per quarter to create profiles
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
        demand_ts_inday[eform] = {}
        demand_tot_inday[eform] = {}
        demand_frac_ts_day[eform] = {}
        demand_frac_ts_year[eform] = {}
        demand_frac_day_year[eform] = {}

        for day, v in hoyinday.items():
            demand_ts_inday[eform][day] = [demand_ts[eform][i] for i in v]
            demand_tot_inday[eform][day] = sum(demand_ts_inday[eform][day])
        demand_tot_inyear[eform] = sum([j for i,j in demand_tot_inday[eform].items()])
        for day in hoyinday:
            demand_frac_ts_day[eform][day] = [i/demand_tot_inday[eform][day] if demand_tot_inday[eform][day] != 0 else 0 for i in demand_ts_inday[eform][day]]
            demand_frac_ts_year[eform][day] = [i/demand_tot_inyear[eform] if demand_tot_inyear[eform] != 0 else 0 for i in demand_ts_inday[eform][day]]
            demand_frac_day_year[eform][day] = demand_tot_inday[eform][day]/demand_tot_inyear[eform] if demand_tot_inyear[eform] !=0 else 0

    for res in re_ts.keys():
        res_ = re.sub(r'[^a-zA-Z0-9 \n\.]', '', res).lower()
        re_ts_inday[res_] = {} #change to lower case
        for day, v in hoyinday.items():
            re_ts_inday[res_][day] = [re_ts[res][i] for i in v]
        for _, tp in tech_p_dict.items():
            if tp['mapname'] == res_:
                tech_p_dict[_]['capfac'] = re_ts_inday[res_]


    #todo: note - unknown technology types are dropped
    #todo: add renewable techs

    ##### MESSAGE #####

    # File .adb
    tdb_s = "tdb: empty\n"
    adb_s = f"adb: {case_name}\n"
    problem_s = f"problem: {case_name}\n"
    description_s = "description:\n"
    drate_s = f"drate: {drate * 100}\n"
    years_incl_by = [baseyear] + years
    timesteps_s = f"timesteps: {' '.join(str(x) for x in years_incl_by)}\n" #todo: change this to match years user specify
    ts = []
    lengths = []

    ts_count = days_year * timesteps_day
    ts_length = round(1/ts_count, 6)

    for i in ascii_lowercase[0:days_year]:
        for j in ascii_lowercase[0:timesteps_day]:
            ts.append(f"{i}a{j}")
            lengths.append(ts_length)


    loadregions_s = (f"loadregions: \n"
                     f"ltype  ordered seasonal 1 0 \n"
                     f"year   {year0} 1 {int(ts_count)} \n"
                     f"name   {' '.join(str(i) for i in ts)} \n"
                     f"length {' '.join(str(i) for i in lengths)} \n"
                     )
    #todo: to be changed to be more dynamic to allow for more technologies and other formulations
    energyforms_fuel_s = ""
    fuel_c = {}
    counter = 11
    for key in fuel_y.keys():
        if activity_count.get(key, 0) >= 1: # if fuel is used for any activity
            fuel_c[key] = ascii_lowercase[counter] # generate a dict which reads fuel
            counter += 1
            energyforms_fuel_s += (f"   {key} {fuel_c[key]}\n"
                                   f"   #\n")

    energyforms_s = ("energyforms: \n"
                     "Final a\n"
                     "# \n"
                     "    ElectricityDemand b l \n"
                     "    # \n"
                     "    HeatDemand c l \n"
                     "    # \n"
                     "*\n"
                     "Distribution d\n"
                     "#\n" 
                     "    ElectricityDistribution e\n"  
                     "    #\n" 
                     "*\n"
                     "Transmission f\n"
                     "#\n" 
                     "    ElectricityNonVRE g\n"  
                     "    #\n" 
                     "    ElectricityVRE h\n"  
                     "    #\n" 
                     "    Heat i\n"  
                     "    #\n" 
                     "*\n"
                     "Fuel j\n"
                     "#\n"
                     f"{energyforms_fuel_s}"
                     "*\n")
    demand_s =       ("demand:\n"
                     f"b-a ts {' '.join([str(round(i*1000000/8760,3)) for i in demand_y['electricity']])}\n"
                     f"c-a ts {' '.join([str(round(i*1000000/8760,3)) for i in demand_y['heat']])}\n"
                      )
    loadcurve_s =    ("loadcurve:\n"
                     f"year {year0}\n"
                     f"b-a {' '.join(str(k) for k in [j for i in demand_frac_ts_year['electricity'].values() for j in i])}\n"
                     f"c-a {' '.join(str(0) for k in [j for i in demand_frac_ts_year['heat'].values() for j in i])}\n" #only 0s
                      )

    loadcurve_systems_s = ""

    relations_s=    ("relationsc:\n"
                     "relationsp:\n"
                     "relationss:\n"
                     "relations1:\n"
                     "relations2:\n"
                     "variables:\n")
    systems_fuel_s = "systems:\n"
    for key in fuel_c.keys():
        fuel_s = (   f"Fuel_{key} a\n"
                     f"    moutp	{fuel_c[key]}-j c 1\n"
                     f"    vom	ts {' '.join(str(i) for i in fuel_y[key])}\n"
                     "#\n" 
                     "*\n")
        systems_fuel_s += fuel_s

    systems_pp_s = ""
    ldb_systems_pp_s = ""
    ldr_loadcurve_systems_s = ""

    for pp in tech_p_dict:
        val = tech_p_dict[pp]
        tp_key = val['mapname']
        if val['status'] == 'existing':
            inv_s = f"    inv  c 0\n"
            hist_s = f"    hisc 0. hc {' '.join(f'{k} {v}' for k, v in val['existing'].items() if (k < year0 and k >= year0-tech_param[tp_key]['lifetime']))}\n"
            bdc = [val['existing'].get(year, 0) for year in years if year >= year0]
            bdc_s = f"    bdc fx ts {' '.join(str(v) for v in bdc)}\n"  #all capacity additions after year0 should be included as bdc constraint
        elif val['status'] == 'exogenous':
            inv_s = f"    inv	ts {' '.join(str(i) for i in tech_param[tp_key]['capex'])}\n"
            hist_s = ""
            bdc = [val['exogenous'].get(year, 0) for year in years if year >= year0]
            bdc_s = f"    bdc fx ts {' '.join(str(v) for v in bdc)}\n"
        elif val['status'] == 'endogenous':
            inv_s = f"    inv	ts {' '.join(str(i) for i in tech_param[tp_key]['capex'])}\n"
            hist_s = ""
            bdc_s = ""


        if tech_param[tp_key]['type'] in ["PP", "RE"]:

            if tech_param[tp_key]['type'] == "PP":
                minp_s = f"    minp	{fuel_c[tech_param[tp_key]['fuel']]}-j 1.\n"
                moutp_lvl = "g-f"
            elif tech_param[tp_key]['type'] == "RE":
                minp_s = "" #if RE, then no input fuel
                moutp_lvl = "h-f"
                try:
                    capfac_list = [j for i in val['capfac'].values() for j in i]
                    loadcurve_systems_s += (f"systems.{pp}.{val['activity']}.capfac {' '.join(str(i) for i in capfac_list)}\n")
                    ldr_loadcurve_systems_s += (f"systems.{pp}.{val['activity']}.capfac\n"
                                                f"{year0}\n")
                    ldr_loadcurve_systems_s += f"{' '.join(str(i) for i in [1]*12)}\n"
                    for i in val['capfac'].values():
                        ldr_loadcurve_systems_s += (f"1.000000\n"
                                                          f"{' '.join(str(j) for j in i)}\n")
                except:
                    print(f"Issue with capfac for technology {pp} so not loaded")

            tech_s = (f"{pp} {val['activity']}\n"
                         f"{minp_s}"
                         f"    moutp	{moutp_lvl} c {tech_param[tp_key]['efficiency']}\n"
                         f"    plf	c {tech_param[tp_key]['availability']}\n"
                         f"    pll	c {tech_param[tp_key]['lifetime']}\n"
                         f"{inv_s}"
                         f"    fom	ts {' '.join(str(i) for i in tech_param[tp_key]['fom'])}\n"
                         f"    vom	ts {' '.join(str(i) for i in tech_param[tp_key]['vom'])}\n"
                         f"{hist_s}"
                         f"    ctime	c {tech_param[tp_key]['construction time']}\n"
                         f"{bdc_s}"
                         "#\n" 
                         "*\n")

            ldb_tech_s = (f"{pp} {val['activity']}\n"
                         "*\n")


            systems_pp_s += tech_s
            ldb_systems_pp_s += ldb_tech_s

    #todo: fix string below to be populated programmatically
    systems_trans_s=("Transmission_ElectricityNonVRE a\n"
                     "    minp	g-f 1.\n"
                     "    moutp	e-d c 0.98\n"
                     "    inv	c 1000.0\n"
                     "    fom	c 10.0\n"
                     "    vom	c 8.76\n"
                     "#\n" 
                     "*\n"
                     "Transmission_ElectricityVRE b\n"
                     "    minp	h-f 1.\n"
                     "    moutp	e-d c 0.98\n"
                     "    inv	c 1000.0\n"
                     "    fom	c 10.0\n"
                     "    vom	c 8.76\n"
                     "#\n" 
                     "*\n"
                     "Distribution a\n"
                     "    minp	e-d 1.\n"
                     "    moutp	b-a c 0.98\n"
                     "    inv	c 1000.0\n"
                     "    fom	c 10.0\n"
                     "    vom	c 8.76\n"
                     "#\n" 
                     "*\n"
                     "heat_network a\n"
                     "    minp	i-f 1.\n"
                     "    moutp	c-a c 0.98\n"
                     "    inv	c 1000.0\n"
                     "    fom	c 10.0\n"
                     "    vom	c 8.76\n"
                     "#\n" 
                     "*\n"
                     "heat_dummy a\n"
                     "    moutp	i-f c 1\n"
                     "    inv	c 1.0\n"
                     "    vom	c 0.1\n"
                     "#\n"
                     "*\n"
                     )
    end_s = ("resources: \n"
             "endata")
    adb_string = (tdb_s + adb_s + problem_s + description_s + drate_s + timesteps_s + loadregions_s + energyforms_s
                  + demand_s + loadcurve_s + loadcurve_systems_s + relations_s + systems_fuel_s + systems_pp_s + systems_trans_s + end_s)

    # File _adb.ldb
    ldb_string = (tdb_s + adb_s + problem_s + description_s + drate_s + timesteps_s + loadregions_s +  energyforms_s
                  + demand_s + loadcurve_s +loadcurve_systems_s + relations_s + systems_fuel_s + ldb_systems_pp_s + systems_trans_s + end_s)

    # File _adb.ldr
    ldr_loadregions_s = (f"loadregions: \n"
                     f"ltype    seasonal \n"
                     f"year  {year0} {' '.join(seasons)} \n" #hardcoded to 12 months here
                     f"range {year0}-01-01 {year0}-02-01 {year0}-03-01 {year0}-04-01 {year0}-05-01 {year0}-06-01 {year0}-07-01 {year0}-08-01 {year0}-09-01 {year0}-10-01 {year0}-11-01 {year0}-12-01 {year0+1}-01-01\n"
                     )
    ldr_season_s = ""
    ts_perday_length = 1/timesteps_day
    for idx, s in enumerate(seasons):
        s_s = (
        f"season    {s} anyday \n"
        f"day   anyday {daysinseason[idx]} \n"
        f"name  {' '.join(str(i) for i in ts[idx*24:(idx+1)*24])}\n"
        f"length    {' '.join(str(i) for i in [ts_perday_length] * timesteps_day)}\n"
        )

        ldr_season_s += s_s

    ldr_loadcurve_season_s = {}
    for k, v in demand_frac_ts_day.items():
        for i in v:
            if k in ldr_loadcurve_season_s:
                ldr_loadcurve_season_s[k] += (f"1.000000\n"
                                             f"{' '.join(str(round(j,6)) for j in v[i])}\n")
            else:
                ldr_loadcurve_season_s[k] = f"{' '.join(str(round(l,6)) for l in demand_frac_day_year[k].values())}\n"
                ldr_loadcurve_season_s[k] += (f"1.000000\n"
                                             f"{' '.join(str(round(j,6)) for j in v[i])}\n")
    ldr_loadcurves_s = (f"loadcurves: \n"
                        f"b-a\n"
                        f"{year0}\n"
                        f"{ldr_loadcurve_season_s['electricity']}"
                        f"c-a\n"
                        f"{year0}\n"
                        f"{ldr_loadcurve_season_s['heat']}"
                        )
    ldr_string = ldr_loadregions_s + ldr_season_s + ldr_loadcurves_s + ldr_loadcurve_systems_s

    #.cin file


    #.gen file

    #.exp file

    # Create batch file to run MESSAGE

    # Create new folder
    output_base_fd = 'E:/Work/benchmark/MESSAGE_generator/MESSAGE_out/'
    if not os.path.exists(output_base_fd):
        os.makedirs(output_base_fd)

    output_fd = f"{output_base_fd}{case_name}"
    #remove folder if exists
    if os.path.exists(output_fd) and os.path.isdir(output_fd):
        shutil.rmtree(output_fd)
    os.makedirs(output_fd)

    # Copy all files over from orig folder to new folder
    orig_base_fd = f'E:/Work/benchmark/MESSAGE_generator/MESSAGE_orig/{orig_name}'
    shutil.copytree(orig_base_fd, output_fd, dirs_exist_ok=True)

    # Replace all file names with new case name

    for root, dirs, filenames in os.walk(output_fd):
        for fn in filenames:
            if orig_name in fn:
                path = os.path.join(root, fn)
                newpath = os.path.join(root, fn.replace(orig_name, case_name))
                os.rename(path, newpath)

    #write out adb file
    adb_new_path =  f'{output_fd}/data/{case_name}.adb'
    with open(adb_new_path,'w') as file:
        file.write(adb_string)

    #write out ldb file
    ldb_new_path =  f'{output_fd}/data/{case_name}_adb.ldb'
    with open(ldb_new_path,'w') as file:
        file.write(ldb_string)

    #write out ldr file
    ldr_new_path =  f'{output_fd}/data/{case_name}_adb.ldr'
    with open(ldr_new_path,'w') as file:
        file.write(ldr_string)

    #Edit MESSAGE directories
    # Copy new folder to MESSAGE directory
    MESSAGE_mod_fd = f"{MESSAGE_root_fd}{case_name}"
    os.makedirs(MESSAGE_mod_fd)
    shutil.copytree(output_fd, MESSAGE_mod_fd, dirs_exist_ok=True)

    # Create batch file


    for sol in solvers:
        if sol == 'cplex':
            sol_s = (f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/csol -v -s adb {case_name}_adb | {MESSAGE_fd}message_bin/tee {case_name}_adb.itl" \n'
                     f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/sol2dbm -s adb -o cplex  {case_name}" \n')

        elif sol == 'HiGHS':
            sol_s = (f'cd {MESSAGE_root_fd}{case_name}\\intm \n'
                     f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/highs --options_file={MESSAGE_fd}message_bin/highs_settings.txt {case_name}_adb.mps --solution_file {case_name}_adb_lin.sol | {MESSAGE_fd}message_bin/tee {case_name}_adb.itl" \n'
                     f'cd {MESSAGE_root_fd}{case_name} \n'
                     f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/sol2dbm -s adb -o glpk  {case_name}" \n'
                     )

        MESSAGE_bat_path = f"{MESSAGE_mod_fd}/run_{sol}_{case_name}_adb.bat"

        if sol == 'cplex':
            # currently bat all only runs cplex
            bat_all_s += (f'cd {MESSAGE_root_fd}{case_name} \n'
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
                    f'cd {MESSAGE_root_fd}{case_name}/intm\n'
                    f'del {case_name}_adb.* \n'
                    f'del {case_name}_adb_lin.sol \n'
                    f'cd {MESSAGE_root_fd}{case_name} \n'
                    f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/mxg -f mxgerr -o cplex -v -n nbd -s adb   -x intm/powerchs.mps -W IAEA   {case_name}" \n'
                    f'{sol_s} \n'
                    f'copy sdbvars.txt sdbvars_{case_name}_adb.txt \n'
                    f'{MESSAGE_fd}message_bin\\tcsh -c "{MESSAGE_fd}message_bin/cap -s adb -c {case_name} -t {case_name}  -T \'{case_name}, adb\'  -g spr -o {case_name}_adb -p \'MESSAGE Int_V2\'  {case_name}" \n'
                    f'pause \n')
        with open(MESSAGE_bat_path,'w') as file:
            file.write(bat_s)

    with open(MESSAGE_bat_all_path, 'w') as file:
        file.write(bat_all_s)

    #Only use the following if creating model for the first time
    if create_reg == 1:

        dir_s = (f"#call              answer\n"
                    f"supply             $MMS_HOME/{case_name}      \n"
                    f"grfdir             /tmp/grf/iaea           \n"
                    f"cin                $MMS_HOME/{case_name}/data \n"
                    f"tdb                $MMS_HOME/tdb           \n"
                    f"adb                $MMS_HOME/{case_name}/data \n"
                    f"ldb                $MMS_HOME/{case_name}/data \n"
                    f"upd                $MMS_HOME/{case_name}/data \n"
                    f"gen                $MMS_HOME/{case_name}/data \n"
                    f"data               $MMS_HOME/{case_name}/data \n"
                 )
        # Add line to mms.pro
        with open(f"{MESSAGE_mms_fils}mms.pro", "a") as f1:
            f1.write(f"{case_name}.dirfile   $MMS_HOME/mms_fils/{case_name}.dir\n")
        with open(f"{MESSAGE_mms_fils}glob.reg", "a") as f2:
            f2.write(f"{case_name}	{case_name}	{case_name}	+	empty\n")
        dir_f = f"{MESSAGE_mms_fils}{case_name}.dir"
        with open(dir_f, "w") as f3:
            f3.write(dir_s)


#todo: essential files to change: ldr, ldb, adb, dic, chkunits, chn