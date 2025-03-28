import collections
import pandas as pd
import openpyxl
import os
import shutil
from string import ascii_lowercase

orig_name = 'CountryA'
input_fd = 'E:/Work/benchmark/input/'
input_fn = 'master'
input_fp = f'{input_fd}{input_fn}.xlsx'
country = 'A'
case_name = f'{country}_{input_fn}'
switch = 0
MESSAGE_root_fd = "C:/Programs/MESSAGE_INT/models/"


# Read excel file
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

#TechnologyData, TechnologyCapex
tech_param = input_df_all["TechnologyData"].set_index('tech').T.to_dict()
tech_capex = input_df_all["TechnologyCapex"].to_dict('list')
for key in tech_param:
    if key in tech_capex:
        tech_param[key]['capex'] = tech_capex[key]
    else:
        print(f"'{key}' does not have capex")
        tech_param[key]['capex'] = 0

#TandDData
td_param = input_df_all["TandDData"].set_index('tech').T.to_dict()


def custom_reader(sheetname = 'Demand'):
    """
    Custom reader to read some excel sheets into dict
    :param sheetname: the sheet name to read
    :return: a dictionary in dictionary, with keys country and parameter
    """
    cols = []
    for i in input_df_all[sheetname].columns:
        cols.append(i.split('.', 1)[0])

    new_columns = list(zip(input_df_all[sheetname].loc[0], cols))
    input_df_all[sheetname].columns = pd.MultiIndex.from_tuples(new_columns)
    input_df = input_df_all[sheetname][1:]

    dict_out = collections.defaultdict(dict)
    for column in input_df:
        dict_out[column[1]][column[0]] = list(input_df[column])

    return dict_out


#Demand
demand_y = custom_reader('Demand')

#DemandProfile
demand_ts = custom_reader('DemandProfile')

#REProfile
re_ts = custom_reader('REProfile')

#FuelPrice
fuel_y = custom_reader('FuelPrice')

##### MESSAGE #####

# .adb, .adb.bu, _adb.exp
tdb_s = "tdb: empty\n"
adb_s = f"adb: {case_name}\n"
problem_s = f"problem: {case_name}\n"
description_s = "description:\n"
drate_s = f"drate: {drate * 100}\n"
timesteps_s = f"timesteps: {' '.join(str(x) for x in list(range(year0, yearx+1)))}\n" #todo: change this to match years user specify
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
                 f"year   {year0 + 1} 1 {int(ts_count)} \n"
                 f"name   {' '.join(str(i) for i in ts)} \n"
                 f"length {' '.join(str(i) for i in lengths)} \n"
                 )
#todo: to be changed to be more dynamic to allow for more technologies and other formulations
energyforms_fuel_s = ""
fuel_c = {}
counter = 11
for key in fuel_y[country].keys():
    fuel_c[key] = ascii_lowercase[counter] #generate a dict which reads fuel
    counter += 1
    energyforms_fuel_s += (f"{key} {fuel_c[key]}\n"
                           f"#\n")

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
                 f"b-a ts {' '.join([str(round(i*1000000/8760,3)) for i in demand_y[country]['electricity']])}\n"
                 f"c-a ts {' '.join([str(round(i*1000000/8760,3)) for i in demand_y[country]['heat']])}\n"
                 "loadcurve:"
                 f"year {year0+1}\n"
                 f"b-a {' '.join(str(i) for i in lengths)}\n"
                  )
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
                 f"    vom	ts {' '.join(str(i) for i in fuel_y[country][key])}\n"
                 "#\n" 
                 "*\n")
    systems_fuel_s += fuel_s

systems_pp_s = ""
for key in tech_param:
    if tech_param[key]['type'] == "PP":
        tech_s = (f"{key} a\n"
                     f"    minp	{fuel_c[tech_param[key]['fuel']]}-j 1.\n"
                     f"    moutp	g-f c {tech_param[key]['efficiency']}\n"
                     f"    plf	c {tech_param[key]['availability']}\n"
                     f"    pll	c {tech_param[key]['lifetime']}\n"
                     f"    inv	ts {' '.join(str(i) for i in tech_param[key]['capex'])}\n"
                     f"    fom	c {tech_param[key]['fom']}\n"
                     f"    vom	c {tech_param[key]['vom']}\n"
                     f"    ctime	c {tech_param[key]['construction time']}\n"
                     "#\n" 
                     "*\n")
        systems_pp_s += tech_s

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
              + demand_s + relations_s + systems_fuel_s + systems_pp_s + systems_trans_s + end_s)




# drate
# timesteps
# year
# name
# length
# read code for energy forms
# demand
# load curve
# fuel cost (vom ts)
# tech params
# transmission

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

#.gen

#write out adb file
adb_new_path =  f'{output_fd}/data/{case_name}.adb'
with open(adb_new_path,'w') as file:
    file.write(adb_string)

#Edit MESSAGE directories
# Copy new folder to MESSAGE directory
MESSAGE_mod_fd = f"{MESSAGE_root_fd}{case_name}"
os.makedirs(MESSAGE_mod_fd)
shutil.copytree(output_fd, MESSAGE_mod_fd, dirs_exist_ok=True)

# Add line to mms.pro
#Only use the following if creating model for the first time
if switch ==1:
    MESSAGE_mms_fils = f"{MESSAGE_root_fd}/mms_fils/"

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

    with open(f"{MESSAGE_mms_fils}mms.pro", "a") as f1:
        f1.write(f"\n{case_name}.dirfile   $MMS_HOME/mms_fils/{case_name}.dir")
    with open(f"{MESSAGE_mms_fils}glob.reg", "a") as f2:
        f2.write(f"\n{case_name}	{case_name}	{case_name}	+	empty")
    dir_f = f"{MESSAGE_mms_fils}{case_name}.dir"
    with open(dir_f, "w") as f3:
        f3.write(dir_s)
