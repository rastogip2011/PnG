from atfunc import *
import pandas as pd
import os
import datetime
import numpy as np
import os

def save_var(month, region):
    f = open("variables.py","w") 
    L = [
        "# Step 1 output:\n",
        "output_pivot_table = 'Output/"+month+'/PivotTable_'+month+".xlsx' \n",
        "output_parcel_cost = 'Output/"+month+'/ParcelCost_'+month+".xlsx' \n",
        "# Step 2 (optimizer) output:\n",
        "optimizer_output = 'Output/"+month+"/optimizer_output_"+month+".xlsx' \n",
        "sheet_BRS = 'BRS cost Daywise' \n",
        "sheet_VU = 'Vehicle Utilisation Daywise' \n"
        "# Files required as input:\n",
        "month = '"+month+"' \n",
        "region = '"+region+"' \n",
        "filename_master = 'Master.xlsx' \n",
        "sheet_cost = 'BRS' \n",
        "sheet_truck_cap = 'Truck' \n",
        "filename_summary = 'Summary/"+month+".xlsx' \n",
        "sheet_summary = 'Summary Sheet' \n",
        "summary_copy_path_local = 'Summary/"+month+'_'+region+".xlsx' \n",
        "summary_copy_path_drive = 'C:/Users/rastogi.l/OneDrive - Procter and Gamble/Desktop/Main/Power BI Input/"+month+'_'+region+".xlsx' \n", 
        ] 

    f.writelines(L) 
    f.close()


month = input('Enter Month: ')
region = input('Enter Region: ')
print('\nStep 1: Initiating process')
save_var(month, region)

parcel_sites = set()
Psheet = pd.read_excel('Master.xlsx','Parcel').values.ravel()
[parcel_sites.add(x) for x in Psheet];

parcel_locations = set()
site = pd.read_excel('Master.xlsx','Site').values
[parcel_locations.add(r[3].split(' ')[0].upper()) for r in site if r[0] in parcel_sites];

Master_SKU = pd.read_excel("Master.xlsx", 'Master SKU')
cm = Master_SKU.columns.tolist()
Master_SKU = Master_SKU[[cm[0],cm[2],cm[9],cm[10]]]
cm = Master_SKU.columns.tolist()

stn_path = 'STNs/'+month+'/' 
STN = pd.DataFrame([])
STN_parcel = pd.DataFrame([])
files = os.listdir(stn_path)


for f in files:
    
    stn = pd.read_excel(stn_path+f)
    c = stn.columns.tolist()
    stn = stn[[c[0],c[2],c[6],c[7],c[9],c[12]]]
    
    t=stn[['TOLOCATION']].values.ravel().tolist()
    t = [x.upper() for x in t]
    df = pd.DataFrame(t)
    stn[['TOLOCATION']]=df 
    
    isparcel = False
    for r in range(1,100,5):
    	if stn['TOLOCATION'][r].split(' ')[0] in parcel_locations:
    		isparcel = True

    if isparcel:
        stn = stn[['TOLOCATION','LocTfrDate','Case']]
        if len(STN_parcel.values.tolist())==0:
            STN_parcel = stn
        else:
            STN_parcel = STN_parcel.append(stn, ignore_index=True)
        
    else:
        stn = pd.merge(stn, Master_SKU[[cm[1],cm[2],cm[3]]],left_on='Product Name',right_on='Item_Des',how = 'left')
        stn = pd.merge(stn, Master_SKU[[cm[0],cm[2],cm[3]]],left_on='Product Code',right_on='Mas_Pcode', how = 'left')
        stn['Weight'] = np.where(stn['Weight_x'].isna(), stn['Weight_y'], stn['Weight_x'])
        stn['Volume'] = np.where(stn['Volume_x'].isna(), stn['Volume_y'], stn['Volume_x'])
        stn = stn.dropna(subset=['Weight', 'Volume'])

        stn["Weight"] = pd.to_numeric(stn['Weight'], errors='coerce')
        stn["Volume"] = pd.to_numeric(stn['Volume'], errors='coerce')
        stn["Case"] = pd.to_numeric(stn['Case'], errors='coerce')
        stn = stn.dropna(subset=['Weight', 'Volume', 'Case'])

        stn = stn.assign(Weight_kg = stn.Weight * stn.Case)
        stn = stn.assign(Volume_m3 = stn.Volume * stn.Case /1000)

        c = stn.columns.tolist()
        stn = stn[[c[0],c[1],c[5],c[14],c[15]]]

        if len(STN.values.tolist())==0:
            STN = stn
        else:
            STN = STN.append(stn, ignore_index=True)

make_pivot_table(STN, month)
make_parcel_cost_table(STN_parcel, month)    


os.system('python optimizer.py')
os.system('python summary.py')
