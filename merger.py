import math

import pandas as pd
import numpy as np
import shutil
import os
import json
from datetime import datetime as dt

with open('BACKUP/fields.json') as f:
    data = json.load(f)
try:
    shutil.copy(data['files']['calculation_file'], 'BACKUP')
except:
    pass
columns=data['files']['COLUMNS']

def file_merge():
    f=open(date_stored,'w+')
    f.write(last_modified)
    f.close()
    try:
        df_accounting = pd.read_excel(data['files']['accounting_file'],header=None,dtype=str)
        df_trucking = pd.read_excel(data['files']['trucking_file'], sheet_name=data['files']['sheet_name'])
        df_calculation = pd.read_excel(data['files']['calculation_file'],usecols=columns,dtype=str)
    except Exception as e:
        pd.DataFrame([],columns=columns).to_excel(data['files']['calculation_file'],index=False)
        df_calculation = pd.read_excel(data['files']['calculation_file'],usecols=columns)
    list_all_rows=list()
    data_acc = data['account']
    data_truck = data['truck']
    start = data_acc['START']
    empty_count=0
    max_size=df_accounting.shape[0]
    while start<=max_size:
        try:
            list_row=[]
            data_row=df_accounting.loc[start]
            Truck_no= data_row.iat[data_acc['TRUCK']]
            if data_row.iat[data_acc['END']]=='On Account of :' or data_row.iat[data_acc['END']]=='To':
                break
            if data_row.empty:
                empty_count+=1
                if empty_count>3:
                    break

            if Truck_no is np.nan:
                start+=1
                continue
            data_truck_row=df_trucking.loc[df_trucking[data_truck['TRUCK']]== Truck_no]
            if data_truck_row.empty:
                start+=1
                continue
            list_row.append(df_accounting.iat[data_acc['VCH.NO'][1],data_acc['VCH.NO'][0]])
            list_row.append(df_accounting.iat[data_acc['REF'][1],data_acc['REF'][0]])
            date_vch=df_accounting.iat[data_acc['VCH.DATE'][1],data_acc['VCH.DATE'][0]]
            list_row.append(date_vch)
            list_row.append(data_row.iat[data_acc['CHALLAN']])
            list_row.append(data_row.iat[data_acc['DATE']])
            list_row.append(Truck_no)
            list_row.append(data_truck_row[data_truck['NAME']].values[0])
            list_row.append(data_truck_row[data_truck['AC.NO']].values[0])
            list_row.append(data_truck_row[data_truck['IFSC']].values[0])
            list_row.append(data_truck_row[data_truck['BANK']].values[0])
            list_row.append(data_row.iat[data_acc['AMOUNT']])
            list_row.append(data_row.iat[data_acc['DEST']])
            list_row.append(data_row.iat[data_acc['QTY']])
            list_row.append(data_truck_row[data_truck['MOB.NO']].values[0])
            list_all_rows.append(list_row)
        except Exception as e:
            print(e)
        start+=1

    final=[""]*14
    final[9]="TOTAL :"
    final[10]=df_accounting.iat[data_acc['TOTAL'][1],data_acc['TOTAL'][0]]
    now=dt.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    empt=[""]*14
    empt[0]='-'
    list_all_rows.append(empt)
    final[2]="CREATED : "+dt_string
    list_all_rows.append(final)
    list_all_rows.append(empt)
    list_all_rows.append(empt)
    new_df=pd.DataFrame(list_all_rows,columns=columns).astype(str)
    df_row = pd.concat([df_calculation, new_df])
    df_row.to_excel(data['files']['calculation_file'],index=False)












date_stored='BACKUP/last.txt'



f=open(date_stored,'r')
stored=f.readline()
last_modified=str(os.path.getmtime(data['files']['accounting_file']))
if last_modified==stored:
    print("ACCOUNTING File was not modified since last run. Do you Still want to continue....")
    reply=  input("Press Y if you want to continue or any other key to exit   ")
    if(reply=='Y' or reply=='y'):
        file_merge()
else:
    file_merge()

