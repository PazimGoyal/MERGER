import math

import pandas as pd
import numpy as np
import shutil
import os
import json
from datetime import datetime as dt
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import DimensionHolder, ColumnDimension

with open('BACKUP/fields.json') as f:
    data = json.load(f)
try:
    shutil.copy(data['files']['calculation_file'], 'BACKUP')
except:
    pass
columns=data['files']['COLUMNS']

def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False,set_width=False,
                       **to_excel_kwargs):
    # Excel file doesn't exist - saving and exiting
    if not os.path.isfile(filename):
        if 'header' in to_excel_kwargs:
            to_excel_kwargs.pop('header')

        df.to_excel(
            filename,
            sheet_name=sheet_name,
            startrow=startrow if startrow is not None else 0,
            **to_excel_kwargs)
        return


    writer = pd.ExcelWriter(filename, engine='openpyxl', mode='a')

    # try to open an existing workbook
    writer.book = load_workbook(filename)

    # get the last row in the existing Excel sheet
    # if it was not specified explicitly
    if startrow is None and sheet_name in writer.book.sheetnames:
        startrow = writer.book[sheet_name].max_row

    # truncate sheet
    if truncate_sheet and sheet_name in writer.book.sheetnames:
        # index of [sheet_name] sheet
        idx = writer.book.sheetnames.index(sheet_name)
        # remove [sheet_name]
        writer.book.remove(writer.book.worksheets[idx])
        # create an empty sheet [sheet_name] using old index
        writer.book.create_sheet(sheet_name, idx)

    # copy existing sheets
    writer.sheets = {ws.title: ws for ws in writer.book.worksheets}
    if set_width:
        size_dict={1:15,3:15,4:30,5:15,6:15,7:40,8:17,9:17,10:30,12:30}
        ws=writer.sheets["Sheet1"]
        dim_holder = DimensionHolder(worksheet=ws)
        for i,j in size_dict.items():
            dim_holder[get_column_letter(i)] = ColumnDimension(ws, min=i, max=i, width=j)
        ws.column_dimensions = dim_holder

    if startrow is None:
        startrow = 0

    # write out the new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, **to_excel_kwargs)

    # save the workbook
    writer.save()


def file_merge():
    f=open(date_stored,'w+')
    f.write(last_modified)
    f.close()
    try:
        df_accounting = pd.read_excel(data['files']['accounting_file'],header=None,dtype=str)
        df_trucking = pd.read_excel(data['files']['trucking_file'], sheet_name=data['files']['sheet_name'],dtype=str)
        # df_calculation = pd.read_excel(data['files']['calculation_file'],usecols=columns,dtype=str)
    except Exception as e:
        # pd.DataFrame([],columns=columns).to_excel(data['files']['calculation_file'],index=False)
        # df_calculation = pd.read_excel(data['files']['calculation_file'],usecols=columns)
        pass
    list_all_rows=list()
    data_acc = data['account']
    data_truck = data['truck']
    start = data_acc['START']
    empty_count=0
    max_size=df_accounting.shape[0]

    firms={'GKF':[],'HTC':[],'NEW':[],'OTHERS':[]}

    empty_truck_df=pd.DataFrame({'SR.NO':'', 'TRUCK NO':'', 'BANK ACCOUNT NO':'', 'IFSC CODE':'', 'BANK NAME':'',
       'OWNER NAME':'', 'MOBILE NO':''},index=[0]
)


    while start<=max_size:
        try:
            list_row=[]
            data_row=df_accounting.loc[start]
            Truck_no= data_row.iat[data_acc['TRUCK']]
            if data_row.iat[data_acc['END']]=='On Account of :' or data_row.iat[data_acc['END']]=='Ta':
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
                data_truck_row=empty_truck_df
            list_row.append(df_accounting.iat[data_acc['VCH.NO'][1],data_acc['VCH.NO'][0]])
            list_row.append(df_accounting.iat[data_acc['REF'][1],data_acc['REF'][0]])
            date_vch=df_accounting.iat[data_acc['VCH.DATE'][1],data_acc['VCH.DATE'][0]]
            list_row.append(date_vch)
            chalan=data_row.iat[data_acc['CHALLAN']]
            list_row.append(chalan)
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
            type=chalan.split('/')[2]
            firm=firms.get(type)
            if firm is None:
                firm = firms.get("OTHERS")
            firm.append(list_row)


        except Exception as e:
            pass
        start+=1

    final=[""]*14
    final[5]="Total Account VCH:"
    final[6]=df_accounting.iat[data_acc['TOTAL'][1],data_acc['TOTAL'][0]]
    final[9]="Total Calculated:"
    final[10]=''

    now=dt.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    empt=[""]*14
    empt[0]='-'
    list_all_rows.append(empt)
    final[2]="Last Modified : "+dt_string
    list_all_rows.append(final)
    list_all_rows.append(empt)
    list_all_rows.append(empt)
    new_df=pd.DataFrame(list_all_rows,columns=columns).astype(str)
    append_df_to_excel(data['files']['calculation_file'],new_df,set_width=True, header=None, index=False)


    # df_row = pd.concat([df_calculation, new_df])
    # writer = pd.ExcelWriter(data['files']['calculation_file'], engine='xlsxwriter')
    # df_row.to_excel(writer, sheet_name='Sheet1',index=False)

    # worksheet = writer.sheets['Sheet1']
    #
    # worksheet.set_column('A:N', 15)
    # worksheet.set_column('D:D', 30)
    # worksheet.set_column('G:G', 40)
    # worksheet.set_column('J:J', 30)
    # worksheet.set_column('L:L', 30)
    # writer.save()

    # writer = pd.ExcelWriter('final.xlsx')
    #
    # for i,j in firms.items():
    #     if j:
    #         new_df = pd.DataFrame(j, columns=columns).astype(str)
    #         new_df.to_excel(writer, sheet_name=i)
    #
    # writer.save()


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

"""
append_df_to_excel('d:/temp/test.xlsx', df)

append_df_to_excel('d:/temp/test.xlsx', df, header=None, index=False)

append_df_to_excel('d:/temp/test.xlsx', df, sheet_name='Sheet2',
                        index=False)

append_df_to_excel('d:/temp/test.xlsx', df, sheet_name='Sheet2',
                        index=False, startrow=25)
"""

