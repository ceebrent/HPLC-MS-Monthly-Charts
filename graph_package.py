import calendar
import itertools
import os
import shutil
import sys
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import load_workbook
import errno
import xlsxwriter

plt.style.use('ggplot')


def flip(items, ncol):
    return itertools.chain(*[items[i::ncol] for i in range(ncol)])

"""Make this take a lab name from PYQT since the dictionary is already there
    and add the Lab name to top of graph as well"""

def silent_remove(filename):
    try:
        os.remove(filename)
    except OSError as e:  # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
            raise  # re-raise exception if a different error occurred

def make_chart(data_csv):
    data = pd.read_csv(data_csv, encoding='utf-8-sig', low_memory=False)
    df = pd.DataFrame(data)
    df.sort_values(['Component Name', 'Sample Name'], inplace=True)
    out_csv_file = os.path.join(os.path.dirname(data_csv), 'Precision_Validation.xlsx')                          
    silent_remove(out_csv_file)
    xlsxwriter.Workbook(out_csv_file)
    list_of_df = []
    grouped = df.groupby(['Component Name'])
    counter = 0
    
    for drug in grouped.size().index:
        dict_to_excel = {}
        drug_group = grouped.get_group(drug)
        number_of_days = list(set(drug_group['Day']))
        drug_name = drug[:-2]
##        drug_group['Calculated Concentration'] = drug_group['Calculated Concentration'].convert_objects(convert_numeric=True).round(2)
        drug_group['Calculated Concentration']= pd.to_numeric(drug_group['Calculated Concentration'], errors='coerce').round(2)
        if(drug_name == 'Modafinil'):
            break
        day_dfs = []
        high_qc_mean = drug_group[drug_group['Sample Name'] == 'HIgh QC']['Calculated Concentration'].mean()
        for days in number_of_days:
            out_dict = {}
            day = drug_group[drug_group['Day'] == days]
            headers = drug_group.dtypes.index
            day_header = 'Day {day}'.format(day=days)
            out_dict[(drug_name, day_header, 'Calc. Conc. (ng/mL)')] = day['Calculated Concentration']##.convert_objects(convert_numeric=True).round(2)
            out_dict[(drug_name, day_header, 'Sample Name')] = day['Sample Name']
            df_to_excel = pd.DataFrame.from_dict(out_dict)
            df_to_excel.reset_index(inplace=True, drop=True)
            df_to_excel = df_to_excel[[(drug_name, day_header, 'Sample Name'), (drug_name, day_header, 'Calc. Conc. (ng/mL)')]]
            df_to_excel.set_index((drug_name, day_header, 'Sample Name'), inplace=True)
            df_to_excel.index.rename('Sample Name', inplace=True)
            day_dfs.append(df_to_excel)

        result = day_dfs[0]
        for x in range(1, len(day_dfs)):
            result = pd.concat([result, day_dfs[x]], axis=1, join_axes=[result.index])


        book = load_workbook(out_csv_file)
        writer = pd.ExcelWriter(out_csv_file, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        result.to_excel(writer, startrow=counter, startcol=0)
        book.active['G1'] = 'Test'
        writer.save()
        counter = counter + 16

##    for x in list_of_df:
##        book = load_workbook(out_csv_file)
##        writer = pd.ExcelWriter(out_csv_file, engine='openpyxl')
##        writer.book = book
##        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
##        x.to_excel(writer, startrow=counter, startcol=0)
##        writer.save()
##        counter = counter + 16


make_chart(r'\\192.168.0.242\profiles$\massspec\Desktop\MPX2 NEW PP VALIDATION\Data\Precision\data.csv')
