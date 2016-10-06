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
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Style, Font
from tkinter import *
from tkinter import filedialog
import glob
import csv


plt.style.use('ggplot')

def silent_remove(filename):
    try:
        os.remove(filename)
    except OSError as e:  # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
            raise  # re-raise exception if a different error occurred

def text_to_df(path_to_folder):
    batch_folder = os.path.dirname(path_to_folder)
    list_of_files = glob.glob(os.path.join(path_to_folder, '*.txt'))

    
    ##For each csv file in list of files we append them to the end
    df = pd.concat((pd.read_csv(f, delimiter='\t') for f in list_of_files))

    df.replace(['BENZOYLECGONINE-COCAINE METABOLITE 1', 'JWH-073-N-Butanoic acid/JWH-018-5OHpentyl 1'],
               ['BENZOYLECGONINE-COCAINE 1', 'JWH-073-N-Butanoic/JWH-018 1'], inplace=True)
    df['Component Name'] = df['Component Name'].str.replace('/', '-')
    df['Calculated Concentration'] = pd.to_numeric(df['Calculated Concentration'], errors='coerce').round(2)        
    columns_to_keep = ['Sample Name', 'Sample Type', 'Component Name', 'Actual Concentration',
                       'Calculated Concentration', 'Inst.']
    return df[columns_to_keep]

def group_sort(df, path_to_folder):
    df = df[df['Component Name'].str.endswith('1')]
    df.sort_values(['Component Name', 'Sample Name', 'Inst.'], inplace=True)
    df.reset_index(inplace=True, drop=True)
    length = df.shape[0]
    df_name_modify = df
##    print(length)
##    rows = df.iterrows()
##    for x in range(length):
##        row = next(rows)
##        print(row)
##        sample_type = row[1]['Sample Type']
##        sample_name = row[1]['Sample Name']
##        if sample_type == 'Unknown':
##            df_name_modify.loc[x, 'Sample Name'] = sample_name[-2:]
            
    csv_file = os.path.join(path_to_folder, 'data.csv')
    new_csv_file = os.path.join(path_to_folder, 'data_m.csv')
    silent_remove(new_csv_file)
    silent_remove(csv_file)
    df.to_csv(csv_file)

    with open(csv_file, 'r') as infile, open(new_csv_file, 'a', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)
        header = next(reader)
        sample_type = header.index('Sample Type')
        sample_name = header.index('Sample Name')
        writer.writerow(header)
        for row in reader:
            if row[sample_type] == 'Unknown':
                row[sample_name] = int(row[sample_name][-2:].strip(' '))
                writer.writerow(row)
            else:
                writer.writerow(row)
    return df['Component Name'].unique(), df['Inst.'].unique(), new_csv_file

def write_styles(csv_file, path_to_folder, drug_names, instruments):
    out_xlsx = os.path.join(path_to_folder, 'output.xlsx')
    silent_remove(out_xlsx)
    wb = openpyxl.Workbook()
    for drugs in range(len(drug_names)):
        wb.create_sheet(index=drugs, title=drug_names[drugs])

    drug_sheets = wb.get_sheet_names()
    num_of_patients = 25


    """Borders formatting"""
    thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))
    medium_border = Border(left=Side(style='medium'), 
                     right=Side(style='medium'), 
                     top=Side(style='medium'), 
                     bottom=Side(style='medium'))
    
    with open(csv_file, 'r') as infile:
        reader = csv.reader(infile)
        reader_header = next(reader)
        header = ['Sample Name', instruments[0], instruments[1], '% Diff.']
        for sheet in drug_sheets:
            ws = wb.get_sheet_by_name(sheet)
##          Write and style title name cell 
            wb[sheet].append([sheet])
            ws.cell(row=1, column=1).border = medium_border
            ws.cell(row=1, column=1).font = Font(bold=True, size=18)

##          Write and style header  
            wb[sheet].append(header)
            for col in range(1,len(header)+1):
                ws.cell(row=2, column=col).border = thin_border
                ws.cell(row=2, column=col).font = Font(bold=True)

##           Begin to write Sample names and machine calc concs.

        
            for sample in range(1, num_of_patients+1):
                ws.cell(row=sample+2, column=1).value = sample
                sample_name = reader_header.index('Sample Name')
                inst_data = reader_header.index('Inst.')
                calc_conc = reader_header.index('Calculated Concentration')
                comp_name = reader_header.index('Component Name')
                inst1_value = 0
                inst2_value = 0
                
                for row in reader:
                    try:
                        if (int(row[sample_name]) == sample and row[comp_name] == sheet):
                            if row[inst_data] == instruments[0]:
                                inst1_value = float(row[calc_conc])
                                ws.cell(row = sample+2, column=2).value = inst1_value
                                
                            elif row[inst_data] == instruments[1]:
                                inst2_value = float(row[calc_conc])
                                ws.cell(row = sample+2, column=3).value = inst2_value
                                
                            else:
                                pass
                    except ValueError:
                        pass
                try:
                    ws.cell(row = sample + 2, column=4).value = round((100 * abs(inst1_value - inst2_value)/inst2_value), 2)
                except ZeroDivisionError:
                    ws.cell(row = sample + 2, column=4).value = 0
                infile.seek(1)
    wb.save(out_xlsx)


def make_file_dialog():
    root = Tk()
    root.withdraw()
    root.file_name = filedialog.askdirectory(parent=root,
                                             title='Choose a Comparison folder')

    if not root.file_name:
        sys.exit(0)
    print(text_to_df(root.file_name).head())

##if __name__ == "__main__":
##    make_file_dialog()
df = text_to_df(r'\\192.168.0.242\profiles$\massspec\Desktop\MPX2 NEW PP VALIDATION\Data\COMPARISON')
drug_names, instruments, csv_file = group_sort(df, r'\\192.168.0.242\profiles$\massspec\Desktop\MPX2 NEW PP VALIDATION\Data\COMPARISON')
write_styles(csv_file, r'\\192.168.0.242\profiles$\massspec\Desktop\MPX2 NEW PP VALIDATION\Data\COMPARISON',
            drug_names, instruments)
