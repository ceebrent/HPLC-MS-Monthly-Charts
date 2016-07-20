import calendar
import csv
import datetime
import errno
import glob
import os
import re
import shutil
from pathlib import Path
import sys
import pandas as pd

##from home_directory import home_folder


class LeveyJennings(object):
    def __init__(self):
        # replace with get_home() when done

        self.homeDirectory = r'D:\Coding\Python\TestFiles'

    # Creates and returns folder to store results into
        def results_folder(self):
            # move_directory_up = Path(self.homeDirectory).parents[0]
            
            result_folder = os.path.join(self.homeDirectory, 'Results')
            os.makedirs(result_folder, exist_ok=True)
            
            machine_results = os.path.join(result_folder, 'Machine Results')

            os.makedirs(machine_results, exist_ok=True)
            return machine_results
        self.machine_results = results_folder(self)

    # Walks through all files and directories in home folder with ending text and containing lab name
        def original_txt(self):
            lab_text_files = []
            for dirpath, dirnames, files in os.walk(os.path.join(self.homeDirectory, 'Data')):
                for filename in files:
                    if filename.endswith('.txt'):
                        lab_text_files.append(os.path.join(dirpath, filename))
            return lab_text_files
        self.text_files = original_txt(self)

    def make_unique_files(self):
        # Copies all appropriate files into result folder and takes newest version of file
        new_result_path = os.path.join(self.machine_results, 'Temp_TXT')
        os.makedirs(new_result_path, exist_ok=True)

        for x in range(len(self.text_files)):
            # Put original files into TXT folder, check for duplicates
            file_name = os.path.join(new_result_path, os.path.basename(self.text_files[x]))
            silent_remove(file_name)
            shutil.copy2(self.text_files[x], new_result_path)

            # Create new name in "Date Method Batch Lab format
            date = get_date_machine(self.text_files[x])
            old_file_name = file_name_regex(os.path.basename(self.text_files[x]))
            unique_file_name = "{date} {old_file}.txt".format(date=date, old_file=old_file_name)

            unique_path = os.path.join(self.machine_results, os.path.basename(unique_file_name))
            silent_remove(unique_path)
            os.rename(file_name, unique_path)


def make_month_folders(result_path):
    list_of_files = glob.glob(result_path+'\\*.txt')
    for file in list_of_files:
        file_name = os.path.basename(file)
        machine_name = file_name[:4]

        machine_folder = os.path.join(result_path, machine_name)
        os.makedirs(machine_folder, exist_ok=True)
        # Move into months folders
        try:
            shutil.move(file, machine_folder)
        # If file exists in month folder, delete it and add new
        except shutil.Error:
            os.remove(os.path.join(machine_folder, os.path.basename(file)))
            shutil.move(file, machine_folder)


# Gets date file was created from cell in original text file
def get_date_machine(text_file):
    with open(text_file, 'r') as original_file:
        row = original_file.readlines()[1].split('\t')
        date_text = os.path.basename(row[2])[:8]
        machine_row = os.path.normpath(row[2]).split(os.sep)
        machine_name = ''
        if 'Projects' in machine_row:
            project_index = machine_row.index('Projects')
            machine_path = machine_row[project_index + 1]
            machine_name = machine_path[:4]
        name_date = '{machine_name} {date}'.format(machine_name=machine_name, date=date_text)
        return name_date


# Gets Unique portion of original file name to append to date file
def file_name_regex(base_name):
    base_name = base_name.strip()
    print(base_name)
    return re.findall('^[A-Z]+\s*\w+-[0-9]+', base_name)[0]


def silent_remove(filename):
    try:
        os.remove(filename)
    except OSError as e:  # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
            raise  # re-raise exception if a different error occurred


def merge_txt_to_csv(path_to_directory):
    """ Takes list of files from directory, makes file new each time and sets designated header values"""
    list_of_files = glob.glob(os.path.join(path_to_directory, '*.txt'))
    out_csv_file = os.path.join(path_to_directory, 'data.csv')
    silent_remove(out_csv_file)
    out_csv = csv.writer(open(out_csv_file, 'a', newline=''))
    headers_in_file = list(csv.reader(open(list_of_files[0], 'rt'), delimiter='\t'))
    # print(headers_in_file)
    # original_filename = headers_in_file[0][headers_in_file[0].index('Original Filename')]
    sample_name = headers_in_file[0][headers_in_file[0].index('Sample Name')]
    component_name = headers_in_file[0][headers_in_file[0].index('Component Name')]
    concentration = headers_in_file[0][headers_in_file[0].index('Calculated Concentration')]
    headers = [component_name, sample_name, concentration, 'Date']
    out_csv.writerow(headers)
    # out_csv_opened = csv.writer(open(out_csv_file, 'a', newline=''))
    for files in list_of_files:
        in_file = list(csv.reader(open(files, 'rt'), delimiter='\t'))
        for rows in in_file:
            if rows[in_file[0].index('Sample Name')] in ('Low QC', 'HIgh QC') and \
                    rows[in_file[0].index('Component Name')].endswith('1'):
                component_name = rows[in_file[0].index('Component Name')]
                sample_name = rows[in_file[0].index('Sample Name')]
                concentration = rows[in_file[0].index('Calculated Concentration')]
                date_name = (os.path.basename(rows[in_file[0].index('Original Filename')])[:8])
                date_formatted = str(pd.to_datetime(date_name, format='%Y%m%d').date())
                date_final = datetime.datetime.strptime(date_formatted, '%Y-%m-%d').strftime('%m-%d-%y')
                values = [component_name, sample_name, concentration, date_final]
                out_csv.writerow(values)


def walk_months(results_directory):
    directories = next(os.walk(results_directory))[1]

    for subdirectories in directories:
        month_directory = os.path.join(results_directory, subdirectories)
        if os.listdir(month_directory) != []:
            merge_txt_to_csv(month_directory)


def get_home():
    if hasattr(sys, 'frozen'):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(sys.argv[0])


def generate_data():
    lab = LeveyJennings()
    lab.make_unique_files()
    make_month_folders(lab.machine_results)
    walk_months(lab.machine_results)


generate_data()