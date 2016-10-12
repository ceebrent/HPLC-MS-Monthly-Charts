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

    sample_name_index = headers_in_file[0].index('Sample Name')
    component_index = headers_in_file[0].index('Component Name')
    concentration_index = headers_in_file[0].index('Calculated Concentration')

    
    sample_name = headers_in_file[0][sample_name_index]
    component_name = headers_in_file[0][component_index]
    concentration = headers_in_file[0][concentration_index]
    
    headers = [sample_name, component_name, concentration, 'Day']
    out_csv.writerow(headers)

    for files in list_of_files:
        in_file = list(csv.reader(open(files, 'rt'), delimiter='\t'))
        day = os.path.basename(os.path.splitext(files)[0])[-1]
        for rows in in_file:
             if rows[in_file[0].index('Sample Name')] in ('Low QC', 'HIgh QC') and \
                    rows[in_file[0].index('Component Name')].endswith('1'):
                component_name = rows[component_index]
                sample_name = rows[sample_name_index]
                concentration = rows[concentration_index]
                values = [sample_name, component_name, concentration, day]
                out_csv.writerow(values)

