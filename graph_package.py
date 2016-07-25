import calendar
import itertools
import os
import shutil
import sys
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

plt.style.use('ggplot')


def flip(items, ncol):
    return itertools.chain(*[items[i::ncol] for i in range(ncol)])

"""Make this take a lab name from PYQT since the dictionary is already there
    and add the Lab name to top of graph as well"""


def make_graph(machine_name, data_csv):
    data = pd.read_csv(data_csv, encoding='utf-8-sig', low_memory=False)
    df = pd.DataFrame(data)
    # pd.set_option('display.float_format', lambda x: '%.2f' % x)
    # df['Date'] = pd.to_datetime(df['Date'])
    # , format='%m-%d-%y').dt.strftime('%m-%d-%y')

    df.sort_values(['Component Name', 'Sample Name'], inplace=True)
##    df.drop_duplicates(['Component Name', 'Sample Name', 'Date'], inplace=True)
    graph_folder = os.path.join(os.path.dirname(data_csv), 'Graphs')

    if os.path.exists(graph_folder):
        shutil.rmtree(graph_folder)
    os.makedirs(graph_folder)

    df.dropna(inplace=True)
    
    pdf_to_save = PdfPages(os.path.join(graph_folder, '{machine_name} May-June Graphs.pdf'.format(machine_name=machine_name)))
    grouped = df.groupby(['Component Name', 'Sample Name'])

    for i in grouped.size().index:
        drug_name = i[0]
        QC = i[1]
        drug_group = grouped.get_group(i)
        y_area = drug_group['Area']
        y_area_is = drug_group['IS Area']
        y_area_ratio = drug_group['Area Ratio']
        x = drug_group['Date']
        x_range = len(x)
        x_values = np.arange(1, x_range+1)
        x_range_full_line = x_range + 2
        x_values_full_line = np.arange(x_range_full_line)

        """Begin to plot Area"""
        y_mean = np.mean(y_area)
        y_mean_values = np.empty(x_range_full_line)
        y_mean_values.fill(y_mean)
        y_neg_20 = y_mean - (y_mean * .2)
        y_pos_20 = y_mean + (y_mean * .2)

        y_neg_values = np.empty(x_range_full_line)
        y_neg_values.fill(y_neg_20)

        y_pos_values = np.empty(x_range_full_line)
        y_pos_values.fill(y_pos_20)

        plt.figure(figsize=(14,8))
        plt.xlim(np.amin(x_values)-.5, np.amax(x_values)+.5)
        plt.xticks(x_values, x, rotation=80)
        plt.tick_params(axis='x',labelsize=10)
        plt.title('Area {machine_name} {drug_name}'.format(machine_name=machine_name, drug_name=drug_name[:-2]))
        plt.ylabel('Area', fontsize=10)
        
        """Begin to plot mean and sd values"""
        plt.plot(x_values_full_line, y_mean_values, 'gold', label='Mean')
        plt.text(x_values_full_line[-1], y_mean_values[-1], 'Mean')
        
        plt.plot(x_values_full_line, y_neg_values, 'green', label='-20%')
        plt.text(x_values_full_line[-1], y_neg_values[-1], '-20%')
        
        plt.plot(x_values_full_line, y_pos_values, 'orange', label='+20%')
        plt.text(x_values_full_line[-1], y_pos_values[-1], '+20%')
        
        plt.plot(x_values, y_area, color='blue', marker='o')
        plt.legend(ncol=6, fontsize=9,loc='upper center')
        pdf_to_save.savefig()
        plt.close()

        """Being to plot IS Area"""

        y_mean = np.mean(y_area_is)
        y_mean_values = np.empty(x_range_full_line)
        y_mean_values.fill(y_mean)
        y_neg_20 = y_mean - (y_mean * .2)
        y_pos_20 = y_mean + (y_mean * .2)

        y_neg_values = np.empty(x_range_full_line)
        y_neg_values.fill(y_neg_20)

        y_pos_values = np.empty(x_range_full_line)
        y_pos_values.fill(y_pos_20)

        plt.figure(figsize=(14,8))
        plt.xlim(np.amin(x_values)-.5, np.amax(x_values)+.5)
        plt.xticks(x_values, x, rotation=80)
        plt.tick_params(axis='x',labelsize=10)
        plt.title('IS Area {machine_name} {drug_name}'.format(machine_name=machine_name, drug_name=drug_name[:-2]))
        plt.ylabel('Area', fontsize=10)
        
        """Begin to plot mean and sd values"""
        plt.plot(x_values_full_line, y_mean_values, 'gold', label='Mean')
        plt.text(x_values_full_line[-1], y_mean_values[-1], 'Mean')
        
        plt.plot(x_values_full_line, y_neg_values, 'green', label='-20%')
        plt.text(x_values_full_line[-1], y_neg_values[-1], '-20%')
        
        plt.plot(x_values_full_line, y_pos_values, 'orange', label='+20%')
        plt.text(x_values_full_line[-1], y_pos_values[-1], '+20%')
        
        plt.plot(x_values, y_area_is, color='red', marker='o')
        plt.legend(ncol=6, fontsize=9,loc='upper center')
        pdf_to_save.savefig()
        plt.close()

        """Begin to plot Area Ratio"""

        y_mean = np.mean(y_area_ratio)
        y_mean_values = np.empty(x_range_full_line)
        y_mean_values.fill(y_mean)
        y_neg_20 = y_mean - (y_mean * .2)
        y_pos_20 = y_mean + (y_mean * .2)

        y_neg_values = np.empty(x_range_full_line)
        y_neg_values.fill(y_neg_20)

        y_pos_values = np.empty(x_range_full_line)
        y_pos_values.fill(y_pos_20)

        plt.figure(figsize=(14,8))
        plt.xlim(np.amin(x_values)-.5, np.amax(x_values)+.5)
        plt.xticks(x_values, x, rotation=80)
        plt.tick_params(axis='x',labelsize=10)
        plt.title('Area Ratio {machine_name} {drug_name}'.format(machine_name=machine_name, drug_name=drug_name[:-2]))
        plt.ylabel('Area/IS', fontsize=10)
        
        """Begin to plot mean and sd values"""
        plt.plot(x_values_full_line, y_mean_values, 'gold', label='Mean')
        plt.text(x_values_full_line[-1], y_mean_values[-1], 'Mean')
        
        plt.plot(x_values_full_line, y_neg_values, 'green', label='-20%')
        plt.text(x_values_full_line[-1], y_neg_values[-1], '-20%')
        
        plt.plot(x_values_full_line, y_pos_values, 'orange', label='+20%')
        plt.text(x_values_full_line[-1], y_pos_values[-1], '+20%')
        
        plt.plot(x_values, y_area_ratio, color='red', marker='o')
        plt.legend(ncol=6, fontsize=9,loc='upper center')
        pdf_to_save.savefig()
        plt.close()
        
    
        
    pdf_to_save.close()

def validate_data_csv(data_csv):
    if data_csv.endswith('data.csv'):
        return True

make_graph('MPX2', r'\\192.168.0.242\profiles$\massspec\Desktop\ChartsTest\Results\Machine Results\MS-MPX2\data.csv')
print('Done MPX2')
                           
make_graph('MPX3', r'\\192.168.0.242\profiles$\massspec\Desktop\ChartsTest\Results\Machine Results\MS-MPX3\data.csv')
print('Done MPX3')
                           
make_graph('MPX4', r'\\192.168.0.242\profiles$\massspec\Desktop\ChartsTest\Results\Machine Results\MS-MPX4\data.csv')
print('Done MPX4')
                           
make_graph('MPX5', r'\\192.168.0.242\profiles$\massspec\Desktop\ChartsTest\Results\Machine Results\MS-MPX5\data.csv')
print('Done MPX5')
