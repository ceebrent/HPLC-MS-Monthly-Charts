from tkinter import *
from tkinter import filedialog
import os
from graph_package import make_chart
from data_main import merge_txt_to_csv

def make_file_dialog():
    root = Tk()
    root.withdraw()
    root.file_name = filedialog.askdirectory(parent=root, title='Choose a TXT folder')
    data_file = os.path.join(root.file_name, 'data.csv')
    
    if not root.file_name:
        sys.exit(0)

    merge_txt_to_csv(root.file_name)
    make_chart(data_file)

if __name__ == "__main__":
    make_file_dialog()
