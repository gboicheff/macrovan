#! python3
# organizer_folders.py - uses excel input data to create folders
# for each organizer and copy their pdf's into them.

import PyPDF2
import os
import pandas as pd
from utils import *
from secrets import *

if __name__ == '__main__':
    print('------------- PROGRAM START -------------')
    fname = r"D:\Stuff\Projects\Pol\macrovan\io\Input\Nov 2020 -Tracking All Voters 20201015 02.xlsx"
    sheet_name = "To Deliver - Reports"

    create_organizer_folders(fname, sheet_name)
