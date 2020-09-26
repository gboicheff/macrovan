#! python3
# read_pdf.py - finds List numbers from all pdf's in a folder
# Gets number from third page by finding string following substring "List"
# plagiarized from https://automatetheboringstuff.com/chapter13/

import PyPDF2
import os
import pandas as pd
from utils import *
from secrets import *

if __name__ == '__main__':
    # Loop through all the PDF files.
    #path = r'D:\Stuff\Projects\Pol\macrovan\io\Output'
    #path = r'D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\PDFExports_20200901125934_35_files'
    #path = r'D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\PDFExports_20200919022945_20_files'
    #path = r'D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\PDFExports_20200921142023_357_files'
    path = r'D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\VBM files'
    # D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\VBM files
    pdf_files = get_fnames(path)

    pdf_dict = extract_pdf_info(path)
    pdf_dict_keys = pdf_dict.keys()
    print('pdf_dict_keys = ', pdf_dict_keys)

    organizer_dict = get_organizer_turfs_dict()

    # for keys in organizer_dict.keys:
    #     for

    df = pd.DataFrame(pdf_dict).transpose()
    write_path = r'..\io\Output\PDFs\VBM files\List Numbers.xlsx'
    write_excel(write_path, df)
