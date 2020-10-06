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

    # path = r'D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\VBM files'
    path = r'C:\Users\Grant\Desktop\macrovan\io\Output\PDFS\VBM files'
    # D:\Stuff\Projects\Pol\macrovan\io\Output\PDFs\VBM files
    pdf_files = get_fnames(path)
    list_dict = extract_list_info(path)
    df = pd.DataFrame(list_dict).transpose()
    # write_path = r'..\io\Output\PDFs\VBM files\List Numbers.xlsx'
    write_path = r'C:\Users\Grant\Desktop\macrovan\io\Output\PDFS\VBM files\List Numbers.xlsx'
    #pdf_file_names = get_fnames(path)
    #print('pdf_file_names = ', pdf_file_names)
    # P123 -vbm Turf 26 Hermitage.pdf

    pdf_dict = extract_pdf_info(path)  # PDF INTERNAL DATA
    # pdf_dict_keys = pdf_dict.keys()  # key is PDF FILE NAME!
    #print('pdf_dict_keys = ', pdf_dict_keys)

    organizer_dict = get_organizer_turfs_dict()  # key is TURF NAME IN VAN!
    # organizer_dict_keys = organizer_dict.keys()

    # Not finding 'P123 -vbm Turf 26 Hermitage' and a few others
    # P123 -vbm Turf 26
    for turf_name_in_van in organizer_dict.keys():
        #print(turf_name_in_van)
        for pdf_file_name in pdf_dict.keys():
            if turf_name_in_van in pdf_file_name:
                pdf_dict[pdf_file_name]['organizer_email'] = organizer_dict[turf_name_in_van]
            #if 'Hermitage' in pdf_file_name:
            # if 'P123 -vbm Turf 26' in turf_name_in_van:
            #     print('Found P123 -vbm Turf 26')
            #     print(turf_name_in_van, ' , ', pdf_file_name)
            #     #print(organizer_dict[turf_name_in_van])
            #     pdf_dict[pdf_file_name]['organizer_email'] = organizer_dict[turf_name_in_van]
            #     print(pdf_dict[pdf_file_name])
            #
            #


    df = pd.DataFrame(pdf_dict).transpose()
    write_path = r'..\io\Output\PDFs\VBM files\List Numbers.xlsx'
    write_excel(write_path, df)
