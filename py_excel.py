# Author : Sujit Mandal
import os
import argparse
import openpyxl # pip install openpyxl
import pandas as pd # pip install pandas
from SystemSlash.SystemSlash import slash # pip install system-slash

system = slash()

# Github: https://github.com/sujitmandal
# Pypi : https://pypi.org/user/sujitmandal/
# LinkedIn : https://www.linkedin.com/in/sujit-mandal-91215013a/

def xls_xls(file_path):
    file_name = pd.ExcelFile(file_path)
    
    for sheet in file_name.sheet_names:
        df = pd.read_excel(file_name,sheet_name=sheet)
        df.to_excel('save_excel' + system + sheet + '.xls', index=False)


def xlsx_xls(file_path):
    file_name = openpyxl.load_workbook(file_path)

    for sheet in file_name.sheetnames:
        df = pd.read_excel(file_path,sheet_name=sheet)
        df.to_excel('save_excel' + system + sheet + '.xls', index=False)
      


def main():
    my_parser = argparse.ArgumentParser(description='Excel Conversion')
    my_parser.add_argument('PATH', help='Enter Excel Path')

    args = my_parser.parse_args()
    excel_path = args.PATH

    save_excel = 'save_excel'

    try:
        if os.path.exists(excel_path) and not os.path.exists(save_excel):
            os.mkdir(save_excel)

            if excel_path[-4:] == '.xls':
                xls_xls(excel_path)

            if excel_path[-5:] == '.xlsx':
                xlsx_xls(excel_path)  


            print('\n')
            print('Directory Creation Completed.')
            print('Directory : {}'.format(os.getcwd()) + system + save_excel)

        elif os.path.exists(save_excel):
            print('\n')
            print('Same Directory Exists On This Directory..')
            print('Directory : {}'.format(os.getcwd()) + system + save_excel)

        elif not os.path.exists(excel_path):
            print('\n')
            print(excel_path + ' File Dose Not Exists')

    except OSError:
        print('Files Dose Not Exists')




if __name__ == "__main__":
    main()