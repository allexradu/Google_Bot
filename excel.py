import pandas as pd
import platform
from openpyxl import load_workbook

# downloaded_images = {'Image 1': [''], 'Image 2': [''], 'Image 3': [''], 'Image 4': [''],
#                      'Image 5': [''], 'Image 6': [''], 'Image 7': [''], 'Image 8': [''], 'Image 9': ['']}


def read_excel_first_column():
    # Reading first column of a local excel file
    try:
        if platform.system() == 'Windows':
            df = pd.read_excel('excel\\a.xlsx', sheet_name = 0)
        else:
            df = pd.read_excel('../excel/a.xlsx', sheet_name = 0)

        print('Excel Read Complete!')

        product_names = df['Name'].tolist()

        for i in range(len(product_names)):
            print(product_names[i])
        return product_names
    except:
        print('Excel File NOT READ. Name your file "a.xls" with the first column "Name"')
        print('and place it in the same directory and the bot file.')
