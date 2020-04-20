import pandas as pd


def read_excel_first_column():
    # Reading first column of a local excel file
    try:
        df = pd.read_excel('excel\\a.xlsx', sheet_name = 0)
        print('Excel Read Complete!')

        product_names = df['Name'].tolist()

        for i in range(len(product_names)):
            print(product_names[i])
        return product_names
    except:
        print('Excel File NOT READ. Name your file "a.xls" with the first column "Name"')
        print('and place it in the same directory and the bot file.')
