import glob
import re
from openpyxl import load_workbook


def main():
    files = glob.glob("./inputs/*.xlsx")

    for file in files:
        print("#############################################")
        print(f'loading file {file}')
        print("#############################################")
        # Excelファイルの読み込み
        work_book = load_workbook(file, data_only=True)

        print("##############################################")
        print('now executing')
        print("##############################################")
        sheet_name_re = re.search(r'A\d+-T\d+', file)
        sheet_name = sheet_name_re.group()

        for ws in work_book:
            if (ws.title == 'WAVEFORM'): ws.title = sheet_name 
        
        work_book.save(f'./outputs/{sheet_name}_renamed.xlsx')

if __name__ == '__main__':
    main()