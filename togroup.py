import os, time
import glob
import win32com.client

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
in_folder = 'in'
out_folder = 'out'

searh_folder = '%s\\**\\*.xlsx' %  os.path.join(BASE_DIR, out_folder)
files_list = glob.glob(searh_folder, recursive=True)

def fileUpdate(xl, files_list):
    new_xlsx_file_path = os.path.join(BASE_DIR, 'Готовый.xlsx')
    neew_wb = xl.Workbooks.Add()
    to_sheet = neew_wb.Worksheets('Лист1')

    for xlsFileName in files_list:
        wb = xl.Workbooks.Open(xlsFileName)
        xl.Visible = True
        xl.DisplayAlerts = False
        sheet = wb.Worksheets('Техоперации совмещенные')
        sheet.Visible = True
        sheet.Activate()
        # sheet.Range("A1:A10").copy
        sheet.Range("A1:K5").Copy(to_sheet.Range("A7:K11"))
        time.sleep(10)

    neew_wb.SaveAs(new_xlsx_file_path)

def worker():
    xl = win32com.client.DispatchEx("Excel.Application")
    fileUpdate(xl, files_list)
    xl.Quit()
    del xl
    print('Все файлы обработаны. Программа закроется автоматически через 3сек.')
    print('Powered by Yegor Kowalew')

if __name__ == "__main__":
    worker()