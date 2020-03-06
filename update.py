import win32com.client
import sys, os, time
import os
from threading import Thread

def update_PQ(wb):
    for x in wb.Connections:
        x.OLEDBConnection.BackgroundQuery = False
    wb.RefreshAll()
    return True

def fileUpdate(xl, xlsFileName):
    wb = xl.Workbooks.Open(xlsFileName)
    xl.Visible = True
    xl.DisplayAlerts = False
    # update_PQ(wb)
    x = Thread(target=update_PQ, args=(wb,))
    x.start()
    x.join()
    time.sleep(10)
    print('yo')


def worker():
    xl = win32com.client.DispatchEx("Excel.Application")
    file = 'C:\\work\\findcolor\\in\\231100043 - ТСЦ-400Ц-22м.xlsx'
    fileUpdate(xl, file)
    xl.Quit()
    del xl
    print('Все файлы обработаны. Программа закроется автоматически через 3сек.')
    print('Powered by Yegor Kowalew')

if __name__ == "__main__":
    worker()