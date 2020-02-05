import win32com.client
import os.path, time, datetime
NEED_OPERATION = ''

def remove_color_rows(wb):
    def remove_rows(color_index):
        mass = []
        zz = []
        for i in range(7, getNumRows):
            this_is = sheet.Cells(i, 2).Font.ColorIndex
            if this_is != color_index:
                if not zz:
                    zz.append(i)
                else:
                    if zz[-1] == i-1:
                        zz.append(i)
                    else:
                        mass.append(zz)
                        zz = []
        for i in reversed(mass):
            delstr = 'A%s:A%s' % (min(i), max(i))
            sheet.Range(delstr).EntireRow.Delete(Shift=color_index)

    sheet = wb.Worksheets('График')
    allData = sheet.UsedRange

    getNumRows = allData.Rows.Count
    print ('Строк до обработки: ',getNumRows)
    print('')
    colors = []
    for i in range(7, getNumRows):
        this_is = sheet.Cells(i, 2).Font.ColorIndex
        if this_is in colors:
            pass
        else:
            colors.append(this_is)
    for oper in colors:
        print(colors.index(oper), oper)
    print('Выберите цвет текста, который должен остаться:')
    answer = input()
    try:
        answer = int(answer)
        if colors[answer]:
            print('Вы выбрали цвет: %s' % colors[answer])
            remove_rows(int(colors[answer]))
            print('')
        else:
            print('Ошибка при выборе цвета. Выход')
    except BaseException as identifier:
        print('Ошибка при выборе цвета: %s. Выход' % (identifier))
        exit()

    getNumRows = allData.Rows.Count
    print ('Строк после обработки: ',getNumRows)

def update_file(wb):
    print('Начало обновления файла')
    for x in wb.Connections:
        x.OLEDBConnection.BackgroundQuery = False
    
    wb.RefreshAll()
    print('Конец обновления файла')

def update_operation(wb):
    # Техоперации совмещенные
    sheet = wb.Worksheets('Техоперации совмещенные')
    sheet.Visible = True
    sheet.Activate()

    wwt = sheet.ListObjects("ПоТехоперации")
    zz = wwt.listcolumns("Техоперация").DataBodyRange
    operlist = []
    for i in zz:
        if str(i) not in operlist:
            operlist.append(str(i))
    print('')
    print('Выберите операцию, по которой фильтровать таблицу:')
    for oper in operlist:
        print(operlist.index(oper), oper)
    
    answer = input()
    try:
        answer = int(answer)
        if operlist[answer]:
            print('Вы выбрали техоперацию: %s' % operlist[answer])
        else:
            print('Ошибка при выборе техоперации. Выход')
    except BaseException as identifier:
        print('Ошибка при выборе техоперации: %s. Выход' % (identifier))
        exit()
       
    sheet.Cells(6,10).WrapText = True
    sheet.ListObjects("ПоТехоперации").DataBodyRange.Font.Size = 14
    sheet.ListObjects("ПоТехоперации").Range.AutoFilter(Field=12, Criteria1=operlist[answer])

def save_as_new_book(wb):
    wb.SaveAs('C:\\work\\findcolor\\out\\test2.xlsx')
    wb.Close()

def fileUpdate(xl, xlsFileName):
    wb = xl.Workbooks.Open(xlsFileName)
    xl.Visible = True

    remove_color_rows(wb)
    update_file(wb)
    update_operation(wb)
    save_as_new_book(wb)
    del wb


def worker():
    xl = win32com.client.DispatchEx("Excel.Application")

    xlsFileName = 'C:\\work\\findcolor\\in\\test.xlsx'
    fileUpdate(xl, xlsFileName)
    # xl.Quit()
    # del xl
    print("Обновлен файл: %s \n" % (xlsFileName))

if __name__ == "__main__":
    worker()