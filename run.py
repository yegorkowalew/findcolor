import win32com.client
import sys, os, time
import os
clear = lambda: os.system('cls') #on Windows System
# clear()

from progress.bar import Bar, ShadyBar
from progress.spinner import Spinner

# spinner = Spinner('Обработка ')
# while state != 'FINISHED':
#     # Do some work
#     spinner.next()

color_base= {
-4105:'Цвет по умолчанию, по идее черный',
1:'Черный',
2:'Белый',
3:'Красный',
4:'Зеленый',
5:'Синий',
6:'Желтый',
7:'Розовый',
8:'Голубой',
9:'Коричневый',
10:'Темно-зеленый',
11:'Темно-синий',
12:'Темно-желтый',
13:'Темно-розовый',
14:'Зеленый',
15:'Серый',
16:'Темно-серый',
17:'Синевато-голубой',
18:'Фиолетовый',
19:'Светло-желтый',
20:'Светло-голубой',
21:'Темно-фиолетовый',
22:'Розоватый',
23:'Синевато-синий',
24:'Светло-голбоватый',
25:'Темно-темно-синий',
26:'Темновато-розовый',
27:'Темновато-желтый',
28:'Темновато-голубой',
29:'Темновато-розовый',
30:'Темновато-коричневый',
31:'Темновато-голубоватый',
32:'Синий',
33:'Голубой',
34:'Светло-голубой',
35:'Светло-салатовый',
36:'Светло-желтоватый',
37:'Светло-голубоватый',
38:'Светло-розоватый',
39:'Светло-фиолеватый',
40:'Светло-коричневатый',
41:'Ярко-синий',
42:'Ярко-голубоватый',
43:'Ярко-зеленоватый',
44:'Ярко-желтоватый',
45:'Ярко-желтоватый',
46:'Ярко-желтый',
47:'Темно-серо-синий',
48:'Серо-серый',
49:'Серо-синий',
50:'Серо-голубо-зеленый',
51:'Серо-зеленый',
52:'Серо-коричневый',
53:'Серо-оранжево-коричневй',
54:'Серо-розоватый',
55:'Серо-сине-голубоватый',
56:'Болотный',
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
in_folder = 'in'
out_folder = 'out'

def remove_color_rows(wb, xlsFileName):
    def remove_rows(color_index):
        mass = []
        zz = []
        # xlsFileName
        clear()
        print('Обработка файла: %s' % xlsFileName)
        barr = ShadyBar('Определяю цвета  ', max=getNumRows+1)
        for i in range(7, getNumRows+1):
            barr.next()
            this_is = sheet.Cells(i, 2).Font.ColorIndex
            if this_is != color_index:
                zz.append(i)

        c = [zz[0]]
        mass.append(c)
        for i in zz[1:]:
          if i == c[-1] + 1:
            c.append(i)
          else:
            c = [i]
            mass.append(c)
        
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar = ShadyBar('Удаляю строки    ', max=len(mass))
        for i in reversed(mass):
            # print('Удаляю строки с %s по %s' % (min(i), max(i)))
            bar.next()
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
        print('%s: %s (%s)' % (colors.index(oper), color_base.get(int(oper), "неизвестно"), oper))
        # print(colors.index(oper), oper, color_base.get(int(oper), "неизвестно"))
    print('Выберите цвет текста, который должен остаться:')
    answer = input()
    try:
        answer = int(answer)
        if colors[answer]:
            # clear()
            print('Вы выбрали цвет: %s (%s)' % (color_base.get(int(colors[answer]), "неизвестно"), colors[answer]))
            remove_rows(int(colors[answer]))
            print('')
        else:
            print('Ошибка при выборе цвета. Выход')
    except BaseException as identifier:
        print('Ошибка при выборе цвета: %s. Выход' % (identifier))
        exit()

    getNumRows = allData.Rows.Count
    print ('Строк после обработки: ',getNumRows)

def update_PQ(wb):
    wb.RefreshAll()
    return True

def update_file(wb):
    print('Нажал пымпу "Обновить всё" ждемс...')
    for x in wb.Connections:
        x.OLEDBConnection.BackgroundQuery = False
    update_PQ(wb)
    # print('Конец обновления файла')

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
        print('%s: %s' % (operlist.index(oper), oper))
    
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
    sheet.PageSetup.Orientation = 2
    sheet.PageSetup.TopMargin = 58
    sheet.PageSetup.BottomMargin = 86
    sheet.PageSetup.LeftMargin = 14
    sheet.PageSetup.RightMargin = 14

def save_as_new_book(wb, xlsFileName):
    filename = xlsFileName.split('\\')[-1]
    filename = filename.split('.')[0]
    filename = '%s%s%s' % (filename, '-new', '.xlsx')
    filepath = os.path.join(BASE_DIR, 'out', filename)
    wb.SaveAs(filepath)
    print('Обработанный файл сохранен по пути: %s' % filepath)
    wb.Close()

def fileUpdate(xl, xlsFileName):
    suffix = '%(percent)d%%'
    with ShadyBar('Прогресс по файлу', suffix=suffix, max=6) as bar:
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar.next()
        print('')
        print('Открываю...')
        wb = xl.Workbooks.Open(xlsFileName)
        xl.Visible = True
        xl.DisplayAlerts = False
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar.next()
        print('')

        remove_color_rows(wb, xlsFileName)
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar.next()
        print('')

        update_file(wb)
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar.next()
        print('')

        update_operation(wb)
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar.next()
        print('')

        save_as_new_book(wb, xlsFileName)
        del wb
        clear()
        print('Обработка файла: %s' % xlsFileName)
        bar.next()
        print('')

def filefinder():
    in_folder_flag = False
    out_folder_flag = False

    # print(in_folder)       
    folder = []
    for i in os.walk(BASE_DIR):
        folder.append(i)
    for address, dirs, files in folder:
        folder = address.split('\\')[-1]
        if in_folder == folder:
            in_folder_flag = True
        if out_folder == folder:
            out_folder_flag = True
    
    if in_folder_flag:
        print('Обнаружена папка файлов для обработки.')
        print('Список файлов с которыми будем работать:')
        folder = []
        for i in os.walk(os.path.join(BASE_DIR, in_folder)):
            folder.append(i)
        file_path_list = []
        for address, dirs, files in folder:
            for file in files:
                if file.split('.')[-1] == 'xlsx':
                    print(file.split('.')[0])
                    file_path_list.append(os.path.join(address, file))
    else:
        print('Папка входящих файлов не обнаружена.')
        in_folder_path = os.path.join(BASE_DIR, in_folder)
        print('Создаю папку: %s' % in_folder_path)
        try:
            os.mkdir(in_folder_path)
        except OSError:
            print("Ошибка при создании папки: %s \nАварийный выход." % in_folder_path)
            exit()
        else:
            print('Скопируйте нужные файлы в созданную папку и запустите скрипт еще раз. \nАварийный выход.')
            exit()

    out_folder_path = os.path.join(BASE_DIR, out_folder)
    if out_folder_flag:
        print('Обнаружена папка результатов обработки: %s' % out_folder_path)
        print('Очищаю')
        time.sleep(3)
        try:
            import shutil
            shutil.rmtree(out_folder_path)
        except OSError:
            print ("Deletion of the directory %s failed" % out_folder_path)
        else:
            try:
                os.mkdir(out_folder_path)
            except OSError:
                print("Ошибка при создании папки: %s \nАварийный выход." % out_folder_path)
                exit()
    else:
        print('Создаю папку результатов обработки: %s' % out_folder_path)
        try:
            os.mkdir(out_folder_path)
        except OSError:
            print("Ошибка при создании папки: %s \nАварийный выход." % in_folder_path)
            exit()
    return file_path_list

def worker():
    xl = win32com.client.DispatchEx("Excel.Application")
    files = filefinder()
    for file in files:
        # clear()
        # print('')
        # print('Обработка файла: %s' % file)
        fileUpdate(xl, file)
    xl.Quit()
    del xl
    clear()
    # print('********************************************************************')
    # print('')
    print('Все файлы обработаны. Программа закроется автоматически через 3сек.')
    print('Powered by Yegor Kowalew')
    time.sleep(3)

if __name__ == "__main__":
    worker()
    # python -OO -m PyInstaller --onefile run.py