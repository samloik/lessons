import pandas as pd  # pip install openpyxl - проблема ушла
import os  # для os.chdir() и problem()
import xlsxwriter
import openpyxl
import shutil  # для модуля problem
from zipfile import ZipFile  # для модуля problem
from collections import defaultdict
import pathlib # для модуля problem
import numpy as np  # для поиска числовых столбцов

PATH_TO_THE_FILES = {
    'РАБОТА': 'C:\Andrew files',
    'ДОМ': 'C:\Andrew files'
}

PATH = PATH_TO_THE_FILES['РАБОТА']

FilesList = []

OUT_FILES = [
    'общий.xlsx',
    'общий2.xlsx'
]

df = {}

# получаем список всех файлов по пути path """

def read_filenames(path):

    f = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        f.extend(filenames)
        break

    count = 1
    for i in f:
        if i != OUT_FILES[0]:
            FilesList.append([i, False])

            # загружаем массивы и ловим ошибки в файлах
            df[count] = try_load(PATH + '\\'+i)

            # подготовка массивов к работе - удаление лишних строк и столбцов
            df[count] = preparation(df[count])

            if count > 1:
                df_out = append_file(df_out, df[count])
            else:
                df_out = df[1]
            count += 1



    # добавление строчек ИТОГО в конец колонки склада со значением суммы
    df_out = append_TOTAL(df_out)

    #print('FilesList: ', FilesList)

    return df_out

# объединение датафраймов в один
def append_file(df_out, df ):

    print ('len(df.columns): ', len(df.columns))
    print('len(df_out.columns): ',len(df_out.columns))

    df_c = len(df.columns)
    df_out_c = len(df_out.columns)

    count = 0
    for x in range( 2, df_out_c ):
        for y in range( 2, df_c ):
            print ( 'df_out.loc[0, df_out.columns[x]]:', df_out.iloc[0, x])
            print('df.loc[0, df.columns[y]]:', df.iloc[0, y])

            # если столцы одиннаковые
            if df_out.iloc[0,x] == df.iloc[0, y]:
                #df_out[df_out.columns[x]] = df_out[df_out.columns[x]] + df[df.columns[y]]
                #df_out[df_out.columns[x]] = \
                df2 = pd.merge(df_out, df, on=['Номенклатура.Код', 'Номенклатура'], how='inner')
                print('df: ', df.info())
                print('df2: ', df2.info())
                df3 = df.drop(df2.index, inplace=True)
                print('df3: ', df3.info())

                    #append_row(df_out[df_out.columns[x]], df[df.columns[y]])
                count += 1
            else:
                #df_out[df.columns[y]] = df[df.columns[y]]
                df_out[df.columns[y]] = append_row(df_out[df_out.columns[x]], df[df.columns[y]])



    print( 'COUNT: ', count)
    #if count == 0:
    #    df_out = pd.merge(df_out, df, on=['Номенклатура.Код', 'Номенклатура'], how='outer')

    return df_out

def append_row( row1, row2 ):
    nrow = row1 + row2

    return nrow




# тест функции enumerate
def test():
    cols = ["A", "B", "C", "D", "E"]
    txt = [0, 1, 2, 3, 4]

    # Loop over the rows and columns and fill in the values
    for num in range(5):
        row = num
        print(row)
        for index, col in enumerate(cols):
            value = txt[index] + num
            print(col, index, ' = ', value)


""" загружаем файл (read_excel) и ловим ошибку "There is no item named 'xl/sharedStrings.xml' in the archive" """


def try_load(f):
    file_name = f

    try:
        DataFrame = pd.read_excel(file_name)
        return DataFrame
    except KeyError as Error:
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            problem(file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            DataFrame = pd.read_excel(file_name)
            for file in FilesList:
                name = PATH + '\\' + file[0]
                if file_name == name:
                    file[1] = True
            return DataFrame
        else:
            print('Ошибка: >>' + str(Error) + '<<')


# вернуть файлам первоначальный вид
def un_pack():
    for file in FilesList:
        name = PATH + '\\' + file[0]
        if file[1] == True:
            print( 'file[0]: ', name)
            un_problem( name )
            file[1] = False



# переименовывание файла 'SharedStrings.xml' в файл 'sharedStrings.xml' в архиве excel-файла filename
def problem(filename):
    tmp_folder = PATH + '\\tmp\\'

    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку
    with ZipFile(filename) as excel_container:
        excel_container.extractall(tmp_folder)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive(filename, 'zip', tmp_folder)
    os.remove(filename)
    os.rename(filename + '.zip', filename)

    # удалить папку tmp и все вложения
    delete_folder(tmp_folder)



# удалить папку и все содержимое
def delete_folder(pth):
    for root, dirs, files in os.walk(pth, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    os.rmdir(pth)




# переименовывание файла обратно 'sharedStrings.xml' в файл 'SharedStrings.xml' в архиве excel-файла filename
def un_problem(filename):
    tmp_folder = PATH + '\\tmp\\'

    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку
    with ZipFile(filename) as excel_container:
        excel_container.extractall(tmp_folder)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive(filename, 'zip', tmp_folder)
    os.remove(filename)
    os.rename(filename + '.zip', filename)

    """ удалить папку tmp и все вложения """
    delete_folder(tmp_folder)



def preparation(df1):

    #print("df1['Номенклатура.Код']: \n", df1.loc[df1['Unnamed: 0'] == 'Номенклатура.Код'] )

    df = df1.loc[df1['Unnamed: 0'] == 'Номенклатура.Код']
    StrIndex = df1.loc[df1['Unnamed: 0'] == 'Номенклатура.Код'].index[0]

    maxcol = len(df.columns)

    for col in range(maxcol):
        name = df.iloc[0,col]
        if not name or pd.isnull(name):
            # удаляем ненужные столбцы
            df1.drop(columns=['Unnamed: '+str(col)], axis=1, inplace=True)
        else:
            # переименовываем столбцы
            df1 = df1.rename(columns={('Unnamed: '+str(col)): name})


    # удаляем ненужные строки первые
    df1.drop(df1.head(StrIndex+2).index, inplace=True)

    # удаляем последнюю строку
    df1.drop(df1.tail(1).index, inplace=True)
    return df1


def append_TOTAL(df ):
    """ добавление строчки ИТОГО в конец колонки склада со значением суммы """

    # - костыль чтобы убрать ошибку:
    # SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame
    # https://coderoad.ru/20625582/%D0%9A%D0%B0%D0%BA-%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%B8%D1%82%D1%8C%D1%81%D1%8F-%D1%81-SettingWithCopyWarning-%D0%B2-Pandas

    pd.options.mode.chained_assignment = None

    # добавляем строчку итого
    df = df.append({'Номенклатура.Код': 'Итого'}, ignore_index=True)

    maxcol = len(df.columns)
    row = df[(df['Номенклатура.Код'] == 'Итого')].index[0]

    for col in range(2, maxcol):
        # записываем в строку итого суммы столоцов с 3го до конца
        df.iloc[row, col] = df[df.columns[col]].sum(axis=0)

    return df


def andrew_task():

    file_name_out = PATH + '\\' + OUT_FILES[0]


    os.chdir(PATH)

    """ загружаем массивы и ловим ошибки в файлах """

    df3_merge_outer = read_filenames(PATH)

    # подготовка массивов к работе - удаление лишних строк и столбцов

    print( 'len(df): ', len(df))


    writer = pd.ExcelWriter(file_name_out)
    df3_merge_outer.to_excel(writer, sheet_name='df3_merge_outer',index=False)
    workbook = writer.book
    worksheet = writer.sheets['df3_merge_outer']
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 100)
    worksheet.set_column(2, len(df3_merge_outer.columns), 15)
    writer.save()
    print (len(df3_merge_outer.columns))

    # вернуть файлам первоначальный вид
    un_pack()




# test()

andrew_task()

print ( FilesList )