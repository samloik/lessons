import pandas as pd  # pip install openpyxl - проблема ушла
import os  # для os.chdir() и problem()
import glob
import xlsxwriter
import openpyxl
import shutil  # для модуля problem
from zipfile import ZipFile  # для модуля problem
import numpy as np  # для поиска числовых столбцов

PATH_TO_THE_FILES = {
    'РАБОТА': 'C:\Andrew files',
    'ДОМ': 'C:\Andrew files'
}

PATH = PATH_TO_THE_FILES['РАБОТА']

IN_FILES = [
    '2021.10.11 продажи за год KYB ТС.xlsx',
    '2021.10.11 продажи за год KYB ШАВ.xlsx',
    '2021.10.11 продажи за год KYB ШСВ.xlsx'
]

FilesList = []

OUT_FILES = [
    'общий.xlsx',
    'общий2.xlsx'
]


# получаем список всех файлов по пути path """

def read_filenames(path):

    f = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        f.extend(filenames)
        break

    for i in f:
        FilesList.append([i, False])

    # print('FilesList: ', FilesList)
    return FilesList


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
            return DataFrame
        else:
            print('Ошибка: >>' + str(Error) + '<<')


# переименовывание файла 'SharedStrings.xml' в файл 'sharedStrings.xml' в архиве excel-файла filename
def problem(filename):
    tmp_folder = '/tmp/convert_wrong_excel/'
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

    """ удалить папку tmp и все вложения
    files = glob.glob(tmp_folder+'**/*.*', recursive=True)
    print( "tmp_folder+'/**/*.*' >>> ",tmp_folder+'**/*.*' )

    for f in files:
        try:
            os.remove(f)
        except OSError as e:
            print("Error: %s : %s" % (f, e.strerror))

    try:
        os.rmdir(tmp_folder)
    except OSError as e:
        print("Error: %s : %s" % (tmp_folder, e.strerror))
    """

def preparation1(df1):

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



def preparation01(df1):
    # удаляем ненужные столбцы
    cols = [1, 2, 4, 5]  # 0, 3, 6
    df1.drop(df1.columns[cols], axis=1, inplace=True)

    # удаляем ненужные строки первые 10
    df1.drop(df1.head(10).index, inplace=True)

    # удаляем последнюю строку
    df1.drop(df1.tail(1).index, inplace=True)

    # переименовываем столбцы
    df1 = df1.rename(columns={'Unnamed: 0': 'Номенклатура.Код'})
    df1 = df1.rename(columns={'Unnamed: 3': 'Номенклатура'})
    df1 = df1.rename(columns={'Unnamed: 6': '05_Pavlovsky'})

    return df1


def preparation2(df2):
    # удаляем ненужные столбцы
    cols = [1, 2, 4, 6]
    df2.drop(df2.columns[cols], axis=1, inplace=True)

    # удаляем ненужные строки первые 10
    # rows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    # df2.drop( rows,  inplace=True  )
    df2.drop(df2.head(11).index, inplace=True)

    # удаляем последнюю строку
    df2.drop(df2.tail(1).index, inplace=True)

    # переименовываем столбцы
    df2 = df2.rename(columns={'Unnamed: 0': 'Номенклатура.Код'})
    df2 = df2.rename(columns={'Unnamed: 3': 'Номенклатура'})
    df2 = df2.rename(columns={'Unnamed: 5': '02_Car'})
    df2 = df2.rename(columns={'Unnamed: 7': '04_Victory'})
    df2 = df2.rename(columns={'Unnamed: 8': '08_Center'})

    return df2


def preparation3(df3):
    # удаляем ненужные столбцы
    cols = [1, 2, 4, 6]  # 0,3,5,7,8  - нужны, 1,2,4,6 - удалить
    df3.drop(df3.columns[cols], axis=1, inplace=True)

    # удаляем ненужные строки первые 10
    df3.drop(df3.head(12).index, inplace=True)

    # удаляем последнюю строку
    df3.drop(df3.tail(1).index, inplace=True)

    # переименовываем столбцы

    df3 = df3.rename(columns={'Unnamed: 0': 'Номенклатура.Код'})
    df3 = df3.rename(columns={'Unnamed: 3': 'Номенклатура'})
    df3 = df3.rename(columns={'Unnamed: 5': '01_Kirova'})
    df3 = df3.rename(columns={'Unnamed: 7': '03_Inter'})
    df3 = df3.rename(columns={'Unnamed: 8': '09_Station'})

    return df3


def unique_values(df1, df2):
    # получить уникальные значения df2 относительно df1, которые находятся только в df2
    df2_unique_vals = df2[~df2['Номенклатура.Код'].isin(df1['Номенклатура.Код'])]

    """
    # получить уникальные значения, которые находятся только в df1
    df1_unique_vals = df1[~df1['Номенклатура.Код'].isin(df2['Номенклатура.Код'])]
    
    # получить как неуникальные значения, которые находятся только в df1
    df1_unique_vals = df1[df1['Номенклатура.Код'].isin(df2['Номенклатура.Код'])]

    Чтобы получить как значения, которые находятся только в df1, 
    так и значения, которые находятся только в df2, 
    вы можете сделать это

    df_unique_vals = df1[~df1['Номенклатура.Код'].isin(df2['Номенклатура.Код'])].append(df2[~df2.['Номенклатура.Код'].isin(df1.['Номенклатура.Код'])], ignore_index=True)

    """

    return df2_unique_vals


def equel_values(df1, df2):
    # получить общие значения df2 относительно df1, которые находятся только в df2
    df2_equel_vals = df2[df2['Номенклатура.Код'].isin(df1['Номенклатура.Код'])]

    return df2_equel_vals


def append_TOTAL(df ):
    """ добавление строчки ИТОГО в конец колонки склада со значением суммы """

    # - костыль чтобы убрать ошибку:
    # SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame
    # https://coderoad.ru/20625582/%D0%9A%D0%B0%D0%BA-%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%B8%D1%82%D1%8C%D1%81%D1%8F-%D1%81-SettingWithCopyWarning-%D0%B2-Pandas

    pd.options.mode.chained_assignment = None

    df = df.append({'Номенклатура.Код': 'Итого'}, ignore_index=True)

    maxcol = len(df.columns)

    #print( "df.loc[df['Unnamed: 0'] == 'Номенклатура.Код'].index",
    #       df.loc[df['Unnamed: 0'] == 'Номенклатура.Код'].index)
    print(len(df.columns) )
    print(df.columns)

    row = df[(df['Номенклатура.Код'] == 'Итого')].index[0]

    print( 'row: \n', row)
    print( df.iloc[row, 0])
    print('maxcol: ', maxcol)

    for col in range(2, maxcol):
        df.iloc[row, col] = df[col].sum(axis=0)

    return df


def andrew_task():
    """ формируем имена файлов из пути и наименований """
    file_name1 = PATH + '\\' + IN_FILES[0]
    file_name2 = PATH + '\\' + IN_FILES[1]
    file_name3 = PATH + '\\' + IN_FILES[2]
    file_name_out = PATH + '\\' + OUT_FILES[0]
    file_name_out2 = PATH + '\\' + OUT_FILES[1]

    os.chdir(PATH)

    """ загружаем массивы и ловим ошибки в файлах """
    df1 = try_load(file_name1)
    df2 = try_load(file_name2)
    df3 = try_load(file_name3)

    """ подготовка массивов к работе - удаление лишних строк и столбцов """
    df1 = preparation1(df1)  # 0, 3, 6    - нужны, 1,2,4,5 - удалить
    df2 = preparation1(df2)  # 0,3,5,7,8  - нужны, 1,2,4,6 - удалить
    df3 = preparation1(df3)  # 0,3,5,7,8  - нужны, 1,2,4,6 - удалить

    """ получить как уникальные значения, которые находятся только в df2 """

    # https: // pandas.pydata.org / docs / user_guide / merging.html
    # Merge, join, concatenate and compare

    # --- добавлены новые столбцы с уникальными позициями

    df2_merge_outer = pd.merge(df1, df2, on=['Номенклатура.Код', 'Номенклатура'], how='outer')
    df3_merge_outer = pd.merge(df2_merge_outer, df3, on=['Номенклатура.Код', 'Номенклатура'], how='outer')

    # print("\ndf2_unique_vals['02_Car'].count():", df2_unique_vals['02_Car'].count())

    # поиск числовых столбцов
    # df = df2_unique_vals
    # https://coderoad.ru/35003138/Python-Pandas-%D0%B2%D1%8B%D0%B2%D0%BE%D0%B4-%D1%82%D0%B8%D0%BF%D0%BE%D0%B2-%D0%B4%D0%B0%D0%BD%D0%BD%D1%8B%D1%85-%D1%81%D1%82%D0%BE%D0%BB%D0%B1%D1%86%D0%BE%D0%B2

    # print(pd.DataFrame(df.apply(pd.api.types.infer_dtype, axis=0)).reset_index().rename(
    #    columns={'index': 'column', 0: 'type'}))

    """
    # поиск числовых столбцов
    colnames_numerics_only = df.select_dtypes(include=np.number).columns.tolist()
    print( '\ncolnames_numerics_only:',colnames_numerics_only)
    """

    """ добавление строчек ИТОГО в конец колонки склада со значением суммы """
    df1 = append_TOTAL(df1 )
    df2 = append_TOTAL(df2 )
    df3 = append_TOTAL(df3 )

    df3_merge_outer = append_TOTAL(df3_merge_outer )

    # print("\ndf2_unique_vals['02_Car'].count():", df2_unique_vals['02_Car'].count())
    # df3_merge_outer.index = pd.date_range( '1900/1/30', periods = df3_merge_outer.shape[0] )

    print(r'df3_merge_outer.shape[0]>>', df3_merge_outer.shape[0])
    print('df3_merge_outer[Номенклатура].unique() >>>', df3_merge_outer['Номенклатура'].nunique())

    #
    # df3_merge_outer = df3_merge_outer.append({'Номенклатура.Код': 'TotalTOTAL'}, ignore_index=True)

    writer = pd.ExcelWriter(file_name_out)
    df3_merge_outer.to_excel(writer, sheet_name='df3_merge_outer')
    workbook = writer.book
    worksheet = writer.sheets['df3_merge_outer']
    worksheet.set_column(1, 1, 11)
    worksheet.set_column(2, 2, 100)
    worksheet.set_column(3, 3, 20)
    worksheet.set_column(4, 4, 20)
    worksheet.set_column(5, 9, 10)
    writer.save()

    # print( '\n>>>', df3_merge_outer.describe(include='all') )
    # ddf3 = df3_merge_outer['Номенклатура'].unique()

    # print("\nddf3>>>", ddf3.describe(include='all'))

    """ Запись результатов обработки в excel файл 
    with pd.ExcelWriter(file_name_out) as writer:
        df2_merge_outer.to_excel(writer, sheet_name='df2_merge_outer')
        
        df1.to_excel(writer, sheet_name='df1')
        df2.to_excel(writer, sheet_name='df2')
        df3.to_excel(writer, sheet_name='df3')
    """

    """
    информация по методу изменения ширины столбцов
    
    https://coderoad.ru/17326973/%D0%95%D1%81%D1%82%D1%8C-%D0%BB%D0%B8-%D1%81%D0%BF%D0%BE%D1%81%D0%BE%D0%B1-%D0%B0%D0%B2%D1%82%D0%BE%D0%BC%D0%B0%D1%82%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8-%D1%80%D0%B5%D0%B3%D1%83%D0%BB%D0%B8%D1%80%D0%BE%D0%B2%D0%B0%D1%82%D1%8C-%D1%88%D0%B8%D1%80%D0%B8%D0%BD%D1%83-%D1%81%D1%82%D0%BE%D0%BB%D0%B1%D1%86%D0%BE%D0%B2-Excel-%D1%81-%D0%BF%D0%BE%D0%BC%D0%BE%D1%89%D1%8C%D1%8E
    """

    # print(df10)

    """
    # создание файла для записи - тест
    
    # открываем новый файл на запись
    workbook = xlsxwriter.Workbook( file_name_out )

    # создаем там "лист"
    worksheet = workbook.add_worksheet()

    # в ячейку A1 пишем текст
    worksheet.write('A1', 'Hello world')

    worksheet.write(0, 0, 'Это A1!')
    worksheet.write(4, 3, 'Колонка D, стока 5')

    # сохраняем и закрываем
    workbook.close()
    """

    """
    # временный вывод для контроля
    print( '\nHead():' )
    print( oDataFrame.head() )

    print( '\ninfo():' )
    print( oDataFrame.info() )
    """

    # for col in oDataFrame:


# test()
andrew_task()

read_filenames(PATH)
