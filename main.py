import pandas as pd # pip install openpyxl - проблема ушла
import os # для os.chdir() и problem()


PATH_TO_THE_FILES = {
    'РАБОТА': "C:\Andrew files",
    'ДОМ': "C:\Andrew files"
}

PATH = PATH_TO_THE_FILES['РАБОТА']

IN_FILES = [
    '2021.10.11 продажи за год KYB ТС.xlsx',
    '2021.10.11 продажи за год KYB ШАВ.xlsx',
    '2021.10.11 продажи за год KYB ШСВ.xlsx'
]

OUT_FILES = [
    'общий.xlsx'
]


# тест функции enumerate
def test():
    cols = ["A", "B", "C", "D", "E"]
    txt = [0, 1, 2, 3, 4]

    # Loop over the rows and columns and fill in the values
    for num in range(5):
        row = num
        print( row)
        for index, col in enumerate(cols):
            value = txt[index] + num
            print(col, index, ' = ', value)


import shutil
from zipfile import ZipFile

# переименовывание файла 'SharedStrings.xml' в файл 'sharedStrings.xml' в архиве excel-файла filename
def problem( filename ):
    tmp_folder = '/tmp/convert_wrong_excel/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку
    with ZipFile( filename ) as excel_container:
        excel_container.extractall(tmp_folder)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive( filename, 'zip', tmp_folder)
    os.remove( filename )
    os.rename( filename+'.zip', filename )


def andrew_task():
    """ временный вывод имен для контроля """
    print('IN_FILES\n', IN_FILES)
    print('\nOUT_FILES\n', OUT_FILES)

    # формируем имена файлов из пути и наименований
    file_name1 = PATH + '\\' + IN_FILES[0]
    file_name2 = PATH + '\\' + IN_FILES[1]
    file_name3 = PATH + '\\' + IN_FILES[2]
    file_name_out = PATH + '\\' + OUT_FILES[0]

    """ временный вывод имен для контроля """
    print(file_name1)
    print(file_name2)
    print(file_name3)
    print(file_name_out)
    print('\n')

    os.chdir( PATH )

    """ загружаем массивы и ловим ошибки в файлах """
    DataFrame1 = try_load( file_name1 )
    DataFrame2 = try_load( file_name2 )
    DataFrame3 = try_load( file_name3 )

    oDataFrame = DataFrame1


    """ временный вывод для контроля """
    print( '\nHead():' )
    print( oDataFrame.head() )

    print( '\ninfo():' )
    print( oDataFrame.info() )

# загружаем файл (read_excel) и ловим ошибку "There is no item named 'xl/sharedStrings.xml' in the archive"
def try_load( f ):
    file_name = f
    try:
         DataFrame = pd.read_excel( file_name )
         return DataFrame
    except KeyError as Error:
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            problem( file_name )
            print ('Исправлена ошибка: ', Error, f'в файле: \"{ file_name }\"\n' )
            DataFrame = pd.read_excel(file_name)
            return DataFrame
        else:
            print ( 'Ошибка: >>' + str(Error) + '<<' )




#test()
andrew_task()



