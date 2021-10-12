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

#print('problem: >>', filename)
def problem( filename ):
    tmp_folder = '/tmp/convert_wrong_excel/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку
    print('problem: >>', filename)                        # -----
    with ZipFile( filename ) as excel_container:
        excel_container.extractall(tmp_folder)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive( filename, 'zip', tmp_folder)
    os.rename( filename+'.zip', filename )


def andrew_task():
    print('IN_FILES\n', IN_FILES)
    print('\nOUT_FILES\n', OUT_FILES)

    file_name1 = PATH + '\\' + IN_FILES[0]
    file_name2 = PATH + '\\' + IN_FILES[1]
    file_name3 = PATH + '\\' + IN_FILES[2]
    file_name_out = PATH + '\\' + OUT_FILES[0]

    print(file_name1)
    print(file_name2)
    print(file_name3)
    print(file_name_out)

    #print( DataFrame1)
    os.chdir( PATH )
    #print ( PATH)

    #DataFrame1 = pd.ExcelFile(file_name1)

    try_load2( file_name2 )

def try_load2( f ):
    file_name = f
    problem( file_name)
    #DataFrame1 = pd.read_excel( file_name )

    #print( DataFrame1 )


def try_load( f ):
    file_name = f
    try:
         DataFrame1 = pd.read_excel( file_name )
         print(DataFrame1)
    except KeyError as Error:
        if Error == "There is no item named 'xl/sharedStrings.xml' in the archive":
            #"There is no item named 'xl/sharedStrings.xml' in the archive":
            print ('####', Error)
            sys.exit(1)
        else:
            print ( 'Error:', Error )

    #print( KeyError, TypeError, NameError )

    #print( DataFrame1 )




#test()
#problem()
andrew_task()
print('\n\n hello ')


