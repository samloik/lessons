import pandas as pd # pip install openpyxl - проблема ушла
import os # для os.chdir() и problem()
import xlsxwriter
import openpyxl
import shutil                   # для модуля problem
from zipfile import ZipFile     # для модуля problem


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



""" загружаем файл (read_excel) и ловим ошибку "There is no item named 'xl/sharedStrings.xml' in the archive" """
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

def preparation1( df1 ):
    # удаляем ненужные столбцы
    cols = [1,2,4,5]                    # 0, 3, 6
    df1.drop( df1.columns[cols], axis = 1, inplace=True )

    # удаляем ненужные строки первые 10
    df1.drop(df1.head(10).index, inplace=True)

    # удаляем последнюю строку
    df1.drop(df1.tail(1).index, inplace=True)
    return df1

def preparation2( df2 ):
    # удаляем ненужные столбцы
    cols = [1,2,4,6]
    df2.drop( df2.columns[cols], axis = 1, inplace=True )

    # удаляем ненужные строки первые 10
    #rows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    #df2.drop( rows,  inplace=True  )
    df2.drop(df2.head(11).index, inplace=True)

    # удаляем последнюю строку
    df2.drop(df2.tail(1).index, inplace=True)
    return df2

def preparation3( df3 ):
    # удаляем ненужные столбцы
    cols = [1,2,4,5]                                    # 0, 3, 6 - нужны эти столбцы
    df3.drop( df3.columns[cols], axis = 1, inplace=True )

    # удаляем ненужные строки первые 10
    df3.drop(df3.head(12).index, inplace=True)

    # удаляем последнюю строку
    df3.drop(df3.tail(1).index, inplace=True)
    return df3


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
    df1 = try_load( file_name1 )
    df2 = try_load( file_name2 )
    df3 = try_load( file_name3 )

    """ подготовка массивов к работе - удаление лишних строк и столбцов """
    df1 = preparation1( df1 )            # 0, 3, 6    - нужны, 1,2,4,5 - удалить
    df2 = preparation2( df2 )            # 0,3,5,7,8  - нужны, 1,2,4,6 - удалить
    df3 = preparation2( df3 )            # 0,3,5,7,8  - нужны, 1,2,4,6 - удалить

    df0 = df3

    #print( df3[ 'Unnamed: 0'])

    # https://coderoad.ru/43544514/Pandas-%D1%81%D1%87%D0%B8%D1%82%D1%8B%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5-%D0%BE%D0%BF%D1%80%D0%B5%D0%B4%D0%B5%D0%BB%D0%B5%D0%BD%D0%BD%D0%BE%D0%B3%D0%BE-%D0%B7%D0%BD%D0%B0%D1%87%D0%B5%D0%BD%D0%B8%D1%8F-%D1%8F%D1%87%D0%B5%D0%B9%D0%BA%D0%B8-Excel-%D0%B2-%D0%BF%D0%B5%D1%80%D0%B5%D0%BC%D0%B5%D0%BD%D0%BD%D1%83%D1%8E

    data = {}
    data[0] = df3['Unnamed: 0'].tolist()
    data[1] = df3['Unnamed: 3'].tolist()
    data[2] = df3['Unnamed: 5'].tolist()
    print('\n\n>>>')

    """
    for i in range( 0,len(data[0])):
        print ( data[0][i], data[1][i] )
        
    print( '>>', data[2][0])
    """

    df10 = pd.DataFrame( data[0], columns = ['col1'] )
    print( df10 )
    print ("\n\n")
    df10['col2' ] = data[1]
    df10['col3'] = data[2]
    print(df10)


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

    #for col in oDataFrame:
    

#test()
andrew_task()



