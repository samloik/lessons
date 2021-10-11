
import pandas as pd

PATH_TO_THE_FILES = {
    'РАБОТА':'C:\Andrew files\\',
    'ДОМ': 'C:\Andrew files\\'
}

PATH = PATH_TO_THE_FILES ['РАБОТА']

IN_FILES = [
    '2021.10.11 продажи за год KYB ТС.xlsx',
    '2021.10.11 продажи за год KYB ШАВ.xlsx',
    '2021.10.11 продажи за год KYB ШСВ.xlsx'
]

OUT_FILES = [
    'общий.xlsx'
]


def andrew_task():
    print( 'IN_FILES\n', IN_FILES )
    print( '\nOUT_FILES\n', OUT_FILES)
    file_name1 = PATH + '\'' + IN_FILES[0] + '\''
    print('*C:\Andrew files\\2021.10.11 продажи за год KYB ТС.xlsx', 'rb')
    #data1 = pd.read_excel( open('C:\Andrew files\\2021.10.11 продажи за год KYB ТС.xlsx','rb'))
    #data1 = pd.read_excel (open(file_name1,'rb')

    print( file_name1 )




andrew_task()
print( '\n\n hello ')
