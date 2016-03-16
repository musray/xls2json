#!python3
# xls2json.py
# This module is used to convert particular excel files into json file
# of which will be used in the T.E.A.M web application.
# The excel file can be converted is listed in docMapping dictionary.
# When excel files get extended either horizentally or vertically,  
# the second element in the value list of docMapping should be updated accordingly.

import win32com.client as win32
import os, re, json, sys


docMapping = {
            '五组标准化文件列表.xlsx': ['doc-list.json', 'H285'],
            'acronym.xlsx':['acronym.json', 'C1027'],
            '五组设计工具清单.xls': ['tool-list.json', 'H33']
        }

def getExcelRows(file):
    '''
      the parameter *file* should be a tuple which
      consists of file_name, abs_path_with_file_name
    '''

    excel = win32.DispatchEx('Excel.Application')
    wb = excel.Workbooks.Open(file[1])
    ws = wb.Worksheets(1)
    bottomRight = docMapping[file[0]][1]
    rows = ws.Range('A1', bottomRight).Value
    wb.Close()
    return rows


def getAllFiles(fileExt):
    files = [file for file in os.listdir() if file.endswith('.' + fileExt)]
    return files

def getArgvFile(file_list):
    # 根据sys.argv清单，返回文件名及绝对路径
    result = []
    for file in file_list:
        if os.path.isfile(file):
            fileAbsPath = os.path.join(os.path.abspath('.'), file)
            result.append( (file, fileAbsPath) )
    return result


def Jgenerator(file):
    # the parameter file should be a tuple which
    # consists of file_name, abs_path_with_file_name

    # getExcelRows is a function
    rows = getExcelRows(file)
    jsonKeys = rows[0]

    Jname = docMapping[file[0]][0]
    with open(Jname, 'w', encoding='utf-8') as f:
        f.write('[\n')
        for numOfRows in range(1, len(rows)):
            aDict = {}
            for numOfKey in range(len(jsonKeys)):
                aDict[jsonKeys[numOfKey]] = rows[numOfRows][numOfKey] 
            aJson = json.dumps(aDict, ensure_ascii=False)
            if not numOfRows == len(rows) - 1:
                f.write(aJson + ',' + '\n')
            else:
                f.write(aJson + '\n]')

excelFiles = getArgvFile(sys.argv[1:])
for file in excelFiles:
    # here file should be a 
    # tuple (file_name, abs_path_with_file_name)
    Jgenerator(file)
