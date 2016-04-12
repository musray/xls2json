#!python3
# io2json.py
# This module is used to convert particular IO list to json file
# of which will be used in the T.E.A.M web application.
# When excel files get extended either horizentally or vertically,  
# the second element in the value list of docMapping should be updated accordingly.

import win32com.client as win32
import os, re, json, sys, copy
header = ['io_tag_no', 'signal_name', 'io_type', 'card_type', 'config', 'master_card.mc_no', 'master_card.ch_no1', 'master_card.ch_no2', 'io_card_location.cf1_no', 'io_card_location.cf2_no', 'io_card_location.iou_no', 'io_card_location.sl_no', 'io_card_location.ch_no', 'output_setting.fail_mode', 'output_setting.clk', 'output_setting.group', 'distribution', 'terminal_block.no', 'terminal_block.terminal', 'device_no', 'signal_condition', 'contact_type', 'connection_source', 'relevent_sheet', 'remark1', 'remark2', 'remark3', 'id', 'sheet_no', 'cnpdc_id_code', 'ext_code', 'cnpdc_desig', 'bdsd_sheet', 'cabinet_id', 'wd_drawing_no', 'wd_index_no', 'single_redundant', 'power_supply' ]

def getExcelRows(file):
    '''
      the parameter *file* should be a tuple which
      consists of file_name, abs_path_with_file_name
    '''

    excel = win32.DispatchEx('Excel.Application')
    wb = excel.Workbooks.Open(file[1])
    ws = wb.Worksheets(1)
    countRows = 0
    for i in range(3, 10000):
        countRows += 1
        if ws.Range('G' + str(i)).Value == None and \
           ws.Range('H' + str(i)).Value == None :
            break
    rows = ws.Range('A1', 'AP'+str(countRows)).Value
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
    # IMPORTANT: file is a turple including
    # both file name and the absolute path
    # such as
    #('HYH3 ESFAC-A DIO.xls', 'D:\\dev\\xls2json\\HYH3 ESFAC-A DIO.xls')

    # getExcelRows is a customized function
    # see above
    rows = getExcelRows(file)[2:]

    # Jname = docMapping[file[0]][0]
    Jname = os.path.splitext(file[1])[0] + ".json"

    with open(Jname, 'w', encoding='utf-8') as f:
        f.write('[\n')

        count = 0
        for row in rows:
            count += 1
            aDic = {}
            for column in header:
                if '.' in column:
                    # here the format of key might like
                    # this: parent.child
                    # we would like to split parent and child
                    sub1, sub2 = column.split('.')
                    aDic.setdefault(sub1, {})
                    cellValue = row[ header.index(column) + 1 ] # +1 to skip the first row in IO List
                    # py32com read number as float, such as 1.0, 3.0
                    # We'd like to make it a string
                    if type(cellValue) == type(1.0):
                        cellValue = str(int(cellValue))
                    aDic[sub1][sub2] = cellValue
                else: 
                    aDic[column] = row[ header.index(column) + 1 ]
            aJson = json.dumps(aDic, ensure_ascii=False);

            # Now let's wirte the row object
            # to the *.json
            # When it is at the last row, exclude the comma
            if count == len(rows):  # is the last row
                f.write(aJson)
            else:  # not the last row
                f.write(aJson + ',' + '\n')

        f.write('\n]')

if __name__ == '__main__':

    excelFiles = getArgvFile(sys.argv[1:])
    for file in excelFiles:
        # here file should be a 
        # tuple (file_name, abs_path_with_file_name)
        Jgenerator(file)

