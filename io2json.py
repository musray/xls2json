#!python3
# io2json.py
# This module is used to convert particular IO list to json file
# of which will be used in the T.E.A.M web application.
# When excel files get extended either horizentally or vertically,  
# the second element in the value list of docMapping should be updated accordingly.

import win32com.client as win32
import os, re, json, sys, copy

headers = {
        'DIO_header' : ['unit', 'io_tag_no', 'signal_name', 'io_type', 'card_type', 'config', 'master_card.mc_no', 'master_card.ch_no1', 'master_card.ch_no2', 'io_card_location.cf1_no', 'io_card_location.cf2_no', 'io_card_location.iou_no', 'io_card_location.sl_no', 'io_card_location.ch_no', 'output_setting.fail_mode', 'distribution', 'terminal_block.no', 'terminal_block.terminal', 'device_no', 'signal_condition', 'contact_type', 'connection_source', 'relevent_sheet', 'remark1', 'remark2', 'remark3', 'id', 'sheet_no', 'rev', 'cnpdc_id_code', 'ext_code', 'cnpdc_desig', 'bdsd_sheet', 'cabinet_id', 'wd_drawing_no', 'wd_index_no', 'single_redundant', 'power_supply' ],
        'PIF_header' : ['unit', 'io_tag_no', 'signal_name', 'io_type', 'card_type', 'config', 'master_card.mc_no', 'master_card.ch_no1', 'master_card.ch_no2', 'io_card_location.cf1_no', 'io_card_location.cf2_no', 'io_card_location.iou_no', 'io_card_location.sl_no', 'io_card_location.ch_no', 'output_setting.fail_mode', 'output_setting.clk', 'output_setting.group', 'distribution', 'terminal_block.no', 'terminal_block.terminal', 'device_no', 'signal_condition', 'contact_type', 'connection_source', 'relevent_sheet', 'remark1', 'remark2', 'remark3', 'id', 'sheet_no', 'rev', 'cnpdc_id_code', 'ext_code', 'cnpdc_desig', 'bdsd_sheet', 'cabinet_id', 'wd_drawing_no', 'wd_index_no', 'single_redundant', 'power_supply'],
        'AIO_header' : ['unit', 'io_tag_no', 'signal_name', 'io_type', 'card_type', 'config', 'master_card.mc_no', 'master_card.ch_no1', 'master_card.ch_no2', 'io_card_location.cf1_no', 'io_card_location.cf2_no', 'io_card_location.iou_no', 'io_card_location.sl_no', 'eng_value.low', 'eng_value.hi', 'eng_value.unit', 'past_value_rate', 'overrange_low_enable', 'overrange_hi_enable', 'overrange_low_value', 'overrange_hi_value', 'input_setting.filter', 'input_setting.digital_filter', 'input_setting.lowcut', 'input_setting.pls_edge', 'input_setting.sq_root', 'input_setting.unused', 'output_setting.fail_mode', 'measurement_range', 'distribution', 'terminal_block.no', 'terminal_block.terminal', 'data_source.tag', 'data_source.connection', 'relevent_sheet', 'remark1', 'remark2', 'remark3', 'id', 'sheet_no', 'rev', 'cnpdc_id_code', 'ext_code', 'cnpdc_desig', 'bdsd_sheet', 'cabinet_id', 'wd_drawing_no', 'wd_index_no', 'single_redundant', 'power_supply']
}

def getHeader(file):
    '''
        take a file name like HYH3 ESFAC-A DIO.xls
        then determine of which type the feed in file name is 
        return the particular header associated to its type of IO
    '''
    all_types = {'AIO': 'AIO_header', 
                 'DIO': 'DIO_header',
                 '16DO':'DIO_header', 
                 'PIF': 'PIF_header'  }

    matcher = re.compile(r'DIO|16DO|AIO|PIF')
    IO_type = matcher.search(file)

    if IO_type == None:
        print('\nError -->')
        print('\t' + file + ' 无法确定IO清单类型（DIO？PIF？AIO？16DO），请检查文件命名并重新运行脚本。')
        sys.exit(0)

    return all_types[IO_type.group(0)]

def getUnit(file):
    '''
        take a file name like HYH3 ESFAC-A DIO.xls
        extract HYH3 from the name then return it
    '''
    matcher = re.compile(r'^([a-zA-Z]{2,3}\d{1})\s*')
    comeout = matcher.search(file)

    # deal non-regular file name
    # in which missing unit index
    while comeout == None:
        print('\nWarning -->')
        unit = input('\t' + file + ' 的机组信息获取失败。\n\t请在此输入(例如HYH3， YJ5等) --> ')
        comeout = matcher.search(unit)

        if comeout:
            break

    return comeout.group(1).lower()

def getExcelRows(file):
    '''
      the parameter *file* should be a tuple which
      consists of file_name, abs_path_with_file_name
    '''

    excel = win32.DispatchEx('Excel.Application')
    wb = excel.Workbooks.Open(file[1])
    ws = wb.Worksheets(1)

    # here we have to do some math to figure out
    # how many rows are actually in this workbook
    countRows = 0
    for i in range(3, 10000):
        if ws.Range('J' + str(i)).Value == None and \
           ws.Range('M' + str(i)).Value == None :
            break
        countRows += 1
    # here we have to do some math, again, to figure out
    # range of actual rows and columns we need
    print(countRows)
    rows = ws.Range('A3', 'AZ'+str(2 + countRows)).Value
    print(len(rows))
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
    rows = getExcelRows(file)

    Jname = os.path.splitext(file[1])[0] + ".json"

    with open(Jname, 'w', encoding='utf-8') as f:

        # extract the unit nubmer from file name
        # If the unit number is not specified in file name
        # getUnit function will ask for a unit number from user.
        unit = getUnit(file[0])

        # check if the IO list is any one of these types
        # AIO, DIO(16DO), PIF
        header = headers[ getHeader(file[0]) ]

        # first, let's write some start points to the json file
        f.write('[\n')

        # starts to parse the rows of IO List
        count = 0
        for row in rows:
            count += 1
            aDic = {}
            for column in header[1:]:

                # get cell value in the tuple of row
                # if the value is float, turn it to string
                cellValue = row[ header.index(column) ] 
                if type(cellValue) == type(1.0):
                    cellValue = str(int(cellValue))

                # write the value to aDic
                # aDic represents a row of IO
                # and will be converted to a json object later 
                if '.' in column:
                    # here the format of key might like
                    # this: parent.child
                    # we would like to split parent as sub1 and child
                    sub1, sub2 = column.split('.')
                    aDic.setdefault(sub1, {})
                    # py32com read number as float, such as 1.0, 3.0
                    # We'd like to make it a string
                    aDic[sub1][sub2] = cellValue
                else: 
                    aDic[column] = cellValue

            # last step of generating aDic:
            # write the unit number
            aDic[header[0]] = unit
            aJson = json.dumps(aDic, ensure_ascii=False);

            # Now let's wirte the row object
            # to the *.json
            # When it is at the last row, exclude the comma
            if count == len(rows):  # is the last row
                f.write(aJson)
            else:  # not the last row
                f.write(aJson + ',' + '\n')

        print(len(rows), count)
        f.write('\n]')

if __name__ == '__main__':

    excelFiles = getArgvFile(sys.argv[1:])
    for file in excelFiles:
        # here file should be a 
        # tuple (file_name, abs_path_with_file_name)
        Jgenerator(file)

