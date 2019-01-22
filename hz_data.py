# -*- coding: utf-8 -*-

import xlrd
import xlwt

path_read = r'E:\data1\hz_data\201901.xlsx'

data = xlrd.open_workbook(path_read)
sheets = data.sheets()


for sheet in sheets:
    sheet_name = sheet.name

    def finder(str):
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                cell = sheet.cell(row, col).value
                if cell == str:
                    return (row,col)


    huixuan_netvalue = sheet.cell_value(finder("华洲慧选")[0],finder("产品份额")[1] + 1)
    wenying9_netvalue = sheet.cell_value(finder("稳赢九期")[0],finder("产品份额")[1] + 1)

    huixuan_account_value = sheet.cell_value(finder("华洲慧选")[0],finder("账户资产余额")[1])
    wenying9_account_value = sheet.cell_value(finder("稳赢九期")[0],finder("账户资产余额")[1])

    huixuan_outincash = sheet.cell_value(finder("华洲慧选")[0],finder("出入金")[1])
    wenying9_outincash = sheet.cell_value(finder("稳赢九期")[0],finder("出入金")[1])
    if huixuan_outincash == "":
        huixuan_outincash = 0.0
    if wenying9_outincash == "":
        wenying9_outincash = 0.0

    huixuan_future_value = sheet.cell_value(finder("华洲慧选")[0],finder("期货")[1])
    wenying9_future_value = sheet.cell_value(finder("稳赢九期")[0],finder("期货")[1])

    huixuan_stock_value = sheet.cell_value(finder("华洲慧选")[0],finder("股票")[1])
    wenying9_stock_value = sheet.cell_value(finder("稳赢九期")[0],finder("股票")[1])

    print(sheet_name,huixuan_netvalue,huixuan_account_value,huixuan_outincash,huixuan_future_value,huixuan_stock_value)
    # print(sheet_name,wenying9_netvalue,wenying9_account_value,wenying9_outincash,wenying9_future_value,wenying9_stock_value)