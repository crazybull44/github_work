# -*- coding: utf-8 -*-

import os
import datetime as dt
import pandas as pd
import xlrd

def file_rename(filepath):
    #filepath = r'E:\data1\gtxt\6黎启君'
    i = 0
    for root, dirs, files in os.walk(filepath):
        for name in files:
            i = i + 1
            old_dir_name = os.path.join(root, name)
            print(old_dir_name)
            print(os.path.getmtime(old_dir_name))
            get_file_mtime = dt.datetime.fromtimestamp(os.path.getmtime(old_dir_name)).strftime('%Y%m%d%H%M%S')
            print(get_file_mtime)
            filetype = os.path.splitext(name)[1]
            new_dir_name = os.path.join(root, get_file_mtime + str(i) + filetype)
            print(new_dir_name)
            os.rename(old_dir_name, new_dir_name)
            print('%s-->>%s' % (old_dir_name, new_dir_name))
            print('-' * 40)
    print('All files are renamed success!')


def get_fund_symbol_dict(filepath):
    # filepath = r'E:\data1\gtxt\6黎启君'
    res_dict = {}
    for root, dirs, files in os.walk(filepath):
        for name in files:
            dir_name = os.path.join(root, name)
            if os.path.splitext(name)[1] == ".XLS":
                open_file = xlrd.open_workbook(dir_name, encoding_override='gbk')
                cell = open_file.sheets()[0].cell(0, 2).value
                if cell != '证券代码':
                    continue
                df = pd.read_excel(open_file, engine='xlrd', encoding='utf-8', dtype=str)
                symbol_list = df['证券代码'].tolist()
                symbol_list.remove('nan')
                date_str = os.path.splitext(name)[0][:8]
                res_dict[date_str] = symbol_list
    print(res_dict)
    print(len(res_dict))
    print('get fund everyday symbol success!')

if __name__ == "__main__":
    filepath = r'E:\data1\gtxt\18启林'
    file_rename(filepath)
    get_fund_symbol_dict(filepath)
