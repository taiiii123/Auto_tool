import os
from datetime import datetime

d = datetime.today()
date = d.strftime('%Y%m%d')

# デフォルト(番号 + フォルダ名)
def mkdir_items_num(dir_path, mkdir_list=None):
    for idx, item in enumerate(mkdir_list):
        if idx < 9:
            os.mkdir(dir_path + '/' + '0{}{}'.format(idx+1, item))
        else:
            os.mkdir(dir_path + '/' + '{}{}'.format(idx+1, item))

# オプション有り(フォルダ名 + 日付)
def mkdir_items_date(dir_path, mkdir_list=None):
    for i in range(len(mkdir_list)):
        os.mkdir(dir_path + '/' + mkdir_list[i] + date)

# オプションなし(フォルダ名)
def mkdir_items_no_slc(dir_path, mkdir_list=None):
    for i in range(len(mkdir_list)):
        os.mkdir(dir_path + '/' + mkdir_list[i])

# オプションを有り(番号 + フォルダ名 + 日付)
def mkdir_items_all_slc(dir_path, mkdir_list=None):
    for idx, item in enumerate(mkdir_list):
        if idx < 9:
            os.mkdir(dir_path + '/' + '0{}{}'.format(idx+1, item) + date)
        else:
            os.mkdir(dir_path + '/' + '{}{}'.format(idx+1, item) + date)
