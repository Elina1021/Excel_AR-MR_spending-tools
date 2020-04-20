#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import xlrd
import pandas as pd
import xlwings as xw
import openpyxl
from openpyxl import load_workbook
from collections import namedtuple


# ### 函数

# In[2]:


def vlookup_combine(df_left, df_right, key_left, key_right, lookup_right):
    '''
    返回df: 查找列key_left; 匹配列key_right; 匹配值lookup_right
    '''
    v_left = df_left[[key_left]].astype('object')
    v_left.columns = ['key_left']
    # 如果匹配列名字相同，手工增加一列
    if key_right == lookup_right:
        v_right = df_right[[key_right]].astype('object')
        v_right[key_right+lookup_right] = v_right[key_right]
    else:
        v_right = df_right[[key_right, lookup_right]].astype('object')
    v_right.columns = ['key_right', 'lookup_right']
    v_right.drop_duplicates(inplace=True)
    v_combine = v_left.merge(v_right, how='left', left_on=[
                             'key_left'], right_on=['key_right'])
    return v_combine


def vlookup_not_exist_key(v_combine):
    '''
    返回list: v_combine中未能匹配的'key_left'
    '''
    v_not_exist = v_combine[v_combine['key_right'].isna()]['key_left']
    list_not_exist = v_not_exist.drop_duplicates().to_list()
    return list_not_exist


def vlookup(df_left, df_right, key_left, key_right, lookup_right, default=None):
    '''
    返回字典: 
    1. 如果有不能匹配的,返回{None:不能匹配的列表}
    2. 如果都能匹配, 返回{True:匹配列表}
    '''
    v_combine = vlookup_combine(
        df_left, df_right, key_left, key_right, lookup_right)
    if default is None:
        not_exist_key = vlookup_not_exist_key(v_combine)
        if not_exist_key:
            return {None: not_exist_key}
        else:
            return {True: v_combine['lookup_right'].to_list()}
    else:
        v_combine['lookup_right'] = v_combine['lookup_right'].fillna(default)
        return {True: v_combine['lookup_right'].to_list()}


def xw_get_book(filename=None):
    if filename is None:
        book = xw.Book.caller()
    else:
        try:
            book = xw.Book(filename)
        except FileNotFoundError:
            book = xw.Book()
            book.save(filename)
    return book


def xw_get_df(book, sheet_name):
    '''
    获取sheet_name数据
    '''
    sheet = book.sheets(sheet_name)
    # 需要再调整
    data = sheet.range((1, 1), sheet.used_range.shape)
    df = pd.DataFrame(data.value[1:], columns=data.value[0])
    return df


def xw_fill_sheet_with_df(cell, df):
    cell.value = df.columns.values
    cell.offset(1, 0).value = df.values


def clear_and_fill_sheet(book, sheet_name, df):
    sheet = book.sheets(sheet_name)
    sheet.clear_contents()
    cell = sheet.cells(1, 1)
    xw_fill_sheet_with_df(cell, df)
    return book


# ### main

# In[3]:


def main(filename_data=None):
    # 1.读取数据
    book = xw_get_book(filename_data)
    # 1.1 获取原始数据
    sheet_name_data = 'AR_Data'
    df_data = xw_get_df(book,sheet_name_data)
    # 1.2 读取list
    sheet_name_right = 'list'
    df_right = xw_get_df(book,sheet_name_right)

    # 2.读取转换后的当月花费数据
    filename_update = 'currentdata.txt'
    df_left = pd.read_csv(filename_update,sep='\t')

    # 3.数据处理
    VReplace = namedtuple('VReplace',
                          ['key_left', 'key_right', 'lookup_right', 'default']
                          )

    multi_processes = [
        # 3.1 替换品牌
        (VReplace('品牌', '品牌', 'Brand_EN', 'OTHER')),
        # 3.2 替换车型_key
        (VReplace('车型', '车型_key', '车型_key', 'OTHER')),
        # 3.3 替换AR_Catetory
        (VReplace('车型', '车型_key', 'AR_Category', 'OTHER')),
        # 3.4 替换AR_Model
        (VReplace('车型', '车型_key', 'AR_Model', 'OTHER')),
    ]

    for key_left, key_right, lookup_right, default in multi_processes:
        df_v = vlookup(df_left, df_right, key_left, key_right, lookup_right,default)
        df_left[lookup_right] = df_v[True]

    # 3.5 折算花费（百万）
    df_left['折算花费（百万）'] = df_left['折算花费（万元）']/100

    # 4 获取结果
    df_res = pd.concat([df_data,df_left])
    df_res.drop_duplicates(inplace=True)

    # 5. 填充数据
    book = clear_and_fill_sheet(book, sheet_name_data, df_res)
    book.save()


# In[51]:


if __name__ == '__main__':
    filename_data = 'spending_data_test.xlsx'
    main(filename_data)

