# -*- coding: utf-8 -*-
"""
Spyder Editor

Author : Akash
Date : 29/02/2020
Purpose : Algorithm to derive Wilcoxon signed rank test
"""

import itertools,xlrd

print 'Enter Key Column Index :'
key_col=input()

sheet_cols=[]
all_perms=[]
col_indexes={}


# get permutations from file columns
def col_permutations(file_cols):
    a=file_cols #[25,7,8,9]
    for num in range (2,5):
        all_perms.extend(itertools.combinations(a,num))
    print 'printing permutations of all non-key columns',all_perms
    return all_perms

# read excel file
def read_excel():
    book=xlrd.open_workbook("C:/Users/acer/Downloads/Akash/test/Python_test/Algo_Test.xlsx",'rb')        
    sheet_obj=book.sheet_by_name('Data')
    for col_ind in range(key_col,sheet_obj.ncols):
        col_val=int(sheet_obj.cell_value(0,col_ind))
        col_indexes.update({col_val:col_ind})
        sheet_cols.append(col_val)
    print sheet_cols,col_indexes
    return sheet_obj

rank_dict={}
# function to generate ranks
def rank_calculate(sheet_ob,key_col,non_key_col):
    for vk,v_col_ind in sorted(col_indexes.items()):#values()):
        print v_col_ind
        diff_list=[]
        conv_list=[]
        rank_list=[]
        if v_col_ind!=key_col:
            for ind in range(sheet_ob.nrows):
                comp_val1=int(sheet_ob.row(ind)[key_col].value)
                comp_val2=int(sheet_ob.row(ind)[v_col_ind].value)
                diff_list.append(comp_val1-comp_val2)
            # converted dataset
            conv_list=[abs(v) for v in diff_list]
            
            # find ranks
            rank_list=list(enumerate(conv_list))
            
            k=str(key_col)+'_'+str(vk)
            # finding signed ranks
            
            
            rank_dict.update({k:rank_list})
            #break
        
        print rank_dict
        


sheet_ob=read_excel()
col_permutations(sheet_cols[key_col+1:])
rank_calculate(sheet_ob,key_col,sheet_cols[key_col+1:])
