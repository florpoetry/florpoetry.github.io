import pandas as pd
from toolz import unique
import xlsxwriter
# This script is for basic usage of organization of files that need to be sorted
def doAll():
    xls = pd.ExcelFile('trial.xlsx')
    df1 = pd.read_excel(xls, 'Sheet1')
    number_ = seriesToList(df1["E_Gen"])
    number_set_list = list(set(number_))
    num_dict = {}
    for unique in number_set_list:
        arr =[]
        num_dict[str(unique)] = arr
    row_size = df1.shape[0]
    all_data = []
    for number in number_set_list:
        curr_data = (df1.loc[df1['E_Gen'] == number])
        ret_val = loadNewSheet(curr_data)
        all_data.append(ret_val)
    final_result = pd.concat(all_data)
    export_excel = final_result.to_excel ("final_result.xlsx", index = None, header=True) #Don't forget to add '.xlsx' at the end of the path

def loadNewSheet(df1):
    p_name = seriesToList(df1["P_Item"])
    p_code = seriesToList(df1["P_Code"])
    p_val = seriesToList(df1["P_Val"])
    dict = getDict(p_name, p_code, p_val)
    single_list = dictToListSingle(dict)
    arr_ultra = []
    for item in single_list:
        arr_ultra.append((df1.iloc[[item]]))
    results = pd.concat(arr_ultra)
    return results
def dictToListSingle(dict):
    single_list = []
    for key in dict.keys():
        curr_list = dict[key]
        for item in curr_list:
            single_list.append(item)
    return single_list

def seriesToList(series_):
    list_ = series_.tolist()
    return list_
def getDict(p_name, p_code, p_val):
    bool = (len(p_name) == len(p_code) == len(p_val))
    dict = {}
    if (bool == True):
        loop_size = len(p_name)
        for i in range(loop_size):
            arr = []
            item_One = p_name[i]
            item_Two = p_code[i]
            item_Three = p_val[i]
            arr.append(item_One)
            arr.append(item_Two)
            arr.append(item_Three)
            string_i = str(i)
            dict[i] = arr
    new_dict = dictToList(dict)
    return new_dict
def dictToList(dict):
    len_dict = len(dict)
    list_of_lists = []
    for key in dict.keys():
        list_of_lists.append(dict[key])
    res = map(list, unique(map(tuple, list_of_lists)))
    list_res = list(res)
    dict = indexPos(list_res, list_of_lists)
    return dict
def indexPos(unique_lists, all_lists):
    dict = {}
    for i in range(len(unique_lists)):
        str_i = str(i)
        arr = []
        dict[str_i] = arr
    for list_O in range (len(all_lists)):
        for list_T in range (len(unique_lists)):
            if (all_lists[list_O] == unique_lists[list_T]):
                arr = dict[str(list_T)]
                new_arr = arr
                new_arr.append(list_O)
                dict[str(list_T)] = new_arr
    return dict
doAll()
