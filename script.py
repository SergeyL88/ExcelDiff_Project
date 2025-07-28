from csv import excel

import pandas as pd
import jinja2
from numpy.f2py.auxfuncs import islong_complex
from openpyxl.descriptors import String
from openpyxl import Workbook
from pandas.core.computation.common import result_type_many

PATH = "/Users/sergej/Documents/ExcelTest/"

def check_percent(data1, data2):
    result = 0
    try:
        result = (((data1 / data2 )-1) * 100)
    except ZeroDivisionError:
        # print("Rais Zero Division")
        return 0
    except ValueError:
        # print("Rais Value Error")
        return 0
    return result

def convert_to_dict(pd_array: pd.DataFrame) -> dict:
    result_dict = {}
    for index, row in pd_array.iterrows():
        result_dict[index] = [pd_array.iloc[index, 0], pd_array.fillna(0).iloc[index, 7]]

    return result_dict

def equal_products(doc1, doc2):
    data1_dict = convert_to_dict(doc1)
    data2_dict = convert_to_dict(doc2)
    sheet_name = ''

    for index1, row1 in list(data1_dict.items()):
        for index2, row2 in list(data2_dict.items()):
            if row1[0] in exclude_list:
                break
            elif row1[0] == row2[0]:
                # print(f"Row name: {row1[0]} = Value: {row1[1]} \t Row name: {row2[0]} = Value: {row2[1]}")
                data2_dict.pop(index2)
                if row1[0] == "Цех №1":
                    sheet_name = "Цех №1"
                    print(f"sheet name: {sheet_name}")
                elif row1[0] == "Цех №2":
                    sheet_name = "Цех №2"
                    print(f"sheet name: {sheet_name}")
                elif row1[0] == "Цех №3":
                    sheet_name = "Цех №3"
                    print(f"sheet name: {sheet_name}")
                elif row1[0] == "Аутсорсинг":
                    sheet_name = "Аутсорсинг"
                    print(f"sheet name: {sheet_name}")

                if sheet_name == "Цех №1":
                    try:
                        mf1_list.append([row1[0], row2[1], row1[1], row2[1] - row1[1], check_percent(row2[1], row1[1])])
                        break
                    except TypeError as ex:
                        print(ex)
                elif  sheet_name == "Цех №2":
                    try:
                        mf2_list.append([row1[0], row2[1], row1[1], row2[1] - row1[1], check_percent(row2[1], row1[1])])
                        break
                    except TypeError as ex:
                        print(ex)
                elif sheet_name == "Цех №3":
                    try:
                        mf3_list.append([row1[0], row2[1], row1[1], row2[1] - row1[1], check_percent(row2[1], row1[1])])
                        break
                    except TypeError as ex:
                        print(ex)
                else:
                    try:
                        outsource_list.append([row1[0], row2[1], row1[1], row2[1] - row1[1], check_percent(row2[1], row1[1])])
                        break
                    except TypeError as ex:
                        print(ex)
            continue

def not_equal_products(data1, data2):
    result_dict = {}
    for index1, row in data1.iterrows():
        for index2, col in data2.iterrows():
            if data1.iloc[index1,0] != data2.iloc[index2,0]:
                print(data1.iloc[index1,0])

def highlight(s):
    if s == 'Цех №1':
        return ['background-color: yellow'] * len(s)
    else:
        return ['background-color: white'] * len(s)

doc1_name = "Остаток ГП и ПФ на 01.07.24.xlsx"
doc2_name = "Остаток ГП и ПФ на 01.07.25.xlsx"
sheet_name = ''
exclude_list = ["Оценка избыточности запасов ПФ и ГП", "Параметры:", "Отбор:", "Цех основной продукции"]
data_columns = ["Наименование", doc2_name[0:-5], doc1_name[0:-5], "Разница", "Процент"]
excel_colors_list = ["Цех №1" , "Цех №2", "Цех №3"]
mf1_list = []
mf2_list = []
mf3_list = []
outsource_list = []


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    doc1 = pd.read_excel(PATH + doc1_name)
    doc2 = pd.read_excel(PATH + doc2_name)

    changes = doc2.fillna(0).iloc[15:,7] - doc1.fillna(0).iloc[15:,7]
    changes_pct = check_percent(doc1.iloc[15:,7] , doc2.iloc[15:,7])

    equal_products(doc1, doc2)

    pd_mf1 = pd.DataFrame(mf1_list, columns=data_columns)
    pd_mf1.style.apply(highlight, data_columns={'Наименование': "Наименование"}, axis=0)
    pd_mf2 = pd.DataFrame(mf2_list, columns=data_columns)
    pd_mf3 = pd.DataFrame(mf3_list, columns=data_columns)
    pd_outsource = pd.DataFrame(outsource_list, columns=data_columns)
    print(pd_mf1.iloc[0:,])

    with pd.ExcelWriter(PATH + 'TestChanges2.xlsx') as writer:
        pd_mf1.to_excel(writer, index=False, sheet_name="Цех 1", float_format='%.2f')
        pd_mf2.to_excel(writer, index=False, sheet_name="Цех 2", float_format='%.2f')
        pd_mf3.to_excel(writer, index=False, sheet_name="Цех 3", float_format='%.2f')
        pd_outsource.to_excel(writer, index=False, sheet_name="Аутсорс", float_format='%.2f')
