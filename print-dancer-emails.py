import xlrd
import re
import random

workbook = xlrd.open_workbook(r'/Users/i505432/Downloads/2022 Dance Gala 报名表.xls')
worksheet = workbook.sheet_by_name(r'工作表1')

dance_type_dict = dict()
dance_dancers_dict = dict()

num_rows = worksheet.nrows
curr_row = 2
while curr_row < num_rows:
    row = worksheet.row(curr_row)
    key = worksheet.cell(curr_row,1).value
    dance_type_dict[key] = worksheet.cell(curr_row,4).value

    dancers = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", worksheet.cell(curr_row,3).value.lower())
    dance_dancers_dict[key] = dancers
    # print(key, type_dict[key], dance_dancers_dict[key])
    curr_row += 1

dancer_dances_dict = dict()
for key in dance_dancers_dict:
    dancers = dance_dancers_dict[key]
    for dancer in dancers:
        if dancer_dances_dict.get(dancer) is None:
            dancer_dances_dict[dancer] = [key]
        else:
            dancer_dances_dict[dancer].append(key)



# print all dancers (unique)
dancers_list = []
for key in dance_dancers_dict:
    dancers_list.extend(dance_dancers_dict[key])
dancers_set = set(dancers_list)

for d in dancers_set:
    print(d)


# print dancers and their number of dances and their dance (from more to less)
# dancer_dances_count_dict = dict()
# for dancer in dancer_dances_dict:
#     dancer_dances_count_dict[dancer] = len(dancer_dances_dict[dancer])
# items = dancer_dances_count_dict. items()
# sorted_items = sorted(items)
# for dancer, num_dances in sorted_items:
#     print(dancer, num_dances, dancer_dances_dict[dancer])