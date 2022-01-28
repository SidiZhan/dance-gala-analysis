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
    dance = worksheet.cell(curr_row,4).value
    if len(dance.strip()) > 0:
        dance_type_dict[key] = dance

    dancers = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", worksheet.cell(curr_row,3).value.lower())
    dance_dancers_dict[key] = dancers
    # print(key, dance_type_dict[key], dance_dancers_dict[key])
    curr_row += 1

dances = dance_type_dict.keys()
print('节目数量', len(dances))

dancers_list = []
for key in dance_dancers_dict:
    dancers_list.extend(dance_dancers_dict[key])
print('表演者人次', len(dancers_list))
print('表演者人数（去重）', len(set(dancers_list)))



dancer_dances_dict = dict()
for key in dance_dancers_dict:
    dancers = dance_dancers_dict[key]
    for dancer in dancers:
        if dancer_dances_dict.get(dancer) is None:
            dancer_dances_dict[dancer] = [key]
        else:
            dancer_dances_dict[dancer].append(key)

dancer_dances_count_dict = dict()
for key in dancer_dances_dict:
    num = len(dancer_dances_dict[key])
    if dancer_dances_count_dict.get(num) is None:
        dancer_dances_count_dict[num] = 1
    else:
        dancer_dances_count_dict[num] = dancer_dances_count_dict[num] + 1
print('单人表演节目数量：')
for key in dancer_dances_count_dict:
    print('-', '同时参与了', key, '个节目的表演者有', dancer_dances_count_dict[key], '位')


print('== 节目单 ==')
# 1. dancer has > 3 dances between her two dances
# 2. two nearby dances are of differnt types


def common_dancer(d1, d2):
    l1 = dance_dancers_dict[d1]
    l2 = dance_dancers_dict[d2]
    l3 = []
    l3.extend(l1)
    l3.extend(l2)
    s = set(l3)
    if len(s) == len(l3):
        return None
    for dancer in s:
        if len(dancer_dances_dict[dancer]) > 1:
            if l1.count(dancer) > 0 and l2.count(dancer) > 0:
                return dancer


def check(l, e):
    # check if the element can be appended to the list, can = True, cannot = False
    # l is the list of dances (string), e is a dance (string)

    i = len(l) - 1 # the index of the last element

    # check if the list is empty
    if i < 0:
        # print('the list is about to append its first element')
        return True

    # check type
    if dance_type_dict[l[i]] == dance_type_dict[e]:
        # print(e, l[i], 'new and last one: the type are the same', dance_type_dict[l[i]])
        return False
    
    # check if it conflicts with last one
    if common_dancer(e, l[i]) is not None:
        # print(e, l[i], 'new and last one: they have a common dancer', common_dancer(e, l[i]))
        return False


    # check if it conflicts with last two
    if i-1 >= 0:
        if common_dancer(e, l[i-1]) is not None:
            # print(e, l[i-1], 'new and last two: they have a common dancer', common_dancer(e, l[i-1]))
            return False


    # check if it conflicts with last three
    if i-2 >= 0:
        if common_dancer(e, l[i-2]) is not None:
            # print(e, l[i-2], 'new and last three: they have a common dancer', common_dancer(e, l[i-2]))
            return False


    # now check with the special requirements 

    # specific requirements: "Run the world" and "BNZ&FIGUB" dist >= 5: dist() = int
    if i-4 >= 0:
        if (e == 'Run the world' and l[i-4] == 'BNZ&FIGUB') or (l[i-4] == 'Run the world' and e == 'BNZ&FIGUB'):
            return False


    # specific requirements: "朝鲜舞" shall be after other "meihua.han@sap.com" dances
    if e == '朝鲜舞':
        dancer = 'meihua.han@sap.com'
        d_count = len(dancer_dances_dict[dancer]) - 1 # number of other dances
        count = 0
        for dance in l:
            if list(dance_dancers_dict[dance]).count(dancer) > 0:
                count = count + 1
        if count != d_count:
            return False

    # specific requirements: '赤伶' shall be before other "tina.chen03@sap.com" dances
    if e == '赤伶':
        dancer = 'tina.chen03@sap.com'
        for dance in l:
            if list(dance_dancers_dict[dance]).count(dancer) > 0:
                return False

    return True


original_order = list(dance_type_dict.keys())
last_dance = original_order.pop() # singing
random.shuffle(original_order)
order_done = False
for i in range(len(original_order)):
    dances = original_order.copy()
    order = []
    order.append(dances.pop())
    old_order_len = len(order)
    while len(dances) > 0:
        old_order_len = len(order)
        for j in range(len(dances)):
            dance = dances.pop()
            if check(order, dance) is True:
                order.append(dance)
            else:
                dances.insert(0,dance)
        if old_order_len == len(order):
            # print(i,j,'===== you have checked all the dances in this position and no fit')
            original_order.insert(0, original_order.pop())
            break
        
    if len(dances) == 0:
        order_done = True
        break


order.append(last_dance)

if order_done is False:
    print('cannot generate order')
else:
    print('ordered')
    for i in range(len(order)):
        print(i+1, order[i], dance_type_dict[order[i]], dance_dancers_dict[order[i]])