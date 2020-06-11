import json
import xlrd
# xlrd是读excel，xlwt是写excel的库
from xlutils.copy import copy
# 原始excel文件的编辑功能 copy
import jieba
# 分词
with open('data/province.json', 'r', encoding='utf-8') as f:
    province = json.load(f)
with open('data/city.json', 'r', encoding='utf-8') as f:
    city = json.load(f)
with open('data/country.json', 'r', encoding='utf-8') as f:
    country = json.load(f)

old = 'data/report--2019.xls'
new = 'data/new--2019.xls'

wb = xlrd.open_workbook(old)
ws = wb.sheets()[0]
nb = copy(wb)
ns = nb.get_sheet(0)

H = ws.col_values(7)
M = ws.col_values(12)

# 手动处理队列
manual = []

i = 1
while i < ws.nrows:

    h = H[i]
    m = M[i]
    print(i, end=' ')

    if h == '':
        print()
        i += 1
        continue

    def mod(loc, col):
        print(loc, end=' -> ')
        seg_list = jieba.lcut(loc)
        # 分词结果为单，若为省、国则不处理，若为市则转化为对应省
        if len(seg_list) == 1:
            # 单词为国，不处理
            if seg_list[0] in country:
                result = seg_list[0]
            # 单词为省，不处理
            elif seg_list[0] in province:
                result = seg_list[0]
            # 单词为市，替换为对应省 !!!!!!!!!!
            elif seg_list[0] in city:
                result = city[loc]
                ns.write(i, col, result)
            # 单词无法识别，手动确认 !!!!!
            else:
                manual.append(i)
                result = ''
            print(result)

        # 分词结果为多
        if len(seg_list) > 1:
            # 若第一个为省，则保留
            if seg_list[0] in province:
                result = seg_list[0]
                ns.write(i, col, result)
            # 若第二个为省，则保留
            elif seg_list[1] in province:
                result = seg_list[1]
                ns.write(i, col, result)
            # 单词无法识别，手动确认 !!!!!
            else:
                manual.append(i)
                result = ''
            print(result)

    mod(h, 7)
    mod(m, 12)

    i += 1

manual = list(set(manual))
manual.sort()

print('以下行的数据需要手动处理')
for index in manual:
    print(index+1, H[index], M[index])
print('总计需要处理: ' + str(len(manual)))

nb.save(new)



