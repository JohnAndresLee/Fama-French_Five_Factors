import numpy as np
import pandas as pd

inputfile = '../temp/stkmnth data.xlsx'
stk_data = pd.read_excel(inputfile)
# 获得去重后的日期
stk_date = stk_data.drop_duplicates(['Trdmnt'])['Trdmnt']
stk_date = stk_date[0:31]

inputpath = '../temp/size_data.xlsx'
size_data = pd.read_excel(inputpath)

size_data = size_data.iloc[:, 0:40]

size_data = size_data.sort_values(by='Msmvosd', ascending=True)
number = len(size_data)
size_data.Size[0:int(0.2 * number)] = 1
size_data.Size[int(0.2 * number):int(0.4 * number)] = 2
size_data.Size[int(0.4 * number):int(0.6 * number)] = 3
size_data.Size[int(0.6 * number):int(0.8 * number)] = 4
size_data.Size[int(0.8 * number):] = 5

size_data = size_data.sort_values(by='BM', ascending=True)
number = len(size_data)
size_data.bm[0:int(0.2 * number)] = 1
size_data.bm[int(0.2 * number):int(0.4 * number)] = 2
size_data.bm[int(0.4 * number):int(0.6 * number)] = 3
size_data.bm[int(0.6 * number):int(0.8 * number)] = 4
size_data.bm[int(0.8 * number):] = 5

size_data = size_data.sort_values(by='inv', ascending=True)
number = len(size_data)
size_data.Inv[0:int(0.2 * number)] = 1
size_data.Inv[int(0.2 * number):int(0.4 * number)] = 2
size_data.Inv[int(0.4 * number):int(0.6 * number)] = 3
size_data.Inv[int(0.6 * number):int(0.8 * number)] = 4
size_data.Inv[int(0.8 * number):] = 5

size_data = size_data.sort_values(by='OP', ascending=True)
number = len(size_data)
size_data.op[0:int(0.2 * number)] = 1
size_data.op[int(0.2 * number):int(0.4 * number)] = 2
size_data.op[int(0.4 * number):int(0.6 * number)] = 3
size_data.op[int(0.6 * number):int(0.8 * number)] = 4
size_data.op[int(0.8 * number):] = 5

size_bm = size_data.drop(['Msmvosd', 'BM', 'inv', 'Inv', 'OP', 'op'], axis=1)
size_inv = size_data.drop(['Msmvosd', 'BM', 'bm', 'inv', 'OP', 'op'], axis=1)
size_op = size_data.drop(['Msmvosd', 'BM', 'inv', 'Inv', 'bm', 'OP'], axis=1)

# means_size_bm = np.zeros((5, 5))
# for i in range(1, 6):
#     for j in range(1, 6):
#         size_temp = size_bm[(size_bm['Size'] == i) & (size_bm['bm'] == j)]
#         size_temp = size_temp.drop(['Stkcd', 'Size', 'bm'], axis=1)
#         xs = size_temp.iloc[:, :].values.mean()
#         means_size_bm[i-1, j-1] = xs
#
# means_size_inv = np.zeros((5, 5))
# for i in range(1, 6):
#     for j in range(1, 6):
#         size_temp = size_inv[(size_inv['Size'] == i) & (size_inv['Inv'] == j)]
#         size_temp = size_temp.drop(['Stkcd', 'Size', 'Inv'], axis=1)
#         xs = size_temp.iloc[:, :].values.mean()
#         means_size_inv[i-1, j-1] = xs
#
# means_size_op = np.zeros((5, 5))
# for i in range(1, 6):
#     for j in range(1, 6):
#         size_temp = size_op[(size_op['Size'] == i) & (size_op['op'] == j)]
#         size_temp = size_temp.drop(['Stkcd', 'Size', 'op'], axis=1)
#         xs = size_temp.iloc[:, :].values.mean()
#         means_size_op[i-1, j-1] = xs

size_bmreg = {'Time': stk_date}
size_bmreg = pd.DataFrame(size_bmreg)

for i in range(1, 6):
    for j in range(1, 6):
        data_temp = []
        str_title = 'Size' + str(i) + '_' + 'BM' + str(j)
        for date in stk_date:
            size_temp = size_bm[['Size', 'bm', date]]
            size_temp = size_temp[(size_temp['Size'] == i) & (size_temp['bm'] == j)]
            data_temp.append(size_temp[date].values.mean())
        size_bmreg[str_title] = data_temp

outputpath = '../temp/25combine/2014-2016Size-BM.xlsx'
size_bmreg.to_excel(outputpath, index=False, header=True)

size_opreg = {'Time': stk_date}
size_opreg = pd.DataFrame(size_opreg)

for i in range(1, 6):
    for j in range(1, 6):
        data_temp = []
        str_title = 'Size' + str(i) + '_' + 'OP' + str(j)
        for date in stk_date:
            size_temp = size_op[['Size', 'op', date]]
            size_temp = size_temp[(size_temp['Size'] == i) & (size_temp['op'] == j)]
            data_temp.append(size_temp[date].values.mean())
        size_opreg[str_title] = data_temp

outputpath = '../temp/25combine/2014-2016Size-OP.xlsx'
size_opreg.to_excel(outputpath, index=False, header=True)

size_invreg = {'Time': stk_date}
size_invreg = pd.DataFrame(size_invreg)
for i in range(1, 6):
    for j in range(1, 6):
        data_temp = []
        str_title = 'Size' + str(i) + '_' + 'Inv' + str(j)
        for date in stk_date:
            size_temp = size_inv[['Size', 'Inv', date]]
            size_temp = size_temp[(size_temp['Size'] == i) & (size_temp['Inv'] == j)]
            data_temp.append(size_temp[date].values.mean())
        size_invreg[str_title] = data_temp

outputpath = '../temp/25combine/2014-2016Size-INV.xlsx'
size_invreg.to_excel(outputpath, index=False, header=True)


print("ok")
