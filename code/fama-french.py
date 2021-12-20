import pandas as pd
import numpy as np
import datetime, time
import dateutil
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker


def xls_read(n):
    inputfile = '../data/ff5f_data/Balance sheet/FS_Combas' + str(n) + '.xlsx'
    data = pd.read_excel(inputfile, index_col=0)

    inputfile2 = '../data/ff5f_data/Income statement/FS_Comins' + str(n) + '.xlsx'
    data2 = pd.read_excel(inputfile2)

    # data.insert(6, 'operation profit', data2['operating profit'])
    return data, data2


def stk_adjust():
    inputfile = '../temp/stkmnth data.xlsx'
    stk_data = pd.read_excel(inputfile)
    # 获得去重后的股票代码
    stk_data_temp1 = stk_data
    stk_data_temp1 = stk_data_temp1.drop_duplicates(['Stkcd'])['Stkcd']

    # 获得去重后的日期
    stk_data_temp2 = stk_data
    stk_data_temp2 = stk_data_temp2.drop_duplicates(['Trdmnt'])['Trdmnt']

    stk_adj = pd.DataFrame(index=stk_data_temp1, columns=stk_data_temp2)
    k = 0
    for date in stk_data_temp2:
        if k == 0:
            date_data = stk_data[(stk_data['Trdmnt'] == date)][['Stkcd', 'Mretwd']]
            date_data.rename(columns={'Mretwd': date}, inplace=True)
            k += 1
        else:
            date_datatemp = stk_data[(stk_data['Trdmnt'] == date)][['Stkcd', 'Mretwd']]
            date_data = pd.merge(date_data, date_datatemp, on='Stkcd', how='outer')
            date_data.rename(columns={'Mretwd': date}, inplace=True)

    inputpath = '../temp/size_data.xlsx'
    size_data = pd.read_excel(inputpath)
    size_data = pd.merge(size_data, date_data, on='Stkcd', how='outer')

    size_data = size_data.dropna(axis=0, how='any')

    outputpath = '../temp/size_data.xlsx'
    size_data.to_excel(outputpath, index=True, header=True)
    print("ok")


def cal_factorappendix(size_data, stk_date, indicate, type1, type2):
    sh_data = size_data[(size_data['Size'] == type1) & (size_data[indicate] == type2)]
    sh_out = []
    for date in stk_date:
        sh = sh_data['Msmvosd'] * sh_data[date]
        sh = sh.sum()
        sh_sum = sh_data['Msmvosd'].sum()
        sh = sh / sh_sum
        sh_out.append(sh)

    title = type1 + type2 + '_' + indicate
    sh_data = {'Time': stk_date, title: sh_out}
    sh_data = pd.DataFrame(sh_data)
    return sh_data


# SH SN SL BH BN BL // SR SN SW BR BN BW // SC SN SA BC BN BA
def cal_factor(stk_date):
    inputpath = '../temp/size_data.xlsx'
    size_data = pd.read_excel(inputpath)

    sh_data = cal_factorappendix(size_data, stk_date, 'bm', 'S', 'H')
    temp = cal_factorappendix(size_data, stk_date, 'bm', 'S', 'N')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'bm', 'S', 'L')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'bm', 'B', 'H')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'bm', 'B', 'N')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'bm', 'B', 'L')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')

    temp = cal_factorappendix(size_data, stk_date, 'Inv', 'S', 'A')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'Inv', 'S', 'N')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'Inv', 'S', 'C')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'Inv', 'B', 'A')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'Inv', 'B', 'N')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'Inv', 'B', 'C')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')

    temp = cal_factorappendix(size_data, stk_date, 'op', 'S', 'R')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'op', 'S', 'N')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'op', 'S', 'W')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'op', 'B', 'R')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'op', 'B', 'N')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')
    temp = cal_factorappendix(size_data, stk_date, 'op', 'B', 'W')
    sh_data = pd.merge(sh_data, temp, on='Time', how='outer')

    outputpath = '../temp/factor_data.xlsx'
    sh_data.to_excel(outputpath, index=True, header=True)

    print("ok")


def cal_final(stk_date):
    inputpath = '../temp/factor_data.xlsx'
    factor_data = pd.read_excel(inputpath)

    SMB_BM = (factor_data['SH_bm'] + factor_data['SN_bm'] + factor_data['SL_bm'] - factor_data['BH_bm'] - factor_data[
        'BN_bm'] - factor_data['BL_bm']) / 3
    SMB_OP = (factor_data['SR_op'] + factor_data['SN_op'] + factor_data['SW_op'] - factor_data['BR_op'] - factor_data[
        'BN_op'] - factor_data['BW_op']) / 3
    SMB_INV = (factor_data['SC_Inv'] + factor_data['SN_Inv'] + factor_data['SA_Inv'] - factor_data['BC_Inv'] -
               factor_data['BN_Inv'] - factor_data['BA_Inv']) / 3
    SMB = (SMB_BM + SMB_OP + SMB_INV) / 3

    HML = (factor_data['SH_bm'] + factor_data['BH_bm'] - factor_data['SL_bm'] - factor_data['BL_bm']) / 2
    RMW = (factor_data['SR_op'] + factor_data['BR_op'] - factor_data['SW_op'] - factor_data['BW_op']) / 2
    CMA = (factor_data['SC_Inv'] + factor_data['BC_Inv'] - factor_data['SA_Inv'] - factor_data['BA_Inv']) / 2

    factor_out = {'Time': stk_date, 'SMB': SMB, 'HML': HML, 'RMW': RMW, 'CMA': CMA}
    factor_out = pd.DataFrame(factor_out)

    inputpath = '../temp/mktmnth data.xlsx'
    mktdata = pd.read_excel(inputpath)
    mktdata.rename(columns={'Trdmnt': 'Time', 'Cmretwdos': 'rm'}, inplace=True)

    factor_out = pd.merge(factor_out, mktdata, on='Time', how='outer')

    inputpath = '../data/ff5f_data/rf/TRD_Nrrate.xlsx'
    rfdata = pd.read_excel(inputpath)
    rfdata.rename(columns={'Clsdt': 'Time', 'Nrrmtdt': 'rf'}, inplace=True)

    rfdata['Time'] = pd.to_datetime(rfdata['Time'])
    rfdata = rfdata[(rfdata['Time'].dt.day == 1) | (
                (rfdata['Time'].dt.day == 30) & (rfdata['Time'].dt.month == 6) & (rfdata['Time'].dt.year == 2014))][
        'rf'].values

    rfdata = {'Time': stk_date, 'rf': rfdata[0:44]}
    rfdata = pd.DataFrame(rfdata)
    factor_out = pd.merge(factor_out, rfdata, on='Time', how='outer')

    factor_out['rm-rf'] = factor_out['rm'] - 0.01*factor_out['rf']

    outputpath = '../temp/five_factor.xlsx'
    factor_out.to_excel(outputpath, index=True, header=True)
    print("ok")


if __name__ == "__main__":
    [data, data2] = xls_read(1)
    for n in range(2, 4):
        [temp, temp2] = xls_read(n)
        data = pd.merge(data, temp, how='outer')
        data2 = pd.merge(data2, temp2, how='outer')

    data_final = pd.merge(data, data2, on=['Stkcd', 'Accper', 'Typrep'], how='outer')

    # debug = data_final[data_final.isnull().T.any()]

    data_final = data_final[-data_final.Typrep.isin(['B'])]

    data_final['Accper'] = pd.to_datetime(data_final['Accper'])
    data_final = data_final[(data_final['Accper'].dt.month == 6) | (data_final['Accper'].dt.month == 12)]

    data_final['Accper'] = data_final['Accper'].dt.date
    # debug = data_final[data_final.isnull().T.any()]
    # print(debug)

    outputpath = '../temp/financial data.xlsx'
    data_final.to_excel(outputpath, index=False, header=True)

    # 处理个股、市场数据
    inputfile3 = '../data/ff5f_data/stockmnth/TRD_Mnth0.xlsx'
    stkmnth_data = pd.read_excel(inputfile3)
    for n in range(1, 3):
        inputfile3 = '../data/ff5f_data/stockmnth/TRD_Mnth' + str(n) + '.xlsx'
        temp = pd.read_excel(inputfile3)
        stkmnth_data = pd.merge(stkmnth_data, temp, how='outer')

    stkmnth_data = stkmnth_data[(stkmnth_data['Markettype'] == 1) | (stkmnth_data['Markettype'] == 4)]

    stkmnth_data = stkmnth_data.sort_values(by=["Stkcd", "Trdmnt"], ascending=True)
    # 处理掉600087/000033/000982这几支异常股票
    stkmnth_data = stkmnth_data[-(stkmnth_data['Stkcd'] == 600087)]
    stkmnth_data = stkmnth_data[-(stkmnth_data['Stkcd'] == 33)]
    stkmnth_data = stkmnth_data[-(stkmnth_data['Stkcd'] == 982)]
    outputpath = '../temp/stkmnth data.xlsx'
    stkmnth_data.to_excel(outputpath, index=False, header=True)

    # 处理综合市场数据
    inputfile3 = '../data/ff5f_data/mktmnth/TRD_Cnmont1.xlsx'
    mktmnth_data = pd.read_excel(inputfile3)
    inputfile3 = '../data/ff5f_data/mktmnth/TRD_Cnmont.xlsx'
    temp = pd.read_excel(inputfile3)
    mktmnth_data = pd.merge(mktmnth_data, temp, how='outer')

    mktmnth_data = mktmnth_data[(mktmnth_data['Markettype'] == 5)]
    mktmnth_data = mktmnth_data[['Trdmnt', 'Cmretwdos']]

    outputpath = '../temp/mktmnth data.xlsx'
    mktmnth_data.to_excel(outputpath, index=False, header=True)

    # 要计算的指标：
    # “市值” (Size)指标是以股票i在第t年6月底的流通市值；
    inputfile = '../temp/stkmnth data.xlsx'
    stkmnth_data = pd.read_excel(inputfile)

    stksize_data = stkmnth_data[['Stkcd', 'Trdmnt', 'Msmvosd']]
    stksize_data['Trdmnt'] = pd.to_datetime(stkmnth_data['Trdmnt'])
    stksize_data = stksize_data[(stksize_data['Trdmnt'].dt.month == 6) & (
                (stksize_data['Trdmnt'].dt.year == 2014) | (stksize_data['Trdmnt'].dt.year == 2015) | (
                    stksize_data['Trdmnt'].dt.year == 2016))]
    stksize_data['Trdmnt'] = stksize_data['Trdmnt'].dt.date
    # 14 15 16年市值数据取平均
    Msmvosd_data = stksize_data.groupby(['Stkcd'])['Msmvosd'].mean()
    Msmvosd_data = Msmvosd_data.sort_values()

    outputpath = '../temp/size_data.xlsx'
    Msmvosd_data.to_excel(outputpath, index=True, header=True)

    stk_adjust()

    # “投资风格”(INV)是用第t - 1年末相对于第t - 2年末的总资产增加额，除以第t - 2年末的总资产
    inputfile = '../temp/financial data.xlsx'
    inv_temp = pd.read_excel(inputfile)

    inv_temp = inv_temp[(inv_temp['Accper'].dt.month == 12) & (
            (inv_temp['Accper'].dt.year == 2014) | (inv_temp['Accper'].dt.year == 2015) | (
            inv_temp['Accper'].dt.year == 2016))]
    inv_temp = inv_temp[['Stkcd', 'Accper', 'total_assets']]

    # 获得inv_temp中去重后的股票代码
    inv_temp2 = inv_temp
    inv_stk = inv_temp2.drop_duplicates(['Stkcd'])['Stkcd']

    inv_2014 = inv_temp[(inv_temp['Accper'].dt.year == 2014)]
    inv_2014 = inv_2014[['Stkcd', 'total_assets']]

    inv_2015 = inv_temp[(inv_temp['Accper'].dt.year == 2015)]
    inv_2015 = inv_2015[['Stkcd', 'total_assets']]

    inv_2016 = inv_temp[(inv_temp['Accper'].dt.year == 2016)]
    inv_2016 = inv_2016[['Stkcd', 'total_assets']]

    inv = pd.merge(inv_2014, inv_2015, on='Stkcd', how='outer')
    inv = pd.merge(inv, inv_2016, on='Stkcd', how='outer')

    inputpath = '../temp/size_data.xlsx'
    size_data = pd.read_excel(inputpath)

    size_data = pd.merge(size_data, inv, on=['Stkcd'], how='outer')
    size_data = size_data.dropna(axis=0, how='any')
    size_data['inv'] = ((size_data['total_assets'] - size_data['total_assets_y']) / size_data['total_assets_y'] + (
                size_data['total_assets_y'] - size_data['total_assets_x']) / size_data['total_assets_x']) / 2
    #  size_data = size_data[['Stkcd', 'Msmvosd', 'inv']]

    size_data = size_data.sort_values(by='inv', ascending=True)
    number = len(size_data)
    size_data['Inv'] = 'N'
    size_data.Inv[0:(int(number * 0.3))] = 'C'
    size_data.Inv[int(number * 0.7) + 1:] = 'A'

    outputpath = '../temp/size_data.xlsx'
    size_data.to_excel(outputpath, index=False, header=True)

    inputpath = '../temp/size_data.xlsx'
    size_data = pd.read_excel(inputpath)
    size_data = size_data.sort_values(by='Msmvosd', ascending=True)
    number = len(size_data)
    size_data['Size'] = 'S'
    size_data.Size[int(number * 0.5):] = 'B'
    size_data = size_data.sort_values(by=['Stkcd'], ascending=True)
    outputpath = '../temp/size_data.xlsx'
    size_data.to_excel(outputpath, index=False, header=True)

    # “账面市值比”(BM)是用第t - 1年末的“账面价值 / 股票i的流通市值”；
    inputfile = '../temp/financial data.xlsx'
    bm = pd.read_excel(inputfile)
    inputfile2 = '../temp/stkmnth data.xlsx'
    stk_temp = pd.read_excel(inputfile2)

    bm['Accper'] = pd.to_datetime(bm['Accper'])
    # bm['Accper'] = bm['Accper'].dt.date
    bm['year'] = bm['Accper'].dt.year
    bm['month'] = bm['Accper'].dt.month
    bm = bm[(bm['Accper'].dt.month == 12) & (
                (bm['Accper'].dt.year == 2014) | (bm['Accper'].dt.year == 2015) | (bm['Accper'].dt.year == 2016))]

    stk_temp = stk_temp[['Stkcd', 'Trdmnt', 'Msmvosd']]
    stk_temp.rename(columns={"Trdmnt": "Accper"}, inplace=True)
    stk_temp['Accper'] = pd.to_datetime(stk_temp['Accper'])
    stk_temp['year'] = stk_temp['Accper'].dt.year
    stk_temp['month'] = stk_temp['Accper'].dt.month
    stk_temp = stk_temp[(stk_temp['Accper'].dt.month == 12) & (
                (stk_temp['Accper'].dt.year == 2014) | (stk_temp['Accper'].dt.year == 2015) | (
                    stk_temp['Accper'].dt.year == 2016))]

    bm = pd.merge(bm, stk_temp, on=['Stkcd', 'year', 'month'], how='outer')
    bm['BM'] = bm.total_equity / bm.Msmvosd
    bm_temp = bm.groupby(['Stkcd'])['BM'].mean()
    bm_temp = bm_temp.dropna()

    bm = {'Stkcd': bm_temp.index, 'BM': bm_temp.values}
    bm = pd.DataFrame(bm)

    inputpath = '../temp/size_data.xlsx'
    size_data = pd.read_excel(inputpath)

    size_data = pd.merge(size_data, bm, on='Stkcd', how='outer')
    size_data = size_data.dropna(axis=0, how='any')
    size_data = size_data.sort_values(by='BM', ascending=True)
    number = len(size_data)
    size_data['bm'] = 'N'
    size_data.bm[0:int(number * 0.3)] = 'L'
    size_data.bm[int(number * 0.7) + 1:] = 'H'

    size_data = size_data.sort_values(by='Stkcd', ascending=True)

    size_data = size_data[['Stkcd', 'Msmvosd', 'Size', 'BM', 'bm', 'inv', 'Inv']]
    outputpath = '../temp/size_data.xlsx'
    size_data.to_excel(outputpath, index=False, header=True)

    # “营运利润率”(OP)是用第t - 1年末的“营业利润 / 股东权益合计”；
    inputfile = '../temp/financial data.xlsx'
    op = pd.read_excel(inputfile)

    op = op[(op['Accper'].dt.month == 12) & (
                (op['Accper'].dt.year == 2014) | (op['Accper'].dt.year == 2015) | (op['Accper'].dt.year == 2016))]
    op['OP'] = op['operating profit'] / op.total_equity
    op = op[['Stkcd', 'Accper', 'OP']]

    op = op.groupby(['Stkcd'])['OP'].mean()
    op = {'Stkcd': op.index, 'OP': op.values}
    op = pd.DataFrame(op)

    inputpath = '../temp/size_data.xlsx'
    size_data = pd.read_excel(inputpath)

    size_data = pd.merge(size_data, op, on=['Stkcd'], how='outer')

    size_data = size_data.dropna(axis=0, how='any')
    size_data = size_data.sort_values(by='OP', ascending=True)
    number = len(size_data)
    size_data['op'] = 'N'
    size_data.op[0:int(number * 0.3)] = 'W'
    size_data.op[int(number * 0.7) + 1:] = 'R'

    size_data = size_data.sort_values(by='Stkcd', ascending=True)

    #   size_data = size_data[['Stkcd', 'Msmvosd', 'Size', 'BM', 'bm', 'OP', 'op', 'inv', 'Inv']]
    outputpath = '../temp/size_data.xlsx'
    size_data.to_excel(outputpath, index=False, header=True)

    stk_adjust()

    inputfile = '../temp/stkmnth data.xlsx'
    stk_data = pd.read_excel(inputfile)
    # 获得去重后的日期
    stk_date = stk_data.drop_duplicates(['Trdmnt'])['Trdmnt']

    cal_factor(stk_date)
    cal_final(stk_date)

    print("ok")
