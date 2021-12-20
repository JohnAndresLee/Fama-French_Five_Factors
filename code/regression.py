import numpy as np
import pandas as pd
import sklearn
import csv
import statsmodels.api as sm
from numpy.linalg import inv
from scipy.stats import f

inputpath = '../temp/five_factor.xlsx'
factor_data = pd.read_excel(inputpath)

factor_data['Time'] = pd.to_datetime(factor_data['Time'])
factor_data = factor_data[(factor_data['Time'].dt.year == 2014) | (factor_data['Time'].dt.year == 2015) |(factor_data['Time'].dt.year == 2016)]
#
# # 拟合CMA
# FX = factor_data[['rm-rf', 'SMB', 'HML', 'RMW']]
# FY = factor_data[['CMA']]
# FX = sm.add_constant(FX)
#
# model = sm.OLS(FY, FX).fit()
# #predictions = model.predict(FX)
#
# print_model = model.summary()
# print("CMA")
# print(model.params)
# print("t values:")
# print(model.tvalues)
# print(model.rsquared)
#
# # 拟合SMB
# FX = factor_data[['rm-rf', 'CMA', 'HML', 'RMW']]
# FY = factor_data[['SMB']]
# FX = sm.add_constant(FX)
#
# model = sm.OLS(FY, FX).fit()
# #predictions = model.predict(FX)
#
# print_model = model.summary()
# print("SMB")
# print(model.params)
# print("t values:")
# print(model.tvalues)
# print(model.rsquared)
#
# # 拟合HML
# FX = factor_data[['rm-rf', 'SMB', 'CMA', 'RMW']]
# FY = factor_data[['HML']]
# FX = sm.add_constant(FX)
#
# model = sm.OLS(FY, FX).fit()
# #predictions = model.predict(FX)
#
# print_model = model.summary()
# print("HML")
# print(model.params)
# print("t values:")
# print(model.tvalues)
# print(model.rsquared)
#
# # 拟合RMW
# FX = factor_data[['rm-rf', 'SMB', 'HML', 'CMA']]
# FY = factor_data[['RMW']]
# FX = sm.add_constant(FX)
#
# model = sm.OLS(FY, FX).fit()
# #predictions = model.predict(FX)
#
# print_model = model.summary()
# print("RMW")
# print(model.params)
# print("t values:")
# print(model.tvalues)
# print(model.rsquared)
#
# # 拟合rm-rf
# FX = factor_data[['CMA', 'SMB', 'HML', 'RMW']]
# FY = factor_data[['rm-rf']]
# FX = sm.add_constant(FX)
#
# model = sm.OLS(FY, FX).fit()
# #predictions = model.predict(FX)
#
# print_model = model.summary()
# print("rm-rf")
# print(model.params)
# print("t values:")
# print(model.tvalues)
# print(model.rsquared)

inputpath = '../temp/25combine/2014-2016Size-BM.xlsx'
Size_BP_df = pd.read_excel(inputpath)

rf_data = factor_data['rf']
Size_BP_df['rf'] = rf_data

inputpath = '../temp/25combine/2014-2016Size-INV.xlsx'
Size_INV_df = pd.read_excel(inputpath)
Size_INV_df['rf'] = rf_data

inputpath = '../temp/25combine/2014-2016Size-OP.xlsx'
Size_OP_df = pd.read_excel(inputpath)
Size_OP_df['rf'] = rf_data

# # 五因子线性回归
# FX = factor_data[['rm-rf', 'SMB', 'HML', 'RMW', 'CMA']]
# temp = Size_BP_df.iloc[:, [1, 26]]
# FY = temp.iloc[:, 0]-0.01*temp.iloc[:, 1]
# FY = FY-FY.mean()
# FX = sm.add_constant(FX)
# model = sm.OLS(FY, FX).fit()
# print_model = model.summary()
# print(print_model)


# %%
def ols(five_factor, Y_df):
    martix = []
    for i in range(1, 26):
        FX = five_factor[['rm-rf', 'SMB', 'HML', 'RMW', 'CMA']]
        FY = Y_df.iloc[:, [i, 26]]
        FY = FY.iloc[:, 0]-0.01*FY.iloc[:, 1]
        FX = sm.add_constant(FX)
        model = sm.OLS(FY, FX).fit()
        row = list(model.params.values) + list(model.tvalues.values)
        martix.append(row)
    return martix


result_Size_BP_df = pd.DataFrame(ols(factor_data, Size_BP_df))
result_Size_BP_df = pd.DataFrame(ols(factor_data, Size_BP_df),
                                 columns=['const', 'beta_rm-rf', 'beta_SMB', 'beta_HML', 'beta_RMW', 'beta_CMA',
                                          't_const', 't_beta_rm-rf', 't_beta_SMB', 't_beta_HML', 't_beta_RMW',
                                          't_beta_CMA'])
result_Size_BP_df.to_excel('../temp/regression/14-16Size_BP_result.xlsx', index=False, header=True)
#
result_Size_INV_df = pd.DataFrame(ols(factor_data, Size_INV_df),
                                  columns=['const', 'beta_market', 'beta_SMB', 'beta_HML', 'beta_RMW', 'beta_CMA',
                                           't_const', 't_beta_market', 't_beta_SMB', 't_beta_HML', 't_beta_RMW',
                                           't_beta_CMA'])
result_Size_INV_df.to_excel('../temp/regression/14-16Size_INV_result.xlsx', index=False, header=True)
#
result_Size_OP_df = pd.DataFrame(ols(factor_data, Size_OP_df),
                                 columns=['const', 'beta_market', 'beta_SMB', 'beta_HML', 'beta_RMW', 'beta_CMA',
                                          't_const', 't_beta_market', 't_beta_SMB', 't_beta_HML', 't_beta_RMW',
                                          't_beta_CMA'])
result_Size_OP_df.to_excel('../temp/regression/14-16Size_OP_result.xlsx', index=False, header=True)

factor_data = factor_data[['rm-rf', 'SMB', 'HML', 'RMW', 'CMA']]
# %%
def GRS(alpha, resids, mu):
    # GRS test statistic
    # N assets, L factors, and T time points
    # alpha is a Nx1 vector of intercepts of the time-series regressions,
    # resids is a TxN matrix of residuals,
    # mu is a TxL matrix of factor returns
    T, N = resids.shape
    L = mu.shape[1]
    mu_mean = np.nanmean(mu, axis=0)
    cov_resids = np.matmul(resids.T, resids) / (T - L - 1)
    cov_fac = np.matmul(np.array(mu - np.nanmean(mu, axis=0)).T, np.array(mu - np.nanmean(mu, axis=0))) / T - 1
    GRS = float((T / N) * ((T - N - L) / (T - L - 1)) * ((np.matmul(np.matmul(alpha.T, inv(cov_resids)), alpha)) / (
                1 + (np.matmul(np.matmul(mu_mean.T, inv(cov_fac)), mu_mean)))))
    pVal = 1 - f.cdf(GRS, N, T - N - L)
    return GRS, pVal


#li = list(range(1, 26)[::2])

def ols_const(five_factor, Y_df):
    const = []
    for i in range(1, 26):
        FX = five_factor
        FY = Y_df.iloc[:, [i, 26]]
        FY = FY.iloc[:, 0]-0.01*FY.iloc[:, 1]
        FX = sm.add_constant(FX)
        model = sm.OLS(FY, FX).fit()
        const.append(model.params[0])
    return const

# Aalpha = np.array(list(map(abs, ols_const(factor_data, Size_BP_df)))).mean()
# # alpha 的平均值
# print(Aalpha)
# debug = Size_BP_df.iloc[:, 1:26]
# rm = np.array(Size_BP_df.iloc[:, 1:26].mean()).mean()
# mean_list = []
# for i in range(1, 26):
#     col = np.array(Size_BP_df.iloc[:, [i]])
#     col = np.array(list(map(abs, list(col - rm))))
#     mean_list.append(col.mean())
# print(Aalpha / np.array(mean_list).mean())


def print_grs(five_factor, Size_BP_df):
    five_factor = five_factor[['rm-rf', 'SMB', 'HML', 'RMW', 'CMA']]
    li = list(range(1, 26))
    rm = np.array(Size_BP_df.iloc[:, 1:26].mean()).mean()
    mean_list = []
    for i in range(1, 26):
        col = np.array(Size_BP_df.iloc[:, [i]])
        col = np.array(list(map(abs, list(col - rm))))
        mean_list.append(col.mean())
    Ar = np.array(mean_list).mean()

    alpha = pd.DataFrame(ols_const(five_factor, Size_BP_df))
    resids = Size_BP_df.iloc[:, li]
    mu = five_factor
    Aalpha = np.array(list(map(abs, ols_const(five_factor, Size_BP_df)))).mean()
    print('Aalpha', Aalpha)
    print('Aa/Ar', Aalpha / Ar)
    print('五因子', GRS(alpha, resids, mu))

    alpha = pd.DataFrame(ols_const(five_factor[['rm-rf', 'SMB', 'HML']], Size_BP_df))
    resids = Size_BP_df.iloc[:, li]
    mu = five_factor[['rm-rf', 'SMB', 'HML']]
    Aalpha = np.array(list(map(abs, ols_const(mu, Size_BP_df)))).mean()
    print('Aalpha', Aalpha)
    print('Aa/Ar', Aalpha / Ar)
    debug1 = list(map(abs, alpha))
    debug2 = GRS(alpha, resids, mu)
    print('HML', GRS(alpha, resids, mu), list(map(abs, alpha)))

    alpha = pd.DataFrame(ols_const(five_factor[['rm-rf', 'SMB', 'RMW', 'CMA']], Size_BP_df))
    resids = Size_BP_df.iloc[:, li]
    mu = five_factor[['rm-rf', 'SMB', 'RMW', 'CMA']]
    Aalpha = np.array(list(map(abs, ols_const(mu, Size_BP_df)))).mean()
    print('Aalpha', Aalpha)
    print('Aa/Ar', Aalpha / Ar)
    print('rm-rf SMB CMA RMW', GRS(alpha, resids, mu), list(map(abs, alpha)))

    alpha = pd.DataFrame(ols_const(five_factor[['rm-rf', 'SMB', 'HML', 'RMW']], Size_BP_df))
    resids = Size_BP_df.iloc[:, li]
    mu = five_factor[['rm-rf', 'SMB', 'HML', 'RMW']]
    Aalpha = np.array(list(map(abs, ols_const(mu, Size_BP_df)))).mean()
    print('Aalpha', Aalpha)
    print('Aa/Ar', Aalpha / Ar)
    print('rm-rf SMB HML RMW', GRS(alpha, resids, mu), list(map(abs, alpha)))

    alpha = pd.DataFrame(ols_const(five_factor[['rm-rf', 'SMB', 'HML', 'CMA']], Size_BP_df))
    resids = Size_BP_df.iloc[:, li]
    mu = five_factor[['rm-rf', 'SMB', 'HML', 'CMA']]
    Aalpha = np.array(list(map(abs, ols_const(mu, Size_BP_df)))).mean()
    print('Aalpha', Aalpha)
    print('Aa/Ar', Aalpha / Ar)
    print('rm-rf SMB HML CMA', GRS(alpha, resids, mu), list(map(abs, alpha)))

    alpha = pd.DataFrame(ols_const(five_factor[['rm-rf', 'RMW', 'HML', 'CMA']], Size_BP_df))
    resids = Size_BP_df.iloc[:, li]
    mu = five_factor[['rm-rf', 'RMW', 'HML', 'CMA']]
    Aalpha = np.array(list(map(abs, ols_const(mu, Size_BP_df)))).mean()
    print('Aalpha', Aalpha)
    print('Aa/Ar', Aalpha / Ar)
    print('rm-rf RMW HML CMA', GRS(alpha, resids, mu), list(map(abs, alpha)))

print_grs(factor_data, Size_BP_df)
print_grs(factor_data, Size_INV_df)
print_grs(factor_data, Size_OP_df)


print("ok")