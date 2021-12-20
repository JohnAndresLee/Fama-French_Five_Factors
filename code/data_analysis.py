import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np


if __name__ == '__main__':
    picoutpath = '../temp/pic'
    inputpath = '../temp/five_factor.xlsx'
    factor_data = pd.read_excel(inputpath)
    factor_data = factor_data[['Time', 'rm-rf', 'SMB', 'HML', 'RMW', 'CMA']]

    # 计算14-16年数据
    factor_data['Time'] = pd.to_datetime(factor_data['Time'])
    factor_data_1416 = factor_data[(factor_data['Time'].dt.year == 2014) | (factor_data['Time'].dt.year == 2015) | (
                factor_data['Time'].dt.year == 2016)]

    data14_16 = factor_data_1416.describe()
    data14_16 = pd.DataFrame(data14_16)
    outputpath = '../temp/14-16因子的概括性描述分析.xlsx'
    data14_16.to_excel(outputpath, index=True, header=True)
    data14_16_corr = factor_data_1416.corr()
    sns.heatmap(data14_16_corr, annot=True, square=True, cmap="Blues")
    plt.savefig("14-16_factor_corr.png")
    plt.show()


    sns.pairplot(factor_data_1416, x_vars=['rm-rf', 'SMB', 'HML', 'RMW', 'CMA'], y_vars=['rm-rf', 'SMB', 'HML', 'RMW', 'CMA'], kind='reg', diag_kind="kde")
    plt.savefig("14-16_factor_reg.png")
    plt.show()

    # # 计算15-17年数据
    # factor_data_1517 = factor_data[(factor_data['Time'].dt.year == 2017) | (factor_data['Time'].dt.year == 2015) | (
    #             factor_data['Time'].dt.year == 2016)]
    #
    # data15_17 = factor_data_1517.describe()
    # data15_17_corr = factor_data_1517.corr()
    # sns.heatmap(data15_17_corr, annot=True, square=True, cmap="Blues")
    # plt.savefig("15-17_factor_corr.png")
    # plt.show()
    #
    # sns.pairplot(factor_data_1517, x_vars=['rm-rf', 'SMB', 'HML', 'RMW', 'CMA'],
    #              y_vars=['rm-rf', 'SMB', 'HML', 'RMW', 'CMA'], kind='reg', diag_kind="kde")
    # plt.savefig("15-17_factor_reg.png")
    # plt.show()
    #
    # # 计算16-18年数据
    # factor_data_1618 = factor_data[(factor_data['Time'].dt.year == 2017) | (factor_data['Time'].dt.year == 2018) | (
    #             factor_data['Time'].dt.year == 2016)]
    #
    # data16_18 = factor_data_1618.describe()
    # data16_18_corr = factor_data_1618.corr()
    # sns.heatmap(data16_18_corr, annot=True, square=True, cmap="Blues")
    # plt.savefig("16-18_factor_corr.png")
    # plt.show()
    #
    # sns.pairplot(factor_data_1618, x_vars=['rm-rf', 'SMB', 'HML', 'RMW', 'CMA'],
    #              y_vars=['rm-rf', 'SMB', 'HML', 'RMW', 'CMA'], kind='reg', diag_kind="kde")
    # plt.savefig("16-18_factor_reg.png")
    # plt.show()

    print("ok")
