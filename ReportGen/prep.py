import pandas as pd
import numpy as np
import os


def consolidate_data(file_root):
    data_folder = os.path.join(os.getcwd(), 'data')
    files = pd.Series(os.listdir(data_folder), name='Document')
    file_filter = files.str.contains(file_root)
    data = [pd.read_excel(os.path.join(data_folder, f), engine='openpyxl')
            for f in files[file_filter]]
    if len(data) > 1:
        return pd.concat(data).reset_index(drop=True)
    elif len(data) == 1:
        return data[0]
    else:
        return 'No data with the suggested file root name found.'


def generate_features(df):
    tmp = df.copy()
    tmp['TX_DATE'] = pd.to_datetime(tmp['TX_DATE'])
    tmp['TX_YER'] = tmp['TX_DATE'].dt.year
    tmp['TX_MTH'] = tmp['TX_DATE'].dt.month
    tmp['TX_DAY'] = tmp['TX_DATE'].dt.day
    tmp['TOT_REV'] = tmp['PRICE'] * tmp['QTY']
    tmp['TOT_COST'] = tmp['PURCH_COST'] * tmp['QTY']
    tmp['NET_SALES'] = tmp['TOT_REV'] - tmp['TOT_COST']
    tmp.sort_values(['TX_DATE', 'CUSTOMER_ID'], inplace=True)
    tmp.reset_index(drop=True, inplace=True)
    return tmp


def get_grouped(df, cat, cols, aggfunc):
    tmp = df.copy()
    return tmp.groupby(cat)[cols].agg(aggfunc)


def calculate_ma(df, col, window):
    tmp = df.copy().reset_index()
    idx_col = tmp.columns[0]
    tmp[col + '_MA_' + str(window)] = tmp[col].rolling(window).mean()
    return tmp.set_index(idx_col)


def group_process(df):
    gp_prod_sum = get_grouped(
        df, 'PRODUCT_NAME', ['TOT_REV', 'TOT_COST', 'NET_SALES'], 'sum')
    gp_prod_avg = get_grouped(
        df, 'PRODUCT_NAME', ['TOT_REV', 'TOT_COST', 'NET_SALES'], 'mean')
    gp_class_sum = get_grouped(df, 'PRODUCT_CLASS', [
                               'TOT_REV', 'TOT_COST', 'NET_SALES'], 'sum')
    gp_class_avg = get_grouped(df, 'PRODUCT_CLASS', [
                               'TOT_REV', 'TOT_COST', 'NET_SALES'], 'mean')
    gp_date = get_grouped(
        df, 'TX_DATE', ['TOT_REV', 'TOT_COST', 'NET_SALES', 'QTY'], 'sum')
    gp_mth = get_grouped(
        df, 'TX_MTH', ['TOT_REV', 'TOT_COST', 'NET_SALES'], 'sum')
    gp_day = get_grouped(
        df, 'TX_DAY', ['TOT_REV', 'TOT_COST', 'NET_SALES'], 'sum')
    return gp_prod_sum, gp_prod_avg, gp_class_sum, gp_class_avg, gp_date, gp_mth, gp_day


def ma_process(df):
    tmp = calculate_ma(df, 'TOT_REV', 7)
    tmp = calculate_ma(tmp, 'TOT_REV', 30)
    tmp = calculate_ma(tmp, 'QTY', 7)
    tmp = calculate_ma(tmp, 'QTY', 30)
    return tmp


def top_process(main, prod_avg, class_avg):
    top_10_prods = prod_avg['NET_SALES'].sort_values(
        ascending=False).head(10).reset_index()
    top_10_classes = class_avg['NET_SALES'].sort_values(
        ascending=False).head(10).reset_index()

    gp_prod_cnt = get_grouped(main, 'PRODUCT_NAME', ['NET_SALES'], 'count').rename(
        columns={'NET_SALES': 'PROD_CNT'})
    gp_class_cnt = get_grouped(main, 'PRODUCT_CLASS', ['NET_SALES'], 'count').rename(
        columns={'NET_SALES': 'PROD_CNT'})
    gp_class_cnt['PROD_CNT_PCT'] = gp_class_cnt['PROD_CNT'] / \
        gp_class_cnt['PROD_CNT'].sum() * 100

    # top 10 products
    top_10_prod_cnt = gp_prod_cnt[gp_prod_cnt.index.isin(
        top_10_prods['PRODUCT_NAME'])]
    top_10_prods2 = top_10_prods.merge(
        top_10_prod_cnt, how='left', on='PRODUCT_NAME')

    # top 10 classes
    top_10_class_cnt = gp_class_cnt[gp_class_cnt.index.isin(
        top_10_classes['PRODUCT_CLASS'])]
    top_10_classes2 = top_10_classes.merge(
        top_10_class_cnt, how='left', on='PRODUCT_CLASS')

    # top 10 product costs
    top_10_prod_cost_idx = main[['PRODUCT_CLASS', 'TOT_COST']].groupby(
        'PRODUCT_CLASS').sum().sort_values('TOT_COST', ascending=False).head(10).index
    top_10_prod_cost = main[main['PRODUCT_CLASS'].isin(top_10_prod_cost_idx)][[
        'PRODUCT_CLASS', 'TOT_COST']]

    # top 10 customer revenue
    top_10_cust_rev = main.groupby('CUSTOMER_ID').sum(
    )[['QTY', 'NET_SALES']].sort_values('NET_SALES', ascending=False).head(10)

    # top 10 customer purchases
    top_10_cust_purch = main.groupby('CUSTOMER_ID').sum(
    )[['QTY', 'NET_SALES']].sort_values('QTY', ascending=False).head(10)
    return top_10_prods, top_10_classes, top_10_classes2, top_10_prod_cost


def gen_top_cost_data(df, top=10):
    idx = df[['PRODUCT_CLASS', 'TOT_COST']].groupby('PRODUCT_CLASS').sum(
    ).sort_values('TOT_COST', ascending=False).head(top).index
    prod_cost = df[df['PRODUCT_CLASS'].isin(
        idx)][['PRODUCT_CLASS', 'TOT_COST']]
    return prod_cost


def gen_top_time_series(df, top=3):
    idx = df.groupby('PRODUCT_CLASS').count().sort_values(
        'PRODUCT_CATEGORY', ascending=False).head(top).index
    large_class = df[df['PRODUCT_CLASS'].isin(idx)].groupby(
        ['TX_DATE', 'PRODUCT_CLASS']).sum()['QTY'].reset_index()
    # top_10_large_class = top_10_large_class.assign(ALPHA=top_10_large_class['PRODUCT_CLASS'].map(class_alpha_map))
    return idx, large_class


def gen_top_name_class_tables(df, col_to_int, rename_dict):
    tmp = df.copy()
    tmp[col_to_int] = tmp[col_to_int].astype(int)
    tmp.rename(columns=rename_dict, inplace=True)
    return tmp
