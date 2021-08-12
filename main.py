import os
from datetime import datetime
from ReportGen.charts import Box, Line, Table, Bar
from ReportGen.ppt import SlideSelect
from ReportGen.prep import *


if __name__ == "__main__":
    img_folder = os.path.join(os.getcwd(), 'img')

    #### Import Data ####
    df = generate_features(consolidate_data('tx_data_'))
    mapper = consolidate_data('map')
    df = df.merge(mapper, how='left', on='PRODUCT_NAME')

    #### Process Data ####
    gp_prod_sum, gp_prod_avg, gp_class_sum, gp_class_avg, gp_date, gp_mth, gp_day = group_process(
        df)
    ma = ma_process(gp_date)
    top_10_prods, top_10_classes, top_10_classes2, top_10_prod_cost = top_process(
        df, gp_prod_avg, gp_class_avg)

    #### Set I/O image paths ####
    img_folder = os.path.join(os.getcwd(), 'img')
    boxplt_path = os.path.join(img_folder, 'boxplot.png')
    timesertop_path = os.path.join(img_folder, 'timeseries_top.png')
    timeserma_path = os.path.join(img_folder, 'timeseries_ma.png')
    combo_path = os.path.join(img_folder, 'combo_top_and_count.png')
    prods_path = os.path.join(img_folder, 'top_10_prods.png')
    classes_path = os.path.join(img_folder, 'top_10_classes.png')

    #### Generate charts ####
    # 1) boxplot - top 10 expensive products
    top_10_prod_cost = gen_top_cost_data(df)
    topCost = Box(8, 5, top_10_prod_cost, boxplt_path)
    fig_tc = topCost.plot('TOT_COST', 'PRODUCT_CLASS',
                          x_lab='Cost ($)', y_lab='Product Class')
    topCost.save()

    # 2) lineplot - top 3 products and qty sold over 1 yr
    top_10_large_class_idx, top_10_large_class = gen_top_time_series(df)
    topClassTS = Line(12, 5.5, top_10_large_class, timesertop_path,
                      idx_data=top_10_large_class_idx, multi_line=True)
    fig_tcts = topClassTS.plot(
        'TX_DATE', 'QTY', x_lab='Date', y_lab='Quantity Sold', selector_col='PRODUCT_CLASS')
    topClassTS.save()

    # 3) Lineplot - revenue and MA over 1-yr
    revMA = Line(12, 5, ma, timeserma_path, multi_line=True)
    fig_rma = revMA.plot(ma.index, '', 2, x_lab='Date', y_lab='Revenue',
                         y_list=['TOT_REV', 'TOT_REV_MA_7', 'TOT_REV_MA_30'],
                         leg_lab_list=['Revenue', 'Revenue (7-day Moving Average)', 'Revenue (30-day Moving Average)'])
    revMA.save()

    # 4) Comboplot - top 10 products net sales and count
    topNetSalesCount = Bar(10, 5.5, top_10_classes2, combo_path, combo=True)
    topProdCount = Line(10, 5.5, top_10_classes2, combo=True)
    fig_tnsc, combo_ax = topNetSalesCount.plot(
        'PRODUCT_CLASS', 'NET_SALES', orient='v', color='lightgrey', x_lab='Products', y_lab='Net Sales ($)', b_rt_spine=True)
    topProdCount.plot('PRODUCT_CLASS', 'PROD_CNT', color='red', y_lab='Product Count',
                      linewidth=0.9, b_rt_spine=True, combo_ax=combo_ax)
    # ax.set(xlabel=x_lab, ylabel=y_lab)
    combo_ax.set(ylabel='Product Count')
    combo_ax.spines['top'].set_visible(False)
    topNetSalesCount.save()

    # 5) Table - top product name, class by net sales
    top_10_prods_fig_data = gen_top_name_class_tables(top_10_prods, 'NET_SALES', {
        'PRODUCT_NAME': 'Product Name', 'NET_SALES': 'Net Sales ($)'})
    top_10_classes_fig_data = gen_top_name_class_tables(top_10_classes, 'NET_SALES', {
        'PRODUCT_CLASS': 'Product Class', 'NET_SALES': 'Net Sales ($)'})
    topProdSalesTab = Table(0.4, 4, top_10_prods_fig_data, prods_path)
    topClassesSalesTab = Table(0.4, 2, top_10_classes_fig_data, classes_path)
    topProdSalesTab.plot()
    topClassesSalesTab.plot()

    #### Copy to pptx ####
    python_logo_path = os.path.join(img_folder, 'python.png')
    bars_logo_path = os.path.join(img_folder, 'bars.png')
    template_path = os.path.join(os.getcwd(), 'design_template.pptx')
    out_path = os.path.join(os.getcwd(), 'report_{}.pptx'.format(
        datetime.now().strftime('%Y-%m-%d')))
    report = SlideSelect(out_path, template_path)

    # slide 1: title page
    title = report.create_slide('title', 'title')
    title.create(label='2021 Profitability Report',
                 sub_label='Generated on - {}'.format(
                     datetime.now().date().strftime('%b %d, %Y')),
                 logo_path_1=python_logo_path,
                 logo_path_2=bars_logo_path)

    # slide 2: summary
    summary = report.create_slide('summary', 'text')
    summary.create(label='Report Summary',
                   text='''Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque purus lorem, eleifend non felis pulvinar, rutrum aliquam lectus. Phasellus risus risus, eleifend vehicula placerat id, luctus eu sem. Praesent in nisl eleifend, volutpat lacus eget, finibus velit. ''Aliquam eget odio varius, bibendum felis vitae, euismod libero. In tempus mi enim, sit amet tempor odio ornare a. Nunc non elementum ex. Quisque ullamcorper, lorem in fermentum posuere, justo lectus feugiat sapien, at sagittis leo metus ac quam. Duis aliquet nisl lobortis massa maximus gravida. Proin pharetra egestas bibendum. Integer sagittis venenatis mi, ut tempor ante. Aliquam erat volutpat. Morbi hendrerit porta iaculis. Nullam vel odio turpis. Aenean dictum, arcu in efficitur hendrerit, purus risus accumsan lorem, a rutrum mi enim ut enim. Aliquam volutpat neque neque, et aliquet nisl blandit vel. Pellentesque a tempor turpis. Nunc accumsan magna velit, ac commodo metus consequat quis. Aliquam erat volutpat. Sed ac gravida magna. Cras velit lacus, sollicitudin ut velit iaculis, elementum gravida urna. Nam turpis ligula, sagittis eu sapien rutrum, laoreet hendrerit dolor. Duis interdum et augue ut faucibus. Sed mi ex, luctus nec iaculis eget, semper vel lacus. Vestibulum vitae quam non massa semper gravida et id nulla.''')

    # slide 3: image - boxplot
    box = report.create_slide('charts', 'blank')
    box.create(boxplt_path, label='Cost Distribution Among Product Classes',
               boxplot=[1.25, 0.25])

    # # slide 4: image - time series 1
    timeseries_top = report.create_slide('charts', 'blank')
    timeseries_top.create(label='Average Top 3 Sold Products (2021)',
                          path=timesertop_path)

    # # slide 5: image - time series 2
    timeseries_ma = report.create_slide('charts', 'blank')
    timeseries_ma.create(label='2021 Revenue and Trends',
                         path=timeserma_path)

    # # slide 6: image - combo
    timeseries_ma = report.create_slide('charts', 'blank')
    timeseries_ma.create(label='Average Top 10 Net Sales & Product Count',
                         path=combo_path)

    # # slide 7: tables
    tables = report.create_slide('datatable', 'two_columns')
    tables.create(label='Top Product and Class Revenue',
                  table_props={
                      'subtitle_prop': ['Top 10 Product Revenue', 1],
                      'table_prop': [prods_path, [-0.25, 0.75]]
                  })
    tables.create(label='Top Product and Class Revenue',
                  table_props={
                      'subtitle_prop': ['Top 10 Product Class Revenue', 3],
                      'table_prop': [classes_path, [5.35, 0.75]]
                  })

    # save pptx
    report.save_file()
