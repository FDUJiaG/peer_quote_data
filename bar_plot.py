import os
import time
from datetime import datetime
import string
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pyecharts.charts import Bar
from pyecharts.commons.utils import JsCode
from pyecharts import options as opts
from op_ns_data import get_ns_info_data
import warnings

warnings.filterwarnings("ignore")


def price_bar_plot(f_name):
    root_path = os.path.abspath(".")
    f_name = "伟创电气.xlsx"
    file_path = os.path.join(root_path, f_name)
    raw_df = get_ns_info_data(file_path)

    raw_df_key = raw_df.keys().tolist()
    for key_item in raw_df_key:
        if "价格" in key_item:
            sg_price = key_item
        elif "数量" in key_item:
            sg_num = key_item

    price_counts = raw_df[sg_price].value_counts()
    x_list = sorted(list(price_counts.keys()), reverse=False)
    y_list = [int(price_counts[item]) for item in x_list]

    p_v_counts = pd.pivot_table(raw_df, index=[sg_price], values=[sg_num], aggfunc=np.sum)
    print(p_v_counts)


if __name__ == '__main__':
    file_name = "伟创电气.xlsx"
    price_bar_plot(file_name)
