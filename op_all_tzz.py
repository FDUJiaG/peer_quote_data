import os
import string
import time
from datetime import datetime
import numpy as np
import pandas as pd
from common_utils import col_temp, state_dict
from op_ns_data import print_info, print_new_info, find_file_path, get_col_list, get_ns_info_data, judge_df
from op_ns_data import get_note, save_df, get_name_and_code
from op_ns_data import op_ns_data
from WindPy import w
import warnings

warnings.filterwarnings("ignore")


def op_all_tzz(file_name):
    root_path = os.path.abspath(".")
    data_dir = os.path.join(root_path, "raw_data")
    file_type = ".xlsx"
    data_name = "同行报价"
    sheet_name = "全部"
    col_list = get_col_list(root_path, data_name, sheet_name, file_type)
    if not col_list:
        return False

    file_path = find_file_path(data_dir, file_name, file_type)
    if not file_path:
        return False

    long_name = os.path.split(file_path)[-1].split(".")[0]
    ipo_name, ipo_code = get_name_and_code(long_name)
    # if ipo_code != "":
    #     data = w.wsd(
    #         ipo_code, "sec_name,ipo_inq_enddate",
    #         "ED-1TD", datetime.now().strftime("%Y-%m-%d")
    #     )
    #     if ipo_name != data.Data[0][0]:
    #         print(print_new_info("E", "R"), end=" ")
    #         print("Name: {} and Code: {} not match!".format(ipo_name, ipo_code))
    #         return False

    raw_df = get_ns_info_data(file_path)
    if type(raw_df) is bool:
        return raw_df

    # 替换
    raw_df.replace("招商基金管理有限公司", "招商基金", inplace=True)

    if not judge_df(raw_df):
        return False

    raw_df_key = raw_df.keys().tolist()
    for item in raw_df_key:
        if "价格" in item:
            sg_jg = item
        elif "投资者" in item or "交易员" in item:
            tzz_mc = item

    for idx, item in zip(range(len(raw_df[tzz_mc])), raw_df[tzz_mc]):
        raw_df[tzz_mc][idx] = item.replace("(", "（").replace(")", "）")

    raw_df_col = raw_df[tzz_mc].tolist()
    union_col = get_all_col(col_list, raw_df_col)

    df_group = get_df_group_one(raw_df, tzz_mc, sg_jg)
    if type(df_group) is bool:
        return df_group

    tzz_list = df_group.index.tolist()
    print(union_col)
    df_note = output_all_df(file_path, tzz_mc, raw_df, df_group, tzz_list, union_col, col_temp, ipo_name, ipo_code)

    try:
        save_df(df_note, root_path, file_path, sheet_name, op_file_type=".xlsx")
        return True
    except:
        return False


def get_all_col(col_list, raw_df_col):
    diff_col = list(set(raw_df_col).difference(set(col_list)))
    all_list = col_list.copy()
    for item in diff_col:
        all_list.append(item)
    return all_list


def output_all_df(f_path, tzz_mc, raw_df, df_group, tzz_list, col_list, col_temp, ipo_n, ipo_c):
    df_output = pd.DataFrame(columns=col_temp)
    desc_list = list()

    for item in col_list:
        output_item = dict()
        output_item[col_temp[0]] = item
        output_item[col_temp[1]] = ""
        output_item[col_temp[3]] = ""
        item = item.rstrip(string.digits)
        item_set = set("".join(item))
        idx, p = 0, 0
        for tzz_item in tzz_list:
            tzz_item_set = set("".join(tzz_item))
            if item_set.issubset(tzz_item_set):
                idx += 1
                p = df_group.loc[tzz_item][0]
                output_item[col_temp[1]] = tzz_item
                output_item[col_temp[3]], desc_item = get_note(raw_df, tzz_mc, item, tzz_item, state_dict)

        if idx > 1:
            idx = 0
            for tzz_item in tzz_list:
                if item in tzz_item:
                    idx += 1
                    p = df_group.loc[tzz_item][0]
                    output_item[col_temp[1]] = tzz_item
                    output_item[col_temp[3]], desc_item = get_note(raw_df, tzz_mc, item, tzz_item, state_dict)

        if p == 0:
            if item == "证券名称":
                # zq_name = os.path.split(f_path)[-1].split(".")[0]
                # if len(zq_name) <= 4:
                #     output_item[col_temp[2]] = zq_name
                # else:
                #     output_item[col_temp[2]] = zq_name[:4]
                output_item[col_temp[2]] = ipo_n
            elif item == "证券代码":
                output_item[col_temp[2]] = ipo_c
            elif item == "询价日期":
                if ipo_c != "":
                    data = w.wsd(
                        ipo_c, "sec_name,ipo_inq_enddate",
                        "ED-1TD", datetime.now().strftime("%Y-%m-%d")
                    )
                    output_item[col_temp[2]] = data.Data[-1][0].strftime("%Y-%m-%d")
            elif item == "我司报价":
                output_item[col_temp[1]] = "上海迎水投资管理有限公司"
                for tzz_item in tzz_list:
                    if "迎水" in tzz_item:
                        output_item[col_temp[2]] = df_group.loc[tzz_item][0]
            else:
                output_item[col_temp[2]] = ""
            df_output = df_output.append(output_item, ignore_index=True)
        else:
            if desc_item != "":
                desc_list.append(desc_item)
            output_item[col_temp[2]] = p
            df_output = df_output.append(output_item, ignore_index=True)

    desc_note = "；".join(desc_list)
    print(print_info(), end=" ")
    print("Note: {}".format(desc_note))
    df_output.set_index(col_temp[0], inplace=True)
    # 删除全称列
    df_output.drop(columns=col_temp[1], inplace=True)
    df_note = pd.DataFrame(df_output.values.T, index=df_output.columns, columns=df_output.index)
    # 查看我司的备注
    if "上海迎水投资管理有限公司" in df_note.columns:
        df_note["我司报价"]["备注"] = df_note["上海迎水投资管理有限公司"]["备注"]
        print(print_info(), end=" ")
        print("我司备注：{}".format(df_note["我司报价"]["备注"]))
        # 删除我司全称列
        df_note.drop(labels="上海迎水投资管理有限公司", inplace=True, axis=1)
    else:
        df_note["我司报价"]["备注"] = ""
    return df_note


def get_df_group_one(df, index_n, price):
    # 仅显示报价的众数
    try:
        index_name = index_n
        df_group = df.groupby(index_name)[price, "备注"].agg(lambda x: x.value_counts().index[0]).reset_index()
        df_group.set_index(index_name, inplace=True)
        print(print_info(), end=" ")
        print("The Group by DataFrame:\n {}".format(df_group))
        df_group.to_excel("test.xlsx")
        return df_group
    except:
        print(print_new_info("E", "R"), end=" ")
        print("There is something wrong in DataFrame:\n {}".format(df.head()))
        return False


if __name__ == '__main__':
    f_name = "卡莱特"
    w.start()
    w.isconnected()
    tf_gz = op_ns_data(f_name)
    tf_all = op_all_tzz(f_name)
    if tf_gz and tf_all:
        print(print_info("S"), end=" ")
        print("Success!")
    else:
        print(print_info("E"))
        print("Error!")
    w.close()
