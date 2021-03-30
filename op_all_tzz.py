import os
import string
import time
from datetime import datetime
import numpy as np
import pandas as pd
from common_utils import col_temp, state_dict
from op_ns_data import print_info, find_file_path, get_col_list, get_ns_info_data, judge_df, get_df_group
from op_ns_data import get_note, save_df
from op_ns_data import op_ns_data
import warnings

warnings.filterwarnings("ignore")


def op_all_tzz(file_name):
    root_path = os.path.abspath(".")
    data_dir = os.path.join(root_path, "raw_data")
    # file_name = input("请输入新股中文名称：") or "恒玄科技"
    file_type = ".xlsx"
    data_name = "同行报价"
    sheet_name = "全部"
    col_list = get_col_list(root_path, data_name, sheet_name, file_type)
    if not col_list:
        return False

    file_path = find_file_path(data_dir, file_name, file_type)
    if not file_path:
        return False

    raw_df = get_ns_info_data(file_path)
    if type(raw_df) is bool:
        return raw_df

    if not judge_df(raw_df):
        return False

    raw_df_key = raw_df.keys().tolist()
    for item in raw_df_key:
        if "申购价格" in item:
            sg_jg = item
        elif "投资者" in item:
            tzz_mc = item

    raw_df_col = raw_df[tzz_mc].tolist()
    union_col = get_all_col(col_list, raw_df_col)

    df_group = get_df_group(raw_df, tzz_mc, sg_jg)
    if type(df_group) is bool:
        return df_group

    tzz_list = df_group.index.tolist()
    print(union_col)
    df_note = output_all_df(file_path, tzz_mc, raw_df, df_group, tzz_list, union_col, col_temp)

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


def output_all_df(f_path, tzz_mc, raw_df, df_group, tzz_list, col_list, col_temp):
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
                zq_name = os.path.split(f_path)[-1].split(".")[0]
                if len(zq_name) <= 4:
                    output_item[col_temp[2]] = zq_name
                else:
                    output_item[col_temp[2]] = zq_name[:4]
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
    # 删除我司列
    df_note.drop(labels="上海迎水投资管理有限公司", inplace=True, axis=1)
    return df_note


if __name__ == '__main__':
    f_name = "晓鸣股份"
    tf_gz = op_ns_data(f_name)
    tf_all = op_all_tzz(f_name)
    if tf_gz and tf_all:
        print(print_info("S"), end=" ")
        print("Success!")
    else:
        print(print_info("E"))
        print("Error!")