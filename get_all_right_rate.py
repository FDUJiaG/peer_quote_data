import os
import time
from datetime import datetime
import numpy as np
import pandas as pd
from WindPy import w
import warnings

warnings.filterwarnings("ignore")


state_dict = {
    "低": "低价剔除",
    "高": "高价剔除",
    "无": "无效报价",
    "有": "有效报价",
    "未": "无效报价"
}

kcb_col = [
    "公司简称",
    "代码",
    "询价日",
    "网下投资者名称",
    "申购价格（元）",
    "报价结果"
]


def get_all_right_rate(r_path, d_dir, o_dir, min_num):
    data_dir = os.path.join(r_path, d_dir)
    output_dir = os.path.join(r_path, o_dir)
    save_name = "科创板网下投资者报价正确率（累计配售对象{}次以上）".format(min_num)
    # file_name = ["科创板", "创业板"]
    file_name = ["科创板"]
    file_type = ".xlsx"
    sheet_name = "基础数据"
    # # 设置本公司加成情况
    my_comp_name = "上海迎水投资管理有限公司"
    # print(print_info(), end=" ")
    # ys_plus = input("请输入迎水的加成次数（默认4次）：") or 4
    #
    # try:
    #     ys_plus = int(ys_plus)
    # except:
    #     while type(ys_plus) is not int:
    #         print(print_info("W"), end=" ")
    #         ys_plus = input("你输入一个整数好吧：") or 4
    #     ys_plus = int(ys_plus)

    # 获取一张精简表
    df_col = [
        "股票名称",
        "标准代码",
        "询价日",
        "投资者名称",
        "申购价格",
        "备注"
    ]
    df_zero = pd.DataFrame(columns=df_col)

    # 设置输出格式
    rate_col = [
        "序号",
        "投资者名称",
        "参与报价新股数量",
        "报价正确次数（部分正确为0）",
        "报价正确次数（部分正确为1）",
        "报价正确率（部分正确为0）",
        "报价正确率（部分正确为1）"
    ]
    df_rate = pd.DataFrame(columns=rate_col)

    for file in file_name:
        file_path = find_file_path(data_dir, file, file_type)
        if not file_path:
            return False

        raw_df = get_all_raw_data(file_path, sheet_name)
        if type(raw_df) is bool:
            return raw_df

        # 创业板和科创板对应不同的表头
        # if file == file_name[-1]:
        #     raw_df = raw_df[cyb_col]
        # elif file == file_name[0]:
        #     raw_df = raw_df[kcb_col]
        raw_df = raw_df[kcb_col]

        # 替换为一个标准表头
        raw_df.columns = df_col
        df_zero = pd.concat([df_zero, raw_df])

    stock_code_dict = dict()
    stock_code_list = list(set(df_zero["标准代码"]))
    ipo_date_data = list()
    len_code = len(stock_code_list)
    max_wind_curl = 30
    n_tmp = len_code // max_wind_curl
    for idx in range(n_tmp + 1):
        if idx * max_wind_curl + max_wind_curl < len_code:
            end_flag = idx * max_wind_curl + max_wind_curl
        else:
            end_flag = None
        stock_code_str = ",".join(
            stock_code_list[idx * max_wind_curl: end_flag]
        )
        ipo_date_data.extend(
            w.wsd(
                stock_code_str, "ipo_inq_enddate",
                "ED0TD", datetime.now().strftime("%Y-%m-%d")
            ).Data[0]
        )
        print("Please Wait 3s!")
        time.sleep(3)

    if len(stock_code_list) == len(ipo_date_data):
        for idx in range(len(stock_code_list)):
            if ipo_date_data[idx] is None:
                return False
            stock_code_dict[stock_code_list[idx]] = ipo_date_data[idx].strftime("%Y-%m-%d")
    else:
        print(print_info("W"), end=" ")
        print(len(stock_code_list), len(ipo_date_data))
        return False

    # 补充询价日信息
    # for code_item in df_zero["标准代码"]:
    #     df_zero[df_zero["标准代码"] == code_item]["询价日"] = stock_code_dict[code_item]
    #
    # df_zero.sort_values(by=["询价日", "股票名称"], inplace=True, ascending=[False, True])
    df_zero.sort_values(by=["股票名称"], inplace=True, ascending=[True])
    print(print_info(), end=" ")
    print("Get the Sample:\n{}".format(df_zero.head()))

    tzz_mc, sg_jg = df_col[3], df_col[4]

    # 获取全部的股票列表
    stock_list = df_zero[df_col[0]].tolist()
    union_stock_col = list(set(stock_list))
    union_stock_col.sort(key=stock_list.index)
    print(print_info(), end=" ")
    print("Get the stock list: \n{}".format(union_stock_col))

    # 获取无重复全部的投资者
    all_tzz_col = df_zero[tzz_mc].tolist()
    union_tzz_col = list(set(all_tzz_col))
    # 删除申购次数较小的投资者
    retain_list = get_small_size(df_zero, tzz_mc, min_num)
    union_tzz_col = list(
        set(union_tzz_col).intersection(set(retain_list))
    )

    print(print_info(), end=" ")
    print("Get the tzz list: \n{}".format(union_tzz_col))

    # 创建参与次数，报价正确次数的字典
    tzz_col_len = len(union_tzz_col)
    join_dict = dict(zip(union_tzz_col, [0] * tzz_col_len))
    s_right_dict = dict(zip(union_tzz_col, [0] * tzz_col_len))
    x_right_dict = dict(zip(union_tzz_col, [0] * tzz_col_len))

    # 对于每一个标的处理
    df_dict = dict()
    # 创建一个我司的处理标签
    df_ysi = pd.DataFrame(columns=["股票名称", "是否参与", "是否部分正确", "是否完全正确"])
    for stock in union_stock_col:
        print(print_info(), end=" ")
        print("Operator the stock: {}".format(stock))
        df_dict[stock] = df_zero[df_zero[df_col[0]] == stock]
        # df_group = get_df_group(df_dict[stock], tzz_mc, sg_jg)
        # if type(df_group) is bool:
        #     return df_group
        tzz_list = list(df_dict[stock][tzz_mc])

        ysi_dict = dict()
        ysi_dict["是否参与"] = 0
        ysi_dict["是否部分正确"] = 0
        ysi_dict["是否完全正确"] = 0
        for tzz_item in union_tzz_col:
            if tzz_item in tzz_list:
                join_dict[tzz_item] += 1
                state_tmp = list(set(df_dict[stock][df_dict[stock][tzz_mc] == tzz_item]["备注"]))
                dt_judge = not(any("低" in tmp for tmp in state_tmp))
                gt_judge = not(any("高" in tmp for tmp in state_tmp))
                if dt_judge and gt_judge:
                    s_right_dict[tzz_item] += 1
                    x_right_dict[tzz_item] += 1
                elif any("有" in tmp for tmp in state_tmp):
                    x_right_dict[tzz_item] += 1

                if tzz_item == my_comp_name:
                    ysi_dict["股票名称"] = stock
                    ysi_dict["是否参与"] = 1
                    if dt_judge and gt_judge:
                        ysi_dict["是否部分正确"] = 1
                        ysi_dict["是否完全正确"] = 1
                    elif any("有" in tmp for tmp in state_tmp):
                        ysi_dict["是否部分正确"] = 1

        df_ysi = df_ysi.append(ysi_dict, ignore_index=True)

    df_ysi.to_excel(my_comp_name + ".xlsx", index=None)

    # 所有投资者
    df_rate[rate_col[1]] = union_tzz_col
    # 其参与的总报价次数
    df_rate[rate_col[2]] = list(join_dict.values())
    # # 迎水加成
    # right_dict[my_comp_name] += ys_plus
    # 正确次数
    df_rate[rate_col[3]] = list(s_right_dict.values())
    df_rate[rate_col[4]] = list(x_right_dict.values())

    # 创建正确率列
    s_rate_list = list()
    x_rate_list = list()
    # rate_dict = dict(zip(union_tzz_col, [""] * len(union_tzz_col)))
    for item, j_v, sr_v, xr_v in zip(union_tzz_col, join_dict.values(), s_right_dict.values(), x_right_dict.values()):
        s_rate_value, x_rate_value = sr_v / j_v, xr_v / j_v
        # rate_dict[item] = "{:.2%}".format(s_rate_value)
        s_rate_list.append(s_rate_value)
        x_rate_list.append(x_rate_value)

    # 将两种方式计算的正确率导入
    df_rate[rate_col[-2]] = s_rate_list
    df_rate[rate_col[-1]] = x_rate_list

    # 优先正确率, 其次正确次数, 最后是参与次数
    df_rate.sort_values(
        by=[rate_col[-2], rate_col[-1], rate_col[3], rate_col[4], rate_col[2], rate_col[1]],
        inplace=True,
        ascending=[False, False, False, False, False, True]
    )
    df_rate[rate_col[0]] = list(range(1, len(union_tzz_col) + 1))
    print(print_info(), end=" ")
    print("Get right rate dataframe:\n{}".format(df_rate))

    judge = save_all_data(output_dir, save_name, file_type, df_rate)
    return judge


def get_time(date=False, utc=False, msl=3):
    if date:
        time_fmt = "%Y-%m-%d %H:%M:%S.%f"
    else:
        time_fmt = "%H:%M:%S.%f"

    if utc:
        return datetime.utcnow().strftime(time_fmt)[:(msl - 6)]
    else:
        return datetime.now().strftime(time_fmt)[:(msl - 6)]


def print_info(status="I"):
    return "\033[0;33;1m[{} {}]\033[0m".format(status, get_time())


def find_file_path(data_path, file_n, f_type=".xlsx"):
    if not os.path.exists(data_path):
        print(print_info("E"), end=" ")
        print("The data directory: {} is not existed, please check again!".format(data_path))
        return False
    else:
        print(print_info(), end=" ")
        print("Find the data in path: {}".format(data_path))

        # 仅考虑符合条件格式的文件列表
        dir_list = os.listdir(data_path)
        f_type_temp = f_type.split(".")[-1]
        dir_list = [item for item in dir_list if item.split(".")[-1] == f_type_temp]

        file_path = list()

        for item in dir_list:
            if file_n in item and "$" not in item:
                file_path.append(item)

        if len(file_path) == 0:
            print(print_info("E"), end=" ")
            print("No {} file in the data path: {}".format(file_n, data_path))
            return False
        elif len(file_path) >= 2:
            print(print_info("E"), end=" ")
            print("Not only one file include {} in the data path: {}".format(file_n, data_path))
            return False
        else:
            file_path = os.path.join(data_path, file_path[0])
            print(print_info(), end=" ")
            print("Get the file path: {}".format(file_path))

    return file_path


def get_all_raw_data(f_path, sheet_n="基础数据"):
    try:
        df = pd.read_excel(f_path, sheet_name=None, header=0)
        df_all = pd.DataFrame()
        # df = pd.read_excel(f_path, sheet_name=sheet_n, header=0)
        print(print_info(), end=" ")
        print("Successfully loading the file: {}".format(f_path))
        op_list = list()
        for sheet_item in list(df.keys()):
            if sheet_n in sheet_item:
                op_list.append(sheet_item)
                df_all = pd.concat([df_all, df[sheet_item]])
                print(print_info(), end=" ")
                print("Get Sheet: {}".format(sheet_item))
        return df_all
    except:
        print(print_info("E"), end=" ")
        print("Could not load the file: {}".format(f_path))
        return False


def get_df_group(df, index_n, price):
    try:
        index_name = index_n
        df_group = df.groupby(index_name)[price, "备注"].agg(lambda x: x.value_counts().index[0]).reset_index()
        df_group.set_index(index_name, inplace=True)
        # print(print_info(), end=" ")
        # print("The Group by DataFrame:\n {}".format(df_group))
        return df_group
    except:
        print(print_info("E"), end=" ")
        print("There is something wrong in DataFrame:\n {}".format(df.head()))
        return False


# 去除少于指定次数的投资者
def get_small_size(df, col_name, num=300):
    df_dict = df[col_name].value_counts().to_dict()
    col_list = list(set(df_dict.keys()))
    for key, value in df_dict.items():
        if value <= num:
            col_list.remove(key)
            print(print_info(), end=" ")
            print("Delete name: {}, value: {}".format(key, value))
    return col_list


def get_yes_or_no(df, tzz_mc, tzz_n, s_dict):
    df_temp = df[df[tzz_mc] == tzz_n]
    note_dict = dict(df_temp["备注"].value_counts())
    note_key = list(note_dict.keys())
    # note_sum = sum(note_dict.values())
    s_key = list(s_dict.keys())

    # 默认“有效报价”不标记
    note = ""
    # valid_len = 0
    for item in note_key:
        if s_key[-1] in item or "入围" == item:
            # valid_len = note_dict[item]
            note_key.remove(item)

    note_len = len(note_key)

    if note_len:
        # 清洗一个备注情况的字典，统计低价，高价，无效的情况
        value_dict = dict()
        for idx in range(len(s_key) - 1):
            for item in note_key:
                if s_key[idx] in item:
                    value_dict[s_key[idx]] = 0
        for item in note_key:
            for idx in range(len(s_key) - 1):
                if s_key[idx] in item:
                    value_dict[s_key[idx]] += note_dict[item]

        # 单个投资者的报价标记，注意优先级
        value_key = list(value_dict.keys())
        if s_key[0] in value_key:
            note = s_dict[s_key[0]]
        elif s_key[1] in value_key:
            note = s_dict[s_key[1]]
        # 未入围：
        elif value_key == "未入围":
            note = s_dict[s_key[0]]
        # 其他种种失败的报价情形
        elif s_key[2] in value_key or len(value_key) == 0:
            note = s_dict[s_key[2]]
        else:
            print(print_info("W"), end=" ")
            print("Unexpected events occurred in {}!".format(tzz_n))
    return note


def save_all_data(o_path, s_name, op_file_type, df):
    op_date = time.strftime('%Y%m%d', time.localtime(time.time()))
    file_name = s_name + op_date + op_file_type
    save_path = os.path.join(o_path, file_name)
    try:
        df.to_excel(save_path, index=None)
        print(print_info(), end=" ")
        print("Save to the path: {}".format(save_path))
    except:
        print(print_info("E"), end=" ")
        print("Can not save to the path: {}".format(save_path))
        return False

    return True


if __name__ == '__main__':
    root_path = os.path.abspath(".")
    data_dir_name = "raw_data"
    output_dir_name = "output"
    # 最小报价次数
    less_num = 0
    # 启动 wind
    w.start()
    w.isconnected()
    t_f = get_all_right_rate(root_path, data_dir_name, output_dir_name, less_num)
    if t_f:
        print(print_info("S"), end=" ")
        print("Success!")
    else:
        print(print_info("E"))
        print("Error!")
    w.close()
