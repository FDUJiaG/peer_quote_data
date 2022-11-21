import os
from op_ns_data import op_ns_data, print_info
from op_all_tzz import op_all_tzz
from WindPy import w


if __name__ == '__main__':
    op_list = os.listdir("raw_data") + os.listdir("raw_data/history")
    w.start()
    w.isconnected()
    for item in op_list:
        if "xlsx" == item.split(".")[-1]:
            f_name = item.split("_")[0]
            tf_gz = op_ns_data(f_name)
            tf_all = op_all_tzz(f_name)
            if tf_gz and tf_all:
                print(print_info("S"), end=" ")
                print("Success!")
            else:
                print(print_info("E"))
                print("Error!")
    w.close()
