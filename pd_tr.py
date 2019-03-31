import pandas as pd
from pyexcel_xls import get_data
import os
import shutil
import numpy as np

# 單純去除資料雜值
str_re = "\u3000\u3000"
# 資料夾路徑
input_dir_name = "./input_data/"
output_dir_name = "./output_data/"
finish_dir_name = "./finish_data/"
dir_list = [input_dir_name, output_dir_name, finish_dir_name]
for i in dir_list:
    if not os.path.isdir(i):
        os.mkdir(i)
# 查看所有input資料
input_list = os.listdir(input_dir_name)
for file_name in input_list:
    # 找檔案資料
    data = get_data(input_dir_name + file_name)
    # .keys() pyexcel_xls.get_data專屬,來找excel的sheet
    for i in data.keys():
                # sheet得數量  就是 i
                # 欄位  如果欄位有雜值,去除,沒有的話不變
                column_list = [column.replace(str_re, "") if (str_re in column) else column for column in data[i][4]]
                # 資料取得年(西元),月份
                t = data[i][1][1].split("/")
                year = str(int(t[0]) + 1911)
                months = t[1] + "~" + t[3] + "月"
                # 資料取得花名
                flower_name = data[i][3][1].split(" ")[1]
                # 單純show資料名稱
                print("正處裡:" + flower_name + year + "年" + months + "資料")

                # 單純因為資料前五筆不重要,每行的最後一欄位是多餘的
                data_arr = np.array(data[i][5:])[:, :-1]
                # 計算有幾筆資料
                count = len(data_arr)
                # 把資料塞進DataFrame
                df = pd.DataFrame(data_arr, columns=column_list)

                # 原本的做法是用for迴圈做Series,但發現nparry超快
                # for d in data[i][5:-1]:
                #     # 計算有幾筆資料+1
                #     count += 1
                #     # 把data加入Series
                #     s = pd.Series(d[:-1], column_list)
                #     print(s)
                #     # 把Series加入dataframe,忽略編碼
                #     df = df.append(s, ignore_index=True)

                # 單純計算有幾筆資料show
                print("總共:" + str(count) + "筆")
                # 存資料換編碼為utf8,並不加入預設編碼
                # 輸出檔案到output
                df.to_csv(output_dir_name + flower_name + year + "年" + months + "資料-" + str(count) + "筆.csv", encoding="utf-8",
                          index=False)
                # 單純show資料名稱
                print("處裡完成:" + flower_name + year + "年" + months + "資料-" + str(count) + "筆")

                # 轉移檔案到finish
                shutil.move(input_dir_name + "/" + file_name, finish_dir_name)
                print("----------------------------------------------")  # 單純分隔線
