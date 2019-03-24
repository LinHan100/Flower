import pandas as pd
from pyexcel_xls import get_data
import os

c = 1  # 檔案數量起始值
str_re = "\u3000\u3000"  # 單純去除資料雜值
file_name = "花卉產品日交易行情 (" + str(c) + ").xls"  # 檔案名稱
is_ok = os.path.exists(file_name)  # 查看檔案是否存在
while is_ok:  # 如果檔案不存在,結束
    data = get_data(file_name)  # 找檔案資料
    for i in data.keys():  # .keys() pyexcel_xls.get_data專屬,來找excel的sheet
        column_list = [j.replace(str_re, "") if (str_re in j) else j for j in data[i][4]]  # 如果欄位有雜值,去除
        year = str(int(data[i][1][1].split("/")[0]) + 1911)  # 單純資料取年(西元)
        flower_name = data[i][3][1].split(" ")[1]  # 單純資料取花名
        print("正處裡:" + flower_name + year + "年資料")  # 單純資料名稱show
        df = pd.DataFrame(columns=column_list)  # 塞column_list進pd.dataframe的欄位
        count = 0  # 計算有幾筆資料初始值
        for d in data[i][5:]:  # 單純因為資料前五筆不重要
            count += 1  # 計算有幾筆資料+1
            d.pop(-1)  # 單純因為資料最後一筆是多的
            s = pd.Series(d, column_list)  # 把data加入Series
            df = df.append(s, ignore_index=True)  # 把Series加入dataframe,忽略編碼
        print("總共:" + str(count) + "筆")  # 單純計算有幾筆資料show

        df.to_csv(flower_name + year + "年資料-" + str(count) + "筆.csv", encoding="utf-8",
                  index=False)  # 存資料換編碼為utf8,並不加入預設編碼
        print("處裡完成:" + flower_name + year + "年資料-" + str(count) + "筆")  # 單純資料名稱show
        os.remove(file_name)  # 移除載入檔案
        c += 1  # 找下個資料
        is_ok = os.path.exists(file_name)  # 查看檔案是否存在
        print("----------------------------------------------")  # 單純分隔線
