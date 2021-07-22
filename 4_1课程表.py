# -*- coding:utf-8 -*-
"""
        File Name     :    4_1课程表.py
        Create Date   :    2021-07-21    20:00
"""
import pandas as pd
import re
import csv


def to_standardcsv(csvwriter):
    # 开始转换
    for i in range(1, df.shape[1]):  # 遍历列
        # print(f"星期{i}")
        for j in df.iloc[:, i].tolist():
            if "节次" in str(j):
                # 正则替换电话空值
                j = re.sub(r"电话.*?\n\n", "", j)
                # j = j.replace("\n", "")
                # print(j)
                # 正则匹配
                obj = re.compile(r"(.*?)\n(.*?)\n节次:(.*?) \n周次:(.*?)\n地点:(.*?)\n开课院系:(.*?)\n", re.S)
                cls = obj.findall(j)
                # print(cls)
                for cl in cls:
                    # 相应的正则后结果
                    name = cl[0]
                    day = str(i)
                    start = cl[2].split("-")[0]
                    end = cl[2].split("-")[1].split("节")[0]
                    teacher = cl[1]
                    location = cl[4]
                    week = cl[3]
                    result = [name, day, start, end, teacher, location, week]
                    # print(result)
                    # 写入csv
                    csvwriter.writerow(result)
        # print("+" * 50)


def main():
    # 打开转换模板，进行写入前准备
    f = open("课程表.csv", "w", encoding="utf-8", newline="")
    csvwriter = csv.writer(f)
    headers = ['课程名称', '星期', '开始节数', '结束节数', '老师', '地点', '周数']
    csvwriter.writerow(headers)

    # 转换函数
    to_standardcsv(csvwriter)

    f.close()
    print("转换完成")


if __name__ == '__main__':
    # 导入原课程表 跳过第一行
    df = pd.read_excel(r"C:\Users\huanity\Desktop\课程表20210721-182312.xls", skiprows=1)
    # print(df.info())
    # print(df)

    main()
