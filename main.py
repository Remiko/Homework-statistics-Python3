import xlwt  # xlwt可以对excel进行写操作
import xlrd  # xlrd可以对excel进行读操作
from xlutils3.copy import copy  # xlutils3对excel进行二次重写操作
import re  # 正则表达式模块
import os  # 系统模块

path = r'C:\Users\Administrator\Desktop\18物联网班c语言第三次作业'  # 提取文件名字的文件夹的位置
file_name = r'C:\Users\Administrator\Desktop\点名\第三次作业统计.xls'  # 最后要创建的excel的路径
main_name = r'C:\Users\Administrator\Desktop\物联网班名单.xls'  # 班级名单的路径
n_str = r'(1812020)?\d{2,3}'  # 此正则可以匹配学号
rows0 = []  # 存储提取文件名字(提取后的学号)后的列表
numbers = []  # 存储物理网名单.xls中的学号
students = []  # 存储物理网名单.xls中的学生名字


def fileNM():  # 引用文件名并筛选
    for root, dirs, files in os.walk(path):  # 调用文件名
        for temp in files:
            judge = re.compile(n_str)
            chr = judge.search(temp).group()  # 用正则筛选文件名
            temp = alter(chr)
            rows0.append(temp)  # 存入rows0的列表


def alter(chr):  # 对不符合规则的数据进行处理
    a = []
    if len(chr) == 2:
        chr = '18120201' + chr
    elif len(chr) == 3:
        chr = '1812020' + chr
    return chr


def write_final():  # 第一次写入数据到第三次作业统计.xls中
    ex = xlwt.Workbook()  # 初始化workbook对象
    sheet = ex.add_sheet('统计', cell_overwrite_ok=True)  # 创建表
    for i in range(len(numbers)):
        sheet.write(i, 0, numbers[i])
        sheet.write(i, 1, students[i])
    for j in range(len(rows0)):
        for m in range(len(numbers)):
            if rows0[j] == numbers[m]:
                sheet.write(m, 2, '已交')
    ex.save(file_name)  # 保存表


def read_ex():  # 读取物理网名单.xls中的学号和姓名
    table = xlrd.open_workbook(main_name)
    sheet = table.sheet_by_name('2018级')
    rows = sheet.nrows
    for x in range(3, rows):
        num_value = sheet.cell(x, 1).value
        st_name = sheet.cell(x, 2).value
        numbers.append(num_value)
        students.append(st_name)


def who_finish():  # 判断新建表中哪儿交了
    table = xlrd.open_workbook(file_name)
    sheet = table.sheet_by_name('统计')
    rows = sheet.nrows
    for x in range(rows):
        value = sheet.cell(x, 2).value
        if value != '已交':
            write_unfinished(x)


def write_unfinished(value):  # 判断新建表中哪为空
    table = xlrd.open_workbook(file_name, formatting_info=True)
    xls = copy(table)
    sheet1 = xls.get_sheet('统计')
    sheet1.write(value, 2, '未交')
    xls.save(file_name)


if __name__ == '__main__':
    fileNM()
    read_ex()
    write_final()
    who_finish()
