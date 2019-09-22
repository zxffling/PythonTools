# -*- encoding: utf-8 -*-
"""
@File    : ExcelToOracle.py
@Time    : 2019/7/22 22:13
@Author  : xf_chief
@Email   : zxffling@gmail.com
@Software: PyCharm
将Excel数据在Oracle数据库中建立一个表，并导入到表中
"""
import cx_Oracle
import csv
import xlrd
import os
import logging

class ImportOracle(object):

    def __init__(self):
        self.filename = ""

    def inoracle(self):
        pass

    def ConnOracle(self):
        dblink = input('请输入数据库连接信息，格式为"testuser/123456@localhost/orcl":')
        os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.utf8'  # 解决编码问题
        # conn = cx_Oracle.connect('zhzs/Ahzd2018@192.168.8.66/zhzs')  # 连接数据库
        conn = cx_Oracle.connect(dblink)  # 连接数据库
        cursor = conn.cursor()  # 创建游标

        fields = []
        # fields1 = [i + ' varchar2(200)' for i in self.title]
        # for i in self.title1:
        #     fields.append(i)
        # print(fields)
        # 解决数据类型的问题
        for n in range(len(self.title1)):
            f = (type(self.title1[n]))  # 获取某列的数据类型
            # print(f)
            if f is float:  # 这里可以根据实际表格数据类型来进行判断
                fields.append(self.title[n] + ' number')
            else:
                fields.append(self.title[n] + ' varchar2(4000)')
        fields_str = ', '.join(fields)  # 把create table 的字段（列）进行格式化
        sql = 'create table %s (%s)' % (self.table_name, fields_str)
        print(sql)  # 打印创建table的SQL
        while 1:
            try:
                cursor.execute(sql)
                print("临时表创建完成！")
                break
            except:
                cursor.execute('drop table {}'.format(
                    self.table_name))  # 通过一个循环来避免tbale名已经被占用的问题，尝试创建，如果存在就执行drop table,直到可以创建就break退出循环。
        logging.info('info message')
        a = [':%s' % i for i in range(len(self.title) + 1)]
        value = ','.join(a[1:])
        sql = 'insert into %s values(%s)' % (self.table_name, value)  # 导入数据语句定义
        print('开始插入数据，请耐心等待...')  # 打印导入数据的sql
        print(self.data[2])  # 打印数据第三行数据（随便打印哪一行都可以，为了验正取到了数据）
        cursor.prepare(sql)
        i_num = 0
        for i in self.data:
            cursor.execute(None, i)  # 执行数据插入，这里是循环单条插入的方式。
            print(i)
            i_num = i_num + 1
            # cursor.executemany(None, self.data)
        # 这是多条插入的方式，但是我在进行多条插入的时候又遇到了浮点数精度不一致的问题。。。

        print('数据插入完成，共插入' + str(i_num) + '条记录')
        cursor.close()  # 关闭游标
        conn.commit()  # 提交执行，增删改都需要进行提交。
        conn.close()  # 关闭数据库连接


class ImportOracleCsv(ImportOracle):
    def inoracle(self):
        with open(self.filename, 'rb') as f:
            reader = csv.reader(f)
            contents = [i for i in reader]

        title = contents[0]
        title1 = contents[1]
        data = contents[1:]

        return (title, title1, data)


# Csv格式读取

class ImportOracleExcel(ImportOracle):
    def inoracle(self):
        wb = xlrd.open_workbook(self.filename)
        sheet1 = wb.sheet_by_index(0)

        title = sheet1.row_values(0)
        title1 = sheet1.row_values(1)
        data = [sheet1.row_values(row) for row in range(1, sheet1.nrows)]
        return (title, title1, data)


# excle格式读取

class ImportError(ImportOracle):
    def inoracle(self):
        print('Undefine file type')
        return 0


# 异常处理

class ChooseFactory(object):
    choose = {}
    choose['csv'] = ImportOracleCsv()
    choose['xlsx'] = ImportOracleExcel()
    choose['xls'] = ImportOracleExcel()

    def choosefile(self, ch):
        if ch in self.choose:
            op = self.choose[ch]
        else:
            op = ImportError()

        return op


# 根据传入的表格的格式去判断用CSV格式读取还是excle格式读取，如果格式不是CSV和EXCLE则跳转报错

if __name__ == "__main__":
    # file_name = 'D:\\DevelopmentData\\Github\\PythonProject\\Tools\\ExcelToOracle\\demo.xlsx'  # 传入表格名
    file_name = input('请输入Excel文件路径，如"D:\gds\\test.xls"：')
    # table_name = 'tmp_spfbaxx'  # 传入需要创建的数据库table名
    table_name = input('请输入存储的数据库表名，如"tmp_spfbaxx":')  # 传入需要创建的数据库table名
    op = file_name.split('.')[-1]  # 切片所传入的表格名字字符串，为了去判断是什么格式的表格
    factory = ChooseFactory()  # 实例判断格式的class
    cal = factory.choosefile(op)  # 调用函数
    cal.filename = file_name
    (cal.title, cal.title1, cal.data) = cal.inoracle()
    cal.table_name = table_name
    cal.ConnOracle()
os.system('pause')
