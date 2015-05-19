# -*- coding: utf8 -*-

import xlrd
import pymysql

host = ""
user = ""
passwd = ""
dbname = ""
filename = ""

def write_db(host, user, password, dbname,data):
    db = pymysql.connect(host, user, passwd, dbname)
    cursor = db.cursor()
    try:
        for item in data:
            sql = "update students set username = '%s', student_num = '%s' where username = '%s'" % (item["xh"],item["xh"], item["yxh"])
            cursor.execute(sql)
        db.commit()
    except:
        db.rollback()
    db.close()

def read_xls(filename):
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    for col_index in range(sheet.ncols):
        if sheet.cell_value(0, col_index) == u"预修号":
            yxh_index = col_index
        elif sheet.cell_value(0, col_index) == u"学号":
            xh_index = col_index
    data = []
    if yxh_index > 0 and xh_index > 0:
        for row_index in range(1,sheet.nrows):
            yxh = sheet.cell_value(row_index, yxh_index)
            xh =sheet.cell_value(row_index, xh_index)
            data.append({
                "yxh":yxh,
                "xh":xh
            })
    return data


if __name__ == '__main__':
    data = read_xls(filename)
    write_db(host,user,passwd,dbname,data)




