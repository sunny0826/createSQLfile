#!/usr/bin/env python
#encoding: utf-8
#Author: guoxudong
import xlrd
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

def open_excel(file='test.xls'):
    try:
        data = xlrd.open_workbook(file)  # 打开excel文件
        return data
    except Exception, e:
        print str(e)


def excel_table_bycol(file='', colindex=[0], table_name='Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(table_name)  # 获取excel里面的某一页
    nrows = table.nrows  # 获取行数
    t_name = table.row_values(0)[0].encode('utf8') #表名
    colnames = table.row_values(1)  # 获取第一行的值，作为key来使用
    list = []
    # （2，nrows）表示取第二行以后的行，第一行为表名，第二行为表头
    list.append(t_name)
    list.append(colnames)
    for rownum in range(2, nrows):
        row = table.row_values(rownum)
        if row:
            app = []
            for i in colindex:
                app.append(str(row[i]).encode("utf-8") )
            list.append(app)  # 将字典加入列表中去
    return list


def main(file_name,colindex):
    # colindex为需要插入的列
    tables = excel_table_bycol(file_name,colindex, table_name=u'Sheet1')
    t_name = tables.pop(0)
    key_list = ','.join(tables.pop(0)).encode('utf8')   #list转为str
    sql_line = "INSERT INTO "+t_name+"（"+key_list+"）VALUE"
    line = ''
    for info in tables:
        content = ','.join(info)
        if line != '':
            line =line + ',(' + content + ')'
        else:
            line = '('+content+')'
    sql_line = sql_line + line + ';'
    with open('./sql_result/insert#' + t_name + '.sql', 'w') as f:  # 创建sql文件，并开启写模式
        f.write(sql_line)  # 往文件里写入sql语句

if __name__ == "__main__":
    file_name = './xls/test.xls'          #导入xls文件名
    colindex = [0, 1, 2, 3, 4]      #需要插入的列
    main(file_name,colindex)