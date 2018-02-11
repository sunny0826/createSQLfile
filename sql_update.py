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


def excel_table_bycol(file='', where=[0], colindex=[0], table_name='Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(table_name)  # 获取excel里面的某一页
    nrows = table.nrows  # 获取行数
    t_name = table.row_values(0)[0].encode('utf8') #表名
    colnames = table.row_values(1)  # 获取第一行的值，作为key来使用
    list = []
    # （2，nrows）表示取第二行以后的行，第一行为表名，第二行为表头
    for rownum in range(2, nrows):
        row = table.row_values(rownum)
        if row:
            whe = {}
            for n in where:
                whe[str(colnames[n]).encode("utf-8")] = str(row[n]).encode("utf-8")  #输入的筛选字段
            app = {}
            for i in colindex:
                app[str(colnames[i]).encode("utf-8")] = str(row[i]).encode("utf-8")  # 将数据填入一个字典中，同时对数据进行utf-8转码，因为有些数据是unicode编码的
            list.append({'where':whe,'app':app})  # 将字典加入列表中去
    return list,t_name


def main(file,where,colindex):
    # colindex为需要更新的列，where为筛选的列
    tables = excel_table_bycol(file,where,colindex, table_name=u'Sheet1')
    with open('./sql_result/update#'+tables[1]+'.sql', 'w') as f:    # 创建sql文件，并开启写模式
        for info in tables[0]:
            sql_line = "UPDATE "+tables[1]+" SET"
            apps = info.get('app')
            for key,value in apps.items():
                if sql_line.endswith('SET'):
                    sql_line += " "+key+"='"+value+"' "
                else:
                    sql_line += ", " + key + "='" + value + "' "
            sql_line += " WHERE"
            where = info.get('where')
            for key, value in where.items():
                if sql_line.endswith('WHERE'):
                    sql_line += " "+key+"='"+value+"' "
                else:
                    sql_line += "AND " + key + "='" + value + "' "
            sql_line+="\n"
            f.write(sql_line)  # 往文件里写入sql语句

if __name__ == "__main__":
    file_name = './xls/test.xls'  # 导入xls文件名
    where = [0,1,2]         # 条件字段
    colindex = [3, 4]       # 需要插入的列
    main(file_name,where,colindex)