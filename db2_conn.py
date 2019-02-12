# coding=utf-8
from __future__ import print_function
import ibm_db

from xlutils.copy import copy
from xlutils.filter import process, XLRDReader, XLWTWriter
import xlrd
import xlwt
import os


def get_data(sql):
    conn = ibm_db.connect(
        "DATABASE=ynbiyj;HOSTNAME=XXX;PORT=10722;PROTOCOL=TCPIP;UID=XXX;PWD=XXX;", "", "")
    data_result = []
    if conn:
        try:
            # sql = "SELECT CURRENT DATE FROM SYSIBM.SYSDUMMY1"
            stmt = ibm_db.exec_immediate(conn, sql)
            result = ibm_db.fetch_tuple(stmt)
            while (result):
                # print ("日期 :", str(result[0]) +'\n')
                # print ("data :", result)
                data_result.append(result)
                # print ('-----------------')
                result = ibm_db.fetch_tuple(stmt)
        except Exception, e:
            print("Transaction couldn't be completed:", e, ibm_db.stmt_errormsg())
        else:
            print("Transaction complete.")
    return data_result


# 转置
def transposition(grid):
    '''转置数据'''
    grid = [[row[i] for row in grid] for i in range(len(grid[0]))]
    # print(grid)
    return grid


# 数据偏移问题
def offset(data, offset_id):
    '''处理数据头'''
    pass


def copy2(wb):
    w = XLWTWriter()
    process(XLRDReader(wb, 'unknown.xls'), w)
    return w.output[0][1], w.style_list


def write_excel():
    file_name = u'data/数据需求表_201802t.xls'
    book = xlrd.open_workbook(u'data/数据需求表_201801.xls', formatting_info=True, on_demand=True)
    sheet_1 = book.get_sheet(17)  # 获取到第一个sheet页
    # 复制一个excel
    # new_book = copy(book)#复制了一份原来的excel
    new_book, s = copy2(book)  # 复制了一份原来的excel
    # 通过获取到新的excel里面的sheet页

    borders = xlwt.Borders()  # 创建一个边框对象
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_RIGHT  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER
    styles = s[sheet_1.cell_xf_index(13, 5)]
    styles.borders = borders
    styles.alignment = alignment
    print(styles)

    wh = u'产品和产品元素业务量（TDS）'
    # sql = "SELECT * FROM bass15.bass1_Multi_cost_dim where SHEET_NAME=\'%s\'" % wh
    sql = "SELECT * FROM bass15.bass1_Multi_cost_dim "
    file_name = u'data/数据需求表_201802t.xls'
    content = get_data(sql)
    for i in content:
        print(i)
        sql = i[2]
        sheet_index = i[0]
        row_position = i[4]
        dt_content = get_data(sql)
        # 处理转置
        if i[3] == '2':
            dt_content = transposition(dt_content)
        # 处理偏移量的问题
        elif i[3] == '3':
            offset_id = 6
            dt_content = offset(dt_content, offset_id)
        print('----dt_content----')
        print(dt_content)
        print('----sheet_index----')
        print(sheet_index)
        print('----row_position----')
        print(row_position)

        line_no = 0  # 控制行数
        line_no, row_no = row_position.split(',')
        line_no, row_no = int(line_no), int(row_no)
        print('line_no:%r,row:%r' % (int(line_no), int(row_no)))
        sheet = new_book.get_sheet(sheet_index)  # 获取到第一个sheet页
        for line in dt_content:
            row = row_no  # 控制列数
            for j in line:
                # 写入excel，第一个值是行，第二个值是列
                sheet.write(line_no, row, j, styles)
                row += 1
            line_no += 1
    # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
    # book.release_resources()
    new_book.save(file_name)

def write_excel1():
    file_name = u'data/数据需求表_201802tt.xls'
    book = xlrd.open_workbook(u'data/数据需求表_201801.xls', formatting_info=True, on_demand=True)
    sheet_1 = book.get_sheet(17)  # 获取到第一个sheet页
    # 复制一个excel
    # new_book = copy(book)#复制了一份原来的excel
    new_book, s = copy2(book)  # 复制了一份原来的excel
    # 通过获取到新的excel里面的sheet页

    borders = xlwt.Borders()  # 创建一个边框对象
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_RIGHT  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER
    styles = s[sheet_1.cell_xf_index(13, 5)]
    styles.borders = borders
    styles.alignment = alignment
    print(styles)

    sheet_index = u'数据校验'

    line_no = 0  # 控制行数
    dt_content=[(u'''=IF(SUMPRODUCT(--('资产分类汇总表（TD公司）'!J9:J50008<'资产分类汇总表（TD公司）'!K9:K50008))=0,"正确","错误")''',)]
    line_no, row_no = 14,5
    print('line_no:%r,row:%r' % (int(line_no), int(row_no)))
    sheet = new_book.get_sheet(sheet_index)  # 获取到第一个sheet页
    for line in dt_content:
        row = row_no  # 控制列数
        for j in line:
            # 写入excel，第一个值是行，第二个值是列
            sheet.write(line_no, row, j, styles)
            row += 1
        line_no += 1
    # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
    # book.release_resources()
    new_book.save(file_name)

def main():
    write_excel1()


if __name__ == '__main__':
    main()
