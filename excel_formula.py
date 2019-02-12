#!/usr/bin/env python     
# -*- coding: utf-8 -*-     
from win32com.client import Dispatch    
import win32com.client
from db2_conn import get_data,transposition
import datetime
from datetime import timedelta
import sys

class easyExcel:    
      """A utility to make it easier to get at Excel.    Remembering  
      to save the data is your problem, as is    error handling.  
      Operates on one workbook at a time."""    
      def __init__(self, filename=None):  #打开文件或者新建文件（如果不存在的话）  
          self.xlApp = Dispatch('Excel.Application.16')    
          if filename:    
              self.filename = filename    
              self.xlBook = self.xlApp.Workbooks.Open(filename)    
          else:    
              self.xlBook = self.xlApp.Workbooks.Add()    
              self.filename = ''  
        
      def save(self, newfilename=None):  #保存文件  
          if newfilename:    
              self.filename = newfilename    
              self.xlBook.SaveAs(newfilename)    
          else:    
              self.xlBook.Save()        
      def close(self):  #关闭文件  
          self.xlBook.Close(SaveChanges=0)    
          del self.xlApp    
      def getCell(self, sheet, row, col):  #获取单元格的数据  
          "Get value of one cell"    
          sht = self.xlBook.Worksheets(sheet)    
          return sht.Cells(row, col).Value    
      def setCell(self, sheet, row, col, value):  #设置单元格的数据  
          "set value of one cell"    
          sht = self.xlBook.Worksheets(sheet)    
          sht.Cells(row, col).Value = value  
      def setCellformat(self, sheet, row, col):  #设置单元格的数据  
          "set value of one cell"    
          sht = self.xlBook.Worksheets(sheet)
          print('Font:') 
          print(sht.Cells(row, col+1).Font.Size)   
          sht.Cells(row, col).Font.Size = sht.Cells(row, col+1).Font.Size  
          sht.Cells(row, col).Font.Bold = True#是否黑体  
          sht.Cells(row, col).Name = "Arial"#字体类型  
          sht.Cells(row, col).Interior.ColorIndex = 3#表格背景  
          #sht.Range("A1").Borders.LineStyle = xlDouble  
          sht.Cells(row, col).BorderAround(1,4)#表格边框  
          sht.Rows(3).RowHeight = 30#行高  
          sht.Cells(row, col).HorizontalAlignment = -4131 #水平居中xlCenter  
          sht.Cells(row, col).VerticalAlignment = -4160 #  
      def deleteRow(self, sheet, row):  
          sht = self.xlBook.Worksheets(sheet)  
          sht.Rows(row).Delete()#删除行  
          sht.Columns(row).Delete()#删除列  
      def getRange(self, sheet, row1, col1, row2, col2):  #获得一块区域的数据，返回为一个二维元组  
          "return a 2d array (i.e. tuple of tuples)"    
          sht = self.xlBook.Worksheets(sheet)  
          return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value    
      def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  #插入图片  
          "Insert a picture in sheet"    
          sht = self.xlBook.Worksheets(sheet)    
          sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)    
      
      def cpSheet(self, before):  #复制工作表  
          "copy sheet"    
          shts = self.xlBook.Worksheets    
          shts(1).Copy(None,shts(1))    

def write_excel(acct_month):

    file_name = u'data\\数据需求表_201805.xls'
    file_to_name=u'data\\数据需求表_%s.xls'%acct_month
    xls = easyExcel(file_name)
    acct_month_now=str(int(acct_month)+1)
    lastDay = (datetime.datetime.strptime('201901','%Y%m')-timedelta(days= 1)).strftime('%Y%m%d')
    print(lastDay)
    sql = "SELECT SHEET_NAME,TABLE_NAME,replace(replace(DATA_SQL,'20180228','%s'),'201802','%s') DATA_SQL,ROW_TYPE,ROW_POSITION,REMARK,OFFSET FROM bass15.bass1_Multi_cost_dim "%(lastDay,acct_month)
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
        line_no, row_no = int(line_no)+1, int(row_no)+1
        print('line_no:%r,row:%r' % (line_no, row_no))
        for line in dt_content:
            row = row_no  # 控制列数
            for j in line:
                # 写入excel，第一个值是行，第二个值是列
                xls.setCell(sheet_index,line_no,row,j)
                row += 1
            line_no += 1
    xls.save(file_to_name)    
    xls.close()

def main():
      #PNFILE = r'c:/screenshot.bmp'  
      xls = easyExcel(r'data\\数据需求表_201802tt.xls')     
      #xls.addPicture('Sheet1', PNFILE, 20,20,1000,1000)    
      #xls.cpSheet('Sheet1') 
      Sheet1='产品和产品元素业务量（总）' 
      row,col=128,5
      xls.setCell(Sheet1,129,6,8888)  
      # xls.setCellformat(Sheet1,row,col)  
      xls.save()    
      xls.close()

   
if __name__ == "__main__":
  acct_month=sys.argv[1]
  print(acct_month)
  write_excel(acct_month)    