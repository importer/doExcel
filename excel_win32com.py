#!/usr/bin/python

 

from win32com.client import Dispatch

import os

pwd = os.getcwd()

xlApp = Dispatch('Word.Application')

# xlApp.Visible = True

xlBook = xlApp.Workbooks.Add()

xlApp.Worksheets.Add().Name = 'test'

xlSheet = xlApp.Worksheets('test')

xlSheet.Cells(1,1).Value = 'title'

xlSheet.Cells(2,1).Value = 123

xlBook.SaveAs(pwd + '\\demo.xlsx')

xlApp.Quit() # exit app