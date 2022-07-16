"""
Created on Fri Apr 10 07:48:42 2020

@author: e.martinez
"""
import xlwings as xw

def gapfiller_xl(): 

    file = r'C:\Users\e.martinez\Documents\gl.xlsX'
    wb = xw.Book(file)
    sht = wb.sheets('Sheet1')
    rng = sht.range('A1:A4860')

    for x in rng:
        if x.value == None:
            x.value = x.offset(-1,0).value
            
gapfiller_xl()
            