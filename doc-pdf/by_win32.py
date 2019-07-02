import sys
import os
import win32com.client as cc

wdFormatPDF = 17

in_file = "C:/Users/eWay/Desktop/aamer/sample.doc"
out_file = "C:/Users/eWay/Desktop/aamer/sample.pdf"

word = cc.Dispatch('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
