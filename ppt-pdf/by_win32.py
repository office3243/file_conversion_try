import sys  
import os  
import glob  
import win32com.client  
  
def convert(filename, formatType = 32):  
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")  
    powerpoint.Visible = 1  
    newname = os.path.splitext(filename)[0] + ".pdf"  
    deck = powerpoint.Presentations.Open(filename)          
    deck.SaveAs(newname, formatType)  
    deck.Close()  
    powerpoint.Quit()


inp = "C:/Users/eWay/Desktop/aamer/Presentations-Tips.ppt"
    
convert(inp)

