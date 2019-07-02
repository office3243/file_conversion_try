from win32com import client

inp = "C:/Users/eWay/Desktop/aamer/tests-example.xls"
out = "C:/Users/eWay/Desktop/aamer/tests-example.pdf"

xlApp = client.Dispatch("Excel.Application")
books = xlApp.Workbooks.Open(inp)
ws = books.Worksheets[0]
ws.Visible = 1
ws.ExportAsFixedFormat(0, out)
