import comtypes.client

def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


inp = "C:/Users/eWay/Desktop/aamer/Presentations-Tips.ppt"
out = "C:/Users/eWay/Desktop/aamer/Presentations-Tips.pdf"

PPTtoPDF(inp, out)
