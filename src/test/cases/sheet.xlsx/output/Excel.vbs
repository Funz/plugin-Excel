Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("D:\tmp\BTSES106\55205\spool.1\sheet.xlsx", 0, True) 
xl.DisplayAlerts = False

wb.Worksheets(1).Range("A1").Value = 1
wb.Worksheets(1).Range("A2").Value = 4
wb.RefreshAll
WScript.StdOut.WriteLine("tata=" & wb.Worksheets(1).Range("A3").Value)

wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing
