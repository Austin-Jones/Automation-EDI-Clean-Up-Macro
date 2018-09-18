Dim ObjExcel, ObjWB
Set ObjExcel = CreateObject("excel.application")
'vbs opens a file specified by the path below
Set ObjWB = ObjExcel.Workbooks.Open("C:\edi\ediAuto.xlam")
objExcel.Application.Run "ediAuto.xlam!Module1.main"

ObjWB.Close False
ObjExcel.Quit
Set ObjExcel = Nothing