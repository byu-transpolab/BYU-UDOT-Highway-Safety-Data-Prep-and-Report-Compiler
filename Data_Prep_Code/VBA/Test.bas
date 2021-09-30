Attribute VB_Name = "Test"
Sub CombineFCFiles()

Dim SR, FedAid, LastRow, wd

wd = ActiveWorkbook.Sheets("Inputs").Range("I2")
wd = replace(wd, " ", "_")
wd = replace(wd, "/", "\")
wd = wd & "\FunctionalClassCombined"

'Create FunctionalClassCombined Workbook
    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=wd

'Open SR & FedAid FC Workbooks
    Workbooks.Open form_CreateIntData.txtFCSR
    SR = ActiveWorkbook.Name
    Workbooks.Open form_CreateIntData.txtFCFED
    FedAid = ActiveWorkbook.Name

'Copy a Range of Data in SR Workbook
    Workbooks(SR).Worksheets("Page1_1").Range("A3:G1000").Copy

'PasteSpecial Values Only into FC Combined
    Workbooks("FunctionalClassCombined").Worksheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteValues
    
LastRow = Workbooks("FunctionalClassCombined").Worksheets("Sheet1").Range("A1").End(xlDown).Address
  
'Copy a Range of Data in FedAid Workbook
    Workbooks(FedAid).Worksheets("Page1_1").Range("A4:G1000").Copy

'PasteSpecial Values Only Into FC Combined
    Workbooks("FunctionalClassCombined").Worksheets("Sheet1").Range(LastRow).PasteSpecial Paste:=xlPasteValues
    Workbooks(FedAid).Activate
    ActiveSheet.Range("A1").Copy
  
'Close SR, FedAid, & FunctionalClassCombined Workbooks
    Workbooks(SR).Close savechanges:=False
    Workbooks(FedAid).Close savechanges:=False
    Workbooks("FunctionalClassCombined").Close savechanges:=True


End Sub
