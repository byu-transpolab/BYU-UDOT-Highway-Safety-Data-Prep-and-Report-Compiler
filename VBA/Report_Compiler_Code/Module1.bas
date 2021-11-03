Attribute VB_Name = "Module1"
Sub sorttest()
Attribute sorttest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sorttest Macro
'

'
    Range(Range("D2"), Range("D2").End(xlDown)).Select
    ActiveWorkbook.Worksheets("SegKey").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SegKey").Sort.SortFields.Add Key:=Range("D2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SegKey").Sort
        .SetRange Range("D2:D233")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=84
End Sub
