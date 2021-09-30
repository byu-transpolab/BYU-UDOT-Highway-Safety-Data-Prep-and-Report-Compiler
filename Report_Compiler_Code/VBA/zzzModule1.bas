Attribute VB_Name = "zzzModule1"

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWorkbook.Worksheets("BlankReport").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BlankReport").Sort.SortFields.Add Key:=Range("A58" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("BlankReport").Sort
        .SetRange Range("A58:A62")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveWorkbook.Worksheets("BlankReport").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BlankReport").Sort.SortFields.Add Key:=Range("A58" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("BlankReport").Sort
        .SetRange Range("A58:A83")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
