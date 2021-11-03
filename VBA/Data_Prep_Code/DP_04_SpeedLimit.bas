Attribute VB_Name = "DP_04_SpeedLimit"
Option Explicit

'PMDP_04_SpeedLimit Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets

Sub Run_SpeedLimit()
'Call Run_SpeedLimit macro:
'   (1) Formats the Speed Limit data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define variables
Dim row1, col1, col2 As Integer 'column number
Dim nrow As Integer 'nrow number
Dim Beg_MP, End_MP, Route, Direction, Speed_Limit As String

'Screen Updating OFF
Application.ScreenUpdating = False

FileName1 = "Speed_Limit"                   'Assign FileName1
Sheets(FileName1).Activate                  'Activate working dataset (Speed Limit)

'Find row and column extent of data
row1 = Range("A1").End(xlDown).row
col1 = Range("A1").End(xlToRight).Column

'Sort data by ROUTE ID and then by BEGINNING MILEPOINT
ActiveWorkbook.Worksheets(FileName1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FileName1).Sort.SortFields.Add Key:=Range( _
    "A2:A" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets(FileName1).Sort.SortFields.Add Key:=Range( _
    "B2:B" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FileName1).Sort
    .SetRange Range("A1:" & Col_Letter(col1) & row1)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'Run Fix_Speed_Limit macro
'   (1) Add a column for LABEL, END_MILEPOINT
'   (2) Complete column of data for Label and End Milepoint
'   (3) Correct ROUTE_ID values
'   (4) Corrects the RECEIVED Values
Fix_Speed_Limit

'Run Remove_Routes macro
'   (1) Removes the crashes on non-state route from the Speed Limit data, copies them to Non-State Routes sheet.
'Remove_Routes         'commented out by Camille

'Sort data by ROUTE ID and then by BEGINNING MILEPOINT
ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Add Key:=Range( _
    "A2:A5077"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Add Key:=Range( _
    "D2:D5077"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Speed_Limit").Sort
    .SetRange Range("A1:H5077")
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Unusual speed limits
nrow = 2
Do While Cells(nrow, 1) <> ""
    Cells(nrow, 6) = Trim(CStr(Cells(nrow, 6)))
    
    Cells(nrow, 6) = Val(Cells(nrow, 6).Value)
    Cells(nrow, 6) = CInt(Cells(nrow, 6).Value)
    
    If Len(CStr(Cells(nrow, 6))) > 2 Then
        Rows(nrow).EntireRow.Delete
    ElseIf Cells(nrow, 6) = "" And Cells(nrow - 1, 3) = Cells(nrow, 3) Then
        Cells(nrow, 6) = Cells(nrow - 1, 6)
        nrow = nrow + 1
    ElseIf Cells(nrow, 6) = "" And Cells(nrow + 1, 3) = Cells(nrow, 3) Then
        Cells(nrow, 6) = Cells(nrow + 1, 6)
        nrow = nrow + 1
    ElseIf Cells(nrow, 6) = "" Or Cells(nrow, 6) = 0 Then
        Cells(nrow, 6) = 25
        nrow = nrow + 1
    Else
        nrow = nrow + 1
    End If
       
    Cells(nrow, 1) = Format(Cells(nrow, 1), "0000")
Loop

'Autofit Columns
Worksheets("Speed_Limit").Columns("A:H").autofit

'Tab color change
Worksheets("Speed_Limit").Tab.ColorIndex = 10

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False
    
End Sub

Sub Fix_Speed_Limit()
'The purpose of this macro is to:
'   (1) Add a column for LABEL, END_MILEPOINT
'   (2) Complete column of data for Label and End Milepoint
'   (3) Correct ROUTE_ID values
'   (4) Corrects the RECEIVED Values
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim Route As String         'Variable for ROUTE_ID
Dim nrow As Long
Dim Direction As String
Dim nstart As Integer
Dim idir As Integer
Dim irou As Integer
Dim BMPcol As Integer

'Identify route ID column
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop

idir = 1
Do Until Cells(1, idir) = "DIRECTION"
    idir = idir + 1
Loop

'Add column for Label
Columns(idir + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir + 1) = "LABEL"

'Correct route values
nrow = 2
Do While Cells(nrow, irou) <> ""
    
    Route = CStr(Cells(nrow, irou))
    
    If Len(Route) >= 5 Then
        Route = Left(Route, 4)
    End If
    
    If Route = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
        Route = "0011"          'Change route to 0011 for the purpose of this process.
    End If
        
    If Len(Route) = 1 Then
        Route = "000" & Route
    ElseIf Len(Route) = 2 Then
        Route = "00" & Route
    ElseIf Len(Route) = 3 Then
        Route = "0" & Route
    Else
    End If
    
    Cells(nrow, irou).NumberFormat = "@"
    Cells(nrow, irou) = Route
    
    nrow = nrow + 1
Loop
Application.CutCopyMode = False

'Correct direction values. Change positive values to "P" and delete negative routes.
nrow = 2
Do While Cells(nrow, irou) <> ""
    If Cells(nrow, idir) = "+" Then
        Cells(nrow, idir) = "P"
        nrow = nrow + 1
    ElseIf Cells(nrow, idir) = "-" Then
        Rows(nrow).EntireRow.Delete
    ElseIf Cells(nrow, idir) = "X" Then
        Rows(nrow).EntireRow.Delete
    ElseIf Cells(nrow, idir) = "N" Then
        Rows(nrow).EntireRow.Delete
    ElseIf Cells(nrow, idir) = "P" Then
        nrow = nrow + 1
    ElseIf Cells(nrow, idir) = "" Then
        Cells(nrow, idir) = "P"
    End If
Loop

'Add a line of data for Interstates and Mountain View Corridor, since they can be identified in the positive or negative direction
nrow = 2
Do While Cells(nrow, irou) <> ""
Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)

    If (Route = "0015" Or Route = "0070" Or Route = "0080" Or _
    Route = "0084" Or Route = "0215" Or Route = "0085") Then
        
        Rows(nrow).Copy
        Rows(nrow + 1).Insert shift:=xlDown

        Cells(nrow + 1, idir) = "N"
        nrow = nrow + 2

    Else
        nrow = nrow + 1
    End If
Loop
Application.CutCopyMode = False

'Sort by route ID and MP
RouteIDSort

'Create the LABEL column data by joining the Route_ID and Direction
nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)
    Cells(nrow, idir + 1) = Route & Direction
    nrow = nrow + 1
Loop
Application.CutCopyMode = False

'corrects the first BMP and last EMP (just for the ISAM - 08/09) for each route (by Camille)
BMPcol = 1
Do While Cells(1, BMPcol) <> "BEG_MILEPOINT"
    BMPcol = BMPcol + 1
Loop
nrow = 2
Do Until Cells(nrow, idir + 1) = ""
    If Cells(nrow + 1, idir + 1) <> Cells(nrow, idir + 1) Then
        Cells(nrow + 1, BMPcol).Value = 0
        If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
            Cells(nrow, BMPcol + 1).Value = Cells(nrow, BMPcol + 1).Value + 1
        End If
    End If
    nrow = nrow + 1
Loop


End Sub


Sub Remove_Routes()              'references to this sub have been commented out by Camille
'Remove_Routes macro:
'   (1) Removes the crashes on non-state route from the Speed Limit data, copies them to Non-State Routes sheet.
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim crow1 As Long            'Row counter for main AADT sheet
Dim prow1 As Long            'Row counter for Non-Routes sheet
Dim row1 As Long            'Row counter 2 for Non-Routes sheet
Dim Route As Long           'Variable for route

'Set inital row values
crow1 = 2
prow1 = 1

'Find first empty row in "Non-State Routes" sheet and add title
Sheets("Non-State Routes").Activate
Do While Cells(prow1, 1) <> ""
    If Cells(prow1, 1) = "Speed_Limit" Then                     'If there is already Speed Limit data, delete it.
        row1 = prow1 + 1
        row1 = Cells(row1, 2).End(xlDown).row
        Rows(prow1 & ":" & row1).EntireRow.Delete
    Else
        prow1 = prow1 + 1
    End If
Loop

'Enter dataset title in first cell to separate these non-state route data from the rest
Cells(prow1, 1) = "Speed_Limit"
prow1 = prow1 + 1
'Copy dataset column headers
Sheets("Speed_Limit").Rows(1).Copy Destination:=Sheets("Non-State Routes").Range("A" & prow1)
prow1 = prow1 + 1

'Steps through each line and checks if the route is greater than 491. Routes above 491 are non-state routes or are ramps.
Sheets("Speed_Limit").Activate
Do Until Cells(crow1, 1) = ""
    Route = CInt(Cells(crow1, 1))
    If (Route > 491) Then
        Rows(crow1).Cut
        ActiveSheet.Paste Destination:=Sheets("Non-State Routes").Range("A" & prow1)
        Rows(crow1).Delete
        Application.CutCopyMode = False
        prow1 = prow1 + 1
    Else
        crow1 = crow1 + 1
    End If
Loop

End Sub

Sub AssignInfo()
'AssignInfo macro:
'   (1) Go through the sheets in the workbook to find the roadway data sheets.
'   (2) Assign sheet names to public variables that can be refered to when working with the datasets.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht As Integer                        'Integer representing worksheet number
Dim WS_Count As Integer                     'Total number of worksheets in the workbook

'Set total worksheet amount to WS_Count
WS_Count = ActiveWorkbook.Worksheets.count

'Cycle through all worksheets. If sheet name correlates to roadway data, assign filename values.
For wksht = 1 To WS_Count
    If InStr(1, Sheets(wksht).Name, "AADT_") > 0 Then
        FN1 = Sheets(wksht).Name
    ElseIf InStr(1, Sheets(wksht).Name, "Functional_") > 0 Then
        FN2 = Sheets(wksht).Name
    ElseIf InStr(1, Sheets(wksht).Name, "Speed_") > 0 Then
        FN3 = Sheets(wksht).Name
    ElseIf InStr(1, Sheets(wksht).Name, "Thru_") > 0 Then
        FN4 = Sheets(wksht).Name
    ElseIf InStr(1, Sheets(wksht).Name, "Urban_") > 0 Then
        FN5 = Sheets(wksht).Name
    End If
Next wksht

End Sub

Public Function Col_Letter(lngCol) As String

Dim vArr

vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)

End Function

