Attribute VB_Name = "DP_11_PavementMessage"
Option Explicit

'DP_11_PavementMessage Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets
Public CurrentWkbk As String                'String variable that holds current workbook name for reference purposes

Sub Run_PavementMessages()
'Call Run_PavementMessage macro:
'   (1) Formats the Pavement Message data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2017
' Commented by: Josh Gibbons, BYU, 2017

'Define Variables
Dim row1 As Long                     'Row counter
Dim col1 As Long

'Screen Updating OFF
Application.ScreenUpdating = False

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'AssignInfo macro:
'   (1) Go through the sheets in the workbook to find the roadway data sheets.
'   (2) Assign sheet names to public variables that can be refered to when working with the datasets.
AssignInfo

FileName1 = "Pavement_Messages"                             'Assign FileName1
Sheets(FileName1).Activate                  'Activate working dataset (Urban Code)

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


'Run Fix_PavementMessages macro
'   (1) Add a column for LABEL
'   (2) Complete column of data for LABEL
Fix_PavementMessages

'Autofit Columns and select cell A1
Columns("A:D").EntireColumn.autofit

'Sort by route ID and MP
RouteIDSort

'Tab color change
Worksheets("Pavement_Messages").Tab.ColorIndex = 10

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub Fix_PavementMessages()
'The purpose of this macro is to:
'   (1) Add a column for LABEL
'   (2) Complete column of data for LABEL
'
' Modified by: Josh Gibbons, BYU, 2017

'Define Variables
Dim Route As String         'Variable for ROUTE_ID
Dim nrow As Double          'row counter
Dim Direction As String     'Variable for DIRECTION
Dim idir As Integer         'Column counter, to find the column with the DIRECTION data
Dim irou As Integer         'Column counter, to find the column with the ROUTE_ID data

'Find where ROUTE_ID is
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop

'Correct the direction values
'   Add another column of data for the interstate routes
' First, convert the Route_ID value to a 4 character integer
nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Left(CStr(Cells(nrow, irou)), 4)
    
    Route = Left(CStr(Cells(nrow, irou)), 4)
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



