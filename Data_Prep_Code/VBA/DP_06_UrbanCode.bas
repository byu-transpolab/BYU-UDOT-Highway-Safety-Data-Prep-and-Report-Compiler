Attribute VB_Name = "DP_06_UrbanCode"
Option Explicit

'PMDP_06_UrbanCode Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets
Public CurrentWkbk As String                'String variable that holds current workbook name for reference purposes

Sub Run_UrbanCode()
'Call Run_UrbanCode macro:
'   (1) Formats the Urban Code data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

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

FileName1 = FN5                             'Assign FileName1
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


'Run Fix_Urban_Code macro
'   (1) Add a column for LABEL
'   (2) Complete column of data for LABEL
Fix_Urban_Code

'Run Remove_Routes macro:
'   (1) Removes the crashes on non-state route and ramps from the Urban Code data, copies to "Non-State Routes" sheet.
'Remove_Routes             'commented out by Camille

'Autofit Columns
Columns("A:G").EntireColumn.autofit

'Sort by route ID and MP
RouteIDSort

'Tab color change
Worksheets("Urban_Code").Tab.ColorIndex = 10

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub Fix_Urban_Code()
'The purpose of this macro is to:
'   (1) Add a column for LABEL
'   (2) Complete column of data for LABEL
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

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

'Add columns for Direction, and Label
Columns(irou + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, irou + 1) = "DIRECTION"
idir = irou + 1
Columns(idir + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir + 1) = "LABEL"

'Correct the direction values
'   Add another column of data for the interstate routes
' First, convert the Route_ID value to a 4 character integer
nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = CStr(Cells(nrow, irou))
    
    'If Len(Route) >= 5 Then
        'Rows(nrow).EntireRow.Delete
    'Else
        If nrow = 1881 Then
            nrow = nrow
        End If
        If Route = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Route = "0011"          'Change route to 0011 for the purpose of this process.
            Cells(nrow, irou).NumberFormat = "@"
            Cells(nrow, irou) = Route
        End If
            
        If Len(Route) = 1 Then
            Route = "000" & Route
            Cells(nrow, irou).NumberFormat = "@"
            Cells(nrow, irou) = Route
        ElseIf Len(Route) = 2 Then
            Route = "00" & Route
            Cells(nrow, irou).NumberFormat = "@"
            Cells(nrow, irou) = Route
        ElseIf Len(Route) = 3 Then
            Route = "0" & Route
            Cells(nrow, irou).NumberFormat = "@"
            Cells(nrow, irou) = Route
        ElseIf Len(Route) > 4 Then
            Cells(nrow, irou).EntireRow.Delete
            nrow = nrow - 1   'Camille 07/08/19 accounts for times where there are two rows with len(route)>4 in a row
        End If
        
        
        nrow = nrow + 1
    'End If
Loop
Application.CutCopyMode = False


nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Int(Cells(nrow, irou))
    If Route > 491 Then
        Cells(nrow, irou).EntireRow.Delete
        nrow = nrow - 1   'Camille 07/08/19 accounts for times where there are two rows with len(route)>4 in a row
    End If
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

'Add a line of data for Interstates and Mountian View Corridor, since they can be identified in the positive or negative direction
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

'Create the LABEL column by joining the Route_ID and Direction
nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)
    Cells(nrow, idir + 1) = Route & Direction
    nrow = nrow + 1
Loop
Application.CutCopyMode = False

End Sub

Sub Remove_Routes()             'references to this sub have been commented out by Camille
'Remove_Routes macro:
'   (1) Removes the crashes on non-state route and ramps from the Urban Code data, copies to "Non-State Routes" sheet.
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim crow As Long            'Row counter for main AADT sheet
Dim prow As Long            'Row counter for Non-Routes sheet
Dim row1 As Long            'Row counter 2 for Non-Routes sheet
Dim Route As Long           'Variable for route

'Set row value
prow = 1

'Find first empty row in "Non-State Routes" sheet and add title
Sheets("Non-State Routes").Activate
Do While Cells(prow, 1) <> ""
    If Cells(prow, 1) = "Urban_Code" Then                     'If there is already AADT data, delete it.
        row1 = prow + 1
        row1 = Cells(row1, 2).End(xlDown).row
        Rows(prow & ":" & row1).EntireRow.Delete
    Else
        prow = prow + 1
    End If
Loop

'Enter dataset title in first cell to separate these non-state route data from the rest
Cells(prow, 1) = "Urban_Code"
prow = prow + 1
'Copy dataset column headers
Sheets("Urban_Code").Rows(1).Copy Destination:=Sheets("Non-State Routes").Range("A" & prow)
prow = prow + 1

crow = 2
'Steps through each line and checks if the route is greater than 491. Routes above 491 are non-state routes or are ramps.
Sheets("Urban_Code").Activate
Do Until Cells(crow, 1) = ""
    Route = CInt(Cells(crow, 1))
    If (Route > 491) Then
        Rows(crow).Cut
        ActiveSheet.Paste Destination:=Sheets("Non-State Routes").Range("A" & prow)
        Rows(crow).Delete
        Application.CutCopyMode = False
        prow = prow + 1
    ElseIf Len(Str(Route)) > 4 Then
        Rows(crow).Cut
        ActiveSheet.Paste Destination:=Sheets("Non-State Routes").Range("A" & prow)
        Rows(crow).Delete
        prow = prow + 1
    Else
        crow = crow + 1
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

