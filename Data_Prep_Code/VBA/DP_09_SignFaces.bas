Attribute VB_Name = "DP_09_SignFaces"
Option Explicit

'PMDP_11_SignFaces Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets

Sub Run_SignFaces()
'Call Run_SignFaces macro:
'   (1) Formats the Speed Limit data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define variables
Dim col1, col2 As Integer 'column number
Dim nrow As Integer 'nrow number
Dim RCol(1 To 5) As String

'Screen Updating OFF
Application.ScreenUpdating = False

'Run Fix_Speed_Limit macro
Fix_Speed_Limit

'Run Remove_Routes macro
'Remove_Routes      'commented out by Camille

'Sort by route number then by milepoint
ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Add Key:=Range( _
    "A2:A5077"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Add Key:=Range( _
    "D2:D5077"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Speed_Limit").Sort
    .SetRange Range("A1:G5077")
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
    
    If Len(CStr(Cells(nrow, 6))) > 2 Then
        Rows(nrow).EntireRow.Delete
    ElseIf Cells(nrow, 6) = "" Then
        Cells(nrow, 6) = Cells(nrow - 1, 6)
        nrow = nrow + 1
    Else
        nrow = nrow + 1
    End If
    
    Cells(nrow, 6) = Val(Cells(nrow, 6).Value)
    Cells(nrow, 6) = CInt(Cells(nrow, 6).Value)
    
    Cells(nrow, 1) = Format(Cells(nrow, 1), "0000")
Loop

Worksheets("Speed_Limit").Columns("A:H").autofit

'Tab color change
Worksheets("Speed_Limit").Tab.ColorIndex = 10

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False
    
End Sub

Sub Fix_Speed_Limit()
'The purpose of this macro is to:
' (1) Add a column for LABEL, END_MILEPOINT
' (2) Complete column of data for Label and End Milepoint
' (3) Correct ROUTE_ID values
' (4) Corrects the RECEIVED Values
'   Comments by Sam Mineer, Brigham Young University, 2015

'Define Variables
Dim Route As String         'Variable for ROUTE_ID
Dim nrow As Long
Dim Direction As String
Dim nstart As Integer
Dim idir As Integer
Dim irou As Integer

'Activate Speed_Limit sheet
Sheets("Speed_Limit").Activate

'Add columns for End Milepoint, Label and Direction
idir = 2
Columns(4).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns(3).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
irou = 1

Cells(1, 5) = "END_MILEPOINT"
Cells(1, 3) = "LABEL"

'Correct the direction values
'   Add another column of data for the interstate routes
'First, convert the Route_ID value to a 4 character integer
nrow = 2
Do While Cells(nrow, irou) <> ""
    
    Route = CStr(Cells(nrow, irou))
    
    If Len(Route) >= 5 Then
        Route = Left(Route, 4)
    End If
    
    If Route = "089A" Then     'Corrects strange route ID. 089A was previously called SR-11, so the route is being changed to 11 for the purpose of running this code
        Route = "0011"
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

'Remove "-" values and Add replaces "+" with "P"
nrow = 2
Do While Cells(nrow, irou) <> ""
    'Sets the END MP equal to the BEG MP
    Cells(nrow, idir + 3) = Cells(nrow, idir + 2)
    'Trims the date values
    Cells(nrow, idir + 5) = CStr(Left(Cells(nrow, idir + 5), 10))
    
    'Correct direction values. Change positive values to "P" and delete negative routes.
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
Application.CutCopyMode = False

'Add a line of data for Interstates and Mountain View Corridor, since they can be identified in the positive or negative direction
nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)

    If (Route = "0015" Or Route = "0070" Or Route = "0080" Or Route = "0084" Or Route = "0215" Or Route = "0085") Then
        
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

Sub Remove_Routes()              'references to this sub have been commented out by Camille
'The purpose of this macro is to remove the crashes on non state route  from the AADT data to the "Non-State Routes" list
'   Comments by Sam Mineer, Brigham Young University, 2015

'Define Variables
Dim crow1 As Double          'nrow counter for main AADT sheet
Dim prow1 As Double          'nrow counter for Non-Routes sheet
Dim Route As Long           'Variable for route

crow1 = 2
prow1 = 1

'Find first empty row in "Non-State Routes" sheet and add title
Sheets("Non-State Routes").Activate
Do While Cells(prow1, 1) <> ""
    prow1 = prow1 + 1
Loop
Cells(prow1, 1) = "Speed_Limit"
prow1 = prow1 + 1
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
