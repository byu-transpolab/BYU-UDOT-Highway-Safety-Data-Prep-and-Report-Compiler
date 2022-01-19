Attribute VB_Name = "DP_03_FunctionalClass"
Option Explicit

'PMDP_03_FunctionalClass Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets

Sub Run_FunctionalClass()
'Call Run_FunctionalClass macro:
'   (1) Formats the Functional Class data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define Variables
Dim col1, col2, row1, row2 As Integer       'Column and row counters
Dim HeaderFound As Boolean                  'True or False for deleting headers
Dim myrow As Long                           'row counter

'Assign beginning column to 1
col1 = 1

'Don't refresh screen
Application.ScreenUpdating = False

'Run AssignInfo macro:
'   (1) Go through the sheets in the workbook to find the roadway data sheets.
'   (2) Assign sheet names to public variables that can be refered to when working with the datasets.
AssignInfo

FileName1 = FN2                             'Assign FileName1
Sheets(FileName1).Activate                  'Activate working dataset (Functional Class)

'change route 194 to route 85 (will be changed back at the end of the roadway process)        'FLAGGED: routes 194 and 85 are still causing problems. I'm not sure how much of a role this plays
If Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    myrow = 2
    Do Until Sheets(FileName1).Cells(myrow, 1) = ""
        If Left(Sheets(FileName1).Cells(myrow, 1), 4) = "0194" Then
            Sheets(FileName1).Cells(myrow, 1) = Replace(Sheets(FileName1).Cells(myrow, 1), "0194", "0085")
        End If
        myrow = myrow + 1
    Loop
End If

'''''''this code is repeated in the Fix_FC sub
''Sort data by ROUTE ID and then by BEGINNING MILEPOINT
'ActiveWorkbook.Worksheets(FileName1).Sort.SortFields.Clear
'ActiveWorkbook.Worksheets(FileName1).Sort.SortFields.Add Key:=Range( _
'    "A2:A" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'    xlSortNormal
'ActiveWorkbook.Worksheets(FileName1).Sort.SortFields.Add Key:=Range( _
'    "B2:B" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'    xlSortNormal
'With ActiveWorkbook.Worksheets(FileName1).Sort
'    .SetRange Range("A1:" & Col_Letter(col1) & row1)
'    .Header = xlYes
'    .MatchCase = False
'    .Orientation = xlTopToBottom
'    .SortMethod = xlPinYin
'    .Apply
'End With


'Run Fix_FC Macro
'   (1) Add a column for Route Name, Direction and Label
'   (2) Translate the Functional Class Code value to a description
Fix_FC

'Run Remove_Routes Macro
'   (1) Removes the crashes on non-state route from the Functional Class data, copies them to Non-State Routes sheet.
'Remove_Routes         'commented out by Camille

'Autofit Columns and delete extra rows
Columns("A:G").EntireColumn.autofit

'Tab color change
Worksheets("Functional_Class").Tab.ColorIndex = 10

'Sort by route ID and MP
RouteIDSort

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub


Sub Fix_FC()
'Fix_FC macro:
'   (1) Add a column for Route Name, Direction and Label
'   (2) Translate the Functional Class Code value to a description
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim Route As String                 'Variable for ROUTE_ID
Dim nrow, nrow2, ncol As Long       'Row and Column Counters
Dim Direction As String             'Variable for Direction
Dim idir As Integer                 'Column counter for finding column with DIRECTION data
Dim irou As Integer                 'Column counter for finding column with ROUTE_IT data
Dim ifc As Integer                  'Column counter for finding column with FC_CODE data
Dim fc As Long                      'Functional Class number value
Dim found As Boolean                'True or False of certain value was found
Dim myrow As Integer                'another row counter
Dim row1 As Integer
Dim col1 As Integer


'Add two columns for DIRECTION and LABEL
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop
idir = irou + 1
Columns(idir).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir) = "DIRECTION"
Columns(idir + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir + 1) = "LABEL"

'A lot here has been changed. Camille: we are given the label and don't need the label or direction. Converts label to route.
''''''''''''''''''''''''''''''''''Actually, we totally need the label for the CAMS combine roadway button process

'delete the rows with N or B
irou = 1
nrow = 2
Do Until Cells(nrow, irou) = ""
    If InStr(1, Cells(nrow, irou), "N") Or InStr(1, Cells(nrow, irou), "B") Then
        Cells(nrow, irou).EntireRow.Delete
        nrow = nrow - 1
    End If
    nrow = nrow + 1
Loop

'Correct route values
nrow = 2
Do While Cells(nrow, irou) <> ""
    
    Route = CStr(Cells(nrow, irou))
    Route = Left(Route, 4)
    
    If Route = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
        Route = "0011"          'Change route to 0011 for the purpose of this process.
    End If
          
    If Len(Route) = 1 Then      'add leading zeros
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

'Fill in the direction column with default all P
nrow = 2
Do Until Cells(nrow, irou) = ""
    Cells(nrow, idir) = "P"
    nrow = nrow + 1
Loop

'Create the LABEL column by joining the Route_ID and Direction
nrow = 2
Do While Cells(nrow, irou) <> ""

    Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)
    Cells(nrow, idir + 1) = Route & Direction

    nrow = nrow + 1
Loop

'Sort data by ROUTE ID and then by BEGINNING MILEPOINT
row1 = Range("A1").End(xlDown).row
col1 = Range("A1").End(xlToRight).Column

ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:=Range( _
    Cells(2, irou), Cells(row1, irou)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:=Range( _
    Cells(2, idir + 2), Cells(row1, idir + 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN2).Sort
    .SetRange Range(Cells(1, 1), Cells(row1, col1))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

ifc = 1
Do Until Cells(1, ifc) = "FC_CODE"
    ifc = ifc + 1
Loop

'condenses the rows
nrow = 2
Do Until Cells(nrow, 1) = ""
If Cells(nrow + 1, 1) = Cells(nrow, 1) And Cells(nrow + 1, ifc) = Cells(nrow, ifc) Then
    Cells(nrow, idir + 3) = Cells(nrow + 1, idir + 3)
    Cells(nrow + 1, 1).EntireRow.Delete
    nrow = nrow - 1
End If
nrow = nrow + 1
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
        Cells(nrow + 1, idir + 1) = Cells(nrow + 1, irou) & Cells(nrow + 1, idir)
        nrow = nrow + 2

    Else
        nrow = nrow + 1
    End If
Loop
Application.CutCopyMode = False

'corrects the first BMP and last EMP (just for the ISAM - 08/08) for each route (by Camille)
nrow = 2
Do Until Cells(nrow, idir + 1) = ""
    If Cells(nrow + 1, idir + 1) <> Cells(nrow, idir + 1) Then
        Cells(nrow + 1, idir + 2).Value = 0
        If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
            Cells(nrow, idir + 3).Value = Cells(nrow, idir + 3).Value + 1
        End If
    End If
    nrow = nrow + 1
Loop

''Correct direction values. Change positive values to "P" and delete negative routes.
'nrow = 2
'Do While Cells(nrow, irou) <> ""
'    If Cells(nrow, idir) = "+" Then
'        Cells(nrow, idir) = "P"
'        nrow = nrow + 1
'    ElseIf Cells(nrow, idir) = "-" Then
'        Rows(nrow).EntireRow.Delete
'    ElseIf Cells(nrow, idir) = "X" Then
'        Rows(nrow).EntireRow.Delete
'    ElseIf Cells(nrow, idir) = "N" Then
'        Rows(nrow).EntireRow.Delete
'    ElseIf Cells(nrow, idir) = "P" Then
'        nrow = nrow + 1
'    ElseIf Cells(nrow, idir) = "" Then
'        Cells(nrow, idir) = "P"
'    End If
'Loop

''Add FC_Type column
'not needed because the update in FC files provides the text value in a separate column
'Cells(1, ifc + 1) = "FC_Type"
''Classifies the roadway by functional class with description
'nrow = 2
'ncol = 57
'Do While Cells(nrow, ifc) <> ""
'    found = False
'    fc = Cells(nrow, ifc)
'    nrow2 = 4
'    Do While Worksheets("OtherData").Cells(nrow2, ncol) <> ""
'        If Worksheets("OtherData").Cells(nrow2, ncol) = fc Then
'            Cells(nrow, ifc + 1) = Worksheets("OtherData").Cells(nrow2, ncol + 1)
'        End If
'        nrow2 = nrow2 + 1
'    Loop
'    nrow = nrow + 1
'Loop
Application.CutCopyMode = False

End Sub

Sub Remove_Routes()               'references to this sub in the ISAM have been commented out by Camille
'Remove_Routes macro:
'   (1) Removes the crashes on non-state route from the Functional Class data, copies them to Non-State Routes sheet.
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim crow As Long            'Row counter for main AADT sheet
Dim prow As Long            'Row counter for Non-Routes sheet
Dim row1 As Long            'Row counter 2 for Non-Routes sheet
Dim Route As Long           'Variable for route
Dim irou As Integer         'Column counter for column with Route_ID
Dim ws As Worksheet         '
'Assign inital row value
prow = 1

Set ws = ThisWorkbook.Sheets.Add(After:= _
ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "Non-State Routes"

'Find first empty row in "Non-State Routes" sheet and add title
Sheets("Non-State Routes").Activate
Do While Cells(prow, 1) <> ""
    If Cells(prow, 1) = "Functional_Class" Then                     'If there is already FClass data, delete it.
        row1 = prow + 1
        row1 = Cells(row1, 2).End(xlDown).row
        Rows(prow & ":" & row1).EntireRow.Delete
    Else
        prow = prow + 1
    End If
Loop

'Enter dataset title in first cell to separate these non-state route data from the rest
Cells(prow, 1) = "Functional_Class"
prow = prow + 1
'Copy dataset column headers
Sheets("Functional_Class").Rows(1).Copy Destination:=Sheets("Non-State Routes").Range("A" & prow)
prow = prow + 1

'Identifies column with Route_ID
Sheets("Functional_Class").Activate
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop

'Steps through each line and checks if the route is greater than 491. Routes above 491 are non-state routes or are ramps.
crow = 2
prow = 1
Do Until Cells(crow, irou) = ""
    Route = CInt(Cells(crow, irou))
    If (Route > 491) Then
        Rows(crow).Cut
        ActiveSheet.Paste Destination:=Sheets("Non-State Routes").Range("A" & prow)
        Rows(crow).Delete
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

