Attribute VB_Name = "DP_05_ThruLanes"
Option Explicit

'PMDP_05_ThruLanes Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets

Sub Run_ThruLanes()
'Call Run_ThruLanes macro:
'   (1) Formats the Thru Lanes data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define Variables
Dim col1, col2, row1, row2 As Integer       'Row and column counters
Dim RCol(1 To 4) As String                  'array used to assign order of columns
Dim HeaderFound As Boolean                  'True or False for deleting headers

'Set inital column value
col1 = 1

'Screen Updating OFF
Application.ScreenUpdating = False

'AssignInfo macro:
'   (1) Go through the sheets in the workbook to find the roadway data sheets.
'   (2) Assign sheet names to public variables that can be refered to when working with the datasets.
AssignInfo

FileName1 = FN4                             'Assign FileName1
Sheets(FileName1).Activate                  'Activate working dataset (Thru Lanes)

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


'Run Fix_Thru_Lanes macro
'   (1) Add a column for LABEL
'   (2) Complete column of data for Label and End Milepoint
Fix_Thru_Lanes

'Run Remove_Routes macro
'   (1) Removes the crashes on non-state route and ramps from the AADT data, copies them to "Non-State Routes" sheet.
'Remove_Routes          'commented out by Camille

'Autofit Columns
Columns("A:F").EntireColumn.autofit

'Sort by route ID and MP
RouteIDSort

'Tab color change
Worksheets("Thru_Lanes").Tab.ColorIndex = 10

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub Fix_Thru_Lanes()
'The purpose of this macro is to:
'   (1) Add a column for LABEL
'   (2) Complete column of data for Label and End Milepoint
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim Route As String         'Variable for ROUTE_ID
Dim nrow1, nrow2 As Long    'Row Counter
Dim totrow, totcol As Long  'Total number of rows and columns
Dim Direction As String     'Variable for DIRECTION
Dim idir As Integer         'Column counter, for finding the column containing DIRECTION data
Dim irou As Integer         'Column counter, for finding the column with the ROUTE_ID data
Dim ibegmp As Integer
Dim iendmp As Integer
Dim inumlanes As Integer
Dim EndMP As Double

Application.Calculation = xlCalculationManual

'Find where columns are
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop
ibegmp = 1
Do Until Cells(1, ibegmp) = "BEG_MILEPOINT"
    ibegmp = ibegmp + 1
Loop
iendmp = 1
Do Until Cells(1, iendmp) = "END_MILEPOINT"
    iendmp = iendmp + 1
Loop
inumlanes = 1
Do Until Cells(1, inumlanes) = "NUM_LANES"
    inumlanes = inumlanes + 1
Loop
Application.CutCopyMode = False


Do While Cells(totrow + 1, 1) <> ""
    totrow = totrow + 1
Loop

Do While Cells(1, totcol + 1) <> ""
    totcol = totcol + 1
Loop

'Consolidate lanes data
nrow1 = 2
nrow2 = 3
Do Until Cells(nrow1, 1) = ""
    If Mid(Cells(nrow1, irou), 5, 1) = "N" Then
        'Range("A" & nrow1 & ":" & ConvertToLetter(totcol) & nrow1).Delete Shift:=xlUp
        Cells(nrow1, 1).EntireRow.Delete
        nrow1 = nrow1 - 1
        nrow2 = nrow1 - 1
    ElseIf Cells(nrow1, irou) = Cells(nrow2, irou) And Cells(nrow1, inumlanes) = Cells(nrow2, inumlanes) Then
        Do Until Cells(nrow1, irou) <> Cells(nrow2 + 1, irou) Or Cells(nrow1, inumlanes) <> Cells(nrow2 + 1, inumlanes)
            nrow2 = nrow2 + 1
        Loop
        Cells(nrow1, iendmp).Value = Cells(nrow2, iendmp).Value
        'Range("A" & nrow1 + 1 & ":" & ConvertToLetter(totcol) & nrow2).Delete Shift:=xlUp
        Range(Cells(nrow1 + 1, 1), Cells(nrow2, 1)).EntireRow.Delete
        'nrow1 = nrow1 + 1
        'nrow2 = nrow1 + 1
    End If
    nrow1 = nrow1 + 1
    nrow2 = nrow1 + 1
Loop


ActiveSheet.UsedRange

'Add columns for Direction, and Label
Columns(irou + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, irou + 1) = "DIRECTION"
Application.CutCopyMode = False
idir = irou + 1
Columns(idir + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir + 1) = "LABEL"
Application.CutCopyMode = False

'Add direction info
nrow1 = 2
Do While Cells(nrow1, irou) <> ""
    Cells(nrow1, idir).Value = "P"
    nrow1 = nrow1 + 1
Loop

'Correct route values
nrow1 = 2
Do While Cells(nrow1, irou) <> ""
    Route = CStr(Cells(nrow1, irou))
    
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
    
    Cells(nrow1, irou).NumberFormat = "@"
    Cells(nrow1, irou) = Route
    
    nrow1 = nrow1 + 1
Loop
Application.CutCopyMode = False

'corrects the first BMP and the last EMP (just for the ISAM - 08/08) for each route
'added by Camille on June 4, 2019; edited on July 29, 2019
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop
ibegmp = 1
Do Until Cells(1, ibegmp) = "BEG_MILEPOINT"
    ibegmp = ibegmp + 1
Loop
iendmp = 1
Do Until Cells(1, iendmp) = "END_MILEPOINT"
    iendmp = iendmp + 1
Loop
nrow1 = 2
Do Until Cells(nrow1, irou) = ""
    If Cells(nrow1 + 1, irou) <> Cells(nrow1, irou) Then
        Cells(nrow1 + 1, ibegmp).Value = 0
        If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
            Cells(nrow1, iendmp).Value = Cells(nrow1, iendmp).Value + 1
        End If
    End If
    nrow1 = nrow1 + 1
Loop


'Add a line of data for Interstates and Mountain View Corridor, since they can be identified in the positive or negative direction
nrow1 = 2
Do While Cells(nrow1, irou) <> ""
    Route = Cells(nrow1, irou)
    Direction = Cells(nrow1, idir)

    If (Route = "0015" Or Route = "0070" Or Route = "0080" Or _
    Route = "0084" Or Route = "0215" Or Route = "0085") Then
        
        Rows(nrow1).Copy
        Rows(nrow1 + 1).Insert shift:=xlDown

        Cells(nrow1 + 1, idir) = "N"
        nrow1 = nrow1 + 2

    Else
        nrow1 = nrow1 + 1
    End If
Loop
Application.CutCopyMode = False

'Create the LABEL column by joining the Route_ID and Direction
nrow1 = 2
Do While Cells(nrow1, irou) <> ""
    Route = Cells(nrow1, irou)
    Direction = Cells(nrow1, idir)
    Cells(nrow1, idir + 1) = Route & Direction
    nrow1 = nrow1 + 1
Loop
Application.CutCopyMode = False

Application.Calculation = xlCalculationAutomatic   'this half of the pair added by Camille on August 6, 2018

End Sub

Sub Remove_Routes()                  'references to this sub have been commented out by Camille
'Remove_Routes macro:
'   (1) Removes the crashes on non-state route and ramps from the AADT data, copies them to "Non-State Routes" sheet.
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim crow As Long            'Row counter for main AADT sheet
Dim prow As Long            'Row counter for Non-Routes sheet
Dim row1 As Long            'Row counter 2 for Non-Routes sheet
Dim Route As Long           'Variable for route

'Assign initial row value
prow = 1

'Find first empty row in "Non-State Routes" sheet and add title
Sheets("Non-State Routes").Activate
Do While Cells(prow, 1) <> ""
    If Cells(prow, 1) = "Thru_Lanes" Then                     'If there is already AADT data, delete it.
        row1 = prow + 1
        row1 = Cells(row1, 2).End(xlDown).row
        Rows(prow & ":" & row1).EntireRow.Delete
    Else
        prow = prow + 1
    End If
Loop

'Enter dataset title in first cell to separate these non-state route data from the rest
Cells(prow, 1) = "Thru_Lanes"
prow = prow + 1
'Copy dataset column headers
Sheets("Thru_Lanes").Rows(1).Copy Destination:=Sheets("Non-State Routes").Range("A" & prow)
prow = prow + 1

crow = 2
'Steps through each line and checks if the route is greater than 491. Routes above 491 are non-state routes or are ramps.
Sheets("Thru_Lanes").Activate
Do Until Cells(crow, 1) = ""
    Route = CInt(Cells(crow, 1))
    If (Route > 491) Then
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

Sub RemoveBlankRows(nrow As Long, datasheet As String)
'MoveBlankRowsToBottom macro:
'   (1) All rows that have been cleared of data are moved to bottom of sheet by a sorting method.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Screen Updating OFF
Sheets(datasheet).Activate

'Declare variable
Dim irou As Integer
    
'Find ROUTE_ID column
irou = 1
Do Until Sheets(datasheet).Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop

'Select all cells, sort by ROUTE_ID to move blank rows to bottom
Cells.AutoFilter
ActiveWorkbook.Worksheets(datasheet).AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets(datasheet).AutoFilter.Sort.SortFields.Add Key:= _
    Range(Cells(2, irou), Cells(nrow, irou)), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
    
With ActiveWorkbook.Worksheets(datasheet).AutoFilter.Sort
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Cells.AutoFilter

End Sub

Sub DeleteRangeMacro()

Dim DelRows() As Variant

Dim DelStr As String, LenStr As Long
Dim CutHere_Str As String
Dim i As Long
Dim nrow As Long

ReDim DelRows(1 To totrow)
nrow1 = 2
nrow2 = 3

For i = 1 To totrow
    If Cells(nrow1, irou) = Cells(nrow2, irou) And Cells(nrow1, inumlanes) = Cells(nrow2, inumlanes) Then
        'Cells(nrow1, iendmp) = Cells(nrow2, iendmp)
        DelRows(i) = nrow2
        nrow2 = nrow2 + 1
    Else
        nrow1 = nrow2
        i = i - 1
    End If
    nrow2 = nrow2 + 1
Next i

'--- How to delete them all together?

LenStr = 0
DelStr = ""

For i = LBound(DelRows) To UBound(DelRows)
    LenStr = LenStr + Len(DelRows(i)) * 2 + 2

    ' One for a comma, one for the colon and the rest for the row number
    ' The goal is to create a string like
    ' DelStr = "15:15,18:18,21:21"

    If LenStr > 200 Then
        LenStr = 0
        CutHere_Str = "!"       ' Demarcator for long strings
    Else
        CutHere_Str = ""
    End If

    DelRows(i) = DelRows(i) & ":" & DelRows(i) & CutHere_Str
Next i

DelStr = Join(DelRows, ",")

Dim DelStr_Cut() As String
DelStr_Cut = Split(DelStr, "!,")
' Each DelStr_Cut(#) string has a usable string

Dim DeleteRng As Range
Set DeleteRng = ActiveSheet.Range(DelStr_Cut(0))

For i = LBound(DelStr_Cut) + 1 To UBound(DelStr_Cut)
    Set DeleteRng = Union(DeleteRng, ActiveSheet.Range(DelStr_Cut(i)))
Next i

DeleteRng.Delete
End Sub

Public Function Col_Letter(lngCol) As String

Dim vArr

vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)

End Function

