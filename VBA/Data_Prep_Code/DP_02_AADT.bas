Attribute VB_Name = "DP_02_AADT"
Option Explicit

'PMDP_02_AADT Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets

Sub Run_AADT()
'Run_AADT Macro:
'   (1) Formats the AADT data in preparation for the combination/segmentation process.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define Variables
Dim col1, col2, row1, row2 As Integer       'column and row counters
Dim AADT As String                          'AADT header variable
Dim HeaderFound As Boolean                  'True or False for deleting headers
Dim myrow As Long                           'row counter

'Screen Updating OFF
Application.ScreenUpdating = False

'Assign chosen Headers to Header variables
AADT = Left(Worksheets("OtherData").Range("A8"), InStr(1, Worksheets("OtherData").Range("A8"), "[") - 1)

'Run AssignInfo macro:
' (1) Go through the sheets in the workbook to find the roadway data sheets.
' (2) Assign sheet names to public variables that can be refered to when working with the datasets.
AssignInfo
FileName1 = "AADT"                             'Assign FileName1
Sheets(FileName1).Activate                  'Activate working dataset (AADT)

'Find row extent of data
row1 = Range("A1").End(xlDown).row
col1 = Range("A1").End(xlToRight).Column

'change route 194 to route 85 (will be changed back at the end of the roadway process)
If Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    myrow = 2
    Do Until Sheets(FileName1).Cells(myrow, 1) = ""
        If Sheets(FileName1).Cells(myrow, 1) = 194 Then
            Sheets(FileName1).Cells(myrow, 1) = 85
        End If
        myrow = myrow + 1
    Loop
End If


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

'Run Fix_Columns macro
' (1) Fix the Route_ID values
' (2) Correct the direction values
' (3) Add a copy of the interstate data, list "X" direction
Fix_Columns

'Run Remove_Routes macro
'   (1) Removes the crashes on non-state routes from the AADT data and places them in the "Non-State Routes" list
'       (Route is a non-state route when route number > 491)
Remove_Routes 'Check to see if this is for the ISAM

'Run Fix_AADT macro:
'   (1) Calculates the total percent truck for AADT and Total Count of Trucks.
Fix_AADT

'Autofit Columns and select cell A1
Columns("A:T").EntireColumn.autofit

'Tab color change
Worksheets("AADT").Tab.ColorIndex = 10

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False
    
End Sub

Sub Fix_Columns()
'Fix_Columns macro:
'   (1) Fix the Route_ID values
'   (2) Correct the direction values
'   (3) Add a copy of the interstate data, list "X" direction
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim Route As String                 'Variable for ROUTE_ID
Dim nrow, nrow2, ncol As Double     'Row and column counters
Dim Direction As String             'Variable for DIRECTION
Dim idir As Integer                 'Column counter, for finding the column with the DIRECTION data
Dim irou As Integer                 'Column counter, for finding the column with the ROUTE_ID data
Dim icou As Integer                 'Column counter, for finding the column with the COUNTY data
Dim iAADT As Integer                'Column counter, for finding the column with the first year of AADT data
Dim CountyCode As Long              'County number found in STATION value
Dim County As String                'County name
Dim RegionCode As Integer           'Region number
Dim ibegmp As Integer               'column for beginning MP
Dim iendmp As Integer               'column for ending MP

'Identify the Column with the correct information
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop
idir = irou + 1

'Create a direction column and assign value to direction column variable
Columns(idir).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir) = "DIRECTION"

'identify column
ibegmp = 1
Do Until Cells(1, ibegmp) = "BEG_MILEPOINT"
    ibegmp = ibegmp + 1
Loop
'identify column
iendmp = 1
Do Until Cells(1, iendmp) = "END_MILEPOINT"
    iendmp = iendmp + 1
Loop

'Convert the Route_ID value to a 4 character integer
nrow = 2
Do While Cells(nrow, irou) <> ""
    
    Route = CStr(Cells(nrow, irou))
    
    'If Len(Route) >= 5 Then
        'Rows(nrow).EntireRow.Delete
    'Else
        If Route = "089A" Or (Route = "89" And Cells(nrow, ibegmp) = 0 And Cells(nrow, iendmp) = 2.945) Then     'Corrects strange route ID. 089A was previously called SR-11, so the route is being changed to 11 for the purpose of running this code
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
    'End If
Loop

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
iAADT = 1
Do Until InStr(1, Cells(1, iAADT), "AADT_") > 0
    iAADT = iAADT + 1
Loop

nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)

    If (Route = "0015" Or Route = "0070" Or Route = "0080" Or _
    Route = "0084" Or Route = "0215" Or Route = "0085") Then
        Cells(nrow, iAADT) = Round(Cells(nrow, iAADT) / 2, 0)
        Cells(nrow, iAADT + 1) = Round(Cells(nrow, iAADT + 1) / 2, 0)
        Cells(nrow, iAADT + 2) = Round(Cells(nrow, iAADT + 2) / 2, 0)
        Cells(nrow, iAADT + 3) = Round(Cells(nrow, iAADT + 3) / 2, 0)
        Cells(nrow, iAADT + 4) = Round(Cells(nrow, iAADT + 4) / 2, 0)
        Cells(nrow, iAADT + 5) = Round(Cells(nrow, iAADT + 5) / 2, 0)
        Cells(nrow, iAADT + 6) = Round(Cells(nrow, iAADT + 6) / 2, 0)

        Rows(nrow).Copy
        Rows(nrow + 1).Insert shift:=xlDown
        Application.CutCopyMode = False

        Cells(nrow + 1, idir) = "N"
        nrow = nrow + 2

    Else
        nrow = nrow + 1
    End If
Loop
Application.CutCopyMode = False

'Create the LABEL column by joining the Route_ID and Direction
Columns(idir + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, idir + 1) = "LABEL"
nrow = 2
Do While Cells(nrow, irou) <> ""
    Route = Cells(nrow, irou)
    Direction = Cells(nrow, idir)
    Cells(nrow, idir + 1) = Route & Direction

    nrow = nrow + 1
Loop

'Find COUNTY column and assign column number to icou
icou = 1
Do Until Cells(1, icou) = "COUNTY"
    icou = icou + 1
Loop

'Change STATION values in the COUNTY column to county names.
'This assumes that the county data is not in the data already
nrow = 2
ncol = 48
Do While Cells(nrow, icou) <> ""
    CountyCode = Left(Cells(nrow, icou), 3)
    nrow2 = 4
    Do While Worksheets("OtherData").Cells(nrow2, ncol) <> ""
        If Worksheets("OtherData").Cells(nrow2, ncol) = CountyCode Then
            Cells(nrow, icou) = Worksheets("OtherData").Cells(nrow2, ncol + 1)
        End If
        nrow2 = nrow2 + 1
    Loop
    nrow = nrow + 1
Loop

'Create a REGION column. Draw region value from the county value using table in OtherData sheet.
Columns(icou + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, icou + 1) = "REGION"
nrow = 2
ncol = 55
Do While Cells(nrow, icou) <> ""
    County = Cells(nrow, icou)
    nrow2 = 4
    Do While Worksheets("OtherData").Cells(nrow2, ncol) <> ""
        If Worksheets("OtherData").Cells(nrow2, ncol) = County Then
            Cells(nrow, icou + 1) = Worksheets("OtherData").Cells(nrow2, ncol + 1)
        End If
        nrow2 = nrow2 + 1
    Loop
    nrow = nrow + 1
Loop

'Delete unwanted AADT year columns
iAADT = 1
Do Until InStr(1, Cells(1, iAADT), "AADT_") > 0
    iAADT = iAADT + 1
Loop

ncol = iAADT

If Sheets("Inputs").Cells(2, 16) = "RSAM" Then
    ncol = ncol + 7
ElseIf Sheets("Inputs").Cells(2, 16) = "CAMS" Then  'added by Camille on 08/23/2019
    Do
        ncol = ncol + 1
    Loop Until Int(Right(Cells(1, ncol), 4)) < 2010
End If

Do Until InStr(1, Cells(1, ncol), "AADT_") < 1
    Columns(ncol).EntireColumn.Delete
Loop

End Sub

Sub Remove_Routes()
'Remove_Routes macro:
'   (1) Removes the crashes on non-state routes from the AADT data and places them in the "Non-State Routes" list
'       (Route is a non-state route when route number > 491)
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim crow As Long            'Row counter for main AADT sheet
Dim prow As Long            'Row counter for Non-Routes sheet
Dim prow2 As Long           'Row counter 2 for Non-Routes sheet
Dim Route As Long           'Variable for route

'Set inital row values
crow = 2
prow = 1

'Find first empty row in "Non-State Routes" sheet and add title
Sheets("Non-State Routes").Activate
Do While Cells(prow, 1) <> ""
    If Cells(prow, 1) = "AADT" Then                     'If there is already AADT data, delete it.
        prow2 = prow + 1
        Do While Cells(prow2 + 1, 2) <> ""
            prow2 = prow2 + 1
        Loop
        Rows(prow & ":" & prow2).EntireRow.Delete
    Else
        prow = prow + 1
    End If
Loop

'Enter dataset title in first cell to separate these non-state route data from the rest
Cells(prow, 1) = "AADT"
prow = prow + 1
'Copy dataset column headers
Sheets("AADT").Rows(1).Copy Destination:=Sheets("Non-State Routes").Range("A" & prow)
prow = prow + 1

'Steps through each line and checks if the route is greater than 491. Routes above 491 are non-state routes or are ramps.
Sheets("AADT").Activate
Do Until Cells(crow, 1) = ""
    Route = CInt(Cells(crow, 1))
    'If the route number is greater than 491, cut and paste it to Non-State Routes sheet
    If (Route > 491) Then
        Rows(crow).Cut
        ActiveSheet.Paste Destination:=Sheets("Non-State Routes").Range("A" & prow)
        Rows(crow).Delete
        prow = prow + 1
    Else
        crow = crow + 1
    End If
Loop
Application.CutCopyMode = False




End Sub

Sub Fix_AADT()
'Fix_AADT macro:
'   (1) Calculates the total percent truck for AADT and Total Count of Trucks.
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Sam Mineer and Josh Gibbons, BYU, 2015-2016

'Define Variables
Dim n As Double             'Row counter
Dim ising As Integer        'Column counter, to find the column with the "Single_Percent" data
Dim icomb As Integer
Dim iAADT As Integer
Dim singp As Double         'value for Single_Percent
Dim combp As Double         'value for Combo_Percent
Dim singc As Double         'value for Single_Count
Dim combc As Double         'value for Combo_Count

'Find the column with "Single_Percent", which will be used as a reference point
ising = 1
Do Until Cells(1, ising) = "Single_Percent"
    ising = ising + 1
Loop

icomb = 1
Do Until Cells(1, icomb) = "Combo_Percent"
    icomb = icomb + 1
Loop

iAADT = 1
Do Until InStr(1, Cells(1, iAADT), "AADT") > 0
    iAADT = iAADT + 1
Loop

'Creates new column headers
Cells(1, ising + 2) = "Single_Count"
Cells(1, ising + 3) = "Combo_Count"
Cells(1, ising + 4) = "Total_Percent_Trucks"
Cells(1, ising + 5) = "Total_Count_Trucks"

'For each row, calculates the Total Percent Trucks and Total Count Trucks
n = 2
Do Until Cells(n, 1) = ""
    singp = Cells(n, ising)
    combp = Cells(n, icomb)
    singc = Round(Cells(n, iAADT) * Cells(n, ising), 0)
    combc = Round(Cells(n, iAADT) * Cells(n, icomb), 0)
    If combc <> 0 And singc <> 0 Then
        Cells(n, ising + 3) = combc
        Cells(n, ising + 2) = singc
        Cells(n, ising + 4) = singp + combp
        Cells(n, ising + 5) = singc + combc
    End If
    n = n + 1
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
        DYear1 = Right(Sheets(wksht).Name, 4)
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

Function Col_Letter(lngCol) As String

Dim vArr

vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)

End Function

