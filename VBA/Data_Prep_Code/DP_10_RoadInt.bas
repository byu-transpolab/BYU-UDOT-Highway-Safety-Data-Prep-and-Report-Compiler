Attribute VB_Name = "DP_10_RoadInt"
Option Explicit

Sub Run_Intersections()

'Declare Variables
Dim colRoute0, ColSR As Integer
Dim row1, LastRow, lastcol As Long
Dim Route, LastColLetter, sname, guiwb, wd As String
Dim StartTime, EndTime As String

'Activate intersections worksheet
Worksheets("Intersections").Activate
guiwb = ActiveWorkbook.Name
wd = Sheets("Inputs").Range("I2").Value

'This makes the intersection data look like what we need
CleanUpIntFile

'Find columns
colRoute0 = 1
Do Until Cells(1, colRoute0) = "ROUTE"
    colRoute0 = colRoute0 + 1
Loop

'ColSR = 0
'Do Until Cells(1, ColSR) = "SR_SR_INT"
'    ColSR = ColSR + 1
'Loop

'Find intersections dataset extent
LastRow = Range("A1").End(xlDown).row
lastcol = Range("A1").End(xlToRight).Column
LastColLetter = Col_Letter(lastcol)

'''''''''commented out by Camille on May 16, 2019
'Copy state route intersections to dataset sheet
''Cells(1, 1).Select
''Selection.AutoFilter
''ActiveSheet.Range("$A$1:$" & LastColLetter & "$" & LastRow).AutoFilter Field:=ColSR, Criteria1:="YES"
''Range("A1:" & LastColLetter & "1").Select
''Range(Selection, Selection.End(xlDown)).Select
''Selection.Copy
''Sheets.Add(Worksheets("AADT")).Name = "Dataset"
''Sheets("Dataset").Activate
''Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
''Cells(1, 1).Select
''Application.CutCopyMode = False
'Delete intersections sheet without prompt
''Application.DisplayAlerts = False
''Sheets("Intersections").Delete
''Application.DisplayAlerts = True


'Change "intersections" sheet name to "Dataset" (the old code used to copy and paste to this sheet, but we don't need to do that)
Sheets("Intersections").Name = "Dataset"

'Correct route number by removing direction letter
row1 = 2
Do While Cells(row1, colRoute0) <> ""
    Route = CStr(Cells(row1, colRoute0))
    
    If Len(Route) > 5 Then
        Rows(row1).EntireRow.Delete
    Else
        Route = Left(Route, 4)
        
        If Route = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.  FLAGGED: we need to figured out what this is about.
            Route = "0011"          'Change route to 0011 for the purpose of this process.
        End If
        
        Cells(row1, colRoute0).NumberFormat = "@"
        Cells(row1, colRoute0) = Route
        
        row1 = row1 + 1
    End If
Loop

'Activate dataset sheet
Sheets("Dataset").Activate

'Run IntControlIdentifier macro
''IntControlIdentifier   'Commented out by Camille: we don't need this becasue there are no blanks in the TRAFFIC_CO column

''Run DupIntAnalysis macro
'DupIntAnalysis         'Commented out by Camille: we don't need this anymore becuase we have the group id macro. If you uncomment this, see macro sub for altering instructions

'Correct city data
CorrectCity

'Changing Region column from strings to integers
Dim RegCol As Integer, RegRow As Integer
Dim RegString, RegInteger As String
RegCol = 1
RegRow = 2

Do Until Cells(1, RegCol) = "REGION"
    RegCol = RegCol + 1
Loop

Do While Cells(RegRow, RegCol) <> ""
    RegString = Cells(RegRow, RegCol)
    RegInteger = Left(RegString, 1)
    Cells(RegRow, RegCol) = RegInteger
    Cells(RegRow, RegCol).NumberFormat = "0"
    Cells(RegRow, RegCol).Value = Cells(RegRow, RegCol).Value
    RegRow = RegRow + 1
Loop

Sheets("Dataset").UsedRange

'Fill in Intersection ID numbers
Columns(1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, 1) = "INT_ID"
Application.CutCopyMode = False

row1 = 2
Do While Cells(row1, 2) <> ""
    Cells(row1, 1) = row1 - 1
    row1 = row1 + 1
Loop


'Add in FC data
FClassCalc

'Add in Pavement Message data
PavMessCalc

'Add in entering vehicle, region, and county data
EntVehCalc

'Add in lanes data
NumLanesCalc

'Add in speed limit data
SpeedLimitCalc

'Add in urban code data
UrbanCodeCalc

Sheets("Dataset").UsedRange

'Adjust order of region, county, and city columns
ColAdjust

'Fill in empty urban code cells
finishUC

'Set tab color
Worksheets("Dataset").Tab.ColorIndex = 10

'Delete other roadway tabs
Dim wksht As Worksheet
For Each wksht In Worksheets
    If wksht.Name = "AADT" Or wksht.Name = "Functional_Class" Or _
    wksht.Name = "Thru_Lanes" Or wksht.Name = "Speed_Limit" Or _
    wksht.Name = "Urban_Code" Or wksht.Name = "Pavement_Messages" Then
        Application.DisplayAlerts = False
        wksht.Delete
        Application.DisplayAlerts = True
    End If
Next

'Save dataset as separate workbook
Sheets("Dataset").Move
sname = wd & "/Intersection_Input" & "_" & Replace(Date, "/", "-") & "_" & Replace(Time, ":", "-") & ".csv"
sname = Replace(sname, " ", "_")
sname = Replace(sname, "/", "\")
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("I5").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")
ActiveWorkbook.Close
Application.DisplayAlerts = True

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Intersections File Compiled."
    .Range("A3") = "Returning to UICPM Input Screen."
    .Range("A5") = ""
    .Range("B5") = ""
    .Range("A6") = "Finish Time"
    .Range("B6") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))

'Assign start time and end time
StartTime = Workbooks(guiwb).Sheets("Progress").Range("B4").Text
EndTime = Workbooks(guiwb).Sheets("Progress").Range("B6").Text

'Message box with start and end times
MsgBox "The Intersection Roadway input file has been successfully created and saved in the working directory folder. The following is a summary of the process:" & Chr(10) & _
Chr(10) & _
"Process: Intersection Roadway Input" & Chr(10) & _
"Start Time: " & StartTime & Chr(10) & _
"End Time: " & EndTime & Chr(10), vbOKOnly, "Process Complete"

End Sub

Sub IntControlIdentifier()
'Macro used to fill in blank intersection control cells.


'Declare variables
Dim row1 As Long
Dim ColIntControl, ColSig As Integer

'Find columns
ColIntControl = 1
Do Until Cells(1, ColIntControl) = "INT_CONTROL"
    ColIntControl = ColIntControl + 1
Loop

ColSig = 1
Do Until Cells(1, ColSig) = "SIGNALIZED"
    ColSig = ColSig + 1
Loop

'Fill in intersection control blanks
row1 = 2
Do While Cells(row1, ColSig) <> ""
    If Cells(row1, ColIntControl) = " " Or Cells(row1, ColIntControl) = "" Then
        If Cells(row1, ColSig) = "YES" Then
            Cells(row1, ColIntControl) = "SIGNAL"
        Else
            Cells(row1, ColIntControl) = "OTHER"
        End If
    End If
    row1 = row1 + 1
Loop

End Sub

Sub DupIntAnalysis()

'If this sub is uncommented it will need to be updated to inclue all Routes (0 through 4)

'Declare variables
Dim row1 As Integer
Dim LatDiffBoundary, LongDiffBoundary, LatDiff, LongDiff, MainMP As Double
Dim MainRoute, colRoute1, ColMP1 As Integer
Dim latcol, longcol, ElevCol As Integer
Dim datasheet As String

'Assign initial values
row1 = 2
'LatCount = 0
'LongCount = 0
'BothCount = 0
LatDiffBoundary = 0.001        'About 365 feet in latitude
LongDiffBoundary = 0.0013       'About 365 feet in longitude
datasheet = "Dataset"

latcol = 1
longcol = 1
ElevCol = 1
colRoute1 = 1
ColMP1 = 1

'Add milepoint columns
Do Until Sheets(datasheet).Cells(1, ColMP1) = "INT_MP_1"
    ColMP1 = ColMP1 + 1
Loop

Sheets(datasheet).Activate

Columns(ColMP1 + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, ColMP1 + 1) = "INT_MP_2"
Application.CutCopyMode = False

Columns(ColMP1 + 2).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, ColMP1 + 2) = "INT_MP_3"
Application.CutCopyMode = False

'Find columns
Do Until Sheets(datasheet).Cells(1, latcol) = "LATITUDE"
    latcol = latcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, longcol) = "LONGITUDE"
    longcol = longcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, ElevCol) = "ELEVATION"
    ElevCol = ElevCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute1) = "ROUTE_1"
    colRoute1 = colRoute1 + 1
Loop

'Sort by latitude, longitude, and elevation
LatLongElevSort

'Combine intersections
Do While Sheets(datasheet).Cells(row1 + 1, 1) <> ""
    LatDiff = Abs(Sheets(datasheet).Cells(row1 + 1, latcol) - Sheets(datasheet).Cells(row1, latcol))
    LongDiff = Abs(Sheets(datasheet).Cells(row1 + 1, longcol) - Sheets(datasheet).Cells(row1, longcol))
    
    MainRoute = Sheets(datasheet).Cells(row1 + 1, colRoute1)
    MainMP = Sheets(datasheet).Cells(row1 + 1, ColMP1)
    
    If LatDiff < LatDiffBoundary And LongDiff < LongDiffBoundary Then
        If MainRoute = Sheets(datasheet).Cells(row1, colRoute1) Then
            'Nothing
        ElseIf Sheets(datasheet).Cells(row1, colRoute1 + 1) = "" Or Sheets(datasheet).Cells(row1, colRoute1 + 1) = " " Then
            Sheets(datasheet).Cells(row1, colRoute1 + 1) = MainRoute
            Sheets(datasheet).Cells(row1, ColMP1 + 1) = MainMP
        ElseIf MainRoute = Sheets(datasheet).Cells(row1, colRoute1 + 1) Then
            Sheets(datasheet).Cells(row1, ColMP1 + 1) = MainMP
        ElseIf Sheets(datasheet).Cells(row1, colRoute1) = Sheets(datasheet).Cells(row1, colRoute1 + 1) Then
            Sheets(datasheet).Cells(row1, colRoute1 + 1) = MainRoute
            Sheets(datasheet).Cells(row1, ColMP1 + 1) = MainMP
        ElseIf MainRoute = Sheets(datasheet).Cells(row1, colRoute1 + 2) Then
            Sheets(datasheet).Cells(row1, ColMP1 + 2) = MainMP
        ElseIf Sheets(datasheet).Cells(row1, colRoute1 + 2) = "" Or Sheets(datasheet).Cells(row1, colRoute1 + 2) = " " Then
            Sheets(datasheet).Cells(row1, colRoute1 + 2) = MainRoute
            Sheets(datasheet).Cells(row1, ColMP1 + 2) = MainMP
        ElseIf Sheets(datasheet).Cells(row1, colRoute1) = Sheets(datasheet).Cells(row1, colRoute1 + 2) Then
            Sheets(datasheet).Cells(row1, colRoute1 + 2) = MainRoute
            Sheets(datasheet).Cells(row1, ColMP1 + 2) = MainMP
        End If
        
        'Average out latitude, longitude, and elevation values
        Sheets(datasheet).Cells(row1, latcol) = Application.WorksheetFunction.Average(Sheets(datasheet).Cells(row1, latcol), Sheets(datasheet).Cells(row1 + 1, latcol))
        Sheets(datasheet).Cells(row1, longcol) = Application.WorksheetFunction.Average(Sheets(datasheet).Cells(row1, longcol), Sheets(datasheet).Cells(row1 + 1, longcol))
        Sheets(datasheet).Cells(row1, ElevCol) = Application.WorksheetFunction.Average(Sheets(datasheet).Cells(row1, ElevCol), Sheets(datasheet).Cells(row1 + 1, ElevCol))
        
        'Delete row
        Rows(row1 + 1).EntireRow.Delete
        row1 = row1 - 1
        
    End If
    row1 = row1 + 1
Loop

'Eliminate intersections that do not have milepoints for at least secondary route.
row1 = 2
Do While Sheets(datasheet).Cells(row1, 1) <> ""
    If Sheets(datasheet).Cells(row1, colRoute1 + 2) = "" Or Sheets(datasheet).Cells(row1, colRoute1 + 2) = " " Then
        If Sheets(datasheet).Cells(row1, ColMP1 + 1) = "" Or Sheets(datasheet).Cells(row1, ColMP1 + 1) = " " Then
            Rows(row1).EntireRow.Delete
        Else
            Sheets(datasheet).Cells(row1, colRoute1 + 2) = ""
            row1 = row1 + 1
        End If
    Else
        If Sheets(datasheet).Cells(row1, ColMP1 + 1) = "" Or Sheets(datasheet).Cells(row1, ColMP1 + 1) = " " Then
            Sheets(datasheet).Cells(row1, colRoute1 + 1) = Sheets(datasheet).Cells(row1, colRoute1 + 2)
            Sheets(datasheet).Cells(row1, colRoute1 + 2) = ""
            
            Sheets(datasheet).Cells(row1, ColMP1 + 1) = Sheets(datasheet).Cells(row1, ColMP1 + 2)
            Sheets(datasheet).Cells(row1, ColMP1 + 2) = ""
            
            If Sheets(datasheet).Cells(row1, ColMP1 + 1) = "" Or Sheets(datasheet).Cells(row1, ColMP1 + 1) = " " Then
                Rows(row1).EntireRow.Delete
            Else
                row1 = row1 + 1
            End If
        ElseIf Sheets(datasheet).Cells(row1, ColMP1 + 2) = "" Or Sheets(datasheet).Cells(row1, ColMP1 + 2) = " " Then
            Sheets(datasheet).Cells(row1, colRoute1 + 2) = ""
            Sheets(datasheet).Cells(row1, ColMP1 + 2) = ""
            row1 = row1 + 1
        Else
            row1 = row1 + 1
        End If
    End If
Loop

End Sub

Sub FClassCalc()

'Declare variables
Dim row1, row2, col1 As Long
Dim routecol, MPcol As Integer
Dim RouteNum, colRoute0, ColMP0 As Integer          'Main route number, route column, and milepoint column
Dim colRoute1, ColMP1 As Integer                    'Secondary route column, and milepoint column
Dim colRoute2, ColMP2 As Integer                    'Tertiary route column, and milepoint column
Dim colRoute3, ColMP3 As Integer                    'Quarternary route column, and milepoint column
Dim colRoute4, ColMP4 As Integer                    'Quinary route column, and milepoint column
'Dim MaxFCCodeCol, MinFCCodeCol, MaxFCTypeCol, MinFCTypeCol As Integer
Dim FCCodeCol0, FCCodeCol1, FCCodeCol2, FCCodeCol3, FCCodeCol4 As Integer
Dim FCTypeCol0, FCTypeCol1, FCTypeCol2, FCTypeCol3, FCTypeCol4 As Integer
Dim FCRouteCol, FCMPCol1, FCMPCol2 As Integer       'Functional class route column, beginning milepoint column, and ending milepoint column
Dim FCCodeCol, FCTypeCol, FC_Code As Integer
Dim MP As Double
Dim datasheet, FCSheet, FC_Type As String
Dim EnterData, StopBool As Boolean
Dim j As Long

'Assign initial values to variables
row1 = 2
row2 = 2
col1 = 1

colRoute0 = 1
ColMP0 = 1
colRoute1 = 1
ColMP1 = 1
colRoute2 = 1
ColMP2 = 1
colRoute3 = 1
ColMP3 = 1
colRoute4 = 1
ColMP4 = 1

FCRouteCol = 1
FCMPCol1 = 1
FCMPCol2 = 1

FCCodeCol = 1
FCTypeCol = 1

datasheet = "Dataset"
FCSheet = "Functional_Class"

'Add column headings
Do Until Sheets(datasheet).Cells(1, col1) = ""
    col1 = col1 + 1
Loop

FCCodeCol0 = col1
FCCodeCol1 = col1 + 1
FCCodeCol2 = col1 + 2
FCCodeCol3 = col1 + 3
FCCodeCol4 = col1 + 4

FCTypeCol0 = col1 + 5
FCTypeCol1 = col1 + 6
FCTypeCol2 = col1 + 7
FCTypeCol3 = col1 + 8
FCTypeCol4 = col1 + 9

Sheets(datasheet).Cells(1, FCCodeCol0) = "FCCode_0"
Sheets(datasheet).Cells(1, FCCodeCol1) = "FCCode_1"
Sheets(datasheet).Cells(1, FCCodeCol2) = "FCCode_2"
Sheets(datasheet).Cells(1, FCCodeCol3) = "FCCode_3"
Sheets(datasheet).Cells(1, FCCodeCol4) = "FCCode_4"

Sheets(datasheet).Cells(1, FCTypeCol0) = "FCType_0"
Sheets(datasheet).Cells(1, FCTypeCol1) = "FCType_1"
Sheets(datasheet).Cells(1, FCTypeCol2) = "FCType_2"
Sheets(datasheet).Cells(1, FCTypeCol3) = "FCType_3"
Sheets(datasheet).Cells(1, FCTypeCol4) = "FCType_4"

'Find columns
Do Until Sheets(datasheet).Cells(1, colRoute0) = "ROUTE"
    colRoute0 = colRoute0 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP0) = "UDOT_BMP"
    ColMP0 = ColMP0 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute1) = "INT_RT_1"
    colRoute1 = colRoute1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP1) = "INT_RT_1_M"
    ColMP1 = ColMP1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute2) = "INT_RT_2"
    colRoute2 = colRoute2 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP2) = "INT_RT_2_M"
    ColMP2 = ColMP2 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute3) = "INT_RT_3"
    colRoute3 = colRoute3 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP3) = "INT_RT_3_M"
    ColMP3 = ColMP3 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute4) = "INT_RT_4"
    colRoute4 = colRoute4 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP4) = "INT_RT_4_M"
    ColMP4 = ColMP4 + 1
Loop


Do Until Sheets(FCSheet).Cells(1, FCRouteCol) = "ROUTE_ID"
    FCRouteCol = FCRouteCol + 1
Loop

Do Until Sheets(FCSheet).Cells(1, FCMPCol1) = "BEG_MILEPOINT"
    FCMPCol1 = FCMPCol1 + 1
Loop

Do Until Sheets(FCSheet).Cells(1, FCMPCol2) = "END_MILEPOINT"
    FCMPCol2 = FCMPCol2 + 1
Loop


Do Until Sheets(FCSheet).Cells(1, FCCodeCol) = "FC_CODE"
    FCCodeCol = FCCodeCol + 1
Loop

Do Until Sheets(FCSheet).Cells(1, FCTypeCol) = "FC_Type"
    FCTypeCol = FCTypeCol + 1
Loop


For j = 0 To 4
    row1 = 2
    row2 = 2
    StopBool = False
    
    If j = 0 Then
        'MAIN ROUTE DATA
        'Sort by main route/MP, secondary route/MP, tertiary route/MP
        MainRouteMPSort
        routecol = colRoute0
        MPcol = ColMP0
        
    ElseIf j = 1 Then
        'SECONDARY ROUTE DATA
        'Sort by secondary route/MP, main route/MP, tertiary route/MP
        SecondaryRouteMPSort
        routecol = colRoute1
        MPcol = ColMP1
        
    ElseIf j = 2 Then
        'TERTIARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP
        TertiaryRouteMPSort
        routecol = colRoute2
        MPcol = ColMP2
        
    ElseIf j = 3 Then
        'QUARTENARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP
        QuartenaryRouteMPSort
        routecol = colRoute3
        MPcol = ColMP3
            
    ElseIf j = 4 Then
        'QUINARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP
        QuinaryRouteMPSort
        routecol = colRoute4
        MPcol = ColMP4
                
    End If

    Dim myrow As Integer
    myrow = 2
    Dim myroute
    Do Until Sheets(FCSheet).Cells(myrow, 1) = ""
        myroute = CStr(Sheets(FCSheet).Cells(myrow, 1))
        myroute = Left(myroute, 4)
        If myroute = "089A" Then
            myroute = "0011"
        End If
        Sheets(FCSheet).Cells(myrow, 1) = myroute
        myrow = myrow + 1
    Loop

    'Assign FC code and label/type values
     Do While Sheets(datasheet).Cells(row1, 1) <> ""
        EnterData = False
        RouteNum = Sheets(datasheet).Cells(row1, routecol) 'Used to be: RouteNum = Int(Sheets(datasheet).Cells(row1, RouteCol))
        MP = Sheets(datasheet).Cells(row1, MPcol)
        
        Do While Sheets(FCSheet).Cells(row2, 1) <> ""
             If RouteNum = "Local" Then
                FC_Code = 7
                FC_Type = "Local"
                EnterData = True
                Exit Do
            ElseIf RouteNum = Empty Then
                StopBool = True
                Exit Do
            ElseIf Int(Sheets(FCSheet).Cells(row2, FCRouteCol)) > Int(RouteNum) Then     'Do if FC row route is higher than actual (meaning the intersection route doesn't exist)
                ''testing begin (by Camille)
                'Dim myMsg As String
                'myMsg = "Route " & Int(RouteNum) & " does not exist in the Functional Class file. Please take note. The code will continue upon clicking OK."
                'MsgBox myMsg, vbOKOnly, "Route Does Not Exist"
                ''testing end
                Sheets(datasheet).Rows(row1).EntireRow.Delete 'commented out by Camille
                row1 = row1 - 1
                Exit Do
            ElseIf Int(Sheets(FCSheet).Cells(row2, FCRouteCol)) = Int(RouteNum) And _
            Sheets(FCSheet).Cells(row2, FCMPCol1).Value < MP And _
            Sheets(FCSheet).Cells(row2, FCMPCol2).Value > MP Then     'Do if FC row has the same route and intersection is between milepoints
                FC_Code = Sheets(FCSheet).Cells(row2, FCCodeCol)
                FC_Type = Sheets(FCSheet).Cells(row2, FCTypeCol)
                EnterData = True
                Exit Do
            ElseIf Int(Sheets(FCSheet).Cells(row2, FCRouteCol)) = Int(RouteNum) And _
            Sheets(FCSheet).Cells(row2, FCMPCol1).Value = MP Then     'Do if FC row has the same route and intersection is at the beginning milepoint (MP = 0)
                FC_Code = Sheets(FCSheet).Cells(row2, FCCodeCol)
                FC_Type = Sheets(FCSheet).Cells(row2, FCTypeCol)
                EnterData = True
                Exit Do
            ElseIf Int(Sheets(FCSheet).Cells(row2, FCRouteCol)) = Int(RouteNum) And _
            Sheets(FCSheet).Cells(row2, FCMPCol2).Value = MP And _
            Sheets(FCSheet).Cells(row2 + 1, FCRouteCol) = Int(RouteNum) Then        'Do if FC row has the same route, the intersection is at the ending MP, and the next FC segment is part of the same route as the previous.
                If Application.min(Sheets(FCSheet).Cells(row2, FCCodeCol), Sheets(FCSheet).Cells(row2 + 1, FCCodeCol)) = Sheets(FCSheet).Cells(row2, FCCodeCol) Then
                    FC_Code = Sheets(FCSheet).Cells(row2, FCCodeCol)
                    FC_Type = Sheets(FCSheet).Cells(row2, FCTypeCol)
                    EnterData = True
                Else
                    FC_Code = Sheets(FCSheet).Cells(row2 + 1, FCCodeCol)
                    FC_Type = Sheets(FCSheet).Cells(row2 + 1, FCTypeCol)
                    EnterData = True
                End If
                Exit Do
            ElseIf Int(Sheets(FCSheet).Cells(row2, FCRouteCol)) = Int(RouteNum) And _
            Sheets(FCSheet).Cells(row2, FCMPCol2).Value = MP Then     'Do if FC row has the same route and the interseciton is at the ending MP
                FC_Code = Sheets(FCSheet).Cells(row2, FCCodeCol)
                FC_Type = Sheets(FCSheet).Cells(row2, FCTypeCol)
                EnterData = True
                Exit Do
            End If
            row2 = row2 + 1
        Loop
        
        'Camille: Assigns functional class to local roads
        If RouteNum = "Local" Then
            FC_Code = 7
            FC_Type = "Local"
            EnterData = True
        End If
        
        If EnterData = True Then
            If j = 0 Then
                Sheets(datasheet).Cells(row1, FCCodeCol0) = FC_Code
                Sheets(datasheet).Cells(row1, FCTypeCol0) = FC_Type
            ElseIf j = 1 Then
                Sheets(datasheet).Cells(row1, FCCodeCol1) = FC_Code
                Sheets(datasheet).Cells(row1, FCTypeCol1) = FC_Type
                'Camille: I'm actually not sure if we want it to make default values. But that's what it does here.
                If Sheets(datasheet).Cells(row1, FCCodeCol1) = "" And (Sheets(datasheet).Cells(row1, colRoute1) <> "" Or Sheets(datasheet).Cells(row2, colRoute2) <> "Local") Then
                    Sheets(datasheet).Cells(row1, FCCodeCol1) = Sheets(datasheet).Cells(row1, FCCodeCol0)
                    Sheets(datasheet).Cells(row1, FCTypeCol1) = Sheets(datasheet).Cells(row1, FCTypeCol0)
                End If
            ElseIf j = 2 Then
                Sheets(datasheet).Cells(row1, FCCodeCol2) = FC_Code
                Sheets(datasheet).Cells(row1, FCTypeCol2) = FC_Type
                'Camille: I'm actually not sure if we want it to make default values. But that's what it does here.
                If Sheets(datasheet).Cells(row1, FCCodeCol2) = "" And (Sheets(datasheet).Cells(row1, colRoute2) <> "" Or Sheets(datasheet).Cells(row2, colRoute2) <> "Local") Then
                    Sheets(datasheet).Cells(row1, FCCodeCol2) = Sheets(datasheet).Cells(row1, FCCodeCol0)
                    Sheets(datasheet).Cells(row1, FCTypeCol2) = Sheets(datasheet).Cells(row1, FCTypeCol0)
                End If
            ElseIf j = 3 Then
                Sheets(datasheet).Cells(row1, FCCodeCol3) = FC_Code
                Sheets(datasheet).Cells(row1, FCTypeCol3) = FC_Type
                 'Camille: I'm actually not sure if we want it to make default values. But that's what it does here.
                If Sheets(datasheet).Cells(row1, FCCodeCol3) = "" And (Sheets(datasheet).Cells(row1, colRoute3) <> "" Or Sheets(datasheet).Cells(row2, colRoute3) <> "Local") Then
                    Sheets(datasheet).Cells(row1, FCCodeCol3) = Sheets(datasheet).Cells(row1, FCCodeCol0)
                    Sheets(datasheet).Cells(row1, FCTypeCol3) = Sheets(datasheet).Cells(row1, FCTypeCol0)
                End If
            ElseIf j = 4 Then
                Sheets(datasheet).Cells(row1, FCCodeCol4) = FC_Code
                Sheets(datasheet).Cells(row1, FCTypeCol4) = FC_Type
                 'Camille: I'm actually not sure if we want it to make default values. But that's what it does here.
                If Sheets(datasheet).Cells(row1, FCCodeCol4) = "" And (Sheets(datasheet).Cells(row1, colRoute4) <> "" Or Sheets(datasheet).Cells(row2, colRoute4) <> "Local") Then
                    Sheets(datasheet).Cells(row1, FCCodeCol4) = Sheets(datasheet).Cells(row1, FCCodeCol0)
                    Sheets(datasheet).Cells(row1, FCTypeCol4) = Sheets(datasheet).Cells(row1, FCTypeCol0)
                End If
            End If
        End If
        If StopBool = True Then
            Exit Do
        End If
        row1 = row1 + 1
    Loop
Next j


'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
MainRouteMPSort

End Sub


Sub EntVehCalc()

'Declare variables
Dim row1, row2, i, j, col1, col2 As Long
Dim routecol, MPcol, NumAADTYears As Integer
Dim RouteNum, colRoute0, ColMP0 As Integer          'Main route number, route column, and milepoint column
Dim colRoute1, ColMP1 As Integer                    'Secondary route column, and milepoint column
Dim colRoute2, ColMP2 As Integer                    'Tertiary route column, and milepoint column
Dim colRoute3, ColMP3 As Integer                    'Quartenary route column, and milepoint column
Dim colRoute4, ColMP4 As Integer                    'Quinary route column, and milepoint column
Dim AADTCountyCol, AADTRegionCol As Integer         'AADT County and Region columns

Dim EntVehCol1 As Integer
Dim AADTRouteCol, AADTMPCol1, AADTMPCol2 As Integer       'AADT route column, beginning milepoint column, and ending milepoint column
Dim AADTCol, STrkPercCol, CTrkPercCol, PercTrkCol As Long
Dim CountyCol, RegionCol As Integer
Dim TotTrkPerc As Double

Dim MP As Double
Dim datasheet, AADTSheet As String
Dim County As String
Dim Region As Integer
Dim EnterData, LocalRoad, StopBool As Boolean                     'LocalRoad added by Camille
Dim LocalAdd As Integer                                           'LocalAdd added by Camille
Dim TControlCol As Integer                                        'TControlCol added by Camille

Dim NumLegsCol As Integer
Dim NumLegs As Integer

'Assign initial values to variables

colRoute0 = 1
ColMP0 = 1
colRoute1 = 1
ColMP1 = 1
colRoute2 = 1
ColMP2 = 1
colRoute3 = 1
ColMP3 = 1
colRoute4 = 1
ColMP4 = 1

AADTRouteCol = 1
AADTMPCol1 = 1
AADTMPCol2 = 1
AADTCountyCol = 1
AADTRegionCol = 1

EntVehCol1 = 1
AADTCol = 1
PercTrkCol = 1
STrkPercCol = 1
CTrkPercCol = 1

TControlCol = 1

datasheet = "Dataset"
AADTSheet = "AADT"

'Find columns

Do Until Sheets(datasheet).Cells(1, colRoute0) = "ROUTE"
    colRoute0 = colRoute0 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP0) = "UDOT_BMP"
    ColMP0 = ColMP0 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute1) = "INT_RT_1"
    colRoute1 = colRoute1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP1) = "INT_RT_1_M"
    ColMP1 = ColMP1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute2) = "INT_RT_2"
    colRoute2 = colRoute2 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP2) = "INT_RT_2_M"
    ColMP2 = ColMP2 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute3) = "INT_RT_3"
    colRoute3 = colRoute3 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP3) = "INT_RT_3_M"
    ColMP3 = ColMP3 + 1
Loop

Do Until Sheets(datasheet).Cells(1, colRoute4) = "INT_RT_4"
    colRoute4 = colRoute4 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ColMP4) = "INT_RT_4_M"
    ColMP4 = ColMP4 + 1
Loop


Do Until Sheets(AADTSheet).Cells(1, AADTRouteCol) = "ROUTE_ID"
    AADTRouteCol = AADTRouteCol + 1
Loop

Do Until Sheets(AADTSheet).Cells(1, AADTMPCol1) = "BEG_MILEPOINT"
    AADTMPCol1 = AADTMPCol1 + 1
Loop

Do Until Sheets(AADTSheet).Cells(1, AADTMPCol2) = "END_MILEPOINT"
    AADTMPCol2 = AADTMPCol2 + 1
Loop


Do Until InStr(1, Sheets(AADTSheet).Cells(1, AADTCol), "AADT") > 0
    AADTCol = AADTCol + 1
Loop

Do Until Sheets(AADTSheet).Cells(1, STrkPercCol) = "Single_Percent"
    STrkPercCol = STrkPercCol + 1
Loop

Do Until Sheets(AADTSheet).Cells(1, CTrkPercCol) = "Combo_Percent"
    CTrkPercCol = CTrkPercCol + 1
Loop

Do Until Sheets(AADTSheet).Cells(1, AADTCountyCol) = "COUNTY"
    AADTCountyCol = AADTCountyCol + 1
Loop

Do Until Sheets(AADTSheet).Cells(1, AADTRegionCol) = "REGION"
    AADTRegionCol = AADTRegionCol + 1
Loop

Do Until Sheets("Dataset").Cells(1, TControlCol) = "TRAFFIC_CO"
    TControlCol = TControlCol + 1
Loop

'Add percent trucks heading
Do Until Sheets(datasheet).Cells(1, PercTrkCol) = ""
    PercTrkCol = PercTrkCol + 1
Loop
CountyCol = PercTrkCol + 1
RegionCol = PercTrkCol + 2

Sheets(datasheet).Cells(1, PercTrkCol) = "PERCENT_TRUCKS"
Sheets(datasheet).Cells(1, CountyCol) = "COUNTY"
Sheets(datasheet).Cells(1, RegionCol) = "REGION"

'Add in entering vehicle column headings
EntVehCol1 = PercTrkCol + 3
col1 = EntVehCol1
col2 = AADTCol
Do Until InStr(1, Sheets(AADTSheet).Cells(1, col2), "AADT") < 1       'finds the last column of AADT data
    Sheets(datasheet).Cells(1, col1) = Replace(Sheets(AADTSheet).Cells(1, col2), "AADT", "ENT_VEH")
    col1 = col1 + 1
    col2 = col2 + 1
Loop

'Add number of legs column
NumLegsCol = 1
Do Until Sheets(datasheet).Cells(1, NumLegsCol) = ""
    NumLegsCol = NumLegsCol + 1
Loop
Sheets(datasheet).Cells(1, NumLegsCol) = "NUM_LEGS"

NumAADTYears = col2 - AADTCol
ReDim AADT(1 To NumAADTYears)


For j = 0 To 4
       
    row1 = 2
    row2 = 2
    StopBool = False
    
    If j = 0 Then
        'MAIN ROUTE DATA
        'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        MainRouteMPSort
        routecol = colRoute0
        MPcol = ColMP0
        
    ElseIf j = 1 Then
        'SECONDARY ROUTE DATA
        'Sort by secondary route/MP, main route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        SecondaryRouteMPSort
        routecol = colRoute1
        MPcol = ColMP1
        
   ElseIf j = 2 Then
        'TERTIARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP, quartenary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute2
        MPcol = ColMP2
        
    ElseIf j = 3 Then
        'QUARTENARY ROUTE DATA
        'Sort by quartenary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quinary route/MP
        QuartenaryRouteMPSort
        routecol = colRoute3
        MPcol = ColMP3
    
    ElseIf j = 4 Then
        'QUINARY ROUTE DATA
        'Sort by quinary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP
        QuinaryRouteMPSort
        routecol = colRoute4
        MPcol = ColMP4
        
    End If
    
    
    'Assign AADT code and label/type values
    Do While Sheets(datasheet).Cells(row1, 1) <> ""
        
        'testing
        'If Sheets(datasheet).Cells(row1, 1) = 519 Or Sheets(datasheet).Cells(row1, 1) = 1730 Or Sheets(datasheet).Cells(row1, 1) = 859 Or Sheets(datasheet).Cells(row1, 1) = 327 Then
        '    row1 = row1
        'End If
        'testing
        If Sheets(datasheet).Cells(row1, 1) = 1730 Or Sheets(datasheet).Cells(row1, 1) = 859 Or Sheets(datasheet).Cells(row1, 1) = 12 Or Sheets(datasheet).Cells(row1, 1) = 435 Then
            row1 = row1
        End If
        'testing
        'If Sheets(datasheet).Cells(row1, 1) = 980 Or Sheets(datasheet).Cells(row1, 1) = 1059 Or Sheets(datasheet).Cells(row1, 1) = 972 Or Sheets(datasheet).Cells(row1, 1) = 954 Then
        '    row1 = row1
        'End If
    
    'testing
    'row1 = 1155
    'row2 = 4805
    
        EnterData = False
        LocalRoad = False                              'added by Camille
        RouteNum = Sheets(datasheet).Cells(row1, routecol)
        MP = Round(Sheets(datasheet).Cells(row1, MPcol), 3)
        
        Do While Sheets(datasheet).Cells(row1, 1) <> ""
            County = Sheets(AADTSheet).Cells(row2, AADTCountyCol)
            Region = Sheets(AADTSheet).Cells(row2, AADTRegionCol)
            If RouteNum = "Local" Then                          'this part of the if statement added by Camille on April 10, 2019
                If Sheets(datasheet).Cells(row1, TControlCol) = "SIGNAL" Then
                    LocalAdd = 1000
                Else
                    LocalAdd = 500
                End If
                LocalRoad = True
                Exit Do
            ElseIf Int(RouteNum) = Empty Then
                StopBool = True
                Exit Do
            ElseIf Int(Sheets(AADTSheet).Cells(row2, AADTRouteCol)) > Int(RouteNum) Then     'Do if AADT row route is higher than actual (meaning the intersection route doesn't exist)
                Sheets(datasheet).Rows(row1).EntireRow.Delete
                row1 = row1 - 1       'Camille: October 29, 2019... does this fix the ghost block problem? or does it mess something else up?
                Exit Do
            ElseIf Int(Sheets(AADTSheet).Cells(row2, AADTRouteCol)) = Int(RouteNum) And _
            Sheets(AADTSheet).Cells(row2, AADTMPCol1).Value < MP And _
            Sheets(AADTSheet).Cells(row2, AADTMPCol2).Value > MP Then     'Case 1: Do if AADT row has the same route and intersection is between milepoints
                For i = 0 To NumAADTYears - 1
                    AADT(i + 1) = Sheets(AADTSheet).Cells(row2, AADTCol + i)
                Next i
                TotTrkPerc = Sheets(AADTSheet).Cells(row2, STrkPercCol) + Sheets(AADTSheet).Cells(row2, CTrkPercCol)
                EnterData = True
                NumLegs = 1   'changed from 2 to 1 beccause I was getting NumLegs total = 6, 7, and 8
                Exit Do
            ElseIf Int(Sheets(AADTSheet).Cells(row2, AADTRouteCol)) = Int(RouteNum) And _
            Sheets(AADTSheet).Cells(row2, AADTMPCol1).Value = MP Then     'Case 2: Do if AADT row has the same route and intersection is at the beginning milepoint (MP = 0)
                For i = 0 To NumAADTYears - 1
                    AADT(i + 1) = 0.5 * Sheets(AADTSheet).Cells(row2, AADTCol + i)
                Next i
                TotTrkPerc = Sheets(AADTSheet).Cells(row2, STrkPercCol) + Sheets(AADTSheet).Cells(row2, CTrkPercCol)
                EnterData = True
                NumLegs = 1
                Exit Do
            ElseIf Int(Sheets(AADTSheet).Cells(row2, AADTRouteCol)) = Int(RouteNum) And _
            Sheets(AADTSheet).Cells(row2, AADTMPCol2).Value = MP And _
            Sheets(AADTSheet).Cells(row2 + 1, AADTRouteCol) = Int(RouteNum) Then        'Case 3: Do if AADT row has the same route, the intersection is at the ending MP, and the next AADT segment is part of the same route as the previous.
                For i = 0 To NumAADTYears - 1
                    AADT(i + 1) = 0.5 * (Sheets(AADTSheet).Cells(row2, AADTCol + i) + Sheets(AADTSheet).Cells(row2 + 1, AADTCol + i))
                Next i
                TotTrkPerc = 0.5 * (Sheets(AADTSheet).Cells(row2, STrkPercCol) + Sheets(AADTSheet).Cells(row2, CTrkPercCol) + _
                    Sheets(AADTSheet).Cells(row2 + 1, STrkPercCol) + Sheets(AADTSheet).Cells(row2 + 1, CTrkPercCol))
                EnterData = True
                NumLegs = 1  'changed from 2 to 1 beccause I was getting NumLegs total = 6, 7, and 8
                Exit Do
            ElseIf Int(Sheets(AADTSheet).Cells(row2, AADTRouteCol)) = Int(RouteNum) And _
            Sheets(AADTSheet).Cells(row2, AADTMPCol2).Value = MP Then     'Case 4: Do if AADT row has the same route and the interseciton is at the ending MP
                For i = 0 To NumAADTYears - 1
                    AADT(i + 1) = 0.5 * Sheets(AADTSheet).Cells(row2, AADTCol + i)
                Next i
                TotTrkPerc = Sheets(AADTSheet).Cells(row2, STrkPercCol) + Sheets(AADTSheet).Cells(row2, CTrkPercCol)
                EnterData = True
                NumLegs = 1
                Exit Do
            ElseIf Sheets(AADTSheet).Cells(row2, AADTRouteCol) = Empty Then
                Sheets(datasheet).Rows(row1).EntireRow.Delete
                row1 = row1 - 1       'Camille: October 29, 2019... does this fix the ghost block problem? or does it mess something else up?
                Exit Do
            End If
            row2 = row2 + 1
        Loop
        If EnterData = True Then
            For i = 0 To NumAADTYears - 1
                Sheets(datasheet).Cells(row1, EntVehCol1 + i) = Sheets(datasheet).Cells(row1, EntVehCol1 + i) + AADT(i + 1)
            Next i
            
            If Sheets(datasheet).Cells(row1, PercTrkCol) = "" Or Sheets(datasheet).Cells(row1, PercTrkCol) = " " Then
                Sheets(datasheet).Cells(row1, PercTrkCol) = TotTrkPerc
            Else
                Sheets(datasheet).Cells(row1, PercTrkCol) = ((Sheets(datasheet).Cells(row1, EntVehCol1) * Sheets(datasheet).Cells(row1, PercTrkCol)) + _
                (TotTrkPerc * AADT(1))) / (AADT(1) + Sheets(datasheet).Cells(row1, EntVehCol1))
            End If
            
            If Sheets(datasheet).Cells(row1, CountyCol) = "" Or Sheets(datasheet).Cells(row1, CountyCol) = " " Then
                Sheets(datasheet).Cells(row1, CountyCol) = County
            End If
            
            If Sheets(datasheet).Cells(row1, RegionCol) = "" Or Sheets(datasheet).Cells(row1, RegionCol) = " " Then
                Sheets(datasheet).Cells(row1, RegionCol) = Region
            End If
            
            Sheets(datasheet).Cells(row1, NumLegsCol) = Sheets(datasheet).Cells(row1, NumLegsCol) + NumLegs
        End If
        
        If LocalRoad = True Then                             ' added by Camille on April 10, 2019  'the lines commented out inside the if statement were copied from the If EnterData = True statement and I don't think they will be needed if it's a local road but tbh I don't really understand what they do.
            For i = 0 To NumAADTYears - 1
                If Sheets(datasheet).Cells(row1, EntVehCol1 + i) <> 0 Then
                    Sheets(datasheet).Cells(row1, EntVehCol1 + i) = Sheets(datasheet).Cells(row1, EntVehCol1 + i) + LocalAdd
                End If
            Next i
      
        '    If Sheets(datasheet).Cells(row1, PercTrkCol) = "" Or Sheets(datasheet).Cells(row1, PercTrkCol) = " " Then
        '        Sheets(datasheet).Cells(row1, PercTrkCol) = TotTrkPerc
        '    Else
        '        Sheets(datasheet).Cells(row1, PercTrkCol) = ((Sheets(datasheet).Cells(row1, EntVehCol1) * Sheets(datasheet).Cells(row1, PercTrkCol)) + _
        '        (TotTrkPerc * AADT(1))) / (AADT(1) + Sheets(datasheet).Cells(row1, EntVehCol1))
        '    End If
        '
        '    If Sheets(datasheet).Cells(row1, CountyCol) = "" Or Sheets(datasheet).Cells(row1, CountyCol) = " " Then
        '        Sheets(datasheet).Cells(row1, CountyCol) = County
        '    End If
        '
        '    If Sheets(datasheet).Cells(row1, RegionCol) = "" Or Sheets(datasheet).Cells(row1, RegionCol) = " " Then
        '        Sheets(datasheet).Cells(row1, RegionCol) = Region
        '    End If
        '
            Sheets(datasheet).Cells(row1, NumLegsCol) = Sheets(datasheet).Cells(row1, NumLegsCol) + 1  'uncommented by Camille; changed from NumLegs to 1 by Camille
        
        End If
        
        If StopBool = True Then
            Exit Do
        End If
        row1 = row1 + 1
    Loop
Next j


'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
MainRouteMPSort

End Sub

Sub SpeedLimitCalc()

'Declare variables
Dim row1, row2, i, j As Long
Dim routecol, MPcol As Integer
Dim RouteNum, colRoute0, ColMP0 As Integer          'Main route number, route column, and milepoint column
Dim colRoute1, ColMP1 As Integer                    'Secondary route column, and milepoint column
Dim colRoute2, ColMP2 As Integer                    'Tertiary route column, and milepoint column
Dim colRoute3, ColMP3 As Integer                    'Quartenary route column, and milepoint column
Dim colRoute4, ColMP4 As Integer                    'Quinary route column, and milepoint column

'Dim MaxSLCol, MinSLCol As Integer
Dim IntSLCol0, IntSLCol1, IntSLCol2, IntSLCol3, IntSLCol4 As Integer
Dim SLRouteCol, SLMPCol1, SLMPCol2, SpeedLimit        'AADT route column, beginning milepoint column, and ending milepoint column
Dim SLcol As Long

Dim MP As Double
Dim datasheet As String, SLSheet As String
Dim EnterData, StopBool As Boolean

'Assign initial values to variables

colRoute1 = 1
ColMP1 = 1
colRoute2 = 1
ColMP2 = 1
colRoute3 = 1
ColMP3 = 1

SLRouteCol = 1
SLMPCol1 = 1
SLMPCol2 = 1

'MaxSLCol = 1
'MinSLCol = 1
IntSLCol0 = 1
IntSLCol1 = 1
IntSLCol2 = 1
IntSLCol3 = 1
IntSLCol4 = 1
SLcol = 1

datasheet = "Dataset"
SLSheet = "Speed_Limit"

'Find columns
colRoute0 = FindColumn("ROUTE", datasheet, 1, True)
colRoute1 = FindColumn("INT_RT_1", datasheet, 1, True)
colRoute2 = FindColumn("INT_RT_2", datasheet, 1, True)
colRoute3 = FindColumn("INT_RT_3", datasheet, 1, True)
colRoute4 = FindColumn("INT_RT_4", datasheet, 1, True)
ColMP0 = FindColumn("UDOT_BMP", datasheet, 1, True)
ColMP1 = FindColumn("INT_RT_1_M", datasheet, 1, True)
ColMP2 = FindColumn("INT_RT_2_M", datasheet, 1, True)
ColMP3 = FindColumn("INT_RT_3_M", datasheet, 1, True)
ColMP4 = FindColumn("INT_RT_4_M", datasheet, 1, True)

Do Until Sheets(SLSheet).Cells(1, SLRouteCol) = "ROUTE_ID"
    SLRouteCol = SLRouteCol + 1
Loop

Do Until Sheets(SLSheet).Cells(1, SLMPCol1) = "BEG_MILEPOINT"
    SLMPCol1 = SLMPCol1 + 1
Loop

Do Until Sheets(SLSheet).Cells(1, SLMPCol2) = "END_MILEPOINT"
    SLMPCol2 = SLMPCol2 + 1
Loop

Do Until Sheets(SLSheet).Cells(1, SLcol) = "SPEED_LIMIT"
    SLcol = SLcol + 1
Loop

'Add max speed limit column headings
'Do Until Sheets(DataSheet).Cells(1, MaxSLCol) = ""
'    MaxSLCol = MaxSLCol + 1
'Loop
'MinSLCol = MaxSLCol + 1
'Sheets(DataSheet).Cells(1, MaxSLCol) = "MAX_SPEED_LIMIT"
'Sheets(DataSheet).Cells(1, MinSLCol) = "MIN_SPEED_LIMIT"

Do Until Sheets(datasheet).Cells(1, IntSLCol0) = ""
    IntSLCol0 = IntSLCol0 + 1
Loop
IntSLCol1 = IntSLCol0 + 1
IntSLCol2 = IntSLCol0 + 2
IntSLCol3 = IntSLCol0 + 3
IntSLCol4 = IntSLCol0 + 4
Sheets(datasheet).Cells(1, IntSLCol0) = "SPEED_LIMIT_0"
Sheets(datasheet).Cells(1, IntSLCol1) = "SPEED_LIMIT_1"
Sheets(datasheet).Cells(1, IntSLCol2) = "SPEED_LIMIT_2"
Sheets(datasheet).Cells(1, IntSLCol3) = "SPEED_LIMIT_3"
Sheets(datasheet).Cells(1, IntSLCol4) = "SPEED_LIMIT_4"

For j = 0 To 4
    row1 = 2
    row2 = 2
    StopBool = False
    
    If j = 0 Then
        'MAIN ROUTE DATA
        'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        MainRouteMPSort
        routecol = colRoute0
        MPcol = ColMP0
        
    ElseIf j = 1 Then
        'SECONDARY ROUTE DATA
        'Sort by secondary route/MP, main route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        SecondaryRouteMPSort
        routecol = colRoute1
        MPcol = ColMP1
        
   ElseIf j = 2 Then
        'TERTIARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP, quartenary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute2
        MPcol = ColMP2
        
    ElseIf j = 3 Then
        'QUARTENARY ROUTE DATA
        'Sort by quartenary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute3
        MPcol = ColMP3
    
    ElseIf j = 4 Then
        'QUINARY ROUTE DATA
        'Sort by quinary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP
        QuinaryRouteMPSort
        routecol = colRoute4
        MPcol = ColMP4
        
    End If
    
    
    'Assign Speed Limit and label/type values
    Do While Sheets(datasheet).Cells(row1, 1) <> ""
        EnterData = False
        RouteNum = Sheets(datasheet).Cells(row1, routecol)
        MP = Sheets(datasheet).Cells(row1, MPcol)
        
        Do While Sheets(SLSheet).Cells(row2, 1) <> ""
            
            If RouteNum = "Local" Then            'If it's Local, assume speed limit is 0.
                SpeedLimit = ""
                EnterData = True
                Exit Do
            ElseIf Int(RouteNum) > 491 Then       'If it's FedAid, assume speed limit is 0.
                SpeedLimit = ""
                EnterData = True
                Exit Do
            ElseIf Int(RouteNum) = Empty Then
                StopBool = True
                Exit Do
            ElseIf Int(Sheets(SLSheet).Cells(row2, SLRouteCol)) > Int(RouteNum) Then     'Do if SL row route is higher than actual (meaning the intersection route doesn't exist)
                'testing begin (by Camille)
                'Dim myMsg As String
                'myMsg = "Route " & Int(RouteNum) & " does not exist in the Speed Limit file. Please take note. The code will continue upon clicking OK."
                'MsgBox myMsg, vbOKOnly, "Route Does Not Exist"
                'testing end
                If j = 0 Then
                    Sheets(datasheet).Rows(row1).EntireRow.Delete    'commented out by Camille
                End If
                Exit Do
            ElseIf Int(Sheets(SLSheet).Cells(row2, SLRouteCol)) = Int(RouteNum) And _
            Sheets(SLSheet).Cells(row2, SLMPCol1).Value < MP And _
            Sheets(SLSheet).Cells(row2, SLMPCol2).Value > MP Then     'Do if SL row has the same route and intersection is between milepoints
                SpeedLimit = Sheets(SLSheet).Cells(row2, SLcol)
                EnterData = True
                Exit Do
            ElseIf Int(Sheets(SLSheet).Cells(row2, SLRouteCol)) = Int(RouteNum) And _
            Sheets(SLSheet).Cells(row2, SLMPCol1).Value = MP Then     'Do if SL row has the same route and intersection is at the beginning milepoint (MP = 0)
                SpeedLimit = Sheets(SLSheet).Cells(row2, SLcol)
                EnterData = True
                Exit Do
            ElseIf Int(Sheets(SLSheet).Cells(row2, SLRouteCol)) = Int(RouteNum) And _
            Sheets(SLSheet).Cells(row2, SLMPCol2).Value = MP And _
            Sheets(SLSheet).Cells(row2 + 1, SLRouteCol) = Int(RouteNum) Then        'Do if SL row has the same route, the intersection is at the ending MP, and the next FC segment is part of the same route as the previous.
                If Application.Max(Sheets(SLSheet).Cells(row2, SLcol), Sheets(SLSheet).Cells(row2 + 1, SLcol)) = Sheets(SLSheet).Cells(row2, SLcol) Then
                    SpeedLimit = Sheets(SLSheet).Cells(row2, SLcol)
                    EnterData = True
                Else
                    SpeedLimit = Sheets(SLSheet).Cells(row2 + 1, SLcol)
                    EnterData = True
                End If
                Exit Do
            ElseIf Int(Sheets(SLSheet).Cells(row2, SLRouteCol)) = Int(RouteNum) And _
            Sheets(SLSheet).Cells(row2, SLMPCol2).Value = MP Then     'Do if SL row has the same route and the interseciton is at the ending MP
                SpeedLimit = Sheets(SLSheet).Cells(row2, SLcol)
                EnterData = True
                Exit Do
            End If
            row2 = row2 + 1
        Loop
        If EnterData = True Then
            'If SpeedLimit > Sheets(DataSheet).Cells(row1, MaxSLCol) Or Sheets(DataSheet).Cells(row1, MaxSLCol) = "" Or _
            'Sheets(DataSheet).Cells(row1, MaxSLCol) = " " Then
            '    Sheets(DataSheet).Cells(row1, MaxSLCol) = SpeedLimit
            'End If
            'If SpeedLimit < Sheets(DataSheet).Cells(row1, MinSLCol) Or Sheets(DataSheet).Cells(row1, MinSLCol) = "" Or _
            'Sheets(DataSheet).Cells(row1, MinSLCol) = " " Then
            '    Sheets(DataSheet).Cells(row1, MinSLCol) = SpeedLimit
            'End If
            
            If j = 0 Then
                Sheets(datasheet).Cells(row1, IntSLCol0) = SpeedLimit
            ElseIf j = 1 Then
                Sheets(datasheet).Cells(row1, IntSLCol1) = SpeedLimit
            ElseIf j = 2 Then
                Sheets(datasheet).Cells(row1, IntSLCol2) = SpeedLimit
            ElseIf j = 3 Then
                Sheets(datasheet).Cells(row1, IntSLCol3) = SpeedLimit
            ElseIf j = 4 Then
                Sheets(datasheet).Cells(row1, IntSLCol4) = SpeedLimit
            End If
        End If
        If StopBool = True Then
            Exit Do
        End If
        row1 = row1 + 1
    Loop
Next j


'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
MainRouteMPSort


End Sub

Sub NumLanesCalc()

'Declare variables
Dim row1, row2, i, j As Long
Dim routecol, MPcol As Integer

Dim RouteNum, colRoute0, ColMP0 As Integer          'Main route number, route column, and milepoint column
Dim colRoute1, ColMP1 As Integer                    'Secondary route column, and milepoint column
Dim colRoute2, ColMP2 As Integer                    'Tertiary route column, and milepoint column
Dim colRoute3, ColMP3 As Integer                    'Quarternary route column, and milepoint column
Dim colRoute4, ColMP4 As Integer                    'Quinary route column, and milepoint column

Dim MaxLanesCol, MinLanesCol As Integer
Dim LanesColRoute, LanesColMP1, LanesColMP2, Num_Lanes, LaneWidth As Integer       'AADT route column, beginning milepoint column, and ending milepoint column
Dim LanesCol, LaneWidthCol, MaxRWCol, MinRWCol As Long
Dim IntDistCol0, IntDistCol1, IntDistCol2, IntDistCol3, IntDistCol4 As Integer
Dim IntDist0, IntDist1, IntDist2, IntDist3, IntDist4 As Integer

Dim MP As Double
Dim datasheet As String
Dim LanesSheet As String
Dim EnterData, StopBool As Boolean

Dim MaxAdjust, MinAdjust As Integer                 'Adjust values to improve road width accuracy. Based on manual measurements of intersections and difference between
                                                    'measured and calculated values. JG - 11/14/17
'Assign initial values to variables
colRoute0 = 1
ColMP0 = 1
colRoute1 = 1
ColMP1 = 1
colRoute2 = 1
ColMP2 = 1
colRoute3 = 1
ColMP3 = 1
colRoute4 = 1
ColMP4 = 1

LanesColRoute = 1
LanesColMP1 = 1
LanesColMP2 = 1

MaxLanesCol = 1
MinLanesCol = 1
MaxRWCol = 1
MinRWCol = 1
LanesCol = 1
LaneWidthCol = 1

datasheet = "Dataset"
LanesSheet = "Thru_Lanes"
                                 'SRSR intersection adjustment --should we measure some SR-nonSR ones by hand?
MaxAdjust = 84                      'These values were found by measuring almost 300 intersections by hand. The max width was off by 84 feet on average.
MinAdjust = 72                          'The min width was off by 72 feet on average.

'Find columns
colRoute0 = FindColumn("ROUTE", datasheet, 1, True)
colRoute1 = FindColumn("INT_RT_1", datasheet, 1, True)
colRoute2 = FindColumn("INT_RT_2", datasheet, 1, True)
colRoute3 = FindColumn("INT_RT_3", datasheet, 1, True)
colRoute4 = FindColumn("INT_RT_4", datasheet, 1, True)
ColMP0 = FindColumn("UDOT_BMP", datasheet, 1, True)
ColMP1 = FindColumn("INT_RT_1_M", datasheet, 1, True)
ColMP2 = FindColumn("INT_RT_2_M", datasheet, 1, True)
ColMP3 = FindColumn("INT_RT_3_M", datasheet, 1, True)
ColMP4 = FindColumn("INT_RT_4_M", datasheet, 1, True)

Do Until Sheets(LanesSheet).Cells(1, LanesColRoute) = "ROUTE_ID"
    LanesColRoute = LanesColRoute + 1
Loop

Do Until Sheets(LanesSheet).Cells(1, LanesColMP1) = "BEG_MILEPOINT"
    LanesColMP1 = LanesColMP1 + 1
Loop

Do Until Sheets(LanesSheet).Cells(1, LanesColMP2) = "END_MILEPOINT"
    LanesColMP2 = LanesColMP2 + 1
Loop


Do Until Sheets(LanesSheet).Cells(1, LanesCol) = "NUM_LANES"
    LanesCol = LanesCol + 1
Loop

Do Until Sheets(LanesSheet).Cells(1, LaneWidthCol) = "LANE_WIDTH"
    LaneWidthCol = LaneWidthCol + 1
Loop


IntDistCol0 = FindColumn("Int_Dist_0", datasheet, 1, True)
IntDistCol1 = FindColumn("Int_Dist_1", datasheet, 1, True)
IntDistCol2 = FindColumn("Int_Dist_2", datasheet, 1, True)
IntDistCol3 = FindColumn("Int_Dist_3", datasheet, 1, True)
IntDistCol4 = FindColumn("Int_Dist_4", datasheet, 1, True)

'Add lanes and roadway width columns
Do Until Sheets(datasheet).Cells(1, MaxLanesCol) = ""
    MaxLanesCol = MaxLanesCol + 1
Loop
MinLanesCol = MaxLanesCol + 1
MaxRWCol = MinLanesCol + 1
MinRWCol = MaxRWCol + 1
Sheets(datasheet).Cells(1, MaxLanesCol) = "MAX_NUM_LANES"
Sheets(datasheet).Cells(1, MinLanesCol) = "MIN_NUM_LANES"
Sheets(datasheet).Cells(1, MaxRWCol) = "MAX_ROAD_WIDTH"
Sheets(datasheet).Cells(1, MinRWCol) = "MIN_ROAD_WIDTH"

For j = 0 To 4
    row1 = 2
    row2 = 2
    StopBool = False
    
    If j = 0 Then
        'MAIN ROUTE DATA
        'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        MainRouteMPSort
        routecol = colRoute0
        MPcol = ColMP0
        
    ElseIf j = 1 Then
        'SECONDARY ROUTE DATA
        'Sort by secondary route/MP, main route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        SecondaryRouteMPSort
        routecol = colRoute1
        MPcol = ColMP1
        
   ElseIf j = 2 Then
        'TERTIARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP, quartenary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute2
        MPcol = ColMP2
        
    ElseIf j = 3 Then
        'QUARTENARY ROUTE DATA
        'Sort by quartenary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute3
        MPcol = ColMP3
    
    ElseIf j = 4 Then
        'QUINARY ROUTE DATA
        'Sort by quinary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP
        QuinaryRouteMPSort
        routecol = colRoute4
        MPcol = ColMP4
        
    End If
    
    

    Do While Sheets(datasheet).Cells(row1, 1) <> ""
        EnterData = False
        RouteNum = Sheets(datasheet).Cells(row1, routecol)
        MP = Sheets(datasheet).Cells(row1, MPcol)
        
        Do While Sheets(LanesSheet).Cells(row2, 1) <> ""
            If RouteNum = "Local" Then       'If it's Local, assume 0 lanes and 0ft width.
                Num_Lanes = 0
                LaneWidth = 0
                EnterData = True
                Exit Do
            ElseIf Int(RouteNum) > 491 Then       'If it's FedAid, assume 0 lanes and 0ft width.
                Num_Lanes = 0
                LaneWidth = 0
                EnterData = True
                Exit Do
            ElseIf Int(RouteNum) = Empty Then
                StopBool = True
                Exit Do
            ElseIf Int(Sheets(LanesSheet).Cells(row2, LanesColRoute)) > Int(RouteNum) Then     'Do if FC row route is higher than actual (meaning the intersection route doesn't exist)
                 'testing begin (by Camille)
                Dim myMsg As String
                'myMsg = "Route " & Int(RouteNum) & " does not exist in the Lanes file. Please take note. The code will continue upon clicking OK."
                'MsgBox myMsg, vbOKOnly, "Route Does Not Exist"
                'testing end
                Sheets(datasheet).Rows(row1).EntireRow.Delete     'commented out by Camille
                Exit Do
            ElseIf Int(Sheets(LanesSheet).Cells(row2, LanesColRoute)) = Int(RouteNum) And _
            Sheets(LanesSheet).Cells(row2, LanesColMP1).Value < MP And _
            Sheets(LanesSheet).Cells(row2, LanesColMP2).Value > MP Then     'Do if FC row has the same route and intersection is between milepoints
                Num_Lanes = Sheets(LanesSheet).Cells(row2, LanesCol)
                LaneWidth = Sheets(LanesSheet).Cells(row2, LaneWidthCol)
                EnterData = True
                Exit Do
            ElseIf Int(Sheets(LanesSheet).Cells(row2, LanesColRoute)) = Int(RouteNum) And _
            Sheets(LanesSheet).Cells(row2, LanesColMP1).Value = MP Then     'Do if FC row has the same route and intersection is at the beginning milepoint (MP = 0)
                Num_Lanes = Sheets(LanesSheet).Cells(row2, LanesCol)
                LaneWidth = Sheets(LanesSheet).Cells(row2, LaneWidthCol)
                EnterData = True
                Exit Do
            ElseIf Int(Sheets(LanesSheet).Cells(row2, LanesColRoute)) = Int(RouteNum) And _
            Sheets(LanesSheet).Cells(row2, LanesColMP2).Value = MP And _
            Sheets(LanesSheet).Cells(row2 + 1, LanesColRoute) = Int(RouteNum) Then        'Do if FC row has the same route, the intersection is at the ending MP, and the next FC segment is part of the same route as the previous.
                If Application.Max(Sheets(LanesSheet).Cells(row2, LanesCol), Sheets(LanesSheet).Cells(row2 + 1, LanesCol)) = Sheets(LanesSheet).Cells(row2, LanesCol) Then
                    Num_Lanes = Sheets(LanesSheet).Cells(row2, LanesCol)
                    LaneWidth = Sheets(LanesSheet).Cells(row2, LaneWidthCol)
                    EnterData = True
                Else
                    Num_Lanes = Sheets(LanesSheet).Cells(row2 + 1, LanesCol)
                    LaneWidth = Sheets(LanesSheet).Cells(row2 + 1, LaneWidthCol)
                    EnterData = True
                End If
                Exit Do
            ElseIf Int(Sheets(LanesSheet).Cells(row2, LanesColRoute)) = Int(RouteNum) And _
            Sheets(LanesSheet).Cells(row2, LanesColMP2).Value = MP Then     'Do if FC row has the same route and the interseciton is at the ending MP
                Num_Lanes = Sheets(LanesSheet).Cells(row2, LanesCol)
                LaneWidth = Sheets(LanesSheet).Cells(row2, LaneWidthCol)
                EnterData = True
                Exit Do
            End If
            row2 = row2 + 1
        Loop
        If EnterData = True Then
            If Num_Lanes > Sheets(datasheet).Cells(row1, MaxLanesCol) Or Sheets(datasheet).Cells(row1, MaxLanesCol) = "" Or _
            Sheets(datasheet).Cells(row1, MaxLanesCol) = " " Then
                Sheets(datasheet).Cells(row1, MaxLanesCol) = Num_Lanes
            End If
            If Num_Lanes < Sheets(datasheet).Cells(row1, MinLanesCol) Or Sheets(datasheet).Cells(row1, MinLanesCol) = "" Or _
            Sheets(datasheet).Cells(row1, MinLanesCol) = " " Then
                Sheets(datasheet).Cells(row1, MinLanesCol) = Num_Lanes
            End If
            
            If Num_Lanes * LaneWidth > Sheets(datasheet).Cells(row1, MaxRWCol) Or Sheets(datasheet).Cells(row1, MaxRWCol) = "" Or _
            Sheets(datasheet).Cells(row1, MaxRWCol) = " " Then
                Sheets(datasheet).Cells(row1, MaxRWCol) = (Num_Lanes * LaneWidth) + MaxAdjust
            End If
            If Num_Lanes * LaneWidth < Sheets(datasheet).Cells(row1, MinRWCol) Or Sheets(datasheet).Cells(row1, MinRWCol) = "" Or _
            Sheets(datasheet).Cells(row1, MinRWCol) = " " Then
                Sheets(datasheet).Cells(row1, MinRWCol) = (Num_Lanes * LaneWidth) + MinAdjust
            End If
            
        End If
        If StopBool = True Then
            Exit Do
        End If
        row1 = row1 + 1
    Loop
Next j

'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
MainRouteMPSort

End Sub


Sub UrbanCodeCalc()

'Declare variables
Dim row1, row2, i, j As Long
Dim routecol, MPcol As Integer
Dim RouteNum, colRoute0, ColMP0 As Integer          'Main route number, route column, and milepoint column
Dim colRoute1, ColMP1 As Integer                    'Secondary route column, and milepoint column
Dim colRoute2, ColMP2 As Integer                    'Tertiary route column, and milepoint column
Dim colRoute3, ColMP3 As Integer                    'Quartenary route column, and milepoint column
Dim colRoute4, ColMP4 As Integer                    'Quinary route column, and milepoint column

Dim UrbCodeCol, UrbDescCol As Integer
Dim UCRouteCol, UCMPCol1, UCMPCol2, UCUrbCodeCol, UCUrbDescCol As Integer       'UC route column, beginning milepoint column, and ending milepoint column

Dim MP As Double
Dim datasheet As String, UCSheet As String
Dim EnterData, StopBool As Boolean

'Assign initial values to variables

colRoute0 = 1
ColMP0 = 1
colRoute1 = 1
ColMP1 = 1
colRoute2 = 1
ColMP2 = 1
colRoute3 = 1
ColMP3 = 1
colRoute4 = 1
ColMP4 = 1

UCRouteCol = 1
UCMPCol1 = 1
UCMPCol2 = 1

UrbCodeCol = 1
UrbDescCol = 1
UCUrbCodeCol = 1
UCUrbDescCol = 1

datasheet = "Dataset"
UCSheet = "Urban_Code"

'Find columns
colRoute0 = FindColumn("ROUTE", datasheet, 1, True)
colRoute1 = FindColumn("INT_RT_1", datasheet, 1, True)
colRoute2 = FindColumn("INT_RT_2", datasheet, 1, True)
colRoute3 = FindColumn("INT_RT_3", datasheet, 1, True)
colRoute4 = FindColumn("INT_RT_4", datasheet, 1, True)
ColMP0 = FindColumn("UDOT_BMP", datasheet, 1, True)
ColMP1 = FindColumn("INT_RT_1_M", datasheet, 1, True)
ColMP2 = FindColumn("INT_RT_2_M", datasheet, 1, True)
ColMP3 = FindColumn("INT_RT_3_M", datasheet, 1, True)
ColMP4 = FindColumn("INT_RT_4_M", datasheet, 1, True)

Do Until Sheets(UCSheet).Cells(1, UCRouteCol) = "ROUTE_ID"
    UCRouteCol = UCRouteCol + 1
Loop

Do Until Sheets(UCSheet).Cells(1, UCMPCol1) = "BEG_MILEPOINT"
    UCMPCol1 = UCMPCol1 + 1
Loop

Do Until Sheets(UCSheet).Cells(1, UCMPCol2) = "END_MILEPOINT"
    UCMPCol2 = UCMPCol2 + 1
Loop

Do Until Sheets(UCSheet).Cells(1, UCUrbCodeCol) = "URBAN_CODE"
    UCUrbCodeCol = UCUrbCodeCol + 1
Loop

Do Until Sheets(UCSheet).Cells(1, UCUrbDescCol) = "URBAN_NAME"
    UCUrbDescCol = UCUrbDescCol + 1
Loop

'Add urban code headings
Do Until Sheets(datasheet).Cells(1, UrbCodeCol) = ""
    UrbCodeCol = UrbCodeCol + 1
Loop
UrbDescCol = UrbCodeCol + 1
Sheets(datasheet).Cells(1, UrbCodeCol) = "URBAN_CODE"
Sheets(datasheet).Cells(1, UrbDescCol) = "URBAN_DESC"


'For j = 0 To 4 commented out by Camille *crossing fingers* This should work. If need to uncommetn, make sure to uncomment the Next j statement as well
    j = 0       'this is a new line that is to replace the line above
    row1 = 2
    row2 = 2
    StopBool = False
    
    If j = 0 Then
        'MAIN ROUTE DATA
        'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        MainRouteMPSort
        routecol = colRoute0
        MPcol = ColMP0
        
    ElseIf j = 1 Then
        'SECONDARY ROUTE DATA
        'Sort by secondary route/MP, main route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        SecondaryRouteMPSort
        routecol = colRoute1
        MPcol = ColMP1
        
   ElseIf j = 2 Then
        'TERTIARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP, quartenary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute2
        MPcol = ColMP2
        
    ElseIf j = 3 Then
        'QUARTENARY ROUTE DATA
        'Sort by quartenary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute3
        MPcol = ColMP3
    
    ElseIf j = 4 Then
        'QUINARY ROUTE DATA
        'Sort by quinary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP
        QuinaryRouteMPSort
        routecol = colRoute4
        MPcol = ColMP4
        
    End If
    
    
    'Assign UC code and label/type values
    '''''don't we only need one value for this? Let's just use the SR (aka route 0)
    Do While Sheets(datasheet).Cells(row1, 1) <> ""
        
        Cells(row1, routecol).NumberFormat = "0000"
        Cells(row1, routecol).Value = Cells(row1, routecol).Value
        
        Cells(row2, 1).NumberFormat = "0000"
        Cells(row2, 1).Value = Cells(row2, 1).Value
        
        RouteNum = Sheets(datasheet).Cells(row1, routecol)
        MP = Sheets(datasheet).Cells(row1, MPcol)
        
        Do While Sheets(UCSheet).Cells(row2, 1) <> ""
            If j = 1 Then
                If Len(Sheets(UCSheet).Cells(row2, 1)) > 4 Then
                    Sheets(UCSheet).Cells(row2, 1).EntireRow.Delete
                End If
                Sheets(UCSheet).Cells(row2, 1).NumberFormat = "0000"
                Sheets(UCSheet).Cells(row2, 1).Value = Sheets(UCSheet).Cells(row2, 1).Value
            End If
            If Int(RouteNum) = Empty Then
                StopBool = True
                Exit Do
            ElseIf RouteNum = "Local" Then
                'do nothing   'added by Camille
            ElseIf Int(Sheets(UCSheet).Cells(row2, UCRouteCol)) > Int(RouteNum) Then     'Do if UC row route is higher than actual (meaning the intersection route doesn't exist)
                'testing begin (by Camille)
                'Dim myMsg As String
                'myMsg = "Route " & Int(RouteNum) & " does not exist in the Urban Code file. Please take note. The code will continue upon clicking OK."
                'MsgBox myMsg, vbOKOnly, "Route Does Not Exist"
                'testing end
                'Sheets(datasheet).Rows(row1).EntireRow.Delete       'commented out by Camille; Sub finishUC has been added to fix any blanks
                Exit Do
            ElseIf Int(Sheets(UCSheet).Cells(row2, UCRouteCol)) = Int(RouteNum) And _
            Sheets(UCSheet).Cells(row2, UCMPCol1).Value < MP And _
            Sheets(UCSheet).Cells(row2, UCMPCol2).Value > MP Then     'Do if UC row has the same route and intersection is between milepoints
                Sheets(datasheet).Cells(row1, UrbCodeCol) = Sheets(UCSheet).Cells(row2, UCUrbCodeCol)
                Sheets(datasheet).Cells(row1, UrbDescCol) = Sheets(UCSheet).Cells(row2, UCUrbDescCol)
                Exit Do
            ElseIf Int(Sheets(UCSheet).Cells(row2, UCRouteCol)) = Int(RouteNum) And _
            Sheets(UCSheet).Cells(row2, UCMPCol1).Value = MP Then     'Do if UC row has the same route and intersection is at the beginning milepoint (MP = 0)
                Sheets(datasheet).Cells(row1, UrbCodeCol) = Sheets(UCSheet).Cells(row2, UCUrbCodeCol)
                Sheets(datasheet).Cells(row1, UrbDescCol) = Sheets(UCSheet).Cells(row2, UCUrbDescCol)
                Exit Do
            ElseIf Int(Sheets(UCSheet).Cells(row2, UCRouteCol)) = Int(RouteNum) And _
            Sheets(UCSheet).Cells(row2, UCMPCol2).Value = MP And _
            Sheets(UCSheet).Cells(row2 + 1, UCRouteCol) = Int(RouteNum) Then        'Do if UC row has the same route, the intersection is at the ending MP, and the next UC segment is part of the same route as the previous.
                If Application.min(Sheets(UCSheet).Cells(row2, UCUrbCodeCol), Sheets(UCSheet).Cells(row2 + 1, UCUrbCodeCol)) = Sheets(UCSheet).Cells(row2, UCUrbCodeCol) Then
                    Sheets(datasheet).Cells(row1, UrbCodeCol) = Sheets(UCSheet).Cells(row2, UCUrbCodeCol)
                    Sheets(datasheet).Cells(row1, UrbDescCol) = Sheets(UCSheet).Cells(row2, UCUrbDescCol)
                Else
                    Sheets(datasheet).Cells(row1, UrbCodeCol) = Sheets(UCSheet).Cells(row2 + 1, UCUrbCodeCol)
                    Sheets(datasheet).Cells(row1, UrbDescCol) = Sheets(UCSheet).Cells(row2 + 1, UCUrbDescCol)
                End If
                Exit Do
            ElseIf Int(Sheets(UCSheet).Cells(row2, UCRouteCol)) = Int(RouteNum) And _
            Sheets(UCSheet).Cells(row2, UCMPCol2).Value = MP Then     'Do if UC row has the same route and the interseciton is at the ending MP
                Sheets(datasheet).Cells(row1, UrbCodeCol) = Sheets(UCSheet).Cells(row2, UCUrbCodeCol)
                Sheets(datasheet).Cells(row1, UrbDescCol) = Sheets(UCSheet).Cells(row2, UCUrbDescCol)
                Exit Do
            End If
            row2 = row2 + 1
        Loop
        If StopBool = True Then
            Exit Do
        End If
        row1 = row1 + 1
    Loop
'Next j


'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
MainRouteMPSort


End Sub

Sub PavMessCalc()

Dim row1, row2, col1 As Long
Dim MessTypeCol As Integer
Dim PavSheet As String
Dim datasheet As String
Dim StopBar, Arrow As String
Dim PavRouteCol, PavRoute, PavMPCol As Integer
Dim routecol, MPcol As Integer
Dim RouteNum, colRoute0, ColMP0 As Integer          'Main route number, route column, and milepoint column
Dim colRoute1, ColMP1 As Integer                    'Secondary route column, and milepoint column
Dim colRoute2, ColMP2 As Integer                    'Tertiary route column, and milepoint column
Dim colRoute3, ColMP3 As Integer                    'Quarternary route column, and milepoint column
Dim colRoute4, ColMP4 As Integer                    'Quinary route column, and milepoint column
Dim IntMP As Double
Dim StopBool As Boolean
Dim i, j As Integer
Dim IntDistCol, IntDistCol0, IntDistCol1, IntDistCol2, IntDistCol3, IntDistCol4 As Integer
Dim SBMP1, SBMP2, SBDiff As Double
Dim IntIntDist As Integer

'Add values to string variables
PavSheet = "Pavement_Messages"
datasheet = "Dataset"
StopBar = "STOP BAR"
Arrow = "ARROW"

'Find columns in pavement sheet
MessTypeCol = FindColumn("MESSAGE_TYPE", PavSheet, 1, True)
PavRouteCol = FindColumn("ROUTE_ID", PavSheet, 1, True)
PavMPCol = FindColumn("MILEPOINT", PavSheet, 1, True)

'Find volumns in data sheet
colRoute0 = FindColumn("ROUTE", datasheet, 1, True)
colRoute1 = FindColumn("INT_RT_1", datasheet, 1, True)
colRoute2 = FindColumn("INT_RT_2", datasheet, 1, True)
colRoute3 = FindColumn("INT_RT_3", datasheet, 1, True)
colRoute4 = FindColumn("INT_RT_4", datasheet, 1, True)
ColMP0 = FindColumn("UDOT_BMP", datasheet, 1, True)
ColMP1 = FindColumn("INT_RT_1_M", datasheet, 1, True)
ColMP2 = FindColumn("INT_RT_2_M", datasheet, 1, True)
ColMP3 = FindColumn("INT_RT_3_M", datasheet, 1, True)
ColMP4 = FindColumn("INT_RT_4_M", datasheet, 1, True)

'Find column extent of data sheet
col1 = Sheets(datasheet).Range("A1").End(xlToRight).Column
IntDistCol0 = col1 + 1
IntDistCol1 = IntDistCol0 + 1
IntDistCol2 = IntDistCol1 + 1
IntDistCol3 = IntDistCol2 + 1
IntDistCol4 = IntDistCol3 + 1
Sheets(datasheet).Cells(1, IntDistCol0) = "Int_Dist_0"
Sheets(datasheet).Cells(1, IntDistCol1) = "Int_Dist_1"
Sheets(datasheet).Cells(1, IntDistCol2) = "Int_Dist_2"
Sheets(datasheet).Cells(1, IntDistCol3) = "Int_Dist_3"
Sheets(datasheet).Cells(1, IntDistCol4) = "Int_Dist_4"


For j = 0 To 4
    row1 = 2
    row2 = 2
    StopBool = False
    
    If j = 0 Then
        'MAIN ROUTE DATA
        'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        MainRouteMPSort
        routecol = colRoute0
        MPcol = ColMP0
        IntDistCol = IntDistCol0
        
    ElseIf j = 1 Then
        'SECONDARY ROUTE DATA
        'Sort by secondary route/MP, main route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
        SecondaryRouteMPSort
        routecol = colRoute1
        MPcol = ColMP1
        IntDistCol = IntDistCol1
        
   ElseIf j = 2 Then
        'TERTIARY ROUTE DATA
        'Sort by tertiary route/MP, main route/MP, secondary route/MP, quartenary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute2
        MPcol = ColMP2
        IntDistCol = IntDistCol2
        
    ElseIf j = 3 Then
        'QUARTENARY ROUTE DATA
        'Sort by quartenary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quinary route/MP
        TertiaryRouteMPSort
        routecol = colRoute3
        MPcol = ColMP3
        IntDistCol = IntDistCol3
    
    ElseIf j = 4 Then
        'QUINARY ROUTE DATA
        'Sort by quinary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP
        QuinaryRouteMPSort
        routecol = colRoute4
        MPcol = ColMP4
        IntDistCol = IntDistCol4
        
    End If
        
    'Assign internal intersection distance to intersections      FLAGGED: there are a lot of do until "" loops which don't work when using R cleaned data because R turns "" into "NA"
    Do While Sheets(datasheet).Cells(row1, 1) <> ""
        RouteNum = Sheets(datasheet).Cells(row1, routecol)
        IntMP = Sheets(datasheet).Cells(row1, MPcol)
        row2 = 2
        
        If RouteNum = "Local" Then
            Sheets(datasheet).Cells(row1, IntDistCol) = 0
        ElseIf Int(RouteNum) = 0 Then
            Sheets(datasheet).Cells(row1, IntDistCol) = 0
        ElseIf Int(RouteNum) > 491 Then
            Sheets(datasheet).Cells(row1, IntDistCol) = 0 'added by Camille on June 4, 2019. This should make all but Route 0 go faster--the PavMess data sheet only has State Routes
        Else
            Do Until (Int(Sheets(PavSheet).Cells(row2, PavRouteCol)) = Int(RouteNum) _
            And Sheets(PavSheet).Cells(row2, PavMPCol) > IntMP _
            And (Sheets(PavSheet).Cells(row2, MessTypeCol) = StopBar Or _
            Sheets(PavSheet).Cells(row2, MessTypeCol) = Arrow)) _
            Or Sheets(PavSheet).Cells(row2, 1) = ""
                row2 = row2 + 1
            Loop
            If Sheets(PavSheet).Cells(row2, 1) = "" Then
                Sheets(datasheet).Cells(row1, IntDistCol) = 0
            Else
                SBMP2 = Sheets(PavSheet).Cells(row2, PavMPCol)
                If IntMP < 0.02 Then
                    SBDiff = (SBMP2 - IntMP) * 2
                ElseIf Sheets(PavSheet).Cells(row2, MessTypeCol) = StopBar Then
                    row2 = row2 - 1
                    Do Until Sheets(PavSheet).Cells(row2, MessTypeCol) = StopBar
                        row2 = row2 - 1
                    Loop
                    SBMP1 = Sheets(PavSheet).Cells(row2, PavMPCol)
                    SBDiff = SBMP2 - SBMP1
                ElseIf Sheets(PavSheet).Cells(row2, MessTypeCol) = Arrow Then
                    row2 = row2 - 1
                    Do Until Sheets(PavSheet).Cells(row2, MessTypeCol) = Arrow
                        row2 = row2 - 1
                    Loop
                    SBMP1 = Sheets(PavSheet).Cells(row2, PavMPCol)
                    SBDiff = SBMP2 - SBMP1 - (30 / 5280)                   'Subtract 30 to account for distance from arrow to stop bar
                End If
                
                If SBDiff > 0.007 And SBDiff < 0.05 Then 'if the distance is between 37ft and 264ft then the value is reasonable
                    IntIntDist = Round(SBDiff * 5280, 0)
                    Sheets(datasheet).Cells(row1, IntDistCol) = IntIntDist
                Else
                    Sheets(datasheet).Cells(row1, IntDistCol) = 0
                End If
            End If
        End If
        row1 = row1 + 1
    Loop
Next j


'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
MainRouteMPSort


End Sub

Sub CorrectCity()

Dim row1, col1 As Long

row1 = 2
col1 = 1

Do Until Sheets("Dataset").Cells(1, col1) = "CITY"
    col1 = col1 + 1
Loop

Do Until Sheets("Dataset").Cells(row1, col1) = ""
    Sheets("Dataset").Cells(row1, col1) = Right(Sheets("Dataset").Cells(row1, col1), Len(Sheets("Dataset").Cells(row1, col1).Value) - InStr(1, Sheets("Dataset").Cells(row1, col1), "-") - 1)
    row1 = row1 + 1
Loop

End Sub
Sub finishUC()

'goes though the Dataset sheet and finds rows with empty or unknown urban code
'matches the city with the list in Column CL of OtherData sheet
'OtherData Sheet columns CL through CU are for this sub (and are hard coded, but could be altered to make it soft coded)

Dim codecol As Integer
Dim textcol As Integer
Dim citycol As Integer
Dim latcol As Integer
Dim longcol As Integer
Dim finderrow As Long
Dim refrow As Integer
Dim city As String
Dim lat As String
Dim lng As String
Dim UCstatus As String

Application.ScreenUpdating = False

'finds columns
codecol = 1
Do Until Sheets("Dataset").Cells(1, codecol) = "URBAN_CODE"
    codecol = codecol + 1
Loop
textcol = 1
Do Until Sheets("Dataset").Cells(1, textcol) = "URBAN_DESC"
    textcol = textcol + 1
Loop
citycol = 1
Do Until Sheets("Dataset").Cells(1, citycol) = "CITY"
    citycol = citycol + 1
Loop
latcol = 1
Do Until Sheets("Dataset").Cells(1, latcol) = "LATITUDE"
    latcol = latcol + 1
Loop
longcol = 1
Do Until Sheets("Dataset").Cells(1, longcol) = "LONGITUDE"
    longcol = longcol + 1
Loop

finderrow = 2

'begin search loop
Do While Sheets("Dataset").Cells(finderrow, 1) <> ""
    If Sheets("Dataset").Cells(finderrow, codecol) = 99999 Then   'if its rural (and says "Rurall")
        Sheets("Dataset").Cells(finderrow, textcol) = "Rural"
    End If
    If Sheets("Dataset").Cells(finderrow, textcol) = "" Or Sheets("Dataset").Cells(finderrow, textcol) = "Unknown" Then
        refrow = 4
        Do Until refrow = 1000   'begin match loop
            If Sheets("Dataset").Cells(finderrow, citycol) = Sheets("OtherData").Cells(refrow, 90) Then
                Sheets("Dataset").Cells(finderrow, codecol) = Sheets("OtherData").Cells(refrow, 91)  'paste code
                Sheets("Dataset").Cells(finderrow, textcol) = Sheets("OtherData").Cells(refrow, 92)  'paste description
                Exit Do
            End If
            If Sheets("OtherData").Cells(refrow, 90) = "" Then   'no matching city found
                Sheets("OtherData").Cells(4, 99) = Sheets("Dataset").Cells(finderrow, textcol)  'paste uc status for reference
                Sheets("OtherData").Cells(5, 99) = Sheets("Dataset").Cells(finderrow, citycol) 'paste city name for reference
                Sheets("OtherData").Cells(6, 99) = Sheets("Dataset").Cells(finderrow, latcol) 'paste lat for reference
                Sheets("OtherData").Cells(7, 99) = Sheets("Dataset").Cells(finderrow, longcol) 'paste long for reference
                Application.ScreenUpdating = True
                formFixUC.Show  'asks for user to select urban area, pastes answer onto OtherData Sheet
                Application.ScreenUpdating = False
                refrow = Sheets("OtherData").Cells(8, 99)
                Sheets("Dataset").Cells(finderrow, codecol) = Sheets("OtherData").Cells(refrow, 95)  'paste code
                Sheets("Dataset").Cells(finderrow, textcol) = Sheets("OtherData").Cells(refrow, 96)  'paste description
                'clear cells
                Sheets("OtherData").Cells(4, 99) = ""
                Sheets("OtherData").Cells(5, 99) = ""
                Sheets("OtherData").Cells(6, 99) = ""
                Sheets("OtherData").Cells(7, 99) = ""
                Sheets("OtherData").Cells(8, 99) = ""
                Exit Do
            End If
            refrow = refrow + 1
        Loop   'match loop      'exits at an exit do
    End If
    finderrow = finderrow + 1
Loop    'search loop        'exits when all dataset rows have been checked






End Sub


Sub ColAdjust()

'new code for adjusting columns Tommy 6/20/19

Dim UDOTCol, IntRT1Col As Long

UDOTCol = 1
IntRT1Col = 1

Do Until Sheets("Dataset").Cells(1, UDOTCol) = "UDOT_BMP"
    UDOTCol = UDOTCol + 1
Loop

Do Until Sheets("Dataset").Cells(1, IntRT1Col) = "INT_RT_1_M"
    IntRT1Col = IntRT1Col + 1
Loop

Worksheets("Dataset").Columns(IntRT1Col).Insert shift:=xlToRight
Worksheets("Dataset").Columns(UDOTCol + 1).Copy
Worksheets("Dataset").Columns(IntRT1Col).Insert
Worksheets("Dataset").Columns(IntRT1Col + 1).Delete
Worksheets("Dataset").Columns(UDOTCol + 1).Delete


'old code for adjusting columns
'Dim CityCol, CountyCol, RegionCol As Long
'Dim IntMP3Col As Long

'CityCol = 1
'CountyCol = 1
'RegionCol = 1
'IntMP3Col = 1

'Do Until Sheets("Dataset").Cells(1, IntMP3Col) = "INT_MP_3"
'    IntMP3Col = IntMP3Col + 1
'Loop


'Do Until Sheets("Dataset").Cells(1, RegionCol) = "REGION"
'    RegionCol = RegionCol + 1
'Loop

'If CityCol <> IntMP3Col + 1 Then
'    Worksheets("Dataset").Columns(RegionCol).Select
'    Selection.Cut
'    Worksheets("Dataset").Columns(IntMP3Col + 1).Select
'    Selection.Insert shift:=xlToRight
'End If

'Do Until Sheets("Dataset").Cells(1, CountyCol) = "COUNTY"
'    CountyCol = CountyCol + 1
'Loop

'If CityCol <> IntMP3Col + 2 Then
'    Worksheets("Dataset").Columns(CountyCol).Select
'    Selection.Cut
'    Worksheets("Dataset").Columns(IntMP3Col + 2).Select
'    Selection.Insert shift:=xlToRight
'End If

'Do Until Sheets("Dataset").Cells(1, CityCol) = "CITY"
'    CityCol = CityCol + 1
'Loop

'If CityCol <> IntMP3Col + 3 Then
'    Worksheets("Dataset").Columns(CityCol).Select
'    Selection.Cut
'    Worksheets("Dataset").Columns(IntMP3Col + 3).Select
'    Selection.Insert shift:=xlToRight
'End If

'Worksheets("Dataset").Cells(1, 1).Select

End Sub

Function FindColumn(colName As String, wksht As String, HeadRow As Integer, Necessary As Boolean)
    Dim col1 As Integer
    Dim continue As Integer
    col1 = 1
    Do Until Left(Worksheets(wksht).Cells(HeadRow, col1), Len(colName)) = colName
        col1 = col1 + 1
        If col1 > 200 Then
            If Necessary = True Then
                continue = MsgBox("The column heading " & colName & " was not found on the " & wksht & " worksheet. Continuing without this data could result in errors. Do you wish to continue anyways?", vbYesNo, "Warning")
                If continue = vbNo Then
                    MsgBox "The column heading " & colName & " was not found on the " & wksht & " worksheet. Macros terminated.", , "Heading Not Found"
                    End
                Else
                    col1 = 0
                End If
            Else
                col1 = 0
            End If
        End If
    Loop
    FindColumn = col1
End Function

Public Function Col_Letter(lngCol) As String

Dim vArr

vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)
End Function
Sub CleanUpIntFile()

'performed on the Dataset [formerly Intersections] sheet
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o
Dim ObjectID, Int_ID, Group, Group_Num, Travel_Dir, Mandli_BMP, Int_Type, Skew_Prese, Rail_Prese, Station, Admin_Unit, Beg_Elev, Comments, Mandli_ID, Coll_Date

a = 1 'objectID is deleted
b = 1 'Intersection ID
c = 1 'Group name
d = 1 'Group number
e = 1 'Travel DIrection is deleted
f = 1 'Mandli BMP is deleted
g = 1 'Intersection type 'Camille: I don't think we should delete this one. It would be helpful to have on the ISAR
h = 1 'Skew Present 'Camille: I have mixed feelings about deleting this one as well. I'll keep it deleted for now
i = 1 'Rail Present is deleted
j = 1 'Station
k = 1 'Admin Unit is deleted
l = 1 'Intersection Elevation is deleted 'Camille: I have mixed feelings about deleting this one as well. I'll keep it deleted for now
m = 1 'Comments is deleted
n = 1 'Mandli ID is deleted
o = 1 'Date Collected is deleted


Do Until Cells(1, a) = ""
    If Cells(1, a) = "OBJECTID" Then
        Cells(1, a).Activate
        ObjectID = ActiveCell.Address(False, False)
        Exit Do
    End If
    a = a + 1
Loop

If ObjectID <> "" Then
    Range(ObjectID).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, e) = ""
    If Cells(1, e) = "TRAVEL_DIR" Then
        Cells(1, e).Activate
        Travel_Dir = ActiveCell.Address(False, False)
        Exit Do
    End If
    e = e + 1
Loop

If Travel_Dir <> "" Then
    Range(Travel_Dir).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, f) = ""
    If Cells(1, f) = "MANDLI_BMP" Then
        Cells(1, f).Activate
        Mandli_BMP = ActiveCell.Address(False, False)
        Exit Do
    End If
    f = f + 1
Loop

If Mandli_BMP <> "" Then
    Range(Mandli_BMP).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, g) = ""
    If Cells(1, g) = "INT_TYPE" Then
        Cells(1, g).Activate
        Int_Type = ActiveCell.Address(False, False)
        Exit Do
    End If
    g = g + 1
Loop

'If Int_Type <> "" Then
'    Range(Int_Type).Activate
'    ActiveCell.EntireColumn.Delete
'End If

Do Until Cells(1, h) = ""
    If Cells(1, h) = "SKEW_PRESE" Then
        Cells(1, h).Activate
        Skew_Prese = ActiveCell.Address(False, False)
        Exit Do
    End If
    h = h + 1
Loop

If Skew_Prese <> "" Then
    Range(Skew_Prese).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, i) = ""
    If Cells(1, i) = "RAIL_PRESE" Then
        Cells(1, i).Activate
        Rail_Prese = ActiveCell.Address(False, False)
        Exit Do
    End If
    i = i + 1
Loop

If Rail_Prese <> "" Then
    Range(Rail_Prese).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, k) = ""
    If Cells(1, k) = "ADMIN_UNIT" Then
        Cells(1, k).Activate
        Admin_Unit = ActiveCell.Address(False, False)
        Exit Do
    End If
    k = k + 1
Loop

If Admin_Unit <> "" Then
    Range(Admin_Unit).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, l) = ""
    If Cells(1, l) = "BEG_ELEV" Then
        Cells(1, l).Activate
        Beg_Elev = ActiveCell.Address(False, False)
        Exit Do
    End If
    l = l + 1
Loop

If Beg_Elev <> "" Then
    Range(Beg_Elev).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, m) = ""
    If Cells(1, m) = "COMMENTS" Then
        Cells(1, m).Activate
        Comments = ActiveCell.Address(False, False)
        Exit Do
    End If
    m = m + 1
Loop

If Comments <> "" Then
    Range(Comments).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, n) = ""
    If Cells(1, n) = "MANDLI_ID" Then
        Cells(1, n).Activate
        Mandli_ID = ActiveCell.Address(False, False)
        Exit Do
    End If
    n = n + 1
Loop

If Mandli_ID <> "" Then
    Range(Mandli_ID).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, o) = ""
    If Cells(1, o) = "COLL_DATE" Then
        Cells(1, o).Activate
        Coll_Date = ActiveCell.Address(False, False)
        Exit Do
    End If
    o = o + 1
Loop

If Coll_Date <> "" Then
    Range(Coll_Date).Activate
    ActiveCell.EntireColumn.Delete
End If

Do Until Cells(1, b) = ""
    If Cells(1, b) = "Int_ID" Then
        Cells(1, b).Activate
        Int_ID = ActiveCell.Address(False, False)
        Exit Do
    End If
    b = b + 1
Loop

Do Until Cells(1, c) = ""
    If Cells(1, c) = "Group" Then
        Cells(1, c).Activate
        Group = ActiveCell.Address(False, False)
        Exit Do
    End If
    c = c + 1
Loop

Do Until Cells(1, d) = ""
    If Cells(1, d) = "Group_Num" Then
        Cells(1, d).Activate
        Group_Num = ActiveCell.Address(False, False)
        Exit Do
    End If
    d = d + 1
Loop

Do Until Cells(1, j) = ""
    If Cells(1, j) = "STATION" Then
        Cells(1, j).Activate
        Station = ActiveCell.Address(False, False)
        Exit Do
    End If
    j = j + 1
Loop


Cells(1, 1).Value = "ORIGINAL_INT_ID"
End Sub
