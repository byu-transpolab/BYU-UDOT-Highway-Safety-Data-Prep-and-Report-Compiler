Attribute VB_Name = "DP_07_Crash"
Option Explicit

'PMDP_07_Crash Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public LocCrashID, LocRouDir, LocRoute, DatCrashID, RollCrashID, VehDetID, LocRampID As String      'Column numbers for specific headers in crash data
Public VehCrashID, VehVehNum, VehCraDate, VehRavDir, VehEvenSeq, VehMostHarm, VehManID As String    'Column numbers for specific headers in crash data
Public CurrentWkbk As String                'String variable that holds current workbook name for reference purposes
Public numCrash, num10000Crash, numTotCrash As Long        'Number of crashes to show in progress boxes

Sub Vehicle_Data_Prep()
'Vehicle_Data_Prep macro:
'   (1) Copy the vehicle data to a new sheet and prepare it for future analysis.
'
' Created by: Samuel Mineer, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Samuel Mineer and Josh Gibbons, BYU, 2016

'Define Variables
Dim icrash As Integer           'Column counter for Crash_ID
Dim iVehNum As Integer          'Column counter for Vehicle Number
Dim idatetime As Integer        'Column counter for datetime
Dim ievent1 As Integer          'Column counter for Event Sequence 1
Dim iharm As Integer            'Column counter for Most Harmful Event
Dim imaneu As Integer           'Column counter for Vehicle Maneuver
Dim numrow As Double            'Row counter
Dim numcol As Double            'Column counter
Dim MyPath As String
Dim MyFileName As String
Dim rowYear As Integer
Dim minyear As Integer
Dim maxyear As Integer
Dim iCol As Integer
Dim strpath, iFileName As String

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'Re-sort Vehicle data by crash_id and then vehicle detail id
Dim row1, row2, col1, col2 As Integer           'row and column counters
Dim nrow, ncol, icol1, icol2 As Double          'row and column counters

'Assign initial values
icol1 = 1
icol2 = 1

'Find extents of data
nrow = Sheets("Vehicle").Range("A1").End(xlDown).row + 1
ncol = Sheets("Vehicle").Range("A1").End(xlToRight).Column + 1

'Find columns for CRASH_ID and VEHICLE_DETAIL_ID
Do Until Sheets("Vehicle").Cells(1, icol1) = "CRASH_ID"
    icol1 = icol1 + 1
Loop
'Do Until Sheets("Vehicle").Cells(1, icol2) = "VEHICLE_DETAIL_ID"
'    icol2 = icol2 + 1
'Loop

'Sort Vehicle data by crash_id and then vehicle detail id
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, icol1), Cells(nrow, icol1)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, icol2), Cells(nrow, icol2)), _
'    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Vehicle").Sort
    .SetRange Range(Cells(1, 1), Cells(nrow, ncol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Find columns with certain column headers and assign column number to variable
imaneu = 1
iharm = 1
ievent1 = 1
icrash = 1
iVehNum = 1
idatetime = 1
Do Until Cells(1, imaneu) = "EVENT_SEQUENCE_1_ID"
    imaneu = imaneu + 1
Loop
Do Until Cells(1, iharm) = "MOST_HARMFUL_EVENT_ID"
    iharm = iharm + 1
Loop
Do Until Cells(1, ievent1) = "VEHICLE_MANEUVER_ID"
    ievent1 = ievent1 + 1
Loop
Do Until Cells(1, icrash) = "CRASH_ID"
    icrash = icrash + 1
Loop
Do Until Cells(1, iVehNum) = "VEHICLE_NUM"
    iVehNum = iVehNum + 1
Loop
Do Until Cells(1, idatetime) = "CRASH_DATETIME"
    idatetime = idatetime + 1
Loop

'Formats the first row and column widths
'Rows("1:1").Select
'With Selection
'    .HorizontalAlignment = xlLeft
'    .VerticalAlignment = xlBottom
'    .WrapText = False
'    .Orientation = 90
'    .AddIndent = False
'    .IndentLevel = 0
'    .ShrinkToFit = False
'    .ReadingOrder = xlContext
'    .MergeCells = False
'End With
'Cells.Select
'Cells.EntireColumn.autofit

'Find max and min years
numrow = 2
Do Until Cells(numrow, 1) = ""
    rowYear = year(Cells(numrow, idatetime).Value)
    If minyear = 0 And maxyear = 0 Then
        minyear = rowYear
        maxyear = rowYear
    ElseIf rowYear < minyear Then
        minyear = rowYear
    ElseIf rowYear > maxyear Then
        maxyear = rowYear
    End If
    numrow = numrow + 1
Loop

'Assign cells AU5 and AU6 to min and max years found in last step
Sheets("OtherData").Range("AU5").Value = minyear
Sheets("OtherData").Range("AU6").Value = maxyear

'Find column extent
numcol = Range("A1").End(xlToRight).Column + 1

'Sort the data by Crash_ID and Vehicle_Num
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, icrash), Cells(numrow, icrash)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, iVehNum), Cells(numrow, iVehNum)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Vehicle").Sort
    .SetRange Range(Cells(1, 1), Cells(numrow, numcol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Select cell A1 to deselect others (to not slow down macros)
'Range("A1").Select

'Highlight the correct Header green
'Union(Columns(icrash), Columns(imaneu), Columns(iharm), Columns(ievent1)).Select
'With Selection.Interior
'    .Pattern = xlSolid
'    .PatternColorIndex = xlAutomatic
'    .Color = 5296274        'Green
'    .TintAndShade = 0
'    .PatternTintAndShade = 0
'End With

'Activate Progress sheet
Workbooks(CurrentWkbk).Activate
Sheets("Progress").Activate

End Sub

Sub DatabaseCleanup()
'DatabaseCleanup macro:
'   (1) Crash data is joined based on Crash_ID.
'   (2) First vehicle direction is determined from vehicle crash data.
'   (3) Route numbers and direction values fixed and labels created.
'
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define variables
Dim guiwb As String

'Screen Updating OFF
Application.ScreenUpdating = False

'Begin to show progress box on Home sheet
'Update progress screen
guiwb = ActiveWorkbook.Name
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Combining Crash Files. Please wait."
    .Range("A3") = "Do not close Excel. Code running."
    .Range("B2") = ""
    .Range("B3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
    .Range("B6") = ""
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

'Run Join Database macro:
'   (1) Joins data from the the Crash data and rollup ID to the Location data, based on the Crash_ID.
Join_Database

'Run Fix_Routes macro:
'   (1) Fix the Route Number, into 4 digit codes
'   (2) Move Non-State Route Crashes to another sheet
'   (3) Move Ramp Crashes to another sheet
Fix_Routes

'Run Direction Code macro:
'   (1) Adds a column for Directional Code, based on vehicle direction
'Direction_Code

'Run Create Label macro:
'   (1) Creates a column for Label Data and fills in the column
Create_Label

'Screen Updating ON
Sheets("Progress").Activate
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub Join_Database()
'Join_Database macro:
'   (1) Joins data from the the Crash data and Location data to the Rollup data, based on the Crash_ID
'
' Created by: Samuel Mineer, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Samuel Mineer and Josh Gibbons, BYU, 2016

'Declare Variables
Dim locID As String                                         'String variable for CrashID on Location sheet
Dim CrashID As String                                       'String variable for CrashID on Crash sheet
Dim rollID As String                                        'String variable for CrashID on Rollup sheet
Dim vehID As String                                         'String variable for CrashID on Vehicle sheet
Dim iLocCID, iCrashDCID, iRollCID, iVehCID As Integer       'Column counter for crash IDs
Dim iVehNum, iVehDir, idir As Integer                       'Column counters for number of vehicles and vehicle direction
Dim iFHarm, iManCol As Integer
Dim row1, row2, row3, row4 As Long                          'Row counter, for row being cycled through
Dim errorrow As Long                                        'Row counter, for rows on error page
Dim erow As Long                                            'Value which the loop will restart; in other words, the macro will work through 20,000 lines at a time, to prevent a crash
Dim iCrashCol, iLocCol, iRollCol, iVehCol As Integer        'Values to find the column extents
Dim iCrashRow, iLocRow, iRollRow, iVehRow As Long           'Values to find the row extent
Dim mact As Integer                                         'Value for user's response to MessageBox before code starts
Dim i, j As Long                                            'Variables for For Next loops
Dim routecol As Integer                                     'variable for Route_ID column on the Rollups sheet (after the locaiton data has been pasted on to the rollups sheet)


'Screen Updating OFF
Application.ScreenUpdating = False

'Find First vehicle direction column
idir = 1
Do While Sheets("Location").Cells(1, idir) <> "DIRECTION"
    idir = idir + 1
Loop
'Sheets("Location").Columns(idir + 1).Insert Shift:=xlShiftToRight
'Sheets("Location").Cells(1, idir + 1) = "First_Veh_Direction"

'Identify the number of vehicles and vehicle direction columns in the vehicles dataset
iVehNum = 1
Do Until Sheets("Vehicle").Cells(1, iVehNum) = "VEHICLE_NUM"
    iVehNum = iVehNum + 1
Loop
iVehDir = 1
Do While Sheets("Vehicle").Cells(1, iVehDir) <> "TRAVEL_DIRECTION_ID"
    iVehDir = iVehDir + 1
Loop

'Find additional vehicle data columns
iFHarm = 1
Do Until Sheets("Crash").Cells(1, iFHarm) = "FIRST_HARMFUL_EVENT_ID"
    iFHarm = iFHarm + 1
Loop
iManCol = 1
Do While Sheets("Crash").Cells(1, iManCol) <> "MANNER_COLLISION_ID"
    iManCol = iManCol + 1
Loop

'Identify the Crash ID columns in each dataset
iLocCID = 1
Do Until Sheets("Location").Cells(1, iLocCID) = "CRASH_ID"
    iLocCID = iLocCID + 1
Loop
iCrashDCID = 1
Do Until Sheets("Crash").Cells(1, iCrashDCID) = "CRASH_ID"
    iCrashDCID = iCrashDCID + 1
Loop
iRollCID = 1
Do Until Sheets("Rollup").Cells(1, iRollCID) = "CRASH_ID"
    iRollCID = iRollCID + 1
Loop
iVehCID = 1
Do Until Sheets("Vehicle").Cells(1, iVehCID) = "CRASH_ID"
    iVehCID = 1
Loop

'Identify the column extent of the Crash, Location, Rollup, and Vehicle data
iCrashCol = Sheets("Crash").Range("A1").End(xlToRight).Column
iLocCol = Sheets("Location").Range("A1").End(xlToRight).Column
iRollCol = Sheets("Rollup").Range("A1").End(xlToRight).Column
iVehCol = Sheets("Vehicle").Range("A1").End(xlToRight).Column

'Identify the row extent of the Crash, Location, Rollup, and Vehicle data
iCrashRow = Sheets("Crash").Range("A1").End(xlDown).row
iLocRow = Sheets("Location").Range("A1").End(xlDown).row
iRollRow = Sheets("Rollup").Range("A1").End(xlDown).row
iVehRow = Sheets("Vehicle").Range("A1").End(xlDown).row

'Set total number of crashes to numTotCrash and 'Reset progress values
numTotCrash = iRollRow
num10000Crash = 0
numCrash = 1

'Re-sort each crash dataset
'Sort location data by Crash_ID
ActiveWorkbook.Worksheets("Location").Activate
ActiveWorkbook.Worksheets("Location").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Location").Sort.SortFields.Add Key:=Range(Cells(2, iLocCID), Cells(iLocRow, iLocCID)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Location").Sort
        .SetRange Range(Cells(1, 1), Cells(iLocRow, iLocCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort crash data by Crash_ID
ActiveWorkbook.Worksheets("Crash").Activate
ActiveWorkbook.Worksheets("Crash").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Crash").Sort.SortFields.Add Key:=Range(Cells(2, iCrashDCID), Cells(iCrashRow, iCrashDCID)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
        For i = 1 To 100
        If ActiveWorkbook.Worksheets("Crash").Cells(1, i) = "CRASH_DATETIME" Then
            ActiveWorkbook.Worksheets("Crash").Cells(1, i).EntireColumn.Delete
        End If
        Next i
    End If
    iCrashRow = Sheets("Crash").Range("A1").End(xlDown).row
    With ActiveWorkbook.Worksheets("Crash").Sort
        .SetRange Range(Cells(1, 1), Cells(iCrashRow, iCrashCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort rollup data by Crash_ID
ActiveWorkbook.Worksheets("Rollup").Activate
ActiveWorkbook.Worksheets("Rollup").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rollup").Sort.SortFields.Add Key:=Range(Cells(2, iRollCID), Cells(iRollRow, iRollCID)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    iRollRow = Sheets("Rollup").Range("A1").End(xlDown).row
    With ActiveWorkbook.Worksheets("Rollup").Sort
        .SetRange Range(Cells(1, 1), Cells(iRollRow, iRollCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort vehicle data by Crash_ID and Vehicle Num
ActiveWorkbook.Worksheets("Vehicle").Activate
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, iVehCID), Cells(iVehRow, iVehCID)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, iVehNum), Cells(iVehRow, iVehNum)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
        For i = 1 To 100
        If ActiveWorkbook.Worksheets("Vehicle").Cells(1, i) = "CRASH_DATETIME" Then
            ActiveWorkbook.Worksheets("Vehicle").Cells(1, i).EntireColumn.Delete
        End If
        Next i
    End If
    iVehRow = Sheets("Vehicle").Range("A1").End(xlDown).row
    With ActiveWorkbook.Worksheets("Vehicle").Sort
        .SetRange Range(Cells(1, 1), Cells(iVehRow, iVehCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



ActiveWorkbook.Worksheets("Home").Activate

'Copy header values to rollup sheet
For i = 2 To iCrashCol
    Sheets("Rollup").Cells(1, iRollCol - 1 + i) = Sheets("Crash").Cells(1, i)
Next i
For i = 2 To iLocCol
    Sheets("Rollup").Cells(1, iRollCol + iCrashCol - 2 + i) = Sheets("Location").Cells(1, i)
Next i
For i = 2 To iVehCol
    Sheets("Rollup").Cells(1, iRollCol + iCrashCol + iLocCol - 3 + i) = Sheets("Vehicle").Cells(1, i)
Next i

'Combine crash, location, vehicle, and rollup data into a single database based on Crash ID
row1 = 2                'Beginning row to begin cycling through
row2 = 2
row3 = 2
row4 = 2
'Cycle through crashes on rollup sheet
Do While Sheets("Rollup").Cells(row1, iRollCID) <> ""
    'Assign Crash ID values from each dataset
    locID = Sheets("Location").Cells(row3, iLocCID)
    CrashID = Sheets("Crash").Cells(row2, iCrashDCID)
    rollID = Sheets("Rollup").Cells(row1, iRollCID)
    vehID = Sheets("Vehicle").Cells(row4, iVehCID)

    If rollID = 11038647 Then
        rollID = rollID  'check to see what happens to rollID #11038653
    End If
    
    
    
    'If all other crash IDs are higher than the rollup ID, then add go to next rollup ID row
    If rollID < CrashID And rollID < locID And rollID < vehID Then
        row1 = row1 + 1
        rollID = Sheets("Rollup").Cells(row1, iRollCID)
        
        'Update progress
        numCrash = row1
        num10000Crash = num10000Crash + 1
        If num10000Crash = 10000 Or numCrash = numTotCrash Then
            Sheets("Progress").Activate
            
            'Screen Updating ON
            Application.ScreenUpdating = True
            
            Range("A2") = "Join Database (1 of 3)"
            Range("A3") = CStr(numCrash) & " of " & CStr(numTotCrash) & " crashes"
            Range("C3") = numCrash / numTotCrash
            Range("A5") = "Update Time"
            Range("B5") = Time
            
            Application.Wait (Now + TimeValue("0:00:01"))
            
            'Screen Updating OFF
            Application.ScreenUpdating = False
            
            num10000Crash = 0
        End If
        
    'If all other crash IDs are blank, then move to next rollup ID row
        ElseIf CrashID = "" And locID = "" Then
        row1 = row1 + 1
        
        'Update progress
        numCrash = row1
        num10000Crash = num10000Crash + 1
        If num10000Crash = 10000 Or numCrash = numTotCrash Then
            Sheets("Progress").Activate
            
            'Screen Updating ON
            Application.ScreenUpdating = True
            
            Range("A2") = "Join Database (1 of 3)"
            Range("A3") = CStr(numCrash) & " of " & CStr(numTotCrash) & " crashes"
            Range("C3") = numCrash / numTotCrash
            Range("A5") = "Update Time"
            Range("B5") = Time
            
            Application.Wait (Now + TimeValue("0:00:02"))
            
            'Screen Updating OFF
            Application.ScreenUpdating = False
            
            num10000Crash = 0
        End If
        
    End If

    'If rollup and crash IDs are equal, then copy crash data over to rollup sheet
    If rollID = CrashID Then
        For i = 2 To iCrashCol
            Sheets("Rollup").Cells(row1, iRollCol - 1 + i) = Sheets("Crash").Cells(row2, i)  'if code breaks change row2 to row3!!!!!!!!!!!!!!!!!!!!!!
        Next i
        row2 = row2 + 1
    'If crash ID is less than rollup ID, go to next crash ID
    ElseIf CrashID < rollID Then
        row2 = row2 + 1
    'Stay at same crash ID if greater than rollup ID
    ElseIf CrashID > rollID Then
        'Nothing
    End If
    
    'If rollup and location IDs are equal, then copy location data over to rollup sheet
    If rollID = locID Then
        For i = 2 To iLocCol
            Sheets("Rollup").Cells(row1, iRollCol + iCrashCol - 2 + i) = Sheets("Location").Cells(row3, i)
        Next i
        row3 = row3 + 1
    'If location ID is less than rollup ID, go to next location ID
    ElseIf locID < rollID Then
        row3 = row3 + 1
    'Stay at same location ID if greater than rollup ID
    ElseIf locID > rollID Then
        'Nothing
    End If
    
    'If rollup and vehicle IDs are equal, then copy vehicle data over to rollup sheet
    If rollID = vehID Then
        For i = 2 To iVehCol
            Sheets("Rollup").Cells(row1, iRollCol + iCrashCol + iLocCol - 3 + i) = Sheets("Vehicle").Cells(row4, i)
        Next i
        row4 = row4 + 1
    'If vehicle ID is less than rollup ID, go to next rollup ID
    ElseIf vehID < rollID Then
        row4 = row4 + 1
    'Stay at same vehicle ID if greater than rollup ID
    ElseIf vehID > rollID Then
        'Nothing
    End If
    
Loop


If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
    routecol = 1
    Do Until Sheets("Rollup").Cells(1, routecol) = "ROUTE_ID"
        routecol = routecol + 1
    Loop
    row1 = 2
    Do Until Sheets("Rollup").Cells(row1, iRollCID) = ""
        If Sheets("Rollup").Cells(row1, routecol) = "" Then
            Sheets("Rollup").Cells(row1, routecol).EntireRow.Delete
            row1 = row1 - 1
        End If
        row1 = row1 + 1
    Loop
End If


'Adjust Columns
Dim Col_ID, Col_Date, Blank_Col As Integer
Col_ID = 1
Col_Date = 1
Blank_Col = 1
Do Until Worksheets("Rollup").Cells(1, Col_ID) = "CRASH_ID"
    Col_ID = Col_ID + 1
Loop
Do Until Worksheets("Rollup").Cells(1, Col_Date) = "CRASH_DATETIME"
    Col_Date = Col_Date + 1
Loop
Worksheets("Rollup").Activate
Worksheets("Rollup").Columns(Col_Date).Copy
Worksheets("Rollup").Columns(Col_ID + 1).Insert
Worksheets("Rollup").Columns(Col_Date + 1).Delete
Do Until Worksheets("Rollup").Cells(1, Blank_Col) = ""
    Blank_Col = Blank_Col + 1
Loop
Worksheets("Rollup").Columns(Blank_Col).Delete
'Format the CRASH_DATETIME column
Dim DateTimeCol As Integer
DateTimeCol = 1
Do Until Worksheets("Rollup").Cells(1, DateTimeCol) = "CRASH_DATETIME"
    DateTimeCol = DateTimeCol + 1
Loop

Range(Worksheets("Rollup").Cells(2, DateTimeCol), Worksheets("Rollup").Cells(2, DateTimeCol).End(xlDown)).NumberFormat = "mmm d, yyyy h:mm:ss AM/PM"
    

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub Fix_Routes()
'Fix_Routes macro:
'   (1) Fix the Route Number, into 4 digit codes
'   (2) Move Non-State Route Crashes to another sheet
'   (3) Move Ramp Crashes to another sheet
'
' Created by: Samuel Mineer, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Samuel Mineer and Josh Gibbons, BYU, 2016

'Declare Variables
Dim nrow, ncol As Long              'Row counter
Dim Route As String                 'Variable for ROUTE_ID
Dim irou As Integer                 'Column counter, to find the ROUTE_ID data
Dim crow, prow, rrow As Double      'Row counter
Dim iramp As Integer                'Column counter, to find the RAMP data
Dim ramp As String                  'Variable for RAMP data
Dim i As Long                       'Counter variable for For Next loops
Dim idir As Integer
Dim ilab As Integer
Dim Label As String
Dim Direction As String

'Screen Updating OFF
Application.ScreenUpdating = False

'Reset progress values
num10000Crash = 0
numCrash = 1

'Find ROUTE_ID column in rollup sheet
irou = 1
Do While Sheets("Rollup").Cells(1, irou) <> "ROUTE_ID"
    irou = irou + 1
Loop

'Find direction and label columns
idir = 1
ilab = 1
Do While Sheets("Rollup").Cells(1, idir) <> "DIRECTION"
    idir = idir + 1
Loop

'Find column extent of rollup sheet
ncol = Sheets("Rollup").Range("A1").End(xlToRight).Column

'Find RAMP_ID column in rollup sheet
iramp = 1
Do While Sheets("Rollup").Cells(1, iramp) <> "RAMP_ID"
    iramp = iramp + 1
Loop

'Copy crash headers to Non-State Route Crashes, Incomplete Data, and Ramp Crashes sheets
For i = 1 To ncol
    Sheets("Incomplete Data").Cells(1, i) = Sheets("Rollup").Cells(1, i)
Next i
For i = 1 To ncol
    Sheets("Non-State Route Crashes").Cells(1, i) = Sheets("Rollup").Cells(1, i)
Next i
For i = 1 To ncol
    Sheets("Ramp Crashes").Cells(1, i) = Sheets("Rollup").Cells(1, i)
Next i

'Assign initial row values
nrow = 2
prow = 2
crow = 2
rrow = 2

'Cycle through each crash in rollup sheet to correct route values and remove ramp crashes
Do
    Route = CStr(Sheets("Rollup").Cells(nrow, irou))                  'Assign route number to Route
    ramp = Sheets("Rollup").Cells(nrow, iramp)                        'Assign ramp number, if any, to ramp
    Direction = Sheets("Rollup").Cells(nrow, idir)
    
    If Route = "089A" Or Route = "89A" Then      'Corrects strange route ID. 089A previously called SR-11.
        Route = "0011"          'Change route to 0011 for the purpose of this process.
    End If
    
    If Sheets("Rollup").Cells(nrow, 1) = 11038650 Then
        Route = Route  'check to see what happens to Crash ID #11038653's crash_datetime
    End If
    
    'Check label and direction. Change to positive if necessary to match route directions
    If Direction = "N" And (CLng(Route) <> 15 And CLng(Route) <> 70 And _
    CLng(Route) <> 80 And CLng(Route) <> 84 And CLng(Route) <> 85 And _
    CLng(Route) <> 215) Then
        Direction = "P"
        
        Sheets("Rollup").Cells(nrow, idir) = Direction
    End If
    
    'If there is no route number or crash ID then cut and paste row to Incomplete Data sheet
    If Route = "" Then
        
        For i = 1 To ncol
            Sheets("Incomplete Data").Cells(prow, i) = Sheets("Rollup").Cells(nrow, i)
            Sheets("Rollup").Cells(nrow, i) = ""
        Next i
        Application.CutCopyMode = False
        prow = prow + 1
        
    'If Route number is greater than 491 then cut and paste row to Non-State Route Crashes sheet
    'Camille Lunt ElseIf (CDbl(Route) > 491) Then
        
        'For i = 1 To ncol
            'Sheets("Non-State Route Crashes").Cells(crow, i) = Sheets("Rollup").Cells(nrow, i)
            'Sheets("Rollup").Cells(nrow, i) = ""
        'Next i
        'Application.CutCopyMode = False
        'crow = crow + 1
        
    'If there is a number for RAMP_ID, then cut and paste row to Ramp Crashes sheet.
    'ElseIf ramp <> "" Or Len(Route) >= 5 Then
        
        'For i = 1 To ncol
            'Sheets("Ramp Crashes").Cells(rrow, i) = Sheets("Rollup").Cells(nrow, i)
            'Sheets("Rollup").Cells(nrow, i) = ""
        'Next i
        'Application.CutCopyMode = False
        'rrow = rrow + 1
        
    Else
        If Len(Route) = 1 Then
            Route = "000" & Route
        ElseIf Len(Route) = 2 Then
            Route = "00" & Route
        ElseIf Len(Route) = 3 Then
            Route = "0" & Route
        Else
        End If
        
        Sheets("Rollup").Cells(nrow, irou).NumberFormat = "@"
        Sheets("Rollup").Cells(nrow, irou) = Route
    End If
    nrow = nrow + 1
        
    'Update progress
    numCrash = nrow
    num10000Crash = num10000Crash + 1
    If num10000Crash = 10000 Or numCrash = numTotCrash Then
        Sheets("Progress").Activate
        
        'Screen Updating ON
        Application.ScreenUpdating = True
        
        Range("A2") = "Fix RouteID (2 of 3)"
        Range("A3") = CStr(numCrash) & " of " & CStr(numTotCrash) & " crashes"
        Range("C3") = numCrash / numTotCrash
        Range("A5") = "Update Time"
        Range("B5") = Time
        
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'Screen Updating OFF
        Application.ScreenUpdating = False
        
        num10000Crash = 0
    End If
        
    Application.CutCopyMode = False
Loop Until Sheets("Rollup").Cells(nrow, 1) = ""           'Loop until no crash id is found in the 1st column

'Run MoveBlankRowsToBottom macro:   (faster than deleting rows)
'   (1) All rows that have been cleared of data are moved to bottom of sheet by a sorting method.
MoveBlankRowsToBottom (nrow)

'Screen Updating ON
Sheets("Progress").Activate
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub Direction_Code()
'Direction_Code macro:
'   (1) Adds a column for Directional Code, based on vehicle direction
'       P for positive
'       X for negative directions, interstates only
'
' Created by: Samuel Mineer, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Samuel Mineer and Josh Gibbons, BYU, 2016

'Define Variables
Dim n As Double
Dim Direction As Integer
Dim Route As String
Dim irou As Integer
Dim idir As Integer

'Screen Updating OFF
Application.ScreenUpdating = False

'Reset progress values
num10000Crash = 0
numCrash = 1

'Find columns for ROUTE_ID and First_Veh_Direction
Sheets("Location").Activate
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop
idir = 1
Do Until Cells(1, idir) = "First_Veh_Direction"
    idir = idir + 1
Loop

'Add column for DIRECTION
Columns(idir + 1).Insert shift:=xlShiftToRight
Columns(idir + 1).NumberFormat = "@"
Cells(1, idir + 1) = "DIRECTION"

'Add second row for negative direction for divided highways/freeways
n = 2
Do
    Route = Cells(n, irou)
    Direction = Cells(n, idir)
    If (Route = "0015" Or Route = "0070" Or Route = "0080" Or Route = "0084" Or Route = "0215" Or Route = "0085") Then
        If (Direction = 2 Or Direction = 4) Then
            Cells(n, idir + 1) = "X"
        Else
            Cells(n, idir + 1) = "P"
        End If
    Else
        Cells(n, idir + 1) = "P"
    End If
    n = n + 1
    
    'Update progress
    numCrash = n
    num10000Crash = num10000Crash + 1
    If num10000Crash = 10000 Or numCrash = numTotCrash Then
        Sheets("Progress").Activate
        
        'Screen Updating ON
        Application.ScreenUpdating = True
        
        Range("A2") = "Fix Direction (3 of 4)"
        Range("A3") = CStr(numCrash) & " of " & CStr(numTotCrash) & " crashes"
        Range("C3") = numCrash / numTotCrash
        Range("A5") = "Update Time"
        Range("B5") = Time
        
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'Screen Updating OFF
        Application.ScreenUpdating = False
        
        num10000Crash = 0
        
        Sheets("Location").Activate
    End If
    
Loop Until Cells(n, 1) = "" And Cells(n + 1, 1) = 0

'Screen Updating ON
Sheets("Progress").Activate
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub IntVehiclePrep()
'
' Samuel Runyan Note: 8/10/21 - It appears this macro was created at least in part to find the true number of vehicles for each crash by finding duplicate crash IDs.
' Editted and Commented by Samuel Runyan BYU - 8/20/21
'
Dim VehSheet As String
Dim IDcol As Integer
Dim VehNumCol, EventSeqCol As Integer
Dim CrashRow As Long
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer


IDcol = 1
VehNumCol = 1
EventSeqCol = 1
CrashRow = 2
VehSheet = "Vehicle"

'Determine Column Numbers
Do Until Sheets(VehSheet).Cells(1, IDcol) = "CRASH_ID"
    IDcol = IDcol + 1
Loop

Do Until Sheets(VehSheet).Cells(1, VehNumCol) = "VEHICLE_NUM"
    VehNumCol = VehNumCol + 1
Loop

Do Until Left(Sheets(VehSheet).Cells(1, EventSeqCol), 5) = "EVENT"
    EventSeqCol = EventSeqCol + 1
Loop


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Fixes Event Sequence Columns
Do Until Sheets(VehSheet).Cells(CrashRow, IDcol) = ""
'This checks for duplicate crash IDs thus signifying multiple vehicles in a crash.
    i = 1
    Do Until Sheets(VehSheet).Cells(CrashRow + i, IDcol) <> Sheets(VehSheet).Cells(CrashRow, IDcol)
        i = i + 1
    Loop
    Sheets(VehSheet).Cells(CrashRow, VehNumCol) = i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This ensures the event sequences from the duplicate crash IDs match and clears the additional rows
    If i > 1 Then
        For j = 2 To i                                                                              'Cycle through each secondary crash ID row
            'ReDim Events(1 To 4) As Integer
            'For k = 1 To 4                                                                         'Cycle through each event in the secondary vehicle row
            '    Events(k) = Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + k - 1)          'Log secondary crash ID row events in events(k)
            'Next k
            For m = 1 To 4                                                                          'Cycle through each event in the vehicle num 1 row
                
                If Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) <> 96 And Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) <> Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m) Then
                    For k = 1 To 4                                                                  'Cycle through Events(k) to see if there are duplicates in events and changes them to 96
                        If Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + k - 1) = Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) Then       'If there are any unique events, we keep them.
                            Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + k - 1) = 96                                                              'Change number to 96 as a N/A value
                        End If
                    Next k
                End If
            Next m
                        
            If Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol) = 96 And Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + 1) = 96 And Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + 2) = 96 And Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + 3) = 96 Then
                'nothing
            Else
                For m = 1 To 4                                                                          'Cycle through each event in the vehicle num 1 row
                    If Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) = 96 Then                  'Check if first vehicle has a N/A event
                        For k = 1 To 4
                            If Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + k - 1) <> 96 Then                                                     'Fill in N/A event with next unique event from next vehicle.
                                Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) = Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + k - 1)
                                Sheets(VehSheet).Cells(CrashRow + j - 1, EventSeqCol + k - 1) = 96
                                Exit For
                            End If
                        Next k
                    End If
                Next m
            End If
        Next j
        Sheets(VehSheet).Rows(CrashRow + 1 & ":" & CrashRow + i - 1).EntireRow.ClearContents
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This determines if the event sequence was repeated and replaces with 96
    For k = 1 To 3
        If Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k - 1) = 96 Then
            'nothing
        Else
        For m = k + 1 To 4
            If Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k - 1) = Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) Then
                Sheets(VehSheet).Cells(CrashRow, EventSeqCol + m - 1) = 96                          '96 is the id number for unknown I think
            End If
        Next m
        End If
    Next k
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This makes sure the first known event is listed as event 1
    For k = 1 To 3
        If Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k - 1) = 96 And Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k) <> 96 Then
            Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k - 1) = Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k)
            Sheets(VehSheet).Cells(CrashRow, EventSeqCol + k) = 96
        End If
    Next k

    CrashRow = CrashRow + i
Loop

VehicleMoveBlankRowsToBottom (CrashRow - 1)

End Sub

Sub Create_Label()
'Create_Label macro:
'   (1) Creates a column for Label Data and fills in the column
'       (LABEL = Route_ID + Direction)
'
' Created by: Samuel Mineer, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Samuel Mineer and Josh Gibbons, BYU, 2016

'Screen Updating OFF
Sheets("Progress").Activate
Application.ScreenUpdating = False

'Define Variables
Dim nrow As Double          'Row counter
Dim irou As Integer         'column counter, for finding column with Route_ID
Dim idir As Integer         'column counter, for finding column with Direction
Dim Route As String         'string variable for Route_ID
Dim Direction As String     'string variable for Direction

'Reset progress values
num10000Crash = 0
numCrash = 1

'Find ROUTE_ID column in rollup sheet
Sheets("Rollup").Activate
irou = 1
Do Until Cells(1, irou) = "ROUTE_ID"
    irou = irou + 1
Loop

'Insert new column called LABEL
Columns(irou + 1).Insert shift:=xlShiftToRight
Cells(1, irou + 1) = "LABEL"

'Find DIRECTION column in rollup sheet
idir = 1
Do Until Cells(1, idir) = "DIRECTION"
    idir = idir + 1
Loop

'Combine route and direction values to form LABEL
nrow = 2
Do
    Route = Sheets("Rollup").Cells(nrow, irou)
    Direction = Sheets("Rollup").Cells(nrow, idir)

    Sheets("Rollup").Cells(nrow, irou + 1).NumberFormat = "@"
    Sheets("Rollup").Cells(nrow, irou + 1) = Route & Direction
        
    nrow = nrow + 1
    
    'Update progress
    numCrash = nrow
    num10000Crash = num10000Crash + 1
    If num10000Crash = 10000 Or numCrash = numTotCrash Then
        Sheets("Progress").Activate
        
        'Screen Updating ON
        Application.ScreenUpdating = True
        
        Range("A2") = "Create Labels (3 of 3)"
        Range("A3") = CStr(numCrash) & " of " & CStr(numTotCrash) & " crashes"
        Range("C3") = numCrash / numTotCrash
        Range("A5") = "Update Time"
        Range("B5") = Time
        
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'Screen Updating OFF
        Application.ScreenUpdating = False
        
        num10000Crash = 0
        
        Sheets("Rollup").Activate
    End If
    
Loop Until Sheets("Rollup").Cells(nrow, irou) = ""

'Assume code has reached 100%
numCrash = numTotCrash
Sheets("Progress").Activate
Range("A3") = CStr(numCrash) & " of " & CStr(numTotCrash) & " crashes"
Range("C3") = numCrash / numTotCrash
Range("B5") = ""
Range("A6") = "End Time"
Range("B6") = Time

'Screen Updating ON
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub MoveBlankRowsToBottom(nrow As Long)
'MoveBlankRowsToBottom macro:
'   (1) All rows that have been cleared of data are moved to bottom of sheet by a sorting method.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Screen Updating OFF
Sheets("Home").Activate
Application.ScreenUpdating = False

'Declare variable
Dim irou As Integer
    'Activate Rollup sheet
    Sheets("Rollup").Activate
    
    'Find ROUTE_ID column
    irou = 1
    Do Until Sheets("Rollup").Cells(1, irou) = "ROUTE_ID"
        irou = irou + 1
    Loop
    
    'Select all cells, sort by ROUTE_ID to move blank rows to bottom
    Cells.AutoFilter
    ActiveWorkbook.Worksheets("Rollup").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rollup").AutoFilter.Sort.SortFields.Add Key:= _
        Range(Cells(2, irou), Cells(nrow, irou)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rollup").AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.AutoFilter

'Screen Updating ON
Sheets("Progress").Activate
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub VehicleMoveBlankRowsToBottom(nrow As Long)
'MoveBlankRowsToBottom macro:
'   (1) All rows that have been cleared of data are moved to bottom of sheet by a sorting method.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Screen Updating OFF
Sheets("Progress").Activate
Application.ScreenUpdating = False

'Declare variable
Dim irou As Integer
    'Activate Location sheet
    Sheets("Vehicle").Activate
    
    'Find ROUTE_ID column
    irou = 1
    Do Until Sheets("Vehicle").Cells(1, irou) = "CRASH_ID"
        irou = irou + 1
    Loop
    
    'Select all cells, sort by ROUTE_ID to move blank rows to bottom
    Cells.AutoFilter
    ActiveWorkbook.Worksheets("Vehicle").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Vehicle").AutoFilter.Sort.SortFields.Add Key:= _
        Range(Cells(2, irou), Cells(nrow, irou)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Vehicle").AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.AutoFilter

Sheets("Vehicle").UsedRange

'Screen Updating ON
Sheets("Progress").Activate
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Public Sub cleanNumetric(sname As String)
    '
    '
    'Created by Samuel Runyan - 8/10/21
    'This macro is used for converting the Numetric data to a format that the model will recognize.
    '
    '
    
Dim col As Integer
Dim col2, row, i As Long
Dim headerNames(24) As String
Dim header                    'placeholder for the array

'turn screen updating OFF
Application.ScreenUpdating = False
'activate sheet
Sheets(sname).Activate

'First checks if the data needs to be cleaned and exits the sub if not.
col = 0
Do
    col = col + 1
    If VarType(Cells(2, col)) = vbString And Cells(2, col) <> "Y" _
    And Cells(2, col) <> "N" And Cells(2, col) <> "P" And Cells(2, col) <> "S" _
    And Cells(2, col) <> "F" And Cells(2, col) <> "L" And Cells(2, col) <> "U" _
    And VarType(Cells(2, col)) <> vbUserDefinedType _
    And VarType(Cells(2, col)) <> vbDate Then
        Exit Do
    ElseIf Cells(1, col) = "" Then
        Exit Sub
    End If
Loop Until Cells(1, col) = ""
'If the data "passes" this first test, it moves forward with the rest of the code.

'set the columns we want to change
headerNames(0) = "IMPROPER_RESTRAINT"
headerNames(1) = "URBAN_COUNTY"
headerNames(2) = "ROUTE_TYPE"
headerNames(3) = "NIGHT_DARK_CONDITION"
headerNames(4) = "SINGLE_VEHICLE"
headerNames(5) = "CRASH_SEVERITY_ID"
headerNames(6) = "LIGHT_CONDITION_ID"
headerNames(7) = "WEATHER_CONDITION_ID"
headerNames(8) = "MANNER_COLLISION_ID"
headerNames(9) = "ROADWAY_SURF_CONDITION_ID"
headerNames(10) = "ROADWAY_JUNCT_FEATURE_ID"
headerNames(11) = "HORIZONTAL_ALIGNMENT_ID"
headerNames(12) = "VERTICAL_ALIGNMENT_ID"
headerNames(13) = "ROADWAY_CONTRIB_CIRCUM_ID"
headerNames(14) = "FIRST_HARMFUL_EVENT_ID"
headerNames(15) = "VEHICLE_NUM"
headerNames(16) = "TRAVEL_DIRECTION_ID"
headerNames(17) = "EVENT_SEQUENCE_1_ID"
headerNames(18) = "EVENT_SEQUENCE_2_ID"
headerNames(19) = "EVENT_SEQUENCE_3_ID"
headerNames(20) = "EVENT_SEQUENCE_4_ID"
headerNames(21) = "MOST_HARMFUL_EVENT_ID"
headerNames(22) = "VEHICLE_MANEUVER_ID"
headerNames(23) = "UTM_Y"
headerNames(24) = "UTM_X"

i = 0
'start looping through headerNames
For Each header In headerNames
    'find the header column number
    col = 1
    Do Until Cells(1, col) = header
        If Cells(1, col) = "" Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   'Skips the header if it doesn't appear on the sheet.
            GoTo nextHeader                                                'GoTo can be problematic for debugging so be careful with it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        col = col + 1
    Loop
    row = 2
    'check for the index number then run through the column to clean the data
    If i = 0 Then
        'clean IMPROPER_RESTRAINT column
        Do Until Cells(row, 1) = ""                                                         'First fill blank cells
            If Cells(row, col) = "" Then
                Cells(row, col) = "N"
            End If
            row = row + 1
        Loop
        
        Find_Replace """", "", col, sname
        Find_Replace "(retired) NOT Used Properly", "Y", col, sname                               'Replace strings with Y or N
        Find_Replace "(retired) Used Properly", "N", col, sname
        Find_Replace "Not Applicable (not restrained)", "N", col, sname
        Find_Replace "Unknown (Due to extent of damage/evidence destroyed)", "N", col, sname
        
        col2 = 1                                                                            'Find UNRESTRAINED column
        Do Until Cells(1, col2) = "UNRESTRAINED"
            col2 = col2 + 1
        Loop
        row = 2
        Do Until Cells(row, 1) = ""                                                         'If the driver was unrestrained, put N for improper restaint
            If Cells(row, col2) = "Y" Then
                Cells(row, col) = "N"
            End If
            row = row + 1
        Loop
        
        row = 2
        Do Until Cells(row, 1) = ""                                                         'Only take the first letter meaning if there was a Y at all it is recorded
            Cells(row, col) = Right(Left(Cells(row, col), 2), 1)                            'In other words, if even one passenger was improperly restrained
            row = row + 1
        Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 1 Then
        'clean URBAN_COUNTY column
        Do Until Cells(row, 1) = ""
            If Cells(row, col) = "Cache" Or Cells(row, col) = "Davis" Or Cells(row, col) = "Salt Lake" Or Cells(row, col) = "Utah" Or Cells(row, col) = "Weber" Then
                Cells(row, col) = "Y"
            Else
                Cells(row, col) = "N"
            End If
            row = row + 1
        Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 2 Then
        'clean ROUTE_TYPE column
        Do Until Cells(row, 1) = ""
            Cells(row, col) = Left(Cells(row, col), 1)
            row = row + 1
        Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 3 Then
        'clean NIGHT_DARK_CONDITION column
        Do Until Cells(row, 1) = ""
            Cells(row, col) = Left(Cells(row, col), 1)
            row = row + 1
        Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 4 Then
        'clean SINGLE_VEHICLE column
        Do Until Cells(row, 1) = ""
            If VarType(Cells(row, col)) = vbString And Cells(row, col) <> "Y" And Cells(row, col) <> "N" Then         'check if string other than N or Y
                Cells(row, col) = CInt(Left(Right(Cells(row, col), 2), 1))              'take last number and convert to integer
            End If
            If Cells(row, col) = 1 Then                                                 'determine if single vehicle
                Cells(row, col) = "Y"
            Else
                Cells(row, col) = "N"
            End If
            row = row + 1
        Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 5 Then
        'clean CRASH_SEVERITY_ID column
        Find_Replace "No injury/PDO", 1, col, sname                                'Replace strings with severity number"
        Find_Replace "Possible injury", 2, col, sname
        Find_Replace "Suspected Minor Injury", 3, col, sname
        Find_Replace "Suspected Serious Injury", 4, col, sname
        Find_Replace "Fatal", 5, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 6 Then
        'clean LIGHT_CONDITION_ID column
        Find_Replace "Daylight", 1, col, sname                                    'Replace strings with id numbers"
        Find_Replace "Dark - Lighted", 2, col, sname
        Find_Replace "Dark - Not Lighted", 3, col, sname
        Find_Replace "Dark - Unknown Lighting", 4, col, sname
        Find_Replace "Dawn", 5, col, sname
        Find_Replace "Dusk", 6, col, sname
        Find_Replace "Other", 89, col, sname
        Find_Replace "Unknown", 99, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 7 Then
        'clean WEATHER_CONDITION_ID column
        Find_Replace "Clear", 1, col, sname                                       'Replace strings with id numbers"
        Find_Replace "Cloudy", 2, col, sname
        Find_Replace "Rain", 3, col, sname
        Find_Replace "Snowing", 4, col, sname
        Find_Replace "Blowing Snow", 5, col, sname
        Find_Replace "Sleet, Hail", 6, col, sname
        Find_Replace "Fog, Smog", 7, col, sname
        Find_Replace "Severe Crosswinds", 8, col, sname
        Find_Replace "(retired) Blowing Sand, Soil, Dirt", 9, col, sname
        Find_Replace "Other", 89, col, sname
        Find_Replace "Unknown", 99, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 8 Then
        'clean MANNER_COLLISION_ID column
        Find_Replace "Angle", 1, col, sname                                       'Replace strings with id numbers"
        Find_Replace "Front to Rear", 2, col, sname
        Find_Replace "Head On (front-to-front)", 3, col, sname
        Find_Replace "Sideswipe Same Direction", 4, col, sname
        Find_Replace "Sideswipe Opposite Direction", 5, col, sname
        Find_Replace "Parked Vehicle", 6, col, sname
        Find_Replace "Rear to Side", 7, col, sname
        Find_Replace "Rear to Rear", 8, col, sname
        Find_Replace "Other", 89, col, sname
        Find_Replace "Not Applicable/Single Vehicle", 96, col, sname
        Find_Replace "Unknown", 99, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 9 Then
        'clean ROADWAY_SURF_CONDITION_ID column
        Find_Replace "Dry", 1, col, sname                                         'Replace strings with id numbers"
        Find_Replace "Wet", 2, col, sname
        Find_Replace "Snow", 3, col, sname
        Find_Replace "Slush", 4, col, sname
        Find_Replace "Ice/Frost", 5, col, sname
        Find_Replace "Water (standing, moving", 6, col, sname
        Find_Replace "Mud", 7, col, sname
        Find_Replace "(retired) Sand, Dirt, Gravel", 8, col, sname
        Find_Replace "Oil", 9, col, sname
        Find_Replace "Unknown", 89, col, sname
        Find_Replace "Other*", 97, col, sname
        Find_Replace "Unknown", 99, col, sname        'not sure why 89 and 99 are the same. 99 will end up being 89
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 10 Then
        'clean ROADWAY_JUNCT_FEATURE_ID column
        Find_Replace "No Special Feature/Junction", 0, col, sname                 'Replace strings with id numbers"
        Find_Replace "Bridge (overpass/underpass)", 1, col, sname
        Find_Replace "Railroad Crossing", 2, col, sname
        Find_Replace "Business Drive", 3, col, sname
        Find_Replace "Farm/Residential Drive", 4, col, sname
        Find_Replace "Alley", 5, col, sname
        Find_Replace "Crossover in Median", 6, col, sname
        Find_Replace "On-Ramp Merge Area (Acceleration Lane)", 7, col, sname
        Find_Replace "Off-Ramp Diverge Area (Deceleration Lane)", 8, col, sname
        Find_Replace "On-Ramp", 9, col, sname
        Find_Replace "Off-Ramp", 10, col, sname
        Find_Replace "4-Leg Intersection", 20, col, sname
        Find_Replace "T-Intersection", 21, col, sname
        Find_Replace "Y-Intersection", 22, col, sname
        Find_Replace "5-Leg or More Intersection", 23, col, sname
        Find_Replace "Roundabout/Traffic Circle", 24, col, sname
        Find_Replace "Ramp Intersection With Crossroad", 25, col, sname
        Find_Replace "Multi Use Path/Trail Intersection", 26, col, sname
        Find_Replace "Unknown", 89, col, sname
        Find_Replace "Other*", 97, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 11 Then
        'clean HORIZONTAL_ALIGNMENT_ID column
        Find_Replace "Straight", 1, col, sname                                    'Replace strings with id numbers"
        Find_Replace "(retired) Curved", 2, col, sname
        Find_Replace "Unknown", 89, col, sname
        Find_Replace "Unknown", 99, col, sname     'not sure why 89 and 99 are the same. 99 will end up being 89
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 12 Then
        'clean VERTICAL_ALIGNMENT_ID column
        Find_Replace "Level", 1, col, sname                                       'Replace strings with id numbers"
        Find_Replace "retired (Grade)", 2, col, sname
        Find_Replace "Hillcrest", 3, col, sname
        Find_Replace "Sag (Bottom)", 4, col, sname
        Find_Replace "Other", 89, col, sname
        Find_Replace "Unknown", 99, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 13 Then
        'clean ROADWAY_CONTRIB_CIRCUM_ID column
        Find_Replace "Shoulder (none, low, soft, high)", 7, col, sname
        Find_Replace "None", 0, col, sname                                        'Replace strings with id numbers"
        Find_Replace "Debris", 1, col, sname
        Find_Replace "Rut, Hole, Bump", 2, col, sname
        Find_Replace "Road Surface Condition (wet, icy, snow, slush, etc.)", 3, col, sname
        Find_Replace "Work Zone (construction/maintenance/utility)", 4, col, sname
        Find_Replace "Worn, Travel-Polished Surface", 5, col, sname
        Find_Replace "Traffic Control Device (inoperative, missing, or obsecured)", 6, col, sname
        Find_Replace "Animal Caused Evasive Action", 8, col, sname
        Find_Replace "Non-Motorist Caused Evasive Action", 9, col, sname
        Find_Replace "Non-Contact Vehicle Caused Evasive Action", 10, col, sname
        Find_Replace "Prior Crash", 11, col, sname
        Find_Replace "Other*", 89, col, sname
        Find_Replace "Other*", 97, col, sname     'not sure why 89 and 97 are the same. 97 will end up being 89
        Find_Replace "Unknown", 99, col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 14 Then
        'clean FIRST_HARMFUL_EVENT_ID column
        cleanEvents col, sname

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 15 Then
        'clean VEHICLE_NUM column
        Do Until Cells(row, 1) = ""
            If VarType(Cells(row, col)) = vbString Then                             'check if string
                Cells(row, col) = CInt(Left(Right(Cells(row, col), 2), 1))          'take last number and convert to integer
            End If
            row = row + 1
        Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 16 Then
        'clean TRAVEL_DIRECTION_ID column
        Find_Replace """", "", col, sname
        Find_Replace "Northbound", 1, col, sname                                            'Replace strings with id numbers"
        Find_Replace "Southbound", 2, col, sname
        Find_Replace "Eastbound", 3, col, sname
        Find_Replace "Westbound", 4, col, sname
        Find_Replace "Not On Roadway (also for parked motor vehicle)", "NA", col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 17 Then
        'clean EVENT_SEQUENCE_1_ID column
        cleanEvents col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 18 Then
        'clean EVENT_SEQUENCE_2_ID column
        cleanEvents col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 19 Then
        'clean EVENT_SEQUENCE_3_ID column
        cleanEvents col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 20 Then
        'clean EVENT_SEQUENCE_4_ID column
        cleanEvents col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 21 Then
        'clean MOST_HARMFUL_EVENT_ID column
        cleanEvents col, sname
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf i = 22 Then
        'clean VEHICLE_MANEUVER_ID column
        Find_Replace """", "", col, sname
        Find_Replace "Straight Ahead", 1, col, sname                                         'Replace strings with id numbers"
        Find_Replace "Backing", 2, col, sname
        Find_Replace "Changing Lanes", 3, col, sname
        Find_Replace "Overtaking/Passing", 4, col, sname
        Find_Replace "Turning Right", 5, col, sname
        Find_Replace "Turning Left", 6, col, sname
        Find_Replace "Making U-turn", 7, col, sname
        Find_Replace "Leaving Traffic Lane", 8, col, sname
        Find_Replace "Entering Traffic Lane", 9, col, sname
        Find_Replace "Stopped in Traffic Lane", 10, col, sname
        Find_Replace "Slowing in Traffic Lane", 11, col, sname
        Find_Replace "Mechanically Disabled in Traffic Lane", 12, col, sname
        Find_Replace "Parked", 13, col, sname
        Find_Replace "Parking Maneuvers", 14, col, sname
        Find_Replace "Unknown", 89, col, sname
        Find_Replace "Other*", 97, col, sname
        Find_Replace "Unknown", 99, col, sname
    End If
nextHeader:                         'skip if header missing
    i = i + 1                       'increase index number
Next header
    
End Sub

Sub cleanEvents(col As Integer, sname As String)
'
'
'Created by Samuel Runyan - 8/12/21
'This was created to accompany the cleanNumetric macro because there are several columns which use event data.
'
'

Find_Replace "(Retired) Ran Off Road Right", 1, col, sname                            'Replace strings with id numbers"
Find_Replace "Ran Off Road Right", 1, col, sname
Find_Replace "(Retired) Ran Off Road Left", 2, col, sname
Find_Replace "Ran Off Road Left", 2, col, sname
Find_Replace "(Retired) Crossed Median/Centerline", 3, col, sname
Find_Replace "Crossed Median/Centerline", 3, col, sname
Find_Replace "(Retired) Equipment Failure (tire, brakes, etc.)", 4, col, sname
Find_Replace "Equipment Failure (tire, brakes, etc.)", 4, col, sname
Find_Replace "Separation of Units", 5, col, sname
Find_Replace "Downhill Runaway", 6, col, sname
Find_Replace "Overturn/Rollover", 7, col, sname
Find_Replace "Cargo Equipment Loss or Shift", 8, col, sname
Find_Replace "Jackknife", 9, col, sname
Find_Replace "Fire/Explosion", 10, col, sname
Find_Replace "Immersion", 11, col, sname
Find_Replace "Fell/Jumped From Motor Vehicle", 12, col, sname
Find_Replace "Other Non-Collision*", 19, col, sname
Find_Replace "Other Non-Motorist*", 19, col, sname
Find_Replace "Collision With Other Motor Vehicle in Transport", 20, col, sname
Find_Replace "Collision With Parked Motor Vehicle", 21, col, sname
Find_Replace "Pedestrian", 22, col, sname
Find_Replace "Pedacycle", 23, col, sname
Find_Replace "Skates, Scooter, Skateboards", 24, col, sname
Find_Replace "Animal - Wild", 25, col, sname
Find_Replace "Animal - Domestic", 26, col, sname
Find_Replace "Work Zone/Maintenance Equipment", 27, col, sname
Find_Replace "Freight Rail", 28, col, sname
Find_Replace "Light Rail", 29, col, sname
Find_Replace "Passenger Heavy Rail", 30, col, sname
Find_Replace "Thrown or Fallen Object", 31, col, sname
Find_Replace "Collision Between Motor Vehicle in Transport and Vehicle Cargo/Part or Object Set in Motion by Motor Vehicle", 32, col, sname
Find_Replace "Other Non-Fixed Object*", 39, col, sname

Find_Replace "Guardrail End Section", 44, col, sname
Find_Replace "Concrete Sloped End Section", 45, col, sname
Find_Replace "Cable Barrier End Section", 46, col, sname
Find_Replace "Guardrail", 40, col, sname
Find_Replace "Concrete Barrier", 41, col, sname
Find_Replace "Cable Barrier", 42, col, sname
Find_Replace "Crash Cushion", 43, col, sname

Find_Replace "Access Control Cable", 47, col, sname
Find_Replace "Bridge Rail", 48, col, sname
Find_Replace "Bridge Pier or Support", 49, col, sname
Find_Replace "Bridge Overhead Structure", 50, col, sname
Find_Replace "Traffic Sign Support", 51, col, sname
Find_Replace "Delineator Post", 52, col, sname
Find_Replace "Other Post, Pole or Support", 53, col, sname
Find_Replace "Utility Pole/Light Support", 54, col, sname
Find_Replace "Traffic Signal Support", 55, col, sname
Find_Replace "Culvert", 56, col, sname
Find_Replace "Ditch", 57, col, sname
Find_Replace "Embankment", 58, col, sname
Find_Replace "Snow Bank", 59, col, sname
Find_Replace "Tree/Shrubbery", 60, col, sname
Find_Replace "Mailbox/Fire Hydrant", 61, col, sname
Find_Replace "Fence", 62, col, sname
Find_Replace "Curb", 63, col, sname
Find_Replace "Cub", 63, col, sname
Find_Replace "Other Fixed Object*", 69, col, sname
Find_Replace "Not Applicable (used only to fill unused box(es))", 96, col, sname
Find_Replace "Unknown", 96, col, sname
Find_Replace "Other", 89, col, sname
Find_Replace """", "", col, sname

End Sub
