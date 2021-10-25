Attribute VB_Name = "DP_08_Combine"
Option Explicit

'PMDP_08_Combine Module:

'Declare public variables
Public DYear1 As Integer                    'Last year of requested AADT data
Public FileName1 As String                  'Working name of data sheet
Public FN1, FN2, FN3, FN4, FN5 As String    'Dataset names of 5 preliminary roadway datasets
Public CurrentWkbk As String                'String variable that holds current workbook name for reference purposes

Sub Run_Combined()
'Run_Combined macro:
'   (1) Finishes formatting and combining the roadway data into one dataset.
'   (2) Routes are segmented to form homogeneous segments.
'   (3) Segment length at every change or at specified length given by user.
'   (4) Segments are combined into larger segments based on minimum and maximum segment lengths.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons and Marlee Seat, BYU, 2016
' Commented by: Samuel Mineer and Josh Gibbons, BYU, 2016

'Declare variables
Dim ws As Worksheet                                     'Worksheet variable used to cycle through worksheets
Dim iCol As Integer                                     'Column counter
Dim wd As String
Dim sname As String
Dim guiwb As String
Dim StartTime, EndTime As String
Dim myrow As Long, mycol As Integer

wd = Sheets("Inputs").Range("I2").Value
guiwb = ActiveWorkbook.Name

'Screen Updating OFF
Application.ScreenUpdating = False

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'AssignInfo macro:
'   (1) Go through the sheets in the workbook to find the roadway data sheets.
'   (2) Assign sheet names to public variables that can be refered to when working with the datasets.
AssignInfo

'Run Delete_Labels macro:
'   (1) Finds which routes are not present in all 5 roadway data sets.
'   (2) Deletes any of those routes in the datasets.
Delete_Labels

'Run Condense_SpeedLimit macro:
'   (1) Combines speed limit data where there are consecutive segments listed with the same speed limit.
iCol = 1
Do Until Worksheets("Speed_Limit").Cells(1, iCol) = ""
    If Worksheets("Speed_Limit").Cells(1, iCol) = "RECEIVED" Then
        Condense_SpeedLimit
    End If
    iCol = iCol + 1
Loop

'Run Condense_ThruLanes macro:
'   (1) Corrects beginning/ending milepoints in order to create the correct milepoint range for each segment.
Condense_ThruLanes

'Run SameMP macro:
'   (1) Verifies that each dataset has the same ending milepoints for each route number.
'   (2) All datasets are compared to the values from the Functional Class data.
SameMP

'Run CombineData macro:
'   (1) Combine all roadway datasets onto the Roadway Data sheet.
CombineData

'Run CombinedHeaders macro:
'   (1) Adds headers to the combined roadway data sheet to match the data that has been placed in there.
CombinedHeaders

'Run FixBlankSpeedLimit macro:
'   (1) Corrects speed limits that have a blank or zero value using default speed limits based on functional class.
FixBlankSpeedLimit

'not needed according to email from Timo on August 28, 2019
''Run CreateInteractions macro:
''   (1) Add columns and calculate values for "interactions" that are used in the R model.
'CreateInteractions

'Run Seg_Length macro:
'   (1) Add column for the segment length of each roadway segment.
'   (2) Calculate the segment lengths by taking the difference of the beginning and ending milepoints.
Seg_Length

'Run Cleanup macro:
'   (1) Combine segments that have the same roadway characteristics.
CleanUp

'Run MaxTrim macro:
'   (1) Trim segments to lengths less than max length, if the "Max Length" option button is selected.
MaxTrim

'for CAMS only: create Seg ID's
If Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    CAMS_SegID
End If

'Delete the now-unneeded worksheets
Application.DisplayAlerts = False                       'Do not show alert messages
Dim wksht As Worksheet
For Each wksht In Worksheets
    If InStr(1, wksht.Name, "AADT") > 0 Then
        wksht.Delete
    ElseIf wksht.Name = "Functional_Class" Then
        wksht.Delete
    ElseIf wksht.Name = "Speed_Limit" Then
        wksht.Delete
    ElseIf wksht.Name = "Thru_Lanes" Then
        wksht.Delete
    ElseIf wksht.Name = "Urban_Code" Then
        wksht.Delete
    End If
Next
Application.DisplayAlerts = True                        'Show alert messages

If Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    myrow = 2
    Do Until Sheets("Roadway Data").Cells(myrow, 1) = ""
        If Sheets("Roadway Data").Cells(myrow, 6) = "0085" And Sheets("Roadway Data").Cells(myrow, 3) < 2.977 Then
            Sheets("Roadway Data").Cells(myrow, 2) = replace(Sheets("Roadway Data").Cells(myrow, 2), "085", "194")
            Sheets("Roadway Data").Cells(myrow, 6) = replace(Sheets("Roadway Data").Cells(myrow, 6), "085", "194")
            Sheets("Roadway Data").Cells(myrow, 6).NumberFormat = "0000"
            Sheets("Roadway Data").Cells(myrow, 6).HorizontalAlignment = xlHAlignLeft
            Sheets("Roadway Data").Cells(myrow, 7) = replace(Sheets("Roadway Data").Cells(myrow, 7), "085", "194")
            Sheets("Roadway Data").Cells(myrow, 7).NumberFormat = "0000"
            Sheets("Roadway Data").Cells(myrow, 7).HorizontalAlignment = xlHAlignLeft
        End If
        myrow = myrow + 1
    Loop
    
    're-sort
    mycol = 1
    Do Until Sheets("Roadway Data").Cells(1, mycol) = ""
        mycol = mycol + 1
    Loop
    
    ActiveWorkbook.Worksheets("Roadway Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Roadway Data").Sort.SortFields.Add Key:=Range( _
        "B2:B" & myrow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Roadway Data").Sort.SortFields.Add Key:=Range( _
        "C2:C" & myrow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Roadway Data").Sort
        .SetRange Range(Cells(1, 1), Cells(myrow - 1, mycol - 1))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    're-number
    myrow = 2
    Do Until Sheets("Roadway Data").Cells(myrow, 1) = ""
        Sheets("Roadway Data").Cells(myrow, 1) = myrow - 1
        myrow = myrow + 1
    Loop
    
End If

'Save dataset as separate workbook
Sheets("Roadway Data").Move
sname = wd & "/CAMSRoadSegments_" & replace(Date, "/", "-") & "_" & replace(Time, ":", "-") & ".csv"
sname = replace(sname, " ", "_")
sname = replace(sname, "/", "\")
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Cells(5, 13).Value = replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")
ActiveWorkbook.Close
Application.DisplayAlerts = True

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Roadway Data Compiled."
    .Range("A3") = "Returning to Input Screen."
    .Range("A5") = ""
    .Range("B5") = ""
    .Range("A6") = "Finish Time"
    .Range("B6") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

'Assign start time and end time
StartTime = Workbooks(guiwb).Sheets("Progress").Range("B4").Text
EndTime = Workbooks(guiwb).Sheets("Progress").Range("B6").Text

'Message box with start and end times
MsgBox "The  Roadway Data input file has been successfully created and saved in the working directory folder. The following is a summary of the process:" & Chr(10) & _
Chr(10) & _
"Process: Combine Roadway Data" & Chr(10) & _
"Start Time: " & StartTime & Chr(10) & _
"End Time: " & EndTime & Chr(10), vbOKOnly, "Process Complete"

'Screen Updating ON
Application.ScreenUpdating = True

End Sub

Sub Delete_Labels()
'Delete_Labels macro:
'   (1) Finds which routes are not present in all 5 roadway data sets.
'   (2) Deletes any of those routes in the datasets.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim i, col1, row1, row2, Route, iCount, lCount As Long
Dim AADTRow, FClassRow, SpeedRow, LanesRow, UCodeRow As Long
Dim AADTCol, FClassCol, SpeedCol, LanesCol, UCodeCol As Long
Dim aAADT(1 To 500), aFClass(1 To 500), aSpeed(1 To 500), aLanes(1 To 500), aUCode(1 To 500), DLabels(1 To 200) As Integer
Dim rrow

Sheets("Urban_Code").Activate

rrow = 2

'delete ramps
Do Until Sheets("Urban_Code").Cells(rrow, 1) = ""
    Route = Int(Right(Sheets("Urban_Code").Cells(rrow, 1), 4))
If Route > 491 Then
        rrow = rrow - 1
        Sheets("Urban_Code").Cells(rrow + 1, 1).EntireRow.Delete
        rrow = rrow + 1
    Else
        rrow = rrow + 1
    End If
Loop

    
'Find data extents
'AADT
AADTRow = Worksheets(FN1).Range("A1").End(xlDown).row
AADTCol = Worksheets(FN1).Range("A1").End(xlToRight).Column

'FClass
FClassRow = Worksheets(FN2).Range("A1").End(xlDown).row
FClassCol = Worksheets(FN2).Range("A1").End(xlToRight).Column

'Speed Limit
SpeedRow = Worksheets(FN3).Range("A1").End(xlDown).row
SpeedCol = Worksheets(FN3).Range("A1").End(xlToRight).Column

'Lanes
LanesRow = Worksheets(FN4).Range("A1").End(xlDown).row
LanesCol = Worksheets(FN4).Range("A1").End(xlToRight).Column

'Urban Code
UCodeRow = Worksheets(FN5).Range("A1").End(xlDown).row
UCodeCol = Worksheets(FN5).Range("A1").End(xlToRight).Column


'Sort AADT data by label then by milepoint
ActiveWorkbook.Worksheets(FN1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN1).Sort.SortFields.Add Key:=Range( _
    "C2:C" & AADTRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets(FN1).Sort.SortFields.Add Key:=Range( _
    "D2:D" & AADTRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN1).Sort
    .SetRange Range("A1:" & ConvertToLetter(AADTCol) & AADTRow)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort Functional Class data by route number then by milepoint
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:=Range( _
    "A2:A" & FClassRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:=Range( _
    "B2:B" & FClassRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN2).Sort
    .SetRange Range("A1:" & ConvertToLetter(FClassCol) & FClassRow)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort Speed Limit data by route number then by milepoint
ActiveWorkbook.Worksheets(FN3).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN3).Sort.SortFields.Add Key:=Range( _
    "C2:C" & SpeedRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets(FN3).Sort.SortFields.Add Key:=Range( _
    "D2:D" & SpeedRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN3).Sort
    .SetRange Range("A1:" & ConvertToLetter(SpeedCol) & SpeedRow)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort Thru Lanes data by route number then by milepoint
ActiveWorkbook.Worksheets(FN4).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN4).Sort.SortFields.Add Key:=Range( _
    "C2:C" & LanesRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets(FN4).Sort.SortFields.Add Key:=Range( _
    "D2:D" & LanesRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN4).Sort
    .SetRange Range("A1:" & ConvertToLetter(LanesCol) & LanesRow)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort Urban Code data by route number then by milepoint
ActiveWorkbook.Worksheets(FN5).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN5).Sort.SortFields.Add Key:=Range( _
    "C2:C" & UCodeRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
ActiveWorkbook.Worksheets(FN5).Sort.SortFields.Add Key:=Range( _
    "D2:D" & UCodeRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN5).Sort
    .SetRange Range("A1:" & ConvertToLetter(UCodeCol) & UCodeRow)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Open AADT sheet, find label column, and assign labels to array
Sheets(FN1).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Route = 0
Do Until Cells(row1, col1) = ""         'Assign labels to array
    If Cells(row1, col1) <> Route Then
        Route = Cells(row1, col1)
        aAADT(Route) = Route
    End If
    row1 = row1 + 1
Loop

'Open Functional Class sheet, find label column, and assign labels to array
Sheets(FN2).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Route = 0
Do Until Cells(row1, col1) = ""         'Assign labels to array                 'FLAGGED: aFClass only goes to 500 so crashes if bigger than that.
    If Cells(row1, col1) <> Route Then
        Route = Cells(row1, col1)
        aFClass(Route) = Route
    End If
    row1 = row1 + 1
Loop

'Open Speed Limit sheet, find label column, and assign labels to array
Sheets(FN3).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Route = 0
Do Until Cells(row1, col1) = ""         'Assign labels to array
    If Cells(row1, col1) <> Route Then
        Route = Cells(row1, col1)
        aSpeed(Route) = Route
    End If
    row1 = row1 + 1
Loop

'Open Lanes sheet, find label column, and assign labels to array
Sheets(FN4).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Route = 0
Do Until Cells(row1, col1) = ""         'Assign labels to array
    If Cells(row1, col1) <> Route Then
        Route = Cells(row1, col1)
        aLanes(Route) = Route
    End If
    row1 = row1 + 1
Loop

'Open Urban Code sheet, find label column, and assign labels to array
Sheets(FN5).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Route = 0
Do Until Cells(row1, col1) = ""         'Assign labels to array
    If Cells(row1, col1) <> Route Then
        Route = Cells(row1, col1)
        aUCode(Route) = Route
    End If
    row1 = row1 + 1
Loop

'Check to see which routes are not present in all 5 roadway data sets
iCount = 1
row2 = 4
For i = 1 To 500
    lCount = 0
    If IsEmpty(aAADT(i)) = True Or aAADT(i) = 0 Then
        lCount = lCount + 1
    End If
    If IsEmpty(aFClass(i)) = True Or aFClass(i) = 0 Then
        lCount = lCount + 1
    End If
    If IsEmpty(aSpeed(i)) = True Or aSpeed(i) = 0 Then
        lCount = lCount + 1
    End If
    If IsEmpty(aLanes(i)) = True Or aLanes(i) = 0 Then
        lCount = lCount + 1
    End If
    If IsEmpty(aUCode(i)) = True Or aUCode(i) = 0 Then
        lCount = lCount + 1
    End If
    If lCount > 0 And lCount < 5 Then           'If labels exist in some but not all datasets, delete them
        DLabels(iCount) = i
        Sheets("OtherData").Activate           'List deletion labels
        Cells(row2, 45) = i
        Cells(row2, 46) = lCount
        row2 = row2 + 1
        iCount = iCount + 1
    End If
Next i

'Delete labels
'Open AADT sheet, find route_id column, and delete labels that are not present in all 5 roadway data sets
Sheets(FN1).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Do Until Cells(row1, col1) = ""         'Delete labels
    For i = 1 To 100
        If Cells(row1, col1) = DLabels(i) Then
            Rows(row1).EntireRow.Delete
            row1 = row1 - 1
        End If
    Next i
    row1 = row1 + 1
Loop

'Open Functional Class sheet, find route_id column, and delete labels that are not present in all 5 roadway data sets
Sheets(FN2).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Do Until Cells(row1, col1) = ""         'Delete labels
    For i = 1 To 100
        If Cells(row1, col1) = DLabels(i) Then
            Rows(row1).EntireRow.Delete
            row1 = row1 - 1
        End If
    Next i
    row1 = row1 + 1
Loop

'Open Speed Limit sheet, find route_id column, and delete labels that are not present in all 5 roadway data sets
Sheets(FN3).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Do Until Cells(row1, col1) = ""         'Delete labels
    For i = 1 To 100
        If Cells(row1, col1) = DLabels(i) Then
            Rows(row1).EntireRow.Delete
            row1 = row1 - 1
        End If
    Next i
    row1 = row1 + 1
Loop

'Open Lanes sheet, find route_id column, and delete labels that are not present in all 5 roadway data sets
Sheets(FN4).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Do Until Cells(row1, col1) = ""         'Delete labels
    For i = 1 To 100
        If Cells(row1, col1) = DLabels(i) Then
            Rows(row1).EntireRow.Delete
            row1 = row1 - 1
        End If
    Next i
    row1 = row1 + 1
Loop

'Open Urban Code sheet, find route_id column, and delete labels that are not present in all 5 roadway data sets
Sheets(FN5).Activate            'Open sheet
col1 = 1
Do Until Cells(1, col1) = "ROUTE_ID"       'Find ROUTE_ID column
    col1 = col1 + 1
Loop
row1 = 2
Do Until Cells(row1, col1) = ""         'Delete labels
    For i = 1 To 100
        If Cells(row1, col1) = DLabels(i) Then
            Rows(row1).EntireRow.Delete
            row1 = row1 - 1
        End If
    Next i
    row1 = row1 + 1
Loop

End Sub

Sub Condense_SpeedLimit()
'Condense_SpeedLimit macro:
'   (1) Combines speed limit data where there are consecutive segments listed with the same speed limit.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define variables
Dim row1, row2, Speed_Limit, Label_Count, i As Integer
Dim Mile1, Mile2 As Double
Dim Label As String
Dim m(1 To 500, 1 To 2) As Variant
Dim col1 As Integer

'Sort Functional Class data
Sheets(FN2).Activate
row1 = Range("A1").End(xlDown).row
col1 = Range("A1").End(xlToRight).Column

ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:= _
    Range("A2:A" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:= _
    Range("B2:B" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:= _
    Range("C2:C" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
With ActiveWorkbook.Worksheets(FN2).Sort
    .SetRange Range(Cells(1, 1), Cells(row1, col1))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Find ending milepoint values
row1 = 2
Label = Cells(row1, 3)
Label_Count = 1
Do While Cells(row1, 3) <> ""
    If Cells(row1, 3) <> Label And Cells(row1 - 1, 2) = "P" Then
        m(Label_Count, 1) = Cells(row1 - 1, 3)
        m(Label_Count, 2) = Cells(row1 - 1, 5)
        Label = Cells(row1, 3)
        Label_Count = Label_Count + 1
    End If
    row1 = row1 + 1
Loop

'Activate SPEED_LIMIT sheet
Sheets(FN3).Activate

'Sort information by Label and then by Beginning Milepoint
ActiveWorkbook.Worksheets(FN3).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN3).Sort.SortFields.Add Key:=Range( _
    "C2:C3069"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets(FN3).Sort.SortFields.Add Key:=Range( _
    "D2:D3069"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets(FN3).Sort
    .SetRange Range("A1:G3069")
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Give value to variable
row1 = 2

Do Until Cells(row1, 1) = ""
    
    'Give new values to variables
    row2 = row1
    Label = Cells(row1, 3)
    Speed_Limit = Cells(row1, 6)
    Mile1 = Cells(row1, 5)
    
    'Go through rows until a change is found in the label or speed limit. End row2 on row before change
    Do Until Cells(row2 + 1, 6) <> Speed_Limit Or Cells(row2 + 1, 3) <> Label
        row2 = row2 + 1
    Loop
    
    'Get value of last milepoint
    Mile2 = Cells(row2, 5)
    
    'Check to see if Mile2 is equal to original milepoint
    If Mile2 <> Mile1 Then
        'Delete rows between and correct row2 number
        Application.CutCopyMode = False
        Rows("" & CStr(row1 + 1) & ":" & CStr(row2)).EntireRow.Delete shift:=xlUp
        row2 = row1
        
        'Correct end milepoint value
        Cells(row2, 5) = Mile2
    End If
    row1 = row1 + 1
Loop

'Delete duplicates
row1 = 3
Do While Cells(row1, 6) <> ""
    If row1 = 1300 Then
        row1 = row1
    End If
    row2 = row1 - 1
    If Cells(row1, 4) = Cells(row2, 4) And Cells(row1, 5) = Cells(row2, 5) And Cells(row1, 3) = Cells(row2, 3) Then
        Rows(row1).EntireRow.Delete
    Else
        row1 = row1 + 1
    End If
Loop

End Sub

Sub Condense_ThruLanes()
' Condense_ThruLanes macro:
'   (1) Corrects beginning/ending milepoints in order to create the correct milepoint range for each segment.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Define variables
Dim row1, row2, Num_Lanes As Integer
Dim Mile1, Mile2 As Double
Dim Label As String

'Activate Thru_Lanes sheet
Sheets("Thru_Lanes").Activate

'Sort information by Label and then by Beginning Milepoint
ActiveWorkbook.Worksheets("Thru_Lanes").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Thru_Lanes").Sort.SortFields.Add Key:=Range( _
    "C2:C50000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets("Thru_Lanes").Sort.SortFields.Add Key:=Range( _
    "D2:D50000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets("Thru_Lanes").Sort.SortFields.Add Key:=Range( _
    "E2:E50000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Thru_Lanes").Sort
    .SetRange Range("A1:F50000")
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Correct beginning and ending milepoints if necessary
row1 = 2
Do Until Cells(row1, 1) = ""
    Label = Cells(row1, 3)
    If Cells(row1 + 1, 3) = Label Then
        Cells(row1, 5) = Cells(row1 + 1, 4)
    Else
        'Nothing
    End If
    row1 = row1 + 1
Loop

'Consolidate thru lane data so that there aren't several segments in a row with the same lane number
row1 = 2
Do Until Cells(row1, 1) = ""
    If Cells(row1, 4) = Cells(row1, 5) Then
        Rows(row1).EntireRow.Delete
    Else
        If Cells(row1, 6) = 0 Then
            Cells(row1, 6) = 1
        End If
        row1 = row1 + 1
    End If
Loop

row1 = 2
Do Until Cells(row1, 1) = ""
    'Give new values to variables
    row2 = row1
    Label = Cells(row1, 3)
    Num_Lanes = Cells(row1, 6)
    Mile1 = Cells(row1, 5)
    
    'Go through rows until a change is found in the label or the number of lanes. End row1 on row before change
    Do Until Cells(row2, 6) <> Num_Lanes Or Cells(row2, 3) <> Label
        row2 = row2 + 1
    Loop
    row2 = row2 - 1
    
    'Get value of last milepoint
    Mile2 = Cells(row2, 5)
    
    'Check to see if Mile2 is equal to original milepoint
    If Mile2 <> Mile1 Then
        'Delete rows between and correct row2 number
        Application.CutCopyMode = False
        Rows("" & CStr(row1 + 1) & ":" & CStr(row2)).EntireRow.Delete shift:=xlUp
        row2 = row1
        
        'Correct end milepoint value
        Cells(row2, 5) = Mile2
    End If
    
    row1 = row1 + 1
Loop

End Sub

Sub SameMP()
'SameMP macro:
'   (1) Verifies that each dataset has the same ending milepoints for each route number.
'   (2) All datasets are compared to the values from the Functional Class data.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Dim row1, row2 As Long
Dim Label_Count As Long
Dim i As Long
Dim Label As String
Dim m(1 To 500, 1 To 2) As Variant
Dim col1 As Integer

'Sort Functional Class data by label, then milepoint
Sheets(FN2).Activate
row1 = Range("A1").End(xlDown).row
col1 = Range("A1").End(xlToRight).Column

ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:= _
    Range(Cells(2, 1), Cells(row1, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:= _
    Range(Cells(2, 2), Cells(row1, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
ActiveWorkbook.Worksheets(FN2).Sort.SortFields.Add Key:= _
    Range(Cells(2, 3), Cells(row1, 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
With ActiveWorkbook.Worksheets(FN2).Sort
    .SetRange Range(Cells(1, 1), Cells(row1, col1))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Functional Class: Correct beginning milepoints if necessary
row1 = 2
Do While Cells(row1, 6) <> ""
    row2 = row1 + 1
    Label = Cells(row1, 3)
    
    'If both beginning and ending milepoints are greater than maximum, delete row
    For i = 1 To 500
        If m(i, 1) = Label Then
            If Cells(row1, 4) > m(i, 2) And Cells(row1, 5) > m(i, 2) Then
                Rows(row1).EntireRow.Delete
                row1 = row1 - 1
                row2 = row1 + 1
            End If
            Exit For
        End If
    Next i
    
    'If beginning and ending milepoints are equal, delete the row.
    If Cells(row1, 4) = Cells(row1, 5) Then
        Rows(row1).EntireRow.Delete
    Else
        'Check beginning milepoint
        If Cells(row1 - 1, 3) <> Label Then
            Cells(row1, 4) = 0
        Else
            Cells(row1, 4) = Cells(row1 - 1, 5)
        End If
        row1 = row1 + 1
    End If
Loop

'Save ending milepoint values from the ending milepoints of each route from the Functional Class dataset
row1 = 2
Label = Cells(row1, 3)
Label_Count = 1
Do While Cells(row1, 3) <> ""
    If Cells(row1, 3) <> Label Then
        m(Label_Count, 1) = Cells(row1 - 1, 3)
        m(Label_Count, 2) = Cells(row1 - 1, 5)
        Label = Cells(row1, 3)
        Label_Count = Label_Count + 1
    End If
    row1 = row1 + 1
Loop

'Speed Limit: Correct beginning and ending milepoints if necessary
Sheets(FN3).Activate
row1 = 2
Do While Cells(row1, 6) <> ""
    row2 = row1 + 1
    Label = Cells(row1, 3)
    
    'If both beginning and ending milepoints are greater than maximum, delete row
    For i = 1 To 500
        If m(i, 1) = Label Then
            If Cells(row1, 4) > m(i, 2) And Cells(row1, 5) > m(i, 2) Then
                Rows(row1).EntireRow.Delete
                row1 = row1 - 1
                row2 = row1 + 1
            End If
            Exit For
        End If
    Next i
    
    'If beginning and ending milepoints are equal, delete the row.
    If Cells(row1, 4) = Cells(row1, 5) Then
        Rows(row1).EntireRow.Delete
        If Cells(row1, 3) <> Label Then
            row1 = row1 - 1
        End If
    Else
        'Check beginning milepoint
        If Cells(row1 - 1, 3) <> Label Then
            Cells(row1, 4) = 0
        Else
            Cells(row1, 4) = Cells(row1 - 1, 5)
        End If
        
        'Check ending milepoint
        If Cells(row2, 3) <> Label Then
            For i = 1 To 500
                If m(i, 1) = Label Then
                   Cells(row1, 5) = m(i, 2)
                   Exit For
                End If
            Next i
        End If
        row1 = row1 + 1
    End If
Loop

'AADT: Correct beginning and ending milepoints if necessary
Sheets(FN1).Activate
row1 = 2
Do While Cells(row1, 6) <> ""
    row2 = row1 + 1
    Label = Cells(row1, 3)
        
    'If both beginning and ending milepoints are greater than maximum, delete row
    For i = 1 To 500
        If m(i, 1) = Label Then
            If Cells(row1, 4) > m(i, 2) And Cells(row1, 5) > m(i, 2) Then
                Rows(row1).EntireRow.Delete
                row1 = row1 - 1
                row2 = row1 + 1
            End If
            Exit For
        End If
    Next i
    
    'If beginning and ending milepoints are equal, delete the row.
    If Cells(row1, 4) = Cells(row1, 5) Then
        Rows(row1).EntireRow.Delete
        If Cells(row1, 3) <> Label Then
            row1 = row1 - 1
        End If
    Else
        'Check beginning milepoint
        If Cells(row1 - 1, 3) <> Label Then
            Cells(row1, 4) = 0
        Else
            Cells(row1, 4) = Cells(row1 - 1, 5)
        End If
        
        'Check ending milepoint
        If Cells(row2, 3) <> Label Then
            For i = 1 To 500
                If m(i, 1) = Label Then
                   Cells(row1, 5) = m(i, 2)
                   Exit For
                End If
            Next i
        End If
        row1 = row1 + 1
    End If
Loop

'Lanes: Correct beginning and ending milepoints if necessary
Sheets(FN4).Activate
row1 = 2
Do While Cells(row1, 6) <> ""
    row2 = row1 + 1
    Label = Cells(row1, 3)
    
    'If both beginning and ending milepoints are greater than maximum, delete row
    For i = 1 To 500
        If m(i, 1) = Label Then
            If Cells(row1, 4) > m(i, 2) And Cells(row1, 5) > m(i, 2) Then
                Rows(row1).EntireRow.Delete
                row1 = row1 - 1
                row2 = row1 + 1
            End If
            Exit For
        End If
    Next i
    
    'Check beginning milepoint
    If Cells(row1 - 1, 3) <> Label Then
        Cells(row1, 4) = 0
    Else
        Cells(row1, 4) = Cells(row1 - 1, 5)
    End If
    
    'Check ending milepoint
    If Cells(row2, 3) <> Label Then
        For i = 1 To 500
            If m(i, 1) = Label Then
               Cells(row1, 5) = m(i, 2)
               Exit For
            End If
        Next i
    End If
    row1 = row1 + 1
Loop

'Urban Code: Correct beginning and ending milepoints if necessary
Sheets(FN5).Activate
row1 = 2
Do While Cells(row1, 6) <> ""
    row2 = row1 + 1
    Label = Cells(row1, 3)
    
    'If both beginning and ending milepoints are greater than maximum, delete row
    For i = 1 To 500
        If m(i, 1) = Label Then
            If Cells(row1, 4) > m(i, 2) And Cells(row1, 5) > m(i, 2) Then
                Rows(row1).EntireRow.Delete
                row1 = row1 - 1
                row2 = row1 + 1
            End If
            Exit For
        End If
    Next i
    
    'If beginning and ending milepoints are equal, delete the row.
    If Cells(row1, 4) = Cells(row1, 5) Then
        Rows(row1).EntireRow.Delete
        If Cells(row1, 3) <> Label Then
            row1 = row1 - 1
        End If
    Else
        'Check beginning milepoint
        If Cells(row1 - 1, 3) <> Label Then
            Cells(row1, 4) = 0
        Else
            Cells(row1, 4) = Cells(row1 - 1, 5)
        End If
        
        'Check ending milepoint
        If Cells(row2, 3) <> Label Then
            For i = 1 To 500
                If m(i, 1) = Label Then
                   Cells(row1, 5) = m(i, 2)
                   Exit For
                End If
            Next i
        End If
        row1 = row1 + 1
    End If
Loop

End Sub

Sub CombineData()
'CombineData macro:
'   (1) Combine all roadway datasets onto the Roadway Data sheet.
'
'   Explanation:
'       This macro keeps track of all five roadway datasets. Starting with the lowest route number and the
'   lowest milepoint, it finds the next closest change in roadway characteristics from any of the five
'   datasets. It then creates a segment with all the corresponding characteristics up to that milepoint
'   where the change happened. Then the next change is found, and a segment is created, and so on. The
'   macro works through all routes and milepoint ranges until all segments are made. They are saved onto
'   the "Roadway Data" sheet.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Dim dRow(1 To 6), dCol(1 To 6), myrow, i As Integer
Dim myLabel As String
Dim MinMile, EndMile(1 To 6) As Double
Dim ws(1 To 6) As Worksheet
Dim MileInt As Single
Dim wksht As Worksheet
Dim j As Integer

'Delete old roadway data sheet
Application.DisplayAlerts = False
For Each wksht In Worksheets
    If wksht.Name = "Roadway Data" Then
        wksht.Delete
    End If
Next
Application.DisplayAlerts = True

'Add a combined roadway data worksheet
Worksheets.Add(After:=Worksheets("Home")).Name = "Roadway Data"
Worksheets("Roadway Data").Tab.ColorIndex = 10

'Assign mile interval value
Sheets("Home").Activate
'If ActiveSheet.optEveryChange.Value = True Then
'    MileInt = 10000                                 '10,000 assigned so that changes in characteristics govern.
                                                    '   (Change will always come before 10,000 miles)
'ElseIf ActiveSheet.optLength.Value = True Then
'    MileInt = ActiveSheet.txtSegLen.Value           'MileInt is the selected segment length
'End If

'Assign initial row values
dRow(1) = 2
dRow(2) = 2
dRow(3) = 2
dRow(4) = 2
dRow(5) = 2
dRow(6) = 2
dCol(1) = 1
dCol(2) = 1
dCol(3) = 1
dCol(4) = 1
dCol(5) = 1
dCol(6) = 1
Set ws(1) = Sheets(FN1)
Set ws(2) = Sheets(FN2)
Set ws(3) = Sheets(FN3)
Set ws(4) = Sheets(FN4)
Set ws(5) = Sheets(FN5)
Set ws(6) = Sheets("Roadway Data")

'Find milepoint columns in each sheet
For i = 1 To 5
    Do Until ws(i).Cells(1, dCol(i)) = "BEG_MILEPOINT"
        dCol(i) = dCol(i) + 1
    Loop
Next i

'Round all milepoints to 3 decimal places
For i = 1 To 5
myrow = 2
    Do Until ws(i).Cells(myrow, 1) = ""
        ws(i).Cells(myrow, dCol(i)) = Format(ws(i).Cells(myrow, dCol(i)), "#.000")
        ws(i).Cells(myrow, dCol(i) + 1) = Format(ws(i).Cells(myrow, dCol(i) + 1), "#.000")
        myrow = myrow + 1
    Loop
Next i

'Check each sheet for the next lowest end milepoint, combine data on combined roadway data sheet
'and check for new labels.
myLabel = "FirstString"             'Set label as "FirstString" as a preliminary value.
'Do Loop until route value is blank (meaning the end of each dataset list has been reached).
Do Until ws(1).Cells(dRow(1), 1) = "" And ws(2).Cells(dRow(2), 1) = "" And ws(3).Cells(dRow(3), 1) = "" _
And ws(4).Cells(dRow(4), 1) = "" And ws(5).Cells(dRow(5), 1) = ""
    'If none of the label values equal the previous myLabel value, new label is assigned,
    'beginning label milepoint is zero.
    If ws(1).Cells(dRow(1), 3) <> myLabel And ws(2).Cells(dRow(2), 3) <> myLabel And _
    ws(3).Cells(dRow(3), 3) <> myLabel And ws(4).Cells(dRow(4), 3) <> myLabel And _
    ws(5).Cells(dRow(5), 3) <> myLabel Then
        myLabel = ws(1).Cells(dRow(1), 3)               'Assign new label
        ws(6).Cells(dRow(6), 2) = 0                     'First segment in label, begmp is zero.
    'Else, enter the following data for new segment
    Else
        ws(6).Cells(dRow(6), 1) = ws(1).Cells(dRow(1), 3)                   'Label
        
        'If current route is equal to the previous, then BEGMP is equal to previous ENDMP
        If ws(6).Cells(dRow(6), 1) = ws(6).Cells(dRow(6) - 1, 1) And _
            ws(6).Cells(dRow(6) - 1, 1) <> "" Then
            ws(6).Cells(dRow(6), 2) = ws(6).Cells(dRow(6) - 1, 3)           'Beg_milepoint
        End If
                
        ws(6).Cells(dRow(6), 4) = ws(1).Cells(dRow(1), 1)                   'Route_name
        ws(6).Cells(dRow(6), 5) = ws(1).Cells(dRow(1), 1)                   'Route_ID1
        ws(6).Cells(dRow(6), 6) = ws(1).Cells(dRow(1), 2)                   'Direction1
        ws(6).Cells(dRow(6), 7) = ws(2).Cells(dRow(2), 6)                   'FC_Code
        ws(6).Cells(dRow(6), 8) = ws(2).Cells(dRow(2), 7)                   'FC_Type
        ws(6).Cells(dRow(6), 9) = ws(1).Cells(dRow(1), 6)                   'County
        ws(6).Cells(dRow(6), 10) = ws(1).Cells(dRow(1), 7)                  'Region
        If Sheets("Inputs").Cells(2, 16) = "RSAM" Then
            ws(6).Cells(dRow(6), 11) = ws(1).Cells(dRow(1), 8)                  'AADT_XXX7
            ws(6).Cells(dRow(6), 12) = ws(1).Cells(dRow(1), 9)                  'AADT_XXX6
            ws(6).Cells(dRow(6), 13) = ws(1).Cells(dRow(1), 10)                 'AADT_XXX5
            ws(6).Cells(dRow(6), 14) = ws(1).Cells(dRow(1), 11)                 'AADT_XXX4
            ws(6).Cells(dRow(6), 15) = ws(1).Cells(dRow(1), 12)                 'AADT_XXX3
            ws(6).Cells(dRow(6), 16) = ws(1).Cells(dRow(1), 13)                 'AADT_XXX2
            ws(6).Cells(dRow(6), 17) = ws(1).Cells(dRow(1), 14)                 'AADT_XXX1
            ws(6).Cells(dRow(6), 18) = ws(1).Cells(dRow(1), 15)                 'Single_Per
            ws(6).Cells(dRow(6), 19) = ws(1).Cells(dRow(1), 16)                 'Combo_Perc
            ws(6).Cells(dRow(6), 20) = ws(1).Cells(dRow(1), 17)                 'Single_Count
            ws(6).Cells(dRow(6), 21) = ws(1).Cells(dRow(1), 18)                 'Combo_Count
            ws(6).Cells(dRow(6), 22) = ws(1).Cells(dRow(1), 19)                 'Total_perc
            ws(6).Cells(dRow(6), 23) = ws(1).Cells(dRow(1), 20)                 'Total_Count
            ws(6).Cells(dRow(6), 24) = ws(3).Cells(dRow(3), 6)                  'Speed_limit
            ws(6).Cells(dRow(6), 25) = ws(4).Cells(dRow(4), 6)                  'Num_Lanes
            ws(6).Cells(dRow(6), 26) = ws(5).Cells(dRow(5), 6)                  'Urban_Rural
            ws(6).Cells(dRow(6), 27) = ws(5).Cells(dRow(5), 7)                  'Urban_Ru_1
        ElseIf Sheets("Inputs").Cells(2, 16) = "CAMS" Then
            j = 0 'counter
            Do While Left(ws(1).Cells(1, 8 + j), 4) = "AADT"
                ws(6).Cells(dRow(6), 11 + j) = ws(1).Cells(dRow(1), 8 + j)
                j = j + 1
            Loop
            ws(6).Cells(dRow(6), 10 + j + 1) = ws(1).Cells(dRow(1), 7 + j + 1)          'Single_Per
            ws(6).Cells(dRow(6), 10 + j + 2) = ws(1).Cells(dRow(1), 7 + j + 2)          'Combo_Perc
            ws(6).Cells(dRow(6), 10 + j + 3) = ws(1).Cells(dRow(1), 7 + j + 3)          'Single_Count
            ws(6).Cells(dRow(6), 10 + j + 4) = ws(1).Cells(dRow(1), 7 + j + 4)          'Combo_Count
            ws(6).Cells(dRow(6), 10 + j + 5) = ws(1).Cells(dRow(1), 7 + j + 5)          'Total_perc
            ws(6).Cells(dRow(6), 10 + j + 6) = ws(1).Cells(dRow(1), 7 + j + 6)          'Total_Count
            If dRow(6) = 5156 Then
                dRow(6) = dRow(6) 'test
            End If
            ws(6).Cells(dRow(6), 10 + j + 7) = ws(3).Cells(dRow(3), 6)                  'Speed_limit
            ws(6).Cells(dRow(6), 10 + j + 8) = ws(4).Cells(dRow(4), 6)                  'Num_Lanes
            ws(6).Cells(dRow(6), 10 + j + 9) = ws(5).Cells(dRow(5), 6)                  'Urban_Rural
            ws(6).Cells(dRow(6), 10 + j + 10) = ws(5).Cells(dRow(5), 7)                 'Urban_Ru_1
        End If
        
        'Cycle through 5 datasets to find end milepoint for each data row.
        For i = 1 To 5
            If ws(i).Cells(dRow(i), dCol(i) + 1) = "" Then
                EndMile(i) = 1000
            Else
                EndMile(i) = ws(i).Cells(dRow(i), dCol(i) + 1)
            End If
        Next i
        
        'Find minimum of all 6 milepoints
        MinMile = WorksheetFunction.min(EndMile(1), EndMile(2), EndMile(3), EndMile(4), EndMile(5))
        
        'Assign current segment EndMP to minimum milepoint
        ws(6).Cells(dRow(6), 3) = Format(MinMile, "#.000")
        
        'adjusting for the concurrent routes
        If Sheets("Inputs").Cells(2, 16) = "CAMS" Then
            If ws(6).Cells(dRow(6), 4) = "0006" And ws(6).Cells(dRow(6), 2) = 160.568 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0006"
                ws(6).Cells(dRow(6), 3) = 173.424
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0009" And ws(6).Cells(dRow(6), 2) = 32.662 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0009"
                ws(6).Cells(dRow(6), 3) = 44.771
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0050" And ws(6).Cells(dRow(6), 2) = 119.723 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0050"
                ws(6).Cells(dRow(6), 3) = 129.816
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0080" And ws(6).Cells(dRow(6), 2) = 119.591 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0080"
                ws(6).Cells(dRow(6), 3) = 122.028
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0084" And ws(6).Cells(dRow(6), 2) = 42.716 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0084"
                ws(6).Cells(dRow(6), 3) = 81.043
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0089" And ws(6).Cells(dRow(6), 2) = 191.74 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0089"
                ws(6).Cells(dRow(6), 3) = 225.362
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0089" And ws(6).Cells(dRow(6), 2) = 312.784 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0089"
                ws(6).Cells(dRow(6), 3) = 322.276
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0089" And ws(6).Cells(dRow(6), 2) = 353.061 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0089"
                ws(6).Cells(dRow(6), 3) = 362.547
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0089" And ws(6).Cells(dRow(6), 2) = 389.531 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0089"
                ws(6).Cells(dRow(6), 3) = 395.586
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0089" And ws(6).Cells(dRow(6), 2) = 433.555 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0089"
                ws(6).Cells(dRow(6), 3) = 458.97
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0191" And ws(6).Cells(dRow(6), 2) = 157.193 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0191"
                ws(6).Cells(dRow(6), 3) = 251.434
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0191" And ws(6).Cells(dRow(6), 2) = 294.847 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0191"
                ws(6).Cells(dRow(6), 3) = 352.611
                GoTo skipahead
            ElseIf ws(6).Cells(dRow(6), 4) = "0276" And ws(6).Cells(dRow(6), 2) = 34.99 Then
                For i = 3 To 29
                    ws(6).Cells(dRow(6), i) = ""
                Next i
                ws(6).Cells(dRow(6), 4) = "0276"
                ws(6).Cells(dRow(6), 3) = 53.743
                GoTo skipahead
            End If
        End If
        
        'Cycle through datasets. If EndMile is equal to MinMile then go to next row.
        'Ideally, the very last EndMP of each label for each dataset should be equal. That is assumed.
        For i = 1 To 5
            If EndMile(i) = MinMile Then
                dRow(i) = dRow(i) + 1
            End If
        Next i
        
        'testing by Camille
        If myLabel = "0130P" Then
            myLabel = "0130P"
        End If
skipahead:
        dRow(6) = dRow(6) + 1               'Go to next segment row.
    

    End If
Loop

myrow = 2
Do Until Sheets("Roadway Data").Cells(myrow, 1) = ""
    If Sheets("Roadway Data").Cells(myrow, 5) = "" Then
        Sheets("Roadway Data").Cells(myrow, 5).EntireRow.Delete
        myrow = myrow - 1
    ElseIf Sheets("Roadway Data").Cells(myrow, 2) - Sheets("Roadway Data").Cells(myrow, 3) = 0# Then
        Sheets("Roadway Data").Cells(myrow, 5).EntireRow.Delete
        myrow = myrow - 1
    End If
    myrow = myrow + 1
Loop


End Sub

Sub CombinedHeaders()
'CombinedHeaders macro:
'   (1) Adds headers to the combined roadway data sheet to match the data that has been placed in there.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Sheets("Roadway Data").Activate
Cells(1, 1) = "LABEL"
Cells(1, 2) = "BEG_MILEPOINT"
Cells(1, 3) = "END_MILEPOINT"
Cells(1, 4) = "Route_Name"
Cells(1, 5) = "ROUTE_ID"
Cells(1, 6) = "DIRECTION"
Cells(1, 7) = "FC_Code"
Cells(1, 8) = "FC_Type"
Cells(1, 9) = "COUNTY"
Cells(1, 10) = "REGION"
If Sheets("Inputs").Cells(2, 16) = "RSAM" Then
    Cells(1, 11) = "AADT_" & DYear1
    Cells(1, 12) = "AADT_" & DYear1 - 1
    Cells(1, 13) = "AADT_" & DYear1 - 2
    Cells(1, 14) = "AADT_" & DYear1 - 3
    Cells(1, 15) = "AADT_" & DYear1 - 4
    Cells(1, 16) = "AADT_" & DYear1 - 5
    Cells(1, 17) = "AADT_" & DYear1 - 6
    Cells(1, 18) = "Single_Percent"
    Cells(1, 19) = "Combo_Percent"
    Cells(1, 20) = "Single_Count"
    Cells(1, 21) = "Combo_Count"
    Cells(1, 22) = "Total_Percent"
    Cells(1, 23) = "Total_Count"
    Cells(1, 24) = "SPEED_LIMIT"
    Cells(1, 25) = "Num_Lanes"
    Cells(1, 26) = "Urban_Rural"
    Cells(1, 27) = "Urban_Ru_1"
ElseIf Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    Dim j As Integer, i As Integer
    j = 0 'counter
    Do While Left(Sheets("AADT").Cells(1, 8 + j), 4) = "AADT"
        j = j + 1
    Loop
    For i = 0 To j
        Cells(1, 11 + i) = "AADT_" & 2010 + j - 1 - i
    Next i
    Cells(1, 10 + j + 1) = "Single_Percent"
    Cells(1, 10 + j + 2) = "Combo_Percent"
    Cells(1, 10 + j + 3) = "Single_Count"
    Cells(1, 10 + j + 4) = "Combo_Count"
    Cells(1, 10 + j + 5) = "Total_Percent"
    Cells(1, 10 + j + 6) = "Total_Count"
    Cells(1, 10 + j + 7) = "SPEED_LIMIT"
    Cells(1, 10 + j + 8) = "Num_Lanes"
    Cells(1, 10 + j + 9) = "Urban_Rural"
    Cells(1, 10 + j + 10) = "Urban_Ru_1"
End If

End Sub

Sub FixBlankSpeedLimit()
'Run FixBlankSpeedLimit macro:
'   (1) Corrects speed limits that have a blank or zero value using default speed limits based on functional class.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim row1, row2 As Long                          'Row counters
Dim colFC, colSL, colFC_SL As Integer           'Functional class, speed limit, and functional class-speed limit column

'Assign default values for variables
row1 = 2                                        'Row 2 on Roadway Data sheet
row2 = 4                                        'Row 4 on OtherData sheet
colFC = 1                                       'Column 1 on Roadway Data sheet
colSL = 1                                       'Column 1 on Roadway Data sheet
colFC_SL = 57                                   'Column 57 on the OtherData sheet

'Activate roadway data sheet
Sheets("Roadway Data").Activate

'Find column for FC_Code
Do Until Cells(1, colFC) = "FC_Code"
    colFC = colFC + 1
Loop

'Find column for SPEED_LIMIT
Do Until Cells(1, colSL) = "SPEED_LIMIT"
    colSL = colSL + 1
Loop

'Cycle through each row. If the speed limit is blank or zero, assign default speed limit based on functional class.
Do While Cells(row1, colFC) <> ""
    If Cells(row1, colSL) = "" Or Cells(row1, colSL) = 0 Then
        row2 = 4                                'Row 4 on OtherData sheet
        Do While Sheets("OtherData").Cells(row2, colFC_SL + 2) <> ""
            If Sheets("OtherData").Cells(row2, colFC_SL) = Cells(row1, colFC) Then
                Cells(row1, colSL) = Sheets("OtherData").Cells(row2, colFC_SL + 2)
            End If
            row2 = row2 + 1
        Loop
    End If
    row1 = row1 + 1
Loop

End Sub

Sub CreateInteractions()
'CreateInteractions macro:
'   (1) Add columns and calculate values for "interactions" that are used in the R model.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim irow, iCol As Long              'row and column counters
Dim iAADT, iendmp, ibegmp, iTotPercTr, iSpeedLimit, iLanes As Integer
Dim VMT As Long                     'VMT value that is calculated for each row

'Activate roadway data sheet
Sheets("Roadway Data").Activate

'Find column values for important columns
irow = 1
iCol = 1
Do While Cells(irow, iCol) <> ""
    If Cells(irow, iCol) = "AADT_" & DYear1 Then
        iAADT = iCol
    ElseIf Cells(irow, iCol) = "END_MILEPOINT" Then
        iendmp = iCol
    ElseIf Cells(irow, iCol) = "BEG_MILEPOINT" Then
        ibegmp = iCol
    ElseIf Cells(irow, iCol) = "Total_Percent" Then
        iTotPercTr = iCol
    ElseIf Cells(irow, iCol) = "SPEED_LIMIT" Then
        iSpeedLimit = iCol
    ElseIf Cells(irow, iCol) = "Num_Lanes" Then
        iLanes = iCol
    End If
    iCol = iCol + 1
Loop

'Find last column with data
iCol = 1
Do While Cells(irow, iCol) <> ""
    iCol = iCol + 1
Loop

'Add interaction headers
Cells(irow, iTotPercTr) = "Total_Percent_Trucks"
Cells(irow, iCol) = "VMT"
Cells(irow, iCol + 1) = "VMT2"
Cells(irow, iCol + 2) = "Total_Percent_Trucks2"
Cells(irow, iCol + 3) = "SPEED_LIMIT2"
Cells(irow, iCol + 4) = "Num_Lanes2"
Cells(irow, iCol + 5) = "VMT_PERCENT_TRUCKS"
Cells(irow, iCol + 6) = "VMT_SPEED_LIMIT"
Cells(irow, iCol + 7) = "VMT_NUM_LANES"
Cells(irow, iCol + 8) = "SPEED_LIMIT_NUM_LANES"
Cells(irow, iCol + 9) = "SPEED_LIMIT_PERCENT_TRUCKS"

'Perform interactions one row at a time
irow = 2
Do While Cells(irow, 1) <> ""
    'Set VMT equal to AADT times by difference in milepoints divided by 1000
    Cells(irow, iCol) = Cells(irow, iAADT) * (Cells(irow, iendmp) - Cells(irow, ibegmp)) / 1000
    'Set VMT^2 equal to VMT squared
    Cells(irow, iCol + 1) = Cells(irow, iCol) ^ 2
    'Set Total_Percent_Trucks^2 equal to Total_Percent_Trucks squared
    Cells(irow, iCol + 2) = Cells(irow, iTotPercTr) ^ 2
    'Set SPEED_LIMIT^2 equal to SPEED_LIMIT squared
    Cells(irow, iCol + 3) = Cells(irow, iSpeedLimit) ^ 2
    'Set Num_Lanes^2 equal to Num_Lanes squared
    Cells(irow, iCol + 4) = Cells(irow, iLanes) ^ 2
    'Set VMT_PERCENT_TRUCKS equal to VMT times Total_Percent_Trucks
    Cells(irow, iCol + 5) = Cells(irow, iCol) * Cells(irow, iTotPercTr)
    'Set VMT_SPEED_LIMIT equal to VMT times SPEED_LIMIT
    Cells(irow, iCol + 6) = Cells(irow, iCol) * Cells(irow, iSpeedLimit)
    'Set VMT_NUM_LANES equal to VMT times Num_Lanes
    Cells(irow, iCol + 7) = Cells(irow, iCol) * Cells(irow, iLanes)
    'Set SPEED_LIMIT_NUM_LANES equal to SPEED_LIMIT times Num_Lanes
    Cells(irow, iCol + 8) = Cells(irow, iSpeedLimit) * Cells(irow, iLanes)
    'Set SPEED_LIMIT_PERCENT_TRUCKS equal to SPEED_LIMIT times Total_Percent_Trucks
    Cells(irow, iCol + 9) = Cells(irow, iSpeedLimit) * Cells(irow, iTotPercTr)
    
    irow = irow + 1
Loop

End Sub

Sub Seg_Length()
'Seg_Length macro:
'   (1) Add column for the segment length of each roadway segment.
'   (2) Calculate the segment lengths by taking the difference of the beginning and ending milepoints.
'
'Created by: Marlee Seat, BYU, 2016
'Modified by: Marlee Seat, BYU, 2016
'Commented by: Marlee Seat, BYU, 2016

Dim SegRow As Integer
Dim Length As Double
Dim ws As Worksheet

Sheets("Roadway Data").Activate

Application.Calculation = xlCalculationManual

'Add a new column for Segment Length
Columns("D:D").Insert shift:=xlToRight, _
    CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, 4) = "Seg_Length"

'Calculate Segment Length
SegRow = 2
Do
    Length = Cells(SegRow, 3) - Cells(SegRow, 2)
    Cells(SegRow, 4) = Length
    SegRow = SegRow + 1
Loop While Cells(SegRow, 3) <> ""

''Two incomplete lines at the end of our segmentation are deleted
'If Worksheets("Roadway Data").Cells(SegRow - 1, 1) = "" Then
'    Rows(SegRow - 1).EntireRow.Delete
'End If
'
'If Worksheets("Roadway Data").Cells(SegRow - 2, 1) = "" Then
'    Rows(SegRow - 2).EntireRow.Delete
'End If


If Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    SegRow = 1
    Do Until Sheets("Roadway Data").Cells(SegRow + 1, 1) = ""
        If Sheets("Roadway Data").Cells(SegRow + 1, 4) < 0 Then
            Sheets("Roadway Data").Cells(SegRow + 1, 4).EntireRow.Delete
            SegRow = SegRow - 1
        End If
        SegRow = SegRow + 1
    Loop
End If

Application.Calculation = xlCalculationAutomatic
End Sub


Sub CleanUp()
'CleanUp macro:
'   (1) Combine segments that have the same roadway characteristics.
'
'Created by: Marlee Seat, BYU, 2016
'Modified by: Marlee Seat, BYU, 2016
'Commented by: Marlee Seat, BYU, 2016

Dim SegRow, NewAC As Integer
Dim NewEnd, Length, Mile2, Seg1, Seg2, Measure As Double
Dim NewACType As String
Dim LabelCol As Integer, begAADTcol As Integer, fifthAADTcol As Integer, FCcol As Integer, NumLanesCol As Integer, SLcol As Integer, UCcol As Integer

Application.Calculation = xlCalculationManual

'find columns
LabelCol = 1
Do Until Worksheets("Roadway Data").Cells(1, LabelCol) = "LABEL"
    LabelCol = LabelCol + 1
Loop
begAADTcol = 1
Do Until Left(Worksheets("Roadway Data").Cells(1, begAADTcol), 4) = "AADT"
    begAADTcol = begAADTcol + 1
Loop
fifthAADTcol = begAADTcol
FCcol = 1
Do Until Worksheets("Roadway Data").Cells(1, FCcol) = "FC_Code"
    FCcol = FCcol + 1
Loop
NumLanesCol = 1
Do Until Worksheets("Roadway Data").Cells(1, NumLanesCol) = "Num_Lanes"
    NumLanesCol = NumLanesCol + 1
Loop
SLcol = 1
Do Until Worksheets("Roadway Data").Cells(1, SLcol) = "SPEED_LIMIT"
    SLcol = SLcol + 1
Loop
UCcol = 1
Do Until Worksheets("Roadway Data").Cells(1, UCcol) = "Urban_Rural"
    UCcol = UCcol + 1
Loop


'Combine Segments that have the same Route, 1st year and 5th year AADT, Number of Lanes, FC, Speed Limit, and UC
SegRow = 2
Do
    If Worksheets("Roadway Data").Cells(SegRow, LabelCol) = Worksheets("Roadway Data").Cells(SegRow + 1, LabelCol) And _
    Worksheets("Roadway Data").Cells(SegRow, begAADTcol) = Worksheets("Roadway Data").Cells(SegRow + 1, begAADTcol) And _
    Worksheets("Roadway Data").Cells(SegRow, fifthAADTcol) = Worksheets("Roadway Data").Cells(SegRow + 1, fifthAADTcol) And _
    Worksheets("Roadway Data").Cells(SegRow, NumLanesCol) = Worksheets("Roadway Data").Cells(SegRow + 1, NumLanesCol) And _
    Worksheets("Roadway Data").Cells(SegRow, FCcol) = Worksheets("Roadway Data").Cells(SegRow + 1, FCcol) And _
    Worksheets("Roadway Data").Cells(SegRow, SLcol) = Worksheets("Roadway Data").Cells(SegRow + 1, SLcol) And _
    Worksheets("Roadway Data").Cells(SegRow, UCcol) = Worksheets("Roadway Data").Cells(SegRow + 1, UCcol) Then
    
        NewEnd = Worksheets("Roadway Data").Cells(SegRow + 1, 3).Value
        Worksheets("Roadway Data").Cells(SegRow, 3) = NewEnd
        
        Length = Worksheets("Roadway Data").Cells(SegRow, 3) - Worksheets("Roadway Data").Cells(SegRow, 2)
        Worksheets("Roadway Data").Cells(SegRow, 4) = Length
        
        Worksheets("Roadway Data").Rows(SegRow + 1).EntireRow.Delete
        
    Else: SegRow = SegRow + 1
    End If
Loop Until Cells(SegRow, 1) = ""

'Only goes through Do Loop if choose to segment on every change
SegRow = 2
If Sheets("Inputs").Cells(2, 16) = "RSAM" Then
    Mile2 = form_CreateSegData.txt_MinSegLength.Value
ElseIf Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    Mile2 = form_CreateCAMSData.txt_MinSegLength.Value
End If

Do
Measure = Worksheets("Roadway Data").Cells(SegRow, 4)
    'Combine segments if the row before it has the same AADT
    If Measure < Mile2 Then
    If Worksheets("Roadway Data").Cells(SegRow, 12) = Worksheets("Roadway Data").Cells(SegRow - 1, 12) Then
        Seg1 = Cells(SegRow, 4)
        Seg2 = Cells(SegRow - 1, 4)
        
        If Seg1 > Seg2 Then
            'Change Beginning Milepoint
            NewEnd = Worksheets("Roadway Data").Cells(SegRow - 1, 2)
            Worksheets("Roadway Data").Cells(SegRow, 2) = NewEnd
            Rows(SegRow - 1).EntireRow.Delete
        
            'Recalculate Segment Length
            Length = Worksheets("Roadway Data").Cells(SegRow - 1, 3) - Worksheets("Roadway Data").Cells(SegRow - 1, 2)
            Worksheets("Roadway Data").Cells(SegRow - 1, 4) = Length
        Else
            'Change End Milepoint
            NewEnd = Worksheets("Roadway Data").Cells(SegRow, 3)
            Worksheets("Roadway Data").Cells(SegRow - 1, 3) = NewEnd
            Rows(SegRow).EntireRow.Delete
        
            'Recalculate Segment Length
            Length = Worksheets("Roadway Data").Cells(SegRow - 1, 3) - Worksheets("Roadway Data").Cells(SegRow - 1, 2)
            Worksheets("Roadway Data").Cells(SegRow - 1, 4) = Length
        End If
        
    'Combine if after segemnt has same AADT
    ElseIf Worksheets("Roadway Data").Cells(SegRow, 12) = Worksheets("Roadway Data").Cells(SegRow + 1, 12) Then
        Seg1 = Cells(SegRow, 4)
        Seg2 = Cells(SegRow + 1, 4)
        
        If Seg1 > Seg2 Then
            'Change End Milepoint
            NewEnd = Worksheets("Roadway Data").Cells(SegRow + 1, 3)
            Worksheets("Roadway Data").Cells(SegRow, 3) = NewEnd
            Rows(SegRow + 1).EntireRow.Delete
        
            'Recalculate Segment Length
            Length = Worksheets("Roadway Data").Cells(SegRow, 3) - Worksheets("Roadway Data").Cells(SegRow, 2)
            Worksheets("Roadway Data").Cells(SegRow, 4) = Length
        Else
            'Change Beginning Milepoint
            NewEnd = Worksheets("Roadway Data").Cells(SegRow, 2)
            Worksheets("Roadway Data").Cells(SegRow + 1, 2) = NewEnd
            Rows(SegRow).EntireRow.Delete
        
            'Recalculate Segment Length
            Length = Worksheets("Roadway Data").Cells(SegRow, 3) - Worksheets("Roadway Data").Cells(SegRow, 2)
            Worksheets("Roadway Data").Cells(SegRow, 4) = Length
        End If
        
    Else: SegRow = SegRow + 1
    End If
    
    Else: SegRow = SegRow + 1
    End If
Loop Until Worksheets("Roadway Data").Cells(SegRow, 1) = ""

Application.Calculation = xlCalculationAutomatic
End Sub

Sub MaxTrim()
'MaxTrim macro:
'   (1) Trim segments to lengths less than max length, if the "Max Length" option button is selected.
'
'Created by: Josh Gibbons, BYU, 2016
'Modified by: Josh Gibbons, BYU, 2016
'Commented by: Josh Gibbons, BYU, 2016

'Define variables
Dim maxLength As Single
Dim row1, col1 As Long

'If the "Max Length" option button is selected, trim all segments so they are the max length or less.
If form_CreateSegData.opt_MaxLength.Value = True Then
    maxLength = form_CreateSegData.txt_MaxSegLength.Value
    row1 = 2
    
    Do While Sheets("Roadway Data").Cells(row1, 4) <> ""
        If Sheets("Roadway Data").Cells(row1, 4) > maxLength Then
            Rows(row1 + 1).EntireRow.Insert
            col1 = 1
            Do While Sheets("Roadway Data").Cells(row1, col1) <> ""
                Sheets("Roadway Data").Cells(row1 + 1, col1) = Sheets("Roadway Data").Cells(row1, col1)
                col1 = col1 + 1
            Loop
            Sheets("Roadway Data").Cells(row1, 3) = Sheets("Roadway Data").Cells(row1, 2) + maxLength
            Sheets("Roadway Data").Cells(row1 + 1, 2) = Sheets("Roadway Data").Cells(row1, 2) + maxLength
            
            'Correct segment length
            Sheets("Roadway Data").Cells(row1, 4) = Sheets("Roadway Data").Cells(row1, 3) - Sheets("Roadway Data").Cells(row1, 2)
            Sheets("Roadway Data").Cells(row1 + 1, 4) = Sheets("Roadway Data").Cells(row1 + 1, 3) - Sheets("Roadway Data").Cells(row1 + 1, 2)
        End If
        row1 = row1 + 1
    Loop
    
End If

End Sub

Sub AssignInfo()
'AssignInfo macro:
'   (1) Go through the sheets in the workbook to find the roadway data sheets.
'   (2) Assign sheet names to public variables that can be refered to when working with the datasets.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Dim wksht As Integer
Dim WS_Count As Integer
Dim col1 As Integer

Windows(CurrentWkbk).Activate

WS_Count = ActiveWorkbook.Worksheets.count

For wksht = 1 To WS_Count
    If InStr(1, Sheets(wksht).Name, "AADT") > 0 Then
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

col1 = 1
DYear1 = 1
Do Until Sheets(FN1).Cells(1, col1) = ""
    If Left(Sheets(FN1).Cells(1, col1), 4) = "AADT" Then
        If CInt(Right(Sheets(FN1).Cells(1, col1), 4)) > DYear1 Then
            DYear1 = CInt(Right(Sheets(FN1).Cells(1, col1), 4))
        End If
    End If
    col1 = col1 + 1
Loop

End Sub

Function CountLabels(oRng As Range) As Integer

'MACRO NOT USED IN MAIN VBA WORK. ONLY USED FOR TESTING DURING CODING PHASE. - JG

'CountLabels macro:
'   This macro was used to count how many routes were listed in each dataset. Used for coding phase only.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Dim myrow, mycol As Integer
Dim Label As String

mycol = oRng.Column
myrow = 2
CountLabels = 1

Do While Cells(myrow, mycol) <> ""
    If Cells(myrow, mycol) <> Label Then
        Label = Cells(myrow, mycol)
        CountLabels = CountLabels + 1
    End If
    myrow = myrow + 1
Loop

End Function

Sub List_Labels()

'MACRO NOT USED IN MAIN VBA WORK. ONLY USED FOR TESTING DURING CODING PHASE. - JG

'CountLabels macro:
'   This macro was used to list the route numbers found in each dataset. Used for coding phase only.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Dim col, row1, lCount As Integer
Dim l(1 To 500) As String
Dim Label As String

Sheets(FN1).Activate
col = 1
row1 = 1
lCount = 1
'Find where Labels are
Do Until Cells(row1, col) = "LABEL"
    col = col + 1
Loop
row1 = row1 + 1
Label = "First"
Do Until Cells(row1, col) = ""
    If Cells(row1, col) <> Label Then
        l(lCount) = Cells(row1, col)
        Label = Cells(row1, col)
        lCount = lCount + 1
    End If
    row1 = row1 + 1
Loop
Sheets("Label Test").Activate
Range("A2:A" & UBound(l) + 2) = WorksheetFunction.Transpose(l)
Erase l

Sheets(FN2).Activate
col = 1
row1 = 1
lCount = 1
'Find where Labels are
Do Until Cells(row1, col) = "LABEL"
    col = col + 1
Loop
row1 = row1 + 1
Label = "First"
Do Until Cells(row1, col) = ""
    If Cells(row1, col) <> Label Then
        l(lCount) = Cells(row1, col)
        Label = Cells(row1, col)
        lCount = lCount + 1
    End If
    row1 = row1 + 1
Loop
Sheets("Label Test").Activate
Range("B2:B" & UBound(l) + 2) = WorksheetFunction.Transpose(l)
Erase l

Sheets(FN3).Activate
col = 1
row1 = 1
lCount = 1
'Find where Labels are
Do Until Cells(row1, col) = "LABEL"
    col = col + 1
Loop
row1 = row1 + 1
Label = "First"
Do Until Cells(row1, col) = ""
    If Cells(row1, col) <> Label Then
        l(lCount) = Cells(row1, col)
        Label = Cells(row1, col)
        lCount = lCount + 1
    End If
    row1 = row1 + 1
Loop
Sheets("Label Test").Activate
Range("C2:C" & UBound(l) + 2) = WorksheetFunction.Transpose(l)
Erase l

Sheets(FN4).Activate
col = 1
row1 = 1
lCount = 1
'Find where Labels are
Do Until Cells(row1, col) = "LABEL"
    col = col + 1
Loop
row1 = row1 + 1
Label = "First"
Do Until Cells(row1, col) = ""
    If Cells(row1, col) <> Label Then
        l(lCount) = Cells(row1, col)
        Label = Cells(row1, col)
        lCount = lCount + 1
    End If
    row1 = row1 + 1
Loop
Sheets("Label Test").Activate
Range("D2:D" & UBound(l) + 2) = WorksheetFunction.Transpose(l)
Erase l

Sheets(FN5).Activate
col = 1
row1 = 1
lCount = 1
'Find where Labels are
Do Until Cells(row1, col) = "LABEL"
    col = col + 1
Loop
row1 = row1 + 1
Label = "First"
Do Until Cells(row1, col) = ""
    If Cells(row1, col) <> Label Then
        l(lCount) = Cells(row1, col)
        Label = Cells(row1, col)
        lCount = lCount + 1
    End If
    row1 = row1 + 1
Loop
Sheets("Label Test").Activate
Range("E2:E" & UBound(l) + 2) = WorksheetFunction.Transpose(l)
Erase l

End Sub

Function ConvertToLetter(ByVal lngCol As Long) As String
'ConvertToLetter macro:
'
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016
    
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    ConvertToLetter = vArr(0)
    
End Function


Sub CAMS_SegID()

Dim myrow As Long

'add segment ID column
Sheets("Roadway Data").Columns(1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Sheets("Roadway Data").Cells(1, 1) = "SEG_ID"
Application.CutCopyMode = False

'populate the segment IDs
myrow = 2
Do While Sheets("Roadway Data").Cells(myrow, 2) <> ""
    Sheets("Roadway Data").Cells(myrow, 1) = myrow - 1
    myrow = myrow + 1
Loop

End Sub
