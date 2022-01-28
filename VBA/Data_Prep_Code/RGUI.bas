Attribute VB_Name = "RGUI"
Option Explicit
'R GUI workbook created for UDOT Roadway Safety Analysis Methodology
'Comments by Sam Mineer, Brigham Young Univerisity, June 2016

Function checkRlibraries() As Boolean
' The purpose of this fuction is to check if the required R libraries exist

' Define variables
Dim n As Integer
Dim libfp As String

'Update list as needed. Update array index with changes
Dim Rlibrary(1 To 9) As String
Rlibrary(1) = "MASS"
Rlibrary(2) = "mnormt"
Rlibrary(3) = "rjags"
Rlibrary(4) = "R2WinBUGS"
Rlibrary(5) = "arm"
Rlibrary(6) = "ggplot2"
Rlibrary(7) = "R2jags"
Rlibrary(8) = "gplots"
Rlibrary(9) = "VGAM"

libfp = ActiveWorkbook.Sheets("Inputs").Range("B3").Value
libfp = Left(libfp, InStr(1, libfp, "R/R-") + 9)

'For each library, check if it exists.
checkRlibraries = True
For n = 1 To 9
    If FolderExists(libfp & "library/" & Rlibrary(n)) Then
    Else
        'If one of the given libraries doesn't exist, an R script will be initialized to download the necessary R libraries
        checkRlibraries = False
        Exit For
    End If
Next n

End Function


Sub BAdataprep(analysisfilepath As String, aadtfilepath As String, crashfilepath As String)
' The purpose of this script is to combine the analysis segments, AADT, and crash data for the BeforeAfter analysis

' Define variables
Dim crashFactorList As Object
    Set crashFactorList = CreateObject("Scripting.Dictionary")
Dim sevlist As Object
    Set sevlist = CreateObject("Scripting.Dictionary")
Dim nrow As Integer

'Dim factorarray As Integer
'dim ifactor as integer

'Dim tfilepath As String     'target directory, where the new worksheet will be saved
'Dim ofilepath As String     'origin of road attribute datasets
'Dim nroad As Integer        'row counter, to find directory of the road characteristics
'Dim jroad As Integer        'column counter, to find the directory of the road characteristics
Dim i As Integer            'integer counter
Dim ilib As Integer
Dim n As Integer            'integer counter
Dim j As Long            'integer counter
Dim sname As String         'name of sheet
Dim wname As String         'name of workbook, with roadway characteristics
Dim guiwb As String         'name of gui workbook
Dim dataworkbook As Workbook
Dim crashworkbook As Workbook
Dim aadtworkbook As Workbook
Dim crashwb As String
Dim CrashSheet As String
Dim datafile As String      'used for cycling through files in a directory
Dim crashheading As String
Dim aadtwb As String
Dim AADTSheet As String
Dim segwb As String
Dim segsheet As String
Dim basheet As String

Dim totalcrash As Integer
Dim SevCount As Integer
Dim factorcount As Integer

Dim iroadroute As Integer
Dim iroadbmp As Integer
Dim iroademp As Integer
Dim ichangedate As Integer
Dim itotalcrash As Integer      'New columns of data
Dim isevcrash As Integer     'New columns of data
Dim icrashid As Integer
Dim icrashdt As Integer

Dim icrashroute As Integer
Dim icrashmp As Integer
Dim icrashsev As Integer

Dim iaadtroute As Integer
Dim iaadtbmp As Integer
Dim iaadtemp As Integer

Dim roadroute As String
Dim roadbmp As Double
Dim roademp As Double
Dim roadyear As String
Dim CrashRoute As String
Dim CrashMP As Double
Dim CrashID As String
Dim crashdt As String
Dim CrashSev As String
Dim aadtroute As String
Dim aadtbmp As Double
Dim aadtemp As Double
Dim Route As String

Dim roadrow As Long
Dim roadcol As Long
Dim AADTRow As Long
Dim AADTCol As Long
Dim CrashRow As Long
Dim CrashCol As Long
Dim barow As Long
Dim bacol As Long

Dim crashmax As Integer
Dim crashmin As Integer
Dim crashrange As Integer

' Set the name of the GUI workbook
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Dim bainput As String
bainput = "BAinput"

' Load the progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Workbooks(guiwb).Sheets("Progress")
    .Range("B2") = "Loading Roadway and Crash Files. Please wait."
    .Range("B3") = "Do not close MS Excel. Code running."
    .Range("A2") = ""
    .Range("A3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
End With
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False

' Adds a new sheet for the BeforeAfter data
Workbooks(guiwb).Sheets.Add After:=Sheets("Key")
Workbooks(guiwb).ActiveSheet.Name = bainput
basheet = Workbooks(guiwb).ActiveSheet.Name
With Workbooks(guiwb).Sheets(basheet)
    .Cells(1, 1) = "Seg_Num"
    .Cells(1, 2) = "Label"
    .Cells(1, 3) = "BegMP"
    .Cells(1, 4) = "EndMP"
    .Cells(1, 5) = "Year"
    .Cells(1, 6) = "AADT"
    .Cells(1, 7) = "BA"
    .Cells(1, 8) = "Crash"
End With

'****** Begin Data Preparation ********

' Open the analysis segment data
    Set dataworkbook = Workbooks.Open(analysisfilepath) 'Opens the file in Read Only mode
    segwb = ActiveWorkbook.Name
    segsheet = ActiveSheet.Name
    segwb = Replace(segwb, ".xlsx", "")  'Removes the ".xls" so later codes will work
    segwb = Replace(segwb, ".xls", "")  'Removes the ".xls" so later codes will work
    segwb = Replace(segwb, ".csv", "")  'Removes the ".xls" so later codes will work
' Keeps it open

Workbooks(guiwb).Activate
Call Seg_Check_Headers(3, segsheet, segwb)

' Identifies the column for Label, BegMP, EndMP, Change Date for anaylsis segments
iroadroute = 1
Do Until Workbooks(segwb).Sheets(segsheet).Cells(1, iroadroute) = Sheets("Key").Range("AK4")    '"Label"         '<--- Verify this matches header in dataset
    iroadroute = iroadroute + 1
Loop
iroadbmp = 1
Do Until Workbooks(segwb).Sheets(segsheet).Cells(1, iroadbmp) = Sheets("Key").Range("AK5")    '"Beginning MP"         '<--- Verify this matches header in dataset
    iroadbmp = iroadbmp + 1
Loop
iroademp = 1
Do Until Workbooks(segwb).Sheets(segsheet).Cells(1, iroademp) = Sheets("Key").Range("AK6")    '"Ending MP"         '<--- Verify this matches header in dataset
    iroademp = iroademp + 1
Loop
ichangedate = 1
Do Until Workbooks(segwb).Sheets(segsheet).Cells(1, ichangedate) = Sheets("Key").Range("AK7")    '"Change Date"         '<--- Verify this matches header in dataset
    ichangedate = ichangedate + 1
Loop

' Open crash data file
    Set crashworkbook = Workbooks.Open(crashfilepath) 'Opens the file in Read Only mode
    Application.Wait (Now + TimeValue("00:00:01"))
    crashwb = ActiveWorkbook.Name
    CrashSheet = ActiveSheet.Name
    crashwb = Replace(crashwb, ".xls", "")  'Removes the ".xls" so later codes will work
    crashwb = Replace(crashwb, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
    crashwb = Replace(crashwb, ".csv", "")  'Removes the ".csv" so later codes will work
' Keep the crash workbook open

Workbooks(guiwb).Activate
Call Seg_Check_Headers(4, CrashSheet, crashwb)

' Identifies the column for Label, MP, Crash Severity
icrashid = 1
Do Until Workbooks(crashwb).Sheets(icrashid).Cells(1, icrashid) = Sheets("Key").Range("AO6")    '"CRASH_ID"         '<--- Verify this matches header in dataset
    icrashid = icrashid + 1
Loop
icrashdt = 1
Do Until Workbooks(crashwb).Sheets(icrashid).Cells(1, icrashdt) = Sheets("Key").Range("AO7")    '"CRASH_DATETIME"         '<--- Verify this matches header in dataset
    icrashdt = icrashdt + 1
Loop
icrashroute = 1
Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, icrashroute) = Sheets("Key").Range("AO4")    '"LABEL"         '<--- Verify this matches header in dataset
    icrashroute = icrashroute + 1
Loop
icrashmp = 1
Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, icrashmp) = Sheets("Key").Range("AO5")    '"MILEPOINT"         '<--- Verify this matches header in dataset
    icrashmp = icrashmp + 1
Loop
icrashsev = 1
Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, icrashsev) = Sheets("Key").Range("AO8")    '"CRASH_SEVERITY_ID"         '<--- Verify this matches header in dataset
    icrashsev = icrashsev + 1
Loop

' Open the AADT data file
    Set aadtworkbook = Workbooks.Open(aadtfilepath) 'Opens the file in Read Only mode
    aadtwb = ActiveWorkbook.Name
    AADTSheet = ActiveSheet.Name
    aadtwb = Replace(aadtwb, ".xls", "")  'Removes the ".xls" so later codes will work
    aadtwb = Replace(aadtwb, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
    aadtwb = Replace(aadtwb, ".csv", "")  'Removes the ".csv" so later codes will work
    
    ' Creates a new sheet for the AADT data in the GUI workbook
    Workbooks(aadtwb).Sheets(AADTSheet).Copy After:=Workbooks(guiwb).Sheets("Key")
    ActiveSheet.Name = "AADT"
    AADTSheet = ActiveSheet.Name
Workbooks(aadtwb).Close False   'Finished with the AADT data

Workbooks(guiwb).Activate
Call Seg_Check_Headers(5, AADTSheet, guiwb)

' Formats the copies AADT data to add the Label field
roadcol = 1
Do Until Workbooks(guiwb).Sheets(AADTSheet).Cells(1, roadcol) = Sheets("Key").Range("AS4")    '"ROUTE"         '<--- Verify this matches header in dataset
    roadcol = roadcol + 1
Loop

Workbooks(guiwb).Sheets(AADTSheet).Activate
Columns(roadcol + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Workbooks(guiwb).Sheets(AADTSheet).Cells(1, roadcol + 1) = "LABEL"

' Adds the label for each line of AADT data
AADTRow = 2
Do While Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, roadcol) <> ""
    
    ' Change Route 89A to 11
    If Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, roadcol) = "089A" Then     'Corrects the strange value
        Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, roadcol) = "11"
    End If
    
    Route = CStr(Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, roadcol))
    
    If Len(Route) = 1 Then
        Route = "000" & Route & "P"
    ElseIf Len(Route) = 2 Then
        Route = "00" & Route & "P"
    ElseIf Len(Route) = 3 Then
        Route = "0" & Route & "P"
    Else
        Route = Route & "P"
    End If
    
    Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, roadcol + 1).NumberFormat = "@"
    Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, roadcol + 1) = Route
    
    'If more modifications to the AADT data are needed, such as calculating the Total Percent Trucks, then include that here.
    
    
    AADTRow = AADTRow + 1
Loop

' Identifies the column for Label, BegMP, EndMP in the AADT data
iaadtroute = 1
Do Until Workbooks(guiwb).Sheets(AADTSheet).Cells(1, iaadtroute) = "LABEL"         '<--- Verify this matches header in dataset
    iaadtroute = iaadtroute + 1
Loop
iaadtbmp = 1
Do Until Workbooks(guiwb).Sheets(AADTSheet).Cells(1, iaadtbmp) = Sheets("Key").Range("AS5")    '"BEGMP"         '<--- Verify this matches header in dataset
    iaadtbmp = iaadtbmp + 1
Loop
iaadtemp = 1
Do Until Workbooks(guiwb).Sheets(AADTSheet).Cells(1, iaadtemp) = Sheets("Key").Range("AS6")    '"ENDMP"         '<--- Verify this matches header in dataset
    iaadtemp = iaadtemp + 1
Loop

' Identifies the column for AADT data by year
n = 1
ilib = 0
Do Until Workbooks(guiwb).Sheets(AADTSheet).Cells(1, n) = ""
    If LCase(Left(Workbooks(guiwb).Sheets(AADTSheet).Cells(1, n), 4)) = "aadt" Then
        ilib = ilib + 1
    End If
    n = n + 1
Loop

ReDim iAADT(1 To 2, 0 To ilib) As String    'first row is the column index, second row is the year of aadt data
    i = 1
    n = 0
    Do Until Workbooks(guiwb).Sheets(AADTSheet).Cells(1, i) = ""
        If LCase(Left(Workbooks(guiwb).Sheets(AADTSheet).Cells(1, i), 4)) = "aadt" Then
            iAADT(1, n) = i
            iAADT(2, n) = Replace(LCase(Workbooks(guiwb).Sheets(AADTSheet).Cells(1, i)), "aadt", "")
            n = n + 1
        End If
    i = i + 1
    Loop

' Sort the analysis segment data by Label and BegMP
' Identify extent of data for sorting purposes.
i = Workbooks(segwb).Sheets(segsheet).Range("A1").End(xlToRight).Column + 1
j = Workbooks(crashwb).Sheets(CrashSheet).Range("A1").End(xlDown).row

' Sort the analysis segment data
With Workbooks(segwb).Sheets(segsheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Workbooks(segwb).Sheets(segsheet).Cells(2, iroadroute), Workbooks(segwb).Sheets(segsheet).Cells(j, iroadroute)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Workbooks(segwb).Sheets(segsheet).Cells(2, iroadbmp), Workbooks(segwb).Sheets(segsheet).Cells(j, iroadbmp)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Workbooks(segwb).Sheets(segsheet).Cells(1, 1), Workbooks(segwb).Sheets(segsheet).Cells(j, i))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Sort the AADT data by Label and BegMP
' Identify extent of data for sorting purposes.
i = Workbooks(guiwb).Sheets(AADTSheet).Range("A1").End(xlToRight).Column
j = Workbooks(guiwb).Sheets(AADTSheet).Range("A1").End(xlDown).row

' Sort the AADT data by Label and Beginning MP
With Workbooks(guiwb).Sheets(AADTSheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Workbooks(guiwb).Sheets(AADTSheet).Cells(2, iaadtroute), Workbooks(guiwb).Sheets(AADTSheet).Cells(j, iaadtroute)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Workbooks(guiwb).Sheets(AADTSheet).Cells(2, iaadtbmp), Workbooks(guiwb).Sheets(AADTSheet).Cells(j, iaadtbmp)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Workbooks(guiwb).Sheets(AADTSheet).Cells(1, 1), Workbooks(guiwb).Sheets(AADTSheet).Cells(j, i))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Sort the crash data by Label and MP
' Identify extent of data for sorting purposes.
i = Workbooks(crashwb).Sheets(CrashSheet).Range("A1").End(xlToRight).Column
j = Workbooks(crashwb).Sheets(CrashSheet).Range("A1").End(xlDown).row + 1

' Identifies the range of years in the crash data, to compare to the AADT data
crashmax = CStr(year(CDate(Application.WorksheetFunction.Max(Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(2, icrashdt), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, icrashdt))))))
crashmin = CStr(year(CDate(Application.WorksheetFunction.min(Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(2, icrashdt), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, icrashdt))))))
crashrange = crashmax - crashmin

' Sort the crash data
With Workbooks(crashwb).Sheets(CrashSheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(2, icrashroute), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, icrashroute)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(2, icrashmp), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, icrashmp)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(1, 1), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, i))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'****** End of Data Preparation ********

'****** Begin crash count process for each segment **********

' For each analysis segment, checks if the crash is a before (0) or after (1) crash and averages the AADT data
' Reset values
roadroute = ""
roadbmp = 0#
roademp = 0#
roadyear = ""
CrashRoute = ""
CrashMP = 0#
CrashID = ""
crashdt = ""
CrashSev = ""

' Beginning row of analysis roadway, AADT, crash, and BeforeAfter data
roadrow = 2
CrashRow = 2
AADTRow = 2
barow = 2

' Sets multiple arrays for crash years, crash count (before, after), and average AADt calculation
ReDim CrashYears(0 To crashrange) As Long
For n = 0 To crashrange
    CrashYears(n) = crashmax - n
Next n
ReDim crashcount(0 To 1, 0 To crashrange) As Integer        'Before = 0 After = 1; Years for crash data
ReDim aadtcalc(1 To 2, 0 To crashrange) As Double           '1 = Sum 2 = Count; Years for AADT data


' Identifies the number of rows in the analysis roadway data
i = 1
Do While Workbooks(segwb).Sheets(segsheet).Cells(i + 1, 1) <> ""
    i = i + 1
Loop

' Updates the progress sheet
Workbooks(guiwb).Sheets("Progress").Activate
With Workbooks(guiwb).Sheets("Progress")
    .Range("A2") = "Roadway Segment"
    .Range("A3") = "Crash Row"
    .Range("A4") = "Start Time"
    .Range("A5") = "Update Time"
    .Range("A6") = "End Time"
    .Range("B2") = roadrow - 1
    .Range("C2") = "of " & Str(i - 1) & " analysis segments..."
    .Range("B3") = (CrashRow - 1) / j
    .Range("B3").Style = "Percent"
    .Range("B3").NumberFormat = "0.00%"
    .Range("D3") = ""
    .Range("B4") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

' For each analysis segment, count the before/after crashes for each year and average AADT
Line88: Do While Workbooks(segwb).Sheets(segsheet).Cells(roadrow, 1) <> ""
    ' Identify analysis segment Label, BegMP, EndMP, Year of Change
    roadroute = Workbooks(segwb).Sheets(segsheet).Cells(roadrow, iroadroute)
    roadbmp = CDbl(Workbooks(segwb).Sheets(segsheet).Cells(roadrow, iroadbmp))
    roademp = CDbl(Workbooks(segwb).Sheets(segsheet).Cells(roadrow, iroademp))
    roadyear = Workbooks(segwb).Sheets(segsheet).Cells(roadrow, ichangedate)
    
    'Debug checker
    If CrashRoute = "0080P" Then
        CrashRoute = CrashRoute
    End If
    
    ' Reset crash and AADT values
    For n = 0 To crashrange
        crashcount(0, n) = 0
        crashcount(1, n) = 0
    Next n
        For n = 0 To crashrange     'reset aadt calc
        aadtcalc(1, n) = 0
        aadtcalc(2, n) = 0
    Next n
    
    ' For the given analysis roadway segment, count number of crashes and calculate average AADT
    Do While Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, 1) <> ""
        ' Set the crash label and MP
        CrashRoute = Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashroute)
        CrashMP = CDbl(Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashmp))
                    
        If (CrashRoute = roadroute) And (CrashMP >= roadbmp) And (CrashMP <= roademp) Then      'Crash belongs to the segment
            ' Set AADT data
            aadtroute = Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, iaadtroute)
            aadtbmp = CDbl(Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, iaadtbmp))
            aadtemp = CDbl(Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, iaadtemp))
                        
            ' If AADT is behind the crash data, prompt the AADT rows to catch up until matchin
            Do Until (aadtroute = CrashRoute) And ((CrashMP >= aadtbmp) And (CrashMP <= aadtemp))
                AADTRow = AADTRow + 1
                aadtroute = Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, iaadtroute)
                aadtbmp = CDbl(Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, iaadtbmp))
                aadtemp = CDbl(Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, iaadtemp))
            Loop
            
            ' Add crash segment to BA input sheet
            crashdt = CDate(Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashdt))
            CrashCol = crashmax - year(crashdt)
            
            'Check if before (0) or after (1)
            If CDate(roadyear) >= CDate(crashdt) Then
                crashcount(0, CrashCol) = crashcount(0, CrashCol) + 1   'Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = "0"      'Before
            Else
                crashcount(1, CrashCol) = crashcount(1, CrashCol) + 1    'Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = "1"      'After
            End If
                        
            ' Sum the AADT data to calculate average AADT later
            For n = 0 To ilib
                If year(crashdt) = iAADT(2, n) Then
                    AADTCol = iAADT(1, n)
                    aadtcalc(1, CrashCol) = aadtcalc(1, CrashCol) + Workbooks(guiwb).Sheets(AADTSheet).Cells(AADTRow, AADTCol)
                    aadtcalc(2, CrashCol) = aadtcalc(2, CrashCol) + 1
                        'Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = Workbooks(guiwb).Sheets(aadtsheet).Cells(aadtrow, aadtcol)
                    Exit For
                End If
            Next n
            
            'Other calculations for the roadway data should be included here, such as Total Percent Trucks, etc.
            
            
            ' Increase crash row by 1
            CrashRow = CrashRow + 1
            
            ' Periodically update the progress screen
            ' Update more frequently when summarizing the crash factors
            If ((CrashRow Mod 2500) = 0) Then
                Workbooks(guiwb).Sheets("Progress").Range("B3") = (CrashRow - 1) / j
                Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
                Workbooks(guiwb).Sheets("Progress").Range("B5") = Time
                Workbooks(guiwb).Sheets("Progress").Activate
                Application.ScreenUpdating = True
                'Application.Wait (Now + TimeValue("00:00:01"))
                Application.ScreenUpdating = False
            End If
            
         ElseIf CrashRoute = roadroute And CrashMP < roadbmp Then    'Same route, crash milepoint behind segment, need to increase crash count
               
            ' Increase crash row by 1
            CrashRow = CrashRow + 1
            
            ' Periodically update the progress screen
            ' Update more frequently when summarizing the crash factors
            If ((CrashRow Mod 2500) = 0) Then
                Workbooks(guiwb).Sheets("Progress").Range("B3") = (CrashRow - 1) / j
                Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
                Workbooks(guiwb).Sheets("Progress").Range("B5") = Time
                Workbooks(guiwb).Sheets("Progress").Activate
                Application.ScreenUpdating = True
                'Application.Wait (Now + TimeValue("00:00:01"))
                Application.ScreenUpdating = False
            End If
            
        ElseIf CrashRoute = roadroute And CrashMP > roademp Then    'Same route, crash ahead of current roadway, need to go to next road segment
               
            ' Print results for the analysis segment before moving to next analysis segment
            For n = 0 To crashrange
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) = roadrow - 1  '"Seg_Num"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 2) = roadroute    '"Label"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 3) = roadbmp      '"BegMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 4) = roademp      '"EndMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 5) = CrashYears(n)    '"Year"
                 If aadtcalc(2, n) = 0 Then
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = 0 '"AADT"
                 Else
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
                 End If
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = 0
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = crashcount(0, n)  '"Before Crashes"
                 barow = barow + 1
             Next n
             For n = 0 To crashrange
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) = roadrow - 1  '"Seg_Num"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 2) = roadroute    '"Label"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 3) = roadbmp      '"BegMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 4) = roademp      '"EndMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 5) = CrashYears(n)    '"Year"
                 If aadtcalc(2, n) = 0 Then
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = 0 '"AADT"
                 Else
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
                 End If
                 'Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = 1
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = crashcount(1, n)  '"After Crashes"
                 barow = barow + 1
             Next n

            ' Increase analysis row counter by 1
            roadrow = roadrow + 1
            
            Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
            Workbooks(guiwb).Sheets("Progress").Range("B3") = (CrashRow - 1) / j
            Workbooks(guiwb).Sheets("Progress").Activate
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
            
            GoTo Line88     'Returns to beginning of DO LOOP for new road segment
                       
        ElseIf CInt(Left(CrashRoute, 4)) < CInt(Left(roadroute, 4)) Then 'If the crash route is less than the route number, then increase the crash row to catch up to roadway
        
            ' Increase the crash row counter by 1
            CrashRow = CrashRow + 1
            
            ' Periodically update the progress screen
            ' Update more frequently when summarizing the crash factors
            If ((CrashRow Mod 2500) = 0) Then
                Workbooks(guiwb).Sheets("Progress").Range("B3") = (CrashRow - 1) / j
                Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
                Workbooks(guiwb).Sheets("Progress").Range("B5") = Time
                Workbooks(guiwb).Sheets("Progress").Activate
                Application.ScreenUpdating = True
                'Application.Wait (Now + TimeValue("00:00:01"))
                Application.ScreenUpdating = False
            End If
               
        ElseIf (CInt(Left(CrashRoute, 4)) > CInt(Left(roadroute, 4))) Or (CInt(Left(CrashRoute, 4)) = CInt(Left(roadroute, 4)) And Right(roadroute, 1) <> Right(CrashRoute, 1)) Then 'And roadbmp = 0 Then     'new road segment. Print results and move to next segment key.
        ' If the crash route is ahead of the road segment OR the same route but different direction (X<>P), then go to next roadway
            
            ' Print results for the analysis segment before moving to next analysis segment
            For n = 0 To crashrange
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) = roadrow - 1  '"Seg_Num"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 2) = roadroute    '"Label"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 3) = roadbmp      '"BegMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 4) = roademp      '"EndMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 5) = CrashYears(n)    '"Year"
                 If aadtcalc(2, n) = 0 Then
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = 0 '"AADT"
                 Else
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
                 End If
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = 0
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = crashcount(0, n)  '"Before Crashes"
                 barow = barow + 1
             Next n
             For n = 0 To crashrange
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) = roadrow - 1  '"Seg_Num"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 2) = roadroute    '"Label"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 3) = roadbmp      '"BegMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 4) = roademp      '"EndMP"
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 5) = CrashYears(n)    '"Year"
                 If aadtcalc(2, n) = 0 Then
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = 0 '"AADT"
                 Else
                    Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
                 End If
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = 1
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = crashcount(1, n)  '"After Crashes"
                 barow = barow + 1
             Next n

            ' Increase analysis row counter by 1
            roadrow = roadrow + 1
            
            Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
            Workbooks(guiwb).Sheets("Progress").Range("B3") = (CrashRow - 1) / j
            Workbooks(guiwb).Sheets("Progress").Activate
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
            
            GoTo Line88     'Returns to beginning of DO LOOP for new road segment
               
        End If
    Loop
    ' No more crash data to summarize. Print the results.
    ' Print results for the analysis segment before moving to next analysis segment
        For n = 0 To crashrange
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) = roadrow - 1  '"Seg_Num"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 2) = roadroute    '"Label"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 3) = roadbmp      '"BegMP"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 4) = roademp      '"EndMP"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 5) = CrashYears(n)    '"Year"
             If aadtcalc(2, n) = 0 Then
               Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = 0 '"AADT"
            Else
               Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
            End If
                 Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = 0
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = crashcount(0, n)  '"Before Crashes"
             barow = barow + 1
         Next n
         For n = 0 To crashrange
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) = roadrow - 1  '"Seg_Num"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 2) = roadroute    '"Label"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 3) = roadbmp      '"BegMP"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 4) = roademp      '"EndMP"
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 5) = CrashYears(n)    '"Year"
             If aadtcalc(2, n) = 0 Then
               Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = 0 '"AADT"
            Else
               Workbooks(guiwb).Sheets(basheet).Cells(barow, 6) = CLng(aadtcalc(1, n) / aadtcalc(2, n)) '"AADT"
            End If
            Workbooks(guiwb).Sheets(basheet).Cells(barow, 7) = 1
             Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = crashcount(1, n)  '"After Crashes"
             barow = barow + 1
         Next n

    ' Increase road row counter by 1
    roadrow = roadrow + 1
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue("00:00:02"))
    Application.ScreenUpdating = False
Loop

' Updates the progress sheet
Workbooks(guiwb).Sheets("Progress").Activate
With Workbooks(guiwb).Sheets("Progress")
    .Range("C3") = "...adding AADT data"
End With
Application.ScreenUpdating = True
'Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False

' Close the crash data and the analysis segment data
Workbooks(crashwb).Close False
Workbooks(segwb).Close False

' Removes the blank number of crashes in the BeforeAfter data
barow = 2
Workbooks(guiwb).Sheets(basheet).Activate
Do While Workbooks(guiwb).Sheets(basheet).Cells(barow, 1) <> ""
    ' If the number of crashes is blank, then delete the row
    If Workbooks(guiwb).Sheets(basheet).Cells(barow, 8) = 0 Then
        Rows(barow).Delete shift:=xlUp
    Else
        barow = barow + 1
    End If
Loop

' Delete the AADT sheet, a temporary file
Application.DisplayAlerts = False
If SheetExists("AADT") Then
    Sheets("AADT").Delete
End If
Application.DisplayAlerts = True

' Updates the progress screen with the end time
Workbooks(guiwb).Sheets("Progress").Range("B2") = i - 1
Workbooks(guiwb).Sheets("Progress").Range("B3") = 1
Workbooks(guiwb).Sheets("Progress").Range("B6") = Time

' Moves the BeforeAfter sheet to a new workbook, to be saved as a CSV file
' Unique date and time given to file
Workbooks(guiwb).Sheets(bainput).Move
sname = "BAinput " & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".csv"
sname = Replace(sname, " ", "_")
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("F8").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")

Workbooks(guiwb).Activate
Workbooks(guiwb).Sheets("Home").Activate
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))


' Bring up the GUI to start up the BeforeAfter model
form_bamodel.Show

End Sub


Sub UCPSMdataprep(segmentfilepath As String, crashfilepath As String, factorsummary As Boolean, sev1 As Boolean, sev2 As Boolean, sev3 As Boolean, sev4 As Boolean, sev5 As Boolean)
' The purpose of this script is to combine the roadway and crash data for the UCPM/UCSM statistical analysis. The input file will be the same for both analyses
' segmentfilepath   = File path to the roadway segment data file
' crashfilepath     = File path to the historic crash data
' factorsummary     = Boolean (T/F) value determining whether the crash factors from the rollup data will be summarized or not
' sev1 to sev5      = Boolean (T/F) value determining whether to include the given crash severity in the input data file summary

' Define variables
Dim crashFactorList As Object
    Set crashFactorList = CreateObject("Scripting.Dictionary")
Dim sevlist As Object
    Set sevlist = CreateObject("Scripting.Dictionary")
Dim sevlist2 As Object
    Set sevlist2 = CreateObject("Scripting.Dictionary")
Dim sevstring As String
Dim sevstring2 As String

Dim i As Integer            ' Integer counter
Dim ilib As Integer         ' Counter for number of items in the library
Dim n As Integer            ' Integer counter
Dim j As Long            'integer counter
Dim sname As String         'name of sheet
Dim wname As String         'name of workbook, with roadway characteristics
Dim guiwb As String         'name of gui workbook
Dim dataworkbook As Workbook
Dim crashworkbook As Workbook
Dim crashwb As String
Dim CrashSheet As String
Dim datafile As String      'used for cycling through files in a directory
Dim crashheading As String
Dim ucpsmsheet As String

Dim totalcrash As Integer
Dim SevCount As Integer
Dim factorcount As Integer

Dim iroadroute As Integer
Dim iroadbmp As Integer
Dim iroademp As Integer

Dim itotalcrash As Integer      'New columns of data
Dim isevcrash As Integer

Dim icrashroute As Integer
Dim icrashmp As Integer
Dim icrashsev As Integer

Dim roadroute As String
Dim roadbmp As Double
Dim roademp As Double
Dim CrashRoute As String
Dim CrashMP As Double

Dim roadrow As Long
Dim roadcol As Integer
Dim CrashRow As Long
Dim CrashCol As Long


' Set the name of the GUI workbook
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

'update by Camille on July 30, 2019
If Sheets("Inputs").Cells(2, 16) = "RSAM" Then
    ucpsmsheet = "UCPSMinput"
ElseIf Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    ucpsmsheet = "CAMSinput"
End If

' Load the progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Workbooks(guiwb).Sheets("Progress")
    .Range("B2") = "Loading Roadway and Crash Files. Please wait."
    .Range("B3") = "Do not close MS Excel. Code running."
    .Range("A2") = ""
    .Range("A3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
End With
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False

' Create a library of crash factors from key sheet
ilib = 1
Do While Workbooks(guiwb).Sheets("Key").Cells(ilib + 7, 31) <> ""
    ilib = ilib + 1
Loop
ReDim crashFactorArray(1 To ilib) As String
For n = 1 To ilib
    crashFactorArray(n) = Workbooks(guiwb).Sheets("Key").Cells(n + 6, 31)
Next n
For n = 1 To ilib
    crashFactorList.Add crashFactorArray(n), n
Next n

' Create an array to count the crash factors, in order of the library list
ReDim factorarray(1 To ilib) As Integer
ReDim ifactor(1 To ilib) As Integer

' Create the list of severities to check for in the crash data
sevstring = ""
i = 1
If sev1 Then
    sevlist.Add "1", i
    i = i + 1
    sevstring = sevstring & "1"
End If
If sev2 Then
    sevlist.Add "2", i
    i = i + 1
    sevstring = sevstring & "2"
End If
If sev3 Then
    sevlist.Add "3", i
    i = i + 1
    sevstring = sevstring & "3"
End If
If sev4 Then
    sevlist.Add "4", i
    i = i + 1
    sevstring = sevstring & "4"
End If
If sev5 Then
    sevlist.Add "5", i
    i = i + 1
    sevstring = sevstring & "5"
End If

sevstring2 = "345"
sevlist2.Add "3", 1
sevlist2.Add "4", 2
sevlist2.Add "5", 3


'****** Begin Data Preparation ********

' Open and copy roadway segment data to the GUI workbook
    Set dataworkbook = Workbooks.Open(segmentfilepath) 'Opens the file in Read Only mode
    wname = ActiveWorkbook.Name
    sname = ActiveSheet.Name
    wname = Replace(wname, ".xls", "")  'Removes the ".xls" so later codes will work
    wname = Replace(wname, ".xlsx", "")  'Removes the ".xlsx" so later codes will work   'Camille: You might want to move the .xlsx line to be above the .xls line. What if it's .xlsx and .xls gets deleted leaving a .x?
    wname = Replace(wname, ".csv", "")  'Removes the ".csv" so later codes will work
        
    'Creates a new sheet, names it after the name of the workbook
    Sheets(sname).Copy After:=Workbooks(guiwb).Sheets("Key")
    ActiveSheet.Name = ucpsmsheet
    sname = ActiveSheet.Name
Workbooks(wname).Close False    'Closes the workbook, after the data has been copied

Workbooks(guiwb).Activate

Call Seg_Check_Headers(1, sname, guiwb)

' Identifies the column for Label, BegMP, EndMP in the roadway data
iroadroute = 1
Do Until Workbooks(guiwb).Sheets(sname).Cells(1, iroadroute) = Sheets("Key").Range("AC4")         '<--- Verify this matches header in dataset
    iroadroute = iroadroute + 1
Loop
iroadbmp = 1
Do Until Workbooks(guiwb).Sheets(sname).Cells(1, iroadbmp) = Sheets("Key").Range("AC5")         '<--- Verify this matches header in dataset
    iroadbmp = iroadbmp + 1
Loop
iroademp = 1
Do Until Workbooks(guiwb).Sheets(sname).Cells(1, iroademp) = Sheets("Key").Range("AC6")         '<--- Verify this matches header in dataset
    iroademp = iroademp + 1
Loop

' Adds new columns to the end of the data
i = 1
Do
    i = i + 1
Loop While Workbooks(guiwb).Sheets(sname).Cells(1, i + 1) <> ""
itotalcrash = i + 1
isevcrash = i + 2
Workbooks(guiwb).Sheets(sname).Cells(1, itotalcrash) = "Total_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, isevcrash) = "Severe_Crashes"

' Statistical interactions have already been calcuatled in the Pre-Model Data Processing Steps

' Open the crash data
    Set crashworkbook = Workbooks.Open(crashfilepath) 'Opens the file in Read Only mode
    crashwb = ActiveWorkbook.Name
    CrashSheet = ActiveSheet.Name
    crashwb = Replace(crashwb, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
    crashwb = Replace(crashwb, ".xls", "")  'Removes the ".xls" so later codes will work
    crashwb = Replace(crashwb, ".csv", "")  'Removes the ".csv" so later codes will work
' Keep the crash workbook open

Workbooks(guiwb).Activate

Call Seg_Check_Headers(2, CrashSheet, crashwb)

' Identifies the column for Label, MP, Crash Severity
icrashroute = 1
Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, icrashroute) = Sheets("Key").Range("AG4")         '<--- Verify this matches header in dataset
    icrashroute = icrashroute + 1
Loop
icrashmp = 1
Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, icrashmp) = Sheets("Key").Range("AG5")         '<--- Verify this matches header in dataset
    icrashmp = icrashmp + 1
Loop
icrashsev = 1
Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, icrashsev) = Sheets("Key").Range("AG6")         '<--- Verify this matches header in dataset
    icrashsev = icrashsev + 1
Loop

' If the crash rollup data is to be summarized, those columns are added to the end of the dataset
If factorsummary Then
    For n = 1 To ilib
        ifactor(n) = 1
        Do Until Workbooks(crashwb).Sheets(CrashSheet).Cells(1, ifactor(n)) = crashFactorArray(n)
            ifactor(n) = ifactor(n) + 1
        Loop
        Workbooks(guiwb).Sheets(sname).Cells(1, isevcrash + n) = crashFactorArray(n)
    Next n
End If

' Sort the segment data by Label and BegMP
' Identify extent of data for sorting purposes.
i = Workbooks(guiwb).Sheets(ucpsmsheet).Range("A1").End(xlToRight).Column + 1
j = Workbooks(guiwb).Sheets(ucpsmsheet).Range("A1").End(xlDown).row + 1

' Sort the segment data
With Workbooks(guiwb).Sheets(ucpsmsheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Workbooks(guiwb).Sheets(ucpsmsheet).Cells(2, iroadroute), Workbooks(guiwb).Sheets(ucpsmsheet).Cells(j, iroadroute)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Workbooks(guiwb).Sheets(ucpsmsheet).Cells(2, iroadbmp), Workbooks(guiwb).Sheets(ucpsmsheet).Cells(j, iroadbmp)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Workbooks(guiwb).Sheets(ucpsmsheet).Cells(1, 1), Workbooks(guiwb).Sheets(ucpsmsheet).Cells(j, i))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Workbooks(guiwb).Sheets("Home").Activate


' Sort the crash data by Label and MP
' Identify the extent of the crash data
i = Workbooks(crashwb).Sheets(CrashSheet).Range("A1").End(xlToRight).Column
j = Workbooks(crashwb).Sheets(CrashSheet).Range("A1").End(xlDown).row

' Sort the crash data
With Workbooks(crashwb).Sheets(CrashSheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(2, icrashroute), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, icrashroute)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(2, icrashmp), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, icrashmp)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Workbooks(crashwb).Sheets(CrashSheet).Cells(1, 1), Workbooks(crashwb).Sheets(CrashSheet).Cells(j, i))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Workbooks(guiwb).Activate

'****** End of Data Preparation ********

'****** Begin crash count process for each segment **********

' For each segment in the input file, count the severity and look at crash factors
' Reset values
roadroute = ""
roadbmp = 0#
roademp = 0#
CrashRoute = ""
CrashMP = 0#

' Begin count of roadway data and crash data
roadrow = 2
CrashRow = 2

' Identifies the number of rows in the roadway data
i = 1
Do While Workbooks(guiwb).Sheets(ucpsmsheet).Cells(i + 1, 1) <> ""
    i = i + 1
Loop

' Updates the progress sheet
Workbooks(guiwb).Sheets("Progress").Activate
Application.ScreenUpdating = True
With Workbooks(guiwb).Sheets("Progress")
    .Range("A2") = "Assigning crashes to roadway segments. Please wait."
    .Range("A3") = "Progress:"
    .Range("A4") = "Start Time"
    .Range("A5") = "Update Time"
    .Range("A6") = "End Time"
    .Range("B2") = ""
    .Range("B3") = roadrow - 1
    .Range("C2") = ""
    .Range("C3") = "of " & Str(i) & " segments..."
    .Range("D3") = ""
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:03"))
Application.ScreenUpdating = False

' The code will step through each line of the crash and roadway data.
Line88: Do While Workbooks(guiwb).Sheets(sname).Cells(roadrow, 1) <> ""
    ' Identify road Label, BegMP, EndMP
    roadroute = Workbooks(guiwb).Sheets(sname).Cells(roadrow, iroadroute)
    roadbmp = CDbl(Workbooks(guiwb).Sheets(sname).Cells(roadrow, iroadbmp))
    roademp = CDbl(Workbooks(guiwb).Sheets(sname).Cells(roadrow, iroademp))
    
    ' Debugging check
    If roadroute = "0006N" Or roadroute = "0006P" Then
        roadroute = roadroute
    End If
    If CrashRow = 120724 Then
        CrashRow = CrashRow
    End If
    
    
    ' Reset crash count values
    SevCount = 0
    totalcrash = 0
    If factorsummary Then
        For n = 1 To ilib
            factorarray(n) = 0
        Next n
    End If
       
    ' For the given roadway segment, check if the crash is within the segment.
    Do While Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, 1) <> ""
        ' Set the crash label and MP
        CrashRoute = Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashroute)
        CrashMP = CDbl(Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashmp))
        
        'debugging check
        If CrashRow = 186497 Then
            CrashRow = CrashRow
        End If
        
        
        If (CrashRoute = roadroute) And (CrashMP >= roadbmp) And (CrashMP <= roademp) Then  'Crash belongs on the segment
            ' Check if crash severity matches user input
            If sevlist.Exists(CStr(Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashsev))) Then

                If sevlist2.Exists(CStr(Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, icrashsev))) Then
                    SevCount = SevCount + 1
                End If
                            
                ' If crash rollup to be summarized, then count the "Y" values for each of the crash factors
                If factorsummary Then
                    For n = 1 To ilib
                        CrashCol = ifactor(n)   ' Column for crash factor already identified
                        If (Workbooks(crashwb).Sheets(CrashSheet).Cells(CrashRow, CrashCol) = "Y") Then
                            factorarray(n) = factorarray(n) + 1
                        End If
                    Next n
                End If
                
                ' Increase total crash count for segment
                totalcrash = totalcrash + 1
            
            End If
            
            ' Increase the crash row counter by 1
            CrashRow = CrashRow + 1
            
        ElseIf CrashRoute = roadroute And CrashMP > roademp Then    'Same route, crash is for the next segment
             
            ' Print the crash count before moving to the next road segment
            Workbooks(guiwb).Sheets(sname).Cells(roadrow, itotalcrash) = totalcrash
            Workbooks(guiwb).Sheets(sname).Cells(roadrow, isevcrash) = SevCount
            ' If crash rollup data being summarized, print those results
            If factorsummary Then
                For n = 1 To ilib
                    Workbooks(guiwb).Sheets(sname).Cells(roadrow, isevcrash + n) = factorarray(n)
                Next n
            End If
            
            ' Increase road row counter by 1
            roadrow = roadrow + 1
            If (roadrow - 1) Mod 250 = 0 Then
                Application.ScreenUpdating = True
                Workbooks(guiwb).Sheets("Progress").Range("B3") = roadrow - 1
                Application.Wait (Now + TimeValue("00:00:02"))
                Application.ScreenUpdating = False
            End If
                        
            GoTo Line88     'Returns to beginning of DO LOOP for new road segment
                       
        ElseIf CInt(Left(CrashRoute, 4)) < CInt(Left(roadroute, 4)) Then 'If the crash route is less than the route number, then increase the crash row to catch up to roadway
        
            ' Increase the crash row counter by 1
            CrashRow = CrashRow + 1
               
        ElseIf (CInt(Left(CrashRoute, 4)) > CInt(Left(roadroute, 4))) Or (CInt(Left(CrashRoute, 4)) = CInt(Left(roadroute, 4)) And Right(roadroute, 1) <> Right(CrashRoute, 1)) Then 'And roadbmp = 0 Then     'new road segment. Print results and move to next segment key.
        ' If the crash route is ahead of the road segment OR the same route but different direction (X<>P), then go to next roadway
            
            ' Print the crash count before moving to the next road segment
            Workbooks(guiwb).Sheets(sname).Cells(roadrow, itotalcrash) = totalcrash
            Workbooks(guiwb).Sheets(sname).Cells(roadrow, isevcrash) = SevCount
            ' If crash rollup data being summarized, print those results
            If factorsummary Then
                For n = 1 To ilib
                    Workbooks(guiwb).Sheets(sname).Cells(roadrow, isevcrash + n) = factorarray(n)
                Next n
            End If
            
            ' Increase road row counter by 1
            roadrow = roadrow + 1
            If (roadrow - 1) Mod 250 = 0 Then
                Application.ScreenUpdating = True
                Workbooks(guiwb).Sheets("Progress").Range("B3") = roadrow - 1
                Application.Wait (Now + TimeValue("00:00:02"))
                Application.ScreenUpdating = False
            End If
                        
            GoTo Line88     'Returns to beginning of DO LOOP for new road segment
               
        End If
    Loop
    
    ' No more crash data to summarize. Print the results.
    ' Print the crash count before moving to the next road segment
    Workbooks(guiwb).Sheets(sname).Cells(roadrow, itotalcrash) = totalcrash
    Workbooks(guiwb).Sheets(sname).Cells(roadrow, isevcrash) = SevCount
    ' If crash rollup data being summarized, print those results
    If factorsummary Then
        For n = 1 To ilib
            Workbooks(guiwb).Sheets(sname).Cells(roadrow, isevcrash + n) = factorarray(n)
        Next n
    End If
    
    ' Increase road row counter by 1
    roadrow = roadrow + 1
    If (roadrow - 1) Mod 100 = 0 Then
        Application.ScreenUpdating = True
        Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
        Application.Wait (Now + TimeValue("00:00:02"))
        Application.ScreenUpdating = False
    End If
Loop

'****** End crash count process for each segment **********

Workbooks(crashwb).Close False      'Close the crash workbook

' Updates the progress screen with the end time
Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
Workbooks(guiwb).Sheets("Progress").Range("B3") = 1
Workbooks(guiwb).Sheets("Progress").Range("B5") = ""
Workbooks(guiwb).Sheets("Progress").Range("B6") = Time

' Moves the UCPMinput sheet to a new workbook, to be saved as a CSV file
' Unique date and time given to file
Workbooks(guiwb).Sheets(ucpsmsheet).Move
If Workbooks(guiwb).Sheets("Inputs").Cells(2, 16) = "RSAM" Then
    sname = "UCPM-UCSMinput_Sev" & sevstring & "_" & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".csv"
ElseIf Workbooks(guiwb).Sheets("Inputs").Cells(2, 16) = "CAMS" Then
    sname = "CAMSinput_Sev" & sevstring & "_" & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".csv"
End If
sname = Replace(sname, " ", "_")
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("B10").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")

MsgBox "The combined file has been saved in the following location:" & Chr(10) & _
sname & Chr(10) & Chr(10) & _
"This file is currently open in the background; no need to browse to the filepath to view it." & Chr(10) & Chr(10) & _
"The code will now open the userform for you to run the combined file with the R code."

Workbooks(guiwb).Activate
Workbooks(guiwb).Sheets("Home").Activate
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:01"))


' Bring up the UCPM Variable Selection Screen
form_ucpsmvariable.Show

End Sub

Sub CAMSdataprep(segmentfilepath As String, crashfilepath As String, factorsummary As Boolean, sev1 As Boolean, sev2 As Boolean, sev3 As Boolean, sev4 As Boolean, sev5 As Boolean, minyr As Integer, maxyr As Integer)
' The purpose of this script is to combine the roadway and crash data for the UCPM/UCSM statistical analysis. The input file will be the same for both analyses
' segmentfilepath   = File path to the roadway segment data file
' crashfilepath     = File path to the historic crash data
' factorsummary     = Boolean (T/F) value determining whether the crash factors from the rollup data will be summarized or not
' sev1 to sev5      = Boolean (T/F) value determining whether to include the given crash severity in the input data file summary

' Define variables
Dim crashFactorList As Object
    Set crashFactorList = CreateObject("Scripting.Dictionary")
Dim sevlist As Object
    Set sevlist = CreateObject("Scripting.Dictionary")
Dim sevlist2 As Object
    Set sevlist2 = CreateObject("Scripting.Dictionary")
Dim sevstring As String
Dim sevstring2 As String

Dim i As Integer            ' Integer counter
Dim ilib As Integer         ' Counter for number of items in the library
Dim n As Integer            ' Integer counter
Dim j As Long               'integer counter
Dim sname As String         'name of sheet
Dim wname As String         'name of workbook, with roadway characteristics
Dim guiwb As String         'name of gui workbook
Dim pname As String         'name of parameters sheet
Dim dataworkbook As Workbook
Dim crashworkbook As Workbook
Dim crashwb As String
Dim CrashSheet As String
Dim datafile As String      'used for cycling through files in a directory
Dim crashheading As String
Dim ucpsmsheet As String
Dim row1 As Long

Dim totalcrash As Integer
Dim SevCount As Integer
Dim factorcount As Integer

Dim iroadroute As Integer
Dim iroadbmp As Integer
Dim iroademp As Integer

Dim begAADTcol As Integer
Dim endAADTcol As Integer

Dim itotalcrash As Integer      'New columns of data
Dim isevcrash As Integer
Dim YearCol As Integer
Dim SegIDcol As Integer
Dim isumtotal As Integer
Dim iSev5 As Integer
Dim iSev4 As Integer
Dim iSev3 As Integer
Dim iSev2 As Integer
Dim iSev1 As Integer

Dim icrashroute As Integer
Dim icrashmp As Integer
Dim icrashsev As Integer
Dim idate As Integer
Dim NumYears As Integer

Dim roadroute As String
Dim roadbmp As Double
Dim roademp As Double
Dim CrashRoute As String
Dim CrashMP As Double

Dim roadrow As Long
Dim roadcol As Integer
Dim CrashRow As Long
Dim CrashCol As Long

Dim CrashDate

' Set the name of the GUI workbook
guiwb = ActiveWorkbook.Name
'guiwb = Replace(guiwb, ".xlsm", "")

ucpsmsheet = "CAMSinput"

' Load the progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Workbooks(guiwb).Sheets("Progress")
    .Range("B2") = "Loading Roadway and Crash Files. Please wait."
    .Range("B3") = "Do not close MS Excel. Code running."
    .Range("A2") = ""
    .Range("A3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
End With
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False

' Create a library of crash factors from key sheet
ilib = 1
Do While Workbooks(guiwb).Sheets("Key").Cells(ilib + 7, 31) <> ""
    ilib = ilib + 1
Loop
ReDim crashFactorArray(1 To ilib) As String
For n = 1 To ilib
    crashFactorArray(n) = Workbooks(guiwb).Sheets("Key").Cells(n + 6, 31)
Next n
For n = 1 To ilib
    crashFactorList.Add crashFactorArray(n), n
Next n

' Create an array to count the crash factors, in order of the library list
ReDim factorarray(1 To ilib) As Integer
ReDim ifactor(1 To ilib) As Integer

' Create the list of severities to check for in the crash data
sevstring = ""
i = 1
If sev1 Then
    sevlist.Add "1", i
    i = i + 1
    sevstring = sevstring & "1"
End If
If sev2 Then
    sevlist.Add "2", i
    i = i + 1
    sevstring = sevstring & "2"
End If
If sev3 Then
    sevlist.Add "3", i
    i = i + 1
    sevstring = sevstring & "3"
End If
If sev4 Then
    sevlist.Add "4", i
    i = i + 1
    sevstring = sevstring & "4"
End If
If sev5 Then
    sevlist.Add "5", i
    i = i + 1
    sevstring = sevstring & "5"
End If

sevstring2 = "345"
sevlist2.Add "3", 1
sevlist2.Add "4", 2
sevlist2.Add "5", 3


'****** Begin Data Preparation ********

' Open and copy roadway segment data to the GUI workbook
    Set dataworkbook = Workbooks.Open(segmentfilepath) 'Opens the file in Read Only mode
    wname = ActiveWorkbook.Name
    sname = ActiveSheet.Name
'    wname = Replace(wname, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
'    wname = Replace(wname, ".xls", "")  'Removes the ".xls" so later codes will work
'    wname = Replace(wname, ".csv", "")  'Removes the ".csv" so later codes will work
        
    'Creates a new sheet, names it after the name of the workbook
    Sheets(sname).Copy After:=Workbooks(guiwb).Sheets("Key")
    ActiveSheet.Name = ucpsmsheet
    sname = ActiveSheet.Name
Workbooks(wname).Close False    'Closes the workbook, after the data has been copied

Workbooks(guiwb).Activate

Call Seg_Check_Headers(1, sname, guiwb)

' Identifies the column for Label, BegMP, EndMP in the roadway data
iroadroute = 1
Do Until Workbooks(guiwb).Sheets(sname).Cells(1, iroadroute) = Sheets("Key").Range("AC4")         '<--- Verify this matches header in dataset
    iroadroute = iroadroute + 1
Loop
iroadbmp = 1
Do Until Workbooks(guiwb).Sheets(sname).Cells(1, iroadbmp) = Sheets("Key").Range("AC5")         '<--- Verify this matches header in dataset
    iroadbmp = iroadbmp + 1
Loop
iroademp = 1
Do Until Workbooks(guiwb).Sheets(sname).Cells(1, iroademp) = Sheets("Key").Range("AC6")         '<--- Verify this matches header in dataset
    iroademp = iroademp + 1
Loop

'add year column
YearCol = iroademp + 2
Workbooks(guiwb).Sheets(sname).Activate
Workbooks(guiwb).Sheets(sname).Columns(YearCol).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Workbooks(guiwb).Sheets(sname).Cells(1, YearCol) = "YEAR"
Application.CutCopyMode = False

'delete any AADT columns for years greater than the maxyr
begAADTcol = 1
Do Until Left(Workbooks(guiwb).Sheets(sname).Cells(1, begAADTcol), 4) = "AADT"
    begAADTcol = begAADTcol + 1
Loop
i = begAADTcol
Do While Left(Workbooks(guiwb).Sheets(sname).Cells(1, i), 4) = "AADT"
    If Int(Right(Workbooks(guiwb).Sheets(sname).Cells(1, i), 4)) > maxyr Or Int(Right(Workbooks(guiwb).Sheets(sname).Cells(1, i), 4)) < minyr Then
    'if a column of AADT data is not in the year range, delete it
        Workbooks(guiwb).Sheets(sname).Columns(i).Delete
        endAADTcol = endAADTcol - 1
        i = i - 1
    End If
    i = i + 1
Loop
endAADTcol = begAADTcol
Do Until Left(Workbooks(guiwb).Sheets(sname).Cells(1, endAADTcol), 4) <> "AADT"
    endAADTcol = endAADTcol + 1
Loop
NumYears = endAADTcol - begAADTcol

'add rows for each year, fill in year-specific AADT values
row1 = 2
Do Until Cells(row1, 1) = ""
    Rows(row1 + 1 & ":" & row1 + NumYears - 1).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A" & row1 & ":BZ" & row1).AutoFill Destination:=Range("A" & row1 & ":BZ" & row1 + NumYears - 1), Type:=xlFillCopy
    
    For i = 0 To NumYears - 1
        Cells(row1 + i, begAADTcol) = Cells(row1, begAADTcol + i)
        Cells(row1 + i, begAADTcol + NumYears) = Cells(row1, begAADTcol + NumYears + i)         'Single Percent
        Cells(row1 + i, begAADTcol + NumYears * 2) = Cells(row1, begAADTcol + NumYears * 2 + i)     'Combo Percent
        Cells(row1 + i, begAADTcol + NumYears * 3) = Cells(row1, begAADTcol + NumYears * 3 + i)     'Single Count
        Cells(row1 + i, begAADTcol + NumYears * 4) = Cells(row1, begAADTcol + NumYears * 4 + i)     'Combo Count
        Cells(row1 + i, begAADTcol + NumYears * 5) = Cells(row1, begAADTcol + NumYears * 5 + i)     'Total Percent
        Cells(row1 + i, begAADTcol + NumYears * 6) = Cells(row1, begAADTcol + NumYears * 6 + i)     'Total Count
        Cells(row1 + i, YearCol) = Int(Right(Workbooks(guiwb).Sheets(sname).Cells(1, begAADTcol), 4)) - i
    Next i
    
    row1 = row1 + NumYears
Loop
Cells(1, begAADTcol) = "AADT"
Cells(1, begAADTcol + NumYears) = "Single_Percent"
Cells(1, begAADTcol + NumYears * 2) = "Combo_Percent"
Cells(1, begAADTcol + NumYears * 3) = "Single_Count"
Cells(1, begAADTcol + NumYears * 4) = "Combo_Count"
Cells(1, begAADTcol + NumYears * 5) = "Total_Percent"
Cells(1, begAADTcol + NumYears * 6) = "Total_Count"

'delete extra AADT and trucks columns
i = 1
Do
    If Left(Workbooks(guiwb).Sheets(sname).Cells(1, i), 4) = "AADT" And IsNumeric(Right(Workbooks(guiwb).Sheets(sname).Cells(1, i), 4)) = True Then
        Workbooks(guiwb).Sheets(sname).Columns(i).Delete
        i = i - 1
    End If
    i = i + 1
Loop While Workbooks(guiwb).Sheets(sname).Cells(1, i + 1) <> ""
i = 1
Do
    If IsNumeric(Left(Workbooks(guiwb).Sheets(sname).Cells(1, i), 4)) = True Then
        Workbooks(guiwb).Sheets(sname).Columns(i).Delete
        i = i - 1
    End If
    i = i + 1
Loop While Workbooks(guiwb).Sheets(sname).Cells(1, i + 1) <> ""

' Adds new columns to the end of the data
itotalcrash = i + 1
isevcrash = i + 2
isumtotal = i + 3
iSev5 = i + 4
iSev4 = i + 5
iSev3 = i + 6
iSev2 = i + 7
iSev1 = i + 8
Workbooks(guiwb).Sheets(sname).Cells(1, itotalcrash) = "Total_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, isevcrash) = "Severe_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, isumtotal) = "Sum_Total_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, iSev5) = "Sev_5_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, iSev4) = "Sev_4_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, iSev3) = "Sev_3_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, iSev2) = "Sev_2_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, iSev1) = "Sev_1_Crashes"


' Statistical interactions have already been calcuatled in the Pre-Model Data Processing Steps

' Open and copy roadway segment data to the GUI workbook
    Set crashworkbook = Workbooks.Open(crashfilepath) 'Opens the file in Read Only mode
    crashwb = ActiveWorkbook.Name
    sname = ActiveSheet.Name
'    crashwb = Replace(crashwb, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
'    crashwb = Replace(crashwb, ".xls", "")  'Removes the ".xls" so later codes will work
'    crashwb = Replace(crashwb, ".csv", "")  'Removes the ".csv" so later codes will work
        
    'Creates a new sheet, names it after the name of the workbook
    Sheets(sname).Copy After:=Workbooks(guiwb).Sheets("Key")
    ActiveSheet.Name = "CrashInput"
    CrashSheet = ActiveSheet.Name
Workbooks(crashwb).Close False    'Closes the workbook, after the data has been copied

Workbooks(guiwb).Activate

Call Seg_Check_Headers(2, CrashSheet, guiwb)

'converts the UTM X and Y coordinates in the Crash data to be Lats and Longs
'UTMtoLatLong       FLAGGED: This was messing up the data because it was already lat long.
Dim utmYCol, utmXCol As Integer
Dim datasheet As String
datasheet = "CrashInput"
utmYCol = 1
utmXCol = 1
Do Until Sheets(datasheet).Cells(1, utmYCol) = "UTM_Y"
    utmYCol = utmYCol + 1
Loop
Do Until Sheets(datasheet).Cells(1, utmXCol) = "UTM_X"
    utmXCol = utmXCol + 1
Loop
Sheets(datasheet).Cells(1, utmYCol) = "LATITUDE"
Sheets(datasheet).Cells(1, utmXCol) = "LONGITUDE"

'assign crashes to segments and years
SumCAMSbyYear   'this sub based loosely off of SumCrashDatabyYear (both are in module RGUI (this module))

'deletes the CrashInput sheet
Application.DisplayAlerts = False
If SheetExists(CrashSheet) Then
    Sheets(CrashSheet).Delete
End If
Application.DisplayAlerts = True

' Moves the UCPMinput sheet to a new workbook, to be saved as a CSV file
' Unique date and time given to file
Workbooks(guiwb).Sheets(ucpsmsheet).Move
sname = "CAMSinput_Sev" & sevstring & "_" & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".csv"
sname = Replace(sname, " ", "_")
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("M10").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")

' Moves the parameters sheet and saves is as an .xlsx file
' Unique date and time given to file
Workbooks(guiwb).Activate
Workbooks(guiwb).Sheets("Parameters").Move
pname = "CAMSParameters_Sev" & sevstring & "_" & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".xlsx"
pname = Replace(pname, " ", "_")
ActiveWorkbook.SaveAs FileName:=pname, FileFormat:=51
Workbooks(guiwb).Sheets("Inputs").Range("M33").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")

' Message to show that the process is complete
MsgBox "The combined file and the parameters file have been saved in the following location:" & Chr(10) & Chr(10) & _
"Input File: " & sname & Chr(10) & _
"Parameters File: " & pname & Chr(10) & Chr(10) & _
"These files are currently open in the background; no need to browse to the filepath to view it." & Chr(10) & Chr(10) & _
"The code will now open the userform for you to run the combined file with the R code."

Workbooks(guiwb).Activate
Workbooks(guiwb).Sheets("Home").Activate
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:01"))


' Bring up the UCPM Variable Selection Screen
form_ucpsmvariable.Show

End Sub

Sub UICPMdataprep(intfilepath As String, crashfilepath As String)
' The purpose of this script is to combine the roadway and crash data for the UICPM statistical analysis.
' intfilepath   = File path to the roadway segment data file
' crashfilepath     = File path to the historic crash data
' factorsummary     = Boolean (T/F) value determining whether the crash factors from the rollup data will be summarized or not
' sev1 to sev5      = Boolean (T/F) value determining whether to include the given crash severity in the input data file summary

' Define variables
Dim crashFactorList As Object
    Set crashFactorList = CreateObject("Scripting.Dictionary")
Dim sevlist As Object
    Set sevlist = CreateObject("Scripting.Dictionary")
Dim sevstring As String

Dim i As Integer            ' Integer counter
Dim ilib As Integer         ' Counter for number of items in the library
Dim n As Integer            ' Integer counter
Dim j As Long               'integer counter
Dim sname As String         'name of sheet
Dim wname As String         'name of workbook, with roadway characteristics
Dim guiwb As String         'name of gui workbook
Dim dataworkbook As Workbook
Dim crashworkbook As Workbook
Dim crashwb As String
Dim newcrashsheet As String
Dim oldcrashsheet As String
Dim datafile As String      'used for cycling through files in a directory
Dim crashheading As String
Dim uicpmsheet As String

Dim totalcrash As Integer
Dim SevCount As Integer
Dim factorcount As Integer

Dim colTotalCrash, colSevereCrash As Integer     'New columns of data
Dim ColSumTotCrash As Integer
Dim ColSev5, ColSev4, ColSev3, ColSev2, ColSev1 As Integer

Dim icrashroute As Integer
Dim icrashmp As Integer
Dim icrashsev As Integer

Dim roadroute As String
Dim roadbmp As Double
Dim roademp As Double
Dim CrashRoute As String
Dim CrashMP As Double

Dim roadrow As Long
Dim roadcol As Integer
Dim CrashRow As Long
Dim CrashCol As Long

Dim crashheadcol As Long
Dim FAkeycol As Long

Dim FAmethod As String
Dim UICPMcol As Integer
Dim FArow As Integer

Dim row1 As Long
Dim col1 As Long
Dim MinCYear, MaxCYear, MinIYear, MaxIYear As Integer
Dim NumYears As Integer

Application.ScreenUpdating = False

' Set the name of the GUI workbook
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

uicpmsheet = "UICPMinput"
newcrashsheet = "CrashInput"

' Load the progress screen
Application.ScreenUpdating = True
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Workbooks(guiwb).Sheets("Progress")
    .Range("A2") = "Loading Roadway and Crash Files. Please wait."
    .Range("A3") = "Do not close MS Excel. Code running."
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B2") = ""
    .Range("B3") = ""
    .Range("B4") = Time
    .Range("B5") = ""
    .Range("B6") = ""
    .Range("C2") = ""
    .Range("C3") = ""
    .Range("C4") = ""
    .Range("C5") = ""
    .Range("C6") = ""
End With
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False


'****** Begin Data Preparation ********

' Open and copy intersection data to the GUI workbook
    Set dataworkbook = Workbooks.Open(intfilepath) 'Opens the file in Read Only mode
    wname = ActiveWorkbook.Name
    sname = ActiveSheet.Name
    
    wname = Replace(wname, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
    wname = Replace(wname, ".xls", "")  'Removes the ".xls" so later codes will work
    wname = Replace(wname, ".csv", "")  'Removes the ".csv" so later codes will work
        
    'Creates a new sheet, names it after the name of the workbook
    Sheets(sname).Copy After:=Workbooks(guiwb).Sheets("Key")
    ActiveSheet.Name = uicpmsheet
    sname = ActiveSheet.Name
Workbooks(wname).Close False    'Closes the workbook, after the data has been copied

Workbooks(guiwb).Activate

Call Int_Check_Headers(1, sname, guiwb)

' Adds new columns to the end of the data
i = 1
Do
    i = i + 1
Loop While Workbooks(guiwb).Sheets(sname).Cells(1, i + 1) <> ""
colTotalCrash = i + 1
colSevereCrash = i + 2
ColSumTotCrash = i + 3
ColSev5 = i + 4
ColSev4 = i + 5
ColSev3 = i + 6
ColSev2 = i + 7
ColSev1 = i + 8
Workbooks(guiwb).Sheets(sname).Cells(1, colTotalCrash) = "Total_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, colSevereCrash) = "Severe_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, ColSumTotCrash) = "Sum_Total_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, ColSev5) = "Sev_5_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, ColSev4) = "Sev_4_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, ColSev3) = "Sev_3_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, ColSev2) = "Sev_2_Crashes"
Workbooks(guiwb).Sheets(sname).Cells(1, ColSev1) = "Sev_1_Crashes"

' Statistical interactions have already been calcuatled in the Pre-Model Data Processing Steps

' Open the crash data
    Set crashworkbook = Workbooks.Open(crashfilepath) 'Opens the file in Read Only mode
    crashwb = ActiveWorkbook.Name
    oldcrashsheet = ActiveSheet.Name
    crashwb = Replace(crashwb, ".xlsx", "")  'Removes the ".xlsx" so later codes will work
    crashwb = Replace(crashwb, ".xls", "")  'Removes the ".xls" so later codes will work
    crashwb = Replace(crashwb, ".csv", "")  'Removes the ".csv" so later codes will work
    
    'Creates a new sheet, names it after the name of the workbook
    Sheets(oldcrashsheet).Copy After:=Workbooks(guiwb).Sheets(uicpmsheet)
    ActiveSheet.Name = newcrashsheet
' Close crash workbook
Workbooks(crashwb).Close False

Workbooks(guiwb).Activate

Call Int_Check_Headers(2, newcrashsheet, guiwb)

'****** End of Data Preparation ********

'****** Identify years of data available ************

'Determine available years for both crash and entering vehicle data
'Crash Years
row1 = 2
col1 = 1
Do Until Sheets(newcrashsheet).Cells(1, col1) = "CRASH_DATETIME"
    col1 = col1 + 1
Loop

MinCYear = year(Sheets(newcrashsheet).Cells(row1, col1))
MaxCYear = year(Sheets(newcrashsheet).Cells(row1, col1))
row1 = row1 + 1
Do Until Sheets(newcrashsheet).Cells(row1, col1) = ""
    If year(Sheets(newcrashsheet).Cells(row1, col1)) < MinCYear Then
        MinCYear = year(Sheets(newcrashsheet).Cells(row1, col1))
    End If
    
    If year(Sheets(newcrashsheet).Cells(row1, col1)) > MaxCYear Then
        MaxCYear = year(Sheets(newcrashsheet).Cells(row1, col1))
    End If
    row1 = row1 + 1
Loop

'Ent Veh Years
col1 = 1
Do Until Left(Sheets(sname).Cells(1, col1), 7) = "ENT_VEH"
    col1 = col1 + 1
Loop

MinIYear = Int(Right(Sheets(sname).Cells(1, col1), 4))
MaxIYear = Int(Right(Sheets(sname).Cells(1, col1), 4))

Do Until Left(Sheets(sname).Cells(1, col1), 7) <> "ENT_VEH"
    If Int(Right(Sheets(sname).Cells(1, col1), 4)) < MinIYear Then
        MinIYear = Int(Right(Sheets(sname).Cells(1, col1), 4))
    End If
    
    If Int(Right(Sheets(sname).Cells(1, col1), 4)) > MaxIYear Then
        MaxIYear = Int(Right(Sheets(sname).Cells(1, col1), 4))
    End If
    col1 = col1 + 1
Loop

'Open data years form for user to determine
If MinCYear < MinIYear Then
    MinCYear = MinIYear
End If

If MaxCYear > MaxIYear Then
    MaxCYear = MaxIYear
End If

NumYears = MaxCYear - MinCYear

row1 = 1
col1 = 2
Sheets("Key").Columns(col1).ClearContents
For i = 0 To NumYears
    Sheets("Key").Cells(row1 + i, col1) = MinCYear + i
Next i

Application.ScreenUpdating = True

form_UICPMyears.cboMinYear.RowSource = "'Key'!B1:B" & NumYears + 1
form_UICPMyears.cboMinYear.ListIndex = 0
form_UICPMyears.cboMaxYear.RowSource = "'Key'!B1:B" & NumYears + 1
form_UICPMyears.cboMaxYear.ListIndex = NumYears
form_UICPMyears.Show

End Sub

Sub UICPMdataprep2()

Dim YearCol As Long
Dim NumYears, MaxEntVehCol As Integer

Dim i As Integer            ' Integer counter
Dim ilib As Integer         ' Counter for number of items in the library
Dim n As Integer            ' Integer counter
Dim j As Long               'integer counter
Dim sname As String         'name of sheet
Dim wname As String         'name of workbook, with roadway characteristics
Dim guiwb As String         'name of gui workbook
Dim dataworkbook As Workbook
Dim crashworkbook As Workbook
Dim crashwb As String
Dim newcrashsheet As String
Dim datafile As String      'used for cycling through files in a directory
Dim crashheading As String
Dim uicpmsheet As String
Dim sevstring As String
Dim PSheet As String

Dim totalcrash As Integer
Dim SevCount As Integer
Dim factorcount As Integer

Dim colTotalCrash, colSevereCrash As Integer     'New columns of data

Dim icrashroute As Integer
Dim icrashmp As Integer
Dim icrashsev As Integer

Dim roadroute As String
Dim roadbmp As Double
Dim roademp As Double
Dim CrashRoute As String
Dim CrashMP As Double

Dim roadrow As Long
Dim roadcol As Integer
Dim CrashRow As Long
Dim CrashCol As Long

Dim crashheadcol As Long
Dim FAkeycol As Long

Dim FAmethod As String
Dim UICPMcol As Integer
Dim FArow As Integer

Dim row1 As Long
Dim col1 As Long
Dim MinCYear, MaxCYear, MinIYear, MaxIYear As Integer
Dim minyear, maxyear As Integer
Dim StartTime, EndTime As String
Dim PercTrkCol As Integer

Dim wd As String


guiwb = ActiveWorkbook.Name
PSheet = "Parameters"
uicpmsheet = "UICPMinput"
newcrashsheet = "CrashInput"
minyear = Sheets("Key").Cells(1, 3).Value
maxyear = Sheets("Key").Cells(2, 3).Value
wd = Sheets("Inputs").Range("I2").Value

'****** Separate data by year *******
NumYears = maxyear - minyear + 1

Sheets(uicpmsheet).Activate

'Add year column
YearCol = 1
Do Until Cells(1, YearCol) = "TRAFFIC_CO"
    YearCol = YearCol + 1
Loop

Columns(YearCol).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, YearCol) = "YEAR"
Application.CutCopyMode = False

'Delete ent veh columns not in year range and find entering vehicle column with max year
MaxEntVehCol = 1
Do Until Cells(1, MaxEntVehCol) = ""
    If Left(Cells(1, MaxEntVehCol), 7) = "ENT_VEH" And IsNumeric(Right(Cells(1, MaxEntVehCol), 4)) = True Then
        If CInt(Right(Cells(1, MaxEntVehCol), 4)) <= maxyear And CInt(Right(Cells(1, MaxEntVehCol), 4)) >= minyear Then
            MaxEntVehCol = MaxEntVehCol + 1
        Else
            Columns(MaxEntVehCol).EntireColumn.Delete
        End If
    Else
       MaxEntVehCol = MaxEntVehCol + 1
    End If
Loop

MaxEntVehCol = 1
Do Until Cells(1, MaxEntVehCol) = ""
    If Left(Cells(1, MaxEntVehCol), 7) = "ENT_VEH" And IsNumeric(Right(Cells(1, MaxEntVehCol), 4)) = True Then
        If CInt(Right(Cells(1, MaxEntVehCol), 4)) = maxyear Then
            Exit Do
        Else
            MaxEntVehCol = MaxEntVehCol + 1
        End If
    Else
        MaxEntVehCol = MaxEntVehCol + 1
    End If
Loop

'Find first percent trucks column
PercTrkCol = 1
Do Until Right(Cells(1, PercTrkCol), 14) = "Percent_Trucks"
    PercTrkCol = PercTrkCol + 1
Loop

Sheets(uicpmsheet).UsedRange

'Add entering vehicle data and truck data for each year
row1 = 2
Do Until Cells(row1, 1) = ""
    Rows(row1 + 1 & ":" & row1 + NumYears - 1).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A" & row1 & ":BZ" & row1).AutoFill Destination:=Range("A" & row1 & ":BZ" & row1 + NumYears - 1), Type:=xlFillCopy
    
    For i = 0 To NumYears - 1
        Cells(row1 + i, MaxEntVehCol) = Cells(row1, MaxEntVehCol + i)
        Cells(row1 + i, PercTrkCol) = Cells(row1, PercTrkCol + i)
    Next i
    
    row1 = row1 + NumYears
Loop
Cells(1, MaxEntVehCol) = "ENT_VEH"
Cells(1, PercTrkCol) = "Percent_Trucks"

'Enter data years
row1 = 2
Do Until Cells(row1, 1) = ""
    For i = 0 To NumYears - 1
        Cells(row1 + i, YearCol) = maxyear - i
    Next i
    row1 = row1 + NumYears
Loop

'Delete uneeded ent veh and Percent_Trucks columns
col1 = 1
Do Until Cells(1, col1) = ""
    If Left(Cells(1, col1), 7) = "ENT_VEH" And Cells(1, col1) <> "ENT_VEH" Or Right(Cells(1, col1), 14) = "Percent_Trucks" And Cells(1, col1) <> "Percent_Trucks" Then
        Columns(col1).EntireColumn.Delete
    Else
        col1 = col1 + 1
    End If
Loop

'Change UTM coordinates to latitude and longitude
'UTMtoLatLong     FLAGGED: We should just do this in R if we need it. This was messing up the data because it was already lat long.
Dim utmYCol, utmXCol As Integer
Dim datasheet As String
datasheet = "CrashInput"
utmYCol = 1
utmXCol = 1
Do Until Sheets(datasheet).Cells(1, utmYCol) = "UTM_Y"
    utmYCol = utmYCol + 1
Loop
Do Until Sheets(datasheet).Cells(1, utmXCol) = "UTM_X"
    utmXCol = utmXCol + 1
Loop
Sheets(datasheet).Cells(1, utmYCol) = "LATITUDE"
Sheets(datasheet).Cells(1, utmXCol) = "LONGITUDE"

'Sort by latitude and longitude
LatLongElevSortCrashUICPM

' Updates the progress sheet
Workbooks(guiwb).Sheets("Progress").Activate
With Workbooks(guiwb).Sheets("Progress")
    .Range("A2") = ""
    .Range("A3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = "Update Time"
    .Range("A6") = "End Time"
    .Range("A2") = "Assigning Crashes"
    .Range("B3") = ""
    .Range("C3") = ""
    .Range("B5") = ""
    .Range("B6") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:03"))
Application.ScreenUpdating = False

'Assign crash data to intersections by year
'Assigns crashes to intersections based on MP
'tallies crash information by year for each intersection
SumCrashDatabyYear

'Change speed limit and functional class values to max and min values
MaxMinCols

'Workbooks(crashwb).Close False      'Close the crash workbook

' Updates the progress screen with the end time
'Workbooks(guiwb).Sheets("Progress").Range("B2") = roadrow - 1
Workbooks(guiwb).Sheets("Progress").Range("B3") = 1
Workbooks(guiwb).Sheets("Progress").Range("B6") = Time

' Find the functional area method that was done
UICPMcol = 1
FArow = 1
Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "UICPM"
    UICPMcol = UICPMcol + 1
Loop

Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "Selected FA Parameter"
    FArow = FArow + 1
Loop

If Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Speed Limit" Then
    FAmethod = "SL"
ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Functional Class" Then
    FAmethod = "FC"
ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Urban Code" Then
    FAmethod = "UC"
End If

Sheets(PSheet).Cells(2, 2) = Sheets("Inputs").Cells(FArow + 1, UICPMcol)

' Find the severities
sevstring = Sheets("Inputs").Cells(12, UICPMcol + 1)

'Save UICPM dataset as separate workbook
Sheets(uicpmsheet).Move
sname = wd & "/UICPMinput_Sev" & sevstring & "_" & FAmethod & "_" & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".csv"
sname = Replace(sname, " ", "_")
sname = Replace(sname, "/", "\")
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("I10").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")
ActiveWorkbook.Close

'Save parameters dataset as separate workbook
Sheets(PSheet).Move
sname = wd & "/UICPM_Parameters_" & sevstring & "_" & FAmethod & "_" & Replace(Date, "/", "-") & " " & Replace(Time, ":", "-") & ".csv"
sname = Replace(sname, " ", "_")
sname = Replace(sname, "/", "\")
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
ActiveWorkbook.Close

Workbooks(guiwb).Activate

Workbooks(guiwb).Sheets(newcrashsheet).Delete
Application.DisplayAlerts = True

Workbooks(guiwb).Sheets("Home").Activate
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:01"))

StartTime = Workbooks(guiwb).Sheets("Progress").Range("B4").Text
EndTime = Workbooks(guiwb).Sheets("Progress").Range("B6").Text

'Message box with start and end times
MsgBox "The UICPM input file has been successfully created and saved in the working directory folder. The following is a summary of the process:" & Chr(10) & _
Chr(10) & _
"Process: UICPM Input (Assigning Crashes)" & Chr(10) & _
"Start Time: " & StartTime & Chr(10) & _
"End Time: " & EndTime & Chr(10), vbOKOnly, "Process Complete"

' Bring up the UICPM Variable Selection Screen
form_uicpmvariable.Show

End Sub

Sub SumCrashDatabyYear()

'Declare variables
Dim IntRow, CrashRow, CrashRowPrev As Long
Dim IntLatCol As Integer, IntLongCol As Integer, CrashLatCol As Integer, CrashLongCol As Integer, UrbanCodeCol, UrbanCode As Long
Dim IntLat As Double, IntLong As Double, LatFtEquiv, LongFtEquiv, MinLatStart, MaxLatEnd As Double
Dim CrashLat As Double, CrashLong As Double, LatDiff, LongDiff, DiffRadius As Double
Dim SpeedLimit, NumYears As Integer
Dim datasheet As String
Dim CrashSheet As String
Dim UCDesc, FCDesc As String
Dim row1, col1, CHeadCol As Long

Dim MaxSLCol, MinSLCol, MaxFCCol, UCcol As Long
Dim IntRoute0Col, IntRoute1Col, IntRoute2Col, IntRoute3Col, IntRoute4Col As Long
Dim IntMP0Col, IntMP1Col, IntMP2Col, IntMP3Col, IntMP4Col As Long
Dim IntNumCol, YearCol, TotCrashCol, SevCrashCol, PedInvCol1, BikeInvCol1, MotInvCol1, ImpResCol1, UnrestCol1, DUICol1 As Long
Dim AggressiveCol1, DistractedCol1, DrowsyCol1, SpeedCol1, IntRelCol1, WeatherCol1, AdvRoadCol1, RoadGeomCol1 As Long
Dim WAnimalCol1, DAnimalCol1, RoadDepartCol1, OverturnCol1, CommVehCol1, InterstateCol1, TeenCol1, OldCol1 As Long
Dim UrbanCol1, NightCol1, SingleVehCol1, TrainCol1, RailCol1, TransitCol1, FixedObCol1, WorkZoneCol1, CollMannerCol1 As Long
Dim SLCol0, SLCol1, SLCol2, SLCol3, SLCol4, FCCodeCol0, FCCodeCol1, FCCodeCol2, FCCodeCol3, FCCodeCol4, FCTypeCol0, FCTypeCol1, FCTypeCol2, FCTypeCol3, FCTypeCol4 As Long
Dim IntDistCol0, IntDistCol1, IntDistCol2, IntDistCol3, IntDistCol4, IntDist0, IntDist1, IntDist2, IntDist3, IntDist4, MaxDist, MinDist As Long
Dim MaxRoadWCol, MinRoadWcol, MaxRoadW, MinRoadW As Long
Dim SL0, SL1, SL2, SL3, SL4, FC0, FC1, FC2, FC3, FC4 As Integer
Dim IntYear As String
Dim FADistLat As Double, FADistLong As Double

Dim crashroutecol, CrashIDCol, crashMPcol As Long
Dim CrashDateCol, CrashSeverityCol, PedInvCol2, BikeInvCol2, MotInvCol2, ImpResCol2, UnrestCol2, DUICol2 As Long
Dim AggressiveCol2, DistractedCol2, DrowsyCol2, SpeedCol2, IntRelCol2, WeatherCol2, AdvRoadCol2, RoadGeomCol2 As Long
Dim WAnimalCol2, DAnimalCol2, RoadDepartCol2, OverturnCol2, CommVehCol2, InterstateCol2, TeenCol2, OldCol2 As Long
Dim UrbanCol2, NightCol2, SingleVehCol2, TrainCol2, RailCol2, TransitCol2, FixedObCol2, WorkZoneCol2, CollMannerCol2 As Long
Dim CrashYear As String

Dim UICPMcol As Long
Dim FArow As Integer
Dim i, j As Long
Dim sevstring As String
ReDim Sevs(1 To 5) As Integer
Dim Route0 As Variant
Dim Route1 As Variant
Dim Route2 As Variant
Dim Route3 As Variant
Dim Route4 As Variant
Dim IntMP0 As Double
Dim IntMP1 As Double
Dim IntMP2 As Double
Dim IntMP3 As Double
Dim IntMP4 As Double

Dim FADist0, FADist1, FADist2, FADist3, FADist4 As Integer

Dim ColSumTotCrash As Integer
Dim ColSev5, ColSev4, ColSev3, ColSev2, ColSev1 As Integer
Dim CrashSev As Integer
Dim CrashRoute As Integer
Dim CrashMP As Double
ReDim SevCount(1 To 5) As Integer
Dim IntNum, CrashID As String
Dim prow As Long
Dim PSheet As String
Dim CrashCol As Long

'Assign initial values
IntRow = 2
CrashRow = 2
prow = 5

datasheet = "UICPMinput"
CrashSheet = "CrashInput"
PSheet = "Parameters"

'Find first blank column in roadway intersection data and add in column headings
row1 = 9
col1 = 1
CHeadCol = 1
Do Until Sheets(datasheet).Cells(1, col1) = ""
    col1 = col1 + 1
Loop

Do Until Sheets("Key").Cells(1, CHeadCol) = "Intersection Check Headers"
    CHeadCol = CHeadCol + 1
Loop

Do Until Sheets("Key").Cells(2, CHeadCol) = 2
    CHeadCol = CHeadCol + 1
Loop

Do Until Sheets("Key").Cells(row1, CHeadCol + 2) = ""
    Sheets(datasheet).Cells(1, col1) = Sheets("Key").Cells(row1, CHeadCol + 2)
    col1 = col1 + 1
    row1 = row1 + 1
Loop

'Find columns                                                              'FLAGGED: This hard codes the condition for if a column is necessary
IntRoute0Col = FindColumn("ROUTE", datasheet, 1, True)
IntRoute1Col = FindColumn("INT_RT_1", datasheet, 1, True)
IntRoute2Col = FindColumn("INT_RT_2", datasheet, 1, True)
IntRoute3Col = FindColumn("INT_RT_3", datasheet, 1, False)
IntRoute4Col = FindColumn("INT_RT_4", datasheet, 1, False)
IntMP0Col = FindColumn("UDOT_BMP", datasheet, 1, True)
IntMP1Col = FindColumn("INT_RT_1_M", datasheet, 1, True)
IntMP2Col = FindColumn("INT_RT_2_M", datasheet, 1, True)
IntMP3Col = FindColumn("INT_RT_3_M", datasheet, 1, True)
IntMP4Col = FindColumn("INT_RT_4_M", datasheet, 1, True)
IntNumCol = FindColumn("INT_ID", datasheet, 1, True)
YearCol = FindColumn("YEAR", datasheet, 1, True)
UCcol = FindColumn("URBAN_DESC", datasheet, 1, True)
UrbanCodeCol = FindColumn("URBAN_CODE", datasheet, 1, True)
SLCol0 = FindColumn("SPEED_LIMIT_0", datasheet, 1, True)
SLCol1 = FindColumn("SPEED_LIMIT_1", datasheet, 1, True)
SLCol2 = FindColumn("SPEED_LIMIT_2", datasheet, 1, True)
SLCol3 = FindColumn("SPEED_LIMIT_3", datasheet, 1, True)
SLCol4 = FindColumn("SPEED_LIMIT_4", datasheet, 1, True)
FCCodeCol0 = FindColumn("FCCode_0", datasheet, 1, True)
FCCodeCol1 = FindColumn("FCCode_1", datasheet, 1, True)
FCCodeCol2 = FindColumn("FCCode_2", datasheet, 1, True)
FCCodeCol3 = FindColumn("FCCode_3", datasheet, 1, True)
FCCodeCol4 = FindColumn("FCCode_4", datasheet, 1, True)
FCTypeCol0 = FindColumn("FCType_0", datasheet, 1, True)
FCTypeCol1 = FindColumn("FCType_1", datasheet, 1, True)
FCTypeCol2 = FindColumn("FCType_2", datasheet, 1, True)
FCTypeCol3 = FindColumn("FCType_3", datasheet, 1, True)
FCTypeCol4 = FindColumn("FCType_4", datasheet, 1, True)
IntDistCol0 = FindColumn("Int_Dist_0", datasheet, 1, True)
IntDistCol1 = FindColumn("INT_DIST_1", datasheet, 1, True)
IntDistCol2 = FindColumn("INT_DIST_2", datasheet, 1, True)
IntDistCol3 = FindColumn("INT_DIST_3", datasheet, 1, True)
IntDistCol4 = FindColumn("Int_Dist_4", datasheet, 1, True)
MaxRoadWCol = FindColumn("MAX_ROAD_WIDTH", datasheet, 1, True)
MinRoadWcol = FindColumn("MIN_ROAD_WIDTH", datasheet, 1, True)
TotCrashCol = FindColumn("Total_Crashes", datasheet, 1, True)
SevCrashCol = FindColumn("Severe_Crashes", datasheet, 1, True)
ColSumTotCrash = FindColumn("Sum_Total_Crashes", datasheet, 1, True)
ColSev5 = FindColumn("Sev_5_Crashes", datasheet, 1, True)
ColSev4 = FindColumn("Sev_4_Crashes", datasheet, 1, True)
ColSev3 = FindColumn("Sev_3_Crashes", datasheet, 1, True)
ColSev2 = FindColumn("Sev_2_Crashes", datasheet, 1, True)
ColSev1 = FindColumn("Sev_1_Crashes", datasheet, 1, True)
WorkZoneCol1 = FindColumn("WORKZONE_RELATED", datasheet, 1, True)
CollMannerCol1 = FindColumn("HEADON_COLLISION", datasheet, 1, True)
PedInvCol1 = FindColumn("PEDESTRIAN_INVOLVED", datasheet, 1, True)
BikeInvCol1 = FindColumn("BICYCLIST_INVOLVED", datasheet, 1, True)
MotInvCol1 = FindColumn("MOTORCYCLE_INVOLVED", datasheet, 1, True)
ImpResCol1 = FindColumn("IMPROPER_RESTRAINT", datasheet, 1, True)
UnrestCol1 = FindColumn("UNRESTRAINED", datasheet, 1, True)
DUICol1 = FindColumn("DUI", datasheet, 1, True)
AggressiveCol1 = FindColumn("AGGRESSIVE_DRIVING", datasheet, 1, True)
DistractedCol1 = FindColumn("DISTRACTED_DRIVING", datasheet, 1, True)
DrowsyCol1 = FindColumn("DROWSY_DRIVING", datasheet, 1, True)
SpeedCol1 = FindColumn("SPEED_RELATED", datasheet, 1, True)
IntRelCol1 = FindColumn("INTERSECTION_RELATED", datasheet, 1, True)
WeatherCol1 = FindColumn("ADVERSE_WEATHER", datasheet, 1, True)
AdvRoadCol1 = FindColumn("ADVERSE_ROADWAY_SURF_CONDITION", datasheet, 1, True)
RoadGeomCol1 = FindColumn("ROADWAY_GEOMETRY_RELATED", datasheet, 1, True)
WAnimalCol1 = FindColumn("WILD_ANIMAL_RELATED", datasheet, 1, True)
DAnimalCol1 = FindColumn("DOMESTIC_ANIMAL_RELATED", datasheet, 1, True)
RoadDepartCol1 = FindColumn("ROADWAY_DEPARTURE", datasheet, 1, True)
OverturnCol1 = FindColumn("OVERTURN_ROLLOVER", datasheet, 1, True)
CommVehCol1 = FindColumn("COMMERCIAL_MOTOR_VEH_INVOLVED", datasheet, 1, True)
InterstateCol1 = FindColumn("INTERSTATE_HIGHWAY", datasheet, 1, True)
TeenCol1 = FindColumn("TEENAGE_DRIVER_INVOLVED", datasheet, 1, True)
OldCol1 = FindColumn("OLDER_DRIVER_INVOLVED", datasheet, 1, True)
UrbanCol1 = FindColumn("URBAN_COUNTY", datasheet, 1, True)
NightCol1 = FindColumn("NIGHT_DARK_CONDITION", datasheet, 1, True)
SingleVehCol1 = FindColumn("SINGLE_VEHICLE", datasheet, 1, True)
TrainCol1 = FindColumn("TRAIN_INVOLVED", datasheet, 1, True)
RailCol1 = FindColumn("RAILROAD_CROSSING", datasheet, 1, True)
TransitCol1 = FindColumn("TRANSIT_VEHICLE_INVOLVED", datasheet, 1, True)
FixedObCol1 = FindColumn("COLLISION_WITH_FIXED_OBJECT", datasheet, 1, True)
IntLatCol = FindColumn("LATITUDE", datasheet, 1, True)
IntLongCol = FindColumn("LONGITUDE", datasheet, 1, True)

crashroutecol = FindColumn("LABEL", CrashSheet, 1, True)
crashMPcol = FindColumn("MILEPOINT", CrashSheet, 1, True)
CrashDateCol = FindColumn("CRASH_DATETIME", CrashSheet, 1, True)
CrashSeverityCol = FindColumn("CRASH_SEVERITY_ID", CrashSheet, 1, True)
CrashCol = FindColumn("", CrashSheet, 1, True)
CrashSeverityCol = FindColumn("CRASH_SEVERITY_ID", CrashSheet, 1, True)
WorkZoneCol2 = FindColumn("WORKZONE_RELATED", CrashSheet, 1, True)
CollMannerCol2 = FindColumn("HEADON_COLLISION", CrashSheet, 1, True)
PedInvCol2 = FindColumn("PEDESTRIAN_INVOLVED", CrashSheet, 1, True)
BikeInvCol2 = FindColumn("BICYCLIST_INVOLVED", CrashSheet, 1, True)
MotInvCol2 = FindColumn("MOTORCYCLE_INVOLVED", CrashSheet, 1, True)
ImpResCol2 = FindColumn("IMPROPER_RESTRAINT", CrashSheet, 1, True)
UnrestCol2 = FindColumn("UNRESTRAINED", CrashSheet, 1, True)
DUICol2 = FindColumn("DUI", CrashSheet, 1, True)
AggressiveCol2 = FindColumn("AGGRESSIVE_DRIVING", CrashSheet, 1, True)
DistractedCol2 = FindColumn("DISTRACTED_DRIVING", CrashSheet, 1, True)
DrowsyCol2 = FindColumn("DROWSY_DRIVING", CrashSheet, 1, True)
SpeedCol2 = FindColumn("SPEED_RELATED", CrashSheet, 1, True)
IntRelCol2 = FindColumn("INTERSECTION_RELATED", CrashSheet, 1, True)
WeatherCol2 = FindColumn("ADVERSE_WEATHER", CrashSheet, 1, True)
AdvRoadCol2 = FindColumn("ADVERSE_ROADWAY_SURF_CONDITION", CrashSheet, 1, True)
RoadGeomCol2 = FindColumn("ROADWAY_GEOMETRY_RELATED", CrashSheet, 1, True)
WAnimalCol2 = FindColumn("WILD_ANIMAL_RELATED", CrashSheet, 1, True)
DAnimalCol2 = FindColumn("DOMESTIC_ANIMAL_RELATED", CrashSheet, 1, True)
RoadDepartCol2 = FindColumn("ROADWAY_DEPARTURE", CrashSheet, 1, True)
OverturnCol2 = FindColumn("OVERTURN_ROLLOVER", CrashSheet, 1, True)
CommVehCol2 = FindColumn("COMMERCIAL_MOTOR_VEH_INVOLVED", CrashSheet, 1, True)
InterstateCol2 = FindColumn("INTERSTATE_HIGHWAY", CrashSheet, 1, True)
TeenCol2 = FindColumn("TEENAGE_DRIVER_INVOLVED", CrashSheet, 1, True)
OldCol2 = FindColumn("OLDER_DRIVER_INVOLVED", CrashSheet, 1, True)
UrbanCol2 = FindColumn("URBAN_COUNTY", CrashSheet, 1, True)
NightCol2 = FindColumn("NIGHT_DARK_CONDITION", CrashSheet, 1, True)
SingleVehCol2 = FindColumn("SINGLE_VEHICLE", CrashSheet, 1, True)
TrainCol2 = FindColumn("TRAIN_INVOLVED", CrashSheet, 1, True)
RailCol2 = FindColumn("RAILROAD_CROSSING", CrashSheet, 1, True)
TransitCol2 = FindColumn("TRANSIT_VEHICLE_INVOLVED", CrashSheet, 1, True)
FixedObCol2 = FindColumn("COLLISION_WITH_FIXED_OBJECT", CrashSheet, 1, True)
CrashLatCol = FindColumn("LATITUDE", CrashSheet, 1, True)
CrashLongCol = FindColumn("LONGITUDE", CrashSheet, 1, True)


FArow = 1
UICPMcol = 1

Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "UICPM"
    UICPMcol = UICPMcol + 1
Loop

Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "Selected FA Parameter"
    FArow = FArow + 1
Loop

Sheets(CrashSheet).Activate
RouteIDSort   'sorts CrashInput sheet
Sheets(datasheet).Activate

'Find number of years of data
NumYears = 1
Do Until Sheets(datasheet).Cells(IntRow + NumYears, YearCol) = Sheets(datasheet).Cells(IntRow, YearCol)
    NumYears = NumYears + 1
Loop

'Insert zeros for all values
IntRow = 2
Do Until Sheets(datasheet).Cells(IntRow, 1) = ""
    Sheets(datasheet).Cells(IntRow, TotCrashCol) = 0
    Sheets(datasheet).Cells(IntRow, SevCrashCol) = 0
    Sheets(datasheet).Cells(IntRow, ColSumTotCrash) = 0
    Sheets(datasheet).Cells(IntRow, ColSev5) = 0
    Sheets(datasheet).Cells(IntRow, ColSev4) = 0
    Sheets(datasheet).Cells(IntRow, ColSev3) = 0
    Sheets(datasheet).Cells(IntRow, ColSev2) = 0
    Sheets(datasheet).Cells(IntRow, ColSev1) = 0
    Sheets(datasheet).Cells(IntRow, PedInvCol1) = 0
    Sheets(datasheet).Cells(IntRow, BikeInvCol1) = 0
    Sheets(datasheet).Cells(IntRow, MotInvCol1) = 0
    Sheets(datasheet).Cells(IntRow, ImpResCol1) = 0
    Sheets(datasheet).Cells(IntRow, UnrestCol1) = 0
    Sheets(datasheet).Cells(IntRow, DUICol1) = 0
    Sheets(datasheet).Cells(IntRow, AggressiveCol1) = 0
    Sheets(datasheet).Cells(IntRow, DistractedCol1) = 0
    Sheets(datasheet).Cells(IntRow, DrowsyCol1) = 0
    Sheets(datasheet).Cells(IntRow, SpeedCol1) = 0
    Sheets(datasheet).Cells(IntRow, IntRelCol1) = 0
    Sheets(datasheet).Cells(IntRow, WeatherCol1) = 0
    Sheets(datasheet).Cells(IntRow, AdvRoadCol1) = 0
    Sheets(datasheet).Cells(IntRow, RoadGeomCol1) = 0
    Sheets(datasheet).Cells(IntRow, WAnimalCol1) = 0
    Sheets(datasheet).Cells(IntRow, DAnimalCol1) = 0
    Sheets(datasheet).Cells(IntRow, RoadDepartCol1) = 0
    Sheets(datasheet).Cells(IntRow, OverturnCol1) = 0
    Sheets(datasheet).Cells(IntRow, CommVehCol1) = 0
    Sheets(datasheet).Cells(IntRow, InterstateCol1) = 0
    Sheets(datasheet).Cells(IntRow, TeenCol1) = 0
    Sheets(datasheet).Cells(IntRow, OldCol1) = 0
    Sheets(datasheet).Cells(IntRow, UrbanCol1) = 0
    Sheets(datasheet).Cells(IntRow, NightCol1) = 0
    Sheets(datasheet).Cells(IntRow, SingleVehCol1) = 0
    Sheets(datasheet).Cells(IntRow, TrainCol1) = 0
    Sheets(datasheet).Cells(IntRow, RailCol1) = 0
    Sheets(datasheet).Cells(IntRow, TransitCol1) = 0
    Sheets(datasheet).Cells(IntRow, FixedObCol1) = 0
    Sheets(datasheet).Cells(IntRow, WorkZoneCol1) = 0
    Sheets(datasheet).Cells(IntRow, CollMannerCol1) = 0
    IntRow = IntRow + 1
Loop

'Enter crash headers into parameters sheet
Sheets(PSheet).Cells(4, 1) = "INT_ID"
For i = 1 To CrashCol + 1
    Sheets(PSheet).Cells(4, i + 1) = Sheets(CrashSheet).Cells(1, i)
Next i

'Loop through intersections and assign crashes that are within the functional area
IntRow = 2
Do While Sheets(datasheet).Cells(IntRow, 1) <> ""
    IntNum = Sheets(datasheet).Cells(IntRow, IntNumCol)
    MaxRoadW = Sheets(datasheet).Cells(IntRow, MaxRoadWCol)
    MinRoadW = Sheets(datasheet).Cells(IntRow, MinRoadWcol)
    IntDist0 = Sheets(datasheet).Cells(IntRow, IntDistCol0)
    IntDist1 = Sheets(datasheet).Cells(IntRow, IntDistCol1)
    IntDist2 = Sheets(datasheet).Cells(IntRow, IntDistCol2)
    IntDist3 = Sheets(datasheet).Cells(IntRow, IntDistCol3)
    IntDist4 = Sheets(datasheet).Cells(IntRow, IntDistCol4)
    Route0 = Sheets(datasheet).Cells(IntRow, IntRoute0Col)
    Route1 = Sheets(datasheet).Cells(IntRow, IntRoute1Col)
    Route2 = Sheets(datasheet).Cells(IntRow, IntRoute2Col)
    Route3 = Sheets(datasheet).Cells(IntRow, IntRoute3Col)
    Route4 = Sheets(datasheet).Cells(IntRow, IntRoute4Col)
    IntMP0 = Sheets(datasheet).Cells(IntRow, IntMP0Col)
    IntMP1 = Sheets(datasheet).Cells(IntRow, IntMP1Col)
    IntMP2 = Sheets(datasheet).Cells(IntRow, IntMP2Col)
    IntMP3 = Sheets(datasheet).Cells(IntRow, IntMP3Col)
    IntMP4 = Sheets(datasheet).Cells(IntRow, IntMP4Col)
    IntLat = Sheets(datasheet).Cells(IntRow, IntLatCol)
    IntLong = Sheets(datasheet).Cells(IntRow, IntLongCol)
    
 '   If IntRoute3Col = 0 Then
 '       Route3 = 0
 '       IntDist3 = 0
 '       IntMP3 = 0
 '   Else
 '       Route3 = Sheets(datasheet).Cells(IntRow, IntRoute3Col)
 '       IntDist3 = Sheets(datasheet).Cells(IntRow, IntDistCol3)
 '       IntMP3 = Sheets(datasheet).Cells(IntRow, IntMP3Col)
 '   End If
    
    'Determine functional area value for each route
    
  '  If Route0 = "Local" Or Route1 = "Local" Or Route2 = "Local" Or Route3 = "Local" Or Route4 = "Local" Then
  '      GoTo FurtherDown
  '  End If
    
           ''''''''''''''''begin if''''''''''''''''''
    
    If Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Speed Limit" Then   'if the selected method of functional area is Speed Limit
        SL0 = Sheets(datasheet).Cells(IntRow, SLCol0)
        For i = 1 To 12                    'find the FA based on the speed limit for Route 0
            If SL0 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                FADist0 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                Exit For
            End If
        Next i
        
        If Sheets(datasheet).Cells(IntRow, SLCol1) = "" Then  'if the speed limit for Route 1 is blank then make the SL and FA dist for Route 1 the same as Route 0 (this is for the max/min SL accuracy)
            SL1 = SL0
            FADist1 = FADist0
        Else
            SL1 = Sheets(datasheet).Cells(IntRow, SLCol1)
            For i = 1 To 12      'find the FA based on the speed limit for Route 1 if the speed limit was given
                If SL1 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                    FADist1 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                    Exit For
                End If
            Next i
        End If
        
        If Sheets(datasheet).Cells(IntRow, SLCol2) = "" Then   'if the speed limit for Route 2 is blank then make the SL and FA dist for Route 2 the same as Route 0 (this is for the max/min SL accuracy)
            SL2 = SL0
            FADist2 = FADist0
        Else
            SL2 = Sheets(datasheet).Cells(IntRow, SLCol2)
            For i = 1 To 12                   'find the FA based on the speed limit for Route 2 if the speed limit was given
                If SL2 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                    FADist2 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                    Exit For
                End If
            Next i
        End If
        
        If Sheets(datasheet).Cells(IntRow, SLCol3) = "" Then     'if the speed limit for Route 3 is blank then make the SL and FA dist for Route 3 the same as Route 1 (this is for the max/min SL accuracy)
            SL3 = SL1
            FADist3 = FADist1
        Else
            SL3 = Sheets(datasheet).Cells(IntRow, SLCol3)
            For i = 1 To 12              'find the FA based on the speed limit for Route 3 if the speed limit was given
                If SL3 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                    FADist3 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                    Exit For
                End If
            Next i
        End If
        
        If Sheets(datasheet).Cells(IntRow, SLCol4) = "" Then           'if the speed limit for Route 4 is blank then make the SL and FA dist for Route 4 the same as Route 2 (this is for the max/min SL accuracy)
            SL4 = SL2
            FADist4 = FADist2
        Else
            SL4 = Sheets(datasheet).Cells(IntRow, SLCol4)
            For i = 1 To 12               'find the FA based on the speed limit for Route 4 if the speed limit was given
                If SL4 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                    FADist4 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                    Exit For
                End If
            Next i
        End If
        
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Functional Class" Then       'still needs update   'if the selected method of functional area is functional area
        FC1 = Sheets(datasheet).Cells(IntRow, FCCodeCol1)
        For i = 1 To 4
            If FC1 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                FADist1 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                Exit For
            End If
        Next i
        
        If Sheets(datasheet).Cells(IntRow, FCCodeCol2) = "" Then
            FC2 = FC1
            FADist2 = FADist1
        Else
            FC2 = Sheets(datasheet).Cells(IntRow, FCCodeCol2)
            For i = 1 To 4
                If FC2 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                    FADist2 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                    Exit For
                End If
            Next i
        End If
        
        If Sheets(datasheet).Cells(IntRow, FCCodeCol3) = "" Then
            FC3 = FC2
            FADist3 = FADist2
        Else
            FC3 = Sheets(datasheet).Cells(IntRow, FCCodeCol3)
            For i = 1 To 4
                If FC3 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                    FADist3 = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                    Exit For
                End If
            Next i
        End If
        
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Urban Code" Then               'still needs update      'if the selected method of functional area is urban code
        UCDesc = Sheets(datasheet).Cells(IntRow, UCcol)
        
        If UCDesc = "Rurall" Then
            FADist1 = Sheets("Inputs").Cells(FArow + 2, UICPMcol + 1)
        ElseIf UCDesc = "Small Urban" Then
            FADist1 = Sheets("Inputs").Cells(FArow + 3, UICPMcol + 1)
        Else
            FADist1 = Sheets("Inputs").Cells(FArow + 4, UICPMcol + 1)
        End If
        
        FADist2 = FADist1
        FADist3 = FADist1
    
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Fixed Length" Then     'if the selected method of functional area is fixed length        'added by Camille on May 9, 2019
        FADist0 = Sheets("Inputs").Cells(FArow + 2, UICPMcol + 1)
        FADist1 = FADist0
        FADist2 = FADist0
        FADist3 = FADist0
        FADist4 = FADist0
        
             ''''''''''''''end if''''''''''''''''''
    End If
    
    'Add internal intersection size to intersection radius
    If IntDist0 = 0 Then
        FADist0 = FADist0 + ((MaxRoadW + MinRoadW) / 4)
    Else
        FADist0 = FADist0 + (IntDist0 / 2)
    End If
    
    If IntDist1 = 0 Then
        FADist1 = FADist1 + ((MaxRoadW + MinRoadW) / 4)
    Else
        FADist1 = FADist1 + (IntDist1 / 2)
    End If
    
    If IntDist2 = 0 Then
        FADist2 = FADist2 + ((MaxRoadW + MinRoadW) / 4)
    Else
        FADist2 = FADist2 + (IntDist2 / 2)
    End If
    
    If IntDist3 = 0 Then
        FADist3 = FADist3 + ((MaxRoadW + MinRoadW) / 4)
    Else
        FADist3 = FADist3 + (IntDist3 / 2)
    End If
    
     If IntDist4 = 0 Then
        FADist4 = FADist4 + ((MaxRoadW + MinRoadW) / 4)
    Else
        FADist4 = FADist4 + (IntDist4 / 2)
    End If
    
    'Convert distances from feet to miles
    FADist0 = FADist0 / 5280
    FADist1 = FADist1 / 5280
    FADist2 = FADist2 / 5280
    FADist3 = FADist3 / 5280
    FADist4 = FADist4 / 5280
    
    'Change max and min roadway width values
    If IntDist0 >= IntDist1 And IntDist0 >= IntDist2 And IntDist0 >= IntDist3 And IntDist0 >= IntDist4 Then
        MaxDist = IntDist0
    ElseIf IntDist1 >= IntDist0 And IntDist1 >= IntDist2 And IntDist1 >= IntDist3 And IntDist1 >= IntDist4 Then
        MaxDist = IntDist1
    ElseIf IntDist2 >= IntDist0 And IntDist2 >= IntDist1 And IntDist2 >= IntDist3 And IntDist2 >= IntDist4 Then
        MaxDist = IntDist2
    ElseIf IntDist3 >= IntDist0 And IntDist3 >= IntDist1 And IntDist3 >= IntDist2 And IntDist3 >= IntDist4 Then
        MaxDist = IntDist3
    Else
        MaxDist = IntDist4
    End If
    
    If IntDist0 <= IntDist1 And IntDist0 <= IntDist2 And IntDist0 <= IntDist3 And IntDist0 <= IntDist4 And IntDist0 <> 0 Then
        MinDist = IntDist0
    ElseIf IntDist1 <= IntDist0 And IntDist1 <= IntDist2 And IntDist1 <= IntDist3 And IntDist1 <= IntDist4 And IntDist1 <> 0 Then
        MinDist = IntDist1
    ElseIf IntDist2 <= IntDist0 And IntDist2 <= IntDist1 And IntDist2 <= IntDist3 And IntDist2 <= IntDist4 And IntDist2 <> 0 Then
        MinDist = IntDist2
    ElseIf IntDist3 <= IntDist0 And IntDist3 <= IntDist1 And IntDist3 <= IntDist2 And IntDist3 <= IntDist4 And IntDist3 <> 0 Then
        MinDist = IntDist3
    Else
        MinDist = IntDist4
    End If
    
        
    If MaxRoadW < MaxDist Then
        MaxRoadW = MaxDist
        For i = 0 To NumYears - 1
            Sheets(datasheet).Cells(IntRow + i, MaxRoadWCol) = MaxRoadW
        Next i
    End If
    
    If MinRoadW > MinDist And MinDist <> 0 Then
        MinRoadW = MinDist
        For i = 0 To NumYears - 1
            Sheets(datasheet).Cells(IntRow, MinRoadWcol) = MinRoadW
        Next i
    End If
    
    'Determine severities
    sevstring = Sheets("Inputs").Cells(FArow - 4, UICPMcol + 1)   'FArow - 4 as of the updates with SRtoFedAid etc. (used to be FArow - 1)
    For i = 1 To Len(sevstring)
        Sevs(i) = Mid(sevstring, i, 1)
    Next i
    
    'Assign crashes to intersection
    CrashRow = 2
    Do Until Sheets(CrashSheet).Cells(CrashRow, crashroutecol) = ""
        
        If Route1 = "Local" Then
            Route1 = 22647
        End If
        
        If Route2 = "Local" Then
            Route2 = 22647
        End If
        
        If Route3 = "Local" Then
            Route3 = 22647
        End If
        
        If Route4 = "Local" Then
            Route4 = 22647
        End If
        
        CrashSev = Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol)
        CrashYear = year(Sheets(CrashSheet).Cells(CrashRow, CrashDateCol))
        CrashRoute = Int(Left(Sheets(CrashSheet).Cells(CrashRow, crashroutecol), 4))
        CrashMP = Sheets(CrashSheet).Cells(CrashRow, crashMPcol)
        CrashLat = Sheets(CrashSheet).Cells(CrashRow, CrashLatCol)
        CrashLong = Sheets(CrashSheet).Cells(CrashRow, CrashLongCol)
        
        Dim GoAhead As Boolean
        GoAhead = False
        
        
        If (CrashRoute = Route0 And ((CrashMP >= IntMP0) And (CrashMP < (IntMP0 + FADist0)) Or ((CrashMP < IntMP0) And (CrashMP > (IntMP0 - FADist0))))) Or _
        (CrashRoute = Route1 And ((CrashMP >= IntMP1) And (CrashMP < (IntMP1 + FADist1)) Or ((CrashMP < IntMP1) And (CrashMP > (IntMP1 - FADist1))))) Or _
        (CrashRoute = Route2 And ((CrashMP >= IntMP2) And (CrashMP < (IntMP2 + FADist2)) Or ((CrashMP < IntMP2) And (CrashMP > (IntMP2 - FADist2))))) Or _
        (CrashRoute = Route3 And ((CrashMP >= IntMP3) And (CrashMP < (IntMP3 + FADist3)) Or ((CrashMP < IntMP3) And (CrashMP > (IntMP3 - FADist3))))) Or _
        (CrashRoute = Route4 And ((CrashMP >= IntMP4) And (CrashMP < (IntMP4 + FADist4)) Or ((CrashMP < IntMP4) And (CrashMP > (IntMP4 - FADist4))))) Then
            GoAhead = True
        
        ElseIf Route1 = 22647 Then
            FADistLat = FADist1 / 111034.605
            FADistLong = FADist1 / 280162.911
            If (CrashLat < (IntLat + FADistLat)) And (CrashLat > (IntLat - FADistLat)) And (CrashLong < (IntLong + FADistLong)) And (CrashLong > (IntLong - FADistLong)) Then
                GoAhead = True
            End If
            
        ElseIf Route2 = 22647 Then
            FADistLat = FADist2 / 111034.605
            FADistLong = FADist2 / 280162.911
            If (CrashLat < (IntLat + FADistLat)) And (CrashLat > (IntLat - FADistLat)) And (CrashLong < (IntLong + FADistLong)) And (CrashLong > (IntLong - FADistLong)) Then
                GoAhead = True
            End If
        
        ElseIf Route3 = 22647 Then
            FADistLat = FADist3 / 111034.605
            FADistLong = FADist3 / 280162.911
            If (CrashLat < (IntLat + FADistLat)) And (CrashLat > (IntLat - FADistLat)) And (CrashLong < (IntLong + FADistLong)) And (CrashLong > (IntLong - FADistLong)) Then
                GoAhead = True
            End If
            
        ElseIf Route4 = 22647 Then
            FADistLat = FADist4 / 111034.605
            FADistLong = FADist4 / 280162.911
            If (CrashLat < (IntLat + FADistLat)) And (CrashLat > (IntLat - FADistLat)) And (CrashLong < (IntLong + FADistLong)) And (CrashLong > (IntLong - FADistLong)) Then
                GoAhead = True
            End If
            
        Else
            GoAhead = False
            
        End If
        
   
        If GoAhead = True Then
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
            
                If IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, ColSumTotCrash) = Sheets(datasheet).Cells(IntRow + i, ColSumTotCrash) + 1
                    'Tally severity numbers for all the crashes at each intersection
                    If CrashSev = 5 Then
                        Sheets(datasheet).Cells(IntRow + i, ColSev5) = Sheets(datasheet).Cells(IntRow + i, ColSev5) + 1
                    ElseIf CrashSev = 4 Then
                        Sheets(datasheet).Cells(IntRow + i, ColSev4) = Sheets(datasheet).Cells(IntRow + i, ColSev4) + 1
                    ElseIf CrashSev = 3 Then
                        Sheets(datasheet).Cells(IntRow + i, ColSev3) = Sheets(datasheet).Cells(IntRow + i, ColSev3) + 1
                    ElseIf CrashSev = 2 Then
                        Sheets(datasheet).Cells(IntRow + i, ColSev2) = Sheets(datasheet).Cells(IntRow + i, ColSev2) + 1
                    ElseIf CrashSev = 1 Then
                        Sheets(datasheet).Cells(IntRow + i, ColSev1) = Sheets(datasheet).Cells(IntRow + i, ColSev1) + 1
                    End If
                End If
            Next i
            
            'what does this do?? 'Camille: I think it is checking to make sure it doesn't start tallying numbers for severities that were not checked in the user form.
            '                     'then maybe the sevstring get erased?
            If CrashSev = Sevs(1) Or CrashSev = Sevs(2) Or CrashSev = Sevs(3) Or CrashSev = Sevs(4) Or CrashSev = Sevs(5) Then
            
                'Input crash data in parameters sheet
                For j = 0 To NumYears - 1
                    IntYear = Sheets(datasheet).Cells(IntRow + j, YearCol).Value

                    If IntYear = CrashYear Then
                        Sheets(PSheet).Cells(prow, 1) = IntNum
                        For i = 1 To CrashCol + 1
                            Sheets(PSheet).Cells(prow, i + 1) = Sheets(CrashSheet).Cells(CrashRow, i)
                        Next i
                        prow = prow + 1
                    End If
                Next j
            
                For i = 0 To NumYears - 1
                    IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                    
                    If IntYear = CrashYear Then
                        'Total crashes
                        Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = Sheets(datasheet).Cells(IntRow + i, TotCrashCol) + 1
                        
                        'Severe crashes
                        If CrashSev >= 3 Then
                            Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = Sheets(datasheet).Cells(IntRow + i, SevCrashCol) + 1
                        End If
                    
                        'Head on collisions
                        If Sheets(CrashSheet).Cells(CrashRow, CollMannerCol2) = 3 Then
                             Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) + 1
                         End If
                                              
                        'Work zone related
                         'If Sheets(CrashSheet).Cells(CrashRow, WorkZoneCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) + 1
                         'End If
                         
                        'Pedestrian Involved
                         If Sheets(CrashSheet).Cells(CrashRow, PedInvCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = Sheets(datasheet).Cells(IntRow + i, PedInvCol1) + 1
                         End If
                         
                        'Bicyclist Involved
                         If Sheets(CrashSheet).Cells(CrashRow, BikeInvCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) + 1
                         End If
                         
                        'Motorcycle Involved
                         If Sheets(CrashSheet).Cells(CrashRow, MotInvCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = Sheets(datasheet).Cells(IntRow + i, MotInvCol1) + 1
                         End If
                         
                        'Improper restraint
                         'If Sheets(CrashSheet).Cells(CrashRow, ImpResCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = Sheets(datasheet).Cells(IntRow + i, ImpResCol1) + 1
                         'End If
                         
                        'Unrestrained
                         If Sheets(CrashSheet).Cells(CrashRow, UnrestCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = Sheets(datasheet).Cells(IntRow + i, UnrestCol1) + 1
                         End If
                         
                        'DUI
                         If Sheets(CrashSheet).Cells(CrashRow, DUICol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, DUICol1) = Sheets(datasheet).Cells(IntRow + i, DUICol1) + 1
                         End If
                         
                        'Aggresive Driving
                         If Sheets(CrashSheet).Cells(CrashRow, AggressiveCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) + 1
                         End If
                         
                        'Distracted Driving
                         If Sheets(CrashSheet).Cells(CrashRow, DistractedCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = Sheets(datasheet).Cells(IntRow + i, DistractedCol1) + 1
                         End If
                         
                        'Drowsy Driving
                         If Sheets(CrashSheet).Cells(CrashRow, DrowsyCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) + 1
                         End If
                         
                        'Speed related
                         If Sheets(CrashSheet).Cells(CrashRow, SpeedCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = Sheets(datasheet).Cells(IntRow + i, SpeedCol1) + 1
                         End If
                         
                        'Intersection related
                         If Sheets(CrashSheet).Cells(CrashRow, IntRelCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = Sheets(datasheet).Cells(IntRow + i, IntRelCol1) + 1
                         End If
                         
                        'Adverse weather
                         If Sheets(CrashSheet).Cells(CrashRow, WeatherCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = Sheets(datasheet).Cells(IntRow + i, WeatherCol1) + 1
                         End If
                         
                        'Adverse roadway surface conditions
                         If Sheets(CrashSheet).Cells(CrashRow, AdvRoadCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) + 1
                         End If
                         
                        'Roadway geometry related
                         If Sheets(CrashSheet).Cells(CrashRow, RoadGeomCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) + 1
                         End If
                         
                        'Wild animal related
                         If Sheets(CrashSheet).Cells(CrashRow, WAnimalCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) + 1
                         End If
                         
                        'Domestic animal related
                         If Sheets(CrashSheet).Cells(CrashRow, DAnimalCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) + 1
                         End If
                         
                        'Roadway departure
                         If Sheets(CrashSheet).Cells(CrashRow, RoadDepartCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) + 1
                         End If
                         
                        'Overturn/rollover
                         If Sheets(CrashSheet).Cells(CrashRow, OverturnCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = Sheets(datasheet).Cells(IntRow + i, OverturnCol1) + 1
                         End If
                         
                        'Commercial motor vehicle involved
                         If Sheets(CrashSheet).Cells(CrashRow, CommVehCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = Sheets(datasheet).Cells(IntRow + i, CommVehCol1) + 1
                         End If
                         
                        'Interstate highway
                         If Sheets(CrashSheet).Cells(CrashRow, InterstateCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = Sheets(datasheet).Cells(IntRow + i, InterstateCol1) + 1
                         End If
                         
                        'Teenage driver involved
                         If Sheets(CrashSheet).Cells(CrashRow, TeenCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, TeenCol1) = Sheets(datasheet).Cells(IntRow + i, TeenCol1) + 1
                          End If
                         
                        'Older driver involved
                         If Sheets(CrashSheet).Cells(CrashRow, OldCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, OldCol1) = Sheets(datasheet).Cells(IntRow + i, OldCol1) + 1
                         End If
                         
                        'Urban county
                         'If Sheets(CrashSheet).Cells(CrashRow, UrbanCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = Sheets(datasheet).Cells(IntRow + i, UrbanCol1) + 1
                         'End If
                         
                        'Night dark condition
                         If Sheets(CrashSheet).Cells(CrashRow, NightCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, NightCol1) = Sheets(datasheet).Cells(IntRow + i, NightCol1) + 1
                         End If
                         
                        'Single vehicle
                         If Sheets(CrashSheet).Cells(CrashRow, SingleVehCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) + 1
                         End If
                         
                        'Train involved
                         If Sheets(CrashSheet).Cells(CrashRow, TrainCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, TrainCol1) = Sheets(datasheet).Cells(IntRow + i, TrainCol1) + 1
                         End If
                         
                        'Railroad crossing
                         If Sheets(CrashSheet).Cells(CrashRow, RailCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, RailCol1) = Sheets(datasheet).Cells(IntRow + i, RailCol1) + 1
                         End If
                         
                        'Transit vehicle involved
                         If Sheets(CrashSheet).Cells(CrashRow, TransitCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, TransitCol1) = Sheets(datasheet).Cells(IntRow + i, TransitCol1) + 1
                         End If
                         
                        'Collision with fixed object
                         If Sheets(CrashSheet).Cells(CrashRow, FixedObCol2) = "Y" Then
                             Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = Sheets(datasheet).Cells(IntRow + i, FixedObCol1) + 1
                         End If
                    End If
                Next i
            End If
        End If
        CrashRow = CrashRow + 1
    Loop
    
If Route1 = "22647" Then
    Route1 = "Local"
End If
        
If Route2 = "22647" Then
    Route2 = "Local"
End If
    
If Route3 = "22647" Then
    Route3 = "Local"
End If

If Route4 = "22647" Then
    Route4 = "Local"
End If
        
'FurtherDown:
    IntRow = IntRow + NumYears
Loop

'Delete IntDist columns (3)
Sheets(datasheet).Columns(IntDistCol1).EntireColumn.Delete
Sheets(datasheet).Columns(IntDistCol1).EntireColumn.Delete
Sheets(datasheet).Columns(IntDistCol1).EntireColumn.Delete

End Sub

Sub SumCAMSbyYear()

'Declare variables
Dim SegRow As Long, CrashRow As Long
Dim CrashLatCol As Integer, CrashLongCol As Integer
Dim CrashLat As Double, CrashLong As Double
Dim NumYears As Integer
Dim datasheet As String
Dim CrashSheet As String
Dim row1 As Long, col1 As Long    'counters for datasheet
Dim row2 As Long, col2 As Long    'counters for CrashSheet
Dim minyr As String, maxyr As String
Dim Crshfp As String

Dim YearCol, TotCrashCol, SevCrashCol, PedInvCol1, BikeInvCol1, MotInvCol1, ImpResCol1, UnrestCol1, DUICol1 As Long
Dim AggressiveCol1, DistractedCol1, DrowsyCol1, SpeedCol1, IntRelCol1, WeatherCol1, AdvRoadCol1, RoadGeomCol1 As Long
Dim WAnimalCol1, DAnimalCol1, RoadDepartCol1, OverturnCol1, CommVehCol1, InterstateCol1, TeenCol1, OldCol1 As Long
Dim UrbanCol1, NightCol1, SingleVehCol1, TrainCol1, RailCol1, TransitCol1, FixedObCol1, WorkZoneCol1, WZworkerCol1, HeadOnCol As Long
Dim SegYear As String

Dim crashlabelcol As Long, crashMPcol As Long
Dim CrashDateCol As Long, CrashSeverityCol As Long
Dim PedInvCol2, BikeInvCol2, MotInvCol2, ImpResCol2, UnrestCol2, DUICol2 As Long
Dim crashroutecol As Long, crashdircol As Long
Dim AggressiveCol2, DistractedCol2, DrowsyCol2, SpeedCol2, IntRelCol2, WeatherCol2, AdvRoadCol2, RoadGeomCol2 As Long
Dim WAnimalCol2, DAnimalCol2, RoadDepartCol2, OverturnCol2, CommVehCol2, InterstateCol2, TeenCol2, OldCol2 As Long
Dim UrbanCol2, NightCol2, SingleVehCol2, TrainCol2, RailCol2, TransitCol2, FixedObCol2, WorkZoneCol2, WZworkerCol2, MannerCollisionCol As Long
Dim CrashYear As String

Dim SegIDcol As Integer, SegID As Integer, SegRoute As Integer
Dim i As Long, j As Long
Dim sevstring As String
ReDim Sevs(1 To 5) As Integer
Dim Label As String, LabelCol As Integer
Dim BMP As Double, EMP As Double
Dim BEGMPcol As Integer, ENDMPcol As Integer
Dim GoAhead As Boolean
Dim NextSegment As Boolean
Dim IntstMrkrBool As Boolean
Dim InterstateMarker As Long
Dim EndofIntstDir As Boolean
Dim ExitDoAfter As Boolean

Dim ColSumTotCrash As Integer
Dim ColSev5, ColSev4, ColSev3, ColSev2, ColSev1 As Integer
Dim CrashSev As Integer
Dim CrashRoute As Integer
Dim CrashMP As Double
ReDim SevCount(1 To 5) As Integer
Dim CrashID As String
Dim prow As Long
Dim PSheet As String

'Assign initial values
SegRow = 2
CrashRow = 2
prow = 5

datasheet = "CAMSinput"
CrashSheet = "CrashInput"
PSheet = "Parameters"

'Find columns and paste appropriate headers to the segment sheet                'FLAGGED: This hard codes whether the column is necessary.
    'columns on the crash sheet
        'columns to go only on the Parameters sheet
crashlabelcol = FindColumn("LABEL", CrashSheet, 1, True)
crashroutecol = FindColumn("ROUTE_ID", CrashSheet, 1, True)
crashdircol = FindColumn("DIRECTION", CrashSheet, 1, True)
crashMPcol = FindColumn("MILEPOINT", CrashSheet, 1, True)
CrashDateCol = FindColumn("CRASH_DATETIME", CrashSheet, 1, True)
CrashSeverityCol = FindColumn("CRASH_SEVERITY_ID", CrashSheet, 1, True)
CrashLatCol = FindColumn("LATITUDE", CrashSheet, 1, True)
CrashLongCol = FindColumn("LONGITUDE", CrashSheet, 1, True)
        'columns that go on the input/segment sheet
WorkZoneCol2 = FindColumn("WORK_ZONE_RELATED_YNU", CrashSheet, 1, True)
WZworkerCol2 = FindColumn("WORK_ZONE_WORKER_PRESENT_YNU", CrashSheet, 1, False)
PedInvCol2 = FindColumn("PEDESTRIAN_INVOLVED", CrashSheet, 1, True)
BikeInvCol2 = FindColumn("BICYCLIST_INVOLVED", CrashSheet, 1, True)
MotInvCol2 = FindColumn("MOTORCYCLE_INVOLVED", CrashSheet, 1, True)
ImpResCol2 = FindColumn("IMPROPER_RESTRAINT", CrashSheet, 1, True)
UnrestCol2 = FindColumn("UNRESTRAINED", CrashSheet, 1, True)
DUICol2 = FindColumn("DUI", CrashSheet, 1, True)
AggressiveCol2 = FindColumn("AGGRESSIVE_DRIVING", CrashSheet, 1, True)
DistractedCol2 = FindColumn("DISTRACTED_DRIVING", CrashSheet, 1, True)
DrowsyCol2 = FindColumn("DROWSY_DRIVING", CrashSheet, 1, True)
SpeedCol2 = FindColumn("SPEED_RELATED", CrashSheet, 1, True)
IntRelCol2 = FindColumn("INTERSECTION_RELATED", CrashSheet, 1, True)
WeatherCol2 = FindColumn("ADVERSE_WEATHER", CrashSheet, 1, True)
AdvRoadCol2 = FindColumn("ADVERSE_ROADWAY_SURF_CONDITION", CrashSheet, 1, True)
RoadGeomCol2 = FindColumn("ROADWAY_GEOMETRY_RELATED", CrashSheet, 1, True)
WAnimalCol2 = FindColumn("WILD_ANIMAL_RELATED", CrashSheet, 1, True)
DAnimalCol2 = FindColumn("DOMESTIC_ANIMAL_RELATED", CrashSheet, 1, True)
RoadDepartCol2 = FindColumn("ROADWAY_DEPARTURE", CrashSheet, 1, True)
OverturnCol2 = FindColumn("OVERTURN_ROLLOVER", CrashSheet, 1, True)
CommVehCol2 = FindColumn("COMMERCIAL_MOTOR_VEH_INVOLVED", CrashSheet, 1, True)
InterstateCol2 = FindColumn("INTERSTATE_HIGHWAY", CrashSheet, 1, True)
TeenCol2 = FindColumn("TEENAGE_DRIVER_INVOLVED", CrashSheet, 1, True)
OldCol2 = FindColumn("OLDER_DRIVER_INVOLVED", CrashSheet, 1, True)
UrbanCol2 = FindColumn("URBAN_COUNTY", CrashSheet, 1, True)
NightCol2 = FindColumn("NIGHT_DARK_CONDITION", CrashSheet, 1, True)
SingleVehCol2 = FindColumn("SINGLE_VEHICLE", CrashSheet, 1, True)
TrainCol2 = FindColumn("TRAIN_INVOLVED", CrashSheet, 1, True)
RailCol2 = FindColumn("RAILROAD_CROSSING", CrashSheet, 1, True)
TransitCol2 = FindColumn("TRANSIT_VEHICLE_INVOLVED", CrashSheet, 1, True)
FixedObCol2 = FindColumn("COLLISION_WITH_FIXED_OBJECT", CrashSheet, 1, True)
MannerCollisionCol = FindColumn("MANNER_COLLISION_ID", CrashSheet, 1, True)

    'columns on the segment sheet
        'columns from the roadway data
SegIDcol = FindColumn("SEG_ID", datasheet, 1, True)
LabelCol = FindColumn("LABEL", datasheet, 1, True)
BEGMPcol = FindColumn("BEG_MILEPOINT", datasheet, 1, True)
ENDMPcol = FindColumn("END_MILEPOINT", datasheet, 1, True)
YearCol = FindColumn("YEAR", datasheet, 1, True)
'SegLenCol = FindColumn("Seg_Length", datasheet, 1, True)
'FCCodeCol = FindColumn("FCCode", datasheet, 1, True)
'FCTypeCol = FindColumn("FCType", datasheet, 1, True)
'UCcol = FindColumn("Urban_Rural", datasheet, 1, True)
'UrbanCodeCol = FindColumn("Urban_Ru_1", datasheet, 1, True)
'SLcol = FindColumn("SPEED_LIMIT", datasheet, 1, True)
        'columns added to the sheet
TotCrashCol = FindColumn("Total_Crashes", datasheet, 1, True)
SevCrashCol = FindColumn("Severe_Crashes", datasheet, 1, True)
ColSumTotCrash = FindColumn("Sum_Total_Crashes", datasheet, 1, True)
ColSev5 = FindColumn("Sev_5_Crashes", datasheet, 1, True)
ColSev4 = FindColumn("Sev_4_Crashes", datasheet, 1, True)
ColSev3 = FindColumn("Sev_3_Crashes", datasheet, 1, True)
ColSev2 = FindColumn("Sev_2_Crashes", datasheet, 1, True)
ColSev1 = FindColumn("Sev_1_Crashes", datasheet, 1, True)
        'columns from the crash sheet
    'Find first blank column in segment data
    col1 = 1
    Do Until Sheets(datasheet).Cells(1, col1) = ""
        col1 = col1 + 1
    Loop
    Sheets(datasheet).Cells(1, col1) = "ADVERSE_ROADWAY_SURF_CONDITION"
    col1 = col1 + 1
AdvRoadCol1 = FindColumn("ADVERSE_ROADWAY_SURF_CONDITION", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "ADVERSE_WEATHER"
    col1 = col1 + 1
WeatherCol1 = FindColumn("ADVERSE_WEATHER", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "AGGRESSIVE_DRIVING"
    col1 = col1 + 1
AggressiveCol1 = FindColumn("AGGRESSIVE_DRIVING", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "BICYCLIST_INVOLVED"
    col1 = col1 + 1
BikeInvCol1 = FindColumn("BICYCLIST_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "COLLISION_WITH_FIXED_OBJECT"
    col1 = col1 + 1
FixedObCol1 = FindColumn("COLLISION_WITH_FIXED_OBJECT", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "COMMERCIAL_MOTOR_VEH_INVOLVED"
    col1 = col1 + 1
CommVehCol1 = FindColumn("COMMERCIAL_MOTOR_VEH_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "DISTRACTED_DRIVING"
    col1 = col1 + 1
DistractedCol1 = FindColumn("DISTRACTED_DRIVING", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "DOMESTIC_ANIMAL_RELATED"
    col1 = col1 + 1
DAnimalCol1 = FindColumn("DOMESTIC_ANIMAL_RELATED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "DROWSY_DRIVING"
    col1 = col1 + 1
DrowsyCol1 = FindColumn("DROWSY_DRIVING", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "DUI"
    col1 = col1 + 1
DUICol1 = FindColumn("DUI", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "HEADON_COLLISION"
    col1 = col1 + 1
HeadOnCol = FindColumn("HEADON_COLLISION", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "IMPROPER_RESTRAINT"
    col1 = col1 + 1
ImpResCol1 = FindColumn("IMPROPER_RESTRAINT", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "INTERSECTION_RELATED"
    col1 = col1 + 1
IntRelCol1 = FindColumn("INTERSECTION_RELATED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "INTERSTATE_HIGHWAY"
    col1 = col1 + 1
InterstateCol1 = FindColumn("INTERSTATE_HIGHWAY", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "MOTORCYCLE_INVOLVED"
    col1 = col1 + 1
MotInvCol1 = FindColumn("MOTORCYCLE_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "NIGHT_DARK_CONDITION"
    col1 = col1 + 1
NightCol1 = FindColumn("NIGHT_DARK_CONDITION", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "OLDER_DRIVER_INVOLVED"
    col1 = col1 + 1
OldCol1 = FindColumn("OLDER_DRIVER_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "OVERTURN_ROLLOVER"
    col1 = col1 + 1
OverturnCol1 = FindColumn("OVERTURN_ROLLOVER", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "PEDESTRIAN_INVOLVED"
    col1 = col1 + 1
PedInvCol1 = FindColumn("PEDESTRIAN_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "RAILROAD_CROSSING"
    col1 = col1 + 1
RailCol1 = FindColumn("RAILROAD_CROSSING", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "ROADWAY_DEPARTURE"
    col1 = col1 + 1
RoadDepartCol1 = FindColumn("ROADWAY_DEPARTURE", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "ROADWAY_GEOMETRY_RELATED"
    col1 = col1 + 1
RoadGeomCol1 = FindColumn("ROADWAY_GEOMETRY_RELATED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "SINGLE_VEHICLE"
    col1 = col1 + 1
SingleVehCol1 = FindColumn("SINGLE_VEHICLE", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "SPEED_RELATED"
    col1 = col1 + 1
SpeedCol1 = FindColumn("SPEED_RELATED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "TEENAGE_DRIVER_INVOLVED"
    col1 = col1 + 1
TeenCol1 = FindColumn("TEENAGE_DRIVER_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "TRAIN_INVOLVED"
    col1 = col1 + 1
TrainCol1 = FindColumn("TRAIN_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "TRANSIT_VEHICLE_INVOLVED"
    col1 = col1 + 1
TransitCol1 = FindColumn("TRANSIT_VEHICLE_INVOLVED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "UNRESTRAINED"
    col1 = col1 + 1
UnrestCol1 = FindColumn("UNRESTRAINED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "URBAN_COUNTY"
    col1 = col1 + 1
UrbanCol1 = FindColumn("URBAN_COUNTY", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "WILD_ANIMAL_RELATED"
    col1 = col1 + 1
WAnimalCol1 = FindColumn("WILD_ANIMAL_RELATED", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "WORK_ZONE_RELATED_YNU"
    col1 = col1 + 1
WorkZoneCol1 = FindColumn("WORK_ZONE_RELATED_YNU", datasheet, 1, True)
    Sheets(datasheet).Cells(1, col1) = "WORK_ZONE_WORKER_PRESENT_YNU"
    col1 = col1 + 1
WZworkerCol1 = FindColumn("WORK_ZONE_WORKER_PRESENT_YNU", datasheet, 1, False)

Sheets(datasheet).Activate

'Find number of years of data
NumYears = 1
Do Until Sheets(datasheet).Cells(SegRow + NumYears, YearCol) = Sheets(datasheet).Cells(SegRow, YearCol)
    NumYears = NumYears + 1
Loop

'Insert zeros for all values
SegRow = 2
Do Until Sheets(datasheet).Cells(SegRow, 1) = ""
    Sheets(datasheet).Cells(SegRow, TotCrashCol) = 0
    Sheets(datasheet).Cells(SegRow, SevCrashCol) = 0
    Sheets(datasheet).Cells(SegRow, ColSumTotCrash) = 0
    Sheets(datasheet).Cells(SegRow, ColSev5) = 0
    Sheets(datasheet).Cells(SegRow, ColSev4) = 0
    Sheets(datasheet).Cells(SegRow, ColSev3) = 0
    Sheets(datasheet).Cells(SegRow, ColSev2) = 0
    Sheets(datasheet).Cells(SegRow, ColSev1) = 0
    Sheets(datasheet).Cells(SegRow, PedInvCol1) = 0
    Sheets(datasheet).Cells(SegRow, BikeInvCol1) = 0
    Sheets(datasheet).Cells(SegRow, MotInvCol1) = 0
    Sheets(datasheet).Cells(SegRow, ImpResCol1) = 0
    Sheets(datasheet).Cells(SegRow, UnrestCol1) = 0
    Sheets(datasheet).Cells(SegRow, DUICol1) = 0
    Sheets(datasheet).Cells(SegRow, AggressiveCol1) = 0
    Sheets(datasheet).Cells(SegRow, DistractedCol1) = 0
    Sheets(datasheet).Cells(SegRow, DrowsyCol1) = 0
    Sheets(datasheet).Cells(SegRow, SpeedCol1) = 0
    Sheets(datasheet).Cells(SegRow, IntRelCol1) = 0
    Sheets(datasheet).Cells(SegRow, WeatherCol1) = 0
    Sheets(datasheet).Cells(SegRow, AdvRoadCol1) = 0
    Sheets(datasheet).Cells(SegRow, RoadGeomCol1) = 0
    Sheets(datasheet).Cells(SegRow, WAnimalCol1) = 0
    Sheets(datasheet).Cells(SegRow, DAnimalCol1) = 0
    Sheets(datasheet).Cells(SegRow, RoadDepartCol1) = 0
    Sheets(datasheet).Cells(SegRow, OverturnCol1) = 0
    Sheets(datasheet).Cells(SegRow, CommVehCol1) = 0
    Sheets(datasheet).Cells(SegRow, InterstateCol1) = 0
    Sheets(datasheet).Cells(SegRow, TeenCol1) = 0
    Sheets(datasheet).Cells(SegRow, OldCol1) = 0
    Sheets(datasheet).Cells(SegRow, UrbanCol1) = 0
    Sheets(datasheet).Cells(SegRow, NightCol1) = 0
    Sheets(datasheet).Cells(SegRow, SingleVehCol1) = 0
    Sheets(datasheet).Cells(SegRow, TrainCol1) = 0
    Sheets(datasheet).Cells(SegRow, RailCol1) = 0
    Sheets(datasheet).Cells(SegRow, TransitCol1) = 0
    Sheets(datasheet).Cells(SegRow, FixedObCol1) = 0
    Sheets(datasheet).Cells(SegRow, WorkZoneCol1) = 0
    Sheets(datasheet).Cells(SegRow, HeadOnCol) = 0
    Sheets(datasheet).Cells(SegRow, WZworkerCol1) = 0
    SegRow = SegRow + 1
Loop

' Sort the segment data by Label and BegMP
' Identify extent of data for sorting purposes.
row1 = Range("A1").End(xlDown).row

' Sort the segment data
With Sheets(datasheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Sheets(datasheet).Cells(2, LabelCol), Sheets(datasheet).Cells(row1, LabelCol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Sheets(datasheet).Cells(2, BEGMPcol), Sheets(datasheet).Cells(row1, BEGMPcol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Sheets(datasheet).Cells(1, 1), Sheets(datasheet).Cells(row1, col1))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Sheets("Home").Activate

' Sort the crash data by Route and MP
' Identify the extent of the crash data
col2 = Sheets(CrashSheet).Range("A1").End(xlToRight).Column
row2 = Sheets(CrashSheet).Range("A1").End(xlDown).row

' Sort the crash data
With Sheets(CrashSheet).Sort
    .SortFields.Clear
    .SortFields.Add _
        Key:=Range(Sheets(CrashSheet).Cells(2, crashroutecol), Sheets(CrashSheet).Cells(row2, crashroutecol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add _
        Key:=Range(Sheets(CrashSheet).Cells(2, crashMPcol), Sheets(CrashSheet).Cells(row2, crashMPcol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range(Sheets(CrashSheet).Cells(1, 1), Sheets(CrashSheet).Cells(row2, col2))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Add Parameters Sheet
Sheets.Add After:=Sheets(datasheet)
ActiveSheet.Name = "Parameters"
Sheets(PSheet).Cells(1, 1) = "Severities"
Sheets(PSheet).Cells(1, 2) = Sheets("Inputs").Cells(32, 13)
Sheets(PSheet).Cells(2, 1) = "Functional Area Definition"
Crshfp = Sheets("Inputs").Cells(6, 13)
Sheets(PSheet).Cells(2, 2) = Mid(Crshfp, InStr(1, Crshfp, "by"), InStr(InStr(1, Crshfp, "by"), Crshfp, "_") - InStr(1, Crshfp, "by"))

Sheets(PSheet).Cells(3, 1) = "Selected Years"
Sheets(PSheet).Cells(3, 2) = Sheets("Inputs").Cells(31, 13)
minyr = Left(Sheets("Inputs").Cells(31, 13), 4)
maxyr = Right(Sheets("Inputs").Cells(31, 13), 4)

'Enter crash headers into parameters sheet
Sheets(PSheet).Cells(4, 1) = "SEG_ID"
For i = 1 To col2 + 1
    Sheets(PSheet).Cells(4, i + 1) = Sheets(CrashSheet).Cells(1, i)
Next i

'Loop through segments and assign crashes
SegRow = 2
NextSegment = False
IntstMrkrBool = False
InterstateMarker = 0
Do While Sheets(datasheet).Cells(SegRow, 1) <> ""
    Label = Sheets(datasheet).Cells(SegRow, LabelCol)
    SegRoute = Int(Left(Label, 4))
    BMP = Sheets(datasheet).Cells(SegRow, BEGMPcol)
    EMP = Sheets(datasheet).Cells(SegRow, ENDMPcol)
    SegID = Sheets(datasheet).Cells(SegRow, SegIDcol)
        
    'Determine severities
    sevstring = Sheets("Inputs").Cells(32, 13)   'FArow - 4 as of the updates with SRtoFedAid etc. (used to be FArow - 1)
    For i = 1 To Len(sevstring)
        Sevs(i) = Mid(sevstring, i, 1)
    Next i
    
    'assign CrashRow
    If NextSegment = False Then
        CrashRow = 2  '201152
        'SegRow = 17956
    End If
    If EndofIntstDir = True Then
        CrashRow = InterstateMarker
    End If

    NextSegment = False
    
    'Assign crashes to segment
    'do until all the crashes have been assigned
    Do Until Sheets(CrashSheet).Cells(CrashRow, crashlabelcol) = ""
        
        'set variables
        CrashSev = Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol)
        CrashYear = year(Sheets(CrashSheet).Cells(CrashRow, CrashDateCol))
        CrashRoute = Int(Left(Sheets(CrashSheet).Cells(CrashRow, crashlabelcol), 4))
        CrashMP = Sheets(CrashSheet).Cells(CrashRow, crashMPcol)
        CrashLat = Sheets(CrashSheet).Cells(CrashRow, CrashLatCol)
        CrashLong = Sheets(CrashSheet).Cells(CrashRow, CrashLongCol)
        
        'debugging check
        If CrashRow = 99347 Then
            CrashRow = CrashRow
        End If
        
        'Interstate
        If CrashRoute = 15 Or CrashRoute = 70 Or CrashRoute = 80 Or CrashRoute = 84 Or CrashRoute = 85 Or CrashRoute = 215 Then
            'if segment is not an interstate
            If SegRoute <> 15 And SegRoute <> 70 And SegRoute <> 80 And SegRoute <> 84 And SegRoute <> 85 And SegRoute <> 215 Then
                NextSegment = True
                Exit Do
            End If
            
            'if this is the first time the Interstate if statement is called for a particular interstate route number
            If IntstMrkrBool = False Then
                InterstateMarker = CrashRow
                IntstMrkrBool = True
            End If
                        
            'Next Segment
            If (CrashRoute = Int(Left(Label, 4)) And CrashMP > EMP) Then
                NextSegment = True
                Exit Do
                        
            'Year does not match
            ElseIf CrashYear < minyr Or CrashYear > maxyr Then
                GoAhead = False
                If CrashRoute <> Int(Left(Sheets(CrashSheet).Cells(CrashRow + 1, crashlabelcol), 4)) Then
                    'if staying on the same interstate but switching directions
                    If Left(Label, 4) = Left(Sheets(datasheet).Cells(SegRow + NumYears, LabelCol), 4) And Right(Label, 1) <> Right(Sheets(datasheet).Cells(SegRow + NumYears, LabelCol), 1) Then
                        EndofIntstDir = True
                        NextSegment = True
                        Exit Do
                    'if completing an interstate
                    ElseIf Label <> Sheets(datasheet).Cells(SegRow + NumYears, LabelCol) Then
                        IntstMrkrBool = False
                        EndofIntstDir = False
                        NextSegment = True
                        CrashRow = CrashRow + 1
                        Exit Do
                    End If
                End If

            'Label, Direction, and MP all match
            ElseIf Sheets(CrashSheet).Cells(CrashRow, crashlabelcol) = Label And CrashMP >= BMP And CrashMP < EMP Then
                GoAhead = True
                'if the next crash's route does not exaclty match the current crash's route
                If CrashRoute <> Int(Left(Sheets(CrashSheet).Cells(CrashRow + 1, crashlabelcol), 4)) Then
                    'if staying on the same interstate but switching directions
                    If Left(Label, 4) = Left(Sheets(datasheet).Cells(SegRow + NumYears, LabelCol), 4) And Right(Label, 1) <> Right(Sheets(datasheet).Cells(SegRow + NumYears, LabelCol), 1) Then
                        EndofIntstDir = True
                        NextSegment = True
                        ExitDoAfter = True
                    'if completing an interstate
                    ElseIf Label <> Sheets(datasheet).Cells(SegRow + NumYears, LabelCol) Then
                        IntstMrkrBool = False
                        EndofIntstDir = False
                        NextSegment = True
                        ExitDoAfter = True
                    End If
                End If

            'Some other condition (ex. direction does not match)
            Else
                GoAhead = False
                If CrashRoute <> Int(Left(Sheets(CrashSheet).Cells(CrashRow + 1, crashlabelcol), 4)) Then
                    'if staying on the same interstate but switching directions
                    If Left(Label, 4) = Left(Sheets(datasheet).Cells(SegRow + NumYears, LabelCol), 4) And Right(Label, 1) <> Right(Sheets(datasheet).Cells(SegRow + NumYears, LabelCol), 1) Then
                        EndofIntstDir = True
                        NextSegment = True
                        Exit Do
                    'if completing an interstate
                    ElseIf Label <> Sheets(datasheet).Cells(SegRow + NumYears, LabelCol) Then
                        IntstMrkrBool = False
                        CrashRow = CrashRow + 1
                        NextSegment = True
                        EndofIntstDir = False
                        Exit Do
                    End If
                End If
            End If

        
        'Next Segment / Next Route
        ElseIf (CrashRoute = Int(Left(Label, 4)) And CrashMP > EMP) Or CrashRoute > Int(Left(Label, 4)) Then
            NextSegment = True
            Exit Do
                    
        'Year does not match
        ElseIf CrashYear < minyr Or CrashYear > maxyr Then
            GoAhead = False
                        
        'if the crash route and MP match with the segment
        ElseIf CrashRoute = Int(Left(Label, 4)) And CrashMP >= BMP And CrashMP <= EMP Then
            GoAhead = True
            
        'none of the other conditions were met
        Else
            GoAhead = False
            
        End If
        
   
        If GoAhead = True Then
            For i = 0 To NumYears - 1
                SegYear = Sheets(datasheet).Cells(SegRow + i, YearCol).Value
            
                If SegYear = CrashYear Then
                    Sheets(datasheet).Cells(SegRow + i, ColSumTotCrash) = Sheets(datasheet).Cells(SegRow + i, ColSumTotCrash) + 1
                    'Tally severity numbers for all the crashes on each segment
                    If CrashSev = 5 Then
                        Sheets(datasheet).Cells(SegRow + i, ColSev5) = Sheets(datasheet).Cells(SegRow + i, ColSev5) + 1
                    ElseIf CrashSev = 4 Then
                        Sheets(datasheet).Cells(SegRow + i, ColSev4) = Sheets(datasheet).Cells(SegRow + i, ColSev4) + 1
                    ElseIf CrashSev = 3 Then
                        Sheets(datasheet).Cells(SegRow + i, ColSev3) = Sheets(datasheet).Cells(SegRow + i, ColSev3) + 1
                    ElseIf CrashSev = 2 Then
                        Sheets(datasheet).Cells(SegRow + i, ColSev2) = Sheets(datasheet).Cells(SegRow + i, ColSev2) + 1
                    ElseIf CrashSev = 1 Then
                        Sheets(datasheet).Cells(SegRow + i, ColSev1) = Sheets(datasheet).Cells(SegRow + i, ColSev1) + 1
                    End If
                End If
            Next i
            
            'what does this do?? 'Camille: I think it is checking to make sure it doesn't start tallying numbers for severities that were not checked in the user form.
            '                     'then maybe the sevstring get erased?
            If CrashSev = Sevs(1) Or CrashSev = Sevs(2) Or CrashSev = Sevs(3) Or CrashSev = Sevs(4) Or CrashSev = Sevs(5) Then
            
                'Input crash data in parameters sheet
                For j = 0 To NumYears - 1
                    SegYear = Sheets(datasheet).Cells(SegRow + j, YearCol).Value

                    If SegYear = CrashYear Then
                        Sheets(PSheet).Cells(prow, 1) = SegID
                        For i = 1 To col2 + 1
                            Sheets(PSheet).Cells(prow, i + 1) = Sheets(CrashSheet).Cells(CrashRow, i)
                        Next i
                        prow = prow + 1
                    End If
                Next j
            
                For i = 0 To NumYears - 1
                    SegYear = Sheets(datasheet).Cells(SegRow + i, YearCol).Value
                    
                    If SegYear = CrashYear Then                                         'FLAGGED: This code assumes certain columns exist.
                        'Total crashes
                        Sheets(datasheet).Cells(SegRow + i, TotCrashCol) = Sheets(datasheet).Cells(SegRow + i, TotCrashCol) + 1
                        
                        'Severe crashes
                        If CrashSev >= 3 Then
                            Sheets(datasheet).Cells(SegRow + i, SevCrashCol) = Sheets(datasheet).Cells(SegRow + i, SevCrashCol) + 1
                        End If
                    
                        'Head on collisions
                        If Sheets(CrashSheet).Cells(CrashRow, MannerCollisionCol) = 3 Then
                             Sheets(datasheet).Cells(SegRow + i, HeadOnCol) = Sheets(datasheet).Cells(SegRow + i, HeadOnCol) + 1
                         End If
                                              
                        'Work zone related
                         'If Sheets(CrashSheet).Cells(CrashRow, WorkZoneCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(SegRow + i, WorkZoneCol1) = Sheets(datasheet).Cells(SegRow + i, WorkZoneCol1) + 1
                         'End If
                         
                        'Pedestrian Involved
                         If Sheets(CrashSheet).Cells(CrashRow, PedInvCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, PedInvCol1) = Sheets(datasheet).Cells(SegRow + i, PedInvCol1) + 1
                         End If
                         
                        'Bicyclist Involved
                         If Sheets(CrashSheet).Cells(CrashRow, BikeInvCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, BikeInvCol1) = Sheets(datasheet).Cells(SegRow + i, BikeInvCol1) + 1
                         End If
                         
                        'Motorcycle Involved
                         If Sheets(CrashSheet).Cells(CrashRow, MotInvCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, MotInvCol1) = Sheets(datasheet).Cells(SegRow + i, MotInvCol1) + 1
                         End If
                         
                        'Improper restraint
                         'If Sheets(CrashSheet).Cells(CrashRow, ImpResCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(SegRow + i, ImpResCol1) = Sheets(datasheet).Cells(SegRow + i, ImpResCol1) + 1
                         'End If
                         
                        'Unrestrained
                         If Sheets(CrashSheet).Cells(CrashRow, UnrestCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, UnrestCol1) = Sheets(datasheet).Cells(SegRow + i, UnrestCol1) + 1
                         End If
                         
                        'DUI
                         If Sheets(CrashSheet).Cells(CrashRow, DUICol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, DUICol1) = Sheets(datasheet).Cells(SegRow + i, DUICol1) + 1
                         End If
                         
                        'Aggresive Driving
                         If Sheets(CrashSheet).Cells(CrashRow, AggressiveCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, AggressiveCol1) = Sheets(datasheet).Cells(SegRow + i, AggressiveCol1) + 1
                         End If
                         
                        'Distracted Driving
                         If Sheets(CrashSheet).Cells(CrashRow, DistractedCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, DistractedCol1) = Sheets(datasheet).Cells(SegRow + i, DistractedCol1) + 1
                         End If
                         
                        'Drowsy Driving
                         If Sheets(CrashSheet).Cells(CrashRow, DrowsyCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, DrowsyCol1) = Sheets(datasheet).Cells(SegRow + i, DrowsyCol1) + 1
                         End If
                         
                        'Speed related
                         If Sheets(CrashSheet).Cells(CrashRow, SpeedCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, SpeedCol1) = Sheets(datasheet).Cells(SegRow + i, SpeedCol1) + 1
                         End If
                         
                        'Intersection related
                         If Sheets(CrashSheet).Cells(CrashRow, IntRelCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, IntRelCol1) = Sheets(datasheet).Cells(SegRow + i, IntRelCol1) + 1
                         End If
                         
                        'Adverse weather
                         If Sheets(CrashSheet).Cells(CrashRow, WeatherCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, WeatherCol1) = Sheets(datasheet).Cells(SegRow + i, WeatherCol1) + 1
                         End If
                         
                        'Adverse roadway surface conditions
                         If Sheets(CrashSheet).Cells(CrashRow, AdvRoadCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, AdvRoadCol1) = Sheets(datasheet).Cells(SegRow + i, AdvRoadCol1) + 1
                         End If
                         
                        'Roadway geometry related
                         If Sheets(CrashSheet).Cells(CrashRow, RoadGeomCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, RoadGeomCol1) = Sheets(datasheet).Cells(SegRow + i, RoadGeomCol1) + 1
                         End If
                         
                        'Wild animal related
                         If Sheets(CrashSheet).Cells(CrashRow, WAnimalCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, WAnimalCol1) = Sheets(datasheet).Cells(SegRow + i, WAnimalCol1) + 1
                         End If
                         
                        'Domestic animal related
                         If Sheets(CrashSheet).Cells(CrashRow, DAnimalCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, DAnimalCol1) = Sheets(datasheet).Cells(SegRow + i, DAnimalCol1) + 1
                         End If
                         
                        'Roadway departure
                         If Sheets(CrashSheet).Cells(CrashRow, RoadDepartCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, RoadDepartCol1) = Sheets(datasheet).Cells(SegRow + i, RoadDepartCol1) + 1
                         End If
                         
                        'Overturn/rollover
                         If Sheets(CrashSheet).Cells(CrashRow, OverturnCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, OverturnCol1) = Sheets(datasheet).Cells(SegRow + i, OverturnCol1) + 1
                         End If
                         
                        'Commercial motor vehicle involved
                         If Sheets(CrashSheet).Cells(CrashRow, CommVehCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, CommVehCol1) = Sheets(datasheet).Cells(SegRow + i, CommVehCol1) + 1
                         End If
                         
                        'Interstate highway
                         If Sheets(CrashSheet).Cells(CrashRow, InterstateCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, InterstateCol1) = Sheets(datasheet).Cells(SegRow + i, InterstateCol1) + 1
                         End If
                         
                        'Teenage driver involved
                         If Sheets(CrashSheet).Cells(CrashRow, TeenCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, TeenCol1) = Sheets(datasheet).Cells(SegRow + i, TeenCol1) + 1
                          End If
                         
                        'Older driver involved
                         If Sheets(CrashSheet).Cells(CrashRow, OldCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, OldCol1) = Sheets(datasheet).Cells(SegRow + i, OldCol1) + 1
                         End If
                         
                        'Urban county
                         'If Sheets(CrashSheet).Cells(CrashRow, UrbanCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(SegRow + i, UrbanCol1) = Sheets(datasheet).Cells(SegRow + i, UrbanCol1) + 1
                         'End If
                         
                        'Night dark condition
                         If Sheets(CrashSheet).Cells(CrashRow, NightCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, NightCol1) = Sheets(datasheet).Cells(SegRow + i, NightCol1) + 1
                         End If
                         
                        'Single vehicle
                         If Sheets(CrashSheet).Cells(CrashRow, SingleVehCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, SingleVehCol1) = Sheets(datasheet).Cells(SegRow + i, SingleVehCol1) + 1
                         End If
                         
                        'Train involved
                         If Sheets(CrashSheet).Cells(CrashRow, TrainCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, TrainCol1) = Sheets(datasheet).Cells(SegRow + i, TrainCol1) + 1
                         End If
                         
                        'Railroad crossing
                         If Sheets(CrashSheet).Cells(CrashRow, RailCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, RailCol1) = Sheets(datasheet).Cells(SegRow + i, RailCol1) + 1
                         End If
                         
                        'Transit vehicle involved
                         If Sheets(CrashSheet).Cells(CrashRow, TransitCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, TransitCol1) = Sheets(datasheet).Cells(SegRow + i, TransitCol1) + 1
                         End If
                         
                        'Collision with fixed object
                         If Sheets(CrashSheet).Cells(CrashRow, FixedObCol2) = "Y" Then
                             Sheets(datasheet).Cells(SegRow + i, FixedObCol1) = Sheets(datasheet).Cells(SegRow + i, FixedObCol1) + 1
                         End If
                         
                         'Work Zone with workers present
                         'If Sheets(CrashSheet).Cells(CrashRow, WZworkerCol2) = "Y" Then
                         '    Sheets(datasheet).Cells(SegRow + i, WZworkerCol1) = Sheets(datasheet).Cells(SegRow + i, WZworkerCol1) + 1
                         'End If
                    End If
                Next i
            End If
            If ExitDoAfter = True Then
                ExitDoAfter = False 'reset it since this is the only line where it gets used
                CrashRow = CrashRow + 1
                Exit Do
            End If
        End If
        CrashRow = CrashRow + 1
    Loop

    SegRow = SegRow + NumYears
Loop

End Sub

Sub SumCrashData()

'Declare variables
Dim IntRow, CrashRow, CrashRowPrev As Long
Dim IntLatCol, IntLongCol, CrashLatCol, CrashLongCol, UrbanCodeCol, UrbanCode As Long
Dim IntLat, IntLong, LatFtEquiv, LongFtEquiv, MinLatStart, MaxLatEnd As Double
Dim CrashLat, CrashLong, LatDiff, LongDiff, DiffRadius As Double
Dim IntRadius, SpeedLimit, NumYears As Integer
Dim datasheet, CrashSheet, UCDesc, FCDesc As String
Dim row1, col1, CHeadCol As Long

Dim MaxSLCol, MaxFCCol, UCcol As Long
Dim IntNumCol, YearCol, TotCrashCol, SevCrashCol, PedInvCol1, BikeInvCol1, MotInvCol1, ImpResCol1, UnrestCol1, DUICol1 As Long
Dim AggressiveCol1, DistractedCol1, DrowsyCol1, SpeedCol1, IntRelCol1, WeatherCol1, AdvRoadCol1, RoadGeomCol1 As Long
Dim WAnimalCol1, DAnimalCol1, RoadDepartCol1, OverturnCol1, CommVehCol1, InterstateCol1, TeenCol1, OldCol1 As Long
Dim UrbanCol1, NightCol1, SingleVehCol1, TrainCol1, RailCol1, TransitCol1, FixedObCol1, WorkZoneCol1, CollMannerCol1 As Long
Dim MaxRoadWCol, MinRoadWcol, MaxRoadW, MinRoadW As Long
Dim IntYear As String

Dim CrashDateCol, CrashSeverityCol, PedInvCol2, BikeInvCol2, MotInvCol2, ImpResCol2, UnrestCol2, DUICol2 As Long
Dim AggressiveCol2, DistractedCol2, DrowsyCol2, SpeedCol2, IntRelCol2, WeatherCol2, AdvRoadCol2, RoadGeomCol2 As Long
Dim WAnimalCol2, DAnimalCol2, RoadDepartCol2, OverturnCol2, CommVehCol2, InterstateCol2, TeenCol2, OldCol2 As Long
Dim UrbanCol2, NightCol2, SingleVehCol2, TrainCol2, RailCol2, TransitCol2, FixedObCol2, WorkZoneCol2, CollMannerCol2 As Long
Dim CrashYear As String

Dim UICPMcol, FArow, i As Long
Dim sevstring As String
ReDim Sevs(1 To 5) As Integer


'Assign initial values
IntRow = 2
CrashRow = 2

IntLatCol = 1
IntLongCol = 1
CrashLatCol = 1
CrashLongCol = 1

datasheet = "UICPMinput"
CrashSheet = "CrashInput"

IntNumCol = 1
YearCol = 1
MaxSLCol = 1
MaxFCCol = 1
UCcol = 1
TotCrashCol = 1
SevCrashCol = 1
UrbanCodeCol = 1
PedInvCol1 = 1
BikeInvCol1 = 1
MotInvCol1 = 1
ImpResCol1 = 1
UnrestCol1 = 1
DUICol1 = 1
AggressiveCol1 = 1
DistractedCol1 = 1
DrowsyCol1 = 1
SpeedCol1 = 1
IntRelCol1 = 1
WeatherCol1 = 1
AdvRoadCol1 = 1
RoadGeomCol1 = 1
WAnimalCol1 = 1
DAnimalCol1 = 1
RoadDepartCol1 = 1
OverturnCol1 = 1
CommVehCol1 = 1
InterstateCol1 = 1
TeenCol1 = 1
OldCol1 = 1
UrbanCol1 = 1
NightCol1 = 1
SingleVehCol1 = 1
TrainCol1 = 1
RailCol1 = 1
TransitCol1 = 1
FixedObCol1 = 1
WorkZoneCol1 = 1
CollMannerCol1 = 1
MaxRoadWCol = 1
MinRoadWcol = 1

CrashDateCol = 1
CrashSeverityCol = 1
PedInvCol2 = 1
BikeInvCol2 = 1
MotInvCol2 = 1
ImpResCol2 = 1
UnrestCol2 = 1
DUICol2 = 1
AggressiveCol2 = 1
DistractedCol2 = 1
DrowsyCol2 = 1
SpeedCol2 = 1
IntRelCol2 = 1
WeatherCol2 = 1
AdvRoadCol2 = 1
RoadGeomCol2 = 1
WAnimalCol2 = 1
DAnimalCol2 = 1
RoadDepartCol2 = 1
OverturnCol2 = 1
CommVehCol2 = 1
InterstateCol2 = 1
TeenCol2 = 1
OldCol2 = 1
UrbanCol2 = 1
NightCol2 = 1
SingleVehCol2 = 1
TrainCol2 = 1
RailCol2 = 1
TransitCol2 = 1
FixedObCol2 = 1
WorkZoneCol2 = 1
CollMannerCol2 = 1

UICPMcol = 1
FArow = 1


'Find first blank column in roadway intersection data and add in column headings
row1 = 9
col1 = 1
CHeadCol = 1
Do Until Sheets(datasheet).Cells(1, col1) = ""
    col1 = col1 + 1
Loop

Do Until Sheets("Key").Cells(1, CHeadCol) = "Intersection Check Headers"
    CHeadCol = CHeadCol + 1
Loop

Do Until Sheets("Key").Cells(2, CHeadCol) = 2
    CHeadCol = CHeadCol + 1
Loop

Do Until Sheets("Key").Cells(row1, CHeadCol + 2) = ""
    Sheets(datasheet).Cells(1, col1) = Sheets("Key").Cells(row1, CHeadCol + 2)
    col1 = col1 + 1
    row1 = row1 + 1
Loop

'Find columns
Do Until Sheets(datasheet).Cells(1, IntNumCol) = "INT_ID"
    IntNumCol = IntNumCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, YearCol) = "YEAR"
    YearCol = YearCol + 1
Loop

MaxSLCol = 1
Do Until Sheets(datasheet).Cells(1, MaxSLCol) = "MAX_SPEED_LIMIT"
    MaxSLCol = MaxSLCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, MaxFCCol) = "MAX FC_TYPE"
    MaxFCCol = MaxFCCol + 1
Loop

UCcol = 1
Do Until Sheets(datasheet).Cells(1, UCcol) = "URBAN_DESC"
    UCcol = UCcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, IntLatCol) = "LATITUDE"
    IntLatCol = IntLatCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, IntLongCol) = "LONGITUDE"
    IntLongCol = IntLongCol + 1
Loop

CrashLatCol = 1
Do Until Sheets(CrashSheet).Cells(1, CrashLatCol) = "LATITUDE"
    CrashLatCol = CrashLatCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CrashLongCol) = "LONGITUDE"
    CrashLongCol = CrashLongCol + 1
Loop

MaxRoadWCol = 1
Do Until Sheets(datasheet).Cells(1, MaxRoadWCol) = "MAX_ROAD_WIDTH"
    MaxRoadWCol = MaxRoadWCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, MinRoadWcol) = "MIN_ROAD_WIDTH"
    MinRoadWcol = MinRoadWcol + 1
Loop


Do Until Sheets(datasheet).Cells(1, TotCrashCol) = "Total_Crashes"
    TotCrashCol = TotCrashCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SevCrashCol) = "Severe_Crashes"
    SevCrashCol = SevCrashCol + 1
Loop

UrbanCodeCol = 1
Do Until Sheets(datasheet).Cells(1, UrbanCodeCol) = "URBAN_CODE"
    UrbanCodeCol = UrbanCodeCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, WorkZoneCol1) = "WORKZONE_RELATED"
    WorkZoneCol1 = WorkZoneCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, CollMannerCol1) = "HEADON_COLLISION"
    CollMannerCol1 = CollMannerCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, PedInvCol1) = "PEDESTRIAN_INVOLVED"
    PedInvCol1 = PedInvCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, BikeInvCol1) = "BICYCLIST_INVOLVED"
    BikeInvCol1 = BikeInvCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, MotInvCol1) = "MOTORCYCLE_INVOLVED"
    MotInvCol1 = MotInvCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ImpResCol1) = "IMPROPER_RESTRAINT"
    ImpResCol1 = ImpResCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, UnrestCol1) = "UNRESTRAINED"
    UnrestCol1 = UnrestCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DUICol1) = "DUI"
    DUICol1 = DUICol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, AggressiveCol1) = "AGGRESSIVE_DRIVING"
    AggressiveCol1 = AggressiveCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DistractedCol1) = "DISTRACTED_DRIVING"
    DistractedCol1 = DistractedCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DrowsyCol1) = "DROWSY_DRIVING"
    DrowsyCol1 = DrowsyCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, SpeedCol1) = "SPEED_RELATED"
    SpeedCol1 = SpeedCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, IntRelCol1) = "INTERSECTION_RELATED"
    IntRelCol1 = IntRelCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, WeatherCol1) = "ADVERSE_WEATHER"
    WeatherCol1 = WeatherCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, AdvRoadCol1) = "ADVERSE_ROADWAY_SURF_CONDITION"
    AdvRoadCol1 = AdvRoadCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, RoadGeomCol1) = "ROADWAY_GEOMETRY_RELATED"
    RoadGeomCol1 = RoadGeomCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, WAnimalCol1) = "WILD_ANIMAL_RELATED"
    WAnimalCol1 = WAnimalCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DAnimalCol1) = "DOMESTIC_ANIMAL_RELATED"
    DAnimalCol1 = DAnimalCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, RoadDepartCol1) = "ROADWAY_DEPARTURE"
    RoadDepartCol1 = RoadDepartCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, OverturnCol1) = "OVERTURN_ROLLOVER"
    OverturnCol1 = OverturnCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, CommVehCol1) = "COMMERCIAL_MOTOR_VEH_INVOLVED"
    CommVehCol1 = CommVehCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, InterstateCol1) = "INTERSTATE_HIGHWAY"
    InterstateCol1 = InterstateCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, TeenCol1) = "TEENAGE_DRIVER_INVOLVED"
    TeenCol1 = TeenCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, OldCol1) = "OLDER_DRIVER_INVOLVED"
    OldCol1 = OldCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, UrbanCol1) = "URBAN_COUNTY"
    UrbanCol1 = UrbanCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, NightCol1) = "NIGHT_DARK_CONDITION"
    NightCol1 = NightCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, SingleVehCol1) = "SINGLE_VEHICLE"
    SingleVehCol1 = SingleVehCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, TrainCol1) = "TRAIN_INVOLVED"
    TrainCol1 = TrainCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, RailCol1) = "RAILROAD_CROSSING"
    RailCol1 = RailCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, TransitCol1) = "TRANSIT_VEHICLE_INVOLVED"
    TransitCol1 = TransitCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, FixedObCol1) = "COLLISION_WITH_FIXED_OBJECT"
    FixedObCol1 = FixedObCol1 + 1
Loop


Do Until Sheets(CrashSheet).Cells(1, CrashDateCol) = "CRASH_DATETIME"
    CrashDateCol = CrashDateCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CrashSeverityCol) = "CRASH_SEVERITY_ID"
    CrashSeverityCol = CrashSeverityCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, WorkZoneCol2) = "WORK_ZONE_RELATED_YNU"
    WorkZoneCol2 = WorkZoneCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CollMannerCol2) = "MANNER_COLLISION_ID"
    CollMannerCol2 = CollMannerCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, PedInvCol2) = "PEDESTRIAN_INVOLVED"
    PedInvCol2 = PedInvCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, BikeInvCol2) = "BICYCLIST_INVOLVED"
    BikeInvCol2 = BikeInvCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, MotInvCol2) = "MOTORCYCLE_INVOLVED"
    MotInvCol2 = MotInvCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, ImpResCol2) = "IMPROPER_RESTRAINT"
    ImpResCol2 = ImpResCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, UnrestCol2) = "UNRESTRAINED"
    UnrestCol2 = UnrestCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DUICol2) = "DUI"
    DUICol2 = DUICol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, AggressiveCol2) = "AGGRESSIVE_DRIVING"
    AggressiveCol2 = AggressiveCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DistractedCol2) = "DISTRACTED_DRIVING"
    DistractedCol2 = DistractedCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DrowsyCol2) = "DROWSY_DRIVING"
    DrowsyCol2 = DrowsyCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, SpeedCol2) = "SPEED_RELATED"
    SpeedCol2 = SpeedCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, IntRelCol2) = "INTERSECTION_RELATED"
    IntRelCol2 = IntRelCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, WeatherCol2) = "ADVERSE_WEATHER"
    WeatherCol2 = WeatherCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, AdvRoadCol2) = "ADVERSE_ROADWAY_SURF_CONDITION"
    AdvRoadCol2 = AdvRoadCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, RoadGeomCol2) = "ROADWAY_GEOMETRY_RELATED"
    RoadGeomCol2 = RoadGeomCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, WAnimalCol2) = "WILD_ANIMAL_RELATED"
    WAnimalCol2 = WAnimalCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DAnimalCol2) = "DOMESTIC_ANIMAL_RELATED"
    DAnimalCol2 = DAnimalCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, RoadDepartCol2) = "ROADWAY_DEPARTURE"
    RoadDepartCol2 = RoadDepartCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, OverturnCol2) = "OVERTURN_ROLLOVER"
    OverturnCol2 = OverturnCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CommVehCol2) = "COMMERCIAL_MOTOR_VEH_INVOLVED"
    CommVehCol2 = CommVehCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, InterstateCol2) = "INTERSTATE_HIGHWAY"
    InterstateCol2 = InterstateCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, TeenCol2) = "TEENAGE_DRIVER_INVOLVED"
    TeenCol2 = TeenCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, OldCol2) = "OLDER_DRIVER_INVOLVED"
    OldCol2 = OldCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, UrbanCol2) = "URBAN_COUNTY"
    UrbanCol2 = UrbanCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, NightCol2) = "NIGHT_DARK_CONDITION"
    NightCol2 = NightCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, SingleVehCol2) = "SINGLE_VEHICLE"
    SingleVehCol2 = SingleVehCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, TrainCol2) = "TRAIN_INVOLVED"
    TrainCol2 = TrainCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, RailCol2) = "RAILROAD_CROSSING"
    RailCol2 = RailCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, TransitCol2) = "TRANSIT_VEHICLE_INVOLVED"
    TransitCol2 = TransitCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, FixedObCol2) = "COLLISION_WITH_FIXED_OBJECT"
    FixedObCol2 = FixedObCol2 + 1
Loop


Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "UICPM"
    UICPMcol = UICPMcol + 1
Loop

Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "Selected FA Parameter"
    FArow = FArow + 1
Loop


LatLongElevSortUICPM

'Find number of years of data
NumYears = 1
Do Until Cells(IntRow + NumYears, YearCol) = Cells(IntRow, YearCol)
    NumYears = NumYears + 1
Loop


'Loop through intersections and assign crashes that fit within the given radius
CrashRowPrev = 2
IntRow = 2
Do While Sheets(datasheet).Cells(IntRow, 1) <> ""
    IntLat = Sheets(datasheet).Cells(IntRow, IntLatCol)
    IntLong = Sheets(datasheet).Cells(IntRow, IntLongCol)
    MaxRoadW = Sheets(datasheet).Cells(IntRow, MaxRoadWCol)
    MinRoadW = Sheets(datasheet).Cells(IntRow, MinRoadWcol)
    
    'Calculate the distance in feet for each lat/long degree for the intersection
    LatFtEquiv = 69.172 * 5280
    LongFtEquiv = Cos(IntLat) * 69.172 * 5280
    
    'Determine functional area value
    If Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Speed Limit" Then
        SpeedLimit = Sheets(datasheet).Cells(IntRow, MaxSLCol)
        For i = 1 To 12
            If SpeedLimit = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                IntRadius = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                Exit For
            End If
        Next i
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Functional Class" Then
        FCDesc = Sheets(datasheet).Cells(IntRow, MaxFCCol)
        For i = 1 To 4
            If FCDesc = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                IntRadius = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                Exit For
            End If
        Next i
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Urban Code" Then
        UCDesc = Sheets(datasheet).Cells(IntRow, UCcol)
        
        If UCDesc = "Rurall" Then
            IntRadius = Sheets("Inputs").Cells(FArow + 2, UICPMcol + 1)
        ElseIf UCDesc = "Small Urban" Then
            IntRadius = Sheets("Inputs").Cells(FArow + 3, UICPMcol + 1)
        Else
            IntRadius = Sheets("Inputs").Cells(FArow + 4, UICPMcol + 1)
        End If
    End If
    
    'Add intersection size to intersection radius
    IntRadius = IntRadius + Sqr(MaxRoadW * MinRoadW)
    
    'Determine severities
    sevstring = Sheets("Inputs").Cells(FArow - 1, UICPMcol + 1)
    For i = 1 To Len(sevstring)
        Sevs(i) = Mid(sevstring, i, 1)
    Next i
    
    MinLatStart = IntLat - (IntRadius / LatFtEquiv)
    MaxLatEnd = IntLat + (IntRadius / LatFtEquiv)
    
    'Go to crash row that is close to range of intersection to save time
    CrashRow = CrashRowPrev
    Do Until Sheets(CrashSheet).Cells(CrashRow, CrashLatCol) >= MinLatStart
        CrashRow = CrashRow + 1
    Loop
    
    CrashRowPrev = CrashRow
    
    'Assign crashes to intersection
    Do Until Sheets(CrashSheet).Cells(CrashRow, CrashLatCol) > MaxLatEnd
        CrashLat = Sheets(CrashSheet).Cells(CrashRow, CrashLatCol)
        CrashLong = Sheets(CrashSheet).Cells(CrashRow, CrashLongCol)
        LatDiff = Abs(CrashLat - IntLat) * LatFtEquiv
        LongDiff = Abs(CrashLong - IntLong) * LongFtEquiv
        DiffRadius = Sqr(LatDiff ^ 2 + LongDiff ^ 2)
        CrashYear = year(Sheets(CrashSheet).Cells(CrashRow, CrashDateCol))
        
        If DiffRadius < IntRadius And (Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(1) Or _
        Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(2) Or Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(3) Or _
        Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(4) Or Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(5)) Then
        
            'Total crahes
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = "" Or Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = Sheets(datasheet).Cells(IntRow + i, TotCrashCol) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = "" Or Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = 0
                    End If
                End If
            Next i
                        
            'Severe crashes
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) >= 3 And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = "" Or Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = Sheets(datasheet).Cells(IntRow + i, SevCrashCol) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = "" Or Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = 0
                    End If
                End If
            Next i
                
            'Head on collisions
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
           
                If Sheets(CrashSheet).Cells(CrashRow, CollMannerCol2) = 3 And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = 0
                    End If
                End If
            Next i
                
                
            'Work zone related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, WorkZoneCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = 0
                    End If
                End If
            Next i
                
            'Pedestrian Involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, PedInvCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = Sheets(datasheet).Cells(IntRow + i, PedInvCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = 0
                    End If
                End If
            Next i
                
            'Bicyclist Involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, BikeInvCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = 0
                    End If
                End If
            Next i
                
            'Motorcycle Involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, MotInvCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = Sheets(datasheet).Cells(IntRow + i, MotInvCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = 0
                    End If
                End If
            Next i
                
            'Improper restraint
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, ImpResCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = Sheets(datasheet).Cells(IntRow + i, ImpResCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = 0
                    End If
                End If
            Next i
                
            'Unrestrained
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, UnrestCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = Sheets(datasheet).Cells(IntRow + i, UnrestCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = 0
                    End If
                End If
            Next i
                
            'DUI
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DUICol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, DUICol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DUICol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, DUICol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, DUICol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, DUICol1) = Sheets(datasheet).Cells(IntRow + i, DUICol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, DUICol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DUICol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, DUICol1) = 0
                    End If
                End If
            Next i
                
            'Aggresive Driving
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, AggressiveCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = 0
                    End If
                End If
            Next i
                
            'Distracted Driving
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DistractedCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = Sheets(datasheet).Cells(IntRow + i, DistractedCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = 0
                    End If
                End If
            Next i
                
            'Drowsy Driving
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DrowsyCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = 0
                    End If
                End If
            Next i
                
            'Speed related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, SpeedCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = Sheets(datasheet).Cells(IntRow + i, SpeedCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = 0
                    End If
                End If
            Next i
                
            'Intersection related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, IntRelCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = Sheets(datasheet).Cells(IntRow + i, IntRelCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = 0
                    End If
                End If
            Next i
                
            'Adverse weather
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, WeatherCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = Sheets(datasheet).Cells(IntRow + i, WeatherCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = 0
                    End If
                End If
            Next i
                
            'Adverse roadway surface conditions
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, AdvRoadCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = 0
                    End If
                End If
            Next i
                
            'Roadway geometry related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, RoadGeomCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = 0
                    End If
                End If
            Next i
                
            'Wild animal related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, WAnimalCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = 0
                    End If
                End If
            Next i
                
            'Domestic animal related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DAnimalCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = 0
                    End If
                End If
            Next i
                
            'Roadway departure
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, RoadDepartCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = 0
                    End If
                End If
            Next i
                
            'Overturn/rollover
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, OverturnCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = Sheets(datasheet).Cells(IntRow + i, OverturnCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = 0
                    End If
                End If
            Next i
                
            'Commercial motor vehicle involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, CommVehCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = Sheets(datasheet).Cells(IntRow + i, CommVehCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = 0
                    End If
                End If
            Next i
                
            'Interstate highway
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, InterstateCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = Sheets(datasheet).Cells(IntRow + i, InterstateCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = 0
                    End If
                End If
            Next i
                
            'Teenage driver involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, TeenCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, TeenCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, TeenCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, TeenCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, TeenCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, TeenCol1) = Sheets(datasheet).Cells(IntRow + i, TeenCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, TeenCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, TeenCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, TeenCol1) = 0
                    End If
                End If
            Next i
                
            'Older driver involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, OldCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, OldCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, OldCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, OldCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, OldCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, OldCol1) = Sheets(datasheet).Cells(IntRow + i, OldCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, OldCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, OldCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, OldCol1) = 0
                    End If
                End If
            Next i
                
            'Urban county
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, UrbanCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = Sheets(datasheet).Cells(IntRow + i, UrbanCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = 0
                    End If
                End If
            Next i
                
            'Night dark condition
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, NightCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, NightCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, NightCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, NightCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, NightCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, NightCol1) = Sheets(datasheet).Cells(IntRow + i, NightCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, NightCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, NightCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, NightCol1) = 0
                    End If
                End If
            Next i
                
            'Single vehicle
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, SingleVehCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = 0
                    End If
                End If
            Next i
                
            'Train involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, TrainCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, TrainCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, TrainCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, TrainCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, TrainCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, TrainCol1) = Sheets(datasheet).Cells(IntRow + i, TrainCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, TrainCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, TrainCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, TrainCol1) = 0
                    End If
                End If
            Next i
                
            'Railroad crossing
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, RailCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, RailCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, RailCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, RailCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, RailCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, RailCol1) = Sheets(datasheet).Cells(IntRow + i, RailCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, RailCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, RailCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, RailCol1) = 0
                    End If
                End If
            Next i
                
            'Transit vehicle involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, TransitCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, TransitCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, TransitCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, TransitCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, TransitCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, TransitCol1) = Sheets(datasheet).Cells(IntRow + i, TransitCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, TransitCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, TransitCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, TransitCol1) = 0
                    End If
                End If
            Next i
                
            'Collision with fixed object
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, FixedObCol2) = "Y" And IntYear = CrashYear Then
                    If Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = " " Or _
                    Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = 0 Then
                        Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = 1
                    Else
                        Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = Sheets(datasheet).Cells(IntRow + i, FixedObCol1) + 1
                    End If
                Else
                    If Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = "" Or Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = " " Then
                        Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = 0
                    End If
                End If
            Next i

        ElseIf Sheets(datasheet).Cells(IntRow, TotCrashCol) = "" Or Sheets(datasheet).Cells(IntRow, TotCrashCol) = " " Then
            For i = 0 To NumYears
                Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = 0
                Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = 0
                Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, DUICol1) = 0
                Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, TeenCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, OldCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, NightCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, TrainCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, RailCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, TransitCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = 0
                Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = 0
            Next i
        End If
        CrashRow = CrashRow + 1
    Loop
    
    IntRow = IntRow + NumYears
Loop

End Sub


Sub executeBA(rscript As String, rcode As String, bawd As String, niter As Long, nburn As Long, datalocation As String)
' The purpose of this script is to send the BeforeAfter parameters to the command line, to start the BeforeAfter analysis
' rscrip = "Rscript.exe" file path
' rcode = file path of the statistical script R file
' bawd = working directory, to save output of statistical analysis
' niter = number of iterations for the statistical analysis
' nburn = number of burn-in iterations for the statistical analysis
' datalocation = file path to the input data file for the statistical analysis

' Define Variables
Dim cmdLine As String

' Prepare command line for Rscript execution
cmdLine = rscript & " " & rcode & " " & bawd & " " & niter & " " & nburn & " " & datalocation
'MsgBox cmdLine, vbOKOnly   'Debugging purposes
    
' Submit Rscript code execution
Shell cmdLine, vbMaximizedFocus

MsgBox "Before After Analysis started." & Chr(10) & "Do not shut down computer until analysis is finished.", vbOKOnly, "Before After Analysis Started"

End Sub

Sub executeUCPSM(rscript As String, rcode As String, wd As String, niter As Long, nburn As Long, datalocation As String, xs As String)
' The purpose of this script is to send the statistical parameters to the command line and execute the Rscript program
' rscrip        = "Rscript.exe" file path
' rcode         = file path of the statistical script R file
' wd            = working directory, to save output of statistical analysis
' niter         = number of iterations for the statistical analysis
' nburn         = number of burn-in iterations for the statistical analysis
' datalocation  = file path to the input data file for the statistical analysis
' xs            = column index of the significant parameters for the analysis. A "()" value will trigger the horseshoe selection method.

' Define Variables
Dim cmdLine As String

' Prepare command line for Rscript execution
cmdLine = rscript & " " & rcode & " " & wd & " " & niter & " " & nburn & " " & datalocation & " " & xs
MsgBox cmdLine, vbOKOnly   'Debugging purposes

' Submit Rscript code execution
Shell cmdLine, vbMaximizedFocus

MsgBox "Statistical analysis started." & Chr(10) & "Do not shut down computer until analysis is finished.", vbOKOnly, "Statistical Analysis Started"

End Sub

Sub executeUICPM(rscript As String, rcode As String, wd As String, niter As Long, nburn As Long, datalocation As String, xs As String)
' The purpose of this script is to send the statistical parameters to the command line and execute the Rscript program
' rscrip        = "Rscript.exe" file path
' rcode         = file path of the statistical script R file
' wd            = working directory, to save output of statistical analysis
' niter         = number of iterations for the statistical analysis
' nburn         = number of burn-in iterations for the statistical analysis
' datalocation  = file path to the input data file for the statistical analysis
' xs            = column index of the significant parameters for the analysis. A "()" value will trigger the horseshoe selection method.

' Define Variables
Dim cmdLine As String

' Prepare command line for Rscript execution
cmdLine = rscript & " " & rcode & " " & wd & " " & niter & " " & nburn & " " & datalocation & " " & xs
MsgBox cmdLine, vbOKOnly   'Debugging purposes

' Submit Rscript code execution
Shell cmdLine, vbMaximizedFocus

MsgBox "Statistical analysis started." & Chr(10) & "Do not shut down computer until analysis is finished.", vbOKOnly, "Statistical Analysis Started"

End Sub

Sub QUICKexecuteUICPM()
' The purpose of this script is to send the statistical parameters to the command line and execute the Rscript program
' rscrip        = "Rscript.exe" file path
' rcode         = file path of the statistical script R file
' wd            = working directory, to save output of statistical analysis
' niter         = number of iterations for the statistical analysis
' nburn         = number of burn-in iterations for the statistical analysis
' datalocation  = file path to the input data file for the statistical analysis
' xs            = column index of the significant parameters for the analysis. A "()" value will trigger the horseshoe selection method.

' Define Variables
Dim cmdLine As String
Dim rscript As String, rcode As String, wd As String, niter As Long, nburn As Long, datalocation As String, xs As String

' Prepare command line for Rscript execution
rscript = "C:/Program Files/R/R-3.3.3/bin/Rscript"
rcode = "J:/groups/udotfy17/03_IntersectionModelDevelopment/Preliminary_Model/UICPM_1.2.R"
wd = "J:/groups/udotfy17/03_IntersectionModelDevelopment/Preliminary_Model/Test"
niter = 10000
nburn = 1000
datalocation = "J:/groups/udotfy17/03_IntersectionModelDevelopment/Preliminary_Model/UICPMinput_Sev12345_SL_4-12-2017_3-13-39_PM.csv"
xs = "(25,29)"


cmdLine = rscript & " " & rcode & " " & wd & " " & niter & " " & nburn & " " & datalocation & " " & xs
MsgBox cmdLine, vbOKOnly   'Debugging purposes

' Submit Rscript code execution
Shell cmdLine, vbMaximizedFocus

MsgBox "Statistical analysis started." & Chr(10) & "Do not shut down computer until analysis is finished.", vbOKOnly, "Statistical Analysis Started"

End Sub

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
' This function checks if a sheet exists in the workbook

    Dim sht As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
    
End Function

Function AlreadyOpen(sFname As String) As Boolean
' This function checks a given file is already open

    Dim wkb As Workbook
    On Error Resume Next
    Set wkb = Workbooks(sFname)
    AlreadyOpen = Not wkb Is Nothing
    Set wkb = Nothing
    
End Function

Function FolderExists(FolderPath As String)
' This function checks if a given folder path exists

    Dim TestStr As String
    TestStr = ""
    
    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If

    On Error Resume Next
    TestStr = Dir(FolderPath)
    On Error GoTo 0
    If TestStr = "" Then
        FolderExists = False
    Else
        FolderExists = True
    End If

End Function

Function FileExists(FilePath As String)
' This function checks if a given file path exists

    Dim TestStr As String
    TestStr = ""

    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function

Function GetFolder(strpath As String) As String
' This function gets a folder, returns a file path

Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strpath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function

Sub MaxMinCols()

'Define variables
Dim datasheet As String
Dim row1, row2 As Long
Dim SLCol0, SLCol1, SLCol2, SLCol3, SLCol4, FCCodeCol0, FCCodeCol1, FCCodeCol2, FCCodeCol3, FCCodeCol4 As Integer
Dim FCTypeCol0, FCTypeCol1, FCTypeCol2, FCTypeCol3, FCTypeCol4 As Integer
Dim SL0, SL1, SL2, SL3, SL4
Dim FC0, FC1, FC2, FC3, FC4
Dim FCT0, FCT1, FCT2, FCT3, FCT4 As String
'Dim SLBlank, FCBlank As Boolean
Dim MaxSLCol, MinSLCol As Integer
Dim MaxFCCodeCol, MinFCCodeCol, MaxFCTypeCol, MinFCTypeCol As Integer


'Assign sheet name
datasheet = "UICPMinput"

'Speed Limit
'Find Columns
SLCol0 = FindColumn("SPEED_LIMIT_0", datasheet, 1, True)
SLCol1 = FindColumn("SPEED_LIMIT_1", datasheet, 1, True)
SLCol2 = FindColumn("SPEED_LIMIT_2", datasheet, 1, True)
SLCol3 = FindColumn("SPEED_LIMIT_3", datasheet, 1, True)
SLCol4 = FindColumn("SPEED_LIMIT_4", datasheet, 1, True)
'SLBlank = True
'FCBlank = True

''Check to see if third column is blank. Delete if blank.
'row1 = 2
'Do Until Worksheets(datasheet).Cells(row1, SLCol1) = ""
'    row1 = row1 + 1
'Loop

'row2 = 2
'Do Until row2 = row1
'    If Worksheets(datasheet).Cells(row2, SLCol3) <> "" Then
'        SLBlank = False
'        FCBlank = False
'    End If
'    row2 = row2 + 1
'Loop

'If SLBlank = True Then
'    Worksheets(datasheet).Columns(SLCol3).EntireColumn.Delete
'End If

''Change column names
'Worksheets(datasheet).Cells(1, SLCol1) = "MAX_SPEED_LIMIT"
'Worksheets(datasheet).Cells(1, SLCol2) = "MIN_SPEED_LIMIT"

''Edit max and min values
'row1 = 2
'If SLBlank = True Then
'    Do Until Worksheets(datasheet).Cells(row1, SLCol1) = ""
'        SL1 = Worksheets(datasheet).Cells(row1, SLCol1)
'        SL2 = Worksheets(datasheet).Cells(row1, SLCol2)
'
'        If SL2 > SL1 Then
'            Worksheets(datasheet).Cells(row1, SLCol1) = SL2
'            Worksheets(datasheet).Cells(row1, SLCol2) = SL1
'        End If
'        row1 = row1 + 1
'    Loop
'Else
'    Do Until Worksheets(datasheet).Cells(row1, SLCol1) = ""
'        SL1 = Worksheets(datasheet).Cells(row1, SLCol1)
'        SL2 = Worksheets(datasheet).Cells(row1, SLCol2)
'        SL3 = Worksheets(datasheet).Cells(row1, SLCol3)
'
'        If SL3 > SL2 Then
'            Worksheets(datasheet).Cells(row1, SLCol2) = SL3
'            Worksheets(datasheet).Cells(row1, SLCol3) = SL2
'        End If
 '
'        If SL2 > SL1 Then
'            Worksheets(datasheet).Cells(row1, SLCol1) = SL2
'            Worksheets(datasheet).Cells(row1, SLCol2) = SL1
'        End If
'
'        If SL3 > SL2 Then
'            Worksheets(datasheet).Cells(row1, SLCol2) = SL3
'            Worksheets(datasheet).Cells(row1, SLCol3) = SL2
'        End If
'        row1 = row1 + 1
'    Loop
'    Worksheets(datasheet).Columns(SLCol3).EntireColumn.Delete
'End If

    
'rename column headings
Worksheets(datasheet).Cells(1, SLCol0) = "MAX_SPEED_LIMIT"
Worksheets(datasheet).Cells(1, SLCol1) = "MIN_SPEED_LIMIT"
    
'set new variable names (for ease of understanding the code)
MaxSLCol = SLCol0
MinSLCol = SLCol1

'Speed limit Max/Min finder and organizer
row1 = 2
Do Until Worksheets(datasheet).Cells(row1, 1) = ""
    'get values
    SL0 = Worksheets(datasheet).Cells(row1, SLCol0)
    SL1 = Worksheets(datasheet).Cells(row1, SLCol1)
    SL2 = Worksheets(datasheet).Cells(row1, SLCol2)
    SL3 = Worksheets(datasheet).Cells(row1, SLCol3)
    SL4 = Worksheets(datasheet).Cells(row1, SLCol4)
        
    'set max SL value
    If SL0 >= SL1 And SL0 >= SL2 And SL0 >= SL3 And SL0 >= SL4 And SL0 <> 0 And SL0 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxSLCol) = SL0
    ElseIf SL1 >= SL0 And SL1 >= SL2 And SL1 >= SL3 And SL1 >= SL4 And SL1 <> 0 And SL1 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxSLCol) = SL1
    ElseIf SL2 >= SL0 And SL2 >= SL1 And SL2 >= SL3 And SL2 >= SL4 And SL2 <> 0 And SL2 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxSLCol) = SL2
    ElseIf SL3 >= SL0 And SL3 >= SL1 And SL3 >= SL2 And SL3 >= SL4 And SL3 <> 0 And SL3 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxSLCol) = SL3
    ElseIf SL4 >= SL0 And SL4 >= SL1 And SL4 >= SL2 And SL4 >= SL3 And SL4 <> 0 And SL4 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxSLCol) = SL4
    End If
        
    'set min SL value
    If SL0 <= SL1 And SL0 <= SL2 And SL0 <= SL3 And SL0 <= SL4 And SL0 <> 0 And SL0 <> "" Then
        Worksheets(datasheet).Cells(row1, MinSLCol) = SL0
    ElseIf SL1 <= SL0 And SL1 <= SL2 And SL1 <= SL3 And SL1 <= SL4 And SL1 <> 0 And SL1 <> "" Then
        Worksheets(datasheet).Cells(row1, MinSLCol) = SL1
    ElseIf SL2 <= SL0 And SL2 <= SL1 And SL2 <= SL3 And SL2 <= SL4 And SL2 <> 0 And SL2 <> "" Then
        Worksheets(datasheet).Cells(row1, MinSLCol) = SL2
    ElseIf SL3 <= SL0 And SL3 <= SL1 And SL3 <= SL2 And SL3 <= SL4 And SL3 <> 0 And SL3 <> "" Then
        Worksheets(datasheet).Cells(row1, MinSLCol) = SL3
    ElseIf SL4 <= SL0 And SL4 <= SL1 And SL4 <= SL2 And SL4 <= SL3 And SL4 <> 0 And SL4 <> "" Then
        Worksheets(datasheet).Cells(row1, MinSLCol) = SL4
    End If
 
    row1 = row1 + 1
Loop

'delete extra columns
Worksheets(datasheet).Columns(SLCol4).EntireColumn.Delete
Worksheets(datasheet).Columns(SLCol3).EntireColumn.Delete
Worksheets(datasheet).Columns(SLCol2).EntireColumn.Delete




'Functional Class Max/Min finder and organizer
'Find Columns
FCCodeCol0 = FindColumn("FCCode_0", datasheet, 1, True)
FCCodeCol1 = FindColumn("FCCode_1", datasheet, 1, True)
FCCodeCol2 = FindColumn("FCCode_2", datasheet, 1, True)
FCCodeCol3 = FindColumn("FCCode_3", datasheet, 1, True)
FCCodeCol4 = FindColumn("FCCode_4", datasheet, 1, True)
FCTypeCol0 = FindColumn("FCType_0", datasheet, 1, True)
FCTypeCol1 = FindColumn("FCType_1", datasheet, 1, True)
FCTypeCol2 = FindColumn("FCType_2", datasheet, 1, True)
FCTypeCol3 = FindColumn("FCType_3", datasheet, 1, True)
FCTypeCol4 = FindColumn("FCType_4", datasheet, 1, True)

'If FCBlank = True Then
'    Worksheets(datasheet).Columns(FCTypeCol3).EntireColumn.Delete
'    Worksheets(datasheet).Columns(FCCodeCol3).EntireColumn.Delete
'    FCTypeCol1 = FCTypeCol1 - 1
'    FCTypeCol2 = FCTypeCol2 - 1
'End If

'rename column headings
Worksheets(datasheet).Cells(1, FCCodeCol0) = "MAX FC_CODE"
Worksheets(datasheet).Cells(1, FCCodeCol1) = "MIN FC_CODE"
Worksheets(datasheet).Cells(1, FCTypeCol0) = "MAX FC_TYPE"
Worksheets(datasheet).Cells(1, FCTypeCol1) = "MIN FC_TYPE"

'set new variable names (for ease of understanding the code)
MaxFCCodeCol = FCCodeCol0
MinFCCodeCol = FCCodeCol1
MaxFCTypeCol = FCTypeCol0
MinFCTypeCol = FCTypeCol1

row1 = 2
Do Until Worksheets(datasheet).Cells(row1, 1) = ""
    'get values
    FC0 = Worksheets(datasheet).Cells(row1, FCCodeCol0)
    FC1 = Worksheets(datasheet).Cells(row1, FCCodeCol1)
    FC2 = Worksheets(datasheet).Cells(row1, FCCodeCol2)
    FC3 = Worksheets(datasheet).Cells(row1, FCCodeCol3)
    FC4 = Worksheets(datasheet).Cells(row1, FCCodeCol4)
    FCT0 = Worksheets(datasheet).Cells(row1, FCTypeCol0)
    FCT1 = Worksheets(datasheet).Cells(row1, FCTypeCol1)
    FCT2 = Worksheets(datasheet).Cells(row1, FCTypeCol2)
    FCT3 = Worksheets(datasheet).Cells(row1, FCTypeCol3)
    FCT4 = Worksheets(datasheet).Cells(row1, FCTypeCol4)
        
    'set max FC code and type values  (FCcode 1 = interstate, FC code 7 = local)
    If FC0 <= FC1 And FC0 <= FC2 And FC0 <= FC3 And FC0 <= FC4 And FC0 <> 0 And FC0 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxFCCodeCol) = FC0
        Worksheets(datasheet).Cells(row1, MaxFCTypeCol) = FCT0
    ElseIf FC1 <= FC0 And FC1 <= FC2 And FC1 <= FC3 And FC1 >= FC4 And FC1 <> 0 And FC1 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxFCCodeCol) = FC1
        Worksheets(datasheet).Cells(row1, MaxFCTypeCol) = FCT1
    ElseIf FC2 <= FC0 And FC2 <= FC1 And FC2 <= FC3 And FC2 <= FC4 And FC2 <> 0 And FC2 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxFCCodeCol) = FC2
        Worksheets(datasheet).Cells(row1, MaxFCTypeCol) = FCT2
    ElseIf FC3 <= FC0 And FC3 <= FC1 And FC3 <= FC2 And FC3 <= FC4 And FC3 <> 0 And FC3 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxFCCodeCol) = FC3
        Worksheets(datasheet).Cells(row1, MaxFCTypeCol) = FCT3
    ElseIf FC4 <= FC0 And FC4 <= FC1 And FC4 <= FC2 And FC4 <= FC3 And FC4 <> 0 And FC4 <> "" Then
        Worksheets(datasheet).Cells(row1, MaxFCCodeCol) = FC4
        Worksheets(datasheet).Cells(row1, MaxFCTypeCol) = FCT4
    End If
    
    'set min FC code and type values  (FCcode 1 = interstate, FC code 7 = local)
    If FC0 >= FC1 And FC0 >= FC2 And FC0 >= FC3 And FC0 >= FC4 And FC0 <> 0 And FC0 <> "" Then
        Worksheets(datasheet).Cells(row1, MinFCCodeCol) = FC0
        Worksheets(datasheet).Cells(row1, MinFCTypeCol) = FCT0
    ElseIf FC1 >= FC0 And FC1 >= FC2 And FC1 >= FC3 And FC1 >= FC4 And FC1 <> 0 And FC1 <> "" Then
        Worksheets(datasheet).Cells(row1, MinFCCodeCol) = FC1
        Worksheets(datasheet).Cells(row1, MinFCTypeCol) = FCT1
    ElseIf FC2 >= FC0 And FC2 >= FC1 And FC2 >= FC3 And FC2 >= FC4 And FC2 <> 0 And FC2 <> "" Then
        Worksheets(datasheet).Cells(row1, MinFCCodeCol) = FC2
        Worksheets(datasheet).Cells(row1, MinFCTypeCol) = FCT2
    ElseIf FC3 >= FC0 And FC3 >= FC1 And FC3 >= FC2 And FC3 >= FC4 And FC3 <> 0 And FC3 <> "" Then
        Worksheets(datasheet).Cells(row1, MinFCCodeCol) = FC3
        Worksheets(datasheet).Cells(row1, MinFCTypeCol) = FCT3
    ElseIf FC4 >= FC0 And FC4 >= FC1 And FC4 >= FC2 And FC4 >= FC3 And FC4 <> 0 And FC4 <> "" Then
        Worksheets(datasheet).Cells(row1, MinFCCodeCol) = FC4
        Worksheets(datasheet).Cells(row1, MinFCTypeCol) = FCT4
    End If
        
    row1 = row1 + 1
Loop



''Edit max and min values
'row1 = 2
'If FCBlank = True Then
'    Do Until Worksheets(datasheet).Cells(row1, FCCodeCol1) = ""
'        FC1 = Worksheets(datasheet).Cells(row1, FCCodeCol1)
'        FC2 = Worksheets(datasheet).Cells(row1, FCCodeCol2)
'        FCT1 = Worksheets(datasheet).Cells(row1, FCTypeCol1)
'        FCT2 = Worksheets(datasheet).Cells(row1, FCTypeCol2)
'
'        If FC2 < FC1 Then
'            Worksheets(datasheet).Cells(row1, FCCodeCol1) = FC2
'            Worksheets(datasheet).Cells(row1, FCCodeCol2) = FC1
'
'            Worksheets(datasheet).Cells(row1, FCTypeCol1) = FCT2
'            Worksheets(datasheet).Cells(row1, FCTypeCol2) = FCT1
'        End If
'        row1 = row1 + 1
'    Loop
'Else
'    Do Until Worksheets(datasheet).Cells(row1, FCCodeCol1) = ""
'        FC1 = Worksheets(datasheet).Cells(row1, FCCodeCol1)
'        FC2 = Worksheets(datasheet).Cells(row1, FCCodeCol2)
'        FC3 = Worksheets(datasheet).Cells(row1, FCCodeCol3)
'
'        If FC3 < FC2 Then
'            Worksheets(datasheet).Cells(row1, FCCodeCol2) = FC3
'            Worksheets(datasheet).Cells(row1, FCCodeCol3) = FC2
'
'            Worksheets(datasheet).Cells(row1, FCTypeCol2) = FCT3
'            Worksheets(datasheet).Cells(row1, FCTypeCol3) = FCT2
'        End If
'
'        If FC2 < FC1 Then
'            Worksheets(datasheet).Cells(row1, FCCodeCol1) = FC2
'            Worksheets(datasheet).Cells(row1, FCCodeCol2) = FC1
'
'            Worksheets(datasheet).Cells(row1, FCTypeCol1) = FCT2
'            Worksheets(datasheet).Cells(row1, FCTypeCol2) = FCT1
'        End If
'
'       If FC3 < FC2 Then
'           Worksheets(datasheet).Cells(row1, FCCodeCol2) = FC3
'            Worksheets(datasheet).Cells(row1, FCCodeCol3) = FC2
'
'            Worksheets(datasheet).Cells(row1, FCTypeCol2) = FCT3
'            Worksheets(datasheet).Cells(row1, FCTypeCol3) = FCT2
'        End If
'        row1 = row1 + 1
'    Loop
    
'delete unneeded columns
Worksheets(datasheet).Columns(FCTypeCol4).EntireColumn.Delete
Worksheets(datasheet).Columns(FCTypeCol3).EntireColumn.Delete
Worksheets(datasheet).Columns(FCTypeCol2).EntireColumn.Delete
Worksheets(datasheet).Columns(FCCodeCol4).EntireColumn.Delete
Worksheets(datasheet).Columns(FCCodeCol3).EntireColumn.Delete
Worksheets(datasheet).Columns(FCCodeCol2).EntireColumn.Delete


'End If


End Sub

Sub Seg_Check_Headers(Data As Integer, datasheet As String, datawb As String)
' Check_Headers Macro
'   (1) If Data = 1, clean up AADT data based on selected year range.
'   (2) Compare incoming column data headers to what is expected. If there is a mismatch
'       the correct headers window opens so that the user can select the header that corresponds
'       to a specific data type. This is needed so that the macros know which columns are which
'       when rearranging the data.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Sam Mineer, BYU, 2016
' Commented by: Sam Mineer, BYU, 2016
   
'Activate data prep workbook
'Workbooks(CurrentWkbk).Activate

'Declare variables
Dim row1, row2, row3, col1, col2 As Integer             'Row and column counters
Dim DupCount As Integer                                 'Counter for multiple similar headers (i.e. AADT)
Dim HeadCount As Integer                                'Number of headers counter
Dim FirstRow As Integer                                 'Row number used in AADT simplification
Dim header, Header2 As String                           'Header values used in AADT simplification
Dim FY1, FY2, FY3, FY4, FY5, FY6, FY7 As Boolean        'Boolean variables for AADT simplification
Dim mainwb As String


' Data Key
'  1 = Segmented Data
'  2 = Crashes
'  3 = BeforeAfter Segments
'  4 = AADT

'Activate home sheet
mainwb = ActiveWorkbook.Name
Workbooks(mainwb).Sheets("Home").Activate

'Clear temporary headings list. Clear previous heading data from check headers info
Sheets("Key").Activate
Range("AX4:AX200").ClearContents
Range("BA4:BE4").ClearContents
Range("BB5").ClearContents

'Copy and paste headings from data file to temporary headings list
Workbooks(datawb).Sheets(datasheet).Activate
    'Windows(FileName1).Activate
Rows("1:1").Copy
    'Windows(CurrentWkbk).Activate
Workbooks(mainwb).Sheets("Key").Activate
    'Sheets("Key").Activate
Range("AX4").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
Application.CutCopyMode = False

'Count how many Headers there are
row1 = 4
col2 = 50 'For row AX
HeadCount = 0
Do While Worksheets("Key").Cells(row1, col2) <> ""
    HeadCount = HeadCount + 1
    row1 = row1 + 1
Loop
Sheets("Key").Range("BD4") = HeadCount    'Set BA4 of Key sheet to HeadCount for later use

'Assign column values based on which file name is selected. Column numbers are of expected headers.
col1 = 27 + (Data - 1) * 4

'Check Header values
row1 = 4                                                'Expected header list row
Sheets("Key").Range("BC4") = col1             'Set AZ4 of Key to column number to be used later

'Cycle through expected headers to see if headers are found in incoming header list
Do While Sheets("Key").Cells(row1, col1) <> ""
    header = Sheets("Key").Cells(row1, col1)          'Expected header
    row2 = 4                                                    'Incoming header list row
    'Cycle through incoming headers to see if expected header is found
    Do While Sheets("Key").Cells(row2, col2) <> ""
        'If header equals expected header, then exit loop so that row2 is left on found header row
        If Sheets("Key").Cells(row2, col2) = header Then
            Exit Do
        End If
        row2 = row2 + 1
    Loop
    
    'If the cell is blank, it means the Header was not found and the counter is at the end of the
    'headers list. Open form to choose new Header.
    If Worksheets("Key").Cells(row2, col2) = "" Then
        'Assign values to be used in correct header process
        Sheets("Key").Range("BA4") = Sheets("Key").Cells(row1, col1)        'Expected header
        Sheets("Key").Range("BB4") = Sheets("Key").Cells(row1, col1 + 1)    'Description
        Sheets("Key").Range("BE4") = Sheets("Key").Cells(row1, col1 + 3)    'Necessary?
        Sheets("Key").Range("BB5") = datasheet                                    'Working dataset filename
        Sheets("Key").Range("BB6") = datawb
        Sheets("Home").Activate
        
        'Screen Updating ON
        Application.ScreenUpdating = True
        
        'Show frmCorrectHeaders:
        '   (1) Asks user to choose header from incoming list that corresponds best to expected header.
        formCorrectHeaders.Show
        
        'Screen Updating OFF
        Application.ScreenUpdating = False
        
        'Change filename to data prep workbook to prevent future crashes
        'If a filename of an unopened workbook is left in cell, the next dataset _
        that runs will return an error.
        Sheets("Key").Range("BB5") = datasheet
    End If
    row1 = row1 + 1                     'Add to row1 to check next expected header
Loop

'Select A1 on Home sheet to clear other selections
Worksheets("Home").Activate
    
End Sub

Sub Int_Check_Headers(Data As Integer, datasheet As String, datawb As String)
' Check_Headers Macro
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Sam Mineer, BYU, 2016
' Commented by: Sam Mineer, BYU, 2016
   
'Activate data prep workbook
'Workbooks(CurrentWkbk).Activate

'Declare variables
Dim row1, row2, row3, col1, col2 As Integer             'Row and column counters
Dim DupCount As Integer                                 'Counter for multiple similar headers (i.e. AADT)
Dim HeadCount As Integer                                'Number of headers counter
Dim FirstRow As Integer                                 'Row number used in AADT simplification
Dim header, Header2 As String                           'Header values used in AADT simplification
Dim FY1, FY2, FY3, FY4, FY5, FY6, FY7 As Boolean        'Boolean variables for AADT simplification
Dim mainwb As String
Dim HeaderFound As Boolean
Dim i, j, k As Integer
Dim RCol(1 To 100) As Variant

' Data Key
'  1 = Intersection Data
'  2 = Crashes

'Activate home sheet
mainwb = ActiveWorkbook.Name
Workbooks(mainwb).Sheets("Home").Activate

'Clear temporary headings list. Clear previous heading data from check headers info
Sheets("Key").Activate
Range("AX4:AX200").ClearContents
Range("BA4:BE4").ClearContents
Range("BB5").ClearContents

'Copy and paste headings from data file to temporary headings list
Workbooks(datawb).Sheets(datasheet).Activate
    'Windows(FileName1).Activate
Rows("1:1").Copy
    'Windows(CurrentWkbk).Activate
Workbooks(mainwb).Sheets("Key").Activate
    'Sheets("Key").Activate
Range("AX4").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
Application.CutCopyMode = False

'Count how many Headers there are
row1 = 4
col2 = 50 'For row AX
HeadCount = 0
Do While Worksheets("Key").Cells(row1, col2) <> ""
    HeadCount = HeadCount + 1
    row1 = row1 + 1
Loop
Sheets("Key").Range("BD4") = HeadCount    'Set BD4 of Key sheet to HeadCount for later use

'Assign column values based on which file name is selected. Column numbers are of expected headers.
col1 = 1

Do Until Sheets("Key").Cells(1, col1) = "Intersection Check Headers"    'Find intersection check headers columns
    col1 = col1 + 1
Loop

Do Until Sheets("Key").Cells(2, col1) = Data        'Find data columns
    col1 = col1 + 1
Loop

'Check Header values
row1 = 4                                                'Expected header list row
Sheets("Key").Range("BC4") = col1             'Set AZ4 of Key to column number to be used later

'Cycle through expected headers to see if headers are found in incoming header list
Do While Sheets("Key").Cells(row1, col1) <> ""
    header = Sheets("Key").Cells(row1, col1)          'Expected header
    row2 = 4                                                    'Incoming header list row
    'Cycle through incoming headers to see if expected header is found
    Do While Sheets("Key").Cells(row2, col2) <> ""
        'If header equals expected header, then exit loop so that row2 is left on found header row
        If Sheets("Key").Cells(row2, col2) = header Then
            Exit Do
        End If
        row2 = row2 + 1
    Loop
    
    'If the cell is blank, it means the Header was not found and the counter is at the end of the
    'headers list. Open form to choose new Header.
    If Worksheets("Key").Cells(row2, col2) = "" Then
        'Assign values to be used in correct header process
        Sheets("Key").Range("BA4") = Sheets("Key").Cells(row1, col1)        'Expected header
        Sheets("Key").Range("BB4") = Sheets("Key").Cells(row1, col1 + 1)    'Description
        Sheets("Key").Range("BE4") = Sheets("Key").Cells(row1, col1 + 3)    'Necessary?
        Sheets("Key").Range("BB5") = datasheet                                    'Working dataset filename
        Sheets("Key").Range("BB6") = datawb
        Sheets("Home").Activate
        
        'Screen Updating ON
        Application.ScreenUpdating = True
        
        'Show frmCorrectHeaders:
        '   (1) Asks user to choose header from incoming list that corresponds best to expected header.
        formCorrectHeaders.Show
        
        'Screen Updating OFF
        Application.ScreenUpdating = False
        
        'Change filename to data prep workbook to prevent future crashes
        'If a filename of an unopened workbook is left in cell, the next dataset _
        that runs will return an error.
        Sheets("Key").Range("BB5") = datasheet
    End If
    row1 = row1 + 1                     'Add to row1 to check next expected header
Loop


'Rename column headers to match the final desired header
col2 = 1
Do While Worksheets(datasheet).Cells(1, col2) <> ""
    row2 = 4
    Do Until Worksheets("Key").Cells(row2, col1) = ""
        If Worksheets(datasheet).Cells(1, col2) = Worksheets("Key").Cells(row2, col1) Then
            Worksheets(datasheet).Cells(1, col2) = Worksheets("Key").Cells(row2, col1 + 2)
            Exit Do
        End If
        row2 = row2 + 1
    Loop
    col2 = col2 + 1
Loop

'Select A1 on Home sheet to clear other selections
Worksheets("Home").Activate
    
End Sub


Sub UTMtoLatLong()

'Declare variables
Dim row1 As Long
Dim ZoneNS, datasheet As String
Dim LongZone, utmYCol, utmXCol As Integer
Dim bConst, aConst, eConst, eSqConst, k0Const As Double
Dim e1Const, C1Const, C2Const, C3Const, C4Const As Double
Dim Northing, Easting, CNorthing, EastPrime As Double
Dim ArcLength, mu, FootprintLat, C1, T1, N1, R1, d As Double
Dim Fact1, Fact2, Fact3, Fact4, LoFact1, LoFact2, LoFact3 As Double
Dim DeltaLong, RawLat, Latitude, Longitude As Double
Dim ZoneCM As Integer

'Variable values
row1 = 2
ZoneNS = "N"
LongZone = 12
utmYCol = 1
utmXCol = 1
datasheet = "CrashInput"

bConst = 6356752.3142                   'Polar axis
aConst = 6378137                        'Equ Rad
eConst = 8.18191909289069E-02           'ecc
eSqConst = 0.006739496756587            'e'^2
k0Const = 0.9996                        'Scale

e1Const = 1.67922038993738E-03          'Footprint latitude values
C1Const = 2.51882658972117E-03
C2Const = 3.70094905128485E-06
C3Const = 7.44781381478857E-09
C4Const = 1.7035993382807E-11

'Find columns
Do Until Sheets(datasheet).Cells(1, utmYCol) = "UTM_Y"
    utmYCol = utmYCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, utmXCol) = "UTM_X"
    utmXCol = utmXCol + 1
Loop


'
Do While Sheets(datasheet).Cells(row1, 1) <> ""
    Northing = Sheets(datasheet).Cells(row1, utmYCol)
    Easting = Sheets(datasheet).Cells(row1, utmXCol)
    
    If ZoneNS = "N" Then
        CNorthing = Northing
    Else
        CNorthing = 10000000 - Northing
    End If
    
    EastPrime = 500000 - Easting
    ArcLength = Northing / k0Const
    mu = ArcLength / (aConst * (1 - (eConst ^ 2 / 4) - (3 * eConst ^ 4 / 64) - (5 * eConst ^ 6 / 256)))
    FootprintLat = mu + (C1Const * Sin(2 * mu)) + (C2Const * Sin(4 * mu)) + (C3Const * Sin(6 * mu)) + (C4Const * Sin(8 * mu))
    C1 = eSqConst * (Cos(FootprintLat)) ^ 2
    T1 = (Tan(FootprintLat)) ^ 2
    N1 = aConst / (1 - (eConst * Sin(FootprintLat)) ^ 2) ^ (1 / 2)
    R1 = aConst * (1 - eConst ^ 2) / (1 - (eConst * Sin(FootprintLat)) ^ 2) ^ (3 / 2)
    d = EastPrime / (N1 * k0Const)
    Fact1 = N1 * Tan(FootprintLat) / R1
    Fact2 = (d ^ 2) / 2
    Fact3 = (5 + (3 * T1) + (10 * C1) - (4 * C1 ^ 2) - (9 * eSqConst)) * (d ^ 4) / 24
    Fact4 = (61 + (90 * T1) + (298 * C1) + (45 * T1 ^ 2) - (252 * eSqConst) - (3 * C1 ^ 2)) * (d ^ 6) / 720
    LoFact1 = d
    LoFact2 = (1 + (2 * T1) + C1) * (d ^ 3) / 6
    LoFact3 = (5 - (2 * C1) + (28 * T1) - (3 * (C1 ^ 2)) + (8 * eSqConst) + (24 * (T1 ^ 2))) * (d ^ 5) / 120
    DeltaLong = (LoFact1 - LoFact2 + LoFact3) / Cos(FootprintLat)
    ZoneCM = (6 * LongZone) - 183
    RawLat = 180 * (FootprintLat - (Fact1 * (Fact2 + Fact3 + Fact4))) / WorksheetFunction.Pi
    
    If ZoneNS = "N" Then
        Latitude = RawLat
    Else
        Latitude = -1 * RawLat
    End If
    
    Longitude = ZoneCM - (DeltaLong * 180 / WorksheetFunction.Pi)
    
    Sheets(datasheet).Cells(row1, utmYCol) = Latitude
    Sheets(datasheet).Cells(row1, utmXCol) = Longitude
    
    row1 = row1 + 1
Loop

Sheets(datasheet).Cells(1, utmYCol) = "LATITUDE"
Sheets(datasheet).Cells(1, utmXCol) = "LONGITUDE"


End Sub

Sub FillBlanks()

Dim row, col As Long

For row = 2 To 1172
    For col = 28 To 60
        If Cells(row, col) = "" Or Cells(row, col) = " " Then
            Cells(row, col) = 0
        End If
    Next col
Next row

End Sub

Function FindColumn(colName As String, wksht As String, HeadRow As Integer, Necessary As Boolean)
    Dim col1 As Integer
    Dim continue As Integer
    col1 = 1
    If colName <> "" Then
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
                        Exit Do
                    End If
                Else
                    col1 = 0
                    Exit Do
                End If
            End If
        Loop
    Else
        Do Until Worksheets(wksht).Cells(HeadRow, col1) = ""
            col1 = col1 + 1
        Loop
    End If
    FindColumn = col1
End Function

Function Col_Letter(lngCol As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
