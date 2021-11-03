Attribute VB_Name = "CompileAnalysisReport"
Option Explicit

Sub CompileAnalysisReports()

'*************************************************
'
'  The purpose of this code is to create Analyis
'    Reports for the "Top 20" list produced from
'    the Statistical Model.
'  These codes should be ran after (1) the crash
'    database has been created and (2) after the
'    "Combo" and "CrashFactors" datasets have
'    been created.
'  General Steps of this code:
'     (1) Save "Report Compiler" as a new file in
'             a new directory, for this specific
'             analysis Group
'     (2) Copy over the "Combo" and "CrashFactors"
'             sheets from the relative "Combo" file
'     (3) For each segment in "Combo" (the "Top 20"
'             list), create a new report sheet and
'             fill out report
'
'   A copy of the CrashFactors and Vehicle data is
'     copied as a reference for error checking.
'   Once report is complete, the report and
'     corresponding copy of the CrashFactors and
'     Vehicle Crash data are moved to a new workbook.
'   This process is repeated for each line on the
'     "Combo" sheet (model built for 1 or 200+
'     segments, as data is avilable)
'
'  Comments by Sam Mineer, Brigham Young University,
'  Updated 11/2015  05/2016  03/2017
'
'*************************************************

'-------------------------------------------------
'
'  Define variables (descriptions given after words)
'
'-------------------------------------------------
Dim n As Integer            'counting variables, used in various manners
Dim i As Integer            'counting variables, used in various manners
Dim j As Integer            'counting variables, used in various manners
Dim k As Integer            'counting variables, used in various manners
Dim combofp As String       'filepath to Combo and CrashFactors data
Dim combofn As String       'filename of Combo Workbook, with Combo and CrashFactors datasets
Dim combosn As String       'sheet name for "Combo" sheet
Dim crafasn As String       'sheet name for "CrashFactors" sheet
Dim crashfp As String       'filepath to crash data
Dim crashfn As String       'filename of Crash Workbook, with vehicle
Dim vehiclesn As String     'sheet name for "vehicle" sheet
Dim reportfp As String      'filepath of where to save Analysis Reports
Dim reportfn As String      'filename of Report Compiler
Dim segname As String       'name of new segment report, changes with each report number
Dim comborow As Integer     'row counter for Combo sheet
Dim crafarow As Integer     'row counter for CrashFactors sheet
Dim reportrow As Integer    'row counter for reportrow (individual ones)
Dim roadname As String      'String variable for Road Name of segment (Table 1)
Dim crashmp As String       'String variable for Crash Milepoint (Table 6, Table 7)
Dim begmp As String         'String variable for Beginning Mile Post (Table 1, Table 2, Table 3)
Dim endmp As String         'String variable for End Mile Post (Table 1, Table 2, Table 3)
Dim County As String        'String variable for County, used to determine UDOT Region (Table 1)
Dim fclas As String         'String variable for Functional Class (Table 2)
Dim nlanes As Integer       'String variable for Number of Lanes (Table 2, Table 4)
Dim AADT As Long            'Value for AADT (Table 2)
Dim Speed As Integer        'Value for Speed Limit (Table 2)
Dim ncrashes As Integer     'Value for total number of Crashes for segment (Table 5)
Dim Median As String        'String variable for Median, combination of type and width (Table 4)
Dim IPM As String           'String variable for IPM, frequency/IPM value (Table 4)
Dim SPM As String           'String variable for SPM, frequency/SPM value (Table 4)
Dim Shoulder As String      'String variable for Shoulder, combination of type and width (Table 4)
Dim Grade As String         'String varaible for Max Grade (Table 4)
Dim Curve As String         'String variable for Curves, combination of class, Length, Radius (Table 4)
Dim Lanes As String         'String variable for Lanes, combination of Thru and Auxiliary lanes (Table 4)
Dim WB As String            'String variable for Wall and Barriers (Table 4)
Dim Rumble As String        'String variable for Rumble Strips, Yes or No (Table 4)
Dim CrashID As String       'String variable for CrashID (Table 6, Table 7)
Dim subseg As Integer       'String variable for SubSegment (Table 6, Table 7)
Dim firhar As String        'String variable for First Harmful Event, uses "FirstHarm()" function to convert Integer to String (Table 7)
Dim mancol As String        'String variable for Manner of Collision, uses "MannerCollision()" function to convert Integer to String (Table 7)
Dim evseq1 As String        'String variable for Event Sequence #1, uses "MannerCollision()" function to convert Integer to String (Table 7)
Dim evseq2 As String        'String variable for Event Sequence #2, uses "EventSequence()" function to convert Integer to String (Table 7)
Dim evseq3 As String        'String variable for Event Sequence #3, uses "EventSequence()" function to convert Integer to String (Table 7)
Dim evseq4 As String        'String variable for Event Sequence #4, uses "EventSequence()" function to convert Integer to String (Table 7)
Dim moharm As String        'String variable for Most Harmful Event, uses "MostHarmful()" function to convert Integer to String (Table 7)
Dim VehMan As String        'String variable for Vehicle Manuever, uses "VehicleManuever()" function to convert Integer to String (Table 7)
Dim vehcrow As Integer      'Row counter for cycling through Vehicle Crash data
Dim sev5 As Integer         'Counter for number of Severe 5 Crashes
Dim sev4 As Integer         'Counter for number of Severe 4 Crashes
Dim sev3 As Integer         'Counter for number of Severe 3 Crashes
Dim sumfac As Integer       'Summation of "Y" factors, for producing a "X/Z" value
Dim vehc As String          'name of sheet containing a copy of Vehicle Crash data, for a specific segment
Dim facc As String          'name of sheet containing a copy of CrashFactors data, for a specific segment
Dim modeltype As String     'name of model used, entered by user
Dim dataranges As String    'range of years for data, entered by user
Dim SRank As Integer
Dim CRank As Integer
Dim RRank As Integer
Dim nrow As Integer
Dim cmdline As String
Dim pcrashes As Double

Dim top8factor As String

Dim factorarray(1 To 8) As String

Dim regnum As String        'Region Number

'-------------------------------------------------
'
'  Start of code
'
'  Sets ups the document before copying data over
'
'-------------------------------------------------

Dim guiwb As String

guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")


'  Prompt user to enter the model used for analysis
modeltype = InputBox("Enter the model used. (UCPM or UCSM)", "Enter the model used", "UCPM/UCSM")
If modeltype = "" Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
Else
    modeltype = UCase(modeltype)
End If

'  Prompt user to enter the range of dates for the data
dataranges = InputBox("Enter the range of dates for data source." & Chr(10) & "Example: [2008-2012]", "Enter the range of dates for data source", "20X1-20X5")
If dataranges = "" Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If
'  Sheet names with the vehicle cand crash factors
vehc = "VehCrash"       '<--- Can be changed
facc = "FacCrash"       '<--- Can be changed

    'Windows("Report Compiler (start here)2.xlsm").Activate
Sheets("Main").Activate

'  Prompts the user to select the location
'    to save the reports
Dim fldr As FileDialog
Dim sItem As String
Dim strPath As String
Dim msgboxstatus As Variant

msgboxstatus = MsgBox("Select the Folder to Save the Report Files", vbOKCancel, "User Prompt")

'  If the user clicks 'Cancel', the macro is aborted
If msgboxstatus = vbOK Then
ElseIf msgboxstatus = vbCancel Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select Location to Save Analysis Reports (Select a Folder)"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode1
    sItem = .SelectedItems(1)
End With
NextCode1:
  'GetFolder = sItem
reportfp = sItem
Set fldr = Nothing

'  If the filepath doesn't have an ending "\",
'    then one is added
If Right(reportfp, 1) <> "\" Then
    reportfp = reportfp & "\"
End If

'  Creates new directory for saving this batch of
'    analysis reports, with unique date and time
reportfp = reportfp & "Output " & Month(Date) & "-" & Day(Date) & " " & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & "\"
MkDir reportfp

'  Resave ReportCompiler, to preserve
'    "Report Compiler (start here)" file,
'    with unique date and time
ActiveWorkbook.SaveAs Filename:= _
    reportfp & "Report Compiler " & Month(Date) & "-" & Day(Date) & " " & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & ".xlsm" _
    , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
reportfn = ActiveWorkbook.Name

'  Communicates to user that the code is running
'    and not to click elsewhere.
'  This ensures that the user doesn't inadvertantly
'    click somewhere else and create and error,
'    the sign will be given.
    Range("H2:L12").Merge
    Range("H2:L12").Font.Bold = True
    Range("H2:L12").Font.Italic = True
    With Range("H2:L12")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
'  Prints "Code Running: Please Wait"
    Range("H2:L12") = "Code Running: Please Wait."
    Range("H2:L12").Font.Size = 24
    With Range("H2:L12").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274    'Light Green
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'  Waits for 1 second, to allow time for the textbox to appear
Application.Wait (Now + TimeValue("00:00:01"))

'  Identify filename of Combo Workbook, as given by the user
msgboxstatus = MsgBox("Open the Workbook with the Combo and CrashFactor sheets", vbOKCancel, "Open Combo and CrashFactor data")

'  If the user clicks 'Cancel', the macro is aborted
If msgboxstatus = vbOK Then
ElseIf msgboxstatus = vbCancel Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If

Application.Dialogs(xlDialogOpen).Show

combofn = ActiveWorkbook.Name

'Identifies the filename of worksheet contatining the "Combo" sheet and "CrashFactors" sheet
combosn = "Combo"
crafasn = "CrashFactors"

'Identify filename of Crash Workbook, with Vehicle Crash Data, as given by the user

msgboxstatus = MsgBox("Open the Workbook with the Vehicle Crash Data", vbOKCancel, "Open Vehicle Crash Workbook")

'  If the user clicks 'Cancel', the macro is aborted
If msgboxstatus = vbOK Then
ElseIf msgboxstatus = vbCancel Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If

'  User selects the Vehicle Crash Workbook
Application.Dialogs(xlDialogOpen).Show

'  Waits for 1 second to give time for the sheet transfer to be successful
Application.Wait (Now + TimeValue("00:00:01"))

crashfn = ActiveWorkbook.Name
vehiclesn = "vehicle"
ActiveSheet.Name = vehiclesn
Sheets.Add After:=Sheets(Sheets.count)
ActiveSheet.Name = "FilterByCrash"     'For filtering the vehicle crash data

'  This line with disable the screen update feature, to prevent a "strobe light" effect
'  Turning this feature off will remove the rapid changing of screen, to be easier on the eyes
Windows(reportfn).Activate
'  Waits for 1 second to give time for the sheet transfer to be successful
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False

'  Moves a copy of the "Combo" and "CrashFactors" sheets to the "Report Compiler" file, to use for future reference
Windows(combofn).Activate       'CrashFactors
Sheets(crafasn).Copy After:=Workbooks(reportfn).Sheets("Main")
Application.Wait (Now + TimeValue("00:00:01"))          'Waits for 1 second to give time for the sheet transfer to be successful
Windows(combofn).Activate       'Combo
Sheets(combosn).Activate
Sheets(combosn).Copy After:=Workbooks(reportfn).Sheets("Main")
Application.Wait (Now + TimeValue("00:00:01"))

'----------------------------------------------------------------------------------
'
'  In this section, the code will loop for each segment listed in the "Combo" sheet
'
'----------------------------------------------------------------------------------

'  Sets the comborow counter for 2, which is the top of the list. The Do Loop will step through each line until a blank entry is found
comborow = 2
Do While Sheets(combosn).Cells(comborow, FindColumn("State_Rank")) <> ""
    '  Resets variables
    ncrashes = 0
    sev5 = 0
    sev4 = 0
    sev3 = 0
    
    '  Creates a new sheet for the CrashFactors
    '  Copies row header for CrashFactors
    Sheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.Name = facc     'Copy of CrashFactors sheet   FLAGGED: Is that statement backwards?
    Sheets(crafasn).Activate
    Rows(1).Copy
    Sheets(facc).Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    '  Creates a new sheet for the Vehicle Crash Data, to be copied later
    Sheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.Name = vehc
    ActiveWindow.Zoom = 80


    '  Gives a name for a segment, starting at 1 (2-1) and ending when the row is empty
    Sheets(combosn).Activate
    segname = modeltype & "-StateRank" & Sheets(combosn).Cells(comborow, FindColumn("State_Rank"))   '               "-Seg" &              ' CStr(comborow - 1)
    'May need to change this...     *********************************************************************************
    
    
    
    
    
    '  Copy the BlankReport template for the new segment report, preserves BlankReport template
    Windows(reportfn).Activate
    Sheets("SegBlankReport").Activate
    Sheets("SegBlankReport").Copy After:=Sheets("SegBlankReport")
    ActiveSheet.Name = segname
        
    '  Extracts data from the copy of the "Combo" sheet, for Table 1, Table 2, Table 3, Table 4 and Table 7
    Windows(combofn).Activate
    Sheets(combosn).Activate
    '  Road Name, Beginning MP, End MP, Functional Class, County, Number of Lanes, Speed
    roadname = Cells(comborow, FindColumn("LABEL"))
    begmp = Cells(comborow, FindColumn("BEG_MILEPOINT"))
    endmp = Cells(comborow, FindColumn("END_MILEPOINT"))
    fclas = Cells(comborow, FindColumn("FC_Type"))
    County = Cells(comborow, FindColumn("COUNTY"))
    nlanes = Cells(comborow, FindColumn("THRU_LANES"))
    Speed = Cells(comborow, FindColumn("SPEED_LIMIT"))
    '-------
    '---AADT
    '-------
    j = 1
    Do Until LCase(Left(Cells(1, j), 4)) = "aadt"
        j = j + 1
    Loop
    AADT = Cells(comborow, j)
    '--------------------
    '---Number of Crashes
    '--------------------
    ncrashes = Cells(comborow, FindColumn("Total_Crashes"))
    pcrashes = Cells(comborow, FindColumn("Predicted_Number_of_Crashes"))
    '---------
    '---Median
    '---------
    j = FindColumn("Type (Median)")
    If Cells(comborow, j) = "" Then     'Checks if blank
        Median = "None"
    Else
        Median = Cells(comborow, j) & "," & Chr(10) & CStr(Round(Cells(comborow, j + 1), 0)) & " ft"     'Joins type and width
    End If
    '------
    '---IPM
    '------
    j = FindColumn("Intersection Count")
    IPM = CStr(CStr(Cells(comborow, j)) & "/" & CStr(Round(Cells(comborow, j + 1), 1)))     'Joins frequency and IPM
    '------
    '---SPM
    '------
    j = FindColumn("Sign Count")
    SPM = CStr(Cells(comborow, j) & "/" & CStr(Round(Cells(comborow, j + 1), 1)))     'Joins frequency and SPM
    '-----------
    '---Shoulder
    '-----------
    j = FindColumn("Common Material (Shoulder)")
    If Cells(comborow, j + 2) = 0 Then      'Checks if blank
        Shoulder = Cells(comborow, j)
    Else
        Shoulder = Cells(comborow, j) & "," & Chr(10) & CStr(Round(Cells(comborow, j + 2), 0)) & " ft"       'Joins Material and width
    End If
    '-----------
    '---MaxGrade
    '-----------
    Grade = CStr(Round(Cells(comborow, FindColumn("Max Grade")), 2)) & " (max)"         'Could add more information for grade changes, etc.
    '--------
    '---Curve
    '--------
    j = FindColumn("Curve Class")
    If Cells(comborow, j) = "" Then
        Curve = "None"
    Else
        Curve = "Class " & Cells(comborow, j) & "," & Chr(10) & "L = " & CStr(Round(Cells(comborow, j + 3), 0)) & "," & Chr(10) & "R = " & CStr(Round(Cells(comborow, j + 2), 0))
    End If
    '--------
    '---Lanes
    '--------
    j = FindColumn("Left Turn Lane")
    Lanes = CStr(nlanes) & " Thru"      'Always prints the number of thru lanes
    For i = j To (j + 7)                'If Auxiliary lane value is greater than 1, then joins on to the Lanes string
        If Cells(comborow, i) > 0 Then
            Lanes = Lanes & "," & Chr(10) & Cells(1, i)
        End If
    Next i
    '---------------
    '---Wall/Barrier
    '---------------
    j = FindColumn("Wall (W)/Barrier (B)")
    If (Left(Cells(comborow, j), 1) = "W") Then
        WB = "Yes (Wall)"
    Else
        WB = "No (Wall)"
    End If
    If Cells(comborow, j + 1) <> "" Then
        WB = WB & "," & Chr(10) & Cells(comborow, j + 1) & " (Center Barrier)," & Chr(10) & Cells(comborow, j + 2) & " (Outside Barrier)"
    Else
        WB = WB & "," & Chr(10) & "No (Barrier)"
    End If
    '---------
    '---Rumble
    '---------
    j = FindColumn("Rumble Strips")
    If Cells(comborow, j) = "Y" Then
        Rumble = "Yes"
    Else
        Rumble = "No"
    End If
    '---------
    '---State, Region, County Ranking
    '---------
    j = FindColumn("State_Rank")
    SRank = Cells(comborow, j)
    RRank = Cells(comborow, j + 1)
    CRank = Cells(comborow, j + 2)
    
    '------------------------------------------------------------------------
    '  Print the gathered data onto Segment Report Sheet
    Windows(reportfn).Activate
    Sheets(segname).Activate

    '  Prints to Table 1, finds the row to print to
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 1"
        n = n + 1
    Loop
        
    If LCase(Right(roadname, 1)) = "p" Then
        Cells(n + 2, 4) = "Positive"
    Else
        Cells(n + 2, 4) = "Negative"
    End If
    
    '  Give Common Road Name to Segment
    roadname = Replace(roadname, "N", "")
    roadname = Replace(roadname, "P", "")
    k = CInt(roadname)
    If k = 15 Or k = 70 Or k = 80 Or k = 84 Or k = 215 Then
        Cells(n + 1, 4) = "I-" & k
    ElseIf k = 6 Or k = 40 Or k = 50 Or k = 89 Or k = 91 Or k = 163 Or k = 189 Or k = 191 Or k = 491 Then
        Cells(n + 1, 4) = "US-" & k
    Else
        Cells(n + 1, 4) = "SR-" & k
    End If
    Cells(n + 3, 4) = Round(CDbl(begmp), 3)
    Cells(n + 3, 5) = Round(CDbl(endmp), 3)
    Cells(n + 4, 4) = Round(CDbl(endmp), 3) - Round(CDbl(begmp), 3)
    Cells(n + 4, 4).NumberFormat = "0.000"
    Cells(n + 5, 4) = dataranges
    Cells(n + 1, 8) = modeltype
    Cells(n + 2, 8) = SRank 'comborow - 1  #srank
    Cells(n + 4, 8) = CRank
    Cells(n + 4, 9) = County
    regnum = UDOTRegion(County)
    Cells(n + 3, 8) = RRank
    Cells(n + 3, 9) = regnum
    Cells(n + 5, 8) = "To be completed by engineer..."
    
    '  Prints to Table 2
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 2"
        n = n + 1
    Loop
    Cells(n + 1, 4) = fclas
    Cells(n + 1, 8) = AADT
    Cells(n + 1, 8).NumberFormat = "#,##0"
    Cells(n + 2, 4) = nlanes
    Cells(n + 2, 8) = Speed
    
    '  Prints to Table 3
    'n = 1
    'Do Until Left(Cells(n, 1), 7) = "Table 3"
    '    n = n + 1
    'Loop
    'Cells(n + 2, 2) = 1
    'Cells(n + 2, 3) = begmp
    'Cells(n + 2, 6) = endmp
    'Cells(n + 2, 9) = Round(CDbl(endmp) - CDbl(begmp), 3)
    
    '  Prints to Table 3
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 3"
        n = n + 1
    Loop
    Cells(n + 2, 1) = CStr(Round(CDbl(begmp), 3)) & "-" & CStr(Round(CDbl(endmp), 3))
    Cells(n + 2, 2) = Median
    Cells(n + 2, 3).NumberFormat = "@"
    Cells(n + 2, 3) = IPM
    Cells(n + 2, 4).NumberFormat = "@"
    Cells(n + 2, 4) = SPM
    Cells(n + 2, 5) = Shoulder
    Cells(n + 2, 6) = Grade
    Cells(n + 2, 7) = Curve
    Cells(n + 2, 8) = Lanes
    Cells(n + 2, 9) = WB
    Cells(n + 2, 10) = Rumble
    
'------------------------------------------------
'
'  This section of code works on the Vehicle Crash
'    Data and Crash Factor Data
'
'------------------------------------------------

    '  Transfer Data from CrashFactors to blank report
    '---Crash count and severity
    '---Crash Milepoint
    '---Data from Crash and Vehicle Datasets
    '---Top CrashFactors
    
    '  Copy CrashFactors to sheet ,as a copy for referencing
    Sheets(crafasn).Activate
    crafarow = 2
    Do Until Cells(crafarow, FindColumn("State_Rank")) = SRank      '(comborow - 1)      'Counts rows until it finds the section or relevant segments
        crafarow = crafarow + 1
        
        If crafarow > 10000 Then
            '  Prints the frequency of Crash Severities to Table 5
            Windows(reportfn).Activate
            Sheets(segname).Activate
            n = 1
            Do Until Left(Cells(n, 1), 7) = "Table 4"
                n = n + 1
            Loop
            'Cells(n + 2, 1) = CStr(Round(CDbl(begmp), 3)) & "-" & CStr(Round(CDbl(endmp), 3))       'begmp & " - " & endmp
            Cells(n + 2, 1) = pcrashes
            Cells(n + 2, 1).NumberFormat = "0.0"
            Cells(n + 2, 3) = ncrashes
            Cells(n + 2, 5) = sev5
            Cells(n + 2, 7) = sev4
            Cells(n + 2, 9) = sev3
        
            GoTo Line999
        End If
    Loop
    
    'Cells(crafarow, 1).Select       'Selects extents of CrashFactors data to copy over
    n = crafarow
    Do Until Cells(n, 1) = ""
        n = n + 1
    Loop
    i = 1
    Do Until Cells(1, i + 1) = ""
        i = i + 1
    Loop
    Range(Cells(crafarow, 1), Cells(n + 1, i)).Copy
    Sheets(facc).Activate
    Cells(2, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit          'Autofit columns for visual appeal
        
    '  Copy Vehicle Crash Data, for referencing
    '  Uses the FilterCrashData function
    Windows(crashfn).Activate
    Sheets("FilterByCrash").Activate
    Columns(1).Clear
    Windows(reportfn).Activate
    '  Get list of CrashID's to filter
    Sheets(facc).Activate
    Range(Cells(2, FindColumn("CRASH_ID")), Cells(2, FindColumn("CRASH_ID")).End(xlDown)).Copy
    Windows(crashfn).Activate
    Sheets("FilterByCrash").Activate
    Cells(1, 1).Select
    ActiveSheet.Paste
    '*********************************************
    Call FilterbyCrashID
    '*********************************************
    'copy filtered crashID to other sheet
    Range(Range(Range("A1"), Range("A1").End(xlDown)), Range(Range("A1"), Range("A1").End(xlDown)).End(xlToRight)).Copy
    Windows(reportfn).Activate
    Sheets(vehc).Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit          'Autofit columns for visual appeal
    
    '  Add J rows to Table 6 for adequate space for Crash Data
    Sheets(facc).Activate
    j = 0
    Do Until Cells(2 + j, FindColumn("CRASH_ID")) = ""
        j = j + 1
    Loop            'j = number of CrashID's
    
    Sheets(segname).Activate
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 6"
        n = n + 1
    Loop
    For i = 1 To j - 1              'With one row already in place, (j-1) number of rows will be added to the table
        Rows(n + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    '  Cycles line by line through the Crash Data
    For i = 1 To j
        '  Reset factors
        CrashID = ""
        crashmp = ""
        firhar = ""
        mancol = ""
        evseq1 = ""
        evseq2 = ""
        evseq3 = ""
        evseq4 = ""
        moharm = ""
        VehMan = ""
        
        '  Defines CrashID, subsegment, First Harmful Event, Manner of Collision from "CrashFactors" sheet
        Sheets(facc).Activate
        CrashID = Round(CDbl(Cells(i + 1, FindColumn("CRASH_ID"))), 1)
        crashmp = CStr(Cells(i + 1, FindColumn("MILEPOINT")))
        firhar = FirstHarmful(Cells(i + 1, FindColumn("FIRST_HARMFUL_EVENT_ID")))
        mancol = MannerCollision(Cells(i + 1, FindColumn("MANNER_COLLISION_ID")))
        
        '  Increase severity count in segment
        If Cells(i + 1, FindColumn("CRASH_SEVERITY_ID")) = 3 Then
            sev3 = sev3 + 1
        ElseIf Cells(i + 1, FindColumn("CRASH_SEVERITY_ID")) = 4 Then
            sev4 = sev4 + 1
        ElseIf Cells(i + 1, FindColumn("CRASH_SEVERITY_ID")) = 5 Then
            sev5 = sev5 + 1
        End If
        
        '  Cycles through each vehicle in the Vehicle Crash sheet
        '  Can be multiple vehicles per CrashID
        vehcrow = 2     'Row counter for Vehicle Crash sheet
        Sheets(vehc).Activate
            Do Until Cells(vehcrow, 1) = ""             'Make sure these values match the "KEY" sheet
                If Round(CDbl(Cells(vehcrow, FindColumn("CRASH_ID"))), 0) = CrashID Then            'If CrashID's match, then extract data
                    If evseq1 = "Not Applicable" Or evseq1 = "Not Provided" Or evseq1 = "Unknown" Or evseq1 = "Invalid" Or evseq1 = "No Damage" Or evseq1 = "" Then  'Finds something more descriptive than blanks, "Not Applicable" or something similar
                        evseq1 = EventSequence(Cells(i + 1, FindColumn("EVENT_SEQUENCE_1_ID")))
                    Else
                    End If
                    If evseq2 = "Not Applicable" Or evseq2 = "Not Provided" Or evseq2 = "Unknown" Or evseq2 = "Invalid" Or evseq2 = "No Damage" Or evseq2 = "" Then
                        evseq2 = EventSequence(Cells(i + 1, FindColumn("EVENT_SEQUENCE_2_ID")))
                    Else
                    End If
                    If evseq3 = "Not Applicable" Or evseq3 = "Not Provided" Or evseq3 = "Unknown" Or evseq3 = "Invalid" Or evseq3 = "No Damage" Or evseq3 = "" Then
                        evseq3 = EventSequence(Cells(i + 1, FindColumn("EVENT_SEQUENCE_3_ID")))
                    Else
                    End If
                    If evseq4 = "Not Applicable" Or evseq4 = "Not Provided" Or evseq4 = "Unknown" Or evseq4 = "Invalid" Or evseq4 = "No Damage" Or evseq4 = "" Then
                        evseq4 = EventSequence(Cells(i + 1, FindColumn("EVENT_SEQUENCE_4_ID")))
                    Else
                    End If
                    
                    If moharm = "Not Applicable" Or moharm = "Not Provided" Or moharm = "Unknown" Or moharm = "Invalid" Or moharm = "No Damage" Or moharm = "" Then
                        moharm = EventSequence(Cells(i + 1, FindColumn("MOST_HARMFUL_EVENT_ID")))
                    Else
                    End If
                    
                    If VehMan = "" Then         'Need to join all vehicle manuever data for a given CrashID
                        VehMan = VehicleManuever(Cells(i + 1, FindColumn("VEHICLE_MANEUVER_ID")))
                    Else
                        VehMan = VehMan & "," & Chr(10) & VehicleManuever(Cells(i + 1, FindColumn("VEHICLE_MANEUVER_ID")))
                    End If
                    
                    vehcrow = vehcrow + 1
                Else
                    vehcrow = vehcrow + 1
                End If
            Loop

        '  Prints results to table 6
        Sheets(segname).Activate
        Cells(n + 1 + i, 1) = CrashID   'crashID
        Cells(n + 1 + i, 2) = crashmp   'Crash milepoint
        Cells(n + 1 + i, 3) = firhar    'FirstHarmful
        Cells(n + 1 + i, 4) = mancol    'MannerOfCollision
        Cells(n + 1 + i, 5) = evseq1    'Event Sequence1
        Cells(n + 1 + i, 6) = evseq2    'Event Sequence2
        Cells(n + 1 + i, 7) = evseq3    'Event Sequence3
        Cells(n + 1 + i, 8) = evseq4    'Event Sequence4
        Cells(n + 1 + i, 9) = moharm    'Most Harmful
        Cells(n + 1 + i, 10) = VehMan   'Vehicle Manuever (single or mulitple)
    Next i
         
    '  Formats Table 5 with lines, correct size text, alignment, etc.
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Font
        .Name = "Cambria"
        .Size = 8
        .Bold = False
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMajor
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    
    '  Fills out Table 5 ------------------------------------------------------------
    '  Adds J rows to Table 5 for adequate space for Crash Factors
    Sheets(segname).Activate
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 5"
        n = n + 1
    Loop
    
    For i = 1 To j - 1             'With one row already in place, (j-1) number of rows will be added to the table
        Rows(n + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
    '  Before picking the Top 8 factors, Sort factors by greatest to smallest (occurences on second to last row)
    Sheets(facc).Activate
    ActiveWorkbook.Worksheets(facc).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(facc).Sort.SortFields.Add Key:=Range(Cells(j + 2, FindColumn("PEDESTRIAN_INVOLVED")), Cells(j + 2, FindColumn("COLLISION_WITH_FIXED_OBJECT"))) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(facc).Sort
        .SetRange Range(Cells(1, FindColumn("PEDESTRIAN_INVOLVED")), Cells(j + 4, FindColumn("COLLISION_WITH_FIXED_OBJECT")))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
        
    '  Extracts CrashID's and Factors from Crash Factors sheet
    '  Copy all CrashID's
    Sheets(segname).Activate  'Sheets(facc).Select
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 6"
        n = n + 1
    Loop
    'FindColumn("CRASH_ID")), Cells(2, FindColumn("CRASH_ID") + 1)).Select
    Range(Range(Cells(n + 2, 1), Cells(n + 2, 2)), Range(Cells(n + 2, 1), Cells(n + 2, 2)).End(xlDown)).Copy
    'Sheets(segname).Select
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 5"
        n = n + 1
    Loop
    Range(Cells(n + 2, 1), Cells(n + 1 + i, 2)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        
    '  Copy all Top 8 Crash Factors
    Sheets(facc).Range(Cells(2, FindColumn("NUMBER_ONE_INJURIES") + 1), Cells(j + 1, FindColumn("NUMBER_ONE_INJURIES") + 8)).Copy
    Sheets(segname).Activate
    Range(Cells(n + 2, 3), Cells(n + 1 + j, 10)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    '  Formatting Table 6 data
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Font
        .Name = "Cambria"
        .Strikethrough = False
        .Bold = False
        .Size = 8
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMajor
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Cells(n + 2, 1), Cells(n + 1 + j, 10)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With

    '  Copies the Columns Headings of the Top 8 Crash Factors
    '  Headers
    Sheets(facc).Activate
    
    'Copy crash factors to an array
   
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 1)
    factorarray(1) = top8factor
    'factorarray.Add top8factor, 1
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 2)
    factorarray(2) = top8factor
    'factorarray.Add top8factor, 2
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 3)
    factorarray(3) = top8factor
    'factorarray.Add top8factor, 3
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 4)
    factorarray(4) = top8factor
    'factorarray.Add top8factor, 4
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 5)
    factorarray(5) = top8factor
    'factorarray.Add top8factor, 5
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 6)
    factorarray(6) = top8factor
    'factorarray.Add top8factor, 6
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 7)
    factorarray(7) = top8factor
    'factorarray.Add top8factor, 7
    top8factor = Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 8)
    factorarray(8) = top8factor
    'factorarray.Add top8factor, 8
    
        
    Sheets(segname).Cells(n + 1, 3) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 1), "_", " ")
    Sheets(segname).Cells(n + 1, 3) = StrConv(Sheets(segname).Cells(n + 1, 3), vbProperCase)
    Sheets(segname).Cells(n + 1, 4) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 2), "_", " ")
    Sheets(segname).Cells(n + 1, 4) = StrConv(Sheets(segname).Cells(n + 1, 4), vbProperCase)
    Sheets(segname).Cells(n + 1, 5) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 3), "_", " ")
    Sheets(segname).Cells(n + 1, 5) = StrConv(Sheets(segname).Cells(n + 1, 5), vbProperCase)
    Sheets(segname).Cells(n + 1, 6) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 4), "_", " ")
    Sheets(segname).Cells(n + 1, 6) = StrConv(Sheets(segname).Cells(n + 1, 6), vbProperCase)
    Sheets(segname).Cells(n + 1, 7) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 5), "_", " ")
    Sheets(segname).Cells(n + 1, 7) = StrConv(Sheets(segname).Cells(n + 1, 7), vbProperCase)
    Sheets(segname).Cells(n + 1, 8) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 6), "_", " ")
    Sheets(segname).Cells(n + 1, 8) = StrConv(Sheets(segname).Cells(n + 1, 8), vbProperCase)
    Sheets(segname).Cells(n + 1, 9) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 7), "_", " ")
    Sheets(segname).Cells(n + 1, 9) = StrConv(Sheets(segname).Cells(n + 1, 9), vbProperCase)
    Sheets(segname).Cells(n + 1, 10) = Replace(Sheets(facc).Cells(1, FindColumn("NUMBER_ONE_INJURIES") + 8), "_", " ")
    Sheets(segname).Cells(n + 1, 10) = StrConv(Sheets(segname).Cells(n + 1, 10), vbProperCase)
        
    '  With the Top 8 Crash Factors prints, Count the number of "Y"'s in the columns, print summation
    Sheets(segname).Activate
    For k = 3 To 10
    sumfac = 0
        For i = 1 To j
            If i <> j Then
            'counts if the cell has a "Y" in it
                If Cells(n + 1 + i, k) = "Y" Then
                    sumfac = sumfac + 1
                End If
            Else
            'count if the cell has a "Y" in it
            'Since the end of the table, Print the Summation. Example: (25/30)
                If Cells(n + 1 + i, k) = "Y" Then
                    sumfac = sumfac + 1
                End If
                Cells(n + 2 + j, k) = "'" & CStr(sumfac) & "/" & CStr(j) & ""
            End If
        Next i
    Next k
        
    '  Prints the frequency of Crash Severities to Table 5
    Windows(reportfn).Activate
    Sheets(segname).Activate
    n = 1
    Do Until Left(Cells(n, 1), 7) = "Table 4"
        n = n + 1
    Loop
    'Cells(n + 2, 1) = CStr(Round(CDbl(begmp), 3)) & "-" & CStr(Round(CDbl(endmp), 3))       'begmp & " - " & endmp
    Cells(n + 2, 1) = pcrashes
    Cells(n + 2, 1).NumberFormat = "0.0"
    Cells(n + 2, 3) = ncrashes
    Cells(n + 2, 5) = sev5
    Cells(n + 2, 7) = sev4
    Cells(n + 2, 9) = sev3
       
    '   Produces list of possible countermeasures
    
    ' Use the array of top 8 crash factors
    ' Row counter for report
    ' row counter for crash factors
    ' For each value in the arrow
        '
        ' Find the column with the crash factors
        ' Prints all the crash factors
        
    ' after the look, run through again to check for duplicates.
    ' If there are two duplicates in a row, then remove the duplicate, recheck
    
    'j = FindColumn("NUMBER_ONE_INJURIES") + 1)
    
    Windows(reportfn).Activate
    Sheets(segname).Activate
    n = 1
    Do Until Cells(n, 1) = "[begin list here]"
        n = n + 1
    Loop
    n = n + 1
        
    For i = 1 To 8
        Sheets("Key").Activate
        j = FindColumn(factorarray(i))
        nrow = 2
        
        Do
            Sheets(segname).Cells(n, 1) = Sheets("Key").Cells(nrow, j)
            n = n + 1
            nrow = nrow + 1
        Loop While Sheets("Key").Cells(nrow, j) <> ""
    
    Next
    
    'Remove duplicate entries
    'sort
    Windows(reportfn).Activate
    Sheets(segname).Activate
    n = 1
    Do Until Cells(n, 1) = "[begin list here]"
        n = n + 1
    Loop
    n = n + 1
    nrow = n
    Do While Cells(nrow + 1, 1) <> ""
        nrow = nrow + 1
    Loop
    
    ActiveWorkbook.Worksheets(segname).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(segname).Sort.SortFields.Add Key:=Range(Cells(n, 1), Cells(nrow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets(segname).Sort
        .SetRange Range(Cells(n, 1), Cells(nrow, 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim mycount As Integer
    'step through and delete duplicates
    n = 1
    Do Until Cells(n, 1) = "[begin list here]"
        n = n + 1
        mycount = n + 1
    Loop
    Cells(n, 1) = ""
    n = n + 1
    Do While Cells(n, 1) <> ""
        If Cells(n, 1) = Cells(n + 1, 1) Then
            Rows(n + 1).Delete Shift:=xlUp
        Else
            n = n + 1
        End If
    Loop
    
Dim mycol, myrow As Integer

          mycol = 1
            Sheets("FacCrash").Activate
                Do Until Sheets("FacCrash").Cells(1, mycol) = "MANNER_COLLIDSION_ID"
                    If Sheets("FacCrash").Cells(1, mycol) <> "MANNER_COLLISION_ID" Then
                        mycol = mycol + 1
                    Else
                        Exit Do
                    End If
                Loop
          myrow = 2
              Do Until Cells(myrow, mycol) = ""
                If Cells(myrow, mycol) = 3 Then
                    Sheets(segname).Activate
            
                    'Find the last non-blank cell in column A(1)
                    myrow = Cells(Rows.count, 1).End(xlUp).Row
                    myrow = myrow + 1
                    Exit Do
                End If
                    myrow = myrow + 1
              Loop
    
    Sheets("Key").Activate
    Range("AQ2:AQ8").Copy
    Sheets(segname).Activate
    Cells(myrow, 1).Select
    ActiveSheet.Paste
    
          mycol = 1
          myrow = 2
            Sheets("FacCrash").Activate
                Do Until Sheets("FacCrash").Cells(1, mycol) = "WORK_ZONE_RELATED_YNU"
                    If Sheets("FacCrash").Cells(1, mycol) <> "WORK_ZONE_RELATED_YNU" Then
                        mycol = mycol + 1
                    Else
                        Exit Do
                    End If
                Loop

              Do Until Cells(myrow, mycol) = ""
                If Cells(myrow, mycol) = "Y" Or Cells(myrow, mycol + 1) = "Y" Then
                    Sheets(segname).Activate
            
                    'Find the last non-blank cell in column A(1)
                    myrow = Cells(Rows.count, 1).End(xlUp).Row
                    myrow = myrow + 1
                    Exit Do
                End If
                    myrow = myrow + 1
              Loop
    
    Sheets("Key").Activate
    Range("AR2:AR10").Copy
    Sheets(segname).Activate
    Cells(myrow, 1).Select
    ActiveSheet.Paste
    
    Sheets(segname).Activate
    Range("A" & mycount).Copy
    myrow = Cells(Rows.count, 1).End(xlUp).Row
    Range("A" & mycount & ":A" & myrow).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    '  Move the current Segment report, Copy of Crash Factors and Copy of Vehicle Crash data to new workbook, Saves and Closes finished Analysis Report
Line999: Sheets(Array(segname, vehc, facc)).Activate
    Sheets(Array(segname, vehc, facc)).Move
    Sheets(segname).Copy After:=ActiveWorkbook.Sheets(segname)
    ActiveSheet.Name = "Auto-populated Data"
    Sheets(segname).Name = "Full Report"
    ActiveWorkbook.SaveAs Filename:=reportfp & " " & segname & " " & "Region " & regnum & " " & Right(dataranges, 4), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close

    '  Before looping, updates message to user on the progress of report compilation
    Windows(reportfn).Activate
    Sheets("Main").Activate

    Application.ScreenUpdating = True       'Enables Screen Refreshing

    '  Updates the message with the name of the Analysis Report completed
    '  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "Report Compilation Running..." & Chr(10) & "Please do not click on other Excel Windows" & Chr(10) & "Estimated Completion Time: 2 minutes (for 20 reports)" & Chr(10) & Chr(10) & "Competed Report #" & (comborow - 1)
    
    Range("H2:L12").Value = "Code Running -- Completed: Report #" & (comborow - 1)
    
    Application.Wait (Now + TimeValue("00:00:01"))      'Waits for the textbox to refresh

    Application.ScreenUpdating = False      'Disables screen refreshing ("strobe light effect")

    '  Loops to the next segment in the "Combo" sheet
    Sheets(combosn).Activate
    comborow = comborow + 1
Loop

Sheets("Main").Activate
Application.ScreenUpdating = True       'Enables Screen Refreshing

'  Deletes message to user on progress of report compilation
Range("H2:L12").Clear

Workbooks(combofn).Close False
Workbooks(crashfn).Close False

'  Displays message box with the location of the saved Reports
MsgBox "Reports Completed." & Chr(10) & "Workbook will now close when you click OK" & Chr(10) & "See Analysis Results in the following folder folder:" & Chr(10) & reportfp, vbOKOnly

cmdline = "explorer.exe" & " " & reportfp

Shell cmdline, vbNormalFocus

'  Saves and closes the "Report Compiler (Date & Time)" workbook.
ActiveWorkbook.Save
ActiveWorkbook.Close

End Sub

Function FindColumn(cname As String) As Integer

FindColumn = 1
Do Until Cells(1, FindColumn) = cname
    FindColumn = FindColumn + 1
Loop

End Function


Function UDOTRegion(County As String) As String

Dim i As Integer

For i = 2 To 30
    If UCase(Sheets("Key").Cells(i, 11)) = UCase(County) Then
        UDOTRegion = CStr(Sheets("Key").Cells(i, 12))
        Exit For
    End If
Next i

End Function

Function FirstHarmful(n As Integer) As String

Dim i As Integer

For i = 2 To 60
    If Sheets("Key").Cells(i, 1) = n Then
        FirstHarmful = Sheets("Key").Cells(i, 2)
        Exit For
    End If
Next i

End Function

Function EventSequence(n As Integer) As String

Dim i As Integer

For i = 2 To 60
    If Sheets("Key").Cells(i, 5) = n Then
        EventSequence = Sheets("Key").Cells(i, 6)
        Exit For
    End If
Next i

End Function

Function MostHarmful(n As Integer) As String

Dim i As Integer

For i = 2 To 60
    If Sheets("Key").Cells(i, 7) = n Then
        MostHarmful = Sheets("Key").Cells(i, 8)
        Exit For
    End If
Next i

End Function

Function MannerCollision(n As Integer) As String

Dim i As Integer

For i = 2 To 15
    If Sheets("Key").Cells(i, 3) = n Then
        MannerCollision = Sheets("Key").Cells(i, 4)
        Exit For
    End If
Next i

End Function

Function VehicleManuever(n As Integer) As String

Dim i As Integer

For i = 2 To 25
    If Sheets("Key").Cells(i, 9) = n Then
        VehicleManuever = Sheets("Key").Cells(i, 10)
        Exit For
    End If
Next i

End Function


Sub FilterbyCrashID()
'The purpose of this function is to sort the large vehicle crash database to those needed for a given segment
'   Comments by Sam Mineer, Brigham Young Univeristy

Dim crashstring() As String     'An array, which uses the CrashID as the string values
Dim n As Integer                'row counter, to dimension the depth of the array of strings
Dim nv As Double                'row counter, to determine the max row of vehicle data
Dim jv As Double                'column counter, to determine the max column of vehicle data
Dim icrash As Double            'column counter, to find the column with "Crash_ID", if not the first row

Sheets("FilterByCrash").Activate         '<--- Verify this is correct, that sheet exists in crashdata
n = 1       'determines the number of crashID's to put into filter, dimenstions the depth of the array
Do While Cells(n, 1) <> ""
    n = n + 1
Loop

ReDim crashstring(0 To n) As String

'With the string array correctly dimensioned, now the values in the string array can be defined
n = 1
Do While Cells(n, 1) <> ""
    crashstring(n - 1) = CStr(Cells(n, 1))
    n = n + 1
Loop

'return to "vehicle" sheet, of Crash Database
Sheets("vehicle").Activate                   '<--- Verify this is correct, that sheet exists in crashdata

'Finds the extends of crash data, for the filter to be applied
nv = 1
Do Until Cells(nv, 1) = ""
    nv = nv + 1
Loop
jv = 1
Do Until Cells(1, jv) = ""
    jv = jv + 1
Loop

'identifies the column with the CrashID
icrash = 1
Do Until LCase(Left(Cells(1, icrash), 5)) = "crash"
    icrash = icrash + 1
Loop

'Turns off the filter, reapplies the filter with
ActiveSheet.Range(Cells(1, 1), Cells(nv, jv)).AutoFilter Field:=icrash
Application.Wait (Now + TimeValue("00:00:01"))
ActiveSheet.Range(Cells(1, 1), Cells(nv, jv)).AutoFilter Field:=icrash, Criteria1:=Array(crashstring), Operator:=xlFilterValues

    
End Sub
