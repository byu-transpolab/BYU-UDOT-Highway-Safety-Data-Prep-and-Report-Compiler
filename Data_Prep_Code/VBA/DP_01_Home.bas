Attribute VB_Name = "DP_01_Home"
Option Explicit

'PMDP_01_Home Module:
' Contains procedures for the following:
'   (1) Opening and copying data into the workbook
'   (2) Checking headers of incoming data to make sure the code knows which columns are which.
'   (3) Resetting the "Home" sheet interface of the workbook.

'Declare public variables
Public DYear1 As Integer            'last year of requested AADT data
Public FileName1 As String          'working name of data sheet
Public AADTCheck As Boolean         'boolean variable for formatting AADT year data
Public CurrentWkbk As String        'String variable that holds current workbook name for reference purposes
Sub OpenCopy2(Data As Integer, FileName As String)
'OpenCopy macro:
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim intChoice As Integer            'integer variable representing the choice in FileDialogOpen selection
Dim BtnChoice As Variant            'button choice variable
Dim strpath, locwb As String        'string variables for saving a string path and the local workbook name
Dim InputEntry As String            'the value of the year entered for the AADT data
Dim row1, col1 As Integer           'row and column counters
Dim DYear2, DYear3 As Integer       'temporary variables to hold year value "XXXX"
Dim DataType As String              'variable to hold name of data type

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'Screen Updating ON
Application.ScreenUpdating = True

'(NOTE: AR4:AR13 on the OtherData sheet keep track of which data sets have been copied to the workbook.)

Workbooks.Open (FileName)
    
    Rows("1:2").Delete shift:=xlUp
    Range("H5").Copy                                                                                'FLAGGED: This seems like it will crash
    Columns("A:G").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Rows("574:574").Delete shift:=xlUp

'Assign filename to variables
locwb = ActiveWorkbook.Name
FileName1 = locwb
 
Line5:                  'Code returns to Line5 if the entered AADT year range is not available
'Activate data prep workbook
Workbooks(CurrentWkbk).Activate

'Screen Updating OFF
Application.ScreenUpdating = False

'Run Check_Headers macro (using Data value depending on working dataset):
'   (1) If Data = 1, clean up AADT data based on selected year range.
'   (2) Correct headers if the incoming headers do not match what is expected.
DP_Check_Headers (Data)

'Check AADTCheck boolean which, if true, means that the entered AADT year range is not available.
If AADTCheck = True Then
    GoTo Line5
End If

'Run CopyDataSets macro (using Data value depending on working dataset):
'   (1) Copy dataset into workbook and begin formatting.
'   (2) Rename column headers to final headers as listed on the OtherData sheet.
'   (3) Rearrange columns in order of how they are listed on the OtherData sheet.
'   (4) Delete any unneeded columns.
'   (5) If Data = 8, ask user to add more rollup data if wanted.
CopyDataSets (Data)

'Activate Home sheet
Sheets("Home").Activate

End Sub

Sub OpenCopy1(Data As Integer, FileName As String)
'OpenCopy macro:
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim intChoice As Integer            'integer variable representing the choice in FileDialogOpen selection
Dim BtnChoice As Variant            'button choice variable
Dim strpath, locwb As String        'string variables for saving a string path and the local workbook name
Dim InputEntry As String            'the value of the year entered for the AADT data
Dim row1, col1 As Integer           'row and column counters
Dim DYear2, DYear3 As Integer       'temporary variables to hold year value "XXXX"
Dim DataType As String              'variable to hold name of data type

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'Screen Updating ON
Application.ScreenUpdating = True

'(NOTE: AR4:AR13 on the OtherData sheet keep track of which data sets have been copied to the workbook.)

Workbooks.Open (FileName)

'Assign filename to variables
locwb = ActiveWorkbook.Name
FileName1 = locwb

Line5:                  'Code returns to Line5 if the entered AADT year range is not available
'Activate data prep workbook
Workbooks(CurrentWkbk).Activate

'Screen Updating OFF
Application.ScreenUpdating = False

'Run Check_Headers macro (using Data value depending on working dataset):
'   (1) If Data = 1, clean up AADT data based on selected year range.
'   (2) Correct headers if the incoming headers do not match what is expected.
DP_Check_Headers (Data)

'Check AADTCheck boolean which, if true, means that the entered AADT year range is not available.
If AADTCheck = True Then
    GoTo Line5
End If

'Run CopyDataSets macro (using Data value depending on working dataset):
'   (1) Copy dataset into workbook and begin formatting.
'   (2) Rename column headers to final headers as listed on the OtherData sheet.
'   (3) Rearrange columns in order of how they are listed on the OtherData sheet.
'   (4) Delete any unneeded columns.
'   (5) If Data = 8, ask user to add more rollup data if wanted.
CopyDataSets (Data)

'Activate Home sheet
Sheets("Home").Activate

End Sub

Sub OpenCopy(Data As Integer, FileName As String)
'OpenCopy macro:
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim intChoice As Integer            'integer variable representing the choice in FileDialogOpen selection
Dim BtnChoice As Variant            'button choice variable
Dim strpath, locwb As String        'string variables for saving a string path and the local workbook name
Dim InputEntry As String            'the value of the year entered for the AADT data
Dim row1, col1 As Integer           'row and column counters
Dim DYear2, DYear3 As Integer       'temporary variables to hold year value "XXXX"
Dim DataType As String              'variable to hold name of data type

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'Screen Updating ON
Application.ScreenUpdating = True

'(NOTE: AR4:AR13 on the OtherData sheet keep track of which data sets have been copied to the workbook.)

Workbooks.Open (FileName)

'Assign filename to variables
locwb = ActiveWorkbook.Name
FileName1 = locwb
 
Line5:                  'Code returns to Line5 if the entered AADT year range is not available
'Activate data prep workbook
Workbooks(CurrentWkbk).Activate

'Screen Updating OFF
Application.ScreenUpdating = False

'Run Check_Headers macro (using Data value depending on working dataset):
'   (1) If Data = 1, clean up AADT data based on selected year range.
'   (2) Correct headers if the incoming headers do not match what is expected.
DP_Check_Headers (Data)

'Check AADTCheck boolean which, if true, means that the entered AADT year range is not available.
If AADTCheck = True Then
    GoTo Line5
End If

'Run CopyDataSets macro (using Data value depending on working dataset):
'   (1) Copy dataset into workbook and begin formatting.
'   (2) Rename column headers to final headers as listed on the OtherData sheet.
'   (3) Rearrange columns in order of how they are listed on the OtherData sheet.
'   (4) Delete any unneeded columns.
'   (5) If Data = 8, ask user to add more rollup data if wanted.
CopyDataSets (Data)

'Activate Progress sheet
Sheets("Progress").Activate

End Sub
Sub DP_Check_Headers(Data As Integer)
' Check_Headers Macro
'   (1) If Data = 1, clean up AADT data based on selected year range.
'   (2) Compare incoming column data headers to what is expected. If there is a mismatch
'       the correct headers window opens so that the user can select the header that corresponds
'       to a specific data type. This is needed so that the macros know which columns are which
'       when rearranging the data.
'   (3) If Data = 8, ask user to add more rollup data if wanted.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016
' Unnecessary code removed by: Samuel Runyan, BYU, 2021

'Declare variables
Dim row1, row2, row3, col1, col2 As Integer             'Row and column counters
Dim DupCount As Integer                                 'Counter for multiple similar headers (i.e. AADT)
Dim HeadCount As Integer                                'Number of headers counter
Dim FirstRow As Integer                                 'Row number used in AADT simplification
Dim header, Header2 As String                           'Header values used in AADT simplification
Dim FY1, FY2, FY3, FY4, FY5, FY6, FY7 As Boolean        'Boolean variables for AADT simplification

'Clear temporary headings list. Clear previous heading data from check headers info
Workbooks(CurrentWkbk).Activate
Sheets("OtherData").Activate
Range("AO4:AO100").ClearContents
col2 = 41
Range("AX4:BB4").ClearContents
Range("AY5").ClearContents

'Copy and paste headings from data file to temporary headings list
Windows(FileName1).Activate
Rows("1:1").SpecialCells(xlCellTypeConstants, 23).Copy
Workbooks(CurrentWkbk).Activate
Sheets("OtherData").Activate
Range("AO4").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
Application.CutCopyMode = False

'Count how many Headers there are
row1 = 4
HeadCount = 0
Do While Worksheets("OtherData").Cells(row1, col2) <> ""
    HeadCount = HeadCount + 1
    row1 = row1 + 1
Loop
Worksheets("OtherData").Range("BA4") = HeadCount    'Set BA4 of OtherData sheet to HeadCount for later use

'Assign column values based on which file name is selected. Column numbers are of expected headers.
If Data = 1 Then
    col1 = 1
ElseIf Data = 2 Then
    col1 = 5
ElseIf Data = 3 Then
    col1 = 9
ElseIf Data = 4 Then
    col1 = 13
ElseIf Data = 5 Then
    col1 = 17
ElseIf Data = 6 Then
    col1 = 25
ElseIf Data = 7 Then
    col1 = 29
ElseIf Data = 8 Then
    col1 = 33
ElseIf Data = 9 Then
    col1 = 37
ElseIf Data = 10 Then
    col1 = 21
ElseIf Data = 11 Then       'Intersections
    col1 = 69
ElseIf Data = 12 Then
    col1 = 73
ElseIf Data = 13 Then
    col1 = 69
End If

'Check Header values
row1 = 4                                                'Expected header list row
Worksheets("OtherData").Range("AZ4") = col1             'Set AZ4 of OtherData to column number to be used later
'Cycle through expected headers to see if headers are found in incoming header list
Do While Worksheets("OtherData").Cells(row1, col1) <> ""
    header = Worksheets("OtherData").Cells(row1, col1)          'Expected header
    If Left(header, 4) = "AADT" Then
        header = "AADT"
    End If
    row2 = 4                                                    'Incoming header list row
    'Cycle through incoming headers to see if expected header is found
    If header = "AADT" Then
        Do While Worksheets("OtherData").Cells(row2, col2) <> ""
            'If header equals expected header, then exit loop so that row2 is left on found header row
            If Left(Worksheets("OtherData").Cells(row2, col2), Len(header)) = Left(header, Len(header)) Then
                Exit Do
            End If
            row2 = row2 + 1
        Loop
    Else
        Do While Worksheets("OtherData").Cells(row2, col2) <> ""
            'If header equals expected header, then exit loop so that row2 is left on found header row
            If Worksheets("OtherData").Cells(row2, col2) = header Then
                Exit Do
            End If
            row2 = row2 + 1
        Loop
    End If
    'If the cell is blank, it means the Header was not found and the counter is at the end of the
    'headers list. Open form to choose new Header.
    If Worksheets("OtherData").Cells(row2, col2) = "" Then
        'Assign values to be used in correct header process
        Worksheets("OtherData").Range("AX4") = Worksheets("OtherData").Cells(row1, col1)        'Expected header
        Worksheets("OtherData").Range("AY4") = Worksheets("OtherData").Cells(row1, col1 + 1)    'Description
        Worksheets("OtherData").Range("BB4") = Worksheets("OtherData").Cells(row1, col1 + 3)    'Necessary?
        Worksheets("OtherData").Range("AY5") = FileName1                                    'Working dataset filename
        Worksheets("Home").Activate
        
        'Screen Updating ON
        Application.ScreenUpdating = True
        
        'Show frmCorrectHeaders:
        '   (1) Asks user to choose header from incoming list that corresponds best to expected header.
        frmCorrectHeaders.Show
        
        'Screen Updating OFF
        Application.ScreenUpdating = False
        
        'Change filename to data prep workbook to prevent future crashes
        'If a filename of an unopened workbook is left in cell, the next dataset _
        that runs will return an error.
        Worksheets("OtherData").Range("AY5") = CurrentWkbk
    End If
    row1 = row1 + 1                     'Add to row1 to check next expected header
Loop

'Find Crash Rollup column headers that are not critical and ask user if they would like to add them to the analysis
If Data = 8 Then
    Worksheets("OtherData").Range("AP4:AP100").ClearContents        'Clear previous rollup headers
    row2 = 4                'Incoming headers row
    row3 = 4                'Rollup data headers row
    'Cycle through incoming headers to see which headers are critical
    Do While Worksheets("OtherData").Cells(row2, col2) <> ""        'Cycle through incoming headers
        row1 = 4
        'Cycle through expected headers to see if incoming header is found
        Do While Worksheets("OtherData").Cells(row1, col1) <> ""
            'If header equals expected header, then exit loop so that row1 is left on found header row
            If Worksheets("OtherData").Cells(row1, col1) = Worksheets("OtherData").Cells(row2, col2) Then
                Exit Do
            End If
            row1 = row1 + 1
        Loop
        'If the cell is blank, it means the Header was not found and the counter is at the end of the
        'headers list and that it is not critical. Header is added to non-critical list.
        If Worksheets("OtherData").Cells(row1, col1) = "" Then
            Worksheets("OtherData").Cells(row3, col2 + 1) = Worksheets("OtherData").Cells(row2, col2)
            row3 = row3 + 1
        End If
        row2 = row2 + 1
    Loop
    
    'Screen Updating ON
    Sheets("Progress").Activate
    Application.ScreenUpdating = True
    
    'Show frmAddHeaders:
    '   (1) Asks user if more rollup data headers should be considered in the analysis.
    frmAddHeaders.Show
    
    'Screen Updating OFF
    Application.ScreenUpdating = False
    Workbooks(CurrentWkbk).Activate
    Sheets("OtherData").Activate
End If

Worksheets("Progress").Activate
    
End Sub

Sub CopyDataSets(Data As Integer)
'CopyDataSets macro:
'   (1) Copy dataset into workbook and begin formatting.
'   (2) Rename column headers to final headers as listed on the OtherData sheet.
'   (3) Rearrange columns in order of how they are listed on the OtherData sheet.
'   (4) Delete any unneeded columns.
'   (5) If Data = 9, format vehicle data.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim col1, col2, col3, intCol As Long    'column number
Dim row1, row2 As Long                  'row1 number
Dim SheetName As String                 'stores working sheet name
Dim HeaderFound As Boolean              'true or false if header was found in search
Dim i, j, k As Integer                  'counters for rearranging columns
Dim RCol(1 To 100) As String            'array used to assign order of columns
Dim AADTCount As Integer                'count of AADT columns
Dim AADTYear As Long                    'year of AADT data
Dim wksht As Worksheet
   
'Depending on data number (based on working dataset) copy data into workbook
If Data = 1 Then                                    'AADT
    'Add AADT worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "AADT"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy AADT data and close workbook
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 2 Then                                'Functional Class
    'Add Functional Class worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "Functional_Class"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy Functional Class data and close workbook
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 3 Then                                'Speed Limit
    'The following variables and code are to delete un-needed sign faces info to speed up process
    'Define variables
    Dim Beg_MP, End_MP, Route, Direction, Speed_Limit As String
    
    'Assign Header values
    Beg_MP = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("I4")
    End_MP = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("I5")
    Route = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("I6")
    Direction = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("I7")
    Speed_Limit = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("I8")
    
    'Add speed limit worksheet
    Workbooks(CurrentWkbk).Activate
    SheetName = "Speed_Limit"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName
    
    'Delete any columns that are not needed
    Workbooks(FileName1).Activate
    col1 = 1 'Reset column count
    Do
        If Cells(1, col1) = Beg_MP Or Cells(1, col1) = End_MP Or Cells(1, col1) = Route Or _
        Cells(1, col1) = Direction Or Cells(1, col1) = Speed_Limit Then
            col1 = col1 + 1
        Else
            Columns(col1).EntireColumn.Delete
        End If
    Loop While Cells(1, col1) <> ""
    
    'Find data extent
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy Speed Limit data and close workbook
    Workbooks(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 4 Then                                'Lanes
    'Add Lanes worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "Thru_Lanes"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy Sign Faces data and close workbook
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 5 Then                                'Urban Code
    'Add Functional Class worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "Urban_Code"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy Sign Faces data and close workbook
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 6 Then                                'Crash Location
    'Add location sheet
    Workbooks(CurrentWkbk).Activate
    SheetName = "Location"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Worksheets.Add(Before:=Worksheets("Non-State Routes")).Name = SheetName
    Worksheets(SheetName).Tab.ColorIndex = 9
    
    'Paste location data to master workbook
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Sheets(SheetName).Activate
    ActiveSheet.Rows(1).Copy Destination:=Sheets("Non-State Routes").Rows(1)
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 7 Then                                'Crash Data (general)
    'Add Crash data sheet
    Workbooks(CurrentWkbk).Activate
    SheetName = "Crash"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Worksheets.Add(Before:=Worksheets("Non-State Routes")).Name = SheetName
    Worksheets(SheetName).Tab.ColorIndex = 9
    
    'Paste data
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Sheets(SheetName).Activate
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 8 Then                                'Crash Rollup
    'Add Rollup sheet
    Workbooks(CurrentWkbk).Activate
    SheetName = "Rollup"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Worksheets.Add(Before:=Worksheets("Non-State Routes")).Name = SheetName
    Worksheets(SheetName).Tab.ColorIndex = 9
    
    'Paste data
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Sheets(SheetName).Activate
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CRUCIAL STEP FOR THE ISAM
    'Deletes rows that don't fit the criteria for ISAM
    If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
        Call FastWB
        'Find the INTERSECTION_RELATED or Intersection Involved column
        intCol = 1
        Do Until replace(LCase(Cells(1, intCol)), " ", "_") = "intersection_related" Or replace(LCase(Cells(1, intCol)), " ", "_") = "intersection_involved"
            intCol = intCol + 1
        Loop
        'Delete non intersection rows
        Call Delete_Bad_Rows(intCol)
        Call FastWB
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CRUCIAL STEP FOR THE ISAM
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 9 Then                                'Crash Vehicle
    'Add Vehicle sheet
    Workbooks(CurrentWkbk).Activate
    SheetName = "Vehicle"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Worksheets.Add(Before:=Worksheets("Non-State Routes")).Name = SheetName
    Worksheets(SheetName).Tab.ColorIndex = 9
    
    'Paste copied data
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Sheets(SheetName).Activate
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 10 Then                                'Sign Faces -> Speed Limit
    'The following variables and code are to delete un-needed sign faces info to speed up process
    'Define variables
    Dim ROUTE_NAME, START_ACCUM, COLLECTED_DATE, LEGEND, MUTCD, ROUTE_DIR As String
    
    'Assign Header values
    START_ACCUM = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("U6")
    COLLECTED_DATE = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("U8")
    ROUTE_NAME = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("U4")
    LEGEND = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("U7")
    MUTCD = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("U9")
    ROUTE_DIR = Workbooks(CurrentWkbk).Worksheets("OtherData").Range("U5")
    
    'Add speed limit worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "Speed_Limit"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName
    
    'Find MUTCD column
    Workbooks(FileName1).Activate
    col1 = 1
    Do While Cells(1, col1) <> ""
        If Cells(1, col1) = MUTCD Then
            col2 = col1
        End If
        col1 = col1 + 1
    Loop
    
    'Sort out non-speed limit signs and copy to "copy" sheet
    ActiveSheet.Range("$A$1:$X$200000").AutoFilter Field:=col2, Criteria1:="R2-1"                   'FLAGGED: I think this range could be problematic if the data changes.
    Sheets.Add.Name = "Copy"
    Sheets(2).Activate
    ActiveSheet.Cells.Copy Destination:=Sheets("Copy").Range("A1")
    Application.CutCopyMode = False
    Sheets("Copy").Activate
    
    'Delete any columns that are not needed
    col1 = 1 'Reset column count
    Do
        If Cells(1, col1) = START_ACCUM Or Cells(1, col1) = COLLECTED_DATE Or _
        Cells(1, col1) = ROUTE_NAME Or Cells(1, col1) = LEGEND Or Cells(1, col1) = ROUTE_DIR Then
            col1 = col1 + 1
        Else
            Columns(col1).EntireColumn.Delete
        End If
    Loop While Cells(1, col1) <> ""
    
    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy Sign Faces data and close workbook
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 11 Then                                'ISAM Intersections
    'Add Intersections worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "Intersections"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy intersections data and close workbook
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
    
    'calls the IntTypeParing code and keeps only the intersection types specified by the user in the CrateIntData form
    Call IntTypesParing
    
    'Activate Progress sheet
    Worksheets("Progress").Activate
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 12 Then                                'Pavement messages
    'Add Intersections worksheet
    Windows(CurrentWkbk).Activate
    SheetName = "Pavement_Messages"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    Sheets.Add(Worksheets("Non-State Routes")).Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy intersections data and close workbook
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Windows(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Data = 13 Then                                'CAMS Intersections
    'Add Intersections worksheet
    Workbooks(CurrentWkbk).Activate
    SheetName = "Intersections"
    
    For Each wksht In Worksheets
        If wksht.Name = SheetName Then
            Application.DisplayAlerts = False
            wksht.Delete
            Application.DisplayAlerts = True
        End If
    Next

    ActiveWorkbook.Sheets.Add After:=Worksheets("Vehicle")
    ActiveWorkbook.ActiveSheet.Name = SheetName

    'Find data extent
    Workbooks(FileName1).Activate
    col1 = 1
    Do Until Cells(2, col1) <> ""            'Find first 2nd row with a value
        col1 = col1 + 1
    Loop
    row1 = Cells(1, col1).End(xlDown).row
    col1 = Range("A1").End(xlToRight).Column
    
    'Copy intersections data and close workbook
    Workbooks(FileName1).Sheets(1).Cells.Copy Destination:=Workbooks(CurrentWkbk).Sheets(SheetName).Cells
    Workbooks(CurrentWkbk).Activate
    Sheets(SheetName).Activate
    
    'Closes the workbook after the data is copied and sorted
    Workbooks(FileName1).Close False

End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------

'Assign column values based on which file name is selected.
'Column number is column on hidden OtherData sheet where the expected column headers for each dataset are found.
If Data = 1 Then
    col1 = 1
ElseIf Data = 2 Then
    col1 = 5
ElseIf Data = 3 Then
    col1 = 9
ElseIf Data = 4 Then
    col1 = 13
ElseIf Data = 5 Then
    col1 = 17
ElseIf Data = 6 Then
    col1 = 25
ElseIf Data = 7 Then
    col1 = 29
ElseIf Data = 8 Then
    col1 = 33
ElseIf Data = 9 Then
    col1 = 37
ElseIf Data = 10 Then
    col1 = 21
ElseIf Data = 11 Then
    col1 = 69
ElseIf Data = 12 Then
    col1 = 73
ElseIf Data = 13 Then
    col1 = 69
End If

'Rename column headers to match the final desired header
col2 = 1
Do While Worksheets(SheetName).Cells(1, col2) <> ""
    row2 = 4
    Do While Worksheets("OtherData").Cells(row2, col1) <> ""
        If Worksheets(SheetName).Cells(1, col2) = Worksheets("OtherData").Cells(row2, col1) Then
            Worksheets(SheetName).Cells(1, col2) = Worksheets("OtherData").Cells(row2, col1 + 2)
            Exit Do
        ElseIf Left(Worksheets(SheetName).Cells(1, col2), 4) = "AADT" Then
            AADTYear = Right(Worksheets(SheetName).Cells(1, col2), 4)
            Worksheets(SheetName).Cells(1, col2) = "AADT_" & CStr(AADTYear)
            Exit Do
        End If
        row2 = row2 + 1
    Loop
    col2 = col2 + 1
Loop

'Sort crash vehicle data
If Data = 9 Then
    Sort_Vehicles (SheetName)
End If

'Delete any columns that are not needed
col2 = 1
AADTCount = 0
DYear1 = 0
Sheets(SheetName).Activate
Do While Worksheets(SheetName).Cells(1, col2) <> ""
    HeaderFound = False
    row2 = 4
    Do While Worksheets("OtherData").Cells(row2, col1) <> ""
        If Left(Worksheets(SheetName).Cells(1, col2), 5) = "AADT_" Then       'Unique to AADT data macros
            HeaderFound = True
            AADTCount = AADTCount + 1
            If Right(Worksheets(SheetName).Cells(1, col2), 4) > DYear1 Then
                DYear1 = Right(Worksheets(SheetName).Cells(1, col2), 4)
            End If
            Exit Do
        ElseIf Worksheets("OtherData").Cells(row2, col1 + 2) = Worksheets(SheetName).Cells(1, col2) And _
        Worksheets("OtherData").Cells(row2, col1 + 3) <> "NOT USED" Then
            HeaderFound = True
            Exit Do
        End If
        row2 = row2 + 1
    Loop
    
    If HeaderFound = False Then
        Worksheets(SheetName).Columns(col2).EntireColumn.Delete
    Else
        col2 = col2 + 1
    End If
Loop

'Sort Crash Data
If Data = 6 Or Data = 7 Or Data = 8 Then
    Sort_Crashes (SheetName)
End If

'Cleans the crash data if it came from Numetric by turning strings into numeric codes.
If Data = 6 Or Data = 7 Or Data = 8 Or Data = 9 Then
    cleanNumetric (SheetName)
End If

'Assign column order for rearranging the columns to RCol() array. Order is based on OtherData sheet order
row2 = 4
i = 0
Do While Worksheets("OtherData").Cells(row2, col1) <> ""
    'If the header is AADT, then assign header names for each AADT year of data
    If row2 = 8 And col1 = 1 Then                                                                   'FLAGGED: It may be better if this is outside the loop but I don't quite understand what it does
        For j = 0 To AADTCount - 1
            i = i + 1
            RCol(i) = "AADT_" & CStr(DYear1 - j)
        Next j
    'If necessary criteria isn't NOT USED, then assign the header name to the next array value
    ElseIf Worksheets("OtherData").Cells(row2, col1 + 3) <> "NOT USED" Then
        i = i + 1
        RCol(i) = Worksheets("OtherData").Cells(row2, col1 + 2)
    End If
    row2 = row2 + 1
Loop

'Rearrange columns based on RCol() order                                                            FLAGGED: This code could be optimized better.
For k = 1 To i
    col1 = 1
    Do While Worksheets(SheetName).Cells(1, col1) <> ""
        If Worksheets(SheetName).Cells(1, col1) = RCol(k) Then
            Exit Do
        End If
        col1 = col1 + 1
    Loop
    If Worksheets(SheetName).Cells(1, col1) <> Worksheets(SheetName).Cells(1, k) Then
        Worksheets(SheetName).Columns(col1).Cut
        Worksheets(SheetName).Columns(k).Insert shift:=xlToRight
    End If
Next k
Application.CutCopyMode = False
'If there are empty columns, give error message
col1 = 1
For k = 1 To i
    If Worksheets(SheetName).Cells(1, col1) = "" Then
        Worksheets("Home").Activate
        MsgBox "While reordering the headers one or more columns were accidentally deleted or shifted leaving an empty column. We are working on fixing this problem, but in the meantime please make sure the order on your input file and the the order on the OtherData sheet match" & _
        vbCrLf & "Please reorder columns.", , "Reorder columns before continuing and run RGUI from the beginning."
        End
    End If
    col1 = col1 + 1
Next k

'If working dataset is Crash Vehicle, run the Vehicle Data Prep macro to format vehicle data
If Data = 9 Then
    'Call Vehicle_Data_Prep macro (PMDP_07_Crash):
    '   (1) Copy the vehicle data to a new sheet and prepare it for future analysis.
    Call Vehicle_Data_Prep
End If

End Sub
Public Sub FastWB(Optional ByVal opt As Boolean = True)
    With Application
        .Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not opt
        .DisplayStatusBar = Not opt
        .EnableAnimations = Not opt
        .EnableEvents = Not opt
        .ScreenUpdating = Not opt
    End With
    FastWS , opt
End Sub

Public Sub FastWS(Optional ByVal ws As Worksheet = Nothing, _
                  Optional ByVal opt As Boolean = True)
    If ws Is Nothing Then
        For Each ws In Application.ActiveWorkbook.Sheets
            EnableWS ws, opt
        Next
    Else
        EnableWS ws, opt
    End If
End Sub

Private Sub EnableWS(ByVal ws As Worksheet, ByVal opt As Boolean)
    With ws
        .DisplayPageBreaks = False
        .EnableCalculation = Not opt
        .EnableFormatConditionsCalculation = Not opt
        .EnablePivotTable = Not opt
    End With
End Sub

Public Sub IntTypesParing()
'Written by Camille Lunt on May 14, 2019
'Takes the inputs given by the user in the CreateIntData User Form
'and uses them to delete the intersections that won't be analyzed

Dim colSRSR As Integer
Dim colTraffCont As Integer
Dim colRoute0 As Integer
Dim colRoute1 As Integer
Dim colRoute2 As Integer
Dim colRoute3 As Integer
Dim colRoute4 As Integer
Dim lastcol As Integer
Dim myrow As Integer

Call Delete_Groups

'finds important column numbers

colSRSR = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colSRSR) = "SR_SR"
    colSRSR = colSRSR + 1
Loop

colTraffCont = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colTraffCont) = "TRAFFIC_CO"
    colTraffCont = colTraffCont + 1
Loop

colRoute0 = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colRoute0) = "ROUTE"
    colRoute0 = colRoute0 + 1
Loop

colRoute1 = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colRoute1) = "INT_RT_1"
    colRoute1 = colRoute1 + 1
Loop

colRoute2 = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colRoute2) = "INT_RT_2"
    colRoute2 = colRoute2 + 1
Loop

colRoute3 = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colRoute3) = "INT_RT_3"
    colRoute3 = colRoute3 + 1
Loop

colRoute4 = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, colRoute4) = "INT_RT_4"
    colRoute4 = colRoute4 + 1
Loop

lastcol = 1
Do Until ActiveWorkbook.Sheets("Intersections").Cells(1, lastcol) = ""
    lastcol = lastcol + 1
Loop

myrow = 2

'fixes formatting, fixes some typos
Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
    'corrects "Local" typos
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) = "LOCAL" Then
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) = "Local"
    End If
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) = "LOCAL" Then
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) = "Local"
    End If
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) = "LOCAL" Then
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) = "Local"
    End If
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) = "LOCAL" Then
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) = "Local"
    End If
    'makes the route numbers in number format (instead of text format)
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) <> "" And _
    ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) <> "Local" Then
        If Cells(myrow, colRoute1) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Cells(myrow, colRoute1) = "0011"          'Change route to 0011 for the purpose of this process.
        End If
        Cells(myrow, colRoute1).NumberFormat = "0000"
        Cells(myrow, colRoute1).Value = Cells(myrow, colRoute1).Value
    End If
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) <> "" And _
    ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) <> "Local" Then
        If Cells(myrow, colRoute2) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Cells(myrow, colRoute2) = "0011"          'Change route to 0011 for the purpose of this process.
        End If
        Cells(myrow, colRoute2).NumberFormat = "0000"
        Cells(myrow, colRoute2).Value = Cells(myrow, colRoute2).Value
    End If
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) <> "" And _
    ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) <> "Local" Then
        If Cells(myrow, colRoute3) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Cells(myrow, colRoute3) = "0011"          'Change route to 0011 for the purpose of this process.
        End If
        Cells(myrow, colRoute3).NumberFormat = "0000"
        Cells(myrow, colRoute3).Value = Cells(myrow, colRoute3).Value
    End If
    If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) <> "" And _
    ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) <> "Local" Then
        If Cells(myrow, colRoute4) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Cells(myrow, colRoute4) = "0011"          'Change route to 0011 for the purpose of this process.
        End If
        Cells(myrow, colRoute4).NumberFormat = "0000"
        Cells(myrow, colRoute4).Value = Cells(myrow, colRoute4).Value
    End If
    'corrects the SR_SR column
    If Left(Cells(myrow, colRoute0).Value, 4) = "089A" Then
        If ((Cells(myrow, colRoute1) <= 491 And Cells(myrow, colRoute1) <> 11 And Cells(myrow, colRoute1) <> "") Or _
        (Cells(myrow, colRoute2) <= 491 And Cells(myrow, colRoute2) <> 11 And Cells(myrow, colRoute2) <> "") Or _
        (Cells(myrow, colRoute3) <= 491 And Cells(myrow, colRoute3) <> 11 And Cells(myrow, colRoute3) <> "") Or _
        (Cells(myrow, colRoute4) <= 491 And Cells(myrow, colRoute4) <> 11 And Cells(myrow, colRoute4) <> "")) Then
            Cells(myrow, colSRSR) = "YES"
        Else
            Cells(myrow, colSRSR) = "NO"
        End If
    ElseIf Int(Left(Cells(myrow, colRoute0).Value, 4)) <= 491 And _
    ((Cells(myrow, colRoute1) <= 491 And Cells(myrow, colRoute1) <> Int(Left(Cells(myrow, colRoute0).Value, 4)) And Cells(myrow, colRoute1) <> "") Or _
    (Cells(myrow, colRoute2) <= 491 And Cells(myrow, colRoute2) <> Int(Left(Cells(myrow, colRoute0).Value, 4)) And Cells(myrow, colRoute2) <> "") Or _
    (Cells(myrow, colRoute3) <= 491 And Cells(myrow, colRoute3) <> Int(Left(Cells(myrow, colRoute0).Value, 4)) And Cells(myrow, colRoute3) <> "") Or _
    (Cells(myrow, colRoute4) <= 491 And Cells(myrow, colRoute4) <> Int(Left(Cells(myrow, colRoute0).Value, 4)) And Cells(myrow, colRoute4) <> "")) Then
        Cells(myrow, colSRSR) = "YES"
    Else
        Cells(myrow, colSRSR) = "NO"
    End If
    'gets rid of intersections with ramps (i.e. deletes intersections if they include I-15, I-70, I-80, I-84, or I-215)
    If Left(Cells(myrow, colRoute0).Value, 4) = "0015" Or Left(Cells(myrow, colRoute0).Value, 4) = "0070" Or Left(Cells(myrow, colRoute0).Value, 4) = "0080" Or Left(Cells(myrow, colRoute0).Value, 4) = "0084" Or Left(Cells(myrow, colRoute0).Value, 4) = "0215" Or _
    Cells(myrow, colRoute1).Value = 15 Or Cells(myrow, colRoute1).Value = 70 Or Cells(myrow, colRoute1).Value = 80 Or Cells(myrow, colRoute1).Value = 84 Or Cells(myrow, colRoute1).Value = 215 Or _
    Cells(myrow, colRoute2).Value = 15 Or Cells(myrow, colRoute2).Value = 70 Or Cells(myrow, colRoute2).Value = 80 Or Cells(myrow, colRoute2).Value = 84 Or Cells(myrow, colRoute2).Value = 215 Or _
    Cells(myrow, colRoute3).Value = 15 Or Cells(myrow, colRoute3).Value = 70 Or Cells(myrow, colRoute3).Value = 80 Or Cells(myrow, colRoute3).Value = 84 Or Cells(myrow, colRoute3).Value = 215 Or _
    Cells(myrow, colRoute4).Value = 15 Or Cells(myrow, colRoute4).Value = 70 Or Cells(myrow, colRoute4).Value = 80 Or Cells(myrow, colRoute4).Value = 84 Or Cells(myrow, colRoute4).Value = 215 Then
        Cells(myrow, colRoute0).EntireRow.Delete      'deletes the row if it has an interstate on it
        myrow = myrow - 1   'to account for anytime there are 2 or more in a row
    End If

myrow = myrow + 1
Loop

myrow = 2

'Option 1: SR-SR only
If ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "" Then
    ActiveWorkbook.Sheets("Intersections").Range("$A$1:$AF$10425").AutoFilter Field:=colSRSR, Criteria1:="NO"           'FLAGGED: Hard Range
    Range(Range(Cells(2, 1), Cells(2, lastcol)), Range(Cells(2, 1), Cells(2, lastcol)).End(xlDown)).EntireRow.Delete    'Deletes all rows that are not SR to SR
    ActiveWorkbook.Sheets("Intersections").Range("$A$1:$AF$10425").AutoFilter Field:=colSRSR                            'FLAGGED: Hard Range
'Option 2: SR-SR and SR-FedAid
ElseIf ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "" Then
    Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
        If Sheets("Intersections").Cells(myrow, colSRSR) = "YES" Or _
        (Sheets("Intersections").Cells(myrow, colRoute1).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute2).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute3).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute4).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "") Then
            myrow = myrow + 1
        Else
            Range(Cells(myrow, 1), Cells(myrow, lastcol)).EntireRow.Delete
        End If
    Loop
'Option 3: SR-SR and SR-FedAid and SR-Signal
ElseIf ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "YES" Then
    Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
        If Sheets("Intersections").Cells(myrow, colSRSR) = "YES" Or _
        Sheets("Intersections").Cells(myrow, colTraffCont) = "SIGNAL" Or _
        (Sheets("Intersections").Cells(myrow, colRoute1).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute2).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute3).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute4).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "") Then
            myrow = myrow + 1
        Else
            Range(Cells(myrow, 1), Cells(myrow, lastcol)).EntireRow.Delete
        End If
    Loop
'Option 4: SR-FedAid only
ElseIf ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "" Then
    Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
        If (Sheets("Intersections").Cells(myrow, colRoute1).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute2).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute3).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute4).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "") Then
            myrow = myrow + 1
        Else
            Range(Cells(myrow, 1), Cells(myrow, lastcol)).EntireRow.Delete
        End If
    Loop
'Option 5: SR-Signal only
ElseIf ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "YES" Then
    ActiveWorkbook.Sheets("Intersections").Range("$A$1:$AF$10425").AutoFilter Field:=11, Criteria1:=Array( _
        "OTHER", "STOP SIGN", "STOP SIGN - ALL WAY", "STOP SIGN - SIDE STREET", _
        "UNCONTROLLED", "YIELD SIGN", "YIELD SIGN - SIDE STREET", _
        "RAILROAD-ACTIVE CONTROL", "HAWK", "YIELD SIGN - ALL WAY", _
        "SHOULDER-MOUNT BEACON", "OVERHEAD BEACON"), Operator:=xlFilterValues
    Range(Range(Cells(2, 1), Cells(2, lastcol)), Range(Cells(2, 1), Cells(2, lastcol)).End(xlDown)).EntireRow.Delete    'Deletes all rows that are not signalized intersections
    ActiveWorkbook.Sheets("Intersections").Range("$A$1:$AF$10425").AutoFilter Field:=colSRSR                            'FLAGGED: Hard Range
'Option 6: SR-SR and SR-Signal
ElseIf ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "YES" Then
    Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
        If Sheets("Intersections").Cells(myrow, colSRSR) = "YES" Or _
        Sheets("Intersections").Cells(myrow, colTraffCont) = "SIGNAL" Then
            myrow = myrow + 1
        Else
            Range(Cells(myrow, 1), Cells(myrow, lastcol)).EntireRow.Delete
        End If
    Loop
'Option 7: SR-FedAid and SR-Signal
ElseIf ActiveWorkbook.Sheets("Inputs").Range("I13").Value = "" And _
ActiveWorkbook.Sheets("Inputs").Range("I14").Value = "YES" And _
ActiveWorkbook.Sheets("Inputs").Range("I15").Value = "YES" Then
    Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
        If Sheets("Intersections").Cells(myrow, colTraffCont) = "SIGNAL" Or _
        (Sheets("Intersections").Cells(myrow, colRoute1).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute1).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute2).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute2).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute3).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute3).Value <> "") Or _
        (Sheets("Intersections").Cells(myrow, colRoute4).Value > 491 And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "Local" And Sheets("Intersections").Cells(myrow, colRoute4).Value <> "") Then
            myrow = myrow + 1
        Else
            Range(Cells(myrow, 1), Cells(myrow, lastcol)).EntireRow.Delete
        End If
    Loop
End If

'clears any filters bar
Rows("1:1").AutoFilter


End Sub

Public Sub Delete_Bad_Rows(intCol As Long)
    'Sam Runyans comment: It seems like this macro deletes the rows that aren't intersection related using the INTERSECTION_RELATED column, but it assumes that is in column 29.
    'MODIFIED BY Samuel Runyan 8/9/21 - Changed the code so it searches for the intersection related column before performing the sort/autofilter/delete action
    
    ActiveWorkbook.Worksheets("Rollup").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rollup").Sort.SortFields.Add Key:=Cells(1, intCol), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rollup").Sort
        .SetRange Range("A2:AV370846")                                                  'FLAGGED: Hard Range
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells(1, intCol).AutoFilter
    ActiveSheet.Range("$A$1:$AV$370846").AutoFilter Field:=intCol, Criteria1:="N"       'FLAGGED: Hard Range
    Range(Range(Range("A2"), Range("A2").End(xlDown)), Range(Range("A2"), Range("A2").End(xlDown)).End(xlToRight)).EntireRow.Delete
    ActiveSheet.Range("$A$1:$AV$142075").AutoFilter Field:=intCol                       'Flagged: Hard Range
        Rows("1:1").AutoFilter
        
End Sub

Public Sub Delete_Groups()

    ActiveWorkbook.Worksheets("Intersections").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Intersections").Sort.SortFields.Add Key:=Range("D1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal                    'FLAGGED: Hard Range
    With ActiveWorkbook.Worksheets("Intersections").Sort
        .SetRange Range("A2:AV370846")                                                          'FLAGGED: Hard Range
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D1").AutoFilter
    ActiveSheet.Range("$A$1:$AV$370846").AutoFilter Field:=4, Criteria1:="2"                    'FLAGGED: Hard Range
    Range(Range(Range("A2"), Range("A2").End(xlDown)), Range(Range("A2"), Range("A2").End(xlDown)).End(xlToRight)).EntireRow.Delete
    ActiveSheet.Range("$A$1:$AV$142075").AutoFilter Field:=4                                    'FLAGGED: Hard Range
   ' Rows("1:1").Select
   ' Selection.AutoFilter
   '
   '     Range("D1").Select
   ' ActiveWorkbook.Worksheets("Intersections").Sort.SortFields.Clear
   ' ActiveWorkbook.Worksheets("Intersections").Sort.SortFields.Add Key:=Range("D1"), _
   '     SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
   ' With ActiveWorkbook.Worksheets("Intersections").Sort
   '     .SetRange Range("A2:AV370846")
    '    .Header = xlNo
    '    .MatchCase = False
   '     .Orientation = xlTopToBottom
    '    .SortMethod = xlPinYin
    '    .Apply
   ' End With
  '  Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AV$370846").AutoFilter Field:=4, Criteria1:="3"                    'FLAGGED: Hard Range
    Range(Range(Range("A2"), Range("A2").End(xlDown)), Range(Range("A2"), Range("A2").End(xlDown)).End(xlToRight)).EntireRow.Delete
    ActiveSheet.Range("$A$1:$AV$142075").AutoFilter Field:=4                                    'FLAGGED: Hard Range
        Rows("1:1").AutoFilter

    
End Sub



