Attribute VB_Name = "DP_00_SheetMacros"
'Sheet2 (Home):
'   (1) The private procedures of all "Home" sheet buttons are found in this group of code.

'Declare public variables
Public LocCrashID, LocRouDir, LocRoute, DatCrashID, RollCrashID, VehDetID, LocRampID As String
Public VehCrashID, VehVehNum, VehCraDate, VehRavDir, VehEvenSeq, VehMostHarm, VehManID As String
Public NSRCheck As Boolean
Public CurrentWkbk As String                'String variable that holds current workbook name for reference purposes

Sub CombineRoadwayButton()
'"Combine Roadway Data" button:
'   (1) Verifies that all data has been compiled as outlined.
'   (2) Combine 5 roadway datasets into a single dataset.
'   (3) Save SegRoadway_Data file with segmentation type, date, and time.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht As Worksheet                                      'Worksheet variable used to cycle through worksheets
Dim strpath, MyFileName, SegType, SaveTime As String        'String variables used to save file

'Screen Updating OFF
Application.ScreenUpdating = False

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
    
'If a previous segmented roadway data file is open, close it.
Dim wkb As Workbook
For Each wkb In Workbooks
    If Left(wkb.Name, 15) = "IntRoadway_Data" Then
        wkb.Close False                                 'Close workbook, "False" to not save file
    End If
Next wkb

'Call Run_Combined macro (PMDP_08_Combine module):
'   (1) Finishes formatting and combining the roadway data into one dataset.
'   (2) Routes are segmented to form homogeneous segments.
'   (3) Segment length at every change or at specified length given by user.
Run_Combined

'Screen Updating ON
Application.ScreenUpdating = True

'Message box asking user to select folder location where to save the segmented data
MsgBox "All roadway data has been combined. Choose where you would like to save the combined segmented data file.", , _
"Choose Folder Location"
     Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
     With fldr
         .Title = "Select a Folder"
         .AllowMultiSelect = False
         .InitialFileName = strpath
         If .Show <> -1 Then GoTo NextCode2
         sItem = .SelectedItems(1) & "\"
     End With
NextCode2:
     Set fldr = Nothing
     If Len(sItem) <= 2 Then            'If the length of the folder name is 2 or less, it means that no folder.
         Exit Sub                       'was selected and/or the user exited or cancelled. Therefore, end process.
     End If

'Screen Updating OFF
Application.ScreenUpdating = False

'Assign value to SegType based on segmentation length
If optEveryChange = True Then
    SegType = "Min-" & txtMinLength.Value & "_EC"                      'Set SegType to Min-#_EC, meaning segmentation was done at every change.
ElseIf optLength = True Then
    SegType = "Min-" & txtMinLength.Value & "_Max-" & txtSegLen.Value  'Set SegType to Min-#_Max-#.#, meaning segmentation had max length.
End If

'Assign current time to SaveTime string
If Hour(Now) < 12 Then                  'Hour of time between 12 AM and 11:59 AM
    SaveTime = Hour(Now) & "-" & Minute(Now) & "-AM"
ElseIf Hour(Now) = 12 Then              'Hour of time is 12 PM
    SaveTime = Hour(Now) & "-" & Minute(Now) & "-PM"
ElseIf Hour(Now) < 24 Then              'Hour of time between 1 PM and 11:59 PM
    SaveTime = (Hour(Now) - 12) & "-" & Minute(Now) & "-PM"
ElseIf Hour(Now) = 24 Then              'Hour of time is 12 AM
    SaveTime = (Hour(Now) - 12) & "-" & Minute(Now) & "-AM"
End If

'Assign filename and path, copy, paste, and save the segmented data as a separate CSV file.
MyFileName = "IntRdwayData_" & SegType & "_(" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & SaveTime & ").csv"
Sheets("Roadway Data").Copy
MyPath = sItem
ActiveWorkbook.SaveAs FileName:=MyPath & MyFileName, FileFormat:=xlCSV, CreateBackup:=False
DoEvents

'Close out of data file, and return to home screen of data prep workbook.
Workbooks(MyFileName).Close False
Workbooks(CurrentWkbk).Activate
Sheets("Home").Activate

'Screen Updating ON
Application.ScreenUpdating = True

'Display message box telling user that SegmentedRoadway_Data file has been saved as a CSV.
MsgBox "Finished. All data has been saved." & vbCrLf & vbCrLf & _
"The roadway data file has been saved in the format 'RdwayData_[Segment Length(s)]_([Date]_[Time]).csv'.", , "Finished"
        
End Sub

Sub CombineCrashButton()   ''''''''this is NOT the code that gets run when you click the combine crash data button. The real code is cmdCombineCrash Click event (see form_CreateIntData)
'"Combine Crash Data" button:
'   (1) Combines and formats crash data in DatabaseCleanup macro.
'   (2) Saves combined crash data sheet as a separate CSV file.
'
' Created by: Samuel Mineer and Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Screen Updating OFF
Application.ScreenUpdating = False

'Assign current workbook value
CurrentWkbk = ActiveWorkbook.Name
   
'Declare variables
Dim Start, Lap1, LapTot
Dim sec, min As Single
Dim wksht As Worksheet
Dim strpath, iFileName, MyFileName As String

'If a previous crash data file is open, close it.
Dim wkb As Workbook
For Each wkb In Workbooks
    If Left(wkb.Name, 10) = "Crash Data" Then
        wkb.Close False
    End If
Next wkb

'Call DatabaseCleanup macro (PMDP_07_Crash module):
'   (1) Crash data is combined based on Crash_ID.
'   (2) First vehicle direction is determined from vehicle crash data.
'   (3) Route numbers and direction values fixed and labels created.
DatabaseCleanup

'Delete worksheets containing crash data besides rollup now that all data is stored with the rollup data
Application.DisplayAlerts = False
For Each wksht In Worksheets
    If wksht.Name = "Location" Then
        wksht.Delete
    ElseIf wksht.Name = "Crash" Then
        wksht.Delete
    End If
Next

'Change Rollup sheet name to "Crash Data". Set color of sheet tab.
Sheets("Rollup").Name = "Crash Data"
Worksheets("Crash Data").Tab.ColorIndex = 9   ''''Camille put a break point here and it never worked

'Activate Home sheet in data prep workbook
Workbooks(CurrentWkbk).Activate
Sheets("Home").Activate

'Screen Updating ON
Application.ScreenUpdating = True

'Message box asking user to choose a folder location to save the combined crash data
Line10:
MsgBox "Choose where you would like to save the crash data file.", , "Choose Folder Location"
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder Location for Combined Crash Data"
    .AllowMultiSelect = False
    .InitialFileName = strpath
    If .Show <> -1 Then GoTo NextCode3
    sItem = .SelectedItems(1) & "\"
End With
NextCode3:
Set fldr = Nothing
If Len(sItem) <= 2 Then                     'If the length of the folder name is 2 or less, folder name is invalid.
    MsgBox "Please select a valid folder location to save the combined crash data.", , "Select Valid Folder Location"
    GoTo Line10
End If

'Screen Updating OFF
Application.ScreenUpdating = False

'Cells AU5 and AU6 keep track of the max and min data years, which are used in naming the crash data file.
'If AU5 and AU6 on the OtherData sheet are blank, then the following code runs to find out those years.
If Sheets("OtherData").Range("AU5").Value = "" And Sheets("OtherData").Range("AU6").Value = "" Then
    
    'Activate vehicle sheet
    Sheets("Vehicle").Activate
    
    'Declare variables
    Dim minyear, maxyear As Double              'Min and max year variables
    Dim numrow, idatetime As Double             'Row counter and Crash_Datetime column identifier
    
    numrow = 2                                  'Initial row value is 2 since headers are in row 1
    idatetime = 1                               'Initial column value is 1 to check each column
        
    'Identify CRASH_DATETIME column
    Do Until Cells(1, idatetime) = "CRASH_DATETIME"
        idatetime = idatetime + 1
    Loop
    
    'Go through each row, identify year, and update max and min years.
    Do Until Cells(numrow, 1) = ""
        rowYear = Year(Cells(numrow, idatetime).Value)
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
Else                        'If cells AU5 and AU6 are not empty, then assign max and min year values
    minyear = Sheets("OtherData").Range("AU5").Value
    maxyear = Sheets("OtherData").Range("AU6").Value
End If

'Assign value to MyFileName variable to save the crash data.
MyFileName = "Crash_Data_" & CStr(minyear) & "-" & CStr(maxyear) & "(" & Month(Now) & _
"-" & Day(Now) & "-" & Year(Now) & ").csv"
Sheets("Crash Data").Copy
MyPath = sItem
ActiveWorkbook.SaveAs FileName:=MyPath & MyFileName, FileFormat:=xlCSV, CreateBackup:=False
DoEvents                    'Finish saving before continuing
Workbooks(MyFileName).Close False

'Show final time
Sheets("Home").Activate
Range("O17") = Now()

'Screen Updating ON
Application.ScreenUpdating = True

'Activate Home sheet and show message box informing user that the crash data file has been saved.
Workbooks(CurrentWkbk).Activate
Sheets("Home").Activate
MsgBox "Finished. All data has been saved." & vbCrLf & vbCrLf & _
"The crash data file has been saved in the format 'Crash_Data_YEAR-YEAR_(Date).csv'.", , "Finished"

'Clear progress box
Worksheets("Home").Activate

Range("J15:O17").ClearContents
Range("J15:O17").Borders(xlDiagonalDown).LineStyle = xlNone
Range("J15:O17").Borders(xlDiagonalUp).LineStyle = xlNone
Range("J15:O17").Borders(xlEdgeLeft).LineStyle = xlNone
Range("J15:O17").Borders(xlEdgeTop).LineStyle = xlNone
Range("J15:O17").Borders(xlEdgeBottom).LineStyle = xlNone
Range("J15:O17").Borders(xlEdgeRight).LineStyle = xlNone
Range("J15:O17").Borders(xlInsideVertical).LineStyle = xlNone
Range("J15:O17").Borders(xlInsideHorizontal).LineStyle = xlNone
With Range("J15:O17").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

End Sub

Sub CrashDataButton()
'"Crash Data" Open and Copy button:
'   (1) Asks user to open the Crash Data file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht, wksht2 As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant                     'Used to assign a yes or no from message box

'If worksheet(s) already exists, user is asked if they would like to replace the data with a new file.
For Each wksht In Worksheets
    'If the "Crash" sheet exists, user is asked whether to keep it or open a new file.
    If wksht.Name = "Crash" Then
        ButtonChoice = MsgBox("Crash Data has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR11") = ""          'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblCrashData")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
        Exit For
    'If the "Crash Data" sheet exists, user is asked whether to keep it or clear all crash data.
    ElseIf wksht.Name = "Crash Data" Then
        ButtonChoice = MsgBox("There is already a Crash Data sheet with combined crash data in this workbook. " _
        & "In order to input new crash data the previous data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous crash data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all crash data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "3" represents resetting the crash data.
            Reset_Workbook (3)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Crash_Visible macro:
            '   (1) Checks to see if the 4 preliminary crash datasets have been copied.
            '   (2) If they have been copied, the Combine Crash data buttons will be shown.
            Crash_Visible
            
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "7" represents that Crash Data is being run.
OpenCopy (7)

'Activate Home sheet and change Crash Data status to green "COMPLETE"
Worksheets("Home").Activate
    With ActiveSheet.Shapes.Range(Array("lblCrashData")).ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(36, 190, 43)
        .Transparency = 0
        .Solid
    End With

Worksheets("OtherData").Range("AR11") = "READY"         'Change tracker cell AR11 to "READY"

'Run Crash_Visible macro:
'   (1) Checks to see if the 4 preliminary crash datasets have been copied.
'   (2) If they have been copied, the Combine Crash data buttons will be shown.
Crash_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box telling user the process is finished
MsgBox "Finished. Crash Data has been copied and formatted."

End Sub

Sub FunctionalClassButton()
'"Functional Class" Open and Copy button:
'   (1) Asks user to open the Functional Class file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant             'Used to assign a yes or no from message box

'Cycle through worksheets and check to see if previously-created worksheets exist in the workbook.
For Each wksht In Worksheets
    'If the "Functional Class" worksheet exists, user is asked whether to keep the old one or create a new one.
    If wksht.Name = "Functional_Class" Then
        ButtonChoice = MsgBox("Functional Class data has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR5") = ""           'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblFClass")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    'If the "Roadway Data" sheet exists, user is asked whether to keep it or clear all roadway data.
    ElseIf wksht.Name = "Roadway Data" Then
        ButtonChoice = MsgBox("There is already a Roadway Data sheet with segmented data in this workbook. " _
        & "In order to input new roadway data the previous segmented data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous segmented data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all roadway data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "2" represents resetting the roadway data.
            Reset_Workbook (2)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Run Roadway_Visible macro:
            '   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
            '   (2) If they have been copied, the Combine Roadway data buttons will be shown.
            Roadway_Visible
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "2" represents that Functional Class data is being run.
OpenCopy (2)

'Call Run_FunctionalClass macro:
'   (1)
'   (2)
Run_FunctionalClass

'Change functional class status to "COMPLETE" and turn it green
Sheets("Home").Activate
    With ActiveSheet.Shapes.Range(Array("lblFClass")).ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(36, 190, 43)
        .Transparency = 0
        .Solid
    End With
Worksheets("OtherData").Range("AR5") = "READY"              'Function class tracking status changed to "READY"

'Activate Home sheet, select cell A1 to deselect others
Sheets("Home").Activate

'Run Roadway_Visible macro:
'   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
'   (2) If they have been copied, the Combine Roadway data buttons will be shown.
Roadway_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box that the process is finished
MsgBox "Finished. Functional Class Data has been copied and formatted.", , "Functional Class Data Copied"

End Sub

End Sub

Sub LanesButton()
'"Thru Lanes" Open and Copy button:
'   (1) Asks user to open the Thru Lanes file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant             'Used to assign a yes or no from message box

'Cycle through worksheets and check to see if previously-created worksheets exist in the workbook.
For Each wksht In Worksheets
    'If the "Thru Lanes" worksheet exists, user is asked whether to keep the old one or create a new one.
    If wksht.Name = "Thru_Lanes" Then
        ButtonChoice = MsgBox("Thru Lanes has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR7") = ""           'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblLanes")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
         'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    'If the "Roadway Data" sheet exists, user is asked whether to keep it or clear all roadway data.
    ElseIf wksht.Name = "Roadway Data" Then
        ButtonChoice = MsgBox("There is already a Roadway Data sheet with segmented data in this workbook. " _
        & "In order to input new roadway data the previous segmented data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous segmented data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all roadway data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "2" represents resetting the roadway data.
            Reset_Workbook (2)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Run Roadway_Visible macro:
            '   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
            '   (2) If they have been copied, the Combine Roadway data buttons will be shown.
            Roadway_Visible
            
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "4" represents that Thru Lanes data is being run.
OpenCopy (4)

'Call Run_ThruLanes macro (PMDP_02_AADT module):
'   (1) Formats the Thru Lane data in preparation for the combination/segmentation process.
Run_ThruLanes

'Change Lanes status to "COMPLETE" and turn it green
Sheets("Home").Activate
With ActiveSheet.Shapes.Range(Array("lblLanes")).ShapeRange.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(36, 190, 43)
    .Transparency = 0
    .Solid
End With
Worksheets("OtherData").Range("AR7") = "READY"              'Lanes tracking status changed to "READY"

'Activate Home sheet, select cell A1 to deselect others
Sheets("Home").Activate

'Run Roadway_Visible macro:
'   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
'   (2) If they have been copied, the Combine Roadway data buttons will be shown.
Roadway_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box that process is finished.
MsgBox "Finished. Lane Data has been copied and formatted."

End Sub

Sub CrashLocationButton()
'"Crash Location" Open and Copy button:
'   (1) Asks user to open the Crash Location Data file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht, wksht2 As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant                     'Used to assign a yes or no from message box

'If worksheet(s) already exists, user is asked if they would like to replace the data with a new file.
For Each wksht In Worksheets
    If wksht.Name = "Location" Then
        ButtonChoice = MsgBox("Location Crash data has already been copied to this workbook. " & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR10") = ""          'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblLocation")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
         'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
        Exit For
    'If the "Crash Data" sheet exists, user is asked whether to keep it or clear all crash data.
    ElseIf wksht.Name = "Crash Data" Then
        ButtonChoice = MsgBox("There is already a Crash Data sheet with combined crash data in this workbook. " _
        & "In order to input new crash data the previous data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous crash data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all crash data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "3" represents resetting the crash data.
            Reset_Workbook (3)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Crash_Visible macro:
            '   (1) Checks to see if the 4 preliminary crash datasets have been copied.
            '   (2) If they have been copied, the Combine Crash data buttons will be shown.
            Crash_Visible
            
         'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "6" represents that Crash Location data is being run.
OpenCopy (6)

'Activate home sheet
Worksheets("Home").Activate

'Activate Home sheet and change Crash Location status to green "COMPLETE"
    With ActiveSheet.Shapes.Range(Array("lblLocation")).ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(36, 190, 43)
        .Transparency = 0
        .Solid
    End With
Worksheets("OtherData").Range("AR10") = "READY"         'Change tracker cell AR11 to "READY"

'Run Crash_Visible macro:
'   (1) Checks to see if the 4 preliminary crash datasets have been copied.
'   (2) If they have been copied, the Combine Crash data buttons will be shown.
Crash_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box that process is finished.
MsgBox "Finished. Crash Location data has been copied and formatted."

End Sub


Sub CrashRollupButton()
'"Crash Rollup" Open and Copy button:
'   (1) Asks user to open the Crash Rollup file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht, wksht2 As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant                     'Used to assign a yes or no from message box

'If worksheet(s) already exists, user is asked if they would like to replace the data with a new file.
For Each wksht In Worksheets
    'If the "Rollup" sheet exists, user is asked whether to keep it or open a new file.
    If wksht.Name = "Rollup" Then
        ButtonChoice = MsgBox("Crash Rollup data has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR12") = ""          'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblRollup")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
        Exit For
    'If the "Crash Data" sheet exists, user is asked whether to keep it or clear all crash data.
    ElseIf wksht.Name = "Crash Data" Then
        ButtonChoice = MsgBox("There is already a Crash Data sheet with combined crash data in this workbook. " _
        & "In order to input new crash data the previous data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous crash data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all crash data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "3" represents resetting the crash data.
            Reset_Workbook (3)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Crash_Visible macro:
            '   (1) Checks to see if the 4 preliminary crash datasets have been copied.
            '   (2) If they have been copied, the Combine Crash data buttons will be shown.
            Crash_Visible
            
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "8" represents that Crash Rollup data is being run.
OpenCopy (8)

'Activate Home sheet and change Crash Rollup status to green "COMPLETE"
Worksheets("Home").Activate
    With ActiveSheet.Shapes.Range(Array("lblRollup")).ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(36, 190, 43)
        .Transparency = 0
        .Solid
    End With
Worksheets("OtherData").Range("AR12") = "READY"

'Crash_Visible macro:
'   (1) Checks to see if the 4 preliminary crash datasets have been copied.
'   (2) If they have been copied, the Combine Crash data buttons will be shown.
Crash_Visible

'Show message box telling user the process is finished
MsgBox "Finished. Crash Rollup data has been copied and formatted."

'Screen Updating ON
Application.ScreenUpdating = True

End Sub

Sub SpeedLimitButton()
'"Speed Limit" Open and Copy button:
'   (1) Asks user to open the Speed Limit file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   Comments by Josh Gibbons, Brigham Young University, 2016

'Declare variables
Dim wksht As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant             'Used to assign a yes or no from message box

'If worksheet already exists, delete it so that new data can be run, based on the user's choice.
For Each wksht In Worksheets
    'If the "Speed_Limit" worksheet exists, user is asked whether to keep the old one or create a new one.
    If wksht.Name = "Speed_Limit" Then
        ButtonChoice = MsgBox("Speed Limit data has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR6") = ""           'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblSignFaces")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    'If the "Roadway Data" sheet exists, user is asked whether to keep it or clear all roadway data.
    ElseIf wksht.Name = "Roadway Data" Then
        ButtonChoice = MsgBox("There is already a Roadway Data sheet with segmented data in this workbook. " _
        & "In order to input new roadway data the previous segmented data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous segmented data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all roadway data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "2" represents resetting the roadway data.
            Reset_Workbook (2)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Run Roadway_Visible macro:
            '   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
            '   (2) If they have been copied, the Combine Roadway data buttons will be shown.
            Roadway_Visible
            
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "3" represents that Speed Limit data is being run.
OpenCopy (3)

'Call Run_SignFaces macro (PMDP_04_SpeedLimit module):
'   (1) Formats the speed limit data in preparation for the combination/segmentation process.
Run_SpeedLimit

'Change Speed Limit status to "COMPLETE" and turn it green
Sheets("Home").Activate
    With ActiveSheet.Shapes.Range(Array("lblSignFaces")).ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(36, 190, 43)
        .Transparency = 0
        .Solid
    End With
Worksheets("OtherData").Range("AR6") = "READY"

'Activate Home sheet, select cell A1 to deselect others
Sheets("Home").Activate

'Run Roadway_Visible macro:
'   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
'   (2) If they have been copied, the Combine Roadway data buttons will be shown.
Roadway_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box that process is finished.
MsgBox "Finished. Speed Limit Data has been copied and formatted.", , "Speed Limit Data Copied"

End Sub

Sub UrbanCodeButton()
'"Urban Code" Open and Copy button:
'   (1) Asks user to open the Urban Code file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant             'Used to assign a yes or no from message box

'If worksheet already exists, delete it so that new data can be run, based on the user's choice.
For Each wksht In Worksheets
    'If the "Urban Code" worksheet exists, user is asked whether to keep the old one or create a new one.
    If wksht.Name = "Urban_Code" Then
        ButtonChoice = MsgBox("Urban Code data has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR8") = ""           'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblUCode")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    'If the "Roadway Data" sheet exists, user is asked whether to keep it or clear all roadway data.
    ElseIf wksht.Name = "Roadway Data" Then
        ButtonChoice = MsgBox("There is already a Roadway Data sheet with segmented data in this workbook. " _
        & "In order to input new roadway data the previous segmented data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous segmented data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all roadway data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "2" represents resetting the roadway data.
            Reset_Workbook (2)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Run Roadway_Visible macro:
            '   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
            '   (2) If they have been copied, the Combine Roadway data buttons will be shown.
            Roadway_Visible
            
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "5" represents that Urban Code data is being run.
OpenCopy (5)

'Call Run_UrbanCode macro (PMDP_06_UrbanCode module):
'   (1) Formats the Urban Code data in preparation for the combination/segmentation process.
Run_UrbanCode

'Change Urban Code status to "COMPLETE" and turn it green
Sheets("Home").Activate
With ActiveSheet.Shapes.Range(Array("lblUCode")).ShapeRange.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(36, 190, 43)
    .Transparency = 0
    .Solid
End With
Worksheets("OtherData").Range("AR8") = "READY"

'Activate Home sheet, select cell A1 to deselect others
Sheets("Home").Activate

'Run Roadway_Visible macro:
'   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
'   (2) If they have been copied, the Combine Roadway data buttons will be shown.
Roadway_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box that process is finished.
MsgBox "Finished. Urban Code Data has been copied and formatted.", , "Urban Code Data Copied"

End Sub

Sub CrashVehicleButton()
'"Crash Vehicle" Open and Copy button:
'   (1) Asks user to open the Crash Vehicle data file.
'   (2) Verifies that column headers match up with what is expected.
'   (3) Copies the data into the workbook.
'   (4) Formats the data in preparation for the combining step.
'   (5) Exports sheet as a separate Excel file that will be used after running the R model(s).
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Declare variables
Dim wksht, wksht2 As Worksheet                  'Used to count through all current worksheets
Dim ButtonChoice As Variant                     'Used to assign a yes or no from message box

'If worksheet(s) already exists, user is asked if they would like to replace the data with a new file.
For Each wksht In Worksheets
    'If the "Vehicle" sheet exists, user is asked whether to keep it or open a new file.
    If wksht.Name = "Vehicle" Then
        ButtonChoice = MsgBox("Crash Vehicle data has already been copied to this workbook." & _
        "Would you like to choose a new file for this dataset?", vbYesNo, "Data Already Exists")
        'If user clicks yes, sheet is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            wksht.Delete                                        'Delete worksheet
            Application.DisplayAlerts = True                    'Show alerts and save messages
            Worksheets("OtherData").Range("AR13") = ""          'Change tracker cell to blank, meaning it's not ready
            'Change status box to blank and black
                With ActiveSheet.Shapes.Range(Array("lblVehicle")).ShapeRange.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
        Exit For
    'If the "Crash Data" sheet exists, user is asked whether to keep it or clear all crash data.
    ElseIf wksht.Name = "Crash Data" Then
        ButtonChoice = MsgBox("There is already a Crash Data sheet with combined crash data in this workbook. " _
        & "In order to input new crash data the previous data must be deleted first. " & vbCrLf & vbCrLf _
        & "Would you like to delete the previous crash data now?", vbYesNo, "Data Already Exists")
        'If user clicks yes, all crash data is deleted.
        If ButtonChoice = vbYes Then
            Application.DisplayAlerts = False                   'Do not show alerts and save messages
            
            'Call Reset_Workbook macro (PMDP_01_Home module):
            '   (1) Clears all previous data in preparation for running the data again.
            '   (2) The "3" represents resetting the crash data.
            Reset_Workbook (3)
            
            Application.DisplayAlerts = True                    'Show alerts and save messages
            
            'Crash_Visible macro:
            '   (1) Checks to see if the 4 preliminary crash datasets have been copied.
            '   (2) If they have been copied, the Combine Crash data buttons will be shown.
            Crash_Visible
            
        'If user clicks no, the open and copy process ends.
        ElseIf ButtonChoice = vbNo Then
            End                                                 'End process
        End If
    End If
Next wksht

'Screen Updating OFF
Application.ScreenUpdating = False

'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "9" represents that Crash Vehicle data is being run.
OpenCopy (9)

'Activate Home sheet and change Crash Data status to green "COMPLETE"
Worksheets("Home").Activate
    With ActiveSheet.Shapes.Range(Array("lblVehicle")).ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(36, 190, 43)
        .Transparency = 0
        .Solid
    End With
Worksheets("OtherData").Range("AR13") = "READY"         'Change tracker cell AR13 to "READY"

'Screen Updating ON
Application.ScreenUpdating = True

Application.DisplayAlerts = False                       'Do not show message alerts

'Save as an Excel document
MsgBox "Please select the folder location where you wish you save the formatted vehicle data.", , "Select Folder"
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strpath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1) & "\"
    End With
NextCode:
    Set fldr = Nothing
    If Len(sItem) <= 2 Then
        Exit Sub
    End If
    
'Screen Updating OFF
Application.ScreenUpdating = False

'If the cells that hold the max and min data years are blank, find max and min years
If Sheets("OtherData").Range("AU5").Value = "" And Sheets("OtherData").Range("AU6").Value = "" Then
    'Activate vehicle sheet
    Sheets("Vehicle").Activate
    
    'Declare variables
    Dim minyear, maxyear As Double
    Dim numrow, idatetime As Double
    numrow = 2
    idatetime = 1
        
    'Find the Crash_Datetime column
    Do Until Cells(1, idatetime) = "CRASH_DATETIME"
        idatetime = idatetime + 1
    Loop
    
    'Go through each crash and find the max and min years of all the data
    Do Until Cells(numrow, 1) = ""
        rowYear = Year(Cells(numrow, idatetime).Value)
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
'If cells that hold years are not blank, then assign max and min years
Else
    minyear = Sheets("OtherData").Range("AU5").Value
    maxyear = Sheets("OtherData").Range("AU6").Value
End If
    
'Assign filename and path variables
MyFileName = "Vehicles_" & CStr(minyear) & "-" & CStr(maxyear) & ".xlsx"
Sheets("Vehicle").Copy
MyPath = sItem

'Save file
ActiveWorkbook.SaveAs FileName:=MyPath & MyFileName, CreateBackup:=False
DoEvents

Application.DisplayAlerts = True                       'Show message alerts

'Close out of vehicles workbook and activate home sheet
ActiveWorkbook.Close False
Sheets("Home").Activate

'Crash_Visible macro:
'   (1) Checks to see if the 4 preliminary crash datasets have been copied.
'   (2) If they have been copied, the Combine Crash data buttons will be shown.
Crash_Visible

'Screen Updating ON
Application.ScreenUpdating = True

'Show message box telling user the process is finished
MsgBox "Finished. Vehicle data has been formatted, copied, and saved as a separate Excel workbook using the format 'Vehicles_YEAR-YEAR.xlsx'.", , "Vehicle Data Saved"

End Sub

Private Sub optEveryChange_Click()
'"Every Change" option button:
'   (1) "Every change" option clicked, therefore it should be selected.
'   (2) HIDES segment length text box as well as "Mile(s)" label.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'If the "Length" option button is selected, then show the length control options
If optLength = True Then
    txtSegLen.Visible = True                'Show segment length text box
    lblMiles.Visible = True                 'Show "Mile(s)" label
    optLength.Caption = "Max Length:"       'Add colon to max length option button caption
'If the "Every Change" option button is selected, then hide the length control options
Else
    txtSegLen.Visible = False               'Hide segment length text box
    lblMiles.Visible = False                'Hide "Mile(s)" label
    optLength.Caption = "Max Length"        'Remove colon from max length option button caption
End If

End Sub

Private Sub optLength_Click()
'"Length" option button:
'   (1) "Length" option clicked, therefore it should be selected.
'   (2) SHOWS segment length text box as well as "Mile(s)" label.
'
' Created by: Josh Gibbons, BYU, 2015
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'If the "Length" option button is selected, then show the length control options
If optLength = True Then
    txtSegLen.Visible = True                'Show segment length text box
    lblMiles.Visible = True                 'Show "Mile(s)" label
    optLength.Caption = "Max Length:"       'Add colon to max length option button caption
'If the "Every Change" option button is selected, then hide the length control options
Else
    txtSegLen.Visible = False               'Hide segment length text box
    lblMiles.Visible = False                'Hide "Mile(s)" label
    optLength.Caption = "Max Length"        'Remove colon from max length option button caption
End If

End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'Worksheet_BeforeDoubleClick action:
'   (1) Activates when screen is double-clicked.
'   (2) Zooms to the range of home screen buttons.
'
' Created by: Josh Gibbons, BYU, 2016
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'Activate Home sheet
Sheets("Home").Activate

'ZoomToRange macro:
'   (1) Zooms the view window to the range indicated as ZoomThisRange
'   (2) If PreserveRows is True, then window will fit all rows. If False, window will fit all columns.
ZoomToRange ZoomThisRange:=Range("A1:Q18"), PreserveRows:=False

End Sub

Sub Roadway_Visible()
'Roadway_Visible macro:
'   (1) Checks to see if the 5 preliminary roadway datasets have been copied.
'   (2) If they have been copied, the Combine Roadway data buttons will be shown.
'
' Created by: Josh Gibbons, BYU, 2016
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'(NOTE: AR4:AR8 on the OtherData sheet keep track of which roadway data sets have been copied to the workbook.)

'If the roadway data tracker cells all say "READY", then show the segmentation controls
If Sheets("OtherData").Range("AR4") = "READY" And Sheets("OtherData").Range("AR5") = "READY" And _
Sheets("OtherData").Range("AR6") = "READY" And Sheets("OtherData").Range("AR7") = "READY" And _
Sheets("OtherData").Range("AR8") = "READY" Then
    lblSegmentation.Visible = True              'Show segmentation label
    optEveryChange.Visible = True               'Show every change option button
    optLength.Visible = True                    'Show length option button
    'If the length radio button is selected, then show the length options
    If optLength = True Then
        txtSegLen.Visible = True                'Show segment length text box
        lblMiles.Visible = True                 'Show "Mile(s)" label
        optLength.Caption = "Max Length:"       'Add colon to max length option button caption
    Else
        txtSegLen.Visible = False               'Hide segment length text box
        lblMiles.Visible = False                'Hide "Mile(s)" label
        optLength.Caption = "Max Length"        'Remove colon from max length option button caption
    End If
    cmdCombineRoadway.Visible = True            'Show combine roadway button
    lblCombineRoadwayCover.Visible = False      'Hide label (cover) from comb. road. status box, revealing the box
    lblMinLength.Visible = True                 'Show Min Length label
    txtMinLength.Visible = True                 'Show Min Length text box
    lblMilesMin.Visible = True                  'Show Mile(s) min label
    
                                                '   (This is done because the actual box cannot be hidden with VBA)
'If at least one roadway data tracker cell does not say "READY" then hide segmentation controls
Else
    lblSegmentation.Visible = False             'Hide segmentation label
    optEveryChange.Visible = False              'Hide every change option button
    optLength.Visible = False                   'Hide length option button
    txtSegLen.Visible = False                   'Hide segment length text box
    lblMiles.Visible = False                    'Hide "Mile(s)" label
    cmdCombineRoadway.Visible = False           'Hide combine roadway button
    lblCombineRoadwayCover.Visible = True       'Show label (cover) on comb. road. status box, hiding the box
    lblMinLength.Visible = False                'Hide Min Length label
    txtMinLength.Visible = False                'Hide Min Length text box
    lblMilesMin.Visible = False                 'Hide Mile(s) min label
                                                '   (This is done because the actual box cannot be hidden with VBA)
End If

End Sub

Sub Crash_Visible()
'Crash_Visible macro:
'   (1) Checks to see if the 4 preliminary crash datasets have been copied.
'   (2) If they have been copied, the Combine Crash data buttons will be shown.
'
' Created by: Josh Gibbons, BYU, 2016
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

'(NOTE: AR10:AR13 on the OtherData sheet keep track of which crash data sets have been copied to the workbook.)

'If the crash data tracker cells all say "READY", then show the combine crash data controls
If Sheets("OtherData").Range("AR10") = "READY" And Sheets("OtherData").Range("AR11") = "READY" And _
Sheets("OtherData").Range("AR12") = "READY" And Sheets("OtherData").Range("AR13") = "READY" Then
    lblCombineCrashCover.Visible = False        'Hide label (cover) from comb. crash status box, revealing the box
                                                '   (This is done because the actual box cannot be hidden with VBA)
    cmdCombineCrash.Visible = True              'Show combine crash button
'If at least one crash data tracker cell does not say "READY" then hide combine crash data controls
Else
    lblCombineCrashCover.Visible = True         'Show label (cover) on comb. crash status box, hiding the box
                                                '   (This is done because the actual box cannot be hidden with VBA)
    cmdCombineCrash.Visible = False             'Hide combine crash button
End If

End Sub


