VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_CreateCAMSData 
   Caption         =   "Create CAMS Data"
   ClientHeight    =   8580.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   OleObjectBlob   =   "form_CreateCAMSData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_CreateCAMSData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_crashdata_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select General Crash Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_crashdata = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_crashdata = ""
Else
    txt_crashdata = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmd_crashlocation_Click()
' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Crash Location Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_crashlocation = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_crashlocation = ""
Else
    txt_crashlocation = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmd_crashrollup_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Crash Rollup Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_crashrollup = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_crashrollup = ""
Else
    txt_crashrollup = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmd_crashvehicle_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Crash Vehicle Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_crashvehicle = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_crashvehicle = ""
Else
    txt_crashvehicle = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmd_intersection_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Intersection Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_intersection = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_intersection = ""
Else
    txt_intersection = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmd_crashspeedlimit_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Speed Limit Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_crashspeedlimit = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_crashspeedlimit = ""
Else
    txt_crashspeedlimit = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmd_pavement_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Pavement Messages Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txt_pavement = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_pavement = ""
Else
    txt_pavement = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdAADT_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select AADT Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtAADT = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtAADT = ""
Else
    txtAADT = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdCAMSCrash_Click()

Dim guiwb, wd As String
Dim ccname As String
Dim StartTime, EndTime

wd = Sheets("Inputs").Range("M2")

'check if all data has been inputted
If txt_crashdata = "" Or txt_crashlocation = "" Or txt_crashrollup = "" Or txt_crashvehicle = "" Or txt_intersection = "" Then
    MsgBox "Select file paths for all given datasets before combining the data.", , "Select All Filepaths"
    Exit Sub
ElseIf chk_SRatSR = False And chk_SRatFA = False And chk_SRatSignal = False Then
    MsgBox "Select at least one type of intersection-related crashes to be removed.", , "Make a Selection"
    Exit Sub
ElseIf (opt_stopbar = False And opt_center = False) Or (opt_250 = False And opt_speedlimit = False) Then
    MsgBox "Make a selection in both groups of radio buttons to define the intersection functional distance.", , "Make a Selection"
    Exit Sub
ElseIf (opt_speedlimit = True And txt_crashspeedlimit = "") Then
    MsgBox "Select file paths for all given datasets before combining the data.", , "Select All Filepaths"
    Exit Sub
End If

'fill in info on the inputs sheet for future use
If chk_SRatSR = True Then Sheets("Inputs").Cells(11, 13) = "YES" Else Sheets("Inputs").Cells(11, 13) = ""
If chk_SRatFA = True Then Sheets("Inputs").Cells(12, 13) = "YES" Else Sheets("Inputs").Cells(12, 13) = ""
If chk_SRatSignal = True Then Sheets("Inputs").Cells(13, 13) = "YES" Else Sheets("Inputs").Cells(13, 13) = ""
If opt_speedlimit = True Then
    Sheets("Inputs").Cells(14, 13) = "Speed Limit"
Else
    Sheets("Inputs").Cells(14, 13) = "250ft"
End If
If opt_stopbar = True Then Sheets("Inputs").Cells(15, 13) = "Stopbar" Else Sheets("Inputs").Cells(15, 13) = "Center"

Sheets("Inputs").Cells(23, 13) = txt_crashdata
Sheets("Inputs").Cells(24, 13) = txt_crashlocation
Sheets("Inputs").Cells(25, 13) = txt_crashrollup
Sheets("Inputs").Cells(26, 13) = txt_crashvehicle
Sheets("Inputs").Cells(27, 13) = txt_intersection
Sheets("Inputs").Cells(28, 13) = txt_pavement
Sheets("Inputs").Cells(29, 13) = txt_crashspeedlimit

form_CreateCAMSData.Hide

guiwb = ActiveWorkbook.Name

StartTime = Time

'loading in all the data
'variables:
'AADT: 1
'Functional Class: 2
'Speed Limit: 3
'Lanes: 4
'Urban Code: 5
'Crash Location: 6
'Crash Data: 7
'Crash Rollup: 8
'Crash Vehicle: 9
'Sign Faces: 10
'ISAM Intersections: 11
'Pavement Messages: 12
'CAMS Intersections: 13

'Delete previous worksheets
Application.DisplayAlerts = False
For Each wksht In Worksheets
    If wksht.Name = "Crash" Then
        wksht.Delete
    ElseIf wksht.Name = "Location" Then
        wksht.Delete
    ElseIf wksht.Name = "Rollups" Then
        wksht.Delete
    ElseIf wksht.Name = "Vehicle" Then
        wksht.Delete
    End If
Next wksht
Application.DisplayAlerts = True

'Update progress screen
guiwb = ActiveWorkbook.Name
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files. Please wait."
    .Range("A3") = "Do not close Excel. Code running."
    .Range("B2") = ""
    .Range("B3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'load in crash data file
Call OpenCopy(7, txt_crashdata.Value)
Worksheets("Crash").UsedRange

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: General Crash Complete"
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'load in crash location file
Call OpenCopy(6, txt_crashlocation.Value)
Worksheets("Location").UsedRange

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: Location Complete"
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'load in crash rollup file
Call OpenCopy(8, txt_crashrollup.Value)
Worksheets("Rollup").UsedRange

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: Rollup Complete"
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'load in crash vehicle file
Call OpenCopy(9, txt_crashvehicle.Value)
Worksheets("Vehicle").UsedRange

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: Vehicle Complete"
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'load in intersections file
Call OpenCopy(13, txt_intersection.Value)
Worksheets("Intersections").UsedRange

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: Intersections Complete"
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'load in speed limit file if necessary
If opt_speedlimit = True Then
    Call OpenCopy(3, txt_crashspeedlimit.Value)
    Worksheets("Speed_Limit").UsedRange
    'cleanup the SL file
    Run_SpeedLimit
End If

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: Speed Limits Complete"
    .Range("A3") = "All Crash Data Imported. Now Combining Files."
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


'These next macros/subs are found in DP_12_CAMSCrash

    'Cleans up typos and bad route names in the Intersections file
    CleanIntersections

    'Process the intersections file
    MP_HitList

    'Run IntVehiclePrep macro
    IntVehiclePrep        'Camille: tbh I'm not entirely sure what this macro does. But I think it's important. Josh used it.
                          'Sam: It appears this macro combines the vehicles in each crash and edits the crash sequence events to reflect a single vehicle crash while storing the number of vehicles in the vehicle_num column.


'Begin to show progress box on Home sheet
'Update progress screen
guiwb = ActiveWorkbook.Name
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Combining Crash Files (Part 1 of 2). Please wait."
    .Range("A3") = "Do not close Excel. Code running."
    .Range("B2") = ""
    .Range("B3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
    .Range("B6") = ""
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


    'Combines the Location and Rollup data together and keeps only the crashes analyzed in the CAMS
    getCAMScrashes
    
    
'Update progress screen
guiwb = ActiveWorkbook.Name
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Combining Crash Files (Part 2 of 2). Please wait."
    .Range("A3") = "Do not close Excel. Code running."
    .Range("B2") = ""
    .Range("B3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
    .Range("B6") = ""
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


    'Finishes combining all of the crash data (general and vehicle)
    'saves it into a new file
    finishCAMScrashes
    

'Save combined crash data file
Sheets("Combined_Crash").Move
ccname = wd & "/CAMSCrash" & "_by" & Left(Workbooks(guiwb).Sheets("Inputs").Cells(14, 13).Value, 5) & "from" & Workbooks(guiwb).Sheets("Inputs").Cells(15, 13).Value & "_" & Left(replace(Date, "/", "-"), Len(replace(Date, "/", "-")) - 5) & "_" & Left(replace(Time, ":", "-"), Len(replace(Time, ":", "-")) - 6) & Right(replace(Time, ":", "-"), 2) & ".csv"
ccname = replace(ccname, " ", "_")
ccname = replace(ccname, "/", "\")
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs FileName:=ccname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("M6").Value = replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")
ActiveWorkbook.Close
Application.DisplayAlerts = True

EndTime = Time

Application.ScreenUpdating = True

'Enter filepath
form_camsinput.txt_camscrashfilepath.Value = Sheets("Inputs").Range("M6").Value
form_camsinput.txt_camsrdwyfilepath.Value = Sheets("Inputs").Range("M5").Value

'Message box with announcement
MsgBox "The Crash input file has been successfully created and saved in the working directory folder. The following is a summary of the process:" & Chr(10) & _
Chr(10) & _
"Process: Crash Input" & Chr(10) & _
"Start Time: " & StartTime & Chr(10) & _
"End Time: " & EndTime & Chr(10), vbOKOnly, "Process Complete"

'Reopen form
Workbooks(guiwb).Sheets("Home").Activate
form_camsinput.Show



End Sub

Private Sub cmdCAMSRoadway_Click()

Dim guiwb As String
Dim wksht As Worksheet

'Check that all text boxes have values.
If txtAADT = "" Or txtFC = "" Or txtLanes = "" Or txtSL = "" Or txtUC = "" Then
    MsgBox "Select file paths for all given roadway datasets before combining the data.", , "Select All Filepaths"
    Exit Sub
End If

'Check that valid values are entered for other values based on option button selection.
If opt_EveryChange.Value = True Then
    'Nothing
ElseIf opt_MaxLength = True Then
    If IsNumeric(txt_MaxSegLength.Value) = False Then
        MsgBox "Enter valid value for the maximum segment length before proceeding.", , "Enter Valid Max Length"
        Exit Sub
    End If
Else
    MsgBox "Select to segment by every change or a maximum length.", , "Make Selection"
    Exit Sub
End If

If IsNumeric(txt_MinSegLength.Value) = False Then
    MsgBox "Enter valid value for the minimum segment length before proceeding.", , "Enter Valid Min Length"
    Exit Sub
End If

Sheets("Inputs").Cells(17, 13) = txtAADT
Sheets("Inputs").Cells(18, 13) = txtFC
Sheets("Inputs").Cells(19, 13) = txtLanes
Sheets("Inputs").Cells(20, 13) = txtSL
Sheets("Inputs").Cells(21, 13) = txtUC

'Hide form
form_CreateCAMSData.Hide

'Delete previous worksheets
Application.DisplayAlerts = False
For Each wksht In Worksheets
    If wksht.Name = "Dataset" Then
        wksht.Delete
    ElseIf wksht.Name = "AADT" Then
        wksht.Delete
    ElseIf wksht.Name = "Functional_Class" Then
        wksht.Delete
    ElseIf wksht.Name = "Thru_Lanes" Then
        wksht.Delete
    ElseIf wksht.Name = "Speed_Limit" Then
        wksht.Delete
    ElseIf wksht.Name = "Urban_Code" Then
        wksht.Delete
    End If
Next wksht
Application.DisplayAlerts = True

'Update progress screen
guiwb = ActiveWorkbook.Name
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Roadway Files. Please wait."
    .Range("A3") = "Do not close Excel. Code running."
    .Range("B2") = ""
    .Range("B3") = ""
    .Range("A4") = "Start Time"
    .Range("A5") = ""
    .Range("A6") = ""
    .Range("B4") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

''AADT Preparation
'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "1" represents that AADT data is being run.
Call OpenCopy(1, txtAADT.Value)

'Call Run_AADT macro (PMDP_02_AADT module):
'   (1) Formats the AADT data in preparation for the combination/segmentation process.
Run_AADT

Worksheets("AADT").UsedRange

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "AADT Complete. Loading Functional Class"
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


''Functional Class Preparation
'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "2" represents that Functional Class data is being run.
Call OpenCopy2(2, txtFC.Value)

'Call Run_FunctionalClass macro:
'   (1)
Run_FunctionalClass

Worksheets("Functional_Class").UsedRange

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "FC Complete. Loading Lanes."
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


''Lanes Preparation
'Call OpenCopy macro (DP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "4" represents that Thru Lanes data is being run.
Call OpenCopy(4, txtLanes.Value)

Worksheets("Thru_Lanes").UsedRange

'Call Run_ThruLanes macro (DP_02_AADT module):
'   (1) Formats the Thru Lane data in preparation for the combination/segmentation process.
Run_ThruLanes

Worksheets("Thru_Lanes").UsedRange

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "Lanes Complete. Loading Speed Limit."
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False


''Speed Limit Preparation
'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "3" represents that Speed Limit data is being run.
Call OpenCopy(3, txtSL.Value)

'Call Run_SignFaces macro (PMDP_04_SpeedLimit module):
'   (1) Formats the speed limit data in preparation for the combination/segmentation process.
Run_SpeedLimit

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "Speed Limit Complete. Loading Urban Code."
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

Worksheets("Speed_Limit").UsedRange


''Urban Code Preparation
'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "5" represents that Urban Code data is being run.
Call OpenCopy(5, txtUC.Value)

'Call Run_UrbanCode macro (PMDP_06_UrbanCode module):
'   (1) Formats the Urban Code data in preparation for the combination/segmentation process.
Run_UrbanCode

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "Urban Code Complete. Now Combining Files."
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

Worksheets("Urban_Code").UsedRange

'Combine files
Run_Combined

Application.ScreenUpdating = True

'Reopen form
Workbooks(guiwb).Sheets("Home").Activate
form_camsinput.Show

'Enter filepath
form_camsinput.txt_camscrashfilepath.Value = Sheets("Inputs").Range("M6").Value
form_camsinput.txt_camsrdwyfilepath.Value = Sheets("Inputs").Range("M5").Value


End Sub

Private Sub cmdData_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select General Crash Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtData = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtData = ""
Else
    txtData = replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdIntersections_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Intersections Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtIntersections = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtIntersections = ""
Else
    txtIntersections = replace(FilePath, "\", "/")
End If

End Sub
Private Sub cmdFC_Click()
' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Lanes Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtFC = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtFC = ""
Else
    txtFC = replace(FilePath, "\", "/")
End If

End Sub
Private Sub cmdCancel_Click()

form_CreateCAMSData.Hide

form_camsinput.Show

End Sub
Private Sub cmdLanes_Click()
' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Lanes Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtLanes = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtLanes = ""
Else
    txtLanes = replace(FilePath, "\", "/")
End If
End Sub

Private Sub cmdSL_Click()
' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Speed Limit Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtSL = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtSL = ""
Else
    txtSL = replace(FilePath, "\", "/")
End If
End Sub

Private Sub cmdUC_Click()
' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Urban Code Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtUC = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtUC = ""
Else
    txtUC = replace(FilePath, "\", "/")
End If
End Sub

Private Sub opt_250_Click()

cmd_crashspeedlimit.Visible = False
txt_crashspeedlimit.Visible = False

End Sub

Private Sub opt_center_Click()

End Sub


Private Sub opt_speedlimit_Click()

cmd_crashspeedlimit.Visible = True
txt_crashspeedlimit.Visible = True

End Sub

Private Sub opt_stopbar_Click()

End Sub

Private Sub UserForm_Activate()

'Blank the values for the user form

txtAADT = Sheets("Inputs").Cells(17, 13)
txtFC = Sheets("Inputs").Cells(18, 13)
txtLanes = Sheets("Inputs").Cells(19, 13)
txtSL = Sheets("Inputs").Cells(20, 13)
txtUC = Sheets("Inputs").Cells(21, 13)

opt_EveryChange.Value = False
opt_MaxLength.Value = False
txt_MaxSegLength.Value = ""
txt_MinSegLength.Value = "0.1"

txt_crashdata = Sheets("Inputs").Cells(23, 13)
txt_crashlocation = Sheets("Inputs").Cells(24, 13)
txt_crashrollup = Sheets("Inputs").Cells(25, 13)
txt_crashvehicle = Sheets("Inputs").Cells(26, 13)

chk_SRatSR.Value = False
chk_SRatFA.Value = False
chk_SRatSignal.Value = False

txt_intersection.Value = Sheets("Inputs").Cells(27, 13)

opt_speedlimit.Value = False
opt_250.Value = False
txt_crashspeedlimit.Value = Sheets("Inputs").Cells(29, 13)

cmd_crashspeedlimit.Visible = False
txt_crashspeedlimit.Visible = False

opt_stopbar.Value = False
opt_center.Value = False

End Sub

