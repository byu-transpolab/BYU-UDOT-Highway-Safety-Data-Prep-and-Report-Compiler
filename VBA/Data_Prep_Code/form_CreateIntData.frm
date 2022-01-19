VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_CreateIntData 
   Caption         =   "Intersection Data Preparation"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   OleObjectBlob   =   "form_CreateIntData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_CreateIntData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chbxSRSignal_Click()

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
    txtAADT = Replace(FilePath, "\", "/")
End If

End Sub


Private Sub cmdCancel_Click()

form_CreateIntData.Hide

form_uicpminput.Show

End Sub

Private Sub cmdCombineCrash_Click()

Dim guiwb, wd As String

wd = Sheets("Inputs").Range("I2")

If txtData = "" Or txtLocation = "" Or txtRollup = "" Or txtVehicle = "" Then
    MsgBox "Select file paths for all given crash datasets before combining the data.", , "Select All Filepaths"
    Exit Sub
End If

'fill in info on the inputs sheet for future use
If chbxSRtoSR = True Then Sheets("Inputs").Cells(13, 9) = "YES" Else Sheets("Inputs").Cells(13, 9) = ""
If chbxSRtoFedAid = True Then Sheets("Inputs").Cells(14, 9) = "YES" Else Sheets("Inputs").Cells(14, 9) = ""
If chbxSRSignal = True Then Sheets("Inputs").Cells(15, 9) = "YES" Else Sheets("Inputs").Cells(15, 9) = ""

Sheets("Inputs").Cells(18, 9) = txtData
Sheets("Inputs").Cells(19, 9) = txtLocation
Sheets("Inputs").Cells(20, 9) = txtRollup
Sheets("Inputs").Cells(21, 9) = txtVehicle


form_CreateIntData.Hide

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


''Crash Location Data Preparation
Call OpenCopy(6, txtLocation.Value)
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


''General Crash Data Preparation
Call OpenCopy(7, txtData.Value)
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


''Crash Rollup Data Preparation
Call OpenCopy1(8, txtRollup.Value)
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


''Crash Vehicle Data Preparation
Call OpenCopy(9, txtVehicle.Value)
Worksheets("Vehicle").UsedRange

'Update progress screen
Workbooks(guiwb).Sheets("Progress").Activate
ActiveWindow.Zoom = 160
With Sheets("Progress")
    .Range("A2") = "Loading Crash Files: Vehicle Complete"
    .Range("A3") = "All Crash Data Imported. Now Combining Files."
    .Range("A5") = "Update Time"
    .Range("B5") = Time
End With
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

'Run IntVehiclePrep macro
IntVehiclePrep

'Run Database Cleanup macro
'   (1) Crash data is joined based on Crash_ID.
'   (2) First vehicle direction is determined from vehicle crash data.
'   (3) Route numbers and direction values fixed and labels created.
DatabaseCleanup

'Delete worksheets containing crash data besides rollup now that all data is stored with the location data
Application.DisplayAlerts = False
For Each wksht In Worksheets
    If wksht.Name = "Location" Then
        wksht.Delete
    ElseIf wksht.Name = "Crash" Then
        wksht.Delete
    ElseIf wksht.Name = "Vehicle" Then
        wksht.Delete
    End If
Next
Application.DisplayAlerts = True

'Change Rollup sheet name to "Crash Data". Set color of sheet tab.
Sheets("Rollup").Name = "Crash Data"
Worksheets("Crash Data").Tab.ColorIndex = 9

'Save crash data file
Sheets("Crash Data").Move
sname = wd & "/IntCrash_Input" & "_" & Replace(Date, "/", "-") & "_" & Replace(Time, ":", "-") & ".csv"
sname = Replace(sname, " ", "_")
sname = Replace(sname, "/", "\")
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs FileName:=sname, FileFormat:=xlCSV
Workbooks(guiwb).Sheets("Inputs").Range("I6").Value = Replace(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, "\", "/")
ActiveWorkbook.Close
Application.DisplayAlerts = True

'Assign start time and end time
StartTime = Workbooks(guiwb).Sheets("Progress").Range("B4").Text
EndTime = Workbooks(guiwb).Sheets("Progress").Range("B6").Text

Application.ScreenUpdating = True

'Message box with start and end times
MsgBox "The Crash input file has been successfully created and saved in the working directory folder. The following is a summary of the process:" & Chr(10) & _
Chr(10) & _
"Process: Crash Input" & Chr(10) & _
"Start Time: " & StartTime & Chr(10) & _
"End Time: " & EndTime & Chr(10), vbOKOnly, "Process Complete"
Workbooks(guiwb).Sheets("Home").Activate

form_uicpminput.txt_intcrashfilepath.Value = Sheets("Inputs").Range("I6")

form_uicpminput.Show


End Sub

Private Sub cmdCombineRoadway_Click()

Dim guiwb As String
Dim wksht As Worksheet

'Check if text boxes are blank   'added by Camille on May 14, 2019
If txtAADT = "" Or txtFCSR = "" Or txtFCFED = "" Or txtIntersections = "" Or txtLanes = "" Or txtSL = "" Or txtUC = "" Or txtPavMess = "" Or _
(chbxSRtoSR.Value = False And chbxSRtoFedAid.Value = False And chbxSRSignal.Value = False) Then
    MsgBox "Select file paths for all given roadway datasets and choose the desired intersection types before combining the data.", , "Select All Filepaths"
    Exit Sub
End If

form_CreateIntData.Hide

'print intersection type to inputs sheet for future use    'added by Camille on May 14, 2019
If chbxSRtoSR = True Then
    Sheets("Inputs").Range("I13").Value = "YES"
End If
If chbxSRtoFedAid = True Then
    Sheets("Inputs").Range("I14").Value = "YES"
End If
If chbxSRSignal = True Then
    Sheets("Inputs").Range("I15").Value = "YES"
End If

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
    ElseIf wksht.Name = "Intersections" Then
        wksht.Delete
    ElseIf wksht.Name = "Pavement_Messages" Then
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
Run_Int_AADT

Worksheets("AADT").UsedRange

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "Loading Roadway Files: AADT Complete"
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

If Sheets("Inputs").Cells(2, 16) = "ISAM" Then
    CombineFCFiles
Else
    Rows("1:2").Delete shift:=xlUp                                                      'FLAGGED: This code seems a bit arbitrary
    Range("I11").Copy
    Range("A1:G573").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Shapes.Range(Array("sort.gif")).Delete
    ActiveSheet.Shapes.Range(Array("mnu_sort_asc.gif")).Delete
    Rows("574:574").Delete shift:=xlUp
End If

' change this to be the combined spreadsheet
Call OpenCopy(2, "FunctionalClassCombined")

'Call Run_FunctionalClass macro:
'   (1)
Run_FunctionalClass

Worksheets("Functional_Class").UsedRange

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "Loading Roadway Files: FC Complete"
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
    .Range("A2") = "Loading Roadway Files: Lanes Complete"
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
    .Range("A2") = "Loading Roadway Files: Speed Limit Complete"
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
    .Range("A2") = "Loading Roadway Files: Urban Code Complete"
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

Worksheets("Urban_Code").UsedRange


''Pavement Message Preparation
'Call OpenCopy macro (PMDP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "12" represents that Pavement Message data is being run.
Call OpenCopy(12, txtPavMess.Value)

'Call Run_UrbanCode macro (PMDP_06_UrbanCode module):
'   (1) Formats the Urban Code data in preparation for the combination/segmentation process.
Run_PavementMessages

'Update progress screen
Sheets("Progress").Activate
ActiveWindow.Zoom = 160
Application.ScreenUpdating = True
With Sheets("Progress")
    .Range("A2") = "Roadway Files Loaded: Pavement Messages Complete"
    .Range("A3") = "Now Importing Intersections file and combining Files."
    .Range("B5") = Time
End With
Application.Wait (Now + TimeValue("00:00:02"))
Application.ScreenUpdating = False

Worksheets("Pavement_Messages").UsedRange



''Intersections Preparation
'Call OpenCopy macro (DP_01_Home module):
'   (1) Asks user to open data file that corresponds with the working dataset and data number.
'   (2) Runs CheckHeaders and CopyDataSets macros to copy data from selected file.
'   (3) The "11" represents that intersections data is being run.
Call OpenCopy(11, txtIntersections.Value)

'Call Run_Intersections macro
'
Run_Intersections

Application.ScreenUpdating = True

Workbooks(guiwb).Sheets("Home").Activate

form_uicpminput.txt_intfilepath.Value = Sheets("Inputs").Range("I5")

form_uicpminput.Show

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
    txtData = Replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdFCSR_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Functional Class Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtFCSR = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtFCSR = ""
Else
    txtFCSR = Replace(FilePath, "\", "/")
End If

End Sub


Private Sub cmdFCFED_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Functional Class Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtFCFED = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtFCFED = ""
Else
    txtFCFED = Replace(FilePath, "\", "/")
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
    txtIntersections = Replace(FilePath, "\", "/")
End If

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
    txtLanes = Replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdLocation_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Crash Location Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtLocation = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtLocation = ""
Else
    txtLocation = Replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdPavMess_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Pavement Message Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtPavMess = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtPavMess = ""
Else
    txtPavMess = Replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdRollup_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Crash Rollup Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtRollup = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtRollup = ""
Else
    txtRollup = Replace(FilePath, "\", "/")
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
    txtSL = Replace(FilePath, "\", "/")
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
    txtUC = Replace(FilePath, "\", "/")
End If

End Sub

Private Sub cmdVehicle_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select Crash Vehicle Data"
    If .Show <> -1 Then
        MsgBox "No file selected.":
        txtVehicle = ""
        Exit Sub
    End If
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txtVehicle = ""
Else
    txtVehicle = Replace(FilePath, "\", "/")
End If

End Sub

Private Sub txtData_Change()

End Sub

Private Sub UserForm_Click()

End Sub
