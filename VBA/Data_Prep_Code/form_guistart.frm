VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_guistart 
   Caption         =   "Safety Statistical Analysis: Start (R GUI)"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   OleObjectBlob   =   "form_guistart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_guistart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'R GUI workbook created for UDOT Roadway Safety Analysis Methodology
'Comments by Sam Mineer, Brigham Young Univerisity, June 2016

Private Sub cmbx_modelselect_Change()

' Do not show the "Next" button until a model has been selected
If cmbx_modelselect.Value = "Select Statistical Model" Then
    cmd_guicreateinput.Visible = False
    cmd_guiexisting.Visible = False
Else
    cmd_guicreateinput.Visible = True
    cmd_guiexisting.Visible = True
End If

End Sub


Private Sub cmd_guicreateinput_Click()

' Define variables
Dim selectedmodel As String

' Save this value to decide which user form to open next
selectedmodel = cmbx_modelselect.Value

' Print the selected inputs to workbook, before moving on
' These values will be important for executing the R code
ActiveWorkbook.Sheets("Inputs").Range("B2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("M2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("B3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("M3").Value = Replace(txt_rfilepath, "\", "/")

' Hide the GUI Start user form
form_guistart.Hide

' Decided which form to show next, based on the model selection
If Left(selectedmodel, 7) = "Segment" Then
    Sheets("Inputs").Cells(2, 16) = "RSAM"
    form_ucpsminput.Show
ElseIf selectedmodel = "Before-After" Then
    Sheets("Inputs").Cells(2, 16) = "Before-After"
    form_bainput.Show
ElseIf Left(selectedmodel, 12) = "Intersection" Then
    Sheets("Inputs").Cells(2, 16) = "ISAM"
    form_uicpminput.Show
ElseIf Left(selectedmodel, 4) = "2019" Then
    Sheets("Inputs").Cells(2, 16) = "CAMS"
    form_camsinput.Show
End If

End Sub

Private Sub cmd_guiexisting_Click()

' Define variables
Dim selectedmodel As String

' Save this value to decide which user form to open next
selectedmodel = cmbx_modelselect.Value

' Print the selected inputs to workbook, before moving on
' These values will be important for executing the R code
ActiveWorkbook.Sheets("Inputs").Range("B2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I2").Value = Replace(txt_workingDirectory, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("B3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I3").Value = Replace(txt_rfilepath, "\", "/")

' Hide the GUI Start user form
form_guistart.Hide

' Decided which form to show next, based on the model selection
If Left(selectedmodel, 7) = "Segment" Then
    Sheets("Inputs").Cells(2, 16) = "RSAM"
    form_ucpsmvariable.Show
ElseIf selectedmodel = "Before-After" Then
    Sheets("Inputs").Cells(2, 16) = "Before-After"
    form_bamodel.Show
ElseIf Left(selectedmodel, 12) = "Intersection" Then
    Sheets("Inputs").Cells(2, 16) = "ISAM"
    form_uicpmvariable.Show
ElseIf Left(selectedmodel, 4) = "2019" Then
    Sheets("Inputs").Cells(2, 16) = "CAMS"
    MsgBox "This button isn't ready yet. Please be patient with Camille  d-;", vbOKOnly, "I'm not ready yet"
End If



End Sub

Private Sub cmd_installRpackages_Click()

' Define variables
Dim cmdLine As String
Dim rcode As Variant
Dim workingdirectory As String
Dim libfp As String

' Set working directory path
workingdirectory = ActiveWorkbook.path

' The R script to download the packages will be included with the GUI file
rcode = workingdirectory & "\downloadPackages.R"
rcode = Replace(rcode, "\", "/")

libfp = ActiveWorkbook.Sheets("Inputs").Range("B3").Value
libfp = Left(libfp, InStr(1, libfp, "R/R-") + 9)
libfp = libfp & "library/"
libfp = Replace(libfp, "\", "/")
libfp = Replace(libfp, " ", "ZZZ")

cmdLine = Replace(txt_rfilepath, "\", "/") & " " & rcode & " " & libfp
    'MsgBox cmdLine, vbOKOnly   'Debugging purposes
Shell cmdLine, vbMaximizedFocus

lbl_selectmodel.Visible = True
lbl_model.Visible = True
cmbx_modelselect.Visible = True
cmd_installRpackages.Visible = False
lbl_packagestatus = "R Packages Updating... See Rscript window"
'MsgBox "Packages Installed. R up to date.", vbOKOnly   'Debugging purposes

End Sub

Private Sub cmd_selectR_Click()

' Define variables
Dim FilePath As Variant

MsgBox "Select the latest version of the Rscript.exe program." & Chr(10) & "Usually found in Programs -> R -> bin", vbOKOnly, "Select Rscript program"

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = "C:\Program Files\R"
    .AllowMultiSelect = False
    .Title = "Select Crash Data"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will default to the expected program location
' Update if a newer version of R is downloaded
If FilePath = False Then
    txt_rfilepath = "C:/Program Files/R/R-3.2.5/bin/Rscript"
Else
    txt_rfilepath = Replace(Replace(FilePath, "\", "/"), ".exe", "")
End If

End Sub

Private Sub cmd_wdfilepath_Click()
' This function gets a folder, returns a file path

'Define variables
Dim fldr As FileDialog
Dim sItem As String
Dim strpath As String

MsgBox "Select a folder as the working directory for the statistical analysis", vbOKOnly, "Select Working Directory"

With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = ActiveWorkbook.path
    .AllowMultiSelect = False
    .Title = "Select Working Directory"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        sItem = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will default to blank
If sItem = "" Then
    txt_workingDirectory = Replace(ActiveWorkbook.path, "\", "/")
Else
    txt_workingDirectory = Replace(sItem, "\", "/")
End If

End Sub

Private Sub txt_rfilepath_Change()
' Check if the Rscript file path actually exists

If FileExists(txt_rfilepath & ".exe") Then
    If checkRlibraries Then
        cmd_installRpackages.Visible = False
        lbl_packagestatus = "R Packages Up To Date"
        lbl_selectmodel.Visible = True
        lbl_model.Visible = True
        cmbx_modelselect.Visible = True
        cmd_guicreateinput.Visible = False
        cmd_guiexisting.Visible = False
    Else
        cmd_installRpackages.Visible = True
        lbl_packagestatus = "Install R Packages to continue"
        lbl_selectmodel.Visible = False
        lbl_model.Visible = False
        cmbx_modelselect.Visible = False
        cmd_guicreateinput.Visible = False
        cmd_guiexisting.Visible = False
    End If
Else
    cmd_installRpackages.Visible = False
    lbl_packagestatus = "Select the correct Rscript Program"
    lbl_selectmodel.Visible = False
    lbl_model.Visible = False
    cmbx_modelselect.Visible = False
    cmd_guicreateinput.Visible = False
    cmd_guiexisting.Visible = False
End If

ActiveWorkbook.Sheets("Inputs").Range("B3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I3").Value = Replace(txt_rfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F3").Value = Replace(txt_rfilepath, "\", "/")

' Show/Hide Certain Options
' Checks if the R packages has already been downloaded
If checkRlibraries Then
    cmd_installRpackages.Visible = False
    lbl_packagestatus = "R Packages Up To Date"
    lbl_selectmodel.Visible = True
    lbl_model.Visible = True
    cmbx_modelselect.Visible = True
    cmd_guicreateinput.Visible = False
    cmd_guiexisting.Visible = False
Else
    cmd_installRpackages.Visible = True
    lbl_packagestatus = "Install R Packages to continue"
    lbl_selectmodel.Visible = False
    lbl_model.Visible = False
    cmbx_modelselect.Visible = False
    cmd_guicreateinput.Visible = False
    cmd_guiexisting.Visible = False
End If


End Sub


Private Sub txt_workingDirectory_Change()

Dim pos As Integer
Dim fpath As String

fpath = Replace(txt_workingDirectory, "\", "/")
pos = InStr(fpath, " ")

If pos = 0 Then
    ActiveWorkbook.Sheets("Inputs").Range("B2").Value = Replace(txt_workingDirectory, "\", "/")
    ActiveWorkbook.Sheets("Inputs").Range("F2").Value = Replace(txt_workingDirectory, "\", "/")
    ActiveWorkbook.Sheets("Inputs").Range("I2").Value = Replace(txt_workingDirectory, "\", "/")
Else
    MsgBox "Select a file path with no spaces in the entire file path." & Chr(10) & "It is recommended to simplify the file path.", vbOKOnly, "Select a Different Working Directory"
    
End If

End Sub

Private Sub UserForm_Activate()
' When the user form is opened, the workbook inputs will be reset

' Define variables
Dim workingdirectory As String
Dim RFolderPath As String
Dim RVersion As Long
Dim MaxRVersion As Long
Dim MaxRStr As String
Dim folder1 As Object
Dim SubFolder As Object
Dim objFSO As Object
Dim i As Long

' Clear input data on Input, Progress, and Key sheets
ActiveWorkbook.Sheets("Inputs").Columns("B").ClearContents
ActiveWorkbook.Sheets("Inputs").Columns("D").ClearContents
ActiveWorkbook.Sheets("Inputs").Columns("F").ClearContents
ActiveWorkbook.Sheets("Inputs").Columns("I").ClearContents
ActiveWorkbook.Sheets("Inputs").Range("H17:H29").ClearContents
ActiveWorkbook.Sheets("Progress").Columns("B").ClearContents
ActiveWorkbook.Sheets("Progress").Columns("C").ClearContents
ActiveWorkbook.Sheets("Progress").Columns("D").ClearContents
ActiveWorkbook.Sheets("Key").Columns("C").ClearContents
ActiveWorkbook.Sheets("Key").Columns("D").ClearContents
ActiveWorkbook.Sheets("Key").Columns("E").ClearContents

' Removes previously created workbooks
Application.DisplayAlerts = False
If SheetExists("UCPMinput") Then
    Sheets("UCPMinput").Delete
End If
If SheetExists("CrashInput") Then
    Sheets("CrashInput").Delete
End If
If SheetExists("UCSMinput") Then
    Sheets("UCSMinput").Delete
End If
If SheetExists("UCPSMinput") Then
    Sheets("UCPSMinput").Delete
End If
If SheetExists("BAinput") Then
    Sheets("BAinput").Delete
End If
If SheetExists("AADT") Then
    Sheets("AADT").Delete
End If
If SheetExists("Parameters") Then
    Sheets("Parameters").Delete
End If
If SheetExists("UICPMinput") Then
    Sheets("UICPMinput").Delete
End If
Application.DisplayAlerts = True

'Find latest R version if folders are available
i = 1
MaxRVersion = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
RFolderPath = "C:/Program Files/R"
Set folder1 = objFSO.GetFolder(RFolderPath)

For Each SubFolder In folder1.subfolders
    RVersion = Replace(Replace(Right(SubFolder.Name, 5), ".", ""), ".", "")
    If RVersion > MaxRVersion Then
        MaxRVersion = RVersion
    End If
Next SubFolder

' Default values for user form
If MaxRVersion <> 1 Then
    MaxRStr = Left(CStr(MaxRVersion), 1) & "." & Mid(CStr(MaxRVersion), 2, 1) & "." & Right(CStr(MaxRVersion), 1)
Else
    MaxRStr = 111
End If

txt_rfilepath = "C:/Program Files/R/R-" & MaxRStr & "/bin/Rscript"
ActiveWorkbook.Sheets("Inputs").Range("B3").Value = Replace(txt_rfilepath, "\", "/")

If FileExists(txt_rfilepath & ".exe") Then
Else
    'If the expected Rscript.exe program doesn't exist, then prompt the user to update the correct location of the program
    txt_rfilepath = ""
    MsgBox "Please specify Rscript program location.", vbOKOnly, "Select Rscript Program"
End If

workingdirectory = ActiveWorkbook.path
txt_workingDirectory.Value = Replace(workingdirectory, "\", "/")
cmbx_modelselect.Value = "Select Statistical Model"

' Show/Hide Certain Options
' Checks if the R packages has already been downloaded
If checkRlibraries Then
    cmd_installRpackages.Visible = False
    lbl_packagestatus = "R Packages Up To Date"
    lbl_selectmodel.Visible = True
    lbl_model.Visible = True
    cmbx_modelselect.Visible = True
    cmd_guicreateinput.Visible = False
    cmd_guiexisting.Visible = False
Else
    cmd_installRpackages.Visible = True
    lbl_packagestatus = "Install R Packages to continue"
    lbl_selectmodel.Visible = False
    lbl_model.Visible = False
    cmbx_modelselect.Visible = False
    cmd_guicreateinput.Visible = False
    cmd_guiexisting.Visible = False
End If

End Sub

