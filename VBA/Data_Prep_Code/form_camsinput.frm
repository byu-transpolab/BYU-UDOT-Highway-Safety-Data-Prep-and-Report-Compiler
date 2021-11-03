VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_camsinput 
   Caption         =   "Input to CAMS"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "form_camsinput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_camsinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chbx_camsseverity1_Click()
    Call camscheckblanks
End Sub

Private Sub chbx_camsseverity2_Click()
    Call camscheckblanks
End Sub

Private Sub chbx_camsseverity3_Click()
    Call camscheckblanks
End Sub

Private Sub chbx_camsseverity4_Click()
    Call camscheckblanks
End Sub

Private Sub chbx_camsseverity5_Click()
    Call camscheckblanks
End Sub

Private Sub cmd_camscrashdata_Click()

' Define variables
Dim FilePath As Variant
Dim wdFP As String
Dim row1, col1 As Integer

'Find working directory location
row1 = 1
col1 = 1
Do Until Sheets("Inputs").Cells(row1, col1) = "CAMS"
    col1 = col1 + 1
Loop
Do Until Sheets("Inputs").Cells(row1, col1) = "Working Directory"
    row1 = row1 + 1
Loop
col1 = col1 + 1

'Assign working directory file path
wdFP = replace(Sheets("Inputs").Cells(row1, col1), "/", "\")

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = wdFP
    .AllowMultiSelect = False
    .Title = "Select Combined Crash Data"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_camscrashfilepath = ""
Else
    txt_camscrashfilepath = replace(FilePath, "\", "/")
    
End If

Call camscheckblanks

End Sub

Private Sub cmd_camsrdwydata_Click()
' Define variables
Dim FilePath As Variant
Dim wdFP As String
Dim row1, col1 As Integer

'Find working directory location
row1 = 1
col1 = 1
Do Until Sheets("Inputs").Cells(row1, col1) = "CAMS"
    col1 = col1 + 1
Loop
Do Until Sheets("Inputs").Cells(row1, col1) = "Working Directory"
    row1 = row1 + 1
Loop
col1 = col1 + 1

'Assign working directory file path
wdFP = replace(Sheets("Inputs").Cells(row1, col1), "/", "\")

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = wdFP
    .AllowMultiSelect = False
    .Title = "Select Combined Roadway Data"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_camsrdwyfilepath = ""
Else
    txt_camsrdwyfilepath = replace(FilePath, "\", "/")
End If

Call camscheckblanks

End Sub

Private Sub cmd_CreateCAMSData_Click()

form_camsinput.Hide
form_CreateCAMSData.Show

End Sub

Private Sub cmd_createcamsinputdata_Click()

If FileExists(txt_camsrdwyfilepath) And txt_camsrdwyfilepath.Value <> "" And FileExists(txt_camscrashfilepath) And txt_camscrashfilepath.Value <> "" And Len(txtCAMSminyr.Value) = 4 And Len(txtCAMSmaxyr.Value) = 4 And _
(chbx_camsseverity5.Value Or chbx_camsseverity4.Value Or chbx_camsseverity3.Value Or chbx_camsseverity2.Value Or chbx_camsseverity1.Value) Then
    'no need for any action
Else
    MsgBox "Please enter valid filepaths and dates for the roadway and crash data before running the code.", vbOKOnly, "Not enough input to run"
    Exit Sub
End If

' Define variables
Dim workingdirectory As String
Dim segmentfilepath As String
Dim crashfilepath As String
Dim severitylist As String

' Extract values from textbox and checkbox selection
segmentfilepath = txt_camsrdwyfilepath
crashfilepath = txt_camscrashfilepath
severitylist = ""
If chbx_camsseverity1 = True Then
    If severitylist = "" Then
        severitylist = "1"
    End If
End If
If chbx_camsseverity2 = True Then
    If severitylist = "" Then
        severitylist = "2"
    Else
        severitylist = severitylist & "2"
    End If
End If
If chbx_camsseverity3 = True Then
    If severitylist = "" Then
        severitylist = "3"
    Else
        severitylist = severitylist & "3"
    End If
End If
If chbx_camsseverity4 = True Then
    If severitylist = "" Then
        severitylist = "4"
    Else
        severitylist = severitylist & "4"
    End If
End If
If chbx_camsseverity5 = True Then
    If severitylist = "" Then
        severitylist = "5"
    Else
        severitylist = severitylist & "5"
    End If
End If

'print inputs to workbook for future information
ActiveWorkbook.Sheets("Inputs").Range("M5").Value = replace(txt_camsrdwyfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("M6").Value = replace(txt_camscrashfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("M31").Value = txtCAMSminyr & "-" & txtCAMSmaxyr
ActiveWorkbook.Sheets("Inputs").Range("M32").Value = severitylist

'hide user form
form_camsinput.Hide

'start process to count crash severity
Call CAMSdataprep(segmentfilepath, crashfilepath, True, chbx_camsseverity1, chbx_camsseverity2, chbx_camsseverity3, chbx_camsseverity4, chbx_camsseverity5, Int(txtCAMSminyr), Int(txtCAMSmaxyr))

End Sub

Private Sub cmdEditParameters_Click()

End Sub

Private Sub cmdSelectAllSev_Click()
    chbx_camsseverity5.Value = True
    chbx_camsseverity4.Value = True
    chbx_camsseverity3.Value = True
    chbx_camsseverity2.Value = True
    chbx_camsseverity1.Value = True
End Sub

Private Sub cmdSelectNoneSev_Click()
    chbx_camsseverity5.Value = False
    chbx_camsseverity4.Value = False
    chbx_camsseverity3.Value = False
    chbx_camsseverity2.Value = False
    chbx_camsseverity1.Value = False
End Sub

Sub camscheckblanks()

' Check if the information has been filled before allowing the user to continue
If FileExists(txt_camsrdwyfilepath) And txt_camsrdwyfilepath.Value <> "" And FileExists(txt_camscrashfilepath) And txt_camscrashfilepath.Value <> "" And _
(chbx_camsseverity5.Value Or chbx_camsseverity4.Value Or chbx_camsseverity3.Value Or chbx_camsseverity2.Value Or chbx_camsseverity1.Value) Then
    cmd_createcamsinputdata.Visible = True
    lbl_camsstop.Visible = False
Else
    cmd_createcamsinputdata.Visible = False
    lbl_camsstop.Visible = True
End If

End Sub

Private Sub opt_camsFAuserdefined_Click()
    lblDefineBy.Visible = True
    cbx_VariableCAMS.Visible = True
    cmdEditParameters.Visible = True
End Sub

Private Sub opt_FArec_Click()
    lblDefineBy.Visible = False
    cbx_VariableCAMS.Visible = False
    cmdEditParameters.Visible = False
End Sub

Private Sub txtCAMSmaxyr_Change()
Dim maxyrstring As String
maxyrstring = txtCAMSmaxyr.Value
If Len(maxyrstring) = 4 And txtCAMSmaxyr.Value < txtCAMSminyr.Value Then
    MsgBox ("The Maximum date must be equal to or larger than the Minimum date")
    txtCAMSmaxyr.Value = ""
    txtCAMSminyr.Value = ""
End If
End Sub

Private Sub txtCAMSminyr_Change()
Dim minyrstring As String
minyrstring = txtCAMSminyr.Value

    If Len(minyrstring) = 4 And txtCAMSminyr.Value < 2010 Then
        MsgBox ("Please enter a date no earlier than 2010.")
        txtCAMSminyr.Value = ""
    End If
End Sub

Private Sub UserForm_Activate()

cmd_createcamsinputdata.Visible = False

End Sub

Private Sub UserForm_Click()

End Sub
