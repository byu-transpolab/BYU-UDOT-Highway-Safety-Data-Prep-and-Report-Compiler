VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_uicpminput 
   Caption         =   "Safety Statistical Analysis: UICPM Input (R GUI)"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   OleObjectBlob   =   "form_uicpminput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_uicpminput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'R GUI workbook created for UDOT Roadway Safety Analysis Methodology
'Comments by Sam Mineer, Brigham Young Univerisity, June 2016

Private Sub cbx_Variable_Change()
    Call checkblanks
    
    If cbx_Variable.Value = "Functional Class" Then
        MsgBox "Sorry, the BYU team has decided to discontinue this option." & Chr(10) & Chr(10) & _
        "Please select either speed limit or fixed length to define functional area.", vbOKOnly, "This option discontinued"
        cbx_Variable.Value = ""
        'form_FAFunctionalClass.Show
    ElseIf cbx_Variable.Value = "Speed Limit" Then
        form_FASpeedLimit.Show
    ElseIf cbx_Variable.Value = "Urban Code" Then
        MsgBox "Sorry, the BYU team has decided to discontinue this option." & Chr(10) & Chr(10) & _
        "Please select either speed limit or fixed length to define functional area.", vbOKOnly, "This option discontinued"
        cbx_Variable.Value = ""
        'form_FAUrbanCode.Show
    ElseIf cbx_Variable.Value = "Fixed Length" Then         'added by Camille on May 9, 2019
        form_FAFixedLength.Show
    End If
    
End Sub

Private Sub chbx_intseverity5_Click()
    Call checkblanks
End Sub

Private Sub chbx_intseverity4_Click()
    Call checkblanks
End Sub

Private Sub chbx_intseverity3_Click()
    Call checkblanks
End Sub

Private Sub chbx_intseverity2_Click()
    Call checkblanks
End Sub

Private Sub chbx_intseverity1_Click()
    Call checkblanks
End Sub

Private Sub txt_crashfilepath_Change()

    ActiveWorkbook.Sheets("Inputs").Range("B6").Value = replace(txt_crashfilepath, "\", "/")
    
    Call checkblanks
End Sub

Private Sub txt_segmentfilepath_Change()
    
    ActiveWorkbook.Sheets("Inputs").Range("B5").Value = replace(txt_segmentfilepath, "\", "/")
    
    Call checkblanks
End Sub

Private Sub cmd_CreateData_Click()

form_uicpminput.Hide
form_CreateIntData.Show

End Sub

Private Sub cmd_intcrashdata_Click()

' Define variables
Dim FilePath As Variant
Dim wdFP As String
Dim row1, col1 As Integer

'Find working directory location
row1 = 1
col1 = 1
Do Until Sheets("Inputs").Cells(row1, col1) = "UICPM"
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
    .Title = "Select Crash Data"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_intcrashfilepath = ""
Else
    txt_intcrashfilepath = replace(FilePath, "\", "/")
    
End If

End Sub

Private Sub cmd_createintinputdata_Click()

' Define variables
Dim workingdirectory As String
Dim intfilepath As String
Dim intcrashfilepath As String
Dim severitylist As String
Dim PSheet As String

' Extract values from textbox and checkbox selection
intfilepath = txt_intfilepath
intcrashfilepath = txt_intcrashfilepath
severitylist = ""
If chbx_intseverity1 = True Then
    If severitylist = "" Then
        severitylist = "1"
    End If
End If
If chbx_intseverity2 = True Then
    If severitylist = "" Then
        severitylist = "2"
    Else
        severitylist = severitylist & "2"
    End If
End If
If chbx_intseverity3 = True Then
    If severitylist = "" Then
        severitylist = "3"
    Else
        severitylist = severitylist & "3"
    End If
End If
If chbx_intseverity4 = True Then
    If severitylist = "" Then
        severitylist = "4"
    Else
        severitylist = severitylist & "4"
    End If
End If
If chbx_intseverity5 = True Then
    If severitylist = "" Then
        severitylist = "5"
    Else
        severitylist = severitylist & "5"
    End If
End If

'print inputs to workbook for future information
ActiveWorkbook.Sheets("Inputs").Range("I5").Value = replace(intfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I6").Value = replace(intcrashfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("I12").Value = severitylist
ActiveWorkbook.Sheets("Inputs").Range("I16").Value = cbx_Variable.Value

'Create parameters sheet to store severities, functional area type, and crash data by crash ID
PSheet = "Parameters"
Sheets.Add.Name = PSheet
Sheets(PSheet).Cells(1, 1) = "Severities:"
Sheets(PSheet).Cells(1, 2) = severitylist
Sheets(PSheet).Cells(2, 1) = "Functional Area Type:"
Sheets(PSheet).Cells(3, 1) = "Selected Years:"
Sheets(PSheet).Cells(4, 1) = "Crash Data:"

'hide user form
form_uicpminput.Hide

'start process to count crash severity
Call UICPMdataprep(intfilepath, intcrashfilepath)

End Sub

Private Sub cmd_intdata_Click()

' Define variables
Dim FilePath As Variant
Dim wdFP As String
Dim row1, col1 As Integer

'Find working directory location
row1 = 1
col1 = 1
Do Until Sheets("Inputs").Cells(row1, col1) = "UICPM"
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
    .Title = "Select Road Intersection Data"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_intfilepath = ""
Else
    txt_intfilepath = replace(FilePath, "\", "/")
End If

End Sub

Sub checkblanks()

' Check if the information has been filled before allowing the user to continue
If FileExists(txt_intfilepath) And txt_intfilepath.Value <> "" And FileExists(txt_intcrashfilepath) And txt_intcrashfilepath.Value <> "" And _
(chbx_intseverity5.Value Or chbx_intseverity4.Value Or chbx_intseverity3.Value Or chbx_intseverity2.Value Or chbx_intseverity1.Value) And _
((opt_FAuserdefined.Value And cbx_Variable.Value <> "") Or opt_FAspeedlimit.Value) Then
    cmd_createintinputdata.Visible = True
    lbl_intstop.Visible = False
Else
    cmd_createintinputdata.Visible = False
    lbl_intstop.Visible = True
End If

End Sub

Private Sub cmdEditParameters_Click()
    
    If cbx_Variable.Value = "Functional Class" Then
        form_FAFunctionalClass.Show
    ElseIf cbx_Variable.Value = "Speed Limit" Then
        form_FASpeedLimit.Show
    ElseIf cbx_Variable.Value = "Urban Code" Then
        form_FAUrbanCode.Show
    ElseIf cbx_Variable.Value = "Fixed Length" Then
        form_FAFixedLength.Show
    End If
    
End Sub

Private Sub cmdSelectAll_Click()
    chbx_intseverity5.Value = True
    chbx_intseverity4.Value = True
    chbx_intseverity3.Value = True
    chbx_intseverity2.Value = True
    chbx_intseverity1.Value = True
End Sub

Private Sub cmdSelectNone_Click()
    chbx_intseverity5.Value = False
    chbx_intseverity4.Value = False
    chbx_intseverity3.Value = False
    chbx_intseverity2.Value = False
    chbx_intseverity1.Value = False
End Sub

Private Sub lbl_intstop_Click()

End Sub

Private Sub opt_FAspeedlimit_Click()

Call checkblanks

If opt_FAuserdefined.Value = True Then
    lblDefineBy.Visible = True
    cbx_Variable.Visible = True
    cmdEditParameters.Visible = True
    opt_FAspeedlimit.Caption = "Recommended Functional Area"
Else
    lblDefineBy.Visible = False
    cbx_Variable.Visible = False
    cmdEditParameters.Visible = False
    opt_FAspeedlimit.Caption = "Recommended Functional Area (based on Speed Limit)"
End If

If opt_FAspeedlimit Then
    Dim colUICPM, rowFA, i As Integer
    Dim colFASpeed As Long
    ReDim FASpeed(1 To 12)
    
    
    colFASpeed = 1
    Do Until Sheets("Key").Cells(1, colFASpeed) = "Functional Area"
        colFASpeed = colFASpeed + 1
    Loop
    
    For i = 1 To 12
        FASpeed(i) = Sheets("Key").Cells(2 + i, colFASpeed + 4)
    Next i
    
    
    colUICPM = 1
    rowFA = 1
    
    Application.ScreenUpdating = False
    
    Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "UICPM"
        colUICPM = colUICPM + 1
    Loop
    
    Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "Selected FA Parameter"
        rowFA = rowFA + 1
    Loop
    
    Sheets("Inputs").Cells(rowFA, colUICPM + 1) = "Speed Limit"
    Sheets("Inputs").Cells(rowFA + 1, colUICPM) = "Speed Limit"   'table header column 1
    Sheets("Inputs").Cells(rowFA + 1, colUICPM + 1) = "Functional Area"     'table header column 2
    'clear the table if it has data in it
    Sheets("Inputs").Activate
    Sheets("Inputs").Range(Cells(rowFA + 2, colUICPM), Cells(rowFA + 20, colUICPM + 1)).ClearContents
    Sheets("Home").Activate
    'paste in the table data for future reference and use
    For i = 1 To 12
        Sheets("Inputs").Cells(rowFA + 1 + i, colUICPM) = 15 + (i * 5)
        Sheets("Inputs").Cells(rowFA + 1 + i, colUICPM + 1) = FASpeed(i)
    Next i
    
    
    Application.ScreenUpdating = True
End If

End Sub

Private Sub opt_FAuserdefined_Click()

Call checkblanks

If opt_FAuserdefined.Value = True Then
    lblDefineBy.Visible = True
    cbx_Variable.Visible = True
    cmdEditParameters.Visible = True
    opt_FAspeedlimit.Caption = "Recommended Functional Area"
Else
    lblDefineBy.Visible = False
    cbx_Variable.Visible = False
    cmdEditParameters.Visible = False
    opt_FAspeedlimit.Caption = "Recommended Functional Area (based on Speed Limit)"
End If

If opt_FAspeedlimit Then
    Dim colUICPM, rowFA, i As Integer
    Dim colFASpeed As Long
    ReDim FASpeed(1 To 12)
    
    
    colFASpeed = 1
    Do Until Sheets("Key").Cells(1, colFASpeed) = "Functional Area"
        colFASpeed = colFASpeed + 1
    Loop
    
    For i = 1 To 12
        FASpeed(i) = Sheets("Key").Cells(2 + i, colFASpeed + 4)
    Next i
    
    
    colUICPM = 1
    rowFA = 1
    
    Application.ScreenUpdating = False
    
    Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "UICPM"
        colUICPM = colUICPM + 1
    Loop
    
    Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "Selected FA Parameter"
        rowFA = rowFA + 1
    Loop
    
    Sheets("Inputs").Cells(rowFA, colUICPM + 1) = "Speed Limit"
    Sheets("Inputs").Cells(rowFA + 1, colUICPM) = "Speed Limit"
    Sheets("Inputs").Cells(rowFA + 1, colUICPM + 1) = "Functional Area"
    
    Sheets("Inputs").Activate
    Sheets("Inputs").Range(Cells(rowFA + 2, colUICPM), Cells(rowFA + 20, colUICPM + 1)).ClearContents
    Sheets("Home").Activate
    
    For i = 1 To 12
        Sheets("Inputs").Cells(rowFA + 1 + i, colUICPM) = 15 + (i * 5)
        Sheets("Inputs").Cells(rowFA + 1 + i, colUICPM + 1) = FASpeed(i)
    Next i
    
    
    Application.ScreenUpdating = True
End If

End Sub

Private Sub txt_intcrashfilepath_Change()
    Call checkblanks
End Sub

Private Sub txt_intfilepath_Change()
    Call checkblanks
End Sub

Private Sub UserForm_Activate()

' Blank the values for the user form
If Sheets("Inputs").Range("I5") = "" Then
    txt_intfilepath.Value = ""
Else
    txt_intfilepath.Value = Sheets("Inputs").Range("I5")
End If

If Sheets("Inputs").Range("I6") = "" Then
    txt_intcrashfilepath.Value = ""
Else
    txt_intcrashfilepath.Value = Sheets("Inputs").Range("I6")
End If

chbx_intseverity5.Value = False
chbx_intseverity4.Value = False
chbx_intseverity3.Value = False
chbx_intseverity2.Value = False
chbx_intseverity1.Value = False

opt_FAspeedlimit.Value = False
opt_FAuserdefined.Value = False

' Hide the creat input data button, until the user fills in the necessary data
lbl_intstop.Visible = True
cmd_createintinputdata.Visible = False

'Hide user-defined functional area labels and text boxes
lblDefineBy.Visible = False
cbx_Variable.Visible = False
cmdEditParameters.Visible = False
opt_FAspeedlimit.Caption = "Recommended Functional Area (based on Speed Limit)"


End Sub

