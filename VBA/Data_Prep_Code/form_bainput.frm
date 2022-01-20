VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_bainput 
   Caption         =   "Safety Statistical Analysis: Before After Model Input (R GUI)"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   OleObjectBlob   =   "form_bainput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_bainput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aadtdata_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
FilePath = Application.GetOpenFilename(, , "Select AADT Data File")

' If the user doesn't select a file, then the box will default to the expected program location
If FilePath = False Then
    txt_aadtfilepath = ""
Else
    txt_aadtfilepath = FilePath
End If

End Sub

Private Sub cmd_analysissegdata_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
FilePath = Application.GetOpenFilename(, , "Select Analysis Segment Data File")

' If the user doesn't select a file, then the box will default to the expected program location
If FilePath = False Then
    txt_analysisfilepath = ""
Else
    txt_analysisfilepath = FilePath
End If

End Sub

Private Sub cmd_crashdata_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
FilePath = Application.GetOpenFilename(, , "Select Crash Data File")

' If the user doesn't select a file, then the box will default to the expected program location
If FilePath = False Then
    txt_crashfilepath = ""
Else
    txt_crashfilepath = FilePath
End If

End Sub

Sub checkBAblanks()

If FileExists(txt_crashfilepath) And FileExists(txt_analysisfilepath) And FileExists(txt_aadtfilepath) And txt_crashfilepath <> "" And txt_analysisfilepath <> "" And txt_aadtfilepath <> "" Then
    cmd_createBAinput.Visible = True
End If

End Sub

Private Sub cmd_createBAinput_Click()

' Define variables
Dim workingdirectory As String
Dim analysisfilepath As String
Dim crashfilepath As String
Dim aadtfilepath As String

'workingdirectory = txt_workingDirectory
analysisfilepath = txt_analysisfilepath
aadtfilepath = txt_aadtfilepath
crashfilepath = txt_crashfilepath

'print inputs to workbook for future information
ActiveWorkbook.Sheets("Inputs").Range("F5").Value = Replace(txt_analysisfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F6").Value = Replace(txt_aadtfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("F7").Value = Replace(txt_crashfilepath, "\", "/")

'hide user form
form_bainput.Hide

'start process to create BA model input file
Call BAdataprep(analysisfilepath, aadtfilepath, crashfilepath)

End Sub

Private Sub txt_aadtfilepath_Change()
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Workbooks(guiwb).Sheets("Inputs").Range("F6") = txt_aadtfilepath.Value
Call checkBAblanks

End Sub

Private Sub txt_analysisfilepath_Change()
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Workbooks(guiwb).Sheets("Inputs").Range("F5") = txt_analysisfilepath.Value
Call checkBAblanks

End Sub

Private Sub txt_crashfilepath_Change()
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Workbooks(guiwb).Sheets("Inputs").Range("F7") = txt_crashfilepath.Value
Call checkBAblanks

End Sub

Private Sub UserForm_Activate()

Dim workingdirectory As String
workingdirectory = ActiveWorkbook.path

txt_crashfilepath = ""
txt_analysisfilepath = ""
txt_aadtfilepath = ""

cmd_createBAinput.Visible = False

End Sub
