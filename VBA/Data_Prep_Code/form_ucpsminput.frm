VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_ucpsminput 
   Caption         =   "Safety Statistical Analysis: UCPM & UCSM Input (R GUI)"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   OleObjectBlob   =   "form_ucpsminput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_ucpsminput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'R GUI workbook created for UDOT Roadway Safety Analysis Methodology
'Comments by Sam Mineer, Brigham Young Univerisity, June 2016

Private Sub chbx_severity5_Click()
    Call checkblanks
End Sub

Private Sub chbx_severity4_Click()
    Call checkblanks
End Sub

Private Sub chbx_severity3_Click()
    Call checkblanks
End Sub

Private Sub chbx_severity2_Click()
    Call checkblanks
End Sub

Private Sub chbx_severity1_Click()
    Call checkblanks
End Sub

Private Sub cmd_CreateData_Click()

form_ucpsminput.Hide
form_CreateSegData.Show

End Sub

Private Sub txt_crashfilepath_Change()

    ActiveWorkbook.Sheets("Inputs").Range("B6").Value = Replace(txt_crashfilepath, "\", "/")
    
    Call checkblanks
End Sub

Private Sub txt_segmentfilepath_Change()
    
    ActiveWorkbook.Sheets("Inputs").Range("B5").Value = Replace(txt_segmentfilepath, "\", "/")
    
    Call checkblanks
End Sub

Private Sub cmd_crashdata_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
FilePath = Application.GetOpenFilename(, , "Select Crash Data")

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_crashfilepath = ""
Else
    txt_crashfilepath = Replace(FilePath, "\", "/")
    
End If

End Sub

Private Sub cmd_createinputdata_Click()

' Define variables
Dim workingdirectory As String
Dim segmentfilepath As String
Dim crashfilepath As String
Dim severitylist As String

' Extract values from textbox and checkbox selection
segmentfilepath = txt_segmentfilepath
crashfilepath = txt_crashfilepath
severitylist = ""
If chbx_severity1 = True Then
    If severitylist = "" Then
        severitylist = "1"
    End If
End If
If chbx_severity2 = True Then
    If severitylist = "" Then
        severitylist = "2"
    Else
        severitylist = severitylist & "2"
    End If
End If
If chbx_severity3 = True Then
    If severitylist = "" Then
        severitylist = "3"
    Else
        severitylist = severitylist & "3"
    End If
End If
If chbx_severity4 = True Then
    If severitylist = "" Then
        severitylist = "4"
    Else
        severitylist = severitylist & "4"
    End If
End If
If chbx_severity5 = True Then
    If severitylist = "" Then
        severitylist = "5"
    Else
        severitylist = severitylist & "5"
    End If
End If

'print inputs to workbook for future information
ActiveWorkbook.Sheets("Inputs").Range("B5").Value = Replace(txt_segmentfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("B6").Value = Replace(txt_crashfilepath, "\", "/")
ActiveWorkbook.Sheets("Inputs").Range("B9").Value = severitylist

'hide user form
form_ucpsminput.Hide

'start process to count crash severity
Call UCPSMdataprep(segmentfilepath, crashfilepath, chbx_factorsummary, chbx_severity1, chbx_severity2, chbx_severity3, chbx_severity4, chbx_severity5)

End Sub

Private Sub cmd_segdata_Click()

' Define variables
Dim FilePath As Variant

' Obtain the file path from the pop-up window
FilePath = Application.GetOpenFilename(, , "Select Road Segment Data")

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_segmentfilepath = ""
Else
    txt_segmentfilepath = Replace(FilePath, "\", "/")
End If

End Sub

Sub checkblanks()

' Check if the information has been filled before allowing the user to continue
If FileExists(txt_segmentfilepath) And txt_segmentfilepath.Value <> "" And FileExists(txt_crashfilepath) And txt_crashfilepath.Value <> "" And (chbx_severity5.Value Or chbx_severity4.Value Or chbx_severity3.Value Or chbx_severity2.Value Or chbx_severity1.Value) Then
    cmd_createinputdata.Visible = True
    lbl_stop.Visible = False
Else
    cmd_createinputdata.Visible = False
    lbl_stop.Visible = True
End If

End Sub

Private Sub UserForm_Activate()

' Blank the values for the user form
If Sheets("Inputs").Range("B5") = "" Then
    txt_segmentfilepath.Value = ""
Else
    txt_segmentfilepath.Value = Sheets("Inputs").Range("B5")
End If

If Sheets("Inputs").Range("B6") = "" Then
    txt_crashfilepath.Value = ""
Else
    txt_crashfilepath.Value = Sheets("Inputs").Range("B6")
End If

chbx_severity5.Value = False
chbx_severity4.Value = False
chbx_severity3.Value = False
chbx_severity2.Value = False
chbx_severity1.Value = False
chbx_factorsummary.Value = False

' Hide the creat input data button, until the user fills in the necessary data
lbl_stop.Visible = True
cmd_createinputdata.Visible = False

End Sub

