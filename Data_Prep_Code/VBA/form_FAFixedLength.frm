VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_FAFixedLength 
   Caption         =   "Functional Area - Fixed Length"
   ClientHeight    =   1812
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   OleObjectBlob   =   "form_FAFixedLength.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_FAFixedLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Dim colUICPM, rowFA, i As Integer

If txtFAFixedLength.Value = "" Then
    MsgBox "Value left blank. Please enter a value for the fixed length."
    Exit Sub
End If

colUICPM = 1
rowFA = 1

Application.ScreenUpdating = False

Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "UICPM"
    colUICPM = colUICPM + 1
Loop

Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "Selected FA Parameter"
    rowFA = rowFA + 1
Loop

Sheets("Inputs").Cells(rowFA, colUICPM + 1) = "Fixed Length"
Sheets("Inputs").Cells(rowFA + 1, colUICPM) = "Fixed Length"
Sheets("Inputs").Cells(rowFA + 1, colUICPM + 1) = "Functional Area"

Sheets("Inputs").Activate
Sheets("Inputs").Range(Cells(rowFA + 2, colUICPM), Cells(rowFA + 20, colUICPM + 1)).ClearContents
Sheets("Home").Activate

Sheets("Inputs").Cells(rowFA + 2, colUICPM) = txtFAFixedLength.Value
Sheets("Inputs").Cells(rowFA + 2, colUICPM + 1) = txtFAFixedLength.Value

form_FAFixedLength.Hide

Application.ScreenUpdating = True

End Sub

Private Sub UserForm_Activate()

txtFAFixedLength.Value = 250

End Sub


