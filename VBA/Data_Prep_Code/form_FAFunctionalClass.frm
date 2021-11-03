VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_FAFunctionalClass 
   Caption         =   "Functional Area - Functional Class"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4665
   OleObjectBlob   =   "form_FAFunctionalClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_FAFunctionalClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

txtFAFreeway = 1045
txtFAPrincipalArt = 700
txtFAMinorArt = 550
txtFAMajorColl = 400


End Sub

Private Sub cmdOK_Click()

Dim colUICPM, rowFA, i As Integer

colUICPM = 1
rowFA = 1

Application.ScreenUpdating = False

Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "UICPM"
    colUICPM = colUICPM + 1
Loop

Do Until Sheets("Inputs").Cells(rowFA, colUICPM) = "Selected FA Parameter"
    rowFA = rowFA + 1
Loop

Sheets("Inputs").Cells(rowFA, colUICPM + 1) = "Functional Class"
Sheets("Inputs").Cells(rowFA + 1, colUICPM) = "Functional Class"
Sheets("Inputs").Cells(rowFA + 1, colUICPM + 1) = "Functional Area"

Sheets("Inputs").Activate
Sheets("Inputs").Range(Cells(rowFA + 2, colUICPM), Cells(rowFA + 20, colUICPM + 1)).ClearContents
Sheets("Home").Activate


Sheets("Inputs").Cells(rowFA + 2, colUICPM) = "Other Freeway & Expressway"
Sheets("Inputs").Cells(rowFA + 3, colUICPM) = "Other Principal Arterial"
Sheets("Inputs").Cells(rowFA + 4, colUICPM) = "Minor Arterial"
Sheets("Inputs").Cells(rowFA + 5, colUICPM) = "Major Collector"


Sheets("Inputs").Cells(rowFA + 2, colUICPM + 1) = txtFAFreeway.Value
Sheets("Inputs").Cells(rowFA + 3, colUICPM + 1) = txtFAPrincipalArt.Value
Sheets("Inputs").Cells(rowFA + 4, colUICPM + 1) = txtFAMinorArt.Value
Sheets("Inputs").Cells(rowFA + 5, colUICPM + 1) = txtFAMajorColl.Value


form_FAFunctionalClass.Hide


Application.ScreenUpdating = True

End Sub
