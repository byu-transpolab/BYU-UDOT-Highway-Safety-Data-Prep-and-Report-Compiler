VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_FAUrbanCode 
   Caption         =   "Functional Area - Urban Code"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4665
   OleObjectBlob   =   "form_FAUrbanCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_FAUrbanCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblFAEnter_Click()

End Sub

Private Sub lblUrbanCode_Click()

End Sub

Private Sub UserForm_Activate()

txtFARural = 500
txtFASuburban = 500
txtFAUrban = 250

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

Sheets("Inputs").Cells(rowFA, colUICPM + 1) = "Urban Code"
Sheets("Inputs").Cells(rowFA + 1, colUICPM) = "Urban Code"
Sheets("Inputs").Cells(rowFA + 1, colUICPM + 1) = "Functional Area"

Sheets("Inputs").Activate
Sheets("Inputs").Range(Cells(rowFA + 2, colUICPM), Cells(rowFA + 20, colUICPM + 1)).ClearContents
Sheets("Home").Activate


Sheets("Inputs").Cells(rowFA + 2, colUICPM) = "Rural"
Sheets("Inputs").Cells(rowFA + 3, colUICPM) = "Small Urban"
Sheets("Inputs").Cells(rowFA + 4, colUICPM) = "Urban"


Sheets("Inputs").Cells(rowFA + 2, colUICPM + 1) = txtFARural.Value
Sheets("Inputs").Cells(rowFA + 3, colUICPM + 1) = txtFASuburban.Value
Sheets("Inputs").Cells(rowFA + 4, colUICPM + 1) = txtFAUrban.Value


form_FAUrbanCode.Hide


Application.ScreenUpdating = True

End Sub

