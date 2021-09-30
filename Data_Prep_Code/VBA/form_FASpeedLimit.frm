VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_FASpeedLimit 
   Caption         =   "Functional Area - Speed Limit"
   ClientHeight    =   7548
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "form_FASpeedLimit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_FASpeedLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

If txtFA20.Value = "" Or txtFA25.Value = "" Or txtFA30.Value = "" Or txtFA35.Value = "" Or _
txtFA40.Value = "" Or txtFA45.Value = "" Or txtFA50.Value = "" Or txtFA55.Value = "" Or _
txtFA60.Value = "" Or txtFA65.Value = "" Or txtFA70.Value = "" Or txtFA75.Value = "" Then
    MsgBox "Value left blank. Please enter a value for the fixed length."
    Exit Sub
End If


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

Sheets("Inputs").Cells(rowFA, colUICPM + 1) = "Speed Limit"
Sheets("Inputs").Cells(rowFA + 1, colUICPM) = "Speed Limit"
Sheets("Inputs").Cells(rowFA + 1, colUICPM + 1) = "Functional Area"

Sheets("Inputs").Activate
Sheets("Inputs").Range(Cells(rowFA + 2, colUICPM), Cells(rowFA + 20, colUICPM + 1)).ClearContents
Sheets("Home").Activate

For i = 1 To 12
    Sheets("Inputs").Cells(rowFA + 1 + i, colUICPM) = 15 + (i * 5)
Next i

Sheets("Inputs").Cells(rowFA + 2, colUICPM + 1) = txtFA20.Value
Sheets("Inputs").Cells(rowFA + 3, colUICPM + 1) = txtFA25.Value
Sheets("Inputs").Cells(rowFA + 4, colUICPM + 1) = txtFA30.Value
Sheets("Inputs").Cells(rowFA + 5, colUICPM + 1) = txtFA35.Value
Sheets("Inputs").Cells(rowFA + 6, colUICPM + 1) = txtFA40.Value
Sheets("Inputs").Cells(rowFA + 7, colUICPM + 1) = txtFA45.Value
Sheets("Inputs").Cells(rowFA + 8, colUICPM + 1) = txtFA50.Value
Sheets("Inputs").Cells(rowFA + 9, colUICPM + 1) = txtFA55.Value
Sheets("Inputs").Cells(rowFA + 10, colUICPM + 1) = txtFA60.Value
Sheets("Inputs").Cells(rowFA + 11, colUICPM + 1) = txtFA65.Value
Sheets("Inputs").Cells(rowFA + 12, colUICPM + 1) = txtFA70.Value
Sheets("Inputs").Cells(rowFA + 13, colUICPM + 1) = txtFA75.Value

form_FASpeedLimit.Hide

Application.ScreenUpdating = True

End Sub

Private Sub UserForm_Activate()

Dim colFASpeed As Long


colFASpeed = 1
Do Until Sheets("Key").Cells(1, colFASpeed) = "Functional Area"
    colFASpeed = colFASpeed + 1
Loop


txtFA20 = Sheets("Key").Cells(3, colFASpeed + 4)
txtFA25 = Sheets("Key").Cells(4, colFASpeed + 4)
txtFA30 = Sheets("Key").Cells(5, colFASpeed + 4)
txtFA35 = Sheets("Key").Cells(6, colFASpeed + 4)
txtFA40 = Sheets("Key").Cells(7, colFASpeed + 4)
txtFA45 = Sheets("Key").Cells(8, colFASpeed + 4)
txtFA50 = Sheets("Key").Cells(9, colFASpeed + 4)
txtFA55 = Sheets("Key").Cells(10, colFASpeed + 4)
txtFA60 = Sheets("Key").Cells(11, colFASpeed + 4)
txtFA65 = Sheets("Key").Cells(12, colFASpeed + 4)
txtFA70 = Sheets("Key").Cells(13, colFASpeed + 4)
txtFA75 = Sheets("Key").Cells(14, colFASpeed + 4)


End Sub

