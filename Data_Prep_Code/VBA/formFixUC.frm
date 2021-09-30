VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formFixUC 
   Caption         =   "Fill in missing Urban Code information"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
   OleObjectBlob   =   "formFixUC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formFixUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGood2Go_Click()

Sheets("OtherData").Cells(8, 99) = cbo_UA.ListIndex + 4
formFixUC.Hide

End Sub

Private Sub cmdIDK_Click()

formFixUC.Hide

End Sub

Private Sub UserForm_Activate()

lblUCstatus.Caption = Sheets("OtherData").Cells(4, 99) 'get UC Status (blank or unknown)
lblCity.Caption = Sheets("OtherData").Cells(5, 99) 'get city name
lblLat.Caption = Sheets("OtherData").Cells(6, 99) 'get latitude
lblLong.Caption = Sheets("OtherData").Cells(7, 99) 'get longitude

End Sub

