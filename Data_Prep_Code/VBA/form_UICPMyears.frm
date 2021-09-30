VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_UICPMyears 
   Caption         =   "UICPM - Data Years"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "form_UICPMyears.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_UICPMyears"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

Dim minyear, maxyear As Integer

minyear = cboMinYear.Value
maxyear = cboMaxYear.Value

Sheets("Key").Columns(3).ClearContents

Sheets("Key").Cells(1, 3) = minyear
Sheets("Key").Cells(2, 3) = maxyear

Sheets("Parameters").Cells(3, 2) = CStr(minyear) & "-" & CStr(maxyear)

form_UICPMyears.Hide

UICPMdataprep2

End Sub

Private Sub UserForm_Click()

End Sub
