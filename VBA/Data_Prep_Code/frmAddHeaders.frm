VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddHeaders 
   Caption         =   "Crash Rollup Data Columns"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10185
   OleObjectBlob   =   "frmAddHeaders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frmAddHeaders:
'   (1) Asks user if more rollup data headers should be considered in the analysis.
'
' Created by: Josh Gibbons, BYU, 2016
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016

Private Sub cmdAdd_Click()

Dim lItem, row1, row2 As Long
Dim col1 As Integer
Dim HeaderDesc, header As String

row1 = 4
col1 = 33
'Find how many crash rollup headers there are
Do While Sheets("OtherData").Cells(row1, col1) <> ""
    row1 = row1 + 1
Loop

'List selected headers and prompt user to enter in a header description
For lItem = 0 To lstFoundHeaders.ListCount - 1
    If lstFoundHeaders.Selected(lItem) = True Then
        header = lstFoundHeaders.List(lItem)
        Sheets("OtherData").Cells(row1, col1) = header
        Sheets("OtherData").Cells(row1, col1 + 2) = header
        lstFoundHeaders.Selected(lItem) = False
        
Line40:
        HeaderDesc = InputBox("Enter a header description for the following header. " & _
        "The description will be used to help the user identify the correct header when this process is run again." _
        & vbCrLf & vbCrLf & "Header: " & header & vbCrLf, "Enter Header Description", "Name: Description")
        
        If HeaderDesc = vbNullString Or HeaderDesc = "" Then
            MsgBox "Please enter a value in for the header description."
            GoTo Line40
        End If
        
        Sheets("OtherData").Cells(row1, col1 + 3) = "NO"
        Sheets("OtherData").Cells(row1, col1 + 1) = HeaderDesc
        row1 = row1 + 1
    End If
Next

frmAddHeaders.Hide

End Sub

Private Sub cmdClearAll_Click()

Dim x As Integer
For x = 0 To lstFoundHeaders.ListCount - 1
    If lstFoundHeaders.Selected(x) = True Then
        lstFoundHeaders.Selected(x) = False
    End If
Next

End Sub

Private Sub cmdNone_Click()

frmAddHeaders.Hide

End Sub



Private Sub UserForm_Activate()
'Declare variables
Dim row1, row2, col1, col2 As Integer

'Find how many non-critical headers are listed
col1 = 41
row1 = 4
Do While Worksheets("OtherData").Cells(row1 + 1, col1 + 1) <> ""
    row1 = row1 + 1
Loop

lstFoundHeaders.RowSource = "OtherData!AP4:AP" & row1

'Find how many critical headers are listed
col2 = 33
row2 = 4
Do While Worksheets("OtherData").Cells(row2 + 1, col2) <> ""
    row2 = row2 + 1
Loop

lstCriticalHeaders.RowSource = "OtherData!AG4:AG" & row2

End Sub

