VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

'Export to Github
Test_MS_Scripting_Runtime
ExportVisualBasicCode

Dim wksht As Worksheet

Sheets("Main").Activate

If ActiveWorkbook.Name = "Report Compiler v2" Then
    For Each wksht In Worksheets
        If wksht.Name = "Results" Then
            wksht.Delete
        ElseIf wksht.Name = "Parameters" Then
            wksht.Delete
        End If
    Next
End If

End Sub

