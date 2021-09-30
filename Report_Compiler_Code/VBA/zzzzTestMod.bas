Attribute VB_Name = "zzzzTestMod"
Sub windowtest()
Dim modeltype As String
Dim dataranges As String

modeltype = InputBox("Enter the model used. (UCPM or UCSM)", "Enter the model used", "UCPM/UCSM")
If modeltype = "" Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If

dataranges = InputBox("Enter the range of dates for data source." & Chr(10) & "Example: [2008-2012]", "Enter the range of dates for data source", "20X1-20X5")
If dataranges = "" Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If


End Sub


Sub testopen()

Dim cmdline As String

reportfp = "J:\groups\udot2015\3 Post-Model Data Analysis\Analysis Reports\Output 5-4 13-9-38"

cmdline = "explorer.exe" & " " & reportfp

Shell cmdline, vbNormalFocus


End Sub

