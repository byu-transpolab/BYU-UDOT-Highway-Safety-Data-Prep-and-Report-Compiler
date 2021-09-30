VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub cmd_openHome_Click()

Application.ScreenUpdating = False

'Clear contents of all sheets
Sheets("Non-State Route Crashes").Cells.ClearContents
Sheets("Non-State Route Crashes").UsedRange

Sheets("Non-State Routes").Cells.ClearContents
Sheets("Non-State Routes").UsedRange

Sheets("Incomplete Data").Cells.ClearContents
Sheets("Incomplete Data").UsedRange

Sheets("Ramp Crashes").Cells.ClearContents
Sheets("Ramp Crashes").UsedRange

form_guistart.Show

Application.ScreenUpdating = True

End Sub

Private Sub cmd_startPMvariable_Click()

form_ucpmvariable.Show

End Sub

Private Sub cmdClear_Click()

'Delete previous worksheets
Application.DisplayAlerts = False
For Each wksht In Worksheets
    If wksht.Name = "Dataset" Then
        wksht.Delete
    ElseIf wksht.Name = "AADT" Then
        wksht.Delete
    ElseIf wksht.Name = "Functional_Class" Then
        wksht.Delete
    ElseIf wksht.Name = "Thru_Lanes" Then
        wksht.Delete
    ElseIf wksht.Name = "Speed_Limit" Then
        wksht.Delete
    ElseIf wksht.Name = "Urban_Code" Then
        wksht.Delete
    ElseIf wksht.Name = "Intersections" Then
        wksht.Delete
    ElseIf wksht.Name = "Parameters" Then
        wksht.Delete
    ElseIf wksht.Name = "UICPMinput" Then
        wksht.Delete
    ElseIf wksht.Name = "UCPSMinput" Then
        wksht.Delete
    ElseIf wksht.Name = "CrashInput" Then
        wksht.Delete
    ElseIf wksht.Name = "Pavement_Messages" Then
        wksht.Delete
    ElseIf wksht.Name = "Roadway Data" Then
        wksht.Delete
    ElseIf wksht.Name = "Crash Data" Then
        wksht.Delete
    ElseIf wksht.Name = "Location" Then
        wksht.Delete
    ElseIf wksht.Name = "Crash" Then
        wksht.Delete
    ElseIf wksht.Name = "Rollup" Then
        wksht.Delete
    ElseIf wksht.Name = "Vehicle" Then
        wksht.Delete
    ElseIf wksht.Name = "Combined_Crash" Then
        wksht.Delete
    ElseIf wksht.Name = "HitList" Then
        wksht.Delete
    ElseIf wksht.Name = "CAMSinput" Then
        wksht.Delete
    End If
Next wksht
Application.DisplayAlerts = True

'Clear contents of all sheets
Sheets("Non-State Route Crashes").Cells.ClearContents
Sheets("Non-State Route Crashes").UsedRange

Sheets("Non-State Routes").Cells.ClearContents
Sheets("Non-State Routes").UsedRange

Sheets("Incomplete Data").Cells.ClearContents
Sheets("Incomplete Data").UsedRange

Sheets("Ramp Crashes").Cells.ClearContents
Sheets("Ramp Crashes").UsedRange

End Sub
