VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCAMSSelection 
   Caption         =   "Segment Selection"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "frmCAMSSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCAMSSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()

MsgBox "Macros aborted.", , "Aborted"

End

End Sub

Private Sub cmdOK_Click()

Dim Data As Integer
Dim i As Integer
Dim j As Integer
Dim NumSeg As Integer
Dim ColName1 As String
Dim ColName2 As String
Dim ColNum1 As Integer
Dim ColNum2 As Integer
Dim IntIDCol As Integer
Dim IdentCol As Integer
Dim numRow As Integer
Dim numCol As Integer
Dim row1 As Integer

'Check which sorting option was selected and make sure selections were valid.
j = 0
If optCounty = True Then
    Data = 3
    For i = 0 To lst_1.ListCount - 1
        If lst_1.Selected(i) = True Then
            j = j + 1
        End If
    Next i
    If j = 0 Then
        MsgBox "Please select at least one county name before proceeding.", , "Make Selection"
        Exit Sub
    End If
    
    If IsNumeric(txtNumSeg.Text) = False Then
        MsgBox "Please enter a valid number of segments value before continuing."
        Exit Sub
    End If
    
ElseIf optState = True Then
    Data = 1
    
    If IsNumeric(txtNumSeg.Text) = False Then
        MsgBox "Please enter a valid number of segments value before continuing."
        Exit Sub
    End If
    
ElseIf optRegion = True Then
    Data = 2
    For i = 0 To lst_1.ListCount - 1
        If lst_1.Selected(i) = True Then
            j = j + 1
        End If
    Next i
    If j = 0 Then
        MsgBox "Please select at least one region before proceeding.", , "Make Selection"
        Exit Sub
    End If
    
    If IsNumeric(txtNumSeg.Text) = False Then
        MsgBox "Please enter a valid number of segments value before continuing."
        Exit Sub
    End If
    
ElseIf optIndInt = True Then
    Data = 4
    For i = 0 To lst_1.ListCount - 1
        If lst_1.Selected(i) = True Then
            j = j + 1
        End If
    Next i
    If j = 0 Then
        MsgBox "Please select at least one segment before proceeding.", , "Make Selection"
        Exit Sub
    End If

Else
    MsgBox "Please select a segment sorting option.", , "Select Option"
    Exit Sub
End If

'Hide user form
frmCAMSSelection.Hide

'List selection on SegKey sheet
If Data = 1 Then            'State
    
    NumSeg = txtNumSeg.Value
    
    Sheets("SegKey").Cells(2, 5).Value = "State"
    
ElseIf Data = 2 Then        'Region
    
    NumSeg = txtNumSeg.Value
    
    ReDim Selection(1 To j) As String
    j = 1
    For i = 0 To lst_1.ListCount - 1
        If lst_1.Selected(i) = True Then
            Selection(j) = Left(lst_1.List(i), InStr(1, lst_1.List(i), "(") - 2)
            j = j + 1
        End If
    Next i
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 5).Value = Selection(i)
    Next i
    
ElseIf Data = 3 Then        'County
    
    NumSeg = txtNumSeg.Value
    
    ReDim Selection(1 To j) As String
    j = 1
    For i = 0 To lst_1.ListCount - 1
        If lst_1.Selected(i) = True Then
            Selection(j) = Left(lst_1.List(i), InStr(1, lst_1.List(i), "(") - 2)
            j = j + 1
        End If
    Next i
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 5).Value = Selection(i)
    Next i
    
ElseIf Data = 4 Then        'Individual Segments
    
    ReDim Selection(1 To j) As String
    j = 1
    For i = 0 To lst_1.ListCount - 1
        If lst_1.Selected(i) = True Then
            Selection(j) = lst_1.List(i)
            j = j + 1
        End If
    Next i
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 5).Value = Selection(i)
    Next i
    
End If


'List Segments on SegKey sheet related to the selection
Sheets("Results").Activate
numRow = 1
numCol = 1

Do Until Cells(numRow + 1, numCol) = ""
    numRow = numRow + 1
Loop

Do Until Cells(1, numCol + 1) = ""
    numCol = numCol + 1
Loop

IntIDCol = FindColumn("SEG_ID")

If Data = 1 Then                'State
    ReDim Segments(1 To NumSeg) As String
    
    ColName1 = "State_Rank"
    ColNum1 = FindColumn(ColName1)
    
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Add Key:=Range( _
        Col_Letter(ColNum1) & "2:" & Col_Letter(ColNum1) & CStr(numRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Results").Sort
        .SetRange Range("A1:" & Col_Letter(numCol) & CStr(numRow))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    row1 = 2
    j = 1
    Do Until Cells(row1, ColNum1) > NumSeg
        Segments(j) = Cells(row1, IntIDCol)
        j = j + 1
        row1 = row1 + 1
    Loop
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 6).Value = Segments(i)
    Next i
    
ElseIf Data = 2 Then            'Region
    ReDim Segments(1 To NumSeg * (j - 1)) As String
    
    ColName1 = "Region_Rank"
    ColName2 = "REGION"
    ColNum1 = FindColumn(ColName1)
    ColNum2 = FindColumn(ColName2)
    
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Add Key:=Range( _
        Col_Letter(ColNum2) & "2:" & Col_Letter(ColNum2) & CStr(numRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Add Key:=Range( _
        Col_Letter(ColNum1) & "2:" & Col_Letter(ColNum1) & CStr(numRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Results").Sort
        .SetRange Range("A1:" & Col_Letter(numCol) & CStr(numRow))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    row1 = 2
    i = 1
    j = 1
    Do Until Cells(row1, ColNum2) = ""
        If CStr(Cells(row1, ColNum2)) = Selection(i) And Cells(row1, ColNum1) > 0 And Cells(row1, ColNum1) < NumSeg + 1 Then
            Segments(j) = Cells(row1, IntIDCol)
            j = j + 1
            row1 = row1 + 1
        ElseIf Cells(row1, ColNum2) < Selection(i) Then
            Do Until Cells(row1, ColNum2) = Selection(i)
                row1 = row1 + 1
            Loop
        ElseIf Cells(row1, ColNum2) > Selection(i) Then
            i = i + 1
        ElseIf Cells(row1, ColNum2) = Selection(i) And Selection(i) = "4" Then
            Exit Do
        ElseIf Cells(row1, ColNum2) = Selection(i) Then
            Do Until Cells(row1, ColNum2) = Selection(i + 1)
                row1 = row1 + 1
            Loop
            i = i + 1
        End If
    Loop
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 6).Value = Segments(i)
    Next i
    
ElseIf Data = 3 Then            'County
    ReDim Segments(1 To NumSeg * (j - 1)) As String
    
    ColName1 = "County_Rank"
    ColName2 = "COUNTY"
    ColNum1 = FindColumn(ColName1)
    ColNum2 = FindColumn(ColName2)
    
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Add Key:=Range( _
        Col_Letter(ColNum2) & "2:" & Col_Letter(ColNum2) & CStr(numRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Add Key:=Range( _
        Col_Letter(ColNum1) & "2:" & Col_Letter(ColNum1) & CStr(numRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Results").Sort
        .SetRange Range("A1:" & Col_Letter(numCol) & CStr(numRow))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    row1 = 2
    i = 1
    j = 1
    Do Until Cells(row1, ColNum2) = ""
        If Cells(row1, ColNum2) = Selection(i) And Cells(row1, ColNum1) > 0 And Cells(row1, ColNum1) < NumSeg + 1 Then
            Segments(j) = Cells(row1, IntIDCol)
            j = j + 1
            row1 = row1 + 1
        ElseIf StrComp(Cells(row1, ColNum2), Selection(i)) = -1 Then
            Do Until Cells(row1, ColNum2) = Selection(i)
                row1 = row1 + 1
            Loop
        ElseIf StrComp(Cells(row1, ColNum2), Selection(i)) = 1 Then
            i = i + 1
        ElseIf Cells(row1, ColNum2) = Selection(i) And Selection(i) = "WEBER" Then
            Exit Do
        ElseIf StrComp(Cells(row1, ColNum2), Selection(i)) = 0 Then
            Do Until Cells(row1, ColNum2) = Selection(i + 1)
                row1 = row1 + 1
            Loop
            i = i + 1
        End If
    Loop
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 6).Value = Segments(i)
    Next i
        
ElseIf Data = 4 Then            'Ind. Segments
    
    For i = 1 To j - 1
        Sheets("SegKey").Cells(1 + i, 6).Value = CInt(Mid(Selection(i), 5, 3))
    Next i
    
End If


'Create segment reports



End Sub


Private Sub cmdSelectAll_Click()

Dim i

For i = 0 To lst_1.ListCount - 1
    lst_1.Selected(i) = True
Next i

End Sub

Private Sub cmdSelectNone_Click()
    
    Dim i
    
    For i = 0 To lst_1.ListCount - 1
        lst_1.Selected(i) = False
    Next i
    
End Sub

Private Sub lblNumSeg1_Click()

End Sub

Private Sub lblNumSeg2_Click()

End Sub

Private Sub lst_1_Click()

End Sub

Private Sub optCounty_Click()

Dim iRow As Integer

lblIntro2.Visible = True
cmdSelectAll.Visible = True
cmdSelectNone.Visible = True
lst_1.Visible = True
lblNumSeg1.Visible = True
txtNumSeg.Visible = True
lblNumSeg2.Visible = True

lblIntro2.Caption = "Select the desired counties below. The # of segments represented from each county is listed in parentheses."
lblNumSeg2.Caption = "segments from each county."
txtNumSeg.Text = "3"

iRow = 2
Do Until Sheets("SegKey").Cells(iRow + 1, 3) = ""
    iRow = iRow + 1
Loop

lst_1.RowSource = "SegKey!C2:C" & iRow

End Sub

Private Sub optIndInt_Click()

Dim iRow As Integer

lblIntro2.Visible = True
cmdSelectAll.Visible = True
cmdSelectNone.Visible = True
lst_1.Visible = True
lblNumSeg1.Visible = False
txtNumSeg.Visible = False
lblNumSeg2.Visible = False

lblIntro2.Caption = "Select the desired segments of interest below. Each segment is listed by model segment ID and route numbers."
txtNumSeg.Text = ""

iRow = 2
Do Until Sheets("SegKey").Cells(iRow + 1, 4) = ""
    iRow = iRow + 1
Loop

lst_1.RowSource = "SegKey!D2:D" & iRow

End Sub

Private Sub optRegion_Click()

Dim iRow As Integer

lblIntro2.Visible = True
cmdSelectAll.Visible = True
cmdSelectNone.Visible = True
lst_1.Visible = True
lblNumSeg1.Visible = True
txtNumSeg.Visible = True
lblNumSeg2.Visible = True

lblIntro2.Caption = "Select the desired regions below. The # of segments represented from each region is listed in parentheses."
lblNumSeg2.Caption = "segments from each region."
txtNumSeg.Text = "10"

iRow = 2
Do Until Sheets("SegKey").Cells(iRow + 1, 1) = ""
    iRow = iRow + 1
Loop

lst_1.RowSource = "SegKey!A2:A" & iRow

End Sub

Private Sub optState_Click()

lblIntro2.Visible = False
cmdSelectAll.Visible = False
cmdSelectNone.Visible = False
lst_1.Visible = False
lblNumSeg1.Visible = True
txtNumSeg.Visible = True
lblNumSeg2.Visible = True

txtNumSeg.Text = "20"

lblNumSeg2.Caption = "segments in the state."

End Sub

Private Sub UserForm_Activate()

lblIntro2.Visible = False
cmdSelectAll.Visible = False
cmdSelectNone.Visible = False
lst_1.Visible = False
lblNumSeg1.Visible = False
txtNumSeg.Visible = False
lblNumSeg2.Visible = False

End Sub

Function Col_Letter(lngCol As Integer) As String

Dim vArr
vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)

End Function

Function FindColumn(cname As String) As Integer

FindColumn = 1
Do Until Cells(1, FindColumn) = cname
    FindColumn = FindColumn + 1
Loop

End Function

