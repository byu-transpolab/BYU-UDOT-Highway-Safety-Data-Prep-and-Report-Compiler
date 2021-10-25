VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorrectHeaders 
   Caption         =   "Correct Headers"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6705
   OleObjectBlob   =   "frmCorrectHeaders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCorrectHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frmCorrectHeaders:
'   (1) Asks user to choose header from incoming list that corresponds best to expected header.
'
' Created by: Josh Gibbons, BYU, 2016
' Modified by: Josh Gibbons, BYU, 2016
' Commented by: Josh Gibbons, BYU, 2016


Private Sub cboChooseHeader_Change()
    'Define variables
    Dim irow, iCol As Integer
    Dim header, FileName1, Description As String
    Dim DExample As Variant
    
    'Assign values to variables
    header = cboChooseHeader.Value
    ThisWorkbook.Activate
    FileName1 = Worksheets("OtherData").Range("AY5")
    'Find example data from header
    irow = 1
    iCol = 1
    Do While Workbooks(FileName1).Worksheets(1).Cells(irow, iCol) <> ""
        If Workbooks(FileName1).Worksheets(1).Cells(irow, iCol) = header Then
            DExample = Workbooks(FileName1).Worksheets(1).Cells(irow + 1, iCol)
            Exit Do
        ElseIf InStr(1, Workbooks(FileName1).Worksheets(1).Cells(irow, iCol), "AADT") > 0 And _
        InStr(1, header, "AADT") > 0 Then
            DExample = Workbooks(FileName1).Worksheets(1).Cells(irow + 1, iCol)
            Exit Do
        End If
        iCol = iCol + 1
    Loop
    
    'Enter data example into lblDataExample so the user knows what kind of data is with a certain header
    lblDataExample = DExample
    
End Sub

Private Sub cmdNotAvail_Click()
    'Declare variables
    Dim strNeeded, strHeader, strHeaderDesc, FileName1 As String
    Dim BtnChoice As Variant
    Dim row1, numcol, numrow As Integer
    
    'Enter values for variables from values on OtherData sheet
    strNeeded = Worksheets("OtherData").Range("BB4")
    strHeader = Worksheets("OtherData").Range("AX4")
    strHeaderDesc = Worksheets("OtherData").Range("AY4")
    FileName1 = Worksheets("OtherData").Range("AY5")
    numcol = Worksheets("OtherData").Range("AZ4")
    numrow = Worksheets("OtherData").Range("BA4")
    
    'Hide user form
    frmCorrectHeaders.Hide
    
    'If the column of data is needed, then the user must obtain the data before continuing.
    If strNeeded = "YES" Then
        Worksheets("Home").Activate
        MsgBox "The Header that is not available is necessary in order to prepare the data for the model." & _
        vbCrLf & "Please obtain the data before proceeding.", , "Obtain Data Before Continuing"
        Workbooks(FileName1).Close False
        End
    'If the column of data is not necessary for the model, the user will be given a choice whether to continue or not.
    ElseIf strNeeded = "NO" Or strNeeded = "NOT USED" Then
        BtnChoice = MsgBox("The Header that is not available is NOT necessary in order to prepare the data for the model." & _
        " However, it is one less variable to be used in the model. If you want to include this data in the model," & _
        " please obtain the correct data file before proceeding with data preparation." & vbCrLf & vbCrLf & _
        "Header Description:" & vbCrLf & strHeader & vbCrLf & vbCrLf & _
        "Do you wish to continue data preparation without this data?", vbYesNo, "Data is NOT Necessary")
        
        If BtnChoice = vbYes Then
            row1 = 4
            Do Until Worksheets("OtherData").Cells(row1, numcol) = strHeader
               row1 = row1 + 1
            Loop
            Worksheets("OtherData").Cells(row1, numcol + 3) = "NOT USED"                                                'FLAGGED: I worry about this. What if they were wrong not to include it? should we say "no" instead?
        ElseIf BtnChoice = vbNo Then
            Worksheets("Home").Activate
            Workbooks(FileName1).Close False
            
            row1 = 4
            Do Until Worksheets("OtherData").Cells(row1, numcol) = strHeader
               row1 = row1 + 1
            Loop
            Worksheets("OtherData").Cells(row1, numcol + 3) = "NO"
            
            MsgBox "The process was canceled. Please obtain the needed data before proceeding.", , "Process Cancelled"
             
            End
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    'Declare variables
    Dim Head1, Head2 As String
    Dim row1, col1 As Integer
    
    'Give values to variables
    Head1 = Worksheets("OtherData").Range("AX4")
    Head2 = cboChooseHeader.Value
    row1 = 4
    col1 = Worksheets("OtherData").Range("AZ4")
    
    'Hide userform
    frmCorrectHeaders.Hide
    
    'Replace previous expected header with the selected header in the userform
    Do Until Worksheets("OtherData").Cells(row1, col1) = ""
        If Worksheets("OtherData").Cells(row1, col1) = Head1 Then
            Worksheets("OtherData").Cells(row1, col1) = Head2
            Exit Do
        End If
        row1 = row1 + 1
    Loop
End Sub

Private Sub cmdWrongFile_Click()
'Declare variables
Dim FileName1, DataSet As String

'Assign values to variables
FileName1 = lblChosenFileName.Caption
DataSet = lblDataSet.Caption

'Close previous file
Windows(FileName1).Close False

'Hide UserForm
frmCorrectHeaders.Hide

'Begin the correct process based on the working dataset
If DataSet = "Historic AADT" Then
    Worksheets("Home").cmdAADT.Value = True
    End
ElseIf DataSet = "Functional Class" Then
    Worksheets("Home").cmdFClass.Value = True
    End
ElseIf DataSet = "Speed Limit" Then
    Worksheets("Home").cmdSpeedLimit.Value = True
    End
ElseIf DataSet = "Thru Lanes" Then
    Worksheets("Home").cmdLanes.Value = True
    End
ElseIf DataSet = "Urban Code" Then
    Worksheets("Home").cmdUCode.Value = True
    End
ElseIf DataSet = "Sign Faces" Then
    Worksheets("Home").cmdSignFaces.Value = True
    End
ElseIf DataSet = "Crash Location" Then
    Worksheets("Home").cmdLocation.Value = True
    End
ElseIf DataSet = "(General) Crash Data" Then
    Worksheets("Home").cmdCrash.Value = True
    End
ElseIf DataSet = "Crash Rollup" Then
    Worksheets("Home").cmdRollup.Value = True
    End
ElseIf DataSet = "Crash Vehicle" Then
    Worksheets("Home").cmdVehicle.Value = True
    End
End If


End Sub


Private Sub UserForm_Activate()
    'Define variables
    Dim numRows, col1 As Integer
    Dim header, Description, DExample, FileName1 As String
    
    'Assign values to variables
    header = Worksheets("OtherData").Range("AX4")
    Description = Worksheets("OtherData").Range("AY4")
    numRows = Worksheets("OtherData").Range("BA4")
    col1 = Worksheets("OtherData").Range("AZ4")
    FileName1 = Worksheets("OtherData").Range("AY5")
    
    lblExpectedHeader.Caption = header
    lblHeaderDescription.Caption = Description
    lblChosenFileName.Caption = FileName1
    cboChooseHeader.RowSource = "OtherData!AO4:AO" & 3 + numRows
    
    If col1 = 1 Then
        lblDataSet.Caption = "Historic AADT"
    ElseIf col1 = 5 Then
        lblDataSet.Caption = "Functional Class"
    ElseIf col1 = 9 Then
        lblDataSet.Caption = "Speed Limit"
    ElseIf col1 = 13 Then
        lblDataSet.Caption = "Thru Lanes"
    ElseIf col1 = 17 Then
        lblDataSet.Caption = "Urban Code"
    ElseIf col1 = 21 Then
        lblDataSet.Caption = "Sign Faces"
    ElseIf col1 = 25 Then
        lblDataSet.Caption = "Crash Location"
    ElseIf col1 = 29 Then
        lblDataSet.Caption = "(General) Crash Data"
    ElseIf col1 = 33 Then
        lblDataSet.Caption = "Crash Rollup"
    ElseIf col1 = 37 Then
        lblDataSet.Caption = "Crash Vehicle"
    End If
End Sub

Private Sub UserForm_Terminate()
'Declare variables
Dim FileName1 As String

'Assign filename to FileName1 variable
FileName1 = lblChosenFileName.Caption

'Close user form
frmCorrectHeaders.Hide

'Close working file
Workbooks(FileName1).Close False

'Cancel process
MsgBox "Process canceled.", , "Canceled"
End

End Sub
