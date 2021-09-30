VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCorrectHeaders 
   Caption         =   "Correct Headers"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6705
   OleObjectBlob   =   "formCorrectHeaders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCorrectHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'frmCorrectHeaders:
'   (1) Asks user to choose header from incoming list that corresponds best to expected header.
'
' Created by: Josh Gibbons, BYU, 2016
' Modified by: Sam Mineer, BYU, 2016
' Commented by: Sam Mineer, BYU, 2016


Private Sub cboChooseHeader_Change()
    'Define variables
    Dim irow, iCol As Integer
    Dim header, datasheet, datawb, Description As String
    Dim DExample As Variant
    
    'Assign values to variables
    header = cboChooseHeader.Value
    datasheet = Worksheets("Key").Range("BB5")
    datawb = Worksheets("Key").Range("BB6")
    
    'Find example data from header
    irow = 1
    iCol = 1
    DExample = ""
    Do While Workbooks(datawb).Worksheets(datasheet).Cells(irow, iCol) <> ""
        If Workbooks(datawb).Worksheets(datasheet).Cells(irow, iCol) = header Then
            Do Until Len(DExample) > 0
                DExample = Workbooks(datawb).Worksheets(datasheet).Cells(irow + 1, iCol)
                irow = irow + 1
            Loop
            DExample = Workbooks(datawb).Worksheets(datasheet).Cells(irow + 1, iCol)
            Exit Do
        ElseIf InStr(1, Workbooks(datawb).Worksheets(datasheet).Cells(irow, iCol), "AADT") > 0 And InStr(1, header, "AADT") > 0 Then
            Do Until Len(DExample) > 0
                DExample = Workbooks(datawb).Worksheets(datasheet).Cells(irow + 1, iCol)
                irow = irow + 1
            Loop
            DExample = Workbooks(datawb).Worksheets(datasheet).Cells(irow + 1, iCol)
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
    strNeeded = Worksheets("Key").Range("BE4")
    strHeader = Worksheets("Key").Range("BA4")
    strHeaderDesc = Worksheets("Key").Range("BB4")
    FileName1 = Worksheets("Key").Range("BB5")
    numcol = Worksheets("Key").Range("BC4")
    numrow = Worksheets("Key").Range("BD4")
    
    'Hide user form
    formCorrectHeaders.Hide
    
    'If the column of data is needed, then the user must obtain the data before continuing.
    If strNeeded = "YES" Then
        MsgBox "The Header that is not available is necessary in order to summarize the roadway and crash data for the Roaday Safety Analysis Reports." & Chr(10) & _
        "Please obtain the data before proceeding.", , "Obtain Data Before Continuing"
        Worksheets("Main").Activate
        'Workbooks(FileName1).Close False
        End
    'If the column of data is not necessary for the model, the user will be given a choice whether to continue or not.
    'ElseIf strNeeded = "NO" Or strNeeded = "NOT USED" Then
    '    BtnChoice = MsgBox("The Header that is not available is NOT necessary in order to prepare the data for the model." & _
    '    " However, it is one less variable to be used in the model. If you want to include this data in the model," & _
    '    " please obtain the correct data file before proceeding with data preparation." & vbCrLf & vbCrLf & _
    '    "Header Description:" & vbCrLf & strHeader & vbCrLf & vbCrLf & _
    '    "Do you wish to continue data preparation without this data?", vbYesNo, "Data is NOT Necessary")
    '
    '    If BtnChoice = vbYes Then
    '        row1 = 4
    '        Do Until Worksheets("OtherData").Cells(row1, NumCol) = strHeader
    '           row1 = row1 + 1
    '        Loop
    '        Worksheets("OtherData").Cells(row1, NumCol + 3) = "NOT USED"
    '    ElseIf BtnChoice = vbNo Then
    '        Worksheets("Home").Activate
    '        Workbooks(FileName1).Close False
    '
    '        row1 = 4
    '        Do Until Worksheets("OtherData").Cells(row1, NumCol) = strHeader
    '           row1 = row1 + 1
    '        Loop
    '        Worksheets("OtherData").Cells(row1, NumCol + 3) = "NO"
    '
    '        MsgBox "The process was cancelled. Please obtain the needed data before proceeding.", , "Process Cancelled"
    '
    '        End
    '    End If
    End If
End Sub

Private Sub cmdOK_Click()
' When the OK command button is clicked, the selected header is updated in the list of critical data columns

    'Declare variables
    Dim Head1, Head2 As String
    Dim row1, col1 As Integer
    
    'Give values to variables
    Head1 = Worksheets("Key").Range("BA4")
    Head2 = cboChooseHeader.Value
    row1 = 4
    col1 = Worksheets("Key").Range("BC4")
    
    'Hide userform
    cboChooseHeader.Value = ""
    formCorrectHeaders.Hide
    
    'Replace previous expected header with the selected header in the userform
    Do Until Worksheets("Key").Cells(row1, col1) = ""
        If Worksheets("Key").Cells(row1, col1) = Head1 Then
            Worksheets("Key").Cells(row1, col1) = Head2
            Worksheets("Key").Cells(row1, col1 + 2) = Head2
            Exit Do
        End If
        row1 = row1 + 1
    Loop
End Sub

Private Sub cmdWrongFile_ClickDELETE()
'Not going to use this option
'Declare variables
'Dim FileName1, DataSet As String

'Assign values to variables
'FileName1 = lblChosenFileName.Caption
'DataSet = lblDataSet.Caption

'Close previous file
'Windows(FileName1).Close False

'Hide UserForm
'frmCorrectHeaders.Hide

'Begin the correct process based on the working dataset
'If DataSet = "Historic AADT" Then
'    Worksheets("Home").cmdAADT.Value = True
'    End
'ElseIf DataSet = "Functional Class" Then
'    Worksheets("Home").cmdFClass.Value = True
'    End
'ElseIf DataSet = "Speed Limit" Then
'    Worksheets("Home").cmdSpeedLimit.Value = True
'    End
'ElseIf DataSet = "Thru Lanes" Then
'    Worksheets("Home").cmdLanes.Value = True
'    End
'ElseIf DataSet = "Urban Code" Then
'    Worksheets("Home").cmdUCode.Value = True
'    End
'ElseIf DataSet = "Sign Faces" Then
'    Worksheets("Home").cmdSignFaces.Value = True
'    End
'ElseIf DataSet = "Crash Location" Then
'    Worksheets("Home").cmdLocation.Value = True
'    End
'ElseIf DataSet = "(General) Crash Data" Then
'    Worksheets("Home").cmdCrash.Value = True
 '   End
'ElseIf DataSet = "Crash Rollup" Then
'    Worksheets("Home").cmdRollup.Value = True
'    End
'ElseIf DataSet = "Crash Vehicle" Then
'    Worksheets("Home").cmdVehicle.Value = True
'    End
'End If


End Sub


Private Sub UserForm_Activate()
' When the user form is activated, the GUI is populated with the working dataset, selected file name, expected column header, header descriptions, and sets the source of the combo box for "Choose Header"

'Define variables
Dim row1, col1 As Integer
Dim header, Description, DExample, FileName1 As String

'Assign values to variables
header = Worksheets("Key").Range("BA4")
Description = Worksheets("Key").Range("BB4")
col1 = Worksheets("Key").Range("BC4")
row1 = Worksheets("Key").Range("BD4")
FileName1 = Worksheets("Key").Range("BB5")

'Update values in the GUI
lblExpectedHeader.Caption = header
lblHeaderDescription.Caption = Description
lblChosenFileName.Caption = FileName1
cboChooseHeader.RowSource = "Key!AX4:AX" & CStr(3 + row1)
lblDataSet.Caption = Worksheets("Key").Cells(2, col1).Value
    
End Sub

Private Sub UserForm_Terminated()
'Don't need to close the file

'Declare variables
'Dim FileName1 As String

'Assign filename to FileName1 variable
'FileName1 = lblChosenFileName.Caption

'Close user form
'frmCorrectHeaders.Hide

'Close working file
'Workbooks(FileName1).Close False

'Cancel process
MsgBox "Process canceled.", vbOKOnly, "Canceled"

End

End Sub
