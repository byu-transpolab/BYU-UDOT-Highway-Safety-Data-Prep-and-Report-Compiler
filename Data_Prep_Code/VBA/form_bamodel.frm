VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_bamodel 
   Caption         =   "Safety Statistical Analysis: Before After Model Start (R GUI)"
   ClientHeight    =   3056
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   OleObjectBlob   =   "form_bamodel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_bamodel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_bacode_Click()

Dim FilePath As Variant
Dim sourcestring As String
Dim inputvariable As String

Dim cmdLine As String
Dim rcode As Variant
Dim sevinput As Variant
Dim inputfile As String
Dim guiwb As String

guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

' Obtain the file path from the pop-up window
    'MsgBox "Select the Before After Statistical R Code", vbOKOnly, "Select Before After Statistical R Code"
    FilePath = Application.GetOpenFilename(, , "Select Before After Statistical R Code")
    
    ' If the user doesn't select a file, then the box will be blank
    If FilePath = False Then
        txt_bainput = ""
        MsgBox "Cannot proceed without Before After Statical R Code."
        Exit Sub
    Else
        txt_bacodefilepath = replace(FilePath, "\", "/")
        Workbooks(guiwb).Sheets("Inputs").Range("F11").Value = replace(txt_bacodefilepath, "\", "/")
    End If

End Sub

Private Sub cmd_bainputfile_Click()

Dim FilePath As Variant
Dim sourcestring As String
Dim inputvariable As String

Dim cmdLine As String
Dim rcode As Variant
Dim sevinput As Variant
Dim inputfile As String
Dim guiwb As String

guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

' Obtain the file path from the pop-up window
    'MsgBox "Select the BAinput .csv file", vbOKOnly, "Select BAinput file"
    FilePath = Application.GetOpenFilename(, , "Select Before After Input File")
    
    ' If the user doesn't select a file, then the box will be blank
    If FilePath = False Then
        txt_bainput = ""
        'MsgBox "Cannot proceed without BA Input file."
        Exit Sub
    Else
        txt_bainput = replace(FilePath, "\", "/")
        Workbooks(guiwb).Sheets("Inputs").Range("F8").Value = replace(txt_bainput, "\", "/")
    End If

End Sub

Private Sub cmd_startBA_Click()

Dim usercheck As Variant
Dim rscript As String
Dim rcode As String
Dim bawd As String
Dim niter As Long
Dim nburn As Long
Dim datalocation As String
Dim xs As String

Dim itemIndex As Integer
Dim xsnum As String

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

usercheck = MsgBox("Are you sure you are ready to begin the Before After Analysis?", vbYesNo, "Ready?")

If usercheck = vbNo Then
    MsgBox "Analysis Aborted", vbOKOnly, "Try Again"
    Exit Sub
Else
End If

rscript = Workbooks(guiwb).Sheets("Inputs").Range("F3")
bawd = Workbooks(guiwb).Sheets("Inputs").Range("F2")
rcode = Workbooks(guiwb).Sheets("Inputs").Range("F11")
bawd = bawd & "/" & "BAanalysis_" & replace(Date, "/", "-") & "_" & replace(replace(Time, ":", "-"), " ", "_") ' & "/"
MkDir bawd
niter = Workbooks(guiwb).Sheets("Inputs").Range("F9")
nburn = Workbooks(guiwb).Sheets("Inputs").Range("F10")
datalocation = Workbooks(guiwb).Sheets("Inputs").Range("F8")
'datalocation = Mid(datalocation, InStr(1, datalocation, "\\UCPMinput"))
'datalocation = Replace(datalocation, "\\", "")

' close form
form_bamodel.Hide

' send to executeBA
Call executeBA(rscript, rcode, bawd, niter, nburn, datalocation)

End Sub

Private Sub txt_baburniterations_Change()

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

If txt_baburniterations = "" Then
Else
    If CDbl(txt_baburniterations.Value) > (0.101 * CDbl(txt_baiterations.Value)) Then
        MsgBox "Do not set Burn-in Iterations greater than 10% of number of iterations.", vbOKOnly, "Warning"
        txt_baburniterations.Value = CInt(txt_baiterations.Value) * 0.1
    End If
    Workbooks(guiwb).Sheets("Inputs").Range("F9") = CDbl(txt_baiterations)
    Workbooks(guiwb).Sheets("Inputs").Range("F10") = CDbl(txt_baburniterations)
End If

Call checkBAblanks

End Sub

Private Sub txt_bacodefilepath_Change()
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

Workbooks(guiwb).Sheets("Inputs").Range("F11") = txt_bacodefilepath
Call checkBAblanks

End Sub

Private Sub txt_bainput_Change()
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

If txt_bainput.Value <> Workbooks(guiwb).Sheets("Inputs").Range("F8") Then
    Workbooks(guiwb).Sheets("Inputs").Range("F8") = txt_bainput
End If
Call checkBAblanks

End Sub

Sub checkBAblanks()

If FileExists(txt_bainput) And FileExists(txt_bacodefilepath) And Len(txt_bainput) > 0 And Len(txt_bacodefilepath) > 0 And Len(txt_baiterations.Value) > 0 And Len(txt_baburniterations.Value) > 0 Then     'And txt_bacadefilepath <> "" Then
    cmd_startBA.Visible = True

Else
    cmd_startBA.Visible = False

End If

End Sub


Private Sub txt_baiterations_Change()

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

If txt_baiterations.Value = "" Then
Else
    txt_baburniterations.Value = txt_baiterations.Value * 0.1
    Workbooks(guiwb).Sheets("Inputs").Range("F9") = CDbl(txt_baiterations)
    Workbooks(guiwb).Sheets("Inputs").Range("F10") = CDbl(txt_baburniterations)
End If

Call checkBAblanks

End Sub

Private Sub UserForm_Activate()

' Define variables
Dim FilePath As Variant
Dim sourcestring As String
Dim inputvariable As String

Dim cmdLine As String
Dim rcode As Variant
Dim sevinput As Variant
Dim inputfile As String
Dim guiwb As String

guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

txt_bainput.Value = ""
txt_bacodefilepath.Value = Workbooks(guiwb).Sheets("Inputs").Range("F8")
    
cmd_startBA.Visible = False

If Workbooks(guiwb).Sheets("Inputs").Range("F8").Value <> "" Then
    txt_bainput = replace(Workbooks(guiwb).Sheets("Inputs").Range("F8").Value, "\", "/")
Else
    ' Obtain the file path from the pop-up window
    ''MsgBox "Select the BAinput .csv file", vbOKOnly, "Select BAinput file"
    ''FilePath = Application.GetOpenFilename(, , "Select Before After Input File")
    
    ' If the user doesn't select a file, then the box will be blank
    ''If FilePath = False Then
    ''    txt_bainput = ""
    ''    MsgBox "Cannot proceed without BA Input file."
    ''    Exit Sub
    ''Else
    ''    txt_bainput = FilePath
    ''    Workbooks(guiwb).Sheets("Inputs").Range("F8").Value = Replace(txt_bainput, "\", "/")
    ''End If
End If

txt_bacodefilepath = ""

' Obtain the file path from the pop-up window
''MsgBox "Select the Before After Statistical R Code", vbOKOnly, "Select Before After Statistical R Code"
''FilePath = Application.GetOpenFilename(, , "Select Before After Statistical R Code")

' If the user doesn't select a file, then the box will be blank
''If FilePath = False Then
''    txt_bainput = ""
''    MsgBox "Cannot proceed without Before After Statical R Code."
''    Exit Sub
''Else
''    txt_bacodefilepath = FilePath
''    Workbooks(guiwb).Sheets("Inputs").Range("F11").Value = Replace(txt_bacodefilepath, "\", "/")
''End If

txt_baiterations = ""
txt_baburniterations = ""
'txt_bacodefilepath = ""

End Sub

