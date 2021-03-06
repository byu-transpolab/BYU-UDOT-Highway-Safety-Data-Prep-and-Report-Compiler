VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_uicpmvariable 
   Caption         =   "Safety Statistical Analysis: UICPM Variable Selection (R GUI)"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   OleObjectBlob   =   "form_uicpmvariable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_uicpmvariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Option Explicit
'R GUI workbook created for UDOT Roadway Safety Analysis Methodology
'Comments by Sam Mineer, Brigham Young Univerisity, June 2016

Private Sub cmd_clearlist_Click()

Dim i As Integer
Dim n As Integer

Dim sourcestring As String
Dim inputvariable As String

    lst_modelvariables.Clear
    cmd_startanalysis.Visible = False
    lst_manual.Clear

    i = 1
    Do While ActiveWorkbook.Sheets("Key").Cells(i + 1, 3) <> ""
        i = i + 1
    Loop

    n = 1
    With lst_manual
        .Clear
        For n = 1 To i
        inputvariable = (Str(n) & "-" & ActiveWorkbook.Sheets("Key").Cells(n, 3))
        .AddItem inputvariable
        Next n
    End With
    
    Call clearnoxs
    
    Call loadmainxs

End Sub

Private Sub cmd_clearmanual_Click()
    lst_manual.MultiSelect = fmMultiSelectSingle
    lst_manual.Value = ""
    lst_manual.MultiSelect = fmMultiSelectMulti
End Sub

Private Sub cmd_horseshoe_Click()
    
    Call startanalysis(True)

End Sub

Private Sub cmd_movebackselected_Click()

    With lst_modelvariables
        Dim itemIndex As Integer
        For itemIndex = .ListCount - 1 To 0 Step -1
            If .Selected(itemIndex) Then
                lst_manual.AddItem .List(itemIndex)
                .RemoveItem itemIndex
            End If
        Next itemIndex
        .MultiSelect = fmMultiSelectSingle
        .Value = ""
        .MultiSelect = fmMultiSelectMulti
    End With
    
    Call SortListBox(lst_modelvariables, 0, 1, 1)
    Call SortListBox(lst_manual, 0, 1, 1)

End Sub

Private Sub cmd_moveselected_Click()
    With lst_manual
        Dim itemIndex As Integer
        For itemIndex = .ListCount - 1 To 0 Step -1
            If .Selected(itemIndex) Then
                lst_modelvariables.AddItem .List(itemIndex)
                .RemoveItem itemIndex
            End If
        Next itemIndex
        .MultiSelect = fmMultiSelectSingle
        .Value = ""
        .MultiSelect = fmMultiSelectMulti
    End With
        
    Call SortListBox(lst_modelvariables, 0, 1, 1)
    Call SortListBox(lst_manual, 0, 1, 1)

End Sub

Private Sub startanalysis(horseshoe As Boolean)

Dim usercheck As Variant
Dim rscript As String
Dim rcode As String
Dim modelwd As String
Dim niter As Long
Dim nburn As Long
Dim datalocation As String
Dim xs As String

Dim itemIndex As Integer
Dim xsnum As String

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

usercheck = MsgBox("Are you sure you are ready to begin the Statistical Analysis?", vbYesNo, "Ready?")

If usercheck = vbNo Then
    MsgBox "Analysis Aborted", vbOKOnly, "Try Again"
    Exit Sub
Else
End If

' Translate the list to a string of numbers and commas
If horseshoe Then
    xs = "()"
Else
    xs = "("
    With lst_modelvariables
        For itemIndex = .ListCount - 1 To 0 Step -1
            xsnum = Left(.List(itemIndex), InStr(1, .List(itemIndex), "-") - 1)
            xs = xs & "," & xsnum
            xs = Replace(xs, " ", "")
            xs = Replace(xs, "(,", "(")
        Next itemIndex
    End With
    xs = xs & ")"
End If
' paste xs to input sheet
Workbooks(guiwb).Sheets("Inputs").Range("I11") = xs

rscript = Workbooks(guiwb).Sheets("Inputs").Range("I3")
modelwd = Workbooks(guiwb).Sheets("Inputs").Range("I2")
rcode = Workbooks(guiwb).Sheets("Inputs").Range("I9")
modelwd = modelwd & "/" & "CrashAnalysis_" & Replace(Date, "/", "-") & "_" & Replace(Replace(Time, ":", "-"), " ", "_")
MkDir modelwd
niter = Workbooks(guiwb).Sheets("Inputs").Range("I7")
nburn = Workbooks(guiwb).Sheets("Inputs").Range("I8")
datalocation = Workbooks(guiwb).Sheets("Inputs").Range("I10")

xs = Workbooks(guiwb).Sheets("Inputs").Range("I11")

' Clear and close form
txt_inputfilepath = ""
txt_uicpmrscript = ""
txt_iterations = ""
txt_burniterations = ""
rad_horseshoe.Value = False
rad_manual.Value = False

Workbooks(guiwb).Sheets("Inputs").Range("I10") = datalocation
Workbooks(guiwb).Sheets("Inputs").Range("I9") = rcode
form_uicpmvariable.Hide

Call executeUICPM(rscript, rcode, modelwd, niter, nburn, datalocation, xs)

End Sub

Private Sub cmd_startanalysis_Click()

Call startanalysis(False)

End Sub

Private Sub cmd_input_Click()

' Define variables
Dim FilePath As Variant
Dim inputfile As String
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")
Dim i As Integer

Dim wdFP As String
Dim row1, col1 As Integer

'Find working directory location
row1 = 1
col1 = 1
Do Until Sheets("Inputs").Cells(row1, col1) = "UICPM"
    col1 = col1 + 1
Loop
Do Until Sheets("Inputs").Cells(row1, col1) = "Working Directory"
    row1 = row1 + 1
Loop
col1 = col1 + 1

'Assign working directory file path
wdFP = Replace(Sheets("Inputs").Cells(row1, col1), "/", "\")

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = wdFP
    .AllowMultiSelect = False
    .Title = "Select the UICPM Model Input File"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_inputfilepath = ""
    
Else
    txt_inputfilepath = Replace(FilePath, "\", "/")
    
    Workbooks(guiwb).Sheets("Inputs").Range("I10") = CStr(txt_inputfilepath)

    inputfile = Replace(txt_inputfilepath, "\", "/")
    inputfile = Mid(inputfile, InStr(1, inputfile, "/UICPM"))
    inputfile = Replace(inputfile, "/", "")
    
    ' Checks if file is open
    If AlreadyOpen(inputfile) Then
        'The file is already open.
    Else
        Workbooks.Open txt_inputfilepath 'Replace(txt_inputfilepath, "\\", "\") 'sPath & sFilename
    End If
    
    inputfile = Replace(inputfile, ".csv", "")
    inputfile = Replace(inputfile, ".xls", "")
    inputfile = Replace(inputfile, ".xlsx", "")
    
    ' Copy the headings from the UCPM Input file to the Key sheet
    Workbooks(inputfile).Activate
    i = 1
    Do While ActiveSheet.Cells(1, i + 1) <> ""
        i = i + 1
    Loop
    Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(1, i)).Copy
    Workbooks(guiwb).Sheets("Key").Activate
    Workbooks(guiwb).Sheets("Key").Range("C1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=True
    Application.CutCopyMode = False
    Workbooks(guiwb).Sheets("Home").Activate
    
    Workbooks(inputfile).Close False
    
End If

End Sub


Sub loadmanuallist()

' Define variables
Dim i As Integer
Dim n As Integer
Dim guiwb As String

guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

lst_modelvariables.Clear
cmd_startanalysis.Visible = False
lst_manual.Clear

    i = 1
    Do Until Workbooks(guiwb).Sheets("Key").Cells(i + 1, 3) = ""
        i = i + 1
    Loop
    
    n = 1
    With lst_manual
        For n = 1 To i
        inputvariable = (Str(n) & "-" & ActiveWorkbook.Sheets("Key").Cells(n, 3))
        .AddItem inputvariable
        Next n
    End With
    
    Call clearnoxs
    
    Call loadmainxs    'Commented out so that defaults are not shown. Text displayed above showing recommended variables.

End Sub
Sub clearnoxs()

' Define variables
Dim itemIndex As Integer
Dim noxs As Object
    Set noxs = CreateObject("Scripting.Dictionary")
    
    ' Remove unlikely variables for the statistical analysis
    With noxs
        .Add "INT_ID", 1
        .Add "LATITUDE", 2
        .Add "LONGITUDE", 3
        .Add "YEAR", 4
        .Add "SR_SR_INT", 5
        .Add "ROUTE_1", 6
        .Add "ROUTE_2", 7
        .Add "ROUTE_3", 8
        .Add "INT_MP_1", 9
        .Add "INT_MP_2", 10
        .Add "INT_MP_3", 11
        .Add "REGION", 12
        .Add "COUNTY", 13
        .Add "CITY", 14
        .Add "Total_Crashes", 15
        .Add "Severe_Crashes", 16
        .Add "Sum_Total_Crashes", 17
        .Add "Sev_5_Crashes", 18
        .Add "Sev_4_Crashes", 19
        .Add "Sev_3_Crashes", 20
        .Add "Sev_2_Crashes", 21
        .Add "Sev_1_Crashes", 22
    End With
    
    With lst_manual
        For itemIndex = .ListCount - 1 To 0 Step -1
            possxs = .List(itemIndex)
            possxs = Mid(possxs, InStr(1, possxs, "-") + 1)
            If noxs.Exists(possxs) Then
                'lst_modelvariables.AddItem .List(itemIndex)
                .RemoveItem itemIndex
            End If
        Next itemIndex
        .MultiSelect = fmMultiSelectSingle
        .Value = ""
        .MultiSelect = fmMultiSelectMulti
    End With

End Sub


Sub loadmainxs()
'This is where default headers will be specified. These will always be loaded, no matter what the user selects on the GUI. This allows for the statistical model to run.

Dim i As Integer
Dim n As Integer
Dim possxs As String
Dim itemIndex As Integer
Dim mainxs As Object
    Set mainxs = CreateObject("Scripting.Dictionary")

With mainxs
    .Add "ENT_VEH", 5
    .Add "MIN_NUM_LANES", 4
    .Add "MAX_SPEED_LIMIT", 3
    .Add "MIN_SPEED_LIMIT", 2
    .Add "PERCENT_TRUCKS", 1
End With

    With lst_manual
        For itemIndex = .ListCount - 1 To 0 Step -1
            possxs = .List(itemIndex)
            possxs = Mid(possxs, InStr(1, possxs, "-") + 1)
            If mainxs.Exists(possxs) Then
                lst_modelvariables.AddItem .List(itemIndex)
                .RemoveItem itemIndex
            End If
        Next itemIndex
        .MultiSelect = fmMultiSelectSingle
        .Value = ""
        .MultiSelect = fmMultiSelectMulti
    End With

    cmd_startanalysis.Visible = True

End Sub

Private Sub cmd_ucpsmrscript_Click()

' Define variables
Dim FilePath As Variant
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Dim wdFP As String
Dim row1, col1 As Integer

'Find working directory location
row1 = 1
col1 = 1
Do Until Sheets("Inputs").Cells(row1, col1) = "UICPM"
    col1 = col1 + 1
Loop
Do Until Sheets("Inputs").Cells(row1, col1) = "Working Directory"
    row1 = row1 + 1
Loop
col1 = col1 + 1

'Assign working directory file path
wdFP = Replace(Sheets("Inputs").Cells(row1, col1), "/", "\")

' Obtain the file path from the pop-up window
With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = wdFP
    .AllowMultiSelect = False
    .Title = "Select UICPM R Model"
    If .Show <> -1 Then MsgBox "No folder selected.": Exit Sub
        FilePath = .SelectedItems(1)
End With

' If the user doesn't select a file, then the box will be blank
If FilePath = False Then
    txt_uicpmrscript = ""
    
Else
    txt_uicpmrscript = Replace(FilePath, "\", "/")
    
End If

Workbooks(guiwb).Sheets("Inputs").Range("I9") = CStr(txt_uicpmrscript)
'MsgBox "Update the UCPM input GUI. Load manual data."

End Sub

Private Sub rad_horseshoe_Change()

If rad_horseshoe.Value = True Then
    lst_modelvariables.Visible = False
    lst_modelvariables.Clear
    cmd_startanalysis.Visible = False
    cmd_horseshoe.Visible = True
    lst_manual.Visible = False
    cmd_clearmanual.Visible = False
    cmd_moveselected.Visible = False
    cmd_movebackselected.Visible = False
    lbl_BasicVariable.Visible = False
Else
    lst_modelvariables.Visible = True
    cmd_horseshoe.Visible = False
    lst_manual.Visible = True
    cmd_clearmanual.Visible = True
    cmd_moveselected.Visible = True
    cmd_movebackselected.Visible = True
    lbl_BasicVariable.Visible = True
    Call loadmanuallist
End If

End Sub


Private Sub rad_manual_Change()
If rad_manual.Value = True Then
    lst_modelvariables.Visible = True
    cmd_horseshoe.Visible = False
    lst_manual.Visible = True
    cmd_clearmanual.Visible = True
    cmd_moveselected.Visible = True
    cmd_movebackselected.Visible = True
    lbl_BasicVariable.Visible = True
    Call loadmanuallist
Else
    lst_modelvariables.Visible = False
    lst_modelvariables.Clear
    cmd_startanalysis.Visible = False
    cmd_horseshoe.Visible = True
    lst_manual.Visible = False
    cmd_clearmanual.Visible = False
    cmd_moveselected.Visible = False
    cmd_movebackselected.Visible = False
    lbl_BasicVariable.Visible = False
End If


End Sub

Private Sub txt_burniterations_Change()

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

If txt_burniterations.Value = "" Then
Else
    If CDbl(txt_burniterations.Value) > (0.101 * CDbl(txt_iterations.Value)) Then
        MsgBox "Do not set Burn-in Iterations greater than 10% of number of iterations.", vbOKOnly, "Warning"
        txt_burniterations.Value = txt_iterations.Value * 0.1
    End If
    txt_burniterations.Value = txt_iterations.Value * 0.1
    Workbooks(guiwb).Sheets("Inputs").Range("I7") = CDbl(txt_iterations)
    Workbooks(guiwb).Sheets("Inputs").Range("I8") = CDbl(txt_burniterations)
End If

Call checkblank

End Sub

Private Sub txt_iterations_Change()

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

If txt_iterations.Value = "" Then
Else
    txt_burniterations.Value = txt_iterations.Value * 0.1
    Workbooks(guiwb).Sheets("Inputs").Range("I7") = CDbl(txt_iterations)
    Workbooks(guiwb).Sheets("Inputs").Range("I8") = CDbl(txt_burniterations)
End If

Call checkblank

End Sub

Private Sub txt_inputfilepath_Change()

'Dim i As Integer
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Call checkblank

Workbooks(guiwb).Sheets("Inputs").Range("I10") = CStr(txt_inputfilepath)

End Sub

Private Sub txt_uicpmrscript_Change()

Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

Call checkblank

Workbooks(guiwb).Sheets("Inputs").Range("I9") = CStr(txt_uicpmrscript)

End Sub


Private Sub checkblank()

If Len(CStr(txt_burniterations)) <> 0 And Len(CStr(txt_iterations)) <> 0 And FileExists(txt_uicpmrscript) And txt_uicpmrscript <> "" And FileExists(txt_inputfilepath) And txt_inputfilepath <> "" Then ' Then

    lbl_heading3ucpm.Visible = True
    rad_horseshoe.Visible = True
    rad_horseshoe.Value = False
    rad_manual.Visible = True
    rad_manual.Value = False
    lst_modelvariables.Visible = False
    cmd_startanalysis.Visible = False
    cmd_horseshoe.Visible = False
    lbl_BasicVariable.Visible = False
    
Else
    
    lbl_heading3ucpm.Visible = False
    rad_horseshoe.Visible = False
    rad_horseshoe.Value = False
    rad_manual.Visible = False
    rad_manual.Value = False
    lst_modelvariables.Visible = False
    cmd_startanalysis.Visible = False
    cmd_horseshoe.Visible = False
    lbl_BasicVariable.Visible = False
    
End If

End Sub

Private Sub UserForm_Activate()

Dim i As Integer
Dim n As Integer

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
    guiwb = Replace(guiwb, ".xlsm", "")
    
    txt_inputfilepath = ""
    txt_uicpmrscript = ""
    txt_iterations = ""
    txt_burniterations = ""
    rad_horseshoe.Value = False
    rad_manual.Value = False
    lst_modelvariables.Clear
    'cmd_startanalysis.Visible = False
        
    'Checks if file has been created
    If Workbooks(guiwb).Sheets("Inputs").Range("I10") <> "" Then
        txt_inputfilepath = Workbooks(guiwb).Sheets("Inputs").Range("I10")
        txt_inputfilepath = Replace(txt_inputfilepath.Value, "\", "/")
       
        inputfile = Replace(txt_inputfilepath, "\", "/")
        inputfile = Mid(inputfile, InStr(1, inputfile, "/UICPM"))
        inputfile = Replace(inputfile, "/", "")
 
        ' Checks if file is open
        If AlreadyOpen(inputfile) Then
            'The file is already open.
        Else
            Workbooks.Open txt_inputfilepath 'Replace(txt_inputfilepath, "\\", "\") 'sPath & sFilename
        End If
    
        inputfile = Replace(inputfile, ".csv", "")
        inputfile = Replace(inputfile, ".xls", "")
        inputfile = Replace(inputfile, ".xlsx", "")
    
        ' Copy the headings from the UCPM Input file to the Key sheet
        Workbooks(inputfile).Activate
        i = 1
        Do While ActiveSheet.Cells(1, i + 1) <> ""
            i = i + 1
       Loop
       Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(1, i)).Copy
       Workbooks(guiwb).Sheets("Key").Activate
       Workbooks(guiwb).Sheets("Key").Range("C1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
               :=False, Transpose:=True
       Application.CutCopyMode = False
       Workbooks(guiwb).Sheets("Home").Activate
    
       If Workbooks(guiwb).Sheets("Inputs").Range("I10") = "" Then
           Workbooks(guiwb).Sheets("Inputs").Range("I10") = Replace(txt_inputfilepath, "\", "/")
       End If
 
    End If
  
    lbl_heading3ucpm.Visible = False
    rad_horseshoe.Value = False
    rad_manual.Value = False
    lst_manual.Visible = False
    cmd_clearmanual.Visible = False
    cmd_moveselected.Visible = False
    cmd_movebackselected.Visible = False
    cmd_startanalysis.Visible = False
    lbl_BasicVariable.Visible = False
    
    End Sub
    
Sub SortListBox(oLb As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
Dim vaItems As Variant
Dim i As Long, j As Long
Dim c As Integer
Dim vTemp As Variant

'Obtained online at http://www.ozgrid.com/forum/showthread.php?t=71509 to sort listboxes with changes in list items.
 
 'Put the items in a variant array
vaItems = oLb.List
 
 'Sort the Array Alphabetically(1)
If sType = 1 Then
    For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
        For j = i + 1 To UBound(vaItems, 1)
             'Sort Ascending (1)
            If sDir = 1 Then
                If vaItems(i, sCol) > vaItems(j, sCol) Then
                    For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                        vTemp = vaItems(i, c)
                        vaItems(i, c) = vaItems(j, c)
                        vaItems(j, c) = vTemp
                    Next c
                End If
                 
                 'Sort Descending (2)
            ElseIf sDir = 2 Then
                If vaItems(i, sCol) < vaItems(j, sCol) Then
                    For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                        vTemp = vaItems(i, c)
                        vaItems(i, c) = vaItems(j, c)
                        vaItems(j, c) = vTemp
                    Next c
                End If
            End If
             
        Next j
    Next i
     'Sort the Array Numerically(2)
     '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
ElseIf sType = 2 Then
    For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
        For j = i + 1 To UBound(vaItems, 1)
             'Sort Ascending (1)
            If sDir = 1 Then
                If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
                    For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                        vTemp = vaItems(i, c)
                        vaItems(i, c) = vaItems(j, c)
                        vaItems(j, c) = vTemp
                    Next c
                End If
                 
                 'Sort Descending (2)
            ElseIf sDir = 2 Then
                If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                    For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                        vTemp = vaItems(i, c)
                        vaItems(i, c) = vaItems(j, c)
                        vaItems(j, c) = vTemp
                    Next c
                End If
            End If
             
        Next j
    Next i
End If
 
 'Set the list to the array
oLb.List = vaItems
End Sub

