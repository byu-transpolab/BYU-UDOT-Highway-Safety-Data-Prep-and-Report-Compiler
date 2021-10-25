Attribute VB_Name = "Export_to_GitHub"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    'Set fso = CreateObject("Scripting.FileSystemObject")
    
    directory = ActiveWorkbook.path & "\Data_Prep_Code\VBA"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    'Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    'Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
    MsgBox ("Successfully exported " & CStr(count) & " VBA files to " & directory)
    Application.StatusBar = False
End Sub

Sub Test_MS_Scripting_Runtime()
    Dim Ref As Object, CheckRefEnabled%
    CheckRefEnabled = 0
    With ThisWorkbook
        For Each Ref In .VBProject.References
            If Ref.Name = "Scripting" Then
                CheckRefEnabled = 1
                Exit For
            End If
        Next Ref
        If CheckRefEnabled = 0 Then
            .VBProject.References.AddFromGUID "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        End If
    End With
End Sub
