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
    'Turn on MS scripting runtime
    Test_MS_Scripting_Runtime
    'Export VBA code to github upon opening the workbook
    ExportVisualBasicCode
End Sub
