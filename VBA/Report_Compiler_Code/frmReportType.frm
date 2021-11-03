VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportType 
   Caption         =   "Report Compiler"
   ClientHeight    =   4008
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4050
   OleObjectBlob   =   "frmReportType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReportType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()

frmReportType.Hide

End Sub

Private Sub cmdOK_Click()

If optSegmentReports Then
    frmReportType.Hide
    Call CompileAnalysisReports
ElseIf optIntReports Then
    frmReportType.Hide
    Call CompileIntAnalysisReports
ElseIf optISAM2019Reports Then
    frmReportType.Hide
    Call Compile2019ISAMAnalysisReports
ElseIf optCAMSReports Then
    frmReportType.Hide
    Call CompileAnalysisReportsCAMS
Else
    MsgBox "Please select a report type before continuing.", , "Select Report Type"
End If

End Sub


Private Sub optCAMSReports_Click()

End Sub

Private Sub UserForm_Activate()

optSegmentReports = False
optIntReports = False

End Sub


