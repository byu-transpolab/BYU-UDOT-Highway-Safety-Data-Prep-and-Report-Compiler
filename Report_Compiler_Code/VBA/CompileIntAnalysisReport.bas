Attribute VB_Name = "CompileIntAnalysisReport"
Option Explicit

Sub CompileIntAnalysisReports()

'*************************************************
'
'  The purpose of this code is to create Analyis
'    Reports for the "Top 20" list produced from
'    the Statistical Model.
'  These codes should be ran after (1) the crash
'    database has been created and (2) after the
'    "Combo" and "CrashFactors" datasets have
'    been created.
'  General Steps of this code:
'     (1) Save "Report Compiler" as a new file in
'             a new directory, for this specific
'             analysis Group
'     (2) Copy over the "Combo" and "CrashFactors"
'             sheets from the relative "Combo" file
'     (3) For each segment in "Combo" (the "Top 20"
'             list), create a new report sheet and
'             fill out report
'
'   A copy of the CrashFactors and Vehicle data is
'     copied as a reference for error checking.
'   Once report is complete, the report and
'     corresponding copy of the CrashFactors and
'     Vehicle Crash data are moved to a new workbook.
'   This process is repeated for each line on the
'     "Combo" sheet (model built for 1 or 200+
'     segments, as data is avilable)
'
'  Comments by Sam Mineer, Brigham Young University,
'  Updated 11/2015  05/2016  03/2017
'
'  Edited by Josh Gibbons, BYU, for Intersection Model
'
'*************************************************

'-------------------------------------------------
'
'  Define variables (descriptions given after words)
'
'-------------------------------------------------
Dim n As Integer            'counting variables, used in various manners
Dim i As Integer            'counting variables, used in various manners
Dim j As Integer            'counting variables, used in various manners
Dim k As Integer            'counting variables, used in various manners
Dim combofp As String       'filepath to Combo and CrashFactors data
Dim combofn As String       'filename of Combo Workbook, with Combo and CrashFactors datasets
Dim combosn As String       'sheet name for "Combo" sheet
Dim crashfp As String       'filepath to crash data
Dim crashfn As String       'filename of Crash Workbook, with vehicle
Dim crashsn As String       'sheet name for "crash/parameter" sheet
Dim reportfp As String      'filepath of where to save Analysis Reports
Dim reportfn As String      'filename of Report Compiler
Dim comborow As Integer     'row counter for Combo sheet
Dim reportrow As Integer    'row counter for reportrow (individual ones)
Dim reportrow2 As Integer   'second row counter for report
Dim crashrow As Long        'row counter for crash data
Dim endcrashrow As Long     'row counter for end of int id rows
Dim modeltype As String     'name of model used, entered by user
Dim dataranges As String    'range of years for data, entered by user
Dim nrow As Integer
Dim ncol As Integer
Dim cmdline As String
Dim pcrashes As Double

'Results sheet column number variables
Dim StateRankCol As Integer
Dim RegionRankCol As Integer
Dim CountyRankCol As Integer
Dim RegionCol As Integer
Dim CountyCol As Integer
Dim CityCol As Integer
Dim IntIDCol As Integer
Dim Route1Col As Integer
Dim Route2Col As Integer
Dim Route3Col As Integer
Dim R1MPCol As Integer
Dim R2MPCol As Integer
Dim R3MPCol As Integer
Dim LatCol As Integer
Dim LongCol As Integer
Dim IntContCol As Integer
Dim EntVehCol As Integer
Dim MaxFCCol As Integer
Dim MinFCCol As Integer
Dim MaxLanesCol As Integer
Dim MinLanesCol As Integer
Dim MaxSLCol As Integer
Dim MinSLCol As Integer
Dim TCrashCol As Integer
Dim PCrashCol As Integer
Dim ACrashCol As Integer
Dim Sev5CrashCol As Integer
Dim Sev4CrashCol As Integer
Dim Sev3CrashCol As Integer
Dim Sev2CrashCol As Integer
Dim Sev1CrashCol As Integer
Dim NumYearsCol As Integer

Dim SRank As Integer
Dim RRank As Integer
Dim CRank As Integer
Dim Region As Integer
Dim County As String
Dim City As String
Dim IntID As Integer
Dim Route1 As Integer
Dim Route2 As Integer
Dim Route3 As Integer
Dim R1MP As Double
Dim R2MP As Double
Dim R3MP As Double
Dim IntLat As Double
Dim IntLong As Double
Dim IntCont As String
Dim EntVeh As Double
Dim MaxFC As String
Dim MinFC As String
Dim MaxLanes As Integer
Dim MinLanes As Integer
Dim MaxSL As Integer
Dim MinSL As Integer
Dim TCrash As Integer
Dim PCrash As Integer
Dim ACrash As Integer
Dim Sev5Crash As Integer
Dim Sev4Crash As Integer
Dim Sev3Crash As Integer
Dim Sev2Crash As Integer
Dim Sev1Crash As Integer
Dim NumYears As Integer

Dim HeadOnCol As Integer
Dim WorkZoneCol2 As Integer
Dim PedCol2 As Integer
Dim BikeCol2 As Integer
Dim MotoCol2 As Integer
Dim ImpRestCol2 As Integer
Dim UnrestCol2 As Integer
Dim DUICol2 As Integer
Dim AggCol2 As Integer
Dim DistrCol2 As Integer
Dim DrowsyCol2 As Integer
Dim SpeedCol2 As Integer
Dim IntCol2 As Integer
Dim AdWeathCol2 As Integer
Dim AdRoadSurfCol2 As Integer
Dim RoadGeomCol2 As Integer
Dim WildAnCol2 As Integer
Dim DomAnCol2 As Integer
Dim RoadDepCol2 As Integer
Dim OverturnCol2 As Integer
Dim CommMotVehCol2 As Integer
Dim InterstateCol2 As Integer
Dim TeenDrivCol2 As Integer
Dim OldDrivCol2 As Integer
Dim UrbanCol2 As Integer
Dim NightDarkCol2 As Integer
Dim SingleVehCol2 As Integer
Dim TrainInvCol2 As Integer
Dim RailCrossCol2 As Integer
Dim TransVehCol2 As Integer
Dim FixedObjCol2 As Integer

Dim WorkZone As Integer
Dim Ped As Integer
Dim Bike As Integer
Dim Moto As Integer
Dim ImpRest As Integer
Dim Unrest As Integer
Dim DUI As Integer
Dim Agg As Integer
Dim Distract As Integer
Dim Drowsy As Integer
Dim SpeedRel As Integer
Dim IntRel As Integer
Dim AdWeath As Integer
Dim AdRoadSurf As Integer
Dim RoadGeom As Integer
Dim WildAnimal As Integer
Dim DomAnimal As Integer
Dim RoadDep As Integer
Dim Overturn As Integer
Dim CommMotVeh As Integer
Dim Interstate As Integer
Dim TeenDriv As Integer
Dim OldDriv As Integer
Dim Urban As Integer
Dim NightDark As Integer
Dim SingleVeh As Integer
Dim TrainInv As Integer
Dim RailCross As Integer
Dim TransVeh As Integer
Dim FixedObj As Integer

'Parameters sheet column number variables
Dim CrashRow1 As Integer
Dim CSevRow As Integer
Dim FARow As Integer
Dim DataYearsRow As Integer
Dim CrashIntIDCol As Integer
Dim CrashIDCol As Integer
Dim NumVehCol As Integer
Dim FirstHarmCol As Integer
Dim ManColCol As Integer
Dim Event1Col As Integer
Dim Event2Col As Integer
Dim Event3Col As Integer
Dim Event4Col As Integer
Dim MostHarmCol As Integer
Dim VehManCol As Integer
Dim CLatCol As Integer
Dim CLongCol As Integer
Dim WorkZoneCol As Integer
Dim PedCol As Integer
Dim BikeCol As Integer
Dim MotoCol As Integer
Dim ImpRestCol As Integer
Dim UnrestCol As Integer
Dim DUICol As Integer
Dim AggCol As Integer
Dim DistrCol As Integer
Dim DrowsyCol As Integer
Dim SpeedCol As Integer
Dim IntCol As Integer
Dim AdWeathCol As Integer
Dim AdRoadSurfCol As Integer
Dim RoadGeomCol As Integer
Dim WildAnCol As Integer
Dim DomAnCol As Integer
Dim RoadDepCol As Integer
Dim OverturnCol As Integer
Dim CommMotVehCol As Integer
Dim InterstateCol As Integer
Dim TeenDrivCol As Integer
Dim OldDrivCol As Integer
Dim UrbanCol As Integer
Dim NightDarkCol As Integer
Dim SingleVehCol As Integer
Dim TrainInvCol As Integer
Dim RailCrossCol As Integer
Dim TransVehCol As Integer
Dim FixedObjCol As Integer

Dim CrashSevs As String
Dim FAMethod As String
Dim DataYears As String
Dim CrashIntID As Integer
Dim CrashID As Integer
Dim NumVeh As Integer
Dim FirstHarm As Integer
Dim ManColl As Integer
Dim Event1 As Integer
Dim Event2 As Integer
Dim Event3 As Integer
Dim Event4 As Integer
Dim MostHarm As Integer
Dim VehMan As Integer
Dim CLat As Integer
Dim CLong As Integer

Dim CountyName As String        'County Name
Dim SRoute1 As String
Dim SRoute2 As String
Dim IntString As String
ReDim RegCount(1 To 4) As Integer
ReDim CountyCount(1 To 29) As Integer
Dim NumInts As Integer

Dim RegionRow As Integer
Dim CountyRow As Integer
Dim IntRow As Integer

Dim fldr As FileDialog
Dim sItem As String
Dim strPath As String
Dim msgboxstatus As Variant
Dim wksht As Worksheet
Dim CurrentReport As String

'-------------------------------------------------
'
'  Start of code
'
'  Sets ups the document before copying data over
'
'-------------------------------------------------

Dim guiwb As String

guiwb = ActiveWorkbook.Name
guiwb = Replace(guiwb, ".xlsm", "")

'   Assign model type
modeltype = "UICPM"

'   Windows("Report Compiler").Activate
Sheets("Main").Activate

'Check if previous sheets are in workbook
reportfn = ActiveWorkbook.Name
For Each wksht In Worksheets
    If wksht.Name = "Results" Or wksht.Name = "Parameters" Then
        Application.DisplayAlerts = False
        wksht.Delete
        Application.DisplayAlerts = True
    End If
Next

'   Identify filename of UICPM Results Workbook, as given by the user
msgboxstatus = MsgBox("Open the Workbook with the UICPM Model Results", vbOKCancel, "Open the Model Results")

'   If the user clicks 'Cancel', the macro is aborted
If msgboxstatus = vbOK Then
    'Continue
ElseIf msgboxstatus = vbCancel Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    End
End If

'User opens the model results
Application.Dialogs(xlDialogOpen).Show

'Get workbook and sheet names
combofn = ActiveWorkbook.Name
combosn = ActiveSheet.Name

'Identify filename of Crash Workbook, with Crash Data, as given by the user
msgboxstatus = MsgBox("Open the Workbook with the model parameters and select crash data.", vbOKCancel, "Open Parameters and Crash Workbook")

'  If the user clicks 'Cancel', the macro is aborted
If msgboxstatus = vbOK Then
ElseIf msgboxstatus = vbCancel Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    Exit Sub
End If

'  User selects the Vehicle Crash Workbook
Application.Dialogs(xlDialogOpen).Show

'  Waits for 1 second to give time for the workbook to open successfully
Application.Wait (Now + TimeValue("00:00:01"))

'Get workbook name and set sheet name
crashfn = ActiveWorkbook.Name
crashsn = "Parameters"
ActiveSheet.Name = crashsn

'Activate report compiler workbook
Windows(reportfn).Activate
Application.ScreenUpdating = False          'Stop screen updating to prevent strobe effect

'----------------------------------------------------------------------------------
'   Load model results sheet into workbook. Find select columns.
'----------------------------------------------------------------------------------

'  Moves a copy of the "Results" sheet to the "Report Compiler" file, to use for future reference
Windows(combofn).Activate           'Results
Sheets(combosn).Activate
Sheets(combosn).Copy After:=Workbooks(reportfn).Sheets("Main")
Workbooks(reportfn).Sheets(combosn).Name = "Results"
combosn = "Results"

'Wait one second to make sure all code runs to this point.
Application.Wait (Now + TimeValue("00:00:01"))

'Activate results sheet in Report Compiler workbook
Sheets(combosn).Activate

'Find column numbers in results data
IntRow = 2
StateRankCol = FindColumn("State_Rank")
RegionRankCol = FindColumn("Region_Rank")
CountyRankCol = FindColumn("County_Rank")
RegionCol = FindColumn("REGION")
CountyCol = FindColumn("COUNTY")
CityCol = FindColumn("CITY")
IntIDCol = FindColumn("INT_ID")
Route1Col = FindColumn("ROUTE_1")
Route2Col = FindColumn("ROUTE_2")
'Route3Col = FindColumn("ROUTE_3")                  'Takes into account if route 3 doesn't exist
Route3Col = 1
Do Until Cells(1, Route3Col) = "ROUTE_3" Or Cells(1, Route3Col) = ""
    Route3Col = Route3Col + 1
Loop
If Cells(1, Route3Col) = "" Then
    Route3Col = 0
End If

R1MPCol = FindColumn("INT_MP_1")
R2MPCol = FindColumn("INT_MP_2")
'R3MPCol = FindColumn("INT_MP_3")                   'Takes into account if route 3 doesn't exist
R3MPCol = 1
Do Until Cells(1, R3MPCol) = "INT_MP_3" Or Cells(1, R3MPCol) = ""
    R3MPCol = R3MPCol + 1
Loop
If Cells(1, R3MPCol) = "" Then
    R3MPCol = 0
End If

LatCol = FindColumn("LATITUDE")
LongCol = FindColumn("LONGITUDE")
IntContCol = FindColumn("INT_CONTROL")
EntVehCol = FindColumn("ENT_VEH")
MaxFCCol = FindColumn("MAX.FC_TYPE")
MinFCCol = FindColumn("MIN.FC_TYPE")
MaxLanesCol = FindColumn("MAX_NUM_LANES")
MinLanesCol = FindColumn("MIN_NUM_LANES")
MaxSLCol = FindColumn("MAX_SPEED_LIMIT")
MinSLCol = FindColumn("MIN_SPEED_LIMIT")
TCrashCol = FindColumn("Sum_Total_Crashes_TotalCount")
PCrashCol = FindColumn("Expected_Total")
ACrashCol = FindColumn("Total_Crashes_TotalCount")
Sev5CrashCol = FindColumn("Sev_5_Crashes_TotalCount")
Sev4CrashCol = FindColumn("Sev_4_Crashes_TotalCount")
Sev3CrashCol = FindColumn("Sev_3_Crashes_TotalCount")
Sev2CrashCol = FindColumn("Sev_2_Crashes_TotalCount")
Sev1CrashCol = FindColumn("Sev_1_Crashes_TotalCount")
NumYearsCol = FindColumn("TotalYears")

HeadOnCol = FindColumn("HEADON_COLLISION_TotalCount")
WorkZoneCol2 = FindColumn("WORKZONE_RELATED_TotalCount")
PedCol2 = FindColumn("PEDESTRIAN_INVOLVED_TotalCount")
BikeCol2 = FindColumn("BICYCLIST_INVOLVED_TotalCount")
MotoCol2 = FindColumn("MOTORCYCLE_INVOLVED_TotalCount")
ImpRestCol2 = FindColumn("IMPROPER_RESTRAINT_TotalCount")
UnrestCol2 = FindColumn("UNRESTRAINED_TotalCount")
DUICol2 = FindColumn("DUI_TotalCount")
AggCol2 = FindColumn("AGGRESSIVE_DRIVING_TotalCount")
DistrCol2 = FindColumn("DISTRACTED_DRIVING_TotalCount")
DrowsyCol2 = FindColumn("DROWSY_DRIVING_TotalCount")
SpeedCol2 = FindColumn("SPEED_RELATED_TotalCount")
'IntCol2 = FindColumn("INTERSECTION_RELATED_TotalCount")
AdWeathCol2 = FindColumn("ADVERSE_WEATHER_TotalCount")
AdRoadSurfCol2 = FindColumn("ADVERSE_ROADWAY_SURF_CONDITION_TotalCount")
RoadGeomCol2 = FindColumn("ROADWAY_GEOMETRY_RELATED_TotalCount")
WildAnCol2 = FindColumn("WILD_ANIMAL_RELATED_TotalCount")
DomAnCol2 = FindColumn("DOMESTIC_ANIMAL_RELATED_TotalCount")
RoadDepCol2 = FindColumn("ROADWAY_DEPARTURE_TotalCount")
OverturnCol2 = FindColumn("OVERTURN_ROLLOVER_TotalCount")
CommMotVehCol2 = FindColumn("COMMERCIAL_MOTOR_VEH_INVOLVED_TotalCount")
InterstateCol2 = FindColumn("INTERSTATE_HIGHWAY_TotalCount")
TeenDrivCol2 = FindColumn("TEENAGE_DRIVER_INVOLVED_TotalCount")
OldDrivCol2 = FindColumn("OLDER_DRIVER_INVOLVED_TotalCount")
'UrbanCol2 = FindColumn("URBAN_COUNTY_TotalCount")
NightDarkCol2 = FindColumn("NIGHT_DARK_CONDITION_TotalCount")
SingleVehCol2 = FindColumn("SINGLE_VEHICLE_TotalCount")
TrainInvCol2 = FindColumn("TRAIN_INVOLVED_TotalCount")
RailCrossCol2 = FindColumn("RAILROAD_CROSSING_TotalCount")
TransVehCol2 = FindColumn("TRANSIT_VEHICLE_INVOLVED_TotalCount")
FixedObjCol2 = FindColumn("COLLISION_WITH_FIXED_OBJECT_TotalCount")


'----------------------------------------------------------------------------------
'  Cycle through results data to index the results by region and county. Then have
'  the user select the desired intersections before proceeding.
'----------------------------------------------------------------------------------

'Reset region count
For i = 1 To 4
    RegCount(i) = 0
Next i

'Reset county count
For i = 1 To 29
    CountyCount(i) = 0
Next i


'Clear previous data
Sheets("SegKey").Range("A2:A10").ClearContents
Sheets("SegKey").Range("C2:F1000").ClearContents

'   Determine which segments to include
'Summarize intersection representation on SegKey sheet
comborow = 2
Do While Sheets(combosn).Cells(comborow, StateRankCol) <> ""
    ' Region
        If Sheets(combosn).Cells(comborow, RegionCol) = 1 Then
            RegCount(1) = RegCount(1) + 1
        ElseIf Sheets(combosn).Cells(comborow, RegionCol) = 2 Then
            RegCount(2) = RegCount(2) + 1
        ElseIf Sheets(combosn).Cells(comborow, RegionCol) = 3 Then
            RegCount(3) = RegCount(3) + 1
        ElseIf Sheets(combosn).Cells(comborow, RegionCol) = 4 Then
            RegCount(4) = RegCount(4) + 1
        End If
        
    ' County
        CountyName = Sheets(combosn).Cells(comborow, CountyCol)
        CountyRow = 2
        Do Until Sheets("SegKey").Cells(CountyRow, 2) = CountyName
            CountyRow = CountyRow + 1
        Loop
        CountyCount(CountyRow - 1) = CountyCount(CountyRow - 1) + 1
        
    ' Individual Intersection
        IntID = Format(Sheets(combosn).Cells(comborow, IntIDCol), "000")
        SRoute1 = Sheets(combosn).Cells(comborow, Route1Col)
        SRoute2 = Sheets(combosn).Cells(comborow, Route2Col)
        IntString = "Int " & IntID & ": Routes " & SRoute1 & " and " & SRoute2
        Sheets("SegKey").Cells(IntRow, 4) = IntString
        IntRow = IntRow + 1
        
    comborow = comborow + 1
Loop

For RegionRow = 2 To 5
    Sheets("SegKey").Cells(RegionRow, 1) = RegionRow - 1 & " (" & RegCount(RegionRow - 1) & ")"
Next RegionRow

For CountyRow = 2 To 30
    Sheets("SegKey").Cells(CountyRow, 3) = Sheets("SegKey").Cells(CountyRow, 2) & " (" & CountyCount(CountyRow - 1) & ")"
Next CountyRow

CountyRow = 2
Do Until Sheets("SegKey").Cells(CountyRow, 3) = ""
    If InStr(1, Sheets("SegKey").Cells(CountyRow, 3), "(0)") > 0 Then
        Sheets("SegKey").Cells(CountyRow, 3).Delete Shift:=xlUp
    Else
        CountyRow = CountyRow + 1
    End If
Loop

'   Sort individual intersections list
IntRow = IntRow - 1
Sheets("SegKey").Activate
ActiveWorkbook.Worksheets("SegKey").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("SegKey").Sort.SortFields.Add Key:=Range("D2"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("SegKey").Sort
    .SetRange Range("D2:D" & IntRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Sheets("Main").Activate

'   Show Intersection Selection user form to select intersections of interest
frmIntSelection.Show

'----------------------------------------------------------------------------------
'   Load parameters/crash sheet into workbook. Find select columns.
'----------------------------------------------------------------------------------

'Copy parameters/crash dataset
Windows(crashfn).Activate
Sheets(crashsn).Copy After:=Workbooks(reportfn).Sheets(combosn)
Application.Wait (Now + TimeValue("00:00:01"))

'Activate report compiler workbook
Windows(reportfn).Activate
Sheets(crashsn).Activate

'Find parameters and crash data rows and columns
CSevRow = 1
Do Until Sheets(crashsn).Cells(CSevRow, 1) = "Severities:"
    CSevRow = CSevRow + 1
Loop

FARow = 1
Do Until Sheets(crashsn).Cells(FARow, 1) = "Functional Area Type:"
    FARow = FARow + 1
Loop

DataYearsRow = 1
Do Until Sheets(crashsn).Cells(DataYearsRow, 1) = "Selected Years:"
    DataYearsRow = DataYearsRow + 1
Loop

CrashRow1 = 1
Do Until Sheets(crashsn).Cells(CrashRow1, 1) = "INT_ID"
    CrashRow1 = CrashRow1 + 1
Loop

CrashIntIDCol = FindPColumn("INT_ID", CrashRow1)
CrashIDCol = FindPColumn("CRASH_ID", CrashRow1)
NumVehCol = FindPColumn("VEHICLE_NUM", CrashRow1)
FirstHarmCol = FindPColumn("FIRST_HARMFUL_EVENT_ID", CrashRow1)
ManColCol = FindPColumn("MANNER_COLLISION_ID", CrashRow1)
Event1Col = FindPColumn("EVENT_SEQUENCE_1_ID", CrashRow1)
Event2Col = FindPColumn("EVENT_SEQUENCE_2_ID", CrashRow1)
Event3Col = FindPColumn("EVENT_SEQUENCE_3_ID", CrashRow1)
Event4Col = FindPColumn("EVENT_SEQUENCE_4_ID", CrashRow1)
MostHarmCol = FindPColumn("MOST_HARMFUL_EVENT_ID", CrashRow1)
VehManCol = FindPColumn("VEHICLE_MANEUVER_ID", CrashRow1)
CLatCol = FindPColumn("LATITUDE", CrashRow1)
CLongCol = FindPColumn("LONGITUDE", CrashRow1)
WorkZoneCol = FindPColumn("WORK_ZONE_RELATED_YNU", CrashRow1)
PedCol = FindPColumn("PEDESTRIAN_INVOLVED", CrashRow1)
BikeCol = FindPColumn("BICYCLIST_INVOLVED", CrashRow1)
MotoCol = FindPColumn("MOTORCYCLE_INVOLVED", CrashRow1)
ImpRestCol = FindPColumn("IMPROPER_RESTRAINT", CrashRow1)
UnrestCol = FindPColumn("UNRESTRAINED", CrashRow1)
DUICol = FindPColumn("DUI", CrashRow1)
AggCol = FindPColumn("AGGRESSIVE_DRIVING", CrashRow1)
DistrCol = FindPColumn("DISTRACTED_DRIVING", CrashRow1)
DrowsyCol = FindPColumn("DROWSY_DRIVING", CrashRow1)
SpeedCol = FindPColumn("SPEED_RELATED", CrashRow1)
IntCol = FindPColumn("INTERSECTION_RELATED", CrashRow1)
AdWeathCol = FindPColumn("ADVERSE_WEATHER", CrashRow1)
AdRoadSurfCol = FindPColumn("ADVERSE_ROADWAY_SURF_CONDITION", CrashRow1)
RoadGeomCol = FindPColumn("ROADWAY_GEOMETRY_RELATED", CrashRow1)
WildAnCol = FindPColumn("WILD_ANIMAL_RELATED", CrashRow1)
DomAnCol = FindPColumn("DOMESTIC_ANIMAL_RELATED", CrashRow1)
RoadDepCol = FindPColumn("ROADWAY_DEPARTURE", CrashRow1)
OverturnCol = FindPColumn("OVERTURN_ROLLOVER", CrashRow1)
CommMotVehCol = FindPColumn("COMMERCIAL_MOTOR_VEH_INVOLVED", CrashRow1)
InterstateCol = FindPColumn("INTERSTATE_HIGHWAY", CrashRow1)
TeenDrivCol = FindPColumn("TEENAGE_DRIVER_INVOLVED", CrashRow1)
OldDrivCol = FindPColumn("OLDER_DRIVER_INVOLVED", CrashRow1)
UrbanCol = FindPColumn("URBAN_COUNTY", CrashRow1)
NightDarkCol = FindPColumn("NIGHT_DARK_CONDITION", CrashRow1)
SingleVehCol = FindPColumn("SINGLE_VEHICLE", CrashRow1)
TrainInvCol = FindPColumn("TRAIN_INVOLVED", CrashRow1)
RailCrossCol = FindPColumn("RAILROAD_CROSSING", CrashRow1)
TransVehCol = FindPColumn("TRANSIT_VEHICLE_INVOLVED", CrashRow1)
FixedObjCol = FindPColumn("COLLISION_WITH_FIXED_OBJECT", CrashRow1)

'----------------------------------------------------------------------------------
'   Create new directory to save reports. Resave report compiler with today's date.
'----------------------------------------------------------------------------------

msgboxstatus = MsgBox("Select the Folder to Save the Report Files", vbOKCancel, "User Prompt")

'   If the user clicks 'Cancel', the macro is aborted
If msgboxstatus = vbOK Then
ElseIf msgboxstatus = vbCancel Then
    MsgBox "Macro aborted.", vbOKOnly, "Macro aborted"
    End
End If

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select Location to Save Analysis Reports (Select a Folder)"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode1
    sItem = .SelectedItems(1)
End With
NextCode1:

'   GetFolder = sItem
reportfp = sItem
Set fldr = Nothing

'   If the filepath doesn't have an ending "\",
'   then one is added
If Right(reportfp, 1) <> "\" Then
    reportfp = reportfp & "\"
End If

'  Creates new directory for saving this batch of
'    analysis reports, with unique date and time
reportfp = reportfp & "Output " & Month(Date) & "-" & Day(Date) & " " & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & "\"
MkDir reportfp

'  Resave ReportCompiler with unique date and time
ActiveWorkbook.SaveAs Filename:= _
    reportfp & "Report Compiler " & Month(Date) & "-" & Day(Date) & " " & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & ".xlsm" _
    , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
reportfn = ActiveWorkbook.Name

'   Hide unneeded segment sheet
Sheets("SegBlankReport").Visible = False
Sheets("SegBlankReport").Visible = False

'Close input files
Workbooks(combofn).Close False
Workbooks(crashfn).Close False

'----------------------------------------------------------------------------------
'   Cycle through select intersections. Load report and enter data.
'----------------------------------------------------------------------------------

'Find number of intersections
nrow = 2
Do Until Sheets("SegKey").Cells(nrow, 6) = ""
    nrow = nrow + 1
Loop
NumInts = nrow - 2

'Enable screen updating
Sheets("Main").Activate

'Show progress box
Range("H2:L3").Merge
Range("H2:L3").Font.Bold = True
Range("H2:L3").Font.Italic = True
With Range("H2:L3")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
End With

' Prints "Code Running: Please Wait"
Range("H2") = "Code Running: Please Wait."
Range("H2").Font.Size = 15
With Range("H2:L12").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 5296274    'Light Green
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Range("J7:L7").Merge
Range("J9:L9").Merge
Range("J11:L11").Merge
With Range("H4:L12")
    .HorizontalAlignment = xlLeft
End With

' Print update info
Range("H4") = "Creating Intersection Safety Analysis Reports"
Range("H5") = "Intersection"
Range("J5") = 1
Range("K5") = "' of"
Range("L5") = NumInts
Range("H7") = "Start Time:"
Range("J7") = Now()
Range("H9") = "Update Time:"
Range("H11") = "End Time:"
Application.ScreenUpdating = True
Application.ScreenUpdating = False

'Sets the comborow counter for 2, which is the top of the list. The Do Loop will step through each line until a blank entry is found
comborow = 2
Do While Sheets(combosn).Cells(comborow, StateRankCol) <> ""
    
    'Assign intersection ID and verify the intersection was selected by user
    IntID = Sheets(combosn).Cells(comborow, IntIDCol)
    nrow = 2
    Do Until Sheets("SegKey").Cells(nrow, 6) = IntID Or Sheets("SegKey").Cells(nrow, 6) = ""
        nrow = nrow + 1
    Loop
    
    If Sheets("SegKey").Cells(nrow, 6) = IntID Then               'If the intersection ID number is found in SegKey list, then create report
        
        'Assign initial values from model results file
        SRank = Sheets(combosn).Cells(comborow, StateRankCol)
        RRank = Sheets(combosn).Cells(comborow, RegionRankCol)
        CRank = Sheets(combosn).Cells(comborow, CountyRankCol)
        Region = Sheets(combosn).Cells(comborow, RegionCol)
        County = Sheets(combosn).Cells(comborow, CountyCol)
        City = Sheets(combosn).Cells(comborow, CityCol)
        Route1 = Sheets(combosn).Cells(comborow, Route1Col)
        Route2 = Sheets(combosn).Cells(comborow, Route2Col)
        If Route3Col = 0 Then
            Route3 = 0
            R3MP = 0
        Else
            Route3 = Sheets(combosn).Cells(comborow, Route3Col)
            R3MP = Sheets(combosn).Cells(comborow, R3MPCol)
        End If
        R1MP = Sheets(combosn).Cells(comborow, R1MPCol)
        R2MP = Sheets(combosn).Cells(comborow, R2MPCol)
        IntLat = Sheets(combosn).Cells(comborow, LatCol)
        IntLong = Sheets(combosn).Cells(comborow, LongCol)
        IntCont = Sheets(combosn).Cells(comborow, IntContCol)
        EntVeh = Sheets(combosn).Cells(comborow, EntVehCol)
        MaxFC = Sheets(combosn).Cells(comborow, MaxFCCol)
        MinFC = Sheets(combosn).Cells(comborow, MinFCCol)
        MaxLanes = Sheets(combosn).Cells(comborow, MaxLanesCol)
        MinLanes = Sheets(combosn).Cells(comborow, MinLanesCol)
        MaxSL = Sheets(combosn).Cells(comborow, MaxSLCol)
        MinSL = Sheets(combosn).Cells(comborow, MinSLCol)
        TCrash = Sheets(combosn).Cells(comborow, TCrashCol)
        PCrash = Sheets(combosn).Cells(comborow, PCrashCol)
        ACrash = Sheets(combosn).Cells(comborow, ACrashCol)
        Sev5Crash = Sheets(combosn).Cells(comborow, Sev5CrashCol)
        Sev4Crash = Sheets(combosn).Cells(comborow, Sev4CrashCol)
        Sev3Crash = Sheets(combosn).Cells(comborow, Sev3CrashCol)
        Sev2Crash = Sheets(combosn).Cells(comborow, Sev2CrashCol)
        Sev1Crash = Sheets(combosn).Cells(comborow, Sev1CrashCol)
        NumYears = Sheets(combosn).Cells(comborow, NumYearsCol)
        
        'Assign initial values from parameters/crash sheet
        CrashSevs = Sheets(crashsn).Cells(CSevRow, 2)
        FAMethod = Sheets(crashsn).Cells(FARow, 2)
        DataYears = Sheets(crashsn).Cells(DataYearsRow, 2)
        
        Sheets("IntBlankReport").Activate
        Sheets("IntBlankReport").Copy After:=Sheets("IntBlankReport")
        CurrentReport = modeltype & "-StateRank" & CStr(SRank) & "-Region" & CStr(Region)
        ActiveSheet.Name = CurrentReport
        
        'Enter date of analysis
        reportrow = 1
        Do Until Sheets(CurrentReport).Cells(reportrow, 1) = "Intersection Identification and Roadway Characteristics"
            reportrow = reportrow + 1
        Loop
        Sheets(CurrentReport).Cells(reportrow, 9) = Date
        
        'Enter values in Table 1
        Do Until Left(Sheets(CurrentReport).Cells(reportrow, 1), 7) = "Table 1"
            reportrow = reportrow + 1
        Loop
        Sheets(CurrentReport).Cells(reportrow + 1, 3) = Route1
        Sheets(CurrentReport).Cells(reportrow + 1, 4) = Round(R1MP, 2)
        Sheets(CurrentReport).Cells(reportrow + 2, 3) = Route2
        Sheets(CurrentReport).Cells(reportrow + 2, 4) = Round(R2MP, 2)
        If Route3 = 0 Then
            Sheets(CurrentReport).Cells(reportrow + 3, 3) = "N/A"
        Else
            Sheets(CurrentReport).Cells(reportrow + 3, 3) = Route3
            Sheets(CurrentReport).Cells(reportrow + 3, 4) = Round(R3MP, 2)
        End If
        
        Sheets(CurrentReport).Cells(reportrow + 4, 3) = Round(IntLat, 5)
        Sheets(CurrentReport).Cells(reportrow + 4, 4) = Round(IntLong, 4)
        Sheets(CurrentReport).Cells(reportrow + 5, 3) = DataYears
        
        Sheets(CurrentReport).Cells(reportrow + 1, 8) = modeltype
        Sheets(CurrentReport).Cells(reportrow + 2, 8) = SRank
        Sheets(CurrentReport).Cells(reportrow + 3, 8) = Region
        Sheets(CurrentReport).Cells(reportrow + 3, 9) = RRank
        Sheets(CurrentReport).Cells(reportrow + 4, 8) = County
        Sheets(CurrentReport).Cells(reportrow + 4, 9) = CRank
        Sheets(CurrentReport).Cells(reportrow + 5, 8) = City
        
        'Enter values in Table 2
        Do Until Left(Sheets(CurrentReport).Cells(reportrow, 1), 7) = "Table 2"
            reportrow = reportrow + 1
        Loop
        If IntCont = 1 Then
            Sheets(CurrentReport).Cells(reportrow + 1, 3) = "Signal"
        Else
            Sheets(CurrentReport).Cells(reportrow + 1, 3) = "Stop"
        End If
        Sheets(CurrentReport).Cells(reportrow + 2, 3) = MaxFC
        Sheets(CurrentReport).Cells(reportrow + 3, 3) = MaxLanes
        Sheets(CurrentReport).Cells(reportrow + 4, 3) = MaxSL
        
        Sheets(CurrentReport).Cells(reportrow + 1, 8) = Format(EntVeh, "###,###")
        Sheets(CurrentReport).Cells(reportrow + 2, 8) = MinFC
        Sheets(CurrentReport).Cells(reportrow + 3, 8) = MinLanes
        Sheets(CurrentReport).Cells(reportrow + 4, 8) = MinSL
        
        'Enter values in Table 3
        Do Until Left(Sheets(CurrentReport).Cells(reportrow, 1), 7) = "Table 3"
            reportrow = reportrow + 1
        Loop
        Sheets(CurrentReport).Cells(reportrow + 3, 1) = CrashSevs
        Sheets(CurrentReport).Cells(reportrow + 3, 2) = FAMethod
        Sheets(CurrentReport).Cells(reportrow + 3, 4) = PCrash
        Sheets(CurrentReport).Cells(reportrow + 3, 5) = ACrash
        Sheets(CurrentReport).Cells(reportrow + 3, 6) = Sev5Crash
        Sheets(CurrentReport).Cells(reportrow + 3, 7) = Sev4Crash
        Sheets(CurrentReport).Cells(reportrow + 3, 8) = Sev3Crash
        Sheets(CurrentReport).Cells(reportrow + 3, 9) = Sev2Crash
        Sheets(CurrentReport).Cells(reportrow + 3, 10) = Sev1Crash
        
        'Cycle through crash data. Find intersection ID and add crash data to Tables 4 and 5
        'Find Table 4
        Do Until Left(Sheets(CurrentReport).Cells(reportrow, 1), 7) = "Table 4"
            reportrow = reportrow + 1
        Loop
        
        crashrow = CrashRow1 + 1
        Do Until Sheets(crashsn).Cells(crashrow, CrashIntIDCol) = IntID
            crashrow = crashrow + 1
        Loop
        endcrashrow = crashrow
        Do Until Sheets(crashsn).Cells(endcrashrow + 1, CrashIntIDCol) <> IntID
            endcrashrow = endcrashrow + 1
        Loop
        
        'Determine top 8 crash factors in order
        Sheets("SegKey").Range("G2:H100").ClearContents                     'Clear previous crash factor data
        
        nrow = 2
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, HeadOnCol)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, HeadOnCol)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, WorkZoneCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, WorkZoneCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, PedCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, PedCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, BikeCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, BikeCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, ImpRestCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, ImpRestCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, UnrestCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, UnrestCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, DUICol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, DUICol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, AggCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, AggCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, DistrCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, DistrCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, DrowsyCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, DrowsyCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, SpeedCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, SpeedCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, AdWeathCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, AdWeathCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, AdRoadSurfCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, AdRoadSurfCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, RoadGeomCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, RoadGeomCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, WildAnCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, WildAnCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, DomAnCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, DomAnCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, RoadDepCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, RoadDepCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, OverturnCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, OverturnCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, CommMotVehCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, CommMotVehCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, InterstateCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, InterstateCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, TeenDrivCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, TeenDrivCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, OldDrivCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, OldDrivCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, NightDarkCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, NightDarkCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, SingleVehCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, SingleVehCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, TrainInvCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, TrainInvCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, RailCrossCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, RailCrossCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, TransVehCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, TransVehCol2)
        nrow = nrow + 1
        Sheets("SegKey").Cells(nrow, 7) = Sheets(combosn).Cells(1, FixedObjCol2)
        Sheets("SegKey").Cells(nrow, 8) = Sheets(combosn).Cells(comborow, FixedObjCol2)
        
        ActiveWorkbook.Worksheets("SegKey").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("SegKey").Sort.SortFields.Add Key:=Range("H2:H" & nrow) _
            , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("SegKey").Sort
            .SetRange Range("G2:H" & nrow)
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        Sheets(CurrentReport).Activate
        reportrow2 = reportrow + 2
        For i = 1 To endcrashrow - crashrow
            Rows(reportrow2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlDiagonalDown).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlDiagonalUp).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeLeft).LineStyle = xlNone
            With Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.499984740745262
                .Weight = xlThin
            End With
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeRight).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlInsideVertical).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlInsideHorizontal).LineStyle = xlNone
        Next i
        
        For i = 0 To (endcrashrow - crashrow)
            Sheets(CurrentReport).Cells(reportrow + 2 + i, 1) = Sheets(crashsn).Cells(crashrow + i, CrashIDCol)
            Sheets(CurrentReport).Cells(reportrow + 2 + i, 2) = Round(Sheets(crashsn).Cells(crashrow + i, CLatCol), 4)
            Sheets(CurrentReport).Cells(reportrow + 2 + i, 3) = Round(Sheets(crashsn).Cells(crashrow + i, CLongCol), 3)
        Next i
        
        For i = 2 To 8
            Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = Replace(Replace(Sheets("SegKey").Cells(i, 7), "_TotalCount", ""), "_", " ")
            Sheets(CurrentReport).Cells(reportrow + 3 + endcrashrow - crashrow, 2 + i) = "'" & CStr(Sheets("SegKey").Cells(i, 8)) & "/" & CStr(ACrash)
            
            reportrow2 = reportrow + 2
            If Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "HEADON COLLISION" Then
                For j = 0 To endcrashrow - crashrow
                    If Sheets(crashsn).Cells(crashrow + j, ManColCol) = 3 Then
                        Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = "Y"
                    Else
                        Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = "N"
                    End If
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "WORKZONE RELATED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, WorkZoneCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "PEDESTRIAN INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, PedCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "BICYCLIST INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, BikeCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "MOTORCYCLE INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, MotoCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "IMPROPER RESTRAINT" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, ImpRestCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "UNRESTRAINED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, UnrestCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "DUI" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, DUICol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "AGGRESSIVE DRIVING" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, AggCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "DISTRACTED DRIVING" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, DistrCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "DROWSY DRIVING" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, DrowsyCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "SPEED RELATED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, SpeedCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "ADVERSE WEATHER" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, AdWeathCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "ADVERSE ROADWAY SURF CONDITION" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, AdRoadSurfCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "ROADWAY GEOMETRY RELATED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, RoadGeomCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "WILD ANIMAL RELATED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, WildAnCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "DOMESTIC ANIMAL RELATED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, DomAnCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "ROADWAY DEPARTURE" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, RoadDepCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "OVERTURN ROLLOVER" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, OverturnCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "COMMERCIAL MOTOR VEH INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, CommMotVehCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "INTERSTATE HIGHWAY" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, InterstateCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "TEENAGE DRIVER INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, TeenDrivCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "OLDER DRIVER INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, OldDrivCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "NIGHT DARK CONDITION" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, NightDarkCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "SINGLE VEHICLE" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, SingleVehCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "TRAIN INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, TrainInvCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "RAILROAD CROSSING" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, RailCrossCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "TRANSIT VEHICLE INVOLVED" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, TransVehCol)
                Next j
            ElseIf Sheets(CurrentReport).Cells(reportrow + 1, 2 + i) = "COLLISION WITH FIXED OBJECT" Then
                For j = 0 To endcrashrow - crashrow
                    Sheets(CurrentReport).Cells(reportrow + 2 + j, 2 + i) = Sheets(crashsn).Cells(crashrow + j, FixedObjCol)
                Next j
            End If
        Next i
                
        'Find Table 5
        Do Until Left(Sheets(CurrentReport).Cells(reportrow, 1), 7) = "Table 5"
            reportrow = reportrow + 1
        Loop
        
        Sheets(CurrentReport).Activate
        reportrow2 = reportrow + 2
        For i = 1 To endcrashrow - crashrow
            Rows(reportrow2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlDiagonalDown).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlDiagonalUp).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeLeft).LineStyle = xlNone
            With Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.499984740745262
                .Weight = xlThin
            End With
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlEdgeRight).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlInsideVertical).LineStyle = xlNone
            Range("A" & reportrow2 & ":J" & reportrow2).Borders(xlInsideHorizontal).LineStyle = xlNone
        Next i
        
        For j = 0 To (endcrashrow - crashrow)
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 1) = Sheets(crashsn).Cells(crashrow + j, CrashIDCol)
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 2) = Sheets(crashsn).Cells(crashrow + j, NumVehCol)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 1) = Sheets(crashsn).Cells(crashrow + j, FirstHarmCol) Or Sheets("Key").Cells(nrow, 1) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 3) = Sheets("Key").Cells(nrow, 2)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 3) = Sheets(crashsn).Cells(crashrow + j, ManColCol) Or Sheets("Key").Cells(nrow, 3) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 4) = Sheets("Key").Cells(nrow, 4)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 5) = Sheets(crashsn).Cells(crashrow + j, Event1Col) Or Sheets("Key").Cells(nrow, 5) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 5) = Sheets("Key").Cells(nrow, 6)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 5) = Sheets(crashsn).Cells(crashrow + j, Event2Col) Or Sheets("Key").Cells(nrow, 5) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 6) = Sheets("Key").Cells(nrow, 6)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 5) = Sheets(crashsn).Cells(crashrow + j, Event3Col) Or Sheets("Key").Cells(nrow, 5) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 7) = Sheets("Key").Cells(nrow, 6)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 5) = Sheets(crashsn).Cells(crashrow + j, Event4Col) Or Sheets("Key").Cells(nrow, 5) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 8) = Sheets("Key").Cells(nrow, 6)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 7) = Sheets(crashsn).Cells(crashrow + j, MostHarmCol) Or Sheets("Key").Cells(nrow, 7) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 9) = Sheets("Key").Cells(nrow, 8)
            
            nrow = 2
            Do Until Sheets("Key").Cells(nrow, 9) = Sheets(crashsn).Cells(crashrow + j, VehManCol) Or Sheets("Key").Cells(nrow, 9) = ""
                nrow = nrow + 1
            Loop
            Sheets(CurrentReport).Cells(reportrow + 2 + j, 10) = Sheets("Key").Cells(nrow, 10)
        Next j
       
        'Add list of possible countermeasures.
       
        ReDim factorarray(1 To 7) As String
        
        nrow = 1
        Do Until Left(Sheets(CurrentReport).Cells(nrow, 1), 7) = "Table 4"
            nrow = nrow + 1
        Loop
        
        For i = 1 To 7
            factorarray(i) = Replace(Sheets(CurrentReport).Cells(nrow + 1, 3 + i), " ", "_")
        Next i
        
        Windows(reportfn).Activate
        Sheets(CurrentReport).Activate
        n = 1
        Do Until Cells(n, 1) = "[begin list here]"
            n = n + 1
        Loop
        n = n + 1
        
        For i = 1 To 7
            Sheets("Key").Activate
            j = FindColumn(factorarray(i))
            nrow = 2
            
            Do
                Sheets(CurrentReport).Cells(n, 1) = Sheets("Key").Cells(nrow, j)
                n = n + 1
                nrow = nrow + 1
            Loop While Sheets("Key").Cells(nrow, j) <> ""
        
        Next
        
        'Remove duplicate entries
        'sort
        Windows(reportfn).Activate
        Sheets(CurrentReport).Activate
        n = 1
        Do Until Cells(n, 1) = "[begin list here]"
            n = n + 1
        Loop
        n = n + 1
        nrow = n
        Do While Cells(nrow + 1, 1) <> ""
            nrow = nrow + 1
        Loop
        
        ActiveWorkbook.Worksheets(CurrentReport).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(CurrentReport).Sort.SortFields.Add Key:=Range(Cells(n, 1), Cells(nrow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        With ActiveWorkbook.Worksheets(CurrentReport).Sort
            .SetRange Range(Cells(n, 1), Cells(nrow, 1))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        Dim mycount As Integer
        'step through and delete duplicates
        n = 1
        Do Until Cells(n, 1) = "[begin list here]"
            n = n + 1
            mycount = n + 1
        Loop
        Cells(n, 1) = ""
        n = n + 1
        Do While Cells(n, 1) <> ""
            If Cells(n, 1) = Cells(n + 1, 1) Then
                Rows(n + 1).Delete Shift:=xlUp
            Else
                n = n + 1
            End If
        Loop
       
       'Export sheet as separate workbook.
        Sheets(CurrentReport).Move
        ChDir reportfp
        ActiveWorkbook.SaveAs Filename:=reportfp & "\" & CurrentReport & ".xlsx" _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Sheets(CurrentReport).Copy Before:=Sheets(CurrentReport)
        ActiveSheet.Name = "2Page-" & CurrentReport
        ActiveWorkbook.Save
        ActiveWorkbook.Close

        'Update progress screen
        Sheets("Main").Activate
        Sheets("Main").Range("J5") = Sheets("Main").Range("J5") + 1
        Sheets("Main").Range("J9") = Now()
        Application.ScreenUpdating = True
        Application.ScreenUpdating = False
        
    End If
    comborow = comborow + 1
Loop

Sheets("Main").Activate
Sheets("Main").Range("J9") = ""
Sheets("Main").Range("J11") = Now()
Application.ScreenUpdating = True
'   Waits for 3 seconds, to allow time for user to see the final time
Application.Wait (Now + TimeValue("00:00:03"))

'  Deletes message to user on progress of report compilation
Range("H2:L12").Clear

'   Waits for 1 second, to allow time for the textbox to appear
Application.Wait (Now + TimeValue("00:00:01"))

'  Displays message box with the location of the saved Reports
MsgBox "Reports Completed." & Chr(10) & "Workbook will now close when you click OK" & Chr(10) & "See Analysis Results in the following folder:" & Chr(10) & reportfp, vbOKOnly

cmdline = "explorer.exe" & " " & reportfp

Shell cmdline, vbNormalFocus

'  Saves and closes the "Report Compiler (Date & Time)" workbook.
ActiveWorkbook.Save
ActiveWorkbook.Close

End Sub

Function FindColumn(cname As String) As Integer

FindColumn = 1
Do Until Cells(1, FindColumn) = cname
    FindColumn = FindColumn + 1
Loop

End Function

Function FindPColumn(cname As String, ColRow As Integer) As Integer

FindPColumn = 1
Do Until Cells(ColRow, FindPColumn) = cname
    FindPColumn = FindPColumn + 1
Loop

End Function

Function UDOTRegion(County As String) As String

Dim i As Integer

For i = 2 To 30
    If UCase(Sheets("Key").Cells(i, 11)) = UCase(County) Then
        UDOTRegion = CStr(Sheets("Key").Cells(i, 12))
        Exit For
    End If
Next i

End Function

Function FirstHarmful(n As Integer) As String

Dim i As Integer

For i = 2 To 60
    If Sheets("Key").Cells(i, 1) = n Then
        FirstHarmful = Sheets("Key").Cells(i, 2)
        Exit For
    End If
Next i

End Function

Function EventSequence(n As Integer) As String

Dim i As Integer

For i = 2 To 60
    If Sheets("Key").Cells(i, 5) = n Then
        EventSequence = Sheets("Key").Cells(i, 6)
        Exit For
    End If
Next i

End Function

Function MostHarmful(n As Integer) As String

Dim i As Integer

For i = 2 To 60
    If Sheets("Key").Cells(i, 7) = n Then
        MostHarmful = Sheets("Key").Cells(i, 8)
        Exit For
    End If
Next i

End Function

Function MannerCollision(n As Integer) As String

Dim i As Integer

For i = 2 To 15
    If Sheets("Key").Cells(i, 3) = n Then
        MannerCollision = Sheets("Key").Cells(i, 4)
        Exit For
    End If
Next i

End Function

Function VehicleManuever(n As Integer) As String

Dim i As Integer

For i = 2 To 25
    If Sheets("Key").Cells(i, 9) = n Then
        VehicleManuever = Sheets("Key").Cells(i, 10)
        Exit For
    End If
Next i

End Function


Sub FilterbyCrashID()
'The purpose of this function is to sort the large vehicle crash database to those needed for a given segment
'   Comments by Sam Mineer, Brigham Young Univeristy

Dim crashstring() As String     'An array, which uses the CrashID as the string values
Dim n As Integer                'row counter, to dimension the depth of the array of strings
Dim nv As Double                'row counter, to determine the max row of vehicle data
Dim jv As Double                'column counter, to determine the max column of vehicle data
Dim icrash As Double            'column counter, to find the column with "Crash_ID", if not the first row

Sheets("FilterByCrash").Activate          '<--- Verify this is correct, that sheet exists in crashdata
n = 1       'determines the number of crashID's to put into filter, dimenstions the depth of the array
Do While Cells(n, 1) <> ""
    n = n + 1
Loop

ReDim crashstring(0 To n) As String

'With the string array correctly dimensioned, now the values in the string array can be defined
n = 1
Do While Cells(n, 1) <> ""
    crashstring(n - 1) = CStr(Cells(n, 1))
    n = n + 1
Loop

'return to "vehicle" sheet, of Crash Database
Sheets("vehicle").Activate                    '<--- Verify this is correct, that sheet exists in crashdata

'Finds the extends of crash data, for the filter to be applied
nv = 1
Do Until Cells(nv, 1) = ""
    nv = nv + 1
Loop
jv = 1
Do Until Cells(1, jv) = ""
    jv = jv + 1
Loop

'identifies the column with the CrashID
icrash = 1
Do Until LCase(Left(Cells(1, icrash), 5)) = "crash"
    icrash = icrash + 1
Loop

'Turns off the filter, reapplies the filter with
ActiveSheet.Range(Cells(1, 1), Cells(nv, jv)).AutoFilter Field:=icrash
Application.Wait (Now + TimeValue("00:00:01"))
ActiveSheet.Range(Cells(1, 1), Cells(nv, jv)).AutoFilter Field:=icrash, Criteria1:=Array(crashstring), Operator:=xlFilterValues
    
End Sub

Function Col_Letter(lngCol As Integer) As String
Dim vArr
vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)
End Function


