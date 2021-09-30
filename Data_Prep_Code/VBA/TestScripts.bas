Attribute VB_Name = "TestScripts"

Sub executeBAtest()
' The purpose of this code is to test the Before After analysis without having to fill in the forms each time.
' The input parameters for the statistical scripts are saved on the 'Inputs' sheet.

Dim rscript As String, rcode As String, bawd As String, niter As Long, nburn As Long, datalocation As String

' Define variables
Dim cmdLine As String
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

rscript = Workbooks(guiwb).Sheets("Inputs").Range("F3")
bawd = Workbooks(guiwb).Sheets("Inputs").Range("F2")
rcode = Workbooks(guiwb).Sheets("Inputs").Range("F11")
bawd = bawd & "/" & "BAanalysis_" & replace(Date, "/", "-") & "_" & replace(replace(Time, ":", "-"), " ", "_") ' & "/"
MkDir bawd
niter = Workbooks(guiwb).Sheets("Inputs").Range("F9")
nburn = Workbooks(guiwb).Sheets("Inputs").Range("F10")
datalocation = Workbooks(guiwb).Sheets("Inputs").Range("F8")

' Prepare command line for Rscript execution
cmdLine = rscript & " " & rcode & " " & bawd & " " & niter & " " & nburn & " " & datalocation

' Submit Rscript code execution
'    MsgBox cmdLine, vbOKOnly   'Debugging purposes
Shell cmdLine, vbMaximizedFocus

'MsgBox "Before After Analysis started." & Chr(10) & "Do not shut down computer until analysis is finished."

End Sub

Sub SumCrashDatadup()

'Declare variables
Dim IntRow, CrashRow, CrashRowPrev As Long
Dim IntLatCol, IntLongCol, CrashLatCol, CrashLongCol, UrbanCodeCol, UrbanCode As Long
Dim IntLat, IntLong, LatFtEquiv, LongFtEquiv, MinLatStart, MaxLatEnd As Double
Dim CrashLat, CrashLong, LatDiff, LongDiff, DiffRadius As Double
Dim IntRadius, SpeedLimit, NumYears As Integer
Dim datasheet, CrashSheet, UCDesc, FCDesc As String
Dim row1, col1, CHeadCol As Long

Dim MaxSLCol, MaxFCCol, UCcol As Long
Dim IntNumCol, YearCol, TotCrashCol, SevCrashCol, PedInvCol1, BikeInvCol1, MotInvCol1, ImpResCol1, UnrestCol1, DUICol1 As Long
Dim AggressiveCol1, DistractedCol1, DrowsyCol1, SpeedCol1, IntRelCol1, WeatherCol1, AdvRoadCol1, RoadGeomCol1 As Long
Dim WAnimalCol1, DAnimalCol1, RoadDepartCol1, OverturnCol1, CommVehCol1, InterstateCol1, TeenCol1, OldCol1 As Long
Dim UrbanCol1, NightCol1, SingleVehCol1, TrainCol1, RailCol1, TransitCol1, FixedObCol1, WorkZoneCol1, CollMannerCol1 As Long
Dim MaxRoadWCol, MinRoadWcol, MaxRoadW, MinRoadW As Long
Dim IntYear As String

Dim CrashDateCol, CrashSeverityCol, PedInvCol2, BikeInvCol2, MotInvCol2, ImpResCol2, UnrestCol2, DUICol2 As Long
Dim AggressiveCol2, DistractedCol2, DrowsyCol2, SpeedCol2, IntRelCol2, WeatherCol2, AdvRoadCol2, RoadGeomCol2 As Long
Dim WAnimalCol2, DAnimalCol2, RoadDepartCol2, OverturnCol2, CommVehCol2, InterstateCol2, TeenCol2, OldCol2 As Long
Dim UrbanCol2, NightCol2, SingleVehCol2, TrainCol2, RailCol2, TransitCol2, FixedObCol2, WorkZoneCol2, CollMannerCol2 As Long
Dim CrashYear As String

Dim UICPMcol, FArow, i As Long
Dim sevstring As String
ReDim Sevs(1 To 5) As Integer


'Assign initial values
IntRow = 2
CrashRow = 2

IntLatCol = 1
IntLongCol = 1
CrashLatCol = 1
CrashLongCol = 1

datasheet = "UICPMinput"
CrashSheet = "CrashInput"

IntNumCol = 1
YearCol = 1
MaxSLCol = 1
MaxFCCol = 1
UCcol = 1
TotCrashCol = 1
SevCrashCol = 1
UrbanCodeCol = 1
PedInvCol1 = 1
BikeInvCol1 = 1
MotInvCol1 = 1
ImpResCol1 = 1
UnrestCol1 = 1
DUICol1 = 1
AggressiveCol1 = 1
DistractedCol1 = 1
DrowsyCol1 = 1
SpeedCol1 = 1
IntRelCol1 = 1
WeatherCol1 = 1
AdvRoadCol1 = 1
RoadGeomCol1 = 1
WAnimalCol1 = 1
DAnimalCol1 = 1
RoadDepartCol1 = 1
OverturnCol1 = 1
CommVehCol1 = 1
InterstateCol1 = 1
TeenCol1 = 1
OldCol1 = 1
UrbanCol1 = 1
NightCol1 = 1
SingleVehCol1 = 1
TrainCol1 = 1
RailCol1 = 1
TransitCol1 = 1
FixedObCol1 = 1
WorkZoneCol1 = 1
CollMannerCol1 = 1
MaxRoadWCol = 1
MinRoadWcol = 1

CrashDateCol = 1
CrashSeverityCol = 1
PedInvCol2 = 1
BikeInvCol2 = 1
MotInvCol2 = 1
ImpResCol2 = 1
UnrestCol2 = 1
DUICol2 = 1
AggressiveCol2 = 1
DistractedCol2 = 1
DrowsyCol2 = 1
SpeedCol2 = 1
IntRelCol2 = 1
WeatherCol2 = 1
AdvRoadCol2 = 1
RoadGeomCol2 = 1
WAnimalCol2 = 1
DAnimalCol2 = 1
RoadDepartCol2 = 1
OverturnCol2 = 1
CommVehCol2 = 1
InterstateCol2 = 1
TeenCol2 = 1
OldCol2 = 1
UrbanCol2 = 1
NightCol2 = 1
SingleVehCol2 = 1
TrainCol2 = 1
RailCol2 = 1
TransitCol2 = 1
FixedObCol2 = 1
WorkZoneCol2 = 1
CollMannerCol2 = 1

UICPMcol = 1
FArow = 1


'Find first blank column in roadway intersection data and add in column headings
row1 = 9
col1 = 1
CHeadCol = 1
Do Until Sheets(datasheet).Cells(1, col1) = ""
    col1 = col1 + 1
Loop

Do Until Sheets("Key").Cells(1, CHeadCol) = "Intersection Check Headers"
    CHeadCol = CHeadCol + 1
Loop

Do Until Sheets("Key").Cells(2, CHeadCol) = 2
    CHeadCol = CHeadCol + 1
Loop

Do Until Sheets("Key").Cells(row1, CHeadCol + 2) = ""
    Sheets(datasheet).Cells(1, col1) = Sheets("Key").Cells(row1, CHeadCol + 2)
    col1 = col1 + 1
    row1 = row1 + 1
Loop

'Find columns
Do Until Sheets(datasheet).Cells(1, IntNumCol) = "INT_ID"
    IntNumCol = IntNumCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, YearCol) = "YEAR"
    YearCol = YearCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, MaxSLCol) = "MAX_SPEED_LIMIT"
    MaxSLCol = MaxSLCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, MaxFCCol) = "MAX FC_TYPE"
    MaxFCCol = MaxFCCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, UCcol) = "URBAN_DESC"
    UCcol = UCcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, IntLatCol) = "LATITUDE"
    IntLatCol = IntLatCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, IntLongCol) = "LONGITUDE"
    IntLongCol = IntLongCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CrashLatCol) = "LATITUDE"
    CrashLatCol = CrashLatCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CrashLongCol) = "LONGITUDE"
    CrashLongCol = CrashLongCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, MaxRoadWCol) = "MAX_ROAD_WIDTH"
    MaxRoadWCol = MaxRoadWCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, MinRoadWcol) = "MIN_ROAD_WIDTH"
    MinRoadWcol = MinRoadWcol + 1
Loop


Do Until Sheets(datasheet).Cells(1, TotCrashCol) = "Total_Crashes"
    TotCrashCol = TotCrashCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SevCrashCol) = "Severe_Crashes"
    SevCrashCol = SevCrashCol + 1
Loop

UrbanCodeCol = 1
Do Until Sheets(datasheet).Cells(1, UrbanCodeCol) = "URBAN_CODE"
    UrbanCodeCol = UrbanCodeCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, WorkZoneCol1) = "WORKZONE_RELATED"
    WorkZoneCol1 = WorkZoneCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, CollMannerCol1) = "HEADON_COLLISION"
    CollMannerCol1 = CollMannerCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, PedInvCol1) = "PEDESTRIAN_INVOLVED"
    PedInvCol1 = PedInvCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, BikeInvCol1) = "BICYCLIST_INVOLVED"
    BikeInvCol1 = BikeInvCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, MotInvCol1) = "MOTORCYCLE_INVOLVED"
    MotInvCol1 = MotInvCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, ImpResCol1) = "IMPROPER_RESTRAINT"
    ImpResCol1 = ImpResCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, UnrestCol1) = "UNRESTRAINED"
    UnrestCol1 = UnrestCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DUICol1) = "DUI"
    DUICol1 = DUICol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, AggressiveCol1) = "AGGRESSIVE_DRIVING"
    AggressiveCol1 = AggressiveCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DistractedCol1) = "DISTRACTED_DRIVING"
    DistractedCol1 = DistractedCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DrowsyCol1) = "DROWSY_DRIVING"
    DrowsyCol1 = DrowsyCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, SpeedCol1) = "SPEED_RELATED"
    SpeedCol1 = SpeedCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, IntRelCol1) = "INTERSECTION_RELATED"
    IntRelCol1 = IntRelCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, WeatherCol1) = "ADVERSE_WEATHER"
    WeatherCol1 = WeatherCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, AdvRoadCol1) = "ADVERSE_ROADWAY_SURF_CONDITION"
    AdvRoadCol1 = AdvRoadCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, RoadGeomCol1) = "ROADWAY_GEOMETRY_RELATED"
    RoadGeomCol1 = RoadGeomCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, WAnimalCol1) = "WILD_ANIMAL_RELATED"
    WAnimalCol1 = WAnimalCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, DAnimalCol1) = "DOMESTIC_ANIMAL_RELATED"
    DAnimalCol1 = DAnimalCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, RoadDepartCol1) = "ROADWAY_DEPARTURE"
    RoadDepartCol1 = RoadDepartCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, OverturnCol1) = "OVERTURN_ROLLOVER"
    OverturnCol1 = OverturnCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, CommVehCol1) = "COMMERCIAL_MOTOR_VEH_INVOLVED"
    CommVehCol1 = CommVehCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, InterstateCol1) = "INTERSTATE_HIGHWAY"
    InterstateCol1 = InterstateCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, TeenCol1) = "TEENAGE_DRIVER_INVOLVED"
    TeenCol1 = TeenCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, OldCol1) = "OLDER_DRIVER_INVOLVED"
    OldCol1 = OldCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, UrbanCol1) = "URBAN_COUNTY"
    UrbanCol1 = UrbanCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, NightCol1) = "NIGHT_DARK_CONDITION"
    NightCol1 = NightCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, SingleVehCol1) = "SINGLE_VEHICLE"
    SingleVehCol1 = SingleVehCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, TrainCol1) = "TRAIN_INVOLVED"
    TrainCol1 = TrainCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, RailCol1) = "RAILROAD_CROSSING"
    RailCol1 = RailCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, TransitCol1) = "TRANSIT_VEHICLE_INVOLVED"
    TransitCol1 = TransitCol1 + 1
Loop

Do Until Sheets(datasheet).Cells(1, FixedObCol1) = "COLLISION_WITH_FIXED_OBJECT"
    FixedObCol1 = FixedObCol1 + 1
Loop


Do Until Sheets(CrashSheet).Cells(1, CrashDateCol) = "CRASH_DATETIME"
    CrashDateCol = CrashDateCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CrashSeverityCol) = "CRASH_SEVERITY_ID"
    CrashSeverityCol = CrashSeverityCol + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, WorkZoneCol2) = "WORK_ZONE_RELATED_YNU"
    WorkZoneCol2 = WorkZoneCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CollMannerCol2) = "MANNER_COLLISION_ID"
    CollMannerCol2 = CollMannerCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, PedInvCol2) = "PEDESTRIAN_INVOLVED"
    PedInvCol2 = PedInvCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, BikeInvCol2) = "BICYCLIST_INVOLVED"
    BikeInvCol2 = BikeInvCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, MotInvCol2) = "MOTORCYCLE_INVOLVED"
    MotInvCol2 = MotInvCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, ImpResCol2) = "IMPROPER_RESTRAINT"
    ImpResCol2 = ImpResCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, UnrestCol2) = "UNRESTRAINED"
    UnrestCol2 = UnrestCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DUICol2) = "DUI"
    DUICol2 = DUICol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, AggressiveCol2) = "AGGRESSIVE_DRIVING"
    AggressiveCol2 = AggressiveCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DistractedCol2) = "DISTRACTED_DRIVING"
    DistractedCol2 = DistractedCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DrowsyCol2) = "DROWSY_DRIVING"
    DrowsyCol2 = DrowsyCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, SpeedCol2) = "SPEED_RELATED"
    SpeedCol2 = SpeedCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, IntRelCol2) = "INTERSECTION_RELATED"
    IntRelCol2 = IntRelCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, WeatherCol2) = "ADVERSE_WEATHER"
    WeatherCol2 = WeatherCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, AdvRoadCol2) = "ADVERSE_ROADWAY_SURF_CONDITION"
    AdvRoadCol2 = AdvRoadCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, RoadGeomCol2) = "ROADWAY_GEOMETRY_RELATED"
    RoadGeomCol2 = RoadGeomCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, WAnimalCol2) = "WILD_ANIMAL_RELATED"
    WAnimalCol2 = WAnimalCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, DAnimalCol2) = "DOMESTIC_ANIMAL_RELATED"
    DAnimalCol2 = DAnimalCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, RoadDepartCol2) = "ROADWAY_DEPARTURE"
    RoadDepartCol2 = RoadDepartCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, OverturnCol2) = "OVERTURN_ROLLOVER"
    OverturnCol2 = OverturnCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, CommVehCol2) = "COMMERCIAL_MOTOR_VEH_INVOLVED"
    CommVehCol2 = CommVehCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, InterstateCol2) = "INTERSTATE_HIGHWAY"
    InterstateCol2 = InterstateCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, TeenCol2) = "TEENAGE_DRIVER_INVOLVED"
    TeenCol2 = TeenCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, OldCol2) = "OLDER_DRIVER_INVOLVED"
    OldCol2 = OldCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, UrbanCol2) = "URBAN_COUNTY"
    UrbanCol2 = UrbanCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, NightCol2) = "NIGHT_DARK_CONDITION"
    NightCol2 = NightCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, SingleVehCol2) = "SINGLE_VEHICLE"
    SingleVehCol2 = SingleVehCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, TrainCol2) = "TRAIN_INVOLVED"
    TrainCol2 = TrainCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, RailCol2) = "RAILROAD_CROSSING"
    RailCol2 = RailCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, TransitCol2) = "TRANSIT_VEHICLE_INVOLVED"
    TransitCol2 = TransitCol2 + 1
Loop

Do Until Sheets(CrashSheet).Cells(1, FixedObCol2) = "COLLISION_WITH_FIXED_OBJECT"
    FixedObCol2 = FixedObCol2 + 1
Loop


Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "UICPM"
    UICPMcol = UICPMcol + 1
Loop

Do Until Sheets("Inputs").Cells(FArow, UICPMcol) = "Selected FA Parameter"
    FArow = FArow + 1
Loop


LatLongElevSortUICPM

Sheets(datasheet).Activate

'Find number of years of data
NumYears = 1
Do Until Sheets(datasheet).Cells(IntRow + NumYears, YearCol) = Sheets(datasheet).Cells(IntRow, YearCol)
    NumYears = NumYears + 1
Loop

'Insert zeros for all values
IntRow = 2
Do Until Sheets(datasheet).Cells(IntRow, 1) = ""
    Sheets(datasheet).Cells(IntRow, TotCrashCol) = 0
    Sheets(datasheet).Cells(IntRow, SevCrashCol) = 0
    Sheets(datasheet).Cells(IntRow, PedInvCol1) = 0
    Sheets(datasheet).Cells(IntRow, BikeInvCol1) = 0
    Sheets(datasheet).Cells(IntRow, MotInvCol1) = 0
    Sheets(datasheet).Cells(IntRow, ImpResCol1) = 0
    Sheets(datasheet).Cells(IntRow, UnrestCol1) = 0
    Sheets(datasheet).Cells(IntRow, DUICol1) = 0
    Sheets(datasheet).Cells(IntRow, AggressiveCol1) = 0
    Sheets(datasheet).Cells(IntRow, DistractedCol1) = 0
    Sheets(datasheet).Cells(IntRow, DrowsyCol1) = 0
    Sheets(datasheet).Cells(IntRow, SpeedCol1) = 0
    Sheets(datasheet).Cells(IntRow, IntRelCol1) = 0
    Sheets(datasheet).Cells(IntRow, WeatherCol1) = 0
    Sheets(datasheet).Cells(IntRow, AdvRoadCol1) = 0
    Sheets(datasheet).Cells(IntRow, RoadGeomCol1) = 0
    Sheets(datasheet).Cells(IntRow, WAnimalCol1) = 0
    Sheets(datasheet).Cells(IntRow, DAnimalCol1) = 0
    Sheets(datasheet).Cells(IntRow, RoadDepartCol1) = 0
    Sheets(datasheet).Cells(IntRow, OverturnCol1) = 0
    Sheets(datasheet).Cells(IntRow, CommVehCol1) = 0
    Sheets(datasheet).Cells(IntRow, InterstateCol1) = 0
    Sheets(datasheet).Cells(IntRow, TeenCol1) = 0
    Sheets(datasheet).Cells(IntRow, OldCol1) = 0
    Sheets(datasheet).Cells(IntRow, UrbanCol1) = 0
    Sheets(datasheet).Cells(IntRow, NightCol1) = 0
    Sheets(datasheet).Cells(IntRow, SingleVehCol1) = 0
    Sheets(datasheet).Cells(IntRow, TrainCol1) = 0
    Sheets(datasheet).Cells(IntRow, RailCol1) = 0
    Sheets(datasheet).Cells(IntRow, TransitCol1) = 0
    Sheets(datasheet).Cells(IntRow, FixedObCol1) = 0
    Sheets(datasheet).Cells(IntRow, WorkZoneCol1) = 0
    Sheets(datasheet).Cells(IntRow, CollMannerCol1) = 0
    IntRow = IntRow + 1
Loop

'Loop through intersections and assign crashes that fit within the given radius
CrashRowPrev = 2
IntRow = 2
Do While Sheets(datasheet).Cells(IntRow, 1) <> ""
    IntLat = Sheets(datasheet).Cells(IntRow, IntLatCol)
    IntLong = Sheets(datasheet).Cells(IntRow, IntLongCol)
    MaxRoadW = Sheets(datasheet).Cells(IntRow, MaxRoadWCol)
    MinRoadW = Sheets(datasheet).Cells(IntRow, MinRoadWcol)
    
    'Calculate the distance in feet for each lat/long degree for the intersection
    LatFtEquiv = 69.172 * 5280
    LongFtEquiv = Cos(IntLat) * 69.172 * 5280
    
    'Determine functional area value
    If Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Speed Limit" Then
        SpeedLimit = Sheets(datasheet).Cells(IntRow, MaxSLCol)
        For i = 1 To 12
            If SpeedLimit = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                IntRadius = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                Exit For
            End If
        Next i
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Functional Class" Then
        FCDesc = Sheets(datasheet).Cells(IntRow, MaxFCCol)
        For i = 1 To 4
            If FCDesc = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol) Then
                IntRadius = Sheets("Inputs").Cells(FArow + 1 + i, UICPMcol + 1)
                Exit For
            End If
        Next i
    ElseIf Sheets("Inputs").Cells(FArow + 1, UICPMcol) = "Urban Code" Then
        UCDesc = Sheets(datasheet).Cells(IntRow, UCcol)
        
        If UCDesc = "Rurall" Then
            IntRadius = Sheets("Inputs").Cells(FArow + 2, UICPMcol + 1)
        ElseIf UCDesc = "Small Urban" Then
            IntRadius = Sheets("Inputs").Cells(FArow + 3, UICPMcol + 1)
        Else
            IntRadius = Sheets("Inputs").Cells(FArow + 4, UICPMcol + 1)
        End If
    End If
    
    'Add intersection size to intersection radius
    IntRadius = IntRadius + Sqr(MaxRoadW * MinRoadW)
    
    'Determine severities
    sevstring = Sheets("Inputs").Cells(FArow - 1, UICPMcol + 1)
    For i = 1 To Len(sevstring)
        Sevs(i) = Mid(sevstring, i, 1)
    Next i
    
    MinLatStart = IntLat - (IntRadius / LatFtEquiv)
    MaxLatEnd = IntLat + (IntRadius / LatFtEquiv)
    
    'Go to crash row that is close to range of intersection to save time
    CrashRow = CrashRowPrev
    Do Until Sheets(CrashSheet).Cells(CrashRow, CrashLatCol) >= MinLatStart
        CrashRow = CrashRow + 1
    Loop
    
    CrashRowPrev = CrashRow
    
    'Assign crashes to intersection
    Do Until Sheets(CrashSheet).Cells(CrashRow, CrashLatCol) > MaxLatEnd
        CrashLat = Sheets(CrashSheet).Cells(CrashRow, CrashLatCol)
        CrashLong = Sheets(CrashSheet).Cells(CrashRow, CrashLongCol)
        LatDiff = Abs(CrashLat - IntLat) * LatFtEquiv
        LongDiff = Abs(CrashLong - IntLong) * LongFtEquiv
        DiffRadius = Sqr(LatDiff ^ 2 + LongDiff ^ 2)
        CrashYear = Year(Sheets(CrashSheet).Cells(CrashRow, CrashDateCol))
        
        If DiffRadius < IntRadius And (Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(1) Or _
        Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(2) Or Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(3) Or _
        Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(4) Or Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) = Sevs(5)) Then
        
            
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If IntYear = CrashYear Then
                    'Total crahes
                    Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = Sheets(datasheet).Cells(IntRow + i, TotCrashCol) + 1
                    
                    'Severe crashes
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) >= 3 And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = Sheets(datasheet).Cells(IntRow + i, SevCrashCol) + 1
                End If
            Next i

                    
                    
                    
                End If
            Next i

        
        
            'Total crahes
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, TotCrashCol) = Sheets(datasheet).Cells(IntRow + i, TotCrashCol) + 1
                End If
            Next i
                        
            'Severe crashes
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, CrashSeverityCol) >= 3 And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, SevCrashCol) = Sheets(datasheet).Cells(IntRow + i, SevCrashCol) + 1
                End If
            Next i
                
            'Head on collisions
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
           
                If Sheets(CrashSheet).Cells(CrashRow, CollMannerCol2) = 3 And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) = Sheets(datasheet).Cells(IntRow + i, CollMannerCol1) + 1
                End If
            Next i
                
                
            'Work zone related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, WorkZoneCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) = Sheets(datasheet).Cells(IntRow + i, WorkZoneCol1) + 1
                End If
            Next i
                
            'Pedestrian Involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, PedInvCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, PedInvCol1) = Sheets(datasheet).Cells(IntRow + i, PedInvCol1) + 1
                End If
            Next i
                
            'Bicyclist Involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, BikeInvCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) = Sheets(datasheet).Cells(IntRow + i, BikeInvCol1) + 1
                End If
            Next i
                
            'Motorcycle Involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, MotInvCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, MotInvCol1) = Sheets(datasheet).Cells(IntRow + i, MotInvCol1) + 1
                End If
            Next i
                
            'Improper restraint
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, ImpResCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, ImpResCol1) = Sheets(datasheet).Cells(IntRow + i, ImpResCol1) + 1
                End If
            Next i
                
            'Unrestrained
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, UnrestCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, UnrestCol1) = Sheets(datasheet).Cells(IntRow + i, UnrestCol1) + 1
                End If
            Next i
                
            'DUI
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DUICol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, DUICol1) = Sheets(datasheet).Cells(IntRow + i, DUICol1) + 1
                End If
            Next i
                
            'Aggresive Driving
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, AggressiveCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) = Sheets(datasheet).Cells(IntRow + i, AggressiveCol1) + 1
                End If
            Next i
                
            'Distracted Driving
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DistractedCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, DistractedCol1) = Sheets(datasheet).Cells(IntRow + i, DistractedCol1) + 1
                End If
            Next i
                
            'Drowsy Driving
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DrowsyCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) = Sheets(datasheet).Cells(IntRow + i, DrowsyCol1) + 1
                End If
            Next i
                
            'Speed related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, SpeedCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, SpeedCol1) = Sheets(datasheet).Cells(IntRow + i, SpeedCol1) + 1
                End If
            Next i
                
            'Intersection related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, IntRelCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, IntRelCol1) = Sheets(datasheet).Cells(IntRow + i, IntRelCol1) + 1
                End If
            Next i
                
            'Adverse weather
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, WeatherCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, WeatherCol1) = Sheets(datasheet).Cells(IntRow + i, WeatherCol1) + 1
                End If
            Next i
                
            'Adverse roadway surface conditions
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, AdvRoadCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) = Sheets(datasheet).Cells(IntRow + i, AdvRoadCol1) + 1
                End If
            Next i
                
            'Roadway geometry related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, RoadGeomCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) = Sheets(datasheet).Cells(IntRow + i, RoadGeomCol1) + 1
                End If
            Next i
                
            'Wild animal related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, WAnimalCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) = Sheets(datasheet).Cells(IntRow + i, WAnimalCol1) + 1
                End If
            Next i
                
            'Domestic animal related
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, DAnimalCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) = Sheets(datasheet).Cells(IntRow + i, DAnimalCol1) + 1
                End If
            Next i
                
            'Roadway departure
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, RoadDepartCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) = Sheets(datasheet).Cells(IntRow + i, RoadDepartCol1) + 1
                End If
            Next i
                
            'Overturn/rollover
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, OverturnCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, OverturnCol1) = Sheets(datasheet).Cells(IntRow + i, OverturnCol1) + 1
                End If
            Next i
                
            'Commercial motor vehicle involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, CommVehCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, CommVehCol1) = Sheets(datasheet).Cells(IntRow + i, CommVehCol1) + 1
                End If
            Next i
                
            'Interstate highway
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, InterstateCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, InterstateCol1) = Sheets(datasheet).Cells(IntRow + i, InterstateCol1) + 1
                End If
            Next i
                
            'Teenage driver involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, TeenCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, TeenCol1) = Sheets(datasheet).Cells(IntRow + i, TeenCol1) + 1
                 End If
            Next i
                
            'Older driver involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, OldCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, OldCol1) = Sheets(datasheet).Cells(IntRow + i, OldCol1) + 1
                End If
            Next i
                
            'Urban county
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, UrbanCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, UrbanCol1) = Sheets(datasheet).Cells(IntRow + i, UrbanCol1) + 1
                End If
            Next i
                
            'Night dark condition
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, NightCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, NightCol1) = Sheets(datasheet).Cells(IntRow + i, NightCol1) + 1
                End If
            Next i
                
            'Single vehicle
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, SingleVehCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) = Sheets(datasheet).Cells(IntRow + i, SingleVehCol1) + 1
                End If
            Next i
                
            'Train involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, TrainCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, TrainCol1) = Sheets(datasheet).Cells(IntRow + i, TrainCol1) + 1
                End If
            Next i
                
            'Railroad crossing
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, RailCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, RailCol1) = Sheets(datasheet).Cells(IntRow + i, RailCol1) + 1
                End If
            Next i
                
            'Transit vehicle involved
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, TransitCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, TransitCol1) = Sheets(datasheet).Cells(IntRow + i, TransitCol1) + 1
                End If
            Next i
                
            'Collision with fixed object
            For i = 0 To NumYears - 1
                IntYear = Sheets(datasheet).Cells(IntRow + i, YearCol).Value
                
                If Sheets(CrashSheet).Cells(CrashRow, FixedObCol2) = "Y" And IntYear = CrashYear Then
                    Sheets(datasheet).Cells(IntRow + i, FixedObCol1) = Sheets(datasheet).Cells(IntRow + i, FixedObCol1) + 1
                End If
            Next i
        End If
        CrashRow = CrashRow + 1
    Loop
    
    IntRow = IntRow + NumYears
Loop


End Sub


Sub executeUCMtest()
' The purpose of this code is to test the UCPM or UCSM without having to fill in the forms each time.
' The input parameters for the statistical scripts are saved on the 'Inputs' sheet.

Dim rscript As String, rcode As String, ucpmwb As String, niter As Long, nburn As Long, datalocation As String

' Define variables
Dim cmdLine As String
Dim guiwb As String
guiwb = ActiveWorkbook.Name
guiwb = replace(guiwb, ".xlsm", "")

rscript = Workbooks(guiwb).Sheets("Inputs").Range("B3")
modelwd = Workbooks(guiwb).Sheets("Inputs").Range("B2")
rcode = Workbooks(guiwb).Sheets("Inputs").Range("B9")
modelwd = modelwd & "/" & "CrashAnalysis_" & replace(Date, "/", "-") & "_" & replace(replace(Time, ":", "-"), " ", "_")
MkDir modelwd
niter = Workbooks(guiwb).Sheets("Inputs").Range("B7")
nburn = Workbooks(guiwb).Sheets("Inputs").Range("B8")
datalocation = Workbooks(guiwb).Sheets("Inputs").Range("B10")
xs = Workbooks(guiwb).Sheets("Inputs").Range("B11")

' Prepare command line for Rscript execution
cmdLine = rscript & " " & rcode & " " & modelwd & " " & niter & " " & nburn & " " & datalocation & " " & xs

' Submit Rscript code execution
'    MsgBox cmdLine, vbOKOnly   'Debugging purposes
Shell cmdLine, vbMaximizedFocus

'MsgBox "Before After Analysis started." & Chr(10) & "Do not shut down computer until analysis is finished."

End Sub
