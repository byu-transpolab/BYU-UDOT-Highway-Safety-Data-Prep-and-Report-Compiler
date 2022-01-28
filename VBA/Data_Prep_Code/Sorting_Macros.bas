Attribute VB_Name = "Sorting_Macros"
Sub LatLongElevSort()
Attribute LatLongElevSort.VB_ProcData.VB_Invoke_Func = " \n14"

'Declare variables
Dim numRows, numCols As Long
Dim latcol, longcol, ElevCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
latcol = 1
longcol = 1
ElevCol = 1
datasheet = "Dataset"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, latcol) = "LATITUDE"
    latcol = latcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, longcol) = "LONGITUDE"
    longcol = longcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, ElevCol) = "ELEVATION"
    ElevCol = ElevCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by Latitude, then Longitude, then Elevation
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(latcol) & "2:" & Col_Letter(latcol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(longcol) & "2:" & Col_Letter(longcol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(ElevCol) & "2:" & Col_Letter(ElevCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub


Sub CrashIDSort()

'Declare variables
Dim numRows, numCols As Long
Dim CrashIDCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
CrashIDCol = 1
datasheet = "Crash_Data"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, CrashIDCol) = "CRASH_ID"
    CrashIDCol = CrashIDCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by Latitude, then Longitude, then Elevation
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(CrashIDCol) & "2:" & Col_Letter(CrashIDCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub


Sub MainRouteMPSort()

'Declare variables
Dim numRows, numCols As Long
Dim MRouteCol, SRouteCol, TRouteCol, Q4RouteCol, Q5RouteCol, MMPCol, SMPCol, TMPCol, Q4MPCol, Q5MPCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
MRouteCol = 1
SRouteCol = 1
TRouteCol = 1
Q4RouteCol = 1
Q5RouteCol = 1
MMPCol = 1
SMPCol = 1
TMPCol = 1
Q4MPCol = 1
Q5MPCol = 1
datasheet = "Dataset"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, MRouteCol) = "ROUTE"
    MRouteCol = MRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SRouteCol) = "INT_RT_1"
    SRouteCol = SRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TRouteCol) = "INT_RT_2"
    TRouteCol = TRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4RouteCol) = "INT_RT_3"
    Q4RouteCol = Q4RouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5RouteCol) = "INT_RT_4"
    Q5RouteCol = Q5RouteCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, MMPCol) = "UDOT_BMP"
    MMPCol = MMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SMPCol) = "INT_RT_1_M"
    SMPCol = SMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TMPCol) = "INT_RT_2_M"
    TMPCol = TMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4MPCol) = "INT_RT_3_M"
    Q4MPCol = Q4MPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5MPCol) = "INT_RT_4_M"
    Q5MPCol = Q5MPCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MRouteCol) & "2:" & Col_Letter(MRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MMPCol) & "2:" & Col_Letter(MMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SRouteCol) & "2:" & Col_Letter(SRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SMPCol) & "2:" & Col_Letter(SMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TRouteCol) & "2:" & Col_Letter(TRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TMPCol) & "2:" & Col_Letter(TMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4RouteCol) & "2:" & Col_Letter(Q4RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4MPCol) & "2:" & Col_Letter(Q4MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5RouteCol) & "2:" & Col_Letter(Q5RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5MPCol) & "2:" & Col_Letter(Q5MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub


Sub SecondaryRouteMPSort()

'Declare variables
Dim numRows, numCols As Long
Dim MRouteCol, SRouteCol, TRouteCol, Q4RouteCol, Q5RouteCol, MMPCol, SMPCol, TMPCol, Q4MPCol, Q5MPCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
MRouteCol = 1
SRouteCol = 1
TRouteCol = 1
Q4RouteCol = 1
Q5RouteCol = 1
MMPCol = 1
SMPCol = 1
TMPCol = 1
Q4MPCol = 1
Q5MPCol = 1
datasheet = "Dataset"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, MRouteCol) = "ROUTE"
    MRouteCol = MRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SRouteCol) = "INT_RT_1"
    SRouteCol = SRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TRouteCol) = "INT_RT_2"
    TRouteCol = TRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4RouteCol) = "INT_RT_3"
    Q4RouteCol = Q4RouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5RouteCol) = "INT_RT_4"
    Q5RouteCol = Q5RouteCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, MMPCol) = "UDOT_BMP"
    MMPCol = MMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SMPCol) = "INT_RT_1_M"
    SMPCol = SMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TMPCol) = "INT_RT_2_M"
    TMPCol = TMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4MPCol) = "INT_RT_3_M"
    Q4MPCol = Q4MPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5MPCol) = "INT_RT_4_M"
    Q5MPCol = Q5MPCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by secondary route/MP, main route/MP, tertiary route/MP, quartenary route/MP, quinary route/MP
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SRouteCol) & "2:" & Col_Letter(SRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SMPCol) & "2:" & Col_Letter(SMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MRouteCol) & "2:" & Col_Letter(MRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MMPCol) & "2:" & Col_Letter(MMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TRouteCol) & "2:" & Col_Letter(TRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TMPCol) & "2:" & Col_Letter(TMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4RouteCol) & "2:" & Col_Letter(Q4RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4MPCol) & "2:" & Col_Letter(Q4MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5RouteCol) & "2:" & Col_Letter(Q5RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5MPCol) & "2:" & Col_Letter(Q5MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub


Sub TertiaryRouteMPSort()

'Declare variables
Dim numRows, numCols As Long
Dim MRouteCol, SRouteCol, TRouteCol, Q4RouteCol, Q5RouteCol, MMPCol, SMPCol, TMPCol, Q4MPCol, Q5MPCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
MRouteCol = 1
SRouteCol = 1
TRouteCol = 1
Q4RouteCol = 1
Q5RouteCol = 1
MMPCol = 1
SMPCol = 1
TMPCol = 1
Q4MPCol = 1
Q5MPCol = 1
datasheet = "Dataset"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, MRouteCol) = "ROUTE"
    MRouteCol = MRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SRouteCol) = "INT_RT_1"
    SRouteCol = SRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TRouteCol) = "INT_RT_2"
    TRouteCol = TRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4RouteCol) = "INT_RT_3"
    Q4RouteCol = Q4RouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5RouteCol) = "INT_RT_4"
    Q5RouteCol = Q5RouteCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, MMPCol) = "UDOT_BMP"
    MMPCol = MMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SMPCol) = "INT_RT_1_M"
    SMPCol = SMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TMPCol) = "INT_RT_2_M"
    TMPCol = TMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4MPCol) = "INT_RT_3_M"
    Q4MPCol = Q4MPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5MPCol) = "INT_RT_4_M"
    Q5MPCol = Q5MPCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by tertiary route/MP, main route/MP, secondary route/MP, quartenary route/MP, quinary route/MP
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TRouteCol) & "2:" & Col_Letter(TRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TMPCol) & "2:" & Col_Letter(TMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MRouteCol) & "2:" & Col_Letter(MRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MMPCol) & "2:" & Col_Letter(MMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SRouteCol) & "2:" & Col_Letter(SRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SMPCol) & "2:" & Col_Letter(SMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4RouteCol) & "2:" & Col_Letter(Q4RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4MPCol) & "2:" & Col_Letter(Q4MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5RouteCol) & "2:" & Col_Letter(Q5RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5MPCol) & "2:" & Col_Letter(Q5MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub

Sub QuartenaryRouteMPSort()

'Declare variables
Dim numRows, numCols As Long
Dim MRouteCol, SRouteCol, TRouteCol, Q4RouteCol, Q5RouteCol, MMPCol, SMPCol, TMPCol, Q4MPCol, Q5MPCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
MRouteCol = 1
SRouteCol = 1
TRouteCol = 1
Q4RouteCol = 1
Q5RouteCol = 1
MMPCol = 1
SMPCol = 1
TMPCol = 1
Q4MPCol = 1
Q5MPCol = 1
datasheet = "Dataset"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, MRouteCol) = "ROUTE"
    MRouteCol = MRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SRouteCol) = "INT_RT_1"
    SRouteCol = SRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TRouteCol) = "INT_RT_2"
    TRouteCol = TRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4RouteCol) = "INT_RT_3"
    Q4RouteCol = Q4RouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5RouteCol) = "INT_RT_4"
    Q5RouteCol = Q5RouteCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, MMPCol) = "UDOT_BMP"
    MMPCol = MMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SMPCol) = "INT_RT_1_M"
    SMPCol = SMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TMPCol) = "INT_RT_2_M"
    TMPCol = TMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4MPCol) = "INT_RT_3_M"
    Q4MPCol = Q4MPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5MPCol) = "INT_RT_4_M"
    Q5MPCol = Q5MPCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by quartenary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quinary route/MP
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4RouteCol) & "2:" & Col_Letter(Q4RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4MPCol) & "2:" & Col_Letter(Q4MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MRouteCol) & "2:" & Col_Letter(MRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MMPCol) & "2:" & Col_Letter(MMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SRouteCol) & "2:" & Col_Letter(SRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SMPCol) & "2:" & Col_Letter(SMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TRouteCol) & "2:" & Col_Letter(TRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TMPCol) & "2:" & Col_Letter(TMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5RouteCol) & "2:" & Col_Letter(Q5RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5MPCol) & "2:" & Col_Letter(Q5MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub

Sub QuinaryRouteMPSort()

'Declare variables
Dim numRows, numCols As Long
Dim MRouteCol, SRouteCol, TRouteCol, Q4RouteCol, Q5RouteCol, MMPCol, SMPCol, TMPCol, Q4MPCol, Q5MPCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
MRouteCol = 1
SRouteCol = 1
TRouteCol = 1
Q4RouteCol = 1
Q5RouteCol = 1
MMPCol = 1
SMPCol = 1
TMPCol = 1
Q4MPCol = 1
Q5MPCol = 1
datasheet = "Dataset"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, MRouteCol) = "ROUTE"
    MRouteCol = MRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SRouteCol) = "INT_RT_1"
    SRouteCol = SRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TRouteCol) = "INT_RT_2"
    TRouteCol = TRouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4RouteCol) = "INT_RT_3"
    Q4RouteCol = Q4RouteCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5RouteCol) = "INT_RT_4"
    Q5RouteCol = Q5RouteCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, MMPCol) = "UDOT_BMP"
    MMPCol = MMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, SMPCol) = "INT_RT_1_M"
    SMPCol = SMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, TMPCol) = "INT_RT_2_M"
    TMPCol = TMPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q4MPCol) = "INT_RT_3_M"
    Q4MPCol = Q4MPCol + 1
Loop

Do Until Sheets(datasheet).Cells(1, Q5MPCol) = "INT_RT_4_M"
    Q5MPCol = Q5MPCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by quinary route/MP, main route/MP, secondary route/MP, tertiary route/MP, quartenary route/MP
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5RouteCol) & "2:" & Col_Letter(Q5RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q5MPCol) & "2:" & Col_Letter(Q5MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MRouteCol) & "2:" & Col_Letter(MRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(MMPCol) & "2:" & Col_Letter(MMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SRouteCol) & "2:" & Col_Letter(SRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(SMPCol) & "2:" & Col_Letter(SMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TRouteCol) & "2:" & Col_Letter(TRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(TMPCol) & "2:" & Col_Letter(TMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4RouteCol) & "2:" & Col_Letter(Q4RouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(Q4MPCol) & "2:" & Col_Letter(Q4MPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal


With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub

Sub RouteIDSort()

'Declare variables
Dim numRows, numCols As Long
Dim MRouteCol, SRouteCol, TRouteCol, MMPCol, SMPCol, TMPCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
MRouteCol = 1
MMPCol = 1

'Find row and column values
Do Until ActiveSheet.Cells(1, MRouteCol) = "ROUTE_ID"    '  Or ActiveSheet.Cells(1, MMPCol) = "MILEPOINT"
    MRouteCol = MRouteCol + 1
Loop

Do Until ActiveSheet.Cells(1, MMPCol) = "BEG_MILEPOINT" Or ActiveSheet.Cells(1, MMPCol) = "MILEPOINT"     '     Or ActiveSheet.Cells(1, MMPCol) = "UDOT_BMP"
    MMPCol = MMPCol + 1
Loop


Do Until ActiveSheet.Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until ActiveSheet.Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by route ID/MP
ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(Col_Letter(MRouteCol) & "2:" & Col_Letter(MRouteCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(Col_Letter(MMPCol) & "2:" & Col_Letter(MMPCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.ActiveSheet.Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub

Sub LatLongElevSortUICPM()

'Declare variables
Dim numRows, numCols As Long
Dim latcol, longcol, ElevCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
latcol = 1
longcol = 1
ElevCol = 1
datasheet = "UICPMinput"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, latcol) = "LATITUDE"
    latcol = latcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, longcol) = "LONGITUDE"
    longcol = longcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, ElevCol) = "ELEVATION"
    ElevCol = ElevCol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by Latitude, then Longitude, then Elevation
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(latcol) & "2:" & Col_Letter(latcol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(longcol) & "2:" & Col_Letter(longcol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(ElevCol) & "2:" & Col_Letter(ElevCol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub

Sub LatLongElevSortCrashUICPM()

'Declare variables
Dim numRows, numCols As Long
Dim latcol, longcol, ElevCol As Long
Dim datasheet As String

'Assign initial variable values
numRows = 1
numCols = 1
latcol = 1
longcol = 1
ElevCol = 1
datasheet = "CrashInput"

'Find row and column values
Do Until Sheets(datasheet).Cells(1, latcol) = "LATITUDE" Or Sheets(datasheet).Cells(1, latcol) = "UTM_Y"
    latcol = latcol + 1
Loop

Do Until Sheets(datasheet).Cells(1, longcol) = "LONGITUDE" Or Sheets(datasheet).Cells(1, longcol) = "UTM_X"
    longcol = longcol + 1
Loop


Do Until Sheets(datasheet).Cells(1, numCols + 1) = ""
    numCols = numCols + 1
Loop

Do Until Sheets(datasheet).Cells(numRows + 1, 1) = ""
    numRows = numRows + 1
Loop


'Sort by Latitude, then Longitude, then Elevation
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(latcol) & "2:" & Col_Letter(latcol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(datasheet).Sort.SortFields.Add Key:=Range(Col_Letter(longcol) & "2:" & Col_Letter(longcol) & numRows), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(datasheet).Sort
    .SetRange Range("A1:" & Col_Letter(numCols) & numRows)
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
End Sub



Public Function Col_Letter(lngCol) As String

Dim vArr

vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)

End Function


Public Sub Find_Replace(find As Variant, repl As Variant, col As Integer, sname As String)
    '
    'Created by Samuel Runyan - 8/10/2021
    'a macro used to find and replace data in a range. Originally created for the cleanNumetric() sub
    '
    '
    'Dim row As Long
    'row = 2
    'Do Until Cells(row, 1) = ""
    '    Sheets(sname).Cells(row, col) = replace(Sheets(sname).Cells(row, col), find, repl)
    '    row = row + 1
    'Loop
    
    Sheets(sname).Columns(col).Replace What:=find, Replacement:=repl, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        
End Sub

Public Sub Sort_Crashes(SheetName As String)

    'Sorts the data by the first column, Crash_ID
    Dim row1, col1, col2 As Long
    col2 = 1
    Do Until Replace(LCase(Sheets(SheetName).Cells(1, col2)), " ", "_") = "crash_id"
        col2 = col2 + 1
    Loop
    row1 = Sheets(SheetName).Cells(1, col2).End(xlDown).row
    col1 = Sheets(SheetName).Range("A1").End(xlToRight).Column
  
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=ActiveWorkbook.Worksheets(SheetName).Range(Cells(2, col2), Cells(row1, col2)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange ActiveWorkbook.Worksheets(SheetName).Range(Cells(1, 1), Cells(row1, col1))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub

Public Sub Sort_Vehicles(SheetName As String)

    'Sorts the data by the first column, Crash_ID, and the column with the Vehicle_Detail_ID
    Dim row1, col1, col2, col3 As Long
    col2 = 1
    col3 = 1
    Do Until Replace(LCase(Sheets(SheetName).Cells(1, col2)), " ", "_") = "crash_id"
        col2 = col2 + 1
    Loop
    Do Until Replace(LCase(Sheets(SheetName).Cells(1, col3)), " ", "_") = "vehicle_detail_id" Or Replace(LCase(Sheets(SheetName).Cells(1, col3)), " ", "_") = "vehicle_type"
        col3 = col3 + 1
    Loop
    row1 = Sheets(SheetName).Cells(1, col2).End(xlDown).row
    col1 = Sheets(SheetName).Range("A1").End(xlToRight).Column

    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=ActiveWorkbook.Worksheets(SheetName).Range(Cells(2, col2), Cells(row1, col2)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=ActiveWorkbook.Worksheets(SheetName).Range(Cells(2, col3), Cells(row1, col3)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange ActiveWorkbook.Worksheets(SheetName).Range(Cells(1, 1), Cells(row1, col1))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub

Sub autofilterLocation(rampcol As Integer, crashroutecol As Integer, crashMPcol As Integer, Lnumrow As Long, Lnumcol As Integer, combo As String)
'
'
'

'
    Dim row As Long
    
    'Activate worksheet
    ActiveWorkbook.Worksheets(combo).Activate
    
    'Sort Fields
    'SortBlankOnTop Range("A2", Cells(Lnumrow, Lnumcol))
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Cells(1, rampcol), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range("A2", Cells(Lnumrow, Lnumcol))
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Autofilter Method
    Worksheets(combo).Range(Cells(1, 1), Cells(Lnumrow, Lnumcol)).AutoFilter
    Worksheets(combo).Range(Cells(1, 1), Cells(Lnumrow, Lnumcol)).AutoFilter Field:=crashroutecol, Criteria1:="0085"
    Worksheets(combo).Range(Cells(1, 1), Cells(Lnumrow, Lnumcol)).AutoFilter Field:=crashMPcol, Criteria1:= _
        "<2.977", Operator:=xlAnd
    
    row = 2
    Do Until Cells(row, crashroutecol) = ""
        If Cells(row, crashroutecol) = "0085" Then
            Cells(row, crashroutecol) = "194"
        End If
        row = row + 1
    Loop

    Range(Cells(1, 1), Cells(1, Lnumcol)).AutoFilter
    
    Worksheets(combo).Range(Cells(1, 1), Cells(Lnumrow, Lnumcol)).AutoFilter
    Worksheets(combo).Range(Cells(1, 1), Cells(Lnumrow, Lnumcol)).AutoFilter Field:=rampcol, Criteria1:="<>NA", Operator:=xlAnd, Criteria2:="<>0"      ', Operator:=xlAnd, Criteria3:="<>" FLAGGED: this would help if it's blank but cant do it.
    Range(Range(Range("A2"), Range("A2").End(xlDown)), Range(Range("A2"), Range("A2").End(xlDown)).End(xlToRight)).EntireRow.Delete
    Range(Cells(1, 1), Cells(1, Lnumcol)).AutoFilter

End Sub

Sub filterRollups(combo As String, Lnumcol As Integer, Lnumrow As Long, Rnumcol As Integer, Rnumrow As Long)
'
'
'

'
    Dim locID, rollID As Long
    Dim IDcol, countme As Integer
    Dim rollrow, locrow As Long
    
    'Activate worksheet
    ActiveWorkbook.Worksheets(combo).Activate
    
    'Find crash id columns
    IDcol = 1
    Do While Sheets(combo).Cells(1, IDcol) <> "CRASH_ID"
        IDcol = IDcol + 1
    Loop
    locID = IDcol
    'locID = Split(Cells(1, IDcol).Address, "$")(1)
    
    IDcol = Lnumcol + 1
    Do While Sheets(combo).Cells(1, IDcol) <> "CRASH_ID"
        IDcol = IDcol + 1
    Loop
    rollID = IDcol
    'rollID = Split(Cells(1, IDcol).Address, "$")(1)
    
    
    locrow = 2
    rollrow = 2
    Do Until Cells(locrow, locID) = ""
        'if the two crash id's match then do nothing
        If Cells(locrow, locID) = Cells(rollrow, rollID) Then
            locrow = locrow + 1
            rollrow = rollrow + 1
        'if the rollup crash id does not match location crash id, then clear that row and move on. This is common since the combined crash sheet got pared for just applicable crashes.
        ElseIf Cells(locrow, locID) > Cells(rollrow, rollID) Then
            Range(Cells(rollrow, Lnumcol + 1), Cells(rollrow, Lnumcol + Rnumcol)).Clear
            rollrow = rollrow + 1
        'if the location crash id does not have a match in the rollup sheet, then clear that crash. I don't expect this number to be high since there shouldn't be any missing data
        ElseIf Cells(locrow, locID) < Cells(rollrow, rollID) Then
            Range(Cells(locrow, 1), Cells(locrow, Lnumcol)).Clear
            countme = countme + 1
            locrow = locrow + 1
        End If
    Loop
    

    'Conditional Formatting
    '   1) compare both crash ID columns
    '   2) clears unique cells
    'Cells(1, Lnumcol + 1).Activate
    'Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions.AddUniqueValues
    'Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions(Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions.Count).SetFirstPriority
    'Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions(1).DupeUnique = xlUnique
    'Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions(1).Font.Color = vbRed
    'Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions(1).StopIfTrue = False
    
    'autofilter and clear red rows
    'Range(Cells(1, 1), Cells(Rnumrow, Lnumcol + Rnumcol)).AutoFilter
    'Range(Cells(1, 1), Cells(Rnumrow, Lnumcol + Rnumcol)).AutoFilter Field:=IDcol, Criteria1:=vbRed, Operator:=xlFilterFontColor
    'Range(Range(Cells(2, Lnumcol + 1), Cells(2, Lnumcol + 1).End(xlDown)), Range(Cells(2, Lnumcol + 1), Cells(2, Lnumcol + 1).End(xlDown)).End(xlToRight)).Clear
    'Range(Cells(1, Lnumcol + 1), Cells(1, Lnumcol + Rnumcol)).AutoFilter
    
    'Clear conditional formatting
    'Range(locID & ":" & locID & "," & rollID & ":" & rollID).FormatConditions.Delete
    
    'sort blank rows to the bottom
    'sort rollups
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add2 Key:=Range(Cells(1, rollID), Cells(Rnumrow, rollID)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range(Cells(1, Lnumcol + 1), Cells(Rnumrow, Lnumcol + Rnumcol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'sort location
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add2 Key:=Range(Cells(1, locID), Cells(Lnumrow, locID)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range(Cells(1, 1), Cells(Lnumrow, Lnumcol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Clear unwanted rows
    Range(Cells(Lnumrow + 1, rollID), Cells(Rnumrow, Lnumcol + Rnumcol)).Clear
    
    'Delete extra crashID column
    Columns(IDcol).Delete
    
    
End Sub

Sub SortBlankOnTop(WorkRng As Range)
'Update 20140318
On Error Resume Next
Dim xMin As Double
'xTitleId = "KutoolsforExcel"
'Set WorkRng = Application.selection
'Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
xMin = Application.WorksheetFunction.Small(WorkRng, 1) - 1
WorkRng.SpecialCells(xlCellTypeBlanks) = xMin
WorkRng.Sort , Key1:=Cells(1, rampcol), Order1:=xlAscending, header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
WorkRng.Replace What:=xMin, Replacement:="", LookAt:=xlWhole
End Sub

