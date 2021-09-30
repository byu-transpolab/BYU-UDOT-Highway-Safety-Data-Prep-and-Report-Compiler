Attribute VB_Name = "DP_12_CAMSCrash"
Option Explicit


Sub MP_HitList()

Dim ifile As String
Dim Rcol0 As Integer, Rcol1 As Integer, Rcol2 As Integer, Rcol3 As Integer, Rcol4 As Integer
Dim Route0, Route1 As Variant, Route2 As Variant, Route3 As Variant, Route4 As Variant, nextR0 As Variant
Dim TraffCoCol As Integer, SRtoSRcol As Integer
Dim TraffCo As String, SRSRYN As String
Dim MPcol As Integer
Dim MPval As Double
Dim rowcount As Integer, HLrow As Integer
Dim JotItDown As Boolean
Dim FAdist As Double
Dim SRYN As String, FAYN As String, SigYN As String


ifile = "Intersections"

'find column numbers

Rcol0 = 1
Do While Sheets(ifile).Cells(1, Rcol0) <> "ROUTE"
    Rcol0 = Rcol0 + 1
Loop

Rcol1 = 1
Do While Sheets(ifile).Cells(1, Rcol1) <> "INT_RT_1"
    Rcol1 = Rcol1 + 1
Loop

Rcol2 = 1
Do While Sheets(ifile).Cells(1, Rcol2) <> "INT_RT_2"
    Rcol2 = Rcol2 + 1
Loop

Rcol3 = 1
Do While Sheets(ifile).Cells(1, Rcol3) <> "INT_RT_3"
    Rcol3 = Rcol3 + 1
Loop

Rcol4 = 1
Do While Sheets(ifile).Cells(1, Rcol4) <> "INT_RT_4"
    Rcol4 = Rcol4 + 1
Loop

TraffCoCol = 1
Do While Sheets(ifile).Cells(1, TraffCoCol) <> "TRAFFIC_CO"
    TraffCoCol = TraffCoCol + 1
Loop

SRtoSRcol = 1
Do While Sheets(ifile).Cells(1, SRtoSRcol) <> "SR_SR"
    SRtoSRcol = SRtoSRcol + 1
Loop

MPcol = 1
Do While Sheets(ifile).Cells(1, MPcol) <> "UDOT_BMP"
    MPcol = MPcol + 1
Loop

'create a new sheet; this is where the list of route, mp, and distances will be saved
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.count)
ActiveWorkbook.ActiveSheet.Name = "HitList"
Sheets("HitList").Cells(1, 1) = "Route"
Sheets("HitList").Cells(1, 2) = "Intersection Center MP"
Sheets("HitList").Cells(1, 3) = "Distance (Radius, in ft)"

'find the intersections to go on the hit list and fill in the information
rowcount = 2
HLrow = 2
SRYN = Sheets("Inputs").Cells(11, 13)
FAYN = Sheets("Inputs").Cells(12, 13)
SigYN = Sheets("Inputs").Cells(13, 13)

Do Until Sheets(ifile).Cells(rowcount, Rcol0) = ""
    Route0 = Sheets(ifile).Cells(rowcount, Rcol0)
    Route1 = Sheets(ifile).Cells(rowcount, Rcol1)
    Route2 = Sheets(ifile).Cells(rowcount, Rcol2)
    Route3 = Sheets(ifile).Cells(rowcount, Rcol3)
    Route4 = Sheets(ifile).Cells(rowcount, Rcol4)
    TraffCo = Sheets(ifile).Cells(rowcount, TraffCoCol)
    SRSRYN = Sheets(ifile).Cells(rowcount, SRtoSRcol)
    JotItDown = False
    
    'if all three have been checked
    If SRYN = "YES" And FAYN = "YES" And SigYN = "YES" Then
        If TraffCo = "SIGNAL" Then   'signalized intersections
            JotItDown = True
        ElseIf SRSRYN = "YES" Then   'SR to SR intersections
            JotItDown = True
        ElseIf Route1 <> "" And Route1 <> "Local" Then   'FedAid intersections
            If Route1 > 491 Then JotItDown = True
        ElseIf Route2 <> "" And Route2 <> "Local" Then
            If Route2 > 491 Then JotItDown = True
        ElseIf Route3 <> "" And Route3 <> "Local" Then
            If Route3 > 491 Then JotItDown = True
        ElseIf Route4 <> "" And Route4 <> "Local" Then
            If Route4 > 491 Then JotItDown = True
        End If
    'If two (SR and FA) have been checked
    ElseIf SRYN = "YES" And FAYN = "YES" And SigYN = "" Then
        If SRSRYN = "YES" Then   'SR to SR intersections
            JotItDown = True
        ElseIf Route1 <> "" And Route1 <> "Local" Then   'FedAid intersections
            If Route1 > 491 Then JotItDown = True
        ElseIf Route2 <> "" And Route2 <> "Local" Then
            If Route2 > 491 Then JotItDown = True
        ElseIf Route3 <> "" And Route3 <> "Local" Then
            If Route3 > 491 Then JotItDown = True
        ElseIf Route4 <> "" And Route4 <> "Local" Then
            If Route4 > 491 Then JotItDown = True
        End If
    'If two (SR and Sig) have been checked
    ElseIf SRYN = "YES" And FAYN = "" And SigYN = "YES" Then
        If TraffCo = "SIGNAL" Then   'signalized intersections
            JotItDown = True
        ElseIf SRSRYN = "YES" Then   'SR to SR intersections
            JotItDown = True
        End If
    'If two (FA and Sig) have been checked
    ElseIf SRYN = "" And FAYN = "YES" And SigYN = "YES" Then
        If TraffCo = "SIGNAL" Then   'signalized intersections
            JotItDown = True
        ElseIf Route1 <> "" And Route1 <> "Local" Then   'FedAid intersections
            If Route1 > 491 Then JotItDown = True
        ElseIf Route2 <> "" And Route2 <> "Local" Then
            If Route2 > 491 Then JotItDown = True
        ElseIf Route3 <> "" And Route3 <> "Local" Then
            If Route3 > 491 Then JotItDown = True
        ElseIf Route4 <> "" And Route4 <> "Local" Then
            If Route4 > 491 Then JotItDown = True
        End If
    'If one (SR) has been checked
    ElseIf SRYN = "YES" And FAYN = "" And SigYN = "" Then
        If SRSRYN = "YES" Then   'SR to SR intersections
            JotItDown = True
        End If
    'If one (FA) has been checked
    ElseIf SRYN = "" And FAYN = "YES" And SigYN = "" Then
        If Route1 <> "" And Route1 <> "Local" Then   'FedAid intersections
            If Route1 > 491 Then JotItDown = True
        ElseIf Route2 <> "" And Route2 <> "Local" Then
            If Route2 > 491 Then JotItDown = True
        ElseIf Route3 <> "" And Route3 <> "Local" Then
            If Route3 > 491 Then JotItDown = True
        ElseIf Route4 <> "" And Route4 <> "Local" Then
            If Route4 > 491 Then JotItDown = True
        End If
    'If one (Sig) has been checked
    ElseIf SRYN = "" And FAYN = "" And SigYN = "YES" Then
        If TraffCo = "SIGNAL" Then   'signalized intersections
            JotItDown = True
        End If
    'End of that nasty If statement
    End If

    If JotItDown = True Then
        Sheets("HitList").Cells(HLrow, 1) = Route0
        MPval = Sheets(ifile).Cells(rowcount, MPcol)
        Sheets("HitList").Cells(HLrow, 2) = MPval
        HLrow = HLrow + 1
    End If

rowcount = rowcount + 1
    
Loop

'call the FAdistCAMS sub to calculate the values for column 3 in the hitlist
FAdistCAMS

'deletes the intersections sheet now that's it's not needed
Application.DisplayAlerts = False
If SheetExists(ifile) Then
    Sheets(ifile).Delete
End If
Application.DisplayAlerts = True



End Sub
Sub FAdistCAMS()

'calculates the intersections' functional area distance according to given parameters

Dim FAdefby As String, FAmeasfrom As String
Dim part1 As Single, part2 As Single
Dim FAdist As Single
Dim Mroute As Integer
Dim MP As Single
Dim myHLrow As Long
Dim SLrow As Long
Dim SLroute As Integer, SLMP As Integer, SLdir As Integer, S_L As Integer

FAdefby = Sheets("Inputs").Cells(14, 13)
If FAdefby = "Speed Limit" Then
    
    Worksheets("Speed_Limit").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Speed_Limit").Sort.SortFields.Add Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Speed_Limit").Sort
    .SetRange Sheets("Speed_Limit").UsedRange
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With
    
    SLroute = 1
    Do While Sheets("Speed_Limit").Cells(1, SLroute) <> "ROUTE_ID"
        SLroute = SLroute + 1
    Loop
    SLMP = 1
    Do While Sheets("Speed_Limit").Cells(1, SLMP) <> "BEG_MILEPOINT"
        SLMP = SLMP + 1
    Loop
    SLdir = 1
    Do While Sheets("Speed_Limit").Cells(1, SLdir) <> "DIRECTION"
        SLdir = SLdir + 1
    Loop
    S_L = 1
    Do While Sheets("Speed_Limit").Cells(1, S_L) <> "SPEED_LIMIT"
        S_L = S_L + 1
    Loop
End If

FAmeasfrom = Sheets("Inputs").Cells(15, 13)
If FAmeasfrom = "Center" Then
    part1 = 0
ElseIf FAmeasfrom = "Stopbar" Then
    part1 = 60
    'this is where pavement messages needs to be used/called
    'perhaps use it in a way where you pass the variables of route and MP to it
    'something like CAMS_PMcalc(Mroute,MP)
    'or this is where you put the assumptions for intersection physical area
End If

myHLrow = 2
Do While Sheets("HitList").Cells(myHLrow, 1) <> ""
    
    Mroute = Sheets("HitList").Cells(myHLrow, 1)
    MP = Sheets("HitList").Cells(myHLrow, 2)
    If FAdefby = "Speed Limit" Then
        part2 = CAMS_FAcalc(Mroute, MP, SLroute, SLMP, SLdir, S_L)
        'this is where speed limit needs to be used/called
        'perhaps use it in a way where you pass the variables of route and MP to it
        'something like CAMS_SLcalc(Mroute, MP)
        'or maybe use an array
    'ElseIf FAdefby = "Functional Class" Then
        'this is where functional class needs to be used/called
        'perhaps use it in a way where you pass the variables of route and MP to it
        'something like CAMS_FCcalc(Mroute,MP)
    Else
        part2 = 250
    End If

    FAdist = part1 + part2
    
    Sheets("HitList").Cells(myHLrow, 3) = FAdist

myHLrow = myHLrow + 1

'reset values for each loop:
FAdist = 0
part2 = 0

Loop

'delete the Speed Limit tab since it is no longer needed
Application.DisplayAlerts = False
If SheetExists("Speed_Limit") Then
    Sheets("Speed_Limit").Delete
End If
Application.DisplayAlerts = True


End Sub


Function CAMS_FAcalc(Mroute As Integer, MP As Single, SLroute As Integer, SLMP As Integer, SLdir As Integer, S_L As Integer) As Integer

Dim keyrow As Integer
Dim mySLrow As Long
mySLrow = 2

If Mroute = 15 Or Mroute = 70 Or Mroute = 80 Or Mroute = 84 Or Mroute = 215 Then
    'message to Camille
    MsgBox "There are interstate intersections"
End If

Do
    If Int(Left(Sheets("Speed_Limit").Cells(mySLrow, SLroute), 4)) < Mroute Then
        'move on to next row
        mySLrow = mySLrow + 1
    ElseIf Int(Left(Sheets("Speed_Limit").Cells(mySLrow, SLroute), 4)) = Mroute Then
        If Sheets("Speed_Limit").Cells(mySLrow, SLMP) <= MP And Sheets("Speed_Limit").Cells(mySLrow, SLroute + 1) >= MP Then
            'get the value
            keyrow = 3
            Do While Sheets("Key").Cells(keyrow, 59) <> Sheets("Speed_Limit").Cells(mySLrow, S_L)
                keyrow = keyrow + 1
            Loop
            CAMS_FAcalc = Sheets("Key").Cells(keyrow, 63)
            Exit Do
        Else
        'move on to next row
        mySLrow = mySLrow + 1
        End If
    Else
        'error message
        CAMS_FAcalc = 250
        Exit Do
    End If
Loop



End Function


Sub CleanIntersections()

Dim colRoute0 As Integer, colRoute1 As Integer, colRoute2 As Integer, colRoute3 As Integer, colRoute4 As Integer
Dim myrow As Integer
Dim groupcol As Integer
Dim deletedrow As Boolean

'find column numbers

colRoute0 = 1
Do While Sheets("Intersections").Cells(1, colRoute0) <> "ROUTE"
    colRoute0 = colRoute0 + 1
Loop

colRoute1 = 1
Do While Sheets("Intersections").Cells(1, colRoute1) <> "INT_RT_1"
    colRoute1 = colRoute1 + 1
Loop

colRoute2 = 1
Do While Sheets("Intersections").Cells(1, colRoute2) <> "INT_RT_2"
    colRoute2 = colRoute2 + 1
Loop

colRoute3 = 1
Do While Sheets("Intersections").Cells(1, colRoute3) <> "INT_RT_3"
    colRoute3 = colRoute3 + 1
Loop

colRoute4 = 1
Do While Sheets("Intersections").Cells(1, colRoute4) <> "INT_RT_4"
    colRoute4 = colRoute4 + 1
Loop

groupcol = 1
Do While Sheets("Intersections").Cells(1, groupcol) <> "Group_Num"
    groupcol = groupcol + 1
Loop

'loop through and fix it all
myrow = 2
Sheets("Intersections").Activate
Application.ScreenUpdating = False
deletedrow = False
Do Until ActiveWorkbook.Sheets("Intersections").Cells(myrow, 1) = ""
    If deletedrow = False Then
        'corrects typos
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) = "LOCAL" Then
            ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) = "Local"
        End If
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) = "LOCAL" Then
            ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) = "Local"
        End If
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) = "LOCAL" Then
            ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) = "Local"
        End If
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) = "local" Or ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) = "LOCAL" Then
            ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) = "Local"
        End If
        'removes the direction indicator from the main route column and stores the route as a number
        Cells(myrow, colRoute0) = Left(Cells(myrow, colRoute0), 4)
        Cells(myrow, colRoute0).NumberFormat = "0000"
        Cells(myrow, colRoute0).Value = Cells(myrow, colRoute0).Value
        'makes the route numbers in number format (instead of text format)
        If Cells(myrow, colRoute0) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Cells(myrow, colRoute0) = "0011"          'Change route to 0011 for the purpose of this process.
        End If
        Cells(myrow, colRoute0).NumberFormat = "0000"
        Cells(myrow, colRoute0).Value = Cells(myrow, colRoute0).Value
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) <> "" And _
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute1) <> "Local" Then
            If Cells(myrow, colRoute1) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
                Cells(myrow, colRoute1) = "0011"          'Change route to 0011 for the purpose of this process.
            End If
            Cells(myrow, colRoute1).NumberFormat = "0000"
            Cells(myrow, colRoute1).Value = Cells(myrow, colRoute1).Value
        End If
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) <> "" And _
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute2) <> "Local" Then
            If Cells(myrow, colRoute2) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
            Cells(myrow, colRoute2) = "0011"          'Change route to 0011 for the purpose of this process.
            End If
            Cells(myrow, colRoute2).NumberFormat = "0000"
            Cells(myrow, colRoute2).Value = Cells(myrow, colRoute2).Value
        End If
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) <> "" And _
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute3) <> "Local" Then
            If Cells(myrow, colRoute3) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
                Cells(myrow, colRoute3) = "0011"          'Change route to 0011 for the purpose of this process.
            End If
            Cells(myrow, colRoute3).NumberFormat = "0000"
            Cells(myrow, colRoute3).Value = Cells(myrow, colRoute3).Value
        End If
        If ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) <> "" And _
        ActiveWorkbook.Sheets("Intersections").Cells(myrow, colRoute4) <> "Local" Then
            If Cells(myrow, colRoute4) = "089A" Then      'Corrects strange route ID. 089A previously called SR-11.
                Cells(myrow, colRoute4) = "0011"          'Change route to 0011 for the purpose of this process.
            End If
            Cells(myrow, colRoute4).NumberFormat = "0000"
            Cells(myrow, colRoute4).Value = Cells(myrow, colRoute4).Value
        End If
    End If
    deletedrow = False
    'delete the ramp rows
    If Len(Cells(myrow + 1, colRoute0)) > 5 Or Cells(myrow + 1, groupcol) > 1 Then
        Sheets("Intersections").Cells(myrow + 1, colRoute0).EntireRow.Delete
        myrow = myrow - 1
        deletedrow = True
    End If
myrow = myrow + 1
Loop
Application.ScreenUpdating = True
Application.ScreenUpdating = False

End Sub

Sub getCAMScrashes()

''created by Camille on July 3, 2019
''optimized by Sam - 2021

'gets rids of any crash not on a state route from the crash location file
'adds the crash rollup data to the crash location data
'deletes crashes that fall within the hit list

Dim rowcounter As Long
Dim crashroutecol As Integer, rampcol As Integer, crashMPcol As Integer, crashintrelcol As Integer
Dim crashonroute As Variant
Dim deletelong As Boolean, deletefedaid As Boolean
Dim Lnumcol As Integer, Lnumrow As Long, Rnumcol As Integer, Rnumrow As Long, Tnumcol As Integer, Tnumrow As Long
Dim datarow As Long
Dim i As Integer
Dim combo As String
Dim HLrow As Long
Dim ICC As Long
Dim SRdone As Boolean
Dim countme As Long

Application.ScreenUpdating = False
Sheets("Location").Activate

crashroutecol = 1
Do While Sheets("Location").Cells(1, crashroutecol) <> "ROUTE_ID"
    crashroutecol = crashroutecol + 1
Loop

rampcol = 1
Do While Sheets("Location").Cells(1, rampcol) <> "RAMP_ID"
    rampcol = rampcol + 1
Loop

crashMPcol = 1
Do While Sheets("Location").Cells(1, crashMPcol) <> "MILEPOINT"
    crashMPcol = crashMPcol + 1
Loop

Lnumcol = Sheets("Location").Range("A1").End(xlToRight).Column
Lnumrow = Sheets("Location").Range("A1").End(xlDown).row

'alphabetically sorts the Location file based on route
ActiveWorkbook.Worksheets("Location").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Location").Sort.SortFields.Add Key:=Range(Cells(2, crashroutecol), Cells(Lnumrow, crashroutecol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("Location").Sort
    .SetRange Range(Cells(1, 1), Cells(Lnumrow, Lnumcol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'adds new sheet
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.count)
ActiveWorkbook.ActiveSheet.Name = "Combined_Crash"
combo = "Combined_Crash"



'pare down the Locations file ''''''''''USING THE SEND TO OTHER SHEET AND FILTER METHOD'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'copy and paste all the state route rows from "Location" to "Combined_Crash"
            'loop through the "Combined_Crash" rows
            'convert route ID to number format (including replacing 089A with 0011)
            'fix UT-85 MP < 2.977 problem
            'if there is a ramp ID (AKA the ramp ID is not blank), delete the row

'Find last row of state routes in the "Location" sheet
rowcounter = 2
Do While Sheets("Location").Cells(rowcounter, crashroutecol).Value < 492
    rowcounter = rowcounter + 1
Loop

'Copy state route rows from "Location" over to "Combined_Crash"
Worksheets("Location").Cells(1, 1).Resize(rowcounter, Lnumcol).Copy Destination:=Worksheets(combo).Range("A1")

'Finish paring location data
autofilterLocation rampcol, crashroutecol, crashMPcol, Lnumrow, Lnumcol, combo

'THIS METHOD HAS BEEN COMMENTED OUT AND UPDATED TO THE SEND TO OTHER SHEET AND FILTER METHOD''''''''''''''''''''''''''''''''''''''''''''''''''''
'pare down the Locations file ''''''''''USING THE SORTING AND SEND TO OTHER SHEET METHOD''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'do until the next row is "", start the counter at row 1
            'if the second to next row has a route lenth longer than 4 characters, delete it
            'convert the 4 digit text in the next row to a number (and change 089A to 0011)
            'if the 4 digit number is > 491, delete it
            'if there is a ramp ID (AKA the ramp ID is not blank), delete it
'THIS METHOD HAS BEEN COMMENTED OUT AND UPDATED TO THE SEND TO OTHER SHEET AND FILTER METHOD''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'pastes column headings
'For i = 1 To Lnumcol
'    Sheets(combo).Cells(1, i) = Sheets("Location").Cells(1, i)
'Next i
'
'rowcounter = 2
'datarow = 2
'SRdone = False
'Do Until SRdone = True
'    'fix UT-85 MP < 2.977 problem (replace 85 with 194)
'    If Sheets("Location").Cells(rowcounter, crashroutecol) = 85 And Sheets("Location").Cells(rowcounter, 6) < 2.977 Then
'        rowcounter = rowcounter
'        Sheets("Location").Cells(rowcounter, crashroutecol) = 194
'    End If
'
'    'change the text to a number
'    Sheets("Location").Cells(rowcounter, crashroutecol).NumberFormat = "0000"
'
'    'if the route number is a state route (AKA less than 492) AND if the ramp id is blank, then copy the information over to the crash combo sheet
'    If Sheets("Location").Cells(rowcounter, crashroutecol).Value < 492 Then
'        If Sheets("Location").Cells(rowcounter, rampcol) = "" Then                              'added by Camille on 08/13/2019
'            For i = 1 To Lnumcol
'                Sheets(combo).Cells(datarow, i) = Sheets("Location").Cells(rowcounter, i)
'            Next i
'            datarow = datarow + 1
'        End If
'    Else
'        SRdone = True
'    End If
'    rowcounter = rowcounter + 1
'Loop

datarow = Sheets(combo).Range("A1").End(xlDown).row + 1
'skip down to the 089A rows
Do Until Sheets("Location").Cells(rowcounter, crashroutecol) = "089A"
    If Sheets("Location").Cells(rowcounter, crashroutecol) = "" Then Exit Do
    rowcounter = rowcounter + 1
Loop
Do Until Sheets("Location").Cells(rowcounter, crashroutecol) = ""
    'change 089A to 0011
    If Sheets("Location").Cells(rowcounter, crashroutecol) = "089A" Then
        Sheets("Location").Cells(rowcounter, crashroutecol) = "0011"
        If Sheets("Location").Cells(rowcounter, rampcol) = "" Then
            For i = 1 To Lnumcol
                Sheets(combo).Cells(datarow, i) = Sheets("Location").Cells(rowcounter, i)
            Next i
            datarow = datarow + 1
        End If
    End If
    rowcounter = rowcounter + 1
Loop

'deletes the location sheet now that's it's not needed (and it's a huge file that probably bogs down excel)
Application.DisplayAlerts = False
If SheetExists("Location") Then
    Sheets("Location").Delete
End If
Application.DisplayAlerts = True

'matches Rollup information to the Locations file ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'sorts both sheets by crash id
            'takes a location crash id (on the combo sheet), goes down the rollup data until the same crash id is found
            'pastes the matching rollup data on that row in the combo sheet
            
'get new rowcount for location data (which is now on the combo sheet)
Tnumrow = Sheets(combo).Range("A1").End(xlDown).row
'get row count for rollup data
Rnumrow = Sheets("Rollup").Range("A1").End(xlDown).row
'get column count for rollup data
Rnumcol = Sheets("Rollup").Range("A1").End(xlToRight).Column

'sort combo data by crash id (always the first column)
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(Tnumrow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With ActiveWorkbook.Worksheets(combo).Sort
    .SetRange Range(Cells(1, 1), Cells(Tnumrow, Lnumcol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'sort rollup data by crash id (always the first column)
ActiveWorkbook.Worksheets("Rollup").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Rollup").Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(Rnumrow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("Rollup").Sort
    .SetRange Range(Cells(1, 1), Cells(Rnumrow, Rnumcol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With



'''''''''''''''Paste Rollup Data using Sort/Autofilter Methods'''''''''''''''''''''''''''''''''''''''''''''''''

'Copy rows from "Rollup" over to "Combined_Crash"
Worksheets("Rollup").Cells(1, 1).Resize(Rnumrow, Rnumcol).Copy Destination:=Worksheets(combo).Cells(1, Lnumcol + 1)

'Use conditional formatting and sorting to eliminate unwanted rollup rows
filterRollups combo, Lnumcol, Tnumrow, Rnumcol, Rnumrow



'pastes the rollup headers onto the combo sheet
'For i = 2 To Rnumcol
'    Sheets(combo).Cells(1, i + Lnumcol - 1) = Sheets("Rollup").Cells(1, i)
'Next i
'begin looping down the combo sheet
'rowcounter = 2   'counter for the rollup sheet
'datarow = 2      'counter for the combo sheet
'Do While Sheets(combo).Cells(datarow, 1) <> ""
'    'if the two crash id's match then paste the rollup info onto the combined data sheet
'    If Sheets(combo).Cells(datarow, 1) = Sheets("Rollup").Cells(rowcounter, 1) Then
'        For i = 2 To Rnumcol
'            Sheets(combo).Cells(datarow, i + Lnumcol - 1) = Sheets("Rollup").Cells(rowcounter, i)
'        Next i
'        rowcounter = rowcounter + 1
'        datarow = datarow + 1
'    'if the rollup crash id does not have a match in the combined crash sheet, then skip that row and move on. This is common since the combined crash sheet got pared for just applicable crashes.
'    ElseIf Sheets(combo).Cells(datarow, 1) > Sheets("Rollup").Cells(rowcounter, 1) Then
'        rowcounter = rowcounter + 1
'    'if the combined crash id does not have a match in the rollup sheet, then delete that crash. I don't expect this number to be high since there shouldn't be any missing data
'    ElseIf Sheets(combo).Cells(datarow, 1) < Sheets("Rollup").Cells(rowcounter, 1) Then
'        Sheets(combo).Cells(datarow, 1).EntireRow.Delete
'        ''MsgBox "row deleted", vbOKOnly, "FYI"
'        'test intead:
'        countme = countme + 1
'    End If
'Loop

'deletes the rollups sheet now that's it's not needed (and it's a huge file that probably bogs down excel)
Application.DisplayAlerts = False
If SheetExists("Rollup") Then
    Sheets("Rollup").Delete
End If
Application.DisplayAlerts = True


'CRUCIAL STEP FOR THE CAMS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''compares the crash combo information to the hitlist and deletes rows that the hitlist says
            'sorts the hitlist by route and then milepost
            'sorts the combo sheet by route and then milepost
            'loops down hitlist, stops at crashes if they fit within the MP boundaries
            'if the crash is intersection-related ("Y") then delete it from the combo sheet

'get row count for HitList
HLrow = Sheets("HitList").Range("A1").End(xlDown).row

'sort hitlist by route and then milepost
ActiveWorkbook.Worksheets("HitList").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("HitList").Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(HLrow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("HitList").Sort.SortFields.Add Key:=Range(Cells(2, 2), Cells(HLrow, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("HitList").Sort
    .SetRange Range(Cells(1, 1), Cells(HLrow, 3))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
'get column number for crash MP column
crashMPcol = 1
Do While Sheets(combo).Cells(1, crashMPcol) <> "MILEPOINT"
    crashMPcol = crashMPcol + 1
Loop
'get column count for the combo sheet
Tnumcol = Sheets(combo).Range("A1").End(xlToRight).Column

'sort combo sheet by route and then milepost
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, crashroutecol), Cells(Tnumrow, crashroutecol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, crashMPcol), Cells(Tnumrow, crashMPcol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(combo).Sort
    .SetRange Range(Cells(1, 1), Cells(Tnumrow, Tnumcol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
'get column number for crash intersection-related Y/N column
crashintrelcol = 1
Do While Sheets(combo).Cells(1, crashintrelcol) <> "INTERSECTION_RELATED"
    crashintrelcol = crashintrelcol + 1
Loop

''''''''''''''''''''''''''''''''''''added by Sam Runyan 2021
'we will perform arithmetic before looping to save computation
Sheets("Hitlist").Activate
Range("D2").FormulaR1C1 = "=RC[-2]-RC[-1]/5280"                         'lower bound
Range("D2").AutoFill Destination:=Range(Cells(2, 4), Cells(HLrow, 4))
Range("E2").FormulaR1C1 = "=RC[-3]+RC[-2]/5280"                         'upper bound
Range("E2").AutoFill Destination:=Range(Cells(2, 5), Cells(HLrow, 5))



'loop down the hitlist
HLrow = 2
datarow = 2
'datarow = 158095
Do While Sheets("HitList").Cells(HLrow, 1) <> ""
    
    'testing by Camille for Bangerter Hwy
    'If Sheets(combo).Cells(datarow, crashroutecol) = 155 Then
    '    datarow = datarow 'meaningless line of code so I can place a breakpoint
    'End If
    'end of testing
    
    'If Sheets("HitList").Cells(HLrow, 1) = Sheets(combo).Cells(datarow, crashroutecol) And _
    (Sheets("HitList").Cells(HLrow, 2) - (Sheets("HitList").Cells(HLrow, 3) / 5280)) < Sheets(combo).Cells(datarow, crashMPcol) And _
    (Sheets("HitList").Cells(HLrow, 2) + (Sheets("HitList").Cells(HLrow, 3) / 5280)) > Sheets(combo).Cells(datarow, crashMPcol) Then 'route matches and MP fits
    If Sheets("HitList").Cells(HLrow, 1) = Sheets(combo).Cells(datarow, crashroutecol) And _
    Sheets("HitList").Cells(HLrow, 4) < Sheets(combo).Cells(datarow, crashMPcol) And _
    Sheets("HitList").Cells(HLrow, 5) > Sheets(combo).Cells(datarow, crashMPcol) Then 'route matches and MP fits
        If Sheets(combo).Cells(datarow, crashintrelcol) = "Y" Then  'is intersection-related
            Sheets(combo).Cells(datarow, 1).EntireRow.Clear   'clear the row
            ICC = ICC + 1
        End If
        datarow = datarow + 1
    ElseIf Sheets("HitList").Cells(HLrow, 1) > Sheets(combo).Cells(datarow, crashroutecol) Then  'crash route lower than hitlist route
        datarow = datarow + 1  'go to next crash row
    ElseIf Sheets("HitList").Cells(HLrow, 1) < Sheets(combo).Cells(datarow, crashroutecol) Then  'crash route higher than hitlist route
        HLrow = HLrow + 1  'go to next hitlist row
    ElseIf Sheets("HitList").Cells(HLrow, 1) = Sheets(combo).Cells(datarow, crashroutecol) And _
    Sheets("Hitlist").Cells(HLrow, 2) > Sheets(combo).Cells(datarow, crashMPcol) Then  'route matches, but MP on hitlist is higher than crash MP
        datarow = datarow + 1  'go to next crash row
    ElseIf Sheets("HitList").Cells(HLrow, 1) = Sheets(combo).Cells(datarow, crashroutecol) And _
    Sheets("Hitlist").Cells(HLrow, 2) < Sheets(combo).Cells(datarow, crashMPcol) Then  'route matches but crash MP is greater than MP on hitlist
        HLrow = HLrow + 1  'go to next hitlist row
    ElseIf Sheets(combo).Cells(datarow, crashroutecol) = "" Then
        Exit Do
    Else
        MsgBox "Uh oh Camille! You need to add some more if statements", vbOKOnly, "Something got skipped"
    End If
Loop

'sort blank rows to bottom
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, crashroutecol), Cells(Tnumrow, crashroutecol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(combo).Sort
    .SetRange Range(Cells(1, 1), Cells(Tnumrow, Tnumcol))
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'deletes the hitlist sheet now that's it's not needed
Application.DisplayAlerts = False
If SheetExists("HitList") Then
    Sheets("HitList").Delete
End If
Application.DisplayAlerts = True


End Sub

Sub finishCAMScrashes()

Dim ComRow As Long, ComCol As Integer, GenRow As Long, GenCol As Integer, VehRow As Long, VehCol As Integer, lastcol As Integer
Dim GenRow2 As Long, VehRow2 As Long
'the first column in each file is "CRASH_ID"
Dim combo As String
Dim VehNumCol As Integer
Dim comCID, genCID, vehCID
Dim i As Integer
Dim DateTimeCol As Integer
Dim routecol As Integer
Dim dircol As Integer
Dim Route, Direction

combo = "Combined_Crash"

Application.ScreenUpdating = False

'Delete the CRASHDATETIME Column in the Vehicle and Crash sheet (don't need it b/c the general crash sheet has it too)
DateTimeCol = 1
Do Until Sheets("Vehicle").Cells(1, DateTimeCol) = "CRASH_DATETIME"
    DateTimeCol = DateTimeCol + 1
Loop
Sheets("Vehicle").Cells(1, DateTimeCol).EntireColumn.Delete

DateTimeCol = 1
Do Until Sheets("Crash").Cells(1, DateTimeCol) = "CRASH_DATETIME"
    DateTimeCol = DateTimeCol + 1
Loop
Sheets("Crash").Cells(1, DateTimeCol).EntireColumn.Delete

'Identify the vehicle number column in the vehilce file
VehNumCol = 1
Do Until Sheets("Vehicle").Cells(1, VehNumCol) = "VEHICLE_NUM"
    VehNumCol = VehNumCol + 1
Loop

'Identify the column extent of the Combined, General, and Vehicle crash data
ComCol = Sheets(combo).Range("A1").End(xlToRight).Column
GenCol = Sheets("Crash").Range("A1").End(xlToRight).Column
VehCol = Sheets("Vehicle").Range("A1").End(xlToRight).Column

'Identify the row extent of the Combined, General, and Vehicle crash data
ComRow = Sheets(combo).Range("A1").End(xlDown).row
GenRow = Sheets("Crash").Range("A1").End(xlDown).row
VehRow = Sheets("Vehicle").Range("A1").End(xlDown).row
'copy these for later
GenRow2 = GenRow
VehRow2 = VehRow

'Re-sort each crash dataset
'Sort combined data by Crash_ID
ActiveWorkbook.Worksheets(combo).Activate
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(ComRow, 1)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range(Cells(1, 1), Cells(ComRow, ComCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort crash data by Crash_ID
ActiveWorkbook.Worksheets("Crash").Activate
ActiveWorkbook.Worksheets("Crash").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Crash").Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(GenRow, 1)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Crash").Sort
        .SetRange Range(Cells(1, 1), Cells(GenRow, GenCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort vehicle data by Crash_ID and Vehicle Num
ActiveWorkbook.Worksheets("Vehicle").Activate
ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(VehRow, 1)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Vehicle").Sort.SortFields.Add Key:=Range(Cells(2, VehNumCol), Cells(VehRow, VehNumCol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Vehicle").Sort
        .SetRange Range(Cells(1, 1), Cells(VehRow, VehCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Combine crash, location, vehicle, and rollup data into a single database based on Crash ID

''''  ***  Camille: This was Josh's code (adapted to my variables). I "kan-de-dong" but still don't understand how it gives the correct solution.
''Cycle through crashes on combined crash sheet
'Do While Sheets(combo).Cells(ComRow, 1) <> ""
'    'Assign Crash ID values from each dataset
'    comCID = Sheets(combo).Cells(ComRow, 1)
'    genCID = Sheets("Crash").Cells(GenRow, 1)
'    vehCID = Sheets("Vehicle").Cells(VehRow, 1)
'
'    'If all other crash IDs are higher than the Combined CID, then add go to next Combined CID row
'    If comCID < genCID And comCID < vehCID Then
'        ComRow = ComRow + 1
'        comCID = Sheets(combo).Cells(ComRow, 1)
'    'If all other crash IDs are blank, then move to next Combined CID row
'    ElseIf genCID = "" And vehCID = "" Then
'        ComRow = ComRow + 1
'    End If
'
'    'If combo and general crash IDs are equal, then copy crash data over to rollup sheet
'    If comCID = genCID Then
'        For i = 2 To GenCol
'            Sheets(combo).Cells(ComRow, ComCol - 1 + i) = Sheets("Crash").Cells(GenRow, i)
'        Next i
'        GenRow = GenRow + 1
'    'If crash ID is less than rollup ID, go to next crash ID
'    ElseIf genCID < comCID Then
'        GenRow = GenRow + 1
'    'Stay at same crash ID if greater than rollup ID
'    ElseIf genCID > comCID Then
'        'Nothing
'    End If
'
'    'If combo and vehicle IDs are equal, then copy vehicle data over to combo sheet
'    If comCID = vehID Then
'        For i = 2 To VehCol
'            Sheets(combo).Cells(ComRow, ComCol + GenCol - 2 + i) = Sheets("Vehicle").Cells(VehRow, i)
'        Next i
'        VehRow = VehRow + 1
'    'If vehicle ID is less than combo ID, go to next vehicle ID
'    ElseIf vehCID < comCID Then
'        VehRow = VehRow + 1
'    'Stay at same vehicle ID if greater than rollup ID
'    ElseIf vehCID > comCID Then
'        'Nothing
'    End If
'
'Loop

'''''''''''''''''''''''''''''''''Updated to the copy then filter method 2021'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'paste from crash sheet
Sheets("Crash").Cells(1, 1).Resize(GenRow, GenCol).Copy Destination:=Sheets(combo).Cells(1, ComCol + 1)

'paste from vehicle sheet
Sheets("Vehicle").Cells(1, 1).Resize(VehRow, VehCol).Copy Destination:=Sheets(combo).Cells(1, ComCol + GenCol + 1)

'change column numbers to match location of CID columns on combo sheet
lastcol = ComCol + GenCol + VehCol
VehCol = ComCol + GenCol + 1
GenCol = ComCol + 1


'loop and filter
ComRow = 2                'Beginning row to begin cycling through
GenRow = 2
VehRow = 2
Sheets(combo).Activate
Do While Cells(ComRow, 1) <> ""
    If Cells(ComRow, 1) = Cells(GenRow, GenCol) Then  'if the general crash CID is equal to the combo CID
        
        'find a matching vehicle CID
        If Cells(ComRow, 1) = Cells(VehRow, VehCol) Then  'if the vehicle CID is equal to the combo CID do nothing
            ComRow = ComRow + 1
            GenRow = GenRow + 1
            VehRow = VehRow + 1
        ElseIf Cells(ComRow, 1) > Cells(VehRow, VehCol) Then  'if the vehicle CID is less than the combo CID
            Range(Cells(VehRow, VehCol), Cells(VehRow, lastcol)).Clear  'clear the row
            VehRow = VehRow + 1  'go to the next vehicle CID row
        ElseIf Cells(ComRow, 1) < Cells(VehRow, VehCol) Then  'if the vehicle CID is greater than the combo CID
            Range(Cells(ComRow, 1), Cells(ComRow, ComCol)).Clear  'clear the combo CID row
            ComRow = ComRow + 1
            MsgBox "row deleted", vbOKOnly, "FYI"
        End If
    
    ElseIf Cells(ComRow, 1) > Cells(GenRow, GenCol) Then  'if the general crash CID is less than the combo CID
        Range(Cells(GenRow, GenCol), Cells(GenRow, VehCol - 1)).Clear 'clear the row
        GenRow = GenRow + 1  'go to the next general crash CID
    ElseIf Cells(ComRow, 1) < Cells(GenRow, GenCol) Then  'if the general CID is greater than the combo CID
        Range(Cells(ComRow, 1), Cells(ComRow, ComCol)).Clear  'clear the combo CID row
        ComRow = ComRow + 1
        MsgBox "row deleted", vbOKOnly, "FYI"
    End If
Loop

'Sort blank rows to bottom
'Sort combined data by Crash_ID
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(ComRow, 1)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range(Cells(1, 1), Cells(ComRow, ComCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort crash data by Crash_ID
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, GenCol), Cells(GenRow2, GenCol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range(Cells(1, GenCol), Cells(GenRow2, VehCol - 1))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort vehicle data by Crash_ID and Vehicle Num
ActiveWorkbook.Worksheets(combo).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(combo).Sort.SortFields.Add Key:=Range(Cells(2, VehCol), Cells(VehRow2, VehCol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(combo).Sort
        .SetRange Range(Cells(1, VehCol), Cells(VehRow2, lastcol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Clear extra rows
Range(Cells(ComRow, GenCol), Cells(GenRow2, lastcol)).Clear

'Delete extra CID columns
Columns(GenCol).Delete
Columns(VehCol - 1).Delete





'ComRow = 2                'Beginning row to begin cycling through
'GenRow = 2
'VehRow = 2

'Copy header values to combination sheet
'For i = 2 To GenCol
'    Sheets(combo).Cells(1, ComCol - 1 + i) = Sheets("Crash").Cells(1, i)
'Next i
'For i = 2 To VehCol
'    Sheets(combo).Cells(1, ComCol + GenCol - 2 + i) = Sheets("Vehicle").Cells(1, i)
'Next i

'Do While Sheets(combo).Cells(ComRow, 1) <> ""
'    If Sheets(combo).Cells(ComRow, 1) = Sheets("Crash").Cells(GenRow, 1) Then  'if the general crash CID is equal to the combo CID
'
'        'find a matching vehicle CID
'        If Sheets(combo).Cells(ComRow, 1) = Sheets("Vehicle").Cells(VehRow, 1) Then  'if the vehicle CID is equal to the combo CID
'            For i = 2 To GenCol 'paste the general crash data
'                Sheets(combo).Cells(ComRow, i + ComCol - 1) = Sheets("Crash").Cells(GenRow, i)
'            Next i
'            For i = 2 To VehCol 'paste the vehicle data
'                Sheets(combo).Cells(ComRow, i + ComCol + GenCol - 2) = Sheets("Vehicle").Cells(VehRow, i)
'            Next i
'            ComRow = ComRow + 1
'            GenRow = GenRow + 1
'            VehRow = VehRow + 1
'        ElseIf Sheets(combo).Cells(ComRow, 1) > Sheets("Vehicle").Cells(VehRow, 1) Then  'if the vehicle CID is less than the combo CID
'            VehRow = VehRow + 1  'go to the next vehicle CID row
'        ElseIf Sheets(combo).Cells(ComRow, 1) < Sheets("Vehicle").Cells(VehRow, 1) Then  'if the vehicle CID is greater than the combo CID
'            ComRow = ComRow - 1
'            Sheets(combo).Cells(ComRow + 1, 1).EntireRow.Delete  'delete the combo CID row
'            ComRow = ComRow + 1
'            MsgBox "row deleted", vbOKOnly, "FYI"
'        End If
'
'    ElseIf Sheets(combo).Cells(ComRow, 1) > Sheets("Crash").Cells(GenRow, 1) Then  'if the general crash CID is less than the combo CID
'        GenRow = GenRow + 1  'go to the next general crash CID
'    ElseIf Sheets(combo).Cells(ComRow, 1) < Sheets("Crash").Cells(GenRow, 1) Then  'if the general CID is greater than the combo CID
'        ComRow = ComRow - 1
'        Sheets(combo).Cells(ComRow + 1, 1).EntireRow.Delete  'delete the combo CID row
'        ComRow = ComRow + 1
'        MsgBox "row deleted", vbOKOnly, "FYI"
'    End If
'Loop

'deletes the general crash sheet now that's it's not needed
Application.DisplayAlerts = False
If SheetExists("Crash") Then
    Sheets("Crash").Delete
End If
Application.DisplayAlerts = True

'deletes the vehicle sheet now that's it's not needed
Application.DisplayAlerts = False
If SheetExists("Vehicle") Then
    Sheets("Vehicle").Delete
End If
Application.DisplayAlerts = True

Application.ScreenUpdating = True
Application.ScreenUpdating = False

'format the CRASH_DATETIME column
DateTimeCol = 1
Do Until Sheets(combo).Cells(1, DateTimeCol) = "CRASH_DATETIME"
    DateTimeCol = DateTimeCol + 1
Loop
Sheets(combo).Activate
Range(Cells(2, DateTimeCol), Cells(2, DateTimeCol).End(xlDown)).NumberFormat = "mmm d, yyyy h:mm:ss AM/PM"

dircol = 1
Do Until Sheets(combo).Cells(1, dircol) = "DIRECTION"
    dircol = dircol + 1
Loop
routecol = 1
Do Until Sheets(combo).Cells(1, routecol) = "ROUTE_ID"
    routecol = routecol + 1
Loop

'adds a new column for the label and fills it in
'also changes N direction to P if route is not an interstate
Columns(dircol + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Sheets(combo).Cells(1, dircol + 1) = "LABEL"
ComRow = 2
Do While Sheets(combo).Cells(ComRow, routecol) <> ""
    Route = CStr(Sheets(combo).Cells(ComRow, routecol))
    Direction = Sheets(combo).Cells(ComRow, dircol)
    If Direction = "N" And _
    (CLng(Route) <> 15 And CLng(Route) <> 70 And _
    CLng(Route) <> 80 And CLng(Route) <> 84 And _
    CLng(Route) <> 85 And CLng(Route) <> 215) Then
        Direction = "P"
        Sheets(combo).Cells(ComRow, dircol) = Direction
    End If
    Direction = Sheets(combo).Cells(ComRow, dircol)
    If Len(Route) = 1 Then
        Route = "000" & Route & Direction
    ElseIf Len(Route) = 2 Then
        Route = "00" & Route & Direction
    ElseIf Len(Route) = 3 Then
        Route = "0" & Route & Direction
    Else
        Route = Route & Direction
    End If
    Sheets(combo).Cells(ComRow, dircol + 1) = Route
    Sheets(combo).Cells(ComRow, dircol + 1).NumberFormat = "@"
    ComRow = ComRow + 1
Loop

End Sub

