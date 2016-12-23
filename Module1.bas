Attribute VB_Name = "Module1"
'   Create_MISGA_Schedule.bas Module:    Last revised Jan 8, 2013       by Tom Rowlett
'
'   Purpose:     To start embedding in code the date of the last revision
'
'               1/8/13  Implement 14 digit Event ID
'
'
Public Const mrAssoc As String = "Musket Ridge MISGA Associates"
Public Const homeCourse As String = "Musket Ridge"
Public Const homeID As String = "232"
Public Const homeMixerCost As Integer = 47
Public Const homeTimeA As String = "9:00 AM"
Public Const homeTimeB As String = "8:30 AM"
Public Const HostPlayerEst As Integer = 40
Public Const homePhone As String = "301-293-9930, x110"
Public Const homePlayerLimit As Integer = 60
Public Const awayPlayerLimit As Integer = 20
Public Const schedFileName As String = "schedule.txt"
Public scheduleFileName As String
Public Const pathPrefix As String = "C:\Users\Public\Documents\"
Public Const pathSuffix As String = " Schedule\"
Public Const Oldpath As String = "C:\Users\Public\Documents\Musket Ridge MISGA\2017 Schedule\"
Public path As String
Public wrkBookName As String
Public scheduleYear As String

Public Const S_No As Integer = 1
Public Const S_Date As Integer = 2
Public Const S_Time As Integer = 3
Public Const S_AH As Integer = 4
Public Const S_Cost As Integer = 5
Public Const S_Event As Integer = 6
Public Const S_Club As Integer = 7
Public Const S_PLimit As Integer = 8
Public Const S_Status As Integer = 9
Public Const S_EstPlay As Integer = 10
Public Const S_EstDay As Integer = 11
Public Const S_Deadline As Integer = 12
Public Const S_ID As Integer = 13
Public Const S_Phone As Integer = 14
Public Const S_SRule As Integer = 15
Public Const S_Guest As Integer = 16
Public Const S_Post As Integer = 17
Public Const S_HostID As Integer = 18
Public Const LastCol As Integer = 18




Public TabName(10) As String
Public selName As String
Public selTabNdx As Integer
Public x        As String
Public tabNames(1 To 6) As String
Public tabNdx(1 To 6) As Integer
Public lastTabNdx As Integer
Public costTxt As String
Public Const eventsTabName As String = "Clubs + Dates"
Public Const scheduleTabName As String = "Schedule"
Public Const misgaTabName As String = "MISGA"
Public Const homeTabName As String = "Home"
Public Const awayTabName As String = "Away"
Public Const mixerDayTabName As String = "HomeDay"
Public Const clubTabName As String = "Club"
Public Const highDate As Date = "12/31/2099"
Public Const lowDate As Date = "1/1/2010"
Public Const dateFormat As String = "ddd, mmm d"
Public Const away As String = "Away"
Public Const home As String = "Home"
Public Const MISGA As String = "MISGA"
Public Const iMISGA As String = "MiSGA"
Public Const club As String = "Club"
Public Const ryderCup As String = "Ryder Cup"
Public Const openDate As String = "open"
Public Const interClub As String = "Inter-club Mixer"
Public Const delimiter As String = ";"
Public Const timeFormat As String = "h:mm a/p"
Public quoteTxt As String
Public eventsTabNdx As Integer
Public scheduleTabNdx As Integer
Public misgaTabNdx As Integer
Public awayTabNdx As Integer
Public homeTabNdx As Integer
Public clubTabNdx As Integer
Public mixerDayTabNdx As Integer
Public tempLastRow As Integer
Public eventName As String
Public eventNumber As Integer
Public fromRowNdx As Integer
Public postDate As String

Public mixerDate As Date
Public schDate As Date
Public schPlace As String
Public curDate As Date
Public prevDate As Date
Public prevPlace As String
Public tDate As String
Public eventNo As Integer
Public playerLimitForDay As Integer
Public clubName As String
' Public cost As Integer

Public schRecord As String
Public clubs As String
Public evDateTxt As String
Public eolTxt As String
Public timeTxt As String
Public deadline As String
Public playerLimitTxt As String
Public haveHome As Boolean
Public prevHome As Boolean
Public mdt As Date
Public sdt As Date
Public print_area As String
Public workID As String
Public clubID As String
Public phoneTxt As String
Public guestTxt As String
Public specialRuleTxt As String
Public IsGuest As String
Public Const Guest As String = "Guest"
Public eventtype As String







' Sub Dates_to_text()
'    Organize_worksheets
'    ActiveWorkbook.Worksheets(eventsTabNdx).Activate
'    Set eventsSheet = ActiveSheet
'
'    For j = 2 To 25
'        tmpDate = eventsSheet.Cells(j, 5).Value
'        eventsSheet.Cells(j, 27).Value = Format(tmpDate, "dddd, mmmm dd, yyyy")
'        tmpDate = vfwEvents.Cells(j, 11).Value
'        vfwEvents.Cells(j, 28).Value = Format(tmpDate, "dddd, mmmm dd, yyyy")
'        Next j
'
' End Sub

Sub Organize_worksheets()
    tabNames(1) = eventsTabName
    tabNames(2) = scheduleTabName
'    tabNames(3) = awayTabName
'    tabNames(4) = homeTabName
    tabNames(3) = misgaTabName
    tabNames(4) = mixerDayTabName
    tabNames(5) = clubTabName
    eventsTabNdx = 0
    scheduleTabNdx = 0
'    awayTabNdx = 0
'    homeTabNdx = 0
    misgaTabNdx = 0
    mixerDayTabNdx = 0
    clubTabNdx = 0
    
    For n = 1 To 5
        tabNdx(n) = 0
        Next n
    
    x = ActiveWorkbook.Worksheets.Count
    For i = 1 To x
        TabName(i) = ActiveWorkbook.Worksheets(i).Name
        For n = 1 To 5
        If TabName(i) = tabNames(n) Then
                tabNdx(n) = i
                End If
            Next n
        Next i
    eventsTabNdx = tabNdx(1)
    scheduleTabNdx = tabNdx(2)
'    awayTabNdx = tabNdx(3)
'    homeTabNdx = tabNdx(4)
    misgaTabNdx = tabNdx(3)
    mixerDayTabNdx = tabNdx(4)
    clubTabNdx = tabNdx(5)
    
    lastTabNdx = i
End Sub




Sub Create_MISGA_Schedule()
Attribute Create_MISGA_Schedule.VB_Description = "Create Schedule"
Attribute Create_MISGA_Schedule.VB_ProcData.VB_Invoke_Func = "s\n14"
'
'   Macro created on December 5, 2008
'
'   Create_MISGA_Schedule takes a worksheet of events and reorganizes
'   it in a chrnological set of events.
'
    wrkBookName = ThisWorkbook.Name
    Dim bookName As String
    bookName = wrkBookName
    scheduleYear = Left(wrkBookName, 4)
    path = pathPrefix + homeCourse + " MISGA\" + scheduleYear + pathSuffix
    scheduleFileName = homeID + "-" + schedFileName
    Organize_worksheets
    
'
'   Clear Schedule Worksheet
'
    ActiveWorkbook.Worksheets(scheduleTabNdx).Activate
    Range("A2:Z150").Clear
    
'
'   Clear Away Worksheet
'
'    ActiveWorkbook.Worksheets(awayTabNdx).Activate
'    Range("A2:Z150").Clear
'
'   Clear Home Worksheet
'
'   ActiveWorkbook.Worksheets(homeTabNdx).Activate
'   Range("A2:Z150").Clear
   
'
'    MsgBox "Schedule, Away, and Home Sheets Cleared"
'
'
    ActiveWorkbook.Worksheets(mixerDayTabNdx).Activate
    Set mixerDaySheet = ActiveSheet
    ActiveWorkbook.Worksheets(misgaTabNdx).Activate
    Set misgaEvents = ActiveSheet
    ActiveWorkbook.Worksheets(clubTabNdx).Activate
    Set clubEvents = ActiveSheet
    ActiveWorkbook.Worksheets(eventsTabNdx).Activate
    Set mixerEvents = ActiveSheet
    ActiveWorkbook.Worksheets(scheduleTabNdx).Activate
    Set scheduleEvents = ActiveSheet
    
    fromRowNdx = 2
    fromColNdx = 2
    toRowNdx = 2
    toColNdx = 2
    workID = ""
    clubID = ""
    
'
'   Take the Club events planned from the EVENTS Sheet and
'   organize them in chronological order in the SCHEDULE Sheet
'
    Do While (mixerEvents.Cells(fromRowNdx, 4).Value <> "")   ' quit when Club col is blank
        eventName = interClub
        clubName = mixerEvents.Cells(fromRowNdx, 4).Value    ' Club Short Name
        eventNumber = mixerEvents.Cells(fromRowNdx, 1).Value
        If (mixerEvents.Cells(fromRowNdx, 2).Value <> "") Then
            eventName = mixerEvents.Cells(fromRowNdx, 2).Value
            End If
        scheduleEvents.Cells(toRowNdx, S_Event).Value = eventName                                     ' Event
        scheduleEvents.Cells(toRowNdx, S_Club).Value = clubName
        If ((mixerEvents.Cells(fromRowNdx, 6).Value = "") Or _
                (mixerEvents.Cells(fromRowNdx, 9).Value = "")) Then
                cost = 0
            Else
            clubID = mixerEvents.Cells(fromRowNdx, 3).Value
            mixerDate = mixerEvents.Cells(fromRowNdx, 6).Value
            workID = homeID + Format(mixerDate, "yymmddhh") + clubID
            scheduleEvents.Cells(toRowNdx, S_Date).Value = mixerEvents.Cells(fromRowNdx, 6).Value    ' date
            scheduleEvents.Cells(toRowNdx, S_Time).Value = mixerEvents.Cells(fromRowNdx, 10).Value    ' time
            scheduleEvents.Cells(toRowNdx, S_Time).NumberFormat = timeFormat
            scheduleEvents.Cells(toRowNdx, S_AH).Value = away                                      ' Away/Home
            scheduleEvents.Cells(toRowNdx, S_PLimit).Value = mixerEvents.Cells(fromRowNdx, 11).Value    ' Host Player Limit
            scheduleEvents.Cells(toRowNdx, S_Post).Value = mixerEvents.Cells(fromRowNdx, 8).Value    ' Away Posting Date
            scheduleEvents.Cells(toRowNdx, S_Post).NumberFormat = "ddd, mmm d"
            scheduleEvents.Cells(toRowNdx, S_EstPlay).Value = mixerEvents.Cells(fromRowNdx, 11).Value    ' Player Limit
            scheduleEvents.Cells(toRowNdx, S_Status).Value = mixerEvents.Cells(fromRowNdx, 9).Value   ' Status
            scheduleEvents.Cells(toRowNdx, S_Deadline).Value = mixerEvents.Cells(fromRowNdx, 7).Value    ' Deadline Date
            scheduleEvents.Cells(toRowNdx, S_Deadline).NumberFormat = "ddd, mmm d"
            Set cost = mixerEvents.Cells(fromRowNdx, 12)
            If (cost = 0) Then
                costTxt = "tbd "
                Else
                costTxt = Application.WorksheetFunction.Text(cost.Value, "$#0")   ' pickup cost
                If (UCase(mixerEvents.Cells(fromRowNdx, 13).Value) = "Y") Then                      ' check for CASH ONLY
                    costTxt = costTxt + "*"
                    End If
                End If
            scheduleEvents.Cells(toRowNdx, S_Cost).Clear
            scheduleEvents.Cells(toRowNdx, S_Cost).Value = costTxt                     ' if so, add an asterisk
            scheduleEvents.Cells(toRowNdx, S_Cost).HorizontalAlignment = xlRight
            scheduleEvents.Cells(toRowNdx, S_ID).Value = workID
            scheduleEvents.Cells(toRowNdx, S_ID).NumberFormat = "00000000000000"
            scheduleEvents.Cells(toRowNdx, S_Phone).Value = mixerEvents.Cells(fromRowNdx, 26)  ' Pro Shop Phone Number
            scheduleEvents.Cells(toRowNdx, S_SRule).Value = mixerEvents.Cells(fromRowNdx, 14)   ' Host Club Special Rule
            scheduleEvents.Cells(toRowNdx, S_Guest).Value = ""
            scheduleEvents.Cells(toRowNdx, S_HostID).Value = clubID
            toRowNdx = toRowNdx + 1
            End If
            '                                                   If no Home Date or not yet confirmed, skip the entry
        If ((mixerEvents.Cells(fromRowNdx, 15).Value = "") Or _
            (mixerEvents.Cells(fromRowNdx, 18).Value = "")) Then
            i = i
            Else
            mixerDate = mixerEvents.Cells(fromRowNdx, 15).Value
'            clubID = mixerEvents.Cells(fromRowNdx, 3).Value
            workID = homeID + Format(mixerDate, "yymmddhh") + homeID
            scheduleEvents.Cells(toRowNdx, S_Event).Value = eventName                                 ' Event
            scheduleEvents.Cells(toRowNdx, S_Club).Value = clubName                                  ' Club Short Name
            scheduleEvents.Cells(toRowNdx, S_Date).Value = mixerEvents.Cells(fromRowNdx, 15).Value   ' Date
            scheduleEvents.Cells(toRowNdx, S_Time).Value = mixerEvents.Cells(fromRowNdx, 19).Value   ' Time
            scheduleEvents.Cells(toRowNdx, S_Time).NumberFormat = timeFormat
            scheduleEvents.Cells(toRowNdx, S_AH).Value = home                                     ' Away/Home
'            scheduleEvents.Cells(toRowNdx, S_PLimit).Value = mixerEvents.Cells(fromRowNdx, 21).Value   ' Visitor Player Limit
            scheduleEvents.Cells(toRowNdx, S_PLimit).Value = homePlayerLimit           ' Visitor Player Limit for Hom Event
            scheduleEvents.Cells(toRowNdx, S_Status).Value = mixerEvents.Cells(fromRowNdx, 18).Value   ' Status
            scheduleEvents.Cells(toRowNdx, S_EstPlay).Value = mixerEvents.Cells(fromRowNdx, 20).Value   ' Estimated Players
            ' ScheduleEvents.Cells(toRowNdx, 10).Value = mixerEvents.Cells(fromRowNdx, 21).Value  ' Player Limit for Visitor
            scheduleEvents.Cells(toRowNdx, S_Deadline).Value = mixerEvents.Cells(fromRowNdx, 16).Value    ' Deadline Date
            scheduleEvents.Cells(toRowNdx, S_Deadline).NumberFormat = "ddd, mmm d"
            Set HCost = mixerEvents.Cells(fromRowNdx, 22)
            costTxt = ""
            If (HCost.Value = "") Then
                HCost.Value = homeMixerCost
                End If
            costTxt = Application.WorksheetFunction.Text(HCost.Value, "$#0")   ' pickup cost
            If (UCase(mixerEvents.Cells(fromRowNdx, 23).Value = "Y")) Then
                costTxt = costTxt + "*"
                End If
            scheduleEvents.Cells(toRowNdx, S_Cost).Value = costTxt
            scheduleEvents.Cells(toRowNdx, S_Cost).HorizontalAlignment = xlRight
'            If (HCost.Value = "") Then
'                scheduleEvents.Cells(toRowNdx, S_Cost).Value = homeMixerCost                        ' pickup cost
'                End If
'            scheduleEvents.Cells(toRowNdx, S_Cost).NumberFormat = "$#0 "
            scheduleEvents.Cells(toRowNdx, S_ID).Value = workID
            scheduleEvents.Cells(toRowNdx, S_ID).NumberFormat = "00000000000000"
            scheduleEvents.Cells(toRowNdx, S_Phone).Value = homePhone        ' Pro Shop Phone Number for MR
            scheduleEvents.Cells(toRowNdx, S_SRule).Value = mixerEvents.Cells(fromRowNdx, 24)   ' Host Club Special Rule
            IsGuest = mixerEvents.Cells(fromRowNdx, 25).Value
            If (UCase(Trim(IsGuest)) = "Y") Then
                scheduleEvents.Cells(toRowNdx, S_Guest).Value = Guest
                Else
                scheduleEvents.Cells(toRowNdx, S_Guest).Value = ""
                End If
            scheduleEvents.Cells(toRowNdx, S_Post).Value = mixerEvents.Cells(fromRowNdx, 17).Value           ' Home Post Date
            scheduleEvents.Cells(toRowNdx, S_Post).NumberFormat = "ddd, mmm d"
            scheduleEvents.Cells(toRowNdx, S_HostID).Value = homeID

            toRowNdx = toRowNdx + 1
            End If
        fromRowNdx = fromRowNdx + 1
        Loop
'
'   Sort the EVENTS by date so that MISGA Events can be integrated
'
    tempLastRow = toRowNdx - 1
    scheduleEvents.Range(Cells(1, 1), Cells(tempLastRow, LastCol)).Sort _
        Key1:=scheduleEvents.Range("B1"), _
        Key2:=scheduleEvents.Range("G1"), _
        Header:=xlYes
'
'   Take the MISGA events and combine them with the Inter-Club mixers
'
    fromRowNdx = 2
    Do While (misgaEvents.Cells(fromRowNdx, 1).Value <> "")
        If (misgaEvents.Cells(fromRowNdx, 9).Value = "C") Then    ' Add if Status = "C" for confirmed
            For i = 2 To 9
                scheduleEvents.Cells(toRowNdx, i).Value = misgaEvents.Cells(fromRowNdx, i).Value
                scheduleEvents.Cells(toRowNdx, i).NumberFormat = misgaEvents.Cells(fromRowNdx, i).NumberFormat
                Next i
            scheduleEvents.Cells(toRowNdx, S_Deadline).Value = misgaEvents.Cells(fromRowNdx, 12).Value  ' Deadline
            scheduleEvents.Cells(toRowNdx, S_Deadline).NumberFormat = "ddd, mmm d"
            scheduleEvents.Cells(toRowNdx, S_Phone).Value = misgaEvents.Cells(fromRowNdx, 13).Value
            scheduleEvents.Cells(toRowNdx, S_Post).Value = misgaEvents.Cells(fromRowNdx, 16).Value      ' Posting Date
            scheduleEvents.Cells(toRowNdx, S_Post).NumberFormat = "ddd, mmm d"
            IsGuest = misgaEvents.Cells(fromRowNdx, 14).Value
            If (UCase(Trim(IsGuest)) = "Y") Then
                scheduleEvents.Cells(toRowNdx, S_Guest).Value = Guest
                Else
                scheduleEvents.Cells(toRowNdx, S_Guest).Value = ""
                End If
            clubID = Format(misgaEvents.Cells(fromRowNdx, 10).Value, "000")
            workID = homeID + Format(misgaEvents.Cells(fromRowNdx, 2).Value, "yymmddhh") + clubID
            scheduleEvents.Cells(toRowNdx, S_ID).Value = workID
            scheduleEvents.Cells(toRowNdx, S_ID).NumberFormat = "00000000000000"
            scheduleEvents.Cells(toRowNdx, S_AH).Value = MISGA
            scheduleEvents.Cells(toRowNdx, S_AH).Value = misgaEvents.Cells(fromRowNdx, 17).Value        ' Event Type
            scheduleEvents.Cells(toRowNdx, S_HostID).Value = clubID
            eventtype = scheduleEvents.Cells(toRowNdx, S_AH).Value
            toRowNdx = toRowNdx + 1
            End If
        fromRowNdx = fromRowNdx + 1
        Loop
'
'   Sort the EVENTS by date so that Club Events can be integrated
'
    tempLastRow = toRowNdx - 1
    scheduleEvents.Range(Cells(1, 1), Cells(tempLastRow, LastCol)).Sort _
        Key1:=scheduleEvents.Range("B1"), _
        Key2:=scheduleEvents.Range("G1"), _
        Header:=xlYes
'
'   Take the Club events and combine them with the Inter-Club mixers
'
    fromRowNdx = 2
    Do While (clubEvents.Cells(fromRowNdx, 1).Value <> "")
        If (clubEvents.Cells(fromRowNdx, 10).Value = "C") Then    ' Add if Status = "C" for confirmed
        j = 2
        For i = 2 To 10
            If (i = 6) Then
                i = i + 1
                End If
            scheduleEvents.Cells(toRowNdx, j).Value = clubEvents.Cells(fromRowNdx, i).Value
            scheduleEvents.Cells(toRowNdx, j).NumberFormat = clubEvents.Cells(fromRowNdx, i).NumberFormat
            j = j + 1
            Next i
        Set HCost = clubEvents.Cells(fromRowNdx, 5)
'         costTxt = ""
        If (HCost.Value = "") Then
            HCost.Value = homeMixerCost
            End If
        costTxt = Application.WorksheetFunction.Text(HCost.Value, "$#0")   ' pickup cost
        If ((UCase(clubEvents.Cells(fromRowNdx, 6).Value) = "Y")) Then
            costTxt = costTxt + "*"
            End If
        scheduleEvents.Cells(toRowNdx, S_Cost).Value = costTxt
        scheduleEvents.Cells(toRowNdx, S_Cost).HorizontalAlignment = xlRight
         
        scheduleEvents.Cells(toRowNdx, S_EstPlay).Value = clubEvents.Cells(fromRowNdx, 12).Value    ' Total Player Estimate
        scheduleEvents.Cells(toRowNdx, S_EstDay).Value = clubEvents.Cells(fromRowNdx, 12).Value     ' Put in Player Estimate for the day
        scheduleEvents.Cells(toRowNdx, S_Deadline).Value = clubEvents.Cells(fromRowNdx, 13).Value  ' Deadline
        scheduleEvents.Cells(toRowNdx, S_Deadline).NumberFormat = "ddd, mmm d"
        scheduleEvents.Cells(toRowNdx, S_Phone).Value = clubEvents.Cells(fromRowNdx, 14).Value
        scheduleEvents.Cells(toRowNdx, S_SRule).Value = clubEvents.Cells(fromRowNdx, 16).Value     ' Tee Selection or Special Rule
        scheduleEvents.Cells(toRowNdx, S_Post).Value = clubEvents.Cells(fromRowNdx, 17).Value      ' Posting Date
        scheduleEvents.Cells(toRowNdx, S_Post).NumberFormat = "ddd, mmm d"
        IsGuest = clubEvents.Cells(fromRowNdx, 15).Value
        If (UCase(Trim(IsGuest)) = "Y") Then
            scheduleEvents.Cells(toRowNdx, S_Guest).Value = Guest
            Else
            scheduleEvents.Cells(toRowNdx, S_Guest).Value = ""
            End If
        clubID = Format(clubEvents.Cells(fromRowNdx, 11).Value, "000")
        workID = homeID + Format(clubEvents.Cells(fromRowNdx, 2).Value, "yymmddhh") + clubID
        scheduleEvents.Cells(toRowNdx, S_ID).Value = workID
        scheduleEvents.Cells(toRowNdx, S_ID).NumberFormat = "00000000000000"
        scheduleEvents.Cells(toRowNdx, S_AH).Value = club
        scheduleEvents.Cells(toRowNdx, S_AH).Value = clubEvents.Cells(fromRowNdx, 18).Value        ' Event Type
        scheduleEvents.Cells(toRowNdx, S_HostID).Value = clubID
        toRowNdx = toRowNdx + 1
        End If
        fromRowNdx = fromRowNdx + 1
        Loop
'
'   Sort the EVENTS by date so that Mixer Days can be integrated
'
    tempLastRow = toRowNdx - 1
    scheduleEvents.Range(Cells(1, 1), Cells(tempLastRow, LastCol)).Sort _
        Key1:=scheduleEvents.Range("B1"), _
        Key2:=scheduleEvents.Range("G1"), _
        Header:=xlYes
'
'   Now add the Tuesdays that are NOT scheduled
'
    i = 2           ' used to track row in ScheduleEvents
    j = 2           ' Row in mixerDaySheet
    Do While (mixerDaySheet.Cells(j, 1).Value <> "")
        mdt = mixerDaySheet.Cells(j, 2).Value
        mixerDate = DateSerial(Year(mdt), Month(mdt), Day(mdt))
        If (i < toRowNdx) Then
            sdt = scheduleEvents.Cells(i, S_Date).Value
            schDate = DateSerial(Year(sdt), Month(sdt), Day(sdt))
            i = i
            Else
            schDate = highDate
            End If
        If (mixerDate < schDate) Then
            If (mixerDaySheet.Cells(j, 3).Value <> "NS") Then
            workID = Format(mixerDate, "yymmdd") + homeID
                scheduleEvents.Cells(toRowNdx, S_ID).Value = workID
                scheduleEvents.Cells(toRowNdx, S_Phone).Value = homePhone
                scheduleEvents.Cells(toRowNdx, S_Date).Value = mdt
                scheduleEvents.Cells(toRowNdx, S_Time).Value = mdt
                scheduleEvents.Cells(toRowNdx, S_Time).NumberFormat = timeFormat
                scheduleEvents.Cells(toRowNdx, S_AH).Value = openDate
                scheduleEvents.Cells(toRowNdx, S_Club).Value = "at " + homeCourse
                scheduleEvents.Cells(toRowNdx, S_Deadline).Value = mixerDate            ' Signup deadline
                scheduleEvents.Cells(toRowNdx, S_Deadline).NumberFormat = "ddd, mmm d"
                scheduleEvents.Cells(toRowNdx, 5).Value = 43                    ' Open Date cost to play at home
                scheduleEvents.Cells(toRowNdx, 5).NumberFormat = "$#0 "
                scheduleEvents.Cells(toRowNdx, S_Post).Value = mixerDate - 31      ' Posting Date
                scheduleEvents.Cells(toRowNdx, S_Post).NumberFormat = "ddd, mmm d"
                If (mixerDaySheet.Cells(j, 4).Value <> "") Then
                    scheduleEvents.Cells(toRowNdx, S_Event).Value = mixerDaySheet.Cells(j, 4).Value
                    End If
                toRowNdx = toRowNdx + 1
                End If
            j = j + 1
            Else
            If (mixerDate = schDate) Then
'                If (mixerDaySheet.Cells(j, 2).Value = "H") Then
'                    scheduleEvents.Cells(toRowNdx, 2).Value = scheduleEvents.Cells(i, 2).Value
'                    scheduleEvents.Cells(toRowNdx, 3).Value = homeTime
'                    scheduleEvents.Cells(toRowNdx, 4).Value = home
'                    scheduleEvents.Cells(toRowNdx, 5).Value = homeMixerCost
'                    scheduleEvents.Cells(toRowNdx, 5).NumberFormat = "$#0"
'                    scheduleEvents.Cells(toRowNdx, 6).Value = mixerDaySheet.Cells(j, 3)
'                    scheduleEvents.Cells(toRowNdx, 7).Value = mrAssoc                  ' Musket Ridge Associates
'                    scheduleEvents.Cells(toRowNdx, 8).Value = "C"       ' status
'                    scheduleEvents.Cells(toRowNdx, 9).Value = mixerDaySheet.Cells(j, 4)
'                    toRowNdx = toRowNdx + 1
'                    End If
                j = j + 1
                Else
                i = i + 1
                End If
            End If
        Loop
'
'   Sort them again in date order
'
    tempLastRow = toRowNdx - 1
    scheduleEvents.Range(Cells(1, 1), Cells(tempLastRow, LastCol)).Sort _
        Key1:=scheduleEvents.Range("M1"), _
        Header:=xlYes
'
'   Do Javascript file output of schedule for input to web site
'
'    Call Create_JavaScript_File(scheduleTabNdx)
'
'   Scan schedule and color home and MISGA entries
'
    fromRowNdx = 2
    Do While (scheduleEvents.Cells(fromRowNdx, S_Date).Value <> "")
        If (scheduleEvents.Cells(fromRowNdx, S_AH).Value = home) Then
            Call tintRowHome((fromRowNdx))
            End If
        If (scheduleEvents.Cells(fromRowNdx, S_AH).Value = away) Then
            Call tintRowAway((fromRowNdx))
            End If
        If (UCase(scheduleEvents.Cells(fromRowNdx, S_AH).Value) = MISGA) Then
            Call tintRowMISGA((fromRowNdx))
            End If
        If (scheduleEvents.Cells(fromRowNdx, S_AH).Value = openDate) Then
            Call tintRowOpen((fromRowNdx))
            End If
        If (scheduleEvents.Cells(fromRowNdx, S_AH).Value = club) Then
            Call tintRowClub((fromRowNdx))
            End If
        fromRowNdx = fromRowNdx + 1
        Loop
'
'   Scan schedule and eliminate displaying the number playing on the same date on multiple lines.
'
    Range("A2:Q2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    eventNo = 1
    prevDate = scheduleEvents.Cells(2, S_Date).Value         ' get first date in the schedule of events
    scheduleEvents.Cells(2, S_No).Value = eventNo          ' set event number col in row 2 = 1
    haveHome = False                                    ' set have Home Event indicator to false
    prevHome = False
    scheduleEvents.Cells(2, S_Date).NumberFormat = dateFormat
    playerLimitForDay = scheduleEvents.Cells(2, S_EstPlay).Value ' Initialize Player Limit For The Day
    If (scheduleEvents.Cells(2, S_AH).Value = home) Then         ' First row is a Home Event
        haveHome = True
        prevHome = True
        playerLimitForDay = HostPlayerEst
        End If
        
    fromRowNdx = 3                                      ' start checking at the second row of events
    Do While (scheduleEvents.Cells(fromRowNdx, S_Date).Value <> "")
        curDate = scheduleEvents.Cells(fromRowNdx, S_Date).Value
        If (curDate = prevDate) Then
        
            If (scheduleEvents.Cells(fromRowNdx, S_AH).Value = home) Then
                scheduleEvents.Cells(fromRowNdx, S_Date).Value = ""              ' Clear Date
                scheduleEvents.Cells(fromRowNdx, S_No).Value = eventNo         ' Set Event Number
                scheduleEvents.Cells(fromRowNdx, S_Time).Value = ""              ' Clear Time
                scheduleEvents.Cells(fromRowNdx, S_AH).Value = ""              ' Clear Home/Away
                scheduleEvents.Cells(fromRowNdx, S_Cost).Value = ""              ' Clear Cost
                scheduleEvents.Cells(fromRowNdx, S_Event).Value = ""              ' Clear Event
                playerLimitForDay = playerLimitForDay + scheduleEvents.Cells(fromRowNdx, S_EstPlay).Value
                haveHome = True
                End If
                
            scheduleEvents.Cells(fromRowNdx, S_Date).NumberFormat = dateFormat
            Else                                ' New Date
            Call UnderLine(fromRowNdx - 1, 17)    ' Underline last entry for date.
            
            If (prevHome) Then   ' was previous event at home
                scheduleEvents.Cells(fromRowNdx - 1, S_EstDay).Value = playerLimitForDay
                End If
            playerLimitForDay = scheduleEvents.Cells(fromRowNdx, S_EstPlay).Value   ' Get the estimated players
            If (scheduleEvents.Cells(fromRowNdx, S_AH).Value = home) Then
                playerLimitForDay = playerLimitForDay + HostPlayerEst
                haveHome = True
                Else
                haveHome = False
                End If
                
            eventNo = eventNo + 1
            End If
            
        scheduleEvents.Cells(fromRowNdx, S_No).Value = eventNo
        prevDate = curDate
        prevHome = haveHome
        scheduleEvents.Cells(fromRowNdx, S_Date).NumberFormat = dateFormat
        fromRowNdx = fromRowNdx + 1
        Loop
        scheduleEvents.Cells(fromRowNdx - 1, S_EstDay).Value = playerLimitForDay
        Call UnderLine(fromRowNdx - 1, 17)
        print_area = "$A$1:$K$" + Format(fromRowNdx - 1, "##0")
        ActiveSheet.PageSetup.PrintArea = print_area

        Call Create_TXT_File

End Sub

