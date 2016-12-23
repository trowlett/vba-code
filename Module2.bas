Attribute VB_Name = "Module2"
Function Close_line() As String
    Dim rtn As String
    rtn = ""
    rtn = clubs + delimiter
    rtn = rtn + costTxt + delimiter
    rtn = rtn + timeTxt + delimiter
    rtn = rtn + deadline + delimiter
    rtn = rtn + phoneTxt + delimiter + playerLimitTxt + delimiter + specialRuleTxt + delimiter + guestTxt
    rtn = rtn + delimiter + postDate
    Close_line = rtn

End Function

Function format_cost(c As String) As String
    If ((Left(c, 1) = "$") Or (Left(c, 1) = "t")) Then
        format_cost = c
        Else
        format_cost = Format(c, "$#0")
        End If
End Function
Sub add_club(c As String, S As String)
    If (c <> mrAssoc) Then
        If (clubs = "") Then
            clubs = c
            Else
            clubs = clubs + S + c
            End If
        End If
End Sub

Sub Create_TXT_File()
'
'
'   Open output file
'
    Dim r As Integer
    Dim eventID As String
    Dim schEN As String
    Dim prevEN As String
    Dim hostID As String
    Dim ans As Integer
    Dim msg As String
    Dim AClub As String
    Dim Connector As String
    Dim initialFileName As String
    Dim tempPlayerLmit As Integer
            
    ActiveWorkbook.Worksheets("Schedule").Activate
    Set SE = ActiveSheet
    initialFileName = path + scheduleFileName
    
    fileSaveName = Application.GetSaveAsFilename(initialFileName, "Text Files (*.txt), *.txt", 1, initialFileName)
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If (fileSaveName <> False) Then             ' Check is cancel not selected
        hFilename = fileSaveName                ' Save Selected
        If (fs.fileExists(hFilename) = True) Then
            fs.DeleteFile (hFilename)
            End If
        Set a = fs.createTextfile(hFilename, True)
    
    
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set a = fs.createTextfile(scheduleFileName, True)
    prevDate = lowDate                          ' pick up low date as previous date to start control loop
    done = False
    clubs = ""
    
    evDateTxt = ""
    costTxt = ""
    timeTxt = ""
    deadline = ""
    quoteTxt = ""
    eolTxt = ""     ' End of Line Text
    playerLimitTxt = ""
    tempPlayerLmit = 0
    Connector = ""
    
    lowID = "00000000000000"
    prevEN = lowID
    r = 2                                       ' start loop at first entry in schedule worksheet
    Do While (SE.Cells(r, S_ID).Value <> "")       ' Event ID is non-blank
        schDate = SE.Cells(r, S_Date).Value
        schPlace = SE.Cells(r, S_AH).Value
        schEN = SE.Cells(r, S_ID).Value           ' Get currrent EventID
        
        If (SE.Cells(r, S_AH).Value <> openDate) Then     ' skip open home dates
            If (prevEN < schEN) Then                                ' change in Event Number?
                If (prevEN <> lowID) Then            ' first time through
                    a.WriteLine (Close_line)
                    End If
                clubs = ""
                AClub = ""
                eolTxt = ""
                quoteTxt = Chr(34)
                a.write ("")                 ' and start a new record
                prevEN = schEN
                eventID = schEN
                
                a.write (eventID)
                a.write (delimiter)
                a.write (Format(schDate, "ddddd ttttt")) '

                a.write (delimiter)
                a.write (LTrim(SE.Cells(r, S_AH).Value)) '

                a.write (delimiter)
                costTxt = format_cost(SE.Cells(r, S_Cost).Value)
                timeTxt = Format(SE.Cells(r, 3).Value, "h:mm a/p")

                deadline = Format(SE.Cells(r, S_Deadline).Value, "ddddd") + " 12:00:00 PM"
                postDate = Format(SE.Cells(r, S_Post).Value, "ddddd") + " 9:00:00 AM"
                phoneTxt = SE.Cells(r, S_Phone).Value
                guestTxt = SE.Cells(r, S_Guest).Value
                specialRuleTxt = SE.Cells(r, S_SRule).Value
                tempPlayerLimit = SE.Cells(r, S_PLimit).Value
                If (tempPlayerLimit = 0) Then
                    If (SE.Cells(r, S_AH).Value = home) Then
                        tempPlayerLimit = homePlayerLimit
                        Else
                        tempPlayerLimit = awayPlayerLimit
                        End If
                    End If
                playerLimitTxt = Format(tempPlayerLimit, "##0")
                    
                End If                  ' prevEN < schEN

            If (SE.Cells(r, 6).Value = ryderCup) Then
                If (SE.Cells(r, S_AH).Value = home) Then
                    Connector = " vs "
                    End If
                If (SE.Cells(r, S_AH).Value = away) Then
                    Connector = " at "
                    End If
                Call add_club(ryderCup + Connector, "")
                Call add_club(SE.Cells(r, S_Club).Value, "")
                Connector = " "
                Else
                
                If (SE.Cells(r, S_AH).Value = home) Then
                    If (SE.Cells(r, S_Event).Value = interClub) Then
                        Call add_club(SE.Cells(r, S_Club).Value, ", ")
                        Else
                        Call add_club(SE.Cells(r, S_Event).Value, ", ")
                        End If
                    End If
                If (SE.Cells(r, S_AH).Value = "") Then
                    Call add_club(SE.Cells(r, S_Club).Value, ", ")
                    End If
                If (SE.Cells(r, S_AH).Value = away) Then
                    AClub = SE.Cells(r, S_Club).Value
                    If (SE.Cells(r, S_Status).Value = "T") Then
                        AClub = "<span style=""color: red"">" + AClub + " **TENTATIVE**</span> "
                        End If
                    Call add_club(AClub, ", ")
                    End If
                If (UCase(SE.Cells(r, 4).Value) = MISGA) Then
                    Call add_club(SE.Cells(r, S_Event).Value, " at ")
                    Call add_club(SE.Cells(r, S_Club).Value, " at ")
                    End If
                If (SE.Cells(r, S_AH).Value = club) Then
                    Call add_club(SE.Cells(r, S_Event).Value, "")
                    End If
                End If
            End If                      ' for skip Open Dates
        r = r + 1                       ' do the next row
        Loop
        a.WriteLine (Close_line)
        a.Close
        msg = hFilename + " created successfully."
    Else            ' Cancel Selected
        msg = "Save of " + scheduleFileName + " Schedule file cancelled"
    End If
    ans = MsgBox(msg, vbInformation, "Schedule Preparation")

    
    
End Sub


