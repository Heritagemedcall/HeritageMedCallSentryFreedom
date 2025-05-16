VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExceptionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' for externs, need report Item to show message received.

Option Base 1

Dim mNumEntries          As Long
Attribute mNumEntries.VB_VarUserMemId = 1141047297
Dim mSumElapsed          As Long
Attribute mSumElapsed.VB_VarUserMemId = 1141047298



' new to Exception report (not in auto reports)
Public ResponseTime      As Long  ' seconds for responding to alarm
Public ResponseIsACK     As Integer  ' non-zero if ACK is response time
Public ReportType        As Integer  ' 0 is exception, 1 is Alarm Count

Public ReportID          As Long
Public Disabled          As Integer  '1 = Disabled, 0=enabled
Public ReportName        As String
Public Comment           As String

Public Rooms             As Collection  ' what rooms to include
Public Events            As Collection  ' what rooms to include
Public DevTypes          As Collection  ' WAS Events in Auto reports, but is DeviceTypes

Public RoomString        As String  ' what rooms to include - delimited list
Public EventString       As String  ' what events for this report  - delimited list
Public DevString         As String  ' what events for this report  - delimited list

Public TimePeriod        As Long  ' Not Used From Dropdown list

Public DayPeriod         As Long  ' days or monthly
Public DAYS              As Long  ' as bitfield
Public DayString         As String  ' as string

Public Shift             As Integer  ' 0=all, 1=first, 2=second, 3= third) ( 3rd not implemented)

Public DayPartStart      As Long  ' if a specific time of day is requested (start and end = 0 ... all day)
Public DayPartEnd        As Long

Public SortOrder         As Long  ' 0 = Chron, 1=long-to-short, 2=By-Room

Public SendHour          As Long  '  NOT USED - 0 to 2400 and end of event ' when to actually create the report

Public EventItems        As Collection

Public SaveAsFile        As Integer  ' always for now
Public SendAsEmail       As Integer

Public recipient         As String  ' email address
Public Subject           As String  ' email subject

Public FileFormat        As Integer  ' TAB, TAB no header, HTML
Public DestFolder        As String  ' where to keep files

Public NextReportDue     As Date  ' when we're up



Private mBusy            As Boolean  ' revent reentrancy

Private Const COL_SERIAL = 0
Private Const COL_ROOM = 0.75
Private Const COL_RESIDENT = 1.5
Private Const COL_EVENT = 3.5
Private Const COL_ALARMTIME = 6.5
Private Const COL_CANCELTIME = 7.5
Private Const COL_ELAPSEDTIME = 8.5
Private Const COL_OPERATOR = 8.75

Public PeriodEnd         As Date
Public PeriodStart       As Date

Private AllRows          As Collection  ' all alarms as rows of report cAlarmHistory
Private rows             As Collection  ' each row chosen for this report

Public Function DoReport()
  '  If mBusy Then Exit Function
  '  mBusy = True

  PeriodEnd = NextReportDue


  NextReportDue = 0
  GenAutoReport
  '  mBusy = False
End Function
Private Function GenAutoReport()




  'Global Const AUTOREPORT_DAILY = 0
  'Global Const AUTOREPORT_SHIFT1 = 1
  'Global Const AUTOREPORT_SHIFT2 = 2
  'Global Const AUTOREPORT_WEEKLY = 7
  'Global Const AUTOREPORT_MONTHLY = 30
  '
  '
  '
  'Global Const AUTOREPORT_SORT_ROOM = 0
  'Global Const AUTOREPORT_SORT_ELAPSED = 1 ' in longest to shortest
  'Global Const AUTOREPORT_SORT_CHRONO = 2
  '
  'Global Const AUTOREPORTFORMAT_TAB = 0
  'Global Const AUTOREPORTFORMAT_TAB_NOHEADER = 1 ' in longest to shortest
  'Global Const AUTOREPORTFORMAT_HTML = 2


  Dim Path               As String
  Dim filename           As String
  Dim ext                As String
  Dim hfile              As Integer

  Dim rs                 As ADODB.Recordset
  Dim SQL                As String
  Dim WhereClause        As String

  
  Dim AlarmHistory       As cAlarmHistory
  Dim CurrentDate        As Date

  Dim rs2                As ADODB.Recordset
  Dim SQL2               As String
  Dim Whereclause2       As String

  Dim Room               As cRoom

  Dim TempEvents         As Collection
  Dim rc                 As Long

  Dim FileText           As String

  ' select all alarms that match criteria with basic info
  ' Walk each alarm
  ' Fetch each event for each alarm
  ' Initially only get initiating events, history comes later
  ' Alarmid = 0 if first event for alarm

  Dim j                  As Integer
  Dim Do_UnAssigned      As Boolean

  Debug.Print "report start " & Now

  Dim HasSecondShift     As Boolean
  Dim HasThirdShift      As Boolean


  If Configuration.EndFirst = Configuration.EndNight Then
    HasSecondShift = False
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = True
  Else
    HasSecondShift = False
    HasThirdShift = False
  End If


  Select Case DayPeriod
    Case AUTOREPORT_DAILY
      PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
      PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
      If PeriodStart >= PeriodEnd Then
        PeriodStart = DateAdd("d", -1, PeriodStart)
      End If

    Case AUTOREPORT_SHIFT1

      PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
      PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
      If PeriodStart >= PeriodEnd Then
        PeriodStart = DateAdd("d", -1, PeriodStart)
      End If

    Case AUTOREPORT_SHIFT2
      PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
      PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
      If PeriodStart >= PeriodEnd Then
        PeriodStart = DateAdd("d", -1, PeriodStart)
      End If

    Case AUTOREPORT_SHIFT3
      PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
      PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
      If PeriodStart >= PeriodEnd Then
        PeriodStart = DateAdd("d", -1, PeriodStart)
      End If



    Case AUTOREPORT_WEEKLY
      PeriodStart = DateAdd("ww", -1, PeriodEnd)

    Case AUTOREPORT_MONTHLY
      PeriodStart = DateAdd("m", -1, PeriodEnd)

  End Select



  'this gets all alarms in the time frame specified. We'll filter out what's not needed later

  WhereClause = " AND (EventDate >= " & DateDelim & PeriodStart & DateDelim & " ) AND  (EventDate < " & DateDelim & PeriodEnd & DateDelim & " )  ) "

  'SQl = "SELECT * FROM alarms WHERE (AlarmID = 0) AND (Devtypes in (" & DevString & ") " & _

   SQL = "SELECT  ID , Serial, EventDate, Alarm, Tamper, Battery, ResidentID, RoomID,  Announce, EventType, UserName, AlarmID, Residents.NameLast, Residents.NameFirst,Rooms.Room , Devices.Model   FROM ALARMS " & _
   " LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID " & _
   " LEFT JOIN Devices ON Alarms.Serial = Devices.Serial " & _
   " LEFT JOIN Rooms ON Alarms.RoomID = Rooms.RoomID " & _
   " WHERE (AlarmID = 0) AND (Devtypes in (" & DevString & ") " & _
   WhereClause & " ORDER BY EventDate"


  Debug.Print "SQL: " & SQL
  Dim Device             As cESDevice
  Dim CLSPTI             As Long

  Set rs = ConnExecute(SQL)
  
  Whereclause2 = "  WHERE (alarms.EventDate >= " & DateDelim & PeriodStart & DateDelim & ") And (alarms.EventDate < " & DateDelim & PeriodEnd & DateDelim & ") "

  SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as EventDate, Alarms_1.EventType as EventType, Alarms_1.info as info, Alarms_1.UserName as UserName " & _
         " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
         Whereclause2 & _
         " ORDER BY Alarms.ID; "

  Set rs2 = ConnExecute(SQL2)




  Set AllRows = New Collection
  Dim counter            As Long
  counter = 0
  Do Until rs.EOF
    If counter > 200 Then
      counter = 0
      DoEvents
    End If
    counter = counter + 1

    Set AlarmHistory = New cAlarmHistory

    AlarmHistory.ID = rs("ID")  ' this record's ID
    AlarmHistory.Serial = rs("Serial") & ""
    AlarmHistory.EventDate = rs("EventDate")
    AlarmHistory.alarm = rs("Alarm")
    AlarmHistory.Tamper = rs("Tamper")
    AlarmHistory.Battery = rs("Battery")
    AlarmHistory.Announce = rs("Announce") & ""
    AlarmHistory.EventType = rs("EventType")
    AlarmHistory.User = rs("Username") & ""
    AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm
    AlarmHistory.RoomID = rs("RoomID")
    ' use lookup table
    
    AlarmHistory.CLSPTI = Val(GetCLSPTI(rs("Model") & ""))

    If Len(rs("NameLast") & "") Then
      If Len(rs("NameFirst") & "") Then
        AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
      Else
        AlarmHistory.Resident = rs("NameLast") & ""
      End If
    ElseIf Len(rs("NameFirst") & "") Then
      AlarmHistory.Resident = rs("NameFirst") & ""
    End If
    AlarmHistory.Room = rs("room") & ""
    rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)
    If AlarmHistory.CLSPTI <> &HEEEE& Then
      AllRows.Add AlarmHistory
    End If
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing

  rs2.Close
  Set rs2 = Nothing


  Debug.Print "AllRows.Count: " & AllRows.Count
  Set rows = New Collection
  If AllRows.Count Then

    ' At this point all alarms are loaded with resident and room info

    Debug.Print "Rooms "; RoomString
    Debug.Print "Room Count "; Rooms.Count



    Select Case SortOrder
      Case AUTOREPORT_SORT_ELAPSED
        'Need to sort by elapsed time
        Set AllRows = ShellSortElapsedTime(AllRows, Descending:=True)

        For Each AlarmHistory In AllRows
          For Each Room In Rooms

            If counter > 200 Then
              counter = 0
              DoEvents
            End If
            counter = counter + 1
            If Room.RoomID = AlarmHistory.RoomID Then

              rows.Add AlarmHistory
            End If
          Next
        Next

      Case AUTOREPORT_SORT_CHRONO
        ' walk the alarms
        'pick up the rooms requested

        For Each AlarmHistory In AllRows
          For Each Room In Rooms

            If counter > 200 Then
              counter = 0
              DoEvents
            End If
            counter = counter + 1


            If Room.RoomID = AlarmHistory.RoomID Then

              rows.Add AlarmHistory
            End If
          Next
        Next


      Case Else                ' AUTOREPORT_SORT_ROOM
        ' walk the room list i
        ' walk the alarm list j
        ' DO Unassigned Last ?

        For Each Room In Rooms

          For Each AlarmHistory In AllRows
            If counter > 200 Then
              counter = 0
              DoEvents
            End If
            counter = counter + 1

            If Room.RoomID = AlarmHistory.RoomID Then

              rows.Add AlarmHistory
            End If
          Next
        Next


    End Select
    Debug.Print
    Debug.Print "Select Rows.Count: " & rows.Count

  End If
  Select Case FileFormat
    Case AUTOREPORTFORMAT_TAB_NOHEADER
      FileText = DoTabReport(rows, Headers:=False)
      ext = ".txt"
    Case AUTOREPORTFORMAT_HTML
      FileText = DoHTMLReport(rows)
      ext = ".html"
    Case Else                  '     Case AUTOREPORTFORMAT_TAB
      FileText = DoTabReport(rows, Headers:=True)
      ext = ".txt"
  End Select
  ' write file to dest folder

  hfile = FreeFile

  DestFolder = App.Path & "\AutoReports\"
  Path = DestFolder

  If Right(Path, 1) <> "\" Then
    Path = Path & "\"
  End If

  If Not DirExists(Path) Then
    On Error Resume Next
    MkDir Path
  End If

  If Not DirExists(Path) Then
    Path = App.Path & "\"
  End If

  filename = Path & CleanFileName(ReportName) & " " & Format$(PeriodEnd, "yyyymmdd hh") & ext

  Open filename For Output As #hfile
  Print #hfile, FileText
  Close #hfile

  'Debug.Print "FileText: " & FileText

  If (SendAsEmail And (Len(recipient) > 0)) Then

    Dim mapi             As Object
    Dim message          As String

    message = "Auto Report Attached: " & filename

    On Error Resume Next

    If (Configuration.UseSMTP = MAIL_SMTP) Then

      If gSMTPMailer Is Nothing Then
        '1070          Set gSMTPMailer = New SendMail
        Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
      End If
      If gSMTPMailer Is Nothing Then
        LogProgramError "Could not create SMTPMailer Object in cExceptionReport.GenAutoReport." & Erl
      Else
        Call gSMTPMailer.Send("", "", recipient, Subject, message, filename)
      End If


    Else
      Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
      If mapi Is Nothing Then
        LogProgramError "Could not create SENTRYMAIL Object in cExceptionReport.GenAutoReport." & Erl
      Else
        Call mapi.SendWithAttachments("", "", recipient, Subject, message, filename)
      End If

    End If


    '// Username, Password, Address, Subject,Body, AttachmentsList ' Attachemnet list is a semicolon ";" delimited list of file attachments
    'Call mapi.Send("", "", Configuration.AssurEmailRecipient, Configuration.AssurEmailSubject, Message)

    Set mapi = Nothing

  End If
  Debug.Print "report end " & Now

End Function

Function DoTabReport(rows As Collection, ByVal Headers As Boolean) As String
  'Private Const COL_SERIAL = 0
  'Private Const COL_ROOM = 0.75
  'Private Const COL_RESIDENT = 1.5
  'Private Const COL_EVENT = 3.5
  'Private Const COL_ALARMTIME = 6.5
  'Private Const COL_CANCELTIME = 7.5
  'Private Const COL_ELAPSEDTIME = 8.5
  'Private Const COL_OPERATOR = 8.75

  Dim row                As cAlarmHistory
  Dim text               As String
  Dim counter            As Long


  If (Headers) Then

    text = Join(Array("Device", "Room", "Resident", "Event", "Announce", "Alarm Time", "Canceled", "Elapsed", "Operator"), vbTab) & vbCrLf
  End If

  For Each row In rows
    If counter > 200 Then
      counter = 0
      DoEvents
    End If
    counter = counter + 1

    text = text & row.ToTABBEDString & vbCrLf  '   Join(Array(row.Serial, row.room, row.Resident, row.EventName, row.AlarmTimeToString, row.CancelTimeToString, row.user), vbTab) & vbCrLf
  Next

  If ReportType = 0 Then

    Dim Average          As Long
    Dim AverageTime      As String
    If mNumEntries > 0 Then
      Average = mSumElapsed / mNumEntries
    End If
    If Average > 0 Then
      AverageTime = Format(DateAdd("s", Average, 0), "hh:nn:ss")
    Else
      AverageTime = "N/A"
    End If

    text = text & Join(Array("Average", "", "", "", "", AverageTime, ""), vbTab) & vbCrLf  '   Join(Array(row.Serial, row.room, row.Resident, row.EventName, row.AlarmTimeToString, row.CancelTimeToString, row.user), vbTab) & vbCrLf

  End If

  DoTabReport = text

End Function

Function DoHTMLReport(rows As Collection) As String
  Dim row                As cAlarmHistory
  Dim odd                As Boolean
  Dim text               As String
  Dim counter            As Long
  Dim Average            As Long
  Dim AverageTime        As Long



  ' html header
  text = "<html>"
  text = text & "<head>"
  text = text + "<style type=""text/css"">"
  text = text + "body {width:900px; font-family:arial,verdana,sans-serif;}"
  text = text + "table.main {width:900px;font-size:1.0em;}"
  text = text + "tr.header td {background-color: #ADD8E6; color: black; margin:0px; padding:2px; font-weight:bold;}"
  text = text + "tr.even td {background-color: #FAFAD2; color: black; margin:0px; padding:2px}"
  text = text + "tr.odd td {background-color: white; color: black; margin:0px; padding:2px}"
  text = text + "h1 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.3em}"
  text = text + "h2 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.0em;}"
  text = text + "p.complete {background-color: white; color:gray;margin:5px;text-align:left;font-size:0.9em;}"

  text = text + "</style>"

  text = text & "</head>" & vbCrLf
  text = text & "<body>"

  text = text + "<h1>" & HTMLEncode(ReportName) & "</h1>"
  text = text + "<h1>" & HTMLEncode(Configuration.Facility) & "</h1>"
  text = text + "<h2>" & "Date: " & HTMLEncode(PeriodStart) & " to " & HTMLEncode(PeriodEnd) & "</h2>"
  text = text + "<br />"

  text = text & "<table class='main'>" & vbCrLf

  text = text + "<tr class='header'>" & vbCrLf
  text = text + "<td>Serial</td> <td>Room</td> <td>Event</td> <td>Announce</td> <td>Time</td> <td>Elapsed</td><td>Operator</td>" & vbCrLf
  text = text + "</tr>" & vbCrLf


  If rows.Count Then
    For Each row In rows
      If ReportType = 0 Then
        Select Case row.EventType
          Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

            If gElapsedEqACK = USE_ELAPSED_ACK Then
              If row.ACKElapsed > 0 Then
                mNumEntries = mNumEntries + 1
                mSumElapsed = mSumElapsed + row.ACKElapsed
              End If
            Else
              If row.ElapsedTime > 0 Then
                mNumEntries = mNumEntries + 1
                mSumElapsed = mSumElapsed + row.ElapsedTime
              End If
            End If
        End Select

      End If

      If counter > 200 Then
        counter = 0
        DoEvents
      End If
      counter = counter + 1



      text = text & "<!-- AlarmID " & row.ID & " RoomID " & row.RoomID & " ResID " & row.ResidentID & "-->" & vbCrLf
      text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
      text = text & "<td>" & row.Serial & "</td><td>" & HTMLEncode(row.Room) & "</td><td>" & _
      HTMLEncode(row.EventName) & "</td><td>" & HTMLEncode(row.Announce) & "</td><td>" & row.AlarmTimeToString & "</td><td>" & row.ElapsedTimeToString & "</td><td>" & HTMLEncode(row.User) & "</td>" & vbCrLf
      text = text & "</tr>" & vbCrLf
      odd = Not odd
    Next
  End If



  If ReportType = 0 Then

    If mNumEntries > 0 Then
      Average = mSumElapsed / mNumEntries
      AverageTime = Format(DateAdd("s", Average, 0), "hh:nn:ss")
    Else
      AverageTime = "N/A"
    End If

    text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
    text = text & "<td>Average</td><td>" & mNumEntries & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>" & AverageTime & "</td><td>&nbsp;</td>" & vbCrLf
    text = text & "</tr>" & vbCrLf
  End If

  text = text & "</table>" & vbCrLf
  text = text + "<p class='complete'>Report Complete</p>"
  text = text & "</body>" & vbCrLf
  text = text & "</html>"



  DoHTMLReport = text

End Function


'Public Function ShellSortElapsedTime(c As Collection, ByVal Descending As Boolean) As Collection
'
'    Dim Hold          As Long ' specific to sort
'    Dim Gap           As Long ' specific to sort
'    Dim i             As Long ' specific to sort
'    Dim Min           As Long ' always 1
'    Dim Max           As Long ' same as number of objects in collection
'    Dim Compare       As Long ' the current value
'    Dim IndexValue    As Long ' the current index
'    Dim TestValue     As Long ' to test against
'    Dim j             As Long
'    Dim sx()          As Long ' array for sorting of indexes
'
'    Dim hist As cAlarmHistory
'
'    Dim newcoll As Collection    ' the new collection
'
'    Set newcoll = New Collection
'
'    ReDim sx(c.Count, 2) As Long ' we need a count number of elements with 2 columns
'
'    For j = 1 To c.Count
'      Set hist = c(j)
'
'      sx(j, 1) = hist.ElapsedTime
'      sx(j, 2) = j
'    Next
'
''        Debug.Print
''    For j = 1 To c.Count
''      Debug.Print sx(j, 1), sx(j, 2)
''    Next
'
'    Min = 1
'    Max = c.Count
'    Gap = Min
'
'    Do    ' figureb optimum gap
'        Gap = 3 * Gap + 1
'    Loop Until Gap > Max
'
'    Dim t
'    t = Timer
'
'    Do
'        Gap = Gap \ 3
'        For i = Gap + Min To Max
'            DoEvents
'            Compare = sx(i, 1)
'            IndexValue = sx(i, 2)
'            Hold = i
'
'            If Descending Then
'
'              TestValue = sx(Hold - Gap, 1)
'              Do While TestValue < Compare
'                DoEvents
'                ' swap the value and the index
'                sx(Hold, 1) = sx(Hold - Gap, 1) ' swap real values
'                sx(Hold, 2) = sx(Hold - Gap, 2) ' swap indexes
'                Hold = Hold - Gap
'                If Hold < Min + Gap Then
'                  Exit Do
'                End If
'                TestValue = sx(Hold - Gap, 1)
'              Loop
'            Else ' ascending
'            TestValue = sx(Hold - Gap, 1)
'              Do While TestValue > Compare
'                DoEvents
'                ' swap the value and the index
'                sx(Hold, 1) = sx(Hold - Gap, 1)
'                sx(Hold, 2) = sx(Hold - Gap, 2)
'                Hold = Hold - Gap
'                If Hold < Min + Gap Then
'                  Exit Do
'                End If
'                TestValue = sx(Hold - Gap, 1)
'              Loop
'            End If
'
'            sx(Hold, 1) = Compare
'            sx(Hold, 2) = IndexValue
'        Next i
'
'    Loop Until Gap = 1
'
'
'
'    'Debug.Print "time " & Timer - t
'
'    For j = 1 To Max
'     ' Debug.Print sx(j, 1), sx(j, 2)
'      newcoll.Add c(sx(j, 2))
'    Next
'    Set c = Nothing
'    Set ShellSortElapsedTime = newcoll
'
'
'End Function

'Private Function GetAlarmHistory(rs As Recordset) As cAlarmHistory
'  Dim History            As cAlarmHistory
'
'  Set History = New cAlarmHistory
'  History.Parse rs
'  If History.ResidentID <> 0 Then
'    History.Resident = GetResidentName(History.ResidentID)
'  End If
'  History.Room = GetRoomName(History.RoomID)
'  History.GetHistory conn
'  Set GetAlarmHistory = History
'
'
'End Function



Public Property Get due() As Boolean
  If mBusy Then Exit Property
  Dim timeleft           As Double
  If 0 = NextReportDue Then    ' first go-round
    NextReportDue = GetNextReportDate()
  End If

  If NextReportDue <> 0 Then
    timeleft = DateDiff("s", Now, NextReportDue)
    'Debug.Print "Next AutoReport Due '" & ReportName & "' " & Format(NextReportDue, "mm/dd/yy h:n:s") & " in " & timeleft \ 60 & ":" & timeleft Mod 60 & " Min:Secs"

    due = (Now > NextReportDue)
  End If

End Property
Public Function GetNextReportDate() As Date
  Dim Makedate           As Date
  Dim CurrentDate        As Date
  Dim CurrentDOW         As Integer
  Dim WeekDayNumber      As Integer

  Dim WeekDays(0 To 7)   As Integer
  Dim ReportDays(0 To 7) As Integer
  Dim DailyReports(0 To 7) As Date

  Dim j                  As Long


  CurrentDate = Now

  GetNextReportDate = 0

  Erase WeekDays()
  Erase ReportDays()
  Erase DailyReports()


  Select Case DayPeriod
    Case AUTOREPORT_DAILY

      If (NextReportDue = 0) Then
        Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
        If DayPartEnd = 0 Then
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        Else
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
        WeekDayNumber = Weekday(Now) - 1
        For j = 0 To 6
          If ((2 ^ j) And DAYS) <> 0 Then

            ReportDays(j) = 1
            DailyReports(j) = DateAdd("d", j - 1, Makedate)  '
            If DayPartEnd = 0 Then
              DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
            End If
            If DailyReports(j) < CurrentDate Then
              DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
            End If
          End If
        Next


      End If

      SortDates DailyReports()
      For j = 0 To 7
        If DailyReports(j) > Now Then
          GetNextReportDate = DailyReports(j)
          Exit For
        End If
      Next


    Case AUTOREPORT_SHIFT1
      If (NextReportDue = 0) Then
        Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
        If DayPartEnd = 0 Then
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        Else
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
        WeekDayNumber = Weekday(Now) - 1
        For j = 0 To 6
          If ((2 ^ j) And DAYS) <> 0 Then

            ReportDays(j) = 1
            DailyReports(j) = DateAdd("d", j - 1, Makedate)  '
            If DayPartEnd = 0 Then
              DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
            End If
            If DailyReports(j) < CurrentDate Then
              DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
            End If
          End If
        Next


      End If

      SortDates DailyReports()
      For j = 0 To 7
        If DailyReports(j) > Now Then
          GetNextReportDate = DailyReports(j)
          Exit For
        End If
      Next
    Case AUTOREPORT_SHIFT2
      If (NextReportDue = 0) Then
        Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
        If DayPartEnd = 0 Then
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        Else
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
        WeekDayNumber = Weekday(Now) - 1
        For j = 0 To 6
          If ((2 ^ j) And DAYS) <> 0 Then

            ReportDays(j) = 1
            DailyReports(j) = DateAdd("d", j - 1, Makedate)  '
            If DayPartEnd = 0 Then
              DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
            End If
            If DailyReports(j) < CurrentDate Then
              DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
            End If
          End If
        Next


      End If

      SortDates DailyReports()
      For j = 0 To 7
        If DailyReports(j) > Now Then
          GetNextReportDate = DailyReports(j)
          Exit For
        End If
      Next
    Case AUTOREPORT_SHIFT3
      If (NextReportDue = 0) Then
        Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
        If DayPartEnd = 0 Then
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        Else
          Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
        End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
        WeekDayNumber = Weekday(Now) - 1
        For j = 0 To 6
          If ((2 ^ j) And DAYS) <> 0 Then

            ReportDays(j) = 1
            DailyReports(j) = DateAdd("d", j - 1, Makedate)  '
            If DayPartEnd = 0 Then
              DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
            End If
            If DailyReports(j) < CurrentDate Then
              DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
            End If
          End If
        Next


      End If

      SortDates DailyReports()
      For j = 0 To 7
        If DailyReports(j) > Now Then
          GetNextReportDate = DailyReports(j)
          Exit For
        End If
      Next

    Case AUTOREPORT_WEEKLY     ' once a week at end of week
      If (NextReportDue = 0) Then
        Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
        CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)  ' returns 1 for sunday
        Makedate = DateAdd("d", 8 - CurrentDOW, Makedate)
        GetNextReportDate = Makedate
      End If

    Case AUTOREPORT_MONTHLY    ' once a month, at end of month
      If (NextReportDue = 0) Then
        Makedate = DateSerial(Year(Now), Month(Now), 1)  ' make first of current month
        Makedate = DateAdd("m", 1, Makedate)  ' add a month
        GetNextReportDate = Makedate  ' DateAdd("n", -1, Makedate)  ' ?? minus
      End If
    Case Else

  End Select




End Function



Function Parse(rs As ADODB.Recordset) As Boolean
  Dim j                  As Integer
  Dim Room               As cRoom
  Dim dataWrapper        As cDataWrapper
  Dim Items()            As String


10 ResponseTime = Val(rs("ResponseTime") & "")  ' seconds for responding to alarm
20 ResponseIsACK = Val(rs("ResponseIsACK") & "")  ' non-zero if ACK is response time


30 ReportID = rs("reportid")
40 Disabled = rs("Disabled")
50 ReportName = rs("reportname") & ""
60 Comment = rs("Comment") & ""
70 RoomString = rs("Rooms") & ""
80 EventString = rs("Events") & ""
90 DevString = rs("DevTypes") & ""
100 TimePeriod = rs("TimePeriod")
110 DayPeriod = rs("DayPeriod")
120 DAYS = rs("Days")
130 Shift = rs("Shift")
140 DayPartStart = rs("DayPartStart")
150 DayPartEnd = rs("DayPartend")
160 SortOrder = rs("SortOrder")
170 SendHour = rs("SendHour")
180 SaveAsFile = 1             '  rs("SaveAsFile")
190 SendAsEmail = rs("SendAsEmail")
200 recipient = rs("Recipient") & ""
210 Subject = rs("Subject") & ""
220 FileFormat = rs("FileFormat")
230 DestFolder = rs("DestFolder") & ""

240 ResponseTime = Val(rs("ResponseTime") & "")
250 ReportType = Val(rs("Reporttype") & "")

260 Set Rooms = New Collection


270 Set DevTypes = New Collection

280 Items = Split(DevString, vbTab)
290 For j = LBound(Items) To UBound(Items)
300 Set dataWrapper = New cDataWrapper
310 dataWrapper.LongValue = Val(Items(j))
320 DevTypes.Add dataWrapper
330 Next

  Dim HasSecondShift     As Boolean
  Dim HasThirdShift      As Boolean



340 If Configuration.EndFirst = Configuration.EndNight Then
350 HasSecondShift = False
360 HasThirdShift = False
370 ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
380 HasSecondShift = True
390 HasThirdShift = False
400 ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
410 HasSecondShift = True
420 HasThirdShift = True
430 Else
440 HasSecondShift = False
450 HasThirdShift = False
460 End If

470 Select Case DayPeriod
    Case AUTOREPORT_SHIFT1
480   If HasThirdShift Then
490     DayPartStart = Configuration.EndThird
500     DayPartEnd = Configuration.EndFirst

510   Else
520     DayPartStart = Configuration.EndNight
530     DayPartEnd = Configuration.EndFirst

540   End If
550 Case AUTOREPORT_SHIFT2

560   DayPartStart = Configuration.EndFirst
570   DayPartEnd = Configuration.EndNight
580 Case AUTOREPORT_SHIFT3
590   DayPartStart = Configuration.EndNight
600   DayPartEnd = Configuration.EndThird

610 End Select



End Function

Function UpdateData(ByVal rs As ADODB.Recordset) As Boolean

  ' convert Rooms collection to string of IDs

  Dim j                  As Integer
  Dim Room               As cRoom
  Dim Items()            As String
  Dim dataWrapper        As cDataWrapper

  RoomString = ""
  If Rooms.Count Then
    ReDim Items(Rooms.Count)
    For j = 1 To Rooms.Count
      Set Room = Rooms(j)
      Items(j) = CStr(Room.RoomID)
    Next
    RoomString = Join(Items, vbTab)
  End If

  DevString = ""
  If DevTypes.Count Then
    ReDim Items(DevTypes.Count)
    For j = 1 To DevTypes.Count
      Set dataWrapper = DevTypes(j)
      Items(j) = dataWrapper.LongValue
    Next
    DevString = Join(Items, vbTab)  ' tab separated for greater compat
  End If

  ' auto increment field:  rs("reportid") = ReportID

  ' new with exceptions
  rs("ResponseTime") = ResponseTime  ' seconds for responding to alarm
  rs("ResponseIsACK") = ResponseIsACK  ' non-zero if ACK is response time
  rs("Reporttype") = ReportType

  ' old

  rs("Disabled") = Disabled
  rs("reportname") = ReportName
  rs("Comment") = Comment
  rs("Rooms") = RoomString
  rs("Events") = EventString
  rs("DevTypes") = DevString
  rs("TimePeriod") = TimePeriod
  rs("DayPeriod") = DayPeriod
  rs("Days") = DAYS
  rs("Shift") = Shift
  rs("DayPartStart") = DayPartStart
  rs("DayPartend") = DayPartEnd
  rs("SortOrder") = SortOrder
  rs("SendHour") = SendHour
  rs("SaveAsFile") = 1         '
  rs("SendAsEmail") = SendAsEmail
  rs("Recipient") = recipient
  rs("Subject") = Subject
  rs("FileFormat") = FileFormat
  rs("DestFolder") = DestFolder

  'Actual call to rs.Update is in calling routine

End Function

Public Sub Finalize()

  Set Rooms = Nothing
  Set Events = Nothing
  Set EventItems = Nothing
End Sub


Private Sub Class_Initialize()
  ReportName = "New Report"
  Subject = "Exception Report"
  Set Rooms = New Collection
  Set Events = New Collection
  Set DevTypes = New Collection
  Set EventItems = New Collection
End Sub

Private Sub Class_Terminate()
  Finalize
End Sub

Public Property Get Busy() As Boolean  ' Public read only
  Busy = mBusy
End Property




