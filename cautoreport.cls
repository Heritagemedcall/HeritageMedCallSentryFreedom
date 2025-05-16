VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' for externs, need report Item to show message received.

Option Base 1

Public ReportID          As Long
Public Disabled          As Integer  '1 = Disabled, 0=enabled
Public ReportName        As String
Public Comment           As String

Public Rooms             As Collection  ' what rooms to include
Public Events            As Collection  ' what events for this report

Public RoomString        As String  ' what rooms to include - delimited list
Public EventString       As String  ' what events for this report  - delimited list

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

Private mNextReportDue   As Date  ' when we're up

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


Private mNumEntries      As Long
Private mSumElapsed      As Double
Private mAverageElapsed  As Double
Private mAverageElapsedString As String

Private mSumFinalized    As Double
Private mNumFinalized    As Double
Private mAverageFinalized As Double
Private mAverageFinalizedString As String

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

  Dim SQL                As String
  Dim WhereClause        As String

  Dim rs                 As ADODB.Recordset
  Dim AlarmHistory       As cAlarmHistory
  Dim CurrentDate        As Date

  Dim BasicEventTypes()  As String
  Dim filterstring       As String
  Dim i                  As Long
  Dim counter            As Long

  Dim SQL2               As String
  Dim Whereclause2       As String
  Dim rs2                As ADODB.Recordset

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

10 Debug.Print "report start " & Now

20 Select Case DayPeriod
    Case AUTOREPORT_DAILY
30    PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
40    PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
50    If PeriodStart >= PeriodEnd Then
60      PeriodStart = DateAdd("d", -1, PeriodStart)
70    End If

80  Case AUTOREPORT_SHIFT1
90    PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
100   PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
110   If PeriodStart >= PeriodEnd Then
120     PeriodStart = DateAdd("d", -1, PeriodStart)
130   End If

140 Case AUTOREPORT_SHIFT2
150   PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
160   PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
170   If PeriodStart >= PeriodEnd Then
180     PeriodStart = DateAdd("d", -1, PeriodStart)
190   End If


200 Case AUTOREPORT_SHIFT3
210   PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
220   PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
230   If PeriodStart >= PeriodEnd Then
240     PeriodStart = DateAdd("d", -1, PeriodStart)
250   End If


260 Case AUTOREPORT_WEEKLY
270   PeriodStart = DateAdd("ww", -1, PeriodEnd)

280 Case AUTOREPORT_MONTHLY
290   PeriodStart = DateAdd("m", -1, PeriodEnd)

300 End Select

  'this gets all alarms in the time frame specified. We'll filter out what's not needed later

310 WhereClause = " AND (Alarms.EventDate >= " & DateDelim & PeriodStart & DateDelim & " ) AND  (Alarms.EventDate < " & DateDelim & PeriodEnd & DateDelim & " )  "

320 SQL = "SELECT Alarms.ID, Alarms.roomid, Alarms.Serial, Alarms.EventDate, Alarms.Alarm, Alarms.Tamper, Alarms.Battery, Alarms.Announce, Alarms.EventType, Alarms.UserName, Alarms.AlarmID, Residents.NameLast, Residents.NameFirst, Rooms.Room ,Devices.Model " & _
          " FROM Alarms left JOIN Residents ON Alarms.ResidentID = Residents.ResidentID  left JOIN Rooms ON Alarms.RoomID = Rooms.RoomID   LEFT JOIN Devices ON Alarms.Serial = Devices.Serial   WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"


  '' First Query
330 Set rs = ConnExecute(SQL)

  'Set rs = FireHoseRecordSet(SQL)

340 BasicEventTypes = Split(EventString, ",")

350 For i = LBound(BasicEventTypes) To UBound(BasicEventTypes)
360 BasicEventTypes(i) = "eventtype = " & BasicEventTypes(i)
370 Next

380 filterstring = Join(BasicEventTypes, " or ")
390 rs.Filter = filterstring

400 Whereclause2 = " WHERE (Alarms.EventDate >= " & DateDelim & PeriodStart & DateDelim & " ) AND  (Alarms.EventDate < " & DateDelim & PeriodEnd & DateDelim & " )  "

410 SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as EventDate, Alarms_1.EventType as eventtype,Alarms_1.info as info, Alarms_1.UserName as username " & _
           " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
           Whereclause2 & _
           " ORDER BY Alarms_1.AlarmID, Alarms.ID; "

  '' Second Query
420 Set rs2 = ConnExecute(SQL2)
  'Set rs2 = FireHoseRecordSet(SQL2)

430 Set AllRows = New Collection

440 Do Until rs.EOF
    mNumEntries = mNumEntries + 1
450 counter = counter + 1
460 If counter > 500 Then
470   counter = 0
480   rc = DoEvents()
490 End If
500 Set AlarmHistory = New cAlarmHistory

510 AlarmHistory.ID = rs("ID")  ' this record's ID

520 AlarmHistory.Serial = rs("Serial") & ""

530 AlarmHistory.EventDate = rs("EventDate")
540 AlarmHistory.alarm = rs("Alarm")
550 AlarmHistory.Tamper = rs("Tamper")
560 AlarmHistory.Battery = rs("Battery")
570 AlarmHistory.Announce = rs("Announce") & ""
580 AlarmHistory.EventType = rs("EventType")
590 AlarmHistory.User = rs("Username") & ""
600 AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm
    AlarmHistory.RoomID = rs("RoomID")


    'Dim device As cESDevice
    'Set device = Devices.device(AlarmHistory.Serial)

    'If Not device Is Nothing Then

    ' use lookup table
    AlarmHistory.CLSPTI = Val(GetCLSPTI(rs("Model") & ""))
    'End If




610 If Len(rs("NameLast") & "") Then
620   If Len(rs("NameFirst") & "") Then
630     AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
640   Else
650     AlarmHistory.Resident = rs("NameLast") & ""
660   End If
670 ElseIf Len(rs("NameFirst") & "") Then
680   AlarmHistory.Resident = rs("NameFirst") & ""
690 End If
700 AlarmHistory.Room = rs("room") & ""
710 rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

720 AllRows.Add AlarmHistory
730 rs.MoveNext
740 Loop
  rs2.Close
  Set rs2 = Nothing
750 rs.Close
760 Set rs = Nothing

770 'Debug.Print "AllRows.Count: " & AllRows.Count
780 Set rows = New Collection
790 If AllRows.Count Then

    ' At this point all alarms are loaded with resident and room info

800 'Debug.Print "Rooms "; RoomString
810 'Debug.Print "Room Count "; Rooms.Count


820 counter = 0

830 Select Case SortOrder
      Case AUTOREPORT_SORT_ELAPSED
        'Need to sort by elapsed time
840     Set AllRows = ShellSortElapsedTime(AllRows, Descending:=True)

850     For Each AlarmHistory In AllRows
860       For Each Room In Rooms
870         If Room.RoomID = AlarmHistory.RoomID Then
880           rows.Add AlarmHistory
890         End If
900       Next
910     Next

920   Case AUTOREPORT_SORT_CHRONO
        ' walk the alarms
        'pick up the rooms requested
930     For Each AlarmHistory In AllRows
940       For Each Room In Rooms
            counter = counter + 1
950         If counter > 500 Then
960           counter = 0
970           rc = DoEvents()
980         End If

990         If Room.RoomID = AlarmHistory.RoomID Then
1000          rows.Add AlarmHistory
1010        End If
1020      Next
1030    Next


1040  Case Else                ' AUTOREPORT_SORT_ROOM
        ' walk the room list i
        ' walk the alarm list j
        ' DO Unassigned Last ?

1050    For Each Room In Rooms

1060      For Each AlarmHistory In AllRows
            counter = counter + 1
1070        If counter > 500 Then
1080          counter = 0
1090          rc = DoEvents()
1100        End If
1110        If Room.RoomID = AlarmHistory.RoomID Then
1120          rows.Add AlarmHistory
1130        End If
1140      Next
1150    Next


1160 End Select
1170 Debug.Print
1180 Debug.Print "Select Rows.Count: " & rows.Count

1190 End If
1200 Select Case FileFormat
    Case AUTOREPORTFORMAT_TAB_NOHEADER
1210  FileText = DoTabReport(rows, Headers:=False)
1220  ext = ".txt"
1230 Case AUTOREPORTFORMAT_HTML
1240  FileText = DoHTMLReport(rows)
1250  ext = ".html"
1260 Case Else                 '     Case AUTOREPORTFORMAT_TAB
1270  FileText = DoTabReport(rows, Headers:=True)
1280  ext = ".txt"
1290 End Select
  ' write file to dest folder

1300 hfile = FreeFile

1310 DestFolder = App.Path & "\AutoReports\"
1320 Path = DestFolder

1330 If Right(Path, 1) <> "\" Then
1340 Path = Path & "\"
1350 End If

1360 If Not DirExists(Path) Then
1370 On Error Resume Next
1380 MkDir Path
1390 End If

1400 If Not DirExists(Path) Then
1410 Path = App.Path & "\"
1420 End If


1430 filename = Path & CleanFileName(ReportName) & " " & Format$(PeriodEnd, "yyyymmdd hh") & ext

1440 Open filename For Output As #hfile
1450 Print #hfile, FileText
1460 Close #hfile

  'Debug.Print "FileText: " & FileText

1470 If (SendAsEmail And (Len(recipient) > 0)) Then

    Dim mapi             As Object
    Dim message          As String

1480 message = "Auto Report Attached: " & filename

1490 On Error Resume Next

1500 If (Configuration.UseSMTP = MAIL_SMTP) Then

1510  If gSMTPMailer Is Nothing Then
        '1070          Set gSMTPMailer = New SendMail
1520    Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
1530  End If
1540  If gSMTPMailer Is Nothing Then
1550    LogProgramError "Could not create SMTPMailer Object in cAutoReport.GenAutoReport." & Erl
1560  Else
1570    Call gSMTPMailer.Send("", "", recipient, Subject, message, filename)
1580  End If


1590 Else
1600  Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
1610  If mapi Is Nothing Then
1620    LogProgramError "Could not create SENTRYMAIL Object in cAutoReport.GenAutoReport." & Erl
1630  Else
1640    Call mapi.SendWithAttachments("", "", recipient, Subject, message, filename)
1650  End If
1660 End If

1670 Set mapi = Nothing

1680 End If
1690 Debug.Print "report end " & Now

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

  mAverageElapsed = 0
  mNumEntries = 0
  mSumElapsed = 0

  For Each row In rows
    counter = counter + 1
    If counter > 200 Then
      counter = 0
      DoEvents
    End If

    Select Case row.EventType

      Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
        If row.CLSPTI <> &HEEE& Then
          If gElapsedEqACK = 1 Then
            If row.ACKElapsed > 0 Then
              mSumElapsed = mSumElapsed + row.ACKElapsed
              mNumEntries = mNumEntries + 1
            End If
          ElseIf row.ElapsedTime > 0 Then
            mSumElapsed = mSumElapsed + row.ElapsedTime
            mNumEntries = mNumEntries + 1
          End If

          If row.FinalizeElapsed > 0 Then
            mSumFinalized = mSumFinalized + row.FinalizeElapsed
            mNumFinalized = mNumFinalized + 1

          End If



        End If

    End Select




    '      mNumEntries = mNumEntries + 1
    '      mSumElapsed = mSumElapsed + row.ElapsedTime

    text = text & row.ToTABBEDString & vbCrLf  '   Join(Array(row.Serial, row.room, row.Resident, row.EventName, row.AlarmTimeToString, row.CancelTimeToString, row.user), vbTab) & vbCrLf
  Next
  If mNumEntries Then
    mAverageElapsed = mSumElapsed / mNumEntries
  End If
  mAverageElapsedString = "N/A"
  If (mAverageElapsed > 0) Then
    mAverageElapsedString = Format$(DateAdd("s", mAverageElapsed, 0), "hh:nn:ss")
  End If

  If (mAverageFinalized > 0) Then
    mAverageFinalizedString = Format$(DateAdd("s", mAverageFinalized, 0), "hh:nn:ss")
  End If


  text = text & Join(Array("Average", "", "", "", "", "", "", mAverageElapsedString, "", mAverageFinalizedString), vbTab) & vbCrLf

  DoTabReport = text

End Function

Function DoHTMLReport(rows As Collection) As String
        Dim row                As cAlarmHistory
        Dim odd                As Boolean
        Dim text               As String
        Dim counter            As Long
        ' html header
10      text = "<html>"
20      text = text & "<head>"
30      text = text + "<style type=""text/css"">"
40      text = text + "body {width:900px; font-family:arial,verdana,sans-serif;}"
50      text = text + "table.main {width:900px;font-size:1.0em;}"
60      text = text + "tr.header td {background-color: #ADD8E6; color: black; margin:0px; padding:2px; font-weight:bold;}"  'Light blue
70      text = text + "tr.even td {background-color: #FAFAD2; color: black; margin:0px; padding:2px}"  ' Light Goldenrod Yellow
80      text = text + "tr.odd td {background-color: white; color: black; margin:0px; padding:2px}"  ' white
90      text = text + "tr.average td {background-color: #d3d3d3; color: black; margin:0px; padding:2px}"
100     text = text + "h1 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.3em}"
110     text = text + "h2 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.0em;}"
120     text = text + "p.complete {background-color: white; color:gray;margin:5px;text-align:left;font-size:0.9em;}"

130     text = text + "</style>"

140     text = text & "</head>" & vbCrLf
150     text = text & "<body>"

160     text = text + "<h1>" & HTMLEncode(ReportName) & "</h1>"
170     text = text + "<h1>" & HTMLEncode(Configuration.Facility) & "</h1>"
180     text = text + "<h2>" & "Date: " & HTMLEncode(PeriodStart) & " to " & HTMLEncode(PeriodEnd) & "</h2>"
190     text = text + "<br />"

200     text = text & "<table class='main'>" & vbCrLf

210     text = text + "<tr class='header'>" & vbCrLf
220     text = text + "<td>Room</td><td>Resident</td><td>Event</td><td>Announce</td><td>Alarm Time</td><td>Elapsed</td><td>Response By</td><td>Finalized</td><td>Resolution</td>" & vbCrLf
230     text = text + "</tr>" & vbCrLf


240     mAverageElapsed = 0
250     mNumEntries = 0
260     mSumElapsed = 0
270     mSumFinalized = 0

280     If rows.Count Then
290       For Each row In rows
300         counter = counter + 1
310         If counter > 200 Then
320           counter = 0
330           DoEvents
340         End If

350         Debug.Print "Row " & row.FinalizeElapsed


360         Select Case row.EventType

              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
370             If row.CLSPTI <> &HEEE& Then
380               If gElapsedEqACK = 1 Then
390                 If row.ACKElapsed > 0 Then
400                   mSumElapsed = mSumElapsed + row.ACKElapsed
410                   mNumEntries = mNumEntries + 1
420                 End If
430               ElseIf row.ElapsedTime > 0 Then
440                 mSumElapsed = mSumElapsed + row.ElapsedTime
450                 mNumEntries = mNumEntries + 1
                    'mSumFinalized = mSumFinalized + row.FinalizedElapsedTime
460               End If

470               If row.FinalizeElapsed > 0 Then
480                 mSumFinalized = mSumFinalized + row.FinalizeElapsed
490                 mNumFinalized = mNumFinalized + 1

500               End If


510             End If


520         End Select
            Dim FinalizeElapsedTimestring As String
530         If row.FinalizeElapsed > 0 Then
540           FinalizeElapsedTimestring = SecondsToTimeString(row.FinalizeElapsed)
550         Else
560           FinalizeElapsedTimestring = " "
570         End If




            '      mNumEntries = mNumEntries + 1
            '      mSumElapsed = mSumElapsed + row.ElapsedTime

580         text = text & "<!-- AlarmID " & row.ID & " RoomID " & row.RoomID & " ResID " & row.ResidentID & "-->" & vbCrLf
590         text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
            '470         text = text & "<td>" & row.Serial & "</td><td>" & HTMLEncode(row.Room) & "</td><td>" & HTMLEncode(row.EventName) & "</td><td>" & HTMLEncode(row.Announce) & "</td><td>" & row.AlarmTimeToString & "</td><td>" & row.ElapsedTimeToString & "</td><td>" & HTMLEncode(row.User) & "</td>" & vbCrLf
600         text = text & "<td>" & HTMLEncode(row.Room) & "</td><td>" & HTMLEncode(row.Resident) & "</td><td>" & HTMLEncode(row.EventName) & "</td><td>" & HTMLEncode(row.Announce) & "</td><td>" & row.AlarmTimeToString & "</td><td>" & row.ElapsedTimeToString & "</td><td>" & HTMLEncode(row.Responder) & "</td><td>" & HTMLEncode(FinalizeElapsedTimestring) & "</td><td>" & HTMLEncode(row.FinalizeText) & "</td>" & vbCrLf
610         text = text & "</tr>" & vbCrLf
620         odd = Not odd
630       Next
640     End If
650     text = text & "<tr " & "class='odd'" & ">" & vbCrLf
660     text = text + "<td colspan='9'><hr/></td>" & vbCrLf
670     text = text & "</tr>" & vbCrLf

680     If mNumEntries Then
690       mAverageElapsed = mSumElapsed / mNumEntries
700     End If

710     If mNumFinalized > 0 Then
720       mAverageFinalized = mSumFinalized / mNumFinalized
730     End If



740     mAverageElapsedString = "N/A"
750     If (mAverageElapsed > 0) Then
760       mAverageElapsedString = Format$(DateAdd("s", mAverageElapsed, 0), "hh:nn:ss")
770     End If

780     mAverageFinalizedString = "N/A"
790     If (mAverageFinalized > 0) Then
800       mAverageFinalizedString = Format$(DateAdd("s", mAverageFinalized, 0), "hh:nn:ss")
810     End If



820     text = text & "<tr " & "class='average'" & ">" & vbCrLf
830     text = text + "<td>Average</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>" & mAverageElapsedString & "</td><td>&nbsp;</td><td>" & mAverageFinalizedString & "</td><td>&nbsp;</td>" & vbCrLf
840     text = text & "</tr>" & vbCrLf

850     text = text & "</table>" & vbCrLf
860     text = text + "<p class='complete'>Report Complete V" & App.Revision & "</p>"
870     text = text & "</body>" & vbCrLf
880     text = text & "</html>"

890     DoHTMLReport = text

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
'  Dim History As cAlarmHistory
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
  If 0 = CLng(NextReportDue) Then  ' first go-round
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
  Dim daybit             As Long
  Dim DOW                As Long

  Dim j                  As Long


10 CurrentDate = Now

20 GetNextReportDate = 0

30 Erase WeekDays()
40 Erase ReportDays()
50 Erase DailyReports()

  '60      Debug.Assert 0

60 Select Case DayPeriod
    Case AUTOREPORT_DAILY
70    If (NextReportDue = 0) Then
80      Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs TODAY
90      If DayPartEnd = 0 Then
100       Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
110     Else
120       Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
130     End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
140     WeekDayNumber = Weekday(Now) - 1
150     For j = 0 To 6
160       DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
170       Makedate = DateAdd("d", -DOW, Makedate)  ' roll back to sunday
180       daybit = (2 ^ j)
          '1 = monday, 2 = tues etc

190       If (daybit And DAYS) <> 0 Then
200         ReportDays(j) = 1

210         DailyReports(j) = DateAdd("d", j, Makedate)  ' add 0 to j days to sunday

220         If DayPartEnd = 0 Then
230           DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '

240         End If
250         If DailyReports(j) < CurrentDate Then
260           DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  ' it was earlier than today, add a week to it

270         End If
280       End If
290       Debug.Print DailyReports(j) & "  " & Format(DailyReports(j), "ddd")
300     Next


310   End If

320   SortDates DailyReports()
330   For j = 0 To 7
340     If DailyReports(j) > Now Then  ' get earliest occurrence after now
350       GetNextReportDate = DailyReports(j)
360       Exit For
370     End If
380   Next




390 Case AUTOREPORT_SHIFT1
400   If (NextReportDue = 0) Then
410     Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
420     If DayPartEnd = 0 Then
430       Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
440     Else
450       Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
460     End If

        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday

470     WeekDayNumber = Weekday(Now) - 1
480     For j = 0 To 6
490       daybit = (2 ^ j)
500       If (daybit And DAYS) <> 0 Then
510         DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
520         Makedate = DateAdd("d", -DOW, Makedate)
530         ReportDays(j) = 1
540         DailyReports(j) = DateAdd("d", j, Makedate)  '
550         If DayPartEnd = 0 Then
560           DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
570         End If
580         If DailyReports(j) < CurrentDate Then
590           DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
600         End If
610       End If
620     Next


630   End If

640   SortDates DailyReports()
650   For j = 0 To 7
660     If DailyReports(j) > Now Then
670       GetNextReportDate = DailyReports(j)
680       Exit For
690     End If
700   Next


710 Case AUTOREPORT_SHIFT2
720   If (NextReportDue = 0) Then
730     Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
740     If DayPartEnd = 0 Then
750       Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
760     Else
770       Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
780     End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
790     WeekDayNumber = Weekday(Now) - 1
800     For j = 0 To 6
810       daybit = (2 ^ j)
820       If (daybit And DAYS) <> 0 Then
830         DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
840         Makedate = DateAdd("d", -DOW, Makedate)

850         ReportDays(j) = 1
860         DailyReports(j) = DateAdd("d", j, Makedate)  '
870         If DayPartEnd = 0 Then
880           DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
890         End If
900         If DailyReports(j) < CurrentDate Then
910           DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
920         End If
930       End If
940     Next
950   End If

960   SortDates DailyReports()
970   For j = 0 To 7
980     If DailyReports(j) > Now Then
990       GetNextReportDate = DailyReports(j)
1000      Exit For
1010    End If
1020  Next


1030 Case AUTOREPORT_SHIFT3
1040  If (NextReportDue = 0) Then
1050    Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
1060    If DayPartEnd = 0 Then
1070      Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
1080    Else
1090      Makedate = DateAdd("h", DayPartEnd, Makedate)  ' add in closing hour
1100    End If
        'CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)                        ' returns 1 for sunday
1110    WeekDayNumber = Weekday(Now) - 1
1120    For j = 0 To 6
1130      daybit = (2 ^ j)
1140      If (daybit And DAYS) <> 0 Then
1150        DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
1160        Makedate = DateAdd("d", -DOW, Makedate)

1170        ReportDays(j) = 1
1180        DailyReports(j) = DateAdd("d", j, Makedate)  '
1190        If DayPartEnd = 0 Then
1200          DailyReports(j) = DateAdd("h", 24, DailyReports(j))  '
1210        End If
1220        If DailyReports(j) < CurrentDate Then
1230          DailyReports(j) = DateAdd("ww", 1, DailyReports(j))  '
1240        End If
1250      End If
1260    Next


1270  End If
1280  SortDates DailyReports()
1290  For j = 0 To 7
1300    If DailyReports(j) > Now Then
1310      GetNextReportDate = DailyReports(j)
1320      Exit For
1330    End If
1340  Next



1350 Case AUTOREPORT_WEEKLY    ' once a week at end of week
1360  If (NextReportDue = 0) Then
1370    Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
1380    CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)  ' returns 1 for sunday
1390    Makedate = DateAdd("d", 8 - CurrentDOW, Makedate)
1400    GetNextReportDate = Makedate
1410  End If

1420 Case AUTOREPORT_MONTHLY   ' once a month, at end of month
1430  If (NextReportDue = 0) Then
1440    Makedate = DateSerial(Year(Now), Month(Now), 1)  ' make first of current month
1450    Makedate = DateAdd("m", 1, Makedate)  ' add a month
1460    GetNextReportDate = Makedate  ' DateAdd("n", -1, Makedate)  ' ?? minus one minute
1470  End If
1480 Case Else

1490 End Select




End Function
'Function GetNextBackupDate() As Date
'  Dim j As Integer
'  Dim MonthDayNumber        As Integer
'  Dim WeekDayNumber         As Integer
'  Dim MakeDate              As Date
'  Dim WeekDays(0 To 7)      As Integer
'  Dim ReportDays(0 To 7)    As Integer
'  Dim ReporReports(0 To 7)  As Date
'
'  Dim CurrentDate           As Date
'
'  Dim MonthlyBackups() As String
'
'  MonthDayNumber = Day(Now)
'  WeekDayNumber = Weekday(Now, vbSunday) - 1
'
'  Select Case Configuration.BackupType
'    Case 1  ' monthly
'
'
'      If Configuration.BackupDOM <= 0 Or Configuration.BackupDOM > 28 Then
'        Configuration.BackupDOM = 1
'      End If
'      MakeDate = DateSerial(Year(Now), Month(Now), Configuration.BackupDOM)
'      MakeDate = DateAdd("h", Configuration.BackupTime / 100, MakeDate)
'      If MakeDate < Now Then
'        MakeDate = DateAdd("m", 1, MakeDate)
'      End If
'      GetNextBackupDate = MakeDate
'
'
'    Case Else  ' days -  weekly
'      MakeDate = DateSerial(Year(Now), Month(Now), Day(Now)) ' create
'      MakeDate = DateAdd("h", Configuration.BackupTime / 100, MakeDate)
'
'      For j = 0 To 6
'        If ((2 ^ j) And Configuration.BackupDOW) <> 0 Then
'          ReportDays(j) = 1
'          WeekDays(j) = j - WeekDayNumber
'          If WeekDays(j) < 0 Then
'            WeekDays(j) = WeekDays(j) + 7
'          End If
'          ReporReports(j) = DateAdd("d", WeekDays(j), MakeDate)
'        End If
'      Next
'
'      If ReportDays(1) = 1 Then
'        ReportDays(7) = 1
'        ReporReports(7) = DateAdd("d", 7, MakeDate)
'      End If
'
'
'      SortDates ReporReports()
'      For j = 0 To 7
'
'        'If ReportDays(j) = 1 Then
'          If ReporReports(j) > Now Then
'            GetNextBackupDate = ReporReports(j)
'
'            Exit For
'          End If
'        'End If
'      Next
''      For j = 0 To 7
''
''        If ReportDays(j) = 1 Then
''          If ReporReports(j) > Now Then
''            GetNextBackupDate = ReporReports(j)
''
''            Exit For
''          End If
''        End If
''      Next
'
'
'
'  End Select
'
'End Function




Function Parse(rs As ADODB.Recordset) As Boolean
  Dim j                  As Integer
  Dim Room               As cRoom
  Dim dataWrapper        As cDataWrapper
  Dim Items()            As String

  ReportID = rs("reportid")
  Disabled = rs("Disabled")
  ReportName = rs("reportname") & ""
  Comment = rs("Comment") & ""
  RoomString = rs("Rooms") & ""
  EventString = rs("Events") & ""
  TimePeriod = rs("TimePeriod")
  DayPeriod = rs("DayPeriod")
  DAYS = rs("Days")
  Shift = rs("Shift")
  DayPartStart = rs("DayPartStart")
  DayPartEnd = rs("DayPartend")
  SortOrder = rs("SortOrder")
  SendHour = rs("SendHour")
  SaveAsFile = 1               '  rs("SaveAsFile")
  SendAsEmail = rs("SendAsEmail")
  recipient = rs("Recipient") & ""
  Subject = rs("Subject") & ""
  FileFormat = rs("FileFormat")
  DestFolder = rs("DestFolder") & ""

  Set Rooms = New Collection

  Items = Split(RoomString, ",")
  For j = LBound(Items) To UBound(Items)
    Set Room = New cRoom
    Room.RoomID = Val(Items(j))
    Rooms.Add Room
  Next

  Set Events = New Collection

  Items = Split(EventString, ",")
  For j = LBound(Items) To UBound(Items)
    Set dataWrapper = New cDataWrapper
    dataWrapper.LongValue = Val(Items(j))
    Events.Add dataWrapper
  Next

  Dim HasSecondShift     As Boolean
  Dim HasThirdShift      As Boolean

  If Configuration.EndFirst = Configuration.EndNight Then  ' no second or third shift' regardless of third shift ending
    HasSecondShift = False
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = True
  End If

  Select Case DayPeriod
    Case AUTOREPORT_SHIFT1
      If HasThirdShift Then
        DayPartStart = Configuration.EndThird
        DayPartEnd = Configuration.EndFirst
      ElseIf HasSecondShift Then
        DayPartStart = Configuration.EndNight
        DayPartEnd = Configuration.EndFirst
      Else                     ' all one shift
        DayPartStart = 0
        DayPartEnd = 24
      End If

    Case AUTOREPORT_SHIFT2
      If HasThirdShift = True Then
        DayPartStart = Configuration.EndFirst
        DayPartEnd = Configuration.EndNight
      Else
        DayPartStart = Configuration.EndFirst
        DayPartEnd = Configuration.EndNight
      End If

    Case AUTOREPORT_SHIFT3
      ' got's to have a third shift
      DayPartStart = Configuration.EndNight
      DayPartEnd = Configuration.EndThird
    Case AUTOREPORT_DAILY
      DayPartStart = Val(rs("daypartstart") & "")
      DayPartEnd = Val(rs("daypartend") & "")
    Case Else
      DayPartStart = 0
      DayPartEnd = 24
  End Select



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
    RoomString = Join(Items, ",")
  End If
  EventString = ""
  If Events.Count Then
    ReDim Items(Events.Count)
    For j = 1 To Events.Count
      Set dataWrapper = Events(j)
      Items(j) = CStr(dataWrapper.LongValue)
    Next
    EventString = Join(Items, ",")
  End If

  ' auto increment field:  rs("reportid") = ReportID
  rs("Disabled") = Disabled
  rs("reportname") = ReportName
  rs("Comment") = Comment
  rs("Rooms") = RoomString
  rs("Events") = EventString
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
  Subject = "Auto Report"
  Set Rooms = New Collection
  Set Events = New Collection
  Set EventItems = New Collection
  PeriodStart = Now
End Sub

Private Sub Class_Terminate()
  Finalize
End Sub

Public Property Get Busy() As Boolean  ' Public read only
  Busy = mBusy
End Property



Public Property Get NextReportDue() As Date
  NextReportDue = mNextReportDue
End Property

Public Property Let NextReportDue(ByVal NextReportDue As Date)
  mNextReportDue = NextReportDue
End Property
