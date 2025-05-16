VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExceptionAutoReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' for externs, need report Item to show message received.

Option Base 0


Public ResponseTime As Long     ' seconds for responding to alarm
Public ResponseIsACK As Long   ' non-zero if ACK is response time



Public ReportType   As Long

Public ReportID     As Long
Public Disabled     As Integer  '1 = Disabled, 0=enabled
Public ReportName   As String
Public Comment      As String

Public Rooms        As Collection  ' what rooms to include
Public Events       As Collection  ' what events for this report
Public DevTypes     As Collection

Public RoomString   As String   ' what rooms to include - delimited list
Public EventString  As String   ' what events for this report  - delimited list
Public DevString    As String   ' what Devices for this report  - delimited list

Public TimePeriod   As Long     ' Not Used From Dropdown list

Public DayPeriod    As Long     ' days or monthly
Public DAYS         As Long     ' as bitfield
Public DayString    As String   ' as string

Public Shift        As Integer  ' 0=all, 1=first, 2=second, 3= third) ( 3rd not implemented)

Public DayPartStart As Long     ' if a specific time of day is requested (start and end = 0 ... all day)
Public DayPartEnd   As Long

Public SortOrder    As Long     ' 0 = Chron, 1=long-to-short, 2=By-Room

Public SendHour     As Long     '  NOT USED - 0 to 2400 and end of event ' when to actually create the report

Public EventItems   As Collection

Public SaveAsFile   As Integer  ' always for now
Public SendAsEmail  As Integer

Public recipient    As String   ' email address
Public Subject      As String   ' email subject

Public FileFormat   As Integer  ' TAB, TAB no header, HTML
Public DestFolder   As String   ' where to keep files

Private mNextReportDue As Date     ' when we're up

Private mBusy       As Boolean  ' revent reentrancy

Private Const COL_SERIAL = 0
Private Const COL_ROOM = 0.75
Private Const COL_RESIDENT = 1.5
Private Const COL_EVENT = 3.5
Private Const COL_ALARMTIME = 6.5
Private Const COL_CANCELTIME = 7.5
Private Const COL_ELAPSEDTIME = 8.5
Private Const COL_OPERATOR = 8.75

Public PeriodEnd    As Date
Public PeriodStart  As Date

Private mTotalElapsed As Long

Private mNumEntries As Long
Private mSumElapsed As Long



Private mAlarmCount   As Long

Private mAverageElapsed As Long


Private AllRows     As Collection     ' all alarms as rows of report cAlarmHistory
Private rows        As Collection        ' each row chosen for this report

Public Function DoReport()

  PeriodEnd = NextReportDue
  NextReportDue = 0
  GenAutoReport

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


        Dim Path          As String
        Dim filename      As String
        Dim ext           As String
        Dim hfile         As Integer

        Dim SQL           As String
        Dim WhereClause   As String

        Dim rs            As ADODB.Recordset
        Dim AlarmHistory  As Object
        Dim CurrentDate   As Date

        Dim SQL2           As String
        Dim Whereclause2   As String

        Dim rs2            As ADODB.Recordset

        Dim Room          As cRoom

        Dim TempEvents    As Collection
        Dim rc            As Long

        Dim FileText      As String

        Dim Criteria      As String  ' not used here

        ' select all alarms that match criteria with basic info
        ' Walk each alarm
        ' Fetch each event for each alarm
        ' Initially only get initiating events, history comes later
        ' Alarmid = 0 if first event for alarm

        Dim j             As Integer
        Dim Do_UnAssigned As Boolean

10      On Error GoTo GenAutoReport_Error

20      Debug.Print "report start " & Now
        Dim HasSecondShift As Boolean
        Dim HasThirdShift As Boolean

30      If Configuration.EndFirst = Configuration.EndNight Then
40        HasSecondShift = False
50        HasThirdShift = False
60      ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
70        HasSecondShift = True
80        HasThirdShift = False
90      ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
100       HasSecondShift = True
110       HasThirdShift = True
120     Else
130       HasSecondShift = False
140       HasThirdShift = False
150     End If

160     Select Case DayPeriod
        Case AUTOREPORT_DAILY
170       PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
180       PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
190       If PeriodStart >= PeriodEnd Then
200         PeriodStart = DateAdd("d", -1, PeriodStart)
210       End If

220     Case AUTOREPORT_SHIFT1
230       PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
240       PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
250       If PeriodStart >= PeriodEnd Then
260         PeriodStart = DateAdd("d", -1, PeriodStart)
270       End If

280     Case AUTOREPORT_SHIFT2
290       PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
300       PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
310       If PeriodStart >= PeriodEnd Then
320         PeriodStart = DateAdd("d", -1, PeriodStart)
330       End If


340     Case AUTOREPORT_SHIFT3
350       PeriodStart = DateSerial(Year(PeriodEnd), Month(PeriodEnd), Day(PeriodEnd))
360       PeriodStart = DateAdd("h", Me.DayPartStart, PeriodStart)
370       If PeriodStart >= PeriodEnd Then
380         PeriodStart = DateAdd("d", -1, PeriodStart)
390       End If


400     Case AUTOREPORT_WEEKLY
410       PeriodStart = DateAdd("ww", -1, PeriodEnd)

420     Case AUTOREPORT_MONTHLY
430       PeriodStart = DateAdd("m", -1, PeriodEnd)

440     End Select

450     'Criteria = ""  ' not used here

460     Select Case ReportType
        Case RPT_INV
          rc = FindInv(Criteria)    ' Function FindInv(ByVal Criteria As String)
        Case RPT_COUNT
470       rc = FindCount(Criteria, PeriodStart, PeriodEnd)
480     Case Else  ' RPT_EXCEPTION
490       rc = FindException(Criteria, PeriodStart, PeriodEnd)
500     End Select



510     Select Case FileFormat
        Case AUTOREPORTFORMAT_TAB_NOHEADER
520       FileText = DoTabReport(rows, Headers:=False)
530       ext = ".txt"
540     Case AUTOREPORTFORMAT_HTML
550       FileText = DoHTMLReport(rows)
560       ext = ".html"
570     Case Else  '     Case AUTOREPORTFORMAT_TAB
580       FileText = DoTabReport(rows, Headers:=True)
590       ext = ".txt"
600     End Select
        ' write file to dest folder

610     hfile = FreeFile

620     DestFolder = App.Path & "\AutoReports\"
630     Path = DestFolder

640     If Right(Path, 1) <> "\" Then
650       Path = Path & "\"
660     End If

670     If Not DirExists(Path) Then
680       On Error Resume Next
690       MkDir Path
700     End If

710     If Not DirExists(Path) Then
720       Path = App.Path & "\"
730     End If



740     filename = Path & CleanFileName(ReportName) & " " & Format$(PeriodEnd, "yyyymmdd hh") & ext

750     Open filename For Output As #hfile
760     Print #hfile, FileText
770     Close #hfile

        'Debug.Print "FileText: " & FileText

780     If (SendAsEmail And (Len(recipient) > 0)) Then

          Dim mapi        As Object
          Dim message     As String

790       message = "Auto Report Attached: " & filename

800       On Error Resume Next

810       If (Configuration.UseSMTP = MAIL_SMTP) Then

820         If gSMTPMailer Is Nothing Then
              '630           'Set gSMTPMailer = New SendMail
830           Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
840         End If
850         If gSMTPMailer Is Nothing Then
860           LogProgramError "Could not create SMTPMailer Object in cExceptionReport.GenAutoReport." & Erl
870         Else
880           Call gSMTPMailer.Send("", "", recipient, Subject, message, filename)
890         End If


900       Else
910         Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
920         If mapi Is Nothing Then
930           LogProgramError "Could not create SENTRYMAIL Object in cException.GenAutoReport." & Erl
940         Else
950           Call mapi.SendWithAttachments("", "", recipient, Subject, message, filename)
960         End If

970       End If


          '// Username, Password, Address, Subject,Body, AttachmentsList ' Attachemnet list is a semicolon ";" delimited list of file attachments
          'Call mapi.Send("", "", Configuration.AssurEmailRecipient, Configuration.AssurEmailSubject, Message)

980       Set mapi = Nothing

990     End If
1000    Debug.Print "report end " & Now

GenAutoReport_Resume:

1010    On Error GoTo 0
1020    Exit Function

GenAutoReport_Error:

1030    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionAutoReport.GenAutoReport." & Erl
1040    Resume GenAutoReport_Resume

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

  Dim row           As Object
  Dim text          As String
  Dim counter       As Long

  If (Headers) Then
    Select Case ReportType
    Case RPT_INV
    text = Join(Array("Device", "Model", "Last", "First", "Room"), vbTab) & vbCrLf
    Case RPT_COUNT
      text = Join(Array("Model", "Desc", "Count"), vbTab) & vbCrLf
    Case Else   ' RPT_EXCEPTION
      text = Join(Array("Device", "Room", "Resident", "Event", "Announce", "Alarm Time", "Canceled", "Elapsed", "Operator"), vbTab) & vbCrLf
    End Select
  End If

  For Each row In rows
    counter = counter + 1
    If counter > 200 Then
      counter = 0
      DoEvents
    End If
    Select Case ReportType
    Case RPT_INV
      text = text & row.Serial & vbTab & row.DeviceType & vbTab & row.NameLast & vbTab & row.NameFirst & vbTab & row.Room & vbCrLf
    Case RPT_COUNT
      text = text & row.Model & vbTab & row.Description & vbTab & row.Count & vbCrLf
    Case Else  ' RPT_EXCEPTION

      text = text & row.ToTABBEDString & vbCrLf  '   Join(Array(row.Serial, row.room, row.Resident, row.EventName, row.AlarmTimeToString, row.CancelTimeToString, row.user), vbTab) & vbCrLf
    End Select
  Next
  
  
  If ReportType = RPT_EXCEPTION Then
              
              Dim AverageTime    As String
              Dim averge         As Long ' need to remove
              
              Dim Average        As Long
              
         If mNumEntries > 0 Then
           Average = mSumElapsed / mNumEntries
           AverageTime = Format(DateAdd("s", Average, 0), "hh:nn:ss")
         Else
           AverageTime = "N/A"
         End If

              

         text = text & Join(Array("Average", "", "", "", "", "", "", AverageTime, ""), vbTab) & vbCrLf


  End If

  
  DoTabReport = text

End Function

Function DoHTMLReport(rows As Collection) As String
        Dim row                As Object
        Dim odd                As Boolean
        Dim text               As String
        Dim Announce           As String
        Dim counter            As Long
        Dim Average            As Long


        ' html header
10      On Error GoTo DoHTMLReport_Error

20      text = "<html>"
30      text = text & "<head>"
40      text = text + "<style type=""text/css"">"
50      text = text + "body {width:900px; font-family:arial,verdana,sans-serif;}"
60      text = text + "table.main {width:900px;font-size:1.0em;}"
70      text = text + "tr.header td {background-color: #ADD8E6; color: black; margin:0px; padding:2px; font-weight:bold;}"
80      text = text + "tr.even td {background-color: #FAFAD2; color: black; margin:0px; padding:2px}"
90      text = text + "tr.odd td {background-color: white; color: black; margin:0px; padding:2px}"
100     text = text + "tr.average td {background-color: #d3d3d3; color: black; margin:0px; padding:2px}"
110     text = text + "h1 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.3em}"
120     text = text + "h2 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.0em;}"
130     text = text + "p.complete {background-color: white; color:gray;margin:5px;text-align:left;font-size:0.9em;}"

140     text = text + "</style>"

150     text = text & "</head>" & vbCrLf
160     text = text & "<body>"

170     text = text + "<h1>" & HTMLEncode(ReportName) & "</h1>"
180     text = text + "<h1>" & HTMLEncode(Configuration.Facility) & "</h1>"

190     If ReportType = RPT_INV Then
200       text = text + "<h2>" & "Date: " & HTMLEncode(Format(Now, "mm/dd/yyyy")) & "</h2>"
210     Else
220       text = text + "<h2>" & "Date: " & HTMLEncode(PeriodStart) & " to " & HTMLEncode(PeriodEnd) & "</h2>"
230     End If

240     text = text + "<br />"

250     Select Case ReportType
          Case RPT_INV
260         text = text & "<table class='main'>" & vbCrLf
270         text = text + "<tr class='header'>" & vbCrLf

280         text = text + "<td>Serial</td><td>Model</td><td>Last</td><td>First</td> <td>Room</td>" & vbCrLf
290         text = text + "</tr>" & vbCrLf
300         If rows.Count Then
310           For Each row In rows
320             counter = counter + 1
330             If counter > 200 Then
340               counter = 0
350               DoEvents
360             End If
370             text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
380             text = text & "<td>" & HTMLEncode(row.Serial) & "</td><td>" & HTMLEncode(row.DeviceType) & "</td><td>" & HTMLEncode(row.NameLast) & "</td><td>" & HTMLEncode(row.NameFirst) & "</td><td>" & row.Room & "</td>" & vbCrLf
390             text = text & "</tr>" & vbCrLf
400             odd = Not odd
410           Next
420         End If
430         text = text & "</table>" & vbCrLf
440       Case RPT_COUNT

450         text = text & "<table class='main'>" & vbCrLf
460         text = text + "<tr class='header'>" & vbCrLf


470         text = text + "<td>Model</td> <td>Description</td> <td>Count</td>" & vbCrLf
480         text = text + "</tr>" & vbCrLf


490         If rows.Count Then
500           For Each row In rows
510             counter = counter + 1
520             If counter > 200 Then
530               counter = 0
540               DoEvents
550             End If

                ' text = text & "<!-- AlarmID " & row.id & " RoomID " & row.RoomID & " ResID " & row.ResidentID & "-->" & vbCrLf  ' just an HTML comment line
560             text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
570             text = text & "<td>" & HTMLEncode(row.Model) & "</td><td>" & HTMLEncode(row.Description) & "</td><td>" & row.Count & "</td>" & vbCrLf
580             text = text & "</tr>" & vbCrLf
590             odd = Not odd
600           Next
610         End If

620         text = text & "</table>" & vbCrLf



630       Case Else                  ' ReportType = RPT_Exception

640         text = text & "<table class='main'>" & vbCrLf
650         text = text + "<tr class='header'>" & vbCrLf
660         text = text + "<td>Serial</td> <td>Resident</td> <td>Room</td> <td>Event</td> <td>Announce</td> <td>Time</td> <td>Elapsed</td><td>Operator</td>" & vbCrLf
670         text = text + "</tr>" & vbCrLf


680         If rows.Count Then
690           For Each row In rows
700             counter = counter + 1
710             If counter > 200 Then
720               counter = 0
730               DoEvents
740             End If

                'text = text & "<!-- AlarmID " & row.id & " RoomID " & row.RoomID & " ResID " & row.ResidentID & "-->" & vbCrLf  ' just an HTML comment line
750             text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf



760             Select Case row.EventType
                  Case EVT_EMERGENCY, EVT_ALERT  ' change 2/15/2012
770                 Announce = row.Announce
780               Case Else
790                 Announce = row.EventName
800             End Select
810             text = text & "<td>" & row.Serial & "</td><td>" & HTMLEncode(row.Resident) & "</td><td>" & HTMLEncode(row.Room) & "</td><td>" & HTMLEncode(row.EventName) & "</td><td>" & HTMLEncode(Announce) & "</td><td>" & row.AlarmTimeToString & "</td><td>" & row.ElapsedTimeToString & "</td><td>" & HTMLEncode(row.User) & "</td>" & vbCrLf
820             text = text & "</tr>" & vbCrLf
830             odd = Not odd
840           Next

850         End If

            Dim AverageTime    As String

860         text = text & "<tr " & "class='odd'" & ">" & vbCrLf
870         text = text + "<td colspan='8'><hr/></td>" & vbCrLf
880         text = text & "</tr>" & vbCrLf

890         If mNumEntries > 0 Then
900           Average = mSumElapsed / mNumEntries
910           AverageTime = Format(DateAdd("s", Average, 0), "hh:nn:ss")
920         Else
930           AverageTime = "N/A"
940         End If
950         text = text & "<tr " & "class='average'" & ">" & vbCrLf
'960         text = text & "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
970         text = text & "<td>Average</td><td></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>" & AverageTime & "</td><td>&nbsp;</td>" & vbCrLf
980         text = text & "</tr>" & vbCrLf




990         text = text & "</table>" & vbCrLf

1000    End Select
1010    text = text + "<p class='complete'>Report Complete</p>"
1020    text = text & "</body>" & vbCrLf
1030    text = text & "</html>"

1040    DoHTMLReport = text

DoHTMLReport_Resume:

1050    On Error GoTo 0
1060    Exit Function

DoHTMLReport_Error:

1070    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionAutoReport.DoHTMLReport." & Erl
1080    Resume DoHTMLReport_Resume

End Function
Function FindInv(ByVal Criteria As String)
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim rsReport           As ADODB.Recordset
        Dim params()           As String
        Dim InventoryReport    As cExceptionReport
        Dim DeviceTypes        As String
        Dim models             As String
        Dim modelList()        As String
        Dim row                As cAlarmHistory

        Dim ESDeviceType       As ESDeviceTypeType

        Dim PTIList()          As String
        Dim PTIString          As String

        Dim RequestedModelString As String

        Dim DEVList()          As ModelPTI

        Dim j                  As Long
        Dim i                  As Long

        Dim PTIs               As String

        Dim rsdev              As ADODB.Recordset


10      On Error GoTo FindInv_Error
20      Set rows = New Collection

30      ReDim DEVList(0)

40      DEVList(0).CLSPTI = 0
50      SQL = "SELECT DISTINCT model FROM Devices"
60      Set rs = ConnExecute(SQL)
        ' we now should have a list of models actually used

70      Do Until rs.EOF
80        i = i + 1
90        ReDim Preserve DEVList(i) As ModelPTI
100       DEVList(i).Model = rs("model") & ""
110       DEVList(i).CLSPTI = GetCLSPTI(DEVList(i).Model)
120       rs.MoveNext
130     Loop


140     rs.Close
150     Set rs = Nothing


160     Set InventoryReport = New cExceptionReport
170     'params = Split(Criteria, ",")
180     SQL = "SELECT * FROM ExceptionReports WHERE ReportID = " & Me.ReportID   '   Val(params(0))
190     Set rsReport = ConnExecute(SQL)
200     If Not (rsReport.EOF) Then

210       InventoryReport.Parse rsReport

220       PTIList() = Split(InventoryReport.DevString, vbTab)

230       ReDim RequestedModels(0)
          Dim n                As Long

240       For j = 1 To UBound(DEVList)
250         For i = LBound(PTIList) To UBound(PTIList)
260           If DEVList(j).CLSPTI = Val(PTIList(i)) Then
270             ReDim Preserve RequestedModels(n)


280             RequestedModels(n) = q(DEVList(j).Model)
290             n = n + 1
300             Exit For
310           End If
320         Next
330       Next

340       RequestedModelString = Join(RequestedModels, ",")



350       SQL = "SELECT Devices.Serial, devices.model, Rooms.Room, Residents.namelast, residents.namefirst FROM "
360       SQL = SQL & " ( Devices  LEFT JOIN residents ON Devices.residentid =  Residents.ResidentID )"
370       SQL = SQL & " LEFT JOIN Rooms ON devices.roomid = Rooms.roomid "
          SQL = SQL & " WHERE devices.model in ( " & RequestedModelString & ")"
380       SQL = SQL & " ORDER BY Rooms.room, devices.model"

390       Set rs = ConnExecute(SQL)

420       Do Until rs.EOF
430         Set row = New cAlarmHistory
440         row.Serial = rs("serial") & ""
            ESDeviceType = GetESDeviceTypeByModel(rs("model") & "")
            row.DeviceType = ESDeviceType.desc & " (" & ESDeviceType.Model & ")"
'450         row.DeviceType = rs("model") & ""
460         row.NameFirst = rs("namefirst") & ""
470         row.NameLast = rs("namelast") & ""
480         row.Room = rs("room") & ""
490         rows.Add row             'Print #hfile, rs("serial") & "," & rs("Model") & "," & rs("namelast") & "," & rs("namefirst") & "," & rs("room")
500         rs.MoveNext
510       Loop
520
530       rs.Close


540     End If
550     rsReport.Close

FindInv_Resume:

560     Set rs = Nothing
570     Set rsReport = Nothing

580     On Error GoTo 0
590     Exit Function

FindInv_Error:

600     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionAutoReport.FindInv." & Erl
610     Resume FindInv_Resume

End Function

'Public Function ShellSortElapsedTime(c As Collection, ByVal Descending As Boolean) As Collection
'
'  Dim Hold          As Long  ' specific to sort
'  Dim Gap           As Long  ' specific to sort
'  Dim i             As Long  ' specific to sort
'  Dim Min           As Long  ' always 1
'  Dim Max           As Long  ' same as number of objects in collection
'  Dim Compare       As Long  ' the current value
'  Dim IndexValue    As Long  ' the current index
'  Dim TestValue     As Long  ' to test against
'  Dim j             As Long
'  Dim sx()          As Long  ' array for sorting of indexes
'
'  Dim hist          As cAlarmHistory
'
'  Dim newcoll       As Collection    ' the new collection
'
'  Set newcoll = New Collection
'
'  ReDim sx(c.Count, 2) As Long  ' we need a count number of elements with 2 columns
'
'  For j = 1 To c.Count
'    Set hist = c(j)
'
'    sx(j, 1) = hist.ElapsedTime
'    sx(j, 2) = j
'  Next
'
'  '        Debug.Print
'  '    For j = 1 To c.Count
'  '      Debug.Print sx(j, 1), sx(j, 2)
'  '    Next
'
'  Min = 1
'  Max = c.Count
'  Gap = Min
'
'  Do    ' figureb optimum gap
'    Gap = 3 * Gap + 1
'  Loop Until Gap > Max
'
'  Dim t
'  t = Timer
'
'  Do
'    Gap = Gap \ 3
'    For i = Gap + Min To Max
'      DoEvents
'      Compare = sx(i, 1)
'      IndexValue = sx(i, 2)
'      Hold = i
'
'      If Descending Then
'
'        TestValue = sx(Hold - Gap, 1)
'        Do While TestValue < Compare
'          DoEvents
'          ' swap the value and the index
'          sx(Hold, 1) = sx(Hold - Gap, 1)  ' swap real values
'          sx(Hold, 2) = sx(Hold - Gap, 2)  ' swap indexes
'          Hold = Hold - Gap
'          If Hold < Min + Gap Then
'            Exit Do
'          End If
'          TestValue = sx(Hold - Gap, 1)
'        Loop
'      Else  ' ascending
'        TestValue = sx(Hold - Gap, 1)
'        Do While TestValue > Compare
'          DoEvents
'          ' swap the value and the index
'          sx(Hold, 1) = sx(Hold - Gap, 1)
'          sx(Hold, 2) = sx(Hold - Gap, 2)
'          Hold = Hold - Gap
'          If Hold < Min + Gap Then
'            Exit Do
'          End If
'          TestValue = sx(Hold - Gap, 1)
'        Loop
'      End If
'
'      sx(Hold, 1) = Compare
'      sx(Hold, 2) = IndexValue
'    Next i
'
'  Loop Until Gap = 1
'
'
'
'  'Debug.Print "time " & Timer - t
'
'  For j = 1 To Max
'    ' Debug.Print sx(j, 1), sx(j, 2)
'    newcoll.Add c(sx(j, 2))
'  Next
'  Set c = Nothing
'  Set ShellSortElapsedTime = newcoll
'
'
'End Function

'Private Function GetAlarmHistory(rs As Recordset) As cAlarmHistory
'  Dim History       As cAlarmHistory
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
  Dim timeleft      As Double
  If 0 = NextReportDue Then  ' first go-round
    NextReportDue = GetNextReportDate()
  End If

  If NextReportDue <> 0 Then
    timeleft = DateDiff("s", Now, NextReportDue)
    'Debug.Print "Next AutoReport Due '" & ReportName & "' " & Format(NextReportDue, "mm/dd/yy h:n:s") & " in " & timeleft \ 60 & ":" & timeleft Mod 60 & " Min:Secs"

    due = (Now > NextReportDue)
  End If

End Property
Public Function GetNextReportDate() As Date
  Dim Makedate      As Date
  Dim CurrentDate   As Date
  Dim CurrentDOW    As Integer
  Dim WeekDayNumber As Integer

  Dim WeekDays(0 To 7) As Integer
  Dim ReportDays(0 To 7) As Integer
  Dim DailyReports(0 To 7) As Date

  Dim j             As Long

  Dim daybit As Long
  Dim DOW As Long

  

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
      
         DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
         Makedate = DateAdd("d", -DOW, Makedate) ' roll back to sunday
         daybit = (2 ^ j)
                '1 = monday, 2 = tues etc
        If (daybit And DAYS) <> 0 Then
          ReportDays(j) = 1
          DailyReports(j) = DateAdd("d", j, Makedate)   '
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
         DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
         Makedate = DateAdd("d", -DOW, Makedate) ' roll back to sunday
         daybit = (2 ^ j)
                '1 = monday, 2 = tues etc
        If (daybit And DAYS) <> 0 Then

          ReportDays(j) = 1
          DailyReports(j) = DateAdd("d", j, Makedate)   '
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
         DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
         Makedate = DateAdd("d", -DOW, Makedate) ' roll back to sunday
         daybit = (2 ^ j)
                '1 = monday, 2 = tues etc
        If (daybit And DAYS) <> 0 Then

          ReportDays(j) = 1
          DailyReports(j) = DateAdd("d", j, Makedate)   '
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
        DOW = Weekday(Makedate, firstdayofweek:=vbSunday) - 1
         Makedate = DateAdd("d", -DOW, Makedate) ' roll back to sunday
         daybit = (2 ^ j)
                '1 = monday, 2 = tues etc
        If (daybit And DAYS) <> 0 Then


          ReportDays(j) = 1
          DailyReports(j) = DateAdd("d", j, Makedate)   '
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


  Case AUTOREPORT_WEEKLY  ' once a week at end of week
    If (NextReportDue = 0) Then
      Makedate = DateSerial(Year(CurrentDate), Month(CurrentDate), Day(CurrentDate))  ' 00 hrs
      CurrentDOW = Weekday(Makedate, firstdayofweek:=vbSunday)  ' returns 1 for sunday
      Makedate = DateAdd("d", 8 - CurrentDOW, Makedate)
      GetNextReportDate = Makedate
    End If

  Case AUTOREPORT_MONTHLY  ' once a month, at end of month
    If (NextReportDue = 0) Then
      Makedate = DateSerial(Year(Now), Month(Now), 1)  ' make first of current month
      Makedate = DateAdd("m", 1, Makedate)  ' add a month
      GetNextReportDate = Makedate  ' DateAdd("n", -1, Makedate)  ' ?? minus one minute
    End If
  Case Else

  End Select




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
  Dim j             As Integer
  Dim Room          As cRoom
  Dim dataWrapper   As cDataWrapper
  Dim Items()       As String

  On Error GoTo Parse_Error

  ResponseTime = Val(rs("ResponseTime") & "")    ' seconds for responding to alarm
  ResponseIsACK = Val(rs("ResponseIsACK") & "")  ' non-zero if ACK is response time
  ReportType = Val(rs("Reporttype") & "")

  ReportID = rs("reportid")
  Disabled = rs("Disabled")
  ReportName = rs("reportname") & ""
  Comment = rs("Comment") & ""
  RoomString = rs("Rooms") & ""
  EventString = rs("Events") & ""
  DevString = rs("DevTypes") & ""
  TimePeriod = rs("TimePeriod")
  DayPeriod = rs("DayPeriod")
  DAYS = rs("Days")
  Shift = rs("Shift")
  DayPartStart = rs("DayPartStart")
  DayPartEnd = rs("DayPartend")
  SortOrder = rs("SortOrder")
  SendHour = rs("SendHour")
  SaveAsFile = 1  '  rs("SaveAsFile")
  SendAsEmail = rs("SendAsEmail")
  recipient = rs("Recipient") & ""
  Subject = rs("Subject") & ""
  FileFormat = rs("FileFormat")
  DestFolder = rs("DestFolder") & ""



  Set Rooms = New Collection

  '  items = Split(RoomString, ",")
  '  For j = LBound(items) To UBound(items)
  '    Set Room = New cRoom
  '    Room.RoomID = Val(items(j))
  '    Rooms.Add Room
  '  Next

  Set Events = New Collection

  '  items = Split(EventString, ",")
  '  For j = LBound(items) To UBound(items)
  '    Set dataWrapper = New cDataWrapper
  '    dataWrapper.LongValue = Val(items(j))
  '    Events.Add dataWrapper
  '  Next

  Set DevTypes = New Collection

  Items = Split(EventString, vbTab)
  For j = LBound(Items) To UBound(Items)
    Set dataWrapper = New cDataWrapper
    dataWrapper.LongValue = Val(Items(j))
    DevTypes.Add dataWrapper
  Next

  Dim HasSecondShift As Boolean
  Dim HasThirdShift As Boolean

  If Configuration.EndFirst = Configuration.EndNight Then  ' no other shifts

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

  '
  Select Case DayPeriod
  Case AUTOREPORT_SHIFT1
    If HasThirdShift Then
      DayPartStart = Configuration.EndThird
      DayPartEnd = Configuration.EndFirst
    Else
      DayPartStart = Configuration.EndNight
      DayPartEnd = Configuration.EndFirst
    End If
  Case AUTOREPORT_SHIFT2
    DayPartStart = Configuration.EndFirst
    DayPartEnd = Configuration.EndNight
  Case AUTOREPORT_SHIFT3
    DayPartStart = Configuration.EndNight
    DayPartEnd = Configuration.EndThird
  End Select



Parse_Resume:

  On Error GoTo 0
  Exit Function

Parse_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionAutoReport.Parse." & Erl
  Resume Parse_Resume

End Function

Function UpdateData(ByVal rs As ADODB.Recordset) As Boolean

' convert Rooms collection to string of IDs

  Dim j             As Integer
  Dim Room          As cRoom
  Dim Items()       As String
  Dim dataWrapper   As cDataWrapper

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
  rs("SaveAsFile") = 1  '
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
  Set DevTypes = New Collection
  Set Events = New Collection
  Set EventItems = New Collection
End Sub

Private Sub Class_Terminate()
  Finalize
End Sub

Public Property Get Busy() As Boolean  ' Public read only
  Busy = mBusy
End Property


Function FindException(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim WhereClause        As String

        Dim SQL2               As String
        Dim Whereclause2       As String
        Dim rs2                As ADODB.Recordset


        Dim AlarmHistory       As cAlarmHistory

        Dim ExceptionLimit     As Double
        'Dim Devices()     As String
        Dim params()           As String
        Dim counter            As Long
        Dim models             As String
        Dim rsd                As ADODB.Recordset
        Dim devdesc            As String

        Dim Dev                As cESDevice

        Dim rc                 As Long
        Dim CLSPTI             As Long

10      On Error GoTo FindException_Error

20      Set rows = New Collection

30      ExceptionLimit = Me.ResponseTime * 60
        'EndDate = DateAdd("d", 1, EndDate)
40      models = Me.DevString



        Dim alarmtypes         As String
50      alarmtypes = EVT_EMERGENCY & "," & EVT_ALERT & "," & EVT_EXTERN

60      SQL = " SELECT Rooms.Room,  Alarms.id, Alarms.serial,alarms.eventdate, Alarms.alarm, Alarms.tamper,alarms.battery, " & _
              " alarms.residentid, alarms.roomid, alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst " & _
              " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
              " WHERE  ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
              " ORDER BY Alarms.id "
70      Set rs = ConnExecute(SQL)

80      rs.Filter = "eventtype = " & EVT_EMERGENCY & " OR eventtype = " & EVT_ALERT & " Or EventType = " & EVT_EXTERN

90      Whereclause2 = " WHERE  ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") ) "

100     SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as EventDate, Alarms_1.EventType as eventtype, Alarms_1.info as info, Alarms_1.UserName as username " & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
               " ORDER BY Alarms_1.alarmID,Alarms_1.ID; "

110     Set rs2 = ConnExecute(SQL2)

120     mSumElapsed = 0

130     Do Until rs.EOF

140       Set Dev = Nothing
150       Set Dev = Devices.Device(rs("Serial") & "")
160       If Not (Dev Is Nothing) Then
170         devdesc = Dev.Description
180         CLSPTI = Dev.CLSPTI      ' 45571 for door/window
190         If CLSPTI Then
200           If InStr(1, models, CStr(CLSPTI), vbTextCompare) Then

210             counter = counter + 1
220             If counter > 200 Then  ' just yield every n hits
230               counter = 0
240               DoEvents
250             End If


260               Set AlarmHistory = New cAlarmHistory
                  'Set AlarmHistory = GetAlarmHistory(rs)
270               AlarmHistory.ID = rs("ID")  ' this record's ID
280               AlarmHistory.Serial = rs("Serial") & ""
290               AlarmHistory.EventDate = rs("EventDate")
300               AlarmHistory.alarm = rs("Alarm")
310               AlarmHistory.Tamper = rs("Tamper")
320               AlarmHistory.Battery = rs("Battery")
330               AlarmHistory.Announce = rs("Announce") & ""
340               AlarmHistory.EventType = rs("EventType")
350               AlarmHistory.User = rs("Username") & ""
360               AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm
370               AlarmHistory.RoomID = rs("RoomID")


380               If Len(rs("NameLast") & "") Then
390                 If Len(rs("NameFirst") & "") Then
400                   AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
410                 Else
420                   AlarmHistory.Resident = rs("NameLast") & ""
430                 End If
440               ElseIf Len(rs("NameFirst") & "") Then
450                 AlarmHistory.Resident = rs("NameFirst") & ""
460               End If
470               AlarmHistory.Room = rs("room") & ""
480               rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

490               If AlarmHistory.ElapsedTime > ExceptionLimit Then
500                 rows.Add AlarmHistory

510                 Select Case AlarmHistory.EventType
                      Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
520                     If gElapsedEqACK = USE_ELAPSED_ACK Then

530                       If AlarmHistory.ACKElapsed > 0 Then
540                         mNumEntries = mNumEntries + 1
550                         mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
560                       End If
570                     Else
580                       If AlarmHistory.ElapsedTime > 0 Then
590                         mNumEntries = mNumEntries + 1
600                         mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
610                       End If
620                     End If
630                 End Select
640               End If
                
650           End If
660         End If
670       End If
680       rs.MoveNext
690     Loop
700     rs.Close
710     rs2.Close





FindException_Resume:

720     On Error GoTo 0

730     Set rs = Nothing
740     Set rs2 = Nothing
750     Exit Function

FindException_Error:

760     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionAutoReport.FindException." & Erl
770     Resume FindException_Resume

End Function


Function FindCount(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL           As String
        Dim rsEvents      As ADODB.Recordset
        Dim rsReport      As ADODB.Recordset
        'Dim AlarmHistory  As cAlarmHistory
        Dim WhereClause   As String
        Dim ExceptionLimit As Double
        Dim devdesc       As String
        Dim params()      As String
        Dim counter       As Long
        Dim models        As String
        Dim ExceptionReport As cExceptionReport

        Dim CLSPTI        As Long

        Dim CountData     As cCountData
        Dim modelList     As Collection

        Dim Dev           As cESDevice

        Dim CountTotal    As cCountData
10      On Error GoTo FindCount_Error

20      Set CountTotal = New cCountData

30      Set modelList = New Collection

40      Set rows = New Collection


50      models = Me.DevString
60      ExceptionLimit = Me.ResponseTime   ' not used
70      EndDate = DateAdd("d", 1, EndDate)


    Dim alarmtypes  As String
    alarmtypes = EVT_EMERGENCY & "," & EVT_ALERT & "," & EVT_EXTERN

'80      SQl = " SELECT Rooms.Room, Alarms.id, Alarms.serial,alarms.eventdate, Alarms.alarm, Alarms.tamper,alarms.battery, " & _
'              " alarms.residentid, alarms.roomid, alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst " & _
'              " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
'              " WHERE (alarms.eventtype in (" & alarmtypes & ")) AND  ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) "

        

80      SQL = " SELECT Rooms.Room, Alarms.id, Alarms.serial,alarms.eventdate, Alarms.alarm, Alarms.tamper,alarms.battery, " & _
              " alarms.residentid, alarms.roomid, alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst " & _
              " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
              " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) "

90      Set rsEvents = ConnExecute(SQL)

95      rsEvents.Filter = "eventtype = " & EVT_EMERGENCY & " OR eventtype = " & EVT_ALERT & " Or EventType = " & EVT_EXTERN

100     Do Until rsEvents.EOF

110       Set Dev = Nothing
120       Set Dev = Devices.Device(rsEvents("Serial") & "")
130       If Not (Dev Is Nothing) Then
140         devdesc = Dev.Description
150         CLSPTI = Dev.CLSPTI  ' 45571 for door/window
160         If CLSPTI Then
170           If InStr(1, models, CStr(CLSPTI), vbTextCompare) Then

180             counter = counter + 1
190             If counter Mod nDoEvents = 0 Then  ' just yield every n hits
200               DoEvents
210             End If

220             Set CountData = Nothing
230             On Error Resume Next
240             Set CountData = modelList(Dev.Model)
250             On Error GoTo 0
260             If CountData Is Nothing Then
270               Set CountData = New cCountData
280               CountData.Model = Dev.Model
290               CountData.Description = Dev.Description
300               CountData.CLSPTI = Dev.CLSPTI
310               CountData.Increment
320               modelList.Add CountData, CountData.Model

330             Else
340               CountData.Increment
350             End If

360           End If
370         End If
380       End If
390       rsEvents.MoveNext

400     Loop
410     rsEvents.Close
420     Set rsEvents = Nothing

430     For Each CountData In modelList
440       rows.Add CountData
450       CountTotal.Count = CountTotal.Count + CountData.Count
460     Next
470     CountTotal.Description = "Total"
480     rows.Add CountTotal

FindCount_Resume:

490     On Error GoTo 0
500     Exit Function

FindCount_Error:

510     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionAutoReport.FindCount." & Erl
520     Resume FindCount_Resume

End Function



Public Property Get NextReportDue() As Date

  NextReportDue = mNextReportDue

End Property

Public Property Let NextReportDue(ByVal NextReportDue As Date)

  mNextReportDue = NextReportDue

End Property
