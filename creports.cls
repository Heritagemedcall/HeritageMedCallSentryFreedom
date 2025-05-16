VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rows             As Collection
Public Dest              As Integer


Private hfile            As Integer
Const DEST_SCREEN = 0
Const DEST_PRINTER = 1
Const DEST_FILE = 2

Private COUNTBYN         As Long

Public lv                As ListView

Public Title             As String
Private ReportDate       As Date

Private ReportType       As String
Private alarmtypes       As String


Private mAverageElapsed  As Double
Private mNumEntries      As Long
Private mSumElapsed      As Double
Private mSumACKElapsed   As Double

Private mAverageElapsedString As String


Private mAverageFinalized  As Double
Private mNumFinalized      As Long
Private mSumFinalized      As Double
Private mAverageFinalizedString As String



Const COL_SERIAL = 0 ' not used
Const COL_ROOM = 0.25
Const COL_RESIDENT = 1.25
Const COL_EVENT = 2.75
Const COL_ALARMTIME = 5.5 ' R Justified

'Const COL_CANCELTIME = 7.5 ' not used
Const COL_ELAPSEDTIME = 6.45 ' R Justified
'Const COL_OPERATOR = 8.75
Const COL_RESPONDER = 6.5
Const COL_FINALIZED = 8.2 ' R Justified
Const COL_RESOLUTION = 8.25


Const LVCOL_ID = 0
Const LVCOL_ROOM = 1
Const LVCOL_RES = 2
Const LVCOL_EVENT = 3
Const LVCOL_ALARM = 4 ' alarm time
Const LVCOL_ELAPSED = 5
Const LVCOL_RESP = 6
Const LVCOL_FINALIZED = 7
Const LVCOL_FINALIZETEXT = 8

'
'Reports generation for Sentry Freedom 1 Software
'
'Reports generated should be available for view, print and/or file.
'Reports should show:
'ID-Room-Resident-Event type-Alarm time-Cancel time-Elapsed time-Software operator
'
'Standard Reports (requiring no user entry) (sorted by alarm activation time)
'
'  Shift report
'  Daily report
'  Weekly report (week begin/end to be determined in setup)
'  Monthly report
'
'Query Reports
'
'  By room and date range
'  By resident and date range
'  By device and date range
'  By event type and date range
'  By date range (all events)
'  Assurance by date
'  Device history(entered - deleted)
'  resident history(entered - deleted)(vacations)

Sub Fill(ByVal ReportID As String)
  If Dest = DEST_SCREEN Then
    ConfigureListView
  End If
  Select Case ReportID
    Case "currentshift"
      ReportType = "S"
      ShiftReport
      PrintReport
    Case "today"
      ReportType = "D"
      DailyReport
      PrintReport
    Case "thisweek"
      ReportType = "W"
      WeeklyReport
      PrintReport
    Case "thismonth"
      ReportType = "M"
      MonthlyReport
      Dim t As Long
      t = Win32.timeGetTime()
      
      PrintReport
      t = Win32.timeGetTime() - t
      Debug.Print "Fill Report Onscreen(Monthly report( took " & Format(t / 1000, "0.0") & " seconds"
      
      
    Case Else
      ReportType = ""
      DailyReport
      PrintReport
  End Select
End Sub

Sub PrintReport()
        Dim AlarmHistory       As cAlarmHistory
        Dim filename           As String
        Dim starttime          As Date
        Dim counter            As Long
10      starttime = Now
        Dim EndTime            As Date

20      On Error GoTo PrintReport_Error

30      Select Case Dest
          Case DEST_SCREEN           ' if To Screen ' Just add to listview
40          lv.ListItems.Clear
            Win32.LockWindowUpdate lv.hwnd
            
50        Case DEST_FILE             ' if To File ' one continuous file
60          hfile = FreeFile
70          filename = Replace(GetReportDateString & "_" & ReportType, "/", "-")
80          EnsurePathExists Configuration.ReportPath


90          Open Configuration.ReportPath & CleanFileName(GetReportDateString & "_" & ReportType) & ".txt" For Output As hfile

100       Case DEST_PRINTER          ' if To Printer
110         Printer.Orientation = vbPRORLandscape
120         Printer.Print " ";
130         Printer.ScaleMode = vbInches
140         Printer.FontSize = 10
150         Printer.FontBold = False

160     End Select

170     PrintBanner
180     PrintRowHeader
190     If Dest = DEST_SCREEN Then
200
210     End If

220     'Debug.Print "CReports.PrintReport Rows.count=" & rows.Count

        Dim begintime          As Date
230     begintime = Now

240     For Each AlarmHistory In rows

250       Select Case Dest
            Case DEST_SCREEN         ' if To Screen ' Just add to listview

260         Case DEST_PRINTER        ' if To Printer
270           'DoEvents
              ' check pagelength
280           If (Printer.ScaleHeight - Printer.CurrentY) < 0.5 Then  '  half inch bottom margin
290             PrintFooter
300             FormFeed Printer
310             PrintRowHeader
320           End If
330         Case DEST_FILE           ' if To File ' one continuous file
340       End Select

350       PrintRow AlarmHistory
360

          counter = counter + 1
          If (counter Mod 500) = 0 Then
            'Debug.Print "Counter  "; Counter; " " & Now
            counter = 0
            DoEvents
          End If
            

          'If DateDiff("s", begintime, Now) > 1 Then
370         'DoEvents
380         'begintime = Now
390       'End If

400     Next

410     Select Case Dest
          Case DEST_SCREEN           ' if To Screen ' Just add to listview
420         Win32.LockWindowUpdate 0
430       Case DEST_PRINTER          ' if To Printer
440         PrintFooter
450         Printer.EndDoc
460         Printer.Orientation = vbPRORPortrait
470       Case DEST_FILE             ' if To File ' one continuous file
480         If hfile <> 0 Then
490           Close hfile
500           hfile = 0
510         End If
520     End Select

PrintReport_Resume:

530     Debug.Print DateDiff("s", starttime, Now) & " Seconds to run report"


540     On Error GoTo 0
550     Exit Sub

PrintReport_Error:

560     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cReports.PrintReport." & Erl
570     Resume PrintReport_Resume


End Sub


Sub PrintRow(AlarmHistory As cAlarmHistory)
  ' ID Room Resident Eventtype Alarmtime Canceltime ElapsedTime Username
  ' tab delimited
  Dim li                 As ListItem
  Select Case Dest
    Case DEST_SCREEN           ' if To Screen:
      If Not lv Is Nothing Then
        If (Len(AlarmHistory.Extra) > 0) Then  ' this is average record
          Set li = lv.ListItems.Add(, "0" & "S" & lv.ListItems.Count, "")
          'If gElapsedEqACK = 1 Then
          li.SubItems(LVCOL_ROOM) = "Average"
          If (AlarmHistory.Count > 0) Then

            li.SubItems(LVCOL_ELAPSED) = AlarmHistory.AverageTimeToString()
          Else
            li.SubItems(LVCOL_ELAPSED) = "N/A"
          End If

          If (AlarmHistory.Count2 > 0) Then

            li.SubItems(LVCOL_FINALIZED) = AlarmHistory.AverageTimeToString2()
          Else
            li.SubItems(LVCOL_FINALIZED) = "N/A"
          End If


          
          

        Else
          Set li = lv.ListItems.Add(, AlarmHistory.ResidentID & "S" & lv.ListItems.Count, AlarmHistory.Serial)
          li.SubItems(LVCOL_ROOM) = AlarmHistory.Room
          li.SubItems(LVCOL_RES) = AlarmHistory.Resident



          Select Case AlarmHistory.EventType
            Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
              li.SubItems(LVCOL_EVENT) = AlarmHistory.Announce
            Case EVT_ASSISTANCE
              li.SubItems(LVCOL_EVENT) = AlarmHistory.Announce
            Case Else
              li.SubItems(LVCOL_EVENT) = AlarmHistory.EventName
          End Select

          li.SubItems(LVCOL_ALARM) = AlarmHistory.AlarmTimeToString



          If gElapsedEqACK = 1 Then
            If AlarmHistory.ACKTime > 0 Then
              'li.SubItems(LVCOL_ELAPSED) = AlarmHistory.ACKTimeToString
              li.SubItems(LVCOL_ELAPSED) = AlarmHistory.ACKElapsedTimeToString

            Else               ' try and use end of alarm time
              'li.SubItems(LVCOL_ELAPSED) = AlarmHistory.CancelTimeToString
              li.SubItems(LVCOL_ELAPSED) = AlarmHistory.ElapsedTimeToString
            End If
          Else
            'li.SubItems(LVCOL_ELAPSED) = AlarmHistory.CancelTimeToString
            li.SubItems(LVCOL_ELAPSED) = AlarmHistory.ElapsedTimeToString
          End If
          
          li.SubItems(LVCOL_RESP) = AlarmHistory.Responder
          li.SubItems(LVCOL_FINALIZED) = AlarmHistory.FinalizeTimeToString
          li.SubItems(LVCOL_FINALIZETEXT) = AlarmHistory.FinalizeText
          
        End If
      End If


    Case DEST_PRINTER          ' if To Printer:
      DoEvents

      If (Len(AlarmHistory.Extra) > 0) Then  ' this is average record
        PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Average"
        'If gElapsedEqACK = 1 Then
        'If (mNumEntries) Then
        If (AlarmHistory.Count > 0) Then

          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.AverageTimeToString()
        Else
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "N/A"
        End If
        
        
        
        'If (mNumFinalized) Then
        If (AlarmHistory.Count2 > 0) Then

          PrintRJ Printer, COL_FINALIZED, Printer.CurrentY, AlarmHistory.AverageTimeToString2()
        Else
          PrintRJ Printer, COL_FINALIZED, Printer.CurrentY, "N/A"
        End If
        
        
        'Else

        'End If


      Else


        'PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial  ' need apostrophy for excel to like it
        PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
        PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident

        Dim EventCol     As String

        Select Case AlarmHistory.EventType
          Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
            EventCol = AlarmHistory.Announce
          Case EVT_ASSISTANCE
            EventCol = AlarmHistory.Announce
          Case Else
            EventCol = AlarmHistory.EventName
        End Select

        Do While Printer.TextWidth(EventCol) > 1.7
          EventCol = left$(EventCol, Len(EventCol) - 1)
        Loop
        Select Case AlarmHistory.EventType
          Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, EventCol
          Case EVT_ASSISTANCE
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, EventCol

          Case Else
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, EventCol
        End Select
        PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString
        'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, AlarmHistory.CancelTimeToString
        PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
        PrintLJ Printer, COL_RESPONDER, Printer.CurrentY, AlarmHistory.Responder
        PrintRJ Printer, COL_FINALIZED, Printer.CurrentY, AlarmHistory.FinalizeTimeToString
        
        Dim finaltextstring As String
        finaltextstring = AlarmHistory.FinalizeText
        
        Do While Printer.TextWidth(finaltextstring) > 2.5
          finaltextstring = left$(finaltextstring, Len(finaltextstring) - 1)
        Loop
        
        PrintLJ Printer, COL_RESOLUTION, Printer.CurrentY, finaltextstring
        Printer.Print " "
      End If
    Case DEST_FILE             ' if To File
      If hfile <> 0 Then
        'If gElapsedEqACK = 1 Then

        If (Len(AlarmHistory.Extra) > 0) Then  ' this is average record
          'PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Average"
          'If gElapsedEqACK = 1 Then
          'If (AlarmHistory.Count) Then

            Print #hfile, Join(Array("Average", "", "", "", "", "", "", AlarmHistory.AverageTimeToString(), "", "", AlarmHistory.AverageTimeToString2()), vbTab)
          'Else
            'Print #hfile, Join(Array("Average", "", "", "", "", "", "", "N/A", ""), vbTab)
          'End If
        Else
          Print #hfile, AlarmHistory.ToTABBEDString
        End If
      End If
  End Select

End Sub
Sub PrintRowHeader()
        ' ID Room Resident Eventtype Alarmtime Canceltime ElapsedTime Username
        ' tab delimited

10      Select Case Dest
          Case DEST_SCREEN  ' if To Screen:
20          ConfigureListView


30        Case DEST_PRINTER  ' if To Printer:
40          'PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
50          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
60          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
70          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
80          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Alarm Time"
90          'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, "Cancel Time"
100         PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "Elapsed Time"
110         PrintLJ Printer, COL_RESPONDER, Printer.CurrentY, "Response By"
111         PrintRJ Printer, COL_FINALIZED, Printer.CurrentY, "Finalized"
112         PrintLJ Printer, COL_RESOLUTION, Printer.CurrentY, "Resolution"
120         Printer.Print " "
130         HR Printer

140       Case DEST_FILE  ' if To File
150         If hfile <> 0 Then
160           Print #hfile, Join(Array("Room", "Resident", "Event", "Announce", "Alarm Time", "Elapsed Time", "Response By", "Finalized", "Resolution"), vbTab)
              'Print #hfile, Join(Array("TXID", "Room", "Resident", "Event", "Announce", "Alarm Time", "Cancel Time", "Elapsed Time", "Operator"), vbTab)
170         End If
180     End Select


End Sub
Sub ConfigureListView()
        Dim ch As ColumnHeader
10      If Not lv Is Nothing Then
20        lv.ColumnHeaders.Clear
30        lv.ColumnHeaders.Add , , "ID", 0 ' 0
40        lv.ColumnHeaders.Add , , "Room", 1400 ' 1
50        lv.ColumnHeaders.Add , , "Resident", 1600 '2
60        lv.ColumnHeaders.Add , , "Event", 1300 '3
70        lv.ColumnHeaders.Add , , "Alarm", 1500 '4
'80        lv.ColumnHeaders.Add , , "Cancel", 1300
90        Set ch = lv.ColumnHeaders.Add(, , "Elapsed", 950) ' d:hh:10:00 ' d:hh:nn:ss '5
          ch.Alignment = lvwColumnRight
'100       lv.ColumnHeaders.Add , , "Operator", 900
100       lv.ColumnHeaders.Add , , "Resp By", 950 '6
105       Set ch = lv.ColumnHeaders.Add(, , "Finalized", 950) '7
          ch.Alignment = lvwColumnRight
          lv.ColumnHeaders.Add , , "Resolution", 1500 '8
110     End If
120     lv.Refresh

End Sub

Sub PrintBanner()
        ' date and time of report
10      Select Case Dest
          Case DEST_SCREEN  ' if To Screen:
20        Case DEST_PRINTER  ' if To Printer:


30          Printer.CurrentX = 0
40          Printer.CurrentY = 0
50          Printer.ScaleMode = vbInches

60          Printer.Print ""
70          PrintCentered Printer, Printer.CurrentY, Title
80          PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "")   ' hh:nn am/pm")

90          Printer.Print ""
100         Printer.Print ""


110       Case DEST_FILE  ' if To File

120     End Select
End Sub
Sub PrintFooter()
        ' copyright etc
10      Select Case Dest
          Case DEST_SCREEN  ' if To Screen:
20        Case DEST_PRINTER  ' if To Printer:
            Dim text           As String
30          text = App.LegalCopyright & " " & App.ProductName & " (Ver. 1." & Format(App.Revision, "000") & ")"
40          PrintCJ Printer, Printer.ScaleWidth / 2, Printer.ScaleHeight - Printer.TextHeight("A") * 1.5, text

50        Case DEST_FILE  ' if To File

60      End Select


End Sub

Sub DailyReport()
  Dim SQL                As String
  Dim WhereClause        As String
  Dim rs                 As Recordset
  Dim AlarmHistory       As cAlarmHistory

  Dim rshist             As Recordset

  Dim CurrentDate        As Date

  Dim BasicEventTypes()  As String
  Dim filterstring       As String
  Dim i                  As Long

  Dim rs2                As Recordset
  Dim SQL2               As String
  Dim Whereclause2       As String


  ReportDate = Now

  CurrentDate = Fix(Now)
  Title = "Current Day Report"

  mAverageElapsed = 0
  mNumEntries = 0
  mSumElapsed = 0

  mAverageFinalized = 0
  mNumFinalized = 0
  mSumFinalized = 0
  mAverageFinalizedString = ""


  ' select all alarms that match criteria with basic info
  ' Walk each alarm
  ' Fetch each event for each alarm
  ' Only get initiating events, history comes later


  WhereClause = " AND (EventDate >= " & DateDelim & CurrentDate & DateDelim & " ) "
  'SQl = "SELECT ID, serial, eventdate, alarm, tamper, battery,residentid, roomid, announce, eventtype, username , alarmid FROM alarms WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"

  SQL = "SELECT Alarms.ID, Alarms.Serial, Alarms.EventDate, Alarms.Alarm, Alarms.Tamper, Alarms.Battery, Alarms.Announce, Alarms.EventType, Alarms.UserName, Alarms.AlarmID, Residents.NameLast, Residents.NameFirst, Rooms.Room  ,Devices.Model " & _
        " FROM Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID left JOIN Rooms ON Alarms.RoomID = Rooms.RoomID  LEFT JOIN Devices ON Alarms.Serial = Devices.Serial  WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"


  Dim t                  As Date
  t = Now

  Set rs = ConnExecute(SQL)
  dbg "Query " & DateDiff("s", t, Now)
  Debug.Print "Query " & DateDiff("s", t, Now)

  'Filter by selected EventTypes (new 3/5/2014 for trying to speed queries that intermittently slow to a crawl)

  'alarmtypes = Join(Array(EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN), ",")

  BasicEventTypes = Split(alarmtypes, ",")

  For i = LBound(BasicEventTypes) To UBound(BasicEventTypes)
    BasicEventTypes(i) = "eventtype = " & BasicEventTypes(i)
  Next

  filterstring = Join(BasicEventTypes, " or ")
  t = Now
  rs.Filter = filterstring



  Debug.Print "Filter " & DateDiff("s", t, Now)


  'WhereClause = " EVENT ID = 0 AND (EventDate >= " & DateDelim & CurrentDate & DateDelim & " ) "
  'SQl = "SELECT ID, serial, eventdate, alarm, tamper, battery, residentid, roomid, announce, eventtype, username , alarmid FROM alarms WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"
  t = Now
  'WhereClause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & " )  "
  Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & CurrentDate & DateDelim & " ) "

  SQL2 = " SELECT Alarms.ID, Alarms.Serial,  Alarms_1.AlarmID, Alarms_1.EventDate as eventdate, Alarms_1.EventType as eventtype,Alarms_1.info as info, Alarms_1.UserName as username " & _
         " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
         Whereclause2 & _
         " ORDER BY Alarms_1.AlarmID,Alarms.ID; "

  Set rs2 = ConnExecute(SQL2)

  Dim RowCount           As Long
  Dim rc                 As Long

  Do Until rs.EOF
    RowCount = RowCount + 1
    'mNumEntries = mNumEntries + 1

    COUNTBYN = COUNTBYN + 1

    If COUNTBYN > 200 Then
      COUNTBYN = 0
      DoEvents
    End If
    Dim CLSPTI           As Long
    '          Dim device           As cESDevice
    '320       CLSPTI = 0
    '330       Set device = Devices.device(Rs("Serial") & "")
    '340       If Not device Is Nothing Then
    '350         CLSPTI = device.CLSPTI
    '360       End If

    CLSPTI = Val(GetCLSPTI(rs("Model") & ""))
    Set AlarmHistory = New cAlarmHistory



    AlarmHistory.ID = rs("ID")  ' this record's ID
    AlarmHistory.Serial = rs("Serial") & ""
    AlarmHistory.EventDate = rs("EventDate")
    AlarmHistory.alarm = rs("Alarm")
    AlarmHistory.Tamper = rs("Tamper")
    AlarmHistory.Battery = rs("Battery")
    'ResidentID = rs("residentID")
    'RoomID = rs("RoomID")
    AlarmHistory.Announce = rs("Announce") & ""
    AlarmHistory.EventType = rs("EventType")
    AlarmHistory.User = rs("Username") & ""
    AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm

    If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
      AlarmHistory.Resident = ""
    Else


      If Len(rs("NameLast") & "") Then
        If Len(rs("NameFirst") & "") Then
          AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
        Else
          AlarmHistory.Resident = rs("NameLast") & ""
        End If
      ElseIf Len(rs("NameFirst") & "") Then
        AlarmHistory.Resident = rs("NameFirst") & ""
      End If

    End If

    AlarmHistory.Room = rs("room") & ""

    rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

    'mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
    'mSumACKElapsed = mSumACKElapsed + AlarmHistory.ACKElapsed


    Select Case rs("eventtype")

      Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

        If CLSPTI <> &HEEEE& Then  ' 61166 dec
          If gElapsedEqACK = 1 Then
            If AlarmHistory.ACKElapsed > 0 Then
              mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
              mNumEntries = mNumEntries + 1
            End If
          ElseIf AlarmHistory.ElapsedTime > 0 Then
            mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
            mNumEntries = mNumEntries + 1
          End If

          If AlarmHistory.FinalizeElapsed > 0 Then
            mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
            mNumFinalized = mNumFinalized + 1
          End If

        End If

      Case Else

    End Select

    rows.Add AlarmHistory

    rs.MoveNext
  Loop

  ' adds an extra alarm history as averages

  Set AlarmHistory = New cAlarmHistory
  AlarmHistory.Extra = "Average"
  AlarmHistory.ACKElapsed = mSumACKElapsed
  AlarmHistory.ElapsedTime = mSumElapsed
  AlarmHistory.FinalizeElapsed = mSumFinalized
  AlarmHistory.Count = mNumEntries
  AlarmHistory.Count2 = mNumFinalized
  If mNumEntries Then
    AlarmHistory.Average = mSumElapsed / mNumEntries
  End If
  If mNumFinalized > 0 Then
    AlarmHistory.Average2 = mSumFinalized / mNumFinalized
  End If

  rows.Add AlarmHistory


  dbg "Rows " & RowCount

  rs2.Close
  Set rs2 = Nothing
  rs.Close
  Set rs = Nothing


End Sub


Sub ShiftReport()
        ' determine time/date for current shift
        Dim SQL                As String
        Dim WhereClause        As String
        Dim rs                 As Recordset
        Dim AlarmHistory       As cAlarmHistory

        Dim CurrentDate        As Date


        Dim ShiftStart         As Date
        Dim ShiftEnd           As Date
        Dim Shift              As Integer  ' shift used to be wither 0 or 1, now 5/10/12 we need another shift


        Dim SecStart           As Long
        Dim SecEnd             As Long
        Dim SecondsSinceMidnight As Long

        Dim HasSecondShift     As Boolean
        Dim HasThirdShift      As Boolean

10      If Configuration.EndFirst = Configuration.EndNight Then
20        HasSecondShift = False
30        HasThirdShift = False
40      ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
50        HasSecondShift = True
60        HasThirdShift = False
70      ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
80        HasSecondShift = True
90        HasThirdShift = True
100     Else
110       HasSecondShift = False
120       HasThirdShift = False
130     End If

140     ReportDate = Now

150     SecondsSinceMidnight = Timer()  ' seconds since midnight



160     Title = "Current Shift Report "
170     Shift = GetCurrentShift()

180     If Configuration.EndNight = Configuration.StartNight And Configuration.EndNight = Configuration.EndThird Then  ' only one shift
          ' just like (to)day report
190       CurrentDate = Format(ReportDate, "mm/dd/yyyy")
200       WhereClause = " AND (EventDate >= " & DateDelim & CurrentDate & DateDelim & " )  "
210     Else
          ' zzzz
220       Select Case Shift
            Case SHIFT_DAY
230           If HasThirdShift Then  ' first,second, and third shifts
240             SecStart = Configuration.EndThird * SECONDSPERHOUR
250             If SecondsSinceMidnight < SecStart Then  ' if now is before start then start was yesterday
260               ShiftStart = DateAdd("s", SecStart, Format(Now - 1, "mm/dd/yyyy"))
270             Else
280               ShiftStart = DateAdd("s", SecStart, Format(Now, "mm/dd/yyyy"))
290             End If
300             WhereClause = " AND (EventDate >= " & DateDelim & ShiftStart & DateDelim & " )  "
310             WhereClause = WhereClause & " AND (EventDate <= " & DateDelim & Now & DateDelim & " )  "
320           ElseIf HasSecondShift Then  ' first and second shift only
330             SecStart = Configuration.EndNight * SECONDSPERHOUR
340             If SecondsSinceMidnight < SecStart Then  ' if now is before start then start was yesterday
350               ShiftStart = DateAdd("s", SecStart, Format(Now - 1, "mm/dd/yyyy"))
360             Else
370               ShiftStart = DateAdd("s", SecStart, Format(Now, "mm/dd/yyyy"))
380             End If
390             WhereClause = " AND (EventDate >= " & DateDelim & ShiftStart & DateDelim & " )  "
400             WhereClause = WhereClause & " AND (EventDate <= " & DateDelim & Now & DateDelim & " )  "
410           Else                   ' only first shift
420             CurrentDate = Format(ReportDate, "mm/dd/yyyy")
430             WhereClause = " AND (EventDate >= " & DateDelim & CurrentDate & DateDelim & " )  "

440           End If

450         Case SHIFT_NIGHT

460           SecStart = Configuration.EndFirst * SECONDSPERHOUR
470           If SecondsSinceMidnight < SecStart Then  ' if now is before start then start was yesterday
480             ShiftStart = DateAdd("s", SecStart, Format(Now - 1, "mm/dd/yyyy"))
490           Else
500             ShiftStart = DateAdd("s", SecStart, Format(Now, "mm/dd/yyyy"))
510           End If
520           WhereClause = " AND (EventDate >= " & DateDelim & ShiftStart & DateDelim & " )  "
530           WhereClause = WhereClause & " AND (EventDate <= " & DateDelim & Now & DateDelim & " )  "

540         Case SHIFT_GRAVE
550           SecStart = Configuration.EndNight * SECONDSPERHOUR
560           If SecondsSinceMidnight < SecStart Then  ' if now is before start then start was yesterday
570             ShiftStart = DateAdd("s", SecStart, Format(Now - 1, "mm/dd/yyyy"))
580           Else
590             ShiftStart = DateAdd("s", SecStart, Format(Now, "mm/dd/yyyy"))
600           End If
610           WhereClause = " AND (EventDate >= " & DateDelim & ShiftStart & DateDelim & " )  "
620           WhereClause = WhereClause & " AND (EventDate <= " & DateDelim & Now & DateDelim & " )  "


630       End Select

640     End If

        ' SQL = "SELECT ID, serial, eventdate, alarm, tamper, battery,residentid, roomid, announce, eventtype, username , alarmid FROM alarms WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"

650     SQL = "SELECT Alarms.ID, Alarms.Serial, Alarms.EventDate, Alarms.Alarm, Alarms.Tamper, Alarms.Battery, Alarms.Announce, Alarms.EventType, Alarms.UserName, Alarms.AlarmID, Residents.NameLast, Residents.NameFirst, Rooms.Room  ,Devices.Model " & _
              " FROM (Alarms left JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) left JOIN Rooms ON Alarms.RoomID = Rooms.RoomID  LEFT JOIN Devices ON Alarms.Serial = Devices.Serial  WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"



660     Set rs = ConnExecute(SQL)

        'Filter by selected EventTypes (new 3/5/2014 for trying to speed queries that intermittently slow to a crawl)
        Dim BasicEventTypes()  As String
        Dim filterstring       As String
        Dim i                  As Long



        'alarmtypes = Join(Array(EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN), ",")

670     BasicEventTypes = Split(alarmtypes, ",")




680     For i = LBound(BasicEventTypes) To UBound(BasicEventTypes)
690       BasicEventTypes(i) = "eventtype = " & BasicEventTypes(i)
700     Next



710     filterstring = Join(BasicEventTypes, " or ")

720     rs.Filter = filterstring

        Dim RowCount           As Long
        Dim rc                 As Long


        Dim rs2                As Recordset
        Dim SQL2               As String
        Dim Whereclause2       As String

730     Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & CurrentDate & DateDelim & " ) "

740     SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as eventdate,Alarms_1.info as info, Alarms_1.EventType as eventtype, Alarms_1.UserName as username" & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
               " ORDER BY  Alarms_1.AlarmID,Alarms.ID; "

750     Set rs2 = ConnExecute(SQL2)

760     mAverageElapsed = 0
770     mNumEntries = 0
780     mSumElapsed = 0


790     mAverageFinalized = 0
800     mNumFinalized = 0
810     mSumFinalized = 0
820     mAverageFinalizedString = ""


830     Do Until rs.EOF

840       COUNTBYN = COUNTBYN + 1
850       If COUNTBYN > 200 Then
860         COUNTBYN = 0
870         DoEvents
880       End If

          Dim CLSPTI           As Long
          '          Dim device           As cESDevice
          '860       CLSPTI = 0
          '870       Set device = Devices.device(Rs("Serial") & "")
          '880       If Not device Is Nothing Then
          '890         CLSPTI = device.CLSPTI
          '900       End If

890       CLSPTI = Val(GetCLSPTI(rs("Model") & ""))


900       Set AlarmHistory = New cAlarmHistory



910       AlarmHistory.ID = rs("ID")  ' this record's ID
920       AlarmHistory.Serial = rs("Serial") & ""
930       AlarmHistory.EventDate = rs("EventDate")
940       AlarmHistory.alarm = rs("Alarm")
950       AlarmHistory.Tamper = rs("Tamper")
960       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
970       AlarmHistory.Announce = rs("Announce") & ""
980       AlarmHistory.EventType = rs("EventType")
990       AlarmHistory.User = rs("Username") & ""
1000      AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm


1010      If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
1020        AlarmHistory.Resident = ""
1030      Else

1040        If Len(rs("NameLast") & "") Then
1050          If Len(rs("NameFirst") & "") Then
1060            AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
1070          Else
1080            AlarmHistory.Resident = rs("NameLast") & ""
1090          End If
1100        ElseIf Len(rs("NameFirst") & "") Then
1110          AlarmHistory.Resident = rs("NameFirst") & ""
1120        End If

1130      End If

1140      AlarmHistory.Room = rs("room") & ""

1150      rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)



1160      Select Case rs("eventtype")



            Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

1170          If CLSPTI <> &HEEEE& Then  ' 61166 dec

                'If AlarmHistory.ElapsedTime > 0 Or AlarmHistory.ACKElapsed > 0 Then
1180            If gElapsedEqACK = 1 Then
1190              If AlarmHistory.ACKElapsed > 0 Then
1200                mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
1210                mNumEntries = mNumEntries + 1
1220              End If
1230            ElseIf AlarmHistory.ElapsedTime > 0 Then
1240              mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
1250              mNumEntries = mNumEntries + 1
1260            End If

               If AlarmHistory.FinalizeElapsed > 0 Then
                  mNumFinalized = mNumFinalized + 1
                  mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                End If


1270          End If
1280        Case Else

1290      End Select

          '1120      mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
          '1130      mSumACKElapsed = mSumACKElapsed + AlarmHistory.ACKElapsed

          '          If AlarmHistory.ElapsedTime > 0 Or AlarmHistory.ACKElapsed > 0 Then
          '            mNumEntries = mNumEntries + 1
          '          End If

1300      rows.Add AlarmHistory


1310      rs.MoveNext
1320    Loop

  Set AlarmHistory = New cAlarmHistory
  AlarmHistory.Extra = "Average"
  AlarmHistory.ACKElapsed = mSumACKElapsed
  AlarmHistory.ElapsedTime = mSumElapsed
  AlarmHistory.FinalizeElapsed = mSumFinalized
  AlarmHistory.Count = mNumEntries
  AlarmHistory.Count2 = mNumFinalized
  If mNumEntries Then
    AlarmHistory.Average = mSumElapsed / mNumEntries
  End If
  If mNumFinalized > 0 Then
    AlarmHistory.Average2 = mSumFinalized / mNumFinalized
  End If


1460    rows.Add AlarmHistory

1470    On Error Resume Next
1480    rs2.Close
1490    Set rs2 = Nothing


1500    rs.Close
1510    Set rs = Nothing


End Sub

'Function GetAlarmHistory(rs As Recordset) As cAlarmHistory
'  Dim History            As cAlarmHistory
'
'  Set History = New cAlarmHistory
'  History.Parse rs
'
'  History.GetHistory conn
'  Set GetAlarmHistory = History
'
'
'End Function

'Function GetAlarmHistory2(rs As Recordset) As cAlarmHistory
'  Dim History As cAlarmHistory
'
'  Set History = New cAlarmHistory
'  History.Parse rs
'  If History.ResidentID <> 0 Then
'    History.Resident = GetResidentName(History.ResidentID)
'  End If
'  If History.RoomID <> 0 Then
'    History.Room = GetRoomName(History.RoomID)
'  End If
'  History.GetHistory conn
'  Set GetAlarmHistory2 = History
'
'
'End Function

Sub WeeklyReport()
        Dim SQL                As String
        Dim WhereClause        As String
        Dim rs                 As Recordset
        Dim AlarmHistory       As cAlarmHistory

        Dim StartDate          As Date
        Dim CurrentDay         As Integer

        Dim BasicEventTypes()  As String
        Dim filterstring       As String
        Dim i                  As Long
        Dim counter            As Long

10      ReportDate = Now

20      Title = "Current Week Report"

30      StartDate = Fix(Now)
40      CurrentDay = Weekday(StartDate, vbSunday)
50      Select Case CurrentDay
          Case vbSunday
            ' no change
60        Case Else
70          StartDate = DateAdd("d", -CurrentDay, StartDate)
80      End Select

        '  WhereClause = " AND (EventDate >= " & DateDelim & StartDate & DateDelim & " ) ) "

        ' select all alarms that match criteria with basic info
        ' Walk each alarm
        ' Fetch each event for each alarm
        ' Only get initiating events, history comes later


        Dim t                  As Date
90      BasicEventTypes = Split(alarmtypes, ",")
100     For i = LBound(BasicEventTypes) To UBound(BasicEventTypes)
110       BasicEventTypes(i) = "eventtype = " & BasicEventTypes(i)
120     Next

130     filterstring = Join(BasicEventTypes, " or ")


140     WhereClause = " AND (EventDate >= " & DateDelim & StartDate & DateDelim & " )  "


        ' this SQL gets the initial alarm and associated data
150     SQL = "SELECT Alarms.ID, Alarms.Serial, Alarms.EventDate, Alarms.Alarm, Alarms.Tamper, Alarms.Battery, Alarms.Announce, Alarms.EventType, Alarms.UserName, Alarms.AlarmID, Residents.NameLast, Residents.NameFirst, Rooms.Room  ,Devices.Model " & _
              " FROM (Alarms left JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) left JOIN Rooms ON Alarms.RoomID = Rooms.RoomID  LEFT JOIN Devices ON Alarms.Serial = Devices.Serial   WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"

160     t = Now

        
170     Set rs = ConnExecute(SQL)

180     Debug.Print "Query " & DateDiff("s", t, Now)
        'Filter by selected EventTypes (new 3/5/2014 for trying to speed queries that intermittently slow to a crawl)
        ' eventtype = 1 or eventtype = 14 or eventtype = 19 or eventtype = 4 or eventtype = 6 or eventtype = 10 or eventtype = 48 or eventtype = 52

190     rs.Filter = filterstring

200     Debug.Print "Filter " & DateDiff("s", t, Now)

        Dim rs2                As Recordset
        Dim SQL2               As String
        Dim Whereclause2       As String

        'info column is resolution text 2019-02-13

210     Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & " )  "

220     SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as EventDate,Alarms_1.info as info,  Alarms_1.EventType as EventType, Alarms_1.UserName as UserName " & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
               " ORDER BY Alarms_1.AlarmID,Alarms.ID; "

        ' SQL2
        ' SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as EventDate, Alarms_1.EventType as EventType, Alarms_1.UserName as UserName  FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID  WHERE (alarms.EventDate >= '02/09/19' )   ORDER BY Alarms_1.AlarmID,Alarms.ID;


'
'Global Const EVT_EMERGENCY_RESPOND = 66
'Global Const EVT_ALERT_RESPOND = 67
'Global Const EVT_GENERIC_RESPOND = 68
'Global Const EVT_EXTERN_RESPOND = 69
'
'Global Const EVT_EMERGENCY_FINALIZE = 70
'Global Const EVT_ALERT_FINALIZE = 71
'Global Const EVT_GENERIC_FINALIZE = 72
'Global Const EVT_EXTERN_FINALIZE = 73
'
'
'Global Const EVT_ASSISTANCE = 74
'Global Const EVT_ASSISTANCE_RESPOND = 75
'Global Const EVT_ASSISTANCE_FINALIZE = 76
'Global Const EVT_ASSISTANCE_ACK = 77
'Global Const EVT_ASSISTANCE_RESTORE = 78
'Global Const EVT_ASSISTANCE_END = 79



'Global Const EVT_EMERGENCY_RESPOND = 66
'Global Const EVT_ALERT_RESPOND = 67
'Global Const EVT_ASSISTANCE_RESPOND = 75

'Global Const EVT_EMERGENCY_FINALIZE = 70
'Global Const EVT_ALERT_FINALIZE = 71
'Global Const EVT_ASSISTANCE_FINALIZE = 76

        ' Debug.Assert 0
230     Set rs2 = ConnExecute(SQL2)

        Dim RowCount           As Long
        Dim rc                 As Long

240     mAverageElapsed = 0
250     mNumEntries = 0
260     mSumElapsed = 0

        mAverageFinalized = 0
        mNumFinalized = 0
        mSumFinalized = 0
        mAverageFinalizedString = ""



270     Do Until rs.EOF
280       RowCount = RowCount + 1
290       counter = counter + 1
          Debug.Print "Weekly EventType " & rs("EventType") & " " & "ID " & rs("ID")
          
          'mNumEntries = mNumEntries + 1
300       If counter > 200 Then
310         counter = 0
320         DoEvents
330       End If

          Dim CLSPTI           As Long
'          Dim device           As cESDevice
'340       CLSPTI = 0
'350       Set device = Devices.device(Rs("Serial") & "")
'360       If Not device Is Nothing Then
'
'370         CLSPTI = device.CLSPTI
'            'If CLSPTI = &HEEEE& Then ' 61166 dec
'            'Debug.Assert 0
'            'End If
'
'380       End If

          CLSPTI = Val(GetCLSPTI(rs("Model") & ""))


390       Set AlarmHistory = New cAlarmHistory



400       AlarmHistory.ID = rs("ID")  ' this record's ID
410       AlarmHistory.Serial = rs("Serial") & ""
420       AlarmHistory.EventDate = rs("EventDate")
430       AlarmHistory.alarm = rs("Alarm")
440       AlarmHistory.Tamper = rs("Tamper")
450       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
460       AlarmHistory.Announce = rs("Announce") & ""
470       AlarmHistory.EventType = rs("EventType")
480       AlarmHistory.User = rs("Username") & ""
490       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm


500       If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
510         AlarmHistory.Resident = ""
520       Else

530         If Len(rs("NameLast") & "") Then
540           If Len(rs("NameFirst") & "") Then
550             AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
560           Else
570             AlarmHistory.Resident = rs("NameLast") & ""
580           End If
590         ElseIf Len(rs("NameFirst") & "") Then
600           AlarmHistory.Resident = rs("NameFirst") & ""
610         End If
620       End If
630       AlarmHistory.Room = rs("room") & ""

640       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)


650       Select Case AlarmHistory.EventType

            Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

660           If CLSPTI <> &HEEEE& Then  ' 61166 dec

670             If gElapsedEqACK = 1 Then
680               If AlarmHistory.ACKElapsed > 0 Then
690                 mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
700                 mNumEntries = mNumEntries + 1
710               End If
720             ElseIf AlarmHistory.ElapsedTime > 0 Then
730               mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
740               mNumEntries = mNumEntries + 1
750             End If

               If AlarmHistory.FinalizeElapsed > 0 Then
                  mNumFinalized = mNumFinalized + 1
                  mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                End If


760           End If





              '        If AlarmHistory.ElapsedTime > 0 Or AlarmHistory.ACKElapsed > 0 Then
              '          mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
              '          mSumACKElapsed = mSumACKElapsed + AlarmHistory.ACKElapsed
              '          mNumEntries = mNumEntries + 1
              '        End If
              '




770         Case Else

780       End Select

          '610       mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
          '620       mSumACKElapsed = mSumACKElapsed + AlarmHistory.ACKElapsed
          '
          '          If AlarmHistory.ElapsedTime > 0 Or AlarmHistory.ACKElapsed > 0 Then
          '            mNumEntries = mNumEntries + 1
          '          End If
790       rows.Add AlarmHistory


800       rs.MoveNext
810     Loop

  Set AlarmHistory = New cAlarmHistory
  AlarmHistory.Extra = "Average"
  AlarmHistory.ACKElapsed = mSumACKElapsed
  AlarmHistory.ElapsedTime = mSumElapsed
  AlarmHistory.FinalizeElapsed = mSumFinalized
  AlarmHistory.Count = mNumEntries
  AlarmHistory.Count2 = mNumFinalized
  If mNumEntries Then
    AlarmHistory.Average = mSumElapsed / mNumEntries
  End If
  If mNumFinalized > 0 Then
    AlarmHistory.Average2 = mSumFinalized / mNumFinalized
  End If


900     rows.Add AlarmHistory

910     'Debug.Print "Rows " & RowCount
920     rs2.Close
930     Set rs2 = Nothing
940     rs.Close
950     Set rs = Nothing

End Sub
Sub MonthlyReport()
        Dim SQL                As String
        Dim WhereClause        As String
        Dim rs                 As Recordset

        Dim SQL2               As String
        Dim Whereclause2       As String
        Dim rs2                As Recordset

        Dim rc                 As Long

        Dim AlarmHistory       As cAlarmHistory

        Dim i                  As Long

        Dim StartDate          As Date

        Dim CurrentDay         As Integer

        Dim t                  As Long



10      Title = "Current Month Report"

20      ReportDate = Now

30      StartDate = Fix(Now)
40      CurrentDay = Day(StartDate)
50      StartDate = DateAdd("d", -(CurrentDay - 1), StartDate)

60      WhereClause = " AND (EventDate >= " & DateDelim & StartDate & DateDelim & " )  "

70      SQL = "SELECT Alarms.ID, Alarms.Serial, Alarms.EventDate, Alarms.Alarm, Alarms.Tamper, Alarms.Battery, Alarms.Announce, Alarms.EventType, Alarms.UserName, Alarms.AlarmID, Residents.NameLast, Residents.NameFirst, Rooms.Room  ,Devices.Model " & _
              " FROM (Alarms left JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) left JOIN Rooms ON Alarms.RoomID = Rooms.RoomID LEFT JOIN Devices ON Alarms.Serial = Devices.Serial  WHERE (AlarmID = 0) " & WhereClause & " ORDER BY ID"

80      t = Win32.timeGetTime()
90      Set rs = ConnExecute(SQL)

        t = Win32.timeGetTime() - t
        Debug.Print "First SQL (Monthly report( took " & Format(t / 1000, "0.0") & " seconds"
        t = Win32.timeGetTime()


        'Filter by selected EventTypes (new 3/5/2014 for trying to speed queries that intermittently slow to a crawl)
        Dim BasicEventTypes()  As String
        Dim filterstring       As String

        'alarmtypes = Join(Array(EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN), ",")

100     BasicEventTypes = Split(alarmtypes, ",")

110     For i = LBound(BasicEventTypes) To UBound(BasicEventTypes)
120       BasicEventTypes(i) = "eventtype = " & BasicEventTypes(i)
130     Next

140     filterstring = Join(BasicEventTypes, " or ")


        

150     rs.Filter = filterstring

        
        t = Win32.timeGetTime() - t
        Debug.Print "Rs.Filter = " & filterstring & " (Monthly report( took " & Format(t / 1000, "0.0") & " seconds"
        t = Win32.timeGetTime()
        

160     Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & " )  "

170     SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as eventdate,Alarms_1.info as info, Alarms_1.EventType as eventtype, Alarms_1.UserName as username" & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
               " ORDER BY  Alarms_1.AlarmID,Alarms.ID; "


        ' Debug.Assert 0
180     Set rs2 = ConnExecute(SQL2)



190     t = Win32.timeGetTime() - t
200     Debug.Print "SQL 2 (Monthly report) took " & Format(t / 1000, "0.0") & " seconds"
        t = Win32.timeGetTime()
210     i = 0
220


230     mAverageElapsed = 0
240     mNumEntries = 0
250     mSumElapsed = 0


        mAverageFinalized = 0
        mNumFinalized = 0
        mSumFinalized = 0
        mAverageFinalizedString = ""


260     Do Until rs.EOF

          Debug.Print "ID " & rs("ID")


          'mNumEntries = mNumEntries + 1

270       COUNTBYN = COUNTBYN + 1
280       If COUNTBYN > 100 Then
290         COUNTBYN = 0
300         'DoEvents
310       End If


          Dim CLSPTI           As Long
          'Dim device           As cESDevice
320       'CLSPTI = 0

          CLSPTI = Val(GetCLSPTI(rs("Model") & ""))

330       'Set device = Devices.device(Rs("Serial") & "")
340       'If Not device Is Nothing Then
350       '  CLSPTI = device.CLSPTI
360       'End If
370       Set AlarmHistory = New cAlarmHistory

          If rs("ID") > 10622 Then
            Debug.Print "rs(ID) " & rs("ID")
          End If

380       AlarmHistory.ID = rs("ID")  ' this record's ID
390       AlarmHistory.Serial = rs("Serial") & ""
400       AlarmHistory.EventDate = rs("EventDate")
410       AlarmHistory.alarm = rs("Alarm")
420       AlarmHistory.Tamper = rs("Tamper")
430       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
440       AlarmHistory.Announce = rs("Announce") & ""
450       AlarmHistory.EventType = rs("EventType")
460       AlarmHistory.User = rs("Username") & ""
470       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm

480       If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
490         AlarmHistory.Resident = ""
500       Else
510         If Len(rs("NameLast") & "") Then
520           If Len(rs("NameFirst") & "") Then
530             AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
540           Else
550             AlarmHistory.Resident = rs("NameLast") & ""
560           End If
570         ElseIf Len(rs("NameFirst") & "") Then
580           AlarmHistory.Resident = rs("NameFirst") & ""
590         End If
600       End If
610       AlarmHistory.Room = rs("room") & ""

620       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

630       Select Case rs("eventtype")

            Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN, EVT_ASSISTANCE

640           If CLSPTI <> &HEEEE& Then  ' 61166 dec

                'If AlarmHistory.ElapsedTime > 0 Or AlarmHistory.ACKElapsed > 0 Then
650             If gElapsedEqACK = 1 Then
660               If AlarmHistory.ACKElapsed > 0 Then
670                 mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
680                 mNumEntries = mNumEntries + 1
690               End If
700             ElseIf AlarmHistory.ElapsedTime > 0 Then
710               mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
720               mNumEntries = mNumEntries + 1
730             End If
                

                If AlarmHistory.FinalizeElapsed > 0 Then
                  mNumFinalized = mNumFinalized + 1
                  mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                End If

740           End If
750         Case Else

760       End Select
770       rows.Add AlarmHistory

        If (rows.Count Mod 100) = 0 Then
            DoEvents
        'Debug.Print "Month Add Row "; Now
        End If

780       rs.MoveNext
790     Loop


  ' adds an extra alarm history as averages

  Set AlarmHistory = New cAlarmHistory
  AlarmHistory.Extra = "Average"
  AlarmHistory.ACKElapsed = mSumACKElapsed
  AlarmHistory.ElapsedTime = mSumElapsed
  AlarmHistory.FinalizeElapsed = mSumFinalized
  AlarmHistory.Count = mNumEntries
  AlarmHistory.Count2 = mNumFinalized
  If mNumEntries Then
    AlarmHistory.Average = mSumElapsed / mNumEntries
  End If
  If mNumFinalized > 0 Then
    AlarmHistory.Average2 = mSumFinalized / mNumFinalized
  End If
  



'800     Set AlarmHistory = New cAlarmHistory
'810     AlarmHistory.Extra = "Average"
'820     AlarmHistory.ACKElapsed = mSumACKElapsed
'830     AlarmHistory.ElapsedTime = mSumElapsed
'840     AlarmHistory.Count = mNumEntries
'850     If mNumEntries > 0 Then
'
'860       AlarmHistory.Average = mSumElapsed / mNumEntries
'
'870     End If
'
'        If mNumFinalized > 0 Then
'          AlarmHistory.Average2 = mSumFinalized / mNumFinalized
'        End If

880     rows.Add AlarmHistory

        
890     rs.Close
900     Set rs = Nothing

        t = Win32.timeGetTime() - t
        Debug.Print "Final Results (Monthly report( took " & Format(t / 1000, "0.0") & " seconds"
        t = Win32.timeGetTime()

End Sub


Private Sub Class_Initialize()

  Set rows = New Collection
  alarmtypes = Join(Array(EVT_EMERGENCY, EVT_ALERT, EVT_TAMPER, EVT_BATTERY_FAIL, EVT_CHECKIN_FAIL, EVT_COMM_TIMEOUT, EVT_EXTERN, EVT_EXTERN_TROUBLE, EVT_ASSISTANCE), ",")

End Sub

Private Sub Class_Terminate()
  Set rows = Nothing
End Sub

Private Function GetReportDateString() As String
  GetReportDateString = Format(ReportDate, "yyyymmdd")
End Function

