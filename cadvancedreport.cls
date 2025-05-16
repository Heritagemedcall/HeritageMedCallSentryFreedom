VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdvancedReport"
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


Public lv                As ListView

Public Title             As String


Private ReportDate       As Date
Private ReportType       As String
Private alarmtypes       As String
Private mReportType      As Integer
Private mStartDate       As Date
Private mEndDate         As Date
Private mCriteria        As String


Const COL_SERIAL = 0
Const COL_ROOM = 0.75
Const COL_RESIDENT = 1.5
Const COL_EVENT = 3.5
Const COL_PHONE = 3.5
Const COL_ANNOUNCE = 4.5
Const COL_NOTES = 5
Const COL_ALARMTIME = 8.5
'Const COL_CANCELTIME = 8.5

Const COL_ELAPSEDTIME = 9.5
Const COL_OPERATOR = 9.75

Private mAverageElapsed  As Double
Private mNumEntries      As Long
Private mSumElapsed      As Double
Private mAverageElapsedString As String


Private mAverageFinalized  As Double
Private mNumFinalized      As Long
Private mSumFinalized      As Double
Private mAverageFinalizedString As String




Public Function FillAdvanced(ByVal ReportType As Integer, ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date)
  mReportType = ReportType
  mStartDate = StartDate
  mEndDate = EndDate
  mCriteria = Criteria
  Dim SQL                As String
  Set rows = New Collection

  ReportDate = StartDate


  Select Case ReportType
    Case RPT_EXCEPTION
      Title = "Exception Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindException Criteria, StartDate, EndDate
      PrintReport

    Case RPT_ROOM
      Title = "Room Event Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindRoom Criteria, StartDate, EndDate
      PrintReport
    Case RPT_RES
      Title = "Resident Event Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindRes Criteria, StartDate, EndDate
      PrintReport
    Case RPT_DEVICE
      Title = "Device Event Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindDevice Criteria, StartDate, EndDate
      PrintReport
    Case RPT_EVENT
      Title = "Event Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindEvent Criteria, StartDate, EndDate
      PrintReport
    Case RPT_ASSUR
      Title = "Check-in Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindAssur StartDate, EndDate
      PrintReport
    Case RPT_DEVHIST  ' add/remove devices
      Title = "Device History Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindDevHist Criteria, StartDate, EndDate
      PrintReport
    Case RPT_RESHIST  ' add/remove residents
      Title = "Resident History Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
      FindResHist Criteria, StartDate, EndDate
      PrintReport
    
    Case Else
      ' no data
      PrintReport
  End Select



End Function
Function FindDevHist(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long

        Dim SQL                As String
        Dim rs                 As Recordset
        Dim AlarmHistory       As cAlarmHistory
        Dim counter            As Long
        'Global Const EVT_VACATION = 29
        'Global Const EVT_VACATION_RETURN = 30
        'Global Const EVT_ADD_RES = 33
        'Global Const EVT_REMOVE_RES = 34
        'Global Const EVT_ADD_DEV = 35
        'Global Const EVT_REMOVE_DEV = 36

        Dim Events             As String
10      Events = Join(Array(EVT_ADD_DEV, EVT_REMOVE_DEV), ",")


20      EndDate = DateAdd("d", 1, EndDate)

30      On Error GoTo FindRoom_Error


40      SQL = "SELECT alarms.* , Residents.NameLast, Residents.NameFirst "
50      SQL = SQL & " FROM Alarms left join residents on alarms.residentid = Residents.ResidentID "
60      SQL = SQL & " WHERE "
70      SQL = SQL & " alarms.eventtype in (" & Events & ")"
80      SQL = SQL & " AND "
90      SQL = SQL & " (EventDate >= " & DateDelim & StartDate & DateDelim & ") And (EventDate < " & DateDelim & EndDate & DateDelim & ")"
100     SQL = SQL & " AND "
110     SQL = SQL & " (AlarmID = 0) "
120     SQL = SQL & " AND "
130     SQL = SQL & " Serial like " & q(Criteria)
140     SQL = SQL & " ORDER BY Serial, EventDate"


150      Set rs = ConnExecute(SQL)


        ' rs.Filter = "eventtype=" & EVT_ADD_DEV & " OR eventtype=" & EVT_REMOVE_DEV

160     Do Until rs.EOF
170       counter = counter + 1
180       If counter > 200 Then
190         DoEvents
200         counter = 0
210       End If
          ' get resident and room info
220       Set AlarmHistory = New cAlarmHistory
230       AlarmHistory.ID = rs("ID")      ' this record's ID
240       AlarmHistory.Serial = rs("Serial") & ""
250       AlarmHistory.EventDate = rs("EventDate")
260       AlarmHistory.alarm = rs("Alarm")
270       AlarmHistory.Tamper = rs("Tamper")
280       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
290       AlarmHistory.Announce = rs("Announce") & ""
300       AlarmHistory.EventType = rs("EventType")
310       AlarmHistory.User = rs("Username") & ""
320       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm



330       If Len(rs("NameLast") & "") Then
340         If Len(rs("NameFirst") & "") Then
350           AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
360         Else
370           AlarmHistory.Resident = rs("NameLast") & ""
380         End If
390       ElseIf Len(rs("NameFirst") & "") Then
400         AlarmHistory.Resident = rs("NameFirst") & ""
410       End If

420       rows.Add AlarmHistory
430       rs.MoveNext
440     Loop
450     rs.Close
460     'rs2.Close

FindRoom_Resume:
470     Set rs = Nothing
480     'Set rs2 = Nothing
490     On Error GoTo 0
500     Exit Function

FindRoom_Error:

510     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindDevHist." & Erl
520     Resume FindRoom_Resume


End Function
Function FindResHist(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As Recordset
        Dim AlarmHistory       As cAlarmHistory
        Dim counter            As Long
        Dim rs2                As ADODB.Recordset
        Dim SQL2               As String
        Dim rc                 As Long
        Dim Whereclause2       As String


        'Global Const EVT_VACATION = 29
        'Global Const EVT_VACATION_RETURN = 30
        'Global Const EVT_ADD_RES = 33
        'Global Const EVT_REMOVE_RES = 34
        'Global Const EVT_ADD_DEV = 35
        'Global Const EVT_REMOVE_DEV = 36

        Dim Events             As String

10      Events = Join(Array(EVT_VACATION, EVT_VACATION_RETURN, EVT_ADD_RES, EVT_REMOVE_RES), ",")


20      EndDate = DateAdd("d", 1, EndDate)

30      On Error GoTo FindRoom_Error

        '  SQl = " SELECT Residents.NameLast, Residents.NameFirst, residents.room, Alarms.*" & _
           '      " FROM Residents INNER JOIN Alarms ON Residents.residentID = Alarms.ResidentID " & _
           '      " WHERE alarms.eventtype in (" & Events & ")  AND ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
           '      " AND (Residents.NameLast like " & q(Criteria) & " OR Residents.Namefirst like " & q(Criteria) & ")" & _
           '      " ORDER BY  Alarms.EventDate , residents.NameLast, residents.Namefirst "

40      SQL = " SELECT Residents.NameLast, Residents.NameFirst, residents.room, Alarms.*" & _
            " FROM Residents INNER JOIN Alarms ON Residents.residentID = Alarms.ResidentID " & _
            " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
            " AND (Residents.NameLast like " & q(Criteria) & " OR Residents.Namefirst like " & q(Criteria) & ")" & _
            " ORDER BY Alarms.ID , residents.NameLast, residents.Namefirst "

50      Set rs = ConnExecute(SQL)
60      Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ")"

70      SQL2 = " SELECT Alarms.ID, Alarms.Serial,  Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as eventdate, Alarms_1.EventType as EventType, Alarms_1.info as info , Alarms_1.UserName as UserName " & _
             " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
             " ORDER BY Alarms_1.AlarmID, Alarms.ID; "


80      Set rs2 = ConnExecute(SQL2)


90      rs.Filter = "eventtype = " & EVT_VACATION & " OR eventtype = " & EVT_VACATION_RETURN & " OR eventtype = " & EVT_ADD_RES & " OR eventtype = " & EVT_REMOVE_RES

100     Do Until rs.EOF
110       counter = counter + 1
120       If counter > 200 Then
130         DoEvents
140         counter = 0
150       End If
160       Set AlarmHistory = New cAlarmHistory



170       AlarmHistory.ID = rs("ID")  ' this record's ID
180       AlarmHistory.Serial = rs("Serial") & ""
190       AlarmHistory.EventDate = rs("EventDate")
200       AlarmHistory.alarm = rs("Alarm")
210       AlarmHistory.Tamper = rs("Tamper")
220       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
230       AlarmHistory.Announce = rs("Announce") & ""
240       AlarmHistory.EventType = rs("EventType")
250       AlarmHistory.User = rs("Username") & ""
260       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm



270       If Len(rs("NameLast") & "") Then
280         If Len(rs("NameFirst") & "") Then
290           AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
300         Else
310           AlarmHistory.Resident = rs("NameLast") & ""
320         End If
330       ElseIf Len(rs("NameFirst") & "") Then
340         AlarmHistory.Resident = rs("NameFirst") & ""
350       End If
360       AlarmHistory.Room = rs("room") & ""

370       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)



          'get resident and room info
          'Set AlarmHistory = GetAlarmHistory(rs)
380       rows.Add AlarmHistory
390       rs.MoveNext
400     Loop
410     rs.Close
420     rs2.Close

FindRoom_Resume:

430     On Error GoTo 0
440     Set rs = Nothing
450     Set rs2 = Nothing

460     Exit Function

FindRoom_Error:

470     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindResHist." & Erl
480     Resume FindRoom_Resume




End Function


Function FindRoom(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim AlarmHistory       As cAlarmHistory
        Dim counter            As Long
        Dim rs2                As ADODB.Recordset
        Dim SQL2               As String
        Dim rc                 As Long
        Dim Whereclause2       As String

        ' get resident-room alarms
        ' AND
        ' get room-only alarms

        

10      EndDate = DateAdd("d", 1, EndDate)

20      On Error GoTo FindRoom_Error

30      SQL = " SELECT Rooms.Room, Alarms.*, Residents.NameLast, Residents.NameFirst " & _
              " FROM (Rooms INNER JOIN Alarms ON Rooms.RoomID = Alarms.RoomID) LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID " & _
              " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
              " AND Rooms.Room like " & q(Criteria) & _
              " ORDER BY alarms.id"

40      Set rs = ConnExecute(SQL)

50      Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ")"

60      SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as eventdate, Alarms_1.EventType as eventtype, Alarms_1.info as info , Alarms_1.UserName as username " & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
               " ORDER BY Alarms.ID; "


70      Set rs2 = ConnExecute(SQL2)
80      counter = 0

90      counter = counter

100     mNumEntries = 0
110     mAverageElapsed = 0
120     mSumElapsed = 0

130   mAverageFinalized = 0
140   mNumFinalized = 0
150   mSumFinalized = 0
160   mAverageFinalizedString = ""

        



170     Do Until rs.EOF
          'Debug.Print "Rooms.room " & rsEvents("room")
180       counter = counter + 1
190       If counter > 500 Then
200         DoEvents
210         counter = 0
220       End If

230       Set AlarmHistory = New cAlarmHistory
240       AlarmHistory.ID = rs("ID")  ' this record's ID
250       AlarmHistory.Serial = rs("Serial") & ""
260       AlarmHistory.EventDate = rs("EventDate")
270       AlarmHistory.alarm = rs("Alarm")
280       AlarmHistory.Tamper = rs("Tamper")
290       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
300       AlarmHistory.Announce = rs("Announce") & ""
310       AlarmHistory.EventType = rs("EventType")
320       AlarmHistory.User = rs("Username") & ""
330       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm
          
          Dim Device As cESDevice
340       Set Device = Devices.Device(AlarmHistory.Serial)
          
350       If Not Device Is Nothing Then
360         AlarmHistory.CLSPTI = Device.CLSPTI
370       End If


380       If Len(rs("NameLast") & "") Then
390         If Len(rs("NameFirst") & "") Then
400           AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
410         Else
420           AlarmHistory.Resident = rs("NameLast") & ""
430         End If
440       ElseIf Len(rs("NameFirst") & "") Then
450         AlarmHistory.Resident = rs("NameFirst") & ""
460       End If
470       AlarmHistory.Room = rs("room") & ""



480       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

490       Select Case AlarmHistory.EventType
            Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

500           If AlarmHistory.CLSPTI <> &HEEEE& Then  ' 61166 dec
              
510             mNumEntries = mNumEntries + 1
520             mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
530           End If
540         Case Else

550       End Select

560       rows.Add AlarmHistory
570       rs.MoveNext
580     Loop
590     rs.Close
600     rs2.Close

610     Set AlarmHistory = New cAlarmHistory
620     AlarmHistory.Extra = "Average"
630     AlarmHistory.ElapsedTime = mSumElapsed



FindRoom_Resume:
640     Set rs = Nothing
650     Set rs2 = Nothing

660     On Error GoTo 0
670     Exit Function

FindRoom_Error:

680     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindRoom." & Erl
690     Resume FindRoom_Resume


End Function

Function FindRes(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As Recordset
        Dim AlarmHistory       As cAlarmHistory
        Dim counter            As Long
        Dim rs2                As ADODB.Recordset
        Dim SQL2               As String
        Dim rc                 As Long
        Dim Whereclause2       As String

        ' get resident-room alarms
        ' AND
        ' get room-only alarms
10      If InStr(Criteria, "%") = 0 Then
20        Criteria = Criteria & "%"
30      End If

40      EndDate = DateAdd("d", 1, EndDate)

50      On Error GoTo FindRoom_Error

60      SQL = " SELECT Rooms.Room, Alarms.*, Residents.NameLast, Residents.NameFirst " & _
            " FROM (Residents INNER JOIN Alarms ON Residents.ResidentID = Alarms.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
            " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
            " AND (Residents.NameLast like " & q(Criteria) & " OR Residents.Namefirst like " & q(Criteria) & ")" & _
            " ORDER BY alarms.id, residents.NameLast, residents.Namefirst"

70      Set rs = ConnExecute(SQL)

80      Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ")"

90      SQL2 = " SELECT Alarms.ID, Alarms.Serial,  Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as EventDate, Alarms_1.EventType as EventType, Alarms_1.info as info , Alarms_1.UserName as UserName " & _
             " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
             " ORDER BY Alarms_1.AlarmID, Alarms.ID; "


100     Set rs2 = ConnExecute(SQL2)


110     Do Until rs.EOF
120       counter = counter + 1
130       If counter > 200 Then
140         DoEvents
150         counter = 0
160       End If
170       Set AlarmHistory = New cAlarmHistory
180       AlarmHistory.ID = rs("ID")  ' this record's ID
190       AlarmHistory.Serial = rs("Serial") & ""
200       AlarmHistory.EventDate = rs("EventDate")
210       AlarmHistory.alarm = rs("Alarm")
220       AlarmHistory.Tamper = rs("Tamper")
230       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
240       AlarmHistory.Announce = rs("Announce") & ""
250       AlarmHistory.EventType = rs("EventType")
260       AlarmHistory.User = rs("Username") & ""
270       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm

280       If Len(rs("NameLast") & "") Then
290         If Len(rs("NameFirst") & "") Then
300           AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
310         Else
320           AlarmHistory.Resident = rs("NameLast") & ""
330         End If
340       ElseIf Len(rs("NameFirst") & "") Then
350         AlarmHistory.Resident = rs("NameFirst") & ""
360       End If
370       AlarmHistory.Room = rs("room") & ""

380       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

          'Set AlarmHistory = GetAlarmHistory(rs)
390       rows.Add AlarmHistory
400       rs.MoveNext
410     Loop
420     rs.Close


FindRoom_Resume:
430     On Error GoTo 0
440     Exit Function

FindRoom_Error:

450     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindRoom." & Erl
460     Resume FindRoom_Resume


End Function

Function FindDevice(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As Recordset
        Dim AlarmHistory       As cAlarmHistory
        Dim rs2                As ADODB.Recordset
        Dim SQL2               As String
        Dim rc                 As Long
        Dim Whereclause2       As String
        Dim counter            As Long


10      On Error GoTo FindDevice_Error

20      EndDate = DateAdd("d", 1, EndDate)


30      SQL = " SELECT Rooms.Room, Alarms.*, Residents.NameLast, Residents.NameFirst " & _
              " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
              " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
              " AND (alarms.serial like " & q(Criteria) & ")" & _
              " ORDER BY alarms.id "

40      Set rs = ConnExecute(SQL)


50      Whereclause2 = " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ")) AND (alarms.serial like " & q(Criteria) & ")"

60      SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as EventDate, Alarms_1.EventType as EventType, Alarms_1.info as info , Alarms_1.UserName as UserName " & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
               " ORDER BY Alarms.ID; "


70      Set rs2 = ConnExecute(SQL2)

80      mNumEntries = 0
90      mAverageElapsed = 0
100     mSumElapsed = 0

mAverageFinalized = 0
mNumFinalized = 0
mSumFinalized = 0
mAverageFinalizedString = ""



110     Do Until rs.EOF
          'Debug.Print "residents.NameLast " & rs("namelast")
120       counter = counter + 1
130       If counter > 200 Then
140         DoEvents
150         counter = 0
160       End If
170       mNumEntries = mNumEntries + 1
180       Set AlarmHistory = New cAlarmHistory
190       AlarmHistory.ID = rs("ID")  ' this record's ID
200       AlarmHistory.Serial = rs("Serial") & ""
210       AlarmHistory.EventDate = rs("EventDate")
220       AlarmHistory.alarm = rs("Alarm")
230       AlarmHistory.Tamper = rs("Tamper")
240       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
250       AlarmHistory.Announce = rs("Announce") & ""
260       AlarmHistory.EventType = rs("EventType")
270       AlarmHistory.User = rs("Username") & ""
280       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm


290       If Len(rs("NameLast") & "") Then
300         If Len(rs("NameFirst") & "") Then
310           AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
320         Else
330           AlarmHistory.Resident = rs("NameLast") & ""
340         End If
350       ElseIf Len(rs("NameFirst") & "") Then
360         AlarmHistory.Resident = rs("NameFirst") & ""
370       End If
380       AlarmHistory.Room = rs("room") & ""

390       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

    Select Case AlarmHistory.EventType
      Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

        If AlarmHistory.CLSPTI <> &HEEEE& Then  ' 61166 dec
          mNumEntries = mNumEntries + 1
          mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
        End If
      Case Else

    End Select

'400       mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
410       rows.Add AlarmHistory
420       rs.MoveNext
430     Loop
440     rs.Close
450     rs2.Close

460    ' Set AlarmHistory = New cAlarmHistory
'        AlarmHistory.EventName = "Average"
'470     AlarmHistory.Extra = "Average"
'480     AlarmHistory.ElapsedTime = mSumElapsed

490     'rows.Add AlarmHistory


FindDevice_Resume:
500     Set rs = Nothing
510     Set rs2 = Nothing
520     On Error GoTo 0
530     Exit Function

FindDevice_Error:

540     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindDevice." & Erl
550     Resume FindDevice_Resume

End Function

Function FindException(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
  Dim SQL                As String
  Dim rs                 As Recordset
  Dim AlarmHistory       As cAlarmHistory
  Dim WhereClause        As String
  Dim ExceptionLimit     As Double
  Dim Devices()          As String
  Dim params()           As String
  Dim counter            As Long


  On Error GoTo FindException_Error

  params = Split(Criteria, ";")

  ExceptionLimit = Val(params(0)) * 60
  Devices() = Split(params(1), ",")

  EndDate = DateAdd("d", 1, EndDate)

  'may need to filter and map devices "custom" field to device types

  ' WhereClause = " AND (alarms.AlarmType in " & params(1) & ") "

  SQL = " SELECT Rooms.Room, Alarms.*, Residents.NameLast, Residents.NameFirst " & _
      " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
      " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
        WhereClause & _
      " ORDER BY Alarms.id "

  Set rs = ConnExecute(SQL)


  Dim rs2                As ADODB.Recordset
  Dim Whereclause2       As String
  Dim SQL2               As String
  Dim rc                 As Long


  SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as EventDate , Alarms_1.EventType as EventType, Alarms_1.info as info , Alarms_1.UserName as UserName " & _
       " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
         Whereclause2 & _
       " ORDER BY Alarms_1.AlarmID, Alarms.ID; "

  Whereclause2 = Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ")"
  Set rs2 = ConnExecute(SQL2)


  Do Until rs.EOF
    'Debug.Print "residents.NameLast " & rs("namelast")
    counter = counter + 1
    If counter Mod nDoEvents = 0 Then  ' just yield every n hits
      DoEvents
    End If

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


    rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

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


    If AlarmHistory.ElapsedTime > ExceptionLimit Then
      rows.Add AlarmHistory
    End If
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing



FindException_Resume:

  On Error GoTo 0
  Exit Function

FindException_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindException." & Erl
  Resume FindException_Resume

End Function

Function FindEvent(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
  Dim SQL                As String
  Dim rs                 As Recordset
  Dim AlarmHistory       As cAlarmHistory
  Dim WhereClause        As String

  Dim rs2                As ADODB.Recordset
  Dim SQL2               As String
  Dim rc                 As Long
  Dim Whereclause2       As String


  On Error GoTo FindEvent_Error

  EndDate = DateAdd("d", 1, EndDate)


  If Val(Criteria) = 0 Then
    WhereClause = ""
  ElseIf Criteria = EVT_CHECKIN_FAIL Then
    WhereClause = " AND ((alarms.eventtype = " & Val(Criteria) & ") or (alarms.eventtype = " & EVT_LINELOSS & "))"
  Else
    WhereClause = " AND (alarms.eventtype = " & Val(Criteria) & ")"
  End If

  '  ID, serial, eventdate, alarm, tamper, battery,residentid, roomid, announce, eventtype, username , alarmid

  SQL = " SELECT Rooms.Room, Alarms.id, alarms.serial, alarms.eventdate, alarms.alarm, alarms.tamper, alarms.battery, alarms.residentid, alarms.roomid," & _
        " alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst " & _
        " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
        " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
        WhereClause & _
        " ORDER BY Alarms.id "

  Set rs = ConnExecute(SQL)



  Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ")"

  SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as AlarmID, Alarms_1.EventDate as EventDate , Alarms_1.EventType as EventType, Alarms_1.info as info , Alarms_1.UserName as UserName " & _
         " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
         Whereclause2 & _
         " ORDER BY Alarms_1.AlarmID, Alarms.ID; "


  Set rs2 = ConnExecute(SQL2)


  Dim COUNTBYN           As Long

  mNumEntries = 0
  mAverageElapsed = 0
  mSumElapsed = 0

mAverageFinalized = 0
mNumFinalized = 0
mSumFinalized = 0
mAverageFinalizedString = ""


  'Debug.Print " "
  Do Until rs.EOF
    'Debug.Print " rs(id)"; rs("id")

    'Debug.Print "residents.NameLast " & rs("namelast")

    COUNTBYN = COUNTBYN + 1
    If COUNTBYN > 200 Then
      COUNTBYN = 0
      DoEvents
    End If

    'mNumEntries = mNumEntries + 1
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


    'need to remove this , done in p[rint report !!!!!!! 2016-02-09

    Select Case AlarmHistory.EventType
      Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

        Dim Device       As cESDevice

        Set Device = Devices.Device(AlarmHistory.Serial)
        If Not Device Is Nothing Then
          AlarmHistory.CLSPTI = Device.CLSPTI
        End If

        'Debug.Print "cAdvancedReport.FindEvent ", AlarmHistory.CLSPTI, AlarmHistory.EventDate

        If AlarmHistory.CLSPTI <> &HEEEE& Then  ' 61166 dec
          If AlarmHistory.ElapsedTime > 0 Then

            mNumEntries = mNumEntries + 1
            mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
            'Debug.Print "cAdvancedReport.FindEvent Added ", AlarmHistory.CLSPTI, AlarmHistory.EventDate, mNumEntries
          Else
            'Debug.Print "cAdvancedReport.FindEvent No Time ", AlarmHistory.CLSPTI, AlarmHistory.EventDate
          End If
        Else
          'Debug.Assert 0
          'Debug.Print "cAdvancedReport.FindEvent Bad CLSPTI ", AlarmHistory.CLSPTI, AlarmHistory.EventDate

        End If
      Case Else

    End Select



    rows.Add AlarmHistory
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing




  rs2.Close
  Set rs2 = Nothing


FindEvent_Resume:

  On Error GoTo 0
  Exit Function

FindEvent_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindEvent." & Erl
  Resume FindEvent_Resume

End Function



Function FindAssur(ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim WhereClause        As String
        Dim AlarmHistory       As cAlarmHistory
        Dim SQL2               As String
        Dim Whereclause2       As String


        Dim counter            As Long

        Dim rc                 As Long

        Dim rs2                As ADODB.Recordset

10      On Error GoTo FindAssur_Error

20      EndDate = DateAdd("d", 1, EndDate)



        WhereClause = " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
            " AND (alarms.Eventtype = " & EVT_ASSUR_FAIL & ")"


'30      SQl = " SELECT ID, serial, eventdate, alarm, tamper, battery,residentid, roomid, announce, eventtype, username , alarmid   " & _
'            " FROM Alarms " & _
'            " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
'            " AND (alarms.Eventtype = " & EVT_ASSUR_FAIL & ")" & _
'            " ORDER BY Alarms.id"

30      SQL = " SELECT Rooms.Room, Alarms.id, alarms.serial, alarms.eventdate, alarms.alarm, alarms.tamper, alarms.battery, alarms.residentid, alarms.roomid," & _
            " alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst " & _
            " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
             WhereClause & _
            " ORDER BY Alarms.id "



40      Set rs = ConnExecute(SQL)


50      Whereclause2 = " WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate <  " & DateDelim & EndDate & DateDelim & ")"

60      SQL2 = " SELECT Alarms.ID, Alarms.Serial,  Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as eventdate, Alarms_1.EventType as EventType, Alarms_1.info as info , Alarms_1.UserName as UserName " & _
             " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
               Whereclause2 & _
             " ORDER BY Alarms_1.AlarmID, Alarms.ID; "

70      Set rs2 = ConnExecute(SQL2)

80      Do Until rs.EOF
90        counter = counter + 1
100       If counter > 200 Then
110         counter = 0
120         DoEvents
130       End If


140       Set AlarmHistory = New cAlarmHistory



150       AlarmHistory.ID = rs("ID")  ' this record's ID
160       AlarmHistory.Serial = rs("Serial") & ""
170       AlarmHistory.EventDate = rs("EventDate")
180       AlarmHistory.alarm = rs("Alarm")
190       AlarmHistory.Tamper = rs("Tamper")
200       AlarmHistory.Battery = rs("Battery")
          'ResidentID = rs("residentID")
          'RoomID = rs("RoomID")
210       AlarmHistory.Announce = rs("Announce") & ""
220       AlarmHistory.EventType = rs("EventType")
230       AlarmHistory.User = rs("Username") & ""
240       AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm





250       If Len(rs("NameLast") & "") Then
260         If Len(rs("NameFirst") & "") Then
270           AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
280         Else
290           AlarmHistory.Resident = rs("NameLast") & ""
300         End If
310       ElseIf Len(rs("NameFirst") & "") Then
320         AlarmHistory.Resident = rs("NameFirst") & ""
330       End If
340       AlarmHistory.Room = rs("room") & ""
350       rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)
360       rows.Add AlarmHistory
370       rs.MoveNext
380     Loop
390     rs.Close
400     rs2.Close

FindAssur_Resume:


410     On Error GoTo 0
420     Set rs = Nothing
430     Set rs2 = Nothing

440     Exit Function

FindAssur_Error:

450     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.FindAssur." & Erl
460     Resume FindAssur_Resume

End Function


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

'2019-02-13 removed this sub

Sub Fill(ByVal ReportID As String)

End Sub
'  '  Select Case reportID
'  '    Case "currentshift"
'  '      ReportType = "S"
'  '      ShiftReport
'  '      PrintReport
'  '    Case "today"
'  '      ReportType = "D"
'  '      DailyReport
'  '      PrintReport
'  '    Case "thisweek"
'  '      ReportType = "W"
'  '      WeeklyReport
'  '      PrintReport
'  '    Case "thismonth"
'  '      ReportType = "M"
'  '      MonthlyReport
'  '      PrintReport
'  '    Case Else
'  '      ReportType = ""
'  '      DailyReport
'  '      PrintReport
'  '  End Select
'End Sub

Sub PrintReport()
        Dim AlarmHistory       As cAlarmHistory
        Dim filename           As String
        Dim COUNTBYN           As Long


10      On Error GoTo PrintReport_Error

20      Select Case Dest
          Case DEST_SCREEN  ' if To Screen ' Just add to listview
30          lv.ListItems.Clear
40          lv.Refresh
50        Case DEST_FILE  ' if To File ' one continuous file
60          hfile = FreeFile
70          filename = Replace(Title, "/", "-")
80          Open Configuration.ReportPath & CleanFileName(filename) & ".txt" For Output As hfile

90        Case DEST_PRINTER  ' if To Printer
100         Printer.Orientation = vbPRORLandscape
110         Printer.Print " ";
120         Printer.ScaleMode = vbInches
130         Printer.FontSize = 10
140         Printer.FontBold = False

150     End Select

160     PrintBanner
170     PrintRowHeader
180     If Dest = DEST_SCREEN Then
190       Win32.LockWindowUpdate lv.hwnd
200     End If

        

210     mNumEntries = 0
220     mAverageElapsed = 0
230     mSumElapsed = 0
        
        mAverageFinalized = 0
        mNumFinalized = 0
        mSumFinalized = 0
        mAverageFinalizedString = ""

        
        
240     For Each AlarmHistory In rows
250
              
260       COUNTBYN = COUNTBYN + 1
270       If COUNTBYN > 200 Then
280         COUNTBYN = 0
290         DoEvents
300       End If

310       Select Case Dest
            Case DEST_SCREEN  ' if To Screen ' Just add to listview
320         Case DEST_PRINTER  ' if To Printer
              ' check pagelength
330           If (Printer.ScaleHeight - Printer.CurrentY) < 0.5 Then  '  half inch bottom margin
340             PrintFooter
350             FormFeed Printer
360             PrintRowHeader
370           End If
380         Case DEST_FILE  ' if To File ' one continuous file
390       End Select

400       PrintRow AlarmHistory
410     Next

420     Select Case Dest
          Case DEST_SCREEN  ' if To Screen ' Just add to listview
            PrintAverage
430         Win32.LockWindowUpdate 0


440       Case DEST_PRINTER  ' if To Printer
450         PrintAverage
460         PrintFooter
470         Printer.EndDoc
480         Printer.Orientation = vbPRORPortrait
490       Case DEST_FILE  ' if To File ' one continuous file
500         PrintAverage
510         If hfile <> 0 Then
520           Close hfile
530           hfile = 0
540         End If
550     End Select

PrintReport_Resume:
560     On Error GoTo 0
570     Exit Sub

PrintReport_Error:

580     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.PrintReport." & Erl
590     Resume PrintReport_Resume


End Sub
Sub PrintAverage()
10      On Error GoTo PrintAverage_Error

  Dim li                 As ListItem

20      mAverageElapsed = 0
30      mAverageElapsedString = "N/A"
40      If mNumEntries > 0 Then
    'mAverageElapsed = Format$(mSumElapsed / mNumEntries, "hh:nn:ss")
50        mAverageElapsed = mSumElapsed / mNumEntries
60        mAverageElapsedString = SecondsToTimeString(mAverageElapsed)
70      End If

80      If mNumFinalized > 0 Then
90        mAverageFinalized = mSumFinalized / mNumFinalized
100       mAverageFinalizedString = SecondsToTimeString(mAverageFinalized)
110     End If



120     Select Case Dest
    Case DEST_SCREEN           ' if To Screen:
130   mAverageElapsedString = "N/A"
140   If (mAverageElapsed > 0) Then
        'mAverageElapsedString = Format$(DateAdd("s", mAverageElapsed, 0), "hh:nn:ss")
150     mAverageElapsedString = SecondsToTimeString(mAverageElapsed)
        
160   End If
      
      
170   mAverageFinalizedString = "N/A"
      
180   If (mAverageFinalized > 0) Then
190     mAverageFinalizedString = SecondsToTimeString(mAverageFinalized)
200   End If
      
210   Select Case mReportType
        
        Case RPT_EXCEPTION, RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT
220       Set li = lv.ListItems.Add(, , "")
230       li.SubItems(1) = "Average"
240       li.SubItems(6) = mAverageElapsedString
250   End Select
260       Case DEST_PRINTER          ' if To Printer:
270   Select Case mReportType
        Case RPT_EXCEPTION
280       PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Average"  ' need a tostring (HH:NN:SS)"
290        PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, mAverageElapsedString  ' need a tostring (HH:NN:SS)
300     Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT
310       PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Average"  ' need a tostring (HH:NN:SS)"
320       PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, mAverageElapsedString  ' need a tostring (HH:NN:SS)
330     Case RPT_ASSUR
340     Case RPT_RESHIST
          ' PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, mAverageElapsedString  ' need a tostring (HH:NN:SS)
350     Case RPT_DEVHIST
          ' PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, mAverageElapsedString  ' need a tostring (HH:NN:SS)
360     Case Else

370   End Select
380       Case DEST_FILE             ' if To File
390   Select Case mReportType
        Case RPT_EXCEPTION
400       Print #hfile, Join(Array("Average", "", "", "", "", "", _
                                   "", "", _
                                   "", mAverageElapsedString, ""), vbTab)

410     Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT
420       Print #hfile, Join(Array("Average", "", "", "", "", "", _
                                   "", "", _
                                   "", mAverageElapsedString, ""), vbTab)

430     Case RPT_ASSUR
440     Case RPT_RESHIST
          '300             Print #hfile, Join(Array("Average", "", "", "", "", "",                                          "", "",                                          "", mAverageElapsedString, ""), vbTab)

450     Case RPT_DEVHIST

          '300             Print #hfile, Join(Array("Average", "", "", "", "", "",                                          "", "",                                          "", mAverageElapsedString, ""), vbTab)
460     Case Else
470   End Select


480     End Select


PrintAverage_Resume:

490     On Error GoTo 0
500     Exit Sub

PrintAverage_Error:

510     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdvancedReport.PrintAverage." & Erl
520     Resume PrintAverage_Resume

End Sub


Sub PrintRow(AlarmHistory As cAlarmHistory)
  ' ID Room Resident Eventtype Alarmtime Canceltime ElapsedTime Username
  ' tab delimited
  Dim li                 As ListItem

  Dim Device             As cESDevice

  Dim CLSPTI             As Long

  Set Device = Devices.Device(AlarmHistory.Serial)
  If Not Device Is Nothing Then
    CLSPTI = Device.CLSPTI
  End If

  Select Case Dest
    Case DEST_SCREEN           ' if To Screen:
      If Not lv Is Nothing Then
        Select Case mReportType
          Case RPT_EXCEPTION

            Set li = lv.ListItems.Add(, , AlarmHistory.Serial)

            li.SubItems(1) = AlarmHistory.Room

            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              li.SubItems(2) = ""
            Else
              li.SubItems(2) = AlarmHistory.Resident
            End If

            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT  ' change 2/15/2012
                li.SubItems(3) = AlarmHistory.Announce
              Case Else
                li.SubItems(3) = AlarmHistory.EventName
            End Select

            li.SubItems(4) = AlarmHistory.AlarmTimeToString

            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

                If gElapsedEqACK = USE_ELAPSED_ACK Then
                  If AlarmHistory.ACKTime > 0 Then
                    li.SubItems(5) = AlarmHistory.ACKTimeToString
                    li.SubItems(6) = AlarmHistory.ACKElapsedTimeToString

                    If CLSPTI <> &HEEEE& Then
                      If AlarmHistory.ACKElapsed > 0 Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                      End If



                    End If


                  Else         ' try and use end of alarm time
                    li.SubItems(5) = AlarmHistory.CancelTimeToString
                    li.SubItems(6) = AlarmHistory.ElapsedTimeToString

                  End If

                  If AlarmHistory.FinalizeElapsed > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumFinalized = mNumFinalized + 1
                      mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                    End If
                  End If


                Else
                  li.SubItems(5) = AlarmHistory.CancelTimeToString
                  li.SubItems(6) = AlarmHistory.ElapsedTimeToString

                  If AlarmHistory.ElapsedTime > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumEntries = mNumEntries + 1
                      mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                    End If
                  End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If


                  End If

                Case Else
                  li.SubItems(5) = AlarmHistory.CancelTimeToString
                  li.SubItems(6) = AlarmHistory.ElapsedTimeToString


              End Select       ' Case AlarmHistory.Eventtype


              li.SubItems(7) = AlarmHistory.User


            Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT

              Set li = lv.ListItems.Add(, , AlarmHistory.Serial)

              li.SubItems(1) = AlarmHistory.Room

              If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
                li.SubItems(2) = ""
              Else

                li.SubItems(2) = AlarmHistory.Resident
              End If

              Select Case AlarmHistory.EventType
                Case EVT_EMERGENCY, EVT_ALERT  ' change 2/15/2012
                  li.SubItems(3) = AlarmHistory.Announce
                Case Else
                  If LCase(AlarmHistory.EventName) = "none" Then
                    Debug.Assert 0
                  End If
                  li.SubItems(3) = AlarmHistory.EventName
              End Select


              li.SubItems(4) = AlarmHistory.AlarmTimeToString


              Select Case AlarmHistory.EventType
                Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN


                  If gElapsedEqACK = USE_ELAPSED_ACK Then
                    If AlarmHistory.ACKTime > 0 Then
                      li.SubItems(5) = AlarmHistory.ACKTimeToString
                      li.SubItems(6) = AlarmHistory.ACKElapsedTimeToString

                      If AlarmHistory.ACKElapsed > 0 Then
                        If CLSPTI <> &HEEEE& Then
                          mNumEntries = mNumEntries + 1
                          mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                        End If
                      End If
                      If AlarmHistory.FinalizeElapsed > 0 Then
                        If CLSPTI <> &HEEEE& Then
                          mNumFinalized = mNumFinalized + 1
                          mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                        End If
                      End If


                    Else       ' try and use end of alarm time
                      li.SubItems(5) = AlarmHistory.CancelTimeToString
                      li.SubItems(6) = AlarmHistory.ElapsedTimeToString
                    End If
                  Else
                    li.SubItems(5) = AlarmHistory.CancelTimeToString
                    li.SubItems(6) = AlarmHistory.ElapsedTimeToString

                    If AlarmHistory.ElapsedTime > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  End If

                Case Else
              End Select       'Case AlarmHistory.Eventtype


              li.SubItems(7) = AlarmHistory.User

            Case RPT_ASSUR

              Set li = lv.ListItems.Add(, , AlarmHistory.Serial)
              li.SubItems(1) = AlarmHistory.Room
              If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
                li.SubItems(2) = ""
              Else
                li.SubItems(2) = AlarmHistory.Resident
              End If
              li.SubItems(3) = AlarmHistory.Phone

              li.SubItems(4) = AlarmHistory.AlarmTimeToString
              li.SubItems(5) = " "
              li.SubItems(6) = AlarmHistory.User

            Case RPT_RESHIST
              If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
                Set li = lv.ListItems.Add(, , AlarmHistory.Resident)
              Else

                Set li = lv.ListItems.Add(, , AlarmHistory.Resident)
              End If
              li.SubItems(1) = AlarmHistory.EventName
              li.SubItems(2) = AlarmHistory.AlarmTimeToString
              li.SubItems(3) = AlarmHistory.User

            Case RPT_DEVHIST

              Set li = lv.ListItems.Add(, , AlarmHistory.Serial)
              li.SubItems(1) = AlarmHistory.EventName
              li.SubItems(2) = AlarmHistory.AlarmTimeToString
              li.SubItems(3) = AlarmHistory.User

            Case Else

              Set li = lv.ListItems.Add(, , AlarmHistory.Serial)
              li.SubItems(1) = AlarmHistory.Room
              If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
                li.SubItems(2) = ""
              Else
                li.SubItems(2) = AlarmHistory.Resident
              End If

              li.SubItems(3) = AlarmHistory.EventName
              li.SubItems(4) = AlarmHistory.AlarmTimeToString


              Select Case AlarmHistory.EventType
                Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN


                  If gElapsedEqACK = USE_ELAPSED_ACK Then
                    If AlarmHistory.ACKTime > 0 Then
                      li.SubItems(5) = AlarmHistory.ACKTimeToString
                      li.SubItems(6) = AlarmHistory.ACKElapsedTimeToString
                      If AlarmHistory.ACKElapsed > 0 Then
                        If CLSPTI <> &HEEEE& Then
                          mNumEntries = mNumEntries + 1
                          mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                        End If
                      End If

                      If AlarmHistory.FinalizeElapsed > 0 Then
                        If CLSPTI <> &HEEEE& Then
                          mNumFinalized = mNumFinalized + 1
                          mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                        End If
                      End If


                    Else       ' try and use end of alarm time
                      li.SubItems(5) = AlarmHistory.CancelTimeToString
                      li.SubItems(6) = AlarmHistory.ElapsedTimeToString


                      ' mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                    End If
                  Else
                    li.SubItems(5) = AlarmHistory.CancelTimeToString
                    li.SubItems(6) = AlarmHistory.ElapsedTimeToString

                    If AlarmHistory.ElapsedTime > 0 Then


                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If



                  End If
                Case Else
                  li.SubItems(5) = AlarmHistory.CancelTimeToString
                  li.SubItems(6) = AlarmHistory.ElapsedTimeToString

              End Select       ' Case AlarmHistory.Eventtype
              li.SubItems(7) = AlarmHistory.User

          End Select
        End If

      Case DEST_PRINTER        ' if To Printer:
        Select Case mReportType
          Case RPT_EXCEPTION
            PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial & " "  ' need apostrophy for excel to like it
            PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, " "
            Else

              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident
            End If
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
            PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, AlarmHistory.Announce
            PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString

            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN

                If gElapsedEqACK = USE_ELAPSED_ACK Then
                  If AlarmHistory.ACKTime > 0 Then
                    PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ACKElapsedTimeToString

                    If AlarmHistory.ACKElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  End If
                Else
                  PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
                  If AlarmHistory.ElapsedTime > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumEntries = mNumEntries + 1
                      mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                    End If

                  End If
                  If AlarmHistory.FinalizeElapsed > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumFinalized = mNumFinalized + 1
                      mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                    End If
                  End If

                End If

            End Select         'Case AlarmHistory.Eventtype

            '          PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, AlarmHistory.CancelTimeToString
            '          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
            PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
            Printer.Print " "

          Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT
            PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial & " "  ' need apostrophy for excel to like it
            PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, " "
            Else

              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident
            End If
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
            PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, AlarmHistory.Announce
            PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString

            Select Case AlarmHistory.EventType

              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN  '

                If gElapsedEqACK = USE_ELAPSED_ACK Then
                  If AlarmHistory.ACKTime > 0 Then
                    PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ACKElapsedTimeToString
                    If CLSPTI <> &HEEEE& Then
                      mNumEntries = mNumEntries + 1
                      mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  Else         ' try and use end of alarm time
                    PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
                  End If
                Else

                  PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
                  If AlarmHistory.ElapsedTime > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumEntries = mNumEntries + 1
                      mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                    End If
                  End If
                  If AlarmHistory.FinalizeElapsed > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumFinalized = mNumFinalized + 1
                      mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                    End If
                  End If

                End If
              Case Else
                ' no average accum
            End Select         'Case AlarmHistory.Eventtype

            '          PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, AlarmHistory.CancelTimeToString
            '          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
            PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
            Printer.Print " "
          Case RPT_ASSUR
            PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial & " "
            PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, " "
            Else

              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident
            End If
            PrintLJ Printer, COL_PHONE, Printer.CurrentY, AlarmHistory.Phone

            'PrintLJ Printer, COL_NOTES, Printer.CurrentY, AlarmHistory.Notes
            PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString
            PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
            Printer.Print " "
            HR Printer

          Case RPT_RESHIST
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              PrintLJ Printer, COL_SERIAL, Printer.CurrentY, " "
            Else

              PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Resident
            End If
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
            PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString
            PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
            Printer.Print " "

          Case RPT_DEVHIST
            PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
            PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString
            PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
            Printer.Print " "


          Case Else
            PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial
            PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, " "
            Else

              PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident
            End If
            PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
            PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, AlarmHistory.Announce
            PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString


            Select Case AlarmHistory.EventType

              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN  ' change 2/15/2012

                If gElapsedEqACK = USE_ELAPSED_ACK Then
                  If AlarmHistory.ACKTime > 0 Then

                    PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ACKElapsedTimeToString

                    If AlarmHistory.ACKElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If


                  Else         ' try and use end of alarm time
                    PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString

                  End If
                Else
                  PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
                  If AlarmHistory.ElapsedTime > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumEntries = mNumEntries + 1
                      mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                    End If
                  End If
                  If AlarmHistory.FinalizeElapsed > 0 Then
                    If CLSPTI <> &HEEEE& Then
                      mNumFinalized = mNumFinalized + 1
                      mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                    End If
                  End If


                End If

              Case Else
            End Select         'Case AlarmHistory.Eventtype

            '          PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, AlarmHistory.CancelTimeToString
            '          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
            PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
            Printer.Print " "

        End Select

      Case DEST_FILE           ' if To File
        If hfile <> 0 Then
          ' need to modify this !
          Select Case mReportType
            Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT


              Select Case AlarmHistory.EventType
                Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                  If gElapsedEqACK = USE_ELAPSED_ACK Then

                    If AlarmHistory.ACKElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  Else
                    If AlarmHistory.ElapsedTime > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  End If
                Case Else

              End Select       'Case AlarmHistory.Eventtype

              'Print #hFile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.user), vbTab)
              Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, _
                                       AlarmHistory.ACKTimeToString, AlarmHistory.ACKElapsedTimeToString, _
                                       AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.User), vbTab)


            Case RPT_ASSUR
              Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.Phone, AlarmHistory.AlarmTimeToString, "", AlarmHistory.User), vbTab)
            Case RPT_RESHIST
              Print #hfile, Join(Array(AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.AlarmTimeToString, AlarmHistory.User), vbTab)
            Case RPT_DEVHIST
              Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.EventName, AlarmHistory.AlarmTimeToString, AlarmHistory.User), vbTab)
            Case Else
              '            Print #hFile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.user), vbTab)
              Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, _
                                       AlarmHistory.ACKTimeToString, AlarmHistory.ACKElapsedTimeToString, _
                                       AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.User), vbTab)


              Select Case AlarmHistory.EventType
                Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                  If gElapsedEqACK = USE_ELAPSED_ACK Then

                    If AlarmHistory.ACKElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  Else
                    If AlarmHistory.ElapsedTime > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumEntries = mNumEntries + 1
                        mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                      End If
                    End If
                    If AlarmHistory.FinalizeElapsed > 0 Then
                      If CLSPTI <> &HEEEE& Then
                        mNumFinalized = mNumFinalized + 1
                        mSumFinalized = mSumFinalized + AlarmHistory.FinalizeElapsed
                      End If
                    End If

                  End If
              End Select       'Case AlarmHistory.Eventtype

          End Select           'Case mReportType
        End If                 ' no file handle

    End Select                 'Case Dest (printer etc)

  End Sub
Sub PrintRowHeader()
  ' ID Room Resident Eventtype Alarmtime Canceltime ElapsedTime Username
  ' tab delimited

  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
      ConfigureListView


    Case DEST_PRINTER  ' if To Printer:
      Select Case mReportType
        Case RPT_EXCEPTION
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, "Announce"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Alarm Time"
          'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, "Cancel Time"
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "Elapsed Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print ""
          HR Printer

        Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT

          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, "Announce"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Alarm Time"
          'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, "Cancel Time"
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "Elapsed Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print ""
          HR Printer
        Case RPT_ASSUR
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_PHONE, Printer.CurrentY, "Phone"
          'PrintLJ Printer, COL_NOTES, Printer.CurrentY, "Notes"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Date/Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print " "
          HR Printer

        Case RPT_RESHIST

          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print " "
          HR Printer
        Case RPT_DEVHIST
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print " "
          HR Printer
        Case Else
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, "Announce"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Alarm Time"
          'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, "Cancel Time"
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "Elapsed Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print " "
          HR Printer

      End Select
    Case DEST_FILE  ' if To File
      If hfile <> 0 Then


        Select Case mReportType
          Case RPT_EXCEPTION
            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
                                     "ACK Time", "ACK Elapsed Time", _
                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)
          Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT
            'Print #hFile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", "Cancel Time", "Elapsed Time", "Operator"), vbTab)
            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
                                     "ACK Time", "ACK Elapsed Time", _
                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)
          Case RPT_ASSUR
            Print #hfile, Join(Array("Serial", "Room", "Resident", "Phone", "Date/Time", "Info", "Operator"), vbTab)
          Case RPT_RESHIST
            Print #hfile, Join(Array("Resident", "Event", "Time", "Operator"), vbTab)
          Case RPT_DEVHIST
            Print #hfile, Join(Array("Serial", "Event", "Time", "Operator"), vbTab)
          Case Else
            '            Print #hFile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", "Cancel Time", "Elapsed Time", "Operator"), vbTab)
            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
                                     "ACK Time", "ACK Elapsed Time", _
                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)

        End Select
      End If
  End Select


End Sub
Sub ConfigureListView()

  'Global Const RPT_ROOM = 1
  'Global Const RPT_RES = 2
  'Global Const RPT_DEVICE = 3
  'Global Const RPT_EVENT = 4
  'Global Const RPT_ASSUR = 5
  'Global Const RPT_DEVHIST = 6  ' add/remove devices
  'Global Const RPT_RESHIST = 7  ' add/remove residents
  'Global Const RPT_EXCEPTION = 8  ' exception (alarm ack/clear time)


  If Not lv Is Nothing Then
    Select Case mReportType
      Case RPT_EXCEPTION
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 0
        lv.ColumnHeaders.Add , , "Room", 1200
        lv.ColumnHeaders.Add , , "Resident", 1600
        lv.ColumnHeaders.Add , , "Event", 1300
        lv.ColumnHeaders.Add , , "Alarm", 1500
        lv.ColumnHeaders.Add , , "Cancel", 1300
        lv.ColumnHeaders.Add , , "Elapsed", 850
        lv.ColumnHeaders.Add , , "Operator", 900

      Case RPT_ROOM, RPT_RES, RPT_DEVICE, RPT_EVENT
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 0
        lv.ColumnHeaders.Add , , "Room", 1200
        lv.ColumnHeaders.Add , , "Resident", 1600
        lv.ColumnHeaders.Add , , "Event", 1300
        lv.ColumnHeaders.Add , , "Alarm", 1500
        lv.ColumnHeaders.Add , , "Cancel", 1300
        lv.ColumnHeaders.Add , , "Elapsed", 850
        lv.ColumnHeaders.Add , , "Operator", 900

      Case RPT_ASSUR


        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 0
        lv.ColumnHeaders.Add , , "Room", 1200
        lv.ColumnHeaders.Add , , "Resident", 1600
        lv.ColumnHeaders.Add , , "Phone", 800
        lv.ColumnHeaders.Add , , "Date/Time", 1500
        lv.ColumnHeaders.Add , , "Notes", 1300
        lv.ColumnHeaders.Add , , "Operator", 900
      Case RPT_RESHIST
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "Resident", 1800
        lv.ColumnHeaders.Add , , "Event", 1800
        lv.ColumnHeaders.Add , , "Date/Time", 1500
        lv.ColumnHeaders.Add , , "Operator", 900

      Case RPT_DEVHIST

        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 1000
        lv.ColumnHeaders.Add , , "Event", 1800
        lv.ColumnHeaders.Add , , "Date/Time", 1500
        lv.ColumnHeaders.Add , , "Operator", 900


      Case Else
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 0
        lv.ColumnHeaders.Add , , "Room", 1200
        lv.ColumnHeaders.Add , , "Resident", 1600
        lv.ColumnHeaders.Add , , "Event", 800
        lv.ColumnHeaders.Add , , "Alarm", 1500
        lv.ColumnHeaders.Add , , "Cancel", 1300
        lv.ColumnHeaders.Add , , "Elapsed", 850
        lv.ColumnHeaders.Add , , "Operator", 900

    End Select
    lv.Refresh
  End If
End Sub

Sub PrintBanner()
  ' date and time of report
  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
    Case DEST_PRINTER  ' if To Printer:


      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.ScaleMode = vbInches

      Printer.Print ""
      PrintCentered Printer, Printer.CurrentY, Title
      PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "")  ' hh:nn am/pm")

      Printer.Print ""
      Printer.Print ""


    Case DEST_FILE  ' if To File

  End Select
End Sub
Sub PrintFooter()
  ' copyright etc
  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
    Case DEST_PRINTER  ' if To Printer:
      Dim text           As String

      text = App.LegalCopyright & " " & App.ProductName & " (Ver. 1." & Format(App.Revision, "000") & ")"
      PrintCJ Printer, Printer.ScaleWidth / 2, Printer.ScaleHeight - Printer.TextHeight("A") * 1.5, text

    Case DEST_FILE  ' if To File

  End Select


End Sub


'Function GetAlarmHistory(rs As Recordset) As cAlarmHistory
'  Dim History            As cAlarmHistory
'  Set History = New cAlarmHistory
'
'  History.Parse rs
'  If History.ResidentID <> 0 Then
'    History.GetResidentInfo History.ResidentID
'    History.Resident = GetResidentName(History.ResidentID)
'    History.Phone = GetResidentName(History.ResidentID)
'  End If
'  History.Room = GetRoomName(History.RoomID)
'  History.GetHistory conn
'  Set GetAlarmHistory = History
'
''  Set GetAlarmHistory = History
'End Function


Private Sub Class_Initialize()

  Set rows = New Collection
  alarmtypes = Join(Array(EVT_EMERGENCY, EVT_ALERT, EVT_TAMPER, EVT_BATTERY_FAIL, EVT_CHECKIN_FAIL, EVT_COMM_TIMEOUT), ",")
End Sub

Private Sub Class_Terminate()
  Set rows = Nothing
End Sub

'Private Function GetReportDateString() As String
'  GetReportDateString = Format(ReportDate, "yyyymmdd")
'End Function



