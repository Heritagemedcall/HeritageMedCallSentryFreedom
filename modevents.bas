Attribute VB_Name = "modEvents"
Option Explicit

'' feiled array indecies
Global Const FLD_ID = 0  ' long
Global Const FLD_STATUS = 1  ' long
Global Const FLD_SERIAL = 2    ' 50 char
Global Const FLD_EVENTDATE = 3  ' date
Global Const FLD_ALARM = 4     ' long
Global Const FLD_TAMPER = 5    ' long
Global Const FLD_ISLOCATOR = 6  ' long
Global Const FLD_BATTERY = 7   ' long
Global Const FLD_HOPS = 8      ' long
Global Const FLD_FIRSTHOP = 9  ' 50 char
Global Const FLD_RESIDENTID = 10  ' long
Global Const FLD_ROOMID = 11   ' long
Global Const FLD_EVENTTYPE = 12  'long
Global Const FLD_ALARMID = 13  ' long - previous alarm's ID
Global Const FLD_USERNAME = 14  ' 50 char
Global Const FLD_SESSIONID = 15  ' long
Global Const FLD_ANNOUNCE = 16  ' 50 char
Global Const FLD_PHONE = 17    ' 50 char
Global Const FLD_INFO = 18     ' 255 char
Global Const FLD_USERDATA = 19  ' 128 char
Global Const FLD_SIGNAL = 20   ' long
Global Const FLD_MARGIN = 21   ' long
Global Const FLD_PACKET = 22   ' 255 char
Global Const FLD_FC1 = 23      ' long
Global Const FLD_FC2 = 24      ' long
Global Const FLD_IDM = 25      ' long
Global Const FLD_IDL = 26      ' long
Global Const FLD_LOCIDM = 27   ' long
Global Const FLD_LOCIDL = 28   ' long
Global Const FLD_INPUTNUM = 29   ' long


Global Const FIELDNAMES_CSV = "Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid,username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum"

Global Const NUM_FLDS = 29

' Status   Hex Decimal  Binary            Device   Event
'         0010     16  00000000 00010000  ??? Smoke Restore ???
'         0020     32  00000000 00100000  ??? TAMPER ???
'         0100    256  00000001 00000000  All Input 1
'         0200    512  00000010 00000000  All Input 2
'         14A1   5281  00010100 10100001  All? TAMPER?
'         4002  16386  01000000 00000010  Repeater EVT_LINELOSS = 40
'         400A  16394  01000000 00001010  Repeater EVT_LINELOSS = 40
'         4012  16402  01000000 00010010  Repeater EVT_LINELOSS = 40
'         4042  16450  01000000 01000010  Repeater EVT_BATTERY_FAIL = 4
'         7FE1  32737  01111111 11100001  All? TAMPER, BATTERY FAILURE  ' B217A3A0,B218DF33

' see Excel file: 'ES Status Recorded.xls'

' older
'         3FE1  16353  00111111 11100001  B218B78C
'         00E0    224  00000000 11100000  B218DF89
'         0FE0   4064  00001111 11100000  B218DF7D
'         1FE1   8161  00011111 11100001  B218B76F 4/22
'         3FE0  16352  00111111 11100000  B218B791 4/22
'         07E0   2016  00000111 11100000  B2181C2B


Global Packets          As New Collection
Global Devices          As New cESDevices
Global Residents        As New cResidents
Global Rooms            As New cRooms
Global Partitions       As New Collection

Global gPush   As Long
Global PushProcessor As cPushProcessor


'Global Const SYSTEM_6080      As Long = 1 ' 0 or 1


Global alarms           As cAlarms
Global Alerts           As cAlarms
Global Troubles         As cAlarms
Global LowBatts         As cAlarms
Global Assurs           As cAlarms
Global Externs          As cAlarms  ' external devices
'Global Assistances      As cAlarms  ' assistance calls



'Event constants

Global Const EVT_NONE = 0      ' NON-EVENT EVENT
Global Const EVT_EMERGENCY = 1  ' ALARM
Global Const EVT_EMERGENCY_RESTORE = 2  ' RESTORE AFTER ALARM
Global Const EVT_EMERGENCY_ACK = 3  ' MANUAL OR AUTO ACKNOWLEDGE
Global Const EVT_BATTERY_FAIL = 4  ' LOW BATTERY
Global Const EVT_BATTERY_RESTORE = 5  ' LOW BATTERY RESTORE
Global Const EVT_CHECKIN_FAIL = 6  ' DEVICE AUTO-CHECKIN FAILED
Global Const EVT_CHECKIN = 7   ' ONLY AFTER A FAILED CHECKIN
Global Const EVT_UNASSIGNED = 8  ' IN SYSTEM, BUT NOT ASSIGNED
Global Const EVT_STRAY = 9     ' NOT IN SYSTEM
Global Const EVT_COMM_TIMEOUT = 10  ' TOO MUCH TIME SINCE LAST COMM DATA / SERIAL PORT DEAD
Global Const EVT_COMM_RESTORE = 11  ' GETTING DATA AGAIN
Global Const EVT_ASSUR_CHECKIN = 12  ' CHECKED IN
Global Const EVT_ASSUR_FAIL = 13  ' FAILED TOCHECK IN
Global Const EVT_ALERT = 14    ' ALERT INSTEAD OF ALARM
Global Const EVT_ALERT_RESTORE = 15  ' ALERT INSTEAD OF ALARM
Global Const EVT_ALERT_ACK = 16  ' ALERT INSTEAD OF ALARM
Global Const EVT_SILENCE = 17  ' SILENCED (BUT NOT ACKNOWLEDGED)
Global Const EVT_LOCATED = 18  ' LOCATOR FORWARD
Global Const EVT_TAMPER = 19   ' TAMPER TRIGGERED
Global Const EVT_TAMPER_RESTORE = 20  ' TAMPER BIT RESTORED
Global Const EVT_ANNOUNCE_1 = 21  ' STANDARD ANNOUNCE
Global Const EVT_ANNOUNCE_2 = 22  ' ESCALATED ANNOUNCE
Global Const EVT_ANNOUNCE_3 = 23  ' 3RD LEVEL ESCALATED ANNOUNCE
Global Const EVT_DATABASE_UPDATE = 24  ' We've written to the database
Global Const EVT_DATABASE_READ = 25  ' We've checked the database for updates from other consoles
Global Const EVT_GENERAL_TROUBLE = 26  ' UNDEFINED TROUBLE
Global Const EVT_ASSUR_START = 27
Global Const EVT_ASSUR_END = 28
Global Const EVT_VACATION = 29
Global Const EVT_VACATION_RETURN = 30
Global Const EVT_EMERGENCY_END = 31
Global Const EVT_ALERT_END = 32
Global Const EVT_ADD_RES = 33
Global Const EVT_REMOVE_RES = 34
Global Const EVT_ADD_DEV = 35
Global Const EVT_REMOVE_DEV = 36
Global Const EVT_ASSIGN_DEV = 37
Global Const EVT_UNASSIGN_DEV = 38
Global Const EVT_LOCATE = 39
Global Const EVT_LINELOSS = 40
Global Const EVT_LINELOSS_RESTORE = 41
Global Const EVT_JAMMED = 42
Global Const EVT_JAMM_RESTORE = 43

Global Const EVT_SYSTEM_START = 44
Global Const EVT_SYSTEM_STOP = 45
Global Const EVT_SYSTEM_LOGIN = 46
Global Const EVT_SYSTEM_LOGOUT = 47

Global Const EVT_EXTERN = 48   ' EXTERN INSTEAD OF ALARM
Global Const EVT_EXTERN_RESTORE = 49
Global Const EVT_EXTERN_ACK = 50
Global Const EVT_EXTERN_END = 51

Global Const EVT_EXTERN_TROUBLE = 52  ' EXTERNAL DEVICE PORT FAILURE/CONNECTOR FAILURE
Global Const EVT_EXTERN_TROUBLE_RESTORE = 53


Global Const EVT_AUTOACK = 54  ' general device alarm autoclear
Global Const EVT_EMERGENCY_AUTOACK = 55  ' device emergency autoclear
Global Const EVT_ALERT_AUTOACK = 56  ' device alert autoclear
Global Const EVT_EXTERN_AUTOACK = 57  ' device extern autoclear

Global Const EVT_PTI_MISMATCH = 58  ' packet PTI and Device PTI don't match
Global Const EVT_STATUS_ERROR = 59  ' packet Status word too big => 32737
Global Const EVT_BATT_TAMPER = 60  ' packet Battery and Tamper in same packet
Global Const EVT_PCA_REG = 61  ' PCA registration packet... not an alarm

Global Const EVT_FORCED_LOGOUT = 62
Global Const EVT_MAXDEVICE = 63


Global Const EVT_SERVER_TROUBLE = 64  ' Output Server Failure
Global Const EVT_SERVER_TROUBLE_RESTORE = 65


Global Const EVT_EMERGENCY_RESPOND = 66
Global Const EVT_ALERT_RESPOND = 67
Global Const EVT_GENERIC_RESPOND = 68
Global Const EVT_EXTERN_RESPOND = 69

Global Const EVT_EMERGENCY_FINALIZE = 70
Global Const EVT_ALERT_FINALIZE = 71
Global Const EVT_GENERIC_FINALIZE = 72
Global Const EVT_EXTERN_FINALIZE = 73


Global Const EVT_ASSISTANCE = 74
Global Const EVT_ASSISTANCE_RESPOND = 75
Global Const EVT_ASSISTANCE_FINALIZE = 76
Global Const EVT_ASSISTANCE_ACK = 77
Global Const EVT_ASSISTANCE_RESTORE = 78
Global Const EVT_ASSISTANCE_END = 79



Public Sub CreateEventtypes()
  Dim j As Integer
  Dim evt As cEventType

10        On Error GoTo CreateEventtypes_Error

20     EventNames = Array("None", "Alarm", "Restore", "Ack", "Low-Batt", "Batt-OK", "Super-Fail", "Super-OK", "Unassigned", "Stray", "Comm-Fail", "Comm_OK", "Assur-OK", "Assur-Fail", "Alert", "Alert-Restore", "Alert-Ack", "Silence", "Located", _
      "Tamper", "Tamper-Restore", "Announce-1", "Announce-2", "Announce-3", "Data-Update", "Data-Read", "General-Fail", "Assure Start", "Assure End", "Vacation", "Vac-Return", "END Emergency", "END Alert", "ADD Resident", "Remove Resident", _
      "ADD Device", "Remove Device", "Assign Device", "Unassign Device", "Locate", "LineLoss", "Line Restore", "Jammed", "Jam Restore", "System Start", "System Stop", "Login", "Logout", _
      "Extern Event", "Extern-Restore", "Extern-ACK", "END Extern", "Extern-Trouble", "Extern-OK", "AutoACK", "Alarm-AutoACK", "Alert-AutoACK", "Extern-AutoACK", "PTI-Mismatch", "Status-Error", "Batt-Tamper", _
      "PCA-Reg", "Forced-Logout", "Too-Many-TX", "Server Trouble", "Server Restore", "Emergency Respond", "Alert Respond", "Generic Respond", "Extern Respond", "Emergency Finalize", "Alert Finalize", "Generic Finalize", "Extern Finalize", _
      "Assistance", "Assistance Respond", "Assistance Finalize", "Assistance ACK", "Assistance Restore", "Assistance End")
      
      
      

30        Set EventTypes = New Collection
40        For j = 0 To UBound(EventNames)
50          Set evt = New cEventType
60          evt.TypeID = j + 1
70          evt.TypeName = EventNames(j)
80          EventTypes.Add evt
90        Next


          For j = 0 To UBound(EventNames)
            Debug.Print EventNames(j)
          Next

CreateEventtypes_Resume:
100       On Error GoTo 0
110       Exit Sub

CreateEventtypes_Error:

120       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.CreateEventtypes." & Erl
130       Resume CreateEventtypes_Resume


End Sub

Public Function CreateAssistanceAlarm(PriorAlarm As cAlarm, ByVal LinkedAlarm As Long) As cESPacket
  Dim packet As cESPacket
  Dim Device As cESDevice
  Dim evt As cEvent
  Dim Inputnum As Long
  Dim alarm As cAlarm
  
  Set alarm = New cAlarm
  Set packet = New cESPacket
  
  Set Device = Devices.Device(PriorAlarm.Serial)
  If Device Is Nothing Then
    Set CreateAssistanceAlarm = Nothing
    Exit Function
  End If
  
  Inputnum = PriorAlarm.Inputnum
  packet.LinkedAlarm = PriorAlarm.AlarmID
  packet.Alarmtype = EVT_ASSISTANCE
  
  Set evt = PostEvent(Device, packet, PriorAlarm, EVT_ASSISTANCE, Inputnum)
  
    
  
      
      
     
  


End Function


Public Function GetEventName(ByVal ID As Long) As String
10        On Error GoTo GetEventName_Error

20        If ID >= LBound(EventNames) And ID <= UBound(EventNames) Then
30          GetEventName = EventNames(ID)
40        Else
50          GetEventName = "ID " & ID
60        End If

GetEventName_Resume:
70        On Error GoTo 0
80        Exit Function

GetEventName_Error:

90        LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.GetEventName." & Erl
100       Resume GetEventName_Resume

End Function



Function DispatchAlarm(alarm As cAlarm)
  ' send data over net to subscribers
  ' serialize object
  
End Function

Function UpdateDispositions(ByVal Disposition As String) As Long
  Dim SQL As String
  Dim rs As ADODB.Recordset
  Dim Count As Long
  
  SQL = "select count(text) as c from dispositions where text = " & q(Disposition)
  Set rs = ConnExecute(SQL)
  Count = rs("c")
  rs.Close
  Set rs = Nothing
  If Count = 0 Then
    SQL = "INSERT INTO dispositions (text) values (" & q(Disposition) & ")"
    ConnExecute SQL
  End If
  
  
  
  
End Function


Function PostEvent(d As cESDevice, p As cESPacket, alarm As cAlarm, ByVal EventType As Long, Inputnum As Long, Optional ByVal Username As String = "System") As cEvent

        Dim PageRequest        As cPageItem
        Dim sn                 As String

        Dim AlarmID            As Long
10      If Not alarm Is Nothing Then
20        AlarmID = alarm.AlarmID
30      End If

40      If Inputnum = 0 Then
50        Inputnum = 1
60      End If

70      On Error GoTo PostEvent_Error

80      If p Is Nothing Then
90        Set p = New cESPacket
100     End If

110     If d Is Nothing Then
120       Set d = New cESDevice
130     End If

140     If MASTER Then

150       Select Case EventType

              ' begin Finalize History


            Case EVT_EMERGENCY_FINALIZE, EVT_ALERT_FINALIZE, EVT_GENERIC_FINALIZE, EVT_EXTERN_FINALIZE
160           LogAlarm alarm, EventType, Username  ' need to track who did this


              ' ************** begin of STAFF ASSIST ***************

170         Case EVT_ASSISTANCE

              'inputnum = p.inputnum  ' need to get inputnum thru chain of paging stuff

              'd.alarm = 1
              'd.LastAlarm = p.DateTime
180           p.Alarmtype = EVT_ASSISTANCE

190           If alarm Is Nothing Then  ' no prior alarm
200             Exit Function
210           End If

              Dim PriorAlarm   As cAlarm
220           Set PriorAlarm = alarm
230           Set alarm = Nothing

240           Set alarm = alarms.AddAssist(p, d, Inputnum, PriorAlarm)  ' need to get inputnum thru prioralarm

250           If Not alarm Is Nothing Then
260             alarm.Alarmtype = EVT_ASSISTANCE
                alarm.PriorID = 0
270             LogAlarm alarm, EVT_ASSISTANCE, ""  ' "Assistance"
280             AddPageRequest alarm, alarm.Alarmtype
290             frmMain.ProcessAlarms
300           End If

            ''''''''''''' EVT_ASSISTANCE_ACK

310         Case EVT_ASSISTANCE_ACK  ' when acked from console, user is warned that ack will end assistance call
320           If alarm.ACKed = 0 Then

330             LogAlarm alarm, EVT_ASSISTANCE_ACK, alarm.Username
                
                'mobileACK Alarm
340             alarms.AcknowledgeAlarm d, alarm, EVT_ASSISTANCE
350             Trace "ASSISTANCE ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
360             frmMain.ProcessAlarms


370           End If

              'If d.alarm = 0 Then ' clear alarm regardless of alarmed state, they can always call for assistance again
380           ClearInfoBox alarm.AlarmID
390           ClearPages alarm.AlarmID
400           alarms.RemoveAlarm alarm, EVT_ASSISTANCE
410           PostEvent d, Nothing, alarm, EVT_ASSISTANCE_END, Inputnum, Username

420           frmMain.ProcessAlarms
              'End If

430         Case EVT_ASSISTANCE_RESTORE  ' there is no restore event at this time of writing

440         Case EVT_ASSISTANCE_END
450           DeleteFromMobile alarm.ID

460           Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_EMERGENCY, InputNum)
470           LogAlarm alarm, EVT_ASSISTANCE_END, Username  ' need to track who did this
480           Trace "Emergency END: " & Right("00000000" & d.Serial, 8) & " " & Now

              ' ************** end of STAFF ASSIST ***************


490         Case EVT_PTI_MISMATCH
500           SpecialLog "PTI Mismatch " & d.CLS & "/" & p.PTI & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now
510           Debug.Print "PTI Mismatch " & d.CLS & "/" & p.PTI & " " & p.Serial & " " & p.HexPacket & " " & Now

520         Case EVT_STATUS_ERROR
530           SpecialLog "Status Error " & p.Status & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now
540           Debug.Print "Status Error " & p.Status & " " & p.Serial & " " & p.HexPacket & " " & Now

550         Case EVT_BATT_TAMPER
560           SpecialLog "Batt-Tamper Error " & p.Status & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now
570           Debug.Print "Batt-Tamper Error " & p.Status & " " & p.Serial & " " & p.HexPacket & " " & Now

580         Case EVT_PCA_REG
590           SpecialLog "PCA Reg " & p.Status & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now
600           Debug.Print "PCA Reg " & p.Status & " " & p.Serial & " " & p.HexPacket & " " & Now

610         Case EVT_EXTERN
620           Inputnum = Win32.timeGetTime()  ' what is this??

630           d.alarm = 0
640           d.LastAlarm = p.DateTime

650           p.Alarmtype = EVT_EXTERN
660           Set alarm = Externs.Add(p, d, Inputnum)
670           If Not alarm Is Nothing Then
680             alarm.Alarmtype = EVT_EXTERN
690             LogAlarm alarm, EVT_EXTERN, gUser.Username
700             AddPageRequest alarm, alarm.Alarmtype

710             Trace "External Event: " & Right("00000000" & d.Serial, 8) & " " & Now & " '" & p.SerialPacket & "'", True
720           End If


              'case EVT_EXTERN_RESTORE

730         Case EVT_EXTERN_ACK, EVT_EXTERN_AUTOACK

740           If alarm.ACKed = 0 Then
750             LogAlarm alarm, EventType, alarm.Username
760             mobileACK alarm
770             alarms.AcknowledgeAlarm d, alarm, EVT_EXTERN
780             Trace "Extern ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
790           End If
800           If d.IsSerialDevice Then
810             d.alarm = 0
820           End If
830           If d.alarm = 0 Then

840             ClearInfoBox alarm.AlarmID
850             ClearPages alarm.AlarmID
860             Externs.RemoveAlarm alarm, EVT_EXTERN
870             PostEvent d, Nothing, alarm, EVT_EXTERN_END, alarm.Inputnum
880             frmMain.ProcessExterns
890             Set alarm = Nothing  ' to eliminate memory leak
900           End If

910         Case EVT_EXTERN_END
              ' alarm should be valid object here

920           Set PageRequest = RemovePageRequest(alarm.ID)  ' Alarm.Serial, EVT_EXTERN, Alarm.InputNum)

              'If Not PageRequest Is Nothing Then
              '  If d.SendCancel = 1 Then
              '    SendEndofEventPage PageRequest, EVT_EXTERN
              '  End If
              'End If
930           LogAlarm alarm, EVT_EXTERN_END, gUser.Username
940           Trace "Extern END: " & Right("00000000" & d.Serial, 8) & " " & Now


              '********** EMERGENCY *************
950         Case EVT_EMERGENCY       ' ALARM  ' MANUAL-CLEAR

960           If Inputnum = 3 Then
970             d.Alarm_B = 1
980             d.LastAlarm_B = p.DateTime

990           ElseIf Inputnum = 2 Then
1000            d.Alarm_A = 1
1010            d.LastAlarm_A = p.DateTime
1020          Else
1030            Inputnum = 1         ' just in csae input num = 0
1040            d.alarm = 1
1050            d.LastAlarm = p.DateTime
1060          End If
1070          p.Alarmtype = EVT_EMERGENCY
              ' get resident/room ' if assursecure and away then '  Add to alarms

1080          Set alarm = alarms.GetAlarm(d, Inputnum)
1090          If alarm Is Nothing Then
1100            Debug.Print "adding alarm to inbounds " & Format(Now, "nn:ss")
1110            Set alarm = InBounds.Add(p, d, Inputnum)
1120            Debug.Print "InBounds Count " & InBounds.Count
1130          End If
1140          If USE6080 Then
1150            alarm.locationtext = d.LastLocationText
1160          End If

1170          LogAlarm alarm, EVT_EMERGENCY, gUser.Username
1180          Trace "Emergency Alarm: " & Right("00000000" & alarm.Serial, 8) & " " & Now

1190        Case EVT_EMERGENCY_RESTORE  ' RESTORE AFTER ALARM

1200          If Inputnum = 3 Then
1210            d.Alarm_B = 0
1220            d.LastRestore_B = p.DateTime
1230            Set alarm = alarms.Restore(d, Inputnum)
1240            If Not alarm Is Nothing Then
1250              alarm.alarm = 0
                  'LogRestore Alarm
1260              LogAlarm alarm, EVT_EMERGENCY_RESTORE, gUser.Username
1270              Trace "Emergency RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1280              If alarm.ACKed <> 0 Then
1290                ClearInfoBox alarm.AlarmID
1300                ClearPages alarm.AlarmID
1310                alarms.RemoveAlarm alarm, EVT_EMERGENCY

1320                PostEvent d, Nothing, alarm, EVT_EMERGENCY_END, Inputnum
1330                frmMain.ProcessAlarms
1340              End If
1350            Else
1360              Set alarm = InBounds.BySerial(d.Serial)
1370              If Not alarm Is Nothing Then
1380                LogAlarm alarm, EVT_EMERGENCY_RESTORE, gUser.Username
1390                Trace "Emergency RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1400              End If
1410            End If
1420          ElseIf Inputnum = 2 Then
1430            d.Alarm_A = 0
1440            d.LastRestore_A = p.DateTime
1450            Set alarm = alarms.Restore(d, Inputnum)
1460            If Not alarm Is Nothing Then
1470              alarm.alarm = 0
                  'LogRestore Alarm
1480              LogAlarm alarm, EVT_EMERGENCY_RESTORE, gUser.Username
1490              Trace "Emergency RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1500              If alarm.ACKed <> 0 Then
1510                ClearInfoBox alarm.AlarmID
1520                ClearPages alarm.AlarmID
1530                alarms.RemoveAlarm alarm, EVT_EMERGENCY
1540                PostEvent d, Nothing, alarm, EVT_EMERGENCY_END, Inputnum
1550                frmMain.ProcessAlarms
1560              End If
1570            Else
1580              Set alarm = InBounds.BySerial(d.Serial)
1590              If Not alarm Is Nothing Then
1600                LogAlarm alarm, EVT_EMERGENCY_RESTORE, gUser.Username
1610                Trace "Emergency RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1620              End If
1630            End If
1640          Else                   ' input 1
1650            d.alarm = 0
1660            d.LastRestore = p.DateTime
1670            Set alarm = alarms.Restore(d, Inputnum)
1680            If Not alarm Is Nothing Then
1690              alarm.alarm = 0
                  'LogRestore Alarm
1700              LogAlarm alarm, EVT_EMERGENCY_RESTORE, gUser.Username
1710              Trace "Emergency RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1720              If alarm.ACKed <> 0 Then
1730                ClearInfoBox alarm.AlarmID
1740                ClearPages alarm.AlarmID
1750                alarms.RemoveAlarm alarm, EVT_EMERGENCY
1760                PostEvent d, Nothing, alarm, EVT_EMERGENCY_END, Inputnum
1770                frmMain.ProcessAlarms
1780              End If
1790            Else
1800              Set alarm = InBounds.BySerial(d.Serial)
1810              If Not alarm Is Nothing Then
1820                LogAlarm alarm, EVT_EMERGENCY_RESTORE, gUser.Username
1830                Trace "Emergency RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1840              End If
1850            End If
1860          End If

1870        Case EVT_EMERGENCY_ACK, EVT_EMERGENCY_AUTOACK  ' MANUAL OR AUTO ACKNOWLEDGE
1880          If Inputnum = 3 Then
1890            If alarm.ACKed = 0 Then
1900              LogAlarm alarm, EventType, alarm.Username
1910              mobileACK alarm
1920              alarms.AcknowledgeAlarm d, alarm, EVT_EMERGENCY

1930              Trace "Emergency ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
1940            End If
                '        If d.IsSerialDevice Then
                '          d.Alarm_A = 0
                '        End If
1950            If d.Alarm_B = 0 Then  ' remove it now if already restored

1960              alarms.RemoveAlarm alarm, EVT_EMERGENCY
1970              ClearInfoBox alarm.AlarmID
1980              ClearPages alarm.AlarmID
1990              PostEvent d, Nothing, alarm, EVT_EMERGENCY_END, Inputnum
2000              frmMain.ProcessAlarms
2010            End If

2020          ElseIf Inputnum = 2 Then
2030            If alarm.ACKed = 0 Then
2040              LogAlarm alarm, EventType, alarm.Username
2050              mobileACK alarm
2060              alarms.AcknowledgeAlarm d, alarm, EVT_EMERGENCY
2070              Trace "Emergency ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
2080            End If
                '        If d.IsSerialDevice Then
                '          d.Alarm_A = 0
                '        End If
2090            If d.Alarm_A = 0 Then  ' remove it now if already restored
2100              alarms.RemoveAlarm alarm, EVT_EMERGENCY
2110              ClearInfoBox alarm.AlarmID
2120              ClearPages alarm.AlarmID
2130              PostEvent d, Nothing, alarm, EVT_EMERGENCY_END, Inputnum
2140              frmMain.ProcessAlarms
2150            End If

2160          Else                   ' input 1
2170            If alarm.ACKed = 0 Then
2180              LogAlarm alarm, EVT_EMERGENCY_ACK, alarm.Username
2190              mobileACK alarm
2200              alarms.AcknowledgeAlarm d, alarm, EVT_EMERGENCY

2210              Trace "Emergency ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
2220            End If
2230            If d.IsSerialDevice Then
2240              d.alarm = 0
2250            End If
2260            If d.alarm = 0 Then

2270              ClearInfoBox alarm.AlarmID
2280              ClearPages alarm.AlarmID
2290              alarms.RemoveAlarm alarm, EVT_EMERGENCY
2300              PostEvent d, Nothing, alarm, EVT_EMERGENCY_END, Inputnum
2310              frmMain.ProcessAlarms
2320            End If

2330          End If


2340        Case EVT_EMERGENCY_END   ' Emergency Ended

2350          If Inputnum = 3 Then
2360            Set PageRequest = RemovePageRequest(alarm.ID)  '.Serial, EVT_EMERGENCY, InputNum)
2370            LogAlarm alarm, EVT_EMERGENCY_END, gUser.Username
2380            Trace "Emergency END: " & Right("00000000" & d.Serial, 8) & " " & Now




2390          ElseIf Inputnum = 2 Then
2400            Set PageRequest = RemovePageRequest(alarm.ID)  '.Serial, EVT_EMERGENCY, InputNum)
2410            LogAlarm alarm, EVT_EMERGENCY_END, gUser.Username
2420            Trace "Emergency END: " & Right("00000000" & d.Serial, 8) & " " & Now

2430          Else
2440            Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_EMERGENCY, InputNum)
2450            LogAlarm alarm, EVT_EMERGENCY_END, gUser.Username
2460            Trace "Emergency END: " & Right("00000000" & d.Serial, 8) & " " & Now
2470          End If

              '********** ALERT *************

2480        Case EVT_ALERT           ' ALERT SCREEN ' NOT+++ AUTO-CLEAR


2490          If Inputnum = 3 Then
2500            d.Alarm_B = 1
2510            d.LastAlarm_B = p.DateTime


2520          ElseIf Inputnum = 2 Then
2530            d.Alarm_A = 1
2540            d.LastAlarm_A = p.DateTime
2550          Else
2560            d.alarm = 1
2570            d.LastAlarm = p.DateTime
2580          End If
2590          p.Alarmtype = EVT_ALERT
              ' get resident/room ' if assursecure and away then '  Add to alarms
2600          Set alarm = InBounds.Add(p, d, Inputnum)
2610          If USE6080 Then
                'Alarm.locationtext = p.LocatedPartionName1
2620            alarm.locationtext = d.LastLocationText
2630          End If

2640          LogAlarm alarm, EVT_ALERT, gUser.Username
2650          Trace "Alert Alarm: " & Right("00000000" & alarm.Serial, 8) & " " & Now

2660        Case EVT_ALERT_RESTORE   ' RESTORE AFTER ALARM


2670          If Inputnum = 3 Then
2680            d.Alarm_B = 0
2690            d.LastRestore_B = p.DateTime
2700            Set alarm = Alerts.Restore(d, Inputnum)
2710            If Not alarm Is Nothing Then
2720              alarm.alarm = 0
                  'LogRestore Alarm
2730              LogAlarm alarm, EVT_ALERT_RESTORE, gUser.Username
2740              Trace "Alert RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
2750              If alarm.ACKed <> 0 Then
2760                ClearInfoBox alarm.AlarmID
2770                ClearPages alarm.AlarmID
2780                Alerts.RemoveAlarm alarm, EVT_ALERT
2790                PostEvent d, Nothing, alarm, EVT_ALERT_END, Inputnum
2800                frmMain.ProcessAlarms

2810              End If
2820              Set alarm = Nothing  ' to eliminate memory leak
2830            End If


2840          ElseIf Inputnum = 2 Then
2850            d.Alarm_A = 0
2860            d.LastRestore_A = p.DateTime
2870            Set alarm = Alerts.Restore(d, Inputnum)
2880            If Not alarm Is Nothing Then
2890              alarm.alarm = 0
                  'LogRestore Alarm
2900              LogAlarm alarm, EVT_ALERT_RESTORE, gUser.Username
2910              Trace "Alert RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
2920              If alarm.ACKed <> 0 Then
2930                ClearInfoBox alarm.AlarmID
2940                ClearPages alarm.AlarmID
2950                Alerts.RemoveAlarm alarm, EVT_ALERT
2960                PostEvent d, Nothing, alarm, EVT_ALERT_END, Inputnum
2970                frmMain.ProcessAlarms
2980              End If
2990              Set alarm = Nothing  ' to eliminate memory leak
3000            End If

3010          Else                   ' input #1
3020            d.alarm = 0
3030            d.LastRestore = p.DateTime
3040            Set alarm = Alerts.Restore(d, Inputnum)
3050            If Not alarm Is Nothing Then
3060              alarm.alarm = 0
3070              LogAlarm alarm, EVT_ALERT_RESTORE, gUser.Username
3080              Trace "Alert RESTORE: " & Right("00000000" & alarm.Serial, 8) & " " & Now
3090              If alarm.ACKed <> 0 Then
3100                ClearInfoBox alarm.AlarmID
3110                ClearPages alarm.AlarmID
3120                Alerts.RemoveAlarm alarm, EVT_ALERT
3130                PostEvent d, Nothing, alarm, EVT_ALERT_END, Inputnum
3140                frmMain.ProcessAlarms
3150              End If
3160              Set alarm = Nothing  ' to eliminate memory leak
3170            End If
3180          End If


3190        Case EVT_ALERT_ACK, EVT_ALERT_AUTOACK  ' MANUAL OR AUTO ACKNOWLEDGE


3200          If Inputnum = 3 Then
3210            If alarm.ACKed = 0 Then
                  'LogAcknowledge Alarm
3220              LogAlarm alarm, EventType, alarm.Username
3230              mobileACK alarm
3240              Alerts.AcknowledgeAlarm d, alarm, EVT_ALERT
3250              Trace "Emergency ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
3260              If d.Alarm_B = 0 Then  ' fixed 2014-02-11 inout 3 would auto clear itself

3270                ClearInfoBox AlarmID
3280                ClearPages AlarmID
3290                Alerts.RemoveAlarm alarm, EVT_ALERT
3300                PostEvent d, Nothing, alarm, EVT_ALERT_END, Inputnum
3310                frmMain.ProcessAlarms
3320                Set alarm = Nothing  ' to eliminate memory leak
3330              End If

3340            End If

3350          ElseIf Inputnum = 2 Then
3360            If alarm.ACKed = 0 Then
                  'LogAcknowledge Alarm
3370              LogAlarm alarm, EventType, alarm.Username
3380              mobileACK alarm
3390              Alerts.AcknowledgeAlarm d, alarm, EVT_ALERT
3400              Trace "Emergency ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
3410              If d.Alarm_A = 0 Then

3420                ClearInfoBox alarm.AlarmID
3430                ClearPages alarm.AlarmID
3440                Alerts.RemoveAlarm alarm, EVT_ALERT
3450                PostEvent d, Nothing, alarm, EVT_ALERT_END, Inputnum
3460                frmMain.ProcessAlarms
3470                Set alarm = Nothing  ' to eliminate memory leak
3480              End If

3490            End If
3500          Else
3510            If alarm.ACKed = 0 Then
3520              LogAlarm alarm, EVT_ALERT_ACK, alarm.Username
3530              mobileACK alarm
3540              Alerts.AcknowledgeAlarm d, alarm, EVT_ALERT
3550              Trace "Alert ACK: " & Right("00000000" & alarm.Serial, 8) & " " & Now
3560              If d.IsSerialDevice Then
3570                d.alarm = 0
3580              End If
3590              If d.alarm = 0 Then

3600                ClearInfoBox alarm.AlarmID
3610                ClearPages alarm.AlarmID
3620                Alerts.RemoveAlarm alarm, EVT_ALERT
3630                PostEvent d, Nothing, alarm, EVT_ALERT_END, Inputnum
3640                frmMain.ProcessAlarms
3650                Set alarm = Nothing  ' to eliminate memory leak
3660              End If

3670            End If
3680          End If

3690        Case EVT_ALERT_END       ' Emergency Ended
3700          If Inputnum = 3 Then

3710            Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_ALERT, InputNum)
3720            LogAlarm alarm, EVT_ALERT_END, gUser.Username
3730            Trace "Alert END: " & Right("00000000" & d.Serial, 8) & " " & Now

3740          ElseIf Inputnum = 2 Then

3750            Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_ALERT, InputNum)
3760            LogAlarm alarm, EVT_ALERT_END, gUser.Username
3770            Trace "Alert END: " & Right("00000000" & d.Serial, 8) & " " & Now


                '2440            SendEndofEventPage PageRequest, EVT_ALERT_END
                '***************
                'If Not PageRequest Is Nothing Then
                '  If d.SendCancel = 1 Then
                '    SendEndofEventPage PageRequest, EVT_ALERT_END
                '  End If
                'End If

3780          Else
3790            Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_ALERT, InputNum)

                'If Not PageRequest Is Nothing Then
                '  If d.SendCancel = 1 Then
                '    SendEndofEventPage PageRequest, EVT_ALERT_END
                '  End If
                'End If

                '2490            SendEndofEventPage PageRequest, EVT_ALERT_END
                '***************

3800            LogAlarm alarm, EVT_ALERT_END, gUser.Username
3810            Trace "Alert END: " & Right("00000000" & d.Serial, 8) & " " & Now
3820          End If

              '********** BATTERY *************
              ' need to align event logging with the rest.
3830        Case EVT_BATTERY_FAIL    ' LOW BATTERY

              'If d.Battery = 0 Then
3840          If 1 = 2 Then          ' = disable eating of low batts for these models ' d.Model = "EN1941" Or d.Model = "EN1223S" Or d.Model = "ES1233S" Then
3850            d.Battery = 1
3860            SpecialLog "Battery  " & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now
3870          Else
3880            d.STAT = p.Status    ' Or &HFFFF& ' threw an error here
3890            d.Battery = 1
3900            p.Alarmtype = EVT_BATTERY_FAIL
3910            Set alarm = LowBatts.Add(p, d, 0)

3920            If Not alarm Is Nothing Then
3930              alarm.Alarmtype = EVT_BATTERY_FAIL
3940              alarm.Announce = "Battery Failure"
3950              LogAlarm alarm, EVT_BATTERY_FAIL, gUser.Username
3960              AddPageRequest alarm, alarm.Alarmtype

3970              Trace "Battery Fail: " & Right("00000000" & d.Serial, 8) & " " & Now
3980            End If
3990          End If
              'End If
4000        Case EVT_BATTERY_RESTORE  ' LOW BATTERY RESTORE
4010          d.STAT = p.Status
4020          d.Battery = 0

4030          Set alarm = LowBatts.RemoveAlarm(d, EVT_BATTERY_FAIL)
4040          If Not alarm Is Nothing Then
4050            alarm.LastRestore = p.DateTime
4060            alarm.STAT = d.STAT
4070            alarm.packet = p.HexPacket
4080            ClearPages AlarmID
4090            LogAlarm alarm, EVT_BATTERY_RESTORE, gUser.Username
4100            Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_BATTERY_FAIL, 0)
                'If Not PageRequest Is Nothing Then
                '  If PageRequest.SendCancel Then
                '    SendEndofEventPage PageRequest, EVT_BATTERY_FAIL
                '  End If
                'End If
4110            Set alarm = Nothing  ' to eliminate memory leak
4120          End If

4130          Trace "Battery Restore: " & Right("00000000" & d.Serial, 8) & " " & Now

              '********** CHECKIN *************

4140        Case EVT_CHECKIN_FAIL    ' DEVICE AUTO-CHECKIN FAILED

              '              If d.Model = "EN6040" Then
              '                Debug.Assert 0
              '              End If

4150          p.Alarmtype = EVT_CHECKIN_FAIL
4160          d.Dead = 1

4170          Set alarm = Troubles.Add(p, d, 0)
4180          If Not (alarm Is Nothing) Then


4190            alarm.Alarmtype = EVT_CHECKIN_FAIL
4200            alarm.Announce = "Checkin Failure"
4210            LogAlarm alarm, EVT_CHECKIN_FAIL, gUser.Username
4220            AddPageRequest alarm, alarm.Alarmtype


4230            Trace "Checkin Fail: " & Right("00000000" & d.Serial, 8) & " " & Now
4240            Set alarm = Nothing  ' to eliminate memory leak
4250          End If

4260        Case EVT_CHECKIN         ' ONLY AFTER A FAILED CHECKIN
4270          d.Dead = 0

              '              If d.Model = "EN6040" Then
              '                Debug.Assert 0
              '              End If

4280          Set alarm = Troubles.RemoveAlarm(d, EVT_CHECKIN_FAIL)
4290          ClearPages AlarmID
4300          If Not alarm Is Nothing Then
4310            Set PageRequest = RemovePageRequest(alarm.ID)  ' Set PageRequest = RemovePageRequest(d.Serial, EVT_CHECKIN_FAIL, 0)
4320            LogAlarm alarm, EVT_CHECKIN, gUser.Username
4330            Set alarm = Nothing  ' to eliminate memory leak
4340          End If
4350          Trace "Checkin Restore: " & Right("00000000" & d.Serial, 8) & " " & Now


              '********** AC Fail *************

4360        Case EVT_LINELOSS

4370          p.Alarmtype = EVT_LINELOSS
4380          d.LineLoss = 1
4390          Set alarm = Troubles.Add(p, d, 0)
4400          If Not alarm Is Nothing Then
4410            alarm.Alarmtype = EVT_LINELOSS
4420            alarm.Announce = "Line Loss"
4430            LogAlarm alarm, EVT_LINELOSS, gUser.Username

4440            AddPageRequest alarm, alarm.Alarmtype
4450          End If

4460          Trace "Line Loss: " & Right("00000000" & d.Serial, 8) & " " & Now

4470        Case EVT_LINELOSS_RESTORE  '

4480          d.LineLoss = 0
4490          Set alarm = Troubles.RemoveAlarm(d, EVT_LINELOSS)

4500          If Not alarm Is Nothing Then
4510            ClearPages alarm.AlarmID
4520            alarm.LastRestore = p.DateTime
4530            alarm.STAT = d.STAT
4540            alarm.packet = p.HexPacket
4550            Set PageRequest = RemovePageRequest(alarm.ID)  'Set PageRequest = RemovePageRequest(Alarm.Serial, EVT_LINELOSS, 0)
4560            LogAlarm alarm, EVT_LINELOSS, gUser.Username
4570            Set alarm = Nothing  ' to eliminate memory leak
4580          End If


4590          Trace "Line Restore: " & Right("00000000" & d.Serial, 8) & " " & Now


              '********** TX ANOMOLIES *************

4600        Case EVT_UNASSIGNED      ' IN SYSTEM, BUT NOT ASSIGNED
4610          p.Alarmtype = EVT_UNASSIGNED

4620        Case EVT_STRAY           ' NOT IN SYSTEM
4630          p.Alarmtype = EVT_STRAY
4640          If gNoStrayData = False Then
4650            LogToStrays p
4660          End If

              'Trace "Transmitter not in system: " & Right("00000000" & p.serial, 8) & " " & Now
              'WriteEventToDB EVT

4670        Case EVT_COMM_TIMEOUT    ' TOO MUCH TIME SINCE LAST COMM DATA / SERIAL PORT DEAD
4680          d.Dead = 1
4690          p.Alarmtype = EVT_COMM_TIMEOUT
4700          Set alarm = Troubles.Add(p, d, 0)
4710          If Not alarm Is Nothing Then
4720            If p.Alarmtype = EVT_COMM_TIMEOUT And USE6080 = 0 Then
                  ' show comm fail
4730              frmMain.ShowCommError True
4740            End If




4750            alarm.Alarmtype = EVT_COMM_TIMEOUT
4760            alarm.Announce = "Communications Error"
4770            LogAlarm alarm, EVT_COMM_TIMEOUT, gUser.Username
4780            AddPageRequest alarm, alarm.Alarmtype
4790            Set alarm = Nothing  ' to eliminate memory leak
4800          End If

4810          Trace "Comm Timeout: " & Right("00000000" & d.Serial, 8) & " " & Now

4820        Case EVT_COMM_RESTORE    ' GETTING DATA AGAIN
4830          d.Dead = 0
4840          Set alarm = Troubles.RemoveAlarm(d, EVT_COMM_TIMEOUT)
4850          If USE6080 = 0 Then
                ' hide comm fail

4860            frmMain.ShowCommError False
4870          End If



4880          If Not alarm Is Nothing Then
4890            alarm.LastRestore = p.DateTime
4900            alarm.STAT = d.STAT
4910            If p.Is6080 Then
4920              alarm.packet = left$("", 255)
4930            Else
4940              alarm.packet = left$(p.HexPacket, 255)
4950            End If
4960            Set PageRequest = RemovePageRequest(alarm.ID)  'Set PageRequest = RemovePageRequest(Alarm.Serial, EVT_COMM_TIMEOUT, 0)
4970            LogAlarm alarm, EVT_COMM_RESTORE, gUser.Username
4980            Set alarm = Nothing  ' to eliminate memory leak
                'LogCommEvent d, p, EventType
4990          End If

5000          Trace "Comm Restore: " & Right("00000000" & d.Serial, 8) & " " & Now

              '********** ASSURANCE *************

5010        Case EVT_ASSUR_START     ' START OF ASSURANCE PERIOD
5020          Trace "Assure START: " & Now
5030          LogAssur EVT_ASSUR_START  ' not device or resident specific

5040        Case EVT_ASSUR_END       ' END OF ASSURANCE PERIOD
5050          Trace "Assure END: " & Now
5060          If Assurs.Count = 0 Then
5070            Assurs.Add p, d, 0   ' not sure why this was in here
5080          End If
5090          LogAssur EVT_ASSUR_END  ' not device or resident specific
5100          frmMain.ProcessAssurs True

5110        Case EVT_ASSUR_CHECKIN   ' CHECKED IN
5120          LogAssurCheckin d, p, EVT_ASSUR_CHECKIN  ', gUser.username
              ' check-in all devices for the same room if room is so flagged
5130          AllCheckin d

5140          Trace "Assure Checkin: " & Right("00000000" & p.Serial, 8) & " " & Now

5150        Case EVT_ASSUR_FAIL      ' FAILED TOCHECK IN
5160          p.Alarmtype = EVT_ASSUR_FAIL
5170          Assurs.Add p, d, 0
5180          LogAssurFail d, EVT_ASSUR_FAIL  ',guser.username
5190          Trace "Assure Fail: " & Right("00000000" & d.Serial, 8) & " " & Now

              '********** LOCATE *************

5200        Case EVT_LOCATE          ' LOCATOR FORWARD
              ' not used
5210        Case EVT_NONE            ' NON-EVENT EVENT

              '********** TAMPER *************

5220        Case EVT_TAMPER          ' TAMPER TRIGGERED ' treat almost like a general alarm
              'WriteEventToDB EVT

5230          If d.NoTamper = 1 Then  ' = "EN1941" Or d.Model = "EN1223S" Or d.Model = "ES1233S" Then
5240            d.Tamper = 1
5250            d.LastTamper = p.DateTime
5260            SpecialLog "Tamper    " & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now

5270          ElseIf d.IgnoreTamper Then
5280            d.Tamper = 0

5290          Else                   ' d.notamper = 0

5300            d.Tamper = 1
5310            d.LastTamper = p.DateTime
5320            p.Alarmtype = EVT_TAMPER
5330            Set alarm = Troubles.Add(p, d, 0)

5340            If Not alarm Is Nothing Then
5350              alarm.Tamper = 1
5360              alarm.Alarmtype = EVT_TAMPER
5370              alarm.Announce = "Device Tamper"
5380              LogAlarm alarm, EVT_TAMPER, gUser.Username
5390              AddPageRequest alarm, alarm.Alarmtype

5400            End If
5410            Trace "Tamper: " & Right("00000000" & d.Serial, 8) & " " & Now

5420          End If

5430        Case EVT_TAMPER_RESTORE  ' TAMPER BIT RESTORED

5440          Set alarm = Troubles.BySerialTamper(d.Serial)
5450          If Not alarm Is Nothing Then
5460            If alarm.Tamper = 1 Then  ' not sure why tamper needs to be set.

5470              alarm.Tamper = 0
5480              alarm.LastRestore = p.DateTime
5490              alarm.STAT = d.STAT
5500              alarm.packet = p.HexPacket
5510              d.Tamper = 0

5520              Set PageRequest = RemovePageRequest(alarm.AlarmID)  ' Set PageRequest = RemovePageRequest(Alarm.Serial, EVT_TAMPER, 0)

5530              ClearInfoBox alarm.AlarmID
5540              ClearPages alarm.AlarmID
5550              Troubles.RemoveAlarm alarm, EVT_TAMPER
5560              LogAlarm alarm, EVT_TAMPER_RESTORE, gUser.Username
5570              Trace "Tamper Restore: " & Right("00000000" & d.Serial, 8) & " " & Now
5580            End If
5590          Else
5600            If d.Tamper = 1 Then
5610              d.Tamper = 0
5620              SpecialLog "Tamper End" & vbTab & p.Serial & vbTab & p.HexPacket & vbTab & Now
5630            End If
5640          End If

              '********** OTHER TROUBLE *************


5650        Case EVT_GENERAL_TROUBLE  ' UNDEFINED TROUBLE
              'WriteEventToDB EVT
              'Trace "System Trouble: " & Right("0000" & Hex(dserial), 4)

              '********** SILENCED *************

5660        Case EVT_SILENCE         ' SILENCED (BUT NOT ACKNOWLEDGED)

              'WriteEventToDB EVT
              'Trace "Silenced: " & Right("0000" & Hex(d.serial), 4)

              '********** ANNOUNCEMENTS/PAGES *************

5670        Case EVT_ANNOUNCE_1      ' STANDARD ANNOUNCE
              'WriteEventToDB EVT
              'Trace "Announce Level 1: " & Right("0000" & Hex(dserial), 4)

5680        Case EVT_ANNOUNCE_2      ' ESCALATED ANNOUNCE
              'WriteEventToDB EVT
              'Trace "Announce Level 2: " & Right("0000" & Hex(dserial), 4)

5690        Case EVT_ANNOUNCE_3      ' 3RD LEVEL ESCALATED ANNOUNCE
              'WriteEventToDB EVT
              'Trace "Announce Level 3: " & Right("0000" & Hex(dserial), 4)

              '********** CHANGES to DATABASE *************

5700        Case EVT_DATABASE_UPDATE  '
              'WriteEventToDB EVT
              'Trace "Database Update: " & Right("0000" & Hex(dserial), 4)

5710        Case EVT_DATABASE_READ   '
              'Trace "Database Read: " & Right("0000" & Hex(d.UniqueID), 4)

5720        Case EVT_SYSTEM_START, EVT_SYSTEM_STOP
5730          LogGeneric EventType

5740        Case EVT_SYSTEM_LOGIN
5750          ResetActivityTime
5760          LogGeneric EventType
5770          frmMain.TimerLogon.Enabled = True

5780        Case EVT_SYSTEM_LOGOUT

5790          Configuration.PCARedirect = 0
5800          LogGeneric EventType
5810          On Error Resume Next

5820      End Select
5830    Else                         ' IF MASTER
          ' THEN IT MUST BE A REMOTE
          ' SEND ACTIONS TO MASTER

5840      Debug.Print "Post Event to Master " & EventType & " : " & GetEventName(EventType)
5850    End If

PostEvent_Resume:
5860    On Error GoTo 0
5870    Exit Function

PostEvent_Error:

5880    If d Is Nothing Then
5890      If alarm Is Nothing Then
5900        sn = "Alarm?"
5910      Else
5920        sn = alarm.Serial
5930      End If
5940    Else
5950      sn = d.Serial
5960    End If


5970    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.PostEvent." & Erl & " Device:" & sn
5980    Resume PostEvent_Resume


End Function


Function mobileACK(alarm As cAlarm)
  Dim SQL As String
  If (Not (alarm Is Nothing)) Then
  
    SQL = "UPDATE mobile SET TimeAcked = " & q(Now) & ", AckUser = " & q(alarm.Username) & " WHERE ID = " & alarm.AlarmID
    ConnExecute SQL
  
  End If

End Function

Function UpdateAnnounceAndLocationText(alarm As cAlarm)
  Dim SQL As String
  SQL = "UPDATE Alarms Set Announce = " & q(Trim$(alarm.Announce)) & ", userdata = " & q(Trim$(alarm.locationtext)) & " WHERE ID = " & alarm.ID
  ConnExecute SQL
  


End Function


Function LogAlarm(alarm As cAlarm, ByVal EventType As Long, ByVal User As String)
        Dim rs                 As Recordset

        Dim AlarmID As Long

10      On Error GoTo LogAlarm_Error

20      If Not alarm Is Nothing Then


          Dim Fields           As String
          Dim f()              As String
          Dim SQL              As String


30        alarm.EventType = EventType

40        ReDim f(1 To NUM_FLDS)
50        f(FLD_STATUS) = alarm.STAT
60        f(FLD_SERIAL) = q(alarm.Serial)
70        f(FLD_EVENTDATE) = DateDelimit(Now)
80        f(FLD_ALARM) = alarm.alarm  'alarm
90        f(FLD_TAMPER) = alarm.Tamper  ' tamper
100       f(FLD_ISLOCATOR) = alarm.IsLocator  'islocator
110       f(FLD_BATTERY) = alarm.Battery  'battery
120       f(FLD_HOPS) = 0            'hops
130       f(FLD_FIRSTHOP) = q(alarm.FirstHopSerial)  ' firsthop
140       f(FLD_RESIDENTID) = alarm.ResidentID
150       f(FLD_ROOMID) = alarm.RoomID

160       f(FLD_EVENTTYPE) = alarm.EventType
170       If alarm.ID = 0 Then
180         f(FLD_ALARMID) = alarm.PriorID  'prior  alarm.id
190       Else
200         f(FLD_ALARMID) = alarm.ID  ' alarmid
210       End If
220       f(FLD_USERNAME) = q(User)  ' username
230       f(FLD_SESSIONID) = gSessionID
240       f(FLD_ANNOUNCE) = q(left$(Trim(alarm.Announce), 50))
250       f(FLD_PHONE) = q(alarm.Phone)

260       If Len(alarm.Disposition) Then
270         f(FLD_INFO) = q(alarm.Disposition)  ' info
280       Else
290         f(FLD_INFO) = q(left$(alarm.info, 255))  ' info
300       End If

310       f(FLD_USERDATA) = q(alarm.locationtext)  ' userdata

320       f(FLD_SIGNAL) = alarm.LEvel  ' signal
330       f(FLD_MARGIN) = alarm.Margin  ' margin
340       f(FLD_PACKET) = q(alarm.packet)  ' packet
350       f(FLD_FC1) = alarm.FC1     ' fc1
360       f(FLD_FC2) = alarm.FC2     ' fc2
370       f(FLD_IDM) = alarm.IDM     ' idm
380       f(FLD_IDL) = alarm.IDL     ' idl
390       f(FLD_LOCIDM) = alarm.LOCIDM  ' locidm
400       f(FLD_LOCIDL) = alarm.LOCIDL  ' locidl

410       f(FLD_INPUTNUM) = alarm.Inputnum  ' locidl
420       Fields = Join(f, ",")


          '420       SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
           '                & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
           '                & Fields & ")"

430       SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"

440       ConnExecute SQL
          Dim InsertID         As String

450       Set rs = ConnExecute("SELECT @@Identity")
460       InsertID = Val(rs.Fields(0) & "")  ' insert is id of newly inserted row
470       If alarm.ID = 0 Then
480         alarm.ID = InsertID
490       End If
500       alarm.Guid = InsertID

          AlarmID = alarm.ID

510       If (gPush) Then
520         If PushProcessor Is Nothing Then
530           Set PushProcessor = New cPushProcessor
540         End If
550         Select Case EventType
              Case 1 To 7, 14 To 17, 19, 20, 31, 32, 40, 41, 48 To 57, 64, 65
560             PushProcessor.AddByID InsertID  ' insert is id of newly inserted row
570           Case Else
580             PushProcessor.AddByID InsertID  ' insert is id of newly inserted row
590         End Select

600         If Not PushProcessor.Busy Then
610           PushProcessor.Send
620         End If

630       End If
640       Debug.Print "Insert ID " & InsertID

650       rs.Close
660       Set rs = Nothing



          '' for mobile phone devices
670       Select Case EventType
            Case EVT_ALERT_END, EVT_EXTERN_END, EVT_EMERGENCY_END, EVT_ASSISTANCE_END
680           SQL = "update mobile set Ended = 1 WHERE AlarmID = " & alarm.ID
690           ConnExecute SQL

700         Case EVT_ASSISTANCE_RESPOND
710           SQL = "update mobile set eacktime = " & q(Now()) & ", Ackuser = " & q(User) & " WHERE AlarmID = " & alarm.ID
720           ConnExecute SQL
730           For Each alarm In alarms.alarms
740             If alarm.ID = AlarmID Then
750               If alarm.Responder = "" Then
760                 alarm.Responder = User
770                 frmMain.ShowResponder alarm
780                 Exit For
790               End If
800             End If
810           Next



820         Case EVT_EMERGENCY_RESPOND
830           SQL = "update mobile set eacktime = " & q(Now()) & ", Ackuser = " & q(User) & " WHERE AlarmID = " & alarm.ID
840           ConnExecute SQL
850           For Each alarm In alarms.alarms
                If alarm.ID = AlarmID Then
860             If alarm.Responder = "" Then
870               alarm.Responder = User
880               frmMain.ShowResponder alarm
                  Exit For
890             End If
                End If
900           Next



910         Case EVT_ALERT_RESPOND
920           SQL = "update mobile set eacktime = " & q(Now()) & ", Ackuser = " & q(User) & " WHERE AlarmID = " & alarm.ID
930           ConnExecute SQL
940           For Each alarm In Alerts.alarms
950             If alarm.ID = AlarmID Then
960               If alarm.Responder = "" Then
970                 alarm.Responder = User
980                 frmMain.ShowResponder alarm
990                 Exit For
1000              End If
1010            End If
1020          Next

1030        Case EVT_EXTERN_RESPOND

1040          For Each alarm In Externs.alarms
1050            If alarm.ID = AlarmID Then
1060              If alarm.Responder = "" Then
1070                alarm.Responder = User
1080                frmMain.ShowResponder alarm
1090                Exit For
1100              End If
1110            End If
1120          Next

1130        Case EVT_GENERIC_RESPOND
1140          SQL = "update mobile set eacktime = " & q(Now()) & ", Ackuser = " & q(User) & " WHERE AlarmID = " & alarm.ID
1150          ConnExecute SQL
              'frmMain.ShowACK alarm
1160      End Select

1170      Exit Function

1180    End If

LogAlarm_Resume:
1190    On Error GoTo 0
1200    Exit Function

LogAlarm_Error:

1210    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogAlarm." & Erl
1220    Resume LogAlarm_Resume


End Function



Public Function GetEventTypeName(ByVal EventType As Integer) As String
  If EventTypes Is Nothing Then
    CreateEventtypes
  End If
  If EventType >= 0 And EventType < EventTypes.Count Then
    GetEventTypeName = EventTypes(EventType + 1).ToString
  End If
End Function

Function LogGeneric(ByVal EventType As Long) As Long
        Dim rs                 As Recordset

10      On Error GoTo LogGeneric_Error

        Dim Fields             As String
        Dim f()                As String
        Dim SQL                As String

20      ReDim f(1 To NUM_FLDS)
30      f(FLD_STATUS) = 0
40      f(FLD_SERIAL) = q("")
50      f(FLD_EVENTDATE) = DateDelimit(Now)
60      f(FLD_ALARM) = 0                     'alarm
70      f(FLD_TAMPER) = 0                     ' tamper
80      f(FLD_ISLOCATOR) = 0                     'islocator
90      f(FLD_BATTERY) = 0                     'battery
100     f(FLD_HOPS) = 0                     'hops
110     f(FLD_FIRSTHOP) = q("")                 ' firsthop
120     f(FLD_RESIDENTID) = 0
130     f(FLD_ROOMID) = 0
140     f(FLD_EVENTTYPE) = EventType
150     f(FLD_ALARMID) = 0                    ' alarmid
160     f(FLD_USERNAME) = q(gUser.Username)    ' username
170     f(FLD_SESSIONID) = gSessionID
180     f(FLD_ANNOUNCE) = q("")
190     f(FLD_PHONE) = q("")
200     f(FLD_INFO) = q("")                ' info
210     f(FLD_USERDATA) = q("")                ' userdata
220     f(FLD_SIGNAL) = 0                    'signal
230     f(FLD_MARGIN) = 0                    ' margin
240     f(FLD_PACKET) = q("")                ' packet
250     f(FLD_FC1) = 0                    'fc1
260     f(FLD_FC2) = 0                    'fc2
270     f(FLD_IDM) = 0                    'idm
280     f(FLD_IDL) = 0                    'idl
290     f(FLD_LOCIDM) = 0                    'locidm
300     f(FLD_LOCIDL) = 0                    'locidl
305     f(FLD_INPUTNUM) = 0                    'locidl

310     Fields = Join(f, ",")


'320     SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'              & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl.inputnum) values (" _
'              & Fields & ")"
              
              
          SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"
              

330     ConnExecute SQL


        ''Remove after testing
      '
      '  Dim InsertID           As String
      '
      '
      '  Set Rs = ConnExecute("SELECT @@Identity")
      '  InsertID = Val(Rs.Fields(0) & "")  ' insert is id of newly inserted row
      '  '380       If alarm.ID = 0 Then
      '  '390         alarm.ID = InsertID
      '  '400       End If
      '  '410       alarm.Guid = InsertID
      '
      '  If (gPush) Then
      '    If PushProcessor Is Nothing Then
      '      Set PushProcessor = New cPushProcessor
      '    End If
      '    Select Case EventType
      '
      '      Case 1 To 7, 14 To 17, 19, 20, 31, 32, 40, 41, 48 To 57, 64, 65
      '        PushProcessor.AddByID InsertID  ' insert is id of newly inserted row
      '      Case Else
      '        PushProcessor.AddByID InsertID  ' insert is id of newly inserted row
      '    End Select
      '
      '    If Not PushProcessor.Busy Then
      '      PushProcessor.Send
      '    End If
      '
      '
      '  End If






        '
        '
        '340     Set rs = New ADODB.Recordset
        '350     rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
        '
        '360     rs.addnew
        '370     rs("Status") = 0
        '380     rs("Serial") = 0
        '390     rs("EventDate") = Now
        '400     rs("Alarm") = 0
        '410     rs("Tamper") = 0
        '420     rs("IsLocator") = 0
        '430     rs("Battery") = 0
        '440     rs("Hops") = 0
        '450     rs("FirstHop") = ""
        '460     rs("ResidentID") = 0
        '470     rs("RoomID") = 0
        '480     rs("EventType") = EventType
        '490     rs("AlarmID") = 0
        '500     rs("username") = gUser.UserName
        '510     rs("sessionid") = gSessionID
        '520     rs("announce") = ""
        '530     rs("Phone") = ""
        '540     rs("Info") = ""
        '
        '550     rs("Userdata") = ""
        '560     rs("Signal") = 0
        '570     rs("Margin") = 0
        '580     rs("Packet") = ""
        '
        '
        '
        '590     rs("FC1") = 0
        '600     rs("FC2") = 0
        '610     rs("IDM") = 0
        '620     rs("IDL") = 0
        '630     rs("LOCIDM") = 0
        '640     rs("LOCIDL") = 0
        '
        '650     rs.Update
        '660     rs.Close
        '670     Set rs = Nothing

LogGeneric_Resume:
340     On Error GoTo 0
350     Exit Function

LogGeneric_Error:

360     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogGeneric." & Erl
370     Resume LogGeneric_Resume


End Function


Function LogAssur(ByVal EventType As Long) As Long
        'Dim rs                 As Recordset

10      On Error GoTo LogAssur_Error

        Dim Fields             As String
        Dim f()                As String
        Dim SQL                As String

20      ReDim f(1 To NUM_FLDS)
30      f(FLD_STATUS) = 0
40      f(FLD_SERIAL) = q("")
50      f(FLD_EVENTDATE) = DateDelimit(Now)
60      f(FLD_ALARM) = 0                     'alarm
70      f(FLD_TAMPER) = 0                     ' tamper
80      f(FLD_ISLOCATOR) = 0                     'islocator
90      f(FLD_BATTERY) = 0                     'battery
100     f(FLD_HOPS) = 0                     'hops
110     f(FLD_FIRSTHOP) = q("")                 ' firsthop
120     f(FLD_RESIDENTID) = 0
130     f(FLD_ROOMID) = 0
140     f(FLD_EVENTTYPE) = EventType
150     f(FLD_ALARMID) = 0                    ' alarmid
160     f(FLD_USERNAME) = q(gUser.Username)    ' username
170     f(FLD_SESSIONID) = gSessionID
180     f(FLD_ANNOUNCE) = q("")
190     f(FLD_PHONE) = q("")
200     f(FLD_INFO) = q("")                ' info
210     f(FLD_USERDATA) = q("")                ' userdata
220     f(FLD_SIGNAL) = 0                    'signal
230     f(FLD_MARGIN) = 0                    ' margin
240     f(FLD_PACKET) = q("")                ' packet
250     f(FLD_FC1) = 0                    'fc1
260     f(FLD_FC2) = 0                    'fc2
270     f(FLD_IDM) = 0                    'idm
280     f(FLD_IDL) = 0                    'idl
290     f(FLD_LOCIDM) = 0                    'locidm
300     f(FLD_LOCIDL) = 0                    'locidl
305     f(FLD_INPUTNUM) = 0                    'locidl

310     Fields = Join(f, ",")


'320     SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'              & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'              & Fields & ")"
              
              
       SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"
              

330     ConnExecute SQL


        '
        '20      Set rs = New ADODB.Recordset
        '30      rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
        '
        '
        '
        '
        '
        '
        '
        '40      rs.addnew
        '50      rs("Status") = 0
        '60      rs("Serial") = 0
        '70      rs("EventDate") = Now
        '80      rs("Alarm") = 0
        '90      rs("Tamper") = 0
        '100     rs("IsLocator") = 0
        '110     rs("Battery") = 0
        '120     rs("Hops") = 0
        '130     rs("FirstHop") = ""
        '
        '140     rs("ResidentID") = 0
        '150     rs("RoomID") = 0
        '160     rs("EventType") = EventType
        '170     rs("AlarmID") = 0
        '180     rs("username") = gUser.UserName
        '190     rs("sessionid") = gSessionID
        '200     rs("announce") = ""
        '210     rs("Phone") = ""
        '220     rs("Info") = ""
        '
        '230     rs("Userdata") = ""
        '240     rs("Signal") = 0
        '250     rs("Margin") = 0
        '260     rs("Packet") = ""
        '
        '270     rs("FC1") = 0
        '280     rs("FC2") = 0
        '290     rs("IDM") = 0
        '300     rs("IDL") = 0
        '310     rs("LOCIDM") = 0
        '320     rs("LOCIDL") = 0
        '
        '330     rs.Update
        '340     rs.Close
        '350     Set rs = Nothing

LogAssur_Resume:
340     On Error GoTo 0
350     Exit Function

LogAssur_Error:

360     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogAssur." & Erl
370     Resume LogAssur_Resume


End Function

Function LogAssurFail(d As cESDevice, Optional ByVal EventType As Long = EVT_ASSUR_FAIL) As Long

        Dim rs                 As Recordset
10      On Error GoTo LogAssurFail_Error

20      If d Is Nothing Then
30        Set d = New cESDevice
40      End If


        Dim Fields             As String
        Dim f()                As String
        Dim SQL                As String

50      ReDim f(1 To NUM_FLDS)
60      f(FLD_STATUS) = 0
70      f(FLD_SERIAL) = q(d.Serial)
80      f(FLD_EVENTDATE) = DateDelimit(Now)
90      f(FLD_ALARM) = 0                     'alarm
100     f(FLD_TAMPER) = 0                     ' tamper
110     f(FLD_ISLOCATOR) = 0                     'islocator
120     f(FLD_BATTERY) = 0                     'battery
130     f(FLD_HOPS) = 0                     'hops
140     f(FLD_FIRSTHOP) = q("")                 ' firsthop
150     f(FLD_RESIDENTID) = d.ResidentID
160     f(FLD_ROOMID) = d.RoomID
170     f(FLD_EVENTTYPE) = EventType
180     f(FLD_ALARMID) = 0                    ' alarmid
190     f(FLD_USERNAME) = q(gUser.Username)    ' username
200     f(FLD_SESSIONID) = gSessionID
210     f(FLD_ANNOUNCE) = q(d.Announce)
220     f(FLD_PHONE) = q(d.Phone)
230     f(FLD_INFO) = q(left$(d.Notes, 255))  ' info
240     f(FLD_USERDATA) = q("")                ' userdata
250     f(FLD_SIGNAL) = 0                    'signal
260     f(FLD_MARGIN) = 0                    ' margin
270     f(FLD_PACKET) = q("")                ' packet
280     f(FLD_FC1) = 0                    'fc1
290     f(FLD_FC2) = 0                    'fc2
300     f(FLD_IDM) = 0                    'idm
310     f(FLD_IDL) = 0                    'idl
320     f(FLD_LOCIDM) = 0                    'locidm
330     f(FLD_LOCIDL) = 0                    'locidl
335     f(FLD_INPUTNUM) = 0                    'locidl

340     Fields = Join(f, ",")


'350     SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'              & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'              & Fields & ")"

        SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"

360     ConnExecute SQL

        '
        '
        '
        '
        '
        '
        '50      Set rs = New ADODB.Recordset
        '60      rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
        '
        '70      rs.addnew
        '80      rs("Status") = 0
        '90      rs("Serial") = d.Serial
        '100     rs("EventDate") = Now
        '110     rs("Alarm") = 0
        '120     rs("Tamper") = 0
        '130     rs("IsLocator") = d.IsLocator
        '140     rs("Battery") = 0
        '150     rs("Hops") = 0
        '160     rs("FirstHop") = ""
        '
        '170     rs("ResidentID") = d.ResidentID
        '180     rs("RoomID") = d.RoomID
        '190     rs("EventType") = EventType
        '200     rs("AlarmID") = 0
        '210     rs("username") = gUser.UserName
        '220     rs("sessionid") = gSessionID
        '230     rs("announce") = ""
        '240     rs("Phone") = d.Phone
        '250     rs("Info") = left$(d.Notes, 255)
        '
        '
        '260     rs("Userdata") = ""
        '270     rs("Signal") = 0
        '280     rs("Margin") = 0
        '290     rs("Packet") = ""
        '
        '300     rs("FC1") = 0
        '310     rs("FC2") = 0
        '320     rs("IDM") = 0
        '330     rs("IDL") = 0
        '340     rs("LOCIDM") = 0
        '350     rs("LOCIDL") = 0
        '
        '360     rs.Update
        '370     rs.Close
        '380     Set rs = Nothing


LogAssurFail_Resume:
370     On Error GoTo 0
380     Exit Function

LogAssurFail_Error:

390     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogAssurFail." & Erl
400     Resume LogAssurFail_Resume


End Function


Function LogAssurCheckin(d As cESDevice, p As cESPacket, Optional ByVal EventType As Long = EVT_ASSUR_CHECKIN) As Long

        '       Dim rs                 As Recordset
10      On Error GoTo LogAssurCheckin_Error

20      If d Is Nothing Then
30        Set d = New cESDevice
40      End If

        Dim Fields             As String
        Dim f()                As String
        Dim SQL                As String

50      ReDim f(1 To NUM_FLDS)
60      f(FLD_STATUS) = 0
70      f(FLD_SERIAL) = q(d.Serial)
80      f(FLD_EVENTDATE) = DateDelimit(Now)
90      f(FLD_ALARM) = 0                     'alarm
100     f(FLD_TAMPER) = 0                     ' tamper
110     f(FLD_ISLOCATOR) = 0                     'islocator
120     f(FLD_BATTERY) = 0                     'battery
130     f(FLD_HOPS) = 0                     'hops
140     f(FLD_FIRSTHOP) = q("")                 ' firsthop
150     f(FLD_RESIDENTID) = d.ResidentID
160     f(FLD_ROOMID) = d.RoomID
170     f(FLD_EVENTTYPE) = EventType
180     f(FLD_ALARMID) = 0                    ' alarmid
190     f(FLD_USERNAME) = q(gUser.Username)    ' username
200     f(FLD_SESSIONID) = gSessionID
210     f(FLD_ANNOUNCE) = q(d.Announce)
220     f(FLD_PHONE) = q(d.Phone)
230     f(FLD_INFO) = q(left$(d.Notes, 255))  ' info
240     f(FLD_USERDATA) = q("")                ' userdata
250     f(FLD_SIGNAL) = 0                    'signal
260     f(FLD_MARGIN) = 0                    ' margin
270     f(FLD_PACKET) = q("")                ' packet
280     f(FLD_FC1) = 0                    'fc1
290     f(FLD_FC2) = 0                    'fc2
300     f(FLD_IDM) = 0                    'idm
310     f(FLD_IDL) = 0                    'idl
320     f(FLD_LOCIDM) = 0                    'locidm
330     f(FLD_LOCIDL) = 0                    'locidl
335     f(FLD_INPUTNUM) = 0                    'locidl

340     Fields = Join(f, ",")


'350     SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'              & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'              & Fields & ")"

        SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"

360     ConnExecute SQL
        Debug.Print "Log Assur Checkin " & d.Serial
370     Exit Function


        '50      Set rs = New ADODB.Recordset
        '60      rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
        '
        '70      rs.addnew
        '80      rs("Status") = 0
        '90      rs("Serial") = d.Serial
        '100     rs("EventDate") = Now
        '110     rs("Alarm") = 0
        '120     rs("Tamper") = 0             ' p.Tamper
        '130     rs("IsLocator") = d.IsLocator
        '140     rs("Battery") = 0            'p.Battery
        '150     rs("Hops") = 0
        '160     rs("FirstHop") = ""
        '170     rs("ResidentID") = d.ResidentID
        '180     rs("RoomID") = d.RoomID
        '190     rs("EventType") = EventType
        '200     rs("AlarmID") = 0
        '210     rs("username") = gUser.UserName
        '220     rs("sessionid") = gSessionID
        '230     rs("announce") = d.Announce
        '240     rs("Phone") = d.Phone
        '250     rs("Info") = left$(d.Notes, 255)
        '
        '260     rs("Userdata") = ""
        '270     rs("Signal") = 0
        '280     rs("Margin") = 0
        '290     rs("Packet") = ""
        '
        '300     rs("FC1") = 0
        '310     rs("FC2") = 0
        '320     rs("IDM") = 0
        '330     rs("IDL") = 0
        '340     rs("LOCIDM") = 0
        '350     rs("LOCIDL") = 0
        '
        '360     rs.Update
        '370     rs.Close
        '


LogAssurCheckin_Resume:
380     On Error GoTo 0
        '380     Set rs = Nothing
390     Exit Function

LogAssurCheckin_Error:

400     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogAssurCheckin." & Erl
410     Resume LogAssurCheckin_Resume


End Function



Function LogCheckInEvent(d As cESDevice, p As cESPacket, ByVal EventType As Long)

        'Dim rs                 As Recordset

10      On Error GoTo LogCheckInEvent_Error

20      If Not d Is Nothing Then

          Dim Fields           As String
          Dim f()              As String
          Dim SQL              As String

30        ReDim f(1 To NUM_FLDS)
40        f(FLD_STATUS) = d.STAT
50        f(FLD_SERIAL) = q(d.Serial)
60        f(FLD_EVENTDATE) = DateDelimit(Now)
70        f(FLD_ALARM) = p.alarm             'alarm
80        f(FLD_TAMPER) = p.Tamper            ' tamper
90        f(FLD_ISLOCATOR) = d.IsLocator         'islocator
100       f(FLD_BATTERY) = p.Battery           'battery
110       f(FLD_HOPS) = 0                   'hops
120       f(FLD_FIRSTHOP) = q("")               ' firsthop
130       f(FLD_RESIDENTID) = d.ResidentID
140       f(FLD_ROOMID) = d.RoomID
150       f(FLD_EVENTTYPE) = EventType
160       f(FLD_ALARMID) = 0                  ' alarmid
170       f(FLD_USERNAME) = q(gUser.Username)  ' username
180       f(FLD_SESSIONID) = gSessionID
190       f(FLD_ANNOUNCE) = q(d.Announce)
200       f(FLD_PHONE) = q(d.Phone)
210       f(FLD_INFO) = q(left$(d.Notes, 255))  ' info
220       f(FLD_USERDATA) = q("")              ' userdata
230       f(FLD_SIGNAL) = 0                  'signal
240       f(FLD_MARGIN) = 0                  ' margin
250       f(FLD_PACKET) = q("")              ' packet
260       f(FLD_FC1) = 0                  'fc1
270       f(FLD_FC2) = 0                  'fc2
280       f(FLD_IDM) = d.IDM              'idm
290       f(FLD_IDL) = d.IDL              'idl
300       f(FLD_LOCIDM) = 0                  'locidm
310       f(FLD_LOCIDL) = 0                  'locidl
315       f(FLD_INPUTNUM) = 0                  'locidl

320       Fields = Join(f, ",")


'330       SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'                & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'                & Fields & ")"

330       SQL = "insert into alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"


340       ConnExecute SQL
          '
          '    Exit Function
          '
          '
          '
          'Set rs = New ADODB.Recordset
          'rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
          '
          'rs.addnew
          'rs("Status") = d.STAT
          'rs("Serial") = d.Serial
          'rs("EventDate") = Now
          'rs("Alarm") = p.Alarm
          'rs("Tamper") = p.Tamper
          'rs("IsLocator") = d.IsLocator
          'rs("Battery") = p.Battery
          '
          'rs("Hops") = 0
          'rs("FirstHop") = ""
          '
          'rs("ResidentID") = d.ResidentID
          'rs("RoomID") = d.RoomID
          'rs("EventType") = EventType
          'rs("Announce") = d.Announce
          'rs("AlarmID") = 0          ' not an alarm
          'rs("username") = gUser.UserName
          'rs("sessionid") = gSessionID
          'rs("Phone") = d.Phone
          'rs("Info") = d.Notes
          '
          '
          'rs("Userdata") = ""
          '
          'rs("FC1") = 0              'D.FC1
          'rs("FC2") = 0              'D.FC2
          'rs("IDM") = d.IDM
          'rs("IDL") = d.IDL
          'rs("LOCIDM") = 0
          'rs("LOCIDL") = 0
          '
          '
          'rs.Update
          'rs.Close
          'Set rs = Nothing
350     End If

LogCheckInEvent_Resume:
360     On Error GoTo 0

370     Exit Function

LogCheckInEvent_Error:

380     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogCheckInEvent." & Erl
390     Resume LogCheckInEvent_Resume


End Function

Function LogToStrays(p As cESPacket)
  Dim hfile As Integer
  Dim filename As String
  filename = App.Path & "\Strays.log"
  
  limitFileSize filename
  
  hfile = FreeFile
  Open filename For Append As #hfile
    Print #hfile, p.Serial, p.HexPacket, Now
  Close hfile
End Function

Function LogCommEvent(d As cESDevice, p As cESPacket, ByVal EventType As Long)

        'Dim rs                 As Recordset

        'P is not used.
10      On Error GoTo LogCommEvent_Error

20      If d Is Nothing Then
30        Set d = New cESDevice
40      End If


        Dim Fields             As String
        Dim f()                As String
        Dim SQL                As String

50      ReDim f(1 To NUM_FLDS)
60      f(FLD_STATUS) = d.STAT
70      f(FLD_SERIAL) = q(d.Serial)
80      f(FLD_EVENTDATE) = DateDelimit(Now)
90      f(FLD_ALARM) = p.alarm               'alarm
100     f(FLD_TAMPER) = p.Tamper              ' tamper
110     f(FLD_ISLOCATOR) = d.IsLocator           'islocator
120     f(FLD_BATTERY) = p.Battery             'battery
130     f(FLD_HOPS) = 0                     'hops
140     f(FLD_FIRSTHOP) = q("")                 ' firsthop
150     f(FLD_RESIDENTID) = d.ResidentID
160     f(FLD_ROOMID) = d.RoomID
170     f(FLD_EVENTTYPE) = EventType
180     f(FLD_ALARMID) = 0                    ' alarmid
190     f(FLD_USERNAME) = q(gUser.Username)    ' username
200     f(FLD_SESSIONID) = gSessionID
210     f(FLD_ANNOUNCE) = q(d.Announce)
220     f(FLD_PHONE) = q(d.Phone)
230     f(FLD_INFO) = q(left$(d.Notes, 255))  ' info
240     f(FLD_USERDATA) = q("")                ' userdata
250     f(FLD_SIGNAL) = 0                    'signal
260     f(FLD_MARGIN) = 0                    ' margin
270     f(FLD_PACKET) = q("")                ' packet
280     f(FLD_FC1) = 0                    'fc1
290     f(FLD_FC2) = 0                    'fc2
300     f(FLD_IDM) = d.IDM                'idm
310     f(FLD_IDL) = d.IDL                'idl
320     f(FLD_LOCIDM) = 0                    'locidm
330     f(FLD_LOCIDL) = 0                    'locidl

335     f(FLD_INPUTNUM) = 0                    'locidl

340     Fields = Join(f, ",")


'350     SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'              & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'              & Fields & ")"
              
        SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"
              

360     ConnExecute SQL




        '  Set rs = New ADODB.Recordset
        '  rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
        '
        '  rs.addnew
        '  rs("FC1") = 0                'D.FC1
        '  rs("FC2") = 0                'D.FC2
        '  rs("IDM") = d.IDM
        '  rs("IDL") = d.IDL
        '  rs("Status") = d.STAT
        '  rs("Serial") = d.Serial
        '  rs("EventDate") = Now
        '  rs("Alarm") = d.Alarm
        '  rs("Tamper") = d.Tamper
        '  rs("IsLocator") = d.IsLocator
        '  rs("Battery") = d.Battery
        '
        '  rs("Hops") = 0
        '  rs("FirstHop") = ""
        '
        '  rs("LOCIDM") = 0
        '  rs("LOCIDL") = 0
        '  rs("ResidentID") = d.ResidentID
        '  rs("RoomID") = d.RoomID
        '  rs("EventType") = EventType
        '  rs("Announce") = d.Announce
        '  rs("AlarmID") = 0            ' not an alarm
        '  rs("username") = gUser.UserName
        '  rs("sessionid") = gSessionID
        '  rs("Phone") = ""
        '  rs("Info") = ""
        '
        '  rs("Userdata") = ""
        '  rs("Signal") = 0
        '  rs("Margin") = 0
        '  rs("Packet") = ""
        '
        '  rs("FC1") = 0
        '  rs("FC2") = 0
        '  rs("IDM") = 0
        '  rs("IDL") = 0
        '  rs("LOCIDM") = 0
        '
        '
        '  rs.Update
        '  rs.Close
        'Set rs = Nothing


LogCommEvent_Resume:
370     On Error GoTo 0
380     Exit Function

LogCommEvent_Error:

390     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogCommEvent." & Erl
400     Resume LogCommEvent_Resume


End Function

Function LogAddRes(ByVal ResidentID As Long) As Long
        Dim rs                 As Recordset
        Dim rsDest             As Recordset
        

        Dim Fields           As String
        Dim f()              As String
        Dim SQL                As String


10      On Error GoTo LogAddRes_Error

20      SQL = "SELECT * FROM residents WHERE residentid = " & ResidentID
30      Set rs = ConnExecute(SQL)
40      If Not rs.EOF Then



50        ReDim f(1 To NUM_FLDS)
60        f(FLD_STATUS) = 0
70        f(FLD_SERIAL) = q("")
80        f(FLD_EVENTDATE) = DateDelimit(Now)
90        f(FLD_ALARM) = 0                   'alarm
100       f(FLD_TAMPER) = 0                   ' tamper
110       f(FLD_ISLOCATOR) = 0                   'islocator
120       f(FLD_BATTERY) = 0                   'battery
130       f(FLD_HOPS) = 0                   'hops
140       f(FLD_FIRSTHOP) = q("")               ' firsthop
150       f(FLD_RESIDENTID) = rs("ResidentID")
160       f(FLD_ROOMID) = rs("RoomID")
170       f(FLD_EVENTTYPE) = EVT_ADD_RES
180       f(FLD_ALARMID) = 0                  ' alarmid
190       f(FLD_USERNAME) = q(gUser.Username)  ' username
200       f(FLD_SESSIONID) = gSessionID
210       f(FLD_ANNOUNCE) = q("")
220       f(FLD_PHONE) = q(rs("Phone"))
230       f(FLD_INFO) = q(rs("info"))      ' info
240       f(FLD_USERDATA) = q("")              ' userdata
250       f(FLD_SIGNAL) = 0                  'signal
260       f(FLD_MARGIN) = 0                  ' margin
270       f(FLD_PACKET) = q("")              ' packet
280       f(FLD_FC1) = 0                  'fc1
290       f(FLD_FC2) = 0                  'fc2
300       f(FLD_IDM) = 0             'idm
310       f(FLD_IDL) = 0             'idl
320       f(FLD_LOCIDM) = 0                  'locidm
330       f(FLD_LOCIDL) = 0                  'locidl

335       f(FLD_INPUTNUM) = 0                  'locidl
340       Fields = Join(f, ",")

'350       SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'                & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'                & Fields & ")"
'
                
          SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"
                

360       ConnExecute SQL


370     End If
380     rs.Close
390     Set rs = Nothing

LogAddRes_Resume:
400     On Error GoTo 0
410     Set rs = Nothing
420     Exit Function

LogAddRes_Error:

430     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogAddRes." & Erl
440     Resume LogAddRes_Resume


End Function


Function LogVacation(ByVal ResidentID As Long, ByVal RoomID As Long, ByVal Away As Integer, ByVal User As String) As Long

        
        Dim EventType          As Long

10      On Error GoTo LogVacation_Error


        Dim Fields             As String
        Dim f()                As String
        Dim SQL                As String

20      If Away <> 0 Then
30        EventType = EVT_VACATION
40      Else
50        EventType = EVT_VACATION_RETURN
60      End If


70      ReDim f(1 To NUM_FLDS)
80      f(FLD_STATUS) = 0
90      f(FLD_SERIAL) = q("")
100     f(FLD_EVENTDATE) = DateDelimit(Now)
110     f(FLD_ALARM) = 0                     'alarm
120     f(FLD_TAMPER) = 0                     ' tamper
130     f(FLD_ISLOCATOR) = 0                     'islocator
140     f(FLD_BATTERY) = 0                     'battery
150     f(FLD_HOPS) = 0                     'hops
160     f(FLD_FIRSTHOP) = q("")                 ' firsthop
170     f(FLD_RESIDENTID) = ResidentID
180     f(FLD_ROOMID) = RoomID
190     f(FLD_EVENTTYPE) = EventType
200     f(FLD_ALARMID) = 0                    ' alarmid
210     f(FLD_USERNAME) = q(User)              ' username
220     f(FLD_SESSIONID) = gSessionID
230     f(FLD_ANNOUNCE) = q("")
240     f(FLD_PHONE) = q("")
250     f(FLD_INFO) = q("")                ' info
260     f(FLD_USERDATA) = q("")                ' userdata
270     f(FLD_SIGNAL) = 0                    'signal
280     f(FLD_MARGIN) = 0                    ' margin
290     f(FLD_PACKET) = q("")                ' packet
300     f(FLD_FC1) = 0                    'fc1
310     f(FLD_FC2) = 0                    'fc2
320     f(FLD_IDM) = 0                    'idm
330     f(FLD_IDL) = 0                    'idl
340     f(FLD_LOCIDM) = 0                    'locidm
350     f(FLD_LOCIDL) = 0                    'locidl
355     f(FLD_INPUTNUM) = 0                    'locidl

360     Fields = Join(f, ",")


'370     SQL = "insert into alarms (Status,Serial,Eventdate,alarm,tamper,islocator,battery,hops,firsthop,residentid,roomid,eventtype,alarmid," _
'              & "username, sessionid,announce,phone,info,userdata,signal,margin,packet,fc1,fc2,idm,idl,locidm,locidl,inputnum) values (" _
'              & Fields & ")"

        SQL = "insert into Alarms (" & FIELDNAMES_CSV & ") values (" & Fields & ")"

380     ConnExecute SQL


LogVacation_Resume:
390     On Error GoTo 0
400     Exit Function

LogVacation_Error:

410     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.LogVacation." & Erl
420     Resume LogVacation_Resume


End Function

Public Function DeleteResident(ByVal ResidentID As Long, ByVal User As String) As Long

        Dim rs        As Recordset
        Dim Device    As cESDevice
        Dim Resident  As cResident
        Dim RoomID    As Long
        Dim j         As Integer
        Dim SQL       As String

10      On Error GoTo DeleteResident_Error



50        Connect

60        Set Resident = New cResident
70        Resident.Fetch ResidentID

80        Resident.GetTransmitters
90        RoomID = GetRoomIDFromResID(ResidentID)

100       'conn.BeginTrans

110       Set rs = New ADODB.Recordset
120       rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType
          ' log event that resident is being removed

130       rs.addnew
140       rs("FC1") = 0
150       rs("FC2") = 0
160       rs("IDM") = 0
170       rs("IDL") = 0
180       rs("Status") = 0
190       rs("Serial") = 0
200       rs("EventDate") = Now
210       rs("Alarm") = 0
220       rs("Tamper") = 0
230       rs("IsLocator") = 0
240       rs("Battery") = 0
250       rs("LOCIDM") = 0
260       rs("LOCIDL") = 0
270       rs("ResidentID") = ResidentID
280       rs("RoomID") = RoomID
290       rs("EventType") = EVT_REMOVE_RES
300       rs("AlarmID") = 0  ' not an alarm
310       rs("username") = User
320       rs("sessionid") = gSessionID
330       rs("announce") = left$(Resident.NameLast & ", " & Resident.NameFirst, 50)
340       rs("Phone") = Resident.Phone
350       rs("Info") = left$(Resident.info, 255)
360       rs.Update
370       rs.Close

          ' actually remove resident
380       SQL = " Update Residents set Deleted = 1, RoomID = 0  WHERE ResidentID = " & ResidentID
390       ConnExecute SQL

          'log any devices being unassigned
400       rs.Open "SELECT * FROM alarms WHERE 1 = 2 ", conn, gCursorType, gLockType

410       For j = 1 To Resident.AssignedTx.Count
420         Set Device = Resident.AssignedTx(j)
            Device.ResidentID = 0 'unassign device in memory (fix april 12 2016)
430         rs.addnew
440         rs("FC1") = 0  'device.FC1
450         rs("FC2") = 0  'device.FC2
460         rs("IDM") = 0  'device.IDM
470         rs("IDL") = 0  'device.IDL
480         rs("Status") = 0
490         rs("Serial") = Device.Serial
500         rs("EventDate") = Now
510         rs("Alarm") = 0
520         rs("Tamper") = 0
530         rs("IsLocator") = 0
540         rs("Battery") = 0
550         rs("LOCIDM") = 0
560         rs("LOCIDL") = 0
570         rs("ResidentID") = ResidentID
580         rs("RoomID") = 0
590         rs("EventType") = EVT_UNASSIGN_DEV
600         rs("AlarmID") = 0  ' not an alarm
610         rs("username") = User
620         rs("sessionid") = gSessionID
630         rs("announce") = Device.Announce
640         rs("Phone") = Resident.Phone
650         rs("Info") = left$(Resident.info, 255)
660         rs.Update
670       Next
680       rs.Close

          'Actually unassign device
          'SQl = "UPDATE Devices Set ResidentID = 0, ROOMID = 0 WHERE ResidentID = " & ResidentID
          Dim Dev As cESDevice
          For Each Dev In Devices.Devices
            If Dev.ResidentID = ResidentID Then
              Dev.ResidentID = 0
            End If
          Next
          
          

700       SQL = "UPDATE Devices Set ResidentID = 0 WHERE ResidentID = " & ResidentID
710       ConnExecute SQL

720       If conn.Errors.Count <> 0 Then
730         'conn.RollbackTrans
740         DeleteResident = 0
750       Else
760         'conn.CommitTrans
770         DeleteResident = 1  ' success!
780       End If
790       Set rs = Nothing



DeleteResident_Resume:
810     On Error GoTo 0
820     Exit Function

DeleteResident_Error:

830     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.DeleteResident." & Erl
840     Resume DeleteResident_Resume


End Function

Public Function DeleteStaff(ByVal ResidentID As Long, ByVal User As String) As Long

        Dim rs        As Recordset
        Dim Device    As cESDevice
        Dim Resident  As cResident
        Dim RoomID    As Long
        Dim j         As Integer
        Dim SQL       As String

10      On Error GoTo DeleteStaff_Error

50        Connect

60        'Set Resident = New cResident
70        'Resident.Fetch ResidentID


100       'conn.BeginTrans

          ' actually remove Staff
'380       SQL = " Update Staff set Deleted = 1  WHERE staffID = " & ResidentID
380       SQL = "Delete from Staff WHERE staffID = " & ResidentID
390       ConnExecute SQL



720       If conn.Errors.Count <> 0 Then
730         'conn.RollbackTrans
740         DeleteStaff = 0
750       Else
760         'conn.CommitTrans
770         DeleteStaff = 1  ' success!
780       End If
790       Set rs = Nothing



DeleteStaff_Resume:
810     On Error GoTo 0
820     Exit Function

DeleteStaff_Error:

830     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modEvents.DeleteStaff." & Erl
840     Resume DeleteStaff_Resume


End Function



Public Function DeleteTransmitter(ByVal ID As Long, ByVal Username As String) As Long
  Dim SQL As String
  Dim rs As Recordset
  Dim Device As cESDevice
  Dim IDL As Integer
  Dim IDM As Integer
  Dim Serial As String
  Dim ZoneID As Long

  On Error Resume Next

  SQL = "select * FROM Devices WHERE DeviceID = " & ID
  Set rs = ConnExecute(SQL)
  If Not rs.EOF Then
    Serial = rs("serial") & ""
  End If
  rs.Close

  If Len(Serial) > 0 Then
    
    Set Device = Devices.Device(Serial)
    If Device Is Nothing Then
      Set Device = New cESDevice
    End If
    'conn.BeginTrans
    
    
    If USE6080 Then

      ZoneID = Device.ZoneID
      If (left$(Device.Model, 2) = "EN") Or (left$(Device.Model, 2) = "ES") Then
        If ZoneID = 0 Then
          ' try to get zone id
          ZoneID = ScanZoneInfoListForSerial(Serial)
        End If
      End If
    End If
    
    SQL = "DELETE FROM Devices WHERE DeviceID = " & ID
    ConnExecute SQL

    SQL = "DELETE FROM Pagers WHERE identifier = " & q(Serial)
    ConnExecute SQL


    'log it
    rs.Open "alarms", conn, gCursorType, gLockType, adCmdTable
    rs.addnew

    rs("FC1") = 0  'device.FC1
    rs("FC2") = 0  'device.FC2
    rs("IDM") = ZoneID
    rs("IDL") = 0  'device.IDL
    rs("Status") = 0
    rs("Serial") = Serial
    rs("EventDate") = Now
    rs("Alarm") = 0
    rs("Tamper") = 0
    rs("IsLocator") = 0
    rs("Battery") = 0
    rs("LOCIDM") = 0
    rs("LOCIDL") = 0
    rs("ResidentID") = Device.ResidentID
    rs("RoomID") = Device.RoomID
    rs("EventType") = EVT_REMOVE_DEV
    rs("AlarmID") = 0  ' not an alarm
    rs("username") = Username
    rs("sessionid") = gSessionID
    rs("announce") = Device.Announce
    rs("Phone") = ""
    rs("Info") = ""
    rs.Update
    rs.Close

    If conn.Errors.Count > 0 Then
      'conn.RollbackTrans
      DeleteTransmitter = -1
      LogProgramError "Error " & conn.Errors(1).Number & " (" & conn.Errors(1).Description & ") at modEvents.DeleteTransmitter.Connection"
    Else
      'conn.CommitTrans
      DeleteTransmitter = 0
      Set Device = Devices.Device(Serial)
      
      If Not Device Is Nothing Then
        'device.swinger = True
        InBounds.ClearAllAlarmsBySerial Serial
        alarms.ClearAllAlarmsBySerial Serial
        Alerts.ClearAllAlarmsBySerial Serial
        Troubles.ClearAllAlarmsBySerial Serial
        LowBatts.ClearAllAlarmsBySerial Serial

        If Alerts.Pending() Then
          frmMain.ProcessAlerts
        End If
        If LowBatts.Pending Then
          frmMain.ProcessBatts
        End If
        If Assurs.Pending Then
          frmMain.ProcessAssurs False
        End If
        If Troubles.Pending Then
          frmMain.ProcessTroubles
        End If
        Devices.RemoveDevice Serial
        RemoveSerialDevice Serial
        If USE6080 Then
           Remove6080Device ZoneID
        End If
      End If
    End If
  End If


End Function

Sub AllCheckin(d As cESDevice)
  Dim RoomID As Long
  Dim Device As cESDevice
  
  
  
  RoomID = d.RoomID
  
  If RoomID <> 0 Then
    d.FetchRoom
  
    If ((d.RoomFlags And 1) = 1) Then
    For Each Device In Devices.Devices
      If Device.RoomID = RoomID Then
        If Device.AssurBit Then
          Device.AssurBit = 0
          LogAssurCheckin d, Nothing, EVT_ASSUR_CHECKIN
        End If
      End If
    Next
  End If
  End If
  
  



End Sub

Sub ClearInfoBox(ByVal AlarmID As Long)
  frmMain.ClearInfoBox AlarmID
End Sub


Sub ClearRemotes(ByVal AlarmID As Long)
  Dim Server    As cPageDevice
  Dim QueItem   As cPageItem
  Dim j As Long
  For Each Server In gPageDevices
    If Server.ProtocolID = PROTOCOL_REMOTE Then
      Server.RemoveByAlarmID AlarmID
    End If
  Next
End Sub

Sub LogOutGoing(ByVal s As String)
  Dim hfile              As Long
  Dim filename As String
  filename = App.Path & "\Push.Log"
  
  limitFileSize filename
  
  On Error Resume Next
    hfile = FreeFile
    Open filename For Append As hfile
    Print #hfile, Format$(Now, "hh:nn:ss") & " " & s
    Close hfile
  

End Sub


Sub ClearPages(ByVal AlarmID As Long)
  Dim pageDevice         As cPageDevice
  If MASTER Then
    For Each pageDevice In gPageDevices
      If pageDevice.ProtocolID <> PROTOCOL_REMOTE Then
        pageDevice.RemoveByAlarmID AlarmID
      End If
    Next
  End If

End Sub
