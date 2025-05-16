VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDefaultBeepTime As Long
Private mLocalBeepTime   As Long
Private mReBeepTime      As Long
Private mLocalLastSilenced As Double  ' date serial

Public alarms            As Collection  ' public for now

Public ManualClear       As Integer



Private mPending         As Boolean
Private mEndBeepTime     As Date

Private mname            As String

Public AssursReady       As Boolean

Public AlarmTime         As Double
Public SilenceTime       As Double

Private Consoles         As Collection

Public LocalAlarmTime    As Double
Public LocalSilenceTime  As Double







Public Function ToXML(ByVal Alarmtype As String, ByVal ConsoleID As String)

  Dim j                  As Long
  Dim alarm              As cAlarm
  Dim str                As String

  str = str & "<" & Alarmtype & ">" & vbCrLf

  For j = 1 To alarms.Count


    Set alarm = alarms(j)
    str = str & "<Alarm>" & vbCrLf
    str = str & alarm.ToXML(ConsoleID)
    'str = str & taggit("LastSilenced", XMLEncode(alarm.ConsoleSilenced(Consoleid))) & vbCrLf
    str = str & "</Alarm>" & vbCrLf
    If Alarmtype = "Troubles" Then
      If j >= MAX_REMOTE_TROUBLES Then
        Exit For               ' prevent overload of data
      End If
    End If

  Next

  ' these are all doubles
  str = str & taggit("SilenceTime", ConsoleSilenceTime(ConsoleID)) & vbCrLf  ' This will get the last silence for requesting console (ID)
  str = str & taggit("AlarmTime", ConsoleAlarmTime(ConsoleID)) & vbCrLf  ' This will get the last silence for requesting console (ID)
  str = str & taggit("HostTime", CDbl(Now)) & vbCrLf  '

  str = str & taggit("Beep", BeepTimer) & vbCrLf  ' beeptimer is either zero or non-zero for remotes

  str = str & "</" & Alarmtype & ">" & vbCrLf

  ToXML = str

End Function



Public Property Get LocalLastSilenced() As Double
  LocalLastSilenced = mLocalLastSilenced
End Property

Public Property Let LocalLastSilenced(ByVal Value As Double)
mLocalLastSilenced = Value

End Property

Public Property Get ConsoleAlarmTime(ByVal ConsoleID As String) As Double
        
        
        Dim Console            As cConsole

        On Error Resume Next

        If Consoles Is Nothing Then Set Consoles = New Collection

10      If Len(Trim$(ConsoleID)) Then

30        Set Console = Consoles(ConsoleID)
40        If Not (Console Is Nothing) Then
55          ConsoleAlarmTime = CDbl(Console.AlarmTime)
          Else
59          Set Console = New cConsole
60          Console.ConsoleID = ConsoleID
70          'Console.User = User
80          Console.AlarmTime = Now
90          Consoles.Add Console, ConsoleID & ""
            ConsoleAlarmTime = Console.AlarmTime
            
71        End If

81      End If
    
End Property

Public Property Let ConsoleAlarmTime(ByVal ConsoleID As String, ByVal TimeValue As Double)
        Dim Console            As cConsole

        On Error Resume Next
        
        If alarms.Count = 0 Then
          Set Consoles = New Collection
          Exit Property
          
        End If
        
        If Consoles Is Nothing Then Set Consoles = New Collection

10      If Len(Trim$(ConsoleID)) Then


30        Set Console = Consoles(ConsoleID)

40        If Console Is Nothing Then
50          Set Console = New cConsole
60          Console.ConsoleID = ConsoleID
70          'Console.User = User
80          Console.AlarmTime = TimeValue
90          Consoles.Add Console, ConsoleID & ""

100       Else
110         'Console.User = User
120         Console.AlarmTime = TimeValue

130       End If
140     End If



End Property





Public Function ConsoleLastSilence(ByVal ConsoleID As String) As String
'        Dim Console            As cConsole
'
'        On Error Resume Next
'
'10      If Len(Trim$(ConsoleID)) Then
'
'30        Set Console = Consoles(ConsoleID)
'40        If Not (Console Is Nothing) Then
'60          ConsoleLastSilence = CDbl(Console.SilenceTime)
'          Else
'            ConsoleLastSilence = 0
'70        End If
'
'80      End If

End Function

Public Property Get ConsoleSilenceTime(ByVal ConsoleID As String) As Double
        Dim Console            As cConsole

        On Error Resume Next
        
        If alarms.Count = 0 Then
          Set Consoles = New Collection
          Exit Property
        End If
        
        If Consoles Is Nothing Then Set Consoles = New Collection

10      If Len(Trim$(ConsoleID)) Then
30        Set Console = Consoles(ConsoleID)
40        If Console Is Nothing Then
            ConsoleSilenceTime = 0
100       Else
120         ConsoleSilenceTime = Console.SilenceTime
130       End If
140     End If


End Property

Public Property Let ConsoleSilenceTime(ByVal ConsoleID As String, ByVal TimeValue As Double)
        Dim Console            As cConsole

        On Error Resume Next
        
        If alarms.Count = 0 Then
          Set Consoles = New Collection
          Exit Property
          
        End If
        
        If Consoles Is Nothing Then Set Consoles = New Collection

10      If Len(Trim$(ConsoleID)) Then

30        Set Console = Consoles(ConsoleID)
40        If Console Is Nothing Then
50          Set Console = New cConsole
60          Console.ConsoleID = ConsoleID
80          Console.SilenceTime = TimeValue
90          Consoles.Add Console, ConsoleID & ""
100       Else



120         Console.SilenceTime = TimeValue

130       End If
140     End If


End Property


Public Property Get BeepTimer() As Long
  Dim a As cAlarm
  Dim j As Integer
  
  ' original backwards compatible. Don't Touch!
  
  If DefaultBeepTime < 0 And alarms.Count > 0 Then
    For j = 1 To alarms.Count
      Set a = alarms(j)
      If a.Silenced = 0 And a.ACKed = 0 Then
        BeepTimer = -1
        Exit For
      End If
    Next
  Else
    Dim Diff As Double
    Diff = DateDiff("s", Now, mEndBeepTime)

    BeepTimer = IIf(Diff > 0, 1, 0)
  End If
End Property

Public Property Let BeepTimer(ByVal Value As Long)
  
  ' original backwards compatible. Don't Touch!
  
  If MASTER Then
    If Value < 0 Then
      mEndBeepTime = DateAdd("yyyy", 1, Now)
    Else
      mEndBeepTime = DateAdd("s", Value, Now)
    End If
  Else
    If Value <> 0 Then
      mEndBeepTime = DateAdd("yyyy", 1, Now)
    Else
      mEndBeepTime = Now
    End If
  End If
End Property



Public Function FetchAssurData() As Boolean
  Dim alarm As cAlarm

  For Each alarm In alarms
    alarm.FetchInfo ' stubbed
  Next

End Function


Public Function CheckAutoClear() As Long
        Dim j As Integer
        Dim alarm As cAlarm
10       On Error GoTo CheckAutoClear_Error

20      For j = alarms.Count To 1 Step -1
30        Set alarm = alarms(j)
40        alarm.CheckAutoClear
50      Next

CheckAutoClear_Resume:
60       On Error GoTo 0
70       Exit Function

CheckAutoClear_Error:

80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.CheckAutoClear." & Erl
90      Resume CheckAutoClear_Resume

End Function

Public Sub Clear()
  Set alarms = New Collection
End Sub



Public Property Get Pending() As Boolean
  Pending = mPending
  mPending = False
End Property

'Public Property Get ByUniqueID(ByVal UniqueID As Long) As cAlarm
'  Dim Alarm As cAlarm
'  For Each Alarm In alarms
'    If Alarm.serial = UniqueID Then
'      Set ByUniqueID = Alarm
'      Exit For
'    End If
'  Next
'End Property

Public Property Get BySerialTamper(ByVal Serial As String) As cAlarm
  Dim alarm As cAlarm
  For Each alarm In alarms
    If alarm.Tamper Then
      If alarm.Serial = Serial Then
        Set BySerialTamper = alarm
        Exit For
      End If
    End If
  Next
End Property


Public Property Get BySerial(ByVal Serial As String) As cAlarm
  Dim alarm As cAlarm
  For Each alarm In alarms
    If alarm.Serial = Serial Then
      Set BySerial = alarm
      Exit For
    End If
  Next
End Property

' commented out 20080605 doesn't seem to haver any callers

'Public Function UpdateLocation(p As cESPacket) As Long
'        Dim serial    As String
'        Dim a         As cAlarm
'        Dim j         As Integer
'        'Exit Function
'
'10       On Error GoTo UpdateLocation_Error
'        If modMain.gDirectedNetwork Then
'          serial = p.LocatedSerial
'        Else
'20        serial = p.serial
'        End If
'
'30      For j = 1 To alarms.count
'40        Set a = alarms(j)
'50        If a.serial = serial Then
'60          If 0 <> a.AddLocator(p) Then ' ES series
''60          If 0 <> A.AddLocater(p.FirstHopMID, p.FirstHopUID, p.Level) Then ' FA series
'70            UpdateLocation = serial
'80          End If
'90          Exit For  ' Done
'100       End If
'110     Next
'
'UpdateLocation_Resume:
'120      On Error GoTo 0
'130      Exit Function
'
'UpdateLocation_Error:
'
'140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.UpdateLocation." & Erl
'150     Resume UpdateLocation_Resume
'
'End Function
Public Property Get Count() As Long
  Count = alarms.Count
End Property

Public Function ProcessPacket(p As cESPacket) As Boolean
  Dim alarm As cAlarm


  For Each alarm In alarms
    Debug.Print "ClassName = "; Me.name
    If (gDirectedNetwork) Then
      If p.IsLocatorPacket Then 'dbg "Have Locator Packet " & p.LocatedSerial & " " & p.Serial
        If p.LocatedSerial = alarm.Serial Then
          alarm.AddLocator p ' used to exit the loop here, but we need to handle devices that are in twice because of two inputs/alarms
        End If
      ElseIf (Configuration.OnlyLocators = 0) Then
        If p.Serial = alarm.Serial Then
          alarm.AddLocator p ' used to exit the loop here, but we need to handle devices that are in twice because of two inputs/alarms
        End If
      End If
    Else ' non directed network
      If p.Serial = alarm.Serial Then
        alarm.AddLocator p
      End If
    End If
  Next

End Function

Function Delete(ByVal index As Integer) As Long
  Dim alarm As cAlarm
  Debug.Print "cAlarms.Delete " & index
  If index > 0 And index <= alarms.Count Then
    Set alarm = alarms(index)
    alarms.Remove index
    Set alarm = Nothing
    Delete = 1
  End If
End Function
Function AddAssist(packet As cESPacket, Device As cESDevice, ByVal inputnum As Long, PriorAlarm As cAlarm) As cAlarm

  '+++++++++++ creates new alarm

  Dim a                  As cAlarm
  Dim j                  As Integer
  Dim Serial             As String
  Dim alarm              As cAlarm


  On Error GoTo AddAssist_Error

  Debug.Print "Adding Assist " & PriorAlarm.Serial




  Serial = PriorAlarm.Serial
  inputnum = PriorAlarm.inputnum  ' passed in, but overwritten here

  ' Check for Duplicate Alarm
  For j = 1 To alarms.Count
    Set a = alarms(j)
    If a.Serial = Serial And a.inputnum = inputnum And a.Alarmtype = EVT_ASSISTANCE Then
      ' duplicate
      'Set AddAssist = a
      Exit For                 ' Already in list
    End If
  Next


  If j <= alarms.Count Then    ' found and we're done
    Exit Function
  End If

  ' didn't find it, so create/add

  Set alarm = New cAlarm       ' might change to a more generic item

  alarm.FC1 = 0                'D.FC1
  alarm.FC2 = 0                'D.FC2
  alarm.IDM = PriorAlarm.IDM
  alarm.IDL = PriorAlarm.IDL

  ' from packet
  alarm.FirstHopSerial = PriorAlarm.FirstHopSerial
  alarm.STAT = PriorAlarm.STAT
  alarm.DateTime = Now
  alarm.alarm = PriorAlarm.alarm
  alarm.Tamper = PriorAlarm.Tamper
  alarm.IsLocator = PriorAlarm.IsLocator
  alarm.Battery = PriorAlarm.Battery
  'Alarm.LineLoss = P.LineLoss
  alarm.packet = PriorAlarm.packet  '     P.Payload
  alarm.Alarmtype = EVT_ASSISTANCE

  alarm.LEvel = PriorAlarm.LEvel
  alarm.Margin = PriorAlarm.Margin

  ' from device
  alarm.Serial = PriorAlarm.Serial



  alarm.Model = PriorAlarm.Model
  alarm.ResidentID = PriorAlarm.ResidentID
  alarm.RoomID = PriorAlarm.RoomID
  alarm.Preamble = PriorAlarm.Preamble
  '420       Alarm.AutoClearTime = d.AutoClear

  alarm.Description = PriorAlarm.Description
  alarm.Custom = PriorAlarm.Custom
  alarm.Announce = "Staff Assist " & PriorAlarm.Announce


  alarm.inputnum = PriorAlarm.inputnum  ' serial devices use a unique system generated inputnum
  GetAlarmText alarm

  alarm.locationtext = PriorAlarm.locationtext

  'Call dbgPackets("cAlarms in Add at AssignLocation Alarm.locationtext " & Alarm.locationtext)


  AlarmTime = Now

  If DefaultBeepTime < 0 Then
    mEndBeepTime = DateAdd("yyyy", 1, Now)
  Else
    mEndBeepTime = DateAdd("s", DefaultBeepTime, Now)
  End If

  alarm.AlarmName = Me.name
  
  '' ************ for linking to originating alarm
  
  alarm.PriorID = PriorAlarm.ID
  Set AddAssist = alarm
  alarms.Add alarm
  
  
  
  
  
  mPending = True ' public porperty Pending is read-only and gets reset thwn read




AddAssist_Resume:

If MASTER Then
  Set Consoles = New Collection
  LocalAlarmTime = Now
  LocalSilenceTime = 0
End If


On Error GoTo 0
Exit Function



AddAssist_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.AddAssist." & Erl
  Resume AddAssist_Resume

End Function


Function Add(p As cESPacket, d As cESDevice, ByVal inputnum As Long) As cAlarm

        Dim a                  As cAlarm
        Dim j                  As Integer
        '        Dim UniqueID  As Long
        Dim Serial             As String
        Dim alarm              As cAlarm

10      Debug.Print "Adding Alarm " & d.Serial


20      On Error GoTo Add_Error

30      Serial = d.Serial

40      For j = 1 To alarms.Count
50        Set a = alarms(j)
          ' When sent to External Alarms Window, Serial Input devices the inputnum is a unique ID to allow for multiple alarms
60        If a.Serial = Serial And a.inputnum = inputnum And a.Alarmtype = p.Alarmtype Then
70          Set Add = a
80          Exit For                 ' Already in list
90        End If
100     Next

110     If j > alarms.Count Then     ' didn't find it, so add
120       Set alarm = New cAlarm     ' might change to a more generic item
130       Set Add = alarm
140       alarm.FC1 = 0              'D.FC1
150       alarm.FC2 = 0              'D.FC2
160       alarm.IDM = d.IDM
170       alarm.IDL = d.IDL

          ' from packet
180       alarm.FirstHopSerial = p.FirstHopSerial
190       alarm.STAT = p.Status
200       alarm.DateTime = Now
210       alarm.alarm = p.alarm
220       alarm.Tamper = p.Tamper
230       alarm.IsLocator = p.IsLocator
240       alarm.Battery = p.Battery
          'Alarm.LineLoss = P.LineLoss
250       alarm.packet = p.HexPacket  '     P.Payload
260       alarm.Alarmtype = p.Alarmtype


          ' boost removed from here

270       'If d.CLSPTI = &H3E1D Then        ' 1221S Pendant
          '  Debug.Print "p.LEvel " & p.LEvel
280       '  alarm.LEvel = p.LEvel * (1 + (Configuration.boost / 100))
          '  Debug.Print "alarm.LEvel " & alarm.LEvel
290       'Else
300       alarm.LEvel = p.LEvel
310       'End If
320       alarm.Margin = p.Margin

          ' from device
330       If d.Serial = "" Then
340         alarm.Serial = p.Serial

350       Else
360         alarm.Serial = d.Serial

370       End If
380       alarm.Model = d.Model
390       alarm.ResidentID = d.ResidentID
400       alarm.RoomID = d.RoomID
410       alarm.Preamble = d.SerialPreamble
420       alarm.AutoClearTime = d.AutoClear

430       alarm.Description = GetESDesc(d.CLSPTI)
440       alarm.Custom = d.Custom
          '361       Alarm.AutoClearTime = d.
450       If d.IsSerialDevice <> 0 Then  ' if it's a serial device

460         alarm.AutoClearTime = d.SerialAutoClear

470         If d.SerialMessageLen > 0 Then  ' and we're extracting messages
480           alarm.Announce = p.SerialMessage
490         Else
500           alarm.Announce = d.Announce
510         End If
520       Else
530         alarm.IsPortable = d.IsPortable Or USE6080
540         If inputnum = 3 Then
550           alarm.Announce = d.Announce_B
560         ElseIf inputnum = 2 Then
570           alarm.Announce = d.Announce_A
580         Else
590           alarm.Announce = d.Announce
600         End If
610       End If

620       alarm.inputnum = inputnum  ' serial devices use a unique system generated inputnum
630       GetAlarmText alarm

          ' new with 6080 device !!!! 5/20/2013
          'Alarm.locationtext = p.LocatedPartionName1  ' up to 4 LocatedPartionNames may be available but not used
          ' changed 8/22/14 , processed in mod6080.Process6080Packet
          'If IsPortable Then


640 AssignLocation:             alarm.locationtext = d.LastLocationText  ' up to 4 LocatedPartionNames may be available but not used

650       Call dbgPackets("cAlarms in Add at AssignLocation Alarm.locationtext " & alarm.locationtext)
          

660       AlarmTime = Now

670       If DefaultBeepTime < 0 Then
680         mEndBeepTime = DateAdd("yyyy", 1, Now)
690       Else
700         mEndBeepTime = DateAdd("s", DefaultBeepTime, Now)
710       End If

720       alarms.Add alarm

730       alarm.AlarmName = Me.name
740       mPending = True
750     End If


Add_Resume:
760     If MASTER Then
770       Set Consoles = New Collection
780       LocalAlarmTime = Now
790       LocalSilenceTime = 0
800     End If


810     On Error GoTo 0
820     Exit Function

Add_Error:

830     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.Add." & Erl
840     Resume Add_Resume


End Function

Public Function GetAlarmText(alarm As cAlarm)
          Dim Rs            As Recordset
        Dim ResidentText  As String
        Dim RoomID        As Long
        Dim RoomText      As String


10      On Error GoTo GetAlarmText_Error

  
20      If Not alarm Is Nothing Then
30        RoomID = alarm.RoomID

40        Set Rs = ConnExecute("Select Phone, NameLast, NameFirst, RoomID  FROM Residents WHERE ResidentID = " & alarm.ResidentID)
50        If Not Rs.EOF Then
      
60          alarm.ResidentText = Trim(Rs("NameFirst") & " " & Rs("NameLast"))
70          alarm.Phone = Rs("phone") & ""
      
80        End If
90        Rs.Close

100       Set Rs = ConnExecute("Select Room FROM Rooms WHERE RoomID = " & RoomID)
110       If Not Rs.EOF Then
120         alarm.RoomText = Trim(Rs("room") & "")
130       End If
140       Rs.Close

150     End If
GetAlarmText_Resume:
160     On Error GoTo 0
170     Exit Function

GetAlarmText_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.GetAlarmText." & Erl
190     Resume GetAlarmText_Resume



End Function


Public Function AddAlarm(Item As Object) As Long

        Dim a         As Object
        Dim j         As Integer
        Dim Serial    As String
        Dim inputnum As Long
10       On Error GoTo AddAlarm_Error

20      Serial = Item.Serial
        If TypeOf Item Is cAlarm Then
          inputnum = Item.inputnum
        End If
30      For j = 1 To alarms.Count
40        Set a = alarms(j)
50        If a.Serial = Serial And a.inputnum = inputnum Then
60          Exit For  ' Already in list
70        End If
80      Next
90      If j > alarms.Count Then  ' didn't find it, so add
          
100       alarms.Add Item
110       AddAlarm = j
120       BeepTimer = DefaultBeepTime
          AlarmTime = Now
          mEndBeepTime = DateAdd("s", DefaultBeepTime, Now)
          
130       mPending = True
140     End If

        

AddAlarm_Resume:
         If MASTER Then
          Set Consoles = New Collection
          LocalAlarmTime = Now
          LocalSilenceTime = 0
         End If

150      On Error GoTo 0
160      Exit Function

AddAlarm_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.AddAlarm." & Erl
180     Resume AddAlarm_Resume


End Function
Function ClearAllAlarms()
  Set alarms = New Collection
End Function
  
Function ClearAllAlarmsBySerial(ByVal Serial As String)
          Dim j As Integer
10       On Error GoTo ClearAllAlarmsBySerial_Error

20        For j = alarms.Count To 1 Step -1
30          If alarms(j).Serial = Serial Then
              RemovePageRequest alarms(j).alarmID
40            alarms.Remove j
50            mPending = True
60            Exit For
70          End If
80        Next

ClearAllAlarmsBySerial_Resume:
90       On Error GoTo 0
100      Exit Function

ClearAllAlarmsBySerial_Error:

110     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.ClearAllAlarmsBySerial." & Erl
120     Resume ClearAllAlarmsBySerial_Resume

End Function

Function RemoveAlarm(alarm As Object, ByVal EventType As Long) As cAlarm
      ' removes but does not update any data
      ' returns 0 if alarm not found, else, returns postion in list

        Dim a As cAlarm
        Dim j As Integer
        Dim Serial As String
        Dim inputnum As Long


10      On Error GoTo RemoveAlarm_Error

20      Serial = alarm.Serial
30      If TypeOf alarm Is cAlarm Then
40        inputnum = alarm.inputnum
50      End If
60      For j = 1 To alarms.Count
70        Set a = alarms(j)
80        If a.Serial = Serial And a.inputnum = inputnum And a.Alarmtype = EventType Then  ' found it
90          Set RemoveAlarm = a
100         alarms.Remove j
105         modEvents.ClearRemotes a.alarmID
110         mPending = True
120         Exit For
130       End If
140     Next

RemoveAlarm_Resume:
150     On Error GoTo 0
160     Exit Function

RemoveAlarm_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.RemoveAlarm." & Erl
180     Resume RemoveAlarm_Resume


End Function
Function GetAlarm(d As cESDevice, ByVal inputnum As Long) As cAlarm
        Dim a         As cAlarm
        Dim j         As Integer
        Dim Serial    As String

10      On Error GoTo GetAlarm_Error

20      Serial = d.Serial
30      For j = 1 To alarms.Count
40        Set a = alarms(j)
50        If a.Serial = Serial Then  ' found it
      
60          If a.inputnum = inputnum Then ' should be inputnum
70            Set GetAlarm = a
80            Exit For
90          End If
100       End If
110     Next

GetAlarm_Resume:
120      On Error GoTo 0
130      Exit Function

GetAlarm_Error:

140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.GetAlarm." & Erl
150     Resume GetAlarm_Resume


End Function

Function Restore(d As cESDevice, ByVal inputnum As Long) As cAlarm
                                    ' should be inputnum
      ' Just need this for the Alarm Reference

        Dim a         As cAlarm
        Dim j         As Integer
        Dim Serial    As String


10      On Error GoTo Restore_Error

20      Serial = d.Serial
30      For j = 1 To alarms.Count
40        Set a = alarms(j)
50        If a.Serial = Serial Then  ' found it
60          If a.inputnum = inputnum Then ' should be inputnum
70            Set Restore = a
80            Exit For
90          End If
100       End If
110     Next

Restore_Resume:
120     On Error GoTo 0
130     Exit Function

Restore_Error:

140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.Restore." & Erl
150     Resume Restore_Resume


End Function

Function RestoreAlarm(alarm As cAlarm, ByVal EventType As Long) As cAlarm

      ' Just need this for the Alarm Reference

        Dim a         As cAlarm
        Dim j         As Integer
10      On Error GoTo Restore_Error

        'Serial = Alarm.Serial
20      For j = 1 To alarms.Count
30        Set a = alarms(j)
40        If a.ID = alarm.ID Then ' found it
50          If a.Alarmtype = EventType Then
60            Set RestoreAlarm = a
70            Exit For
80          End If
90        End If
100     Next

Restore_Resume:
110     On Error GoTo 0
120     Exit Function

Restore_Error:

130     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.Restore." & Erl
140     Resume Restore_Resume


End Function

Function Acknowledge(d As cESDevice, ByVal EventType As Long) As cAlarm

      ' Sets Ack to current time
      ' returns 0 if alarm not found, else, returns postion in list

        Dim a         As cAlarm
        Dim j         As Integer
        Dim Serial  As String

10       On Error GoTo Acknowledge_Error

20      Serial = d.Serial
30      For j = 1 To alarms.Count
40        Set a = alarms(j)
    
50        If a.Serial = Serial Then  ' found it
            If a.Alarmtype = EventType Then
60          a.alarm = d.alarm
70          a.STAT = d.STAT
80          a.ACKTime = Now
90          a.ACKed = 1
100         Set Acknowledge = a
110         mPending = True
120         Exit For
            End If
130       End If
140     Next

Acknowledge_Resume:
150      On Error GoTo 0
160      Exit Function

Acknowledge_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.Acknowledge." & Erl
180     Resume Acknowledge_Resume


End Function
Public Function AcknowledgeAlarm(d As cESDevice, a As cAlarm, ByVal EventType As Long) As cAlarm

      ' Sets Ack to current time
      ' returns 0 if alarm not found, else, returns postion in list

      '       Dim A         As cAlarm
      '       Dim j         As Integer
        Dim Serial  As String
        Dim inputnum  As Long

10      On Error GoTo Acknowledge_Error

20      Serial = d.Serial
30      inputnum = a.inputnum
        '30      For j = 1 To alarms.Count
        '40        Set A = alarms(j)

40      If a.Serial = Serial And a.inputnum = inputnum Then  ' found it
          '            If A.AlarmType = eventtype Then
50        If inputnum = 2 Then
60          a.alarm = d.Alarm_A  ' why set here?
70        Else
80          a.alarm = d.alarm  ' why set here?
90        End If
100       a.STAT = d.STAT
110       a.ACKTime = Now
120       a.ACKed = 1
130       Set AcknowledgeAlarm = a
140       mPending = True
          '120         Exit For
          '            End If
150     End If
        '140     Next

Acknowledge_Resume:
160     On Error GoTo 0
170     Exit Function

Acknowledge_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarms.Acknowledge." & Erl
190     Resume Acknowledge_Resume


End Function


Private Sub Class_Initialize()
  Set alarms = New Collection
  Set Consoles = New Collection
  DefaultBeepTime = 60
  mEndBeepTime = Now
End Sub

Private Sub Class_Terminate()
  Set alarms = Nothing
End Sub


Public Function RemoveAlarmByID(ByVal ID As Long) As Long
  Dim j As Integer
  Dim a As cAlarm
  For j = 1 To alarms.Count
    Set a = alarms(j)
    If a.ID = ID Then
      alarms.Remove j
      Exit For
    End If
  Next
End Function

Public Property Get DefaultBeepTime() As Long

  DefaultBeepTime = mDefaultBeepTime

End Property

Public Property Let DefaultBeepTime(ByVal DefaultBeepTime As Long)
  mDefaultBeepTime = DefaultBeepTime
End Property

Public Property Get ReBeepTime() As Long
  ReBeepTime = mReBeepTime
End Property

Public Property Let ReBeepTime(ByVal ReBeepTime As Long)

  mReBeepTime = ReBeepTime

End Property

Public Property Get LocalBeepTime() As Long

  LocalBeepTime = mLocalBeepTime

End Property

Public Property Let LocalBeepTime(ByVal LocalBeepTime As Long)

  mLocalBeepTime = LocalBeepTime

End Property

Public Property Get name() As String
  name = mname
End Property

Public Property Let name(ByVal name As String)
  mname = name
End Property
