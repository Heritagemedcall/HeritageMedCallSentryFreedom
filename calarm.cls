VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Disposition    As String


Public AlarmName      As String

Private mSendCancel   As Integer
Private mMarquisCode  As String

Public Description    As String
Public Custom         As String

Public UID            As String  ' four hex bytes

Public Model          As String

Private MID           As Long  ' me, this alarm's ID
Private mAlarmID      As Long  ' for referencial linking to events for this alarm


Public Guid           As Long  ' Actual ID of record

Public PriorID        As Long

Public LastAlarm      As Date  ' date/time of last alarm

Public LastSupervise  As Date  ' date/time of last Supervise
Public LastTest       As Date  ' date/time of last test

Public alarm          As Integer  ' 0 or 1
Public Tamper         As Integer  ' 0 or 1
Public Battery        As Integer  ' 0 or 1
Public IsLocator      As Integer  ' 0 or 1

Public ResidentText   As String
Public RoomText       As String
Public locationtext   As String

Private mAnnounce     As String
Public Phone          As String
Public info           As String  ' notes
Public Preamble       As String

Public Serial         As String

Public LEvel          As Integer ' to 255
Public Margin         As Integer ' to 255


Public LOCIDM         As Long     ' not used FA
Public LOCIDL         As Long     ' not used FA
Public FC1            As Integer  ' not used FA
Public FC2            As Integer  ' not used FA
Public IDM            As Long     ' (6080 device ID ) not used FA
Public IDL            As Long     ' (6080 device fixed, mobile, etc ) not used FA
Public STAT           As Long
Public LOCSIG         As Integer  ' not used FA
Public DateTime       As Date

Public IsPortable     As Integer  ' Flag... used for locating

Public Latching       As Integer  ' 0 or 1

Public ACKTime        As Date
Public ACKed          As Integer  ' 0 or 1
'Public AckUser        As String
Public Username       As String

Private mResponder    As String

Public Restored       As Integer

Public SilenceTime    As Date
Public Silenced       As Integer  ' 0 or 1
Public SilenceUser    As String  ' 0 or 1

Public ResidentID     As Long     ' if 0 then unassigned
Public RoomID         As Long     ' if 0 then unassigned ' this is +NOT+ the resident's room
Public inputnum       As Long

Public TimeStamp      As Long  ' miliseconds from TimeGetTime
Private mAlarmtype    As Long
Private mEventType    As Long


' not needed for Client Console
Public Owner          As String   ' not used

Private mRepeater1    As String
Public Repeater2      As String
Public Repeater3      As String

Private mSignal1      As Double
Public Signal2        As Double
Public Signal3        As Double

Public AVGSignal      As Double
Public FirstHopSerial As String

Private mPacket       As String
Private mLastRestore  As Date  ' date/time of last restore
Private mFirstHops    As Collection  ' ES only
Private mReadyToPost  As Boolean
Private mCancelText   As String


Private mPosted       As Boolean

Private mAutoClearTime As Long

Private Consoles As Collection

Public Function ReadFromDB(ByVal AlarmID As Long) As cAlarm

  Dim rs As ADODB.Recordset
  Dim SQL As String
  Dim TempAlarm As cAlarm
  
  SQL = "SELECT * from Alarms where ID = " & AlarmID
  Set rs = ConnExecute(SQL)
  If Not rs.EOF Then
    Set TempAlarm = New cAlarm
    TempAlarm.AlarmID = AlarmID
    TempAlarm.Alarmtype = Val(rs("eventtype") & "")
    TempAlarm.RoomID = Val(rs("roomid") & "")
    TempAlarm.ResidentID = Val(rs("residentid") & "")
    TempAlarm.inputnum = Val(rs("inputnum") & "")
    TempAlarm.Announce = rs("announce") & ""
    TempAlarm.Serial = rs("serial") & ""
    TempAlarm.locationtext = rs("userdata") & ""
    
    Set ReadFromDB = TempAlarm
  End If
  rs.Close
  Set rs = Nothing
  
  
  
  





End Function


Public Function Clone() As cAlarm
  Dim TempClone As cAlarm
  Set TempClone = New cAlarm
  
  TempClone.alarm = alarm
  TempClone.AlarmID = AlarmID
  TempClone.AlarmName = AlarmName
  TempClone.Alarmtype = Alarmtype
  TempClone.Announce = Announce
  TempClone.AutoClearTime = AutoClearTime
  TempClone.AVGSignal = AVGSignal
  TempClone.Battery = Battery
  TempClone.CancelText = CancelText
  TempClone.Custom = Custom
  TempClone.Description = Description
  TempClone.EventType = EventType
  TempClone.FC1 = FC1
  TempClone.FC2 = FC2
  TempClone.FirstHopSerial = FirstHopSerial
  TempClone.Guid = Guid
  TempClone.ID = ID
  TempClone.IDL = IDL
  TempClone.IDM = IDM
  TempClone.info = info
  TempClone.inputnum = inputnum
  TempClone.IsLocator = IsLocator
  TempClone.IsPortable = IsPortable
  TempClone.LastAlarm = LastAlarm
  TempClone.LastRestore = LastRestore
  TempClone.LastSupervise = LastSupervise
  TempClone.LastTest = LastTest
  TempClone.Latching = Latching
  TempClone.LEvel = LEvel
  TempClone.locationtext = locationtext
  
  
  TempClone.LOCIDL = LOCIDL
  TempClone.LOCIDM = LOCIDM
  TempClone.LOCSIG = LOCSIG
  TempClone.Margin = Margin
  TempClone.Model = Model
  TempClone.Owner = Owner
  TempClone.packet = packet
  'TempClone.packetbytes = packetbytes
  TempClone.Phone = Phone
  TempClone.Posted = Posted
  TempClone.Preamble = Preamble
  TempClone.ReadyToPost = ReadyToPost
  TempClone.Repeater1 = Repeater1
  TempClone.Repeater2 = Repeater2
  TempClone.Repeater3 = Repeater3
  
  TempClone.ResidentID = ResidentID
  TempClone.ResidentText = ResidentText
  TempClone.Restored = Restored
  TempClone.RoomID = RoomID
  TempClone.RoomText = RoomText
  TempClone.SendCancel = SendCancel
  TempClone.Serial = Serial
  TempClone.Signal1 = Signal1
  TempClone.Signal2 = Signal2
  TempClone.Signal3 = Signal3
  TempClone.Silenced = Silenced
  TempClone.SilenceTime = SilenceTime
  TempClone.SilenceUser = SilenceUser
  TempClone.STAT = STAT
  TempClone.Tamper = Tamper
  TempClone.TimeStamp = TimeStamp
  TempClone.UID = UID
  Set Clone = TempClone
  Set TempClone = Nothing

End Function

Public Property Get Announce() As String
    Announce = mAnnounce
  
End Property

Public Property Let Announce(ByVal Announce As String)
  mAnnounce = Announce
End Property

Public Property Get EventType() As Long
  EventType = mEventType
End Property

Public Property Let EventType(ByVal EventType As Long)
  mEventType = EventType
End Property

Public Property Get Alarmtype() As Long
  Alarmtype = mAlarmtype
End Property

Public Property Let Alarmtype(ByVal Alarmtype As Long)
  mAlarmtype = Alarmtype
End Property

Public Function ConsoleUnSilence(ByVal ConsoleID As String, ByVal User As String) As Long
        Dim Console            As cConsole
        
10      On Error Resume Next

20      If Len(Trim$(ConsoleID)) Then

30        Set Console = Consoles(ConsoleID)

40        If Console Is Nothing Then  ' add it if not silenced before
50          Set Console = New cConsole
60          Console.ConsoleID = ConsoleID
70          Console.User = User
80          Console.SilenceTime = 0
85          Console.AlarmTime = CDbl(Now)
90          Consoles.Add Console, ConsoleID & ""
          Else
91          Console.User = User
95          Console.SilenceTime = 0
99          Console.AlarmTime = CDbl(Now)
          
100       End If
110     End If

End Function


Public Function ConsoleSilence(ByVal ConsoleID As String, ByVal User As String) As Long
        Dim Console            As cConsole
        
10      On Error Resume Next

20      If Len(Trim$(ConsoleID)) Then

30        Set Console = Consoles(ConsoleID)

40        If Console Is Nothing Then  ' add it if not silenced before
50          Set Console = New cConsole
60          Console.ConsoleID = ConsoleID
70          Console.User = User
80          Console.SilenceTime = Now
90          Consoles.Add Console, ConsoleID & ""
          Else
91          Console.User = User
92          Console.SilenceTime = Now
          
100       End If
110     End If

End Function

Public Function ConsoleSilenceTime(ByVal ConsoleID As String) As Double
  Dim Console            As cConsole

  On Error Resume Next
  ConsoleSilenceTime = 0

  If Len(Trim$(ConsoleID)) Then
    Set Console = Consoles(ConsoleID)
    If Not (Console Is Nothing) Then  ' add it if not silenced before
      ConsoleSilenceTime = Console.SilenceTime
    End If
  End If

End Function

Public Function FetchInfo() As String
  ' stubbed
End Function

Public Function AssurToXML() As String
  'Dim fs As String '
  
  AssurToXML = taggit("serial", Me.Serial) & vbCrLf & _
  taggit("resident", XMLEncode(Me.ResidentText)) & vbCrLf & _
  taggit("room", XMLEncode(Me.RoomText)) & vbCrLf & _
  taggit("phone", XMLEncode(Me.Phone)) & vbCrLf

  'AssurToXML = fs '.ToString

  'Set fs = Nothing
End Function




Public Function ToXML(ByVal ConsoleID As String) As String
  'Dim fs As String '
  'called from modclients.Client_GetAlarms for each alarm, alert etc

  ToXML = taggit("serial", Me.Serial) & vbCrLf & taggit("id", Me.ID) & vbCrLf & taggit("alarmid", Me.AlarmID) & vbCrLf & _
  taggit("inputnum", Me.inputnum) & vbCrLf & _
  taggit("locationtext", XMLEncode(Me.locationtext)) & vbCrLf & _
  taggit("datetime", Me.DateTime) & vbCrLf & _
  taggit("announce", XMLEncode(Me.Announce)) & vbCrLf & _
  taggit("acked", XMLEncode(Me.ACKed)) & vbCrLf & _
  taggit("acktime", XMLEncode(Me.ACKTime)) & vbCrLf & _
  taggit("alarmtype", XMLEncode(Me.Alarmtype)) & vbCrLf & _
  taggit("eventtype", XMLEncode(Me.EventType)) & vbCrLf ' only can have so many line-continuations
  
  
  ToXML = ToXML & taggit("timestamp", XMLEncode(Me.TimeStamp)) & vbCrLf & _
  taggit("silenced", XMLEncode(Me.Silenced)) & vbCrLf & _
  taggit("silencetime", XMLEncode(Me.SilenceTime)) & vbCrLf & _
  taggit("description", XMLEncode(Me.Description)) & vbCrLf & _
  taggit("custom", XMLEncode(Me.Custom)) & vbCrLf & _
  taggit("resident", XMLEncode(Me.ResidentText)) & vbCrLf & _
  taggit("roomtext", XMLEncode(Me.RoomText)) & vbCrLf & _
  taggit("roomid", XMLEncode(Me.RoomID)) & vbCrLf & _
  taggit("model", XMLEncode(Me.Model)) & vbCrLf
  
  ToXML = ToXML & taggit("responder", XMLEncode(Me.Responder)) & vbCrLf & _
  taggit("priorid", XMLEncode(Me.PriorID)) & vbCrLf
  
  ToXML = ToXML & taggit("silencetime", XMLEncode(ConsoleSilenceTime(ConsoleID))) & vbCrLf
  

End Function


Public Property Get AutoClear() As Boolean
  AutoClear = (mAutoClearTime > 0)
End Property

Public Property Get AutoClearTime() As Long
  AutoClearTime = mAutoClearTime / 60
End Property

Public Property Let AutoClearTime(ByVal Value As Long)
  mAutoClearTime = Value * 60 ' convert minutes to seconds
End Property

Public Function CheckAutoClear() As Long
  Dim i As Long
  Dim d As cESDevice
  Dim EventType As Long
  Select Case Alarmtype
    Case EVT_EMERGENCY
      EventType = EVT_EMERGENCY_AUTOACK
    Case EVT_ALERT
      EventType = EVT_ALERT_AUTOACK
    Case EVT_EXTERN
      EventType = EVT_EXTERN_AUTOACK
    Case Else
      EventType = 0
  End Select
    
  If (AutoClear) And (EventType <> 0) Then  ' AutoClear is a BOOLEAN!
    i = DateDiff("s", DateTime, Now)   ' returns seconds  'not minutes
    'Debug.Print "autoclear seconds " & mAutoClearTime & " elapsed " & DateDiff("s", DateTime, Now)
    If i >= mAutoClearTime Then 'Autoclear
        Set d = Devices.Device(Serial)
        If Not d Is Nothing Then
        'If d.Serial = Serial Then
          ' no packet
          PostEvent d, Nothing, Me, EventType, Me.inputnum
          'Exit For
        End If
      
      'Next
    End If
  End If
  '
End Function




Public Function AddLocator(p As cESPacket) As Integer
      
     ' Debug.Print "cAlarm.addlocator " & p.LocatedSerial
      ' returns 0 if locator not added.
      ' else returns 1 upon first locator

        Dim i As Integer

10      On Error GoTo AddLocator_Error
20      i = DateDiff("s", DateTime, p.DateTime)
  
30      If i <= LocatorWaitTime Then
40        mFirstHops.Add p
          'dbgloc "cAlarm.addlocator " & IIf(gDirectedNetwork, p.serial, p.FirstHopSerial) & " " & IIf(gDirectedNetwork, p.LocatedLevel, p.Level) & "  " & Me.serial
        Else
          'dbgloc "Wait time expired"
50      End If

AddLocator_Resume:
60      On Error GoTo 0
70      Exit Function

AddLocator_Error:

80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarm.AddLocator." & Erl
90      Resume AddLocator_Resume

End Function
'Public Function DNProcessLocations() As Boolean
'  ' sort on average FirstHopUID
'
'  Dim j           As Integer
'  Dim i           As Integer
'  Dim count       As Integer
'  Dim Repeaters As Collection
'  Dim Repeater  As cRepeater
'
'' need fixed
''Stop
'
'  count = mFirstHops.count
'  If IsPortable Then
'    If count > 0 Then
'      'dbg "Processing Locations " & vbCrLf
'      i = DateDiff("s", DateTime, Now)
'      If i > LocatorWaitTime Then
'        ' ******************* the new way
'        If Repeaters Is Nothing Then
'          Set Repeaters = New Collection
'        End If
'
'        For j = 1 To count  ' accumulate distinct repeaters
'          For i = 1 To Repeaters.count
'            Set Repeater = Repeaters(i)
'            If Repeater.Serial = mFirstHops(j).LocatedSerial Then
'              Exit For
'            End If
'          Next
'          If i > Repeaters.count Then
'            Set Repeater = New cRepeater
'            Repeaters.Add Repeater
'            Repeater.Serial = mFirstHops(j).LocatedSerial
'          End If
'
'          Repeater.Level = Repeater.Level + mFirstHops(j).LocatedLevel
'          Repeater.HitCount = Repeater.HitCount + 1
'        Next
'
'        For j = 1 To Repeaters.count  ' avarege signals for each repeater
'          Set Repeater = Repeaters(j)
'          Repeater.Level = Repeater.Level / Repeater.HitCount
'        Next
'
'
'        SortRepeaters Repeaters  ' get in decending order
'        'Debug.Print "Process Locations 'Repeaters '"
'        For j = 1 To Repeaters.count  ' test printout
'          'Debug.Print Repeaters(j).Serial, Repeaters(j).Level
'        Next
'
'        ' not sure why we need to clear this here
'        Set mFirstHops = New Collection
'
'        If Repeaters.count >= 1 Then
'          Repeater1 = Repeaters(1).Serial  ' Right("00000000" & Hex(loc1(1).FirstHopMID) & Hex(loc1(1).FirstHopUID), 8)
'          Signal1 = Format(Repeaters(1).Level, "0")
'        End If
'        If Repeaters.count >= 2 Then
'          Repeater2 = Repeaters(2).Serial
'          Signal2 = Format(Repeaters(2).Level, "0")
'        End If
'        If Repeaters.count >= 3 Then
'          Repeater3 = Repeaters(3).Serial
'          Signal3 = Format(Repeaters(3).Level, "0")
'        End If
'
'        AVGSignal = (Signal1 + Signal2 + Signal3) / 3
'
'        DNProcessLocations = True
'        ReadyToPost = True
'      End If
'    End If
'  Else
'    DNProcessLocations = True
'    ReadyToPost = True
'  End If
'
'
'
'
'
'
'
'End Function

Public Function ProcessLocations() As Boolean
      ' sort on average FirstHopUID

        Dim j                  As Integer
        Dim i                  As Double
        Dim Count              As Integer
        Dim Repeaters          As Collection
        Dim Repeater           As cRepeater
        Dim Serial             As String
        Dim LEvel              As Single
        Dim p                  As cESPacket

        'mfirsthops is a misnomer for directed network... It's accumulation of locators
        'Repeaters is a misnomer for directed network...  It too is accumulation of locators

10      On Error GoTo ProcessLocations_Error

20      If IsPortable And (USE6080 = 0) Then   ' only locate portable devices
30        Count = mFirstHops.Count

40        i = DateDiff("s", DateTime, Now)


          If i > LocatorWaitTime Then  ' it's time to see what we got!
          Debug.Print "calarm ProcessLocations count "; Count
          
50        If Count > 0 Then
            'dbg "Processing Locations " & vbCrLf


60

              ' ******************* the new way
70            If Repeaters Is Nothing Then
80              Set Repeaters = New Collection
90            End If

100           For j = 1 To mFirstHops.Count  ' accumulate distinct repeaters/locators
110             Set p = mFirstHops(j)
120             For i = 1 To Repeaters.Count  ' note: repeaters are neccessarily Repeaters
130               Set Repeater = Repeaters(i)
140               If gDirectedNetwork Then

150                 If Configuration.OnlyLocators = 0 Then  ' use both repeaters and locators

160                   If p.IsLocatorPacket Then             ' use LOCATOR data format
170                     Serial = mFirstHops(j).Serial
180                     LEvel = mFirstHops(j).LocatedLevel
190                   Else                                  ' use REPEATER data format
200                     Serial = mFirstHops(j).FirstHopSerial
210                     LEvel = mFirstHops(j).LEvel
220                   End If
230                 Else                                    ' ONLY LOCATORS
240                   Serial = mFirstHops(j).Serial         ' use LOCATOR data format
250                   LEvel = mFirstHops(j).LocatedLevel
260                 End If


270               Else                                      ' NO Locators in NON-DNet system
280                 Serial = mFirstHops(j).FirstHopSerial   ' use REPEATER data format
290                 LEvel = mFirstHops(j).LEvel
300               End If
310               If Repeater.Serial = Serial Then
320                 Exit For
330               End If
340             Next

350             If i > Repeaters.Count Then                 ' not really a repeater, just old nomenclature
360               Set Repeater = New cRepeater
370               Repeaters.Add Repeater
380               Repeater.Serial = Serial
390             End If

400             Repeater.LEvel = Repeater.LEvel + LEvel
410             Repeater.HitCount = Repeater.HitCount + 1
420           Next

430           For j = 1 To Repeaters.Count  ' redo average of signals for each repeater
440             Set Repeater = Repeaters(j)
450             Repeater.LEvel = Repeater.LEvel / Repeater.HitCount
460           Next

470           SortRepeaters Repeaters  ' get in decending order
              'Debug.Print "Process Locations 'Repeaters '"
480           For j = 1 To Repeaters.Count  ' test printout
                'Debug.Print Repeaters(j).Serial, Repeaters(j).Level
490           Next

              ' not sure why we need to clear this here, other than to clean up
500           Set mFirstHops = New Collection

510           If Repeaters.Count >= 1 Then
520             Repeater1 = Repeaters(1).Serial  ' Right("00000000" & Hex(loc1(1).FirstHopMID) & Hex(loc1(1).FirstHopUID), 8)
530             Signal1 = Format(Repeaters(1).LEvel, "0")
540           End If
550           If Repeaters.Count >= 2 Then
560             Repeater2 = Repeaters(2).Serial
570             Signal2 = Format(Repeaters(2).LEvel, "0")
580           End If
590           If Repeaters.Count >= 3 Then
600             Repeater3 = Repeaters(3).Serial
610             Signal3 = Format(Repeaters(3).LEvel, "0")
620           End If

630           AVGSignal = (Signal1 + Signal2 + Signal3) / 3

640           ProcessLocations = True
650           ReadyToPost = True

            Else '  Count = 0
               If i > LocatorWaitTime Then
               ProcessLocations = True
               ReadyToPost = True
               ' DoEvents
               End If
660         End If  ' i > LocatorWaitTime
            ' just returns if nothing to process
670       Else  ' count = 0
            ' what do we do if no count?
680         If i > LocatorWaitTime Then  ' it's time out for locators... we got nothin'!
690           ProcessLocations = True
700           ReadyToPost = True
710         End If
720       End If
730     Else  ' not a portable or not a 6080
740       ProcessLocations = True
750       ReadyToPost = True
760     End If

ProcessLocations_Resume:
770     On Error GoTo 0
780     Exit Function

ProcessLocations_Error:

790     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarm.ProcessLocations." & Erl
800     Resume ProcessLocations_Resume


End Function
'Sub DNSortLocations(a() As cESPacket)
'  Dim offset  As Integer
'  Dim maxrow  As Integer
'  Dim limit   As Integer
'  Dim Row     As Integer
'  Dim switch  As Integer
'  Dim MinRow  As Integer
'  Dim Temp
'
'  maxrow = UBound(a)
'  MinRow = 1
'
'  offset = maxrow \ 2
'  Do While offset > 0
'    limit = maxrow - offset
'    Do
'      switch = 0
'      For Row = MinRow To limit
'        'If a(row) > a(row + offset) Then
'        If a(Row).LocatedLevel < a(Row + offset).LocatedLevel Then    ' may need to incorporate margin
'          Set Temp = a(Row)
'          Set a(Row) = a(Row + offset)
'          Set a(Row + offset) = Temp
'          Set Temp = Nothing
'          switch = Row
'        End If
'      Next Row
'      limit = switch - offset
'    Loop While switch
'
'    offset = offset \ 2
'  Loop
'
'
'End Sub
'Sub SortLocations(a() As cESPacket)
'  Dim offset  As Integer
'  Dim maxrow  As Integer
'  Dim limit   As Integer
'  Dim Row     As Integer
'  Dim switch  As Integer
'  Dim MinRow  As Integer
'  Dim Temp
'
'  maxrow = UBound(a)
'  MinRow = 1
'
'  offset = maxrow \ 2
'  Do While offset > 0
'    limit = maxrow - offset
'    Do
'      switch = 0
'      For Row = MinRow To limit
'        'If a(row) > a(row + offset) Then
'        If a(Row).level < a(Row + offset).level Then  ' may need to incorporate margin
'          Set Temp = a(Row)
'          Set a(Row) = a(Row + offset)
'          Set a(Row + offset) = Temp
'          Set Temp = Nothing
'          switch = Row
'        End If
'      Next Row
'      limit = switch - offset
'    Loop While switch
'
'    offset = offset \ 2
'  Loop
'
'
'End Sub
'Public Property Get UniqueID() As Long
'  UniqueID = IDM * 256 + IDL
'End Property

Public Sub ACK(ByVal TimeStamp As Date)
  ACKTime = TimeStamp
End Sub

Public Property Let packetbytes(b() As Byte)
  Dim s As String
  s = StrConv(b, vbUnicode)
End Property

Public Property Let packet(ByVal Value As String)
  mPacket = left$(Value, 255)
End Property
Public Property Get packet() As String
  packet = mPacket
End Property

Private Sub Class_Initialize()
  'Set mLocations = New Collection
  Set Consoles = New Collection
  Set mFirstHops = New Collection
  DateTime = Now
End Sub

Public Property Get LastRestore() As Date
  LastRestore = mLastRestore
End Property

Public Property Let LastRestore(ByVal LastRestore As Date)
  mLastRestore = LastRestore
End Property

Public Property Get ReadyToPost() As Boolean
  ReadyToPost = mReadyToPost
End Property

Public Property Let ReadyToPost(ByVal ReadyToPost As Boolean)
  mReadyToPost = ReadyToPost
End Property

Public Property Get CancelText() As String
  CancelText = mCancelText
End Property

Public Property Let CancelText(ByVal CancelText As String)
  mCancelText = CancelText
End Property

Public Property Get SendCancel() As Integer
  SendCancel = mSendCancel
End Property

Public Property Let SendCancel(ByVal SendCancel As Integer)
  mSendCancel = SendCancel
End Property
Public Property Get AlarmID() As Long
  AlarmID = mAlarmID
End Property

Public Property Let AlarmID(ByVal AlarmID As Long)
  mAlarmID = AlarmID
End Property

Public Property Get Posted() As Boolean
  Posted = mPosted
End Property

Public Property Let Posted(ByVal Value As Boolean)
  dbgloc "Alarm Posting " & Value
  mPosted = Value
End Property

Public Property Get Repeater1() As String
  Repeater1 = mRepeater1
End Property

Public Property Let Repeater1(ByVal Repeater1 As String)
  mRepeater1 = Repeater1
End Property

Public Property Get Signal1() As Double
  Signal1 = mSignal1
End Property

Public Property Let Signal1(ByVal Signal1 As Double)
  mSignal1 = Signal1
End Property


Public Property Get ID() As Long

  ID = MID

End Property

Public Property Let ID(ByVal ID As Long)

  MID = ID

End Property

Public Property Get Responder() As String

  Responder = mResponder

End Property

Public Property Let Responder(ByVal Responder As String)

  mResponder = Responder

End Property
