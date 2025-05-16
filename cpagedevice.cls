VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPageDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private inproc           As Boolean

Private TAP2_PROGRESS    As Long  '
Private TAP2_PROGRESS_SENT_CR As Long
Const TAP2_PROGRESS_NOTSTARTED = 0
Const TAP2_PROGRESS_STARTED = 1
Const TAP2_PROGRESS_ID_RECEIVED = 2  ' RECEIVED 'ID=
Const TAP2_PROGRESS_BRACKET_P_RECIEVED = 3  ' RECEIVED "<ESC>[P<CR>"
Const TAP2_PROGRESS_ACK_RECIEVED = 4  ' RECEIVED "<ESC>[P<CR>"

Private TAP2ResponseTimeout As Date

Private TAP2message      As String
Private TAP2_ProgressWaitFor As Long  ' seconds

Public Que               As Collection
Private CurrentItem      As cPageItem

Private UseKeepAlive     As Boolean
Private NoDataCount      As Long
Const MaxNoData          As Long = 3

Private PingWaiting      As Boolean
Private ResponseWaiting  As Boolean

Private NextPing         As Date  ' for pinging against Keltron SDACT2


Private Const PING_DELAY As Long = 10  ' wait at least 10 seconds
Private mSDACTError      As Long


Private Const SDACT2_PINGMESSAGE = "@" & vbCr

' For TAP_IP (and probably TAP) , if we monitor for "ID=" ...
' If we're not connected,

' set last error on AckError
' clear error on ACK Good


' bits maybe

Const SERVER_STATUS_UNINIT = -1
Const SERVER_STATUS_CLOSED = 0
Const SERVER_STATUS_CONNECTED = 1  ' bit 0
Const SERVER_STATUS_DATALIVE = 2  ' must be this level to send, we've recieved "ID="
Const SERVER_STATUS_DATAMISSING = 4  ' bit 2 ' keep alive fail
Const SERVER_STATUS_ACKMISSING = 8  ' bit 3
Const SERVER_STATUS_PORTERROR = 16  ' bit 5
Const SERVER_STATUS_DISCONNECTED = 32  ' bit 5

Const SERVER_STATUS_GENERAL_ERROR = 64  ' bit 6
Const SERVER_STATUS_TIMOUT = 256  ' bit 8


Dim TAP_IP_Error         As Long  ' yes or no
Dim TAP_Error            As Long  ' yes or no

Dim TAP2_Error           As Long

Dim TAP_Status           As Long
Dim TAP_LastError        As Long
Dim TapWaiting           As Boolean

Dim SDACT_ERROR          As Long
Dim PingDue              As Double
Dim DeviceSerial         As String

'Global Const PROTOCOL_NONE  As Integer = 0
'Global Const PROTOCOL_TAP   As Integer = 1
'Global Const PROTOCOL_COMP1 As Integer = 2
'Global Const PROTOCOL_COMP2 As Integer = 3
'Global Const PROTOCOL_TTS    As Integer = 4
'Global Const PROTOCOL_DIALER  As Integer = 6

'Global Const PROTOCOL_TAP_IP      As Integer = 14

' Winsock state
'sckClosed              0 Closed (Default)
'sckOpen                1 Open
'sckListening           2 Listening
'sckConnectionPending   3 Connection pending
'sckResolvingHost       4 Resolving host
'sckHostResolved        5 Host resolved
'sckConnecting          6 Connecting
'sckConnected           7 Connected
'sckClosing             8 Peer is closing the connection
'sckError               9 Error


'Private Const SET_RELAY = "SK"
'Private Const RESET_RELAY = "RK"


' Must send a heartbeat message from the host at least once every 60 Seconds
' The heartbeat message format is:
' @<CR>

' Whether or not port supervision is enabled
' the SDACT-2 will respond to the heartbeat message with a single character between ASCII '@' and ASCII 'O'
' i.e. between 0x40 and 0x4F. The low four bits encode the four
' SDACT-2 dialer communication status values as:

' Bit 0: 0 = Phone Line 1 OK
' Bit 0: 1 = Phone Line 1 Failure
' Bit 1: 0 = Phone Line 2 OK
' Bit 1: 1 = Phone Line 2 Failure
' Bit 2: 0 = Primary phone number OK
' Bit 2: 1 = Communication to Primary phone number Failure
' Bit 4: 0 = Second phone number OK
' Bit 4: 1 = Communication to Second phone number Failure
' For example

' '@' (0x40) indicates the SDACT-2 is all normal
' 'B' (0x42) indicates a Phone Line 2 failure.
' Note that each of the four faults above also activates the SDACT-2 trouble relay
' which can be used as an alternative to provide local annunciation of Keltron SDACT-2 dialer faults

Const HEARTBEAT_DELAY = 20     ' seconds


Private mRelay1          As Integer
Private mRelay2          As Integer
Private mRelay3          As Integer
Private mRelay4          As Integer
Private mRelay5          As Integer
Private mRelay6          As Integer
Private mRelay7          As Integer
Private mRelay8          As Integer

Private mLastRelayID     As Long

Public Serial            As String

Private KeepAliveTimeout As Date  ' we should get a ping back every second

Private NextRetry        As Date
Private IgnoreMe         As Boolean

Private ServerStatus     As Long  ' 0 is OK, non 0 = error
Private WSock            As Winsock
Attribute WSock.VB_VarHelpID = -1

Private WithEvents TTS   As spvoice
Attribute TTS.VB_VarHelpID = -1
Private PCA              As cPCA

Public ADU200            As cADU200

Private Comm             As cComm

Private ExtPageDevice    As cADU200

Private Const TTS_DEFAULTPAUSE = 2  ' seconds


Private RetriesRemaining As Long

'Private Declare Function OpenAduDevice Lib "AduHid.DLL" (ByVal iTimeout As Long) As Long
'Private Declare Function WriteAduDevice Lib "AduHid.DLL" (ByVal aduHandle As Long, ByVal lpBuffer As String, ByVal lNumberOfBytesToWrite As Long, ByRef lBytesWritten As Long, ByVal iTimeout As Long) As Long
'Private Declare Function ReadAduDevice Lib "AduHid.DLL" (ByVal aduHandle As Long, ByVal lpBuffer As String, ByVal lNumberOfBytesToRead As Long, ByRef lBytesRead As Long, ByVal iTimeout As Long) As Long
'Private Declare Function CloseAduDevice Lib "AduHid.DLL" (ByVal iHandle As Long) As Long

'Private Const ADU_NO_TIMEOUTS = 0
'Private Const ADU_USE_TIMEOUTS = 1

Public ErrorStatus       As Long

Public TIMEOUTDELAY      As Long  ' device fail in seconds
Private Const KEEPALAIVE_DELAY As Long = 10  ' device fail in seconds

Private mDeviceID        As Long  ' which one in the system

Public RemoteAddress     As String


Public Description       As String

Private mPort            As Integer
Public BaudRate          As String
Public BITS              As String
Public Parity            As String
Public Stopbits          As String
Public Settings          As String  ' also will be for audio device

Public MarquisControlCode As Integer


Public AudioDevice       As String


'*************** AUTO DIALER ************
Private WithEvents Dialer As cTapiCall
Attribute Dialer.VB_VarHelpID = -1


Private mDialerModem     As Long
Private mDialerPhone     As String
Public DialerTag         As String
Public DialerMsgDelay    As Long
Public DialerMsgRepeats  As Long
Public DialerMsgSpacing  As Long
Public DialerRedials     As Long
Public DialerRedialDelay As Long
Public DialerVoice       As String
Public DialerTimeout     As Long

Private mKeepPaging      As Long


Private mDialerTerminateDigit As Long

Private WithEvents DivaDialer As cDivaCall
Attribute DivaDialer.VB_VarHelpID = -1



'*********** END AUTO DIALER ************

' with TTS engine, optionally allow keying of microphone/trnasmitter
' Need to insert a pause/silence into audio stream

Public PASystemKey       As Integer  ' 0 or 1
Public PASystemPort      As Integer
Public PASystemHandle    As Long
Public PARepeatTwice     As Integer




Public Checked           As Boolean  ' used to update system live
Public NoWait            As Integer  ' use default wait between TTS sends

Public Pause             As Long  'pause in seconds: how fast can this device can send pages?

Public LFeeds            As Long

Private NUL              As String
Private SOH              As String
Private STX              As String
Private ETX              As String
Private EOT              As String
Private ENQ              As String
Private ACK              As String
Private BEL              As String
Private BS               As String
Private HT               As String
Private VT               As String
Private FF               As String
Private CR               As String
Private SO               As String
Private si               As String
Private LF               As String
Private XOFF             As String
Private XON              As String
Private NAK              As String
Private ETB              As String
Private SUBS             As String
Private ESC              As String
Private rs               As String
Private US               As String
Private DEL              As String

Private SUB_CR           As String
Private SUB_LF           As String




Private NextSend         As Date
Private mBusy            As Boolean
Private mProtocolID      As Integer

Private ACKFailTime      As Date

Private VoiceStart       As Date

Private rxBuffer         As String

Public Function ToESPacket() As cESPacket
  Dim packet             As cESPacket
  Set packet = New cESPacket

  packet.DateTime = Now
  packet.alarm = 1
  packet.PacketType = 1
  packet.SerialPacket = 0

  packet.Serial = Right$("00000000" & Serial, 8)

  '  packet.SetMIDClassPTI (&HD0 + DukeType), DukeType, &HDD

  Set ToESPacket = packet

  Set packet = Nothing


End Function




'create error
'      Case &H41 ' line 1 fail
'      Case &H42 ' line 2 fail
'      Case &H43 ' line 1 and line 2 fail
'      Case &H44 ' primary number fail
'      Case &H45 ' line 1 and primary number fail
'      Case &H46 ' primary number fail and secondary number fail
'      Case &H47 ' line 1 and primary number fail and line 2 fail
'      Case &H48 ' secondary number fail
'      Case &H49 ' line 1 fail and secondary number fail
'      Case &H4A ' line 2 fail and secondary number fail
'      Case &H4B ' line 1 fail and line 2 fail and secondary number fail
'      Case &H4C ' primary number fail and secondary number fail
'      Case &H4D ' line 1 fail and primary number fail and secondary number fail
'      Case &H4E ' primary number fail and line 2 fail and secondary number fail
'      Case &H4F ' line 1 fail and primary number fail and line 2 fail and secondary number fail

Public Function CheckPingReponse() As Long
  Dim packet             As cESPacket
  Dim Response           As String
  Dim d                  As cESDevice

  Set d = Devices.Device(AudioDevice)

  If d Is Nothing Then
    Exit Function
  End If

  If CDbl(d.LastAlarm = 0) Then
    d.LastAlarm = Now
  End If

  If Comm.inbuffercount Then

    Response = String$(Comm.inbuffercount, " ")
    Comm.CommRead Response, Len(Response)  'comm.CommFlush

    Select Case Asc(Response)
      Case &H41 To &H4F
        If Not CBool(d.alarm) Then
          Debug.Print Now & " BAD Response"
          Set packet = New cESPacket
          packet.alarm = 1
          packet.PacketType = 0
          packet.SerialPacket = ""
          packet.Serial = Right$("00000000" & AudioDevice, 8)
          packet.MID = &HD6
          packet.PTI = &HD6
          packet.ClassByte = &H7
          packet.Extra = "Code " & Hex(Asc(Response))
          packet.DateTime = Now
          ProcessESPacket packet
          'Debug.Print
          'Debug.Print Now & " CheckPingReponse Error mBusy " & mBusy

        End If

      Case &H40                ' good

        If CBool(d.alarm) Then

          Set packet = New cESPacket
          packet.alarm = 0
          packet.PacketType = 0
          packet.SerialPacket = ""
          packet.Serial = Right$("00000000" & AudioDevice, 8)
          packet.MID = &HD6
          packet.PTI = &HD6
          packet.ClassByte = &H7
          packet.Extra = "Code " & Hex(Asc(Response))
          packet.DateTime = Now
          ProcessESPacket packet
          'Debug.Print
          'Debug.Print Now & " CheckPingReponse Good mBusy " & mBusy

        End If


        ' clear error
    End Select
  Else                         ' no ping returned

    If Not CBool(d.alarm) Then
      Debug.Print Now & " No Response"
      Set packet = New cESPacket
      packet.alarm = 1
      packet.PacketType = 0
      packet.SerialPacket = ""
      packet.Serial = Right$("00000000" & AudioDevice, 8)
      packet.MID = &HD6
      packet.PTI = &HD6
      packet.ClassByte = &H7
      packet.Extra = "Status Dead"
      packet.DateTime = Now
      ProcessESPacket packet
      'Debug.Print
      'Debug.Print Now & " CheckPingReponse No Ping mBusy " & mBusy
    End If

  End If

  NextPing = DateAdd("s", PING_DELAY, Now)

  'Debug.Print "Next Ping at " & NextPing


End Function

Private Function PollSDACT2() As Long

  'DoEvents


10 If CurrentItem Is Nothing Then  ' first time
20  If Que.Count > 0 Then
30    Set CurrentItem = Que(1)
40    Que.Remove 1
50  End If
60 End If



70 On Error GoTo PollSDACT2_Error


80 If inproc Then
90  'Debug.Print Now & " inproc"
100 Exit Function
110 End If

120 inproc = True


130 'Debug.Print Now & " mBusy " & mBusy
140 If mBusy Then

    'If IsPingDue Then
150 If PingWaiting Then
160   PingWaiting = False
170   'Debug.Print Now & " CheckPingReponse"
180   CheckPingReponse
190   mBusy = False


200 ElseIf ResponseWaiting Then
210   ResponseWaiting = False
220   'Debug.Print Now & " CheckHandShake"
230   CheckHandShake
240   mBusy = False
250 Else
260   mBusy = False
270 End If
280 Else                       ' not busy
    '        If Comm.inbuffercount Then ' clear buffer for new send
    '          Dim Temp As String
    '          Temp = String$(Comm.inbuffercount, " ")
    '          Comm.CommRead Temp, Len(Temp)  'comm.CommFlush
    '        End If

290 If IsPingDue Then
300   mBusy = True
310   SendSDACTPing

320   'Debug.Print Now & " SendSDACTPing"
330 Else
340   If NextSend <= Now Then
350     mBusy = True
360     If CurrentItem Is Nothing Then
370       Debug.Print Now & " CurrentItem is nothing"
380     Else
390       Send CurrentItem
          Debug.Print "********************"
400       Debug.Print Now & " Send CurrentItem"
410       Set CurrentItem = Nothing
420     End If
430   End If
440 End If

450 End If


PollSDACT2_Resume:
460 inproc = False
470 On Error GoTo 0
480 Exit Function

PollSDACT2_Error:

490 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.PollSDACT2." & Erl
500 Resume PollSDACT2_Resume

End Function


Private Function IsPingDue() As Boolean

  Debug.Print "Now, Next Ping "; Now, NextPing

  If NextPing <= Now Then
    IsPingDue = True
  End If
End Function

Private Function SendSDACTPing() As Boolean
  Dim rc                 As Long
  Dim Buffer             As String
  ' make sure buffer is empty
  If Comm.inbuffercount Then
    rc = Comm.CommRead(Buffer, Comm.inbuffercount)
  End If
  Comm.CommWrite SDACT2_PINGMESSAGE
  PingWaiting = True
End Function


Private Function GetSDACT2Response() As Long
  Static inproc          As Boolean
  Dim doeventsloops      As Long

  Dim rc                 As Long
  Dim Buffer             As String

10 On Error GoTo GetSDACT2Response_Error

20 If inproc Then
30  Exit Function
40 End If
50 inproc = True


60 Do

70  If Comm.inbuffercount Then
80    rc = Comm.CommRead(Buffer, Comm.inbuffercount)
90    rxBuffer = rxBuffer & Buffer
100   If InStr(rxBuffer, ACK) Then
        '        LogPageEvent "ACK OK " & vbTab & ProtocolString(Me.ProtocolID)
        '        Debug.Print "ACK OK " & vbTab & ProtocolString(Me.ProtocolID)
110     CurrentItem.SendComplete = 1

120     SDACT_ERROR = 0
130     NextSend = DateAdd("s", Pause, Now)
140     mBusy = False
150     Exit Do
160   End If
170 End If



180 If Now > ACKFailTime Then
190   LogPageEvent "ACK FAIL TIMEOUT " & vbTab & ProtocolString(Me.ProtocolID)
      'Debug.Print "ACK FAIL " & vbTab & ProtocolString(Me.ProtocolID)
200   NextRetry = DateAdd("s", 30, Now)
210   SDACT_ERROR = 1
      'TAP_LastError = SERVER_STATUS_TIMOUT
      'TAP_Status = SERVER_STATUS_TIMOUT

220   mBusy = False

230   Exit Do
240 End If
250 doeventsloops = doeventsloops + 1
260 If doeventsloops > 100 Then
270   doeventsloops = 0
280   DoEvents
290 End If

300 Loop

310 inproc = False





GetSDACT2Response_Resume:

320 On Error GoTo 0
330 Exit Function

GetSDACT2Response_Error:

340 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.GetSDACT2Response." & Erl
350 Resume GetSDACT2Response_Resume

End Function



Private Function CheckForTTSTimeout()
  Dim SinceStart         As Date
  SinceStart = DateAdd("s", 15, VoiceStart)
  If Now > SinceStart Then
    mBusy = False
  End If


  ' causes intermittant or no voice announce.

  '  If NoWait Then
  '    NextSend = DateAdd("s", 0, Now)  ' pause between messages
  '  Else
  '    NextSend = DateAdd("s", Pause, Now)  ' pause between messages
  '  End If
  '
  '  mBusy = False


End Function

Public Function GetPingResponse() As Boolean
  Dim PingResponseTime   As Double
  Dim starttime          As Date
  PingResponseTime = 1         ' one second to reply
  starttime = DateAdd("s", PingResponseTime, Now)

  Do While starttime > Now
    If Comm.inbuffercount Then
      'Comm.CommFlush
      Exit Do
    End If
    DoEvents
  Loop




End Function

Public Function HasError() As Long


  HasError = TAP_IP_Error Or TAP_Error
End Function

Public Function GetActiveAlarmsByAddress(ByVal Address As String) As Collection
  Dim AlarmList          As Collection
  Dim PageItem           As cPageItem

  Set AlarmList = New Collection

  For Each PageItem In Que
    If 0 = StrComp(PageItem.Address, Address, vbTextCompare) Then
      AlarmList.Add PageItem
    End If
  Next
  Set GetActiveAlarmsByAddress = AlarmList

End Function




Public Function RemoveByAlarmID(ByVal AlarmID As Long) As Long
  Dim PageItem           As cPageItem
  Dim j                  As Long
  On Error Resume Next

  Debug.Print "Removing AlarmID "; AlarmID
  For j = Que.Count To 1 Step -1
    If Que(j).AlarmID = AlarmID Then
      Que.Remove j             ' delete it
    End If
  Next

End Function

Public Function GetQueAlarmIDs(ByVal Subscriber As String) As Long()

  Dim PageItem           As cPageItem
  Dim ids()              As Long
  Dim Count              As Long
10 On Error GoTo GetQueAlarmIDs_Error
If 1 Then
'  Debug.Assert 0
End If

20 ReDim ids(0)
30 For Each PageItem In Que
40  If 0 = StrComp(Subscriber, PageItem.Address, vbTextCompare) Then
50    Count = Count + 1
60    ReDim Preserve ids(Count)
70    ids(Count) = PageItem.AlarmID
80  End If
90 Next
100 GetQueAlarmIDs = ids

GetQueAlarmIDs_Resume:

110 On Error GoTo 0
120 Exit Function

GetQueAlarmIDs_Error:

130 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.GetQueAlarmIDs." & Erl
140 Resume GetQueAlarmIDs_Resume

End Function

Private Sub Dialer_OnDone()
        Dim PageItem           As cPageItem
        Dim j                  As Integer


10      If (Dialer.TerminateDigit <> 0) Then  ' a non-zero is active

20        If (Dialer.DigitReceived) Then  ' and someone pressed the correct one...

30          If (KeepPaging And 1) Then  ' should finish this one, but continue on with other pagers
40            For j = Que.Count To 1 Step -1
50              If Que(j).AlarmID = CurrentItem.AlarmID Then
60                If Que(j).Address = CurrentItem.Address Then
70                  If Not (Que(j).IsCancel) Then
80                    dbgTAPI "Que Item Removed (duplicate)" & CurrentItem.AlarmID & " @ " & CurrentItem.Address
90                    Que.Remove j
100                 End If
110               End If
120             End If
130           Next


140         Else ' if NOT KeepPaging And 1 , Need to kill off all other pages for dialer for this event


              ' remove them if repeats for the same alarm, but keep any cancel announcements
150           For j = Que.Count To 1 Step -1
160             If Que(j).AlarmID = CurrentItem.AlarmID Then
                  'If Que(j).Address = CurrentItem.Address Then
170               If Not (Que(j).IsCancel) Then
180                 dbg "Que Item Removed (duplicate)" & CurrentItem.AlarmID
190                 Que.Remove j
200               'End If
                  End If
210             End If
220           Next

230           For j = gPageRequests.Count To 1 Step -1   ' we'll want to stop accepting page requests for this alarm
240             Set PageItem = gPageRequests.PageItem(j)
250             If PageItem.AlarmID = CurrentItem.AlarmID Then  ' came from same alarm
                  
260               PageItem.ProcessDialerACK  ' tell the pagerequest not to talk to me anymor
                  
                  gPageRequests.Remove j
270             End If
280           Next
290         End If
300       End If
310     End If
320     NextSend = DateAdd("s", Pause, Now)  ' pause between messages
330     mBusy = False
End Sub

Private Function SendOnTrak() As Long

  Do While Que.Count
    ADU200.AddItem Que(1)
    Que.Remove 1
  Loop
  NextSend = Now
  mBusy = False

End Function

Private Function SendDIALOGIC(ByVal PhoneNumber As String, ByVal message As String) As Long


  On Error Resume Next
  Dim t                  As Long
  Dim BytesWritten       As Long
  Dim Buffer             As String

  dbgTAPI "PageDevice.SendDIALOGIC " & PhoneNumber

  message = Replace(message, vbCrLf, " ")

  '  t = Win32.timeGetTime()
  If IsDate(message) Then
    message = ConvertDatetoTTS(message)
  End If

  If Trim(PhoneNumber) = "" Then
    DivaDialer.PhoneNumber = Me.DialerPhone
  Else
    DivaDialer.PhoneNumber = PhoneNumber
  End If
  DivaDialer.message = message
  DivaDialer.BeginCall

  If Err.Number <> 0 Then
    SendDIALOGIC = Err.Number
    NextSend = DateAdd("s", Pause, Now)
    mBusy = False
  End If

End Function



Private Function SendDIALER(ByVal PhoneNumber As String, ByVal message As String) As Long


10 On Error GoTo SendDIALER_Error

  'On Error Resume Next
  '  Dim t As Long
  Dim BytesWritten       As Long
  Dim Buffer             As String

20 Debug.Print "SendDIALER "; PhoneNumber, message

30 message = Replace(message, vbCrLf, " ")

  '  t = Win32.timeGetTime()
40 If IsDate(message) Then
50  message = ConvertDatetoTTS(message)
60 End If
  PhoneNumber = Trim(PhoneNumber)
70 If PhoneNumber = "" Then
80  Dialer.PhoneNumber = Trim(Me.DialerPhone)
85  If Trim(Me.DialerPhone) = "" Then
      NextSend = DateAdd("s", Pause, Now)
      SendDIALER = -99
      mBusy = False
      LogProgramError "Error No Phone Number at cPageDevice.SendDIALER.85"
      Exit Function
    End If
90 Else
100 Dialer.PhoneNumber = PhoneNumber
110 End If
120 Dialer.message = message
130 Dialer.BeginCall

  '  If Err.Number <> 0 Then
  '  End If

SendDIALER_Resume:

140 On Error GoTo 0
150 Exit Function

SendDIALER_Error:

160 SendDIALER = Err.Number
170 NextSend = DateAdd("s", Pause, Now)
180 mBusy = False

190 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendDIALER." & Erl
200 Resume SendDIALER_Resume

End Function
Function Ready() As Boolean
  Select Case mProtocolID


    Case PROTOCOL_TAP2
      Ready = Not mBusy


    Case PROTOCOL_COMP1
      Ready = Not mBusy

    Case PROTOCOL_COMP2
      Ready = Not mBusy

    Case PROTOCOL_TAP
      Ready = Not mBusy

    Case PROTOCOL_TAP_IP
      Ready = Not mBusy

    Case PROTOCOL_TTS          ' TTS Engine
      Ready = Not mBusy

    Case PROTOCOL_EMAIL
      Ready = Not mBusy

    Case PROTOCOL_DIALER
      Ready = (Dialer.Status = 0)

    Case PROTOCOL_DIALOGIC
      If Not DivaDialer Is Nothing Then
        Ready = (DivaDialer.Status = 0)
      End If

    Case PROTOCOL_REMOTE
      Ready = True

    Case PROTOCOL_CENTRAL, PROTOCOL_SDACT2
      Ready = Not mBusy

    Case PROTOCOL_ONTRAK
      Ready = Not mBusy

    Case Else                  'PROTOCOL_NONE
      Ready = Not mBusy
  End Select
End Function



'Auto sends when polled, removes page from this object's que

Public Function Poll() As Boolean

  ' Called every second via CheckPagers which loops thru each pagedevice


  Dim RC2                As Long
  Dim j                  As Long




  If IgnoreMe Then             ' we're in the midst of making a connection
    Exit Function
  End If

  'Debug.Print "Poll ", Now

  On Error GoTo Poll_Error
  'Debug.Print "PageDevice.Poll.Que.Count " & Que.Count
  'Debug.Print "PageDevice.Poll.ready " & Ready

  If CDbl(NextSend) <= 0 Then
    NextSend = DateAdd("s", 1, Now)
  End If

  If CDbl(NextPing) <= 0 Then
    NextPing = DateAdd("s", PING_DELAY, Now)  ' ping delay is typically 2 seconds

  End If

  Select Case ProtocolID
    Case PROTOCOL_SDACT2
      Dim IsOpen         As Boolean
      'isopen = Comm.PortOpen
      If Not Comm.PortOpen Then
        OpenConnection
      End If

      'CheckForSDACT2Timeout
      PollSDACT2

      '      If SDACT2_Error Then
      '        If SDACT2_Error = SDACT2_timeout Then
      '           Comm.CommClose
      '           SDACT2_Error = 0
      '           OpenConnection
      '
      '        End If
      '      Else
      '      End If

      ' CheckForHeartBeat

    Case PROTOCOL_TAP2

      ' need to duplicate , kinda like the above
      If TAP2_Error And 0 Then
        Comm.CommClose
        OpenConnection
      Else
        If Not Comm.PortOpen Then
          OpenConnection
        End If
      End If

      If mBusy Then
        POLL_SENDTAP2
      Else                     ' JUST EMPTY BUFFER
        If Comm.inbuffercount Then
          'Debug.Print "PROTOCOL_TAP2 cPagedevice.Poll InbufferCount (flush) ", Comm.inbuffercount
          Dim Temp       As String
          Temp = String$(Comm.inbuffercount, " ")
          Comm.CommRead Temp, Len(Temp)  'comm.CommFlush
          'Debug.Print " Temp read " & Temp
        End If

      

      If CurrentItem Is Nothing Then  ' first time
        If Que.Count > 0 Then
          Set CurrentItem = Que(1)
          Que.Remove 1
        End If
      End If

      If Not (CurrentItem Is Nothing) Then  ' See if it was sent
        If CurrentItem.SendComplete Then
          Set CurrentItem = Nothing
          If Que.Count > 0 Then
            Set CurrentItem = Que(1)
            Que.Remove 1
          End If
        End If
      End If

      If Not (CurrentItem Is Nothing) Then  ' something maybe to send
        If CurrentItem.SendComplete = 0 Then  ' if it's not sent, let's try
          If Now > NextSend Then  ' we send no page before its time
            Send CurrentItem
          End If
        End If
      End If
      End If
      Exit Function

    Case PROTOCOL_TAP
      ' need to duplicate , kinda like the above
      If TAP_Error Then
        Comm.CommClose
        OpenConnection
      Else
        If Not Comm.PortOpen Then
          OpenConnection
        End If
      End If

      If mBusy Then
        CheckHandShake         ' only used for TAP
      Else

        If Comm.inbuffercount Then
          'Debug.Print "PROTOCOL_TAP cPagedevice.Poll InbufferCount (flush) ", Comm.inbuffercount
          'Dim Temp       As String
          Temp = String$(Comm.inbuffercount, " ")
          Comm.CommRead Temp, Len(Temp)  'comm.CommFlush
          'Debug.Print " Temp read " & Temp
        End If

      End If

      If CurrentItem Is Nothing Then  ' first time
        If Que.Count > 0 Then
          Set CurrentItem = Que(1)
          Que.Remove 1
        End If
      End If

      If Not (CurrentItem Is Nothing) Then  ' See if it was sent
        If CurrentItem.SendComplete Then
          If Que.Count > 0 Then
            Set CurrentItem = Que(1)
            Que.Remove 1
          End If
        End If
      End If

      If Not (CurrentItem Is Nothing) Then  ' something maybe to send
        If CurrentItem.SendComplete = 0 Then  ' if it's not sent, let's try
          If Now > NextSend Then  ' we send no page before its time
            Send CurrentItem
          End If
        End If
      End If
      Exit Function

    Case PROTOCOL_TAP_IP
      If WSock Is Nothing Then
        OpenConnection         ' simple and sweet
      End If

      If TAP_IP_Error Then
        If Now > NextRetry Then
          OpenConnection       ' simple and sweet
        End If
      ElseIf WSock.State <> sckConnected Then
        OpenConnection         ' simple and sweet
      Else
        If mBusy Then          ' sent and waiting for handshake
          CheckHandShake       '
        Else                   ' between sends, see if we have anything to send
          If WSock.BytesReceived Then
            NoDataCount = 0
            PurgeWinsockInBuffer
          Else
            If NoDataCount > 4 Then
              TAP_IP_Error = 1
              Exit Function
            Else
              NoDataCount = NoDataCount + 1
            End If
          End If

          If CurrentItem Is Nothing Then  ' first time
            If Que.Count > 0 Then
              Set CurrentItem = Que(1)
              Que.Remove 1
            End If
          End If

          If Not (CurrentItem Is Nothing) Then  ' See if it was sent
            If CurrentItem.SendComplete Then
              If Que.Count > 0 Then
                Set CurrentItem = Que(1)
                Que.Remove 1
              End If
            End If
          End If

          If Not (CurrentItem Is Nothing) Then  ' something maybe to send
            If CurrentItem.SendComplete = 0 Then  ' if it's not sent, let's try
              If Now > NextSend Then  ' we send no page before its time
                Send CurrentItem
              End If
            End If
          End If

        End If
      End If
      Exit Function

    Case PROTOCOL_MOBILE
      If Que.Count > 0 Then
        Set CurrentItem = Que(1)
        Que.Remove 1
        Send CurrentItem
      End If
      Exit Function            '<<<<<<<<<<<<<<<<<
    Case PROTOCOL_REMOTE
      ' ???????????????? not used here ?????????????
      'Debug.Print "Polling REMOTE PAGE DEVICE ID"; DeviceID
      'Debug.Print "QueCount "; Que.Count

      'For j = 1 To Que.Count
      '  Debug.Print "Alarm "; Que(j).AlarmID
      '  Debug.Print "Address "; Que(j).Address
      'Next

      Exit Function            '<<<<<<<<<<<<<<<<<

    Case PROTOCOL_TTS
      CheckForTTSTimeout



    Case Else                  ' all the other protocols

      If Not Dialer Is Nothing Then
        Dialer.UpdateClock
      End If

      If Not DivaDialer Is Nothing Then
        DivaDialer.UpdateClock
      End If

      If Not ADU200 Is Nothing Then
        ADU200.UpdateClock
      End If

  End Select

  'Debug.Print "Page Device QueCount "; Que.Count
  If Que.Count > 0 Then
    ' check this
    'Debug.Print "Page Device QueCount "; Que.Count
    If Ready Then
      If Now > NextSend Then
        mBusy = True

        If ProtocolID = PROTOCOL_ONTRAK Then
          ' ONTRAK Does its own dequeing
          Set CurrentItem = Que(1)
          Send CurrentItem
        Else
          Set CurrentItem = Que(1)
          RC2 = Send(CurrentItem)
          If 0 = RC2 Then
            Que.Remove 1
          End If
        End If
      End If
    Else
      'CheckHandShake ' only used for TAP and TAP_IP
    End If
  End If


Poll_Resume:
  On Error GoTo 0
  Exit Function

Poll_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.Poll." & Erl
  Resume Poll_Resume


End Function
Private Sub CheckHandShake()

  ' Debug.Print "Check Handshake ", Now

  Select Case mProtocolID

    Case PROTOCOL_MOBILE


    Case PROTOCOL_COMP1
      ' no response codes
    Case PROTOCOL_COMP2
      ' no response codes
    Case PROTOCOL_TAP_IP
      If TapWaiting Then
        GetResponse_IP
      End If

    Case PROTOCOL_SDACT2
      ' handled elsewhere
    Case PROTOCOL_TAP2

      If mBusy Then
        'Debug.Print ""
        'Debug.Print "GetResponse is busy"
        'GetResponse
      Else
        'Debug.Print ""
        'Debug.Print "GetResponse NOT busy"
      End If


    Case PROTOCOL_TAP

      If mBusy Then
        'Debug.Print ""
        'Debug.Print "GetResponse is busy"
        GetResponse
      Else
        'Debug.Print ""
        'Debug.Print "GetResponse NOT busy"
      End If

    Case PROTOCOL_TTS          ' TTS Engine
      ' no response codes
    Case PROTOCOL_EMAIL
      ' MAPI MAIL
    Case PROTOCOL_PCA
      ' PCA Device
    Case PROTOCOL_DIALER
      mBusy = (Dialer.Status = 0)

    Case PROTOCOL_DIALOGIC
      If Not DivaDialer Is Nothing Then
        mBusy = (DivaDialer.Status = 0)
      End If




    Case PROTOCOL_CENTRAL
      ' Central monitoring
    Case PROTOCOL_ONTRAK

    Case Else                  'PROTOCOL_NONE
      ' no response codes
  End Select

End Sub
Function GetResponse_IP() As Long
  ' FOR TAP_IP only

  Dim rc                 As Long
  Dim Buffer             As String


10 On Error GoTo GetResponse_IP_Error

20 If WSock Is Nothing Then Exit Function

  'Debug.Print CDbl(ACKFailTime)
30 Do
    ' if has data then
40  If WSock.BytesReceived > 0 Then  ' clear out old data
50    KeepAliveTimeout = DateAdd("s", KEEPALAIVE_DELAY, Now)
60    WSock.GetData Buffer, vbString
70  End If
80  rxBuffer = rxBuffer & Buffer
90  If InStr(rxBuffer, ACK) Then  ' ************* YEA! ************
100   LogPageEvent "ACK OK " & vbTab & ProtocolString(Me.ProtocolID)
110   Debug.Print "ACK OK " & vbTab & ProtocolString(Me.ProtocolID)
      ' no change in status
120   CurrentItem.SendComplete = 1
130   NextSend = DateAdd("s", Pause, Now)
140   TAP_IP_Error = 0
150   TAP_LastError = 0
160   mBusy = False
170   Exit Do
180 End If

190 If Now > ACKFailTime Then  ' ********** ZONK **************
200   LogPageEvent "ACK FAIL TIMEOUT" & vbTab & ProtocolString(Me.ProtocolID)
210   Debug.Print "ACK FAIL " & vbTab & ProtocolString(Me.ProtocolID)
220   TAP_IP_Error = 1
230   NextRetry = DateAdd("s", 30, Now)
240   TAP_LastError = SERVER_STATUS_TIMOUT
250   TAP_Status = SERVER_STATUS_TIMOUT
260   mBusy = False
270   Exit Do
280 End If
290 DoEvents
300 Loop

GetResponse_IP_Resume:
310 On Error GoTo 0
320 TapWaiting = False
330 Exit Function

GetResponse_IP_Error:

340 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.GetResponse_IP." & Erl
350 TAP_Status = SERVER_STATUS_GENERAL_ERROR
360 TAP_LastError = SERVER_STATUS_GENERAL_ERROR
370 TAP_IP_Error = 1
380 Resume GetResponse_IP_Resume




End Function
Private Function PurgeWinsockInBuffer() As String
  Dim Buffer             As String

  If Not WSock Is Nothing Then
    If WSock.BytesReceived > 0 Then
      WSock.GetData Buffer, vbString
      PurgeWinsockInBuffer = Buffer
    End If
  End If

End Function

Function GetResponseSDACT() As Long

  Static inproc          As Boolean
  Dim doeventsloops      As Long

  Dim rc                 As Long
  Dim Buffer             As String




  On Error GoTo GetResponseSDACT_Error



  If inproc Then
    Exit Function
  End If
  inproc = True

  Dim BufferChar         As Long

  Do

    If Comm.inbuffercount Then
      rc = Comm.CommRead(Buffer, Comm.inbuffercount)
      BufferChar = Asc(Buffer) And &H4F  ' gets ascii of (first if more than one) character

      If BufferChar = &H40 Then
        SDACT_ERROR = 0
        mBusy = False
        Exit Do

      ElseIf BufferChar And &H40 Then
        SDACT_ERROR = 1
        mBusy = False
        Exit Do
      Else
        SDACT_ERROR = SERVER_STATUS_DISCONNECTED
        mBusy = False
        Exit Do
      End If
    End If

    If Now > ACKFailTime Then
      LogPageEvent "ACK FAIL TIMEOUT " & vbTab & ProtocolString(Me.ProtocolID)
      Debug.Print "ACK FAIL " & vbTab & ProtocolString(Me.ProtocolID)
      NextRetry = DateAdd("s", 30, Now)
      'SDACT_ERROR = ACK_FAIL_TIMEOUT
      mBusy = False

      Exit Do
    End If
    doeventsloops = doeventsloops + 1
    If doeventsloops > 100 Then
      doeventsloops = 0
      DoEvents
    End If

  Loop

  inproc = False

GetResponseSDACT_Resume:
  On Error GoTo 0
  Exit Function

GetResponseSDACT_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.GetResponseSDACT." & Erl
  SDACT_ERROR = &HFF
  Resume GetResponseSDACT_Resume


End Function

Function GetResponse() As Long
  ' FOR serial TAP

  Static inproc          As Boolean
  Dim doeventsloops      As Long

  Dim rc                 As Long
  Dim Buffer             As String


10 On Error GoTo GetResponse_Error



20 If inproc Then
30  Exit Function
40 End If
50 inproc = True


60 Do

70  If Comm.inbuffercount Then
80    rc = Comm.CommRead(Buffer, Comm.inbuffercount)
90    rxBuffer = rxBuffer & Buffer
100   If InStr(rxBuffer, ACK) Then
110     LogPageEvent "ACK OK " & vbTab & ProtocolString(Me.ProtocolID)
120     Debug.Print "ACK OK " & vbTab & ProtocolString(Me.ProtocolID)
130     CurrentItem.SendComplete = 1
140     NextSend = DateAdd("s", Pause, Now)
150     TAP_Error = 0

160     mBusy = False
170     Exit Do
180   End If
190 End If
200 If Now > ACKFailTime Then
210   LogPageEvent "ACK FAIL TIMEOUT " & vbTab & ProtocolString(Me.ProtocolID)
220   Debug.Print "ACK FAIL " & vbTab & ProtocolString(Me.ProtocolID)
230   NextRetry = DateAdd("s", 30, Now)
240   TAP_Error = 1
250   TAP_LastError = SERVER_STATUS_TIMOUT
260   TAP_Status = SERVER_STATUS_TIMOUT
270   mBusy = False

280   Exit Do
290 End If
    doeventsloops = doeventsloops + 1
    If doeventsloops > 100 Then
      doeventsloops = 0
      DoEvents
    End If

310 Loop

320 inproc = False

GetResponse_Resume:
330 On Error GoTo 0
340 Exit Function

GetResponse_Error:

350 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.GetResponse." & Erl
360 TAP_Error = 1
370 Resume GetResponse_Resume

End Function

Private Function WaitForWinsockConnect() As Long
  Dim Timeout            As Date
  Dim doeventsloops      As Long

  IgnoreMe = True              ' revent reentry

  Timeout = DateAdd("s", 3, Now)  ' just 3 seconds to make a connection.. should be nearly instant
  DoEvents
  Do Until Timeout < Now
    doeventsloops = doeventsloops + 1
    If doeventsloops > 100 Then
      doeventsloops = 0
      DoEvents
    End If
    If WSock.State = sckConnected Then
      WaitForWinsockConnect = 1  ' success
      Exit Do
    End If

  Loop

  IgnoreMe = False

End Function

Public Function OpenConnection() As Boolean
  Dim rc                 As Long
  On Error GoTo OpenConnection_Error

  Select Case mProtocolID

    Case PROTOCOL_SDACT2

      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc


    Case PROTOCOL_TAP_IP
      If WSock Is Nothing Then
        Set WSock = frmTimer.GetGenericWinsock
      End If
      If WSock.State <> sckClosed Then
        WSock.Close
      End If

      WSock.Connect Me.DialerPhone, Port  ' ***** dialerphone is IP Address *****!!!!!

      WaitForWinsockConnect

      If WSock.State = sckConnected Then
        TAP_Status = SERVER_STATUS_CONNECTED
        TAP_IP_Error = 0
      Else
        TAP_Status = SERVER_STATUS_TIMOUT
        NextRetry = DateAdd("s", 30, Now)
        TAP_IP_Error = 1
      End If

      NoDataCount = 0
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & TAP_Status



    Case PROTOCOL_TAP
      mBusy = False
      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      If rc = 0 Then
        Comm.CommWrite "ATI" & CR
        Debug.Print "TAP OPEN Send  ATI  CR"
        LogPageEvent "TAP OPEN PORT OK"

        TAP_Error = 0
      Else
        'Debug.Print "TAP OPEN PORT FAIL"
        LogPageEvent "TAP OPEN PORT FAIL"
        TAP_Error = 1
      End If

    Case PROTOCOL_TAP2
      mBusy = False
      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      If rc = 0 Then
        
        Debug.Print "TAP2 OPEN"
        LogPageEvent "TAP2 OPEN PORT OK"

        TAP2_Error = 0
      Else
        'Debug.Print "TAP OPEN PORT FAIL"
        LogPageEvent "TAP OPEN PORT FAIL"
        TAP2_Error = 1
      End If



    Case PROTOCOL_COMP1
      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc

    Case PROTOCOL_COMP2
      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc

    Case PROTOCOL_CENTRAL
      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc

    Case PROTOCOL_TTS          ' TTS Engine
      rc = SendTTS(GetSystemTimeString)  ' send tts takes care of operating the relay
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc
    Case PROTOCOL_EMAIL
      rc = 0


    Case PROTOCOL_PCA
      ' PCA Device
      rc = 0

    Case PROTOCOL_DIALOGIC
      Set DivaDialer = DialogicSystem.GetCallHandlerByID(Me.DialerModem)
      SpecialLog "PageDevice.OpenConnection PROTOCOL_DIALOGIC"

      If Not DivaDialer Is Nothing Then
        SpecialLog "PageDevice.OpenConnection DivaDialer is DivaDialer"
        DivaDialer.Pretext = Me.DialerTag  ' prepended to message
        DivaDialer.PhoneNumber = Me.DialerPhone
        ' redial
        DivaDialer.RedialAttempts = Me.DialerRedials
        DivaDialer.RedialWait = Me.DialerRedialDelay
        'message settings
        DivaDialer.MessageDelay = Me.DialerMsgDelay
        DivaDialer.MessageRepeats = Me.DialerMsgRepeats
        DivaDialer.MessageRepeatDelay = Me.DialerMsgSpacing

        DivaDialer.VoiceName = Me.DialerVoice


        DivaDialer.TerminateDigit = Me.DialerTerminateDigit
        DivaDialer.Timeout = 45
        ' do this last?
        rc = DivaDialer.Active
      Else
        SpecialLog "PageDevice.OpenConnection DivaDialer is NOTHING"
      End If
      SpecialLog "PageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc

    Case PROTOCOL_DIALER

      Dialer.Pretext = Me.DialerTag  ' prepended to message
      Dialer.PhoneNumber = Me.DialerPhone
      ' redial
      Dialer.RedialAttempts = Me.DialerRedials
      Dialer.RedialWait = Me.DialerRedialDelay
      'message settings
      Dialer.MessageDelay = Me.DialerMsgDelay
      Dialer.MessageRepeats = Me.DialerMsgRepeats
      Dialer.MessageRepeatDelay = Me.DialerMsgSpacing

      Dialer.VoiceName = Me.DialerVoice
      Dialer.PermanentLineID = Me.DialerModem  ' Me.DeviceID  ' do this last? DialerModem value ???

      Dialer.TerminateDigit = Me.DialerTerminateDigit
      Dialer.Timeout = 45

      rc = Dialer.Active
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc


    Case PROTOCOL_ONTRAK
      ' nothing to do
    Case PROTOCOL_MOBILE

    Case Else                  'PROTOCOL_NONE, but use serial port
      If Comm.PortOpen Then
        Comm.CommClose
      End If
      rc = Comm.CommOpen(Port, Settings)
      Debug.Print "cPageDevice.OpenConnection." & ProtocolString(mProtocolID) & " " & rc

  End Select
  ErrorStatus = rc

OpenConnection_Resume:
  On Error GoTo 0
  Exit Function

OpenConnection_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.OpenConnection." & Erl
  Resume OpenConnection_Resume

End Function

Public Function CloseConnection() As Boolean
  On Error Resume Next
  Select Case mProtocolID
    Case PROTOCOL_COMP1
      CloseComm
    Case PROTOCOL_COMP2
      CloseComm
    Case PROTOCOL_TAP
      CloseComm
    Case PROTOCOL_TAP2
      CloseComm
    Case PROTOCOL_SDACT2
      CloseComm

    Case PROTOCOL_TAP_IP
      If Not WSock Is Nothing Then
        WSock.Close
        TAP_IP_Error = 0
      End If

    Case PROTOCOL_CENTRAL
      CloseComm
    Case PROTOCOL_TTS          ' TTS Engine
      ' no connection exists
    Case PROTOCOL_EMAIL
      ' handled in send
    Case PROTOCOL_PCA
      ' PCA Device
    Case PROTOCOL_DIALER
      ' cTapicCall Device
    Case PROTOCOL_DIALOGIC
    Case PROTOCOL_ONTRAK
    Case PROTOCOL_MOBILE

    Case Else                  'PROTOCOL_NONE
      CloseComm
  End Select

End Function


Function Send(Item As cPageItem) As Long
  Dim rc                 As Long


  dbg "Send ID/message " & ProtocolString(mProtocolID) & ": " & Item.Address & "/" & Item.message
  'Debug.Print "Send ID/message " & ProtocolString(mProtocolID) & ": " & item.Address & "/" & item.message
  On Error GoTo Send_Error

  Select Case mProtocolID



    Case PROTOCOL_SDACT2
      SendSDACT2 "", Item.RoomText, Item.message  ' pagerid is not used

    Case PROTOCOL_MOBILE
      rc = SendMobile(Item)

    Case PROTOCOL_MARQUIS
      rc = SendMarquis(Item.message)  ' no address required

    Case PROTOCOL_APOLLO
      rc = SendApollo(Item.message)  ' no address required

    Case PROTOCOL_COMP1
      rc = SendComp1(Item.message)  ' no address required

    Case PROTOCOL_COMP2
      rc = SendComp2(Item.Address, Item.message)

    Case PROTOCOL_TAP
      rc = SendTAP(Item.Address, Item.message)

    Case PROTOCOL_TAP2
      rc = SendTAP2(Item.Address, Item.message)


    Case PROTOCOL_TAP_IP
      rc = SendTAP_IP(Item.Address, Item.message)

    Case PROTOCOL_CENTRAL
      rc = SendCentral("000", Item.RoomText, Item.message)

    Case PROTOCOL_TTS          ' TTS Engine
      NoWait = Item.NoWait
      rc = SendTTS(Item.message, Val(Item.RelayID))

    Case PROTOCOL_EMAIL
      rc = SendMAPI2(Item.Address, Item.message)

    Case PROTOCOL_PCA
      rc = SendPCA(Item.Address, Item.message)

    Case PROTOCOL_DIALER
      rc = SendDIALER(Item.Address, Item.message)

    Case PROTOCOL_DIALOGIC
      dbgTAPI "PageDevice.Send SendDIALOGIC"
      rc = SendDIALOGIC(Item.Address, Item.message)

    Case PROTOCOL_ONTRAK
      rc = SendOnTrak()

    Case PROTOCOL_REMOTE
      ' do nothing, it is polled by system
    Case Else                  'PROTOCOL_NONE
      rc = SendComp1(Item.message, False)  ' no address, no CR required
  End Select
  Send = rc

Send_Resume:
  On Error GoTo 0
  Exit Function

Send_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.Send." & Erl
  Resume Send_Resume

End Function
Public Property Get ProtocolID() As Long
  ProtocolID = mProtocolID
End Property
Public Property Let ProtocolID(ByVal Value As Long)

  mProtocolID = Value
  Select Case mProtocolID

    Case PROTOCOL_MOBILE

    Case PROTOCOL_COMP1
      Set Comm = New cComm
    Case PROTOCOL_COMP2
      Set Comm = New cComm
    Case PROTOCOL_TAP
      Set Comm = New cComm

    Case PROTOCOL_TAP2
      Set Comm = New cComm

    Case PROTOCOL_TAP_IP
      If WSock Is Nothing Then
        Set WSock = frmTimer.GetGenericWinsock
      End If

    Case PROTOCOL_CENTRAL
      Set Comm = New cComm

    Case PROTOCOL_TTS          ' TTS Engine
      Set TTS = New spvoice
    Case PROTOCOL_EMAIL
      ' handled in send
      'case PROTOCOL_WEB
      'set XMLHTTP = new
    Case PROTOCOL_ONTRAK
      Set ADU200 = New cADU200

    Case PROTOCOL_DIALER
      Set Dialer = New cTapiCall

    Case PROTOCOL_DIALOGIC
      Set DivaDialer = New cDivaCall
      dbgTAPI "pagedevice.ProtocolID DivaDialer=New DivaCall " & IIf(DivaDialer Is Nothing, "Fail", "DivaDialer")
    Case Else                  'PROTOCOL_NONE
      Set Comm = New cComm
  End Select

End Property
Public Property Get protocol() As String
  protocol = ProtocolString(mProtocolID)
End Property

Public Function AddPage(PageItem As cPageItem)



  Dim DevicePageItem                  As cPageItem
  
10 On Error GoTo AddPage_Error

'20 Set PageItem = New cPageItem
25 'p.group = PageItem.group

20 Set DevicePageItem = New cPageItem

30 DevicePageItem.AlarmID = PageItem.AlarmID  ' so we can remove them upon end of alarm
40 DevicePageItem.message = PageItem.message  ' the message
50 DevicePageItem.Address = PageItem.Address  ' pager address
60 DevicePageItem.NoWait = PageItem.NoWait  ' For TTS, send next one immediatly
70 DevicePageItem.RoomText = PageItem.RoomText  ' Room Text for Central Office
80 DevicePageItem.IsCancel = PageItem.IsCancel
90 DevicePageItem.RelayID = PageItem.RelayID
95 DevicePageItem.AssistRequest = PageItem.AssistRequest
  ' only used here in cPageDevice
100 DevicePageItem.TimeAdded = Now
110 DevicePageItem.SendComplete = 0

120 Que.Add DevicePageItem

AddPage_Resume:

130 On Error GoTo 0
140 Exit Function

AddPage_Error:

150 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.AddPage." & Erl
160 Resume AddPage_Resume
End Function

'Function SendMAPI(ByVal Address As String, ByVal Message As String) As Long
'         Dim mapi As cMAPI
'10       On Error GoTo SendMAPI_Error
'
'20        Set mapi = New cMAPI
'          mapi.Username = "tomk@anonabay.com"
'          mapi.Password = "mypassword"
'30        mapi.Logon
'40        SendMAPI = mapi.Send(Address, "Sentry Freedom", Message)
'50        Set mapi = Nothing
'
'
'SendMAPI_Resume:
'60       On Error GoTo 0
'70       Exit Function
'
'SendMAPI_Error:
'
'80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendMAPI." & Erl
'90      Resume SendMAPI_Resume
'
'
'End Function
Function SendMAPI2(ByVal Address As String, ByVal message As String) As Long
  Dim mapi               As Object
  Dim s                  As String

10 On Error Resume Next
  #If brookdale Then
20  s = "TechConnect"
  #ElseIf esco Then
30  s = "CareConnect"
  #Else
40  s = "Sentry"
  #End If


50 If (Configuration.UseSMTP = MAIL_SMTP) Then
60  If gSMTPMailer Is Nothing Then
      '70          Set gSMTPMailer = New SendMail
70    Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
80  End If
90  If gSMTPMailer Is Nothing Then
100   LogProgramError "Could not create SMTPMailer Object in CpageDevice.SENDMAPI2." & Erl
110 Else
120   Call gSMTPMailer.Send("", "", Address, s & " E-Call Notification", message, "")
130 End If

140 Else

150 Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
160 If mapi Is Nothing Then
170   LogProgramError "Could not create SENTRYMAIL Object in CpageDevice.SENDMAPI2." & Erl
180 Else
190   Call mapi.Send("", "", Address, s & " E-Call Notification", message)
200 End If
210 End If

220 Set mapi = Nothing
230 mBusy = False
240 Exit Function

  'SendMAPI2_Error:
  '
  '  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendMAPI2." & Erl
  '  Resume SendMAPI2_Resume


End Function



Function SendMobile(Item As cPageItem) As Long

  Dim s                  As String
  Dim values             As String
  Dim SQL                As String
  Dim message            As String
  Dim rs                 As ADODB.Recordset
  Dim AckUser            As String
  Dim ACKTime            As String
  Dim Group              As String


  On Error GoTo SendMobile_Error

  If Item.AlarmID = 0 Then
    NextSend = Now
    mBusy = False
    Exit Function
  End If
    

  If InStr(1, message, "cancel", vbTextCompare) > 0 Then
    SQL = "DELETE FROM mobile WHERE AlarmID = " & Item.AlarmID
  Else

    SQL = "SELECT ackuser, eacktime FROM mobile WHERE AckUser > '' and AlarmID = " & Item.AlarmID

    Set rs = ConnExecute(SQL)
    If Not rs.EOF Then
      AckUser = rs("ackuser") & ""
      ACKTime = rs("eAckTime") & ""
    End If
    rs.Close
    Set rs = Nothing


    SQL = "SELECT * FROM mobile WHERE " & Parentheses("AlarmID = " & Item.AlarmID) & " and " & Parentheses("pagerID = " & Item.Address)
    Set rs = ConnExecute(SQL)
    If Not rs.EOF Then
      ' update
      'SQL = "update mobile set ackuser = " & AckUser & ", acktime = " & Acktime & ,

    Else
      ' insert
      
      Item.EventType = Item.EventType
      Item.Phone = Item.Phone
      Item.ResidentText = Item.ResidentText
      s = Item.TimeAdded
      message = Item.RoomText & " " & Replace(Item.message, vbCrLf, " ")
      message = StripWhiteSpace(message)

      values = Item.AlarmID & "," & q(Item.Address) & "," & q(message) & "," & _
               DateDelim & Now & DateDelim & "," & q(ACKTime) & "," & q(AckUser) & "," & _
               "0" & "," & q(Item.TimeAdded) & "," & q(ACKTime) & ",0"

      SQL = "Insert Into mobile ( AlarmID,PagerID,message,TimeAdded,TimeAcked,AckUser,sent,eAlarmTime,eAckTime,Ended ) values ( " & values & " )"
    End If

    rs.Close
    Set rs = Nothing

  End If

  ConnExecute SQL

  NextSend = Now
  mBusy = False


SendMobile_Resume:
  On Error GoTo 0
  Exit Function

SendMobile_Error:
  SendMobile = -1
  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendMobile." & Erl
  Resume SendMobile_Resume


End Function

Function SendCentral(ByVal PagerID As String, ByVal RoomNumber As String, ByVal message As String) As Long
  'rc = SendCentral(Item.RoomText, Item.Message)
  Dim s                  As String
  ' same format as COMP2
  ' PagerID<CR>Message<CR>
  ' to resync <DEL><DEL><DEL>
  ' response ID<CR><CR><RS><CR> if error (device not listed in TX database)
10 On Error GoTo SendCentral_Error
  message = Replace(message, vbCrLf, " ")
20 rxBuffer = ""
30 Comm.CommFlush
  ' pager id is not used
  PagerID = "0"
40 s = PagerID & CR & "#" & CStr(Val(RoomNumber)) & " " & message & CR
  Debug.Print "cPageDevice.SendCentral " & Replace(s, vbCr, "^")
  's = Replace(s, vbCr, "^") & vbCrLf
50 SendCentral = Comm.CommWrite(s)

  LogPageEvent "SEND CO " & Replace(s, vbCr, "^")

60 NextSend = DateAdd("s", Pause, Now)
70 ACKFailTime = DateAdd("s", 10, Now)  ' must receive an ACK within 10 seconds
80 mBusy = False

SendCentral_Resume:
90 On Error GoTo 0
100 Exit Function

SendCentral_Error:

110 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendCentral." & Erl
120 Resume SendCentral_Resume

End Function

Function PingSDACT2()
20 rxBuffer = ""
30 Comm.CommFlush
50 PingSDACT2 = Comm.CommWrite("@" & vbCr)

  '' get return of chr$(&h40) to chr$(&h4F)
  'Bit 0: 1 = Phone Line 1 Failure, 0 = Phone Line 1 OK
  'Bit 1: 1 = Phone Line 2 Failure, 0 = Phone Line 2 OK
  'Bit 2: 1 = Communication to Primary phone number Failure, 0 = Primary phone number OK
  'Bit 4: 1 = Communication to Second phone number Failure, 0 = Second phone number OK



End Function

Function SendSDACT2(ByVal PagerID As String, ByVal RoomNumber As String, ByVal message As String) As Long

  Dim s                  As String
  ' same format as COMP2
  ' PagerID<CR>Message<CR>
  ' to resync <DEL><DEL><DEL>
  ' response ID<CR><CR><RS><CR> if error (device not listed in TX database)
10 On Error GoTo SendSDACT2_Error

20 On Error GoTo SendSDACT2_Error
30 message = Replace(message, vbCrLf, " ")
40 rxBuffer = ""
50 Comm.CommFlush
  ' pager id is not used
60 PagerID = "0"
70 s = PagerID & CR & "#" & CStr(Val(RoomNumber)) & " " & message & CR
80 Debug.Print "cPageDevice.SendCentral " & Replace(s, vbCr, "^")

90 SendSDACT2 = Comm.CommWrite(s)

100 LogPageEvent "SEND SDACT2 " & Replace(s, vbCr, "^")


120 ACKFailTime = DateAdd("s", 10, Now)  ' must receive an ACK within 10 seconds
  ResponseWaiting = True

SendSDACT2_Resume:

130 NextSend = DateAdd("s", Pause, Now)
140 On Error GoTo 0
150 Exit Function

SendSDACT2_Error:

  mBusy = False
160 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendSDACT2." & Erl
170 Resume SendSDACT2_Resume

End Function



'default for ALPHA pagers is 9600 N81

Function SendTAP_IP(ByVal PagerID As String, ByVal message As String) As Long
  Dim cmess              As String

  'TAP
  '<STX>ID<CR>Message<CR><ETX>Checkum<CR>
  ' a customer wants <STX>pagerid<CR>message<ETX>Checkum<CR><LF><LF>.

10 If WSock.State <> sckConnected Then
20  Debug.Print "Socket State " & WSock.State & " SendTAP_IP"
30  Exit Function
40 End If

50 On Error GoTo SendTAP_IP_Error

60 mBusy = True

70 TapWaiting = True


80 rxBuffer = ""

90 message = TAPClean(message)
100 message = STX & PagerID & CR & message & CR & ETX
110 If LFeeds > 0 Then
120 message = message & TAPCheckSum(message) & CR & String$(LFeeds, 10)  ' added variable # of LF chars  4/8/13
130 Else
140 message = message & TAPCheckSum(message) & CR
150 End If

160 WSock.SendData message

170 ACKFailTime = DateAdd("s", 5, Now)  ' must receive an ACK within 10 seconds , probably can shorten this


180 cmess = CleanTAPMessage(message)
190 Debug.Print "SendTAPIP ID/message "; cmess, Now, ACKFailTime
200 LogPageEvent "SEND TAP IP" & vbTab & cmess

210 SendTAP_IP = 1

SendTAP_IP_Resume:
220 On Error GoTo 0

230 Exit Function

SendTAP_IP_Error:

240 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendTAP_IP." & Erl
250 TapWaiting = False

260 mBusy = False
  TAP_IP_Error = 1
270 SendTAP_IP = 0
280 Resume SendTAP_IP_Resume



End Function

Function POLL_SENDTAP2()       ' every second
  Dim Response           As String

  If Now > TAP2ResponseTimeout Then
    TAP2ResponseTimeout = DateAdd("s", 10, Now)
    TAP2_PROGRESS = TAP2_PROGRESS_STARTED
  End If


  Select Case TAP2_PROGRESS
    
    Case TAP2_PROGRESS_STARTED
      rxBuffer = ""
      If Comm.inbuffercount = 3 Then
        Response = Comm.CommInput(3)
        If Response = "ID=" Then
          TAP2_PROGRESS = TAP2_PROGRESS_ID_RECEIVED
          Comm.CommWrite ESC & "PG1" & CR
        Else
          Comm.CommInput (Comm.inbuffercount)
          Comm.CommWrite CR
        End If
      Else
        Comm.CommInput (Comm.inbuffercount)
        Comm.CommWrite CR
      End If
    
    Case TAP2_PROGRESS_ID_RECEIVED
      rxBuffer = rxBuffer & Comm.CommInput(Comm.inbuffercount)
      If InStr(1, rxBuffer, ESC & "[p") Then
        rxBuffer = ""
        TAP2_PROGRESS = TAP2_PROGRESS_BRACKET_P_RECIEVED
        Comm.CommWrite TAP2message
      End If
    
    Case TAP2_PROGRESS_BRACKET_P_RECIEVED
      rxBuffer = rxBuffer & Comm.CommInput(Comm.inbuffercount)
      If InStr(1, rxBuffer, ACK, vbTextCompare) Then
        ' done
        TAP2_PROGRESS = TAP2_PROGRESS_ACK_RECIEVED
        If Not (CurrentItem Is Nothing) Then  ' See if it was sent
          CurrentItem.SendComplete = 1
        End If
        rxBuffer = ""
      End If
    Case TAP2_PROGRESS_ACK_RECIEVED
      TAP2_PROGRESS = TAP2_PROGRESS_NOTSTARTED  ' do nothing
      mBusy = False
    Case Else
      ' error
  End Select




End Function

Function SendTAP2(ByVal PagerID As String, ByVal message As String) As Long

  On Error Resume Next
  mBusy = True

  If InIDE Then
    If PagerID = "" Then
      PagerID = "101"
    End If
  End If
  

  TAP2ResponseTimeout = DateAdd("s", 10, Now)
  message = TAPClean(message)
  message = STX & PagerID & CR & message & CR & ETX
  TAP2message = message & TAPCheckSum(message) & CR
  rxBuffer = ""
  Comm.CommFlush

  TAP2_PROGRESS = TAP2_PROGRESS_STARTED

  Exit Function
  '**************  EXIT **************





  '<STX>ID<CR>Message<CR><ETX>Checkum<CR>
  ' a customer wants <STX>pagerid<CR>message<ETX>Checkum<CR><LF><LF>.
  '10      On Error GoTo SendTAP_Error
  '
  '20      mBusy = True
  '
  '
  '60      rxBuffer = ""
  '70      Comm.CommFlush
  '80      message = TAPClean(message)
  '90      message = STX & PagerID & CR & message & CR & ETX
  '100     If LFeeds > 0 Then
  '110       message = message & TAPCheckSum(message) & CR & String$(LFeeds, 10)   ' added variable # of LF chars  4/8/13
  '120     Else
  '130       message = message & TAPCheckSum(message) & CR
  '140     End If
  '150     SendTAP2 = Comm.CommWrite(message)
  '
  '160     ACKFailTime = DateAdd("s", 5, Now)  ' must receive an ACK within 10 seconds (reduced to 5)
  '
  '        Dim CleanedMessage As String
  '170     CleanedMessage = Replace(message, vbCr, "^")
  '180     CleanedMessage = Replace(CleanedMessage, STX, "[")
  '190     CleanedMessage = Replace(CleanedMessage, ETX, "]")
  '200     CleanedMessage = Replace(CleanedMessage, vbLf, "!")
  '
  '210     Debug.Print "SendTAP2 ID/message "; CleanedMessage, Now, ACKFailTime
  '
  '220     LogPageEvent "SEND TAP2 " & vbTab & CleanedMessage
  '
  '
  'SendTAP2_Resume:
  '230     On Error GoTo 0
  '
  '240     RetriesRemaining = RetriesRemaining - 1
  '
  '250     SendTAP2 = 0
  '260     Exit Function
  '
  'SendTAP2_Error:
  '
  '270     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendTAP2." & Erl
  '        TAP2_Error = 1
  '280     mBusy = False
  '290     Resume SendTAP2_Resume
  '
End Function


Private Function CleanTAPMessage(ByVal message As String) As String
  message = Replace(message, vbCr, "^")
  message = Replace(message, STX, "[")
  message = Replace(message, ETX, "]")
  message = Replace(message, vbLf, "!")

  CleanTAPMessage = message
End Function

Function SendTAP(ByVal PagerID As String, ByVal message As String) As Long
  'TAP
  '<STX>ID<CR>Message<CR><ETX>Checkum<CR>
  ' a customer wants <STX>pagerid<CR>message<ETX>Checkum<CR><LF><LF>.
10 On Error GoTo SendTAP_Error

20 mBusy = True


60 rxBuffer = ""
70 Comm.CommFlush
80 message = TAPClean(message)
90 message = STX & PagerID & CR & message & CR & ETX
100 If LFeeds > 0 Then
110 message = message & TAPCheckSum(message) & CR & String$(LFeeds, 10)  ' added variable # of LF chars  4/8/13
120 Else
130 message = message & TAPCheckSum(message) & CR
140 End If
150 SendTAP = Comm.CommWrite(message)

160 ACKFailTime = DateAdd("s", 5, Now)  ' must receive an ACK within 10 seconds (reduced to 5)

  Dim CleanedMessage     As String
170 CleanedMessage = Replace(message, vbCr, "^")
180 CleanedMessage = Replace(CleanedMessage, STX, "[")
190 CleanedMessage = Replace(CleanedMessage, ETX, "]")
200 CleanedMessage = Replace(CleanedMessage, vbLf, "!")

210 Debug.Print "SendTAP ID/message "; CleanedMessage, Now, ACKFailTime

220 LogPageEvent "SEND TAP " & vbTab & CleanedMessage



SendTAP_Resume:
230 On Error GoTo 0

240 RetriesRemaining = RetriesRemaining - 1

250 SendTAP = 0
260 Exit Function

SendTAP_Error:

270 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendTAP." & Erl
  TAP_Error = 1
280 mBusy = False
290 Resume SendTAP_Resume


End Function

Function SendComp1(ByVal message As String, Optional ByVal SendCR As Boolean = True) As Long
  'COMP1
  'Message<CR>  (10 seconds will also terminate input as well as 256 in TX buffer)
  'All data is sent to all pagers
10 On Error GoTo SendComp1_Error

  message = Replace(message, vbCrLf, " ")
20 rxBuffer = ""
30 Comm.CommFlush
40 SendComp1 = Comm.CommWrite(message & CR)
50 LogPageEvent "SEND COMP1 " & message
  NextSend = DateAdd("s", Pause, Now)
60 mBusy = False

SendComp1_Resume:
70 On Error GoTo 0
80 Exit Function

SendComp1_Error:

90 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendComp1." & Erl
100 Resume SendComp1_Resume

End Function

Function SendApollo(ByVal message As String, Optional ByVal SendCR As Boolean = True) As Long




10 On Error GoTo SendApollo_Error

  Dim CR_BYTE            As String
  If SendCR Then
    CR_BYTE = vbCr
  End If


20 message = Trim$(Replace(message, vbCrLf, " "))
30 If InStr(1, message, "cancel", vbTextCompare) > 0 Then
    message = Replace(message, "cancel", "RESET", , , vbTextCompare)
  ElseIf InStr(1, message, "cancelled", vbTextCompare) > 0 Then
    message = Replace(message, "cancelled", "RESET", , , vbTextCompare)
50 End If
60 rxBuffer = ""
70 Comm.CommFlush

80 SendApollo = Comm.CommWrite(message & CR_BYTE)
90 LogPageEvent "SEND APOLLO " & message
100 NextSend = DateAdd("s", Pause, Now)


SendApollo_Resume:
110 mBusy = False
120 On Error GoTo 0
130 Exit Function

SendApollo_Error:

140 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendApollo." & Erl
150 Resume SendApollo_Resume

End Function

Function SendMarquis(ByVal message As String, Optional ByVal SendCR As Boolean = True) As Long

  ' change cancel into RESET, Append control code

  Dim MarquisControlCodeChar As String

10 Select Case MarquisControlCode
    Case 0
20    MarquisControlCodeChar = "$"
30  Case 1
40    MarquisControlCodeChar = "+"
50  Case 2
60    MarquisControlCodeChar = "="
70  Case 3
80    MarquisControlCodeChar = "#"
90  Case Else
100   MarquisControlCodeChar = "*"
110 End Select





  'All data is sent to all pagers
120 On Error GoTo SendMarquis_Error

130 message = Trim$(Replace(message, vbCrLf, " ")) & MarquisControlCodeChar
140 If (0 = StrComp(left$(message, 5), "CANCEL", vbTextCompare)) Then
150 Mid$(message, 1, 5) = "RESET"
160 End If
170 rxBuffer = ""
180 Comm.CommFlush
190 SendMarquis = Comm.CommWrite(message & CR)
200 LogPageEvent "SEND MQ " & message
210 NextSend = DateAdd("s", Pause, Now)


SendMarquis_Resume:
220 mBusy = False
230 On Error GoTo 0
240 Exit Function

SendMarquis_Error:

250 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendMarquis." & Erl
260 Resume SendMarquis_Resume

End Function





Function SendComp2(ByVal PagerID As String, ByVal message As String) As Long
  Dim s                  As String



  'COMP2
  ' PagerID<CR>Message<CR>
  ' to resync <DEL><DEL><DEL>
  ' response ID<CR><CR><RS><CR> if error (device not listed in TX database)
10 On Error GoTo SendComp2_Error
  message = Replace(message, vbCrLf, " ")
20 rxBuffer = ""
30 Comm.CommFlush
40 s = PagerID & CR & message & CR
50 SendComp2 = Comm.CommWrite(s)
  LogPageEvent "SEND COMP2 " & PagerID & message
60 NextSend = DateAdd("s", Pause, Now)
70 ACKFailTime = DateAdd("s", 10, Now)  ' must receive an ACK within 10 seconds
80 mBusy = False

SendComp2_Resume:
90 On Error GoTo 0
100 Exit Function

SendComp2_Error:

110 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendComp2." & Erl
120 Resume SendComp2_Resume

End Function

Function ReSyncComp2() As Long
10 On Error GoTo ReSyncComp2_Error

20 ReSyncComp2 = Comm.CommWrite(String(3, DEL))  ' <DEL><DEL><DEL>

ReSyncComp2_Resume:
30 On Error GoTo 0
40 Exit Function

ReSyncComp2_Error:

50 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.ReSyncComp2." & Erl
60 Resume ReSyncComp2_Resume

End Function


Function SendTTS(ByVal message As String, Optional RelayID As Long = 0) As Long
  ' if address is 0 then it's autofind

10 On Error Resume Next
  'Dim t             As Long
  Dim BytesWritten       As Long
  Dim Buffer             As String
  Dim pd                 As cPageDevice
  Dim ErrNum             As Long

  VoiceStart = Now


20 message = Replace(message, vbCrLf, " ")

30 VoiceStart = Now            ' t = Win32.timeGetTime()
40 If IsDate(message) Then
50  message = ConvertDatetoTTS(message)
60 End If

70 mLastRelayID = RelayID


80 If PASystemKey = 1 Then

90  For Each pd In gPageDevices
100   If pd.ProtocolID = PROTOCOL_ONTRAK Then  ' we have an ontrak server to use
110     Set ExtPageDevice = pd.ADU200
120     Exit For
130   End If
140 Next

150 If ExtPageDevice Is Nothing Then  ' legacy, no ontrak object, handle the relay directly
160   If PASystemHandle > 0 Then
170     CloseAduDevice (PASystemHandle)
180     PASystemHandle = 0
190   End If

200   PASystemHandle = OpenAduDevice(ADU_USE_TIMEOUTS)  ' relay board
210   If PASystemHandle > 0 Then
220     Buffer = "SK0"
        ' Key PA System
230     WriteAduDevice PASystemHandle, Buffer, Len(Buffer), BytesWritten, 400

240   End If

250 Else                       'handle via the ontrak object
260   ExtPageDevice.SetDedicatedRelay RELAY_ON, mLastRelayID  ' if address is 0 then it's autofind
270 End If

    '    t = t + 900  ' < one second wait

    'Dim Pause As Long
    Dim Pausetime        As Date
    Pausetime = DateAdd("s", 1, Now)
280 'Pause = DateAdd("s", 1, Now)

    '290       Do While t > Win32.timeGetTime()
290 Do While Pausetime > Now
      'Sleep 50
300   DoEvents
310 Loop

320 End If
  dbg "TTS.SpeakMessage"
  Dim ttsRC              As Long
330 ttsRC = TTS.Speak(message, SVSFlagsAsync)
  dbg "TTS" & " " & message
340 ErrNum = Err.Number
350 If ErrNum <> 0 Then
360 If ErrNum = 53 Then        ' relay problem

370 Else                       ' other problem
380   SendTTS = Err.Number
390   NextSend = DateAdd("s", Pause, Now)
400   If PASystemHandle > 0 Then
410     CloseAduDevice (PASystemHandle)
420     PASystemHandle = 0
430   End If
440 End If
450 mBusy = False
460 End If

SendTTS_Resume:
470 On Error GoTo 0
480 Exit Function

SendTTS_Error:

490 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.SendTTS." & Erl
500 Resume SendTTS_Resume


End Function

Function TAPClean(ByVal message As String) As String
  ' convert all these to spaces
  message = Replace(message, "^", " ")
  message = Replace(message, "~", " ")
  message = Replace(message, "_", " ")
  message = Replace(message, CR, SUB_CR)
  message = Replace(message, LF, SUB_LF)
  TAPClean = message
End Function

Function TAPCheckSum(ByVal message As String) As String
  Dim j                  As Long
  Dim Sum                As Long
  Dim CK1                As String
  Dim CK2                As String
  Dim CK3                As String
  'These go into the checksum
  'STX & PagerID & CR & Message & CR & ETX

  For j = 1 To Len(message)
    Sum = Sum + (Asc(MID(message, j, 1)) And &H7F)
  Next
  Sum = Sum And &HFFF&
  CK3 = Chr$(48 + (Sum - (Sum \ 16) * 16))
  Sum = Sum \ 16
  CK2 = Chr$(48 + (Sum - (Sum \ 16) * 16))
  Sum = Sum \ 16
  CK1 = Chr$(48 + (Sum - (Sum \ 16) * 16))
  TAPCheckSum = CK1 & CK2 & CK3

End Function

Private Sub Class_Terminate()
10 On Error GoTo Class_Terminate_Error
  '        If Not FC Is Nothing Then
  '          'FC.Disable
  '          Set FC = Nothing
  '        End If



20 If Not Comm Is Nothing Then
30  CloseComm
40  Set Comm = Nothing
50 End If
60 If Not TTS Is Nothing Then
70  Set TTS = Nothing
80 End If

90 If Not Dialer Is Nothing Then
100 Set Dialer = Nothing
110 End If



120 If Not ADU200 Is Nothing Then
130 Set ADU200 = Nothing
140 End If

150 If Not DivaDialer Is Nothing Then
160 Set DivaDialer = Nothing
170 End If


180 If Not WSock Is Nothing Then
190 WSock.Close
    WSock.tag = ""
200 Set WSock = Nothing
210 End If



Class_Terminate_Resume:
220 On Error GoTo 0
230 Exit Sub

Class_Terminate_Error:

240 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.Class_Terminate." & Erl
250 Resume Class_Terminate_Resume

End Sub


Public Function GetErrorStatus() As Long
  GetErrorStatus = ErrorStatus
End Function

Public Sub ClearError()
  ErrorStatus = 0
End Sub
Private Sub CloseComm()
  ' Just say goodbye
  On Error Resume Next
  Comm.CommFlush
  Comm.CommClose
End Sub


Private Sub Class_Initialize()
  NUL = Chr(0)
  SOH = Chr(1)
  STX = Chr(2)
  ETX = Chr(3)
  EOT = Chr(4)
  ENQ = Chr(5)
  ACK = Chr(6)
  BEL = Chr(7)
  BS = vbBack
  HT = vbTab
  VT = Chr(&HB)
  FF = Chr(&HC)
  CR = vbCr
  SO = Chr(&HE)
  si = Chr(&HF)
  LF = vbLf
  XOFF = Chr(&H11)
  XON = Chr(&H13)
  NAK = Chr(&H15)
  ETB = Chr(&H17)
  SUBS = Chr(&H1A)
  ESC = Chr(&H1B)
  rs = Chr(&H1E)
  US = Chr(&H1F)
  DEL = Chr(&H7F)
  '
  SUB_CR = SUBS & "M"
  SUB_LF = SUBS & "J"

  TIMEOUTDELAY = 30            ' 30 seconds default timeout, can be overridden
  NextRetry = Now              ' for winsock, may also be applied for comm port
  KeepAliveTimeout = DateAdd("s", KEEPALAIVE_DELAY, Now)

  If CLng(NextSend) <= 0 Then
    NextSend = DateAdd("s", 10, Now)
  End If

  If CLng(NextPing) <= 0 Then
    NextPing = DateAdd("s", PING_DELAY, Now)
  End If

  'OpenConnection

  Set Que = New Collection



End Sub



Private Sub TTS_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
  ' This is event driven
  ' flags when voice has completed talking
  Dim Buffer             As String
  Dim BytesWritten       As Long
  ' Dim t             As Long

  Dim Pausetime          As Date

  dbg "TTS Endstream"



  On Error GoTo TTS_EndStream_Error

10 If NoWait Then
20  NextSend = DateAdd("s", 0, Now)  ' pause between messages
30 Else
40  NextSend = DateAdd("s", Pause, Now)  ' pause between messages
50 End If

60 If Not ExtPageDevice Is Nothing Then
70  ExtPageDevice.SetDedicatedRelay RELAY_OFF, mLastRelayID
80  Set ExtPageDevice = Nothing
90 ElseIf PASystemHandle > 0 Then
100 Buffer = "RK0"
110 WriteAduDevice PASystemHandle, Buffer, Len(Buffer), BytesWritten, 500

120 't = Win32.timeGetTime() + 500  ' half second wait
    Pausetime = DateAdd("s", 1, Now)
130 Do While Pausetime > Now   'Do While t > Win32.timeGetTime()
140   DoEvents
150 Loop

160 CloseAduDevice (PASystemHandle)
170 PASystemHandle = 0
180 End If

  '  End If
190 mBusy = False

TTS_EndStream_Resume:
  On Error GoTo 0
  Exit Sub

TTS_EndStream_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageDevice.TTS_EndStream." & Erl
  Resume TTS_EndStream_Resume

End Sub



'protocols

'COMP1
'Message<CR>  (10 seconds will also terminate input as well as 256 in TX buffer)
'All data is sent to all pagers

'COMP2
'PagerID<CR>Message<CR>
' to resync <DEL><DEL><DEL>
' response ID<CR><CR><RS><CR> if error (device not listed in TX database)

' TIMED "Watchdog" 1.06 minutes
' If enabled
'1) any valid TAP message resets
'2) send ATI<CR> to reset

'TAP
' login in is <CR> every two seconds
' response to login is ID=<CR>
' Reply with <ESC>PG1<CR>
' system Announces 110 1.8<CR>System Name<CR><ACK><CR> note: 1.8 is TAP version
' system is ready with reponse: <ESC>[p<CR>
' <STX>ID<CR>Message<CR><ETX>Checkum<CR>
' returns <ACK>
' or      <NAK>
' logout is <EOT><CR>


' Message can be 235 Alpha 7 bit data except control chars and "^~_"

' ID of 1 to 4 digits does internal DB lookup


'The following represents a typical call flow in sending one page request into a paging network.

'   Remote Entry Device               Paging Terminal
'   -------------------               ---------------
'1) Dials paging terminal             Modem Answers
'2) Modem Connects
'3) <CR>
'4)                                   ID=
'5) <ESC>PG1<CR>
'6)                                   110 1.7<CR>
'                                     Thank you for calling the PCIA<CR>
'                                     <ACK><CR>
'                                     <ESC>[p<CR>
'7) <STX>123<CR>ABC<CR><ETX>17;<CR>   211 Page accepted <CR>
'                                     <ACK><CR>
'8) <EOT><CR>
'9)                                   115 Thank you for calling <CR>
'                                     <ESC><EOT><CR>
'10)                                  Drops Carrier
'11) Drops Carrier

' TTS

Private Function SendPCA(ByVal Address As String, ByVal message As String) As Long
  On Error Resume Next
  '  Dim t         As Long
  Dim ErrNum             As Long
  Dim PCA                As cPCA

  Set PCA = New cPCA

  message = Replace(message, "|", "'")
  message = Replace(message, vbCrLf, "|")

  '  t = Win32.timeGetTime()
  If IsDate(message) Then
    message = ConvertDatetoTTS(message)
  End If

  PCA.Send Address, message, Alert

  SendPCA = Err.Number
  NextSend = DateAdd("s", Pause, Now)
  mBusy = False


End Function



Public Property Get DialerPhone() As String

  DialerPhone = mDialerPhone

End Property

Public Property Let DialerPhone(ByVal Value As String)

  mDialerPhone = Value

End Property

Public Property Get DialerTerminateDigit() As Long
  DialerTerminateDigit = mDialerTerminateDigit
End Property

Public Property Let DialerTerminateDigit(ByVal Value As Long)
  mDialerTerminateDigit = Value
End Property

Public Property Get Relay1() As Integer
  Relay1 = mRelay1
End Property
Public Property Let Relay1(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 1, Value
    mRelay1 = Value
  End If
End Property
Public Property Get Relay2() As Integer
  Relay2 = mRelay2
End Property
Public Property Let Relay2(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 2, Value
    mRelay2 = Value
  End If
End Property
Public Property Get Relay3() As Integer
  Relay3 = mRelay3
End Property
Public Property Let Relay3(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 3, Value
    mRelay3 = Value
  End If
End Property
Public Property Get Relay4() As Integer
  Relay4 = mRelay4
End Property
Public Property Let Relay4(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 4, Value
    mRelay4 = Value
  End If
End Property

Public Property Get DeviceID() As Long

  DeviceID = mDeviceID

End Property

Public Property Let DeviceID(ByVal DeviceID As Long)

  mDeviceID = DeviceID

End Property

Public Property Get Relay5() As Integer

  Relay5 = mRelay5

End Property

Public Property Let Relay5(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 5, Value
    mRelay5 = Value
  End If

End Property

Public Property Get Relay6() As Integer

  Relay6 = mRelay6

End Property

Public Property Let Relay6(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 6, Value
    mRelay6 = Value
  End If

End Property

Public Property Get Relay7() As Integer

  Relay7 = mRelay7

End Property

Public Property Let Relay7(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 7, Value
    mRelay7 = Value
  End If

End Property

Public Property Get Relay8() As Integer

  Relay8 = mRelay8

End Property

Public Property Let Relay8(ByVal Value As Integer)
  If Not ADU200 Is Nothing Then
    ADU200.ConfigureRelay 8, Value
    mRelay8 = Value
  End If

End Property

Public Property Get DialerModem() As Long

  DialerModem = mDialerModem

End Property

Public Property Let DialerModem(ByVal Value As Long)


  mDialerModem = Value

End Property

Sub LogPageEvent(ByVal s As String)
  Dim hfile              As Long
  Dim RightNow           As Date
  Dim filename           As String

  '   Debug.Print s
  If gLogTAP Then

    On Error Resume Next
    filename = App.Path & "\pages.log"
    limitFileSize filename
    RightNow = Now
    hfile = FreeFile
    Open filename For Append As hfile
    Print #hfile, s & vbTab & Format(RightNow, "yyyy-mm-dd") & " " & Format(RightNow, "hh:nn:ss")
    Close hfile

  End If
End Sub

Private Sub WSock_Close()
  '
End Sub

Private Sub WSock_Connect()
  '
End Sub

Private Sub WSock_ConnectionRequest(ByVal requestID As Long)
  '
End Sub

Private Sub WSock_DataArrival(ByVal bytesTotal As Long)
  '
End Sub

Private Sub WSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  '
End Sub

Private Sub WSock_SendComplete()
  '
End Sub

Private Sub WSock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  '
End Sub

Private Sub TTS_StartStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
  dbg "TTS_StartStream"
End Sub

Public Property Get KeepPaging() As Long

  KeepPaging = mKeepPaging

End Property

Public Property Let KeepPaging(ByVal KeepPaging As Long)

  mKeepPaging = KeepPaging

End Property

Public Property Get SDACTError() As Long

  SDACTError = mSDACTError

End Property

Public Property Let SDACTError(ByVal SDACTError As Long)

  mSDACTError = SDACTError

End Property

Public Property Get Port() As Integer

  Port = mPort

End Property

Public Property Let Port(ByVal Port As Integer)

  mPort = Port

End Property
