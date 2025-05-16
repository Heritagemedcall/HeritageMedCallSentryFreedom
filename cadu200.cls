VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADU200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' get handle
' send commands
' rx data

' close handle
' vendor ID = &h0a07
' productid typically 200
' serial number &h000000 to &hFFFFFF

' commands

'Port K   Relays ************
'SKn      Sets relay n (0 to 3)
'RKn      Clears relay n (0 to 3)
'MKdd     Sets entire port to decimal dd (0 to 15)
'SPKxxxx  Sets entire port bitwise (0 or 1)
'RPKn     Returns status of relay n (0 t 3) one Byte
'RPK      Returns status of entire port as binary (bitwise 0 or 1) 4 Bytes
'PK       Returns status of entire port as decimal (0 to 15) 2 Bytes

'Port A   Input Ports ************
'RPAn     Returns status of input n (0 t 3) one Byte
'RPA      Returns status of entire port as binary (bitwise 0 or 1) 4 Bytes
'PA       Returns status of entire port as decimal (0 to 15) 2 Bytes

'Event Counters
'REx      Returns count of event counter (x=0 to 3) 5 Bytes 00000 to 65535
'RCx      Returns count of event counter and clears value (x=0 to 3) 5 Bytes 00000 to 65535
'DBn      Set Debounce Timer, all ports 0=10ms, 1 = 1ms , 2 = 100us
'DB       Returns Debounce setting  one Byte (0 to 2)

'Special settings
'SBn      Set Baud Rate of RS232 (0=9600, 1=19.2K, 2=38K, 3=56K)
'SB       Returns Baud Rate setting
'WDn      Set Watchdog timeout (0=OFF, 1=1s, 2=10s, 3=60s)
'WD       Returns Watchdog setting


'moved to modLib

'Const MODE_RESERVED = 0   ' control is by other provider (KEY PA SYSTEM)
'Const MODE_ALWAYSON = 3  ' keeps relay closed as long as there is something in the Que
'                          ' QUE does NOT get emptied automatically
'Const MODE_FLASHER = 2    ' toggles relay at once second on/off times as long as there is something in the Que
'                          ' QUE does NOT get emptied automatically
'Const MODE_ONESHOT = 1    ' sets relay closed for one second
'                          ' Empties QUE on each pass (rising pulse)

'Const SET_RELAY = "SK"
'Const RESET_RELAY = "RK"

'Const RELAY_OFF = False
'Const RELAY_ON = True
'


Private mDelayTime As Long
Private mHandle  As Long

Private mWatchDogTimeout As Integer
  
Private mRelays(1 To MAX_RELAYS) As ADU200Relay
Private RelayQue(1 To MAX_RELAYS) As Collection

Public Sub SetDedicatedRelay(ByVal Value As Boolean, Optional RelayNum As Long = 0)
        Dim j    As Long
        Debug.Print "SetDedicatedRelay On/Off Relay " & Value, RelayNum
10      If RelayNum = 0 Then  ' auto find
20        For j = 1 To MAX_RELAYS
30          If mRelays(j).Mode = MODE_RESERVED Then
40            If (Value) Then
50              SendCommand SET_RELAY & CStr(j - 1)
60            Else
70              SendCommand RESET_RELAY & CStr(j - 1)
80            End If
90            Exit For
100         End If

110       Next
120     Else  ' choose specific relay
130       For j = 1 To MAX_RELAYS
140         If j = RelayNum Then
150           If (Value) Then
160             SendCommand SET_RELAY & CStr(j - 1)
170           Else
180             SendCommand RESET_RELAY & CStr(j - 1)
190           End If
200           Exit For
210         End If
220       Next
230     End If

End Sub



Public Sub ConfigureRelay(ByVal index As Integer, ByVal Value As Integer)
  If index > 0 And index <= MAX_RELAYS Then
    mRelays(index).Mode = Value
    SendCommand RESET_RELAY & CStr(index)
    mRelays(index).OnOff = RELAY_OFF
    mRelays(index).LastChange = Now
    Set RelayQue(index) = New Collection
  End If
End Sub

Public Sub UpdateClock()
  Dim Relay As ADU200Relay
  Dim Que   As Collection
  Dim p     As cPageItem
  Dim j     As Integer


  For j = 1 To MAX_RELAYS

    Set Relay = mRelays(j)
    Set Que = RelayQue(j)
    If Que.Count > 0 Then

      Select Case Relay.Mode

        Case MODE_ALWAYSON  ' anything in this que turns it on
          If (Relay.OnOff = RELAY_OFF) Then  'trigger it on
            SendCommand SET_RELAY & CStr(j - 1)
            Relay.OnOff = RELAY_ON
            Relay.LastChange = Now
          End If

        Case MODE_FLASHER  ' anything in this que turns it on for flashing
          If (1 < DateDiff("s", Relay.LastChange, Now)) Then
            If (Relay.OnOff = RELAY_OFF) Then  'trigger it on
              SendCommand SET_RELAY & CStr(j - 1)
              Relay.OnOff = RELAY_ON
              Relay.LastChange = Now
            Else
              SendCommand RESET_RELAY & CStr(j - 1)
              Relay.OnOff = RELAY_OFF
              Relay.LastChange = Now
            End If
          End If

        Case MODE_ONESHOT ' one flash per request
            ' relay is only ready if relay is off AND one-second has passed since it was turned off
          If ((Relay.OnOff = RELAY_OFF) And (1 < DateDiff("s", Relay.LastChange, Now))) Then
            RelayQue(j).Remove 1  ' never to be used again
            SendCommand SET_RELAY & CStr(j - 1)
            Relay.OnOff = RELAY_ON
            Relay.LastChange = Now
          ElseIf ((Relay.OnOff = RELAY_ON) And (1 < DateDiff("s", Relay.LastChange, Now))) Then

            SendCommand RESET_RELAY & CStr(j - 1)
            Relay.OnOff = RELAY_OFF
            Relay.LastChange = Now
          
          End If



        Case Else  '  MODE_RESERVED
          'we never update via this way with reserved (PA MIC KEYING)
      End Select
    Else

      Select Case Relay.Mode

        Case MODE_ALWAYSON
          If (Relay.OnOff = RELAY_ON) Then  'trigger it OFF
            SendCommand RESET_RELAY & CStr(j - 1)
            Relay.OnOff = RELAY_OFF
            Relay.LastChange = Now
          End If

        Case MODE_FLASHER
          If (Relay.OnOff = RELAY_ON) Then  'trigger it OFF
            SendCommand RESET_RELAY & CStr(j - 1)
            Relay.OnOff = RELAY_OFF
            Relay.LastChange = Now
          End If

        Case MODE_ONESHOT
          If (Relay.OnOff = RELAY_ON) Then  'trigger it OFF
            SendCommand RESET_RELAY & CStr(j - 1)
            Relay.OnOff = RELAY_OFF
            Relay.LastChange = Now
          End If
        Case Else  '  MODE_RESERVED
          'we never update via this way with reserved (PA MIC KEYING)
      End Select


    End If
  Next


End Sub

Public Function AddItem(Item As cPageItem) As Long
  Dim p As cPageItem
  Dim j As Integer
  Dim Que As Collection

  
  
  Set p = New cPageItem
  p.AlarmID = Item.AlarmID
  p.RelayID = Item.RelayID
  p.IsCancel = Item.IsCancel

  If p.RelayID > 0 And p.RelayID <= MAX_RELAYS Then
    
    
    
    
    
    Set Que = RelayQue(p.RelayID)
    If (p.IsCancel) Then                   ' end of this alarm
      For j = Que.Count To 1 Step -1
        If (Que(j).AlarmID = p.AlarmID) Then ' remove it if alarm IDs match
          Que.Remove j
          'dbg "Removed cADU200 AlarmID, RelayID, Mode " & Item.AlarmID & ", " & Item.RelayID & ", " & mRelays(p.RelayID).Mode
        End If
      Next
    Else ' just add it
      'dbg "Added to cADU200 AlarmID, RelayID, Mode " & Item.AlarmID & ", " & Item.RelayID & ", " & mRelays(p.RelayID).Mode
      Que.Add p
    End If
  End If


End Function


Public Sub SetRelayStatus(ByVal RelayID As Long, ByVal Status As Boolean)
  
  mRelays(RelayID).Status = Status

End Sub
Public Function GetRelayStatus(ByVal RelayID As Long) As Boolean
 GetRelayStatus = mRelays(RelayID).Status
End Function


'Public Sub UpdateClock()
'  'called once a second from cpagedevice (which is called from the master clock)
'  Dim CurrentTime As Double
'  Dim elapsed As Long
'
'
'
'  CurrentTime = Timer
'  If StartTime > CurrentTime Then
'    'Debug.Print "Rollover"
'    elapsed = (86400 - StartTime) + CurrentTime
'  Else
'    elapsed = CurrentTime - StartTime
'  End If
'  If elapsed >= 1 Then
'    StartTime = Timer
'    'CheckStatus
'  End If
'
'
'
'
'  'Debug.Print "Elapsed = " & Format(Elapsed, "0.00")
'End Sub
'Sub CheckStatus()
'  Dim j As Integer
'
'  For j = 1 To 4
'    If RelayType(j).status = True Then
'    Select Case RelayType(j).Mode
'      Case MODE_ALWAYSON
'        If Not (RelayType(j).OnOff) Then
'          RelayType(j).OnOff = True
'        End If
'
'      Case MODE_ONESHOT
'          If RelayType(j).OnOff = True Then
'            RelayType(j).OnOff = False
'          End If
'      Case MODE_BLINK
'        RelayType(j).OnOff = Not (RelayType(j).OnOff)
'      Case Else
'    End Select
'    End If
'  Next
'
'End Sub




Public Function SendCommand(ByVal Command As String) As Long
  Dim BytesWritten As Long
  'dbg "************ ADU command " & Command
  If (Handle > 0) Then
    WriteAduDevice Handle, Command, Len(Command), BytesWritten, 400
  Else
    dbg "ADU has no open channel " & Command
  End If
  SendCommand = BytesWritten
End Function

'Public Function GetResult()
'
'End Function


'Public Function SetBaudRate(ByVal Value As Integer) As Integer
'  If Value >= 0 And Value <= 3 Then
'    SetBaudRate = GetBaudRate
'  Else
'    SetBaudRate = -1
'  End If
'End Function
'Public Function GetBaudRate() As Integer
'  'SB
'End Function
'
'Public Function GetAllPortA() As Integer
'
'End Function
'
'Public Function GetPortA(ByVal InputNum As Integer) As Integer
'  If InputNum >= 0 And InputNum <= 3 Then
'    GetPortA = 0
'  Else
'    GetPortA = -1
'  End If
'
'End Function
'Public Function OpenDevice(ByVal Serial As String) As Long
'  'returns handle if successful
'End Function
'Public Function SetPortK(ByVal Value As Integer) As Integer
'  'MKdd dd = 0 to 15
'End Function
'Private Function GetPortK(ByVal RelayNum As Integer) As Integer
'  ' returns 0 or 1
'  ' RPKn
'End Function

'Public Function SetRelay(ByVal RelayNum As Integer) As Integer
'  ' sets relay to on if within range
'  ' returns -1 if out of range
'
'
'
'  If RelayNum >= 0 And RelayNum <= 3 Then
'    'dbg "Set Relay " & RelayNum + 1
'
'    SetRelay = RelayNum
'  Else
'    'dbg "Set Relay Out Of Range Error"
'    SetRelay = -1
'  End If
'End Function
'Public Function GetRelay(ByVal RelayNum As Integer) As Integer
'  ' gets status of relay to on if within range
'  ' returns -1 if out of range
'  If RelayNum >= 0 And RelayNum <= 3 Then
'    'GetRelay = GetStatus(RelayNum)
'  Else
'    GetRelay = -1
'  End If
'End Function

'Public Function SetAllRelays() As Integer
'  ' turn all 4 relays on
'
'End Function
'Public Function GetAllRelays() As Integer
'  'get status of all 4 relays
' 'PK
'
'End Function

Public Property Get Handle() As Long
  Handle = mHandle
End Property

Public Property Let Handle(ByVal Handle As Long)
  mHandle = Handle
End Property

Public Property Get DelayTime() As Long
  DelayTime = mDelayTime
End Property
Public Property Let DelayTime(ByVal DelayTime As Long)
  mDelayTime = DelayTime
End Property

'Public Function GetWatchDogTimeout() As Integer
'  'WD
'  GetWatchDogTimeout = 0
'End Function

'Public Function SetWatchDogTimeout(ByVal Value As Integer) As Integer
'  'WDn
'  SendCommand "WD" & CStr(Value)
'  SetWatchDogTimeout = GetWatchDogTimeout()
'End Function


Private Sub Class_Initialize()
  Dim j As Integer
  For j = 1 To MAX_RELAYS
    Set RelayQue(j) = New Collection
    Set mRelays(j) = New ADU200Relay
  Next
  Handle = OpenAduDevice(ADU_USE_TIMEOUTS)
     
  If Handle > 0 Then ' turn off all relays
    For j = 1 To MAX_RELAYS
      SendCommand RESET_RELAY & CStr(j)
    Next
  End If
  
End Sub

Sub Finalize()
  Dim j As Integer
  If Handle > 0 Then ' turn off all relays
    For j = 1 To MAX_RELAYS
      SendCommand RESET_RELAY & CStr(j)
    Next
  End If

  CloseAduDevice Handle
  Handle = 0
End Sub

Private Sub Class_Terminate()
  Finalize
End Sub
