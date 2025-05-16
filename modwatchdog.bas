Attribute VB_Name = "modWatchDog"
Option Explicit
Global LastWatchDog         As Date
Global Const WD_NONE = 0
Global Const WD_UL = 1
Global Const WD_BERKSHIRE = 2
Global Const WD_ARK3510 = 3

Global BerkshireWD     As cBerkshire

Global ARK3510_started As Boolean

Dim value As Integer
Dim PortAddress As Integer
Dim start As Date

Const UNLOCK_DATA = &H87
Const INDEX_PORT = &H2E
Const DATA_PORT = &H2F
Const DEVICE_REGISTER = &H7
Const WATCH_DOG_TIMER_DEVICE_NO = &H8
Const ACTIVE_PORT_SELECT = &H30
Const ACTIVE_PORT1_ADD = &H1
Const SELECT_TIMETYPE = &HF5
Const TIME_SECOND = &H0
Const TIME_MINUTE = &H8
Const TIMER_VALUE = &HF6
Const LOCK_DATA = &HAA

'EC_Command_Port = 0x29Ah
'EC_Data_Port = 0x299h
'Write EC HW ram = 0x89
'Watch dog event flag = 0x57
'Watchdog reset delay time = 0x5E
'Reset event = 0x04
'Start WDT function = 0x28


'mov dX, EC_COMMAND_PORT
'mov al,89h ; Write EC HW ram.
'out dX, al
'mov dX, EC_DATA_PORT
'mov al, 5Fh ; Watchdog reset delay time low byte (5Eh is high byte) index, Timebase:
'100 ms
'out dX, al
'mov dX, EC_DATA_PORT
'mov al, 64h ;Set 10 seconds delay time.
'out dX, al
'mov dX, EC_COMMAND_PORT
'mov al,89h ; Write EC HW ram.
'out dX, al
'mov dX, EC_DATA_PORT
'mov al, 57h ; Watch dog event flag.
'out dX, al
'mov dX, EC_DATA_PORT
'mov al, 04h ; Reset event.
'out dX, al
'mov dX, EC_COMMAND_PORT
'mov al,28h ; start WDT function. (Stop: 0x29, Reset: 0x2A)


'UL2_
'EC Watchdog Timer sample code
'Const EC_Command_Port = &H29A
'Const EC_Data_Port = &H299

'Write EC HW ram = &h89
'Watch dog event flag = &h57
'Watchdog reset delay time = &h5E
'Reset event = &h04
'Start WDT function = &h28

'mov dX, EC_Command_Port
'mov al,89h ; Write EC HW ram.
'out dX, al
'mov dX, EC_Command_Port
'mov al, 5Fh ; Watchdog reset delay time low byte (5Eh is high byte) index.
'out dX, al
'mov dX, EC_Data_Port
'mov al, 30h ;Set 3 seconds delay time.
'out dX, al
'mov dX, EC_Command_Port
'mov al,89h ; Write EC HW ram.
'out dX, al
'mov dX, EC_Command_Port
'mov al, 57h ; Watch dog event flag.
'out dX, al
'mov dX, EC_Data_Port
'mov al, 04h ; Reset event.
'out dX, al
'mov dX, EC_Command_Port
'mov al,28h ; start WDT function.
'out dX, al





'Inp and Out declarations for port I/O using inpout32.dll.

'Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" _
'    (ByVal PortAddress As Integer) _
'    As Integer
'
'Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" _
'    (ByVal PortAddress As Integer, _
'    ByVal Value As Integer)
'

Public Declare Function inportb Lib "inpout32.dll" Alias "Inp32" _
    (ByVal PortAddress As Integer) _
    As Integer
    
Public Declare Sub outportb Lib "inpout32.dll" Alias "Out32" _
    (ByVal PortAddress As Integer, _
    ByVal value As Integer)

'SUSI4.DLL is used by Advantech ARK-3510

Const SUSI_STATUS_SUCCESS = 0
Const SUSI_STATUS_NOT_INITIALIZED = &HFFFFFFFF
Const SUSI_STATUS_INITIALIZED = &HFFFFFFFE
Const SUSI_STATUS_UNSUPPORTED = &HFFFFFCFF ' -769 or 3327

Const SUSI_ID_WDT_DELAY_MAXIMUM = &H1
Const SUSI_ID_WDT_DELAY_MINIMUM = &H2
Const SUSI_ID_WDT_EVENT_MAXIMUM = &H3
Const SUSI_ID_WDT_EVENT_MINIMUM = &H4
Const SUSI_ID_WDT_RESET_MAXIMUM = &H5
Const SUSI_ID_WDT_RESET_MINIMUM = &H6
Const SUSI_ID_WDT_UNIT_MINIMUM = &HF
Const SUSI_ID_WDT_DELAY_TIME = &H10001
Const SUSI_ID_WDT_EVENT_TIME = &H10002
Const SUSI_ID_WDT_RESET_TIME = &H10003
Const SUSI_ID_WDT_EVENT_TYPE = &H10004

Private Declare Function SusiWDogTrigger Lib "Susi4.dll" (ByVal ID As Long) As Long
Private Declare Function SusiWDogStop Lib "Susi4.dll" (ByVal ID As Long) As Long
Private Declare Function SusiWDogStart Lib "Susi4.dll" (ByVal ID As Long, ByVal DelayTime As Long, ByVal EventTime As Long, ByVal ResetTime As Long, ByVal EventType As Long) As Long
Private Declare Function SusiWDogGetCaps Lib "Susi4.dll" (ByVal ID As Long, ByVal ItemID As Long, ByRef value As Long) As Long
Private Declare Function SusiLibInitialize Lib "Susi4.dll" () As Long




Public Sub SendTestPing()
  PingMonitor True, "facilityid=" & Configuration.MonitorFacilityID & "&eventcode=0"
End Sub

'Public CountdownNo    As Byte
'Public io_index_port  As Integer
'Public io_data_port   As Integer
'Public time_type      As Byte

Sub ARK3510_WatchDogTimer(ByVal Timeout As Long, Optional ByVal TimeType As Integer = 0)

  ' Uses Susi4.DLL


  Select Case Timeout
    Case Is >= 1
      If ARK3510_started Then
        Susi_Trigger
      Else
        Timeout = Configuration.WatchdogTimeout
        ARK3510_started = Susi_Start(Timeout * 1000)
      End If
    Case Is < 1
      Susi_Stop
      ARK3510_started = False
    Case Else
      Susi_Stop
      ARK3510_started = False
  End Select


End Sub

Function Susi_Trigger()
  Dim rc As Long
  rc = SusiWDogTrigger(0)
  If rc <> SUSI_STATUS_SUCCESS Then
    rc = SusiLibInitialize()
    rc = SusiWDogTrigger(0)
  End If
End Function


Function Susi_Stop()
  Dim rc As Long
  rc = SusiWDogStop(0)
  If rc <> SUSI_STATUS_SUCCESS Then
    rc = SusiLibInitialize()
    rc = SusiWDogStop(0)
  End If
  
End Function

Function Susi_Start(ByVal ResetTime As Long) As Boolean
  Dim rc            As Long
  Dim StoredValue   As Long
  
  'reset time is in ms.
  
  rc = SusiLibInitialize()
  rc = Susi_Stop()
  rc = SusiWDogStart(0, 0, 0, ResetTime, 0)
  rc = SusiWDogGetCaps(0, SUSI_ID_WDT_RESET_TIME, StoredValue)
  Susi_Start = ResetTime = StoredValue
  ' returns true on success
  
End Function

Sub wdsave(ByVal s As String)
  Dim hfile As Long
  Dim filename As String
  filename = App.Path & "\WDTrace.log"
  limitFileSize filename
  On Error Resume Next
  hfile = FreeFile
  Open filename For Append As hfile
  Print #hfile, s
  Close #hfile
  
End Sub




Sub WinbondWatchDogTimer(ByVal Timeout As Integer, ByVal TimeType As Integer)

' SET TIMEOUT TO 0 TO DISABLE
  
'
'  outportb io_index_port, UNLOCK_DATA
'  outportb io_index_port, UNLOCK_DATA
'  outportb io_index_port, DEVICE_REGISTER
'
'  outportb io_data_port, WATCH_DOG_TIMER_DEVICE_NO
'
'  outportb io_index_port, ACTIVE_PORT_SELECT
'  outportb io_data_port, ACTIVE_PORT1_ADD
'
'  outportb io_index_port, SELECT_TIMETYPE
'  outportb io_data_port, timetype
'
'  outportb io_index_port, TIMER_VALUE
'  outportb io_data_port, CountdownNo
'
'  outportb &H2E, &HAA

  If TimeType <> TIME_MINUTE Then
    TimeType = TIME_SECOND
  End If
    
  Timeout = Timeout And 255  ' max is 255
  
  
  outportb &H2E, &H87
  outportb &H2E, &H87
  outportb &H2E, &H7
  
  outportb &H2F, &H8
  
  outportb &H2E, &H30
  outportb &H2F, &H1
  
  outportb &H2E, &HF5
  outportb &H2F, &H0 ' Seconds or minutes 0 = seconds
  
  outportb &H2E, &HF6
  outportb &H2F, Timeout ' 15
  
  outportb &H2E, &HAA
  


End Sub


Public Function CheckWatchdog()
  Dim WDTO  As Long
  WDTO = Configuration.WatchdogTimeout * 0.3 ' trigger at 50% of requested time
  
    If DateDiff("s", LastWatchDog, Now) > WDTO Or CDbl(LastWatchDog) = 0 Then
        Select Case Configuration.WatchdogType
          Case WD_BERKSHIRE
            BerkshireWD.Tickle
          Case WD_UL
            SetWatchdog Configuration.WatchdogTimeout
          Case WD_ARK3510
            SetWatchdog Configuration.WatchdogTimeout
          Case Else
            ' do nothing
        End Select
      LastWatchDog = Now
  End If


End Function

Public Sub SetWatchdog(ByVal Timeout As Long)
  If (MASTER) And (Configuration.WatchdogType > 0) Then
    Select Case Configuration.WatchdogType
    Case WD_ARK3510
      ARK3510_WatchDogTimer Timeout, TIME_SECOND
    Case WD_UL
      WinbondWatchDogTimer Timeout, TIME_SECOND
    Case WD_BERKSHIRE
      
      'Set BershireWD = Nothing
    Case Else  ' WD_NONE
      ' do nothing
    End Select
  End If

End Sub

