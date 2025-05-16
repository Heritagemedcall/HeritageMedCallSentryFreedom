Attribute VB_Name = "modMain"
Option Explicit
Global USE6080             As Long  ' non-zero for new 6080
Global IP1                 As String
Global USER1               As String
Global PW1                 As String
Global UseSecureSockets    As Long
Global ZoneInfoList        As cZoneInfoList

Global Adapters            As cAdapters
Global DefaultAdapter      As String
Global DefaultIP           As String

Global RemoteStartUpTimer  As Date

Global MemCheck As cMemory

Global NoACG               As Boolean

Global i6080               As c6080
Global Enroller            As c6080

Global gSPForwardAccount   As String
Global gForwardSoftPoints  As Long

Global nDoEvents            As Long   ' used for iterations that take a long time, set in Sub Main
' as a variable for testing or changing in setup
Global Const DoEventYield   As Long = 25  ' yield to event que every nDoEvents count ' default

Global gLockTimeRemaining   As Long
Global gInactivityTimeRemaining   As Long
Global Const INACTIVITY_DELAY = 60 * 5
Global Const INACTIVITY_DELAY_FACTORY = 60 * 60
Global Const LOCK_DELAY = 30

' true if master , false if remote console
Global MASTER               As Boolean
Global gDirectedNetwork     As Boolean
Global gMyAlarms            As Long  ' 0 or 1

Global NODIVA               As Boolean

' Global Sendmail object
Global gSMTPMailer          As Object

' these are toggles for sending debug info to external debug window
Global gShowLocationData    As Boolean
Global gShowTAPIData        As Boolean
Global gShowHostRemoteData  As Boolean
Global gShowGeneralData     As Boolean
Global gShowPacketData      As Boolean

Global gExtendFactory       As Boolean

Global gNoStrayData         As Boolean  ' don't record stray data
Global gNoDataErrorLog      As Boolean  ' don't log BAD data
Global gLogTAP              As Boolean


'Global EthernetAdapter      As cEthernetAdapter
Global ConsoleID            As String
Global LoggedIn             As Boolean



Global gAllowedDeviceCount  As Long
Global gDaysLeft            As Long
Global gRegistered          As Boolean

Global bytespermin          As Double
Global packetspermin        As Double
Global lastupdate           As Date


Global ms                   As Long
Global TickCounter          As Long
Global HalfSecondCounter    As Long
Global SecondCounter        As Long
Global ThreeSecondCounter   As Long
Global MinuteCounter        As Long
Global TenMinuteCounter     As Long
Global ThirtyMinuteCounter  As Long

Global RemoteRefreshCounter As Long
Global RemoteRefresh_Delay  As Long  '  = 3

Global packetizer           As New cPacketizer

'Global RebuildQue           As cRebuildQue

Global SoundAlert()         As Byte
Global SoundAlarm()         As Byte
Global SoundTrouble()       As Byte
Global SoundLowBatt()       As Byte
Global SoundAssur()         As Byte

Global EventNames()
Global EventIDs()           As Long

Global EventTypes           As Collection

Global Configuration        As New cConfiguration

Global gAssurDisableScreenOutput As Integer

Global DeletingAlarm    As Boolean
Global gLogRawData      As Boolean
Global gWindow          As Integer
Global gLocateWindow    As Long  ' time to accept packets for location
Global gSupervisePeriod As Long
Global gSuperviseGroup  As Integer  ' 0 to 255
Global gLastSupervise   As Date
Global gSystemStart     As Date
Global gLastPacket      As Date

Global gSuspendPackets  As Boolean

Global gUser            As cUser

Global gTimeFormat       As Integer  ' 1 = 24 hr or 0 = am-pm
Global gTimeFormatString As String  ' default is : "hh:nnA/P" else, "hh:nn"
Global gElapsedEqACK     As Integer

Global Const USE_ELAPSED_ACK = 1

Global ActiveIDMs(0 To 255) As Boolean

Global gCurrentNID       As Integer  ' RX NID

Global StopIt            As Boolean
Global starttime         As Date
Global gSessionID        As Long

'Global gNewVersion       As Boolean

Global gLogDevice        As Integer

Global Const gCursorType = adOpenKeyset  'adOpenDynamic, adOpenStatic adOpenkeyst
Global Const gLockType = adLockOptimistic  'adLockPessimistic

Global InIDE             As Boolean

'Global Const BAD_FILE_CHARS as string = "\/:*?<>|" & """"

Global Const TRACE_FILE = 0
Global Const TRACE_TRACE = 1
Global Const TRACE_DEBUG = 2
Global Const TRACE_NONE = 3
Global Const TRACE_MSGBOX = 4

Global Const PCA_MODE = 0
Global Const TWO_BUTTON_MODE = 1
Global Const EN1221_MODE = 2

Global Const SURVEY_RC0 = 0
Global Const SURVEY_RC1 = 1
Global Const SURVEY_RC2 = 2
Global Const SURVEY_RC3 = 3


Global Const COM_DEV_NAME = "SERIAL-IN"
Global PCA_DEV_NAME         As String

' used to determine polling frequency
Global TimerPeriod        As Long
Global TimerStart         As Long

Global gTracing           As Boolean

Global conn               As ADODB.Connection
Global gHoldOff           As Boolean  ' stops the heart beat from doing ay work

Global gIsJET             As Boolean

Global gSentinel          As cSentinel

Global gShowLogonScreen    As Boolean

Global gDukane            As cDukane

Global LastMonitorPing    As Date
Global MonitorPinger      As cHTTPRequest

Global Const EVENT_FACILITY_NONE = 0
Global Const EVENT_FACILITY_STARTUP = 3
Global Const EVENT_FACILITY_SHUTDOWN = 4
Global Const EVENT_FACILITY_NETFAIL = 5
Global Const EVENT_FACILITY_NETRESTORE = 6

Global Const BAD_FILE_CHARS = "\/:*?<>|" & """"

Global PingMonitorBusy As Boolean








Sub PingMonitor(ByVal ForcePing As Boolean, Optional ByVal message As String = "")
  
  Dim READYSTATE As Long
  
  If PingMonitorBusy Then
    Exit Sub
  End If
  
  PingMonitorBusy = True
  
  If Configuration.MonitorEnabled Then
    
    On Error Resume Next
    
    If MonitorPinger Is Nothing Then
      Set MonitorPinger = New cHTTPRequest
    End If
    
    READYSTATE = MonitorPinger.READYSTATE
    
    'Debug.Print "HTTP Readystate " & READYSTATE

    If (DateAdd("s", Configuration.MonitorInterval, LastMonitorPing) <= Now) Or ForcePing Then
      
      If READYSTATE = 0 Or READYSTATE = 4 Then
      
      'Do Until MonitorPinger.READYSTATE = 0 Or MonitorPinger.READYSTATE = 4
      '  DoEvents
      'Loop
        Debug.Print " PingWatchdogHost " & Now
        MonitorPinger.PingWatchdogHost "http://" & Configuration.MonitorDomain, Configuration.MonitorRequest, Configuration.MonitorPort, message, "", ""
        LastMonitorPing = Now
      Else
        Debug.Print "Wait PingWatchdogHost " & Now
      
      End If
      
    End If
  End If
  PingMonitorBusy = False
End Sub



Sub InvalidateHostLogin()
  gShowLogonScreen = True
End Sub


Public Function ProcessLogin(ByVal Password As String) As Boolean
  'LEVEL_FACTORY = 256
  'LEVEL_ADMIN = 128
  'LEVEL_SUPERVISOR = 32
  'LEVEL_USER = 1



  Dim User               As cUser
  Dim Session            As cUser
  Dim j                  As Integer

  If MASTER Then
    Set User = GetUser(Password)
    If User.LEvel > 0 Then
      If User.LEvel = LEVEL_FACTORY Then
        For j = HostSessions.Count To 1 Step -1
          Set Session = HostSessions(j)
          If Session.LEvel >= LEVEL_SUPERVISOR Then
            If Session.Session <> User.Session Then
              dbg "Bumping Admin"
              HostSessions.Remove j
              LogRemoteSession Session.Session, 0, "Process Login Bumping Admin"
            End If
          End If
        Next
        dbg "Local Log on as Factory"
        Set gUser = User
        HostSessions.Add gUser
        ProcessLogin = True

      ElseIf User.LEvel > LEVEL_USER Then  ' two admin levels
        For j = HostSessions.Count To 1 Step -1
          Set Session = HostSessions(j)
          If Session.LEvel >= LEVEL_SUPERVISOR Then
            ProcessLogin = False
            Exit For
          End If
        Next

        If j = 0 Then
          dbg "Local Log on as Admin " & User.LEvel
          Set gUser = User
          HostSessions.Add gUser
          ProcessLogin = True
        Else
          ProcessLogin = False
        End If


      ElseIf User.LEvel = LEVEL_USER Then
        Set gUser = User
        HostSessions.Add gUser
        ProcessLogin = True
      End If

    Else
      ProcessLogin = False
    End If


    'NOT MASTER: this is REMOTE
  Else
    ' get user token back from Host
    ResetRemoteRefreshCounter (-5)

    Set User = GetUser(Password)
    ' factory must be able to login regardless if remote or not
    If User.LEvel = LEVEL_FACTORY Then
      Set gUser = User
      ProcessLogin = True
      'remote logoff admins

    Else
      'dbgHostRemote
      dbg "modmain.ProcessLogin call remotegetuser"
      Set User = RemoteGetUser(Password)
      ResetRemoteRefreshCounter (-5)
      'dbgHostRemote
      dbg "modmain.ProcessLogin, user.level " & User.LEvel & "  LoggedOn " & User.LoggedOn
      If User.LEvel > 0 Then
        If User.LoggedOn Then
          Set gUser = User
          If gUser.LEvel = LEVEL_FACTORY Then
            'remote logoff admins

          End If
          ProcessLogin = True
        Else
          ProcessLogin = False
        End If
      Else
        ProcessLogin = False
      End If
    End If
  End If

End Function


Private Function RunningInIDE() As Boolean
  'On Error Resume Next
  'Debug.Assert 1 / 0
  'RunningInIDE = Err.Number <> 0
  
  Dim filename      As String
  Dim Count         As Long
  On Error Resume Next
  
  filename = String(255, 0)
  Count = GetModuleFileName(App.hInstance, filename, 255)
  filename = left(filename, Count)
  filename = Right(filename, 7)
  If 0 = StrComp(filename, "VB6.EXE", vbTextCompare) Then
    RunningInIDE = True
  Else
    RunningInIDE = False
  End If


End Function
Sub Main()
        Dim s                  As String              ' just for comm open error

        Dim Adapter            As cAdapter
        Dim MasterIP           As String
        Dim ServiceName        As String



Set alarms = New cAlarms
Set Alerts = New cAlarms
Set Troubles = New cAlarms
Set LowBatts = New cAlarms
Set Assurs = New cAlarms
Set Externs = New cAlarms          ' external devices

        


10      MASTER = True
20      If InStr(1, App.exename, "REMOTE", vbTextCompare) Or InStr(1, Command$, "REMOTE", vbTextCompare) Then
30        MASTER = False
          
          
          
40      End If

        RemoteStartUpTimer = DateAdd("s", 60, Now)

50      If MASTER Then

60        On Error Resume Next

70        On Error GoTo 0
80        Set Enroller = New c6080
90        Set MonitorPinger = New cHTTPRequest
          Set MemCheck = New cMemory
          Set PushProcessor = New cPushProcessor
          'PushProcessor.SetAsyncControl frmTimer.AsyncReader1
          'Set frmTimer.AsyncReader1.ParentObject = PushProcessor
          
          gPush = IIf(Val(ReadSetting("Push", "Enabled", "0")) <> 0, 1, 0)
          
          
        Else ' Remote
        
100     End If
110     nDoEvents = DoEventYield


120     Randomize Timer
130     'LastWatchDog = Now
140     Set BerkshireWD = New cBerkshire

        'If InStr(1, Command$, "nodiva", vbTextCompare) Then
        
        If InStr(1, Command$, "TAPI", vbTextCompare) Then
          gShowTAPIData = True
        End If
        
        
150     NODIVA = True
        'End If
160     SetVars

170     GetConfig    ' gets configuration from inifile

180     If Not MASTER Then
          
          Dim oldMAC As cEthernetAdapterOLD
          Set oldMAC = New cEthernetAdapterOLD
          ConsoleID = oldMAC.MAC
          Set oldMAC = Nothing
          
290     End If


300     Set HostSessions = New Collection

310     Set gAutoReports = New Collection

320     TimerPeriod = 20
330     InIDE = RunningInIDE()

        'gTracing = True
340     gWindow = SAME_EVENT_PERIOD    '15  ' seconds where transmission would be the same event
350     gLocateWindow = LOCATOR_WAIT_TIME

360     If App.PrevInstance And MASTER Then
370       MsgBox "Program Already Running" & vbCrLf & "Switching to Running Instance", vbInformation, App.Title
380       End
390     End If

400     If Configuration.WatchdogType = WD_BERKSHIRE Then
410       BerkshireWD.InitWatchdog
420       BerkshireWD.Enable
430       BerkshireWD.Tickle
440     End If


450     Set gSentinel = New cSentinel
460     If Not GetLicensing() Then
470       MsgBox "Unlicensed Version Has Expired"
480     End If


490     Set ClientConnections = New Collection
500     Set WirelessPort = New cComm
        'Set EthernetAdapter = New cEthernetAdapter

        ' reminder system

510     Set gRemindersToSend = New Collection
520     Set EmailQue = New Collection
530     Set VoiceMailQue = New Collection
540     LeadTime = 0






550     frmSplash.Show vbModeless
560     frmSplash.Refresh
570     DoEvents
580     frmSplash.Refresh

        ' OK for REMOTE
590     GetLogDevice
600     frmSplash.progress.Value = 10

        ' OK for REMOTE
610     Init    ' collections and global variables
        ' OK for REMOTE


620     frmSplash.progress.Value = 20

        ' OK for REMOTE
630     If StartupConnect() Then    'Connect to database

640       frmSplash.progress.Value = 30
          'If gNewVersion Or InIDE Then
650       If MASTER Then
660         FixDatabase
670       End If

          ' clear out mobile alarms list
          ClearAllMobiles

680       GetGlobals    ' Those held in the database

690       If MASTER Then
700         VerifyPCAOutputServer
710       End If

720       frmSplash.progress.Value = 40

          ' OK for REMOTE
730       LoadSounds

740       gLastSupervise = DateAdd("n", 2, Now)    ' start gSupervisePeriod minutes from now

750       frmSplash.progress.Value = 50

          ' OK for REMOTE
760       ReadESDeviceTypes

770       frmSplash.progress.Value = 60

          ' OK for REMOTE (partial)
780       If MASTER Then
790         Set i6080 = New c6080



800         dbg "Next: ReadResidents"
810         ReadResidents
820         dbg "Next: ReadRooms"
830         ReadRooms
840         frmSplash.progress.Value = 62
850         dbg "Next: ReadDevices"
860         ReadDevices  ' get configured transmitters
            
870         frmSplash.progress.Value = 66
880         If USE6080 Then  ' remember this is only master
              'dbg "Next: Read 6080"

              'If Not (NoACG) Then
              '  i6080.Get6080Data
              'End If

              'Get6080Info

              'dbg "Next: Correlate"
              'CorrelateDevicesToZones
890         End If
900         frmSplash.progress.Value = 75
910         dbg "Next: KillExternalMessages"
920         KillExternalMessages

930       End If
940       frmSplash.progress.Value = 78



          ' OK for REMOTE
950       If USE6080 Then
960       Else
970         dbg "Next: FetchWaypoints"
980         FetchWaypoints
990       End If

1000      frmSplash.progress.Value = 80

          ' OK for REMOTE (partial)

1010      dbg "Next: InitPageDevices"

1020      If MASTER Then
1030        InitPageDevices
1040      End If

1050      frmSplash.progress.Value = 90

1060      gSystemStart = Now

1070      dbg "Next: StartSession"

1080      If MASTER Then
            SetPushMobileEntries
1090        StartSession  ' just adds an entry into the DB
1100      End If

1110      frmSplash.progress.Value = 100

1120      dbg "Next: PostEvent EVT_SYSTEM_START"

1130      If MASTER Then
1140        PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_START, 0
1150      Else
            'PostEvent Nothing, Nothing, Nothing, EVT_REMOTE_START, 0
1160      End If



1170      frmMain.Show vbModeless
1180      frmMain.Enabled = False
1190      frmSplash.Show vbModeless


          'moved here 1/28/07 for different startup sequence... was after starting timers
1200      If MASTER Then
1210        If USE6080 Then
1220          dbg "Next: Read 6080"

1230          i6080.Username = USER1
1240          i6080.Password = PW1
1250          i6080.useSSL = UseSecureSockets
1260          i6080.IP = IP1
1270          i6080.SetRequestString 0

1280          If InIDE Then
1290            NoACG = False
1300          Else
1310            If Not Ping(IP1) Then
1320              NoACG = True
1330              messagebox frmMain, "No Server on " & IP1, App.Title, vbInformation
1340            End If
1350          End If
1360          Err.Clear
1370          On Error Resume Next
1380          If Not NoACG Then
1390            If i6080.Get6080Data() Then
1400              i6080.Connect
1410            End If
1420          End If
1430          If Err.Number Then
1440            messagebox frmMain, "Could Not Connect to ACG at " & IP1 & vbCrLf & Err.Description & " " & Err.Number, App.Title, vbInformation
1450          End If
1460          Err.Clear
1470          On Error GoTo 0

1480        Else
1490          InitComm WirelessPort, Configuration.CommPort, "baud=9600 parity=N data=8 stop=1"
1500          s = WirelessPort.CommGetError
1510          If Len(s) Then  ' problem connecting to serial port
1520            messagebox frmMain, s, App.Title, vbInformation
1530          End If
1540        End If

1550        Set gDukane = New cDukane
1560        gDukane.Init

1570      End If

          ' Master/Host uses this!
1580      Set RemoteAutoEnroller = New cRemoteAutoEnroller

1590      Load frmTimer
1600      frmTimer.StartTimer


1610      If MASTER Then

            '      Dim w              As Object
            '      Set w = frmTimer.WinsockHost

1620        On Error Resume Next

1630        BootLog "----------- Session Start " & Now & "-----------"

1640        Set Adapters = Nothing
1650        Set Adapters = New cAdapters



1660        Adapters.RefreshAdapters

1670        BootLog "Adpaters Count: " & Adapters.Count & " #" & Err.Number


            ' get settings for adapter by GUID

1680        ServiceName = Trim$(ReadSetting("Adapter", "ServiceName", ""))

1690        BootLog "Default SN: " & ServiceName & " #" & Err.Number

1700        Set Adapter = Adapters.GetAdapterByServiceName(ServiceName)

1710        BootLog "Adapter MAC: " & Adapter.MacAddress & " #" & Err.Number

1720        MasterIP = "127.0.0.1"  ' default to loopback
1730        If Not (Adapter Is Nothing) Then
1740          MasterIP = Adapter.DhcpIPAddress
1750          BootLog "DhcpIPAddress: " & Adapter.DhcpIPAddress & " #" & Err.Number
1760        End If


            Dim IPOrig         As String

1770        BootLog "MasterIP: " & MasterIP


1780        If MasterIP = "127.0.0.1" Or MasterIP = "0.0.0.0" Then
1790          IPOrig = frmTimer.WinsockHost(0).LocalIP
1800          BootLog "WinsockHost IP: " & IPOrig & " #" & Err.Number
1810          If IPOrig <> "0.0.0.0" Or IPOrig <> "127.0.0.1" Then

1820            MasterIP = IPOrig



1830            Set Adapter = Adapters.GetAdapterByIP(MasterIP)



1840            If Not Adapter Is Nothing Then
1850              BootLog "Adapter " & "MasterIP " & Adapter.DhcpIPAddress & " #" & Err.Number
1860              WriteSetting "Adapter", "MasterIP", Adapter.DhcpIPAddress
1870              BootLog "Adapter " & "MAC " & Adapter.MacAddress & " #" & Err.Number
1880              WriteSetting "Adapter", "MAC", Adapter.MacAddress
1890              BootLog "Adapter " & "AdapterName " & Adapter.Description & " #" & Err.Number
1900              WriteSetting "Adapter", "AdapterName", Adapter.Description
1910              BootLog "Adapter " & "ServiceName " & Adapter.ServiceName & " #" & Err.Number

1920              WriteSetting "Adapter", "ServiceName", Adapter.ServiceName
1930            Else
1940              BootLog "Adapter not available via " & MasterIP & " #" & Err.Number
1950            End If
1960          End If
1970        End If



1980        frmTimer.WinsockHost(0).Close
1990        BootLog "WinsockHost Closed"
2000        frmTimer.WinsockHost(0).Bind Configuration.HostPort, MasterIP
2010        BootLog "WinsockHost Bind " & Configuration.HostPort & " " & MasterIP
            'frmTimer.WinsockHost(0).Close
            'frmTimer.WinsockHost(0).Bind Configuration.HostPort, Configuration.HostIP

            'frmTimer.WinsockHost(0).LocalPort = Configuration.HostPort
2020        frmTimer.WinsockHost(0).Listen
2030        BootLog "WinsockHost Listening"

2040        On Error GoTo 0

2050      Else   ' not a master, but a remote
2060        Set HostConnection = New cHostConnection
2070        Set HostConnection.Socket = frmTimer.WinsockClient
2080        HostConnection.Connect Configuration.HostIP, Configuration.HostPort

2090        Set HostInterraction = New cHostConnection
2100        Set HostInterraction.Socket = frmTimer.WinsockClientInterraction
2110        HostInterraction.Connect Configuration.HostIP, Configuration.HostPort
2120      End If
          'Debug.Print "frmTimer.WinsockHost(0).LocalPort", frmTimer.WinsockHost(0).LocalPort
          'Debug.Print "frmTimer.WinsockHost(0).LocalIP ", frmTimer.WinsockHost(0).LocalIP
          'Debug.Print "frmTimer.WinsockHost(0).state", frmTimer.WinsockHost(0).State

          'moved here 11/17/06 for different startup sequence... was after setlisttabs in frmMain Load
          'InitComm WirelessPort, Configuration.CommPort, "baud=9600 parity=N data=8 stop=1"
2130      If MASTER Then
2140        If (USE6080 = 0) Then

2150          GetNCNID
2160          If gDirectedNetwork Then
2170            Configuration.RxSerial = GetNCSerial()
2180          End If

2190        End If

2200        LoadAutoReports
2210        LoadAutoExReports
2220        If (Configuration.MonitorInterval > 0) And (Configuration.MonitorEnabled = 1) Then
2230          If PingMonitorBusy = False Then
2240            PingMonitor True, "facilityid=" & Configuration.MonitorFacilityID & "&" & "eventcode=" & EVENT_FACILITY_STARTUP
2250          End If
2260        End If
2270      End If
2280      Unload frmSplash
2290      LoggedIn = False
          
2300      frmMain.DoLogin "0000"
          frmMain.TimerLogon.Enabled = True
2310    Else
2320      MsgBox "Failed to Connect to Database" & vbCrLf & "Please Check Your Database Settings.", vbInformation Or vbCritical
2330      Unload frmSplash
2340      DestroyObjects
2350      End
2360    End If


End Sub


Function StartupConnect() As Boolean
  On Error Resume Next
  Connect
  StartupConnect = (Err.Number = 0)
End Function
Function GetLicensing() As Boolean
  gAllowedDeviceCount = GetAllowedDeviceCount()


  If MASTER Then
    If (gAllowedDeviceCount > 0) Then  ' we have a valid count, good to go
      gRegistered = True
      GetLicensing = True
      gDaysLeft = -1
    ElseIf gSentinel.Expired Then  ' expired, double check
      gRegistered = False
      GetLicensing = gAllowedDeviceCount > 0
      gDaysLeft = -1
    Else                       ' the trial run
      gRegistered = False
      gDaysLeft = gSentinel.DaysLeft  ' how many days left
      GetLicensing = True
      gAllowedDeviceCount = 32000
    End If
  Else

    If (gAllowedDeviceCount = -1) Then  ' we have a valid count, good to go
      gRegistered = True
      GetLicensing = True
      gDaysLeft = -1
    ElseIf gSentinel.Expired Then  ' expired, double check
      gRegistered = False
      GetLicensing = gAllowedDeviceCount > 0
      gDaysLeft = -1
    Else                       ' the trial run
      gRegistered = False
      gDaysLeft = gSentinel.DaysLeft  ' how many days left
      GetLicensing = True
      gAllowedDeviceCount = 32000
    End If
  End If

End Function
Function GetAllowedDeviceCount() As Long
  Dim AllowedDevices As Long
  AllowedDevices = gSentinel.GetDeviceCount
  
  GetAllowedDeviceCount = Min(32000, AllowedDevices)
End Function

Function VerifyPCAOutputServer() As Boolean
  Dim SQL           As String
  Dim rs            As Recordset
  Dim IsPCAServer   As Boolean

  ' we need to create an entry in pagerdevices for the PCA Always.
  ' so the field defs may need to be updated

  SQL = "Select count(*) FROM PagerDevices WHERE ProtocolID = -1 AND Deleted = 0 "
  Set rs = ConnExecute(SQL)
  IsPCAServer = (rs(0) > 0)
  rs.Close
  If Not (IsPCAServer) Then
    SQL = "INSERT INTO PagerDevices (Description,Port,BaudRate,Bits,Parity,StopBits,Flow,Settings,AudioDevice,ProtocolID,Pause,Deleted,IncludePhone,Pin,KeyPA,Twice,DialerVoice,DialerModem,dialerphone,dialertag,dialermsgdelay,dialermsgrepeats,dialermsgspacing,dialerredials,dialerredialdelay,DialerAckDigit,Marquiscode,relay1,relay2,relay3,relay4,relay5,relay6,relay7,relay8,LF) " & _
          " Values ('PCA',0,0,0,'','',0,'','',-1,0,0,0,'',0,0,'',0,'','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)"
    ConnExecute SQL
  End If
  Set rs = Nothing



End Function

Function UpdateWaypoints()
  Dim j             As Integer
  Dim rs            As Recordset
  Dim SQL           As String
  Dim Match         As Boolean
  Dim w             As cWayPoint

  For j = 1 To Waypoints.Count
    Set w = Waypoints.waypoint(j)
    SQL = "SELECT count(id) FROM Waypoints where ID = " & w.ID
    ConnExecute SQL
    Match = (rs(0) <> 0)
    rs.Close
    If Not Match Then
      ' oops!
    Else
      SQL = "Update Waypoints Set repeater1 = " & q(w.Repeater1) & _
            ", repeater2 = " & q(w.Repeater2) & _
            ", repeater3 = " & q(w.Repeater3) & _
            ", signal1 = " & q(w.Signal1) & _
            ", signal2 = " & q(w.Signal2) & _
            ", signal3 = " & q(w.Signal3)
      ConnExecute SQL
      w.Checked = True
    End If

  Next
End Function

Function UpdateWaypoint(waypoint As cWayPoint)

  Dim rs            As Recordset
  Dim SQL           As String
  Dim Match         As Boolean

  SQL = "SELECT count(id) FROM Waypoints where ID = " & waypoint.ID
  Set rs = ConnExecute(SQL)
  Match = (rs(0) <> 0)
  rs.Close
  If Not Match Then
    ' oops!
  Else
    SQL = "Update Waypoints Set " & _
          "  repeater1 = " & q(waypoint.Repeater1) & _
          ", repeater2 = " & q(waypoint.Repeater2) & _
          ", repeater3 = " & q(waypoint.Repeater3) & _
          ", signal1 = " & q(waypoint.Signal1) & _
          ", signal2 = " & q(waypoint.Signal2) & _
          ", signal3 = " & q(waypoint.Signal3) & _
          " WHERE ID = " & waypoint.ID
    ConnExecute SQL
    waypoint.Checked = True
  End If


End Function


Function FetchWaypoints() As Boolean
  Dim rs            As Recordset
  Dim SQL           As String
  Dim ID            As Long

  Waypoints.ClearChecked
  SQL = "SELECT * FROM Waypoints order by ID"
  Set rs = ConnExecute(SQL)
  Do Until rs.EOF
    ID = rs("ID")
    If (Not Waypoints.Exists(ID)) Then
      DoEvents
      Waypoints.AddWayPoint rs
    End If
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Waypoints.RemoveDeadWood



End Function

Sub GetLogDevice()
  Dim s             As String
  Dim commands()    As String
  Dim j             As Integer
  Dim NVP()         As String
  s = Trim(Command$)
  commands = Split(s, " ")

  For j = LBound(commands) To UBound(commands)
    If InStr(1, commands(j), "Debug", vbTextCompare) Then
      NVP = Split(commands(j), "=")
      If UBound(NVP) >= 1 Then
        gLogDevice = Val(NVP(1))
      End If
    End If
    Exit For
  Next

End Sub

Sub Init()
  'gTracing = True
  If gDirectedNetwork Then
    LocatorWaitTime = 6
  Else
    LocatorWaitTime = 6
  End If
  Set gUser = New cUser

  SetDeviceTypes

  Set Packets = New Collection
  Set Devices = New cESDevices
  Set alarms = New cAlarms
  Set Alerts = New cAlarms
  Set LowBatts = New cAlarms
  Set Troubles = New cAlarms
  Set Assurs = New cAlarms
  Set SerialIns = New Collection
  Set Externs = New cAlarms

  Set Waypoints = New cWaypoints
  Set Outbounds = New cOutBounds
  Set InBounds = New cAlarms
  InBounds.name = "InBounds"

  alarms.ManualClear = 1
  Alerts.ManualClear = 1

  CreateEventtypes

End Sub
Sub SetVars()
  CHAR_NUL = Chr(0)
  CHAR_SOH = Chr(1)
  CHAR_STX = Chr(2)
  CHAR_ETX = Chr(3)
  CHAR_EOT = Chr(4)
  CHAR_ENQ = Chr(5)
  CHAR_ACK = Chr(6)
  CHAR_BEL = Chr(7)
  CHAR_BS = vbBack
  CHAR_HT = vbTab
  CHAR_VT = Chr(&HB)
  CHAR_FF = Chr(&HC)
  CHAR_CR = vbCr
  CHAR_SO = Chr(&HE)
  CHAR_SI = Chr(&HF)
  CHAR_LF = vbLf
  CHAR_XOFF = Chr(&H11)
  CHAR_XON = Chr(&H13)
  CHAR_NAK = Chr(&H15)
  CHAR_ETB = Chr(&H17)
  CHAR_SUB = Chr(&H1A)
  CHAR_ESC = Chr(&H1B)
  CHAR_RS = Chr(&H1E)
  CHAR_US = Chr(&H1F)
  CHAR_DEL = Chr(&H7F)
  '
  CHAR_SUB_CR = CHAR_SUB & "M"
  CHAR_SUB_LF = CHAR_SUB & "J"
  CHAR_CRLF = CHAR_CR & CHAR_LF

End Sub

Sub GetGlobals()
  Dim rs            As Recordset


  Set rs = ConnExecute("SELECT * FROM GlobalSettings")
  If Not rs.EOF Then
    gTimeFormat = IIf(rs("MilitaryTime") = 1, 1, 0)
    gElapsedEqACK = IIf(rs("ElapsedEqACK") = 1, 1, 0)
    gAssurDisableScreenOutput = IIf(rs("DisableScreenOutput") = 1, 1, 0)
  End If

  rs.Close
  Set rs = Nothing
  If gTimeFormat = 1 Then
    gTimeFormatString = "hh:nn"
  Else
    gTimeFormatString = "hh:nnA/P"
  End If

End Sub

Function CorrelateDevicesToZones() As Long
  Dim j             As Long
  Dim AllZones      As Collection
  Dim d             As cESDevice
  Dim z             As cZoneInfo
  Dim Serial        As String

  Set AllZones = ZoneInfoList.ZoneList
  For j = 1 To AllZones.Count
    DoEvents
    Set z = AllZones(j)
    Debug.Print "Zone "; j, " pti "; z.PTI; " "; z.TypeName
    Serial = Right$("00" & z.MID, 2) & Right$("000000" & z.HexID, 6)

    On Error Resume Next
    Set d = Devices.Device(Serial)
    If d Is Nothing Then
      z.Validated = 0
    Else
      z.Validated = 1
      d.ZoneID = z.ID
      d.Validated = 1
    End If
  Next



End Function

Sub GetConfig()
  'Dim Major   As Long
  'Dim Minor     As Long
  'Dim Build     As Long

  ' get last build version Run
  'Major = Val(ReadSetting("Version", "Major", "0"))
  'Minor = Val(ReadSetting("Version", "Minor", "0"))
  'Build = Val(ReadSetting("Version", "Build", "0"))





  On Error GoTo GetConfig_Error


  USE6080 = IIf(Val(ReadSetting("Configuration", "USE6080", 0)), 1, 0)
  IP1 = ReadSetting("Configuration", "IP06080", "192.168.60.80")
  USER1 = ReadSetting("Configuration", "U06080", "Admin")
  PW1 = ReadSetting("Configuration", "P06080", "Admin")

  If USER1 = "Admin" Then
    ' no decrypt
  Else
    USER1 = MakeDeCryptedString(USER1)
  End If

  If PW1 = "Admin" Then
    ' no decrypt
  Else
    PW1 = MakeDeCryptedString(PW1)
  End If

  ' USER1 = "Admin"
  ' PW1 = "Admin"


  gForwardSoftPoints = Val(ReadSetting("Configuration", "ForwardSoftPoints", "0")) And 1

  gSPForwardAccount = Trim$(ReadSetting("Configuration", "SPForwardAccount", ""))
  If Len(gSPForwardAccount) = 0 Then
    gForwardSoftPoints = 0
  End If


  UseSecureSockets = IIf(Val(ReadSetting("Configuration", "SSL6080", "0")) = 1, 1, 0)

  gMyAlarms = IIf(Val(ReadSetting("Remote", "MyAlarms", 0)), 1, 0)

  gLoBattDelay = Max(0, Val(ReadSetting("Configuration", "LowBattDelay", 0)))


  gSupervisePeriod = ReadSetting("Configuration", "SupervisePeriod", 10)
  RemoteRefresh_Delay = Val(ReadSetting("Configuration", "RemoteRefresh_Delay", "5"))  ' seconds for fetching alarms

  gNoStrayData = CBool(ReadSetting("Debug", "NoStrayData", "false"))
  gNoDataErrorLog = CBool(ReadSetting("Debug", "NoDataErrorLog", "false"))

  gExtendFactory = CBool(ReadSetting("Debug", "ExtendFactory", "false"))

  gLogTAP = CBool(ReadSetting("Debug", "LogTAP", "false"))

  Configuration.Facility = ReadSetting("Configuration", "Facility", "<New>")
  Configuration.RxTimeout = Val(ReadSetting("Configuration", "RxTimeout", 60 * 12))
  'Configuration.ID = Val(ReadSetting("Configuration", "ID", Configuration.ID))
  Configuration.CommPort = Val(ReadSetting("Configuration", "CommPort", "1"))

  Configuration.HostPort = Val(ReadSetting("Configuration", "HostPort", "2500"))
  Configuration.HostIP = ReadSetting("Configuration", "HostIP", "127.0.0.1")

  Configuration.ReportPath = ReadSetting("Configuration", "ReportPath", App.Path)
  If Right(Configuration.ReportPath, 1) <> "\" Then
    Configuration.ReportPath = Configuration.ReportPath & "\"
  End If

  Configuration.EscTimer = Val(ReadSetting("Configuration", "EscTimer", 60))


  Configuration.AlarmFile = ReadSetting("Configuration", "AlarmFile", Configuration.AlarmFile)
  Configuration.AlertFile = ReadSetting("Configuration", "AlertFile", Configuration.AlertFile)
  Configuration.LowBattFile = ReadSetting("Configuration", "LowBattFile", Configuration.LowBattFile)
  Configuration.TroubleFile = ReadSetting("Configuration", "TroubleFile", Configuration.TroubleFile)
  Configuration.AssurFile = ReadSetting("Configuration", "assurFile", Configuration.AssurFile)
  Configuration.ExtFile = ReadSetting("Configuration", "ExtFile", Configuration.ExtFile)

  Configuration.locationtext = Val(ReadSetting("Configuration", "LocationText", "0")) And 1



  Configuration.AlarmBeep = Val(ReadSetting("Configuration", "Alarmbeep", Configuration.AlarmBeep))
  Configuration.AlertBeep = Val(ReadSetting("Configuration", "Alertbeep", Configuration.AlertBeep))
  Configuration.LowBattBeep = Val(ReadSetting("Configuration", "LowBattbeep", Configuration.LowBattBeep))
  Configuration.TroubleBeep = Val(ReadSetting("Configuration", "Troublebeep", Configuration.TroubleBeep))
  Configuration.AssurBeep = Val(ReadSetting("Configuration", "AssurBeep", Configuration.AssurBeep))
  Configuration.ExtBeep = Val(ReadSetting("Configuration", "ExtBeep", Configuration.ExtBeep))


  Configuration.AlarmReBeep = Val(ReadSetting("Configuration", "AlarmRebeep", Configuration.AlarmReBeep))
  Configuration.AlertReBeep = Val(ReadSetting("Configuration", "AlertRebeep", Configuration.AlertReBeep))
  Configuration.LowBattReBeep = Val(ReadSetting("Configuration", "LowBattRebeep", Configuration.LowBattReBeep))
  Configuration.TroubleReBeep = Val(ReadSetting("Configuration", "TroubleRebeep", Configuration.TroubleReBeep))
  Configuration.AssurReBeep = Val(ReadSetting("Configuration", "AssurReBeep", Configuration.AssurReBeep))
  Configuration.ExtReBeep = Val(ReadSetting("Configuration", "ExtReBeep", Configuration.ExtReBeep))

  Configuration.BeepControl = Val(ReadSetting("Configuration", "BeepControl", 0))



  Configuration.AssurStart = Val(ReadSetting("Configuration", "AssurStart", "0"))
  Configuration.AssurEnd = Val(ReadSetting("Configuration", "AssurEnd", "0"))



  Configuration.AssurStart2 = Val(ReadSetting("Configuration", "AssurStart2", "0"))
  Configuration.AssurEnd2 = Val(ReadSetting("Configuration", "AssurEnd2", "0"))

  Configuration.StartNight = Val(ReadSetting("Configuration", "StartNight", "0"))
  Configuration.EndNight = Val(ReadSetting("Configuration", "EndNight", "0"))
  Configuration.EndThird = Val(ReadSetting("Configuration", "EndThird", Configuration.EndNight))


  Configuration.ESLastMessage = ReadSetting("Configuration", "ESLastMessage", 1)

  Configuration.SurveyDevice = ReadSetting("Configuration", "SurveyDevice", "00000000")
  Configuration.SurveyPCA = ReadSetting("Configuration", "SurveyPCA", "00000000")

  Configuration.WaypointDevice = ReadSetting("Configuration", "WaypointDevice", "00000000")

  Configuration.boost = Min(BOOST_LIMIT, Val(ReadSetting("Configuration", "Boost", "0")))

  Configuration.surveymode = Val(ReadSetting("Configuration", "SurveyMode", "0"))
  Configuration.SurveyPager = Val(ReadSetting("Configuration", "SurveyPager", "0"))
  Configuration.OnlyLocators = 0

  Configuration.NoNCs = Val(ReadSetting("Configuration", "NoNCs", "0"))

  Configuration.RxSerial = ReadSetting("Configuration", "RxSerial", "00000000")
  Configuration.RxLocation = ReadSetting("Configuration", "RxLocation", "Receiver")

  Configuration.BackupDOW = Val(ReadSetting("Backup", "DOW", "1"))
  Configuration.BackupDOM = Val(ReadSetting("Backup", "DOM", "1"))



  Configuration.BackupTime = Val(ReadSetting("Backup", "Time", "100"))
  Configuration.BackupEnabled = Val(ReadSetting("Backup", "Enabled", "0"))
  Configuration.BackupFolder = ReadSetting("Backup", "Folder", "")
  Configuration.BackupType = Val(ReadSetting("Backup", "Type", "0"))

  Configuration.BackupDOWRemote = Val(ReadSetting("RemoteBackup", "DOW", "1"))
  Configuration.BackupDOMRemote = Val(ReadSetting("RemoteBackup", "DOM", "1"))



  Configuration.BackupTimeRemote = Val(ReadSetting("RemoteBackup", "Time", "100"))
  Configuration.BackupEnabledRemote = Val(ReadSetting("RemoteBackup", "Enabled", "0"))
  Configuration.BackupFolderRemote = ReadSetting("RemoteBackup", "Folder", "")
  Configuration.BackupTypeRemote = Val(ReadSetting("RemoteBackup", "Type", "0"))

  Configuration.BackupHost = ReadSetting("RemoteBackup", "Host", "")
  Configuration.BackupUser = ReadSetting("RemoteBackup", "User", "")
  Configuration.BackupPassword = UnScramble(ReadSetting("RemoteBackup", "Password", ""))


  gDirectedNetwork = CBool(ReadSetting("Configuration", "DNet", False))


  Configuration.AssurSaveAsFile = IIf(Val(ReadSetting("Assurance", "SaveAsFile", "0")) <> 0, 1, 0)
  Configuration.AssurSendAsEmail = IIf(Val(ReadSetting("Assurance", "SendAsEmail", "0")) <> 0, 1, 0)
  Configuration.AssurFileFormat = Val(ReadSetting("Assurance", "FileFormat", "0"))
  Configuration.AssurEmailRecipient = Trim$(ReadSetting("Assurance", "EmailRecipient", ""))
  Configuration.AssurEmailSubject = Trim$(ReadSetting("Assurance", "EmailSubject", "Assurance Report"))





  ' mail
  Configuration.UseSMTP = ReadSetting("Configuration", "UseSMTP", "0")
  Configuration.MailUserName = ReadSetting("Configuration", "MailUserName", "")
  Configuration.MailSMTPserver = ReadSetting("Configuration", "MailSMTPserver", "")
  Configuration.MailPOP3Server = ReadSetting("Configuration", "MailPOP3Server", "")
  Configuration.MailSenderEmail = ReadSetting("Configuration", "MailSenderEmail", "")
  Configuration.MailPassword = UnScramble(ReadSetting("Configuration", "MailPassword", ""))
  Configuration.MailSenderName = ReadSetting("Configuration", "MailSenderName", "")
  Configuration.MailRequirePopLogin = Val(ReadSetting("Configuration", "MailRequirePopLogin", "0"))
  Configuration.MailRequireLogin = Val(ReadSetting("Configuration", "MailRequireLogin", "0"))
  Configuration.MailDebug = Val(ReadSetting("Configuration", "MailDebug", "0"))
  Configuration.MailPort = Min(Val(ReadSetting("Configuration", "MailPort", "25")), 99999)

  Configuration.ReminderMsgDelay = Val(ReadSetting("Reminders", "MsgDelay", "0"))
  Configuration.ReminderMsgRepeats = Val(ReadSetting("Reminders", "MsgRepeats", "3"))
  Configuration.ReminderMsgSpacing = Val(ReadSetting("Reminders", "MsgSpacing", "2"))
  Configuration.ReminderRedials = Val(ReadSetting("Reminders", "Redials", "3"))
  Configuration.ReminderRedialDelay = Val(ReadSetting("Reminders", "RedialDelay", "10"))
  Configuration.ReminderAckDigit = Val(ReadSetting("Reminders", "AckDigit", "48"))  ' the 0 on the keypad

  Configuration.WatchdogTimeout = Val(ReadSetting("Configuration", "WDtimeout", "90"))  ' UL standard is 90 sec
  Configuration.WatchdogType = Val(ReadSetting("Configuration", "WDtype", "0"))  ' new systems might not have WD

  Configuration.MonitorDomain = Trim$(ReadSetting("Monitoring", "Domain", ""))
  Configuration.MonitorRequest = Trim$(ReadSetting("Monitoring", "Request", ""))
  Configuration.MonitorInterval = Val(ReadSetting("Monitoring", "Interval", "90"))
  Configuration.MonitorPort = Val(ReadSetting("Monitoring", "Port", "80"))
  Configuration.MonitorEnabled = Val(ReadSetting("Monitoring", "Enabled", 0)) And 1
  Configuration.MonitorFacilityID = ReadSetting("Monitoring", "FacilityID", "")

  Configuration.RemoteSerial = ReadSetting("Configuration", "RemoteSerial", "00000000")
  
  
  Configuration.AdminContact = Trim$(ReadSetting("Configuration", "AdminContact", ""))

  Configuration.HideHIPPANames = Val(ReadSetting("Configuration", "HIPPAHideNames", 0)) And 1
  Configuration.HideHIPPASidebar = Val(ReadSetting("Configuration", "HideHIPPASideBar", 1)) And 1

  NODIVA = IIf(Val(ReadSetting("Configuration", "Dialogic", False)), True, False)  ' force to true if problematic .. must hand edit INI file

  ' Auto Report List Printing

  Configuration.AutoReportsListFolder = ReadSetting("Configuration", "AutoReportsListFolder", App.Path)
  Configuration.AutoReportsListPrinter = ReadSetting("Configuration", "AutoReportsListPrinter", GetPrinterDeviceName)

  ' Mobile
  
  Configuration.MobilehtPasswordPath = ReadSetting("Mobile", "PasswordPath", "")
  Configuration.MobilehtPasswordEXEPath = ReadSetting("Mobile", "PasswordEXEPath", "")
  Configuration.MobileWebRoot = ReadSetting("Mobile", "Root", "")
  Configuration.MobileWebEnabled = Val(ReadSetting("Mobile", "Enabled", "0"))
    
    
  
  Configuration.MobileClearAssist = Val(ReadSetting("Mobile", "ClearAssist", "60"))
  Configuration.MobileClearHistory = Val(ReadSetting("Mobile", "ClearHist", "60"))




GetConfig_Resume:
  On Error GoTo 0
  Exit Sub

GetConfig_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.GetConfig." & Erl
  Resume GetConfig_Resume


End Sub

Public Function GetPrinterDeviceName() As String
  Dim PrinterName   As String
  On Error Resume Next
  PrinterName = Printer.DeviceName
  If PrinterName = "" Then
    MsgBox "Error Finding Default Printer." & vbCrLf & "Please Restart Program After Installing a Printer."
  End If
  GetPrinterDeviceName = PrinterName

End Function

Sub LoadSounds()


10      On Error GoTo LoadSounds_Error

20      If FileExists(Configuration.AlarmFile) Then
30        SoundAlarm = GetWaveData(Configuration.AlarmFile)
40      Else
50        Configuration.AlarmFile = ""
60        SoundAlarm = LoadResData(101, "SOUND")
70      End If
80      alarms.DefaultBeepTime = Configuration.AlarmBeep
90      alarms.ReBeepTime = Configuration.AlarmReBeep

100     If FileExists(Configuration.AssurFile) Then
110       SoundAssur = GetWaveData(Configuration.AssurFile)
120     Else
130       Configuration.AssurFile = ""
140       SoundAssur = LoadResData(101, "SOUND")
150     End If
160     Assurs.DefaultBeepTime = Configuration.AssurBeep

170     Assurs.ReBeepTime = Configuration.AssurReBeep

180     If FileExists(Configuration.AlertFile) Then
190       SoundAlert = GetWaveData(Configuration.AlarmFile)
200     Else
210       Configuration.AlertFile = ""
220       SoundAlert = LoadResData(101, "SOUND")
230     End If
240     Alerts.DefaultBeepTime = Configuration.AlertBeep

250     Alerts.ReBeepTime = Configuration.AlertReBeep

260     If FileExists(Configuration.ExtFile) Then
270       SoundAlert = GetWaveData(Configuration.ExtFile)
280     Else
290       Configuration.ExtFile = ""
300       SoundAlert = LoadResData(101, "SOUND")
310     End If
320     Externs.DefaultBeepTime = Configuration.ExtBeep

330     Externs.ReBeepTime = Configuration.ExtReBeep

340     If FileExists(Configuration.TroubleFile) Then
350       SoundTrouble = GetWaveData(Configuration.AlarmFile)
360     Else
370       Configuration.TroubleFile = ""
380       SoundTrouble = LoadResData(101, "SOUND")
390     End If

400     Troubles.DefaultBeepTime = Configuration.TroubleBeep

410     Troubles.ReBeepTime = Configuration.TroubleReBeep

420     If FileExists(Configuration.LowBattFile) Then
430       SoundLowBatt = GetWaveData(Configuration.AlarmFile)
440     Else
450       Configuration.LowBattFile = ""
460       SoundLowBatt = LoadResData(101, "SOUND")
470     End If

480     LowBatts.DefaultBeepTime = Configuration.LowBattBeep

490     LowBatts.ReBeepTime = Configuration.LowBattReBeep

LoadSounds_Resume:
500     On Error GoTo 0
510     Exit Sub

LoadSounds_Error:

520     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.LoadSounds." & Erl
530     Resume LoadSounds_Resume


End Sub

'Public Function ReadBinaryData(ByVal filename As String, Data As String)
'  Dim hfile As Integer
'  hfile = FreeFile
'  On Error Resume Next
'  Open Configuration.AlarmFile For Binary Access Read As #hfile
'  Data = String(LOF(hfile), vbNullChar)
'  'ReDim Data(0 To LOF(hFile) - 1) As Byte
'  Get #hfile, , Data
'  Close #hfile
'  ReadBinaryData = Err.Number = 0
'End Function

Sub KillExternalMessages()
  On Error Resume Next
  Dim SQL           As String
  SQL = "DELETE FROM ExternalPages"
  ConnExecute SQL

End Sub

Sub ReadResidents()
  Dim rs            As Recordset
  Dim SQL           As String
  Dim Device        As cESDevice
  Dim Count         As Long

  Dim t             As Long: t = Win32.timeGetTime

  If Residents Is Nothing Then
    Set Residents = New cResidents
  End If

  Residents.ClearAll

  SQL = "SELECT * FROM Residents WHERE deleted = 0"
  Set rs = ConnExecute(SQL)
  Do Until rs.EOF
    Count = Count + 1
    Residents.ParseAndAdd rs
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Debug.Print Count & " Residents Loaded in " & Win32.timeGetTime - t
End Sub

Sub ReadRooms()
  Dim rs            As Recordset
  Dim SQL           As String
  Dim Device        As cESDevice
  Dim Count         As Long

  Dim t             As Long: t = Win32.timeGetTime

  If Rooms Is Nothing Then
    Set Rooms = New cResidents
  End If

  Rooms.ClearAll

  SQL = "SELECT * FROM Rooms WHERE deleted = 0"
  Set rs = ConnExecute(SQL)
  Do Until rs.EOF
    Count = Count + 1
    Rooms.ParseAndAdd rs
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Debug.Print Count & " Rooms Loaded in " & Win32.timeGetTime - t
End Sub



Sub ReadDevices()
  Dim rs                 As Recordset
  Dim SQL                As String
  Dim Device             As cESDevice
  Dim Count              As Long

  Dim t                  As Long

  If MASTER Then ' only ever called by master

    ' clear all alarms
    ' serial port this machine
    Set Device = New cESDevice
    Device.Serial = "00000000"
    Device.Description = "Receiver"
    Device.SupervisePeriod = Configuration.RxTimeout

    Devices.AddDevice Device
    Device.Serial = Configuration.RxSerial
    Device.Description = Configuration.RxLocation

    ' the rest of the devices




    SQL = "SELECT * FROM Devices WHERE devices.Deleted = 0"
    Set rs = ConnExecute(SQL)
    Do Until rs.EOF
      If IsNull(rs("serial")) = False Then
        Count = Count + 1
        If Count > gAllowedDeviceCount Then
          MsgBox "There Are More Devices Than Licensed" & vbCrLf & "Allowed:" & gAllowedDeviceCount & vbCrLf & "Loaded:" & Count, vbCritical, "Licensing Error"
          Exit Do
        Else
          ResetDevice rs, True
        End If
      End If
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

  End If

End Sub

Function GetLastPort() As Integer
  Dim cp
  On Error Resume Next
  Open "PortSet" For Input As #11
  Input #11, cp
  Close #11
  cp = Fix(Val(cp & ""))
  If cp < 1 Or cp > 255 Then
    cp = 1
  End If
  GetLastPort = cp
End Function

Public Sub LogProgramError(ByVal ErrorString As String)
  Select Case gLogDevice
    Case TRACE_MSGBOX
      MsgBox ErrorString
    Case TRACE_FILE
      On Error Resume Next
      Dim hfile     As Integer
      Dim filename As String
      filename = App.Path & "\Err.log"
      limitFileSize filename
      hfile = FreeFile
      Open filename For Append As hfile
      Print #hfile, ErrorString & " v" & App.Revision & IIf(USE6080, " 6080 ", " 6040 ") & Now
      Close hfile

    Case TRACE_DEBUG
      Debug.Print ErrorString
    Case TRACE_TRACE
      Trace ErrorString, True
    Case Else
      ' nothing
  End Select
End Sub


Public Function Trace(ByVal s As String, Optional override As Boolean = False)
  On Error Resume Next
  If gTracing Or override Then
    Win32.OutputDebugString s & vbCrLf
  End If
End Function


Public Function SpecialLog(ByVal s As String)
  Dim hfile         As Integer
  Dim filename As String
  filename = App.Path & "/EVT.TXT"
  limitFileSize filename
  hfile = FreeFile
  Open filename For Append As hfile
  Print #hfile, s
  Close hfile
End Function
Public Function CheckAutoClears() As Long

  alarms.CheckAutoClear
  Alerts.CheckAutoClear
  Externs.CheckAutoClear



End Function

'Public Function CheckWatchdog()
'  Dim WDTO          As Long
'
'  WDTO = Configuration.WatchdogTimeout * 0.75  ' trigger at 75% of requested time
'
'  If DateAdd("s", -WDTO, LastWatchDog) > Now Then
'
'    Select Case Configuration.WatchdogType
'      Case WD_BERKSHIRE
'          BerkshireWD.Tickle
'      Case WD_UL
'        SetWatchdog Configuration.WatchdogTimeout
'      Case Else
'    End Select
'    LastWatchDog = Now
'
'  End If
'End Function
Public Function CheckComm() As Long
  Dim d             As cESDevice

10 On Error GoTo CheckComm_Error

20 Set d = Devices.Item(1)
30 If d.Dead = 0 Then
40  If DateDiff("n", d.LastSupervise, Now) >= d.SupervisePeriod Then          ' was:  Configuration.RxTimeout
50    PostEvent d, Nothing, Nothing, EVT_COMM_TIMEOUT, 0
60  End If
70 End If


CheckComm_Resume:
  CheckComm = Err.Number
80 On Error GoTo 0
90 Exit Function

CheckComm_Error:

100 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.CheckComm." & Erl
110 Resume CheckComm_Resume


End Function

Public Sub AuxLoop()
  Static SecondCounter As Long
  SecondCounter = SecondCounter + 1
  If SecondCounter > 30 Then  ' 30 seconds
    CheckIfReportsDue
    SecondCounter = 0
  End If

End Sub

Public Sub HeartBeat()
  Static Busy       As Boolean      ' prevent reentry

  ' hearbeat is at timer intervals... as short as 1 ms
  If gHoldOff Then
    dbg "Holdoff Heartbeat"
    Exit Sub  ' pending operation flag
  End If

  If StopIt Then
    dbg "Stopping Heartbeat"
    Exit Sub  ' global stop flag. used all events are stopped when exiting program.
  End If

  If Busy Then
    dbg "Busy Heartbeat"
    Exit Sub
  End If

  Busy = True

  ' call on every "tick"
  TickCounter = TickCounter + TimerPeriod

  If TickCounter >= 20 Then  ' at least every 20 ms
    If MASTER Then
      'dbg "Doread"
      DOREAD

    End If
    TickCounter = 0
  End If


  HalfSecondCounter = HalfSecondCounter + TimerPeriod
  SecondCounter = SecondCounter + TimerPeriod  ' our second counter

  If HalfSecondCounter >= 500 Then
    Call HalfSecond
    HalfSecondCounter = 0
  End If


  If SecondCounter >= 1000 Then  ' one second or thereabouts
    'Debug.Print "One Second " & Format(Now, "hh:nn:ss")


    Call OneSecond
    
    SecondCounter = 0
    
    ' three seconds delay in getting packet messages from mobile phones
    ThreeSecondCounter = ThreeSecondCounter + 1
    If ThreeSecondCounter >= 3 Then
      ThreeSeconds
      ThreeSecondCounter = 0
    End If
    
    MinuteCounter = MinuteCounter + 1
    TenMinuteCounter = TenMinuteCounter + 1
    ThirtyMinuteCounter = ThirtyMinuteCounter + 1

    If MASTER Then
      If USE6080 Then
        CheckIf6080Alive
      End If
      CheckLogins
    End If

    If Not MASTER Then
      RemoteRefreshCounter = RemoteRefreshCounter + 1
      'Debug.Print "RemoteRefresh in "; RemoteRefresh_Delay - RemoteRefreshCounter; " seconds"
      If RemoteRefreshCounter >= RemoteRefresh_Delay Then
        ResetRemoteRefreshCounter

        If RemoteAutoEnrollEnabled Then
          If RemotePollAutoEnroll() <> 0 Then
            frmTransmitter.DisableAutoEnroll
          End If
        End If
        
        
        Call RefreshRemoteAlarms
      End If


    End If
  End If



  If MinuteCounter >= 60 Then
    'Debug.Print "One Minute " & Format(Now, "hh:nn:ss")
    Call OneMinute
    MinuteCounter = 0
  End If



  If TenMinuteCounter >= 600 Then
    'Debug.Print "TEN Minute " & Format(Now, "hh:nn:ss")
    Call TenMinute
    TenMinuteCounter = 0
  End If

  If ThirtyMinuteCounter >= 1800 Then
    'Debug.Print "30 Minute " & Format(Now, "hh:nn:ss")
    Call ThirtyMinute
    ThirtyMinuteCounter = 0
  End If




  If MASTER Then

    If lastupdate = 0 Then
      lastupdate = Now
    End If
    If DateDiff("s", lastupdate, Now) >= 10 Then
      lastupdate = Now
      bytespermin = packetizer.TotalBytes
      packetizer.TotalBytes = 0
      packetspermin = packetizer.totalpackets
      packetizer.totalpackets = 0
      ' Debug.Print "Total Bytes 10 sec    "; bytespermin
      ' Debug.Print "Total Packets  10 sec "; packetspermin
      If packetspermin > 0 Then
        ' Debug.Print "Bytes / Packet        "; Format(CDbl(bytespermin) / CDbl(packetspermin), "0")
      End If
    End If
  End If



  Busy = False
End Sub
Sub SendRebuilds()
  ' If RebuildQue.HasNext Then
  'Outbounds.AddPreparedMessage RebuildQue.GetNext
  ' End If
End Sub

Sub HalfSecond()
  If MASTER Then
    'If InBounds.Count Then
      ProcessInbounds
    'End If
    GetRemoteRequests
  End If
End Sub
Sub OneSecond()


  MinuteCounter = MinuteCounter + 1
  frmMain.UpdateUptimeStats  ' toggles colors on buttons, does beepers
  If MASTER Then
    CheckWatchdog
    If USE6080 = 0 Then
      CheckComm
    End If
    CheckAssur
    CheckPagers         ' actually transmits pages
    CheckPageRequests   ' adds to device queues
    CheckAutoClears
    CheckExternalPages
    ProcessRemoteMonitoring
    CheckForPushEnabled
    CheckforMobilePackets ' moved here for better response times
    
    
    ' master only
    If RemoteAutoEnroller.CheckTimeout() Then
      'dbg "RemoteAutoEnroller.CheckTimeout Timeout = True" & vbCrLf
      RemoteAutoEnroller.RemoteEnrollEnabled = False
    End If
    
    
    PingMonitor False, "facilityid=" & Configuration.MonitorFacilityID & "&eventcode=" & EVENT_FACILITY_NONE

  Else
    ' see if connection is lost
    CheckRemoteConnectStatus
  End If
End Sub
Sub CheckExternalPages()
  Dim rs            As Recordset
  Dim GroupID       As Long
  Dim PagerID       As Long
  Dim message       As String
  Dim ID            As Long
  Dim SQL           As String

10 On Error GoTo CheckExternalPages_Error

20 SQL = "SELECT TOP 1 * FROM externalpages ORDER BY ID"
30 Set rs = ConnExecute(SQL)
40 If Not rs.EOF Then
50  GroupID = rs("groupid")
60  PagerID = rs("pagerid")
70  ID = rs("ID")
80  message = Trim$(rs("message") & "")
90 End If
100 rs.Close
110 Set rs = Nothing
120 If ID Then
130 SQL = "DELETE FROM externalpages WHERE ID = " & ID
140 ConnExecute SQL
150 End If
160 If ID Then
170 If Len(message) Then
180   If PagerID Then
190     SendToPager message, PagerID, 0, "", "", PAGER_NORMAL, left$(message, 19), 0, 0
200   ElseIf GroupID Then
210     SendToGroup message, GroupID, "", "", PAGER_NORMAL, left$(message, 19), 0, 0
220   End If
230 End If
240 End If

CheckExternalPages_Resume:

250 On Error GoTo 0
260 Exit Sub

CheckExternalPages_Error:

270 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.CheckExternalPages." & Erl
280 Resume CheckExternalPages_Resume
End Sub

Sub RefreshRemoteAlarms()
  If Not MASTER Then
    If (gMyAlarms) Then
      ClientGetSubscribedAlarms
    Else
      ClientGetAlarms
    End If
  End If
  ResetRemoteRefreshCounter

End Sub


Sub ThreeSeconds()

  ThreeSecondCounter = 0
End Sub
Sub CheckforMobilePackets()
10      If MASTER Then
          Dim ids()            As String
          Dim Packets()        As String
          Dim SQL              As String
          Dim rs               As ADODB.Recordset
          Dim Count            As Long
          Dim j                As Long

20        ReDim RecordIDs(1) ' 0
30        ReDim Packets(1) ' 0

40        ReDim ids(1)
          
50        Count = 0
        

60        SQL = "SELECT * FROM Packets "
70        Set rs = ConnExecute(SQL)
80        Do Until rs.EOF
90          Count = Count + 1
100         ReDim Preserve ids(Count)
110         ReDim Preserve Packets(Count)
120         ids(Count) = rs("ID")
130         Packets(Count) = rs("text") & ""
140         rs.MoveNext
150       Loop
160       rs.Close
170       Set rs = Nothing
          Dim IDs2Delete       As String

180       If Count > 0 Then
            ' delete packets returned
190         ids(0) = "0"
200         IDs2Delete = Join(ids, ",")
              
210         SQL = "DELETE FROM Packets WHERE ID in (" & IDs2Delete & ")"
220         ConnExecute SQL


230         For j = 1 To Count
              ' process mobile packets
240           MobileParseAndProcess Packets(j)
250         Next

260       End If

270     End If
End Sub

Function MobileParseAndProcess(ByVal PacketData As String)


  Dim isXML              As Boolean
  Dim xmlDoc             As DOMDocument
  Dim RootNode           As IXMLDOMNode
  Dim NodeList           As MSXML2.IXMLDOMNodeList
  Dim childnode          As IXMLDOMNode
  Dim i                  As Long

  Dim Action             As String
  Dim AlarmID            As String
  Dim GroupID            As String
  Dim PagerID            As String
  Dim EventType          As String
  Dim Disposition        As String
  Dim devicetime         As String
  Dim Username           As String

  Dim j                  As Long
  Dim rc                 As Long

  Dim Device             As cESDevice

  Dim alarm              As cAlarm

  Dim SQL                As String
  Dim rs                 As ADODB.Recordset


  i = InStr(1, PacketData, "<?xml", vbTextCompare)
  isXML = i > 0                ' alternative may be JSON, not for now

  


  If isXML Then
    Set xmlDoc = New DOMDocument
    rc = xmlDoc.LoadXML(PacketData)
    Set RootNode = xmlDoc.childnodes(1)
    Set NodeList = RootNode.childnodes
    For Each childnode In NodeList
      Select Case LCase$(childnode.baseName)

        Case "action"
          Action = LCase(childnode.text)
        Case "alarmid"
          AlarmID = childnode.text
        Case "groupid"
          GroupID = Val(childnode.text)
        Case "pagerid"
          PagerID = Val(childnode.text)
        Case "eventtype"
          EventType = Val(childnode.text)
        Case "disposition"
          Disposition = childnode.text
        Case "username"
          Username = childnode.text
        Case "devicetime"
          devicetime = FixISODateTime(childnode.text)
        Case "disposition"
          Disposition = childnode.text
        Case Else
          ' undefined
      End Select
    Next
  End If

  Action = LCase$(Action)

  
  Debug.Print "AlarmID " & AlarmID
  
  Select Case Action
    Case "accept"
      ' respond just creates a response entry in database, no visual changes to host
      Set alarm = New cAlarm

      SQL = "SELECT * FROM Alarms WHERE ID = " & AlarmID  ' alarm id should be parent alarm's ID
      Set rs = ConnExecute(SQL)
      Do Until rs.EOF
        alarm.ID = rs("id")
        alarm.AlarmID = Val(rs("id") & "")
        alarm.PriorID = Val(rs("id") & "")
        alarm.STAT = rs("status")
        alarm.Serial = rs("serial") & ""
        alarm.ResidentID = Val(rs("residentid") & "")
        alarm.RoomID = Val(rs("roomid") & "")
        alarm.inputnum = Val(rs("inputnum") & "")
        alarm.Username = Username & ""
        alarm.Announce = rs("announce") & ""
        alarm.Phone = rs("phone") & ""
        alarm.locationtext = rs("userdata") & ""
        Select Case Val(rs("eventtype") & "")
          Case EVT_ASSISTANCE
            alarm.EventType = EVT_ASSISTANCE_RESPOND
          Case EVT_EMERGENCY
            alarm.EventType = EVT_EMERGENCY_RESPOND
          Case EVT_ALERT
            alarm.EventType = EVT_ALERT_RESPOND
          Case Else
            alarm.EventType = EVT_GENERIC_RESPOND
        End Select
        alarm.LOCIDL = Val(PagerID)


        Exit Do
      Loop
      rs.Close
      Set rs = Nothing

      If alarm.AlarmID <> 0 Then
        'Alarm.PriorID = Alarm.alarmid
        LogAlarm alarm, alarm.EventType, Username
      End If




    Case "assist"
      Dim NoAssistancePending   As Boolean
      Dim inputnum       As Long
      Dim TempAlarm As cAlarm
      ' adds assist alarm if it doesn't already exist, update main screen if new (assist)alarm created


      SQL = "SELECT * FROM Alarms where eventtype = " & EVT_ASSISTANCE & " AND alarmID = " & AlarmID
      Set rs = ConnExecute(SQL)
      NoAssistancePending = rs.EOF
      rs.Close
      Set rs = Nothing

      If NoAssistancePending Then

        Set TempAlarm = New cAlarm
        
        Debug.Print "AlarmID " & AlarmID
        Set TempAlarm = TempAlarm.ReadFromDB(Val(AlarmID))
        If Not (TempAlarm Is Nothing) Then


          'For j = 1 To alarms.alarms.Count
          'Set alarm = alarms.alarms(j)
          'If alarm.ID = Val(AlarmID) Then
          'If alarm.Alarmtype <> EVT_ASSISTANCE Then  ' can't staff assist a staff assist

          'Set Device = Devices.Device(alarm.Serial)
          Set Device = Devices.Device(TempAlarm.Serial)
          inputnum = TempAlarm.inputnum
          Call PostEvent(Device, Nothing, TempAlarm, EVT_ASSISTANCE, inputnum, Username)
          ''            Exit For
        End If
      End If
      'Next


      '      For j = 1 To Alerts.alarms.count
      '        Set alarm = Alerts.alarms(j)
      '        If alarm.ID = Val(alarmid) Then
      '          If alarm.Alarmtype <> EVT_ASSISTANCE Then ' can't staff assist a staff assist
      '            Set Device = Devices.Device(alarm.Serial)
      '            Call PostEvent(Device, Nothing, alarm, EVT_ASSISTANCE, alarm.inputnum)
      '            Exit For
      '          End If
      '        End If
      '      Next



    Case "finalize", "closeout"
      ' removes from history AND if EVT_Assistance then removes from alarms too
      Dim AlarmCount     As Long
      AlarmCount = alarms.alarms.Count
      For j = 1 To AlarmCount
        Set alarm = alarms.alarms(j)
        If alarm.ID = Val(AlarmID) Then
          If alarm.Alarmtype = EVT_ASSISTANCE Then  ' can't staff assist a staff assist
            Set Device = Devices.Device(alarm.Serial)
            alarm.Disposition = Disposition
            alarm.LOCIDL = Val(PagerID)
            Call PostEvent(Device, Nothing, alarm, EVT_ASSISTANCE_ACK, alarm.inputnum, Username)
            Exit For
          End If
        End If
      Next
      If j > AlarmCount Then   ' not found in active alarms
        Set alarm = New cAlarm

        SQL = "SELECT * FROM Alarms WHERE ID = " & AlarmID  ' alarm id should be parent alarm's ID
        Set rs = ConnExecute(SQL)
        Do Until rs.EOF
          alarm.ID = rs("id")
          alarm.AlarmID = rs("id")
          alarm.PriorID = rs("id")
          alarm.STAT = rs("status")
          alarm.Serial = rs("serial")
          alarm.ResidentID = rs("residentid")
          alarm.RoomID = rs("roomid")
          alarm.Username = Username
          alarm.Announce = rs("announce")
          alarm.Phone = rs("phone")
          alarm.locationtext = rs("userdata")
          'Alarm.inputnum = rs("inputnum")
          alarm.Disposition = Disposition
          Select Case rs("eventtype")

            Case EVT_EMERGENCY
              alarm.EventType = EVT_EMERGENCY_FINALIZE
            Case EVT_ALERT
              alarm.EventType = EVT_ALERT_FINALIZE
            Case Else
              alarm.EventType = EVT_GENERIC_FINALIZE
              'case ???
              alarm.EventType = EVT_EXTERN_FINALIZE
          End Select



          Exit Do
        Loop
        rs.Close
        Set rs = Nothing

        alarm.LOCIDL = Val(PagerID)
        Set Device = Devices.Device(alarm.Serial)
        Call PostEvent(Device, Nothing, alarm, alarm.EventType, alarm.inputnum, Username)
      End If

      SQL = "DELETE FROM Mobile WHERE AlarmID = " & AlarmID
      ConnExecute SQL

    Case Else

  End Select
End Function


Sub DeleteFromMobile(ByVal AlarmID As Long)
  Dim SQL As String
      SQL = "DELETE FROM Mobile WHERE AlarmID = " & AlarmID
      ConnExecute SQL

End Sub

Sub CheckLogins()
  Dim Session       As cUser
  Dim j             As Integer

  'dbg "Session Count " & HostSessions.count

  For j = HostSessions.Count To 1 Step -1
    Set Session = HostSessions(j)
    'dbgHostRemote "Session Last Seen " & Session.LastSeen & " " & j
    If DateDiff("s", Session.LastSeen, Now) > 30 Then  ' Now 10 seconds remove dead logins  after two minutes (as one minute before 2008-09-29)

      If Session.Session <> gUser.Session Then  ' Local login never times out
        dbgHostRemote "Host kill session (overdue)" & Session.Session
        Debug.Print "Host kill session (overdue)" & Session.Session
        LogRemoteSession Session.Session, 0, "Host kill session (overdue) " & Session.LastSeen
        HostSessions.Remove j
      End If

    End If
  Next

End Sub

Sub CheckMobileDeadWood()
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim list()             As String
        Dim Count              As Long
        Dim j                  As Long
        Dim csvList            As String
        Dim XML                As String
        Dim Action             As String
        Dim AlarmID            As String
        Dim CutOffTime         As String


10      CutOffTime = DateAdd("n", -Configuration.MobileClearHistory, Now)


20      ReDim list(0)
30      list(0) = "0"

        Dim AlarmTime          As String

        '' might do Alarms.eventtype != EVT_ASSISTANCE

40      SQL = "SELECT Alarms.ID, Mobile.AlarmID, Alarms.EventType, Mobile.eAlarmTime , mobile.ended FROM  Alarms INNER JOIN Mobile ON Alarms.ID = Mobile.AlarmID  WHERE (mobile.ended = 1) AND   (Alarms.EventType != " & EVT_ASSISTANCE & ")"

50      Set rs = ConnExecute(SQL)
60      Do Until rs.EOF
70        AlarmTime = rs("eAlarmTime") & ""
80        If IsDate(AlarmTime) Then
90          If DateDiff("n", CDate(AlarmTime), CutOffTime) > 0 Then
100           Count = Count + 1
110           ReDim Preserve list(Count)
120           list(Count) = rs("ID")
130         End If
140       End If
150       rs.MoveNext
160     Loop
170     rs.Close
180     Set rs = Nothing

190     For j = 1 To UBound(list)

200       XML = "<?xml version=""1.0""?>"
210       XML = XML & "<update>"
220       XML = XML & taggit("action", "finalize")
230       XML = XML & taggit("alarmid", list(Count))
240       XML = XML & taggit("disposition", "System Timeout")
250       XML = XML & taggit("pagerid", 0)
260       XML = XML & taggit("apikey", 0)
270       XML = XML & taggit("username", "System")
280       XML = XML & taggit("devicetime", Now)
290       XML = XML & "</update>"
300       SQL = "INSERT INTO Packets(Posted,PostDate,Session,Text) values (0,'" & Now & "',0," & q(XML) & ")"
310       ConnExecute SQL
320     Next

330     If (UBound(list) > 0) Then
340       csvList = Join(list, ",")
350       SQL = "DELETE FROM mobile WHERE AlarmID in (" & csvList & ")"
360       ConnExecute SQL
370     End If



380     CutOffTime = DateAdd("n", -Configuration.MobileClearAssist, Now)

390     ReDim list(0)
400     list(0) = "0"
410     SQL = "SELECT Alarms.ID, Mobile.AlarmID, Alarms.EventType, Mobile.eAlarmTime FROM  Alarms INNER JOIN Mobile ON Alarms.ID = Mobile.AlarmID  WHERE (Alarms.EventType = " & EVT_ASSISTANCE & ")"
420     Set rs = ConnExecute(SQL)
430     Do Until rs.EOF
440       AlarmTime = rs("eAlarmTime") & ""
450       If IsDate(AlarmTime) Then
460         If DateDiff("n", CDate(AlarmTime), CutOffTime) > 0 Then
470           Count = Count + 1
480           ReDim Preserve list(Count)
490           list(Count) = rs("ID")
500         End If
510       End If
520       rs.MoveNext
530     Loop
540     rs.Close
550     Set rs = Nothing

560     For j = 1 To UBound(list)

570       XML = "<?xml version=""1.0""?>"
580       XML = XML & "<update>"
590       XML = XML & taggit("action", "finalize")
600       XML = XML & taggit("alarmid", list(Count))
610       XML = XML & taggit("disposition", "System Timeout")
620       XML = XML & taggit("pagerid", 0)
630       XML = XML & taggit("apikey", 0)
640       XML = XML & taggit("username", "System")
650       XML = XML & taggit("devicetime", Now)
660       XML = XML & "</update>"
670       SQL = "INSERT INTO Packets(Posted,PostDate,Session,Text) values (0,'" & Now & "',0," & q(XML) & ")"
680       ConnExecute SQL
690     Next

700     If (UBound(list) > 0) Then
710       csvList = Join(list, ",")
720       SQL = "DELETE FROM mobile WHERE AlarmID in (" & csvList & ")"
730       ConnExecute SQL
740     End If





End Sub

Sub CheckMasterLogin()
  ' see if i've been bumped by a remote!

  Dim j             As Integer
  Dim Session       As cUser

  For j = HostSessions.Count To 1 Step -1
    Set Session = HostSessions(j)
    If Session.Session = gUser.Session Then  ' Local login doesn't timeout
      Exit For
    End If

  Next
  If (j = 0) And (LoggedIn = True) Then
    frmMain.DoLogin
    dbg "Logout bumped by remote CheckMasterLogin"
  End If
End Sub

Sub CheckOutputServers()
  Dim p             As cPageDevice
  For Each p In gPageDevices
    If p.ErrorStatus <> 0 Then  ' ok
      ' make entry via devices
    End If
  Next

End Sub

Sub OneMinute()

  'Debug.Print "One minute start"

  If MASTER Then
    

    DoSupervise
    CheckIfBackupDue
        
    CheckMasterLogin
    CheckOutputServers
    CheckMobileDeadWood
    
  

  Else
    CheckIfHasAssur

  End If


  'Debug.Print "One minute end"

End Sub

Sub CheckIfHasAssur()

End Sub

Sub TenMinute()
  If MASTER Then
    PingSession  ' stores time in database... to see when it all went to shit.
  End If
End Sub
Sub ThirtyMinute()
  ' update PCA time
  Dim j             As Integer
  Dim d             As cESDevice
  If MASTER Then
    For j = 1 To Devices.Count
      Set d = Devices.Devices(j)
      If d.IsPCA Then
        ' upate time
        Outbounds.AddMessage d.Serial, MSGTYPE_SET_TIME, "", 0
      End If
    Next
  End If
End Sub

Public Sub ResetRemoteRefreshCounter(Optional ByVal Seconds As Long = 0)
  RemoteRefreshCounter = Seconds
End Sub


Sub ProcessInbounds()
        Dim a                  As cAlarm
        Dim j                  As Integer
        Dim Incount            As Integer
        Dim PagerID            As Long


10      On Error GoTo ProcessInbounds_Error

20      j = 1  ' j no longer used as iterator


30      Incount = InBounds.Count
        '40      Do While j <= Incount
        ' changed how these are removed, from j to the first one
        'Debug.Print "ProcessInbounds count " & Incount
        
40      Do While InBounds.Count > 0  'j <= Incount
          ' maybe output debug string to show number of inbounds and j
          '50        Set a = InBounds.alarms(j)   ' inbounds alarms are pending alarms, not yet posted to the alarm list

50        Set a = InBounds.alarms(1)   ' inbounds alarms are pending alarms, not yet posted to the alarm list
          
          ' always pull off first one

60        If (a.ProcessLocations) Then    ' all data is ready and timed out
            Debug.Print "Ready to process " & Format(Now, "nn:ss")
70          Select Case a.Alarmtype
              Case EVT_EMERGENCY
                'dbgloc "ProcessInbounds EVT_EMERGENCY " & Incount
80              If a.IsPortable = 1 Then  ' we only fo location on portables
90                If USE6080 Then
                    ' Skip it, we have location in Location Text
100               Else
                    ' >>>>>>>>>>>>>> do the waypoints <<<<<<<<<<<<<<<
110                 Waypoints.Locate a

120               End If
130             Else  ' else leave the location blank
140               a.locationtext = ""
150             End If

160             If (a.Serial = Configuration.SurveyDevice) Then  ' (Configuration.PCARedirect Or 1) And (a.Serial = Configuration.SurveyDevice) Then  ' shunts device alarm from alarm list

170               If Configuration.surveymode = TWO_BUTTON_MODE Then
180                 If a.alarm = 1 Then
                      'dbgloc "SendToPager: OK Got it, Press Button Two."

                      'PagerID = GetPagerIDFromSerial(Configuration.SurveyDevice)
                      ' was : SendToPager "OK Got it, Press Button Two.", SurveyDevice.PagerID, 0, "", "", PAGER_NORMAL, "", 0
190                   SendToPager "OK Got it, Press Button Two.", Configuration.SurveyPager, 0, "", "", PAGER_NORMAL, "", 0, a.inputnum
200                 End If
210               ElseIf Configuration.surveymode = EN1221_MODE Then
220                 If a.alarm = 1 Then
                      'dbgloc "SendToPager: OK Got it, Press Button Two."

                      'PagerID = GetPagerIDFromSerial(Configuration.SurveyDevice)
                      ' was : SendToPager "OK Got it, Press Button Two.", SurveyDevice.PagerID, 0, "", "", PAGER_NORMAL, "", 0
230                   SendToPager "OK Got it, Clear the Alarm.", Configuration.SurveyPager, 0, "", "", PAGER_NORMAL, "", 0, a.inputnum
240                 End If
250               Else
                    'dbgloc "PCARedirect "
260                 If (Configuration.PCARedirect) Then
270                   PCARedirect a, Configuration.SurveyPCA   ' sends to innovonics pager
280                 End If
290               End If

300             ElseIf (a.Serial = Configuration.WaypointDevice) Then  ' it's a waypointb verifier
310               SendToPager "@ " & a.locationtext, Configuration.SurveyPager, 0, "", "", PAGER_NORMAL, "", 0, a.inputnum

320             Else  ' it's a real alarm

330               If (alarms.AddAlarm(a)) Then  ' it's a new alarm
340                 AddPageRequest a, a.Alarmtype
350               End If
                  '320               InBounds.Delete j
                  'InBounds.Delete 1
360             End If
370             InBounds.Delete 1



380           Case EVT_ALERT
                'dbgloc "ProcessInbounds EVT_ALERT " & Incount
390             If a.IsPortable = 1 Then
                  'If SYSTEM_6080 Then
400               If USE6080 Then
                    ' we have location in Location Text
410               Else
420                 Waypoints.Locate a
430               End If
440             Else
450               a.locationtext = ""
460             End If

470             If (Configuration.PCARedirect = 1) And (a.Serial = Configuration.SurveyDevice) Then    ' shunts device alarm from alarm list
                  ' no op
480             Else

490               If (Alerts.AddAlarm(a)) Then    ' it's a new alert
500                 AddPageRequest a, a.Alarmtype
510               End If
520             End If

                '480               InBounds.Delete j
530             InBounds.Delete 1
540           Case EVT_BATTERY_FAIL
                'Do I need to add this here? to avoid overflow ?
550             InBounds.Delete 1
560           Case EVT_CHECKIN_FAIL
                'Do I need to add this here? to avoid overflow ?
570             InBounds.Delete 1
580           Case Else
590             InBounds.Delete 1

600         End Select
610         a.Posted = True
620       Else
630         j = j + 1  ' why would this overflow ?


650         If j > 10 Then
              'Sleep 100  ' give things time to arrive
              j = 0
660           'LogProgramError "Error 0xDEADC0DE " & " (j overflow) at modMain.ProcessInbounds.550"
670           Exit Do
680         End If
690       End If
700       Incount = InBounds.Count
          
710     Loop

ProcessInbounds_Resume:
720     On Error GoTo 0
730     Exit Sub

ProcessInbounds_Error:

740     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.ProcessInbounds." & Erl
750     Resume ProcessInbounds_Resume

End Sub
'Function GetPagerIDSerial(ByVal SurveyDeviceSerial As String) As Long
'
'End Function


Sub DOREAD()
  'serial polling routine

  Dim t             As Double
  Dim haspacket     As Boolean
  Dim Passes        As Long
  Dim Pakcet        As cESPacket

  t = Timer

  'dbg "DoRead ****"
  'dbg "DoRead Start " & Format(t, "0.000")

  haspacket = False

  Do  ' check at end of loop for no packet

    If (0 = USE6080) Then

      packetizer.FetchData  ' pulls serial data into local buffer
      packetizer.Process  ' scans local buffer and parses packets, also rejects bad packets

      If packetizer.PacketReady Then
        ProcessESPacket packetizer.GetPacket
        haspacket = True
      Else
        haspacket = False
      End If


      'dbg "DoRead 80  " & Format(Timer - t, "0.000")

    Else ' a 6080 system

      i6080.GetData
      If i6080.HasMessages Then
        If (First6080packet = False) Then
          First6080packet = True
        End If
        'Set packet = i6080.ConvertToPacket(i6080.GetNextMessage)
        'Debug.Print "i6080 Has Messages"
        ProcessESPacket i6080.ConvertToPacket(i6080.GetNextMessage)
        haspacket = True
      Else
        haspacket = False
      End If

    End If

    Outbounds.SendMessages

        
    

    GetDukaneRequests  ' rename this

    ProcessSerialIns  ' much like packetizer, but one serialin for each watched port/device
    
    
    
    

    If alarms.Pending() Then
      frmMain.ProcessAlarms
    End If

    If Alerts.Pending() Then
      frmMain.ProcessAlerts
    End If

    If LowBatts.Pending Then
      frmMain.ProcessBatts
    End If

    If Troubles.Pending Then
      frmMain.ProcessTroubles
    End If

    If Externs.Pending Then
      frmMain.ProcessExterns
    End If

'    If InBounds.Count Then
'      ProcessInbounds
'    End If

    Passes = Passes + 1

    If Passes > 60 Then  ' failsafe at eating 60 packets at a time
      Exit Do
    End If
  Loop While haspacket

  If Assurs.Pending Then
    frmMain.ProcessAssurs False
  End If


exitdoread:

DOREAD_Resume:
  On Error GoTo 0
  Exit Sub

DOREAD_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.DOREAD." & Erl
  Resume DOREAD_Resume

End Sub

Function CheckForPushEnabled()
  
          ' Debug.Print "CheckForPushEnabled PushProcessor.Send QueCount = " & PushProcessor.Que.Count & "  Busy = "; PushProcessor.Busy
  'PushProcessor.Busy = False
490        If Not PushProcessor.Busy Then
500           PushProcessor.Send
510        End If
  
  
End Function

Function ProcessRemoteMonitoring()
  Dim Device             As cESDevice
  Dim packet             As cESPacket
  Dim Elapsed            As Double
  
  Static Busy            As Boolean

  If MASTER Then

    If Busy Then Exit Function
    Busy = True
    
    'Debug.Print "***********"
    'Debug.Print Now
    
    For Each Device In Devices.Devices
      'Debug.Print "ProcessRemoteMonitoring", device.Serial, device.MIDPTI
      'If device.Model = "REMOTE" Then
      If Device.MIDPTI = &H1EE Then
      'If device.MIDPTI = (&H1 * 256&) + &HEE Then
        If Device.LastSeen = 0 Then
          Device.LastSeen = Now
        End If
        
        If Device.Model = "REMOTE" Then
          If Device.alarm = 0 Then
            Elapsed = DateDiff("s", Device.LastSeen, Now)
            If Elapsed > 120 Then  ' seconds
              Set packet = New cESPacket
              packet.DateTime = Now
              packet.alarm = 1
              packet.PacketType = 2
              packet.Serial = Device.Serial
              packet.SetMIDClassPTI &H1, &HEE, &HEE
              ProcessESPacket packet
            End If
          ElseIf Device.alarm = 1 Then
              Elapsed = DateDiff("s", Device.LastSeen, Now)
              If Elapsed < 20 Then ' seconds
                Set packet = New cESPacket
                packet.DateTime = Now
                packet.alarm = 0
                packet.PacketType = 2
                packet.Serial = Device.Serial
                packet.SetMIDClassPTI &H1, &HEE, &HEE
                ProcessESPacket packet
              End If
            End If
          
        End If
      End If
    Next


    Busy = False

  End If
End Function

Sub ProcessSerialIns()
  Dim j             As Integer
  Dim si            As cSerialInput
  Static Busy       As Boolean

10 On Error GoTo ProcessSerialIns_Error

20 If Busy Then Exit Sub
30 Busy = True


  'If SerialIns.count Then Stop

40 For j = 1 To SerialIns.Count
50  Set si = SerialIns(j)
60  si.FetchData
70  If si.PacketReady Then
80    ProcessESPacket si.GetPacket
90  End If
100 Next

110 Busy = False

ProcessSerialIns_Resume:
120 On Error GoTo 0
130 Exit Sub

ProcessSerialIns_Error:

140 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.ProcessSerialIns." & Erl
150 Resume ProcessSerialIns_Resume

End Sub

Sub DoSupervise()

  Dim d             As cESDevice
  Dim rx            As cESDevice
  Dim start         As Long

  Dim j             As Integer
  Dim a             As cAlarm
  ' Check supervisory at "gSupervisePeriod" intervals
10 On Error GoTo DoSupervise_Error

20 Set rx = Devices.Item(1)       ' serial RX port ' always first device
30 If rx.Dead = 0 Then        ' RX working
    ' time to check?
    'OutputDebugString "Next Supervise in " & Format(gSupervisePeriod - DateDiff("s", gLastSupervise, Now) / 60, "0.00") & " minutes" & vbCrLf
40  'Debug.Print "since last supervise "; DateDiff("n", gLastSupervise, Now)
50  'If DateDiff("n", gLastSupervise, Now) >= 0.5 Then
    If 1 Then
60    'Debug.Print "Supervise check at: " & Now
70    start = Win32.timeGetTime()

80    If 1 Then
        ' skip this
90
100     For j = 1 To Devices.Devices.Count
110       Set d = Devices.Devices(j)
          If j = 22 Then
            If InIDE Then
'              Stop
            End If
          End If
120       If d.CLSPTI = 0 And j > 1 Then        ' "EN6040"
            ' skip RX checkin enrolled as "device"
130       ElseIf d.IsPCA Then        'Or d.Serial = "00000000" Then
            ' skip check-in for PCAs
140       ElseIf d.Model = COM_DEV_NAME Then
            ' skip check-in Serial-Inputs
150       Else
            'Trace "Check Supervise " & D.Serial
160         If d.Dead = 0 Then ' already dead ?? No?? then...
170           If d.IsLate Then        ' on 6080 devices translates from d.ismissing ' DateDiff("n", d.LastSupervise, Now) > d.SupervisePeriod Then
                'Debug.Print "************* MISSING *********** in DO Supervise"
180             PostEvent d, Nothing, a, EVT_CHECKIN_FAIL, 0
190           End If
200         End If
210       End If
220     Next

230     'Debug.Print "Time to check all Troubles: "; Win32.timeGetTime() - start
240   End If
250   gLastSupervise = Now
260 End If
270 End If

280 start = Win32.timeGetTime()
  ' clear any that have communicated
  'dbg "Troubles.alarms.count " & Troubles.alarms.count
290 For j = Troubles.alarms.Count To 1 Step -1
    'Debug.Print "x";
    Dim zx          As Long
300 zx = DoEvents()  ' might not need this, or at least defer to every n iterations
310 Set a = Troubles.alarms(j)
320 If Not a Is Nothing Then
330   If a.Alarmtype = EVT_CHECKIN_FAIL Then
340     Set d = Devices.Device(a.Serial)
        'Debug.Print "+";
350     If Not d Is Nothing Then
          'Debug.Print "+";

360       If USE6080 Then
370         If d.Dead = 0 Then
            Debug.Print "************* BACK FROM THE DEAD *********** in DO Supervise"
380           PostEvent d, Nothing, a, EVT_CHECKIN, 0
390         End If
400       Else
            
            
            
410         If Not d.IsLate Then        ' DateDiff("n", d.LastSupervise, Now) <= d.SupervisePeriod Then
420           PostEvent d, Nothing, a, EVT_CHECKIN, 0
430         End If
440       End If
450     End If
460   End If
470 End If
480 Next
  'Debug.Print
  'dbg "Time to restore any Troubles: " & Win32.timeGetTime() - Start
490 start = Win32.timeGetTime()
500 If Troubles.Pending() Then
510 frmMain.ProcessTroubles
520 End If
  'dbg "Time to process any Troubles: " & Win32.timeGetTime() - Start
DoSupervise_Resume:
530 On Error GoTo 0
540 Exit Sub

DoSupervise_Error:

550 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.DoSupervise." & Erl
560 Resume DoSupervise_Resume


End Sub


Sub CheckPagers()
  Dim pageDevice             As cPageDevice
10 On Error GoTo CheckPagers_Error
  'Debug.Print "Check Pagers ", Now, "pageDevices ", gPageDevices.count
20 For Each pageDevice In gPageDevices
30  pageDevice.Poll
40 Next

CheckPagers_Resume:
50 On Error GoTo 0
60 Exit Sub

CheckPagers_Error:

70 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.CheckPagers." & Erl
80 Resume CheckPagers_Resume

End Sub

Function GetDBVer() As Long
  Dim rs            As Recordset
  Dim Revision      As Long

  Set rs = ConnExecute("SELECT Revision FROM SystemState")
  If Not rs.EOF Then
    If IsNull(rs(0).Value) Then
      Revision = 0
    Else
      Revision = rs(0).Value
    End If

  End If
  rs.Close
  Set rs = Nothing
  GetDBVer = Revision

End Function

Function WriteDBVersion() As Boolean
  Dim rs            As Recordset
  Dim SQL           As String

  Set rs = ConnExecute("SELECT count(*) FROM SystemState")
  If rs(0).Value = 0 Then
    SQL = "INSERT INTO SystemState (Major, Minor, Revision) values (" & App.Major & "," & App.Minor & "," & App.Revision & ")"
  Else
    SQL = "UPDATE SystemState SET Major = " & App.Major & ", Minor =" & App.Minor & ", Revision =" & App.Revision
  End If
  rs.Close
  Set rs = Nothing
  ConnExecute SQL

End Function

Function FixDatabase() As String
        Dim SQL                As String
        Dim Revision           As Long
        Dim DBVER              As Long


10      On Error GoTo FixDatabase_Error

20      Revision = App.Revision

        ' just for testing
        '  EnsureFieldExists "AATable", "Afield", adInteger, "-1", "0"
        '  EnsureFieldExists "AATable", "Description", adVarWChar, 50, ""
        '  EnsureFieldExists "AATable", "StartTime", adDate, "", ""
        '  EnsureFieldExists "AATable", "ImageData", adLongVarBinary, "", ""
        '  EnsureFieldExists "AATable", "AMemo", adLongVarWChar, "", ""



30      EnsureFieldExists "SystemState", "Major", adInteger, "0", "0"
40      EnsureFieldExists "SystemState", "Minor", adInteger, "0", "0"
50      EnsureFieldExists "SystemState", "Revision", adInteger, "0", "0"



60      EnsureFieldExists "DeviceTypes", "AutoClear", adInteger, "0", "0"
70      EnsureFieldExists "GlobalSettings", "MilitaryTime", adInteger, "0", "0"
80      EnsureFieldExists "GlobalSettings", "ElapsedEqACK", adInteger, "0", "0"
90      EnsureFieldExists "GlobalSettings", "DisableScreenOutput", adInteger, "0", "0"  ' for assurance, added 2009-11-2

100     DBVER = GetDBVer()
110     If Revision < DBVER Then
120       FixDatabase = Revision & " " & DBVER
130     End If


140     Select Case DBVER
          Case Is > 200: GoTo Version200Plus
150       Case Is > 190: GoTo Version190Plus
160       Case Else
170     End Select


180     EnsureFieldExists "Waypoints", "ID", adInteger, "-1", "0"
190     EnsureFieldExists "Waypoints", "Description", adVarWChar, 50, ""
200     EnsureFieldExists "Waypoints", "Building", adVarWChar, 50, ""
210     EnsureFieldExists "Waypoints", "Floor", adVarWChar, 50, ""
220     EnsureFieldExists "Waypoints", "Wing", adVarWChar, 50, ""

230     EnsureFieldExists "Waypoints", "Repeater1", adVarWChar, 10, ""
240     EnsureFieldExists "Waypoints", "Repeater2", adVarWChar, 10, ""
250     EnsureFieldExists "Waypoints", "Repeater3", adVarWChar, 10, ""

260     EnsureFieldExists "Waypoints", "Signal1", adInteger, "0", "0"
270     EnsureFieldExists "Waypoints", "Signal2", adInteger, "0", "0"
280     EnsureFieldExists "Waypoints", "Signal3", adInteger, "0", "0"



290     EnsureFieldExists "Sessions", "SessionID", adInteger, "-1", "0"
300     EnsureFieldExists "Sessions", "StartTime", adDate, "", ""
310     EnsureFieldExists "Sessions", "QuitTime", adDate, "", ""
320     EnsureFieldExists "Sessions", "LastPing", adDate, "", ""

330     EnsureFieldExists "Alarms", "Phone", adVarWChar, 50, ""
        'copytophone

340     EnsureFieldExists "Alarms", "Info", adVarWChar, 255, ""
350     EnsureFieldExists "Alarms", "Announce", adVarWChar, 50, ""
360     EnsureFieldExists "Alarms", "UserName", adVarWChar, 50, ""
370     EnsureFieldExists "Alarms", "SessionID", adInteger, "0", "0"


380     EnsureFieldExists "Devices", "IsPortable", adInteger, "0", "0"
390     EnsureFieldExists "Devices", "NumInputs", adInteger, "0", "0"
400     EnsureFieldExists "Devices", "Deleted", adInteger, "0", "0"
410     EnsureFieldExists "Devices", "ClearByReset", adInteger, "0", "0"

420     EnsureFieldExists "Devices", "SendCancel", adInteger, "0", "0"
430     EnsureFieldExists "Devices", "NG1", adInteger, "0", "0"
440     EnsureFieldExists "Devices", "NG2", adInteger, "0", "0"
450     EnsureFieldExists "Devices", "NG3", adInteger, "0", "0"
460     EnsureFieldExists "Devices", "AlarmMask", adInteger, "0", "0"
470     EnsureFieldExists "Devices", "UseAssur2", adInteger, "0", "0"
480     EnsureFieldExists "Devices", "Announce", adVarWChar, 50, ""
490     EnsureFieldExists "Devices", "Repeats", adInteger, "0", "0"
500     EnsureFieldExists "Devices", "Pause", adInteger, "0", "0"
510     EnsureFieldExists "Devices", "RepeatUntil", adInteger, "0", "0"
520     EnsureFieldExists "Devices", "OG1", adInteger, "0", "0"
530     EnsureFieldExists "Devices", "OG2", adInteger, "0", "0"
540     EnsureFieldExists "Devices", "OG3", adInteger, "0", "0"
550     EnsureFieldExists "Devices", "VacationSuper", adInteger, "0", "0"
560     EnsureFieldExists "Devices", "DisableStart", adInteger, "0", "0"
570     EnsureFieldExists "Devices", "DisableEnd", adInteger, "0", "0"


580     EnsureFieldExists "Devices", "SendCancel_A", adInteger, "0", "0"
590     EnsureFieldExists "Devices", "NG1_A", adInteger, "0", "0"
600     EnsureFieldExists "Devices", "NG2_A", adInteger, "0", "0"
610     EnsureFieldExists "Devices", "NG3_A", adInteger, "0", "0"
620     EnsureFieldExists "Devices", "AlarmMask_A", adInteger, "0", "0"
630     EnsureFieldExists "Devices", "UseAssur_A", adInteger, "0", "0"
640     EnsureFieldExists "Devices", "UseAssur2_A", adInteger, "0", "0"
650     EnsureFieldExists "Devices", "Announce_A", adVarWChar, 50, ""
660     EnsureFieldExists "Devices", "Repeats_A", adInteger, "0", "0"
670     EnsureFieldExists "Devices", "Pause_A", adInteger, "0", "0"
680     EnsureFieldExists "Devices", "RepeatUntil_A", adInteger, "0", "0"
690     EnsureFieldExists "Devices", "OG1_A", adInteger, "0", "0"
700     EnsureFieldExists "Devices", "OG2_A", adInteger, "0", "0"
710     EnsureFieldExists "Devices", "OG3_A", adInteger, "0", "0"
720     EnsureFieldExists "Devices", "VacationSuper_A", adInteger, "0", "0"
730     EnsureFieldExists "Devices", "DisableStart_A", adInteger, "0", "0"
740     EnsureFieldExists "Devices", "DisableEnd_A", adInteger, "0", "0"
750     EnsureFieldExists "Devices", "ResidentID_A", adInteger, "0", "0"
760     EnsureFieldExists "Devices", "RoomID_A", adInteger, "0", "0"
770     EnsureFieldExists "Devices", "ClearByReset_A", adInteger, "0", "0"
780     EnsureFieldExists "Devices", "AssurInput", adInteger, "0", "0"

790     EnsureFieldExists "DeviceTypes", "AllowDisable", adInteger, "0", "0"
800     EnsureFieldExists "DeviceTypes", "Announce", adVarWChar, 50, ""
810     EnsureFieldExists "DeviceTypes", "Announce2", adVarWChar, 50, ""
820     EnsureFieldExists "DeviceTypes", "IsPortable", adInteger, "0", "0"
830     EnsureFieldExists "DeviceTypes", "ClearByReset", adInteger, "0", "0"



840     EnsureFieldExists "Residents", "AssurDays", adInteger, "0", "0"
850     EnsureFieldExists "Residents", "Away", adInteger, "0", "0"
860     EnsureFieldExists "Residents", "Phone", adVarWChar, 50, ""
870     EnsureFieldExists "Residents", "Deleted", adInteger, "0", "0"

880     EnsureFieldExists "Rooms", "AssurDays", adInteger, "0", "0"
890     EnsureFieldExists "Rooms", "Away", adInteger, "0", "0"
900     EnsureFieldExists "Rooms", "Deleted", adInteger, "0", "0"

910     EnsureFieldExists "PagerDevices", "Pause", adInteger, "0", "0"
920     EnsureFieldExists "PagerDevices", "Deleted", adInteger, "0", "0"
930     EnsureFieldExists "PagerDevices", "IncludePhone", adInteger, "0", "0"
940     EnsureFieldExists "PagerDevices", "Pin", adVarWChar, 50, """"
950     EnsureFieldExists "PagerDevices", "KeyPA", adInteger, "0", "0"
960     EnsureFieldExists "PagerDevices", "Twice", adInteger, "0", "0"

970     EnsureFieldExists "Pagers", "IncludePhone", adInteger, "0", "0"


        ' do updates
        '880     Sql = "UPDATE DeviceTypes SET Model = 'EN1941' WHERE Model = 'ES1941'"
        '890     connexecute Sql
        '900     Sql = "UPDATE Devices SET Model = 'EN1941' WHERE Model = 'ES1941'"
        '910     connexecute Sql

        '920     Sql = "UPDATE DeviceTypes SET Model = 'EN3954' WHERE Model = 'ES3954'"
        '930     connexecute Sql
        '940     Sql = "UPDATE Devices SET Model = 'EN3954' WHERE Model = 'ES3954'"
        '950     connexecute Sql


Version190Plus:

980     EnsureFieldExists "CannedMessages", "Message", adVarWChar, 80, """"


990     EnsureFieldExists "ScreenMasks", "Screen", adInteger, "0", "0"
1000    EnsureFieldExists "ScreenMasks", "OG1", adInteger, "0", "0"
1010    EnsureFieldExists "ScreenMasks", "OG2", adInteger, "0", "0"
1020    EnsureFieldExists "ScreenMasks", "OG3", adInteger, "0", "0"
1030    EnsureFieldExists "ScreenMasks", "NG1", adInteger, "0", "0"
1040    EnsureFieldExists "ScreenMasks", "NG2", adInteger, "0", "0"
1050    EnsureFieldExists "ScreenMasks", "NG3", adInteger, "0", "0"
1060    EnsureFieldExists "ScreenMasks", "Repeats", adInteger, "0", "0"
1070    EnsureFieldExists "ScreenMasks", "RepeatUntil", adInteger, "0", "0"
1080    EnsureFieldExists "ScreenMasks", "SendCancel", adInteger, "0", "0"
1090    EnsureFieldExists "ScreenMasks", "Pause", adInteger, "0", "0"
1100    EnsureFieldExists "ScreenMasks", "ScreenName", adVarWChar, 50, """"


1110    EnsureFieldExists "Devices", "SerialSkip", adInteger, "0", "0"
1120    EnsureFieldExists "Devices", "SerialMessageLen", adInteger, "0", "0"
1130    EnsureFieldExists "Devices", "SerialAutoClear", adInteger, "0", "0"
1140    EnsureFieldExists "Devices", "SerialInclude", adVarWChar, 255, """"
1150    EnsureFieldExists "Devices", "SerialExclude", adVarWChar, 255, """"

1160    EnsureFieldExists "Devices", "SerialPort", adInteger, "0", "0"
1170    EnsureFieldExists "Devices", "SerialBaud", adInteger, "0", "0"
1180    EnsureFieldExists "Devices", "SerialBits", adInteger, "0", "0"
1190    EnsureFieldExists "Devices", "SerialParity", adVarWChar, 1, """"
1200    EnsureFieldExists "Devices", "SerialStopBits", adVarWChar, 3, """"
1210    EnsureFieldExists "Devices", "SerialFlow", adInteger, "0", "0"
1220    EnsureFieldExists "Devices", "SerialSettings", adVarWChar, 20, """"
1230    EnsureFieldExists "Devices", "SerialEOLChar", adInteger, "0", "0"

Version200Plus:

1240    EnsureFieldExists "Devices", "SerialPreamble", adVarWChar, 100, """"
        '

1250    If DBVER < 235 Then          ' we'll check every time up to version 235
          'new with build 226
1260      EnsureFieldExists "Devicetypes", "Repeats", adInteger, "0", "0"
1270      EnsureFieldExists "Devicetypes", "Pause", adInteger, "0", "0"
1280      EnsureFieldExists "Devicetypes", "repeatuntil", adInteger, "0", "0"
1290      EnsureFieldExists "Devicetypes", "SendCancel", adInteger, "0", "0"

1300      EnsureFieldExists "Devicetypes", "Repeats_A", adInteger, "0", "0"
1310      EnsureFieldExists "Devicetypes", "Pause_A", adInteger, "0", "0"
1320      EnsureFieldExists "Devicetypes", "repeatuntil_A", adInteger, "0", "0"
1330      EnsureFieldExists "Devicetypes", "SendCancel_A", adInteger, "0", "0"



1340      EnsureFieldExists "Devicetypes", "OG1", adInteger, "0", "0"
1350      EnsureFieldExists "Devicetypes", "OG2", adInteger, "0", "0"
1360      EnsureFieldExists "Devicetypes", "NG1", adInteger, "0", "0"
1370      EnsureFieldExists "Devicetypes", "NG2", adInteger, "0", "0"

1380      EnsureFieldExists "Devicetypes", "OG1_A", adInteger, "0", "0"
1390      EnsureFieldExists "Devicetypes", "OG2_A", adInteger, "0", "0"
1400      EnsureFieldExists "Devicetypes", "NG1_A", adInteger, "0", "0"
1410      EnsureFieldExists "Devicetypes", "NG2_A", adInteger, "0", "0"

1420    End If


1430    If DBVER < 250 Then          ' we'll check every time up to version 235
1440      EnsureFieldExists "PagerDevices", "DialerVoice", adVarWChar, 100, """"
1450      EnsureFieldExists "PagerDevices", "DialerModem", adInteger, "0", "0"
1460      EnsureFieldExists "PagerDevices", "DialerPhone", adVarWChar, 50, """"
1470      EnsureFieldExists "PagerDevices", "DialerTag", adVarWChar, 100, """"
1480      EnsureFieldExists "PagerDevices", "DialerMsgDelay", adInteger, "0", "0"
1490      EnsureFieldExists "PagerDevices", "DialerMsgRepeats", adInteger, "0", "0"
1500      EnsureFieldExists "PagerDevices", "DialerMsgSpacing", adInteger, "0", "0"
1510      EnsureFieldExists "PagerDevices", "DialerMsgSpacing", adInteger, "0", "0"
1520      EnsureFieldExists "PagerDevices", "DialerRedials", adInteger, "0", "0"
1530      EnsureFieldExists "PagerDevices", "DialerRedialDelay", adInteger, "0", "0"



1540    End If

1550    If DBVER < 470 Then          ' we'll check every time up to version 470
1560      EnsureFieldExists "PagerDevices", "DialerAckDigit", adInteger, "0", "0"
1570    End If

1580    If DBVER < 500 Then
1590      EnsureFieldExists "Devices", "Ignored", adInteger, "0", "0"

1600      EnsureFieldExists "PagerDevices", "MarquisCode", adInteger, "0", "0"
1610      EnsureFieldExists "Pagers", "MarquisCode", adInteger, "0", "0"

1620    End If


1630    If DBVER < 506 Then

1640      EnsureFieldExists "Pagers", "RelayNum", adInteger, "0", "0"
1650      EnsureFieldExists "PagerDevices", "Relay1", adInteger, "0", "0"
1660      EnsureFieldExists "PagerDevices", "Relay2", adInteger, "0", "0"
1670      EnsureFieldExists "PagerDevices", "Relay3", adInteger, "0", "0"
1680      EnsureFieldExists "PagerDevices", "Relay4", adInteger, "0", "0"

1690    End If

1700    If DBVER < 525 Then


1710      EnsureFieldExists "AutoReports", "ReportID", adInteger, "-1", "0"
1720      EnsureFieldExists "AutoReports", "Disabled", adInteger, "0", "0"
1730      EnsureFieldExists "AutoReports", "ReportName", adVarWChar, 50, """"
1740      EnsureFieldExists "AutoReports", "Comment", adVarWChar, 50, """"
1750      EnsureFieldExists "AutoReports", "Rooms", adLongVarWChar, 50, """"  ' memo field
1760      EnsureFieldExists "AutoReports", "Events", adVarWChar, 250, """"

1770      EnsureFieldExists "AutoReports", "TimePeriod", adInteger, "0", "0"
1780      EnsureFieldExists "AutoReports", "DayPeriod", adInteger, "0", "0"
1790      EnsureFieldExists "AutoReports", "Days", adInteger, "0", "0"
1800      EnsureFieldExists "AutoReports", "Shift", adInteger, "0", "0"
1810      EnsureFieldExists "AutoReports", "DayPartStart", adInteger, "0", "0"
1820      EnsureFieldExists "AutoReports", "DayPartEnd", adInteger, "0", "0"
1830      EnsureFieldExists "AutoReports", "SortOrder", adInteger, "0", "0"
1840      EnsureFieldExists "AutoReports", "SendHour", adInteger, "0", "0"

1850      EnsureFieldExists "AutoReports", "SaveAsFile", adInteger, "0", "0"
1860      EnsureFieldExists "AutoReports", "DestFolder", adVarWChar, 255, """"

1870      EnsureFieldExists "AutoReports", "SendAsEmail", adInteger, "0", "0"

1880      EnsureFieldExists "AutoReports", "Recipient", adVarWChar, 150, """"
1890      EnsureFieldExists "AutoReports", "Subject", adVarWChar, 150, """"

1900      EnsureFieldExists "AutoReports", "FileFormat", adInteger, "0", "0"


          '  ReportID = rs("reportid")
          '  Enabled = rs("enabled")
          '  ReportName = rs("reportname") & ""
          '  Comment = rs("Comment ") & ""
          '  RoomString = rs("Rooms") & ""
          '  EventString = rs("Events") & ""
          '  TimePeriod = rs("TimePeriod")
          '  DayPeriod = rs("DayPeriod")
          '  Days = rs("Days")
          '  Shift = rs("Shift")
          '  DayPartStart = rs("DayPartStart")
          '  DayPartEnd = rs("DayPartend")
          '  SortOrder = rs("SortOrder")
          '  SendHour = rs("SendHour")
          '  SaveAsFile = 1  '  rs("SaveAsFile")
          '  SendAsEmail = rs("SendAsEmail")
          '  Recipient = rs("Recipient") & ""
          '  Subject = rs("Subject") & ""
          '  FileFormat = rs("FileFormat")
          '  DestFolder = rs("DestFolder") & ""


1910    End If


1920    If DBVER < 530 Then
1930      EnsureFieldExists "ReminderSubscribers", "ReminderID", adInteger, "0", "0"
1940      EnsureFieldExists "ReminderSubscribers", "SubscriberID", adInteger, "0", "0"
1950      EnsureFieldExists "ReminderSubscribers", "ResidentID", adInteger, "0", "0"

1960      EnsureFieldExists "Reminders", "ReminderID", adInteger, "-1", "0"
1970      EnsureFieldExists "Reminders", "OwnerID", adInteger, "0", "0"
1980      EnsureFieldExists "Reminders", "ResidentID", adInteger, "0", "0"
1990      EnsureFieldExists "Reminders", "IsPublic", adInteger, "0", "0"

2000      EnsureFieldExists "Reminders", "Description", adVarWChar, 150, """"
2010      EnsureFieldExists "Reminders", "Message", adVarWChar, 150, """"

2020      EnsureFieldExists "Reminders", "Coordinator", adVarWChar, 50, """"
2030      EnsureFieldExists "Reminders", "LeadTime", adInteger, "0", "0"
2040      EnsureFieldExists "Reminders", "Disabled", adInteger, "0", "0"
2050      EnsureFieldExists "Reminders", "Cancelled", adInteger, "0", "0"
2060      EnsureFieldExists "Reminders", "Recurring", adInteger, "0", "0"
2070      EnsureFieldExists "Reminders", "Frequency", adInteger, "0", "0"
2080      EnsureFieldExists "Reminders", "DaysActive", adVarWChar, 150, """"
2090      EnsureFieldExists "Reminders", "SpecificDay", adVarWChar, 20, """"
2100      EnsureFieldExists "Reminders", "TimeOfDay", adInteger, "0", "0"  'in hhmm 0000 , 0600, 0630

2110      EnsureFieldExists "Reminders", "DOW", adInteger, "0", "0"  'bitfield
2120      EnsureFieldExists "Reminders", "DOM", adInteger, "0", "0"  'value


          ' staff table mimics residents table

2130      EnsureFieldExists "Staff", "StaffID", adInteger, "-1", "0"
2140      EnsureFieldExists "Staff", "Name", adVarWChar, 50, ""
2150      EnsureFieldExists "Staff", "NameLast", adVarWChar, 50, ""
2160      EnsureFieldExists "Staff", "NameFirst", adVarWChar, 50, ""
2170      EnsureFieldExists "Staff", "Room", adVarWChar, 50, ""
2180      EnsureFieldExists "Staff", "Active", adInteger, "0", "0"
2190      EnsureFieldExists "Staff", "GroupID", adInteger, "0", "0"
2200      EnsureFieldExists "Staff", "RoomID", adInteger, "0", "0"
2210      EnsureFieldExists "Staff", "ImagePath", adVarWChar, 255, ""
2220      EnsureFieldExists "Staff", "ImageData", adLongVarBinary, "", ""



2230      EnsureFieldExists "Staff", "AssurDays", adInteger, "0", "0"
2240      EnsureFieldExists "Staff", "Vacation", adInteger, "0", "0"
2250      EnsureFieldExists "Staff", "Away", adInteger, "0", "0"
2260      EnsureFieldExists "Staff", "Info", adVarWChar, 255, ""
2270      EnsureFieldExists "Staff", "Phone", adVarWChar, 50, ""
2280      EnsureFieldExists "Staff", "Deleted", adInteger, "0", "0"



2290    End If

2300    If DBVER < 550 Then          ' adding more escalation levels/groups

2310      EnsureFieldExists "Devices", "OG4", adInteger, "0", "0"
2320      EnsureFieldExists "Devices", "OG5", adInteger, "0", "0"
2330      EnsureFieldExists "Devices", "OG6", adInteger, "0", "0"

2340      EnsureFieldExists "Devices", "NG4", adInteger, "0", "0"
2350      EnsureFieldExists "Devices", "NG5", adInteger, "0", "0"
2360      EnsureFieldExists "Devices", "NG6", adInteger, "0", "0"

2370      EnsureFieldExists "Devices", "OG4_A", adInteger, "0", "0"
2380      EnsureFieldExists "Devices", "OG5_A", adInteger, "0", "0"
2390      EnsureFieldExists "Devices", "OG6_A", adInteger, "0", "0"

2400      EnsureFieldExists "Devices", "NG4_A", adInteger, "0", "0"
2410      EnsureFieldExists "Devices", "NG5_A", adInteger, "0", "0"
2420      EnsureFieldExists "Devices", "NG6_A", adInteger, "0", "0"

2430      EnsureFieldExists "Devicetypes", "OG3", adInteger, "0", "0"
2440      EnsureFieldExists "Devicetypes", "OG4", adInteger, "0", "0"
2450      EnsureFieldExists "Devicetypes", "OG5", adInteger, "0", "0"
2460      EnsureFieldExists "Devicetypes", "OG6", adInteger, "0", "0"


2470      EnsureFieldExists "Devicetypes", "NG3", adInteger, "0", "0"
2480      EnsureFieldExists "Devicetypes", "NG4", adInteger, "0", "0"
2490      EnsureFieldExists "Devicetypes", "NG5", adInteger, "0", "0"
2500      EnsureFieldExists "Devicetypes", "NG6", adInteger, "0", "0"


2510      EnsureFieldExists "Devicetypes", "OG3_A", adInteger, "0", "0"
2520      EnsureFieldExists "Devicetypes", "OG4_A", adInteger, "0", "0"
2530      EnsureFieldExists "Devicetypes", "OG5_A", adInteger, "0", "0"
2540      EnsureFieldExists "Devicetypes", "OG6_A", adInteger, "0", "0"


2550      EnsureFieldExists "Devicetypes", "NG3_A", adInteger, "0", "0"
2560      EnsureFieldExists "Devicetypes", "NG4_A", adInteger, "0", "0"
2570      EnsureFieldExists "Devicetypes", "NG5_A", adInteger, "0", "0"
2580      EnsureFieldExists "Devicetypes", "NG6_A", adInteger, "0", "0"



2590      EnsureFieldExists "ScreenMasks", "OG4", adInteger, "0", "0"
2600      EnsureFieldExists "ScreenMasks", "OG5", adInteger, "0", "0"
2610      EnsureFieldExists "ScreenMasks", "OG6", adInteger, "0", "0"

2620      EnsureFieldExists "ScreenMasks", "NG4", adInteger, "0", "0"
2630      EnsureFieldExists "ScreenMasks", "NG5", adInteger, "0", "0"
2640      EnsureFieldExists "ScreenMasks", "NG6", adInteger, "0", "0"




2650    End If

2660    If DBVER < 555 Then          ' adding more escalation levels/groups

2670      EnsureFieldExists "Devices", "OG1d", adInteger, "0", "0"
2680      EnsureFieldExists "Devices", "OG2d", adInteger, "0", "0"
2690      EnsureFieldExists "Devices", "OG3d", adInteger, "0", "0"
2700      EnsureFieldExists "Devices", "OG4d", adInteger, "0", "0"
2710      EnsureFieldExists "Devices", "OG5d", adInteger, "0", "0"
2720      EnsureFieldExists "Devices", "OG6d", adInteger, "0", "0"

2730      EnsureFieldExists "Devices", "NG1d", adInteger, "0", "0"
2740      EnsureFieldExists "Devices", "NG2d", adInteger, "0", "0"
2750      EnsureFieldExists "Devices", "NG3d", adInteger, "0", "0"
2760      EnsureFieldExists "Devices", "NG4d", adInteger, "0", "0"
2770      EnsureFieldExists "Devices", "NG5d", adInteger, "0", "0"
2780      EnsureFieldExists "Devices", "NG6d", adInteger, "0", "0"


2790      EnsureFieldExists "Devices", "OG1_Ad", adInteger, "0", "0"
2800      EnsureFieldExists "Devices", "OG2_Ad", adInteger, "0", "0"
2810      EnsureFieldExists "Devices", "OG3_Ad", adInteger, "0", "0"
2820      EnsureFieldExists "Devices", "OG4_Ad", adInteger, "0", "0"
2830      EnsureFieldExists "Devices", "OG5_Ad", adInteger, "0", "0"
2840      EnsureFieldExists "Devices", "OG6_Ad", adInteger, "0", "0"

2850      EnsureFieldExists "Devices", "NG1_Ad", adInteger, "0", "0"
2860      EnsureFieldExists "Devices", "NG2_Ad", adInteger, "0", "0"
2870      EnsureFieldExists "Devices", "NG3_Ad", adInteger, "0", "0"
2880      EnsureFieldExists "Devices", "NG4_Ad", adInteger, "0", "0"
2890      EnsureFieldExists "Devices", "NG5_Ad", adInteger, "0", "0"
2900      EnsureFieldExists "Devices", "NG6_Ad", adInteger, "0", "0"



2910      EnsureFieldExists "DeviceTypes", "OG1d", adInteger, "0", "0"
2920      EnsureFieldExists "DeviceTypes", "OG2d", adInteger, "0", "0"
2930      EnsureFieldExists "DeviceTypes", "OG3d", adInteger, "0", "0"
2940      EnsureFieldExists "DeviceTypes", "OG4d", adInteger, "0", "0"
2950      EnsureFieldExists "DeviceTypes", "OG5d", adInteger, "0", "0"
2960      EnsureFieldExists "DeviceTypes", "OG6d", adInteger, "0", "0"

2970      EnsureFieldExists "DeviceTypes", "NG1d", adInteger, "0", "0"
2980      EnsureFieldExists "DeviceTypes", "NG2d", adInteger, "0", "0"
2990      EnsureFieldExists "DeviceTypes", "NG3d", adInteger, "0", "0"
3000      EnsureFieldExists "DeviceTypes", "NG4d", adInteger, "0", "0"
3010      EnsureFieldExists "DeviceTypes", "NG5d", adInteger, "0", "0"
3020      EnsureFieldExists "DeviceTypes", "NG6d", adInteger, "0", "0"


3030      EnsureFieldExists "DeviceTypes", "OG1_Ad", adInteger, "0", "0"
3040      EnsureFieldExists "DeviceTypes", "OG2_Ad", adInteger, "0", "0"
3050      EnsureFieldExists "DeviceTypes", "OG3_Ad", adInteger, "0", "0"
3060      EnsureFieldExists "DeviceTypes", "OG4_Ad", adInteger, "0", "0"
3070      EnsureFieldExists "DeviceTypes", "OG5_Ad", adInteger, "0", "0"
3080      EnsureFieldExists "DeviceTypes", "OG6_Ad", adInteger, "0", "0"

3090      EnsureFieldExists "DeviceTypes", "NG1_Ad", adInteger, "0", "0"
3100      EnsureFieldExists "DeviceTypes", "NG2_Ad", adInteger, "0", "0"
3110      EnsureFieldExists "DeviceTypes", "NG3_Ad", adInteger, "0", "0"
3120      EnsureFieldExists "DeviceTypes", "NG4_Ad", adInteger, "0", "0"
3130      EnsureFieldExists "DeviceTypes", "NG5_Ad", adInteger, "0", "0"
3140      EnsureFieldExists "DeviceTypes", "NG6_Ad", adInteger, "0", "0"


3150      EnsureFieldExists "ScreenMasks", "OG1d", adInteger, "0", "0"
3160      EnsureFieldExists "ScreenMasks", "OG2d", adInteger, "0", "0"
3170      EnsureFieldExists "ScreenMasks", "OG3d", adInteger, "0", "0"
3180      EnsureFieldExists "ScreenMasks", "OG4d", adInteger, "0", "0"
3190      EnsureFieldExists "ScreenMasks", "OG5d", adInteger, "0", "0"
3200      EnsureFieldExists "ScreenMasks", "OG6d", adInteger, "0", "0"

3210      EnsureFieldExists "ScreenMasks", "NG1d", adInteger, "0", "0"
3220      EnsureFieldExists "ScreenMasks", "NG2d", adInteger, "0", "0"
3230      EnsureFieldExists "ScreenMasks", "NG3d", adInteger, "0", "0"
3240      EnsureFieldExists "ScreenMasks", "NG4d", adInteger, "0", "0"
3250      EnsureFieldExists "ScreenMasks", "NG5d", adInteger, "0", "0"
3260      EnsureFieldExists "ScreenMasks", "NG6d", adInteger, "0", "0"




3270    End If

3280    If DBVER < 600 Then          ' TAP protocol for input device
3290      EnsureFieldExists "Devices", "SerialTapProtocol", adInteger, "0", "0"


3300    End If

3310    If DBVER < 600 Then          ' Reminder QUE
3320      EnsureFieldExists "Reminders", "DeliveryPoint", adVarWChar, 255, """"

3330      EnsureFieldExists "Residents", "DeliveryPoints", adVarWChar, 255, """"
3340      EnsureFieldExists "Staff", "DeliveryPoints", adVarWChar, 255, """"

3350      ExpandResInfo              ' does res table
3360      ExpandStaffInfo            ' does staff table


3370      EnsureFieldExists "ReminderSubscribers", "PagerID", adInteger, "0", "0"
3380      EnsureFieldExists "ReminderSubscribers", "GroupID", adInteger, "0", "0"


3390      EnsureFieldExists "PagerDevices", "Relay5", adInteger, "0", "0"
3400      EnsureFieldExists "PagerDevices", "Relay6", adInteger, "0", "0"
3410      EnsureFieldExists "PagerDevices", "Relay7", adInteger, "0", "0"
3420      EnsureFieldExists "PagerDevices", "Relay8", adInteger, "0", "0"

3430    End If

3440    If DBVER < 650 Then
3450      EnsureFieldExists "ThinSessions", "ID", adInteger, "-1", "0"
3460      EnsureFieldExists "ThinSessions", "username", adVarWChar, 255, """"
3470      EnsureFieldExists "ThinSessions", "sessionid", adVarWChar, 255, """"
3480      EnsureFieldExists "ThinSessions", "sessiontime", adDate, "", ""
3490      EnsureFieldExists "ThinSessions", "IP", adVarWChar, 255, """"

3500      EnsureFieldExists "ExternalPages", "ID", adInteger, "-1", "0"
3510      EnsureFieldExists "ExternalPages", "groupid", adInteger, "0", "0"
3520      EnsureFieldExists "ExternalPages", "pagerid", adInteger, "0", "0"
3530      EnsureFieldExists "ExternalPages", "message", adVarWChar, 255, """"

3540      EnsureFieldExists "Pagers", "NoName", adInteger, "0", "0"


3550    End If



3560    If DBVER < 650 Then
3570      EnsureFieldExists "Devicetypes", "ignoretamper", adInteger, "0", "0"
3580      EnsureFieldExists "Devices", "ignoretamper", adInteger, "0", "0"
3590    End If




3600    If DBVER < 667 Then
          ' add field for custom device
3610      EnsureFieldExists "Devices", "custom", adVarWChar, 50, """"
          ' populate custom field with default
3620      UpdateCustom


          ' add ExeptionReports Table (not all fields will be used
3630      EnsureFieldExists "ExceptionReports", "ReportID", adInteger, "-1", "0"
3640      EnsureFieldExists "ExceptionReports", "Disabled", adInteger, "0", "0"
3650      EnsureFieldExists "ExceptionReports", "ReportName", adVarWChar, 50, """"
3660      EnsureFieldExists "ExceptionReports", "Comment", adVarWChar, 50, """"
3670      EnsureFieldExists "ExceptionReports", "Rooms", adLongVarWChar, 50, """"  ' memo field
3680      EnsureFieldExists "ExceptionReports", "Events", adLongVarWChar, 50, """"  ' memo as well
3690      EnsureFieldExists "ExceptionReports", "DevTypes", adLongVarWChar, 50, """"  ' memo as well
3700      EnsureFieldExists "ExceptionReports", "TimePeriod", adInteger, "0", "0"
3710      EnsureFieldExists "ExceptionReports", "DayPeriod", adInteger, "0", "0"
3720      EnsureFieldExists "ExceptionReports", "Days", adInteger, "0", "0"
3730      EnsureFieldExists "ExceptionReports", "Shift", adInteger, "0", "0"
3740      EnsureFieldExists "ExceptionReports", "DayPartStart", adInteger, "0", "0"
3750      EnsureFieldExists "ExceptionReports", "DayPartEnd", adInteger, "0", "0"
3760      EnsureFieldExists "ExceptionReports", "SortOrder", adInteger, "0", "0"
3770      EnsureFieldExists "ExceptionReports", "SendHour", adInteger, "0", "0"
3780      EnsureFieldExists "ExceptionReports", "SaveAsFile", adInteger, "0", "0"
3790      EnsureFieldExists "ExceptionReports", "DestFolder", adVarWChar, 255, """"
3800      EnsureFieldExists "ExceptionReports", "SendAsEmail", adInteger, "0", "0"
3810      EnsureFieldExists "ExceptionReports", "Recipient", adVarWChar, 150, """"
3820      EnsureFieldExists "ExceptionReports", "Subject", adVarWChar, 150, """"
3830      EnsureFieldExists "ExceptionReports", "FileFormat", adInteger, "0", "0"
3840      EnsureFieldExists "ExceptionReports", "ResponseTime", adInteger, "0", "0"
3850      EnsureFieldExists "ExceptionReports", "ResponseIsACK", adInteger, "0", "0"
3860      EnsureFieldExists "ExceptionReports", "ReportType", adInteger, "0", "0"


3870    End If

3880    If DBVER < 675 Then          ' added 1723 temperature device
3890      EnsureFieldExists "Devices", "EnableTemp", adInteger, "0", "0"  ' 0=disabled, 1=rise, 2=fall
3900      EnsureFieldExists "Devices", "EnableTemp_A", adInteger, "0", "0"
3910      EnsureFieldExists "Devices", "LowSet", adDouble, "0", "0"
3920      EnsureFieldExists "Devices", "LowSet_A", adDouble, "0", "0"
3930      EnsureFieldExists "Devices", "HiSet", adDouble, "0", "0"
3940      EnsureFieldExists "Devices", "HiSet_A", adDouble, "0", "0"

3950    End If


3960    If DBVER < 695 Then          ' upgrade to 3 shifts

          ' upgrade to 3 shifts
3970      EnsureFieldExists "Devices", "GG1", adInteger, "0", "0"
3980      EnsureFieldExists "Devices", "GG2", adInteger, "0", "0"
3990      EnsureFieldExists "Devices", "GG3", adInteger, "0", "0"
4000      EnsureFieldExists "Devices", "GG4", adInteger, "0", "0"
4010      EnsureFieldExists "Devices", "GG5", adInteger, "0", "0"
4020      EnsureFieldExists "Devices", "GG6", adInteger, "0", "0"

4030      EnsureFieldExists "Devices", "GG1d", adInteger, "0", "0"
4040      EnsureFieldExists "Devices", "GG2d", adInteger, "0", "0"
4050      EnsureFieldExists "Devices", "GG3d", adInteger, "0", "0"
4060      EnsureFieldExists "Devices", "GG4d", adInteger, "0", "0"
4070      EnsureFieldExists "Devices", "GG5d", adInteger, "0", "0"
4080      EnsureFieldExists "Devices", "GG6d", adInteger, "0", "0"

4090      EnsureFieldExists "Devices", "GG1_A", adInteger, "0", "0"
4100      EnsureFieldExists "Devices", "GG2_A", adInteger, "0", "0"
4110      EnsureFieldExists "Devices", "GG3_A", adInteger, "0", "0"
4120      EnsureFieldExists "Devices", "GG4_A", adInteger, "0", "0"
4130      EnsureFieldExists "Devices", "GG5_A", adInteger, "0", "0"
4140      EnsureFieldExists "Devices", "GG6_A", adInteger, "0", "0"

4150      EnsureFieldExists "Devices", "GG1_Ad", adInteger, "0", "0"
4160      EnsureFieldExists "Devices", "GG2_Ad", adInteger, "0", "0"
4170      EnsureFieldExists "Devices", "GG3_Ad", adInteger, "0", "0"
4180      EnsureFieldExists "Devices", "GG4_Ad", adInteger, "0", "0"
4190      EnsureFieldExists "Devices", "GG5_Ad", adInteger, "0", "0"
4200      EnsureFieldExists "Devices", "GG6_Ad", adInteger, "0", "0"

          ' and device types
4210      EnsureFieldExists "Devicetypes", "GG1", adInteger, "0", "0"
4220      EnsureFieldExists "Devicetypes", "GG2", adInteger, "0", "0"
4230      EnsureFieldExists "Devicetypes", "GG3", adInteger, "0", "0"
4240      EnsureFieldExists "Devicetypes", "GG4", adInteger, "0", "0"
4250      EnsureFieldExists "Devicetypes", "GG5", adInteger, "0", "0"
4260      EnsureFieldExists "Devicetypes", "GG6", adInteger, "0", "0"

4270      EnsureFieldExists "Devicetypes", "GG1d", adInteger, "0", "0"
4280      EnsureFieldExists "Devicetypes", "GG2d", adInteger, "0", "0"
4290      EnsureFieldExists "Devicetypes", "GG3d", adInteger, "0", "0"
4300      EnsureFieldExists "Devicetypes", "GG4d", adInteger, "0", "0"
4310      EnsureFieldExists "Devicetypes", "GG5d", adInteger, "0", "0"
4320      EnsureFieldExists "Devicetypes", "GG6d", adInteger, "0", "0"

4330      EnsureFieldExists "Devicetypes", "GG1_A", adInteger, "0", "0"
4340      EnsureFieldExists "Devicetypes", "GG2_A", adInteger, "0", "0"
4350      EnsureFieldExists "Devicetypes", "GG3_A", adInteger, "0", "0"
4360      EnsureFieldExists "Devicetypes", "GG4_A", adInteger, "0", "0"
4370      EnsureFieldExists "Devicetypes", "GG5_A", adInteger, "0", "0"
4380      EnsureFieldExists "Devicetypes", "GG6_A", adInteger, "0", "0"

4390      EnsureFieldExists "Devicetypes", "GG1_Ad", adInteger, "0", "0"
4400      EnsureFieldExists "Devicetypes", "GG2_Ad", adInteger, "0", "0"
4410      EnsureFieldExists "Devicetypes", "GG3_Ad", adInteger, "0", "0"
4420      EnsureFieldExists "Devicetypes", "GG4_Ad", adInteger, "0", "0"
4430      EnsureFieldExists "Devicetypes", "GG5_Ad", adInteger, "0", "0"
4440      EnsureFieldExists "Devicetypes", "GG6_Ad", adInteger, "0", "0"

          ' and masks

4450      EnsureFieldExists "ScreenMasks", "GG1", adInteger, "0", "0"
4460      EnsureFieldExists "ScreenMasks", "GG2", adInteger, "0", "0"
4470      EnsureFieldExists "ScreenMasks", "GG3", adInteger, "0", "0"
4480      EnsureFieldExists "ScreenMasks", "GG4", adInteger, "0", "0"
4490      EnsureFieldExists "ScreenMasks", "GG5", adInteger, "0", "0"
4500      EnsureFieldExists "ScreenMasks", "GG6", adInteger, "0", "0"

4510      EnsureFieldExists "ScreenMasks", "GG1d", adInteger, "0", "0"
4520      EnsureFieldExists "ScreenMasks", "GG2d", adInteger, "0", "0"
4530      EnsureFieldExists "ScreenMasks", "GG3d", adInteger, "0", "0"
4540      EnsureFieldExists "ScreenMasks", "GG4d", adInteger, "0", "0"
4550      EnsureFieldExists "ScreenMasks", "GG5d", adInteger, "0", "0"
4560      EnsureFieldExists "ScreenMasks", "GG6d", adInteger, "0", "0"

4570      EnsureFieldExists "ScreenMasks", "GG1_A", adInteger, "0", "0"
4580      EnsureFieldExists "ScreenMasks", "GG2_A", adInteger, "0", "0"
4590      EnsureFieldExists "ScreenMasks", "GG3_A", adInteger, "0", "0"
4600      EnsureFieldExists "ScreenMasks", "GG4_A", adInteger, "0", "0"
4610      EnsureFieldExists "ScreenMasks", "GG5_A", adInteger, "0", "0"
4620      EnsureFieldExists "ScreenMasks", "GG6_A", adInteger, "0", "0"

4630      EnsureFieldExists "ScreenMasks", "GG1_Ad", adInteger, "0", "0"
4640      EnsureFieldExists "ScreenMasks", "GG2_Ad", adInteger, "0", "0"
4650      EnsureFieldExists "ScreenMasks", "GG3_Ad", adInteger, "0", "0"
4660      EnsureFieldExists "ScreenMasks", "GG4_Ad", adInteger, "0", "0"
4670      EnsureFieldExists "ScreenMasks", "GG5_Ad", adInteger, "0", "0"
4680      EnsureFieldExists "ScreenMasks", "GG6_Ad", adInteger, "0", "0"


4690    End If

4700    If DBVER < 770 Then
4710      EnsureFieldExists "pagerdevices", "lf", adInteger, "0", "0"

          ' added for third input

4720      EnsureFieldExists "Devices", "UseTamperAsInput", adInteger, "0", "0"

4730      EnsureFieldExists "Devices", "OG1_b", adInteger, "0", "0"
4740      EnsureFieldExists "Devices", "OG2_b", adInteger, "0", "0"
4750      EnsureFieldExists "Devices", "OG3_b", adInteger, "0", "0"
4760      EnsureFieldExists "Devices", "OG4_b", adInteger, "0", "0"
4770      EnsureFieldExists "Devices", "OG5_b", adInteger, "0", "0"
4780      EnsureFieldExists "Devices", "OG6_b", adInteger, "0", "0"

4790      EnsureFieldExists "Devices", "NG1_b", adInteger, "0", "0"
4800      EnsureFieldExists "Devices", "NG2_b", adInteger, "0", "0"
4810      EnsureFieldExists "Devices", "NG3_b", adInteger, "0", "0"
4820      EnsureFieldExists "Devices", "NG4_b", adInteger, "0", "0"
4830      EnsureFieldExists "Devices", "NG5_b", adInteger, "0", "0"
4840      EnsureFieldExists "Devices", "NG6_b", adInteger, "0", "0"

4850      EnsureFieldExists "Devices", "GG1_b", adInteger, "0", "0"
4860      EnsureFieldExists "Devices", "GG2_b", adInteger, "0", "0"
4870      EnsureFieldExists "Devices", "GG3_b", adInteger, "0", "0"
4880      EnsureFieldExists "Devices", "GG4_b", adInteger, "0", "0"
4890      EnsureFieldExists "Devices", "GG5_b", adInteger, "0", "0"
4900      EnsureFieldExists "Devices", "GG6_b", adInteger, "0", "0"




4910      EnsureFieldExists "Devices", "OG1_bd", adInteger, "0", "0"
4920      EnsureFieldExists "Devices", "OG2_bd", adInteger, "0", "0"
4930      EnsureFieldExists "Devices", "OG3_bd", adInteger, "0", "0"
4940      EnsureFieldExists "Devices", "OG4_bd", adInteger, "0", "0"
4950      EnsureFieldExists "Devices", "OG5_bd", adInteger, "0", "0"
4960      EnsureFieldExists "Devices", "OG6_bd", adInteger, "0", "0"

4970      EnsureFieldExists "Devices", "NG1_bd", adInteger, "0", "0"
4980      EnsureFieldExists "Devices", "NG2_bd", adInteger, "0", "0"
4990      EnsureFieldExists "Devices", "NG3_bd", adInteger, "0", "0"
5000      EnsureFieldExists "Devices", "NG4_bd", adInteger, "0", "0"
5010      EnsureFieldExists "Devices", "NG5_bd", adInteger, "0", "0"
5020      EnsureFieldExists "Devices", "NG6_bd", adInteger, "0", "0"


5030      EnsureFieldExists "Devices", "GG1_bd", adInteger, "0", "0"
5040      EnsureFieldExists "Devices", "GG2_bd", adInteger, "0", "0"
5050      EnsureFieldExists "Devices", "GG3_bd", adInteger, "0", "0"
5060      EnsureFieldExists "Devices", "GG4_bd", adInteger, "0", "0"
5070      EnsureFieldExists "Devices", "GG5_bd", adInteger, "0", "0"
5080      EnsureFieldExists "Devices", "GG6_bd", adInteger, "0", "0"

5090      EnsureFieldExists "Devices", "Announce_B", adVarWChar, 50, ""

5100      EnsureFieldExists "Devices", "UseAssur_B", adInteger, "0", "0"
5110      EnsureFieldExists "Devices", "UseAssur2_B", adInteger, "0", "0"
5120      EnsureFieldExists "Devices", "SendCancel_B", adInteger, "0", "0"
5130      EnsureFieldExists "Devices", "repeatuntil_B", adInteger, "0", "0"
5140      EnsureFieldExists "Devices", "repeats_B", adInteger, "0", "0"
5150      EnsureFieldExists "Devices", "Pause_B", adInteger, "0", "0"
5160      EnsureFieldExists "Devices", "VacationSuper_B", adInteger, "0", "0"
5170      EnsureFieldExists "Devices", "AlarmMask_B", adInteger, "0", "0"



5180      EnsureFieldExists "Devices", "disablestart_B", adInteger, "0", "0"
5190      EnsureFieldExists "Devices", "disableEnd_B", adInteger, "0", "0"

5200    End If



5210    If DBVER < 830 Then
5220      EnsureFieldExists "Devices", "LocKW", adVarWChar, 255, ""
5230      EnsureFieldExists "Rooms", "LocKW", adVarWChar, 255, ""
5240    End If

5250    If DBVER < 900 Then
5260      EnsureFieldExists "Rooms", "Flags", adInteger, "0", "0"
5270    End If

5280    If DBVER < 1005 Then

          '' FOR MOBILE APP
5290      EnsureFieldExists "Mobile", "ID", adInteger, "-1", "0"
5300      EnsureFieldExists "Mobile", "AlarmID", adInteger, "0", "0"
5310      EnsureFieldExists "Mobile", "PagerID", adVarWChar, 255, """"  ' destination pager(s)
5320      EnsureFieldExists "Mobile", "Message", adVarWChar, 255, """"  ' not used anymore
5330      EnsureFieldExists "Mobile", "TimeAdded", adDate, "", ""  ' when alarm is CREATED
5340      EnsureFieldExists "Mobile", "TimeAcked", adVarWChar, 55, """"  ' when it's ACKED
5350      EnsureFieldExists "Mobile", "AckUser", adVarWChar, 255, """"  ' Who ACKED it
5360      EnsureFieldExists "Mobile", "sent", adInteger, "0", "0"  ' probably not goint to use this
5370      EnsureFieldExists "Mobile", "eAlarmTime", adVarWChar, 25, """"  ' unix epoc time
5380      EnsureFieldExists "Mobile", "eAckTime", adVarWChar, 25, """"  ' unix epoc time


          ' to enable choice where hitting the ACK # on phone still allows all other phones and devices to receive pages
5390      EnsureFieldExists "PagerDevices", "KeepPaging", adInteger, "0", "0"
          ' might need this for lazy hangups
5400      EnsureFieldExists "PagerDevices", "Timeout", adInteger, "0", "0"
5410    End If
5420    If DBVER < 1010 Then
5430      EnsureFieldExists "Users", "Permissions", adInteger, "0", "0"
5440    End If

5450    If DBVER < 1030 Then         ' 1023
5460      EnsureFieldExists "Devices", "ConfigurationString", adVarWChar, 50, """"
5470    End If

        If DBVER < 1040 Then ' 2018-08-03
          EnsureFieldExists "Dispositions", "ID", adInteger, "-1", "0"
          EnsureFieldExists "Dispositions", "Text", adVarWChar, 255, """"
          
          EnsureFieldExists "Packets", "ID", adInteger, "-1", "0"
          EnsureFieldExists "Packets", "Posted", adInteger, "0", "0"
          EnsureFieldExists "Packets", "PostDate", adDate, "", ""
          EnsureFieldExists "Packets", "Session", adInteger, "0", "0"
          EnsureFieldExists "Packets", "Text", adVarWChar, 2048, """"
          
          EnsureFieldExists "Mobile", "Ended", adInteger, "0", "0"
          
          
        End If

        If DBVER < 1050 Then ' 2018-08-03
          EnsureFieldExists "Alarms", "InputNum", adInteger, "0", "0"
        End If

        Dim rs                 As ADODB.Recordset
        Dim Colsize            As Long

5480    Set rs = ConnExecute("select recipient from ExceptionReports")
5490    If Not rs.EOF Then
5500      Colsize = rs("recipient").DefinedSize
5510    End If
5520    rs.Close

5530    If Colsize > 0 And Colsize < 255 Then

5540      If gIsJET Then

5550        SQL = "ALTER TABLE ExceptionReports ALTER COLUMN Recipient text(255)"
5560        ConnExecute SQL
5570      Else
5580        SQL = "ALTER TABLE ExceptionReports ALTER COLUMN Recipient NVARCHAR(255)"
5590        ConnExecute SQL
5600      End If

5610    End If

5620    Colsize = 0

5630    Set rs = ConnExecute("select recipient from AutoReports")
5640    If Not rs.EOF Then
5650      Colsize = rs("recipient").DefinedSize
5660    End If
5670    rs.Close

5680    If Colsize > 0 And Colsize < 255 Then

5690      If gIsJET Then
5700        SQL = "ALTER TABLE AutoReports ALTER COLUMN Recipient text(255)"
5710        ConnExecute SQL
5720      Else
5730        SQL = "ALTER TABLE AutoReports ALTER COLUMN Recipient NVARCHAR(255)"
5740        ConnExecute SQL
5750      End If

5760    End If

5770    Set rs = Nothing


5780    If Err.Number = 0 Then
5790      WriteDBVersion
5800      UpdateCLSPTIs              ' only called if jet database
5810      Call WriteSetting("Version", "Major", App.Major)
5820      Call WriteSetting("Version", "Minor", App.Minor)
5830      Call WriteSetting("Version", "Build", App.Revision)

5840    End If



FixDatabase_Resume:
5850    On Error GoTo 0
5860    Exit Function

FixDatabase_Error:

5870    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.FixDatabase." & Erl
5880    Resume FixDatabase_Resume


End Function

Function UpdateCustom()
  Dim SQL           As String
  Dim DeviceType    As ESDeviceTypeType
  Dim j             As Integer
  Debug.Print "FillCustom"

  For j = LBound(ESDeviceType) To UBound(ESDeviceType)
    DeviceType = ESDeviceType(j)
    DeviceType.desc = Trim$(Replace(DeviceType.desc, "'", " "))
    SQL = "UPDATE Devices set Custom = '" & DeviceType.desc & "' WHERE Model = '" & DeviceType.Model & "'"
    ConnExecute SQL
  Next


End Function


Function ExpandResInfo()


  Dim tbl           As ADOX.Table
  Dim cat           As ADOX.Catalog
  Dim fld           As ADOX.Column
  Dim i             As Integer
  Dim SQL           As String
  Dim tempfld       As ADOX.Column


  Connect

  Set cat = New ADOX.Catalog
  cat.ActiveConnection = conn
  On Error Resume Next
  Set tbl = cat.Tables("Residents")
  If Not (tbl Is Nothing) Then

    For i = 0 To tbl.Columns.Count - 1
      If 0 = StrComp(tbl.Columns(i).name, "Info", vbTextCompare) Then
        Set fld = tbl.Columns(i)
        Exit For
      End If
    Next

    If Not (fld Is Nothing) Then
      If (fld.Type = adVarWChar) Then
        For i = 0 To tbl.Columns.Count - 1
          If 0 = StrComp(tbl.Columns(i).name, "TempInfo", vbTextCompare) Then
            Set tempfld = tbl.Columns(i)
            Exit For
          End If
        Next
        If tempfld Is Nothing Then
          tbl.Columns.Append "TempInfo", adVarWChar
          tbl.Columns.Item("TempInfo").Properties("Jet OLEDB:Allow Zero Length") = True
          tbl.Columns.Item("TempInfo").Properties("Jet OLEDB:Required") = False
          tbl.Columns.Item("TempInfo").Properties("Required") = False
          tbl.Columns.Item("TempInfo").Properties("Nullable") = True
        End If

        SQL = "UPDATE Residents SET TempInfo  = '' "  ' clear out temp column data if any residual
        ConnExecute SQL

        SQL = "UPDATE Residents SET TempInfo  = Info "  ' copy to temp column
        ConnExecute SQL


        SQL = "ALTER TABLE Residents DROP COLUMN Info "  ' drop old Info column
        ConnExecute SQL

        tbl.Columns.Append "Info", adLongVarWChar     ' recreate it
        tbl.Columns.Item("Info").Properties("Jet OLEDB:Allow Zero Length") = True
        tbl.Columns.Item("Info").Properties("Jet OLEDB:Required") = False
        tbl.Columns.Item("Info").Properties("Required") = False
        tbl.Columns.Item("Info").Properties("Nullable") = True

        SQL = "UPDATE Residents SET Info  = TempInfo "  ' copy data back
        ConnExecute SQL

        SQL = "UPDATE Residents SET Info  = '' WHERE TempInfo is null"  ' eliminate any nulls
        ConnExecute SQL

        SQL = "ALTER TABLE Residents DROP COLUMN TempInfo "  ' get rid of temp column
        ConnExecute SQL




      End If  ' fld.Type = adChar
    End If  '(fld Is Nothing)


  End If
  Set tempfld = Nothing
  Set fld = Nothing
  Set tbl = Nothing
  Set cat = Nothing


End Function

Function ExpandStaffInfo() As Boolean


  Dim tbl           As ADOX.Table
  Dim cat           As ADOX.Catalog
  Dim fld           As ADOX.Column
  Dim i             As Integer
  Dim SQL           As String
  Dim tempfld       As ADOX.Column

  Dim TableName     As String
  Dim Fieldname     As String

  TableName = "Staff"
  Fieldname = "Info"


  Connect

  Set cat = New ADOX.Catalog
  cat.ActiveConnection = conn
  On Error Resume Next
  Set tbl = cat.Tables(TableName)
  If Not (tbl Is Nothing) Then

    For i = 0 To tbl.Columns.Count - 1
      If 0 = StrComp(tbl.Columns(i).name, Fieldname, vbTextCompare) Then
        Set fld = tbl.Columns(i)
        Exit For
      End If
    Next

    If Not (fld Is Nothing) Then
      If (fld.Type = adVarWChar) Then
        For i = 0 To tbl.Columns.Count - 1
          If 0 = StrComp(tbl.Columns(i).name, "TempInfo", vbTextCompare) Then
            Set tempfld = tbl.Columns(i)
            Exit For
          End If
        Next
        If tempfld Is Nothing Then
          tbl.Columns.Append "TempInfo", adVarWChar
          tbl.Columns.Item("TempInfo").Properties("Jet OLEDB:Allow Zero Length") = True
          tbl.Columns.Item("TempInfo").Properties("Jet OLEDB:Required") = False
          tbl.Columns.Item("TempInfo").Properties("Required") = False
          tbl.Columns.Item("TempInfo").Properties("Nullable") = True
        End If

        SQL = "UPDATE " & TableName & " SET TempInfo  = '' "  ' clear out temp column data if any residual
        ConnExecute SQL

        SQL = "UPDATE " & TableName & " SET TempInfo  = " & Fieldname  ' copy to temp column
        ConnExecute SQL


        SQL = "ALTER TABLE " & TableName & " DROP COLUMN " & Fieldname  ' drop old Info column
        ConnExecute SQL

        tbl.Columns.Append "Info", adLongVarWChar     ' recreate it
        tbl.Columns.Item("Info").Properties("Jet OLEDB:Allow Zero Length") = True
        tbl.Columns.Item("Info").Properties("Jet OLEDB:Required") = False
        tbl.Columns.Item("Info").Properties("Required") = False
        tbl.Columns.Item("Info").Properties("Nullable") = True

        SQL = "UPDATE " & TableName & " SET " & Fieldname & "  = TempInfo "  ' copy data back
        ConnExecute SQL

        SQL = "UPDATE " & TableName & " SET " & Fieldname & "  = '' WHERE TempInfo is null"  ' eliminate any nulls
        ConnExecute SQL

        SQL = "ALTER TABLE " & TableName & " DROP COLUMN TempInfo "  ' get rid of temp column
        ConnExecute SQL




      End If  ' fld.Type = adChar
    End If  '(fld Is Nothing)


  End If
  Set tempfld = Nothing
  Set fld = Nothing
  Set tbl = Nothing
  Set cat = Nothing


End Function


Function UpdateCLSPTIs()
  Dim SQL           As String
  Dim DeviceType    As ESDeviceTypeType
  Dim j             As Integer
  Debug.Print "DeviceTypes"

  If gIsJET Then
    For j = LBound(ESDeviceType) To UBound(ESDeviceType)
      DeviceType = ESDeviceType(j)
      SQL = "UPDATE Devicetypes set MIDPTI = " & DeviceType.CLSPTI & " WHERE Model = '" & DeviceType.Model & "'"
      ConnExecute SQL
      Debug.Print HexFormat(DeviceType.CLSPTI, 4) & " " & DeviceType.Model
    Next
  End If
End Function


Function copytophone()
  Dim rs            As Recordset
  Dim Count         As Long

  Set rs = ConnExecute("SELECT count(phone) FROM Residents WHERE phone > '' ")
  Count = rs(0)
  rs.Close
  If Count = 0 Then
    ConnExecute "UPDATE Residents SET Phone = Name"
  End If


End Function
Function SQLEnsureFieldExists(ByVal TableName As String, ByVal Fieldname As String, ByVal FieldType As String, ByVal FieldLen As String, ByVal DefaultValue As String) As Boolean

  Dim SQL           As String

  SQL = " IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'" & TableName & "') AND type in (N'U'))"
  SQL = SQL & " BEGIN "
  If FieldType = adVarWChar Then
    SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " NVARCHAR(" & FieldLen & ") NULL ) ON [PRIMARY] "
  ElseIf FieldType = adLongVarWChar Then
    SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " NVARCHAR(MAX) NULL ) ON [PRIMARY] "
  ElseIf FieldType = adLongVarBinary Then  ' Image
    SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " IMAGE NULL ) ON [PRIMARY] "
  ElseIf FieldType = adDate Then  ' DDL = DATATIME
    SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " DATETIME NULL ) ON [PRIMARY] "
  ElseIf FieldType = adDouble Then  ' DDL = INT, LONG
    SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " FLOAT NULL DEFAULT 0) ON [PRIMARY] "
  ElseIf FieldType = adInteger Then  ' DDL = INT, LONG
    If Val(FieldLen) < 0 Then  ' IDENTITY
      SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " INT IDENTITY NOT NULL ) ON [PRIMARY] "
    Else
      SQL = SQL & " CREATE TABLE " & TableName & "(" & Fieldname & " INT NULL DEFAULT 0) ON [PRIMARY] "
    End If
  End If
  SQL = SQL & " END "

  ConnExecute SQL

  If FieldType = adVarWChar Then  '  DDL = NVARCHAR(n)
    SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
    SQL = SQL & " BEGIN "
    SQL = SQL & "ALTER TABLE " & TableName & " ADD " & Fieldname & " NVARCHAR(" & FieldLen & ") NULL"
    SQL = SQL & " END "
    ConnExecute SQL

    SQL = "UPDATE " & TableName & " SET " & Fieldname & " = '' WHERE " & Fieldname & " is null"
    ConnExecute SQL

  ElseIf FieldType = adLongVarWChar Then  ' (memo) -  DDL = NVARCHAR(MAX)

    SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
    SQL = SQL & " BEGIN "
    SQL = SQL & "ALTER TABLE " & TableName & " ADD " & Fieldname & " NVARCHAR(MAX) NULL"
    SQL = SQL & " END "
    ConnExecute SQL

    SQL = "UPDATE " & TableName & " SET " & Fieldname & " = '' WHERE " & Fieldname & " is null"
    ConnExecute SQL

  ElseIf FieldType = adLongVarBinary Then  ' DDL = Image

    SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
    SQL = SQL & " BEGIN "
    SQL = SQL & "ALTER TABLE " & TableName & " ADD " & Fieldname & " IMAGE NULL"
    SQL = SQL & " END "
    ConnExecute SQL


  ElseIf FieldType = adDate Then  ' DDL = DATATIME

    SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
    SQL = SQL & " BEGIN "
    SQL = SQL & "ALTER TABLE " & TableName & " ADD " & Fieldname & " DATETIME NULL"
    SQL = SQL & " END "
    ConnExecute SQL

  ElseIf FieldType = adDouble Then  ' DDL = Double
    SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
    SQL = SQL & " BEGIN "
    SQL = SQL & " ALTER TABLE " & TableName & " ADD " & Fieldname & " FLOAT NULL DEFAULT 0"
    SQL = SQL & " END "
    conn.Execute SQL
    SQL = "UPDATE " & TableName & " SET " & Fieldname & " = 0 WHERE " & Fieldname & " is null"
    ConnExecute SQL

  ElseIf FieldType = adInteger Then  ' DDL = INT, LONG
    If Val(FieldLen) < 0 Then  ' IDENTITY

      SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
      SQL = SQL & " BEGIN "
      SQL = SQL & " ALTER TABLE " & TableName & " ADD " & Fieldname & " INT IDENTITY NOT NULL "
      SQL = SQL & " END "
      conn.Execute SQL


    Else                      ' Normal int

      SQL = "IF NOT EXISTS(select * from sys.columns where Name = N'" & Fieldname & "' and Object_ID = Object_ID(N'" & TableName & "'))"
      SQL = SQL & " BEGIN "
      SQL = SQL & " ALTER TABLE " & TableName & " ADD " & Fieldname & " INT NULL DEFAULT 0"
      SQL = SQL & " END "
      conn.Execute SQL
      SQL = "UPDATE " & TableName & " SET " & Fieldname & " = 0 WHERE " & Fieldname & " is null"
      ConnExecute SQL

    End If
  End If


End Function


Function EnsureFieldExists(ByVal TableName As String, ByVal Fieldname As String, ByVal FieldType As String, ByVal FieldLen As String, ByVal DefaultValue As String) As Boolean

  If Not gIsJET Then
    EnsureFieldExists = SQLEnsureFieldExists(TableName, Fieldname, FieldType, FieldLen, DefaultValue)
    Exit Function
  End If



  Dim tbl           As ADOX.Table
  Dim cat           As ADOX.Catalog
  Dim fld           As ADOX.Column
  Dim i             As Integer
  Dim SQL           As String

  Connect

  Set cat = New ADOX.Catalog
  cat.ActiveConnection = conn
  On Error Resume Next
  Set tbl = cat.Tables(TableName)

  If tbl Is Nothing Then
    Set tbl = New ADOX.Table
    tbl.name = TableName
    cat.Tables.Append tbl
  End If


  For i = 0 To tbl.Columns.Count - 1
    If 0 = StrComp(tbl.Columns(i).name, Fieldname, vbTextCompare) Then
      Set fld = tbl.Columns(i)
      Exit For
    End If
  Next



  ''?? Attributes = adColNullable  remove the required attribute ??

  If fld Is Nothing Then
    If FieldType = adVarWChar Then


      tbl.Columns.Append Fieldname, FieldType, FieldLen
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Allow Zero Length") = True
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Required") = False
      tbl.Columns.Item(Fieldname).Properties("Required") = False
      tbl.Columns.Item(Fieldname).Properties("Nullable") = True

      SQL = "UPDATE " & TableName & " SET " & Fieldname & " = '' WHERE " & Fieldname & " is null"
      ConnExecute SQL

    ElseIf FieldType = adLongVarWChar Then
      tbl.Columns.Append Fieldname, FieldType
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Allow Zero Length") = True
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Required") = False
      tbl.Columns.Item(Fieldname).Properties("Required") = False
      tbl.Columns.Item(Fieldname).Properties("Nullable") = True
      SQL = "UPDATE " & TableName & " SET " & Fieldname & " = '' WHERE " & Fieldname & " is null"
      ConnExecute SQL

    ElseIf FieldType = adLongVarBinary Then
      tbl.Columns.Append Fieldname, FieldType
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Allow Zero Length") = True
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Required") = False
      tbl.Columns.Item(Fieldname).Properties("Required") = False
      tbl.Columns.Item(Fieldname).Properties("Nullable") = True

    ElseIf FieldType = adDate Then
      tbl.Columns.Append Fieldname, FieldType
      tbl.Columns.Item(Fieldname).Properties("Jet OLEDB:Required") = False
      tbl.Columns.Item(Fieldname).Properties("Required") = False
      tbl.Columns.Item(Fieldname).Properties("Nullable") = True

      'Sql = "UPDATE " & Tablename & " SET " & FieldName & " = 0 WHERE " & FieldName & " is null"
      'connexecute Sql

    ElseIf FieldType = adDouble Then
      tbl.Columns.Append Fieldname, FieldType
      SQL = "UPDATE " & TableName & " SET " & Fieldname & " = 0 WHERE " & Fieldname & " is null"
      ConnExecute SQL

    ElseIf FieldType = adInteger Then
      If Val(FieldLen) < 0 Then
        SQL = "ALTER TABLE " & TableName & " ADD COLUMN " & Fieldname & " COUNTER"
        ConnExecute SQL
      Else
        tbl.Columns.Append Fieldname, FieldType
        SQL = "UPDATE " & TableName & " SET " & Fieldname & " = 0 WHERE " & Fieldname & " is null"
        ConnExecute SQL
      End If
    End If


    '    tbl.Columns.Item("Announce").Properties("Nullable") = True

  End If
  Set fld = Nothing
  Set tbl = Nothing
  Set cat = Nothing

End Function

Function CalcESChecksum(ByVal s As String)
  Dim Checksum      As Integer
  Dim hexsum        As String
  Dim Msglen        As Integer
  'Generates checksum from hex string


  Dim j             As Integer

  Dim HexByte       As String
  'Debug.Print ""
  Msglen = Val("&h" & MID(s, 3, 2))
  s = left(s, Msglen * 2)
  hexsum = MID(s, Msglen * 2 + 1, 2)

  's = "7212B20E087E00116BCD00003E1800104F4F" '17" <- checksum
  For j = 1 To Msglen * 2 - 1 Step 2

    HexByte = "&h" & MID(s, j, 2)
    'Debug.Print hexbyte
    Checksum = (Checksum + Val(HexByte)) And &HFF&

  Next
  'Debug.Print "Data          Hex (decimal)"
  'Debug.Print "Message Length " & Right("00" & Hex(MsgLen), 2) & " (" & MsgLen & ")"
  'Debug.Print "Checksum       " & Right("00" & Hex(checksum), 2) & " (" & checksum & ")"
  'Debug.Print ""

  CalcESChecksum = Hex(Checksum)

End Function


Public Function ValidateLogin(ByVal login As String) As Boolean
  Dim User          As cUser

  If 0 = StrComp(login, FACTORY_PWD, vbTextCompare) Then
    gUser.UserID = 0
    gUser.LEvel = LEVEL_FACTORY
    gUser.Username = "Factory"
    gUser.Password = login
    gUser.ConsoleID = ConsoleID
    ValidateLogin = True
  Else
    Set User = GetUser(login)
    If User.LEvel > 0 Then
      Set gUser = User
      ValidateLogin = True
    End If
  End If
End Function


Public Function PCARedirect(a As cAlarm, ByVal SurveyPCA As String) As Boolean
  Dim Location      As String
  Location = a.locationtext
  Outbounds.AddMessage SurveyPCA, &H100, a.locationtext, 0
End Function


Public Function SaveDevice(Device As cESDevice, Username As String) As Long  ' returns deviceid
        Dim rs            As Recordset
        Dim SQL           As String
        Dim adding        As Boolean
        Dim Serial        As String


10      Debug.Print "modmain.SaveDevice device.ZoneID " & Device.ZoneID


20      dbgGeneral "Modmain.SaveDevice Entry " & vbCrLf

30      On Error GoTo SaveDevice_Error

40      Set rs = New ADODB.Recordset

50      SQL = "SELECT * FROM Devices WHERE serial = " & q(Device.Serial)

        'conn.BeginTrans
60      rs.Open SQL, conn, gCursorType, gLockType

70      If rs.EOF Then
80        dbgGeneral "Modmain.SaveDevice rs.eof" & vbCrLf
90        rs.Close
100       adding = True
110       SQL = "Devices"
120       rs.Open SQL, conn, gCursorType, gLockType, adCmdTable
130       rs.addnew
140     End If
150     dbgGeneral "Modmain.SaveDevice rs(...)=" & vbCrLf
160     Serial = Device.Serial
170     rs("Serial") = Device.Serial
180     rs("model") = Device.Model

190     rs("Deleted") = 0
200     rs("Assurinput") = Device.AssurInput
210     rs("IsPortable") = Device.IsPortable    ' = DeviceType.Portable
220     rs("NumInputs") = Device.NumInputs    ' =DeviceType.NumInputs

        ' Pairs
230     rs("Announce") = Device.Announce
240     rs("Announce_A") = Device.Announce_A
250     rs("Announce_B") = Device.Announce_B

260     rs("UseAssur") = Device.UseAssur
270     rs("UseAssur2") = Device.UseAssur2

280     rs("UseAssur_A") = Device.UseAssur_A
290     rs("UseAssur2_A") = Device.UseAssur2_A

300     rs("UseAssur_B") = Device.UseAssur_B
310     rs("UseAssur2_B") = Device.UseAssur2_B
 
     

320     rs("ClearByReset") = 0    ' Device.ClearByReset

        'rs("ClearByReset_A") = Device.ClearByReset_A

330     rs("custom") = Device.Custom
340     rs("ignoretamper").Value = Device.IgnoreTamper



350     rs("OG1") = Device.OG1
360     rs("OG2") = Device.OG2
370     rs("OG3") = Device.OG3
380     rs("OG4") = Device.OG4
390     rs("OG5") = Device.OG5
400     rs("OG6") = Device.OG6


410     rs("NG1") = Device.NG1
420     rs("NG2") = Device.NG2
430     rs("NG3") = Device.NG3
440     rs("NG4") = Device.NG4
450     rs("NG5") = Device.NG5
460     rs("NG6") = Device.NG6

470     rs("GG1") = Device.GG1
480     rs("GG2") = Device.GG2
490     rs("GG3") = Device.GG3
500     rs("GG4") = Device.GG4
510     rs("GG5") = Device.GG5
520     rs("GG6") = Device.GG6

530     rs("OG1_A") = Device.OG1_A
540     rs("OG2_A") = Device.OG2_A
550     rs("OG3_A") = Device.OG3_A
560     rs("OG4_A") = Device.OG4_A
570     rs("OG5_A") = Device.OG5_A
580     rs("OG6_A") = Device.OG6_A

590     rs("NG1_A") = Device.NG1_A
600     rs("NG2_A") = Device.NG2_A
610     rs("NG3_A") = Device.NG3_A
620     rs("NG4_A") = Device.NG4_A
630     rs("NG5_A") = Device.NG5_A
640     rs("NG6_A") = Device.NG6_A

650     rs("GG1_A") = Device.GG1_A
660     rs("GG2_A") = Device.GG2_A
670     rs("GG3_A") = Device.GG3_A
680     rs("GG4_A") = Device.GG4_A
690     rs("GG5_A") = Device.GG5_A
700     rs("GG6_A") = Device.GG6_A



710     rs("OG1_b") = Device.OG1_B
720     rs("OG2_b") = Device.OG2_B
730     rs("OG3_b") = Device.OG3_B
740     rs("OG4_b") = Device.OG4_B
750     rs("OG5_b") = Device.OG5_B
760     rs("OG6_b") = Device.OG6_B

770     rs("NG1_b") = Device.NG1_B
780     rs("NG2_b") = Device.NG2_B
790     rs("NG3_b") = Device.NG3_B
800     rs("NG4_b") = Device.NG4_B
810     rs("NG5_b") = Device.NG5_B
820     rs("NG6_b") = Device.NG6_B

830     rs("GG1_b") = Device.GG1_B
840     rs("GG2_b") = Device.GG2_B
850     rs("GG3_b") = Device.GG3_B
860     rs("GG4_b") = Device.GG4_B
870     rs("GG5_b") = Device.GG5_B
880     rs("GG6_b") = Device.GG6_B




890     rs("OG1d") = Device.OG1D
900     rs("OG2d") = Device.OG2D
910     rs("OG3d") = Device.OG3D
920     rs("OG4d") = Device.OG4D
930     rs("OG5d") = Device.OG5D
940     rs("OG6d") = Device.OG6D


950     rs("NG1d") = Device.NG1D
960     rs("NG2d") = Device.NG2D
970     rs("NG3d") = Device.NG3D
980     rs("NG4d") = Device.NG4D
990     rs("NG5d") = Device.NG5D
1000    rs("NG6d") = Device.NG6D

1010    rs("GG1d") = Device.GG1D
1020    rs("GG2d") = Device.GG2D
1030    rs("GG3d") = Device.GG3D
1040    rs("GG4d") = Device.GG4D
1050    rs("GG5d") = Device.GG5D
1060    rs("GG6d") = Device.GG6D


1070    rs("OG1_Ad") = Device.OG1_AD
1080    rs("OG2_Ad") = Device.OG2_AD
1090    rs("OG3_Ad") = Device.OG3_AD
1100    rs("OG4_Ad") = Device.OG4_AD
1110    rs("OG5_Ad") = Device.OG5_AD
1120    rs("OG6_Ad") = Device.OG6_AD

1130    rs("NG1_Ad") = Device.NG1_AD
1140    rs("NG2_Ad") = Device.NG2_AD
1150    rs("NG3_Ad") = Device.NG3_AD
1160    rs("NG4_Ad") = Device.NG4_AD
1170    rs("NG5_Ad") = Device.NG5_AD
1180    rs("NG6_Ad") = Device.NG6_AD

1190    rs("GG1_Ad") = Device.GG1_AD
1200    rs("GG2_Ad") = Device.GG2_AD
1210    rs("GG3_Ad") = Device.GG3_AD
1220    rs("GG4_Ad") = Device.GG4_AD
1230    rs("GG5_Ad") = Device.GG5_AD
1240    rs("GG6_Ad") = Device.GG6_AD

1250    rs("OG1_bd") = Device.OG1_BD
1260    rs("OG2_bd") = Device.OG2_BD
1270    rs("OG3_bd") = Device.OG3_BD
1280    rs("OG4_bd") = Device.OG4_BD
1290    rs("OG5_bd") = Device.OG5_BD
1300    rs("OG6_bd") = Device.OG6_BD

1310    rs("NG1_bd") = Device.NG1_BD
1320    rs("NG2_bd") = Device.NG2_BD
1330    rs("NG3_bd") = Device.NG3_BD
1340    rs("NG4_bd") = Device.NG4_BD
1350    rs("NG5_bd") = Device.NG5_BD
1360    rs("NG6_bd") = Device.NG6_BD

1370    rs("GG1_bd") = Device.GG1_BD
1380    rs("GG2_bd") = Device.GG2_BD
1390    rs("GG3_bd") = Device.GG3_BD
1400    rs("GG4_bd") = Device.GG4_BD
1410    rs("GG5_bd") = Device.GG5_BD
1420    rs("GG6_bd") = Device.GG6_BD




        'Stop
        ' save for 3rd shift data



1430    rs("VacationSuper") = Device.AssurSecure
1440    rs("VacationSuper_A") = Device.AssurSecure_A
1441    rs("VacationSuper_B") = Device.AssurSecure_B
1450    rs("SendCancel") = Device.SendCancel
1460    rs("SendCancel_A") = Device.SendCancel_A
1470    rs("SendCancel_B") = Device.SendCancel_B

1480    rs("DisableStart") = Device.DisableStart
1490    rs("DisableEnd") = Device.DisableEnd

1500    rs("DisableStart_A") = Device.DisableStart_A
1510    rs("DisableEnd_A") = Device.DisableEnd_A

1520    rs("DisableStart_B") = Device.DisableStart_B
1530    rs("DisableEnd_B") = Device.DisableEnd_B


1540    rs("repeatuntil") = Device.RepeatUntil
1550    rs("repeatuntil_A") = Device.RepeatUntil_A
1560    rs("repeatuntil_B") = Device.RepeatUntil_B

1570    rs("repeats") = Device.Repeats
1580    rs("repeats_A") = Device.Repeats_A
1590    rs("repeats_B") = Device.Repeats_B

1600    rs("Pause") = Device.Pause
1610    rs("Pause_A") = Device.Pause_A
1620    rs("Pause_B") = Device.Pause_B

1630    rs("AlarmMask") = Device.AlarmMask
1640    rs("AlarmMask_A") = Device.AlarmMask_A
1650    rs("AlarmMask_B") = Device.AlarmMask_B

1660    rs("ResidentID") = Device.ResidentID
1670    rs("ResidentID_A") = 0   ' Device.ResidentID_A

1680    rs("RoomID") = Device.RoomID
1690    rs("RoomID_A") = 0   'Device.RoomID_A

1700    rs("ClearByReset") = Device.ClearByReset

        ' bogus
1710    rs("RoomD_A") = 0   'Device.RoomID_A
1720    rs("ClearByReset_A") = 0   'Device.ClearByReset_a

1730    rs("lowset") = Device.LowSet
1740    rs("lowset_a") = Device.LowSet_A
1750    rs("hiset") = Device.HiSet
1760    rs("hiset_a") = Device.HiSet_A
1770    rs("EnableTemp") = Device.EnableTemperature
1780    rs("EnableTemp_a") = Device.EnableTemperature_A

1790    rs("UseTamperAsInput") = Device.UseTamperAsInput

1795    rs("LocKW") = ""


        rs("ConfigurationString") = Device.Configurationstring
        
1800    If (Device.Model <> COM_DEV_NAME) Or adding Then
1810      rs("SerialTapProtocol") = 0
1820      rs("SerialSkip") = 0
1830      rs("SerialMessageLen") = 0
1840      rs("SerialAutoClear") = 0
1850      rs("SerialPort") = 0
1860      rs("SerialBaud") = 0
1870      rs("SerialBits") = 0
1880      rs("SerialParity") = ""
1890      rs("SerialStopBits") = ""
1900      rs("SerialInclude") = ""
1910      rs("SerialExclude") = ""
1920      rs("SerialFlow") = 0
1930      rs("SerialEOLChar") = 0
1940      rs("SerialPreamble") = ""
1950    End If
1960    rs("SerialSettings") = rs("SerialBaud") & rs("SerialParity") & rs("SerialBits") & rs("SerialStopBits")


        ' NOW FOR 6080
        
1965   rs("LocKW") = ""
        
1970    rs("IDM") = Device.ZoneID
1980    rs("IDL") = Device.IDL

1990    rs("ignored") = Device.Ignored


2000    dbgGeneral "Modmain.SaveDevice rs.update" & vbCrLf
2010    rs.Update
        'conn.CommitTrans
2020    Device.DeviceID = rs("DeviceID")
2030    dbgGeneral "Modmain.SaveDevice rs.update/DeviceID" & Device.DeviceID & vbCrLf
2040    rs.Close
2050    Set rs = Nothing

2060    If adding Then   ' log adding the device
2070      Dim rslog       As ADODB.Recordset: Set rslog = New ADODB.Recordset

2080      SQL = "alarms"
          'conn.BeginTrans

2090      rslog.Open SQL, conn, gCursorType, gLockType, adCmdTable
2100      rslog.addnew

2110      rslog("FC1") = 0
2120      rslog("FC2") = 0
2130      rslog("IDM") = Device.IDM
2140      rslog("IDL") = Device.IDL
2150      rslog("Status") = 0
2160      rslog("Serial") = Device.Serial
2170      rslog("EventDate") = Now
2180      rslog("Alarm") = 0
2190      rslog("Tamper") = 0
2200      rslog("IsLocator") = 0
2210      rslog("Battery") = 0
2220      rslog("LOCIDM") = 0
2230      rslog("LOCIDL") = 0
2240      rslog("ResidentID") = Device.ResidentID
2250      rslog("RoomID") = Device.RoomID
2260      rslog("EventType") = EVT_ADD_DEV
2270      rslog("AlarmID") = 0   ' not an alarm
2280      rslog("username") = Username
2290      rslog("sessionid") = gSessionID
2300      rslog("announce") = Device.Announce
2310      rslog("Phone") = ""
2320      rslog("Info") = ""
2330      rslog.Update
          'conn.CommitTrans
2340      rslog.Close
2350    End If

2360    SaveDevice = Device.DeviceID

2370    dbgGeneral "Modmain.SaveDevice Exit" & vbCrLf

2380    If adding Then
2390      Dim NewDevice   As cESDevice: Set NewDevice = New cESDevice
2400      NewDevice.Serial = Device.Serial
2410      Devices.AddDevice NewDevice
2420      Set Device = NewDevice
2430    End If

2440    Devices.RefreshBySerial Device.Serial   ' make sure all params are set

2450    SetupSerialDevice Device   ' if it's a serial device, set it up too
2460    If USE6080 = 0 Then
2470      Select Case left$(UCase$(Device.Model), 6)
            Case "EN5000", "EN5040", "EN5081"
2480          Outbounds.AddMessage Device.Serial, MSGTYPE_REPEATERNID, "", 0
              ' create outbound message to set NID
2490        Case "EN3954"
2500          Outbounds.AddMessage Device.Serial, MSGTYPE_TWOWAYNID, "", 0
              ' create outbound message to set NID
2510      End Select
2520    End If

2530    If Device.Ignored Then
2540      InBounds.ClearAllAlarmsBySerial Serial
2550      alarms.ClearAllAlarmsBySerial Serial
2560      Alerts.ClearAllAlarmsBySerial Serial
2570      Troubles.ClearAllAlarmsBySerial Serial
2580      LowBatts.ClearAllAlarmsBySerial Serial

2590    End If


SaveDevice_Resume:
2600    On Error GoTo 0
2610    Exit Function

SaveDevice_Error:
        Dim rc
        rc = messagebox(frmMain, "Error " & Err.Number & " (" & Err.Description & ") at modMain.SaveDevice." & Erl & vbCrLf, "Save Error", vbCritical Or vbOKOnly)
2620    dbg "Error " & Err.Number & " (" & Err.Description & ") at modMain.SaveDevice." & Erl & vbCrLf
        'LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modMain.SaveDevice." & Erl
2630    Resume SaveDevice_Resume


End Function

Public Sub DestroyObjects()

  ' cleans house at shutdown

  Dim j             As Integer
  Dim d             As Object

  Dim si            As cSerialInput
  Dim Dev           As cESDevice

  On Error Resume Next
  For j = Devices.Devices.Count To 1 Step -1
    Set Dev = Devices.Devices(j)
    Devices.Devices.Remove j
    Set Dev = Nothing
  Next
  For j = SerialIns.Count To 1 Step -1
    Set si = SerialIns(j)
    SerialIns.Remove j
    Set si = Nothing
  Next


  Set Devices = Nothing
  For j = gPageDevices.Count To 1 Step -1
    Set d = gPageDevices(j)
    gPageDevices.Remove j
    Set d = Nothing

  Next


End Sub

Public Function GetLastCheckInLocation(ByVal ResidentID As Long, Optional Serial As String = "") As String
  ' get a/the portable device for a given resident



End Function

Public Function Get6080Info()
  Dim HTTPRequest   As cHTTPRequest
  Dim rc            As Long

  Set ZoneInfoList = New cZoneInfoList
  Set HTTPRequest = New cHTTPRequest
  Call HTTPRequest.GetZoneList(GetHTTP & "://" & IP1, USER1, PW1)
  Do Until HTTPRequest.Ready
    DoEvents
  Loop
  Select Case HTTPRequest.StatusCode
    Case 200, 201
    Case Else
  End Select
  If Len(HTTPRequest.XML) Then
    rc = ZoneInfoList.LoadXML(HTTPRequest.XML)
  End If
  Set HTTPRequest = Nothing

End Function


Public Function CleanFileName(ByVal filename As String) As String
  ' BAD_FILE_CHARS
  Dim j             As Long
  'Global Const BAD_FILE_CHARS = "\/:*?<>|" & """"
  For j = 1 To Len(BAD_FILE_CHARS)
    filename = Replace(filename, MID$(BAD_FILE_CHARS, j, 1), " ")
  Next
 
  CleanFileName = filename
 
End Function


Public Sub BootLog(ByVal s As String)
  Dim hfile As Long
  If App.Revision > 813 Then Exit Sub
  Dim filename As String
  filename = App.Path & "\AppBoot.log"
  limitFileSize filename
  
  hfile = FreeFile
  Open filename For Append As hfile
  Print #hfile, s
  Close #hfile
End Sub

Public Sub ResetLockTime()
  gLockTimeRemaining = LOCK_DELAY
End Sub
Public Sub ResetActivityTime()
  If gUser.LEvel >= LEVEL_FACTORY And gExtendFactory Then
    gInactivityTimeRemaining = INACTIVITY_DELAY_FACTORY
  Else
    gInactivityTimeRemaining = INACTIVITY_DELAY
  End If
End Sub



Public Sub ClearAllMobiles()
  Dim SQL As String
  SQL = "DELETE FROM Mobile"
  ConnExecute SQL

End Sub


Function ToUnixTime(ByVal DateAndTime As String) As String
  ToUnixTime = Format$(DateAndTime, "yyyy-mm-dd") & "T" & Format$(DateAndTime, "hh:nn:ss") & "Z"
End Function

Function FromUnixTime(ByVal DateAndTime As String) As String
  DateAndTime = Replace(DateAndTime, "T", " ", Compare:=vbTextCompare)
  FromUnixTime = Replace(DateAndTime, "Z", " ", Compare:=vbTextCompare)
End Function


Function FixISODateTime(ByVal DateAndTime As String) As String
  Dim DateParts()        As String
  DateAndTime = Trim$(FromUnixTime(DateAndTime))
  DateParts = Split(DateAndTime, ".")
  FixISODateTime = DateParts(0)
End Function

Function FormatAsISO(ByVal DateAndTime As Date, Optional ByVal IncludeTSeperator As Boolean = False) As String
  Dim TSeperator         As String
  TSeperator = IIf(IncludeTSeperator, "T", " ")
  FormatAsISO = Format(DateAndTime, "YYYY-MM-DD" & TSeperator & "HH:NN:SS")
End Function
