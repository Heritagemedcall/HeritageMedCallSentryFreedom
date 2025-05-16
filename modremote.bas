Attribute VB_Name = "modRemote"
Option Explicit


'This is the fake XML-RPC on the Remote(s)
Global nosessioncount As Long
Global WaitingForHost As Boolean
Global HostQue        As New Collection

Global RemoteAutoEnrollEnabled As Boolean
Global RemoteQuitting As Boolean
Global RemoteConnectStatus As Long

Global HostTime As Double ' the time on the host computer
Global HostOffset As Long ' seconds
'Global SyncHostTime As Long ' flag to sync time or not (in config settings)


Private mRemoteConnectError As Long

Global QuedEvents As Collection


Public Function ClientRequestAssist(ByVal panel As String, ByVal Serial As String, ByVal Inputnum As Long, ByVal AlarmID As Long) As Boolean

  Dim XML As String
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("ClientRequestAssist")
  XML = XML & taggit("action", "requestassist") & vbCrLf
  XML = XML & taggit("panel", panel) & vbCrLf
  XML = XML & taggit("alarmid", AlarmID) & vbCrLf
  XML = XML & taggit("myalarms", gMyAlarms) & vbCrLf
  XML = XML & taggit("serial", Serial) & vbCrLf

  XML = XML & taggit("inputnum", CStr(Inputnum)) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ClientRequestAssist = ParseandPostAlarms(XML)




End Function



Public Function GetQuedEvents() As String
  Dim XML                As String
  Dim QuedEvent          As cQuedEvent
  If Not QuedEvents Is Nothing Then
    XML = XML & "<Events>"
    For Each QuedEvent In QuedEvents
      XML = XML & "<Event>" & QuedEvent.ToXML & "</Event>" & vbCrLf
    Next
    XML = XML & "</Events>" & vbCrLf
    GetQuedEvents = XML
    Set QuedEvents = New Collection
  End If

End Function


Public Function QueEvent(ByVal EventName As String, ByVal ConsoleID As String, ByVal Alarmtype As String)
    Dim QuedEvent As cQuedEvent
    Dim ExistingEvent As cQuedEvent
    Dim j As Long
    
    If QuedEvents Is Nothing Then
      Set QuedEvents = New Collection
    End If
  
    EventName = LCase$(EventName)
    
    Select Case EventName
      Case "clientunsilencealarms"
        Set QuedEvent = New cQuedEvent
        QuedEvent.EventName = EventName
        QuedEvent.ConsoleID = ConsoleID
        QuedEvent.Alarmtype = Alarmtype
        
        For j = QuedEvents.Count To 1 Step -1
          Set ExistingEvent = QuedEvents(j)
          If ExistingEvent.EventName = QuedEvent.EventName And QuedEvent.EventName = QuedEvent.Alarmtype Then
            Exit For
          End If
        Next
        If j = 0 Then
          QuedEvents.Add QuedEvent
          Debug.Print "QueEvent.Unsilence Added"
        End If
        
      Case "clientsilencealarms"
        Set QuedEvent = New cQuedEvent
        QuedEvent.EventName = EventName
        QuedEvent.ConsoleID = ConsoleID
        QuedEvent.Alarmtype = Alarmtype
        For j = QuedEvents.Count To 1 Step -1
          Set ExistingEvent = QuedEvents(j)
          If ExistingEvent.EventName = QuedEvent.EventName And QuedEvent.EventName = QuedEvent.Alarmtype Then
            Exit For
          End If
        Next
        
        If j = 0 Then
          QuedEvents.Add QuedEvent
          Debug.Print "QueEvent.Silence  Added"

          QuedEvents.Add QuedEvent
        
        End If
    
    End Select
  
End Function
Public Sub ClearQuedEvents()
    If QuedEvents Is Nothing Then
      Set QuedEvents = New Collection
    End If

End Sub

Public Function DeQueEvent() As cQuedEvent
  If Not QuedEvents Is Nothing Then
     Set DeQueEvent = QuedEvents(1)
     QuedEvents.Remove 1
  End If



End Function


'Private Type SYSTEMTIME
'        wYear As Integer
'        wMonth As Integer
'        wDayOfWeek As Integer
'        wDay As Integer
'        wHour As Integer
'        wMinute As Integer
'        wSecond As Integer
'        wMilliseconds As Integer
'End Type



'private declare function GetSystemTimeAsFileTime (tFiletime as FILETIME)
'Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Sub SyncTime()
  Dim HostTimeAsDate As Date
  
  If HostTime > 42370 Then ' 1/1/2016
    On Error Resume Next ' in case user does not have priveledges
    HostTimeAsDate = CDate(HostTime)
    Date = HostTimeAsDate
    Time = HostTimeAsDate
  End If


End Sub


Function RemoteLogout(User As cUser)
  Dim XML As String
  If Not HostInterraction Is Nothing Then
    If HostInterraction.Socket.State = sckConnected Then
      XML = "<?xml version=""1.0""?>" & vbCrLf
      XML = XML & MakeXMLAttributes("remotelogout")
      XML = XML & taggit("action", "logout") & vbCrLf
      XML = XML & "</HMC>"
      
      HostInterraction.Send "post", XML
      XML = GetHostInterractionResponse()
      LogRemoteSession 0, 0, "Remote RemoteLogout"
      
    Else
      Set gUser = New cUser
      dbg "RemoteLogout, no connection"
      LogRemoteSession 0, 0, "Remote RemoteLogout, no connection"
    End If
  End If

End Function




'Sub ValidateMyLogon(Node As IXMLDOMNode)
' ' this is called by modRemote.ParseandPostAlarms on every alarm update packet
' ' if sessionid is zero then host has logged us out
'
'' node is "logon"
'' node has children
'' currentadmin (NIC ID)
'' force logoff
'' if i'm an admin, and the current admin isn't me then
'' log me out
'' else leave me logged in
'' unless all logout forced
'
'
'  Dim childnode   As IXMLDOMNode
'  Dim HostAdmin   As String
'  Dim SessionID   As Long
'
''  For Each childnode In Node.childNodes
''
''    Select Case childnode.baseName
''      Case "session"
''        SessionID = Val(childnode.text)
''        Select Case gUser.Level
''          Case LEVEL_FACTORY
''            ' Factory does not get logged off by host
''
''          Case LEVEL_ADMIN, LEVEL_SUPERVISOR
''            If SessionID = 0 Then
''              If LoggedIn Then
''                PostEvent Nothing, Nothing, Nothing, EVT_FORCED_LOGOUT, 0
''                frmMain.DoLogin ' do login sets loggedin to False
''              End If
''            End If
''          Case Else ' LEVEL_USER
''            'Allow continued interraction
''            ' we won't log out a non-admin USER
''
''        End Select
''    End Select
''  Next
'
'End Sub

Public Function RemoteIsAssurActive() As Boolean
  Dim XML           As String
  Dim Root          As IXMLDOMNode
  Dim Node          As IXMLDOMNode
  Dim doc           As DOMDocument60: Set doc = New DOMDocument60

  If HostInterraction Is Nothing Then Exit Function

  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("isassuractive")
  XML = XML & taggit("action", "isassuractive") & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    ' no change
  Else
    Set Node = Root.nextSibling
    If Not Node Is Nothing Then
      If LCase(Node.baseName) = "value" Then
        RemoteIsAssurActive = IIf(Node.text = "True", True, False)
      End If
    End If
    
    For Each Node In Root.childnodes
      Select Case LCase(Node.baseName)
      
      Case "isassuractive"
        'RemoteIsAssurActive = IIf(Val(Node.text) = 1, True, False)
      End Select
      'RemoteIsAssurActive = IIf(Val(Node.text) = "isassuractive", True, False)
      Exit For
    Next
  End If
  
  Set Root = Nothing
  Set doc = Nothing



  End Function

Public Function RemoteDeleteResident(User As cUser, ByVal ResidentID As Long)
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "deleteresident") & vbCrLf
  XML = XML & taggit("residentid", ResidentID) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    'dbg "RemoteDeleteResident Failure" & vbCrLf
    ' did not fly
    RemoteDeleteResident = -1 ' no response or could be error
  Else
    'dbg "RemoteDeleteResident OK" & vbCrLf
    RemoteDeleteResident = 0 ' no errors
  End If

  Set Root = Nothing
  Set doc = Nothing


End Function

Public Function RemoteGetUser(ByVal Password As String) As cUser

  Dim XML As String
 
  Dim delay As Date
  
  Dim TempUser As cUser
  
  If HostInterraction Is Nothing Then Exit Function
  
  Set TempUser = New cUser
  
  
  delay = DateAdd("s", 5, Now)
  

  
  Do While delay > Now
    If HostInterraction.IsConnected Then
      Exit Do
    End If
    dbg "RemoteGetUser Waiting for connection"
    DoEvents
  Loop
  
  


  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("remotegetuser")
  
  XML = XML & taggit("action", "logon") & vbCrLf
  XML = XML & taggit("password", Password) & vbCrLf
  
  XML = XML & "</HMC>"
  
  'dbg "Send " & xml
  RemoteConnectStatus = 0
  HostInterraction.Send "post", XML
  dbg "RemoteGetUser POST " & vbCrLf & XML
  XML = GetHostInterractionResponse()
  dbg "RemoteGetUser RESP" & vbCrLf & XML
  Set TempUser = GetUserFromXML(XML)
  Set RemoteGetUser = TempUser

End Function
Public Function GetUserFromXML(ByVal XML As String) As cUser
  
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  
  Dim doc As DOMDocument60:   Set doc = New DOMDocument60
  Dim User As cUser:        Set User = New cUser
  
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/logon")
  If Root Is Nothing Then
    ' just the blank user
  Else
    
    For Each Node In Root.childnodes
      Select Case LCase(Node.baseName)
        Case "loggedon"
          User.LoggedOn = IIf(Val(Node.text) = 1, True, False)
        Case "level"
          User.LEvel = Val(Node.text)
        Case "userid"
          User.UserID = Val(Node.text)
        Case "password"
          User.Password = Node.text
        Case "username"
          User.Username = Node.text
        Case "name"
          User.Username = Node.text
        Case "session"
          User.Session = Val(Node.text)
        Case "userpermissions"
          User.UserPermissions.ParseUserPermissions Val(Node.text)
      End Select
    Next
  End If
  
  Set GetUserFromXML = User
  
  Set Node = Nothing
  Set Root = Nothing
  Set doc = Nothing


End Function



Public Function MakeXMLAttributes(Optional ByVal Caller As String = "unk") As String
  Dim XML As String
  
  XML = "<HMC revision=" & q(App.Revision) & " ConsoleID=" & q(ConsoleID) & " RemoteSerial=" & q(Configuration.RemoteSerial) & " User=" & q(gUser.Username) & " session=" & q(gUser.Session) & ">" & vbCrLf
  
'  Debug.Assert Not (Caller = "unk")
  
  'Debug.Print "!!!! MakeXMLAttributes "; Caller
  'Debug.Print XML
  
    
  
  MakeXMLAttributes = XML
End Function


Public Function RemoteClearAssurs() As Long
  ' this will tell the Master to clear out assur failures, and to reset devices for the next assur period
  
  'dbg "ModRemote.RemoteClearAssurs"
  
   
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "clearassurs") & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()


  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Failure")
  If Root Is Nothing Then
    Set Root = doc.selectSingleNode("HMC/Success")
      If Root Is Nothing Then
        
       ' dbg "ModRemote.RemoteClearAssurs Root = nothing "
      '  RemoteClearAssurs = -1 ' no response or could be error
      Else
        If Len(XML) Then
          ParseandPostAlarms XML
          'dbg "ModRemote.RemoteClearAssurs Call to ParseandPostAlarms "
        Else
          ' no return!
        End If
        
      '  RemoteClearAssurs = 0 ' no errors
      End If
   Else
    'dbg "ModRemote.RemoteClearAssurs Failure"
  End If
  
  Set Root = Nothing
  Set doc = Nothing




End Function


Public Function RemoteSetResidentAwayStatus(ByVal Away As Integer, ByVal ResID As Long) As Long
  
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "setresidentaway") & vbCrLf
  XML = XML & taggit("resid", CStr(ResID)) & vbCrLf
  XML = XML & taggit("away", CStr(Away)) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    RemoteSetResidentAwayStatus = -1 ' no response or could be error
  Else
    ' could get away status
    RemoteSetResidentAwayStatus = 0 ' no errors
  End If

  Set Root = Nothing
  Set doc = Nothing


End Function

Public Function RemoteSetRoomAwayStatus(ByVal Away As Integer, ByVal RoomID As Long) As Long
  
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "setroomaway") & vbCrLf
  XML = XML & taggit("roomid", CStr(RoomID)) & vbCrLf
  XML = XML & taggit("away", CStr(Away)) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    RemoteSetRoomAwayStatus = -1 ' no response or could be error
  Else
    ' could get away status
    RemoteSetRoomAwayStatus = 0 ' no errors
  End If

  Set Root = Nothing
  Set doc = Nothing


End Function



Public Function RemoteDeleteTransmitter(ByVal DeviceID As Long) As Long

  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "deletedevice") & vbCrLf
  XML = XML & taggit("deviceid", DeviceID) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    'dbg "RemoteDeleteTransmitter Failure" & vbCrLf
    ' did not fly
    RemoteDeleteTransmitter = -1 ' no response or could be error
  Else
    'dbg "RemoteDeleteTransmitter OK" & vbCrLf
    RemoteDeleteTransmitter = 0 ' no errors
  End If

  Set Root = Nothing
  Set doc = Nothing

End Function

Public Function RemoteUpdateDevice(Device As cESDevice) As Long
  
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "savedevice") & vbCrLf
  XML = XML & "<device>" & vbCrLf
  XML = XML & Device.ToXML  ' ********* THIS IS WHERE DEVICE IS CONVERTED TO XML
  XML = XML & "</device>" & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  'dbg "RemoteUpdateDevice" & vbCrLf & xml & vbCrLf
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    'dbg "RemoteUpdateDevice Failure" & vbCrLf
    ' did not fly
    RemoteUpdateDevice = -1 ' no response or could be error
  Else
    'dbg "RemoteUpdateDevice OK" & vbCrLf
    RemoteUpdateDevice = 0 ' no errors
  End If
  
  Set Root = Nothing
  Set doc = Nothing


End Function


Public Function RemoteSaveTemperatureDevice(Device As cESDevice) As Long
  
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "savetemperaturedevice") & vbCrLf
  XML = XML & "<device>" & vbCrLf
  XML = XML & Device.SerialToXML
  XML = XML & "</device>" & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    'dbg " Failure" & vbCrLf
    ' did not fly
    RemoteSaveTemperatureDevice = -1 ' no response or could be error
  Else
    'dbg " OK" & vbCrLf
    RemoteSaveTemperatureDevice = 0 ' no errors
  End If

  Set Root = Nothing
  Set doc = Nothing


End Function

Public Function RemoteSaveSerialDevice(Device As cESDevice) As Long
  
  Dim XML As String
  Dim Root As IXMLDOMNode
  Dim doc As DOMDocument60: Set doc = New DOMDocument60

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "saveserialdevice") & vbCrLf
  XML = XML & "<device>" & vbCrLf
  XML = XML & Device.SerialToXML
  XML = XML & "</device>" & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    'dbg " Failure" & vbCrLf
    ' did not fly
    RemoteSaveSerialDevice = -1 ' no response or could be error
  Else
    'dbg " OK" & vbCrLf
    RemoteSaveSerialDevice = 0 ' no errors
  End If

  Set Root = Nothing
  Set doc = Nothing


End Function


Public Function RemoteStartAutoEnroll() As Long
  ' see: modClients.Client_AutoEnroll
  ' enters thru: modClients.ProcessClientRequest
  Dim XML As String
  Dim Root As IXMLDOMNode

  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  
  'dbg "RemoteStartAutoEnroll" & vbCrLf
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "autoenroll") & vbCrLf
  XML = XML & taggit("SubFunction", "start") & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  
  doc.LoadXML XML
  Set Root = doc.selectSingleNode("HMC/Success")
  If Root Is Nothing Then
    ' did not fly
    RemoteStartAutoEnroll = -5 ' no response or could be error
  Else
    RemoteStartAutoEnroll = 0 ' no errors
  End If
  
  Set Root = Nothing
  Set doc = Nothing


End Function
Public Function RemotePollAutoEnroll() As Long

' see: modClients.Client_AutoEnroll
' enters thru: modClients.ProcessClientRequest

' Polls Host to get status of auto enroll...
' a 0 return is good
' anything else is error


  Dim XML         As String
  Dim ErrorNumber As Long
  Dim Serial      As String
  Dim MIDPTI      As Long
  Dim CLSPTI      As Long
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode

  Dim Model As String


  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "autoenroll") & vbCrLf
  XML = XML & taggit("SubFunction", "check") & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  'dbg "RemotePollAutoEnroll " & vbCrLf
  XML = GetHostInterractionResponse()

  
  If Len(XML) > 0 Then
    doc.LoadXML XML
    

    Set Root = doc.selectSingleNode("HMC/Waiting")  ' most frequent response to our request
    
    If (Root Is Nothing) Then  'no longer waiting, maybe we're done

      Set Root = doc.selectSingleNode("HMC/Success") ' hopefully!
      If Not (Root Is Nothing) Then  ' we have success, process it
        Set Node = doc.selectSingleNode("HMC/Serial")
        If Not Node Is Nothing Then
          Serial = Node.text
        Else
          ErrorNumber = -1
        End If
        Set Node = doc.selectSingleNode("HMC/CLSPTI")
        If Not Node Is Nothing Then
          CLSPTI = Val(Node.text)
          ' was midpti
        Else
          ErrorNumber = -2
        End If
        If ErrorNumber = 0 Then
          'push Serial And MIDPTI to frmtransmitter
          Model = GetDeviceModelByCLSPTI(CLSPTI) 'MIDPTI)
          'dbg "Remote AutoEnroll " & Serial & " " & model & vbCrLf
          On Error Resume Next
          frmTransmitter.txtSerial.text = Serial
          frmTransmitter.cboDeviceType.ListIndex = CboFindExact(frmTransmitter.cboDeviceType, Model)
          frmTransmitter.DisableAutoEnroll
        End If
  
      Else  ' failure
        Set Root = doc.selectSingleNode("HMC/Fail")
        If Not (Root Is Nothing) Then  ' failure
          'dbg "Remote AutoEnroll Failure" & vbCrLf
          
          ErrorNumber = -3  ' may expand later
        Else
          Set Root = doc.selectSingleNode("HMC/TimeOut")
          If Not (Root Is Nothing) Then  ' timeout
            'dbg "Remote AutoEnroll Server TimeOut" & vbCrLf
            ' cancel it
            ErrorNumber = -9
          Else  ' no response, general failure, no connection
            'dbg "Remote AutoEnroll Invalid Response" & vbCrLf
            ErrorNumber = -4
            ' cancel it
          End If
        End If
      End If
    Else
      'dbg "Remote AutoEnrollWaiting" & vbCrLf
    End If
  Else  ' no xml
    ' cancel it
    ErrorNumber = -5
  End If
  RemotePollAutoEnroll = ErrorNumber
  
  Set Node = Nothing
  Set Root = Nothing
  Set doc = Nothing

End Function
Public Function RemoteCancelAutoEnroll() As Long

  ' see: modClients.Client_AutoEnroll
  ' enters thru: modClients.ProcessClientRequest
  Dim XML As String
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "autoenroll") & vbCrLf
  XML = XML & taggit("SubFunction", "cancel") & vbCrLf
  XML = XML & "</HMC>"
  
  On Error Resume Next
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  
  'dbg xml & vbCrLf
  
  ' doesn't really matter what we get back
  ' we're cancelling here on the remote
  ' and if the Host got the message, it will cancel there too
  RemoteCancelAutoEnroll = 0


End Function

Public Function ClientGetAlarms() As Boolean

  Dim t As Date
  Dim t1 As Date
  Dim t2 As Date
  Dim t3 As Date
  
  Dim User As cUser
  
  Dim XML As String
  
  If gRegistered = False Then
    frmMain.ShowDisconnect 2
    Exit Function
  End If
  
  If gUser.LEvel = LEVEL_FACTORY Then ' only do this for remote Factory
    If gUser.Session = 0 Then
      Set User = RemoteGetUser(gUser.Password)
      If User.Session <> 0 And User.LEvel = LEVEL_FACTORY Then
        Set gUser = User
      End If
    End If
  End If

  

  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("getalarms")
  XML = XML & taggit("action", "getalarms") & vbCrLf
  XML = XML & GetQuedEvents()
  
  
  XML = XML & "</HMC>"
  
  Debug.Print "ClientGetAlarms"
  t = Now
  
  HostConnection.Send "post", XML
  
  t1 = DateDiff("s", Now, t)
  t = Now
  
  Dim Response As String
  
  Response = GetHostResponse()
  
  t2 = DateDiff("s", Now, t)
  t = Now
  
  ClientGetAlarms = ParseandPostAlarms(Response)
  
  t3 = DateDiff("s", Now, t)
  'Debug.Print "Post " & t1 & vbCrLf & "Response " & t2 & vbCrLf & "Parse and Post " & t3
    
  
End Function

Public Function ClientGetSubscribedAlarms() As Boolean
        Dim t                  As Long
        Dim t1                 As Long
        Dim t2                 As Long
        Dim t3                 As Long
        Dim User               As cUser
        Dim Response           As String

        Dim lastget            As Date
        Static thisget         As Date


        Dim XML                As String

10      If gRegistered = False Then
          frmMain.ShowDisconnect 2
20        Exit Function
30      End If


40      On Error GoTo ClientGetSubscribedAlarms_Error

50      If gUser.LEvel = LEVEL_FACTORY Then  ' only do this for remote Factory
60        If gUser.Session = 0 Then
70          Set User = RemoteGetUser(gUser.Password)
80          If User.Session <> 0 And User.LEvel = LEVEL_FACTORY Then
90            Set gUser = User
100         End If
110       End If
120     End If


130     Debug.Print "***** gUser.Session  "; gUser.Session



140     XML = "<?xml version=""1.0""?>" & vbCrLf
150     XML = XML & MakeXMLAttributes("getsubscribedalarms")
160     XML = XML & taggit("action", "getsubscribedalarms") & vbCrLf
170     XML = XML & GetQuedEvents()
180     XML = XML & "</HMC>"



190     Debug.Print "ClientGetSubscribedAlarms"
200     Debug.Print XML
210     Debug.Print


220     t = Win32.timeGetTime
230     HostConnection.Send "post", XML

240     t1 = Win32.timeGetTime() - t
250     t = Win32.timeGetTime()
260     Response = GetHostResponse()


270     t2 = Win32.timeGetTime() - t


280     t = Win32.timeGetTime()

290     Debug.Print "ClientGetSubscribedAlarms - Response"
300     Debug.Print Response
        'Debug.Print Now
310     Debug.Print

320     If CDbl(thisget) = 0 Then
330       thisget = Now
340     Else
350       Debug.Print "Last Get "; DateDiff("s", thisget, Now)
360       thisget = Now
370     End If
380     t = Win32.timeGetTime()

390     ClientGetSubscribedAlarms = ParseandPostAlarms(Response)

400     t3 = Win32.timeGetTime() - t

410     Debug.Print "Post " & t1 & vbCrLf & "Response " & t2 & vbCrLf & "Parse and Post " & t3

ClientGetSubscribedAlarms_Resume:

420     On Error GoTo 0
430     Exit Function

ClientGetSubscribedAlarms_Error:

440     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modRemote.ClientGetSubscribedAlarms." & Erl
450     Resume ClientGetSubscribedAlarms_Resume

End Function

Public Function ClientSilenceAlarms(ByVal ConsoleID As String, ByVal RemoteSerial As String, ByVal Alarmtype As String) As Boolean
  Dim XML As String

  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("silencealarms")
  XML = XML & taggit("action", "silencealarms") & vbCrLf
  XML = XML & taggit("consoleid", XMLEncode(ConsoleID)) & vbCrLf
  XML = XML & taggit("alarmtype", XMLEncode(Alarmtype)) & vbCrLf
  XML = XML & taggit("remoteserial", XMLEncode(RemoteSerial)) & vbCrLf
  XML = XML & taggit("myalarms", gMyAlarms) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ClientSilenceAlarms = ParseandPostAlarms(XML)

End Function

Public Function ClientUnSilenceAlarms(ByVal ConsoleID As String, ByVal RemoteSerial As String, ByVal Alarmtype As String) As Boolean
  Dim XML As String

  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("unsilencealarms")
  XML = XML & taggit("action", "unsilencealarms") & vbCrLf
  XML = XML & taggit("consoleid", XMLEncode(ConsoleID)) & vbCrLf
  XML = XML & taggit("remoteserial", XMLEncode(RemoteSerial)) & vbCrLf
  XML = XML & taggit("alarmtype", XMLEncode(Alarmtype)) & vbCrLf
  XML = XML & taggit("myalarms", gMyAlarms) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ' just discard results

End Function



' these are send functions! Client->Server
'Public Function ClientGreeting() As Boolean
'
'
'  'Simple Connection Greeting to HOST
'  Dim XML As String
'  Dim HostRequest As cHostRequest: Set HostRequest = New cHostRequest
'
'  XML = "<?xml version=""1.0""?>" & vbCrLf
'  XML = XML & MakeXMLAttributes
'  XML = XML & taggit("action", "Greeting") & vbCrLf
'  XML = XML & taggit("value", "Hello") & vbCrLf
'
'  XML = XML & "</HMC>"
'
'  HostInterraction.Send "post", XML
'  ClientGreeting = ParseandPostAlarms(GetHostInterractionResponse())
'
'
'End Function


Public Function ClientUpdateDeviceResidentID(ByVal ResidentID As Long, ByVal DeviceID As Long) As Boolean
  
  ' Request to  have HOST change ResidentID associated with Device
  Dim XML As String

  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "UpdateDeviceResidentID") & vbCrLf
  XML = XML & taggit("ResidentID", ResidentID) & vbCrLf
  XML = XML & taggit("DeviceID", DeviceID) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  ClientUpdateDeviceResidentID = ParseandPostAlarms(GetHostInterractionResponse())
  
  
End Function

Public Function ClientUpdateDeviceRoomID(ByVal RoomID As Long, ByVal DeviceID As Long) As Boolean
  
  ' Request to  have HOST change RoomID associated with Device
  Dim XML As String
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Node As IXMLDOMNode
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "UpdateDeviceRoomID") & vbCrLf
  XML = XML & taggit("roomid", CStr(RoomID)) & vbCrLf
  XML = XML & taggit("deviceid", CStr(DeviceID)) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  
  ' need success or failue, not parse and post!
  XML = GetHostInterractionResponse()
  

  doc.LoadXML XML

  Set Node = doc.selectSingleNode("HMC/Success")
  If Not (Node Is Nothing) Then
    ClientUpdateDeviceRoomID = True
  Else
    ClientUpdateDeviceRoomID = False
  End If
  Set Node = Nothing
  Set doc = Nothing


End Function

Public Function ClientNotify(ByVal Category As String) As Boolean
  'notify of a direct DataBase change by client, only needs an ACK
  Dim XML As String
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "ClientNotify") & vbCrLf
  XML = XML & taggit("category", Category) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  ClientNotify = ParseandPostAlarms(GetHostInterractionResponse())
  

End Function

Public Function ClientSendToPager(ByVal message As String, ByVal PagerID As Long, ByVal NoWait As Integer, ByVal Phone As String, ByVal Inputnum As Long) As Boolean
  ' sends message to pager, only needs ack
  Dim XML As String
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "SendToPager") & vbCrLf
  XML = XML & taggit("message", XMLEncode(message)) & vbCrLf
  XML = XML & taggit("pagerid", PagerID) & vbCrLf
  XML = XML & taggit("nowait", NoWait) & vbCrLf
  XML = XML & taggit("phone", XMLEncode(Phone)) & vbCrLf
  XML = XML & taggit("inputnum", 0) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ClientSendToPager = True

End Function

Public Function ClientSendToGroup(ByVal message As String, ByVal GroupID As Long, ByVal Phone As String, ByVal Inputnum As Long) As Boolean
  Dim XML As String
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "SendToGroup") & vbCrLf
  XML = XML & taggit("message", message) & vbCrLf
  XML = XML & taggit("groupid", GroupID) & vbCrLf
  XML = XML & taggit("phone", Phone) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ClientSendToGroup = True


End Function

Public Function Remote_GetRoom(Room As cRoom, RoomID As Long) As Boolean
  Dim XML As String
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Node As IXMLDOMNode
  Dim Root As IXMLDOMNode

  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "getroom") & vbCrLf
  XML = XML & taggit("roomid", CStr(RoomID)) & vbCrLf
  XML = XML & "</HMC>"

  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()


  doc.LoadXML XML

  Set Root = doc.selectSingleNode("HMC/getroom")
  If Not (Root Is Nothing) Then
    For Each Node In Root.childnodes
      Select Case LCase(Node.baseName)
        Case "roomid"
          Room.RoomID = Val(Node.text)
        Case "assurdays"
          Room.Assurdays = Val(Node.text)
        Case "vacation"
          Room.Away = Val(Node.text)
        Case "room"
          Room.Room = Node.text
        Case "flags"
          Room.flags = Val(Node.text)
        Case "lockw"
          Room.locKW = Node.text
      End Select
    Next
    Remote_GetRoom = True
  Else
    Remote_GetRoom = False
  End If

  Set Root = Nothing
  Set Node = Nothing
  Set doc = Nothing



End Function

Public Function ClientUpdateRoom(Room As cRoom) As Boolean
  Dim XML As String
  
  
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Node As IXMLDOMNode
  
  'Dim HostRequest As cHostRequest: Set HostRequest = New cHostRequest

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "saveroom") & vbCrLf
  XML = XML & "<room>"
  '..... build update here
  XML = XML & taggit("assurdays", CStr(Room.Assurdays))
  XML = XML & taggit("away", CStr(Room.Away))
  XML = XML & taggit("building", XMLEncode(Room.Building))
  XML = XML & taggit("deleted", CStr(Room.Deleted))
  XML = XML & taggit("description", XMLEncode(Room.Description))
  XML = XML & taggit("room", XMLEncode(Room.Room))
  XML = XML & taggit("lockw", XMLEncode(Trim$(Room.locKW)))
  XML = XML & taggit("roomid", CStr(Room.RoomID))
  XML = XML & taggit("flags", CStr(Room.flags))
  XML = XML & taggit("vacation", "0")
  XML = XML & "</room>"
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  

  doc.LoadXML XML

  Set Node = doc.selectSingleNode("HMC/Success")
  If Not (Node Is Nothing) Then
    Set Node = doc.selectSingleNode("HMC/Value")
    If Not (Node Is Nothing) Then
      Room.RoomID = Val(Node.text)
      ClientUpdateRoom = True
    Else
      ClientUpdateRoom = False
    End If
  End If

  Set Node = Nothing
  Set doc = Nothing


End Function
Public Function ClientUpdateResident(Resident As cResident) As Boolean
  Dim XML As String
  
  
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Node As IXMLDOMNode
  
  'Dim HostRequest As cHostRequest: Set HostRequest = New cHostRequest

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "UpdateResident") & vbCrLf
  
  '..... build update here
  XML = XML & taggit("residentID", CStr(Resident.ResidentID))
  XML = XML & taggit("namelast", XMLEncode(Resident.NameLast))
  XML = XML & taggit("nameFirst", XMLEncode(Resident.NameFirst))
  XML = XML & taggit("phone", XMLEncode(Resident.Phone))
  XML = XML & taggit("Room", XMLEncode(Resident.Room))
  XML = XML & taggit("Info", XMLEncode(Resident.info))
  XML = XML & taggit("AssurDays", CStr(Resident.Assurdays And &HFF))
  XML = XML & taggit("Vacation", "0")
  XML = XML & taggit("Away", CStr(Resident.Vacation))
  XML = XML & taggit("DeliveryPoints", Resident.DeliveryPointsToString())
  XML = XML & taggit("Deleted", "0")
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()

  doc.LoadXML XML

  Set Node = doc.selectSingleNode("HMC/Success")
  If Not (Node Is Nothing) Then
    Set Node = doc.selectSingleNode("HMC/Value")
    If Not (Node Is Nothing) Then
      Resident.ResidentID = Val(Node.text)
      ClientUpdateResident = True
    Else
      ClientUpdateResident = False
    End If
  End If
'
  
  Set Node = Nothing
  Set doc = Nothing

End Function

Public Function ClientUpdateStaff(Resident As cResident) As Boolean
  Dim XML As String
  
  
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Node As IXMLDOMNode
  
  'Dim HostRequest As cHostRequest: Set HostRequest = New cHostRequest

  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes
  XML = XML & taggit("action", "UpdateStaff") & vbCrLf
  
  '..... build update here
  XML = XML & taggit("residentID", CStr(Resident.ResidentID))
  XML = XML & taggit("namelast", XMLEncode(Resident.NameLast))
  XML = XML & taggit("nameFirst", XMLEncode(Resident.NameFirst))
  XML = XML & taggit("phone", XMLEncode(Resident.Phone))
  XML = XML & taggit("Room", XMLEncode(Resident.Room))
  XML = XML & taggit("Info", XMLEncode(Resident.info))
  XML = XML & taggit("AssurDays", CStr(Resident.Assurdays And &HFF))
  XML = XML & taggit("Vacation", "0")
  XML = XML & taggit("Away", CStr(Resident.Vacation))
  XML = XML & taggit("Deleted", "0")
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()

  doc.LoadXML XML

  Set Node = doc.selectSingleNode("HMC/Success")
  If Not (Node Is Nothing) Then
    Set Node = doc.selectSingleNode("HMC/Value")
    If Not (Node Is Nothing) Then
      Resident.ResidentID = Val(Node.text)
      ClientUpdateStaff = True
    Else
      ClientUpdateStaff = False
    End If
  End If
'
  
  Set Node = Nothing
  Set doc = Nothing

End Function

Public Function ClientFinalizeAlarm(ByVal panel As String, ByVal Serial As String, ByVal Inputnum As Long, ByVal AlarmID As Long, ByVal Disposition As String) As Boolean
  Dim XML As String
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("ClientFinalizeAlarm")
  XML = XML & taggit("action", "finalizealarm") & vbCrLf
  XML = XML & taggit("disposition", XMLEncode(Disposition)) & vbCrLf
  XML = XML & taggit("panel", panel) & vbCrLf
  XML = XML & taggit("alarmid", AlarmID) & vbCrLf
  XML = XML & taggit("serial", Serial) & vbCrLf
  XML = XML & taggit("myalarms", gMyAlarms) & vbCrLf
  XML = XML & taggit("inputnum", CStr(Inputnum)) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ClientFinalizeAlarm = ParseandPostAlarms(XML)
  'dbg "ClientACKAlarm"

End Function



Public Function ClientACKAlarm(ByVal panel As String, ByVal Serial As String, ByVal Inputnum As Long, ByVal AlarmID As Long) As Boolean
  Dim XML As String
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & MakeXMLAttributes("ClientACKAlarm")
  XML = XML & taggit("action", "ackalarm") & vbCrLf
  XML = XML & taggit("panel", panel) & vbCrLf
  XML = XML & taggit("alarmid", AlarmID) & vbCrLf
  XML = XML & taggit("serial", Serial) & vbCrLf
  XML = XML & taggit("myalarms", gMyAlarms) & vbCrLf
  XML = XML & taggit("inputnum", CStr(Inputnum)) & vbCrLf
  XML = XML & "</HMC>"
  
  HostInterraction.Send "post", XML
  XML = GetHostInterractionResponse()
  ClientACKAlarm = ParseandPostAlarms(XML)
  'dbg "ClientACKAlarm"

End Function



Public Sub CheckRemoteConnectStatus()
'
' this is where we see if we've lost communications with the host
' called every second by master clock

  'Debug.Print "RemoteConnectStatus " & RemoteConnectStatus & " " & Now
  
  'dbg "Checking Host Connection Counter " & RemoteConnectStatus
  If RemoteConnectStatus <= 100 Then  ' 2 minute timeout
    RemoteConnectStatus = RemoteConnectStatus + 1
  Else
    If LoggedIn Then
      If gUser.LEvel <> LEVEL_FACTORY Then  ' log off all except Factory due to timeout
        If RemoteConnectStatus >= 100 Then
          dbg "logging off due to timeout (Check Remote Connect Status)"
          frmMain.DoLogin
        End If
      End If
    Else
      dbg "Reponse overdue (Check Remote Connect Status)"
    End If
  End If


End Sub


Function ParseandPostAlarms(ByVal XML As String) As Boolean
  

  
  Dim doc As DOMDocument60: Set doc = New DOMDocument60
  Dim Root            As IXMLDOMNode  ' main node of interest
  Dim Node            As IXMLDOMNode
  Dim FunctionName    As String
  Dim SessionID       As Long

  'dbg "ParseandPostAlarms"

  

  doc.LoadXML XML
  'Set Root = doc.selectSingleNode("HMC/Response")  ' mixed case

  'If Root Is Nothing Then
  Set Root = doc.selectSingleNode("HMC/response")  'lower case
  'End If

  If Not Root Is Nothing Then

    RemoteConnectStatus = 0
    FunctionName = LCase$(Root.text)
    Select Case FunctionName
      Case "getalarms"
        frmMain.RefreshAlarms XML
        dbg "ParseandPostAlarms " & XML
        ' also handle forced logoff message
        Set Node = doc.selectSingleNode("HMC/session")
        If Node Is Nothing Then
          dbg "ParseandPostAlarms No Session Node"
          LogRemoteSession nosessioncount, -1, "ParseandPostAlarms No Session Node"
        End If
        
        If LoggedIn Then
          If Not Node Is Nothing Then
            SessionID = Val(Node.text)
            If SessionID = 0 Then
              If gUser.LEvel <> LEVEL_FACTORY Then
                If nosessioncount > 2 Then
                  dbg "Logging out due to no sessionID (Parse and Post Alarms)"
                  LogRemoteSession nosessioncount, SessionID, "nosessioncount, Logging out due to no sessionID (Parse and Post Alarms)"
                  frmMain.DoLogin
                  'nosessioncount = 0
                  nosessioncount = Min(99, nosessioncount + 3)
                End If
              End If
            Else ' has session ID
                nosessioncount = Min(99, nosessioncount + 3)
            End If
          End If

        End If
        frmMain.PacketToggle
        'dbg "Session " & SessionID
      Case "ackalarm"
        frmMain.RefreshAlarms XML
        frmMain.PacketToggle
        dbg "ParseandPostAlarms ACKAlarm " & left(XML, 100)
      Case "error'"
    End Select
  Else
    dbg "ParseandPostAlarms No Data xml '" & left(XML, 150) & "'"
  End If

  Set Root = Nothing
  Set doc = Nothing


End Function
Function ParseResidentUpdateResponse(doc As DOMDocument60, Resident As Object) As Boolean

  Dim Node As IXMLDOMNode
  If Not Resident Is Nothing Then
    Set Node = doc.selectSingleNode("HMC/Success")
    If Not (Node Is Nothing) Then
      Set Node = doc.selectSingleNode("HMC/Value")
      If Not (Node Is Nothing) Then
        Resident.ResidentID = Val(Node.text)
        ParseResidentUpdateResponse = True
      Else
        ParseResidentUpdateResponse = False
      End If
    End If
  End If
End Function



Function GetHostResponse() As String
  Dim s                  As String
  'Dim t As Long

  On Error GoTo GetHostResponse_Error

  Dim timedelay          As Date

  't = Win32.timeGetTime + 10000 '2 second timout now 10, now 20

  timedelay = DateAdd("s", 20, Now)

  Do Until HostConnection.ResponseReady
    DoEvents
    If RemoteQuitting Then
      Exit Do
    End If
    If Now > timedelay Then
      Exit Do
    End If
  Loop
  dbg "HostConnection.ResponseReady"
  s = HostConnection.ResponseText
  If HostConnection.GetWinsockError <> 0 Or Len(s) = 0 Then
    RemoteConnectError = 1
    'frmMain.imgPacket.Picture = LoadResPicture(1005, vbResBitmap)  ' red
  Else
    RemoteConnectError = 0
  End If


  GetHostResponse = s
  'Debug.Print "response"; s
GetHostResponse_Resume:
  On Error GoTo 0
  Exit Function

GetHostResponse_Error:

  'LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modRemote.GetHostResponse." & Erl
  Resume GetHostResponse_Resume


End Function


Function GetHostInterractionResponse() As String
        'same as GetHostResponse but gets it via a different channel
        Dim s As String
        
        Dim i As Long
        
        Dim j As Currency
        Dim freq As Currency
        
        Dim SysTime As SYSTEMTIME
        'Win32.GetLocalTime SysTime
        
        Dim tstart As Double
        
        Dim fractional As Double
        
        
        
        'fractional = (SysTime.wHour * 60 * 60 * 60) + (SysTime.wMinute * 60 * 60) + (SysTime.wSecond * 60)
        'start = CLng(Now)
        'Dim TheTime As FILETIME
        'Win32.GetSystemTimeAsFileTime TheTime
        
        

10      On Error GoTo GetHostInterractionResponse_Error
        
        If Win32.timeGetTime + 200 >= 2 ^ 32 Then
          Sleep 200
        End If
        
20      i = Win32.timeGetTime + 50  ' 50 ms
        
        tstart = DateAdd("s", 5, Now)

30      Do Until HostInterraction.ResponseReady
          
          If Win32.timeGetTime > i Then ' yield every (100) ms
40          DoEvents
            If Win32.timeGetTime + 200 >= 2 ^ 32 Then
              Sleep 200
            End If
            i = Win32.timeGetTime + 100
            
          End If
50        If RemoteQuitting Then
            dbg "GetHostInterractionResponse RemoteQuitting"
60          Exit Do
            
70        End If
80        If Now > tstart Then
            dbg "GetHostInterractionResponse Timeout"
90          Exit Do
100       End If
110     Loop
120     s = HostInterraction.ResponseText
        If Len(s) = 0 Then
        '  Debug.Assert 0
        End If
130     If HostInterraction.GetWinsockError <> 0 Or Len(s) = 0 Then
          
          RemoteConnectError = 1
140       '  done elsewhere frmMain.imgPacket.Picture = LoadResPicture(1005, vbResBitmap)  ' red
        Else
            RemoteConnectError = 0
150     End If

160     GetHostInterractionResponse = s

GetHostInterractionResponse_Resume:
170      On Error GoTo 0
180      Exit Function

GetHostInterractionResponse_Error:

        'LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modRemote.GetHostInterractionResponse." & Erl
190     Resume GetHostInterractionResponse_Resume


End Function



Public Property Get RemoteConnectError() As Long

  RemoteConnectError = mRemoteConnectError

End Property

Public Property Let RemoteConnectError(ByVal Value As Long)

  If Value Then
    frmMain.imgPacket.Picture = LoadResPicture(1005, vbResBitmap)  ' red
    frmMain.ShowDisconnect 1 ' show
  Else
    frmMain.ShowDisconnect 0 ' hide
  End If

  mRemoteConnectError = Value

End Property
