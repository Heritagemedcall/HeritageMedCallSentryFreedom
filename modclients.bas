Attribute VB_Name = "modClients"
Option Explicit

'This is the XML-RPC on the host


Global Const MAX_REMOTE_TROUBLES = 50


'Global AdminConsoleID As String
'Global AdminForceOff  As Boolean
Global HostSessions        As Collection
'Global HostSessionNum      As Long



' These are the client connections
Global Listener            As Winsock  ' this is the one that receives the connect request and ands off to newly created winsocks

Global ClientConnections   As Collection

' this is the remote's winsock
Global Client              As Winsock
' used for remote autoenroll
Global RemoteAutoEnroller  As cRemoteAutoEnroller


' this is for client requests for alarm status every 5 seconds
Global HostConnection     As cHostConnection
  
' this is for client requests for anything else
Global HostInterraction   As cHostConnection

Function AlarmFromXML(ByVal XML As String)




End Function

Function GetNextSession() As Long
  Dim SessionNum As Long
  Static NextSession As Long
  NextSession = NextSession + 1

  If NextSession > 999999999 Then
    NextSession = 1
  End If

  GetNextSession = NextSession

  
End Function

' need to handle multiple request from multiple clients


Function Client_DoLogon(doc As DOMDocument60) As String

        Dim Root      As IXMLDOMNode
        Dim Node      As IXMLDOMNode
        Dim attr      As IXMLDOMAttribute
        Dim Password  As String
        Dim User      As cUser
        Dim XML       As String
        Dim LoggedOn  As Boolean
        Dim Session   As cUser
        Dim j         As Integer

        'dbg "Client_DoLogon Start"

        Dim ConsoleID As String

10       On Error GoTo Client_DoLogon_Error

20      Set Root = doc.selectSingleNode("HMC")

30      If Not Root Is Nothing Then
40        If Not Root.attributes Is Nothing Then
50          For Each attr In Root.attributes
60            If attr.baseName = "ConsoleID" Then
70              ConsoleID = attr.text
                'dbg "Client_DoLogon ConsoleID: " & ConsoleID
80            End If
90          Next
100       End If

110       For Each Node In Root.childnodes
120         Select Case LCase(Node.baseName)
              Case "password"
130             Password = Node.text
                'dbg "Client_DoLogon Password: " & Password
140             Set User = GetUser(Password)
150             User.ConsoleID = ConsoleID
160             User.Session = GetNextSession()
170             Select Case User.LEvel
                  Case LEVEL_FACTORY
                    ' bump any other admin

180                 For j = HostSessions.Count To 1 Step -1
190                   Set Session = HostSessions(j)
200                   If Session.LEvel >= LEVEL_SUPERVISOR Then
                        dbg "Bumping Admin, Factory Logging In"
                        ' log off master
210                     HostSessions.Remove j
220                     If Session.Session = gUser.Session Then
                          LogRemoteSession Session.Session, 0, "Bumping Admin, Factory Logging In"
230                       InvalidateHostLogin

                          'frmMain.DoLogin
240                     End If
250                   End If
260                 Next
                    dbg ">>>>>>>>> Logging on as Factory <<<<<<<<<"
270                 LoggedOn = True
280                 HostSessions.Add User
                    User.LastSeen = Now
            
290               Case LEVEL_ADMIN, LEVEL_SUPERVISOR
                    ' logon if no other LEVEL_ADMIN or LEVEL_SUPERVISOR active
300                 LoggedOn = True ' lets be optomistic!
310                 For Each Session In HostSessions
320                    If Session.LEvel >= LEVEL_SUPERVISOR Then
                          dbg ">>>>>>>>> Admin already online Cannot logon <<<<<<<<<"
                          LogRemoteSession Session.Session, 0, "Admin already online Cannot logon"
330                       LoggedOn = False
340                       Exit For
350                    End If
360                 Next
370                 If LoggedOn = True Then
                       dbg ">>>>>>>>> Logging on Admin <<<<<<<<<"
                       dbg ">>>>>>>>> " & User.Username & " <<<<<<<<<"
380                   HostSessions.Add User
                      User.LastSeen = Now
390                 End If
400               Case LEVEL_USER
                    ' logon OK
                    dbg ">>>>>>>>> Logging on User <<<<<<<<<"
                    dbg ">>>>>>>>> " & User.Username & " <<<<<<<<<"
410                 LoggedOn = True
                    User.LastSeen = Now
420                 HostSessions.Add User
430               Case Else
                    ' no valid user
                    dbg ">>>>>>>>> Logon Failure No Such User<<<<<<<<<"
440                 LoggedOn = False
450             End Select

460             Client_DoLogon = UserToXML(User, LoggedOn) ' passes loggedon status to function
                dbg ">>>>>>>>> Response to client sent <<<<<<<<<"
470         End Select
480       Next

          ' here is where we check password, return user token


490     End If


Client_DoLogon_Resume:
500      On Error GoTo 0
510      Exit Function

Client_DoLogon_Error:

520     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_DoLogon." & Erl
530     Resume Client_DoLogon_Resume


End Function

Function UserToXML(User As cUser, ByVal LoggedOn As Boolean) As String
  Dim XML As String
  
  XML = "<?xml version=""1.0""?>" & vbCrLf
  XML = XML & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  XML = XML & "<logon>" & vbCrLf
  XML = XML & taggit("loggedon", IIf(LoggedOn, 1, 0)) & vbCrLf
  XML = XML & taggit("name", User.Username) & vbCrLf
  XML = XML & taggit("level", User.LEvel) & vbCrLf
  XML = XML & taggit("consoleid", User.ConsoleID) & vbCrLf
  XML = XML & taggit("userid", User.UserID) & vbCrLf
  XML = XML & taggit("session", User.Session) & vbCrLf
  XML = XML & taggit("userpermissions", User.UserPermissions.UnParseUserPermissions) & vbCrLf
  
  XML = XML & "</logon>"
  XML = XML & "</HMC>"
  UserToXML = XML
  'dbg "UserToXML = " & xml
End Function

Function Client_DoLogOff(doc As DOMDocument60, ByVal User As String) As String
  'dbg "DoLogOff " & user
End Function



Function Client_AutoEnroll(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode
        Dim attr As IXMLDOMAttribute
        Dim SubFunction   As String

10       On Error GoTo Client_AutoEnroll_Error

20      Set Node = doc.selectSingleNode("HMC/SubFunction")
30      If Not (Node Is Nothing) Then
40        SubFunction = LCase(Node.text)
50        Select Case LCase(SubFunction)
            Case "start"
              ' create new autoenroll object
60            Client_AutoEnroll = CreateRemoteAutoEnroll("")
70          Case "check"
              ' poll the autoenroll object
80            Client_AutoEnroll = CheckRemoteAutoEnroll("")
90          Case "cancel"
              ' cancel the autoenroll object
100           Client_AutoEnroll = CancelRemoteAutoEnroll("")
110         Case Else
120           Client_AutoEnroll = Client_InvalidRequest(SubFunction)
130       End Select
140     Else
          'dbg " No Response" & vbCrLf
          'dbg doc.xml & vbCrLf
150     End If

Client_AutoEnroll_Resume:
160      On Error GoTo 0
170      Exit Function

Client_AutoEnroll_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_AutoEnroll." & Erl
190     Resume Client_AutoEnroll_Resume

End Function
Function CreateRemoteAutoEnroll(ByVal Client As String) As String
  Dim str As String '
  Dim Success As Boolean


  Success = RemoteAutoEnroller.start("")
  'dbg "Create RemoteAuto Enroll" & vbCrLf

  If Success Then
    str = "<?xml version=""1.0""?>" & vbCrLf
    str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
    str = str & taggit("Success", "CreateRemoteAutoEnroll") & vbCrLf
    str = str & "</HMC>"
    CreateRemoteAutoEnroll = str
  Else
    str = "<?xml version=""1.0""?>" & vbCrLf
    str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
    str = str & taggit("Fail", "CreateRemoteAutoEnroll") & vbCrLf
    str = str & "</HMC>"
    CreateRemoteAutoEnroll = str
  End If

End Function
Function CheckRemoteAutoEnroll(ByVal Client As String) As String
        Dim XML As String

        'dbg "CheckRemoteAutoEnroll" & vbCrLf

10       On Error GoTo CheckRemoteAutoEnroll_Error

20      If RemoteAutoEnroller Is Nothing Then
          ' return AutoEnrollerError
30        XML = "<?xml version=""1.0""?>" & vbCrLf
40        XML = XML & "<HMC revision=""" & App.Revision & """>" & vbCrLf
50        XML = XML & taggit("Failure", "CheckRemoteAutoEnroll") & vbCrLf
60        XML = XML & "</HMC>"

          'dbg " No Object" & vbCrLf


70      ElseIf RemoteAutoEnroller.Ready Then
          ' return AutoEnroller.Serial, AutoEnroller.MIDPTI

          'dbg " Ready" & vbCrLf
80        CheckRemoteAutoEnroll = RemoteAutoEnroller.GetAutoEnrollResult()

90      ElseIf RemoteAutoEnroller.CheckTimeout Then
          'dbg " Time Out" & vbCrLf

          ' return AutoEnrollerCancel/Timedout
100       CheckRemoteAutoEnroll = RemoteAutoEnroller.TimeOutError
110     Else
120       CheckRemoteAutoEnroll = RemoteAutoEnroller.ReturnAutoEnrollWaiting()
          'dbg " Waiting" & vbCrLf


130     End If

CheckRemoteAutoEnroll_Resume:
140      On Error GoTo 0
150      Exit Function

CheckRemoteAutoEnroll_Error:

160     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.CheckRemoteAutoEnroll." & Erl
170     Resume CheckRemoteAutoEnroll_Resume


End Function
Function CancelRemoteAutoEnroll(ByVal Client As String) As String
      'dbg " CancelRemoteAutoEnroll" & vbCrLf
10       On Error GoTo CancelRemoteAutoEnroll_Error

20      If RemoteAutoEnroller Is Nothing Then
          ' return AutoEnrollerCancel
          'dbg " No Object" & vbCrLf
30        CancelRemoteAutoEnroll = ""
40      Else
          ' kill autorenroller
          'dbg " Cancelling" & vbCrLf
50        CancelRemoteAutoEnroll = RemoteAutoEnroller.Cancel()
60      End If


CancelRemoteAutoEnroll_Resume:
70       On Error GoTo 0
80       Exit Function

CancelRemoteAutoEnroll_Error:

90      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.CancelRemoteAutoEnroll." & Erl
100     Resume CancelRemoteAutoEnroll_Resume


End Function

Function GetDukaneRequests() As Long
  ' rename to process Dukanes
  
  Dim packet As cESPacket
  'On Error Resume Next
  If gDukane.Enabled Then
    gDukane.ReadDuke
    gDukane.ProcessPackets
    Do While gDukane.ESPackets.Count
      
      Set packet = gDukane.GetNextESPacket
      
      If Not packet Is Nothing Then
        ProcessESPacket packet
      End If
    Loop
  End If
End Function


' first request comes in....
' handle it, return with response
' client cannot send new request until ACKed or Timeout (10 seconds?)
' Transmission is held in client object until processed.
' client may send new request after either ACK or timeout.
' response it sent, but no ACK from Client



Function GetRemoteRequests()

        Static Busy As Boolean

        Dim Client As cClientConnection
10      On Error GoTo GetRemoteRequests_Error
        'dbg "ClientConnections.count = " & ClientConnections.count
20      For Each Client In ClientConnections
30        If Client.RequestPending Then
40          ProcessClientRequest Client
            'Exit For ' if we wish to handle one request per pass at a time
50        End If
60      Next

GetRemoteRequests_Resume:
70       On Error GoTo 0
80       Exit Function

GetRemoteRequests_Error:

90      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.GetRemoteRequests." & Erl
100     Resume GetRemoteRequests_Resume


End Function


Function ReturnFailure(ByVal Action As String) As String
  Dim str As String '

  str = "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & taggit("Fail", Action) & vbCrLf
  str = str & "</HMC>"
  ReturnFailure = str

End Function
Function ReturnSuccess(ByVal Action As String, ByVal Value As String) As String
  Dim str As String '

  str = "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & taggit("Success", Action) & vbCrLf
  str = str & taggit("Value", Value) & vbCrLf
  str = str & "</HMC>"

  ReturnSuccess = str

End Function



Function ProcessClientRequest(Client As cClientConnection)
  'call appropriate functionality

  Dim doc                As DOMDocument60: Set doc = New DOMDocument60
  Dim Root               As IXMLDOMNode
  Dim Node               As IXMLDOMNode   ' node matching criteria
  Dim NodeList           As IXMLDOMNodeList   ' collection of nodes matching criteria
  Dim attr               As IXMLDOMAttribute
  Dim FunctionName       As String
  Dim Request            As String
  Dim User               As String
  Dim ConsoleID          As String
  Dim Session            As Long
  Dim RemoteSerial           As String
  Dim Device             As cESDevice

  On Error GoTo ProcessClientRequest_Error

  Request = Client.Request

  'dbg "modclients.ProcessClientRequest ()" & Request & vbCrLf

  doc.LoadXML Request


  Set Root = doc.selectSingleNode("HMC")
  If Not Root Is Nothing Then
    If Not Root.attributes Is Nothing Then
      For Each attr In Root.attributes
        If attr.baseName = "User" Then
          User = attr.text
        End If

        If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
          RemoteSerial = Trim$(attr.text)
          If Len(RemoteSerial) Then
            Set Device = Devices.Device(RemoteSerial)
            If Not Device Is Nothing Then
              Device.LastSeen = Now
            End If
          End If

        End If


        If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
          ConsoleID = attr.text
          'dbg "ConsoleID: " & ConsoleID
        End If

        If attr.baseName = "session" Then
          Session = Val(attr.text)
          UpdateSessionTime Session
          'dbg "Session: " & Session
        End If

      Next
    End If
  End If

  Set Node = doc.selectSingleNode("HMC/action")
  'dbg "***********"

  'dbg "Process " & MID$(doc.xml, 21, 50)
  dbg "Process " & doc.XML    ' MID$(doc.XML, 43, 58)
  'dbg "***********"
  If Not Node Is Nothing Then
    FunctionName = LCase$(Node.text)
    'dbgHostRemote "FunctionName " & FunctionName
    Select Case FunctionName
    
      Case "requestassist"
        Client.Respond Client_RequestAssist(doc), 200
      
      Case "finalizealarm"
        Client.Respond Client_FinalizeAlarm(doc, User), 200
    
      Case "isassuractive"
        Client.Respond Client_IsAssurActive(), 200


      Case "logout"
        Client.Respond Client_DoLogout(doc, Session), 200
      Case "logon"
        Client.Respond Client_DoLogon(doc), 200
      Case "getalarms"
        DeQuePendingEvents doc
        Client.Respond Client_GetAlarms(ConsoleID, User, Session), 200
      Case "getsubscribedalarms"
        DeQuePendingEvents doc
        Client.Respond Client_GetSubscribedAlarms(ConsoleID, User, Session), 200
      Case "silencealarms"
        Client.Respond Client_SilenceAlarms(doc, User, Session), 200
      Case "unsilencealarms"
        Client.Respond Client_UnSilenceAlarms(doc, User, Session), 200
      
      Case "ackalarm"
        Client.Respond Client_AckAlarm(doc, User, Session), 200
        
        
        
      Case "updatedeviceroomid"
        Client.Respond Client_UpdateDeviceRoomID(doc, User), 200
      Case "updatedeviceresidentid"
        Client.Respond Client_UpdateDeviceResidentID(doc, User), 200
      Case "updateresident"
        Client.Respond Client_UpdateResident(doc, User), 200
      Case "updatestaff"
        Client.Respond Client_UpdateStaff(doc, User), 200
      Case "autoenroll"
        Client.Respond Client_AutoEnroll(doc, User), 200
      Case "saveserialdevice"
        Client.Respond Client_UpdateSerialDevice(doc, User), 200
      Case "savetemperturedevice"
        Client.Respond Client_UpdateTemperatureDevice(doc, User), 200
      Case "savedevice"
        Client.Respond Client_UpdateDevice(doc, User), 200
      Case "deletedevice"
        Client.Respond Client_DeleteDevice(doc, User), 200
      Case "saveroom"
        Client.Respond Client_UpdateRoom(doc, User), 200
      Case "sendtopager"
        Client.Respond Client_SendToPager(doc, User), 200
      Case "sendtogroup"
        Client.Respond Client_SendToGroup(doc, User), 200
      Case "setresidentaway"
        Client.Respond Client_SetResidentAway(doc, User), 200
      Case "setroomaway"
        Client.Respond Client_SetRoomAway(doc, User), 200
      Case "clearassurs"
        Client.Respond Client_ClearAssurs(doc, User, Session), 200
      Case "getroom"
        'dbghostremote "getroom called at host "
        Client.Respond Client_GetRoom(doc, User), 200
      Case "deleteresident"
        Client.Respond Client_DeleteResident(doc, User), 200


    End Select
  Else
    dbg Client_InvalidRequest("No Node HMC/action")

    dbg "Request " & Request
    Client.Respond Client_InvalidRequest("No Node HMC/action"), 404
  End If




ProcessClientRequest_Resume:
  On Error GoTo 0
  Exit Function

ProcessClientRequest_Error:

  'dbg "Error " & Err.Number & " (" & Err.Description & ") at modClients.ProcessClientRequest." & Erl
  Resume ProcessClientRequest_Resume


End Function

Function DeQuePendingEvents(doc As DOMDocument60) As Long  ' # of events processed
  'Dim nodelist           As IXMLDOMNodeList

  Dim EventsNode         As IXMLDOMNode
  Dim EventNode          As IXMLDOMNode
  Dim childnode          As IXMLDOMNode
  Dim QuedEvent          As cQuedEvent

 If InStr(1, doc.XML, "Events", vbTextCompare) <> 0 Then
  'Debug.Assert 0
End If

  Set EventsNode = doc.selectSingleNode("HMC/Events")
  If Not EventsNode Is Nothing Then

    For Each EventNode In EventsNode.childnodes  ' Each event"
      Set QuedEvent = New cQuedEvent
      For Each childnode In EventNode.childnodes  ' Particulars for Each Event"
        Select Case LCase$(childnode.baseName)
          Case "eventname"
            QuedEvent.EventName = childnode.text
          Case "clientunsilencealarms"
            QuedEvent.EventName = childnode.text
          Case "consoleid"
            QuedEvent.ConsoleID = childnode.text
          Case "alarmtype"
            QuedEvent.Alarmtype = childnode.text
        End Select
      Next 'ChildNode
      
      ProcessQuedEvent QuedEvent
    Next 'EventNode
    
  End If

End Function

Function ProcessQuedEvent(QuedEvent As cQuedEvent)
  
  
  Select Case QuedEvent.EventName
    Case "clientsilencealarms"
      Select Case LCase(QuedEvent.Alarmtype)
        Case "alarms"
          alarms.ConsoleAlarmTime(QuedEvent.ConsoleID) = 0
          alarms.ConsoleSilenceTime(QuedEvent.ConsoleID) = CDbl(Now)
        Case "alerts"
          Alerts.ConsoleAlarmTime(QuedEvent.ConsoleID) = 0
          Alerts.ConsoleSilenceTime(QuedEvent.ConsoleID) = CDbl(Now)
        Case "troubles"
          Troubles.ConsoleAlarmTime(QuedEvent.ConsoleID) = 0
          Troubles.ConsoleSilenceTime(QuedEvent.ConsoleID) = CDbl(Now)
        Case "lowbatts"
          LowBatts.ConsoleAlarmTime(QuedEvent.ConsoleID) = 0
          LowBatts.ConsoleSilenceTime(QuedEvent.ConsoleID) = CDbl(Now)
        Case "externs"
          Externs.ConsoleAlarmTime(QuedEvent.ConsoleID) = 0
          Externs.ConsoleSilenceTime(QuedEvent.ConsoleID) = CDbl(Now)
      End Select

    Case "clientunsilencealarms"
      Select Case LCase(QuedEvent.Alarmtype)
        Case "alarms"
          alarms.ConsoleAlarmTime(QuedEvent.ConsoleID) = CDbl(Now)
          alarms.ConsoleSilenceTime(QuedEvent.ConsoleID) = 0
        Case "alerts"
          Alerts.ConsoleAlarmTime(QuedEvent.ConsoleID) = CDbl(Now)
          Alerts.ConsoleSilenceTime(QuedEvent.ConsoleID) = 0
        Case "troubles"
          Troubles.ConsoleAlarmTime(QuedEvent.ConsoleID) = CDbl(Now)
          Troubles.ConsoleSilenceTime(QuedEvent.ConsoleID) = 0
        Case "lowbatts"
          LowBatts.ConsoleAlarmTime(QuedEvent.ConsoleID) = CDbl(Now)
          LowBatts.ConsoleSilenceTime(QuedEvent.ConsoleID) = 0
        Case "externs"
          Externs.ConsoleAlarmTime(QuedEvent.ConsoleID) = CDbl(Now)
          Externs.ConsoleSilenceTime(QuedEvent.ConsoleID) = 0
      End Select




  End Select
End Function


Function Client_DoLogout(doc As DOMDocument60, ByVal SessionID As Long)
  
        Dim Session     As cUser
        Dim j           As Integer
        Dim Node As IXMLDOMNode
10       On Error GoTo Client_DoLogout_Error

20      Set Node = doc.selectSingleNode("HMC/logout")
  
  
  
30      For j = HostSessions.Count To 1 Step -1
40        Set Session = HostSessions(j)
50        If Session.Session = SessionID Then
60          HostSessions.Remove j
            dbg "removing sessionid " & SessionID
            LogRemoteSession Session.Session, 0, "Client Do Logout"
70          Exit For
80        End If
90      Next
        dbg "HostSessions.Count " & HostSessions.Count
100     If j > 0 Then
110       Client_DoLogout = ReturnSuccess("logout", "OK")
120     Else
130       Client_DoLogout = ReturnFailure("logout")
140     End If

Client_DoLogout_Resume:
150      On Error GoTo 0
160      Exit Function

Client_DoLogout_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_DoLogout." & Erl
180     Resume Client_DoLogout_Resume


End Function

Sub UpdateSessionTime(ByVal SessionID As Long)
        Dim Session As cUser
10       On Error GoTo UpdateSessionTime_Error

20      For Each Session In HostSessions
30        If Session.Session = SessionID Then
40          Session.LastSeen = Now
50          Exit For
60        End If
70      Next

UpdateSessionTime_Resume:
80       On Error GoTo 0
90       Exit Sub

UpdateSessionTime_Error:

100     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.UpdateSessionTime." & Erl
110     Resume UpdateSessionTime_Resume

  
End Sub


Function Client_ClearAssurs(doc As DOMDocument60, ByVal User As String, ByVal Session As Long) As String


  Dim myalarms           As Long
  Dim RemoteSerial       As String
  Dim ConsoleID          As String
  Dim Root               As IXMLDOMNode
  Dim Node               As IXMLDOMNode
  Dim NodeList           As IXMLDOMNodeList
  Dim attr               As IXMLDOMAttribute

  On Error GoTo Client_ClearAssurs_Error

  Set Root = doc.selectSingleNode("HMC")
  If Not Root Is Nothing Then
    If Not Root.attributes Is Nothing Then
      For Each attr In Root.attributes
        If 0 = StrComp(attr.baseName, "User", vbTextCompare) Then
          User = attr.text
        End If
        If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
          RemoteSerial = Trim$(attr.text)
        End If
        If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
          ConsoleID = attr.text
        End If
      Next
    End If
    Set NodeList = Root.childnodes
    For Each Node In NodeList
      Select Case LCase$(Node.baseName)
        Case "myalarms"
          myalarms = Val(Node.text)
          Exit For
      End Select
    Next
  End If


  If InAssurPeriod = False Then
    Assurs.Clear
    Set AssureVacationDevices = New Collection
    frmMain.ProcessAssurs False
    frmMain.Assur = False

    If myalarms Then
      Client_ClearAssurs = Client_GetSubscribedAlarms(ConsoleID, User, Session)
    Else
      Client_ClearAssurs = Client_GetAlarms(ConsoleID, User, Session)
    End If

  Else
    Client_ClearAssurs = ReturnFailure("clearassurs")
  End If

Client_ClearAssurs_Resume:
  On Error GoTo 0
  Exit Function

Client_ClearAssurs_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_ClearAssurs." & Erl
  Resume Client_ClearAssurs_Resume


End Function

Function Client_SendToPager(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode

        Dim message As String
        Dim PagerID As Long
        Dim Phone   As String

10       On Error GoTo Client_SendToPager_Error

20      Set Root = doc.selectSingleNode("HMC")
30      If Not Root Is Nothing Then
40        For Each Node In Root.childnodes
50          Select Case LCase(Node.baseName)
              Case "message"
60              message = Node.text
70            Case "pagerid"
80              PagerID = Val(Node.text)

90          End Select
100       Next
110       SendToPager message, PagerID, 0, Phone, "", PAGER_NORMAL, message, 0, 0 ' need to fix this maybe for Contral office thing
120       Client_SendToPager = ReturnSuccess("sendtopager", "OK")
130     Else
140       Client_SendToPager = ReturnFailure("sendtopager")
150     End If

Client_SendToPager_Resume:
160      On Error GoTo 0
170      Exit Function

Client_SendToPager_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_SendToPager." & Erl
190     Resume Client_SendToPager_Resume


End Function

Function Client_IsAssurActive()
  Client_IsAssurActive = ReturnSuccess("isassuractive", IsAssurActive())
End Function
Function Client_SetResidentAway(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode

        Dim Away        As Integer
        Dim ResidentID  As Long


10       On Error GoTo Client_SetResidentAway_Error

20      Set Root = doc.selectSingleNode("HMC")
30      If Not Root Is Nothing Then
40        For Each Node In Root.childnodes
50          Select Case LCase(Node.baseName)
              Case "resid"
60              ResidentID = Val(Node.text)
70            Case "away"
80              Away = Val(Node.text)

90          End Select
100       Next
110       Away = SetResidentAwayStatus(Away, ResidentID, User)
120       Client_SetResidentAway = ReturnSuccess("setresidentaway", CStr(Away))
130     Else
140       Client_SetResidentAway = ReturnFailure("setresidentaway")
150     End If

Client_SetResidentAway_Resume:
160      On Error GoTo 0
170      Exit Function

Client_SetResidentAway_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_SetResidentAway." & Erl
190     Resume Client_SetResidentAway_Resume


End Function

Function Client_SetRoomAway(doc As DOMDocument60, ByVal User As String) As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode

  Dim Away        As Integer
  Dim RoomID  As Long


   On Error GoTo Client_SetRoomAway_Error

  Set Root = doc.selectSingleNode("HMC")
  If Not Root Is Nothing Then
    For Each Node In Root.childnodes
      Select Case LCase(Node.baseName)
        Case "roomid"
          RoomID = Val(Node.text)
        Case "away"
          Away = Val(Node.text)

      End Select
    Next
    Away = SetRoomAwayStatus(Away, RoomID, User)
    Client_SetRoomAway = ReturnSuccess("setroomaway", CStr(Away))
  Else
    Client_SetRoomAway = ReturnFailure("setroomaway")
  End If

Client_SetRoomAway_Resume:
   On Error GoTo 0
   Exit Function

Client_SetRoomAway_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_SetRoomAway." & Erl
  Resume Client_SetRoomAway_Resume


End Function


Function Client_SendToGroup(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode

        Dim message As String
        Dim GroupID As Long
        Dim Phone   As String

   On Error GoTo Client_SendToGroup_Error

10      Set Root = doc.selectSingleNode("HMC")
20      If Not Root Is Nothing Then
30        For Each Node In Root.childnodes
40          Select Case LCase(Node.baseName)
              Case "message"
50              message = Node.text
                 
60            Case "groupid"
70              GroupID = Val(Node.text)

80          End Select
90        Next
100       SendToGroup message, GroupID, Phone, "", PAGER_NORMAL, message, 0, 0 ' fix this for Central Office
110       Client_SendToGroup = ReturnSuccess("sendtogroup", "OK")
120     Else
130       Client_SendToGroup = ReturnFailure("sendtogroup")
140     End If





Client_SendToGroup_Resume:
   On Error GoTo 0
   Exit Function

Client_SendToGroup_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_SendToGroup." & Erl
  Resume Client_SendToGroup_Resume


End Function

Function Client_UpdateSerialDevice(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode
        Dim attr As IXMLDOMAttribute

        Dim Device As cESDevice: Set Device = New cESDevice
        Dim Action   As String
        ' ?? RoomID ???
10       On Error GoTo Client_UpdateSerialDevice_Error

20      Set Root = doc.selectSingleNode("HMC/device")
30      If Not Root Is Nothing Then
40        For Each Node In Root.childnodes
50          Select Case LCase(Node.baseName)
              Case "deviceid"
60              Device.DeviceID = Val(Node.text)
70            Case "serial"
80              Device.Serial = Node.text
              Case "serialtapprotocol"
                Device.SerialTapProtocol = Val(Node.text)
90            Case "serialskip"
100             Device.SerialSkip = Val(Node.text)
110           Case "serialmessagelen"
120             Device.SerialMessageLen = Val(Node.text)
130           Case "serialautoclear"
140             Device.SerialAutoClear = Val(Node.text)
150           Case "serialinclude"
160             Device.SerialInclude = Node.text
170           Case "serialexclude"
180             Device.SerialExclude = Node.text
190           Case "serialport"
200             Device.SerialPort = Val(Node.text)
210           Case "serialbaud"
220             Device.SerialBaud = Val(Node.text)
230           Case "serialparity"
240             Device.SerialParity = Node.text
250           Case "serialbits"
260             Device.Serialbits = Val(Node.text)
270           Case "serialflow"
280             Device.SerialFlow = Val(Node.text)
290           Case "serialstopbits"
300             Device.SerialStopbits = Val(Node.text)
310           Case "serialsettings"
320             Device.SerialSettings = Node.text
330           Case "serialeolchar"
340             Device.SerialEOLChar = Val(Node.text)
350           Case "serialpreamble"
360             Device.SerialPreamble = Node.text



370         End Select
380       Next

390       If SaveSerialDevice(Device, User) Then
400         Client_UpdateSerialDevice = ReturnSuccess("SaveSerialDevice", Device.Serial)
410       Else
420         Client_UpdateSerialDevice = ReturnFailure("SaveSerialDevice")
430       End If
440     Else
450       Client_UpdateSerialDevice = ReturnFailure("SaveSerialDevice")
460     End If


Client_UpdateSerialDevice_Resume:
470      On Error GoTo 0
480      Exit Function

Client_UpdateSerialDevice_Error:

490     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateSerialDevice." & Erl
500     Resume Client_UpdateSerialDevice_Resume


End Function

Function Client_UpdateTemperatureDevice(doc As DOMDocument60, ByVal User As String) As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  Dim attr As IXMLDOMAttribute

  Dim Device As cESDevice: Set Device = New cESDevice
  Dim Action   As String
  
   On Error GoTo Client_UpdateTemperatureDevice_Error

  Set Root = doc.selectSingleNode("HMC/device")
  If Not Root Is Nothing Then
    For Each Node In Root.childnodes
      Select Case LCase(Node.baseName)
        Case "deviceid"
          Device.DeviceID = Val(Node.text)
        Case "serial"
          Device.Serial = Node.text
        Case "serialpreamble"
          Device.SerialPreamble = Node.text
        Case "lowset"
          Device.LowSet = Val(Node.text)
        Case "lowset_a"
          Device.LowSet_A = Val(Node.text)
        Case "hiset"
          Device.HiSet = Val(Node.text)
        Case "hiset_a"
          Device.HiSet_A = Val(Node.text)
        Case "enabletemperature"
          Device.EnableTemperature = Val(Node.text)
        Case "enabletemperature_a"
          Device.EnableTemperature_A = Val(Node.text)

      End Select
    Next

    If SaveTemperatureDevice(Device, User) Then
      Client_UpdateTemperatureDevice = ReturnSuccess("SaveTemperatureDevice", Device.Serial)
    Else
      Client_UpdateTemperatureDevice = ReturnFailure("SaveTemperatureDevice")
    End If
  Else
    Client_UpdateTemperatureDevice = ReturnFailure("SaveTemperatureDevice")
  End If


Client_UpdateTemperatureDevice_Resume:
   On Error GoTo 0
   Exit Function

Client_UpdateTemperatureDevice_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateTemperatureDevice." & Erl
  Resume Client_UpdateTemperatureDevice_Resume


End Function
Function Client_DeleteResident(doc As DOMDocument60, ByVal User As String) As String

        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode
        Dim attr As IXMLDOMAttribute
        Dim ResidentID   As Long

        Dim Action   As String: Action = "deleteresident"
  
10       On Error GoTo Client_DeleteResident_Error

20      Set Root = doc.selectSingleNode("HMC/residentid")
30      If Not (Root Is Nothing) Then
40        ResidentID = Val(Root.text)
50      End If
        'dbg "Attempting to delete ResidentID " & ResidentID
60      If ResidentID <> 0 Then
70        If DeleteResident(ResidentID, User) = 0 Then
80          Client_DeleteResident = ReturnSuccess(Action, CStr(ResidentID))
90        Else
100         Client_DeleteResident = ReturnFailure(Action)
110       End If
120     Else
130       Client_DeleteResident = ReturnFailure(Action)
140     End If


Client_DeleteResident_Resume:
150      On Error GoTo 0
160      Exit Function

Client_DeleteResident_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_DeleteResident." & Erl
180     Resume Client_DeleteResident_Resume




End Function

Function Client_DeleteDevice(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode
        Dim attr As IXMLDOMAttribute
        Dim DeviceID   As Long

        Dim Action   As String: Action = "deletedevice"
        ' ?? RoomID ???
10       On Error GoTo Client_DeleteDevice_Error

20      Set Root = doc.selectSingleNode("HMC/deviceid")
30      If Not (Root Is Nothing) Then
40        DeviceID = Val(Root.text)
50      End If
        'dbg "Attempting to delete DeviceID " & DeviceID
60      If DeviceID <> 0 Then
70        If DeleteTransmitter(DeviceID, User) = 0 Then
80          Client_DeleteDevice = ReturnSuccess(Action, CStr(DeviceID))
90        Else
100         Client_DeleteDevice = ReturnFailure(Action)
110       End If
120     Else
130       Client_DeleteDevice = ReturnFailure(Action)
140     End If


Client_DeleteDevice_Resume:
150      On Error GoTo 0
160      Exit Function

Client_DeleteDevice_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_DeleteDevice." & Erl
180     Resume Client_DeleteDevice_Resume


End Function
Function Client_UpdateDevice(doc As DOMDocument60, ByVal User As String) As String
        Dim Root          As IXMLDOMNode
        Dim Node          As IXMLDOMNode
        Dim attr          As IXMLDOMAttribute

10      Dim Device        As cESDevice: Set Device = New cESDevice
20      Dim Action        As String: Action = "savedevice"
        ' ?? RoomID ???
30      On Error GoTo Client_UpdateDevice_Error

40      Set Root = doc.selectSingleNode("HMC/device")
50      If Not Root Is Nothing Then
          'dbg "Root.childNodes.length " & Root.childNodes.Length & vbCrLf
60        For Each Node In Root.childnodes
70          Select Case LCase(Node.baseName)
              Case "deviceid"
80              Device.DeviceID = Val(Node.text)
90            Case "serial"
100             Device.Serial = Node.text
110           Case "alarmmask"
120             Device.AlarmMask = Val(Node.text)
130           Case "alarmmask_a"
140             Device.AlarmMask_A = Val(Node.text)
150           Case "alarmmask_b"
160             Device.AlarmMask_B = Val(Node.text)

170           Case "alarmmask2"
180             Device.AlarmMask2 = Val(Node.text)
190           Case "alarmmask2_a"
200             Device.AlarmMask2_A = Val(Node.text)
210           Case "alarmmask2_b"
220             Device.AlarmMask2_B = Val(Node.text)

230           Case "alert"
240             Device.Alert = Val(Node.text)
250           Case "alert_a"
260             Device.Alert_A = Val(Node.text)
270           Case "alert_b"
280             Device.Alert_B = Val(Node.text)

290           Case "announce"
300             Device.Announce = Node.text
310           Case "announce_a"
320             Device.Announce_A = Node.text
330           Case "announce_b"
340             Device.Announce_B = Node.text
350           Case "assurbit"
360             Device.AssurBit = Val(Node.text)
370           Case "assurbit_a"
380             Device.AssurBit_A = Val(Node.text)
390           Case "assurbit_b"
400             Device.AssurBit_B = Val(Node.text)

410           Case "assurinput"
420             Device.AssurInput = Val(Node.text)
430           Case "assursecure"
440             Device.AssurSecure = Val(Node.text)
450           Case "assursecure_a"
460             Device.AssurSecure_A = Val(Node.text)
470           Case "assursecure_b"
480             Device.AssurSecure_B = Val(Node.text)

490           Case "custom"
500             Device.Custom = Node.text

510           Case "building"
520             Device.Building = Node.text
530           Case "clearbyreset"
540             Device.ClearByReset = Val(Node.text)
550           Case "description"
560             Device.Description = Node.text
570           Case "disableend"
580             Device.DisableEnd = Val(Node.text)
590           Case "disableend_a"
600             Device.DisableEnd_A = Val(Node.text)
610           Case "disableend_b"
620             Device.DisableEnd_B = Val(Node.text)

630           Case "disableend2"
640             Device.DisableEnd2 = Val(Node.text)
650           Case "disableend2_a"
660             Device.DisableEnd2_A = Val(Node.text)
670           Case "disableend2_b"
680             Device.DisableEnd2_B = Val(Node.text)

690           Case "disablestart"
700             Device.DisableStart = Val(Node.text)
710           Case "disablestart_a"
720             Device.DisableStart_A = Val(Node.text)
730           Case "disablestart_b"
740             Device.DisableStart_B = Val(Node.text)

750           Case "disablestart2"
760             Device.DisableStart2 = Val(Node.text)
770           Case "disablestart2_a"
780             Device.DisableStart2_A = Val(Node.text)
790           Case "disablestart2_b"
800             Device.DisableStart2_B = Val(Node.text)

810           Case "isaway"
820             Device.IsAway = Val(Node.text)
830           Case "isaway_a"
840             Device.IsAway_A = Val(Node.text)
850           Case "islatching"
860             Device.IsLatching = Val(Node.text)
870           Case "islocator"
880             Device.IsLocator = Val(Node.text)
890           Case "isportable"
900             Device.IsPortable = Val(Node.text)
910           Case "midpti"
920             Device.MIDPTI = Val(Node.text)
930           Case "clspti"
940             Device.CLSPTI = Val(Node.text)
950           Case "model"
960             Device.Model = Node.text
970           Case "notamper"
980             Device.NoTamper = Val(Node.text)

                ' output groups button 1
990           Case "og1"
1000            Device.OG1 = Val(Node.text)
1010          Case "og2"
1020            Device.OG2 = Val(Node.text)
1030          Case "og3"
1040            Device.OG3 = Val(Node.text)
1050          Case "og4"
1060            Device.OG4 = Val(Node.text)
1070          Case "og5"
1080            Device.OG5 = Val(Node.text)
1090          Case "og6"
1100            Device.OG6 = Val(Node.text)

1110          Case "ng1"
1120            Device.NG1 = Val(Node.text)
1130          Case "ng2"
1140            Device.NG2 = Val(Node.text)
1150          Case "ng3"
1160            Device.NG3 = Val(Node.text)
1170          Case "ng4"
1180            Device.NG4 = Val(Node.text)
1190          Case "ng5"
1200            Device.NG5 = Val(Node.text)
1210          Case "ng6"
1220            Device.NG6 = Val(Node.text)

1230          Case "gg1"
1240            Device.GG1 = Val(Node.text)
1250          Case "gg2"
1260            Device.GG2 = Val(Node.text)
1270          Case "gg3"
1280            Device.GG3 = Val(Node.text)
1290          Case "gg4"
1300            Device.GG4 = Val(Node.text)
1310          Case "gg5"
1320            Device.GG5 = Val(Node.text)
1330          Case "gg6"
1340            Device.GG6 = Val(Node.text)


                ' output groups button 2

1350          Case "og1_a"
1360            Device.OG1_A = Val(Node.text)
1370          Case "og2_a"
1380            Device.OG2_A = Val(Node.text)
1390          Case "og3_a"
1400            Device.OG3_A = Val(Node.text)
1410          Case "og4_a"
1420            Device.OG4_A = Val(Node.text)
1430          Case "og5_a"
1440            Device.OG5_A = Val(Node.text)
1450          Case "og6_a"
1460            Device.OG6_A = Val(Node.text)

1470          Case "ng1_a"
1480            Device.NG1_A = Val(Node.text)
1490          Case "ng2_a"
1500            Device.NG2_A = Val(Node.text)
1510          Case "ng3_a"
1520            Device.NG3_A = Val(Node.text)
1530          Case "ng4_a"
1540            Device.NG4_A = Val(Node.text)
1550          Case "ng5_a"
1560            Device.NG5_A = Val(Node.text)
1570          Case "ng6_a"
1580            Device.NG6_A = Val(Node.text)

1590          Case "gg1_a"
1600            Device.GG1_A = Val(Node.text)
1610          Case "gg2_a"
1620            Device.GG2_A = Val(Node.text)
1630          Case "gg3_a"
1640            Device.GG3_A = Val(Node.text)
1650          Case "gg4_a"
1660            Device.GG4_A = Val(Node.text)
1670          Case "gg5_a"
1680            Device.GG5_A = Val(Node.text)
1690          Case "gg6_a"
1700            Device.GG6_A = Val(Node.text)


                ' output groups button 3
1710          Case "og1_b"
1720            Device.OG1_B = Val(Node.text)
1730          Case "og2_b"
1740            Device.OG2_B = Val(Node.text)
1750          Case "og3_b"
1760            Device.OG3_B = Val(Node.text)
1770          Case "og4_b"
1780            Device.OG4_B = Val(Node.text)
1790          Case "og5_b"
1800            Device.OG5_B = Val(Node.text)
1810          Case "og6_b"
1820            Device.OG6_B = Val(Node.text)

1830          Case "ng1_b"
1840            Device.NG1_B = Val(Node.text)
1850          Case "ng2_b"
1860            Device.NG2_B = Val(Node.text)
1870          Case "ng3_b"
1880            Device.NG3_B = Val(Node.text)
1890          Case "ng4_b"
1900            Device.NG4_B = Val(Node.text)
1910          Case "ng5_b"
1920            Device.NG5_B = Val(Node.text)
1930          Case "ng6_b"
1940            Device.NG6_B = Val(Node.text)

1950          Case "gg1_b"
1960            Device.GG1_B = Val(Node.text)
1970          Case "gg2_b"
1980            Device.GG2_B = Val(Node.text)
1990          Case "gg3_b"
2000            Device.GG3_B = Val(Node.text)
2010          Case "gg4_b"
2020            Device.GG4_B = Val(Node.text)
2030          Case "gg5_b"
2040            Device.GG5_B = Val(Node.text)
2050          Case "gg6_b"
2060            Device.GG6_B = Val(Node.text)

                ' duration timers button 1
2070          Case "og1d"
2080            Device.OG1D = Val(Node.text)
2090          Case "og2d"
2100            Device.OG2D = Val(Node.text)
2110          Case "og3d"
2120            Device.OG3D = Val(Node.text)
2130          Case "og4d"
2140            Device.OG4D = Val(Node.text)
2150          Case "og5d"
2160            Device.OG5D = Val(Node.text)
2170          Case "og6d"
2180            Device.OG6D = Val(Node.text)

2190          Case "ng1d"
2200            Device.NG1D = Val(Node.text)
2210          Case "ng2d"
2220            Device.NG2D = Val(Node.text)
2230          Case "ng3d"
2240            Device.NG3D = Val(Node.text)
2250          Case "ng4d"
2260            Device.NG4D = Val(Node.text)
2270          Case "ng5d"
2280            Device.NG5D = Val(Node.text)
2290          Case "ng6d"
2300            Device.NG6D = Val(Node.text)

2310          Case "gg1d"
2320            Device.GG1D = Val(Node.text)
2330          Case "gg2d"
2340            Device.GG2D = Val(Node.text)
2350          Case "gg3d"
2360            Device.GG3D = Val(Node.text)
2370          Case "gg4d"
2380            Device.GG4D = Val(Node.text)
2390          Case "gg5d"
2400            Device.GG5D = Val(Node.text)
2410          Case "gg6d"
2420            Device.GG6D = Val(Node.text)

                ' duration timers button 2

2430          Case "og1_ad"
2440            Device.OG1_AD = Val(Node.text)
2450          Case "og2_ad"
2460            Device.OG2_AD = Val(Node.text)
2470          Case "og3_ad"
2480            Device.OG3_AD = Val(Node.text)
2490          Case "og4_ad"
2500            Device.OG4_AD = Val(Node.text)
2510          Case "og5_ad"
2520            Device.OG5_AD = Val(Node.text)
2530          Case "og6_ad"
2540            Device.OG6_AD = Val(Node.text)


2550          Case "ng1_ad"
2560            Device.NG1_AD = Val(Node.text)
2570          Case "ng2_ad"
2580            Device.NG2_AD = Val(Node.text)
2590          Case "ng3_ad"
2600            Device.NG3_AD = Val(Node.text)
2610          Case "ng4_ad"
2620            Device.NG4_AD = Val(Node.text)
2630          Case "ng5_ad"
2640            Device.NG5_AD = Val(Node.text)
2650          Case "ng6_ad"
2660            Device.NG6_AD = Val(Node.text)

2670          Case "gg1_ad"
2680            Device.GG1_AD = Val(Node.text)
2690          Case "gg2_ad"
2700            Device.GG2_AD = Val(Node.text)
2710          Case "gg3_ad"
2720            Device.GG3_AD = Val(Node.text)
2730          Case "gg4_ad"
2740            Device.GG4_AD = Val(Node.text)
2750          Case "gg5_ad"
2760            Device.GG5_AD = Val(Node.text)
2770          Case "gg6_ad"
2780            Device.GG6_AD = Val(Node.text)


                ' duration timers button

2790          Case "og1_bd"
2800            Device.OG1_BD = Val(Node.text)
2810          Case "og2_bd"
2820            Device.OG2_BD = Val(Node.text)
2830          Case "og3_bd"
2840            Device.OG3_BD = Val(Node.text)
2850          Case "og4_bd"
2860            Device.OG4_BD = Val(Node.text)
2870          Case "og5_bd"
2880            Device.OG5_BD = Val(Node.text)
2890          Case "og6_bd"
2900            Device.OG6_BD = Val(Node.text)


2910          Case "ng1_bd"
2920            Device.NG1_BD = Val(Node.text)
2930          Case "ng2_bd"
2940            Device.NG2_BD = Val(Node.text)
2950          Case "ng3_bd"
2960            Device.NG3_BD = Val(Node.text)
2970          Case "ng4_bd"
2980            Device.NG4_BD = Val(Node.text)
2990          Case "ng5_bd"
3000            Device.NG5_BD = Val(Node.text)
3010          Case "ng6_bd"
3020            Device.NG6_BD = Val(Node.text)

3030          Case "gg1_bd"
3040            Device.GG1_BD = Val(Node.text)
3050          Case "gg2_bd"
3060            Device.GG2_BD = Val(Node.text)
3070          Case "gg3_bd"
3080            Device.GG3_BD = Val(Node.text)
3090          Case "gg4_bd"
3100            Device.GG4_BD = Val(Node.text)
3110          Case "gg5_bd"
3120            Device.GG5_BD = Val(Node.text)
3130          Case "gg6_bd"
3140            Device.GG6_BD = Val(Node.text)


3150          Case "owner"
3160            Device.Owner = Val(Node.text)
3170          Case "owner_a"
                'Device.Owner_A = ""
3180          Case "pause"
3190            Device.Pause = Val(Node.text)
3200          Case "pause_a"
3210            Device.Pause_A = Val(Node.text)
3220          Case "pause_b"
3230            Device.Pause_B = Val(Node.text)

3240          Case "repeats"
3250            Device.Repeats = Val(Node.text)
3260          Case "repeats_a"
3270            Device.Repeats_A = Val(Node.text)
3280          Case "repeats_b"
3290            Device.Repeats_B = Val(Node.text)

3300          Case "repeatuntil"
3310            Device.RepeatUntil = Val(Node.text)
3320          Case "repeatuntil_a"
3330            Device.RepeatUntil_A = Val(Node.text)
3340          Case "repeatuntil_b"
3350            Device.RepeatUntil_B = Val(Node.text)


3360          Case "residentid"
3370            Device.ResidentID = Val(Node.text)
3380          Case "residentid_a"
                'Device.ResidentID_A = Val(Node.text)
3390          Case "room"
3400            Device.Room = Node.text
3410          Case "room_a"
                'Device.Room_A = Node.text
3420          Case "roomid"
3430            Device.RoomID = Val(Node.text)
3440          Case "roomid_a"
                'Device.RoomID_A = Val(Node.text)
3450          Case "sendcancel"
3460            Device.SendCancel = Val(Node.text)
3470          Case "sendcancel_a"
3480            Device.SendCancel_A = Val(Node.text)
3490          Case "sendcancel_b"
3500            Device.SendCancel_B = Val(Node.text)

3510          Case "superviseperiod"
3520            Device.SupervisePeriod = Val(Node.text)
3530          Case "useassur"
3540            Device.UseAssur = Val(Node.text)
3550          Case "useassur_a"
3560            Device.UseAssur_A = Val(Node.text)
3570          Case "useassur_b"
3580            Device.UseAssur_B = Val(Node.text)

3590          Case "useassur2"
3600            Device.UseAssur2 = Val(Node.text)
3610          Case "useassur2_a"
3620            Device.UseAssur2_A = Val(Node.text)
3630          Case "useassur2_b"
3640            Device.UseAssur2_B = Val(Node.text)

              Case "tamperasinput"
                Device.UseTamperAsInput = Val(Node.text)
              
              Case "ignored"
                Device.Ignored = Val(Node.text)
              
              Case "configurationstring"
                Device.Configurationstring = Node.text
              Case "lockw"
                

3650        End Select
3660      Next

          Dim ID          As Long
          'dbg "modclients.ClientUpdateDevice Before SaveDevice " & vbCrLf
3670      ID = SaveDevice(Device, User)
          'dbg "modclients.ClientUpdateDevice After SaveDevice ID=" & ID & vbCrLf
3680      If ID <> 0 Then
            'dbg "modclients.ClientUpdateDevice Returning Success" & vbCrLf
3690        Client_UpdateDevice = ReturnSuccess(Action, Device.Serial)

3700      Else
            'dbg "modclients.ClientUpdateDevice Returning Failure" & vbCrLf
3710        Client_UpdateDevice = ReturnFailure(Action)
3720      End If
3730    Else
          'dbg "No Root " & vbCrLf
3740      Client_UpdateDevice = ReturnFailure(Action)
3750    End If



Client_UpdateDevice_Resume:
3760    On Error GoTo 0
3770    Exit Function

Client_UpdateDevice_Error:

3780    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateDevice." & Erl
3790    Resume Client_UpdateDevice_Resume


End Function
Function Client_GetRoom(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode
        Dim attr As IXMLDOMAttribute
        Dim RoomID As Long
        Dim Room As cRoom: Set Room = New cRoom
        Dim Action   As String: Action = "getroom"
        Dim XML    As String

10       On Error GoTo Client_GetRoom_Error

20    On Error GoTo 0

30      Set Root = doc.selectSingleNode("HMC/roomid")
40      If Not Root Is Nothing Then
50        RoomID = Val(Root.text)
60        XML = ClientGetRoom(RoomID)
70      Else
80        XML = ClientGetRoom(0)
90      End If
100   'Debug.Print xml

110   Client_GetRoom = XML
120     On Error GoTo 0
130     Exit Function


Client_GetRoom_Resume:
140      On Error GoTo 0
150      Exit Function

Client_GetRoom_Error:

160     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_GetRoom." & Erl
170     Resume Client_GetRoom_Resume


End Function

Function ClientGetRoom(RoomID As Long) As String
        Dim SQL As String
        Dim rs  As ADODB.Recordset
        Dim XML As String
  
10       On Error GoTo ClientGetRoom_Error

20      SQL = "Select * FROM Rooms WHERE RoomID = " & RoomID
30      Set rs = ConnExecute(SQL)
40      If rs.EOF Then
50        XML = "<?xml version=""1.0""?>" & vbCrLf
60        XML = XML & "<HMC revision=""" & App.Revision & """>" & vbCrLf
70        XML = XML & "<getroom>" & vbCrLf
80        XML = XML & taggit("roomid", "0") & vbCrLf
90        XML = XML & "</getroom>"
100       XML = XML & "</HMC>"
    
    
110     Else
  
120       XML = "<?xml version=""1.0""?>" & vbCrLf
130       XML = XML & "<HMC revision=""" & App.Revision & """>" & vbCrLf
140       XML = XML & "<getroom>" & vbCrLf
150       XML = XML & taggit("roomid", rs("Roomid")) & vbCrLf
160       XML = XML & taggit("room", XMLEncode(rs("room") & "")) & vbCrLf
170       XML = XML & taggit("assurdays", rs("AssurDays")) & vbCrLf
180       XML = XML & taggit("vacation", rs("Away")) & vbCrLf
185       XML = XML & taggit("lockw", XMLEncode(rs("lockw") & "")) & vbCrLf

190       XML = XML & "</getroom>"
200       XML = XML & "</HMC>"
  
210     End If
220     rs.Close
230     Set rs = Nothing
  
240     ClientGetRoom = XML
  



ClientGetRoom_Resume:
250      On Error GoTo 0
260      Exit Function

ClientGetRoom_Error:

270     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.ClientGetRoom." & Erl
280     Resume ClientGetRoom_Resume


End Function
Function Client_UpdateRoom(doc As DOMDocument60, ByVal User As String) As String
        Dim Root As IXMLDOMNode
        Dim Node As IXMLDOMNode
        Dim attr As IXMLDOMAttribute

        Dim Room As cRoom: Set Room = New cRoom
        Dim Action   As String: Action = "saveroom"


10      On Error GoTo Client_UpdateRoom_Error

20      Set Root = doc.selectSingleNode("HMC/room")
  
30      If Not Root Is Nothing Then
40        For Each Node In Root.childnodes
50          Select Case LCase(Node.baseName)
              Case "assurdays"
60              Room.Assurdays = Val(Node.text)
70            Case "away"
80              Room.Away = Val(Node.text)
90            Case "building"
100             Room.Building = Node.text
110           Case "deleted"
120             Room.Deleted = Val(Node.text)
130           Case "description"
140             Room.Description = Node.text
150           Case "room"
160             Room.Room = Node.text
              Case "lockw"
                Room.locKW = Node.text
170           Case "roomid"
180             Room.RoomID = Val(Node.text)
              Case "flags"
                Room.flags = Val(Node.text)
190           Case "vacation"
200             Room.Vacation = Val(Node.text)
210         End Select
220       Next
          Dim rc As Boolean
230       rc = SaveRoom(Room, User)

240       'RefreshJet
250       'Sleep 1
260       RefreshJet

          'dbg "modclients SaveRoom RC " & rc
270       If rc Then
            'dbg "modclients SaveRoom OK"
280         Client_UpdateRoom = ReturnSuccess(Action, CStr(Room.RoomID))
290       Else
            'dbg "modclients SaveRoom Fail"
300         Client_UpdateRoom = ReturnFailure(Action)
310       End If
320     End If

Client_UpdateRoom_Resume:
330     On Error GoTo 0
340     Exit Function

Client_UpdateRoom_Error:

350     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateRoom." & Erl
360     Resume Client_UpdateRoom_Resume


End Function

Function Client_UpdateResident(doc As DOMDocument60, ByVal User As String) As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  Dim attr As IXMLDOMAttribute

  Dim Resident As cResident: Set Resident = New cResident
  Dim Action   As String
  ' ?? RoomID ???


10         On Error GoTo Client_UpdateResident_Error

20        Set Root = doc.selectSingleNode("HMC")
30        If Not Root Is Nothing Then
40          For Each Node In Root.childnodes
50            Select Case LCase(Node.baseName)
                Case "residentid"
60                Resident.ResidentID = Val(Node.text)
70              Case "namelast"
80                Resident.NameLast = Node.text
90              Case "namefirst"
100               Resident.NameFirst = Node.text
110             Case "phone"
120               Resident.Phone = Node.text
130             Case "room"
140               Resident.Room = Node.text
150             Case "info"
160               Resident.info = Node.text
170             Case "assurdays"
180               Resident.Assurdays = Val(Node.text)
185             Case "deliverypoints"
186                Resident.DeliveryPointsString = Node.text
                   Resident.ParseDeliveryPoints

  'case "vacation"
190             Case "away"
200               Resident.Vacation = IIf(Val(Node.text) = 1, 1, 0)
210             Case "action"
220               Action = Node.text
230           End Select
240         Next
  'dbg "modClients.Client_UpdateResident.start" & vbCrLf
250         If UpdateResident(Resident, User) Then
260           Client_UpdateResident = ReturnSuccess(Action, CStr(Resident.ResidentID))
  '  dbg "modClients.Client_UpdateResident.OK" & vbCrLf
270         Else
280           Client_UpdateResident = ReturnFailure(Action)
  '  dbg "modClients.Client_UpdateResident.ERROR" & vbCrLf
290         End If
  'dbg "modClients.Client_UpdateResident.Done" & vbCrLf
300       End If


Client_UpdateResident_Resume:
310        On Error GoTo 0
320        Exit Function

Client_UpdateResident_Error:

330       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateResident." & Erl
340       Resume Client_UpdateResident_Resume


End Function
Function Client_UpdateStaff(doc As DOMDocument60, ByVal User As String) As String
  Dim Root As IXMLDOMNode
  Dim Node As IXMLDOMNode
  Dim attr As IXMLDOMAttribute

  Dim Resident As cResident: Set Resident = New cResident
  Dim Action   As String
  ' ?? RoomID ???


10         On Error GoTo Client_UpdateStaff_Error

20        Set Root = doc.selectSingleNode("HMC")
30        If Not Root Is Nothing Then
40          For Each Node In Root.childnodes
50            Select Case LCase(Node.baseName)
                Case "residentid"
60                Resident.ResidentID = Val(Node.text)
70              Case "namelast"
80                Resident.NameLast = Node.text
90              Case "namefirst"
100               Resident.NameFirst = Node.text
110             Case "phone"
120               Resident.Phone = Node.text
130             Case "room"
140               Resident.Room = Node.text
150             Case "info"
160               Resident.info = Node.text
170             Case "assurdays"
180               Resident.Assurdays = Val(Node.text)
  'case "vacation"
190             Case "away"
200               Resident.Vacation = IIf(Val(Node.text) = 1, 1, 0)
210             Case "action"
220               Action = Node.text
230           End Select
240         Next
  'dbg "modClients.Client_UpdateResident.start" & vbCrLf
250         If UpdateStaff(Resident, User) Then
260           Client_UpdateStaff = ReturnSuccess(Action, CStr(Resident.ResidentID))
  '  dbg "modClients.Client_UpdateResident.OK" & vbCrLf
270         Else
280           Client_UpdateStaff = ReturnFailure(Action)
  '  dbg "modClients.Client_UpdateResident.ERROR" & vbCrLf
290         End If
  'dbg "modClients.Client_UpdateResident.Done" & vbCrLf
300       End If


Client_UpdateStaff_Resume:
310        On Error GoTo 0
320        Exit Function

Client_UpdateStaff_Error:

330       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateStaff." & Erl
340       Resume Client_UpdateStaff_Resume


End Function



Function Client_UpdateDeviceResidentID(doc As DOMDocument60, ByVal User As String) As String
  Dim Root        As IXMLDOMNode
  Dim Node        As IXMLDOMNode
  Dim Action      As String
  Dim ResidentID  As Long
  Dim DeviceID    As Long
  Dim SQL         As String
  Dim OK          As Boolean

10        Set Node = doc.selectSingleNode("HMC/ResidentID")
20        If Not Node Is Nothing Then
30          ResidentID = Val(Node.text)
40          Set Node = doc.selectSingleNode("HMC/DeviceID")
50          If Not Node Is Nothing Then
60            DeviceID = Val(Node.text)
70              On Error Resume Next
80              SQL = "UPDATE Devices SET ResidentID = " & ResidentID & " WHERE DeviceID = " & DeviceID
90              conn.BeginTrans
100             ConnExecute SQL
110             conn.CommitTrans

120             Devices.RefreshByID DeviceID
130             OK = Err.Number = 0

140         End If
150       End If


160       If OK Then
170         Client_UpdateDeviceResidentID = ReturnSuccess(Action, CStr(ResidentID))

180       Else
190         Client_UpdateDeviceResidentID = ReturnFailure(Action)

200       End If

End Function
Function Client_UpdateDeviceRoomID(doc As DOMDocument60, ByVal User As String) As String
        Dim Root        As IXMLDOMNode
        Dim Node        As IXMLDOMNode
        Dim Action      As String
        Dim RoomID      As Long
        Dim DeviceID    As Long
        Dim SQL         As String
        Dim OK          As Boolean

10       On Error GoTo Client_UpdateDeviceRoomID_Error

20      Set Node = doc.selectSingleNode("HMC/roomid")
30      If Not Node Is Nothing Then
40        RoomID = Val(Node.text)
50        Set Node = doc.selectSingleNode("HMC/deviceid")
60        If Not Node Is Nothing Then
70          DeviceID = Val(Node.text)
80          On Error Resume Next
            
90          SQL = "UPDATE Devices SET RoomID = " & RoomID & " WHERE DeviceID = " & DeviceID
100         conn.BeginTrans
110         ConnExecute SQL
120         conn.CommitTrans

130         Devices.RefreshByID DeviceID
140         OK = (Err.Number = 0)

150       End If
160     End If


170     If OK Then
180       Client_UpdateDeviceRoomID = ReturnSuccess(Action, CStr(RoomID))

190     Else
200       Client_UpdateDeviceRoomID = ReturnFailure(Action)

210     End If

Client_UpdateDeviceRoomID_Resume:
220      On Error GoTo 0
230      Exit Function

Client_UpdateDeviceRoomID_Error:

240     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UpdateDeviceRoomID." & Erl
250     Resume Client_UpdateDeviceRoomID_Resume


End Function

Function RemoteClientFinalizeAlarm(ByVal Serial As String, ByVal Inputnum As Long, ByVal User As String, ByVal AlarmID As Long, ByVal Disposition As String)
  'RemoteClientFinalizeAlarm panel, Serial, Inputnum, User, AlarmID, Disposition
  
  ' called from Client_FinalizeAlarm(doc As DOMDocument60, Optional ByVal User As String = "", Optional ByVal Session As Long = 0) As String

  Dim d                  As cESDevice
  Set d = Devices.Devices(Serial)
  Dim alarm              As cAlarm

  For Each alarm In alarms.alarms
    If alarm.ID = AlarmID Then
      alarm.Disposition = Disposition
      PostEvent d, Nothing, alarm, EVT_ASSISTANCE_ACK, alarm.Inputnum, User
      Exit For
    End If
  Next

End Function




Function Client_FinalizeAlarm(doc As DOMDocument60, Optional ByVal User As String = "", Optional ByVal Session As Long = 0) As String

  Dim panel              As String
  Dim Serial             As String
  'Dim Inputs             As String
  Dim Inputnum           As Integer
  Dim myalarms           As Long
  Dim ConsoleID          As String
  Dim RemoteSerial       As String
  Dim AlarmID            As Long
  Dim PriorID            As Long

  Dim Root               As IXMLDOMNode
  Dim Node               As IXMLDOMNode
  Dim NodeList           As IXMLDOMNodeList  ' collection of nodes matching criteria
  Dim attr               As IXMLDOMAttribute
  Dim Disposition        As String

  On Error GoTo Client_FinalizeAlarm_Error

  

  'On Error GoTo Client_AckAlarm_Error

  Set Root = doc.selectSingleNode("HMC")

  If Not Root Is Nothing Then
    If Not Root.attributes Is Nothing Then
      For Each attr In Root.attributes
        If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
          ConsoleID = attr.text
        End If

        If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
          RemoteSerial = attr.text
        End If

      Next
    End If


    Set NodeList = Root.childnodes
    For Each Node In NodeList
      Select Case LCase(Node.baseName)

        Case "panel"
          panel = Node.text
        Case "serial"
          Serial = Node.text
        Case "inputnum"
          Inputnum = Val(Node.text)
        Case "alarmid"
          AlarmID = Val(Node.text)
        Case "priorid"
          PriorID = Val(Node.text)
        Case "myalarms"
          myalarms = Val(Node.text)
        Case "disposition"
          Disposition = Node.text
      End Select
    Next



    If Len(panel) > 0 And Len(Serial) > 0 Then
      'frmMain.ClientACKSelected panel, Serial, Val(Inputs), User, AlarmID
      'frmMain.RemoteClientFinalizeAlarm panel, Serial, Val(Inputs), User, AlarmID, Disposition
      
      Call RemoteClientFinalizeAlarm(Serial, Inputnum, User, AlarmID, Disposition)
    Else
      ' error?
    End If
    If (myalarms) Then
      Client_FinalizeAlarm = Client_GetSubscribedAlarms(ConsoleID, User, Session)
    Else
      Client_FinalizeAlarm = Client_GetAlarms(ConsoleID, User, Session)
    End If
  End If





Client_FinalizeAlarm_Resume:

  On Error GoTo 0
  Exit Function

Client_FinalizeAlarm_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_FinalizeAlarm." & Erl
  Resume Client_FinalizeAlarm_Resume

End Function

Function Client_RequestAssist(doc As DOMDocument60, Optional ByVal User As String = "", Optional ByVal Session As Long = 0) As String

  Dim panel              As String
  Dim Serial             As String
  Dim Inputs             As String
  Dim Inputnum           As Integer
  Dim myalarms           As Long
  Dim ConsoleID          As String
  Dim RemoteSerial       As String
  Dim AlarmID            As Long
  Dim PriorID            As Long

  Dim Root               As IXMLDOMNode
  Dim Node               As IXMLDOMNode
  Dim NodeList           As IXMLDOMNodeList  ' collection of nodes matching criteria
  Dim attr               As IXMLDOMAttribute

  On Error GoTo Client_RequestAssist_Error

  Set Root = doc.selectSingleNode("HMC")

  If Not Root Is Nothing Then
    If Not Root.attributes Is Nothing Then
      For Each attr In Root.attributes
        If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
          ConsoleID = attr.text
        End If

        If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
          RemoteSerial = attr.text
        End If

      Next
    End If


    Set NodeList = Root.childnodes
    For Each Node In NodeList
      Select Case LCase(Node.baseName)

        Case "panel"
          panel = Node.text
        Case "serial"
          Serial = Node.text
        Case "inputnum"
          Inputs = Node.text
        Case "alarmid"
          AlarmID = Val(Node.text)
        Case "priorid"
          PriorID = Val(Node.text)
        Case "myalarms"
          myalarms = Val(Node.text)
      End Select
    Next

    Debug.Print "modclients.Client_AckAlarm "

    If Len(panel) > 0 And Len(Serial) > 0 And Len(Inputs) > 0 Then
      'frmMain.ClientACKSelected panel, Serial, Val(Inputs), User, AlarmID
      frmMain.RemoteClientRequestAssist panel, Serial, Val(Inputs), User, AlarmID
      
    Else
      ' error?
    End If
    If (myalarms) Then
      Client_RequestAssist = Client_GetSubscribedAlarms(ConsoleID, User, Session)
    Else
      Client_RequestAssist = Client_GetAlarms(ConsoleID, User, Session)
    End If
  End If




Client_RequestAssist_Resume:

  On Error GoTo 0
  Exit Function

Client_RequestAssist_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_RequestAssist." & Erl
  Resume Client_RequestAssist_Resume

End Function

Function Client_AckAlarm(doc As DOMDocument60, ByVal User As String, ByVal Session As Long) As String
  Dim panel              As String
  Dim Serial             As String
  Dim Inputs             As String
  Dim Inputnum           As Integer
  Dim myalarms           As Long
  Dim ConsoleID          As String
  Dim RemoteSerial       As String
  Dim AlarmID            As Long
  Dim PriorID            As Long

  Dim Root               As IXMLDOMNode
  Dim Node               As IXMLDOMNode
  Dim NodeList           As IXMLDOMNodeList  ' collection of nodes matching criteria
  Dim attr               As IXMLDOMAttribute

  On Error GoTo Client_AckAlarm_Error

  Set Root = doc.selectSingleNode("HMC")

  If Not Root Is Nothing Then
    If Not Root.attributes Is Nothing Then
      For Each attr In Root.attributes
        If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
          ConsoleID = attr.text
        End If

        If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
          RemoteSerial = attr.text
        End If

      Next
    End If


    Set NodeList = Root.childnodes
    For Each Node In NodeList
      Select Case LCase(Node.baseName)

        Case "panel"
          panel = Node.text
        Case "serial"
          Serial = Node.text
        Case "inputnum"
          Inputs = Node.text
        Case "alarmid"
          AlarmID = Val(Node.text)
        Case "priorid"
          PriorID = Val(Node.text)
        Case "myalarms"
          myalarms = Val(Node.text)
      End Select
    Next

    Debug.Print "modclients.Client_AckAlarm "

    If Len(panel) > 0 And Len(Serial) > 0 And Len(Inputs) > 0 Then
      frmMain.ClientACKSelected panel, Serial, Val(Inputs), User, AlarmID
    Else
      ' error?
    End If
    If (myalarms) Then
      Client_AckAlarm = Client_GetSubscribedAlarms(ConsoleID, User, Session)
    Else
      Client_AckAlarm = Client_GetAlarms(ConsoleID, User, Session)
    End If
  End If
Client_AckAlarm_Resume:
  On Error GoTo 0
  Exit Function

Client_AckAlarm_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_AckAlarm." & Erl
  Resume Client_AckAlarm_Resume

End Function



Function Client_InvalidRequest(ByVal Request As String) As String

  Dim str As String '

  str = "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & "<Error>" & vbCrLf
  str = str & "Invalid Request " & Request & vbCrLf
  str = str & "</Error>" & vbCrLf
  str = str & "<ErrorCode>" & vbCrLf
  str = str & "404" & vbCrLf
  str = str & "</ErrorCode>" & vbCrLf
  
  str = str & "</HMC>"
  Client_InvalidRequest = str
End Function

Function Client_GetAlarms(ByVal ConsoleID As String, ByVal User As String, ByVal Session As Long) As String

  Dim str                As String  '
  On Error GoTo Client_GetAlarms_Error



  Dim alarm              As cAlarm
  Dim j                  As Integer

  ' user doesn't matter with this routine... just a read of alarms
  ' REPLICATE FOR SUBSCRIBED ALARMS


  'do for:
  'Alarms
  'Alerts
  'Troubles
  'LowBatts
  'Externs
  'Assurs

  'Creates an XML of all alarms, transport will create envelope
  'Alarms to be reconciled by Client
  'listview has varied data based upon type of alarm (Alarm, trouble, batt etc)
  'see frmmain.ProcessAlarms() for typcial data display

  ' ***************** ALARMS *********

  str = str & "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & taggit("response", "getalarms") & vbCrLf
  str = str & alarms.ToXML("Alarms", ConsoleID)
  str = str & Alerts.ToXML("Alerts", ConsoleID)
  str = str & Troubles.ToXML("Troubles", ConsoleID)
  str = str & LowBatts.ToXML("LowBatts", ConsoleID)
  str = str & Externs.ToXML("Externs", ConsoleID)
  str = str & Assurs.ToXML("Assurs", ConsoleID)

'  str = str & "<Alerts>" & vbCrLf
'  For j = 1 To Alerts.Count
'    Set alarm = Alerts.alarms(j)
'    str = str & "<Alarm>" & vbCrLf
'    str = str & alarm.ToXML(ConsoleID)
'    str = str & "</Alarm>" & vbCrLf
'  Next
'  str = str & taggit("Beep", Alerts.BeepTimer)
'  str = str & "</Alerts>" & vbCrLf

  
'
'  str = str & "<Troubles>" & vbCrLf
'  For j = 1 To Troubles.Count
'
'
'
'    Set alarm = Troubles.alarms(j)
'    str = str & "<Alarm>" & vbCrLf
'    str = str & alarm.ToXML(ConsoleID)
'    str = str & "</Alarm>" & vbCrLf
'
'    If j >= MAX_REMOTE_TROUBLES Then
'      Exit For                 ' prevent overload of data
'    End If
'
'  Next
'
'
'  str = str & taggit("Beep", Troubles.BeepTimer)
'  str = str & "</Troubles>" & vbCrLf
'
'  str = str & "<LowBatts>" & vbCrLf
'  For j = 1 To LowBatts.Count
'    Set alarm = LowBatts.alarms(j)
'    str = str & "<Alarm>" & vbCrLf
'    str = str & alarm.ToXML(ConsoleID)
'    str = str & "</Alarm>" & vbCrLf
'  Next
'  str = str & taggit("Beep", LowBatts.BeepTimer)
'  str = str & "</LowBatts>" & vbCrLf

'  str = str & "<Externs>" & vbCrLf
'  For j = 1 To Externs.Count
'    Set alarm = Externs.alarms(j)
'    str = str & "<Alarm>" & vbCrLf
'    str = str & alarm.ToXML(ConsoleID)
'    str = str & "</Alarm>" & vbCrLf
'  Next
'  str = str & taggit("Beep", Externs.BeepTimer)
'  str = str & "</Externs>" & vbCrLf

'  str = str & "<Assurs>" & vbCrLf
'  For j = 1 To Assurs.Count
'    Set alarm = Assurs.alarms(j)
'    str = str & "<Alarm>" & vbCrLf
'    str = str & alarm.AssurToXML
'    str = str & "</Alarm>" & vbCrLf
'  Next
'  str = str & taggit("Beep", Assurs.BeepTimer)
'  str = str & "</Assurs>" & vbCrLf

  str = str & taggit("HostTime", CDbl(Now))

  str = str & GetSessionStatus(Session)  ' get my session status

  str = str & "</HMC>"

  '  Debug.Print str

  Client_GetAlarms = str

Client_GetAlarms_Resume:
  On Error GoTo 0
  Exit Function

Client_GetAlarms_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_GetAlarms." & Erl
  Resume Client_GetAlarms_Resume


End Function


Function Client_GetSubscribedAlarms(ByVal ConsoleID As String, ByVal User As String, ByVal Session As Long) As String

        Dim str As String '
10      On Error GoTo Client_GetSubscribedAlarms_Error

20

        Dim alarm As cAlarm
        Dim j As Integer
        
        
        Static lastsession As Long
        
        If Session <> 0 Then
          lastsession = Session
'          dbg "Client_GetSubscribedAlarms  Last Now " & lastsession & " " & Session
        End If
        
        If lastsession <> 0 Then
          If Session = 0 Then
'            dbg "Client_GetSubscribedAlarms Last Now " & lastsession & " " & Session
            lastsession = 0
          End If
            
        End If
        

        ' user doesn't matter with this routine... just a read of alarms
        'do for:
        'Alarms
        'Alerts
        'Troubles
        'LowBatts
        'Externs
        'Assurs

        'Creates an XML of all alarms, transport will create envelope
        'Alarms to be reconciled by Client
        'listview has varied data based upon type of alarm (Alarm, trouble, batt etc)
        'see frmmain.ProcessAlarms() for typcial data display

        'makes calls into Alarm.ToXML

        Dim AlarmIDs() As Long


        Dim pageDevice As cPageDevice
        Dim listcount As Long
        Dim i As Long
30      ReDim AlarmIDs(0)

40      For Each pageDevice In gPageDevices
50        If pageDevice.ProtocolID = PROTOCOL_REMOTE Then
60          AlarmIDs = pageDevice.GetQueAlarmIDs(ConsoleID)  ' get list of alarms assigned to me
70        End If
80      Next

90      listcount = UBound(AlarmIDs)

'''************** NEED TO HANDLE REBEEP AS IN CLIENT_GETALARMS ********************************

100     str = "<?xml version=""1.0""?>" & vbCrLf
110     str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
120     str = str & taggit("response", "getalarms") & vbCrLf

130     str = str & "<Alarms>" & vbCrLf
140     For j = 1 To alarms.Count
150       Set alarm = alarms.alarms(j)
160       For i = 1 To listcount
            alarm.ID = alarm.ID ' needed to change from alarm.alarmID to  alarm.ID

170         If alarm.ID = AlarmIDs(i) Then
180           str = str & "<Alarm>" & vbCrLf
190           str = str & alarm.ToXML(ConsoleID)
200           str = str & "</Alarm>" & vbCrLf
210           Exit For
220         End If
230       Next
240     Next
250     str = str & taggit("Beep", alarms.BeepTimer)  ' beeptimer is either zero or non-zero for remotes
        str = str & taggit("SilenceTime", alarms.ConsoleSilenceTime(ConsoleID))   ' new for local beep control
        str = str & taggit("AlarmTime", alarms.ConsoleAlarmTime(ConsoleID))   ' new for local beep control
        
260     str = str & "</Alarms>" & vbCrLf

270     str = str & "<Alerts>" & vbCrLf
280     For j = 1 To Alerts.Count
290       Set alarm = Alerts.alarms(j)
300       For i = 1 To listcount
310         If alarm.ID = AlarmIDs(i) Then
320           str = str & "<Alarm>" & vbCrLf
330           str = str & alarm.ToXML(ConsoleID)
340           str = str & "</Alarm>" & vbCrLf
350           Exit For
360         End If
370       Next
380     Next
390     str = str & taggit("Beep", Alerts.BeepTimer)
400     str = str & "</Alerts>" & vbCrLf

410     str = str & "<Troubles>" & vbCrLf


420     For j = 1 To Troubles.Count
430       Set alarm = Troubles.alarms(j)
440       For i = 1 To listcount
450         If alarm.ID = AlarmIDs(i) Then
460           str = str & "<Alarm>" & vbCrLf
470           str = str & alarm.ToXML(ConsoleID)
480           str = str & "</Alarm>" & vbCrLf
490           Exit For
500         End If
510       Next
520       If j >= MAX_REMOTE_TROUBLES Then Exit For  ' prevent overload of data
530     Next
540     str = str & taggit("Beep", Troubles.BeepTimer)
550     str = str & "</Troubles>" & vbCrLf

560     str = str & "<LowBatts>" & vbCrLf
570     For j = 1 To LowBatts.Count
580       Set alarm = LowBatts.alarms(j)
590       For i = 1 To listcount
600         If alarm.ID = AlarmIDs(i) Then
610           str = str & "<Alarm>" & vbCrLf
620           str = str & alarm.ToXML(ConsoleID)
630           str = str & "</Alarm>" & vbCrLf
640           Exit For
650         End If
660       Next
670     Next
680     str = str & taggit("Beep", LowBatts.BeepTimer)
690     str = str & "</LowBatts>" & vbCrLf

700     str = str & "<Externs>" & vbCrLf
710     For j = 1 To Externs.Count
720       Set alarm = Externs.alarms(j)
730       For i = 1 To listcount
740         If alarm.ID = AlarmIDs(i) Then
750           str = str & "<Alarm>" & vbCrLf
760           str = str & alarm.ToXML(ConsoleID)
770           str = str & "</Alarm>" & vbCrLf
780           Exit For
790         End If
800       Next
810     Next
820     str = str & taggit("Beep", Externs.BeepTimer)
830     str = str & "</Externs>" & vbCrLf

840     str = str & "<Assurs>" & vbCrLf
850     For j = 1 To Assurs.Count
860       Set alarm = Assurs.alarms(j)
870       str = str & "<Alarm>" & vbCrLf
880       str = str & alarm.AssurToXML
890       str = str & "</Alarm>" & vbCrLf
900     Next
910     str = str & taggit("Beep", Assurs.BeepTimer)
920     str = str & "</Assurs>" & vbCrLf

       If Session Then
        UpdateSessionTime Session
       End If

        str = str & taggit("ts", Now)
        
        str = str & taggit("HostTime", CDbl(Now))

        Dim SessionStatus As String
        
        SessionStatus = GetSessionStatus(Session) ' get my session status

930     str = str & SessionStatus  ' get my session status

940     str = str & "</HMC>"

        dbg "Client Session,  SessionStatus " & Session & ", " & SessionStatus
        
          
        
960     Client_GetSubscribedAlarms = str


Client_GetSubscribedAlarms_Resume:
970     On Error GoTo 0
980     Exit Function

Client_GetSubscribedAlarms_Error:

990     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_GetSubscribedAlarms." & Erl
1000    Resume Client_GetSubscribedAlarms_Resume


End Function

Sub LogRemoteSession(ByVal SessionID As Long, ByVal SessionStatus As Long, ByVal Comment As String)
  On Error Resume Next
  Dim hfile As Long
  Dim filename As String
  filename = App.Path & "\Session.Log"
  limitFileSize filename
  hfile = FreeFile
  Open filename For Append As hfile
  Print #hfile, Now, " SessionID, SessionStatus ", SessionID, SessionStatus, Comment
  Close #hfile
End Sub

Public Function GetSessionStatus(ByVal SessionID As Long) As String   ' get my status
        Dim Session            As cUser
        Dim j                  As Integer
        Dim Found              As Boolean


10    'Debug.Print "****** Session Count " & HostSessions.Count
20      dbg "GetSessionStatus Session ID " & SessionID
30      On Error GoTo GetSessionStatus_Error

40      For j = HostSessions.Count To 1 Step -1

50        Set Session = HostSessions(j)
60        'Debug.Print "****** Session.Session, j "; Session.Session, j
70        If Session.Session = SessionID Then
80          Exit For
90        End If
100     Next


110     If j = 0 Then  ' I've been bumped off
120       dbg "Not Found"
130       If SessionID Then
140         LogRemoteSession SessionID, 0, "Session Not Found GetSessionStatus"
150         For j = HostSessions.Count To 1 Step -1
              Set Session = HostSessions(j)
160           LogRemoteSession Session.Session, 0, "Active Session Last Seen GetSessionStatus " & Session.LastSeen
170         Next
180       End If

190       GetSessionStatus = taggit("session", "0")
200     Else
210       dbg "OK"
220       GetSessionStatus = taggit("session", CStr(SessionID))
230     End If

GetSessionStatus_Resume:
240     On Error GoTo 0
250     Exit Function

GetSessionStatus_Error:

260     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.GetSessionStatus." & Erl
270     Resume GetSessionStatus_Resume


End Function

Function Client_SilenceAlarms(doc As DOMDocument60, ByVal User As String, ByVal Session As Long) As String

10      On Error GoTo Client_SilenceAlarms_Error

        Dim myalarms           As Long
        Dim ConsoleID          As String
        Dim Root               As IXMLDOMNode
        Dim Node               As IXMLDOMNode
        Dim NodeList           As IXMLDOMNodeList  ' collection of nodes matching criteria
        Dim attr               As IXMLDOMAttribute
        Dim RemoteSerial       As String
        Dim Alarmtype          As String


20      Set Root = doc.selectSingleNode("HMC")


30      If Not Root.attributes Is Nothing Then
40        For Each attr In Root.attributes
50          If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
60            ConsoleID = attr.text

70          End If
80          If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
90            RemoteSerial = attr.text
100         End If
110       Next
120     End If

130     Set NodeList = Root.childnodes
140     For Each Node In NodeList
150       Select Case LCase(Node.baseName)
            Case "alarmtype"
160           Alarmtype = Node.text
170         Case "myalarms"
180           myalarms = Val(Node.text)
190       End Select
200     Next

210     frmMain.SilenceAlarms User, ConsoleID, RemoteSerial, Alarmtype

220     If (myalarms) Then
230       Client_SilenceAlarms = Client_GetSubscribedAlarms(ConsoleID, User, Session)
240     Else
250       Client_SilenceAlarms = Client_GetAlarms(ConsoleID, User, Session)
260     End If
Client_SilenceAlarms_Resume:
270     On Error GoTo 0
280     Exit Function

Client_SilenceAlarms_Error:

290     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_SilenceAlarms." & Erl
300     Resume Client_SilenceAlarms_Resume


End Function


Function Client_UnSilenceAlarms(doc As DOMDocument60, ByVal User As String, ByVal Session As Long) As String

10      On Error GoTo Client_UnSilenceAlarms_Error

        Dim myalarms           As Long
        Dim ConsoleID          As String
        Dim Root               As IXMLDOMNode
        Dim Node               As IXMLDOMNode
        Dim NodeList           As IXMLDOMNodeList  ' collection of nodes matching criteria
        Dim attr               As IXMLDOMAttribute
        Dim RemoteSerial       As String
        Dim Alarmtype          As String


20      Set Root = doc.selectSingleNode("HMC")


30      If Not Root.attributes Is Nothing Then
40        For Each attr In Root.attributes
50          If 0 = StrComp(attr.baseName, "ConsoleID", vbTextCompare) Then
60            ConsoleID = attr.text

70          End If
80          If 0 = StrComp(attr.baseName, "RemoteSerial", vbTextCompare) Then
90            RemoteSerial = attr.text
100         End If
110       Next
120     End If

130     Set NodeList = Root.childnodes
140     For Each Node In NodeList
150       Select Case LCase(Node.baseName)
            Case "alarmtype"
160           Alarmtype = Node.text
170         Case "myalarms"
180           myalarms = Val(Node.text)
190       End Select
200     Next
210     frmMain.UnSilenceAlarms User, ConsoleID, RemoteSerial, Alarmtype

220     If (myalarms) Then
230       Client_UnSilenceAlarms = Client_GetSubscribedAlarms(ConsoleID, User, Session)
240     Else
250       Client_UnSilenceAlarms = Client_GetAlarms(ConsoleID, User, Session)
260     End If
Client_UnSilenceAlarms_Resume:
270     On Error GoTo 0
280     Exit Function

Client_UnSilenceAlarms_Error:

290     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modClients.Client_UnSilenceAlarms." & Erl
300     Resume Client_UnSilenceAlarms_Resume


End Function

