VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c6080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IP           As String
Public Username     As String
Public Password     As String
Public useSSL       As Boolean

Dim Messages        As Collection
Dim WebSocket       As WebSocketSocket

Public RequestString As String



Public DeviceName   As String  ' ACG, RX whatever
Public DeviceID     As String  ' long GUID
Public DeviceDescription As String
Public DeviceLocation As String
Public SystemContact As String
Public Model        As String
Public SerialNumber As String  ' actually a decimal number
Public MacAddress   As String
Public FirmwareVersion As String
Public BootVersion  As String
Public ManufactureDate As String  ' mm-dd-yyyy


Const EventStreamReconnect = "/PSIA/Metadata/stream?AreaControlEvents=true"
Const EventStream = "/PSIA/Metadata/stream?AreaControlEvents=true"
Const RegistrationStream = "/PSIA/Metadata/stream?Registration=true"
Function ParseACGData(ByVal XML As String)
  Dim doc           As DOMDocument60
  Dim Node          As IXMLDOMNode
  Dim Cnode         As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList
  Set doc = New DOMDocument60
  If doc.LoadXML(XML) Then
    Set Node = doc.selectSingleNode("DeviceInfo")
    If Not Node Is Nothing Then
      For Each Cnode In Node.childnodes
        Select Case Cnode.baseName
          Case "deviceName"
            DeviceName = Cnode.text
          Case "deviceID"
            DeviceID = Cnode.text

          Case "deviceDescription"
            DeviceDescription = Cnode.text
          Case "deviceLocation"
            DeviceLocation = Cnode.text
          Case "systemContact"
            SystemContact = Cnode.text
          Case "model"
            Model = Cnode.text
          Case "serialNumber"
            SerialNumber = Cnode.text
          Case "macAddress"
            MacAddress = Cnode.text
          Case "firmwareVersion"
            FirmwareVersion = Cnode.text
          Case "bootVersion"
            BootVersion = Cnode.text
          Case "manufactureDate"
            ManufactureDate = Cnode.text

          
      End Select
      Next
    End If


  End If



  '<DeviceInfo>
  '<deviceName>ACG</deviceName>
  '<deviceID> bfbabc10-4946-11e1-b86c-0800200c9a66 </DeviceID
  '<deviceDescription>EchoStream Gateway</deviceDescription>
  '<deviceLocation>South Campus</deviceLocation>
  '<systemContact>Joe,Smith,212 346-2323, 212 338-9930</systemContact>
  '<model>EN6080</model>
  '<serialNumber>12345678</serialNumber>
  '<macAddress> E00B28040506 </macAddress>
  '<firmwareVersion>0.59.10</firmwareVersion>
  '<bootVersion>0.57.1</bootVersion>
  '<manufactureDate>04-15-2012</manufactureDate>
  '</DeviceInfo>


End Function


Function Get6080Data()

  Dim XML           As String

  Dim HTTPRequest   As cHTTPRequest
  
  On Error Resume Next
  
  Set HTTPRequest = New cHTTPRequest
  XML = HTTPRequest.Get6080Data(GetHTTP & "://" & IP1, USER1, PW1)
  Set HTTPRequest = Nothing
  If XML = "" Then
    Get6080Data = False
  Else
    ParseACGData XML
    Get6080Data = True
  End If

  If InIDE Then
    Get6080Data = True
  End If


  '/PSIA/System/deviceInfo



End Function

'              192.168.1.122/PSIA/Metadata/stream?Registration=true
Public Property Get Is6080() As Long
  Is6080 = 1
End Property

Public Property Get Status() As Long
  On Error Resume Next
  If WebSocket Is Nothing Then
    Status = -1
  Else
    Status = WebSocket.StatusCode
  End If

End Property

Public Sub OverrideRequestString(ByVal NewRequestString As String)
  RequestString = NewRequestString
End Sub

Public Sub SetRequestString(ByVal ID As Long)
  Select Case ID
    Case 1
      RequestString = RegistrationStream
    Case Else
      RequestString = EventStream
  End Select
End Sub

Public Function ReConnect(ByVal TimeString As String) As Long

  Dim url           As String
  On Error Resume Next
  If WebSocket Is Nothing Then
    Set WebSocket = New WebSocketSocket
  Else
    WebSocket.DisConnect
  End If


  WebSocket.Init "5002EgAtIrEh"

  Call WebSocket.UserNamePassword(Username, Password)
  url = Get6080WSAddress & RequestString & "&time=2014-03-10T12:29:06.001Z"
  WebSocket.SetURL url
  WebSocket.Connect


End Function

Public Function Connect() As Long

  Dim url                As String
  Dim rc                 As Long
  Dim Address            As String
  On Error Resume Next
  If WebSocket Is Nothing Then
    Set WebSocket = New WebSocketSocket
  Else
    WebSocket.DisConnect
  End If

  Dim TimeString As String
  TimeString = ReadSetting("i6080", "DateStamp", "")
  
  


  WebSocket.Init "5002EgAtIrEh"

  Call WebSocket.UserNamePassword(Username, Password)
  Address = Get6080WSAddress
  'If Len(TimeString) Then
  
  'URL = Address & RequestString & "&time=" & TimeString
  'Else
  
  '4/15/2014 14:46:33.739 -400
  Dim LastTime As String
  
  LastTime = ReadSetting("i6080", "DateStamp", "")
  If Len(LastTime) Then
    LastTime = "&time=" & LastTime
  End If
  
  'LastTime = "&time=2014-04-15T14:46:33.739Z"
  
  url = Address & RequestString '& LastTime     ' & "&time=2014-04-15T14:46:33.739Z"
  
  ''URL = Address & RequestString & "&time=2014-03-10T12:29:06.001Z"
  
  'End If

  WebSocket.SetURL url
  rc = WebSocket.Connect()
  Sleep 200
  Dim Status As Long
  Status = WebSocket.StatusCode
'  If Status = 0 Then
'    RC = WebSocket.Connect()
'  End If
  Debug.Print "Status Code: "; WebSocket.StatusCode; " Last Error: "; WebSocket.GetLastError
'  If RC <> 1 Then
'    WebSocket.DisConnect
'    Sleep 200
'    Set WebSocket = Nothing
'    Set WebSocket = New WebSocketSocket
'
'    Debug.Print "Attempting ReConnect"
'
'    WebSocket.Init "5002EgAtIrEh"
'
'    Call WebSocket.UserNamePassword(UserName, Password)
'    Address = Get6080WSAddress
'    URL = Address & RequestString
'
'    WebSocket.SetURL URL
'    RC = WebSocket.Connect()
'    Status = WebSocket.StatusCode
'    Debug.Print "Return Code: "; RC; " Last Error: "; WebSocket.GetLastError
'
'  End If
  'If WebSocket.StatusCode = 0 Then
  '  MsgBox "NOT Connected in c6080.connect"
  'End If

  

End Function

Public Function HasMessages() As Long
  On Error Resume Next
  HasMessages = Messages.Count
End Function
Public Function GetNextMessage() As String
  Dim XML                As String
  Static firstmessagetime As Date
  Static messagecount    As Long
  Static Done            As Boolean
  Static ending          As Long

  Static Beginning       As Long
  If Beginning = 0 Then
    Beginning = Win32.timeGetTime
    ending = Win32.timeGetTime
  End If

  If CDbl(firstmessagetime = 0) Then
    firstmessagetime = Now
  End If


  On Error GoTo GetNextMessageError
  If Messages.Count Then
    messagecount = messagecount + 1
    If RequestString = RegistrationStream Then
      XML = Messages.Item(1) & ""
    Else
      XML = Messages.Item(1) & ""

    End If

    
    
    If InStr(1, XML, ">OK<", vbTextCompare) <= 0 Then
      Debug.Print Replace(XML, "><", ">" & vbCrLf & "<")
      
      
'      Debug.Assert 0
    End If
    
    
    If Not Done Then


      'If InStr(1, XML, "2014-04-15T20:17:55.228Z", vbTextCompare) Then
      If InStr(1, XML, "hello", vbTextCompare) Then
        Done = True
        Debug.Print messagecount & " Messages in " & Format$(ending - Beginning, "#,##0")  ' DateDiff("s", firstmessagetime, Now)
      Else
        ending = Win32.timeGetTime
      End If
    End If
    GetNextMessage = XML  'Messages.item(1) & ""
        
    Messages.Remove 1
  End If
  Exit Function

GetNextMessageError:

End Function




Public Function PushData(ByVal XML As String)

    Messages.Add XML

End Function

Public Function GetData()

10      On Error GoTo GetData_Error

20      If Not WebSocket Is Nothing Then
30        If WebSocket.StatusCode = 0 Then
            '  WebSocket.Connect
40        Else
50          Do While WebSocket.HasMessages

60            Messages.Add WebSocket.GetNextMessage
70          Loop
80        End If
90      End If


GetData_Resume:

100     On Error GoTo 0
110     Exit Function

GetData_Error:

120     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at c6080.GetData." & Erl
130     Resume GetData_Resume

End Function
Public Function ClearMessages() As Long
  GetData
  Set Messages = New Collection
End Function
Public Function DisConnect() As Long
  On Error Resume Next
  If Not WebSocket Is Nothing Then
    WebSocket.DisConnect
  End If
End Function

Private Sub Class_Initialize()
  Set Messages = New Collection
  SetRequestString 0
  IP = IP1
  Username = "Admin"
  Password = "Admin"

End Sub

Public Function Reset()
  On Error Resume Next

  If Not WebSocket Is Nothing Then
    WebSocket.DisConnect
    Set WebSocket = Nothing
  End If
  Set Messages = New Collection
  Set WebSocket = New WebSocketSocket
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  If Not WebSocket Is Nothing Then
    WebSocket.DisConnect
    Set WebSocket = Nothing
  End If

  Set Messages = Nothing
End Sub

Function ConvertToPacket(ByVal XML As String) As cESPacket
  
  ' Called from modMain.DoREAD / Packet is sent to ProcessESPacket -> Process6080Packet
  
  
  Dim NewPacket     As cESPacket
  Dim ThisEvent     As cAreaControlEvent
  ' converts 6080 XML data  to our packet object
  
  Set ThisEvent = New cAreaControlEvent
  Set NewPacket = New cESPacket
'
'  If InStr(1, XML, "base", vbTextCompare) < 1 Then
'
'
'  Debug.Print "c8060.ConvertToPacket XML" & vbCrLf & XML
'  End If

   'If InStr(1, XML, "EN6080 ACG", vbTextCompare) Then
    'Debug.Print XML
   'Else
    'Debug.Print XML
    'Debug.Assert 0
   
   'End If
  
  If ThisEvent.LoadXML(XML) Then
  
  ' If InStr(1, XML, "<SerialData>", vbTextCompare) Then

 ' End If

  NewPacket.TimeStamp = ThisEvent.EventTime
  

  If ThisEvent.SCICode <> 125 Then  ' InStr(1, XML, "base receiver", vbTextCompare) < 1 Then ' "hello" messages not stored in 6080 logs
    Call dbgPackets(XML)
    WriteSetting "i6080", "DateStamp", ThisEvent.TimeString
    WriteSetting "i6080", "ZoneType", ThisEvent.ZoneType
    WriteSetting "i6080", "SCICode", ThisEvent.SCICode
  End If
  ' typical Alarm Packet
  '<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.Zone/alarmState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>1061482</MetaSourceLocalID><MetaTime>2013-10-22T17:19:09.652Z</MetaTime><MetaPriority>1</MetaPriority></MetadataHeader><EventData><Info><ID>46</ID><Type>EN1210W Door/Window</Type><Description>B210326A</Description><PartitionList></PartitionList><SCI>ALARM 1</SCI><SCICode>1</SCICode></Info><ValueState><IntrusionAlarm>Intrusion</IntrusionAlarm></ValueState></EventData></AreaControlEvent>
  ' typical keepalive packet
  '<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.System/systemState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>7260310</MetaSourceLocalID><MetaTime>2013-10-22T17:18:07.732Z</MetaTime><MetaPriority>6</MetaPriority></MetadataHeader><EventData><Info><ID>1</ID><Type>EN6080 Receiver</Type><Description>base receiver</Description><PartitionList></PartitionList><SCI>hello</SCI><SCICode>125</SCICode></Info><ValueState><System>OK</System></ValueState></EventData></AreaControlEvent>
  
  
  
  
  

  Select Case left$(ThisEvent.ZoneType, 6)
    Case "EN6080", "EN6040"
      NewPacket.MID = &H0
    Case "ES1723"
      NewPacket.MID = &HC0
    Case "EN5000"
      NewPacket.MID = &H1
      NewPacket.PTI = 0
    Case "EN5040"
      NewPacket.MID = &H1
      NewPacket.PTI = &H1
    Case "EN5080", "EN5081"
      NewPacket.MID = &H1
    Case "EN1210"
      NewPacket.MID = &HB2
      'Debug.Assert 0
    Case Else
      NewPacket.MID = &HB2
  End Select

  NewPacket.Serial = Right$("00" & Hex$(NewPacket.MID), 2) & Right$("000000" & ThisEvent.HexID, 6)
  
  
  
'  If InStr(1, XML, "base receiver", vbTextCompare) < 1 Then
'    Debug.Print
'    Debug.Print XML
'    Debug.Print
'  End If
  
  NewPacket.LocalID = ThisEvent.LocalID
  Dim j             As Long
  
  If ThisEvent.Partitions.Count Then
    j = j
  End If
  'NewPacket.LocatedPartionName1 = "104"
  For j = 1 To ThisEvent.Partitions.Count
    Select Case j
      Case 1
        NewPacket.LocatedPartionName1 = ThisEvent.Partitions(j).Description
      Case 2
        NewPacket.LocatedPartionName2 = ThisEvent.Partitions(j).Description
      Case 3
        NewPacket.LocatedPartionName3 = ThisEvent.Partitions(j).Description

        'NewPacket.LocatedPartionID4 = ThisEvent.partitions(j).Description
    End Select
  Next

  'NewPacket.HexPacket = XML
  NewPacket.Is6080 = 1
  NewPacket.SCICode = ThisEvent.SCICode
  If ThisEvent.SCICode = SCI_CODE_SERIALDATA Then
    NewPacket.SerialDataValue = ThisEvent.SerialDataValue
  End If
  Set ConvertToPacket = NewPacket
  Set NewPacket = Nothing
  
  End If
End Function

'  Select Case ThisEvent.SCICode
'    Case SCI_CODE_ALARM1
'      NewPacket.Alarm0 = 1
'    Case SCI_CODE_ALARM1_CLEAR
'      NewPacket.Alarm0 = 0
'    Case SCI_CODE_ALARM2
'      NewPacket.Alarm1 = 1
'    Case SCI_CODE_ALARM2_CLEAR
'      NewPacket.Alarm1 = 0
'    Case SCI_CODE_ALARM3
'    Case SCI_CODE_ALARM3_CLEAR
'    Case SCI_CODE_ALARM4
'    Case SCI_CODE_ALARM4_CLEAR
'    Case SCI_CODE_DEVICE_INACTIVE
'      NewPacket.Missing = 1
'    Case SCI_CODE_DEVICE_INACTIVE_CLEARED
'      NewPacket.Missing = 0
'    Case SCI_CODE_TAMPER
'      NewPacket.Tamper = 1
'    Case SCI_CODE_TAMPER_CLEAR
'      NewPacket.Tamper = 0
'    Case SCI_CODE_EOL_TAMPER
'      NewPacket.Tamper = 1
'    Case SCI_CODE_EOL_TAMPER_CLEAR
'      NewPacket.Tamper = 0
'    Case SCI_CODE_LOW_BATT
'      NewPacket.Battery = 1
'    Case SCI_CODE_LOW_BATT_CLEAR
'      NewPacket.Battery = 0
'    Case SCI_CODE_MAINT_DUE
'      NewPacket.Trouble = 1
'    Case SCI_CODE_MAINT_DUE_CLEAR
'      NewPacket.Trouble = 0
'      'Case SCI_CODE_19 = 19
'      'Case SCI_CODE_20 = 20
'
'    Case SCI_CODE_DEVICE_RESET
'      NewPacket.Reset = 1
'    Case SCI_CODE_ENDPOINT_FAIL
'      NewPacket.Trouble = 1
'    Case SCI_CODE_ENDPOINT_SUCCESS
'      NewPacket.Trouble = 0
'    Case Else
'
'  End Select





