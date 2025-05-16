Attribute VB_Name = "mod6080"
Option Explicit

Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" (ByVal lDestIPAddr As Long, ByRef lHopCount As Long, ByVal lMaxHops As Long, ByRef lRTT As Long) As Long

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long


Global Const MIN_CHECKIN = 3 * 1 * 60 ' 3 x 3  s * 60 seconds per

Global GettingAllZoneStatus   As Boolean
Global LastAllZoneStatus      As Date


Global BackUp6080Handler As cHTTPRequest
Global BackUp6080Busy    As Boolean

Private mFirst6080packet      As Boolean

Function MakeResetPacket(ByVal Serial As String) As cESPacket
        Dim NewPacket          As cESPacket

        Dim Device             As cESDevice
10      Set Device = Devices.Device(Serial)

20      If Device Is Nothing Then
30        Exit Function
40      End If

50      Set NewPacket = New cESPacket
60      NewPacket.TimeStamp = Now
70      NewPacket.DateTime = Now
        'NewPacket.MID = "B2"
80      NewPacket.Serial = Device.Serial
90      NewPacket.Is6080 = 0
100     NewPacket.ClassByte = Device.CLS
110     NewPacket.CLSPTI = Device.CLSPTI
120     NewPacket.SCICode = SCI_CODE_DEVICE_RESET
130     NewPacket.Reset = 1
140     NewPacket.XML = ""
150     Set MakeResetPacket = NewPacket
160     Set NewPacket = Nothing


End Function


Function MakeFakePacket(ZoneInfo As cZoneInfo, ByVal SCICode As Long) As cESPacket
        Dim NewPacket     As cESPacket
        
10      Set NewPacket = New cESPacket
20      NewPacket.TimeStamp = Now
30      NewPacket.MID = ZoneInfo.MID
40      NewPacket.Serial = Right$("00" & Hex$(NewPacket.MID), 2) & Right$("000000" & ZoneInfo.HexID, 6)
50      NewPacket.LocalID = ZoneInfo.DeviceID ' serial as a decimal
60      NewPacket.Is6080 = 1
70      NewPacket.SCICode = SCICode
        
80      NewPacket.XML = ZoneInfo.XML
90      Set MakeFakePacket = NewPacket
        
100     Set NewPacket = Nothing

End Function



Public Function GetAllZoneStatus() As Boolean
        Dim HTTPRequest   As cHTTPRequest
        Dim http          As String
        Dim XML           As String
        Dim ZoneInfo      As cZoneInfo
        Dim SQL           As String
        Dim ZoneDevices   As Collection
        Dim Count         As Long
        Dim t             As Long
        Dim rc            As Long
        Dim Serial        As String



        Const Alarm1 = 1
        Const Alarm2 = 2
        Const Missing = 4

        Dim Rs            As ADODB.Recordset

10      If GettingAllZoneStatus Then
20        Exit Function
30      End If

40      t = Win32.timeGetTime
50      Count = 0
60      Set ZoneDevices = New Collection
70      GettingAllZoneStatus = True


80      SQL = " SELECT devices.Serial, devices.deviceid, devices.idm, devices.idl FROM devices WHERE deleted = 0 order by deviceid"
        '
90      Set Rs = ConnExecute(SQL)

100     Do Until Rs.EOF
          Dim BaseSerial  As String
110       BaseSerial = Right$(Rs("Serial" & ""), 6)
          On Error Resume Next
120       ZoneDevices.Add CStr(BaseSerial), CStr(BaseSerial)

130       Rs.MoveNext
140       DoEvents
150     Loop
160     Rs.Close
170     Set Rs = Nothing

180     Set ZoneInfoList = New cZoneInfoList

        Dim t2            As Long
190     t2 = Win32.timeGetTime

200     Set HTTPRequest = New cHTTPRequest
210     Call HTTPRequest.GetZoneList(GetHTTP & "://" & IP1, USER1, PW1)
220     Do Until HTTPRequest.Ready
230       DoEvents
240     Loop

250     Select Case HTTPRequest.StatusCode
          Case 200, 201
260         rc = 1
270       Case Else
280         rc = 0
290     End Select

300     Debug.Print "Time to Fetch Zonelist from 6080 " & Format(Win32.timeGetTime - t2, "0,000")

        '181 devices took 15.56 seconds or 90 ms / zone

310     If rc Then
320       If Len(HTTPRequest.XML) Then
330         rc = ZoneInfoList.LoadXML(HTTPRequest.XML)
340       End If
350     End If
360     Set HTTPRequest = Nothing

        Dim OKCount       As Long
        Dim notfound      As Long
        Dim SCICode       As Long

370     If rc Then
380       For Each ZoneInfo In ZoneInfoList.ZoneList
390         Serial = ""
400         On Error Resume Next
410         Serial = ZoneDevices.Item(CStr(Right$("000000" & ZoneInfo.HexID, 6)))
420         'Debug.Print "Zone serial", Val("&H" & ZoneInfo.HexID)
430         On Error GoTo 0
440         If Len(Serial) Then
              ' get status from zoneinfo

450           If GetZoneMissingStatus(ZoneInfo.ID) Then


460             SCICode = 0
470             Select Case ZoneInfo.MID
                  Case &HB2
480                 SCICode = SCI_CODE_DEVICE_INACTIVE
490               Case Else
500                 If InStr(1, ZoneInfo.TypeName, "6080") > 0 Then
510                   SCICode = 0
520                 Else
530                   SCICode = SCI_CODE_REPEATER_INACTIVE
540                 End If
550               End Select

560               If (SCICode <> 0) Then
570                 ProcessESPacket MakeFakePacket(ZoneInfo, SCICode)
                    
580                 Debug.Print "Missing ", ZoneInfo.DeviceID
590                 Count = Count + 1
                    ' create a missing packet
600               End If
610             Else
620               'Debug.Print "OK    ", ZoneInfo.DeviceID
630               OKCount = OKCount + 1
640             End If

650           Else
660             Debug.Print "Not Found ", ZoneInfo.DeviceID, ZoneInfo.HexID
670             notfound = notfound + 1
680           End If
690           DoEvents
700         Next
710       End If




          ' get all zones

          ' walk the zones v enrolled devices

          ' if missing, create fake event for the missing ONE AT A TIME

720       Debug.Print "Time to get all Missing Status " & Format(Win32.timeGetTime - t, "0,000")
730       Debug.Print "Missing Status Count =  " & Count
740       Debug.Print "OK   Count =  " & OKCount
750       Debug.Print "Not Found Status Count =  " & notfound
760       Debug.Print "ZoneInfoList.ZoneList.Count " & ZoneInfoList.ZoneList.Count

770       GettingAllZoneStatus = False


  End Function

Public Function CheckIf6080Alive()
        Dim Device6080 As cESDevice
        
10      Set Device6080 = Devices.Item(1)
        
        Dim SecondsOff As Long
        
        
        
20      SecondsOff = DateDiff("s", Device6080.LastSupervise, Now)
        'Debug.Print "Seconds Off "; SecondsOff
        
30      If Device6080.Dead = 0 Then
40        If SecondsOff > 60 Then
          
          ' need to reconnect
         ' Device6080.LastSupervise
        
50          First6080packet = False
            
60          Debug.Print "Missing at "; Now; "   "; Device6080.LastSupervise
70          PostEvent Device6080, Nothing, Nothing, EVT_COMM_TIMEOUT, 0
80          i6080.Connect
90          Device6080.LastSupervise = Now
100       End If
110     Else
120         First6080packet = False
130         If SecondsOff > 15 Then
140           Debug.Print "Attempt Reconnect "; Now
150           i6080.Connect
160           Device6080.LastSupervise = Now
170         End If
180     End If
        
End Function


Public Function UpdateCheckinTimeByModel(ByVal Model As String, ByVal CheckInTime As Long)


        Dim HTTPRequest   As cHTTPRequest
        
        Dim XML           As String
        Dim ZoneList      As Collection
        Dim Device        As cESDevice
        Dim ZoneID        As Long
        Dim rc            As Long
        Dim ZoneInfoList As cZoneInfoList
        Dim ZoneInfo      As cZoneInfo
        Dim MID           As Long
        Dim PTI           As Long
        
        
        
10      Set ZoneInfoList = New cZoneInfoList
        
20      If left$(Model, 2) = "ES" Then
30        Mid$(Model, 2, 1) = "N"
40      End If

50      Set HTTPRequest = New cHTTPRequest
60      Call HTTPRequest.GetZoneList(GetHTTP & "://" & IP1, USER1, PW1)
70      Do Until HTTPRequest.Ready
80        DoEvents
90      Loop
100     Select Case HTTPRequest.StatusCode
          Case 200, 201
110         rc = 1
120       Case Else
130         rc = 0
140     End Select
150     If rc Then
160       If Len(HTTPRequest.XML) Then
170         rc = ZoneInfoList.LoadXML(HTTPRequest.XML)
180       End If
190     End If
200     Set HTTPRequest = Nothing
        
210     If rc Then
220       For Each ZoneInfo In ZoneInfoList.ZoneList
230         If (left$(ZoneInfo.TypeName, Len(Model)) = Model) Then
240           Set HTTPRequest = New cHTTPRequest
250           XML = HTTPRequest.UpdateZoneCheckinTime(GetHTTP & "://" & IP1, USER1, PW1, ZoneInfo.ID, CheckInTime)
260           Set HTTPRequest = Nothing
270           DoEvents
280         End If
290       Next
300     End If
        
        
        
      '  ZoneInfoList.
      '  for each zone
      '
      '  http = "http"
      '  If UseSecureSockets Then
      '    http = "https"
      '  End If
      '
      '  Set ZoneList = New Collection
      '  ' get list of zones to change
      '  For Each Device In Devices.Devices
      '    If 0 = StrComp(Device.Model, Model, vbTextCompare) Then
      '      ZoneList.Add Device
      '    End If
      '  Next
      '
      '  For Each Device In ZoneList
      '    If Device.ZoneID Then
      '
      '      Set HTTPRequest = New cHTTPRequest
      '      XML = HTTPRequest.UpdateZoneCheckinTime(gethttp & "://" & IP1, USER1, PW1, Device.ZoneID, CheckInTime)
      '      Set HTTPRequest = Nothing
      '    End If
      '
      '  Next



End Function

Public Function UpdateZoneCheckinTimeByID(ByVal ZoneID As Long, ByVal CheckInTime As Long) As Boolean

        Dim HTTPRequest   As cHTTPRequest
        
        Dim XML           As String


        

10      If ZoneID Then
20        Set HTTPRequest = New cHTTPRequest
30        XML = HTTPRequest.UpdateZoneCheckinTime(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, CheckInTime)
40        Set HTTPRequest = Nothing
50      End If

End Function

Public Function RegisterSP(SoftPoint As cSoftPoint) As Long  ' maybe return New ID

        Dim HTTPRequest   As cHTTPRequest
        
        Dim XML           As String

        

10      Set HTTPRequest = New cHTTPRequest
20      XML = HTTPRequest.RegisterSoftPoint(GetHTTP & "://" & IP1, USER1, PW1, SoftPoint)

30      If Len(XML) Then
            Dim doc As DOMDocument60
            Dim Node As IXMLDOMNode
40          Set doc = New DOMDocument60
50          If doc.LoadXML(XML) Then
60            Set Node = doc.selectSingleNode("ResponseStatus/id")
70            If Not Node Is Nothing Then
80              SoftPoint.ID = Val(Node.text)
90              RegisterSP = SoftPoint.ID
100           End If
            
110         End If
120     Else

130     End If




End Function

Public Function UnregisterSP(ByVal ID As Long) As Boolean
  Dim HTTPRequest   As cHTTPRequest

  
  Dim rc As Boolean

  

  Set HTTPRequest = New cHTTPRequest
  UnregisterSP = HTTPRequest.UnRegisterSoftPoint(GetHTTP & "://" & IP1, USER1, PW1, ID)

  Set HTTPRequest = Nothing
  

End Function


Public Function GetCurrentPartionlist()


  Dim part          As cPartition
  Dim li            As ListItem
  Dim Found         As Boolean
  Dim XML           As String


  Dim HTTPRequest   As cHTTPRequest


  Set HTTPRequest = New cHTTPRequest
  XML = HTTPRequest.GetPartitionList(GetHTTP & "://" & IP1, USER1, PW1)

  If Len(XML) Then
    Set Partitions = ParsePartionList(XML)
  Else
    Set Partitions = New Collection
  End If

End Function
Public Function MakeEnCryptedString(ByVal PlainText As String) As String
  Dim s             As String
  s = SimpleEncrypt(PlainText, "Boogeddy")
  MakeEnCryptedString = SimpleHexit(s)
End Function
Public Function MakeDeCryptedString(ByVal CypherText As String) As String
  Dim s             As String
  s = SimpleDeHexit(CypherText)
  MakeDeCryptedString = SimpleEncrypt(s, "Boogeddy")

End Function


Private Function SimpleEncrypt(ByVal inputString As String, ByVal Password As String) As String
  Dim L             As Integer
  Dim x             As Integer
  Dim Char          As String

  L = Len(Password)
  For x = 1 To Len(inputString)
    Char = Asc(MID$(Password, (x Mod L) - L * ((x Mod L) = 0), 1))
    Mid$(inputString, x, 1) = Chr$(Asc(MID$(inputString, x, 1)) Xor Char)
  Next
  SimpleEncrypt = inputString
End Function
Private Function SimpleHexit(ByVal s As String)
  Dim j             As Long
  Dim outs          As String
  For j = 1 To Len(s)
    outs = outs & Right("00" & Hex(Asc(MID(s, j, 1))), 2)
  Next
  SimpleHexit = outs
End Function
Private Function SimpleDeHexit(ByVal s As String)
  Dim j             As Long
  Dim outs          As String
  For j = 1 To Len(s) Step 2
    outs = outs & Chr(Val("&h" & MID(s, j, 2)))
  Next
  SimpleDeHexit = outs
End Function



Function Set6080NID(ByVal NewNID As Long) As Long

  Dim HTTPRequest   As cHTTPRequest
  Set HTTPRequest = New cHTTPRequest


  Set6080NID = HTTPRequest.SetNID(GetHTTP & "://" & IP1, USER1, PW1, NewNID)
  Set HTTPRequest = Nothing



  '<RFNetwork><NID>10</NID></RFNetwork>
End Function

Function Get6080NID() As Long
  Dim HTTPRequest   As cHTTPRequest
  Set HTTPRequest = New cHTTPRequest


  Get6080NID = HTTPRequest.GetNID(GetHTTP & "://" & IP1, USER1, PW1)
  Set HTTPRequest = Nothing

End Function

Function GetZoneMissingStatus(ByVal ZoneID As Long) As Long

        ' returns 0 (false) or 1 (true)
        
        
        Dim rc            As Long
        Dim ZoneInfo      As cZoneInfo
        Dim doc           As DOMDocument60

        

        Dim HTTPRequest   As cHTTPRequest
10      Set HTTPRequest = New cHTTPRequest


20      Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, ZoneID)
30      If Not (ZoneInfo Is Nothing) Then
40        If 0 = StrComp(ZoneInfo.IsMissing, "true", vbTextCompare) Then
50          GetZoneMissingStatus = 1
60        Else
70          GetZoneMissingStatus = 0
80        End If
90      Else
100       GetZoneMissingStatus = 0
110     End If


End Function

Function GetZoneIDL(ByVal ZoneID As Long) As Long
        
        Dim rc            As Long
        Dim ZoneInfo      As cZoneInfo
        Dim doc           As DOMDocument60

        

        Dim HTTPRequest   As cHTTPRequest
10      Set HTTPRequest = New cHTTPRequest

        

20      Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, ZoneID)
30      If Not (ZoneInfo Is Nothing) Then
40        If ZoneInfo.IsFixedDevice Then
50          GetZoneIDL = 2
60        ElseIf ZoneInfo.IsLocatable Then
70          GetZoneIDL = 1
80        ElseIf ZoneInfo.IsSoftPointer Then
90          GetZoneIDL = 3
100       Else
110         GetZoneIDL = 0
120       End If
130     Else
140       GetZoneIDL = 0
150     End If




End Function
Function ChangeZoneParameter(ByVal ZoneID As Long, ByVal ParamName As String, ByVal Value As String) As Long

        
        Dim rc            As Long
        

        Dim HTTPRequest   As cHTTPRequest
10      Set HTTPRequest = New cHTTPRequest

20      Select Case ParamName   ' for these three, we need to set other conflicting params to false
          Case "IsRef"
30          If Value = "true" Then
40            rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "Locatable", "false")
50            rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "IsSPDevice", "false")
60            rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, ParamName, Value)
70          End If
80        Case "Locatable"
90          If Value = "true" Then
100           rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "IsRef", "false")
110           rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "IsSPDevice", "false")
120           rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, ParamName, Value)
130         End If

140       Case "IsSPDevice"
150         If Value = "true" Then
160           rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "IsRef", "false")
170           rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "Locatable", "false")
180           rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, ParamName, Value)
190         End If
200       Case "Null"
210         rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "Locatable", "false")
220         rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "IsRef", "false")
230         rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, "IsSPDevice", "false")

240       Case Else
250         rc = HTTPRequest.ChangeZoneParameter(GetHTTP & "://" & IP1, USER1, PW1, ZoneID, ParamName, Value)
260     End Select




270     Set HTTPRequest = Nothing




End Function

Function Reboot6080()
  'PUT http://192.168.1.122/PSIA/System/reboot
End Function


Function Get6080WSAddress() As String
  Dim protocol      As String
  protocol = "ws"
  If UseSecureSockets Then
    protocol = "wss"
  End If
  Get6080WSAddress = protocol & "://" & IP1
End Function


Function UpgradeDevice(Device As cSimpleDevice) As Long
        Dim ZoneID        As Long
        Dim XML           As String
          ' we need to register
        

10      XML = "<ZoneInfo><DeviceID>" & Device.DecimalSerial & "</DeviceID><PTI>" & Device.PTI & "</PTI>" & _
              "<MID>" & Device.MID & "</MID><Locatable>" & IIf(Device.IDL = 1, "true", "false") & "</Locatable>" & _
              "<Description>" & XMLEncode(Device.Serial) & "</Description><IsRef>" & IIf(Device.IDL = 2, "true", "false") & "</IsRef>" & _
              "<SyncWindow>0</SyncWindow><SyncTimeout>0</SyncTimeout><MessageExpirationTime>0</MessageExpirationTime><CheckInTime>0</CheckInTime>" & _
              "<SupervisionWindow>" & Device.Checkin6080 & "</SupervisionWindow><IsSPDevice>" & IIf(Device.IDL = 3, "true", "false") & "</IsSPDevice></ZoneInfo>"




        Dim HTTPRequest   As cHTTPRequest
20      Set HTTPRequest = New cHTTPRequest
30      XML = HTTPRequest.RegisterDevice(GetHTTP & "://" & IP1, USER1, PW1, XML)

40      Set HTTPRequest = Nothing

50      If InStr(1, XML, "Device already registered", vbTextCompare) Then
60        UpgradeDevice = 0
          'See if serial is in our zoneinfolist
          'scan it and return ID if found
          'if not, get a Fresh zoneinfolist and check it.
          'UpgradeDevice = ScanZoneInfoListForSerial(Device.Serial)' not for upgrading system, takes too long
70      Else
          ' get Zone ID from XML
80        UpgradeDevice = ZoneFromXML(XML)
90      End If
End Function
Function GetACGSerial() As String

  GetACGSerial = ""
End Function

Function RegisterDevice(Device As cESDevice) As Long
        Dim ZoneID        As Long

10      ZoneID = Device.ZoneID
20      If ZoneID = 0 Then
          ' we need to register
          Dim XML         As String
          Dim http        As String


          Dim PTI         As Long
30        PTI = Device.PTI


40        XML = "<ZoneInfo><DeviceID>" & Device.DecimalSerial & "</DeviceID><PTI>" & Device.PTI & "</PTI>" & _
                "<MID>" & Device.MID & "</MID><Locatable>" & IIf(Device.IsPortable, "true", "false") & "</Locatable>" & _
                "<Description>" & XMLEncode(Device.Serial) & "</Description><IsRef>" & IIf(Device.IsRef, "true", "false") & "</IsRef>" & _
                "<SyncWindow>0</SyncWindow><SyncTimeout>0</SyncTimeout><MessageExpirationTime>0</MessageExpirationTime><CheckInTime>0</CheckInTime>" & _
                "<SupervisionWindow>" & Device.Checkin6080 & "</SupervisionWindow><IsSPDevice>" & IIf(Device.IsSPDevice, "true", "false") & "</IsSPDevice></ZoneInfo>"


50        http = GetHTTP

          Dim HTTPRequest As cHTTPRequest
60        Set HTTPRequest = New cHTTPRequest
70        XML = HTTPRequest.RegisterDevice(GetHTTP & "://" & IP1, USER1, PW1, XML)

80        Set HTTPRequest = Nothing

90        If InStr(1, XML, "Device already registered", vbTextCompare) Then
100         RegisterDevice = 0
            'See if serial is in our zoneinfolist
            'scan it and return ID if found
            'if not, get a Fresh zoneinfolist and check it.
            
            
            
110         RegisterDevice = ScanZoneInfoListForSerial(Device.Serial)
            
120       Else
            ' get Zone ID from XML
130         RegisterDevice = ZoneFromXML(XML)
140       End If
150     Else
          ' maybe later verify it
160       RegisterDevice = ZoneID
170     End If
End Function
Function ScanZoneInfoListForSerial(ByVal Serial As String)
  Dim ZoneInfoList  As cZoneInfoList
  Dim ZoneID        As Long


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
  If rc Then
    ZoneID = ZoneInfoList.ScanforSerial(Serial)
  End If
  ScanZoneInfoListForSerial = ZoneID

End Function

Function ZoneFromXML(ByVal XML As String) As Long
  Dim doc           As DOMDocument60
  Dim Node          As IXMLDOMNode
  Set doc = New DOMDocument60
  If doc.LoadXML(XML) Then
    Set Node = doc.selectSingleNode("ResponseStatus/id")
    If Not Node Is Nothing Then
      ZoneFromXML = Val(Node.text)
    End If
  End If

End Function
Function ParseSoftPointList(ByVal XML As String) As Collection

  Dim doc           As DOMDocument60

  Dim Node          As IXMLDOMNode
  Dim subnode       As IXMLDOMNode
  Dim ssnode        As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList

  Dim SoftPoint     As cSoftPoint
  Dim SoftPoints    As Collection

  Set SoftPoints = New Collection


  Set doc = New DOMDocument60
  If doc.LoadXML(XML) Then

    Set Node = doc.selectSingleNode("SoftPointList")
    If Not Node Is Nothing Then
      For Each subnode In Node.childnodes
        If subnode.baseName = "SoftPoint" Then
          Set SoftPoint = New cSoftPoint
          If SoftPoint.ParseXML(subnode.XML) Then
            SoftPoints.Add SoftPoint
          End If
        End If
      Next
    End If
  Else
  
    LogXML XML, "Error-SoftPointList"
  End If

  Set ParseSoftPointList = SoftPoints
  Set SoftPoints = Nothing



End Function


Function ParsePartionList(ByVal XML As String) As Collection


  '<PartitionInfoList>
  ' <PartitionInfo>
  '   <ID>1</ID>
  '   <Description>Test Partition 1</Description >
  '   <IsMobile>false</IsMobile>
  ' </PartitionInfo>
  ' <PartitionInfo>
  '   <ID>2</ID>
  '   <Description>Test Partition 2</Description >
  '   <IsMobile>false</IsMobile>
  ' </PartitionInfo>
  '</PartitionInfoList>


  Dim doc           As DOMDocument60
  Dim list          As Collection
  Dim Node          As IXMLDOMNode
  Dim subnode       As IXMLDOMNode
  Dim ssnode        As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList
  Dim ZoneInfo      As cZoneInfo
  Dim partition     As cPartition

  Set list = New Collection
  Set doc = New DOMDocument60
  
  
  
  If doc.LoadXML(XML) Then

    Set list = New Collection
    Set Node = doc.selectSingleNode("PartitionInfoList")
    If Not Node Is Nothing Then
      For Each subnode In Node.childnodes
        If subnode.baseName = "PartitionInfo" Then
          Set partition = New cPartition

          For Each ssnode In subnode.childnodes
            Select Case ssnode.baseName
              Case "ID"

                partition.PartitionID = Val(ssnode.text)
                list.Add partition, ssnode.text & ""
              Case "Description"
                partition.Description = XMLDecode(Trim$(ssnode.text))
              Case "IsMobile"
                partition.IsLocation = (CBool(ssnode.text)) And 1
            End Select
          Next
        End If
      Next
    End If
  Else
    LogXML XML, "Error-PartitionList"
    
  End If

  Set ParsePartionList = list
  Set list = Nothing

End Function
Public Sub LogXML(ByVal XML As String, ByVal SourceOfXML As String)
    Dim hfile As Long
    Dim filename As String
    On Error Resume Next
    filename = App.Path & "\" & SourceOfXML & ".xml"
    hfile = FreeFile
    Kill filename
    Open filename For Output As hfile
    Print #hfile, XML
    Close hfile

End Sub

Public Function Remove6080Device(ByVal ZoneID As Long)

  Dim HTTPRequest As cHTTPRequest

  If ZoneID <> 0 Then
    Set HTTPRequest = New cHTTPRequest
    Remove6080Device = HTTPRequest.UnRegisterDevice(GetHTTP & "://" & IP1, USER1, PW1, ZoneID)
    Set HTTPRequest = Nothing
  End If

End Function

Public Function Process6080Packet(packet As cESPacket) As Long

        Dim d                  As cESDevice
        Dim a                  As cAlarm
        'comes from ProcessESPacket in modES
        Dim IsAlarm            As Boolean
        Dim IsAlarm_A          As Boolean
        Dim IsAlarm_B          As Boolean
        Dim SerialDataValue    As Long
        Dim locationtext       As String
        Dim roomname           As String


10      frmMain.PacketToggle

20      Set d = Devices.Item(1)
        ' we'll want to keep track of this for replay when reboot

        'Debug.Print "Packet at "; Now
        'Debug.Print "Process6080Packet Device 1 " & d.Serial, packet.Serial
30      d.LastSupervise = Now        ' set time for receiver com port activity
        'Debug.Print "Packet at "; d.LastSupervise
        'dbg "ProcessESPacket  " & Packet.serial
40      If d.Dead = 1 Then
50        dbgPackets "6080 Back on line"
60        PostEvent d, packet, Nothing, EVT_COMM_RESTORE, 0
70      End If

80      If gShowPacketData Then
          'Call dbgPackets(packet.XML)
90      End If

100     If packet.Serial = "B2851BB5" Then
          'Debug.Assert 0
110     End If

120     Set d = Devices.Device(packet.Serial)
130     If d Is Nothing Then
140       Process6080Packet = -1     ' stray/failure
150       Exit Function
160     Else
170       Process6080Packet = 0      ' not failure
180     End If


        If Len(d.Configurationstring) > 0 Then
          Dim resetpacket As cESPacket
          Select Case packet.SCICode
            Case SCI_CODE_ALARM1, SCI_CODE_ALARM2, SCI_CODE_ALARM3, SCI_CODE_ALARM4
              Set resetpacket = MakeResetPacket(d.Configurationstring)
          End Select
          If resetpacket Is Nothing Then
            ' do nothing
          Else
            Debug.Print "mod6080.ProcessESPacket resetpacket"
            ProcessESPacket resetpacket
            
          End If
          Exit Function
        End If

        ' try to get best match on location 8/15/14

190     If d.Ignored Then
200       If d.Alarm_A And packet.SCICode = SCI_CODE_ALARM1_CLEAR Then
            ' LET IT PASS
210       ElseIf d.Alarm_B And packet.SCICode = SCI_CODE_ALARM2_CLEAR Then
            ' LET IT PASS
220       Else
            ' EAT IT
230         Exit Function            ' maybe 10/30/2013
240       End If
250     End If

260     If d.IsPortable Then

270       Select Case packet.SCICode
            Case SCI_CODE_ALARM1, SCI_CODE_ALARM2, SCI_CODE_ALARM3, SCI_CODE_ALARM4

280           d.FetchRoom
290           roomname = Trim$(d.Room)
300           If Len(roomname) Then

                'Call dbgPackets("We Want " & roomname)
                'Call dbgPackets("Partition Name 1 " & packet.LocatedPartionName1)
                'Call dbgPackets("Partition Name 2 " & packet.LocatedPartionName2)

310             d.FetchRoom

320             If (Len(packet.LocatedPartionName1) > 0) And (0 = StrComp(roomname, packet.LocatedPartionName1, vbTextCompare)) Then
330               d.LastLocationText = roomname
340               packet.LocatedPartion = roomname
                  'Call dbgPackets("Exact Match 1 " & roomname)

350             ElseIf (Len(packet.LocatedPartionName2) > 0) And (0 = StrComp(roomname, packet.LocatedPartionName2, vbTextCompare)) Then
360               d.LastLocationText = roomname
370               packet.LocatedPartion = roomname
                  'Call dbgPackets("Exact Match 2 " & roomname)

380             Else                 ' no direct match, look in keywords
390               If d.IsInLocKW(packet.LocatedPartionName1, d.locKW) Then  ' match on first partition
400                 d.LastLocationText = roomname
410                 packet.LocatedPartion = roomname
                    'Call dbgPackets("KW Match 1 " & roomname)

420               Else
430                 If d.IsInLocKW(packet.LocatedPartionName2, d.locKW) Then  ' match on second partition
440                   d.LastLocationText = roomname
450                   packet.LocatedPartion = roomname
                      'Call dbgPackets("KW Match 2 " & roomname)
460                 Else             ' no match at all
470                   If Len(packet.LocatedPartionName1) > 0 Then
480                     d.LastLocationText = packet.LocatedPartionName1
490                     packet.LocatedPartion = packet.LocatedPartionName1
                        'Call dbgPackets("No Match Using Partition Name 1 " & packet.LocatedPartionName1)
500                   End If
510                 End If
520               End If
530             End If

540           Else                   ' no room assigned, use keywords for best match

                'dbgPackets ("no room assigned, use keywords for best match")

550             locationtext = d.GetLocKW(packet.LocatedPartionName1, d.locKW)
560             If Len(locationtext) Then  ' use partition 1 as the most likely location
570               d.LastLocationText = locationtext
580               packet.LocatedPartion = locationtext
                  'dbgPackets ("no room assigned, KW 1 Match " & locationtext)
590             Else
600               locationtext = d.GetLocKW(packet.LocatedPartionName2, d.locKW)
610               If Len(locationtext) Then  ' use partition 2 as the most likely location
620                 d.LastLocationText = locationtext
630                 packet.LocatedPartion = locationtext
                    'dbgPackets ("no room assigned, KW 2 Match " & locationtext)
640               Else               ' default to whatever is in the first partition
650                 d.LastLocationText = packet.LocatedPartionName1
660                 packet.LocatedPartion = packet.LocatedPartionName1
                    'dbgPackets ("no room assigned, NO Match " & d.LastLocationText)
670               End If
680             End If
690           End If
700       End Select                 ' only alarm codes
710     End If                       ' only portables (processing location data)

        'dbgPackets ("Location for Device " & d.LastLocationText)
        'dbgPackets ("Location for Packet " & packet.LocatedPartion)

720     Select Case packet.SCICode

          Case SCI_CODE_SERIALDATA
730         SerialDataValue = packet.SerialDataValue

740         If ((SerialDataValue And CODE_1941XS_TAMPER) = 0) And ((SerialDataValue And CODE_1941XS_RESET) = CODE_1941XS_RESET) Then
750           If d.Tamper Then
760             PostEvent d, packet, Nothing, EVT_TAMPER_RESTORE, 0
770           End If
780         End If

790         If (SerialDataValue And CODE_1941XS_TAMPER) = CODE_1941XS_TAMPER Then
800           If d.Tamper = 0 Then
810             PostEvent d, packet, Nothing, EVT_TAMPER, 0
820           End If
830         End If

840         If (SerialDataValue And CODE_1941XS_RESET) = CODE_1941XS_RESET Then
              'reset sent
850         End If

860         If (SerialDataValue And CODE_1941XS_PULLCORD) = CODE_1941XS_PULLCORD Then
              ' Pull Cord/Lever Active
870           IsAlarm = True
880           If (d.isDisabled) Then
890             Exit Function        ' maybe
900           End If
910           If d.IsAway = 0 Then   ' if it's not on vacation then
920             If d.AssurInput = 1 Then
930               If d.Assur = 1 Then  ' an assurance device
940                 IsAlarm = False
950               End If
960             End If
970           Else                   ' if it's on vacation
980             If d.AssurInput = 1 Then
990               If d.Assur = 1 Then  ' it's an assurance device
1000                If d.AssurSecure = 0 Then  ' they took it home
1010                  IsAlarm = False
1020                End If
1030              End If
1040            End If
1050          End If
1060          If d.AssurInput = 1 Then
1070            If d.AssurBit = 1 Then  ' is it set ?
1080              d.AssurBit = 0     ' clear it
1090              PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 1
1100              IsAlarm = False
1110            End If
1120          End If
1130          If IsAlarm Then
                '1040          If (d.IsAway = 0) Then
1140            If (1) Then
1150              If d.alarm = 0 Then
1160                d.alarm = 1
1170                Select Case d.AlarmMask
                      Case 2
1180                    PostEvent d, packet, Nothing, EVT_EXTERN, 1
1190                  Case 1
1200                    PostEvent d, packet, Nothing, EVT_ALERT, 1
1210                  Case Else
1220                    Trace " EVT_EMERGENCY 1"
1230                    PostEvent d, packet, Nothing, EVT_EMERGENCY, 1
1240                End Select
1250              End If
1260            End If
1270          End If
1280        End If
            
1290        If ((SerialDataValue And CODE_1941XS_PUSHBUTTON) = CODE_1941XS_PUSHBUTTON) Or ((SerialDataValue And CODE_1941XS_PUSHBUTTON_DOWN) = CODE_1941XS_PUSHBUTTON_DOWN) Then
              'pusbutton alarm
1300          IsAlarm_A = True
1310          If (d.isDisabled_A) Then
1320            Exit Function        ' maybe
1330          End If
1340          If d.IsAway = 0 Then   ' if it's not on vacation then
1350            If d.AssurInput = 2 Then
1360              If d.Assur = 1 Then  ' an assurance device
1370                IsAlarm_A = False
1380              End If
1390            End If
1400          Else                   ' if it's on vacation
1410            If d.AssurInput = 2 Then
1420              If d.Assur = 1 Then  ' it's an assurance device
1430                If d.AssurSecure_A = 0 Then  ' they took it home
1440                  IsAlarm_A = False
1450                End If
1460              End If
1470            End If
1480          End If
1490          If d.AssurInput = 2 Then
1500            If d.AssurBit = 1 Then  ' is it set ?
1510              d.AssurBit = 0     ' clear it
1520              PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 2
1530              IsAlarm_A = False
1540            End If
1550          End If
1560          If IsAlarm_A Then
                '1040          If (d.IsAway = 0) Then
1570            If (1) Then
1580              If d.Alarm_A = 0 Then
1590                d.Alarm_A = 1
1600                Select Case d.AlarmMask_A
                      Case 2
1610                    PostEvent d, packet, Nothing, EVT_EXTERN, 2
1620                  Case 1
1630                    PostEvent d, packet, Nothing, EVT_ALERT, 2
1640                  Case Else
1650                    Trace " EVT_EMERGENCY 2"
1660                    PostEvent d, packet, Nothing, EVT_EMERGENCY, 2
1670                End Select
1680              End If
1690            End If
1700          End If
1710        End If


1720        If ((SerialDataValue And &H40) = 0) And ((SerialDataValue And &H4) = 4) Then  ' pullcord restored
              
              'If d.ClearByReset Then
              ' defer to reset
              'Else
              
1730          If d.alarm Then
1740            d.alarm = 0
1750            If d.AlarmMask = 1 Then
1760              PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 1
1770            Else
1780              Trace " EVT_EMERGENCY_RESTORE 1"
1790              PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 1
1800            End If
1810          End If
1820        End If


1830        If ((SerialDataValue And CODE_1941XS_PUSHBUTTON) = CODE_1941XS_OK) And ((SerialDataValue And CODE_1941XS_RESET) = CODE_1941XS_RESET) Then  ' pushbutton restored
              
              'If d.ClearByReset Then
              ' defer to reset
              'Else

1840          If d.Alarm_A Then
1850            d.Alarm_A = 0
1860            If d.AlarmMask_A = 1 Then
1870              PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 2
1880            Else
1890              Trace " EVT_EMERGENCY_RESTORE 2"
1900              PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 2
1910            End If
1920          End If
1930        End If

            ' handle serial data here
1940        Select Case SerialDataValue
              Case &H0               ' Pushbutton (tamper) restored, no alarms
1950          Case &H4               ' reset"
1960          Case &H8               ' pushbutton unplugged (tamper)
1970          Case &HC               ' Reset w/ pushbutton unplugged (tamper)
1980          Case &H20              ' pushbutton
1990          Case &H40              ' pullcord
2000          Case &H44              ' reset w/ pullcord active
2010          Case &H48              ' pushbutton unplugged w/ pullcord active
2020          Case &H60              ' pull cord and pushbutton active
                'SerialValue = SCI_CODE_ALARM1_AND_ALARM2
2030          Case Else
                ' UNDEFINED
2040        End Select


2050      Case SCI_CODE_ALARM1
2060        IsAlarm = True

2070        If d.isDisabled Then
2080          Exit Function          ' maybe
2090        End If

2100        If d.IsAway = 0 Then     ' if it's not on vacation then
2110          If d.AssurInput < 1 Then  ' just a sanity check
2120            If d.Assur = 1 Then
2130              d.AssurInput = 1
2140            End If
2150          End If
2160        End If

2170        If d.IsAway = 0 Then     ' if it's not on vacation then
2180          If d.AssurInput = 1 Then  ' a single input device that somehow didn't get the input # flagged
2190            If d.Assur = 1 Then  ' an assurance device
2200              IsAlarm = False
2210            End If
2220          End If
2230        Else                     ' if it's on vacation
2240          If d.AssurInput = 1 Then
2250            If d.Assur = 1 Then  ' it's an assurance device
2260              If d.AssurSecure = 0 Then  ' they took it home
2270                IsAlarm = False
2280              End If
2290            End If
2300          End If
2310        End If

2320        If d.AssurInput = 1 Then
2330          If d.AssurBit = 1 Then  ' is it set ?
2340            d.AssurBit = 0       ' clear it
2350            PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 1
2360            IsAlarm = False
2370          End If
2380        End If

2390        If IsAlarm Then

              '530         If (d.IsAway = 0) Then
2400          If (1) Then
2410            If d.alarm = 0 Then
2420              d.alarm = 1
2430              Select Case d.AlarmMask
                    Case 2
2440                  PostEvent d, packet, Nothing, EVT_EXTERN, 1
2450                Case 1
2460                  PostEvent d, packet, Nothing, EVT_ALERT, 1
2470                Case Else
2480                  Trace " EVT_EMERGENCY 1"
2490                  PostEvent d, packet, Nothing, EVT_EMERGENCY, 1
2500              End Select
2510            End If
2520          End If
2530        End If

2540      Case SCI_CODE_ALARM1_CLEAR
2550        If d.ClearByReset Then
              ' defer to reset
2560        Else
2570          d.alarm = 0
2580          If d.AlarmMask = 1 Then
2590            PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 1
2600          Else
2610            Trace " EVT_EMERGENCY_RESTORE 1"
2620            PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 1
2630          End If
2640        End If


2650      Case SCI_CODE_ALARM2

2660        IsAlarm_A = True
2670        If (d.isDisabled_A) Then
2680          Exit Function          ' maybe
2690        End If

2700        If d.IsAway = 0 Then     ' if it's not on vacation then
2710          If d.AssurInput = 2 Then
2720            If d.Assur = 1 Then  ' an assurance device
2730              IsAlarm_A = False
2740            End If
2750          End If
2760        Else                     ' if it's on vacation
2770          If d.AssurInput = 2 Then
2780            If d.Assur = 1 Then  ' it's an assurance device
2790              If d.AssurSecure_A = 0 Then  ' they took it home
2800                IsAlarm_A = False
2810              End If
2820            End If
2830          End If
2840        End If

2850        If d.AssurInput = 2 Then
2860          If d.AssurBit = 1 Then  ' is it set ?
2870            d.AssurBit = 0       ' clear it
2880            PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 2
2890            IsAlarm_A = False
2900          End If
2910        End If

2920        If IsAlarm_A Then
              '1040          If (d.IsAway = 0) Then
2930          If (1) Then
2940            If d.Alarm_A = 0 Then
2950              d.Alarm_A = 1
2960              Select Case d.AlarmMask_A
                    Case 2
2970                  PostEvent d, packet, Nothing, EVT_EXTERN, 2
2980                Case 1
2990                  PostEvent d, packet, Nothing, EVT_ALERT, 2
3000                Case Else
3010                  Trace " EVT_EMERGENCY 2"
3020                  PostEvent d, packet, Nothing, EVT_EMERGENCY, 2
3030              End Select
3040            End If
3050          End If
3060        End If

3070      Case SCI_CODE_ALARM2_CLEAR
3080        If d.ClearByReset Then
              ' defer to reset
3090        Else
3100          d.Alarm_A = 0
3110          If d.AlarmMask_A = 1 Then
3120            PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 2
3130          Else
3140            Trace " EVT_EMERGENCY_RESTORE 2"
3150            PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 2
3160          End If
3170        End If

3180      Case SCI_CODE_ALARM3
3190        If d.isDisabled_B Then
3200          Exit Function          ' maybe
3210        End If

3220      Case SCI_CODE_ALARM3_CLEAR

3230      Case SCI_CODE_ALARM4
3240        If d.isDisabled Then
3250          Exit Function          ' maybe
3260        End If

3270      Case SCI_CODE_ALARM4_CLEAR

            'MISSING
            ' these are handled by DoSupervise in ModMain
3280      Case SCI_CODE_DEVICE_INACTIVE
3290        d.IsMissing = 1
            'Debug.Print "************* MISSING DEVICE ***********"
3300      Case SCI_CODE_REPEATER_INACTIVE
3310        d.IsMissing = 1
            'Debug.Print "************* MISSING REPEATER ***********"
3320      Case SCI_CODE_ACG_INACTIVE
3330        d.IsMissing = 1
            'Debug.Print "************* MISSING ACG ***********"

            ' these are handled by DoSupervise in ModMain
3340      Case SCI_CODE_DEVICE_INACTIVE_CLEARED
3350        d.Dead = 0               ' also resets IsMissing flag

3360      Case SCI_CODE_REPEATER_INACTIVE_CLEAR
3370        d.Dead = 0
3380      Case SCI_CODE_ACG_INACTIVE_CLEAR
3390        d.Dead = 0

3400      Case SCI_CODE_TAMPER
3410        If CBool(d.UseTamperAsInput) And (d.NumInputs = 2) Then
              ' New use tamper as input
3420          IsAlarm_B = True
3430          If d.isDisabled_B Then
3440            Exit Function        ' maybe
3450          End If

3460          If d.IsAway = 0 Then   ' if it's not on vacation then
3470            If d.AssurInput = 3 Then
3480              If d.Assur = 1 Then  ' an assurance device
3490                IsAlarm_B = False
3500              End If
3510            End If
3520          Else                   ' if it's on vacation
3530            If d.AssurInput = 3 Then
3540              If d.Assur = 1 Then  ' it's an assurance device
3550                If d.AssurSecure_B = 0 Then  ' they took it home
3560                  IsAlarm_B = False
3570                End If
3580              End If
3590            End If
3600          End If

3610          If d.AssurInput = 3 Then
3620            If d.AssurBit = 1 Then  ' is it set ?
3630              d.AssurBit = 0     ' clear it
3640              PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 3
3650              IsAlarm_B = False
3660            End If
3670          End If

3680          If IsAlarm_B Then

3690            If (1) Then
3700              If d.Alarm_B = 0 Then
3710                d.Alarm_B = 1
3720                Select Case d.AlarmMask_B
                      Case 2
3730                    PostEvent d, packet, Nothing, EVT_EXTERN, 3
3740                  Case 1
3750                    PostEvent d, packet, Nothing, EVT_ALERT, 3
3760                  Case Else
3770                    Trace " EVT_EMERGENCY 3"
3780                    PostEvent d, packet, Nothing, EVT_EMERGENCY, 3
3790                End Select
3800              End If
3810            End If
3820          End If

3830        Else                     ' normal tamper handling

3840          d.LastLocationText = ""
3850          packet.LocatedPartion = ""


3860          If d.Tamper = 0 Then
3870            PostEvent d, packet, Nothing, EVT_TAMPER, 0
3880          End If
3890        End If
3900      Case SCI_CODE_EOL_TAMPER

3910        d.LastLocationText = ""
3920        packet.LocatedPartion = ""

3930        If d.Tamper = 0 Then
3940          PostEvent d, packet, Nothing, EVT_TAMPER, 0
3950        End If
3960      Case SCI_CODE_REPEATER_TAMPER

3970        d.LastLocationText = ""
3980        packet.LocatedPartion = ""

3990        If d.Tamper = 0 Then
4000          PostEvent d, packet, Nothing, EVT_TAMPER, 0
4010        End If
4020      Case SCI_CODE_ACG_TAMPER

4030        d.LastLocationText = ""
4040        packet.LocatedPartion = ""

4050        If d.Tamper = 0 Then
4060          PostEvent d, packet, Nothing, EVT_TAMPER, 0
4070        End If

4080      Case SCI_CODE_TAMPER_CLEAR


4090        If CBool(d.UseTamperAsInput) And (d.NumInputs = 2) Then
4100          If d.ClearByReset Then
                ' defer to reset
4110          Else
4120            d.Alarm_B = 0
4130            If d.AlarmMask_B = 1 Then
4140              PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 3
4150            Else
4160              Trace " EVT_EMERGENCY_RESTORE 3"
4170              PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 3
4180            End If
4190          End If


4200        Else

4210          d.LastLocationText = ""
4220          packet.LocatedPartion = ""


4230          If d.Tamper = 1 Then
4240            PostEvent d, packet, Nothing, EVT_TAMPER_RESTORE, 0
4250          End If
4260        End If
4270      Case SCI_CODE_EOL_TAMPER_CLEAR

4280        d.LastLocationText = ""
4290        packet.LocatedPartion = ""


4300        If d.Tamper = 1 Then
4310          PostEvent d, packet, Nothing, EVT_TAMPER_RESTORE, 0
4320        End If
4330      Case SCI_CODE_REPEATER_TAMPER_CLEAR

4340        d.LastLocationText = ""
4350        packet.LocatedPartion = ""

4360        If d.Tamper = 1 Then
4370          PostEvent d, packet, Nothing, EVT_TAMPER_RESTORE, 0
4380        End If
4390      Case SCI_CODE_ACG_TAMPER_CLEAR

4400        d.LastLocationText = ""
4410        packet.LocatedPartion = ""


4420        If d.Tamper = 1 Then
4430          PostEvent d, packet, Nothing, EVT_TAMPER_RESTORE, 0
4440        End If


4450      Case SCI_CODE_LOW_BATT

4460        d.LastLocationText = ""
4470        packet.LocatedPartion = ""

4480        PostEvent d, packet, Nothing, EVT_BATTERY_FAIL, 0
4490      Case SCI_CODE_REPEATER_LOW_BATTERY
4500        d.LastLocationText = ""
4510        packet.LocatedPartion = ""


4520        PostEvent d, packet, Nothing, EVT_BATTERY_FAIL, 0
4530      Case SCI_CODE_ACG_LOW_BATTERY

4540        d.LastLocationText = ""
4550        packet.LocatedPartion = ""


4560        PostEvent d, packet, Nothing, EVT_BATTERY_FAIL, 0

4570      Case SCI_CODE_LOW_BATT_CLEAR

4580        d.LastLocationText = ""
4590        packet.LocatedPartion = ""

4600        PostEvent d, packet, Nothing, EVT_BATTERY_RESTORE, 0
4610      Case SCI_CODE_REPEATER_LOW_BATTERY_CLEAR
4620        PostEvent d, packet, Nothing, EVT_BATTERY_RESTORE, 0
4630      Case SCI_CODE_ACG_BATTERY_OK
4640        PostEvent d, packet, Nothing, EVT_BATTERY_RESTORE, 0


4650      Case SCI_CODE_MAINT_DUE
4660      Case SCI_CODE_MAINT_DUE_CLEAR

4670      Case SCI_CODE_DEVICE_RESET
            ' may have to poll device on 6080 to see it's status



4680        If d.ClearByReset Then
4690          d.alarm = 0
4700          If d.AlarmMask = 1 Then
4710            PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 1
4720          Else
4730            Trace " EVT_EMERGENCY_RESTORE 1"
4740            PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 1
4750          End If
4760        End If


4770      Case SCI_CODE_REPEATER_RESET
4780      Case SCI_CODE_ACG_RESET

4790      Case SCI_CODE_ENDPOINT_FAIL
4800      Case SCI_CODE_ENDPOINT_SUCCESS

            ' REPEATERS
4810      Case SCI_CODE_REPEATER_POWER_LOSS

4820        d.LastLocationText = ""
4830        packet.LocatedPartion = ""

4840        PostEvent d, packet, Nothing, EVT_LINELOSS, 0
4850      Case SCI_CODE_REPEATER_POWER_LOSS_CLEAR

4860        d.LastLocationText = ""
4870        packet.LocatedPartion = ""


4880        PostEvent d, packet, Nothing, EVT_LINELOSS_RESTORE, 0

4890      Case SCI_CODE_REPEATER_JAM

4900        d.LastLocationText = ""
4910        packet.LocatedPartion = ""

4920        d.IncrementJam 1
4930      Case SCI_CODE_ACG_JAMMED

4940      Case SCI_CODE_REPEATER_JAM_CLEAR
4950      Case SCI_CODE_ACG_JAM_CLEAR

4960      Case SCI_CODE_REPEATER_CONFIG_FAIL
4970      Case SCI_CODE_REPEATER_CONFIG_SUCCESS
            ''AGC
4980      Case SCI_CODE_ACG_CONFIG_FAIL
4990      Case SCI_CODE_ACG_CONFIG_SUCCESS
5000      Case SCI_CODE_ACG_CONFIG_CRC_FAIL
5010      Case SCI_CODE_ACG_FW_SUCCESS
5020      Case SCI_CODE_ACG_FW_FAIL
5030      Case SCI_CODE_ACG_BATTERY_FAIL
5040      Case SCI_CODE_ACG_SHUTDOWN_IMMINENT
5050      Case SCI_CODE_ACG_FW_PENDING
5060      Case SCI_CODE_ACG_IP_PROCESSOR_CRC_FAIL
5070      Case SCI_CODE_ACG_REBOOT_REQUESTED
5080      Case SCI_CODE_ACG_HELLO

5090    End Select

End Function

Public Function Ping(ByVal IPAddr As String) As Boolean
        Dim ResolveResult      As RESOLVE_ERRORS
        Dim PingIPv4           As PingIPv4
        Dim IP                 As String

10      On Error GoTo Ping_Error

20      Ping = False
30      IPAddr = Trim$(IPAddr)
40      If Len(IPAddr) Then
          
50        Set PingIPv4 = New PingIPv4
60        ResolveResult = PingIPv4.Resolve(IPAddr, IP)
70        If ResolveResult = RES_SUCCESS Then

80          If PingIPv4.Ping(IP) Then
              'lblRoundtrip.Caption = CStr(.RoundTripTime)
90            Ping = True
100         Else
              ' "Failure"
              ' lblReason.Caption = CStr(.Reason)
110         End If
            'If .Reason <> PFR_BAD_IP Then lblStatus.Caption = CStr(.Status)
120       Else
            'lblResults.Caption = "Bad Name or IP"
            'lblStatus.Caption = CStr(ResolveResult)
130         Ping = False
140       End If

150     End If




Ping_Resume:

160     Set PingIPv4 = Nothing
170     On Error GoTo 0
180     Exit Function

Ping_Error:

190     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at mod6080.Ping." & Erl
200     Resume Ping_Resume

End Function

'Public Function Ping(prmIPaddr As String, Optional ByVal MaxHops As Long = 20) As Boolean
'
'  Dim IPAddr As Long, HopsCount As Long, RTT As Long
'  'Dim MaxHops       As Long
'  Const Success = 1
'  'MaxHops = 20                 ' should be enough ...
'  IPAddr = inet_addr(prmIPaddr)
'
'  Ping = (GetRTTAndHopCount(IPAddr, HopsCount, MaxHops, RTT) = Success)
'
'End Function
Public Function GetWS() As String
  Dim ws          As String
  ws = "ws"
  If UseSecureSockets Then
    ws = "wss"
  End If
  GetWS = ws

End Function
Public Function GetHTTP() As String
  Dim http          As String
  http = "http"
  If UseSecureSockets Then
    http = "https"
  End If
  GetHTTP = http
End Function

Public Property Get First6080packet() As Boolean

  First6080packet = mFirst6080packet

End Property

Public Property Let First6080packet(ByVal Value As Boolean)
  ' used at startup when 6080 is ready, we need to sync missing devices
  ' this is also for recovering from 6080 failure
  If mFirst6080packet = False Then
    mFirst6080packet = Value ' set it so we don't call it again
    If Value Then
    ' start polling for missing devices
'      Debug.Assert 0
      GetAllZoneStatus
    End If
  End If

  mFirst6080packet = Value

End Property

'Public Function Backup6080Config() As Boolean
'  If BackUp6080Handler Is Nothing Then
'    Set BackUp6080Handler = New cHTTPRequest
'
'  End If
'  If Not (BackUp6080Handler.Busy) Then
'    BackUp6080Handler.Backup6080 "", GetHTTP & "://" & IP1, USER1, PW1
'  End If
'End Function


