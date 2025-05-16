VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAreaControlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Meta Section
Public Version      As String
Public MetaID       As String
Public SourceID     As String  ' GUID of 6080
Public LocalID      As String  ' serial of device
Public TimeString   As String  ' unix time
Public EventTime    As Date
Public Priority     As Long

Dim currstate As String

Public InfoID       As Long    ' ID as stored in 6080

Public SCICode      As Long
Public ValueState   As String
Public SerialData   As String
Public SerialDataValue   As Long


Public doc          As DOMDocument60
Public RootNode     As IXMLDOMNode
Public RootName     As String

Public ZoneType     As String

Private Node        As IXMLDOMNode
Private NodeList    As IXMLDOMNodeList


Public Partitions   As Collection



'Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'Priority/Severity 'Category 'Description
'0 'Emergency 'System or device is unusable; Hard fault/Total failure or catastrophic occurrence.
'1 'Alert 'Action must be taken immediately
'2 'Critical 'Critical condition(s) occurred
'3 'Error 'Error condition(s) occurred
'4 'Warning 'Significant/abnormal/warning conditions have occurred
'5 'Notice 'Normal but noteworthy conditions occurred
'6 'Informational 'Informative messages
'7 'Debug/Diagnostic 'System/Device debug messages

'Endpoint Device Codes:
'1 Alarm1
'2 Alarm1 has cleared
'3 Alarm2
'4 Alarm2 has cleared
'5 Alarm3
'6 Alarm3 has cleared
'7 Alarm3
'8 Alarm3 has cleared
'9 Device is Inactive
'10 Device is Inactive has cleared (is now active)
'11 tamper activated
'12 tamper cleared
'13 EOL tamper activated
'14 EOL tamper cleared
'15 Low battery
'16 Low battery cleared
'17 Maintenance required
'18 Maintenance required cleared
'21 Device Reset
'25 Endpoint Configuration Fail
'26 Endpoint Configuration Success
'27 Repeater Power Loss
'28 Repeater Power Loss clear
'29 Repeater Reset
'31 Repeater Tamper
'32 Repeater tamper clear
'33 Repeater low battery
'34 Repeater low battery clear
'35 Repeater Jam
'36 Repeater jam clear
'43 Repeater is inactive
'44 Repeater is inactive clear (is now active)
'45 Repeater Configuration fail
'46 Repeater Configuration Success
'49 ACG Reset
'51 ACG Tamper
'52 ACG Tamper clear
'55 ACG Jammed
'56 ACG Jam clear
'57 ACG Inactive
'58 ACG Inactive clear
'59 ACG Configuration Fail
'60 ACG Configuration success
'61 ACG CRC Check Fail
'71 ACG F/W Update Success
'72 ACG F/W Update Failed
'91 ACG Battery Failed
'92 ACG Battery low
'93 ACG Battery OK
'96 ACG Shutdown Imminent
'97 ACG F/W Update Pending
'99 ACG IP Processor CRC Invalid
'100 ACG Reboot Requested


' MID PTI is found only in Registration message

' actual sample
'<AreaControlEvent>
'  <MetadataHeader>
'    <MetaVersion>1.0</MetaVersion>
'    <MetaID>psiaalliance.org.AreaControl.Zone.alarmState</MetaID>
'    <MetaSourceID>{207DF398-5015-9281-4F2C-71CE056CEA4C}</MetaSourceID>
'    <MetaSourceLocalID>4046736</MetaSourceLocalID>
'    <MetaTime>2013-04-02T15:05:13.695Z</MetaTime>
'    <MetaPriority>1</MetaPriority>
'  </MetadataHeader>
'  <EventData>
'    <Info>
'      <ID>10</ID>
'      <Type>EN1223S Pendant</Type>
'      <Description>Pendant 3</Description>
'      <PartitionList>
'        <Partition>
'          <PartitionID>5</PartitionID>
'          <Description>Location 5</Description>
'        </Partition>
'        <Partition>
'          <PartitionID>4</PartitionID>
'          <Description>Location 4</Description>
'        </Partition>
'      </PartitionList>
'      <SCI>ALARM 1</SCI>
'      <SCICode>1</SCICode>
'    </Info>
'    <ValueState>
'      <IntrusionAlarm>Intrusion</IntrusionAlarm>
'    </ValueState>
'  </EventData>
'</AreaControlEvent>


' per manual
' Alarm Event
'<AreaControlEvent>
'<MetadataHeader>
'<MetaVersion>1.0</MetaVersion>
'<MetaID>psiaalliance.org/AreaControl.Zone/alarmState</MetaID>
'<MetaSourceID>{C15768C8-E695-4315-A06E-CF49E1409654}</MetaSourceID>
'<MetaSourceLocalID>6475870</MetaSourceLocalID>
'<MetaTime>2012-06-21T23:18:56.8969Z</MetaTime>
'<MetaPriority>1</MetaPriority>
'</MetadataHeader>
'<EventData>
'<Info>
'<ID>32</ID>
'<Type>EN1223S Pendant</Type>
'<Description>Device 6475870</Description>
'<PartitionList>
'<Partition>
'<ID>15</ID>
'<Description>Test Pendant</Description>
'</Partition>
'</PartitionList>
'<SCI>ALARM 1</SCI>
'<SCICode>1</SCICode>
'</Info>
'<ValueState>
'<IntrusionAlarm>Intrusion</IntrusionAlarm>
'</ValueState>
'</EventData>
'</AreaControlEvent>


' HELLO message every 5 seconds

'<AreaControlEvent>
'<MetadataHeader>
'<MetaVersion>1.0</MetaVersion>
'<MetaID>psiaalliance.org/AreaControl.System/systemState</MetaID>
'<MetaSourceID>{207DF398-5015-9281-4F2C-71CE056CEA4C}</MetaSourceID>
'<MetaSourceLocalID>7034681</MetaSourceLocalID>
'<MetaTime>2013-01-24T17:02:42.709Z</MetaTime>
'<MetaPriority>6</MetaPriority>
'</MetadataHeader>
'<EventData>
'<Info>
'<ID>1</ID>
'<Type>EN6080 Receiver</Type>
'<Description>base receiver</Description>
'<PartitionList></PartitionList>
'<SCI>hello</SCI>
'<SCICode>125</SCICode>
'</Info>
'<ValueState>
'<System>OK</System>
'</ValueState>
'</EventData>
'</AreaControlEvent>



Private Function ParseControlEvent(doc As DOMDocument60) As Long

        ' in <AreaControlEvent><MetadataHeader>

        Dim Node               As IXMLDOMNode
        Dim childnode          As IXMLDOMNode
        Dim LeafNode           As IXMLDOMNode
        Dim NodeList           As IXMLDOMNodeList

        Dim partition          As cPartition
10      Set Partitions = New Collection

        ' get source ID <AreaControlEvent><MetadataHeader><MetaSourceID> (GUID)
        ' Device id <AreaControlEvent><MetadataHeader><MetasourceLocalID> (serial in decimal)
        ' packet time <AreaControlEvent><MetadataHeader><MetaTime> (unix time yyyy-mm-ddThh:nn:ss.mmmmZ)
        ' Priority <AreaControlEvent><MetadataHeader><MetaPriority>


20      Set Node = doc.selectSingleNode("AreaControlEvent/MetadataHeader/MetaSourceID")
30      If Not Node Is Nothing Then
40        SourceID = Node.text
50      End If

60      Set Node = doc.selectSingleNode("AreaControlEvent/MetadataHeader/MetaSourceLocalID")
70      If Not Node Is Nothing Then
80        LocalID = Val(Node.text)
90      End If

100     If LocalID = 16359193 Then
          'Debug.Print doc.XML
          'Debug.Assert 0
110     End If

120     Set Node = doc.selectSingleNode("AreaControlEvent/MetadataHeader/MetaTime")
130     If Not Node Is Nothing Then
140       TimeString = Node.text
          Dim utc              As String
150       utc = ParseUnixTime(TimeString)
160       If IsDate(utc) Then
170         EventTime = Format(utc, "mm/dd/yyyy hh:nn:ss")
180       End If
190     End If


200     Set Node = doc.selectSingleNode("AreaControlEvent/MetadataHeader/MetaPriority")
210     If Not Node Is Nothing Then
220       Priority = Val(Node.text)
230     End If

        ' in <AreaControlEvent><EventData>
        ' in <AreaControlEvent><EventData><info>

240     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/ID")
250     If Not Node Is Nothing Then
260       InfoID = Val(Node.text)
270     End If


280     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/Type")
290     If Not Node Is Nothing Then
300       ZoneType = Node.text
310     End If


320     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/SCICode")
330     If Not Node Is Nothing Then

340       SCICode = Val(Node.text)
350     End If

360     If SCICode <> 125 Then
          'Debug.Assert 0
370     End If

380     If SCICode = SCI_CODE_SERIALDATA Then


390       Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/SerialData")  ' handle 1941XS Custom "Status Solutions" Device
400       If Not Node Is Nothing Then
410         SerialData = Right$("0000" & Node.text, 4)  ' string of hex bytes typically &h0C##

420         If InIDE Then
430           logserialdata SerialData
440         End If

            Dim AlarmsState    As String
            Dim AlarmsStateValue As Long

            Dim ResetPressed      As Boolean
            Dim PullCordRestored  As Boolean
            Dim ButtonPress       As Boolean
            Dim ButtonAlarm       As Boolean
            Dim PullcordPulled    As Boolean
            Dim ButtonUnplugged   As Boolean
            Dim ResetNoAlarms     As Boolean
            Dim Tamper            As Boolean



450         If left$(SerialData, 2) = "0C" Then  ' handle 1941XS Custom "Status Solutions" Device

              'MUST USE CLEAR BY RESET

              
              
460           AlarmsState = Right$(SerialData, 2)
470           AlarmsStateValue = Val("&h" & AlarmsState)

480           SerialDataValue = AlarmsStateValue

      '        ResetPressed = CBool(AlarmsStateValue And 4)
      '
      '        Reset = CBool((AlarmsStateValue And &HF4) = 4)
      '
      '        PullCordRestored = CBool(AlarmsStateValue = 0)  ' alarm in memory
      '
      '        ButtonPress = CBool(AlarmsStateValue And &H10)
      '        ButtonAlarm = CBool(AlarmsStateValue And &H20)
      '
      '        ButtonAlarm = ButtonAlarm Or ButtonPress
      '
      '        PullcordPulled = CBool(AlarmsStateValue And &H40)
      '
      '        ButtonUnplugged = CBool(AlarmsStateValue And &H8)  ' TAMPER
      '
      '        Tamper = ButtonUnplugged
      '
      '        ResetNoAlarms = CBool(AlarmsStateValue And &H4) And CBool((AlarmsStateValue And &HF0) = 0)
      '
      '
      '




      '        Select Case AlarmsState
      '          Case "00"            ' Pushbutton (tamper) restored, no alarms
      '            SerialValue = SCI_CODE_ALARM1_AND_ALARM2_AND_TAMPER_CLEAR
      '          Case "02"            ' UNDEFINED
      '          Case "04"            ' reset
      '            SerialValue = SCI_CODE_ALARM1_AND_ALARM2_AND_TAMPER_CLEAR
      '          Case "08"            ' pushbutton unplugged (tamper)
      '            SerialValue = SCI_CODE_TAMPER_XS
      '          Case "0C"            ' Reset w/ pushbutton unplugged (tamper)
      '            SerialValue = SCI_CODE_TAMPER_XS
      '          Case ""
      '          Case "20"            ' pushbutton
      '            SerialValue = SCI_CODE_ALARM1
      '          Case "22"            ' UNDEFINED
      '          Case "24"            ' UNDEFINED
      '          Case "28"            ' UNDEFINED
      '          Case "2C"            ' UNDEFINED
      '          Case "40"            ' pullcord
      '            SerialValue = SCI_CODE_ALARM2
      '          Case "42"            ' UNDEFINED
      '          Case "44"            ' reset w/ pullcord active
      '            SerialValue = SCI_CODE_ALARM2
      '          Case "48"            ' pushbutton unplugged w/ pullcord active
      '            SerialValue = SCI_CODE_ALARM2_AND_TAMPER_XS
      '          Case "4C"            ' UNDEFINED
      '          Case "60"            ' pull cord and pushbutton active
      '            SerialValue = SCI_CODE_ALARM1_AND_ALARM2
      '          Case Else
      '            ' UNDEFINED
      '        End Select
490         End If
500       End If
510     End If

520     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/CurrState")
530     If Not Node Is Nothing Then

540       currstate = Node.text
          'If InfoID = 46 Then
          '  Debug.Print "Area Control Event"
          '  Debug.Print doc.XML
          '

          'Debug.Print "***** ZoneType " & ZoneType & " CurrState " & currstate
          'End If
550     End If


560     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/ValueState")
570     If Not Node Is Nothing Then
580       ValueState = Node.text
590     End If


600     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/PartitionList")



610     If Not Node Is Nothing Then
620       For Each childnode In Node.childnodes
630         If childnode.baseName = "Partition" Then
640           Set partition = New cPartition
650           Partitions.Add partition
660           For Each LeafNode In childnode.childnodes
670             Select Case LeafNode.baseName
                  Case "PartitionID"
680                 Debug.Print "Partion ID " & LeafNode.text;
690                 partition.PartitionID = Val(LeafNode.text)
700               Case "Description"
710                 Debug.Print " " & LeafNode.text
720                 partition.Description = LeafNode.text
730               Case Else
                    ' nada
740             End Select

750           Next
760         End If

770       Next
780     End If

790     Set Node = doc.selectSingleNode("AreaControlEvent/EventData/Info/PartitionList")
800     If Not Node Is Nothing Then
810       For Each childnode In Node.childnodes

820       Next
830     End If




        '        <ID>10</ID> ID of unit in RX
        '    <Type>EN1223S Pendant</Type> ' type of device
        '    <Description>Pendant 3</Description> ' user supplied name (not used by us)
        '    <PartitionList>  ' list of possible locations in order of likelyhood
        '      <Partition>
        '        <PartitionID>5</PartitionID>
        '        <Description>Location 5</Description>
        '      </Partition>
        '      <Partition>
        '        <PartitionID>4</PartitionID>
        '        <Description>Location 4</Description>
        '      </Partition>
        '    </PartitionList>
        '    <SCI>ALARM 1</SCI>     ' verbose alarm
        '    <SCICode>1</SCICode>   ' code for alarm

        ' in <AreaControlEvent><EventData><ValueState>
        '<IntrusionAlarm>Intrusion</IntrusionAlarm> ' description of event state





End Function
Public Property Get HexID() As String
  HexID = Hex$(LocalID)
End Property

Private Function ParseUnixTime(ByVal TimeString As String) As Date
  Dim DatePart      As String
  Dim TimePart      As String
  Dim Parts()       As String

  Dim TimeZoneInformation As Win32.TIME_ZONE_INFORMATION
  Win32.GetTimeZoneInformation TimeZoneInformation

  Parts = Split(TimeString, "T", -1, vbTextCompare)
  DatePart = Parts(0)
  TimePart = left$(Parts(1), Len(Parts(1)) - 5)
  ParseUnixTime = DateAdd("n", -TimeZoneInformation.Bias - TimeZoneInformation.DaylightBias, Format$(DatePart, "mm-dd-yyyy") & " " & TimePart)



End Function

Public Function LoadXML(ByVal XML As String) As Long
  Dim rc            As Boolean

  On Error Resume Next
  Set doc = Nothing
  Set doc = New DOMDocument60
  doc.Async = False
  rc = doc.LoadXML(XML)
  

  If rc Then
    Set RootNode = doc.firstChild
    If RootNode Is Nothing Then
      rc = False
    Else
      RootName = RootNode.baseName
      ParseControlEvent doc
    End If
  End If
  LoadXML = rc
  
End Function
Sub logserialdata(ByVal SerialData As String)
  Dim hfile As Long
  Dim filename As String
  filename = App.Path & "\SerialData.log"
  limitFileSize filename
  
  hfile = FreeFile
  Open filename For Append As #hfile
  Print #hfile, SerialData
  Close #hfile
End Sub


Private Sub Class_Initialize()
  Set doc = New DOMDocument60
  Set Partitions = New Collection
End Sub

Private Sub Class_Terminate()
  Set doc = Nothing
End Sub
