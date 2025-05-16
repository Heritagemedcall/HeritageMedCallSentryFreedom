VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cESPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LocatedPartion As String ' this is what is finally used during location

Public LocatedPartionName1 As String
Public LocatedPartionName2 As String
Public LocatedPartionName3 As String

Public LocatedPartionID1 As Long
Public LocatedPartionID2 As Long
Public LocatedPartionID3 As Long

Public PartionName1 As String  ' Description is a 64 character text description
Public PartionName2 As String
Public PartionName3 As String
Public PartionName4 As String

Public PartionID1   As Long
Public PartionID2   As Long
Public PartionID3   As Long
Public PartionID4   As Long


Private mLinkedAlarm As Long

Public Partions     As Collection

Public XML          As String

Public Trouble      As Long
Public SCICode      As Long
Public SCIPriority  As Long
Public SCIValueState As String ' text of SCI Code Status

Public SerialDataValue As Long

Public Is6080       As Long
Public i6080data    As String

Public LocalID      As String

Private mAlarm2     As Long

Public Extra        As String



' Public Alarm3       as Long  not used for now, but alarm 3 would be a fourth input

'SEE Property MID and mMid Public MID          As Long ' used to bridge from old system to 6080

' public sub GetPartions
'
' public sub Delete Partiton from Zone
'                                                              v = Zone ID      v partiion ID
' DELETE /PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/2/PartitionList/12


' asign partion
'                                                           v = Zone ID to assign
'POST /PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/2/PartitionList HTTP/1.1\r\n
'                          v partiion ID
' <Partition> <PartitionID>1</PartitionID> </Partition>

' XML response (http 200)
'<ResponseStatus version="1.0" xmlns:urn="psialliance-org">
'  <requestURL>/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/ZoneInfo/PartitionLost/Partition</requestURL>
'  <statusCode>1</statusCode>
'  <statusString>201 Created</statusString>
'      v partition sequence
'  <id>1</id>
'</ResponseStatus>



'4 partitions to any single zone


Private mPacketSequence As Long

' UID

' First Byte  = Market ID (MID)
'          Byte = NODE ID   (NID)

'MID Application                MID Value (hex)   NID value
'Security End Device (one way)  B2                always 0
'Security End Device (two-way)  B2                0, assignable
'Environmental End Device       C0                0, assignable
'Submetering products           A0                0, assignable
'High-power repeaters           01                0, assignable
'Network coordinator (RX)       00                factory 1 to 31 (dec), assignable

'  Three Bytes of UID = Serial#
Private mHexPacket  As String             ' only calc it once, see function HexPacket

Private LocatorFormatByte As Integer

Private mPCAConfigured As Integer
Private mRegistering As Integer             ' for PCA, Repeater, Locator

Private mBADPacket  As Boolean
Private mAggregatePacket As Boolean
Public SubPackets   As Collection

Public TimeStamp    As Double
Public DateTime     As Date
'Public DataLength               As Integer  ' really limited to 255
Private mLastByte   As Integer               ' local variable, called by DataLength
Public Alarmtype    As Long
Public AvgCount     As Integer              'How many items to figure average signal strength

' from serial device
Public SerialPacket As String
Private mSerialMessage As String

Public BAD          As Boolean

' PacketType = 1 for serial-in device
Public PacketType   As Integer

Private mSerial     As String
Private PacketByte(255) As Byte

Private mOriginatorMID As Long
Private mOriginatorUID As Long

Private mFirstHopMID As Long
Private mFirstHopUID As Long

Private mTransType  As Byte

Private mMessageLength As Byte

Private mHOPs       As Byte
Private mTraces     As Byte
Private mClass      As Byte
Private mPTI        As Byte


Private mStat1      As Byte
Private mStat0      As Byte


Private mTemperature0 As Single
Private mTemperature1 As Single

Private mIEEE754(0 To 3) As Byte

Private mOpt1       As Byte
Private mOpt0       As Byte

Private mLevel      As Byte
Private mMargin     As Byte

Private mUnACKCount As Byte

Private mCheckSum   As Integer

Private mTraceMIDs() As Integer
Private mTraceUIDs() As Long

Private mDeliveryCode As Byte
Private mCmd        As Byte

Private mResponseCode As Byte
Private mSequence   As Long


' packet type x31 only
Private mNID        As Byte
Private mOutputType As Byte              ' either x11 or x30
Private mCFG        As Byte              ' x03 default
Private mCheckInSeconds As Long          ' two bytes
Private mLinkSuperSeconds As Long        ' two bytes


Private mLowBatt    As Byte
Private mLinkFailure As Byte
Private mReset      As Byte
Private mNoChange   As Byte
Private mTamper     As Byte
Private mJammed     As Byte
Private mLineLoss   As Byte

Private mRegisterPCA As Byte

Private mAck        As Byte

Private mSubCommand As Byte


Private mCleanMe    As Byte

'Pendant
Private mArm        As Byte
Private mDisarm     As Byte

'Others
Private mAlarm0     As Byte
Private mAlarm1     As Byte

Private mAckReq     As Boolean
Private mTraceCount As Byte

'Private mPayload(0 To 255)      As Byte
Private mRangeCheck As Boolean

Private mLocatedMID As Long
Private mLocatedUID As Long
Private mLocatorPacket As Boolean
Private mLocatedLevel As Byte
Private mLocatedMargin As Byte

Private mLayerAssigned As Byte
Private mOneNeighbor As Byte
Private mIsAck      As Byte
'unused = IIf(mStat1 And &H10, 1, 0)
Private mLayerID3   As Byte
Private mLayerID2   As Byte
Private mLayerID1   As Byte
Private mLayerID0   As Byte


' direct network repeater accumulated data
Private mLoPriRxCount As Long           ' max 0xFFFF
Private mLoPriTxCount As Long           ' max 0xFFFF
Private mLoPriDropCount As Long         ' max 0xFFFF
Private mHiPriRxCount As Long           ' max 0xFFFF
Private mHiPriTxCount As Long           ' max 0xFFFF
Private mHiPriDropCount As Long         ' max 0xFFFF
Private mNeighborBuildCount As Long     ' max 0xFFFF

Private mHardReset  As Boolean

' not used
Private Const MCB_LOCATOR = 99

'x36
' not used
Private Const PTI_LOCATOR = 99
'x01
Public Sub SetAlarm0(ByVal Value As Integer)
  mAlarm0 = (mAlarm0 Or Value)
End Sub
Public Sub SetAlarm1(ByVal Value As Integer)
  mAlarm1 = (mAlarm1 Or Value)
End Sub

Public Property Get LayerID() As Integer
    LayerID = CInt(mLayerID0 + 2 * mLayerID1 + 4 * mLayerID2 + 4 * mLayerID3)
End Property

Public Property Get LocatedSerial()
  LocatedSerial = HexFormat(mLocatedMID, 2) & HexFormat(mLocatedUID, 6)
End Property

Public Property Get LocatedLevel() As Byte
  LocatedLevel = mLocatedLevel
End Property

Public Property Get LocatedMargin() As Byte
  LocatedMargin = mLocatedMargin
End Property

Public Property Get IsLocatorPacket() As Boolean
 IsLocatorPacket = mLocatorPacket
End Property

Public Property Get IsAggregatePacket() As Boolean
  IsAggregatePacket = mAggregatePacket
End Property

Public Property Get LocatedMID() As Long
  LocatedMID = mLocatedMID
End Property
Public Property Get LocatedUID() As Long
  LocatedUID = mLocatedUID
End Property


Public Property Get TransType() As Integer
  TransType = mTransType
End Property
Public Property Get SubCommand() As Integer
  SubCommand = mSubCommand
End Property

Public Property Get NID() As Integer
  NID = mNID
End Property


Public Property Get RegisterPCA() As Byte
  RegisterPCA = mRegisterPCA
End Property

Public Property Get LinkFailure() As Byte
  LinkFailure = mLinkFailure
End Property

Public Property Get Jammed() As Byte
  Jammed = mJammed
End Property
Public Property Get LineLoss() As Byte
  If mLineLoss = 1 Then
    mLineLoss = mLineLoss
  End If
  LineLoss = mLineLoss
End Property

Public Property Get inputnum() As Long
  If mAlarm0 = 1 Then
    inputnum = 1
  ElseIf mAlarm1 = 1 Then
    inputnum = 2
  End If
End Property

Public Property Get CONFIGACK() As Boolean
  CONFIGACK = IIf(mAck = 0, False, True)
End Property

Public Property Get ACKREQ() As Boolean
  ACKREQ = mAckReq
End Property
Public Property Let Alarm(ByVal Value As Integer)
  ' 1 is alarm0, 2 is alarm1 and 0 is no alarm
  If Value = 1 Then
    mAlarm0 = 1
    mAlarm1 = 0
    mAlarm2 = 0
  ElseIf Value = 2 Then
    mAlarm0 = 0
    mAlarm1 = 1
    mAlarm2 = 0
  ElseIf Value = 3 Then
    mAlarm0 = 0
    mAlarm1 = 0
    mAlarm2 = 1
  Else
    mAlarm0 = 0
    mAlarm1 = 0
    mAlarm2 = 0
  End If

End Property
Public Property Get Alarm() As Integer
' generic alarm, either alarm of two alarms
  'Alarm = mAlarm0 Or mAlarm1

  If mAlarm0 = 1 Then
    Alarm = 1
  ElseIf mAlarm1 = 1 Then
    Alarm = 2
  ElseIf Alarm2 = 1 Then
    Alarm = 3
  End If

End Property



Public Property Get Alarm0() As Integer
  Alarm0 = mAlarm0
End Property

Public Property Get Alarm1() As Integer
  Alarm1 = mAlarm1
End Property

Public Property Get Arm() As Integer
  Arm = IIf(Button And 1, 1, 0)
End Property

Public Property Get Battery() As Integer
'  If Serial = "B210326A" Then
'    Battery = 1
'  Else
    Battery = mLowBatt
'  End If
End Property

Public Property Get Button() As Byte
  Button = mStat1
End Property

Public Sub SetMIDPTI(ByVal MIDValue As Long, ByVal PTIValue As Long)
  mOriginatorMID = MIDValue
  mPTI = PTIValue

End Sub

Public Sub SetMIDClassPTI(ByVal MIDValue As Long, ByVal ClassValue As Long, ByVal PTIValue As Long)
  mOriginatorMID = MIDValue
  mClass = ClassValue
  mPTI = PTIValue

End Sub

'EN3954 Two Way Communicator (0)
'MID = &HB2
'PTI = &H5B


' ********** ONE (1) input ***********

'ES1210 Universal End Device (1)
'MID = &HB2
'PTI = &H0

'ES1210SK Survey Kit Universal End Device (1)
'MID = &HB2
'PTI = &H4



'S1215 Universal End Device w/ Wall Tamper (1)
'MID = &HB2
'PTI = &H8

'ES1235S One Button Necklace Pendant (1)
'MID = &HB2
'PTI = &H10

'ES1233S One Button Necklace Pendant (1)
'MID = &HB2
'PTI = &H14

'EN1223S One Button Pendant (1)
'MID = &HB2
'PTI = &H18

'EN1223SK Survey Kit One Button Pendant (1)
'MID = &HB2
'PTI = &H1E


'ES1260 Inovonics PIR (1)
'MID = &HB2
'PTI = &H28

'ES1262 Bosch PIR (1)
'MID = &HB2
'PTI = &H29

'ES1265 Ceiling Mount PIR (1)
'MID = &HB2
'PTI = &H2A

'ES1242 Residential Smoke Detector (1)
'MID = &HB2
'PTI = &H2C

'ES1244 Residential Smoke Detector (1)
'MID = &HB2
'PTI = &H21


'ES1249 Bill Trap (1)
'MID = &HB2
'PTI = &H30

'ES1247 Glass Break Detector (1)
'MID = &HB2
'PTI = &H32


' ********** TWO (2) input ***********

'ES1212 Two Input Universal (2)
'MID = &HB2
'PTI = &H1

'ES1210W  Door/Window (2)
'MID = &HB2
'PTI = &H3

'ES1216 Two Input Universal w/ Wall Tamper (2)
'MID = &HB2
'PTI = &H9

'ES1215W Door/Window w/ Wall Tamper (2)
'MID = &HB2
'PTI = &HB

'ES1235D Two Button Necklace Pendant (2)
'MID = &HB2
'PTI = &H11

'ES1233D Two Button Necklace Pendant (2)
'MID = &HB2
'PTI = &H15

'EN1223D Two Button Pendant (2)
'MID = &HB2
'PTI = &H19

'ES1255 End Device Arm/Disarm (2)
'MID = &HB2
'PTI = &H24


'ES1941 OEM device
'MID = &HB2
'PTI = &HC



' ********** THREE (3) input ***********

'ES1234D Three Channel Necklace Pendant (3)
'MID = &HB2
'PTI = &H38


'ES1236D Three Channel Beltclip Pendant (3)
'MID = &HB2
'PTI = &H39

' ********** External Serial Input ***********

Public Property Get DataLength() As Integer
  DataLength = mLastByte
End Property

Public Property Get BadPacket() As Boolean
  BadPacket = mBADPacket
  
End Property
 Public Property Get HardReset() As Boolean
  HardReset = mHardReset
 End Property


Public Property Get PacketCount() As Integer
  PacketCount = SubPackets.Count
End Property

Public Property Let Bytes(Value() As Byte)
        ' main parser

        Const HeaderByte = 0
        Const LengthByte = 1
        Const MIDByte = 2
        Const UIDByte0 = 3
        Const UIDByte1 = 4
        Const UIDByte2 = 5

        Const FirstHopMIDByte = 6
        Const FirstHopUIDByte0 = 7
        Const FirstHopUIDByte1 = 8
        Const FirstHopUIDByte2 = 9

        Const TraceByte = 10
        Const ClassByte = 12
        Const PTIByte = 13
        Const PayloadStart = 14
        Const Stat1Byte = 14
        Const Stat0Byte = 15
        Const SignalByte = 16
        Const MarginByte = 17
        Const CheckSumByte = 18

        Dim j                  As Integer
        Dim i                  As Integer
        Dim k                  As Integer
        Dim SubArray()         As Byte  ' temp byte array to hold subpacket bytes
        Dim AggregatedCount    As Integer  ' how many subpackets to create
        Dim SubPacket          As cESPacket
        Dim SubMessageLength   As Integer  ' length of submessage
        Dim SubPacketCount     As Integer  ' just for looping

10      Set SubPackets = New Collection

        'dbg "Packet " & Packet2Hex(value)

20      On Error GoTo Bytes_Error

30      mLastByte = UBound(Value)

40      If mLastByte >= 1 Then
50        mTransType = Value(HeaderByte)
60        mMessageLength = Value(LengthByte)
70        If mMessageLength < mLastByte Then
80          Debug.Print "ShortMessage, Expected: " & mMessageLength & " got " & mLastByte
90          mBADPacket = True
100         Exit Property
110       End If
120     Else
130       Debug.Print "ShortMessage < 2 bytes"
140       mBADPacket = True
150       Exit Property
160     End If

        'ReDim PacketByte(0 To mLastByte)
170     For j = 0 To mLastByte
180       PacketByte(j) = Value(j)
190     Next
        'Debug.Print "packet " & Packet2Hex(value)

200     If mTransType = &H72 Or mTransType = &H1C Then

          'Debug.Print "mTransType   0x" & Hex(mTransType)

210     Else
          'Debug.Print "mTransType   0x" & Hex(mTransType)
220     End If

230     Select Case mTransType

          Case &H6                   ' Data ACK from RX
            'always  x06 x02 x08
240         mCheckSum = PacketByte(2)
250       Case &H15                  ' NAK from RX / invalid command or bad checksum
            ' reason codes?
            ' case &h23 ' button pressed
            ' case &h11 ' NodeAck
            ' case &h13 ' systemACL


260       Case &H1C
270         mUnACKCount = PacketByte(2)
280         mLinkFailure = IIf(PacketByte(4) And &H1, 1, 0)
            '&h2 Reserved
            '&h4 Reserved
290         mReset = IIf(PacketByte(4) And &H8, 1, 0)

300         mRegisterPCA = mReset
310         mNoChange = IIf(PacketByte(4) And &H10, 1, 0)

320         mTamper = IIf(PacketByte(4) And &H20, 1, 0)
            '&H40 Reserved
330         mJammed = IIf(PacketByte(4) And &H80, 1, 0)

            'Case &H11 ' NodeAck
            '  mSequence = PacketByte(2) * 256& + PacketByte(3)
            '  mResponseCode = (PacketByte(4) = 1)

            'Case &H13 ' systemACL
            '  mSequence = PacketByte(2) * 256& + PacketByte(3)
            '  mCmd = PacketByte(5)
            '  mResponseCode = (PacketByte(5) = 1)


340       Case &H30                  ' Config Request
350       Case &H31                  ' Network Controller Config Response ' standard ES
360         mSubCommand = PacketByte(2)
370         mNID = PacketByte(3)
380         mOutputType = PacketByte(4)
390         mCFG = PacketByte(5)
            ' check documentation for byte ordering!
400         mCheckInSeconds = CLng(PacketByte(6)) * 256& + PacketByte(7)
410         mLinkSuperSeconds = CLng(PacketByte(8)) * 256& + PacketByte(9)
420         mCheckSum = PacketByte(10)


430       Case &H35                  ' Network Controller Config Responses Directed Network
440         mSubCommand = PacketByte(2)
450         Select Case SubCommand
              Case &H82              ' reporting NID
460             mNID = PacketByte(3)
470             mCheckSum = PacketByte(4)

480           Case &H85              ' reporting check-in time
                'checkintime = PacketByte(3) * 256
                'checkintime = checkintime + PacketByte(4)
                'configbyte =PacketByte(5)
490             mCheckSum = PacketByte(6)


500           Case &H90              ' reporting Serial Number
510             mOriginatorMID = 0   ' always 0
520             mOriginatorUID = CLng(PacketByte(3)) * 65536
530             mOriginatorUID = mOriginatorUID + CLng(PacketByte(4)) * 256&
540             mOriginatorUID = mOriginatorUID + PacketByte(5)
550             mCheckSum = PacketByte(6)
560           Case Else
                ' not valid packet

570         End Select

            '********** regular packets **********

580       Case &H72                  ' typical transmission

590         mOriginatorMID = PacketByte(MIDByte)
600         If mOriginatorMID = &HC0 Then
610           mOriginatorMID = mOriginatorMID
620         End If
630         mOriginatorUID = PacketByte(UIDByte0) * 65536
640         mOriginatorUID = mOriginatorUID + CLng(PacketByte(UIDByte1)) * 256&
650         mOriginatorUID = mOriginatorUID + PacketByte(UIDByte2)

660         mSerial = Right("00" & Hex(mOriginatorMID), 2) & Right("000000" & Hex(mOriginatorUID), 6)

670         If mSerial = "B26D1239" Then
              '     Debug.Assert 0
680         End If

690         If mSerial = "B23D5443" Then
              'Debug.Assert 0
700         End If




            ' first hop
710         mFirstHopMID = PacketByte(FirstHopMIDByte)
720         mFirstHopUID = CLng(PacketByte(FirstHopUIDByte0)) * 65536
730         mFirstHopUID = mFirstHopUID + CLng(PacketByte(FirstHopUIDByte1)) * 256&
740         mFirstHopUID = mFirstHopUID + PacketByte(FirstHopUIDByte2)

750         mTraces = PacketByte(TraceByte)

760         If mTraces > 0 Then
770           ReDim mTraceMIDs(mTraces - 1)
780           ReDim mTraceUIDs(mTraces - 1)
790         Else
800           ReDim mTraceMIDs(0)
810           ReDim mTraceUIDs(0)
820         End If
830         i = 11

            'i = 11 + mTraces * 4   ' variable data begins here
            'mTraces = 0
840         For j = 0 To mTraces - 1
850           mTraceMIDs(j) = PacketByte(i): i = i + 1
860           mTraceUIDs(j) = PacketByte(i) * 65536: i = i + 1
870           mTraceUIDs(j) = mTraceUIDs(j) + CLng(PacketByte(i)) * 256&: i = i + 1
880           mTraceUIDs(j) = mTraceUIDs(j) + PacketByte(i): i = i + 1
890         Next

900         mHOPs = PacketByte(i): i = i + 1
910         mClass = PacketByte(i): i = i + 1  ' MCB ' 36 is locator ' 62/x3E pendant
920         mPTI = PacketByte(i): i = i + 1  ' PTI   ' 1  is locator ' 24/x18 pendant
            ' class is 65, pti 1 for locator
            ' locator class on reset  = 0
            ' locator pti on reset  = 5?




930         Select Case mClass       ' primary decision byte
                '72 15 B2 0F 69 50 00 11 6B CD 00 00 39 5B 80 23 00 03 01 47 47 13

                ' ************** REPEATER RESET ?? **************


              Case &H0               ' Message Class = 0 = repeater reset ?
                ' next byte if repeater reset
                ' mOriginatorMID = &h01 for repeater
                ' mPTI = &h00 for repeater reset
                ' mOriginatorMID = &hB2 for PCA


940             If mOriginatorMID = &H1 Then  ' always 1 for hard reset
                  'PTI holds NID upon a hard reset of repeater
950               mNID = mPTI
960               mLevel = PacketByte(i): i = i + 1
970               mMargin = PacketByte(i): i = i + 1
980               mReset = 1
990               mHardReset = True
1000              mCheckSum = PacketByte(i)
                  'dbg "Hard Reset of Serial, MID, NID, seq " & serial & ", " & mOriginatorMID & ", " & mNID & ", " & PacketSequence
                  'dbg "len packet " & Len(HexPacket) / 2 - 1 & " " & HexPacket


                  'REPEATER RESET
                  'Stop

1010            ElseIf mOriginatorMID = &HC0 Then

1020              mNID = mPTI
1030              mLevel = PacketByte(i): i = i + 1
1040              mMargin = PacketByte(i): i = i + 1
1050              mReset = 1
1060              mHardReset = True
1070              mCheckSum = PacketByte(i)

1080            ElseIf mOriginatorMID = &HB2 Then
1090              mNID = mPTI
                  'mNID = -1
1100              mLevel = PacketByte(i): i = i + 1
1110              mMargin = PacketByte(i): i = i + 1
1120              mReset = 1
1130              mHardReset = True
1140              mCheckSum = PacketByte(i)
1150              dbg "Registering mOriginatorMID=&HB2, mClass=0 " & HexPacket
1160            End If


                ' ************** AGGREGATED MESSAGE (DN) **************

1170          Case &H2               ' aggregated directed network
                ' multi packet
                ' 7231012B9D62002E1005000002030BB2300BF43E1801001A170BB2300C823E1801001C180BB2300BF43E1800001917222246
1180            mAggregatePacket = True

1190            AggregatedCount = mPTI  ' aggregated count is mpti from above
1200            For SubPacketCount = 1 To AggregatedCount

1210              Set SubPacket = New cESPacket
1220              SubMessageLength = PacketByte(i) + 7: i = i + 1

1230              ReDim SubArray(0 To SubMessageLength)  ' message plus checksum
1240              SubArray(HeaderByte) = &H72
1250              SubArray(LengthByte) = UBound(SubArray)

                  ' these are standard positions for these bytes
1260              SubArray(MIDByte) = PacketByte(i): i = i + 1
1270              SubArray(UIDByte0) = PacketByte(i): i = i + 1
1280              SubArray(UIDByte1) = PacketByte(i): i = i + 1
1290              SubArray(UIDByte2) = PacketByte(i): i = i + 1

1300              If SubMessageLength < 9 Then  ' sanity check
1310                mBADPacket = True
1320                Debug.Print "Short Packet in sub packet creation"
1330                Debug.Print Packet2Hex(Value)
1340                Exit Property
1350              End If

1360              SubArray(FirstHopMIDByte) = PacketByte(MIDByte)
1370              SubArray(FirstHopUIDByte0) = PacketByte(UIDByte0)
1380              SubArray(FirstHopUIDByte1) = PacketByte(UIDByte1)
1390              SubArray(FirstHopUIDByte2) = PacketByte(UIDByte2)

1400              SubArray(TraceByte) = 0  ' no trace info

1410              SubArray(ClassByte) = PacketByte(i): i = i + 1
1420              SubArray(PTIByte) = PacketByte(i): i = i + 1

1430              For k = PayloadStart To SubMessageLength - 1
1440                SubArray(k) = PacketByte(i): i = i + 1
1450              Next

1460              SubArray(SubMessageLength) = MakeCheckSum(SubArray)
1470              SubPacket.Bytes = SubArray
1480              SubPackets.Add SubPacket

1490            Next

                ' ************** UNK **************
1500          Case &H31              ' UNK
1510            dbg "Unk Packet type &H31 " & Me.HexPacket
1520            Select Case mPTI
                  Case &H1           'PTI_LOCATOR ' ??? FA series ???
                    'formatbyte= PacketByte(i):
                    ' i = i + 1
                    'located device
                    'mlocateddeviceMID = PacketByte(i): i = i + 1
                    'mlocateddeviceUID = PacketByte(i) * 65536: i = i + 1
                    'mlocateddeviceUID = mlocateddeviceUID + PacketByte(i) * 256&: i = i + 1
                    'mlocateddeviceUID = mlocateddeviceUID + PacketByte(i): i = i + 1
                    'mlocateddeviceSIG =PacketByte(i): i = i + 1
                    'mlocateddeviceMAR =PacketByte(i): i = i + 1
                    'mfirsthopSIG =PacketByte(i): i = i + 1
                    'mfirsthopMAR =PacketByte(i): i = i + 1
                    'mCheckSum = PacketByte(i)
1530            End Select

                ' ************** LOCATOR **************



                '                [0x72] - Header for inbound complete message.
                '[LEN] - Message length, excluding checksum.
                '[UID (4) originator] - Unique ID of the locator originating the
                'message. Locator MID is 0xB2.
                '[UID (4) first hop] - Unique ID of the device which received
                'original message, as well as the signal level and margin. If the
                'UID begins with 00, then the RF gateway is the first hop; if the
                'UID begins with 01, then a repeater is the first hop.
                '[Trace count (1)] - Number of trace unique IDs that follow. The
                'default is zero, indicating no trace data.
                '[Trace UID1 (4)]...[Trace UIDn (4)] - Unique ID of each device that
                'repeats this message.
                '[Hop count (1)] - The number of times this message was relayed by a
                'repeater.
                '[0x36] - Message class byte for locator message.
                '[0x01] = Product type identifier for E*5080.
                '[0x00] = Data format byte for locator status message.
                '[STAT1] - Application flags.
                'Bit 7 - Reserved.
                'Bit 6 - Reserved.
                'Bit 5 - Acknowledgement, status is sent as an acknowledgement of
                'a configuration message.
                'Bit 4 - Activator state.
                'Bit 3 - Reserved.
                'Bit 2 - Reserved.
                'Bit 1 - Reserved.
                'Bit 0 - Reserved.
                '[STAT0] - Locator status flags.
                'Bit 7 - Locator is jammed.
                'Bit 6 - Low battery.
                'Bit 5 - Case tamper.
                'Bit 4 - Supervisory message, no change in status since last
                'transmission.
                'Bit 3 - Reset, this status message was transmitted immediately
                'following reset.
                'Bit 2 - Reserved.
                'Bit 1 - Loss of line power, locator is operating on battery
                'power.
                'Bit 0 - Reserved.
                '[Level] - Signal level of message as measured by the first hop
                'device, or by the gateway if survey information did not exist in
                'the message.
                '[Margin] - Signal margin of message as measured by the first hop
                'device, or by the gateway if survey information did not exist in
                'the message.
                '[CKSUM] - Checksum.

                'End Device Location Message
                '[0x72] - Header for inbound complete message.
                '[LEN] - Message length, excluding checksum.
                '[UID (4) originator] - Unique ID of the locator originating the
                'message. Locator MID is 0xB2.
                '[UID (4) first hop] - Unique ID of the device which received this
                'message from the locator, as well as the signal level and margin.
                'If the UID begins with 00, then the RF gateway is the first hop; if
                'the UID begins with 01, then a repeater is the first hop.
                '[Trace count (1)] - Number of trace unique IDs that follow. The
                'default is zero, indicating no trace data.
                '[Trace UID1 (4)]...[Trace UIDn (4)] - Unique ID of each device that
                'repeats this message.
                '[Hop count (1)] - The number of times this message was relayed by a
                'repeater.
                '[0x36] - Message class byte for locator message.
                '[0x01] = Product type identifier for EE/EN5080.
                '[DFB] = Data format byte for locator status message.
                '[0x01] - Location in alarm.
                '[0x02] - Location on tracked mobile asset.
                '[UID (4) end device] - Unique ID of the end device being located.
                '[RSSI Level] - Signal level of the end device message as measured
                'by the locator.
                '[RSSI Margin] - Signal margin of the end device message as measured
                'by the locator.
                '[Locator Level] - Signal level of the locator message as measured
                'by the first hop device, or by the gateway if survey information
                'did not exist in the message.
                '[Locator Margin] - Signal margin of the locator message as measured
                'by the first hop device, or by the gateway if survey information
                'did not exist in the message.
                '[CKSUM] - Checksum.
                'Configure List of Product Type Identifiers
                '[0x36] - Message class byte for a configuration message to a
                'locator.
                '[0x02] = Subclass byte indicating list of PTI to locate.
                '[PTI1 ... PTI24] =List of PTI of 24 device types to locate on (set
                'unused to 0xFF)
                'Configure Minimum Location Threshold
                ' [0x36] - Message class byte for a configuration message to a
                'locator.
                '[0x03] = Subclass byte indicating the minimal qualifying
                'threshold.
                '[LOCATION THRESHOLD] = A value of 0 to 99, indicating the level
                'below which a message from a device will not located on.

                'Configure Minimum Actuator Threshold
                '[0x36] = Message class byte for a configuration message to a
                'locator.
                '[0x04] = Subclass byte indicating the minimal actuator threshold.
                '[ACTUATOR THESHOLD] = A value of 0 to 99, indicating the level
                'below which a message from a device will not activate actuator.


1540          Case &H36
                'class &H36
1550            mBADPacket = True
1560            LogProgramError "Unsupported Product PTI =  " & PTI
1570            Exit Property
                'dbg "Locator Status Packet Received"

1580            If mPTI <> 1 Then
1590              mBADPacket = True
1600              LogProgramError "Error in Locator Packet PTI =  " & PTI
1610              Exit Property
1620            End If

                'mLocatorPacket = True
1630            LocatorFormatByte = PacketByte(i): i = i + 1
1640            Select Case LocatorFormatByte  ' Data format byte for locator status message.
                  Case 0
                    ' it's the locator itself
1650                mStat1 = PacketByte(i): i = i + 1
                    'Bit 5 - Acknowledgement, status is sent as an acknowledgement of a configuration message.
                    'Bit 4 - Activator state.

1660                mStat0 = PacketByte(i): i = i + 1
                    'Bit 1 - Loss of line power, locator is operating on battery
1670                mLevel = PacketByte(i): i = i + 1
1680                mMargin = PacketByte(i): i = i + 1
1690                mCheckSum = PacketByte(i)

                    ' parse out status bytes
                    ' stat 1
                    'mLayerAssigned = IIf(mStat1 And &H80, 1, 0)
                    'mOneNeighbor = IIf(mStat1 And &H40, 1, 0)
1700                mAck = IIf(mStat1 And &H20, 1, 0)

                    '               If mSerial = "B23B5D17" Then
                    '                If mAck Then
                    '                  'Stop
                    '                End If
                    '               End If

                    'unused = IIf(mStat1 And &H10, 1, 0)
                    'mLayerID3 = IIf(mStat1 And &H8, 1, 0)
                    'mLayerID2 = IIf(mStat1 And &H4, 1, 0)
                    'mLayerID1 = IIf(mStat1 And &H2, 1, 0)
                    'mLayerID0 = IIf(mStat1 And &H1, 1, 0)

                    'stat0
1710                mJammed = IIf(mStat0 And &H80, 1, 0)
1720                mLowBatt = IIf(mStat0 And &H40, 1, 0)
                    'Debug.Print "Locator lowbatt "; mLowBatt
                    'mTamper = IIf(mStat0 And &H20, 1, 0)
1730                mLineLoss = IIf(mStat0 And &H20, 1, 0)

1740                Debug.Print "Locator LineLoss (Was Tamper) "; mLineLoss
                    'dbg "Locator Reset " & IIf(mStat0 And &H8, 1, 0)
                    'dbg "Locator ACK " & IIf(mStat1 And &H2, 1, 0)

1750                mNoChange = IIf(mStat0 And &H10, 1, 0)
                    'mReset = IIf(mStat0 And &H8, 1, 0)
1760                If mReset Then
                      ' mReset = mReset
1770                End If
                    'mRegistering = mAck

                    'dbg "Broadcast ACK to configure = " & serial & ", " & mAck

1780              Case 1
                    'LOC 72170128D425002E10050000360101B2300BF41E1E2A2A97
                    ' it's locating some other device
                    ' 4 bytes of located device
1790                mLocatorPacket = True

                    '[DFB]

1800                mLocatedMID = PacketByte(i): i = i + 1
1810                mLocatedUID = PacketByte(i) * 65536: i = i + 1
1820                mLocatedUID = mLocatedUID + CLng(PacketByte(i)) * 256&: i = i + 1
1830                mLocatedUID = mLocatedUID + PacketByte(i): i = i + 1




                    ' only signal data is transmitted, no status info
1840                mLocatedLevel = PacketByte(i): i = i + 1
1850                mLocatedMargin = PacketByte(i): i = i + 1

                    ' level of locator itself
1860                mLevel = PacketByte(i): i = i + 1
1870                mMargin = PacketByte(i): i = i + 1
1880                mCheckSum = PacketByte(i)

                    'dbg "Broadcast ACK to configure = " & serial & ", " & mAck
                    'dbg "LOC " & Me.HexPacket
                    'dbg "Device  " & HexFormat(mLocatedMID, 2) & HexFormat(mLocatedUID, 6)
                    'dbg "Locator " & HexFormat(mOriginatorMID, 2) & HexFormat(mOriginatorUID, 6)
                    'dbg "Level   " & mLocatedLevel

1890              Case Else
                    'BadPacket = True
1900            End Select

                ' ************** PCA **************
1910          Case &H39              ' PCA (Portable communications assistant)
1920            Select Case mPTI
                  Case &H5B          ' PTI for PCA
1930                mDeliveryCode = PacketByte(i): i = i + 1  ' Bit seven set if ACK is required, all others reserved

                    '72 15 B2 0F 69 50 00 11 6B CD 00 00 39 5B 80 23 00 03 01 47 47 13
                    '                                          ^
1940                mAckReq = (mDeliveryCode And (2 ^ 7)) <> 0

1950                mCmd = PacketByte(i): i = i + 1  ' this can be either a command or response to a command

                    'If mCmd = &H17 Then Stop

1960                Select Case mCmd

                      Case &H17

1970                    mStat1 = PacketByte(i): i = i + 1
1980                    mPCAConfigured = IIf(mStat1 And &H20, 1, 0)
1990                    mStat0 = PacketByte(i): i = i + 1
2000                    mLevel = PacketByte(i): i = i + 1
2010                    mMargin = PacketByte(i): i = i + 1
2020                    mCheckSum = PacketByte(i)
2030                    mReset = IIf(mStat0 And &H8, 1, 0)
2040                    mRegisterPCA = mPCAConfigured

2050                  Case &H11      ' ACK

2060                    mSequence = CLng(PacketByte(i)) * CLng(256): i = i + 1
2070                    mSequence = mSequence + PacketByte(i): i = i + 1
2080                    mResponseCode = PacketByte(i): i = i + 1
2090                    mLevel = PacketByte(i): i = i + 1
2100                    mMargin = PacketByte(i): i = i + 1
2110                    mCheckSum = PacketByte(i)

                        '72 15 B2 0F 69 50 00 11 6B CD 00 00 39 5B 80 23 00 03 01 47 47 13

                        'Range check:
                        '72 12 B2 11 D4 84 00 11 6B CD 00 00 39 5B 80 26 3F 3A
2120                  Case &H26
2130                    mRangeCheck = True

2140                  Case &H23
2150                    mSequence = CLng(PacketByte(i)) * CLng(256): i = i + 1
2160                    mSequence = mSequence + PacketByte(i): i = i + 1
2170                    mResponseCode = PacketByte(i): i = i + 1
2180                    mLevel = PacketByte(i): i = i + 1
2190                    mMargin = PacketByte(i): i = i + 1
2200                    mCheckSum = PacketByte(i)
2210                  Case Else
2220                    mStat1 = PacketByte(i): i = i + 1
2230                    mStat0 = PacketByte(i): i = i + 1
2240                    mLevel = PacketByte(i): i = i + 1
2250                    mMargin = PacketByte(i): i = i + 1
2260                    mCheckSum = PacketByte(i)
2270                    mReset = IIf(mStat0 And &H8, 1, 0)
2280                End Select
2290            End Select


                ' ************** TYPICAL SECURITY DEVICE **************
2300          Case &H3E              ' Security Device

2310            mStat1 = PacketByte(i): i = i + 1
2320            mStat0 = PacketByte(i): i = i + 1
2330            mLevel = PacketByte(i): i = i + 1
2340            mMargin = PacketByte(i): i = i + 1
2350            mCheckSum = PacketByte(i)

2360            Select Case mPTI     ' handle by PTI

                    '&H1 is 1212 w input
                    '&H2 is 1210-60
                    '&H3 is 1210W
                    '&H4 is 1210SK survey device
                    '&H6 is 1210-240
                    '&H7 is 1252 2-input extended range
                    '&H8 is 1215
                    '&H9 is 1216 2-input
                    '&HA is 1212-60
                    '&HB is 1215W 2-input
                    '&HC is 1941 2-input
                    '&HF is 1941-60 2-input

                    '&H10 is 1235S  one button pendant
                    '&H11 is 1235D pendant 2 channel
                    '&H12 is 1235SF  One button Fixed
                    '&H14 is 1223S pendant
                    '&H15 is 1223D 2-input pendant
                    '&H18 is 1223 pendant
                    '&H19 is 1223D 2-input
                    '&H1A is 1223S-60 pendant
                    '&H1D is 1221S-60 pendant
                    '&H1E is 1223SK Survey device obsolete

                    '&H21 is 1244 Smoke detector
                    '&H23 is 1751 water detect
                    '&H24 is 1255 Arm Disarm 2-input
                    '&H28 is 1260 PIR
                    '&H29 is 1262 PIR
                    '&H2A is 1265 PIR
                    '&H2C is 1242 Smoke detector
                    '&H2D is 1241-60 activity sensor

                    '&H30 is 1249 Bill Trap
                    '&H31 is 1240 activity sensor
                    '&H32 is 1247 Glass break
                    '&H38 is 1224D pendant 3 channel
                    '&H3A is 1238D Two channel Belt Clip
                    '&H3B is 1254 4-channel Belt
                    '&H3E is 1236D Three channel belt clip

                    '&H5B is 3954 PCA

                    '&H73 is 3942XS Bideirectional Serial
                    '&H74 is 3941XS Bideirectional Serial






                  Case &H0, &H2, &H4, &H6, &H18, &H29, &H2A, &H2C, &H2D, &H30, &H31, &H32  ' ONE (1) input device

2370                If mPTI = &H2C Then  ' product type is ES1232 smoke detector
2380                  mCleanMe = IIf(mStat1 And &H40, 1, 0)

2390                End If

2400                mAlarm1 = 0      ' always 0 for single button
2410                mAlarm0 = IIf(mStat1 And &H1, 1, 0)  ' single button press
2420                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2430                mTamper = IIf(mStat0 And &H20, 1, 0)
2440                mNoChange = IIf(mStat0 And &H10, 1, 0)
2450                mReset = IIf(mStat0 And &H8, 1, 0)


2460              Case &H21

2470                mAlarm1 = 0      ' always 0 for smoke


                    ' mStat1 And &H40 not used
2480                mCleanMe = IIf(mStat1 And &H40, 1, 0)
2490                mCleanMe = mCleanMe Or IIf(mStat1 And &H20, 1, 0)
                    'mStat1 And &H20 not used
                    'mStat1 And &H10 not used (Test Alarm)
                    ' 02,04,08 not used
2500                mAlarm0 = IIf(mStat1 And &H1, 1, 0)  ' alarm


2510                mLowBatt = IIf(mStat0 And &H80, 1, 0)
                    ''mLowBatt = IIf(mStat0 And &H40, 1, 0) not used
2520                mTamper = IIf(mStat0 And &H20, 1, 0)
2530                mNoChange = IIf(mStat0 And &H10, 1, 0)
2540                mReset = IIf(mStat0 And &H8, 1, 0)
                    ' 01,02,04 not used

                    'Debug.Print "cESPacket.bytes mAlarm0=" & mAlarm0 & " Serial=" & Me.HexPacket
2550              Case &H28          ' Innovonics PIR

2560                mAlarm1 = 0
2570                mAlarm0 = IIf(mStat1 And &H1, 1, 0)

2580                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2590                mTamper = IIf(mStat0 And &H20, 1, 0)
2600                mNoChange = IIf(mStat0 And &H10, 1, 0)
2610                mReset = IIf(mStat0 And &H8, 1, 0)



2620              Case &H10, &H14, &H1E, &H1D  ' portable 1-button pendants

2630                mAlarm1 = 0
2640                mAlarm0 = IIf(mStat1 And &H1, 1, 0)
2650                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2660                mTamper = IIf(mStat0 And &H20, 1, 0)
2670                mNoChange = IIf(mStat0 And &H10, 1, 0)
2680                mReset = IIf(mStat0 And &H8, 1, 0)

2690              Case &H24          ' arming device

2700                mDisarm = IIf(mStat1 And &H2, 1, 0)
2710                mArm = IIf(mStat1 And &H1, 1, 0)

2720                mAlarm1 = 0
2730                mAlarm0 = 0

2740                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2750                mTamper = IIf(mStat0 And &H20, 1, 0)
2760                mNoChange = IIf(mStat0 And &H10, 1, 0)
2770                mReset = IIf(mStat0 And &H8, 1, 0)


2780              Case &H1, &H3, &H7, &H9, &HB, &H11, &H15, &H19, &H24, &H38, &H3A, &H99  '  TWO (2) input device
2790                mAlarm0 = IIf(mStat1 And &H2, 1, 0)  ' flipped these
2800                mAlarm1 = IIf(mStat1 And &H1, 1, 0)
2810                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2820                mTamper = IIf(mStat0 And &H20, 1, 0)
2830                mNoChange = IIf(mStat0 And &H10, 1, 0)
2840                mReset = IIf(mStat0 And &H8, 1, 0)



2850              Case &HC, &HF      ' 1941 OEM

                    '    ESDeviceType(j).Model = "EN1941XS"
                    '    ESDeviceType(j).desc = "End Device Serial Input"
                    '    ESDeviceType(j).MID = &HB2
                    '    ESDeviceType(j).PTI = &HC    ' shoud be 12-decimal &h0C
                    '    ESDeviceType(j).CLS = &H3E   ' assumed
                    '    ESDeviceType(j).NumInputs = 2

                    'Debug.Print HexPacket
2860                mAlarm1 = IIf(mStat1 And &H2, 1, 0)
2870                mAlarm0 = IIf(mStat1 And &H1, 1, 0)
2880                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2890                mTamper = IIf(mStat0 And &H20, 1, 0)
2900                mNoChange = IIf(mStat0 And &H10, 1, 0)
2910                mReset = IIf(mStat0 And &H8, 1, 0)


2920              Case Else

2930                mAlarm1 = IIf(mStat1 And &H2, 1, 0)
2940                mAlarm0 = IIf(mStat1 And &H1, 1, 0)
2950                mLowBatt = IIf(mStat0 And &H40, 1, 0)
2960                mTamper = IIf(mStat0 And &H20, 1, 0)
2970                mNoChange = IIf(mStat0 And &H10, 1, 0)
2980                mReset = IIf(mStat0 And &H8, 1, 0)

2990            End Select

                ' ************** ENVIRONMENTAL DEVICE **************
3000          Case &H3C              ' Environmental (1723 temperatur device)
                ' what we do here is handle the info as it comes in
                ' depending on the configuration  of the device
                ' then afterwards, set the PTI to a generic 99 so that
                ' the 1723 device type is is only included once in the system

3010            If mOriginatorMID = &HC0 Then
                  ' the pti will change if you change the jumper or other programmable settings
                  ' we'll handle all here and then
                  ' change the mPTI at the end
3020              Select Case mPTI

                    Case &HA         ' Environmental Internal Resistor

3030                  mIEEE754(3) = PacketByte(i): i = i + 1
3040                  mIEEE754(2) = PacketByte(i): i = i + 1
3050                  mIEEE754(1) = PacketByte(i): i = i + 1
3060                  mIEEE754(0) = PacketByte(i): i = i + 1
3070                  mTemperature0 = IEEE_754_ByteArray_to_Single(mIEEE754)
                      'PacketByte(i): i = i + 4    ' convert 3 bytes to float value

3080                  mStat0 = PacketByte(i): i = i + 1
3090                  mLevel = PacketByte(i): i = i + 1
3100                  mMargin = PacketByte(i): i = i + 1
3110                  mCheckSum = PacketByte(i)
3120                  mReset = IIf(mStat0 And &H8, 1, 0)
3130                Case &H2A        ' Environmental RESET ? or With Options ?
3140                  mIEEE754(3) = PacketByte(i): i = i + 1
3150                  mIEEE754(2) = PacketByte(i): i = i + 1
3160                  mIEEE754(1) = PacketByte(i): i = i + 1
3170                  mIEEE754(0) = PacketByte(i): i = i + 1
3180                  mTemperature0 = IEEE_754_ByteArray_to_Single(mIEEE754)

                      'mTemperature= PacketByte(i): i = i + 4  ' convert 3 bytes to float value
3190                  mOpt1 = PacketByte(i): i = i + 1
3200                  mOpt0 = PacketByte(i): i = i + 1
3210                  mStat1 = PacketByte(i): i = i + 1
3220                  mStat0 = PacketByte(i): i = i + 1
3230                  mLevel = PacketByte(i): i = i + 1
3240                  mMargin = PacketByte(i): i = i + 1
3250                  mCheckSum = PacketByte(i)
3260                  mReset = IIf(mStat0 And &H8, 1, 0)

3270                Case &H17        ' Environmental Two Temps
3280                  mIEEE754(3) = PacketByte(i): i = i + 1
3290                  mIEEE754(2) = PacketByte(i): i = i + 1
3300                  mIEEE754(1) = PacketByte(i): i = i + 1
3310                  mIEEE754(0) = PacketByte(i): i = i + 1
3320                  mTemperature0 = IEEE_754_ByteArray_to_Single(mIEEE754)

3330                  mIEEE754(3) = PacketByte(i): i = i + 1
3340                  mIEEE754(2) = PacketByte(i): i = i + 1
3350                  mIEEE754(1) = PacketByte(i): i = i + 1
3360                  mIEEE754(0) = PacketByte(i): i = i + 1
3370                  mTemperature1 = IEEE_754_ByteArray_to_Single(mIEEE754)



                      'mTemperature= PacketByte(i): i = i + 4  ' convert 3 bytes to float value
                      'mTemperature1 = PacketByte(i): i = i + 4  ' convert 3 bytes to float value
3380                  mStat1 = PacketByte(i): i = i + 1
3390                  mStat0 = PacketByte(i): i = i + 1  ' 16 is "no change"
3400                  mLevel = PacketByte(i): i = i + 1
3410                  mMargin = PacketByte(i): i = i + 1
3420                  mCheckSum = PacketByte(i)
3430                  mReset = IIf(mStat0 And &H8, 1, 0)

3440                Case &H37        ' Environmental Two Temps w/ options
3450                  mIEEE754(3) = PacketByte(i): i = i + 1
3460                  mIEEE754(2) = PacketByte(i): i = i + 1
3470                  mIEEE754(1) = PacketByte(i): i = i + 1
3480                  mIEEE754(0) = PacketByte(i): i = i + 1
3490                  mTemperature0 = IEEE_754_ByteArray_to_Single(mIEEE754)

3500                  mIEEE754(3) = PacketByte(i): i = i + 1
3510                  mIEEE754(2) = PacketByte(i): i = i + 1
3520                  mIEEE754(1) = PacketByte(i): i = i + 1
3530                  mIEEE754(0) = PacketByte(i): i = i + 1
3540                  mTemperature1 = IEEE_754_ByteArray_to_Single(mIEEE754)


                      '        mTemperature= PacketByte(i): i = i + 4  ' convert 3 bytes to float value
                      '        mTemperature1 = PacketByte(i): i = i + 4  ' convert 3 bytes to float value
3550                  mOpt1 = PacketByte(i): i = i + 1  ' not sent normally
3560                  mOpt0 = PacketByte(i): i = i + 1  ' not sent normally
3570                  mStat1 = PacketByte(i): i = i + 1
3580                  mStat0 = PacketByte(i): i = i + 1  ' 16 is "no change"
3590                  mLevel = PacketByte(i): i = i + 1
3600                  mMargin = PacketByte(i): i = i + 1
3610                  mCheckSum = PacketByte(i)
3620                  mReset = IIf(mStat0 And &H8, 1, 0)
3630                Case Else
3640              End Select
3650              Debug.Print "CurrentTemperature I is " & mTemperature0
3660              Debug.Print "CurrentTemperature X is " & mTemperature1

                  '      If mOpt1 Then Stop
                  '      If mOpt0 Then Stop
                  'If mStat1 Then Stop
                  'If (mStat0 And 16) = 0 Then Stop
3670              mPTI = &H99
3680            End If


                ' ************** REPEATER **************
3690          Case &H41              ' Repeater Class Byte
                ' class 65
3700            Select Case mPTI
                    ' **************  BROADCAST REPEATER PTI
                  Case &H0
3710                mStat1 = PacketByte(i): i = i + 1
3720                mStat0 = PacketByte(i): i = i + 1
3730                mLevel = PacketByte(i): i = i + 1
3740                mMargin = PacketByte(i): i = i + 1
3750                mCheckSum = PacketByte(i)
3760                mAck = IIf(mStat1 And &H20, 1, 0)  ' RESPONSE TO CONFIGURATION MESSAGE
3770                mJammed = IIf(mStat0 And &H80, 1, 0)
3780                mLowBatt = IIf(mStat0 And &H40, 1, 0)
3790                mTamper = IIf(mStat0 And &H20, 1, 0)
3800                mNoChange = IIf(mStat0 And &H10, 1, 0)
3810                mReset = IIf(mStat0 And &H8, 1, 0)
                    ' &h4 reserved
3820                mLineLoss = IIf(mStat0 And &H2, 1, 0)
                    ' &h1 reserved

                    'dbg "Broadcast ACK to configure = " & serial & ", " & mAck


                    ' ************** DIRECTED MESSAGE REPEATER
3830              Case &H1

3840                mLoPriRxCount = CLng(PacketByte(i)) * 256&: i = i + 1  'LoPriRxCount
3850                mLoPriRxCount = mLoPriRxCount + PacketByte(i): i = i + 1

3860                mLoPriTxCount = CLng(PacketByte(i)) * 256&: i = i + 1  'LoPriRxCount
3870                mLoPriTxCount = mLoPriTxCount + PacketByte(i): i = i + 1

3880                mLoPriDropCount = CLng(PacketByte(i)) * 256&: i = i + 1  'LoPriRxCount
3890                mLoPriDropCount = mLoPriDropCount + PacketByte(i): i = i + 1


3900                mHiPriRxCount = CLng(PacketByte(i)) * 256&: i = i + 1  'LoPriRxCount
3910                mHiPriRxCount = mHiPriRxCount + PacketByte(i): i = i + 1


3920                mHiPriTxCount = CLng(PacketByte(i)) * 256&: i = i + 1  'LoPriRxCount
3930                mHiPriTxCount = mHiPriTxCount + PacketByte(i): i = i + 1


3940                mHiPriDropCount = CLng(PacketByte(i)) * 256&: i = i + 1  'LoPriRxCount
3950                mHiPriDropCount = mHiPriDropCount + PacketByte(i): i = i + 1

3960                mNeighborBuildCount = CLng(PacketByte(i)) * 256& * 256&: i = i + 1  'LoPriRxCount
3970                mNeighborBuildCount = mNeighborBuildCount + PacketByte(i): i = i + 1

3980                mStat1 = PacketByte(i): i = i + 1
3990                mStat0 = PacketByte(i): i = i + 1
4000                mLevel = PacketByte(i): i = i + 1
4010                mMargin = PacketByte(i): i = i + 1

4020                mCheckSum = PacketByte(i)



4030                mLayerAssigned = IIf(mStat1 And &H80, 1, 0)
4040                mOneNeighbor = IIf(mStat1 And &H40, 1, 0)
4050                mAck = IIf(mStat1 And &H20, 1, 0)  ' RESPONSE TO CONFIGURATION MESSAGE
                    'unused = IIf(mStat1 And &H10, 1, 0)
4060                mLayerID3 = IIf(mStat1 And &H8, 1, 0)
4070                mLayerID2 = IIf(mStat1 And &H4, 1, 0)
4080                mLayerID1 = IIf(mStat1 And &H2, 1, 0)
4090                mLayerID0 = IIf(mStat1 And &H1, 1, 0)

4100                mJammed = IIf(mStat0 And &H80, 1, 0)
4110                mLowBatt = IIf(mStat0 And &H40, 1, 0)
4120                mTamper = IIf(mStat0 And &H20, 1, 0)
4130                mNoChange = IIf(mStat0 And &H10, 1, 0)
4140                mReset = IIf(mStat0 And &H8, 1, 0)
4150                mLineLoss = IIf(mStat0 And &H2, 1, 0)

                    'mRegistering = mReset
                    'dbg "DN registration mACK = " & mAck




                    'dbg "New Netwerk Repeater check-in " & Me.HexPacket()


                    ' ************** UNKNOWN DEVICE **************
4160              Case Else
                    ' Invalid
4170                mBADPacket = True
4180            End Select           ' mPTI


4190          Case Else              ' mClass
4200            If mClass = &H18 Then
4210              If mPTI = &HC Then
4220                Debug.Print mSerial, HexPacket





4230              End If
4240            End If

4250            dbg "UNK Class " & Hex(mClass)
4260        End Select

4270      Case Else                  'mTransType
4280        dbg "UNK TransType " & Hex(mTransType)

4290    End Select
4300    mSerial = Right("00" & Hex(mOriginatorMID), 2) & Right("000000" & Hex(mOriginatorUID), 6)

Bytes_Resume:
4310    On Error GoTo 0
4320    Exit Property

Bytes_Error:
4330    mBADPacket = True
4340    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cESPacket.Bytes." & Erl
4350    Resume Bytes_Resume

End Property

Public Property Get Checksum() As Integer
  Checksum = mCheckSum
End Property

Public Property Get PacketSequence() As Long
  PacketSequence = mPacketSequence
End Property
Private Sub Class_Initialize()
  TimeStamp = Now
  mPacketSequence = GetNextPacketSeq()
  Set SubPackets = New Collection
  Set Partions = New Collection
End Sub

Public Property Get CMD() As Byte
  CMD = mCmd
End Property

Public Property Get DataChanged() As Integer
  DataChanged = IIf(Stat0 And &H10, 1, 0)
End Property

Public Property Get Disarm() As Integer
  Disarm = IIf(Button And 2, 1, 0)
End Property

Public Property Get FirstHopMID() As Long
  FirstHopMID = mFirstHopMID
End Property

Public Property Get FirstHopSerial() As String
  FirstHopSerial = Right("00" & Hex(mFirstHopMID), 2) & Right("000000" & Hex(mFirstHopUID), 6)
End Property

Public Property Get FirstHopUID() As Long
  FirstHopUID = mFirstHopUID
End Property

Public Property Get HexPacket() As String
    
  
  Dim j As Integer
  Dim s As String
  
  If Is6080 Then
    mHexPacket = i6080data
  Else
  
  If PacketType = 0 Then ' otherwise serial packet
    If Len(mHexPacket) = 0 Then
      For j = 0 To mLastByte
        s = s & Right("00" & Hex(PacketByte(j)), 2)
      Next
      mHexPacket = s
    End If
    HexPacket = mHexPacket
  Else
    HexPacket = SerialPacket
  End If
  End If
End Property

Public Property Get HOPs() As Byte
  HOPs = mHOPs
End Property

Public Property Get IsLocator() As Integer ' repeater doubles as locator
  IsLocator = IIf(mOriginatorMID = 1, 1, 0)
End Property

Public Function LEvel() As Byte

'1530    ESDeviceType(j).Model = "EN1223S"
'1540    ESDeviceType(j).desc = "One Button Pendant"
'1550    ESDeviceType(j).MID = &HB2   '3E class   18 pti
'1560    ESDeviceType(j).PTI = &H18
'1570    ESDeviceType(j).CLS = &H3E


'1660    ESDeviceType(j).Model = "EN1221S-60"
'1670    ESDeviceType(j).desc = "One Button Pendant"
'1680    ESDeviceType(j).MID = &HB2   '3E class   1D pti
'1690    ESDeviceType(j).PTI = &H1D
'1700    ESDeviceType(j).CLS = &H3E


  'If CLSPTI() = &H3E18 Then
  If CLSPTI() = &H3E1D Then
    LEvel = mLevel * (1 + (Configuration.boost / 100))
  Else
    LEvel = mLevel
  End If
End Function

Public Function Margin() As Byte
  Margin = mMargin
End Function
Public Property Let MID(ByVal Value As Long)
  mOriginatorMID = Value
End Property


Public Property Get MID() As Long
  MID = mOriginatorMID
End Property

Public Property Get MIDPTI() As Long
  MIDPTI = mOriginatorMID * 256& + mPTI
End Property

Public Property Get Payload() As String
  'Payload = mPayload
End Property
Public Property Let PTI(ByVal Value As Byte)
  mPTI = Value
End Property

Public Property Get PTI() As Byte
  PTI = mPTI
End Property

Public Property Get RangeCheck() As Boolean
  RangeCheck = mRangeCheck
End Property

Public Property Get Reset() As Byte
  Reset = mReset
End Property

Public Property Let Reset(Value As Byte)
  mReset = Value
End Property


Public Property Get ResponseCode() As Byte
  ResponseCode = mResponseCode
End Property

Public Property Get Packet2Hex(Bytes() As Byte) As String
  Dim j As Integer
  Dim s As String
  Dim LastByte As Long
  LastByte = UBound(Bytes)
  For j = 0 To LastByte
    s = s & Right("00" & Hex(Bytes(j)), 2)
  Next
  Packet2Hex = s
End Property


Public Property Let ResponseCode(ByVal ResponseCode As Byte)
  mResponseCode = ResponseCode
End Property

Public Property Get Sequence() As Long
  Sequence = mSequence
End Property

Public Property Let Sequence(ByVal Sequence As Long)
  mSequence = Sequence
End Property

Public Property Let Serial(ByVal Value As String)
  mSerial = Value
End Property

Public Property Get Serial() As String
  Serial = mSerial
End Property

Public Function Stat0() As Byte
  Stat0 = mStat0
End Function

Public Function Stat1() As Byte
  Stat1 = mStat1
End Function

Public Property Get Status() As Long
  Status = mStat1 * 256& + mStat0
End Property

Public Property Get Tamper() As Integer
  Tamper = mTamper
End Property

Public Property Get TraceCount() As Byte
  TraceCount = mTraceCount
End Property

Public Property Get UnACKCount() As Byte
  UnACKCount = mUnACKCount
End Property

Public Function ValidateChecksum() As Boolean
  Dim j As Integer
  Dim Checksum As Integer: Checksum = 0
  For j = 0 To mLastByte - 1
    Checksum = (Checksum + PacketByte(j)) And &HFF&
  Next
  ValidateChecksum = (Checksum = PacketByte(mLastByte))
End Function
Function MakeCheckSum(DataBytes() As Byte) As Byte
  Dim j As Integer
  Dim CheckDigit As Integer
  For j = LBound(DataBytes) To UBound(DataBytes) - 1
    CheckDigit = (CheckDigit + DataBytes(j)) And &HFF&
  Next
  MakeCheckSum = (CheckDigit And &HFF)

End Function
Public Property Let CLSPTI(ByVal Value As Long)
 mClass = Value \ 256&
 mPTI = Value And &HFF
 
 
End Property



Public Property Get CLSPTI() As Long
 CLSPTI = CLng(mClass) * 256& + mPTI
End Property

Public Property Let ClassByte(ByVal Value As Byte)
  mClass = Value
End Property


Public Property Get ClassByte() As Byte
  ClassByte = mClass
End Property

Public Property Get SequenceID() As Long
  SequenceID = mSequence
End Property

Public Property Get Registering() As Integer

  Registering = mRegistering

End Property

Public Property Let Registering(ByVal Registering As Integer)

  mRegistering = Registering

End Property

Public Property Get SerialMessage() As String
  SerialMessage = mSerialMessage
End Property

Public Property Let SerialMessage(ByVal SerialMessage As String)
  mSerialMessage = SerialMessage
End Property

Public Property Get Temperature0() As Single

  Temperature0 = mTemperature0

End Property

Public Property Let Temperature0(ByVal Value As Single)

  mTemperature0 = Value

End Property

Public Property Get Temperature1() As Single

  Temperature1 = mTemperature1

End Property

Public Property Let Temperature1(ByVal Value As Single)

  mTemperature1 = Value

End Property

Public Property Get Alarm2() As Long

  Alarm2 = mAlarm2

End Property

Public Property Let Alarm2(ByVal Alarm2 As Long)

  mAlarm2 = Alarm2

End Property

Public Property Get LinkedAlarm() As Long

  LinkedAlarm = mLinkedAlarm

End Property

Public Property Let LinkedAlarm(ByVal LinkedAlarm As Long)

  mLinkedAlarm = LinkedAlarm

End Property
