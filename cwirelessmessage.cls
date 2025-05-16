VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWirelessMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mSerial         As String
Private mDeviceType     As Long
Private mCommand        As Long
Private mMessage        As String  ' the packetdata as a hex string
Private mTimeStamp      As Date
Private mPacket         As cESPacket
Private mMessageType    As Integer
Private mMessageData    As String
Private mSendCount      As Long
Private mNeedsAck       As Boolean
Private mResendDelayMs  As Long ' milliseconds
Private mACKnowledged   As Boolean
Private mSequenceID     As Long

Const ReQueDelay = 1

Public Sub Reque()
  TimeStamp = DateAdd("s", ReQueDelay, Now)
End Sub

Public Function ComparesTo(ByVal Serial As String, ByVal MessageType As Integer, ByVal Message As String) As Boolean
  ComparesTo = (mSerial = Serial) And (mMessageType = MessageType) And (mMessage = Message)
End Function

Public Function CreateMessageFromString(ByVal MessageString As String) As String
  mMessageData = MessageString
  TimeStamp = DateAdd("s", -1, Now)  ' No Delay
End Function

Public Function CreateMessage(ByVal Serial As String, ByVal MessageType As Integer, Message As String, ByVal id As Integer) As String
  
  'Note: to send a messages immediately, place a -1 in the delay (see CreateCannedACK)
  mSerial = Serial
  mMessageType = MessageType
  mMessage = Message
  Select Case MessageType
    
    
    
    Case MSGTYPE_SETNID
      mMessageData = SetNCNID(Message)
      TimeStamp = DateAdd("s", -1, Now)  ' No Delay
      
    Case MSGTYPE_CANNEDACK
      mMessageData = CreateCannedACK(Serial, Message)
      TimeStamp = DateAdd("s", -1, Now)  ' No Delay
    Case MSGTYPE_CREATEFIELDMSG
      mMessageData = CreateFieldMessage(Serial, id, Message)
      TimeStamp = DateAdd("s", 1, Now) ' send 1 seconds from now
    Case MSGTYPE_TWOWAYNID
      mMessageData = ConfigureTwoWayNID(Serial)
      TimeStamp = DateAdd("s", 1, Now)  ' send 1 second from now
    
    Case MSGTYPE_REPEATERNID
      mMessageData = ConfigureTwoWayNID(Serial)
      TimeStamp = DateAdd("s", 1, Now)  ' send 1 second from now
    
    Case MSGTYPE_REQTXSTAT
      mMessageData = RequestTXStatus()
      TimeStamp = DateAdd("s", 1, Now)  ' send 1 second from now
      
    Case MSGTYPE_GETNID
      
      mMessageData = RequestTXNID()
      TimeStamp = DateAdd("s", 1, Now)  ' send 1 second from now
      
    Case MSGTYPE_SET_TIME
      mMessageData = CreateSetTimeMsg(Serial)
      TimeStamp = DateAdd("s", 1, Now)  ' send 1 second from now
    Case MSGTYPE_DELETEALL
      mMessageData = CreateDeleteAll(Serial)
      TimeStamp = DateAdd("s", 1, Now)  ' send 1 second from now
    Case MSGTYPE_CUSTOM_CUSTOM
    Case MSGTYPE_GETNCSERIAL
      mMessageData = ReportNCSerial
      TimeStamp = DateAdd("s", -1, Now)
    Case MSGTYPE_GENERIC_PAGE
      mMessageData = CreatePageMessage(Serial, Message, False)
      TimeStamp = DateAdd("s", -1, Now)  ' no delay
      NeedsAck = (Val("&h" & mid(mMessageData, 17, 2)) And (BIT_7)) = (BIT_7)
      SequenceID = Val("&h" & mid(mMessageData, 21, 4))
  End Select
  Message = mMessageData
End Function



Public Function RepeaterAssignNeighbors(ByVal Neighbor1 As String, ByVal Neighbor2 As String, ByVal Layer As Integer) As String

  Dim Header          As String
  Dim Msglen          As String
  Dim RadioType       As String
  Dim UID             As String
  Dim MsgClass        As String
  Dim SubClass        As String
  Dim DeliveryCode    As String
  Dim commandbyte     As String
  Dim Content         As String
  Dim Message         As String
  Dim numbytes        As Integer


  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "40"  ' Repeater Config
  SubClass = "0F"  ' Subclass byte to assign neighbors to repeater
  

  Message = Header & Msglen & RadioType & UID & MsgClass & SubClass & Neighbor1 & Neighbor2 & Right("00" & Hex$(Layer), 2)
  
  numbytes = Len(Message) / 2
  Mid(Message, 3, 2) = Right("00" & Hex$(numbytes), 2)  ' insert message length

  RepeaterAssignNeighbors = Message & HexChecksum(Message)


'[0x50]  Header for Broadcast Outbound message
'[LEN]  Length. Total number of bytes in this message excluding checksum.
'[radio type] = 0x00
'[SN(4) dest]
'[payload] =
  '[0x40] = Message class byte for repeater configuration message
  '[0x0F] = Subclass byte to assign neighbors to repeater
  '[UID1] = UID of first neighbor (4 bytes)
  '[UID2] = UID of second neighbor (4 bytes)
  '[LAYER] = New layer of repeater (1 byte)
'[CKSUM]

End Function

Private Function RepeaterNeighborRebuiid(ByVal Serial As String)

'7.2.2 Force Neighbor Rebuild
'When the repeater receives this message it will rebuild its neighbor list using its
'currently assigned layer.


  Dim Header          As String
  Dim Msglen          As String
  Dim RadioType       As String
  Dim UID             As String
  Dim MsgClass        As String
  Dim SubClass        As String
  Dim DeliveryCode    As String
  Dim commandbyte     As String
  Dim Content         As String
  Dim Message         As String
  Dim numbytes        As Integer


  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "40"  ' Repeater Config
  SubClass = "0E"  ' Force Neighbor Rebuild

  Message = Header & Msglen & RadioType & UID & MsgClass & SubClass
  
  numbytes = Len(Message) / 2
  Mid(Message, 3, 2) = Right("00" & Hex(numbytes), 2)  ' insert message length

  RepeaterNeighborRebuiid = Message & HexChecksum(Message)


  

'[0x50]  Header for Broadcast Outbound message
'[LEN]  Length. Total number of bytes in this message excluding checksum.
'[radio type] = 0x00
'[SN(4) dest]
'[payload] =
  '[0x40] = Message class byte for repeater configuration message
  '[0x0E] = Subclass byte to force repeater neighbor rebuild
'[CKSUM]


End Function

Public Property Set Packet(Packet As cESPacket)
  Set mPacket = Packet
End Property
Public Property Get Packet() As cESPacket
  Set Packet = mPacket
End Property

Private Sub Class_Initialize()
  Set Packet = New cESPacket
  
End Sub

Public Property Get Serial() As String
  Serial = mSerial
End Property

Public Property Let Serial(ByVal Serial As String)
  mSerial = Serial
End Property

Public Property Get Message() As String
  Message = mMessage
End Property

Public Property Let Message(ByVal Message As String)
  mMessage = Message
End Property

Public Property Get TimeStamp() As Date
  TimeStamp = mTimeStamp
End Property

Public Property Let TimeStamp(ByVal Value As Date)
  mTimeStamp = Value
End Property

Public Property Get DeviceType() As Long
  DeviceType = mDeviceType
End Property

Public Property Let DeviceType(ByVal DeviceType As Long)
  mDeviceType = DeviceType
End Property

Public Property Get Command() As Long
  Command = mCommand
End Property

Public Property Let Command(ByVal Command As Long)
  mCommand = Command
End Property

Public Property Get MessageType() As String
  MessageType = mMessageType
End Property

Public Property Let MessageType(ByVal MessageType As String)
  mMessageType = MessageType
End Property

Public Property Get MessageData() As String
  MessageData = mMessageData
End Property

Public Property Let MessageData(ByVal MessageData As String)
  mMessageData = MessageData
End Property

Public Property Get SendCount() As Long
  SendCount = mSendCount
End Property

Public Property Let SendCount(ByVal SendCount As Long)
  mSendCount = SendCount
End Property

Public Property Get NeedsAck() As Boolean
  NeedsAck = mNeedsAck
End Property
Public Property Let NeedsAck(ByVal NeedsAck As Boolean)
  mNeedsAck = NeedsAck
End Property

Public Property Get ResendDelayMs() As Long
  ResendDelayMs = mResendDelayMs
End Property

Public Property Let ResendDelayMs(ByVal ResendDelayMs As Long)
  mResendDelayMs = ResendDelayMs
End Property

Public Property Get ACKnowledged() As Boolean
  ACKnowledged = mACKnowledged
End Property

Public Property Let ACKnowledged(ByVal ACKnowledged As Boolean)
  mACKnowledged = ACKnowledged
End Property

Public Property Get SequenceID() As Long
  SequenceID = mSequenceID
End Property

Public Property Let SequenceID(ByVal SequenceID As Long)
  mSequenceID = SequenceID
End Property
