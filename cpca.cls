VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Serial As String

Public Messages As Collection
Public Groups  As Collection  '6 groups per PCA, 32 groups per system, group IDs from 1 to 60
'If first group is xFE, then responds to ALL groups



Public DateTime
Public FirmwareVersion


Public CodedPageData        As Collection  ' 64 per PCA, NO IDs          cCodedPageData
Public CodedResponseData    As Collection  ' 32 per PCA, IDs 1 to 250    cCodedResponseString
Public CodedFieldMessages   As Collection  ' 32 per PCA, IDs 1 to 250    cCodedResponseString
Public GroupCodes           As Collection  ' 32 per PCA, IDs 1 to 63     cCodedResponseString

Public HostLinkDownMessage  As String  ' 31 char or two lines
Public MessageUndeliverable As String  ' 31 char or two lines
Public ResponseUndelivered  As String  ' 31 char or two lines

Public PageMessages         As Collection

Public Function Send(ByVal Address As String, ByVal Message As String, Optional ByVal Alert As Boolean) As Boolean
  Dim pCancel As Long
  Dim NewMessage  As cWirelessMessage
  pCancel = InStr(1, Message, "cancel", vbTextCompare)
  If pCancel > 2 Then
    If pCancel = 1 + Len(Message) - Len("Cancel") Then
      Message = left(Message, pCancel - 2)  ' 1 based, plus space
      Message = "Cancel  " & Message
    End If
  End If

  'Message = CreatePageMessage(Address, Message, 0)
  Outbounds.AddPreparedMessage CreatePageMessage(Address, Message, 0)

  'Outbounds.AddMessageString Message  ' Address, &H28, text, getnextid
End Function
Private Function CreatePageMessage(ByVal Serial As String, ByVal Message As String, ByVal Alert As Boolean) As cWirelessMessage

  Dim Msglen        As Integer  ' up to 90
  Dim NeedACK       As Boolean
  Dim DeliveryCode  As Byte
  Dim Buffer        As String
  Dim MessageID     As String
  Dim SequenceID    As Long
  Dim NewMessage    As cWirelessMessage

  NeedACK = True
  Buffer = "500000" & Serial  'Header, Length placeholder, Radiotype and serial
  Buffer = Buffer & "19"  ' PCA class byte
  DeliveryCode = DeliveryCode Or IIf(NeedACK, BIT_7, 0)
  Buffer = Buffer & Right("00" & Hex(DeliveryCode), 2)
  Buffer = Buffer & "26"
  MessageID = GetNextMessageID()
  SequenceID = Val("&h" & MessageID)
  Buffer = Buffer & MessageID
  '  [Message content]  Application-specific information. Maximum message content length is 85 bytes (90 - 5).
  If Len(Message) > 62 Then
    Message = left(Message, 62) ' chop off at 62 bytes
  End If
  Buffer = Buffer & StringToHex(Message & vbNullChar)
  Buffer = Buffer & "00" ' no reponse string

  ' get message length
  Msglen = Len(Buffer) / 2
  Mid(Buffer, 3, 2) = Right("00" & Hex(Msglen), 2)

  ' append checksum and return it all
  Message = Buffer & HexChecksum(Buffer)

  Set NewMessage = New cWirelessMessage
  NewMessage.MessageData = Message
  NewMessage.TimeStamp = DateAdd("s", -1, Now)  ' no delay
  NewMessage.NeedsAck = True
  NewMessage.SequenceID = SequenceID
  
  Set CreatePageMessage = NewMessage
  Set NewMessage = Nothing



End Function

'OUT bound message codes
'x12 Application ACK
'x20 Configure NID
'x22 Coded Page/Coded Repsonses
'x24 Coded Page/Custom Repsonses
'x26 Custom Page/Coded Repsonses
'x28 Custom Page/Custom Repsonses
'x40 Update date/time
'x41 Delete all messages in log
'x42 Delete Specific message in log
'x44 Request PCA Status
'x46 Query Message Status
'x48 Force Alert Activation
'x51 Configure PCA Options 86
'x52 Configure PCA Parameters 87
'x56 Request Configuration 87
'x57 Set Group Membership 88
'x58 Request Group Membership List 88
'x59 Request Date/Time 89
'x5E Request Firmware Version 89
'x61 Clear Coded Data 89
'x62 Load Coded Page Data 90
'x63 Request List of Page Codes 90
'x64 Load Coded Response Data 90
'x65 Request List of Response Codes 91
'x66 Load Coded Field Messages 91
'x67 Request List of Field Message Codes 92
'x68 Load Coded Group List 92
'x69 Request List of Group Codes 92
'x6A Load Network Message Strings 93
'x6B Request List of Network Message Codes 93
'x6E Request Message Log 93


'IN bound message codes
'x00 Registration 78
'x11 0x11 Node ACK 97
'x13 0x13 System ACK 97
'x17 0x17 Supervisory Message 96
'x23 0x23 Response to Paging Message 97
'x25 0x25 Field Message 96
'x26 0x26 Range Check 97
'x5A 0x5A User Changed Configuration Message 97
'x5C 0x5C User Deleted Message 98
'x5D 0x5D User Cleared Message Log 98

Private Sub Class_Initialize()
  Set Messages = New Collection
End Sub



'Private Function Setgroupmembership() As String
''[0x57] - Set group membership command.
''[MsgSeq (2)] Application controller generated unique sequence number for this message.
''[Message content]
''[Count n]
''[Array n of group codes (up to six) to which the PCA should respond]
'
'End Function
'

Function BroadcastMessage(ByVal commandbyte As Integer)

  Dim Message       As String
  Dim Msglen        As Integer  ' up to 90
  Dim NeedACK       As Boolean
  Dim DeliveryCode  As Byte
  Dim Sequence      As Long


  '[0x50]  Header for broadcast outbound message.
  '[LEN]  Message length, excluding checksum.
  '[0x00] - Radio type: Enhanced two-way
  '[UID (4) dest]  The unique ID of the destination node.
  Message = "500000" & Serial  'Header, Length placeholder, Radiotype and serial

  'Payload: The maximum length of the payload is 90 bytes, including all message bytes below.
  '[0x19]  Message class byte for messages to the PCA.
  Message = Message & "19"  ' PCA class byte

  '[Delivery Code]  Identifies specific protocol for message delivery.
  '  Bit 7 Set if Node ACK is requested.
  '  Bit 6 Reserved.
  '  Bits 0-5 Group code (0-63), to address message to a group of PCAs.
  DeliveryCode = DeliveryCode Or IIf(NeedACK, &HF0, 0)
  Message = Message & Right("00" & Hex(DeliveryCode), 2)

  '  [CMD] Command Byte.
  Message = Message & Right("00" & Hex(commandbyte), 2)
  '  [MsgSeq (2)]  Application controller generated two-byte, unique sequence number for this message.
  Sequence = GetNextMessageID()
  Message = Message & Right("0000" & Hex(Sequence), 4)


  '  [Message content]  Application-specific information. Maximum message content length is 85 bytes (90 - 5).
  '  [CKSUM]  Checksum

  ' get message length
  Msglen = Len(Message) / 2
  Mid(Message, 3, 2) = Right("00" & Hex(Msglen), 2)

  ' append checksum
  Message = Message & HexChecksum(Message)

End Function
