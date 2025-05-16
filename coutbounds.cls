VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mReady          As Boolean
Private mMessages       As Collection
Private mLastTimeSent   As Date

Private mSentMessages   As Collection

Public Sub ACK(Packet As cESPacket)
  Dim j As Integer
  Dim m As cWirelessMessage

  For j = 1 To mSentMessages.count
    Set m = mSentMessages(j)
    If m.SequenceID = Packet.SequenceID Then
      mSentMessages.Remove j
      dbgGeneral "Outbound Message ACK"
    Exit For
    End If
  Next
    
  
End Sub

Private Sub Class_Initialize()
  Set mMessages = New Collection
  Set mSentMessages = New Collection
  
  mLastTimeSent = Now
  mReady = True
End Sub
Public Property Get count() As Long
  count = mMessages.count
End Property
Public Property Get ResendCount() As Long
  ResendCount = mSentMessages.count
End Property
Public Function AddPreparedMessage(WirelessMessage As cWirelessMessage) As Long
  mMessages.Add WirelessMessage
  AddPreparedMessage = mMessages.count
End Function
Public Function AddMessageString(ByVal MessageString As String) As Long
  Dim NewMessage As cWirelessMessage
  Set NewMessage = New cWirelessMessage
  NewMessage.CreateMessageFromString MessageString
  mMessages.Add NewMessage
End Function
Public Function AddMessage(ByVal Serial As String, ByVal MessageType As Integer, ByVal Message As String, ByVal id As Integer) As Long
  Dim NewMessage As cWirelessMessage
  Dim m          As cWirelessMessage
  For Each m In mMessages
    If m.ComparesTo(Serial, MessageType, Message) Then
      Exit Function
    End If
  Next

  Set NewMessage = New cWirelessMessage
  NewMessage.CreateMessage Serial, MessageType, Message, id
  mMessages.Add NewMessage

End Function
Function CheckForResends() As Long
  Dim j  As Integer
  Dim m As cWirelessMessage
  If mSentMessages.count Then
    CheckForResends = mSentMessages.count
  End If
  If mReady Then
  
    For j = 1 To mSentMessages.count
      Set m = mSentMessages(j)
      If DateDiff("s", m.TimeStamp, Now) >= 0 Then
        mSentMessages.Remove j
        'DebugString "SendCount: " & m.SendCount
        If m.SendCount < MAX_PCA_RESENDS Then
          mMessages.Add m
          dbgGeneral "ReQue Message"
        Else
          dbgGeneral "Device Not Responding"  ' & m.SendCount
          ' log error too many resends
        End If
        Exit For
      End If
    Next
  End If


End Function
Sub SendMessages()
'typically called from modMain_DoRead
  Dim m As cWirelessMessage
  Dim j As Integer
  Dim ResendsPending As Integer
  
  
  
  ResendsPending = CheckForResends()
  
  If mReady Then

    For j = 1 To mMessages.count
      Set m = mMessages(j)
      
      If DateDiff("s", m.TimeStamp, Now) >= 0 Then
        mReady = False
        mMessages.Remove j
        m.SendCount = m.SendCount + 1
        If m.NeedsAck Then
          m.Reque
          mSentMessages.Add m
        End If
        DispatchMessage m.MessageData
        mLastTimeSent = Now
        'DebugString "sent outbound " & m.MessageData
        Exit For
      End If
    Next
  Else
    ' handle timeout reset of ready flag
   
    If DateDiff("s", mLastTimeSent, Now) >= OUTBOUND_TIMEOUT Then
      dbgGeneral "timeout/reset for NC Response"
      mReady = True
   
    End If

  End If
End Sub



Sub DispatchMessage(ByVal Message As String)
  Dim j As Integer
  Dim buffsize As Integer
  Dim Buffer() As Byte
  buffsize = Len(Message) \ 2 ' num bytes
  
  dbgGeneral "Dispatched Message ==> " & Message
  
  If buffsize > 0 Then
    If buffsize = 1 Then
       ReDim Buffer(0)
    Else
       ReDim Buffer(0 To buffsize - 1)
    End If
    For j = 0 To buffsize - 1
      Buffer(j) = Val("&h" & MID(Message, j * 2 + 1, 2))
    Next
    WirelessPort.CommWriteByteArray Buffer, buffsize
   
  End If
 
  
End Sub

Public Property Get Ready() As Boolean
  Ready = mReady
End Property

Public Property Let Ready(ByVal Value As Boolean)
  mReady = Value
End Property
