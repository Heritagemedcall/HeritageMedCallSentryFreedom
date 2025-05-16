VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerialTAPMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mState          As Integer
Private ID_Buffer       As String
Private MSG_Buffer      As String
Private Checksum_Buffer As String
Private Request_Buffer  As String
Private Buffer          As String

'Timing Params

Const RequestForID = "ID="

'conversation:
'Rx:<CR>
'Tx:ID=
'Rx:<ESC>PC1<CR>
'Tx:<CR><LF><110 1.8 TAP In Ready<CR>ACK><CR><ESC>[p<CR>
'Rx:<STX>0770100<CR>TEST PAGE<CR><ETX>37;<CR>
'Tx:211 Message sent.<CR><ACK><CR>
'Rx:<EOT><CR>
'Tx:115 Goodbye.<CR><ESC><EOT><CR>

'
'<Message sequence>
'<CR><ACK><CR>        'Logon accepted
'  or
'<Message sequence>
'<CR><NAK><CR>        'Requested again
'  or
'<Message sequence>
'<CR><ESC><EOT><CR>   'Forced disconnect

'The Paging server will do this when no RX is happening
' <CR> ' 1 sec
' <CR> ' 1 sec
' <CR> ' 1 sec
' <CR> ' 1 sec
' <EOT><CR> ' 4 sec

' throws error 101





Const T1 = 2
Const T2 = 1
Const T3 = 10
Const T4 = 4
Const T5 = 8

'     t1 -  2 secs.
'     t2 -  1 sec.
'     t3 - 10 secs.
'     t4 -  4 secs.
'     t5 -  8 secs.
'Retry Parameters

Const N1 = 3
Const N2 = 3
Const N3 = 3

'     n1 -3
'     n2 -  3 (undefined in rev. 1.0)
'     n3 -  3 (undefined in rev. 1.0)



' cheksum is on stx thru etx inclusive
'<STX>ID<CR>Message<CR><ETX>Checkum<CR>




Private mStarted         As Boolean
Private mEndTime         As Date


Private Const TIMEOUTDELAY = 10 ' ten seconds
Public Property Get Timeout() As Boolean
  If mStarted Then
    Timeout = (Now > mEndTime)
  End If
End Property


Public Property Get Packet() As String
  Packet = Buffer
End Property
Public Property Get Message() As String
  Message = MSG_Buffer
End Property
Public Property Get PagerNum() As String
  PagerNum = ID_Buffer
End Property



Public Sub clear()
  mStarted = False
  State = TAP_WAITING
  Buffer = ""
  ID_Buffer = ""
  MSG_Buffer = ""
  Checksum_Buffer = ""
  Request_Buffer = ""
End Sub

Public Property Get IsCheckSumValid() As Boolean
' calculate checksum on ID_Buffer
' compare to checksum
  Dim Message As String
  Message = Buffer ' Chr$(2) & ID_Buffer & vbCr & MSG_Buffer & vbCr & Chr$(3)
  IsCheckSumValid = VerifyTAPCheckSum(Message, Checksum_Buffer)

End Property

Public Sub BufferAddChar(ByVal b As Byte)
  Buffer = Buffer & Chr$(b)
End Sub

Public Sub IDAddChar(ByVal b As Byte)
  ID_Buffer = ID_Buffer & Chr$(b)
End Sub

Public Sub MessageAddChar(ByVal b As Byte)
  MSG_Buffer = MSG_Buffer & Chr$(b)
End Sub


Public Sub REQAddChar(ByVal b As Byte)
  Request_Buffer = Request_Buffer & Chr$(b)
End Sub


Public Sub ChecksumAddChar(ByVal b As Byte)
  Checksum_Buffer = Checksum_Buffer & Chr$(b)
End Sub

Function VerifyTAPCheckSum(ByVal Message As String, ByVal Checksum As String) As Boolean
  
  Dim j           As Long
  Dim Sum         As Long
  Dim CK1         As String
  Dim CK2         As String
  Dim CK3         As String
  
  
  For j = 1 To Len(Message)
    Sum = Sum + (Asc(mid$(Message, j, 1)) And &H7F)
  Next
  Sum = Sum And &HFFF&
  CK3 = Chr(48 + (Sum - (Sum \ 16) * 16))
  Sum = Sum \ 16
  CK2 = Chr(48 + (Sum - (Sum \ 16) * 16))
  Sum = Sum \ 16
  CK1 = Chr(48 + (Sum - (Sum \ 16) * 16))
  

  VerifyTAPCheckSum = (Checksum = CK1 & CK2 & CK3)

End Function


Public Property Get State() As Integer
  State = mState
End Property

Public Property Let State(ByVal Value As Integer)
  Select Case Value
  
  
  Case TAP_STATUS.TAP_ATTENTION
    mStarted = True
    mEndTime = DateAdd("s", TIMEOUTDELAY, Now)
  Case TAP_STATUS.TAP_TIMEOUT, TAP_STATUS.TAP_ERROR
    clear
    mStarted = False
  Case TAP_STATUS.TAP_HAS_STX
    mStarted = True
    mEndTime = DateAdd("s", TIMEOUTDELAY, Now)
  End Select
  mState = Value
End Property
