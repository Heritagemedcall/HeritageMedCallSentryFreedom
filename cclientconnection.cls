VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClientConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' remote ack just needs to send Alarm type (Alarm, Alert, etc) and
' key: Serial and inputnum

' Host acts as much like a web service / soap but terse
' Header is action (post) followed by ### bytes in the message

' Client is the consumer of the data returned

Public LastError        As String
Public LastErrorCode    As Long
Public Header           As String


Private Waiting         As Boolean
Private BytesNeeded     As Long
'Private StartTimer      As Long
Private mReadyState     As Integer
Private mResponseText   As String
Private FastString      As String '
Private starttime       As Date

'Public WithEvents Socket As Winsock
Public Socket As Winsock

Public Property Get Closed() As Boolean
  Closed = (Socket.State = sckClosed)
End Property
Function AssignSocket(NewSocket As Winsock)
  If Not (Socket Is Nothing) Then
    Socket.Close
  End If

  Set Socket = NewSocket

End Function


Public Property Get RequestPending() As Boolean
  RequestPending = (mReadyState = 4)
End Property
Public Property Get Request() As String
  'request is XML Doc as text
  Request = mResponseText
  mResponseText = ""
  FastString = "" '.Clear
  mReadyState = 1
End Property

Public Property Get ResponseSent() As Boolean
' indicates that response has been created and
' response data is transmitted and
' outgoing buffer is empty and ready
' might be redundant
End Property

Public Function Respond(ByVal Response As String, ByVal Status As Integer) As Long
      ' response to send... if no-one is in the forest to hear the tree fall... so be it
  
        Dim statusword As String
        Dim packet As String
  
10       On Error GoTo Respond_Error

20      'Debug.Print "Respond "; response

  
30      packet = Join(Array("response", CStr(Status), CStr(Len(Response))), " ") & vbCrLf
40      packet = packet & Response
  
50      Socket.SendData packet

60      mReadyState = 1
70      Waiting = True
  
        'dbg "Respond response sent"

Respond_Resume:
80       On Error GoTo 0
90       Exit Function

Respond_Error:

100     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cClientConnection.Respond." & Erl
110     Resume Respond_Resume

  
End Function

Private Sub Class_Initialize()
  
  Waiting = True
End Sub

Public Sub CloseConnection()
  If Not (Socket Is Nothing) Then
    mReadyState = 0
    Socket.Close
  End If
End Sub

Private Sub Socket_Close()
  mReadyState = 0
  FastString = "" '.Clear
  mResponseText = ""
  
End Sub

Public Sub Socket_Connect()
  mReadyState = 1
  dbg "cClient Socket_Connect() " & Socket.RemoteHostIP & " " & Socket.RemotePort
End Sub

Public Sub Socket_ConnectionRequest(ByVal requestID As Long)
  ' nada
End Sub

Public Sub Socket_DataArrival(ByVal bytesTotal As Long)
        Dim DataBuffer As String
        Dim HeaderBytes As Long
        'bytes in buffer = socket.BytesReceived

10       On Error GoTo Socket_DataArrival_Error

20      'Debug.Print "cClient Socket_DataArrival " & vbCrLf & bytesTotal
30      Socket.PeekData DataBuffer, vbString
40      'Debug.Print "cClient Socket_DataArrival.DataBuffer " & vbCrLf & DataBuffer

50      DataBuffer = ""

60      If mReadyState = 2 Then
70        mReadyState = 3
80      End If

90      If mReadyState = 1 Then
100       mReadyState = 2
110     End If
120     'Debug.Print "Readystate "; mReadyState
130     'Debug.Print "Waiting "; Waiting
  
140     If Waiting Then
150       Socket.PeekData DataBuffer, vbString  ' peekdata gets data but does not remove it
    
160       HeaderBytes = Checkdata(DataBuffer)
170       If HeaderBytes > 0 Then
180         Socket.GetData DataBuffer, vbString, HeaderBytes  ' get just the header including \r\n
190         bytesTotal = bytesTotal - HeaderBytes
200         Waiting = False
210         'StartTimer = Win32.timeGetTime + 2000
            starttime = DateAdd("s", 4, Now)
220       End If
230       Socket.GetData DataBuffer, vbString  ' , optional maxlength (do not use maxlength!)
240       FastString = FastString & DataBuffer  ' .Append DataBuffer

250       BytesNeeded = MaxL(0, BytesNeeded - bytesTotal)
260       If BytesNeeded <= 0 Then
270         mResponseText = FastString '.ToString
280         mReadyState = 4
290       End If

300     Else
310       If starttime < Now Then
            ' abort
320         mResponseText = GetTimeoutResponse()
330         mReadyState = 4
340         Exit Sub
350       End If
360       Socket.GetData DataBuffer, vbString  ' , optional maxlength (do not use maxlength!)
370       FastString = FastString & DataBuffer '.Append DataBuffer
380       BytesNeeded = MaxL(0, BytesNeeded - bytesTotal)
390       If BytesNeeded <= 0 Then
400         mResponseText = FastString '.ToString
410         mReadyState = 4
420       End If
430     End If

Socket_DataArrival_Resume:
440      On Error GoTo 0
450      Exit Sub

Socket_DataArrival_Error:

460     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cClientConnection.Socket_DataArrival." & Erl
470     Resume Socket_DataArrival_Resume


End Sub
Public Function GetTimeoutResponse() As String
  Dim timoutstring As String
  timoutstring = "<?xml version=""1.0""?><HMC revision=""" & App.Revision & """><error>Timeout</error><errorcode>500</errorcode><module>ClientConnection.GetTimeoutResponse</module></HMC>"
  dbg "GetTimeoutResponse " & timoutstring & vbCrLf
  GetTimeoutResponse = timoutstring

End Function

Public Property Get getAllResponseHeaders() As String
  ' does not remove messages
End Property

Public Property Get ResponseReady() As Boolean
  ResponseReady = (mReadyState = 4)
End Property
Public Property Get ResponseText() As String
  'Clears response text and faststring
  'resets readystate to 1
  'request is XML
  'FastString.Clear
  FastString = ""
  ResponseText = mResponseText
  mResponseText = ""

  mReadyState = 1
  Waiting = True
End Property

'Public Function Send(ByVal Method As String, ByVal Request As String) As Long
'
'        Dim packet As String
'
'10       On Error GoTo Send_Error
'
'20      packet = Join(Array(Method, CStr(Len(Request))), " ") & vbCrLf
'30      packet = packet & Request
'40      Socket.SendData packet
'50      mReadyState = 2
'
'Send_Resume:
'60       On Error GoTo 0
'70       Exit Function
'
'Send_Error:
'
'80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cClientConnection.Send." & Erl
'90      Resume Send_Resume
'
'End Function

Public Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  LastErrorCode = Number
  LastError = Description

End Sub

Public Function Checkdata(ByVal Header As String) As Long

        Dim Ptr As Long
        Dim ParsedHeaders() As String

10       On Error GoTo Checkdata_Error

20      Ptr = InStr(Header, vbCrLf)
30      If Ptr > 0 And Ptr < 20 Then  ' max header is 19 chars
40        Header = left(Header, Ptr)
50        ParsedHeaders() = Split(Header, " ")
60        Select Case LCase(ParsedHeaders(0))
            Case "post"
70            BytesNeeded = Val(ParsedHeaders(1))
80          Case "get"
90            BytesNeeded = Val(ParsedHeaders(1))
100         Case Else
110       End Select
120       Checkdata = Ptr + 1
130     End If

Checkdata_Resume:
140      On Error GoTo 0
150      Exit Function

Checkdata_Error:

160     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cClientConnection.Checkdata." & Erl
170     Resume Checkdata_Resume


End Function
'winsock access wrappers
Public Function Accept(ByVal requestID As Long) As Long
  Socket.Accept requestID
End Function

Public Sub Bind(ByVal Port As Long, ByVal IP As String)
  'Use this for multiple network cards
  Socket.Bind Port, IP
End Sub

Public Function Connect(ByVal Host As String, ByVal Port As String) As Boolean
  Socket.Connect Host, Port
End Function

Public Function Listen()
  Socket.Listen
End Function

Public Property Get LocalPort() As Long
  LocalPort = Socket.LocalPort
End Property

Public Property Let LocalPort(ByVal Port As Long)
  Socket.LocalPort = Port
End Property

Public Property Get LocalIP() As String
  ' read only
  LocalIP = Socket.LocalIP
End Property

Public Property Get RemotePort() As Long
  RemotePort = Socket.RemotePort
End Property

Public Property Let RemotePort(ByVal Port As Long)
  Socket.RemotePort = Port
End Property

Public Property Get RemoteIP() As String
  ' read only
  RemoteIP = Socket.RemoteHostIP
End Property

Public Property Get State() As Integer
  State = Socket.State
End Property

Public Property Get RemoteHost() As String
  RemoteHost = Socket.RemoteHost
End Property

Public Property Let RemoteHost(ByVal Host As String)
  Socket.RemoteHost = Host
End Property
Public Sub Socket_SendComplete()
  Debug.Print "Response Socket_SendComplete"
End Sub

Public Sub Socket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  Debug.Print "Response Socket_SendProgress bytesSent "; bytesSent
End Sub

Private Sub Class_Terminate()
   CloseConnection
End Sub
