VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHostConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this connects TO host, lives on the Client

' remote ack just needs to send Alarm type (Alarm, Alert, etc) and
' key: Serial and inputnum

' Host acts as much like a web service / soap but terse
' Sent is a header and action (post) followed by ### bytes for the message body length

' This Client is the consumer of the data returned
' Header received is action (reponse) followed by ### bytes for the message body length


Public Socket As Winsock

Public LastError        As String
Public LastErrorCode    As Long

Public Header            As String


Private Waiting         As Boolean      ' waiting means that there is no pending inbound message
Private BytesNeeded     As Long         ' bytes needed are the body length without the header
'Private StartTimer      As Long         ' used for RX timeout
Private StartTimer      As Date         ' used for RX timeout
Private mResponseText   As String       ' Body of message
Private FastString      As String '   ' Collects incoming message body... faster that concatenation

Private Const TIMEOUT_DELAY = 4000      ' two seconds before giving up on response

Private mReadyState     As Integer      ' Works a lot like like HTTP object
  '0 (Uninitialized) The object has been created, but not initialized (the open method has not been called).
  '1 (Open) The object has been created, but the send method has not been called.
  '2 (Sent) The send method has been called, but the reply status and header(s) are not yet available.
  '3 (Receiving) Some data has been received. Calling the responseBody and responseText properties
      'at this state to obtain partial results will return an error,
      'because status and response headers are not fully available.
  '4 (Loaded) All the data has been received, and the complete data is available.
  '
  
Public Function GetWinsockError() As Long
  GetWinsockError = LastErrorCode
  LastErrorCode = 0
  LastError = ""
End Function

  
  
Private Sub Class_Initialize()
  Waiting = True
  
  
End Sub

Public Sub CloseConnection()
  If Not (Socket Is Nothing) Then
    Socket.Close
  End If
End Sub

Public Sub Socket_Close()
  CloseConnection
  mReadyState = 0
End Sub

Public Sub Socket_Connect()
  mReadyState = 1
End Sub

Public Sub Socket_ConnectionRequest(ByVal requestID As Long)
  ' nada
End Sub

Public Sub Socket_DataArrival(ByVal bytesTotal As Long)

        Dim DataBuffer As String
        Dim HeaderBytes As Long
        'bytes in buffer = socket.BytesReceived
         'On Error GoTo Socket_DataArrival_Error
10       On Error GoTo 0

20      If mReadyState = 2 Then
30        mReadyState = 3
40      End If
        'Debug.Print "Waiting, ReadyState " & Waiting & "  " & ReadyState
  
  
50      If Waiting Then
60        Socket.PeekData DataBuffer, vbString  ' peekdata gets data but does not remove it
70        HeaderBytes = Checkdata(DataBuffer)

          'Debug.Print "peeked data, HeaderBytes " & DataBuffer & "  " & HeaderBytes

80        If HeaderBytes > 0 Then
90          Socket.GetData DataBuffer, vbString, HeaderBytes  ' get just the header including \r\n
100         bytesTotal = bytesTotal - HeaderBytes
110         Waiting = False
120         StartTimer = DateAdd("s", TIMEOUT_DELAY, Now) 'Win32.timeGetTime + TIMEOUT_DELAY
130       End If

          

140       Socket.GetData DataBuffer, vbString  ' , optional maxlength (do not use maxlength!)
'150       FastString .Append DataBuffer
150       FastString = FastString & DataBuffer
          
160       BytesNeeded = MaxL(0, BytesNeeded - bytesTotal)
170       If BytesNeeded <= 0 Then
180         mResponseText = FastString '.ToString
190         mReadyState = 4
200       End If

210     Else
220       If Now > StartTimer Then
            ' abort
230         'mResponseText = GetTimeoutResponse()
240         'mReadyState = 4
250         'Exit Sub
260       End If

270       Socket.GetData DataBuffer, vbString  ' , optional maxlength (do not use maxlength!)
280       'FastString.Append DataBuffer
          FastString = FastString & DataBuffer
          'Debug.Print "Main Data " & DataBuffer

290       BytesNeeded = MaxL(0, BytesNeeded - bytesTotal)
300       If BytesNeeded <= 0 Then
310         mResponseText = FastString '.ToString
320         mReadyState = 4
330       End If
340     End If

Socket_DataArrival_Resume:
350      On Error GoTo 0
360      Exit Sub

Socket_DataArrival_Error:

370     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cHostConnection.Socket_DataArrival." & Erl
380     Resume Socket_DataArrival_Resume


End Sub

'Public Sub Socket_DataArrival(ByVal bytesTotal As Long)
'  Dim DataBuffer As String
'  Dim HeaderBytes As Long
'  'bytes in buffer = socket.BytesReceived
'
'  Debug.Print "cClient Socket_DataArrival " & vbCrLf & bytesTotal
'  Socket.PeekData DataBuffer, vbString
'  Debug.Print "cClient Socket_DataArrival.DataBuffer " & vbCrLf & DataBuffer
'
'  DataBuffer = ""
'
'  If mReadyState = 2 Then
'    mReadyState = 3
'  End If
'
'  If mReadyState = 1 Then
'    mReadyState = 2
'  End If
'  Debug.Print "Readystate "; mReadyState
'  Debug.Print "Waiting "; Waiting
'
'  If Waiting Then
'    Socket.PeekData DataBuffer, vbString  ' peekdata gets data but does not remove it
'    HeaderBytes = Checkdata(DataBuffer)
'    ' typically: "response 200 219"
'    If HeaderBytes > 15 Then
'      Socket.GetData DataBuffer, vbString, HeaderBytes  ' get just the header including \r\n
'      Debug.Print "HeaderBytes "; DataBuffer
'      bytesTotal = bytesTotal - HeaderBytes
'      Waiting = False
'      StartTimer = Win32.timeGetTime + 2000
'    End If
'    Socket.GetData DataBuffer, vbString  ' , optional maxlength (do not use maxlength!)
'    FastString.Append DataBuffer
'
'    BytesNeeded = MaxL(0, BytesNeeded - bytesTotal)
'    If BytesNeeded <= 0 Then
'      mResponseText = FastString.ToString
'      mReadyState = 4
'    End If
'
'  Else
'    If StartTimer < Win32.timeGetTime Then
'      ' abort
'      mResponseText = GetTimeoutResponse()
'      mReadyState = 4
'      Exit Sub
'    End If
'    Socket.GetData DataBuffer, vbString  ' , optional maxlength (do not use maxlength!)
'    FastString.Append DataBuffer
'    BytesNeeded = MaxL(0, BytesNeeded - bytesTotal)
'    If BytesNeeded <= 0 Then
'      mResponseText = FastString.ToString
'      mReadyState = 4
'    End If
'  End If
'
'End Sub
'



Public Function GetTimeoutResponse() As String
  Dim timoutstring As String
  timoutstring = "<?xml version=""1.0""?><HMC revision=""" & App.Revision & """><error>Timeout</error><errorcode>500</errorcode><module>HostConnection.GetTimeoutResponse</module></HMC>"
  dbg timoutstring & vbCrLf
  GetTimeoutResponse = timoutstring

  
End Function

Public Function AssignSocket(NewSocket As Winsock)
  If Not (Socket Is Nothing) Then
    Socket.Close
  End If
  Set Socket = NewSocket
End Function

Public Property Get getAllResponseHeaders() As String
  ' does not remove messages
End Property

Public Property Get ResponseReady() As Boolean
  ResponseReady = (READYSTATE = 4)
End Property
Public Property Get ResponseText() As String
  'Clears response text and faststring
  'resets readystate to 1
  'request is XML
  'FastString.Clear
  FastString = ""
  ResponseText = mResponseText
  mResponseText = ""
  Waiting = True
  mReadyState = 1
End Property

Public Function Send(ByVal Method As String, ByVal Request As String) As Long
  
        Dim packet As String
  
        Dim start As Date
        'dbg "cHostconnection Send " & Method
  
10       On Error GoTo Send_Error
        start = DateAdd("s", 2, Now)
  
30      packet = Join(Array(Method, CStr(Len(Request))), " ") & vbCrLf
40      packet = packet & Request
  
        'dbg "cHostconnection Socket.State " & Socket.State
        'If Socket.State = sckConnecting Then
        '  Socket.Close
        'End If
  
50      If Socket.State = sckClosed Then
60        Socket.Connect RemoteHost, RemotePort
70        Do While Socket.State = sckConnecting
            'dbg "cHostconnection connecting "
            Sleep 100
80          DoEvents
90          If Now > start Then  ' Win32.timeGetTime > t Then
100           Exit Do
110         End If
120       Loop
130     End If
140     If Socket.State = sckConnected Then
          'dbg "cHostconnection Socket.SendData " & left(Packet, 50)
          'dbg "cHostconnection Socket Connected"
150       Socket.SendData packet
160       mReadyState = 1
170     Else
          dbg "cHostconnection Socket NOT Connected"
180     End If

Send_Resume:
190      On Error GoTo 0
200      Exit Function

Send_Error:

210     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cHostConnection.Send." & Erl
220     Resume Send_Resume

End Function

Public Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  LastErrorCode = Number
  LastError = Description

End Sub


Public Sub Socket_SendComplete()
  If mReadyState < 2 Then
    mReadyState = 2
  End If
End Sub

Public Property Get READYSTATE() As Integer
  READYSTATE = mReadyState
End Property

Public Function Checkdata(ByVal Header As String) As Long

        Dim Ptr As Long
        Dim ParsedHeaders() As String

10      On Error GoTo Checkdata_Error

20      Ptr = InStr(Header, vbCrLf)
30      If Ptr > 0 And Ptr < 40 Then  ' max header is 19 chars
40        Header = left(Header, Ptr)
50        ParsedHeaders() = Split(Header, " ")
60        Select Case LCase(ParsedHeaders(0))
            Case "post"
70            BytesNeeded = Val(ParsedHeaders(1))
80          Case "get"
90            BytesNeeded = Val(ParsedHeaders(1))
100         Case "response"
110           BytesNeeded = Val(ParsedHeaders(2))
              'Debug.Print "Checkdata BytesNeeded " & BytesNeeded
120         Case Else
130       End Select


140       If BytesNeeded = 0 Then  ' added 10/27/2008 for connection issues
150       Else
160         Checkdata = Ptr + 1
170       End If


180     End If

Checkdata_Resume:
190     On Error GoTo 0
200     Exit Function

Checkdata_Error:

210     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cHostConnection.Checkdata." & Erl
220     Resume Checkdata_Resume


End Function

'winsock access wrappers

Public Sub Bind(ByVal Port As Long, ByVal IP As String)
  'Use this for multiple network cards
  Socket.Bind Port, IP
End Sub
Public Function ReConnect() As Boolean
  If Not (Socket Is Nothing) Then
    
    If Socket.State <> sckClosed Then
      Socket.Close
    End If
    Socket.Connect
  End If
End Function
Public Property Get IsConnected() As Boolean
  If (Socket Is Nothing) Then
    IsConnected = False
  Else
    IsConnected = (Socket.State = sckConnected)
  End If
  
End Property

Public Function Connect(ByVal Host As String, ByVal Port As String) As Boolean
  If Not (Socket Is Nothing) Then
    
    If Socket.State <> sckClosed Then
      Socket.Close
    End If
    
    Socket.RemoteHost = Host
    Socket.RemotePort = Port
    Socket.Connect
  End If
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

Public Sub Socket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  ' nada
End Sub
