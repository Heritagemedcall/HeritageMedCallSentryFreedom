Attribute VB_Name = "modNet"
Option Explicit

Private mLastFTPError As String

'The settings for the Winsock State property are:
'
'Constant Value Description
'sckClosed 0 Default. Closed
'sckOpen 1 Open
'sckListening 2 Listening
'sckConnectionPending  3 Connection pending
'sckResolvingHost  4 Resolving host
'sckHostResolved  5 Host resolved
'sckConnecting  6 Connecting
'sckConnected  7 Connected
'sckClosing  8 Peer is closing the connection
'sckError  9 Error


' Winsock Error codes
'sckOutOfMemory 7 Out of memory
'sckInvalidPropertyValue 380 The property value is invalid.
'sckGetNotSupported 394 The property can't be read.
'sckSetNotSupported 383 The property is read-only.
'sckBadState 40006 Wrong protocol or connection state for the requested transaction or request.
'sckInvalidArg 40014 The argument passed to a function was not in the correct format or in the specified range.
'sckSuccess 40017 Successful.
'sckUnsupported 40018 Unsupported variant type.
'sckInvalidOp 40020 Invalid operation at current state
'sckOutOfRange 40021 Argument is out of range.
'sckWrongProtocol 40026 Wrong protocol for the requested transaction or request
'sckOpCanceled 1004 The operation was canceled.
'sckInvalidArgument 10014 The requested address is a broadcast address, but flag is not set.
'sckWouldBlock 10035 Socket is non-blocking and the specified operation will block.
'sckInProgress 10036 A blocking Winsock operation in progress.
'sckAlreadyComplete 10037 The operation is completed. No blocking operation in progress
'sckNotSocket 10038 The descriptor is not a socket.
'sckMsgTooBig 10040 The datagram is too large to fit into the buffer and is truncated.
'sckPortNotSupported 10043 The specified port is not supported.
'sckAddressInUse 10048 Address in use.
'sckAddressNotAvailable 10049 Address not available from the local machine.
'sckNetworkSubsystemFailed 10050 Network subsystem failed.
'sckNetworkUnreachable 10051 The network cannot be reached from this host at this time.
'sckNetReset 10052 Connection has timed out when SO_KEEPALIVE is set.
'sckConnectAborted 11053 Connection is aborted due to timeout or other failure.
'sckConnectionReset 10054 The connection is reset by remote side.
'sckNoBufferSpace 10055 No buffer space is available.
'sckAlreadyConnected 10056 Socket is already connected.
'sckNotConnected 10057 Socket is not connected.
'sckSocketShutdown 10058 Socket has been shut down.
'sckTimedout 10060 Socket has been shut down.
'sckConnectionRefused 10061 Connection is forcefully rejected.
'sckNotInitialized 10093 WinsockInit should be called first.
'sckHostNotFound 11001 Authoritative answer: Host not found.
'sckHostNotFoundTryAgain 11002 Non-Authoritative answer: Host not found.
'sckNonRecoverableError 11003 Non-recoverable errors.
'sckNoData 11004 Valid name, no data record of requested type.




Public Function GetLastFTPError() As String

  GetLastFTPError = mLastFTPError
End Function

Public Sub SetPushMobileEntries()
  Dim url As String
  Dim Enabled As Long
  Dim Retries As Long
  
  url = ReadSetting("Push", "URL", "")
  Enabled = Val(ReadSetting("Push", "Enabled", "0"))
  Retries = Val(ReadSetting("Push", "Retries", "0"))
  
  WriteSetting "Push", "URL", url
  WriteSetting "Push", "Enabled", Enabled And 1
  WriteSetting "Push", "Retries", Fix(Retries)
  

  url = ReadSetting("Mobile", "Root", "")
  Enabled = Val(ReadSetting("Mobile", "Enabled", "0"))
  
  
  WriteSetting "Mobile", "Root", url
  WriteSetting "Mobile", "Enabled", Enabled And 1


End Sub



Public Function URLEncode(ByVal s As String) As String

  Dim StringLen          As Long
  Dim i                  As Long
  Dim CharCode           As Integer
  Dim Char               As String

  StringLen = Len(s)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    For i = 1 To StringLen
      Char = MID$(s, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = "+"
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If



  ' RESERVED
  'ampersand ("&")
  'dollar ("$")
  'plus sign("+")
  'comma (",")
  'Forward slash("/")
  'colon (":")
  'semi -colon(";")
  'equals ("=")
  'question mark("?")
  ''At' symbol ("@")
  'pound ("#").
  
  'UNSAFE
  'Space (" ")
  'less than and greater than ("<>")
  'open and close brackets ("[]")
  'open and close braces ("{}")
  'pipe ("|")
  'backslash ("\")
  'caret ("^")
  'Percent ("%")



End Function


Public Function XMLEncode(ByVal s As String) As String
  s = Replace(s, "&", "&amp;")
  s = Replace(s, "'", "&apos;")
  s = Replace(s, """", "&quot;")
  s = Replace(s, ">", "&gt;")
  s = Replace(s, "<", "&lt;")
  XMLEncode = s

End Function
Public Function XMLDecode(ByVal s As String) As String
If InStr(1, s, "&amp;", vbTextCompare) > 0 Then
  'Stop
End If
  s = Replace(s, "&amp;", "&")
  s = Replace(s, "&apos;", "'")
  s = Replace(s, "&quot;", """")
  s = Replace(s, "&gt;", ">")
  s = Replace(s, "&lt;", "<")
  XMLDecode = s
 

End Function

Public Function HTMLEncode(ByVal s As String) As String

  Dim Char          As String
  Dim result        As String
  Dim code          As Long
  Dim j As Long
  
  For j = 1 To Len(s)
    Char = MID$(s, j, 1)
    code = Asc(Char)

    Select Case code

      Case Is < 32
        result = result & "&#" & CStr(code) & ";"
      Case Is > 127
        result = result & "&#" & CStr(code) & ";"
      Case 34
        result = result & "&quot;"
      Case 38
        result = result & "&amp;"
      Case 39
        result = result & "&apos;"
      Case 60
        result = result & "&lt;"
      Case 62
        result = result & "&gt;"
      Case Else
        result = result & Char
    End Select
  Next

  HTMLEncode = result

  's = Replace(s, "&", "&amp;")
  's = Replace(s, "'", "&apos;")
  's = Replace(s, """", "&quot;")
  's = Replace(s, ">", "&gt;")
  's = Replace(s, "<", "&lt;")



End Function


Public Function taggit(ByVal tag As String, ByVal text As String) As String
  taggit = "<" & tag & ">" & text & "</" & tag & ">"

End Function
Public Function Indent(ByVal Num As Integer) As String
  Indent = String(Num * 2, " ")
End Function

Public Function GetValidIP(ByVal IP As String, ByVal default As String) As String
  Dim Octets()      As String
  Dim Octet         As Long
  Dim j             As Integer


  Octets = Split(IP, ".", 4)
  If UBound(Octets) < 3 Then
    GetValidIP = default
  Else
    For j = 0 To 3
      Octet = Val(Octets(j))
      If Octet < 0 Or Octet > 255 Then
        Octets(j) = "0"
      Else
        Octets(j) = CStr(CInt(Octet))
      End If
    Next
    GetValidIP = Join(Octets, ".")
  End If
End Function


Function GetValidRemotePort(ByVal RemotePort As Long, ByVal default As Long) As Long
  If RemotePort < 0 Or RemotePort > 65535 Then
    GetValidRemotePort = default
  Else
    GetValidRemotePort = RemotePort
  End If
End Function

Function TestFTPConnection(ByVal hostname As String, ByVal Username As String, ByVal Password As String) As Boolean
  Dim ftp As cFTP
  Dim rc  As Boolean
  mLastFTPError = ""
  Set ftp = New cFTP
  rc = ftp.OpenConnection(hostname, Username, Password)
  If False = rc Then
    mLastFTPError = ftp.GetLastErrorMessage()
  End If
  ftp.CloseConnection
  Set ftp = Nothing
  TestFTPConnection = rc
  

End Function
