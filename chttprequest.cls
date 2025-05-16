VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTTPRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Public Variables

Public XML               As String
Public Event Done()
Public Ready             As Boolean
Public FunctionMode      As Long
Public StatusCode        As Long
Public ZoneInfoList      As cZoneInfoList



Public Busy              As Boolean

Const BACKUP_6080 = 1
Const RESTORE_6080 = 2
Const GENERIC_ASYNC = 3

Private OperationType    As Long

Private SaveAsFileName   As String

' Private Variables

Private XMLHttpRequest   As MSXML2.XMLHTTP60

Private WithEvents XMLHandler As XMLHTTPRequestHandler
Attribute XMLHandler.VB_VarHelpID = -1



'  Enum READYSTATE
'    READYSTATE_UNINITIALIZED = 0    'Default initialization state
'    READYSTATE_LOADING = 1          'Object is currently loading its properties.
'    READYSTATE_LOADED = 2           'Object has been initialized
'    READYSTATE_INTERACTIVE = 3      'Object is interactive, but not all of its data is available
'    READYSTATE_COMPLETE = 4         'Object has received all of its data
'  End Enum


'12001          ERROR_INTERNET_OUT_OF_HANDLES
'               No more handles could be generated at this time.
'
'12002          ERROR_INTERNET_TIMEOUT
'               The request has timed out.
'
'12003          ERROR_INTERNET_EXTENDED_ERROR
'               An extended error was returned from the server. This is
'               typically a string or buffer containing a verbose error
'               message. Call InternetGetLastResponseInfo to retrieve the
'               error text.
'
'12004          ERROR_INTERNET_INTERNAL_ERROR
'               An internal error has occurred.
'
'12005          ERROR_INTERNET_INVALID_URL
'               The URL is invalid.
'
'12006          ERROR_INTERNET_UNRECOGNIZED_SCHEME
'               The URL scheme could not be recognized or is not supported.
'
'12007          ERROR_INTERNET_NAME_NOT_RESOLVED
'               The server name could not be resolved.
'
'12008          ERROR_INTERNET_PROTOCOL_NOT_FOUND
'               The requested protocol could not be located.
'
'12009          ERROR_INTERNET_INVALID_OPTION
'               A request to InternetQueryOption or InternetSetOption
'               specified an invalid option value.
'
'12010          ERROR_INTERNET_BAD_OPTION_LENGTH
'               The length of an option supplied to InternetQueryOption or
'               InternetSetOption is incorrect for the type of option
'               specified.
'
'12011          ERROR_INTERNET_OPTION_NOT_SETTABLE
'               The request option cannot be set, only queried.
'
'12012          ERROR_INTERNET_SHUTDOWN
'               The Win32 Internet function support is being shut down or
'               unloaded.
'
'12013          ERROR_INTERNET_INCORRECT_USER_NAME
'               The request to connect and log on to an FTP server could
'               not be completed because the supplied user name is
'               incorrect.
'
'12014          ERROR_INTERNET_INCORRECT_PASSWORD
'               The request to connect and log on to an FTP server could
'               not be completed because the supplied password is
'               incorrect.
'
'12015          ERROR_INTERNET_LOGIN_FAILURE
'               The request to connect to and log on to an FTP server
'               failed.
'
'12016          ERROR_INTERNET_INVALID_OPERATION
'               The requested operation is invalid.
'
'12017          ERROR_INTERNET_OPERATION_CANCELLED
'               The operation was canceled, usually because the handle on
'               which the request was operating was closed before the
'               operation completed.
'
'12018          ERROR_INTERNET_INCORRECT_HANDLE_TYPE
'               The type of handle supplied is incorrect for this
'               operation.
'
'12019          ERROR_INTERNET_INCORRECT_HANDLE_STATE
'               The requested operation cannot be carried out because the
'               handle supplied is not in the correct state.
'
'12020          ERROR_INTERNET_NOT_PROXY_REQUEST
'               The request cannot be made via a proxy.
'
'12021          ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND
'               A required registry value could not be located.
'
'12022          ERROR_INTERNET_BAD_REGISTRY_PARAMETER
'               A required registry value was located but is an incorrect
'               type or has an invalid value.
'
'12023          ERROR_INTERNET_NO_DIRECT_ACCESS
'               Direct network access cannot be made at this time.
'
'12024          ERROR_INTERNET_NO_CONTEXT
'               An asynchronous request could not be made because a zero
'               context value was supplied.
'
'12025          ERROR_INTERNET_NO_CALLBACK
'               An asynchronous request could not be made because a
'               callback function has not been set.
'
'12026          ERROR_INTERNET_REQUEST_PENDING
'               The required operation could not be completed because one
'               or more requests are pending.
'
'12027          ERROR_INTERNET_INCORRECT_FORMAT
'               The format of the request is invalid.
'
'12028          ERROR_INTERNET_ITEM_NOT_FOUND
'               The requested item could not be located.
'
'12029          ERROR_INTERNET_CANNOT_CONNECT
'               The attempt to connect to the server failed.
'
'12030          ERROR_INTERNET_CONNECTION_ABORTED
'               The connection with the server has been terminated.
'
'12031          ERROR_INTERNET_CONNECTION_RESET
'               The connection with the server has been reset.
'
'12032          ERROR_INTERNET_FORCE_RETRY
'               Calls for the Win32 Internet function to redo the request.
'
'12033          ERROR_INTERNET_INVALID_PROXY_REQUEST
'               The request to the proxy was invalid.
'
'12036          ERROR_INTERNET_HANDLE_EXISTS
'               The request failed because the handle already exists.
'
'12037          ERROR_INTERNET_SEC_CERT_DATE_INVALID
'               SSL certificate date that was received from the server is
'               bad. The certificate is expired.
'
'12038          ERROR_INTERNET_SEC_CERT_CN_INVALID
'               SSL certificate common name (host name field) is incorrect.
'               For example, if you entered www.server.com and the common
'               name on the certificate says www.different.com.
'
'12039          ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR
'               The application is moving from a non-SSL to an SSL
'               connection because of a redirect.
'
'12040          ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR
'               The application is moving from an SSL to an non-SSL
'               connection because of a redirect.
'
'12041          ERROR_INTERNET_MIXED_SECURITY
'               Indicates that the content is not entirely secure. Some of
'               the content being viewed may have come from unsecured
'               servers.
'
'12042          ERROR_INTERNET_CHG_POST_IS_NON_SECURE
'               The application is posting and attempting to change
'               multiple lines of text on a server that is not secure.
'
'12043          ERROR_INTERNET_POST_IS_NON_SECURE
'               The application is posting data to a server that is not
'               secure.
'
'12110          ERROR_FTP_TRANSFER_IN_PROGRESS
'               The requested operation cannot be made on the FTP session
'               handle because an operation is already in progress.
'
'12111          ERROR_FTP_DROPPED
'               The FTP operation was not completed because the session was
'               aborted.
'
'12130          ERROR_GOPHER_PROTOCOL_ERROR
'               An error was detected while parsing data returned from the
'               gopher server.
'
'12131          ERROR_GOPHER_NOT_FILE
'               The request must be made for a file locator.
'
'12132          ERROR_GOPHER_DATA_ERROR
'               An error was detected while receiving data from the gopher
'               server.
'
'12133          ERROR_GOPHER_END_OF_DATA
'               The end of the data has been reached.
'
'12134          ERROR_GOPHER_INVALID_LOCATOR
'               The supplied locator is not valid.
'
'12135          ERROR_GOPHER_INCORRECT_LOCATOR_TYPE
'               The type of the locator is not correct for this operation.
'
'12136          ERROR_GOPHER_NOT_GOPHER_PLUS
'               The requested operation can only be made against a Gopher+
'               server or with a locator that specifies a Gopher+
'               operation.
'
'12137          ERROR_GOPHER_ATTRIBUTE_NOT_FOUND
'               The requested attribute could not be located.
'
'12138          ERROR_GOPHER_UNKNOWN_LOCATOR
'               The locator type is unknown.
'
'12150          ERROR_HTTP_HEADER_NOT_FOUND
'               The requested header could not be located.
'
'12151          ERROR_HTTP_DOWNLEVEL_SERVER
'               The server did not return any headers.
'
'12152          ERROR_HTTP_INVALID_SERVER_RESPONSE
'               The server response could not be parsed.
'
'12153          ERROR_HTTP_INVALID_HEADER
'               The supplied header is invalid.
'
'12154          ERROR_HTTP_INVALID_QUERY_REQUEST
'               The request made to HttpQueryInfo is invalid.
'
'12155          ERROR_HTTP_HEADER_ALREADY_EXISTS
'               The header could not be added because it already exists.
'
'12156          ERROR_HTTP_REDIRECT_FAILED
'               The redirection failed because either the scheme changed
'               (for example, HTTP to FTP) or all attempts made to redirect
'               failed (default is five attempts).
'

Private mLastStatus      As Long



Public Function SendAsync(ByVal url As String, ByVal Querystring As String, Optional ByVal GenericAsync As Boolean = True)

  Dim Response      As String
  Dim auth          As String
  Dim Async         As Boolean
  Dim XML           As String
  Dim UserName      As String ' leave blank for now
  Dim Password      As String ' leave blank for now

  Ready = False
  
  Dim starttime As Date
  starttime = Now
  
  If GenericAsync Then
    OperationType = GENERIC_ASYNC
  End If
  
  Debug.Print "<<<<< SendAsync (Start)"
  
'   Timeout values are in milli-seconds
'    lResolve = 10 * 1000
'    lConnect = 10 * 1000
'    lSend = 10 * 1000
'    lReceive = 15 * 1000 'waiting time to receive data from server
'    XMLHTTP.setTimeouts lResolve, lConnect, lSend, lReceive
'
  
  Set XMLHandler = New XMLHTTPRequestHandler
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  
  XMLHttpRequest.OnReadyStateChange = XMLHandler
  url = url & Querystring
  
  Async = True
  XMLHttpRequest.Open "GET", url, Async, UserName, Password
  XMLHttpRequest.Send

End Function



Private Function SaveToFile() As Boolean
  
  ' called upon completion of ASYNC operation. example: 6080 configuration backup
  Dim hfile As Long
  Dim rc As Long
  
  On Error GoTo SaveToFile_Error

  rc = KillFile(SaveAsFileName)
  hfile = FreeFile
  Open SaveAsFileName For Binary Access Write As hfile
  Put #hfile, , XML
  Close hfile
  

SaveToFile_Resume:

  On Error GoTo 0
  Exit Function

SaveToFile_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cHTTPRequest.SaveToFile." & Erl
  Resume SaveToFile_Resume

End Function




Public Function Backup6080(ByVal exportfilename As String, ByVal Address As String, ByVal UserName As String, ByVal Password As String) As Byte()

  Dim url                As String
  Dim UpdateXML          As String
  Dim Response           As String
  Dim auth               As String
  Dim hfile              As Long
  Dim Async              As Boolean
  Dim BinaryData()       As Byte


  OperationType = BACKUP_6080

  On Error Resume Next

  If (Len(exportfilename)) > 0 Then

    KillFile exportfilename



    Async = False

    Set XMLHttpRequest = New MSXML2.XMLHTTP60
    'Set XMLHandler = New XMLHTTPRequestHandler
    ' Assign the wrapper class object to onreadystatechange.
    'Set XMLHandler.XMLHttpRequest = XMLHttpRequest
    'XMLHttpRequest.OnReadyStateChange = XMLHandler

    url = Address & "/PSIA/System/configurationData"

    XMLHttpRequest.Open "GET", url, Async, UserName, Password
    Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
    auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
    Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
    Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
    XMLHttpRequest.Send


    Select Case XMLHttpRequest.Status
      Case 200, 201
        'ReDim BinaryData(200000)
      
        BinaryData = XMLHttpRequest.responseBody
        hfile = FreeFile
        Open exportfilename For Binary Access Write As hfile
        Put #hfile, , BinaryData
        Close #hfile
      


      'Backup6080 = XMLHttpRequest.responseBody
    Case Else
      ReDim BinaryData(0)
  End Select

  Set XMLHttpRequest = Nothing
End If

End Function
Public Function Restore6080(ByVal filename As String, ByVal Address As String, ByVal UserName As String, ByVal Password As String) As Boolean

  Dim Response           As String
  Dim auth               As String
  Dim url                As String
  Dim Async              As Boolean
  Dim XML                As String
  Dim Blocks             As Long
  Dim Block              As Long
  Dim BlockData          As String
  Dim Buffer             As String
  Dim OK                 As Boolean
  Const BLOCKSIZE        As Long = 256

  Dim FileLen            As Long

  Dim hfile              As Long

  hfile = FreeFile

  Open filename For Binary Access Read As #hfile
  FileLen = LOF(hfile)
  If FileLen < 1000 Then
    ' don't do it!
  Else

    Buffer = Space$(FileLen)
    Get #hfile, , Buffer
    Blocks = FileLen / BLOCKSIZE
    If Blocks * BLOCKSIZE < FileLen Then
      Blocks = Blocks + 1
    End If

    Async = False  ' !!!!!!!!!!!!
    Set XMLHttpRequest = New MSXML2.XMLHTTP60

    For Block = 0 To Blocks - 1
      BlockData = MID$(Buffer, (Block * BLOCKSIZE) + 1, BLOCKSIZE)
      If Len(BlockData) Then
        url = Address & "/PSIA/System/configurationData/" & Block
        XMLHttpRequest.Open "PUT", url, Async, UserName, Password
        Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
        auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
        Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
        Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
        XMLHttpRequest.Send BlockData

        Select Case XMLHttpRequest.Status
          Case 200, 201
            OK = True
          Case Else
            OK = False
        End Select

        If Not (OK) Then
          Exit For
        End If
      End If
    Next
    Set XMLHttpRequest = Nothing


  End If

End Function



Public Property Get READYSTATE() As Long
  If Not XMLHttpRequest Is Nothing Then
    READYSTATE = XMLHttpRequest.READYSTATE
  Else
    READYSTATE = 0
  End If

End Property

Public Function Reboot6080(ByVal Address As String, ByVal UserName As String, ByVal Password As String) As Boolean
  ''http://192.168.1.122/PSIA/System/reboot
  
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  
  
  Async = False
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  url = Address & "/PSIA/System/reboot"
  
  XMLHttpRequest.Open "PUT", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send
  
  
  Select Case XMLHttpRequest.Status
    Case 200, 201
      Reboot6080 = True
    Case Else
      Reboot6080 = False
  End Select
  
  Set XMLHttpRequest = Nothing

End Function



Public Function setNewIDL(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneID As Long, ByVal IDL As Long) As Boolean
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  ' set each to zero then
  ' set the one we want
  '  "<ZoneInfo><DeviceID>" & Device.DecimalSerial & "</DeviceID><PTI>" & Device.PTI & "</PTI>" & _
     '        "<MID>" & Device.MID & "</MID><Locatable>" & IIf(Device.IDL = 1, "true", "false") & "</Locatable>" & _
     '        "<Description>" & Device.Serial & "</Description><IsRef>" & IIf(Device.IDL = 2, "true", "false") & "</IsRef>" & _
     '        "<SyncWindow>0</SyncWindow><SyncTimeout>0</SyncTimeout><MessageExpirationTime>0</MessageExpirationTime><CheckInTime>0</CheckInTime>" & _
     '        "<SupervisionWindow>" & Device.Checkin6080 & "</SupervisionWindow><IsSPDevice>" & IIf(Device.IDL = 3, "true", "false") & "</IsSPDevice></ZoneInfo>"
  '
  UpdateZone Address, UserName, Password, ZoneID, "IsSPDevice", "false"
  UpdateZone Address, UserName, Password, ZoneID, "Locatable", "false"
  UpdateZone Address, UserName, Password, ZoneID, "IsRef", "false"



  Select Case IDL
    Case 3
      UpdateZone Address, UserName, Password, ZoneID, "IsSPDevice", "true"
    Case 2
      UpdateZone Address, UserName, Password, ZoneID, "IsRef", "true"
    Case 1
      UpdateZone Address, UserName, Password, ZoneID, "Locatable", "true"
    Case Else
  End Select



End Function


Public Function UpdateZoneCheckinTime(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneID As Long, ByVal CheckInTime As Long) As Boolean
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String


'<ZoneInfo>
'<Description>Joe’s New Pendant</Description>
'</ZoneInfo>

  XML = "<ZoneInfo><SupervisionWindow>" & CheckInTime & "</SupervisionWindow></ZoneInfo>"

  Async = False
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneID
  
  XMLHttpRequest.Open "PUT", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML
  
  Select Case XMLHttpRequest.Status
    Case 200, 201
      UpdateZoneCheckinTime = True
    Case Else
      UpdateZoneCheckinTime = False
  End Select
  
  Set XMLHttpRequest = Nothing

End Function


Public Function UnRegisterSoftPoint(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ID As Long) As Boolean
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String

  Async = False
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  url = Address & "/PSIA/AreaControl/SoftPointList/" & ID
  
  XMLHttpRequest.Open "DELETE", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send
  
  Select Case XMLHttpRequest.Status
    Case 200, 201
      UnRegisterSoftPoint = True
    Case Else
      UnRegisterSoftPoint = False
  End Select

  Set XMLHttpRequest = Nothing


End Function

Public Function RegisterSoftPoint(ByVal Address As String, ByVal UserName As String, ByVal Password As String, SoftPoint As cSoftPoint) As String ' XML
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String

  XML = SoftPoint.GetRegistrationXML

  Async = False
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  url = Address & "/PSIA/AreaControl/SoftPointList"
  
  XMLHttpRequest.Open "POST", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML
  
  Select Case XMLHttpRequest.Status
    Case 200, 201
      RegisterSoftPoint = XMLHttpRequest.ResponseText
    Case Else
  End Select

  Set XMLHttpRequest = Nothing


End Function
Public Function GetSoftPointList(ByVal Address As String, ByVal UserName As String, ByVal Password As String) As String
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String

  Async = True



  Set XMLHttpRequest = New MSXML2.XMLHTTP60
 ' Set XMLHandler = New XMLHTTPRequestHandler
 ' Set XMLHandler.XMLHttpRequest = XMLHttpRequest
 ' XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/SoftPointList"

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler ' << doesn't really work
  XMLHttpRequest.Open "GET", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  Do Until XMLHttpRequest.READYSTATE = 0 Or XMLHttpRequest.READYSTATE = 4
    DoEvents
  Loop
  

  Select Case XMLHttpRequest.Status
    Case 200, 201
      GetSoftPointList = XMLHttpRequest.ResponseText
      StatusCode = XMLHttpRequest.Status
    Case Else
      StatusCode = XMLHttpRequest.Status
  End Select

  Me.XML = XMLHttpRequest.ResponseText
  Ready = True

  On Error Resume Next
  Set XMLHandler.XMLHttpRequest = Nothing
  Set XMLHandler = Nothing
  Set XMLHttpRequest = Nothing
  

End Function


Public Function SetNID(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal NewNID As Long) As Long
  SetNID = -1
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String

  Dim Node          As IXMLDOMNode

  Dim OldNID        As Long

  OldNID = GetNID(Address, UserName, Password)

  If OldNID = NewNID Then
    Exit Function
  End If

  Async = False

  XML = "<RFNetwork><NID>" & NewNID & "</NID></RFNetwork>"

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/System/RFNetwork"

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "PUT", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML

  Select Case XMLHttpRequest.Status
    Case 200, 201

      Dim doc       As DOMDocument60
      Set doc = New DOMDocument60

      XML = XMLHttpRequest.ResponseText

      If doc.LoadXML(XML) Then
        '<RFNetwork><NID>10</NID></RFNetwork>
        Set Node = doc.selectSingleNode("RFNetwork/NID")
        If Not Node Is Nothing Then
          SetNID = Val(Node.text)
        End If
      End If
    Case Else

      XML = XMLHttpRequest.ResponseText
      If InStr(1, XML, "Duplicate NID", vbTextCompare) Then
        SetNID = -6
      End If
        '<?xml version="1.0" encoding="utf-8" ?>
        '<ResponseStatus version="1.0" xmlns:urn="psialliance-org">
        '  <requestURL>/PSIA/System/RFNetwork</requestURL>
        '  <statusCode>6</statusCode>
        '  <statusString>Duplicate NID, Nothing to do</statusString>
        '</ResponseStatus>

    End Select



    Set XMLHandler.XMLHttpRequest = Nothing

    Set XMLHandler = Nothing  ' may need to switch these two
    Set XMLHttpRequest = Nothing





  End Function
Public Function GetNID(ByVal Address As String, ByVal UserName As String, ByVal Password As String) As Long
  ' return -1 on error
  GetNID = -1
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String

  Dim Node          As IXMLDOMNode
  Async = False


  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/System/RFNetwork"

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "GET", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  Select Case XMLHttpRequest.Status
    Case 200, 201

      Dim doc       As DOMDocument60
      Set doc = New DOMDocument60

      XML = XMLHttpRequest.ResponseText

      If doc.LoadXML(XML) Then
        '<RFNetwork><NID>10</NID></RFNetwork>
        Set Node = doc.selectSingleNode("RFNetwork/NID")
        If Not Node Is Nothing Then
          GetNID = Val(Node.text)
        End If
      End If
    Case Else



  End Select



  Set XMLHandler.XMLHttpRequest = Nothing

  Set XMLHandler = Nothing    ' may need to switch these two
  Set XMLHttpRequest = Nothing




  'http://192.168.1.122/PSIA/System/RFNetwork
End Function



Public Function ChangeZoneParameter(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneID As Long, ByVal ParamName As String, ByVal Value As String) As Long

'PUT /PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/n
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  Dim Success       As Boolean
  Async = False
  
  XML = "<ZoneInfo><" & ParamName & ">" & Value & "</" & ParamName & "></ZoneInfo>"
  
  
  
  
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneID

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "PUT", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
    Case Else
      Success = 0
  End Select
  Set XMLHandler.XMLHttpRequest = Nothing

  Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing

  ChangeZoneParameter = Success

End Function

Public Function DeletePartition(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal PartitionID As Long) As Boolean
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  Dim Success       As Boolean
  Async = False
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/Partitions/PartitionInfoList/" & PartitionID
  ' page 31 guide 1.5 has URL Wrong as follows
  'DELETE /PSIA/AreaControl/Partitions/PartitionList/18
  'should be
  'http://" & ip1 & "/PSIA/AreaControl/Partitions/PartitionInfoList/9

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "DELETE", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
    Case Else
      Success = 0
  End Select
  Set XMLHandler.XMLHttpRequest = Nothing

  Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing

  DeletePartition = Success
End Function

Public Function GetZonesForPartition(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal PartitionID As Long) As Collection
  Dim rc            As Long
  Dim ZoneInfo      As cZoneInfo
  Dim ZoneID        As Long
  Dim part          As cPartition


  rc = GetZoneList(Address, UserName, Password)

  Do Until Ready
    DoEvents
  Loop

  ProcessZonelist
  For Each ZoneInfo In ZoneInfoList.ZoneList
    DoEvents
    For Each part In ZoneInfo.Partitionlist
      If part.PartitionID = PartitionID Then
        ZoneID = ZoneInfo.ID
        rc = UnassignPartition(Address, UserName, Password, PartitionID, ZoneID)
        Exit For
      End If
    Next
  Next


End Function


Public Function AddPartitionList(ByVal Address As String, ByVal UserName As String, ByVal Password As String, Parts As Collection, ByVal ZoneID As Long)
  Dim part As cPartition
  For Each part In Parts
    AssignPartition Address, UserName, Password, part.PartitionID, ZoneID
  Next


End Function
Public Function AssignPartition(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal PartitionID As Long, ByVal ZoneID As Long)

  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  Dim Success       As Boolean

 
'<Partition>
'<PartitionID>1</PartitionID>
'</Partition>
  Async = False


'http://192.168.1.122/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/16/PartitionList?%3CPartition%3E%3CPartitionID%3E14%3C%2FPartitionID%3E%3C%2FPartition%3E=

  XML = "<Partition><PartitionID>" & PartitionID & "</PartitionID></Partition>"
  '<Partition><PartitionID>14</PartitionID></Partition>
  '<Partition><PartitionID>14</PartitionID></Partition>
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneID & "/PartitionList"
  'http://192.168.1.122/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/32/PartitionList
  'http://192.168.1.122/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/16/PartitionList
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "POST", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
    Case Else
      Success = 0
  End Select




End Function


Public Function RemovePartitionList(ByVal Address As String, ByVal UserName As String, ByVal Password As String, Parts As Collection, ByVal ZoneID As Long)
  Dim part As cPartition
  For Each part In Parts
    UnassignPartition Address, UserName, Password, part.PartitionID, ZoneID
  Next


End Function

Public Function UnassignPartition(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal PartitionID As Long, ByVal ZoneID As Long)

  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  Dim Success       As Boolean

  ' /PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & zoneid & "/PartitionList/" & partitionid

  Async = False

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneID & "/PartitionList/" & PartitionID

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "DELETE", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
    Case Else
      Success = 0
  End Select


  UnassignPartition = Success

End Function




Public Function CreatePartition(ByVal Address As String, ByVal UserName As String, ByVal Password As String, partition As cPartition) As String

  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  Dim Success       As Boolean

  Dim ms As Long
  ms = Win32.timeGetTime
  '<PartitionInfo>
  '  <Description>Front Door</Description>
  '  <IsMobile>false</IsMobile>
  '</PartitionInfo>

  XML = "<PartitionInfo><Description>" & XMLEncode(partition.Description) & "</Description><IsMobile>" & IIf(partition.IsLocation, "true", "false") & "</IsMobile></PartitionInfo>"

  Async = False

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/Partitions/PartitionInfoList"

  '/PSIA/AreaControl/Partitions/PartitionInfoList/

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "POST", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
    Case Else
      Success = 0
  End Select

  'response = XMLHttpRequest.responseXML.XML
  'If response = "" Or IsGarbage(response) Then
  '  response = XMLHttpRequest.responseBody
  'End If
  'If response = "" Or IsGarbage(response) Then
    Response = XMLHttpRequest.ResponseText
  'End If
  'If response = "" Or IsGarbage(response) Then
  '  response = ""
  'End If

  Set XMLHandler.XMLHttpRequest = Nothing

  Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing

  CreatePartition = Response
  ms = Win32.timeGetTime - ms
  
  'MsgBox "Time in ms " & ms

End Function

Public Function UpdatePartition(ByVal Address As String, ByVal UserName As String, ByVal Password As String, partition As cPartition) As Long
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim oldxml        As String
  Dim Success       As Boolean

  Async = False

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  'Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  'Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  'XMLHttpRequest.OnReadyStateChange = XMLHandler

  'XML = "<PartitionInfo><Description>" & XMLEncode(partition.Description) & "</Description></PartitionInfo>"

  XML = "<PartitionInfo><Description>" & XMLEncode(partition.Description) & "</Description><IsMobile>" & IIf(partition.IsLocation, "true", "false") & "</IsMobile></PartitionInfo>"

  url = Address & "/PSIA/AreaControl/Partitions/PartitionInfoList/" & partition.PartitionID
  '/PSIA/AreaControl/Partitions/PartitionInfoList/
  
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  'Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "PUT", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
    Case Else
      Success = 0
  End Select


  Response = XMLHttpRequest.responseXML.XML
  If Response = "" Or IsGarbage(Response) Then
    Response = XMLHttpRequest.responseBody
  End If
  If Response = "" Or IsGarbage(Response) Then
    Response = XMLHttpRequest.ResponseText
  End If
  If Response = "" Or IsGarbage(Response) Then
    Response = ""
  End If

  'Set XMLHandler.XMLHttpRequest = Nothing
  'Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing


  UpdatePartition = IIf(Success, 1, 0)


End Function


Public Function GetPartitionList(ByVal Address As String, ByVal UserName As String, ByVal Password As String) As String
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim oldxml        As String
  Dim Success       As Boolean

  Async = True

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler



  url = Address & "/PSIA/AreaControl/Partitions/PartitionInfoList"  ' Get Partition List Good

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "GET", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  Do Until XMLHttpRequest.READYSTATE = 0 Or XMLHttpRequest.READYSTATE = 4
    DoEvents
  Loop

  Select Case XMLHttpRequest.Status
    Case 200, 201
      Success = 1
      StatusCode = XMLHttpRequest.Status
    Case Else
      StatusCode = XMLHttpRequest.Status
      Success = 0
  End Select


  Debug.Print "Done"


  Me.Ready = True

  'response = XMLHttpRequest.responseXML.XML

  Response = XMLHttpRequest.ResponseText
  XML = Response

  On Error Resume Next
  

  Set XMLHandler.XMLHttpRequest = Nothing

  Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing


  GetPartitionList = Response


End Function


Public Function UnRegisterDevice(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneID As Long) As Long

  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  
  Async = False

  Set XMLHttpRequest = New MSXML2.XMLHTTP60

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneID
                  '/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/2
  Debug.Print XML

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  'Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "DELETE", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  StatusCode = XMLHttpRequest.Status
  
  Select Case XMLHttpRequest.Status
    Case 200, 201
      UnRegisterDevice = 1
    Case Else
      UnRegisterDevice = 0
  End Select
  
  Set XMLHttpRequest = Nothing


End Function
Public Function Get6080Data(ByVal Address As String, ByVal UserName As String, ByVal Password As String) As String
  Dim XMLHttpRequest As MSXML2.XMLHTTP60
  Dim Async         As Boolean
  Dim url           As String
  Dim auth          As String

  Async = False

  Set XMLHttpRequest = New MSXML2.XMLHTTP60

  url = Address & "/PSIA/System/deviceInfo"  '

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  
  XMLHttpRequest.Open "GET", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  Select Case XMLHttpRequest.Status
    Case 200, 201
      'Status = 1
      Get6080Data = XMLHttpRequest.ResponseText
    Case Else
      Get6080Data = ""
     ' Status = 0
  End Select


  Set XMLHttpRequest = Nothing
End Function


Public Function RegisterDevice(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal XML As String) As String

  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  
  Async = False

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
'  Set XMLHandler = New XMLHTTPRequestHandler
'  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
'  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList"

  ' UpdateXML = "<ZoneInfo><DeviceID>6853285</DeviceID><PTI>0</PTI><MID>178</MID><Description>New Device</Description><IsRef>false</IsRef><Locatable>false</Locatable><IsSPDevice>false</IsSPDevice><SupervisionWindow>3600</SupervisionWindow></ZoneInfo>"
 
  Debug.Print XML

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  'Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "POST", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send XML

  Select Case XMLHttpRequest.Status
    Case 200, 201
      StatusCode = XMLHttpRequest.Status
    Case Else
      StatusCode = XMLHttpRequest.Status
  End Select
  
  RegisterDevice = XMLHttpRequest.ResponseText

  Set XMLHttpRequest = Nothing

End Function

Public Function GetRegisterResponse() As Long

  Dim doc           As DOMDocument60
  Set doc = New DOMDocument60
  If doc.LoadXML(Me.XML) Then
    Debug.Print XML
  End If


End Function


Function DeleteDevice(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneNumber As Long) As Long
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Async = False

  On Error GoTo DeleteDevice_Error


  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler


  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneNumber

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XMLHttpRequest.Open "DELETE", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send

  If XMLHttpRequest.Status = 200 Then
    DeleteDevice = ZoneNumber
  Else
    DeleteDevice = 0
  End If



DeleteDevice_Resume:

  On Error GoTo 0
  Exit Function

DeleteDevice_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cHTTPRequest.DeleteDevice." & Erl
  Resume DeleteDevice_Resume

End Function

Public Function GetDeleteResponse() As Long
  ' returns 200 if OK
  Dim doc           As DOMDocument60
  Set doc = New DOMDocument60
  If doc.LoadXML(Me.XML) Then
    Debug.Print XML

  End If


End Function

Function UpdateZone(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneID As Long, ByVal Fieldname As String, ByVal Value As String)
  Dim url           As String
  Dim UpdateXML     As String
  Dim Response      As String
  Dim auth          As String
  Dim IP            As String


  Dim Async As Boolean
  
  Async = True

  UpdateXML = "<ZoneInfo><" & Fieldname & ">" & XMLEncode(Value) & "</" & Fieldname & "></ZoneInfo>"



  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  ' Assign the wrapper class object to onreadystatechange.
  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  IP = Address

  url = IP & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList/" & ZoneID
  ' http://msdn.microsoft.com/en-us/library/windows/desktop/ms757030(v=vs.85).aspx
  'XMLHttpRequest.Open "GET", URL, True, user1 , pw1

  XMLHttpRequest.Open "PUT", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send UpdateXML



End Function

Function ProcessZonelist() As Long  ' return # of zones
  Set ZoneInfoList = New cZoneInfoList
  If Len(XML) Then
    ZoneInfoList.LoadXML XML
  End If
  ProcessZonelist = ZoneInfoList.ZoneList.Count

End Function
Function GetSingleZoneInfo(ByVal Address As String, ByVal UserName As String, ByVal Password As String, ByVal ZoneID As Long) As cZoneInfo
'/PSIA/AreaControl/PartitionMembers/Zones/info/n
  Dim Response      As String
  Dim auth          As String
  Dim url           As String
  Dim Async         As Boolean
  Dim XML           As String
  
  Async = False

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/info/" & ZoneID

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  XMLHttpRequest.Open "GET", url, Async, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  XMLHttpRequest.Send
  
  StatusCode = XMLHttpRequest.Status
  Dim ZoneInfo As cZoneInfo
  Set ZoneInfo = New cZoneInfo
  XML = XMLHttpRequest.ResponseText
  Select Case StatusCode
    Case 200, 201
      ZoneInfo.ParseXML XML
      
      'RegisterDevice = 1
    Case Else
      'RegisterDevice = 0
  End Select
  
  Set GetSingleZoneInfo = ZoneInfo

End Function


Function GetZoneList(ByVal Address As String, ByVal UserName As String, ByVal Password As String) As Long
  ' http://msdn.microsoft.com/en-us/library/windows/desktop/ms757030(v=vs.85).aspx
  ' XMLHttpRequest.Open "GET", URL, True, user1 , pw1
  ' Assign the wrapper class object to onreadystatechange.
  
  'http://192.168.60.80/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList
  Dim url           As String
  Dim Response      As String
  Dim auth          As String
'  On Error Resume Next
  ' This is an ASYNC CALL!

  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set XMLHandler = New XMLHTTPRequestHandler
  XML = ""

  XMLHandler.RequestType = "ZoneInfoList"

  Set XMLHandler.XMLHttpRequest = XMLHttpRequest
  XMLHttpRequest.OnReadyStateChange = XMLHandler

  url = Address & "/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList"

  XMLHttpRequest.Open "GET", url, True, UserName, Password
  Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
  auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
  Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")

  XMLHttpRequest.Send  ' nothing for GETs

End Function

Function PingWatchdogHost(ByVal Address As String, ByVal Request As String, ByVal Port As Long, ByVal message As String, ByVal UserName As String, ByVal Password As String) As Long

        Dim url           As String
        Dim Response      As String
        Dim auth          As String
        Dim Async         As Boolean

10      Async = True
        
        '*************** This is an ASYNC CALL! ****************

20      Set XMLHttpRequest = New MSXML2.XMLHTTP60
30      Set XMLHandler = New XMLHTTPRequestHandler
40      XML = ""

50      XMLHandler.RequestType = "PingWatchdogHost"

60      Set XMLHandler.XMLHttpRequest = XMLHttpRequest
70      XMLHttpRequest.OnReadyStateChange = XMLHandler

        'Address should have http://ip [or domain]
        
80      If Right$(Address, 1) = "/" Then
90        Address = left$(Address, Len(Address) - 1)
100     End If
        
110     If Not (Port = 80 Or Port = 0) Then
120       Address = Address & ":" & Port
130     End If

140     If left$(Request, 1) <> "/" Then
150       Request = "/" & Request
160     End If


170     url = Address & Request & IIf(Len(message), "?" & message, "")

180     XMLHttpRequest.Open "GET", url, Async, UserName, Password
190     Call XMLHttpRequest.setRequestHeader("Content-Type", "text/xml")  ' not needed but safe
200     auth = EncodeBase64(UserName & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
210     Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
220     Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")

230     XMLHttpRequest.Send ' nothing for GETs


End Function



Private Sub Class_Terminate()
  
  
  On Error Resume Next
  
  If Not XMLHttpRequest Is Nothing Then
  
  If (XMLHttpRequest.READYSTATE = 0) Or (XMLHttpRequest.READYSTATE = 4) Then
    ' OK to shut down
  Else
    XMLHttpRequest.abort

  End If
  RaiseEvent Done
  If Not XMLHandler Is Nothing Then
     Set XMLHandler.XMLHttpRequest = Nothing
  End If
  Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing
  End If
  
End Sub

Private Sub XMLHandler_Done()

  Dim Response           As String

  StatusCode = XMLHttpRequest.Status


  If StatusCode > INTERNET_ERROR_BASE Then
        'Debug.Print InetErrorToString(StatusCode) & " Status: " & StatusCode & " " & Now
        LastStatus = XMLHttpRequest.Status
        Busy = False
        
  Else

    Select Case OperationType
        
      Case GENERIC_ASYNC
         LastStatus = 0
         XML = XMLHttpRequest.ResponseText
         
      Case BACKUP_6080

        XML = XMLHttpRequest.ResponseText
        SaveToFile
        Busy = False

      Case RESTORE_6080

      Case Else
        Response = XMLHttpRequest.responseXML.XML
        If Response = "" Then
          If Not (IsGarbage(Response)) Then
            Response = XMLHttpRequest.responseBody & ""
          End If
        End If
        If Response = "" Then
          If Not (IsGarbage(Response)) Then
            Response = XMLHttpRequest.ResponseText & ""
          End If
        End If
        If Response = "" Then
          If Not (IsGarbage(Response)) Then
            Response = "Nothing" & vbCrLf
          End If
        End If


    End Select
  End If

  Set XMLHandler.XMLHttpRequest = Nothing

  Set XMLHandler = Nothing  ' may need to switch these two
  Set XMLHttpRequest = Nothing

  XML = Response

  RaiseEvent Done
  Ready = True
  Busy = False

End Sub

Public Property Get LastStatus() As Long

  LastStatus = mLastStatus

End Property

Public Property Let LastStatus(ByVal LastStatus As Long)

  mLastStatus = LastStatus

End Property



