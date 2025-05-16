VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIVALine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''' CLASS NOT USED '''''''''''''''''

'Events
Public Event MakeCallResult(ByVal errorCode As Long)
Public Event DropCallResult(ByVal errorCode As Long)
Public Event Connected()
Public Event Idle()
Public Event Disconnected(ByVal Reason As Long)
Public Event Proceeding()
Public Event Dialing()
Public Event RingBack()
Public Event Unknown(ByVal HexReason As String)
Public Event SpecialInfo(ByVal HexReason As String)
Public Event GotDigit(ByVal Digit As Integer)





'line selection variables
Private m_numLines As Long
Private m_CurLineID As Long


'Call-state and asynch function state variables
Private mhLine As Long
Private mhCall As Long
Private m_RequestingCall As Long
Private m_DroppingCall As Long
Private m_ReqPrivileges As Long
Private m_ReqMediaModes As Long

Private mLastDisconnectReason As Long

Private m_LineName            As String
Private m_LastTAPIError       As Long

Friend Sub LineProcHandler(ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
'Handle callbacks here in the class itself via ITapiCallbackSink
'dbgtapi   "Entering LineProcHandler " & Timer
'dbgtapi   "Msg, param1, param1, param3 "
'dbgtapi   Hex(dwMsg) & "  " & Hex(dwParam1) & "  " & Hex(dwParam2) & "  " & Hex(dwParam3)
'  Select Case dwMsg
'    Case LINE_REPLY  ' &hC
'      If dwParam1 = m_RequestingCall Then  ' success if a match
' '       dbgTAPI "LINE_REPLY-CALL REQUEST"
'        '(earlier we called the asynch lineMakeCall() function from the
'        'MakeCallAsynch method and stored the requester in m_RequestingCall
'        'now we can see that this is the asynchronous reply to that call)
'        'don't need ID anymore - it matched
'        m_RequestingCall = 0
'
'        'if it was an error make sure the line is closed
'        If dwParam2 <> 0 Then
'          Call CloseLine
'  '        dbgTAPI "ERROR: cTapiLine.LineProcHandler Call CloseLine Param2 " & GetLineErrString(dwParam2)
'        End If
'        'send result to parent via event
'        RaiseEvent MakeCallResult(dwParam2)
'
'      ElseIf dwParam1 = m_DroppingCall Then
'   '     dbgTAPI "LINE_REPLY-CALL DROP"
'        'asynch reply to lineDrop() call
'        'Call CloseLine
'        m_DroppingCall = 0
'        RaiseEvent DropCallResult(dwParam2)
'
'      End If
'
'    Case LINE_CALLSTATE  ' & h2
'      Select Case dwParam1
'        Case LINECALLSTATE_DISCONNECTED
'          'call has disconnected
'          ' announced elsewhere: dbgtapi  "LINECALLSTATE_DISCONNECTED"
'          mLastDisconnectReason = dwParam2
'          RaiseEvent Disconnected(dwParam2)
'          'Call DropCallAsynch
'
'        Case LINECALLSTATE_IDLE
'          'no call exists - the line is idle
'          dbgTAPI "LINECALLSTATE_IDLE"
'          If mhCall <> 0 Then
'            Call lineDeallocateCall(mhCall)
'            mhCall = 0
'          End If
'          RaiseEvent Idle
'
'        Case LINECALLSTATE_CONNECTED
'          RaiseEvent Connected
'
'        Case LINECALLSTATE_BUSY
'          'NOT found in modems.
'          dbgTAPI "LINECALLSTATE_BUSY"
'
'
'        Case LINECALLSTATE_DIALTONE
'          'NOT found in modems.
'          dbgTAPI "LINECALLSTATE_DIALTONE"
'
'        Case LINECALLSTATE_RINGBACK
'          'NOT found in modems.
'          dbgTAPI "LINECALLSTATE_RINGBACK"
'          RaiseEvent RingBack
'
'        Case LINECALLSTATE_DIALING
'          RaiseEvent Dialing
'
'        Case LINECALLSTATE_PROCEEDING
'          RaiseEvent Proceeding
'
'        Case LINECALLSTATE_SPECIALINFO
'          'possible network error
'          dbgTAPI "LINECALLSTATE_SPECIALINFO"
'          RaiseEvent SpecialInfo(Hex(dwParam1))
'        Case Else
'          dbgTAPI UCase("Unknown CallSTATE: " & Hex(dwParam1))
'          RaiseEvent Unknown(Hex(dwParam1))
'      End Select
'    Case LINE_MONITORDIGITS  ' & h9
'      RaiseEvent GotDigit(dwParam1)
''      dbgTAPI UCase("LINE_MONITORDIGITS " & Hex(dwParam1))
'    Case Else
'
'
'  End Select

End Sub




Public Sub DropCallAsynch()

End Sub
Public Property Get LineName() As String
  LineName = m_LineName
End Property
Public Function MakeCallAsynch(Optional ByVal PhoneNumber As String = "") As Long

End Function
Public Function CloseLine() As Boolean

End Function

Public Property Get CallInProgress() As Boolean
'CallInProgress = (mhCall <> 0)
End Property

Public Function MonitorDigits(ByVal value As String)
'  Dim Mode As Integer
'  dbgTAPI "TAPILine Class MonitorDigits Set to " & Value
'  If mhCall <> 0 Then
'    Value = LCase(Value)
'    Select Case Value
'      Case "touchtone", "dtmf"
'        Mode = LINEDIGITMODE_DTMF
'      Case "pulse", "rotary"
'        Mode = LINEDIGITMODE_PULSE
'      Case "edge", "dtmfend"
'        Mode = LINEDIGITMODE_DTMFEND
'      Case Else
'        Mode = LINEDIGITMODE_NONE
'    End Select
'    lineMonitorDigits mhCall, Mode
'  End If
End Function
Public Property Get LastDisconnectReason() As Long
  LastDisconnectReason = mLastDisconnectReason
End Property
Public Property Get LastError() As Long
  LastError = m_LastTAPIError
  'reset error value after access
  m_LastTAPIError = 0&
End Property
Public Function ErrorString(ByVal ErrCode As Long) As String
'  ErrorString = GetLineErrString(ErrCode)
End Function
Public Function Create() As Boolean

End Function

Public Function GetReasonDisconnected(ByVal LineCallState As Long) As String
'  Select Case LineCallState
'    Case LINEDISCONNECTMODE_NORMAL
'      GetReasonDisconnected = "Normal Disconnect"
'    Case LINEDISCONNECTMODE_UNKNOWN
'      GetReasonDisconnected = "Disconnected Unknown reason"
'    Case LINEDISCONNECTMODE_REJECT
'      GetReasonDisconnected = "Remote Party rejected call"
'    Case LINEDISCONNECTMODE_PICKUP
'      GetReasonDisconnected = "Disconnected Call was picked up on another phone"
'    Case LINEDISCONNECTMODE_FORWARDED
'      GetReasonDisconnected = "Disconnected Forwarded"
'    Case LINEDISCONNECTMODE_BUSY
'      GetReasonDisconnected = "Busy"
'    Case LINEDISCONNECTMODE_NOANSWER
'      GetReasonDisconnected = "No Answer"
'    Case LINEDISCONNECTMODE_BADADDRESS
'      GetReasonDisconnected = "Disconnected Bad Address"
'    Case LINEDISCONNECTMODE_UNREACHABLE
'      GetReasonDisconnected = "Disconnected Unreachable"
'    Case LINEDISCONNECTMODE_CONGESTION
'      GetReasonDisconnected = "Disconnected Congestion"
'    Case LINEDISCONNECTMODE_INCOMPATIBLE
'      GetReasonDisconnected = "Disconnected Incompatible"
'    Case LINEDISCONNECTMODE_UNAVAIL
'      GetReasonDisconnected = "Disconnected Unavailable"
'    Case LINEDISCONNECTMODE_NODIALTONE
'      GetReasonDisconnected = "No Dial Tone"
'    Case Else
'      GetReasonDisconnected = "LINECALLSTATE Bad Reason"
'  End Select

End Function
Private Sub Class_Terminate()
  Finalize "Class_Terminate"
End Sub
Public Function GetLineErrString(lparam As Long) As String
'Returns a String description of a TAPI Line Error code
  Dim msg As String

'  Select Case lparam
'    Case LINEERR_ALLOCATED  '( = &H80000001)
'      msg = "Allocated"
'    Case LINEERR_BADDEVICEID  '(= &H80000002)
'      msg = "Bad Device ID"
'    Case LINEERR_BEARERMODEUNAVAIL  '(= &H80000003)
'      msg = "Bearer Mode Unavail"
'    Case LINEERR_CALLUNAVAIL  '(= &H80000005)
'      msg = "Call UnAvail"
'    Case LINEERR_COMPLETIONOVERRUN  '(= &H80000006
'      msg = "Completion Overrun"
'    Case LINEERR_CONFERENCEFULL  '(= &H80000007
'      msg = "Conference Full"
'    Case LINEERR_DIALBILLING  '(= &H80000008
'      msg = "Dial Billing"
'    Case LINEERR_DIALDIALTONE  '(= &H80000009
'      msg = "Dial Dialtone"
'    Case LINEERR_DIALPROMPT  '(= &H8000000A
'      msg = "Dial Prompt"
'    Case LINEERR_DIALQUIET  '(= &H8000000B
'      msg = "Dial Quiet"
'    Case LINEERR_INCOMPATIBLEAPIVERSION  '(= &H8000000C
'      msg = "Incompatible API Version"
'    Case LINEERR_INCOMPATIBLEEXTVERSION  '(= &H8000000D
'      msg = "Incompatible Ext Version"
'    Case LINEERR_INIFILECORRUPT  '(= &H8000000E
'      msg = "Ini File Corrupt"
'    Case LINEERR_INUSE  '(= &H8000000F
'      msg = "In Use"
'    Case LINEERR_INVALADDRESS  '(= &H80000010
'      msg = "Invalid Address"
'    Case LINEERR_INVALADDRESSID  '(= &H80000011
'      msg = "Invalid Address ID"
'    Case LINEERR_INVALADDRESSMODE  '(= &H80000012
'      msg = "Invalid Address Mode"
'    Case LINEERR_INVALADDRESSSTATE  '(= &H80000013
'      msg = "Invalid Address State"
'    Case LINEERR_INVALAPPHANDLE  '(= &H80000014
'      msg = "Invalid App Handle"
'    Case LINEERR_INVALAPPNAME  '(= &H80000015
'      msg = "Invalid App Name"
'    Case LINEERR_INVALBEARERMODE  '(= &H80000016
'      msg = "Invalid Bearer Mode"
'    Case LINEERR_INVALCALLCOMPLMODE  '(= &H80000017
'      msg = "Invalid Call Completion Mode"
'    Case LINEERR_INVALCALLHANDLE  '(= &H80000018
'      msg = "Invalid Call Handle"
'    Case LINEERR_INVALCALLPARAMS  '(= &H80000019
'      msg = "Invalid Call Params"
'    Case LINEERR_INVALCALLPRIVILEGE  '(= &H8000001A
'      msg = "Invalid Call Privilege"
'    Case LINEERR_INVALCALLSELECT  '(= &H8000001B
'      msg = "Invalid Call Select"
'    Case LINEERR_INVALCALLSTATE  '(= &H8000001C
'      msg = "Invalid Call State"
'    Case LINEERR_INVALCALLSTATELIST  '(= &H8000001D
'      msg = "Invalid Call State List"
'    Case LINEERR_INVALCARD  '(= &H8000001E
'      msg = "Invalid Card"
'    Case LINEERR_INVALCOMPLETIONID  '(= &H8000001F
'      msg = "Invalid Completion ID"
'    Case LINEERR_INVALCONFCALLHANDLE  '(= &H80000020
'      msg = "Invalid Conf Call Handle"
'    Case LINEERR_INVALCONSULTCALLHANDLE  '(= &H80000021
'      msg = "Invalid Consult Call Handle"
'    Case LINEERR_INVALCOUNTRYCODE  '(= &H80000022
'      msg = "Invalid Country Code"
'    Case LINEERR_INVALDEVICECLASS  '(= &H80000023
'      msg = "Invalid Device Class"
'    Case LINEERR_INVALDEVICEHANDLE  '(= &H80000024
'      msg = "Invalid Device Handle"
'    Case LINEERR_INVALDIGITLIST  '(= &H80000026
'      msg = "Invalid Digit List"
'    Case LINEERR_INVALDIGITMODE  '(= &H80000027
'      msg = "Invalid Digit Mode"
'    Case LINEERR_INVALDIGITS  '(= &H80000028
'      msg = "Invalid Digits"
'    Case LINEERR_INVALEXTVERSION  '(= &H80000029
'      msg = "Invalid Ext Version"
'    Case LINEERR_INVALGROUPID  '(= &H8000002A
'      msg = "Invalid Group ID"
'    Case LINEERR_INVALLINEHANDLE  '(= &H8000002B
'      msg = "Invalid Line Handle"
'    Case LINEERR_INVALLINESTATE  '(= &H8000002C
'      msg = "Invalid Line State"
'    Case LINEERR_INVALLOCATION  '(= &H8000002D
'      msg = "Invalid Location"
'    Case LINEERR_INVALMEDIALIST  '(= &H8000002E
'      msg = "Invalid Media List"
'    Case LINEERR_INVALMEDIAMODE  '(= &H8000002F
'      msg = "Invalid Media Mode"
'    Case LINEERR_INVALMESSAGEID  '(= &H80000030
'      msg = "Invalid Message ID"
'    Case LINEERR_INVALPARAM  '(= &H80000032
'      msg = "Invalid Param"
'    Case LINEERR_INVALPARKID  '(= &H80000033
'      msg = "Invalid Park ID"
'    Case LINEERR_INVALPARKMODE  '(= &H80000034
'      msg = "Invalid Park Mode"
'    Case LINEERR_INVALPOINTER  '(= &H80000035
'      msg = "Invalid Pointer"
'    Case LINEERR_INVALPRIVSELECT  '(= &H80000036
'      msg = "Invalid Priv Select"
'    Case LINEERR_INVALRATE  '(= &H80000037
'      msg = "Invalid Rate"
'    Case LINEERR_INVALREQUESTMODE  '(= &H80000038
'      msg = "Invalid Request Mode"
'    Case LINEERR_INVALTERMINALID  '(= &H80000039
'      msg = "Invalid Terminal ID"
'    Case LINEERR_INVALTERMINALMODE  '(= &H8000003A
'      msg = "Invalid Terminal Mode"
'    Case LINEERR_INVALTIMEOUT  '(= &H8000003B
'      msg = "Invalid Time Out"
'    Case LINEERR_INVALTONE  '(= &H8000003C
'      msg = "Invalid Tone"
'    Case LINEERR_INVALTONELIST  '(= &H8000003D
'      msg = "Invalid Tone List"
'    Case LINEERR_INVALTONEMODE  '(= &H8000003E
'      msg = "Invalid Tone Mode"
'    Case LINEERR_INVALTRANSFERMODE  '(= &H8000003F
'      msg = "Invalid Transfer Mode"
'    Case LINEERR_LINEMAPPERFAILED  '(= &H80000040
'      msg = "Line Mapper Failed"
'    Case LINEERR_NOCONFERENCE  '(= &H80000041
'      msg = "No Conference"
'    Case LINEERR_NODEVICE  '(= &H80000042
'      msg = "No Device"
'    Case LINEERR_NODRIVER  '(= &H80000043
'      msg = "No Driver"
'    Case LINEERR_NOMEM  '(= &H80000044
'      msg = "No Memory"
'    Case LINEERR_NOREQUEST  '(= &H80000045
'      msg = "No Request"
'    Case LINEERR_NOTOWNER  '(= &H80000046
'      msg = "Not Owner"
'    Case LINEERR_NOTREGISTERED  '(= &H80000047
'      msg = "Not Registered"
'    Case LINEERR_OPERATIONFAILED  '(= &H80000048
'      msg = "Operation Failed"
'    Case LINEERR_OPERATIONUNAVAIL  '(= &H80000049
'      msg = "Operation Unavailable"
'    Case LINEERR_RATEUNAVAIL  '(= &H8000004A
'      msg = "Rate Unavailable"
'    Case LINEERR_RESOURCEUNAVAIL  '(= &H8000004B
'      msg = "Resource Unavailable"
'    Case LINEERR_REQUESTOVERRUN  '(= &H8000004C
'      msg = "Request Overrun"
'    Case LINEERR_STRUCTURETOOSMALL  '(= &H8000004D
'      msg = "Structure Too Small"
'    Case LINEERR_TARGETNOTFOUND  '(= &H8000004E
'      msg = "Target Not found"
'    Case LINEERR_TARGETSELF  '(= &H8000004F
'      msg = "Target Self"
'    Case LINEERR_UNINITIALIZED  '(= &H80000050
'      msg = "Uninitialized"
'    Case LINEERR_USERUSERINFOTOOBIG  '(= &H80000051
'      msg = "UserUser Info Too Big"
'    Case LINEERR_REINIT  '(= &H80000052
'      msg = "Re-init"
'    Case LINEERR_ADDRESSBLOCKED  '(= &H80000053
'      msg = "Address Blocked"
'    Case LINEERR_BILLINGREJECTED  '(= &H80000054
'      msg = "Billing Rejected"
'    Case LINEERR_INVALFEATURE  '(= &H80000055
'      msg = "Invalid Feature"
'    Case LINEERR_NOMULTIPLEINSTANCE  '(= &H80000056
'      msg = "No Multiple Instance"
'    Case Else
'      msg = "Unknown Error"  ' undefined
'  End Select

  GetLineErrString = msg
End Function
Sub Finalize(Optional ByVal Source As String)
'  dbgTAPI "TAPILINE Finalize " & Source
'  If mhCall <> 0 Then
'    Call DropCallAsynch
'  End If
'
'  If mhCall <> 0 Then
'    Call lineDeallocateCall(mhCall)
'    mhCall = 0
'  End If
'
'  If mhLine <> 0 Then
'    Call CloseLine
'    mhLine = 0
'  End If
'
'  If mhLineApp <> 0 Then
'    Call lineShutdown(mhLineApp)
'    mhLineApp = 0
'  End If
'  dbgTAPI "End TAPILINE Finalize " & Source
End Sub
