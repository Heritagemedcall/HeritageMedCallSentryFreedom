VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPILine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Private DigitToGet As Integer

'DrawIconEx Constants
Private Const DI_MASK          As Long = &H1&
Private Const DI_IMAGE         As Long = &H2&
Private Const DI_NORMAL        As Long = &H3&
Private Const DI_COMPAT        As Long = &H4&
Private Const DI_DEFAULTSIZE   As Long = &H8&

'Initialization, error & version control variables
Private mhLineApp As Long  'handle to TAPI
Private m_APIversions() As Long
Private m_ExtAPIversions() As LINEEXTENSIONID
Private m_ApplicationVersion As Long  'some TAPI calls need this instead of negotiated API version
Private m_LowAPI As Long  'default TAPI 1.3 (&H00010003)
Private m_HiAPI As Long  'default TAPI 3.0 (&H00030000)
Private m_LastTAPIError As Long

'line selection variables
Private m_numLines As Long
Private m_CurLineID As Long

'devcaps variables
Private m_linecaps As LineDevCaps
Private m_LineName As String
Private m_ProviderInfo As String
Private m_SwitchInfo As String
Private m_PermanentLineID As Long
Private m_StringFormat As Long
Private m_numAddresses As Long
Private m_maxDataRate As Long
Private m_BearerModes As Long
Private m_AddressModes As Long
Private m_MediaModes As Long
Private m_GenerateToneMaxNumFreq As Long
Private m_GenerateToneModes As Long
Private m_numTerminals As Long

'Call-state and asynch function state variables
Private mhLine As Long
Private mhCall As Long
Private m_RequestingCall As Long
Private m_DroppingCall As Long
Private m_ReqPrivileges As Long
Private m_ReqMediaModes As Long

Private mLastDisconnectReason As Long



'*************** TAPI *************

Private Declare Function lineInitialize Lib "tapi32.dll" (ByRef lphLineApp As Long, ByVal hInstance As Long, ByVal lpfnCallback As Long, ByVal lpszAppName As String, ByRef lpdwNumDevs As Long) As Long
Private Declare Function lineInitializeEx Lib "tapi32.dll" Alias "lineInitializeExA" (ByRef lphLineApp As Long, ByVal hInstance As Long, ByVal lpfnCallback As Long, ByVal lpszFriendlyAppName As String, ByRef lpdwNumDevs As Long, ByRef lpdwAPIVersion As Long, ByRef lpLineInitializeExParams As LINEINITIALIZEEXPARAMS) As Long
Private Declare Function lineGetDevCaps Lib "tapi32.dll" Alias "lineGetDevCapsA" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByRef lpLineDevCaps As LineDevCaps) As Long
Private Declare Function lineConfigDialog Lib "tapi32.dll" Alias "lineConfigDialogA" (ByVal dwDeviceID As Long, ByVal hwndOwner As Long, ByVal lpszDeviceClass As String) As Long
Private Declare Function lineTranslateDialog Lib "tapi32.dll" Alias "lineTranslateDialogA" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal hwndOwner As Long, ByVal lpszAddressIn As String) As Long
Private Declare Function lineShutdown Lib "tapi32.dll" (ByVal hLineApp As Long) As Long

Private Declare Function lineMakeCall Lib "tapi32.dll" Alias "lineMakeCallA" (ByVal hLine As Long, ByRef lphCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long, ByRef lpCallParams As Any) As Long  'LINECALLPARAMS declared As Any so Null value can be passed

Private Declare Function lineDeallocateCall Lib "tapi32.dll" (ByVal hCall As Long) As Long
Private Declare Function lineDrop Lib "tapi32.dll" (ByVal hCall As Long, ByVal lpsUserUserInfo As String, ByVal dwSize As Long) As Long
Private Declare Function lineGetIcon Lib "tapi32.dll" Alias "lineGetIconA" (ByVal dwDeviceID As Long, ByVal lpszDeviceClass As Long, ByRef lphIcon As Long) As Long

Private Declare Function lineClose Lib "tapi32.dll" (ByVal hLine As Long) As Long
Private Declare Function lineOpen Lib "tapi32.dll" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByRef lphLine As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByVal dwCallbackInstance As Long, ByVal dwPrivileges As Long, ByVal dwMediaModes As Long, ByRef lpCallParams As Any) As Long  'LINECALLPARAMS declared As Any so NULL can be passed

Private Declare Function lineAnswer Lib "Tapi32" (ByVal hCall As Long, ByRef lpsUserUserInfo As String, ByVal dwSize As Long) As Long
Private Declare Function lineGetID Lib "Tapi32" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, ByRef lpDevice As varString, ByVal lpszDeviceClass As String) As Long
Private Declare Function lineMonitorDigits Lib "Tapi32" (ByVal hCall As Long, ByVal dwDigitModes As Long) As Long
Private Declare Function lineGenerateDigits Lib "Tapi32" (ByVal hCall As Long, ByVal dwDigitMode As Long, ByVal lpszDigits As String, ByVal dwDuration As Long) As Long
Private Declare Function lineGetCallInfo Lib "Tapi32" (ByVal hCall As Long, ByRef lpCallInf As lineCallInfo) As Long
Private Declare Function lineSetStatusMessages Lib "Tapi32" (ByVal hLine As Long, ByVal dwLineStates As Long, ByVal dwAddressStates As Long) As Long

Private Declare Function lineNegotiateAPIVersion Lib "tapi32.dll" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPILowVersion As Long, ByVal dwAPIHighVersion As Long, ByRef lpdwAPIVersion As Long, ByRef lpExtensionID As LINEEXTENSIONID) As Long

Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal left As Long, ByVal top As Long, ByVal hIcon As Long, ByVal Width As Long, ByVal Height As Long, ByVal step As Long, ByVal hBrush As Long, ByVal uFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long  'BOOL

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pvDest As Any, ByVal pvSource As Any, ByVal lBytes As Long)



'*********** TAPI types ***********

Private Type LINETERMCAPS
  dwTermDev As Long
  dwTermModes As Long
  dwTermSharing As Long
End Type


Private Type LINEINITIALIZEEXPARAMS
  dwTotalSize As Long
  dwNeededSize As Long
  dwUsedSize As Long
  dwOptions As Long
  hEvent As Long  'union hEvent and Completion port
  dwCompletionKey As Long
End Type

Private Type LINEEXTENSIONID
  dwExtensionID0 As Long
  dwExtensionID1 As Long
  dwExtensionID2 As Long
  dwExtensionID3 As Long
End Type


Private Type LINEDIALPARAMS
  dwDialPause As Long
  dwDialSpeed As Long
  dwDigitDuration As Long
  dwWaitForDialtone As Long
End Type

Private Type LineDevCaps
  dwTotalSize As Long
  dwNeededSize As Long
  dwUsedSize As Long
  dwProviderInfoSize As Long
  dwProviderInfoOffset As Long
  dwSwitchInfoSize As Long
  dwSwitchInfoOffset As Long
  dwPermanentLineID As Long
  dwLineNameSize As Long
  dwLineNameOffset As Long
  dwStringFormat As Long
  dwAddressModes As Long
  dwNumAddresses As Long
  dwBearerModes As Long
  dwMaxRate As Long
  dwMediaModes As Long
  dwGenerateToneModes As Long
  dwGenerateToneMaxNumFreq As Long
  dwGenerateDigitModes As Long
  dwMonitorToneMaxNumFreq As Long
  dwMonitorToneMaxNumEntries As Long
  dwMonitorDigitModes As Long
  dwGatherDigitsMinTimeout As Long
  dwGatherDigitsMaxTimeout As Long
  dwMedCtlDigitMaxListSize As Long
  dwMedCtlMediaMaxListSize As Long
  dwMedCtlToneMaxListSize As Long
  dwMedCtlCallStateMaxListSize As Long
  dwDevCapFlags As Long
  dwMaxNumActiveCalls As Long
  dwAnswerMode As Long
  dwRingModes As Long
  dwLineStates As Long
  dwUUIAcceptSize As Long
  dwUUIAnswerSize As Long
  dwUUIMakeCallSize As Long
  dwUUIDropSize As Long
  dwUUISendUserUserInfoSize As Long
  dwUUICallInfoSize As Long
  MinDialParams As LINEDIALPARAMS
  MaxDialParams As LINEDIALPARAMS
  DefaultDialParams As LINEDIALPARAMS
  dwNumTerminals As Long
  dwTerminalCapsSize As Long
  dwTerminalCapsOffset As Long
  dwTerminalTextEntrySize As Long
  dwTerminalTextSize As Long
  dwTerminalTextOffset As Long
  dwDevSpecificSize As Long
  dwDevSpecificOffset As Long
  dwLineFeatures As Long  '// TAPI v1.4
  '#if (TAPI_CURRENT_VERSION >= 0x00020000)
  dwSettableDevStatus As Long  '// TAPI v2.0
  dwDeviceClassesSize As Long  ' // TAPI v2.0
  dwDeviceClassesOffset As Long  ' // TAPI v2.0
  '#End If
  'my way of handling TAPI variable sized structures (yech!)
  vbByteBuffer(0 To 2048) As Byte
  'note*  if you get LINEERR_STRUCTURETOOSMALL and you know that you are
  'doing everything else right (Like initializing the dwActualSize parameter
  'of structs you are passing) then you *might* need to increase this buffer
  'size and recompile.  However, I have not had any problems with this size yet.
End Type




Private Type LINECALLPARAMS                 '// DEFAULTS
  dwTotalSize As Long  '// ---------
  dwBearerMode As Long  '// voice
  dwMinRate As Long  '// (3.1kHz)
  dwMaxRate As Long  '// (3.1kHz)
  dwMediaMode As Long  '// interactiveVoice
  dwCallParamFlags As Long  '// 0
  dwAddressMode As Long  '// addressID
  dwAddressID As Long  '// (any available)
  DialParams As LINEDIALPARAMS  '// (0, 0, 0, 0)
  dwOrigAddressSize As Long  '// 0
  dwOrigAddressOffset As Long
  dwDisplayableAddressSize As Long
  dwDisplayableAddressOffset As Long
  dwCalledPartySize As Long  '// 0
  dwCalledPartyOffset As Long
  dwCommentSize As Long  '// 0
  dwCommentOffset As Long
  dwUserUserInfoSize As Long  '// 0
  dwUserUserInfoOffset As Long
  dwHighLevelCompSize As Long  '// 0
  dwHighLevelCompOffset As Long
  dwLowLevelCompSize As Long  '// 0
  dwLowLevelCompOffset As Long
  dwDevSpecificSize As Long  '// 0
  dwDevSpecificOffset As Long
  '#if (TAPI_CURRENT_VERSION >= 0x00020000)
  dwPredictiveAutoTransferStates As Long  '// TAPI v2.0
  dwTargetAddressSize As Long  '// TAPI v2.0
  dwTargetAddressOffset As Long  '// TAPI v2.0
  dwSendingFlowspecSize As Long  '// TAPI v2.0
  dwSendingFlowspecOffset As Long  '// TAPI v2.0
  dwReceivingFlowspecSize As Long  '// TAPI v2.0
  dwReceivingFlowspecOffset As Long  '// TAPI v2.0
  dwDeviceClassSize As Long  '// TAPI v2.0
  dwDeviceClassOffset As Long  '// TAPI v2.0
  dwDeviceConfigSize As Long  '// TAPI v2.0
  dwDeviceConfigOffset As Long  '// TAPI v2.0
  dwCallDataSize As Long  '// TAPI v2.0
  dwCallDataOffset As Long  '// TAPI v2.0
  dwNoAnswerTimeout As Long  '// TAPI v2.0
  dwCallingPartyIDSize As Long  '// TAPI v2.0
  dwCallingPartyIDOffset As Long  '// TAPI v2.0
  '#End If
End Type

Private Type lineCallInfo

  dwTotalSize As Long
  dwNeededSize As Long
  dwUsedSize As Long
  hLine As Long
  dwLineDeviceID As Long
  dwAddressID As Long
  dwBearerMode As Long
  dwRate As Long
  dwMediaMode As Long
  dwAppSpecific As Long
  dwCallID As Long
  dwRelatedCallID As Long
  dwCallParamFlags As Long
  dwCallStates As Long
  dwMonitorDigitModes As Long
  dwMonitorMediaModes As Long
  DialParams As LINEDIALPARAMS
  dwOrigin As Long
  dwReason As Long
  dwCompletionID As Long
  dwNumOwners As Long
  dwNumMonitors As Long
  dwCountryCode As Long
  dwTrunk As Long
  dwCallerIDFlags As Long
  dwCallerIDSize As Long
  dwCallerIDOffset As Long
  dwCallerIDNameSize As Long
  dwCallerIDNameOffset As Long
  dwCalledIDFlags As Long
  dwCalledIDSize As Long
  dwCalledIDOffset As Long
  dwCalledIDNameSize As Long
  dwCalledIDNameOffset As Long
  dwConnectedIDFlags As Long
  dwConnectedIDSize As Long
  dwConnectedIDOffset As Long
  dwConnectedIDNameSize As Long
  dwConnectedIDNameOffset As Long
  dwRedirectionIDFlags As Long
  dwRedirectionIDSize As Long
  dwRedirectionIDOffset As Long
  dwRedirectionIDNameSize As Long
  dwRedirectionIDNameOffset As Long
  dwRedirectingIDFlags As Long
  dwRedirectingIDSize As Long
  dwRedirectingIDOffset As Long
  dwRedirectingIDNameSize As Long
  dwRedirectingIDNameOffset As Long
  dwAppNameSize As Long
  dwAppNameOffset As Long
  dwDisplayableAddressSize As Long
  dwDisplayableAddressOffset As Long
  dwCalledPartySize As Long
  dwCalledPartyOffset As Long
  dwCommentSize As Long
  dwCommentOffset As Long
  dwDisplaySize As Long
  dwDisplayOffset As Long
  dwUserUserInfoSize As Long
  dwUserUserInfoOffset As Long
  dwHighLevelCompSize As Long
  dwHighLevelCompOffset As Long
  dwLowLevelCompSize As Long
  dwLowLevelCompOffset As Long
  dwChargingInfoSize As Long
  dwChargingInfoOffset As Long
  dwTerminalModesSize As Long
  dwTerminalModesOffset As Long
  dwDevSpecificSize As Long
  dwDevSpecificOffset As Long
  ' tapi > 2.0
  dwCallTreatment As Long  ' TAPI v2.0
  dwCallDataSize As Long  ' TAPI v2.0
  dwCallDataOffset As Long  ' TAPI v2.0
  dwSendingFlowspecSize As Long  ' TAPI v2.0
  dwSendingFlowspecOffset As Long  ' TAPI v2.0
  dwReceivingFlowspecSize As Long  ' TAPI v2.0
  dwReceivingFlowspecOffset As Long  ' TAPI v2.0

  bBytes(2000) As Byte  'Added to TAPI structure for callinfo data.
End Type

Const LINEINITIALIZEEXOPTION_USEHIDDENWINDOW   As Long = &H1&  '// TAPI v2.0
Const LINEINITIALIZEEXOPTION_USEEVENT    As Long = &H2&  '// TAPI v2.0
Const LINEINITIALIZEEXOPTION_USECOMPLETIONPORT    As Long = &H3&  '// TAPI v2.0

Const LINEBEARERMODE_VOICE = &H1&
Const LINEBEARERMODE_SPEECH = &H2&
Const LINEBEARERMODE_MULTIUSE = &H4&
Const LINEBEARERMODE_DATA = &H8&
Const LINEBEARERMODE_ALTSPEECHDATA = &H10&
Const LINEBEARERMODE_NONCALLSIGNALING = &H20&


Const LINEDISCONNECTMODE_NORMAL = &H1
Const LINEDISCONNECTMODE_UNKNOWN = &H2
Const LINEDISCONNECTMODE_REJECT = &H4
Const LINEDISCONNECTMODE_PICKUP = &H8
Const LINEDISCONNECTMODE_FORWARDED = &H10
Const LINEDISCONNECTMODE_BUSY = &H20
Const LINEDISCONNECTMODE_NOANSWER = &H40
Const LINEDISCONNECTMODE_BADADDRESS = &H80
Const LINEDISCONNECTMODE_UNREACHABLE = &H100
Const LINEDISCONNECTMODE_CONGESTION = &H200
Const LINEDISCONNECTMODE_INCOMPATIBLE = &H400
Const LINEDISCONNECTMODE_UNAVAIL = &H800
Const LINEDISCONNECTMODE_NODIALTONE = &H1000
Const LINEDISCONNECTMODE_NUMBERCHANGED = &H2000
Const LINEDISCONNECTMODE_OUTOFORDER = &H4000
Const LINEDISCONNECTMODE_TEMPFAILURE = &H8000
Const LINEDISCONNECTMODE_QOSUNAVAIL = &H10000
Const LINEDISCONNECTMODE_BLOCKED = &H20000
Const LINEDISCONNECTMODE_DONOTDISTURB = &H40000
Const LINEDISCONNECTMODE_CANCELLED = &H80000

Public Property Get hCall() As Long
  hCall = mhCall
End Property

Public Property Let hCall(ByVal Value As Long)
  
  LogProgramError "STATUS at CTAPILine.hcall " & Value
  mhCall = Value
End Property



Public Function MonitorDigits(ByVal Value As String)
        Dim Mode As Integer
10      On Error GoTo MonitorDigits_Error

20      dbgTAPI "TAPILine Class MonitorDigits Set to " & Value
30      If mhCall <> 0 Then
40        Value = LCase(Value)
50        Select Case Value
            Case "touchtone", "dtmf"
60            Mode = LINEDIGITMODE_DTMF
70          Case "pulse", "rotary"
80            Mode = LINEDIGITMODE_PULSE
90          Case "edge", "dtmfend"
100           Mode = LINEDIGITMODE_DTMFEND
110         Case Else
120           Mode = LINEDIGITMODE_NONE
130       End Select
140       lineMonitorDigits mhCall, Mode
150     End If

MonitorDigits_Resume:

160     On Error GoTo 0
170     Exit Function

MonitorDigits_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.MonitorDigits." & Erl
190     Resume MonitorDigits_Resume
End Function


'Private Type varString
'  dwTotalSize As Long
'  dwNeededSize As Long
'  dwUsedSize As Long
'  dwStringFormat As Long
'  dwStringSize As Long
'  dwStringOffset As Long
'  bBytes(2000) As Byte  'HACK Added to TAPI structure for lineGetID data.
'End Type

Public Property Get IsLineOpen() As Boolean
  IsLineOpen = (mhLine <> 0)
End Property

Public Function GetLineID(ByVal DeviceClass As String) As Long
        ' DeviceClass typically "wave/out"


10      On Error GoTo GetLineID_Error

20      On Error Resume Next

        Dim rc                 As Long
        Dim LineID             As String
        Dim oVar               As varString
        Dim lineGetIdError     As String
        
        'Type varString
        '  dwTotalSize As Long
        '  dwNeededSize As Long
        '  dwUsedSize As Long
        '  dwStringFormat As Long
        '  dwStringSize As Long
        '  dwStringOffset As Long
        '  bBytes(2000) As Byte  'Added to TAPI structure for lineGetID data.
        'End Type
        
        'If a string cannot be returned in a variable structure, the dwStringSize and dwStringOffset members are set in one of the following ways:
        'dwStringSize and dwStringOffset members are both set to zero.
        'dwStringOffset is nonzero and dwStringSize is zero.
        'dwStringOffset is nonzero, dwStringSize is 1, and the byte at the given offset is zero.
        

30      oVar.dwTotalSize = Len(oVar)

40      dbgTAPI "hline, hcall " & mhLine & " " & mhCall


        'Possible Errors (all -n)
        'LINEERR_INVALLINEHANDLE, LINEERR_NOMEM, LINEERR_INVALADDRESSID, LINEERR_OPERATIONUNAVAIL, LINEERR_INVALCALLHANDLE, LINEERR_OPERATIONFAILED, LINEERR_INVALCALLSELECT
        'LINEERR_INVALDEVICECLASS, LINEERR_RESOURCEUNAVAIL, LINEERR_INVALPOINTER, LINEERR_STRUCTURETOOSMALL, LINEERR_NODEVICE, LINEERR_UNINITIALIZED
        'https://msdn.microsoft.com/en-us/library/ms735751(v=vs.85).aspx
        
                    ' hLine,AddressID,hCall,Select,(ref)Device,DeviceClass
50      rc = lineGetID(mhLine, 0, mhCall, LINECALLSELECT_CALL, oVar, DeviceClass)
      
        'LogProgramError "STATUS at CTAPILine.GetLineID.lineGetID Line: " & mhLine & " Call: " & mhCall & " RC: " & rc


60      If rc <> 0 Then

70        Select Case rc
            Case LINEERR_INVALLINEHANDLE
80            lineGetIdError = "LINEERR_INVALLINEHANDLE"
90          Case LINEERR_NOMEM
100           lineGetIdError = "LINEERR_NOMEM"
110         Case LINEERR_INVALADDRESSID
120           lineGetIdError = "LINEERR_INVALADDRESSID"
130         Case LINEERR_OPERATIONUNAVAIL
140           lineGetIdError = "LINEERR_OPERATIONUNAVAIL"
150         Case LINEERR_INVALCALLHANDLE
160           lineGetIdError = "INVALCALLHANDLE"
170         Case LINEERR_OPERATIONFAILED
180           lineGetIdError = "LINEERR_OPERATIONFAILED"
190         Case LINEERR_INVALCALLSELECT
200           lineGetIdError = "LINEERR_INVALCALLSELECT"
210         Case LINEERR_INVALDEVICECLASS
220           lineGetIdError = "LINEERR_INVALDEVICECLASS"
230         Case LINEERR_RESOURCEUNAVAIL
240           lineGetIdError = "LINEERR_RESOURCEUNAVAIL"
250         Case LINEERR_INVALPOINTER
260           lineGetIdError = "LINEERR_INVALPOINTER"
270         Case LINEERR_STRUCTURETOOSMALL
280           lineGetIdError = "LINEERR_STRUCTURETOOSMALL"
290         Case LINEERR_NODEVICE
300           lineGetIdError = "LINEERR_NODEVICE"
310         Case LINEERR_UNINITIALIZED
320           lineGetIdError = "LINEERR_UNINITIALIZED"
330         Case Else
340           lineGetIdError = "LINEERR_UNKNOWN"
350       End Select


360        dbgTAPI "Error " & lineGetIdError & " ( " & rc & ") " & " at CTAPILine.GetLineID.60 Line: " & mhLine & " Call: " & mhCall

          '  ProcessTAPIError RC
370     Else
380       If oVar.dwStringOffset = 0 Then  'Nothing to get!
390         GetLineID = -1
400         Exit Function
410       End If
420       LineID = Trim(left(oVar.bBytes(0), oVar.dwStringSize))
430       GetLineID = Val(LineID)

440     End If

GetLineID_Resume:

450     On Error GoTo 0
460     Exit Function

GetLineID_Error:

470     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.GetLineID." & Erl
480     Resume GetLineID_Resume

End Function


Public Property Get LastDisconnectReason() As Long
  LastDisconnectReason = mLastDisconnectReason
End Property

Private Sub Class_Initialize()

10      dbgTAPI "TAPILine Class Init"
20      mLastDisconnectReason = 1
        'default TAPI 1.3 - 3.0
30      m_LowAPI = &H10003
40      m_HiAPI = &H30000

50      Create
End Sub
Public Property Get LowAPI() As Long
  LowAPI = m_LowAPI
End Property
Public Property Let LowAPI(ByVal ver As Long)
  m_LowAPI = ver
End Property

Public Property Get HiAPI() As Long
  HiAPI = m_HiAPI
End Property
Public Property Let HiAPI(ByVal ver As Long)
  m_HiAPI = ver
End Property


Public Property Get LastError() As Long
  LastError = m_LastTAPIError
  'reset error value after access
  m_LastTAPIError = 0&
End Property

Public Function ErrorString(ByVal ErrCode As Long) As String
  ErrorString = GetLineErrString(ErrCode)
End Function

Public Function Create() As Boolean

        Dim rc As Long
        Dim line As Long
        Dim lip As LINEINITIALIZEEXPARAMS
        Dim lxid As LINEEXTENSIONID

        'init params
10      lip.dwTotalSize = Len(lip)
20      lip.dwOptions = LINEINITIALIZEEXOPTION_USEHIDDENWINDOW

30      'SpecialLog "in Tapiline.Create"
        
        'initialize TAPI lines and get handle
40      rc = lineInitializeEx(mhLineApp, App.hInstance, AddressOf modTapiLine.LineCallbackProc, App.Title, m_numLines, m_HiAPI, lip)
        
50      'SpecialLog "in Tapiline.Create lineInitializeEx"
        
60      If rc <> TAPI_SUCCESS Then
70        m_LastTAPIError = rc
80        'SpecialLog "in Tapiline.Create Exit 80"
90        Exit Function
100     Else
          
110       If m_numLines < 1 Then
              'SpecialLog "in Tapiline.Create Exit 110"
              Exit Function  'no lines!
          End If
          'negotiate and cache API versions for each line
120       ReDim m_APIversions(0 To m_numLines - 1)
130       ReDim m_ExtAPIversions(0 To m_numLines - 1)
          'SpecialLog "in Tapiline.Create Numlines " & m_numLines
140       For line = 0 To m_numLines - 1
            'SpecialLog "in Tapiline.Create LIne " & line
150         rc = lineNegotiateAPIVersion(mhLineApp, line, m_LowAPI, m_HiAPI, m_APIversions(line), lxid)
            'SpecialLog "in Tapiline.Create lineNegotiateAPIVersion"
160         If rc <> TAPI_SUCCESS Then
              'SpecialLog "in Tapiline.Create TAPI FAIL 160"
170           m_APIversions(line) = 0  'no compatible API negotiated for this line
180           rc = 0
190         Else
              'also cache extension version (dev-specific features) if available
200           m_ExtAPIversions(line).dwExtensionID0 = lxid.dwExtensionID0
210           m_ExtAPIversions(line).dwExtensionID1 = lxid.dwExtensionID1
220           m_ExtAPIversions(line).dwExtensionID2 = lxid.dwExtensionID2
230           m_ExtAPIversions(line).dwExtensionID3 = lxid.dwExtensionID3
              'SpecialLog "in Tapiline.Create after 230"
              'TODO! add call to lineNegotiateExtVersion()
              'and save the highest negotiated version as our app version
240           If m_APIversions(line) > m_ApplicationVersion Then
                
250             m_ApplicationVersion = m_APIversions(line)
                'SpecialLog "in Tapiline.Create after m_ApplicationVersion 250"
260           End If
270         End If

280       Next
290     End If
        'SpecialLog "in Tapiline.Create Call  GetLineDevCaps  290"
300     rc = GetLineDevCaps()
        'SpecialLog "in Tapiline.Create Returning True"
310     Create = True  'return success

End Function

Public Property Get numLines() As Long
  numLines = m_numLines
End Property

Public Property Get CurrentLineID() As Long
  CurrentLineID = m_CurLineID
End Property

Public Property Let CurrentLineID(ByVal ID As Long)

'called once for each device found
  If ID < 0 Then Err.Raise 380
  If ID > m_numLines - 1 Then Err.Raise 380
  m_CurLineID = ID
  Call GetLineDevCaps
End Property

Public Property Get NegotiatedAPIVersion() As Long
  NegotiatedAPIVersion = m_APIversions(m_CurLineID)
End Property

Public Property Get ExtAPI_ID0() As Long
  ExtAPI_ID0 = m_ExtAPIversions(m_CurLineID).dwExtensionID0
End Property

Public Property Get ExtAPI_ID1() As Long
  ExtAPI_ID1 = m_ExtAPIversions(m_CurLineID).dwExtensionID1
End Property

Public Property Get ExtAPI_ID2() As Long
  ExtAPI_ID2 = m_ExtAPIversions(m_CurLineID).dwExtensionID2
End Property

Public Property Get ExtAPI_ID3() As Long
  ExtAPI_ID3 = m_ExtAPIversions(m_CurLineID).dwExtensionID3
End Property

Private Function GetLineDevCaps() As Boolean
        Dim rc As Long

        'init required fields in devcaps UDT
10      On Error GoTo GetLineDevCaps_Error

20      m_linecaps.dwTotalSize = Len(m_linecaps)

        'the LINEDEVCAPS UDT is just hard-allocated to be big enough to hold most
        'variable-length structures that TAPI may throw at you

        'request TAPI to fill UDT with info
30      rc = lineGetDevCaps(mhLineApp, m_CurLineID, m_APIversions(m_CurLineID), 0&, m_linecaps)
40      If rc <> TAPI_SUCCESS Then
50        m_LastTAPIError = rc
60        Exit Function
70      End If

        'store UDT info in class local variables


80      m_LineName = Replace(GetTAPIStructString(VarPtr(m_linecaps), m_linecaps.dwLineNameOffset, m_linecaps.dwLineNameSize), vbNullChar, "")
        '= Zoom V.92 PCI Voice Faxmodem
90      m_ProviderInfo = Replace(GetTAPIStructString(VarPtr(m_linecaps), m_linecaps.dwProviderInfoOffset, m_linecaps.dwProviderInfoSize), vbNullChar, "")
        '= Windows Telephony Service Provider for Universal Modem Driver
100     m_SwitchInfo = GetTAPIStructString(VarPtr(m_linecaps), m_linecaps.dwSwitchInfoOffset, m_linecaps.dwSwitchInfoSize)
110     m_PermanentLineID = m_linecaps.dwPermanentLineID  '327260
120     m_StringFormat = m_linecaps.dwStringFormat  '1 for my machine
130     m_numAddresses = m_linecaps.dwNumAddresses  '1
140     m_maxDataRate = m_linecaps.dwMaxRate  '56000
150     m_BearerModes = m_linecaps.dwBearerModes  '65 ' bitfield of voice enabled capabilities etc
160     m_AddressModes = m_linecaps.dwAddressModes  '1
170     m_MediaModes = m_linecaps.dwMediaModes  '30
180     m_GenerateToneMaxNumFreq = m_linecaps.dwGenerateToneMaxNumFreq  '0
190     m_GenerateToneModes = m_linecaps.dwGenerateToneModes  '0
200     m_numTerminals = m_linecaps.dwNumTerminals  '0
        

        'return success
210     GetLineDevCaps = True

GetLineDevCaps_Resume:

220     On Error GoTo 0
230     Exit Function

GetLineDevCaps_Error:

240     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.GetLineDevCaps." & Erl
250     Resume GetLineDevCaps_Resume
End Function
Public Property Get LineName() As String
  LineName = m_LineName
End Property
Public Property Get ProviderInfo() As String
  ProviderInfo = m_ProviderInfo
End Property
Public Property Get SwitchInfo() As String
  SwitchInfo = m_SwitchInfo
End Property
Public Property Get PermanentLineID() As Long
  PermanentLineID = m_PermanentLineID
End Property

Public Property Get StringFormat() As Long
  StringFormat = m_StringFormat
End Property
Public Property Get numAddresses() As Long
  numAddresses = m_numAddresses
End Property

Public Property Get maxDataRate() As Long
  maxDataRate = m_maxDataRate
End Property
Public Property Get BearerModes() As Long
  BearerModes = m_BearerModes
End Property
Public Property Get AddressModes() As Long
  AddressModes = m_AddressModes
End Property
Public Property Get mediamodes() As Long
  mediamodes = m_MediaModes
End Property
Public Property Get GenerateToneMaxNumFreq() As Long
  GenerateToneMaxNumFreq = m_GenerateToneMaxNumFreq
End Property
Public Property Get GenerateToneModes() As Long
  GenerateToneModes = m_GenerateToneModes
End Property
Public Property Get numTerminals() As Long
  numTerminals = m_numTerminals
End Property

Public Property Get LineSupportsVoiceCalls() As Boolean
  If m_BearerModes And LINEBEARERMODE_VOICE Then
    If m_MediaModes And LINEMEDIAMODE_INTERACTIVEVOICE Then
      LineSupportsVoiceCalls = True
      'Debug.Print "Voice Supported " & m_LineName, m_ProviderInfo
    Else
      LineSupportsVoiceCalls = False
      'Debug.Print "Voice NOT Supported MM " & m_LineName, m_ProviderInfo
    End If
  Else
    'Debug.Print "Voice NOT Supported BM " & m_LineName, m_ProviderInfo
  End If
End Property


Public Function OpenLine(Optional ByVal privileges As Long = LINECALLPRIVILEGE_NONE, Optional ByVal mediamodes As Long = LINEMEDIAMODE_INTERACTIVEVOICE) As Boolean
      'returns false on error or if a line is already open
      'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
10      On Error GoTo OpenLine_Error

20      If mhLine <> 0 Then
30        dbgTAPI "OpenLine called, line already open"
40        Exit Function
50      End If
        Dim rc As Long

        'open the line for outgoing call
        '(passes a reference to itself in the dwCallbackinstance parameter
        'tapi will pass that reference back to the callback procedure and
        'we will use the ITapiCallbackSink interface to call back to this
        'class instance)
60      rc = lineOpen(mhLineApp, m_CurLineID, mhLine, m_APIversions(m_CurLineID), 0&, ByVal ObjPtr(Me), privileges, mediamodes, ByVal 0&)
70      If rc <> TAPI_SUCCESS Then
80        m_LastTAPIError = rc
90        dbgTAPI ErrorString(rc)
100       Exit Function
110     Else
120       dbgTAPI "LINE OPEN " & mhLine
130     End If
        'Debug.Assert mhLine
140     OpenLine = True

OpenLine_Resume:

150     On Error GoTo 0
160     Exit Function

OpenLine_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.OpenLine." & Erl
180     Resume OpenLine_Resume

End Function
Public Function MakeCallAsynch(Optional ByVal PhoneNumber As String = "") As Long
      'returns false on error or if mhLine has not been initialized(by calling OpenLine)
      'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
        Dim rc As Long
        Dim timeDue As Date
        Dim i As Long
        
        
10      On Error GoTo MakeCallAsynch_Error

      

20      mLastDisconnectReason = LINEDISCONNECTMODE_NORMAL
30      m_LastTAPIError = 0
40      m_RequestingCall = 0
        
50      dbgTAPI "TAPILine MakeCallAsynch " & PhoneNumber
        
60      If 0 = mhLine Then
70        dbgTAPI "MakeCallAsynch entered with No active line"
80        Exit Function
90      End If
        
100     timeDue = DateAdd("s", 10, Now)
        
110     DropCallAsynch
        
120     Do While mhCall <> 0
130       If i > 10000 Then
140         i = 0
150         DoEvents
160       End If
170       i = i + 1
          
180       If Now > timeDue Then
190         mLastDisconnectReason = LINEDISCONNECTMODE_CANCELLED
200         dbgTAPI "LINE MAKE CALL ERR DROP CALL TIMEOUT"
210         mhCall = 0
220         Exit Do
230       End If
240     Loop
          
        
        Dim localhcall As Long
        
        ' API calls
250     If PhoneNumber = "" Then  'just get dialtone
260       rc = lineMakeCall(mhLine, mhCall, vbNullString, 0&, ByVal 0&)
          
          
270     Else  'dial a number
280       dbgTAPI "TAPILine lineMakeCall " & PhoneNumber
290       rc = lineMakeCall(mhLine, mhCall, PhoneNumber, 0&, ByVal 0&) ' actual API Call
300     End If
        
        LogProgramError "STATUS at CTAPILine.MakeCallAsynch.300 lineMakeCall hCall = " & mhCall & " hLine = " & mhLine & " RC " & rc & " ph " & PhoneNumber
        
        
310     If rc > 0 Then
          'now placing call - completion will be signalled by a LINE_REPLY event in the callback
          'store positive request identifier in class member variable for use in the callback handler
320       m_RequestingCall = rc
330       dbgTAPI "LINE MAKE CALL " & rc
340     Else
350       m_LastTAPIError = rc
360       dbgTAPI "LINE MAKE CALL ERR " & rc
          LogProgramError "LINE MAKE CALL ERR (" & GetLineErrString(rc) & ") at CTAPILine.MakeCallAsynch.290 hCall = " & mhCall & " hLine = " & mhLine
370       Exit Function
380     End If
390     MakeCallAsynch = rc
        
        

MakeCallAsynch_Resume:

400     On Error GoTo 0
410     Exit Function

MakeCallAsynch_Error:

420     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.MakeCallAsynch." & Erl
430     Resume MakeCallAsynch_Resume
        
End Function
Public Property Get CallInProgress() As Boolean
  CallInProgress = (mhCall <> 0)
End Property


Public Function DropCallAsynch() As Boolean
      'returns false on error or if mhLine has not been initialized(by calling OpenLine)
      'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
        Dim rc As Long

10      On Error GoTo DropCallAsynch_Error

20      If mhCall = 0 Then
30        dbgTAPI "Dropcall... no Call in progress"
          'Status = STAT_IDLE  ' testing
40        Exit Function
50      End If


60      dbgTAPI "Dropcall->LineDrop hCall " & mhCall
        
70      rc = lineDrop(mhCall, vbNullString, 0&)
        
        
80      If rc > 0 Then
          'dropping call
90        m_DroppingCall = rc
100     Else
110       m_LastTAPIError = rc
          'Exit Function
120     End If
        LogProgramError "STATUS at CTAPILine.DropCallAsynch.LineDrop hCall " & hCall & " RC " & rc

130     DelayLoop

        LogProgramError "STATUS at CTAPILine.DropCallAsynch.LineDrop (delay) hCall " & hCall & " RC " & rc

140     mhCall = 0
150     dbgTAPI "Dropcall->LineDrop RC " & rc

        

160     DropCallAsynch = True
170     Debug.Print "DropCallAsynch" & Now

DropCallAsynch_Resume:

180     On Error GoTo 0
190     Exit Function

DropCallAsynch_Error:

200     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.DropCallAsynch." & Erl
210     Resume DropCallAsynch_Resume
        
End Function

Public Function CloseLine() As Boolean
      'returns false on error or if a line is already closed
      'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
        Dim rc As Long

10      If mhLine = 0 Then
20        dbgTAPI "TAPILine.CloseLine() No hLine to close 20"

30        Exit Function
40      End If
50      dbgTAPI "TAPILine.CloseLine() Closing Line 50"
60      rc = lineClose(mhLine)
        'line handle is now invalid
70      dbgTAPI "TAPILine.CloseLine() (After)  " & mhLine & IIf(rc = TAPI_SUCCESS, " OK", " ERROR " & rc)
80      mhLine = 0
        'store errors
90      If rc <> TAPI_SUCCESS Then
100       m_LastTAPIError = rc
110       CloseLine = True 'TESTING added for TESTING

120       Exit Function
130     End If
140     CloseLine = True

End Function

Public Function ConfigDialog(Optional ByVal OwnerHwnd As Long = 0&, Optional ByVal DeviceClass As String = "") As Boolean
        Dim rc As Long

10      On Error GoTo ConfigDialog_Error

20      If DeviceClass = "" Then
30        rc = lineConfigDialog(m_CurLineID, OwnerHwnd, vbNullString)
40      Else
50        rc = lineConfigDialog(m_CurLineID, OwnerHwnd, DeviceClass)
60      End If
70      If rc <> TAPI_SUCCESS Then
80        m_LastTAPIError = rc
90        Exit Function
100     End If
        'indicate success
110     ConfigDialog = True

ConfigDialog_Resume:

120     On Error GoTo 0
130     Exit Function

ConfigDialog_Error:

140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.ConfigDialog." & Erl
150     Resume ConfigDialog_Resume

End Function

Public Function DialingPropertiesDialog(Optional ByVal OwnerHwnd As Long = 0&, Optional ByVal PhoneNumber As String = "") As Boolean
        Dim rc As Long
10      On Error GoTo DialingPropertiesDialog_Error

20      If PhoneNumber = "" Then
30        rc = lineTranslateDialog(mhLineApp, m_CurLineID, m_ApplicationVersion, OwnerHwnd, vbNullString)
40      Else
50        rc = lineTranslateDialog(mhLineApp, m_CurLineID, m_ApplicationVersion, OwnerHwnd, PhoneNumber)
60      End If
70      If rc <> TAPI_SUCCESS Then
80        m_LastTAPIError = rc
90        Exit Function
100     End If
        'indicate success
110     DialingPropertiesDialog = True

DialingPropertiesDialog_Resume:

120     On Error GoTo 0
130     Exit Function

DialingPropertiesDialog_Error:

140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.DialingPropertiesDialog." & Erl
150     Resume DialingPropertiesDialog_Resume
End Function
Public Function PaintDevIcon(ByVal hDC As Long, Optional ByVal left As Long = 0, Optional ByVal top As Long = 0, Optional ByVal Width As Long = 0, Optional ByVal Height As Long = 0) As Boolean
        Dim hIcon As Long
        Dim rc As Long

10      On Error GoTo PaintDevIcon_Error

20      rc = lineGetIcon(m_CurLineID, 0&, hIcon)
30      If rc <> TAPI_SUCCESS Then
40        m_LastTAPIError = rc
50        Exit Function
60      End If
70      rc = DrawIconEx(hDC, left, top, hIcon, Width, Height, 0&, 0&, DI_NORMAL)
        '(function succeeds even when there is no icon in the TSP)
80      If 0 = hIcon Then  'function failed (check actual hIcon instead of return value)
90        Exit Function
100     End If
110     Call DestroyIcon(hIcon)

120     PaintDevIcon = True

PaintDevIcon_Resume:

130     On Error GoTo 0
140     Exit Function

PaintDevIcon_Error:

150     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.PaintDevIcon." & Erl
160     Resume PaintDevIcon_Resume

End Function


Private Function GetTAPIStructString(ByVal ptrTapistruct As Long, ByVal offset As Long, ByVal Length As Long) As String
'ugly C-hacker way to deal with ugly C-hacker TAPI structs (UDTs)
  Dim Buffer() As Byte

  If Length < 1 Then Exit Function  'handle erroneous input

  If offset Then  '
    ReDim Buffer(0 To Length - 1)
    'CopyMemory2 buffer(0), ByVal ptrTapistruct + offset, length
    CopyMemory Buffer(0), (ptrTapistruct + offset), Length
    GetTAPIStructString = StrConv(Buffer, vbUnicode)
  End If

End Function

Private Sub Class_Terminate()
  Finalize "Class_Terminate"
End Sub
Sub Finalize(Optional ByVal Source As String)
10      On Error GoTo Finalize_Error

20      dbgTAPI "TAPILINE Finalize " & Source
30      If mhCall <> 0 Then
40        Call DropCallAsynch
50      End If

60      If mhCall <> 0 Then
70        Call lineDeallocateCall(mhCall)
80        mhCall = 0
90      End If

100     If mhLine <> 0 Then
110       Call CloseLine
120       mhLine = 0
130     End If

140     If mhLineApp <> 0 Then
150       Call lineShutdown(mhLineApp)
160       mhLineApp = 0
170     End If
180     dbgTAPI "End TAPILINE Finalize " & Source

Finalize_Resume:

190     On Error GoTo 0
200     Exit Sub

Finalize_Error:

210     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.Finalize." & Erl
220     Resume Finalize_Resume
End Sub
Friend Sub LineProcHandler(ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
      'Handle callbacks here in the class itself via ITapiCallbackSink
      dbgTAPI "Entering LineProcHandler " & Timer
      dbgTAPI "Msg, param1, param1, param3 "
      dbgTAPI Hex(dwMsg) & "  " & Hex(dwParam1) & "  " & Hex(dwParam2) & "  " & Hex(dwParam3)
10      On Error GoTo LineProcHandler_Error

20      Select Case dwMsg
          Case LINE_REPLY  ' &hC
30          If dwParam1 = m_RequestingCall Then  ' success if a match
40            dbgTAPI "LINE_REPLY-CALL REQUEST"
              '(earlier we called the asynch lineMakeCall() function from the
              'MakeCallAsynch method and stored the requester in m_RequestingCall
              'now we can see that this is the asynchronous reply to that call)
              'don't need ID anymore - it matched
50            m_RequestingCall = 0

              'if it was an error make sure the line is closed
60            If dwParam2 <> 0 Then
70              Call CloseLine
80              dbgTAPI "ERROR: cTapiLine.LineProcHandler Call CloseLine Param2 " & GetLineErrString(dwParam2)
90            End If
              'send result to parent via event
100           RaiseEvent MakeCallResult(dwParam2)

110         ElseIf dwParam1 = m_DroppingCall Then
120           dbgTAPI "LINE_REPLY-CALL DROP"
              'asynch reply to lineDrop() call
              'Call CloseLine
130           m_DroppingCall = 0
140           RaiseEvent DropCallResult(dwParam2)

            Else
              'dbgTAPI "LineProcHandler"

150         End If

160       Case LINE_CALLSTATE  ' & h2
170         Select Case dwParam1
              Case LINECALLSTATE_DISCONNECTED
                'call has disconnected
                ' announced elsewhere: dbgtapi  "LINECALLSTATE_DISCONNECTED"
180             mLastDisconnectReason = dwParam2
190             RaiseEvent Disconnected(dwParam2)
                'Call DropCallAsynch

200           Case LINECALLSTATE_IDLE
                'no call exists - the line is idle
210             dbgTAPI "LINECALLSTATE_IDLE"
220             If mhCall <> 0 Then
230               Call lineDeallocateCall(mhCall)
240               mhCall = 0
250             End If
260             RaiseEvent Idle

270           Case LINECALLSTATE_CONNECTED
280             RaiseEvent Connected

290           Case LINECALLSTATE_BUSY
                'NOT found in modems.
300             dbgTAPI "LINECALLSTATE_BUSY"


310           Case LINECALLSTATE_DIALTONE
                'NOT found in modems.
320             dbgTAPI "LINECALLSTATE_DIALTONE"

330           Case LINECALLSTATE_RINGBACK
                'NOT found in modems.
340             dbgTAPI "LINECALLSTATE_RINGBACK"
350             RaiseEvent RingBack

360           Case LINECALLSTATE_DIALING
370             RaiseEvent Dialing

380           Case LINECALLSTATE_PROCEEDING
390             RaiseEvent Proceeding

400           Case LINECALLSTATE_SPECIALINFO
                'possible network error
410             dbgTAPI "LINECALLSTATE_SPECIALINFO"
420             RaiseEvent SpecialInfo(Hex(dwParam1))
430           Case Else
440             dbgTAPI UCase("Unknown CallSTATE: " & Hex(dwParam1))
450             RaiseEvent Unknown(Hex(dwParam1))
460         End Select
470       Case LINE_MONITORDIGITS  ' & h9
480         RaiseEvent GotDigit(dwParam1)
490         dbgTAPI UCase("LINE_MONITORDIGITS " & Hex(dwParam1))
500       Case Else


510     End Select

LineProcHandler_Resume:

520     On Error GoTo 0
530     Exit Sub

LineProcHandler_Error:

540     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at CTAPILine.LineProcHandler." & Erl
550     Resume LineProcHandler_Resume

End Sub

Public Function GetReasonDisconnected(ByVal LineCallState As Long) As String
  Select Case LineCallState
    Case LINEDISCONNECTMODE_NORMAL
      GetReasonDisconnected = "Normal Disconnect"
    Case LINEDISCONNECTMODE_UNKNOWN
      GetReasonDisconnected = "Disconnected Unknown reason"
    Case LINEDISCONNECTMODE_REJECT
      GetReasonDisconnected = "Remote Party rejected call"
    Case LINEDISCONNECTMODE_PICKUP
      GetReasonDisconnected = "Disconnected Call was picked up on another phone"
    Case LINEDISCONNECTMODE_FORWARDED
      GetReasonDisconnected = "Disconnected Forwarded"
    Case LINEDISCONNECTMODE_BUSY
      GetReasonDisconnected = "Busy"
    Case LINEDISCONNECTMODE_NOANSWER
      GetReasonDisconnected = "No Answer"
    Case LINEDISCONNECTMODE_BADADDRESS
      GetReasonDisconnected = "Disconnected Bad Address"
    Case LINEDISCONNECTMODE_UNREACHABLE
      GetReasonDisconnected = "Disconnected Unreachable"
    Case LINEDISCONNECTMODE_CONGESTION
      GetReasonDisconnected = "Disconnected Congestion"
    Case LINEDISCONNECTMODE_INCOMPATIBLE
      GetReasonDisconnected = "Disconnected Incompatible"
    Case LINEDISCONNECTMODE_UNAVAIL
      GetReasonDisconnected = "Disconnected Unavailable"
    Case LINEDISCONNECTMODE_NODIALTONE
      GetReasonDisconnected = "No Dial Tone"
    Case Else
      GetReasonDisconnected = "LINECALLSTATE Bad Reason"
  End Select

End Function


Public Function GetLineErrString(lparam As Long) As String
'Returns a String description of a TAPI Line Error code
  Dim msg As String

  Select Case lparam
    Case LINEERR_ALLOCATED  '( = &H80000001)
      msg = "Allocated"
    Case LINEERR_BADDEVICEID  '(= &H80000002)
      msg = "Bad Device ID"
    Case LINEERR_BEARERMODEUNAVAIL  '(= &H80000003)
      msg = "Bearer Mode Unavail"
    Case LINEERR_CALLUNAVAIL  '(= &H80000005)
      msg = "Call UnAvail"
    Case LINEERR_COMPLETIONOVERRUN  '(= &H80000006
      msg = "Completion Overrun"
    Case LINEERR_CONFERENCEFULL  '(= &H80000007
      msg = "Conference Full"
    Case LINEERR_DIALBILLING  '(= &H80000008
      msg = "Dial Billing"
    Case LINEERR_DIALDIALTONE  '(= &H80000009
      msg = "Dial Dialtone"
    Case LINEERR_DIALPROMPT  '(= &H8000000A
      msg = "Dial Prompt"
    Case LINEERR_DIALQUIET  '(= &H8000000B
      msg = "Dial Quiet"
    Case LINEERR_INCOMPATIBLEAPIVERSION  '(= &H8000000C
      msg = "Incompatible API Version"
    Case LINEERR_INCOMPATIBLEEXTVERSION  '(= &H8000000D
      msg = "Incompatible Ext Version"
    Case LINEERR_INIFILECORRUPT  '(= &H8000000E
      msg = "Ini File Corrupt"
    Case LINEERR_INUSE  '(= &H8000000F
      msg = "In Use"
    Case LINEERR_INVALADDRESS  '(= &H80000010
      msg = "Invalid Address"
    Case LINEERR_INVALADDRESSID  '(= &H80000011
      msg = "Invalid Address ID"
    Case LINEERR_INVALADDRESSMODE  '(= &H80000012
      msg = "Invalid Address Mode"
    Case LINEERR_INVALADDRESSSTATE  '(= &H80000013
      msg = "Invalid Address State"
    Case LINEERR_INVALAPPHANDLE  '(= &H80000014
      msg = "Invalid App Handle"
    Case LINEERR_INVALAPPNAME  '(= &H80000015
      msg = "Invalid App Name"
    Case LINEERR_INVALBEARERMODE  '(= &H80000016
      msg = "Invalid Bearer Mode"
    Case LINEERR_INVALCALLCOMPLMODE  '(= &H80000017
      msg = "Invalid Call Completion Mode"
    Case LINEERR_INVALCALLHANDLE  '(= &H80000018
      msg = "Invalid Call Handle"
    Case LINEERR_INVALCALLPARAMS  '(= &H80000019
      msg = "Invalid Call Params"
    Case LINEERR_INVALCALLPRIVILEGE  '(= &H8000001A
      msg = "Invalid Call Privilege"
    Case LINEERR_INVALCALLSELECT  '(= &H8000001B
      msg = "Invalid Call Select"
    Case LINEERR_INVALCALLSTATE  '(= &H8000001C
      msg = "Invalid Call State"
    Case LINEERR_INVALCALLSTATELIST  '(= &H8000001D
      msg = "Invalid Call State List"
    Case LINEERR_INVALCARD  '(= &H8000001E
      msg = "Invalid Card"
    Case LINEERR_INVALCOMPLETIONID  '(= &H8000001F
      msg = "Invalid Completion ID"
    Case LINEERR_INVALCONFCALLHANDLE  '(= &H80000020
      msg = "Invalid Conf Call Handle"
    Case LINEERR_INVALCONSULTCALLHANDLE  '(= &H80000021
      msg = "Invalid Consult Call Handle"
    Case LINEERR_INVALCOUNTRYCODE  '(= &H80000022
      msg = "Invalid Country Code"
    Case LINEERR_INVALDEVICECLASS  '(= &H80000023
      msg = "Invalid Device Class"
    Case LINEERR_INVALDEVICEHANDLE  '(= &H80000024
      msg = "Invalid Device Handle"
    Case LINEERR_INVALDIGITLIST  '(= &H80000026
      msg = "Invalid Digit List"
    Case LINEERR_INVALDIGITMODE  '(= &H80000027
      msg = "Invalid Digit Mode"
    Case LINEERR_INVALDIGITS  '(= &H80000028
      msg = "Invalid Digits"
    Case LINEERR_INVALEXTVERSION  '(= &H80000029
      msg = "Invalid Ext Version"
    Case LINEERR_INVALGROUPID  '(= &H8000002A
      msg = "Invalid Group ID"
    Case LINEERR_INVALLINEHANDLE  '(= &H8000002B
      msg = "Invalid Line Handle"
    Case LINEERR_INVALLINESTATE  '(= &H8000002C
      msg = "Invalid Line State"
    Case LINEERR_INVALLOCATION  '(= &H8000002D
      msg = "Invalid Location"
    Case LINEERR_INVALMEDIALIST  '(= &H8000002E
      msg = "Invalid Media List"
    Case LINEERR_INVALMEDIAMODE  '(= &H8000002F
      msg = "Invalid Media Mode"
    Case LINEERR_INVALMESSAGEID  '(= &H80000030
      msg = "Invalid Message ID"
    Case LINEERR_INVALPARAM  '(= &H80000032
      msg = "Invalid Param"
    Case LINEERR_INVALPARKID  '(= &H80000033
      msg = "Invalid Park ID"
    Case LINEERR_INVALPARKMODE  '(= &H80000034
      msg = "Invalid Park Mode"
    Case LINEERR_INVALPOINTER  '(= &H80000035
      msg = "Invalid Pointer"
    Case LINEERR_INVALPRIVSELECT  '(= &H80000036
      msg = "Invalid Priv Select"
    Case LINEERR_INVALRATE  '(= &H80000037
      msg = "Invalid Rate"
    Case LINEERR_INVALREQUESTMODE  '(= &H80000038
      msg = "Invalid Request Mode"
    Case LINEERR_INVALTERMINALID  '(= &H80000039
      msg = "Invalid Terminal ID"
    Case LINEERR_INVALTERMINALMODE  '(= &H8000003A
      msg = "Invalid Terminal Mode"
    Case LINEERR_INVALTIMEOUT  '(= &H8000003B
      msg = "Invalid Time Out"
    Case LINEERR_INVALTONE  '(= &H8000003C
      msg = "Invalid Tone"
    Case LINEERR_INVALTONELIST  '(= &H8000003D
      msg = "Invalid Tone List"
    Case LINEERR_INVALTONEMODE  '(= &H8000003E
      msg = "Invalid Tone Mode"
    Case LINEERR_INVALTRANSFERMODE  '(= &H8000003F
      msg = "Invalid Transfer Mode"
    Case LINEERR_LINEMAPPERFAILED  '(= &H80000040
      msg = "Line Mapper Failed"
    Case LINEERR_NOCONFERENCE  '(= &H80000041
      msg = "No Conference"
    Case LINEERR_NODEVICE  '(= &H80000042
      msg = "No Device"
    Case LINEERR_NODRIVER  '(= &H80000043
      msg = "No Driver"
    Case LINEERR_NOMEM  '(= &H80000044
      msg = "No Memory"
    Case LINEERR_NOREQUEST  '(= &H80000045
      msg = "No Request"
    Case LINEERR_NOTOWNER  '(= &H80000046
      msg = "Not Owner"
    Case LINEERR_NOTREGISTERED  '(= &H80000047
      msg = "Not Registered"
    Case LINEERR_OPERATIONFAILED  '(= &H80000048
      msg = "Operation Failed"
    Case LINEERR_OPERATIONUNAVAIL  '(= &H80000049
      msg = "Operation Unavailable"
    Case LINEERR_RATEUNAVAIL  '(= &H8000004A
      msg = "Rate Unavailable"
    Case LINEERR_RESOURCEUNAVAIL  '(= &H8000004B
      msg = "Resource Unavailable"
    Case LINEERR_REQUESTOVERRUN  '(= &H8000004C
      msg = "Request Overrun"
    Case LINEERR_STRUCTURETOOSMALL  '(= &H8000004D
      msg = "Structure Too Small"
    Case LINEERR_TARGETNOTFOUND  '(= &H8000004E
      msg = "Target Not found"
    Case LINEERR_TARGETSELF  '(= &H8000004F
      msg = "Target Self"
    Case LINEERR_UNINITIALIZED  '(= &H80000050
      msg = "Uninitialized"
    Case LINEERR_USERUSERINFOTOOBIG  '(= &H80000051
      msg = "UserUser Info Too Big"
    Case LINEERR_REINIT  '(= &H80000052
      msg = "Re-init"
    Case LINEERR_ADDRESSBLOCKED  '(= &H80000053
      msg = "Address Blocked"
    Case LINEERR_BILLINGREJECTED  '(= &H80000054
      msg = "Billing Rejected"
    Case LINEERR_INVALFEATURE  '(= &H80000055
      msg = "Invalid Feature"
    Case LINEERR_NOMULTIPLEINSTANCE  '(= &H80000056
      msg = "No Multiple Instance"
    Case Else
      msg = "Unknown Error"  ' undefined
  End Select

  GetLineErrString = msg
End Function

