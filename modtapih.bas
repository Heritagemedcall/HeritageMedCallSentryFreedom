Attribute VB_Name = "modTapiH"
Option Explicit


'http://msdn.microsoft.com/en-us/library/ms735996(VS.85).aspx Line monitor digits

'LineMonitorDigits
'Returns zero if the request succeeds or a negative error number if an error occurs.
'Possible error return values are:
'LINEERR_INVALCALLHANDLE
'LINEERR_OPERATIONUNAVAIL
'LINEERR_INVALCALLSTATE
'LINEERR_OPERATIONFAILED
'LINEERR_INVALDIGITMODE
'LINEERR_RESOURCEUNAVAIL
'LINEERR_NOMEM
'LINEERR_UNINITIALIZED


'LINEDIGITMODE_DTMF = 2
'Uses DTMF tones to signal digits. Valid digits are 0 through 9, '*', '#', 'A', 'B', 'C', and 'D'.
'LINEDIGITMODE_DTMFEND = 4
'Uses DTMF tones to signal digits and detect the down edges. Valid digits are 0 through 9, '*', '#', 'A', 'B', 'C', and 'D'.
'LINEDIGITMODE_PULSE = 1
'Uses rotary pulse sequences to signal digits. Valid digits are 0 through 9.
'LINEDIGITMODE_NONE = 0
' stops monitoring




Public Enum TapiEvent

  LINE_ADDRESSSTATE = 0
  LINE_CALLINFO
  LINE_CALLSTATE
  LINE_CLOSE
  LINE_DEVSPECIFIC
  LINE_DEVSPECIFICFEATURE
  LINE_GATHERDIGITS
  LINE_GENERATE
  LINE_LINEDEVSTATE
  LINE_MONITORDIGITS
  LINE_MONITORMEDIA
  LINE_MONITORTONE
  LINE_REPLY
  LINE_REQUEST
  PHONE_BUTTON
  PHONE_CLOSE
  PHONE_DEVSPECIFIC
  PHONE_REPLY
  PHONE_STATE
  LINE_CREATE  ' TAPI v1.4
  PHONE_CREATE  ' TAPI v1.4

End Enum

Global Const LINECALLSELECT_LINE = &H1
Global Const LINECALLSELECT_ADDRESS = &H2
Global Const LINECALLSELECT_CALL = &H4
Global Const LINECALLSELECT_DEVICEID = &H8

Global Const LINEDIGITMODE_NONE = 0
Global Const LINEDIGITMODE_PULSE = 1
Global Const LINEDIGITMODE_DTMF = 2
Global Const LINEDIGITMODE_DTMFEND = 4

'Global  Const HIGHTAPIVERSION = &H20001 'Also available as an upgrade to Win9x
'Global  Const WIN95TAPIVERSION = &H10004
'Only support 1.4, with lineInitialize that is the best we can do anyhow.
Global Const TAPIVERSION = &H10004

Global Const LINECALLPRIVILEGE_NONE = &H1
Global Const LINECALLPRIVILEGE_MONITOR = &H2
Global Const LINECALLPRIVILEGE_OWNER = &H4

 Const LINECALLINFOSTATE_CALLERID = 32768

'LINECALLPARTYID_ Constants
Global Const LINECALLPARTYID_BLOCKED = &H1
Global Const LINECALLPARTYID_OUTOFAREA = &H2
Global Const LINECALLPARTYID_NAME = &H4
Global Const LINECALLPARTYID_ADDRESS = &H8
Global Const LINECALLPARTYID_PARTIAL = &H10
Global Const LINECALLPARTYID_UNKNOWN = &H20
Global Const LINECALLPARTYID_UNAVAIL = &H40

Global Const LINEMEDIAMODE_UNKNOWN = &H2
Global Const LINEMEDIAMODE_INTERACTIVEVOICE = &H4
Global Const LINEMEDIAMODE_AUTOMATEDVOICE = &H8
Global Const LINEMEDIAMODE_DATAMODEM = &H10
Global Const LINEMEDIAMODE_G3FAX = &H20
Global Const LINEMEDIAMODE_TDD = &H40
Global Const LINEMEDIAMODE_G4FAX = &H80
Global Const LINEMEDIAMODE_DIGITALDATA = &H100
Global Const LINEMEDIAMODE_TELETEX = &H200
Global Const LINEMEDIAMODE_VIDEOTEX = &H400
Global Const LINEMEDIAMODE_TELEX = &H800
Global Const LINEMEDIAMODE_MIXED = &H1000
Global Const LINEMEDIAMODE_ADSI = &H2000
Global Const LINEMEDIAMODE_VOICEVIEW = &H4000   ' TAPI v1.4
Global Const LINEMEDIAMODE_VIDEO = &H8000   ' TAPI v2.1

'LINECALLSTATE Constants
Global Const LINECALLSTATE_IDLE = &H1
Global Const LINECALLSTATE_OFFERING = &H2
Global Const LINECALLSTATE_ACCEPTED = &H4
Global Const LINECALLSTATE_DIALTONE = &H8
Global Const LINECALLSTATE_DIALING = &H10
Global Const LINECALLSTATE_RINGBACK = &H20
Global Const LINECALLSTATE_BUSY = &H40
Global Const LINECALLSTATE_SPECIALINFO = &H80
Global Const LINECALLSTATE_CONNECTED = &H100
Global Const LINECALLSTATE_PROCEEDING = &H200
Global Const LINECALLSTATE_ONHOLD = &H400
Global Const LINECALLSTATE_CONFERENCED = &H800
Global Const LINECALLSTATE_ONHOLDPENDCONF = &H1000
Global Const LINECALLSTATE_ONHOLDPENDTRANSFER = &H2000
Global Const LINECALLSTATE_DISCONNECTED = &H4000
Global Const LINECALLSTATE_UNKNOWN = &H8000

'LINEDEVSTATE Constants
Global Const LINEDEVSTATE_OTHER = &H1
Global Const LINEDEVSTATE_RINGING = &H2
Global Const LINEDEVSTATE_CONNECTED = &H4
Global Const LINEDEVSTATE_DISCONNECTED = &H8
Global Const LINEDEVSTATE_MSGWAITON = &H10
Global Const LINEDEVSTATE_MSGWAITOFF = &H20
Global Const LINEDEVSTATE_INSERVICE = &H40
Global Const LINEDEVSTATE_OUTOFSERVICE = &H80
Global Const LINEDEVSTATE_MAINTENANCE = &H100
Global Const LINEDEVSTATE_OPEN = &H200
Global Const LINEDEVSTATE_CLOSE = &H400
Global Const LINEDEVSTATE_NUMCALLS = &H800
Global Const LINEDEVSTATE_NUMCOMPLETIONS = &H1000
Global Const LINEDEVSTATE_TERMINALS = &H2000
Global Const LINEDEVSTATE_ROAMMODE = &H4000
Global Const LINEDEVSTATE_BATTERY = &H8000
Global Const LINEDEVSTATE_SIGNAL = &H10000
Global Const LINEDEVSTATE_DEVSPECIFIC = &H20000
Global Const LINEDEVSTATE_REINIT = &H40000
Global Const LINEDEVSTATE_LOCK = &H80000
Global Const LINEDEVSTATE_CAPSCHANGE = &H100000   ' TAPI v1.4
Global Const LINEDEVSTATE_CONFIGCHANGE = &H200000   ' TAPI v1.4
Global Const LINEDEVSTATE_TRANSLATECHANGE = &H400000   ' TAPI v1.4
Global Const LINEDEVSTATE_COMPLCANCEL = &H800000   ' TAPI v1.4
Global Const LINEDEVSTATE_REMOVED = &H1000000   ' TAPI v1.4


Type LINEDIALPARAMS
  dwDialPause As Long
  dwDialSpeed As Long
  dwDigitDuration As Long
  dwWaitForDialtone As Long
End Type

Type lineCallInfo

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

  ''#if (TAPI_CURRENT_VERSION >= 0x00020000)
  '    dwCallTreatment As Long                                ' TAPI v2.0
  '    dwCallDataSize As Long                                 ' TAPI v2.0
  '    dwCallDataOffset As Long                               ' TAPI v2.0
  '    dwSendingFlowspecSize As Long                          ' TAPI v2.0
  '    dwSendingFlowspecOffset As Long                        ' TAPI v2.0
  '    dwReceivingFlowspecSize As Long                        ' TAPI v2.0
  '    dwReceivingFlowspecOffset As Long                      ' TAPI v2.0
  ''#End If
  bBytes(2000) As Byte  'HACK Added to TAPI structure for callinfo data.

End Type

Type LINEEXTENSIONID
  dwExtensionID0 As Long
  dwExtensionID1 As Long
  dwExtensionID2 As Long
  dwExtensionID3 As Long
End Type

Type LineDevCaps
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
  dwLineFeatures As Long  ' TAPI v1.4
  ''#if (TAPI_CURRENT_VERSION >= 0x00020000)
  dwSettableDevStatus As Long  ' TAPI v2.0
  dwDeviceClassesSize As Long  ' TAPI v2.0
  dwDeviceClassesOffset As Long  ' TAPI v2.0
  ''#End If
  bBytes(2000) As Byte
End Type

Type varString
  dwTotalSize As Long
  dwNeededSize As Long
  dwUsedSize As Long
  dwStringFormat As Long
  dwStringSize As Long
  dwStringOffset As Long
  bBytes(2000) As Byte  'Added to TAPI structure for lineGetID data.
End Type


Global Const TAPI_SUCCESS As Long = 0&   'declared for convenience
Global Const LINEERR_ALLOCATED As Long = &H80000001
Global Const LINEERR_BADDEVICEID As Long = &H80000002
Global Const LINEERR_BEARERMODEUNAVAIL As Long = &H80000003
Global Const LINEERR_CALLUNAVAIL As Long = &H80000005
Global Const LINEERR_COMPLETIONOVERRUN As Long = &H80000006
Global Const LINEERR_CONFERENCEFULL As Long = &H80000007
Global Const LINEERR_DIALBILLING As Long = &H80000008
Global Const LINEERR_DIALDIALTONE As Long = &H80000009
Global Const LINEERR_DIALPROMPT As Long = &H8000000A
Global Const LINEERR_DIALQUIET As Long = &H8000000B
Global Const LINEERR_INCOMPATIBLEAPIVERSION As Long = &H8000000C
Global Const LINEERR_INCOMPATIBLEEXTVERSION As Long = &H8000000D
Global Const LINEERR_INIFILECORRUPT As Long = &H8000000E
Global Const LINEERR_INUSE As Long = &H8000000F
Global Const LINEERR_INVALADDRESS As Long = &H80000010
Global Const LINEERR_INVALADDRESSID As Long = &H80000011
Global Const LINEERR_INVALADDRESSMODE As Long = &H80000012
Global Const LINEERR_INVALADDRESSSTATE As Long = &H80000013
Global Const LINEERR_INVALAPPHANDLE As Long = &H80000014
Global Const LINEERR_INVALAPPNAME As Long = &H80000015
Global Const LINEERR_INVALBEARERMODE As Long = &H80000016
Global Const LINEERR_INVALCALLCOMPLMODE As Long = &H80000017
Global Const LINEERR_INVALCALLHANDLE As Long = &H80000018
Global Const LINEERR_INVALCALLPARAMS As Long = &H80000019
Global Const LINEERR_INVALCALLPRIVILEGE As Long = &H8000001A
Global Const LINEERR_INVALCALLSELECT As Long = &H8000001B
Global Const LINEERR_INVALCALLSTATE As Long = &H8000001C
Global Const LINEERR_INVALCALLSTATELIST As Long = &H8000001D
Global Const LINEERR_INVALCARD As Long = &H8000001E
Global Const LINEERR_INVALCOMPLETIONID As Long = &H8000001F
Global Const LINEERR_INVALCONFCALLHANDLE As Long = &H80000020
Global Const LINEERR_INVALCONSULTCALLHANDLE As Long = &H80000021
Global Const LINEERR_INVALCOUNTRYCODE As Long = &H80000022
Global Const LINEERR_INVALDEVICECLASS As Long = &H80000023
Global Const LINEERR_INVALDEVICEHANDLE As Long = &H80000024
Global Const LINEERR_INVALDIALPARAMS = &H80000025   ' from answ mach
Global Const LINEERR_INVALDIGITLIST As Long = &H80000026
Global Const LINEERR_INVALDIGITMODE As Long = &H80000027
Global Const LINEERR_INVALDIGITS As Long = &H80000028
Global Const LINEERR_INVALEXTVERSION As Long = &H80000029
Global Const LINEERR_INVALGROUPID As Long = &H8000002A
Global Const LINEERR_INVALLINEHANDLE As Long = &H8000002B
Global Const LINEERR_INVALLINESTATE As Long = &H8000002C
Global Const LINEERR_INVALLOCATION As Long = &H8000002D
Global Const LINEERR_INVALMEDIALIST As Long = &H8000002E
Global Const LINEERR_INVALMEDIAMODE As Long = &H8000002F
Global Const LINEERR_INVALMESSAGEID As Long = &H80000030
Global Const LINEERR_INVALPARAM As Long = &H80000032
Global Const LINEERR_INVALPARKID As Long = &H80000033
Global Const LINEERR_INVALPARKMODE As Long = &H80000034
Global Const LINEERR_INVALPOINTER As Long = &H80000035
Global Const LINEERR_INVALPRIVSELECT As Long = &H80000036
Global Const LINEERR_INVALRATE As Long = &H80000037
Global Const LINEERR_INVALREQUESTMODE As Long = &H80000038
Global Const LINEERR_INVALTERMINALID As Long = &H80000039
Global Const LINEERR_INVALTERMINALMODE As Long = &H8000003A
Global Const LINEERR_INVALTIMEOUT As Long = &H8000003B
Global Const LINEERR_INVALTONE As Long = &H8000003C
Global Const LINEERR_INVALTONELIST As Long = &H8000003D
Global Const LINEERR_INVALTONEMODE As Long = &H8000003E
Global Const LINEERR_INVALTRANSFERMODE As Long = &H8000003F
Global Const LINEERR_LINEMAPPERFAILED As Long = &H80000040
Global Const LINEERR_NOCONFERENCE As Long = &H80000041
Global Const LINEERR_NODEVICE As Long = &H80000042
Global Const LINEERR_NODRIVER As Long = &H80000043
Global Const LINEERR_NOMEM As Long = &H80000044
Global Const LINEERR_NOREQUEST As Long = &H80000045
Global Const LINEERR_NOTOWNER As Long = &H80000046
Global Const LINEERR_NOTREGISTERED As Long = &H80000047
Global Const LINEERR_OPERATIONFAILED As Long = &H80000048
Global Const LINEERR_OPERATIONUNAVAIL As Long = &H80000049
Global Const LINEERR_RATEUNAVAIL As Long = &H8000004A
Global Const LINEERR_RESOURCEUNAVAIL As Long = &H8000004B
Global Const LINEERR_REQUESTOVERRUN As Long = &H8000004C
Global Const LINEERR_STRUCTURETOOSMALL As Long = &H8000004D
Global Const LINEERR_TARGETNOTFOUND As Long = &H8000004E
Global Const LINEERR_TARGETSELF As Long = &H8000004F
Global Const LINEERR_UNINITIALIZED As Long = &H80000050
Global Const LINEERR_USERUSERINFOTOOBIG As Long = &H80000051
Global Const LINEERR_REINIT As Long = &H80000052
Global Const LINEERR_ADDRESSBLOCKED As Long = &H80000053
Global Const LINEERR_BILLINGREJECTED As Long = &H80000054
Global Const LINEERR_INVALFEATURE As Long = &H80000055
Global Const LINEERR_NOMULTIPLEINSTANCE As Long = &H80000056

Global Const LINEFEATURE_DEVSPECIFIC As Long = &H1&
Global Const LINEFEATURE_DEVSPECIFICFEAT As Long = &H2&
Global Const LINEFEATURE_FORWARD As Long = &H4&
Global Const LINEFEATURE_MAKECALL As Long = &H8&
Global Const LINEFEATURE_SETMEDIACONTROL As Long = &H10&
Global Const LINEFEATURE_SETTERMINAL As Long = &H20&

Global Const LINECALLFEATURE_ACCEPT As Long = &H1&
Global Const LINECALLFEATURE_ADDTOCONF As Long = &H2&
Global Const LINECALLFEATURE_ANSWER As Long = &H4&
Global Const LINECALLFEATURE_BLINDTRANSFER As Long = &H8&
Global Const LINECALLFEATURE_COMPLETECALL As Long = &H10&
Global Const LINECALLFEATURE_COMPLETETRANSF As Long = &H20&
Global Const LINECALLFEATURE_DIAL As Long = &H40&
Global Const LINECALLFEATURE_DROP As Long = &H80&
Global Const LINECALLFEATURE_GATHERDIGITS As Long = &H100&
Global Const LINECALLFEATURE_GENERATEDIGITS As Long = &H200&
Global Const LINECALLFEATURE_GENERATETONE As Long = &H400&
Global Const LINECALLFEATURE_HOLD As Long = &H800&
Global Const LINECALLFEATURE_MONITORDIGITS As Long = &H1000&
Global Const LINECALLFEATURE_MONITORMEDIA As Long = &H2000&
Global Const LINECALLFEATURE_MONITORTONES As Long = &H4000&
Global Const LINECALLFEATURE_PARK As Long = &H8000&
Global Const LINECALLFEATURE_PREPAREADDCONF As Long = &H10000
Global Const LINECALLFEATURE_REDIRECT As Long = &H20000
Global Const LINECALLFEATURE_REMOVEFROMCONF As Long = &H40000
Global Const LINECALLFEATURE_SECURECALL As Long = &H80000
Global Const LINECALLFEATURE_SENDUSERUSER As Long = &H100000
Global Const LINECALLFEATURE_SETCALLPARAMS As Long = &H200000
Global Const LINECALLFEATURE_SETMEDIACONTROL As Long = &H400000
Global Const LINECALLFEATURE_SETTERMINAL As Long = &H800000
Global Const LINECALLFEATURE_SETUPCONF As Long = &H1000000
Global Const LINECALLFEATURE_SETUPTRANSFER As Long = &H2000000
Global Const LINECALLFEATURE_SWAPHOLD As Long = &H4000000
Global Const LINECALLFEATURE_UNHOLD As Long = &H8000000



'#if (TAPI_CURRENT_VERSION >0x00020000)
Global Const LINECALLTREATMENT_SILENCE                 As Long = &H1&  '// TAPI v2.0
Global Const LINECALLTREATMENT_RINGBACK                As Long = &H2&  '// TAPI v2.0
Global Const LINECALLTREATMENT_BUSY                    As Long = &H3&  '// TAPI v2.0
Global Const LINECALLTREATMENT_MUSIC                   As Long = &H4&  '// TAPI v2.0
'#End If

Global Const TAPI_DTMF_0                               As Long = &H30
Global Const TAPI_DTMF_1                               As Long = &H31
Global Const TAPI_DTMF_2                               As Long = &H32
Global Const TAPI_DTMF_3                               As Long = &H33
Global Const TAPI_DTMF_4                               As Long = &H34
Global Const TAPI_DTMF_5                               As Long = &H35
Global Const TAPI_DTMF_6                               As Long = &H36
Global Const TAPI_DTMF_7                               As Long = &H37
Global Const TAPI_DTMF_8                               As Long = &H38
Global Const TAPI_DTMF_9                               As Long = &H39
Global Const TAPI_DTMF_STAR                            As Long = &H2A
Global Const TAPI_DTMF_POUND                           As Long = &H23




'[2784] 9  39  2  F6120AE  9?
'[2784] 9  38  2  F612E91
'[2784] 9  37  2  F613DFB
'[2784] 9  38  2  F61474A
'[2784] 9  39  2  F6149C1
'[2784] 9  36  2  F6159CB
'[2784] 9  23  2  F616826 3?

