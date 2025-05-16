VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBerkshire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Temperature  As Double
Public Tickles      As Long


' ----------------------------------------------------
' This module for use with WDog_Univrsl.dll
' Created Feb 25/2010
' www.berkprod.com
' ----------------------------------------------------

' ----------------------------------------------------
' DLL Return codes
' ----------------------------------------------------
Private Enum WDogFlags
  WD_NO_DLL = -1
  WD_OK = 0
  WD_DEVICE_NOT_FOUND
  WD_DEVICE_NOT_OPENED
  WD_INVALID_HANDLE
  WD_FAIL_COMM_INIT
  WD_IO_ERROR
  WD_IO_READ_FAILED
  WD_IO_WRITE_FAILED
  WD_RX_CHECKSUM_ERROR                 ' Not Used on PCI
  WD_TX_CHECKSUM_ERROR                 ' Not Used on PCI
  WD_INVALID_COMMAND                   'board firmware Rev may be to old for new command?

  WD_OPEN_EX_ERROR
  WD_OPEN_EX_COMM_ERR                  ' COM port specified was greater than MAX_COMM_PORTS
  WD_COMM_VAL_ERR
  WD_BD_TYPE_UNKNOWN                   ' called a function where the type was unknown
  ' could be because board was not yet opened.
  WD_USB_DISABLE_FAIL = &H100           ' USB Specific
  WD_USB_ENABLE_FAIL
  WD_USB_FIND_COMM_ERROR               ' found a USB but errored on initial comm

  WD_PCI_DISABLE_FAIL = &H140       ' PCI Specific
  WD_PCI_ENABLE_FAIL
  WD_PCI_FIND_COMM_ERROR               ' found a PCI but errored on initial comm

  WD_EUSB_DISABLE_FAIL = &H180          ' ETH-USB Specific
  WD_EUSB_ENABLE_FAIL
  WD_EUSB_FIND_COMM_ERROR              ' found a PCI but errored on initial comm
  WD_FTD2XX_DLL_NOT_FOUND
  WD_FTD2XX_DLL_VER_ERR
  WD_INVALID_INFO_STRUCTURE
  WD_INVALID_IP_ADDRESS

  WD_SERIAL_NO_COMM_PORT = &H200       ' Could not find a comm port in the PC
  WD_SERIAL_NO_DEVICE                  ' Check all comm ports - no device found
  WD_SERIAL_FIND_COMM_ERROR            ' found a Serial but errored on initial comm

  ' Special Flag/Status for serial WDogs - if possble it will be used as
  ' a flag - otherwise it will be a WD_STATUS
  WD_SERIAL_DATA_CONVERTED = &H80000000        ' data saved - but not an exact converion


  ' Special for firmware - you can use upper byte to distinguish that it is firmware
  WD_WDOG_FIRMWARE_ERR_FLAG = &HFF000000        ' upper byte = 0xff indicates firmware error
  ' for any type WDog
  WD_USB_FIRMWARE_ERROR                 ' firmware rev too low to accept command
  WD_PCI_FIRMWARE_ERROR                 ' firmware rev too low to accept command
  WD_EUSB_FIRMWARE_ERROR            ' firmware rev too low to accept command
  WD_SERIAL_FIRMWARE_ERROR              ' firmware rev too low to accept command

  WD_OTHER_ERROR = -1
End Enum

' ----------------------------------------------------
' DLL Flags
' ----------------------------------------------------
' Status flags
Private Const WD_STAT_PCI_ENTP As Long = &H200000        ' PCI Board Only - Bit in Port #2
Private Const WD_STAT_PCI_WDIS As Long = &H100000        ' PCI Board Only - Bit in Port #2
Private Const WD_STAT_PCI_TTRP As Long = &H40000         ' PCI Board Only - Bit in Port #1
Private Const WD_STAT_PCI_WTRP As Long = &H10000         ' PCI Board Only - Bit in Port #1

Private Const WD_STAT_ETH_IP_SET As Long = &H40          ' Ethernet has a valid IP
Private Const WD_STAT_ETH_ENABLED As Long = &H20         ' Ehternet enabled
Private Const WD_STAT_CMD_DISABLED As Long = &H8         ' WDog has been disabled by command
Private Const WD_STAT_POD_DSW_DELAY As Long = &H4        ' WDog is armed - SW set to wait for first "tickle"
Private Const WD_STAT_POD_ACTIVE As Long = &H2           ' WDog is still in 2.5 minute (or user) delay
Private Const WD_STAT_ACTIVE_ARMED As Long = &H1         ' WDog is armed and done with POD time


' Diagnostic flags
Private Const WD_DIAG_ARP_ERROR As Long = &H40            'IP address conflict
Private Const WD_DIAG_NV_WRITE_FAIL As Long = &H20        ' data not written to NV memory
Private Const WD_DIAG_MAC_INVALID As Long = &H10          ' MAC Address corrupted
Private Const WD_DIAG_NV_CORRUPTED As Long = &H8          ' Non-Voaltile Memory Corrupted
Private Const WD_DIAG_ETHER_OK As Long = &H4              ' Ethernet Chip OK
Private Const WD_DIAG_NVMEM_OK As Long = &H2              ' Non-Volatile Memory OK
Private Const WD_DIAG_TEMP_OK As Long = &H1               ' Temp Sensor IC OK

'------------------------------
' DIP Switch Bit Defines
' Common to all Boards
Private Const WD_DSW_OVERTEMP_RESET As Long = &H10
Private Const WD_DSW_POWER_ON_DELAY As Long = &H8        ' Wait for first "tickle" after POD time
Private Const WD_DSW_WDOG_TIME4 As Long = &H4
Private Const WD_DSW_WDOG_TIME2 As Long = &H2
Private Const WD_DSW_WDOG_TIME1 As Long = &H1
' USB & ETH-USB
Private Const WD_DSW_BUZZER_DISABLE As Long = &H20
' ETH-USB Switches
Private Const WD_DSW_ETH_OPTION1 As Long = &H80           ' Ethernet option switches
Private Const WD_DSW_ETH_OPTION2 As Long = &H40
' USB Switches
Private Const WD_DSW_PASSIVE_MODE As Long = &H80          ' SW1 Passive Mode
Private Const WD_DSW_ACTIVE_MODE As Long = &H40           ' SW2 Active Mode
' PCI & PCIe Switches
Private Const WD_DSW_RLY1_TEMP_MODE As Long = &H80        ' Activate RELAY #1 for overtemp
Private Const WD_DSW_RLY1_MODE As Long = &H40             ' Relay #1 Pulse or Latch
Private Const WD_DSW_RLY2_MODE As Long = &H20             ' Relay #2 Pulse or Latch
' Internal & External Serial
Private Const WD_DSW_IES_MONITOR As Long = &H80           ' Passive Monitor Mode
Private Const WD_DSW_IES_COMMAND As Long = &H40           ' Command Monitor Mode
Private Const WD_DSW_IES_RINGMON As Long = &H20           ' Ring RI Monitor Mode
Private Const WD_DSW_IES_POD As Long = &H10               ' Wait for first "tickle" after POD time
' Internal Serial
Private Const WD_DSW_IS_OVERTEMP_RST As Long = &H8
' External Serial
Private Const WD_DSW_EXS_NOTUSED As Long = &H8        ' Spare Switch - External Serial

' ------------------------------------------
' Board types for WD_Open_Ex() Function
Private Enum WDogTypes
  WD_TYPE_NONE_FOUND = 0
  WD_TYPE_USB_V1                   ' PN: 1140
  WD_TYPE_PCI                      ' PN: 1130 & 1134
  WD_TYPE_PCI_TYP2                 ' PN: 1131
  WD_TYPE_PCI_EXP                  ' PN: 1170
  WD_TYPE_ETH_USB                  ' PN: 1160
  WD_TYPE_INT_SER                  ' PN: 1120
  WD_TYPE_EXT_SER                  ' PN: 1110
End Enum


' ----------------------------------------------------
' Function Calling Flags
' ------------------------------------------
' Max Buffer Sizes in DLL
Private Const INFO_ERR_BUFF_MAXSIZE As Integer = 256

' ----------------------------------------
' Flags for POD Set operations
Private Const WD_POD_SETNVPOD As Long = &H2             ' Write new POD to non-volatile memory
Private Const WD_POD_SETPOD As Long = &H1               ' Update POD time if WD not active
' Flags from POD Get operations
Private Const WD_POD_SET_IGNORE As Long = &H1           ' new POD ignored - WD was already active

' ------------------------------------------
' Flags for WDog Timer Set operations
Private Const WD_SETNV_TIMEOUT As Long = &H2             ' Write new time to nv memory
Private Const WD_SET_TIMEOUT As Long = &H1               ' Set new time

' ------------------------------------------
' Clear Reset Count Flag
Private Const WD_CLEAR_RST_CNT As Long = &H1             ' clear the Reset Count & LED

' ------------------------------------------
' Temp Offset Write Enable Flag
Private Const WD_TEMP_OFF_WREN As Long = &H1             ' enable nv memory write

' ------------------------------------------
' Enable/Disable Flags for Send and Get
Private Const WD_WDOG_ENABLE As Long = &H2               ' enable the Watchdog
Private Const WD_WDOG_DISABLE As Long = &H1              ' disable the Watchdog

' ------------------------------------------
' Relay Pulse Write Enable Flag
Private Const WD_RLY_PLS_WREN As Long = &H1              ' enable nv memory write

' ------------------------------------------
' Set the Buzzer Flags
Private Const WD_BUZZ_GET_NOT_NV As Long = &H10000         ' get current buzzer time

Private Const WD_BUZZ_NV_CNTL_EN As Long = &H8000            ' must be set for NV disable/enable operations
Private Const WD_BUZZ_CNTL_EN As Long = &H4000            ' must be set for disable/enable operations
Private Const WD_BUZZ_NV_DISABLE As Long = &H2000         ' save disable in non-volatile memory
Private Const WD_BUZZ_DISABLE As Long = &H1000         ' buzzer disable

Private Const WD_BUZZ_NV_WREN As Long = &H2              ' enable nv memory write
Private Const WD_BUZZ_ON_EN As Long = &H1                ' turn buzzer on (off)

' ------------------------------------------
' User Code Write Enable Flag
Private Const WD_USE_CODE_WREN As Long = &H1             ' enable nv memory write

' ------------------------------------------
' Enable - Disable PC Reset Flags
Private Const WD_PC_RESET_NO_PEND As Long = &H80000       ' there was not a reset pending yet - issue reset command first
Private Const WD_PC_RESET_NO_TR As Long = &H40000         ' board firmware does not allow read of time remaining
Private Const WD_PC_RESET_NO_DIS As Long = &H20000        ' board firmware does not allow disable after enable
Private Const WD_PC_RESET_REJECT As Long = &H10000        ' command rejected - board was not ARMed yet

Private Const WD_PC_RESET_DIS As Long = &H2              ' stop the reset sequence if possible
Private Const WD_PC_RESET_EN As Long = &H1               ' reset the pc w/timer

' ------------------------------------------
' Flags for Aux Get/Set Relay
Private Const WD_AUX_RLY_INVRT_EN As Long = &H800         ' invert relay
Private Const WD_AUX_RLY_RST_EN As Long = &H200           ' enable relay reset action
Private Const WD_AUX_RLY_EN As Long = &H100               ' enable ON/OFF operation

Private Const WD_AUX_RLY_INVRT As Long = &H80             ' invert relay at power-up
Private Const WD_AUX_RLY_LATCH As Long = &H4              ' latch relay on at reboot
Private Const WD_AUX_RLY_PULSE As Long = &H2              ' relay pulse at reboot
Private Const WD_AUX_RLY_ON As Long = &H1                 ' turn relay on

' ------------------------------------------
' Flags for Digital In
Private Const WD_DIG_CLR_EXTCNT As Long = &H100           ' Clear External Trigger Count after read
Private Const WD_DIG_IN_EDGE_EN As Long = &H80            ' enable edge detection command bit
Private Const WD_DIG_IN_EDGE As Long = &H40               ' turn on edge detection
Private Const WD_DIG_IN_NVWR_EN As Long = &H20            ' nv write enable
Private Const WD_DIG_IN_NVEDGE As Long = &H10             ' nv turn on edge detection
' Flags returned
Private Const WD_DIG_IN_ERROR As Long = &H8000            ' wrong firmware - too low
Private Const WD_DIG_IN_EDGE_ACT As Long = &H80           ' edges are active for tickling
Private Const WD_DIG_IN_NVEDGE_ACT As Long = &H40         ' edges active are stored in NV memory
Private Const WD_DIGITAL_IN_EDGE As Long = &H2            ' edge has occurred
Private Const WD_DIGITAL_IN_HIGH As Long = &H1            ' input is currently high


' --------------------------------------------------
' Flag defines for PCI only functions
' ------------------------------------------
' Flags for PCI Digital IO
Private Const WD_PCI_DIG_OUT_EN As Long = &H1             ' enable writes to Digital out pins on PCI

' ------------------------------------------
' Flags for PCI Get Set Relays
' Flags for Set operations
Private Const WD_PCI_INVRT_EN As Long = &H800            ' invert relay #1
Private Const WD_PCI_HDW2_EN As Long = &H400             ' enable relay #2 hardware Port #1 options
Private Const WD_PCI_RLY2_EN As Long = &H200             ' enable ON/OFF operation of Relay #2
Private Const WD_PCI_RLY1_EN As Long = &H100             ' enable ON/OFF operation of Relay #1
' Flags for Get/Set operations
Private Const WD_PCI_NV_INVRT As Long = &H20             ' make invert a non-volatile setting
Private Const WD_PCI_INVRT_ON As Long = &H10             ' Invert Relay #1
Private Const WD_PCI_HDW_EXCL_ON As Long = &H8           ' set bit R2DS exclusive Relay #2 in hardware Port #1
Private Const WD_PCI_HDW_RLY2_ON As Long = &H4           ' turn on Relay #2 with hardware Port #1
Private Const WD_PCI_RLY2_ON As Long = &H2               ' turn on Relay #1
Private Const WD_PCI_RLY1_ON As Long = &H1               ' turn on Relay #1

' --------------------------------------------------
' Flag defines for External Serial only functions
' ------------------------------------------
' Flags Ext Serial Power Module
Private Const WD_POWERMODULE_ON As Long = &H1             ' turn on the power module
Private Const WD_POWERMODULE_OFF As Long = &H2            ' turn off

' --------------------------------------------------
' Flag defines for Ethernet-USB only functions
' --------------------------------------------------
' Get - Set Ethernet Command Allow
Private Const WD_ETH_ALLOW_WREN As Long = &H1             ' write the codes to non-volatile

' ------------------------------------------
' Set IP Addresses Flag
Private Const WD_NV_IPADD_WREN As Long = &H2              ' For writes - of NV Memory
Private Const WD_IP_ACTIVATE_NOW As Long = &H1            ' restart Ethernet with new address

' ------------------------------------------
' Get IP Addresses Flag
Private Const WD_NV_IPADD_RDEN As Long = &H1              ' Get the NV addresses

' ------------------------------------------
' Get / Set UDP Port Number
Private Const WD_NV_UDP_PORT_WREN As Long = &H1           ' Wrtite new port number to non-volatile memory

' -------------------------------------------------------------------------
' Ethernet Reset
Private Const WD_OVRR_DHCP_IP As Long = &HC0              ' DHCP IP override
Private Const WD_OVRR_NVMEM_IP As Long = &H80             ' EEPROM IP override
Private Const WD_OVRR_FIXED_IP As Long = &H40             ' fixed IP override

Private Const WD_REBOOT_ETHERNET As Long = &H2            ' Perform soft reset
Private Const WD_RESET_ETHERNET As Long = &H1             ' Perform Hard reset

' ----------------------------------------
' Get Set USB Suspend Mode
Private Const WD_SUSPEND_IGNORE As Long = &H4             ' if set - then ignore suspend
Private Const WN_NV_IGNORE_WREN As Long = &H2             ' Write status to nv memory
Private Const WN_NV_SUSP_TIME_WREN As Long = &H1          ' Write time to NV memory

' -------------------------------
' Ethernet Commands Allowed
Private Const E_Allow_WD_GetDeviceInfo As Long = &H1                 ' Bit 0
Private Const E_Allow_WD_GetTempTickle As Long = &H2                 ' Bit 1
Private Const E_Allow_WD_SetPowerOnDlyTimes As Long = &H4            ' Bit 2
Private Const E_Allow_WD_GetPowerOnDlyTimes As Long = &H8            ' Bit 3
Private Const E_Allow_WD_SetWdogTimes As Long = &H10                 ' Bit 4
Private Const E_Allow_WD_GetWdogTimes As Long = &H20                 ' Bit 5
Private Const E_Allow_WD_GetAnalogDigitalIn As Long = &H40           ' Bit 6
Private Const E_Allow_WD_GetSetAuxRelay As Long = &H80               ' Bit 7
Private Const E_Allow_WD_EnableDisable As Long = &H100               ' Bit 8
Private Const E_Allow_WD_GetResetCount As Long = &H200               ' Bit 9
Private Const E_Allow_WD_GetSetNvTempOffset As Long = &H400          ' Bit 10
Private Const E_Allow_WD_SetBuzzer As Long = &H800                   ' Bit 11
Private Const E_Allow_WD_GetBuzzer As Long = &H1000                  ' Bit 12
Private Const E_Allow_WD_GetSetNvRelayPulse As Long = &H2000         ' Bit 13
Private Const E_Allow_WD_GetSetNvUsbUserCode As Long = &H4000        ' Bit 14
Private Const E_Allow_WD_EnableDisablePcReset As Long = &H8000       ' Bit 15


' ----------------------------------------------------
' Private declarations of variables
' ----------------------------------------------------
Private wdStatus    As Long
Private wdHandle    As Long


' --------------------------------------------------------
'   Start Function Defs
' ----------------------------------------
' Open Device - Must be called first
Private Declare Function WD_Open Lib "WDog_Univrsl.dll" (ByRef pwdHandle As Long) As Long

' ----------------------------------------
' Open Device Extended - Must be called first
Private Declare Function WD_OpenEx Lib "WDog_Univrsl.dll" (ByRef pwdHandle As Long, _
                                                           ByVal iType As Long) As Long

' ----------------------------------------
' Close Device - Should be called last
Private Declare Function WD_ForceCom Lib "WDog_Univrsl.dll" (ByVal iCom As Long) As Long

' ----------------------------------------
' Force COM port - Call before OpenEx
Private Declare Function WD_Close Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long) As Long

' ----------------------------------------
' Get DLL Version
'    Private Declare Function WD_GetDllVersion Lib "WDog_Univrsl.dll" _
     '                (ByRef pDllVersion As long) As long

Private Declare Function WD_GetDllVersion Lib "WDog_Univrsl.dll" (ByRef pDllVersion As Long) As Long


' ----------------------------------------
'Get Error / Info Messages
Private Declare Function WD_GetErrorInfoMsg Lib "WDog_Univrsl.dll" (ByVal ErrorMsg As String, _
                                                                    ByVal InfoMsg As String) As Long

' ----------------------------------------
' Get board type
Private Declare Function WD_GetWDogType Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                ByRef iType As Long) As Long

' ----------------------------------------
' Get Com Port Open number
Private Declare Function WD_GetComPortNum Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                  ByRef iCpn As Long) As Long

' ----------------------------------------
' Get Info from board
Private Declare Function WD_GetDeviceInfo Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                  ByRef pStat As Long, ByRef pDipSw As Long, ByRef pVer As Long, _
                                                                  ByRef pTick As Long, ByRef pDiag As Long) As Long

' ------------------------------------------
' Get Temp & Tickle the board
Private Declare Function WD_GetTempTickle Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                  ByRef pTempw As Long, ByRef pTempf As Long, ByRef pTick As Long) _
                                                                  As Long

' ------------------------------------------
' Set POD times
Private Declare Function WD_SetPowerOnDlyTimes Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                       ByVal iPod As Long, ByVal iNvPod As Long, ByVal iSetFlag As Long, _
                                                                       ByRef pResFlag As Long) As Long

' ------------------------------------------
' Get POD times
Private Declare Function WD_GetPowerOnDlyTimes Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                       ByRef pPod As Long, ByRef pNvPod As Long) As Long

' ------------------------------------------
' Set Watchdog Countdown Times
Private Declare Function WD_SetWdogTimes Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                 ByVal iWdTime As Long, ByVal iNvWdTime As Long, ByVal iFlag As Long) _
                                                                 As Long

' ------------------------------------------
' Get Watchdog Countdown Times
Private Declare Function WD_GetWdogTimes Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                 ByRef pWdTime As Long, ByRef pNvWdTime As Long, ByRef pHoldRegTime As Long, _
                                                                 ByRef pDipSwTime As Long) As Long

' ------------------------------------------
' Get (Clear) Reset Count
Private Declare Function WD_GetResetCount Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                  ByVal iFlag As Long, ByRef pRstCnt As Long) As Long

' ------------------------------------------
' Get / Set NV Temp Offset
Private Declare Function WD_GetSetNvTempOffset Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                       ByVal iFlag As Long, ByVal iNvOffset As Long, _
                                                                       ByRef pCurOffset As Long) As Long

' ------------------------------------------
' WD Enable Disable
Private Declare Function WD_EnableDisable Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                  ByVal iFlagSet As Long) As Long

' ------------------------------------------
' Get / Set NV Relay Pulse
Private Declare Function WD_GetSetNvRelayPulse Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                       ByVal iFlag As Long, ByVal iNvRelayPulse As Long, _
                                                                       ByRef pCurRelayPulse As Long) As Long

' ------------------------------------------
' Set the Buzzer Times
Private Declare Function WD_SetBuzzer Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                              ByVal iBuzzTime As Long, ByVal iNvBuzzTime As Long, _
                                                              ByVal iFlags As Long) As Long

' ------------------------------------------
' Get the Buzzer Times
Private Declare Function WD_GetBuzzer Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                              ByRef pFlag As Long, ByRef pBuzzTime As Long) As Long

' ------------------------------------------
' Get / Set NV User Code
Private Declare Function WD_GetSetNvUserCode Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                     ByVal iFlag As Long, ByRef pNvUserCode As Byte, ByRef pCurNvUserCode As Byte) _
                                                                     As Long

' ------------------------------------------
' Enable / Disable PC Reset
Private Declare Function WD_EnableDisablePcReset Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                         ByVal iFlag As Long, ByVal iResetTime As Long, _
                                                                         ByRef pGetTime As Long) As Long

' ----------------------------------------
' Get & Set Aux Relay
Private Declare Function WD_GetSetAuxRelay Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                   ByVal iRelaySet As Long, ByRef pRelayGet As Long) As Long

' ----------------------------------------
' Get Digital In
Private Declare Function WD_GetDigitalIn Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                 ByRef pExtCount As Long, ByRef pDi As Long, ByVal iFlag As Long) As Long

' ------------------------------------------
' Get Analog Input
Private Declare Function WD_GetAnalogIn Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                ByRef pAi As Long) As Long

' ----------------------------------------
' Get / Set PCI Digital IO
Private Declare Function WD_GetSetPciDigitalInOut Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                          ByVal iDigOut As Long, ByRef pDigIn As Long, ByVal iFlag As Long) As Long

' ------------------------------------------
' Get & set PCI Relays
Private Declare Function WD_GetSetPciRelays Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                    ByVal iRelaySet As Long, ByRef pRelayGet As Long) As Long

' ----------------------------------------
' Get Set External Serial Power Module
Private Declare Function WD_GetSetPowerModule Lib "WDog_Univrsl.dll" (ByVal wdHandle As Long, _
                                                                      ByVal iFlagSet As Long, ByRef iFlagGet As Long) As Long

Public Property Get IsOpen() As Boolean

  IsOpen = wdStatus <> 0

End Property

Public Sub SetTimeOut(ByVal Seconds As Long)
  Dim SomeValue     As Long
  On Error Resume Next

  wdStatus = WD_SetWdogTimes(wdHandle, Seconds, SomeValue, WD_SET_TIMEOUT)
  'WD_BD_TYPE_UNKNOWN = 14
End Sub
Public Property Get RemainingTime() As Long

  Dim WDTime        As Long
  Dim NVWDTime      As Long
  Dim HoldReg       As Long
  Dim DipSwTime     As Long

  On Error Resume Next

  wdStatus = WD_GetWdogTimes(wdHandle, WDTime, NVWDTime, HoldReg, DipSwTime)
  'If wdStatus = WDogFlags.WD_OK Then
  RemainingTime = WDTime
  'End If


End Property

Public Function CloseWD() As Long
  On Error Resume Next
  If wdHandle <> 0 Then
    Disable
    wdStatus = WD_Close(wdHandle)
  End If
  wdHandle = 0

End Function

Public Function Enable() As Boolean
  Dim Status As Long, DSW As Long, Ver As Long, Tick As Long, Diag As Long
  On Error Resume Next


  wdStatus = WD_EnableDisable(wdHandle, WD_WDOG_ENABLE)
  wdStatus = wdStatus Or (WD_GetDeviceInfo(wdHandle, Status, DSW, Ver, Tick, Diag))

  If (Status And (WD_STAT_ACTIVE_ARMED)) Then
    Enable = True
  End If

End Function


Public Function Disable() As Boolean
  Dim Status As Long, DSW As Long, Ver As Long, Tick As Long, Diag As Long
  On Error Resume Next


  wdStatus = WD_EnableDisable(wdHandle, WD_WDOG_DISABLE)
  wdStatus = wdStatus Or (WD_GetDeviceInfo(wdHandle, Status, DSW, Ver, Tick, Diag))

  If (Status And (WD_STAT_CMD_DISABLED Or WD_STAT_PCI_WDIS)) Then
    Disable = True
  End If

End Function

Public Property Get Handle() As Long
  Handle = wdHandle
End Property

Public Property Get Status() As Long
  Status = wdStatus
End Property

Public Property Get Version() As String
  Dim VersionData   As Long
  Dim Ver(2)        As Byte
  Dim i             As Long

  On Error Resume Next


  wdStatus = WD_GetDllVersion(VersionData)

  For i = 2 To 0 Step -1
    Ver(i) = CByte(VersionData And &HFF)
    VersionData = VersionData / 256               ' iVersion >> 8
    VersionData = VersionData And &HFFFFFF
  Next

  Version = Hex(Ver(0)) & "." & Hex(Ver(1)) & "." & Hex(Ver(2))

  'txtOutput.Text = txtOutput.Text & "DLL Version: " & Hex(bVer(0)) & "." _
   & Hex(bVer(1)) & "." & Hex(bVer(2)) & vbNewLine
  'txtOutput.Text = txtOutput.Text & "DLL Type #: " & Hex(iVersion And &HFF) & vbNewLine


End Property

Public Function InitWatchdog() As Boolean
  On Error Resume Next
  If Handle = 0 Then
    wdStatus = WD_Open(wdHandle)
    ' WD_DEVICE_NOT_FOUND = 1
  End If
  InitWatchdog = wdHandle <> 0

End Function

Public Function Tickle() As Boolean
  Dim TempWhole     As Long
  Dim TempFraction  As Long
  Dim Ticks         As Long
  On Error Resume Next
  wdStatus = WD_GetTempTickle(wdHandle, TempWhole, TempFraction, Ticks)
  Temperature = Val(TempWhole & "." & TempFraction)
  Tickles = Ticks


End Function

Private Sub Class_Terminate()
  CloseWD
End Sub
