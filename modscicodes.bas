Attribute VB_Name = "modSCICodes"
Option Explicit


'Priority/Severity
'Category
'Description
'0 emergency
'System or device is unusable; Hard fault/Total failure or catastrophic occurrence.
'1 Alert
'Action must be taken immediately
'2 Critical
'Critical condition(s) occurred
'3 Error
'Error condition(s) occurred
'4 Warning
'Significant/abnormal/warning conditions have occurred
'5 Notice
'Normal but noteworthy conditions occurred
'6 Informational
'Informative Messages
'7 Debug/Diagnostic
'System/Device debug messages

Global Const SCI_PRIORITY_EMERGENCY = 0
Global Const SCI_PRIORITY_ALERT = 1
Global Const SCI_PRIORITY_CRITICAL = 2
Global Const SCI_PRIORITY_ERROR = 3
Global Const SCI_PRIORITY_WARNING = 4
Global Const SCI_PRIORITY_NOTICE = 5
Global Const SCI_PRIORITY_INFOMATIONAL = 6
Global Const SCI_PRIORITY_DEBUG_DIAG = 7
Global Const SCI_PRIORITY_8 = 8
Global Const SCI_PRIORITY_9 = 9
Global Const SCI_PRIORITY_10 = 10


'SCI_Codes
Global Const SCI_CODE_ALARM1 = 1
Global Const SCI_CODE_ALARM1_CLEAR = 2
Global Const SCI_CODE_ALARM2 = 3
Global Const SCI_CODE_ALARM2_CLEAR = 4
Global Const SCI_CODE_ALARM3 = 5
Global Const SCI_CODE_ALARM3_CLEAR = 6
Global Const SCI_CODE_ALARM4 = 7
Global Const SCI_CODE_ALARM4_CLEAR = 8
Global Const SCI_CODE_DEVICE_INACTIVE = 9
Global Const SCI_CODE_DEVICE_INACTIVE_CLEARED = 10
Global Const SCI_CODE_TAMPER = 11
Global Const SCI_CODE_TAMPER_CLEAR = 12
Global Const SCI_CODE_EOL_TAMPER = 13
Global Const SCI_CODE_EOL_TAMPER_CLEAR = 14
Global Const SCI_CODE_LOW_BATT = 15
Global Const SCI_CODE_LOW_BATT_CLEAR = 16
Global Const SCI_CODE_MAINT_DUE = 17
Global Const SCI_CODE_MAINT_DUE_CLEAR = 18
Global Const SCI_CODE_19 = 19
Global Const SCI_CODE_20 = 20
Global Const SCI_CODE_DEVICE_RESET = 21
Global Const SCI_CODE_ENDPOINT_FAIL = 25
Global Const SCI_CODE_ENDPOINT_SUCCESS = 26

Global Const CODE_1941XS_OK = 0
Global Const CODE_1941XS_RESET = 4
Global Const CODE_1941XS_TAMPER = 8
Global Const CODE_1941XS_PUSHBUTTON_DOWN = &H10
Global Const CODE_1941XS_PUSHBUTTON = &H20
Global Const CODE_1941XS_PULLCORD = &H40



'Global Const SCI_CODE_ALARM1_AND_ALARM2 = -(SCI_CODE_ALARM1 + SCI_CODE_ALARM2)
'Global Const SCI_CODE_ALARM1_AND_ALARM2_CLEAR = -(SCI_CODE_ALARM1_CLEAR + SCI_CODE_ALARM2_CLEAR)
'Global Const SCI_CODE_TAMPER_XS = -(SCI_CODE_TAMPER)
'Global Const SCI_CODE_TAMPER_CLEAR_XS = -(SCI_CODE_TAMPER_CLEAR)
'Global Const SCI_CODE_ALARM2_AND_TAMPER_XS = -(SCI_CODE_ALARM2 + SCI_CODE_TAMPER)
'
'Global Const SCI_CODE_ALARM1_AND_ALARM2_AND_TAMPER = -(SCI_CODE_ALARM1 + SCI_CODE_ALARM2 + SCI_CODE_TAMPER)
'Global Const SCI_CODE_ALARM1_AND_ALARM2_AND_TAMPER_CLEAR = -(SCI_CODE_ALARM1_CLEAR + SCI_CODE_ALARM2_CLEAR + SCI_CODE_TAMPER_CLEAR)



' REPEATERS
Global Const SCI_CODE_REPEATER_POWER_LOSS = 27
Global Const SCI_CODE_REPEATER_POWER_LOSS_CLEAR = 28
Global Const SCI_CODE_REPEATER_RESET = 29
Global Const SCI_CODE_REPEATER_TAMPER = 31
Global Const SCI_CODE_REPEATER_TAMPER_CLEAR = 32
Global Const SCI_CODE_REPEATER_LOW_BATTERY = 33
Global Const SCI_CODE_REPEATER_LOW_BATTERY_CLEAR = 34
Global Const SCI_CODE_REPEATER_JAM = 35
Global Const SCI_CODE_REPEATER_JAM_CLEAR = 36
Global Const SCI_CODE_REPEATER_INACTIVE = 43
Global Const SCI_CODE_REPEATER_INACTIVE_CLEAR = 44
Global Const SCI_CODE_REPEATER_CONFIG_FAIL = 45
Global Const SCI_CODE_REPEATER_CONFIG_SUCCESS = 46


''AGC
'
Global Const SCI_CODE_ACG_RESET = 49
Global Const SCI_CODE_ACG_TAMPER = 51
Global Const SCI_CODE_ACG_TAMPER_CLEAR = 52
Global Const SCI_CODE_ACG_JAMMED = 55
Global Const SCI_CODE_ACG_JAM_CLEAR = 56
Global Const SCI_CODE_ACG_INACTIVE = 57
Global Const SCI_CODE_ACG_INACTIVE_CLEAR = 58
Global Const SCI_CODE_ACG_CONFIG_FAIL = 59
Global Const SCI_CODE_ACG_CONFIG_SUCCESS = 60
Global Const SCI_CODE_ACG_CONFIG_CRC_FAIL = 61
Global Const SCI_CODE_ACG_FW_SUCCESS = 71
Global Const SCI_CODE_ACG_FW_FAIL = 72
Global Const SCI_CODE_ACG_BATTERY_FAIL = 91
Global Const SCI_CODE_ACG_LOW_BATTERY = 92
Global Const SCI_CODE_ACG_BATTERY_OK = 93
Global Const SCI_CODE_ACG_SHUTDOWN_IMMINENT = 96
Global Const SCI_CODE_ACG_FW_PENDING = 97
Global Const SCI_CODE_ACG_IP_PROCESSOR_CRC_FAIL = 99
Global Const SCI_CODE_ACG_REBOOT_REQUESTED = 100
Global Const SCI_CODE_ACG_HELLO = 125

' 1941XS from "Status Solutions"
Global Const SCI_CODE_SERIALDATA = 147



'1 Alarm1
'2 Alarm1 has cleared
'3 Alarm2
'4 Alarm2 has cleared
'5 Alarm3
'6 Alarm3 has cleared
'7 Alarm3
'8 Alarm3 has cleared
'9 Device is Inactive
'10 Device is Inactive has cleared (is now active)
'11 Tamper activated
'12 Tamper cleared
'13 EOL tamper activated
'14 EOL tamper cleared
'15 low Battery
'16 Low battery cleared
'17 Maintenance required
'18 Maintenance required cleared
'21 device Reset
'25 Endpoint Configuration Fail
'26 Endpoint Configuration Success
'
'' Repeaters
'27 Repeater Power Loss
'28 Repeater Power Loss clear
'29 Repeater Reset
'31 Repeater Tamper
'32 Repeater tamper clear
'33 Repeater low battery
'34 Repeater low battery clear
'35 Repeater Jam
'36 Repeater jam clear
'43 Repeater is inactive
'44 Repeater is inactive clear (is now active)
'45 Repeater Configuration fail
'46 Repeater Configuration Success
'
''AGC
'
'49 ACG Reset
'51 ACG Tamper
'52 ACG Tamper clear
'55 ACG Jammed
'56 ACG Jam clear
'57 ACG Inactive
'58 ACG Inactive clear
'59 ACG Configuration Fail
'60 ACG Configuration success
'61 ACG CRC Check Fail
'71 ACG F/W Update Success
'72 ACG F/W Update Failed
'91 ACG Battery Failed
'92 ACG Battery low
'93 ACG Battery OK
'96 ACG Shutdown Imminent
'97 ACG F/W Update Pending
'99 ACG IP Processor CRC Invalid
'100 ACG Reboot Requested
'
'ValueState
'System
'OK , Other, Fault, Maintenance
'device
'OK , Other, Failed
'Comms
'OK , Other, Marginal, Failed
'Power
'OK , Other, Failed, LowBattery
'Arm
'Other , Away, Stay, disarmed
'ArmReady
'Other , NotReady, Ready
'IntrusionAlarm
'OK , Other, Panic, Tamper, Intrusion, Fire, Duress, Technical
'IntrusionFault
'OK , Other, Fault
'IntrusionTrouble
'OK , Other, Trouble
'bypass
'Other , Active, bypass
'Latch
'Other , Latched, Unlatched, Unknown
'Door
'Other, Open, Closed, Forced, held, Unknown
'AccessOverride
'Other , Normal, Overridden
'mask
'Other , Masked, Monitored
'Tamper
'OK , Other, Tamper
'Granted
'Other , OK
'Denied
'Other , UnknownCredential, InvalidCredential, InactiveCredential, ExpiredCredential, LostCredential, StolenCredential, InvalidCredentialFormat, InvalidFacilityCode, InvalidIssueCode, AuthenticationTimeout, AuthenticationFailure, MaxRetriesReached, InactiveCredentialHolder, ExpiredCredentialHolder, NotPremitted, NotPermittedAtThisTime, AntiPassback, AccessOverridden, NoAsset, NoEscort, OccupancyLimitReached, CredentialNotPresented, UseLimitReached, PartitionClosed, Unauthorized, BiometricMismatch, InvalidPIN, AccessOverridden
'Connection
'OK , Other, Disconnected, Unknown
'InputFault
'OK , Other, Cut, Short
'Active
'Other , Inactive, Active
'
'






'Global Const EVT_NONE = 0               ' NON-EVENT EVENT
'Global Const EVT_EMERGENCY = 1          ' ALARM
'Global Const EVT_EMERGENCY_RESTORE = 2  ' RESTORE AFTER ALARM
'Global Const EVT_EMERGENCY_ACK = 3      ' MANUAL OR AUTO ACKNOWLEDGE
'Global Const EVT_BATTERY_FAIL = 4       ' LOW BATTERY
'Global Const EVT_BATTERY_RESTORE = 5    ' LOW BATTERY RESTORE
'Global Const EVT_CHECKIN_FAIL = 6       ' DEVICE AUTO-CHECKIN FAILED
'Global Const EVT_CHECKIN = 7            ' ONLY AFTER A FAILED CHECKIN
'Global Const EVT_UNASSIGNED = 8         ' IN SYSTEM, BUT NOT ASSIGNED
'Global Const EVT_STRAY = 9              ' NOT IN SYSTEM
'Global Const EVT_COMM_TIMEOUT = 10      ' TOO MUCH TIME SINCE LAST COMM DATA / SERIAL PORT DEAD
'Global Const EVT_COMM_RESTORE = 11      ' GETTING DATA AGAIN
'Global Const EVT_ASSUR_CHECKIN = 12     ' CHECKED IN
'Global Const EVT_ASSUR_FAIL = 13        ' FAILED TOCHECK IN
'Global Const EVT_ALERT = 14             ' ALERT INSTEAD OF ALARM
'Global Const EVT_ALERT_RESTORE = 15     ' ALERT INSTEAD OF ALARM
'Global Const EVT_ALERT_ACK = 16         ' ALERT INSTEAD OF ALARM
'Global Const EVT_SILENCE = 17           ' SILENCED (BUT NOT ACKNOWLEDGED)
'Global Const EVT_LOCATED = 18           ' LOCATOR FORWARD
'Global Const EVT_TAMPER = 19            ' TAMPER TRIGGERED
'Global Const EVT_TAMPER_RESTORE = 20    ' TAMPER BIT RESTORED
'Global Const EVT_ANNOUNCE_1 = 21        ' STANDARD ANNOUNCE
'Global Const EVT_ANNOUNCE_2 = 22        ' ESCALATED ANNOUNCE
'Global Const EVT_ANNOUNCE_3 = 23        ' 3RD LEVEL ESCALATED ANNOUNCE
'Global Const EVT_DATABASE_UPDATE = 24   ' We've written to the database
'Global Const EVT_DATABASE_READ = 25     ' We've checked the database for updates from other consoles
'Global Const EVT_GENERAL_TROUBLE = 26   ' UNDEFINED TROUBLE
'Global Const EVT_ASSUR_START = 27
'Global Const EVT_ASSUR_END = 28
'Global Const EVT_VACATION = 29
'Global Const EVT_VACATION_RETURN = 30
'Global Const EVT_EMERGENCY_END = 31
'Global Const EVT_ALERT_END = 32
'Global Const EVT_ADD_RES = 33
'Global Const EVT_REMOVE_RES = 34
'Global Const EVT_ADD_DEV = 35
'Global Const EVT_REMOVE_DEV = 36
'Global Const EVT_ASSIGN_DEV = 37
'Global Const EVT_UNASSIGN_DEV = 38
'Global Const EVT_LOCATE = 39
'Global Const EVT_LINELOSS = 40
'Global Const EVT_LINELOSS_RESTORE = 41
'Global Const EVT_JAMMED = 42
'Global Const EVT_JAMM_RESTORE = 43
'
'Global Const EVT_SYSTEM_START = 44
'Global Const EVT_SYSTEM_STOP = 45
'Global Const EVT_SYSTEM_LOGIN = 46
'Global Const EVT_SYSTEM_LOGOUT = 47
'
'Global Const EVT_EXTERN = 48          ' EXTERN INSTEAD OF ALARM
'Global Const EVT_EXTERN_RESTORE = 49
'Global Const EVT_EXTERN_ACK = 50
'Global Const EVT_EXTERN_END = 51
'
'Global Const EVT_EXTERN_TROUBLE = 52   ' EXTERNAL DEVICE PORT FAILURE/CONNECTOR FAILURE
'Global Const EVT_EXTERN_TROUBLE_RESTORE = 53
'
'
'Global Const EVT_AUTOACK = 54           ' general device alarm autoclear
'Global Const EVT_EMERGENCY_AUTOACK = 55  ' device emergency autoclear
'Global Const EVT_ALERT_AUTOACK = 56     ' device alert autoclear
'Global Const EVT_EXTERN_AUTOACK = 57    ' device extern autoclear
'
'Global Const EVT_PTI_MISMATCH = 58      ' packet PTI and Device PTI don't match
'Global Const EVT_STATUS_ERROR = 59      ' packet Status word too big => 32737
'Global Const EVT_BATT_TAMPER = 60       ' packet Battery and Tamper in same packet
'Global Const EVT_PCA_REG = 61           ' PCA registration packet... not an alarm
'
'Global Const EVT_FORCED_LOGOUT = 62
'Global Const EVT_MAXDEVICE = 63

