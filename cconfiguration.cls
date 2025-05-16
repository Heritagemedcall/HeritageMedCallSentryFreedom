VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public KeepOnPaging   As Long

Public HideHIPPANames As Long
Public HideHIPPASidebar As Long

Private mAdminContact As String

Public Facility     As String
Public RxTimeout    As Long
Public ID           As Long
Public ConnectString As String
Private mCommPort   As Integer

Public EscTimer     As Long

Public RemoteSerial  As String


Public AlarmFile    As String
Public AlertFile    As String
Public LowBattFile  As String
Public TroubleFile  As String
Public AssurFile    As String
Public ExtFile      As String

Public ReportPath   As String

Public WatchdogType As Long
Public WatchdogTimeout As Long

Public MonitorDomain As String
Public MonitorRequest As String
Public MonitorPort  As Long
Public MonitorInterval As Long
Public MonitorEnabled As Long
Public MonitorFacilityID As String


Public AlarmBeep    As Long
Public AlertBeep    As Long
Public LowBattBeep  As Long
Public TroubleBeep  As Long
Public AssurBeep    As Long
Public ExtBeep      As Long


Public AlarmReBeep    As Long
Public AlertReBeep    As Long
Public LowBattReBeep  As Long
Public TroubleReBeep  As Long
Public AssurReBeep    As Long
Public ExtReBeep      As Long

Public BeepControl    As Long


Public AssurStart   As Long
Public AssurEnd     As Long

Public AssurStart2  As Long
Public AssurEnd2    As Long

Public StartNight   As Long
'Public EndFisrt     As Long ' for readability, a property that just echos StartNight
Public EndNight     As Long
Public EndThird     As Long

'Public StartNight2  As Long ' new 5/10/12 for 3 shifts
'Public EndNight2    As Long


Public ESLastMessage As Long

Public SendAckMSG   As Integer

Public PCARedirect  As Integer

Public RxLocation   As String

Public RxSerial     As String

Private mSurveyDevice As String
Private mSurveyPCA  As String


Private mWaypointDevice As String

Public locationtext     As Integer


Private mHostIP     As String
Private mHostPort   As Long

Public surveymode   As Integer
Public SurveyPager  As Long

Public OnlyLocators As Integer   ' 0 use all dvices, 1 only locators (DNet)
Public NoNCs        As Integer   ' 0 use ncs for locating

'Local backups
Public BackupType   As Integer    ' 0 = days, 1 = monthly
Public BackupDOW    As Integer    ' bitfield
Public BackupDOM2   As String     ' from a string array
Public BackupDOM    As Integer    ' number
Public BackupEnabled As Integer   ' 0 or 1
Public BackupTime   As Single     ' integer for hour, decimal for fraction
Public BackupFolder As String     ' path to backup files

'Remote backups (FTP)
Public BackupTypeRemote As Integer      ' 0 = days, 1 = monthly
Public BackupDOWRemote As Integer       ' bitfield
Public BackupDOMRemote As Integer       ' number
Public BackupDOMRemote2 As String       ' from a string array
Public BackupEnabledRemote As Integer   ' 0 or 1
Public BackupTimeRemote As Single       ' integer for hour, decimal for fraction
Public BackupFolderRemote As String     ' path to backup files

Public BackupHost   As String
Public BackupUser   As String
Public BackupPassword As String

' assurance

Private mAssurSaveAsFile As Integer    ' 0 or 1
Public AssurSendAsEmail As Integer     ' 0 or 1
Private mAssurFileFormat As Integer    ' 0=TabDelimited; 1=TabDelimited-noheader;2=HTML
Public AssurEmailRecipient As String
Public AssurEmailSubject As String

' auto reports

Public AutoReportsSaveAsFile As Integer      ' 0 or 1
Public AutoReportsSendAsEmail As Integer     ' 0 or 1
Public AutoReportsFileFormat As Integer      ' 0=TabDelimited; 1=TabDelimited-noheader;2=HTML
Public AutoReportsEmailRecipient As String
Public AutoReportsEmailSubject As String


' autoreports list print folder and printer
Public AutoReportsListFolder As String
Public AutoReportsListPrinter As String


' SMTP settings
Public UseSMTP      As Integer                    ' 0 = mapi, 1 = smtp , 2... future:(html or google secure mail, or Imap)
Public MailSMTPserver As String
Public MailPOP3Server As String
Public MailSenderName As String
Public MailSenderEmail As String
Public MailUserName As String
Public MailPassword As String
Public MailRequireLogin As Integer               '0 or 1
Public MailRequirePopLogin As Integer            '0 or 1
Public MailUseSMTP  As Integer
Public MailDebug    As Integer
Public MailPort     As Long


Public ReminderMsgDelay As Long
Public ReminderMsgRepeats As Long
Public ReminderMsgSpacing As Long
Public ReminderRedials As Long
Public ReminderRedialDelay As Long
Public ReminderAckDigit As Long


Public SyncHostTime As Long

' Mobile settings
Public MobilehtPasswordPath As String ' where we keep the .htpasswd file
Public MobilehtPasswordEXEPath As String   ' path to htAccess EXE
Public MobileWebRoot  As String         ' root of our website
Public MobileWebEnabled As Long  ' 0 or 1

Public MobileClearAssist As Long
Public MobileClearHistory As Long



Public boost As Double ' for weaker tx devices  * (1 + (boost/100))





Public Property Get SurveyDevice() As String
  SurveyDevice = mSurveyDevice
End Property

Public Property Let SurveyDevice(ByVal Value As String)
  Dim checkvalue As Double
  On Error Resume Next
  checkvalue = Val("&h" & Value & "&")
  mSurveyDevice = Right("00000000" & Hex(checkvalue), 8)
End Property

Public Property Get SurveyPCA() As String
  SurveyPCA = mSurveyPCA
End Property

Public Property Let SurveyPCA(ByVal Value As String)
  Dim checkvalue As Double
  On Error Resume Next
  checkvalue = Val("&h" & Value & "&")
  mSurveyPCA = Right("00000000" & Hex(checkvalue), 8)
End Property



Public Property Get HostPort() As Long
  HostPort = mHostPort
End Property

Public Property Let HostPort(ByVal Value As Long)
  Value = Max(1, Min(65000, Value))
  mHostPort = Value
End Property

Private Sub Class_Initialize()
  mHostPort = 2500
  mHostIP = "127.0.0.1"
End Sub

Public Property Get HostIP() As String
  HostIP = GetValidIP(mHostIP, "127.0.0.1")
End Property

Public Property Let HostIP(ByVal IP As String)
  mHostIP = IP
End Property

Public Property Get CommPort() As Integer
  CommPort = Max(1, mCommPort)
End Property

Public Property Let CommPort(ByVal CommPort As Integer)
  mCommPort = CommPort
End Property

Public Property Get AssurFileFormat() As Integer

  AssurFileFormat = mAssurFileFormat

End Property

Public Property Let AssurFileFormat(ByVal Value As Integer)
' for now limit to values 0 thru 2
  Value = Min(2, Value)
  Value = Max(0, Value)
  mAssurFileFormat = Value

End Property

Public Property Get AssurSaveAsFile() As Integer

  AssurSaveAsFile = 1 ' mAssurSaveAsFile

End Property

Public Property Let AssurSaveAsFile(ByVal Value As Integer)

  mAssurSaveAsFile = 1

End Property

Public Property Get WaypointDevice() As String

  WaypointDevice = mWaypointDevice

End Property

Public Property Let WaypointDevice(ByVal WaypointDevice As String)

  mWaypointDevice = WaypointDevice

End Property

Public Property Get EndFirst() As Long

  EndFirst = StartNight

End Property

Public Property Let EndFirst(ByVal Value As Long)

  StartNight = Value

End Property

Public Property Get AdminContact() As String

  If Len(mAdminContact) = 0 Then
    AdminContact = "Contact Your System Administrator"
  Else
    AdminContact = mAdminContact
  End If
End Property

Public Property Let AdminContact(ByVal Value As String)

  mAdminContact = Trim$(Value)

End Property
