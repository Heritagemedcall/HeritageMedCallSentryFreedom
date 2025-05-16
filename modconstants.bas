Attribute VB_Name = "modConstants"
Option Explicit


Global Const DELETE_RESIDENTS = 2
Global Const DELETE_ROOMS = 4
Global Const DELETE_TRANSMITTERS = 8

Global Const MAIL_MAPI = 0
Global Const MAIL_SMTP = 1
Global Const MAIL_IMAP = 2

Global Const TRANSPORT_NONE = 0
Global Const TRANSPORT_EMAIL = 1
Global Const TRANSPORT_FTP = 2
Global Const TRANSPORT_POST = 3

Global Const PROTOCOL_PCA         As Integer = -1
Global Const PROTOCOL_NONE        As Integer = 0
Global Const PROTOCOL_TAP         As Integer = 1
Global Const PROTOCOL_COMP1       As Integer = 2
Global Const PROTOCOL_COMP2       As Integer = 3
Global Const PROTOCOL_TTS         As Integer = 4
Global Const PROTOCOL_EMAIL       As Integer = 5
Global Const PROTOCOL_DIALER      As Integer = 6
Global Const PROTOCOL_WEB         As Integer = 71
Global Const PROTOCOL_ONTRAK      As Integer = 8
Global Const PROTOCOL_CENTRAL     As Integer = 9 ' SDACT2 is similar
Global Const PROTOCOL_MARQUIS     As Integer = 10
Global Const PROTOCOL_PET         As Integer = 11
Global Const PROTOCOL_DIALOGIC    As Integer = 12
Global Const PROTOCOL_REMOTE      As Integer = 13
Global Const PROTOCOL_TAP_IP      As Integer = 14
Global Const PROTOCOL_APOLLO      As Integer = 15
Global Const PROTOCOL_MOBILE      As Integer = 16
Global Const PROTOCOL_SDACT2      As Integer = 17
Global Const PROTOCOL_TAP2        As Integer = 18


Global Const PROTOCOL_NONE_TEXT     As String = "NONE/TEXT"
Global Const PROTOCOL_TAP_TEXT      As String = "TAP"
Global Const PROTOCOL_COMP1_TEXT    As String = "COMP1"
Global Const PROTOCOL_COMP2_TEXT    As String = "COMP2"
Global Const PROTOCOL_TTS_TEXT      As String = "PA VOICE"
Global Const PROTOCOL_EMAIL_TEXT    As String = "EMAIL"
Global Const PROTOCOL_PCA_TEXT      As String = "PCA"
Global Const PROTOCOL_DIALER_TEXT   As String = "DIALER"
Global Const PROTOCOL_WEB_TEXT      As String = "WEB"
Global Const PROTOCOL_ONTRAK_TEXT   As String = "RELAY"
Global Const PROTOCOL_CENTRAL_TEXT  As String = "CENTRAL MONITOR"
Global Const PROTOCOL_MARQUIS_TEXT  As String = "MARQUIS"
Global Const PROTOCOL_PET_TEXT      As String = "TAP VERBOSE"
Global Const PROTOCOL_DIALOGIC_TEXT As String = "DIALOGIC"
Global Const PROTOCOL_REMOTE_TEXT   As String = "REMOTE"
Global Const PROTOCOL_TAP_IP_TEXT   As String = "TAP IP"
Global Const PROTOCOL_APOLLO_TEXT   As String = "APOLLO"
Global Const PROTOCOL_MOBILE_TEXT   As String = "MOBILE"
Global Const PROTOCOL_SDACT2_TEXT   As String = "SDACT2"
Global Const PROTOCOL_TAP2_TEXT     As String = "TAP2"


'Internet API Error Returns
Global Const INTERNET_ERROR_BASE = 12000
Global Const ERROR_INTERNET_OUT_OF_HANDLES = 12001
Global Const ERROR_INTERNET_TIMEOUT = 12002
Global Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Global Const ERROR_INTERNET_INTERNAL_ERROR = 12004
Global Const ERROR_INTERNET_INVALID_URL = 12005
Global Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = 12006
Global Const ERROR_INTERNET_NAME_NOT_RESOLVED = 12007
Global Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = 12008
Global Const ERROR_INTERNET_INVALID_OPTION = 12009
Global Const ERROR_INTERNET_BAD_OPTION_LENGTH = 12010
Global Const ERROR_INTERNET_OPTION_NOT_SETTABLE = 12011
Global Const ERROR_INTERNET_SHUTDOWN = 12012
Global Const ERROR_INTERNET_INCORRECT_USER_NAME = 12013
Global Const ERROR_INTERNET_INCORRECT_PASSWORD = 12014
Global Const ERROR_INTERNET_LOGIN_FAILURE = 12015
Global Const ERROR_INTERNET_INVALID_OPERATION = 12016
Global Const ERROR_INTERNET_OPERATION_CANCELLED = 12017
Global Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = 12018
Global Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = 12019
Global Const ERROR_INTERNET_NOT_PROXY_REQUEST = 12020
Global Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = 12021
Global Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = 12022
Global Const ERROR_INTERNET_NO_DIRECT_ACCESS = 12023
Global Const ERROR_INTERNET_NO_CONTEXT = 12024
Global Const ERROR_INTERNET_NO_CALLBACK = 12025
Global Const ERROR_INTERNET_REQUEST_PENDING = 12026
Global Const ERROR_INTERNET_INCORRECT_FORMAT = 12027
Global Const ERROR_INTERNET_ITEM_NOT_FOUND = 12028
Global Const ERROR_INTERNET_CANNOT_CONNECT = 12029
Global Const ERROR_INTERNET_CONNECTION_ABORTED = 12030
Global Const ERROR_INTERNET_CONNECTION_RESET = 12031
Global Const ERROR_INTERNET_FORCE_RETRY = 12032
Global Const ERROR_INTERNET_INVALID_PROXY_REQUEST = 12033
Global Const ERROR_INTERNET_NEED_UI = 12034
Global Const ERROR_INTERNET_HANDLE_EXISTS = 12036
Global Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = 12037
Global Const ERROR_INTERNET_SEC_CERT_CN_INVALID = 12038
Global Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = 12039
Global Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = 12040
Global Const ERROR_INTERNET_MIXED_SECURITY = 12041
Global Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = 12042
Global Const ERROR_INTERNET_POST_IS_NON_SECURE = 12043
Global Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = 12044
Global Const ERROR_INTERNET_INVALID_CA = 12045
Global Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = 12046
Global Const ERROR_INTERNET_ASYNC_THREAD_FAILED = 12047
Global Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = 12048
Global Const ERROR_INTERNET_DIALOG_PENDING = 12049
Global Const ERROR_INTERNET_RETRY_DIALOG = 12050
Global Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = 12052
Global Const ERROR_INTERNET_INSERT_CDROM = 12053
Global Const ERROR_INTERNET_FORTEZZA_LOGIN_NEEDED = 12054
Global Const ERROR_INTERNET_SEC_CERT_ERRORS = 12055
Global Const ERROR_INTERNET_SEC_CERT_NO_REV = 12056
Global Const ERROR_INTERNET_SEC_CERT_REV_FAILED = 12057
Global Const ERROR_HTTP_INVALID_SERVER_RESPONSE = 12152
Global Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = 12157
Global Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = 12158
Global Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = 12159
Global Const ERROR_INTERNET_DISCONNECTED = 12163
Global Const ERROR_INTERNET_SERVER_UNREACHABLE = 12164
Global Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = 12165
Global Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = 12166
Global Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = 12167
Global Const ERROR_INTERNET_SEC_INVALID_CERT = 12169
Global Const ERROR_INTERNET_SEC_CERT_REVOKED = 12170


Global Const ELAPSED_FORMAT = "d:hh:nn:ss"


Global Const LEVEL_FACTORY = 256
Global Const LEVEL_ADMIN = 128
Global Const LEVEL_SUPERVISOR = 32
Global Const LEVEL_USER = 1


Global Const SHIFT_DAY = 0
Global Const SHIFT_NIGHT = 1
Global Const SHIFT_GRAVE = 2

#If brookdale Then

  Global Const FACTORY_PWD = "brook2009"
  Global Const PRODUCT_NAME = "Brookdale TechConnect"
  Global Const G_REG_KEY = "Brookdale"
  Global Const G_REG_SUB_KEY = "TechConnect"
  Global Const G_REG_FILEMASK = "BROOKDALE"

  Global Const COMPANY_NAME = "ESCO Technologies, LLC"
  Global Const COMPANY_800 = "866-968-6225"
  Global Const COMPANY_FAX = "Fax: 513-674-8980"
  Global Const COMPANY_EMAIL = "Customer.Service@esco-tech.net"

#ElseIf esco Then
  Global Const FACTORY_PWD = "esco2009"
  Global Const PRODUCT_NAME = "Esco CareConnect"
  Global Const G_REG_KEY = "Esco"
  Global Const G_REG_SUB_KEY = "CareConnect"
  Global Const G_REG_FILEMASK = "ESCO"

  Global Const COMPANY_NAME = "ESCO Technologies, LLC"
  Global Const COMPANY_800 = "866-968-6225"
  Global Const COMPANY_FAX = "Fax: 513-674-8980"
  Global Const COMPANY_EMAIL = "Customer.Service@esco-tech.net"

#Else
  Global Const FACTORY_PWD = "heritage2005"
  Global Const PRODUCT_NAME = "Heritage MedCall Sentry Freedom II E-Call System"
  Global Const G_REG_KEY = "HeritageMedcall"
  Global Const G_REG_SUB_KEY = "Freedom2"
  Global Const G_REG_FILEMASK = "HERITAGE"

  Global Const COMPANY_NAME = "Heritage MedCall, Inc."
  Global Const COMPANY_800 = "813-221-1000 / 800-396-6157"
  Global Const COMPANY_FAX = "Fax: 813-223-1405"
  Global Const COMPANY_EMAIL = "email@heritagemedcall.com"


#End If

Global Const RPT_ROOM = 1
Global Const RPT_RES = 2
Global Const RPT_DEVICE = 3
Global Const RPT_EVENT = 4
Global Const RPT_ASSUR = 5
Global Const RPT_DEVHIST = 6   ' add/remove devices
Global Const RPT_RESHIST = 7   ' add/remove residents
Global Const RPT_EXCEPTION = 8  ' exception (alarm ack/clear time)
Global Const RPT_COUNT = 9     ' Generate Count of events
Global Const RPT_INV = 10      ' Generate Count of events

Global Const LV_WIDTH = 7665   ' standard lists width
Global Const LV_HEIGHT = 2985  ' standard lists height
Global Const BT_LEFT = 7725    ' standard buttons left

Global Const SCREEN_ALARM = 1
Global Const SCREEN_ALERT = 2
Global Const SCREEN_TROUBLE = 3
Global Const SCREEN_TAMPER = 4
Global Const SCREEN_BATTERY = 5

Global Const LOCATOR_WAIT_TIME = 5  ' sec before processing locators
Global Const SAME_EVENT_PERIOD = 15  ' sec between events for same event
Global Const TIMER_PERIOD = 20  ' ms ' clock tick period

Global Const BIT_NONE = 2 ^ -1
Global Const BIT_0 = 2 ^ 0
Global Const BIT_1 = 2 ^ 1
Global Const BIT_2 = 2 ^ 2
Global Const BIT_3 = 2 ^ 3
Global Const BIT_4 = 2 ^ 4
Global Const BIT_5 = 2 ^ 5
Global Const BIT_6 = 2 ^ 6
Global Const BIT_7 = 2 ^ 7

Global Const MAX_PCA_RESENDS = 4
