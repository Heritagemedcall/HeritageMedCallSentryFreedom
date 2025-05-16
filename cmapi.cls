VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Const MAPI_UNREAD_ONLY = &H20
Const MAPI_GUARANTEE_FIFO = &H100

Private Type MAPIMessage
  lReserved As Long
  Subject As String
  MessageText As String
  MessageType As String
  DateReceived As String
  ConversationID As String
  Flags As Long
  RecipCount As Long
  FileCount As Long
End Type

Private Type MapiRecip
  lReserved As Long
  RecipClass As Long
  Name As String
  Address As String
  EIDSize As Long
  EntryID As String
End Type

'MapiFile structure, holds information about file attachments
Private Type MapiFile
  lReserved As Long
  Flags As Long
  Position As Long
  PathName As String
  Filename As String
  FileType As String
End Type



Private mSessionID    As Long
Private mParentHwnd   As Long
Private mUserName     As String
Private mPassword     As String
Private mShowDialogs  As Long


'Error message API
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

'MAPI API
Private Declare Function MAPILogon Lib "MAPI32.DLL" (ByVal lUIParam As Long, ByVal user As String, ByVal password As String, ByVal lFlags As Long, ByVal lReserved As Long, lSession As Long) As Long
Private Declare Function MAPILogoff Lib "MAPI32.DLL" (ByVal lSession As Long, ByVal lUIParam As Long, ByVal lFlags As Long, ByVal lReserved As Long) As Long

Private Declare Function MAPISaveMail Lib "MAPI32.DLL" Alias "BMAPISaveMail" (ByVal Session As Long, ByVal UIParam As Long, Message As MAPIMessage, Recipient As MapiRecip, File As MapiFile, ByVal Reserved As Long, MsgID As String) As Long
Private Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" (ByVal Session As Long, ByVal UIParam As Long, Message As MAPIMessage, Recipient() As MapiRecip, File() As MapiFile, ByVal Flags As Long, ByVal Reserved As Long) As Long

Private Declare Function MAPIReadMail Lib "MAPI32.DLL" Alias "BMAPIReadMail" (lMsg As Long, nRecipients As Long, nFiles As Long, ByVal Session As Long, ByVal UIParam As Long, MessageID As String, ByVal Flag As Long, ByVal Reserved As Long) As Long
Private Declare Function MAPIGetReadMail Lib "MAPI32.DLL" Alias "BMAPIGetReadMail" (ByVal lMsg As Long, Message As MAPIMessage, Recip() As MapiRecip, File() As MapiFile, Originator As MapiRecip) As Long
Private Declare Function MAPIFindNext Lib "MAPI32.DLL" Alias "BMAPIFindNext" (ByVal lSession As Long, ByVal lUIParam As Long, MsgType As String, SeedMsgID As String, ByVal Flag As Long, ByVal lReserved As Long, MsgID As String) As Long

Private Declare Function MAPISendDocuments Lib "MAPI32.DLL" (ByVal lUIParam As Long, ByVal DelimStr As String, ByVal FilePaths As String, ByVal FileNames As String, ByVal lReserved As Long) As Long
Private Declare Function MAPIDeleteMail Lib "MAPI32.DLL" (ByVal lSession As Long, ByVal lUIParam As Long, ByVal MsgID As String, ByVal lFlags As Long, ByVal lReserved As Long) As Long
Private Declare Function MAPIAddress Lib "MAPI32.DLL" (ByVal lSession As Long, ByVal lUIParam As Long, ByVal Caption As String, ByVal nEditFields As Long, ByVal Label As String, ByVal nRecipients As Long, Recip() As MapiRecip, ByVal lFlags As Long, ByVal lReserved As Long) As Long
Private Declare Function MAPIGetAddress Lib "MAPI32.DLL" (ByVal lInfo As Long, ByVal nRecipients As Long, Recipients() As MapiRecip) As Long
Private Declare Function MAPIDetails Lib "MAPI32.DLL" Alias "BMAPIDetails" (ByVal lSession As Long, ByVal lUIParam As Long, Recipient As MapiRecip, ByVal lFlags As Long, ByVal lReserved As Long) As Long
Private Declare Function MAPIResolveName Lib "MAPI32.DLL" Alias "BMAPIResolveName" (ByVal lSession As Long, ByVal lUIParam As Long, ByVal username As String, ByVal lFlags As Long, ByVal lReserved As Long, Recipient As MapiRecip) As Long

Private Const SUCCESS_SUCCESS = 0
Private Const MAPI_USER_ABORT = 1
Private Const MAPI_E_USER_ABORT = MAPI_USER_ABORT
Private Const MAPI_E_FAILURE = 2
Private Const MAPI_E_LOGIN_FAILURE = 3
Private Const MAPI_E_LOGON_FAILURE = MAPI_E_LOGIN_FAILURE
Private Const MAPI_E_DISK_FULL = 4
Private Const MAPI_E_INSUFFICIENT_MEMORY = 5
Private Const MAPI_E_BLK_TOO_SMALL = 6
Private Const MAPI_E_TOO_MANY_SESSIONS = 8
Private Const MAPI_E_TOO_MANY_FILES = 9
Private Const MAPI_E_TOO_MANY_RECIPIENTS = 10
Private Const MAPI_E_ATTACHMENT_NOT_FOUND = 11
Private Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 12
Private Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 13
Private Const MAPI_E_UNKNOWN_RECIPIENT = 14
Private Const MAPI_E_BAD_RECIPTYPE = 15
Private Const MAPI_E_NO_MESSAGES = 16
Private Const MAPI_E_INVALID_MESSAGE = 17
Private Const MAPI_E_TEXT_TOO_LARGE = 18
Private Const MAPI_E_INVALID_SESSION = 19
Private Const MAPI_E_TYPE_NOT_SUPPORTED = 20
Private Const MAPI_E_AMBIGUOUS_RECIPIENT = 21
Private Const MAPI_E_AMBIG_RECIP = MAPI_E_AMBIGUOUS_RECIPIENT
Private Const MAPI_E_MESSAGE_IN_USE = 22
Private Const MAPI_E_NETWORK_FAILURE = 23
Private Const MAPI_E_INVALID_EDITFIELDS = 24
Private Const MAPI_E_INVALID_RECIPS = 25
Private Const MAPI_E_NOT_SUPPORTED = 26
Private Const MAPI_E_NO_LIBRARY = 999
Private Const MAPI_E_INVALID_PARAMETER = 998

Private Const MAPI_ORIG = 0
Private Const MAPI_TO = 1
Private Const MAPI_CC = 2
Private Const MAPI_BCC = 3

'Flags for MAPILogon
Private Const MAPI_LOGON_UI = &H1&
Private Const MAPI_NEW_SESSION = &H2&
Private Const MAPI_FORCE_DOWNLOAD = &H1000&

'Flags for MAPILogoff
Private Const MAPI_LOGOFF_SHARED = &H1
Private Const MAPI_LOGOFF_UI = &H2

'Flags for MAPISendMail
Private Const MAPI_DIALOG = &H8
Private Const MAPI_NODIALOG = &H1

'Flags for ReadMail
Private Const MAPI_ENVELOPE_ONLY = &H40
Private Const MAPI_PEEK = &H80
Private Const MAPI_BODY_AS_FILE = &H200
Private Const MAPI_SUPPRESS_ATTACH = &H800

'Flags for MAPIDetails
Private Const MAPI_AB_NOMODIFY = &H400

'Flags for Attachments
Private Const MAPI_OLE = &H1
Private Const MAPI_OLE_STATIC = &H2

'Flags for MapiMessage
Private Const MAPI_UNREAD = &H1
Private Const MAPI_RECEIPT_REQUESTED = &H2
Private Const MAPI_SENT = &H4


Public Function Send(ByVal Address As String, ByVal Subject As String, ByVal Body As String) As Long  ' 0 on success
  Dim Message   As MAPIMessage
  Dim Recipient(1) As MapiRecip
  Dim FileInfo(1)  As MapiFile

  Message.Subject = Subject
  Message.MessageText = Body
  Message.FileCount = 0
  Message.RecipCount = 1
  
  
  Recipient(1).Address = Address
  Recipient(1).Name = "Sentry Alert"

  Send = MAPISendMail(mSessionID, mParentHwnd, Message, Recipient, FileInfo, &H1&, 0&)
End Function


'Private Function ReadMail(Message As MAPIMessage, Originator As MapiRecip, Recipients() As MapiRecip, Attachments() As MapiFile, Optional ByRef MessageResumeID As String, Optional UnRead As Long = MAPI_UNREAD_ONLY) As Long
'  Dim lMessageID As Long
'  Dim lNumFiles As Long, lNumRecips As Long
'
'  If Len(MessageResumeID) = 0 Then
'    'First call to routine, create buffer
'    MessageResumeID = Space(256)
'    ReadMail = MAPIFindNext(mSessionID, mParentHwnd, "", "", UnRead, 0&, MessageResumeID)
'  Else
'    'User specified a MessageResumeID
'    ReadMail = MAPIFindNext(mSessionID, mParentHwnd, "", MessageResumeID, UnRead, 0&, MessageResumeID)
'  End If
'
'  If (ReadMail = SUCCESS_SUCCESS) Then
'    ReadMail = MAPIReadMail(lMessageID, lNumRecips, lNumFiles, mSessionID, mParentHwnd, MessageResumeID, MAPI_PEEK, 0&)
'    If (ReadMail = SUCCESS_SUCCESS) Then
'      'Read mail into types
'      ReadMail = MAPIGetReadMail(lMessageID, Message, Recipients, Attachments, Originator)
'    End If
'  End If
'End Function

Function Logon() As String
  Dim RC As Long

  'On Error Resume Next
  If mSessionID Then
    'End existing session
    LogOff
    mSessionID = 0
  End If

  RC = MAPILogon(mParentHwnd, mUserName, mPassword, MAPI_LOGON_UI, 0&, mSessionID)
  If RC <> 0 Then
  Logon = ErrorDescription(RC)
  End If
End Function

Property Get password() As String
  password = mPassword
End Property

Property Let password(Value As String)
  mPassword = Value
End Property
Property Get username() As String
  password = mUserName
End Property
Property Let username(Value As String)
  mUserName = Value
End Property
Function LogOff() As Long
  If mSessionID Then
    LogOff = MAPILogoff(mSessionID, mParentHwnd, 0&, 0&)
    mSessionID = 0
  End If
End Function
'Private Function ResolveName(ByRef Recipient As MapiRecip) As Long
'  ResolveName = MAPIResolveName(mSessionID, mParentHwnd, Recipient.Name, mShowDialogs, 0&, Recipient)
'End Function
Public Function ShowSendFileDialog(DocumentList As String, DocumentNamesList As String, Optional Delimeter As String = ";") As Long
  ShowSendFileDialog = MAPISendDocuments(mParentHwnd, Delimeter, DocumentList, DocumentNamesList, 0&)
End Function
'Private Function SendMail(Message As MAPIMessage, Recipient() As MapiRecip, tFile() As MapiFile) As Long
'  Dim ThisRecip As Long
'
'  'Resolve the recipients before sending mail
'  For ThisRecip = LBound(Recipient) To UBound(Recipient)
'    SendMail = MAPIResolveName(mSessionID, mParentHwnd, Recipient(ThisRecip).Name, 0, 0, Recipient(ThisRecip))
'  Next
'  'MAPI_DIALOG if mShowDialogs <>  0
'  SendMail = MAPISendMail(mSessionID, mParentHwnd, Message, Recipient, tFile, mShowDialogs, 0&)
'  If SendMail Then
'    Debug.Print ErrorDescription(SendMail)
'  End If
'End Function
Property Get ShowDialogs() As Long
  ShowDialogs = mShowDialogs
End Property

Property Let ShowDialogs(Value As Long)
  mShowDialogs = Value
End Property
Property Get ParentHwnd() As Long
  ParentHwnd = mParentHwnd
End Property
Property Let ParentHwnd(Value As Long)
  mParentHwnd = Value
End Property
Private Sub Class_Initialize()
'mShowDialogs = MAPI_DIALOG
'mParentHwnd = GetActiveWindow      'Seed parent window handle
End Sub
Private Sub Class_Terminate()
  LogOff
End Sub
Function DeleteMail(sMessageID As String) As Long
  DeleteMail = MAPIDeleteMail(mSessionID, mParentHwnd, sMessageID, mShowDialogs, 0&)
End Function
Function ErrorDescription(ByVal ErrorNumber As Long) As String
  Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_SYSTEM = &H1000
  Const NERR_BASE = 2100, MAX_NERR = NERR_BASE + 899
  Const LOAD_LIBRARY_AS_DATAFILE = &H2

  Dim RC      As String
  Dim Flags   As Long
  Dim hModule As Long
  Dim Ret     As Long

  hModule = 0
  RC = Space$(256)
  Flags = FORMAT_MESSAGE_FROM_SYSTEM

  'If Ret is in the network range, load the message source
  If (ErrorNumber >= NERR_BASE And ErrorNumber <= MAX_NERR) Then
    hModule = LoadLibraryEx("netmsg.dll", 0&, LOAD_LIBRARY_AS_DATAFILE)
    If (hModule <> 0) Then
      Flags = Flags Or FORMAT_MESSAGE_FROM_HMODULE
    End If
  End If

  'Call FormatMessage to allow for message text to be acquired
  'from the system or the supplied module handle.
  Ret = FormatMessage(Flags, hModule, ErrorNumber, 0&, RC, 256&, 0&)

  If (hModule <> 0) Then
    'Unloaded message source
    FreeLibrary hModule
  End If

  ErrorDescription = "ERROR: " & ErrorNumber & " - " & RC

  'Clean message
  Ret = InStr(1, ErrorDescription, vbNullChar)
  If Ret Then
    ErrorDescription = left$(ErrorDescription, Ret - 1)
  End If
  Ret = InStr(1, ErrorDescription, vbNewLine)
  If Ret Then
    ErrorDescription = left$(ErrorDescription, Ret - 1)
  End If
End Function



' Place the following code in a standard module:

'Option Explicit
'
''---------------------Public structures and constants-----------------
''Flags for MAPIFindNext
'Public Const MAPI_UNREAD_ONLY = &H20
'Public Const MAPI_GUARANTEE_FIFO = &H100
'
'
''MAPI message holds information about a Message
'Public Type MAPIMessage
'    lReserved As Long
'    Subject As String
'    MessageText As String
'    MessageType As String
'    DateReceived As String
'    ConversationID As String
'    Flags As Long
'    RecipCount As Long
'    FileCount As Long
'End Type
'
''MAPIRecip structure, holds information about a message
''originator or recipient
'Public Type MapiRecip
'    lReserved As Long
'    RecipClass As Long
'    Name As String
'    Address As String
'    EIDSize As Long
'    EntryID As String
'End Type
'
''MapiFile structure, holds information about file attachments
'Public Type MapiFile
'    lReserved As Long
'    Flags As Long
'    Position As Long
'    PathName As String
'    FileName As String
'    FileType As String
'End Type
'
'
'
''Demostration routine (for Excel 2000 and VB)
'
'Sub Test()
'    Dim Mail As New MAPI    'Assumes you called the class MAPI
'    Dim Message As MAPIMessage
'    Dim aRecipients() As MapiRecip
'    Dim tOriginator As MapiRecip
'    Dim atMapiFile(1 To 1) As MapiFile
'    Dim lRetVal As Long
'    Dim sMesageResumeID As String
'
''---Logon to mail account
'    Mail.UserName = ""          'Enter Username
'    Mail.Password = ""          'Enter Password
'    Debug.Print Mail.Logon
'
'    'Get a mail address (resolve a name)
'    ReDim aRecipients(1 To 1)
'    aRecipients(1).Name = InputBox("Please enter an email address to find...")
'    If Mail.ResolveName(aRecipients(1)) = 0 Then
'        MsgBox "Found address: " & aRecipients(1).Address
'    End If
'
'    'Send a message
'    Message.MessageText = "Test"
'    Message.FileCount = 1          'Send a file
'    Message.RecipCount = 2         'Send to two email addresses
'    Message.Subject = "Test Mail"
'
'    atMapiFile(1).FileName = "Autoexec"
'    atMapiFile(1).PathName = "C:\test.txt"
'
'    ReDim aRecipients(1 To 2)
'    aRecipients(1).Name = "Andrew Baker"
'    aRecipients(1).Address = "someone@hotmail.com"
'
'    aRecipients(2).Name = "Andrew Baker"
'    aRecipients(2).Address = "someone@somewhere.com"
'
'    Mail.ShowDialogs = 1    'Don't show dialogs
'
'    Debug.Print "Send Mail Result: " & Mail.SendMail(Message, aRecipients, atMapiFile)
'
''---Get all new mail
'    sMesageResumeID = ""
'    lRetVal = Mail.ReadMail(Message, tOriginator, aRecipients, atMapiFile, sMesageResumeID)
'    Do While lRetVal = 0
'        Debug.Print "--------------------------------"
'        Debug.Print "SUBJECT: " & Message.Subject
'        Debug.Print "TEXT: " & Message.MessageText
'        Debug.Print "DATE: " & Message.DateReceived
'        Debug.Print "FROM: " & tOriginator.Name & " {" & tOriginator.Address & "}"
'        Debug.Print "--------------------------------"
'        lRetVal = Mail.ReadMail(Message, tOriginator, aRecipients, atMapiFile, sMesageResumeID)
'        If MsgBox("Delete new mail :" & Message.Subject, vbQuestion + vbYesNo) = vbYes Then
'            Mail.DeleteMail sMesageResumeID
'        End If
'    Loop
'    'Logoff
'    Mail.LogOff
'End Sub
'

