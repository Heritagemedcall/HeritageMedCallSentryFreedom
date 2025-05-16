VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Output Control Lines (CommSetLine)
Private Const LINE_BREAK = 1
Private Const LINE_DTR = 2
Private Const LINE_RTS = 3

' Input Control Lines  (CommGetLine)
Private Const LINE_CTS = &H10&
Private Const LINE_DSR = &H20&
Private Const LINE_RING = &H40&
Private Const LINE_RLSD = &H80&  ' same as CD
Private Const LINE_CD = &H80&

'-------------------------------------------------------------------------------
' System Constants
'-------------------------------------------------------------------------------
Private Const ERROR_IO_INCOMPLETE = 996&
Private Const ERROR_IO_PENDING = 997&
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Const OPEN_EXISTING = 3

' COMM Functions
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&

Private Const PURGE_TXABORT = &H1
Private Const PURGE_RXABORT = &H2
Private Const PURGE_TXCLEAR = &H4
Private Const PURGE_RXCLEAR = &H8

' COMM Escape Functions
Private Const SET_RTS = 3
Private Const CLR_RTS = 4
Private Const SET_DTR = 5
Private Const CLR_DTR = 6
Private Const SET_BREAK = 8
Private Const CLR_BREAK = 9


Private Type COMSTAT
  fBitFields As Long  ' See Comment in Win32API.Txt
  InQue As Long
  OutQue As Long
End Type

Private Type CommTimeOuts
  ReadIntervalTimeout As Long
  ReadTotalTimeoutMultiplier As Long
  ReadTotalTimeoutConstant As Long
  WriteTotalTimeoutMultiplier As Long
  WriteTotalTimeoutConstant As Long
End Type


Private Type DCB
  DCBlength As Long
  BaudRate As Long
  
  fBitFields As Long  ' See Comments in Win32API.Txt
    '  DWORD fBinary  :1;
    '  DWORD fParity  :1;
    '  DWORD fOutxCtsFlow  :1;
    '  DWORD fOutxDsrFlow  :1;
    '  DWORD fDtrControl  :2;
    '  DWORD fDsrSensitivity  :1;
    '  DWORD fTXContinueOnXoff  :1;
    '  DWORD fOutX  :1;
    '  DWORD fInX  :1;
    '  DWORD fErrorChar  :1;
    '  DWORD fNull  :1;
    '  DWORD fRtsControl  :2;
    '  DWORD fAbortOnError  :1;
    '  DWORD fDummy2  :17;
    
  wReserved As Integer
  XonLim As Integer
  XoffLim As Integer
  
  ByteSize As Byte
  Parity As Byte
  Stopbits As Byte
  
  XonChar As Byte
  XoffChar As Byte
  ErrorChar As Byte
  EofChar As Byte
  EvtChar As Byte
  
  wReserved1 As Integer  'Reserved; Do Not Use
End Type


'typedef struct _DCB {
'  DWORD DCBlength;
'  DWORD BaudRate;
'  DWORD fBinary  :1;
'  DWORD fParity  :1;
'  DWORD fOutxCtsFlow  :1;
'  DWORD fOutxDsrFlow  :1;
'  DWORD fDtrControl  :2;
'  DWORD fDsrSensitivity  :1;
'  DWORD fTXContinueOnXoff  :1;
'  DWORD fOutX  :1;
'  DWORD fInX  :1;
'  DWORD fErrorChar  :1;
'  DWORD fNull  :1;
'  DWORD fRtsControl  :2;
'  DWORD fAbortOnError  :1;
'  DWORD fDummy2  :17;

'  WORD  wReserved;

'  WORD  XonLim;
'  WORD  XoffLim;
'  BYTE  ByteSize;
'  BYTE  Parity;
'  BYTE  StopBits;

'  char  XonChar;
'  char  XoffChar;
'  char  ErrorChar;
'  char  EofChar;
'  char  EvtChar;

'  WORD  wReserved1;

'} DCB, *LPDCB;



Private Type OVERLAPPED
  Internal As Long
  InternalHigh As Long
  offset As Long
  OffsetHigh As Long
  hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type COMM_ERROR
  errorCode As Long
  FunctionName As String
  ErrorMessage As String
End Type

Private Type COMM_PORT
  Handle As Long
  PortOpen As Boolean
  DCB As DCB
End Type

Private CommOverlap  As Win32.OVERLAPPED
Private CommError    As COMM_ERROR
Private Port         As COMM_PORT

Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare Function ClearCommError Lib "kernel32" (ByVal hfile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function EscapeCommFunction Lib "kernel32" (ByVal nCid As Long, ByVal nFunc As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetCommModemStatus Lib "kernel32" (ByVal hfile As Long, lpModemStat As Long) As Long
Private Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hfile As Long, lpOverlapped As Win32.OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Private Declare Function PurgeComm Lib "kernel32" (ByVal hfile As Long, ByVal dwFlags As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hfile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, lpOverlapped As Win32.OVERLAPPED) As Long
Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hfile As Long, lpCommTimeouts As CommTimeOuts) As Long
Private Declare Function SetupComm Lib "kernel32" (ByVal hfile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hfile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Win32.OVERLAPPED) As Long

Public OutBufferSize As Long
Public InBufferSize As Long

Private mPortID     As Integer
Public Property Get PortID() As Integer
  ' gets com 1 thru n
  PortID = mPortID
End Property

Public Function PortOpen() As Boolean
  'returns status of port being open
  PortOpen = Port.PortOpen
End Function

Public Function GetSystemMessage(errorCode As Long) As String
  Dim intPos As Integer
  Dim strMessage As String, strMsgBuff As String * 256
  Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, errorCode, 0, strMsgBuff, 255, 0)

  intPos = InStr(1, strMsgBuff, vbNullChar)
  If intPos > 0 Then
    strMessage = Trim$(left$(strMsgBuff, intPos - 1))
  Else
    strMessage = Trim$(strMsgBuff)
  End If

  GetSystemMessage = strMessage

End Function

Function GetErrorCode() As Long
  Dim RC          As Long
  Dim errorflags  As Long
  Dim commstat    As COMSTAT

  'If the function succeeds, the return value is nonzero.
  RC = ClearCommError(Port.Handle, errorflags, commstat)
  If RC = 0 Then
    GetErrorCode = errorflags

    If CE_BREAK And errorflags Then  ' The hardware detected a break condition.
      ClearCommBreak Port.Handle
    End If

    If CE_FRAME And errorflags Then  ' The hardware detected a framing error.
      GetErrorCode = CE_FRAME  '8
    ElseIf CE_IOE And errorflags Then  ' An I/O error occurred during communications with the device.
      GetErrorCode = CE_IOE  ' 1024
    ElseIf CE_OVERRUN And errorflags Then  ' A character-buffer overrun has occurred. The next character is lost.
      GetErrorCode = CE_OVERRUN  '2
    ElseIf CE_RXOVER And errorflags Then  ' An input buffer overflow has occurred. There is either no room in the input buffer, or a character was received after the end-of-file (EOF) character.
      GetErrorCode = CE_RXOVER  '1
    ElseIf CE_RXPARITY And errorflags Then  ' The hardware detected a parity error.
      GetErrorCode = CE_RXPARITY  ' 4
    ElseIf CE_TXFULL And errorflags Then  '  The application tried to transmit a character, but the output buffer was full.
      GetErrorCode = CE_TXFULL  '256
    End If
    '    End If
    '    If GetErrorCode = 0 Then
    '      GetErrorCode = -1
    '    End If

    ' UNUSED
    'CE_MODE The requested mode is not supported, or the hFile parameter is invalid. If this value is specified, it is the only valid error.
    'CE_OOP Windows Me/98/95:  A parallel device signaled that it is out of paper.
    'CE_DNS Windows Me/98/95:  A parallel device is not selected.
    'CE_PTO Windows Me/98/95:  A time-out occurred on a parallel device.

  End If
End Function


'   PortID   - COM port ID. (1,2,3,4 ... 256)
'   Settings - Communication settings. Ex: "baud=1200 parity=E data=8 stop=1"
'   Error Code  - 0 = No Error.
Public Function CommOpen(PortID As Integer, ByVal Settings As String) As Long

  Dim CommTimeOuts    As CommTimeOuts

  If Settings = "" Then
    Settings = "baud=9600 parity=N data=8 stop=1"
  End If
  On Error GoTo Error_Handler
  ' See if port already in use.
  If Port.PortOpen Then
    CommOpen = -1
    CommError.errorCode = 0
    CommError.FunctionName = "CommOpen"
    CommError.ErrorMessage = "Port in use"
    Exit Function
  End If
  ' Open serial port.

  ' note must use "\\.\COM" to access ports above 16
   'Port.Handle = CreateFile("\\.\COM" & PortID, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  Port.Handle = CreateFile("\\.\COM" & PortID, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
  If Port.Handle = -1 Then
    CommOpen = SetCommError("CommOpen (CreateFile)")
    CommError.ErrorMessage = "Comm Port " & PortID & " Unavailable"
    
    Port.PortOpen = False
    PortID = 0
    Exit Function
  End If

  Port.PortOpen = True

  ' Setup device buffers (1K each).
  If 0 = SetupComm(Port.Handle, InBufferSize, OutBufferSize) Then
    CommOpen = SetCommError("CommOpen (SetupComm)")
    Port.PortOpen = False
    Exit Function
  End If

  ' Purge buffers.
  '  If 0 = PurgeComm(Port.Handle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR) Then
  '    CommOpen = SetCommError("CommOpen (PurgeComm)")
  '    Exit Function
  '  End If

  ' Set serial port timeouts.

  CommTimeOuts.ReadIntervalTimeout = &HFFFFFFFF '; -1
  CommTimeOuts.ReadTotalTimeoutMultiplier = 0
  'CommTimeOuts.ReadTotalTimeoutConstant = 1000
  CommTimeOuts.ReadTotalTimeoutConstant = 0
  CommTimeOuts.WriteTotalTimeoutMultiplier = 0
  'CommTimeOuts.WriteTotalTimeoutConstant = 1000
  CommTimeOuts.WriteTotalTimeoutConstant = 0


  If 0 = SetCommTimeouts(Port.Handle, CommTimeOuts) Then
    CommOpen = SetCommError("CommOpen (SetCommTimeouts)")
    Exit Function
  End If

  ' Get the current state (DCB).
  If 0 = GetCommState(Port.Handle, Port.DCB) Then
    CommOpen = SetCommError("CommOpen (GetCommState)")
    Port.PortOpen = False
    Exit Function
  End If

  ' Modify the DCB to reflect the desired settings.
  
  Dim tDCB As DCB
  If 0 = BuildCommDCB(Settings, tDCB) Then
  
  End If
  
  
  
  If 0 = BuildCommDCB(Settings, Port.DCB) Then
    CommOpen = SetCommError("CommOpen (BuildCommDCB)")
    Port.PortOpen = False
    Exit Function
  End If


'  If Port.DCB.StopBits = 0 Then
'    Port.DCB.StopBits = 1
'    Port.DCB.Parity = 1
'  End If

  
  ' Set the new state.
  If 0 = SetCommState(Port.Handle, Port.DCB) Then
    CommOpen = SetCommError("CommOpen (SetCommState)")
    Port.PortOpen = False
    Exit Function
  End If
  mPortID = PortID

Error_Handler:
  If Err.Number <> 0 Then
    CommOpen = Err.Number
    SetError Err.Number, "CommOpen", Err.Description
  End If
End Function
Private Sub SetError(errorCode, FunctionName, Description)
  CommError.errorCode = errorCode
  CommError.FunctionName = FunctionName
  CommError.ErrorMessage = Description
End Sub


Private Function SetCommError(FunctionName As String) As Long
  CommError.errorCode = Err.LastDllError
  CommError.FunctionName = FunctionName
  CommError.ErrorMessage = GetSystemMessage(CommError.errorCode)
  SetCommError = CommError.errorCode
End Function


Private Function SetCommErrorEx(FunctionName As String, lngHnd As Long) As Long
  Dim errorflags  As Long
  Dim commstat    As COMSTAT
  Dim RC          As Long

  CommError.errorCode = GetLastError
  CommError.FunctionName = FunctionName
  CommError.ErrorMessage = GetSystemMessage(CommError.errorCode)
  RC = ClearCommError(lngHnd, errorflags, commstat)
  CommError.ErrorMessage = CommError.ErrorMessage + "  COMM Error Flags = " & Hex$(errorflags)
  SetCommErrorEx = CommError.errorCode


End Function

Public Function CommSet(ByVal Settings As String) As Long

  On Error GoTo Err_Handler

  If 0 = GetCommState(Port.Handle, Port.DCB) Then
    CommSet = SetCommError("CommSet (GetCommState)")
    Exit Function
  End If

  If 0 = BuildCommDCB(Settings, Port.DCB) Then
    CommSet = SetCommError("CommSet (BuildCommDCB)")
    Exit Function
  End If


  If 0 = SetCommState(Port.Handle, Port.DCB) Then
    CommSet = SetCommError("CommSet (SetCommState)")
    Exit Function
  End If

Err_Handler:
  If Err.Number <> 0 Then
    CommSet = Err.Number
    SetError Err.Number, "CommSet", Err.Description
  End If
End Function

Public Function CommClose() As Long

  On Error GoTo Err_Handler

  If Port.PortOpen Then
    If 0 = CloseHandle(Port.Handle) Then
      CommClose = SetCommError("CommClose (CloseHandle)")
      Exit Function
    End If
  End If
  Port.PortOpen = False

Err_Handler:
  If Err.Number <> 0 Then
    CommClose = Err.Number
    SetError Err.Number, "CommClose", Err.Description
  End If

End Function

Public Function CommFlush() As Long
  On Error GoTo Err_Handler
  'If 0 = PurgeComm(Port.Handle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR) Then
  ' we'll just clear input
  If Port.Handle = 0 Then Exit Function
  
  If 0 = PurgeComm(Port.Handle, PURGE_RXABORT Or PURGE_RXCLEAR) Then
    CommFlush = SetCommError("CommFlush (PurgeComm)")
    Exit Function
  End If

Err_Handler:
  If Err.Number <> 0 Then
    CommFlush = Err.Number
    SetError Err.Number, "CommFlush", Err.Description
  End If
End Function

Public Function CommInput(Optional ByVal numbytes As Long) As String
  Dim Buffer As String
  Call CommRead(Buffer, numbytes)
  CommInput = Buffer
End Function
Public Function CommReadByte(ReadBuffer As Byte) As Long
  Dim bytesread As Long
  Call Win32.ReadFile(Port.Handle, ReadBuffer, 1, bytesread, CommOverlap)
  CommReadByte = bytesread
End Function


'   Buffer   - Data buffer.
'   Size     - requested number of bytes to be read.
Public Function CommRead(Buffer As String, Size As Long) As Long

  Dim Status            As Long
  Dim ReadSize          As Long
  Dim bytesread         As Long
  Dim ReadBuffer        As String
  Dim errorflags        As Long
  Dim commstat          As COMSTAT

10          ReadBuffer = String(1024, vbNullChar)


20          On Error GoTo Err_Handler

30          Buffer = ""
40          bytesread = 0
  ' Clear any previous errors and get current status.

50          If 0 = ClearCommError(Port.Handle, errorflags, commstat) Then
60            bytesread = -1
70            CommRead = SetCommError("CommRead (ClearCommError)")
80            Exit Function
90          End If


100         If commstat.InQue > 0 Then
110           If Size = 0 Then  ' get all data
120             ReadSize = commstat.InQue
130           Else
140             If Size > commstat.InQue Then
150               ReadSize = commstat.InQue
160             Else
170               ReadSize = Size
180             End If
190           End If
200           ReadBuffer = String(ReadSize, vbNullChar)
210         Else
220           ReadSize = 0
230         End If


240         If ReadSize > 0 Then
250           If 0 = ReadFile(Port.Handle, ReadBuffer, ReadSize, bytesread, CommOverlap) Then
260             Status = GetLastError
270             If ERROR_IO_PENDING = GetLastError() Then
  ' Wait for read to complete.
  ' This function will timeout according to the
  ' COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
  ' Every time it times out, check for port errors.
  ' Loop until operation is complete.
280               While GetOverlappedResult(Port.Handle, CommOverlap, bytesread, True) = 0
290                 If ERROR_IO_INCOMPLETE = GetLastError() Then
300                   bytesread = -1
310                   CommRead = SetCommErrorEx("CommRead (GetOverlappedResult)", Port.Handle)
320    'Debug.Print "Error: CommRead (GetOverlappedResult) 320"
330                   Exit Function
340                 End If
350               Wend
360             Else
  ' Some other error occurred.
370               bytesread = -1
380               CommRead = SetCommErrorEx("CommRead (ReadFile)", Port.Handle)
390               Exit Function
400             End If
410           End If
420           Buffer = left$(ReadBuffer, bytesread)
430         End If

440         CommRead = bytesread

Err_Handler:
450         If Err.Number <> 0 Then
460           Debug.Print "Comm Read Error " & Erl
470           CommRead = -1
480           SetError Err.Number, "CommRead", Err.Description
490         End If
End Function

Public Function CommWriteByte(Buffer As Byte) As Long
  Dim WriteStatus  As Long
  Dim WriteSize    As Long

  WriteStatus = Win32.WriteFile(Port.Handle, Buffer, 1, WriteSize, CommOverlap)
  CommWriteByte = WriteSize

Err_Handler:
  If Err.Number <> 0 Then
    SetError Err.Number, "CommWriteByte", Err.Description
  End If

End Function
Public Function CommWriteByteArray(Buffer() As Byte, ByVal numbytes As Long) As Long
  Dim WriteStatus  As Long
  Dim WriteSize    As Long

'  Dim j As Integer
'  Dim s As String
'  For j = 0 To NumBytes - 1
'    s = s & Right("00" & Hex(buffer(j)), 2) & " "
'  Next
'  Trace s

  WriteStatus = Win32.WriteFile(Port.Handle, Buffer(0), numbytes, WriteSize, CommOverlap)
  CommWriteByteArray = WriteSize

Err_Handler:
  If Err.Number <> 0 Then
    SetError Err.Number, "CommWriteByte", Err.Description
  End If

End Function

Public Function CommWrite(Buffer As String) As Long

  Dim WriteSize    As Long

  On Error GoTo Err_Handler


  If 0 = WriteFile(Port.Handle, Buffer, Len(Buffer), WriteSize, CommOverlap) Then
    ' Note that normally the following code will not execute because the driver
    ' caches write operations. Small I/O requests (up to several thousand bytes)
    ' will normally be accepted immediately and WriteFile will return true even
    ' though an overlapped operation was specified.
    ' let's assume success

    '    Status = GetLastError
    '    If Status = 0 Then
    '      GoTo Routine_Exit
    '    ElseIf Status = ERROR_IO_PENDING Then
    '      ' We should wait for the completion of the write operation so we know
    '      ' if it worked or not.
    '      '
    '      ' This is only one way to do this. It might be beneficial to place the
    '      ' writing operation in a separate thread so that blocking on completion
    '      ' will not negatively affect the responsiveness of the UI.
    '      '
    '      ' If the write takes long enough to complete, this function will timeout
    '      ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
    '      ' At that time we can check for errors and then wait some more.
    '
    '      ' Loop until operation is complete.
    '      Do While GetOverlappedResult(Port.Handle, CommOverlap, WriteSize, True) = 0
    '        Status = GetLastError
    '        If Status <> ERROR_IO_INCOMPLETE Then
    '          Status = SetCommErrorEx("CommWrite (GetOverlappedResult)", Port.Handle)
    '          GoTo Routine_Exit
    '        End If
    '      Loop
    '    Else
    '      ' Some other error occurred.
    '      WriteSize = -1
    '
    '      Status = SetCommErrorEx("CommWrite (WriteFile)", Port.Handle)
    '      GoTo Routine_Exit
    '
    '    End If

  End If

  CommWrite = WriteSize
Err_Handler:
  If Err.Number <> 0 Then
    SetError Err.Number, "CommWrite", Err.Description
  End If

End Function
Public Property Get CTS() As Boolean
  Dim RC            As Long
  Dim ModemStatus   As Long
  RC = GetCommModemStatus(Port.Handle, ModemStatus)
  CTS = (ModemStatus And LINE_CTS) = LINE_CTS
End Property
Public Property Get DSR() As Boolean
  Dim RC            As Long
  Dim ModemStatus   As Long
  RC = GetCommModemStatus(Port.Handle, ModemStatus)
  DSR = (ModemStatus And LINE_DSR) = LINE_DSR
End Property
Public Property Get RING() As Boolean
  Dim RC            As Long
  Dim ModemStatus   As Long
  On Error Resume Next
  RC = GetCommModemStatus(Port.Handle, ModemStatus)
  RING = (ModemStatus And LINE_RING) = LINE_RING
End Property
Public Property Get RLSD() As Boolean
  Dim RC            As Long
  Dim ModemStatus   As Long
  On Error Resume Next
  RC = GetCommModemStatus(Port.Handle, ModemStatus)
  RLSD = (ModemStatus And LINE_RLSD) = LINE_RLSD
End Property
Public Property Get CD() As Boolean
  Dim RC            As Long
  Dim ModemStatus   As Long
  On Error Resume Next
  RC = GetCommModemStatus(Port.Handle, ModemStatus)
  CD = (ModemStatus And LINE_CD) = LINE_CD
End Property

Public Sub SetBreak()
  Dim RC As Long
  RC = EscapeCommFunction(Port.Handle, SET_BREAK)
End Sub

Public Sub ClearBreak()
  Dim RC As Long
  RC = EscapeCommFunction(Port.Handle, CLR_BREAK)
End Sub

Public Sub SetDTR()
  Dim RC As Long
  RC = EscapeCommFunction(Port.Handle, SET_DTR)
End Sub
Public Sub ClearDTR()
  Dim RC As Long
  RC = EscapeCommFunction(Port.Handle, CLR_DTR)
End Sub
Public Sub SetRTS()
  Dim RC As Long
  RC = EscapeCommFunction(Port.Handle, SET_RTS)
End Sub

Public Sub ClearRTS()
  Dim RC As Long
  RC = EscapeCommFunction(Port.Handle, CLR_RTS)
End Sub
Public Function OutBufferCount() As Long
  Dim errorflags As Long
  Dim commstat As COMSTAT
  If Port.Handle <> 0 Then
    Call ClearCommError(Port.Handle, errorflags, commstat)
  End If
  OutBufferCount = commstat.OutQue
End Function

Public Function inbuffercount() As Long
  Dim errorflags As Long
  Dim commstat As COMSTAT
  If Port.Handle <> 0 Then
    Call ClearCommError(Port.Handle, errorflags, commstat)
  End If
  inbuffercount = commstat.InQue
End Function

Public Function CommGetError() As String
  CommGetError = CommError.ErrorMessage
End Function

Private Sub Class_Initialize()
  InBufferSize = 2048
  OutBufferSize = 1024
End Sub

Public Function ReadPortSettings() As String


  Dim NewDCB As DCB
  NewDCB.DCBlength = Len(NewDCB)
  If 0 = GetCommState(Port.Handle, NewDCB) Then
    ReadPortSettings = "Unable to read port settings"
  Else
   
'  DCBlength As Long
'  BaudRate As Long
'  fBitFields As Long  ' See Comments in Win32API.Txt
'  wReserved As Integer
'  XonLim As Integer
'  XoffLim As Integer
'  ByteSize As Byte
'  Parity As Byte
'  StopBits As Byte
'  XonChar As Byte
'  XoffChar As Byte
'  ErrorChar As Byte
'  EofChar As Byte
'  EvtChar As Byte
'  wReserved1 As Integer  'Reserved; Do Not Use
   
    ReadPortSettings = "Baud " & NewDCB.BaudRate & " " & NewDCB.ByteSize & " " & NewDCB.Parity & " " & NewDCB.Stopbits
  End If
  
End Function
