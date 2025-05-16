Attribute VB_Name = "ModComm"
Option Explicit

Global WirelessPort As cComm

Global Const BYTE_STX = &H2
Global Const BYTE_ETX = &H3
Global Const BYTE_EOT = &H4
Global Const BYTE_ENQ = &H5
Global Const BYTE_ACK = &H6
Global Const BYTE_CR = &HD
Global Const BYTE_NAK = 15
Global Const BYTE_ESC = &H1B
Global Const BYTE_US = &H1F
Global Const BYTE_LF = &HA
Global Const BYTE_ETB = &H17
Global Const BYTE_SUB = &H1A ' to prefix char less than 0x20 , then add 0x40 to control char


Global CHAR_NUL       As String
Global CHAR_SOH       As String
Global CHAR_STX       As String
Global CHAR_ETX       As String
Global CHAR_EOT       As String
Global CHAR_ENQ       As String
Global CHAR_ACK       As String
Global CHAR_BEL       As String
Global CHAR_BS        As String
Global CHAR_HT        As String
Global CHAR_LF        As String
Global CHAR_VT        As String
Global CHAR_FF        As String
Global CHAR_CR        As String
Global CHAR_SO        As String
Global CHAR_SI        As String
Global CHAR_XOFF      As String
Global CHAR_XON       As String
Global CHAR_NAK       As String
Global CHAR_ETB       As String
Global CHAR_SUB       As String
Global CHAR_ESC       As String
Global CHAR_RS        As String
Global CHAR_US        As String
Global CHAR_DEL       As String
'
Global CHAR_SUB_CR    As String
Global CHAR_SUB_LF    As String
Global CHAR_CRLF      As String


Public Enum TAP_STATUS
  TAP_TIMEOUT = -1
  TAP_WAITING = 0
  TAP_ATTENTION
  TAP_LOGON
  TAP_ACCEPT
  TAP_HAS_STX
  TAP_HAS_ID
  TAP_HAS_MSG
  TAP_HAS_ETX
  TAP_HAS_CHKSUM
  TAP_IS_VALID
  TAP_SENT
  TAP_EOT
  TAP_DONE
  TAP_ERROR
End Enum


Public Function InitComm(SerialPort As cComm, ByVal PortID As Integer, ByVal Settings As String) As Long
        
                
        
        ' 9600,n,8,1 default
10       On Error GoTo InitComm_Error

20      If Settings = "" Then
30        Settings = "baud=9600 parity=N data=8 stop=1"
40      End If
50      If SerialPort Is Nothing Then
60        Set SerialPort = New cComm
70      End If
80      SerialPort.CommClose
90      Sleep 100
100     InitComm = SerialPort.CommOpen(PortID, "")

InitComm_Resume:
110      On Error GoTo 0
120      Exit Function

InitComm_Error:

130     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at ModComm.InitComm." & Erl
140     Resume InitComm_Resume

  
End Function

Public Sub CloseComm(SerialPort As cComm)
10       On Error GoTo CloseComm_Error

20      SerialPort.CommClose

CloseComm_Resume:
30       On Error GoTo 0
40       Exit Sub

CloseComm_Error:

50      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at ModComm.CloseComm." & Erl
60      Resume CloseComm_Resume

End Sub

