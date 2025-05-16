Attribute VB_Name = "modLib"
Option Explicit

Public Declare Function GetInputState Lib "user32" () As Long


Private Type SYSTEMTIME
 wYear As Integer
 wMonth As Integer
 wDayOfWeek As Integer
 wDay As Integer
 wHour As Integer
 wMinute As Integer
 wSecond As Integer
 wMilliseconds As Integer
End Type
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)


Public Enum HKEY_Type
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

'Registry Entry Types
'------------------------------------------------------------------------

Public Enum Reg_Type
  REG_NONE = 0  'No data type.
  REG_SZ = 1  'A string terminated by a null character.
  REG_EXPAND_SZ = 2  'A null-terminated string which contains unexpanded environment variables.
  REG_BINARY = 3  'A non-text sequence of bytes.
  REG_DWORD = 4  'Same as REG_DWORD_LITTLE_ENDIAN.
  REG_DWORD_LITTLE_ENDIAN = 4  'A 32-bit integer stored in little-endian format. This is the way Intel-based computers normally store numbers.
  REG_DWORD_BIG_ENDIAN = 5  'A 32-bit integer stored in big-endian format. This is the opposite of the way Intel-based computers normally store numbers -- the word order is reversed.
  REG_LINK = 6  'A Unicode symbolic link.
  REG_MULTI_SZ = 7  'A series of strings, each separated by a null character and the entire set terminated by a two null characters.
  REG_RESOURCE_LIST = 8  'A list of resources in the resource map.
End Enum


' Shell Constants
Global Const WAIT_INFINITE As Long = -1&
Global Const SYNCHRONIZE As Long = &H100000


'Security Constants
'------------------------------------------------------------------------

Global Const KEY_ALL_ACCESS = &HF003F  'Permission for all types of access.
Global Const KEY_CREATE_LINK = &H20  'Permission to create symbolic links.
Global Const KEY_CREATE_SUB_KEY = &H4  'Permission to create subkeys.
Global Const KEY_ENUMERATE_SUB_KEYS = &H8  'Permission to enumerate subkeys.
Global Const KEY_EXECUTE = &H20019  'Same as KEY_READ.
Global Const KEY_NOTIFY = &H10  'Permission to give change notification.
Global Const KEY_QUERY_VALUE = &H1  'Permission to query subkey data.
Global Const KEY_READ = &H20019  'Permission for general read access.
Global Const KEY_SET_VALUE = &H2  'Permission to set subkey data.
Global Const KEY_WRITE = &H20006  'Permission for general write access.

Global Const REG_OPTION_NON_VOLATILE = 0

'Error Numbers
'------------------------------------------------------------------------

Global Const REG_ERR_OK = 0  'No Problems
Global Const REG_ERR_NOT_EXIST = 1  'Key does not exist
Global Const REG_ERR_NOT_STRING = 2  'Value is not a string
Global Const REG_ERR_NOT_DWORD = 4  'Value not DWORD

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_ARENA_TRASHED = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Public Const LVM_FIRST = &H1000
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)

Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

' ONTRAK IO
Public Declare Function OpenAduDevice Lib "AduHid.DLL" (ByVal iTimeout As Long) As Long
Public Declare Function WriteAduDevice Lib "AduHid.DLL" (ByVal aduHandle As Long, ByVal lpBuffer As String, ByVal lNumberOfBytesToWrite As Long, ByRef lBytesWritten As Long, ByVal iTimeout As Long) As Long
Public Declare Function ReadAduDevice Lib "AduHid.DLL" (ByVal aduHandle As Long, ByVal lpBuffer As String, ByVal lNumberOfBytesToRead As Long, ByRef lBytesRead As Long, ByVal iTimeout As Long) As Long
Public Declare Function CloseAduDevice Lib "AduHid.DLL" (ByVal iHandle As Long) As Long

Public Const SET_RELAY = "SK"
Public Const RESET_RELAY = "RK"

Public Const MODE_RESERVED = 0   ' control is by other provider (KEY PA SYSTEM)

Public Const MODE_ONESHOT = 1    ' sets relay closed for one second

Public Const MODE_FLASHER = 2    ' toggles relay at once second on/off times as long as there is something in the Que
                                 ' QUE does NOT get emptied automatically
                             
Public Const MODE_ALWAYSON = 3   ' keeps relay closed as long as there is something in the Que
                                 ' QUE does NOT get emptied automatically
                                 ' Empties QUE on each pass (rising pulse)
Public Const RELAY_OFF = False
Public Const RELAY_ON = True

Public Const MAX_RELAYS = 8

Public Const ADU_NO_TIMEOUTS = 0
Public Const ADU_USE_TIMEOUTS = 1


Public PASystemKey      As Integer ' 0 or 1
Public PASystemPort     As Integer
Public PASystemHandle   As Long
Public PARepeatTwice    As Integer



Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal section$, ByVal Key$, ByVal default$, ByVal returnstring$, ByVal nSize&, ByVal filename$) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal section$, ByVal Key$, ByVal Value$, ByVal filename$) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal str As String) As Long


'local sound outputs
' called directly by alarm beeps
Public Declare Function PlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
' called testing a sound via PLayASound function
Public Declare Function PlayWavSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' called directly verifying alarm sound filenames
Public Declare Function PlayMemSound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long






Public sound1() As Byte
Public sound2() As Byte

Public Const SECONDSPERHOUR = 3600&
Public Const SecondsPerDay = 86400
Public Const SECONDSPERWEEK = 604800



Public Const BOOST_LIMIT = 50


Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long


Function PlayASound(ByVal AlarmFile As String, ByVal flags As Long)

  If MASTER Then
      PlayWavSound AlarmFile, flags
  Else
    PlayWavSound AlarmFile, flags
  End If
End Function


Public Function HexFormat(ByVal Value As Long, ByVal digits As Integer) As String
  HexFormat = Right(String$(digits, "0") & Hex(Value), digits)
End Function

Public Function ListViewGetVisibleCount(lv As ListView) As Long
  ListViewGetVisibleCount = SendMessage(lv.hwnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
End Function
Public Function ListViewAutoSize(lv As ListView) As Long

  Dim Col As Long
  For Col = 0 To lv.ColumnHeaders.Count - 1
    SendMessage lv.hwnd, LVM_SETCOLUMNWIDTH, Col, LVSCW_AUTOSIZE_USEHEADER
  Next Col
End Function
'Public Sub CenterFormOnForm(parent As Form, child As Form)
'
'  If Not (parent Is Nothing) Then
'    child.left = (parent.left) + (parent.width) / 2 - child.width / 2
'    child.top = (parent.top) + (parent.height) / 2 - child.height / 2
'  Else
'    Centerform child
'  End If
'
'End Sub


Public Sub Centerform(f As Form)
  f.left = (Screen.Width - f.Width) \ 2
  f.top = (Screen.Height - f.Height) \ 2
End Sub

Public Sub SaveFormPos(f As Form)
  If f.WindowState = vbNormal Then
  End If
End Sub
Public Sub GetFormPos(f As Form)
  Centerform f
End Sub


Public Sub SelAll(t As Control)
  t.SelStart = 0
  t.SelLength = Len(t.text)
End Sub
Public Function KeyProcMax(TextControl As Object, Key As Integer, Optional minus As Boolean, Optional decimals As Integer, Optional MaxLen As Single, Optional maxvalue As Double) As Integer
'Place this function in the KeyPress event of a text Control
'to control numeric input.
'Pass the control and the KeyAscii keypress with optional True/False for Minus Sign and Decimals to the right
'if minus is not passed it is assumed to be false
'If decimals is omitted, then there will be no decimal point or fractional component to the number
'The return value will be 0 (no keypress) if it would violate the parameters
'

'the call:
'Keyascii = KeyProcMax(TextControlControl, KeyAscii [,True/False [,0 to 9 ])
'
  Const vbKeyMinus = 45
  Dim Temp  As String
  Dim lTextControl  As String
  Dim rTextControl  As String
  Dim IP    As Long
  Dim IL    As Long

  Const vbKeyDecimalPoint = 46

  IP = TextControl.SelStart
  IL = TextControl.SelLength
  Temp = TextControl.text
  lTextControl = left(Temp, IP)
  rTextControl = MID(Temp, (IP + IL) + 1)



  KeyProcMax = Key
  '  If IsMissing(decimals) Then
  '    decimals = 0
  '  Else
  '    decimals = Val(decimals)
  '  End If

  '  If IsMissing(minus) Then
  '    minus = False
  '  Else
  '    minus = minus
  '  End If

  Select Case Key
    Case vbKeyBack, vbKeySubtract, vbKeyDecimalPoint, vbKeyMinus, vbKey0 To vbKey9
      Select Case Key
        Case vbKeyBack
          'always good!
        Case vbKeyDecimalPoint
          If decimals = 0 Then  'whole numbers please !
            KeyProcMax = 0
          Else
            Temp = lTextControl & Chr$(Key) & rTextControl
            If InStr(InStr(Temp, ".") + 1, Temp, ".") Then  'double decimals
              KeyProcMax = 0
            Else
              If Len(Temp) - InStr(Temp, ".") > decimals Then  'too many decimals
                KeyProcMax = 0
              End If
            End If
          End If
        Case vbKeySubtract, vbKeyMinus
          If minus = False Then
            KeyProcMax = 0
          Else
            If TextControl.SelStart > 0 Then  'must be starting character
              KeyProcMax = 0
            Else
              Temp = Chr$(vbKeySubtract) & rTextControl
              If Temp Like "*-*-*" Then  'double minus
                KeyProcMax = 0
              End If
            End If
          End If
        Case vbKey0 To vbKey9
          Temp = lTextControl & Chr(Key) & rTextControl
          If decimals Then
            IP = InStr(Temp, ".")
            IL = Len(Temp)
            If (IP > 0) And (IL - IP) > decimals Then  'too many decimals
              KeyProcMax = 0
            End If
          End If
      End Select
    Case Else
      KeyProcMax = 0
  End Select
  If MaxLen Then
    If Len(Temp) > MaxLen Then
      KeyProcMax = 0
    End If
  End If
  If maxvalue Then
    If Abs(Val(Temp)) > maxvalue Then
      KeyProcMax = 0
    End If
  End If

End Function

Public Function KeyProcHex(TextControl As Object, Key As Integer, MaxLen As Long) As Integer
'Place this function in the KeyPress event of a text Control
'to control numeric input.
'Pass the control and the KeyAscii keypress with optional True/False for Minus Sign and Decimals to the right
'if minus is not passed it is assumed to be false
'If decimals is omitted, then there will be no decimal point or fractional component to the number
'The return value will be 0 (no keypress) if it would violate the parameters
'

'the call:
'Keyascii = KeyProcMax(TextControlControl, KeyAscii [,True/False [,0 to 9 ])
'
  Dim Temp  As String
  Dim lTextControl  As String
  Dim rTextControl  As String
  Dim IP    As Long
  Dim IL    As Long


  IP = TextControl.SelStart
  IL = TextControl.SelLength
  Temp = TextControl.text
  lTextControl = left(Temp, IP)
  rTextControl = MID(Temp, (IP + IL) + 1)



  KeyProcHex = Key
  '  If IsMissing(decimals) Then
  '    decimals = 0
  '  Else
  '    decimals = Val(decimals)
  '  End If

  '  If IsMissing(minus) Then
  '    minus = False
  '  Else
  '    minus = minus
  '  End If

  Select Case Key
    Case vbKeyBack, vbKey0 To vbKeyF
      Select Case Key
        Case vbKeyBack
          'always good!
        Case vbKey0 To vbKeyF
          Temp = lTextControl & Chr(Key) & rTextControl
      End Select
    Case Else
      KeyProcHex = 0
  End Select
  If MaxLen Then
    If Len(Temp) > MaxLen Then
      KeyProcHex = 0
    End If
  End If
  '  If MaxValue Then
  '    If Abs(Val(temp)) > MaxValue Then
  '      KeyProcMax = 0
  '    End If
  '  End If

End Function

Public Sub CenterFormOnForm(child As Form, Optional Parent As Form)
  If Parent Is Nothing Then
    child.Move (Screen.Width - child.Width) / 2, (Screen.Height - child.Height) / 2
  Else
    child.Move (Parent.Width - child.Width) / 2 + Parent.left, (Parent.Height - child.Height) / 2 + Parent.top
  End If
End Sub
Public Function ToUpper(ByVal KeyAscii As Integer) As Integer
  ToUpper = Asc(UCase(Chr(KeyAscii)))
End Function

Public Function Parentheses(ByVal s As String) As String
  Parentheses = "(" & s & ")"
End Function

Public Function q(ByVal s As String) As String
  q = "'" & FixQuotes(s) & "'"
End Function

Public Function FixQuotes(ByVal s As String) As String
  FixQuotes = Replace(s, "'", "''")
End Function
Public Function DQ(ByVal s As String) As String

  DQ = """" & s & """"


End Function

Function DateDelimit(ByVal d As Date) As String
  DateDelimit = DateDelim & d & DateDelim
End Function

Function StripWhiteSpace(ByVal s As String) As String

  Do While InStr(1, s, "  ", vbTextCompare)
    s = Replace(s, "  ", " ")
  Loop
  StripWhiteSpace = s
End Function

Public Function ValidateHEXChecksum(ByVal Value As String) As Boolean
  Dim j As Integer
  Dim Checksum  As Long
  For j = 1 To Len(Value) - 2 Step 2
    Checksum = Checksum + HexToInt(MID(Value, j, 2))
  Next
  ValidateHEXChecksum = (Checksum And 255) = HexToInt(MID(Value, j, 2))

End Function
Public Function ValidateChecksum(Words() As String, wordcount) As Boolean
  Dim j As Integer
  Dim Checksum  As Long
  For j = 0 To wordcount - 2
    Checksum = Checksum + HexToInt(Words(j))
  Next
  ValidateChecksum = (Checksum And 255) = HexToInt(Words(wordcount - 1))

End Function

Public Function HexToInt(ByVal HexString As String) As Integer
'For j = 0 To 128: Print "Case """ & Right("00" & Hex(j), 2) & """" & vbCrLf & "HexToInt = " &  j: Next

  Select Case HexString
    Case "00"
      HexToInt = 0
    Case "01"
      HexToInt = 1
    Case "02"
      HexToInt = 2
    Case "03"
      HexToInt = 3
    Case "04"
      HexToInt = 4
    Case "05"
      HexToInt = 5
    Case "06"
      HexToInt = 6
    Case "07"
      HexToInt = 7
    Case "08"
      HexToInt = 8
    Case "09"
      HexToInt = 9
    Case "0A"
      HexToInt = 10
    Case "0B"
      HexToInt = 11
    Case "0C"
      HexToInt = 12
    Case "0D"
      HexToInt = 13
    Case "0E"
      HexToInt = 14
    Case "0F"
      HexToInt = 15
    Case "10"
      HexToInt = 16
    Case "11"
      HexToInt = 17
    Case "12"
      HexToInt = 18
    Case "13"
      HexToInt = 19
    Case "14"
      HexToInt = 20
    Case "15"
      HexToInt = 21
    Case "16"
      HexToInt = 22
    Case "17"
      HexToInt = 23
    Case "18"
      HexToInt = 24
    Case "19"
      HexToInt = 25
    Case "1A"
      HexToInt = 26
    Case "1B"
      HexToInt = 27
    Case "1C"
      HexToInt = 28
    Case "1D"
      HexToInt = 29
    Case "1E"
      HexToInt = 30
    Case "1F"
      HexToInt = 31
    Case "20"
      HexToInt = 32
    Case "21"
      HexToInt = 33
    Case "22"
      HexToInt = 34
    Case "23"
      HexToInt = 35
    Case "24"
      HexToInt = 36
    Case "25"
      HexToInt = 37
    Case "26"
      HexToInt = 38
    Case "27"
      HexToInt = 39
    Case "28"
      HexToInt = 40
    Case "29"
      HexToInt = 41
    Case "2A"
      HexToInt = 42
    Case "2B"
      HexToInt = 43
    Case "2C"
      HexToInt = 44
    Case "2D"
      HexToInt = 45
    Case "2E"
      HexToInt = 46
    Case "2F"
      HexToInt = 47
    Case "30"
      HexToInt = 48
    Case "31"
      HexToInt = 49
    Case "32"
      HexToInt = 50
    Case "33"
      HexToInt = 51
    Case "34"
      HexToInt = 52
    Case "35"
      HexToInt = 53
    Case "36"
      HexToInt = 54
    Case "37"
      HexToInt = 55
    Case "38"
      HexToInt = 56
    Case "39"
      HexToInt = 57
    Case "3A"
      HexToInt = 58
    Case "3B"
      HexToInt = 59
    Case "3C"
      HexToInt = 60
    Case "3D"
      HexToInt = 61
    Case "3E"
      HexToInt = 62
    Case "3F"
      HexToInt = 63
    Case "40"
      HexToInt = 64
    Case "41"
      HexToInt = 65
    Case "42"
      HexToInt = 66
    Case "43"
      HexToInt = 67
    Case "44"
      HexToInt = 68
    Case "45"
      HexToInt = 69
    Case "46"
      HexToInt = 70
    Case "47"
      HexToInt = 71
    Case "48"
      HexToInt = 72
    Case "49"
      HexToInt = 73
    Case "4A"
      HexToInt = 74
    Case "4B"
      HexToInt = 75
    Case "4C"
      HexToInt = 76
    Case "4D"
      HexToInt = 77
    Case "4E"
      HexToInt = 78
    Case "4F"
      HexToInt = 79
    Case "50"
      HexToInt = 80
    Case "51"
      HexToInt = 81
    Case "52"
      HexToInt = 82
    Case "53"
      HexToInt = 83
    Case "54"
      HexToInt = 84
    Case "55"
      HexToInt = 85
    Case "56"
      HexToInt = 86
    Case "57"
      HexToInt = 87
    Case "58"
      HexToInt = 88
    Case "59"
      HexToInt = 89
    Case "5A"
      HexToInt = 90
    Case "5B"
      HexToInt = 91
    Case "5C"
      HexToInt = 92
    Case "5D"
      HexToInt = 93
    Case "5E"
      HexToInt = 94
    Case "5F"
      HexToInt = 95
    Case "60"
      HexToInt = 96
    Case "61"
      HexToInt = 97
    Case "62"
      HexToInt = 98
    Case "63"
      HexToInt = 99
    Case "64"
      HexToInt = 100
    Case "65"
      HexToInt = 101
    Case "66"
      HexToInt = 102
    Case "67"
      HexToInt = 103
    Case "68"
      HexToInt = 104
    Case "69"
      HexToInt = 105
    Case "6A"
      HexToInt = 106
    Case "6B"
      HexToInt = 107
    Case "6C"
      HexToInt = 108
    Case "6D"
      HexToInt = 109
    Case "6E"
      HexToInt = 110
    Case "6F"
      HexToInt = 111
    Case "70"
      HexToInt = 112
    Case "71"
      HexToInt = 113
    Case "72"
      HexToInt = 114
    Case "73"
      HexToInt = 115
    Case "74"
      HexToInt = 116
    Case "75"
      HexToInt = 117
    Case "76"
      HexToInt = 118
    Case "77"
      HexToInt = 119
    Case "78"
      HexToInt = 120
    Case "79"
      HexToInt = 121
    Case "7A"
      HexToInt = 122
    Case "7B"
      HexToInt = 123
    Case "7C"
      HexToInt = 124
    Case "7D"
      HexToInt = 125
    Case "7E"
      HexToInt = 126
    Case "7F"
      HexToInt = 127
    Case "80"
      HexToInt = 128
    Case "81"
      HexToInt = 129
    Case "82"
      HexToInt = 130
    Case "83"
      HexToInt = 131
    Case "84"
      HexToInt = 132
    Case "85"
      HexToInt = 133
    Case "86"
      HexToInt = 134
    Case "87"
      HexToInt = 135
    Case "88"
      HexToInt = 136
    Case "89"
      HexToInt = 137
    Case "8A"
      HexToInt = 138
    Case "8B"
      HexToInt = 139
    Case "8C"
      HexToInt = 140
    Case "8D"
      HexToInt = 141
    Case "8E"
      HexToInt = 142
    Case "8F"
      HexToInt = 143
    Case "90"
      HexToInt = 144
    Case "91"
      HexToInt = 145
    Case "92"
      HexToInt = 146
    Case "93"
      HexToInt = 147
    Case "94"
      HexToInt = 148
    Case "95"
      HexToInt = 149
    Case "96"
      HexToInt = 150
    Case "97"
      HexToInt = 151
    Case "98"
      HexToInt = 152
    Case "99"
      HexToInt = 153
    Case "9A"
      HexToInt = 154
    Case "9B"
      HexToInt = 155
    Case "9C"
      HexToInt = 156
    Case "9D"
      HexToInt = 157
    Case "9E"
      HexToInt = 158
    Case "9F"
      HexToInt = 159
    Case "A0"
      HexToInt = 160
    Case "A1"
      HexToInt = 161
    Case "A2"
      HexToInt = 162
    Case "A3"
      HexToInt = 163
    Case "A4"
      HexToInt = 164
    Case "A5"
      HexToInt = 165
    Case "A6"
      HexToInt = 166
    Case "A7"
      HexToInt = 167
    Case "A8"
      HexToInt = 168
    Case "A9"
      HexToInt = 169
    Case "AA"
      HexToInt = 170
    Case "AB"
      HexToInt = 171
    Case "AC"
      HexToInt = 172
    Case "AD"
      HexToInt = 173
    Case "AE"
      HexToInt = 174
    Case "AF"
      HexToInt = 175
    Case "B0"
      HexToInt = 176
    Case "B1"
      HexToInt = 177
    Case "B2"
      HexToInt = 178
    Case "B3"
      HexToInt = 179
    Case "B4"
      HexToInt = 180
    Case "B5"
      HexToInt = 181
    Case "B6"
      HexToInt = 182
    Case "B7"
      HexToInt = 183
    Case "B8"
      HexToInt = 184
    Case "B9"
      HexToInt = 185
    Case "BA"
      HexToInt = 186
    Case "BB"
      HexToInt = 187
    Case "BC"
      HexToInt = 188
    Case "BD"
      HexToInt = 189
    Case "BE"
      HexToInt = 190
    Case "BF"
      HexToInt = 191
    Case "C0"
      HexToInt = 192
    Case "C1"
      HexToInt = 193
    Case "C2"
      HexToInt = 194
    Case "C3"
      HexToInt = 195
    Case "C4"
      HexToInt = 196
    Case "C5"
      HexToInt = 197
    Case "C6"
      HexToInt = 198
    Case "C7"
      HexToInt = 199
    Case "C8"
      HexToInt = 200
    Case "C9"
      HexToInt = 201
    Case "CA"
      HexToInt = 202
    Case "CB"
      HexToInt = 203
    Case "CC"
      HexToInt = 204
    Case "CD"
      HexToInt = 205
    Case "CE"
      HexToInt = 206
    Case "CF"
      HexToInt = 207
    Case "D0"
      HexToInt = 208
    Case "D1"
      HexToInt = 209
    Case "D2"
      HexToInt = 210
    Case "D3"
      HexToInt = 211
    Case "D4"
      HexToInt = 212
    Case "D5"
      HexToInt = 213
    Case "D6"
      HexToInt = 214
    Case "D7"
      HexToInt = 215
    Case "D8"
      HexToInt = 216
    Case "D9"
      HexToInt = 217
    Case "DA"
      HexToInt = 218
    Case "DB"
      HexToInt = 219
    Case "DC"
      HexToInt = 220
    Case "DD"
      HexToInt = 221
    Case "DE"
      HexToInt = 222
    Case "DF"
      HexToInt = 223
    Case "E0"
      HexToInt = 224
    Case "E1"
      HexToInt = 225
    Case "E2"
      HexToInt = 226
    Case "E3"
      HexToInt = 227
    Case "E4"
      HexToInt = 228
    Case "E5"
      HexToInt = 229
    Case "E6"
      HexToInt = 230
    Case "E7"
      HexToInt = 231
    Case "E8"
      HexToInt = 232
    Case "E9"
      HexToInt = 233
    Case "EA"
      HexToInt = 234
    Case "EB"
      HexToInt = 235
    Case "EC"
      HexToInt = 236
    Case "ED"
      HexToInt = 237
    Case "EE"
      HexToInt = 238
    Case "EF"
      HexToInt = 239
    Case "F0"
      HexToInt = 240
    Case "F1"
      HexToInt = 241
    Case "F2"
      HexToInt = 242
    Case "F3"
      HexToInt = 243
    Case "F4"
      HexToInt = 244
    Case "F5"
      HexToInt = 245
    Case "F6"
      HexToInt = 246
    Case "F7"
      HexToInt = 247
    Case "F8"
      HexToInt = 248
    Case "F9"
      HexToInt = 249
    Case "FA"
      HexToInt = 250
    Case "FB"
      HexToInt = 251
    Case "FC"
      HexToInt = 252
    Case "FD"
      HexToInt = 253
    Case "FE"
      HexToInt = 254
    Case "FF"
      HexToInt = 255
  End Select

End Function

Public Function ReadSetting(ByVal section As String, ByVal Key As String, ByVal default As String) As String
  Dim rc As Long
  Dim AppPath   As String
  Dim Buffer    As String
  Dim Ptr       As Long

  AppPath = App.Path
  If Right(AppPath, 1) <> "\" Then
    AppPath = AppPath & "\"
  End If
  Buffer = String$(255, 0)
  rc = GetPrivateProfileString(section, Key, default, Buffer, Len(Buffer), AppPath & "FREEDOM2.INI")
  
  Ptr = InStr(Buffer, vbNullChar)
  If Ptr Then
    ReadSetting = left$(Buffer, Ptr - 1)
  End If
  


End Function
Public Sub WriteSetting(ByVal section As String, ByVal Key As String, ByVal Setting As Variant)

  Dim AppPath As String
  AppPath = App.Path
  If Right(AppPath, 1) <> "\" Then
    AppPath = AppPath & "\"
  End If


  Setting = Setting & ""
  WritePrivateProfileString section, Key, Setting, AppPath & "FREEDOM2.INI"

End Sub

Public Function FileExists(ByVal filename As String) As Boolean
  Dim Temp As String
  filename = Trim(filename)
  On Error Resume Next
  If Len(filename) Then
    Temp = Dir(filename, vbNormal)
    FileExists = (Len(Temp) > 0) And (Err.Number = 0)
  End If

End Function
Public Function GetWaveData(ByVal filename As String)
  Dim hfile As Integer
  Dim flen As Long
  Dim d() As Byte
  On Error Resume Next

  hfile = FreeFile
  Open filename For Binary Access Read As hfile
  flen = LOF(hfile)
  ReDim d(0 To flen - 1)
  Get #hfile, , d
  Close hfile
  GetWaveData = d



End Function

Public Function AddToCombo(cb As ComboBox, ByVal text As String, ByVal ItemData As Long) As Long
' add text and item data to combo box
' returns new index
  cb.AddItem text
  cb.ItemData(cb.NewIndex) = ItemData
  AddToCombo = cb.NewIndex
End Function


Public Function GetComboItemData(cb As ComboBox) As Long
' defaults to 0 if nothing selected
  If cb.ListIndex > -1 Then
    GetComboItemData = cb.ItemData(cb.ListIndex)
  End If
End Function
Public Function AddToListBox(lb As ListBox, ByVal text As String, ByVal ItemData As Long) As Long
' add text and item data to combo box
' returns new index
  lb.AddItem text
  lb.ItemData(lb.NewIndex) = ItemData
  AddToListBox = lb.NewIndex
End Function


Public Function GetListBoxItemData(lb As ListBox) As Long
' defaults to 0 if nothing selected
  If lb.ListIndex > -1 Then
    GetListBoxItemData = lb.ItemData(lb.ListIndex)
  End If
End Function

Public Sub ListBoxClearSelections(list As ListBox)
  Dim j As Integer
  If list.listcount > 0 Then
    For j = 0 To list.listcount - 1
      list.Selected(j) = False
    Next
    list.ListIndex = 0
  End If

End Sub


Public Function CboFindExact(cbo, ByVal text As String) As Integer
  CboFindExact = SendMessageByString(cbo.hwnd, Win32.CB_FINDSTRINGEXACT, -1, text)
End Function

Public Function CboGetIndexByItemData(cbo As ComboBox, ByVal ItemData As Long) As Integer
'gets listindex of matching itemdata, or returns -1 if not found

  Dim j As Integer
  For j = cbo.listcount - 1 To 0 Step -1
    Debug.Print cbo.list(j),
    Debug.Print cbo.ItemData(j)
    If cbo.ItemData(j) = ItemData Then
      Exit For
    End If
  Next
  CboGetIndexByItemData = j

End Function
Function GetComboByText(cbo As ComboBox, ByVal text As String) As Integer

  Dim j As Integer
  For j = cbo.listcount - 1 To 0 Step -1
    If 0 = StrComp(cbo.list(j), text, vbTextCompare) Then
      Exit For
    End If
  Next
  GetComboByText = j

End Function


Public Function ProtocolString(ByVal ProtocolID As Integer) As String
'Coverts protocolID to string
  Select Case ProtocolID
    Case PROTOCOL_MOBILE
      ProtocolString = PROTOCOL_MOBILE_TEXT
    Case PROTOCOL_APOLLO
      ProtocolString = PROTOCOL_APOLLO_TEXT
    Case PROTOCOL_REMOTE
      ProtocolString = PROTOCOL_REMOTE_TEXT
    Case PROTOCOL_TAP
      ProtocolString = PROTOCOL_TAP_TEXT
    Case PROTOCOL_TAP2
      ProtocolString = PROTOCOL_TAP2_TEXT
    
    
    Case PROTOCOL_TAP_IP
      ProtocolString = PROTOCOL_TAP_IP_TEXT
    Case PROTOCOL_COMP1
      ProtocolString = PROTOCOL_COMP1_TEXT
    Case PROTOCOL_COMP2
      ProtocolString = PROTOCOL_COMP2_TEXT
    Case PROTOCOL_TTS
      ProtocolString = PROTOCOL_TTS_TEXT
    Case PROTOCOL_EMAIL
      ProtocolString = PROTOCOL_EMAIL_TEXT
    Case PROTOCOL_PCA
      ProtocolString = PROTOCOL_PCA_TEXT
    Case PROTOCOL_DIALER
      ProtocolString = PROTOCOL_DIALER_TEXT
    Case PROTOCOL_CENTRAL
      ProtocolString = PROTOCOL_CENTRAL_TEXT
    Case PROTOCOL_MARQUIS
      ProtocolString = PROTOCOL_MARQUIS_TEXT
    Case PROTOCOL_ONTRAK
      ProtocolString = PROTOCOL_ONTRAK_TEXT
    Case PROTOCOL_DIALOGIC
      ProtocolString = PROTOCOL_DIALOGIC_TEXT
    Case PROTOCOL_SDACT2
      ProtocolString = PROTOCOL_SDACT2_TEXT
    Case Else  '  PROTOCOL_NONE
      ProtocolString = PROTOCOL_NONE_TEXT
  End Select

End Function


Public Function GetParityString(ByVal ID As Integer) As String
' converts parity id to string
  Select Case ID
    Case 1
      GetParityString = "O"
    Case 2
      GetParityString = "N"
    Case 3
      GetParityString = "M"
    Case 4
      GetParityString = "S"
    Case Else
      GetParityString = "E"
  End Select

End Function
Public Function GetParityID(ByVal Parity As String) As Long

' converts parity string to ID
  Select Case Parity
    Case "O"
      GetParityID = 1
    Case "N"
      GetParityID = 2
    Case "M"
      GetParityID = 3
    Case "S"
      GetParityID = 4
    Case Else  ' "E"
      GetParityID = 0
  End Select

End Function
'Public Function MinMax(a, b, c)
'  Dim temp
'  temp = Max(a, b)
'  temp = Min(temp, c)
'
'
'End Function

Public Function Max(a, b)
  If a > b Then
    Max = a
  Else
    Max = b
  End If
End Function
Public Function Min(a, b)
  If a < b Then
    Min = a
  Else
    Min = b
  End If

End Function
Public Sub NewSendkeys(ByVal s As String, Optional ByVal Wait As Boolean = False)
  SendKeyAPI s, Wait
End Sub

Public Sub SendKeyAPI(ByVal c As String, ByVal Wait As Boolean)
  Dim vk As Integer
  Dim scan As Integer
  Dim oemchar As String

  Select Case UCase$(c)  ' special cases as defined
    Case "{TAB}"
      vk = vbKeyTab
    Case "{ENTER}", "~"
      vk = vbKeyReturn
    Case "{ESC}"
      vk = vbKeyEscape
    Case "{UP}"
      vk = vbKeyUp
    Case "{LEFT}"
      vk = vbKeyLeft
    Case "{RIGHT}"
      vk = vbKeyRight
    Case "{DOWN}"
      vk = vbKeyDown
    Case "{HOME}"
      vk = vbKeyHome
    Case "{END}"
      vk = vbKeyEnd
    Case "{PGDN}"
      vk = vbKeyPageDown
    Case "{PGUP}"
      vk = vbKeyPageUp
    Case "{DELETE}", "{DEL}"
      vk = vbKeyDelete
    Case "{INSERT}", "{INS}"
      vk = vbKeyInsert
    Case "{BACKSPACE}", "{BS}", "{BKSP}"
      vk = vbKeyBack
    Case "{NUMLOCK}"
      vk = vbKeyNumlock
    Case "{F1}"
      vk = vbKeyF1
    Case "{F2}"
      vk = vbKeyF2
    Case "{F3}"
      vk = vbKeyF3
    Case "{F4}"
      vk = vbKeyF4
    Case "{F5}"
      vk = vbKeyF5
    Case "{F6}"
      vk = vbKeyF6
    Case "{F7}"
      vk = vbKeyF7
    Case "{F8}"
      vk = vbKeyF8
    Case "{F9}"
      vk = vbKeyF9
    Case "{F10}"
      vk = vbKeyF10
    Case "{F11}"
      vk = vbKeyF11
    Case "{F12}"
      vk = vbKeyF12


    Case Else
      If Len(c) = 1 Then
        vk = Asc(c)
      Else
        SendKeys c, Wait
        Exit Sub

      End If
  End Select


  ' Get the virtual key code for this character
  If vk <> 0 Then
    oemchar = "  "  ' needs to be 2 bytes
    CharToOem left$(c, 1), oemchar  ' Get the OEM character - preinitialize the buffer
    scan = OemKeyScan(Asc(oemchar)) And &HFF  ' Get the scan code for this key

    ' make it happen
    keybd_event vk, scan, 0, 0  ' Send the key down
    keybd_event vk, scan, KEYEVENTF_KEYUP, 0  ' Send the key up
  End If
End Sub





Public Function KeyProcAlpha(ByVal KeyAscii As Integer) As Integer
  KeyAscii = ToUpper(KeyAscii)
  KeyProcAlpha = KeyAscii
  Select Case KeyAscii
    Case vbKeyA To vbKeyZ
    Case vbKey0 To vbKey9
    Case vbKeyBack
    Case Else
      KeyProcAlpha = 0
  End Select
End Function

Public Function GetLevelString(LEvel) As String
  If IsNull(LEvel) Then
    GetLevelString = "User"
  Else
    Select Case LEvel
      Case LEVEL_FACTORY
        GetLevelString = "Factory"
      Case LEVEL_ADMIN
        GetLevelString = "Admin 2"
      Case LEVEL_SUPERVISOR
        GetLevelString = "Admin 1"
      Case Else
        GetLevelString = "User"
    End Select
  End If


End Function

'Sub TestShifts()
'  Dim e1 As Long
'  Dim e2 As Long
'  Dim e3 As Long
'  Dim testtime As Date
'  Dim j As Long
'
'  e1 = 18
'  e2 = 1
'  e3 = 1
'
'  Debug.Print "Shift Times ", e1; " "; e2; " "; e3
'  For j = 0 To 23
'    testtime = Format(Now, "mm/dd/yyyy")
'    testtime = DateAdd("n", j * 60 + 30, testtime)
'    'If j > 17 Then Stop
'
'    Debug.Print "Shift "; Format(testtime, "HH:NN:SS"); " "; GetCurrentShift(testtime, e1, e2, e3)
'  Next
'
'
'
'
'End Sub

Public Function GetCurrentShift() As Long ' (Optional ByVal testtime As Date, Optional ByVal e1 As Long, Optional ByVal e2 As Long, Optional ByVal e3 As Long) As Integer

  Dim EndFirstSec   As Long
  Dim EndNightSec   As Long
  Dim EndThirdSec   As Long

  Dim HasSecondShift As Boolean
  Dim HasThirdShift As Boolean


  
  Dim SecondsSinceMidnight As Long

  Dim ShiftNumber   As Integer
  
  On Error GoTo GetCurrentShift_Error
  
  GetConfig
  SecondsSinceMidnight = Timer()  ' seconds since midnight

  ' for testing only using optional params
'    If Not IsMissing(testtime) Then
'      Configuration.EndFirst = e1
'      Configuration.EndNight = e2 ' am
'      Configuration.EndThird = e3 ' am
'      SecondsSinceMidnight = DateDiff("s", Format(testtime, "mm/dd/yyyy"), testtime)
'    End If

  

  EndFirstSec = Configuration.EndFirst * SECONDSPERHOUR ' convert to seconds since midnight
  EndNightSec = Configuration.EndNight * SECONDSPERHOUR
  EndThirdSec = Configuration.EndThird * SECONDSPERHOUR

  ShiftNumber = SHIFT_DAY

  HasSecondShift = False
  HasThirdShift = False


  If Configuration.EndFirst = Configuration.EndNight Then   ' no second or third shift' regardless of third shift ending
    ShiftNumber = SHIFT_DAY
    HasSecondShift = False
    HasThirdShift = False

  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = False
    
    
    
    If Configuration.EndFirst > Configuration.EndNight Then
'      If Configuration.EndNight = 0 Then
'         EndNightSec = 24 * SecondsPerHour
'      End If
      
      If SecondsSinceMidnight >= EndNightSec Or SecondsSinceMidnight < EndFirstSec Then
        ShiftNumber = SHIFT_DAY
      End If
    Else 'if Configuration.EndFirst < Configuration.EndNight
      'If Configuration.EndFirst = 0 Then
      '   EndFirstSec = 24 * SecondsPerHour
      'End If
      
      If SecondsSinceMidnight >= EndNightSec And SecondsSinceMidnight < EndFirstSec Then
        ShiftNumber = SHIFT_DAY
      End If
    End If

    If Configuration.EndFirst > Configuration.EndNight Then
      If SecondsSinceMidnight >= EndFirstSec Or SecondsSinceMidnight < EndNightSec Then
        ShiftNumber = SHIFT_NIGHT
      End If
    Else
      If SecondsSinceMidnight >= EndFirstSec And SecondsSinceMidnight < EndNightSec Then
        ShiftNumber = SHIFT_NIGHT
      End If
    End If

  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = True
'    Debug.Print "Since Mid " & SecondsSinceMidnight
'
'    Debug.Print "End First " & EndFirstSec, Configuration.EndFirst
'    Debug.Print "End Night " & EndNightSec, Configuration.EndNight
'    Debug.Print "End Third " & EndThirdSec, Configuration.EndThird
    
    If Configuration.EndFirst > Configuration.EndThird Then
      If SecondsSinceMidnight >= EndThirdSec Or SecondsSinceMidnight < EndFirstSec Then
        ShiftNumber = SHIFT_DAY
      End If
    Else ' Configuration.EndFirst < Configuration.EndThird
      If SecondsSinceMidnight >= EndThirdSec And SecondsSinceMidnight < EndFirstSec Then
        ShiftNumber = SHIFT_DAY
      End If
    End If

    If Configuration.EndFirst > Configuration.EndNight Then
      If SecondsSinceMidnight >= EndFirstSec Or SecondsSinceMidnight < EndNightSec Then
        ShiftNumber = SHIFT_NIGHT
      End If
    Else
      If SecondsSinceMidnight >= EndFirstSec And SecondsSinceMidnight < EndNightSec Then
        ShiftNumber = SHIFT_NIGHT
      End If
    End If

    If Configuration.EndNight > Configuration.EndThird Then
      If SecondsSinceMidnight >= EndNightSec Or SecondsSinceMidnight < EndThirdSec Then
        ShiftNumber = SHIFT_GRAVE
      End If
    Else
      If SecondsSinceMidnight >= EndNightSec And SecondsSinceMidnight < EndThirdSec Then
        ShiftNumber = SHIFT_GRAVE
      End If
    End If

  Else
    ShiftNumber = SHIFT_DAY
    HasSecondShift = False
    HasThirdShift = False
  End If

  'Trace "GetCurrentShift = " & ShiftNumber, True
  GetCurrentShift = ShiftNumber

GetCurrentShift_Resume:
  On Error GoTo 0
  Exit Function

GetCurrentShift_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modLib.GetCurrentShift." & Erl
  Resume GetCurrentShift_Resume


End Function

Function ConvertHourToAMPM(ByVal hour As Integer) As String
  If hour < 12 Then
    If hour = 0 Then
      ConvertHourToAMPM = "MidNight"
    Else
      ConvertHourToAMPM = hour & " AM"
    End If
  ElseIf hour = 12 Then
    ConvertHourToAMPM = "Noon"
  Else
    ConvertHourToAMPM = hour - 12 & " PM"
  End If
End Function

Function ConvertLastFirst(ByVal LastName As String, ByVal FirstName As String) As String
  Dim s As String
  If Len(LastName) > 0 Then
    s = LastName
    If Len(FirstName) > 0 Then
      s = s & ", " & FirstName
    End If
  Else
    s = FirstName
  End If
  ConvertLastFirst = s
End Function



'   Global hWndConsole As Long, hConsoleDC As Long, rcfConsole As RECT
'Sub GetConsoleDCRC()
'  Dim pt As POINTAPI
'  hConsoleDC = GetWindowDC(hWndConsole)
'  GetWindowRect hWndConsole, rcfConsole
'  ClientToScreen hWndConsole, pt
'  pt.x = pt.x - rcfConsole.nLeft
'  pt.y = pt.y - rcfConsole.nTop
'  GetClientRect hWndConsole, rcfConsole
'  rcfConsole.nLeft = rcfConsole.nLeft + pt.x
'  rcfConsole.nRight = rcfConsole.nRight + pt.x
'  rcfConsole.nTop = rcfConsole.nTop + pt.y
'  rcfConsole.nBottom = rcfConsole.nBottom + pt.y
'End Sub
'Function PrintConsoleWindow(TextMode As Long) As Long
'  Dim hMemDC As Long, hMemBmp As Long, hMemBmpOld As Long
'  Dim bmi As BITMAPINFO, bm As BITMAP, rctConsole As RECT
'  Dim lf As LOGFONT, hFont As Long, hFontOld As Long, tm As TEXTMETRIC
'  Dim i As Long, hFile As Long
'  Dim PrinterName As String * MAX_PATH
'  Dim pd As PRINTER_DEFAULTS, hPrinter As Long
'  Dim hGlobal As Long, hGlobal2 As Long, dwNeeded As Long, pi2 As PRINTER_INFO_2
'  Dim hPrDC As Long, hBmpDC As Long, di As DOCINFO, dn As String * 64
'  GetConsoleDCRC
'
'  hMemDC = CreateCompatibleDC(hConsoleDC)
'  rctConsole = rcfConsole
'  '--- DIB-section for console screen ----
'  bmi.bmiHeader.biSize = SizeOf(bmi.bmiHeader)
'  bmi.bmiHeader.biWidth = (rctConsole.nRight - rctConsole.nLeft)
'  bmi.bmiHeader.biHeight = (rctConsole.nBottom - rctConsole.nTop)
'  bmi.bmiHeader.biPlanes = 1
'  bmi.bmiHeader.biBitCount = 24
'  bmi.bmiHeader.biCompression = BI_RGB
'  hMemBmp = CreateDIBSection(hMemDC, bmi, DIB_RGB_COLORS, 0, 0, 0)
'
'  GlobalLock hMemBmp
'  hMemBmpOld = SelectObject(hMemDC, hMemBmp)
'  GetObject hMemBmp, SizeOf(bm), bm
'  BitBlt hMemDC, 0, 0, bm.bmWidth, bm.bmHeight, hConsoleDC, rctConsole.nLeft, rctConsole.nTop, SRCCOPY
'  SelectObject hMemDC, hMemBmpOld
'  DeleteDC hMemDC
'  '--- Default printer ---
'  '----------------- Printing ------------
'  'hPrDC = CreateDC(ByVal Null, PrinterName$, ByVal %Null, ByVal %Null)
'  hPrDC = Printer.HDC
'  If hPrDC Then
'    Dim PrLogPixelsX As Long, PrLogPixelsY As Long, PrPageX As Long, PrPageY As Long
'    PrLogPixelsX = GetDeviceCaps(hPrDC, LOGPIXELSX)
'    PrLogPixelsY = GetDeviceCaps(hPrDC, LOGPIXELSY)
'    PrPageX = GetDeviceCaps(hPrDC, HORZRES)
'    PrPageY = GetDeviceCaps(hPrDC, VERTRES)
'
'    Dim szPrinter_X As Double, szPrinter_Y As Double
'    szPrinter_X = PrPageX / PrLogPixelsX  ' sm (max)
'    szPrinter_Y = PrPageY / PrLogPixelsY
'
'    Dim kScale As Double  ' 10% for borders
'    kScale = 0.9 * Min(szPrinter_X / szConsole_X, szPrinter_Y / szConsole_Y)
'    szPrinter_X = kScale * szConsole_X  ' sm
'    szPrinter_Y = kScale * szConsole_Y
'
'    Dim x As Long, y As Long  ' pixels
'    x = PrLogPixelsX * szPrinter_X
'    y = PrLogPixelsY * szPrinter_Y
'    dn = "Ltr"
'    di.cbSize = SizeOf(di)
'    di.lpszDocName = VarPtr(dn)
'    hBmpDC = CreateCompatibleDC(hPrDC)
'    SelectObject hBmpDC, hMemBmp
'
'    If StartDoc(hPrDC, di) > 0 Then
'      If StartPage(hPrDC) > 0 Then
'        StretchBlt hPrDC, (PrPageX - x) / 2, (PrPageY - y) / 2, _
'            x, y, hBmpDC, 0, 0, _
'            rctConsole.nRight - rctConsole.nLeft, _
'            rctConsole.nBottom - rctConsole.nTop, SRCCOPY
'        EndPage hPrDC
'      End If
'      EndDoc hPrDC
'    End If
'    DeleteDC hPrDC
'    DeleteDC hBmpDC
'  End If
'
'  If pi2 Then GlobalUnlock hGlobal
'  If hGlobal2 Then
'    If pi2.pDevMode Then
'      GlobalUnlock hGlobal2
'    End If
'  End If
'
'  If hGlobal Then GlobalFree hGlobal
'  If hGlobal2 Then GlobalFree hGlobal2
'  If hPrinter Then ClosePrinter hPrinter
'  DeleteObject hMemBmp
'  ReleaseDC hWndConsole, hConsoleDC
'End Function
'Function PbMain()
'  hWndConsole = GetForegroundWindow  ' <-------------- For test only
'  GetConsoleDCRC
'  Dim hBmp As Long, hBmpDC As Long
'  hBmp = LoadImage(ByVal Null, FileBmp, IMAGE_BITMAP, _
'      rcfConsole.nRight - rcfConsole.nLeft, rcfConsole.nBottom - rcfConsole.nTop, LR_LOADFROMFILE)
'  hBmpDC = CreateCompatibleDC(hConsoleDC)
'  SelectObject hBmpDC, hBmp
'  BitBlt hConsoleDC, rcfConsole.nLeft, rcfConsole.nTop, _
'      rcfConsole.nRight - rcfConsole.nLeft, rcfConsole.nBottom - rcfConsole.nTop, _
'      hBmpDC, 0, 0, SRCCOPY
'  DeleteDC hBmpDC
'  ReleaseDC hWndConsole, hConsoleDC
'  Dim i As Long
'  'color 0, 15: Cursor Off
'  For i = 1 To ScreenY
'    'Locate i, 1: Print "|";
'    'Locate i, 11: Print "   Line " + Format$(i, "00") + "   ";
'    'Locate i, ScreenX: Print "|";
'  Next
'  PrintConsoleWindow True
'  PrintConsoleWindow False
'
'
'
'
'End Function
Public Function messagebox(f As Form, ByVal text As String, ByVal Caption As String, ByVal style As Long) As Long
  ' a non-blocking message box
  'messagebox = Win32.MessageBoxEx(frmMain.hwnd, text, Caption, style, 0)
  messagebox = Win32.MessageBoxEx(f.hwnd, text, Caption, style, 0)
End Function
'
'To indicate the modality of the dialog box, specify one of the following values.
'MB_APPLMODAL
'The user must respond to the message box before continuing work in the window identified by the hWnd parameter. However, the user can move to the windows of other threads and work in those windows.
'Depending on the hierarchy of windows in the application, the user may be able to move to other windows within the thread. All child windows of the parent of the message box are automatically disabled, but pop-up windows are not.
'
'MB_APPLMODAL is the default if neither MB_SYSTEMMODAL nor MB_TASKMODAL is specified.
'
'MB_SYSTEMMODAL
'Same as MB_APPLMODAL except that the message box has the WS_EX_TOPMOST style. Use system-modal message boxes to notify the user of serious, potentially damaging errors that require immediate attention (for example, running out of memory). This flag has no effect on the user's ability to interact with windows other than those associated with hWnd.
'MB_TASKMODAL
'Same as MB_APPLMODAL except that all the top-level windows belonging to the current thread are disabled if the hWnd parameter is NULL. Use this flag when the calling application or library does not have a window handle available but still needs to prevent input to other windows in the calling thread without suspending other threads.
'To specify other options, use one or more of the following values.
'MB_DEFAULT_DESKTOP_ONLY
'Windows NT/2000/XP: Same as MB_SERVICE_NOTIFICATION except that the system will display the message box only on the default desktop of the interactive window station. For more information, see Window Stations.
'Windows NT 4.0 and earlier: If the current input desktop is not the default desktop, MessageBoxEx fails.
'
'Windows 2000/XP: If the current input desktop is not the default desktop, MessageBoxEx does not return until the user switches to the default desktop.
'
'Windows 95/98/Me: This flag has no effect.
'
'MB_RIGHT
'The text is right-justified.
'MB_RTLREADING
'Displays message and caption text using right-to-left reading order on Hebrew and Arabic systems.
'MB_SETFOREGROUND
'The message box becomes the foreground window. Internally, the system calls the SetForegroundWindow function for the message box.
'MB_TOPMOST


'Public Function AutoSel(cbo As ComboBox, keycode As Integer)
'  Dim text    As String
'  Dim i       As Long
'  Dim Temp    As String
'
'  Select Case keycode
'    Case vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyDelete, vbKeyPageUp, vbKeyPageDown, vbKeyEnd, vbKeyHome, vbKeyInsert
'    Case Else
'
'  End Select
'
'  text = cbo.text
'  If Len(text) > 0 Then
'    For i = 0 To cbo.ListCount
'      Temp = Left(cbo.List(i), Len(text))
'      If 0 = StrComp(Temp, text, vbTextCompare) Then
'        cbo.text = cbo.List(i)
'        cbo.ListIndex = i
'        cbo.SelStart = Len(text)
'        cbo.SelLength = Len(cbo.List(i))
'      End If
'    Next
'  End If
'End Function

Function AutoSel(cbo As ComboBox, KeyCode As Integer)
    Dim i As Integer
    Dim OriginalTextLength As Long
    If KeyCode = 8 Or KeyCode = 48 Or cbo.text = "" Then Exit Function

    For i = 0 To cbo.listcount - 1

        'If Text matches with the currenty list item :: Using LCase for ignoring Case Sensitive
        If 0 = StrComp(cbo.text, left(cbo.list(i), Len(cbo.text)), vbTextCompare) Then
            'Remember Text Length
            OriginalTextLength = Len(cbo.text)
            'Completes the Text
            cbo.text = cbo.list(i)
            'Set The start of the selection in the original text lenght
            cbo.SelStart = OriginalTextLength
            ' Select All the new text added
            cbo.SelLength = Len(cbo.text) - OriginalTextLength
            'As it has now searched for the first item that matches, stop searching
            Exit For
        End If
    Next


End Function
'Public Sub SendKeyAPI(ByVal c As String, ByVal wait As Boolean)
'  Dim vk As Integer
'  Dim scan As Integer
'  Dim oemchar As String
'  Dim dl As Long
'
'  Select Case UCase$(c)  ' special cases as defined
'    Case "{TAB}"
'      vk = vbKeyTab
'    Case "{ENTER}", "~"
'      vk = vbKeyReturn
'    Case "{ESC}"
'      vk = vbKeyEscape
'    Case "{UP}"
'      vk = vbKeyUp
'    Case "{LEFT}"
'      vk = vbKeyLeft
'    Case "{RIGHT}"
'      vk = vbKeyRight
'    Case "{DOWN}"
'      vk = vbKeyDown
'    Case "{HOME}"
'      vk = vbKeyHome
'    Case "{END}"
'      vk = vbKeyEnd
'    Case "{PGDN}"
'      vk = vbKeyPageDown
'    Case "{PGUP}"
'      vk = vbKeyPageUp
'    Case "{DELETE}", "{DEL}"
'      vk = vbKeyDelete
'    Case "{INSERT}", "{INS}"
'      vk = vbKeyInsert
'    Case "{BACKSPACE}", "{BS}", "{BKSP}"
'      vk = vbKeyBack
'    Case "{NUMLOCK}"
'      vk = vbKeyNumlock
'    Case "{F1}"
'      vk = vbKeyF1
'    Case "{F2}"
'      vk = vbKeyF2
'    Case "{F3}"
'      vk = vbKeyF3
'    Case "{F4}"
'      vk = vbKeyF4
'    Case "{F5}"
'      vk = vbKeyF5
'    Case "{F6}"
'      vk = vbKeyF6
'    Case "{F7}"
'      vk = vbKeyF7
'    Case "{F8}"
'      vk = vbKeyF8
'    Case "{F9}"
'      vk = vbKeyF9
'    Case "{F10}"
'      vk = vbKeyF10
'    Case "{F11}"
'      vk = vbKeyF11
'    Case "{F12}"
'      vk = vbKeyF12
'    Case Else
'        vk = Asc(c)
'  End Select
'
'
'  ' Get the virtual key code for this character
'  If vk <> 0 Then
'    oemchar = "  "  ' needs to be 2 bytes
'    CharToOem Left$(c, 1), oemchar  ' Get the OEM character - preinitialize the buffer
'    scan = OemKeyScan(Asc(oemchar)) And &HFF  ' Get the scan code for this key
'
'    ' make it happen
'    keybd_event vk, scan, 0, 0  ' Send the key down
'    keybd_event vk, scan, KEYEVENTF_KEYUP, 0  ' Send the key up
'  End If
'End Sub
Sub SendKeypress2Window(ByVal hwnd As Long, ByVal c As String)

  Dim vk As Long
  Dim rc As Long

  Select Case UCase$(c)  ' special cases as defined
    Case "{TAB}"
      vk = vbKeyTab
    Case "{ENTER}", "~"
      vk = vbKeyReturn
    Case "{ESC}"
      vk = vbKeyEscape
    Case "{UP}"
      vk = vbKeyUp
    Case "{LEFT}"
      vk = vbKeyLeft
    Case "{RIGHT}"
      vk = vbKeyRight
    Case "{DOWN}"
      vk = vbKeyDown
    Case "{HOME}"
      vk = vbKeyHome
    Case "{END}"
      vk = vbKeyEnd
    Case "{PGDN}"
      vk = vbKeyPageDown
    Case "{PGUP}"
      vk = vbKeyPageUp
    Case "{DELETE}", "{DEL}"
      vk = vbKeyDelete
    Case "{INSERT}", "{INS}"
      vk = vbKeyInsert
    Case "{BACKSPACE}", "{BS}", "{BKSP}"
      vk = vbKeyBack
    Case "{NUMLOCK}"
      vk = vbKeyNumlock
    Case "{F1}"
      vk = vbKeyF1
    Case "{F2}"
      vk = vbKeyF2
    Case "{F3}"
      vk = vbKeyF3
    Case "{F4}"
      vk = vbKeyF4
    Case "{F5}"
      vk = vbKeyF5
    Case "{F6}"
      vk = vbKeyF6
    Case "{F7}"
      vk = vbKeyF7
    Case "{F8}"
      vk = vbKeyF8
    Case "{F9}"
      vk = vbKeyF9
    Case "{F10}"
      vk = vbKeyF10
    Case "{F11}"
      vk = vbKeyF11
    Case "{F12}"
      vk = vbKeyF12
    Case Else
      If Len(c) = 1 Then
        vk = Asc(c)
      End If
  End Select

  If vk > 0 Then
    APIKeyPress hwnd, vk
  End If
End Sub

Public Sub APIKeyPress(ByVal hwnd As Long, ByVal VirtualKey As Long)
' pretend we're the keyboard
' send keydown, keypress, and keyup


  Dim rc As Long
  rc = SendMessage(hwnd, WM_KEYDOWN, ByVal VirtualKey, ByVal 0&)
  rc = SendMessage(hwnd, WM_CHAR, ByVal VirtualKey, ByVal 0&)
  rc = SendMessage(hwnd, WM_KEYUP, ByVal VirtualKey, ByVal 0&)

End Sub

Public Sub HyperLink(ByVal url As String, ByVal params As String)
  On Error Resume Next
  If Len(params) > 0 Then
    url = url & "?" & params
  End If
  Call ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)
  
End Sub


Public Function ValueToBinary(ByVal Value As Long) As String
  Dim j As Long
  Dim s As String
  Dim x As Long
  
  For j = 15 To 0 Step -1
    If j = 7 Then
      s = s & " "
    End If
    x = Value And 2 ^ j
    s = s & IIf(x, "1", "0")
  Next
  ValueToBinary = s
End Function

Public Function Timer2Ms() As String
  Dim Seconds As Long
  Dim ms      As Integer
  Dim timerdata As Double
  timerdata = Timer
  Seconds = Fix(timerdata)
  Timer2Ms = Right(Format(timerdata - Seconds, ".000"), 4)
  
End Function

Public Function MaxL(ByVal a As Long, ByVal b As Long) As Long
  If a > b Then
    MaxL = a
  Else
    MaxL = b
  End If
End Function
Public Function MinL(ByVal a As Long, ByVal b As Long) As Long
  If a < b Then
    MinL = a
  Else
    MinL = b
  End If
End Function

Public Sub SetFocusTo(ctrl As Object)
  On Error Resume Next
  ctrl.SetFocus
End Sub

Public Function HTMLTD(ByVal text As String, Optional ByVal attributes As String = "") As String
  Dim Td As String
  'if attributes string is provided, it's added in the <td as an attribute
  'text passed in is checked for valid html, converting to entities as needed
  
  attributes = Trim$(attributes)
  If Len(attributes) > 0 Then
    attributes = " " & attributes
  End If
  HTMLTD = "<td" & attributes & ">" & HTMLEncode(text) & "</td>"
  
End Function

Public Function DaypartToString(ByVal DayPart As Long) As String
  
  If DayPart = 0 Then
    DaypartToString = "12 Midnight"
  ElseIf DayPart = 12 Then
    DaypartToString = "12 Noon"
  ElseIf DayPart > 12 Then
    DaypartToString = DayPart - 12 & " PM"
  Else
    DaypartToString = DayPart & " AM"
  End If
  
End Function


Public Function DaysToString(ByVal DAYS As Long) As String

  Const DayList = "SMTWTFS"
  
  Dim j       As Long
  Dim result  As String
  
  For j = 0 To 6
    If DAYS And (2 ^ j) Then
      result = result & MID$(DayList, j + 1, 1)
    Else
      result = result & "_"
    End If
  Next
  DaysToString = result

End Function

Public Function EventToString(ByVal EventID As Long) As String
  
  Select Case EventID
    Case EVT_EMERGENCY
      EventToString = "Alarms"
    Case EVT_ALERT
      EventToString = "Alerts"
    Case EVT_BATTERY_FAIL
      EventToString = "Low Battery"
    Case EVT_CHECKIN_FAIL
      EventToString = "Trouble"
    Case EVT_TAMPER
      EventToString = "Tamper"
    Case EVT_EXTERN
      EventToString = "External"
    Case Else
      EventToString = "Unspecified"
  End Select
  
End Function

Public Function FileFormatToString(ByVal Format As Long) As String
  Select Case Format
    Case AUTOREPORTFORMAT_TAB_NOHEADER
      FileFormatToString = "Tab Delimited, No Headers"
    Case AUTOREPORTFORMAT_HTML
      FileFormatToString = "HTML"
    Case Else '      AUTOREPORTFORMAT_TAB
      FileFormatToString = "Tab Delimited"
    End Select

End Function

Public Function PeriodToString(ByVal period As Long) As String
  Select Case period
    Case AUTOREPORT_DAILY: PeriodToString = "Daily"
    Case AUTOREPORT_SHIFT1: PeriodToString = "1st Shift"
    Case AUTOREPORT_SHIFT2: PeriodToString = "2nd Shift"
    Case AUTOREPORT_SHIFT3: PeriodToString = "3rd Shift"
    
    Case AUTOREPORT_WEEKLY: PeriodToString = "Weekly"
    Case AUTOREPORT_MONTHLY: PeriodToString = "Monthly"
    Case Else: PeriodToString = "Unspecified"
  End Select
End Function

Public Function IEEE_754_Bytes_to_Single(ByVal Byte0 As Byte, ByVal Byte1 As Byte, ByVal Byte2 As Byte, ByVal Byte3 As Byte) As Single
  ' reverse order!
  ' 429D999A which is 78 F byte0 is 9A, byte1 is 99, byte2 is 9D , byte 3 is
  Dim result          As Single
  Dim Bytes(0 To 3)   As Byte
  
  Bytes(0) = Byte0
  Bytes(1) = Byte1
  Bytes(2) = Byte2
  Bytes(3) = Byte3
  
  Call CopyMemory(result, Bytes(0), 4)
  
  IEEE_754_Bytes_to_Single = result
  
End Function


Public Function IEEE_754_ByteArray_to_Single(Bytes() As Byte) As Single
  Dim result As Single
  
  Call CopyMemory(result, Bytes(0), 4)
  IEEE_754_ByteArray_to_Single = result
End Function
Public Function IEEE_754_Hex_to_Single(ByVal HexString As String) As Single

  'Dim result          As Single
  Dim Bytes(0 To 3)   As Byte
  HexString = Right("00000000" & HexString, 8)
  
  '429D999A = 78.8 F
  
  Bytes(0) = Val("&h" & MID$(HexString, 7, 2))
  Bytes(1) = Val("&h" & MID$(HexString, 5, 2))
  Bytes(2) = Val("&h" & MID$(HexString, 3, 2))
  Bytes(3) = Val("&h" & MID$(HexString, 1, 2))
  
  IEEE_754_Hex_to_Single = IEEE_754_ByteArray_to_Single(Bytes)
  'Call CopyMemory(result, Bytes(0), 4)
  'IEEE_754_Hex_to_Single = result

End Function

Public Function Single_to_HexString(ByVal Value As Single) As String
  

End Function

Public Function Single_to_ByteArray(ByVal Value As Single) As Byte

End Function

Private Function DecodeBase64(ByVal EncodedText As String) As String
  Dim B64           As Base64

  Set B64 = New Base64
  DecodeBase64 = B64.Decode(EncodedText)   '  "QWRtaW46qwrTAw4=" =  ("Admin:Admin")
  Set B64 = Nothing

End Function
Public Function EncodeBase64(ByVal PlainText As String) As String
  Dim B64           As Base64
  Set B64 = New Base64
  EncodeBase64 = B64.Encode(PlainText)  ' ("Admin:Admin")   = "QWRtaW46qwrTAw4="
  Set B64 = Nothing
End Function

Public Function IsGarbage(ByVal s As String) As Boolean
  Dim j             As Long
  Dim character     As String
  Dim garbagechars  As Long
  For j = 1 To Len(s)
    character = MID$(s, j, 1)
    If Asc(character) = 63 Then
      garbagechars = garbagechars + 1
      If garbagechars > 3 Then
        IsGarbage = True
        Exit For
      End If
    Else
      garbagechars = garbagechars - 1
    End If
  Next
End Function

Public Function ValidateIPV4(ByVal IP As String) As Boolean
  Dim Octets
  Dim j As Long
  Dim Count As Long
  Dim Octet As String
  Dim OctetLen  As Long

  Octets = Split(IP, ".")
  For j = LBound(Octets) To UBound(Octets)
    Octet = Octets(j)
    OctetLen = Len(Octet)
    If OctetLen >= 1 And OctetLen <= 3 Then
      If Val(Octet) >= 0 And Val(Octet) <= 255 Then
        Count = Count + 1
      End If
    End If
  Next
  
  ValidateIPV4 = (Count = 4) And ((UBound(Octets) - LBound(Octets)) = 3)
  
End Function

Public Function InetErrorToString(ByVal InetError As Long) As String
  Dim s                  As String
  Select Case InetError
    Case ERROR_INTERNET_OUT_OF_HANDLES
      s = "No more handles could be generated at this time."

    Case ERROR_INTERNET_TIMEOUT
      s = "The request has timed out."

    Case ERROR_INTERNET_EXTENDED_ERROR
      s = "An extended error was returned from the server."
      '               This is typically a string or buffer containing a verbose error
      '               message. Call InternetGetLastResponseInfo to retrieve the
      '               error text.

    Case ERROR_INTERNET_INTERNAL_ERROR
      s = "An internal error has occurred."

    Case ERROR_INTERNET_INVALID_URL
      s = "The URL is invalid."

    Case ERROR_INTERNET_UNRECOGNIZED_SCHEME
      s = "The URL scheme could not be recognized or is not supported."

    Case ERROR_INTERNET_NAME_NOT_RESOLVED
      s = "The server name could not be resolved."

    Case ERROR_INTERNET_PROTOCOL_NOT_FOUND
      s = "The requested protocol could not be located."

    Case ERROR_INTERNET_INVALID_OPTION
      s = "A request to InternetQueryOption or InternetSetOption"
      s = "specified an invalid option value."

    Case ERROR_INTERNET_BAD_OPTION_LENGTH
      s = "The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified."

    Case ERROR_INTERNET_OPTION_NOT_SETTABLE
      s = "The request option cannot be set, only queried."

    Case ERROR_INTERNET_SHUTDOWN
      s = "The Win32 Internet function support is being shut down or unloaded."

    Case ERROR_INTERNET_INCORRECT_USER_NAME
      s = "The request to connect and log on to an FTP server could not be completed because the supplied user name is incorrect."

    Case ERROR_INTERNET_INCORRECT_PASSWORD
      s = "The request to connect and log on to an FTP server could not be completed because the supplied password is incorrect."

    Case ERROR_INTERNET_LOGIN_FAILURE
      s = "The request to connect to and log on to an FTP server failed."

    Case ERROR_INTERNET_INVALID_OPERATION
      s = "The requested operation is invalid."

    Case ERROR_INTERNET_OPERATION_CANCELLED
      s = "The operation was canceled."
      '               usually because the handle on
      '               which the request was operating was closed before the
      '               operation completed.

    Case ERROR_INTERNET_INCORRECT_HANDLE_TYPE
      s = "The type of handle supplied is incorrect for this operation."

    Case ERROR_INTERNET_INCORRECT_HANDLE_STATE
      s = "The requested operation cannot be carried out because the handle supplied is not in the correct state."

    Case ERROR_INTERNET_NOT_PROXY_REQUEST
      s = "The request cannot be made via a proxy."

    Case ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND
      s = "A required registry value could not be located."

    Case ERROR_INTERNET_BAD_REGISTRY_PARAMETER
      s = "A required registry value was located but is an incorrect type or has an invalid value."

    Case ERROR_INTERNET_NO_DIRECT_ACCESS
      s = "Direct network access cannot be made at this time."

    Case ERROR_INTERNET_NO_CONTEXT
      s = "An asynchronous request could not be made because a zero context value was supplied."

    Case ERROR_INTERNET_NO_CALLBACK
      s = "An asynchronous request could not be made because a callback function has not been set."

    Case ERROR_INTERNET_REQUEST_PENDING
      s = "The required operation could not be completed because one or more requests are pending."

    Case ERROR_INTERNET_INCORRECT_FORMAT
      s = "The format of the request is invalid."

    Case ERROR_INTERNET_ITEM_NOT_FOUND
      s = "The requested item could not be located."

    Case ERROR_INTERNET_CANNOT_CONNECT
      s = "The attempt to connect to the server failed."

    Case ERROR_INTERNET_CONNECTION_ABORTED
      s = "The connection with the server has been terminated."

    Case ERROR_INTERNET_CONNECTION_RESET
      s = "The connection with the server has been reset."

    Case ERROR_INTERNET_FORCE_RETRY
      s = "Calls for the Win32 Internet function to redo the request."

    Case ERROR_INTERNET_INVALID_PROXY_REQUEST
      s = "The request to the proxy was invalid."

    Case ERROR_INTERNET_HANDLE_EXISTS
      s = "The request failed because the handle already exists."

    Case ERROR_INTERNET_SEC_CERT_DATE_INVALID
      s = "SSL certificate date that was received from the server is bad. The certificate is expired."

    Case ERROR_INTERNET_SEC_CERT_CN_INVALID
      s = "SSL certificate common name (host name field) is incorrect."
      '               For example, if you entered www.server.com and the common
      '               name on the certificate says www.different.com.

    Case ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR
      s = "The application is moving from a non-SSL to an SSL connection because of a redirect."

    Case ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR
      s = "The application is moving from an SSL to an non-SSL connection because of a redirect."

    Case ERROR_INTERNET_MIXED_SECURITY
      s = "Indicates that the content is not entirely secure. Some of the content being viewed may have come from unsecured servers."

    Case ERROR_INTERNET_CHG_POST_IS_NON_SECURE
      s = "The application is posting and attempting to change multiple lines of text on a server that is not secure."

    Case ERROR_INTERNET_POST_IS_NON_SECURE
      s = "The application is posting data to a server that is not secure."

'    Case ERROR_FTP_TRANSFER_IN_PROGRESS
'      s = "The requested operation cannot be made on the FTP session handle because an operation is already in progress."

'    Case ERROR_FTP_DROPPED
'      s = "The FTP operation was not completed because the session was aborted."

'    Case ERROR_GOPHER_PROTOCOL_ERROR
'      s = "An error was detected while parsing data returned from the gopher server."
'
'    Case ERROR_GOPHER_NOT_FILE
'      s = "The request must be made for a file locator."
'
'    Case ERROR_GOPHER_DATA_ERROR
'      s = "An error was detected while receiving data from the gopher server."

'    Case ERROR_GOPHER_END_OF_DATA
'      s = "The end of the data has been reached."
'
'    Case ERROR_GOPHER_INVALID_LOCATOR
'      s = "The supplied locator is not valid."
'
'    Case ERROR_GOPHER_INCORRECT_LOCATOR_TYPE
'      s = "The type of the locator is not correct for this operation."
'
'    Case ERROR_GOPHER_NOT_GOPHER_PLUS
'      s = "The requested operation can only be made against a Gopher+ server."
'      ' or with a locator that specifies a Gopher+ operation.
'
'    Case ERROR_GOPHER_ATTRIBUTE_NOT_FOUND
'      s = "The requested attribute could not be located."
'
'    Case ERROR_GOPHER_UNKNOWN_LOCATOR
'      s = "The locator type is unknown."

'    Case ERROR_HTTP_HEADER_NOT_FOUND
'      s = "The requested header could not be located."

'    Case ERROR_HTTP_DOWNLEVEL_SERVER
'      s = "The server did not return any headers."

    Case ERROR_HTTP_INVALID_SERVER_RESPONSE
      s = "The server response could not be parsed."
'
'    Case ERROR_HTTP_INVALID_HEADER
'      s = "The supplied header is invalid."

'    Case ERROR_HTTP_INVALID_QUERY_REQUEST
'      s = "The request made to HttpQueryInfo is invalid."

'    Case ERROR_HTTP_HEADER_ALREADY_EXISTS
'      s = "The header could not be added because it already exists."

'    Case ERROR_HTTP_REDIRECT_FAILED
'      s = "The redirection failed."
      '               because either the scheme changed
      '               (for example, HTTP to FTP) or all attempts made to redirect
      '               failed (default is five attempts).
    Case 0
      s = "OK"
    Case Else
      s = "Unknown Error"
  End Select
  InetErrorToString = s
End Function

'Public Function SecondsToDaysTimeString(ByVal Seconds As Double)
'  Dim DAYS As Double
'  If Seconds >= SecondsPerDay Then
'    DAYS = Fix(Seconds / SecondsPerDay)
'    Seconds = Seconds - (DAYS * SecondsPerDay)
'    SecondsToDaysTimeString = DAYS & "." & Format$(DateAdd("s", Seconds, 0), "hh:nn:ss")
'  Else
'     SecondsToDaysTimeString = Format$(DateAdd("s", Seconds, 0), "hh:nn:ss")
' End If
'
'End Function



Public Function SecondsToTimeString(ByVal Seconds As Double)
  Dim DAYS As Double
  If Seconds >= SecondsPerDay Then
    DAYS = Fix(Seconds / SecondsPerDay)
    Seconds = Seconds - (DAYS * SecondsPerDay)
    SecondsToTimeString = DAYS & "." & Format$(DateAdd("s", Seconds, 0), "hh:nn:ss")
  Else
     SecondsToTimeString = Format$(DateAdd("s", Seconds, 0), "hh:nn:ss")
 End If

End Function

Public Function KillFile(ByVal filename As String) As Boolean
  On Error Resume Next
  If FileExists(filename) Then
      
    Kill filename
  End If
  KillFile = (Err.Number = 0)
End Function


'InIDE = RunningInIDE()
Private Function RunningInIDE() As Boolean
  'On Error Resume Next
  'Debug.Assert 1 / 0
  'RunningInIDE = Err.Number <> 0
  
  Dim filename      As String
  Dim Count         As Long
  On Error Resume Next
  
  filename = String(255, 0)
  Count = GetModuleFileName(App.hInstance, filename, 255)
  filename = left(filename, Count)
  filename = Right(filename, 7)
  If 0 = StrComp(filename, "VB6.EXE", vbTextCompare) Then
    RunningInIDE = True
  Else
    RunningInIDE = False
  End If

End Function

Public Sub SyncApacheUsers()

  Dim hfile              As Long
  Dim Buffer             As String
  Dim SQL                As String
  Dim rs                 As ADODB.Recordset
  Dim ActiveUsers        As Collection
  Dim ApacheUsers        As Collection
  Dim User               As cUser
  Dim Username           As String
  Dim j                  As Long
  Dim rc                 As Long
  Dim Commandline        As String
  Dim Process            As Long

  On Error GoTo SyncApacheUsers_Error
  ' read in users
  If (Configuration.MobileWebEnabled = 1) Then
    If (Len(Configuration.MobilehtPasswordPath) > 0) Then 'do sanity checks
      If (DirExists(Configuration.MobilehtPasswordPath)) Then
        If (Len(Configuration.MobilehtPasswordEXEPath) > 0) Then
          Set ActiveUsers = New Collection
          Set ApacheUsers = New Collection
          SQL = "SELECT username, password FROM users"
          Set rs = ConnExecute(SQL)
          On Error Resume Next
          Do Until rs.EOF
            Set User = New cUser
            User.Username = rs("username") & ""
            User.Password = rs("password") & ""

            ActiveUsers.Add User, User.Username
            
            rs.MoveNext
          Loop
          rs.Close
          Set rs = Nothing



          If (FileExists(Configuration.MobilehtPasswordPath & "\.htpasswd")) Then
            ' read in exisiting file
            hfile = FreeFile
            Open Configuration.MobilehtPasswordPath & "\.htpasswd" For Input As #hfile
            Do Until EOF(hfile)
              Dim UserString As String
              Line Input #hfile, UserString
              UserString = Trim$(UserString)
              If Len(UserString) Then
                Set User = New cUser
                If (User.ParseHtPassword(UserString)) Then
                  ApacheUsers.Add User, User.Username
                End If
              End If
            Loop
            Close #hfile
            'We should have our active and apache users here
            On Error Resume Next

            For j = ApacheUsers.Count To 1 Step -1
              Set User = Nothing
              Username = ApacheUsers(j).Username
              Set User = ActiveUsers(Username & "")  ' returns nothing if not found in the collection
              If User Is Nothing Then
                rc = Shell(Configuration.MobilehtPasswordEXEPath & " -D " & Configuration.MobilehtPasswordPath & "\.htpasswd " & Username, vbHide)  ' ApacheUsers.Remove j ' might not need this
               Process = Win32.OpenProcess(SYNCHRONIZE, True, rc)
               Call Win32.WaitForSingleObject(Process, WAIT_INFINITE)
               CloseHandle Process
              End If
            Next
            For Each User In ActiveUsers
              ' update if needed or not
              rc = Shell(Configuration.MobilehtPasswordEXEPath & " -b " & Configuration.MobilehtPasswordPath & "\.htpasswd " & User.Username & " " & User.Password, vbHide)
                             Process = Win32.OpenProcess(SYNCHRONIZE, True, rc)
               Call Win32.WaitForSingleObject(Process, WAIT_INFINITE)
               CloseHandle Process
            Next

          Else                 ' we need to create it
            On Error Resume Next
            For j = 1 To ActiveUsers.Count
              Set User = ActiveUsers(j)
              Commandline = Configuration.MobilehtPasswordEXEPath & " -b " & IIf(j = 1, "-c ", "") & Configuration.MobilehtPasswordPath & "\.htpasswd " & User.Username & " " & User.Password
              rc = Shell(Commandline, vbHide)
               Process = Win32.OpenProcess(SYNCHRONIZE, True, rc)
               Call Win32.WaitForSingleObject(Process, WAIT_INFINITE)
               CloseHandle Process
            Next
          End If
          
        End If
      End If
    End If
  End If



SyncApacheUsers_Resume:

  On Error GoTo 0
  Exit Sub

SyncApacheUsers_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modLib.SyncApacheUsers." & Erl
  Resume SyncApacheUsers_Resume

End Sub

Public Sub limitFileSize(ByVal filename As String, Optional ByVal Size As Long = 10 ^ 6)
        Dim hfile              As Long
        Dim hfile2             As Long
        '  Dim x                  As Byte
        '  Dim loc                As Long
        Dim fl                 As Long
        Dim start              As Long

        Dim bytesToCopy()      As Byte

        Dim MaxLen             As Long
        Dim NewLen             As Long


        Dim bakFile            As String

10      If FileExists(filename) Then
20        bakFile = filename & ".bak"
30        NewLen = Size              ' 1 meg as default
40        MaxLen = NewLen * 2        ' can grow to 2 meg

45          On Error Resume Next
50        If (FileLen(filename) > (MaxLen)) Then  ' it's gotten too big

70          Win32.DeleteFile bakFile
80          Name filename As bakFile  ' we've backed up our file by renaming it... very fast
90          Win32.DeleteFile filename

100         On Error GoTo 0
110         fl = FileLen(bakFile)    ' see how big it is

120         start = fl - NewLen      ' get our starting offset

130         hfile = FreeFile
140         Open bakFile For Binary As hfile  ' reading from backup file
150         hfile2 = FreeFile
160         Open filename For Binary As hfile2  ' writng to new file

170         ReDim bytesToCopy(1 To NewLen)  ' we'll create the new file with this number of bytes (default is 1 meg)

180         If NewLen < fl Then      ' sanity check
190           Get #hfile, start, bytesToCopy
200           Put #hfile2, , bytesToCopy
210         End If

            ' done
220         Close hfile
230         Close hfile2

240       End If
250     End If

End Sub

Public Sub DelayLoop(Optional ByVal Seconds As Long = 1, Optional ByVal Iterations As Long = 10, Optional ByVal UseDoevents As Boolean = True)

  ' delays execution n seconds (default 1)
  ' iterations = call to doevents per delay period default is 10
  ' optional doevents, typically yes, otherwise is total blocking
  ' with defaults, doevents is called every 100ms  (10/second)
  
  Dim rc
  Dim Count              As Long

  Dim st                 As SYSTEMTIME


  Dim WinEpoc            As Double
  
  Dim UnixTime           As Double
  Dim CurrentTime        As Double
  Dim BreakTime          As Double
  Dim FinishTime         As Double
  
  Dim SkipTime           As Double
  
  Const SecondsPerDay    As Double = 86400
  
  WinEpoc = DateSerial(1970, 1, 1)
  
  If Iterations <= 0 Then
    Iterations = 1
  End If
  
  SkipTime = Seconds / Iterations
  
  If SkipTime < 0.01 Then
    SkipTime = 0.01
  End If
    
  GetLocalTime st
  UnixTime = (Now - WinEpoc) * SecondsPerDay + (st.wMilliseconds * 0.001)

  CurrentTime = UnixTime  ' (Now - WinEpoc) * SecondsPerDay + (st.wMilliseconds * 0.001)
  BreakTime = CurrentTime + SkipTime
  FinishTime = CurrentTime + Seconds + SkipTime
  Do
    If (UseDoevents) Then
      If CurrentTime > BreakTime Then
        BreakTime = BreakTime + SkipTime
        'Debug.Print "Break "; BreakTime
        Count = Count + 1
        DoEvents
      End If
    End If

    GetLocalTime st
    
    CurrentTime = (Now - WinEpoc) * SecondsPerDay + (st.wMilliseconds * 0.001)

    If CurrentTime > FinishTime Then
      Exit Do
    End If
  Loop

End Sub


Public Sub UpdateApacheUser(ByVal Username As String, ByVal Password As String)

End Sub

