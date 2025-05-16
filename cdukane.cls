VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDukane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ParsedBytes(0 To 256)   As Byte
Private Pointer                 As Long


Public Enabled        As Long

' port settingsLong
Public SerialPort     As Integer
Public Baud           As Long
Private mParity       As String
Public BITS           As Long
Public Flow           As Long
Public Stopbits       As String

Private mAlarmWords   As String
Private mRestoreWords As String

Public Preamble       As String
Public AutoClear      As Long  ' 0 = none

Public CommPort       As cComm  ' change to correct type

Public Packets        As Collection

Private Const section = "Dukane"

'"baud=9600 parity=N data=8 stop=1"

Private RestoreWordList()  As String
Private AlarmWordList()    As String

'Private Buffer()        As Byte

Public ESPackets As Collection

Function GetNextESPacket() As cESPacket
    Dim packet As cESPacket
    If ESPackets.Count Then
      Set GetNextESPacket = ESPackets(1)
      ESPackets.Remove 1
    End If
  
End Function

Sub ProcessPackets() ' called by main message loop
  Dim PacketCount As Long
  Dim j           As Long
  Dim packet      As cDuke
  
  
  Dim s As String
  Dim bytepacket As String
  
  
  PacketCount = Packets.Count
  
  For j = 1 To PacketCount
    Set packet = Packets(1)
    Packets.Remove 1
    ESPackets.Add packet.ToESPacket()
  Next
  

End Sub



Sub ReadDuke()

  'Global Const BYTE_STX = &H2
  'Global Const BYTE_ETX = &H3
  'Global Const BYTE_EOT = &H4
  'Global Const BYTE_ENQ = &H5
  'Global Const BYTE_ACK = &H6
  'Global Const BYTE_CR = &HD
  'Global Const BYTE_NAK = 15
  'Global Const BYTE_ESC = &H1B
  'Global Const BYTE_US = &H1F
  'Global Const BYTE_LF = &HA
  'Global Const BYTE_ETB = &H17
  'Global Const BYTE_SUB = &H1A

  ' &h82 is Start get ready for next record
  ' record begins with a single "@" sign
  ' Sync is "@@[0x80]

  'Record is
  ' 1 byte [0x82]
  ' 2 bytes [0xOD][0xOA] ' crlf
  ' 1 byte @
  ' 2 bytes sequence ? length ?
  ' 6 bytes for room identifier
  ' 1 byte tab
  ' 1 byte devicetype
  ' 1 byte tab
  ' 13 bytes message
  ' 1 byte [0x03] ETX '
  ' 1 byte tab
  ' 8 bytes date
  ' 1 byte space
  ' 8 bytes time
  ' 1 byte checksum
  ' 1 byte [0x83]
  ' 2 bytes [0xOD][0xOA] ' crlf
  ' 1 byte


  ' @ sign is [x80]

  ' [x80] END OF TRANSMISSION
  ' @@[x2C] END OF TRANSMISSION ?? bye ??

  '[x82][xOD][xOA]@[x3B][x05].10027[x09]3[x09]CANCEL.......[x03][x09]06/03/01.99:99:99[xNN][x83][xOD][xOA]
  '[x82][xOD][xOA]@[x3B][x05].10027[x09]3[x09]CANCEL.......[x03][x09]06/03/01.99:99:99[xNN][x83][xOD][xOA]
  '[x82][xOD][xOA]@[x3C][x04].10021[x09]5[x09]CALL_PLACE...[x03][x09]06/03/01.99:99:99[xNN][x83][xOD][xOA]
  ' ^^^@^^.10021t5tCALL_PLACE...^t06/03/01.99:99:99^^^^  <-  49 char from 0x82 to 0x83 inclusive


  '  Dim Packetstring As String
  '  Dim PacketBytes(0 To 20) As Byte
  '  Dim Fields() As String




  Dim ParsedString       As String

  Dim i                  As Long

  Dim numbytes           As Long
  Dim Char               As Byte
  Dim bytesread          As Long
  Dim DukePacket         As cDuke

  If Not CommPort Is Nothing Then
    If CommPort.PortOpen Then
    Do While CommPort.inbuffercount > 0

      If bytesread > 1200 Then Exit Do  ' yield to other processes

      If Pointer > 127 Then    ' hopelessly hosed... come back later
        Pointer = 0
        Erase ParsedBytes
        Exit Do
      End If

      bytesread = bytesread + CommPort.CommReadByte(Char)

      If Char >= 32 And Char < 127 Then
        Debug.Print Chr(Char) & " ";
      Else
        Debug.Print Right("00" & Hex(Char), 2) & " ";
      End If

      Select Case Char
        Case &H82              'begin record
          Erase ParsedBytes
          Pointer = 0
          ParsedBytes(Pointer) = Char
        Case &HD               'CR

          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char
        Case &HA               'LF
          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char

        Case &H40              ' At sign hex 40
          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char


        Case &H20 To &H7E      ' ASCII printable chars, space thru tilde
          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char

        Case &H9               '  tab
          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char


        Case &H3               'This is all the data we really need
          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char

        Case &H83              ' end of record
          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char

          ParsedString = StrConv(ParsedBytes, vbUnicode)

          i = InStr(1, ParsedString, Chr(3))
          If i >= 25 Then
            Set DukePacket = New cDuke
            If DukePacket.ParsePacket(ParsedString) < 0 Then
              ' error
            Else
              If IsAlarm(DukePacket) Then
                Packets.Add DukePacket
              ElseIf IsRestore(DukePacket) Then
                Packets.Add DukePacket
              Else
                ' toss it
              End If
            End If
          End If

          Erase ParsedBytes
          Pointer = 0

        Case &H80              ' non-data outside of packets sync ???
          Erase ParsedBytes
          Pointer = 0

        Case Else              ' bad

          Pointer = Pointer + 1
          ParsedBytes(Pointer) = Char

      End Select



      i = i + 1
      DoEvents
    Loop
    End If

  End If

End Sub
Function IsAlarm(packet As cDuke) As Boolean
  Dim j As Long
  Dim message As String
  message = Trim$(packet.DukeMessage)
 
  For j = 1 To UBound(AlarmWordList)
    If 0 = StrComp(AlarmWordList(j), message, vbTextCompare) Then
      packet.DukeStatus = 1
      Exit For
    End If
  Next
  
  IsAlarm = j <= UBound(AlarmWordList)
  
End Function
Function IsRestore(packet As cDuke) As Boolean
  Dim j As Long
  Dim message As String
  message = Trim$(packet.DukeMessage)
 
  For j = 1 To UBound(RestoreWordList)
    If 0 = StrComp(RestoreWordList(j), message, vbTextCompare) Then
      packet.DukeStatus = 0
      Exit For
    End If
  Next
  
  IsRestore = j <= UBound(RestoreWordList)
  
End Function



Public Function GetSerialSettings() As String
  GetSerialSettings = "baud=" & Baud & " parity=" & Parity & " data=" & BITS & " stop=" & Stopbits
End Function


Public Function Read() As Boolean
' check for data and process as needed
  Dim bite      As Byte
  Dim rc        As Long
  Dim j         As Integer

  If CommPort Is Nothing Then
    Exit Function
  Else
    If CommPort.PortOpen Then
      For j = 1 To CommPort.inbuffercount
        rc = CommPort.CommReadByte(bite)
        ' stuff into read buffer
      Next
    End If
  End If

End Function


Public Sub Init()
' read settings and start if needed
  Set Packets = New Collection
  ReLoad
  InitSerialPort

End Sub
Public Function InitSerialPort() As Long

10      Dispose

20      If (SerialPort > 0) And (Enabled > 0) Then
          ' start serial port
30        Set CommPort = New cComm
40        CommPort.CommOpen SerialPort, GetSerialSettings
50        Debug.Print "Duke Settings ", SerialPort, GetSerialSettings
60        If CommPort.PortOpen Then
70          CommPort.CommFlush
80        Else
90          LogProgramError "Error (Dukane Port Not Opened) at cDukane.InitSerialPort." & 40
100       End If
110     End If

End Function


Public Property Get SerialSettings() As String
  SerialSettings = Baud & Parity & BITS & Stopbits
End Property


Public Function Save() As Long

  WriteSetting section, "AlarmWords", AlarmWords
  WriteSetting section, "RestoreWords", RestoreWords

  WriteSetting section, "Preamble", Preamble

  WriteSetting section, "SerialPort", SerialPort
  WriteSetting section, "Baud", Baud
  WriteSetting section, "Parity", Parity
  WriteSetting section, "Bits", BITS
  WriteSetting section, "Flow", Flow
  WriteSetting section, "StopBits", Stopbits


  WriteSetting section, "AutoClear", AutoClear
  WriteSetting section, "Enabled", Enabled
  ' close serial port if open
  ' if enabled then start serial port



  InitSerialPort

End Function

Public Function ReLoad() As Long
  AlarmWords = ReadSetting(section, "AlarmWords", "Alarm")
  RestoreWords = ReadSetting(section, "RestoreWords", "Restore")

  Preamble = Trim$(ReadSetting(section, "Preamble", ""))

  SerialPort = Val(ReadSetting(section, "SerialPort", 0))
  Baud = Val(ReadSetting(section, "Baud", 9600))
  Parity = ReadSetting(section, "Parity", "N")
  BITS = Val(ReadSetting(section, "Bits", 8))
  Flow = Val(ReadSetting(section, "Flow", 0))
  Stopbits = Val(ReadSetting(section, "StopBits", 1))


  AutoClear = Val(ReadSetting(section, "AutoClear", 0))
  Enabled = Val(ReadSetting(section, "Enabled", 0))



End Function

Private Sub Dispose()
' deallocate serial port
' on error resume next
  If Not (CommPort Is Nothing) Then
    CommPort.CommClose
    Set CommPort = Nothing
  End If
  
  
End Sub

Private Sub Class_Initialize()
  'ReDim Buffer(0 To 32000)
  Set ESPackets = New Collection
  Set Packets = New Collection
  ReDim RestoreWordList(0)
  ReDim AlarmWordList(0)

End Sub

Private Sub Class_Terminate()
  Dispose
End Sub

Public Property Get AlarmWords() As String

  AlarmWords = mAlarmWords

End Property

Public Property Let AlarmWords(ByVal WordList As String)

  Dim j           As Long
  Dim i           As Long
  Dim Word        As String
  Dim Words()     As String
  
  i = 0
  ReDim AlarmWordList(i)
  Words = Split(WordList, " ")

  For j = LBound(Words) To UBound(Words)

    Word = Trim$(Words(j))
    If Len(Word) > 0 Then
      i = i + 1
      ReDim Preserve AlarmWordList(i)
      AlarmWordList(i) = Word
    End If
  Next

  mAlarmWords = WordList

End Property

Public Property Get RestoreWords() As String

  RestoreWords = mRestoreWords

End Property

Public Property Let RestoreWords(ByVal WordList As String)

  Dim j           As Long
  Dim i           As Long
  Dim Word        As String
  Dim Words()     As String

  i = 0
  ReDim RestoreWordList(i)
  Words = Split(WordList, " ")

  For j = LBound(Words) To UBound(Words)
    Word = Trim$(Words(j))
    If Len(Word) > 0 Then
      i = i + 1
      ReDim Preserve RestoreWordList(i)
      RestoreWordList(i) = Word
    End If
  Next

  mRestoreWords = WordList

End Property


Public Property Get Parity() As String

  Parity = mParity

End Property

Public Property Let Parity(ByVal Value As String)

  mParity = Value

End Property
