VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPacketizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
' 980

Private Packets  As Collection
Private packet   As cESPacket

Const MAXBUFFER  As Long = 4096& * 8&
Const MINBUFFER  As Long = 3
'Const MAXBYTE    As Integer = &HFF

Const MAX_BADPACKETS As Long = 50
Const BADPACKET_WARNING_CAUTION = 4
Const BADPACKET_WARNING_DANGER = 10

Public i6080 As c6080



Private Buffer(0 To MAXBUFFER)   As Byte

Private StartPtr           As Long
Private EndPtr             As Long
Private Checksum           As Integer

Private Header             As Byte
Private DataLength         As Byte

Public TotalBytes          As Long
Public totalpackets        As Long

Public BadPackets          As Long
Public BadPacketRetries    As Long

Private HighWaterMark      As Long

Public Function AddPacket(packet As cESPacket) As Long
  packet.DateTime = Now
  Packets.Add packet
End Function


Public Function GetPacket() As cESPacket
  Dim packet As cESPacket
  If Packets.Count > 0 Then
    totalpackets = totalpackets + 1
    Set packet = Packets(1)
    Packets.Remove 1
    packet.TimeStamp = packet.DateTime
  Else
    Set packet = New cESPacket  ' empty packet
  End If
  Set GetPacket = packet

End Function
Public Function PacketReady() As Boolean
  If gSuspendPackets Then
    PacketReady = False
  Else
    PacketReady = Packets.Count > 0  'PacketReady = Not (packet Is Nothing)
  End If
End Function



Public Function BufferCount() As Long
  Dim Count As Long
  Count = EndPtr - StartPtr  ' possible 0, positive or negative
  If Count < 0 Then
    Count = Count + MAXBUFFER + 1
  End If
  BufferCount = Count
End Function

Public Function Process() As Long
  Dim LengthPtr   As Long
  Dim BuffLen     As Long
  Dim CurrentPtr  As Long

  Dim j           As Long
  Dim bytebuffer() As Byte
  Dim Checksum    As Integer
  Dim packet      As cESPacket
  Dim SubPacket   As cESPacket

  BuffLen = BufferCount
  
  
  
  'dbg "Process.BufferLen " & BuffLen & vbCrLf

  If BuffLen >= MINBUFFER Then  ' 060208 is the shortest packet, 3 bytes
    ' ok to try an process packet
    LengthPtr = StartPtr + 1
    If LengthPtr > MAXBUFFER Then  ' do wrap
      LengthPtr = 0
    End If
    Header = Buffer(StartPtr)
    DataLength = Buffer(LengthPtr)

    
    If DataLength = &H11 Then
     ' Debug.Assert 0
    End If


    If DataLength < BuffLen Then  ' we need datalength +1
      ReDim bytebuffer(DataLength)  ' 0 to datalength + checksum
      ' validate checksum
      
      
      CurrentPtr = StartPtr

        

      For j = 0 To DataLength - 1
        bytebuffer(j) = Buffer(CurrentPtr)
        Checksum = (Checksum + bytebuffer(j)) And &HFF&
        CurrentPtr = CurrentPtr + 1
        If CurrentPtr > MAXBUFFER Then  ' handle wrap of pointer
          CurrentPtr = 0
        End If
      Next
      bytebuffer(j) = Buffer(CurrentPtr)
       

      If Buffer(CurrentPtr) = Checksum Then
        
        Set packet = New cESPacket
        
        packet.Bytes = bytebuffer
        'Debug.Print "Packet ";
        'Debug.Print Packet.HexPacket
        
        If (packet.BadPacket) Then ' toss packet if it's bad
          StartPtr = CurrentPtr + 1
          If StartPtr > MAXBUFFER Then
            StartPtr = 0
          End If
          BadPackets = BadPackets + 1
          
        Else  ' Good Packet here, may contain sub packets from Aggregated repeater packets
            
            packet.DateTime = Now
            Packets.Add packet
            
            StartPtr = CurrentPtr + 1
            If StartPtr > MAXBUFFER Then
              StartPtr = 0
            End If
            BadPackets = BadPackets \ 2
            For Each SubPacket In packet.SubPackets
              SubPacket.DateTime = Now
              Packets.Add SubPacket
            Next
          'End If
        End If
      Else
        Dim s As String
        If gNoDataErrorLog = False Then
          s = "Bad Data " & Format(Now, "mm-dd-yy hh:nn:ss") & ": " & BytesToHex(bytebuffer)
          SpecialLog s
          Trace s, True
        End If
        
        ' try to resync
        StartPtr = StartPtr + 1
        If StartPtr > MAXBUFFER Then
          StartPtr = 0
        End If
        BadPackets = BadPackets + 1
      End If
    Else
      'Trace "Short data " & Format(Win32.timeGetTime, "0")
    End If
  End If

End Function

Private Function BytesToHex(Buffer() As Byte)
  Dim j As Long
  Dim s As String
  For j = LBound(Buffer) To UBound(Buffer)
    s = s & Right("00" & Hex(Buffer(j)), 2)
  Next
  BytesToHex = s
End Function

Public Sub ClearBuffer()
  Erase Buffer  ' set all bytes to 0
  StartPtr = 0
  EndPtr = 0    ' point head and tail to start
  BadPackets = 0
End Sub

Public Sub ClearPort()
  WirelessPort.CommFlush
End Sub

'Public Sub ResetPort()
' for future use if we need to close and reopen main port

'  If MASTER Then
'    WirelessPort.CommFlush
'    CloseComm WirelessPort
'    InitComm WirelessPort, Configuration.CommPort, "" ' use defaults settings
'  End If
'
'
'End Sub

Public Function FetchData() As Long
  Dim Count         As Long
  Dim inbuffercount As Long
  Dim NewPtr        As Long
  Dim bInput        As Byte
  
  If BadPackets > MAX_BADPACKETS Then
    ClearBuffer
  End If
  
  
  inbuffercount = WirelessPort.inbuffercount
   
  'dbg "Buffer count " & inbuffercount

  For Count = 0 To inbuffercount - 1
    
    ' binput holds new byte
    WirelessPort.CommReadByte bInput
    
    If TotalBytes > 2000000 Then ' statistical counter
      TotalBytes = 1
    End If
    TotalBytes = TotalBytes + 1
    
    'Debug.Print Right("0" & Hex(bInput), 2);
    
    Buffer(EndPtr) = bInput
    NewPtr = EndPtr + 1
    If NewPtr > HighWaterMark Then
      HighWaterMark = EndPtr
      'dbg "New High Water Mark " & HighWaterMark & vbCrLf
    End If
    If NewPtr > MAXBUFFER Then
      NewPtr = 0
    End If
    If NewPtr = StartPtr Then
      ' leave in port's buffer for now
      ' raise error / buffer overrun
        Dim s As String
        s = "Bad Data " & " Buffer Overrrun " & Format(Now, "mm-dd-yy hh:nn:ss")
        SpecialLog s
        dbg " Buffer Overrrun " & Format(Now, "mm-dd-yy hh:nn:ss") & vbCrLf
      
      Exit For
    End If

    EndPtr = NewPtr
    
    
    Dim cbufferlen As Long
    If StartPtr > EndPtr Then
      cbufferlen = -1
    Else
      cbufferlen = EndPtr - StartPtr
    End If
    
    'dbg " StartPtr, EndPtr, len " & StartPtr & "  " & EndPtr & "  " & cbufferlen & vbCrLf
    ' get byte
  Next
  ' return # of bytes processed
  FetchData = Count
End Function


Private Sub Class_Initialize()
  Set Packets = New Collection
  Set i6080 = New c6080
End Sub

Private Sub Class_Terminate()
  If Not i6080 Is Nothing Then
    i6080.DisConnect
  End If
  Set i6080 = Nothing
End Sub
