VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerialInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary





Private TAP_STATE As Integer


Private TAPMessage As cSerialTAPMessage

Private mProtocol     As Integer

Public Serial         As String

Public Skip           As Long
Private mPhraseLength As Long
Public EOLChar        As Integer

Public Port           As Integer
Public Baud           As Long
Public DataBits       As Integer
Public Stopbits       As String
Public Parity         As String

Public Settings       As String

' if it's less than a space char, replace it with space char, unless it's a CR, LF , or EOL char.
Private Const MAXBUFFER = 255

Private Buffer(0 To MAXBUFFER) As Byte

Private mIncludeWords As Collection
Private mExcludeWords As Collection

Private StartPtr      As Long
Private NewPtr        As Long
Private EndPtr        As Long
Private DataLen       As Long
Private Busy          As Boolean

Public SerialPacket   As String
Public PagerNum       As String

Private mAnnounce     As String

Private mPacketReady   As Boolean

Private Const vbQuote = """"
Private Const vbSpace = " "

Private mSerialComm    As cComm


'Private NUL     As String
'Private SOH     As String
'Private STX     As String
'Private ETX     As String
'Private EOT     As String
'Private ENQ     As String
'Private ACK     As String
'Private BEL     As String
'Private BS      As String
'Private HT      As String
'Private LF      As String
'Private VT      As String
'Private FF      As String
'Private CR      As String
'Private SO      As String
'Private si      As String
'Private XOFF    As String
'Private XON     As String
'Private NAK     As String
'Private ETB     As String
'Private SUBS    As String
'Private ESC     As String
'Private rs      As String
'Private US      As String
'Private DEL     As String
''
'Private SUB_CR  As String
'Private SUB_LF  As String

Private mSerialTapProtocol As Integer


Public Function FetchData() As Long
  Dim Count         As Long
  Dim inbuffercount As Long
  Dim NewPtr        As Long
  Dim bInput        As Byte
  Dim t             As Long
  Dim j             As Long
  Dim ReadyBuffer   As String

  If mSerialComm Is Nothing Then Exit Function


  Busy = True

  If mSerialComm.PortOpen Then

    inbuffercount = mSerialComm.inbuffercount

    If inbuffercount > 0 Then
      'Stop 't = t
    End If

    If mProtocol = PROTOCOL_TAP Or mProtocol = PROTOCOL_PET Then
      If TAPMessage Is Nothing Then
        Set TAPMessage = New cSerialTAPMessage
      End If
      If TAPMessage.Timeout Then
        TAPMessage.Clear
        PagerNum = ""
        EndPtr = 0
        NewPtr = 0
        StartPtr = 0
        Erase Buffer
        SendNAK
        Exit Function
      End If
    End If


    For Count = 0 To inbuffercount - 1
      DoEvents
      mSerialComm.CommReadByte bInput

      '******** choose protocol ************
      If mProtocol = PROTOCOL_PET Then

        If TAPMessage Is Nothing Then
          Set TAPMessage = New cSerialTAPMessage
        End If

        Select Case TAPMessage.State
          Case TAP_Status.TAP_TIMEOUT, TAP_Status.TAP_Error
            TAPMessage.State = TAP_Status.TAP_WAITING
          Case TAP_Status.TAP_WAITING
            If bInput = BYTE_CR Then
              TAPMessage.State = TAP_Status.TAP_ATTENTION
              SendIDReq
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
            End If

          Case TAP_Status.TAP_ATTENTION
            If bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
          
            ElseIf bInput = BYTE_ESC Then
              TAPMessage.State = TAP_Status.TAP_LOGON
            End If
          Case TAP_Status.TAP_LOGON
            If bInput = BYTE_CR Then
              TAPMessage.State = TAP_Status.TAP_ACCEPT
              mSerialComm.CommWrite CHAR_CRLF & "110 HMC TAP In Ready" & CHAR_CR & CHAR_ACK & CHAR_CRLF & CHAR_ESC & "[p" & CHAR_CR
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
            Else
              TAPMessage.REQAddChar bInput
            End If
          Case TAP_Status.TAP_ACCEPT
            If bInput = BYTE_STX Then  ' Ignore until <STX>
              TAPMessage.State = TAP_Status.TAP_HAS_STX
              TAPMessage.BufferAddChar bInput
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
            
            End If

          Case TAP_Status.TAP_HAS_STX  ' keep accumulating until <CR>
            If bInput = BYTE_CR Then
              TAPMessage.State = TAP_Status.TAP_HAS_ID
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
            
            Else
              TAPMessage.IDAddChar bInput
            End If
            TAPMessage.BufferAddChar bInput

          Case TAP_Status.TAP_HAS_ID  ' keep accumulating until <CR>
            If bInput = BYTE_CR Then
              TAPMessage.State = TAP_Status.TAP_HAS_MSG
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
            
            Else
              TAPMessage.MessageAddChar bInput
            End If
            TAPMessage.BufferAddChar bInput

          Case TAP_Status.TAP_HAS_MSG  ' keep accumulating until <ETX>
            If bInput = BYTE_ETX Then
              TAPMessage.State = TAP_Status.TAP_HAS_ETX
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
            End If
            TAPMessage.BufferAddChar bInput

          Case TAP_Status.TAP_HAS_ETX
            If bInput = BYTE_CR Then  ' build checksum until <CR>
              TAPMessage.State = TAP_Status.TAP_HAS_CHKSUM
              ' parse it
              If TAPMessage.IsCheckSumValid() Then  ' All is Good
                SerialPacket = TAPMessage.packet
                PagerNum = TAPMessage.PagerNum
                mPacketReady = ProcessString(TAPMessage.message)
                'processTAPMessage Tapmessage.message
                mSerialComm.CommWrite TAPMessage.PagerNum & " Message Sent." & CHAR_CR & CHAR_ACK & CHAR_CR
                TAPMessage.State = TAP_SENT
            ElseIf bInput = BYTE_EOT Then
              TAPMessage.State = TAP_Status.TAP_TIMEOUT
              Else  ' bad checksum
                TAPMessage.State = TAP_Status.TAP_Error
                SendNAK
              End If
            Else
              TAPMessage.ChecksumAddChar bInput
            End If

          Case TAP_Status.TAP_SENT
            If bInput = BYTE_EOT Then
              TAPMessage.State = TAP_EOT
            End If
          Case TAP_Status.TAP_EOT
            If bInput = BYTE_CR Then  '
              mSerialComm.CommWrite "115 GoodBye." & CHAR_CR & CHAR_ESC & CHAR_EOT & CHAR_CR
              TAPMessage.Clear
              'PagerNum = ""
              EndPtr = 0
              NewPtr = 0
              StartPtr = 0
              Erase Buffer
              TAPMessage.State = TAP_Status.TAP_WAITING
              Exit For

            End If
          Case Else
            ' shouldn't be here
          

        End Select
        
        ' something about  Comm.CommWrite "ATI" & CR
        ' see SendTAP in Cpagedevice
        'RC = SendTAP(Item.Address, Item.Message)
        '<STX>ID<CR>Message<CR><ETX>Checkum<CR>


        '  Function TAPClean(ByVal Message As String) As String
        '  convert all these to spaces
        '  Message = Replace(Message, "^", " ")
        '  Message = Replace(Message, "~", " ")
        '  Message = Replace(Message, "_", " ")
        '  Message = Replace(Message, CR, SUB_CR)
        '  Message = Replace(Message, LF, SUB_LF)
        '  TAPClean = Message
        '  End Function

        'look for last byte and return ACK char



      ElseIf mProtocol = PROTOCOL_TAP Then
        If TAPMessage Is Nothing Then
          Set TAPMessage = New cSerialTAPMessage
        End If

        Select Case TAPMessage.State
          Case TAP_TIMEOUT

          Case TAP_WAITING
            If bInput = BYTE_STX Then  ' Ignore until <STX>
              TAPMessage.State = TAP_HAS_STX
              TAPMessage.BufferAddChar bInput
            End If

          Case TAP_HAS_STX  ' keep accumulating until <CR>
            If bInput = BYTE_CR Then
              TAPMessage.State = TAP_HAS_ID
            Else
              TAPMessage.IDAddChar bInput
            End If
            TAPMessage.BufferAddChar bInput

          Case TAP_HAS_ID  ' keep accumulating until <CR>
            If bInput = BYTE_CR Then
              TAPMessage.State = TAP_HAS_MSG
            Else
              TAPMessage.MessageAddChar bInput
            End If
            TAPMessage.BufferAddChar bInput

          Case TAP_HAS_MSG  ' keep accumulating until <ETX>
            If bInput = BYTE_ETX Then
              TAPMessage.State = TAP_HAS_ETX
            End If
            TAPMessage.BufferAddChar bInput

          Case TAP_HAS_ETX
            If bInput = BYTE_CR Then  ' build checksum until <CR>
              TAPMessage.State = TAP_HAS_CHKSUM
              ' parse it
              If TAPMessage.IsCheckSumValid() Then  ' All is Good
                SerialPacket = TAPMessage.packet
                PagerNum = TAPMessage.PagerNum
                mPacketReady = ProcessString(TAPMessage.message)
                'processTAPMessage Tapmessage.message
                SendACK
              Else  ' bad checksum
                SendNAK
              End If
              TAPMessage.Clear
              'PagerNum = ""
              EndPtr = 0
              NewPtr = 0
              StartPtr = 0
              Erase Buffer
              Exit For
            Else
              TAPMessage.ChecksumAddChar bInput
            End If

        End Select
        ' something about  Comm.CommWrite "ATI" & CR
        ' see SendTAP in Cpagedevice
        'RC = SendTAP(Item.Address, Item.Message)
        '<STX>ID<CR>Message<CR><ETX>Checkum<CR>


        '  Function TAPClean(ByVal Message As String) As String
        '  convert all these to spaces
        '  Message = Replace(Message, "^", " ")
        '  Message = Replace(Message, "~", " ")
        '  Message = Replace(Message, "_", " ")
        '  Message = Replace(Message, CR, SUB_CR)
        '  Message = Replace(Message, LF, SUB_LF)
        '  TAPClean = Message
        '  End Function

        'look for last byte and return ACK char



      Else
        Select Case bInput
          Case &HD, &HA, EOLChar
            If EndPtr > 0 Then
              ReadyBuffer = left$(StrConv(Buffer, vbUnicode), EndPtr)
              SerialPacket = ReadyBuffer  ' the "Packet" that we're using
              mPacketReady = ProcessString(ReadyBuffer)
            End If
            SerialPacket = ""
            PagerNum = ""
            EndPtr = 0
            NewPtr = 0
            StartPtr = 0
            Erase Buffer
            Exit For
          Case Else
            If bInput < &H20 Then  ' no control chars
              bInput = &H20
            ElseIf bInput > &H7E Then  ' no high chars
              bInput = &H20
            End If
        End Select
      End If

      Buffer(EndPtr) = bInput
      NewPtr = EndPtr + 1

      If NewPtr > MAXBUFFER Then  ' slide em over
        'Debug.Print "Buffer Before "; StrConv(Buffer, vbUnicode)
        For j = 0 To MAXBUFFER - 1
          Buffer(j) = Buffer(j + 1)
        Next
        Buffer(MAXBUFFER) = 0
        'Debug.Print "Buffer After "; StrConv(Buffer, vbUnicode)
      Else
        EndPtr = NewPtr
      End If
    Next

    ' return # of bytes processed
    FetchData = Count
  End If
  Busy = False

End Function

Sub SendIDReq()
  mSerialComm.CommWrite "ID="
End Sub


Sub SendACK()
  mSerialComm.CommWriteByte BYTE_ACK
End Sub
Sub SendNAK()
  mSerialComm.CommWriteByte BYTE_NAK
End Sub

Public Function start() As Long
  start = OpenComm(Port, Settings)
End Function

Public Function OpenComm(ByVal Port As Integer, ByVal Settings As String)
  
  
  If mSerialComm Is Nothing Then
    Set mSerialComm = New cComm
  Else
    mSerialComm.CommFlush
    mSerialComm.CommClose
  End If
  
  Sleep 100
  
  OpenComm = mSerialComm.CommOpen(Port, Settings)

End Function
Public Function CloseComm()
  If Not mSerialComm Is Nothing Then
    mSerialComm.CommClose
  End If
End Function


Public Function PacketReady() As Boolean
  PacketReady = mPacketReady
End Function

Public Function GetPacket() As cESPacket
  Dim packet As cESPacket

  Set packet = New cESPacket

  packet.Alarm = 1
  packet.Serial = Serial
  packet.DateTime = Now
  packet.PacketType = 1
  packet.SerialPacket = SerialPacket


  If mProtocol = PROTOCOL_TAP Or mProtocol = PROTOCOL_PET Then
    packet.SerialMessage = Announce ' Trim$(PagerNum & " ") & Announce
  Else
    packet.SerialMessage = Announce
  End If
  packet.SetMIDPTI 0, 255

  Set GetPacket = packet
  Set packet = Nothing
  ' clear ready for next event
  mPacketReady = False
  SerialPacket = ""

  'packet.Tamper = mSerialComm.DSR ' enable for trouble screen

End Function
'Function ProcessString() As String
'  Dim StringBuffer      As String
'  Dim StringBuffArray() As String
'  Dim BuffLen           As Long
'  Dim Word              As String
'  Dim ExcludeWord       As String
'  Dim j                 As Long
'  Dim i                 As Integer
'  Dim K                 As Integer
'  Dim ptr               As Long
'  Dim ArrayStart        As Integer
'  Dim ArrayEnd          As Integer
'  Dim ArrayElement      As String
'  Dim Phrase            As String
'  Dim Match             As Boolean
'
'  BuffLen = EndPtr
'
'  If BuffLen > 0 Then
'    StringBuffer = Left(StrConv(Buffer, vbUnicode), BuffLen)
'    SerialPacket = StringBuffer
'    StringBuffArray = Split(StringBuffer, " ")
'    ArrayStart = LBound(StringBuffArray)
'    ArrayEnd = UBound(StringBuffArray)
'    If ArrayStart <= ArrayEnd Then
'
'
'
'      For i = ArrayStart To ArrayEnd
'
'        Match = False
'        DoEvents
'
'          For j = 1 To mIncludeWords.Count
'          ArrayElement = StringBuffArray(i)
'          Word = mIncludeWords(j)
'          If InStr(Word, "*") > 0 Then
'            If ArrayElement Like Word Then
'              Match = True
'              Exit For
'            End If
'          Else
'            ptr = InStr(1, ArrayElement, Word, vbBinaryCompare)
'
'            If ptr = 1 Then
'              Match = True
'              Exit For
'              ''            If Skip < BuffLen Then
'              ''              If Skip = 0 Then
'              ''                Phrase = StringBuffer
'              ''              Else
'              ''                Phrase = MID(StringBuffer, Skip + 1) ' convert to 1 based
'              ''              End If
'              ''              Announce = Left(Phrase, PhraseLength)
'              ''              mPacketReady = True
'              ''              ProcessString = Word
'              ''              Exit Function
'              ''            Else
'              ''              Announce = ""
'              ''              mPacketReady = True
'              ''              ProcessString = Word
'              ''              Exit Function
'              ''            End If
'            End If
'
'          End If
'        Next
'        If (Match) Then
'          ' try  excludes
'          For K = 1 To mExcludeWords.Count
'            ExcludeWord = mExcludeWords(K)
'            ptr = InStr(1, ArrayElement, ExcludeWord, vbBinaryCompare)
'            If ptr > 0 Then
'               Match = False
'               Exit For
'            End If
'          Next
'
'          ' else, return alarm
'          If (Match) Then
'          If Skip < BuffLen Then
'            If Skip = 0 Then
'              Phrase = StringBuffer
'            Else
'              Phrase = MID(StringBuffer, Skip + 1)  ' convert to 1 based
'            End If
'            Announce = Left(Phrase, PhraseLength)
'            mPacketReady = True
'            ProcessString = Word
'            Exit For
'          Else
'            Announce = ""
'            mPacketReady = True
'            ProcessString = Word
'            Exit For
'          End If
'          End If
'        End If
'      Next
'    End If
'  End If
'
'End Function


Private Function ProcessString(ByVal StringBuffer As String) As Boolean
  'Dim StringBuffer      As String
  Dim StringBuffArray() As String
  Dim BuffLen           As Long
  Dim Word              As String
  Dim ExcludeWord       As String
  Dim j                 As Long
  Dim i                 As Integer
  Dim k                 As Integer
  Dim Ptr               As Long
  Dim ArrayStart        As Integer
  Dim ArrayEnd          As Integer
  Dim ArrayElement      As String
  Dim Phrase            As String
  Dim Match             As Boolean
  
  Dim TAPMessage        As cSerialTAPMessage

  Dim includewords()    As String

  Dim debugthis         As Boolean
  On Error GoTo ProcessString_Error

  debugthis = False

  Announce = ""
  'BuffLen = EndPtr

  If Len(StringBuffer) > 0 Then
    'StringBuffer = left$(StrConv(Buffer, vbUnicode), BuffLen)  ' convert from bytes to string
    BuffLen = Len(StringBuffer)
    
    StringBuffArray = Split(StringBuffer, " ")
    ArrayStart = LBound(StringBuffArray)
    ArrayEnd = UBound(StringBuffArray)


    If debugthis Then
      dbg "***"
      dbg "Buffer WordCount " & StringBuffer & ", " & ArrayEnd
    End If

    If ArrayStart <= ArrayEnd Then

      includewords = Split(StringBuffer, " ")

      For i = ArrayStart To ArrayEnd  ' do the whole string
        Match = False
        ArrayElement = StringBuffArray(i)
        If debugthis Then
          dbg "Word # " & ArrayElement & ", " & i
        End If
        'includewords
        For j = 1 To mIncludeWords.Count
          ' word is the qualifier word to look for
          Word = mIncludeWords(j)
          If debugthis Then
            dbg "Include Word # " & Word & ", " & j
          End If

          If InStr(Word, "*") > 0 Then  ' wildcard match
            If ArrayElement Like Word Then
              If debugthis Then
                dbg "Wildcard Match " & Word
              End If
              Match = True
              Exit For
            End If

          Else  ' exact match only
            Ptr = InStr(1, ArrayElement, Word, vbBinaryCompare)
            If Ptr = 1 Then
              If debugthis Then
                dbg "Exact Match " & Word
              End If
              Match = True
              Exit For  ' exit for j
            End If
          End If
        Next j

        If (Match) Then  ' we're done for now...
          Exit For  ' exit for i
        End If
      Next i

      If debugthis Then
        If Match Then
          dbg "We have a match and try to exclude it next"
        Else
          dbg "No match We're done"
        End If
  
      End If

      If Match Then  ' see if we can exclude it
        For i = ArrayStart To ArrayEnd
          ArrayElement = StringBuffArray(i)
          For k = 1 To mExcludeWords.Count
            ExcludeWord = mExcludeWords(k)
            Ptr = InStr(1, ArrayElement, ExcludeWord, vbBinaryCompare)
            If Ptr > 0 Then
              If debugthis Then
                dbg "We're using the exclude word '" & ExcludeWord & "'"
              End If
              Match = False
              Exit For
            End If
          Next k
          If Match = False Then  ' then we've excluded it
            Exit For
          End If
        Next i
      End If
      ' else, return alarm

      If (Match) Then
        If debugthis Then
          dbg "The match remains, process it"
        End If

        If Skip < BuffLen Then
          If Skip = 0 Then
            Phrase = StringBuffer
          Else
            Phrase = MID(StringBuffer, Skip + 1)  ' convert to 1 based
          End If
          Announce = left(Phrase, PhraseLength)
          'mPacketReady = True
          'ProcessString = Word
          ProcessString = True
        Else
          Announce = ""
          'Announce = left(StringBuffer, PhraseLength)
          'mPacketReady = True
          ProcessString = True
        End If
      Else
        If debugthis Then
          dbg "Return a no-match condition"
        End If
        
        
      End If

    End If
  End If

ProcessString_Resume:
  On Error GoTo 0
  Exit Function

ProcessString_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cSerialInput.ProcessString." & Erl
  Resume ProcessString_Resume


End Function

Public Sub SetWords(ByVal s As String)
  Dim Words() As String
  Dim j       As Long

  Set mIncludeWords = New Collection
  Words = Split(s, " ")
  For j = LBound(Words) To UBound(Words)
     Words(j) = Words(j)
     's = Trim(s)
     If Len(Words(j)) Then
      mIncludeWords.Add Words(j)
    End If
  Next
  
End Sub
Public Sub SetExclude(ByVal s As String)
  Dim Words() As String
  Dim j       As Long

  Set mExcludeWords = New Collection
  Words = Split(s, " ")
  For j = LBound(Words) To UBound(Words)
     Words(j) = Words(j)
     's = Trim(s)
     If Len(Words(j)) Then
      mExcludeWords.Add Words(j)
    End If
  Next
  
End Sub


Private Sub Class_Initialize()
'  NUL = vbNullChar 'Chr$(0)
'  SOH = Chr$(1)
'  STX = Chr$(2)
'  ETX = Chr$(3)
'  EOT = Chr$(4)
'  ENQ = Chr$(5)
'  ACK = Chr$(6)
'  BEL = Chr$(7)
'  BS = vbBack
'  HT = vbTab
'  LF = vbLf
'  VT = vbVerticalTab
'  FF = vbFormFeed
'  CR = vbCr
'  SO = Chr$(&HE)
'  si = Chr$(&HF)
'  XOFF = Chr$(&H11)
'  XON = Chr$(&H13)
'  NAK = Chr$(&H15)
'  ETB = Chr$(&H17)
'  SUBS = Chr$(&H1A)
'  ESC = Chr$(&H1B)
'  rs = Chr$(&H1E)
'  US = Chr$(&H1F)
'  DEL = Chr$(&H7F)
'  '
'  SUB_CR = SUBS & "M"
'  SUB_LF = SUBS & "J"
  
  Set mSerialComm = New cComm
End Sub

Private Sub Class_Terminate()
  ' exit gracefully
  mSerialComm.CommFlush
  mSerialComm.CommClose
  'Wait 100 ' 100 ms for port to close
  Set mSerialComm = Nothing
End Sub

Public Property Get protocol() As Integer
  protocol = mProtocol
End Property

Public Property Let protocol(ByVal Value As Integer)
  mProtocol = Value
End Property

Public Property Get SerialTapProtocol() As Integer

  SerialTapProtocol = mSerialTapProtocol

End Property

Public Property Let SerialTapProtocol(ByVal Value As Integer)
  If Value = 1 Then
    protocol = PROTOCOL_TAP
  ElseIf Value = 2 Then
    protocol = PROTOCOL_PET
  Else
    protocol = 0
  End If
  mSerialTapProtocol = Value

End Property

Public Property Get Announce() As String
  Announce = mAnnounce
End Property

Public Property Let Announce(ByVal Announce As String)
  mAnnounce = Announce
End Property

Public Property Get PhraseLength() As Long

  PhraseLength = mPhraseLength

End Property

Public Property Let PhraseLength(ByVal PhraseLength As Long)

  mPhraseLength = PhraseLength

End Property
