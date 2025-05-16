VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDivaCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DIVALine As Object ' divacallhandler/singlecallhandler

Private mTerminateDigit As Long
Public DigitReceived    As Boolean

Public Channel          As Long
Public Device           As Long
Public index            As Long

Private mhCall          As Long

Private mBusy           As Boolean

Private SpeechFilename As String

Public Status As Long
Public StatusString As String

Private Enum LineStatus
  STAT_IDLE = 0
  STAT_NORMALDISCONNECT
  STAT_STARTED
  STAT_BUSY
  STAT_nodialtone
  STAT_REDIALING
  STAT_CANCELLED
  STAT_PROCEEDING
  STAT_DIALING
  STAT_CONNECTED
  STAT_SPEAKING
  STAT_SPEAKDELAY
  STAT_DONESPEAKING
  STAT_REPEATING
  STAT_UNK '= 99
  STAT_NOANSWER = 200
End Enum

Public Event OnDone()

Public Pretext            As String ' extra prefix to message

Public PhoneNumber        As String ' obvious
Private mMessage          As String ' system generated message
Public MessageDelay       As Long   ' seconds to wait before blindly speaking message

Public MessageRepeats     As Long  ' how many times to repeat message
Public MessageRepeatDelay As Long  ' seconds

Public RedialAttempts     As Long  ' how many times to redial if busy, or no dial tone
Public RedialWait         As Long  ' seconds

Private mTimeout          As Long   ' seconds: general failure timer

'Public TAPIProviders        As Collection ' list of modems available

Private WithEvents DIVALine As DIVASDKLib.DivaCall ' diva call object
Attribute DIVALine.VB_VarHelpID = -1
'Private WithEvents DIVALine As DIVASDKLib.DivaCall ' diva call object

Public SpeechEngine As cSpeechEngine

Private voice As spvoice


Private mPermanentLineID  As Long   ' unique hardware ID for voice modem

'Private Lasttime          As Double
Private Elapsed           As Double
Private starttime         As Double

Private MediaID           As Long   ' channel to communicate from speech to modem

Private MessageDelayTimer   As Long  ' timer for connect-dialing-speak message
Private MessageRepeatTimer  As Long  ' timer between message repeats
Private MessageCounter      As Long  ' number of messages repeats
Private RedialCounter       As Long  ' number of redial attempts
Private RedialWaitTimer     As Long  ' timer between redial attempts
'Private DialtoneTimer       As Long  ' maximum time to wait for dialtone
Private TimeoutTimer        As Long  ' general failure timer

'Private mActive             As Boolean '

Private Speaking            As Boolean
Public Function SetDivaLine(Diva As DIVASDKLib.DivaCall) As Long
  Set DIVALine = Diva
End Function



Public Sub ResetClock()
  starttime = Timer
End Sub
Public Sub UpdateClock()
  'called once a second from cpagedevice (which is called from the master clock)
  Dim CurrentTime As Double
  
  CurrentTime = Timer
  If starttime > CurrentTime Then
    'Debug.Print "Rollover"
    Elapsed = (86400 - starttime) + CurrentTime
  Else
    Elapsed = CurrentTime - starttime
  End If
  If Elapsed >= 1 Then
    starttime = Timer
    CheckStatus
  End If
  
  
  'Debug.Print "Elapsed = " & Format(Elapsed, "0.00")
End Sub

Public Property Get Active() As Boolean
  
  Active = Not (DIVALine Is Nothing) ' inverted logic?
  
End Property

Sub SignalDone()
  On Error Resume Next
  Kill SpeechFilename
  Busy = False
  RaiseEvent OnDone
End Sub

Private Function Status2Text(ByVal Status As Long) As String
  Select Case Status
  Case STAT_IDLE: Status2Text = "Idle"
  Case STAT_NORMALDISCONNECT: Status2Text = "Normal Disc"
  Case STAT_STARTED: Status2Text = "Started"
  Case STAT_BUSY: Status2Text = "Busy"
  Case STAT_nodialtone: Status2Text = "No Dialtone"
  Case STAT_REDIALING: Status2Text = "Redialing"
  Case STAT_CANCELLED: Status2Text = "Cancelled"
  Case STAT_PROCEEDING: Status2Text = "Proceeding"
  Case STAT_DIALING: Status2Text = "Dialing"
  Case STAT_CONNECTED: Status2Text = "Connected"
  Case STAT_SPEAKING: Status2Text = "Speaking"
  Case STAT_SPEAKDELAY: Status2Text = "Speak Delay"
  Case STAT_DONESPEAKING: Status2Text = "Done Speaking"
  Case STAT_REPEATING: Status2Text = "Repeating"
  Case STAT_NOANSWER = 200: Status2Text = "No Answer"
  Case Else: Status2Text = "UNK"
    'STAT_UNK '= 99

  End Select

End Function

'Public Property Get Voices()
'  Set Voices = SpeechEngine.Voices
'End Property

Sub CheckStatus()
  'todo
  If Status <> STAT_IDLE Then
    dbgTAPI "cDivaCall.CheckStatus Channel: " & index & " " & Status2Text(Status)
    
  End If
   
'    Debug.Print "DivaCall.CheckStatus Disconnected reason" & DIVALine.DisconnectReason
   
  If TimeoutTimer <= 0 And Status <> STAT_IDLE Then
    HangUp
    Status = STAT_IDLE
    dbgTAPI "cDivaCall.CheckStatus Timeout Error Line: " & index
    SignalDone
  
    
  
  Else
    Select Case Status
      Case STAT_REPEATING
        Debug.Print "STAT_REPEATING Line: " & index
      Case STAT_SPEAKING
        Debug.Print "STAT_SPEAKING Line: " & index
      Case STAT_DONESPEAKING
        Debug.Print "STAT_DONESPEAKING Line: " & index
      Case STAT_REDIALING
          Debug.Print "STAT_REDIALING Line: " & index
      Case STAT_SPEAKDELAY
        ' don't decrement TimeoutTimer
      Case Else
        If TimeoutTimer > 0 Then
          TimeoutTimer = TimeoutTimer - 1
        End If
    End Select
  End If

  Select Case Status
    Case STAT_IDLE
      ' nothing to do
      'Debug.Print "STAT_IDLE"
      If mTerminateDigit Then
         'Debug.Print " STAT_IDLE, mTerminateDigit, DigitReceived, RedialCounter " & mTerminateDigit & "  " & DigitReceived & "  " & RedialCounter
        If DigitReceived = False And (RedialCounter > 0) Then
            RedialCounter = RedialCounter - 1
            RedialWaitTimer = RedialWait
            Status = STAT_REDIALING
            TimeoutTimer = Timeout
             'Debug.Print "IDLE - > STAT_REDIALING  Line: " & Index
           Status = STAT_REDIALING
           Exit Sub
        End If
      End If
         


    Case STAT_NORMALDISCONNECT
       'Debug.Print "STAT_NORMALDISCONNECT  Line: " & Index
      
      ' new gather digits redial if not found
      If mTerminateDigit <> 0 Then  'STAT_NOANSWER = 200
          ' Debug.Print "DIGIT NOT RECEIVED "
          
          If RedialCounter > 0 Then
            RedialCounter = RedialCounter - 1
            RedialWaitTimer = RedialWait
            Status = STAT_REDIALING
           '  Debug.Print "STAT_REDIALING " & RedialWaitTimer
          Else
            ' Debug.Print "REDIALS EXHAUSTED STAT_NORMALDISCONNECT"
            Status = STAT_IDLE
            SignalDone
          End If
         
      Else
        ' Debug.Print "DIGIT NOT NEEDED "
        Status = STAT_IDLE
        SignalDone
      End If

    Case STAT_REDIALING
       Debug.Print "STAT_REDIALING " & RedialWaitTimer
      If RedialWaitTimer <= 0 Then
        MakeCallAsynch
      Else
        RedialWaitTimer = RedialWaitTimer - 1
      End If

    Case STAT_BUSY
       Debug.Print "STAT_BUSY " & RedialCounter
       
      If RedialCounter > 0 Then
        RedialCounter = RedialCounter - 1
        RedialWaitTimer = RedialWait
        Status = STAT_REDIALING
      Else
        DIVALine.DisConnect
        Status = STAT_IDLE
        SignalDone
      End If

    Case STAT_nodialtone
       Debug.Print "STAT_NODIALTONE " & RedialCounter
      If RedialCounter > 0 Then
        RedialCounter = RedialCounter - 1
        RedialWaitTimer = RedialWait
        Status = STAT_REDIALING
      Else
        DIVALine.DisConnect
        Status = STAT_IDLE
        SignalDone
      End If

    Case STAT_CANCELLED
       Debug.Print "STAT_CANCELLED"
      Status = STAT_IDLE
      SignalDone

    Case STAT_PROCEEDING
      ' OK
       Debug.Print "STAT_PROCEEDING"
    Case STAT_DIALING
      ' OK
       Debug.Print "STAT_DIALING"

    Case STAT_CONNECTED
       Debug.Print "STAT_CONNECTED"
      MessageDelayTimer = MessageDelay
      MessageCounter = MessageRepeats
      Status = STAT_SPEAKDELAY
      If MessageDelayTimer <= 0 Then
        Status = STAT_SPEAKING
        PlayMessage Pretext & " " & message
      Else
        MessageDelayTimer = MessageDelayTimer - 1
      End If


    Case STAT_SPEAKING
      'dbgTAPI "STAT_SPEAKING"

    Case STAT_SPEAKDELAY

      dbgTAPI "STAT_SPEAKDELAY Seconds " & MessageDelayTimer & " Remaining"
      If MessageDelayTimer <= 0 Then
        ' ???here?? MessageCounter = MessageCounter - 1
        'MessageCounter = MessageCounter - 1
        Status = STAT_SPEAKING
        PlayMessage Pretext & " " & message

      Else
        MessageDelayTimer = MessageDelayTimer - 1
      End If

    Case STAT_DONESPEAKING
       dbgTAPI "STAT_DONESPEAKING " & MessageCounter & " Repeats Pending"
      MessageRepeatTimer = MessageRepeatDelay

      If MessageCounter > 0 Then
        MessageCounter = MessageCounter - 1
        Status = STAT_REPEATING
      Else
        HangUp
        Status = STAT_IDLE
        SignalDone
      End If

    Case STAT_REPEATING
      dbgTAPI "STAT_REPEATING in " & MessageRepeatTimer & " Seconds"
      If MessageRepeatTimer <= 0 Then
        Status = STAT_SPEAKING
        PlayMessage Pretext & " " & message
      Else
        MessageRepeatTimer = MessageRepeatTimer - 1
      End If

    Case STAT_STARTED
      'Debug.Print "STAT_STARTED " & Index '  & Device & "," & Channel

    Case Else  ' STAT_UNK
       'Debug.Print "STAT_UNK " & Status

  End Select

  
End Sub
Public Function BeginCall() As Long
'todo
  Busy = True

  
  dbgTAPI "DIVACALL.BeginCall"
  Status = STAT_STARTED
  
  TimeoutTimer = Timeout
  RedialCounter = RedialAttempts
  MessageCounter = MessageRepeats
  
  
  DIVALine.DisConnect

  DIVALine.ClearDetectedDigits

  If mTerminateDigit <> 0 Then
    
    'DIVALine.GetDigits 1, 0, 0, 0, mTerminateDigit
  Else
    'DIVALine.DisableDigitDetection
    
    
  End If
  'DIVALine.DisableAMD
  
  MakeCallAsynch


End Function

Private Function MakeCallAsynch() As Long

  Dim rc As Long
  
'  m_LastTAPIError = 0
'  m_RequestingCall = 0
  
  Dim tstart As Date
  
  Dim i As Long
  
  tstart = DateAdd("s", 10, Now)
  
  
  HangUp
  
  Do While mhCall <> 0
    
    If i > 10000 Then
      i = 0
      DoEvents
    End If
    i = i + 1
    
    If (Now > tstart) Then
      Exit Do
    End If
    
  Loop
  
  mhCall = 0

' **************** new
  ' TAPI Way:     RC = lineMakeCall(mhLine, mhCall, PhoneNumber, 0&, ByVal 0& )
  ' mhCall is changed by linemakecall
  TimeoutTimer = Timeout
  
  DigitReceived = False
  DIVALine.ClearDetectedDigits
  
  DIVALine.AsyncMode = True
  DIVALine.SignalEvents = True
  DIVALine.Device = Device
  DIVALine.Channel = Channel
  If mTerminateDigit <> 0 Then
    DIVALine.EnableDigitDetection = True
  Else
    DIVALine.EnableDigitDetection = False
  End If
  
  'divaline.DialingComplete
  
  'DIVALine.DisableAMD =DivaResultAMDAnsweringMachine
  
'  DIVALine.EnableExtendedToneDetection = True
  
  Dim bc
  bc = DIVALine.EnableAMD(3000, 5000, 1000)
  dbgTAPI "DIVALine.EnableAMD " & rc
  
  
  DIVALine.ClearDetectedDigits
  DIVALine.Connect PhoneNumber, DivaCallType_Voice
  mhCall = 1
  
  dbgTAPI "cDivaCall.MakeCallAsynch " & Now & " on line " & index & " ph#" & PhoneNumber

End Function

Public Function PlayMessage(ByVal text As String) As Long
  'todo
  'DIVALine.SendVoiceFile SpeechFilename, DivaAudioDefault
  
  If Not FileExists(SpeechFilename) Then
    dbgTAPI "Recreating Voice File Line: " & index
    Call SaveSpeech2File(text, SpeechFilename)
    
  End If
  dbgTAPI "DIVALine Sending speech file (playmessage) Line: " & index
  DIVALine.SendVoiceFile SpeechFilename, DivaAudioAutodetect
End Function

'Function SendVoiceFile(ByVal filename As String)
'  DIVALine.SendVoiceFile "8x8.wav", DivaAudioAutodetect
'  Status = STAT_SPEAKING
'End Function


Public Function CancelCall() As Boolean
  'todo
  DIVALine.DisConnect
End Function

Public Function HangUp() As Boolean
  'todo
  dbgTAPI "DivaCall.HangUp"
  DIVALine.DisConnect
End Function


Private Sub Class_Initialize()
  
  
  dbgTAPI "DivaCall.Init"
  
  Set voice = New spvoice
 ' voice.EventInterests = SVEAllEvents
  'Set f = New SpAudioFormat
  'f.Type = SAFT8kHz8BitMono
 ' voice.AudioOutputStream.Format.Type = SAFT8kHz8BitMono
 ' voice.AllowAudioOutputFormatChangesOnNextSet = False
  Init
  starttime = Timer
  UpdateClock
 
End Sub

Private Sub Class_Terminate()
  Finalize
  dbgTAPI "DivaCall.Terminate"
End Sub

Public Property Get VoiceName() As String
  VoiceName = voice.voice.GetDescription()
End Property

Public Property Let VoiceName(ByVal Value As String)
  
  

  Dim token As ISpeechObjectToken  ' token to hold the registry voice tokens

  If Len(Value) > 0 Then  ' else use default voice
    For Each token In voice.GetVoices
      'Debug.Print "Voice "; Token.GetDescription()
      If InStr(1, token.GetDescription(), Value, vbTextCompare) > 0 Then
        Set voice.voice = token
        Debug.Print "c DivaCall Voice changed to: " & token.GetDescription()
        Exit For
      End If
    Next
  End If
  
  'Set token =
  'SetVoice = token.GetDescription()
End Property

Sub Init()
  Timeout = 60
  DigitReceived = False
'  If TapiLine Is Nothing Then
'    Set TapiLine = New CTAPILine
'    'TapiLine.OpenLine
'  End If
'  If SpeechEngine Is Nothing Then
'    Set SpeechEngine = New cSpeechEngine
'    SpeechEngine.Init
'    SpeechEngine.Volume = 100
'    'dbgDIVA "SpeechEngine = " & IIf(SpeechEngine Is Nothing, "DEAD", "OK")
'  Else
'    'dbgDIVA "SpeechEngine OK"
'  End If
  EnumerateLines ' for pickin and choosin
  
  
  'SetDIVALineByPermanentLineID PermanentLineID
 
End Sub

'Function SetDIVALineByPermanentLineID(ByVal ID As Long) As Long
  ' todo
  'Dim Device    As Long
  'Dim Channel   As Long
'  Set DIVALine = DialogicSystem.CreateLine(ID)
'End Function

Public Sub EnumerateLines()
  ' mebbe nothin to do here
End Sub

Public Sub Finalize()
  'Debug.Print "cDivaCall Finalize " & Now
  On Error Resume Next
  If Not DIVALine Is Nothing Then
    DIVALine.DisConnect
  End If
  Set voice = Nothing
  'Set SpeechEngine = Nothing
  Set DIVALine = Nothing
End Sub

Public Property Get DeviceName() As String
  DeviceName = DIVALine.LineName
End Property
'Public Property Get VoiceName() As String
'  VoiceName = SpeechEngine.VoiceName
'End Property
'
'Public Property Let VoiceName(ByVal Value As String)
Private Function SetVoiceName(ByVal VoiceName As String) As String

  Dim token As ISpeechObjectToken  ' token to hold the registry voice tokens

  If Len(VoiceName) > 0 Then  ' else use default voice
    For Each token In voice.GetVoices
      'Debug.Print "Voice "; Token.GetDescription()
      If InStr(1, token.GetDescription(), VoiceName, vbTextCompare) > 0 Then
        Set voice.voice = token
        dbg "Voice changed to: " & token.GetDescription()
        Exit For
      End If
    Next
  End If
  SetVoiceName = voice.voice.GetDescription()


End Function

Public Property Get TerminateDigit() As Long

  TerminateDigit = mTerminateDigit

End Property

Public Property Let TerminateDigit(ByVal Value As Long)
  'a setting of 0 turns off Monitor Digits
  ' it will be passed to the call upon "connected"
  mTerminateDigit = Value

End Property

Public Property Get PermanentLineID() As Long
  PermanentLineID = mPermanentLineID
End Property

Public Property Let PermanentLineID(ByVal Value As Long)
'  mPermanentLineID = SetTAPILineByPermanentLineID(Value)
  
End Property



Public Property Get message() As String

  message = mMessage

End Property

Public Property Let message(ByVal text As String)
  Dim rc As Long
  Dim NeededFolder As String
  NeededFolder = App.Path & "\TTS"
  mMessage = text
  
  
  If Not (DirExists(NeededFolder)) Then
    MkDir NeededFolder
  End If
  
  dbgTAPI "DivaCall."
  
  SpeechFilename = NeededFolder & "\TTS" & Format(Now, "MMDD") & Win32.timeGetTime & Format(Me.index, "00") & ".wav"

  rc = SaveSpeech2File(Pretext & " " & text, SpeechFilename)
  
  
End Property

Function SaveSpeech2File(ByVal text As String, filename As String) As Long
        Dim cpFileStream As SpFileStream
        
        dbgTAPI "DivaCall.SaveSpeech2File " & filename
        
10       On Error GoTo SaveSpeech2File_Error

20      Set cpFileStream = New SpFileStream
30      cpFileStream.Format.Type = SAFT8kHz8BitMono ' SAFTADPCM_8kHzMono '
40      voice.AllowAudioOutputFormatChangesOnNextSet = False
50      cpFileStream.Open filename, SSFMCreateForWrite, False
        
60      Set voice.AudioOutputStream = cpFileStream
        text = Trim$(text)
70      voice.Speak text, SVSFlagsAsync
80      Do
90        DoEvents
          'Debug.Print "speak"
100     Loop Until voice.WaitUntilDone(1)
110     cpFileStream.Close
120     Set cpFileStream = Nothing

130      On Error GoTo 0
140      Exit Function

SaveSpeech2File_Error:

150     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSpeech2File of Class Module cDivaCall " & Erl
End Function


Private Sub DIVALine_OnAMDFinished(ByVal result As DIVASDKLib.DivaAMDResult)
  dbgTAPI "*** DIVALine_OnAMDFinished " & TranslateAMDResult(result)

  Select Case result
    Case DivaAMDUserTerminated

    Case DivaAMDHumanTalker

    Case DivaAMDAnsweringMachine

    Case DivaAMDAnsweringMachineTone

    Case DivaAMDSilence

    Case DivaAMDFaxOrModem
      HangUp
    Case Else
    
  End Select

End Sub
Private Function TranslateAMDResult(ByVal result As Long) As String
  Select Case result
    Case DivaAMDUserTerminated
      TranslateAMDResult = "User Terminated"
    Case DivaAMDHumanTalker
      TranslateAMDResult = "Human Talker"
    Case DivaAMDAnsweringMachine
      TranslateAMDResult = "Answering Machine Detected"
    Case DivaAMDAnsweringMachineTone
      TranslateAMDResult = "Answering Machine Tone"
    Case DivaAMDSilence
      TranslateAMDResult = "Silence"
    Case DivaAMDFaxOrModem
      TranslateAMDResult = "Fax Or Modem"
    Case Else
      TranslateAMDResult = "Unknown"
  End Select

End Function


Private Sub DIVALine_OnCallProgress(ByVal State As Long)
  
  dbgTAPI "DIVALine_OnCallProgress  " & CallProgressToText(State)
  
  Select Case State
  Case DivaCallState_Idle

  Case DivaCallState_Listening

  Case DivaCallState_Connecting
    Status = STAT_STARTED
  Case DivaCallState_Ringing
    
  'Case DivaCallState_Offering

  'Case DivaCallState_Alerting

  Case DivaCallState_Connected

  'Case DivaCallState_OnHold

  Case DivaCallState_Disconnecting

  Case DivaCallState_Disconnected

  Case DivaCallState_Proceeding
      Status = STAT_STARTED

  End Select
'  StatusString = "DIALING"
'  'dbg  StatusString
'  Status = STAT_DIALING
End Sub

Private Sub DIVALine_OnConnected()
  dbgTAPI "*** DIVALine_OnConnected Line: " & index
  

  StatusString = "CONNECTED"
  Status = STAT_CONNECTED

  CheckStatus

End Sub
Function TranslateDiconnectReasonCode(ByVal Reason As Long) As String
  Dim s As String
  Select Case Reason
  Case DivaDiscReasonUnspecified
    s = "DivaDiscReasonUnspecified"
    
  Case DivaDiscReasonNormalClearing
    s = "DivaDiscReasonNormalClearing"
  Case DivaDiscReasonUserInitiated
    s = "DivaDiscReasonUserInitiated"
  Case DivaDiscReasonBusy
    s = "DivaDiscReasonBusy"
  Case DivaDiscReasonReject  ' ????
    s = "DivaDiscReasonReject"
  Case DivaDiscReasonNoAnswer
    s = "DivaDiscReasonNoAnswer"
  Case DivaDiscReasonCableError
    s = "DivaDiscReasonCableError"
  Case DivaDiscReasonUnknownNumber
    s = "DivaDiscReasonUnknownNumber"
  Case DivaDiscReasonInvalidNumber
    s = "DivaDiscReasonInvalidNumber"
  Case DivaDiscReasonNumberChanged
    s = "DivaDiscReasonNumberChanged"
  Case DivaDiscReasonIncompatibleDest
    s = "DivaDiscReasonIncompatibleDest"
  Case DivaDiscReasonNoResource
    s = "DivaDiscReasonNoResource"
  Case DivaDiscReasonNoDataChannel
    s = "DivaDiscReasonNoDataChannel"
  Case Else
    s = "DivaDiscReason UNK"
  End Select
  TranslateDiconnectReasonCode = s
End Function
Private Sub DIVALine_OnDisconnected()
  
  
 ' dbgTAPI "DIVALine_OnDisconnected Line: " & Index
  
'   If DIVALine.DisconnectReason <= 3 Then
'
'    Status = STAT_IDLE
'    End If
  mhCall = 0
  



  dbgTAPI "*** Disconnected Because " & TranslateDiconnectReasonCode(DIVALine.DisconnectReason)
  
  Select Case DIVALine.DisconnectReason
  Case DivaDiscReasonUnspecified
    Status = STAT_IDLE
    
  Case DivaDiscReasonNormalClearing
    Status = STAT_IDLE
  Case DivaDiscReasonUserInitiated
    Status = STAT_IDLE
  Case DivaDiscReasonBusy
    Status = STAT_BUSY
    TimeoutTimer = Timeout
  Case DivaDiscReasonReject  ' ????
    Status = STAT_nodialtone
  Case DivaDiscReasonNoAnswer
    Status = STAT_CANCELLED
  Case DivaDiscReasonCableError
    Status = STAT_CANCELLED
  Case DivaDiscReasonUnknownNumber
    Status = STAT_CANCELLED
  Case DivaDiscReasonInvalidNumber
    Status = STAT_CANCELLED
  Case DivaDiscReasonNumberChanged
    Status = STAT_CANCELLED
  Case DivaDiscReasonIncompatibleDest
    Status = STAT_CANCELLED
  Case DivaDiscReasonNoResource
    Status = STAT_CANCELLED
  Case DivaDiscReasonNoDataChannel
    Status = STAT_nodialtone

  End Select





End Sub

Private Sub DIVALine_OnToneReceived(ByVal cTone As Byte)
  dbgTAPI "DIVALine_OnToneReceived " & cTone & " Line: " & index
  If cTone = mTerminateDigit Then
    MessageCounter = 0 ' to suppress repeats of speaking (stat_speaking endless status)
    RedialCounter = 0
    HangUp
    DigitReceived = True

    SignalDone
  End If
  
End Sub
Private Function CallProgressToText(ByVal State As Long) As String
  Dim StateText As String
  'Debug.Print "CallProgressToText " & State & " " & Now
  Select Case State
  Case DivaCallState_Idle
    StateText = "Idle"
  Case DivaCallState_Listening
    StateText = "Listening"
  Case DivaCallState_Connecting
    StateText = "Connecting"
  Case DivaCallState_Ringing
    StateText = "Ringing"
  Case DivaCallState_Offering
    StateText = "Offering"
  Case DivaCallState_Alerting
    StateText = "Alerting"
  Case DivaCallState_Connected
    StateText = "Connected"
  Case DivaCallState_OnHold
    StateText = "OnHold"
  Case DivaCallState_Disconnecting
    StateText = "Disconnecting"
  Case DivaCallState_Disconnected
    StateText = "Disconnected"
  Case DivaCallState_Proceeding
    StateText = "Proceeding"
  Case Else
    StateText = "Unknown State"
  End Select

CallProgressToText = StateText

End Function

Private Sub DIVALine_OnVoiceStreamed(ByVal bWrapped As Long)
  Debug.Print "DIVALine_OnVoiceStreamed  Line: " & index
  Status = STAT_DONESPEAKING


End Sub

Public Property Get Busy() As Boolean

  Busy = mBusy

End Property

Public Property Let Busy(ByVal Value As Boolean)

  mBusy = Value

End Property

Public Property Get Timeout() As Long

  Timeout = mTimeout

End Property

Public Property Let Timeout(ByVal Value As Long)
  
  mTimeout = Value
  

End Property
