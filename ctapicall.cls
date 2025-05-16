VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTapiCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1.000  s       for TapiLine.MakeCallAsynch typical
' 0.080  ms      for LineReplyCallRequest
' 4.557  s       for connected
'
' 4.816  s       detect busy ' busy will not time out as long as there is speaking!!!!
'- or -
' 8.993  s       no dialtone ' no voice will be heard, no dialing if no dialtone
'
' 0.151  seconds to Callstate Idle
'-or-
' 0.000  s       manual disconnect
'
' 0.060  s       to Line Idle
' 0.000  s       CallDrop
' 0.000  s       DropCall result

Private mTerminateDigit As Long
Public DigitReceived As Boolean


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

Private mPhoneNumber      As String ' obvious
Public message            As String ' system generated message
Public MessageDelay       As Long   ' seconds to wait before blindly speaking message

Public MessageRepeats     As Long  ' how many times to repeat message
Public MessageRepeatDelay As Long  ' seconds

Public RedialAttempts     As Long  ' how many times to redial if busy, or no dial tone
Public RedialWait         As Long  ' seconds

Public Timeout            As Long   ' seconds: general failure timer

Const ModemWatchDogTimeout As Long = 30

Public TAPIProviders        As Collection ' list of modems available

Private WithEvents TapiLine As CTAPILine
Attribute TapiLine.VB_VarHelpID = -1
Private WithEvents SpeechEngine As cSpeechEngine
Attribute SpeechEngine.VB_VarHelpID = -1

'Private mVoiceName        As String ' voice to use

Private mPermanentLineID  As Long   ' unique hardware ID for voice modem

'Private Lasttime          As Double
Private Elapsed           As Double
'Private starttime         As Double
Private StartDate         As Date

Private MediaID           As Long   ' channel to communicate from speech to modem

Private MessageDelayTimer   As Long  ' timer for connect-dialing-speak message
Private MessageRepeatTimer  As Long  ' timer between message repeats
Private MessageCounter      As Long  ' number of messages repeats
Private RedialCounter       As Long  ' number of redial attempts
Private RedialWaitTimer     As Long  ' timer between redial attempts
'Private DialtoneTimer       As Long  ' maximum time to wait for dialtone
Private TimeoutTimer        As Long  ' general failure timer
Private TimeoutTime         As Date
Private ModemWatchDog       As Long

'Private mActive             As Boolean '

Private Speaking            As Boolean

Public Sub ResetClock()
'  starttime = Timer
  StartDate = Now
End Sub

Public Sub UpdateClock()
  'called once a second from cpagedevice (which is called from the master clock)
  Dim CurrentTime As Double
  CurrentTime = Timer
  
  
  If DateDiff("s", StartDate, Now) >= 1 Then
    StartDate = Now
    CheckStatus
  End If
  
  
End Sub

Public Property Get Voices()
  Set Voices = SpeechEngine.Voices
End Property
Sub CheckStatus()
      ' overall timeout
       ' dbgTAPI ">>>>>>>>>>>>>>> CheckStatus TimoutTimer " & TimeoutTimer
        
10      If TimeoutTimer <= 0 And Status <> STAT_IDLE Then
20        HangUp
30        Status = STAT_IDLE  ' testing

40        dbgTAPI "CTapicCall.CheckStatus Timeout Error"
50        RaiseEvent OnDone
        
60      ElseIf ModemWatchDog <= 0 And Status = STAT_SPEAKING Then
          
        
70        HangUp
80        Status = STAT_IDLE ' testing
          
90        dbgTAPI "CTapicCall.CheckStatus ModemWatchDog Timeout Error"
100       RaiseEvent OnDone
110     Else
120       Select Case Status
            Case STAT_SPEAKING
130           If ModemWatchDog > 0 Then
140             ModemWatchDog = ModemWatchDog - 1
150           End If
          
160         Case STAT_REPEATING, STAT_DONESPEAKING, STAT_REDIALING, STAT_SPEAKDELAY
              ' don't decrement TimeoutTimer
170         Case Else
180           If TimeoutTimer > 0 Then
190             TimeoutTimer = TimeoutTimer - 1
200           End If
210       End Select
220     End If

230     Select Case Status
          Case STAT_IDLE
            ' nothing to do
            'dbgtapi  "STAT_IDLE"
240         If mTerminateDigit Then
250           dbgTAPI " STAT_IDLE, mTerminateDigit, DigitReceived, RedialCounter " & mTerminateDigit & "  " & DigitReceived & "  " & RedialCounter
260           If DigitReceived = False And (RedialCounter > 0) Then
270               RedialCounter = RedialCounter - 1
280               RedialWaitTimer = RedialWait
290               Status = STAT_REDIALING
300               TimeoutTimer = Timeout ' 30
310               dbgTAPI "IDLE - > STAT_REDIALING " & RedialWaitTimer
320              Status = STAT_REDIALING
330              Exit Sub
340           End If
350         End If
               


360       Case STAT_NORMALDISCONNECT
            
370         dbgTAPI "STAT_NORMALDISCONNECT in checkstatus "
            
            ' new gather digits redial if not found
380         If mTerminateDigit <> 0 Then  'STAT_NOANSWER = 200
390             dbgTAPI "DIGIT NOT RECEIVED "
                
400             If RedialCounter > 0 Then
410               RedialCounter = RedialCounter - 1
420               RedialWaitTimer = RedialWait
430               Status = STAT_REDIALING
440               dbgTAPI "STAT_REDIALING " & RedialWaitTimer
450             Else
460               dbgTAPI "REDIALS EXHAUSTED STAT_NORMALDISCONNECT"
470               Status = STAT_IDLE
480               RaiseEvent OnDone
490             End If
               
500         Else
              
510           dbgTAPI "DIGIT NOT NEEDED "
520           Status = STAT_IDLE  ' testing
530           RaiseEvent OnDone
540         End If

550       Case STAT_REDIALING
560         TimeoutTimer = Timeout ' 30
            
570         dbgTAPI "STAT_REDIALING " & RedialWaitTimer
580         If RedialWaitTimer <= 0 Then
590           MakeCallAsynch
600         Else
610           RedialWaitTimer = RedialWaitTimer - 1
620         End If

630       Case STAT_BUSY
            
640         dbgTAPI "STAT_BUSY " & RedialCounter

'            If Not (CallInProgress) Then
'              Status = STAT_IDLE  ' testing
'            Else

650         If RedialCounter > 0 Then
660           RedialCounter = RedialCounter - 1
670           RedialWaitTimer = RedialWait
680           Status = STAT_REDIALING
690         Else

700           TapiLine.DropCallAsynch
              
710           Status = STAT_IDLE  ' testing
720           RaiseEvent OnDone
730         End If
            

740       Case STAT_nodialtone
750         dbgTAPI "STAT_NODIALTONE " & RedialCounter
760         If RedialCounter > 0 Then
770           RedialCounter = RedialCounter - 1
780           RedialWaitTimer = RedialWait
790           Status = STAT_REDIALING
800         Else
810           TapiLine.DropCallAsynch
820           Status = STAT_IDLE  ' testing
830           RaiseEvent OnDone
840         End If

850       Case STAT_CANCELLED
860         dbgTAPI "STAT_CANCELLED"
870         Status = STAT_IDLE  ' testing
880         RaiseEvent OnDone

890       Case STAT_PROCEEDING
            ' OK
900         dbgTAPI "STAT_PROCEEDING"
910       Case STAT_DIALING
            ' OK
920         dbgTAPI "STAT_DIALING"

930       Case STAT_CONNECTED
940         dbgTAPI "STAT_CONNECTED"
950         MessageDelayTimer = MessageDelay
960         MessageCounter = MessageRepeats
970         Status = STAT_SPEAKDELAY

980       Case STAT_SPEAKING
990         dbgTAPI "STAT_SPEAKING"

1000      Case STAT_SPEAKDELAY

1010        dbgTAPI "STAT_SPEAKDELAY Seconds " & MessageDelayTimer & " Remaining"
1020        If MessageDelayTimer <= 0 Then
              ' ???here?? MessageCounter = MessageCounter - 1
              'MessageCounter = MessageCounter - 1
1030          Status = STAT_SPEAKING
1040          PlayMessage Pretext & " " & message

1050        Else
1060          MessageDelayTimer = MessageDelayTimer - 1
1070        End If

1080      Case STAT_DONESPEAKING
1090        dbgTAPI "STAT_DONESPEAKING " & MessageCounter & " Repeats Pending"
1100        MessageRepeatTimer = MessageRepeatDelay

1110        If MessageCounter > 0 Then
1120          MessageCounter = MessageCounter - 1
1130          Status = STAT_REPEATING
1140        Else
              DelayLoop 'TESTING by inserting a delay w/ doevents,
1150          HangUp
1160          Status = STAT_IDLE  ' testing
1170          RaiseEvent OnDone
1180        End If

1190      Case STAT_REPEATING
1200        dbgTAPI "STAT_REPEATING in " & MessageRepeatTimer & " Seconds"
1210        If MessageRepeatTimer <= 0 Then
1220          Status = STAT_SPEAKING
1230          PlayMessage Pretext & " " & message
1240        Else
1250          MessageRepeatTimer = MessageRepeatTimer - 1
1260        End If

1270      Case STAT_STARTED
1280        dbgTAPI "STAT_STARTED"

1290      Case Else  ' STAT_UNK
1300        dbgTAPI "STAT_UNK " & Status

1310    End Select

End Sub



Public Function BeginCall() As Long  ' returns 0 on success of object init and call start

10      On Error GoTo BeginCall_Error

        MemCheck.UpdateMemoryInformation
        
        LogProgramError "STATUS at CTAPILine.MakeCallAsynch.Memory " & MemCheck.MemoryLoad



20      Debug.Print "BeginCall " & Now

30      dbgTAPI "TAPICALL.BeginCall"

40      TimeoutTimer = Timeout       ' 30
50      ModemWatchDog = ModemWatchDogTimeout

60      RedialCounter = RedialAttempts
70      MessageCounter = MessageRepeats
80      Status = STAT_STARTED

90      Call MakeCallAsynch

        Dim MediaID            As Long



100     MediaID = TapiLine.GetLineID("wave/out")
        LogProgramError "STATUS at cTapiCall.BeginCall.100 TapiLine.GetLineID('wave/out')"


        Dim LoopDelayStart     As Date
        Dim LoopDelay          As Long

110     LoopDelay = 3 ' three seconds
120     LoopDelayStart = DateAdd("s", LoopDelay, Now)

        Dim loopcounter As Long

130     If MediaID = 0 Then
140       Do While LoopDelayStart > Now
            DoEvents
            If loopcounter > 1000 Then
              loopcounter = 0
              Sleep 100
150           MediaID = TapiLine.GetLineID("wave/out")
              LogProgramError "STATUS at cTapiCall.BeginCall.140 TapiLine.GetLineID('wave/out') MediaID= " & MediaID
            End If
            
160         If MediaID <> 0 Then
              LogProgramError "STATUS Exit Delay Loop at cTapiCall.BeginCall.Media_ID " & MediaID
170           Exit Do
180         End If
            loopcounter = loopcounter + 1
190       Loop
          If LoopDelayStart > Now Then
            LogProgramError "STATUS Delay Timeout at cTapiCall.BeginCall.Media_ID " & MediaID
          End If
            
200     End If
NoMedia_ID:
210     If MediaID = 0 Then
220       If TapiLine.IsLineOpen Then
230         HangUp
240       End If
250       LogProgramError "Error No MediaID at cTapiCall.BeginCall.NoMedia_ID " & TapiLine.LastError
260       Exit Function
270     End If
       
280     SpeechEngine.MediaID = MediaID
       

BeginCall_Resume:

290     On Error GoTo 0
300     Exit Function

BeginCall_Error:

310     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.BeginCall." & Erl
320     Resume BeginCall_Resume
End Function
Private Function MakeCallAsynch() As Long
        
        Dim delayCounter As Long
        
        
        
        
                
10      DigitReceived = False
        
20      Status = STAT_STARTED
        
        
30      TimeoutTimer = Timeout '30
40      ModemWatchDog = ModemWatchDogTimeout
        'dbgtapi  "TimeoutTimer " & TimeoutTimer
50      If TapiLine.CallInProgress Then
60        dbgTAPI "TapiLine.MakeCallAsynch (CallInProgress) -> TapiLine.DropCallAsync 60"
70        TapiLine.DropCallAsynch
80      End If
        
        '90      Do While TapiLine.hCall <> 0 '' TESTING
        '100       delayCounter = delayCounter + 1
        '110       If delayCounter Mod 500 Then
        '120         DoEvents
        '130       End If
        '140       If delayCounter > 10000 Then
        '145          dbgTAPI "TapiLine.MakeCallAsynch -> delayCounter exceeded 145"
        '             TapiLine.hCall = 0
        '150         Exit Do
        '160       End If
        '170     Loop
                
        
180     If TapiLine.IsLineOpen = False Then
190       dbgTAPI "TapiLine.MakeCallAsynch -> TapiLine.OpenLine (False) 190"
          
200       TapiLine.OpenLine
210     Else
          
220       dbgTAPI "TapiLine.MakeCallAsynch <- TapiLine.IsLineOpen = True"
          HangUp '' TESTING
230       TapiLine.CloseLine
240       dbgTAPI "TapiLine.MakeCallAsynch -> TapiLine.Closeline 240"
250       TapiLine.OpenLine
260       dbgTAPI "TapiLine.MakeCallAsynch -> TapiLine.OpenLine 260"
270     End If
        
280     MakeCallAsynch = TapiLine.MakeCallAsynch(PhoneNumber)
        
        
End Function

Public Function PlayMessage(ByVal text As String) As Long
  If TapiLine.CallInProgress Then
    ModemWatchDog = ModemWatchDogTimeout
    SpeechEngine.Speak text, SVSFlagsAsync
  Else ' don't speak if no call in progress (no hCall)
    dbgTAPI "Can't TAPICAll.PlayMessage without active call"
  End If
End Function

Public Function CancelCall() As Boolean   ' returns 0 on success of call cancelled
10      On Error GoTo CancelCall_Error

20      If TapiLine.CallInProgress Then
30        CancelCall = TapiLine.DropCallAsynch()
40      End If

CancelCall_Resume:

50      On Error GoTo 0
60      Exit Function

CancelCall_Error:

70      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.CancelCall." & Erl
80      Resume CancelCall_Resume
End Function
Public Function HangUp() As Boolean  'Alias for CancelCall returns 0 on success of call cancelled
  'If TapiLine.CallInProgress() Then

10 On Error GoTo HangUp_Error

20 HangUp = TapiLine.DropCallAsynch()
30 DelayLoop
40 Status = STAT_IDLE

  'End If

HangUp_Resume:

50 On Error GoTo 0
60 Exit Function

HangUp_Error:

70 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.HangUp." & Erl
80 Resume HangUp_Resume
End Function

Private Sub Class_Initialize()
  Init
  ResetClock
  UpdateClock
  dbgTAPI "Tapi call INIT done"
End Sub

Private Sub SpeechEngine_EndStream()
  StatusString = "DONE TALKING"
  dbgTAPI StatusString
  Status = STAT_DONESPEAKING
  Speaking = False
End Sub

Private Sub SpeechEngine_StartStream()
  
  StatusString = "START TALKING"
  dbgTAPI StatusString
  Speaking = True
End Sub

Private Sub Class_Terminate()
  Finalize
End Sub

Private Sub TapiLine_Connected()
10      On Error GoTo TapiLine_Connected_Error
        Dim MediaID            As Long

20      StatusString = "CONNECTED"
30      Status = STAT_CONNECTED

        ' need a quick sleep?

40      MediaID = TapiLine.GetLineID("wave/out")


        Dim LoopDelayStart     As Date
        Dim LoopDelay          As Long

50      LoopDelay = 3
60      LoopDelayStart = DateAdd("s", LoopDelay, Now)

70      If MediaID = 0 Then
80        Do While LoopDelayStart > Now
            DoEvents
90          MediaID = TapiLine.GetLineID("wave/out")
100         If MediaID <> 0 Then
110           Exit Do
120         End If
130       Loop
140     End If

NoMedia_ID:

150     If MediaID = 0 Then
160       If TapiLine.IsLineOpen Then
170         HangUp
180       End If
190       LogProgramError "Error No MediaID at cTapiCall.Connected.NoMedia_ID " & TapiLine.LastError
200       Exit Sub
210     End If

220     SpeechEngine.MediaID = MediaID

        'dbgTAPI TapiLine.GetLineID("wave/out")

230     TapiLine.MonitorDigits "dtmf"
240     CheckStatus
250     CheckStatus

TapiLine_Connected_Resume:

260     On Error GoTo 0
270     Exit Sub

TapiLine_Connected_Error:

280     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.TapiLine_Connected." & Erl
290     Resume TapiLine_Connected_Resume

End Sub

Private Sub TapiLine_Dialing()
  StatusString = "DIALING"
  'dbgtapi  StatusString
  Status = STAT_DIALING
End Sub

Private Sub TapiLine_Disconnected(ByVal Reason As Long)
10      On Error GoTo TapiLine_Disconnected_Error

20      StatusString = "DISCONNECTED " & UCase(TapiLine.GetReasonDisconnected(Reason))
        
30      Select Case Reason
          Case 0, 1
          
            '2008-09-22 come back here for revisit new digit received
            'If DigitReceived = False And mTerminateDigit <> 0 Then
            '  Status = STAT_NOANSWER
            'Else
              'Status = STAT_IDLE
40            dbgTAPI "DigitReceived " & DigitReceived
            'End If
50        Case &H20
60          Status = STAT_BUSY
70        Case &H1000 ' 4096
80          Status = STAT_nodialtone
90        Case &H80000 ' 524288
100          Status = STAT_CANCELLED
        
110       Case Else
120          Status = STAT_UNK
130     End Select
140     dbgTAPI StatusString & " " & Reason
        
          
        
150     HangUp '' TESTING both
160     dbgTAPI "TapiLine_DISCONNECTED -> TapiLine.CloseLine "
170     TapiLine.CloseLine
180     Status = STAT_IDLE
        
        
        
        

TapiLine_Disconnected_Resume:

190     On Error GoTo 0
200     Exit Sub

TapiLine_Disconnected_Error:

210     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.TapiLine_Disconnected." & Erl
220     Resume TapiLine_Disconnected_Resume
        
End Sub

Private Sub TapiLine_DropCallResult(ByVal errorCode As Long)
  StatusString = "TAPILINE DROPCALL RESULT " & errorCode
  Status = STAT_IDLE
  'dbgtapi  StatusString
End Sub

Private Sub TapiLine_GotDigit(ByVal Digit As Integer)
10      On Error GoTo TapiLine_GotDigit_Error

20      DigitReceived = False
30      dbgTAPI "TapiLine_GotDigit Digit/TerminateDigit" & Digit & mTerminateDigit

40      If Digit = mTerminateDigit Then
50        MessageCounter = 0 ' to suppress repeats of speaking (stat_speaking endless status)
60        RedialCounter = 0
70        HangUp
80        DigitReceived = True
          'Status = STAT_IDLE
90        dbgTAPI "CTAPICALL GOT DIGIT " & Digit
100       RaiseEvent OnDone
110     End If

TapiLine_GotDigit_Resume:

120     On Error GoTo 0
130     Exit Sub

TapiLine_GotDigit_Error:

140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.TapiLine_GotDigit." & Erl
150     Resume TapiLine_GotDigit_Resume
End Sub

Private Sub TapiLine_Idle()
10      dbgTAPI "Sub TapiLine_Idle , stat, RedialCounter  " & Status & "  " & RedialCounter
        
20      If Status = STAT_nodialtone And RedialCounter > 0 Then
30        Status = STAT_nodialtone
40      ElseIf Status = STAT_BUSY And RedialCounter > 0 Then
50        Status = STAT_BUSY
        
60      Else
70        StatusString = "TAPILINE IDLE"
80        Status = STAT_IDLE
90      End If
        
End Sub

Private Sub TapiLine_MakeCallResult(ByVal errorCode As Long)
  If errorCode <> 0 Then
    StatusString = "TAPILINE MAKECALL ERROR " & TapiLine.GetLineErrString(errorCode)
  Else
    StatusString = "TAPILINE MAKECALL SUCCESS"
  End If
  'dbgtapi  StatusString
End Sub

Sub Init()
10      On Error GoTo Init_Error

20      Timeout = 30
30      ModemWatchDog = ModemWatchDogTimeout
40      DigitReceived = False
50      If TapiLine Is Nothing Then
60        Set TapiLine = New CTAPILine
          'TapiLine.OpenLine
70      End If
80      If SpeechEngine Is Nothing Then
90        Set SpeechEngine = New cSpeechEngine
100       SpeechEngine.Init
110       SpeechEngine.Volume = 100
120       dbgTAPI "SpeechEngine = " & IIf(SpeechEngine Is Nothing, "DEAD", "OK")
130     Else
140       dbgTAPI "SpeechEngine OK"
150     End If
160     EnumerateLines ' for pickin and choosin
        'PermanentLineID = SetTAPILineByPermanentLineID(PermanentLineID)
        'PermanentLineID = SetTAPILineByPermanentLineID(PermanentLineID)
170     SetTAPILineByPermanentLineID (PermanentLineID)
        'SpeechEngine.MediaID = TapiLine.GetLineID("wave/out")

Init_Resume:

180     On Error GoTo 0
190     Exit Sub

Init_Error:

200     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.Init." & Erl
210     Resume Init_Resume
End Sub
Public Function SetTAPILineByPermanentLineID(ByVal PermanentLineID As Long) As Long
        Dim j As Long
        
        
        
10      On Error GoTo SetTAPILineByPermanentLineID_Error

20      If TapiLine.IsLineOpen Then
          'Sleep 50
30        TapiLine.CloseLine
40        Status = STAT_IDLE
          'Sleep 50
          'dbgtapi  "SetTAPILineByPermanentLineID (If TapiLine.IsLineOpen)"
50      End If
          
        'dbgtapi  "TapiLine.numLines  " & TapiLine.numLines
60      For j = TapiLine.numLines - 1 To 0 Step -1
70        TapiLine.CurrentLineID = j
80        If TapiLine.PermanentLineID = PermanentLineID Then
90          Exit For
100       End If
110     Next
120     SetTAPILineByPermanentLineID = TapiLine.PermanentLineID
130     TapiLine.OpenLine
        'SpeechEngine.MediaID = TapiLine.GetLineID("wave/out")
        'dbgtapi  "SpeechEngine.MediaID " & SpeechEngine.MediaID

SetTAPILineByPermanentLineID_Resume:

140     On Error GoTo 0
150     Exit Function

SetTAPILineByPermanentLineID_Error:

160     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.SetTAPILineByPermanentLineID." & Erl
170     Resume SetTAPILineByPermanentLineID_Resume
End Function
Public Sub Finalize()
  If TapiLine.CallInProgress Then
    TapiLine.DropCallAsynch
  End If
  Set SpeechEngine = Nothing
  Set TapiLine = Nothing
End Sub
Private Sub TapiLine_Proceeding()
  StatusString = "TAPILINE PROCEEDING"
  'dbgtapi  StatusString
  Status = STAT_PROCEEDING
End Sub
Sub EnumerateLines()
        Dim TapiLine           As CTAPILine
        Dim j                  As Long
        'Dim s         As String
        Dim provider           As cTAPIProvider

10      On Error GoTo EnumerateLines_Error

20      Set TAPIProviders = New Collection

30      Set TapiLine = New CTAPILine
40      If TapiLine.Create <> 0 Then
50        For j = 0 To TapiLine.numLines - 1
60          TapiLine.CurrentLineID = j
70          If TapiLine.NegotiatedAPIVersion > 0 Then
80            If TapiLine.LineSupportsVoiceCalls Then
90              If 1 = 1 Then        'InStr(1, TapiLine.ProviderInfo, "Modem", vbTextCompare) > 0 Then
100               Set provider = New cTAPIProvider
110               provider.ProviderInfo = TapiLine.ProviderInfo
120               provider.LineName = TapiLine.LineName
130               provider.PermanentLineID = TapiLine.PermanentLineID
140               provider.ID = j
150               TAPIProviders.Add provider
160             End If
170           End If
180         End If
190       Next
200     End If
210     TapiLine.Finalize "Enumerate Lines"


EnumerateLines_Resume:

220     Set provider = Nothing
230     Set TapiLine = Nothing

240     On Error GoTo 0
250     Exit Sub

EnumerateLines_Error:

260     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.EnumerateLines." & Erl
270     Resume EnumerateLines_Resume
End Sub
Public Property Get DeviceName() As String
  On Error Resume Next
  DeviceName = TapiLine.LineName
End Property
Public Property Get VoiceName() As String
  On Error Resume Next
  VoiceName = SpeechEngine.VoiceName
End Property

Public Property Let VoiceName(ByVal Value As String)
        
10      On Error GoTo VoiceName_Error

20       If SpeechEngine.VoiceName <> Value Then
30        SpeechEngine.VoiceName = Value
40      End If
50      Debug.Print "SpeechEngine.VoiceName " & SpeechEngine.VoiceName

VoiceName_Resume:

60      On Error GoTo 0
70      Exit Property

VoiceName_Error:

80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.VoiceName." & Erl
90      Resume VoiceName_Resume
End Property

Public Property Get PermanentLineID() As Long
  PermanentLineID = mPermanentLineID
End Property

Public Property Let PermanentLineID(ByVal Value As Long)
  mPermanentLineID = SetTAPILineByPermanentLineID(Value)
  
End Property

Public Property Get Active() As Boolean
        
10      On Error GoTo Active_Error

20      Active = TapiLine.IsLineOpen

Active_Resume:

30      On Error GoTo 0
40      Exit Property

Active_Error:

50      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTapiCall.Active." & Erl
60      Resume Active_Resume
End Property

'Public Property Let Active(ByVal Value As Boolean)
'  TapiLine.OpenLine
'  SetTAPILineByPermanentLineID (PermanentLineID)
'End Property

Public Property Get TerminateDigit() As Long

  TerminateDigit = mTerminateDigit

End Property

Public Property Let TerminateDigit(ByVal Value As Long)
  'a setting of 0 turns off Monitor Digits
  ' it will be passed to the call upon "connected"
  mTerminateDigit = Value

End Property

Public Property Get PhoneNumber() As String

  PhoneNumber = mPhoneNumber

End Property

Public Property Let PhoneNumber(ByVal PhoneNumber As String)

  mPhoneNumber = PhoneNumber

End Property
