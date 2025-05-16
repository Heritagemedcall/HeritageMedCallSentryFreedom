VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpeechEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents spvoice As spvoice
Attribute spvoice.VB_VarHelpID = -1
Public MMSysAudioOut As ISpeechMMSysAudio  ' SAPI
'Public MMSysAudioIn          As ISpeechMMSysAudio ' SAPI not used for now
Public Speaking As Boolean
'Private mVoiceName            As String
Public Voices As Collection

Public Event StartStream()
Public Event EndStream()


Public Function Init() As Long
  dbg "Init"
  Set spvoice = New spvoice
  GetVoices
  'VoiceName = SetVoice(VoiceName)
  spvoice.EventInterests = SVEAllEvents
  ResetAudioOut

End Function
Public Sub ResetAudioOut()
  Dim w As ISpeechWaveFormatEx
  Dim f As SpAudioFormat

  dbg "ResetAudioOut"

  Set spvoice.AudioOutputStream = Nothing
  Set MMSysAudioOut = Nothing
  Set MMSysAudioOut = New SpMMAudioOut  ' SAPI

  Set f = New SpAudioFormat
  'Const SAFT8kHz16BitMono = 6
  'Const SAFT8kHz8BitMono = 4
  'Const SAFTTrueSpeech_8kHz1BitMono = 40 (&H28)
  
  
  
  f.Type = SAFT8kHz16BitMono  '= 6 'SAFTADPCM_8kHzMono or try 4 or 57
''PCM 8,000 Hz, 16-bit, Mono  [wfPCM08000M16]

' Const SAFTADPCM_8kHzMono = 57 (&H39)
  Set w = f.GetWaveFormatEx()
  MMSysAudioOut.Format.SetWaveFormatEx w  ' set the wave format

  spvoice.AllowAudioOutputFormatChangesOnNextSet = False  ' prevent accidental format changes
  Set spvoice.AudioOutputStream = MMSysAudioOut
  Debug.Print "MMSysAudioOut Reset **************"
  dbg "ResetAudioOut End"
End Sub

Public Sub GetVoices()
  Dim token As ISpeechObjectToken

  dbg "GetVoices"
  Set Voices = New Collection
  If Not spvoice Is Nothing Then
    For Each token In spvoice.GetVoices
      Voices.Add token
    Next
  End If
End Sub
Public Function SetVoiceByID(ByVal ID As Long) As String
  On Error Resume Next
  dbg "SetVoiceByID " & ID
  Set spvoice.voice = spvoice.GetVoices().Item(ID)


End Function

Private Function SetVoice(ByVal VoiceName As String) As String
'Dim j       As Integer
  Dim token As ISpeechObjectToken  ' token to hold the registry voice tokens
  dbg "SetVoice " & VoiceName

  If Len(VoiceName) > 0 Then  ' else use default voice
    For Each token In spvoice.GetVoices
      'Debug.Print "Voice "; Token.GetDescription()
      If InStr(1, token.GetDescription(), VoiceName, vbTextCompare) > 0 Then
        Set spvoice.voice = token
        dbgTAPI "Voice changed to: " & token.GetDescription()
        Exit For
      End If
    Next
  End If
  SetVoice = spvoice.voice.GetDescription()
  'Set token =
  'SetVoice = token.GetDescription()
End Function


Public Function Speak(ByVal text As String, ByVal flags As Long) As Long

  dbg "Speak"
  If spvoice Is Nothing Then
    Init  ' configure it
  End If
  If spvoice.AudioOutputStream Is Nothing Then
    Set spvoice.AudioOutputStream = MMSysAudioOut
  End If

  MMSysAudioOut.DeviceID = MediaID  '-1 ' MEDIA ID

  spvoice.Speak text, flags

End Function
Public Function SpeakToFile(ByVal text As String, ByVal flags As Long, ByVal filename As String) As Long

  Dim cpFileStream As New SpFileStream

  If spvoice Is Nothing Then
    Init  ' configure it
  End If

  cpFileStream.Format.Type = SAFTADPCM_8kHzMono
  cpFileStream.Open filename, SSFMCreateForWrite, False  ' make it true?
  spvoice.AllowAudioOutputFormatChangesOnNextSet = False
  Set spvoice.AudioOutputStream = cpFileStream
  spvoice.Speak text, flags
  Do
    DoEvents
    Debug.Print "speak"
  Loop Until spvoice.WaitUntilDone(1)

  If spvoice.AudioOutputStream Is Nothing Then
    Set spvoice.AudioOutputStream = MMSysAudioOut
  End If

  cpFileStream.Close
  Set cpFileStream = Nothing

  If Not cpFileStream Is Nothing Then
    Set cpFileStream = Nothing
  End If

End Function


Private Sub Class_Initialize()
  Init
  Set Voices = New Collection

  MediaID = -1  ' default audio out
End Sub

Private Sub Class_Terminate()
  Finalize
End Sub
Public Sub Finalize()
  Do While Speaking
    DoEvents  ' maybe add a timeout
  Loop
  Set Voices = Nothing
  Set spvoice.AudioOutputStream = Nothing
  Set spvoice = Nothing
  Set MMSysAudioOut = Nothing
End Sub

Private Sub spVoice_AudioLevel(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal AudioLevel As Long)
  dbg "spVoice_AudioLevel " & AudioLevel
End Sub

Private Sub spVoice_Bookmark(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Bookmark As String, ByVal BookmarkId As Long)
'Debug.Print "spVoice_Bookmark "; Bookmark
End Sub

Private Sub spVoice_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
  dbg "spVoice_EndStream " & Now
  RaiseEvent EndStream
  Speaking = False

End Sub

Private Sub spVoice_EnginePrivate(ByVal StreamNumber As Long, ByVal StreamPosition As Long, ByVal EngineData As Variant)
'Debug.Print "spVoice_EnginePrivate"
End Sub

Private Sub spVoice_Phoneme(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Duration As Long, ByVal NextPhoneId As Integer, ByVal Feature As SpeechLib.SpeechVisemeFeature, ByVal CurrentPhoneId As Integer)
'Debug.Print "spVoice_Phoneme"
End Sub

Private Sub spVoice_Sentence(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal CharacterPosition As Long, ByVal Length As Long)
'Debug.Print "spVoice_Sentence"
End Sub

Private Sub spVoice_StartStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
  Debug.Print "spVoice_StartStream " & Now
  Speaking = True
  RaiseEvent StartStream
End Sub

Private Sub spVoice_Viseme(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Duration As Long, ByVal NextVisemeId As SpeechLib.SpeechVisemeType, ByVal Feature As SpeechLib.SpeechVisemeFeature, ByVal CurrentVisemeId As SpeechLib.SpeechVisemeType)
'Debug.Print "spVoice_Viseme"
End Sub

Private Sub spVoice_VoiceChange(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal VoiceObjectToken As SpeechLib.ISpeechObjectToken)
  Dim token As SpeechLib.ISpeechObjectToken
  
  dbg "spVoice_VoiceChange # pos " & StreamNumber & "  " & StreamPosition
  dbg "VoiceObjectToken.GetDescription " & VoiceObjectToken.GetDescription
End Sub

Private Sub spVoice_Word(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal CharacterPosition As Long, ByVal Length As Long)
'Debug.Print "spVoice_Word, CharacterPosition,Length "; CharacterPosition, length
End Sub
'Private Function ShowUI(ByVal theUI As String, ByVal hwnd As Long)
'  Dim token   As ISpeechObjectToken  ' token to hold the registry voice tokens
'  If Not spVoice Is Nothing Then
'    Set token = spVoice.Voice
'    If token.IsUISupported(theUI) = True Then
'      ' Display Master Volume window.
'      token.DisplayUI hwnd, "", theUI, ""
'    End If
'  End If
'End Function

' to send a wave file, use speakstream instead of speak
' to "speak" a spFileStream opened as an AudioOutputStream object

' possibly later for voice analysis/busy tone/ringback
' SpMemoryStream inherits from ISpeechBaseStream interface
'Format Property Gets and sets the cached wave format of the stream as an SpAudioFormat object.
'Methods
'GetData Method Gets the contents of the stream.
'Read Method Reads data from an audio stream.
'Seek Method Returns the current read position of the audio stream in bytes.
'SetData Method Sets the contents of the stream.
'Write Method Writes data to the audio stream

' Often, the Alert Voice is declared without events
' use SPVoice.Priority for verbal alerts
'    Enum SpeechVoicePriority
'        SVPNormal = 0
'        SVPAlert = 1 ' interrups normal voice
'        SVPOver = 2  ' can talk over voice
'    End Enum

' these can also be used to set the AlertBondary where alert voice can interrupt
'    Enum SpeechVoiceEvents
'        SVEStartInputStream = 2
'        SVEEndInputStream = 4
'        SVEVoiceChange = 8
'        SVEBookmark = 16
'        SVEWordBoundary = 32
'        SVEPhoneme = 64
'        SVESentenceBoundary = 128
'        SVEViseme = 256
'        SVEAudioLevel = 512
'        SVEPrivate = 32768
'        SVEAllEvents = 33790
'    End Enum

' SpVoice.GetVoices( [RequiredAttributes As String = ""], [OptionalAttributes As String = ""])) As ISpeechObjectTokens
'Parameters
'RequiredAttributes [Optional] Specifies the RequiredAttributes. All voices selected will match these specifications. If no voices match the selection, the selection returned will contain no voices. By default, no attributes are required and so the list returns all the tokens discovered.
'OptionalAttributes [Optional] Specifies the OptionalAttributes. Voices which match these specifications will be returned at the front of the selection. By default, no attribute is specified and the list returned from the speech configuration database is in the order that attributes were discovered.


Public Property Get Volume() As Long
  If spvoice Is Nothing Then
    Volume = -1
  Else
    Volume = spvoice.Volume
  End If

End Property

Public Property Let Volume(ByVal Volume As Long)

  If Not spvoice Is Nothing Then
    spvoice.Volume = Volume
  End If

End Property

Public Property Get MediaID() As Long
        
        
10      On Error GoTo MediaID_Error

20      dbg "Get MediaID"
30      If MMSysAudioOut Is Nothing Then
40        dbgTAPI "Get MediaID, No MMSysAudioOut " & -2
50        MediaID = -2  ' bogus
60      Else
70        MediaID = MMSysAudioOut.DeviceID
80        dbgTAPI "Get MediaID " & MMSysAudioOut.DeviceID
90      End If

MediaID_Resume:

100     On Error GoTo 0
110     Exit Property

MediaID_Error:

120     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cSpeechEngine.MediaID.Getter." & Erl
130     Resume MediaID_Resume

End Property

Public Property Let MediaID(ByVal Value As Long)
      'dbgtapi  "Let MediaID " & value
10      On Error GoTo MediaID_Error

20      dbg "Let MediaID " & Value

30      If Not MMSysAudioOut Is Nothing Then
          If Value = 0 Then
            LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cSpeechEngine.MediaID.Setter." & "30"
            Exit Property
          Else
40          MMSysAudioOut.DeviceID = Value
          End If
50      Else
60        Debug.Print "MMSysAudioOut letMediaID failed"
70      End If

MediaID_Resume:

80      On Error GoTo 0
90      Exit Property

MediaID_Error:

100     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cSpeechEngine.MediaID.Setter." & Erl & " Media ID=" & Value
110     Resume MediaID_Resume

End Property

Public Property Get VoiceName() As String
  VoiceName = spvoice.voice.GetDescription()
End Property
Public Property Let VoiceName(ByVal Value As String)
  SetVoice Value
End Property

