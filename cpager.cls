VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' new 7/10/1018
Public OuputGroup        As Long
Public OuputGroupName    As String

Public InputNum          As Long


' new 3/9/09
Public AlarmID           As Long
Public Address           As String
Private mPIN             As String

Public PagerID           As Long
Public Phone             As String
Public RoomText          As String
Private mMarquisMessage  As String

Public Twice             As Integer
Public KeyPA             As Integer

Public SendCount         As Long

Public RelayNum          As Integer

Public Description       As String
Public Identifier        As String
Public DefaultMessage    As String

Public DeviceID          As Long
Public NoName            As Long
Public NoRepeats         As Integer
Public RepeatTwice       As Integer
Public Repeats           As Integer
Public SendCancel        As Integer
Public RepeatTilReset    As Integer
Public RepeatSeconds     As Integer

Public ProtocolID        As Long
Public NoCancel          As Integer  ' Central office gets no cancel?

Public Sendphone         As Integer
Public MarquisChar       As String

Public RelayNumber       As Integer

Private mMarquisCode     As Integer

Private mMessage         As String
'Public Property Get Done() As Boolean
'  ' we are done if this device does not do repeats
'  If NoRepeats Then
'    Done = True
'  End If
'End Property


'Public Function MustCancel() As Boolean
'  MustCancel = mMarquisCode > 0
'End Function

Public Function IsApollo() As Boolean
    IsApollo = mMarquisCode = (MARQUIS_APOLLO)
End Function

Public Function IsMarquis() As Boolean
  IsMarquis = (mMarquisCode > 0) And (mMarquisCode <> MARQUIS_APOLLO)
End Function


Public Property Get MarquisCode() As Integer
  MarquisCode = mMarquisCode
End Property

Public Property Let MarquisCode(ByVal code As Integer)

  mMarquisCode = code
  If mMarquisCode > 0 Then
    RepeatTwice = 0
    NoRepeats = 1
    Sendphone = 0
    If mMarquisCode <> MARQUIS_APOLLO Then
     MarquisChar = MarquiCode2MarquiChar(mMarquisCode)
    End If
  End If
End Property


Public Property Get message() As String

  message = mMessage

End Property

Public Property Let message(ByVal NewMessage As String)

  mMessage = NewMessage

End Property

Public Property Get MarquisMessage() As String

  MarquisMessage = mMarquisMessage

End Property

Public Property Let MarquisMessage(ByVal MarquisMessage As String)

  mMarquisMessage = MarquisMessage

End Property

Public Property Get PIN() As String

  PIN = mPIN

End Property

Public Property Let PIN(ByVal PIN As String)

  mPIN = PIN

End Property
