VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFastClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Stopped As Boolean

Implements ICcrpTimerNotify

Public Event TimerEvent()

Private TmrNotify As ccrpTimer

Public Property Let interval(ByVal Value As Long)
  TmrNotify.interval = Value
  TmrNotify.Stats.Frequency = Value
End Property

Public Property Get interval() As Long
  interval = TmrNotify.interval
End Property

Public Sub RunIt()
  Stopped = False
  If Not TmrNotify Is Nothing Then
    Set TmrNotify.Notify = Me
    TmrNotify.EventType = TimerPeriodic
    TmrNotify.Enabled = True
  End If
End Sub

Public Sub StopIt()
  Stopped = True
  If Not TmrNotify Is Nothing Then
    TmrNotify.Enabled = False
  End If
End Sub

Public Sub Dispose()
  Class_Terminate
End Sub

Private Sub Class_Initialize()
  Set TmrNotify = New ccrpTimer
End Sub

Private Sub Class_Terminate()
  If Not TmrNotify Is Nothing Then
    Set TmrNotify.Notify = Nothing
    TmrNotify.Enabled = False
    Set TmrNotify.Notify = Nothing
    Set TmrNotify = Nothing
  End If
End Sub



Public Sub Enable()
  If Not TmrNotify Is Nothing Then
    TmrNotify.Enabled = True
  End If
End Sub

Public Sub Disable()
  If Not TmrNotify Is Nothing Then
    TmrNotify.Enabled = False
  End If
End Sub

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
   
    If Stopped Then Exit Sub
    Static busy As Boolean
    If busy Then Exit Sub
      
      busy = True
      RaiseEvent TimerEvent
    busy = False
End Sub
