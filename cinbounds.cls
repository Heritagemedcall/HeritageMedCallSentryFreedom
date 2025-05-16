VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInbounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAlarms As Collection

Private Sub Class_Initialize()
  Set mAlarms = New Collection
  
End Sub
Public Property Get Count() As Long
  Count = mAlarms.Count
End Property


Public Property Get Alarm(ByVal index As Integer) As cAlarm
  If index > 0 And index <= alarms.Count Then
    Set Alarm = mAlarms(index)
  Else
    Set Alarm = New cAlarm
  End If
End Property
Public Sub RemoveAlarm(ByVal index As Integer)
  If index > 0 And index <= alarms.Count Then
    mAlarms.Remove index
  End If
End Sub

