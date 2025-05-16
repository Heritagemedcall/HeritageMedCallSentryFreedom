VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuedEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EventName As String

'Public Param1    As String
'Public Param2    As String
'Public Param3    As String
'Public Param4    As String

Private mConsoleID As String
Private mAlarmtype As String


Public Function ToXML() As String
  Dim XML As String
  XML = taggit("EventName", EventName)
  XML = XML & taggit("ConsoleID", ConsoleID)
  XML = XML & taggit("AlarmType", Alarmtype)
  ToXML = XML
End Function

Public Property Get ConsoleID() As String

  ConsoleID = mConsoleID

End Property

Public Property Let ConsoleID(ByVal ConsoleID As String)

  mConsoleID = ConsoleID

End Property

Public Property Get Alarmtype() As String

  Alarmtype = mAlarmtype

End Property

Public Property Let Alarmtype(ByVal Alarmtype As String)

  mAlarmtype = Alarmtype

End Property
