VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DeviceID     As Long
Public Serial       As String
Public ResidentID   As Long
Public RoomID       As Long
Public Console      As Long


Public UniqueID     As Long
Public IDM          As Long '0 to 255
Public IDL          As Long '0 to 255
Public FC1          As Long '0 to 255
Public FC2          As Long '0 to 255
Public Status       As Long

Public IsCheckIn    As Integer ' 0 or 1
Public IsLocator    As Integer ' 0 or 1


Public LOCIDM       As Long '0 TO 255
Public LOCIDL       As Long '0 TO 255

Public Alarm        As Integer  '0 or 1
Public Battery      As Integer  '0 or 1
Public Tamper       As Integer  '0 or 1


Public SubItem0     As String
Public SubItem1     As String
Public SubItem2     As String
Public SubItem3     As String
Public SubItem4     As String
Public SubItem5     As String
Public SubItem6     As String


Public EventDate    As String
Public TimeEvent    As String
Public TimeSilent   As String
Public TimeAck      As String

Public TimeLast     As Date

Public Room         As String
Public NameFirst    As String
Public NameLast     As String
Public Location     As String
Public Building     As String


' No Event
' Alarm/emergency
' Alert
' Restore
' Assurance
' Silence
' Acknowledge
' CheckinFail
' CheckinRestore
' BatteryFail
' BatteryRestore


Public Eventtype    As Long

Public Property Get NameFirstLast() As String
  Dim s As String
  If Len(NameFirst) > 0 Then
    s = NameFirst
  End If
  If Len(s) > 0 Then
    If Len(NameLast) > 0 Then
      s = s & " " & NameLast
    End If
  Else
    s = NameLast
  End If
End Property
Public Property Get NameLastFirst() As String
  Dim s As String
  If Len(NameLast) > 0 Then
    s = NameLast
  End If
  If Len(s) > 0 Then
    If Len(NameFirst) > 0 Then
      s = s & ", " & NameFirst
    End If
  Else
    s = NameFirst
  End If
End Property

Public Property Get SubItems(ByVal index As Integer) As String
  'Map to listview columns
  Select Case index
    Case 0
      SubItems = Serial
    Case 1
      SubItems = NameLastFirst
    Case 2
      SubItems = Room
    Case 3
      SubItems = Location
    Case 4
      SubItems = Format(TimeEvent, "mm/dd hh:nn")
    Case 5
      SubItems = Format(TimeSilent, "mm/dd hh:nn")
  End Select

End Property
