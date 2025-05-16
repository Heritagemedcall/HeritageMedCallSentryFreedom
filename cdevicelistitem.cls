VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Public Serial      As String
Public DeviceID    As Long
Public ResidentID  As Long
Public RoomID      As Long
Public Model       As String
Public UseAssur    As Integer
Public UseAssur2   As Integer
Public AssurInput  As Integer
Public Ignored     As Integer

Public NameLast    As String
Public NameFirst   As String
Public Room        As String
Public Custom      As String ' custom device description

Public Property Get AssurString() As String
  Dim s As String
  s = IIf(UseAssur = 1, "Y", "N") & IIf(UseAssur2 = 1, "Y", "N")
  If UseAssur = 1 Or UseAssur2 = 1 Then
      s = s & AssurInput
  End If
  AssurString = s
End Property
Public Property Get NameFull() As String
  If Len(NameLast) > 0 Or Len(NameFirst) > 0 Then
    NameFull = NameLast & ", " & NameFirst
  End If
End Property
Public Property Get ListKey() As String
  ListKey = DeviceID & "B"
End Property
Public Property Get ModelAndStatus() As String
  ModelAndStatus = Model & IIf(Ignored And 1, " *X*", "")
End Property
