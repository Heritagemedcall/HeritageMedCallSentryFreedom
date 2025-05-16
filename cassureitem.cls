VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAssureItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public DeviceID     As Long
Public Serial       As String
Public Model        As String
Public ResidentID   As Long
Public NameLast     As String
Public NameFirst    As String
Public NameFull     As String
Public Phone        As String
Public Room         As String
Public RoomID       As Long
Public CheckInTime  As String


