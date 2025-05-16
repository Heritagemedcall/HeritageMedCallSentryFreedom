VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cREgistrationHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DeviceID     As Long
Public PTI          As Long
Public MID          As Long
Public locatable    As Boolean
Public Description  As String
Public IsRef        As Boolean
Public SyncWindow   As Long
Public SyncTimeOut  As Long
Public MessageExpirationTime As Long
Public CheckInTime  As Long
Public SupervisionWindow As Long
Public IsSPDevice   As Boolean



'
'End Function
Private Sub Class_Initialize()

End Sub
