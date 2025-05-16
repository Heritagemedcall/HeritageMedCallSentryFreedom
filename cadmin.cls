VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const AdminTimeout = 10 * 60

Public LastSeen  As Date
Public Session   As Long
Public Available As Boolean

Public Property Get IsAvailable()
  If IsDate(LastSeen) Then
    If DateDiff("s", LastSeen, Now) > AdminTimeout Then
      IsAvailable = True
    Else
      IsAvailable = False
    End If
  Else
    IsAvailable = True
  End If
End Property


