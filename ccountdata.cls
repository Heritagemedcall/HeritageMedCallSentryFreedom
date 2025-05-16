VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCountData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Model              As String
Public CLSPTI             As Long
Public Description        As String
Public Count              As Long
Public AveragePerDay      As Double

'Public MinPerDay          As Double
'Public MaxPerDay          As Double

Public Sub Increment()
  Count = Count + 1
End Sub
