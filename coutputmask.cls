VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutputMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Announce     As String

Public Repeats      As Long
Public RepeatUntil  As Long
Public SendCancel   As Long

Public Pause        As Long

Public RepeatTwice  As Boolean

Public ScreenName   As String

Public OG1          As Long
Public OG2          As Long
Public OG3          As Long
Public OG4          As Long
Public OG5          As Long
Public OG6          As Long

Public OG1D         As Long
Public OG2D         As Long
Public OG3D         As Long
Public OG4D         As Long
Public OG5D         As Long
Public OG6D         As Long


Public NG1          As Long
Public NG2          As Long
Public NG3          As Long
Public NG4          As Long
Public NG5          As Long
Public NG6          As Long


Public NG1D         As Long
Public NG2D         As Long
Public NG3D         As Long
Public NG4D         As Long
Public NG5D         As Long
Public NG6D         As Long

Public GG1          As Long
Public GG2          As Long
Public GG3          As Long
Public GG4          As Long
Public GG5          As Long
Public GG6          As Long

Public GG1D         As Long
Public GG2D         As Long
Public GG3D         As Long
Public GG4D         As Long
Public GG5D         As Long
Public GG6D         As Long



'Public Function ToString() As String
'
'End Function
'Public Function Verbose() As String
'
'Verbose = " Repeats " & Repeats & _
 '          " Pause  " & Repeats & _
 '          " OG1, OG2 " & OG1 & " " & OG2
'
'
'End Function

