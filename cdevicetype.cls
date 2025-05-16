VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Custom ESDEVICETYPE

Public id           As Long

Public Model        As String '
Public Description  As String
Public IsLatching   As Integer ' 0 or 1
Public IsPortable   As Integer ' 0 or 1
Public Checkin      As Integer
Public Announce     As String
Public Announce2    As String
Public AllowDisable As Integer ' 0 or 1
Public MIDPTI       As Long
Public CLSPTI       As Long
Public ClearByReset As Integer ' 0 or 1
Public NumInputs    As Integer ' 0,1,2,3
Public NoTamper     As Integer ' 0 or 1

Public IgnoreTamper As Integer

'New With build

Private mAutoClear  As Long     ' 1 to 10

'new with build 226
' input 1
Public Repeats      As Long     ' 1 to 10
Public Pause        As Long     ' 1 to 999
Public RepeatUntil  As Integer  ' 0 or 1
Public SendCancel   As Integer  ' 0 or 1

' input 2
Public Repeats_A      As Long     ' 1 to 10
Public Pause_A        As Long     ' 1 to 999
Public RepeatUntil_A  As Integer  ' 0 or 1
Public SendCancel_A   As Integer  ' 0 or 1

' input 3
Public Repeats_B      As Long     ' 1 to 10
Public Pause_B        As Long     ' 1 to 999
Public RepeatUntil_B  As Integer  ' 0 or 1
Public SendCancel_B   As Integer  ' 0 or 1

' day group first button

Public OG1            As Long
Public OG2            As Long
Public OG3            As Long
Public OG4            As Long
Public OG5            As Long
Public OG6            As Long

' night group first button

Public NG1            As Long
Public NG2            As Long
Public NG3            As Long
Public NG4            As Long
Public NG5            As Long
Public NG6            As Long

' graveyard group first button

Public GG1            As Long
Public GG2            As Long
Public GG3            As Long
Public GG4            As Long
Public GG5            As Long
Public GG6            As Long

' Second button

Public OG1_A          As Long
Public OG2_A          As Long
Public OG3_A          As Long
Public OG4_A          As Long
Public OG5_A          As Long
Public OG6_A          As Long


Public NG1_A          As Long
Public NG2_A          As Long
Public NG3_A          As Long
Public NG4_A          As Long
Public NG5_A          As Long
Public NG6_A          As Long

Public GG1_A          As Long
Public GG2_A          As Long
Public GG3_A          As Long
Public GG4_A          As Long
Public GG5_A          As Long
Public GG6_A          As Long


' this is the delay for above groups

Public OG1d           As Long
Public OG2d           As Long
Public OG3d           As Long
Public OG4d           As Long
Public OG5d           As Long
Public OG6d           As Long


Public NG1d           As Long
Public NG2d           As Long
Public NG3d           As Long
Public NG4d           As Long
Public NG5d           As Long
Public NG6d           As Long

Public GG1d           As Long
Public GG2d           As Long
Public GG3d           As Long
Public GG4d           As Long
Public GG5d           As Long
Public GG6d           As Long


Public OG1_Ad         As Long
Public OG2_Ad         As Long
Public OG3_Ad         As Long
Public OG4_Ad         As Long
Public OG5_Ad         As Long
Public OG6_Ad         As Long


Public NG1_Ad         As Long
Public NG2_Ad         As Long
Public NG3_Ad         As Long
Public NG4_Ad         As Long
Public NG5_Ad         As Long
Public NG6_Ad         As Long

Public GG1_Ad         As Long
Public GG2_Ad         As Long
Public GG3_Ad         As Long
Public GG4_Ad         As Long
Public GG5_Ad         As Long
Public GG6_Ad         As Long


'need third shift

Public Property Get AutoClear() As Long

  AutoClear = mAutoClear

End Property

Public Property Let AutoClear(ByVal AutoClear As Long)

  mAutoClear = AutoClear

End Property
