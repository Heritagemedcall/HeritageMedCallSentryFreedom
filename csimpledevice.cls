VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimpleDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID           As Long  ' id from database
Private mDecimalSerial As Long
Public Serial       As String  ' hex string w/MID
Public MID          As Long  ' 1-255
Public PTI          As Long  ' 1-255
Public CLS          As Long  ' 1-255 ' not really used
Public Model        As String  ' what we call it
Public AltModel     As String  ' what Inovonics calls it
Public ACGID        As Long  ' what the ACG assigns this

Private mIsPortable As Long  ' is it locatable, pendant
Private mIsRef      As Long  ' is it a fixed point for location
Private mIsRepeater As Long  '
Private mIsSPDevice As Long  ' Soft Pointer Device

Public Checkin6080  As Long  ' seconds 3 margin * 3 minutes * 60 second per minute

Public UpGradeStatus As Long ' 1 good, 2 bad 3 non

Public IDL As Long  ' fixed, mobile, N/A , Soft Pointer
' IDL = 2 = Reference/fixed device
Public IDM As Long ' same as inovonics ID

Public Property Get IsPortable() As Long
  IsPortable = IDL = 1
End Property

Public Property Let IsPortable(ByVal Value As Long)
  If Value Then
    IDL = 1
  End If
  mIsPortable = Value
End Property

Public Property Get IsRef() As Long
  
  IsRef = IDL = 2
End Property

Public Property Let IsRef(ByVal Value As Long)
  If Value Then
    IDL = 2
  End If
  mIsRef = Value
End Property

Public Property Get IsSPDevice() As Long
  IsSPDevice = IDL = 3
End Property

Public Property Let IsSPDevice(ByVal Value As Long)
  If Value Then
    IDL = 3
  End If
  mIsSPDevice = Value
End Property

Public Property Get IsRepeater() As Long
  IsRepeater = mIsRepeater
End Property

Public Property Let IsRepeater(ByVal IsRepeater As Long)
  mIsRepeater = IsRepeater
End Property

Public Property Get DecimalSerial() As Long

  DecimalSerial = Val("&h" & Right$(Serial, 6))

End Property

Public Property Let DecimalSerial(ByVal DecimalSerial As Long)

  mDecimalSerial = DecimalSerial

End Property
