VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPagerGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPagers As Collection

Public PagerGroupID As Long
Public PagerGroupName As Long
Public TimesSent    As Long


Private mEscalationLevel As Integer
Private mExpires       As Date



Public Property Get PagerCount() As Long
  PagerCount = mPagers.Count
End Property


Private Sub Class_Initialize()
  Set Pagers = New Collection
  mExpires = DateAdd("n", 5, Now)
End Sub

Private Sub Class_Terminate()
  Set mPagers = Nothing
End Sub

Public Property Get Pagers() As Collection
  Set Pagers = mPagers
End Property

Public Property Set Pagers(Pagers As Collection)
  Set mPagers = Pagers
End Property

Public Property Get EscalationLevel() As Integer
  EscalationLevel = mEscalationLevel
End Property

Public Property Let EscalationLevel(ByVal EscalationLevel As Integer)
  mEscalationLevel = EscalationLevel
End Property

Public Property Get Expires() As Date
  Expires = mExpires
End Property

Public Property Let Expires(ByVal Expires As Date)
    
  mExpires = Expires
End Property
Public Property Get Expired() As Boolean
  Expired = Now > mExpires
End Property
