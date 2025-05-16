VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTAPIProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mProviderInfo   As String
Private mLineName       As String
Private mid             As Long
Private mPermanentLineID As Long

Public Property Get ProviderInfo() As String
  ProviderInfo = mProviderInfo
End Property

Public Property Let ProviderInfo(ByVal ProviderInfo As String)
  mProviderInfo = ProviderInfo
End Property

Public Property Get LineName() As String
  LineName = mLineName
End Property

Public Property Let LineName(ByVal LineName As String)
  mLineName = LineName
End Property

Public Property Get id() As Long
  id = mid
End Property

Public Property Let id(ByVal id As Long)
  mid = id
End Property

Public Property Get PermanentLineID() As Long
  PermanentLineID = mPermanentLineID
End Property

Public Property Let PermanentLineID(ByVal PermanentLineID As Long)
  mPermanentLineID = PermanentLineID
End Property
