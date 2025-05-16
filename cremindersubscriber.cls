VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReminderSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public NameLast           As String
Public NameFirst          As String

Public DeliveryPointID    As Long

Public PagerID            As Long
Public GroupID            As Long



Public PagerName          As String

Private mResidentID As Long
Private mStaffID          As Long
Private mDeliveryPoints   As String
Private mDeliveryPoint    As String
Private mPublicDeliveryPoint As cDeliveryPoint


Public Function GetPublicDeliveryPoint() As cDeliveryPoint
  
  Dim j As Integer
  Dim rows()    As String
  Dim cells()   As String
  Dim Status As Long
  

  Set mPublicDeliveryPoint = New cDeliveryPoint
  On Error Resume Next
  rows = Split(mDeliveryPoints, vbCrLf, 3)
  For j = LBound(rows) To UBound(rows)
    cells = Split(rows(j), vbTab)
    Status = Val(cells(0))
    If Status = DELIVERY_POINT_STATUS.BOTH_STATUS Or Status = DELIVERY_POINT_STATUS.PUBLIC_STATUS Then     ' public
      mPublicDeliveryPoint.Status = Status
      mPublicDeliveryPoint.Address = Trim$(cells(1))
      mPublicDeliveryPoint.AddressType = Trim$(cells(2))
      Exit For
    End If
  Next

  Set GetPublicDeliveryPoint = mPublicDeliveryPoint

End Function


Public Property Get DeliveryPoints() As String
  DeliveryPoints = mDeliveryPoints
End Property

Public Property Let DeliveryPoints(ByVal DeliveryPoints As String)
  mDeliveryPoints = DeliveryPoints
End Property
Public Property Get IsResText() As String
  IsResText = IIf(IsResident, "R", "S")
End Property

Public Property Get NameAll() As String
  NameAll = NameLast & ", " & NameFirst
End Property

Public Property Get PagerKey() As String
  PagerKey = PagerID & "|" & GroupID
End Property

Public Property Let PagerKey(ByVal value As String)
  Dim ids() As String
  PagerID = 0
  GroupID = 0
  
  On Error Resume Next
  
  ids = Split(value, "|")
  PagerID = ids(0)
  GroupID = ids(1)

End Property


Public Property Get ResidentKey() As String
  ResidentKey = ResidentID & "|" & StaffID
End Property

Public Property Let ResidentKey(ByVal value As String)
  Dim ids() As String
  ResidentID = 0
  StaffID = 0
  
  On Error Resume Next
  
  ids = Split(value, "|")
  ResidentID = ids(0)
  StaffID = ids(1)

End Property

Public Property Get IsResident() As Long
  IsResident = ResidentID <> 0
End Property

Public Property Get ResidentID() As Long
  ResidentID = mResidentID
End Property

Public Property Let ResidentID(ByVal ResidentID As Long)
  mResidentID = ResidentID
End Property

Public Property Get StaffID() As Long
  StaffID = mStaffID
End Property

Public Property Let StaffID(ByVal StaffID As Long)
  mStaffID = StaffID
End Property


Public Property Get IsGroup() As Long

  IsGroup = (GroupID <> 0)

End Property

Public Property Get IsPager() As Long
  IsPager = (PagerID <> 0)
End Property
