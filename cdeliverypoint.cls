VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeliveryPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DELIVERY_POINT
  Phone = 0
  phone_ack = 1
  EMAIL = 2
End Enum

Public Enum DELIVERY_POINT_STATUS
  NEITHER_STATUS = 0
  PRIVATE_STATUS = 1
  PUBLIC_STATUS = 2
  BOTH_STATUS = 3
End Enum

Public Status       As Long  ' 0 = no status, 1 = private default, 2 = public Default
Public Address      As String
Public AddressType  As Long   ' phone email etc

