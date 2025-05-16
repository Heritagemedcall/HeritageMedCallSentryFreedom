VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' just a wrapper to use in collections

Public StringValue    As String
Public IntValue       As Integer
Public LongValue      As Long
Public DoubleValue    As Double
Public ByteValue      As Byte
Public SingleValue    As Single
Public ObjectValue    As Object
Public VariantValue   As Variant
