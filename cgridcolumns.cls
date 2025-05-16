VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGridColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Private Type ColumnType
  Size      As Single
  Caption   As String
End Type

Private Columns() As ColumnType




Private Sub Class_Initialize()
  ReDim Columns(1)
End Sub

Public Property Get Size(ByVal index As Integer) As Single
  If index > 0 And index <= UBound(Columns) Then
    Size = Columns(index).Size
  End If
End Property

Public Property Let Size(ByVal index As Integer, ByVal Value As Single)
  If index > UBound(Columns) Then
    ReDim Preserve Columns(index)
  End If
  Columns(index).Size = Value
End Property
