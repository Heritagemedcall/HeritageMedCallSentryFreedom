VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWayPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID           As String
Public InternalID   As String
Public Description  As String
Public Building     As String
Public Floor        As String
Public Wing         As String

Public Repeater1    As String
Public Repeater2    As String
Public Repeater3    As String

Public Signal1      As Double
Public Signal2      As Double
Public Signal3      As Double

Public Repeater1Desc   As String
Public Repeater2Desc   As String
Public Repeater3Desc   As String


Private mFirstHops  As Collection


Public DateTime     As Date
Public Serial       As String
Public Checked      As Boolean


Sub SortLocations(a())
  Dim offset  As Integer
  Dim maxrow  As Integer
  Dim limit   As Integer
  Dim row     As Integer
  Dim switch  As Integer
  Dim MinRow  As Integer
  Dim Temp
  

  maxrow = UBound(a)
  MinRow = 1


  offset = maxrow \ 2
  Do While offset > 0
    limit = maxrow - offset
    Do
      switch = 0
      For row = MinRow To limit
        'If a(row) > a(row + offset) Then
        If a(row).LEvel < a(row + offset).LEvel Then  ' may need to incorporate margin
          Set Temp = a(row)
          Set a(row) = a(row + offset)
          Set a(row + offset) = Temp
          Set Temp = Nothing
          switch = row
        End If
      Next row
      limit = switch - offset
    Loop While switch

    offset = offset \ 2
  Loop


End Sub


Private Sub Class_Initialize()
  Set mFirstHops = New Collection
  
End Sub

Public Property Get average() As Double

  average = CDbl(Signal1 + Signal2 + Signal3) / 3

End Property

