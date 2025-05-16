VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRebuildQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mQue As Collection
Public RetryCount As Integer


Public Function Add(Item As cRebuildItem) As Integer
'Inserts non-duplicate que item, return index of newly added item (count of que items)
  Dim j As Integer
  For j = mQue.count To 1 Step -1
    If mQue(j).Serial = Item.Serial Then
      Exit For
    End If
  Next
  If j = 0 Then
    If RetryCount > -1 Then
      Item.RetryCount = RetryCount
    End If
    mQue.Add Item
  End If
  Add = mQue.count

End Function

Public Function Remove(ByVal Serial As String) As Integer
  Dim j As Integer
  For j = mQue.count To 1 Step -1
    If mQue(j).Serial = Serial Then
      mQue.Remove j
    End If
  Next
End Function
Public Sub Clear()
  Dim j As Integer
  For j = mQue.count To 1 Step -1
    mQue.Remove j
  Next
  Set mQue = New Collection
End Sub
Private Sub Class_Initialize()
  Set mQue = New Collection
  RetryCount = -1
End Sub
Private Sub Class_Terminate()
  Finalize
End Sub
Public Sub Finalize()
  Set mQue = Nothing
End Sub
Public Function HasNext() As Boolean
  HasNext = mQue.count > 0
End Function
Public Function GetNext() As cWirelessMessage
  Dim WirelessMessage As cWirelessMessage
  Dim RebuildItem As cRebuildItem
  If mQue.count Then
    Set RebuildItem = mQue(1)
    'wirelessmessage.stuff = rebuilditem.stuff
    Set WirelessMessage = New cWirelessMessage
    Set GetNext = WirelessMessage
    mQue.Remove mQue(1)
    Set WirelessMessage = Nothing
    Set RebuildItem = Nothing
  End If
End Function

