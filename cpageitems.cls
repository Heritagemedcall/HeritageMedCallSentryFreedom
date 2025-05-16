VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPageItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mPageItems As Collection

Private mLastPageTime As Date

Private Sub Class_Initialize()
  Clear
End Sub

Public Sub Clear()
  Set mPageItems = New Collection
End Sub

Public Property Get Count() As Integer
'  If mPageItems.count > 0 Then MsgBox "1 to send"
  Count = mPageItems.Count
  
End Property

Public Property Get LastPageTime() As Date
  LastPageTime = mLastPageTime
End Property

Public Function AddPageItem(PageItem As cPageItem) As Integer
  mPageItems.Add PageItem
  
End Function

Public Property Get PageItem(ByVal Index As Long) As cPageItem
  If Index > 0 And Index <= mPageItems.Count Then
    Set PageItem = mPageItems(Index)
  End If
End Property

Public Property Get PageItemByAlarmID(ByVal AlarmID As Long) As cPageItem
  Dim PageItem As cPageItem
  For Each PageItem In mPageItems
    If PageItem.AlarmID = AlarmID Then
      Set PageItemByAlarmID = PageItem
      Exit For
    End If
  Next
End Property

Public Function Remove(ByVal Index As Long) As Long
10      On Error GoTo Remove_Error
20      mPageItems.Remove Index
Remove_Resume:
30      On Error GoTo 0
40      Exit Function
Remove_Error:
50      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageItems.Remove." & Erl
60      Resume Remove_Resume
End Function

Private Sub Class_Terminate()
  Dispose
End Sub
Public Sub Dispose()
  'Set mPageItems = Nothing
End Sub
