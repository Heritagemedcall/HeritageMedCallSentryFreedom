VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BeepTimer        As Long ' minutes
Public DefaultBeepTime  As Long

Public Items As Collection
Private mPending As Boolean
Public Property Get Pending() As Boolean
  Pending = mPending
  mPending = False
End Property
Public Function Add(dev As cESDevice) As Boolean
        Dim j As Integer
10       On Error GoTo Add_Error

20      For j = 1 To Items.count
30        If Items(j).Serial = dev.Serial Then
40          Exit For
50        End If
60      Next
70      If j > Items.count Then
80        Items.Add dev
90        Add = True
100       BeepTimer = DefaultBeepTime
110       mPending = True
120     End If

Add_Resume:
130      On Error GoTo 0
140      Exit Function

Add_Error:

150     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cEventItems.Add." & Erl
160     Resume Add_Resume

End Function
Public Function Remove(dev As cESDevice) As Boolean
        Dim j As Integer
10       On Error GoTo Remove_Error

20      For j = 1 To Items.count
30        If Items(j).Serial = dev.Serial Then
40          Items.Remove j
50          Remove = True
60          mPending = True
70          Exit For
80        End If
90      Next
100     If Items.count = 0 Then
110       BeepTimer = 0
120     End If

Remove_Resume:
130      On Error GoTo 0
140      Exit Function

Remove_Error:

150     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cEventItems.Remove." & Erl
160     Resume Remove_Resume

End Function


Private Sub Class_Initialize()
  Set Items = New Collection
  DefaultBeepTime = 5
End Sub



