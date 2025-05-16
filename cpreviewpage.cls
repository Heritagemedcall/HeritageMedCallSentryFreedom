VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPreviewPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mhDC As Long
Private mhBM As Long


Public Property Get HDC() As Long
  HDC = mhDC
End Property

Public Function CreatePage(ByVal Device As Object) As Long
  If mhDC = 0 Then
    mhDC = CreateCompatibleDC(Device.HDC)
  End If
  If mhBM = 0 Then
    mhBM = CreateBitmap(mhDC)
    SelectObject mhDC, mhBM
  End If
  If mhBM = 0 Then
    If mhDC <> 0 Then
      DeleteDC mhDC
      mhDC = 0
    End If
  End If
  CreatePage = mhDC
  


End Function

Private Function CreateBitmap(ByVal HDC As Long) As Long
  If HDC <> 0 Then
    CreateBitmap = CreateCompatibleBitmap(HDC, GetDeviceCaps(HDC, HORZRES), GetDeviceCaps(HDC, VERTRES))
  End If
End Function


 

Private Sub Class_Terminate()
  If mhBM <> 0 Then
    SelectObject mhDC, 0
    DeleteObject mhBM
    mhBM = 0
  End If
  If mhDC <> 0 Then
    DeleteDC mhDC
    mhDC = 0
  End If
End Sub
