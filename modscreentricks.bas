Attribute VB_Name = "modScreenTricks"
Option Explicit


Private Const LWA_COLORKEY = 1
Private Const LWA_ALPHA = 2
Private Const LWA_BOTH = 3
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = -20
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Sub SetTranslucent(ByVal hwnd As Long, ByVal Trans As Integer, Optional Red As Integer = 255, Optional Green As Integer = 255, Optional Blue As Integer = 255)
  On Error GoTo ErrorExit
  Dim attrib As Long
  If Red > 255 Then Red = 255
  If Red < 0 Then Red = 0

  If Green > 255 Then Green = 255
  If Green < 0 Then Green = 0

  If Blue > 255 Then Blue = 255
  If Blue < 0 Then Blue = 0
  
  If Trans > 100 Then Trans = 100
  If Trans < 0 Then Trans = 0


  attrib = GetWindowLong(hwnd, GWL_EXSTYLE)
  SetWindowLong hwnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
  SetLayeredWindowAttributes hwnd, RGB(Red, Green, Blue), Trans, LWA_ALPHA
  Exit Sub
ErrorExit:
End Sub

