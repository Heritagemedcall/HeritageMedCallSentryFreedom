Attribute VB_Name = "modTapiLine"

Option Explicit



Private Declare Sub CopyMemory2 Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal Length As Long)
Public Sub LineCallbackProc(ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
'the callbackInstance parameter contains a pointer to the cTAPILine class
'this sub just routes all callbacks back to the class for handling there

  Dim PassedObj   As CTAPILine
  Dim objTemp     As CTAPILine

  'dbgtapi  "LineCALLBACK : dwCallbackInst = " & dwCallbackInstance
  If dwCallbackInstance <> 0 Then
    'turn pointer into lightweight uncounted reference
    'passes byref objtemp and byref dwCallbackInstance
    CopyMemory2 objTemp, dwCallbackInstance, 4
    'Assign to legal reference
    Set PassedObj = objTemp
    'Destroy the illegal reference
    CopyMemory2 objTemp, 0&, 4
    'use the interface to call back to the class
    PassedObj.LineProcHandler hDevice, dwMsg, dwParam1, dwParam2, dwParam3
  End If
End Sub

'Lower 16 bits of a 32 bit value
Function LoWord(ByVal value As Long) As Long
  If value And &H8000& Then
    LoWord = value Or &HFFFF0000
  Else
    LoWord = value And &HFFFF
  End If
End Function
'Only works for positive numbers
Function LShiftWord(ByVal value As Long, ByVal Bits As Integer) As Long
  LShiftWord = value * (2 ^ Bits)
End Function




