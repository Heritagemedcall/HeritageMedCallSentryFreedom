VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cESDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Base 0
Private mDevices As Collection

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator"
'Collection class enumerator
'Open the Collection class editor window.
'Select the Tools menu's Procedure Attributes command.
'Select NewEnum in the Name box.
'Click Advanced >>
'Click on the Procedure ID box and type -4.
'Check the "Hide this member"

  Set NewEnum = mDevices.[_NewEnum]
End Function
Public Property Get Item(ByVal index As Integer) As cESDevice
  If Not MASTER Then Exit Function
  If index > 0 And index <= mDevices.Count Then
    Set Item = mDevices(index)
  End If
End Property

Public Property Get Devices() As Collection
  Set Devices = mDevices
End Property
Public Property Get Count() As Integer
  Count = mDevices.Count
End Property

'Public Property Get DeviceSlow(ByVal Serial As String) As cESDevice
'  Dim j As Integer
'  If Not MASTER Then Exit Function
'  For j = 1 To mDevices.Count
'    If 0 = StrComp(Serial, mDevices(j).Serial) Then
'      Set DeviceSlow = mDevices(j)
'      Exit For
'    End If
'  Next
'
'End Property


'Public Property Get DeviceFast(ByVal serial As String) As cESDevice
'  Dim j As Long
'
'  On Error Resume Next
'  If Not MASTER Then Exit Property
'  If mDevices Is Nothing Then
'    Exit Property
'  End If
'
'  Dim Device As cESDevice
'
'  Set Device = mDevices.item(serial)
'  If Err.Number Then Exit Property
'  'For j = 1 To mDevices.Count
'  '  If 0 = StrComp(Serial, mDevices(j).Serial) Then
'  '    Set DeviceSlow = mDevices(j)
'  '    Exit For
'  '  End If
'  'Next
'  Set DeviceFast = Device
'End Property





Public Property Get device(ByVal Serial As String) As cESDevice
  Dim Rs            As ADODB.Recordset
  Dim SQL           As String
  Dim NewDevice     As cESDevice


  On Error Resume Next
  If Not MASTER Then
  
    SQL = "Select * from devices where serial = '" & Serial & "'"
    Set Rs = ConnExecute(SQL)
    If Not Rs.EOF Then
    Set NewDevice = New cESDevice
    If NewDevice.Parse(Rs) Then
      Set device = NewDevice
    End If
    Else
      ' Stop
    End If
    Exit Property
  Else
    
    If mDevices Is Nothing Then
      Exit Property
    End If
    Set device = mDevices.Item(Serial)
  End If
  'Exit Sub


  
  


End Property
Public Function AddDevice(device As cESDevice) As Long

' only adds it to the devices collection... does not init the divice



  Dim d As cESDevice

  If Not MASTER Then Exit Function


  On Error Resume Next
  Set d = mDevices(device.Serial)
  If d Is Nothing Then
    mDevices.Add device, (device.Serial) & ""
    AddDevice = mDevices.Count
  Else
    Exit Function
  End If

End Function
Public Sub RefreshByID(ByVal ID As Long)
  Dim SQL As String
  Dim d   As cESDevice
  Dim Rs  As Recordset

10        On Error GoTo RefreshByID_Error
  If Not MASTER Then Exit Sub

20        SQL = "SELECT * FROM Devices WHERE DeviceID = " & ID
30        Set Rs = ConnExecute(SQL)
40        If Not Rs.EOF Then
50          Set d = device(Rs("Serial") & "")
60          If Not d Is Nothing Then
70            d.Refresh
80          End If
90        End If
100       Rs.Close
110       Set Rs = Nothing
120       Set d = Nothing

RefreshByID_Resume:
130       On Error GoTo 0
140       Exit Sub

RefreshByID_Error:

150       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cESDevices.RefreshByID." & Erl
160       Resume RefreshByID_Resume


End Sub
Public Sub RefreshBySerial(ByVal Serial As String)

  Dim d   As cESDevice
  If Not MASTER Then Exit Sub

  On Error Resume Next
  Set d = device(Serial)
  If Not d Is Nothing Then

    d.Refresh  ' this will gather all new data for device
    d.LastSupervise = Now
  End If
  Set d = Nothing

End Sub


Public Sub RemoveDevice(Serial As String)

  Dim j As Integer
  If Not MASTER Then Exit Sub

  For j = 1 To mDevices.Count
    If Serial = mDevices(j).Serial Then
      mDevices.Remove j
      Exit For
    End If
  Next


End Sub

Private Sub Class_Initialize()
  Set mDevices = New Collection
End Sub
