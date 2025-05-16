VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ResidentID     As Long
Public NameLast       As String
Public NameFirst      As String
Public Room           As String
Public RoomID         As Long
Public Group          As String
Public GroupID        As Long
Public info           As String
Public Phone          As String


Public Assurdays      As Long
Public Vacation       As Long

Public AssignedTx    As Collection

Public Picture       As IPictureDisp

Public DeliveryPoints         As Collection

Public DeliveryPointsString   As String

Public Function ParseDeliveryPoints() As Long

  Dim dp As cDeliveryPoint
  Dim j As Integer
  
  
  Dim rows()    As String
  Dim cells()   As String
  
  Set DeliveryPoints = New Collection
  rows = Split(DeliveryPointsString, vbCrLf, 3)
  For j = LBound(rows) To UBound(rows)
    Set dp = New cDeliveryPoint
    cells = Split(rows(j), vbTab)
    On Error Resume Next
    dp.Status = Val(cells(0))
    dp.Address = Trim$(cells(1))
    dp.AddressType = Trim$(cells(2))
    DeliveryPoints.Add dp
    
  Next
  ParseDeliveryPoints = DeliveryPoints.Count

End Function

Public Function DeliveryPointsToString() As String
  
  Dim dp As cDeliveryPoint
  
  Dim s       As String
  Dim row     As String
  
  For Each dp In DeliveryPoints
    row = Join(Array(dp.Status, dp.Address, dp.AddressType), vbTab) ' field sep
    If Len(s) Then
      s = s & vbCrLf ' record sep
    End If
    s = s & row
  Next
  DeliveryPointsToString = s
End Function


Public Function Fetch(ByVal ID As Long) As Long
  Dim Rs As Recordset
  Set Rs = ConnExecute("Select * FROM Residents WHERE ResidentID = " & ID)
  If Not Rs.EOF Then
    Fetch = Rs("ResidentID")
    Parse Rs
  End If
  Rs.Close
  Set Rs = Nothing
  
End Function
Function Parse(Rs As Recordset) As Long
  ResidentID = Rs("ResidentID")
  NameLast = Rs("NameLast") & ""
  NameFirst = Rs("Namefirst") & ""
  Room = Rs("room") & ""
  RoomID = Val(Rs("RoomID") & "")
  info = Rs("info") & ""
  Phone = Rs("phone") & ""
  Vacation = IIf(Rs("Away") = True, 1, 0)
  Assurdays = Val("" & Rs("AssurDays"))
  DeliveryPointsString = "" & Rs("Deliverypoints")
End Function





Public Property Let AssurDay(ByVal index As Long, ByVal value As Long)
'index (bit) 1 thur 7
'index 1 is monday
'index 2 is tues
' Value is either 1 or 0 (on or off)
  
  value = IIf(value = 0, 0, 1)
  If index > 0 And index < 8 Then  ' only 1 thru 7 ' bit 0 is reserved
    If value = 1 Then  ' set the bit
      Assurdays = Assurdays Or (2 ^ index)
    Else
      Assurdays = Assurdays And (Not (2 ^ index))
    End If
  End If

End Property


Public Property Get AssurDay(ByVal index As Long) As Long
'index (bit) 1 thur 7
'index 1 is monday
'index 2 is tues etc

' ReturnValue is either 1 or 0 (on or off)
  
  If index > 0 And index < 8 Then  ' only 1 thru 7 ' bit 0 is reserved
      AssurDay = IIf((Assurdays And (2 ^ index)) = 0, 0, 1)
  End If

End Property



Public Function AddTx(device As cESDevice) As Long
  Dim j As Integer
  For j = 1 To AssignedTx.Count
    If AssignedTx(j).DeviceID = device.DeviceID Then
      Exit For
    End If
  Next
  If j > AssignedTx.Count Then
    AssignedTx.Add device
  End If


End Function

Sub GetTransmitters()

        Dim SQL As String
        Dim Rs  As Recordset
        Dim tx  As cESDevice
        Dim i   As Long
        

10       On Error GoTo GetTransmitters_Error

20      Set AssignedTx = New Collection
30      If ResidentID <> 0 Then
40        SQL = "SELECT Devices.* FROM Devices WHERE ResidentID <> 0 AND ResidentID = " & ResidentID
50        Set Rs = ConnExecute(SQL)
60        Do Until Rs.EOF
            i = i + 1
            Debug.Print "cResident.GetTransmitters count "; i
70          Set tx = New cESDevice
80          tx.Parse Rs
            tx.FetchResident
            tx.FetchRoom
90          AssignedTx.Add tx
      
100         Rs.MoveNext
110       Loop
120       Rs.Close
130       Set Rs = Nothing
140     End If

GetTransmitters_Resume:
150      On Error GoTo 0
160      Exit Sub

GetTransmitters_Error:

170     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cResident.GetTransmitters." & Erl
180     Resume GetTransmitters_Resume

End Sub

Private Sub Class_Initialize()
  Set AssignedTx = New Collection
  Set DeliveryPoints = New Collection
  
End Sub

Public Sub Dispose()

  Set AssignedTx = Nothing
  Set DeliveryPoints = Nothing

End Sub
Private Sub Class_Terminate()
  
  Dispose

End Sub
