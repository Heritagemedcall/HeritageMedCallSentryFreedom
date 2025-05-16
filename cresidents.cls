VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResidents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Residents As Collection



Private Sub Class_Initialize()
  Set Residents = New Collection
End Sub

Private Sub Class_Terminate()
  Dispose
End Sub
Private Sub Dispose()
  If Not Residents Is Nothing Then
    Set Residents = Nothing
  End If
End Sub

Public Property Get Resident(ByVal ID As Long) As cResident
  On Error Resume Next
  If Not MASTER Then Exit Function
  If Residents Is Nothing Then
    Exit Property
  End If
  Set Resident = Residents.Item(ID & "R")

End Property

Public Function AddResident(Resident As cResident) As Long
  Dim r As cResident
  If Not MASTER Then Exit Function
  
  On Error Resume Next
  Set r = Residents(Resident.ResidentID & "R")
  If r Is Nothing Then
    Set r = New cResident
    Residents.Add Resident, (Resident.ResidentID & "R")
    AddResident = Residents.Count
  Else
    Exit Function
  End If
  
End Function
Public Sub ClearAll()
  Set Residents = New Collection
End Sub

Public Function ParseAndAdd(rs As ADODB.Recordset) As Long
  Dim r As cResident
  Set r = New cResident
  r.Parse rs
  Residents.Add r, r.ResidentID & "R"
End Function

Public Function FetchAndUpdate(ResidentID As Long) As Long
  Dim rs                 As ADODB.Recordset
  Dim ActiveRes          As cResident
  Set ActiveRes = Resident(ResidentID)
  If Not ActiveRes Is Nothing Then
    Set rs = ConnExecute("SELECT * from Residents where residentID = " & ResidentID)
    If Not rs.EOF Then
      ActiveRes.Parse rs
    End If
    rs.Close
    Set rs = Nothing
  End If

End Function
