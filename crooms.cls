VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Rooms As Collection

Private Sub Class_Initialize()
  Set Rooms = New Collection
End Sub

Private Sub Class_Terminate()
  Dispose
End Sub
Private Sub Dispose()


  
  If Not Rooms Is Nothing Then
    Set Rooms = Nothing
  End If
End Sub

Public Property Get Room(ByVal id As Long) As cRoom
  On Error Resume Next
  If Not MASTER Then Exit Property
  If Rooms Is Nothing Then
    Exit Property
  End If
  Set Room = Rooms.item(id & "R")

End Property

Public Function AddRoom(Room As cRoom) As Long
  Dim r As cRoom
  If Not MASTER Then Exit Function
  
  On Error Resume Next
  Set r = Rooms(Room.RoomID & "R")
  If r Is Nothing Then
    Set r = New cRoom
    Rooms.Add Room, (Room.RoomID & "R")
    AddRoom = Rooms.Count
  Else
    Exit Function
  End If
  
End Function
Public Sub ClearAll()
  Set Rooms = New Collection
End Sub

Public Function ParseAndAdd(rs As ADODB.Recordset) As Long
  Dim r As cRoom
  Set r = New cRoom
  r.Parse rs
  Rooms.Add r, r.RoomID & "R"
End Function


