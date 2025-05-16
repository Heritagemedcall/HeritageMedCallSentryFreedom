VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPushEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Serial As String
Public EventType As String
Public AlarmID As String
Public Announce As String
Public EventDate As String
Public FirstName As String
Public LastName As String
Public Room As String
Public Location As String


Public Retries As Long
Public IsValid As Boolean

Public Function CreateFromRowID(ByVal RowID As Long)
  ' BIG SQL HERE
  ' if ALARMID Column <> 0 then get the Parent RowID
  Dim Columns            As String
  Dim FromClause         As String
  Dim FirstJoin          As String
  Dim SecondJoin         As String
  Dim WhereClause        As String
  Dim SQL                As String
  Dim Rs                 As ADODB.Recordset


'SELECT
'Alarms.Serial, Alarms.ID, Alarms.AlarmID, Alarms.Announce, Alarms.EventDate, Alarms.EventType, Alarms.Userdata as Location,Residents.NameLast, Residents.NameFirst, Rooms.Room
'From alarms
'LEFT JOIN Rooms ON Alarms.RoomID = Rooms.RoomID
'LEFT  JOIN Residents ON Alarms.ResidentID = Residents.ResidentID
' Where (alarms.ID = 280307)

  Columns = "Alarms.Serial, Alarms.ID,Alarms.AlarmID, Alarms.Announce, Alarms.EventDate, Alarms.EventType, Alarms.Userdata as Location," & _
            "Residents.NameLast, Residents.NameFirst, Rooms.Room "

  FromClause = " FROM Alarms "
  FirstJoin = " LEFT JOIN Rooms ON Alarms.RoomID = Rooms.RoomID "
  SecondJoin = " LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID "
  WhereClause = " WHERE  (Alarms.ID = " & RowID & ") "

  SQL = "SELECT " & Columns & FromClause & FirstJoin & SecondJoin & WhereClause


  Set Rs = ConnExecute(SQL)
  If Not Rs Is Nothing Then
    If Not Rs.EOF Then
      IsValid = True
      Serial = Rs("Serial") & ""
      EventType = Rs("EventType") & ""
      If (Rs("AlarmID") = 0) Then
        AlarmID = RowID
      Else
        AlarmID = Rs("AlarmID")
      End If
      Announce = Rs("Announce") & ""
      EventDate = Format$(Rs("EventDate") & "", "yyyymmddhhnnss")
      FirstName = Rs("NameFirst") & ""
      LastName = Rs("NameLast") & ""
      Room = Rs("Room") & ""
      Location = Rs("Location") & ""
    End If
  End If
  Rs.Close
  Set Rs = Nothing

End Function

Public Function ToQueryString() As String
  
  ToQueryString = "?Serial=" & URLEncode(Serial) & "&EventType=" & URLEncode(EventType) & "&AlarmID=" & URLEncode(AlarmID) & _
    "&EventDate=" & URLEncode(EventDate) & "&Announce=" & URLEncode(Announce) & "&Fname=" & URLEncode(FirstName) & "&Lname=" & URLEncode(LastName) & "&Room=" & URLEncode(Room) & "&Location=" & URLEncode(Location)

End Function



