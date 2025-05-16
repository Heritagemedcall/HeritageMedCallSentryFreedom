Attribute VB_Name = "modDialogic"
Option Explicit
Global DialogicSystem As DialogicSystem
'Global DialogicSystem As Object

Sub SetPriorityChannels()
  Dim Rs      As ADODB.Recordset
  Dim j       As Long
  Dim SQL     As String
  On Error Resume Next
  
  If Not MASTER Then Exit Sub
  
  If DialogicSystem Is Nothing Then
    Set DialogicSystem = New DialogicSystem
  End If
  
  
  For j = 1 To Diva.MAX_CHANNELS
    DialogicSystem.Reserved(j) = False
  Next
  
  SQL = "SELECT DialerModem FROM PagerDevices WHERE ProtocolID = " & PROTOCOL_DIALOGIC
  
  Set Rs = ConnExecute(SQL)
  Do Until Rs.EOF
    DialogicSystem.Reserved(Val(Rs("DialerModem") & "")) = True
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing


End Sub
