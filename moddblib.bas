Attribute VB_Name = "modDBLib"
Option Explicit

Global je As JetEngine

Global DateDelim As String

Public Function ConnectionString() As String
  If FileExists(App.Path & "\freedom2.udl") Then
    ' remote connections MUST NOT have spaces between key phrases and '=' (equal signs)
    ConnectionString = "File Name=" & App.Path & "\freedom2.udl"
  Else
    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & App.Path & "\freedom2.mdb" & ";Mode=Share Deny None"
  End If
End Function

Public Function Connect()
  Dim conn2              As ADODB.Connection
  Dim Timeout As Long
  
  If conn Is Nothing Then
    Set conn = New ADODB.Connection
    'timeout = conn.CommandTimeout
  End If
  
  Dim Command As ADODB.Command
  
  Set Command = CreateObject("ADODB.command")
  'Set command.ActiveConnection = conn
  Timeout = Command.CommandTimeout
  
  
  
  If InIDE Then
    ' Debug.Assert 0
  End If
  If conn.State = ADODB.adStateClosed Then
    'Set conn2 = conn
    'conn.provider = "Microsoft.Jet.OLEDB.4.0"



    On Error Resume Next
    '             Debug.Print "Path " & conn.Properties("Jet OLEDB:Registry Path")
    '        '  Debug.Print "Lock Delay " & conn.Properties("Jet OLEDB:Lock Delay")
    '          Debug.Print "Max Locks " & conn.Properties("Jet OLEDB:Max Locks Per File")
    '          Debug.Print "Conn Mode " & conn.Mode
    '          Debug.Print "Tx Timeout " & conn.Properties("Jet OLEDB:Flush Transaction Timeout")

    conn.Open ConnectionString
    'Debug.Print "Path " & conn.Properties("Jet OLEDB:Registry Path")
    '
    '          Debug.Print "Lock Delay " & conn.Properties("Jet OLEDB:Lock Delay")
    '          Debug.Print "Max Locks " & conn.Properties("Jet OLEDB:Max Locks Per File")
    Debug.Print "Conn Mode " & conn.Mode
    '          Debug.Print "Tx Timeout " & conn.Properties("Jet OLEDB:Flush Transaction Timeout")

    'Debug.Print conn.Properties("'Jet OLEDB:Registry Path'") & " " & Len(conn.Properties("Jet OLEDB:Registry Path"))

    ' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Jet 4.0
    Set conn2 = conn
    If InStr(1, conn.ConnectionString, ".Jet.", vbTextCompare) > 0 Then
      gIsJET = True
      Set je = New JetEngine
      DateDelim = "#"
      'If conn.State = ADODB.adStateOpen Then
      'End If

    ElseIf InStr(1, conn.ConnectionString, ".ACE.", vbTextCompare) > 0 Then
      gIsJET = True
      Set je = Nothing
      DateDelim = "#"
    Else
      DateDelim = "'"
    End If




  End If

End Function

Function FireHoseRecordSet(SQL As String) As ADODB.Recordset
  Dim Rs As ADODB.Recordset
  Dim Retries As Long
  Dim rc
  
  On Error Resume Next
  Retries = 3
  
  Set Rs = New ADODB.Recordset
  Do
    Err.Clear
    Retries = Retries - 1
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic
    If Retries < 1 Then
      Exit Do
    End If
    If Err.Number <> 0 Then
      LogProgramError "FHRS Could Not Retreive Query Data. " & vbCrLf & ConnectionString & vbCrLf & SQL & vbCrLf & Err.Description
    End If
    
  Loop While Err.Number <> 0
  Set FireHoseRecordSet = Rs
  Set Rs = Nothing


End Function

Function ConnExecute(SQL As String) As ADODB.Recordset
        Dim Rs            As ADODB.Recordset
        Dim ErrNum As Long
        Dim ErrDesc As String

10      On Error Resume Next
20      Err.Clear
30      Set Rs = conn.Execute(SQL)

40      If Err.Number <> 0 Then
50        LogProgramError "Could Not Execute Query, Retrying Connection. " & vbCrLf & ConnectionString & vbCrLf & SQL & vbCrLf & Err.Description
60        conn.Close
70        Err.Clear
80        Sleep 100
90        Connect
100       Sleep 100
110       If Err.Number <> 0 Then
120           ErrNum = Err.Number
130           ErrDesc = Err.Description
140       End If
150       If ErrNum <> 0 Then
160         LogProgramError "Could Not Re-Connect. " & vbCrLf & ConnectionString & vbCrLf & Err.Description
170       Else
180         Set Rs = conn.Execute(SQL)
190         If Err.Number <> 0 Then
200           LogProgramError "Could Not Retreive Query Data. " & vbCrLf & ConnectionString & vbCrLf & SQL & vbCrLf & Err.Description
210         End If
220       End If
230     End If
240     Set ConnExecute = Rs

        Set Rs = Nothing


End Function

Sub RefreshJet()
10      On Error GoTo RefreshJet_Error
20      If Not (gIsJET) Then Exit Sub
        ' trying to avoid jet crashes  11/2/2009
30      If MASTER Then
40        Sleep 1
50      Else

60        Sleep 100
70      End If
80      Exit Sub
        'Dim t                  As Long
90      't = Win32.timeGetTime

100     On Error Resume Next
110     If gIsJET Then

120       If Not je Is Nothing Then
130         If MASTER Then
              ' no refresh
140         Else
              'If Not conn Is Nothing Then
              '  If conn.State <> ADODB.adStateClosed Then
              '    conn.Close
              '  End If
150           je.RefreshCache conn

              'End If
              ' Connect
160         End If
170       End If
180     End If

RefreshJet_Resume:
190     't = Win32.timeGetTime - t
200   'Debug.Print "refresh jet time " & t
210     On Error GoTo 0
220     Exit Sub

RefreshJet_Error:

230     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.RefreshJet." & Erl
240     Resume RefreshJet_Resume

End Sub

Sub UpdateAssignedTx(Rs As Recordset)
        Dim Device As cESDevice
        Dim ResidentID As Long

10      ResidentID = Rs("residentID")

20      Set Device = Devices.Device(Rs("serial"))

30      If Device Is Nothing Then  ' wasn't loaded into system
40        Set Device = New cESDevice
50        Device.Parse Rs
60        Device.FetchResident
70        Device.FetchRoom
          
80      ElseIf Device.ResidentID <> ResidentID Then  ' loaded but not assigned to this person
90        Device.Parse Rs
100       Device.FetchResident
110       Device.FetchRoom
        
120     Else
          ' no need to update
130     End If

End Sub


Function UpdateResident(Resident As cResident, ByVal User As String) As Boolean
        Dim addnew As Boolean
        Dim Rs As ADODB.Recordset
        
        Dim ClearAlarms As Integer

10      Set Rs = New ADODB.Recordset
20      Rs.Open "SELECT * FROM Residents WHERE residentID = " & Resident.ResidentID, conn, gCursorType, gLockType
        'conn.BeginTrans

30      If Rs.EOF Then
40        addnew = True
50        Rs.addnew
60        Rs("Active") = 1
70        Resident.Vacation = 0
80      Else
90        ClearAlarms = IIf(Resident.Vacation = Rs("Away"), 0, 1)
100     End If

110     Rs("namelast") = Resident.NameLast
120     Rs("nameFirst") = Resident.NameFirst
130     Rs("phone") = Resident.Phone
140     Rs("room") = Resident.Room
150     Rs("info") = Resident.info
160     Rs("AssurDays") = Resident.Assurdays And &HFF
170     Rs("Vacation") = Resident.Vacation  ' error bug fix , we''l probably  delete this column
180     Rs("Away") = Resident.Vacation
190     Rs("Deliverypoints") = Resident.DeliveryPointsString
200     Rs("Deleted") = 0
210     Rs.Update
220     Rs.MoveLast
230     Resident.ResidentID = Rs("ResidentID")

240     Rs.Close

        Dim tx As cESDevice

250     For Each tx In Resident.AssignedTx
260       tx.ResidentID = Resident.ResidentID
270       ConnExecute "UPDATE devices SET  ResidentID = " & Resident.ResidentID & " WHERE devices.deviceid = " & tx.DeviceID
280       Set Rs = ConnExecute("SELECT Devices.* FROM Devices,Devicetypes WHERE Devices.model = Devicetypes.model AND deviceid = " & tx.DeviceID)
290       If Not Rs.EOF Then
300         UpdateAssignedTx Rs
310       End If
320       Rs.Close
330     Next

340     If conn.Errors.Count = 0 Then
          'conn.CommitTrans
          'If ClearAlarms Then
          '  frmMain.ClearAwayAlarms Resident.ResidentID
          'End If
350     Else
          'conn.RollbackTrans
360       Exit Function
370     End If
380     Set Rs = Nothing

390     If addnew Then
400       LogAddRes Resident.ResidentID
410     End If

        'If Away <> resident.Vacation Then
420     ConnExecute "UPDATE Residents SET Away = " & Resident.Vacation & " WHERE ResidentID = " & Resident.ResidentID
        '  LogVacation resident.ResidentID, 0, resident.Vacation
        '  End If

430     SetDevicesAwayByResident Resident.ResidentID, Resident.Vacation
        
        
        

440     UpdateResident = True
End Function

Function UpdateStaff(Resident As cResident, ByVal User As String) As Boolean

  Dim addnew As Boolean
  Dim Rs As ADODB.Recordset
  Dim byteArray(0) As Byte

  Set Rs = New ADODB.Recordset
  Rs.Open "SELECT * FROM Staff WHERE StaffID = " & Resident.ResidentID, conn, gCursorType, gLockType
  'conn.BeginTrans

  If Rs.EOF Then
    addnew = True
    Rs.addnew
    Rs("Active") = 1
    Resident.Vacation = 0
  Else
    'Resident.Vacation = IIf(rs("Away") = 1, 1, 0)
  End If

  Rs("namelast") = Resident.NameLast
  Rs("nameFirst") = Resident.NameFirst
  Rs("name") = Resident.NameLast & ", " & Resident.NameFirst
  Rs("phone") = Resident.Phone
  Rs("room") = Resident.Room
  Rs("info") = Resident.info
  Rs("AssurDays") = Resident.Assurdays And &HFF
  Rs("Vacation") = Resident.Vacation  ' error bug fix , we''l probably  delete this column
  Rs("Away") = Resident.Vacation
  Rs("groupid") = 0
  Rs("roomid") = 0
  Rs("Imagepath") = ""
  Rs("Imagedata") = byteArray
  Rs("Deleted") = 0
  Rs("Deliverypoints") = Resident.DeliveryPointsString
  Rs.Update
  Rs.MoveLast
  Resident.ResidentID = Rs("staffID")

  Rs.Close

  

  If conn.Errors.Count = 0 Then
    'conn.CommitTrans

  Else
    'conn.RollbackTrans
    Exit Function
  End If
  Set Rs = Nothing

  'If addnew Then
  '  LogAddRes Resident.ResidentID
  'End If

  'If Away <> resident.Vacation Then
  'connexecute "UPDATE Residents SET Away = " & Resident.Vacation & " WHERE ResidentID = " & Resident.ResidentID
  '  LogVacation resident.ResidentID, 0, resident.Vacation
  'End If

  'SetDevicesAwayByResident Resident.ResidentID, Resident.Vacation

  UpdateStaff = True
End Function



Function FieldToString(dbfield As Variant) As String
  FieldToString = dbfield & ""
End Function
Function FieldToNumber(dbfield As Variant) As Double
  If IsNull(dbfield) Then
    FieldToNumber = 0
  Else
    FieldToNumber = dbfield
  End If
End Function

Function GetResidentRoomID(ByVal ResidentID As Long) As Long


  Dim Rs As Recordset
  If (ResidentID) Then
  Set Rs = ConnExecute("SELECT RoomID FROM Devices WHERE ResidentID = " & ResidentID)
  If Not Rs.EOF Then
    GetResidentRoomID = Val("" & Rs("RoomID"))
  End If
  Rs.Close
  End If
End Function


Function GetRoomName(ByVal RoomID As Long) As String
  Dim Rs As Recordset
10         On Error GoTo GetRoomName_Error

20        If (RoomID) Then
30          Set Rs = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & RoomID)
40          If Not Rs.EOF Then
50            GetRoomName = Rs("Room") & ""
60          End If
70          Rs.Close
80        End If

GetRoomName_Resume:
90         On Error GoTo 0
100        Exit Function

GetRoomName_Error:

110       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetRoomName." & Erl
120       Resume GetRoomName_Resume

End Function

Public Function GetResidentName(ByVal ResidentID As Long) As String
  Dim Rs As Recordset
10         On Error GoTo GetResidentName_Error

20        Set Rs = ConnExecute("SELECT NameLast, NameFirst FROM residents WHERE residentid = " & ResidentID)
30        If Not Rs.EOF Then
40          GetResidentName = Rs("NameLast") & ", " & Rs("NameFirst")
50        End If
60        Rs.Close
70        Set Rs = Nothing

GetResidentName_Resume:
80         On Error GoTo 0
90         Exit Function

GetResidentName_Error:

100       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetResidentName." & Erl
110       Resume GetResidentName_Resume


End Function
Public Function GetResidentRooms(ByVal ResidentID As Long) As String
  Dim Rs As Recordset
  Dim SQL As String
  Dim Rooms As String
  Dim t As Long
'  t = Win32.timeGetTime()
  
  SQL = " SELECT Distinct Rooms.Room FROM Devices INNER JOIN Rooms ON Devices.RoomID = Rooms.RoomID WHERE Devices.ResidentID = " & ResidentID
  Set Rs = ConnExecute(SQL)
  Do Until Rs.EOF
    If Len(Rooms) Then
      Rooms = Rooms & "\"
    End If
    Rooms = Rooms & Rs("Room") & ""
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing

  GetResidentRooms = Rooms

  'Debug.Print "Get Resident Rooms " & Win32.timeGetTime() - t

End Function
'Public Function GetRoomNameByRoomID(ByVal RoomID As Long) As String
'  Dim rs As Recordset
'  Dim sql As String
'
'  sql = "SELECT RoomID, Room FROM rooms WHERE RoomID = " & RoomID
'  Set rs = connexecute(sql)
'  Do Until rs.EOF
'
'    rs.MoveNext
'  Loop
'  rs.Close
'
'End Function


Public Function GetResidentRoom(ByVal RoomID As Long) As String
  Dim Rs As Recordset
10         On Error GoTo GetResidentRoom_Error

20        Set Rs = ConnExecute("SELECT Room FROM residents WHERE residentid = " & RoomID)
30        If Not Rs.EOF Then
40          GetResidentRoom = Rs("Room") & ""
50        End If
60        Rs.Close
70        Set Rs = Nothing

GetResidentRoom_Resume:
80         On Error GoTo 0
90         Exit Function

GetResidentRoom_Error:

100       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetResidentRoom." & Erl
110       Resume GetResidentRoom_Resume


End Function

Public Function GetLocater(ByVal Serial As String) As String
  Dim SQL As String
  Dim Rs As Recordset

10         On Error GoTo GetLocater_Error

20        SQL = " SELECT Rooms.Room, Rooms.Locator, Rooms.Building FROM Devices LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
      " WHERE  Devices.serial =" & q(Serial)

30        Set Rs = ConnExecute(SQL)
40        If Not Rs.EOF Then
50          GetLocater = Rs("Room") & ""
60        End If
70        Rs.Close

GetLocater_Resume:
80         On Error GoTo 0
90         Exit Function

GetLocater_Error:

100       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetLocater." & Erl
110       Resume GetLocater_Resume


End Function

Public Function GetUser(ByVal Password As String) As cUser
        Dim Rs                 As Recordset
        Dim User               As cUser

10      On Error GoTo GetUser_Error
20      Set User = New cUser

30      If 0 = StrComp(Password, FACTORY_PWD, vbTextCompare) Then

40        User.UserID = 0
50        User.LEvel = LEVEL_FACTORY
60        User.Username = "Factory"
70        User.Password = Password
80        User.ConsoleID = ConsoleID
          User.UserPermissions.SetUserPermissions 1, 1, 1
          
          'user.Session = GetNextSession
          'LOGOFF OTHER ADMINS

90      Else

100       Set Rs = ConnExecute("SELECT * FROM Users WHERE Password = " & q(Password))
110       If Not Rs.EOF Then
120         User.UserID = Rs("userid")
130         User.LEvel = Rs("Level")
140         User.Username = Rs("UserName") & ""
            'Debug.Assert 0
            User.UserPermissions.ParseUserPermissions Val(Rs("permissions") & "")
150         User.Password = Password
160         User.ConsoleID = ConsoleID
170         If Not MASTER Then
180           User.Session = GetNextSession
190         End If
200       End If
210       Rs.Close

220     End If

230     Set GetUser = User

GetUser_Resume:
240     On Error GoTo 0
250     Exit Function

GetUser_Error:

260     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetUser." & Erl
270     Resume GetUser_Resume

End Function
'
'
'Public Function GetUserold(ByVal Password As String) As cUser
'        Dim rs          As Recordset
'        Dim user        As cUser
'
'10       On Error GoTo GetUser_Error
'        Set user = New cUser
'
'        If 0 = StrComp(Password, FACTORY_PWD, vbTextCompare) Then
'
'           user.UserID = 0
'           user.Level = LEVEL_FACTORY
'           user.username = "Factory"
'          user.Password = Password
'        Else
'20
'
'30      Set rs = connexecute("SELECT * FROM Users WHERE Password = " & Q(Password))
'40      If Not rs.EOF Then
'50        user.UserID = rs("userid")
'60        user.Level = rs("Level")
'70        user.username = rs("UserName") & ""
'80        user.Password = Password
'90      End If
'100     rs.Close
'        End If
'110     Set GetUserold = user
'
'GetUser_Resume:
'120      On Error GoTo 0
'130      Exit Function
'
'GetUser_Error:
'
'140     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetUser." & Erl
'150     Resume GetUser_Resume
'
'End Function

Function GetRoomIDFromResID(ByVal ResidentID As Long) As Long
  Dim Rs As Recordset
  Set Rs = ConnExecute("SELECT RoomID FROM Residents WHERE ResidentID = " & ResidentID)
  If Not Rs.EOF Then
    If Not IsNull(Rs("roomid")) Then
      GetRoomIDFromResID = Rs("roomid")
    End If
  End If
  Rs.Close
  Set Rs = Nothing


End Function

Function GetResidentIDFromSerial(ByVal Serial As String) As Long
  Dim SQL As String
  Dim Rs As Recordset
  SQL = " SELECT DeviceID, Serial, ResidentID, RoomID FROM Devices WHERE serial =" & q(Serial)
  Set Rs = ConnExecute(SQL)
  If Not Rs.EOF Then
    GetResidentIDFromSerial = Rs("ResidentID")
  End If
  Rs.Close
  Set Rs = Nothing

End Function

Function GetResidentIDFromDeviceID(ByVal DeviceID As Long) As Long
  Dim SQL As String
  Dim Rs As Recordset
  SQL = " SELECT  ResidentID FROM Devices WHERE DeviceID =" & DeviceID
  Set Rs = ConnExecute(SQL)
  If Not Rs.EOF Then
    GetResidentIDFromDeviceID = IIf(IsNull(Rs("ResidentID")), 0, Rs("ResidentID"))
  End If
  Rs.Close
  Set Rs = Nothing

End Function

Function GetRoomIDFromSerial(ByVal Serial As String) As Long
  Dim SQL As String
  Dim Rs As Recordset
  SQL = " SELECT DeviceID, Serial, ResidentID, RoomID FROM Devices WHERE serial =" & q(Serial)
  Set Rs = ConnExecute(SQL)
  If Not Rs.EOF Then
    GetRoomIDFromSerial = Rs("RoomID")
  End If
  Rs.Close
  Set Rs = Nothing

End Function


Sub StartSession()
10      On Error GoTo StartSession_Error

20      Connect
        Dim Rs As Recordset
30      Set Rs = New Recordset

40      Rs.Open "sessions", conn, gCursorType, gLockType
50      Rs.addnew

60      Rs("StartTime") = Now
70      Rs("LastPing") = Now
80      Rs("QuitTime") = 0
90      Rs.Update

100     Rs.MoveLast
110     gSessionID = Rs("sessionID")
120     Rs.Close
130     Set Rs = Nothing

StartSession_Resume:

140     On Error GoTo 0
150     Exit Sub

StartSession_Error:

160     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.StartSession." & Erl
170     Resume StartSession_Resume

End Sub
Sub QuitSession()
  Dim SQL As String

  PingSession
  SQL = "UPDATE Sessions SET QuitTime = " & q(Now) & " WHERE SessionID = " & gSessionID
  ConnExecute SQL
End Sub
Sub PingSession()
  ConnExecute "UPDATE Sessions SET LastPing = " & q(Now) & " WHERE SessionID = " & gSessionID
End Sub

Function ExistsUser0000() As Boolean
  Dim SQL As String
  Dim Rs As ADODB.Recordset
  SQL = "SELECT count(*) as User0000 FROM users WHERE password = '0000'"
  Set Rs = ConnExecute(SQL)
  On Error Resume Next
  ExistsUser0000 = Rs(0) > 0
  Rs.Close
  Set Rs = Nothing
  
End Function
