VERSION 5.00
Begin VB.Form frmImport 
   Caption         =   "Import Data"
   ClientHeight    =   3540
   ClientLeft      =   300
   ClientTop       =   2355
   ClientWidth     =   10305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   10305
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      Begin VB.CommandButton cmdTransmitters 
         Caption         =   "Import Transmitters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7605
         TabIndex        =   8
         Top             =   1905
         Width           =   1320
      End
      Begin VB.CommandButton cmdWaypoints 
         Caption         =   "Import Waypoints"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7605
         TabIndex        =   7
         Top             =   1320
         Width           =   1320
      End
      Begin VB.FileListBox lstFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   2625
         TabIndex        =   3
         Top             =   135
         Width           =   2700
      End
      Begin VB.DriveListBox lstDrives 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   2835
         Width           =   2490
      End
      Begin VB.DirListBox lstFolders 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   2535
      End
      Begin VB.CommandButton cmdNames 
         Caption         =   "Import  Names"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7605
         TabIndex        =   5
         Top             =   150
         Width           =   1320
      End
      Begin VB.CommandButton cmdRooms 
         Caption         =   "Import  Rooms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7605
         TabIndex        =   6
         Top             =   735
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7605
         TabIndex        =   9
         Top             =   2490
         Width           =   1320
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5430
         TabIndex        =   4
         Top             =   240
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private LastDrive As String
Private Sub EnableControls()
  lstFolders.Enabled = True
  lstDrives.Enabled = True
  lstFiles.Enabled = True
  cmdNames.Enabled = True
  cmdRooms.Enabled = True
  cmdWaypoints.Enabled = True
  cmdTransmitters.Enabled = True
  cmdExit.Enabled = True
  
End Sub
Private Sub DisableControls()
  
  lstFolders.Enabled = False
  lstDrives.Enabled = False
  lstFiles.Enabled = False
  cmdNames.Enabled = False
  cmdRooms.Enabled = False
  cmdWaypoints.Enabled = False
  cmdTransmitters.Enabled = False
  cmdExit.Enabled = False
  
  
End Sub

Private Sub cmdExit_Click()
  ResetActivityTime
  PreviousForm
  Unload Me
End Sub

Private Sub cmdNames_Click()
  ResetActivityTime
  DisableControls
  ImportNames
  EnableControls
  SetFocusTo cmdNames
End Sub

Private Sub cmdRooms_Click()
  ResetActivityTime
  DisableControls
  ImportRooms
  EnableControls
  SetFocusTo cmdRooms
End Sub
Function FullPath() As String
  Dim Path As String
  Path = lstFolders.Path
  If Right(Path, 1) <> "\" Then
    Path = Path & "\"
  End If
  FullPath = Path & lstFiles.filename
End Function
Sub ImportWaypoints()

        Dim hfile           As Integer
        Dim s               As String
        Dim rows()          As String
        Dim Header()        As String
        Dim Body()          As String
        Dim Col             As Integer
        Dim j               As Long
        Dim SQL             As String
        Dim Fields          As String
        Dim filename        As String
        Dim RowCount        As Long
        Dim LastCol         As Integer

        Dim Description     As String
        Dim Building        As String
        Dim Floor           As String
        Dim Wing            As String
        Dim Repeater1       As String
        Dim Repeater2       As String
        Dim Repeater3       As String
        Dim Signal1         As Integer
        Dim Signal2         As Integer
        Dim Signal3         As Integer

        Dim ColDescription  As Integer
        Dim colBuilding     As Integer
        Dim colFloor        As Integer
        Dim colWing         As Integer
        Dim colRepeater1    As Integer
        Dim colRepeater2    As Integer
        Dim colRepeater3    As Integer
        Dim ColSignal1      As Integer
        Dim ColSignal2      As Integer
        Dim ColSignal3      As Integer

10      On Error GoTo ImportWaypoints_Error

20      ClearError
30      If Len(lstFiles.filename) > 0 Then
40        filename = FullPath

50        hfile = FreeFile
60        Open filename For Binary As #hfile
70        s = Space(LOF(hfile))
80        Get #hfile, , s
90        Close hfile
100       rows = Split(s, vbCrLf)
110       RowCount = UBound(rows) - 1
120       ColDescription = -1
130       colBuilding = -1
140       colFloor = -1
150       colWing = -1
160       colRepeater1 = -1
170       colRepeater2 = -1
180       colRepeater3 = -1
190       ColSignal1 = -1
200       ColSignal2 = -1
210       ColSignal3 = -1

220       If RowCount > 0 Then
230         Header = Split(rows(0), vbTab)
240         For Col = LBound(Header) To UBound(Header)
250           Select Case LCase(Header(Col))
                Case ""
260               Exit For
270             Case "description", "waypoint"
280               ColDescription = Col
290             Case "floor"
300               colFloor = Col
310             Case "building"
320               colBuilding = Col
330             Case "wing"
340               colWing = Col
350             Case "repeater1"
360               colRepeater1 = Col
370             Case "repeater2"
380               colRepeater2 = Col
390             Case "repeater3"
400               colRepeater3 = Col
410             Case "signal1"
420               ColSignal1 = Col
430             Case "signal2"
440               ColSignal2 = Col
450             Case "signal3"
460               ColSignal3 = Col
470           End Select
480         Next

490         If ColDescription > -1 Then

500           For j = LBound(rows) + 1 To UBound(rows)
510             DoEvents
520             Body = Split(rows(j), vbTab)
530             Description = ""
540             Floor = ""
550             Building = ""
560             Wing = ""
570             Repeater1 = ""
580             Repeater2 = ""
590             Repeater3 = ""
600             Signal1 = 0
610             Signal2 = 0
620             Signal3 = 0


630             LastCol = UBound(Body)
640             If LastCol >= ColDescription Then

650               Description = Trim(Body(ColDescription))
660               If Len(Description) > 0 Then
670                 If colBuilding > -1 And LastCol >= colBuilding Then
680                   Building = Trim(Body(colBuilding))
690                 End If
700                 If colFloor > -1 And LastCol >= colFloor Then
710                   Floor = Trim(Body(colFloor))
720                 End If
730                 If colWing > -1 And LastCol >= colWing Then
740                   Wing = Trim(Body(colWing))
750                 End If

760                 If colRepeater1 > -1 And LastCol >= colRepeater1 Then
770                   Repeater1 = Trim(Body(colRepeater1))
780                 End If
790                 If colRepeater2 > -1 And LastCol >= colRepeater2 Then
800                   Repeater2 = Trim(Body(colRepeater2))
810                 End If
820                 If colRepeater3 > -1 And LastCol >= colRepeater3 Then
830                   Repeater3 = Trim(Body(colRepeater3))
840                 End If
850                 If ColSignal1 > -1 And LastCol >= ColSignal1 Then
860                   Signal1 = Val(Body(ColSignal1))
870                 End If
880                 If ColSignal2 > -1 And LastCol >= ColSignal2 Then
890                   Signal2 = Val(Body(ColSignal2))
900                 End If
910                 If ColSignal3 > -1 And LastCol >= ColSignal3 Then
920                   Signal3 = Val(Body(ColSignal3))
930                 End If

940                 Fields = Join(Array(q(Description), q(Building), q(Floor), q(Wing), q(Repeater1), q(Repeater2), q(Repeater3), Signal1, Signal2, Signal3), ",")
950                 SQL = "insert into Waypoints (Description, Building, Floor,Wing,Repeater1,Repeater2,Repeater3,Signal1,Signal2,Signal3)"
960                 SQL = SQL & " values ( " & Fields & ")"
970                 ConnExecute SQL
980               End If
990             Else
1000              Exit For
1010            End If
1020          Next
1030        End If
1040      End If
1050      lblMsg.Caption = "Done"
1060    Else
1070      lblMsg.Caption = "File Not Processed"
1080    End If



ImportWaypoints_Resume:
1090    On Error GoTo 0
1100    Exit Sub

ImportWaypoints_Error:

1110    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.ImportWaypoints." & Erl
1120    Resume ImportWaypoints_Resume


End Sub
Sub ImportTransmitters()

        Dim hfile           As Integer
        Dim s               As String
        Dim rows()          As String
        Dim Header()        As String
        Dim Body()          As String
        Dim Col             As Integer
        Dim j               As Long
        Dim i               As Long
        Dim SQL             As String
        'Dim Fields          As String
        Dim filename        As String
        Dim RowCount        As Long
        Dim LastCol         As Integer

        Dim Serial          As String
        Dim Model           As String

        Dim d               As cESDevice
        'Dim DeviceType      As cDeviceType

        Dim ColSerial       As Integer
        Dim ColModel        As Integer

        Dim RoomSerials     As Collection
        Dim RoomSerial      As cRoomSerial
        Dim Rooms           As Collection
        Dim Room            As String

        Dim ColRoom         As Integer
        'Dim HasRooms        As Boolean

        Dim RoomID          As Long

10      On Error GoTo ImportTransmitters_Error

20      ClearError
30      Set RoomSerials = New Collection
40      Set Rooms = New Collection

50      If Len(lstFiles.filename) > 0 Then
60        filename = FullPath

70        hfile = FreeFile
80        Open filename For Binary As #hfile
90        s = Space(LOF(hfile))
100       Get #hfile, , s
110       Close hfile
120       rows = Split(s, vbCrLf)
130       RowCount = UBound(rows)
140       ColSerial = -1
150       ColModel = -1

160       If RowCount > 0 Then
170         Header = Split(rows(0), vbTab)
180         For Col = LBound(Header) To UBound(Header)
190           Select Case LCase(Header(Col))
                Case ""
200               Exit For
210             Case "serial", "transmitter", "device", "deviceid", "device id", "id"
220               ColSerial = Col
230             Case "model", "type"
240               ColModel = Col
250             Case "room", "suite", "unit", "rooms"
260               ColRoom = Col
270           End Select
280         Next

290         If ColSerial > -1 And ColModel > -1 Then

300           For j = LBound(rows) + 1 To UBound(rows)
310             DoEvents
320             Body = Split(rows(j), vbTab)
330             Serial = ""
340             Model = ""
350             Room = ""
360             Set d = Nothing

370             LastCol = UBound(Body)
380             If LastCol >= ColSerial Then
390               Serial = ValidateSerial(Body(ColSerial))
400               If ColModel > -1 And LastCol >= ColModel Then
410                 Model = Trim(Body(ColModel))
420               End If
430               If ColRoom > -1 And LastCol >= ColRoom Then
440                 Room = Trim(Body(ColRoom))
450                 Rooms.Add Room
460               End If

470               If Len(Serial) = 8 And Len(Model) > 0 Then
                    ' get details for model

480                 For i = 0 To MAX_ESDEVICETYPES
490                   If 0 = StrComp(ESDeviceType(i).Model, Model, vbTextCompare) Then  ' must have both model and serial

500                     Set d = Devices.device(Serial)
510                     If d Is Nothing Then
520                       Set d = New cESDevice

530                       d.Serial = Serial
540                       d.Model = ESDeviceType(i).Model
550                       d.ClearByReset = ESDeviceType(i).ClearByReset
560                       d.SupervisePeriod = ESDeviceType(i).Checkin
570                       d.Announce = ESDeviceType(i).Announce
580                       d.Announce_A = ESDeviceType(i).Announce2
590                       d.IsPortable = ESDeviceType(i).Portable
                          
600                       d.AutoClear = ESDeviceType(i).AutoClear
                          d.IgnoreTamper = ESDeviceType(i).IgnoreTamper

610                       d.NumInputs = ESDeviceType(i).NumInputs
620                       d.Description = ESDeviceType(i).desc
630                       d.IsLatching = ESDeviceType(i).Latching
640                       d.IsPortable = ESDeviceType(i).Portable
650                       d.NoTamper = ESDeviceType(i).NoTamper
660                       d.SendCancel = ESDeviceType(i).SendCancel
670                       d.SendCancel_A = ESDeviceType(i).SendCancel_A
680                       d.RepeatUntil = ESDeviceType(i).RepeatUntil
690                       d.RepeatUntil_A = ESDeviceType(i).RepeatUntil_A
700                       d.Repeats = ESDeviceType(i).Repeats
710                       d.Repeats_A = ESDeviceType(i).Repeats_A
720                       d.Pause = ESDeviceType(i).Pause
730                       d.Pause_A = ESDeviceType(i).Pause_A
                          'd.AlarmMask = 0
                          'd.AlarmMask_A = 0

740                       d.NumInputs = ESDeviceType(i).NumInputs

750                       d.OG1 = ESDeviceType(i).OG1
760                       d.OG2 = ESDeviceType(i).OG2
770                       d.OG3 = ESDeviceType(i).OG3
780                       d.OG4 = ESDeviceType(i).OG4
790                       d.OG5 = ESDeviceType(i).OG5
800                       d.OG6 = ESDeviceType(i).OG6


810                       d.NG1 = ESDeviceType(i).NG1
820                       d.NG2 = ESDeviceType(i).NG2
830                       d.NG3 = ESDeviceType(i).NG3
840                       d.NG4 = ESDeviceType(i).NG4
850                       d.NG5 = ESDeviceType(i).NG5
860                       d.NG6 = ESDeviceType(i).NG6

870                       d.OG1_A = ESDeviceType(i).OG1_A
880                       d.OG2_A = ESDeviceType(i).OG2_A
890                       d.OG3_A = ESDeviceType(i).OG3_A
900                       d.OG4_A = ESDeviceType(i).OG4_A
910                       d.OG5_A = ESDeviceType(i).OG5_A
920                       d.OG6_A = ESDeviceType(i).OG6_A

930                       d.NG1_A = ESDeviceType(i).NG1_A
940                       d.NG2_A = ESDeviceType(i).NG2_A
950                       d.NG3_A = ESDeviceType(i).NG3_A
960                       d.NG4_A = ESDeviceType(i).NG4_A
970                       d.NG5_A = ESDeviceType(i).NG5_A
980                       d.NG6_A = ESDeviceType(i).NG6_A


990                       d.OG1D = ESDeviceType(i).OG1D
1000                      d.OG2D = ESDeviceType(i).OG2D
1010                      d.OG3D = ESDeviceType(i).OG3D
1020                      d.OG4D = ESDeviceType(i).OG4D
1030                      d.OG5D = ESDeviceType(i).OG5D
1040                      d.OG6D = ESDeviceType(i).OG6D

1050                      d.NG1D = ESDeviceType(i).NG1D
1060                      d.NG2D = ESDeviceType(i).NG2D
1070                      d.NG3D = ESDeviceType(i).NG3D
1080                      d.NG4D = ESDeviceType(i).NG4D
1090                      d.NG5D = ESDeviceType(i).NG5D
1100                      d.NG6D = ESDeviceType(i).NG6D

1110                      d.OG1_AD = ESDeviceType(i).OG1_AD
1120                      d.OG2_AD = ESDeviceType(i).OG2_AD
1130                      d.OG3_AD = ESDeviceType(i).OG3_AD
1140                      d.OG4_AD = ESDeviceType(i).OG4_AD
1150                      d.OG5_AD = ESDeviceType(i).OG5_AD
1160                      d.OG6_AD = ESDeviceType(i).OG6_AD

1170                      d.NG1_AD = ESDeviceType(i).NG1_AD
1180                      d.NG2_AD = ESDeviceType(i).NG2_AD
1190                      d.NG3_AD = ESDeviceType(i).NG3_AD
1200                      d.NG4_AD = ESDeviceType(i).NG4_AD
1210                      d.NG5_AD = ESDeviceType(i).NG5_AD
1220                      d.NG6_AD = ESDeviceType(i).NG6_AD





1230                      If SaveDevice(d) Then  ' successfully saved
1240                        Devices.AddDevice d  ' add it
1250                        Devices.RefreshBySerial Serial  ' refresh it
1260                        SetupSerialDevice d  ' configure serial settings
1270                        Set RoomSerial = New cRoomSerial
1280                        RoomSerial.Room = Room
1290                        RoomSerial.Serial = Serial
1300                        RoomSerials.Add RoomSerial

1310                        Select Case UCase(d.Model)
                              Case "EN5000", "EN5040"  ' create outbound message to set NID
1320                            Outbounds.AddMessage d.Serial, MSGTYPE_REPEATERNID, "", 0
1330                          Case "EN3954"  ' create outbound message to set NID
1340                            Outbounds.AddMessage d.Serial, MSGTYPE_TWOWAYNID, "", 0
1350                        End Select
1360                        Debug.Print "Imported " & d.Serial & " " & d.Model
1370                      End If

1380                    End If
1390                    Exit For
1400                  End If
1410                Next
1420              End If
1430            Else
1440              Exit For
1450            End If
1460          Next
1470        End If
1480      End If

          ' get distinct rooms
1490      Set Rooms = RemoveDupeRooms(Rooms)

          ' add rooms to database
1500      AppendRooms Rooms

1510      For j = 1 To RoomSerials.Count
1520        Room = RoomSerials(j).Room

1530        If Len(Room) > 0 Then  ' is there a room number?
1540          Serial = RoomSerials(j).Serial
1550          If Len(Serial) > 0 Then  ' is there a serial number ?
                'get roomID for room
1560            RoomID = GetRoomID(Room)
1570            If RoomID <> 0 Then
1580              SQL = "UPDATE Devices SET RoomID = " & RoomID & " WHERE Serial = " & q(Serial)
1590              ConnExecute SQL
1600              Devices.RefreshBySerial Serial
1610            End If
1620          End If
1630        End If
1640      Next
          ' add rooms to system
1650      lblMsg.Caption = "Done"
1660    Else
1670      lblMsg.Caption = "File Not Processed"
1680    End If

ImportTransmitters_Resume:
1690    On Error GoTo 0
1700    Exit Sub

ImportTransmitters_Error:

1710    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.ImportTransmitters." & Erl
1720    Resume ImportTransmitters_Resume


End Sub
Private Function GetRoomID(ByVal roomname As String) As Long
        Dim Rs      As Recordset
        Dim SQL     As String

10      On Error GoTo GetRoomID_Error

20      GetRoomID = 0
30      SQL = "select roomid from rooms where room = " & q(roomname)
40      Set Rs = ConnExecute(SQL)
50      If Not (Rs.EOF) Then
60        GetRoomID = IIf(IsNull(Rs("roomid")), 0, Rs("roomid"))
70      End If
80      Rs.Close
90      Set Rs = Nothing




GetRoomID_Resume:
100     On Error GoTo 0
110     Exit Function

GetRoomID_Error:

120     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.GetRoomID." & Erl
130     Resume GetRoomID_Resume


End Function
Private Function AppendRooms(Rooms As Collection) As Integer
        Dim Rs      As Recordset
        Dim j       As Integer
        Dim Count   As Long
        Dim Room    As String
        Dim Fields  As String
        Dim SQL     As String

10       On Error GoTo AppendRooms_Error

20      For j = 1 To Rooms.Count
30        Room = Rooms(j)
40        SQL = "select count(*) from rooms where room = " & q(Room)
50        Set Rs = ConnExecute(SQL)
60        If Rs(0) < 1 Then
70          Count = Count + 1
80          Fields = Join(Array(q(Room), q(""), q(""), q(""), 0, 0, 0, 0), ",")
90          SQL = "insert into rooms (room, building, locator,lockw,assurdays,vacation,away,deleted) "
100         SQL = SQL & " values ( " & Fields & ")"
110         ConnExecute SQL
120       End If
130       Rs.Close
140     Next
150     Set Rs = Nothing
160     AppendRooms = Count


AppendRooms_Resume:
170      On Error GoTo 0
180      Exit Function

AppendRooms_Error:

190     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.AppendRooms." & Erl
200     Resume AppendRooms_Resume


End Function
Private Function RemoveDupeRooms(Rooms As Collection) As Collection
        ' sorts list and removes duplicate rooms
  
        Dim j         As Integer
        Dim i         As Integer
        Dim Prev      As String
        Dim Current   As String
        Dim RoomList  As Collection

        Dim NewRooms  As Collection

10       On Error GoTo RemoveDupeRooms_Error
        

20      Set NewRooms = SortRooms(Rooms)
30      Set RoomList = New Collection

40      For j = 1 To NewRooms.Count
50        Current = NewRooms(j)
60        If Len(Current) > 0 Then
70          If 0 <> StrComp(Current, Prev, vbTextCompare) Then
80            RoomList.Add Current
90            Prev = Current
100         End If
110       End If
120     Next
130     Set RemoveDupeRooms = RoomList
140     Set RoomList = Nothing
150     Set NewRooms = Nothing

RemoveDupeRooms_Resume:
160      On Error GoTo 0
170      Exit Function

RemoveDupeRooms_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.RemoveDupeRooms." & Erl
190     Resume RemoveDupeRooms_Resume

End Function
Public Function SortRooms(Rooms As Collection) As Collection  ' get in decending order
        Dim offset      As Integer
        Dim maxrow      As Integer
        Dim limit       As Integer
        Dim row         As Integer
        Dim switch      As Integer
        Dim MinRow      As Integer
        Dim Temp        As String
        Dim SortedRooms As Collection

        Dim a() As String
  
10       On Error GoTo SortRooms_Error

20      maxrow = Rooms.Count
  
30      ReDim a(1 To maxrow)
40      For row = 1 To maxrow
50         a(row) = Rooms(row)
60      Next
  
  
70      MinRow = 1

80      offset = maxrow \ 2
90      Do While offset > 0
100       limit = maxrow - offset
110       Do
120         switch = 0
130         For row = MinRow To limit
              'If a(Row) < a(Row + Offset).Level Then  ' may need to incorporate margin
140           If -1 = StrComp(a(row), a(row + offset), vbTextCompare) Then
150             Temp = a(row)
160             a(row) = a(row + offset)
170             a(row + offset) = Temp
180             switch = row
190           End If
200         Next row
210         limit = switch - offset
220       Loop While switch

230       offset = offset \ 2
240     Loop
  
250     Set SortedRooms = New Collection
  
260     For row = 1 To maxrow
270       SortedRooms.Add a(row)
280     Next
290     Set SortRooms = SortedRooms

SortRooms_Resume:
300      On Error GoTo 0
310      Exit Function

SortRooms_Error:

320     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.SortRooms." & Erl
330     Resume SortRooms_Resume

End Function

Function SaveDevice(d As cESDevice) As Boolean
        Dim Rs            As Recordset
        Dim Exists        As Boolean
        Dim Saved         As Boolean

10      On Error GoTo SaveDevice_Error

20      Set Rs = ConnExecute("SELECT COUNT(*) FROM devices WHERE Serial = " & q(d.Serial))
30      Exists = Rs(0).value > 0
40      Rs.Close
50      Set Rs = Nothing
60      If Exists Then
70        Saved = False
80      Else

90        Set Rs = New Recordset
100       Rs.Open "devices", conn, gCursorType, gLockType
110       Rs.addnew

          ' these are set to actual values
120       Rs("serial") = d.Serial
130       Rs("Model") = d.Model

140       Rs("IsPortable") = d.IsPortable

150       Rs("NumInputs") = d.NumInputs
160       Rs("Announce") = d.Announce
170       Rs("Announce_A") = d.Announce_A
180       Rs("Announce_b") = d.Announce_B
190       Rs("ClearByReset") = d.ClearByReset
200       Rs("deleted") = 0

          ' these default to zero or empty strings
210       Rs("Assurinput") = 0
220       Rs("UseAssur") = 0
230       Rs("UseAssur2") = 0
240       Rs("UseAssur_A") = 0
250       Rs("UseAssur2_A") = 0
260       Rs("UseAssur_B") = 0
270       Rs("UseAssur2_B") = 0


280       Rs("OG1") = d.OG1
290       Rs("OG2") = d.OG2
300       Rs("OG3") = d.OG3
310       Rs("OG4") = d.OG4
320       Rs("OG5") = d.OG5
330       Rs("OG6") = d.OG6


340       Rs("NG1") = d.NG1
350       Rs("NG2") = d.NG2
360       Rs("NG3") = d.NG3
370       Rs("NG4") = d.NG4
380       Rs("NG5") = d.NG5
390       Rs("NG6") = d.NG6

400       Rs("gG1") = d.GG1
410       Rs("gG2") = d.GG2
420       Rs("gG3") = d.GG3
430       Rs("gG4") = d.GG4
440       Rs("gG5") = d.GG5
450       Rs("gG6") = d.GG6



460       Rs("OG1_A") = d.OG1_A
470       Rs("OG2_A") = d.OG2_A
480       Rs("OG3_A") = d.OG3_A
490       Rs("OG4_A") = d.OG4_A
500       Rs("OG5_A") = d.OG5_A
510       Rs("OG6_A") = d.OG6_A


520       Rs("NG1_A") = d.NG1_A
530       Rs("NG2_A") = d.NG2_A
540       Rs("NG3_A") = d.NG3_A
550       Rs("NG4_A") = d.NG4_A
560       Rs("NG5_A") = d.NG5_A
570       Rs("NG6_A") = d.NG6_A

580       Rs("gG1_A") = d.GG1_A
590       Rs("gG2_A") = d.GG2_A
600       Rs("gG3_A") = d.GG3_A
610       Rs("gG4_A") = d.GG4_A
620       Rs("gG5_A") = d.GG5_A
630       Rs("gG6_A") = d.GG6_A


640       Rs("OG1_b") = d.OG1_B
650       Rs("OG2_b") = d.OG2_B
660       Rs("OG3_b") = d.OG3_B
670       Rs("OG4_b") = d.OG4_B
680       Rs("OG5_b") = d.OG5_B
690       Rs("OG6_b") = d.OG6_B


700       Rs("NG1_b") = d.NG1_B
710       Rs("NG2_b") = d.NG2_B
720       Rs("NG3_b") = d.NG3_B
730       Rs("NG4_b") = d.NG4_B
740       Rs("NG5_b") = d.NG5_B
750       Rs("NG6_b") = d.NG6_B

760       Rs("gG1_b") = d.GG1_B
770       Rs("gG2_b") = d.GG2_B
780       Rs("gG3_b") = d.GG3_B
790       Rs("gG4_b") = d.GG4_B
800       Rs("gG5_b") = d.GG5_B
810       Rs("gG6_b") = d.GG6_B


820       Rs("OG1d") = d.OG1D
830       Rs("OG2d") = d.OG2D
840       Rs("OG3d") = d.OG3D
850       Rs("OG4d") = d.OG4D
860       Rs("OG5d") = d.OG5D
870       Rs("OG6d") = d.OG6D


880       Rs("NG1d") = d.NG1D
890       Rs("NG2d") = d.NG2D
900       Rs("NG3d") = d.NG3D
910       Rs("NG4d") = d.NG4D
920       Rs("NG5d") = d.NG5D
930       Rs("NG6d") = d.NG6D


940       Rs("gG1d") = d.GG1D
950       Rs("gG2d") = d.GG2D
960       Rs("gG3d") = d.GG3D
970       Rs("gG4d") = d.GG4D
980       Rs("gG5d") = d.GG5D
990       Rs("gG6d") = d.GG6D


1000      Rs("OG1_Ad") = d.OG1_AD
1010      Rs("OG2_Ad") = d.OG2_AD
1020      Rs("OG3_Ad") = d.OG3_AD
1030      Rs("OG4_Ad") = d.OG4_AD
1040      Rs("OG5_Ad") = d.OG5_AD
1050      Rs("OG6_Ad") = d.OG6_AD


1060      Rs("NG1_Ad") = d.NG1_AD
1070      Rs("NG2_Ad") = d.NG2_AD
1080      Rs("NG3_Ad") = d.NG3_AD
1090      Rs("NG4_Ad") = d.NG4_AD
1100      Rs("NG5_Ad") = d.NG5_AD
1110      Rs("NG6_Ad") = d.NG6_AD

1120      Rs("gG1_Ad") = d.GG1_AD
1130      Rs("gG2_Ad") = d.GG2_AD
1140      Rs("gG3_Ad") = d.GG3_AD
1150      Rs("gG4_Ad") = d.GG4_AD
1160      Rs("gG5_Ad") = d.GG5_AD
1170      Rs("gG6_Ad") = d.GG6_AD

            ' b

1180      Rs("OG1_bd") = d.OG1_BD
1190      Rs("OG2_bd") = d.OG2_BD
1200      Rs("OG3_bd") = d.OG3_BD
1210      Rs("OG4_bd") = d.OG4_BD
1220      Rs("OG5_bd") = d.OG5_BD
1230      Rs("OG6_bd") = d.OG6_BD


1240      Rs("NG1_bd") = d.NG1_BD
1250      Rs("NG2_bd") = d.NG2_BD
1260      Rs("NG3_bd") = d.NG3_BD
1270      Rs("NG4_bd") = d.NG4_BD
1280      Rs("NG5_bd") = d.NG5_BD
1290      Rs("NG6_bd") = d.NG6_BD

1300      Rs("gG1_bd") = d.GG1_BD
1310      Rs("gG2_bd") = d.GG2_BD
1320      Rs("gG3_bd") = d.GG3_BD
1330      Rs("gG4_bd") = d.GG4_BD
1340      Rs("gG5_bd") = d.GG5_BD
1350      Rs("gG6_bd") = d.GG6_BD




1360      Rs("VacationSuper") = 0
1370      Rs("VacationSuper_A") = 0
1380      Rs("VacationSuper_B") = 0
1390      Rs("SendCancel") = d.SendCancel
1400      Rs("SendCancel_A") = d.SendCancel_A
1410      Rs("SendCancel_b") = d.SendCancel_B
          
1420      Rs("DisableStart") = 0
1430      Rs("DisableEnd") = 0
1440      Rs("DisableStart_A") = 0
1450      Rs("DisableEnd_A") = 0
1460      Rs("DisableStart_b") = 0
1470      Rs("DisableEnd_b") = 0
          
1480      Rs("UseTamperAsInput") = 0
          
1490      Rs("repeatuntil") = d.RepeatUntil
1500      Rs("repeatuntil_A") = d.RepeatUntil_A
1510      Rs("repeatuntil_b") = d.RepeatUntil_B
1520      Rs("repeats") = d.Repeats
1530      Rs("repeats_A") = d.Repeats_A
1540      Rs("repeats_b") = d.Repeats_B
1550      Rs("Pause") = d.Pause
1560      Rs("Pause_A") = d.Pause_A
1570      Rs("Pause_b") = d.Pause_B
1580      Rs("AlarmMask") = d.AlarmMask
1590      Rs("AlarmMask_A") = d.AlarmMask_A
1600      Rs("AlarmMask_b") = d.AlarmMask_B
1610      Rs("ResidentID") = 0
1620      Rs("ResidentID_A") = 0
1630      Rs("RoomID") = 0
1640      Rs("RoomID_A") = 0
1650      Rs("RoomD_A") = 0  ' bogus field
1660      Rs("ClearByReset_A") = 0  ' d.ClearByReset
1670      Rs("SerialTapProtocol") = 0
1680      Rs("SerialSkip") = 0
1690      Rs("SerialMessageLen") = 0
1700      Rs("SerialAutoClear") = 0
1710      Rs("SerialPort") = 0
1720      Rs("SerialBaud") = 0
1730      Rs("SerialBits") = 0
1740      Rs("SerialParity") = ""
1750      Rs("SerialStopBits") = ""
1760      Rs("SerialInclude") = ""
1770      Rs("SerialExclude") = ""
1780      Rs("SerialFlow") = 0
1790      Rs("SerialEOLChar") = 0
1800      Rs("SerialPreamble") = ""
1810      Rs("SerialSettings") = ""
1820      Rs("IDM") = 0
1830      Rs("IDL") = 0
1840      Rs("ignored") = 0
1850      Rs("ignoretamper") = d.IgnoreTamper

1860      Rs("custom") = ""

1865      Rs("LocKW") = ""

          ' temperature device
1870      Rs("lowset") = d.LowSet
1880      Rs("lowset_a") = d.LowSet_A
1890      Rs("hiset") = d.HiSet
1900      Rs("hiset_a") = d.HiSet_A
1910      Rs("EnableTemp") = d.EnableTemperature
1920      Rs("EnableTemp_a") = d.EnableTemperature_A



1930      Rs.Update
1940      Rs.Close
1950      Set Rs = Nothing
1960      Saved = True
1970    End If


SaveDevice_Resume:
1980    SaveDevice = (Saved And (Err.Number = 0))
1990    On Error GoTo 0

2000    Exit Function

SaveDevice_Error:

2010    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.SaveDevice." & Erl
2020    Resume SaveDevice_Resume


End Function

Function ValidateSerial(ByVal Serial As String) As String
  Dim j               As Integer
  Dim ValidHexValue   As Boolean
  
  Serial = UCase(Trim(Serial))
  If Len(Serial) < 1 Then
    ' bail
  ElseIf Len(Serial) <= 8 Then
    For j = 1 To Len(Serial)
      ValidHexValue = InStr("0123456789ABCDEF", MID(Serial, j, 1)) > 0
      If Not (ValidHexValue) Then
        ' invalid character, bail
        Exit For
      End If
    Next
    If ValidHexValue Then
      ' all checked out, pad with leading zeros
      ValidateSerial = Right("00000000" & Serial, 8)
    Else
      ' bail
    End If
  End If
    
End Function

Sub ImportNames()
        Dim hfile As Integer
        Dim s As String
        Dim rows() As String
        Dim Header() As String
        Dim Body()   As String
        Dim ColLast  As Integer
        Dim ColFirst As Integer
        Dim ColRoom  As Integer
        Dim ColPhone As Integer
        Dim Col      As Integer
        Dim j        As Long
        Dim SQL As String
        Dim NameLast As String
        Dim NameFirst As String
        Dim Phone As String
        Dim Room As String
        Dim Fields As String
        Dim filename As String
        Dim RowCount As Long

10      On Error GoTo ImportNames_Error

20      ClearError
30      If Len(lstFiles.filename) > 0 Then
40        filename = FullPath


50        hfile = FreeFile
60        Open filename For Binary As #hfile
70        s = Space(LOF(hfile))
80        Get #hfile, , s
90        Close hfile
100       rows = Split(s, vbCrLf)
110       RowCount = UBound(rows) - 1
120       ColLast = -1
130       ColFirst = -1
140       ColPhone = -1
150       ColRoom = -1

160       If RowCount > 0 Then
170         Header = Split(rows(0), vbTab)
180         For Col = LBound(Header) To UBound(Header)
190           Select Case LCase(Header(Col))
                Case ""
200               Exit For
210             Case "last", "lastname", "last name", "name last", "namelast"
220               ColLast = Col
230             Case "first", "firstname", "first name", "name first", "namefirst"
240               ColFirst = Col
250             Case "room"
260               ColRoom = Col
270             Case "ph", "phone", "phone number", "phonenumber"
280               ColPhone = Col
290           End Select
300         Next

310         For j = LBound(rows) + 1 To UBound(rows)
              DoEvents
320           Body = Split(rows(j), vbTab)
330           NameLast = ""
340           NameFirst = ""
350           Phone = ""
360           Room = ""
370           If UBound(Body) >= 1 Then
380             If ColLast > -1 Then
390               NameLast = Trim(Body(ColLast))
400             End If
410             If ColFirst > -1 Then
420               NameFirst = Trim(Body(ColFirst))
430             End If
440             If ColPhone > -1 Then
450               Phone = Trim(Body(ColPhone))
460             End If
470             If Len(NameLast) > 0 Then
                  'room = body(ColRoom)
480               Fields = Join(Array(q(""), q(NameLast), q(NameFirst), q(""), 1, 0, 0, q(""), 254, 0, 0, q(""), q(Phone), 0, q("")), ",")
490               SQL = "insert into residents (name, namelast, namefirst,room,active,groupid,roomid,imagepath,assurdays,vacation,away,info, phone,Deleted,deliverypoints) "
500               SQL = SQL & " values ( " & Fields & ")"
510               ConnExecute SQL
520             End If
530           Else
540             Exit For
550           End If
560         Next

570       End If
580       lblMsg.Caption = "Done"
590     Else
600       lblMsg.Caption = "File Not Processed"
610     End If

ImportNames_Resume:
620     On Error GoTo 0
630     Exit Sub

ImportNames_Error:

640     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.ImportNames." & Erl
650     lblMsg.Caption = "Error, File Not Processed"
660     Resume ImportNames_Resume

End Sub
Sub ClearError()
  lblMsg.Caption = ""
End Sub
Sub ImportRooms()
        Dim hfile As Integer
        Dim s As String
        Dim rows() As String
        Dim Header() As String
        Dim Body()   As String
        Dim ColLast  As Integer
        Dim ColFirst As Integer
        Dim ColRoom  As Integer
        Dim ColPhone As Integer
        Dim Col      As Integer
        Dim j        As Long
        Dim SQL As String
        Dim NameLast As String
        Dim NameFirst As String
        Dim Phone As String
        Dim Room As String
        Dim Fields As String
        Dim filename As String
        Dim RowCount As Long
        Dim Rs As Recordset


10      On Error GoTo ImportRooms_Error

20      ClearError
30      If Len(lstFiles.filename) > 0 Then
40        filename = FullPath


50        hfile = FreeFile
60        Open filename For Binary As #hfile
70        s = Space(LOF(hfile))
80        Get #hfile, , s
90        Close hfile
100       rows = Split(s, vbCrLf)
110       RowCount = UBound(rows) - 1
120       ColRoom = -1
130       If RowCount > 0 Then
140         Header = Split(rows(0), vbTab)
150         For Col = LBound(Header) To UBound(Header)
160           Select Case LCase(Header(Col))
                Case ""
170               Exit For
180             Case "last", "lastname", "last name", "name last", "namelast"
190               ColLast = Col
200             Case "first", "firstname", "first name", "name first", "namefirst"
210               ColFirst = Col
220             Case "room", "suite", "unit"
230               ColRoom = Col
240             Case "ph", "phone", "phone number", "phonenumber"
250               ColPhone = Col
260           End Select
270         Next

280         For j = LBound(rows) + 1 To UBound(rows)
              DoEvents
290           Body = Split(rows(j), vbTab)
300           NameLast = ""
310           NameFirst = ""
320           Phone = ""
330           Room = ""
340           If UBound(Body) >= ColRoom Then
345             Room = Trim(Body(ColRoom))
350             If Len(Room) > 0 Then

370               Fields = Join(Array(q(Room), q(""), q(""), q(""), 254, 0, 0, 0, 0), ",")
380               SQL = " select count(*) from rooms where room = " & q(Room)
390               Set Rs = ConnExecute(SQL)
400               If Rs(0) < 1 Then
410                 SQL = "insert into rooms (room, building, locator,lockw,assurdays,vacation,away,deleted,flags) "
420                 SQL = SQL & " values ( " & Fields & ")"
430                 ConnExecute SQL
440               End If
450               Rs.Close
460             End If
470           Else
480             Exit For
490           End If
500         Next

510       End If
520       lblMsg.Caption = "Done"
530     Else
540       lblMsg.Caption = "File Not Processed"
550     End If


ImportRooms_Resume:
560     On Error GoTo 0
570     Exit Sub

ImportRooms_Error:

580     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmImport.ImportRooms." & Erl
590     lblMsg.Caption = "Error, File Not Processed"
600     Resume ImportRooms_Resume


End Sub



Private Sub Command1_Click()

End Sub

Private Sub cmdTransmitters_Click()
  ResetActivityTime
  DisableControls
  ImportTransmitters
  EnableControls
  SetFocusTo cmdTransmitters
End Sub

Private Sub cmdWaypoints_Click()
  ResetActivityTime
  DisableControls
  If USE6080 Then
    ConvertRoomsToPartitions
  Else

    ImportWaypoints
  End If
    EnableControls
    SetFocusTo cmdWaypoints

End Sub
Sub ConvertRoomsToPartitions()
  



End Sub


Private Sub Form_Activate()

EnableControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select

End Sub

Private Sub Form_Load()
  ResetActivityTime
  Connect
  EnableControls
  fraEnabler.BackColor = Me.BackColor
  lstFiles.Pattern = "*.txt"
  lstFiles.Path = lstFolders.Path
  LastDrive = lstDrives.Drive

End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub

Private Sub lstFiles_Click()
  ClearError
End Sub

Private Sub lstFolders_Change()
  ClearError
  Static Busy As Boolean

  On Error Resume Next

  If Not Busy Then

    Busy = True
    ' Change the current directory
    ChDir lstFolders.Path

    If Err.Number = 0 Then
      lstFiles.Path = lstFolders.Path
      lstDrives.Drive = left(lstFolders.Path, 2)
    Else
      Err.Clear
    End If

    Busy = False

  End If

End Sub

Private Sub lstDrives_Change()
  ClearError
  Dim Retry As Boolean
  On Error Resume Next

  Retry = True
  Do While Retry
    Retry = False
    lstFiles.Path = lstFolders.Path
    lstFolders.Path = lstDrives.Drive
    ' If and error occurs
    Select Case Err.Number
      Case 68  ' Not accessable Error
        If vbRetry = messagebox(Me, lstDrives.Drive & " is not accessible", App.Title, vbRetryCancel Or vbCritical) Then
          Retry = True
        Else  'Switch to previous known drive
          lstDrives.Drive = LastDrive
        End If
      Case 0
        ' done
      Case Else
        ' Ooops
        messagebox Me, "Unexpected File Access Error " & Err.Number & " : " & Err.Description, App.Title, vbInformation
        lstDrives.Drive = LastDrive
    End Select
  Loop

End Sub
Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub


