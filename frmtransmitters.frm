VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmTransmitters 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transmitters"
   ClientHeight    =   3030
   ClientLeft      =   3210
   ClientTop       =   5115
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
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
         Left            =   7725
         TabIndex        =   5
         Top             =   1785
         Visible         =   0   'False
         Width           =   1175
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Left            =   7725
         TabIndex        =   4
         Top             =   1200
         Width           =   1175
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         Left            =   7725
         TabIndex        =   3
         Top             =   615
         Width           =   1175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "New"
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
         Left            =   7725
         TabIndex        =   2
         Top             =   30
         Width           =   1175
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   7725
         TabIndex        =   6
         Top             =   2370
         Width           =   1175
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgLst"
         SmallIcons      =   "imgLst"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTransmitters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRoomID As Long
Private mResidentID As Long
Public OnlyPCAs     As Boolean
Public Serial       As String
Public returnobject As Object
Private quitting    As Boolean

Private LastIndex   As Long
Public Property Get RoomID() As Long
  RoomID = mRoomID
End Property

Public Property Let RoomID(ByVal RoomID As Long)
  mRoomID = RoomID
  If mRoomID <> 0 Then
    cmdApply.Visible = True
  End If
End Property

Public Property Get ResidentID() As Long
  ResidentID = mResidentID
End Property

Public Property Let ResidentID(ByVal ResidentID As Long)
  mResidentID = ResidentID
  If mResidentID <> 0 Then
    cmdApply.Visible = True
  End If
End Property

Sub Apply()
  Dim DeviceID  As Long
  Dim SQL       As String

  If lvMain.SelectedItem Is Nothing Then
    ' nada
    Beep
  Else
    DeviceID = Val(lvMain.SelectedItem.Key)

    If Not returnobject Is Nothing Then
      returnobject.text = lvMain.SelectedItem.text
      PreviousForm
      Unload Me
      
    ElseIf mResidentID <> 0 Then  ' we're setting for a resident
      If MASTER Then
        SQL = "UPDATE Devices SET ResidentID = " & ResidentID & " WHERE DeviceID = " & DeviceID
        ConnExecute SQL
        Devices.RefreshByID DeviceID
      Else
        ClientUpdateDeviceResidentID ResidentID, DeviceID
      End If
      frmMain.SetListTabs
      PreviousForm
      Unload Me
    ElseIf mRoomID <> 0 Then  ' we're setting for the room
      If MASTER Then
        SQL = "UPDATE Devices SET RoomID = " & RoomID & " WHERE DeviceID = " & DeviceID
        ConnExecute SQL
        Devices.RefreshByID DeviceID
      Else
        ClientUpdateDeviceRoomID RoomID, DeviceID
      End If
      frmMain.SetListTabs
      
      PreviousForm
      Unload Me
      
    End If
  End If
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

Public Sub Fill()
  ResetActivityTime
  'cmdDelete.Enabled = gUser.LEvel >= LEVEL_ADMIN
  cmdDelete.Enabled = gUser.UserPermissions.CanDeleteTransmitters = 1 Or gUser.LEvel = LEVEL_FACTORY
  ShowTransmitters
End Sub

Private Sub cmdAdd_Click()
  EditTransmitter 0
End Sub

Public Sub ShowTransmitters()
  RefreshJet
  DisableButtons
  RefreshDevices
  EnableButtons
End Sub
  

Sub Configurelvmain()
  Dim ch As ColumnHeader
  
  If lvMain.ColumnHeaders.Count < 6 Then
  
  
  lvMain.ColumnHeaders.Clear
  lvMain.Sorted = True
  Set ch = lvMain.ColumnHeaders.Add(, "S", "Serial", 10)
  Set ch = lvMain.ColumnHeaders.Add(, "M", "Model", 2300)
  
  Set ch = lvMain.ColumnHeaders.Add(, "Res", "Resident", 2100)
  Set ch = lvMain.ColumnHeaders.Add(, "Room", "Room", 2100)
 ' Set ch = lvMain.ColumnHeaders.Add(, "Bldg", "Building", 1350)
  Set ch = lvMain.ColumnHeaders.Add(, "Assur", "Assur", 700)
  
  
  End If
End Sub

Sub RefreshDevices()

10      On Error GoTo RefreshDevices_Error

20      DisableButtons

        Dim SQL        As String
        Dim Rs         As Recordset
        Dim li         As ListItem
        Dim j          As Integer

        Dim assure     As String
        Dim index      As Long
        Dim rsRoom     As Recordset

        Dim si         As Object

        Dim k As Long

        Dim SortedDevices As Collection

        Dim CurrentPass As Long
        Dim t As Long

        Dim Items As Collection

        't = Win32.timeGetTime

        Static passnumber As Long

30      passnumber = passnumber + 1
40      If passnumber >= MAXLONG Then
50        passnumber = 1
60      End If
70      CurrentPass = passnumber

80      Set SortedDevices = New Collection

90      lvMain.ListItems.Clear
100     LockWindowUpdate lvMain.hwnd
'110     lvMain.Sorted = True
'120     lvMain.SortKey = 0
130     lvMain.SortOrder = lvwAscending

140     If (OnlyPCAs) Then
150       SQL = " SELECT Serial,deviceid,ResidentID,RoomID,Model,UseAssur,UseAssur2,Assurinput,custom, ignored FROM Devices WHERE Devices.model = 'EN3954' ORDER BY Devices.Serial"
          'sql = " SELECT Serial,deviceid,ResidentID,RoomID,Model,UseAssur,UseAssur2,Assurinput, ignored FROM Devices WHERE Devices.model = 'EN3954' ORDER BY Devices.Serial"

160       SQL = " SELECT devices.Serial,devices.deviceid, devices.ResidentID,devices.RoomID,devices.Model,UseAssur,UseAssur2,Assurinput, ignored, devices.custom "
170       SQL = SQL & " residents.namelast, residents.nameFirst , rooms.room FROM "
180       SQL = SQL & " (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID WHERE Devices.model = 'EN3954' ORDER BY Devices.Serial"

190     Else
          'sql = " SELECT Serial,deviceid,ResidentID,RoomID,Model,UseAssur,UseAssur2,Assurinput, ignored FROM Devices ORDER BY Devices.Serial"
          'sql = " SELECT devices.Serial,devices.deviceid, devices.ResidentID,devices.RoomID,devices.Model,UseAssur,UseAssur2,Assurinput, ignored, "
          'sql = sql & " residents.namelast, residents.nameFirst FROM Devices  left join residents on (residents.residentID = devices.residentid) ORDER BY Devices.Serial"

200       SQL = " SELECT devices.Serial,devices.deviceid, devices.ResidentID, devices.custom, devices.RoomID,devices.Model,UseAssur,UseAssur2,Assurinput, ignored, "
210       SQL = SQL & " residents.namelast, residents.nameFirst , rooms.room FROM "
220       SQL = SQL & " (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID ORDER BY Devices.Serial"


230     End If
        '235     RefreshJet

240     Set Rs = ConnExecute(SQL)



        Dim Item    As cDeviceListItem

250     Set Items = New Collection

260     Do Until Rs.EOF
          'K = K + 1
          'If CurrentPass <> passnumber Then Exit Do
          ' DoEvents
          'If quitting Then Exit Do
270       Set Item = New cDeviceListItem
          
280       Item.DeviceID = Rs("deviceid")
290       Item.Serial = Rs("serial")
300       Item.Model = Rs("model")
310       Item.Ignored = Rs("ignored")
320       Item.ResidentID = Rs("residentid")
330       Item.AssurInput = Rs("assurinput")
340       Item.UseAssur = Rs("UseAssur")
350       Item.UseAssur2 = Rs("UseAssur2")
360       Item.NameFirst = Rs("namefirst") & ""
370       Item.NameLast = Rs("namelast") & ""
380       Item.Room = Rs("room") & ""
          Item.Custom = Rs("custom") & ""
          
          Item.Model = Item.Model & " " & Item.Custom
390       Items.Add Item, Item.Serial
400       Rs.MoveNext
410     Loop

420     Rs.Close
430     Set Rs = Nothing

        

440     For Each Item In Items
          'DoEvents
450       Set li = lvMain.ListItems.Add(, Item.ListKey, Right("00000000" & Item.Serial, 8))
460       li.SubItems(1) = Item.ModelAndStatus
470       li.ListSubItems(1).ForeColor = IIf(Item.Ignored And 1, vbRed, vbBlack)
480       li.SubItems(2) = Item.NameFull
490       li.SubItems(3) = Item.Room
500       li.SubItems(4) = Item.AssurString
          '    If OnlyPCAs Then  ' not sure shy this is here!!!!
510       If Item.Serial = Serial Then
520         index = li.index
530       End If
          '    End If

540     Next



        '  'rs.MoveFirst
        '  Do Until 1 = 1 'Do Until rs.EOF
        '    K = K + 1
        '    If CurrentPass <> passnumber Then Exit Do
        '    DoEvents
        '    If quitting Then Exit Do
        '    ' REMOTE todo
        '
        '    Set li = lvMain.ListItems.Add(, rs("DeviceID") & "B", Right("00000000" & rs("serial"), 8))
        '    li.SubItems(1) = rs("Model") & IIf(rs("ignored") And 1, " *X*", "")
        '
        '    li.ListSubItems(1).ForeColor = IIf(rs("ignored") And 1, vbRed, vbBlack)
        '
        '    li.SubItems(2) = GetResidentName(rs("ResidentID"))
        '    li.SubItems(2) = rs("namelast") & " " & rs("namefirst")
        '    sql = "SELECT * from rooms where RoomID = " & rs("RoomID")
        '    Set rsRoom = connexecute(sql)
        '    If Not rsRoom.EOF Then
        '      li.SubItems(3) = rsRoom("Room") & ""
        '
        '    Else
        '      li.SubItems(3) = " "
        '
        '    End If
        '    rsRoom.Close
        '    Set rsRoom = Nothing
        '    assure = IIf(rs("UseAssur") = 1, "Y", "N") & IIf(rs("UseAssur2") = 1, "Y", "N")
        '    If rs("UseAssur") = 1 Or rs("UseAssur2") = 1 Then
        '      assure = assure & rs("AssurInput")
        '    End If
        '
        '    '  assure = IIf(d.UseAssur Or d.UseAssur2 = 1, "Y", "N")
        '
        '    li.SubItems(4) = assure
        '    If OnlyPCAs Then  ' not sure shy this is here!!!!
        '      'If d.serial = serial Then
        '      index = li.index
        '      'End If
        '    End If
        '    rs.MoveNext
        '  Loop
        '


550     If index > 0 And index <= lvMain.ListItems.Count Then
560       lvMain.ListItems(index).Selected = True
570       lvMain.ListItems(index).EnsureVisible
580     End If



590     If LastIndex > 0 And index = 0 Then
600       If LastIndex <= lvMain.ListItems.Count Then
610         lvMain.ListItems(LastIndex).EnsureVisible
620         lvMain.ListItems(LastIndex).Selected = True
630       End If
640     End If

650     LockWindowUpdate 0
660     EnableButtons



RefreshDevices_Resume:

670     On Error Resume Next
680     If Not Rs Is Nothing Then
690       Rs.Close
700     End If
710     Set Rs = Nothing

720     On Error GoTo 0
730     Exit Sub

RefreshDevices_Error:

740     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitters.RefreshDevices." & Erl
750     Resume RefreshDevices_Resume

End Sub
Sub RefreshDevicesOLD()

10       On Error GoTo RefreshDevices_Error

20      DisableButtons

        Dim SQL        As String
        Dim Rs         As Recordset
        Dim li         As ListItem
        Dim j          As Integer
        Dim d          As cESDevice
        Dim assure     As String
        Dim index      As Long
        Dim rsRoom     As Recordset

        Dim k As Long

        Dim SortedDevices As Collection
30      Set SortedDevices = New Collection

40      lvMain.ListItems.Clear
50      LockWindowUpdate lvMain.hwnd
60      lvMain.Sorted = True
70      lvMain.SortKey = 0
80      lvMain.SortOrder = lvwAscending




90      If (OnlyPCAs) Then
100       SQL = " SELECT Serial,deviceid,ResidentID,RoomID,Model,UseAssur,UseAssur2,Assurinput FROM Devices WHERE Devices.model = 'EN3954' ORDER BY Devices.Serial"
110     Else
120       SQL = " SELECT Serial FROM Devices ORDER BY Devices.Serial"
130     End If

140     Set Rs = ConnExecute(SQL)

150     Do Until Rs.EOF
          k = k + 1
160       DoEvents
170       If quitting Then Exit Do
          ' REMOTE todo



180       For j = 1 To Devices.Devices.Count
190         Set d = Devices.Devices(j)
200         If d.Serial = Rs("Serial") Then
210           Exit For
220         End If
230         Set d = Nothing
240       Next
250       If Not d Is Nothing Then

260         Set li = lvMain.ListItems.Add(, d.DeviceID & "B", Right("00000000" & d.Serial, 8))
270         li.SubItems(1) = d.Model
280         li.SubItems(2) = GetResidentName(d.ResidentID)
290         SQL = "SELECT * from rooms where RoomID = " & d.RoomID
300         Set rsRoom = ConnExecute(SQL)
310         If Not rsRoom.EOF Then
320           li.SubItems(3) = rsRoom("Room") & ""
330           'li.SubItems(4) = rsRoom("building") & ""
340         Else
350           li.SubItems(3) = " "
360           'li.SubItems(4) = " "

370         End If
380         rsRoom.Close
390         Set rsRoom = Nothing
400         assure = IIf(d.UseAssur = 1, "Y", "N") & IIf(d.UseAssur2 = 1, "Y", "N")
410         If d.UseAssur = 1 Or d.UseAssur2 = 1 Then
420           assure = assure & d.AssurInput
430         End If

            '  assure = IIf(d.UseAssur Or d.UseAssur2 = 1, "Y", "N")
            'End If
440         li.SubItems(4) = assure
450         If OnlyPCAs Then
460           If d.Serial = Serial Then
470             index = li.index
480           End If
490         End If
500       Else
510         Set li = lvMain.ListItems.Add(, 0 & "B", Right("00000000" & Rs("Serial"), 8))
520         li.SubItems(1) = Rs("Serial")
530         li.SubItems(2) = "List Error"
540         li.SubItems(3) = " "
550         li.SubItems(4) = " "
560       End If
570       Rs.MoveNext
580     Loop
590     Rs.Close
600     Set Rs = Nothing

610     If index > 0 And index <= lvMain.ListItems.Count Then
620       lvMain.ListItems(index).Selected = True
630       lvMain.ListItems(index).EnsureVisible
640     End If
650     LockWindowUpdate 0
660     EnableButtons

RefreshDevices_Resume:
670      On Error GoTo 0
680      Exit Sub

RefreshDevices_Error:

690     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitters.RefreshDevices." & Erl
700     Resume RefreshDevices_Resume

End Sub

Function RoomFromResidentRoom(ByVal RoomID As Long) As String
  Dim SQL As String
  Dim Rs As Recordset

  SQL = "SELECT Room FROM Rooms WHERE RoomID = " & RoomID
  Set Rs = ConnExecute(SQL)
  If Not Rs.EOF Then
    RoomFromResidentRoom = Rs("Room") '& "," & rs("Building")
  End If
  Rs.Close

End Function


Private Sub cmdApply_Click()
  Apply
End Sub

Private Sub cmdDelete_Click()
  Dim i As Long
  Dim Key As Long
  
  If Not lvMain.SelectedItem Is Nothing Then
    If vbYes = messagebox(Me, "Delete Selected Device?", App.Title, vbQuestion Or vbYesNo) Then
      i = lvMain.SelectedItem.index
      Key = Val(lvMain.SelectedItem.Key)
      If MASTER Then
        DeleteTransmitter Key, gUser.Username
      Else
        RemoteDeleteTransmitter Key
      End If
      Fill
      DisableButtons
      If lvMain.ListItems.Count > 0 Then
        i = Min(i, lvMain.ListItems.Count)
        lvMain.ListItems(i).Selected = True
      End If
      frmMain.SetListTabs
      EnableButtons
    End If
  Else
    Beep
  End If
End Sub

Private Sub cmdEdit_Click()
  If Not lvMain.SelectedItem Is Nothing Then
    EditTransmitter Val(lvMain.SelectedItem.Key)
  Else
    Beep
  End If
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub Form_Load()

  quitting = False
  Configurelvmain
  lvMain.Sorted = True
  lvMain.SortKey = 3

  If (Not MASTER) Then
    If (USE6080) Then
      cmdAdd.Visible = False
    End If
  End If
  
  
End Sub

'Sub SetButtons()
'  If gUser.LEvel >= LEVEL_ADMIN Then
'    cmdDelete.Enabled = gUser.LEvel >= LEVEL_ADMIN
'  End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  quitting = True
  Set returnobject = Nothing
  UnHost
End Sub
Sub DisableButtons()
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdApply.Enabled = False
  cmdExit.Enabled = False
End Sub
Sub EnableButtons()
  cmdAdd.Enabled = True And cmdAdd.Visible
  cmdEdit.Enabled = True
  'cmdDelete.Enabled = True And gUser.LEvel >= LEVEL_ADMIN
  cmdDelete.Enabled = gUser.UserPermissions.CanDeleteTransmitters = 1 Or gUser.LEvel = LEVEL_FACTORY
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub

Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  If lvMain.SortKey = ColumnHeader.index - 1 Then
    If lvMain.SortOrder = lvwAscending Then
      lvMain.SortOrder = lvwDescending
    Else
      lvMain.SortOrder = lvwAscending
    End If
  Else
    lvMain.SortOrder = lvwAscending
  End If
  lvMain.SortKey = ColumnHeader.index - 1

End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    LastIndex = Item.index
  End If
  'ResetRemoteRefreshCounter
End Sub
