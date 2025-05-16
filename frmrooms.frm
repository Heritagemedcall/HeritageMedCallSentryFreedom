VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmRooms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rooms"
   ClientHeight    =   3180
   ClientLeft      =   105
   ClientTop       =   2295
   ClientWidth     =   9345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   60
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
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   5265
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mResidentID     As Long
Private mTransmitterID  As Long
Private mRoomID       As Long
Private quitting        As Boolean
Public Caller           As String

Private LastIndex   As Long
Function DeleteRoom(ByVal RoomID As Long)

        Dim AssignedTx         As Collection
        Dim device             As cESDevice
        Dim Resident           As cResident
        Dim ResidentID         As Long
        Dim SQL                As String
        Dim Rs                 As ADODB.Recordset


        'need to revisit and provide for remote update

10      If vbYes = messagebox(Me, "Delete Selected Room?", App.Title, vbQuestion Or vbYesNo) Then



20        If MASTER Then
30          Set AssignedTx = New Collection
40          If RoomID <> 0 Then
50            SQL = "SELECT Devices.serial FROM Devices WHERE  RoomID = " & RoomID
60            Set Rs = ConnExecute(SQL)
70            Do Until Rs.EOF
80              Set device = Devices.device(Rs("serial") & "")
90              If Not (device Is Nothing) Then
100               device.RoomID = 0
110               device.Room = ""

120             End If
130             Rs.MoveNext
140           Loop
150           Rs.Close
160           Set Rs = Nothing

170         End If
180       End If
190       If RoomID <> 0 Then

200         SQL = "SELECT residentID FROM residents WHERE RoomID = " & RoomID
210         Set Rs = ConnExecute(SQL)
220         Do Until Rs.EOF
230           ResidentID = Resident("ID")
240           For Each Resident In Residents
250             If Resident.ResidentID = ResidentID Then
260               Resident.RoomID = 0
270               Resident.Room = ""
280             End If
290           Next
300           Rs.MoveNext
310         Loop
320         Rs.Close
330         Set Rs = Nothing

340       End If

350       conn.BeginTrans
360       ConnExecute "UPDATE Residents SET RoomID = 0 WHERE RoomID = " & RoomID
370       ConnExecute "UPDATE Devices SET RoomID = 0 WHERE RoomID = " & RoomID
380       ConnExecute "UPDATE Devices SET RoomID_A = 0 WHERE RoomID_A = " & RoomID
390       ConnExecute "DELETE FROM Rooms WHERE RoomID = " & RoomID

400       conn.CommitTrans
410     End If

End Function


Public Property Get TransmitterID() As Long
  TransmitterID = mTransmitterID
End Property

Public Property Let TransmitterID(ByVal TransmitterID As Long)
  mTransmitterID = TransmitterID
  
End Property



Public Property Get ResidentID() As Long
  ResidentID = mResidentID
End Property

Public Property Let ResidentID(ByVal ResidentID As Long)
  mResidentID = ResidentID

End Property

Sub Apply()
  Dim RoomID As Long
  Dim SQL    As String

  Select Case Caller
    Case "RES"
      'If mResidentID <> 0 Then
      If lvMain.SelectedItem Is Nothing Then
        ' nada
        Beep
      Else
        RoomID = Val(lvMain.SelectedItem.Key)
        If MASTER Then
          SQL = "UPDATE devices SET RoomID = " & RoomID & "  WHERE deviceid = " & mTransmitterID
          ConnExecute (SQL)
          Devices.RefreshByID mTransmitterID
        Else
          ClientUpdateDeviceRoomID RoomID, mTransmitterID
          RefreshJet
        End If
        
        frmMain.SetListTabs
        PreviousForm
        Unload Me
      End If
      'End If
    Case "TX"
      If mTransmitterID <> 0 Then
        If lvMain.SelectedItem Is Nothing Then
          ' nada
        Else
          RoomID = Val(lvMain.SelectedItem.Key)
          If MASTER Then
            SQL = "UPDATE Devices SET RoomID = " & RoomID & "  WHERE deviceid = " & mTransmitterID
            ConnExecute (SQL)
            Devices.RefreshByID mTransmitterID
          Else
            ClientUpdateDeviceRoomID RoomID, mTransmitterID
            RefreshJet
          End If
          frmMain.SetListTabs
          PreviousForm
          Unload Me
          
        End If
      End If
    End Select





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
  cmdDelete.Enabled = gUser.UserPermissions.CanDeleteRooms = 1 Or gUser.LEvel = LEVEL_FACTORY
  
  ShowRooms
  
End Sub

Private Sub cmdAdd_Click()
  EditRoom 0
End Sub

Public Sub ShowRooms()
  DisableButtons
  cmdApply.Visible = (Caller > "")

  RefreshRooms
  EnableButtons
End Sub
Sub DisableButtons()
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdApply.Enabled = False
  cmdExit.Enabled = False
End Sub
Sub EnableButtons()
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
    'cmdDelete.Enabled = True And gUser.LEvel >= LEVEL_ADMIN
  cmdDelete.Enabled = gUser.UserPermissions.CanDeleteRooms = 1 Or gUser.LEvel = LEVEL_FACTORY
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub


Sub RefreshRooms()

  Dim Rs            As Recordset
  Dim li            As ListItem
  Dim index         As Long

  Dim Items         As Collection
  Dim Item          As cRoomListItem
  Dim SQL           As String



  On Error GoTo RefreshRooms_Error

  lvMain.ListItems.Clear
  LockWindowUpdate lvMain.hwnd

  Set Items = New Collection
  If gIsJET Then
    SQL = "SELECT roomid, room FROM Rooms order by room"
  Else
    SQL = "SELECT  Rooms.roomid, Rooms.Room, count(distinct Residents.ResidentID) AS CountOfResidents, Count(distinct Devices.DeviceID) AS CountOfDevices From (Rooms LEFT JOIN Devices ON Rooms.RoomID = Devices.RoomID) LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID GROUP BY Rooms.Room, Rooms.RoomID  ORDER BY Rooms.Room;"
  End If
  Set Rs = ConnExecute(SQL)




  Do Until Rs.EOF
    Set Item = New cRoomListItem
    Item.RoomID = Rs("RoomID")
    Item.Room = Rs("Room") & ""
    If gIsJET Then
      Item.ResidentCount = GetResidentCount(Item.RoomID)
      Item.DeviceCount = GetDeviceCount(Item.RoomID)
    Else
      Item.ResidentCount = Val(Rs("CountOfResidents") & "")
      Item.DeviceCount = Val(Rs("CountOfDevices") & "")

    End If
    Items.Add Item
    Rs.MoveNext
  Loop

  For Each Item In Items
    Set li = lvMain.ListItems.Add(, Item.RoomKey, Item.Room)
    li.SubItems(1) = Item.ResidentCount
    li.SubItems(2) = Item.DeviceCount
    If Item.RoomID = RoomID Then
      index = li.index
    End If
  Next

  If index > 0 Then
    lvMain.ListItems(index).Selected = True
    lvMain.ListItems(index).EnsureVisible
    LastIndex = 0
  End If

  If LastIndex > 0 And LastIndex <= lvMain.ListItems.Count Then
    lvMain.ListItems(LastIndex).Selected = True
    lvMain.ListItems(LastIndex).EnsureVisible
  End If
  'End If

RefreshRooms_Resume:
  On Error Resume Next
  Rs.Close
  Set Rs = Nothing
  LockWindowUpdate 0
  On Error GoTo 0
  Exit Sub

RefreshRooms_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmRooms.RefreshRooms." & Erl
  Resume RefreshRooms_Resume


End Sub

Sub Configurelvmain()
  Dim ch As ColumnHeader
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  
  Set ch = lvMain.ColumnHeaders.Add(, "Room", "Room", 2500, lvwColumnLeft)
  'Set ch = lvMain.ColumnHeaders.Add(, "Bldg", "Building")
  Set ch = lvMain.ColumnHeaders.Add(, "Res", "Residents", , lvwColumnCenter)
  Set ch = lvMain.ColumnHeaders.Add(, "Dev", "Devices", , lvwColumnCenter)
  lvMain.Sorted = False
End Sub
Function GetResidentCount(ByVal RoomID As Long) As String
  Dim Rs As Recordset
  Dim SQL As String
  Dim rsres As Recordset
  Dim Total As Long
  
  SQL = "select distinct residentid from devices where residentid > 0 and roomid = " & RoomID
  Set Rs = ConnExecute(SQL)
    Do Until Rs.EOF
      SQL = "select count(residentid) as numres from residents where residentid = " & Rs("residentID")
      Set rsres = ConnExecute(SQL)
      Total = Total + rsres("numres")
      Rs.MoveNext
    Loop
  Rs.Close
    
'  Set rs = connexecute("SELECT count(Residentid) FROM Residents WHERE RoomID = " & RoomID)
  GetResidentCount = CStr(Total)
  
End Function
Function GetDeviceCount(ByVal RoomID As Long) As String
  Dim Rs As Recordset
  Dim SQL As String
  Dim Total As Long
  
  SQL = "select distinct Deviceid from devices where roomid = " & RoomID
  Set Rs = ConnExecute(SQL)
    Do Until Rs.EOF
      Total = Total + 1
      Rs.MoveNext
    Loop
  Rs.Close
  Set Rs = Nothing
  GetDeviceCount = CStr(Total)
  
End Function


Private Sub cmdApply_Click()
  Apply
End Sub

Private Sub cmdDelete_Click()
  DeleteRoom Val(lvMain.SelectedItem.Key)
  Fill
  frmMain.SetListTabs
End Sub

Private Sub cmdEdit_Click()
  If lvMain.SelectedItem Is Nothing Then
    ' nada
  Else
    EditRoom Val(lvMain.SelectedItem.Key)
  End If
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub Form_Load()
  
  Configurelvmain
  quitting = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  quitting = True
  UnHost
End Sub

Public Property Get RoomID() As Long
  RoomID = mRoomID
End Property

Public Property Let RoomID(ByVal RoomID As Long)
  mRoomID = RoomID
End Property

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
End Sub
