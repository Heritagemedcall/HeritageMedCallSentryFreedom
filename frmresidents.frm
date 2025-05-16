VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmResidents 
   Caption         =   "Residents"
   ClientHeight    =   3165
   ClientLeft      =   405
   ClientTop       =   5265
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   9555
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
Attribute VB_Name = "frmResidents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRoomID As Long
Private mTransmitterID As Long
Private mResidentID As Long
Private quitting As Boolean
Public Caller As String
Private LastIndex   As Long
Public Property Get TransmitterID() As Long
  TransmitterID = mTransmitterID
End Property

Public Property Let TransmitterID(ByVal TransmitterID As Long)
  mTransmitterID = TransmitterID
  If mTransmitterID <> 0 Then
    cmdApply.Visible = True
  End If
End Property


Public Property Get RoomID() As Long
  RoomID = mRoomID
End Property

Public Property Let RoomID(ByVal RoomID As Long)
  mRoomID = RoomID
  If mRoomID <> 0 Then
    cmdApply.Visible = True
  End If
End Property
Sub Apply()
  Dim ResidentID As Long
  Dim SQL    As String

  If mRoomID <> 0 Then
    If lvMain.SelectedItem Is Nothing Then
      ' nada
    Else
      ResidentID = Val(lvMain.SelectedItem.Key)
      If MASTER Then
        SQL = "UPDATE Devices SET RoomID = " & RoomID & "  WHERE  Deviceid = " & mTransmitterID
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
  ElseIf mTransmitterID <> 0 Then
    If lvMain.SelectedItem Is Nothing Then
      ' nada
    Else
      ResidentID = Val(lvMain.SelectedItem.Key)
      If MASTER Then
        SQL = "UPDATE Devices SET residentID = " & ResidentID & "  WHERE Deviceid = " & mTransmitterID
        ConnExecute (SQL)
        Devices.RefreshByID mTransmitterID
      Else
        ClientUpdateDeviceResidentID ResidentID, mTransmitterID
      
        RefreshJet
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
  On Error Resume Next
  ResetActivityTime
  'cmdDelete.Enabled = gUser.LEvel >= LEVEL_ADMIN
  cmdDelete.Enabled = gUser.UserPermissions.CanDeleteResidents = 1 Or gUser.LEvel = LEVEL_FACTORY
  ShowResidents
End Sub

Private Sub cmdAdd_Click()
  EditResident 0
End Sub

Public Sub ShowResidents()
  RefreshJet
  RefreshResidents
End Sub

Sub RefreshResidents()

  Dim SQL           As String
  Dim Rs            As Recordset
  Dim li            As ListItem
  Dim index         As Long
  Dim residentname  As String
  Dim Phone         As String
  Dim Rooms         As String
  Dim t             As Long

  Dim Items         As Collection
  Dim Item          As cResListItem

  Dim ResidentID    As Long

  Set Items = New Collection



  Dim CurrentPass   As Long
  Static passnumber As Long

  On Error GoTo RefreshResidents_Error

  'passnumber = passnumber + 1
  'If passnumber >= MAXLONG Then
  '  passnumber = 1
  'End If
  'CurrentPass = passnumber

  DisableButtons


  t = Win32.timeGetTime

  lvMain.ListItems.Clear
  'RefreshJet

  LockWindowUpdate lvMain.hwnd

'  SQL = " SELECT ResidentID, Name, phone, NameFirst, NameLast " & _
'        " FROM Residents " & _
'        " WHERE Residents.Deleted = 0 " & _
'        " ORDER BY NameLast, NameFirst"


  SQL = "SELECT DISTINCT  Residents.ResidentID, Residents.NameLast, Residents.NameFirst, residents.phone, Rooms.Room " & _
        "FROM Rooms RIGHT JOIN (Residents LEFT JOIN Devices ON Residents.ResidentID = Devices.ResidentID) ON Rooms.RoomID = Devices.RoomID " & _
        "WHERE Residents.deleted = 0 " & _
        "ORDER BY NameLast, NameFirst ;"



  Set Rs = ConnExecute(SQL)

  Debug.Print "Residents Execute " & Win32.timeGetTime - t
  t = Win32.timeGetTime
  Do Until Rs.EOF
    DoEvents
    
    If ResidentID <> Rs("ResidentID") Then
    
      Set Item = New cResListItem
      Item.ResidentID = Rs("ResidentID")

      Item.residentrooms = Rs("Room") & ""
      Item.NameLast = Rs("namelast") & ""
      Item.NameFirst = Rs("namefirst") & ""
      Item.Phone = Rs("phone") & ""
      Items.Add Item
      ResidentID = Rs("ResidentID")
    Else
      If Len(Rs("Room") & "") Then
        If Len(Item.residentrooms) Then
          Item.residentrooms = Item.residentrooms & "\" & Rs("Room")
        End If
      End If
    End If
    
      Rs.MoveNext
  Loop
  Debug.Print "Residents Query (includes GetResidentRooms) " & Win32.timeGetTime - t
  t = Win32.timeGetTime

  For Each Item In Items
    Set li = lvMain.ListItems.Add(, Item.ResidentKey, Item.NameFull)
    li.SubItems(1) = Item.residentrooms
    li.SubItems(2) = Item.Phone
    If ResidentID = Item.ResidentID Then
      'Index = li.Index
    End If
  Next

  Debug.Print "Residents Fill Count " & lvMain.ListItems.Count & " Time: " & Win32.timeGetTime - t
  t = Win32.timeGetTime

  If lvMain.ListItems.Count > 0 Then
  If LastIndex > 0 And index = 0 Then
    If LastIndex <= lvMain.ListItems.Count Then
      lvMain.ListItems(LastIndex).EnsureVisible
      lvMain.ListItems(LastIndex).Selected = True
    Else
      
      lvMain.ListItems(1).Selected = True
      lvMain.ListItems(1).EnsureVisible
    End If
  End If

  End If

  LockWindowUpdate 0
  EnableButtons

RefreshResidents_Resume:
  On Error Resume Next
  Rs.Close
  Set Rs = Nothing

  On Error GoTo 0
  Exit Sub

RefreshResidents_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmResidents.RefreshResidents." & Erl
  Resume RefreshResidents_Resume




End Sub

'Sub RefreshResidents2()
'
'        Dim SQL           As String
'        Dim rs            As Recordset
'        Dim li            As ListItem
'        Dim Index         As Long
'        Dim residentname  As String
'        Dim Phone         As String
'        Dim Rooms         As String
'        Dim t             As Long
'
'        Dim items         As Collection
'        Dim item          As cResListItem
'
'10      Set items = New Collection
'
'
'
'        Dim CurrentPass   As Long
'        Static passnumber As Long
'
'20      On Error GoTo RefreshResidents2_Error
'
'30      passnumber = passnumber + 1
'40      If passnumber >= MAXLONG Then
'50        passnumber = 1
'60      End If
'70      CurrentPass = passnumber
'
'80      DisableButtons
'
'
'90      t = Win32.timeGetTime
'
'100     lvMain.ListItems.Clear
'        'RefreshJet
'
'110     LockWindowUpdate lvMain.hwnd
'
'120     SQL = " SELECT ResidentID, Name, phone, NameFirst, NameLast " & _
'              " FROM Residents " & _
'              " WHERE Residents.Deleted = 0 " & _
'              " ORDER BY NameLast, NameFirst"
'
'130     Set rs = ConnExecute(SQL)
'
'140     Debug.Print "Residents Execute " & Win32.timeGetTime - t
'150     t = Win32.timeGetTime
'160     Do Until rs.EOF
'170       DoEvents
'180       Set item = New cResListItem
'190       item.ResidentID = rs("ResidentID")
'          'If ResidentID = item.ResidentID Then
'          '  index =
'
'200       item.residentrooms = GetResidentRooms(item.ResidentID)
'210       item.NameLast = rs("namelast") & ""
'220       item.NameFirst = rs("namefirst") & ""
'230       item.Phone = rs("phone") & ""
'240       items.Add item
'250       rs.MoveNext
'260     Loop
'270     Debug.Print "Residents Query (includes GetResidentRooms) " & Win32.timeGetTime - t
'280     t = Win32.timeGetTime
'
'290     For Each item In items
'300       Set li = lvMain.ListItems.Add(, item.ResidentKey, item.NameFull)
'310       li.SubItems(1) = item.residentrooms
'320       li.SubItems(2) = item.Phone
'330       If ResidentID = item.ResidentID Then
'340         Index = li.Index
'350       End If
'360     Next
'
'370     Debug.Print "Residents Fill Count " & lvMain.ListItems.Count & " Time: " & Win32.timeGetTime - t
'380     t = Win32.timeGetTime
'
'
'
'
'390     If CurrentPass = passnumber Then
'400       If Index > 0 And Index <= lvMain.ListItems.Count Then
'410         lvMain.ListItems(Index).Selected = True
'420         lvMain.ListItems(Index).EnsureVisible
'430       End If
'440     End If
'
'450     If LastIndex > 0 And Index = 0 Then
'460       If LastIndex <= lvMain.ListItems.Count Then
'470         lvMain.ListItems(LastIndex).EnsureVisible
'480         lvMain.ListItems(LastIndex).Selected = True
'490       End If
'500     End If
'
'
'510     LockWindowUpdate 0
'520     EnableButtons
'
'RefreshResidents2_Resume:
'530     On Error Resume Next
'540     rs.Close
'550     Set rs = Nothing
'
'560     On Error GoTo 0
'570     Exit Sub
'
'RefreshResidents2_Error:
'
'580     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmResidents.RefreshResidents." & Erl
'590     Resume RefreshResidents2_Resume
'
'
'End Sub

Sub Configurelvmain()
  Dim ch As ColumnHeader
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  Set ch = lvMain.ColumnHeaders.Add(, "Res", "Resident", 2500)
  Set ch = lvMain.ColumnHeaders.Add(, "Room", "Room", 2500)
  Set ch = lvMain.ColumnHeaders.Add(, "Ann", "Phone", 2300)
  lvMain.Sorted = False
  'lvMain.Sorted = True
End Sub

Private Sub cmdApply_Click()
  Apply
End Sub

Private Sub cmdDelete_Click()
  Dim ResidentID    As Long

  If Not lvMain.SelectedItem Is Nothing Then
    ResidentID = Val(lvMain.SelectedItem.Key)
    If ResidentID <> 0 Then
      If vbYes = messagebox(Me, "Delete Selected Resident?", App.Title, vbYesNo Or vbQuestion) Then
        If MASTER Then
          DeleteResident ResidentID, gUser.Username
        Else
          RemoteDeleteResident gUser, ResidentID
          RefreshJet
        End If
        RefreshResidents
      End If
    End If
  End If
  frmMain.SetListTabs
End Sub


Private Sub cmdEdit_Click()
  If lvMain.SelectedItem Is Nothing Then
    Beep
  Else
    EditResident Val(lvMain.SelectedItem.Key)
  End If
End Sub


Private Sub cmdExit_Click()
  PreviousForm
  Unload Me

End Sub

Private Sub Form_Load()
ResetActivityTime
  quitting = False
  Configurelvmain
End Sub

Private Sub Form_Unload(Cancel As Integer)
  LastIndex = 0
  
  quitting = True
  UnHost
End Sub


Public Property Get ResidentID() As Long

  ResidentID = mResidentID

End Property

Public Property Let ResidentID(ByVal ResidentID As Long)

  mResidentID = ResidentID

End Property

Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  lvMain.Sorted = True
  
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
  cmdDelete.Enabled = gUser.UserPermissions.CanDeleteResidents = 1 Or gUser.LEvel = LEVEL_FACTORY
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    LastIndex = Item.index
  End If
End Sub
