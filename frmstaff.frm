VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmStaff 
   Caption         =   "Staff"
   ClientHeight    =   3015
   ClientLeft      =   5175
   ClientTop       =   5370
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   9150
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1175
      End
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
         TabIndex        =   1
         Top             =   1785
         Visible         =   0   'False
         Width           =   1175
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   6
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
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRoomID          As Long
Private mTransmitterID   As Long
Private mResidentID      As Long
Private quitting         As Boolean
Private mCaller          As String
Public IsPublic          As Long

Private LastIndex        As Long
Public Property Get TransmitterID() As Long
  TransmitterID = mTransmitterID
End Property

Public Property Let TransmitterID(ByVal TransmitterID As Long)
  cmdApply.Visible = True

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
  Dim ResidentID         As Long
  Dim SQl                As String
  Dim rs                 As ADODB.Recordset

  If lvMain.SelectedItem Is Nothing Then
    ' nada
  Else
    ResidentID = Val(lvMain.SelectedItem.Key)
    PreviousFormWithValue ResidentID
    Unload Me
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
  ShowResidents
End Sub

Private Sub cmdAdd_Click()
  EditStaff 0
End Sub

Public Sub ShowResidents()
  RefreshResidents
End Sub


Sub RefreshResidents()

  Dim SQl                As String
  Dim rs                 As Recordset
  Dim li                 As ListItem
  Dim index              As Long
  Dim residentname       As String
  Dim Phone              As String
  Dim Rooms              As String
  Dim t                  As Long

  Dim Items              As Collection
  Dim Item               As cResListItem

10 Set Items = New Collection



  Dim CurrentPass        As Long
  Static passnumber      As Long

20 On Error GoTo RefreshResidents_Error

30 passnumber = passnumber + 1
40 If passnumber >= MAXLONG Then
50  passnumber = 1
60 End If
70 CurrentPass = passnumber

80 DisableButtons


90 t = Win32.timeGetTime

100 lvMain.ListItems.Clear
110 RefreshJet

120 LockWindowUpdate lvMain.hwnd

130 SQl = " SELECT StaffID, Name, phone, NameFirst, NameLast " & _
          " FROM Staff " & _
          " WHERE Deleted = 0 " & _
          " ORDER BY NameLast, NameFirst"

140 Set rs = ConnExecute(SQl)

150 Do Until rs.EOF
160 Set Item = New cResListItem
170 Item.ResidentID = rs("StaffID")
    'If ResidentID = item.ResidentID Then
    '  index =
180 Item.residentrooms = ""    'GetResidentRooms(item.ResidentID)
190 Item.NameLast = rs("namelast") & ""
200 Item.NameFirst = rs("namefirst") & ""
210 Item.Phone = rs("phone") & ""
220 Items.Add Item
230 rs.MoveNext
240 Loop
  '  MsgBox "Query " & Win32.timeGetTime - t
250 t = Win32.timeGetTime

260 For Each Item In Items
270 Set li = lvMain.ListItems.Add(, Item.ResidentKey, Item.NameFull)

300 If ResidentID = Item.ResidentID Then
310   index = li.index
320 End If
330 Next

  'MsgBox "Fill " & Win32.timeGetTime - t
340 t = Win32.timeGetTime


350 If CurrentPass = passnumber Then
360 If index > 0 And index <= lvMain.ListItems.Count Then
370   lvMain.ListItems(index).Selected = True
380   lvMain.ListItems(index).EnsureVisible
390 End If
400 End If

410 If LastIndex > 0 And index = 0 Then
420 If LastIndex <= lvMain.ListItems.Count Then
430   lvMain.ListItems(LastIndex).EnsureVisible
440   lvMain.ListItems(LastIndex).Selected = True
450 End If
460 End If


470 LockWindowUpdate 0
480 EnableButtons

RefreshResidents_Resume:
490 On Error Resume Next
500 rs.Close
510 Set rs = Nothing

520 On Error GoTo 0
530 Exit Sub

RefreshResidents_Error:

540 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmResidents.RefreshResidents." & Erl
550 Resume RefreshResidents_Resume


End Sub

Sub Configurelvmain()
  Dim ch                 As ColumnHeader
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  Set ch = lvMain.ColumnHeaders.Add(, "Res", "Staff", 2500)
  'Set ch = lvMain.ColumnHeaders.Add(, "Room", "Room", 2500)
  'Set ch = lvMain.ColumnHeaders.Add(, "Ann", "Phone", 2300)
  lvMain.Sorted = False
  'lvMain.Sorted = True
End Sub

Private Sub cmdApply_Click()
  Apply
End Sub

Private Sub cmdDelete_Click()
  Dim ResidentID         As Long

  If Not lvMain.SelectedItem Is Nothing Then
    ResidentID = Val(lvMain.SelectedItem.Key)
    If ResidentID <> 0 Then
      If vbYes = messagebox(Me, "Delete Selected Staff?", App.Title, vbYesNo Or vbQuestion) Then
        If MASTER Then
          'DeleteResident ResidentID, gUser.username
          DeleteStaff ResidentID, gUser.UserName
          '
        Else
          'RemoteDeleteResident gUser, ResidentID
          DeleteStaff ResidentID, gUser.UserName
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
    EditStaff Val(lvMain.SelectedItem.Key)
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
  cmdApply.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
  cmdDelete.Enabled = True And gUser.LEvel >= LEVEL_ADMIN

  cmdApply.Enabled = True
  cmdExit.Enabled = True

End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    LastIndex = Item.index

  Else

  End If

End Sub



Public Property Get Caller() As String

  Caller = mCaller

End Property

Public Property Let Caller(ByVal Caller As String)

  cmdApply.Visible = (Caller <> "Announce")

  mCaller = Caller

End Property
