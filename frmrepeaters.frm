VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmRepeaters 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   960
   ClientTop       =   5835
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   9660
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
         Caption         =   "Refresh"
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
Attribute VB_Name = "frmRepeaters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OnlyPCAs     As Boolean
Public Serial       As String
Public returnobject As Object



Private mRoomID As Long
Private mResidentID As Long
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
  ShowRepeaters

End Sub

Private Sub cmdAdd_Click()
  EditTransmitter 0
End Sub

Public Sub ShowRepeaters()
  DisableButtons
  RefreshDevices
  EnableButtons
End Sub


Sub Configurelvmain()
  Dim ch As ColumnHeader

  If lvMain.ColumnHeaders.Count < 8 Then


    lvMain.ColumnHeaders.Clear
    lvMain.Sorted = True
    Set ch = lvMain.ColumnHeaders.Add(, "Serial", "Serial", 1100)
    Set ch = lvMain.ColumnHeaders.Add(, "Model", "Model", 900)
    Set ch = lvMain.ColumnHeaders.Add(, "LastSeen", "LastSeen", 1300)
    Set ch = lvMain.ColumnHeaders.Add(, "FirstHop", "FirstHop", 1100)
    Set ch = lvMain.ColumnHeaders.Add(, "Level", "L", 400)
    Set ch = lvMain.ColumnHeaders.Add(, "Margin", "M", 400)
    Set ch = lvMain.ColumnHeaders.Add(, "LastConfig", "Last Config", 1300)
    Set ch = lvMain.ColumnHeaders.Add(, "NID", "NID", 600)
    Set ch = lvMain.ColumnHeaders.Add(, "Ly", "Ly", 600)
    Set ch = lvMain.ColumnHeaders.Add(, "Jam", "Jam", 1100)
  End If
End Sub

Sub RefreshDevices()

  On Error GoTo RefreshDevices_Error

  DisableButtons

  

  Dim SQl        As String
  Dim rs         As Recordset
  Dim li         As ListItem
  Dim j          As Integer
  Dim d          As cESDevice
  Dim assure     As String
  Dim index      As Long
  Dim rsRoom     As Recordset

  Dim k As Long

  Dim SortedDevices As Collection

  Dim CurrentPass As Long
  Static passnumber As Long

  passnumber = passnumber + 1
  If passnumber >= MAXLONG Then
    passnumber = 1
  End If
  CurrentPass = passnumber

  Set SortedDevices = New Collection

  lvMain.ListItems.Clear
  LockWindowUpdate lvMain.hwnd
  lvMain.Sorted = True
  lvMain.SortKey = 0
  lvMain.SortOrder = lvwAscending

  SQl = " SELECT Serial,deviceid,ResidentID,RoomID,Model,UseAssur,UseAssur2,Assurinput FROM Devices " & _
      " WHERE (devices.model = 'EN5000'  or devices.model = 'EN5040' or  devices.model = 'EN5081' )" & _
      " ORDER BY Devices.Serial "


  RefreshJet
  Set rs = ConnExecute(SQl)


  Do Until rs.EOF
    k = k + 1
    If CurrentPass <> passnumber Then Exit Do
    DoEvents
    If quitting Then Exit Do
    ' REMOTE todo

    Set li = lvMain.ListItems.Add(, rs("DeviceID") & "B", Right("00000000" & rs("serial"), 8))
    li.SubItems(1) = rs("Model") & ""
    Set d = Devices.device(rs("serial") & "")
    If Not d Is Nothing Then
      If d.LastSeen <= 0 Then
        li.SubItems(2) = " -- "
      Else
        li.SubItems(2) = Format(d.LastSeen, "mm-dd hh:nn")
      End If
      li.SubItems(3) = d.FirstHop
      li.SubItems(4) = d.LastLevel
      li.SubItems(5) = d.LastMargin
      If d.LastConfigResponse <= 0 Then
        li.SubItems(6) = " -- "
      Else
        li.SubItems(6) = Format(d.LastConfigResponse, "mm-dd hh:nn")
      End If
      If d.NID < 0 Then
        li.SubItems(7) = " -- "
      Else
        li.SubItems(7) = d.NID
      End If
      If d.Layer < 0 Then
        li.SubItems(8) = " -- "
      Else
        li.SubItems(8) = d.Layer
      End If
      li.SubItems(9) = d.JamCount
      
      
    End If
    rs.MoveNext
  Loop

  If LastIndex > 0 Then
    If LastIndex <= lvMain.ListItems.Count Then
      lvMain.ListItems(LastIndex).EnsureVisible
      lvMain.ListItems(LastIndex).Selected = True
    End If
  End If
 
  LockWindowUpdate 0
  
  
  
  EnableButtons



RefreshDevices_Resume:
  On Error Resume Next
  rs.Close
  Set rs = Nothing

  On Error GoTo 0
  Exit Sub

RefreshDevices_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitters.RefreshDevices." & Erl
  Resume RefreshDevices_Resume

End Sub


Private Sub cmdApply_Click()
  Fill
End Sub

Private Sub cmdDelete_Click()



  Dim i As Long
  Dim Key As Long

  If Not lvMain.SelectedItem Is Nothing Then
    If vbYes = messagebox(Me, "Delete Selected Device?", App.Title, vbQuestion Or vbYesNo) Then
      i = lvMain.SelectedItem.index
      Key = Val(lvMain.SelectedItem.Key)
      If MASTER Then
        DeleteTransmitter Key, gUser.UserName
      Else
        RemoteDeleteTransmitter Key
        RefreshJet
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
ResetActivityTime
  quitting = False
  Configurelvmain
End Sub

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
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True And gUser.LEvel >= LEVEL_ADMIN
  
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
'ResetRemoteRefreshCounter
  If Not Item Is Nothing Then
    LastIndex = Item.index
  End If
  
End Sub

