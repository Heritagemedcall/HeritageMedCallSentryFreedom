VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmOutputGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output Groups"
   ClientHeight    =   3345
   ClientLeft      =   555
   ClientTop       =   6075
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3060
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
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
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmOutputGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRoomID As Long
Private mResidentID As Long
Public Property Get RoomID() As Long
  RoomID = mRoomID
End Property

Public Property Let RoomID(ByVal RoomID As Long)
  mRoomID = RoomID
'  If mRoomID <> 0 Then
'    cmdApply.Visible = True
'  End If
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
  RefreshGroups
End Sub

Private Sub cmdAdd_Click()
  EditGroup 0
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
  cmdDelete.Enabled = True
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub

Sub RefreshGroups()
  DisableButtons
  Dim SQl   As String
  Dim rs    As Recordset
  Dim li    As ListItem
  lvMain.ListItems.Clear

  SQl = " SELECT * FROM PagerGroups "

  Set rs = ConnExecute(SQl)
  Do Until rs.EOF
    Set li = lvMain.ListItems.Add(, rs("groupID") & "B", rs("Description") & "")
    li.SubItems(1) = rs("Notes") & ""
    'Li.SubItems(2) =
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  EnableButtons
End Sub

Sub Configurelvmain()
  Dim ch As ColumnHeader
  Dim w  As Long
  w = lvMain.Width - 500
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  
  'Set ch = lvMain.ColumnHeaders.Add(, "G", "Group", 1000)
  'w = w - ch.Width
  Set ch = lvMain.ColumnHeaders.Add(, "D", "Group Description", 2500)
  w = w - ch.Width
  Set ch = lvMain.ColumnHeaders.Add(, "N", "Notes", w)
  lvMain.Sorted = True
End Sub

Private Sub cmdApply_Click()
' we dont use the apply button on this form
' Apply
End Sub

Private Sub cmdDelete_Click()
  Dim GroupID As Long
  If Not lvMain.SelectedItem Is Nothing Then
    GroupID = Val(lvMain.SelectedItem.Key)
    If GroupID <> 0 Then
      If vbYes = messagebox(Me, "Delete Output Group?", App.Title, vbYesNo Or vbQuestion) Then
        ConnExecute "DELETE FROM PagerGroups WHERE GroupID = " & GroupID
        RefreshGroups
      End If
    End If
  End If
End Sub

Private Sub cmdEdit_Click()
  If lvMain.SelectedItem Is Nothing Then
    Beep
  Else
    EditGroup Val(lvMain.SelectedItem.Key)
  End If
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub Form_Load()
  
  Configurelvmain
  ResetActivityTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
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



