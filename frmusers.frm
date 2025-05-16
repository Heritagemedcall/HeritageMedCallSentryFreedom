VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmUsers 
   Caption         =   "Users"
   ClientHeight    =   3225
   ClientLeft      =   165
   ClientTop       =   2340
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   9600
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
         TabIndex        =   6
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
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Fill()
  ResetActivityTime
    
  DisableButtons
  Dim SQL   As String
  Dim Rs    As Recordset
  Dim li    As ListItem
  
  lvMain.ListItems.Clear
  
  SQL = "SELECT * FROM users"
  Set Rs = ConnExecute(SQL)
  
  Do Until Rs.EOF
    Set li = lvMain.ListItems.Add(, Rs("userID") & "U", Rs("username") & "")
    li.SubItems(1) = GetLevelString(Rs("level"))
    Rs.MoveNext

  Loop
  Rs.Close
  Set Rs = Nothing
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
  cmdDelete.Enabled = True And gUser.LEvel >= LEVEL_ADMIN
  
  
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub

Private Sub cmdAdd_Click()
  EditUser 0
End Sub

Private Sub cmdDelete_Click()
  DeleteUser
  SyncApacheUsers
End Sub
Sub DeleteUser()
  Dim UserID As Long
  If Not lvMain.SelectedItem Is Nothing Then
    UserID = Val(lvMain.SelectedItem.Key)
    ConnExecute "Delete from users where userid = " & UserID
    Fill
  End If
  
End Sub
Private Sub Form_Load()
ResetActivityTime
  Configurelvmain
End Sub

Private Sub Form_Unload(Cancel As Integer)

  UnHost
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

Sub Configurelvmain()
  Dim ch As ColumnHeader
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  Set ch = lvMain.ColumnHeaders.Add(, "User", "User", 2500)
  Set ch = lvMain.ColumnHeaders.Add(, "Level", "Level")
  lvMain.Sorted = True
End Sub
Private Sub cmdEdit_Click()
  If lvMain.SelectedItem Is Nothing Then
    Beep
  Else
    EditUser Val(lvMain.SelectedItem.Key)
  End If
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me

End Sub

