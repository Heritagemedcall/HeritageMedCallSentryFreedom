VERSION 5.00
Begin VB.Form frmExternalApp 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   1245
   ClientTop       =   5925
   ClientWidth     =   9225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   9225
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9150
      Begin VB.TextBox txtFileName 
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
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   8535
      End
      Begin VB.ComboBox cboFileType 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3090
         Width           =   1935
      End
      Begin VB.FileListBox flist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   180
         TabIndex        =   3
         Top             =   690
         Width           =   3675
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
         Left            =   7725
         TabIndex        =   10
         Top             =   2400
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
         TabIndex        =   9
         Top             =   1785
         Width           =   1175
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
         Height          =   2340
         Left            =   3930
         TabIndex        =   4
         Top             =   705
         Width           =   3675
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
         Left            =   5100
         TabIndex        =   8
         Top             =   3090
         Width           =   2490
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
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
         Left            =   240
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lblDrive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive"
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
         Left            =   4260
         TabIndex        =   7
         Top             =   3150
         Width           =   705
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Type"
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
         Left            =   390
         TabIndex        =   5
         Top             =   3120
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmExternalApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As String
Dim LastGoodDrive As String

Public Sub Fill()

End Sub

Private Sub cboFileType_Click()
  On Error Resume Next
  flist.Pattern = "*." & cboFileType.text
  
End Sub

Private Sub cmdApply_Click()
  ResetActivityTime
  If Save() Then
    PreviousForm
    Unload Me
  End If


End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Sub FillCombo()
  cboFileType.Clear
  cboFileType.AddItem "exe"
  cboFileType.AddItem "bat"
  cboFileType.AddItem "lnk"
  cboFileType.AddItem "*"
  cboFileType.ListIndex = 0
End Sub

Private Sub flist_Click()
  Dim newpath As String
  newpath = lstFolders.path
  If Right(newpath, 1) <> "\" Then
    newpath = newpath & "\"
  End If
  txtFileName.text = newpath & flist.filename
End Sub

Private Sub Form_Activate()
  Fill
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
  fraEnabler.BackColor = Me.BackColor
  FillCombo
  ArrangeControls
End Sub
Private Sub ArrangeControls()

End Sub
Private Sub lstDrives_Change()
  On Error Resume Next
  lstFolders.path = lstDrives.Drive
  If Err.Number = 0 Then
    LastGoodDrive = lstDrives.Drive
  Else
    If LastGoodDrive = "" Then
      LastGoodDrive = App.path
    End If
    lstDrives.Drive = LastGoodDrive
  End If
End Sub

Private Sub lstFolders_Change()
  On Error Resume Next
  flist.path = lstFolders.path
End Sub

Private Function Save() As Boolean
  Save = True
  Select Case Val(Source)
    Case 1
      WriteSetting "ExtApps", "App1", Trim(txtFileName.text)
    Case 2
      WriteSetting "ExtApps", "App2", Trim(txtFileName.text)
    Case 3
      WriteSetting "ExtApps", "App3", Trim(txtFileName.text)
    Case Else
      Save = False
  End Select
End Function
Public Sub Host(ByVal hwnd As Long)
  
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  SetParent fraEnabler.hwnd, hwnd
  fraEnabler.BackColor = Me.BackColor
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
  'Caller = ""
End Sub

