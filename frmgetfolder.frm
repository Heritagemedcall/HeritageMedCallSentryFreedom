VERSION 5.00
Begin VB.Form frmGetFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Folder"
   ClientHeight    =   3570
   ClientLeft      =   375
   ClientTop       =   5670
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   9360
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9150
      Begin VB.TextBox txtNewFolder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4170
         TabIndex        =   6
         Top             =   180
         Width           =   4275
      End
      Begin VB.CommandButton cmdNewFolder 
         Caption         =   "New Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   600
         Width           =   1500
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
         Top             =   2730
         Width           =   3960
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
         Width           =   3960
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
         TabIndex        =   3
         Top             =   1785
         Width           =   1175
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
         TabIndex        =   4
         Top             =   2370
         Width           =   1175
      End
   End
End
Attribute VB_Name = "frmGetFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Caller As String

Private mFolder As String
Private mFolderRemote As Boolean ' flag if this is for remote FTP backup
Private LastDrive As String


Private Sub cmdApply_Click()
  ResetActivityTime
  Save
  PreviousForm
  Unload Me
End Sub

Function Save() As Boolean
  mFolder = lstFolders.path

  If (mFolderRemote) Then
    WriteSetting "RemoteBackup", "Folder", mFolder
    Configuration.BackupFolderRemote = mFolder
  Else
    WriteSetting "Backup", "Folder", mFolder
    Configuration.BackupFolder = mFolder

  End If

End Function

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub



Private Sub cmdNewFolder_Click()
  
  Dim folder As String
  ResetActivityTime
  
  folder = Trim$(txtNewFolder.text)
  If Len(folder) > 0 Then
    On Error Resume Next
    MkDir lstFolders.path & "\" & folder
    lstFolders.Refresh
  End If
  
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
  LastDrive = lstDrives.Drive
  
End Sub

Public Property Let FolderRemote(ByVal Value As Boolean)
  mFolderRemote = Value
End Property

Public Property Get folder() As String
  folder = mFolder
End Property


Public Property Let folder(ByVal folder As String)
  mFolder = folder
End Property







Public Sub Fill()
  On Error Resume Next
  lstFolders.path = folder
  lstDrives.Drive = lstFolders.path
  

End Sub
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
End Sub
  
  
Private Sub lstDrives_Change()
  Dim Retry As Boolean
  On Error Resume Next

  Retry = True
  Do While Retry
    Retry = False
    lstFolders.path = lstDrives.Drive
    ' If and error occurs
    Select Case Err.Number
      Case 68  ' Not accessable Error
        If vbRetry = messagebox(Me, lstDrives.Drive & " is not accessible", App.Title, vbRetryCancel + vbCritical) Then
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

Private Sub lstFolders_Change()
  Static Busy As Boolean

  On Error Resume Next

  If Not Busy Then
  
    Busy = True
    ' Change the current directory
    ChDir lstFolders.path

    If Err.Number = 0 Then
      lstDrives.Drive = left(lstFolders.path, 2)
    Else
      Err.Clear
    End If

    Busy = False

  End If

End Sub

Private Sub lstFolders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
End Sub

Private Sub lstFolders_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
End Sub
