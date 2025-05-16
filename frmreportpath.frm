VERSION 5.00
Begin VB.Form frmReportPath 
   Caption         =   "Report Path"
   ClientHeight    =   3525
   ClientLeft      =   285
   ClientTop       =   2325
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   10215
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
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
         Left            =   6825
         TabIndex        =   4
         Top             =   2370
         Width           =   1175
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
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
         Left            =   6825
         TabIndex        =   3
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
         Height          =   2790
         Left            =   60
         TabIndex        =   2
         Top             =   135
         Width           =   2370
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
         TabIndex        =   1
         Top             =   3015
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmReportPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LastDrive As String

Private Sub cmdApply_Click()
  If Save() Then
    PreviousForm
    Unload Me
  End If

End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
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

Public Sub Fill()

End Sub

Function FullPath() As String
  Dim path As String
  path = lstFolders.path
  If Right(path, 1) <> "\" Then
    path = path & "\"
  End If
  FullPath = path
End Function

Function Save() As Boolean
    Configuration.ReportPath = FullPath
    WriteSetting "Configuration", "ReportPath", Configuration.ReportPath
    Save = True
End Function

