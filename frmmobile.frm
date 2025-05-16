VERSION 5.00
Begin VB.Form frmMobile 
   Caption         =   "Mobile"
   ClientHeight    =   3195
   ClientLeft      =   3780
   ClientTop       =   6060
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.ComboBox cboClearHistory 
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
         Left            =   2835
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2430
         Width           =   1005
      End
      Begin VB.ComboBox cboClearAssist 
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
         Left            =   5685
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2430
         Width           =   1005
      End
      Begin VB.CommandButton cmdGetPWDFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7065
         TabIndex        =   9
         Top             =   2010
         Width           =   450
      End
      Begin VB.CommandButton cmdGetEXE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7065
         TabIndex        =   6
         Top             =   1335
         Width           =   450
      End
      Begin VB.CommandButton cmdGetRoot 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7065
         TabIndex        =   3
         Top             =   675
         Width           =   450
      End
      Begin VB.TextBox txtPwdPath 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         MaxLength       =   256
         TabIndex        =   8
         Top             =   2010
         Width           =   6720
      End
      Begin VB.TextBox txtEXEPath 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         MaxLength       =   256
         TabIndex        =   5
         Top             =   1350
         Width           =   6720
      End
      Begin VB.CommandButton cmdCancel 
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
         TabIndex        =   12
         Top             =   2370
         Width           =   1175
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save"
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
         TabIndex        =   11
         Top             =   1785
         Width           =   1175
      End
      Begin VB.TextBox txtURL 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         MaxLength       =   256
         TabIndex        =   2
         Top             =   690
         Width           =   6720
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
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
         Left            =   300
         TabIndex        =   10
         Top             =   2430
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear Assistance"
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
         Left            =   4185
         TabIndex        =   17
         Top             =   2490
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear Hist"
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
         Left            =   1950
         TabIndex        =   16
         Top             =   2490
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Security Configuration"
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
         Left            =   2175
         TabIndex        =   13
         Top             =   105
         Width           =   2340
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path to Password File   (.htpasswd)"
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
         Left            =   270
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path to Password EXE"
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
         Left            =   270
         TabIndex        =   4
         Top             =   1125
         Width           =   1920
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path to Root of Web Server"
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
         Left            =   270
         TabIndex        =   1
         Top             =   465
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frmMobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Shell            As Shell32.Shell
'Private Folder As Shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1

Sub FillCombos()
  cboClearAssist.Clear
  cboClearHistory.Clear

  cboClearAssist.AddItem "30"
  cboClearHistory.AddItem "30"
  cboClearAssist.AddItem "60"
  cboClearHistory.AddItem "60"
  cboClearAssist.AddItem "90"
  cboClearHistory.AddItem "90"
  cboClearAssist.AddItem "120"
  cboClearHistory.AddItem "120"
  cboClearAssist.AddItem "150"
  cboClearHistory.AddItem "150"
  cboClearAssist.AddItem "180"
  cboClearHistory.AddItem "180"
  
  cboClearAssist.ListIndex = 0
  cboClearHistory.ListIndex = 0
  
  
  
  
  
End Sub


Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdGetEXE_Click()
  Dim Path             As String
  Path = GetFile("Choose Path to htPasswd.exe File")
  If Len(Path) Then
    Me.txtEXEPath.text = Path
  End If
  
End Sub

Private Sub cmdGetPWDFile_Click()
  Dim Folder             As String
  Folder = BrowseForFolder("Choose Path to Password File")
  If Len(Folder) Then
    Me.txtPwdPath.text = Folder
  End If

End Sub

Private Sub cmdOK_Click()
  ResetActivityTime
  Save

End Sub
Public Sub Host(ByVal hwnd As Long)
  fraEnabler.BackColor = Me.BackColor
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Sub Fill()
  Dim MobileEnabled      As Long

  MobileEnabled = Configuration.MobileWebEnabled
  txtURL.text = Configuration.MobileWebRoot
  txtPwdPath.text = Configuration.MobilehtPasswordPath
  txtEXEPath.text = Configuration.MobilehtPasswordEXEPath
  
  If Len(Configuration.MobilehtPasswordPath) = 0 Then
    MobileEnabled = 0
  End If
  
  If Len(Configuration.MobileWebRoot) = 0 Then
    MobileEnabled = 0
  End If
  
  If Len(Configuration.MobilehtPasswordEXEPath) = 0 Then
    MobileEnabled = 0
  End If
  
  chkEnabled.Value = IIf(MobileEnabled = 1, 1, 0)
  Dim j As Long
  For j = 0 To cboClearAssist.listcount - 1
    If cboClearAssist.list(j) = CStr(Configuration.MobileClearAssist) Then
      cboClearAssist.ListIndex = j
      Exit For
    End If
  Next

  For j = 0 To Me.cboClearHistory.listcount - 1
    If cboClearHistory.list(j) = CStr(Configuration.MobileClearHistory) Then
      cboClearHistory.ListIndex = j
      Exit For
    End If
  Next



End Sub
Sub Save()
  Dim MobileEnabled      As Long
  Dim Root               As String
  Dim PasswordEXEPath    As String
  Dim PasswordPath       As String

  MobileEnabled = chkEnabled.Value
  Root = Trim$(txtURL.text)
  PasswordEXEPath = Trim$(Me.txtEXEPath.text)
  PasswordPath = Trim$(Me.txtPwdPath.text)

  If Len(Root) = 0 Then
    MobileEnabled = 0
  End If

  If Len(PasswordEXEPath) = 0 Then
    MobileEnabled = 0
  End If

  If Len(PasswordPath) = 0 Then
    MobileEnabled = 0
  End If

  WriteSetting "Mobile", "Root", Root
  WriteSetting "Mobile", "Enabled", MobileEnabled And 1
  WriteSetting "Mobile", "PasswordPath", PasswordPath
  WriteSetting "Mobile", "PasswordEXEPath", PasswordEXEPath

  WriteSetting "Mobile", "ClearAssist", cboClearAssist.text
  WriteSetting "Mobile", "ClearHist", cboClearHistory.text
  

  Configuration.MobilehtPasswordPath = ReadSetting("Mobile", "PasswordPath", "")
  Configuration.MobilehtPasswordEXEPath = ReadSetting("Mobile", "PasswordEXEPath", "")
  Configuration.MobileWebRoot = ReadSetting("Mobile", "Root", "")
  
  Configuration.MobileClearAssist = Val(ReadSetting("Mobile", "ClearAssist", "60"))
  Configuration.MobileClearHistory = Val(ReadSetting("Mobile", "ClearHist", "60"))
  
  
  
  MobileEnabled = Val(ReadSetting("Mobile", "Enabled", "0"))

  If Len(Root) = 0 Then
    MobileEnabled = 0
  End If

  If Len(PasswordEXEPath) = 0 Then
    MobileEnabled = 0
  End If

  If Len(PasswordPath) = 0 Then
    MobileEnabled = 0
  End If
  Configuration.MobileWebEnabled = MobileEnabled And 1

  If Configuration.MobileWebEnabled = 1 Then
    SyncApacheUsers
    WriteApacheHint
  End If


  Fill

End Sub

Sub WriteApacheHint()
  Dim hfile As Long
  On Error Resume Next
  If DirExists(Configuration.MobileWebRoot) Then
    DeleteFile (Configuration.MobileWebRoot & "\.htaccess2")
    hfile = FreeFile
    Open Configuration.MobileWebRoot & "\.htaccess2" For Output As hfile
    Print #hfile, "# Add/Change These Lines in at the Beginning of .htaccess"
    Print #hfile, "AuthName " & """Heritage Medcall"""
    Print #hfile, "AuthType Basic"
    Print #hfile, "AuthUserFile " & Configuration.MobilehtPasswordPath & "\.htpasswd"
    Print #hfile, "Require valid-user"
    Close hfile
'AuthName "Heritage Medcall"
'AuthType Basic
'AuthUserFile D:\HeritageMedCall\Freedom2\.htpasswd
'Require valid - User

    
 
  End If


End Sub

Private Sub cmdGetRoot_Click()
  Dim Folder             As String
  Folder = BrowseForFolder("Choose Web Root")
  If Len(Folder) Then
    Me.txtURL.text = Folder
  End If
End Sub
Private Function BrowseForFolder(Optional ByVal Title As String = "Choose a Folder:")
  Dim Shell As Shell32.Shell
  Dim Folder As Shell32.Folder
  
  Set Shell = New Shell32.Shell
  Set Folder = Shell.BrowseForFolder(Me.hwnd, Title, BIF_RETURNONLYFSDIRS)
  If Not Folder Is Nothing Then
    BrowseForFolder = Folder.Self.Path
  End If
End Function

Function GetFile(Optional ByVal Title As String = "Select a File") As String
  ' non - blocking
  Dim tpOpenFname As OPENFILENAME

  tpOpenFname.lpstrFile = String(256, 0)
  tpOpenFname.nMaxFile = 255
  
  tpOpenFname.lpstrDefExt = "exe"
  tpOpenFname.lpstrFilter = "*.exe"
  tpOpenFname.lpstrTitle = Title
  tpOpenFname.hwndOwner = Me.hwnd
  
  tpOpenFname.lStructSize = Len(tpOpenFname)
  

  If GetOpenFileName(tpOpenFname) <> 0 Then
    GetFile = left$(tpOpenFname.lpstrFile, tpOpenFname.nMaxFile)
  Else
    'Debug.Print "Open Canceled"
  End If

End Function

Private Sub Form_Load()
  FillCombos
End Sub

Private Sub txtEXEPath_GotFocus()
  SelAll txtEXEPath
End Sub

Private Sub txtPwdPath_GotFocus()
  SelAll txtPwdPath
End Sub

Private Sub txtURL_GotFocus()
  SelAll txtURL
End Sub

