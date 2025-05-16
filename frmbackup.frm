VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Settngs"
   ClientHeight    =   11970
   ClientLeft      =   915
   ClientTop       =   2550
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11970
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9195
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9285
      Begin VB.Frame fraFTPsettings 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   2565
         Left            =   90
         TabIndex        =   35
         Top             =   6270
         Width           =   8925
         Begin VB.CommandButton cmdSave3 
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
            Left            =   7620
            TabIndex        =   45
            Top             =   1200
            Width           =   1175
         End
         Begin VB.CommandButton cmdExit3 
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
            Left            =   7620
            TabIndex        =   46
            Top             =   1830
            Width           =   1175
         End
         Begin VB.TextBox txtHost 
            Height          =   345
            Left            =   330
            TabIndex        =   37
            Top             =   465
            Width           =   3795
         End
         Begin VB.TextBox txtUser 
            Height          =   345
            Left            =   330
            TabIndex        =   39
            Top             =   1125
            Width           =   3795
         End
         Begin VB.TextBox txtPass 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   330
            PasswordChar    =   "*"
            TabIndex        =   41
            Top             =   1785
            Width           =   3795
         End
         Begin VB.OptionButton optPassive 
            Caption         =   "Passive Connection"
            Height          =   255
            Left            =   4440
            TabIndex        =   43
            Top             =   720
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Active Connection"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4440
            TabIndex        =   42
            Top             =   420
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "Test Connection"
            Height          =   495
            Left            =   4950
            TabIndex        =   44
            Top             =   1170
            Width           =   1150
         End
         Begin VB.Label lblLastFTPMessage 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   450
            TabIndex        =   48
            Top             =   2250
            Width           =   75
         End
         Begin VB.Label lblhost 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Host:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   36
            Top             =   210
            Width           =   375
         End
         Begin VB.Label lbluser 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   38
            Top             =   870
            Width           =   375
         End
         Begin VB.Label lblPass 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   40
            Top             =   1530
            Width           =   735
         End
      End
      Begin VB.Frame fraFTP 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Rooms"
         Height          =   2820
         Left            =   60
         TabIndex        =   21
         Top             =   3270
         Width           =   9000
         Begin VB.CommandButton cmdRemoteBackupNow 
            Caption         =   "Backup Now"
            Height          =   585
            Left            =   5430
            TabIndex        =   47
            Top             =   1380
            Width           =   1150
         End
         Begin VB.CommandButton cmdExit2 
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
            TabIndex        =   34
            Top             =   2085
            Width           =   1175
         End
         Begin VB.CommandButton cmdApply2 
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
            TabIndex        =   33
            Top             =   1455
            Width           =   1175
         End
         Begin VB.CommandButton cmdGetfolderRemote 
            Height          =   330
            Left            =   6960
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmBackup.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2340
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtFolderRemote 
            Height          =   375
            Left            =   420
            TabIndex        =   31
            Top             =   2310
            Width           =   6495
         End
         Begin VB.ComboBox cboTimeRemote 
            Height          =   315
            ItemData        =   "frmBackup.frx":052A
            Left            =   360
            List            =   "frmBackup.frx":052C
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   960
            Width           =   1965
         End
         Begin VB.OptionButton optMonthlyRemote 
            Caption         =   "Monthly"
            Height          =   300
            Left            =   420
            TabIndex        =   27
            Top             =   1695
            Width           =   1335
         End
         Begin VB.OptionButton optDailyRemote 
            Caption         =   "Days (Every Week)"
            Height          =   300
            Left            =   420
            TabIndex        =   26
            Top             =   1365
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.ListBox lstDOMRemote 
            Height          =   1605
            IntegralHeight  =   0   'False
            ItemData        =   "frmBackup.frx":052E
            Left            =   2940
            List            =   "frmBackup.frx":0535
            TabIndex        =   29
            Top             =   420
            Width           =   2235
         End
         Begin VB.CheckBox chkEnableRemote 
            Caption         =   "Remote Backups"
            Height          =   315
            Left            =   360
            TabIndex        =   22
            Top             =   270
            Width           =   1815
         End
         Begin VB.ListBox lstDOWRemote 
            Height          =   1635
            Left            =   2910
            Style           =   1  'Checkbox
            TabIndex        =   28
            Top             =   390
            Width           =   2235
         End
         Begin VB.Label lblFolderRemote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder"
            Height          =   195
            Left            =   480
            TabIndex        =   30
            Top             =   2070
            Width           =   540
         End
         Begin VB.Label lblHeaderremote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Times"
            Height          =   195
            Left            =   2880
            TabIndex        =   25
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblrtime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            Height          =   195
            Left            =   330
            TabIndex        =   23
            Top             =   735
            Width           =   420
         End
      End
      Begin VB.Frame fraLocal 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Rooms"
         Height          =   2820
         Left            =   30
         TabIndex        =   2
         Top             =   330
         Width           =   9030
         Begin VB.CommandButton cmdExportDevices 
            Caption         =   "Export Devices"
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
            Left            =   5550
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1020
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
            TabIndex        =   20
            Top             =   2070
            Width           =   1175
         End
         Begin VB.CommandButton cmdApply 
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
            TabIndex        =   19
            Top             =   1455
            Width           =   1175
         End
         Begin VB.ListBox lstDOM 
            Height          =   1605
            IntegralHeight  =   0   'False
            Left            =   2940
            TabIndex        =   10
            Top             =   450
            Width           =   2175
         End
         Begin VB.CheckBox chkEnable 
            Caption         =   "Local Backups"
            Height          =   315
            Left            =   360
            TabIndex        =   3
            Top             =   270
            Width           =   1815
         End
         Begin VB.ListBox lstDOW 
            Height          =   1635
            ItemData        =   "frmBackup.frx":0547
            Left            =   2910
            List            =   "frmBackup.frx":054E
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   420
            Width           =   2235
         End
         Begin VB.OptionButton OptDaily 
            Caption         =   "Days (Every Week)"
            Height          =   300
            Left            =   420
            TabIndex        =   6
            Top             =   1365
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.OptionButton optMonthly 
            Caption         =   "Monthly"
            Height          =   300
            Left            =   420
            TabIndex        =   7
            Top             =   1695
            Width           =   1335
         End
         Begin VB.ComboBox cboTime 
            Height          =   315
            ItemData        =   "frmBackup.frx":055A
            Left            =   360
            List            =   "frmBackup.frx":055C
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   1965
         End
         Begin VB.TextBox txtFolder 
            Height          =   375
            Left            =   420
            TabIndex        =   17
            Top             =   2310
            Width           =   6495
         End
         Begin VB.CommandButton cmdGetfolder 
            Height          =   330
            Left            =   6960
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmBackup.frx":055E
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2340
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdDoBackup 
            Caption         =   "Backup Now"
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
            Left            =   5550
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1175
         End
         Begin VB.CommandButton cmdPurgeData 
            Caption         =   "Purge Data"
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
            Left            =   5520
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   30
            Width           =   1175
         End
         Begin VB.TextBox txtAge 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5520
            MaxLength       =   3
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "365"
            Top             =   660
            Width           =   525
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            Height          =   195
            Left            =   330
            TabIndex        =   4
            Top             =   735
            Width           =   420
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Times"
            Height          =   195
            Left            =   2880
            TabIndex        =   8
            Top             =   120
            Width           =   510
         End
         Begin VB.Label blbFolder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder"
            Height          =   195
            Left            =   480
            TabIndex        =   16
            Top             =   2070
            Width           =   540
         End
         Begin VB.Label lblAge 
            BackStyle       =   0  'Transparent
            Caption         =   "Days Old"
            Height          =   225
            Left            =   6240
            TabIndex        =   13
            Top             =   720
            Width           =   915
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   3195
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5636
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Local Backups"
               Key             =   "local"
               Object.ToolTipText     =   "Configure Local Backups"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Remote Backups"
               Key             =   "remote"
               Object.ToolTipText     =   "Configure Remote Backups"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "FTP Settings"
               Key             =   "ftp"
               Object.ToolTipText     =   "FTP Configuration"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  
  Save

End Sub

Private Sub cmdApply2_Click()
  Save
End Sub

Private Sub cmdDoBackup_Click()
  ResetActivityTime
  DoBackup
End Sub

Private Sub cmdExit_Click()

  PreviousForm
  Unload Me
End Sub

Public Sub Fill()

  Dim j As Integer
  Dim DAYS() As String
  Dim DaysRemote() As String
  
  ShowOptions
  'txtFolder.Text = Configuration.BackupFolder

  'local
  For j = cboTime.listcount - 1 To 0 Step -1
    If Configuration.BackupTime = cboTime.ItemData(j) Then

      cboTime.ListIndex = j
      Exit For
    End If
  Next

  If j = -1 Then
    cboTime.ListIndex = 0
  End If

  ' remote
  For j = cboTimeRemote.listcount - 1 To 0 Step -1
    If Configuration.BackupTimeRemote = cboTimeRemote.ItemData(j) Then

      cboTimeRemote.ListIndex = j
      Exit For
    End If
  Next

  If j = -1 Then
    cboTimeRemote.ListIndex = 0
  End If

  ' local
  For j = 0 To 6
    If 2 ^ j And Configuration.BackupDOW Then
      lstDOW.Selected(j) = True
    Else
      lstDOW.Selected(j) = False
    End If
  Next

  'remote
  For j = 0 To 6
    If 2 ^ j And Configuration.BackupDOWRemote Then
      lstDOWRemote.Selected(j) = True
    Else
      lstDOWRemote.Selected(j) = False
    End If
  Next
  
  
  
  'local
  

  For j = lstDOM.listcount - 1 To 0 Step -1
    If j = Configuration.BackupDOM - 1 Then
      lstDOM.Selected(j) = True
      Exit For
    End If
  Next
  If j < 0 Then
    lstDOM.Selected(0) = True
    lstDOM.ListIndex = 0
  End If


  ' remote
  

  For j = lstDOMRemote.listcount - 1 To 0 Step -1
    If j = Configuration.BackupDOMRemote - 1 Then
      lstDOMRemote.Selected(j) = True
      Exit For
    End If
  Next
  If j < 0 Then
    lstDOMRemote.Selected(0) = True
    lstDOMRemote.ListIndex = 0
  End If

'
'  For j = 0 To lstDOMRemote.ListCount - 1
'      lstDOMRemote.Selected(j) = False
'  Next
'
'
'  For j = LBound(DaysRemote) To UBound(DaysRemote)
'    If j >= 28 Then Exit For
'    If DaysRemote(j) = "1" Then
'      lstDOMRemote.Selected(j) = True
'    End If
'  Next

  
'  For j = lstDOMRemote.ListCount - 1 To 0 Step -1
'    If lstDOMRemote.ItemData(j) = Configuration.BackupDOMRemote Then
'      Exit For
'    End If
'  Next
'
'  If j = -1 Then
'    j = 0
'  End If
'  lstDOMRemote.ListIndex = j



  chkEnable.Value = IIf(Configuration.BackupEnabled = 1, 1, 0)
  txtFolder.text = Configuration.BackupFolder
  If Configuration.BackupType = 1 Then
    optMonthly.Value = True
  Else
    OptDaily.Value = True
  End If



  chkEnableRemote.Value = IIf(Configuration.BackupEnabledRemote = 1, 1, 0)
  txtFolderRemote.text = Configuration.BackupFolderRemote
  If Configuration.BackupTypeRemote = 1 Then
    optMonthlyRemote.Value = True
  Else
    optDailyRemote.Value = True
  End If
  
  txtHost.text = Configuration.BackupHost
  txtUser.text = Configuration.BackupUser
  txtPass.text = Configuration.BackupPassword
  

  updatescreen

End Sub

Function ValidateBackupFolder(ByVal folder As String) As Boolean
  On Error Resume Next
  Dim s As String
  s = Dir(folder, vbDirectory)
  If Err.Number = 0 Then
    ValidateBackupFolder = (Len(s) > 0)
  End If
End Function
Sub updatescreen()
  If Me.OptDaily.Value Then
    lstDOW.Visible = True
    lstDOM.Visible = False
    lblHeader.Caption = "Days"
  Else
    lstDOM.Visible = True
    lstDOW.Visible = False
    lblHeader.Caption = "Day of Month"
  End If

  If Me.optDailyRemote.Value Then
    lstDOWRemote.Visible = True
    lstDOMRemote.Visible = False
    lblHeaderremote.Caption = "Days"
  Else
    lstDOMRemote.Visible = True
    lstDOWRemote.Visible = False
    lblHeaderremote.Caption = "Day of Month"
  End If



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

Private Sub cmdExit2_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdExit3_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdExportDevices_Click()
  Dim path    As String
  Dim hfile   As Integer
  Dim d As cESDevice

  path = App.path
  If Right(path, 1) <> "\" Then
    path = path & "\"
  End If


  hfile = FreeFile


  Open path & "deviceexport.txt" For Output As #hfile
  Print #hfile, "Serial" & vbTab & "Model" & vbTab & "Room" & vbCrLf;


  For Each d In Devices.Devices
    Print #hfile, d.Serial & vbTab & d.Model & vbTab & d.Room & vbCrLf;
  Next



  Close hfile


End Sub

Private Sub cmdGetfolder_Click()
  If Save() Then
    GetFolder txtFolder.text, Me.name
  End If
End Sub

Private Sub cmdGetfolderRemote_Click()
  If Save() Then
    GetFolderremote txtFolderRemote.text, Me.name
  End If
End Sub



Private Sub cmdPurgeData_Click()
  Dim Predate As Date
  Predate = DateAdd("d", Now, -Val(txtAge.text))


  If vbYes = MsgBox("This will delete all Alarm data Prior to " & vbCrLf & Format(Predate, "mm/dd/yyyy") & vbCrLf & "Are you sure?", vbYesNo Or vbDefaultButton2, "Purge Old Data") Then
    Dim SQl As String
    
    Dim cutoff As String
    cutoff = Format(Predate, "mm/dd/yyyy")
    
    SQl = "Delete from Alarms WHERE EventDate < " & DateDelim & cutoff & DateDelim
    ConnExecute SQl


  End If
End Sub

Private Sub cmdRemoteBackupNow_Click()
  ResetActivityTime
  DoBackupRemote
End Sub

Private Sub cmdSave3_Click()
 Save
End Sub

Private Sub cmdTest_Click()
  Dim rc As Long
  'On Error Resume Next
  ResetActivityTime
  lblLastFTPMessage.Caption = ""
  rc = TestFTPConnection(txtHost.text, txtUser.text, txtPass.text)
  If (rc) Then
    lblLastFTPMessage.Caption = "Success"
  Else
    lblLastFTPMessage.Caption = GetLastFTPError()
  End If
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
    ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
  ArrangeControls

  ShowOptions
  FillLists
End Sub
Sub ArrangeControls()
  fraEnabler.BackColor = Me.BackColor

  fraLocal.left = TabStrip.ClientLeft
  fraLocal.top = TabStrip.ClientTop
  fraLocal.Height = TabStrip.ClientHeight
  fraLocal.Width = TabStrip.ClientWidth
  fraLocal.BackColor = Me.BackColor

  fraFTP.left = TabStrip.ClientLeft
  fraFTP.top = TabStrip.ClientTop
  fraFTP.Height = TabStrip.ClientHeight
  fraFTP.Width = TabStrip.ClientWidth
  fraFTP.BackColor = Me.BackColor


  fraFTPsettings.left = TabStrip.ClientLeft
  fraFTPsettings.top = TabStrip.ClientTop
  fraFTPsettings.Height = TabStrip.ClientHeight
  fraFTPsettings.Width = TabStrip.ClientWidth
  fraFTPsettings.BackColor = Me.BackColor
  
  SetTabs

End Sub
Sub ShowOptions()
  cmdPurgeData.Visible = False
  txtAge.Visible = False
  lblAge.Visible = False

  If MASTER And (gUser.LEvel >= LEVEL_ADMIN) Then
    cmdPurgeData.Visible = True
    txtAge.Visible = True
    lblAge.Visible = True
  End If

End Sub

Private Sub FillLists()

  Dim j As Integer


  cboTime.Clear
  cboTimeRemote.Clear
  For j = 1 To 24
    cboTime.AddItem Format(j, "00") & ":00" & IIf(j = 12, " (noon)", IIf(j = 24, " (midnight)", ""))
    cboTime.ItemData(cboTime.NewIndex) = j * 100
    cboTimeRemote.AddItem Format(j, "00") & ":00" & IIf(j = 12, " (noon)", IIf(j = 24, " (midnight)", ""))
    cboTimeRemote.ItemData(cboTime.NewIndex) = j * 100

  Next
  cboTime.ListIndex = 0
  cboTimeRemote.ListIndex = 0

  lstDOW.Clear
  lstDOWRemote.Clear
  For j = 1 To 7
    lstDOW.AddItem Format(j, "dddd")
    lstDOW.ItemData(lstDOW.NewIndex) = j - 1
    lstDOWRemote.AddItem Format(j, "dddd")
    lstDOWRemote.ItemData(lstDOW.NewIndex) = j - 1
  Next
  lstDOW.ListIndex = 0
  lstDOWRemote.ListIndex = 0

  lstDOM.Clear
  lstDOMRemote.Clear
  For j = 1 To 28
    lstDOM.AddItem Format(j)
    lstDOM.ItemData(lstDOM.NewIndex) = j
    lstDOMRemote.AddItem Format(j)
    lstDOMRemote.ItemData(lstDOM.NewIndex) = j

  Next
  lstDOM.ListIndex = 0
  lstDOMRemote.ListIndex = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub
Function Save() As Boolean

  Dim Bitfield        As Integer
  Dim BitFieldRemote  As Integer
  Dim j               As Integer

  ResetActivityTime

  For j = 0 To lstDOW.listcount - 1
    If lstDOW.Selected(j) Then
      Bitfield = Bitfield Or 2 ^ j
    End If
  Next

  If Bitfield = 0 Then Bitfield = 1

  For j = 0 To lstDOWRemote.listcount - 1
    If lstDOWRemote.Selected(j) Then
      BitFieldRemote = BitFieldRemote Or 2 ^ j
    End If
  Next

  If BitFieldRemote = 0 Then BitFieldRemote = 1

  
  Dim DayOM As Integer

  For j = 0 To lstDOM.listcount - 1
    If lstDOM.Selected(j) Then
      DayOM = j + 1
      Exit For
    End If
  Next

  'Dim DaysRemote(1 To 28) As String
  Dim DayRemote As Integer

  For j = 0 To lstDOMRemote.listcount - 1
    If lstDOMRemote.Selected(j) Then
      DayRemote = j + 1
      Exit For
    End If
  Next



  '  If BitField = 0 And chkEnable.value = 1 And OptDaily.value = True Then
  '    Beep
  '    cmdApply.Caption = "Not Saved"
  '
  '  Else
  Configuration.BackupDOW = Bitfield

  'Configuration.BackupDOM = Join(Days, ",")
  Configuration.BackupDOM = Max(DayOM, 1)
  Configuration.BackupTime = cboTime.ItemData(cboTime.ListIndex)
  Configuration.BackupEnabled = chkEnable.Value
  Configuration.BackupFolder = txtFolder.text
  Configuration.BackupType = IIf(OptDaily.Value = True, 0, 1)


  Configuration.BackupDOWRemote = BitFieldRemote
  'Configuration.BackupDOMRemote2 = Join(DaysRemote, ",")
  Configuration.BackupDOMRemote = Max(DayRemote, 1)
  Configuration.BackupTimeRemote = cboTimeRemote.ItemData(cboTimeRemote.ListIndex)
  Configuration.BackupEnabledRemote = chkEnableRemote.Value
  Configuration.BackupFolderRemote = txtFolderRemote.text
  Configuration.BackupTypeRemote = IIf(optDailyRemote.Value = True, 0, 1)


  Configuration.BackupHost = txtHost.text
  Configuration.BackupUser = txtUser.text
  Configuration.BackupPassword = txtPass.text





  WriteSetting "Backup", "Enabled", Configuration.BackupEnabled
  WriteSetting "Backup", "Time", Configuration.BackupTime
  WriteSetting "Backup", "Type", Configuration.BackupType
  WriteSetting "Backup", "DOW", Configuration.BackupDOW
  WriteSetting "Backup", "DOM", Configuration.BackupDOM
  'WriteSetting "Backup", "DOM2", Configuration.BackupDOM2
  WriteSetting "Backup", "Folder", Configuration.BackupFolder


  WriteSetting "RemoteBackup", "Enabled", Configuration.BackupEnabledRemote
  WriteSetting "RemoteBackup", "Time", Configuration.BackupTimeRemote
  WriteSetting "RemoteBackup", "Type", Configuration.BackupTypeRemote
  WriteSetting "RemoteBackup", "DOW", Configuration.BackupDOWRemote
  WriteSetting "RemoteBackup", "DOM", Configuration.BackupDOMRemote
  'WriteSetting "RemoteBackup", "DOM2", Configuration.BackupDOMRemote2
  WriteSetting "RemoteBackup", "Folder", Configuration.BackupFolderRemote



  WriteSetting "RemoteBackup", "Host", Configuration.BackupHost
  WriteSetting "RemoteBackup", "User", Configuration.BackupUser
  WriteSetting "RemoteBackup", "Password", Scramble(Configuration.BackupPassword)




  Save = True
  ClearBackupDate
  cmdApply.Caption = "Save"
  Fill
  '  End If
End Function

Private Sub Label2_Click()

End Sub

Private Sub optDaily_Click()
  updatescreen
End Sub

Private Sub optDailyRemote_Click()
  updatescreen
End Sub

Private Sub optMonthly_Click()
  updatescreen
End Sub


Private Sub tabSettings_BeforeClick(Cancel As Integer)

End Sub


Private Sub optMonthlyRemote_Click()
  updatescreen
End Sub

Private Sub TabStrip_Click()
  SetTabs
End Sub
Private Sub SetTabs()
  Select Case TabStrip.SelectedItem.Key
    Case "remote"
      fraFTP.Visible = True
      fraLocal.Visible = False
      fraFTPsettings.Visible = False
    Case "ftp"
      fraFTPsettings.Visible = True
      fraFTP.Visible = False
      fraLocal.Visible = False
    
    Case Else
      fraLocal.Visible = True
      fraFTP.Visible = False
      fraFTPsettings.Visible = False
  End Select

End Sub

Private Sub txtFolder_Change()
  If ValidateBackupFolder(txtFolder.text) Then
    blbFolder.Caption = "Folder"
    blbFolder.ForeColor = vbBlack
  Else
    blbFolder.Caption = "Folder Unavailable"
    blbFolder.ForeColor = vbRed
  End If
End Sub

Private Sub txtFolder_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If Save() Then
      GetFolder txtFolder.text, Me.name
    End If
  End If
End Sub

Private Sub txtFolderRemote_Change()
  If ValidateBackupFolder(txtFolderRemote.text) Then
    lblFolderRemote.Caption = "Folder"
    lblFolderRemote.ForeColor = vbBlack
  Else
    lblFolderRemote.Caption = "Folder Unavailable"
    lblFolderRemote.ForeColor = vbRed
  End If
End Sub

Private Sub txtFolderRemote_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If Save() Then
      GetFolder txtFolderRemote.text, Me.name
    End If
  End If
End Sub
