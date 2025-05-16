VERSION 5.00
Begin VB.Form frmSMTPSetup 
   Caption         =   "SMTP Setup"
   ClientHeight    =   2625
   ClientLeft      =   2790
   ClientTop       =   7905
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   9645
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
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
         Left            =   8010
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Click to Send Test Email"
         Top             =   300
         Width           =   1175
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
         Left            =   8010
         TabIndex        =   22
         Top             =   1995
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
         Left            =   8010
         TabIndex        =   21
         Top             =   1410
         Width           =   1175
      End
      Begin VB.Frame fraEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   30
         TabIndex        =   2
         Top             =   0
         Width           =   7905
         Begin VB.TextBox txtPort 
            Alignment       =   1  'Right Justify
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
            Left            =   5730
            MaxLength       =   5
            TabIndex        =   19
            Text            =   "25"
            ToolTipText     =   "Standard is 25  Gmail uses 465"
            Top             =   1530
            Width           =   735
         End
         Begin VB.CheckBox chkDebug 
            Caption         =   "Debug"
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
            Left            =   5700
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Mail Body is HTML / Plain Text"
            Top             =   2010
            Width           =   1035
         End
         Begin VB.CheckBox chkUseSMTP 
            Caption         =   "Use SMTP Mailer"
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
            Left            =   1320
            TabIndex        =   1
            ToolTipText     =   "Mail Body is HTML / Plain Text"
            Top             =   60
            Width           =   2505
         End
         Begin VB.TextBox txtUsername 
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
            Left            =   1320
            TabIndex        =   12
            Top             =   1800
            Width           =   3030
         End
         Begin VB.TextBox txtPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   2190
            Width           =   3030
         End
         Begin VB.CheckBox ckLogin 
            Caption         =   "Require Login"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5670
            TabIndex        =   15
            ToolTipText     =   "Use Login Authorization When Connecting to a Host"
            Top             =   390
            Width           =   2055
         End
         Begin VB.CheckBox ckHtml 
            Caption         =   "Html"
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
            Left            =   5700
            TabIndex        =   18
            ToolTipText     =   "Mail Body is HTML / Plain Text"
            Top             =   1170
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CheckBox ckPopLogin 
            Caption         =   "Require POP Login"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5670
            TabIndex        =   16
            ToolTipText     =   "Use Login Authorization When Connecting to a Host"
            Top             =   750
            Width           =   2085
         End
         Begin VB.TextBox txtServer 
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
            Left            =   1320
            TabIndex        =   4
            ToolTipText     =   "Use SSL:// Prefix for Gmail and other Secure Sockets enabled Servers"
            Top             =   360
            Width           =   4200
         End
         Begin VB.TextBox txtFromName 
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
            Left            =   1320
            TabIndex        =   8
            Top             =   1080
            Width           =   4200
         End
         Begin VB.TextBox txtFrom 
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
            Left            =   1320
            TabIndex        =   10
            Top             =   1440
            Width           =   4200
         End
         Begin VB.TextBox txtPopServer 
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
            Left            =   1320
            TabIndex        =   6
            Top             =   720
            Width           =   4200
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port"
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
            Left            =   6600
            TabIndex        =   24
            Top             =   1590
            Width           =   360
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
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
            Left            =   330
            TabIndex        =   11
            Top             =   1830
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   420
            TabIndex        =   13
            Top             =   2220
            Width           =   825
         End
         Begin VB.Label lblServer 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Server"
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
            Left            =   150
            TabIndex        =   3
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label lblFromName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sender Name"
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
            Left            =   150
            TabIndex        =   7
            Top             =   1110
            Width           =   1155
         End
         Begin VB.Label lblFrom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sender Email"
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
            Left            =   150
            TabIndex        =   9
            Top             =   1470
            Width           =   1125
         End
         Begin VB.Label lblPopServer 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "POP3 Server"
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
            Left            =   210
            TabIndex        =   5
            Top             =   780
            Width           =   1110
         End
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Used"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmSMTPSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdTest_Click()
  SendTestSMTPMessage
End Sub

Sub SendTestSMTPMessage()
        'Dim mapi    As Object
        Dim Subject As String
        Dim message As String

10      message = "SMTP Setup Test"



20      On Error Resume Next



30      If gSMTPMailer Is Nothing Then
'40        Set gSMTPMailer = new sendmail

        'Set gSMTPMailer = New SMTPMailer.SendMail

40        Set gSMTPMailer = CreateObject("smtpmailer.sendmail")
50      End If
60      If gSMTPMailer Is Nothing Then
70        LogProgramError "Could not create SMTPMailer Object in frmSMTPSetup." & Erl
80      Else
          If 0 Then
            Call gSMTPMailer.Send("", "", Configuration.MailSenderEmail & ";" & "tkiehl1@tampabay.rr.com", "SMTP Setup Test", message, "")
          Else
90          Call gSMTPMailer.Send("", "", Configuration.MailSenderEmail, "SMTP Setup Test", message, "")
          End If
100     End If

        If Err.Number Then
          LogProgramError Err.Description & " " & Err.Number & vbCrLf & "Error in frmSMTPSetup." & Erl
        End If
        '// Username, Password, Address, Subject,Body, AttachmentsList ' Attachemnet list is a semicolon ";" delimited list of file attachments

        'Set mapi = Nothing

End Sub

Private Sub Form_Load()
ResetActivityTime
 SetControls
End Sub
Private Sub SetControls()
    Dim f As Control

  For Each f In Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor

    End If
  Next
End Sub

Private Sub Form_Unload(Cancel As Integer)

  UnHost
End Sub

Private Sub cmdOK_Click()
  ResetActivityTime
  
  Configuration.UseSMTP = IIf(chkUseSMTP.Value = 1, 1, 0)
  Configuration.MailUserName = txtUsername.text
  Configuration.MailPassword = txtPassword.text
  Configuration.MailSMTPserver = txtServer.text
  Configuration.MailPOP3Server = txtPopServer.text
  Configuration.MailSenderEmail = txtFrom.text
  Configuration.MailSenderName = txtFromName.text
  
  
  Configuration.MailRequireLogin = IIf(ckLogin.Value = 1, 1, 0)
  Configuration.MailRequirePopLogin = IIf(ckPopLogin.Value = 1, 1, 0)
  Configuration.MailDebug = IIf(Me.chkDebug.Value = 1, 1, 0)
  Configuration.MailPort = Val(txtPort.text)
  
  
  WriteSetting "Configuration", "UseSMTP", Configuration.UseSMTP
  WriteSetting "Configuration", "MailUserName", Configuration.MailUserName
  WriteSetting "Configuration", "MailSMTPserver", Configuration.MailSMTPserver
  WriteSetting "Configuration", "MailPOP3Server", Configuration.MailPOP3Server
  WriteSetting "Configuration", "MailSenderEmail", Configuration.MailSenderEmail
  WriteSetting "Configuration", "MailSenderName", Configuration.MailSenderName
  WriteSetting "Configuration", "MailRequirePopLogin", Configuration.MailRequirePopLogin
  WriteSetting "Configuration", "MailRequireLogin", Configuration.MailRequireLogin
  WriteSetting "Configuration", "MailPassword", Scramble(Configuration.MailPassword)
  WriteSetting "Configuration", "MailDebug", Configuration.MailDebug
  WriteSetting "Configuration", "MailPort", Configuration.MailPort
  
End Sub
Public Sub Fill()

'  dbg "Filling Form"
'  dbg "SMTP UseSMTP     " & Configuration.UseSMTP
'  dbg "SMTP Username    " & Configuration.MailUserName
'  dbg "SMTP SMTPserver  " & Configuration.MailSMTPserver
'  dbg "SMTP SenderEmail " & Configuration.MailSenderEmail
'  dbg "SMTP SenderName  " & Configuration.MailSenderName
'  dbg "SMTP MailPort    " & Configuration.MailPort

  chkUseSMTP.Value = IIf(Configuration.UseSMTP, 1, 0)
  txtUsername.text = Configuration.MailUserName
  txtPassword.text = Configuration.MailPassword
  txtServer.text = Configuration.MailSMTPserver
  txtPopServer.text = Configuration.MailPOP3Server
  txtFrom.text = Configuration.MailSenderEmail
  txtFromName.text = Configuration.MailSenderName
  txtPort.text = Configuration.MailPort

  ckLogin.Value = IIf(Configuration.MailRequireLogin, 1, 0)
  ckPopLogin.Value = IIf(Configuration.MailRequirePopLogin, 1, 0)
  chkDebug.Value = IIf(Configuration.MailDebug, 1, 0)


End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Private Sub cmdExit_Click()


End Sub
Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  SetParent fraEnabler.hwnd, hwnd
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtPort, KeyAscii, False, 0, 5, 99999)
End Sub
