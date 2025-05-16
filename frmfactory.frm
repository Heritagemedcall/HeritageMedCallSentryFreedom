VERSION 5.00
Begin VB.Form frmFactory 
   Caption         =   "Factory"
   ClientHeight    =   3150
   ClientLeft      =   840
   ClientTop       =   10320
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3060
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.TextBox txtAdminContact 
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
         Left            =   330
         MaxLength       =   80
         TabIndex        =   3
         Top             =   360
         Width           =   5025
      End
      Begin VB.CommandButton cmdApply 
         Cancel          =   -1  'True
         Caption         =   "Apply"
         Default         =   -1  'True
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
         Left            =   7740
         TabIndex        =   1
         Top             =   1755
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
         TabIndex        =   2
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Contact Info"
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
         Left            =   345
         TabIndex        =   4
         Top             =   75
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Fill()
  txtAdminContact.text = Configuration.AdminContact
End Sub

Private Sub cmdApply_Click()
  Configuration.AdminContact = txtAdminContact.text
  Call WriteSetting("Configuration", "AdminContact", Configuration.AdminContact)
End Sub

Private Sub cmdExit_Click()
  Unload Me
  
End Sub

Private Sub Form_Load()

  ResetActivityTime
  Fill
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ResetActivityTime
  UnHost
End Sub

Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  fraEnabler.BackColor = Me.BackColor
  SetParent fraEnabler.hwnd, hwnd
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

