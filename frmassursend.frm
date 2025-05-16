VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmAssurSend 
   Caption         =   "Check-ins Send"
   ClientHeight    =   7665
   ClientLeft      =   2205
   ClientTop       =   2460
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   6465
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Frame fraEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   7425
         Begin VB.CommandButton Command1 
            Caption         =   "End Assur"
            Height          =   435
            Left            =   5040
            TabIndex        =   17
            Top             =   600
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox txtEmailRecipient 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   7
            Top             =   495
            Width           =   3795
         End
         Begin VB.TextBox txtEmailSubject 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   6
            Top             =   1155
            Width           =   3795
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Recipient"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Line"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   900
            Width           =   1080
         End
      End
      Begin VB.Frame fraGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   7425
         Begin VB.Frame fraFileType 
            BorderStyle     =   0  'None
            Caption         =   "FileType"
            Height          =   1605
            Left            =   2940
            TabIndex        =   12
            Top             =   270
            Width           =   3945
            Begin VB.OptionButton optTabDelimitedNoHeader 
               Caption         =   "Tab Delimited Table / NO Headers"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   15
               Top             =   750
               Width           =   3585
            End
            Begin VB.OptionButton optHTML 
               Caption         =   "HTML Document"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   14
               Top             =   1110
               Width           =   3405
            End
            Begin VB.OptionButton optTabDelimited 
               Caption         =   "Tab Delimited Table"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   13
               Top             =   390
               Value           =   -1  'True
               Width           =   3285
            End
            Begin VB.Label lblFileFormat 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "File Format:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   90
               Width           =   1005
            End
         End
         Begin VB.CheckBox chkSendAsEmail 
            Caption         =   "Send As Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   390
            TabIndex        =   11
            Top             =   330
            Width           =   2295
         End
         Begin VB.CheckBox chkSaveAsFile 
            Caption         =   "Save As File"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   390
            TabIndex        =   10
            Top             =   300
            Visible         =   0   'False
            Width           =   2295
         End
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   1785
         Width           =   1175
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2970
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   5239
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "General"
               Object.Tag             =   "general"
               Object.ToolTipText     =   "General Settings"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Email Settings"
               Key             =   "email"
               Object.ToolTipText     =   "Configure Email Settings"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAssurSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Save
  
End Sub
Function Save() As Boolean

'Public AssurSaveAsFile      As Integer ' 0 or 1
'Public AssurSendAsEmail     As Integer ' 0 or 1
'Public AssurFileFormat      As Integer ' 0=TabDelimited; 1=TabDelimited-noheader;2=HTML
'Public AssurEmailRecipient  As String
'Public AssurEmailSubject    As String


  
    Select Case True
      Case optTabDelimitedNoHeader.Value
        Configuration.AssurFileFormat = 1
      Case optHTML.Value
        Configuration.AssurFileFormat = 2
      Case Else  ' optTabDelimited.Value
        Configuration.AssurFileFormat = 0
    End Select
      
  
  Configuration.AssurSaveAsFile = IIf(chkSaveAsFile.Value <> 0, 1, 0)
  Configuration.AssurSendAsEmail = IIf(chkSendAsEmail.Value <> 0, 1, 0)
  Configuration.AssurEmailRecipient = Trim$(txtEmailRecipient.text)
  Configuration.AssurEmailSubject = Trim$(txtEmailSubject.text)
  
  
  WriteSetting "Assurance", "SaveAsFile", Configuration.AssurSaveAsFile
  WriteSetting "Assurance", "SendAsEmail", Configuration.AssurSendAsEmail
  WriteSetting "Assurance", "FileFormat", Configuration.AssurFileFormat
  WriteSetting "Assurance", "EmailRecipient", Configuration.AssurEmailRecipient
  WriteSetting "Assurance", "EmailSubject", Configuration.AssurEmailSubject

End Function

Private Sub Command1_Click()
  EndAssure
End Sub

Private Sub Form_Load()
  ResetActivityTime
  Command1.Visible = InIDE
  SetControls
  ShowPanel TabStrip.SelectedItem.Key
End Sub
Sub SetControls()
  Dim f As Control
  
  For Each f In Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor
    End If
  Next
  
  
  fraGeneral.left = TabStrip.ClientLeft
  fraGeneral.top = TabStrip.ClientTop
  fraGeneral.Height = TabStrip.ClientHeight
  fraGeneral.Width = TabStrip.ClientWidth
  
  fraEmail.left = TabStrip.ClientLeft
  fraEmail.top = TabStrip.ClientTop
  fraEmail.Height = TabStrip.ClientHeight
  fraEmail.Width = TabStrip.ClientWidth
  





End Sub

Sub ShowPanel(ByVal Key As String)
  Select Case LCase(Key)
    Case "email"
      fraEmail.Visible = True
      fraGeneral.Visible = False
    Case Else ' general
      fraGeneral.Visible = True
      fraEmail.Visible = False
  End Select
End Sub

Public Sub Fill()
  txtEmailRecipient.text = Configuration.AssurEmailRecipient
  txtEmailSubject.text = Configuration.AssurEmailSubject

  Select Case Configuration.AssurFileFormat
    Case 1
      optTabDelimitedNoHeader.Value = True
    Case 2
      optHTML.Value = True
    Case Else
      optTabDelimited.Value = True
  End Select

  chkSaveAsFile.Value = 1 ' IIf(Configuration.AssurSaveAsFile, 1, 0)
  chkSendAsEmail.Value = IIf(Configuration.AssurSendAsEmail, 1, 0)

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

Private Sub TabStrip_Click()
  ShowPanel TabStrip.SelectedItem.Key
End Sub
