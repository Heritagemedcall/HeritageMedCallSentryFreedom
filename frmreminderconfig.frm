VERSION 5.00
Begin VB.Form frmReminderConfig 
   Caption         =   "Reminder Configuration"
   ClientHeight    =   3855
   ClientLeft      =   5175
   ClientTop       =   5595
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   10230
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         TabIndex        =   3
         Top             =   1785
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
         Left            =   7725
         TabIndex        =   2
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Frame fraGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   150
         TabIndex        =   1
         Top             =   60
         Width           =   7425
         Begin VB.Frame fraProto6 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   0
            TabIndex        =   4
            Tag             =   "Voice Dialer"
            Top             =   0
            Width           =   4695
            Begin VB.ComboBox cboVoices 
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
               Left            =   750
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   2100
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.TextBox txtMsgDelay 
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
               Left            =   1215
               MaxLength       =   2
               TabIndex        =   12
               Top             =   720
               Width           =   510
            End
            Begin VB.TextBox txtMsgSpacing 
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
               Left            =   3945
               MaxLength       =   2
               TabIndex        =   11
               Top             =   720
               Width           =   510
            End
            Begin VB.TextBox txtRedials 
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
               Left            =   1215
               MaxLength       =   2
               TabIndex        =   10
               Top             =   1125
               Width           =   510
            End
            Begin VB.TextBox txtMsgRepeats 
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
               Left            =   2595
               MaxLength       =   2
               TabIndex        =   9
               Top             =   720
               Width           =   510
            End
            Begin VB.TextBox txtRedialDelay 
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
               Left            =   2940
               MaxLength       =   2
               TabIndex        =   8
               Top             =   1125
               Width           =   510
            End
            Begin VB.TextBox txtTag 
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
               Left            =   720
               MaxLength       =   70
               TabIndex        =   7
               Top             =   2310
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.TextBox txtTimeout 
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
               Left            =   1215
               MaxLength       =   3
               TabIndex        =   6
               Top             =   1530
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.ComboBox cboAckDigit 
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
               Left            =   2940
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1530
               Width           =   900
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reminder Voicemail Settings"
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
               Left            =   1140
               TabIndex        =   23
               Top             =   210
               Width           =   2430
            End
            Begin VB.Label lblv 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Voice"
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
               Left            =   195
               TabIndex        =   22
               Top             =   2130
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label z7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Msg Delay"
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
               TabIndex        =   21
               Top             =   780
               Width           =   900
            End
            Begin VB.Label z8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Spacing"
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
               Left            =   3165
               TabIndex        =   20
               Top             =   780
               Width           =   705
            End
            Begin VB.Label z9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Redials"
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
               Left            =   495
               TabIndex        =   19
               Top             =   1185
               Width           =   645
            End
            Begin VB.Label z6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Repeats"
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
               Left            =   1815
               TabIndex        =   18
               Top             =   780
               Width           =   720
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Redial Delay"
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
               Left            =   1770
               TabIndex        =   17
               Top             =   1185
               Width           =   1095
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tag"
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
               Left            =   315
               TabIndex        =   16
               Top             =   2385
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label z10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Timeout"
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
               Left            =   450
               TabIndex        =   15
               Top             =   1590
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Label lblAckDigit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ACK Digit"
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
               Left            =   2025
               TabIndex        =   14
               Top             =   1590
               Width           =   825
            End
         End
      End
   End
End
Attribute VB_Name = "frmReminderConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Save
End Sub

Private Sub Form_Load()
  SetFrames
  FillCombos
  Fill
  ResetActivityTime
End Sub
Sub FillCombos()



  cboAckDigit.Clear
  AddToCombo cboAckDigit, "None", 0
  AddToCombo cboAckDigit, "0", TAPI_DTMF_0
  AddToCombo cboAckDigit, "1", TAPI_DTMF_1
  AddToCombo cboAckDigit, "2", TAPI_DTMF_2
  AddToCombo cboAckDigit, "3", TAPI_DTMF_3
  AddToCombo cboAckDigit, "4", TAPI_DTMF_4
  AddToCombo cboAckDigit, "5", TAPI_DTMF_5
  AddToCombo cboAckDigit, "6", TAPI_DTMF_6
  AddToCombo cboAckDigit, "7", TAPI_DTMF_7
  AddToCombo cboAckDigit, "8", TAPI_DTMF_8
  AddToCombo cboAckDigit, "9", TAPI_DTMF_9
  AddToCombo cboAckDigit, "*", TAPI_DTMF_STAR
  AddToCombo cboAckDigit, "#", TAPI_DTMF_POUND
  cboAckDigit.ListIndex = 0
End Sub

Private Sub SetFrames()
  fraEnabler.BackColor = Me.BackColor
  fraGeneral.BackColor = Me.BackColor
  fraProto6.BackColor = Me.BackColor

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys "{tab}"
  End Select

End Sub
Sub Fill()
  txtMsgDelay.text = Configuration.ReminderMsgDelay
  txtMsgRepeats.text = Configuration.ReminderMsgRepeats
  txtMsgSpacing.text = Configuration.ReminderMsgSpacing
  txtRedials.text = Configuration.ReminderRedials
  txtRedialDelay.text = Configuration.ReminderRedialDelay
  cboAckDigit.ListIndex = Max(0, CboGetIndexByItemData(cboAckDigit, Configuration.ReminderAckDigit))
End Sub
Sub Save()
  Configuration.ReminderMsgDelay = Val(txtMsgDelay.text)
  Configuration.ReminderMsgRepeats = Val(txtMsgRepeats.text)
  Configuration.ReminderMsgSpacing = Val(txtMsgSpacing.text)
  Configuration.ReminderRedials = Val(txtRedials.text)
  Configuration.ReminderRedialDelay = Val(txtRedialDelay.text)
  If cboAckDigit.ListIndex > -1 Then
    Configuration.ReminderAckDigit = cboAckDigit.ItemData(cboAckDigit.ListIndex)
  Else
    Configuration.ReminderAckDigit = 48
  End If

  
  WriteSetting "Reminders", "MsgDelay", Configuration.ReminderMsgDelay
  WriteSetting "Reminders", "MsgRepeats", Configuration.ReminderMsgRepeats
  WriteSetting "Reminders", "MsgSpacing", Configuration.ReminderMsgSpacing
  WriteSetting "Reminders", "Redials", Configuration.ReminderRedials
  WriteSetting "Reminders", "RedialDelay", Configuration.ReminderRedialDelay
  WriteSetting "Reminders", "AckDigit", Configuration.ReminderAckDigit



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

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub
Private Sub txtMsgDelay_GotFocus()
  SelAll txtMsgDelay
End Sub

Private Sub txtMsgDelay_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMsgDelay, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtMsgRepeats_GotFocus()
  SelAll txtMsgRepeats
End Sub

Private Sub txtMsgRepeats_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMsgRepeats, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtMsgSpacing_GotFocus()
  SelAll txtMsgSpacing
End Sub

Private Sub txtMsgSpacing_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMsgSpacing, KeyAscii, False, 0, 2, 99)
End Sub

'Private Sub txtPause_GotFocus()
'  SelAll txtPause
'End Sub
'
'Private Sub txtPause_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtPause, KeyAscii, False, 0, 2, 60)
'End Sub
Private Sub txtRedialDelay_GotFocus()
  SelAll txtRedialDelay
End Sub

Private Sub txtRedialDelay_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtRedialDelay, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtRedials_GotFocus()
  SelAll txtRedials
End Sub

Private Sub txtRedials_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtRedials, KeyAscii, False, 0, 2, 99)
End Sub
Private Sub txtTimeout_GotFocus()
  SelAll txtTimeout
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtTimeout, KeyAscii, False, 0, 3, 999)
End Sub
Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub
