VERSION 5.00
Begin VB.Form frmPush 
   AutoRedraw      =   -1  'True
   Caption         =   "Push"
   ClientHeight    =   3165
   ClientLeft      =   3675
   ClientTop       =   9090
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   8925
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
         Left            =   705
         TabIndex        =   3
         Top             =   1380
         Width           =   1350
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
         Left            =   675
         TabIndex        =   2
         Top             =   945
         Width           =   6720
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
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         Caption         =   "URL to Push Data"
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
         Left            =   675
         TabIndex        =   1
         Top             =   600
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmPush"
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
  Dim PushEnabled As Long
  
  PushEnabled = Val(ReadSetting("Push", "Enabled", "0"))
  txtURL.text = ReadSetting("Push", "URL", "")
  chkEnabled.value = IIf(PushEnabled = 1, 1, 0)
  
End Sub
Sub Save()
  Dim PushEnabled As Long
  
  PushEnabled = chkEnabled.value
  txtURL.text = Trim$(txtURL.text)
  
  If Len(txtURL.text) = 0 Then
    PushEnabled = 0
  End If
  
  WriteSetting "Push", "URL", txtURL.text
  WriteSetting "Push", "Enabled", PushEnabled And 1
  
  Fill
  
End Sub

Private Sub txtURL_GotFocus()
  SelAll txtURL
End Sub
