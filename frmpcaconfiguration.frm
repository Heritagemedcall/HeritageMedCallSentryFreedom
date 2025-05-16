VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPCAConfiguration 
   Caption         =   "PCA Configuration"
   ClientHeight    =   3270
   ClientLeft      =   -30
   ClientTop       =   2295
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   9105
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Frame fraSetup 
         BorderStyle     =   0  'None
         Height          =   2475
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7365
         Begin VB.TextBox txtSerial 
            BackColor       =   &H00FFFFFF&
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
            Left            =   930
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   150
            Width           =   1500
         End
         Begin VB.Label lblSerial 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Serial"
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
            Left            =   195
            TabIndex        =   8
            Top             =   195
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
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
         TabStop         =   0   'False
         Top             =   30
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   2370
         Width           =   1175
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   3015
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5318
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PCA Setup"
               Key             =   "tx"
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
      Begin VB.Label lblDecimal1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11310
         TabIndex        =   5
         Top             =   1395
         Visible         =   0   'False
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmPCAConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Serial As String

Private Function Save() As Boolean

End Function

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub Form_Load()
ResetActivityTime
   UpdateScreenElements
End Sub
Sub UpdateScreenElements()
  fraEnabler.BackColor = Me.BackColor


  fraSetup.left = TabStrip.ClientLeft
  fraSetup.top = TabStrip.ClientTop
  fraSetup.Width = TabStrip.ClientWidth
  fraSetup.Height = TabStrip.ClientHeight
End Sub
Public Sub Fill()

End Sub

Public Sub Display()
  txtSerial.text = Serial
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

