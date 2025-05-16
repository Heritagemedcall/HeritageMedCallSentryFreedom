VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmDukane 
   Caption         =   "Dukane"
   ClientHeight    =   3105
   ClientLeft      =   7290
   ClientTop       =   5970
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   9105
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   1785
         Width           =   1175
      End
      Begin VB.Frame fraGeneral 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   15
         TabIndex        =   2
         Top             =   330
         Width           =   7605
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
            Left            =   1560
            TabIndex        =   5
            Top             =   450
            Width           =   1485
         End
         Begin VB.TextBox txtAutoClear 
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
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   4
            ToolTipText     =   "Enter 1 to 999. Zero Disables"
            Top             =   465
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtIncludeWords 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1560
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "Separate Words with Spaces"
            Top             =   1170
            Width           =   5970
         End
         Begin VB.TextBox txtExcludeWords 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1560
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   12
            ToolTipText     =   "Separate Words with Spaces"
            Top             =   1845
            Width           =   5970
         End
         Begin VB.TextBox txtPream 
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
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   8
            ToolTipText     =   "Number of characters to skip from beginning"
            Top             =   810
            Visible         =   0   'False
            Width           =   5970
         End
         Begin VB.Label lbl3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm Words"
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
            TabIndex        =   9
            Top             =   1245
            Width           =   1080
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Auto Clear"
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
            Left            =   510
            TabIndex        =   3
            Top             =   525
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minutes"
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
            Left            =   2340
            TabIndex        =   6
            Top             =   525
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restore Words"
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
            Left            =   135
            TabIndex        =   11
            Top             =   1920
            Width           =   1275
         End
         Begin VB.Label lblPream 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preamble"
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
            Left            =   615
            TabIndex        =   7
            Top             =   870
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VB.Frame fraPort 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2445
         Left            =   60
         TabIndex        =   13
         Top             =   405
         Width           =   5490
         Begin VB.ComboBox cboParity 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   855
            Width           =   2340
         End
         Begin VB.ComboBox cboPort 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   75
            Width           =   2340
         End
         Begin VB.ComboBox cboFlow 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2025
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox cboStop 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1635
            Width           =   2340
         End
         Begin VB.ComboBox cboBits 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1245
            Width           =   2340
         End
         Begin VB.ComboBox cboBaud 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   465
            Width           =   2340
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Parity"
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
            Left            =   1020
            TabIndex        =   18
            Top             =   915
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Bits per second"
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
            Left            =   180
            TabIndex        =   16
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Flow Control"
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
            TabIndex        =   24
            Top             =   2085
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Stop bits"
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
            Left            =   750
            TabIndex        =   22
            Top             =   1695
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Data bits"
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
            Left            =   735
            TabIndex        =   20
            Top             =   1305
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Comm Port"
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
            Left            =   600
            TabIndex        =   14
            Top             =   150
            Width           =   915
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2970
         Left            =   0
         TabIndex        =   1
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
               Caption         =   "Port Setup"
               Key             =   "port"
               Object.ToolTipText     =   "Port and data format setup"
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
Attribute VB_Name = "frmDukane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Serial As String
Private rs As Recordset
Function Save() As Boolean

  If gDukane Is Nothing Then
    Save = False
  Else

    gDukane.AutoClear = Min(Val(txtAutoClear.text), 999)
    gDukane.AlarmWords = Trim(txtIncludeWords.text)
    gDukane.RestoreWords = Trim(txtExcludeWords.text)
    gDukane.SerialPort = GetComboItemData(cboPort)
    gDukane.Baud = GetComboItemData(cboBaud)
    gDukane.Parity = GetParityString(GetComboItemData(Me.cboParity))
    gDukane.BITS = Val(cboBits.text)
    gDukane.Flow = GetComboItemData(cboFlow)
    gDukane.Stopbits = cboStop.text
    '      Device.SerialSettings = Device.serialbaud & Device.SerialParity & Device.Serialbits & Device.SerialStopbits
    gDukane.Preamble = left$(txtPream.text, 100)
    gDukane.Enabled = chkEnabled.Value
    Save = gDukane.Save
    Fill
  End If








End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
  SetControls
  ShowPanel TabStrip.SelectedItem.Key
End Sub
Sub SetControls()
  fraEnabler.BackColor = Me.BackColor
  
  fraGeneral.left = TabStrip.ClientLeft
  fraGeneral.top = TabStrip.ClientTop
  fraGeneral.Height = TabStrip.ClientHeight
  fraGeneral.Width = TabStrip.ClientWidth
  
  fraGeneral.BackColor = Me.BackColor
  fraPort.left = TabStrip.ClientLeft
  fraPort.top = TabStrip.ClientTop
  fraPort.Height = TabStrip.ClientHeight
  fraPort.Width = TabStrip.ClientWidth
  fraPort.BackColor = Me.BackColor
  FillCombos



End Sub
Sub FillCombos()

  Dim j As Integer

  cboPort.Clear
  AddToCombo cboPort, "None", 0
  For j = 1 To 256
    AddToCombo cboPort, "COM " & j, j
  Next
  cboPort.ListIndex = 0

  cboBits.Clear
  For j = 4 To 8
    AddToCombo cboBits, j, j
  Next
  cboBits.ListIndex = cboBits.listcount - 1



  cboBaud.Clear
  'AddToCombo cboBaud, "75", 75
  'AddToCombo cboBaud, "110", 110
  'AddToCombo cboBaud, "150", 150
  AddToCombo cboBaud, "300", 300
  AddToCombo cboBaud, "600", 600
  AddToCombo cboBaud, "1200", 1200
  AddToCombo cboBaud, "2400", 2400
  AddToCombo cboBaud, "4800", 4800
  AddToCombo cboBaud, "7200", 7200
  AddToCombo cboBaud, "9600", 9600
  AddToCombo cboBaud, "14400", 14400
  AddToCombo cboBaud, "19200", 19200
  AddToCombo cboBaud, "38400", 38400
  AddToCombo cboBaud, "57600", 57600
  AddToCombo cboBaud, "115200", 115200
  AddToCombo cboBaud, "128000", 128000

  cboBaud.ListIndex = 6

  cboParity.Clear
  AddToCombo cboParity, "Even", 0
  AddToCombo cboParity, "Odd", 1
  AddToCombo cboParity, "None", 2
  AddToCombo cboParity, "Mark", 3
  AddToCombo cboParity, "Space", 4

  cboParity.ListIndex = 2

  cboStop.Clear
  AddToCombo cboStop, 1, 10
  AddToCombo cboStop, 1.5, 15
  AddToCombo cboStop, 2, 20
  cboStop.ListIndex = 0

  cboFlow.Clear
  AddToCombo cboFlow, "None", 0
  AddToCombo cboFlow, "Hardware", 1
  AddToCombo cboFlow, "Xon/Xoff", 2

  cboFlow.ListIndex = 0


End Sub
Sub ShowPanel(ByVal Key As String)
  Select Case LCase(Key)
    Case "port"
      fraPort.Visible = True
      fraGeneral.Visible = False
    Case Else ' general
      fraGeneral.Visible = True
      fraPort.Visible = False
  End Select
  
End Sub

Private Sub TabStrip_Click()
  ShowPanel TabStrip.SelectedItem.Key
End Sub

Private Sub txtAutoClear_GotFocus()
  SelAll txtAutoClear
End Sub
Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  ResetActivityTime
  Save
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
Sub ResetForm()
    

  txtAutoClear.text = "0"
  txtIncludeWords.text = ""
  txtExcludeWords.text = ""
  cboPort.ListIndex = 0
  cboBaud.ListIndex = 6
  cboBits.ListIndex = cboBits.listcount - 1
  cboParity.ListIndex = 2
  cboStop.ListIndex = 0
  cboFlow.ListIndex = 0

  
End Sub

Public Sub Fill()

  ResetForm


  If Not (gDukane Is Nothing) Then

    chkEnabled.Value = IIf(gDukane.Enabled = 1, 1, 0)
    txtPream.text = Trim$(gDukane.Preamble)
    txtAutoClear.text = gDukane.AutoClear
    txtIncludeWords.text = gDukane.AlarmWords
    txtExcludeWords.text = gDukane.RestoreWords
    cboPort.ListIndex = Max(0, CboGetIndexByItemData(cboPort, gDukane.SerialPort))
    cboBaud.ListIndex = Max(0, CboGetIndexByItemData(cboBaud, gDukane.Baud))
    cboParity.ListIndex = Max(0, CboGetIndexByItemData(cboParity, GetParityID(gDukane.Parity)))
    cboBits.ListIndex = Max(0, CboGetIndexByItemData(cboBits, gDukane.BITS))
    cboStop.ListIndex = GetComboByText(cboStop, gDukane.Stopbits)
    If cboStop.ListIndex < 0 Then
      cboStop.ListIndex = 0
    End If

    cboFlow.ListIndex = Max(0, CboGetIndexByItemData(cboFlow, gDukane.Flow))

  End If


End Sub
