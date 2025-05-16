VERSION 5.00
Begin VB.Form frmReportChoices 
   Caption         =   "Reports"
   ClientHeight    =   3150
   ClientLeft      =   1245
   ClientTop       =   7035
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdPerfReports 
         Caption         =   "Exception Reports"
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
         Left            =   6240
         TabIndex        =   20
         Top             =   135
         Width           =   1350
      End
      Begin VB.CommandButton cmdAutoReports 
         Caption         =   "Automatic Reports"
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
         Left            =   4890
         TabIndex        =   19
         Top             =   135
         Width           =   1350
      End
      Begin VB.CommandButton cmdFolder 
         Caption         =   "Folder"
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   1175
      End
      Begin VB.Frame fra1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   300
         TabIndex        =   5
         Top             =   930
         Width           =   7890
         Begin VB.ComboBox cboReporttype 
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
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   540
            Width           =   1965
         End
         Begin VB.ComboBox cboEventType 
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
            Left            =   1978
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   1995
         End
         Begin VB.TextBox txtSearch 
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
            Left            =   1978
            TabIndex        =   13
            Top             =   540
            Width           =   2040
         End
         Begin VB.TextBox txtdateFrom 
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
            Left            =   4050
            TabIndex        =   14
            Top             =   540
            Width           =   1245
         End
         Begin VB.TextBox txtDateTo 
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
            Left            =   5340
            TabIndex        =   15
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label lblStart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
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
            Left            =   4050
            TabIndex        =   9
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
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
            Left            =   5400
            TabIndex        =   10
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report Type"
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
            Left            =   15
            TabIndex        =   7
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label lblCriteria 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "         "
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
            Left            =   2040
            TabIndex        =   8
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblDateError 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000CC&
            Height          =   195
            Left            =   5280
            TabIndex        =   6
            Top             =   0
            Width           =   90
         End
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
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
         Left            =   5595
         TabIndex        =   16
         Top             =   2370
         Width           =   1175
      End
      Begin VB.CommandButton cmdThisWeek 
         Caption         =   "This Week"
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
         Left            =   2550
         TabIndex        =   3
         Top             =   135
         Width           =   1175
      End
      Begin VB.CommandButton cmdToday 
         Caption         =   "Today"
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
         Left            =   1380
         TabIndex        =   2
         Top             =   135
         Width           =   1175
      End
      Begin VB.CommandButton cmdCurrentShift 
         Caption         =   "Current Shift"
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
         Left            =   210
         TabIndex        =   1
         Top             =   135
         Width           =   1175
      End
      Begin VB.CommandButton cmdThisMonth 
         Caption         =   "This Month"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   135
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
         TabIndex        =   17
         Top             =   2370
         Width           =   1175
      End
   End
End
Attribute VB_Name = "frmReportChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cboReporttype_Click()
  Dim i As Long
  i = GetComboItemData(cboReporttype)

  Select Case i
    Case 0
      lblCriteria.Caption = ""
      txtSearch.Visible = False
      cboEventType.Visible = False
      txtdateFrom.Visible = False
      txtDateTo.Visible = False
    Case RPT_ROOM  ' "room"
      lblCriteria.Caption = "Search"
      txtSearch.Visible = True
      cboEventType.Visible = False
      txtdateFrom.Visible = True
      txtDateTo.Visible = True

    Case RPT_RES  '"resident"
      txtSearch.Visible = True
      lblCriteria.Caption = "Search"
      cboEventType.Visible = False
      txtdateFrom.Visible = True
      txtDateTo.Visible = True

    Case RPT_DEVICE  '"device"
      lblCriteria.Caption = "Search"
      txtSearch.Visible = True
      cboEventType.Visible = False
      txtdateFrom.Visible = True
      txtDateTo.Visible = True

    Case RPT_EVENT  '"event"
      lblCriteria.Caption = "Event"
      cboEventType.Visible = True
      txtSearch.Visible = False
      cboEventType.ListIndex = 0
      txtdateFrom.Visible = True
      txtDateTo.Visible = True

    Case RPT_ASSUR  '"assurance"
      lblCriteria.Caption = ""
      txtSearch.Visible = False
      cboEventType.Visible = False
      txtdateFrom.Visible = True
      txtDateTo.Visible = True

    Case RPT_DEVHIST  '"device history"
      lblCriteria.Caption = "Search"
      txtSearch.Visible = True
      cboEventType.Visible = False
      txtdateFrom.Visible = True
      txtDateTo.Visible = True

    Case RPT_RESHIST '"resident history"
      lblCriteria.Caption = "Search"
      txtSearch.Visible = True
      cboEventType.Visible = False
      txtdateFrom.Visible = True
      txtDateTo.Visible = True
    Case Else
      lblCriteria.Caption = ""
      txtSearch.Visible = False
      cboEventType.Visible = False
      txtdateFrom.Visible = False
      txtDateTo.Visible = False


  End Select
End Sub

Private Sub cmdAutoReports_Click()
  ListAutoReports
End Sub

Private Sub cmdCurrentShift_Click()
  BasicReport "currentshift"


End Sub

Private Sub cmdFolder_Click()
  HostForm frmReportPath
End Sub

Private Sub cmdGo_Click()
  If ValidateDates() Then
    Dim i As Long
    i = GetComboItemData(cboReporttype)
    If i = RPT_EVENT Then
      AdvancedReport GetComboItemData(cboReporttype), GetComboItemData(cboEventType), Format(Me.txtdateFrom.text, "mm/dd/yyyy"), Format(Me.txtDateTo.text, "mm/dd/yyyy")
    Else
      AdvancedReport GetComboItemData(cboReporttype), Trim(Me.txtSearch), Format(Me.txtdateFrom.text, "mm/dd/yyyy"), Format(Me.txtDateTo.text, "mm/dd/yyyy")
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub


Private Sub cmdPerfReports_Click()
  HostForm frmExceptionReports
  frmExceptionReports.Fill

End Sub

Private Sub cmdThisMonth_Click()
  BasicReport "thismonth"
End Sub

Private Sub cmdThisWeek_Click()
  BasicReport "thisweek"
End Sub

Private Sub cmdToday_Click()
  BasicReport "today"
End Sub

Sub ResetForm()
  lblDateError.Caption = ""
  txtdateFrom.text = Format(Now, "mm/dd/yy")
  txtDateTo.text = Format(Now, "mm/dd/yy")

End Sub



Private Sub Form_Activate()
  updatescreen
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
  FillCombos
  updatescreen
End Sub
Sub updatescreen()
  Select Case gUser.LEvel
  
  Case LEVEL_FACTORY  ' Factory
      cmdAutoReports.Visible = True
      cmdPerfReports.Visible = True
    Case LEVEL_ADMIN  ' Admin 2
      cmdAutoReports.Visible = True
      cmdPerfReports.Visible = True
    Case Else
      cmdAutoReports.Visible = False
      cmdPerfReports.Visible = False
  End Select


End Sub
Sub FillCombos()
'AddToCombo cboReporttype, "Select", 0
  cboReporttype.Clear
  AddToCombo cboReporttype, "Room", 1
  AddToCombo cboReporttype, "Resident", 2
  AddToCombo cboReporttype, "Device", 3
  AddToCombo cboReporttype, "Event", 4
  AddToCombo cboReporttype, "Check-in", 5
  AddToCombo cboReporttype, "Device History", 6
  AddToCombo cboReporttype, "Resident History", 7
  cboReporttype.ListIndex = 0

  cboEventType.Clear
  AddToCombo cboEventType, "All", 0
  AddToCombo cboEventType, "Alarms", EVT_EMERGENCY
  AddToCombo cboEventType, "Alerts", EVT_ALERT
  AddToCombo cboEventType, "Low Battery", EVT_BATTERY_FAIL
  AddToCombo cboEventType, "Trouble", EVT_CHECKIN_FAIL
  AddToCombo cboEventType, "Tamper", EVT_TAMPER
  'AddToCombo cboEventType, "Comm Error", EVT_COMM_TIMEOUT
  cboEventType.ListIndex = 0

  ResetForm
End Sub

Private Sub Form_Paint()
'  Select Case gUser.Level
'    Case LEVEL_FACTORY  ' Factory
'      cmdAutoReports.Visible = True
'    Case LEVEL_ADMIN  ' Admin 2
'      cmdAutoReports.Visible = True
'    Case Else
'      cmdAutoReports.Visible = False
'  End Select
End Sub

Private Sub Form_Resize()
'    Select Case gUser.Level
'    Case LEVEL_FACTORY  ' Factory
'      cmdAutoReports.Visible = True
'    Case LEVEL_ADMIN  ' Admin 2
'      cmdAutoReports.Visible = True
'    Case Else
'      cmdAutoReports.Visible = False
'  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
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

Public Sub Fill()
 updatescreen
 ' empty for now
End Sub

Private Sub txtdateFrom_Change()
  Dim s As String
  s = Trim(txtdateFrom.text)
  If IsDate(s) Then
    lblStart.ForeColor = &H80000012
  Else
    lblStart.ForeColor = vbRed
  End If

End Sub

Private Sub txtdateFrom_GotFocus()
  SelAll txtdateFrom
  
End Sub

Private Sub txtdateFrom_KeyPress(KeyAscii As Integer)
  Dim newval As Date
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      If IsDate(txtdateFrom.text) Then
        newval = DateAdd("d", 1, txtdateFrom.text)
        txtdateFrom.text = Format(newval, "mm/dd/yy")
      End If
    Case vbKeySubtract, 45
      KeyAscii = 0
      If IsDate(txtdateFrom.text) Then
        newval = DateAdd("d", -1, txtdateFrom.text)
        txtdateFrom.text = Format(newval, "mm/dd/yy")
      End If

    Case Else
  End Select
End Sub

Private Sub txtdateFrom_LostFocus()
  ValidateDates
End Sub

Private Sub txtDateTo_Change()
  Dim s As String
  s = Trim(txtDateTo.text)
  If IsDate(s) Then
    lblEnd.ForeColor = &H80000012
  Else
    lblEnd.ForeColor = vbRed
  End If
End Sub

Private Sub txtDateTo_GotFocus()
  SelAll txtDateTo
  
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
  Dim newval As Date
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      If IsDate(txtDateTo.text) Then
        newval = DateAdd("d", 1, txtDateTo.text)
        txtDateTo.text = Format(newval, "mm/dd/yy")
      End If
    Case vbKeySubtract, 45
      KeyAscii = 0
      If IsDate(txtDateTo.text) Then
        newval = DateAdd("d", -1, txtDateTo.text)
        txtDateTo.text = Format(newval, "mm/dd/yy")
      End If

    Case Else
  End Select

End Sub

Private Sub txtDateTo_LostFocus()
  ValidateDates

End Sub
Function ValidateDates() As Boolean
  Dim s As String
  
  lblDateError.Caption = ""
  s = Trim(txtdateFrom.text)
  If IsDate(s) Then
    txtdateFrom.text = Format(s, "mm/dd/yy")
    s = Trim(txtDateTo.text)
    If IsDate(s) Then
      txtDateTo.text = Format(s, "mm/dd/yy")
      ValidateDates = True
    Else
      lblDateError.Caption = "End Date Invalid"
    End If
  Else
    lblDateError.Caption = "Start Date Invalid"
  End If

End Function

Private Sub txtSearch_GotFocus()
  SelAll txtSearch
End Sub
