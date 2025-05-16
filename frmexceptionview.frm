VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmExceptionView 
   Caption         =   "Exception Viewer"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   6255
   ClientWidth     =   9840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   9840
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10440
      Begin VB.CommandButton cmdFile 
         Caption         =   "File"
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
         Left            =   8910
         Picture         =   "frmExceptionView.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   810
         Width           =   870
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
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
         Left            =   8910
         Picture         =   "frmExceptionView.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   870
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
         Left            =   8910
         TabIndex        =   4
         Top             =   2415
         Width           =   870
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   5265
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Loading..."
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmExceptionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ReportID As String
Private ThisReport As Object
Public Sub Fill()
  Dim NewReport As cExceptionReportData
  cmdExit.Enabled = False
  Screen.MousePointer = vbHourglass
  Set NewReport = New cReports
  Set ThisReport = NewReport
  lvMain.ListItems.Clear
  Set NewReport.lv = lvMain
  NewReport.Fill ReportID
  Set NewReport = Nothing
  
  Screen.MousePointer = vbNormal
  cmdExit.Enabled = True
End Sub



Public Function AdvancedReport(ByVal ReportType As Integer, ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
  Dim NewReport As cExceptionReportData
  
  cmdExit.Enabled = False
  Screen.MousePointer = vbHourglass
  
  Set NewReport = New cExceptionReportData
  Set ThisReport = NewReport
  lvMain.ListItems.Clear
  Set NewReport.lv = lvMain
  NewReport.FillAdvanced ReportType, Criteria, StartDate, EndDate
  Set NewReport = Nothing
  Screen.MousePointer = vbNormal
  cmdExit.Enabled = True
End Function

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub



Private Sub cmdFile_Click()
  ResetActivityTime
  cmdExit.Enabled = False
  Screen.MousePointer = vbHourglass
  ThisReport.Dest = 2
  ThisReport.PrintReport
  Screen.MousePointer = vbNormal
  cmdExit.Enabled = True

End Sub

Private Sub cmdPrint_Click()
  ResetActivityTime
If Printer Is Nothing Then Exit Sub
  cmdExit.Enabled = False
  Screen.MousePointer = vbHourglass
  
  ThisReport.Dest = 1
  ThisReport.PrintReport
  Screen.MousePointer = vbNormal
  cmdExit.Enabled = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
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

Private Sub lvMain_Click()
  Dim Key As Long
  If Not lvMain.SelectedItem Is Nothing Then
    Key = Val(lvMain.SelectedItem.Key)
    frmMain.DisplayResidentInfo Key, 0
  End If
End Sub

