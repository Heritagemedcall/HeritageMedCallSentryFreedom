VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmExceptionReports 
   Caption         =   "Exception  Reports List"
   ClientHeight    =   3300
   ClientLeft      =   6270
   ClientTop       =   11505
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   9075
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         Left            =   7725
         TabIndex        =   5
         Top             =   1785
         Visible         =   0   'False
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
         TabIndex        =   6
         Top             =   2370
         Width           =   1175
      End
      Begin VB.CommandButton cmdAdd 
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
         TabIndex        =   2
         Top             =   30
         Width           =   1175
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         Top             =   615
         Width           =   1175
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Top             =   1200
         Width           =   1175
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7665
         _ExtentX        =   13520
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmExceptionReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private quitting        As Boolean
Public Caller           As String

Private LastIndex   As Long
Private Reports As Collection



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

Public Sub Fill()
  ShowExeptionReports ' if I could spell
End Sub

Private Sub cmdAdd_Click()
  EditExceptionReport 0
End Sub

Public Sub ShowExeptionReports()
  DisableButtons
  Dim SQl As String
  Dim Report As cExceptionReport
  Dim rs As ADODB.Recordset
  
  SQl = "SELECT * FROM ExceptionReports ORDER BY reportname"
  
  Set Reports = New Collection
  Set rs = ConnExecute(SQl)
  Do Until rs.EOF
    Set Report = New cExceptionReport
    Report.Parse rs
    Reports.Add Report
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  RefreshReports
  'RefreshRooms
  EnableButtons
End Sub
Sub DisableButtons()
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdPrint.Enabled = False
  cmdExit.Enabled = False
End Sub
Sub EnableButtons()
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True
  cmdPrint.Enabled = True
  cmdExit.Enabled = True
  
End Sub


Sub RefreshReports()


  Dim li As ListItem
  Dim index As Long

  Dim Items As Collection
  Dim Item  As cExceptionReport

  lvMain.ListItems.Clear
  LockWindowUpdate lvMain.hwnd

  For Each Item In Reports
    Set li = lvMain.ListItems.Add(, Item.ReportID & "s", Item.ReportName)
    li.SubItems(1) = IIf(Item.Disabled, "X", " ")
    li.SubItems(2) = Item.Comment
  Next
  

  LockWindowUpdate 0


End Sub

Sub Configurelvmain()
  Dim ch As ColumnHeader
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  Me.FontBold = True
  
  Set ch = lvMain.ColumnHeaders.Add(, "Report", "Report", 2500, lvwColumnLeft)
  Set ch = lvMain.ColumnHeaders.Add(, "Disabled", "Off", 500, lvwColumnLeft)
  Set ch = lvMain.ColumnHeaders.Add(, "Comment", "Comment", 2500, lvwColumnLeft)
  lvMain.Sorted = False
End Sub


Private Sub cmdApply_Click()
  'Apply
End Sub

Private Sub cmdDelete_Click()
  Dim SQl As String
  Dim ReportID As Long
  If Not lvMain.SelectedItem Is Nothing Then
      ReportID = Val(lvMain.SelectedItem.Key)
      On Error Resume Next
      SQl = "DELETE FROM ExceptionReports WHERE ReportID = " & ReportID
      ConnExecute SQl
      
      'DeleteExeptionReport ReportID
      Fill
  End If
  
End Sub


Private Sub cmdEdit_Click()
  If lvMain.SelectedItem Is Nothing Then
    ' nada
  Else
    
    EditExceptionReport Val(lvMain.SelectedItem.Key)
    'EditAutoReport Val(lvMain.SelectedItem.Key)
  End If
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdPrint_Click()
If Printer Is Nothing Then Exit Sub
  MsgBox "PrintExeptionReportList"
  'PrintExeptionReportList
  'PrintAutoReportList
End Sub

Private Sub Form_Initialize()
  Set Reports = New Collection
End Sub

Private Sub Form_Load()
  ResetActivityTime
  Configurelvmain
  quitting = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  quitting = True
  UnHost
End Sub


Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  If lvMain.SortKey = ColumnHeader.index - 1 Then
    If lvMain.SortOrder = lvwAscending Then
      lvMain.SortOrder = lvwDescending
    Else
      lvMain.SortOrder = lvwAscending
    End If
  Else
    lvMain.SortOrder = lvwAscending
  End If
  lvMain.SortKey = ColumnHeader.index - 1

End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    LastIndex = Item.index
  End If
End Sub




