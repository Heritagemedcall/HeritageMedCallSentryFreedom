VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmAutoReports 
   Caption         =   "Auto Reports"
   ClientHeight    =   3075
   ClientLeft      =   6225
   ClientTop       =   8865
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   9150
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
Attribute VB_Name = "frmAutoReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private quitting        As Boolean
Public Caller           As String

Private LastIndex   As Long
Private AutoReports As Collection
'Function DeleteRoom(ByVal RoomID As Long)
'
'  If vbYes = messagebox(Me, "Delete Selected Room?", App.Title, vbQuestion Or vbYesNo) Then
'
'  conn.BeginTrans
'    ConnExecute "UPDATE Residents SET RoomID = 0 WHERE RoomID = " & RoomID
'    ConnExecute "UPDATE Devices SET RoomID = 0 WHERE RoomID = " & RoomID
'    ConnExecute "UPDATE Devices SET RoomID_A = 0 WHERE RoomID_A = " & RoomID
'    ConnExecute "DELETE FROM Rooms WHERE RoomID = " & RoomID
'
'  conn.CommitTrans
'  End If
'
'End Function







Sub Apply()

  Dim SQl    As String

'  Select Case Caller
'    Case "RES"
'      'If mResidentID <> 0 Then
'      If lvMain.SelectedItem Is Nothing Then
'        ' nada
'        Beep
'      Else
'        RoomID = Val(lvMain.SelectedItem.key)
'        If MASTER Then
'          sql = "UPDATE devices SET RoomID = " & RoomID & "  WHERE deviceid = " & mTransmitterID
'          connexecute (sql)
'          Devices.RefreshByID mTransmitterID
'        Else
'          ClientUpdateDeviceRoomID RoomID, mTransmitterID
'          RefreshJet
'        End If
'
'        frmMain.SetListTabs
'        PreviousForm
'        Unload Me
'      End If
'      'End If
'    Case "TX"
'      If mTransmitterID <> 0 Then
'        If lvMain.SelectedItem Is Nothing Then
'          ' nada
'        Else
'          RoomID = Val(lvMain.SelectedItem.key)
'          If MASTER Then
'            sql = "UPDATE Devices SET RoomID = " & RoomID & "  WHERE deviceid = " & mTransmitterID
'            connexecute (sql)
'            Devices.RefreshByID mTransmitterID
'          Else
'            ClientUpdateDeviceRoomID RoomID, mTransmitterID
'            RefreshJet
'          End If
'          frmMain.SetListTabs
'          PreviousForm
'          Unload Me
'
'        End If
'      End If
'    End Select
'
'
'
'

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

Public Sub Fill()
  ShowAutoReports
End Sub

Private Sub cmdAdd_Click()
  EditAutoReport 0
End Sub

Public Sub ShowAutoReports()
  DisableButtons
  Dim SQl As String
  Dim Report As cAutoReport
  SQl = "SELECT * FROM AutoReports ORDER BY reportname"
  Dim rs As ADODB.Recordset
  Set AutoReports = New Collection
  Set rs = ConnExecute(SQl)
  Do Until rs.EOF
    Set Report = New cAutoReport
    Report.Parse rs
    AutoReports.Add Report
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

  Dim rs As Recordset
  Dim li As ListItem
  Dim index As Long

  Dim Items As Collection
  Dim Item  As cAutoReport
  

  Dim CurrentPass As Long
  Static passnumber As Long

  passnumber = passnumber + 1
  If passnumber >= MAXLONG Then
    passnumber = 1
  End If
  CurrentPass = passnumber

  lvMain.ListItems.Clear
  LockWindowUpdate lvMain.hwnd

  For Each Item In AutoReports
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
  Apply
End Sub

Private Sub cmdDelete_Click()
  Dim SQl As String
  Dim ReportID As Long
  If Not lvMain.SelectedItem Is Nothing Then
      ReportID = Val(lvMain.SelectedItem.Key)
      On Error Resume Next
      SQl = "DELETE FROM Autoreports WHERE ReportID = " & ReportID
      ConnExecute SQl
      DeleteAutoReport ReportID
      Fill
  End If
  
End Sub


Private Sub cmdEdit_Click()
  If lvMain.SelectedItem Is Nothing Then
    ' nada
  Else
    EditAutoReport Val(lvMain.SelectedItem.Key)
  End If
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdPrint_Click()
If Printer Is Nothing Then Exit Sub
  PrintAutoReportList
End Sub

Private Sub Form_Initialize()
  Set AutoReports = New Collection
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



