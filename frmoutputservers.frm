VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmOutputServers 
   Caption         =   "Output Servers"
   ClientHeight    =   3420
   ClientLeft      =   150
   ClientTop       =   2265
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   8910
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
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
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgLst"
         SmallIcons      =   "imgLst"
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
Attribute VB_Name = "frmOutputServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mOutPutID As Long


Public Property Get OutputID() As Long
  OutputID = mOutPutID
End Property

Public Property Let OutputID(ByVal OutputID As Long)
  mOutPutID = OutputID
  cmdApply.Visible = mOutPutID <> 0
End Property


Private Sub cmdAdd_Click()
  SpecialLog "Call EditServer 0"
  EditServer 0
  
End Sub

Private Sub cmdDelete_Click()
  Dim ID As Long

  If Not lvMain.SelectedItem Is Nothing Then
    ID = Val(lvMain.SelectedItem.Key)
    If vbYes = messagebox(Me, "Delete this output server?", App.Title, vbYesNo Or vbQuestion) Then
      conn.BeginTrans
      ConnExecute "DELETE FROM PagerDevices WHERE ID = " & ID
      ConnExecute "UPDATE Pagers SET DeviceID = 0 WHERE DeviceID = " & ID
      conn.CommitTrans
      ChangePageDevice ID, 3  ' delete
      Fill
    End If
  End If
End Sub

Private Sub cmdEdit_Click()
  EditServer GetActiveServer
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me

End Sub

Private Sub Form_Load()
  Configurelvmain
  ResetActivityTime
End Sub


Private Sub Configurelvmain()
  Dim ch As ColumnHeader
  lvMain.ListItems.Clear
  lvMain.ColumnHeaders.Clear
  Set ch = lvMain.ColumnHeaders.Add(, "DX", "Server Description", 2500)
  Set ch = lvMain.ColumnHeaders.Add(, "PX", "Protocol", 2500)
  Set ch = lvMain.ColumnHeaders.Add(, "SX", "Settings", 2250)
  lvMain.Sorted = True
End Sub
Function GetActiveServer() As Long
  If Not lvMain.SelectedItem Is Nothing Then
    GetActiveServer = Val(lvMain.SelectedItem.Key)
  End If
End Function
Sub DisableButtons()
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdApply.Enabled = False
  cmdExit.Enabled = False
End Sub
Sub EnableButtons()
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub

Public Sub Fill()
  DisableButtons

  Dim rs As Recordset
  Dim li As ListItem
  Dim ProtocolID As Long
  
  lvMain.ListItems.Clear

  Set rs = ConnExecute("SELECT * FROM PagerDevices order by ID")
  Do Until rs.EOF
    ProtocolID = rs("ProtocolID")
    Set li = lvMain.ListItems.Add(, rs("ID") & "D", rs("Description") & "")
    li.SubItems(1) = ProtocolString(ProtocolID)
    If rs("Port") = 0 Then
      li.SubItems(2) = rs("Settings") & ""
    Else
      Select Case ProtocolID
        Case PROTOCOL_TAP_IP
          li.SubItems(2) = rs("DialerPhone") & ":" & rs("port") & ""
        Case Else
          li.SubItems(2) = "COM" & rs("Port") & " " & rs("Settings") & ""
      End Select
    End If
    rs.MoveNext

  Loop
  rs.Close
  Set rs = Nothing
  EnableButtons
  If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems.Item(1).Selected = True
  End If
End Sub


Sub EditServer(ByVal ID As Long)
  EditOutputServer ID

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

