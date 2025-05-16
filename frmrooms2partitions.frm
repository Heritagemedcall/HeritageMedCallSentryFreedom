VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmRooms2Partitions 
   Caption         =   "Rooms to Partitions"
   ClientHeight    =   3300
   ClientLeft      =   615
   ClientTop       =   3705
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   9240
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdUncheckAll 
         Caption         =   "Uncheck All"
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
         TabIndex        =   2
         Top             =   420
         Width           =   1350
      End
      Begin VB.CommandButton cmdCheckAll 
         Caption         =   "Check All"
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
         TabIndex        =   1
         Top             =   60
         Width           =   1350
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
         Height          =   525
         Left            =   7680
         TabIndex        =   8
         Top             =   2460
         Width           =   1175
      End
      Begin VB.CheckBox chkLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "Set as Fixed Location"
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
         Left            =   2700
         TabIndex        =   3
         ToolTipText     =   "Check this Box to Create  Location Partition"
         Top             =   120
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7680
         TabIndex        =   4
         Top             =   120
         Width           =   1175
      End
      Begin VB.TextBox txtSearchBox 
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
         Left            =   3300
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Type Search String"
         Top             =   600
         Width           =   2370
      End
      Begin VB.Timer tmrSearch 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6540
         Top             =   1500
      End
      Begin MSComctlLib.ListView lvAvailRooms 
         Height          =   2115
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "desc"
            Text            =   "Description"
            Object.Width           =   7850
         EndProperty
      End
      Begin VB.Label lblSearch 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   660
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRooms2Partitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HTTPRequest As cHTTPRequest
Private Items       As Collection

Sub RefreshRooms()
  Dim rs            As Recordset
  Dim li            As ListItem
  Dim index         As Long


  Dim Item          As cRoomListItem
  Dim SQl           As String





  On Error GoTo RefreshRooms_Error

  Set Items = New Collection

  If gIsJET Then
    SQl = "SELECT roomid, room FROM Rooms order by room"
  Else
    SQl = "SELECT  roomid, Room  FROM Rooms  ORDER BY Rooms.Room;"
  End If
  Set rs = ConnExecute(SQl)




  Do Until rs.EOF
    Set Item = New cRoomListItem
    Item.RoomID = rs("RoomID")
    Item.Room = rs("Room") & ""
    Items.Add Item
    rs.MoveNext
  Loop

RefreshRooms_Resume:
  On Error Resume Next
  rs.Close
  Set rs = Nothing

  On Error GoTo 0
  Exit Sub

RefreshRooms_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmRooms2Partitons.RefreshRooms." & Erl
  Resume RefreshRooms_Resume

End Sub


Sub Fill()
  Dim li            As ListItem
  Dim Item          As cRoomListItem
  Dim searchstring  As String
  lvAvailRooms.ListItems.Clear
  If Items Is Nothing Then
    Set Items = New Collection
  End If
  searchstring = Trim$(txtSearchBox.text)
  For Each Item In Items
    'i = InStr(item.Room, Trim$(Me.txtSearchBox.text))
    'If i > 0 Then

    If Len(searchstring) Then
      If InStr(1, Item.Room, txtSearchBox.text, vbTextCompare) Then
        Set li = lvAvailRooms.ListItems.Add(, Item.RoomID & "_", Item.Room)
      End If
    Else
      Set li = lvAvailRooms.ListItems.Add(, Item.RoomID & "_", Item.Room)
    End If
  Next



End Sub

Private Sub chkCheckAll_Click()

End Sub

Private Sub chkClear_Click()


End Sub

Private Sub cmdNew_Click()
  On Error Resume Next
  DisableControls
  ConvertRooms
  EnableControls
End Sub

Sub ConvertRooms()


  
  Dim response      As String
  Dim li            As ListItem
  Dim j             As Long
  Dim partition     As cPartition
  Dim IsLocation    As Boolean
  Dim Count         As Long
  Dim XML           As String
  Dim Success       As Boolean
  Dim HTTPRequest   As cHTTPRequest

  
  Set HTTPRequest = New cHTTPRequest
  IsLocation = chkLocation.Value = 1

  For j = 1 To lvAvailRooms.ListItems.Count
    Set li = lvAvailRooms.ListItems(j)
    If li.Checked Then
      'LastIndex = j
      Set partition = New cPartition
      partition.Description = Trim$(li.text)
      partition.IsLocation = IsLocation
      response = HTTPRequest.CreatePartition(GetHTTP & "://" & IP1, USER1, PW1, partition)
      DoEvents
      If IsPartitionCreated(response) Then
        Count = Count + 1
      End If
    End If
  Next
  Call messagebox(Me, Count & " Partitions Created", "Rooms to Partitions", vbOKOnly)

  '  fill


End Sub
Function IsPartitionCreated(ByVal response As String) As Long
  Dim doc           As DOMDocument60
  Dim Node          As IXMLDOMNode
  Dim StatusCode    As Long
  Dim StatusString  As Long
  Dim NewID         As Long
  Dim Success       As Long

  Set doc = New DOMDocument60
  Success = doc.LoadXML(response)
  If Success Then
    Set Node = doc.selectSingleNode("ResponseStatus/statusCode")
    If Not Node Is Nothing Then
      StatusCode = Val(Node.text)
      If StatusCode = 1 Then
        Set Node = doc.selectSingleNode("ResponseStatus/statusString")
        If Not Node Is Nothing Then
          'StatusString = Node.text
        End If
        Set Node = doc.selectSingleNode("ResponseStatus/id")
        If Not Node Is Nothing Then
          NewID = Val(Node.text)
        End If
      End If
    End If
  End If
  IsPartitionCreated = NewID

End Function

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  fraEnabler.BackColor = Me.BackColor
  SetParent fraEnabler.hwnd, hwnd
End Sub

Private Sub cmdCheckAll_Click()
    Dim li            As ListItem
  For Each li In Me.lvAvailRooms.ListItems
    li.Checked = True
  Next
End Sub

Private Sub cmdUncheckAll_Click()
  Dim li            As ListItem
  For Each li In Me.lvAvailRooms.ListItems
    li.Checked = False
  Next
End Sub

Private Sub txtSearchBox_Change()
  tmrSearch.Enabled = False
  tmrSearch.Enabled = True

End Sub
Private Sub tmrSearch_Timer()
  tmrSearch.Enabled = False
  Fill

End Sub
Sub DisableControls()
  lblSearch.Enabled = False
  cmdCheckAll.Enabled = False
  cmdUncheckAll.Enabled = False
  tmrSearch.Enabled = False
  txtSearchBox.Enabled = False
  chkLocation.Enabled = False
  lvAvailRooms.Enabled = False
  cmdCancel.Enabled = False
  cmdNew.Enabled = False
End Sub
Sub EnableControls()

  tmrSearch.Enabled = False
  
  cmdCancel.Enabled = True
  cmdCheckAll.Enabled = True
  cmdUncheckAll.Enabled = True
  lblSearch.Enabled = True
  txtSearchBox.Enabled = True
  cmdNew.Enabled = True
  chkLocation.Enabled = True
  lvAvailRooms.Enabled = True




End Sub

Private Sub Form_Load()
ResetActivityTime
  setcolumns
  fraEnabler.BackColor = Me.BackColor
  RefreshRooms
  Fill
End Sub
Sub setcolumns()

  If lvAvailRooms.ColumnHeaders.Count = 0 Then
    lvAvailRooms.ColumnHeaders.Add , , "Room Desc"
  End If
  Me.lvAvailRooms.ColumnHeaders(1).text = "Room Desc"
  Me.lvAvailRooms.ColumnHeaders(1).Width = 2450

End Sub
