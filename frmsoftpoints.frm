VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSoftPoints 
   Caption         =   "Soft Points"
   ClientHeight    =   6180
   ClientLeft      =   810
   ClientTop       =   3375
   ClientWidth     =   13095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   13095
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9705
      Begin VB.Frame fraLoading 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   3060
         TabIndex        =   14
         Top             =   660
         Width           =   3615
         Begin VB.Label lblLoading 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Loading..."
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
            Left            =   1200
            TabIndex        =   15
            Top             =   720
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3300
         TabIndex        =   4
         ToolTipText     =   "Register Checked Discovered Soft Points"
         Top             =   60
         Width           =   1290
      End
      Begin VB.CommandButton cmdParts 
         Caption         =   "Partitions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5025
         TabIndex        =   5
         ToolTipText     =   "Register Checked Discovered Soft Points"
         Top             =   60
         Width           =   1290
      End
      Begin VB.Timer tmrSearchPartitions 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   9360
         Top             =   4260
      End
      Begin VB.Timer tmrSearchSP 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   9420
         Top             =   3720
      End
      Begin VB.CommandButton cmdUnRegister 
         Caption         =   "Unregister"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         TabIndex        =   9
         ToolTipText     =   "Unregister Soft Point"
         Top             =   60
         Width           =   1290
      End
      Begin VB.TextBox txtSearchPartitions 
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
         Left            =   4305
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Type Search String"
         Top             =   427
         Width           =   1770
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
         Left            =   7245
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Type Search String"
         Top             =   427
         Width           =   2370
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1965
         TabIndex        =   2
         ToolTipText     =   "Clear Discoverd List"
         Top             =   60
         Width           =   1290
      End
      Begin VB.Timer tmrFetch 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   9480
         Top             =   2880
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Discover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Start / Stop Discovery"
         Top             =   60
         Width           =   1290
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
         Height          =   345
         Left            =   8325
         TabIndex        =   13
         ToolTipText     =   "Exit Screen"
         Top             =   60
         Width           =   1290
      End
      Begin MSComctlLib.ListView lvSP 
         Height          =   2535
         Left            =   6360
         TabIndex        =   12
         ToolTipText     =   "Active Soft Points"
         Top             =   780
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   20
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Desc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvDisc 
         Height          =   2535
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "List of New Discoverd Soft Points (Unregistered)"
         Top             =   780
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Received"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID1"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rssi1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rssi2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Rssi3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ID4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Rssi4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ID5"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Rssi5"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ID6"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Rssi6"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ID7"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Rssi7"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "ID8"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Rssi8"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvAvailPartitions 
         Height          =   2535
         Left            =   3300
         TabIndex        =   8
         ToolTipText     =   "Select AvailablePartition"
         Top             =   780
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "id"
            Text            =   "ID"
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "desc"
            Text            =   "Desc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "loc"
            Text            =   "Loc"
            Object.Width           =   1129
         EndProperty
      End
      Begin VB.Label Label2 
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
         Left            =   3540
         TabIndex        =   6
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Left            =   6480
         TabIndex        =   10
         Top             =   495
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSoftPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ManagingPartitions As Boolean

Private mDiscovering As Boolean
Private Messages    As Collection

Private Discovered  As Collection
Private SoftPoints  As Collection
Private Partitions  As Collection


'Const CHTime = 1

Const CHID1 = 1
Const CHID2 = 3
Const CHID3 = 5
Const CHID4 = 7
Const CHID5 = 9
Const CHID6 = 11
Const CHID7 = 13
Const CHID8 = 15


Const CHR1 = 2
Const CHR2 = 4
Const CHR3 = 6
Const CHR4 = 8
Const CHR5 = 10
Const CHR6 = 12
Const CHR7 = 14
Const CHR8 = 16

Const IDWidth = 800
Const RWidth = 500

Private wsSP        As WebSocketSocket

Private Partitionlist As Collection




Private Sub cmdBegin_Click()
  If Discovering Then
    EndDiscovering
  Else
    BeginDiscovering
  End If
End Sub

Private Sub cmdCancel_Click()
  tmrFetch.Enabled = False
  PreviousForm
  Unload Me
End Sub

Private Sub cmdClear_Click()
  Set Discovered = New Collection
  FillDiscovered

End Sub

Private Sub cmdParts_Click()
    ManageAvailablePartitions
  ManagingPartitions = True
End Sub

Private Sub cmdRegister_Click()
  EndDiscovering
  RegisterSelected
End Sub

Sub RegisterSelected()
  Dim li            As ListItem
  Dim lis           As ListItem
  Dim PartID        As Long
  Dim sp            As cSoftPoint
  Dim RegisterItems As Collection
  Set RegisterItems = New Collection


  For Each li In lvAvailPartitions.ListItems
    If li.Checked Then
      PartID = Val(li.text)
      Exit For
    End If
  Next
  If PartID = 0 Then
    Beep
    Exit Sub
  Else

    For Each li In lvDisc.ListItems
      If li.Checked Then
        Set sp = New cSoftPoint
        sp.TimeString = li.text
        sp.PartitionID = PartID
        sp.DeviceID1 = Val(li.SubItems(CHID1))
        sp.Rssi1 = Val(li.SubItems(CHR1))
        sp.DeviceID2 = Val(li.SubItems(CHID2))
        sp.Rssi2 = Val(li.SubItems(CHR2))
        sp.DeviceID3 = Val(li.SubItems(CHID3))
        sp.Rssi3 = Val(li.SubItems(CHR3))
        sp.DeviceID4 = Val(li.SubItems(CHID4))
        sp.Rssi4 = Val(li.SubItems(CHR4))
        sp.DeviceID5 = Val(li.SubItems(CHID5))
        sp.Rssi5 = Val(li.SubItems(CHR5))
        sp.DeviceID6 = Val(li.SubItems(CHID6))
        sp.Rssi6 = Val(li.SubItems(CHR6))
        sp.DeviceID7 = Val(li.SubItems(CHID7))
        sp.Rssi7 = Val(li.SubItems(CHR7))
        sp.DeviceID8 = Val(li.SubItems(CHID8))
        sp.Rssi8 = Val(li.SubItems(CHR8))
        RegisterItems.Add sp

      End If
    Next

  End If

  Dim j             As Long
  Dim k As Long
  For j = RegisterItems.Count To 1 Step -1
    Set sp = RegisterItems(j)
    ' do register
    RegisterSP sp
    If sp.ID Then
      RegisterItems.Remove j
      SoftPoints.Add sp
      ' remove from Discovered
      For k = 1 To Discovered.Count
        If Discovered(k).TimeString = sp.TimeString Then
          Discovered.Remove k
          Exit For
        End If
      Next
      
      
    End If
  Next

  FillSoftPoints
  FillDiscovered

End Sub


Private Sub cmdUnRegister_Click()
  DoUnregister
End Sub
Sub DoUnregister()
  Dim j             As Long
  Dim i             As Long

  Dim ID            As Long
  Dim li            As ListItem
  For j = lvSP.ListItems.Count To 1 Step -1
    Set li = lvSP.ListItems(j)
    If li.Checked Then
      ID = Val(li.text)
      If UnregisterSP(ID) Then
        For i = 1 To SoftPoints.Count
          If SoftPoints(i).ID = ID Then
            SoftPoints.Remove i
            Exit For
          End If
        Next
      End If
    End If
  Next

  FillSoftPoints

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    KeyCode = 0
    RefreshPartsAndSPs
  End If
End Sub
Sub RefreshPartsAndSPs()
  RefreshSPs

End Sub

Sub RefreshSPs()

End Sub

Private Sub Form_Load()
ResetActivityTime
  If wsSP Is Nothing Then
    Set wsSP = New WebSocketSocket
  End If
  Set Messages = New Collection
  Set SoftPoints = New Collection
  Set Discovered = New Collection
  Set Partitions = New Collection
  SetLVColumns
  'GetPartitionList

End Sub

Private Function GetPartitionList()

  Dim HTTPRequest   As cHTTPRequest
  Dim XML           As String
  
  Set HTTPRequest = New cHTTPRequest
  XML = HTTPRequest.GetPartitionList(GetHTTP & "://" & IP1, USER1, PW1)
  If Len(XML) Then
    Set Partitions = ParsePartionList(XML)
  Else
    Set Partitions = New Collection
  End If
  Set HTTPRequest = Nothing


End Function

Private Sub SetLVColumns()
  Dim i             As Long
  i = 1
  lvDisc.ColumnHeaders(i).Width = 1400
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID1"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R1"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID2"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R2"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID3"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R3"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID4"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R4"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID5"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R5"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID6"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R6"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID7"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R7"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = IDWidth
  lvDisc.ColumnHeaders(i).text = "ID8"
  i = i + 1
  lvDisc.ColumnHeaders(i).Width = RWidth
  lvDisc.ColumnHeaders(i).text = "R8"


  i = 1

  Me.lvAvailPartitions.ColumnHeaders(i).Width = 1400


  i = 1

  lvSP.ColumnHeaders(i).Width = 1400
  i = i + 1

  Me.lvSP.ColumnHeaders(i).Width = 1200
  lvSP.ColumnHeaders(i).text = "Desc"
  i = i + 1

  Me.lvSP.ColumnHeaders(i).Width = 1200
  lvSP.ColumnHeaders(i).text = "PID"
  i = i + 1

  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID1"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R1"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID2"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R2"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID3"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R3"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID4"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R4"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID5"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R5"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID6"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R6"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID7"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R7"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = IDWidth
  lvSP.ColumnHeaders(i).text = "ID8"
  i = i + 1
  lvSP.ColumnHeaders(i).Width = RWidth
  lvSP.ColumnHeaders(i).text = "R8"


End Sub

Private Sub Form_Unload(Cancel As Integer)

  tmrFetch.Enabled = False
  Discovering = False
  If Not wsSP Is Nothing Then
    wsSP.DisConnect
  End If
  Set wsSP = Nothing

End Sub
Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  fraEnabler.BackColor = Me.BackColor
  SetParent fraEnabler.hwnd, hwnd
  
  Me.fraLoading.Visible = True
  Me.fraLoading.ZOrder
  
  Me.Refresh
  Me.fraEnabler.Enabled = False
  GetPartitionList
  Me.fraEnabler.Enabled = True
  Me.fraLoading.Visible = False
End Sub



Public Sub UnHost()
  Discovering = False
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Sub Fill()
  
  
  Me.fraLoading.Visible = True
  Me.fraLoading.ZOrder
  
  Me.fraEnabler.Enabled = False
    
  
  If ManagingPartitions Then
    ManagingPartitions = False
    GetPartitionList
  End If
  FetchSoftPoints
  FillSoftPoints
  FillDiscovered
  
  FillPartitions

  Me.fraLoading.Visible = False
  Me.fraEnabler.Enabled = True

End Sub
Sub FillPartitions()
  Dim part          As cPartition
  Dim Found         As Boolean
  Dim li            As ListItem

  lvAvailPartitions.ListItems.Clear

  For Each part In Partitions
    Found = False

    If Len(Trim$(Me.txtSearchPartitions.text)) > 0 Then
      Found = InStr(1, part.Description, txtSearchPartitions.text, vbTextCompare) <> 0
      Found = Found And (part.IsLocation <> 0)
    Else
      Found = True And (part.IsLocation <> 0)
    End If
    If Found Then
      Set li = lvAvailPartitions.ListItems.Add(, , part.PartitionID)
      li.SubItems(1) = part.Description
      li.SubItems(2) = part.IsLocation
    End If
  Next

End Sub

Sub ClearDiscovered()
  Set Discovered = New Collection
  FillDiscovered
End Sub
Sub FillDiscovered()
  Dim sp            As cSoftPoint
  Dim li            As ListItem
  Dim i             As Long
  lvDisc.ListItems.Clear
  For Each sp In Discovered
    Set li = lvDisc.ListItems.Add(, , sp.TimeString)
    i = 1
    li.SubItems(i) = Val(sp.DeviceID1)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi1)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID2)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi2)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID3)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi3)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID4)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi4)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID5)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi5)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID6)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi6)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID7)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi7)
    i = i + 1
    li.SubItems(i) = Val(sp.DeviceID8)
    i = i + 1
    li.SubItems(i) = Val(sp.Rssi8)
  Next


End Sub

Sub EndDiscovering()
  tmrFetch.Enabled = False
  If Not wsSP Is Nothing Then
    wsSP.DisConnect
  End If
  Discovering = False
End Sub
Sub BeginDiscovering()
  Dim URL           As String
  Dim rc            As Long

  If Me.lvDisc.ListItems.Count Then
    rc = messagebox(Me, "Clear Discoverd List?", App.Title, vbYesNo Or vbQuestion)
  End If
  If rc = vbYes Then
    ClearDiscovered
  End If


  If wsSP Is Nothing Then
    Set wsSP = New WebSocketSocket
  Else
    Debug.Print "BeginDiscovering", wsSP.StatusCode
    If wsSP.StatusCode = 1 Then
      wsSP.DisConnect
    End If
  End If
  wsSP.Init "5002EgAtIrEh"
  Call wsSP.UserNamePassword(USER1, PW1)
  wsSP.SetURL GetWS & "://" & IP1 & "/PSIA/Metadata/stream?SoftPoint=true"

  wsSP.Connect
  Sleep 200
  Do While wsSP.HasMessages  ' clear out any garbage
    wsSP.GetNextMessage
  Loop


  Discovering = True
  '  Debug.Print "Status, code,info ", wsSP.StatusCode, wsSP.StatusString
  tmrFetch.Enabled = True



End Sub
Sub FillSoftPoints()
  Dim sp            As cSoftPoint
  Dim part          As cPartition
  'Dim partDesc As String
  Dim Found         As Boolean
  Dim desc          As String
  Dim i             As Long

  Dim li            As ListItem
  lvSP.ListItems.Clear
  For Each sp In SoftPoints

    Found = False
    desc = ""
    On Error Resume Next

    Set part = Nothing
    Set part = Partitions(sp.PartitionID & "")
    If Not part Is Nothing Then
      desc = part.Description
    End If


    If Len(Me.txtSearchBox.text) Then
      Found = InStr(1, desc, txtSearchBox.text, vbTextCompare) <> 0
    Else
      Found = True
    End If
    If Found Then
      Set li = lvSP.ListItems.Add(, , sp.ID)
      li.SubItems(1) = desc
      li.SubItems(2) = sp.PartitionID
      i = 3
      li.SubItems(i) = Val(sp.DeviceID1)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi1)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID2)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi2)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID3)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi3)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID4)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi4)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID5)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi5)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID6)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi6)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID7)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi7)
      i = i + 1
      li.SubItems(i) = Val(sp.DeviceID8)
      i = i + 1
      li.SubItems(i) = Val(sp.Rssi8)

    End If


  Next



End Sub


Sub FetchSoftPoints()


  Dim HTTPRequest   As cHTTPRequest
  Dim Found         As Boolean
  Dim XML           As String
  


  Dim sp            As cSoftPoint
  Set SoftPoints = New Collection

  


  Set HTTPRequest = New cHTTPRequest
  XML = HTTPRequest.GetSoftPointList(GetHTTP & "://" & IP1, USER1, PW1)

  If Len(XML) Then
    Set SoftPoints = ParseSoftPointList(XML)
  Else
    Set SoftPoints = New Collection
  End If








End Sub


Private Sub lvAvailPartitions_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim li            As ListItem
  For Each li In lvAvailPartitions.ListItems
    If li Is Item Then
    Else
      li.Checked = False
    End If
  Next


End Sub

Private Sub tmrFetch_Timer()
  Dim XML           As String
  If Not wsSP Is Nothing Then
    'Debug.Print "tmrFetch wsSP.GetLastError", wsSP.GetLastError
    'Debug.Print "tmrFetch wsSP.StatusCode", wsSP.StatusCode
    If wsSP.HasMessages Then
      Debug.Print "tmrFetch wsSP.HasMessages", wsSP.HasMessages
      XML = wsSP.GetNextMessage

      AddSP XML
      
      FillDiscovered
    End If
  Else
    tmrFetch.Enabled = False
    EndDiscovering
    ' quit
  End If
End Sub
Sub AddSP(ByVal XML As String)
  Dim sp            As cSoftPoint
  Set sp = New cSoftPoint
  If sp.ParseEventXML(XML) Then
    sp.Num = Discovered.Count
    Discovered.Add sp
    If gForwardSoftPoints And Len(gSPForwardAccount) > 0 Then
      ForwardToSMS sp, gSPForwardAccount
    End If
  End If

End Sub

Sub ForwardToSMS(sp As cSoftPoint, ByVal recipient As String)
          
          Dim mapi    As Object
          Dim message As String
          Dim Subject As String
          Dim filename As String ' none!
          

10        Subject = "Soft Point " & sp.TimeString
20        message = ":" & vbCr & "ID1 " & Right$("    " & sp.DeviceID1, 8) & " " & sp.Rssi1 & vbCr
30        message = message & "ID2 " & Right$("        " & sp.DeviceID2, 8) & " " & Right$("  " & sp.Rssi2, 2) & vbCr
40        message = message & "ID3 " & Right$("        " & sp.DeviceID3, 8) & " " & Right$("  " & sp.Rssi3, 2) & vbCr
50        message = message & "ID4 " & Right$("        " & sp.DeviceID4, 8) & " " & Right$("  " & sp.Rssi4, 2) & vbCr
60        message = message & "ID5 " & Right$("        " & sp.DeviceID5, 8) & " " & Right$("  " & sp.Rssi5, 2) & vbCr
70        message = message & "ID6 " & Right$("        " & sp.DeviceID6, 8) & " " & Right$("  " & sp.Rssi6, 2) & vbCr
80        message = message & "ID7 " & Right$("        " & sp.DeviceID7, 8) & " " & Right$("  " & sp.Rssi7, 2) & vbCr
90        message = message & "ID8 " & Right$("        " & sp.DeviceID8, 8) & " " & Right$("  " & sp.Rssi8, 2) & vbCr
          

100       On Error Resume Next

110       If (Configuration.UseSMTP = MAIL_SMTP) Then

120         If gSMTPMailer Is Nothing Then
130           Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
140         End If
150         If gSMTPMailer Is Nothing Then
160           LogProgramError "Could not create SMTPMailer Object in frmSoftPoints.ForwardToSMS." & Erl
170         Else
180           Call gSMTPMailer.Send("", "", recipient, Subject, message, "")
190         End If


200       Else
210         Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
220         If mapi Is Nothing Then
230           LogProgramError "Could not create SENTRYMAIL Object in frmSoftPoints.ForwardToSMS." & Erl
240         Else
250           Call mapi.Send("", "", recipient, Subject, message)
260         End If

270       End If


          '// Username, Password, Address, Subject,Body, AttachmentsList ' Attachemnet list is a semicolon ";" delimited list of file attachments
          'Call mapi.Send("", "", Configuration.AssurEmailRecipient, Configuration.AssurEmailSubject, Message)

280       Set mapi = Nothing


End Sub

Public Property Get Discovering() As Boolean

  Discovering = mDiscovering

End Property

Public Property Let Discovering(ByVal Value As Boolean)
  If Value Then
    cmdBegin.Caption = "Stop"
    cmdRegister.Enabled = False
    cmdClear.Enabled = False
    cmdUnRegister.Enabled = False
    txtSearchPartitions.Enabled = False
    txtSearchBox.Enabled = False
    cmdParts.Enabled = False

  Else
    Me.cmdBegin.Caption = "Discover"
    cmdClear.Enabled = True
    cmdRegister.Enabled = True
    cmdUnRegister.Enabled = True
    txtSearchPartitions.Enabled = True
    txtSearchBox.Enabled = True
    cmdParts.Enabled = True
  End If
  mDiscovering = Value

End Property

Private Sub tmrSearchPartitions_Timer()
  tmrSearchPartitions.Enabled = False
  FillPartitions
End Sub

Private Sub tmrSearchSP_Timer()
  tmrSearchSP.Enabled = False
  FillSoftPoints
  'FillSPs
End Sub

Private Sub txtSearchBox_Change()

  tmrSearchSP.Enabled = False
  tmrSearchSP.Enabled = True
End Sub

Private Sub txtSearchPartitions_Change()
  tmrSearchPartitions.Enabled = False
  tmrSearchPartitions.Enabled = True


End Sub

