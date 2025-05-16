VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPartitions 
   Caption         =   "Partitions"
   ClientHeight    =   3165
   ClientLeft      =   405
   ClientTop       =   2940
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   9045
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Frame fraLoading 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   2160
         TabIndex        =   13
         Top             =   720
         Width           =   1635
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
            Left            =   300
            TabIndex        =   14
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdCreateFromRooms 
         Caption         =   "Create from Rooms (no!)"
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
         Left            =   5880
         TabIndex        =   12
         Top             =   2460
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton cmdUnassign 
         Caption         =   "Unassign"
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
         TabIndex        =   9
         Top             =   1290
         Width           =   1175
      End
      Begin VB.Timer tmrSearch 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6540
         Top             =   1500
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
         TabIndex        =   5
         ToolTipText     =   "Type Search String"
         Top             =   600
         Width           =   2370
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
         Height          =   525
         Left            =   7680
         TabIndex        =   8
         Top             =   705
         Width           =   1175
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Update"
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
         TabIndex        =   7
         Top             =   120
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
         Height          =   525
         Left            =   7680
         TabIndex        =   10
         Top             =   1875
         Width           =   1175
      End
      Begin VB.CheckBox chkLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "Location"
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
         Left            =   4080
         TabIndex        =   3
         ToolTipText     =   "Check this Box to Create  Location Partition"
         Top             =   270
         Width           =   1275
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
         TabIndex        =   11
         Top             =   2460
         Width           =   1175
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Description of Partition"
         Top             =   285
         Width           =   2670
      End
      Begin MSComctlLib.ListView lvAvailPartitions 
         Height          =   2115
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "id"
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "desc"
            Text            =   "Description"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "loc"
            Text            =   "Location"
            Object.Width           =   1852
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mange Partitions"
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
         Left            =   105
         TabIndex        =   15
         Top             =   30
         Width           =   1440
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
         Left            =   2520
         TabIndex        =   4
         Top             =   660
         Width           =   615
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPartitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const EditNone = 0
Const EditActive = 1


Private CurrentPartition As cPartition
Private mEditmode As Long
Private LastIndex   As Long

Private HTTPRequest As cHTTPRequest

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  fraEnabler.BackColor = Me.BackColor
  SetParent fraEnabler.hwnd, hwnd
  Me.Refresh
End Sub
Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Private Sub cmdDelete_Click()
  DeleteSelectedPartitions
End Sub
Function DeleteSelectedPartitions()

  Dim response      As String
  Dim li As ListItem
  Dim j As Long

  Dim XML           As String
  Dim Success       As Boolean

  Dim HTTPRequest   As cHTTPRequest
  Set HTTPRequest = New cHTTPRequest
  
  'For Each li In lvAvailPartitions.ListItems
  For j = 1 To lvAvailPartitions.ListItems.Count
    Set li = lvAvailPartitions.ListItems(j)
    If li.Selected Then
      response = HTTPRequest.DeletePartition(GetHTTP & "://" & IP1, USER1, PW1, Val(li.text))
      LastIndex = j
    End If
  Next
  Set HTTPRequest = Nothing

  Fill
  
  If LastIndex <= lvAvailPartitions.ListItems.Count Then
    Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(LastIndex)
    lvAvailPartitions.SelectedItem.EnsureVisible
  Else
    If lvAvailPartitions.ListItems.Count Then
      Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(lvAvailPartitions.ListItems.Count)
      lvAvailPartitions.SelectedItem.EnsureVisible
    End If
  End If


  
  
    
  
End Function

Function UnassignSelectedPartition()
  
  Dim response      As String
  Dim li As ListItem
  Dim j As Long

    
  
  Dim XML           As String
  Dim Success       As Boolean

  Dim HTTPRequest   As cHTTPRequest
  Set HTTPRequest = New cHTTPRequest
  
  For j = 1 To lvAvailPartitions.ListItems.Count
    Set li = lvAvailPartitions.ListItems(j)
    If li.Selected Then
      LastIndex = j
      Call HTTPRequest.GetZonesForPartition(GetHTTP & "://" & IP1, USER1, PW1, Val(li.text))
      Exit For
    End If
  Next
  

  Fill
  
  
  If LastIndex <= lvAvailPartitions.ListItems.Count Then
    Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(LastIndex)
    lvAvailPartitions.SelectedItem.EnsureVisible
  Else
    If lvAvailPartitions.ListItems.Count Then
      Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(lvAvailPartitions.ListItems.Count)
      lvAvailPartitions.SelectedItem.EnsureVisible
    End If
  End If

End Function


Private Sub cmdEdit_Click()
  Dim li            As ListItem
  Dim j             As Long
  If Editmode = EditNone Then
    Editmode = EditActive
    'For Each li In lvAvailPartitions.ListItems
    For j = 1 To lvAvailPartitions.ListItems.Count
      Set li = lvAvailPartitions.ListItems(j)
      If li.Selected Then
        LastIndex = j
        Set CurrentPartition = New cPartition
        CurrentPartition.PartitionID = Val(li.text)
        CurrentPartition.Description = li.SubItems(1)
        CurrentPartition.IsLocation = IIf(li.SubItems(2) = "X", 1, 0)
        Editmode = EditActive
        txtDescription.text = CurrentPartition.Description
        chkLocation.Value = IIf(CurrentPartition.IsLocation, 1, 0)

        Exit For
      End If

    Next
  Else
    Editmode = EditNone
  End If


End Sub


Private Sub cmdNew_Click()
  Dim Description   As String
  Description = Trim$(txtDescription.text)
  If Len(Description) Then
    If Editmode = EditActive Then
      ' update on 6080 and locally
      CurrentPartition.Description = Description
      CurrentPartition.IsLocation = (chkLocation.Value And 1)
      UpdatePartition CurrentPartition
      Editmode = EditNone
      Fill
      If LastIndex <= lvAvailPartitions.ListItems.Count Then
        Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(LastIndex)
        lvAvailPartitions.SelectedItem.EnsureVisible
      Else
        If lvAvailPartitions.ListItems.Count Then
          Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(lvAvailPartitions.ListItems.Count)
          lvAvailPartitions.SelectedItem.EnsureVisible
        End If
      End If
    Else  ' add it
      Set CurrentPartition = New cPartition
      CurrentPartition.Description = Description
      CurrentPartition.IsLocation = chkLocation.Value And 1
      CreatePartition CurrentPartition
      ' update on 6080 and locally
    End If
  End If
End Sub
Function CreatePartition(partition As cPartition) As Long
  
  Dim response      As String
  Dim XML           As String
  Dim Success       As Boolean
  Dim HTTPRequest   As cHTTPRequest
  
  Set HTTPRequest = New cHTTPRequest
  response = HTTPRequest.CreatePartition(GetHTTP & "://" & IP1, USER1, PW1, partition)
  Set HTTPRequest = Nothing

  Dim doc           As DOMDocument60
  Dim Node          As IXMLDOMNode
  Dim StatusCode    As Long
  Dim StatusString  As Long
  Dim NewID         As Long

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

  CreatePartition = NewID

  '<ResponseStatus version="1.0" xmlns:urn="psialliance-org">
  '<requestURL>/PSIA/AreaControl/Partitions/PartitionInfoList/partitionInfo</requestURL>
  '<statusCode>1</statusCode>
  '<statusString>201 Created</statusString>
  '<id>3</id>
  '</ResponseStatus>
  Fill
  
  If lvAvailPartitions.ListItems.Count Then
      Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(lvAvailPartitions.ListItems.Count)
      lvAvailPartitions.SelectedItem.EnsureVisible
   End If

  

End Function
Function UpdatePartition(partition As cPartition) As Long
  
  Dim XML           As String
  Dim Success       As Boolean
  
  Dim HTTPRequest As cHTTPRequest
  Set HTTPRequest = New cHTTPRequest
  UpdatePartition = HTTPRequest.UpdatePartition(GetHTTP & "://" & IP1, USER1, PW1, partition)
  Set HTTPRequest = Nothing
  
  Editmode = EditNone
  
  Fill

  If LastIndex <= lvAvailPartitions.ListItems.Count Then
    Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(LastIndex)
    lvAvailPartitions.SelectedItem.EnsureVisible
  Else
    If lvAvailPartitions.ListItems.Count Then
      Set lvAvailPartitions.SelectedItem = lvAvailPartitions.ListItems(lvAvailPartitions.ListItems.Count)
      lvAvailPartitions.SelectedItem.EnsureVisible
    End If
  End If



End Function

Sub Fill()
  Dim part          As cPartition
  Dim li            As ListItem
  Dim Found         As Boolean
  Dim XML           As String

  Dim Partitions      As Collection
  
  Me.fraLoading.Visible = True
  Me.fraLoading.ZOrder
  Me.fraEnabler.Enabled = False
  'Screen.MousePointer = vbHourglass
  
  Me.Refresh

  Set HTTPRequest = New cHTTPRequest
  XML = HTTPRequest.GetPartitionList(GetHTTP & "://" & IP1, USER1, PW1)
  
  If Len(XML) Then
    Set Partitions = ParsePartionList(XML)
  Else
    Set Partitions = New Collection
  End If
  
  lvAvailPartitions.ListItems.Clear

  For Each part In Partitions
    If Len(Trim$(Me.txtSearchBox.text)) > 0 Then
      Found = InStr(1, part.Description, Me.txtSearchBox.text, vbTextCompare) <> 0
    Else
      Found = True
    End If
    If Found Then
      Set li = lvAvailPartitions.ListItems.Add(, , part.PartitionID)
      li.SubItems(1) = part.Description
      li.SubItems(2) = IIf(part.IsLocation, "X", "")
    End If

  Next

  For Each li In lvAvailPartitions.ListItems
    li.Selected = 0
  Next
  'Screen.MousePointer = vbNormal
  Me.fraLoading.Visible = False
  Editmode = EditNone
  Me.fraEnabler.Enabled = True
End Sub


Private Sub cmdUnassign_Click()
  
  UnassignSelectedPartition
  
End Sub

Private Sub cmdCreateFromRooms_Click()
  CreatePartitonsFromRooms
End Sub

Private Sub Form_Load()
ResetActivityTime
  Editmode = EditNone
  fraEnabler.BackColor = Me.BackColor
  
  
End Sub
Sub setcolumns()
  Me.lvAvailPartitions.ColumnHeaders(1).Width = 800
  Me.lvAvailPartitions.ColumnHeaders(2).Width = 2450
  Me.lvAvailPartitions.ColumnHeaders(3).Width = 1050
End Sub

Private Sub tmrSearch_Timer()
  tmrSearch.Enabled = False
  Fill
End Sub

Private Sub txtSearchBox_Change()
  tmrSearch.Enabled = False
  tmrSearch.Enabled = True
  
End Sub

Public Property Get Editmode() As Long

  Editmode = mEditmode

End Property

Public Property Let Editmode(ByVal Value As Long)
  
  cmdNew.Caption = IIf(Value = 1, "Update", "Add")
  cmdEdit.Caption = IIf(Value = 1, "Cancel", "Edit")
  mEditmode = Value

End Property
