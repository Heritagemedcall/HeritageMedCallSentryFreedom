VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmWaypoints 
   Caption         =   "Waypoints"
   ClientHeight    =   3165
   ClientLeft      =   2760
   ClientTop       =   4275
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   9105
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdLocate 
         Caption         =   "Locate"
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
         Width           =   1155
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
         Width           =   1155
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
         Width           =   1155
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
         Width           =   1155
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
         TabIndex        =   1
         Top             =   2370
         Width           =   1155
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmWaypoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLocating     As Boolean
'Private SurveyDevice  As cESSurveyDevice
Private SelectedPtr   As Long
Private Selected      As Collection
Private waypoint      As cWayPoint

Private SortOrder     As Integer

Private Sub cmdAdd_Click()
  EditWaypoint 0


End Sub

Public Sub ProcessPacket(packet As cESPacket)

  ' From ModES.BatchSurvey
  
  ' at this point if configuration.noncs is true and first hop was NC, then it should not get here

  Dim Serial As String
  ' this is were we handle events from the system
  If SurveyDevice Is Nothing Then Exit Sub
    
  If gDirectedNetwork Then
    If packet.IsLocatorPacket Then
      Serial = packet.LocatedSerial
    Else
      Serial = packet.Serial
    End If
  Else
    Serial = packet.Serial
  End If
    
  If SurveyDevice.Serial = Serial Then
    If gDirectedNetwork Then
      If Configuration.OnlyLocators Then ' NOTE: NOT using in new installs
        If packet.IsLocatorPacket Then
           SurveyDevice.AddLocater packet
        End If
      Else
        If packet.IsLocatorPacket Then ' NOT using in new installs
          SurveyDevice.AddLocater packet
        ElseIf packet.Alarm0 = 1 Then
          SurveyDevice.AddLocater packet
        End If
      End If
    ElseIf packet.Alarm0 = 1 Then ' NOT Directed Network
      SurveyDevice.AddLocater packet
    End If
  End If
    
  If SurveyDevice.PCASerial = packet.Serial Then
    If packet.CMD = &H11 Then
      ' skip it
    Else
      
      Select Case SurveyDevice.ResponseCode(packet, Configuration.surveymode)
        Case 0  ' nothing... supervisory
          'Debug.Assert 0
          Debug.Print "nothing... supervisory MULTI"
        Case 1  ' OK
          Debug.Print "SurveyDevice.ProcessLocations MULTI"
          SurveyDevice.ProcessLocations
          waypoint.Repeater1 = SurveyDevice.Location1
          waypoint.Repeater2 = SurveyDevice.Location2
          waypoint.Repeater3 = SurveyDevice.Location3
          waypoint.Signal1 = SurveyDevice.Signal1
          waypoint.Signal2 = SurveyDevice.Signal2
          waypoint.Signal3 = SurveyDevice.Signal3
          ' update screen??

          dbgloc "Assigning Repeaters and Levels: " & waypoint.Description & vbCrLf
          UpdateWaypoint waypoint
          NextWayPoint

        Case 2  ' cancel this one
          Debug.Print "Waypoint Update Cancelled MULTI"
          dbgloc "Waypoint Update Cancelled: " & waypoint.Description & vbCrLf
          NextWayPoint
        Case 3  ' quit
          Debug.Print "Quitting Waypoint Updates MULTI"
          dbgloc "Quitting Waypoint Updates: " & waypoint.Description & vbCrLf
          Locating = False
          ShowWaypoints

      End Select
    End If
  End If
End Sub


Private Sub cmdDelete_Click()
  Dim ID    As Long
  Dim li    As ListItem
  Dim SQL   As String

  If lvMain.SelectedItem Is Nothing Then
    ' nada
    Fill
  Else
    Set li = lvMain.SelectedItem
    ID = Val(li.Key)
    SQL = "DELETE FROM Waypoints WHERE ID = " & ID
    ConnExecute SQL
    Fill
  End If


End Sub

Private Sub cmdEdit_Click()
  Dim li As ListItem
  Dim ID As Long
  If lvMain.SelectedItem Is Nothing Then
    EditWaypoint 0
  Else
    Set li = lvMain.SelectedItem
    ID = Val(li.Key)
    EditWaypoint ID
  End If

End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Sub Configurelvmain()
  Dim ch As ColumnHeader

  If lvMain.ColumnHeaders.Count < 6 Then


    lvMain.ColumnHeaders.Clear
    'lvMain.Sorted = True
    'Set ch = lvMain.ColumnHeaders.Add(, "Serial", "Serial", 1100)
    Set ch = lvMain.ColumnHeaders.Add(, "Desc", "Description", 1440)
    Set ch = lvMain.ColumnHeaders.Add(, "Bldg", "Building", 1350)
    Set ch = lvMain.ColumnHeaders.Add(, "Floor", "Floor", 1440)
    Set ch = lvMain.ColumnHeaders.Add(, "Wing", "Wing", 1440)
    Set ch = lvMain.ColumnHeaders.Add(, "Levels", "Levels", 1440)
    'Set ch = lvMain.ColumnHeaders.Add(, "Rpt", "Repeaters", 700)


  End If
End Sub


Public Property Get Locating() As Boolean
  Locating = mLocating
End Property
Public Property Let Locating(ByVal Value As Boolean)

  BatchSurveyEnabled = Value
  If Value Then
    cmdLocate.Caption = "Cancel"
  Else
    Set BatchForm = Nothing
    cmdLocate.Caption = "Locate"
  End If
  mLocating = Value

End Property

Private Sub cmdLocate_Click()
  If mLocating Then
    Locating = False
  Else
    'Configuration.PCARedirect = 0
    AutoLocate
  End If
End Sub



Sub AutoLocate()

  Dim li As ListItem
  Set Selected = New Collection
  For Each li In lvMain.ListItems
    If li.Selected Then
      Selected.Add li
      Debug.Print li.text
    End If
  Next
  If Selected.Count > 0 Then
    If Configuration.surveymode = PCA_MODE Then
      Outbounds.AddMessage Configuration.SurveyPCA, MSGTYPE_TWOWAYNID, "", 0 ' surveypca is string set NID
    ElseIf Configuration.surveymode = TWO_BUTTON_MODE Then
      'SendToPager "Begin Survey.", Configuration.SurveyPager, 0, "", "", PAGER_NORMAL, "", 0
      'Outbounds.AddMessage Configuration.SurveyPCA, MSGTYPE_TWOWAYNID, "", 0 ' surveypca is string
    ElseIf Configuration.surveymode = EN1221_MODE Then
      
    End If
    dbgloc "FrmWaypoints.Autolocate Count: " & Selected.Count  ' how many to do
    
    Set BatchForm = Me ' batchform is this form!
    SelectedPtr = 0
    NextWayPoint
    Locating = True
    BatchSurveyEnabled = True
  End If

End Sub
Sub NextWayPoint()


  Dim ID                 As Long
  Dim li                 As ListItem
  Dim j                  As Integer

  Dim w                  As cWayPoint
  Dim WaypointDescription As String
  Dim NewMessage         As cWirelessMessage

  Dim PagerID            As Long

  On Error GoTo NextWayPoint_Error

  Sleep 1

  SelectedPtr = SelectedPtr + 1
  If SelectedPtr > Selected.Count Then
    Locating = False
    ShowWaypoints
  Else

    Set li = Selected(SelectedPtr)
    ID = Val(li.Key)

    For j = 1 To Waypoints.Count
      Set w = Waypoints.waypoint(j)

      If w.ID = ID Then  ' it's one of the selected waypoints, use it

        dbgloc "frmWaypoints.Nextwaypoint (in loop) " & ID & " " & w.Description
        Set waypoint = w  'waypoint is form level object

        Set SurveyDevice = New cESSurveyDevice  'SurveyDevice is now global level object

        If Configuration.surveymode = PCA_MODE Then  'PCA mode

          ' set up serials
          SurveyDevice.Serial = Configuration.SurveyDevice
          SurveyDevice.PCASerial = Configuration.SurveyPCA

          ' make the prompt that shows on the PCA

          SurveyDevice.RequestSurvey BuildPrompt(waypoint), 3

          Set NewMessage = New cWirelessMessage
          NewMessage.MessageData = SurveyDevice.MessageString
          NewMessage.TimeStamp = DateAdd("s", -1, Now)  ' send right away, no delay
          NewMessage.NeedsAck = SurveyDevice.RequireACK  ' (Val("&h" & MID(mMessageData, 17, 2)) And (BIT_7)) = (BIT_7)
          NewMessage.SequenceID = SurveyDevice.Sequence  ' Val("&h" & MID(mMessageData, 21, 4))

          'Send it to the PCA
          Outbounds.AddPreparedMessage NewMessage
          Set NewMessage = Nothing

          'dbg "Sent PCA Message, Get " & s & vbCrLf

        ElseIf Configuration.surveymode = EN1221_MODE Then  ' New tiny pendant
          PagerID = Configuration.SurveyPager
          SurveyDevice.Serial = Configuration.SurveyDevice
          If PagerID <> 0 Then
            SurveyDevice.PagerID = PagerID
            SurveyDevice.Serial = Configuration.SurveyDevice  ' use just one device
            SurveyDevice.PCASerial = Configuration.SurveyDevice
            ' create outbound message for ouput "Pager"
            SurveyDevice.RequestSurvey BuildPrompt(waypoint), 2, True  ' only OK, cancel
            dbgloc "Sending message to pager " & SurveyDevice.MessageString
            SendToPager SurveyDevice.MessageString, SurveyDevice.PagerID, 0, "", "", PAGER_NORMAL, "", 0, 0
          End If

        Else  ' TWO Button Pendant Mode

          PagerID = Configuration.SurveyPager
          SurveyDevice.Serial = Configuration.SurveyDevice
          If PagerID <> 0 Then
            SurveyDevice.PagerID = PagerID
            SurveyDevice.Serial = Configuration.SurveyDevice  ' use just one device
            SurveyDevice.PCASerial = Configuration.SurveyDevice
            ' create outbound message for ouput "Pager"
            SurveyDevice.RequestSurvey BuildPrompt(waypoint), 2, True  ' only OK, cancel
            dbgloc "Sending message to pager " & SurveyDevice.MessageString
            SendToPager SurveyDevice.MessageString, SurveyDevice.PagerID, 0, "", "", PAGER_NORMAL, "", 0, 0
          End If

        End If
        Exit For

      End If
    Next

  End If
  If SurveyDevice Is Nothing Then
    Locating = False
  End If

NextWayPoint_Resume:
  On Error GoTo 0
  Exit Sub

NextWayPoint_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmWaypoints.NextWayPoint." & Erl
  Resume NextWayPoint_Resume

End Sub
Private Function BuildPrompt(waypoint As cWayPoint) As String
  BuildPrompt = Trim$(waypoint.Description & " " & waypoint.Building & " " & waypoint.Floor & " " & waypoint.Wing)
End Function
Private Sub Form_Deactivate()
'  Locating = False
End Sub

Private Sub Form_Load()
  ResetActivityTime
  Set Selected = New Collection
  cmdLocate.Caption = "Locate"
  Configurelvmain
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
  ResetActivityTime
  cmdDelete.Enabled = gUser.LEvel >= LEVEL_ADMIN
  ShowWaypoints
End Sub
Sub ShowWaypoints()
  Dim j As Integer
  Dim wp As cWayPoint
  Dim li As ListItem
  DisableButtons

  FetchWaypoints

  lvMain.ListItems.Clear
  LockWindowUpdate lvMain.hwnd
  For j = 1 To Waypoints.Count
    'DoEvents
    Set wp = Waypoints.waypoint(j)
        
    Set li = lvMain.ListItems.Add(, wp.ID & "s", wp.Description)
    li.SubItems(1) = wp.Building
    li.SubItems(2) = wp.Floor
    li.SubItems(3) = wp.Wing
    li.SubItems(4) = Format(wp.Signal1, "00") & "-" & Format(wp.Signal2, "00") & "-" & Format(wp.Signal3, "00")


  Next
  Debug.Print "Waypoints " & Waypoints.Count
  Win32.LockWindowUpdate 0
  EnableButtons

End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Set WayPointForm = Nothing
  Set BatchForm = Nothing
  Set SurveyDevice = Nothing
  BatchSurveyEnabled = False
  Locating = False
  
End Sub


Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Select Case ColumnHeader.index
    Case 1
      If SortOrder = 0 Then
        SortOrder = 1
      ElseIf SortOrder = 1 Then
        SortOrder = 0
      End If
    Case Else
  End Select
  lvMain.Sorted = True
  lvMain.SortOrder = IIf(SortOrder = 1, lvwDescending, lvwAscending)
End Sub
Sub DisableButtons()
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdLocate.Enabled = False
  cmdExit.Enabled = False
End Sub
Sub EnableButtons()
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True
  cmdLocate.Enabled = True
  cmdExit.Enabled = True

End Sub
