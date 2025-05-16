VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmStaffEdit 
   Caption         =   "Form1"
   ClientHeight    =   15390
   ClientLeft      =   5070
   ClientTop       =   450
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   15390
   ScaleWidth      =   9225
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      Begin VB.CheckBox chkVacation 
         Alignment       =   1  'Right Justify
         Caption         =   "Vacation"
         Height          =   360
         Left            =   1560
         TabIndex        =   44
         Top             =   330
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtLastName 
         Height          =   315
         Left            =   4695
         MaxLength       =   50
         TabIndex        =   39
         Top             =   0
         Width           =   2310
      End
      Begin VB.TextBox txtFirstName 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   38
         Top             =   0
         Width           =   2310
      End
      Begin VB.CommandButton cmdEditResident 
         Caption         =   "Save"
         Height          =   585
         Left            =   7725
         TabIndex        =   37
         Top             =   1785
         Width           =   1175
      End
      Begin VB.CommandButton cmdAddRes 
         Caption         =   "New"
         Height          =   585
         Left            =   7725
         TabIndex        =   36
         Top             =   30
         Visible         =   0   'False
         Width           =   1175
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Exit"
         Height          =   585
         Left            =   7725
         TabIndex        =   35
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Frame fraTx 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   30
         TabIndex        =   29
         Top             =   7500
         Width           =   7500
         Begin VB.CommandButton cmdResConfigureTx 
            Caption         =   "Configure Transmitter"
            Height          =   525
            Left            =   3405
            TabIndex        =   33
            Top             =   1500
            Width           =   1175
         End
         Begin VB.CommandButton cmdUnassign 
            Caption         =   "Remove Transmitter"
            Height          =   525
            Left            =   2055
            TabIndex        =   32
            Top             =   1500
            Width           =   1175
         End
         Begin VB.CommandButton cmdAssign 
            Caption         =   "Add Transmitter"
            Height          =   525
            Left            =   720
            TabIndex        =   31
            Top             =   1500
            Width           =   1175
         End
         Begin VB.CommandButton cmdAssignRoom 
            Caption         =   "Assign Room"
            Height          =   525
            Left            =   4740
            TabIndex        =   30
            Top             =   1500
            Width           =   1175
         End
         Begin MSComctlLib.ListView lvDevices 
            Height          =   1440
            Left            =   -15
            TabIndex        =   34
            Top             =   -15
            Width           =   7380
            _ExtentX        =   13018
            _ExtentY        =   2540
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Serial #"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Model"
               Object.Width           =   1411
            EndProperty
         End
      End
      Begin VB.Frame fraInfo 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   30
         TabIndex        =   26
         Top             =   9630
         Width           =   7500
         Begin VB.CommandButton cmdEditPicture 
            Caption         =   "Picture"
            Height          =   465
            Left            =   225
            TabIndex        =   28
            Top             =   1530
            Width           =   1035
         End
         Begin VB.TextBox txtInfo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1980
            Left            =   1935
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   -15
            Width           =   5490
         End
         Begin VB.Image imgPic 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1500
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.Frame fraReminders 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   60
         TabIndex        =   21
         Top             =   3270
         Width           =   7500
         Begin VB.CommandButton cmdDeleteReminder 
            Caption         =   "Delete Reminder"
            Height          =   525
            Left            =   4260
            TabIndex        =   24
            Top             =   1530
            Width           =   1125
         End
         Begin VB.CommandButton cmdEditReminder 
            Caption         =   "Edit Reminder"
            Height          =   525
            Left            =   3015
            TabIndex        =   23
            Top             =   1530
            Width           =   1125
         End
         Begin VB.CommandButton cmdAddReminder 
            Caption         =   "Add Reminder"
            Height          =   525
            Left            =   1770
            TabIndex        =   22
            Top             =   1530
            Width           =   1125
         End
         Begin MSComctlLib.ListView lvMain 
            Height          =   1485
            Left            =   -30
            TabIndex        =   25
            Top             =   30
            Width           =   7380
            _ExtentX        =   13018
            _ExtentY        =   2619
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
      Begin VB.Frame fraContact 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   60
         TabIndex        =   1
         Top             =   11820
         Width           =   7500
         Begin VB.TextBox txtContact1 
            Height          =   375
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   15
            Top             =   450
            Width           =   3855
         End
         Begin VB.TextBox txtContact2 
            Height          =   375
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   14
            Top             =   855
            Width           =   3855
         End
         Begin VB.TextBox txtContact3 
            Height          =   375
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   13
            Top             =   1260
            Width           =   3855
         End
         Begin VB.Frame fra2 
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   180
            TabIndex        =   9
            Top             =   480
            Width           =   495
            Begin VB.OptionButton optPrivate1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
            Begin VB.OptionButton optPrivate2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   11
               Top             =   405
               Width           =   315
            End
            Begin VB.OptionButton optPrivate3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   10
               Top             =   810
               Width           =   315
            End
         End
         Begin VB.Frame fra3 
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   750
            TabIndex        =   5
            Top             =   480
            Width           =   495
            Begin VB.OptionButton optPublic3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   8
               Top             =   810
               Width           =   315
            End
            Begin VB.OptionButton optPublic2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   7
               Top             =   405
               Width           =   315
            End
            Begin VB.OptionButton optPublic1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
         End
         Begin VB.ComboBox cboType1 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   450
            Width           =   1845
         End
         Begin VB.ComboBox cboType2 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   870
            Width           =   1845
         End
         Begin VB.ComboBox cboType3 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   1845
         End
         Begin VB.Label lblPrivate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Private"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   20
            Top             =   270
            Width           =   615
         End
         Begin VB.Label lblPublic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Public"
            Height          =   195
            Index           =   2
            Left            =   750
            TabIndex        =   19
            Top             =   270
            Width           =   540
         End
         Begin VB.Label lblDef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Defaults"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   18
            Top             =   60
            Width           =   720
         End
         Begin VB.Label z 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number/Email"
            Height          =   195
            Index           =   1
            Left            =   2370
            TabIndex        =   17
            Top             =   150
            Width           =   1935
         End
         Begin VB.Label lblDef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Method"
            Height          =   195
            Index           =   1
            Left            =   5820
            TabIndex        =   16
            Top             =   150
            Width           =   1365
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2505
         Left            =   15
         TabIndex        =   40
         Top             =   720
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   4419
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Reminders"
               Key             =   "reminders"
               Object.Tag             =   "reminders"
               Object.ToolTipText     =   "Reminders"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Contact Info"
               Key             =   "contact"
               Object.Tag             =   "contact"
               Object.ToolTipText     =   "How Reminders Can Contact Resident"
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
      Begin VB.Label z 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Index           =   4
         Left            =   3705
         TabIndex        =   43
         Top             =   60
         Width           =   915
      End
      Begin VB.Label z 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   42
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblassurdaylist 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4785
         TabIndex        =   41
         Top             =   420
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmStaffEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStaffID As Long
Public CallingForm As String
Private LastIndex   As Long
Private CurrentResident As New cResident


Sub FetchReminders()
 Dim rs               As ADODB.Recordset
  Dim li              As ListItem
  Dim Reminder        As cReminder
  Dim c               As Collection
  
  Set c = New Collection
  
  lvMain.ListItems.Clear
  
  If StaffID <> 0 Then
  
  Set rs = ConnExecute("SELECT * FROM Reminders WHERE ispublic = 1 and OwnerID = " & StaffID & " ORDER BY Description")
  Do Until rs.EOF
    DoEvents
    Set Reminder = New cReminder
    Reminder.Parse rs
    c.Add Reminder
    
    Set li = lvMain.ListItems.Add(, Reminder.reminderid & "s", Reminder.ReminderName)
    ' set an expired flag
    li.SubItems(1) = LCase$(Reminder.FrequencyToString())
    li.SubItems(2) = Reminder.ScheduleToString()
    li.SubItems(3) = Reminder.TimeOfDayToString()
    ' show the date?
    
    
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  End If


End Sub

Private Sub FetchResident()

  Dim rs As Recordset
  RefreshJet
  Set CurrentResident = New cResident
  
  Set rs = ConnExecute("select * from staff where StaffID = " & StaffID)
  If rs.EOF Then
    StaffID = 0
    imgPic.Picture = LoadPicture("")
  Else
    CurrentResident.ResidentID = StaffID
    CurrentResident.NameFirst = rs("nameFirst") & ""
    CurrentResident.NameLast = rs("namelast") & ""
    'CurrentResident.Name = rs("name") & ""
    CurrentResident.Phone = rs("phone") & ""
    'CurrentResident.RoomID = IIf(IsNull(rs("RoomID")), 0, rs("RoomID"))
    'CurrentResident.GroupID = IIf(IsNull(rs("groupID")), 0, rs("groupID"))
    CurrentResident.info = rs("info") & ""
    'CurrentResident.Assurdays = Val(0 & rs("AssurDays")) And &HFF
    CurrentResident.Vacation = IIf(rs("Away") = 1, 1, 0)
    'If rs("imagedata").ActualSize > 0 Then
    'GetImageFromDB imgPic, rs("imagedata")
    
    CurrentResident.DeliveryPointsString = rs("deliverypoints") & ""
    'Else
    ' imgPic.Picture = LoadPicture("")
    'End If
  End If
  rs.Close
'  Set rs = connexecute("select * from Rooms where Roomid = " & CurrentResident.RoomID)
'  If rs.EOF Then
'    CurrentResident.room = ""
'  Else
'    CurrentResident.room = rs("Room") & ""
'  End If
'  rs.Close

'  Set rs = connexecute("select * from pagergroups where groupid = " & CurrentResident.GroupID)
'  If rs.EOF Then
'    CurrentResident.Group = ""
'  Else
'    CurrentResident.Group = rs("Groupname") & ""
'  End If
'  rs.Close
  Set rs = Nothing

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




Function SaveResident() As Boolean
  
  'dbg "SaveResident clicked" & vbCrLf

  If Len(Trim(txtFirstName.text)) = 0 Then
    messagebox Me, "First Name must be filled in.", App.Title, vbInformation
    Exit Function
  End If

  If Len(Trim(Me.txtLastName.text)) = 0 Then
    messagebox Me, "Last Name must be filled in.", App.Title, vbInformation
    Exit Function
  End If

  Me.Enabled = False

  CurrentResident.NameFirst = Trim(txtFirstName.text)
  CurrentResident.NameLast = Trim(txtLastName.text)
  CurrentResident.info = Trim(txtInfo.text)
   
  
  
  CurrentResident.Vacation = chkVacation.Value
  CurrentResident.ResidentID = StaffID
  SetDeliveryPoints ' sets delivery point info
  
  If MASTER Then
    SaveResident = UpdateStaff(CurrentResident, gUser.UserName)
    StaffID = CurrentResident.ResidentID
  Else

    SaveResident = UpdateStaff(CurrentResident, gUser.UserName)
    StaffID = CurrentResident.ResidentID

'    If ClientUpdateStaff(CurrentResident) Then
'      StaffID = CurrentResident.ResidentID
'      SaveResident = True
'    Else
'      SaveResident = False
'    End If
    
  End If
  'lblassurdaylist.Caption = GetAssurDaysFromValue(CurrentResident.Assurdays) & "   " & IIf(CurrentResident.Vacation, "Vac", "")
  Me.Enabled = True
End Function

Sub SetDeliveryPoints()

  Dim dp As cDeliveryPoint
  Dim DeliveryPoints As Collection
  Set DeliveryPoints = New Collection
  
   
  Set dp = New cDeliveryPoint
 
    dp.Address = txtContact1.text
    dp.AddressType = cboType1.ListIndex
    
    If optPrivate1.Value = True And optPublic1.Value = False Then
      dp.Status = 1
    ElseIf optPublic1.Value = True And optPrivate1.Value = False Then
      dp.Status = 2
    ElseIf optPublic1.Value = True And optPrivate1.Value = True Then
      dp.Status = 3
    Else
      dp.Status = 0
    End If
  DeliveryPoints.Add dp
  
    

  Set dp = New cDeliveryPoint
 
    dp.Address = txtContact2.text
    dp.AddressType = cboType2.ListIndex
    If optPrivate2.Value = True And optPublic2.Value = False Then
      dp.Status = 1
    ElseIf optPublic2.Value = True And optPrivate2.Value = False Then
      dp.Status = 2
    ElseIf optPublic2.Value = True And optPrivate2.Value = True Then
      dp.Status = 3
    Else
      dp.Status = 0
    End If
  DeliveryPoints.Add dp
  
  
  Set dp = New cDeliveryPoint
 
    dp.Address = txtContact3.text
    dp.AddressType = cboType3.ListIndex
    If optPrivate3.Value = True And optPublic3.Value = False Then
      dp.Status = 1
    ElseIf optPublic3.Value = True And optPrivate3.Value = False Then
      dp.Status = 2
    ElseIf optPublic3.Value = True And optPrivate3.Value = True Then
      dp.Status = 3
    Else
      dp.Status = 0
    End If
  DeliveryPoints.Add dp
  
  Set CurrentResident.DeliveryPoints = DeliveryPoints
  CurrentResident.DeliveryPointsString = CurrentResident.DeliveryPointsToString
  
End Sub


Sub AssignTransmitter()
  'ShowTransmitters CurrentResident.ResidentID, 0
End Sub

Sub Fill()
  
  RefreshJet
  FetchResident
  txtFirstName.text = CurrentResident.NameFirst
  txtLastName.text = CurrentResident.NameLast
  txtInfo.text = CurrentResident.info
  chkVacation.Value = CurrentResident.Vacation
  
  FillDeliveryPoints
  
      
  lvDevices.ListItems.Clear
  FetchReminders
  
  

End Sub

Private Sub FillDeliveryPoints()
  Dim dp As cDeliveryPoint
  
  optPublic1.Value = True
  optPrivate1.Value = True
  
  txtContact1.text = ""
  txtContact2.text = ""
  txtContact3.text = ""
  
  cboType1.ListIndex = 0
  cboType2.ListIndex = 0
  cboType3.ListIndex = 0


  CurrentResident.ParseDeliveryPoints
  If CurrentResident.DeliveryPoints.Count >= 1 Then
    Set dp = CurrentResident.DeliveryPoints(1)
    txtContact1.text = dp.Address
    Select Case dp.AddressType
      Case DELIVERY_POINT.Phone
        cboType1.ListIndex = DELIVERY_POINT.Phone
      Case DELIVERY_POINT.phone_ack
        cboType1.ListIndex = DELIVERY_POINT.phone_ack
      Case DELIVERY_POINT.EMAIL
        cboType1.ListIndex = DELIVERY_POINT.EMAIL
      Case Else
        cboType1.ListIndex = 0
    End Select
    Select Case dp.Status
      Case DELIVERY_POINT_STATUS.PRIVATE_STATUS
        optPrivate1.Value = True
        optPublic1.Value = False
      Case DELIVERY_POINT_STATUS.PUBLIC_STATUS
        optPublic1.Value = True
        optPrivate1.Value = False
      Case DELIVERY_POINT_STATUS.BOTH_STATUS
        optPublic1.Value = True
        optPrivate1.Value = True
    
      Case Else
        optPublic1.Value = False
        optPrivate1.Value = False
    End Select
  End If
  If CurrentResident.DeliveryPoints.Count >= 2 Then
    Set dp = CurrentResident.DeliveryPoints(2)
    txtContact2.text = dp.Address
    Select Case dp.AddressType
      Case DELIVERY_POINT.Phone
        cboType2.ListIndex = DELIVERY_POINT.Phone
      Case DELIVERY_POINT.phone_ack
        cboType2.ListIndex = DELIVERY_POINT.phone_ack
      Case DELIVERY_POINT.EMAIL
        cboType2.ListIndex = DELIVERY_POINT.EMAIL
      Case Else
        cboType2.ListIndex = 0
    End Select
    Select Case dp.Status
      Case DELIVERY_POINT_STATUS.PRIVATE_STATUS
        optPrivate2.Value = True
        optPublic2.Value = False
      Case DELIVERY_POINT_STATUS.PUBLIC_STATUS
        optPublic2.Value = True
        optPrivate2.Value = False
      Case DELIVERY_POINT_STATUS.BOTH_STATUS
        optPublic2.Value = True
        optPrivate2.Value = True
    
      Case Else
        optPublic2.Value = False
        optPrivate2.Value = False
    End Select
  End If
  If CurrentResident.DeliveryPoints.Count >= 3 Then
    Set dp = CurrentResident.DeliveryPoints(3)
    txtContact3.text = dp.Address
    Select Case dp.AddressType
      Case DELIVERY_POINT.Phone
        cboType3.ListIndex = DELIVERY_POINT.Phone
      Case DELIVERY_POINT.phone_ack
        cboType3.ListIndex = DELIVERY_POINT.phone_ack
      Case DELIVERY_POINT.EMAIL
        cboType3.ListIndex = DELIVERY_POINT.EMAIL
      Case Else
        cboType3.ListIndex = 0
    End Select
    Select Case dp.Status
      Case DELIVERY_POINT_STATUS.PRIVATE_STATUS
        optPrivate3.Value = True
        optPublic3.Value = False
      Case DELIVERY_POINT_STATUS.PUBLIC_STATUS  '   2
        optPublic3.Value = True
        optPrivate3.Value = False
      Case DELIVERY_POINT_STATUS.BOTH_STATUS
        optPublic3.Value = True
        optPrivate3.Value = True
    
      Case Else ' delivery_point_status.neither_STATUS
        optPublic3.Value = False
        optPrivate3.Value = False
    End Select
  End If
      

End Sub


Private Sub cmdAddReminder_Click()
  EditPublicEvent 0, StaffID
End Sub

Private Sub cmdAddRes_Click()
  StaffID = 0
  Set CurrentResident = New cResident
  Fill

End Sub



Private Sub cmdDeleteReminder_Click()
  Dim reminderid As Long
  
  If Not lvMain.SelectedItem Is Nothing Then
    reminderid = Val(lvMain.SelectedItem.Key)
    DeleteReminder reminderid
    
  End If
  FetchReminders
End Sub

Private Sub cmdEditPicture_Click()
  If SaveResident = True Then
    EditPicture
  End If
End Sub
Private Sub EditPicture()
  If CurrentResident.ResidentID <> 0 Then
    ShowPictures CurrentResident.ResidentID
  End If

End Sub

Private Sub cmdEditReminder_Click()
  Dim reminderid As Long
  
  If Not lvMain.SelectedItem Is Nothing Then
    reminderid = Val(lvMain.SelectedItem.Key)
    EditPublicEvent reminderid, StaffID
    
  End If
  
  


  
End Sub

Private Sub cmdResConfigureTx_Click()
  ConfigureTX
End Sub
Sub ConfigureTX()


End Sub



Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If Me.ActiveControl Is Me.txtInfo Then
    Exit Sub
  End If
  Select Case KeyAscii
    
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select

End Sub



Private Sub TabStrip_Click()
  ShowTabData TabStrip.SelectedItem.Key


End Sub

Private Sub txtFirstName_GotFocus()
  SelAll txtFirstName
  
End Sub

Private Sub txtInfo_GotFocus()
  SelAll txtInfo
End Sub

Private Sub txtLastName_GotFocus()
  SelAll txtLastName
End Sub

'Private Sub txtName_GotFocus()
'  SelAll txtName
'End Sub

Private Sub txtRoom_DblClick()
  'AssignRoom
End Sub
Private Sub AssignRoom()
'  If CurrentResident.ResidentID <> 0 Then
'    ShowRooms CurrentResident.ResidentID, 0, CurrentResident.RoomID, "RES"
'  End If
End Sub
Private Sub cmdClose_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdEditResident_Click()
  SaveResident
End Sub

Private Sub cmdUnassign_Click()
  
  
  
  'ResetRemoteRefreshCounter
'  Unassign
  'ResetRemoteRefreshCounter
  Fill
  
End Sub

Private Sub Form_Load()
  ResetActivityTime
  ArrangeControls
  
  ConfigureLVDevices
  Connect
End Sub
Sub ArrangeControls()
  fraEnabler.BackColor = Me.BackColor
  
  fraInfo.left = TabStrip.ClientLeft
  fraInfo.top = TabStrip.ClientTop
  fraInfo.Width = TabStrip.ClientWidth
  fraInfo.Height = TabStrip.ClientHeight
  fraInfo.BackColor = Me.BackColor

  fraTx.left = TabStrip.ClientLeft
  fraTx.top = TabStrip.ClientTop
  fraTx.Width = TabStrip.ClientWidth
  fraTx.Height = TabStrip.ClientHeight
  fraTx.BackColor = Me.BackColor

'  fraAssur.left = TabStrip.ClientLeft
'  fraAssur.top = TabStrip.ClientTop
'  fraAssur.width = TabStrip.ClientWidth
'  fraAssur.height = TabStrip.ClientHeight
'  fraAssur.BackColor = Me.BackColor

  fraReminders.left = TabStrip.ClientLeft
  fraReminders.top = TabStrip.ClientTop
  fraReminders.Width = TabStrip.ClientWidth
  fraReminders.Height = TabStrip.ClientHeight
  fraReminders.BackColor = Me.BackColor

  fraContact.left = TabStrip.ClientLeft
  fraContact.top = TabStrip.ClientTop
  fraContact.Width = TabStrip.ClientWidth
  fraContact.Height = TabStrip.ClientHeight
  fraContact.BackColor = Me.BackColor
  
  ShowTabData TabStrip.SelectedItem.Key

  lvMain.ColumnHeaders.Clear
  lvMain.ColumnHeaders.Add , , "Name", 3000
  lvMain.ColumnHeaders.Add , , "F", 400
  lvMain.ColumnHeaders.Add , , "Days", 1500
  lvMain.ColumnHeaders.Add , , "Time", 1200

  cboType1.Clear
  cboType2.Clear
  cboType3.Clear
  
  AddToCombo cboType1, "Phone", DELIVERY_POINT.Phone
  AddToCombo cboType2, "Phone", DELIVERY_POINT.Phone
  AddToCombo cboType3, "Phone", DELIVERY_POINT.Phone
  
  AddToCombo cboType1, "Phone-ACK", DELIVERY_POINT.phone_ack
  AddToCombo cboType2, "Phone-ACK", DELIVERY_POINT.phone_ack
  AddToCombo cboType3, "Phone-ACK", DELIVERY_POINT.phone_ack
  
  AddToCombo cboType1, "Email", DELIVERY_POINT.EMAIL
  AddToCombo cboType2, "Email", DELIVERY_POINT.EMAIL
  AddToCombo cboType3, "Email", DELIVERY_POINT.EMAIL
  
  cboType1.ListIndex = 0
  cboType2.ListIndex = 0
  cboType3.ListIndex = 0


End Sub
Sub ShowTabData(ByVal TabKey As String)
  Select Case LCase(TabKey)

    Case "info"
      fraInfo.Visible = True
      fraTx.Visible = False
     ' fraAssur.Visible = False
      fraReminders.Visible = False
      fraContact.Visible = False
    Case "assur"
      'fraAssur.Visible = True
      fraInfo.Visible = False
      fraTx.Visible = False
      fraReminders.Visible = False
      fraContact.Visible = False
    Case "reminders"
      fraReminders.Visible = True
      fraInfo.Visible = False
      fraTx.Visible = False
   '   fraAssur.Visible = False
      fraContact.Visible = False
    Case "contact"
      fraContact.Visible = True
      fraReminders.Visible = False
      fraInfo.Visible = False
      fraTx.Visible = False
  '    fraAssur.Visible = False

    Case Else

      fraTx.Visible = True
      'fraAssur.Visible = False
      fraInfo.Visible = False
      fraReminders.Visible = False
      fraContact.Visible = False
  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub



Sub ConfigureLVDevices()
  Dim ch   As ColumnHeader
  lvDevices.ColumnHeaders.Clear
  lvDevices.Sorted = True
  Set ch = lvDevices.ColumnHeaders.Add(, "S", "Serial", 1100)
  Set ch = lvDevices.ColumnHeaders.Add(, "M", "Model", 1200)
  'Set ch = lvDevices.ColumnHeaders.Add(, "Res", "Resident", 1440)
  Set ch = lvDevices.ColumnHeaders.Add(, "Room", "Room", 2000)
  'Set ch = lvDevices.ColumnHeaders.Add(, "Bldg", "Building", 1350)
  Set ch = lvDevices.ColumnHeaders.Add(, "Assur", "Assur", 700)

End Sub

Public Property Get StaffID() As Long

  StaffID = mStaffID

End Property

Public Property Let StaffID(ByVal ID As Long)
  cmdAddReminder.Enabled = ID <> 0
  cmdDeleteReminder.Enabled = ID <> 0
  cmdEditReminder.Enabled = ID <> 0
  
  mStaffID = ID

End Property

