VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmResident 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Resident"
   ClientHeight    =   15435
   ClientLeft      =   390
   ClientTop       =   2010
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15435
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
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
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   9120
      Begin VB.Frame fraContact 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2070
         Left            =   60
         TabIndex        =   41
         Top             =   11820
         Width           =   7500
         Begin VB.ComboBox cboType3 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1320
            Width           =   1845
         End
         Begin VB.ComboBox cboType2 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   870
            Width           =   1845
         End
         Begin VB.ComboBox cboType1 
            Height          =   315
            Left            =   5430
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   450
            Width           =   1845
         End
         Begin VB.Frame fra3 
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   750
            TabIndex        =   49
            Top             =   480
            Width           =   495
            Begin VB.OptionButton optPublic1 
               Height          =   300
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
            Begin VB.OptionButton optPublic2 
               Height          =   300
               Left            =   0
               TabIndex        =   51
               Top             =   405
               Width           =   315
            End
            Begin VB.OptionButton optPublic3 
               Height          =   300
               Left            =   0
               TabIndex        =   50
               Top             =   810
               Width           =   315
            End
         End
         Begin VB.Frame fra2 
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   180
            TabIndex        =   45
            Top             =   480
            Width           =   495
            Begin VB.OptionButton optPrivate3 
               Height          =   300
               Left            =   0
               TabIndex        =   48
               Top             =   810
               Width           =   315
            End
            Begin VB.OptionButton optPrivate2 
               Height          =   300
               Left            =   0
               TabIndex        =   47
               Top             =   405
               Width           =   315
            End
            Begin VB.OptionButton optPrivate1 
               Height          =   300
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
         End
         Begin VB.TextBox txtContact3 
            Height          =   375
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   44
            Top             =   1260
            Width           =   3855
         End
         Begin VB.TextBox txtContact2 
            Height          =   375
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   43
            Top             =   855
            Width           =   3855
         End
         Begin VB.TextBox txtContact1 
            Height          =   375
            Left            =   1410
            MaxLength       =   255
            TabIndex        =   42
            Top             =   450
            Width           =   3855
         End
         Begin VB.Label lblDef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Method"
            Height          =   195
            Index           =   1
            Left            =   5820
            TabIndex        =   60
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label z 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number/Email"
            Height          =   195
            Index           =   1
            Left            =   2370
            TabIndex        =   59
            Top             =   150
            Width           =   1935
         End
         Begin VB.Label lblDef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Defaults"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   60
            Width           =   720
         End
         Begin VB.Label lblPublic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Public"
            Height          =   195
            Index           =   2
            Left            =   750
            TabIndex        =   54
            Top             =   270
            Width           =   540
         End
         Begin VB.Label lblPrivate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Private"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   53
            Top             =   270
            Width           =   615
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
         TabIndex        =   36
         Top             =   3270
         Width           =   7500
         Begin VB.CommandButton cmdAddReminder 
            Caption         =   "Add Reminder"
            Height          =   525
            Left            =   1770
            TabIndex        =   40
            Top             =   1530
            Width           =   1125
         End
         Begin VB.CommandButton cmdEditReminder 
            Caption         =   "Edit Reminder"
            Height          =   525
            Left            =   3015
            TabIndex        =   39
            Top             =   1530
            Width           =   1125
         End
         Begin VB.CommandButton cmdDeleteReminder 
            Caption         =   "Delete Reminder"
            Height          =   525
            Left            =   4260
            TabIndex        =   38
            Top             =   1530
            Width           =   1125
         End
         Begin MSComctlLib.ListView lvMain 
            Height          =   1485
            Left            =   -30
            TabIndex        =   37
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
         TabIndex        =   27
         Top             =   9630
         Width           =   7500
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
            MaxLength       =   1024
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   -15
            Width           =   5490
         End
         Begin VB.CommandButton cmdEditPicture 
            Caption         =   "Picture"
            Height          =   465
            Left            =   225
            TabIndex        =   29
            Top             =   1530
            Width           =   1035
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
         TabIndex        =   22
         Top             =   7500
         Width           =   7500
         Begin VB.CommandButton cmdAssignRoom 
            Caption         =   "Assign Room"
            Height          =   525
            Left            =   4740
            TabIndex        =   30
            Top             =   1500
            Width           =   1175
         End
         Begin VB.CommandButton cmdAssign 
            Caption         =   "Add Transmitter"
            Height          =   525
            Left            =   720
            TabIndex        =   24
            Top             =   1500
            Width           =   1175
         End
         Begin VB.CommandButton cmdUnassign 
            Caption         =   "Remove Transmitter"
            Height          =   525
            Left            =   2055
            TabIndex        =   25
            Top             =   1500
            Width           =   1175
         End
         Begin VB.CommandButton cmdResConfigureTx 
            Caption         =   "Configure Transmitter"
            Height          =   525
            Left            =   3405
            TabIndex        =   26
            Top             =   1500
            Width           =   1175
         End
         Begin MSComctlLib.ListView lvDevices 
            Height          =   1440
            Left            =   -15
            TabIndex        =   23
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
      Begin VB.Frame fraAssur 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   8
         Top             =   5370
         Width           =   7500
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "Clear All"
            Height          =   345
            Left            =   5655
            TabIndex        =   11
            Top             =   225
            Width           =   1175
         End
         Begin VB.CommandButton cmdSetAll 
            Caption         =   "Set All"
            Height          =   345
            Left            =   4425
            TabIndex        =   10
            Top             =   225
            Width           =   1175
         End
         Begin VB.TextBox txtUntil 
            Height          =   315
            Left            =   3435
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1500
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkVacation 
            Alignment       =   1  'Right Justify
            Caption         =   "Vacation"
            Height          =   360
            Left            =   420
            TabIndex        =   19
            Top             =   1455
            Width           =   1335
         End
         Begin VB.CheckBox chkMon 
            Alignment       =   1  'Right Justify
            Caption         =   "Mon"
            Height          =   360
            Left            =   465
            TabIndex        =   12
            Top             =   780
            Width           =   705
         End
         Begin VB.CheckBox chkThu 
            Alignment       =   1  'Right Justify
            Caption         =   "Thu"
            Height          =   360
            Left            =   3396
            TabIndex        =   15
            Top             =   780
            Width           =   675
         End
         Begin VB.CheckBox chkFri 
            Alignment       =   1  'Right Justify
            Caption         =   "Fri"
            Height          =   360
            Left            =   4343
            TabIndex        =   16
            Top             =   780
            Width           =   585
         End
         Begin VB.CheckBox chkSat 
            Alignment       =   1  'Right Justify
            Caption         =   "Sat"
            Height          =   360
            Left            =   5200
            TabIndex        =   17
            Top             =   780
            Width           =   645
         End
         Begin VB.CheckBox chkSun 
            Alignment       =   1  'Right Justify
            Caption         =   "Sun"
            Height          =   360
            Left            =   6120
            TabIndex        =   18
            Top             =   780
            Width           =   675
         End
         Begin VB.CheckBox chkWed 
            Alignment       =   1  'Right Justify
            Caption         =   "Wed"
            Height          =   360
            Left            =   2389
            TabIndex        =   14
            Top             =   780
            Width           =   735
         End
         Begin VB.CheckBox chkTues 
            Alignment       =   1  'Right Justify
            Caption         =   "Tue"
            Height          =   360
            Left            =   1442
            TabIndex        =   13
            Top             =   780
            Width           =   675
         End
         Begin VB.Label lblUntil 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Until"
            Height          =   195
            Left            =   2940
            TabIndex        =   20
            Top             =   1545
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblassurdays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check-in Days"
            Height          =   195
            Index           =   0
            Left            =   435
            TabIndex        =   9
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Exit"
         Height          =   585
         Left            =   7725
         TabIndex        =   33
         Top             =   2370
         Width           =   1175
      End
      Begin VB.CommandButton cmdAddRes 
         Caption         =   "New"
         Height          =   585
         Left            =   7725
         TabIndex        =   31
         Top             =   30
         Width           =   1175
      End
      Begin VB.CommandButton cmdEditResident 
         Caption         =   "Save"
         Height          =   585
         Left            =   7725
         TabIndex        =   32
         Top             =   1785
         Width           =   1175
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   6
         Top             =   360
         Width           =   2310
      End
      Begin VB.TextBox txtFirstName 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   2
         Top             =   0
         Width           =   2310
      End
      Begin VB.TextBox txtLastName 
         Height          =   315
         Left            =   4695
         MaxLength       =   50
         TabIndex        =   4
         Top             =   0
         Width           =   2310
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2505
         Left            =   15
         TabIndex        =   7
         Top             =   720
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   4419
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Transmitters"
               Key             =   "tx"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Info/Picture"
               Key             =   "Info"
               Object.ToolTipText     =   "General Information"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Check-in"
               Key             =   "assur"
               Object.ToolTipText     =   "Assurance Setup"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Reminders"
               Key             =   "reminders"
               Object.Tag             =   "reminders"
               Object.ToolTipText     =   "Reminders"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
      Begin VB.Label lblassurdaylist 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4785
         TabIndex        =   35
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label z 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assur"
         Height          =   195
         Index           =   0
         Left            =   4140
         TabIndex        =   34
         Top             =   420
         Width           =   480
      End
      Begin VB.Label z 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   195
         Index           =   5
         Left            =   660
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
      Begin VB.Label z 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   1
         Top             =   75
         Width           =   915
      End
      Begin VB.Label z 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Index           =   4
         Left            =   3705
         TabIndex        =   3
         Top             =   75
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmResident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mResidentID As Long
Public CallingForm As String
Private LastIndex   As Long
Private CurrentResident As New cResident

Sub FetchReminders()
 Dim Rs      As ADODB.Recordset
  Dim li      As ListItem
  Dim Reminder       As cReminder
  Dim c       As Collection
  Set c = New Collection
  
  lvMain.ListItems.Clear
  
  If ResidentID <> 0 Then
  
  Set Rs = ConnExecute("SELECT * FROM Reminders WHERE ispublic = 0 and OwnerID = " & ResidentID & " ORDER BY Description")
  Do Until Rs.EOF
    DoEvents
    Set Reminder = New cReminder
    Reminder.Parse Rs
    c.Add Reminder
    
    Set li = lvMain.ListItems.Add(, Reminder.reminderid & "s", Reminder.ReminderName)
    ' set an expired flag
    li.SubItems(1) = LCase$(Reminder.FrequencyToString())
    li.SubItems(2) = Reminder.ScheduleToString()
    li.SubItems(3) = Reminder.TimeOfDayToString()
    ' show the date?
    
    
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing
  End If


End Sub

Private Sub FetchResident()

  Dim Rs As Recordset

  Set CurrentResident = New cResident
  CurrentResident.AssurDay(1) = 1
  CurrentResident.AssurDay(2) = 1
  CurrentResident.AssurDay(3) = 1
  CurrentResident.AssurDay(4) = 1
  CurrentResident.AssurDay(5) = 1
  CurrentResident.AssurDay(6) = 1
  CurrentResident.AssurDay(7) = 1
  'Set rs = connexecute("select * from residents where residentid = " & ResidentID)
  Set Rs = ConnExecute("select * from residents where residentid = " & ResidentID)
  If Rs.EOF Then
    ResidentID = 0
    imgPic.Picture = LoadPicture("")
  Else
    CurrentResident.ResidentID = ResidentID
    CurrentResident.NameFirst = Rs("nameFirst") & ""
    CurrentResident.NameLast = Rs("namelast") & ""
    'CurrentResident.Name = rs("name") & ""
    CurrentResident.Phone = Rs("phone") & ""
    CurrentResident.RoomID = IIf(IsNull(Rs("RoomID")), 0, Rs("RoomID"))
    CurrentResident.GroupID = IIf(IsNull(Rs("groupID")), 0, Rs("groupID"))
    CurrentResident.info = Rs("info") & ""
    CurrentResident.Assurdays = Val(0 & Rs("AssurDays")) And &HFF
    CurrentResident.Vacation = IIf(Rs("Away") = 1, 1, 0)
    'If rs("imagedata").ActualSize > 0 Then
    CurrentResident.DeliveryPointsString = Rs("DeliveryPoints") & ""
    GetImageFromDB imgPic, Rs("imagedata")
    'Else
    ' imgPic.Picture = LoadPicture("")
    'End If
  End If
  Rs.Close
  Set Rs = ConnExecute("select * from Rooms where Roomid = " & CurrentResident.RoomID)
  If Rs.EOF Then
    CurrentResident.Room = ""
  Else
    CurrentResident.Room = Rs("Room") & ""
  End If
  Rs.Close

  Set Rs = ConnExecute("select * from pagergroups where groupid = " & CurrentResident.GroupID)
  If Rs.EOF Then
    CurrentResident.Group = ""
  Else
    CurrentResident.Group = Rs("Groupname") & ""
  End If
  Rs.Close
  Set Rs = Nothing

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
  CurrentResident.Phone = Trim(txtName.text)
  'CurrentResident.Room = Trim(txtRoom.Text)
  CurrentResident.info = Trim(txtInfo.text)
  CurrentResident.AssurDay(1) = chkSun.value  ' sunday is 1
  CurrentResident.AssurDay(2) = chkMon.value
  CurrentResident.AssurDay(3) = chkTues.value
  CurrentResident.AssurDay(4) = chkWed.value
  CurrentResident.AssurDay(5) = chkThu.value
  CurrentResident.AssurDay(6) = chkFri.value
  CurrentResident.AssurDay(7) = chkSat.value
  
  'Dim ClearAlarms As Integer
'  ClearAlarms = CurrentResident.Vacation <> chkVacation.Value
  
  CurrentResident.Vacation = chkVacation.value
  
  SetDeliveryPoints
  
  If MASTER Then
    SaveResident = UpdateResident(CurrentResident, gUser.Username)
    ResidentID = CurrentResident.ResidentID
    Residents.FetchAndUpdate ResidentID
  Else

    If ClientUpdateResident(CurrentResident) Then
      ResidentID = CurrentResident.ResidentID
      Residents.FetchAndUpdate ResidentID
      SaveResident = True
    Else
      SaveResident = False
    End If
    
  End If
  lblassurdaylist.Caption = GetAssurDaysFromValue(CurrentResident.Assurdays) & "   " & IIf(CurrentResident.Vacation, "Vac", "")
  Me.Enabled = True
End Function
Sub SetDeliveryPoints()

  Dim dp As cDeliveryPoint
  Dim DeliveryPoints As Collection
  Set DeliveryPoints = New Collection
  
   
  Set dp = New cDeliveryPoint
 
    dp.Address = txtContact1.text
    dp.AddressType = cboType1.ListIndex
    
    If optPrivate1.value = True And optPublic1.value = False Then
      dp.Status = 1
    ElseIf optPublic1.value = True And optPrivate1.value = False Then
      dp.Status = 2
    ElseIf optPublic1.value = True And optPrivate1.value = True Then
      dp.Status = 3
    Else
      dp.Status = 0
    End If
  DeliveryPoints.Add dp
  
    

  Set dp = New cDeliveryPoint
 
    dp.Address = txtContact2.text
    dp.AddressType = cboType2.ListIndex
    If optPrivate2.value = True And optPublic2.value = False Then
      dp.Status = 1
    ElseIf optPublic2.value = True And optPrivate2.value = False Then
      dp.Status = 2
    ElseIf optPublic2.value = True And optPrivate2.value = True Then
      dp.Status = 3
    Else
      dp.Status = 0
    End If
  DeliveryPoints.Add dp
  
  
  Set dp = New cDeliveryPoint
 
    dp.Address = txtContact3.text
    dp.AddressType = cboType3.ListIndex
    If optPrivate3.value = True And optPublic3.value = False Then
      dp.Status = 1
    ElseIf optPublic3.value = True And optPrivate3.value = False Then
      dp.Status = 2
    ElseIf optPublic3.value = True And optPrivate3.value = True Then
      dp.Status = 3
    Else
      dp.Status = 0
    End If
  DeliveryPoints.Add dp
  
  Set CurrentResident.DeliveryPoints = DeliveryPoints
  CurrentResident.DeliveryPointsString = CurrentResident.DeliveryPointsToString
  
End Sub


Sub AssignTransmitter()
  ShowTransmitters CurrentResident.ResidentID, 0
End Sub

Sub Fill()
  Dim li As ListItem
  Dim j As Integer
  Dim d As cESDevice
  Dim assure As String

'  Sleep 100
  'Debug.Print "frmResident.Fill"
  'dbg "frmResident.Fill" & vbCrLf
  
  RefreshJet
  FetchResident

  
  txtFirstName.text = CurrentResident.NameFirst
  txtLastName.text = CurrentResident.NameLast
  txtName.text = CurrentResident.Phone
  txtInfo.text = CurrentResident.info


  chkSun.value = CurrentResident.AssurDay(1)
  chkMon.value = CurrentResident.AssurDay(2)
  chkTues.value = CurrentResident.AssurDay(3)
  chkWed.value = CurrentResident.AssurDay(4)
  chkThu.value = CurrentResident.AssurDay(5)
  chkFri.value = CurrentResident.AssurDay(6)
  chkSat.value = CurrentResident.AssurDay(7)
  
  lblassurdaylist.Caption = GetAssurDaysFromValue(CurrentResident.Assurdays) & "   " & IIf(CurrentResident.Vacation, "Vac", "")
  chkVacation.value = CurrentResident.Vacation

  FillDeliveryPoints

  lvDevices.ListItems.Clear


  CurrentResident.GetTransmitters
  
  For j = 1 To CurrentResident.AssignedTx.Count
    Set d = CurrentResident.AssignedTx(j)
        
    Set li = lvDevices.ListItems.Add(, d.DeviceID & "S", Right("00000000" & d.Serial, 8))
    li.SubItems(1) = d.Model    'rs("model") & ""
    li.SubItems(2) = d.Description     'rs("model") & ""
    li.SubItems(3) = d.Room
    'li.SubItems(3) = " " 'd.Building   'need to populate

    
    'If d.NumInputs > 1 Then
      assure = IIf(d.UseAssur = 1, "Y", "N") & IIf(d.UseAssur2 = 1, "Y", "N")
      If d.UseAssur = 1 Or d.UseAssur2 = 1 Then
        assure = assure & d.AssurInput
      End If
    li.SubItems(4) = assure
        
  Next

  FetchReminders


End Sub

Private Sub FillDeliveryPoints()
  Dim dp As cDeliveryPoint
  
  optPublic1.value = True
  optPrivate1.value = True
  
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
      Case 1
        optPrivate1.value = True
        optPublic1.value = False
      Case 2
        optPublic1.value = True
        optPrivate1.value = False
      Case 3
        optPublic1.value = True
        optPrivate1.value = True
    
      Case Else
        optPublic1.value = False
        optPrivate1.value = False
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
      Case 1
        optPrivate2.value = True
        optPublic2.value = False
      Case 2
        optPublic2.value = True
        optPrivate2.value = False
      Case 3
        optPublic2.value = True
        optPrivate2.value = True
    
      Case Else
        optPublic2.value = False
        optPrivate2.value = False
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
      Case 1
        optPrivate3.value = True
        optPublic3.value = False
      Case 2
        optPublic3.value = True
        optPrivate3.value = False
      Case 3
        optPublic3.value = True
        optPrivate3.value = True
    
      Case Else
        optPublic3.value = False
        optPrivate3.value = False
    End Select
  End If
      

End Sub

Private Sub cmdAddReminder_Click()
  EditPrivateEvent 0, ResidentID
End Sub

Private Sub cmdAddRes_Click()
  ResidentID = 0
  Set CurrentResident = New cResident
  CurrentResident.AssurDay(1) = 1
  CurrentResident.AssurDay(2) = 1
  CurrentResident.AssurDay(3) = 1
  CurrentResident.AssurDay(4) = 1
  CurrentResident.AssurDay(5) = 1
  CurrentResident.AssurDay(6) = 1
  CurrentResident.AssurDay(7) = 1

  
  Fill

End Sub

Private Sub cmdAssign_Click()
  
  If CurrentResident.ResidentID = 0 Then
    'ResetRemoteRefreshCounter
    If SaveResident() Then
      AssignTransmitter
    End If
  Else
    AssignTransmitter

  End If
End Sub


Private Sub cmdAssignRoom_Click()
  'ResetRemoteRefreshCounter
  
  Dim d As cESDevice
  If ResidentID <> 0 Then
    If lvDevices.SelectedItem Is Nothing Then
      Beep
    Else
      Set d = Devices.device(lvDevices.SelectedItem.text)
      If d Is Nothing Then
        Beep
      Else
        ShowRooms ResidentID, d.DeviceID, d.RoomID, "RES"
      End If
    End If
  Else
    Beep
  End If

End Sub

Private Sub cmdClearAll_Click()
  chkMon.value = 0
  chkTues.value = 0
  chkWed.value = 0
  chkThu.value = 0
  chkFri.value = 0
  chkSat.value = 0
  chkSun.value = 0
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
    EditPrivateEvent reminderid, ResidentID
    
  End If
  
  


  
End Sub

Private Sub cmdResConfigureTx_Click()
  ConfigureTX
End Sub
Sub ConfigureTX()
  If Not lvDevices.SelectedItem Is Nothing Then
    EditTransmitter Val(lvDevices.SelectedItem.Key)
    Fill
  Else
    Beep
  End If

End Sub


Private Sub cmdSetAll_Click()
  chkMon.value = 1
  chkTues.value = 1
  chkWed.value = 1
  chkThu.value = 1
  chkFri.value = 1
  chkSat.value = 1
  chkSun.value = 1
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

Private Sub lvDevices_ItemClick(ByVal Item As MSComctlLib.ListItem)
  'ResetRemoteRefreshCounter
End Sub

Private Sub lvMain_DblClick()
  Dim reminderid As Long
  
  If Not lvMain.SelectedItem Is Nothing Then
    reminderid = Val(lvMain.SelectedItem.Key)
    EditPrivateEvent reminderid, ResidentID
    
  End If
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

Private Sub txtName_GotFocus()
  SelAll txtName
End Sub

Private Sub txtRoom_DblClick()
  'AssignRoom
End Sub
Private Sub AssignRoom()
  If CurrentResident.ResidentID <> 0 Then
    ShowRooms CurrentResident.ResidentID, 0, CurrentResident.RoomID, "RES"
  End If
End Sub
Private Sub cmdClose_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdEditResident_Click()
  SaveResident
  frmMain.SetListTabs
  frmMain.DisplayResidentInfo Val(frmMain.txtHiddenResID.text), 0
End Sub

Private Sub cmdUnassign_Click()
  
  
  
  'ResetRemoteRefreshCounter
  Unassign
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

  fraassur.left = TabStrip.ClientLeft
  fraassur.top = TabStrip.ClientTop
  fraassur.Width = TabStrip.ClientWidth
  fraassur.Height = TabStrip.ClientHeight
  fraassur.BackColor = Me.BackColor

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


If NO_REMINDERS Then
  TabStrip.Tabs.Remove ("contact")
  TabStrip.Tabs.Remove ("reminders")
End If

End Sub
Sub ShowTabData(ByVal TabKey As String)
  Select Case LCase(TabKey)

    Case "info"
      fraInfo.Visible = True
      fraTx.Visible = False
      fraassur.Visible = False
      fraReminders.Visible = False
      fraContact.Visible = False
    Case "assur"
      fraassur.Visible = True
      fraInfo.Visible = False
      fraTx.Visible = False
      fraReminders.Visible = False
      fraContact.Visible = False
    Case "reminders"
      fraReminders.Visible = True
      fraInfo.Visible = False
      fraTx.Visible = False
      fraassur.Visible = False
      fraContact.Visible = False
    Case "contact"
      fraContact.Visible = True
      fraReminders.Visible = False
      fraInfo.Visible = False
      fraTx.Visible = False
      fraassur.Visible = False

    Case Else

      fraTx.Visible = True
      fraassur.Visible = False
      fraInfo.Visible = False
      fraReminders.Visible = False
      fraContact.Visible = False
  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub

Private Sub Unassign()
  Dim TxID As Long
  Dim li As ListItem
  Dim Rs As Recordset
  Dim device As cESDevice
  Dim t As Long


  Set li = lvDevices.SelectedItem
  If Not li Is Nothing Then
    TxID = Val(li.Key)
    If MASTER Then
      ConnExecute "UPDATE devices SET ResidentID = 0 WHERE deviceid = " & TxID
      Set Rs = ConnExecute("SELECT Devices.Serial FROM Devices WHERE DeviceID = " & TxID)
      If Not Rs.EOF Then
        Set device = Devices.device(Rs("serial"))
        If device Is Nothing Then
          ' nothing to do
        Else
          device.Clear
          device.Refresh
        End If
      End If
      Rs.Close
      Set Rs = Nothing
    Else
      'Debug.Print "ClientUpdateDeviceResidentID Enter"
      
      ClientUpdateDeviceResidentID 0, TxID
      
      'Debug.Print "ClientUpdateDeviceResidentID Done"
    End If
  Else
    Beep
  End If

End Sub

Sub ConfigureLVDevices()
  Dim ch   As ColumnHeader
  lvDevices.ColumnHeaders.Clear
  lvDevices.Sorted = True
  Set ch = lvDevices.ColumnHeaders.Add(, "S", "Serial", 1100)
  Set ch = lvDevices.ColumnHeaders.Add(, "M", "Model", 1100)
  Set ch = lvDevices.ColumnHeaders.Add(, "D", "Desc", 2000)
  Set ch = lvDevices.ColumnHeaders.Add(, "Room", "Room", 2000)

  Set ch = lvDevices.ColumnHeaders.Add(, "Assur", "Assur", 700)

End Sub

Public Property Get ResidentID() As Long

  ResidentID = mResidentID

End Property

Public Property Let ResidentID(ByVal ResidentID As Long)
  cmdAddReminder.Enabled = ResidentID <> 0
  cmdDeleteReminder.Enabled = ResidentID <> 0
  cmdEditReminder.Enabled = ResidentID <> 0
  
  mResidentID = ResidentID

End Property
