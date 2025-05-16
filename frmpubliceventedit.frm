VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPublicEventEdit 
   Caption         =   "Public Event Edit"
   ClientHeight    =   15375
   ClientLeft      =   540
   ClientTop       =   1920
   ClientWidth     =   10680
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
   ScaleHeight     =   15375
   ScaleWidth      =   10680
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   15495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10695
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
         Height          =   585
         Left            =   8025
         TabIndex        =   34
         Top             =   2280
         Width           =   1175
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save"
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
         Left            =   8025
         TabIndex        =   33
         Top             =   1695
         Width           =   1175
      End
      Begin VB.Frame fraSystem 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   0
         TabIndex        =   30
         Top             =   11550
         Width           =   9225
         Begin VB.CommandButton cmdAddPager 
            Caption         =   "<="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3510
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   690
            Width           =   450
         End
         Begin VB.CommandButton cmdRemovePager 
            Caption         =   "=X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3510
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1650
            Width           =   450
         End
         Begin MSComctlLib.ListView lvActivePagers 
            Height          =   2265
            Left            =   90
            TabIndex        =   37
            Top             =   270
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3995
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
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
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvPagers 
            DragIcon        =   "frmPublicEventEdit.frx":0000
            Height          =   2265
            Left            =   4020
            TabIndex        =   38
            Top             =   270
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3995
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
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
            NumItems        =   0
         End
      End
      Begin VB.Frame fraSchedule 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   0
         TabIndex        =   4
         Top             =   5970
         Width           =   9225
         Begin VB.Frame fraDOM 
            BorderStyle     =   0  'None
            Height          =   2085
            Left            =   4350
            TabIndex        =   19
            Top             =   180
            Width           =   2415
            Begin VB.ListBox lstDOM 
               Height          =   1605
               IntegralHeight  =   0   'False
               ItemData        =   "frmPublicEventEdit.frx":030A
               Left            =   60
               List            =   "frmPublicEventEdit.frx":0311
               TabIndex        =   20
               Top             =   300
               Width           =   2235
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Day of Month"
               Height          =   195
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   1155
            End
         End
         Begin VB.ComboBox cboMinute 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   990
            Width           =   675
         End
         Begin VB.ComboBox cboHour 
            Height          =   315
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   990
            Width           =   1575
         End
         Begin VB.Frame fraDays 
            BorderStyle     =   0  'None
            Height          =   2085
            Left            =   4350
            TabIndex        =   23
            Top             =   180
            Width           =   2415
            Begin VB.ListBox lstDOW 
               Height          =   1635
               Left            =   90
               Style           =   1  'Checkbox
               TabIndex        =   26
               Top             =   330
               Width           =   2235
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               Height          =   195
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Width           =   435
            End
         End
         Begin VB.Frame fraCalendar 
            BorderStyle     =   0  'None
            Height          =   2385
            Left            =   4320
            TabIndex        =   21
            Top             =   150
            Width           =   3165
            Begin MSComCtl2.MonthView mvCalendar 
               CausesValidation=   0   'False
               Height          =   2310
               Left            =   60
               TabIndex        =   22
               Top             =   60
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   4075
               _Version        =   393216
               ForeColor       =   -2147483640
               BackColor       =   -2147483633
               Appearance      =   0
               ShowToday       =   0   'False
               StartOfWeek     =   100532225
               CurrentDate     =   40080
               MaxDate         =   55153
               MinDate         =   39814
            End
         End
         Begin VB.ComboBox cboWhen 
            Height          =   315
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   405
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "When"
            Height          =   195
            Left            =   270
            TabIndex        =   25
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            Height          =   195
            Left            =   270
            TabIndex        =   24
            Top             =   750
            Width           =   420
         End
      End
      Begin VB.Frame fraAttendees 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   0
         TabIndex        =   3
         Top             =   8790
         Width           =   9225
         Begin VB.CommandButton cmdRemove 
            Caption         =   "=X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3510
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1650
            Width           =   450
         End
         Begin VB.CommandButton cmdSubscribe 
            Caption         =   "<="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3510
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   690
            Width           =   450
         End
         Begin MSComctlLib.ListView lvSubscribers 
            Height          =   2265
            Left            =   90
            TabIndex        =   5
            Top             =   270
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3995
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
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
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvPeople 
            DragIcon        =   "frmPublicEventEdit.frx":031D
            Height          =   2265
            Left            =   4020
            TabIndex        =   8
            Top             =   270
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3995
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
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
            NumItems        =   0
         End
      End
      Begin VB.Frame framain 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   0
         TabIndex        =   2
         Top             =   3210
         Width           =   9225
         Begin VB.CommandButton cmdGetStaff 
            Height          =   330
            Left            =   4050
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPublicEventEdit.frx":0627
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1410
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboType1 
            Height          =   315
            Left            =   300
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2055
            Width           =   5445
         End
         Begin VB.TextBox txtOwner 
            Height          =   345
            Left            =   270
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1395
            Width           =   3705
         End
         Begin VB.TextBox txtMessage 
            Height          =   345
            Left            =   270
            MaxLength       =   100
            TabIndex        =   10
            Top             =   795
            Width           =   7455
         End
         Begin VB.TextBox txtName 
            Height          =   375
            Left            =   270
            TabIndex        =   9
            Top             =   225
            Width           =   3255
         End
         Begin VB.Label lblContact 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner Delivery Point"
            Height          =   195
            Left            =   300
            TabIndex        =   32
            Top             =   1830
            Width           =   1800
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner Name"
            Height          =   195
            Left            =   270
            TabIndex        =   31
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reminder Message"
            Height          =   195
            Left            =   270
            TabIndex        =   15
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reminder Name"
            Height          =   195
            Left            =   270
            TabIndex        =   14
            Top             =   30
            Width           =   1350
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2970
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   5239
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Event"
               Key             =   "main"
               Object.Tag             =   "main"
               Object.ToolTipText     =   "General Event Info"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Schedule"
               Key             =   "schedule"
               Object.Tag             =   "schedule"
               Object.ToolTipText     =   "Days and Times"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Participants"
               Key             =   "attendees"
               Object.Tag             =   "attendees"
               Object.ToolTipText     =   "Who is Participating"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "System Paging"
               Key             =   "system"
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
   End
   Begin VB.Label lblHeaderremote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Times"
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "frmPublicEventEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReturnValue As String

Public reminderid As Long
Public OwnerID    As Long
Private mIsPublic As Long


Private Resident   As cResident
Private Reminder As cReminder

Private DragSource As ListView

Private lastx    As Single
Private lasty    As Single

Private Busy     As Boolean

Private Cancelling As Boolean

Private Type OutputType
  IsGroup         As Long
  ID              As Long
  Description     As String


End Type

Dim outputs() As OutputType


Function FillActiveOutputs()
  Dim li          As ListItem
  Dim Subscriber  As cReminderSubscriber
  On Error Resume Next

  LockWindowUpdate lvActivePagers.hwnd
  lvActivePagers.ListItems.Clear
  For Each Subscriber In Reminder.subscribers
    DoEvents
    If Subscriber.PagerID <> 0 Or Subscriber.GroupID <> 0 Then
    Set li = lvActivePagers.ListItems.Add(, Subscriber.PagerKey, Subscriber.PagerName)
    li.SubItems(1) = IIf(Subscriber.IsGroup, "G", "P")
    End If
  Next
  LockWindowUpdate 0
End Function

Function FillOutputs() As Long
  Dim rs    As ADODB.Recordset
  Dim SQl   As String
  Dim i     As Long
  Dim li    As ListItem
  Dim PagerOrGroup As String
  Dim Key As String

  On Error GoTo FillOutputs_Error


  ReDim outputs(0) As OutputType

  SQl = "Select * from pagergroups order by description"
  Set rs = ConnExecute(SQl)
  Do Until rs.EOF
    i = i + 1
    ReDim Preserve outputs(i)
    outputs(i).IsGroup = 1
    outputs(i).ID = rs("groupid")
    outputs(i).Description = rs("Description") & ""
    rs.MoveNext
  Loop
  rs.Close

  SQl = "Select * from pagers order by description"
  Set rs = ConnExecute(SQl)
  Do Until rs.EOF
    i = i + 1
    ReDim Preserve outputs(i)
    outputs(i).ID = rs("pagerid")
    outputs(i).Description = rs("Description") & ""
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing


  lvPagers.ListItems.Clear
  For i = 1 To UBound(outputs)
    PagerOrGroup = IIf(outputs(i).IsGroup, "G", "P")
    If outputs(i).IsGroup Then
      Key = "0|" & CStr(outputs(i).ID)
    Else
      Key = CStr(outputs(i).ID) & "|0"
    End If
    Set li = lvPagers.ListItems.Add(, Key, outputs(i).Description)
    li.SubItems(1) = PagerOrGroup
  Next




FillOutputs_Resume:
  On Error GoTo 0
  Exit Function

FillOutputs_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmPublicEventEdit.FillOutputs." & Erl
  Resume FillOutputs_Resume


End Function
Public Function GetAvailableOutputs() As Collection
  
'  Dim rs    As ADODB.Recordset
'  Dim SQL   As String
'  Dim i     As Long
'  Dim li    As ListItem
'  Dim PagerOrGroup As String
'  Dim Key As String
'
'
'  SQL = "Select * from pagergroups order by description"
'  Set rs = ConnExecute(SQL)
'  Do Until rs.EOF
'    i = i + 1
'    ReDim Preserve outputs(i)
'    outputs(i).IsGroup = 1
'    outputs(i).ID = rs("groupid")
'    outputs(i).Description = rs("Description") & ""
'    rs.MoveNext
'  Loop
'  rs.Close
'
'  SQL = "Select * from pagers order by description"
'  Set rs = ConnExecute(SQL)
'  Do Until rs.EOF
'    i = i + 1
'    ReDim Preserve outputs(i)
'    outputs(i).ID = rs("pagerid")
'    outputs(i).Description = rs("Description") & ""
'    rs.MoveNext
'  Loop
'  rs.Close
'  Set rs = Nothing
'

End Function


Function FillOwner() As Long
' fills only owner info

  Set Resident = New cResident
  If IsPublic = 0 Then
    If OwnerID = 0 Then
      OwnerID = Reminder.OwnerID
    End If
    Resident.Fetch OwnerID
  Else
    ' fetch Staff
    Dim rs As Recordset
    Set rs = ConnExecute("Select * FROM staff WHERE staffID = " & OwnerID)
    If Not rs.EOF Then
      Resident.NameLast = rs("NameLast") & ""
      Resident.NameFirst = rs("Namefirst") & ""
      Resident.Room = ""
      Resident.RoomID = 0
      Resident.info = rs("info") & ""
      Resident.Phone = rs("phone") & ""
      Resident.Vacation = 0
      Resident.Assurdays = 0
      'Resident.Parse rs
      Resident.ResidentID = OwnerID
      Resident.DeliveryPointsString = rs("DeliveryPoints") & ""
    End If
    rs.Close
    Set rs = Nothing

    ' Resident.Fetch OwnerID
  End If
  txtOwner.text = Resident.NameLast & ", " & Resident.NameFirst

  Resident.ParseDeliveryPoints

  Dim dp As cDeliveryPoint
  Dim j As Long

  cboType1.Clear
  For j = 1 To Resident.DeliveryPoints.Count
    Set dp = Resident.DeliveryPoints(j)
    cboType1.AddItem dp.Address & " " & ReminderStatusToString(dp.Status)
  Next

End Function

Function FillPeople() As Long

  Dim rs      As ADODB.Recordset
  Dim SQl     As String
  Dim Key     As String
  Dim li      As ListItem


  Dim people As Collection
  Dim Person As cReminderSubscriber

  Set people = GetPeople()

  lvPeople.ListItems.Clear

  LockWindowUpdate lvPeople.hwnd
  For Each Person In people
    DoEvents
    Set li = lvPeople.ListItems.Add(, Person.ResidentKey, Person.NameAll)
    li.SubItems(1) = IIf(Person.IsResident, "R", "N")

  Next
  LockWindowUpdate 0

End Function

Public Function GetPeople() As Collection
  Dim people  As Collection
  Dim Person  As cReminderSubscriber
  Dim rs      As ADODB.Recordset
  Dim SQl     As String
  
  Set people = New Collection
  
  SQl = "SELECT Namelast, NameFirst, ResidentID, 0 as StaffID, 1 as IsResident FROM Residents WHERE deleted = 0   "
  SQl = SQl & "UNION ALL "
  SQl = SQl & "SELECT  Namelast, NameFirst, 0 as ResidentID, StaffID, 0 as IsResident   FROM Staff  WHERE deleted = 0  "
  SQl = SQl & "ORDER BY NameLast, NameFirst"
  Set rs = ConnExecute(SQl)
  Do Until rs.EOF

    Set Person = New cReminderSubscriber
    Person.NameLast = rs("Namelast") & ""
    Person.NameFirst = rs("Namefirst") & ""
    Person.ResidentID = rs("ResidentID")
    Person.StaffID = rs("staffID")
    people.Add Person

    rs.MoveNext

  Loop
  rs.Close
  Set rs = Nothing
  
  Set GetPeople = people
  
  
End Function

Function FillSubscribers() As Long

  Dim li          As ListItem
  Dim Subscriber  As cReminderSubscriber

  LockWindowUpdate lvSubscribers.hwnd
  lvSubscribers.ListItems.Clear
  For Each Subscriber In Reminder.subscribers
    DoEvents
    If Subscriber.ResidentID <> 0 Or Subscriber.StaffID <> 0 Then
    Set li = lvSubscribers.ListItems.Add(, Subscriber.ResidentKey, Subscriber.NameAll)
    li.SubItems(1) = IIf(Subscriber.IsResident, "R", "N")
    End If
  Next
  LockWindowUpdate 0

End Function

Public Function ReminderStatusToString(ByVal Status As String) As String
  Select Case Status
    Case DELIVERY_POINT_STATUS.PUBLIC_STATUS
      ReminderStatusToString = "(Public)"
    Case DELIVERY_POINT_STATUS.PRIVATE_STATUS
      ReminderStatusToString = "(Private)"
    Case DELIVERY_POINT_STATUS.BOTH_STATUS
      ReminderStatusToString = "(Public/Private)"
    Case Else
      ReminderStatusToString = ""
  End Select
End Function

Function Save() As Boolean

  Dim Bitfield    As Long
  Dim j           As Integer


  Dim li          As ListItem

  cmdOK.Enabled = False

  Dim DayCodes As String

  DayCodes = "1234567"

  For j = 0 To lstDOW.listcount - 1
    If lstDOW.Selected(j) Then
      Bitfield = Bitfield Or 2 ^ j
    ElseIf j < 7 Then
      Mid(DayCodes, j + 1, 1) = "_"
    End If

  Next



  ' fill object
  Reminder.Cancelled = 0  'iif(me.chk1
  Reminder.OwnerID = OwnerID
  Reminder.IsPublic = IsPublic
  Reminder.ReminderName = Trim$(txtName.text)
  Reminder.ReminderMessage = Trim$(txtMessage.text)

  Reminder.DeliveryPointID = cboType1.ListIndex

  Reminder.Coordinator = ""
  Reminder.LeadTime = 0
  Reminder.Disabled = 0
  Reminder.Cancelled = 0
  Reminder.Recurring = 0
  Reminder.Frequency = GetComboItemData(cboWhen)
  Reminder.DOW = Bitfield
  Reminder.DOM = GetListBoxItemData(lstDOM)
  Reminder.DayString = DayCodes

  'Days = parsedaysstring
  If Reminder.Frequency = 1 Then
    Reminder.SpecificDay = Format$(mvCalendar.Value, "mm/dd/yyyy")
  Else
    Reminder.SpecificDay = Format$(Now, "mm/dd/yyyy")
  End If

  Reminder.TimeHours = GetComboItemData(cboHour)
  Reminder.TimeMinutes = GetComboItemData(cboMinute)


  Reminder.ClearSubscribers

  Dim Subscriber As cReminderSubscriber

  For Each li In lvSubscribers.ListItems
    Set Subscriber = New cReminderSubscriber
    Subscriber.ResidentKey = li.Key
    Reminder.subscribers.Add Subscriber
  Next

  'Actually save it
  ' may need to marshall this for remotes
  For Each li In lvActivePagers.ListItems
    Set Subscriber = New cReminderSubscriber
    Subscriber.PagerKey = li.Key
    Reminder.subscribers.Add Subscriber
  Next



  reminderid = modReminders.SaveReminder(Reminder)
  modReminders.SaveSubscribers Reminder.subscribers, reminderid
 ' modReminders.SaveReminderPagers Reminder.subscribers, ReminderID
  '  If ReminderID <> 0 Then
  '      ' save attendees to database in the ReminderSubscribers table
  '      modReminders.SaveSubScribers ReminderID, Reminder.SubScribers
  '
  '  End If

  Fill

  cmdOK.Enabled = True
End Function

Public Property Get IsPublic() As Long
  IsPublic = mIsPublic
End Property

Public Property Get ReturnValue() As String
  ReturnValue = mReturnValue
End Property

Public Property Let IsPublic(ByVal Value As Long)
  Dim j As Integer
  cmdGetStaff.Visible = (Value <> 0)
  If (Value = 0) Then
    For j = 1 To TabStrip.Tabs.Count
      If TabStrip.Tabs(j).Key = "system" Then
        TabStrip.Tabs.Remove j
        Exit For
      End If
    Next
  Else
    For j = 1 To TabStrip.Tabs.Count
      If TabStrip.Tabs(j).Key = "system" Then
        Exit For
      End If
    Next
    If j > TabStrip.Tabs.Count Then
      TabStrip.Tabs.Add , "system", "System Paging"
    End If
  End If
  mIsPublic = Value
  ShowPanel TabStrip.SelectedItem.Key
End Property

Public Property Let ReturnValue(ByVal Value As String)
  OwnerID = Val(Value)
  mReturnValue = Value
  FillOwner
End Property

Private Sub cboWhen_Click()
  Dim index As Long
  index = GetComboItemData(cboWhen)
  Select Case GetComboItemData(cboWhen)

    Case 6  ' Annualy
    Case 5  ' Quarterly
    Case 4  ' Monthly
      fraDOM.Visible = True
      fraDays.Visible = False
      fraCalendar.Visible = False
    Case 3  ' Weekly
    Case 2  ' Daily
      fraDays.Visible = True
      fraCalendar.Visible = False
      fraDOM.Visible = False
    Case 1  ' Date
      fraCalendar.Visible = True
      fraDOM.Visible = False
      fraDays.Visible = False
    Case Else  ' 0 = inactive
      fraCalendar.Visible = False
      fraDOM.Visible = False
      fraDays.Visible = False
  End Select
End Sub

Private Sub cmdAddPager_Click()
  PageSelected
End Sub

Private Sub cmdCancel_Click()
  Debug.Print "Cancel Click"
  Cancelling = True
  PreviousForm
  Unload Me
  Cancelling = False
  Debug.Print "Cancel Click Done"
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Debug.Print "Cancel Mousedown"
  If Not Cancelling Then
    Debug.Print "doing prevoius form"
    PreviousForm
    Unload Me
  End If
End Sub

Private Sub cmdGetStaff_Click()
  ShowStaff OwnerID, reminderid, 1, "EventEdit"
End Sub

Private Sub cmdOK_Click()
  Debug.Print " frmpublicevent.OK.Click"

  Save

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
  Save

End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
  Save
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Save
End Sub

Private Sub cmdRemove_Click()
  UnsubscribeSelected
End Sub

Private Sub cmdRemovePager_Click()
  UnpageSelected
End Sub

Private Sub cmdSubscribe_Click()
  SubscribeSelected
End Sub

Public Sub Fill()

  Busy = True

  Set Reminder = modReminders.GetReminder(reminderid)
  If OwnerID = 0 Then
    OwnerID = Reminder.OwnerID
  End If
  modReminders.FetchSubScribers Reminder

  FillOwner

  FillForm
  FillSubscribers
  FillPeople
  FillOutputs
  FillActiveOutputs
  On Error Resume Next
  If cboType1.listcount > 0 Then
    If cboType1.listcount > Reminder.DeliveryPointID Then
      cboType1.ListIndex = Reminder.DeliveryPointID
    Else
      cboType1.ListIndex = 0
    End If
  Else
    cboType1.ListIndex = -1
  End If


  Busy = False
End Sub

Sub FillForm()
  Dim j As Long

  Dim CalendarDay As Date
  ' fill object
  txtName.text = Reminder.ReminderName
  txtMessage.text = Reminder.ReminderMessage
  'txtOwner.Text = Resident.NameLast & ", " & Resident.NameFirst
  ' get contact info for resident


  Reminder.LeadTime = 0

  For j = cboWhen.listcount - 1 To 1 Step -1
    If cboWhen.ItemData(j) = Reminder.Frequency Then
      Exit For
    End If
  Next
  cboWhen.ListIndex = j




  If IsDate(Reminder.SpecificDay) Then
    CalendarDay = CDate(Reminder.SpecificDay)
  End If
  If 0 = CalendarDay Then
    mvCalendar.Value = Now
  Else
    mvCalendar.Value = CalendarDay
  End If

  For j = cboHour.listcount - 1 To 1 Step -1
    If cboHour.ItemData(j) = Reminder.TimeHours Then
      Exit For
    End If
  Next

  cboHour.ListIndex = j

  For j = cboMinute.listcount - 1 To 1 Step -1
    If cboMinute.ItemData(j) = Reminder.TimeMinutes Then
      Exit For
    End If
  Next
  cboMinute.ListIndex = j


  For j = 0 To 6
    If 2 ^ j And Reminder.DOW Then
      lstDOW.Selected(j) = True
    Else
      lstDOW.Selected(j) = False
    End If
  Next

  For j = lstDOM.listcount - 1 To 1 Step -1
    If lstDOM.ItemData(j) = Reminder.DOM Then
      Exit For
    End If
  Next
  lstDOM.ListIndex = j


  ' future
  'chkCancelled =    iif(Reminder.Cancelled,1,0 )
  'Reminder.OwnerID = 0
  'Reminder.IsPublic = 1
  'txtcoordinator = Reminder.Coordinator
  'Reminder.Disabled = 0
  'Reminder.Cancelled = 0
  'Reminder.Recurring = 0


End Sub

Sub FillLists()
  Dim j As Long

  lstDOW.Clear
  For j = 1 To 7
    lstDOW.AddItem Format(j, "dddd")
    lstDOW.ItemData(lstDOW.NewIndex) = j - 1
  Next
  lstDOW.ListIndex = 0

  lstDOM.Clear
  For j = 1 To 28
    lstDOM.AddItem Format(j)
    lstDOM.ItemData(lstDOM.NewIndex) = j

  Next
  lstDOM.ListIndex = 0

  cboWhen.Clear
  AddToCombo cboWhen, "Inactive", 0
  AddToCombo cboWhen, "Date", 1
  AddToCombo cboWhen, "Daily", 2
  'AddToCombo cboWhen, "Weekly", 3
  AddToCombo cboWhen, "Monthly", 4
  'AddToCombo cboWhen, "Quarterly", 5
  'AddToCombo cboWhen, "Annualy", 6
  cboWhen.ListIndex = 1

  cboHour.Clear


  AddToCombo cboHour, "12 AM", 0
  AddToCombo cboHour, "1 AM", 1
  AddToCombo cboHour, "2 AM", 2
  AddToCombo cboHour, "3 AM", 3
  AddToCombo cboHour, "4 AM", 4
  AddToCombo cboHour, "5 AM", 5
  AddToCombo cboHour, "6 AM", 6
  AddToCombo cboHour, "7 AM", 7
  AddToCombo cboHour, "8 AM", 8
  AddToCombo cboHour, "9 AM", 9
  AddToCombo cboHour, "10 AM", 10
  AddToCombo cboHour, "11 AM", 11
  AddToCombo cboHour, "12 Noon", 12
  AddToCombo cboHour, "1 PM", 13
  AddToCombo cboHour, "2 PM", 14
  AddToCombo cboHour, "3 PM", 15
  AddToCombo cboHour, "4 PM", 16
  AddToCombo cboHour, "5 PM", 17
  AddToCombo cboHour, "6 PM", 18
  AddToCombo cboHour, "7 PM", 19
  AddToCombo cboHour, "8 PM", 20
  AddToCombo cboHour, "9 PM", 21
  AddToCombo cboHour, "10 PM", 22
  AddToCombo cboHour, "11 PM", 23

  cboHour.ListIndex = 0

  cboMinute.Clear
  For j = 0 To 59
    AddToCombo cboMinute, Format(j, "00"), j
  Next
  cboMinute.ListIndex = 0
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  Set Reminder = New cReminder
  SetControls
  FillLists
  ShowPanel TabStrip.SelectedItem.Key
  ResetActivityTime
  'updatescreen
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Do While Busy
    DoEvents
  Loop
  Set Reminder = Nothing
  UnHost
End Sub

Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub

Private Sub lvActivePagers_DblClick()
  UnpageSelected
End Sub

Private Sub lvPagers_DblClick()
  PageSelected
End Sub

Private Sub lvPeople_DblClick()
  SubscribeSelected
End Sub

Private Sub lvPeople_DragDrop(Source As Control, X As Single, Y As Single)
  If Source Is lvPeople Then lvPeople.Drag vbCancel

End Sub

Private Sub lvPeople_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyInsert
      KeyCode = 0
      SubscribeSelected
  End Select
End Sub

Private Sub lvPeople_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lastx = X
  lasty = Y
End Sub

Private Sub lvSubscribers_DblClick()
  UnsubscribeSelected
End Sub

Private Sub lvSubscribers_DragDrop(Source As Control, X As Single, Y As Single)
  If Source Is lvPeople Then
    SubscribeSelected
  End If
End Sub

Private Sub lvSubscribers_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDelete
      KeyCode = 0
      UnsubscribeSelected
    Case Else
  End Select
End Sub

Private Sub lvSubscribers_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
End Sub

Sub SetControls()
  Dim f As Control

  mvCalendar.Value = Now

  For Each f In Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor
    End If
  Next

  framain.left = TabStrip.ClientLeft
  framain.top = TabStrip.ClientTop
  framain.Height = TabStrip.ClientHeight
  framain.Width = TabStrip.ClientWidth

  fraAttendees.left = TabStrip.ClientLeft
  fraAttendees.top = TabStrip.ClientTop
  fraAttendees.Height = TabStrip.ClientHeight
  fraAttendees.Width = TabStrip.ClientWidth

  fraSchedule.left = TabStrip.ClientLeft
  fraSchedule.top = TabStrip.ClientTop
  fraSchedule.Height = TabStrip.ClientHeight
  fraSchedule.Width = TabStrip.ClientWidth



  fraSystem.left = TabStrip.ClientLeft
  fraSystem.top = TabStrip.ClientTop
  fraSystem.Height = TabStrip.ClientHeight
  fraSystem.Width = TabStrip.ClientWidth


  lvSubscribers.ColumnHeaders.Clear
  lvSubscribers.ColumnHeaders.Add , , "Participant", 2500
  lvSubscribers.ColumnHeaders.Add , , "?", 350

  lvPeople.ColumnHeaders.Clear
  lvPeople.ColumnHeaders.Add , , "Name", 2500
  lvPeople.ColumnHeaders.Add , , "?", 350

  lvActivePagers.ColumnHeaders.Clear

  lvActivePagers.ColumnHeaders.Add , , "Active Pagers/Groups", 2500
  lvActivePagers.ColumnHeaders.Add , , "?", 350


  lvPagers.ColumnHeaders.Clear
  lvPagers.ColumnHeaders.Add , , "System Pagers/Groups", 2500
  lvPagers.ColumnHeaders.Add , , "?", 350


End Sub

Sub ShowPanel(ByVal Key As String)

  Select Case LCase(Key)
    Case "attendees"
      fraAttendees.Visible = True
      fraSchedule.Visible = False
      framain.Visible = False
      fraSystem.Visible = False
    Case "schedule"
      fraSchedule.Visible = True
      framain.Visible = False
      fraAttendees.Visible = False
      fraSystem.Visible = False
    Case "system"
      fraSystem.Visible = True
      fraSchedule.Visible = False
      framain.Visible = False
      fraAttendees.Visible = False

    Case Else  ' general
      framain.Visible = True
      fraAttendees.Visible = False
      fraSchedule.Visible = False
      fraSystem.Visible = False
  End Select
End Sub

Sub PageSelected()
  Dim li        As ListItem
  Dim destli    As ListItem
  Dim Key       As String
  On Error Resume Next

  For Each li In lvPagers.ListItems
    If li.Selected Then
      Key = li.Key
      Set destli = Nothing
      Set destli = lvActivePagers.ListItems(Key)
      If destli Is Nothing Then
        Set destli = lvActivePagers.ListItems.Add(, Key, li.text)
        destli.SubItems(1) = li.SubItems(1)
      End If
    End If
  Next

End Sub
Sub SubscribeSelected()
  Dim li        As ListItem
  Dim destli    As ListItem
  Dim Key       As String
  On Error Resume Next

  For Each li In lvPeople.ListItems
    If li.Selected Then
      Key = li.Key
      Set destli = Nothing
      Set destli = lvSubscribers.ListItems(Key)
      If destli Is Nothing Then
        Set destli = lvSubscribers.ListItems.Add(, Key, li.text)
        destli.SubItems(1) = li.SubItems(1)
      End If
    End If
  Next
End Sub

Private Sub TabStrip_Click()
  ShowPanel TabStrip.SelectedItem.Key
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Sub UnpageSelected()
  Dim li As ListItem
  Dim j As Long
  For j = lvActivePagers.ListItems.Count To 1 Step -1
    If lvActivePagers.ListItems(j).Selected Then
      lvActivePagers.ListItems.Remove (j)
    End If
  Next

End Sub


Sub UnsubscribeSelected()
  Dim li As ListItem
  Dim j As Long
  For j = lvSubscribers.ListItems.Count To 1 Step -1
    If lvSubscribers.ListItems(j).Selected Then
      lvSubscribers.ListItems.Remove (j)
    End If
  Next

End Sub

Sub updatescreen()


End Sub
