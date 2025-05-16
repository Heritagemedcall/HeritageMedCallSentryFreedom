VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmRoomEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Room"
   ClientHeight    =   3870
   ClientLeft      =   90
   ClientTop       =   2235
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   9000
      Begin VB.Frame fraAssur 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   30
         TabIndex        =   17
         Top             =   705
         Width           =   7200
         Begin VB.CheckBox chkInAuto 
            Caption         =   "Auto Check-In Same Room Devices"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   420
            TabIndex        =   39
            Top             =   1740
            Width           =   3615
         End
         Begin VB.CommandButton cmdSetAll 
            Caption         =   "Set All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4440
            TabIndex        =   18
            Top             =   225
            Width           =   1175
         End
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "Clear All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5670
            TabIndex        =   19
            Top             =   225
            Width           =   1175
         End
         Begin VB.CheckBox chkTues 
            Alignment       =   1  'Right Justify
            Caption         =   "Tue"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1442
            TabIndex        =   21
            Top             =   780
            Width           =   675
         End
         Begin VB.CheckBox chkWed 
            Alignment       =   1  'Right Justify
            Caption         =   "Wed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2389
            TabIndex        =   22
            Top             =   780
            Width           =   735
         End
         Begin VB.CheckBox chkSun 
            Alignment       =   1  'Right Justify
            Caption         =   "Sun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6120
            TabIndex        =   26
            Top             =   780
            Width           =   675
         End
         Begin VB.CheckBox chkSat 
            Alignment       =   1  'Right Justify
            Caption         =   "Sat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5200
            TabIndex        =   25
            Top             =   780
            Width           =   645
         End
         Begin VB.CheckBox chkFri 
            Alignment       =   1  'Right Justify
            Caption         =   "Fri"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4343
            TabIndex        =   24
            Top             =   780
            Width           =   585
         End
         Begin VB.CheckBox chkThu 
            Alignment       =   1  'Right Justify
            Caption         =   "Thu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3396
            TabIndex        =   23
            Top             =   780
            Width           =   675
         End
         Begin VB.CheckBox chkMon 
            Alignment       =   1  'Right Justify
            Caption         =   "Mon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   465
            TabIndex        =   20
            Top             =   780
            Width           =   705
         End
         Begin VB.CheckBox chkVacation 
            Alignment       =   1  'Right Justify
            Caption         =   "Vacation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   420
            TabIndex        =   27
            Top             =   1305
            Width           =   1335
         End
         Begin VB.TextBox txtUntil 
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
            Left            =   3435
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1350
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblAssurDays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check-in Days"
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
            Left            =   435
            TabIndex        =   30
            Top             =   225
            Width           =   1245
         End
         Begin VB.Label lblUntil 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Until"
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
            Left            =   2940
            TabIndex        =   28
            Top             =   1395
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin VB.Frame fraLocation 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   180
         TabIndex        =   35
         Top             =   630
         Width           =   7320
         Begin VB.TextBox txtLocKW 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   90
            MaxLength       =   254
            TabIndex        =   36
            ToolTipText     =   "Location Keywords Are ONLY Effective for PORTABLE Decices"
            Top             =   630
            Width           =   7035
         End
         Begin VB.Label lblSerial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter a comma separated list of partitions associated with this room"
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
            Index           =   0
            Left            =   450
            TabIndex        =   38
            ToolTipText     =   "Enter Preferred Location Keywords Comma Separated"
            Top             =   1080
            Width           =   5745
         End
         Begin VB.Label lblSerial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location Keywords"
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
            Index           =   1
            Left            =   120
            TabIndex        =   37
            ToolTipText     =   "Enter Preferred Location Keywords Comma Separated"
            Top             =   360
            Width           =   1620
         End
      End
      Begin VB.Frame fraResidents 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   390
         TabIndex        =   6
         Top             =   690
         Width           =   7320
         Begin VB.CommandButton cmdRemoveResident 
            Caption         =   "Remove"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2970
            TabIndex        =   10
            Top             =   2100
            Width           =   1175
         End
         Begin VB.CommandButton cmdEditResident 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1620
            TabIndex        =   9
            Top             =   2100
            Width           =   1175
         End
         Begin VB.CommandButton cmdAddResident 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   255
            TabIndex        =   8
            Top             =   2100
            Width           =   1175
         End
         Begin MSComctlLib.ListView lvResidents 
            Height          =   1935
            Left            =   90
            TabIndex        =   7
            Top             =   120
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   3413
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
      Begin VB.Frame fraTransmitters 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   150
         TabIndex        =   11
         Top             =   855
         Width           =   7395
         Begin VB.CommandButton cmdAssignRes 
            Caption         =   "Assign Resident"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   4515
            TabIndex        =   31
            Top             =   1800
            Width           =   1175
         End
         Begin VB.CommandButton cmdDeleteTransmitter 
            Caption         =   "Remove Transmitter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1825
            TabIndex        =   15
            Top             =   1800
            Width           =   1175
         End
         Begin VB.CommandButton cmdEditDevice 
            Caption         =   "Configure Transmitter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   3170
            TabIndex        =   14
            Top             =   1800
            Width           =   1175
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add Transmitter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   480
            TabIndex        =   13
            Top             =   1800
            Width           =   1175
         End
         Begin MSComctlLib.ListView lvDevices 
            Height          =   1620
            Left            =   15
            TabIndex        =   12
            Top             =   15
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   2858
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
      Begin VB.CommandButton cmdAddRoom 
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
         TabIndex        =   16
         Top             =   30
         Width           =   1175
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
         Height          =   585
         Left            =   7725
         TabIndex        =   33
         Top             =   2370
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
         Left            =   7725
         TabIndex        =   32
         Top             =   1785
         Width           =   1175
      End
      Begin VB.TextBox txtBuilding 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4785
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2940
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.TextBox txtRoom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   2
         Top             =   0
         Width           =   2000
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2940
         Left            =   30
         TabIndex        =   5
         Top             =   360
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   5186
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Transmitters"
               Key             =   "tx"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Check-in"
               Key             =   "assur"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Location"
               Key             =   "location"
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
      Begin VB.Label lblError 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3480
         TabIndex        =   34
         Top             =   60
         Width           =   75
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Building"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3615
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblRoom 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   1
         Top             =   60
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmRoomEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CallingForm  As String
Public RoomID       As Long
Private Room        As New cRoom
Private mAssurDays  As Long
Private Vacation    As Long

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
Function Save() As Boolean
  Dim Rs      As Recordset
  Dim Room    As cRoom: Set Room = New cRoom
  
  txtBuilding.text = Trim(txtBuilding.text)
  txtRoom.text = Trim(txtRoom.text)
  
  Room.Room = txtRoom.text
  Room.RoomID = RoomID
  Room.Building = ""
  Room.Assurdays = GetAssurDays()
  Room.Away = chkVacation.value
  Room.Vacation = 0
  Room.Deleted = 0
  Room.locKW = Trim$(txtLocKW.text)
  Room.flags = chkInAuto.value And 1
  
  If MASTER Then
    Save = SaveRoom(Room, gUser.Username)
    RoomID = Room.RoomID
  Else
    
    If ClientUpdateRoom(Room) Then
      RoomID = Room.RoomID
      dbgGeneral "ClientUpdateRoom called, RoomID = " & RoomID
      Save = True
    Else
      Save = False
    End If
  End If

  Fill

End Function

Private Property Let AssurDay(ByVal index As Long, ByVal value As Long)
'index (bit) 1 thur 7
'index 1 is monday
'index 2 is tues
' Value is either 1 or 0 (on or off)

  value = IIf(value = 0, 0, 1)
  If index > 0 And index < 8 Then  ' only 1 thru 7 ' bit 0 is reserved
    If value = 1 Then  ' set the bit
      mAssurDays = mAssurDays Or (2 ^ index)
    Else
      mAssurDays = mAssurDays And (Not (2 ^ index))
    End If
  End If

End Property
Function FillAssurDays() As Long
  chkSun.value = IIf(AssurDay(1) = 0, 0, 1)
  chkMon.value = IIf(AssurDay(2) = 0, 0, 1)
  chkTues.value = IIf(AssurDay(3) = 0, 0, 1)
  chkWed.value = IIf(AssurDay(4) = 0, 0, 1)
  chkThu.value = IIf(AssurDay(5) = 0, 0, 1)
  chkFri.value = IIf(AssurDay(6) = 0, 0, 1)
  chkSat.value = IIf(AssurDay(7) = 0, 0, 1)

End Function

Private Property Get AssurDay(ByVal index As Long) As Long
  If index > 0 And index < 8 Then  ' only 1 thru 7 ' bit 0 is reserved
    AssurDay = IIf((mAssurDays And (2 ^ index)) = 0, 0, 1)
  End If
End Property


Function GetAssurDays() As Long
  mAssurDays = 0
  AssurDay(1) = chkSun.value
  AssurDay(2) = chkMon.value
  AssurDay(3) = chkTues.value
  AssurDay(4) = chkWed.value
  AssurDay(5) = chkThu.value
  AssurDay(6) = chkFri.value
  AssurDay(7) = chkSat.value
  GetAssurDays = mAssurDays And &HFF
End Function

Sub ConfigureViews()


  Dim ch   As ColumnHeader
  
  If lvDevices.ColumnHeaders.Count < 5 Then
    lvDevices.ColumnHeaders.Clear
    lvDevices.Sorted = True
    Set ch = lvDevices.ColumnHeaders.Add(, "S", "Serial", 1100)
    Set ch = lvDevices.ColumnHeaders.Add(, "M", "Model", 1200)
    Set ch = lvDevices.ColumnHeaders.Add(, "Res", "Resident", 2500)
    Set ch = lvDevices.ColumnHeaders.Add(, "Phon", "Phone", 1440)
    Set ch = lvDevices.ColumnHeaders.Add(, "Assur", "Assur", 700)
  End If

  lvResidents.ColumnHeaders.Clear
  Set ch = lvResidents.ColumnHeaders.Add(, , "Name", 2400)
  Set ch = lvResidents.ColumnHeaders.Add(, , "Phone", 2400)
  


End Sub
Private Sub ClearForm()
  AssurDay(1) = 1
  AssurDay(2) = 1
  AssurDay(3) = 1
  AssurDay(4) = 1
  AssurDay(5) = 1
  AssurDay(6) = 1
  AssurDay(7) = 1

  ''mAssurDays = 0
  Vacation = 0
  txtBuilding.text = ""
  txtRoom.text = ""
  txtLocKW.text = ""
  lvDevices.ListItems.Clear
  lvResidents.ListItems.Clear
  chkInAuto.value = 0
End Sub

Public Sub Fill()
  Dim SQL As String
  Dim Rs  As Recordset

  RefreshJet
  ClearForm

  If MASTER Then

    dbgGeneral "filling RoomID " & RoomID

    Set Room = New cRoom
    Room.RoomID = RoomID

    SQL = "select * from rooms where roomid =" & RoomID

    Set Rs = ConnExecute(SQL)
    If Not Rs.EOF Then
      dbgGeneral "rs.eof " & Rs.EOF
      'Room.Building = rs("Building") & ""
      Room.Room = Rs("room") & ""
      txtRoom.text = Rs("room") & ""
      mAssurDays = Rs("AssurDays")
      Vacation = Rs("Away")
      Room.RoomID = RoomID
      txtLocKW.text = Trim$(Rs("lockw") & "")
      Room.locKW = Trim$(Rs("lockw") & "")
      Room.flags = Val(Rs("flags") & "")
      chkInAuto.value = Room.flags And 1
      
      
      dbgGeneral "filling room " & Room.Room
    Else
      dbgGeneral "rs.eof " & Rs.EOF
    End If
    Rs.Close
    Set Rs = Nothing
  Else
    On Error GoTo 0
    Dim temproom As cRoom
    Set temproom = New cRoom
    Remote_GetRoom temproom, RoomID

    Room.Room = temproom.Room
    txtRoom.text = temproom.Room
    mAssurDays = temproom.Assurdays
    Vacation = temproom.Away
    Room.RoomID = temproom.RoomID
    txtLocKW.text = temproom.locKW  ' Trim$(rs("lockw") & "")
    chkInAuto.value = temproom.flags And 1
    RoomID = temproom.RoomID
  End If

  chkVacation.value = IIf(Vacation = 1, 1, 0)
  FillAssurDays


  ' get transmitters specifically for this room
  RefreshlvDevices

  ' get residents
  RefreshlvResidents

End Sub

Sub RefreshlvDevices()

  Dim SQL As String
  Dim Rs  As Recordset
  Dim li As ListItem
  Dim d As cESDevice
  Dim j  As Integer
  Dim rsres As Recordset
  lvDevices.ListItems.Clear
'  sql = " SELECT Devices.DeviceID, Devices.Serial, Devices.model,Devices.midpti " & _
'      " FROM Devices " & _
'      " WHERE RoomID <> 0 and residentid = 0 and roomid = " & roomid

  SQL = " SELECT devices.useassur, devices.model , devices.useassur2, devices.assurinput,devices.deviceid, Devices.Serial , Devices.residentid FROM Devices WHERE RoomID <> 0 and roomid = " & RoomID


  Set Rs = ConnExecute(SQL)
  
  Do Until Rs.EOF
    
    Set d = New cESDevice
    d.Serial = Rs("serial")
    d.UseAssur = Rs("UseAssur")
    d.UseAssur2 = Rs("UseAssur2")
    d.AssurInput = Rs("AssurInput")
    d.DeviceID = Rs("deviceid")
    d.Model = Rs("model") & ""
    d.ResidentID = IIf(IsNull(Rs("residentID")), "0", Rs("residentID"))
    
    If Not d Is Nothing Then


      Dim assure As String
      
      Set li = lvDevices.ListItems.Add(, d.DeviceID & "D", d.Serial)
      li.SubItems(1) = d.Model
      
      SQL = "select * from residents where residentID = " & d.ResidentID
      Set rsres = ConnExecute(SQL)
      If Not rsres.EOF Then
        li.SubItems(2) = ConvertLastFirst(rsres("namelast") & "", rsres("namefirst") & "")
        li.SubItems(3) = rsres("phone") & ""
      Else
        li.SubItems(2) = " "
        li.SubItems(3) = " "
      
      End If
      rsres.Close
      'If d.NumInputs > 1 Then
        
        assure = IIf(d.UseAssur = 1, "Y", "N") & IIf(d.UseAssur2 = 1, "Y", "N")
        If d.UseAssur = 1 Or d.UseAssur2 = 1 Then
          assure = assure & d.AssurInput
        End If
      'Else
       ' assure = IIf(d.UseAssur = 1, "Y", "N")
      'End If
      li.SubItems(4) = assure

    End If
    Rs.MoveNext
  Loop
  Rs.Close

End Sub
Sub RefreshlvResidents()
  Dim SQL As String
  Dim Rs  As Recordset
  Dim li As ListItem

  Exit Sub

  lvResidents.ListItems.Clear
  If RoomID <> 0 Then
    SQL = "SELECT * FROM residents WHERE roomID <> 0 AND roomid = " & RoomID

    Set Rs = ConnExecute(SQL)
    Do Until Rs.EOF
      Set li = lvResidents.ListItems.Add(, Rs("residentid") & "R", ConvertLastFirst(Rs("namelast") & "", Rs("namefirst") & ""))
      li.SubItems(1) = Rs("phone") & "" ' phone
      Rs.MoveNext
    Loop
    Rs.Close
  End If

End Sub



Private Sub cmdAdd_Click()
  If RoomID <> 0 Then
    ShowTransmitters 0, RoomID
  Else
    Beep
  End If
End Sub

Private Sub cmdAddResident_Click()
  If RoomID <> 0 Then
    ShowResidents RoomID, 0, ""
  End If
End Sub

Private Sub cmdAddRoom_Click()
  RoomID = 0
  Fill

End Sub

Private Sub cmdAssignRes_Click()
  
  Dim ResID As Long
  Dim DeviceID As Long
  
  
  If RoomID <> 0 Then
    
    If lvDevices.SelectedItem Is Nothing Then
      Beep
    Else
      DeviceID = Val(lvDevices.SelectedItem.Key)
      
      ResID = GetResidentIDFromDeviceID(DeviceID)
      
      'Set d = Devices.Device(lvDevices.SelectedItem.text)
      'If d Is Nothing Then
      '  Beep
      'Else
        ShowResidents ResID, DeviceID, "Room Edit"
      'End If
    End If
  Else
    Beep
  End If
End Sub

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
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

Private Sub cmdDeleteTransmitter_Click()
  DeleteTransmitter
End Sub

Private Sub cmdEditDevice_Click()
  ConfigureTX
End Sub
Private Sub ConfigureTX()
  If Not lvDevices.SelectedItem Is Nothing Then
    EditTransmitter GetSelectedTransmitter()
    Fill
  Else
    Beep
  End If
  
  

End Sub
Private Sub cmdEditResident_Click()
  EditResident GetSelectedResident()

End Sub

Function GetSelectedTransmitter() As Long
  If Not lvDevices.SelectedItem Is Nothing Then
    GetSelectedTransmitter = Val(lvDevices.SelectedItem.Key)
  End If
End Function
Function GetSelectedResident() As Long
  If Not lvResidents.SelectedItem Is Nothing Then
    GetSelectedResident = Val(lvResidents.SelectedItem.Key)
  End If
End Function

Private Sub cmdOK_Click()
  cmdOK.Enabled = False
  RefreshJet
  If Validate() Then
    Save
    
  Else
    Beep
  End If
  RefreshJet
  cmdOK.Enabled = True
End Sub

Private Function Validate() As Boolean

  txtRoom.text = Trim(txtRoom.text)
  If Len(txtRoom.text) = 0 Then
    Validate = False
    Exit Function
  End If

  

  'If MASTER Then

    If RoomID = 0 Then
      If GetRoomByName(txtRoom.text) <> 0 Then
        Beep
        Validate = False
        txtRoom.ForeColor = vbRed
        lblError.Caption = "Error: Duplicate Room"
        Exit Function
      Else
        Validate = True
      End If
    Else
      Validate = True
    End If
  'Else
  '  Dim temproom As cRoom:    Set temproom = New cRoom
    
    
  '  Set temproom = Nothing
  'End If

End Function
Private Sub cmdRemoveResident_Click()
  RemoveResident
End Sub

Private Sub Command2_Click()

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
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select

End Sub

Private Sub Form_Load()
ResetActivityTime
  ArrangeControls
  Connect
  ConfigureViews
End Sub
Sub ArrangeControls()
  fraEnabler.BackColor = Me.BackColor

  fraassur.left = TabStrip.ClientLeft
  fraassur.top = TabStrip.ClientTop
  fraassur.Height = TabStrip.ClientHeight
  fraassur.Width = TabStrip.ClientWidth
  fraassur.BackColor = Me.BackColor

  fraResidents.left = TabStrip.ClientLeft
  fraResidents.top = TabStrip.ClientTop
  fraResidents.Height = TabStrip.ClientHeight
  fraResidents.Width = TabStrip.ClientWidth
  fraResidents.BackColor = Me.BackColor

  fraTransmitters.left = TabStrip.ClientLeft
  fraTransmitters.top = TabStrip.ClientTop
  fraTransmitters.Height = TabStrip.ClientHeight
  fraTransmitters.Width = TabStrip.ClientWidth
  fraTransmitters.BackColor = Me.BackColor

  
  fraLocation.left = TabStrip.ClientLeft
  fraLocation.top = TabStrip.ClientTop
  fraLocation.Height = TabStrip.ClientHeight
  fraLocation.Width = TabStrip.ClientWidth
  fraLocation.BackColor = Me.BackColor
  
  

  fraTransmitters.Visible = True
  fraResidents.Visible = False
  fraLocation.Visible = False
  fraassur.Visible = False

  lblError.Caption = ""

End Sub

Sub DeleteTransmitter()
  Dim SQL       As String
  Dim DeviceID  As Long

  DeviceID = GetSelectedTransmitter()
  If DeviceID <> 0 Then
    If vbYes = messagebox(Me, "Remove Selected Transmitter?", App.Title, vbYesNo Or vbQuestion) Then
      SQL = "UPDATE devices Set RoomID = 0 WHERE DeviceID = " & DeviceID
      ConnExecute SQL
      Devices.RefreshByID DeviceID
      RefreshlvDevices
    End If
  Else
    Beep
  End If


End Sub

Sub RemoveResident()
  Dim SQL       As String
  Dim ResidentID  As Long

  ResidentID = GetSelectedResident()
  If ResidentID <> 0 Then
    If vbYes = messagebox(Me, "Remove Selected Resident?", App.Title, vbYesNo Or vbQuestion) Then
      SQL = "UPDATE Residents Set RoomID = 0 WHERE ResidentID = " & ResidentID
      ConnExecute SQL
      RefreshlvResidents
    End If
  End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
  lblError.Caption = ""
  UnHost
End Sub

Private Sub TabStrip_Click()
  Select Case TabStrip.SelectedItem.Key
    Case "location"
      fraLocation.Visible = True
      fraResidents.Visible = False
      fraTransmitters.Visible = False
      fraassur.Visible = False
    
    
    Case "assur"
      fraassur.Visible = True
      fraResidents.Visible = False
      fraTransmitters.Visible = False
      fraLocation.Visible = False
    Case "res"

      fraResidents.Visible = True
      fraassur.Visible = False
      fraTransmitters.Visible = False
      fraLocation.Visible = False

    Case Else
      fraTransmitters.Visible = True
      fraassur.Visible = False
      fraResidents.Visible = False
      fraLocation.Visible = False

  End Select

End Sub

Private Sub txtBuilding_GotFocus()
  SelAll txtBuilding
End Sub

Private Sub txtRoom_Change()
  txtRoom.ForeColor = vbBlack
  lblError.Caption = ""
End Sub

Private Sub txtRoom_GotFocus()
  SelAll txtRoom
End Sub
