VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    "
   ClientHeight    =   10935
   ClientLeft      =   165
   ClientTop       =   1935
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   15240
   Begin VB.PictureBox picCommError 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   3180
      ScaleHeight     =   2625
      ScaleWidth      =   4965
      TabIndex        =   82
      Top             =   5040
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Frame fraExternal 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   1425
      TabIndex        =   72
      Top             =   11550
      Visible         =   0   'False
      Width           =   8835
      Begin MSComctlLib.ListView lvExternal 
         Height          =   2130
         Left            =   -30
         TabIndex        =   73
         Top             =   285
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3757
         View            =   3
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tx ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label lblExternal 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "External Alarms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   9045
      End
   End
   Begin VB.Frame fraAlerts 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   1245
      TabIndex        =   16
      Top             =   2505
      Width           =   8835
      Begin MSComctlLib.ListView lvAlerts 
         Height          =   2130
         Left            =   0
         TabIndex        =   18
         Top             =   300
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3757
         View            =   3
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tx ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label lblAlerts 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alerts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   9045
      End
   End
   Begin VB.CommandButton CmdEditInfo 
      Caption         =   "Edit Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5520
      Width           =   720
   End
   Begin VB.CommandButton cmdClearInfo 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6210
      Width           =   720
   End
   Begin VB.Frame fraResinfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   1245
      TabIndex        =   28
      Top             =   5055
      Width           =   8970
      Begin VB.PictureBox picResident 
         BackColor       =   &H00FFFFFF&
         Height          =   2310
         Left            =   0
         ScaleHeight     =   2250
         ScaleWidth      =   8805
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   8865
         Begin VB.TextBox txtAssurDays 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   390
            Width           =   2070
         End
         Begin VB.ListBox lstAssurDevs 
            Appearance      =   0  'Flat
            Height          =   1275
            IntegralHeight  =   0   'False
            Left            =   5685
            TabIndex        =   71
            Top             =   915
            Width           =   1425
         End
         Begin VB.TextBox txtHiddenAlarmID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6630
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   465
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtInfoNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1290
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   915
            Width           =   5520
         End
         Begin VB.TextBox txtHiddenRoomID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6630
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CommandButton cmdChangeVacation 
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
            Height          =   540
            Left            =   7305
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1245
         End
         Begin VB.TextBox txtHiddenResID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6630
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   255
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtInfox 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   585
            Width           =   3645
         End
         Begin VB.TextBox txtInfoRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   105
            Width           =   2070
         End
         Begin VB.TextBox txtInfoMessage 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   345
            Width           =   3645
         End
         Begin VB.TextBox txtInfoFullName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   105
            Width           =   3645
         End
         Begin VB.Label lblAssrDayList 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Assur"
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
            Left            =   3675
            TabIndex        =   76
            Top             =   390
            Width           =   480
         End
         Begin VB.Label lblInfoRoom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   3675
            TabIndex        =   32
            Top             =   105
            Width           =   495
         End
         Begin VB.Image imgResPic 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1500
            Left            =   7200
            Stretch         =   -1  'True
            Top             =   60
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.Label lblInformation 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   9330
      End
   End
   Begin VB.Frame fraControlPanel 
      BorderStyle     =   0  'None
      Height          =   2910
      Left            =   11280
      TabIndex        =   55
      Top             =   7980
      Width           =   3945
      Begin VB.CommandButton cmdAnnouncements 
         Caption         =   "Announce"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdRooms 
         Caption         =   "Rooms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":09B8
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdResidents 
         Caption         =   "Residents"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":0FAE
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdTransmitters 
         Caption         =   "Transmitters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":1600
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   795
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdOutputs 
         Caption         =   "Output groups"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2490
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":1C44
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":2180
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   795
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdReports 
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":23F0
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdOutputDevices 
         Caption         =   "Outputs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2490
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":2B2E
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   795
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdOutputServers 
         Caption         =   "Output servers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2490
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":305A
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.Image imgPacket 
         Height          =   255
         Left            =   240
         Top             =   2550
         Width           =   255
      End
      Begin VB.Image imgUptime 
         Height          =   255
         Left            =   1020
         Picture         =   "frmMain.frx":365A
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "         "
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
         Left            =   3135
         TabIndex        =   65
         Top             =   2595
         Width           =   555
      End
   End
   Begin VB.Frame fraAssur 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   1245
      TabIndex        =   42
      Top             =   7680
      Visible         =   0   'False
      Width           =   9945
      Begin VB.CommandButton cmdAssurExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9045
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2085
         Width           =   720
      End
      Begin VB.CommandButton cmdAsurrDown 
         Caption         =   "down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9045
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   900
         Width           =   720
      End
      Begin VB.CommandButton cmdAssurUp 
         Caption         =   "up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9045
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdAssurPrintList 
         Caption         =   "Print List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9045
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1485
         Width           =   720
      End
      Begin MSComctlLib.ListView lvAssur 
         Height          =   2430
         Left            =   0
         TabIndex        =   44
         Top             =   285
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4286
         View            =   3
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Room"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label lblAssur 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check-ins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   9045
      End
   End
   Begin VB.CommandButton cmdMultiPrint 
      Caption         =   "Print List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3945
      Width           =   720
   End
   Begin VB.CommandButton cmdAlarmPrintList 
      Caption         =   "Print List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1515
      Width           =   720
   End
   Begin VB.CommandButton cmdPrintInfo 
      Caption         =   "Print Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6900
      Width           =   720
   End
   Begin VB.Frame fraHost 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   1245
      TabIndex        =   40
      Top             =   7620
      Width           =   10005
      Begin VB.Label lblHostPanel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Host Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4035
         TabIndex        =   41
         Top             =   1545
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdMultiUp 
      Caption         =   "up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2805
      Width           =   720
   End
   Begin VB.CommandButton cmdMultiDown 
      Caption         =   "down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3375
      Width           =   720
   End
   Begin VB.Frame fraLoBatt 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2490
      Left            =   1245
      TabIndex        =   22
      Top             =   3765
      Width           =   8835
      Begin MSComctlLib.ListView lvLoBatt 
         Height          =   2130
         Left            =   0
         TabIndex        =   24
         Top             =   300
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3757
         View            =   3
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tx ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Low Battery Time"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label lblLowBattery 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Low Battery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   9045
      End
   End
   Begin VB.CommandButton cmdAlarmdown 
      Caption         =   "down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   945
      Width           =   720
   End
   Begin VB.CommandButton cmdAlarmUp 
      Caption         =   "up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10230
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   375
      Width           =   720
   End
   Begin VB.Frame fraCheckin 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   1245
      TabIndex        =   25
      Top             =   3540
      Width           =   8835
      Begin MSComctlLib.ListView lvCheckIn 
         Height          =   2130
         Left            =   0
         TabIndex        =   27
         Top             =   300
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3757
         View            =   3
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tx ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Checkin"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label lblTrouble 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trouble"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   9045
      End
   End
   Begin VB.PictureBox picAlarms 
      BorderStyle     =   0  'None
      Height          =   10755
      Left            =   0
      ScaleHeight     =   717
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1200
      Begin VB.Timer TimerLogon 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   210
         Top             =   10530
      End
      Begin VB.CommandButton cmdExternal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "External"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FF0000&
         Picture         =   "frmMain.frx":3A65
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3420
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrintScreen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Help File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":442B
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   9360
         Width           =   1050
      End
      Begin VB.CommandButton cmdAssur 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check-ins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FF0000&
         Picture         =   "frmMain.frx":48A1
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   6750
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdTrouble 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Trouble"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":4F55
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4530
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdBattery 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Battery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":58A1
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdAlert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":5F49
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2310
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdAck 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":6797
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdSilence 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Silence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":7157
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8250
         Width           =   1050
      End
      Begin VB.TextBox txtLogin 
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
         IMEMode         =   3  'DISABLE
         Left            =   0
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   7860
         Width           =   1140
      End
   End
   Begin MSComctlLib.ListView lvEmergency 
      Height          =   2130
      Left            =   1245
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   330
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   3757
      View            =   3
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tx ID"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Time"
         Object.Width           =   5010
      EndProperty
   End
   Begin VB.Frame fraLocate 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   12705
      TabIndex        =   66
      Top             =   10965
      Width           =   9090
   End
   Begin VB.Frame fraHippaList 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   7440
      Left            =   11040
      TabIndex        =   77
      Top             =   150
      Width           =   4155
      Begin VB.CommandButton cmdResListDown 
         Caption         =   "down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2220
         Width           =   720
      End
      Begin VB.CommandButton cmdResListUp 
         Caption         =   "up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1650
         Width           =   720
      End
      Begin VB.CommandButton cmdResListPrintList 
         Caption         =   "Print List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4995
         Width           =   720
      End
      Begin VB.CommandButton cmdResListPrintView 
         Caption         =   "Print View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4425
         Width           =   720
      End
      Begin VB.CommandButton cmdResListPgUp 
         Caption         =   "page up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3090
         Width           =   720
      End
      Begin VB.CommandButton cmdResListPgDown 
         Caption         =   "page down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3660
         Width           =   720
      End
      Begin VB.Frame fraTransmitters 
         BorderStyle     =   0  'None
         Height          =   6825
         Left            =   90
         TabIndex        =   78
         Top             =   375
         Width           =   3120
         Begin MSComctlLib.ListView lvtx 
            Height          =   6720
            Left            =   30
            TabIndex        =   81
            Top             =   90
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   11853
            View            =   3
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Model"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Serial"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.Image imgEditTxRes 
            Height          =   285
            Left            =   2325
            Picture         =   "frmMain.frx":79B3
            Stretch         =   -1  'True
            Top             =   6885
            Width           =   285
         End
         Begin VB.Image imgDelTx 
            Height          =   285
            Left            =   2670
            Picture         =   "frmMain.frx":850F
            Top             =   6885
            Width           =   285
         End
         Begin VB.Image imgEditTx 
            Height          =   285
            Left            =   1965
            Picture         =   "frmMain.frx":8B19
            Stretch         =   -1  'True
            Top             =   6885
            Width           =   285
         End
      End
      Begin MSComctlLib.TabStrip tabList 
         Height          =   7380
         Left            =   60
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   0
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   13018
         TabFixedWidth   =   2290
         TabFixedHeight  =   526
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Residents"
               Key             =   "res"
               Object.ToolTipText     =   "Sort By Resident"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Rooms"
               Key             =   "room"
               Object.ToolTipText     =   "Sort By Room"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Transmitters"
               Key             =   "tx"
               Object.ToolTipText     =   "Sort By Transmitter"
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
      Begin VB.Frame fraPagers 
         BorderStyle     =   0  'None
         Height          =   6780
         Left            =   105
         TabIndex        =   80
         Top             =   375
         Width           =   3165
      End
   End
   Begin VB.Label lblAlarms 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alarms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1245
      TabIndex        =   11
      Top             =   30
      Width           =   8865
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ShowCountdown As Boolean

Private lvtxbusy As Boolean
Private mCurrentGrid      As Integer
Private mCurrentSubGrid   As Integer

Private Resident As cResident
Private CurrentResident As cResident

Private DeletingAlarm As Boolean

Private mAssur   As Boolean

Enum InfoPane
  emergency
  Alert
  Checkin
  LowBatt
  extern
End Enum

Const SORT_RESIDENT = 0
Const SORT_SERIAL = 1
Const SORT_ROOM = 2

Private lvtxCols_Res As cGridColumns
Private lvtxCols_Serial As cGridColumns
Private lvtxCols_Room As cGridColumns


Const LED_AMBER = 1006
Const LED_GREEN = 101
Const LED_GRAY = 103
Const LED_RED = 1008


Const ColorActive = &H80000002
Const ColorInActive = &H8000000F

Const BADPACKET_WARNING_CAUTION = 5
Const BADPACKET_WARNING_DANGER = 10

Const LVEMERGENCY_ID = 0
Const LVEMERGENCY_NAME = 1
Const LVEMERGENCY_ROOM = 2
Const LVEMERGENCY_LOCATION = 3
Const LVEMERGENCY_TIME = 4
Const LVEMERGENCY_ANNOUNCE = 5
Const LVEMERGENCY_ACK = 6
Const LVEMERGENCY_RESP = 7


Const LVALERT_ID = 0
Const LVALERT_NAME = 1
Const LVALERT_ROOM = 2
Const LVALERT_LOCATION = 3
Const LVALERT_TIME = 4
Const LVALERT_ANNOUNCE = 5
Const LVALERT_ACK = 6
Const LVALERT_RESP = 7



Sub RefreshAlarms(ByVal XML As String)


  ' used only by remotes !!!!!!!!



  Dim Node               As IXMLDOMNode
  Dim NodeList           As IXMLDOMNodeList
  Dim childnode          As IXMLDOMNode
  Dim NewAlarms          As cAlarms
  Dim alarm              As cAlarm
  Dim BeepTimer          As Long

  Dim AlarmTime          As Double
  Dim SilenceTime        As Double

  Dim doc                As DOMDocument60

  If Len(XML) > 0 Then

starthere:

    On Error GoTo RefreshAlarms_Error
    Set doc = New DOMDocument60
    doc.LoadXML XML

    '****** ALARMS

    Set NewAlarms = New cAlarms


    Set Node = doc.selectSingleNode("HMC/Alarms/Beep")
    BeepTimer = 0
    If Not Node Is Nothing Then
      BeepTimer = Val(Node.text)
    End If


    'check parsing here

    AlarmTime = 0
    Set Node = doc.selectSingleNode("HMC/Alarms/AlarmTime")
    If Not Node Is Nothing Then
      AlarmTime = Val(Node.text)
    End If

    SilenceTime = 0
    Set Node = doc.selectSingleNode("HMC/Alarms/SilenceTime")
    If Not Node Is Nothing Then
      SilenceTime = Val(Node.text)
    End If

    HostTime = 0
    Set Node = doc.selectSingleNode("HMC/Alarms/HostTime")
    If Not Node Is Nothing Then
      HostTime = Val(Node.text)
      Debug.Print "Host Time " & CDate(HostTime)
      If HostTime > 42370 Then  ' 1/1/2016
        If Configuration.SyncHostTime Then  ' if true then sync time if more than 1 minute off
          If Abs(DateDiff("s", Now, CDate(HostTime))) >= 60 Then
            SyncTime
          End If
        End If
      End If
    End If


    Set NodeList = doc.selectNodes("HMC/Alarms/Alarm")
    If Not (NodeList Is Nothing) Then
      For Each childnode In NodeList
        Set alarm = New cAlarm
        For Each Node In childnode.childnodes
          Select Case LCase(Node.baseName)
            Case "resident"
              If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
                alarm.ResidentText = ""
              Else
                alarm.ResidentText = Node.text
              End If
            Case "eventtype"
              alarm.EventType = Val(Node.text)
            Case "roomtext"
              alarm.RoomText = Node.text
            Case "roomid"
              alarm.RoomID = Val(Node.text)
            Case "model"
              alarm.Model = Node.text
            Case "serial"      ' hex B20E087E
              alarm.Serial = Node.text
            Case "id"
              alarm.ID = Val(Node.text)
            Case "alarmid"
              alarm.AlarmID = Val(Node.text)
            Case "inputnum"    ' 1
              alarm.Inputnum = Val(Node.text)
            Case "locationtext"  ' Lobby
              alarm.locationtext = XMLDecode(Node.text)
            Case "datetime"    ' 02/05/07 2:50:55 PM
              alarm.DateTime = Node.text
            Case "announce"
              alarm.Announce = XMLDecode(Node.text)
            Case "acked"
              alarm.ACKed = Node.text
            Case "acktime"     ' 12:00:00 AM
              alarm.ACKTime = Node.text
            Case "responder"     ' 12:00:00 AM
              alarm.Responder = Node.text
            Case "priorid"
              alarm.PriorID = Val(Node.text)
            Case "alarmtype"   ' 1
              alarm.Alarmtype = Val(Node.text)
            Case "timestamp"   ' 1
              alarm.TimeStamp = Val(Node.text)
            Case "silenced"    ' 0?
              alarm.Silenced = Val(Node.text)
            Case "silencetime"  ' 02/05/07 2:52:02 PM
              alarm.SilenceTime = Node.text
            Case "description"
              alarm.Description = Node.text
            Case "custom"
              alarm.Custom = Node.text
          End Select
        Next

        NewAlarms.alarms.Add alarm
        
      Next

SetBeepTimer:

      NewAlarms.BeepTimer = BeepTimer

    End If

    Set alarms = NewAlarms
    If alarms.Count Then
      Debug.Print "frmMain.RefreshAlarms AlarmTime " & IIf(AlarmTime, CDate(AlarmTime), "0")
      Debug.Print "frmMain.RefreshAlarms SilenceTime " & IIf(SilenceTime, CDate(SilenceTime), "0")

      'If AlarmTime = 0 Then Stop
      alarms.LocalAlarmTime = AlarmTime
      alarms.LocalSilenceTime = SilenceTime
    End If


    ProcessAlarms



    '****** ALERTS

    Set NewAlarms = New cAlarms

    Set Node = doc.selectSingleNode("HMC/Alerts/Beep")
    BeepTimer = 0
    If Not Node Is Nothing Then
      BeepTimer = Val(Node.text)
    End If


    Set Node = doc.selectSingleNode("HMC/Alerts/AlarmTime")
    If Not Node Is Nothing Then
      AlarmTime = Val(Node.text)
    End If

    SilenceTime = 0
    Set Node = doc.selectSingleNode("HMC/Alerts/SilenceTime")
    If Not Node Is Nothing Then
      SilenceTime = Val(Node.text)
    End If



    Set NodeList = doc.selectNodes("HMC/Alerts/Alarm")
    If Not (NodeList Is Nothing) Then
      For Each childnode In NodeList
        Set alarm = New cAlarm
        For Each Node In childnode.childnodes
          Select Case LCase(Node.baseName)

            Case "resident"
              alarm.ResidentText = Node.text
            Case "roomtext"
              alarm.RoomText = Node.text
            Case "roomid"
              alarm.RoomID = Val(Node.text)

            Case "serial"      ' B20E087E
              alarm.Serial = Node.text
            Case "id"
              alarm.ID = Val(Node.text)
            Case "alarmid"
              alarm.AlarmID = Val(Node.text)
  
            Case "inputnum"    ' 1
              alarm.Inputnum = Val(Node.text)
            Case "locationtext"  ' Lobby
              alarm.locationtext = XMLDecode(Node.text)
            Case "datetime"    ' 02/05/07 2:50:55 PM
              alarm.DateTime = Node.text
            Case "announce"
              alarm.Announce = XMLDecode(Node.text)
            Case "acked"
              alarm.ACKed = Node.text
            Case "acktime"     ' 12:00:00 AM
              alarm.ACKTime = Node.text
            Case "responder"     ' 12:00:00 AM
              alarm.Responder = Node.text
            Case "priorid"
              alarm.PriorID = Val(Node.text)
            Case "alarmtype"   ' 1
              alarm.Alarmtype = Val(Node.text)
            Case "timestamp"   ' 1
              alarm.TimeStamp = Val(Node.text)
            Case "silenced"    ' 0?
              alarm.Silenced = Val(Node.text)
            Case "silencetime"  ' 02/05/07 2:52:02 PM
              alarm.SilenceTime = Node.text
            Case "description"
              alarm.Description = Node.text
            Case "custom"
              alarm.Custom = Node.text
          End Select
        Next
        NewAlarms.alarms.Add alarm
      Next
      NewAlarms.BeepTimer = BeepTimer
    End If

    Set Alerts = NewAlarms
    
    If Alerts.Count Then
      Alerts.LocalAlarmTime = AlarmTime
      Alerts.LocalSilenceTime = SilenceTime
    End If
    
    ProcessAlerts


    '****** TROUBLES

    Set NewAlarms = New cAlarms
    Set Node = doc.selectSingleNode("HMC/Troubles/Beep")
    BeepTimer = 0
    If Not Node Is Nothing Then
      BeepTimer = Val(Node.text)
    End If

    Set Node = doc.selectSingleNode("HMC/Troubles/SilenceTime")
    SilenceTime = 0
    If Not Node Is Nothing Then
      SilenceTime = Val(Node.text)
    End If

    Set Node = doc.selectSingleNode("HMC/Troubles/AlarmTime")
    If Not Node Is Nothing Then
      AlarmTime = Val(Node.text)
    End If


    Set NodeList = doc.selectNodes("HMC/Troubles/Alarm")
    If Not (NodeList Is Nothing) Then
      For Each childnode In NodeList
        Set alarm = New cAlarm

        For Each Node In childnode.childnodes
          Select Case LCase(Node.baseName)

            Case "resident"
              alarm.ResidentText = Node.text
            Case "roomtext"
              alarm.RoomText = Node.text
            Case "roomid"
              alarm.RoomID = Val(Node.text)
            Case "serial"      ' B20E087E
              alarm.Serial = Node.text
            Case "id"
              alarm.ID = Val(Node.text)
            Case "alarmid"
              alarm.AlarmID = Val(Node.text)
            Case "inputnum"    ' 1
              alarm.Inputnum = Val(Node.text)
            Case "locationtext"  ' Lobby
              alarm.locationtext = XMLDecode(Node.text)
            Case "datetime"    ' 02/05/07 2:50:55 PM
              alarm.DateTime = Node.text
            Case "announce"
              alarm.Announce = XMLDecode(Node.text)
            Case "acked"
              alarm.ACKed = Node.text
            Case "acktime"     ' 12:00:00 AM
              alarm.ACKTime = Node.text
            Case "responder"     ' 12:00:00 AM
              alarm.Responder = Node.text
            Case "priorid"
              alarm.PriorID = Val(Node.text)

            Case "alarmtype"   ' 1
              alarm.Alarmtype = Val(Node.text)
            Case "timestamp"   ' 1
              alarm.TimeStamp = Val(Node.text)
            Case "silenced"    ' 0?
              alarm.Silenced = Val(Node.text)
            Case "silencetime"  ' 02/05/07 2:52:02 PM
              alarm.SilenceTime = Node.text
            Case "description"
              alarm.Description = Node.text
            Case "custom"
              alarm.Custom = Node.text
          End Select
        Next
        NewAlarms.alarms.Add alarm
      Next
      NewAlarms.BeepTimer = BeepTimer
    End If

    Set Troubles = NewAlarms
    
    If Troubles.Count Then
      Troubles.LocalAlarmTime = AlarmTime
      Troubles.LocalSilenceTime = SilenceTime
    End If
    
    
    ProcessTroubles


    '****** LOW BATTS

    Set NewAlarms = New cAlarms
    Set Node = doc.selectSingleNode("HMC/LowBatts/Beep")
    BeepTimer = 0
    If Not Node Is Nothing Then
      BeepTimer = Val(Node.text)
    End If

    Set Node = doc.selectSingleNode("HMC/LowBatts/SilenceTime")
    SilenceTime = 0
    If Not Node Is Nothing Then
      SilenceTime = Val(Node.text)
    End If

    Set Node = doc.selectSingleNode("HMC/LowBatts/AlarmTime")
    If Not Node Is Nothing Then
      AlarmTime = Val(Node.text)
    End If


    Set NodeList = doc.selectNodes("HMC/LowBatts/Alarm")
    If Not (NodeList Is Nothing) Then
      For Each childnode In NodeList
        Set alarm = New cAlarm
        For Each Node In childnode.childnodes
          Select Case LCase(Node.baseName)
            Case "resident"
              alarm.ResidentText = Node.text
            Case "roomtext"
              alarm.RoomText = Node.text
            Case "roomid"
              alarm.RoomID = Val(Node.text)
            Case "serial"      ' B20E087E
              alarm.Serial = Node.text
            Case "id"
              alarm.ID = Val(Node.text)
            Case "alarmid"
              alarm.AlarmID = Val(Node.text)
            Case "inputnum"    ' 1
              alarm.Inputnum = Val(Node.text)
            Case "locationtext"  ' Lobby
              alarm.locationtext = XMLDecode(Node.text)
            Case "datetime"    ' 02/05/07 2:50:55 PM
              alarm.DateTime = Node.text
            Case "announce"
              alarm.Announce = XMLDecode(Node.text)
            Case "acked"
              alarm.ACKed = Node.text
            Case "acktime"     ' 12:00:00 AM
              alarm.ACKTime = Node.text
            Case "responder"     ' 12:00:00 AM
              alarm.Responder = Node.text
            Case "priorid"
              alarm.PriorID = Val(Node.text)
              
              
            Case "alarmtype"   ' 1
              alarm.Alarmtype = Val(Node.text)
            Case "timestamp"   ' 1
              alarm.TimeStamp = Val(Node.text)
            Case "silenced"    ' 0?
              alarm.Silenced = Val(Node.text)
            Case "silencetime"  ' 02/05/07 2:52:02 PM
              alarm.SilenceTime = Node.text
            Case "description"
              alarm.Description = Node.text
            Case "custom"
              alarm.Custom = Node.text
          End Select
        Next
        NewAlarms.alarms.Add alarm
      Next
      NewAlarms.BeepTimer = BeepTimer
    End If

    Set LowBatts = NewAlarms
    
    If LowBatts.Count Then
      LowBatts.LocalAlarmTime = AlarmTime
      LowBatts.LocalSilenceTime = SilenceTime
    End If
    
    ProcessBatts


    '****** EXTERNS

    Set NewAlarms = New cAlarms
    Set Node = doc.selectSingleNode("HMC/Externs/Beep")
    BeepTimer = 0
    If Not Node Is Nothing Then
      BeepTimer = Val(Node.text)
    End If


    Set Node = doc.selectSingleNode("HMC/Externs/SilenceTime")
    SilenceTime = 0
    If Not Node Is Nothing Then
      SilenceTime = Val(Node.text)
    End If

    Set Node = doc.selectSingleNode("HMC/Externs/AlarmTime")
    If Not Node Is Nothing Then
      AlarmTime = Val(Node.text)
    End If


    Set NodeList = doc.selectNodes("HMC/Externs/Alarm")
    If Not (NodeList Is Nothing) Then
      For Each childnode In NodeList
        Set alarm = New cAlarm
        For Each Node In childnode.childnodes
          Select Case LCase(Node.baseName)
            Case "resident"
              alarm.ResidentText = Node.text
            Case "roomtext"
              alarm.RoomText = Node.text
            Case "roomid"
              alarm.RoomID = Val(Node.text)
            Case "serial"      ' B20E087E
              alarm.Serial = Node.text
            Case "id"
              alarm.ID = Val(Node.text)
            Case "alarmid"
              alarm.AlarmID = Val(Node.text)
            Case "inputnum"    ' 1
              alarm.Inputnum = Val(Node.text)
            Case "locationtext"  ' Lobby
              alarm.locationtext = XMLDecode(Node.text)
            Case "datetime"    ' 02/05/07 2:50:55 PM
              alarm.DateTime = Node.text
            Case "announce"
              alarm.Announce = XMLDecode(Node.text)
            Case "acked"
              alarm.ACKed = Node.text
            Case "acktime"     ' 12:00:00 AM
              alarm.ACKTime = Node.text
            Case "responder"
              alarm.Responder = Node.text
            Case "priorid"
              alarm.PriorID = Val(Node.text)

            Case "alarmtype"   ' 1
              alarm.Alarmtype = Val(Node.text)
            Case "timestamp"   ' 1
              alarm.TimeStamp = Val(Node.text)
            Case "silenced"    ' 0?
              alarm.Silenced = Val(Node.text)
            Case "silencetime"  ' 02/05/07 2:52:02 PM
              alarm.SilenceTime = Node.text
            Case "description"
              alarm.Description = Node.text
            Case "custom"
              alarm.Custom = Node.text
          End Select

        Next
        NewAlarms.alarms.Add alarm

      Next
      NewAlarms.BeepTimer = BeepTimer
    End If

    Set Externs = NewAlarms
    
    If Externs.Count Then
      Externs.LocalAlarmTime = AlarmTime
      Externs.LocalSilenceTime = SilenceTime
    End If
    
    
    ProcessExterns

    '****** ASSURS
    BeepTimer = 0
    Set NewAlarms = New cAlarms
    Set Node = doc.selectSingleNode("HMC/Assurs/Beep")
    If Not Node Is Nothing Then
      BeepTimer = Val(Node.text)
    End If

    Set Node = doc.selectSingleNode("HMC/Assurs/SilenceTime")
    SilenceTime = 0
    If Not Node Is Nothing Then
      SilenceTime = Val(Node.text)
    End If



    If gAssurDisableScreenOutput = 0 Then

      Set NodeList = doc.selectNodes("HMC/Assurs/Alarm")
      If Not (NodeList Is Nothing) Then
        For Each childnode In NodeList
          Set alarm = New cAlarm
          For Each Node In childnode.childnodes
            Select Case LCase(Node.baseName)
              Case "serial"    ' B20E087E
                alarm.Serial = Node.text
              Case "id"
                alarm.ID = Val(Node.text)
              Case "alarmid"
                alarm.AlarmID = Val(Node.text)
            End Select
          Next
          NewAlarms.alarms.Add alarm
        Next
      End If
      NewAlarms.BeepTimer = BeepTimer
      Set Assurs = NewAlarms
      ProcessAssurs False
    End If
  End If


RefreshAlarms_Resume:
  On Error GoTo 0
  Exit Sub

RefreshAlarms_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.RefreshAlarms." & Erl
  Resume RefreshAlarms_Resume


End Sub



Public Property Let CurrentGrid(ByVal Value As Integer)
  On Error Resume Next
  mCurrentGrid = Value
  If Value <> InfoPane.emergency Then
    mCurrentSubGrid = Value
  End If
  UpdateLayout
End Property




Public Sub PacketToggle()
  Static i As Integer
  ' cycle thru "data" images
  Select Case i
    Case 0
      imgPacket.Picture = LoadResPicture(1000, vbResBitmap)
      i = i + 1
    Case 1
      imgPacket.Picture = LoadResPicture(1001, vbResBitmap)
      i = i + 1
    Case 2
      imgPacket.Picture = LoadResPicture(1002, vbResBitmap)
      i = i + 1
    Case 3
      imgPacket.Picture = LoadResPicture(1003, vbResBitmap)
      i = i + 1
    Case 4
      imgPacket.Picture = LoadResPicture(1004, vbResBitmap)
      i = 1
  End Select

End Sub


Sub EmergencyUp()
  Dim i As Long
  SetFocusTo lvEmergency
  If Not lvEmergency.SelectedItem Is Nothing Then
    i = lvEmergency.SelectedItem.index - 1
    If i > 0 Then
      lvEmergency.SelectedItem = lvEmergency.ListItems(i)
      lvEmergency.SelectedItem.EnsureVisible
    End If
  End If
  ShowEmergencyInfo

End Sub
Sub EmergencyDown()
  Dim i As Long
  SetFocusTo lvEmergency
  If Not lvEmergency.SelectedItem Is Nothing Then
    i = lvEmergency.SelectedItem.index + 1
    If i <= lvEmergency.ListItems.Count Then
      lvEmergency.SelectedItem = lvEmergency.ListItems(i)
      lvEmergency.SelectedItem.EnsureVisible
    End If
  End If
  ShowEmergencyInfo

End Sub
Sub AssuranceUp()
  Dim i As Long
  If lvAssur.SelectedItem Is Nothing Then Exit Sub
  i = lvAssur.SelectedItem.index - 1
  If i > 0 Then
    lvAssur.SelectedItem = lvAssur.ListItems(i)
    lvAssur.SelectedItem.EnsureVisible
  End If
  ShowAssurInfo

End Sub
Sub AssuranceDown()
  Dim i As Long
  If lvAssur.SelectedItem Is Nothing Then Exit Sub
  i = lvAssur.SelectedItem.index + 1
  If i <= lvAssur.ListItems.Count Then
    lvAssur.SelectedItem = lvAssur.ListItems(i)
    lvAssur.SelectedItem.EnsureVisible
  End If
  ShowAssurInfo

End Sub

Function GetKey(li As ListItem) As Long
  GetKey = Val(li.Key)
End Function

Sub CenterPanelUp()
  Dim lv As ListView
  Dim i As Long
  Select Case mCurrentSubGrid
    Case InfoPane.Alert
      Set lv = lvAlerts
    Case InfoPane.Checkin
      Set lv = lvCheckIn
    Case InfoPane.LowBatt
      Set lv = lvLoBatt
    Case InfoPane.extern
      Set lv = lvExternal
  End Select

  If lv Is Nothing Then Exit Sub
  SetFocusTo lv
  If lv.SelectedItem Is Nothing Then Exit Sub
  i = lv.SelectedItem.index - 1
  If i > 0 Then
    lv.SelectedItem = lv.ListItems(i)
    lv.SelectedItem.EnsureVisible
  End If
  ShowResidentInfo 0


End Sub

Sub CenterPanelDown()
  Dim lv As ListView
  Dim i As Long
  Select Case mCurrentSubGrid
    Case InfoPane.Alert
      Set lv = lvAlerts
    Case InfoPane.Checkin
      Set lv = lvCheckIn
    Case InfoPane.LowBatt
      Set lv = lvLoBatt
    Case InfoPane.extern
      Set lv = lvExternal
  End Select

  If lv Is Nothing Then Exit Sub
  SetFocusTo lv
  If Not lv.SelectedItem Is Nothing Then
    i = lv.SelectedItem.index + 1
    If i <= lv.ListItems.Count Then
      lv.SelectedItem = lv.ListItems(i)
      lv.SelectedItem.EnsureVisible
    End If
  End If
  ShowResidentInfo 0


End Sub

Sub CenterPanelPrintList()


  Select Case mCurrentSubGrid
    Case 0, InfoPane.Alert
      Dim Report As cAlarmReport
      Set Report = New cAlarmReport
      Report.PrintList lvAlerts, Alerts, "Alerts"
      Set Report = Nothing

    Case InfoPane.Checkin
      Dim Report2 As cTroubleReport
      Set Report2 = New cTroubleReport
      Report2.PrintList lvCheckIn, Troubles, "Device Troubles"
      Set Report2 = Nothing
    Case InfoPane.LowBatt
      Dim Report3 As cTroubleReport
      Set Report3 = New cTroubleReport
      Report3.PrintList lvLoBatt, LowBatts, "Low Battery Report"
      Set Report3 = Nothing
    Case InfoPane.extern
      Dim Report4 As cExternReport
      Set Report4 = New cExternReport
      Report4.PrintList lvExternal, Externs, "External Alarm Report"
      Set Report4 = Nothing
    
      
      
  End Select



End Sub



Sub PrintInfoWindow(ByVal ID As Long)
  If ID <> 0 Then

    Dim InfoPrinter As cInfoPrinter
    Set InfoPrinter = New cInfoPrinter
    InfoPrinter.PrintInfo ID
    Set InfoPrinter = Nothing
  End If
End Sub

Public Property Let Assur(ByVal Value As Boolean)
  mAssur = Assur
  fraassur.Visible = Value
  fraHost.Visible = Not Value
  Assurs.BeepTimer = 0

End Property
Public Property Get Assur() As Boolean
  Assur = mAssur
End Property

Private Sub cmdAck_Click()
'  ResetActivityTime
  If MASTER Then
    AckSelected
  Else
    ResetRemoteRefreshCounter
    DoClientACKAlarm
    ResetRemoteRefreshCounter
  End If
End Sub

Sub DoClientRequestAssist()

  Dim panel              As String
  Dim Key                As String
  Dim Inputnum           As Long
  Dim Serial             As String
  Dim alarm              As cAlarm
  Dim j                  As Long
  Dim AlarmID            As Long

  ResetRemoteRefreshCounter

  If mCurrentGrid = InfoPane.emergency Then
    panel = "alarms"
    If Not lvEmergency.SelectedItem Is Nothing Then
      
      'alarmid = GetlvEmergencySelectedID
      AlarmID = lvKey2ID(lvEmergency.SelectedItem.Key)
      Key = lvEmergency.SelectedItem.Key
      
      
      Serial = left(Key, 8)
      If Len(Key) > 9 Then
        Inputnum = Val(MID(Key, 10))
      End If
      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Serial = alarm.Serial
          Inputnum = alarm.Inputnum
          If alarm.Alarmtype = EVT_ASSISTANCE Then
            ' can't do assistance on assistance
'            If vbYes = modLib.messagebox(Me, "Acknowledging an Assistance Call Will Terminate the Call" & vbCrLf & "Proceed?", App.Title, vbYesNo Or vbExclamation) Then
'              ClientACKAlarm panel, Serial, inputnum, AlarmID
'            End If
            Exit For
          Else
            'ClientACKAlarm panel, Serial, inputnum
            
            Call ClientRequestAssist(panel, Serial, Inputnum, AlarmID)
            Exit For
          End If
        End If
      Next
    End If

  ElseIf mCurrentGrid = InfoPane.Alert Then
    panel = "alerts"
    If Not lvAlerts.SelectedItem Is Nothing Then
      'alarmid = alarmid = GetlvEmergencySelectedID
      
      AlarmID = lvKey2ID(lvAlerts.SelectedItem.Key)
      Key = lvAlerts.SelectedItem.Key
'      Serial = left(key, 8)
'      If Len(key) > 9 Then
'        inputnum = Val(MID(key, 10))
'      End If
      Serial = alarm.Serial
      Inputnum = alarm.Inputnum
      'ClientACKAlarm panel, Serial, inputnum
      ClientRequestAssist panel, Serial, Inputnum, AlarmID
    End If

  ElseIf mCurrentGrid = InfoPane.extern Then
    panel = "externs"
    If Not lvExternal.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvExternal.SelectedItem.Key)
'      key = lvExternal.SelectedItem.key
'      Serial = left(key, 8)
'      Serial = alarm.Serial
'      If Len(key) > 9 Then
'        inputnum = Val(MID(key, 10))
'      End If
      Inputnum = alarm.Inputnum
      ClientRequestAssist panel, Serial, Inputnum, AlarmID
    End If
  End If



End Sub

Sub DoClientACKAlarm()
  Dim panel              As String
  Dim Key                As String
  Dim Inputnum           As Long
  Dim Serial             As String
  Dim alarm              As cAlarm
  Dim j                  As Long
  Dim AlarmID            As Long
  Dim f                  As frmAssistCancel
  Dim Disposition        As String
  
  
  If mCurrentGrid = InfoPane.emergency Then
    panel = "alarms"
    If Not lvEmergency.SelectedItem Is Nothing Then

      'alarmid = GetlvEmergencySelectedID
      AlarmID = lvKey2ID(lvEmergency.SelectedItem.Key)
      Key = lvEmergency.SelectedItem.Key


      '      Serial = left(key, 8)
      '      If Len(key) > 9 Then
      '        inputnum = Val(MID(key, 10))
      '      End If
      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Serial = alarm.Serial
          Inputnum = alarm.Inputnum
          If alarm.Alarmtype = EVT_ASSISTANCE Then
            If vbYes = modLib.messagebox(Me, "Acknowledging an Assistance Call Will Terminate the Call" & vbCrLf & "Proceed?", App.Title, vbYesNo Or vbExclamation) Then

              Set f = New frmAssistCancel
              Load f
              f.Show vbModal, Me
              Disposition = Trim$(f.Disposition)
              Unload f
              Set f = Nothing
              If Len(Disposition) Then
                alarm.Disposition = Disposition
                alarm.Username = gUser.Username
                
                modRemote.ClientFinalizeAlarm "alarms", alarm.Serial, Inputnum, AlarmID, Disposition
                ' sends event over the wire
                
                UpdateDispositions Disposition
                
              End If
            End If
            Exit For
          Else
            'ClientACKAlarm panel, Serial, inputnum
            ClientACKAlarm panel, Serial, Inputnum, AlarmID
            Exit For
          End If
        End If
      Next
    End If

  ElseIf mCurrentGrid = InfoPane.Alert Then
    panel = "alerts"
    If Not lvAlerts.SelectedItem Is Nothing Then
      'alarmid = alarmid = GetlvEmergencySelectedID

      AlarmID = lvKey2ID(lvAlerts.SelectedItem.Key)
      '      key = lvAlerts.SelectedItem.key
      '      Serial = left(key, 8)
      '      If Len(key) > 9 Then
      '        inputnum = Val(MID(key, 10))
      '      End If
      Serial = alarm.Serial
      Inputnum = alarm.Inputnum
      'ClientACKAlarm panel, Serial, inputnum
      ClientACKAlarm panel, Serial, Inputnum, AlarmID
    End If

  ElseIf mCurrentGrid = InfoPane.extern Then
    panel = "externs"
    If Not lvExternal.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvExternal.SelectedItem.Key)
      '      key = lvExternal.SelectedItem.key
      '      Serial = left(key, 8)
      '      Serial = alarm.Serial
      '      If Len(key) > 9 Then
      '        inputnum = Val(MID(key, 10))
      '      End If
      Inputnum = alarm.Inputnum
      ClientACKAlarm panel, Serial, Inputnum, AlarmID
    End If
  End If


End Sub

Sub RemoteClientRequestAssist(ByVal panel As String, ByVal Serial As String, ByVal Inputnum As Long, ByVal User As String, ByVal AlarmID As Long)

  Dim j                  As Integer
  Dim alarm              As cAlarm
  Dim d                  As cESDevice
  Debug.Print
  Debug.Print "ClientRequestAssist"
  Debug.Print

  

  Select Case LCase(panel)
    Case "alarms"
      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then

          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = User
            Call PostEvent(d, Nothing, alarm, EVT_ASSISTANCE, alarm.Inputnum, gUser.Username)
          End If
          Exit For
        End If
      Next

    Case "alerts"
      For j = 1 To Alerts.alarms.Count
        Set alarm = Alerts.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = User
            Call PostEvent(d, Nothing, alarm, EVT_ASSISTANCE, alarm.Inputnum, gUser.Username)
          End If
          Exit For
        End If
      Next
    Case "externs"
      For j = 1 To Externs.alarms.Count
        Set alarm = Externs.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = User
            Call PostEvent(d, Nothing, alarm, EVT_ASSISTANCE, alarm.Inputnum, gUser.Username)
          End If
          Exit For
        End If
      Next
  End Select


End Sub

Sub ClientACKSelected(ByVal panel As String, ByVal Serial As String, ByVal Inputnum As Long, ByVal User As String, ByVal AlarmID As Long)
  Dim j                  As Integer
  Dim alarm              As cAlarm
  Dim d                  As cESDevice
  Debug.Print
  Debug.Print "ClientAckSelected"
  Debug.Print
  DeletingAlarm = True
  Select Case LCase(panel)
    Case "alarms"
      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then

          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            If alarm.EventType = EVT_ASSISTANCE Then
              alarm.Username = User
              PostEvent d, Nothing, alarm, EVT_ASSISTANCE_ACK, alarm.Inputnum
              ShowACK alarm

            Else
              alarm.Username = User
              PostEvent d, Nothing, alarm, EVT_EMERGENCY_ACK, alarm.Inputnum
              ShowACK alarm
            End If
          End If
          Exit For
        End If
      Next
    Case "alerts"
      For j = 1 To Alerts.alarms.Count
        Set alarm = Alerts.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = User
            PostEvent d, Nothing, alarm, EVT_ALERT_ACK, alarm.Inputnum
            ShowACK alarm
          End If
          Exit For
        End If
      Next
    Case "externs"
      For j = 1 To Externs.alarms.Count
        Set alarm = Externs.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = User
            PostEvent d, Nothing, alarm, EVT_EXTERN_ACK, alarm.Inputnum
            ShowACK alarm
          End If
          Exit For
        End If
      Next

  End Select
  DeletingAlarm = False

End Sub

Sub AckSelected()
  Dim Key
  Dim j                  As Integer
  Dim alarm              As cAlarm
  Dim d                  As cESDevice
  Dim Inputnum           As Long
  Dim Serial             As String
  Dim AlarmID            As Long
  Dim f                  As frmAssistCancel
  Dim Disposition        As String


  DeletingAlarm = True


  If mCurrentGrid = InfoPane.emergency Then

    If Not lvEmergency.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvEmergency.SelectedItem.Key)  ' GetlvEmergencySelectedID()
'      key = lvEmergency.SelectedItem.key
'      Serial = left(key, 8)
'      If Len(key) > 9 Then
'       inputnum = Val(MID(lvEmergency.SelectedItem.key, 10))
'      End If

      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Inputnum = alarm.Inputnum
          If alarm.Alarmtype = EVT_ASSISTANCE Then  ' special assistance alarm

            Set f = New frmAssistCancel
            Load f
            f.Show vbModal, Me
            Disposition = Trim$(f.Disposition)
            Unload f
            Set f = Nothing

            If Len(Disposition) Then
              alarm.Disposition = Disposition

              'If vbYes = modLib.messagebox(Me, "Acknowledging an Assistance Call Will Terminate the Call" & vbCrLf & "Proceed?", App.Title, vbYesNo Or vbExclamation) Then
              Set d = Devices.Device(alarm.Serial)
              If Not d Is Nothing Then
                alarm.Username = gUser.Username
                alarm.AlarmID = AlarmID
                'ShowACK alarm
                'If MASTER Then
                  PostEvent d, Nothing, alarm, EVT_ASSISTANCE_ACK, alarm.Inputnum
                  UpdateDispositions Disposition
                'Else
                '  ClientFinalizeAlarm "alarms", alarm.Serial, inputnum, AlarmID, Disposition
                'End If
                
                
              End If
              ' update disposition table
            Else
              DeletingAlarm = False
            End If

            Exit For

          Else                 ' regular alarm
            Set d = Devices.Device(alarm.Serial)
            If Not d Is Nothing Then
              alarm.Username = gUser.Username
              ShowACK alarm
              PostEvent d, Nothing, alarm, EVT_EMERGENCY_ACK, alarm.Inputnum
            End If
            Exit For
          End If
        End If
      Next
    End If
  ElseIf mCurrentGrid = InfoPane.Alert Then


    If Not lvAlerts.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvAlerts.SelectedItem.Key)
'      key = lvAlerts.SelectedItem.key
'      Serial = left(key, 8)
'      If Len(key) > 9 Then
'        inputnum = Val(MID(lvAlerts.SelectedItem.key, 10))
'      End If


      For j = 1 To Alerts.alarms.Count
        Set alarm = Alerts.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = gUser.Username
            ShowACK alarm
            PostEvent d, Nothing, alarm, EVT_ALERT_ACK, alarm.Inputnum

          End If
          Exit For
        End If
      Next
    End If
  ElseIf mCurrentGrid = InfoPane.extern Then

    If Not lvExternal.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvExternal.SelectedItem.Key)
    
'      key = lvExternal.SelectedItem.key
'      Serial = left(key, 8)
'      If Len(key) > 9 Then
'        inputnum = Val(MID(lvExternal.SelectedItem.key, 10))
'      End If


      For j = 1 To Externs.alarms.Count
        Set alarm = Externs.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          Set d = Devices.Device(alarm.Serial)
          If Not d Is Nothing Then
            alarm.Username = gUser.Username
            ShowACK alarm
            PostEvent d, Nothing, alarm, EVT_EXTERN_ACK, alarm.Inputnum

          End If
          Exit For
        End If
      Next
    End If


  End If

  DeletingAlarm = False

End Sub

Public Sub ShowResponder(alarm As cAlarm)
        Dim li                 As ListItem
        Dim Inputnum           As Long
        Dim Key                As String
        Dim Serial             As String
        Dim KeyInput()         As String
        Dim AlarmID            As Long
        'Debug.Assert 0

10      On Error GoTo ShowResponder_Error

20      If alarm Is Nothing Then
          'Debug.Assert 0
30        Exit Sub

40      End If



50      For Each li In lvEmergency.ListItems
60        Key = li.Key

70        KeyInput = Split(Key, "s", , vbTextCompare)
80        If UBound(KeyInput) = 1 Then
90          AlarmID = Val(KeyInput(0))
100         Inputnum = Val(KeyInput(1))
110       Else
120         AlarmID = left(li.Key, 8)
130         If Len(Key) > 9 Then
140           Inputnum = Val(MID$(Key, 10))
150         End If
160       End If
170       If alarm.ID = AlarmID Then

            'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
180         'If alarm.EventType = EVT_ASSISTANCE Or alarm.EventType = EVT_ASSISTANCE_ACK Or alarm.EventType = EVT_ASSISTANCE_RESTORE Then
190           If Trim$(li.ListSubItems(LVEMERGENCY_RESP).text) = "" Then
200             li.ListSubItems(LVEMERGENCY_RESP).text = alarm.Responder
210           End If
220         'ElseIf alarm.EventType = EVT_EMERGENCY Or alarm.EventType = EVT_EMERGENCY_ACK Or alarm.EventType = EVT_BATTERY_RESTORE Then
230          ' If Trim$(li.ListSubItems(LVEMERGENCY_RESP).text) = "" Then
240           '  li.ListSubItems(LVEMERGENCY_RESP).text = alarm.Responder
250           'End If
260         'End If
            lvEmergency.Refresh
            
270         Exit For
280       End If
290     Next

300     For Each li In lvAlerts.ListItems
310       Key = li.Key
320       KeyInput = Split(Key, "s", , vbTextCompare)
330       If UBound(KeyInput) = 1 Then
340         AlarmID = Val(KeyInput(0))
350         Inputnum = Val(KeyInput(1))
360       Else
370         AlarmID = left(li.Key, 8)
380         If Len(Key) > 9 Then
390           Inputnum = Val(MID$(Key, 10))
400         End If
410       End If
420       If alarm.ID = AlarmID Then

      '    Serial = left(li.Key, 8)
      '    If Len(Key) > 9 Then
      '      inputnum = Val(MID(Key, 10))
      '    End If
      '
      '    If alarm.ID = Val(Key) Then
430         If Trim$(li.ListSubItems(LVALERT_RESP).text) = "" Then
440           li.ListSubItems(LVALERT_RESP).text = alarm.Responder
450         End If
            lvAlerts.Refresh
460         Exit For
470       End If
480     Next


ShowResponder_Resume:

490     On Error GoTo 0
500     Exit Sub

ShowResponder_Error:

510     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowResponder." & Erl
520     Resume ShowResponder_Resume

End Sub


Public Sub ShowACK(alarm As cAlarm)
        Dim li                 As ListItem
        Dim Inputnum           As Long
        Dim Key                As String
        Dim Serial             As String
        Dim KeyInput()         As String
        Dim AlarmID            As Long

        
        
10      On Error GoTo ShowACK_Error

20      If alarm Is Nothing Then
          'Debug.Assert 0
30        Exit Sub

40      End If

50      For Each li In lvEmergency.ListItems
60        Key = li.Key

70        KeyInput = Split(Key, "s", , vbTextCompare)
80        If UBound(KeyInput) = 1 Then
90          AlarmID = Val(KeyInput(0))
100         Inputnum = Val(KeyInput(1))
110       Else
120         AlarmID = left(li.Key, 8)
130         If Len(Key) > 9 Then
140           Inputnum = Val(MID$(Key, 10))
150         End If
160       End If
170       If alarm.ID = AlarmID Then


            '70        Serial = left(li.Key, 8)
            '80        If Len(Key) > 9 Then
            '90          inputnum = Val(MID$(Key, 10))
            '100       End If
            '
            '
            '110       If alarm.ID = Val(Key) Then

            'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
180         If alarm.EventType = EVT_ASSISTANCE Then
190           If CDbl(alarm.ACKTime) = 0 Then
200             li.ListSubItems(LVEMERGENCY_ACK).text = ""
210           Else
220             li.ListSubItems(LVEMERGENCY_ACK).text = Format(alarm.ACKTime, gTimeFormatString)
230           End If


240         Else
250           If CDbl(alarm.ACKTime) = 0 Then
260             li.ListSubItems(LVEMERGENCY_ACK).text = ""
270           Else
280             li.ListSubItems(LVEMERGENCY_ACK).text = Format(alarm.ACKTime, gTimeFormatString)
290           End If
300         End If
310         Exit For
320       End If
330     Next

340     For Each li In lvAlerts.ListItems
350       Key = li.Key

360       KeyInput = Split(Key, "s", , vbTextCompare)
370       If UBound(KeyInput) = 1 Then
380         AlarmID = Val(KeyInput(0))
390         Inputnum = Val(KeyInput(1))
400       Else
410         AlarmID = left(li.Key, 8)
420         If Len(Key) > 9 Then
430           Inputnum = Val(MID$(Key, 10))
440         End If
450       End If
460       If alarm.ID = AlarmID Then



            '360       Serial = left(li.Key, 8)
            '370       If Len(Key) > 9 Then
            '380         inputnum = Val(MID(Key, 10))
            '390       End If
            '
            '
            '400       If alarm.ID = Val(Key) Then

470         If CDbl(alarm.ACKTime) = 0 Then
480           li.ListSubItems(LVALERT_ACK).text = ""
490         Else
500           li.ListSubItems(LVALERT_ACK).text = Format(alarm.ACKTime, gTimeFormatString)
510         End If

520         Exit For
530       End If
540     Next


ShowACK_Resume:
550     On Error GoTo 0
560     Exit Sub

ShowACK_Error:

570     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowACK." & Erl
580     Resume ShowACK_Resume


End Sub



Private Sub cmdAlarmdown_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.emergency
  EmergencyDown
End Sub
Private Sub cmdAlarmPrintList_Click()
'  ResetActivityTime
  If Printer Is Nothing Then Exit Sub
  Dim Report As cAlarmReport
  ResetRemoteRefreshCounter
  
  Static Busy As Boolean
  If Busy Then Exit Sub
  Busy = True
  Set Report = New cAlarmReport
  Report.PrintList lvEmergency, alarms, "Emergency"
  Set Report = Nothing
  SetFocusTo lvEmergency
  Busy = False
End Sub
Private Sub cmdAlarmUp_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.emergency
  EmergencyUp
End Sub

Private Sub cmdAlert_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  ShowAlert
End Sub
Sub ShowAlert()
  CurrentGrid = InfoPane.Alert
  ShowAlertInfo
End Sub

Private Sub cmdAnnouncements_Click()
  ResetActivityTime
  ClearHostedForms
  ShowAnnouncementForm

End Sub


Private Sub cmdAssur_Click()
  'ResetActivityTime
  lvAssur.left = 0
  lblAssur.Width = lvAssur.Width
  Assur = True
End Sub

Private Sub cmdAssurExit_Click()
  fraHost.Visible = True
  fraassur.Visible = False
End Sub

Private Sub cmdAssurPrintList_Click()
If Printer Is Nothing Then Exit Sub
  
  Static Busy As Boolean
  Dim Report As cAssuranceReport
  If Busy Then Exit Sub
  Busy = True
  Set Report = New cAssuranceReport
  If Report.AssurPrintList(lvAssur, "Check-ins") Then
    If InAssurPeriod = False Then
      If MASTER Then
        Assurs.Clear
        Set AssureVacationDevices = New Collection
        ProcessAssurs False
        Assur = False
        lvAssur.ListItems.Clear
      Else
        RemoteClearAssurs
      End If
  
    Else
      lvAssur.ListItems.Clear
    End If
    
  End If
  Busy = False
  ResetActivityTime
End Sub

Private Sub cmdAssurUp_Click()
  AssuranceUp
End Sub

Private Sub cmdAsurrDown_Click()
  AssuranceDown
End Sub

Private Sub cmdBattery_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  AckLowBatt
  ShowLowBattInfo
  
End Sub






Private Sub cmdChangeVacation_Click()
'  ResetActivityTime
  cmdChangeVacation.Enabled = False
  ResetRemoteRefreshCounter
  ToggleVacation IIf(cmdChangeVacation.Caption = "Return from Vacation", 0, 1)
  If Val(txtHiddenResID.text) <> 0 Then
    'RefreshJet
    DisplayResidentInfo Val(txtHiddenResID.text), Val(txtHiddenAlarmID.text)
  ElseIf Val(txtHiddenRoomID.text) <> 0 Then
    DisplayResOrRoomInfo 0, Val(txtHiddenRoomID.text), Val(txtHiddenAlarmID.text)
  End If
  
  cmdChangeVacation.Enabled = True
  On Error Resume Next
  SetFocusTo cmdChangeVacation
End Sub
Sub ToggleVacation(ByVal Away As Integer)
        Dim ID As Long

        Dim rs As Recordset
        'Dim away As Integer
        Dim SQL  As String

10      On Error GoTo ToggleVacation_Error

20      If Val(txtHiddenResID.text) <> 0 Then
30        ID = Val(txtHiddenResID.text)
40        'Set rs = ConnExecute("SELECT Away FROM Residents WHERE ResidentID = " & ID)
50        'If Not rs.EOF Then
60          'away = IIf(rs("Away") = 1, 1, 0)
70          'rs.Close

80          'If away = 1 Then
90          '  away = 0
100         'Else
110         '  away = 1
120         'End If
130         ID = Val(txtHiddenResID.text)
140         If MASTER Then
150           SetResidentAwayStatus Away, ID, gUser.Username
160         Else
170           If 0 = RemoteSetResidentAwayStatus(Away, ID) Then  ' no errors
                ' get away status
                RefreshJet
180           End If
190         End If
200       'Else
210         'rs.Close
220       'End If


230     ElseIf Val(txtHiddenRoomID.text) <> 0 Then
240       ID = Val(txtHiddenRoomID.text)
250       'Set rs = ConnExecute("SELECT Away FROM Rooms WHERE RoomID = " & ID)
260       'If Not rs.EOF Then
270         'away = IIf(rs("Away") = 1, 1, 0)
280         'rs.Close
290         'If away = 1 Then
300         '  away = 0
310         'Else
320         '  away = 1
330         'End If

340         ID = Val(txtHiddenRoomID.text)
350         If MASTER Then
360           SetRoomAwayStatus Away, ID, gUser.Username
370         Else

380           If 0 = RemoteSetRoomAwayStatus(Away, ID) Then  ' no errors
                ' get away status
                RefreshJet
390           End If
              
400         End If
410       'Else
420       '  rs.Close
430       'End If
440     End If

ToggleVacation_Resume:
450     On Error GoTo 0
460     Exit Sub

ToggleVacation_Error:

470     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ToggleVacation." & Erl
480     Resume ToggleVacation_Resume


End Sub

Public Sub ClearAwayAlarmsByRoomID(ByVal RoomID As Long)
  Dim d         As cESDevice
  Dim Serial    As String
  Dim alarm     As cAlarm
  Dim Alarmtype As Long
  Dim Inputnum  As Integer
  Dim PageRequest As cPageItem
  
  For Each d In Devices.Devices
    If d.RoomID = RoomID Then
      Serial = d.Serial
      Troubles.ClearAllAlarmsBySerial Serial
      LowBatts.ClearAllAlarmsBySerial Serial
      If d.AssurSecure = 0 Then
        Do While 1
        Set alarm = alarms.BySerial(Serial)
          If alarm Is Nothing Then Exit Do
          Set PageRequest = RemovePageRequest(alarm.ID) ' Serial, Alarm.AlarmType, Alarm.InputNum)
          PostEvent d, Nothing, alarm, alarm.Alarmtype, alarm.Inputnum
          alarms.RemoveAlarm alarm, alarm.Alarmtype
          Set alarm = Nothing ' to eliminate memory leak
        Loop
        Do While 1
        Set alarm = Alerts.BySerial(Serial)
          If alarm Is Nothing Then Exit Do
          Set PageRequest = RemovePageRequest(alarm.ID) ' Serial, Alarm.AlarmType, Alarm.InputNum)
          PostEvent d, Nothing, alarm, alarm.Alarmtype, alarm.Inputnum
          Alerts.RemoveAlarm alarm, alarm.Alarmtype
          Set alarm = Nothing ' to eliminate memory leak
        Loop
      End If
    End If
  Next

End Sub
Public Sub ClearAwayAlarms(ByVal ResidentID As Long)
  Dim d         As cESDevice
  Dim Serial    As String
  Dim alarm     As cAlarm
  Dim Alarmtype As Long
  Dim Inputnum  As Integer
  Dim PageRequest As cPageItem
  
  For Each d In Devices.Devices
    If d.ResidentID = ResidentID Then
      Serial = d.Serial
      Troubles.ClearAllAlarmsBySerial Serial
      LowBatts.ClearAllAlarmsBySerial Serial
      If d.AssurSecure = 0 Then
        Do While 1
        Set alarm = alarms.BySerial(Serial)
          If alarm Is Nothing Then Exit Do
          Set PageRequest = RemovePageRequest(alarm.ID) ' Serial, Alarm.AlarmType, Alarm.InputNum)
          'Set PageRequest = RemovePageRequest(Serial, Alarm.AlarmType, Alarm.InputNum)
          'If Not PageRequest Is Nothing Then
          '  SendEndofEventPage PageRequest, Alarm.AlarmType
          'End If
          PostEvent d, Nothing, alarm, alarm.Alarmtype, alarm.Inputnum
          alarms.RemoveAlarm alarm, alarm.Alarmtype
          Set alarm = Nothing ' to eliminate memory leak
        Loop
        Do While 1
        Set alarm = Alerts.BySerial(Serial)
          If alarm Is Nothing Then Exit Do
          Set PageRequest = RemovePageRequest(alarm.ID) ' Serial, Alarm.AlarmType, Alarm.InputNum)
          'Set PageRequest = RemovePageRequest(Serial, Alarm.AlarmType, Alarm.InputNum)
          'If Not PageRequest Is Nothing Then
          '  SendEndofEventPage PageRequest, Alarm.AlarmType
          'End If
          PostEvent d, Nothing, alarm, alarm.Alarmtype, alarm.Inputnum
          Alerts.RemoveAlarm alarm, alarm.Alarmtype
          Set alarm = Nothing ' to eliminate memory leak
        Loop
      End If
    End If
  Next
  
End Sub
Private Sub cmdClearInfo_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  DisplayResidentInfo 0, 0
End Sub

Private Sub CmdEditInfo_Click()
'  ResetActivityTime
  Dim ID As Long
  ResetRemoteRefreshCounter
  ID = Val(txtHiddenResID.text)
  If ID <> 0 Then
    EditResident ID
  End If
End Sub

Private Sub cmdExternal_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  ShowExternal
End Sub
Private Sub ShowExternal()
  CurrentGrid = InfoPane.extern
  ShowResidentInfo 0

End Sub
Private Sub cmdLogin_Click()
    
  RemoveHostedForms
  PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_LOGOUT, 0
  DoLogin
  
  Form_Paint
  
End Sub

Public Function LocalLogout(User As cUser)
  Dim j                  As Integer
  Dim Session            As cUser

  ' this should not log out all users! 12/4/14

  For j = HostSessions.Count To 1 Step -1
    Set Session = HostSessions(j)
    If Session.Session = User.Session Then
      If Session.LEvel >= LEVEL_SUPERVISOR Then
        dbg "Bumping Admin"
        HostSessions.Remove j
        LogRemoteSession Session.Session, 0, "LocalLogout"
      Else
        dbg "Bumping User " & Session.Username
        HostSessions.Remove j
        LogRemoteSession Session.Session, 0, "LocalLogout2"
      End If
    End If
  Next


End Function
Public Function DoLogin(Optional bypass As String) As Boolean
  Dim login As String
  Dim CurrentUserlevel As Long
  
  login = Trim(txtLogin.text)
  CurrentUserlevel = gUser.LEvel
  
  If MASTER Then
    LocalLogout gUser
    If CurrentUserlevel >= LEVEL_SUPERVISOR Then
      If Len(login) = 0 Then
        login = "0000"
      End If
    End If
        
    
  Else
    dbg "Logging out (Main DoLogin)"
    RemoteLogout gUser
    
    If CurrentUserlevel >= LEVEL_SUPERVISOR Then
      If Len(login) = 0 Then
        login = "0000"
      End If
    End If
    
    
  End If
  
  Set gUser = New cUser
  LoggedIn = False
  
  UpdateScreenElements
  
  
  If Len(bypass) > 0 Then
     login = bypass
  End If
  
  If (Len(login) > 0) Then
    
    If ProcessLogin(login) Then
      UpdateScreenElements
      txtLogin.text = ""
      DoLogin = True
      PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_LOGIN, 0
      Me.Enabled = True
      
      frmMain.ShowCountdown = ExistsUser0000()
      Exit Function
      
    End If
  End If
    
  Me.Enabled = False
  
  
  frmLogin.Show vbModeless, Me
  frmLogin.ShowCountdown = ExistsUser0000()
  frmMain.ShowCountdown = frmLogin.ShowCountdown
  ResetLockTime
  


End Function

Private Sub cmdMultiDown_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  CenterPanelDown
End Sub

Private Sub cmdMultiPrint_Click()
'  ResetActivityTime
If Printer Is Nothing Then Exit Sub
  
  Static Busy As Boolean
  ResetRemoteRefreshCounter
  If Busy Then Exit Sub
  Busy = True

  CenterPanelPrintList

  Busy = False
  Select Case mCurrentSubGrid
    Case InfoPane.Alert
      SetFocusTo lvAlerts
    Case InfoPane.Checkin
      SetFocusTo lvCheckIn
    Case InfoPane.LowBatt
      SetFocusTo lvLoBatt
    Case InfoPane.extern
      SetFocusTo lvExternal
  End Select
End Sub

Private Sub cmdMultiUp_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  CenterPanelUp

End Sub

Private Sub cmdOutputDevices_Click()
  ResetActivityTime
  ClearHostedForms
  ShowOutputs 0

End Sub

Private Sub cmdOutputs_Click()
  ResetActivityTime
  ClearHostedForms
  ShowGroups 0, 0
End Sub

Private Sub cmdOutputServers_Click()
  ResetActivityTime
  ClearHostedForms
  ShowOutputServers 0

End Sub

Private Sub cmdPrintInfo_Click()
'  ResetActivityTime
  If Printer Is Nothing Then Exit Sub
  ResetRemoteRefreshCounter
  PrintInfoWindow Val(txtHiddenResID.text)
End Sub

Private Sub cmdPrintScreen_Click()
  ShowHelp
End Sub
Sub ShowHelp()
        
        'Global Const LEVEL_FACTORY = 256
        'Global Const LEVEL_ADMIN = 128
        'Global Const LEVEL_SUPERVISOR = 32
        'Global Const LEVEL_USER = 1

        'User level logged in points to C:\HeritageMedCall\Help\HelpUser.chm
        'Admin1 level logged in points to C:\HeritageMedCall\Help\HelpAdmin1.chm
        'Admin2 level logged in points to C:\HeritageMedCall\Help\HelpAdmin2.chm
        'Factory level logged in points to C:\HeritageMedCall\Help\HelpFactory.chm
        Dim AppPath As String
        Dim filename As String

10      On Error GoTo ShowHelp_Error

20      AppPath = App.Path
30      If Right$(AppPath, 1) <> "\" Then
40        AppPath = AppPath & "\"
50      End If

60      'On Error Resume Next

70      If gUser.LEvel >= LEVEL_FACTORY Then
80        filename = AppPath & "Help\HelpFactory.chm"

90      ElseIf gUser.LEvel >= LEVEL_ADMIN Then
100       filename = AppPath & "Help\HelpAdmin2.chm"

110     ElseIf gUser.LEvel >= LEVEL_SUPERVISOR Then
120       filename = AppPath & "Help\HelpAdmin1.chm"

130     Else
140       filename = AppPath & "Help\HelpUser.chm"
150     End If

160     HtmlHelp Me.hwnd, filename, 0&, 0&


ShowHelp_Resume:
170     On Error GoTo 0
180     Exit Sub

ShowHelp_Error:

190     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowHelp." & Erl
200     Resume ShowHelp_Resume


End Sub

Sub EchostreamRequestConfig()
  Dim s As String
  Dim Buffer() As Byte
  ReDim Buffer(0 To 3)
  Dim Checksum As Integer


  s = Chr(&H30) & Chr(&H3) & Chr(&H7) & Chr(&H3D)
  Dim j As Integer
  Buffer(0) = &H30
  Buffer(1) = &H3
  Buffer(2) = &H7
  Buffer(3) = &H3A
  For j = 0 To 2
    Checksum = Checksum + Buffer(j)
  Next
  Buffer(4) = Checksum And &HFF
  WirelessPort.CommWriteByteArray Buffer, 5
  'WirelessPort.CommWrite s
End Sub


Private Sub cmdReports_Click()
  ResetActivityTime
  ClearHostedForms
  ShowReportMenu
End Sub

Private Sub cmdResidents_Click()
  ResetActivityTime
  Static Busy As Boolean
  If Busy Then Exit Sub
  Busy = True
  
  ClearHostedForms
  ShowResidents 0, 0, ""
  Busy = False
End Sub

Private Sub cmdResListDown_Click()
'  ResetActivityTime
  lvTxMoveDown
End Sub

Private Sub cmdResListPgDown_Click()
'  ResetActivityTime
  SendMessage lvtx.hwnd, WM_KEYDOWN, Win32.VK_NEXT, ByVal 0&
  SendMessage lvtx.hwnd, WM_KEYDOWN, Win32.VK_NEXT, ByVal 0&

End Sub

Private Sub cmdResListPgUp_Click()

'  ResetActivityTime
  SendMessage lvtx.hwnd, WM_KEYDOWN, Win32.VK_PRIOR, ByVal 0&
  SendMessage lvtx.hwnd, WM_KEYDOWN, Win32.VK_PRIOR, ByVal 0&


End Sub


Private Sub cmdResListPrintList_Click()
'  ResetActivityTime
If Printer Is Nothing Then Exit Sub
  Static Busy As Boolean
  If Busy Then Exit Sub
  Busy = True
  Select Case tabList.SelectedItem.index
    Case 2  ' by Room
      PrintRoomList lvtx, False
    Case 3  ' by Device
      PrintDeviceList lvtx, False
    Case Else  ' by Resident (index 1  or 0)
      PrintResList lvtx, False
  End Select


  Busy = False

End Sub

Private Sub cmdResListPrintView_Click()
'  ResetActivityTime
If Printer Is Nothing Then Exit Sub
  Static Busy As Boolean
  If Busy Then Exit Sub
  Busy = True
  Select Case tabList.SelectedItem.index
    Case 2  ' by Room
      PrintRoomList lvtx, True
    Case 3  ' by Device
      PrintDeviceList lvtx, True
    Case Else  ' by Resident (index 1  or 0)
      PrintResList lvtx, True
  End Select


  Busy = False
End Sub


'Private Sub cmdResListPrintView_Click()
'  Static busy As Boolean
'  If busy Then Exit Sub
'  busy = True
'    Select Case tabList.SelectedItem.index
'      Case 2  ' by Room
'        'PrintRoomList lvtx, True
'      Case 3  ' by Device
'        'PrintDeviceList lvtx, True
'      Case Else ' by Resident 1  or none
'        PrintResList lvtx, True
'    End Select
'
'
'  busy = False
'End Sub

Sub PrintResList(lvtx As ListView, ByVal Partial As Boolean)
  Dim ResList As cResList
  Set ResList = New cResList
  ResList.Partial = Partial
  ResList.PrintList lvtx
  Set ResList = Nothing
End Sub
Sub PrintRoomList(lvtx As ListView, ByVal Partial As Boolean)
  Dim RoomList As cRoomList
  Set RoomList = New cRoomList
  RoomList.Partial = Partial
  RoomList.PrintList lvtx
  Set RoomList = Nothing

End Sub
Sub PrintDeviceList(lvtx As ListView, ByVal Partial As Boolean)
  Dim DeviceList As cDeviceList
  Set DeviceList = New cDeviceList
  DeviceList.Partial = Partial
  DeviceList.PrintList lvtx
  Set DeviceList = Nothing


End Sub


Private Sub cmdResListUp_Click()
'  ResetActivityTime
  lvTxMoveUp

End Sub
Sub lvTxMoveUp()
  Dim i As Long
  If Not lvtx.SelectedItem Is Nothing Then
    i = lvtx.SelectedItem.index - 1
    If i > 0 Then
      lvtx.SelectedItem = lvtx.ListItems(i)
      lvtx.SelectedItem.EnsureVisible
    End If
    ShowLVTXData lvtx.SelectedItem.Key
  End If
End Sub
Sub lvTxMoveDown()
  Dim i As Long
  If Not lvtx.SelectedItem Is Nothing Then
    i = lvtx.SelectedItem.index + 1
    If i <= lvtx.ListItems.Count Then
      lvtx.SelectedItem = lvtx.ListItems(i)
      lvtx.SelectedItem.EnsureVisible
    End If
    ShowLVTXData lvtx.SelectedItem.Key
  End If
End Sub

Private Sub cmdRooms_Click()
  ResetActivityTime
  Static Busy As Boolean
  If Busy Then Exit Sub
  Busy = True
  ClearHostedForms
  ShowRooms 0, 0, 0, ""
  Busy = False
End Sub

Private Sub cmdSetup_Click()
  ResetActivityTime
  ClearHostedForms
  ShowConfigure1
End Sub

Private Sub cmdSilence_Click()
'  ResetActivityTime
  
  If MASTER Then
    Call Silence(ConsoleID, "")
  Else
  
    ResetRemoteRefreshCounter
    If Configuration.BeepControl Then
      Silence ConsoleID, ""
      If gMyAlarms Then
        ClientGetSubscribedAlarms
      Else
        ClientGetAlarms
      End If
      
    Else
      ClientSilenceAlarms ConsoleID, Configuration.RemoteSerial, ""
    ResetRemoteRefreshCounter
    End If
  End If
End Sub
Sub Silence(ByVal ConsoleID As String, ByVal Alarmtype As String)
  If MASTER Then
    SilenceAlarms gUser.Username, "MASTER", "MASTER", Alarmtype
  Else
    'call ClientSilenceAlarms(ConsoleID, RemoteSerial, Alarmtype)"
    QueEvent "ClientSilenceAlarms", ConsoleID, "Alarms"
    QueEvent "ClientSilenceAlarms", ConsoleID, "Alerts"
    QueEvent "ClientSilenceAlarms", ConsoleID, "Troubles"
    QueEvent "ClientSilenceAlarms", ConsoleID, "LowBatts"
    QueEvent "ClientSilenceAlarms", ConsoleID, "Externs"
  End If
End Sub

Sub UnSilence(ByVal ConsoleID As String, ByVal RemoteSerial As String, ByVal Alarmtype As String)
  If MASTER Then
    UnSilenceAlarms gUser.Username, "MASTER", "MASTER", Alarmtype
  Else
    Call ClientUnSilenceAlarms(ConsoleID, RemoteSerial, Alarmtype)
  End If

End Sub

Public Sub UnSilenceAlarms(ByVal User As String, ByVal ConsoleID As String, ByVal RemoteSerial As String, ByVal Alarmtype As String)

  Dim Key                As Long
  Dim li                 As ListItem
  Dim alarm              As cAlarm
  Dim rc                 As Long
  On Error GoTo UnSilenceAlarms_Error

  If (Alarmtype = "Alarms") Or (Alarmtype = "") Then


    For Each alarm In alarms.alarms
      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
        rc = alarm.ConsoleUnSilence(ConsoleID, User)
      End If
    Next
    If MASTER Then
      alarms.ConsoleSilenceTime(ConsoleID) = 0
      alarms.ConsoleAlarmTime(ConsoleID) = CDbl(Now)
      
    End If
    If ConsoleID = "MASTER" Then
      Debug.Print "FrmMain.UnSilenceAlarms"
      alarms.LocalSilenceTime = 0
      alarms.LocalAlarmTime = CDbl(Now)
      
    End If

  End If

  If (Alarmtype = "Alerts") Or (Alarmtype = "") Then

    For Each alarm In Alerts.alarms
      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
        rc = alarm.ConsoleUnSilence(ConsoleID, User)
      End If
    Next

    If MASTER Then
      Alerts.ConsoleAlarmTime(ConsoleID) = CDbl(Now)
      Alerts.ConsoleSilenceTime(ConsoleID) = 0
    End If
    If ConsoleID = "MASTER" Then
      Alerts.LocalAlarmTime = CDbl(Now)
      Alerts.LocalSilenceTime = 0
    End If

  End If

  If (Alarmtype = "LowBatts") Or (Alarmtype = "") Then

    For Each alarm In LowBatts.alarms
      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
        rc = alarm.ConsoleUnSilence(ConsoleID, User)
      End If
    Next
    If MASTER Then
      LowBatts.ConsoleAlarmTime(ConsoleID) = CDbl(Now)
      LowBatts.ConsoleSilenceTime(ConsoleID) = 0
    End If

    If ConsoleID = "MASTER" Then
      LowBatts.LocalAlarmTime = CDbl(Now)
      LowBatts.LocalSilenceTime = 0
    End If

  End If


  If (Alarmtype = "Troubles") Or (Alarmtype = "") Then
    For Each alarm In Troubles.alarms
      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
        rc = alarm.ConsoleUnSilence(ConsoleID, User)
      End If
    Next
    If MASTER Then
      Troubles.ConsoleAlarmTime(ConsoleID) = CDbl(Now)
      Troubles.ConsoleSilenceTime(ConsoleID) = 0
    End If
    If ConsoleID = "MASTER" Then
      Troubles.LocalAlarmTime = CDbl(Now)
      Troubles.LocalSilenceTime = 0
    End If
  End If

  If (Alarmtype = "Externs") Or (Alarmtype = "") Then
    For Each alarm In Externs.alarms
      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
        rc = alarm.ConsoleUnSilence(ConsoleID, User)
      End If
    Next
    If MASTER Then
      Externs.ConsoleAlarmTime(ConsoleID) = CDbl(Now)
      Externs.ConsoleSilenceTime(ConsoleID) = 0
    End If
    If ConsoleID = "MASTER" Then
      Externs.LocalAlarmTime = CDbl(Now)
      Externs.LocalSilenceTime = 0
    End If


  End If

UnSilenceAlarms_Resume:
  On Error GoTo 0
  Exit Sub

UnSilenceAlarms_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.UnSilenceAlarms." & Erl
  Resume UnSilenceAlarms_Resume


End Sub

Public Sub SilenceAlarms(ByVal User As String, ByVal ConsoleID As String, ByVal RemoteSerial As String, ByVal Alarmtype As String)

        Dim Key                As Long
        Dim li                 As ListItem
        Dim alarm              As cAlarm
        Dim rc                 As Long
10      On Error GoTo SilenceAlarms_Error



20      If Alarmtype = "Alarms" Or Alarmtype = "" Then

30        For Each alarm In alarms.alarms

            '      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
            '        rc = alarm.ConsoleSilence(ConsoleID, User)
            '      Else
40          If alarm.Silenced = 0 Then
50            alarm.Silenced = 1
60            alarm.SilenceTime = Now
70            alarm.SilenceUser = User
80            alarm.ConsoleSilence ConsoleID, User
90          End If
            'End If
100       Next
110       alarms.BeepTimer = 0

          ' repeat for alerts etc
120       If MASTER Then
130         alarms.ConsoleAlarmTime(ConsoleID) = 0
140         alarms.ConsoleSilenceTime(ConsoleID) = CDbl(Now)
150       End If

160       Debug.Print "alarms.ConsoleAlarmTime(ConsoleID) " & alarms.ConsoleAlarmTime(ConsoleID)

170       If ConsoleID = "MASTER" Then
180         Debug.Print "FrmMain.SilenceAlarms"
190         alarms.LocalAlarmTime = 0
200         alarms.LocalSilenceTime = CDbl(Now)
210       End If

220     End If

230     If Alarmtype = "Alerts" Or Alarmtype = "" Then


240       For Each alarm In Alerts.alarms
            '      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
            '        rc = alarm.ConsoleSilence(ConsoleID, User)
            '      Else


250         If alarm.Silenced = 0 Then
260           alarm.Silenced = 1
270           alarm.SilenceTime = Now
280           alarm.SilenceUser = User
              '      For Each li In lvAlerts.ListItems
              '        If Alarm.Serial = Left(li.key, 8) Then
              '        End If
              '      Next
290         End If
            'End If
300       Next
310       Alerts.BeepTimer = 0

320       If MASTER Then
330         Alerts.ConsoleAlarmTime(ConsoleID) = 0
340         Alerts.ConsoleSilenceTime(ConsoleID) = CDbl(Now)
350       End If
360       If ConsoleID = "MASTER" Then
370         Alerts.LocalAlarmTime = 0
380         Alerts.LocalSilenceTime = CDbl(Now)
390       End If

400     End If

410     If Alarmtype = "LowBatts" Or Alarmtype = "" Then


420       For Each alarm In LowBatts.alarms
            '      If CBool(Configuration.BeepControl) Then  ' used when remote console has control of own alarm Beeps
            '        rc = alarm.ConsoleSilence(ConsoleID, User)
            '      Else

430         If alarm.Silenced = 0 Then
440           alarm.Silenced = 1
450           alarm.SilenceTime = Now
460           alarm.SilenceUser = User
470           If MASTER Then
480             For Each li In lvLoBatt.ListItems
490               If alarm.Serial = left(li.Key, 8) Then
500                 li.ListSubItems(6).text = Format(alarm.SilenceTime, gTimeFormatString)
510               End If
520             Next
530           End If
540         End If
            'End If
550       Next
560       LowBatts.BeepTimer = 0

570       If MASTER Then
580         LowBatts.ConsoleAlarmTime(ConsoleID) = 0
590         LowBatts.ConsoleSilenceTime(ConsoleID) = CDbl(Now)
600       End If

610       If ConsoleID = "MASTER" Then
620         LowBatts.LocalAlarmTime = 0
630         LowBatts.LocalSilenceTime = CDbl(Now)
640       End If

650     End If

660     If Alarmtype = "Troubles" Or Alarmtype = "" Then

670       For Each alarm In Troubles.alarms
680         If alarm.Silenced = 0 Then
690           alarm.Silenced = 1
700           alarm.SilenceTime = Now
710           alarm.SilenceUser = User
720           For Each li In lvCheckIn.ListItems
730             If alarm.Serial = left(li.Key, 8) Then
740               li.ListSubItems(6).text = Format(alarm.SilenceTime, gTimeFormatString)
750             End If
760           Next
770         End If
780       Next
790       Troubles.BeepTimer = 0
800       If MASTER Then
810         Troubles.ConsoleAlarmTime(ConsoleID) = 0
820         Troubles.ConsoleSilenceTime(ConsoleID) = CDbl(Now)
830       End If

840       If ConsoleID = "MASTER" Then
850         Troubles.LocalAlarmTime = 0
860         Troubles.LocalSilenceTime = CDbl(Now)
870       End If

880     End If


890     If Alarmtype = "Externs" Or Alarmtype = "" Then
900       For Each alarm In Externs.alarms
910         If alarm.Silenced = 0 Then
920           alarm.Silenced = 1
930           alarm.SilenceTime = Now
940           alarm.SilenceUser = User
              '      For Each li In lvext.ListItems
              '        If Alarm.Serial = Left(li.key, 8) Then
              '420             li.ListSubItems(5).text = Format(Alarm.SilenceTime, gTimeFormatString)
              '        End If
              '      Next
950         End If
960       Next
970       Externs.BeepTimer = 0
980       If MASTER Then
990         Externs.ConsoleAlarmTime(ConsoleID) = 0
1000        Externs.ConsoleSilenceTime(ConsoleID) = CDbl(Now)
1010      End If

1020      If ConsoleID = "MASTER" Then
1030        Externs.LocalAlarmTime = 0
1040        Externs.LocalSilenceTime = CDbl(Now)
1050      End If

1060    End If

SilenceAlarms_Resume:
1070    On Error GoTo 0
1080    Exit Sub

SilenceAlarms_Error:

1090    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.SilenceAlarms." & Erl
1100    Resume SilenceAlarms_Resume


End Sub

Private Sub cmdTransmitters_Click()
  
  ResetActivityTime
  Static Busy As Boolean
  If Busy Then Exit Sub
  Busy = True

  ClearHostedForms
  ShowTransmitters 0, 0
  Busy = False

End Sub

Private Sub cmdTrouble_Click()
'  ResetActivityTime
  ResetRemoteRefreshCounter
  AckTrouble
  ShowCheckinInfo
End Sub




Private Sub Form_Click()
'  ResetActivityTime
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  ResetActivityTime
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  'Debug.Print "Form_KeyPress FrmMain"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  'Debug.Print "Form_KeyUP FrmMain"
End Sub

Private Sub Form_Load()
  Dim s As String
  sizeform
  Me.top = 0
  Me.left = 0
  starttime = Now
  Connect
  Set CurrentResident = New cResident
  ArrangeControls

  SetListTabs

  ' was here!

  CurrentGrid = InfoPane.emergency

  Me.Refresh
  DisplayResidentInfo 0, 0
  imgPacket.Picture = LoadResPicture(1005, vbResBitmap)  '
  'Load frmTimer ' dupe in sub_main
  UpdateScreenElements


End Sub

Sub Fill_lvtx(ByVal SortOrder As Integer, Optional searchstring As String)
        ' only one caller allowed!
        ' caller must prevent reentrancy

        Dim t                  As Long

        Dim CurrentPass        As Long
        Static passnumber      As Long

        '        Debug.Print "Entering Pass " & CurrentPass

        Dim rs                 As Recordset
        Dim li                 As ListItem

        Dim modrow             As Boolean
        Dim SQL                As String
        Dim text               As String

        Dim counter            As Long

        Dim ResidentID         As Long

10      On Error GoTo Fill_lvtx_Error

20      searchstring = FixQuotes(searchstring)
30      lvtx.ListItems.Clear


40      Select Case SortOrder

          Case SORT_ROOM             ' by room


50          If lvtxCols_Room Is Nothing Then
60            Set lvtxCols_Room = New cGridColumns
70            lvtxCols_Room.Size(1) = 2000
80            lvtxCols_Room.Size(2) = 2000
90          End If

            Do While lvtx.ColumnHeaders.Count >= 3
               lvtx.ColumnHeaders.Remove 3
            Loop


100         lvtx.ColumnHeaders(1).text = "Room"  '
110         lvtx.ColumnHeaders(1).Width = lvtxCols_Room.Size(1)
120         lvtx.ColumnHeaders(2).text = "Resident"
130         lvtx.ColumnHeaders(2).Width = lvtxCols_Room.Size(2)

140         LockWindowUpdate lvtx.hwnd

150         t = Win32.timeGetTime

160         SQL = "SELECT distinct Rooms.RoomID, Rooms.Room, Rooms.Building, Devices.ResidentID, Residents.Namefirst, Residents.NameLast , residents.Phone " & _
                  " FROM (Devices RIGHT JOIN Rooms ON Devices.RoomID = Rooms.RoomID) " & _
                  " LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID ORDER BY Rooms.Room"

170         Set rs = ConnExecute(SQL)

180         Debug.Print "SQL Sort by rooms " & Format(0.001 * (Win32.timeGetTime - t), "0.000")

190         lvtx.Sorted = True
200         lvtx.SortOrder = lvwAscending
210         lvtx.SortKey = 0

220         t = Win32.timeGetTime

230         Do Until rs.EOF

240           counter = counter + 1
250           Set li = lvtx.ListItems.Add(, rs("residentID") & "R" & rs("RoomID") & "ID" & counter, rs("Room"))

260           text = ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")

270           If Len(text) = 0 Then
280             li.SubItems(1) = " "
290           Else
300             li.SubItems(1) = text
310           End If
320           rs.MoveNext

330         Loop
340         rs.Close
350         LockWindowUpdate 0
360         Debug.Print "Fill Sort by rooms " & Format(0.001 * (Win32.timeGetTime - t), "0.000")


370       Case SORT_SERIAL           ' by device serial

380         t = Win32.timeGetTime

            Do While lvtx.ColumnHeaders.Count >= 3
               lvtx.ColumnHeaders.Remove 3
            Loop


390         If lvtxCols_Serial Is Nothing Then
400           Set lvtxCols_Serial = New cGridColumns
410           lvtxCols_Serial.Size(1) = 1050
420           lvtxCols_Serial.Size(2) = 2400
430         End If
440         lvtx.ColumnHeaders(1).text = "ID"
450         lvtx.ColumnHeaders(1).Width = lvtxCols_Serial.Size(1)
460         lvtx.ColumnHeaders(2).text = "Resident/Room"
470         lvtx.ColumnHeaders(2).Width = lvtxCols_Serial.Size(2)
480         lvtx.Sorted = True

490         LockWindowUpdate lvtx.hwnd
500         SQL = " SELECT Devices.Serial,Devices.deviceid, Devices.ResidentID, Residents.NameLast, Residents.NameFirst, Rooms.Room, Rooms.Building " & _
                  " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) " & _
                  " LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID order by  Devices.Serial "

510         Set rs = ConnExecute(SQL)

520         Debug.Print "SQL Sort by serial " & Format(0.001 * (Win32.timeGetTime - t), "0.000")


530         Do Until rs.EOF
540           Set li = lvtx.ListItems.Add(, rs("ResidentID") & "@" & rs("deviceid"), Right("00000000" & rs("serial"), 8))
550           text = ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
560           li.SubItems(1) = text & " / " & rs("Room")
570           rs.MoveNext
580         Loop
590         rs.Close
600         LockWindowUpdate 0

610         Debug.Print "Fill Sort by serial " & Format(0.001 * (Win32.timeGetTime - t), "0.000")

620       Case Else                  ' by resident

630         t = Win32.timeGetTime

640         If lvtxCols_Res Is Nothing Then
650           Set lvtxCols_Res = New cGridColumns
660           lvtxCols_Res.Size(1) = 2000
670           lvtxCols_Res.Size(2) = 2000
680           lvtxCols_Res.Size(3) = 2000
690         End If


700         lvtx.ColumnHeaders(1).text = "Name"  ' full name
710         lvtx.ColumnHeaders(1).Width = lvtxCols_Res.Size(1)
720         lvtx.ColumnHeaders(2).text = "Room"  ' Room
730         lvtx.ColumnHeaders(2).Width = lvtxCols_Res.Size(2)
            
            If lvtx.ColumnHeaders.Count < 3 Then
               lvtx.ColumnHeaders.Add 3
            End If
740         lvtx.ColumnHeaders(3).text = "Phone"  ' Room
750         lvtx.ColumnHeaders(3).Width = lvtxCols_Res.Size(3)


760         lvtx.Sorted = False

770         LockWindowUpdate lvtx.hwnd

780         SQL = "SELECT DISTINCT  Residents.ResidentID, Residents.NameLast, Residents.NameFirst, residents.phone, Rooms.Room " & _
                  "FROM Rooms RIGHT JOIN (Residents LEFT JOIN Devices ON Residents.ResidentID = Devices.ResidentID) ON Rooms.RoomID = Devices.RoomID " & _
                  "WHERE Residents.deleted = 0 " & _
                  "ORDER BY NameLast, NameFirst ;"


790         Set rs = ConnExecute(SQL)

800         Debug.Print "SQL Sort by resident " & Format(0.001 * (Win32.timeGetTime - t), "0.000")
810         t = Win32.timeGetTime

820         Do Until rs.EOF
              'DoEvents

              'If CurrentPass <> passnumber Then Exit Do
              'Debug.Print "CurrentPass Data " & CurrentPass
              'Set Li = lvtx.ListItems.Add(, rs("ResidentID") & " " & rs("DeviceID"), rs("Namelast") & ", " & rs("NameFirst"))
830           If ResidentID <> rs("ResidentID") Then
840             Set li = lvtx.ListItems.Add(, rs("ResidentID") & "@", ConvertLastFirst(rs("namelast") & "", rs("namefirst") & ""))

850             If Len(rs("Room") & "") > 0 Then
860               li.SubItems(1) = rs("Room")
870             Else
880               li.SubItems(1) = ""
890             End If

900             If Len(rs("Phone") & "") > 0 Then
910               li.SubItems(2) = rs("Phone")
920             Else
930               li.SubItems(2) = ""
940             End If


950           Else
960             If Len(rs("Room") & "") > 0 Then
970               If Len(li.SubItems(1)) Then
980                 li.SubItems(1) = "\" & rs("Room")
990               Else
1000                li.SubItems(1) = rs("Room")
1010              End If
1020            End If

1030            If Len(rs("Phone") & "") > 0 Then
1040              li.SubItems(2) = rs("Phone")
1050            Else
1060              li.SubItems(2) = ""
1070            End If

1080          End If
1090          ResidentID = rs("ResidentID")
              'li.SubItems(1) = GetResidentRooms(rs("ResidentID") & "")  '              rs("serial") & ""
1100          rs.MoveNext

1110        Loop
1120        rs.Close
1130        LockWindowUpdate 0
1140        Debug.Print "Fill Sort by resident " & Format(0.001 * (Win32.timeGetTime - t), "0.000")
1150    End Select
        'Debug.Print "Exiting Pass " & CurrentPass
Fill_lvtx_Resume:
1160    Set rs = Nothing
1170    On Error GoTo 0
1180    Exit Sub

Fill_lvtx_Error:

1190    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.Fill_lvtx." & Erl
1200    Resume Fill_lvtx_Resume


End Sub


Sub sizeform()
  If Me.WindowState = vbNormal Then
    Me.Width = 1024 * Screen.TwipsPerPixelX
    'If large Then
    If 1 Then
      Me.Height = 768 * Screen.TwipsPerPixelX
    Else
      Me.Height = fraAlerts.top + fraAlerts.Height + (Me.Height - Me.ScaleHeight)
    End If

  End If
  picAlarms.top = 0
  picAlarms.left = 0

  picAlarms.Height = Me.ScaleHeight
  ' pic alarms is the container for buttons down the left side of main screen
  
End Sub

Sub ArrangeControls()
  Dim f As Control
  For Each f In Me.Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor
    End If
  Next




  'Center windows
  fraAlerts.Width = lvEmergency.Width + 60
  lblAlerts.Width = lblAlarms.Width
  lblLowBattery.Width = lblAlarms.Width
  lblTrouble.Width = lblAlarms.Width
  lblInformation.Width = lblAlarms.Width
  lblAlerts.left = 15
  lblInformation.left = 15
  lblLowBattery.left = 15
  lblTrouble.left = 15
  lblInformation.left = 15
  picResident.Width = lblAlarms.Width
  
  
  
  
  lvEmergency.top = lblAlarms.Height
  lvAlerts.top = lblAlerts.Height
  lvLoBatt.top = lblLowBattery.Height
  lvCheckIn.top = lblTrouble.Height
  picResident.top = lblInformation.Height
  
  
  fraLoBatt.left = fraAlerts.left
  fraLocate.left = fraAlerts.left
  fraExternal.left = fraAlerts.left
  fraCheckin.left = fraAlerts.left
  fraResinfo.left = fraAlerts.left
  
  fraLoBatt.Width = fraAlerts.Width
  fraLocate.Width = fraAlerts.Width
  fraExternal.Width = fraAlerts.Width
  fraCheckin.Width = fraAlerts.Width
  fraResinfo.Width = fraAlerts.Width

  fraLoBatt.top = fraAlerts.top
  fraLocate.top = fraAlerts.top
  fraExternal.top = fraAlerts.top
  fraCheckin.top = fraAlerts.top

  fraHippaList.BackColor = Me.BackColor


  fraHost.left = fraResinfo.left
  fraHost.top = fraResinfo.top + fraResinfo.Height
  'fraHost.BackColor = vbRed

  ' righthand  controls
  fraTransmitters.Width = tabList.ClientWidth
  fraTransmitters.left = tabList.ClientLeft
  fraTransmitters.top = tabList.ClientTop

  lvtx.left = 0
  lvtx.top = 0
  lvtx.Width = fraTransmitters.Width


  ConfigurelvEmergency
  ConfigurelvAlerts
  ConfigurelvLoBatt
  ConfigurelvTrouble
  ConfigurelvAssur
  ConfigurelvExtern

  txtInfoFullName.BackColor = picResident.BackColor
  txtInfoMessage.BackColor = picResident.BackColor
  txtInfoRoom.BackColor = picResident.BackColor
  txtAssurDays.BackColor = picResident.BackColor
  txtInfox.BackColor = picResident.BackColor
  txtInfoNotes.BackColor = picResident.BackColor
End Sub

Sub UpdateLayout()

  lblAlarms.BackColor = ColorInActive
  lblAlerts.BackColor = ColorInActive
  lblLowBattery.BackColor = ColorInActive
  lblTrouble.BackColor = ColorInActive
  lblExternal.BackColor = ColorInActive
  Select Case mCurrentGrid


    Case InfoPane.Checkin  ' trouble
      fraCheckin.Visible = True
      fraLoBatt.Visible = False
      fraAlerts.Visible = False
      fraExternal.Visible = False
      fraExternal.Visible = False
      lblTrouble.BackColor = ColorActive
      SetFocusTo lvCheckIn

    Case InfoPane.LowBatt  ' Lo Batt

      fraLoBatt.Visible = True
      fraCheckin.Visible = False
      fraAlerts.Visible = False
      fraExternal.Visible = False
      lblLowBattery.BackColor = ColorActive
      SetFocusTo lvLoBatt

    Case InfoPane.Alert  ' Alerts
      fraAlerts.Visible = True
      fraCheckin.Visible = False
      fraLoBatt.Visible = False
      fraExternal.Visible = False
      lblAlerts.BackColor = ColorActive
      SetFocusTo lvAlerts
    
    Case InfoPane.extern ' external devices
      fraExternal.Visible = True
      fraCheckin.Visible = False
      fraLoBatt.Visible = False
      fraAlerts.Visible = False
      lblExternal.BackColor = ColorActive
      SetFocusTo lvExternal
    
    
    Case InfoPane.emergency
      lblAlarms.BackColor = ColorActive
  
  End Select


End Sub

Sub FitAlarms()
  Dim w As Double
  w = lblAlarms.left
  w = lblAlarms.Width
  w = lvEmergency.left
  w = lvEmergency.Width
  w = Me.lvAlerts.Width ' 8895
w = fraHippaList.left
w = fraHippaList.Width
w = Me.Width

    Dim PanelWidth As Double

    If Configuration.HideHIPPASidebar <> 0 Then
      
      cmdAlarmUp.left = Me.Width - (120 + cmdAlarmUp.Width)
      
      PanelWidth = cmdAlarmUp.left - (lvEmergency.left + 120)
      
      lvEmergency.Width = PanelWidth
      lblAlarms.Width = PanelWidth
      lvAlerts.Width = PanelWidth
      fraAlerts.Width = PanelWidth
      lblAlerts.Width = PanelWidth
      
      fraExternal.Width = PanelWidth
      lvExternal.Width = PanelWidth
      lblExternal.Width = PanelWidth
      
      
      fraCheckin.Width = PanelWidth
      lvCheckIn.Width = PanelWidth
      lblTrouble.Width = PanelWidth
      
      
      
      fraLoBatt.Width = PanelWidth
      lvLoBatt.Width = PanelWidth
      lblLowBattery.Width = PanelWidth
      
      fraResinfo.Width = PanelWidth
      lblInformation.Width = PanelWidth
      
      picResident.Width = PanelWidth
      lblInformation.Width = PanelWidth
      
      
      
      
      
      cmdAlarmUp.left = cmdAlarmUp.left  ' Me.ScaleWidth '+ 60
      cmdAlarmdown.left = cmdAlarmUp.left
      cmdAlarmPrintList.left = cmdAlarmUp.left
      cmdMultiUp.left = cmdAlarmUp.left
      cmdMultiDown.left = cmdAlarmUp.left
      cmdMultiPrint.left = cmdAlarmUp.left
      CmdEditInfo.left = cmdAlarmUp.left
      cmdClearInfo.left = cmdAlarmUp.left
      cmdPrintInfo.left = cmdAlarmUp.left
    Else
      cmdAlarmUp.left = fraHippaList.left - (cmdAlarmUp.Width + 120)
      
      PanelWidth = cmdAlarmUp.left - (lvEmergency.left + 120)
      
      
      lvEmergency.Width = PanelWidth
      lblAlarms.Width = PanelWidth
      lvAlerts.Width = PanelWidth
      fraAlerts.Width = PanelWidth
      lblAlerts.Width = PanelWidth
      
      fraExternal.Width = PanelWidth
      lvExternal.Width = PanelWidth
      lblExternal.Width = PanelWidth
      
      
      fraCheckin.Width = PanelWidth
      lvCheckIn.Width = PanelWidth
      lblTrouble.Width = PanelWidth
      
      
      
      fraLoBatt.Width = PanelWidth
      lvLoBatt.Width = PanelWidth
      lblLowBattery.Width = PanelWidth
      
      fraResinfo.Width = PanelWidth
      lblInformation.Width = PanelWidth
      
      picResident.Width = PanelWidth
      lblInformation.Width = PanelWidth
      
      

      cmdAlarmdown.left = cmdAlarmUp.left
      cmdAlarmPrintList.left = cmdAlarmUp.left
      cmdMultiUp.left = cmdAlarmUp.left
      cmdMultiDown.left = cmdAlarmUp.left
      cmdMultiPrint.left = cmdAlarmUp.left
      CmdEditInfo.left = cmdAlarmUp.left
      cmdClearInfo.left = cmdAlarmUp.left
      cmdPrintInfo.left = cmdAlarmUp.left
      
    End If
    
    
End Sub


Private Sub Form_Paint()
  If Configuration.HideHIPPASidebar <> 0 Then
    If fraHippaList.Visible Then
      fraHippaList.Visible = False
      FitAlarms
    End If
  Else

    If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
      If fraHippaList.Visible Then
        fraHippaList.Visible = False
        FitAlarms
      End If
    Else
      If Not fraHippaList.Visible Then
        fraHippaList.Visible = True
        FitAlarms
      End If
    End If
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'vbFormControlMenu  The user chose the Close command from the Control menu on the form.
'vbFormCode         The Unload statement is invoked from code.
'vbAppWindows       The current Microsoft Windows operating environment session is ending.
'vbAppTaskManager   The Microsoft Windows Task Manager is closing the application.
'vbFormMDIForm      MDI child form is closing because the MDI form is closing.
'vbFormOwner        Form is closing because its owner is closing


  Select Case UnloadMode
    Case vbFormControlMenu
      Cancel = True
    Case vbFormCode
    Case vbAppTaskManager
    Case vbAppWindows
  End Select
  StopIt = Not Cancel
End Sub

Private Sub Form_Resize()

  If Configuration.HideHIPPASidebar <> 0 Then
    If fraHippaList.Visible Then
      fraHippaList.Visible = False
      FitAlarms
    End If
  Else

    If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
      If fraHippaList.Visible Then
        fraHippaList.Visible = False
        FitAlarms
      End If
    Else
      If Not fraHippaList.Visible Then
        fraHippaList.Visible = True
        FitAlarms
      End If
    End If
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Error Resume Next
  frmTimer.Timer1.Enabled = False
  
  ' cleanup listener socket
  If MASTER Then
    
'    If Not (Listener Is Nothing) Then
'      Listener.Close
'      Set Listener = Nothing
'
'    End If
    
  Else
    RemoteQuitting = True
    Sleep 50
    If Not HostConnection Is Nothing Then
      HostConnection.CloseConnection
    End If
    If Not HostInterraction Is Nothing Then
      HostInterraction.CloseConnection
    End If
    Sleep 50
    Set HostConnection = Nothing
    Set HostInterraction = Nothing

  End If
  
  
  If MASTER Then
    Set Enroller = Nothing
    Set gDukane = Nothing
    CloseComm WirelessPort
    Set DialogicSystem = Nothing
    Sleep 50
  End If
  

  Set WirelessPort = Nothing
  conn.Close
  Close ' closes all filehandles
  
  Dim f As Form
  
  For Each f In Forms
    If Not f Is Me Then
      Unload f
      DoEvents
    End If
  Next
  
  For Each f In Forms
    If Not f Is Me Then
      Unload f
      DoEvents
    End If
  Next


End Sub


Sub ProcessAssurs(ByVal EndOfPeriod As Boolean)

  Dim li                 As ListItem
  Dim alarm              As cAlarm
  Dim j                  As Integer


  Dim rs                 As Recordset
  Dim SQL                As String


  Dim Col0               As String
  Dim col1               As String
  Dim col2               As String
  Dim col3               As String

  Dim Assuritems         As Collection
  Dim Assuritem          As cAssureItem

  On Error GoTo ProcessAssurs_Error

  Set Assuritems = New Collection

  Dim InClause           As String

  Dim AssurListSerial()  As String
  ReDim AssurListSerial(0)


  If Assurs.alarms.Count Then




    ReDim AssurListSerial(1 To Assurs.alarms.Count)
    For j = 1 To Assurs.alarms.Count
      Set alarm = Assurs.alarms(j)
      AssurListSerial(j) = "'" & alarm.Serial & "'"
    Next
    InClause = Join(AssurListSerial, ",")



    SQL = "SELECT Devices.Serial, Devices.DeviceID, Residents.NameLast, Rooms.Room, Residents.NameFirst,  Residents.phone,  Devices.RoomID, Devices.ResidentID " & _
        " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
        " WHERE Devices.residentid <> 0 AND (Devices.Serial In (" & InClause & ")) " & _
        " ORDER BY Residents.NameLast, Residents.NameFirst, Rooms.Room; "

    Set rs = ConnExecute(SQL)
    Do Until rs.EOF
      Set Assuritem = New cAssureItem
      Assuritem.DeviceID = rs("deviceid")
      Assuritem.RoomID = Val("" & rs("RoomID"))
      Assuritem.ResidentID = Val("" & rs("ResidentID"))
      Assuritem.Serial = Right$("    " & rs("Serial"), 8)

      
      
      
      ' resident name and phone
      If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
        Assuritem.NameFirst = ""
        Assuritem.NameLast = ""
        Assuritem.NameFull = ""
      Else

        Assuritem.NameFirst = rs("namefirst") & ""
        Assuritem.NameLast = rs("namelast") & ""
        Assuritem.NameFull = ConvertLastFirst(Assuritem.NameLast, Assuritem.NameFirst)
      End If

      Assuritem.Phone = rs("Phone") & ""


      ' roomname
      Assuritem.Room = rs("Room") & ""

      'Alarm.ResidentText = Assuritem.NameFull  ' ?? not needed
      'Alarm.RoomText = Assuritem.room  ' ?? not needed
      'Alarm.Phone = Assuritem.Phone  ' ?? not needed
      Assuritems.Add Assuritem
      rs.MoveNext
    Loop
    rs.Close


    SQL = "SELECT Devices.Serial, Devices.DeviceID, Residents.NameLast, Rooms.Room, Residents.NameFirst,  Residents.phone,  Devices.RoomID, Devices.ResidentID " & _
        " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
        " WHERE Devices.residentid = 0 AND (Devices.Serial In (" & InClause & ")) " & _
        " ORDER BY Rooms.Room; "




    Set rs = ConnExecute(SQL)
    Do Until rs.EOF
      Set Assuritem = New cAssureItem
      Assuritem.DeviceID = rs("deviceid")
      Assuritem.RoomID = Val("" & rs("RoomID"))
      Assuritem.ResidentID = Val("" & rs("ResidentID"))
      Assuritem.Serial = Right$("    " & rs("Serial"), 8)




      ' resident name and phone
      If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
        Assuritem.NameFirst = ""
        Assuritem.NameLast = ""
        Assuritem.NameFull = ""
      Else

        Assuritem.NameFirst = rs("namefirst") & ""
        Assuritem.NameLast = rs("namelast") & ""
        Assuritem.NameFull = ConvertLastFirst(Assuritem.NameLast, Assuritem.NameFirst)
      End If



      Assuritem.NameFirst = rs("namefirst") & ""
      Assuritem.NameLast = rs("namelast") & ""
      Assuritem.NameFull = ConvertLastFirst(Assuritem.NameLast, Assuritem.NameFirst)




      Assuritem.Phone = rs("Phone") & ""


      ' roomname
      Assuritem.Room = rs("Room") & ""

      '450         Alarm.ResidentText = Assuritem.NameFull  ' ?? not needed
      '460         Alarm.RoomText = Assuritem.room  ' ?? not needed
      '470         Alarm.Phone = Assuritem.Phone  ' ?? not needed
      Assuritems.Add Assuritem
      rs.MoveNext
    Loop
    rs.Close

    Set rs = Nothing

  End If



'Debug.Print "ProcessAssurs " & Now

  If EndOfPeriod Then

    If Configuration.AssurSaveAsFile Or Configuration.AssurSendAsEmail Then
      AutoSendAssur Assuritems
    End If

  End If

  lvAssur.ListItems.Clear

  If (gAssurDisableScreenOutput = 0) Then

    lvAssur.Sorted = False
    LockWindowUpdate lvAssur.hwnd
    lvAssur.ListItems.Clear


    Dim SR               As String
    For j = 1 To Assuritems.Count
      Set Assuritem = Assuritems(j)

      SR = Assuritem.Serial
      SR = Assuritems(j).Serial



      Set li = lvAssur.ListItems.Add(, SR & "S", SR)
      If Len(Assuritem.NameFull) = 0 Then  ' name
        li.SubItems(1) = " "
      Else
        li.SubItems(1) = Assuritem.NameFull
      End If
      If Len(Assuritem.Room) = 0 Then  ' room
        li.SubItems(2) = " "
      Else
        li.SubItems(2) = Assuritem.Room
      End If
      If Len(Assuritem.Phone) = 0 Then  ' phone
        li.SubItems(3) = " "
      Else
        li.SubItems(3) = Assuritem.Phone
      End If
    Next

  End If
  '''' ------- old




ProcessAssurs_Resume:
  On Error GoTo 0
  LockWindowUpdate 0
  Exit Sub

ProcessAssurs_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ProcessAssurs." & Erl
  Resume ProcessAssurs_Resume


End Sub

Sub ProcessBatts()
        ' 32k alarms max

        Dim li                 As ListItem
        Dim j                  As Integer
        'Dim Device        As cESDevice
        Dim rs                 As Recordset
        Dim SQL                As String


        Dim COL_SERIAL         As String
        Dim COL_NAME           As String
        Dim COL_ROOM           As String
        Dim col_model          As String
        Dim col_desc           As String

        Dim alarm              As cAlarm

10      On Error GoTo ProcessBatts_Error


        Dim SelectedItem       As ListItem
        Dim SelectedSerialnum  As String

20      Set SelectedItem = lvLoBatt.SelectedItem
30      If Not SelectedItem Is Nothing Then
40        SelectedSerialnum = lvLoBatt.SelectedItem.Key
50      End If


60      LockWindowUpdate lvLoBatt.hwnd

70      lvLoBatt.ListItems.Clear

'80      For j = 1 To LowBatts.alarms.Count

80      For j = LowBatts.alarms.Count To 1 Step -1

90        COL_SERIAL = ""
100       COL_NAME = ""
110       COL_ROOM = ""
120       col_model = ""
130       col_desc = ""


140       Set alarm = LowBatts.alarms(j)
          'Set Device = Devices.Device(Alarm.Serial) ' validates device
          'If Not Device Is Nothing Then


'150       SQl = " SELECT Devices.Serial,Devices.Model, Residents.NameLast, Residents.NameFirst, Rooms.Room " & _
'              " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
'              " WHERE  Devices.serial =" & q(Alarm.Serial)
'          '100       SQl = " SELECT custom, DeviceID, Serial,model, ResidentID, RoomID FROM Devices WHERE Devices.serial =" & q(Alarm.Serial)

160       COL_SERIAL = alarm.Serial
170       col_desc = IIf(Len(alarm.Custom), alarm.Custom, alarm.Description)

'180       Set rs = ConnExecute(SQl)
'190       If Not rs.EOF Then
200
          'Set Device = Devices.Device(Alarm.Serial)
            col_model = alarm.Model
            
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              COL_NAME = ""
            Else
            
210           COL_NAME = alarm.ResidentText '  ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
            End If
220         COL_ROOM = alarm.RoomText ' rs("Room") & ""
'230       End If
'240       rs.Close

250       Set li = lvLoBatt.ListItems.Add(, alarm.Serial & "S" & alarm.Alarmtype, alarm.Serial)




260       If Len(COL_NAME) = 0 Then
270         li.SubItems(1) = " "
280       Else
290         li.SubItems(1) = COL_NAME
300       End If

310       If Len(COL_ROOM) = 0 Then
320         li.SubItems(2) = " "
330       Else
340         li.SubItems(2) = COL_ROOM
350       End If

360       If Len(col_model) = 0 Then
370         li.SubItems(3) = " "
380       Else
390         li.SubItems(3) = col_model
400       End If

410       li.SubItems(4) = col_desc & ""

420       li.SubItems(5) = Format(alarm.DateTime, gTimeFormatString)

430       li.SubItems(6) = IIf(alarm.SilenceTime <> 0, Format(alarm.SilenceTime, gTimeFormatString), " ")
          'End If
440     Next

450     For j = 1 To lvLoBatt.ListItems.Count
460       If SelectedSerialnum = lvLoBatt.ListItems(j).Key Then
470         lvLoBatt.ListItems(j).EnsureVisible
480         lvLoBatt.ListItems(j).Selected = True
490         Exit For
500       End If
510     Next


ProcessBatts_Resume:

520     On Error GoTo 0
530     LockWindowUpdate 0
540     Exit Sub

ProcessBatts_Error:

550     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ProcessBatts." & Erl
560     Resume ProcessBatts_Resume

End Sub
'Sub ShowAlertInfo()
'        Dim lv As ListView
'        Dim j As Integer
'        Dim Alarm As Object
'        Dim d     As cESDevice
'        Dim ResidentID As Long
'        Dim AlarmID As Long
'        Dim Serial  As String
'
'10      On Error GoTo ShowAlertInfo_Error
'
'20      Set lv = Me.lvAlerts
'
'30      Serial = left(GetlvAlertSelectedKey(), 8)
'40      For j = 1 To Alerts.alarms.count
'50        Set Alarm = Alerts.alarms(j)
'60        If 0 = StrComp(Alarm.Serial, Serial, vbTextCompare) Then
'            ResidentID = GetResidentIDFromSerial(Alarm.Serial)
'
'
''70          Set d = Devices.Device(alarm.serial)
''80          If Not d Is Nothing Then
'90            'ResidentID = d.ResidentID
'100           AlarmID = Alarm.id
'110         'End If
'120         Exit For
'130       End If
'140     Next
'150     DisplayResidentInfo ResidentID, AlarmID
'
'ShowAlertInfo_Resume:
'160     On Error GoTo 0
'170     Exit Sub
'
'ShowAlertInfo_Error:
'
'180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowAlertInfo." & Erl
'190     Resume ShowAlertInfo_Resume
'
'
'
'End Sub
Sub ShowEmergencyInfo()
        Dim lv                 As ListView
        Dim j                  As Integer
        Dim alarm              As cAlarm
        Dim d                  As cESDevice
        Dim ResidentID         As Long
        Dim RoomID             As Long
        Dim AlarmID            As Long
        Dim Serial             As String


10      On Error GoTo ShowEmergencyInfo_Error

20      Set lv = lvEmergency

30      AlarmID = lvKey2ID(lvEmergency.SelectedItem.Key) '  GetlvEmergencySelectedID()

        'Serial = left(GetlvEmergencySelectedKey(), 8)
40      For j = 1 To alarms.alarms.Count
50        Set alarm = alarms.alarms(j)

60        If (alarm.ID = AlarmID) Then
            'If 0 = StrComp(alarm.Serial, Serial, vbTextCompare) Then

70          ResidentID = alarm.ResidentID

80          RoomID = alarm.RoomID


            '            ResidentID = GetResidentIDFromSerial(alarm.Serial)
            '            RoomID = GetRoomIDFromSerial(alarm.Serial)

            '70          Set d = Devices.Device(alarm.serial)
            '80          If Not d Is Nothing Then
            'ResidentID = d.ResidentID
            ' AlarmID = alarm.ID
            'End If
90          Exit For
100       End If
110     Next
        '    DisplayResidentInfo ResidentID, AlarmID

120     DisplayResOrRoomInfo ResidentID, RoomID, AlarmID
ShowEmergencyInfo_Resume:
130     On Error GoTo 0
140     Exit Sub

ShowEmergencyInfo_Error:

150     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowEmergencyInfo." & Erl
160     Resume ShowEmergencyInfo_Resume


End Sub

Sub ShowCheckinInfo()
        Dim lv As ListView
        Dim j As Integer
        Dim alarm As Object
        Dim d     As cESDevice
        Dim ResidentID As Long
        Dim RoomID     As Long
        Dim AlarmID As Long
        Dim Serial  As String

10      On Error GoTo ShowCheckinInfo_Error

20      Set lv = lvCheckIn

30      Serial = left(GetlvCheckinSelectedKey(), 8)
40      For j = 1 To Troubles.alarms.Count
50        Set alarm = Troubles.alarms(j)
60        If 0 = StrComp(alarm.Serial, Serial, vbTextCompare) Then
            ResidentID = GetResidentIDFromSerial(alarm.Serial)
            RoomID = GetRoomIDFromSerial(alarm.Serial)

'70          Set d = Devices.Device(alarm.serial)
'80          If Not d Is Nothing Then
90            'ResidentID = d.ResidentID
100           AlarmID = alarm.ID
110         'End If
120         Exit For
130       End If
140     Next
   '    DisplayResidentInfo ResidentID, AlarmID

150     DisplayResOrRoomInfo ResidentID, RoomID, AlarmID
ShowCheckinInfo_Resume:
160     On Error GoTo 0
170     Exit Sub

ShowCheckinInfo_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowCheckinInfo." & Erl
190     Resume ShowCheckinInfo_Resume


End Sub


Function GetlvCheckinSelectedKey() As String
  If Not lvCheckIn.SelectedItem Is Nothing Then
    GetlvCheckinSelectedKey = lvCheckIn.SelectedItem.Key
  End If

End Function


Function GetlvAssurSelectedKey() As String
  If Not lvAssur.SelectedItem Is Nothing Then
    GetlvAssurSelectedKey = lvAssur.SelectedItem.Key
  End If

End Function


Function GetlvLowBattSelectedKey() As String
  If Not lvLoBatt.SelectedItem Is Nothing Then
    GetlvLowBattSelectedKey = lvLoBatt.SelectedItem.Key
  End If

End Function


Function GetlvEmergencySelectedKey() As String
  If Not lvEmergency.SelectedItem Is Nothing Then
    GetlvEmergencySelectedKey = lvEmergency.SelectedItem.Key
  End If

End Function
Function GetlvAlertSelectedKey() As String
  If Not lvAlerts.SelectedItem Is Nothing Then
    GetlvAlertSelectedKey = lvAlerts.SelectedItem.Key
  End If

End Function
Sub ShowAlertInfo()
        Dim lv As ListView
        Dim j As Integer
        Dim alarm As Object
        Dim d     As cESDevice
        Dim ResidentID As Long
        Dim RoomID     As Long
        Dim AlarmID As Long
        Dim Serial  As String

10      On Error GoTo ShowAlertInfo_Error

20      Set lv = lvAlerts

30      Serial = left(GetlvAlertSelectedKey(), 8)
40      For j = 1 To Alerts.alarms.Count
50        Set alarm = Alerts.alarms(j)
60        If 0 = StrComp(alarm.Serial, Serial, vbTextCompare) Then
            ResidentID = GetResidentIDFromSerial(alarm.Serial)
            RoomID = GetRoomIDFromSerial(alarm.Serial)

'70          Set d = Devices.Device(alarm.serial)
'80          If Not d Is Nothing Then
90            'ResidentID = d.ResidentID
100           AlarmID = alarm.ID
110         'End If
120         Exit For
130       End If
140     Next
   '    DisplayResidentInfo ResidentID, AlarmID

150     DisplayResOrRoomInfo ResidentID, RoomID, AlarmID
ShowAlertInfo_Resume:
160     On Error GoTo 0
170     Exit Sub

ShowAlertInfo_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowAlertInfo." & Erl
190     Resume ShowAlertInfo_Resume


End Sub

Sub ShowLowBattInfo()
        Dim lv As ListView
        Dim j As Integer
        Dim alarm As Object
        Dim d     As cESDevice
        Dim ResidentID As Long
        Dim RoomID     As Long
        Dim AlarmID As Long
        Dim Serial  As String

10      On Error GoTo ShowLowBattInfo_Error

20      Set lv = lvLoBatt

30      Serial = left(GetlvLowBattSelectedKey(), 8)
40      For j = 1 To LowBatts.alarms.Count
50        Set alarm = LowBatts.alarms(j)
60        If 0 = StrComp(alarm.Serial, Serial, vbTextCompare) Then
            ResidentID = GetResidentIDFromSerial(alarm.Serial)
            RoomID = GetRoomIDFromSerial(alarm.Serial)

'70          Set d = Devices.Device(alarm.serial)
'80          If Not d Is Nothing Then
90            'ResidentID = d.ResidentID
100           AlarmID = alarm.ID
110         'End If
120         Exit For
130       End If
140     Next
   '    DisplayResidentInfo ResidentID, AlarmID

150     DisplayResOrRoomInfo ResidentID, RoomID, AlarmID
ShowLowBattInfo_Resume:
160     On Error GoTo 0
170     Exit Sub

ShowLowBattInfo_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowLowBattInfo." & Erl
190     Resume ShowLowBattInfo_Resume


End Sub


Sub ShowResidentInfo(ByVal ResID As Long)

      ' get device associated with resident
        Dim lv As ListView

        Dim Key
        Dim j As Integer
        Dim alarm As cAlarm
        Dim d     As cESDevice
        Dim ResidentID As Long
        Dim AlarmID    As Long
        Dim CurrentAlarmID As Long

10      On Error GoTo ShowResidentInfo_Error

20      Select Case mCurrentGrid
          Case InfoPane.Checkin
30          Set lv = lvCheckIn
40          If Not lv.SelectedItem Is Nothing Then
              'key = lv.SelectedItem.key
50            AlarmID = lvKey2ID(lv.SelectedItem.Key)
60            For j = 1 To Troubles.alarms.Count
70              Set alarm = Troubles.alarms(j)
80              If AlarmID = alarm.AlarmID Then
                'If alarm.Serial = left(key, 8) Then
90                ResidentID = GetResidentIDFromSerial(alarm.Serial)
100               CurrentAlarmID = alarm.ID
110               Exit For
120             End If
130           Next
140         End If

150       Case InfoPane.Alert

160         Set lv = lvAlerts
170         If Not lv.SelectedItem Is Nothing Then
              'key = lv.SelectedItem.key
180           AlarmID = lvKey2ID(lv.SelectedItem.Key)
190           For j = 1 To Alerts.alarms.Count
200             Set alarm = Alerts.alarms(j)
210             If AlarmID = alarm.ID Then
                'If alarm.Serial = left(key, 8) Then
220               ResidentID = GetResidentIDFromSerial(alarm.Serial)
230               CurrentAlarmID = alarm.ID
240               Exit For
250             End If

      '          If Alarm.Serial = left(key, 8) Then
      '            Set d = Devices.Device(Alarm.Serial)
      '            If Not d Is Nothing Then
      '              ResidentID = d.ResidentID
      '              AlarmID = Alarm.id
      '            End If
      '            Exit For
      '          End If
260           Next
270         End If

280       Case InfoPane.LowBatt
290         Set lv = lvLoBatt
300         If Not lv.SelectedItem Is Nothing Then
              'key = lv.SelectedItem.key
310           AlarmID = lvKey2ID(lv.SelectedItem.Key)
320           For j = 1 To LowBatts.alarms.Count
330             Set alarm = LowBatts.alarms(j)
340             If AlarmID = alarm.ID Then
                'If alarm.Serial = left(key, 8) Then
350               ResidentID = GetResidentIDFromSerial(alarm.Serial)
360               CurrentAlarmID = alarm.ID
370               Exit For
380             End If
390           Next
400         End If



410       Case InfoPane.extern
420         Set lv = lvExternal
430         If Not lv.SelectedItem Is Nothing Then
440           Key = lv.SelectedItem.Key
450           For j = 1 To Externs.alarms.Count
460             Set alarm = Externs.alarms(j)
470             If AlarmID = alarm.ID Then
                'If alarm.Serial = left(key, 8) Then
480               ResidentID = GetResidentIDFromSerial(alarm.Serial)
490               CurrentAlarmID = alarm.ID
500               Exit For
510             End If
520           Next
530         End If



540     End Select

        
550     DisplayResidentInfo ResidentID, CurrentAlarmID

ShowResidentInfo_Resume:
560     On Error GoTo 0
570     Exit Sub

ShowResidentInfo_Error:

580     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowResidentInfo." & Erl
590     Resume ShowResidentInfo_Resume


End Sub

Public Sub DisplayResidentInfo(ByVal ID As Long, ByVal AlarmID As Long)
  
  DisplayResOrRoomInfo ID, 0, AlarmID

End Sub

Public Sub DisplayResOrRoomInfo(ByVal ResidentID As Long, ByVal RoomID As Long, ByVal AlarmID As Long)
  Dim rs As Recordset
  picResident.CLS
  picResident.FontSize = 14
  Dim Away As Integer
  Dim text As String

  lstAssurDevs.Clear  ' lstAssurdevs is a misnomer, just devices assigned to room or resident
  cmdChangeVacation.Visible = False
  If ResidentID <> 0 Then
    Set rs = ConnExecute("SELECT * FROM Residents WHERE ResidentID = " & ResidentID)


    If rs.EOF Then
      txtHiddenAlarmID.text = "0"
      txtHiddenResID.text = "0"
      txtHiddenRoomID.text = "0"
      txtInfoFullName.text = "NO DATA"
      txtInfoMessage.text = ""
      txtInfoRoom.text = ""
      txtInfoNotes.text = ""
      txtInfoNotes.Visible = False
      imgResPic.Visible = False
      imgResPic.Picture = LoadPicture()
      CmdEditInfo.Enabled = False
      cmdPrintInfo.Enabled = False
      lblInfoRoom.Visible = False
      cmdChangeVacation.Visible = False
      lstAssurDevs.Visible = False
      lblAssrDayList.Visible = False
      txtAssurDays.text = ""
      txtAssurDays.Visible = False
    Else
      Away = IIf(Val("0" & rs("Away")) = 1, 1, 0)
      
      txtHiddenAlarmID.text = AlarmID
      txtHiddenResID.text = ResidentID
      txtHiddenRoomID.text = "0"
      text = ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
      txtAssurDays.text = GetAssurDaysFromValue(Val("" & rs("AssurDays"))) & "   " & IIf(rs("Away") = 1, "Vac", "")
      txtAssurDays.Visible = True
      lblAssrDayList.Visible = True
      If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
          txtInfoFullName.text = "N/A"
      Else
          txtInfoFullName.text = text
      End If
      
        
      
      txtInfoMessage.text = rs("Phone") & ""
      txtInfoRoom.text = GetResidentRooms(ResidentID)
      txtInfoNotes.text = rs("info") & ""
      GetImageFromDB imgResPic, rs("imagedata")

      txtInfoNotes.Visible = True
      imgResPic.Visible = True
      CmdEditInfo.Enabled = True
      cmdPrintInfo.Enabled = True
      lblInfoRoom.Visible = True
      cmdChangeVacation.Visible = IsAssurActive() Or (Away = 1)

      cmdChangeVacation.Caption = IIf(Away = 1, "Return from Vacation", "Place on Vacation")
      lstAssurDevs.Visible = True  ' assurance devices
      GetAssurDevs ResidentID, lstAssurDevs


    End If
    rs.Close
    Set rs = Nothing
  ElseIf RoomID <> 0 Then
    Set rs = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & RoomID)
        
    If rs.EOF Then
      txtHiddenAlarmID.text = "0"
      txtHiddenResID.text = "0"
      txtHiddenRoomID.text = "0"
      txtInfoFullName.text = "NO DATA"
      txtInfoMessage.text = ""
      txtInfoRoom.text = ""
      txtInfoNotes.text = ""
      txtInfoNotes.Visible = False
      imgResPic.Visible = False
      imgResPic.Picture = LoadPicture()
      CmdEditInfo.Enabled = False
      cmdPrintInfo.Enabled = False
      lblInfoRoom.Visible = False
      cmdChangeVacation.Visible = False
      lstAssurDevs.Visible = False
      lblAssrDayList.Visible = False
      txtAssurDays.text = ""
      txtAssurDays.Visible = False
    Else
    
      'RoomID , room, Building, Locator, Assurdays, Away, Deleted
      
      Away = IIf(Val("0" & rs("Away")) = 1, 1, 0)
      
      txtHiddenAlarmID.text = AlarmID
      txtHiddenResID.text = "0"
      txtHiddenRoomID.text = RoomID
      text = "Room Data"
      txtAssurDays.text = GetAssurDaysFromValue(Val("" & rs("AssurDays"))) & "   " & IIf(rs("Away") = 1, "Vac", "")
      txtAssurDays.Visible = True
      lblAssrDayList.Visible = True
      If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
          txtInfoFullName.text = "N/A"
      Else
          txtInfoFullName.text = text
      End If
      
      txtInfoMessage.text = ""
      txtInfoRoom.text = "" & rs("Room")
      txtInfoNotes.text = ""
      imgResPic.Picture = LoadPicture()

      txtInfoNotes.Visible = True
      imgResPic.Visible = False
      CmdEditInfo.Enabled = False
      cmdPrintInfo.Enabled = True
      lblInfoRoom.Visible = True
      cmdChangeVacation.Visible = IsAssurActive() Or (Away = 1)

      cmdChangeVacation.Caption = IIf(Away = 1, "Return from Vacation", "Place on Vacation")
      lstAssurDevs.Visible = True  ' assurance devices
      GetAssurDevs_Room RoomID, lstAssurDevs


    End If
    
    
    
    rs.Close
    Set rs = Nothing
  Else
      txtHiddenAlarmID.text = "0"
      txtHiddenResID.text = "0"
      txtHiddenRoomID.text = "0"
      txtInfoFullName.text = "NO DATA"
      txtInfoMessage.text = ""
      txtInfoRoom.text = ""
      txtInfoNotes.text = ""
      txtInfoNotes.Visible = False
      imgResPic.Visible = False
      imgResPic.Picture = LoadPicture()
      CmdEditInfo.Enabled = False
      cmdPrintInfo.Enabled = False
      lblInfoRoom.Visible = False
      cmdChangeVacation.Visible = False
      lstAssurDevs.Visible = False
      lblAssrDayList.Visible = False
      txtAssurDays.text = ""
      txtAssurDays.Visible = False

    
  End If
End Sub

Sub GetAssurDevs_Room(ByVal RoomID As Long, list As ListBox)

  Dim rs As Recordset
  Dim SQL As String


  list.Clear

  If RoomID <> 0 Then
    SQL = "Select Distinct * From Devices Where Deleted <> 1 and Roomid = " & RoomID

    Set rs = ConnExecute(SQL)
    Do Until rs.EOF
      list.AddItem rs("custom") & ""
      list.ItemData(list.NewIndex) = rs("deviceid")
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
  End If
End Sub

Sub GetAssurDevs(ByVal ResidentID As Long, list As ListBox)
 ' GetAssurdevs is a misnomer, just devices assigned to room or resident
  Dim rs As Recordset
  Dim SQL As String
  Dim RoomID As Long
  
  list.Clear
  RoomID = GetResidentRoomID(ResidentID) ' get room assigned to resident
  If RoomID <> 0 Then ' if there is a room, get all devices assigned to room and/or resident
    SQL = "Select Distinct * From Devices Where Roomid = " & RoomID & " or Residentid = " & ResidentID
  Else  ' no room assigned, just get those devices assigned to resident
    SQL = "Select Distinct * From Devices Where Residentid = " & ResidentID
  End If
  Set rs = ConnExecute(SQL)
  Do Until rs.EOF
    list.AddItem rs("custom") & ""
    list.ItemData(list.NewIndex) = rs("deviceid")
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
End Sub



Sub ProcessExterns()

        Dim li                 As ListItem
        Dim alarm              As cAlarm
        Dim j                  As Integer
        '        Dim Device        As cESDevice

        Dim rs                 As Recordset
        Dim SQL                As String


        Dim Col0               As String
        Dim col1               As String
        Dim col2               As String
        Dim col3               As String
        Dim Col4               As String
        Dim Col5               As String

        Dim SelectedItem       As ListItem
        Dim SelectedSerialnum  As String

10      On Error GoTo ProcessExterns_Error

20      If Externs.alarms.Count > lvExternal.ListItems.Count Then
30        CurrentGrid = InfoPane.extern
40      End If


50      Set SelectedItem = lvExternal.SelectedItem
60      If Not SelectedItem Is Nothing Then
70        SelectedSerialnum = lvExternal.SelectedItem.Key
80      End If


90      LockWindowUpdate lvExternal.hwnd

100     lvExternal.ListItems.Clear

'110     For j = 1 To Externs.alarms.Count

110     For j = Externs.alarms.Count To 1 Step -1


120       Col0 = ""
130       col1 = ""
140       col2 = ""
150       col3 = ""
160       Col4 = ""
170       Col5 = ""

180       Set alarm = Externs.alarms(j)
          'Set Device = Devices.Device(Alarm.Serial)


'190       SQl = " SELECT Devices.Serial, Devices.Model, Residents.NameLast, Residents.NameFirst, Rooms.Room " & _
'              " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
'              " WHERE (Devices.Serial=" & q(Alarm.Serial) & ");"


200       Col0 = alarm.Serial
210       col3 = Format(alarm.DateTime, gTimeFormatString)
220       Col4 = alarm.Announce

'230       Set rs = ConnExecute(SQl)

            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              col1 = ""
            Else
250           col1 = alarm.ResidentText '  ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
            End If
260         col2 = alarm.RoomText ' rs("Room") & ""
'270       End If
'2'80       rs.Close


290       Set li = lvExternal.ListItems.Add(, alarm.Serial & "S" & alarm.Inputnum, Col0)
300       If Len(col1) = 0 Then
310         li.SubItems(1) = " "
320       Else
330         li.SubItems(1) = col1
340       End If
350       If Len(col2) = 0 Then
360         li.SubItems(2) = " "
370       Else
380         li.SubItems(2) = col2
390       End If
400       If Len(col3) = 0 Then
410         li.SubItems(3) = " "
420       Else
430         li.SubItems(3) = col3
440       End If

450       If Len(Col4) = 0 Then
460         li.SubItems(4) = " "
470       Else
480         li.SubItems(4) = Col4
490       End If



500     Next

510     For j = 1 To lvExternal.ListItems.Count
520       If SelectedSerialnum = lvExternal.ListItems(j).Key Then
530         lvExternal.ListItems(j).EnsureVisible
540         lvExternal.ListItems(j).Selected = True
550         Exit For
560       End If
570     Next




        '560     If lvExternal.ListItems.count Then
        '570       lvExternal.ListItems(lvExternal.ListItems.count).Selected = True
        '580       lvExternal.ListItems(lvExternal.ListItems.count).EnsureVisible
        '590     End If
        '
580     LockWindowUpdate 0


ProcessExterns_Resume:
590     On Error GoTo 0
600     Exit Sub

ProcessExterns_Error:

610     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ProcessExterns." & Erl
620     Resume ProcessExterns_Resume


End Sub

Sub ProcessAlerts()
        Dim li                 As ListItem
        Dim alarm              As cAlarm
        Dim j                  As Integer
        '        Dim Device        As cESDevice

        Dim rs                 As Recordset
        Dim SQL                As String


        Dim Col0               As String
        Dim col1               As String
        Dim col2               As String
        Dim col3               As String
        Dim Col4               As String
        Dim Col5               As String

10      On Error GoTo ProcessAlerts_Error

20      If Alerts.alarms.Count > lvAlerts.ListItems.Count Then
30        CurrentGrid = InfoPane.Alert
40      End If

        Dim SelectedItem       As ListItem
        Dim SelectedSerialnum  As String

50      Set SelectedItem = lvAlerts.SelectedItem
60      If Not SelectedItem Is Nothing Then
70        SelectedSerialnum = lvAlerts.SelectedItem.Key
80      End If


90      LockWindowUpdate lvAlerts.hwnd

100     lvAlerts.ListItems.Clear



'110     For j = 1 To Alerts.alarms.Count

110     For j = Alerts.alarms.Count To 1 Step -1

120       Col0 = ""
130       col1 = ""
140       col2 = ""
150       col3 = ""
160       Col4 = ""
170       Col5 = ""

180       Set alarm = Alerts.alarms(j)
          'Set Device = Devices.Device(Alarm.Serial)


'190       SQl = " SELECT Devices.Serial,Devices.Model, Residents.NameLast, Residents.NameFirst, Rooms.Room " & _
'              " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
'              " WHERE  Devices.serial =" & q(Alarm.Serial)
'
'200       Set rs = ConnExecute(SQl)

210       col3 = alarm.locationtext
220       Col5 = alarm.Announce
230       Col0 = alarm.Serial

240       'If Not rs.EOF Then
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              col1 = ""
            Else
250           col1 = alarm.ResidentText ' ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
            End If
260         col2 = alarm.RoomText ' rs("Room") & ""
270       'End If
280       'rs.Close


290       Set li = lvAlerts.ListItems.Add(, alarm.Serial & "S" & alarm.Inputnum, Col0)
300       If Len(col1) = 0 Then
310         li.SubItems(1) = " "
320       Else
330         li.SubItems(1) = col1
340       End If
350       If Len(col2) = 0 Then
360         li.SubItems(2) = " "
370       Else
380         li.SubItems(2) = col2
390       End If
400       If Len(col3) = 0 Then
410         li.SubItems(3) = " "
420       Else
430         li.SubItems(3) = col3
440       End If

450       li.SubItems(4) = Format(alarm.DateTime, gTimeFormatString)
460       If Len(Col5) = 0 Then
470         Col5 = " "
480       End If
490       li.SubItems(5) = Col5  ' IIf(Alarm.SilenceTime <> 0, Format(Alarm.SilenceTime, gTimeFormatString), " ")

500       li.SubItems(6) = IIf(alarm.ACKTime <> 0, Format(alarm.ACKTime, gTimeFormatString), " ")
505       li.SubItems(7) = alarm.Responder
510     Next

520     For j = 1 To lvAlerts.ListItems.Count
530       If SelectedSerialnum = lvAlerts.ListItems(j).Key Then
540         lvAlerts.ListItems(j).EnsureVisible
550         lvAlerts.ListItems(j).Selected = True
560         Exit For
570       End If
580     Next


ProcessAlerts_Resume:
590     On Error GoTo 0
600     LockWindowUpdate 0
610     Exit Sub

ProcessAlerts_Error:

620     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ProcessAlerts." & Erl
630     Resume ProcessAlerts_Resume


End Sub

Function lvKey2ID(ByVal Key As String) As Long
    Dim KeyParts() As String
    On Error Resume Next
    
    Key = lvEmergency.SelectedItem.Key
    KeyParts = Split(Key, "S", , vbTextCompare)
    lvKey2ID = Val(KeyParts(0))


End Function

'Function GetlvEmergencySelectedID()
'  If Not lvEmergency.SelectedItem Is Nothing Then
'    Dim key As String
'    Dim KeyParts() As String
'    key = lvEmergency.SelectedItem.key
'    KeyParts = Split(key, "S", , vbTextCompare)
'    GetlvEmergencySelectedID = Val(KeyParts(0))
'    'GetlvEmergencySelectedKey = lvEmergency.SelectedItem.Key
'  End If
'
'End Function

'Function GetlvAlertsSelectedID()
'  If Not lvEmergency.SelectedItem Is Nothing Then
'    Dim key As String
'    Dim KeyParts() As String
'    key = lvAlerts.SelectedItem.key
'    KeyParts = Split(key, "S", , vbTextCompare)
'    GetlvAlertsSelectedID = Val(KeyParts(0))
'    'GetlvEmergencySelectedKey = lvEmergency.SelectedItem.Key
'  End If
'
'End Function


Sub ProcessAlarms()
        ' 32k alarms max
        Dim li                 As ListItem
        Dim alarm              As cAlarm
        Dim j                  As Integer
        '  Dim Device        As cESDevice
        Dim rs                 As Recordset
        Dim SQL                As String


        Dim Col0               As String
        Dim col1               As String
        Dim col2               As String
        Dim col3               As String
        Dim Col4               As String
        Dim Col5               As String
        Dim Col6               As String

10      On Error GoTo ProcessAlarms_Error

        'On Error GoTo 0
20      If alarms.alarms.Count <> lvEmergency.ListItems.Count Then
30        CurrentGrid = InfoPane.emergency
40        If lvEmergency.Visible Then
50          SetFocusTo lvEmergency
60        End If
70      End If


        Dim SelectedItem       As ListItem
        Dim SelectedSerialnum  As String

80      Set SelectedItem = lvEmergency.SelectedItem
90      If Not SelectedItem Is Nothing Then
100       SelectedSerialnum = lvEmergency.SelectedItem.Key
110     End If


120     LockWindowUpdate lvEmergency.hwnd

130     lvEmergency.ListItems.Clear

        ' Debug.Print "ProcessAlarms >> alarms.alarms.Count "; alarms.alarms.Count
        
        ' May need to percolate Assist Calls to top ???
        
        

'140     For j = 1 To alarms.alarms.Count
140     For j = alarms.alarms.Count To 1 Step -1 ' newest first

150       Col0 = ""
160       col1 = ""
170       col2 = ""
180       col3 = ""
190       Col4 = ""
200       Col5 = ""
205       Col6 = ""
210       Set alarm = alarms.alarms(j)

220       'SQl = " SELECT Devices.Serial,Devices.Model, Residents.NameLast, Residents.NameFirst, Rooms.Room " & _
          '    " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
          '    " WHERE  Devices.serial =" & q(Alarm.Serial)
230       Col0 = alarm.Serial

240       'Set rs = ConnExecute(SQl)

250       'If Not rs.EOF Then  ' fetch resident and room for device
            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              col1 = ""
            Else
260           col1 = alarm.ResidentText ' ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")

            End If

270         col2 = alarm.RoomText
280       'End If
290       'rs.Close

300       Set li = lvEmergency.ListItems.Add(, CStr(alarm.ID) & "s", Col0)
310       If Len(col1) = 0 Then
320         li.SubItems(1) = " "
330       Else
340         li.SubItems(1) = col1
350       End If
360       If Len(col2) = 0 Then
370         li.SubItems(2) = " "
380       Else
390         li.SubItems(2) = col2
400       End If
410       If Len(Col5) = 0 Then
420         Col5 = " "
430       End If
          
440       li.SubItems(3) = alarm.locationtext
          Call dbgPackets("MainForm ProcessAlarms Alarm.locationtext " & alarm.locationtext)
450       li.SubItems(4) = Format(alarm.DateTime, gTimeFormatString)
460       li.SubItems(5) = alarm.Announce
470       li.SubItems(6) = IIf(alarm.ACKTime <> 0, Format(alarm.ACKTime, gTimeFormatString), " ")
475       li.SubItems(7) = alarm.Responder

          'End If
480     Next

490     For j = 1 To lvEmergency.ListItems.Count
500       If SelectedSerialnum = lvEmergency.ListItems(j).Key Then
510         lvEmergency.ListItems(j).EnsureVisible
520         lvEmergency.ListItems(j).Selected = True
530         Exit For
540       End If
550     Next

        '      If lvEmergency.ListItems.count Then
        '        lvEmergency.ListItems(lvEmergency.ListItems.count).Selected = True  ' select the last one on the list
        '        lvEmergency.ListItems(lvEmergency.ListItems.count).EnsureVisible    ' make sure it can be seen onscreen
        '      End If
ProcessAlarms_Resume:
560     On Error GoTo 0
570     LockWindowUpdate 0
580     Exit Sub

ProcessAlarms_Error:

590     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ProcessAlarms." & Erl
600     Resume ProcessAlarms_Resume


End Sub

Sub AckLowBatt()
'LowBatts.BeepTimer = 0
  CurrentGrid = InfoPane.LowBatt
  ShowResidentInfo 0

End Sub

Sub AckTrouble()
'Troubles.BeepTimer = 0
  CurrentGrid = InfoPane.Checkin
  ShowResidentInfo 0
End Sub
Public Sub ShowCommError(ByVal Status As Boolean)

10      On Error Resume Next

        Dim ErrorMessage       As String

20      If MASTER Then

30        If USE6080 Then
40          ErrorMessage = "Communication with ACG Lost"
50        Else
60          ErrorMessage = "Communication with NC Lost"
70        End If



80        If Status Then

90          If Now > RemoteStartUpTimer Then

100           If picCommError.Visible = False Then

110             picCommError.CLS
120             picCommError.Visible = True
130             picCommError.Enabled = False
140             picCommError.BackColor = vbRed
150             picCommError.ForeColor = vbYellow
160             picCommError.CurrentY = (picCommError.Height - picCommError.TextHeight("A") * 4) / 2

170             picCommError.Width = Max(4995, picCommError.TextWidth(Configuration.AdminContact) + 120)
180             picCommError.left = (Me.Width / 2 - picCommError.Width / 2)
190             picCommError.CurrentX = (picCommError.Width - picCommError.TextWidth(ErrorMessage)) / 2
200             picCommError.Print ErrorMessage
210             picCommError.Print ""
220             picCommError.CurrentX = (picCommError.Width - picCommError.TextWidth(Configuration.AdminContact)) / 2
230             picCommError.Print Configuration.AdminContact

240           End If

250           picCommError.ZOrder 0
260           If Len(Configuration.TroubleFile) Then
270             PlayASound Configuration.TroubleFile, Win32.SND_ASYNC
280           End If

290         End If
300       Else
310         If picCommError.Visible Then
320           picCommError.CLS
330           picCommError.Visible = False
340           picCommError.Enabled = False
350           picCommError.BackColor = vbWhite
360           picCommError.ForeColor = vbBlack
370           picCommError.ZOrder 1
380         End If
390       End If



400     End If

End Sub


Public Sub ShowDisconnect(ByVal Status As Boolean)

  Dim Elapsed            As Double

  On Error Resume Next

  

  If Not MASTER Then

    If Status Then

      If Now > RemoteStartUpTimer Then
        picCommError.CLS
        picCommError.Visible = True
        picCommError.Enabled = False
        picCommError.BackColor = vbRed
        picCommError.ForeColor = vbYellow
        picCommError.CurrentY = (picCommError.Height - picCommError.TextHeight("A") * 4) / 2
        picCommError.CurrentX = (picCommError.Width - picCommError.TextWidth("Communication with Host Lost")) / 2
        picCommError.Print "Communication with Host Lost"
        picCommError.Print ""
        picCommError.CurrentX = (picCommError.Width - picCommError.TextWidth(Configuration.AdminContact)) / 2
        picCommError.Print Configuration.AdminContact



        picCommError.ZOrder 0
        If Len(Configuration.TroubleFile) Then
          PlayASound Configuration.TroubleFile, Win32.SND_ASYNC
        End If

      End If
    Else
      If picCommError.Visible Then
        picCommError.CLS
        picCommError.Visible = False
        picCommError.Enabled = False
        picCommError.BackColor = vbWhite
        picCommError.ForeColor = vbBlack
        picCommError.ZOrder 1
      End If
    End If

  End If

End Sub

Private Sub imgUptime_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If imgUptime.ToolTipText <> DateDiff("n", starttime, Now) & " Minutes Uptime" Then
    imgUptime.ToolTipText = DateDiff("n", starttime, Now) & " Minutes Uptime"
  End If

End Sub

Private Sub lblAlarms_Click()
  SetFocusTo lvEmergency
End Sub

Private Sub lblAlerts_Click()
  CurrentGrid = InfoPane.Alert
End Sub

Private Sub lblLowBattery_Click()
  CurrentGrid = InfoPane.LowBatt
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If lblTime.ToolTipText <> Format(Now, "dddd, mmmm dd, yyyy") Then
    lblTime.ToolTipText = Format(Now, "dddd, mmmm dd, yyyy")
  End If
End Sub

Private Sub lblTrouble_Click()
  CurrentGrid = InfoPane.Checkin
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lstAssurDevs_DblClick()
  ' misnomer , includes all transmitters for room and or residnet
  If lstAssurDevs.ListIndex > -1 Then
    If gUser.LEvel > LEVEL_USER Then
      EditTransmitter lstAssurDevs.ItemData(lstAssurDevs.ListIndex)
    End If
  End If
End Sub

Private Sub lvAlerts_GotFocus()
'  ResetActivityTime
  CurrentGrid = InfoPane.Alert
End Sub

Private Sub lvAlerts_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.Alert
  ShowAlertInfo
  
  'ShowResidentInfo 0

End Sub

Private Sub lvAlerts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim j                  As Long
  Dim AlarmID            As Long
  Dim Device             As cESDevice
  Dim EventType          As Long
  Dim Inputnum           As Long
  Dim Serial             As String
  Dim packet             As cESPacket
  Dim Key                As String
  Dim Disposition        As String
  Dim assistalarm        As cAlarm



  If Not MASTER Then
    '  Exit Sub
  End If


  If Button = vbRightButton Then

    Dim alarm            As cAlarm
    If Not lvAlerts.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvAlerts.SelectedItem.Key)  '   GetlvEmergencySelectedID()
      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          If alarm.Alarmtype <> EVT_ASSISTANCE Then  ' can't staff assist a staff assist
            For Each assistalarm In alarms.alarms
              If assistalarm.PriorID = AlarmID Then
                messagebox Me, "Already Has an Active Staff Assist", "Sentry Freedom II", vbOKOnly Or vbInformation
                Exit Sub
              End If
            Next
            If vbYes = messagebox(Me, "Call for Staff Assist?", "Sentry Freedom II", vbYesNo Or vbQuestion) Then
              ' create assist call
              If MASTER Then
                Set Device = Devices.Device(alarm.Serial)
                Call PostEvent(Device, Nothing, alarm, EVT_ASSISTANCE, alarm.Inputnum, gUser.Username)
              Else
                'Call ClientRequestAssist("alarms", alarm.Serial, alarm.inputnum, alarm.AlarmID)
                DoClientRequestAssist
                ' create remote call for assistance 'Call PostEvent(Device, Nothing, alarm, EVT_ASSISTANCE, alarm.inputnum, gUser.Username)
              End If
              Exit For
            End If             ' vbYes = MsgBox
          End If               ' alarm.Alarmtype <
        End If                 ' alarm.Serial = Serial
      Next                     ' next j
    End If
  End If



End Sub

Private Sub lvAssur_Click()
  ResetActivityTime
  ShowAssurInfo

End Sub

Private Sub lvAssur_ItemClick(ByVal Item As MSComctlLib.ListItem)
  ShowAssurInfo
End Sub

Private Sub ShowAssurInfo()
  Dim lv          As ListView
  Dim Key         As String
  Dim j           As Integer
  Dim alarm       As cAlarm
  Dim d           As cESDevice
  Dim ResidentID  As Long
  Dim AlarmID     As Long

10          On Error GoTo ShowAssurInfo_Error

20          Set lv = lvAssur

30          Key = GetlvAssurSelectedKey()
40          For j = 1 To Assurs.alarms.Count
50            Set alarm = Assurs.alarms(j)
60            If alarm.Serial = left(Key, 8) Then
70              Set d = Devices.Device(alarm.Serial)
80              If Not d Is Nothing Then
90                ResidentID = d.ResidentID
100               AlarmID = alarm.ID
110             End If
120             Exit For
130           End If
140         Next
150         DisplayResidentInfo ResidentID, AlarmID

ShowAssurInfo_Resume:
160         On Error GoTo 0
170         Exit Sub

ShowAssurInfo_Error:

180         LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowAssurInfo." & Erl
190         Resume ShowAssurInfo_Resume


End Sub

Private Sub lvCheckIn_GotFocus()
  CurrentGrid = InfoPane.Checkin
End Sub

Private Sub lvCheckIn_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.Checkin
  UpdateLayout
  'ShowResidentInfo 0
  ShowCheckinInfo

End Sub


Private Sub lvEmergency_Click()
'  ResetActivityTime
  ' When clicking in the lists, or the up-down/print list buttons,
  ' reset the refresh interval to a whole interval

  ShowEmergencyInfo
End Sub

Private Sub lvEmergency_DblClick()
' get alarm
  Dim lv As ListView

  Dim j As Integer
  Dim alarm As cAlarm
  Dim AlarmID As Long
  Dim WaypointID As Long

  Dim Serial  As String

  Dim w As cWayPoint
  Dim s As String

  Set lv = lvEmergency

  If lvEmergency.SelectedItem Is Nothing Then
    Exit Sub
  End If

  'Serial = GetlvEmergencySelectedKey()
  AlarmID = lvKey2ID(lvEmergency.SelectedItem.Key)  '   GetlvEmergencySelectedID()
  
  For j = 1 To alarms.alarms.Count
    Set alarm = alarms.alarms(j)
    'If 0 = StrComp(alarm.Serial, left(Serial, 8), vbTextCompare) Then
      If AlarmID = alarm.AlarmID Then
      dbgloc "" & vbCrLf
      s = left("DEVICE" & "         ", 10) & " "
      dbgloc s & alarm.Repeater1 & " " & Format(alarm.Signal1, "0") & " " & alarm.Repeater2 & " " & Format(alarm.Signal2, "0") & " " & alarm.Repeater3 & " " & Format(alarm.Signal3, "0") & vbCrLf
      
      For WaypointID = 1 To Waypoints.Count
        Set w = Waypoints.waypoint(WaypointID)
        If w.Description = alarm.locationtext Then
        s = left(w.Description & "         ", 10) & "*"
        dbgloc s & w.Repeater1 & " " & Format(w.Signal1, "0") & " " & w.Repeater2 & " " & Format(w.Signal2, "0") & " " & w.Repeater3 & " " & Format(w.Signal3, "0") & vbCrLf
        Exit For
        End If
      Next
      
      
      For WaypointID = 1 To Waypoints.Count
        Set w = Waypoints.waypoint(WaypointID)
        s = left(w.Description & "         ", 10) & " "
        dbgloc s & w.Repeater1 & " " & Format(w.Signal1, "0") & " " & w.Repeater2 & " " & Format(w.Signal2, "0") & " " & w.Repeater3 & " " & Format(w.Signal3, "0") & vbCrLf
      Next
      dbgloc "" & vbCrLf
      Exit For
    End If
  Next





End Sub

Private Sub lvEmergency_GotFocus()
  CurrentGrid = InfoPane.emergency
End Sub

Private Sub lvEmergency_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.emergency
  ShowEmergencyInfo

End Sub

Private Sub lvEmergency_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim j                  As Long
  Dim AlarmID            As Long
  Dim Device             As cESDevice
  Dim EventType          As Long
  Dim Inputnum           As Long
  Dim Serial             As String
  Dim packet             As cESPacket
  Dim Key                As String
  Dim Disposition        As String
  Dim assistalarm        As cAlarm



  If Not MASTER Then
    '  Exit Sub
  End If


  If Button = vbRightButton Then

    Dim alarm            As cAlarm
    If Not lvEmergency.SelectedItem Is Nothing Then
      AlarmID = lvKey2ID(lvEmergency.SelectedItem.Key)  '   GetlvEmergencySelectedID()
      For j = 1 To alarms.alarms.Count
        Set alarm = alarms.alarms(j)
        'If alarm.Serial = Serial And alarm.inputnum = inputnum Then
        If alarm.ID = AlarmID Then
          If alarm.Alarmtype <> EVT_ASSISTANCE Then  ' can't staff assist a staff assist
            For Each assistalarm In alarms.alarms
              If assistalarm.PriorID = AlarmID Then
                messagebox Me, "Already Has an Active Staff Assist", "Sentry Freedom II", vbOKOnly Or vbInformation
                Exit Sub
              End If
            Next
            If vbYes = messagebox(Me, "Call for Staff Assist?", "Sentry Freedom II", vbYesNo Or vbQuestion) Then
              ' create assist call
              If MASTER Then
                Set Device = Devices.Device(alarm.Serial)
                Call PostEvent(Device, Nothing, alarm, EVT_ASSISTANCE, alarm.Inputnum, gUser.Username)
              Else
                'Call ClientRequestAssist("alarms", alarm.Serial, alarm.inputnum, alarm.AlarmID)
                DoClientRequestAssist
                ' create remote call for assistance 'Call PostEvent(Device, Nothing, alarm, EVT_ASSISTANCE, alarm.inputnum, gUser.Username)
              End If
              Exit For
            End If             ' vbYes = MsgBox
          End If               ' alarm.Alarmtype <
        End If                 ' alarm.Serial = Serial
      Next                     ' next j
    End If
  End If
End Sub

Private Sub lvExternal_GotFocus()
  CurrentGrid = InfoPane.extern
End Sub

Private Sub lvExternal_ItemClick(ByVal Item As MSComctlLib.ListItem)
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.extern
  ShowResidentInfo 0
End Sub

Private Sub lvLoBatt_GotFocus()
  CurrentGrid = InfoPane.LowBatt
End Sub

Private Sub lvLoBatt_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
  ResetRemoteRefreshCounter
  CurrentGrid = InfoPane.LowBatt
  ShowResidentInfo 0

End Sub

Private Sub lvtx_ItemClick(ByVal Item As MSComctlLib.ListItem)
'  ResetActivityTime
  ShowLVTXData Item.Key
End Sub
Sub ShowLVTXData(ByVal Key As String)

      ' Get Resident Info
      ' or get room info

        Dim ResidentID As Long
        Dim Ptr        As Long

10      On Error GoTo ShowLVTXData_Error

20      ResidentID = Val(Key)
30      If ResidentID <> 0 Then
40        DisplayResidentInfo ResidentID, 0
50      Else
60        Ptr = InStr(Key, "@")
70        If Ptr > 0 Then
80          Key = MID(Key, Ptr + 1)
90          ResidentID = GetResidentIDFromTransmitter(Val(Key))
100         If ResidentID <> 0 Then
110           DisplayResidentInfo ResidentID, 0
120         End If
130       Else
140         Ptr = InStr(Key, "R")
150         If Ptr > 0 Then
160           Key = MID(Key, Ptr + 1)
              'DisplayResidentInfo 0, 0
170           DisplayResOrRoomInfo 0, Val(Key), 0
180         End If
190       End If
200     End If



ShowLVTXData_Resume:
210     On Error GoTo 0
220     Exit Sub

ShowLVTXData_Error:

230     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ShowLVTXData." & Erl
240     Resume ShowLVTXData_Resume


End Sub

Function GetResidentIDFromTransmitter(ByVal TxID As Long) As Long
  Dim rs As Recordset
  Set rs = ConnExecute("SELECT ResidentID FROM devices WHERE DeviceID = " & TxID)
  If Not rs.EOF Then
    GetResidentIDFromTransmitter = rs(0)
  End If
  rs.Close


End Function

Private Sub lvtx_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'MsgBox "lvtx_MouseUp"
End Sub

Private Sub tabList_BeforeClick(Cancel As Integer)
  'Static busy
'  ResetActivityTime
  If Not lvtxbusy Then
    
    If Not tabList.SelectedItem Is Nothing Then
      Select Case LCase(tabList.SelectedItem.Key)
        Case "tx"
          If lvtxCols_Serial Is Nothing Then
            Set lvtxCols_Serial = New cGridColumns
          End If
          lvtxCols_Serial.Size(1) = lvtx.ColumnHeaders(1).Width
          lvtxCols_Serial.Size(2) = lvtx.ColumnHeaders(2).Width
        Case "room"
          If lvtxCols_Room Is Nothing Then
            Set lvtxCols_Room = New cGridColumns
          End If
          lvtxCols_Room.Size(1) = lvtx.ColumnHeaders(1).Width
          lvtxCols_Room.Size(2) = lvtx.ColumnHeaders(2).Width
        Case Else
          If lvtxCols_Res Is Nothing Then
            Set lvtxCols_Res = New cGridColumns
          End If
          lvtxCols_Res.Size(1) = lvtx.ColumnHeaders(1).Width
          lvtxCols_Res.Size(2) = lvtx.ColumnHeaders(2).Width
          lvtxCols_Res.Size(3) = lvtx.ColumnHeaders(3).Width
      End Select
    End If
    
  Else
    Cancel = True
    Beep
  End If
End Sub

Private Sub tabList_Click()
  If Not lvtxbusy Then
    SetListTabs
  End If
End Sub
Public Sub SetListTabs()
  Dim i As Integer

  
  lvtxbusy = True
  
  If Not tabList.SelectedItem Is Nothing Then
    Select Case LCase(tabList.SelectedItem.Key)
      Case "tx"
        i = SORT_SERIAL
      Case "room"
        i = SORT_ROOM
      Case Else
        i = 0
    End Select
    
    
    Fill_lvtx i
  End If
  lvtxbusy = False

End Sub




Sub UpdateScreenElements()
  Dim t As Object

  On Error GoTo UpdateScreenElements_Error
  Dim MasterSlug As String
  If USE6080 Then
    MasterSlug = "ACG "
  Else
    MasterSlug = "NC "
  End If
  frmMain.Caption = IIf(MASTER, MasterSlug, "REMOTE CONSOLE ") & PRODUCT_NAME & " (Ver. 1" & Format(App.Revision, "000") & ")  Logged In: " & gUser.Username & "  Level: " & gUser.LevelString
  
  #If brookdale Then
    frmMain.Icon = frmBrookdale.Icon ' LoadResPicture("BROOKDALE", vbResIcon)
  #ElseIf esco Then
    frmMain.Icon = frmEsco.Icon 'LoadResPicture("ESCO", vbResIcon)
  #Else
    'frmMain.Icon = LoadResPicture("SENTRY", vbResIcon)
  #End If
  
  'Heritage MedCall Sentry Freedom I E-Call System (Ver.
  
  
  cmdAssur.Visible = ((Configuration.AssurStart <> Configuration.AssurEnd) Or (Configuration.AssurStart2 <> Configuration.AssurEnd2))
  
  cmdChangeVacation.Visible = IsAssurActive() And (txtInfoFullName.text <> "NO DATA")
  
  cmdAssur.Visible = cmdAssur.Visible And (gAssurDisableScreenOutput = 0)
  
  
  
  
  ' also do items by access level
  Select Case gUser.LEvel
    Case LEVEL_FACTORY
      cmdResidents.Visible = True
      cmdTransmitters.Visible = True
      cmdRooms.Visible = True
      cmdSetup.Visible = True
      cmdOutputServers.Visible = (True And MASTER)
      CmdEditInfo.Visible = True

      cmdOutputDevices.Visible = True
      cmdOutputs.Visible = True
      If tabList.Tabs.Count = 2 Then
        Set t = tabList.Tabs.Add(3, "tx", "Transmitters")
        t.ToolTipText = "Sort By Transmitter"
      End If

    Case LEVEL_ADMIN
      cmdResidents.Visible = True
      cmdTransmitters.Visible = True
      cmdRooms.Visible = True
      cmdSetup.Visible = True
      cmdOutputServers.Visible = (True And MASTER)
      cmdOutputDevices.Visible = True
      cmdOutputs.Visible = True
      CmdEditInfo.Visible = True
      If tabList.Tabs.Count = 2 Then
        Set t = tabList.Tabs.Add(3, "tx", "Transmitters")
        t.ToolTipText = "Sort By Transmitter"
      End If

    Case LEVEL_SUPERVISOR
      cmdResidents.Visible = True
      cmdRooms.Visible = True
      cmdTransmitters.Visible = True
      cmdSetup.Visible = False
      cmdOutputServers.Visible = (False And MASTER)
      cmdOutputDevices.Visible = False
      cmdOutputs.Visible = False
      CmdEditInfo.Visible = True
      If tabList.Tabs.Count = 2 Then
        Set t = tabList.Tabs.Add(3, "tx", "Transmitters")
        t.ToolTipText = "Sort By Transmitter"
      End If

    Case Else
      cmdResidents.Visible = False
      cmdTransmitters.Visible = False
      cmdRooms.Visible = False
      cmdSetup.Visible = False
      cmdOutputServers.Visible = (False And MASTER)
      cmdOutputDevices.Visible = False
      cmdOutputs.Visible = False
      CmdEditInfo.Visible = False
      If tabList.Tabs.Count = 3 Then
        If tabList.SelectedItem.index = 3 Then
          Set tabList.SelectedItem = tabList.Tabs(1)
        End If
        tabList.Tabs.Remove 3
      End If

  End Select



UpdateScreenElements_Resume:
  On Error GoTo 0
  Exit Sub

UpdateScreenElements_Error:

  'LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.UpdateScreenElements." & Erl
  Resume UpdateScreenElements_Resume


End Sub

Sub UpdateUptimeStats()
        Static toggle          As Boolean
        Dim beeptimers         As Long
        Dim LocalLastSilenced


10      On Error GoTo UpdateUptimeStats_Error
        'On Error GoTo 0

20      toggle = Not toggle


30      lblTime.Caption = Format(Now, "ddd, mmm. dd, yy   " & gTimeFormatString)  ' h:nn AM/PM")


40      If alarms.Count = 0 Then
50        alarms.BeepTimer = 0
60      End If

70      If Alerts.Count = 0 Then
80        Alerts.BeepTimer = 0
90      End If

100     If Troubles.Count = 0 Then
110       Troubles.BeepTimer = 0
120     End If

130     If LowBatts.Count = 0 Then
140       LowBatts.BeepTimer = 0
150     End If

160     If Assurs.Count = 0 Then
170       Assurs.BeepTimer = 0
180     End If

190     If Externs.Count = 0 Then
200       Externs.BeepTimer = 0
210     End If


        'Global HostTime As Double ' the time on the host computer
        'Global HostOffset As Long ' seconds
220     'Debug.Print ""
230     'Debug.Print " <<<<<<<<<<<< update uptime >>>>>>>>>>>>"

240     If toggle Then

beeptimerscode:


250       If CBool(Configuration.BeepControl) Then  '  Or True Then  ' use local settings
260         If MASTER Then
270           HostTime = Now
280         End If


            ' sanity check
290         If HostTime < 42370 Or HostTime > 51136 Then  ' bad/no host time returned pre 1/1/2006 or after 12/31/2040
300           HostTime = Now         ' use local time
310         End If
320         HostOffset = DateDiff("s", CDate(HostTime), CDate(Now))


            ' **************** repeat for Alerts etc ***************

            ' Still in use local settings

330         If alarms.Count = 0 Then
340           beeptimers = 0
350           alarms.BeepTimer = 0   ' need to add each type of alarm
360         Else                     ' alarms.Count > 0

370           If alarms.LocalSilenceTime = 0 Then
380             beeptimers = -1
390             alarms.BeepTimer = -1
400             'Debug.Print "alarms.LocalSilenceTime = 0", "alarms.LocalAlarmTime = " & DateDiff("s", CDate(alarms.LocalAlarmTime), CDate(HostTime)), Configuration.AlarmBeep
                ' might need to move to end of routine to avoid reentry
410             If Configuration.AlarmBeep > 0 Then
420               If alarms.LocalAlarmTime = 0 Then
430                 alarms.LocalAlarmTime = Now  ' fix for deadlock
440               End If
450               'Debug.Print " DateDiff('s', CDate(alarms.LocalAlarmTime), CDate(HostTime)) " & DateDiff("s", CDate(alarms.LocalAlarmTime), CDate(HostTime))

460               If DateDiff("s", CDate(alarms.LocalAlarmTime), CDate(HostTime)) > Configuration.AlarmBeep Then
                    ' Still in use local settings
470                 If MASTER Then
480                   Call Silence("MASTER", "Alarms")
490                 Else

500                   'Debug.Print "ClientSilenceAlarms"
510                   QueEvent "ClientSilenceAlarms", ConsoleID, "Alarms"
520                 End If
530               End If
540             ElseIf Configuration.AlarmBeep < 0 Then  ' beep forever
                  ' no need to auto-silence

550             End If
560           Else                   ' if alarms.LocalLastSilenced = 0
                ' Still in use local settings
'570             Debug.Print "alarms.LocalSilenceTime non-zero"
'580             Debug.Print "datediff " & DateDiff("s", CDate(alarms.LocalSilenceTime), CDate(HostTime))
'590             Debug.Print "rebeep " & Configuration.AlarmReBeep



600             If DateDiff("s", CDate(alarms.LocalSilenceTime), CDate(HostTime)) > Configuration.AlarmReBeep And Configuration.AlarmReBeep > 0 Then

610               alarms.BeepTimer = -1
620               beeptimers = -1
630               If MASTER Then
640                 Call UnSilence("MASTER", "", "Alarms")
650               Else
                    'ClientUnSilenceAlarms ConsoleID, "", "Alarms"
660                 QueEvent "ClientUnSilenceAlarms", ConsoleID, "Alarms"

                    'Debug.Print "ClientUnSilenceAlarms "
670               End If
680             Else                 ' still in silence mode
690               alarms.BeepTimer = 0
700               beeptimers = 0
710             End If
720           End If
730         End If                   ' alarms.Count

            ' **************** ALERTS ***************

740         If Alerts.Count = 0 Then
750           beeptimers = 0
760           Alerts.BeepTimer = 0   ' need to add each type of alarm
770         Else                     ' alerts.Count > 0

780           If Alerts.LocalSilenceTime = 0 Then
790             beeptimers = -1
800             Alerts.BeepTimer = -1
810             Debug.Print "alerts.LocalSilenceTime = 0", "alerts.LocalAlarmTime = " & CDate(Alerts.LocalAlarmTime)
                ' might need to move to end of routine to avoid reentry
820             If Configuration.AlarmBeep > 0 Then
830               If Alerts.LocalAlarmTime = 0 Then
840                 Alerts.LocalAlarmTime = Now  ' fix for deadlock
850               End If
860               Debug.Print " DateDiff('s', CDate(alerts.LocalAlarmTime), CDate(HostTime)) " & DateDiff("s", CDate(Alerts.LocalAlarmTime), CDate(HostTime))

870               If DateDiff("s", CDate(Alerts.LocalAlarmTime), CDate(HostTime)) > Configuration.AlertBeep Then

880                 If MASTER Then
890                   Call Silence("MASTER", "Alerts")
900                 Else
910                   QueEvent "ClientSilenceAlarms", ConsoleID, "Alerts"
                      'ClientSilenceAlarms ConsoleID, "", "Alerts"
920                 End If
930               End If
940             End If
950           Else                   'alerts.LocalLastSilenced > 0
960             Debug.Print "alerts.LocalSilenceTime non-zero"
970             Debug.Print "datediff " & DateDiff("s", CDate(Alerts.LocalSilenceTime), CDate(HostTime))
980             Debug.Print "rebeep " & Configuration.AlertReBeep
990             If DateDiff("s", CDate(Alerts.LocalSilenceTime), CDate(HostTime)) > Configuration.AlertReBeep And Configuration.AlertReBeep > 0 Then
1000              Alerts.BeepTimer = -1
1010              beeptimers = -1
1020              If MASTER Then
1030                Call UnSilence("MASTER", "", "Alerts")
1040              Else
                    'ClientUnSilenceAlarms ConsoleID, "", "Alerts"
1050                QueEvent "ClientUnSilenceAlarms", ConsoleID, "Alerts"
                    'Debug.Print "ClientUnSilenceAlarms "
1060              End If
1070            Else                 ' still in silence mode
1080              Alerts.BeepTimer = 0
1090              beeptimers = 0
1100            End If
1110          End If
1120        End If                   ' alerts.Count

            ' TROUBLES

1130        If Troubles.Count = 0 Then
1140          beeptimers = 0
1150          Troubles.BeepTimer = 0  ' need to add each type of alarm
1160        Else                     ' troubles.Count > 0

1170          If Troubles.LocalSilenceTime = 0 Then
1180            beeptimers = -1
1190            Troubles.BeepTimer = -1
                ' might need to move to end of routine to avoid reentry

1200            If Configuration.AlarmBeep > 0 Then
1210              If Troubles.LocalAlarmTime = 0 Then
1220                Troubles.LocalAlarmTime = Now  ' fix for deadlock
1230              End If
1240              Debug.Print " DateDiff('s', CDate(troubles.LocalAlarmTime), CDate(HostTime)) " & DateDiff("s", CDate(Troubles.LocalAlarmTime), CDate(HostTime))

1250              If DateDiff("s", CDate(Troubles.LocalAlarmTime), CDate(HostTime)) > Configuration.TroubleBeep Then

1260                If MASTER Then
1270                  Call Silence("MASTER", "Troubles")
1280                Else
                      'Debug.Print "ClientSilenceAlarms", "Troubles"
                      'ClientSilenceAlarms ConsoleID, "", "Troubles"
1290                  QueEvent "ClientSilenceAlarms", ConsoleID, "Troubles"
1300                End If
1310              End If
1320            End If
1330          Else                   'troubles.LocalLastSilenced > 0
1340            If DateDiff("s", CDate(Troubles.LocalSilenceTime), CDate(HostTime)) > Configuration.TroubleReBeep And Configuration.TroubleReBeep > 0 Then
1350              Troubles.BeepTimer = -1
1360              beeptimers = -1
1370              If MASTER Then
1380                Call UnSilence("MASTER", "", "Troubles")
1390              Else
                    'ClientUnSilenceAlarms ConsoleID, "", "Troubles"
1400                QueEvent "ClientUnSilenceAlarms", ConsoleID, "Troubles"
                    'Debug.Print "ClientUnSilenceAlarms "
1410              End If
1420            Else                 ' still in silence mode
1430              Troubles.BeepTimer = 0
1440              beeptimers = 0
1450            End If
1460          End If
1470        End If                   ' troubles.Count


            ' **************** LOWBATTS ******************


1480        If LowBatts.Count = 0 Then
1490          beeptimers = 0
1500          LowBatts.BeepTimer = 0  ' need to add each type of alarm
1510        Else                     ' lowbatts.Count > 0

1520          If LowBatts.LocalSilenceTime = 0 Then
1530            beeptimers = -1
1540            LowBatts.BeepTimer = -1
1550            'Debug.Print "lowbatts.LocalSilenceTime = 0", "lowbatts.LocalAlarmTime = " & CDate(LowBatts.LocalAlarmTime)
                ' might need to move to end of routine to avoid reentry
1560            If Configuration.AlarmBeep > 0 Then
1570              If LowBatts.LocalAlarmTime = 0 Then
1580                LowBatts.LocalAlarmTime = Now  ' fix for deadlock
1590              End If
1600              'Debug.Print " DateDiff('s', CDate(lowbatts.LocalAlarmTime), CDate(HostTime)) " & DateDiff("s", CDate(LowBatts.LocalAlarmTime), CDate(HostTime))

1610              If DateDiff("s", CDate(LowBatts.LocalAlarmTime), CDate(HostTime)) > Configuration.LowBattBeep Then

1620                If MASTER Then
1630                  Call Silence("MASTER", "LowBatts")
1640                Else
1650                  Debug.Print "ClientSilenceAlarms,  LowBatts"
                      'ClientSilenceAlarms ConsoleID, "", "LowBatts"
1660                  QueEvent "ClientSilenceAlarms", ConsoleID, "LowBatts"
1670                End If
1680              End If
1690            End If
1700          Else                   'lowbatts.LocalLastSilenced > 0
'1710            Debug.Print "lowbatts.LocalSilenceTime non-zero"
'1720            Debug.Print "datediff " & DateDiff("s", CDate(LowBatts.LocalSilenceTime), CDate(HostTime))
'1730            Debug.Print "rebeep " & Configuration.AlarmReBeep
1740            If DateDiff("s", CDate(LowBatts.LocalSilenceTime), CDate(HostTime)) > Configuration.LowBattReBeep And Configuration.LowBattReBeep > 0 Then
1750              LowBatts.BeepTimer = -1
1760              beeptimers = -1
1770              If MASTER Then
1780                Call UnSilence("MASTER", "", "LowBatts")
1790              Else
                    'ClientUnSilenceAlarms ConsoleID, "", "LowBatts"
1800                QueEvent "ClientUnSilenceAlarms", ConsoleID, "LowBatts"
                    'Debug.Print "ClientUnSilenceAlarms "
1810              End If
1820            Else                 ' still in silence mode
1830              LowBatts.BeepTimer = 0
1840              beeptimers = 0
1850            End If
1860          End If
1870        End If                   ' lowbatts.Count


            ' **************** EXTERNS ******************



1880        If Externs.Count = 0 Then
1890          beeptimers = 0
1900          Externs.BeepTimer = 0  ' need to add each type of alarm
1910        Else                     ' externs.Count > 0

1920          If Externs.LocalSilenceTime = 0 Then
1930            beeptimers = -1
1940            Externs.BeepTimer = -1
1950            Debug.Print "externs.LocalSilenceTime = 0", "externs.LocalAlarmTime = " & CDate(Externs.LocalAlarmTime)
                ' might need to move to end of routine to avoid reentry
1960            If Configuration.AlarmBeep > 0 Then
1970              If Externs.LocalAlarmTime = 0 Then
1980                Externs.LocalAlarmTime = Now  ' fix for deadlock
1990              End If
2000              Debug.Print " DateDiff('s', CDate(externs.LocalAlarmTime), CDate(HostTime)) " & DateDiff("s", CDate(Externs.LocalAlarmTime), CDate(HostTime))

2010              If DateDiff("s", CDate(Externs.LocalAlarmTime), CDate(HostTime)) > Configuration.ExtBeep Then

2020                If MASTER Then
2030                  Call Silence("MASTER", "Externs")
2040                Else
                      'Debug.Print "ClientSilenceAlarms"
                      'ClientSilenceAlarms ConsoleID, "", "Externs"
2050                  QueEvent "ClientSilenceAlarms", ConsoleID, "Externs"
2060                End If
2070              End If
2080            End If
2090          Else                   'externs.LocalLastSilenced > 0
2100            Debug.Print "externs.LocalSilenceTime non-zero"
2110            Debug.Print "datediff " & DateDiff("s", CDate(Externs.LocalSilenceTime), CDate(HostTime))
2120            Debug.Print "rebeep " & Configuration.AlarmReBeep
2130            If DateDiff("s", CDate(Externs.LocalSilenceTime), CDate(HostTime)) > Configuration.ExtReBeep And Configuration.ExtReBeep > 0 Then
2140              Externs.BeepTimer = -1
2150              beeptimers = -1
2160              If MASTER Then
2170                Call UnSilence("MASTER", "", "Externs")
2180              Else
                    'ClientUnSilenceAlarms ConsoleID, "", "Externs"
2190                QueEvent "ClientUnSilenceAlarms", ConsoleID, "Externs"
                    'Debug.Print "ClientUnSilenceAlarms "
2200              End If
2210            Else                 ' still in silence mode
2220              Externs.BeepTimer = 0
2230              beeptimers = 0
2240            End If
2250          End If
2260        End If                   ' externs.Count

            ' ********************* FINISH ***************

2270        beeptimers = (alarms.BeepTimer <> 0) _
                         Or (Alerts.BeepTimer <> 0) _
                         Or (Troubles.BeepTimer <> 0) _
                         Or (LowBatts.BeepTimer <> 0) _
                         Or (Externs.BeepTimer <> 0)

            '2280        beeptimers = (alarms.LocalAlarmTime <> 0)
            '2290        Debug.Print "beeptimers line 2260 " & beeptimers


2280      Else                       ' Configuration.BeepControl = 1 Or 1
            ' use host settings
2290        Debug.Print "Old Beeptimer " & alarms.BeepTimer
2300        beeptimers = (alarms.BeepTimer <> 0) _
                         Or (Alerts.BeepTimer <> 0) _
                         Or (Troubles.BeepTimer <> 0) _
                         Or (LowBatts.BeepTimer <> 0) _
                         Or (Externs.BeepTimer <> 0)
2310      End If

2320      If MASTER Then
2330        If USE6080 Then
              'Me.Caption = "i6080.status " & i6080.Status
2340          If i6080.Status > 0 Then
2350            ShowCommError False
2360          Else
2370            ShowCommError True
2380          End If


2390        Else
'              Dim NC As cESDevice
'2400          Set NC = Devices.Devices(1)
'2410          If NC.CheckinFail Then
'2420            ShowCommError True
'2430          Else
'2440            ShowCommError False
'
'2450          End If
2460        End If
2470      Else

2480      End If


2490      If toggle Then             ' this is the "Waveform" indicator
2500        If (packetizer.BadPackets > BADPACKET_WARNING_CAUTION) Then
2510          imgUptime.Picture = LoadResPicture(LED_AMBER, vbResBitmap)  ' amber
2520        ElseIf (packetizer.BadPackets > BADPACKET_WARNING_DANGER) Then
2530          imgUptime.Picture = LoadResPicture(LED_RED, vbResBitmap)  ' RED
2540        Else
2550          imgUptime.Picture = LoadResPicture(LED_GREEN, vbResBitmap)  ' green
2560        End If
2570      End If                     ' toggle




          Dim playalarm        As Boolean

2580      If CBool(Configuration.BeepControl) Then
2590        playalarm = alarms.LocalAlarmTime <> 0 And alarms.Count > 0
2600      Else
2610        playalarm = alarms.BeepTimer <> 0 And alarms.Count > 0
2620      End If

          ' header above emergency list

2630      If playalarm Then
2640        PlayASound Configuration.AlarmFile, Win32.SND_ASYNC Or Win32.SND_NOSTOP
2650        lblAlarms.BackColor = vbRed
2660        lblAlarms.ForeColor = vbYellow
2670        lblAlarms.Refresh

2680      Else
2690        If alarms.Count > 0 Then
2700          lblAlarms.BackColor = vbYellow
2710          lblAlarms.ForeColor = vbBlack
2720          lblAlarms.Refresh
2730        Else
2740          lblAlarms.BackColor = &H80000002
2750          lblAlarms.ForeColor = &H8000000E
2760          lblAlarms.Refresh
2770        End If
2780      End If

          Dim PlayAlert        As Boolean

2790      If CBool(Configuration.BeepControl) Then
2800        PlayAlert = Alerts.LocalAlarmTime <> 0 And Alerts.Count > 0
2810      Else
2820        PlayAlert = Alerts.BeepTimer <> 0 And Alerts.Count > 0
2830      End If



          '2630      If Alerts.BeepTimer <> 0 And Alerts.Count > 0 Then
2840      If PlayAlert Then


2850        PlayASound Configuration.AlertFile, Win32.SND_ASYNC Or Win32.SND_NOSTOP
2860        cmdAlert.BackColor = vbRed
2870        cmdAlert.Refresh
2880      Else
2890        If Alerts.Count > 0 Then
2900          cmdAlert.BackColor = vbYellow
2910        Else
2920          cmdAlert.BackColor = vbWhite
2930        End If
2940      End If


          Dim PlayTrouble      As Boolean


2950      If CBool(Configuration.BeepControl) Then
2960        PlayTrouble = Troubles.LocalAlarmTime <> 0 And Troubles.Count > 0
2970      Else
2980        PlayTrouble = Troubles.BeepTimer <> 0 And Troubles.Count > 0
2990      End If


          '2740      If Troubles.BeepTimer <> 0 And Troubles.Count > 0 Then
3000      If PlayTrouble Then
3010        PlayASound Configuration.TroubleFile, Win32.SND_ASYNC Or Win32.SND_NOSTOP
3020        cmdTrouble.BackColor = vbRed
3030        cmdTrouble.Refresh
3040      Else
3050        If Troubles.Count > 0 Then
3060          cmdTrouble.BackColor = vbYellow
3070        Else
3080          cmdTrouble.BackColor = vbWhite
3090        End If
3100      End If

          Dim PlayLowBatt      As Boolean

3110      If CBool(Configuration.BeepControl) Then
3120        PlayLowBatt = LowBatts.LocalAlarmTime <> 0 And LowBatts.Count > 0
3130      Else
3140        PlayLowBatt = LowBatts.BeepTimer <> 0 And LowBatts.Count > 0
3150      End If



          '2850      If LowBatts.BeepTimer <> 0 And LowBatts.Count > 0 Then
3160      If PlayLowBatt Then
3170        PlayASound Configuration.LowBattFile, Win32.SND_ASYNC Or Win32.SND_NOSTOP
3180        cmdBattery.BackColor = vbRed
3190        cmdBattery.Refresh
3200      Else
3210        If LowBatts.Count > 0 Then
3220          cmdBattery.BackColor = vbYellow
3230        Else
3240          cmdBattery.BackColor = vbWhite
3250        End If
3260      End If

          Dim PlayExtern       As Boolean


3270      If CBool(Configuration.BeepControl) Then
3280        PlayExtern = Externs.LocalAlarmTime <> 0 And Externs.Count > 0
3290      Else
3300        PlayExtern = Externs.BeepTimer <> 0 And Externs.Count > 0
3310      End If


          '2960      If Externs.BeepTimer <> 0 And Externs.Count > 0 Then
3320      If PlayExtern Then


3330        PlayASound Configuration.ExtFile, Win32.SND_ASYNC Or Win32.SND_NOSTOP
3340        cmdExternal.BackColor = vbRed
3350        cmdExternal.Refresh
3360      Else
3370        If Externs.Count > 0 Then
3380          cmdExternal.BackColor = vbYellow
3390        Else
3400          cmdExternal.BackColor = vbWhite
3410        End If
3420      End If


3430      If playalarm Or PlayAlert Or PlayTrouble Or PlayExtern Or PlayLowBatt Then  ' 0 when all silenced
3440        cmdSilence.BackColor = vbRed  ' toggle on
3450      Else
3460        If cmdSilence.BackColor <> vbWhite Then
3470          cmdSilence.BackColor = vbWhite  ' stay white if no beeptimers
3480        End If
3490      End If



3500      If Assurs.BeepTimer <> 0 Then
3510        If (gAssurDisableScreenOutput = 0) Then
3520          PlayASound Configuration.AssurFile, Win32.SND_ASYNC Or Win32.SND_NOSTOP
3530          cmdAssur.BackColor = vbRed
3540          cmdAssur.Refresh
3550        End If
3560      End If


3570    Else                         ' (not) toggle
          ' makes all backgrounds "normal" on this toggle pass
3580      imgUptime.Picture = LoadResPicture(LED_GRAY, vbResBitmap)  ' gray
3590      cmdBattery.BackColor = vbWhite
3600      cmdTrouble.BackColor = vbWhite
3610      cmdSilence.BackColor = vbWhite
3620      cmdAlert.BackColor = vbWhite
3630      cmdAssur.BackColor = vbWhite
3640      cmdExternal.BackColor = vbWhite
3650      lblAlarms.BackColor = &H80000002
3660      lblAlarms.ForeColor = &H8000000E
3670    End If

        'The sound events that are predefined by the system can vary with the platform. The following list gives the sound events that are defined for all implementations of the Win32 API:
        'SystemAsterisk
        'SystemExclamation
        'SystemExit
        'SystemHand
        'SystemQuestion
        'SystemStart

UpdateUptimeStats_Resume:
3680    On Error GoTo 0
3690    Exit Sub

UpdateUptimeStats_Error:

3700    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.UpdateUptimeStats." & Erl
3710    Resume UpdateUptimeStats_Resume


End Sub



Sub ProcessTroubles()
        ' 32k alarms max
        Dim li                 As ListItem
        Dim j                  As Integer


        Dim alarm              As cAlarm
        'Dim rs                 As Recordset
        'Dim SQl                As String

        Dim COL_SERIAL         As String
        Dim COL_NAME           As String
        Dim COL_ROOM           As String
        Dim col_model          As String
        Dim col_desc           As String
        Dim Col5               As String
        Dim COL_EVENT          As String

        Dim MAXTROUBLES        As Long
10      MAXTROUBLES = 2000

        'Dim selectedid
        Dim SelectedItem       As ListItem
        Dim SelectedSerialnum  As String

20      Set SelectedItem = lvCheckIn.SelectedItem
30      If Not SelectedItem Is Nothing Then
40        SelectedSerialnum = lvCheckIn.SelectedItem.Key
50      End If


60      On Error GoTo ProcessTroubles_Error
70      Call LockWindowUpdate(lvCheckIn.hwnd)

80      lvCheckIn.ListItems.Clear



'90      For j = 1 To Min(Troubles.alarms.Count, MAXTROUBLES)

90      For j = Min(Troubles.alarms.Count, MAXTROUBLES) To 1 Step -1

100       COL_SERIAL = ""
110       COL_NAME = ""
120       COL_ROOM = ""
130       col_model = ""
140       col_desc = ""
          ''Col_silenced = ""
150       COL_EVENT = ""

160       Set alarm = Troubles.alarms(j)


          'If Not Device Is Nothing Then

170       COL_SERIAL = alarm.Serial
180       col_desc = IIf(Len(alarm.Custom), alarm.Custom, alarm.Description)
'190       SQl = " SELECT Devices.Serial, Devices.Model, Residents.NameLast, Residents.NameFirst, Rooms.Room " & _
'              " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
'              " WHERE (Devices.Serial=" & q(Alarm.Serial) & ");"

'200       Set rs = ConnExecute(SQl)
'210       If Not rs.EOF Then
            'Set Device = Devices.Device(Alarm.Serial)
            
220         col_model = alarm.Model ' ders("Model") & ""

            If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
              COL_NAME = ""
            Else
230           COL_NAME = alarm.ResidentText '  ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
            End If
240         COL_ROOM = alarm.RoomText ' rs("Room") & ""

270       Set li = lvCheckIn.ListItems.Add(, alarm.Serial & "S" & alarm.Alarmtype, COL_SERIAL)

          ' name
280       If Len(COL_NAME) = 0 Then
290         li.SubItems(1) = " "
300       Else
310         li.SubItems(1) = COL_NAME
320       End If


          ' room
330       If Len(COL_ROOM) = 0 Then
340         li.SubItems(2) = " "
350       Else
360         li.SubItems(2) = COL_ROOM
370       End If

          'Device
380       If Len(col_model) = 0 Then  ' desc
390         li.SubItems(3) = " "
400       Else
410         li.SubItems(3) = col_model
420       End If

          ' desc
430       li.SubItems(4) = col_desc


          'time
440       li.SubItems(5) = Format(alarm.DateTime, gTimeFormatString)


          ' silenced
450       li.SubItems(6) = IIf(alarm.SilenceTime <> 0, Format(alarm.SilenceTime, gTimeFormatString), " ")


          ' event type
460       Select Case alarm.Alarmtype
            Case EVT_CHECKIN_FAIL
470           COL_EVENT = "Supv"
480         Case EVT_TAMPER
490           COL_EVENT = "Tamp"
500         Case EVT_COMM_TIMEOUT
510           COL_EVENT = "Comm"
520         Case EVT_STRAY
530           COL_EVENT = "Stray"
540         Case EVT_LINELOSS
550           COL_EVENT = "NoAC"
560         Case EVT_JAMMED
570           COL_EVENT = "Jamm"
580         Case EVT_SERVER_TROUBLE
590           COL_EVENT = "Pager Trouble"
600         Case Else
610           COL_EVENT = "UNK"
620       End Select
630       li.SubItems(7) = COL_EVENT

640     Next

650     LockWindowUpdate 0

660     If lvCheckIn.ListItems.Count Then
670       For j = 1 To lvCheckIn.ListItems.Count
680         If SelectedSerialnum = lvCheckIn.ListItems(j).Key Then
690           lvCheckIn.ListItems(j).EnsureVisible
700           lvCheckIn.ListItems(j).Selected = True
710           Exit For
720         End If
730       Next

          'If selectedid > 0 And selectedid <= lvCheckIn.ListItems.count Then
          'lvCheckIn.ListItems(selectedid).Selected = True
          'lvCheckIn.ListItems(selectedid).EnsureVisible

          'Else

          ' lvCheckIn.ListItems(lvCheckIn.ListItems.count).Selected = True
          ' lvCheckIn.ListItems(lvCheckIn.ListItems.count).EnsureVisible
          ' End If
740     End If


ProcessTroubles_Resume:
750     On Error GoTo 0
760     LockWindowUpdate 0
770     Exit Sub

ProcessTroubles_Error:

780     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.ProcessTroubles." & Erl
790     Resume ProcessTroubles_Resume


End Sub

'Public Sub UpdateLocations(P As cESPacket)
'  Dim li As ListItem
'  Dim lv As ListView
'  Dim ID As String
'  Dim Location As String
'
'  Exit Sub
'
'  ' dead wood
'  On Error GoTo UpdateLocations_Error
'
'  'Location = GetLocater(p.Serial)
'
'  'id = p.Serial
'  'If id <> 0 Then
'  If 0 <> alarms.UpdateLocation(P) Then
'    Set lv = lvEmergency
'    Trace "UpdateLocations: " & P.Serial
'    For Each li In lv.ListItems
'      If Val(li.Key) = P.Serial Then
'        li.ListSubItems(3).text = Location
'        UpdatePageLocation ID, Location  ' uses unique ID
'        Trace "Located: " & Right("00000000" & Hex(P.Serial), 4) & " " & Now
'        Exit For
'      End If
'    Next
'    ' push location info to paging
'  End If
'
'  If 0 <> Alerts.UpdateLocation(P) Then
'    Set lv = lvAlerts
'    For Each li In lv.ListItems
'      If Val(li.Key) = P.Serial Then
'        li.ListSubItems(3).text = Location
'        UpdatePageLocation ID, Location  ' uses unique ID
'        Trace "Located: " & Right("00000000" & Hex(P.Serial), 4) & " " & Now
'        Exit For
'      End If
'    Next
'    ' push location info to paging
'  End If
'
'
'  If 0 <> LowBatts.UpdateLocation(P) Then
'    Set lv = Me.lvLoBatt
'    For Each li In lv.ListItems
'      If Val(li.Key) = P.Serial Then
'        li.ListSubItems(3).text = Location
'        Trace "Located: " & Right("00000000" & Hex(P.Serial), 4) & " " & Now
'        Exit For
'      End If
'    Next
'  End If
'  '        If p.serial = 12554 Then Stop
'  If Troubles.count > 0 Then
'    If 0 <> Troubles.UpdateLocation(P) Then
'      Set lv = Me.lvCheckIn
'      For Each li In lv.ListItems
'        If Val(li.Key) = P.Serial Then
'          li.ListSubItems(3).text = Location
'          Trace "Located: " & Right("00000000" & Hex(P.Serial), 4) & " " & Now
'          Exit For
'        End If
'      Next
'    End If
'  End If
'
'  'End If
'
'
'
'
'
'
'  ' if we recieved a location event, then it's not a trouble, is it?
'  ' note we don't do troubles since there's no data being received to give us a location.
'
'UpdateLocations_Resume:
'  On Error GoTo 0
'  Exit Sub
'
'UpdateLocations_Error:
'
'  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmMain.UpdateLocations." & Erl
'  Resume UpdateLocations_Resume
'
'
'End Sub

Sub ConfigurelvEmergency()
  Dim ch As ColumnHeader
  Dim lv As ListView
  Set lv = lvEmergency
  lv.ColumnHeaders.Clear

  Set ch = lv.ColumnHeaders.Add(, "ID", "ID", 0)          ' 0
  Set ch = lv.ColumnHeaders.Add(, "Name", "Name", 1800)   ' 1
  Set ch = lv.ColumnHeaders.Add(, "Room", "Room")         ' 2
  Set ch = lv.ColumnHeaders.Add(, "Location", "Location") ' 3
  Set ch = lv.ColumnHeaders.Add(, "Time", "Time", 900)    ' 4
  Set ch = lv.ColumnHeaders.Add(, "Announce", "Announce", 1600) ' 5
  Set ch = lv.ColumnHeaders.Add(, "Ack", "Ack", 900)      ' 6
  Set ch = lv.ColumnHeaders.Add(, "Resp", "Resp", 900)    ' 7
  

End Sub

Sub ConfigurelvAlerts()
  Dim ch As ColumnHeader
  Dim lv As ListView
  Set lv = lvAlerts

  lv.ColumnHeaders.Clear
  Set ch = lv.ColumnHeaders.Add(, "ID", "ID", 0)
  Set ch = lv.ColumnHeaders.Add(, "Name", "Name", 1800)
  Set ch = lv.ColumnHeaders.Add(, "Room", "Room")
  Set ch = lv.ColumnHeaders.Add(, "Location", "Location")
  Set ch = lv.ColumnHeaders.Add(, "Time", "Time", 900)
  Set ch = lv.ColumnHeaders.Add(, "Announce", "Announce", 1600)
  Set ch = lv.ColumnHeaders.Add(, "Ack", "Ack", 900)
  Set ch = lv.ColumnHeaders.Add(, "Resp", "Resp", 900)

End Sub

Sub ConfigurelvTrouble()
  Dim ch As ColumnHeader
  lvCheckIn.ColumnHeaders.Clear
  '  Col  SubItem
  Set ch = lvCheckIn.ColumnHeaders.Add(, "ID", "ID", 0)                 '   1
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Name", "Name", 1800)          '   2
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Room", "Room")                '   3
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Device", "Device", 1000)      '   4
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Desc", "Desc")                '   5
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Time", "Time", 900)           '   6
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Silenced", "Silenced", 1200)  '   7
  Set ch = lvCheckIn.ColumnHeaders.Add(, "Type", "Type", 900)           '   8
End Sub
Sub ConfigurelvLoBatt()
  Dim ch As ColumnHeader
  lvLoBatt.ColumnHeaders.Clear
  ' Col  SubItem
  Set ch = lvLoBatt.ColumnHeaders.Add(, "ID", "ID", 0)                '1
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Name", "Name", 1800)         '2
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Room", "Room")               '3
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Device", "Device", 1000)           '4
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Desc", "Desc")               '5
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Time", "Time", 900)          '6
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Silenced", "Silenced", 1200) '7
  Set ch = lvLoBatt.ColumnHeaders.Add(, "Type", " ", 900)                 '8
End Sub

Sub ConfigurelvAssur()
  Dim ch As ColumnHeader
  lvAssur.ColumnHeaders.Clear

  Set ch = lvAssur.ColumnHeaders.Add(, "ID", "ID", 0)
  Set ch = lvAssur.ColumnHeaders.Add(, "Name", "Name", 2800)
  Set ch = lvAssur.ColumnHeaders.Add(, "Room", "Room")
  Set ch = lvAssur.ColumnHeaders.Add(, "Phone", "Phone")

End Sub
Sub ConfigurelvExtern()
  Dim ch As ColumnHeader
  lvExternal.ColumnHeaders.Clear

  Set ch = lvExternal.ColumnHeaders.Add(, "ID", "ID", 0)
  Set ch = lvExternal.ColumnHeaders.Add(, "Name", "Name", 1800)
  Set ch = lvExternal.ColumnHeaders.Add(, "Room", "Room")
  Set ch = lvExternal.ColumnHeaders.Add(, "Time", "Time", 900)  '   5       4
  Set ch = lvExternal.ColumnHeaders.Add(, "Event", "Event", 4600)

End Sub


Private Sub TimerLogon_Timer()


  If frmLogin.Visible Then
    
    gLockTimeRemaining = Max(0, gLockTimeRemaining - 1)
    frmLogin.CountDown = gLockTimeRemaining
    If gLockTimeRemaining <= 0 Then
      If ExistsUser0000() Then
        '        frmLogin.Caption = "System Login " & gLockTimeRemaining
        frmLogin.txtLogin.text = "0000"
        frmLogin.DoOK
        RemoveHostedForms
        ResetLockTime
        ResetActivityTime
      Else
      
      End If
      
    End If
  Else
    ' do the same for inactivity timer if user level is >= LEVEL_SUPERVISOR
    'levels are user - supervisor - admin - factory
    If gUser.LEvel >= LEVEL_SUPERVISOR Then
      gInactivityTimeRemaining = Max(0, gInactivityTimeRemaining - 1)
      If frmMain.ShowCountdown Then
        If cmdLogin.BackColor = vbYellow Then
           cmdLogin.BackColor = Me.BackColor
        Else
          cmdLogin.BackColor = vbYellow
        End If
        If gInactivityTimeRemaining <= 60 Then
          cmdLogin.Caption = "Logoff " & vbCrLf & gInactivityTimeRemaining
        Else
          cmdLogin.Caption = "Logoff"
        End If
      Else
        If cmdLogin.Caption <> "Logoff" Then
          cmdLogin.Caption = "Logoff"
        End If
      End If

      If gInactivityTimeRemaining <= 0 Then
        ResetActivityTime
        If ExistsUser0000() Then
          DoLogin
          RemoveHostedForms
        End If
      End If
    Else
      If gUser.LEvel <= LEVEL_USER Then
        If cmdLogin.Caption <> "Login" Then
           cmdLogin.Caption = "Login"
           
        End If
        If cmdLogin.BackColor <> Me.BackColor Then
          cmdLogin.BackColor = Me.BackColor
        End If
        
      ElseIf cmdLogin.Caption <> "Logoff" Then
        cmdLogin.Caption = "Logoff"
      End If
    End If
  End If

End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      DoLogin
      Form_Paint
  End Select
End Sub


Public Sub ClearInfoBox(ByVal AlarmID As Long)
  Dim CurrrentAlarmID As Long
  'When an alarm is dismissed, clear the info box

  CurrrentAlarmID = Val(txtHiddenAlarmID.text)
  If CurrrentAlarmID <> 0 Then
    If CurrrentAlarmID = AlarmID Then
      DisplayResidentInfo 0, 0
    End If
  End If



End Sub
