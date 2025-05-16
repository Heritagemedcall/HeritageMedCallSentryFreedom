VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmConfigure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Configuration"
   ClientHeight    =   15450
   ClientLeft      =   7440
   ClientTop       =   3810
   ClientWidth     =   20550
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
   ScaleHeight     =   15450
   ScaleWidth      =   20550
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
      Height          =   15210
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   18765
      Begin VB.Frame fraSoftPoints 
         BackColor       =   &H0000FFFF&
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
         Height          =   2835
         Left            =   -105
         TabIndex        =   183
         Top             =   12300
         Visible         =   0   'False
         Width           =   9255
         Begin VB.CheckBox chkLocationText 
            Caption         =   "Use ""Location"" Phrase"
            Height          =   285
            Left            =   4110
            TabIndex        =   188
            ToolTipText     =   "Forward Discovered Soft Points to Text Message"
            Top             =   450
            Width           =   2505
         End
         Begin VB.CommandButton cmdCreateFromRooms 
            Caption         =   "Rooms to Partitions"
            Height          =   450
            Left            =   1740
            TabIndex        =   179
            ToolTipText     =   "Convert Rooms to Partitions"
            Top             =   360
            Width           =   1590
         End
         Begin VB.TextBox txtSMSAccount 
            Height          =   300
            Left            =   240
            MaxLength       =   40
            TabIndex        =   182
            ToolTipText     =   "SMS Account Identifier"
            Top             =   2040
            Width           =   4020
         End
         Begin VB.CommandButton cmdPartitions 
            Caption         =   "Partitons"
            Height          =   450
            Left            =   240
            TabIndex        =   178
            ToolTipText     =   "Mange Partitions"
            Top             =   360
            Width           =   1365
         End
         Begin VB.CheckBox chkForwardSP 
            Caption         =   "Forward via SMS"
            Height          =   285
            Left            =   240
            TabIndex        =   181
            ToolTipText     =   "Forward Discovered Soft Points to Text Message"
            Top             =   1380
            Width           =   2505
         End
         Begin VB.CommandButton cmdEditSoftPoints 
            Caption         =   "Soft Points"
            Height          =   450
            Left            =   240
            TabIndex        =   180
            ToolTipText     =   "Manage Soft Points"
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recipient"
            Height          =   195
            Left            =   270
            TabIndex        =   186
            Top             =   1740
            Width           =   825
         End
      End
      Begin VB.Frame fraWatchdog 
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
         Height          =   3315
         Left            =   9120
         TabIndex        =   174
         Top             =   11640
         Width           =   9255
         Begin VB.TextBox txtMonitorFacilityID 
            Height          =   300
            Left            =   1590
            MaxLength       =   25
            TabIndex        =   141
            ToolTipText     =   "Uniquely Identify Facility to Monitoring Station"
            Top             =   1980
            Width           =   1830
         End
         Begin VB.CheckBox chkMonitorEnabled 
            Alignment       =   1  'Right Justify
            Caption         =   "Enable Monitoring"
            Height          =   315
            Left            =   540
            TabIndex        =   142
            Top             =   2460
            Width           =   2205
         End
         Begin VB.CommandButton cmdTestMonitor 
            Caption         =   "Monitor Test"
            Height          =   315
            Left            =   3720
            TabIndex        =   143
            Top             =   2460
            Width           =   1395
         End
         Begin VB.TextBox txtMonitorRequest 
            Height          =   300
            Left            =   1590
            MaxLength       =   100
            TabIndex        =   137
            ToolTipText     =   "Folder of Sub Folder to Send Updates to"
            Top             =   1620
            Width           =   3870
         End
         Begin VB.TextBox txtMonitorPort 
            Height          =   300
            Left            =   6735
            MaxLength       =   5
            TabIndex        =   135
            ToolTipText     =   "Port to Use Other than Port 80"
            Top             =   1260
            Width           =   750
         End
         Begin VB.TextBox txtMonitorDomain 
            Height          =   300
            Left            =   1590
            MaxLength       =   100
            TabIndex        =   133
            ToolTipText     =   "IP Address or Domain Name to Send Updates to"
            Top             =   1260
            Width           =   3870
         End
         Begin VB.TextBox txtMonitorInterval 
            Height          =   300
            Left            =   6735
            MaxLength       =   5
            TabIndex        =   139
            ToolTipText     =   "How Often to Send Updates to Monitoring Station"
            Top             =   1620
            Width           =   750
         End
         Begin VB.Timer TimerWD 
            Enabled         =   0   'False
            Interval        =   2000
            Left            =   7500
            Top             =   2100
         End
         Begin VB.ComboBox cboWDType 
            Height          =   315
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtWDTimeout 
            Height          =   300
            Left            =   3555
            MaxLength       =   5
            TabIndex        =   127
            Top             =   502
            Width           =   750
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facility ID"
            Height          =   195
            Left            =   555
            TabIndex        =   140
            Top             =   2040
            Width           =   870
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Page/Folder"
            Height          =   195
            Left            =   360
            TabIndex        =   136
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special Port"
            Height          =   195
            Left            =   5580
            TabIndex        =   134
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP/Domain"
            Height          =   195
            Left            =   510
            TabIndex        =   132
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Onboard Watchdog"
            Height          =   195
            Left            =   420
            TabIndex        =   123
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interval"
            Height          =   195
            Left            =   5970
            TabIndex        =   138
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sec"
            Height          =   195
            Left            =   7620
            TabIndex        =   144
            Top             =   1680
            Width           =   345
         End
         Begin VB.Label lblRemote 
            AutoSize        =   -1  'True
            Caption         =   "Remote Monitoring"
            Height          =   195
            Left            =   540
            TabIndex        =   131
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "    "
            Height          =   195
            Left            =   6360
            TabIndex        =   130
            Top             =   555
            Width           =   255
         End
         Begin VB.Label lblStatuslbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   195
            Left            =   5700
            TabIndex        =   129
            Top             =   555
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seconds"
            Height          =   195
            Left            =   4440
            TabIndex        =   128
            Top             =   555
            Width           =   750
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   195
            Left            =   420
            TabIndex        =   124
            Top             =   540
            Width           =   435
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Timeout"
            Height          =   195
            Left            =   2760
            TabIndex        =   126
            Top             =   555
            Width           =   690
         End
      End
      Begin VB.Frame fraExternal 
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
         Height          =   2835
         Left            =   9150
         TabIndex        =   161
         Top             =   9510
         Width           =   9255
         Begin VB.CommandButton cmdMobileSettings 
            Caption         =   "Mobile Settings"
            Height          =   540
            Left            =   4080
            TabIndex        =   202
            Top             =   375
            Width           =   1095
         End
         Begin VB.CommandButton cmdMobile 
            Caption         =   "Web Security"
            Height          =   540
            Left            =   2835
            TabIndex        =   200
            Top             =   390
            Width           =   1095
         End
         Begin VB.CommandButton cmdPush 
            Caption         =   "Push"
            Height          =   540
            Left            =   1627
            TabIndex        =   199
            Top             =   390
            Width           =   1095
         End
         Begin VB.CommandButton cmdStressTest 
            Caption         =   "Stress Test"
            Height          =   570
            Left            =   390
            TabIndex        =   189
            Top             =   1455
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdDukane 
            Caption         =   "Dukane"
            Height          =   540
            Left            =   420
            TabIndex        =   164
            Top             =   390
            Width           =   1095
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Not Used"
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
            Left            =   240
            TabIndex        =   162
            Top             =   120
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.Frame fraRegistration 
         BackColor       =   &H00FF00FF&
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
         Height          =   2820
         Left            =   9435
         TabIndex        =   116
         Top             =   6570
         Width           =   9180
         Begin VB.CommandButton cmdRegister 
            Caption         =   "Register"
            Height          =   405
            Left            =   5370
            TabIndex        =   122
            Top             =   1980
            Width           =   1485
         End
         Begin VB.TextBox txtMID 
            Height          =   360
            Left            =   645
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   375
            Width           =   6825
         End
         Begin VB.TextBox txtEID 
            Height          =   360
            Left            =   645
            TabIndex        =   120
            Top             =   990
            Width           =   6825
         End
         Begin VB.Label lblMailTo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "email@heritagemedcall.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   1155
            TabIndex        =   148
            Top             =   2385
            Width           =   3180
         End
         Begin VB.Label lblFax 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "fax: 813-223-1405"
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
            Left            =   1755
            TabIndex        =   147
            Top             =   2115
            Width           =   1980
         End
         Begin VB.Label lblPhone 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "813-221-1000 / 800-396-6157"
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
            TabIndex        =   146
            Top             =   1845
            Width           =   3105
         End
         Begin VB.Label lblCompany 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Heritage MedCall, Inc."
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
            Left            =   1500
            TabIndex        =   145
            Top             =   1575
            Width           =   2490
         End
         Begin VB.Label lblDevices 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   6075
            TabIndex        =   121
            Top             =   1620
            Width           =   75
         End
         Begin VB.Label lbl4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registration Code"
            Height          =   195
            Left            =   645
            TabIndex        =   119
            Top             =   765
            Width           =   1530
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Code"
            Height          =   195
            Left            =   645
            TabIndex        =   117
            Top             =   150
            Width           =   1290
         End
      End
      Begin VB.Frame fraSystem 
         BackColor       =   &H00FFC0C0&
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
         Height          =   2910
         Left            =   75
         TabIndex        =   68
         Top             =   6480
         Width           =   8055
         Begin VB.CheckBox chkHippaSidebar 
            Alignment       =   1  'Right Justify
            Caption         =   "Hide Sidebar"
            Height          =   240
            Left            =   5520
            TabIndex        =   203
            Top             =   2595
            Width           =   2205
         End
         Begin VB.CommandButton cmdFactory 
            Caption         =   "Factory Settings"
            Height          =   855
            Left            =   7050
            TabIndex        =   76
            Top             =   870
            Width           =   990
         End
         Begin VB.CheckBox chkHIPPANames 
            Alignment       =   1  'Right Justify
            Caption         =   "Hide Resident Names"
            Height          =   240
            Left            =   5520
            TabIndex        =   190
            Top             =   2355
            Width           =   2205
         End
         Begin VB.ComboBox cboThird 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   2543
            Width           =   1035
         End
         Begin VB.ComboBox cboSecond 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   172
            Top             =   2190
            Width           =   1035
         End
         Begin VB.ComboBox cboFirst 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   1845
            Width           =   1035
         End
         Begin VB.TextBox txtDBInfo 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   166
            Top             =   420
            Width           =   7095
         End
         Begin VB.CommandButton cmdEmailSetup 
            Caption         =   "Email Setup"
            Height          =   855
            Left            =   6045
            TabIndex        =   75
            Top             =   870
            Width           =   990
         End
         Begin VB.CommandButton cmdDebug 
            Caption         =   "Debug"
            Height          =   855
            Left            =   3030
            TabIndex        =   72
            Top             =   870
            Width           =   990
         End
         Begin VB.CommandButton cmdExternalApps 
            Caption         =   "Run Other Programs"
            Height          =   855
            Left            =   5040
            TabIndex        =   74
            Top             =   870
            Width           =   990
         End
         Begin VB.CheckBox chkElapsedEqACK 
            Alignment       =   1  'Right Justify
            Caption         =   "Elapsed Time = Ack"
            CausesValidation=   0   'False
            Height          =   255
            Left            =   5700
            TabIndex        =   85
            Top             =   2100
            Width           =   2025
         End
         Begin VB.CheckBox chkTimeFormat 
            Alignment       =   1  'Right Justify
            Caption         =   "24Hr Alarm Format"
            Height          =   255
            Left            =   5700
            TabIndex        =   84
            Top             =   1860
            Width           =   2025
         End
         Begin VB.CommandButton cmdBackupSettings 
            Caption         =   "Backup Settings"
            Height          =   855
            Left            =   4035
            TabIndex        =   73
            Top             =   870
            Width           =   990
         End
         Begin VB.TextBox txtAppPath 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   157
            Top             =   60
            Width           =   7095
         End
         Begin VB.CommandButton cmdScreenMask 
            Caption         =   "Battery/ Trouble Setup"
            Height          =   855
            Left            =   2025
            TabIndex        =   71
            Top             =   870
            Width           =   990
         End
         Begin VB.CommandButton cmdEditUsers 
            Caption         =   "Edit Users"
            Height          =   855
            Left            =   1020
            TabIndex        =   70
            Top             =   870
            Width           =   990
         End
         Begin VB.CommandButton cmdDeviceTypes 
            Caption         =   "Edit Device Types"
            Height          =   855
            Left            =   15
            TabIndex        =   69
            Top             =   870
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(End)"
            Height          =   195
            Left            =   473
            TabIndex        =   170
            Top             =   2220
            Width           =   465
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0-23 hr)"
            Height          =   195
            Left            =   4005
            TabIndex        =   169
            Top             =   2610
            Width           =   735
         End
         Begin VB.Label lblEndThird 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2970
            TabIndex        =   168
            Top             =   2610
            Width           =   555
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Third"
            Height          =   195
            Left            =   1410
            TabIndex        =   167
            Top             =   2610
            Width           =   450
         End
         Begin VB.Label lblNightPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shift Times"
            Height          =   195
            Left            =   225
            TabIndex        =   77
            Top             =   1965
            Width           =   960
         End
         Begin VB.Label lblStartNight 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First"
            Height          =   195
            Left            =   1485
            TabIndex        =   78
            Top             =   1905
            Width           =   375
         End
         Begin VB.Label lblEndNight 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Second"
            Height          =   195
            Left            =   1200
            TabIndex        =   81
            Top             =   2250
            Width           =   660
         End
         Begin VB.Label lblEndNightHR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2970
            TabIndex        =   82
            Top             =   2250
            Width           =   555
         End
         Begin VB.Label lblStartNightHR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2970
            TabIndex        =   79
            Top             =   1905
            Width           =   555
         End
         Begin VB.Label lblNightRange2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0-23 hr)"
            Height          =   195
            Left            =   4005
            TabIndex        =   83
            Top             =   2250
            Width           =   735
         End
         Begin VB.Label lblNightRange1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0-23 hr)"
            Height          =   195
            Left            =   4005
            TabIndex        =   80
            Top             =   1905
            Width           =   735
         End
      End
      Begin VB.Frame framain 
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
         Height          =   2895
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   8175
         Begin VB.TextBox txtRemoteSerial 
            Height          =   300
            Left            =   1815
            MaxLength       =   8
            TabIndex        =   105
            ToolTipText     =   "Device Serial Number For Monitoring This Remote"
            Top             =   1380
            Width           =   1740
         End
         Begin VB.CommandButton cmdChangeHostAdapter 
            Caption         =   "Adapter"
            Height          =   330
            Left            =   4680
            TabIndex        =   187
            ToolTipText     =   "Change Username and Password on ACG"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txt6080Password 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3360
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   114
            ToolTipText     =   "ACG Password"
            Top             =   2400
            Width           =   1500
         End
         Begin VB.TextBox txt6080UserName 
            Height          =   300
            Left            =   1815
            MaxLength       =   15
            TabIndex        =   113
            ToolTipText     =   "ACGT User Name"
            Top             =   2400
            Width           =   1500
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Download"
            Height          =   390
            Left            =   4680
            TabIndex        =   176
            ToolTipText     =   "Capture ACG Data"
            Top             =   1470
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CheckBox chkSubscribedAlarms 
            Alignment       =   1  'Right Justify
            Caption         =   "My Alarms Only"
            Height          =   315
            Left            =   3510
            TabIndex        =   165
            Top             =   1005
            Width           =   2205
         End
         Begin VB.Frame fraNetwork 
            BorderStyle     =   0  'None
            Caption         =   "Network"
            Height          =   2685
            Left            =   6450
            TabIndex        =   149
            Top             =   120
            Width           =   1635
            Begin VB.ComboBox cboNID 
               Height          =   315
               Left            =   120
               TabIndex        =   150
               Text            =   "cboNID"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton cmdUpgrade 
               Caption         =   "Upgrade"
               Height          =   390
               Left            =   120
               TabIndex        =   154
               ToolTipText     =   "Upgrade From Legacy to 6080 ACG"
               Top             =   1620
               Width           =   1365
            End
            Begin VB.CommandButton cmdRepeaters 
               Caption         =   "Repeaters..."
               Height          =   390
               Left            =   120
               TabIndex        =   153
               ToolTipText     =   "Syncronizes NIDs of Two-Way Devices"
               Top             =   1590
               Width           =   1365
            End
            Begin VB.CommandButton cmdImport 
               Caption         =   "Import..."
               Height          =   420
               Left            =   120
               TabIndex        =   155
               ToolTipText     =   "Import Names and Rooms"
               Top             =   2070
               Width           =   1365
            End
            Begin VB.CommandButton cmdSetNID 
               Caption         =   "Change NID"
               Height          =   390
               Left            =   120
               TabIndex        =   151
               ToolTipText     =   "Changes Network ID of NC"
               Top             =   690
               Width           =   1365
            End
            Begin VB.CommandButton cmdSyncNIDs 
               Caption         =   "Sync NIDs"
               Height          =   390
               Left            =   120
               TabIndex        =   152
               ToolTipText     =   "Syncronizes NIDs of Two-Way Devices"
               Top             =   1110
               Width           =   1365
            End
            Begin VB.Label lblNID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Network ID"
               Height          =   195
               Left            =   165
               TabIndex        =   156
               Top             =   45
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkDirectedNet 
            Alignment       =   1  'Right Justify
            Caption         =   "Directed Message NC"
            Height          =   315
            Left            =   3510
            TabIndex        =   158
            Top             =   690
            Width           =   2205
         End
         Begin VB.TextBox txtIP 
            Height          =   300
            Left            =   3015
            MaxLength       =   15
            TabIndex        =   111
            Top             =   2040
            Width           =   1470
         End
         Begin VB.TextBox txtHostPort 
            Height          =   300
            Left            =   1815
            MaxLength       =   5
            TabIndex        =   109
            Top             =   2040
            Width           =   750
         End
         Begin VB.TextBox txtRxLocation 
            Height          =   300
            Left            =   1815
            MaxLength       =   25
            TabIndex        =   107
            Top             =   1710
            Width           =   2550
         End
         Begin VB.TextBox txtRxSerial 
            Height          =   300
            Left            =   1815
            MaxLength       =   8
            TabIndex        =   104
            Top             =   1380
            Width           =   1500
         End
         Begin VB.TextBox txtCommPort 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1815
            MaxLength       =   3
            TabIndex        =   9
            ToolTipText     =   "Com Port for Receiver"
            Top             =   720
            Width           =   780
         End
         Begin VB.TextBox txtCommTimeout 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1815
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1050
            Width           =   780
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3735
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   6
            Top             =   390
            Width           =   1770
         End
         Begin VB.TextBox txtFacility 
            Height          =   300
            Left            =   90
            MaxLength       =   25
            TabIndex        =   5
            Top             =   390
            Width           =   3600
         End
         Begin VB.CommandButton cmd6080 
            Caption         =   "6080"
            Height          =   330
            Left            =   4920
            TabIndex        =   115
            ToolTipText     =   "Change Username and Password on ACG"
            Top             =   2400
            Width           =   825
         End
         Begin VB.CommandButton cmdUpdate6080UP 
            Caption         =   "Set"
            Enabled         =   0   'False
            Height          =   330
            Left            =   4920
            TabIndex        =   177
            ToolTipText     =   "Syncronizes NIDs of Two-Way Devices"
            Top             =   2385
            Width           =   825
         End
         Begin VB.TextBox txtAGPIP 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1815
            MaxLength       =   15
            TabIndex        =   8
            ToolTipText     =   "IP Address of ACG"
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label lblRemoteSerial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "My DeviceID"
            Height          =   195
            Left            =   570
            TabIndex        =   191
            Top             =   1440
            Width           =   1110
         End
         Begin VB.Label lblACGUP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACG User/Pwd"
            Height          =   195
            Left            =   285
            TabIndex        =   112
            Top             =   2460
            Width           =   1290
         End
         Begin VB.Label lblACGIP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACG IP"
            Height          =   195
            Left            =   1080
            TabIndex        =   175
            Top             =   765
            Width           =   630
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP"
            Height          =   195
            Left            =   2745
            TabIndex        =   110
            Top             =   2100
            Width           =   195
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Host Port"
            Height          =   195
            Left            =   915
            TabIndex        =   108
            Top             =   2100
            Width           =   810
         End
         Begin VB.Label lblRxLocation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receiver Location"
            Height          =   195
            Left            =   135
            TabIndex        =   106
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label lblCommPort 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Com Port"
            Height          =   195
            Left            =   930
            TabIndex        =   7
            Top             =   765
            Width           =   780
         End
         Begin VB.Label lblRxMinutes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minutes"
            Height          =   195
            Left            =   2670
            TabIndex        =   12
            Top             =   1110
            Width           =   675
         End
         Begin VB.Label lblrxTimout 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receiver Timeout"
            Height          =   195
            Left            =   195
            TabIndex        =   10
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label lblID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Console ID"
            Height          =   195
            Left            =   3990
            TabIndex        =   4
            Top             =   105
            Width           =   945
         End
         Begin VB.Label lblFacility 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facilty / Station Name"
            Height          =   195
            Left            =   255
            TabIndex        =   3
            Top             =   105
            Width           =   1920
         End
         Begin VB.Label lblRxSerial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receiver Serial"
            Height          =   195
            Left            =   390
            TabIndex        =   103
            Top             =   1440
            Width           =   1320
         End
      End
      Begin VB.Frame frawaypts 
         BackColor       =   &H00FF8080&
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
         Height          =   2745
         Left            =   10350
         TabIndex        =   86
         Top             =   450
         Width           =   8160
         Begin VB.TextBox txtBoost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   6285
            MaxLength       =   2
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   1890
            Width           =   525
         End
         Begin VB.CheckBox chkLocationPhrase 
            Caption         =   "Use ""Location"" Phrase"
            Height          =   285
            Left            =   5370
            TabIndex        =   90
            ToolTipText     =   "Forward Discovered Soft Points to Text Message"
            Top             =   780
            Width           =   2505
         End
         Begin VB.CheckBox chkNoNCs 
            Caption         =   "Ignore NCs"
            Height          =   285
            Left            =   5370
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   360
            Width           =   2115
         End
         Begin VB.TextBox txtWaypointDevice 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   3525
            MaxLength       =   8
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   1860
            Width           =   1500
         End
         Begin VB.TextBox txtPCA 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3540
            MaxLength       =   8
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1500
         End
         Begin VB.ComboBox cboWaypointMode 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   1260
            Width           =   2685
         End
         Begin VB.ComboBox cboAvail 
            Height          =   315
            Left            =   3510
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1260
            Width           =   2355
         End
         Begin VB.CheckBox chkUseOnlyLocators 
            Caption         =   "Only Use Locators"
            Height          =   285
            Left            =   5370
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.CommandButton cmdPrintWaypoints 
            Caption         =   "Print Waypoints"
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
            Left            =   3015
            TabIndex        =   88
            Top             =   315
            Width           =   1980
         End
         Begin VB.CheckBox chkPCARedirect 
            Caption         =   "Redirect to PCA"
            Height          =   285
            Left            =   3540
            TabIndex        =   101
            Top             =   2340
            Width           =   2505
         End
         Begin VB.TextBox txtSurveyDeviceID 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            MaxLength       =   8
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   1890
            Width           =   1500
         End
         Begin VB.CommandButton cmdWayppoints 
            Caption         =   "Edit Waypoints"
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
            Left            =   600
            TabIndex        =   87
            Top             =   315
            Width           =   1980
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1221S-60 Boost %"
            Height          =   195
            Left            =   5910
            TabIndex        =   201
            Top             =   1650
            Width           =   1560
         End
         Begin VB.Label lblVerifier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Walker"
            Height          =   195
            Left            =   3540
            TabIndex        =   163
            Top             =   1650
            Width           =   615
         End
         Begin VB.Label lblSurveyPager 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Survey Pager"
            Height          =   195
            Left            =   3555
            TabIndex        =   96
            Top             =   990
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Survey Mode"
            Height          =   195
            Left            =   600
            TabIndex        =   91
            Top             =   990
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Survey Transmitter"
            Height          =   195
            Left            =   600
            TabIndex        =   93
            Top             =   1650
            Width           =   1605
         End
         Begin VB.Label lblPCA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Survey PCA"
            Height          =   195
            Left            =   3540
            TabIndex        =   94
            Top             =   990
            Width           =   1020
         End
      End
      Begin VB.Frame fraAssur 
         BackColor       =   &H008080FF&
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
         Height          =   2910
         Left            =   10170
         TabIndex        =   13
         Top             =   3510
         Width           =   8265
         Begin VB.CheckBox chkDisableScreenOutput 
            Caption         =   "Disable Screen Output"
            Height          =   345
            Left            =   4500
            TabIndex        =   159
            Top             =   2010
            Width           =   3195
         End
         Begin VB.CommandButton cmdAssurDest 
            Caption         =   "Send To"
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
            Left            =   6600
            TabIndex        =   160
            Top             =   660
            Width           =   1175
         End
         Begin VB.TextBox txtAssurStart2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   26
            ToolTipText     =   "0 is Midnight"
            Top             =   1755
            Width           =   585
         End
         Begin VB.TextBox txtAssurEnd2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   31
            ToolTipText     =   "0 is Midnight"
            Top             =   2085
            Width           =   585
         End
         Begin VB.TextBox txtAssurEnd 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   21
            ToolTipText     =   "0 is Midnight"
            Top             =   945
            Width           =   585
         End
         Begin VB.TextBox txtAssurStart 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   16
            ToolTipText     =   "0 is Midnight"
            Top             =   615
            Width           =   585
         End
         Begin VB.Label lblNote2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "One Hour minimum delay between Check-in Periods"
            Height          =   765
            Left            =   4365
            TabIndex        =   29
            Top             =   1245
            Width           =   1830
         End
         Begin VB.Label lblNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Check-in Periods may not overlap.  "
            Height          =   525
            Left            =   4365
            TabIndex        =   19
            Top             =   705
            Width           =   1830
         End
         Begin VB.Label lblAssurPd2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check-in Period #2"
            Height          =   195
            Left            =   1005
            TabIndex        =   24
            Top             =   1470
            Width           =   1650
         End
         Begin VB.Label lblAssurPd1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check-in Period #1"
            Height          =   195
            Left            =   1005
            TabIndex        =   14
            Top             =   315
            Width           =   1650
         End
         Begin VB.Label lblAssurStart2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Hour"
            Height          =   195
            Left            =   435
            TabIndex        =   25
            Top             =   1815
            Width           =   885
         End
         Begin VB.Label lblAssurStop2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Hour"
            Height          =   195
            Left            =   510
            TabIndex        =   30
            Top             =   2130
            Width           =   810
         End
         Begin VB.Label lblHrsPrompt3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "( 0 to 23 hr)"
            Height          =   195
            Left            =   2910
            TabIndex        =   28
            Top             =   1815
            Width           =   1020
         End
         Begin VB.Label lblHrsPrompt4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "( 0 to 23 hr)"
            Height          =   195
            Left            =   2910
            TabIndex        =   33
            Top             =   2130
            Width           =   1020
         End
         Begin VB.Label lblEndHr2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2070
            TabIndex        =   32
            Top             =   2130
            Width           =   555
         End
         Begin VB.Label lblStartHr2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2070
            TabIndex        =   27
            Top             =   1815
            Width           =   555
         End
         Begin VB.Label lblStartHr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2070
            TabIndex        =   17
            Top             =   675
            Width           =   555
         End
         Begin VB.Label lblEndHr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12 PM"
            Height          =   195
            Left            =   2070
            TabIndex        =   22
            Top             =   990
            Width           =   555
         End
         Begin VB.Label lblHrsPrompt2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "( 0 to 23 hr)"
            Height          =   195
            Left            =   2910
            TabIndex        =   23
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label lblHrsPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "( 0 to 23 hr)"
            Height          =   195
            Left            =   2910
            TabIndex        =   18
            Top             =   675
            Width           =   1020
         End
         Begin VB.Label lblAssurStop 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Hour"
            Height          =   195
            Left            =   510
            TabIndex        =   20
            Top             =   990
            Width           =   810
         End
         Begin VB.Label lblAssurStart 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Hour"
            Height          =   195
            Left            =   435
            TabIndex        =   15
            Top             =   675
            Width           =   885
         End
      End
      Begin VB.Frame fraSounds 
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
         Height          =   2790
         Left            =   60
         TabIndex        =   34
         Top             =   3330
         Width           =   9285
         Begin VB.CheckBox chkLocaltControl 
            Caption         =   "Local Control"
            Height          =   285
            Left            =   1845
            TabIndex        =   192
            ToolTipText     =   "Beep Timers Controlled by Host  Compuiter"
            Top             =   2460
            Width           =   2505
         End
         Begin VB.TextBox txtAlarmRebeep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   40
            Top             =   480
            Width           =   630
         End
         Begin VB.TextBox txtAlertRebeep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   45
            Top             =   795
            Width           =   630
         End
         Begin VB.TextBox txtBattRebeep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   55
            Top             =   1455
            Width           =   630
         End
         Begin VB.TextBox txtTroubleRebeep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   60
            Top             =   1785
            Width           =   630
         End
         Begin VB.TextBox txtAssurRebeep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   65
            Top             =   2115
            Width           =   630
         End
         Begin VB.TextBox txtExternRebeep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   50
            Top             =   1125
            Width           =   630
         End
         Begin VB.TextBox txtExtBeepTimer 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   49
            Top             =   1140
            Width           =   630
         End
         Begin VB.TextBox txtExtFileName 
            Height          =   300
            Left            =   3855
            TabIndex        =   51
            Top             =   1140
            Width           =   3825
         End
         Begin VB.CommandButton cmdGetExtFilename 
            Height          =   330
            Left            =   7695
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Configure.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1125
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtAssurBeepTimer 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   64
            Top             =   2130
            Width           =   630
         End
         Begin VB.TextBox txtAssurFileName 
            Height          =   300
            Left            =   3855
            TabIndex        =   66
            Top             =   2130
            Width           =   3825
         End
         Begin VB.CommandButton cmdGetAssurFilename 
            Height          =   330
            Left            =   7695
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Configure.frx":052A
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   2115
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdGetTroubleFilename 
            Height          =   330
            Left            =   7695
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Configure.frx":0A54
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1785
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdGetLowBattFilename 
            Height          =   330
            Left            =   7695
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Configure.frx":0F7E
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1455
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdGetAlertFilename 
            Height          =   330
            Left            =   7695
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Configure.frx":14A8
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   795
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdGetAlarmFilename 
            Height          =   330
            Left            =   7695
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Configure.frx":19D2
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   465
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtTroubleFileName 
            Height          =   300
            Left            =   3855
            TabIndex        =   61
            Top             =   1800
            Width           =   3825
         End
         Begin VB.TextBox txtLowBattFileName 
            Height          =   300
            Left            =   3855
            TabIndex        =   56
            Top             =   1470
            Width           =   3825
         End
         Begin VB.TextBox txtAlertFileName 
            Height          =   300
            Left            =   3855
            TabIndex        =   46
            Top             =   810
            Width           =   3825
         End
         Begin VB.TextBox txtAlarmFileName 
            Height          =   300
            Left            =   3855
            TabIndex        =   41
            Top             =   480
            Width           =   3825
         End
         Begin VB.TextBox txtTroubleBeepTimer 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   59
            Top             =   1800
            Width           =   630
         End
         Begin VB.TextBox txtLowBattBeepTimer 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   54
            Top             =   1470
            Width           =   630
         End
         Begin VB.TextBox txtAlertBeepTimer 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   44
            Top             =   810
            Width           =   630
         End
         Begin VB.TextBox txtAlarmBeepTimer 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   39
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lblsec6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Timed by Host)"
            Height          =   195
            Left            =   2370
            TabIndex        =   198
            Top             =   2160
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblsec5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Timed by Host)"
            Height          =   195
            Left            =   2355
            TabIndex        =   197
            Top             =   1836
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblsec4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Timed by Host)"
            Height          =   195
            Left            =   2370
            TabIndex        =   196
            Top             =   1512
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblsec3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Timed by Host)"
            Height          =   195
            Left            =   2370
            TabIndex        =   195
            Top             =   1188
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblsec2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Timed by Host)"
            Height          =   195
            Left            =   2370
            TabIndex        =   194
            Top             =   864
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblsec1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Timed by Host)"
            Height          =   195
            Left            =   2355
            TabIndex        =   193
            Top             =   540
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblRechime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Chime (sec)"
            Height          =   195
            Left            =   2385
            TabIndex        =   36
            Top             =   195
            Width           =   1305
         End
         Begin VB.Label lblExtBeep 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ext Alarm Beep"
            Height          =   195
            Left            =   300
            TabIndex        =   48
            Top             =   1170
            Width           =   1305
         End
         Begin VB.Label lblAssurB 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Assurance Beep"
            Height          =   195
            Left            =   210
            TabIndex        =   63
            Top             =   2160
            Width           =   1395
         End
         Begin VB.Label lblAlarmTimers 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm Timers (sec)"
            Height          =   195
            Left            =   630
            TabIndex        =   35
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label lblBeepSounds 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sound File"
            Height          =   195
            Left            =   3870
            TabIndex        =   37
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lblTroubleB 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trouble Beep"
            Height          =   195
            Left            =   450
            TabIndex        =   58
            Top             =   1830
            Width           =   1155
         End
         Begin VB.Label lblBattB 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Low Battery Beep"
            Height          =   195
            Left            =   90
            TabIndex        =   53
            Top             =   1500
            Width           =   1515
         End
         Begin VB.Label lblAlertB 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alert Beep"
            Height          =   195
            Left            =   705
            TabIndex        =   43
            Top             =   840
            Width           =   900
         End
         Begin VB.Label lblAlarmB 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm Beep"
            Height          =   195
            Left            =   630
            TabIndex        =   38
            Top             =   510
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Exit"
         Height          =   585
         Left            =   8400
         TabIndex        =   185
         Top             =   1935
         Width           =   1175
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   8400
         TabIndex        =   184
         Top             =   660
         Width           =   1175
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   3105
         Left            =   45
         TabIndex        =   1
         Top             =   0
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   5477
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   8
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Facility"
               Key             =   "main"
               Object.ToolTipText     =   "Facility Name"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Sounds"
               Key             =   "sounds"
               Object.ToolTipText     =   "Edit System Sounds"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "System"
               Key             =   "system"
               Object.ToolTipText     =   "System Setup"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Check-ins"
               Key             =   "assur"
               Object.ToolTipText     =   "Set Up Check-ins"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Waypoints"
               Key             =   "waypoints"
               Object.ToolTipText     =   "Set up Locator Waypoints"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Registration"
               Key             =   "registration"
               Object.ToolTipText     =   "Registration and Licensing"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "External Sources"
               Key             =   "external"
               Object.ToolTipText     =   "External Alarm Sources"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Watchdog"
               Key             =   "Watchdog"
               Object.Tag             =   "Watchdog"
               Object.ToolTipText     =   "Watchdog Settings"
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
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mBusy As Boolean

Private Const BEEP_LIMIT = 99999

Private Const FIVE_DIGITS = 5

'5.8.6 0x90 Report Network Coordinator Serial Number
'[0x34]     Network coordinator configuration header
'[LEN]      Length of this message, 0x03
'[0x90]     subcommand to report network coordinator serial number
'[CKSUM]    (0xC7)

'Network coordinator responds with:
'[0x35]     Configuration response header
'[LEN]      Length of message, 0x06
'[0x90]     subcommand to read network coordinator serial number
'[SN]       three byte serial number (network coordinator MID always 00), 3 bytes
'[CKSUM]

Sub Set6080UserPass()
  Dim NewUser As String
  Dim NewPass As String
  
  NewUser = Trim$(txt6080UserName.text)
  NewPass = Trim$(txt6080Password.text)
  
  
'  i6080.DisConnect
'  Set i6080 = Nothing
'  Set i6080 = New c6080
'  i6080.Connect
'
  
  
  
  
  'Save changes, close websocket, reopen
  
End Sub


Private Sub cboAvail_Click()
  Dim index As Long
  index = cboAvail.ListIndex
  If index > -1 Then
    Configuration.SurveyPager = cboAvail.ItemData(index)
  End If
End Sub



Private Sub cboFirst_Click()
  If cboFirst.ListIndex > -1 Then
    lblStartNightHR.Caption = ConvertHourToAMPM(cboFirst.ItemData(cboFirst.ListIndex))
  Else
    lblStartNightHR.Caption = ""
  End If

End Sub

Private Sub cboSecond_Click()
  If cboSecond.ListIndex > -1 Then
    lblEndNightHR.Caption = ConvertHourToAMPM(cboSecond.ItemData(cboSecond.ListIndex))
  Else
    lblEndNightHR.Caption = ""
  End If

End Sub

Private Sub cboThird_Click()
  If cboThird.ListIndex > -1 Then
    lblEndThird.Caption = ConvertHourToAMPM(cboThird.ItemData(cboThird.ListIndex))
  Else
    lblEndThird.Caption = ""
  End If

End Sub

Private Sub cboWaypointMode_Click()
  Configuration.surveymode = Max(PCA_MODE, Min(EN1221_MODE, cboWaypointMode.ListIndex))
  SetWaypointModeElements
End Sub
Sub SetWaypointModeElements()
  txtPCA.Visible = (Configuration.surveymode = PCA_MODE)
  lblPCA.Visible = (Configuration.surveymode = PCA_MODE)
  lblSurveyPager.Visible = (Configuration.surveymode = TWO_BUTTON_MODE Or Configuration.surveymode = EN1221_MODE)
  chkPCARedirect.Visible = True
  cboAvail.Visible = (Configuration.surveymode = TWO_BUTTON_MODE Or Configuration.surveymode = EN1221_MODE)
  
  'If Configuration.surveymode = 1 Then
  '  lblPCA.Caption = "Survey Controller"
  'Else
  '  lblPCA.Caption = "Survey PCA"
  'End If
End Sub

Private Sub chkLocaltControl_Click()
  UpdateScreenElements
End Sub

Private Sub chkLocationPhrase_Click()
  If chkLocationText.Value <> chkLocationPhrase.Value Then
    chkLocationText.Value = chkLocationPhrase.Value
  End If
End Sub

Private Sub chkLocationText_Click()
  If chkLocationText.Value <> chkLocationPhrase.Value Then
    chkLocationPhrase.Value = chkLocationText.Value
  End If

End Sub

Private Sub chkNoNCs_Click()
  Configuration.NoNCs = chkNoNCs.Value
End Sub

Private Sub chkPCARedirect_Click()
  Configuration.PCARedirect = chkPCARedirect.Value
End Sub


Private Sub chkSubscribedAlarms_Click()
  gMyAlarms = IIf(chkSubscribedAlarms.Value = 1, 1, 0)
  
  
End Sub

Private Sub cmd6080_Click()
  If vbYes = messagebox(Me, "This Will Change the System Over to the EN6080 ACG" & vbCrLf & "Continue?", App.Title, vbQuestion Or vbYesNo) Then
    If vbYes = messagebox(Me, "Are You Sure?", App.Title, vbQuestion Or vbYesNo) Then
      ChangeTo6080
    End If
  End If
  Fill
  
End Sub
Sub ChangeTo6080()
  USE6080 = 1
  WriteSetting "Configuration", "USE6080", USE6080
  SaveSettings
  LoadSounds
  SetMainCaption
  UpdateScreenElements
  



End Sub
Private Sub cmdAssurDest_Click()
   ShowAssurSend
End Sub

Private Sub cmdBackupSettings_Click()
  ShowBackupSettings
End Sub

Private Sub cmdChangeHostAdapter_Click()
  ChangeHostAdapter
  Fill
End Sub



Private Sub cmdClose_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdCreateFromRooms_Click()
  CreatePartitonsFromRooms
End Sub

Private Sub cmdDebug_Click()
  ShowDebugScreen
End Sub

Private Sub cmdDeviceTypes_Click()
  ShowTransmitterTypes
End Sub




Private Sub cmdDukane_Click()
  ShowDukane
End Sub

Private Sub cmdEditSoftPoints_Click()
  SaveSettings
  EditSoftPoints
  
End Sub

Private Sub cmdEditUsers_Click()
  ShowUsers
End Sub

Private Sub cmdEmailSetup_Click()
  ShowEmailSettings
End Sub

Private Sub cmdExternalApps_Click()
  ShowOtherPrograms
End Sub

Private Sub cmdFactory_Click()
  ShowFactorySettings
End Sub

Private Sub cmdGetAlarmFilename_Click()
  SaveTimers
  GetWaveFile 1
End Sub

Private Sub cmdGetAlertFilename_Click()
  SaveTimers
  GetWaveFile 2
End Sub

Private Sub cmdGetAssurFilename_Click()
  SaveTimers
  GetWaveFile 5
End Sub

Private Sub cmdGetExtFilename_Click()
  SaveTimers
  GetWaveFile 6

End Sub

Private Sub cmdGetLowBattFilename_Click()
  SaveTimers
  GetWaveFile 3
End Sub

Private Sub cmdGetTroubleFilename_Click()
  SaveTimers
  GetWaveFile 4
End Sub


Private Sub cmdImport_Click()
  SaveSettings
  DoImports
End Sub

Private Sub cmdMobile_Click()
  ShowMobile
  
End Sub

Private Sub cmdMobileSettings_Click()
  ShowMobileSettings
  
End Sub


Private Sub cmdPrintWaypoints_Click()
  Dim rpt As cWaypointReport
  Set rpt = New cWaypointReport
  rpt.PrintReport
  Set rpt = Nothing

End Sub

Private Sub cmdPush_Click()
  ShowPush
End Sub

Private Sub cmdRegister_Click()
  Dim CustomerCode       As String
  Dim RegistrationCode   As String
  CustomerCode = txtMID.text
  RegistrationCode = Trim(txtEID.text)
  Dim Count              As Long
  Count = gSentinel.DoRegistration(CustomerCode, RegistrationCode)

  If MASTER Then
    If Count > 0 Then
      cmdRegister.Caption = "Done"
    Else

      cmdRegister.Caption = "Invalid Code"
    End If
  Else
    If Count = -1 Then
      cmdRegister.Caption = "Done"
    Else

      cmdRegister.Caption = "Invalid Code"
    End If
  End If

  RefreshLicensing
End Sub
Sub RefreshLicensing()
  Call GetLicensing
  
  If MASTER Then
  
  If gRegistered Then
    lblDevices.Caption = "Registered for " & gAllowedDeviceCount & " Devices (" & Devices.Devices.Count & " used)"
  Else
    lblDevices.Caption = "Expires in " & gSentinel.DaysLeft & " Days"
  End If
  
  Else
  
  If gRegistered Then
    lblDevices.Caption = "Registered as Remote Console"
  Else
    lblDevices.Caption = "Expires in " & gSentinel.DaysLeft & " Days"
  End If
  
  End If

End Sub

Private Sub cmdRepairMesh_Click()

End Sub

Private Sub cmdReminders_Click()
  ShowReminders
End Sub

Private Sub cmdRepeaters_Click()
  SaveSettings
  DoRepeaters
End Sub

Private Sub cmdSave_Click()
  
  Dim Key                As String
  On Error Resume Next
  
  Key = TabStrip.SelectedItem.Key
  
  On Error GoTo 0
  
  If StrComp(Key, "waypoints", vbTextCompare) = 0 Then
  
    SaveSettings
  
  ElseIf StrComp(Key, "sounds", vbTextCompare) = 0 Then
    SaveTimers
    LoadSounds
    SetMainCaption
    UpdateScreenElements
  ElseIf StrComp(Key, "assur", vbTextCompare) = 0 Then
    
    UpdateCheckintab
    
    UpdateScreenElements
  Else
    
    SaveSettings
    SaveTimers
    LoadSounds
    SetMainCaption
    UpdateScreenElements

    'start any watchdog
    Select Case Configuration.WatchdogType
      Case WD_BERKSHIRE
        SetWatchdog 0          ' kills UL watchdog
        Set BerkshireWD = New cBerkshire
        BerkshireWD.InitWatchdog
        BerkshireWD.Enable
        BerkshireWD.Tickle
      Case WD_UL
        SetWatchdog 0          ' kills UL watchdog
        Set BerkshireWD = New cBerkshire
        SetWatchdog Configuration.WatchdogTimeout
      Case WD_ARK3510
        SetWatchdog 0          ' kills UL watchdog
        Set BerkshireWD = New cBerkshire
        SetWatchdog -1
      Case Else
        SetWatchdog 0          ' kills UL watchdog
        Set BerkshireWD = New cBerkshire

    End Select
  End If
End Sub

Private Sub cmdScreenMask_Click()
  ShowScreenMask
End Sub

Private Sub cmdSetNID_Click()

  Dim NID           As Long

  If cboNID.ListIndex < 0 Then
    Beep
    Exit Sub
  End If


  If vbYes <> messagebox(Me, "Changing the NID " & vbCrLf & "Requires Resetting ALL Repeaters!" & vbCrLf & "Continue?", App.Title, vbQuestion Or vbYesNo) Then
    Exit Sub
  End If
          
          


  ' to Get NID send 30 03 07 CS
  ' returns typically where 0E is the NID
  ' 31 13 07 0E 30 03 00 1E 00 00 00 00 1E B2 C0 A0 80 90 01 EB
  '          ^^ NID

  ' to SET NID send
  ' 30 04 05 xx CS
  ' NID value must be bewteen 1 and 31 inclusive!


  ' to Sync NIDs send
  ' 20 07 00 nn xx yy zz CS

  NID = Max(0, cboNID.ItemData(cboNID.ListIndex))

  If USE6080 Then
    Dim OldNID      As Long
    OldNID = Get6080NID
    If NID <> OldNID Then
      Set6080NID NID
    End If
    NID = Get6080NID()
    DisplayNID NID
  Else
    If mBusy Then Exit Sub
    mBusy = True
    Screen.MousePointer = vbHourglass
    SetNID NID
    Sleep 200
    Call GetNCNID
    DisplayNID GlobalNID
    mBusy = False
    Screen.MousePointer = vbDefault
  End If

End Sub

Private Sub DisplayNID(ByVal NID As Long)
  Dim j As Long
  For j = cboNID.listcount - 1 To 1 Step -1
    If Val(cboNID.list(j)) = NID Then
      Exit For
    End If
  Next
  cboNID.ListIndex = Max(0, j)
End Sub

Private Sub cmdStressTest_Click()
  frmPhantom.Show
End Sub

Private Sub cmdSyncNIDs_Click()
  
  If mBusy Then Exit Sub
  mBusy = True
  Screen.MousePointer = vbHourglass
  SyncNIDs
  mBusy = False
  Screen.MousePointer = vbDefault
End Sub



Private Sub cmdTestMonitor_Click()
  SaveSettings
  SendTestPing
End Sub



Private Sub cmdUpdate6080UP_Click()
  Dim rc As Long
  rc = messagebox(Me, "This Will Change The Password of the 6080 ACG " & vbCrLf & "Continue?", App.Title, vbQuestion Or vbYesNo)
  If rc = vbYes Then
    Set6080UserPass
  End If
End Sub

Private Sub cmdUpgrade_Click()
  ShowUpgrade
End Sub



Private Sub cmdWayppoints_Click()
  SaveSettings
  LoadSounds
  SetMainCaption
  ShowWaypoints
End Sub




Private Sub cmdPartitions_Click()
 ManageAvailablePartitions
End Sub

Private Sub Command1_Click()
  Download6080
End Sub
Sub Download6080()
  Dim ZoneInfoList As cZoneInfoList
  Dim Zone         As cZoneInfo
  Dim ZoneID       As Long
  
  Dim HTTPRequest   As cHTTPRequest
  Dim rc            As Long

  Set ZoneInfoList = New cZoneInfoList
  Set HTTPRequest = New cHTTPRequest
  Call HTTPRequest.GetZoneList(GetHTTP & "://" & IP1, USER1, PW1)
  Do Until HTTPRequest.Ready
    DoEvents
  Loop
  Select Case HTTPRequest.StatusCode
    Case 200, 201
    Case Else
  End Select
  If Len(HTTPRequest.XML) Then
    rc = ZoneInfoList.LoadXML(HTTPRequest.XML)
  End If
  Set HTTPRequest = Nothing
  Debug.Print
  Debug.Print "ID|Description|DeviceID|HexID|MID|PTI|TypeName|SupervisionWindow|IsFixedDevice|IsLocatable"
  For Each Zone In ZoneInfoList.ZoneList
    Debug.Print Zone.ID & "|" & Zone.Description & "|" & Zone.DeviceID & "|" & Zone.HexID & "|" & Zone.MID & "|" & Zone.PTI & "|" & Zone.TypeName & "|" & Zone.SupervisionWindow & "|" & Zone.IsFixedDevice & "|" & Zone.IsLocatable
    
  Next
  Debug.Print
  
  
  
End Sub

Private Sub Form_Activate()
  UpdateScreenElements

End Sub

Private Sub Form_Click()
'UpdateScreenElements
End Sub

Private Sub Form_DblClick()
  UpdateScreenElements
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Debug.Print "Form_KeyPress FrmConfigure"
End Sub

Private Sub Form_Load()
  ResetActivityTime
  If InIDE Or InStr(1, Command$, "stress", vbTextCompare) Then
    cmdStressTest.Visible = True And MASTER
  Else
    cmdStressTest.Visible = False
  End If
  GetConfig
  
  LoadCombos
  
  lblCompany.Caption = COMPANY_NAME
  lblPhone.Caption = COMPANY_800
  lblFax.Caption = COMPANY_FAX
  lblMailTo.Caption = COMPANY_EMAIL
  
  SetControls
  ShowPanel "main"
  Fill
  UpdateScreenElements
End Sub

Sub LoadCombos()
  Dim j As Long
  cboFirst.Clear
  cboSecond.Clear
  cboThird.Clear
  cboNID.Clear
  For j = 0 To 31
    AddToCombo cboNID, j, j
  Next
  cboNID.ListIndex = 0
  
 
  'cboSecond.AddItem "<None>"
  'cboSecond.ItemData(cboSecond.NewIndex) = -1
  
 ' cboThird.AddItem "<None>"
 ' cboThird.ItemData(cboThird.NewIndex) = -1
  
  
  For j = 0 To 23
      cboFirst.AddItem j
      cboFirst.ItemData(cboFirst.NewIndex) = j
      cboSecond.AddItem j
      cboSecond.ItemData(cboSecond.NewIndex) = j
      cboThird.AddItem j
      cboThird.ItemData(cboThird.NewIndex) = j
  Next
  
  cboWDType.Clear
  cboWDType.AddItem "<none>"
  cboWDType.ItemData(cboWDType.NewIndex) = WD_NONE
  cboWDType.AddItem "WindBond"
  cboWDType.ItemData(cboWDType.NewIndex) = WD_UL
  cboWDType.AddItem "Bershire"
  cboWDType.ItemData(cboWDType.NewIndex) = WD_BERKSHIRE
  cboWDType.AddItem "ARK 3510"
  cboWDType.ItemData(cboWDType.NewIndex) = WD_ARK3510
  
  cboWDType.ListIndex = 0
  
  
  
End Sub

Sub FillAvailablePagers()
  Dim rs As Recordset
  Dim j As Integer
  
  cboAvail.Clear
  AddToCombo cboAvail, "[Select Pager]", 0
  Set rs = ConnExecute("SELECT * FROM Pagers")
  Do Until rs.EOF
    AddToCombo cboAvail, rs("Description") & "", rs("pagerid")
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  For j = cboAvail.listcount - 1 To 1 Step -1
    If cboAvail.ItemData(j) = Configuration.SurveyPager Then
      Exit For
    End If
  Next
  cboAvail.ListIndex = j
End Sub


Sub SetControls()
  Dim f As Control
  Dim PanelWidth As Double

  For Each f In Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor
    End If
  Next

  fraEnabler.BackColor = Me.BackColor

  fraEnabler.left = 0
  fraEnabler.top = 0
  Debug.Print " fraEnabler.width  "; fraEnabler.Width
  fraEnabler.Width = 9735
  fraEnabler.Height = ENABLER_HEIGHT
  

  TabStrip.left = 0
  TabStrip.top = 0
  TabStrip.Width = fraEnabler.Width
  TabStrip.Height = fraEnabler.Height

  PanelWidth = TabStrip.ClientWidth - 1400

  cmdSave.left = TabStrip.ClientWidth - cmdSave.Width

  cmdClose.left = cmdSave.left
  cmdClose.top = TabStrip.ClientHeight + TabStrip.ClientTop - (cmdClose.Height + 60)

  If Not MASTER Then
    TabStrip.Tabs.Remove ("system")
    TabStrip.Tabs.Remove ("assur")
    TabStrip.Tabs.Remove ("waypoints")
  End If
  
  framain.left = TabStrip.ClientLeft
  framain.top = TabStrip.ClientTop
  framain.Width = PanelWidth
  framain.Height = TabStrip.ClientHeight

  fraSounds.BackColor = Me.BackColor
  fraSounds.left = TabStrip.ClientLeft
  fraSounds.top = TabStrip.ClientTop
  fraSounds.Width = PanelWidth
  fraSounds.Height = TabStrip.ClientHeight

  fraSystem.BackColor = Me.BackColor
  fraSystem.left = TabStrip.ClientLeft
  fraSystem.top = TabStrip.ClientTop
  fraSystem.Width = PanelWidth
  fraSystem.Height = TabStrip.ClientHeight

  fraAssur.BackColor = Me.BackColor
  fraAssur.left = TabStrip.ClientLeft
  fraAssur.top = TabStrip.ClientTop
  fraAssur.Width = PanelWidth
  fraAssur.Height = TabStrip.ClientHeight

 
  fraExternal.BackColor = Me.BackColor
  fraExternal.left = TabStrip.ClientLeft
  fraExternal.top = TabStrip.ClientTop
  fraExternal.Width = PanelWidth
  fraExternal.Height = TabStrip.ClientHeight
  
  fraSoftPoints.BackColor = Me.BackColor
  fraSoftPoints.left = TabStrip.ClientLeft
  fraSoftPoints.top = TabStrip.ClientTop
  fraSoftPoints.Width = PanelWidth
  fraSoftPoints.Height = TabStrip.ClientHeight
  
  
  frawaypts.BackColor = Me.BackColor
  frawaypts.left = TabStrip.ClientLeft
  frawaypts.top = TabStrip.ClientTop
  frawaypts.Width = PanelWidth
  frawaypts.Height = TabStrip.ClientHeight

  fraRegistration.BackColor = Me.BackColor
  fraRegistration.left = TabStrip.ClientLeft
  fraRegistration.top = TabStrip.ClientTop
  fraRegistration.Width = PanelWidth
  fraRegistration.Height = TabStrip.ClientHeight

  fraWatchdog.BackColor = Me.BackColor
  fraWatchdog.left = TabStrip.ClientLeft
  fraWatchdog.top = TabStrip.ClientTop
  fraWatchdog.Width = PanelWidth
  fraWatchdog.Height = TabStrip.ClientHeight


  chkTimeFormat.BackColor = Me.BackColor
  
  txtAppPath.BackColor = Me.BackColor
  txtAppPath.text = "App Path: " & App.Path & "\" & App.exename

  cboWaypointMode.AddItem "PCA"
  cboWaypointMode.AddItem "Two Button w/ Pager"
  cboWaypointMode.AddItem "EN1221S Pendant"
  
  
  Configuration.surveymode = Max(PCA_MODE, Min(EN1221_MODE, Configuration.surveymode)) '  Or Configuration.surveymode = EN1221_MODE
  
  cboWaypointMode.ListIndex = Configuration.surveymode
  SetWaypointModeElements
  chkUseOnlyLocators.Visible = gDirectedNetwork
  
  cmdDukane.Visible = ((MASTER) And (gUser.LEvel >= LEVEL_ADMIN))
  
  
End Sub


Sub SetReBeepVisible(ByVal Visible As Boolean)



  lblRechime.Visible = Visible
  txtAlarmRebeep.Visible = Visible
  txtAlertRebeep.Visible = Visible
  txtAssurRebeep.Visible = False
  txtExternRebeep.Visible = Visible
  txtBattRebeep.Visible = Visible
  txtTroubleRebeep.Visible = Visible

End Sub

Private Sub UpdateScreenElements()
10      cmdImport.Visible = False

20      chkDirectedNet.Enabled = False

30      AddRemoveTabs

        cmdFactory.Visible = False

40      Select Case gUser.LEvel

          Case LEVEL_FACTORY         ' Factory

            cmdFactory.Visible = True
50          cmdDeviceTypes.Visible = True
60          cmdEditUsers.Visible = True
70          txtFacility.Locked = False
80          txtID.Locked = True
90          txtCommPort.Locked = False
100         txtCommTimeout.Locked = False
110         txtAssurStart.Locked = False
120         txtAssurStart2.Locked = False
130         txtAssurEnd.Locked = False
140         txtAssurEnd2.Locked = False
150         txtHostPort.Visible = True

160         lblCommPort.Visible = True And MASTER And USE6080 = 0
170         txtCommPort.Visible = True And MASTER And USE6080 = 0
180         lblrxTimout.Visible = True And MASTER And USE6080 = 0
190         txtCommTimeout.Visible = True And MASTER And USE6080 = 0
200         lblRxMinutes.Visible = True And MASTER And USE6080 = 0
210         lblRxSerial.Visible = True And MASTER
220         txtRxSerial.Visible = True And MASTER
230         lblRxLocation.Visible = True And MASTER
240         txtRxLocation.Visible = True And MASTER
250         cmdImport.Visible = True And MASTER
260         fraNetwork.Visible = True And MASTER
270         cmdExternalApps.Visible = True

280         txtIP.Locked = MASTER

290         chkDirectedNet.Visible = True And MASTER And (USE6080 = 0)
300         chkDirectedNet.Enabled = True And MASTER

310         txt6080Password.Visible = True And USE6080 And MASTER
320         txt6080UserName.Visible = True And USE6080 And MASTER

330         lblACGUP.Visible = True And USE6080 And MASTER
340         cmdUpdate6080UP.Visible = True And USE6080 And MASTER

            'cmdRepairMesh.Visible = True And MASTER

350       Case LEVEL_ADMIN           ' Admin 2
360         cmdEditUsers.Visible = True
370         cmdDeviceTypes.Visible = False
380         txtFacility.Locked = False
390         txtID.Locked = True
400         txtCommPort.Locked = False
410         txtCommTimeout.Locked = False
420         txtAssurStart.Locked = False
430         txtAssurStart2.Locked = False
440         txtAssurEnd.Locked = False
450         txtAssurEnd2.Locked = False
460         txtIP.Locked = MASTER

470         lblCommPort.Visible = True And MASTER And USE6080 = 0
480         txtCommPort.Visible = True And MASTER And USE6080 = 0
490         lblrxTimout.Visible = True And MASTER And USE6080 = 0
500         txtCommTimeout.Visible = True And MASTER And USE6080 = 0
510         lblRxMinutes.Visible = True And MASTER And USE6080 = 0
520         lblRxSerial.Visible = True And MASTER
530         txtRxSerial.Visible = True And MASTER
540         lblRxLocation.Visible = True And MASTER
550         txtRxLocation.Visible = True And MASTER
560         cmdImport.Visible = True And MASTER
570         fraNetwork.Visible = True And MASTER

580         cmdExternalApps.Visible = True
590         chkDirectedNet.Visible = True And USE6080 = 0
600         chkDirectedNet.Enabled = False

610         txt6080Password.Visible = True And USE6080 And MASTER
620         txt6080UserName.Visible = True And USE6080 And MASTER

630         lblACGUP.Visible = True And USE6080 And MASTER
640         cmdUpdate6080UP.Visible = False And MASTER

            'cmdRepairMesh.Visible = True And MASTER

650       Case LEVEL_SUPERVISOR      ' Admin 1
660         cmdEditUsers.Visible = False
670         cmdDeviceTypes.Visible = False
680         txtFacility.Locked = True
690         txtID.Locked = True
700         txtCommPort.Locked = True
710         txtCommTimeout.Locked = True
720         txtAssurStart.Locked = True
730         txtAssurStart2.Locked = True
740         txtAssurEnd.Locked = True
750         txtAssurEnd2.Locked = True
760         txtIP.Locked = True

770         lblACGUP.Visible = False
780         txt6080Password.Visible = False
790         txt6080UserName.Visible = False
800         lblACGUP.Visible = False
810         cmdUpdate6080UP.Visible = False


820         lblCommPort.Visible = True And MASTER And USE6080 = 0
830         txtCommPort.Visible = True And MASTER And USE6080 = 0
840         lblrxTimout.Visible = True And MASTER And USE6080 = 0
850         txtCommTimeout.Visible = True And MASTER And USE6080 = 0
860         lblRxMinutes.Visible = True And MASTER And USE6080 = 0
870         lblRxSerial.Visible = True And MASTER
880         txtRxSerial.Visible = True And MASTER
890         lblRxLocation.Visible = True And MASTER
900         txtRxLocation.Visible = True And MASTER

910         cmdImport.Visible = False
920         fraNetwork.Visible = False

930         cmdExternalApps.Visible = True

940         chkDirectedNet.Visible = True And MASTER And USE6080 = 0
950         chkDirectedNet.Enabled = False

960         lblACGUP.Visible = True And USE6080
970         cmdUpdate6080UP.Visible = False


980       Case Else                  ' General users
990         cmdEditUsers.Visible = False
1000        cmdDeviceTypes.Visible = False
1010        txtFacility.Locked = True
1020        txtID.Locked = True
1030        txtCommPort.Locked = True
1040        txtCommTimeout.Locked = True
1050        txtAssurStart.Locked = True
1060        txtAssurStart2.Locked = True
1070        txtAssurEnd.Locked = True
1080        txtAssurEnd2.Locked = True
1090        txtIP.Locked = True
1100        fraNetwork.Visible = False
1110        lblCommPort.Visible = True And MASTER And USE6080 = 0
1120        txtCommPort.Visible = True And MASTER And USE6080 = 0
1130        lblrxTimout.Visible = True And MASTER And USE6080 = 0
1140        txtCommTimeout.Visible = True And MASTER And USE6080 = 0
1150        lblRxMinutes.Visible = True And MASTER And USE6080 = 0
1160        lblRxSerial.Visible = True And MASTER
1170        txtRxSerial.Visible = True And MASTER
1180        lblRxLocation.Visible = True And MASTER
1190        txtRxLocation.Visible = True And MASTER

1200        cmdExternalApps.Visible = False

1210        chkDirectedNet.Visible = True And MASTER And USE6080 = 0
1220        chkDirectedNet.Enabled = False

1230        txt6080Password.Visible = False
1240        txt6080UserName.Visible = False
1250        lblACGUP.Visible = False
1260        cmdUpdate6080UP.Visible = False

1270    End Select



11640        lblsec2.left = lblsec1.left
11650        lblsec3.left = lblsec1.left
11660        lblsec4.left = lblsec1.left
11670        lblsec5.left = lblsec1.left
11680        lblsec6.left = lblsec1.left


1280    If MASTER Then               ' BEEPTIMERS



1290      txtRemoteSerial.Visible = False
1300      lblRemoteSerial.Visible = False

1310      txtAlarmBeepTimer.Visible = True
1320      txtAlertBeepTimer.Visible = True
1330      txtAssurBeepTimer.Visible = True
1340      txtExtBeepTimer.Visible = True
1350      txtLowBattBeepTimer.Visible = True
1360      txtTroubleBeepTimer.Visible = True

1370      If chkLocaltControl.Value And 1 Then  '
1380        SetReBeepVisible True
1390      Else
1400        SetReBeepVisible False
1410      End If

1420      chkSubscribedAlarms.Visible = False

1430      lblsec1.Visible = False
1440      lblsec2.Visible = False
1450      lblsec3.Visible = False
1460      lblsec4.Visible = False
1470      lblsec5.Visible = False
1480      lblsec6.Visible = False

1490    Else                         ' If MASTER Then

1500      txtRemoteSerial.Visible = True
1510      lblRemoteSerial.Visible = True

1520      chkSubscribedAlarms.Visible = True

1530      If chkLocaltControl.Value And 1 Then  ' chkLocaltControl.Value And 1 Then

1540        SetReBeepVisible True

  txtAlarmBeepTimer.Visible = True
  txtAlertBeepTimer.Visible = True
  'txtAssurBeepTimer.Visible = True
  txtExtBeepTimer.Visible = True
  txtLowBattBeepTimer.Visible = True
  txtTroubleBeepTimer.Visible = True


1550        lblsec1.Visible = False
1560        lblsec2.Visible = False
1570        lblsec3.Visible = False
1580        lblsec4.Visible = False
1590        lblsec5.Visible = False
1600        lblsec6.Visible = True

1610      Else
1620        SetReBeepVisible False

  txtAlarmBeepTimer.Visible = False
  txtAlertBeepTimer.Visible = False
  txtAssurBeepTimer.Visible = False
  txtExtBeepTimer.Visible = False
  txtLowBattBeepTimer.Visible = False
  txtTroubleBeepTimer.Visible = False


1630        lblsec1.Visible = True
1640        lblsec2.Visible = True
1650        lblsec3.Visible = True
1660        lblsec4.Visible = True
1670        lblsec5.Visible = True
1680        lblsec6.Visible = True

1690      End If


1700    End If                       'If MASTER Then

        '    txtAlarmBeepTimer.Visible = False
        '    txtAlertBeepTimer.Visible = False
        '    txtAssurBeepTimer.Visible = False
        '    txtExtBeepTimer.Visible = False
        '    txtLowBattBeepTimer.Visible = False
        '    txtTroubleBeepTimer.Visible = False
        '
        '
        '    txtAlarmRebeep.Visible = False
        '    txtAlertRebeep.Visible = False
        '    txtAssurRebeep.Visible = False
        '    txtExternRebeep.Visible = False
        '    txtBattRebeep.Visible = False
        '    txtTroubleRebeep.Visible = False
        '

        '  End If






End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  Do While mBusy
'    DoEvents
'  Loop
End Sub

Private Sub fraRemoteConsole_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub lblMailTo_Click()
'do mailto:

  HyperLink "mailto:" & lblMailTo.Caption, "Subject=Support"

End Sub

Private Sub TabStrip_Click()
  AddRemoveTabs
  ShowPanel TabStrip.SelectedItem.Key

End Sub
Sub AddRemoveTabs()

  Dim t             As Object
  Dim haswatchdogtab As Boolean
  Dim Key As String
  Dim j As Long
  
'  For j = 1 To TabStrip.Tabs.Count
'    Set t = TabStrip.Tabs(j)
'  Next
  


  For Each t In TabStrip.Tabs
    'Debug.Print "Key Caption " & t.Key, t.Caption
    
    If 0 = StrComp(t.Key, "waypoints", vbTextCompare) Then
      
      't.Key = "watchdog"
      Key = t.Key
      'haswaypointstab = True
      If USE6080 Then
        t.Caption = "Location"
       Else
        t.Caption = "Waypoints"
      End If
      'Exit For
    End If
  Next
  
  
  
  For Each t In TabStrip.Tabs
    If 0 = StrComp(t.Key, "watchdog", vbTextCompare) Then
      haswatchdogtab = True
    End If
  Next

  If haswatchdogtab Then
    If (gUser.LEvel < LEVEL_FACTORY) Then
      If Not MASTER Then
        Call TabStrip.Tabs.Remove(TabStrip.Tabs.Count)
      End If
    End If
  Else
    If gUser.LEvel = LEVEL_FACTORY And MASTER Then
      TabStrip.Tabs.Add , "watchdog", "Watchdog"
    End If
  End If
  TabStrip.Refresh
End Sub

Sub ShowPanel(ByVal Key As String)

  TimerWD.Enabled = False

  Select Case LCase(Key)
  


    Case "watchdog"
      If gUser.LEvel = LEVEL_FACTORY Then
        TimerWD.Enabled = True
        fraWatchdog.Visible = True
        fraExternal.Visible = False
        fraAssur.Visible = False
        fraSystem.Visible = False
        framain.Visible = False

        fraSounds.Visible = False
        frawaypts.Visible = False
        fraRegistration.Visible = False
        fraSoftPoints.Visible = False
      Else
        Beep
      End If

    Case "external"
      fraExternal.Visible = True
      fraAssur.Visible = False
      fraSystem.Visible = False
      framain.Visible = False

      fraSounds.Visible = False
      frawaypts.Visible = False
      fraRegistration.Visible = False
      fraWatchdog.Visible = False
      fraSoftPoints.Visible = False

    Case "assur"
      fraAssur.Visible = True
      fraSystem.Visible = False
      framain.Visible = False

      fraSounds.Visible = False
      frawaypts.Visible = False
      fraRegistration.Visible = False
      fraExternal.Visible = False
      fraWatchdog.Visible = False
      fraSoftPoints.Visible = False
    Case "sounds"
      fraSounds.Visible = True
      fraAssur.Visible = False
      fraSystem.Visible = False
      framain.Visible = False

      frawaypts.Visible = False
      fraRegistration.Visible = False
      fraExternal.Visible = False
      fraWatchdog.Visible = False
      fraSoftPoints.Visible = False
    Case "system"
      fraSystem.Visible = True
      fraAssur.Visible = False
      framain.Visible = False

      fraSounds.Visible = False
      frawaypts.Visible = False
      fraRegistration.Visible = False
      fraExternal.Visible = False
      fraWatchdog.Visible = False
      fraSoftPoints.Visible = False
    Case "waypoints"

      If USE6080 Then
        fraSoftPoints.Visible = True
        frawaypts.Visible = False
        fraAssur.Visible = False
        framain.Visible = False

        fraSounds.Visible = False
        fraSystem.Visible = False
        fraRegistration.Visible = False
        fraExternal.Visible = False
        fraWatchdog.Visible = False
      Else


        frawaypts.Visible = True
        fraSoftPoints.Visible = False
        fraAssur.Visible = False
        framain.Visible = False

        fraSounds.Visible = False
        fraSystem.Visible = False
        fraRegistration.Visible = False
        fraExternal.Visible = False
        fraWatchdog.Visible = False
      End If
    Case "registration"

      fraRegistration.Visible = True

      framain.Visible = False

      fraAssur.Visible = False
      fraSystem.Visible = False
      fraSounds.Visible = False
      frawaypts.Visible = False
      fraSoftPoints.Visible = False
      txtMID.text = gSentinel.MachineID
      txtEID.text = ""
      fraExternal.Visible = False
      fraWatchdog.Visible = False
      RefreshLicensing

    Case Else
      framain.Visible = True
      If USE6080 Then
        cmdPartitions.Visible = True
        cmdUpgrade.Visible = True
        lblACGIP.Visible = True And MASTER
        txtAGPIP.Visible = True And MASTER
        cmdRepeaters.Visible = False
        cmdSyncNIDs.Visible = False
        chkDirectedNet.Visible = False
        lblCommPort.Visible = False
        txtCommPort.Visible = False
        fraSoftPoints.Visible = False
      Else

        cmdRepeaters.Visible = True
        cmdSyncNIDs.Visible = True
        chkDirectedNet.Visible = True And USE6080 = 0
        lblCommPort.Visible = True And USE6080 = 0
        txtCommPort.Visible = True And USE6080 = 0

        cmdPartitions.Visible = False
        cmdUpgrade.Visible = False
        lblACGIP.Visible = False
        txtAGPIP.Visible = False
        fraSoftPoints.Visible = False
      End If

      fraAssur.Visible = False
      fraSystem.Visible = False
      fraSounds.Visible = False
      frawaypts.Visible = False
      fraRegistration.Visible = False
      fraExternal.Visible = False
      fraWatchdog.Visible = False
      fraSoftPoints.Visible = False
  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  UnHost
End Sub


Private Sub Text1_Change()

End Sub

Private Sub TimerWD_Timer()
  On Error Resume Next
  Select Case Configuration.WatchdogType
    Case WD_BERKSHIRE
    
      lblStatus.Caption = "Countdown: " & BerkshireWD.RemainingTime
      If BerkshireWD.RemainingTime < 10 Then
        BerkshireWD.Tickle
      End If
    Case WD_ARK3510
    
      lblStatus.Caption = IIf(ARK3510_started, "Started", "Stopped")
      
    Case WD_UL
    
      inportb &H2F
      If Err.Number = 0 Then
        lblStatus.Caption = "OK"
      Else
        lblStatus.Caption = "No DLL"
      End If
    Case Else
      lblStatus.Caption = ""
  End Select
End Sub

Private Sub txt6080Password_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case vbKeyReturn
  Case Else
    txt6080Password.tag = "1"
  End Select
End Sub

Private Sub txtAlarmBeepTimer_GotFocus()
 SelAll txtAlarmBeepTimer
End Sub

Private Sub txtAlarmBeepTimer_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAlarmBeepTimer.text) + 1
      txtAlarmBeepTimer.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAlarmBeepTimer.text) - 1
      txtAlarmBeepTimer.text = Max(newval, -1)
    Case Else

      KeyAscii = KeyProcMax(txtAlarmBeepTimer, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Private Sub txtAlarmBeepTimer_LostFocus()
  txtAlarmBeepTimer.text = Max(-1, Min(Val(txtAlarmBeepTimer.text), BEEP_LIMIT))
  SaveTimers
End Sub

Private Sub txtAlarmFileName_DblClick()

  Dim Temp() As Byte
  On Error Resume Next
  Temp = GetWaveData(txtAlarmFileName.text)
  PlayMemSound Temp(0), Win32.SND_MEMORY

End Sub

Private Sub txtAlarmFileName_GotFocus()
  SelAll txtAlarmFileName
End Sub

Private Sub txtAlarmRebeep_GotFocus()
  SelAll txtAlarmRebeep
End Sub

Private Sub txtAlarmRebeep_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAlarmRebeep.text) + 1
      txtAlarmRebeep.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAlarmRebeep.text) - 1
      txtAlarmRebeep.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtAlarmRebeep, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Private Sub txtAlertBeepTimer_GotFocus()
  SelAll txtAlertBeepTimer
End Sub

Private Sub txtAlertBeepTimer_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAlertBeepTimer.text) + 1
      txtAlertBeepTimer.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAlertBeepTimer.text) - 1
      txtAlertBeepTimer.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtAlertBeepTimer, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Private Sub txtAlertBeepTimer_LostFocus()
  txtAlertBeepTimer.text = Max(-1, Min(Val(txtAlertBeepTimer.text), BEEP_LIMIT))
End Sub

Private Sub txtAlertFileName_DblClick()
  Dim Temp() As Byte
  On Error Resume Next
  Temp = GetWaveData(txtAlertFileName.text)
  PlayMemSound Temp(0), Win32.SND_MEMORY

End Sub

Private Sub txtAlertFileName_GotFocus()
  SelAll txtAlertFileName
End Sub

Private Sub txtAlertRebeep_GotFocus()
SelAll txtAlertRebeep
End Sub

Private Sub txtAlertRebeep_KeyPress(KeyAscii As Integer)
Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAlertRebeep.text) + 1
      txtAlertRebeep.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAlertRebeep.text) - 1
      txtAlertRebeep.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtAlertRebeep, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Private Sub txtAssurBeepTimer_GotFocus()
  SelAll txtAssurBeepTimer
End Sub

Private Sub txtAssurBeepTimer_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAssurBeepTimer.text) + 1
      txtAssurBeepTimer.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAssurBeepTimer.text) - 1
      txtAssurBeepTimer.text = Max(newval, -1)
    Case Else

      KeyAscii = KeyProcMax(txtAssurBeepTimer, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select

End Sub

Private Sub txtAssurBeepTimer_LostFocus()
  txtAssurBeepTimer.text = Max(-1, Min(Val(txtAssurBeepTimer.text), BEEP_LIMIT))
End Sub

Private Sub txtAssurEnd_Change()
  lblEndHr.Caption = ConvertHourToAMPM(Val(txtAssurEnd.text))
End Sub

Private Sub txtAssurEnd_GotFocus()
  SelAll txtAssurEnd
End Sub

Private Sub txtAssurEnd_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAssurEnd.text) + 1
      txtAssurEnd.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAssurEnd.text) - 1
      txtAssurEnd.text = Max(newval, 0)

    Case Else

      KeyAscii = KeyProcMax(txtAssurEnd, KeyAscii, False, 0, 2, 23)
  End Select
End Sub

Private Sub txtAssurEnd2_Change()
  lblEndHr2.Caption = ConvertHourToAMPM(Val(txtAssurEnd2.text))
End Sub

Private Sub txtAssurEnd2_GotFocus()
  SelAll txtAssurEnd2
End Sub

Private Sub txtAssurEnd2_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAssurEnd2.text) + 1
      txtAssurEnd2.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAssurEnd2.text) - 1
      txtAssurEnd2.text = Max(newval, 0)

    Case Else

      KeyAscii = KeyProcMax(txtAssurEnd2, KeyAscii, False, 0, 2, 23)
  End Select
End Sub

Private Sub txtAssurFileName_DblClick()
  Dim Temp() As Byte
  On Error Resume Next
  Temp = GetWaveData(txtAssurFileName.text)
  PlayMemSound Temp(0), Win32.SND_MEMORY
End Sub

Private Sub txtAssurFileName_GotFocus()
  SelAll txtAssurFileName
End Sub

Private Sub txtAssurRebeep_GotFocus()
  SelAll txtAssurRebeep
End Sub

Private Sub txtAssurRebeep_KeyPress(KeyAscii As Integer)
Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAssurRebeep.text) + 1
      txtAssurRebeep.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAssurRebeep.text) - 1
      txtAssurRebeep.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtAssurRebeep, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Private Sub txtAssurStart_Change()
  lblStartHr.Caption = ConvertHourToAMPM(Val(txtAssurStart.text))
End Sub

Private Sub txtAssurStart_GotFocus()
  SelAll txtAssurStart
End Sub

Private Sub txtAssurStart_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAssurStart.text) + 1
      txtAssurStart.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAssurStart.text) - 1
      txtAssurStart.text = Max(newval, 0)

    Case Else
      KeyAscii = KeyProcMax(txtAssurStart, KeyAscii, False, 0, 2, 23)
  End Select


End Sub

Private Sub txtAssurStart2_Change()
  lblStartHr2.Caption = ConvertHourToAMPM(Val(txtAssurStart2.text))

End Sub

Private Sub txtAssurStart2_GotFocus()
  SelAll txtAssurStart2
End Sub

Private Sub txtAssurStart2_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtAssurStart2.text) + 1
      txtAssurStart2.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtAssurStart2.text) - 1
      txtAssurStart2.text = Max(newval, 0)

    Case Else

      KeyAscii = KeyProcMax(txtAssurStart2, KeyAscii, False, 0, 2, 23)
  End Select
End Sub

Private Sub txtBattRebeep_GotFocus()
  SelAll txtBattRebeep
End Sub

Private Sub txtBattRebeep_KeyPress(KeyAscii As Integer)
Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtBattRebeep.text) + 1
      txtBattRebeep.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtBattRebeep.text) - 1
      txtBattRebeep.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtBattRebeep, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select

End Sub

Private Sub txtBoost_GotFocus()
  SelAll txtBoost
  
End Sub

Private Sub txtBoost_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtBoost.text) + 1
      txtBoost.text = Max(-1, Min(newval, BOOST_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtBoost.text) - 1
      txtBoost.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtBoost, KeyAscii, False, 0, 2, BOOST_LIMIT)
  End Select


End Sub

Private Sub txtCommPort_GotFocus()
  SelAll txtCommPort
End Sub

Private Sub txtCommTimeout_GotFocus()
  SelAll txtCommTimeout
End Sub

Private Sub txtEID_Change()
  cmdRegister.Caption = "Register"
End Sub

'Private Sub txtEndNight_Change()
'  Dim t As Integer
'  t = Val(txtEndNight.text)
'  If t < 12 Then
'    If t = 0 Then
'      lblEndNightHR.Caption = "MidNight"
'    Else
'      lblEndNightHR.Caption = t & " AM"
'    End If
'  ElseIf t = 12 Then
'    lblEndNightHR.Caption = t & " PM"
'  Else
'    lblEndNightHR.Caption = t - 12 & " PM"
'  End If
'
'End Sub
'
'Private Sub txtEndNight_GotFocus()
'  SelAll txtEndNight
'End Sub

'Private Sub txtEndNight_KeyPress(KeyAscii As Integer)
'  Dim newval As Integer
'  Select Case KeyAscii
'    Case vbKeyAdd, 43
'      KeyAscii = 0
'      newval = Val(txtEndNight.text) + 1
'      txtEndNight.text = Min(newval, 23)
'    Case vbKeySubtract, 45
'      KeyAscii = 0
'      newval = Val(txtEndNight.text) - 1
'      txtEndNight.text = Max(newval, 0)
'
'    Case Else
'      KeyAscii = KeyProcMax(txtEndNight, KeyAscii, False, 0, 2, 23)
'  End Select
'
'End Sub

'Private Sub txtEscalate_GotFocus()
'  SelAll txtEscalate
'End Sub

'Private Sub txtEscalate_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtEscalate, KeyAscii, False, 0, 4, 9999)
'End Sub

Private Sub txtExtBeepTimer_GotFocus()
  SelAll txtExtBeepTimer
End Sub

Private Sub txtExtBeepTimer_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtExtBeepTimer.text) + 1
      txtExtBeepTimer.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtExtBeepTimer.text) - 1
      txtExtBeepTimer.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtExtBeepTimer, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select

End Sub

Private Sub txtExtBeepTimer_LostFocus()
  txtExtBeepTimer.text = Max(-1, Min(Val(txtExtBeepTimer.text), BEEP_LIMIT))
  SaveTimers

End Sub

Private Sub txtExternRebeep_GotFocus()
SelAll txtExternRebeep
End Sub

Private Sub txtExternRebeep_KeyPress(KeyAscii As Integer)
Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtExternRebeep.text) + 1
      txtExternRebeep.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtExternRebeep.text) - 1
      txtExternRebeep.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtExternRebeep, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select

End Sub

Private Sub txtExtFileName_DblClick()
  Dim Temp() As Byte
  On Error Resume Next
  Temp = GetWaveData(txtExtFileName.text)
  PlayMemSound Temp(0), Win32.SND_MEMORY

End Sub

Private Sub txtExtFileName_GotFocus()
  SelAll txtExtFileName
End Sub

Private Sub txtFacility_GotFocus()
  SelAll txtFacility
End Sub

Private Sub txtHostPort_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtHostPort, KeyAscii, False, 0, 5, 65000)
End Sub

Private Sub txtHostPort_LostFocus()
  txtHostPort.text = GetValidRemotePort(Val(txtHostPort.text), 2500)
End Sub

Private Sub txtID_GotFocus()
  SelAll txtID
End Sub

Private Sub txtIP_GotFocus()
  SelAll txtIP
End Sub

Private Sub txtIP_LostFocus()
  txtIP.text = GetValidIP(txtIP.text, "127.0.0.1")
End Sub

Private Sub txtLowBattBeepTimer_GotFocus()
  SelAll txtLowBattBeepTimer
End Sub

Private Sub txtLowBattBeepTimer_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtLowBattBeepTimer.text) + 1
      txtLowBattBeepTimer.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtLowBattBeepTimer.text) - 1
      txtLowBattBeepTimer.text = Max(newval, -1)
    Case Else



      KeyAscii = KeyProcMax(txtLowBattBeepTimer, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Private Sub txtLowBattBeepTimer_LostFocus()
  txtLowBattBeepTimer.text = Max(-1, Min(Val(txtLowBattBeepTimer.text), BEEP_LIMIT))
End Sub

Private Sub txtLowBattFileName_DblClick()
  Dim Temp() As Byte
  On Error Resume Next
  Temp = GetWaveData(txtLowBattFileName.text)
  PlayMemSound Temp(0), Win32.SND_MEMORY
End Sub

Private Sub txtLowBattFileName_GotFocus()
  SelAll txtLowBattFileName
End Sub

Private Sub txtMonitorInterval_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMonitorInterval, KeyAscii, False, 0, 4, 9999)
End Sub

Private Sub txtMonitorPort_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMonitorPort, KeyAscii, False, 0, 5, 65000)
End Sub

Private Sub txtRemoteSerial_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select

End Sub

'Private Sub txtNID_Change()
'  Dim Value As Integer
'  Value = Val(txtNID.text)
'  cmdSetNID.Enabled = Value >= 0 And Value < 32
'  cmdSyncNIDs.Enabled = Value >= 0 And Value < 32
'End Sub
'
'Private Sub txtNID_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtNID, KeyAscii, False, 0, 2, 31)
'End Sub

Private Sub txtPCA_GotFocus()
  SelAll txtPCA
End Sub

Private Sub txtPCA_KeyPress(KeyAscii As Integer)
' handle hex data only
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
  'KeyAscii = KeyProcHex(txtSerial, KeyAscii, False, 0, 8)

End Sub


Private Sub txtRxLocation_GotFocus()
  SelAll txtRxLocation
End Sub

Private Sub txtRxSerial_GotFocus()
  SelAll txtRxSerial
End Sub

Private Sub txtRxSerial_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select

End Sub

'Private Sub txtStartNight_Change()
'  Dim t As Integer
'  t = Val(txtStartNight.text)
'  If t < 12 Then
'    If t = 0 Then
'      lblStartNightHR.Caption = "MidNight"
'    Else
'      lblStartNightHR.Caption = t & " AM"
'    End If
'  ElseIf t = 12 Then
'    lblStartNightHR.Caption = t & " PM"
'  Else
'    lblStartNightHR.Caption = t - 12 & " PM"
'  End If
'
'End Sub
'
'Private Sub txtStartNight_GotFocus()
'  SelAll txtStartNight
'End Sub
'
'Private Sub txtStartNight_KeyPress(KeyAscii As Integer)
'
'  Dim newval As Integer
'  Select Case KeyAscii
'    Case vbKeyAdd, 43
'      KeyAscii = 0
'      newval = Val(txtStartNight.text) + 1
'      txtStartNight.text = Min(newval, 23)
'    Case vbKeySubtract, 45
'      KeyAscii = 0
'      newval = Val(txtStartNight.text) - 1
'      txtStartNight.text = Max(newval, 0)
'
'    Case Else
'      KeyAscii = KeyProcMax(txtStartNight, KeyAscii, False, 0, 2, 23)
'  End Select
'
'End Sub

Private Sub txtSurveyDeviceID_GotFocus()
  SelAll txtSurveyDeviceID
End Sub

Private Sub txtSurveyDeviceID_KeyPress(KeyAscii As Integer)
' handle hex data only
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
  'KeyAscii = KeyProcHex(txtSerial, KeyAscii, False, 0, 8)

End Sub

Private Sub txtTroubleBeepTimer_GotFocus()
  SelAll txtTroubleBeepTimer
End Sub

Private Sub txtTroubleBeepTimer_KeyPress(KeyAscii As Integer)
  Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtTroubleBeepTimer.text) + 1
      txtTroubleBeepTimer.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtTroubleBeepTimer.text) - 1
      txtTroubleBeepTimer.text = Max(newval, -1)
    Case Else

      KeyAscii = KeyProcMax(txtTroubleBeepTimer, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select
End Sub

Sub UpdateCheckintab()
  

  If Configuration.AssurStart <> Val(txtAssurStart.text) Then
    Configuration.AssurStart = Val(txtAssurStart.text)
    Configuration.AssurEnd = Val(txtAssurEnd.text)
    CancelAssur
    gAssurStartTime = 0
  Else
    Configuration.AssurStart = Val(txtAssurStart.text)
    Configuration.AssurEnd = Val(txtAssurEnd.text)
  End If

  If Configuration.AssurStart2 <> Val(txtAssurStart2.text) Then
    Configuration.AssurStart2 = Val(txtAssurStart2.text)
    Configuration.AssurEnd2 = Val(txtAssurEnd2.text)
    CancelAssur
    gAssurStartTime2 = 0
  Else
    Configuration.AssurStart2 = Val(txtAssurStart2.text)
    Configuration.AssurEnd2 = Val(txtAssurEnd2.text)
  End If

  WriteSetting "Configuration", "AssurStart", Configuration.AssurStart
  WriteSetting "Configuration", "AssurEnd", Configuration.AssurEnd

  WriteSetting "Configuration", "AssurStart2", Configuration.AssurStart2
  WriteSetting "Configuration", "AssurEnd2", Configuration.AssurEnd2


  Configuration.AssurStart = Val(ReadSetting("Configuration", "AssurStart", "0"))
  Configuration.AssurEnd = Val(ReadSetting("Configuration", "AssurEnd", "0"))

  Configuration.AssurStart2 = Val(ReadSetting("Configuration", "AssurStart2", "0"))
  Configuration.AssurEnd2 = Val(ReadSetting("Configuration", "AssurEnd2", "0"))
  
  gAssurDisableScreenOutput = IIf(chkDisableScreenOutput.Value = 1, 1, 0)
  
  SaveGlobals

  ' end checkin tab

End Sub

Private Sub txtTroubleBeepTimer_LostFocus()
  txtTroubleBeepTimer.text = Max(-1, Min(Val(txtTroubleBeepTimer.text), BEEP_LIMIT))
End Sub

Private Sub txtTroubleFileName_DblClick()
  Dim Temp() As Byte
  On Error Resume Next
  Temp = GetWaveData(txtTroubleFileName.text)
  PlayMemSound Temp(0), Win32.SND_MEMORY
End Sub

Private Sub txtCommPort_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtCommPort, KeyAscii, False, 0, 3, 255)
End Sub

Private Sub txtCommTimeout_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtCommTimeout, KeyAscii, False, 0, FIVE_DIGITS, BEEP_LIMIT)
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtID, KeyAscii, False, 0, 2, 99)
End Sub
Function SaveTimers()
  
  ResetActivityTime
  
  Configuration.AlarmFile = Trim(txtAlarmFileName.text)
  Configuration.AlertFile = Trim(txtAlertFileName.text)
  Configuration.LowBattFile = Trim(txtLowBattFileName.text)
  Configuration.TroubleFile = Trim(txtTroubleFileName.text)
  Configuration.AssurFile = Trim(txtAssurFileName.text)
  
  
  
  Configuration.AlarmBeep = Max(-1, Val(txtAlarmBeepTimer.text))
  Configuration.AlertBeep = Max(-1, Val(txtAlertBeepTimer.text))
  Configuration.ExtBeep = Max(-1, Val(txtExtBeepTimer.text))
  Configuration.LowBattBeep = Max(-1, Val(txtLowBattBeepTimer.text))
  Configuration.TroubleBeep = Max(-1, Val(txtTroubleBeepTimer.text))
  Configuration.AssurBeep = Max(-1, Val(txtAssurBeepTimer.text))

  Configuration.AlarmReBeep = Max(0, Val(txtAlarmRebeep.text)) ' 0 if off
  Configuration.AlertReBeep = Max(0, Val(txtAlertRebeep.text))
  Configuration.ExtReBeep = Max(0, Val(txtExternRebeep.text))
  Configuration.LowBattReBeep = Max(0, Val(txtBattRebeep.text))
  Configuration.TroubleReBeep = Max(0, Val(txtTroubleRebeep.text))
  Configuration.AssurReBeep = Max(0, Val(txtAssurRebeep.text))

  Configuration.BeepControl = chkLocaltControl.Value And 1  ' Local Control = <> 0

  
  WriteSetting "Configuration", "AlarmFile", Configuration.AlarmFile
  WriteSetting "Configuration", "AlertFile", Configuration.AlertFile
  WriteSetting "Configuration", "LowBattFile", Configuration.LowBattFile
  WriteSetting "Configuration", "TroubleFile", Configuration.TroubleFile
  WriteSetting "Configuration", "AssurFile", Configuration.AssurFile
  
  
  
  WriteSetting "Configuration", "AlarmBeep", Configuration.AlarmBeep
  WriteSetting "Configuration", "AlertBeep", Configuration.AlertBeep
  WriteSetting "Configuration", "ExtBeep", Configuration.ExtBeep
  WriteSetting "Configuration", "LowBattBeep", Configuration.LowBattBeep
  WriteSetting "Configuration", "TroubleBeep", Configuration.TroubleBeep
  WriteSetting "Configuration", "AssurBeep", Configuration.AssurBeep


  WriteSetting "Configuration", "AlarmReBeep", Configuration.AlarmReBeep
  WriteSetting "Configuration", "AlertReBeep", Configuration.AlertReBeep
  WriteSetting "Configuration", "ExtReBeep", Configuration.ExtReBeep
  WriteSetting "Configuration", "LowBattReBeep", Configuration.LowBattReBeep
  WriteSetting "Configuration", "TroubleReBeep", Configuration.TroubleReBeep
  WriteSetting "Configuration", "AssurReBeep", Configuration.AssurReBeep

  WriteSetting "Configuration", "BeepControl", Configuration.BeepControl

End Function

Function SaveSettings() As Boolean

  Dim newport            As Integer
  Dim ClientConnection   As cClientConnection
  Dim New6080UserName    As String
  Dim New6080Password    As String

  ResetActivityTime
  
  On Error GoTo SaveSettings_Error
  newport = Val(txtCommPort.text)
  If MASTER And (USE6080 = 0) Then
    If newport <> WirelessPort.PortID Then
      WirelessPort.CommClose
      Sleep 100
      On Error Resume Next
      WirelessPort.CommOpen newport, ""
      If Err.Number Then
        messagebox Me, "Error Opening Port to Receiver, Port Com:" & newport, App.Title, vbInformation
      End If
    End If

  End If

  


  Dim NeedReconnect      As Boolean
  Dim EntcryptedValue    As String
  Dim NewIP              As String

  

  If USE6080 And (MASTER) Then

    If i6080.Status <> 0 Then
      NeedReconnect = NeedReconnect Or True
    End If

    If NoACG Then
      NeedReconnect = NeedReconnect Or True
    End If

    NewIP = Trim$(txtAGPIP.text)
    If Not ValidateIPV4(NewIP) Then
      ' actually, validate the IP
      messagebox Me, "IP Missing or Invalid", App.Title, vbInformation
      Exit Function

    End If

    WriteSetting "Configuration", "IP06080", NewIP




    If NewIP <> IP1 Then
      NeedReconnect = NeedReconnect Or True
    End If
    IP1 = ReadSetting("Configuration", "IP06080", "192.168.60.80")


    New6080UserName = Trim$(txt6080UserName.text)
    New6080Password = Trim$(txt6080Password.text)
    ' save off
    If Len(New6080UserName) Then
      ' encrypt & save off
      EntcryptedValue = MakeEnCryptedString(New6080UserName)
      Call WriteSetting("Configuration", "U06080", EntcryptedValue)
      NeedReconnect = NeedReconnect Or True
    End If



    If Len(New6080Password) Then
      ' encrypt & save off
      EntcryptedValue = MakeEnCryptedString(New6080Password)
      Call WriteSetting("Configuration", "P06080", EntcryptedValue)
      NeedReconnect = NeedReconnect Or True
    End If

    USER1 = ReadSetting("Configuration", "U06080", "Admin")
    PW1 = ReadSetting("Configuration", "P06080", "Admin")

    If USER1 = "Admin" Then
      ' no decrypt
    Else
      USER1 = MakeDeCryptedString(USER1)
    End If

    If PW1 = "Admin" Then
      ' no decrypt
    Else
      PW1 = MakeDeCryptedString(PW1)
    End If

    ' let's try and reconnect ACG

    If NeedReconnect Then
      i6080.DisConnect
    End If
    NoACG = False              ' reset an try again

    If Not Ping(IP1) Then
      NoACG = True
      messagebox Me, "No Server on " & IP1, App.Title, vbInformation
      Exit Function
    End If

    If Not NoACG Then

      i6080.Username = USER1
      i6080.Password = PW1
      i6080.useSSL = False
      i6080.IP = IP1
      i6080.SetRequestString 0

      If i6080.Get6080Data() Then
        i6080.Connect
      Else
        messagebox Me, "No ACG Response at " & IP1, App.Title, vbInformation
        Exit Function

      End If
    End If
  End If

  txtSMSAccount.text = Trim$(txtSMSAccount.text)

  gSPForwardAccount = txtSMSAccount.text
  If Len(gSPForwardAccount) = 0 Then
    chkForwardSP.Value = 0
  End If

  gForwardSoftPoints = (chkForwardSP.Value = 1) And 1

  WriteSetting "Configuration", "ForwardSoftPoints", gForwardSoftPoints
  WriteSetting "Configuration", "SPForwardAccount", gSPForwardAccount


  Configuration.HideHIPPANames = chkHIPPANames.Value
  Configuration.HideHIPPASidebar = chkHippaSidebar.Value
  WriteSetting "Configuration", "HIPPAHideNames", Configuration.HideHIPPANames
  WriteSetting "Configuration", "HideHIPPASideBar", Configuration.HideHIPPASidebar

  WriteSetting "Remote", "MyAlarms", gMyAlarms

  ' Configuration.EscTimer = Val(txtEscalate.text)






  Configuration.Facility = Trim(txtFacility.text)
  Configuration.RxTimeout = Val(txtCommTimeout.text)

  'Configuration.ID = Val(txtID.text)
  Configuration.CommPort = newport
  Configuration.locationtext = chkLocationText.Value And 1
  Configuration.locationtext = (chkLocationPhrase.Value And 1) Or Configuration.locationtext
  'chkLocationPhrase
  'WDog

  Configuration.WatchdogTimeout = Val(txtWDTimeout.text)
  If cboWDType.ListIndex > -1 Then
    Configuration.WatchdogType = Max(0, cboWDType.ItemData(cboWDType.ListIndex))
  Else
    Configuration.WatchdogType = 0
  End If

  Configuration.HostPort = Val(txtHostPort.text)
  Configuration.HostIP = GetValidIP(txtIP.text, "127.0.0.1")


  Configuration.MonitorDomain = Trim$(txtMonitorDomain.text)
  Configuration.MonitorRequest = Trim$(txtMonitorRequest.text)
  Configuration.MonitorInterval = Val(txtMonitorInterval.text)
  Configuration.MonitorPort = Val(txtMonitorPort.text)
  Configuration.MonitorEnabled = chkMonitorEnabled.Value
  Configuration.MonitorFacilityID = Trim$(txtMonitorFacilityID.text)


  gElapsedEqACK = IIf(chkElapsedEqACK.Value = 1, 1, 0)
  gTimeFormat = IIf(chkTimeFormat.Value = 1, 1, 0)

  SaveGlobals
  'WriteSetting "Configuration", "TimeFormat", CStr(gTimeFormat)

  If gTimeFormat = 1 Then
    gTimeFormatString = "hh:nn"
  Else
    gTimeFormatString = "hh:nnA/P"
  End If




  If MASTER Then

    ' stop listener
    Do Until frmTimer.WinsockHost(0).State = sckClosed
      frmTimer.WinsockHost(0).Close
      DoEvents
    Loop

    ' close all connections
    For Each ClientConnection In ClientConnections
      ClientConnection.CloseConnection
    Next

    ' restart listener
    frmTimer.WinsockHost(0).LocalPort = Configuration.HostPort
    frmTimer.WinsockHost(0).Listen

  Else
    ' close connection  to Host
    ResetRemoteRefreshCounter  ' = -5 ' should complete in 10 seconds
    Do Until HostConnection.State = sckClosed
      HostConnection.CloseConnection
      Sleep 100
      DoEvents
    Loop
    HostConnection.Connect Configuration.HostIP, Configuration.HostPort
    ResetRemoteRefreshCounter  ' should complete in  5 to 10 seconds


  End If




  


  If cboFirst.ListIndex > -1 Then
    Configuration.StartNight = cboFirst.ItemData(cboFirst.ListIndex)
  Else
    Configuration.StartNight = 0
  End If




  If cboSecond.ListIndex > -1 Then
    Configuration.EndNight = cboSecond.ItemData(cboSecond.ListIndex)
  Else
    Configuration.EndNight = 0
  End If

  If cboThird.ListIndex > -1 Then
    Configuration.EndThird = cboThird.ItemData(cboThird.ListIndex)
  Else
    Configuration.EndThird = 0
  End If


  txtRxLocation.text = Trim(txtRxLocation.text)
  Configuration.RxLocation = txtRxLocation.text

  txtRxSerial.text = Trim(txtRxSerial.text)
  Configuration.RxSerial = Right("00000000" & txtRxSerial.text, 8)

  ' serial port
  If MASTER Then

    Devices.Item(1).Description = Configuration.RxLocation
    Devices.Item(1).Serial = Configuration.RxSerial
    Devices.Item(1).SupervisePeriod = Configuration.RxTimeout
  End If

  WriteSetting "Configuration", "LocationText", Configuration.locationtext

  WriteSetting "Configuration", "Facility", Configuration.Facility
  WriteSetting "Configuration", "CommPort", Configuration.CommPort
  WriteSetting "Configuration", "RxTimeout", Configuration.RxTimeout
  WriteSetting "Configuration", "ID", Configuration.ID



  WriteSetting "Configuration", "HostPort", Configuration.HostPort
  WriteSetting "Configuration", "HostIP", Configuration.HostIP

  WriteSetting "Configuration", "EscTimer", Configuration.EscTimer


  WriteSetting "Configuration", "WDType", Configuration.WatchdogType
  WriteSetting "Configuration", "WDTimeout", Configuration.WatchdogTimeout




  WriteSetting "Configuration", "AssurStart", Configuration.AssurStart
  WriteSetting "Configuration", "AssurEnd", Configuration.AssurEnd

  WriteSetting "Configuration", "AssurStart2", Configuration.AssurStart2
  WriteSetting "Configuration", "AssurEnd2", Configuration.AssurEnd2

  WriteSetting "Configuration", "StartNight", Configuration.StartNight
  WriteSetting "Configuration", "EndNight", Configuration.EndNight
  WriteSetting "Configuration", "EndThird", Configuration.EndThird

  WriteSetting "Configuration", "RxLocation", Configuration.RxLocation
  WriteSetting "Configuration", "RxSerial", Configuration.RxSerial

  WriteSetting "Monitoring", "Domain", Configuration.MonitorDomain
  WriteSetting "Monitoring", "Request", Configuration.MonitorRequest
  WriteSetting "Monitoring", "Interval", Configuration.MonitorInterval
  WriteSetting "Monitoring", "Port", Configuration.MonitorPort
  WriteSetting "Monitoring", "Enabled", Configuration.MonitorEnabled And 1
  WriteSetting "Monitoring", "FacilityID", Configuration.MonitorFacilityID

  '  Configuration.MonitorDomain = Trim$(txtMonitorDomain.text)
  '  Configuration.Request = Trim$(txtMonitorRequest.text)
  '  Configuration.MonitorInterval = Val(txtMonitorInterval.text)
  '  Configuration.MonitorPort = Val(txtMonitorPort.text)


  Configuration.RemoteSerial = UCase$(Right$("00000000" & Trim$(txtRemoteSerial.text), 8))
  WriteSetting "Configuration", "RemoteSerial", Configuration.RemoteSerial

  ' Waypoint Survey
  If MASTER Then
    Configuration.SurveyPCA = Right("00000000" & Hex(Val("&h" & txtPCA.text)), 8)

    WriteSetting "Configuration", "SurveyPCA", Configuration.SurveyPCA
    WriteSetting "Configuration", "Surveymode", Configuration.surveymode
    If cboAvail.ListIndex > -1 Then
      Configuration.SurveyPager = cboAvail.ItemData(cboAvail.ListIndex)
    Else
      Configuration.SurveyPager = 0
    End If
    WriteSetting "Configuration", "Surveypager", Configuration.SurveyPager

    WriteSetting "Configuration", "NoNCs", Configuration.NoNCs

    Configuration.OnlyLocators = chkUseOnlyLocators.Value
    WriteSetting "Configuration", "OnlyLocators", Configuration.OnlyLocators

    gDirectedNetwork = chkDirectedNet.Value
    WriteSetting "Configuration", "DNet", gDirectedNetwork

    Configuration.WaypointDevice = Right("00000000" & Hex(Val("&h" & txtWaypointDevice.text)), 8)
    WriteSetting "Configuration", "WaypointDevice", Configuration.WaypointDevice

    Configuration.SurveyDevice = Right("00000000" & Hex(Val("&h" & txtSurveyDeviceID.text)), 8)
    WriteSetting "Configuration", "SurveyDevice", Configuration.SurveyDevice
    
    Configuration.boost = Min(BOOST_LIMIT, Val(Me.txtBoost.text))
    WriteSetting "Configuration", "Boost", Configuration.boost
    
    
  End If

  'Configuration.SendAckMSG = chkSendACKMessage.value
  'WriteSetting "Configuration", "SendAckMsg", Configuration.SendAckMSG

SaveSettings_Resume:
  On Error GoTo 0
  Exit Function

SaveSettings_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmConfigure.SaveSettings." & Erl
  Resume SaveSettings_Resume


End Function

Public Sub SaveGlobals()
  Dim rs As Recordset
  Dim Count As Long
  Dim SQL   As String
  
  ' two fields with build 377
  '        EnsureFieldExists "Global", "MilitaryTime", adInteger, "0", "0"
  '        EnsureFieldExists "Global", "ElapsedEqACK", adInteger, "0", "0"

  
  Set rs = ConnExecute("Select Count(*) From GlobalSettings")
  Count = rs(0)
  rs.Close
  Set rs = Nothing
  If Count = 0 Then
    SQL = "INSERT into GlobalSettings (MilitaryTime,ElapsedEqACK,DisableScreenOutput) values (" & gTimeFormat & "," & gElapsedEqACK & "," & gAssurDisableScreenOutput & ")"
  Else
    SQL = "UPDATE GlobalSettings SET MilitaryTime = " & gTimeFormat & ", ElapsedEqACK = " & gElapsedEqACK & ", DisableScreenOutput = " & gAssurDisableScreenOutput
  End If
  
  ConnExecute SQL
  
  
  
End Sub


Public Sub Fill()


10      AddRemoveTabs

20      cmdChangeHostAdapter.Visible = (gUser.LEvel >= LEVEL_ADMIN) And MASTER

30      GetConfig

40      Configuration.EndFirst = Configuration.StartNight

50      If Configuration.EndFirst = Configuration.EndNight Then
60        cboFirst.ListIndex = 0  ' only first shift
70        cboSecond.ListIndex = 0  ' no second shift
80        cboThird.ListIndex = 0  ' no third
90      Else
100       cboFirst.ListIndex = Max(0, CboGetIndexByItemData(cboFirst, Configuration.EndFirst))
110       cboSecond.ListIndex = Max(0, CboGetIndexByItemData(cboSecond, Configuration.EndNight))
120       cboThird.ListIndex = Max(0, CboGetIndexByItemData(cboThird, Configuration.EndThird))

130     End If

        chkHIPPANames.Value = IIf(Configuration.HideHIPPANames, 1, 0)
        chkHippaSidebar.Value = IIf(Configuration.HideHIPPASidebar, 1, 0)


140     chkSubscribedAlarms.Value = IIf(gMyAlarms, 1, 0)

150     txt6080Password.text = ""
160     txt6080UserName.text = ""

170     txtSMSAccount.text = gSPForwardAccount
180     chkForwardSP.Value = 1 And gForwardSoftPoints

190     chkLocationText.Value = Configuration.locationtext And 1
200     chkLocationPhrase.Value = Configuration.locationtext And 1

        'txtEscalate.text = Configuration.EscTimer
210     txtAlarmFileName.text = Configuration.AlarmFile
220     txtAlertFileName.text = Configuration.AlertFile
230     txtExtFileName.text = Configuration.ExtFile
240     txtLowBattFileName.text = Configuration.LowBattFile
250     txtTroubleFileName.text = Configuration.TroubleFile
260     txtAssurFileName.text = Configuration.AssurFile

270     txtAlarmBeepTimer.text = Configuration.AlarmBeep
280     txtAlertBeepTimer.text = Configuration.AlertBeep
290     txtExtBeepTimer.text = Configuration.ExtBeep
300     txtLowBattBeepTimer.text = Configuration.LowBattBeep
310     txtTroubleBeepTimer.text = Configuration.TroubleBeep
320     txtAssurBeepTimer.text = Configuration.AssurBeep


     txtAlarmRebeep.text = Configuration.AlarmReBeep
     txtAlertRebeep.text = Configuration.AlertReBeep
     txtExternRebeep.text = Configuration.ExtReBeep
     txtBattRebeep.text = Configuration.LowBattReBeep
     txtTroubleRebeep.text = Configuration.TroubleReBeep
     txtAssurRebeep.text = Configuration.AssurReBeep


     chkLocaltControl.Value = IIf(Configuration.BeepControl, 1, 0)



330     txtFacility.text = Configuration.Facility
340     txtCommTimeout.text = Configuration.RxTimeout

350     If MASTER Then

          Dim Adapter          As cAdapter
360       If Adapters Is Nothing Then
370         Set Adapters = New cAdapters
380       End If
390       Adapters.RefreshAdapters

400       Set Adapter = Adapters.GetAdapterByIP(frmTimer.WinsockHost(0).LocalIP)
410       If Not (Adapter Is Nothing) Then
420         ConsoleID = Adapter.MacAddress
430       Else
440         ConsoleID = "000000000000"
450       End If

460     Else


          Dim oldMAC           As cEthernetAdapterOLD
470       Set oldMAC = New cEthernetAdapterOLD
480       ConsoleID = oldMAC.MAC
490       Set oldMAC = Nothing



500     End If

510     txtID.text = ConsoleID
520     txtCommPort.text = Configuration.CommPort
530     txtHostPort.text = Configuration.HostPort

540     txtMonitorDomain.text = Configuration.MonitorDomain
550     txtMonitorRequest.text = Configuration.MonitorRequest
560     txtMonitorInterval.text = Configuration.MonitorInterval
570     txtMonitorPort.text = Configuration.MonitorPort
580     chkMonitorEnabled.Value = Configuration.MonitorEnabled And 1
590     txtMonitorFacilityID.text = Configuration.MonitorFacilityID

600     If MASTER Then

610       txtIP.text = frmTimer.WinsockHost(0).LocalIP
620       If USE6080 Then
630         txtAGPIP.text = IP1
640         If Not NoACG Then
650           On Error Resume Next
660           DisplayNID Get6080NID
670         End If
680       End If
690       On Error GoTo 0

700     Else
710       txtIP.text = Configuration.HostIP
720       txtAGPIP.text = IP1
          '    DisplayNID Get6080NID

730     End If

740     chkDirectedNet.Value = IIf(gDirectedNetwork, 1, 0)

750     txtAssurStart.text = Configuration.AssurStart
760     txtAssurEnd.text = Configuration.AssurEnd
770     txtAssurStart2.text = Configuration.AssurStart2
780     txtAssurEnd2.text = Configuration.AssurEnd2


790     txtSurveyDeviceID.text = Configuration.SurveyDevice
800     txtWaypointDevice.text = Configuration.WaypointDevice

        txtBoost.text = Min(BOOST_LIMIT, Configuration.boost)


810     txtPCA.text = Configuration.SurveyPCA
820     chkPCARedirect.Value = IIf(Configuration.PCARedirect = 1, 1, 0)

830     txtRxLocation.text = Configuration.RxLocation
840     chkDisableScreenOutput.Value = IIf(gAssurDisableScreenOutput = 1, 1, 0)

        txtRemoteSerial.text = UCase$(Right$("00000000" & Trim$(Configuration.RemoteSerial), 8))

850     cmd6080.Visible = False
860     If MASTER Then
870       If USE6080 Then
880         cmd6080.Visible = False
890         On Error Resume Next
900         If Not NoACG Then
910           i6080.Get6080Data
920           Configuration.RxSerial = i6080.SerialNumber
930         End If
940         On Error GoTo 0
950       Else
960         If gUser.LEvel = LEVEL_FACTORY Then
970           cmd6080.Visible = True
980         Else
990           cmd6080.Visible = False
1000        End If
1010        Call GetNCNID
1020        Configuration.RxSerial = GetNCSerial()
1030      End If

1040    End If

1050    txtRxSerial.text = Configuration.RxSerial

1060    If USE6080 Then
1070      On Error Resume Next
1080      If Not NoACG Then
1090        DisplayNID Get6080NID
1100      End If
1110      On Error GoTo 0
1120    Else
1130      If GlobalNID >= 0 And GlobalNID < 32 Then
1140        DisplayNID GlobalNID
1150      End If

1160    End If



1170    chkTimeFormat.Value = IIf(gTimeFormat = 1, 1, 0)
1180    chkElapsedEqACK.Value = IIf(gElapsedEqACK = 1, 1, 0)

1190    chkUseOnlyLocators.Value = IIf(Configuration.OnlyLocators = 1, 1, 0)

1200    FillAvailablePagers

1210    RefreshLicensing

1220    chkNoNCs.Value = IIf(Configuration.NoNCs = 1, 1, 0)


        ' watchdog settings
        Dim j                  As Long

1230    txtWDTimeout.text = Configuration.WatchdogTimeout

1240    For j = cboWDType.listcount - 1 To 1 Step -1
1250      If cboWDType.ItemData(j) = Configuration.WatchdogType Then
1260        Exit For
1270      End If
1280    Next

1290    On Error Resume Next
1300    j = Max(0, j)
1310    cboWDType.ListIndex = j
1320    Err.Clear
1330    txtDBInfo.text = conn.provider & " " & conn.Properties("Data Source Name")

1340    If Err.Number <> 0 Then
1350      txtDBInfo.text = "DB Data not retrieved"
1360    End If
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

Private Sub txtTroubleFileName_GotFocus()
  SelAll txtTroubleFileName
End Sub

Private Sub txtTroubleRebeep_GotFocus()
  SelAll txtTroubleRebeep
End Sub

Private Sub txtTroubleRebeep_KeyPress(KeyAscii As Integer)
Dim newval As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtTroubleRebeep.text) + 1
      txtTroubleRebeep.text = Max(-1, Min(newval, BEEP_LIMIT))
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtTroubleRebeep.text) - 1
      txtTroubleRebeep.text = Max(newval, -1)
    Case Else
      KeyAscii = KeyProcMax(txtTroubleRebeep, KeyAscii, True, 0, FIVE_DIGITS, BEEP_LIMIT)
  End Select

End Sub

Private Sub txtWDTimeout_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyProcMax(txtWDTimeout, KeyAscii, False, 0, 3, 255)
End Sub
