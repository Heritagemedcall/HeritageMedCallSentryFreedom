VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmTransmitterTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transmitter Types"
   ClientHeight    =   15165
   ClientLeft      =   3255
   ClientTop       =   1500
   ClientWidth     =   9045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15165
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   14700
      Left            =   -30
      TabIndex        =   0
      Top             =   180
      Width           =   9030
      Begin VB.Frame fradef 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   240
         TabIndex        =   84
         Top             =   3180
         Width           =   7125
         Begin VB.CheckBox chkUpdateCheckin 
            Alignment       =   1  'Right Justify
            Caption         =   "Update Checkin"
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
            Left            =   1140
            TabIndex        =   117
            Top             =   1500
            Width           =   2115
         End
         Begin VB.CheckBox chkClearByReset 
            Alignment       =   1  'Right Justify
            Caption         =   "Clear by Reset"
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
            Left            =   1140
            TabIndex        =   119
            Top             =   900
            Width           =   2085
         End
         Begin VB.CommandButton cmdGlobalMain 
            Caption         =   "Live Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5580
            TabIndex        =   118
            Top             =   1620
            Width           =   1395
         End
         Begin VB.CheckBox chkIgnoreTamper 
            Alignment       =   1  'Right Justify
            Caption         =   "Ignore Tamper"
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
            Left            =   4440
            TabIndex        =   116
            Top             =   900
            Width           =   2115
         End
         Begin VB.TextBox txtDescription 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   510
            Width           =   3165
         End
         Begin VB.TextBox txtCheckin 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   5745
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   86
            TabStop         =   0   'False
            ToolTipText     =   "Checkin Period for this Device Type in Minutes"
            Top             =   510
            Width           =   825
         End
         Begin VB.TextBox txtAutoClear 
            Alignment       =   1  'Right Justify
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
            Left            =   5940
            MaxLength       =   3
            TabIndex        =   85
            ToolTipText     =   "Enter 1 to 999. Zero Disables"
            Top             =   150
            Width           =   630
         End
         Begin VB.ComboBox cboDeviceType 
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
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   150
            Width           =   2265
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   150
            Width           =   1995
         End
         Begin VB.Label lblModel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
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
            Left            =   480
            TabIndex        =   94
            Top             =   225
            Width           =   525
         End
         Begin VB.Label lblDescription 
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
            Left            =   30
            TabIndex        =   93
            Top             =   570
            Width           =   975
         End
         Begin VB.Label lblCheckin 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Checkin Period"
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
            Left            =   4380
            TabIndex        =   92
            Top             =   570
            Width           =   1305
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min."
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
            Left            =   6600
            TabIndex        =   91
            Top             =   210
            Width           =   375
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Auto Clear"
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
            Left            =   4950
            TabIndex        =   90
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min."
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
            Left            =   6600
            TabIndex        =   89
            Top             =   570
            Width           =   375
         End
      End
      Begin VB.Frame fraGroups1 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   2385
         Left            =   150
         TabIndex        =   30
         Top             =   9960
         Width           =   7245
         Begin VB.TextBox txtGG1 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   155
            ToolTipText     =   "Escalation Timeout"
            Top             =   75
            Width           =   585
         End
         Begin VB.TextBox txtGG2 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   154
            ToolTipText     =   "Escalation Timeout"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox txtGG3 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   153
            Top             =   705
            Width           =   585
         End
         Begin VB.TextBox txtGG4 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   152
            ToolTipText     =   "Escalation Timeout"
            Top             =   1035
            Width           =   585
         End
         Begin VB.TextBox txtGG5 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   151
            ToolTipText     =   "Escalation Timeout"
            Top             =   1350
            Width           =   585
         End
         Begin VB.TextBox txtGG6 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   150
            ToolTipText     =   "Escalation Timeout"
            Top             =   1665
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cboGroupG3 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   705
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG4 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   1020
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG5 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   1335
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG6 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   1650
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG2 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   375
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG1 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   60
            Width           =   1425
         End
         Begin VB.TextBox txtOG1D 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   32
            ToolTipText     =   "Escalation Timeout"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txtOG2D 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   34
            ToolTipText     =   "Escalation Timeout"
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox txtOG3D 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   36
            ToolTipText     =   "Escalation Timeout"
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox txtOG4D 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   38
            ToolTipText     =   "Escalation Timeout"
            Top             =   990
            Width           =   585
         End
         Begin VB.TextBox txtOG5D 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   40
            ToolTipText     =   "Escalation Timeout"
            Top             =   1305
            Width           =   585
         End
         Begin VB.TextBox txtOG6D 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   42
            ToolTipText     =   "Escalation Timeout"
            Top             =   1620
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtNG1D 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   44
            ToolTipText     =   "Escalation Timeout"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txtNG2D 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   46
            ToolTipText     =   "Escalation Timeout"
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox txtNG3D 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   48
            ToolTipText     =   "Escalation Timeout"
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox txtNG4D 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   50
            ToolTipText     =   "Escalation Timeout"
            Top             =   1020
            Width           =   585
         End
         Begin VB.TextBox txtNG5D 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   52
            ToolTipText     =   "Escalation Timeout"
            Top             =   1335
            Width           =   585
         End
         Begin VB.TextBox txtNG6D 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   54
            ToolTipText     =   "Escalation Timeout"
            Top             =   1650
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cboGroupN6 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1620
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN5 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1305
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN4 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   990
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN3 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   675
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup6 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1620
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup5 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1305
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup4 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   990
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup3 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   675
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup2 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup1 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   45
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN1 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   45
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN2 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3rd"
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
            Left            =   4800
            TabIndex        =   125
            Top             =   105
            Width           =   285
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E"
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
            Left            =   4950
            TabIndex        =   124
            Top             =   1710
            Width           =   135
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D"
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
            Left            =   4935
            TabIndex        =   123
            Top             =   1395
            Width           =   150
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
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
            Left            =   4950
            TabIndex        =   122
            Top             =   1065
            Width           =   135
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
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
            Left            =   4950
            TabIndex        =   121
            Top             =   750
            Width           =   135
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
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
            Left            =   4950
            TabIndex        =   120
            Top             =   420
            Width           =   135
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
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
            TabIndex        =   115
            Top             =   405
            Width           =   135
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
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
            TabIndex        =   114
            Top             =   735
            Width           =   135
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
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
            TabIndex        =   113
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D"
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
            Left            =   2505
            TabIndex        =   112
            Top             =   1380
            Width           =   150
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E"
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
            TabIndex        =   111
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
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
            Left            =   180
            TabIndex        =   110
            Top             =   405
            Width           =   135
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
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
            Left            =   180
            TabIndex        =   109
            Top             =   735
            Width           =   135
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
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
            Left            =   180
            TabIndex        =   108
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D"
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
            Left            =   165
            TabIndex        =   107
            Top             =   1380
            Width           =   150
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E"
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
            Left            =   180
            TabIndex        =   106
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label lblOG1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1st"
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
            Left            =   45
            TabIndex        =   56
            Top             =   90
            Width           =   270
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2nd"
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
            Left            =   2370
            TabIndex        =   55
            Top             =   90
            Width           =   330
         End
      End
      Begin VB.Frame fraGroups2 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   2385
         Left            =   150
         TabIndex        =   57
         Top             =   12390
         Width           =   7215
         Begin VB.ComboBox cboGroupG1_A 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   60
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG2_A 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   375
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG6_A 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   141
            Top             =   1650
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG5_A 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   1335
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG4_A 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   139
            Top             =   1020
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG3_A 
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
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   705
            Width           =   1425
         End
         Begin VB.TextBox txtGG6_A 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   137
            ToolTipText     =   "Escalation Timeout"
            Top             =   1665
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtGG5_A 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   136
            ToolTipText     =   "Escalation Timeout"
            Top             =   1350
            Width           =   585
         End
         Begin VB.TextBox txtGG4_A 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   135
            ToolTipText     =   "Escalation Timeout"
            Top             =   1035
            Width           =   585
         End
         Begin VB.TextBox txtGG3_A 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   134
            Top             =   705
            Width           =   585
         End
         Begin VB.TextBox txtGG2_A 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   133
            ToolTipText     =   "Escalation Timeout"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox txtGG1_A 
            Alignment       =   1  'Right Justify
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   132
            ToolTipText     =   "Escalation Timeout"
            Top             =   75
            Width           =   585
         End
         Begin VB.TextBox txtOG1_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   59
            ToolTipText     =   "Escalation Timeout"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txtOG2_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   61
            ToolTipText     =   "Escalation Timeout"
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox txtOG3_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   63
            ToolTipText     =   "Escalation Timeout"
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox txtOG4_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   65
            ToolTipText     =   "Escalation Timeout"
            Top             =   990
            Width           =   585
         End
         Begin VB.TextBox txtOG5_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   67
            ToolTipText     =   "Escalation Timeout"
            Top             =   1305
            Width           =   585
         End
         Begin VB.TextBox txtOG6_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   69
            ToolTipText     =   "Escalation Timeout"
            Top             =   1620
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtNG1_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   71
            ToolTipText     =   "Escalation Timeout"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txtNG2_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   73
            ToolTipText     =   "Escalation Timeout"
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox txtNG3_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   75
            ToolTipText     =   "Escalation Timeout"
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox txtNG4_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   77
            ToolTipText     =   "Escalation Timeout"
            Top             =   1020
            Width           =   585
         End
         Begin VB.TextBox txtNG5_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   79
            ToolTipText     =   "Escalation Timeout"
            Top             =   1335
            Width           =   585
         End
         Begin VB.TextBox txtNG6_AD 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   81
            ToolTipText     =   "Escalation Timeout"
            Top             =   1650
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cboGroupN6_A 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1620
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN5_A 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1305
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN4_A 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   990
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN3_A 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   675
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup6_A 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   1620
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup5_A 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1305
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup4_A 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   990
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup3_A 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   675
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup2_A 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   360
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup1_A 
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
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   45
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN1_A 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   45
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN2_A 
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
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3rd"
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
            Left            =   4800
            TabIndex        =   131
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E"
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
            Left            =   4950
            TabIndex        =   130
            Top             =   1725
            Width           =   135
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D"
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
            Left            =   4935
            TabIndex        =   129
            Top             =   1410
            Width           =   150
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
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
            Left            =   4950
            TabIndex        =   128
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
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
            Left            =   4950
            TabIndex        =   127
            Top             =   765
            Width           =   135
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
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
            Left            =   4950
            TabIndex        =   126
            Top             =   435
            Width           =   135
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
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
            TabIndex        =   105
            Top             =   405
            Width           =   135
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " B"
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
            Left            =   2460
            TabIndex        =   104
            Top             =   735
            Width           =   195
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
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
            TabIndex        =   103
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " D"
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
            Left            =   2445
            TabIndex        =   102
            Top             =   1380
            Width           =   210
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " E"
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
            Left            =   2460
            TabIndex        =   101
            Top             =   1695
            Width           =   195
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
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
            Left            =   180
            TabIndex        =   100
            Top             =   405
            Width           =   135
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
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
            Left            =   180
            TabIndex        =   99
            Top             =   735
            Width           =   135
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
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
            Left            =   180
            TabIndex        =   98
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D"
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
            Left            =   165
            TabIndex        =   97
            Top             =   1380
            Width           =   150
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E"
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
            Left            =   180
            TabIndex        =   96
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1st"
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
            Left            =   45
            TabIndex        =   83
            Top             =   90
            Width           =   270
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2nd"
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
            Left            =   2370
            TabIndex        =   82
            Top             =   90
            Width           =   330
         End
      End
      Begin VB.Frame fraInput0 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1485
         Left            =   7680
         TabIndex        =   11
         Top             =   6120
         Width           =   2085
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Inputs"
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
            Left            =   690
            TabIndex        =   12
            Top             =   600
            Width           =   840
         End
      End
      Begin VB.Frame fraInput2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   180
         TabIndex        =   10
         Top             =   7680
         Width           =   7185
         Begin VB.CheckBox chkClearByReset_A 
            Alignment       =   1  'Right Justify
            Caption         =   "Clear by Reset"
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
            Left            =   2460
            TabIndex        =   25
            Top             =   390
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.CheckBox chkSendCancel_A 
            Alignment       =   1  'Right Justify
            Caption         =   "Send Cancel Notice"
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
            Left            =   2460
            TabIndex        =   24
            Top             =   780
            Width           =   2085
         End
         Begin VB.TextBox txtPause_A 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   23
            Top             =   780
            Width           =   510
         End
         Begin VB.TextBox txtRepeats_A 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            MaxLength       =   2
            TabIndex        =   22
            Top             =   390
            Width           =   510
         End
         Begin VB.CheckBox chkRepeatUntil_A 
            Alignment       =   1  'Right Justify
            Caption         =   "Repeat Until Reset"
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
            Left            =   0
            TabIndex        =   21
            Top             =   1110
            Width           =   2355
         End
         Begin VB.TextBox txtAnnounce2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   30
            Width           =   3165
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repeats"
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
            Left            =   1020
            TabIndex        =   27
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repeat Every (Sec.)"
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
            Left            =   30
            TabIndex        =   26
            Top             =   810
            Width           =   1740
         End
         Begin VB.Label lblAnnounce2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Announce 2"
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
            Left            =   720
            TabIndex        =   14
            Top             =   90
            Width           =   1035
         End
      End
      Begin VB.Frame fraInput1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   210
         TabIndex        =   9
         Top             =   5460
         Width           =   7155
         Begin VB.TextBox txtRepeats 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            MaxLength       =   2
            TabIndex        =   28
            Top             =   390
            Width           =   510
         End
         Begin VB.CheckBox chkSendCancel 
            Alignment       =   1  'Right Justify
            Caption         =   "Send Cancel Notice"
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
            Left            =   2460
            TabIndex        =   19
            Top             =   810
            Width           =   2085
         End
         Begin VB.TextBox txtPause 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   18
            Top             =   780
            Width           =   510
         End
         Begin VB.CheckBox chkRepeatUntil 
            Alignment       =   1  'Right Justify
            Caption         =   "Repeat Until Reset"
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
            Left            =   0
            TabIndex        =   17
            Top             =   1110
            Width           =   2355
         End
         Begin VB.TextBox txtAnnounce 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   30
            Width           =   3165
         End
         Begin VB.Label lblRepeats 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repeats"
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
            Left            =   1020
            TabIndex        =   29
            Top             =   450
            Width           =   720
         End
         Begin VB.Label lblPause 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repeat Every (Sec.)"
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
            Left            =   30
            TabIndex        =   20
            Top             =   810
            Width           =   1740
         End
         Begin VB.Label lblAnnounce 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Announce 1"
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
            Left            =   720
            TabIndex        =   16
            Top             =   90
            Width           =   1035
         End
      End
      Begin MSComctlLib.TabStrip tabstrip 
         Height          =   2955
         Left            =   150
         TabIndex        =   8
         Top             =   60
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   5212
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Input 1"
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
      Begin VB.CheckBox chkAllowDisable 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow Disable"
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
         Left            =   2985
         TabIndex        =   2
         Top             =   2265
         Width           =   2325
      End
      Begin VB.CheckBox chkPortable 
         Alignment       =   1  'Right Justify
         Caption         =   "Portable"
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
         Left            =   2985
         TabIndex        =   1
         Top             =   1905
         Width           =   2325
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
         TabIndex        =   6
         Top             =   1785
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdSave 
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   615
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdNew 
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
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
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
         TabIndex        =   7
         Top             =   2370
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmTransmitterTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Model      As String
Public MIDPTI     As Long
Public CLSPTI     As Long
Public AutoClear  As Integer

Private DeviceType  As cDeviceType
Private ESDevice    As ESDeviceTypeType
Private DeviceTypes As Collection
Private mEditmode   As Integer

Private oldindex    As Integer

Private mBusy       As Boolean


Private Sub txtGG1_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG1_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG1_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG1, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG2_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG2_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG2_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG2, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG3_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG3_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG3_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG3, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG4_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG4_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG4_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG4, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG5_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG5_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG5_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG5, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG6_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG6_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG6_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG6, KeyAscii, False, 0, 3, 999)
End Sub


'Private Sub txtGG1_A_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtGG1_A, KeyAscii, False, 0, 3, 999)
'End Sub

'Private Sub txtGG2_A_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtGG2_A, KeyAscii, False, 0, 3, 999)
'End Sub

'Private Sub txtGG3_A_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtGG3_A, KeyAscii, False, 0, 3, 999)
'End Sub
'
'Private Sub txtGG4_A_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtGG4_A, KeyAscii, False, 0, 3, 999)
'End Sub
'
'Private Sub txtGG5_A_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtGG5_A, KeyAscii, False, 0, 3, 999)
'End Sub
'
'Private Sub txtGG6_A_KeyPress(KeyAscii As Integer)
'  KeyAscii = KeyProcMax(txtGG6_A, KeyAscii, False, 0, 3, 999)
'End Sub

Sub ArrangeControls()
  fraEnabler.BackColor = Me.BackColor
  
  fradef.BackColor = Me.BackColor
  fradef.left = TabStrip.ClientLeft
  fradef.top = TabStrip.ClientTop
  fradef.Width = TabStrip.ClientWidth
  fradef.Height = TabStrip.ClientHeight
  
  
  fraInput0.BackColor = Me.BackColor
  fraInput0.left = TabStrip.ClientLeft
  fraInput0.top = TabStrip.ClientTop
  fraInput0.Width = TabStrip.ClientWidth
  fraInput0.Height = TabStrip.ClientHeight
  
  fraInput1.BackColor = Me.BackColor
  fraInput1.left = TabStrip.ClientLeft
  fraInput1.top = TabStrip.ClientTop
  fraInput1.Width = TabStrip.ClientWidth
  fraInput1.Height = TabStrip.ClientHeight
  
  fraInput2.BackColor = Me.BackColor
  fraInput2.left = TabStrip.ClientLeft
  fraInput2.top = TabStrip.ClientTop
  fraInput2.Width = TabStrip.ClientWidth
  fraInput2.Height = TabStrip.ClientHeight

  fraGroups1.left = TabStrip.ClientLeft
  fraGroups1.top = TabStrip.ClientTop
  fraGroups1.Width = TabStrip.ClientWidth
  fraGroups1.Height = TabStrip.ClientHeight
  fraGroups1.BackColor = Me.BackColor

  fraGroups2.left = TabStrip.ClientLeft
  fraGroups2.top = TabStrip.ClientTop
  fraGroups2.Width = TabStrip.ClientWidth
  fraGroups2.Height = TabStrip.ClientHeight
  fraGroups2.BackColor = Me.BackColor

End Sub

Private Sub cboDeviceType_Click()

  Display
  If oldindex <> cboDeviceType.ListIndex Then
    If TabStrip.Tabs.Count > 1 Then
      TabStrip.Tabs(1).Selected = True
    End If
  End If
  oldindex = cboDeviceType.ListIndex
  
End Sub

Sub ClearForm()
  cboDeviceType.ListIndex = 0
  txtModel.text = ""
  txtDescription.text = ""
  txtAnnounce.text = ""
  txtAnnounce2.text = ""
  'chkLatching.value = 0
  chkPortable.Value = 0
  chkAllowDisable.Value = 0
  txtCheckin.text = 0
  chkClearByReset.Value = 0
  chkIgnoreTamper.Value = 0
  txtAutoClear = 0

  txtRepeats.text = 0
  txtPause.text = 0
  chkRepeatUntil.Value = 0
  chkSendCancel.Value = 0
  

End Sub

Private Sub cmdDelete_Click()
'  DeleteDevicetype
End Sub

Private Sub cmdEdit_Click()
  If Editmode <> 2 Then
    Editmode = 2  ' edit
    Win32.SetActiveWindow Me.txtDescription.hwnd
  End If

End Sub

Private Sub cmdExit_Click()
  
  If Busy Then Exit Sub
  
  If Editmode = 0 Then
    PreviousForm
    Unload Me
  Else
    Editmode = 0
  End If
End Sub

Private Sub cmdGlobalInput1_Click()
  If DoSave() Then
  
    ' apply this page globally
  End If

End Sub

Private Sub cmdGlobalMain_Click()
  Busy = True
  If DoSave() Then
    If (cboDeviceType.ListIndex > -1) Then
      ResetRunningDevices cboDeviceType.text
    End If
    ' apply this page globally
  End If
  
  Busy = False
End Sub

Private Sub cmdNew_Click()
  Model = 0
  ClearForm
  Editmode = 1
End Sub

Private Sub cmdSave_Click()
  Busy = True
  DoSave
  Busy = False
End Sub

Function DoSave() As Boolean
  If Validate() Then
    If Save() Then
      Editmode = 0
      'Fill
      DoSave = True
    End If
  End If

End Function

Private Sub Display()
        Dim index As Integer


10      index = cboDeviceType.ListIndex
20      If index > 0 Then
30        ESDevice = GetDeviceTypeByModel(cboDeviceType.text)
          
40        Set DeviceType = DeviceTypes(index)
          
50        If Not DeviceType Is Nothing Then
            
60          Debug.Print "Device # "; DeviceType.Model; " "; DeviceType.NumInputs
            
70          If DeviceType.NumInputs > 1 Then
80            If TabStrip.Tabs.Count <= 1 Then
90              TabStrip.Tabs.Add 2, "input1", "Input 1"
100             TabStrip.Tabs.Add 3, "groups1", "Outputs 1"
110           End If
120           If TabStrip.Tabs.Count <= 3 Then
130             TabStrip.Tabs.Add 4, "input2", "Input 2"
140             TabStrip.Tabs.Add 5, "groups2", "Outputs 2"
150             TabStrip.Tabs(1).Selected = True

160           End If
170           fraInput0.Visible = False
180         ElseIf DeviceType.NumInputs <= 1 Then
190           If TabStrip.Tabs.Count > 3 Then
200             Do Until TabStrip.Tabs.Count = 3
210               TabStrip.Tabs.Remove TabStrip.Tabs.Count
220             Loop
230           End If
240           If TabStrip.Tabs.Count < 2 Then
250             TabStrip.Tabs.Add 2, "input1", "Input 1"
260             TabStrip.Tabs.Add 3, "groups1", "Outputs 1"
270             TabStrip.Tabs(1).Selected = True
280           End If

290           fraInput0.Visible = False
300         End If
            
310       Else 'DeviceType Is Nothing
320         If TabStrip.Tabs.Count > 0 Then
330           Do Until TabStrip.Tabs.Count = 3
340             TabStrip.Tabs.Remove TabStrip.Tabs.Count
350           Loop
360         End If
            
370         fraInput0.Visible = False
380         fraInput1.Visible = False
390         fraInput2.Visible = False
400         fraGroups1.Visible = False
410         fraGroups2.Visible = False

420       End If ' Not DeviceType Is Nothing
          
          
          
430       Model = ESDevice.Model
440       MIDPTI = ESDevice.MIDPTI
450       CLSPTI = ESDevice.CLSPTI
460       txtModel.text = ESDevice.Model
          ' it's mangling this
          
470       txtDescription.text = DeviceType.Description  ' cboDeviceType
          'chkLatching.value = DeviceType.IsLatching
480       chkPortable.Value = DeviceType.IsPortable
490       chkAllowDisable.Value = DeviceType.AllowDisable
500       txtCheckin.text = DeviceType.Checkin
          
510       txtAutoClear.text = DeviceType.AutoClear
          
520       txtAnnounce.text = DeviceType.Announce
530       txtAnnounce2.text = DeviceType.Announce2
540       chkClearByReset.Value = DeviceType.ClearByReset
550       If DeviceType.NumInputs > 1 Then
560         txtAnnounce2.Visible = True
570         lblAnnounce2.Visible = True
580       Else
590         txtAnnounce2.Visible = False
600         lblAnnounce2.Visible = False
610       End If

620       cboGroup1.ListIndex = CboGetIndexByItemData(cboGroup1, DeviceType.OG1)
630       cboGroup2.ListIndex = CboGetIndexByItemData(cboGroup2, DeviceType.OG2)
640       cboGroup3.ListIndex = CboGetIndexByItemData(cboGroup3, DeviceType.OG3)
650       cboGroup4.ListIndex = CboGetIndexByItemData(cboGroup4, DeviceType.OG4)
660       cboGroup5.ListIndex = CboGetIndexByItemData(cboGroup5, DeviceType.OG5)
670       cboGroup6.ListIndex = CboGetIndexByItemData(cboGroup6, DeviceType.OG6)
          
680       cboGroupN1.ListIndex = CboGetIndexByItemData(cboGroupN1, DeviceType.NG1)
690       cboGroupN2.ListIndex = CboGetIndexByItemData(cboGroupN2, DeviceType.NG2)
700       cboGroupN3.ListIndex = CboGetIndexByItemData(cboGroupN3, DeviceType.NG3)
710       cboGroupN4.ListIndex = CboGetIndexByItemData(cboGroupN4, DeviceType.NG4)
720       cboGroupN5.ListIndex = CboGetIndexByItemData(cboGroupN5, DeviceType.NG5)
730       cboGroupN6.ListIndex = CboGetIndexByItemData(cboGroupN6, DeviceType.NG6)

740       cboGroupG1.ListIndex = CboGetIndexByItemData(cboGroupG1, DeviceType.GG1)
750       cboGroupG2.ListIndex = CboGetIndexByItemData(cboGroupG2, DeviceType.GG2)
760       cboGroupG3.ListIndex = CboGetIndexByItemData(cboGroupG3, DeviceType.GG3)
770       cboGroupG4.ListIndex = CboGetIndexByItemData(cboGroupG4, DeviceType.GG4)
780       cboGroupG5.ListIndex = CboGetIndexByItemData(cboGroupG5, DeviceType.GG5)
790       cboGroupG6.ListIndex = CboGetIndexByItemData(cboGroupG6, DeviceType.GG6)


800       cboGroup1_A.ListIndex = CboGetIndexByItemData(cboGroup1_A, DeviceType.OG1_A)
810       cboGroup2_A.ListIndex = CboGetIndexByItemData(cboGroup2_A, DeviceType.OG2_A)
820       cboGroup3_A.ListIndex = CboGetIndexByItemData(cboGroup3_A, DeviceType.OG3_A)
830       cboGroup4_A.ListIndex = CboGetIndexByItemData(cboGroup4_A, DeviceType.OG4_A)
840       cboGroup5_A.ListIndex = CboGetIndexByItemData(cboGroup5_A, DeviceType.OG5_A)
850       cboGroup6_A.ListIndex = CboGetIndexByItemData(cboGroup6_A, DeviceType.OG6_A)
          
860       cboGroupN1_A.ListIndex = CboGetIndexByItemData(cboGroupN1_A, DeviceType.NG1_A)
870       cboGroupN2_A.ListIndex = CboGetIndexByItemData(cboGroupN2_A, DeviceType.NG2_A)
880       cboGroupN3_A.ListIndex = CboGetIndexByItemData(cboGroupN3_A, DeviceType.NG3_A)
890       cboGroupN4_A.ListIndex = CboGetIndexByItemData(cboGroupN4_A, DeviceType.NG4_A)
900       cboGroupN5_A.ListIndex = CboGetIndexByItemData(cboGroupN5_A, DeviceType.NG5_A)
910       cboGroupN6_A.ListIndex = CboGetIndexByItemData(cboGroupN6_A, DeviceType.NG6_A)

920       cboGroupG1_A.ListIndex = CboGetIndexByItemData(cboGroupG1_A, DeviceType.GG1_A)
930       cboGroupG2_A.ListIndex = CboGetIndexByItemData(cboGroupG2_A, DeviceType.GG2_A)
940       cboGroupG3_A.ListIndex = CboGetIndexByItemData(cboGroupG3_A, DeviceType.GG3_A)
950       cboGroupG4_A.ListIndex = CboGetIndexByItemData(cboGroupG4_A, DeviceType.GG4_A)
960       cboGroupG5_A.ListIndex = CboGetIndexByItemData(cboGroupG5_A, DeviceType.GG5_A)
970       cboGroupG6_A.ListIndex = CboGetIndexByItemData(cboGroupG6_A, DeviceType.GG6_A)

        ' new with extended escalation
        
        
980       txtOG1D.text = DeviceType.OG1D
990       txtOG2D.text = DeviceType.OG2D
1000      txtOG3D.text = DeviceType.OG3D
1010      txtOG4D.text = DeviceType.OG4D
1020      txtOG5D.text = DeviceType.OG5D
1030      txtOG6D.text = DeviceType.OG6D


1040      txtNG1D.text = DeviceType.NG1D
1050      txtNG2D.text = DeviceType.NG2D
1060      txtNG3D.text = DeviceType.NG3D
1070      txtNG4D.text = DeviceType.NG4D
1080      txtNG5D.text = DeviceType.NG5D
1090      txtNG6D.text = DeviceType.NG6D

1100      txtGG1.text = DeviceType.GG1D
1110      txtGG2.text = DeviceType.GG2D
1120      txtGG3.text = DeviceType.GG3D
1130      txtGG4.text = DeviceType.GG4D
1140      txtGG5.text = DeviceType.GG5D
1150      txtGG6.text = DeviceType.GG6D


1160      txtOG1_AD.text = DeviceType.OG1_AD
1170      txtOG2_AD.text = DeviceType.OG2_AD
1180      txtOG3_AD.text = DeviceType.OG3_AD
1190      txtOG4_AD.text = DeviceType.OG4_AD
1200      txtOG5_AD.text = DeviceType.OG5_AD
1210      txtOG6_AD.text = DeviceType.OG6_AD

1220      txtNG1_AD.text = DeviceType.NG1_AD
1230      txtNG2_AD.text = DeviceType.NG2_AD
1240      txtNG3_AD.text = DeviceType.NG3_AD
1250      txtNG4_AD.text = DeviceType.NG4_AD
1260      txtNG5_AD.text = DeviceType.NG5_AD
1270      txtNG6_AD.text = DeviceType.NG6_AD

1280      txtGG1_A.text = DeviceType.GG1_AD
1290      txtGG2_A.text = DeviceType.GG2_AD
1300      txtGG3_A.text = DeviceType.GG3_AD
1310      txtGG4_A.text = DeviceType.GG4_AD
1320      txtGG5_A.text = DeviceType.GG5_AD
1330      txtGG6_A.text = DeviceType.GG6_AD


          'new with build 226

1340      txtRepeats.text = DeviceType.Repeats
1350      txtPause.text = DeviceType.Pause
1360      chkRepeatUntil.Value = IIf(DeviceType.RepeatUntil = 1, 1, 0)
1370      chkSendCancel.Value = IIf(DeviceType.SendCancel = 1, 1, 0)

1380      txtRepeats_A.text = DeviceType.Repeats_A
1390      txtPause_A.text = DeviceType.Pause_A
1400      chkRepeatUntil_A.Value = IIf(DeviceType.RepeatUntil_A = 1, 1, 0)
1410      chkSendCancel_A.Value = IIf(DeviceType.SendCancel_A = 1, 1, 0)

1420      chkIgnoreTamper.Value = IIf(DeviceType.IgnoreTamper = 1, 1, 0)


1430      Editmode = 0  ' none
1440      'cmdDelete.Enabled = True
1450
1460    Else
1470      Model = 0
1480      ClearForm
1490      cmdDelete.Enabled = False
1500      cmdEdit.Enabled = False
1510      cmdSave.Enabled = False
1520      Do Until TabStrip.Tabs.Count = 1
1530        TabStrip.Tabs.Remove TabStrip.Tabs.Count
1540      Loop
        
1550    End If

        cmdSave.Enabled = Not Busy

End Sub

Private Property Let Editmode(ByVal Value As Integer)
  txtCheckin.Locked = False
  'cboDeviceType.Visible = False
  cmdNew.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  txtDescription.Locked = False
  txtAnnounce.Locked = False
  'chkLatching.Enabled = True
  chkPortable.Enabled = True
  chkAllowDisable.Enabled = True
  cmdSave.Enabled = True

  mEditmode = Value
End Property

Private Property Get Editmode() As Integer
  Editmode = mEditmode
End Property

Public Sub Fill()

  Dim d   As cDeviceType
  Dim i   As Integer
  
  ClearForm
  

  Set DeviceTypes = New Collection
  cboDeviceType.Clear
  AddToCombo cboDeviceType, "<Select>", 0

  For i = 0 To MAX_ESDEVICETYPES
    If Len(ESDeviceType(i).Model) Then
    Set d = New cDeviceType
    d.Model = ESDeviceType(i).Model
    d.Description = ESDeviceType(i).desc
    d.IsLatching = ESDeviceType(i).Latching
    d.Checkin = ESDeviceType(i).Checkin
    d.Announce = ESDeviceType(i).Announce
    d.AllowDisable = ESDeviceType(i).AllowDisable
    d.IsPortable = ESDeviceType(i).Portable
    d.MIDPTI = ESDeviceType(i).MIDPTI
    d.CLSPTI = ESDeviceType(i).CLSPTI
    d.NumInputs = ESDeviceType(i).NumInputs
    d.NoTamper = ESDeviceType(i).NoTamper
    d.AutoClear = ESDeviceType(i).AutoClear
    
    d.IgnoreTamper = ESDeviceType(i).IgnoreTamper
    
    'new with build 226
    d.Repeats = ESDeviceType(i).Repeats
    d.Pause = ESDeviceType(i).Pause
    d.RepeatUntil = ESDeviceType(i).RepeatUntil
    d.SendCancel = ESDeviceType(i).SendCancel
    
    d.Repeats_A = ESDeviceType(i).Repeats_A
    d.Pause_A = ESDeviceType(i).Pause_A
    d.RepeatUntil_A = ESDeviceType(i).RepeatUntil_A
    d.SendCancel_A = ESDeviceType(i).SendCancel_A
    
    d.OG1 = ESDeviceType(i).OG1
    d.OG2 = ESDeviceType(i).OG2
    d.OG3 = ESDeviceType(i).OG3
    d.OG4 = ESDeviceType(i).OG4
    d.OG5 = ESDeviceType(i).OG5
    d.OG6 = ESDeviceType(i).OG6
    
    d.OG1D = ESDeviceType(i).OG1D
    d.OG2D = ESDeviceType(i).OG2D
    d.OG3D = ESDeviceType(i).OG3D
    d.OG4D = ESDeviceType(i).OG4D
    d.OG5D = ESDeviceType(i).OG5D
    d.OG6D = ESDeviceType(i).OG6D
    
    
    d.NG1 = ESDeviceType(i).NG1
    d.NG2 = ESDeviceType(i).NG2
    d.NG3 = ESDeviceType(i).NG3
    d.NG4 = ESDeviceType(i).NG4
    d.NG5 = ESDeviceType(i).NG5
    d.NG6 = ESDeviceType(i).NG6

    d.NG1D = ESDeviceType(i).NG1D
    d.NG2D = ESDeviceType(i).NG2D
    d.NG3D = ESDeviceType(i).NG3D
    d.NG4D = ESDeviceType(i).NG4D
    d.NG5D = ESDeviceType(i).NG5D
    d.NG6D = ESDeviceType(i).NG6D


    d.OG1_A = ESDeviceType(i).OG1_A
    d.OG2_A = ESDeviceType(i).OG2_A
    d.OG3_A = ESDeviceType(i).OG3_A
    d.OG4_A = ESDeviceType(i).OG4_A
    d.OG5_A = ESDeviceType(i).OG5_A
    d.OG6_A = ESDeviceType(i).OG6_A
    
    d.OG1_AD = ESDeviceType(i).OG1_AD
    d.OG2_AD = ESDeviceType(i).OG2_AD
    d.OG3_AD = ESDeviceType(i).OG3_AD
    d.OG4_AD = ESDeviceType(i).OG4_AD
    d.OG5_AD = ESDeviceType(i).OG5_AD
    d.OG6_AD = ESDeviceType(i).OG6_AD
    
    
    d.NG1_A = ESDeviceType(i).NG1_A
    d.NG2_A = ESDeviceType(i).NG2_A
    d.NG3_A = ESDeviceType(i).NG3_A
    d.NG4_A = ESDeviceType(i).NG4_A
    d.NG5_A = ESDeviceType(i).NG5_A
    d.NG6_A = ESDeviceType(i).NG6_A
    
    
    d.NG1_AD = ESDeviceType(i).NG1_AD
    d.NG2_AD = ESDeviceType(i).NG2_AD
    d.NG3_AD = ESDeviceType(i).NG3_AD
    d.NG4_AD = ESDeviceType(i).NG4_AD
    d.NG5_AD = ESDeviceType(i).NG5_AD
    d.NG6_AD = ESDeviceType(i).NG6_AD
    
   ' Stop
    'need third shift
    d.GG1 = ESDeviceType(i).GG1
    d.GG2 = ESDeviceType(i).GG2
    d.GG3 = ESDeviceType(i).GG3
    d.GG4 = ESDeviceType(i).GG4
    d.GG5 = ESDeviceType(i).GG5
    d.GG6 = ESDeviceType(i).GG6
    
    d.GG1D = ESDeviceType(i).GG1D
    d.GG2D = ESDeviceType(i).GG2D
    d.GG3D = ESDeviceType(i).GG3D
    d.GG4D = ESDeviceType(i).GG4D
    d.GG5D = ESDeviceType(i).GG5D
    d.GG6D = ESDeviceType(i).GG6D
    
    d.GG1_A = ESDeviceType(i).GG1_A
    d.GG2_A = ESDeviceType(i).GG2_A
    d.GG3_A = ESDeviceType(i).GG3_A
    d.GG4_A = ESDeviceType(i).GG4_A
    d.GG5_A = ESDeviceType(i).GG5_A
    d.GG6_A = ESDeviceType(i).GG6_A
    
    
    d.GG1_AD = ESDeviceType(i).GG1_AD
    d.GG2_AD = ESDeviceType(i).GG2_AD
    d.GG3_AD = ESDeviceType(i).GG3_AD
    d.GG4_AD = ESDeviceType(i).GG4_AD
    d.GG5_AD = ESDeviceType(i).GG5_AD
    d.GG6_AD = ESDeviceType(i).GG6_AD
    
    DeviceTypes.Add d
    AddToCombo cboDeviceType, d.Model, d.CLSPTI ' MIDPTI
    
    End If

  Next
  ReFreshDeviceTypes  ' local working copy


  cboDeviceType.ListIndex = CboGetIndexByItemData(cboDeviceType, CLSPTI)


End Sub

Sub FillCombos()

        Dim rs            As Recordset
        'Set rs = connexecute("SELECT * FROM Devicetypes ORDER BY model")
        Dim j             As Integer



10      For j = 0 To MAX_ESDEVICETYPES
          If Len((ESDeviceType(j).Model)) Then
20           AddToCombo cboDeviceType, ESDeviceType(j).Model, ESDeviceType(j).CLSPTI  ' MIDPTI
          End If
          'rs.MoveNext
30      Next
        'rs.Close
40      If cboDeviceType.listcount > 0 Then
50        cboDeviceType.ListIndex = 0
60      End If

70      Set rs = ConnExecute("SELECT * FROM pagergroups ORDER BY groupname")
        '  AddToCombo cboGroup1, "< none > ", 0
        '  AddToCombo cboGroup2, "< none > ", 0
        '  AddToCombo cboGroupN1, "< none > ", 0
        '  AddToCombo cboGroupN2, "< none > ", 0
        '
        '  AddToCombo cboGroup1_A, "< none > ", 0
        '  AddToCombo cboGroup2_A, "< none > ", 0
        '  AddToCombo cboGroupN1_A, "< none > ", 0
        '  AddToCombo cboGroupN2_A, "< none > ", 0
80      AddToCombo cboGroup1, "< none > ", 0
90      AddToCombo cboGroup2, "< none > ", 0
100     AddToCombo cboGroup3, "< none > ", 0
110     AddToCombo cboGroup4, "< none > ", 0
120     AddToCombo cboGroup5, "< none > ", 0
130     AddToCombo cboGroup6, "< none > ", 0

140     AddToCombo cboGroupN1, "< none > ", 0
150     AddToCombo cboGroupN2, "< none > ", 0
160     AddToCombo cboGroupN3, "< none > ", 0
170     AddToCombo cboGroupN4, "< none > ", 0
180     AddToCombo cboGroupN5, "< none > ", 0
190     AddToCombo cboGroupN6, "< none > ", 0


200     AddToCombo cboGroup1_A, "< none > ", 0
210     AddToCombo cboGroup2_A, "< none > ", 0
220     AddToCombo cboGroup3_A, "< none > ", 0
230     AddToCombo cboGroup4_A, "< none > ", 0
240     AddToCombo cboGroup5_A, "< none > ", 0
250     AddToCombo cboGroup6_A, "< none > ", 0



260     AddToCombo cboGroupN1_A, "< none > ", 0
270     AddToCombo cboGroupN2_A, "< none > ", 0
280     AddToCombo cboGroupN3_A, "< none > ", 0
290     AddToCombo cboGroupN4_A, "< none > ", 0
300     AddToCombo cboGroupN5_A, "< none > ", 0
310     AddToCombo cboGroupN6_A, "< none > ", 0

320     AddToCombo cboGroupG1, "< none > ", 0
330     AddToCombo cboGroupG2, "< none > ", 0
340     AddToCombo cboGroupG3, "< none > ", 0
350     AddToCombo cboGroupG4, "< none > ", 0
360     AddToCombo cboGroupG5, "< none > ", 0
370     AddToCombo cboGroupG6, "< none > ", 0


380     AddToCombo cboGroupG1_A, "< none > ", 0
390     AddToCombo cboGroupG2_A, "< none > ", 0
400     AddToCombo cboGroupG3_A, "< none > ", 0
410     AddToCombo cboGroupG4_A, "< none > ", 0
420     AddToCombo cboGroupG5_A, "< none > ", 0
430     AddToCombo cboGroupG6_A, "< none > ", 0




440     Do Until rs.EOF
450       AddToCombo cboGroup1, rs("description") & "", rs("groupID")
460       AddToCombo cboGroup2, rs("description") & "", rs("groupID")
470       AddToCombo cboGroup3, rs("description") & "", rs("groupID")
480       AddToCombo cboGroup4, rs("description") & "", rs("groupID")
490       AddToCombo cboGroup5, rs("description") & "", rs("groupID")
500       AddToCombo cboGroup6, rs("description") & "", rs("groupID")



510       AddToCombo cboGroupN1, rs("description") & "", rs("groupID")
520       AddToCombo cboGroupN2, rs("description") & "", rs("groupID")
530       AddToCombo cboGroupN3, rs("description") & "", rs("groupID")
540       AddToCombo cboGroupN4, rs("description") & "", rs("groupID")
550       AddToCombo cboGroupN5, rs("description") & "", rs("groupID")
560       AddToCombo cboGroupN6, rs("description") & "", rs("groupID")



570       AddToCombo cboGroup1_A, rs("description") & "", rs("groupID")
580       AddToCombo cboGroup2_A, rs("description") & "", rs("groupID")
590       AddToCombo cboGroup3_A, rs("description") & "", rs("groupID")
600       AddToCombo cboGroup4_A, rs("description") & "", rs("groupID")
610       AddToCombo cboGroup5_A, rs("description") & "", rs("groupID")
620       AddToCombo cboGroup6_A, rs("description") & "", rs("groupID")

630       AddToCombo cboGroupN1_A, rs("description") & "", rs("groupID")
640       AddToCombo cboGroupN2_A, rs("description") & "", rs("groupID")
650       AddToCombo cboGroupN3_A, rs("description") & "", rs("groupID")
660       AddToCombo cboGroupN4_A, rs("description") & "", rs("groupID")
670       AddToCombo cboGroupN5_A, rs("description") & "", rs("groupID")
680       AddToCombo cboGroupN6_A, rs("description") & "", rs("groupID")


690       AddToCombo cboGroupG1, rs("description") & "", rs("groupID")
700       AddToCombo cboGroupG2, rs("description") & "", rs("groupID")
710       AddToCombo cboGroupG3, rs("description") & "", rs("groupID")
720       AddToCombo cboGroupG4, rs("description") & "", rs("groupID")
730       AddToCombo cboGroupG5, rs("description") & "", rs("groupID")
740       AddToCombo cboGroupG6, rs("description") & "", rs("groupID")

750       AddToCombo cboGroupG1_A, rs("description") & "", rs("groupID")
760       AddToCombo cboGroupG2_A, rs("description") & "", rs("groupID")
770       AddToCombo cboGroupG3_A, rs("description") & "", rs("groupID")
780       AddToCombo cboGroupG4_A, rs("description") & "", rs("groupID")
790       AddToCombo cboGroupG5_A, rs("description") & "", rs("groupID")
800       AddToCombo cboGroupG6_A, rs("description") & "", rs("groupID")


810       rs.MoveNext
820     Loop
830     rs.Close
840     Set rs = Nothing
End Sub


Private Sub Form_Initialize()
  Set DeviceTypes = New Collection

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
  Connect
  cmdGlobalMain.Visible = MASTER
  ArrangeControls
  Do Until TabStrip.Tabs.Count = 0
    TabStrip.Tabs.Remove 1
  Loop
  TabStrip.Tabs.Add 1, "main", "Main"
  
  FillCombos
  
  Fill
  If cboDeviceType.listcount > 0 Then
    cboDeviceType.ListIndex = 0
    Editmode = 0
  Else
    Editmode = 1
  End If
End Sub

Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  SetParent fraEnabler.hwnd, hwnd
  chkUpdateCheckin.Visible = False
  
End Sub

Private Sub ReFreshDeviceTypes()
        Dim SQl           As String
        Dim rs            As Recordset
        Dim d             As cDeviceType
10      For Each d In DeviceTypes
          '    If d.Model = "EN1223D" Then Stop
20        SQl = "SELECT * FROM Devicetypes WHERE Model = " & q(d.Model)
30        Set rs = ConnExecute(SQl)
40        If Not rs.EOF Then

50          d.AllowDisable = IIf(rs("AllowDisable") = 1, 1, 0)
60          d.IsLatching = IIf(rs("IsLatching") = 1, 1, 0)
70          d.Checkin = Val("" & rs("checkin"))
80          d.ClearByReset = IIf(rs("ClearByReset") = 1, 1, 0)
90          d.IsPortable = IIf(rs("IsPortable") = 1, 1, 0)
100         d.Description = rs("description") & ""
110         d.Announce = rs("Announce") & ""
120         d.Announce2 = rs("Announce2") & ""

            'new with build 226
130         d.Repeats = Val("" & rs("Repeats"))
140         d.Pause = Val("" & rs("Pause"))
150         d.RepeatUntil = IIf(rs("repeatuntil") = 1, 1, 0)
160         d.SendCancel = IIf(rs("SendCancel") = 1, 1, 0)

170         d.AutoClear = Val(rs("AutoClear") & "")



180         d.Repeats_A = Val("" & rs("Repeats_A"))
190         d.Pause_A = Val("" & rs("Pause_A"))
200         d.RepeatUntil_A = IIf(rs("repeatuntil_A") = 1, 1, 0)
210         d.SendCancel_A = IIf(rs("SendCancel_A") = 1, 1, 0)

220         d.IgnoreTamper = IIf(rs("ignoretamper") = 1, 1, 0)

230         d.OG1 = Val(rs("OG1") & "")
240         d.OG2 = Val(rs("OG2") & "")
250         d.OG3 = Val(rs("OG3") & "")
260         d.OG4 = Val(rs("OG4") & "")
270         d.OG5 = Val(rs("OG5") & "")
280         d.OG6 = Val(rs("OG6") & "")

290         d.OG1D = Val(rs("OG1D") & "")
300         d.OG2D = Val(rs("OG2D") & "")
310         d.OG3D = Val(rs("OG3D") & "")
320         d.OG4D = Val(rs("OG4D") & "")
330         d.OG5D = Val(rs("OG5D") & "")
340         d.OG6D = Val(rs("OG6D") & "")


350         d.NG1 = Val(rs("NG1") & "")
360         d.NG2 = Val(rs("NG2") & "")
370         d.NG3 = Val(rs("NG3") & "")
380         d.NG4 = Val(rs("NG4") & "")
390         d.NG5 = Val(rs("NG5") & "")
400         d.NG6 = Val(rs("NG6") & "")

410         d.NG1D = Val(rs("NG1D") & "")
420         d.NG2D = Val(rs("NG2D") & "")
430         d.NG3D = Val(rs("NG3D") & "")
440         d.NG4D = Val(rs("NG4D") & "")
450         d.NG5D = Val(rs("NG5D") & "")
460         d.NG6D = Val(rs("NG6D") & "")

470         d.GG1 = Val(rs("GG1") & "")
480         d.GG2 = Val(rs("GG2") & "")
490         d.GG3 = Val(rs("GG3") & "")
500         d.GG4 = Val(rs("GG4") & "")
510         d.GG5 = Val(rs("GG5") & "")
520         d.GG6 = Val(rs("GG6") & "")

530         d.GG1D = Val(rs("GG1D") & "")
540         d.GG2D = Val(rs("GG2D") & "")
550         d.GG3D = Val(rs("GG3D") & "")
560         d.GG4D = Val(rs("GG4D") & "")
570         d.GG5D = Val(rs("GG5D") & "")
580         d.GG6D = Val(rs("GG6D") & "")

590         d.OG1_A = Val(rs("OG1_A") & "")
600         d.OG2_A = Val(rs("OG2_A") & "")
610         d.OG3_A = Val(rs("OG3_A") & "")
620         d.OG4_A = Val(rs("OG4_A") & "")
630         d.OG5_A = Val(rs("OG5_A") & "")
640         d.OG6_A = Val(rs("OG6_A") & "")


650         d.OG1_AD = Val(rs("OG1_AD") & "")
660         d.OG2_AD = Val(rs("OG2_AD") & "")
670         d.OG3_AD = Val(rs("OG3_AD") & "")
680         d.OG4_AD = Val(rs("OG4_AD") & "")
690         d.OG5_AD = Val(rs("OG5_AD") & "")
700         d.OG6_AD = Val(rs("OG6_AD") & "")

710         d.NG1_A = Val(rs("NG1_A") & "")
720         d.NG2_A = Val(rs("NG2_A") & "")
730         d.NG3_A = Val(rs("NG3_A") & "")
740         d.NG4_A = Val(rs("NG4_A") & "")
750         d.NG5_A = Val(rs("NG5_A") & "")
760         d.NG6_A = Val(rs("NG6_A") & "")

770         d.NG1_AD = Val(rs("NG1_Ad") & "")
780         d.NG2_AD = Val(rs("NG2_Ad") & "")
790         d.NG3_AD = Val(rs("NG3_Ad") & "")
800         d.NG4_AD = Val(rs("NG4_Ad") & "")
810         d.NG5_AD = Val(rs("NG5_Ad") & "")
820         d.NG6_AD = Val(rs("NG6_Ad") & "")


830         d.GG1_A = Val(rs("GG1_A") & "")
840         d.GG2_A = Val(rs("GG2_A") & "")
850         d.GG3_A = Val(rs("GG3_A") & "")
860         d.GG4_A = Val(rs("GG4_A") & "")
870         d.GG5_A = Val(rs("GG5_A") & "")
880         d.GG6_A = Val(rs("GG6_A") & "")

890         d.GG1_AD = Val(rs("GG1_AD") & "")
900         d.GG2_AD = Val(rs("GG2_AD") & "")
910         d.GG3_AD = Val(rs("GG3_AD") & "")
920         d.GG4_AD = Val(rs("GG4_AD") & "")
930         d.GG5_AD = Val(rs("GG5_AD") & "")
940         d.GG6_AD = Val(rs("GG6_AD") & "")


950       End If
960       rs.Close
970     Next
980     Set rs = Nothing

End Sub

'Private Function IsDupeDevice(ByVal MIDPTI As Long) As Boolean
'  Dim rs As Recordset
'  Set rs = connexecute("SELECT Count(*) FROM devicetypes WHERE MIDPTI = " & MIDPTI)
'  IsDupeDevice = (rs(0) <> 0)
'  rs.Close
'
'End Function
Private Function Save() As Boolean
        Dim rs            As Recordset
        Dim Checkin       As Long

        Dim i             As Integer
10      i = cboDeviceType.ListIndex
20      If i > 0 Then
30        AutoClear = Val(txtAutoClear.text)
40        MIDPTI = cboDeviceType.ItemData(i)
50        CLSPTI = cboDeviceType.ItemData(i)
60        Model = cboDeviceType.text
70        Checkin = Val(txtCheckin.text)
80        Set rs = New ADODB.Recordset

90        rs.Open "SELECT * FROM devicetypes WHERE model = '" & Model & "'", conn, gCursorType, gLockType

100       If rs.EOF Then
110         rs.addnew
120       End If
130       rs("model") = Model
140       rs("Description") = Trim(txtDescription.text)

150       rs("IsLatching") = 0  'chkLatching.value
160       rs("ClearByReset") = chkClearByReset.Value
170       rs("CheckIn") = Val(txtCheckin.text)
180       rs("IsPortable") = chkPortable.Value
190       rs("AllowDisable") = chkAllowDisable.Value
200       rs("Announce") = Trim(txtAnnounce.text)
210       rs("Announce2") = Trim(txtAnnounce2.text)
          'rs("MIDPTI") = MIDPTI
220       rs("MIDPTI") = CLSPTI
230       rs("AutoClear") = AutoClear



          'new with build 226
240       rs("Repeats") = Val(txtRepeats.text)
250       rs("Pause") = Val(txtPause.text)
260       rs("repeatuntil") = chkRepeatUntil.Value
270       rs("SendCancel") = chkSendCancel.Value

280       rs("Repeats_A") = Val(txtRepeats_A.text)
290       rs("Pause_A") = Val(txtPause_A.text)
300       rs("repeatuntil_A") = chkRepeatUntil_A.Value
310       rs("SendCancel_A") = chkSendCancel_A.Value

320       rs("ignoretamper") = chkIgnoreTamper.Value

          ' these are default groups for this device type
330       rs("OG1") = GetComboItemData(cboGroup1)
340       rs("OG2") = GetComboItemData(cboGroup2)
350       rs("OG3") = GetComboItemData(cboGroup3)
360       rs("OG4") = GetComboItemData(cboGroup4)
370       rs("OG5") = GetComboItemData(cboGroup5)
380       rs("OG6") = GetComboItemData(cboGroup6)

390       rs("OG1d") = Val(txtOG1D.text)
400       rs("OG2d") = Val(txtOG2D.text)
410       rs("OG3d") = Val(txtOG3D.text)
420       rs("OG4d") = Val(txtOG4D.text)
430       rs("OG5d") = Val(txtOG5D.text)
440       rs("OG6d") = Val(txtOG6D.text)

450       rs("OG1_A") = GetComboItemData(cboGroup1_A)
460       rs("OG2_A") = GetComboItemData(cboGroup2_A)
470       rs("OG3_A") = GetComboItemData(cboGroup3_A)
480       rs("OG4_A") = GetComboItemData(cboGroup4_A)
490       rs("OG5_A") = GetComboItemData(cboGroup5_A)
500       rs("OG6_A") = GetComboItemData(cboGroup6_A)



510       rs("OG1_Ad") = Val(txtOG1_AD.text)
520       rs("OG2_Ad") = Val(txtOG2_AD.text)
530       rs("OG3_Ad") = Val(txtOG3_AD.text)
540       rs("OG4_Ad") = Val(txtOG4_AD.text)
550       rs("OG5_Ad") = Val(txtOG5_AD.text)
560       rs("OG6_Ad") = Val(txtOG6_AD.text)



570       rs("NG1") = GetComboItemData(cboGroupN1)
580       rs("NG2") = GetComboItemData(cboGroupN2)
590       rs("NG3") = GetComboItemData(cboGroupN3)
600       rs("NG4") = GetComboItemData(cboGroupN4)
610       rs("NG5") = GetComboItemData(cboGroupN5)
620       rs("NG6") = GetComboItemData(cboGroupN6)

630       rs("NG1d") = Val(txtNG1D.text)
640       rs("NG2d") = Val(txtNG2D.text)
650       rs("NG3d") = Val(txtNG3D.text)
660       rs("NG4d") = Val(txtNG4D.text)
670       rs("NG5d") = Val(txtNG5D.text)
680       rs("NG6d") = Val(txtNG6D.text)

690       rs("NG1_A") = GetComboItemData(cboGroupN1_A)
700       rs("NG2_A") = GetComboItemData(cboGroupN2_A)
710       rs("NG3_A") = GetComboItemData(cboGroupN3_A)
720       rs("NG4_A") = GetComboItemData(cboGroupN4_A)
730       rs("NG5_A") = GetComboItemData(cboGroupN5_A)
740       rs("NG6_A") = GetComboItemData(cboGroupN6_A)

750       rs("NG1_Ad") = Val(txtNG1_AD.text)
760       rs("NG2_Ad") = Val(txtNG2_AD.text)
770       rs("NG3_Ad") = Val(txtNG3_AD.text)
780       rs("NG4_Ad") = Val(txtNG4_AD.text)
790       rs("NG5_Ad") = Val(txtNG5_AD.text)
800       rs("NG6_Ad") = Val(txtNG6_AD.text)

810       rs("GG1") = GetComboItemData(cboGroupG1)
820       rs("GG2") = GetComboItemData(cboGroupG2)
830       rs("GG3") = GetComboItemData(cboGroupG3)
840       rs("GG4") = GetComboItemData(cboGroupG4)
850       rs("GG5") = GetComboItemData(cboGroupG5)
860       rs("GG6") = GetComboItemData(cboGroupG6)

870       rs("GG1D") = Val(txtGG1.text)
880       rs("GG2D") = Val(txtGG2.text)
890       rs("GG3D") = Val(txtGG3.text)
900       rs("GG4D") = Val(txtGG4.text)
910       rs("GG5D") = Val(txtGG5.text)
920       rs("GG6D") = Val(txtGG6.text)




930       rs("GG1_A") = GetComboItemData(cboGroupG1_A)
940       rs("GG2_A") = GetComboItemData(cboGroupG2_A)
950       rs("GG3_A") = GetComboItemData(cboGroupG3_A)
960       rs("GG4_A") = GetComboItemData(cboGroupG4_A)
970       rs("GG5_A") = GetComboItemData(cboGroupG5_A)
980       rs("GG6_A") = GetComboItemData(cboGroupG6_A)

990       rs("GG1_AD") = Val(txtGG1_A.text)
1000      rs("GG2_AD") = Val(txtGG2_A.text)
1010      rs("GG3_AD") = Val(txtGG3_A.text)
1020      rs("GG4_AD") = Val(txtGG4_A.text)
1030      rs("GG5_AD") = Val(txtGG5_A.text)
1040      rs("GG6_AD") = Val(txtGG6_A.text)


1050      rs.Update
1060      rs.Close
1070      Save = True

           Dim CheckInTime As Long
           If Len(Model) Then
              If MASTER Then
                If USE6080 Then
                  CheckInTime = Max(MIN_CHECKIN, Val(Me.txtCheckin.text) * 1 * 60)
                  UpdateCheckinTimeByModel Model, CheckInTime
                End If
              End If
            End If
          

1080      ReFreshDeviceTypes  ' local working copy
1090      ReadESDeviceTypes  ' into global system
1100      UpdateDeviceCheckin Model, Checkin  ' working devices
1110      Display
1120    End If
End Function

Private Sub TabStrip_Click()
  Dim Selected  As Object
  Dim Key       As String

  Set Selected = TabStrip.SelectedItem
  If Selected Is Nothing Then
  Else
    Key = Selected.Key
  End If

  Select Case Selected.Key
  
  
    Case "input1"
      
      fraInput1.Visible = True
      fradef.Visible = False
      fraInput2.Visible = False
      fraInput0.Visible = False
      fraGroups1.Visible = False
      fraGroups2.Visible = False
    
    Case "input2"
      fraInput2.Visible = True
      fradef.Visible = False
      fraInput0.Visible = False
      fraInput1.Visible = False
      fraGroups1.Visible = False
      fraGroups2.Visible = False
    
    Case "groups1"
      fraGroups1.Visible = True
      fradef.Visible = False
      fraGroups2.Visible = False
      fraInput1.Visible = False
      fraInput2.Visible = False
      fraInput0.Visible = False

    Case "groups2"
      fraGroups2.Visible = True
      fradef.Visible = False
      fraGroups1.Visible = False
      fraInput1.Visible = False
      fraInput2.Visible = False
      fraInput0.Visible = False


    Case Else
      fradef.Visible = True
      fraInput0.Visible = False
      
      fraInput1.Visible = False
      fraInput2.Visible = False
      fraGroups1.Visible = False
      fraGroups2.Visible = False
  End Select

End Sub

Private Sub txtAnnounce_GotFocus()
  SelAll txtAnnounce
End Sub

Private Sub txtAnnounce2_GotFocus()
  SelAll txtAnnounce2
End Sub

Private Sub txtAutoClear_GotFocus()
  SelAll txtAutoClear
End Sub

Private Sub txtAutoClear_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtAutoClear, KeyAscii, 0, 0, 3, 999)
End Sub

Private Sub txtCheckin_GotFocus()
  SelAll txtCheckin
End Sub

Private Sub txtCheckin_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtCheckin, KeyAscii, 0, 0, 4, 9999)

End Sub

Private Sub txtDescription_GotFocus()
  SelAll txtDescription
End Sub

Private Sub txtModel_GotFocus()
  SelAll txtModel
End Sub

Private Sub txtnG1_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG1_AD, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtNG1D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG1D, KeyAscii, False, 0, 3, 999)
End Sub



Private Sub txtnG2_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG2_AD, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtNG2D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG2D, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtnG3_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG3_AD, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtNG3D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG3D, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtnG4_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG4_AD, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtNG4D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG4D, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtnG5_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG5_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtNG5D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG5D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtnG6_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG6_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtNG6D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtNG6D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG1_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG1_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG1D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG1D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG2_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG2_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG2D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG2D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG3_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG3_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG3D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG3D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG4_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG4_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG4D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG4D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG5_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG5_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG5D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG5D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG6_AD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG6_AD, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtOG6D_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyProcMax(txtOG6D, KeyAscii, False, 0, 3, 999)
End Sub


Private Sub txtPause_GotFocus()
  SelAll txtPause
End Sub

Private Sub txtPause_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtPause, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtRepeats_GotFocus()
  SelAll txtRepeats
End Sub

Private Sub txtRepeats_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtRepeats, KeyAscii, False, 0, 2, 10)
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

Private Function Validate() As Boolean
  Dim s       As String
  

  Validate = False

  'txtModel.text = UCase(txtModel.text)
  's = Trim(txtModel.text)

  If Me.cboDeviceType.ListIndex <= 0 Then
    Beep
    Exit Function
  End If
  
  's = Me.cboDeviceType.text
  
  s = Trim(txtDescription.text)
  
  If Len(s) = 0 Then
    Beep
    messagebox Me, "Description Cannot Be Blank", App.Title, vbInformation
    Exit Function
  End If

  Validate = True

End Function

Public Property Get Busy() As Boolean

  Busy = mBusy
  fraEnabler.Enabled = Not mBusy
  cmdGlobalMain.Enabled = Not mBusy
  cmdExit.Enabled = Not mBusy
  cmdSave.Enabled = Not Busy
End Property

Public Property Let Busy(ByVal Value As Boolean)

  mBusy = Value
  fraEnabler.Enabled = Not mBusy
  cmdGlobalMain.Enabled = Not mBusy
  cmdExit.Enabled = Not mBusy
  cmdSave.Enabled = Not Busy
End Property
