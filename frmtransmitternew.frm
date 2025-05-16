VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmTransmitter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transmitter"
   ClientHeight    =   13275
   ClientLeft      =   4125
   ClientTop       =   2100
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13275
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrEnroller 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9000
      Top             =   1320
   End
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   13515
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin VB.Frame fraPartitions 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   60
         TabIndex        =   243
         Top             =   3240
         Width           =   7515
         Begin VB.Timer tmrSearch 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   0
            Top             =   0
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
            Left            =   3660
            MaxLength       =   50
            TabIndex        =   37
            Top             =   120
            Width           =   2370
         End
         Begin VB.CommandButton cmdCreatePartion 
            Caption         =   "Manage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6240
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Add new device"
            Top             =   120
            Width           =   1175
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add <"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2760
            TabIndex        =   40
            Top             =   660
            Width           =   615
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "> Del"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2760
            TabIndex        =   41
            Top             =   1395
            Width           =   615
         End
         Begin MSComctlLib.ListView lvPartitions 
            Height          =   1455
            Left            =   60
            TabIndex        =   39
            Top             =   600
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Desc"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lvAvailPartitions 
            Height          =   1995
            Left            =   3480
            TabIndex        =   42
            Top             =   600
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3519
            View            =   3
            LabelEdit       =   1
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
         Begin VB.Label Label41 
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
            Left            =   2880
            TabIndex        =   245
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblMembers 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Device Partitions"
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
            Left            =   120
            TabIndex        =   244
            Top             =   195
            Width           =   1470
         End
      End
      Begin VB.Frame fraassur 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   30
         TabIndex        =   88
         Top             =   3180
         Visible         =   0   'False
         Width           =   7020
         Begin VB.OptionButton optAssur2 
            Caption         =   "Input 2"
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
            Left            =   705
            TabIndex        =   91
            Top             =   870
            Width           =   1545
         End
         Begin VB.OptionButton optAssur1 
            Caption         =   "Input 1"
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
            Left            =   705
            TabIndex        =   90
            Top             =   405
            Width           =   1530
         End
         Begin VB.CheckBox chkAssurance2 
            Alignment       =   1  'Right Justify
            Caption         =   "Check-in Period 2"
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
            Left            =   4785
            TabIndex        =   94
            ToolTipText     =   "Check this box to use as Assurance Device"
            Top             =   1125
            Width           =   2100
         End
         Begin VB.CheckBox chkAssurance 
            Alignment       =   1  'Right Justify
            Caption         =   "Check-in Period 1"
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
            Left            =   4785
            TabIndex        =   92
            ToolTipText     =   "Check this box to use as Assurance Device"
            Top             =   435
            Width           =   2100
         End
         Begin VB.Label lblAssurInput 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check-in Input"
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
            Left            =   750
            TabIndex        =   89
            Top             =   135
            Width           =   1260
         End
         Begin VB.Label lblAssurTime2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "          "
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
            Left            =   4860
            TabIndex        =   95
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label lblAssurTime1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "          "
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
            Left            =   4875
            TabIndex        =   93
            Top             =   765
            Width           =   615
         End
      End
      Begin VB.Frame fraOutput 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         TabIndex        =   43
         Top             =   6660
         Visible         =   0   'False
         Width           =   7560
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
            Left            =   5445
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   60
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
            Left            =   5445
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   375
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
            Left            =   5445
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   1650
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
            Left            =   5445
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1335
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
            Left            =   5445
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   1020
            Width           =   1425
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
            Left            =   5445
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   705
            Width           =   1425
         End
         Begin VB.TextBox txtGG6d 
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
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   80
            ToolTipText     =   "Escalation Timeout"
            Top             =   1665
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtGG5d 
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
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   78
            ToolTipText     =   "Escalation Timeout"
            Top             =   1350
            Width           =   585
         End
         Begin VB.TextBox txtGG4d 
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
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   76
            ToolTipText     =   "Escalation Timeout"
            Top             =   1035
            Width           =   585
         End
         Begin VB.TextBox txtGG3d 
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
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   74
            Top             =   705
            Width           =   585
         End
         Begin VB.TextBox txtGG2d 
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
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   72
            ToolTipText     =   "Escalation Timeout"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox txtGG1D 
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
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   70
            ToolTipText     =   "Escalation Timeout"
            Top             =   75
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   58
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   60
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   62
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   64
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   66
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   68
            ToolTipText     =   "Escalation Timeout"
            Top             =   1650
            Visible         =   0   'False
            Width           =   585
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
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   46
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
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   48
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
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   50
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
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   52
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
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   54
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
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   56
            ToolTipText     =   "Escalation Timeout"
            Top             =   1620
            Visible         =   0   'False
            Width           =   585
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   690
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1005
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1320
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1635
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
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   690
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
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1005
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
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1320
            Width           =   1425
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
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1635
            Width           =   1425
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
            Left            =   3630
            TabIndex        =   82
            ToolTipText     =   "Check this box to send Cancel announcement"
            Top             =   1980
            Width           =   2025
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
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   360
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
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   45
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   75
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   390
            Width           =   1425
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
            Left            =   5235
            MaxLength       =   3
            TabIndex        =   84
            Top             =   2325
            Width           =   510
         End
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
            Left            =   2715
            MaxLength       =   2
            TabIndex        =   83
            Top             =   2325
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
            Left            =   1050
            TabIndex        =   81
            ToolTipText     =   "Check this box to repeat announcements until cleared"
            Top             =   1980
            Width           =   1995
         End
         Begin VB.Label labelg 
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
            Index           =   30
            Left            =   5145
            TabIndex        =   235
            Top             =   105
            Width           =   285
         End
         Begin VB.Label labelg 
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
            Index           =   31
            Left            =   5295
            TabIndex        =   234
            Top             =   420
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   32
            Left            =   5295
            TabIndex        =   233
            Top             =   750
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   33
            Left            =   5295
            TabIndex        =   232
            Top             =   1065
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   34
            Left            =   5280
            TabIndex        =   231
            Top             =   1395
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   35
            Left            =   5295
            TabIndex        =   230
            Top             =   1710
            Width           =   135
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "Input 1"
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
            Left            =   180
            TabIndex        =   229
            Top             =   2340
            Width           =   690
         End
         Begin VB.Label labelg 
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
            Index           =   23
            Left            =   255
            TabIndex        =   228
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   29
            Left            =   2820
            TabIndex        =   227
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   28
            Left            =   2805
            TabIndex        =   226
            Top             =   1380
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   27
            Left            =   2820
            TabIndex        =   225
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   26
            Left            =   2820
            TabIndex        =   224
            Top             =   735
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   25
            Left            =   2820
            TabIndex        =   223
            Top             =   405
            Width           =   135
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esc E"
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
            Left            =   1110
            TabIndex        =   222
            Top             =   1725
            Width           =   510
         End
         Begin VB.Label labelg 
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
            Index           =   22
            Left            =   270
            TabIndex        =   221
            Top             =   1410
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   21
            Left            =   285
            TabIndex        =   220
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   20
            Left            =   285
            TabIndex        =   219
            Top             =   765
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   19
            Left            =   285
            TabIndex        =   218
            Top             =   435
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   24
            Left            =   2625
            TabIndex        =   85
            Top             =   90
            Width           =   330
         End
         Begin VB.Label labelg 
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
            Index           =   18
            Left            =   150
            TabIndex        =   44
            Top             =   120
            Width           =   270
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
            Index           =   0
            Left            =   1860
            TabIndex        =   86
            Top             =   2385
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
            Index           =   0
            Left            =   3360
            TabIndex        =   87
            Top             =   2385
            Width           =   1740
         End
      End
      Begin VB.Frame fraTimes 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2700
         Left            =   -120
         TabIndex        =   31
         Top             =   5220
         Width           =   7455
         Begin VB.Frame fraIgnore 
            BorderStyle     =   0  'None
            Height          =   1305
            Left            =   4470
            TabIndex        =   208
            Top             =   240
            Width           =   2985
            Begin VB.CheckBox chkIgnore 
               Caption         =   "Ignore This device (ALL)"
               CausesValidation=   0   'False
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
               Left            =   120
               TabIndex        =   209
               Top             =   480
               Width           =   2715
            End
         End
         Begin VB.Frame fraDisable_B 
            BorderStyle     =   0  'None
            Caption         =   "Disable"
            Height          =   825
            Left            =   1620
            TabIndex        =   270
            Top             =   600
            Width           =   1545
            Begin VB.TextBox txtDisableEnd_B 
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
               Left            =   30
               MaxLength       =   2
               TabIndex        =   272
               Top             =   420
               Width           =   525
            End
            Begin VB.TextBox txtDisableStart_B 
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
               Left            =   30
               MaxLength       =   2
               TabIndex        =   271
               Top             =   60
               Width           =   525
            End
            Begin VB.Label lblendHr_B 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AM"
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
               Left            =   630
               TabIndex        =   274
               Top             =   480
               Width           =   285
            End
            Begin VB.Label lblStartHr_B 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AM"
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
               Left            =   630
               TabIndex        =   273
               Top             =   120
               Width           =   285
            End
         End
         Begin VB.Frame fraDisable_A 
            BorderStyle     =   0  'None
            Caption         =   "Disable"
            Height          =   825
            Left            =   1620
            TabIndex        =   203
            Top             =   630
            Width           =   1545
            Begin VB.TextBox txtDisableStart_A 
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
               Left            =   30
               MaxLength       =   2
               TabIndex        =   205
               Top             =   60
               Width           =   525
            End
            Begin VB.TextBox txtDisableEnd_A 
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
               Left            =   30
               MaxLength       =   2
               TabIndex        =   204
               Top             =   420
               Width           =   525
            End
            Begin VB.Label lblStartHr_A 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AM"
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
               Left            =   630
               TabIndex        =   207
               Top             =   120
               Width           =   285
            End
            Begin VB.Label lblEndHr_A 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AM"
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
               Left            =   630
               TabIndex        =   206
               Top             =   480
               Width           =   285
            End
         End
         Begin VB.Frame fraDisable 
            BorderStyle     =   0  'None
            Caption         =   "Disable"
            Height          =   825
            Left            =   1620
            TabIndex        =   195
            Top             =   630
            Width           =   1545
            Begin VB.TextBox txtDisableEnd 
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
               Left            =   30
               MaxLength       =   2
               TabIndex        =   199
               Top             =   420
               Width           =   525
            End
            Begin VB.TextBox txtDisableStart 
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
               Left            =   30
               MaxLength       =   2
               TabIndex        =   197
               Top             =   60
               Width           =   525
            End
            Begin VB.Label lblEndHr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AM"
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
               Left            =   630
               TabIndex        =   202
               Top             =   480
               Width           =   285
            End
            Begin VB.Label lblStartHr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AM"
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
               Left            =   630
               TabIndex        =   201
               Top             =   120
               Width           =   285
            End
         End
         Begin VB.Label lblHrsPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "( 0 to 23 hr)"
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
            Left            =   3165
            TabIndex        =   34
            Top             =   765
            Width           =   1020
         End
         Begin VB.Label lblHrsPrompt2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "( 0 to 23 hr)"
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
            Left            =   3165
            TabIndex        =   36
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblAssurStop 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Hour"
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
            Left            =   780
            TabIndex        =   35
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblAssurStart 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Hour"
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
            Left            =   705
            TabIndex        =   33
            Top             =   765
            Width           =   885
         End
         Begin VB.Label lblActive 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the time this Alarm is disabled"
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
            Left            =   780
            TabIndex        =   32
            Top             =   405
            Width           =   3135
         End
      End
      Begin VB.Frame fraOutput_A 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         TabIndex        =   101
         Top             =   8820
         Visible         =   0   'False
         Width           =   7575
         Begin VB.TextBox txtGG1_Ad 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   127
            ToolTipText     =   "Escalation Timeout"
            Top             =   15
            Width           =   585
         End
         Begin VB.TextBox txtGG2_Ad 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   129
            ToolTipText     =   "Escalation Timeout"
            Top             =   330
            Width           =   585
         End
         Begin VB.TextBox txtGG3_Ad 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   131
            Top             =   645
            Width           =   585
         End
         Begin VB.TextBox txtGG4_Ad 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   133
            ToolTipText     =   "Escalation Timeout"
            Top             =   975
            Width           =   585
         End
         Begin VB.TextBox txtGG5_Ad 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   135
            ToolTipText     =   "Escalation Timeout"
            Top             =   1290
            Width           =   585
         End
         Begin VB.TextBox txtGG6_Ad 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   137
            ToolTipText     =   "Escalation Timeout"
            Top             =   1605
            Visible         =   0   'False
            Width           =   585
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   645
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   132
            Top             =   960
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   134
            Top             =   1275
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   1590
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   315
            Width           =   1425
         End
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   0
            Width           =   1425
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   115
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   117
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   119
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   121
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   123
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   125
            ToolTipText     =   "Escalation Timeout"
            Top             =   1650
            Visible         =   0   'False
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   103
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   105
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   107
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   109
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   111
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   113
            ToolTipText     =   "Escalation Timeout"
            Top             =   1620
            Visible         =   0   'False
            Width           =   585
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   660
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   975
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   1290
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   1605
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   690
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   1005
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   1320
            Width           =   1425
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   1635
            Width           =   1425
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
            Left            =   3570
            TabIndex        =   139
            ToolTipText     =   "Check this box to send Cancel announcement"
            Top             =   1950
            Width           =   2115
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   360
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   45
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   45
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   360
            Width           =   1425
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
            Left            =   5265
            MaxLength       =   3
            TabIndex        =   141
            Top             =   2325
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
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   140
            Top             =   2325
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
            Left            =   930
            TabIndex        =   138
            ToolTipText     =   "Check this box to repeat announcements until cleared"
            Top             =   1950
            Width           =   1995
         End
         Begin VB.Label labelg 
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
            Index           =   53
            Left            =   5310
            TabIndex        =   242
            Top             =   1650
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   52
            Left            =   5295
            TabIndex        =   241
            Top             =   1335
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   47
            Left            =   5310
            TabIndex        =   240
            Top             =   1005
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   46
            Left            =   5310
            TabIndex        =   239
            Top             =   690
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   45
            Left            =   5310
            TabIndex        =   238
            Top             =   360
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   44
            Left            =   5160
            TabIndex        =   237
            Top             =   45
            Width           =   285
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Input 2"
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
            Left            =   165
            TabIndex        =   236
            Top             =   2340
            Width           =   690
         End
         Begin VB.Label labelg 
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
            Index           =   51
            Left            =   2730
            TabIndex        =   217
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   50
            Left            =   2715
            TabIndex        =   216
            Top             =   1365
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   43
            Left            =   2730
            TabIndex        =   215
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   42
            Left            =   2730
            TabIndex        =   214
            Top             =   720
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   49
            Left            =   225
            TabIndex        =   213
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   48
            Left            =   210
            TabIndex        =   212
            Top             =   1365
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   39
            Left            =   225
            TabIndex        =   211
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   38
            Left            =   225
            TabIndex        =   210
            Top             =   720
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   41
            Left            =   2730
            TabIndex        =   187
            Top             =   405
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   40
            Left            =   2535
            TabIndex        =   186
            Top             =   90
            Width           =   330
         End
         Begin VB.Label labelg 
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
            Index           =   36
            Left            =   90
            TabIndex        =   145
            Top             =   90
            Width           =   270
         End
         Begin VB.Label labelg 
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
            Index           =   37
            Left            =   225
            TabIndex        =   144
            Top             =   405
            Width           =   135
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
            Index           =   1
            Left            =   1710
            TabIndex        =   143
            Top             =   2370
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
            Index           =   1
            Left            =   3390
            TabIndex        =   142
            Top             =   2385
            Width           =   1740
         End
      End
      Begin VB.Frame fraOutput_B 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   60
         TabIndex        =   247
         Top             =   10800
         Visible         =   0   'False
         Width           =   7575
         Begin VB.TextBox txtGG1_BD 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   171
            ToolTipText     =   "Escalation Timeout"
            Top             =   15
            Width           =   585
         End
         Begin VB.TextBox txtGG2_BD 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   173
            ToolTipText     =   "Escalation Timeout"
            Top             =   330
            Width           =   585
         End
         Begin VB.TextBox txtGG3_BD 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   175
            Top             =   645
            Width           =   585
         End
         Begin VB.TextBox txtGG4_BD 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   177
            ToolTipText     =   "Escalation Timeout"
            Top             =   975
            Width           =   585
         End
         Begin VB.TextBox txtGG5_BD 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   179
            ToolTipText     =   "Escalation Timeout"
            Top             =   1290
            Width           =   585
         End
         Begin VB.TextBox txtGG6_BD 
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
            Left            =   6900
            MaxLength       =   3
            TabIndex        =   181
            ToolTipText     =   "Escalation Timeout"
            Top             =   1605
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cboGroupG3_B 
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   645
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG4_B 
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   960
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG5_B 
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   178
            Top             =   1275
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG6_B 
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   1590
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG2_B 
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   172
            Top             =   315
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupG1_B 
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
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   0
            Width           =   1425
         End
         Begin VB.TextBox txtNG1_BD 
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   159
            ToolTipText     =   "Escalation Timeout"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txtNG2_BD 
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   161
            ToolTipText     =   "Escalation Timeout"
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox txtNG3_BD 
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   163
            ToolTipText     =   "Escalation Timeout"
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox txtNG4_BD 
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   165
            ToolTipText     =   "Escalation Timeout"
            Top             =   1020
            Width           =   585
         End
         Begin VB.TextBox txtNG5_BD 
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   167
            ToolTipText     =   "Escalation Timeout"
            Top             =   1335
            Width           =   585
         End
         Begin VB.TextBox txtNG6_BD 
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
            Left            =   4410
            MaxLength       =   3
            TabIndex        =   169
            ToolTipText     =   "Escalation Timeout"
            Top             =   1650
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtOG1_BD 
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   147
            ToolTipText     =   "Escalation Timeout"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txtOG2_BD 
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   149
            ToolTipText     =   "Escalation Timeout"
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox txtOG3_BD 
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   151
            ToolTipText     =   "Escalation Timeout"
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox txtOG4_BD 
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   153
            ToolTipText     =   "Escalation Timeout"
            Top             =   990
            Width           =   585
         End
         Begin VB.TextBox txtOG5_BD 
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   155
            ToolTipText     =   "Escalation Timeout"
            Top             =   1305
            Width           =   585
         End
         Begin VB.TextBox txtOG6_BD 
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
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   157
            ToolTipText     =   "Escalation Timeout"
            Top             =   1620
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cboGroup3_B 
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   660
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup4_B 
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   975
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup5_B 
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   1290
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup6_B 
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   1605
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN3_B 
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   162
            Top             =   690
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN4_B 
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   1005
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN5_B 
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   1320
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN6_B 
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   1635
            Width           =   1425
         End
         Begin VB.CheckBox chkSendCancel_B 
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
            Left            =   3570
            TabIndex        =   183
            ToolTipText     =   "Check this box to send Cancel announcement"
            Top             =   1950
            Width           =   2115
         End
         Begin VB.ComboBox cboGroupN2_B 
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   360
            Width           =   1425
         End
         Begin VB.ComboBox cboGroupN1_B 
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
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   45
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup1_B 
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   45
            Width           =   1425
         End
         Begin VB.ComboBox cboGroup2_B 
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   360
            Width           =   1425
         End
         Begin VB.TextBox txtPause_B 
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
            Left            =   5265
            MaxLength       =   3
            TabIndex        =   185
            Top             =   2325
            Width           =   510
         End
         Begin VB.TextBox txtRepeats_B 
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
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   184
            Top             =   2325
            Width           =   510
         End
         Begin VB.CheckBox chkRepeatUntil_B 
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
            Left            =   930
            TabIndex        =   182
            ToolTipText     =   "Check this box to repeat announcements until cleared"
            Top             =   1950
            Width           =   1995
         End
         Begin VB.Label labelg 
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
            Index           =   17
            Left            =   5310
            TabIndex        =   268
            Top             =   1650
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   16
            Left            =   5295
            TabIndex        =   267
            Top             =   1335
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   15
            Left            =   5310
            TabIndex        =   266
            Top             =   1005
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   14
            Left            =   5310
            TabIndex        =   265
            Top             =   690
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   13
            Left            =   5310
            TabIndex        =   264
            Top             =   360
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   12
            Left            =   5160
            TabIndex        =   263
            Top             =   45
            Width           =   285
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFCCCC&
            Caption         =   "Input 3"
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
            Left            =   165
            TabIndex        =   262
            Top             =   2340
            Width           =   690
         End
         Begin VB.Label labelg 
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
            Index           =   11
            Left            =   2730
            TabIndex        =   261
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   10
            Left            =   2715
            TabIndex        =   260
            Top             =   1365
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   9
            Left            =   2730
            TabIndex        =   259
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   8
            Left            =   2730
            TabIndex        =   258
            Top             =   720
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   4
            Left            =   225
            TabIndex        =   257
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   3
            Left            =   210
            TabIndex        =   256
            Top             =   1365
            Width           =   150
         End
         Begin VB.Label labelg 
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
            Index           =   2
            Left            =   225
            TabIndex        =   255
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   1
            Left            =   225
            TabIndex        =   254
            Top             =   720
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   7
            Left            =   2730
            TabIndex        =   253
            Top             =   405
            Width           =   135
         End
         Begin VB.Label labelg 
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
            Index           =   6
            Left            =   2535
            TabIndex        =   252
            Top             =   90
            Width           =   330
         End
         Begin VB.Label labelg 
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
            Index           =   5
            Left            =   90
            TabIndex        =   251
            Top             =   90
            Width           =   270
         End
         Begin VB.Label labelg 
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
            Index           =   0
            Left            =   225
            TabIndex        =   250
            Top             =   405
            Width           =   135
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
            Index           =   2
            Left            =   1710
            TabIndex        =   249
            Top             =   2370
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
            Index           =   2
            Left            =   3390
            TabIndex        =   248
            Top             =   2385
            Width           =   1740
         End
      End
      Begin VB.Frame fratx 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   30
         TabIndex        =   2
         Top             =   330
         Width           =   7560
         Begin VB.CheckBox chkTamperInput 
            Alignment       =   1  'Right Justify
            Caption         =   "Tamper as Input"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5100
            TabIndex        =   194
            Top             =   1305
            Width           =   2055
         End
         Begin VB.OptionButton optInput3 
            BackColor       =   &H00FFCCCC&
            Caption         =   "Input 3"
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
            Left            =   4980
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   60
            Width           =   780
         End
         Begin VB.ComboBox cboDeviceMode 
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
            Left            =   2940
            TabIndex        =   7
            Text            =   "cboDeviceMode"
            Top             =   480
            Width           =   795
         End
         Begin VB.TextBox txtCustom 
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
            Left            =   3780
            MaxLength       =   35
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   450
            Width           =   3360
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
            Height          =   255
            Left            =   5100
            TabIndex        =   193
            Top             =   1050
            Width           =   2055
         End
         Begin VB.CheckBox chkClearByReset 
            Alignment       =   1  'Right Justify
            Caption         =   "Clear By Reset"
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
            Left            =   5100
            TabIndex        =   192
            ToolTipText     =   "Check this box to require pressing reset to clear alarm"
            Top             =   780
            Width           =   2055
         End
         Begin VB.Frame fraInput 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Caption         =   "Input1"
            Height          =   1095
            Left            =   5070
            TabIndex        =   27
            Top             =   1590
            Width           =   2085
            Begin VB.CheckBox chkExtern 
               Alignment       =   1  'Right Justify
               Caption         =   "Display as Extern"
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
               Left            =   0
               TabIndex        =   198
               ToolTipText     =   "Check this box to display as Ecxternal Alarm"
               Top             =   300
               Width           =   2100
            End
            Begin VB.CheckBox chkVacationSuper 
               Alignment       =   1  'Right Justify
               Caption         =   "Vacation Security"
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
               Left            =   0
               TabIndex        =   200
               ToolTipText     =   "Check this box to use for Vacation Supervision"
               Top             =   630
               Width           =   2100
            End
            Begin VB.CheckBox chkAlarmAlert 
               Alignment       =   1  'Right Justify
               Caption         =   "Display as Alert"
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
               Left            =   0
               TabIndex        =   196
               ToolTipText     =   "Check this box to diaplay as Alert"
               Top             =   0
               Width           =   2100
            End
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
            Left            =   1470
            MaxLength       =   30
            TabIndex        =   13
            Top             =   795
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Frame fraAssignment 
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
            Height          =   1200
            Left            =   60
            TabIndex        =   18
            Top             =   1125
            Width           =   4710
            Begin VB.TextBox txtClearingDevice 
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
               Height          =   315
               Left            =   1185
               MaxLength       =   8
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   810
               Width           =   1365
            End
            Begin VB.CommandButton cmdUnAssignRes 
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
               Height          =   375
               Left            =   3720
               TabIndex        =   25
               ToolTipText     =   "Clear Resident Assignment"
               Top             =   420
               Width           =   930
            End
            Begin VB.TextBox txtAssignRes 
               BackColor       =   &H8000000F&
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
               Left            =   1192
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   435
               Width           =   2445
            End
            Begin VB.TextBox txtAssigned 
               BackColor       =   &H8000000F&
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
               Left            =   1192
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   30
               Width           =   2445
            End
            Begin VB.CommandButton cmdUnassign 
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
               Height          =   375
               Left            =   3720
               TabIndex        =   22
               ToolTipText     =   "Clear Room Assignment"
               Top             =   15
               Width           =   930
            End
            Begin VB.CommandButton cmdAsignResident 
               BackColor       =   &H00FFFF80&
               Caption         =   "Resident"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   15
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Click to Assign Resident to this device"
               Top             =   420
               Width           =   1095
            End
            Begin VB.CommandButton cmdAssignRoom 
               BackColor       =   &H0080FF80&
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
               Height          =   375
               Left            =   15
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Click to Assign Room to this device"
               Top             =   30
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Reset"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   0
               TabIndex        =   275
               Top             =   885
               Width           =   1035
            End
            Begin VB.Label lblAssignType 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " "
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
               Left            =   1875
               TabIndex        =   19
               Top             =   120
               Width           =   75
            End
         End
         Begin VB.CommandButton cmdAutoEnroll 
            Caption         =   "Enroll"
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
            Left            =   5820
            TabIndex        =   17
            ToolTipText     =   "Click to start autoenroll of this device"
            Top             =   75
            Width           =   840
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
            Left            =   1470
            MaxLength       =   30
            TabIndex        =   12
            Top             =   795
            Width           =   3375
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
            Left            =   630
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   450
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox txtSerial 
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
            Left            =   630
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   90
            Width           =   1500
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
            Height          =   315
            Left            =   630
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   465
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.OptionButton optInput1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Input 1"
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
            Left            =   3300
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   75
            Width           =   780
         End
         Begin VB.OptionButton optInput2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Input 2"
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
            Left            =   4140
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   75
            Width           =   780
         End
         Begin VB.CommandButton cmdConfigureSerial 
            Caption         =   "Setup"
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
            Left            =   6720
            TabIndex        =   100
            ToolTipText     =   "Configure Serio I/O parameters"
            Top             =   75
            Width           =   840
         End
         Begin VB.TextBox txtAnnounce3 
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
            Left            =   1470
            MaxLength       =   30
            TabIndex        =   269
            Top             =   795
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Frame fraInput_B 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Input1"
            Height          =   975
            Left            =   5040
            TabIndex        =   29
            Top             =   1560
            Width           =   2145
            Begin VB.CheckBox chkAlarmAlert_B 
               Alignment       =   1  'Right Justify
               Caption         =   "Display as Alert 3"
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
               Left            =   0
               TabIndex        =   191
               ToolTipText     =   "Check this box to diaplay as Alert"
               Top             =   0
               Width           =   2100
            End
         End
         Begin VB.Frame fraInput_A 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Caption         =   "Input2"
            Height          =   1215
            Left            =   5040
            TabIndex        =   28
            Top             =   1710
            Width           =   2115
            Begin VB.CheckBox chkExtern_A 
               Alignment       =   1  'Right Justify
               Caption         =   "Display as Extern 2"
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
               Left            =   0
               TabIndex        =   190
               ToolTipText     =   "Check this box to display as Ecxternal Alarm"
               Top             =   300
               Width           =   2100
            End
            Begin VB.CheckBox chkVacationSuper_A 
               Alignment       =   1  'Right Justify
               Caption         =   "Vacation Security 2"
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
               Left            =   0
               TabIndex        =   189
               ToolTipText     =   "Check this box to use for Vacation Supervision"
               Top             =   630
               Width           =   2100
            End
            Begin VB.CheckBox chkAlarmAlert_A 
               Alignment       =   1  'Right Justify
               Caption         =   "Display as Alert 2"
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
               Left            =   0
               TabIndex        =   188
               ToolTipText     =   "Check this box to diaplay as Alert"
               Top             =   0
               Width           =   2100
            End
         End
         Begin VB.Label lblIDM 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "   "
            Height          =   195
            Left            =   180
            TabIndex        =   246
            Top             =   2400
            Width           =   135
         End
         Begin VB.Label lblTypeDesc 
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
            Left            =   3780
            TabIndex        =   10
            Top             =   495
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label lblAlert 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "."
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
            Top             =   2370
            Width           =   2400
         End
         Begin VB.Label lblAnnounce 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Announce"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   11
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lblSerial 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Serial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -105
            TabIndex        =   3
            Top             =   135
            Width           =   690
         End
         Begin VB.Label lblDesc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
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
            Left            =   120
            TabIndex        =   5
            Top             =   525
            Width           =   435
         End
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
         TabIndex        =   99
         ToolTipText     =   "Exit this screen"
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
         TabIndex        =   98
         ToolTipText     =   "Save changes"
         Top             =   1785
         Width           =   1175
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
         TabIndex        =   97
         TabStop         =   0   'False
         ToolTipText     =   "Add new device"
         Top             =   30
         Width           =   1175
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   3135
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   5530
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Transmitter"
               Key             =   "tx"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Check-in"
               Key             =   "assure"
               Object.ToolTipText     =   "Assurance Settings"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Output Groups"
               Key             =   "output"
               Object.ToolTipText     =   "Set How Events are Paged or Announced"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Times"
               Key             =   "times"
               Object.ToolTipText     =   "Set when this transmitter alarms are active"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Partitions"
               Key             =   "partitions"
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
      Begin VB.Label lblDecimal1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal"
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
         Left            =   11310
         TabIndex        =   96
         Top             =   1395
         Visible         =   0   'False
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmTransmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Partitions  As Collection

Private mDeviceID   As Long

Private RoomID      As Long
Private ResidentID  As Long

Private mDevice   As cESDevice

'Private mInput2     As Boolean

Public InputNumber As Long

Private mDeviceType As ESDeviceTypeType
Private ManagingPartitions As Boolean

Private RegDevice   As cRegDevice

Private bycode      As Boolean

Private regws       As WebSocketSocket

Function EnrollButtonVisible() As Boolean
  EnrollButtonVisible = True
  If (Not MASTER) Then
    If (USE6080) Then
      EnrollButtonVisible = False
    End If
  End If

End Function

Sub LoadRegDevice()
  Dim sn            As String
  Dim dt            As String
  dt = RegDevice.DeviceType

  SetDevice RegDevice.FullHexSerial, RegDevice.CLSPTI
  SetNewDeviceType

End Sub


Sub AdvanceModel()
  If bycode Then Exit Sub
  If cboDeviceType.ListIndex < cboDeviceType.listcount - 1 Then
    cboDeviceType.ListIndex = cboDeviceType.ListIndex + 1
  Else
    If cboDeviceType.listcount > 0 Then
      cboDeviceType.ListIndex = 0
    End If
  End If
End Sub

Public Sub AutoEnroll(p As cESPacket)



10 On Error GoTo AutoEnroll_Error

20 AutoEnrollEnabled = False
  'SetDevice p.Serial, p.MIDPTI
30 SetDevice p.Serial, p.ClassByte * 256& + p.PTI ' 1941 standard is Class 62 dec 0x3E pti 12 dec 0x0C sample packet 7212B26D12390020640E00003E0C000848455F
40 cmdAutoEnroll.Enabled = (mDeviceID = 0) And EnrollButtonVisible()
50 SetNewDeviceType

AutoEnroll_Resume:
60 On Error GoTo 0
70 Exit Sub

AutoEnroll_Error:

80 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.AutoEnroll." & Erl
90 Resume AutoEnroll_Resume

End Sub

Private Sub cboDeviceType_Click()
  Dim index         As Long

  'ClearAlerts
10 On Error GoTo cboDeviceType_Click_Error

20 If Not bycode Then
30  SetNewDeviceType
40  If 0 = StrComp(cboDeviceType.text, COM_DEV_NAME, vbTextCompare) Then
50    cmdConfigureSerial.Visible = True
60    cmdAutoEnroll.Visible = False And EnrollButtonVisible()
70    chkClearByReset.Value = 0
80    chkClearByReset.Visible = False
90    chkExtern.Visible = True
100   chkExtern_A.Visible = True
110 Else
120   cmdAutoEnroll.Visible = True And EnrollButtonVisible()
130   chkClearByReset.Visible = True
140   cmdConfigureSerial.Visible = False
150   cmdAutoEnroll.Enabled = (mDeviceID = 0) And EnrollButtonVisible()
160   chkExtern.Visible = False
170   chkExtern_A.Visible = False
180 End If

    If mDeviceType.NumInputs = 2 Then
        chkTamperInput.Visible = True
    Else
        chkTamperInput.Visible = False
        chkTamperInput.Value = 0
    End If
    
190 End If  ' not by code

  'cmdPagerSettings.Visible = (UCase(mDeviceType.Model) = "ES3954") And (mDevice.DeviceID <> 0)

cboDeviceType_Click_Resume:
200 On Error GoTo 0
210 Exit Sub

cboDeviceType_Click_Error:

220 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.cboDeviceType_Click." & Erl
230 Resume cboDeviceType_Click_Resume

End Sub

Private Sub cboGroup1_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup1_A)
  mDevice.OG1_A = i

End Sub

Private Sub cboGroup1_Click()
  Debug.Print "cboGroup1_Click , bycode "; bycode

  If bycode Then Exit Sub
  Dim i             As Long

  i = GetComboItemData(cboGroup1)
  mDevice.OG1 = i

End Sub

Private Sub cboGroup2_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup2_A)
  mDevice.OG2_A = i
End Sub

Private Sub cboGroup2_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup2)
  mDevice.OG2 = i


End Sub

Private Sub cboGroup3_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup3_A)
  mDevice.OG3_A = i
End Sub

Private Sub cboGroup3_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup3)
  mDevice.OG3 = i
End Sub

Private Sub cboGroup4_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup4_A)
  mDevice.OG4_A = i
End Sub

Private Sub cboGroup4_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup4)
  mDevice.OG4 = i
End Sub

Private Sub cboGroup5_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup5_A)
  mDevice.OG5_A = i
End Sub

Private Sub cboGroup5_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup5)
  mDevice.OG5 = i
End Sub

Private Sub cboGroup6_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup6_A)
  mDevice.OG6_A = i
End Sub

Private Sub cboGroup6_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroup6)
  mDevice.OG6 = i
End Sub

Private Sub cboGroupN1_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN1_A)
  mDevice.NG1_A = i

End Sub

Private Sub cboGroupN1_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN1)
  mDevice.NG1 = i


End Sub

Private Sub cboGroupN2_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN2_A)
  mDevice.NG2_A = i
End Sub

Private Sub cboGroupN2_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN2)
  mDevice.NG2 = i
End Sub

Private Sub cboGroupN3_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN3_A)
  mDevice.NG3_A = i

End Sub

Private Sub cboGroupN3_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN3)
  mDevice.NG3 = i
End Sub

Private Sub cboGroupN4_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN4_A)
  mDevice.NG4_A = i

End Sub

Private Sub cboGroupN4_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN4)
  mDevice.NG4 = i
End Sub

Private Sub cboGroupN5_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN5_A)
  mDevice.NG5_A = i

End Sub

Private Sub cboGroupN5_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN5)
  mDevice.NG5 = i
End Sub

Private Sub cboGroupN6_A_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN6_A)
  mDevice.NG6_A = i

End Sub

Private Sub cboGroupN6_Click()
  If bycode Then Exit Sub
  Dim i             As Long
  i = GetComboItemData(cboGroupN6)
  mDevice.NG6 = i
End Sub

Private Sub chkAlarmAlert_A_Click()
  If bycode Then Exit Sub
  If chkAlarmAlert_A.Value = 1 Then
    mDevice.AlarmMask_A = 1
    If chkExtern_A.Value <> 0 Then
      chkExtern_A.Value = 0
    End If
  ElseIf chkExtern_A.Value = 1 Then
    If chkAlarmAlert_A.Value <> 0 Then
      chkAlarmAlert_A.Value = 0
    End If
    mDevice.AlarmMask_A = 2
  Else
    mDevice.AlarmMask_A = 0
  End If

End Sub

Private Sub chkAlarmAlert_B_Click()
  If bycode Then Exit Sub
  If chkAlarmAlert_B.Value = 1 Then
    mDevice.AlarmMask_B = 1
  Else
    mDevice.AlarmMask_B = 0
  End If


End Sub

Private Sub chkAlarmAlert_Click()
  If bycode Then Exit Sub
  If chkAlarmAlert.Value = 1 Then
    mDevice.AlarmMask = 1
    If chkExtern.Value <> 0 Then
      chkExtern.Value = 0
    End If
  ElseIf chkExtern.Value = 1 Then
    If chkAlarmAlert.Value <> 0 Then
      chkAlarmAlert.Value = 0
    End If
    mDevice.AlarmMask = 2
  Else
    mDevice.AlarmMask = 0
  End If
End Sub

Private Sub chkAssurance_Click()
  If bycode Then Exit Sub
  'mDevice.UseAssur_A = chkAssurance.value
  mDevice.UseAssur = chkAssurance.Value
End Sub

Private Sub chkAssurance2_Click()

  If bycode Then Exit Sub
  'mDevice.UseAssur2_A = chkAssurance2.value
  mDevice.UseAssur2 = chkAssurance2.Value

End Sub

Private Sub chkClearByReset_Click()
  mDevice.ClearByReset = chkClearByReset.Value
End Sub

Private Sub chkExtern_A_Click()
  If chkExtern_A.Value = 1 Then
    If chkAlarmAlert_A.Value <> 0 Then
      chkAlarmAlert_A.Value = 0
    End If
    mDevice.AlarmMask_A = 2

  ElseIf chkAlarmAlert_A.Value = 1 Then
    mDevice.AlarmMask_A = 1
    If chkExtern_A.Value <> 0 Then
      chkExtern_A.Value = 0
    End If

  Else
    mDevice.AlarmMask_A = 0
  End If

End Sub

Private Sub chkExtern_Click()
  If chkExtern.Value = 1 Then
    If chkAlarmAlert.Value <> 0 Then
      chkAlarmAlert.Value = 0
    End If
    mDevice.AlarmMask = 2

  ElseIf chkAlarmAlert.Value = 1 Then
    mDevice.AlarmMask = 1
    If chkExtern.Value <> 0 Then
      chkExtern.Value = 0
    End If

  Else
    mDevice.AlarmMask = 0
  End If


End Sub

Private Sub chkRepeatUntil_A_Click()
  mDevice.RepeatUntil_A = chkRepeatUntil_A.Value
End Sub

Private Sub chkRepeatUntil_Click()
  If bycode Then Exit Sub
  mDevice.RepeatUntil = chkRepeatUntil.Value

End Sub

Private Sub chkSendCancel_A_Click()
  If bycode Then Exit Sub
  mDevice.SendCancel_A = chkSendCancel_A.Value
End Sub

Private Sub chkSendCancel_Click()
  If bycode Then Exit Sub
  On Error Resume Next
  Device.SendCancel = chkSendCancel.Value

End Sub

Private Sub chkTamperInput_Click()
  mDevice.UseTamperAsInput = chkTamperInput.Value And 1
  optInput3.Visible = CBool(mDevice.UseTamperAsInput) And (mDeviceType.NumInputs = 2)
    
  
  'Display
End Sub

Private Sub chkVacationSuper_A_Click()
  If bycode Then Exit Sub
  mDevice.AssurSecure_A = chkVacationSuper_A.Value
End Sub

Private Sub chkVacationSuper_Click()
  If bycode Then Exit Sub
  mDevice.AssurSecure = chkVacationSuper.Value

End Sub

Sub ClearAlerts()
  On Error Resume Next
  lblAlert.Caption = ""
End Sub

Sub ClearResAssignment(ByVal DeviceID As Long)
  Dim SQL           As String
10 On Error GoTo ClearResAssignment_Error

20 SQL = "UPDATE Devices set  ResidentID = 0 WHERE Deviceid = " & DeviceID
30 ConnExecute SQL
40 If MASTER Then
50  Devices.RefreshByID DeviceID
60 End If
70 Fill

ClearResAssignment_Resume:
80 On Error GoTo 0
90 Exit Sub

ClearResAssignment_Error:

100 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.ClearResAssignment." & Erl
110 Resume ClearResAssignment_Resume


End Sub

Sub ClearRoomAssignment(ByVal DeviceID As Long)
  Dim SQL           As String
  SQL = "UPDATE Devices set RoomID = 0 WHERE Deviceid = " & DeviceID
  ConnExecute SQL
  If MASTER Then
    Devices.RefreshByID DeviceID
  End If
  Fill

End Sub

Private Sub cmdAdd_Click()
  AddPartition
End Sub

Function AddPartition()
  Dim currentcount  As Long

  Dim Partitionlist As Collection
  Dim part          As cPartition
  Dim li            As ListItem
  Dim ZoneID        As Long

  ZoneID = mDevice.ZoneID

  If ZoneID = 0 Then
    Beep
    ' throw error
    Exit Function
  End If
  Set Partitionlist = New Collection
  currentcount = lvPartitions.ListItems.Count
  For Each li In lvAvailPartitions.ListItems
    If li.Selected Then
      Set part = New cPartition
      part.PartitionID = Val(li.text)
      Partitionlist.Add part
    End If
  Next

  If (Partitionlist.Count + currentcount > 4) Then
    Beep
  ElseIf (Partitionlist.Count = 0) Then
    Beep
  Else

    Dim XML         As String

    Dim HTTPRequest As cHTTPRequest
    Set HTTPRequest = New cHTTPRequest
    XML = HTTPRequest.AddPartitionList(GetHTTP & "://" & IP1, USER1, PW1, Partitionlist, ZoneID)

    Set HTTPRequest = Nothing
    FillActivePartitionList
  End If




End Function

Private Sub cmdAsignResident_Click()
  ClearAlerts
  On Error Resume Next
  If bycode Then Exit Sub
  If DoSave() Then
    ShowResidents mDevice.ResidentID, DeviceID, "TX"
  End If
End Sub

Private Sub cmdAssignRoom_Click()

10 On Error GoTo cmdAssignRoom_Click_Error

20 ClearAlerts
30 If bycode Then Exit Sub
40 If DoSave() Then
    
50  ShowRooms 0, DeviceID, mDevice.RoomID, "TX"
    On Error Resume Next
  
60 End If

cmdAssignRoom_Click_Resume:
70 On Error GoTo 0
80 Exit Sub

cmdAssignRoom_Click_Error:

90 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.cmdAssignRoom_Click." & Erl
100 Resume cmdAssignRoom_Click_Resume

End Sub

Private Sub cmdAutoEnroll_Click()
  
  If bycode Then Exit Sub
  ClearAlerts  ' clears caption info
  EnableAutoEnroll

End Sub

Private Sub cmdCancel_Click()
  ClearAlerts
  DisableAutoEnroll
  PreviousForm
  Unload Me
End Sub

Private Sub cmdConfigureSerial_Click()

  If Save() Then
    Dim Clone As cESDevice
    Set Clone = mDevice
    'If mDevice.IsTemperatureDev Then
    If 0 = StrComp(cboDeviceType.text, "ES1723", vbTextCompare) Then
      EditTemperatureDevice Clone
    ElseIf mDevice.IsSerialDevice Then
      EditSerialDevice Clone
    End If
  End If
End Sub



Private Sub cmdCreatePartion_Click()
  ManageAvailablePartitions
  ManagingPartitions = True

End Sub

Private Sub cmdNew_Click()
  ClearAlerts
  Set mDevice = New cESDevice
  DeviceID = 0
  ResetForm

End Sub

Private Sub cmdOK_Click()
  DoSave
  On Error Resume Next
  SetFocusTo cmdOK
End Sub
Function DoSave() As Boolean
  Dim Saved              As Boolean

  Me.Enabled = False

  cmdOK.Enabled = False
  cmdCancel.Enabled = False
  cmdAssignRoom.Enabled = False
  cmdAssignRoom.BackColor = Me.BackColor
  cmdAdd.Enabled = False
  cmdAsignResident.Enabled = False
  cmdAsignResident.BackColor = Me.BackColor
  cmdUnassign.Enabled = False
  cmdUnAssignRes.Enabled = False

  Dim Allowed As Long
  Allowed = GetAllowedDeviceCount()
  
  If Devices.Count = Allowed Then
    MsgBox "Cannot Add Device. Maximum Devices Allowed is " & Allowed, vbCritical, "Add Device"
  Else
    Saved = Save()
  End If

  If Saved Then

    cmdAssignRoom.Enabled = True
    cmdAssignRoom.BackColor = &H80FF80
    cmdAsignResident.Enabled = True
    cmdAsignResident.BackColor = &HFFFF80
    cmdUnAssignRes.Enabled = True
    cmdUnassign.Enabled = True
  End If

  cmdAdd.Enabled = True
  cmdCancel.Enabled = True

  Me.Enabled = True
  cmdOK.Enabled = True

  DoSave = Saved
End Function


Private Sub cmdRemove_Click()

  TxRemovePartitions

End Sub
Function TxRemovePartitions() As Long
  Dim Partitions    As Collection
  Dim li            As ListItem
  Dim part          As cPartition

  Dim ZoneID        As Long

  ZoneID = mDevice.ZoneID
  If ZoneID = 0 Then
    Beep
    ' show error
    Exit Function
  End If

  Set Partitions = New Collection



  For Each li In lvPartitions.ListItems
    If li.Selected Then
      Set part = New cPartition
      part.PartitionID = Val(li.text)
      part.Description = li.SubItems(1)
      'part.IsLocation = li.SubItems(2)
      Partitions.Add part
    End If
  Next

  Dim XML           As String
  
  If Partitions.Count Then

    
    Dim HTTPRequest As cHTTPRequest
    Set HTTPRequest = New cHTTPRequest
    XML = HTTPRequest.RemovePartitionList(GetHTTP & "://" & IP1, USER1, PW1, Partitions, ZoneID)

    Set HTTPRequest = Nothing

    FillActivePartitionList

  End If  ' no parts chosen


End Function


Private Sub cmdUnassign_Click()
10 On Error GoTo cmdUnassign_Click_Error

20 If bycode Then Exit Sub

30 'If vbYes = messagebox(frmMain, "Remove Room assignment for this Device?", App.Title, vbYesNo Or vbQuestion Or Win32.MB_TASKMODAL) Then
40 ClearRoomAssignment DeviceID
50 frmMain.SetListTabs
60 Fill
70 'End If

cmdUnassign_Click_Resume:
80 On Error GoTo 0
90 Exit Sub

cmdUnassign_Click_Error:

100 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.cmdUnassign_Click." & Erl
110 Resume cmdUnassign_Click_Resume

End Sub

Private Sub cmdUnAssignRes_Click()
10 On Error GoTo cmdUnAssignRes_Click_Error

20 If bycode Then Exit Sub
30 'If vbYes = messagebox(frmMain, "Remove Resident assignment for this Device?", App.Title, vbYesNo Or vbQuestion Or Win32.MB_TASKMODAL) Then
40 ClearResAssignment DeviceID
50 frmMain.SetListTabs
60 Fill
70 'End If

cmdUnAssignRes_Click_Resume:
80 On Error GoTo 0
90 Exit Sub

cmdUnAssignRes_Click_Error:

100 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.cmdUnAssignRes_Click." & Erl
110 Resume cmdUnAssignRes_Click_Resume

End Sub
Public Sub SerialEditByFactory()
  
  If gUser.LEvel <> LEVEL_FACTORY Then
    txtSerial.Locked = True
  End If

End Sub



Public Property Let DeviceID(ByVal DeviceID As Long)
10      On Error GoTo DeviceID_Error

20      mDeviceID = DeviceID
30      If mDeviceID = 0 Then
40        txtModel.Visible = False
50        cboDeviceType.Visible = True
55        txtSerial.Locked = False

60        SerialEditByFactory


70        cmdAssignRoom.Enabled = False
80        cmdAssignRoom.BackColor = Me.BackColor
90        cmdAsignResident.Enabled = False
100       cmdAsignResident.BackColor = Me.BackColor
110       cmdUnassign.Enabled = False
120       cmdUnAssignRes.Enabled = False
130     Else
140       txtModel.Visible = False
150       cboDeviceType.Visible = True

160       txtSerial.Locked = False

170       SerialEditByFactory

180       cmdAssignRoom.Enabled = True
190       cmdAssignRoom.BackColor = &H80FF80
200       cmdAsignResident.Enabled = True
210       cmdAsignResident.BackColor = &HFFFF80
220       cmdUnassign.Enabled = True
230       cmdUnAssignRes.Enabled = True
240       cboDeviceType.Visible = True

250     End If
260     cmdAutoEnroll.Enabled = (mDeviceID = 0) And EnrollButtonVisible()

DeviceID_Resume:
270     On Error GoTo 0
280     Exit Property

DeviceID_Error:

290     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.DeviceID." & Erl
300     Resume DeviceID_Resume

End Property

Public Property Get DeviceID() As Long
  DeviceID = mDeviceID
End Property

Public Sub DisableAutoEnroll()
10 On Error GoTo DisableAutoEnroll_Error
20 If USE6080 Then
30  If MASTER Then

40    tmrEnroller.Enabled = False
50    AutoEnrollEnabled = False
60    cmdAutoEnroll.Enabled = True And EnrollButtonVisible()
      
70    If Not regws Is Nothing Then
80      regws.DisConnect
90      Set regws = Nothing
100   End If

110 End If
120 Else
130 If MASTER Then
140   AutoEnrollEnabled = False
150   cmdAutoEnroll.Enabled = True And EnrollButtonVisible()
160 Else
170   RemoteAutoEnrollEnabled = False        ' kill it (Polling) immediately
180   RemoteCancelAutoEnroll
190   cmdAutoEnroll.Enabled = True And EnrollButtonVisible()
200 End If

210 End If

DisableAutoEnroll_Resume:
220 On Error GoTo 0
230 Exit Sub

DisableAutoEnroll_Error:

240 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.DisableAutoEnroll." & Erl
250 Resume DisableAutoEnroll_Resume


End Sub

Sub Display()
  On Error GoTo Display_Error

  chkTamperInput.Visible = False

  If DeviceID = 0 Then
    SetNewDeviceType
    If 0 = StrComp(cboDeviceType.text, COM_DEV_NAME, vbTextCompare) Then
      cmdConfigureSerial.Visible = True
      cmdAutoEnroll.Visible = False And EnrollButtonVisible()
      chkClearByReset.Value = 0

      chkClearByReset.Visible = False
      chkExtern.Visible = True
      chkExtern_A.Visible = True
    ElseIf 0 = StrComp(cboDeviceType.text, "ES1723", vbTextCompare) Then
      cmdConfigureSerial.Visible = True
      cmdAutoEnroll.Visible = True And EnrollButtonVisible()
      chkClearByReset.Value = 0

      chkClearByReset.Visible = False
      chkExtern.Visible = False
      chkExtern_A.Visible = False

    Else
      cmdAutoEnroll.Visible = True And EnrollButtonVisible()
      chkClearByReset.Visible = True
      cmdConfigureSerial.Visible = False
      cmdAutoEnroll.Enabled = (mDeviceID = 0) And EnrollButtonVisible()
      chkExtern.Visible = False
      chkExtern_A.Visible = False
    End If

    Exit Sub
  End If


  If 0 = StrComp(mDevice.Model, COM_DEV_NAME, vbTextCompare) Then
    cmdConfigureSerial.Visible = True
    cmdAutoEnroll.Visible = False And EnrollButtonVisible()
    chkClearByReset.Value = 0
    chkClearByReset.Visible = False
    chkExtern.Visible = True
    chkExtern_A.Visible = True

  ElseIf 0 = StrComp(cboDeviceType.text, "ES1723", vbTextCompare) Then
    cmdConfigureSerial.Visible = True
    cmdAutoEnroll.Enabled = (mDeviceID = 0) And EnrollButtonVisible()
    chkClearByReset.Value = 0

    chkClearByReset.Visible = False
    chkExtern.Visible = False
    chkExtern_A.Visible = False


  Else
    cmdAutoEnroll.Visible = True And EnrollButtonVisible()
    chkClearByReset.Visible = True
    cmdConfigureSerial.Visible = False
    cmdAutoEnroll.Enabled = (mDeviceID = 0) And EnrollButtonVisible()
    chkExtern.Visible = False
    chkExtern_A.Visible = False

  End If


  If bycode Then Exit Sub

  bycode = True
  txtSerial.text = mDevice.Serial
  
  txtSerial.Locked = False
  
  SerialEditByFactory
  
  If DeviceID <> 0 Then
    txtSerial.Locked = True
  End If
  
  txtModel.text = mDevice.Model
  txtClearingDevice.text = mDevice.Configurationstring
  
  
  
  cboDeviceType.ListIndex = CboFindExact(cboDeviceType, mDeviceType.Model)

  If cboDeviceType.ListIndex > -1 Then
    'mDeviceType = GetDeviceTypeByMIDPTI(cboDeviceType.ItemData(cboDeviceType.ListIndex))
    mDeviceType = GetDeviceTypeByModel(cboDeviceType.text)
    lblTypeDesc.Caption = mDeviceType.desc

  Else
    'mDeviceType = GetDeviceTypeByMIDPTI(0)
    mDeviceType = GetDeviceTypeByModel(cboDeviceType.text)
    lblTypeDesc.Caption = mDeviceType.desc
  End If

  chkIgnoreTamper.Value = IIf(mDeviceType.IgnoreTamper = 1, 1, 0)
  optInput3.Visible = CBool(mDevice.UseTamperAsInput) And (mDeviceType.NumInputs = 2)

  Select Case mDeviceType.NumInputs

    Case 3
    
      chkClearByReset.Visible = True
      optInput1.Visible = True
      optInput2.Visible = True
      optAssur1.Visible = True
      optAssur2.Visible = True
      optInput3.Visible = True
      chkTamperInput.Visible = False
    Case 2
      chkClearByReset.Visible = True
      optInput1.Visible = True
      optInput2.Visible = True
      optAssur1.Visible = True
      optAssur2.Visible = True
      chkTamperInput.Visible = True
      If mDeviceType.NumInputs = 2 And CBool(mDevice.UseTamperAsInput) Then ' Device.UseTamperAsInput
        optInput3.Visible = True
        chkTamperInput.Value = vbChecked
      Else
        optInput3.Visible = False
        chkTamperInput.Value = vbUnchecked
      End If
    Case 1
      chkClearByReset.Visible = True
      optInput1.Visible = False
      optInput2.Visible = False
      optInput1.Value = True
      InputNumber = 1  'mInput2 = False
      optAssur1.Visible = True
      optAssur2.Visible = False
      optInput3.Visible = False
      
    Case Else
      chkClearByReset.Visible = True
      optInput1.Visible = False
      optInput2.Visible = False
      optInput1.Value = True
      InputNumber = 1  'mInput2 = False
      optAssur1.Visible = False
      optAssur2.Visible = False
      optInput3.Visible = False
  End Select
  ' global for device

  If cboDeviceType.text = COM_DEV_NAME Then
    chkClearByReset.Visible = False
    chkClearByReset.Value = False
    cmdConfigureSerial.Visible = True
    cmdAutoEnroll.Visible = False And EnrollButtonVisible()

  ElseIf 0 = StrComp(cboDeviceType.text, "ES1723", vbTextCompare) Then
    chkClearByReset.Visible = False
    cmdAutoEnroll.Visible = True And EnrollButtonVisible()
    cmdConfigureSerial.Visible = True
    chkClearByReset.Value = 0

  Else
    chkClearByReset.Visible = True
    cmdAutoEnroll.Visible = True And EnrollButtonVisible()
    cmdConfigureSerial.Visible = False

  End If



  chkClearByReset.Value = mDevice.ClearByReset

'  If mInput2 Then
'    If optInput2.Value <> True Then
'      optInput2.Value = True
'    End If
'  ElseIf optInput1.Value <> True Then
'    optInput1.Value = True
'  End If
  
  
    Select Case InputNumber
      Case 3
        If optInput3.Value <> True Then
          optInput3.Value = True
        End If
      
      Case 2
        If optInput2.Value <> True Then
          optInput2.Value = True
        End If
      
      Case Else
      
        If optInput1.Value <> True Then
          optInput1.Value = True
        End If

      
    End Select

  If DeviceID <> 0 Then
    txtAnnounce.text = mDevice.Announce
    txtAnnounce2.text = mDevice.Announce_A
    txtAnnounce3.text = mDevice.Announce_B
    
    
    txtCustom.text = IIf(Len(mDevice.Custom) > 0, mDevice.Custom, lblTypeDesc.Caption)
  End If

  '

  chkIgnoreTamper.Value = IIf(mDevice.IgnoreTamper = 1, 1, 0)

  ' input 1
  cboGroup1.ListIndex = CboGetIndexByItemData(cboGroup1, mDevice.OG1)
  cboGroup2.ListIndex = CboGetIndexByItemData(cboGroup2, mDevice.OG2)
  cboGroup3.ListIndex = CboGetIndexByItemData(cboGroup3, mDevice.OG3)
  cboGroup4.ListIndex = CboGetIndexByItemData(cboGroup4, mDevice.OG4)
  cboGroup5.ListIndex = CboGetIndexByItemData(cboGroup5, mDevice.OG5)
  cboGroup6.ListIndex = CboGetIndexByItemData(cboGroup6, mDevice.OG6)


  cboGroupN1.ListIndex = CboGetIndexByItemData(cboGroupN1, mDevice.NG1)
  cboGroupN2.ListIndex = CboGetIndexByItemData(cboGroupN2, mDevice.NG2)
  cboGroupN3.ListIndex = CboGetIndexByItemData(cboGroupN3, mDevice.NG3)
  cboGroupN4.ListIndex = CboGetIndexByItemData(cboGroupN4, mDevice.NG4)
  cboGroupN5.ListIndex = CboGetIndexByItemData(cboGroupN5, mDevice.NG5)
  cboGroupN6.ListIndex = CboGetIndexByItemData(cboGroupN6, mDevice.NG6)


  cboGroupG1.ListIndex = CboGetIndexByItemData(cboGroupG1, mDevice.GG1)
  cboGroupG2.ListIndex = CboGetIndexByItemData(cboGroupG2, mDevice.GG2)
  cboGroupG3.ListIndex = CboGetIndexByItemData(cboGroupG3, mDevice.GG3)
  cboGroupG4.ListIndex = CboGetIndexByItemData(cboGroupG4, mDevice.GG4)
  cboGroupG5.ListIndex = CboGetIndexByItemData(cboGroupG5, mDevice.GG5)
  cboGroupG6.ListIndex = CboGetIndexByItemData(cboGroupG6, mDevice.GG6)

  txtOG1D.text = mDevice.OG1D
  txtOG2D.text = mDevice.OG2D
  txtOG3D.text = mDevice.OG3D
  txtOG4D.text = mDevice.OG4D
  txtOG5D.text = mDevice.OG5D
  txtOG6D.text = mDevice.OG6D

  txtNG1D.text = mDevice.NG1D
  txtNG2D.text = mDevice.NG2D
  txtNG3D.text = mDevice.NG3D
  txtNG4D.text = mDevice.NG4D
  txtNG5D.text = mDevice.NG5D
  txtNG6D.text = mDevice.NG6D


  txtGG1D.text = mDevice.GG1D
  txtGG2d.text = mDevice.GG2D
  txtGG3d.text = mDevice.GG3D
  txtGG4d.text = mDevice.GG4D
  txtGG5d.text = mDevice.GG5D
  txtGG6d.text = mDevice.GG6D


  chkRepeatUntil.Value = IIf(mDevice.RepeatUntil = 1, 1, 0)
  txtRepeats.text = mDevice.Repeats
  txtPause.text = mDevice.Pause

  chkAlarmAlert.Value = IIf(mDevice.AlarmMask = 1, 1, 0)
  chkExtern.Value = IIf(mDevice.AlarmMask = 2, 1, 0)
  chkSendCancel.Value = IIf(mDevice.SendCancel = 1, 1, 0)

  chkVacationSuper.Value = IIf(mDevice.AssurSecure = 1, 1, 0)

  txtDisableStart.text = mDevice.DisableStart
  txtDisableEnd.text = mDevice.DisableEnd

  'input 2

  txtOG1_AD.text = mDevice.OG1_AD
  txtOG2_AD.text = mDevice.OG2_AD
  txtOG3_AD.text = mDevice.OG3_AD
  txtOG4_AD.text = mDevice.OG4_AD
  txtOG5_AD.text = mDevice.OG5_AD
  txtOG6_AD.text = mDevice.OG6_AD

  txtNG1_AD.text = mDevice.NG1_AD
  txtNG2_AD.text = mDevice.NG2_AD
  txtNG3_AD.text = mDevice.NG3_AD
  txtNG4_AD.text = mDevice.NG4_AD
  txtNG5_AD.text = mDevice.NG5_AD
  txtNG6_AD.text = mDevice.NG6_AD

  txtGG1_Ad.text = mDevice.GG1_AD
  txtGG2_Ad.text = mDevice.GG2_AD
  txtGG3_Ad.text = mDevice.GG3_AD
  txtGG4_Ad.text = mDevice.GG4_AD
  txtGG5_Ad.text = mDevice.GG5_AD
  txtGG6_Ad.text = mDevice.GG6_AD

  cboGroup1_A.ListIndex = CboGetIndexByItemData(cboGroup1_A, mDevice.OG1_A)    '
  cboGroup2_A.ListIndex = CboGetIndexByItemData(cboGroup2_A, mDevice.OG2_A)
  cboGroup3_A.ListIndex = CboGetIndexByItemData(cboGroup3_A, mDevice.OG3_A)    '
  cboGroup4_A.ListIndex = CboGetIndexByItemData(cboGroup4_A, mDevice.OG4_A)
  cboGroup5_A.ListIndex = CboGetIndexByItemData(cboGroup5_A, mDevice.OG5_A)    '
  cboGroup6_A.ListIndex = CboGetIndexByItemData(cboGroup6_A, mDevice.OG6_A)



  cboGroupN1_A.ListIndex = CboGetIndexByItemData(cboGroupN1_A, mDevice.NG1_A)
  cboGroupN2_A.ListIndex = CboGetIndexByItemData(cboGroupN2_A, mDevice.NG2_A)
  cboGroupN3_A.ListIndex = CboGetIndexByItemData(cboGroupN3_A, mDevice.NG3_A)
  cboGroupN4_A.ListIndex = CboGetIndexByItemData(cboGroupN4_A, mDevice.NG4_A)
  cboGroupN5_A.ListIndex = CboGetIndexByItemData(cboGroupN5_A, mDevice.NG5_A)
  cboGroupN6_A.ListIndex = CboGetIndexByItemData(cboGroupN6_A, mDevice.NG6_A)

  cboGroupG1_A.ListIndex = CboGetIndexByItemData(cboGroupG1_A, mDevice.GG1_A)
  cboGroupG2_A.ListIndex = CboGetIndexByItemData(cboGroupG2_A, mDevice.GG2_A)
  cboGroupG3_A.ListIndex = CboGetIndexByItemData(cboGroupG3_A, mDevice.GG3_A)
  cboGroupG4_A.ListIndex = CboGetIndexByItemData(cboGroupG4_A, mDevice.GG4_A)
  cboGroupG5_A.ListIndex = CboGetIndexByItemData(cboGroupG5_A, mDevice.GG5_A)
  cboGroupG6_A.ListIndex = CboGetIndexByItemData(cboGroupG6_A, mDevice.GG6_A)


  chkRepeatUntil_A.Value = IIf(mDevice.RepeatUntil_A = 1, 1, 0)    ' IIf(rs("repeatuntil") = 1, 1, 0)
  txtRepeats_A.text = mDevice.Repeats_A
  txtPause_A.text = mDevice.Pause_A

  chkAlarmAlert_A.Value = IIf(mDevice.AlarmMask_A = 1, 1, 0)
  chkExtern_A.Value = IIf(mDevice.AlarmMask_A = 2, 1, 0)
  chkSendCancel_A.Value = IIf(mDevice.SendCancel_A = 1, 1, 0)

  chkVacationSuper_A.Value = IIf(mDevice.AssurSecure_A = 1, 1, 0)

  txtDisableStart_A.text = mDevice.DisableStart_A
  txtDisableEnd_A.text = mDevice.DisableEnd_A
  
  txtDisableStart_B.text = mDevice.DisableStart_B
  txtDisableEnd_B.text = mDevice.DisableEnd_B



' INPUT 3 (Also tamper as input

  txtOG1_BD.text = mDevice.OG1_BD
  txtOG2_BD.text = mDevice.OG2_BD
  txtOG3_BD.text = mDevice.OG3_BD
  txtOG4_BD.text = mDevice.OG4_BD
  txtOG5_BD.text = mDevice.OG5_BD
  txtOG6_BD.text = mDevice.OG6_BD

  txtNG1_BD.text = mDevice.NG1_BD
  txtNG2_BD.text = mDevice.NG2_BD
  txtNG3_BD.text = mDevice.NG3_BD
  txtNG4_BD.text = mDevice.NG4_BD
  txtNG5_BD.text = mDevice.NG5_BD
  txtNG6_BD.text = mDevice.NG6_BD

  txtGG1_BD.text = mDevice.GG1_BD
  txtGG2_BD.text = mDevice.GG2_BD
  txtGG3_BD.text = mDevice.GG3_BD
  txtGG4_BD.text = mDevice.GG4_BD
  txtGG5_BD.text = mDevice.GG5_BD
  txtGG6_BD.text = mDevice.GG6_BD

  cboGroup1_B.ListIndex = CboGetIndexByItemData(cboGroup1_B, mDevice.OG1_B)    '
  cboGroup2_B.ListIndex = CboGetIndexByItemData(cboGroup2_B, mDevice.OG2_B)
  cboGroup3_B.ListIndex = CboGetIndexByItemData(cboGroup3_B, mDevice.OG3_B)    '
  cboGroup4_B.ListIndex = CboGetIndexByItemData(cboGroup4_B, mDevice.OG4_B)
  cboGroup5_B.ListIndex = CboGetIndexByItemData(cboGroup5_B, mDevice.OG5_B)    '
  cboGroup6_B.ListIndex = CboGetIndexByItemData(cboGroup6_B, mDevice.OG6_B)



  cboGroupN1_B.ListIndex = CboGetIndexByItemData(cboGroupN1_B, mDevice.NG1_B)
  cboGroupN2_B.ListIndex = CboGetIndexByItemData(cboGroupN2_B, mDevice.NG2_B)
  cboGroupN3_B.ListIndex = CboGetIndexByItemData(cboGroupN3_B, mDevice.NG3_B)
  cboGroupN4_B.ListIndex = CboGetIndexByItemData(cboGroupN4_B, mDevice.NG4_B)
  cboGroupN5_B.ListIndex = CboGetIndexByItemData(cboGroupN5_B, mDevice.NG5_B)
  cboGroupN6_B.ListIndex = CboGetIndexByItemData(cboGroupN6_B, mDevice.NG6_B)

  cboGroupG1_B.ListIndex = CboGetIndexByItemData(cboGroupG1_B, mDevice.GG1_B)
  cboGroupG2_B.ListIndex = CboGetIndexByItemData(cboGroupG2_B, mDevice.GG2_B)
  cboGroupG3_B.ListIndex = CboGetIndexByItemData(cboGroupG3_B, mDevice.GG3_B)
  cboGroupG4_B.ListIndex = CboGetIndexByItemData(cboGroupG4_B, mDevice.GG4_B)
  cboGroupG5_B.ListIndex = CboGetIndexByItemData(cboGroupG5_B, mDevice.GG5_B)
  cboGroupG6_B.ListIndex = CboGetIndexByItemData(cboGroupG6_B, mDevice.GG6_B)


  chkRepeatUntil_B.Value = IIf(mDevice.RepeatUntil_B = 1, 1, 0)    ' IIf(rs("repeatuntil") = 1, 1, 0)
  txtRepeats_B.text = mDevice.Repeats_B
  txtPause_B.text = mDevice.Pause_B

' check this
  chkAlarmAlert_B.Value = IIf(mDevice.AlarmMask_B = 1, 1, 0)
' and this  chkExtern_b.Value = IIf(mDevice.AlarmMask_B = 2, 1, 0)
  chkSendCancel_B.Value = IIf(mDevice.SendCancel_B = 1, 1, 0)

' and this  chkVacationSuper_b.Value = IIf(mDevice.AssurSecure_B = 1, 1, 0)

'  txtDisableStart_b.text = mDevice.DisableStart_B
'  txtDisableEnd_b.text = mDevice.DisableEnd_B




  ' all inputs


  chkAssurance.Value = IIf(mDevice.UseAssur = 1, 1, 0)
  chkAssurance2.Value = IIf(mDevice.UseAssur2 = 1, 1, 0)

  chkIgnore.Value = mDevice.Ignored And 1

  ' show room/resident assignment
  If mDevice.RoomID <> 0 Then
    txtAssigned.text = mDevice.Room
  Else
    txtAssigned.text = ""
  End If
  If mDevice.ResidentID <> 0 Then
    txtAssignRes.text = mDevice.LastFirst
  Else
    txtAssignRes.text = ""
  End If

  Select Case mDevice.AssurInput
    Case 0
      optAssur1.Value = False
      optAssur2.Value = False
    Case 1
      optAssur1.Value = True
      optAssur2.Value = False
    Case 2
      optAssur1.Value = False
      optAssur2.Value = True
  End Select

  If Configuration.AssurStart = Configuration.AssurEnd Then
    lblAssurTime1.Caption = "(Disabled)"
  Else
    lblAssurTime1.Caption = "(" & ConvertHourToAMPM(Configuration.AssurStart) & " - " & ConvertHourToAMPM(Configuration.AssurEnd) & ")"
  End If
  If Configuration.AssurStart2 = Configuration.AssurEnd2 Then
    lblAssurTime2.Caption = "(Disabled)"
  Else
    lblAssurTime2.Caption = "(" & ConvertHourToAMPM(Configuration.AssurStart2) & " - " & ConvertHourToAMPM(Configuration.AssurEnd2) & ")"
  End If

  If ManagingPartitions Then
    FillAvailPartitionList   ' MUST CALL THIS FIRST
    FillActivePartitionList
  End If

  Dim j             As Long

  For j = cboDeviceMode.listcount - 1 To 0 Step -1
    If cboDeviceMode.ItemData(j) = mDevice.IDL Then
      Exit For
    End If
  Next
  j = Max(0, j)
  cboDeviceMode.ListIndex = j

  lblIDM.Caption = mDevice.ZoneID

  ManagingPartitions = False



  bycode = False

Display_Resume:
  On Error GoTo 0
  Exit Sub

Display_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.Display." & Erl
  Resume Display_Resume

End Sub

Sub EnableAutoEnroll()
        'Public Sub AutoEnroll(p As cESPacket)
        '  AutoEnrollEnabled = False
        '  setdevice p.serial, p.MIDPTI
        '  cmdAutoEnroll.Enabled = (mDeviceID = 0)
        'End Sub
        'Sub setdevice(ByVal serial As String, ByVal MIDPTI As Long)
        '  txtSerial.text = serial
        '  Dim i As Long
        '  i = CboGetIndexByItemData(cboDeviceType, MIDPTI)
        '  If i < 0 Then i = 0
        '  cboDeviceType.ListIndex = i
        '
        'End Sub

10      On Error GoTo EnableAutoEnroll_Error

20      If USE6080 Then
30        If MASTER Then
40          Set mDevice = New cESDevice
50          cmdAutoEnroll.Enabled = False And EnrollButtonVisible()
60          If regws Is Nothing Then
70            Set regws = New WebSocketSocket
80          End If
90          regws.Init "5002EgAtIrEh"
100         regws.UserNamePassword USER1, PW1
            'ws.SetURL "ws://echo.websocket.org"

110         regws.SetURL GetWS & "://" & IP1 & "/PSIA/Metadata/stream?Registration=true"
120         regws.Connect
130         Sleep 100
140         Do While regws.HasMessages
150           regws.GetNextMessage
160         Loop


170         tmrEnroller.Enabled = True
180       End If
190     Else
200       If MASTER Then
210         cmdAutoEnroll.Enabled = False And EnrollButtonVisible()
220         AutoEnrollEnabled = True
230       Else
240         cmdAutoEnroll.Enabled = False And EnrollButtonVisible()
250         RemoteStartAutoEnroll
260         RemoteAutoEnrollEnabled = True
270       End If
280     End If

EnableAutoEnroll_Resume:
290     On Error GoTo 0
300     Exit Sub

EnableAutoEnroll_Error:

310     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.EnableAutoEnroll." & Erl
320     Resume EnableAutoEnroll_Resume


End Sub

Public Sub Fill()
        Dim rs            As Recordset
        Dim SQL           As String
10      Set mDevice = New cESDevice

        SerialEditByFactory



20      SQL = "select devices.* from devices where deviceid = " & DeviceID
30      Set rs = ConnExecute(SQL)
40      If Not rs.EOF Then
50        mDevice.Parse rs
60        mDevice.FetchResident
70        mDevice.FetchRoom
80        mDevice.GetZoneInfo
90      Else
100       DeviceID = 0
110       Set mDevice = New cESDevice
120     End If
130     rs.Close
140     Set rs = Nothing
150     mDeviceType = GetDeviceTypeByModel(mDevice.Model)
160     Display



End Sub

Sub FillAvailPartitionList()
        Dim part          As cPartition
        Dim li            As ListItem
        Dim Found         As Boolean
        Dim XML           As String
        Dim HTTPRequest   As cHTTPRequest

10      lvAvailPartitions.ListItems.Clear

20      If USE6080 Then

30        Set HTTPRequest = New cHTTPRequest
40        XML = HTTPRequest.GetPartitionList(GetHTTP & "://" & IP1, USER1, PW1)

50        If Len(XML) Then
60          Set Partitions = ParsePartionList(XML)
70        Else
80          Set Partitions = New Collection
90        End If

100       For Each part In Partitions
110         If Len(Trim$(Me.txtSearchBox.text)) > 0 Then
120           Found = InStr(1, part.Description, Me.txtSearchBox.text, vbTextCompare) <> 0
130         Else
140           Found = True
150         End If
160         If Found Then
170           Set li = lvAvailPartitions.ListItems.Add(, , part.PartitionID)
180           li.SubItems(1) = part.Description
190           li.SubItems(2) = IIf(part.IsLocation, "X", "")
200         End If

210       Next

220       For Each li In lvAvailPartitions.ListItems
230         li.Selected = 0
240       Next

250     End If

End Sub

Private Sub Form_Initialize()
  Set mDevice = New cESDevice
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ClearAlerts
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select
End Sub

Private Sub Form_Load()

  ResetActivityTime


  SerialEditByFactory

  On Error Resume Next


  If (Not MASTER) Then
    If (USE6080) Then
      cmdNew.Visible = False

    End If
  End If



  If mDevice Is Nothing Then
    Set mDevice = New cESDevice
  End If

  cboDeviceMode.Clear

  AddToCombo cboDeviceMode, "NA", 0
  AddToCombo cboDeviceMode, "Mob", 1
  AddToCombo cboDeviceMode, "Fix", 2
  AddToCombo cboDeviceMode, "SP", 3

  cboDeviceMode.ListIndex = 0

  If Enroller Is Nothing Then
    Set Enroller = New c6080
  End If

  If USE6080 = 0 Then
    TabStrip.Tabs.Remove ("partitions")
    cboDeviceMode.Visible = False
    lblSerial(1).Visible = False

  Else
    Me.lblSerial(1).Visible = False
    cboDeviceMode.Visible = True
    lblSerial(1).Visible = True


  End If

  On Error GoTo Form_Load_Error

  ClearAlerts
  cmdAsignResident.BackColor = &HFFFF80
  cmdAssignRoom.BackColor = &H80FF80


  cmdAsignResident.Enabled = False
  cmdAsignResident.BackColor = Me.BackColor
  cmdAssignRoom.Enabled = False
  cmdAssignRoom.BackColor = Me.BackColor
  cmdUnassign.Enabled = False
  cmdUnAssignRes.Enabled = False
  optInput1.Value = True

  SetControls
  LoadCombos
  Connect
  ResetForm

  If Configuration.AssurStart = Configuration.AssurEnd Then
    lblAssurTime1.Caption = "(Disabled)"
    chkAssurance.Enabled = False
  Else
    lblAssurTime1.Caption = "(" & ConvertHourToAMPM(Configuration.AssurStart) & " - " & ConvertHourToAMPM(Configuration.AssurEnd) & ")"
    chkAssurance.Enabled = True
  End If
  If Configuration.AssurStart2 = Configuration.AssurEnd2 Then
    lblAssurTime2.Caption = "(Disabled)"
    chkAssurance2.Enabled = False
  Else
    lblAssurTime2.Caption = "(" & ConvertHourToAMPM(Configuration.AssurStart2) & " - " & ConvertHourToAMPM(Configuration.AssurEnd2) & ")"
    chkAssurance2.Enabled = True
  End If



Form_Load_Resume:

  'cmdAssignRoom.Refresh
  'Me.cmdAsignResident.Refresh

  On Error GoTo 0
  Exit Sub

Form_Load_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.Form_Load." & Erl
  Resume Form_Load_Resume





End Sub

Private Sub Form_Paint()
    
  SerialEditByFactory

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ClearAlerts
  DisableAutoEnroll
  UnHost
  Set mDevice = Nothing
End Sub

Function GetDeviceIDFromSerial(ByVal Serial As String) As Long
  Dim rs            As Recordset
  Dim t2            As Double

10 On Error GoTo GetDeviceIDFromSerial_Error

20 RefreshJet
  t2 = Timer + 1
  Do Until Timer > t2
    DoEvents
  Loop


30 Set rs = ConnExecute("Select deviceid from devices where serial =" & q(Serial))
40 If Not rs.EOF Then
50  GetDeviceIDFromSerial = rs(0)
60 End If
70 rs.Close
80 Set rs = Nothing



GetDeviceIDFromSerial_Resume:
90 On Error GoTo 0
100 Exit Function

GetDeviceIDFromSerial_Error:

110 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.GetDeviceIDFromSerial." & Erl
120 Resume GetDeviceIDFromSerial_Resume


End Function

Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub

Public Property Get Input2() As Boolean
  
  'Input2 = mInput2
  Input2 = InputNumber = 2
End Property

Public Property Let Input2(ByVal Value As Boolean)
  On Error Resume Next
  txtAnnounce2.Visible = Value
  txtAnnounce.Visible = Not Value
  txtAnnounce3.Visible = Not Value
  
  
  'mInput2 = Value
End Property

Sub LoadCombos()

        Dim rs            As Recordset
        'Set rs = connexecute("SELECT * FROM Devicetypes ORDER BY model")
        Dim j             As Integer


10      On Error GoTo LoadCombos_Error

20      For j = 0 To MAX_ESDEVICETYPES
30        If ESDeviceType(j).Model = "ES1723" Then
40          ESDeviceType(j).Model = "ES1723"
50        End If
60        AddToCombo cboDeviceType, ESDeviceType(j).Model, ESDeviceType(j).CLS * 256& + ESDeviceType(j).PTI
          'rs.MoveNext
70      Next
        'rs.Close
80      If cboDeviceType.listcount > 0 Then
90        cboDeviceType.ListIndex = 0
100     End If

110     Set rs = ConnExecute("SELECT * FROM pagergroups ORDER BY Description")

120     AddToCombo cboGroup1, "< none > ", 0
130     AddToCombo cboGroup2, "< none > ", 0
140     AddToCombo cboGroup3, "< none > ", 0
150     AddToCombo cboGroup4, "< none > ", 0
160     AddToCombo cboGroup5, "< none > ", 0
170     AddToCombo cboGroup6, "< none > ", 0

180     AddToCombo cboGroupN1, "< none > ", 0
190     AddToCombo cboGroupN2, "< none > ", 0
200     AddToCombo cboGroupN3, "< none > ", 0
210     AddToCombo cboGroupN4, "< none > ", 0
220     AddToCombo cboGroupN5, "< none > ", 0
230     AddToCombo cboGroupN6, "< none > ", 0

240     AddToCombo cboGroup1_A, "< none > ", 0
250     AddToCombo cboGroup2_A, "< none > ", 0
260     AddToCombo cboGroup3_A, "< none > ", 0
270     AddToCombo cboGroup4_A, "< none > ", 0
280     AddToCombo cboGroup5_A, "< none > ", 0
290     AddToCombo cboGroup6_A, "< none > ", 0

300     AddToCombo cboGroupN1_A, "< none > ", 0
310     AddToCombo cboGroupN2_A, "< none > ", 0
320     AddToCombo cboGroupN3_A, "< none > ", 0
330     AddToCombo cboGroupN4_A, "< none > ", 0
340     AddToCombo cboGroupN5_A, "< none > ", 0
350     AddToCombo cboGroupN6_A, "< none > ", 0

360     AddToCombo cboGroup1_B, "< none > ", 0
370     AddToCombo cboGroup2_B, "< none > ", 0
380     AddToCombo cboGroup3_B, "< none > ", 0
390     AddToCombo cboGroup4_B, "< none > ", 0
400     AddToCombo cboGroup5_B, "< none > ", 0
410     AddToCombo cboGroup6_B, "< none > ", 0

420     AddToCombo cboGroupN1_B, "< none > ", 0
430     AddToCombo cboGroupN2_B, "< none > ", 0
440     AddToCombo cboGroupN3_B, "< none > ", 0
450     AddToCombo cboGroupN4_B, "< none > ", 0
460     AddToCombo cboGroupN5_B, "< none > ", 0
470     AddToCombo cboGroupN6_B, "< none > ", 0

480     AddToCombo cboGroupG1, "< none > ", 0
490     AddToCombo cboGroupG2, "< none > ", 0
500     AddToCombo cboGroupG3, "< none > ", 0
510     AddToCombo cboGroupG4, "< none > ", 0
520     AddToCombo cboGroupG5, "< none > ", 0
530     AddToCombo cboGroupG6, "< none > ", 0

540     AddToCombo cboGroupG1_A, "< none > ", 0
550     AddToCombo cboGroupG2_A, "< none > ", 0
560     AddToCombo cboGroupG3_A, "< none > ", 0
570     AddToCombo cboGroupG4_A, "< none > ", 0
580     AddToCombo cboGroupG5_A, "< none > ", 0
590     AddToCombo cboGroupG6_A, "< none > ", 0

600     AddToCombo cboGroupG1_B, "< none > ", 0
610     AddToCombo cboGroupG2_B, "< none > ", 0
620     AddToCombo cboGroupG3_B, "< none > ", 0
630     AddToCombo cboGroupG4_B, "< none > ", 0
640     AddToCombo cboGroupG5_B, "< none > ", 0
650     AddToCombo cboGroupG6_B, "< none > ", 0




660     Do Until rs.EOF
670       AddToCombo cboGroup1, rs("description") & "", rs("groupID")
680       AddToCombo cboGroup2, rs("description") & "", rs("groupID")
690       AddToCombo cboGroup3, rs("description") & "", rs("groupID")
700       AddToCombo cboGroup4, rs("description") & "", rs("groupID")
710       AddToCombo cboGroup5, rs("description") & "", rs("groupID")
720       AddToCombo cboGroup6, rs("description") & "", rs("groupID")



730       AddToCombo cboGroupN1, rs("description") & "", rs("groupID")
740       AddToCombo cboGroupN2, rs("description") & "", rs("groupID")
750       AddToCombo cboGroupN3, rs("description") & "", rs("groupID")
760       AddToCombo cboGroupN4, rs("description") & "", rs("groupID")
770       AddToCombo cboGroupN5, rs("description") & "", rs("groupID")
780       AddToCombo cboGroupN6, rs("description") & "", rs("groupID")


790       AddToCombo cboGroupG1, rs("description") & "", rs("groupID")
800       AddToCombo cboGroupG2, rs("description") & "", rs("groupID")
810       AddToCombo cboGroupG3, rs("description") & "", rs("groupID")
820       AddToCombo cboGroupG4, rs("description") & "", rs("groupID")
830       AddToCombo cboGroupG5, rs("description") & "", rs("groupID")
840       AddToCombo cboGroupG6, rs("description") & "", rs("groupID")

850       AddToCombo cboGroup1_A, rs("description") & "", rs("groupID")
860       AddToCombo cboGroup2_A, rs("description") & "", rs("groupID")
870       AddToCombo cboGroup3_A, rs("description") & "", rs("groupID")
880       AddToCombo cboGroup4_A, rs("description") & "", rs("groupID")
890       AddToCombo cboGroup5_A, rs("description") & "", rs("groupID")
900       AddToCombo cboGroup6_A, rs("description") & "", rs("groupID")

910       AddToCombo cboGroupN1_A, rs("description") & "", rs("groupID")
920       AddToCombo cboGroupN2_A, rs("description") & "", rs("groupID")
930       AddToCombo cboGroupN3_A, rs("description") & "", rs("groupID")
940       AddToCombo cboGroupN4_A, rs("description") & "", rs("groupID")
950       AddToCombo cboGroupN5_A, rs("description") & "", rs("groupID")
960       AddToCombo cboGroupN6_A, rs("description") & "", rs("groupID")


970       AddToCombo cboGroupG1_A, rs("description") & "", rs("groupID")
980       AddToCombo cboGroupG2_A, rs("description") & "", rs("groupID")
990       AddToCombo cboGroupG3_A, rs("description") & "", rs("groupID")
1000      AddToCombo cboGroupG4_A, rs("description") & "", rs("groupID")
1010      AddToCombo cboGroupG5_A, rs("description") & "", rs("groupID")
1020      AddToCombo cboGroupG6_A, rs("description") & "", rs("groupID")


1030      AddToCombo cboGroup1_B, rs("description") & "", rs("groupID")
1040      AddToCombo cboGroup2_B, rs("description") & "", rs("groupID")
1050      AddToCombo cboGroup3_B, rs("description") & "", rs("groupID")
1060      AddToCombo cboGroup4_B, rs("description") & "", rs("groupID")
1070      AddToCombo cboGroup5_B, rs("description") & "", rs("groupID")
1080      AddToCombo cboGroup6_B, rs("description") & "", rs("groupID")

1090      AddToCombo cboGroupN1_B, rs("description") & "", rs("groupID")
1100      AddToCombo cboGroupN2_B, rs("description") & "", rs("groupID")
1110      AddToCombo cboGroupN3_B, rs("description") & "", rs("groupID")
1120      AddToCombo cboGroupN4_B, rs("description") & "", rs("groupID")
1130      AddToCombo cboGroupN5_B, rs("description") & "", rs("groupID")
1140      AddToCombo cboGroupN6_B, rs("description") & "", rs("groupID")


1150      AddToCombo cboGroupG1_B, rs("description") & "", rs("groupID")
1160      AddToCombo cboGroupG2_B, rs("description") & "", rs("groupID")
1170      AddToCombo cboGroupG3_B, rs("description") & "", rs("groupID")
1180      AddToCombo cboGroupG4_B, rs("description") & "", rs("groupID")
1190      AddToCombo cboGroupG5_B, rs("description") & "", rs("groupID")
1200      AddToCombo cboGroupG6_B, rs("description") & "", rs("groupID")


1210      rs.MoveNext
1220    Loop
1230    rs.Close
1240    Set rs = Nothing

LoadCombos_Resume:
1250    On Error GoTo 0
1260    Exit Sub

LoadCombos_Error:

1270    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.LoadCombos." & Erl
1280    Resume LoadCombos_Resume

End Sub

Private Sub optInput1_Click()
  ' save settings to input2

  'Input2 = False
  InputNumber = 1
  ShowPanels
  If Not bycode Then
    '    Display
  End If

End Sub

Private Sub optInput2_Click()
  ' save settings to input2

  'Input2 = True
  InputNumber = 2
  ShowPanels
  If Not bycode Then
    '  Display
  End If

End Sub

Private Sub ResetForm()
10      On Error GoTo ResetForm_Error

         
        SerialEditByFactory



20      Input2 = False
30      txtSerial.text = ""
40      txtAssigned.text = ""
50      txtAssignRes.text = ""



55      txtAnnounce.text = ""
60      txtAnnounce2.text = ""
        txtAnnounce3.text = ""

70      txtCustom.text = ""

80      lblTypeDesc.Caption = ""

90      chkAssurance.Value = vbUnchecked
100     chkAssurance2.Value = vbUnchecked

110     cboDeviceType.ListIndex = 0

120     chkRepeatUntil.Value = vbUnchecked
130     txtRepeats.text = "0"
140     txtPause.text = "0"

150     txtDisableStart.text = "0"
160     txtDisableEnd.text = "0"

165     chkTamperInput.Value = vbUnchecked
        chkTamperInput.Visible = False

170     chkVacationSuper.Value = vbUnchecked
180     chkSendCancel.Value = vbChecked
190     chkAlarmAlert.Value = vbUnchecked

200     cboGroup1.ListIndex = 0
210     cboGroup2.ListIndex = 0
220     cboGroup3.ListIndex = 0
230     cboGroup4.ListIndex = 0
240     cboGroup5.ListIndex = 0
250     cboGroup6.ListIndex = 0



260     cboGroupN1.ListIndex = 0
270     cboGroupN2.ListIndex = 0
280     cboGroupN3.ListIndex = 0
290     cboGroupN4.ListIndex = 0
300     cboGroupN5.ListIndex = 0
310     cboGroupN6.ListIndex = 0



320     cboGroupG1.ListIndex = 0
330     cboGroupG2.ListIndex = 0
340     cboGroupG3.ListIndex = 0
350     cboGroupG4.ListIndex = 0
360     cboGroupG5.ListIndex = 0
370     cboGroupG6.ListIndex = 0


380     txtOG1D.text = 0
390     txtOG2D.text = 0
400     txtOG3D.text = 0
410     txtOG4D.text = 0
420     txtOG5D.text = 0
430     txtOG6D.text = 0


440     txtNG1D.text = 0
450     txtNG2D.text = 0
460     txtNG3D.text = 0
470     txtNG4D.text = 0
480     txtNG5D.text = 0
490     txtNG6D.text = 0

500     txtGG1D.text = 0
510     txtGG2d.text = 0
520     txtGG3d.text = 0
530     txtGG4d.text = 0
540     txtGG5d.text = 0
550     txtGG6d.text = 0




        ' _A settings

560     chkRepeatUntil_A.Value = vbUnchecked
570     txtRepeats_A.text = "0"
580     txtPause_A.text = "0"


590     chkVacationSuper_A.Value = vbUnchecked
600     chkSendCancel_A.Value = vbChecked
610     chkAlarmAlert_A.Value = 0

620     cboGroup1_A.ListIndex = 0
630     cboGroup2_A.ListIndex = 0
640     cboGroup3_A.ListIndex = 0
650     cboGroup4_A.ListIndex = 0
660     cboGroup5_A.ListIndex = 0
670     cboGroup6_A.ListIndex = 0

680     cboGroupN1_A.ListIndex = 0
690     cboGroupN2_A.ListIndex = 0
700     cboGroupN3_A.ListIndex = 0
710     cboGroupN4_A.ListIndex = 0
720     cboGroupN5_A.ListIndex = 0
730     cboGroupN6_A.ListIndex = 0



740     cboGroupG1_A.ListIndex = 0
750     cboGroupG2_A.ListIndex = 0
760     cboGroupG3_A.ListIndex = 0
770     cboGroupG4_A.ListIndex = 0
780     cboGroupG5_A.ListIndex = 0
790     cboGroupG6_A.ListIndex = 0


800     txtOG1_AD.text = 0
810     txtOG2_AD.text = 0
820     txtOG3_AD.text = 0
830     txtOG4_AD.text = 0
840     txtOG5_AD.text = 0
850     txtOG6_AD.text = 0


860     txtNG1_AD.text = 0
870     txtNG2_AD.text = 0
880     txtNG3_AD.text = 0
890     txtNG4_AD.text = 0
900     txtNG5_AD.text = 0
910     txtNG6_AD.text = 0

920     txtGG1_Ad.text = 0
930     txtGG2_Ad.text = 0
940     txtGG3_Ad.text = 0
950     txtGG4_Ad.text = 0
960     txtGG5_Ad.text = 0
970     txtGG6_Ad.text = 0



        ' _B settings

980     chkRepeatUntil_B.Value = vbUnchecked
990     chkSendCancel_B.Value = vbChecked
1000    txtRepeats_B.text = 0
1010    txtPause_B.text = 0



        'chkVacationSuper_b.Value = 0
1020    chkSendCancel_B.Value = 0
        'chkAlarmAlert_b.Value = 0

1030    cboGroup1_B.ListIndex = 0
1040    cboGroup2_B.ListIndex = 0
1050    cboGroup3_B.ListIndex = 0
1060    cboGroup4_B.ListIndex = 0
1070    cboGroup5_B.ListIndex = 0
1080    cboGroup6_B.ListIndex = 0

1090    cboGroupN1_B.ListIndex = 0
1100    cboGroupN2_B.ListIndex = 0
1110    cboGroupN3_B.ListIndex = 0
1120    cboGroupN4_B.ListIndex = 0
1130    cboGroupN5_B.ListIndex = 0
1140    cboGroupN6_B.ListIndex = 0



1150    cboGroupG1_B.ListIndex = 0
1160    cboGroupG2_B.ListIndex = 0
1170    cboGroupG3_B.ListIndex = 0
1180    cboGroupG4_B.ListIndex = 0
1190    cboGroupG5_B.ListIndex = 0
1200    cboGroupG6_B.ListIndex = 0



1210    txtOG1_BD.text = 0
1220    txtOG2_BD.text = 0
1230    txtOG3_BD.text = 0
1240    txtOG4_BD.text = 0
1250    txtOG5_BD.text = 0
1260    txtOG6_BD.text = 0


1270    txtNG1_BD.text = 0
1280    txtNG2_BD.text = 0
1290    txtNG3_BD.text = 0
1300    txtNG4_BD.text = 0
1310    txtNG5_BD.text = 0
1320    txtNG6_BD.text = 0

1330    txtGG1_BD.text = 0
1340    txtGG2_BD.text = 0
1350    txtGG3_BD.text = 0
1360    txtGG4_BD.text = 0
1370    txtGG5_BD.text = 0
1380    txtGG6_BD.text = 0



1390    chkIgnore.Value = 0

ResetForm_Resume:
1400    On Error GoTo 0
1410    Exit Sub

ResetForm_Error:

1420    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.ResetForm." & Erl
1430    Resume ResetForm_Resume


End Sub

Function Save() As Boolean

        Dim rs                 As Recordset
        Dim Count              As Long
        Dim Serial             As String
        Dim Model              As String
        Dim index              As Integer
        Dim SQL                As String

        Dim CLSPTI             As Long

        Dim adding             As Boolean
        Dim rslog              As Recordset

        Dim hextype            As String
        Dim newserial          As String
        Dim SerialValue        As String
        Dim ZoneInfo           As cZoneInfo
        Dim SCICode            As Long

        Dim HTTPRequest        As cHTTPRequest


10      On Error GoTo Save_Error

20      ClearAlerts
30      If mDevice.ZoneID <> 0 Then
          'Stop
40      End If


50      index = cboDeviceType.ListIndex
60      If index >= 0 Then


70        hextype = Hex(cboDeviceType.ItemData(index))
80        Select Case hextype
            Case "3DD"  ' dukane 3
              ' strip off Dn prefix
90            newserial = Right$("000000" & txtSerial.text, 6)
100           SerialValue = Val(newserial)
110           newserial = "D3" & Right$("000000" & Format$(SerialValue, "0"), 6)
120           txtSerial.text = "D3" & Right$("000000" & Format$(SerialValue, "0"), 6)
130         Case "5DD  ' dukane 5"
140           newserial = Right$("000000" & txtSerial.text, 6)
150           SerialValue = Val(newserial)
160           newserial = "D5" & Right$("000000" & Format$(SerialValue, "0"), 6)
170           txtSerial.text = "D5" & Right$("000000" & Format$(SerialValue, "0"), 6)
180       End Select
190     End If



200     If MASTER Then    '********************** master ****************

210       DisableAutoEnroll

220       txtSerial.text = Trim(txtSerial.text)
230       If Len(txtSerial.text) < 8 Then
240         Beep
250         lblAlert.Caption = "Serial Number Must Be 8 Characters"
260         Exit Function
270       End If

280       index = cboDeviceType.ListIndex

290       If index < 0 Then
300         lblAlert.Caption = "Device Type Not Selected"
310         Exit Function
320       End If

330       Model = UCase(cboDeviceType.text)
340       Serial = UCase(Trim(txtSerial.text))

350       mDeviceType = GetDeviceTypeByModel(Model)

360       mDevice.MID = mDeviceType.MID
370       mDevice.PTI = mDeviceType.PTI
380       mDevice.CLS = mDeviceType.CLS
390       mDevice.Description = Serial
400       mDevice.Checkin6080 = Max(MIN_CHECKIN, mDeviceType.Checkin * 1 * 60)

          mDevice.Configurationstring = Trim$(txtClearingDevice.text)

410       mDevice.Serial = Serial
          Dim modenum          As Long
420       modenum = Max(0, Me.cboDeviceMode.ListIndex)

430       '

440       Select Case modenum
            Case 0
450           mDevice.IsPortable = 0
460           mDevice.IsRef = 0
470           mDevice.IsSPDevice = 0
480         Case 1
490           mDevice.IsPortable = 1
500           mDevice.IsRef = 0
510           mDevice.IsSPDevice = 0

520         Case 2
530           mDevice.IsPortable = 0
540           mDevice.IsRef = 1
550           mDevice.IsSPDevice = 0

560         Case 3
570           mDevice.IsPortable = 0
580           mDevice.IsRef = 0
590           mDevice.IsSPDevice = 1

600         Case Else
610           mDevice.IsPortable = 0
620           mDevice.IsRef = 0
630           mDevice.IsSPDevice = 0

640       End Select

650       If USE6080 Then
660         If mDevice.Model = "" Then
670           mDevice.Model = mDeviceType.Model
680         End If
            
            ' only process ZoneInfo for ES and EN Devices
690         If (left$(mDevice.Model, 2) = "EN") Or (left$(mDevice.Model, 2) = "ES") Then

700           If mDevice.ZoneID = 0 Then  ' no zoneID, probably not registered with 6080
710             mDevice.ZoneID = RegisterDevice(mDevice)  ' will get current assignment if already registered with 6080
720             If mDevice.ZoneID Then
730               Set HTTPRequest = New cHTTPRequest
740               Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, mDevice.ZoneID)
750               Set HTTPRequest = Nothing
760             End If
770           Else  ' have ZoneID
                ' get device zoneinfo
780             Set HTTPRequest = New cHTTPRequest
790             Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, mDevice.ZoneID)
800             Set HTTPRequest = Nothing
810             If ZoneInfo Is Nothing Then  ' we have an ID but missing from 6080
820               mDevice.ZoneID = RegisterDevice(mDevice)  ' will get current assignment if already registered with 6080
830               If mDevice.ZoneID Then
840                 Set HTTPRequest = New cHTTPRequest
850                 Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, mDevice.ZoneID)

860                 Set HTTPRequest = Nothing
870               End If
880             End If

890           End If

900           If Not ZoneInfo Is Nothing Then

910             If ZoneInfo.IDL <> mDevice.IDL Then  ' IDL is whether it's a fixed, portable or ref device
920               Set HTTPRequest = New cHTTPRequest
930               Call HTTPRequest.setNewIDL(GetHTTP & "://" & IP1, USER1, PW1, mDevice.ZoneID, mDevice.IDL)
940               Set HTTPRequest = Nothing
950             End If
960           End If  'Zoneinfo = nothing

970         End If  ' Device.Model, 2) = "EN"
980       End If  ' If USE6080 Then

990       SQL = "SELECT COUNT(deviceid) FROM Devices WHERE serial = " & q(Serial)
1000      Set rs = ConnExecute(SQL)
1010      Count = rs(0)
1020      rs.Close
1030      Set rs = Nothing
1040      If Count > 0 And DeviceID = 0 Then  ' if device id = 0 then we need to fetch it
1050        Beep
1060        lblAlert.Caption = "Duplicate Serial Number ... Please Verify Serial Number"
1070        Exit Function
1080      Else
            ' singletons

1090        mDevice.Serial = Serial
1100        mDevice.Model = Model

1110        mDevice.Announce = Trim(txtAnnounce.text)
1120        mDevice.Announce_A = Trim(txtAnnounce2.text)
1130        mDevice.Announce_B = Trim(txtAnnounce3.text)

1140        mDeviceType = GetDeviceTypeByModel(mDevice.Model)

1150        Select Case mDeviceType.NumInputs

              Case 1
1160            mDevice.AssurInput = 1
1170          Case 2
1180            If optAssur1.Value Then
1190              mDevice.AssurInput = 1
1200            ElseIf optAssur2.Value Then
1210              mDevice.AssurInput = 2
1220            Else
1230              mDevice.AssurInput = 1
1240            End If
1250          Case 3

1260          Case Else
1270            mDevice.AssurInput = 0
1280        End Select

1290        mDevice.NumInputs = mDeviceType.NumInputs
1300        mDevice.IsPortable = mDeviceType.Portable
1310        mDevice.AutoClear = mDeviceType.AutoClear

1320        mDevice.ClearByReset = chkClearByReset.Value And 1
            ' ************ common
1330        mDevice.IgnoreTamper = chkIgnoreTamper.Value And 1

1340        mDevice.UseTamperAsInput = chkTamperInput.Value And 1


1350        mDevice.OG1 = GetComboItemData(cboGroup1)
1360        mDevice.OG2 = GetComboItemData(cboGroup2)
1370        mDevice.OG3 = GetComboItemData(cboGroup3)
1380        mDevice.OG4 = GetComboItemData(cboGroup4)
1390        mDevice.OG5 = GetComboItemData(cboGroup5)
1400        mDevice.OG6 = GetComboItemData(cboGroup6)

1410        mDevice.OG1D = Val(txtOG1D.text)
1420        mDevice.OG2D = Val(txtOG2D.text)
1430        mDevice.OG3D = Val(txtOG3D.text)
1440        mDevice.OG4D = Val(txtOG4D.text)
1450        mDevice.OG5D = Val(txtOG5D.text)
1460        mDevice.OG6D = Val(txtOG6D.text)


1470        mDevice.OG1_A = GetComboItemData(cboGroup1_A)
1480        mDevice.OG2_A = GetComboItemData(cboGroup2_A)
1490        mDevice.OG3_A = GetComboItemData(cboGroup3_A)
1500        mDevice.OG4_A = GetComboItemData(cboGroup4_A)
1510        mDevice.OG5_A = GetComboItemData(cboGroup5_A)
1520        mDevice.OG6_A = GetComboItemData(cboGroup6_A)

1530        mDevice.OG1_AD = Val(txtOG1_AD.text)
1540        mDevice.OG2_AD = Val(txtOG2_AD.text)
1550        mDevice.OG3_AD = Val(txtOG3_AD.text)
1560        mDevice.OG4_AD = Val(txtOG4_AD.text)
1570        mDevice.OG5_AD = Val(txtOG5_AD.text)
1580        mDevice.OG6_AD = Val(txtOG6_AD.text)


1590        mDevice.OG1_B = GetComboItemData(cboGroup1_B)
1600        mDevice.OG2_B = GetComboItemData(cboGroup2_B)
1610        mDevice.OG3_B = GetComboItemData(cboGroup3_B)
1620        mDevice.OG4_B = GetComboItemData(cboGroup4_B)
1630        mDevice.OG5_B = GetComboItemData(cboGroup5_B)
1640        mDevice.OG6_B = GetComboItemData(cboGroup6_B)

1650        mDevice.OG1_BD = Val(txtOG1_BD.text)
1660        mDevice.OG2_BD = Val(txtOG2_BD.text)
1670        mDevice.OG3_BD = Val(txtOG3_BD.text)
1680        mDevice.OG4_BD = Val(txtOG4_BD.text)
1690        mDevice.OG5_BD = Val(txtOG5_BD.text)
1700        mDevice.OG6_BD = Val(txtOG6_BD.text)



1710        mDevice.NG1 = GetComboItemData(cboGroupN1)
1720        mDevice.NG2 = GetComboItemData(cboGroupN2)
1730        mDevice.NG3 = GetComboItemData(cboGroupN3)
1740        mDevice.NG4 = GetComboItemData(cboGroupN4)
1750        mDevice.NG5 = GetComboItemData(cboGroupN5)
1760        mDevice.NG6 = GetComboItemData(cboGroupN6)

1770        mDevice.NG1D = Val(txtNG1D.text)
1780        mDevice.NG2D = Val(txtNG2D.text)
1790        mDevice.NG3D = Val(txtNG3D.text)
1800        mDevice.NG4D = Val(txtNG4D.text)
1810        mDevice.NG5D = Val(txtNG5D.text)
1820        mDevice.NG6D = Val(txtNG6D.text)


1830        mDevice.NG1_A = GetComboItemData(cboGroupN1_A)
1840        mDevice.NG2_A = GetComboItemData(cboGroupN2_A)
1850        mDevice.NG3_A = GetComboItemData(cboGroupN3_A)
1860        mDevice.NG4_A = GetComboItemData(cboGroupN4_A)
1870        mDevice.NG5_A = GetComboItemData(cboGroupN5_A)
1880        mDevice.NG6_A = GetComboItemData(cboGroupN6_A)

1890        mDevice.NG1_AD = Val(txtNG1_AD.text)
1900        mDevice.NG2_AD = Val(txtNG2_AD.text)
1910        mDevice.NG3_AD = Val(txtNG3_AD.text)
1920        mDevice.NG4_AD = Val(txtNG4_AD.text)
1930        mDevice.NG5_AD = Val(txtNG5_AD.text)
1940        mDevice.NG6_AD = Val(txtNG6_AD.text)


1950        mDevice.NG1_B = GetComboItemData(cboGroupN1_B)
1960        mDevice.NG2_B = GetComboItemData(cboGroupN2_B)
1970        mDevice.NG3_B = GetComboItemData(cboGroupN3_B)
1980        mDevice.NG4_B = GetComboItemData(cboGroupN4_B)
1990        mDevice.NG5_B = GetComboItemData(cboGroupN5_B)
2000        mDevice.NG6_B = GetComboItemData(cboGroupN6_B)

2010        mDevice.NG1_BD = Val(txtNG1_BD.text)
2020        mDevice.NG2_BD = Val(txtNG2_BD.text)
2030        mDevice.NG3_BD = Val(txtNG3_BD.text)
2040        mDevice.NG4_BD = Val(txtNG4_BD.text)
2050        mDevice.NG5_BD = Val(txtNG5_BD.text)
2060        mDevice.NG6_BD = Val(txtNG6_BD.text)



2070        mDevice.GG1 = GetComboItemData(cboGroupG1)
2080        mDevice.GG2 = GetComboItemData(cboGroupG2)
2090        mDevice.GG3 = GetComboItemData(cboGroupG3)
2100        mDevice.GG4 = GetComboItemData(cboGroupG4)
2110        mDevice.GG5 = GetComboItemData(cboGroupG5)
2120        mDevice.GG6 = GetComboItemData(cboGroupG6)

2130        mDevice.GG1D = Val(txtGG1D.text)
2140        mDevice.GG2D = Val(txtGG2d.text)
2150        mDevice.GG3D = Val(txtGG3d.text)
2160        mDevice.GG4D = Val(txtGG4d.text)
2170        mDevice.GG5D = Val(txtGG5d.text)
2180        mDevice.GG6D = Val(txtGG6d.text)


2190        mDevice.GG1_A = GetComboItemData(cboGroupG1_A)
2200        mDevice.GG2_A = GetComboItemData(cboGroupG2_A)
2210        mDevice.GG3_A = GetComboItemData(cboGroupG3_A)
2220        mDevice.GG4_A = GetComboItemData(cboGroupG4_A)
2230        mDevice.GG5_A = GetComboItemData(cboGroupG5_A)
2240        mDevice.GG6_A = GetComboItemData(cboGroupG6_A)

2250        mDevice.GG1_AD = Val(txtGG1_Ad.text)
2260        mDevice.GG2_AD = Val(txtGG2_Ad.text)
2270        mDevice.GG3_AD = Val(txtGG3_Ad.text)
2280        mDevice.GG4_AD = Val(txtGG4_Ad.text)
2290        mDevice.GG5_AD = Val(txtGG5_Ad.text)
2300        mDevice.GG6_AD = Val(txtGG6_Ad.text)



2310        mDevice.GG1_B = GetComboItemData(cboGroupG1_B)
2320        mDevice.GG2_B = GetComboItemData(cboGroupG2_B)
2330        mDevice.GG3_B = GetComboItemData(cboGroupG3_B)
2340        mDevice.GG4_B = GetComboItemData(cboGroupG4_B)
2350        mDevice.GG5_B = GetComboItemData(cboGroupG5_B)
2360        mDevice.GG6_B = GetComboItemData(cboGroupG6_B)

2370        mDevice.GG1_BD = Val(txtGG1_BD.text)
2380        mDevice.GG2_BD = Val(txtGG2_BD.text)
2390        mDevice.GG3_BD = Val(txtGG3_BD.text)
2400        mDevice.GG4_BD = Val(txtGG4_BD.text)
2410        mDevice.GG5_BD = Val(txtGG5_BD.text)
2420        mDevice.GG6_BD = Val(txtGG6_BD.text)


2430        mDevice.RepeatUntil = IIf(chkRepeatUntil.Value, 1, 0)
2440        mDevice.Repeats = Val(txtRepeats.text)
2450        mDevice.Pause = Val(txtPause.text)

2460        mDevice.RepeatUntil_A = IIf(chkRepeatUntil_A.Value, 1, 0)
2470        mDevice.Repeats_A = Val(txtRepeats_A.text)
2480        mDevice.Pause_A = Val(txtPause_A.text)


2490        mDevice.RepeatUntil_B = IIf(chkRepeatUntil_B.Value, 1, 0)
2500        mDevice.Repeats_B = Val(txtRepeats_B.text)
2510        mDevice.Pause_B = Val(txtPause_B.text)



2520        mDevice.AlarmMask = IIf(chkAlarmAlert.Value, 1, 0)
2530        mDevice.AlarmMask = IIf(chkExtern.Value, 2, mDevice.AlarmMask)
2540        mDevice.SendCancel = IIf(chkSendCancel.Value, 1, 0)

2550        mDevice.AlarmMask_A = IIf(chkAlarmAlert_A.Value, 1, 0)
2560        mDevice.AlarmMask_A = IIf(chkExtern_A.Value, 2, mDevice.AlarmMask_A)
2570        mDevice.SendCancel_A = IIf(chkSendCancel_A.Value, 1, 0)

2580        mDevice.AlarmMask_B = IIf(chkAlarmAlert_B.Value, 1, 0)
            'mDevice.AlarmMask_B = IIf(chkExtern_b.Value, 2, mDevice.AlarmMask_B)
2590        mDevice.SendCancel_B = IIf(chkSendCancel_B.Value, 1, 0)



2600        mDevice.AssurSecure = chkVacationSuper.Value
2610        mDevice.AssurSecure_A = chkVacationSuper_A.Value
            'mDevice.AssurSecure_B = chkVacationSuper_b.Value

2620        mDevice.Ignored = IIf(chkIgnore.Value, 1, 0)

            ' Temperature device settings not done here

            mDevice.Configurationstring = Trim$(txtClearingDevice.text)

2630        mDevice.Custom = IIf((Trim$(txtCustom.text) <> lblTypeDesc.Caption) And (Len(Trim$(txtCustom.text)) > 0), Trim$(txtCustom.text), lblTypeDesc.Caption)

2640        mDevice.IDL = Max(0, cboDeviceMode.ListIndex)

            ' ************ end common

            'On Error Resume Next

2650        adding = (DeviceID = 0)

            'mDevice.ZoneID = 44 ' RegisterDevice(mDevice) '

2660        If SaveDevice(mDevice, gUser.Username) Then  ' ok, it saved
2670          DeviceID = mDevice.DeviceID  ' local deviceid variable
2680        End If


2690        If (DeviceID = 0) Or (Err.Number <> 0) Then
2700          Save = False
2710        Else
2720          Save = True

2730          If USE6080 Then
              If (left$(mDevice.Model, 2) = "EN") Or (left$(mDevice.Model, 2) = "ES") Then
2740            If mDevice.ZoneID Then
2750              Select Case mDevice.IDL
                    Case 1  ' mobile
2760                  ChangeZoneParameter mDevice.ZoneID, "Locatable", "true"
2770                Case 2  ' fixed
2780                  ChangeZoneParameter mDevice.ZoneID, "IsRef", "true"
2790                Case 3  ' ispdevice
2800                  ChangeZoneParameter mDevice.ZoneID, "IsSPDevice", "true"
2810                Case Else
2820                  ChangeZoneParameter mDevice.ZoneID, "Null", "true"

2830              End Select



                  Dim ActualIDL As Long
                 ' DelayLoop 1
2840              ActualIDL = GetZoneIDL(mDevice.ZoneID) ' blew an error here
2850              If ActualIDL <> mDevice.IDL Then
2860                mDevice.IDL = ActualIDL
2870                SQL = "UPDATE devices SET IDL = " & mDevice.IDL & " WHERE serial = '" & mDevice.Serial & "'"
2880                ConnExecute SQL
2890              End If

2900            End If ' mDevice.ZoneID
              End If
2910          End If ' USE6080
2920          Fill

2930          If adding Then
          Dim NewDevice  As cESDevice:
2940            Set NewDevice = New cESDevice
2950            NewDevice.Serial = Serial
2960            Devices.AddDevice NewDevice
2970            Set mDevice = NewDevice
2980          End If

2990          Devices.RefreshBySerial mDevice.Serial  ' make sure all params are set

3000          Set mDevice = Devices.Devices(mDevice.Serial)
3010          SetupSerialDevice mDevice  ' if it's a serial device, set it up too
3020          If USE6080 = 0 Then
3030            Select Case UCase(mDevice.Model)
                  Case "EN5000", "EN5040", "EN5081"
3040                Outbounds.AddMessage mDevice.Serial, MSGTYPE_REPEATERNID, "", 0
                    ' create outbound message to set NID
3050              Case "EN3954"
3060                Outbounds.AddMessage mDevice.Serial, MSGTYPE_TWOWAYNID, "", 0
                    ' create outbound message to set NID
3070            End Select
3080          Else

3090          End If



3100        End If
3110      End If

3120    Else    ' (NOT) IF MASTER REMOTE ONLY *************************************************



3130      txtSerial.text = Trim(txtSerial.text)
3140      If Len(txtSerial.text) < 8 Then
3150        Beep
3160        lblAlert.Caption = "Serial Number Must Be 8 Characters"
3170        Exit Function
3180      End If
3190      index = cboDeviceType.ListIndex
3200      If index < 0 Then
3210        Beep
3220        lblAlert.Caption = "Device Type Not Selected"
3230        Exit Function
3240      End If

3250      Serial = UCase(Trim(txtSerial.text))
3260      Model = UCase(cboDeviceType.text)


3270      SQL = "SELECT COUNT(deviceid) FROM Devices WHERE serial = " & q(Serial)
3280      Set rs = ConnExecute(SQL)
3290      Count = rs(0)
          Dim TempID           As Long

3300      rs.Close
3310      Set rs = Nothing

3320      If Count > 0 And DeviceID = 0 Then
3330        Beep
3340        lblAlert.Caption = "Duplicate Serial Number ... Please Verify Serial Number"
3350        Exit Function
3360      End If


3370      mDevice.Serial = Serial
3380      mDevice.Model = Model


3390      mDevice.Announce = Trim(txtAnnounce.text)
3400      mDevice.Announce_A = Trim(txtAnnounce2.text)
3410      mDevice.Announce_B = Trim(txtAnnounce3.text)

3420      mDevice.ClearByReset = chkClearByReset.Value
          ' singletons
3430      Select Case mDeviceType.NumInputs
            Case 1
3440          mDevice.AssurInput = 1
3450        Case 2
3460          If optAssur1.Value Then
3470            mDevice.AssurInput = 1
3480          ElseIf optAssur2.Value Then
3490            mDevice.AssurInput = 2
3500          Else
3510            mDevice.AssurInput = 1
3520          End If
3530        Case Else
3540          mDevice.AssurInput = 0
3550      End Select

3560      mDevice.NumInputs = mDeviceType.NumInputs
3570      mDevice.IsPortable = mDeviceType.Portable


3580      mDevice.AutoClear = mDeviceType.AutoClear
          ' ************ common

3590      mDevice.OG1 = GetComboItemData(cboGroup1)
3600      mDevice.OG2 = GetComboItemData(cboGroup2)
3610      mDevice.OG3 = GetComboItemData(cboGroup3)
3620      mDevice.OG4 = GetComboItemData(cboGroup4)
3630      mDevice.OG5 = GetComboItemData(cboGroup5)
3640      mDevice.OG6 = GetComboItemData(cboGroup6)

3650      mDevice.OG1D = Val(txtOG1D.text)
3660      mDevice.OG2D = Val(txtOG2D.text)
3670      mDevice.OG3D = Val(txtOG3D.text)
3680      mDevice.OG4D = Val(txtOG4D.text)
3690      mDevice.OG5D = Val(txtOG5D.text)
3700      mDevice.OG6D = Val(txtOG6D.text)





3710      mDevice.OG1_A = GetComboItemData(cboGroup1_A)
3720      mDevice.OG2_A = GetComboItemData(cboGroup2_A)
3730      mDevice.OG3_A = GetComboItemData(cboGroup3_A)
3740      mDevice.OG4_A = GetComboItemData(cboGroup4_A)
3750      mDevice.OG5_A = GetComboItemData(cboGroup5_A)
3760      mDevice.OG6_A = GetComboItemData(cboGroup6_A)

3770      mDevice.OG1_AD = Val(txtOG1_AD.text)
3780      mDevice.OG2_AD = Val(txtOG2_AD.text)
3790      mDevice.OG3_AD = Val(txtOG3_AD.text)
3800      mDevice.OG4_AD = Val(txtOG4_AD.text)
3810      mDevice.OG5_AD = Val(txtOG5_AD.text)
3820      mDevice.OG6_AD = Val(txtOG6_AD.text)


3830      mDevice.OG1_B = GetComboItemData(cboGroup1_B)
3840      mDevice.OG2_B = GetComboItemData(cboGroup2_B)
3850      mDevice.OG3_B = GetComboItemData(cboGroup3_B)
3860      mDevice.OG4_B = GetComboItemData(cboGroup4_B)
3870      mDevice.OG5_B = GetComboItemData(cboGroup5_B)
3880      mDevice.OG6_B = GetComboItemData(cboGroup6_B)

3890      mDevice.OG1_AD = Val(txtOG1_AD.text)
3900      mDevice.OG2_AD = Val(txtOG2_AD.text)
3910      mDevice.OG3_AD = Val(txtOG3_AD.text)
3920      mDevice.OG4_AD = Val(txtOG4_AD.text)
3930      mDevice.OG5_AD = Val(txtOG5_AD.text)
3940      mDevice.OG6_AD = Val(txtOG6_AD.text)


3950      mDevice.OG1_BD = Val(txtOG1_BD.text)
3960      mDevice.OG2_BD = Val(txtOG2_BD.text)
3970      mDevice.OG3_BD = Val(txtOG3_BD.text)
3980      mDevice.OG4_BD = Val(txtOG4_BD.text)
3990      mDevice.OG5_BD = Val(txtOG5_BD.text)
4000      mDevice.OG6_BD = Val(txtOG6_BD.text)



4010      mDevice.NG1 = GetComboItemData(cboGroupN1)
4020      mDevice.NG2 = GetComboItemData(cboGroupN2)
4030      mDevice.NG3 = GetComboItemData(cboGroupN3)
4040      mDevice.NG4 = GetComboItemData(cboGroupN4)
4050      mDevice.NG5 = GetComboItemData(cboGroupN5)
4060      mDevice.NG6 = GetComboItemData(cboGroupN6)

4070      mDevice.NG1D = Val(txtNG1D.text)
4080      mDevice.NG2D = Val(txtNG2D.text)
4090      mDevice.NG3D = Val(txtNG3D.text)
4100      mDevice.NG4D = Val(txtNG4D.text)
4110      mDevice.NG5D = Val(txtNG5D.text)
4120      mDevice.NG6D = Val(txtNG6D.text)


4130      mDevice.NG1_A = GetComboItemData(cboGroupN1_A)
4140      mDevice.NG2_A = GetComboItemData(cboGroupN2_A)
4150      mDevice.NG3_A = GetComboItemData(cboGroupN3_A)
4160      mDevice.NG4_A = GetComboItemData(cboGroupN4_A)
4170      mDevice.NG5_A = GetComboItemData(cboGroupN5_A)
4180      mDevice.NG6_A = GetComboItemData(cboGroupN6_A)

4190      mDevice.NG1_AD = Val(txtNG1_AD.text)
4200      mDevice.NG2_AD = Val(txtNG2_AD.text)
4210      mDevice.NG3_AD = Val(txtNG3_AD.text)
4220      mDevice.NG4_AD = Val(txtNG4_AD.text)
4230      mDevice.NG5_AD = Val(txtNG5_AD.text)
4240      mDevice.NG6_AD = Val(txtNG6_AD.text)



4250      mDevice.NG1_B = GetComboItemData(cboGroupN1_B)
4260      mDevice.NG2_B = GetComboItemData(cboGroupN2_B)
4270      mDevice.NG3_B = GetComboItemData(cboGroupN3_B)
4280      mDevice.NG4_B = GetComboItemData(cboGroupN4_B)
4290      mDevice.NG5_B = GetComboItemData(cboGroupN5_B)
4300      mDevice.NG6_B = GetComboItemData(cboGroupN6_B)

4310      mDevice.NG1_BD = Val(txtNG1_BD.text)
4320      mDevice.NG2_BD = Val(txtNG2_BD.text)
4330      mDevice.NG3_BD = Val(txtNG3_BD.text)
4340      mDevice.NG4_BD = Val(txtNG4_BD.text)
4350      mDevice.NG5_BD = Val(txtNG5_BD.text)
4360      mDevice.NG6_BD = Val(txtNG6_BD.text)


4370      mDevice.GG1 = GetComboItemData(cboGroupG1)
4380      mDevice.GG2 = GetComboItemData(cboGroupG2)
4390      mDevice.GG3 = GetComboItemData(cboGroupG3)
4400      mDevice.GG4 = GetComboItemData(cboGroupG4)
4410      mDevice.GG5 = GetComboItemData(cboGroupG5)
4420      mDevice.GG6 = GetComboItemData(cboGroupG6)

4430      mDevice.GG1D = Val(txtGG1D.text)
4440      mDevice.GG2D = Val(txtGG2d.text)
4450      mDevice.GG3D = Val(txtGG3d.text)
4460      mDevice.GG4D = Val(txtGG4d.text)
4470      mDevice.GG5D = Val(txtGG5d.text)
4480      mDevice.GG6D = Val(txtGG6d.text)


4490      mDevice.GG1_A = GetComboItemData(cboGroupG1_A)
4500      mDevice.GG2_A = GetComboItemData(cboGroupG2_A)
4510      mDevice.GG3_A = GetComboItemData(cboGroupG3_A)
4520      mDevice.GG4_A = GetComboItemData(cboGroupG4_A)
4530      mDevice.GG5_A = GetComboItemData(cboGroupG5_A)
4540      mDevice.GG6_A = GetComboItemData(cboGroupG6_A)

4550      mDevice.GG1_AD = Val(txtGG1_Ad.text)
4560      mDevice.GG2_AD = Val(txtGG2_Ad.text)
4570      mDevice.GG3_AD = Val(txtGG3_Ad.text)
4580      mDevice.GG4_AD = Val(txtGG4_Ad.text)
4590      mDevice.GG5_AD = Val(txtGG5_Ad.text)
4600      mDevice.GG6_AD = Val(txtGG6_Ad.text)



4610      mDevice.GG1_B = GetComboItemData(cboGroupG1_B)
4620      mDevice.GG2_B = GetComboItemData(cboGroupG2_B)
4630      mDevice.GG3_B = GetComboItemData(cboGroupG3_B)
4640      mDevice.GG4_B = GetComboItemData(cboGroupG4_B)
4650      mDevice.GG5_B = GetComboItemData(cboGroupG5_B)
4660      mDevice.GG6_B = GetComboItemData(cboGroupG6_B)

4670      mDevice.GG1_BD = Val(txtGG1_BD.text)
4680      mDevice.GG2_BD = Val(txtGG2_BD.text)
4690      mDevice.GG3_BD = Val(txtGG3_BD.text)
4700      mDevice.GG4_BD = Val(txtGG4_BD.text)
4710      mDevice.GG5_BD = Val(txtGG5_BD.text)
4720      mDevice.GG6_BD = Val(txtGG6_BD.text)





4730      mDevice.RepeatUntil = IIf(chkRepeatUntil.Value, 1, 0)
4740      mDevice.Repeats = Val(txtRepeats.text)
4750      mDevice.Pause = Val(txtPause.text)

4760      mDevice.RepeatUntil_A = IIf(chkRepeatUntil_A.Value, 1, 0)
4770      mDevice.Repeats_A = Val(txtRepeats_A.text)
4780      mDevice.Pause_A = Val(txtPause_A.text)

4790      mDevice.RepeatUntil_B = IIf(chkRepeatUntil_B.Value, 1, 0)
4800      mDevice.Repeats_B = Val(txtRepeats_B.text)
4810      mDevice.Pause_B = Val(txtPause_B.text)

4820      mDevice.AlarmMask = IIf(chkAlarmAlert.Value, 1, 0)
4830      mDevice.AlarmMask = IIf(chkExtern.Value, 2, mDevice.AlarmMask)
4840      mDevice.SendCancel = IIf(chkSendCancel.Value, 1, 0)

4850      mDevice.AlarmMask_A = IIf(chkAlarmAlert_A.Value, 1, 0)
4860      mDevice.AlarmMask_A = IIf(chkExtern_A.Value, 2, mDevice.AlarmMask_A)
4870      mDevice.SendCancel_A = IIf(chkSendCancel_A.Value, 1, 0)

4880      mDevice.AlarmMask_B = IIf(chkAlarmAlert_B.Value, 1, 0)
          'mDevice.AlarmMask_B = IIf(chkExtern_b.Value, 2, mDevice.AlarmMask_B)
4890      mDevice.SendCancel_B = IIf(chkSendCancel_B.Value, 1, 0)


4900      mDevice.AssurSecure = chkVacationSuper.Value
4910      mDevice.AssurSecure_A = chkVacationSuper_A.Value

4920      mDevice.Ignored = IIf(chkIgnore.Value, 1, 0)
4930      '


          mDevice.Configurationstring = txtClearingDevice.text
          ' ************ end common


4940      dbgHostRemote "saving via remote"
  
4950      If RemoteUpdateDevice(mDevice) = 0 Then    ' ********* THIS IS WHERE IT REALLY HAPPENS FOR THE REMOTE
4960        dbgHostRemote "saving via remote success" & vbCrLf
4970        Save = True
            'Select Case UCase(mDevice.model)
            '  Case "EN5000"
            'RemoteRequestOutbounds.AddMessage mDevice.serial, MSGTYPE_REPEATERNID, "", 0
            ' create outbound message to set NID
            '  Case "EN3954"
            'RemoteRequestOutbounds.AddMessage mDevice.serial, MSGTYPE_TWOWAYNID, "", 0
            ' create outbound message to set NID
            'End Select
4980      Else
4990        dbgHostRemote "saving via remote failure" & vbCrLf
5000      End If    'RemoteUpdateDevice(mDevice) Then


5010      DeviceID = GetDeviceIDFromSerial(mDevice.Serial)    ' assumes we had success saving

5020      Fill    ' get data via ADO back channel
5030    End If    'If Master


5040    If MASTER Then
5050      If USE6080 Then
            'Debug.Assert mDevice.ZoneID
            If (left$(mDevice.Model, 2) = "EN") Or (left$(mDevice.Model, 2) = "ES") Then
            
            
5060        Set HTTPRequest = New cHTTPRequest

5070        Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, mDevice.ZoneID)
5080        Set HTTPRequest = Nothing
5090        If Not ZoneInfo Is Nothing Then
5100          If 0 = StrComp(ZoneInfo.IsMissing, "true", vbTextCompare) Then
5110            SCICode = 0
5120            Select Case ZoneInfo.MID
                  Case &HB2
5130                SCICode = SCI_CODE_DEVICE_INACTIVE
5140              Case Else
5150                If InStr(1, ZoneInfo.TypeName, "6080") > 0 Then
5160                  SCICode = 0
5170                Else
5180                  SCICode = SCI_CODE_REPEATER_INACTIVE
5190                End If
5200            End Select
5210            If SCICode Then
5220              ProcessESPacket MakeFakePacket(ZoneInfo, SCICode)
5230            End If
5240          End If
5250        End If
            End If
5260      End If
5270    End If


Save_Resume:
5280    On Error GoTo 0
5290    Exit Function

Save_Error:

5300    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.Save." & Erl
5310    Resume Save_Resume


End Function

'Sub SetCodeVisiblility(ByVal Value As Boolean)
'  txtSerial.Locked = Not (Value)
'  'lblIDM.Visible = Value
'  'txtIDM.Visible = Value
'  'lblIDL.Visible = Value
'  'txtIDL.Visible = Value
'  'lblIDL.Visible = Value
'
'
'End Sub

Sub SetControls()
  On Error GoTo setControls_Error

  fraEnabler.BackColor = Me.BackColor

  fraTx.left = TabStrip.ClientLeft
  fraTx.top = TabStrip.ClientTop
  fraTx.Height = TabStrip.ClientHeight
  fraTx.Width = TabStrip.ClientWidth
  fraTx.BackColor = Me.BackColor

  fraassur.left = TabStrip.ClientLeft
  fraassur.top = TabStrip.ClientTop
  fraassur.Height = TabStrip.ClientHeight
  fraassur.Width = TabStrip.ClientWidth
  fraassur.BackColor = Me.BackColor


  fraOutput.left = TabStrip.ClientLeft
  fraOutput.top = TabStrip.ClientTop
  fraOutput.Height = TabStrip.ClientHeight
  fraOutput.Width = TabStrip.ClientWidth
  fraOutput.BackColor = Me.BackColor

  fraOutput_A.left = TabStrip.ClientLeft
  fraOutput_A.top = TabStrip.ClientTop
  fraOutput_A.Height = TabStrip.ClientHeight
  fraOutput_A.Width = TabStrip.ClientWidth
  fraOutput_A.BackColor = Me.BackColor
  
  fraOutput_B.left = TabStrip.ClientLeft
  fraOutput_B.top = TabStrip.ClientTop
  fraOutput_B.Height = TabStrip.ClientHeight
  fraOutput_B.Width = TabStrip.ClientWidth
  fraOutput_B.BackColor = Me.BackColor


  fraTimes.left = TabStrip.ClientLeft
  fraTimes.top = TabStrip.ClientTop
  fraTimes.Height = TabStrip.ClientHeight
  fraTimes.Width = TabStrip.ClientWidth
  fraTimes.BackColor = Me.BackColor

  fraInput.left = chkClearByReset.left
  fraInput_A.left = fraInput.left
  fraInput_A.Width = TabStrip.ClientWidth
  fraInput_A.top = fraInput.top


  fraInput_B.left = fraInput.left
  fraInput_B.Width = TabStrip.ClientWidth
  fraInput_B.top = fraInput.top



  fraPartitions.left = TabStrip.ClientLeft
  fraPartitions.top = TabStrip.ClientTop
  fraPartitions.Height = TabStrip.ClientHeight
  fraPartitions.Width = TabStrip.ClientWidth

  fraPartitions.BackColor = Me.BackColor

  fraInput.BackColor = Me.BackColor
  fraInput_A.BackColor = Me.BackColor
  fraInput_B.BackColor = Me.BackColor

  fraTx.Visible = True
  
  fraOutput.Visible = False
  fraOutput_A.Visible = False
  fraOutput_B.Visible = False
  
  fraTimes.Visible = False
  fraassur.Visible = False
  

  fraDisable.BackColor = Me.BackColor
  fraDisable_A.BackColor = Me.BackColor
  fraDisable_B.BackColor = Me.BackColor

  fraDisable_A.left = fraDisable.left
  fraDisable_A.top = fraDisable.top

  fraDisable_B.left = fraDisable.left
  fraDisable_B.top = fraDisable.top

  fraIgnore.Visible = (gUser.LEvel >= LEVEL_ADMIN)
  

  lvAvailPartitions.ColumnHeaders(1).Width = 800
  lvAvailPartitions.ColumnHeaders(2).Width = 2250
  lvAvailPartitions.ColumnHeaders(3).Width = 500
  lvAvailPartitions.ColumnHeaders(3).text = "L"


setControls_Resume:
  On Error GoTo 0
  Exit Sub

setControls_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.setControls." & Erl
  Resume setControls_Resume


End Sub


Sub SetDeviceByModel(ByVal Serial As String, ByVal Model As String)
  Dim i             As Long
  txtSerial.text = Serial
  
  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If

  
  For i = cboDeviceType.listcount - 1 To 1 Step -1
    If 0 = StrComp(cboDeviceType.list(i), Model, vbTextCompare) Then
      Exit For
    End If
  Next
  cboDeviceType.ListIndex = i
  '  i = CboGetIndexByItemData(cboDeviceType, CLSPTI)  ' was midpti
  '  If i < 0 Then i = 0
  '  cboDeviceType.ListIndex = i
  '
End Sub


Sub SetDevice(ByVal Serial As String, ByVal CLSPTI As Long)
  txtSerial.text = Serial
  Dim i             As Long
  i = CboGetIndexByItemData(cboDeviceType, CLSPTI)  ' was midpti
  If i < 0 Then i = 0
  cboDeviceType.ListIndex = i

End Sub

Sub SetNewDeviceType()
  Dim index         As Integer
10 On Error GoTo SetNewDeviceType_Error

20 index = cboDeviceType.ListIndex
30 txtModel.text = cboDeviceType.text

40 If index > -1 Then
    'mDeviceType = GetDeviceTypeByMIDPTI(cboDeviceType.ItemData(index))
50  mDeviceType = GetDeviceTypeByModel(cboDeviceType.text)
60  lblTypeDesc.Caption = mDeviceType.desc
70  chkClearByReset.Visible = IIf(mDeviceType.ClearByReset = 1, True, False)
80  chkClearByReset.Value = mDeviceType.ClearByReset        '= 1

90  If DeviceID = 0 Then
100   txtAnnounce.text = mDeviceType.Announce

110   txtAnnounce2.text = mDeviceType.Announce2

111   txtAnnounce3.text = "" ' mDeviceType.Announce3

120   txtRepeats.text = mDeviceType.Repeats
130   txtRepeats_A.text = mDeviceType.Repeats_A

140   mDevice.Repeats = mDeviceType.Repeats
150   mDevice.Repeats_A = mDeviceType.Repeats_A

160   txtPause.text = mDeviceType.Pause
170   txtPause_A.text = mDeviceType.Pause_A

180   mDevice.Pause = mDeviceType.Pause
190   mDevice.Pause_A = mDeviceType.Pause_A

      'lblTypeDesc.Caption = mDeviceType.Desc

200   chkRepeatUntil.Value = mDeviceType.RepeatUntil
210   chkRepeatUntil_A.Value = mDeviceType.RepeatUntil_A
220   mDevice.RepeatUntil = mDeviceType.RepeatUntil
230   mDevice.RepeatUntil_A = mDeviceType.RepeatUntil_A

240   chkSendCancel.Value = mDeviceType.SendCancel
250   chkSendCancel_A.Value = mDeviceType.SendCancel_A
260   mDevice.SendCancel = mDeviceType.SendCancel
270   mDevice.SendCancel_A = mDeviceType.SendCancel_A

280   mDevice.IgnoreTamper = mDeviceType.IgnoreTamper
290   chkIgnoreTamper.Value = IIf(mDevice.IgnoreTamper = 1, 1, 0)

300   mDevice.OG1 = mDeviceType.OG1
310   mDevice.OG2 = mDeviceType.OG2
320   mDevice.OG3 = mDeviceType.OG3
330   mDevice.OG4 = mDeviceType.OG4
340   mDevice.OG5 = mDeviceType.OG5
350   mDevice.OG6 = mDeviceType.OG6


360   mDevice.NG1 = mDeviceType.NG1
370   mDevice.NG2 = mDeviceType.NG2
380   mDevice.NG3 = mDeviceType.NG3
390   mDevice.NG4 = mDeviceType.NG4
400   mDevice.NG5 = mDeviceType.NG5
410   mDevice.NG6 = mDeviceType.NG6

420   mDevice.GG1 = mDeviceType.GG1
430   mDevice.GG2 = mDeviceType.GG2
440   mDevice.GG3 = mDeviceType.GG3
450   mDevice.GG4 = mDeviceType.GG4
460   mDevice.GG5 = mDeviceType.GG5
470   mDevice.GG6 = mDeviceType.GG6



480   mDevice.OG1_A = mDeviceType.OG1_A
490   mDevice.OG2_A = mDeviceType.OG2_A
500   mDevice.OG3_A = mDeviceType.OG3_A
510   mDevice.OG4_A = mDeviceType.OG4_A
520   mDevice.OG5_A = mDeviceType.OG5_A
530   mDevice.OG6_A = mDeviceType.OG6_A


540   mDevice.NG1_A = mDeviceType.NG1_A
550   mDevice.NG2_A = mDeviceType.NG2_A
560   mDevice.NG3_A = mDeviceType.NG3_A
570   mDevice.NG4_A = mDeviceType.NG4_A
580   mDevice.NG5_A = mDeviceType.NG5_A
590   mDevice.NG6_A = mDeviceType.NG6_A


600   mDevice.GG1_A = mDeviceType.GG1_A
610   mDevice.GG2_A = mDeviceType.GG2_A
620   mDevice.GG3_A = mDeviceType.GG3_A
630   mDevice.GG4_A = mDeviceType.GG4_A
640   mDevice.GG5_A = mDeviceType.GG5_A
650   mDevice.GG6_A = mDeviceType.GG6_A





660   cboGroup1.ListIndex = CboGetIndexByItemData(cboGroup1, mDevice.OG1)
670   cboGroup2.ListIndex = CboGetIndexByItemData(cboGroup2, mDevice.OG2)
680   cboGroup3.ListIndex = CboGetIndexByItemData(cboGroup3, mDevice.OG3)
690   cboGroup4.ListIndex = CboGetIndexByItemData(cboGroup4, mDevice.OG4)
700   cboGroup5.ListIndex = CboGetIndexByItemData(cboGroup5, mDevice.OG5)
710   cboGroup6.ListIndex = CboGetIndexByItemData(cboGroup6, mDevice.OG6)

720   cboGroupN1.ListIndex = CboGetIndexByItemData(cboGroupN1, mDevice.NG1)
730   cboGroupN2.ListIndex = CboGetIndexByItemData(cboGroupN2, mDevice.NG2)
740   cboGroupN3.ListIndex = CboGetIndexByItemData(cboGroupN3, mDevice.NG3)
750   cboGroupN4.ListIndex = CboGetIndexByItemData(cboGroupN4, mDevice.NG4)
760   cboGroupN5.ListIndex = CboGetIndexByItemData(cboGroupN5, mDevice.NG5)
770   cboGroupN6.ListIndex = CboGetIndexByItemData(cboGroupN6, mDevice.NG6)

780   cboGroupG1.ListIndex = CboGetIndexByItemData(cboGroupG1, mDevice.GG1)
790   cboGroupG2.ListIndex = CboGetIndexByItemData(cboGroupG2, mDevice.GG2)
800   cboGroupG3.ListIndex = CboGetIndexByItemData(cboGroupG3, mDevice.GG3)
810   cboGroupG4.ListIndex = CboGetIndexByItemData(cboGroupG4, mDevice.GG4)
820   cboGroupG5.ListIndex = CboGetIndexByItemData(cboGroupG5, mDevice.GG5)
830   cboGroupG6.ListIndex = CboGetIndexByItemData(cboGroupG6, mDevice.GG6)



840   cboGroup1_A.ListIndex = CboGetIndexByItemData(cboGroup1_A, mDevice.OG1_A)
850   cboGroup2_A.ListIndex = CboGetIndexByItemData(cboGroup2_A, mDevice.OG2_A)
860   cboGroup3_A.ListIndex = CboGetIndexByItemData(cboGroup3_A, mDevice.OG3_A)
870   cboGroup4_A.ListIndex = CboGetIndexByItemData(cboGroup4_A, mDevice.OG4_A)
880   cboGroup5_A.ListIndex = CboGetIndexByItemData(cboGroup5_A, mDevice.OG5_A)
890   cboGroup6_A.ListIndex = CboGetIndexByItemData(cboGroup6_A, mDevice.OG6_A)

900   cboGroupN1_A.ListIndex = CboGetIndexByItemData(cboGroupN1_A, mDevice.NG1_A)
910   cboGroupN2_A.ListIndex = CboGetIndexByItemData(cboGroupN2_A, mDevice.NG2_A)
920   cboGroupN3_A.ListIndex = CboGetIndexByItemData(cboGroupN3_A, mDevice.NG3_A)
930   cboGroupN4_A.ListIndex = CboGetIndexByItemData(cboGroupN4_A, mDevice.NG4_A)
940   cboGroupN5_A.ListIndex = CboGetIndexByItemData(cboGroupN5_A, mDevice.NG5_A)
950   cboGroupN6_A.ListIndex = CboGetIndexByItemData(cboGroupN6_A, mDevice.NG6_A)

960   cboGroupG1_A.ListIndex = CboGetIndexByItemData(cboGroupG1_A, mDevice.GG1_A)
970   cboGroupG2_A.ListIndex = CboGetIndexByItemData(cboGroupG2_A, mDevice.GG2_A)
980   cboGroupG3_A.ListIndex = CboGetIndexByItemData(cboGroupG3_A, mDevice.GG3_A)
990   cboGroupG4_A.ListIndex = CboGetIndexByItemData(cboGroupG4_A, mDevice.GG4_A)
1000  cboGroupG5_A.ListIndex = CboGetIndexByItemData(cboGroupG5_A, mDevice.GG5_A)
1010  cboGroupG6_A.ListIndex = CboGetIndexByItemData(cboGroupG6_A, mDevice.GG6_A)



1020  mDevice.OG1D = mDeviceType.OG1D
1030  mDevice.OG2D = mDeviceType.OG2D
1040  mDevice.OG3D = mDeviceType.OG3D
1050  mDevice.OG4D = mDeviceType.OG4D
1060  mDevice.OG5D = mDeviceType.OG5D
1070  mDevice.OG6D = mDeviceType.OG6D


1080  mDevice.NG1D = mDeviceType.NG1D
1090  mDevice.NG2D = mDeviceType.NG2D
1100  mDevice.NG3D = mDeviceType.NG3D
1110  mDevice.NG4D = mDeviceType.NG4D
1120  mDevice.NG5D = mDeviceType.NG5D
1130  mDevice.NG6D = mDeviceType.NG6D

1140  mDevice.GG1D = mDeviceType.GG1D
1150  mDevice.GG2D = mDeviceType.GG2D
1160  mDevice.GG3D = mDeviceType.GG3D
1170  mDevice.GG4D = mDeviceType.GG4D
1180  mDevice.GG5D = mDeviceType.GG5D
1190  mDevice.GG6D = mDeviceType.GG6D


1200  mDevice.OG1_AD = mDeviceType.OG1_AD
1210  mDevice.OG2_AD = mDeviceType.OG2_AD
1220  mDevice.OG3_AD = mDeviceType.OG3_AD
1230  mDevice.OG4_AD = mDeviceType.OG4_AD
1240  mDevice.OG5_AD = mDeviceType.OG5_AD
1250  mDevice.OG6_AD = mDeviceType.OG6_AD


1260  mDevice.NG1_AD = mDeviceType.NG1_AD
1270  mDevice.NG2_AD = mDeviceType.NG2_AD
1280  mDevice.NG3_AD = mDeviceType.NG3_AD
1290  mDevice.NG4_AD = mDeviceType.NG4_AD
1300  mDevice.NG5_AD = mDeviceType.NG5_AD
1310  mDevice.NG6_AD = mDeviceType.NG6_AD

1320  mDevice.GG1_AD = mDeviceType.GG1_AD
1330  mDevice.GG2_AD = mDeviceType.GG2_AD
1340  mDevice.GG3_AD = mDeviceType.GG3_AD
1350  mDevice.GG4_AD = mDeviceType.GG4_AD
1360  mDevice.GG5_AD = mDeviceType.GG5_AD
1370  mDevice.GG6_AD = mDeviceType.GG6_AD


1380  txtOG1D.text = mDevice.OG1D
1390  txtOG2D.text = mDevice.OG2D
1400  txtOG3D.text = mDevice.OG3D
1410  txtOG4D.text = mDevice.OG4D
1420  txtOG5D.text = mDevice.OG5D
1430  txtOG6D.text = mDevice.OG6D

1440  txtNG1D.text = mDevice.NG1D
1450  txtNG2D.text = mDevice.NG2D
1460  txtNG3D.text = mDevice.NG3D
1470  txtNG4D.text = mDevice.NG4D
1480  txtNG5D.text = mDevice.NG5D
1490  txtNG6D.text = mDevice.NG6D

1500  txtGG1D.text = mDevice.GG1D
1510  txtGG2d.text = mDevice.GG2D
1520  txtGG3d.text = mDevice.GG3D
1530  txtGG4d.text = mDevice.GG4D
1540  txtGG5d.text = mDevice.GG5D
1550  txtGG6d.text = mDevice.GG6D


1560  txtOG1_AD.text = mDevice.OG1_AD
1570  txtOG2_AD.text = mDevice.OG2_AD
1580  txtOG3_AD.text = mDevice.OG3_AD
1590  txtOG4_AD.text = mDevice.OG4_AD
1600  txtOG5_AD.text = mDevice.OG5_AD
1610  txtOG6_AD.text = mDevice.OG6_AD

1620  txtNG1_AD.text = mDevice.NG1_AD
1630  txtNG2_AD.text = mDevice.NG2_AD
1640  txtNG3_AD.text = mDevice.NG3_AD
1650  txtNG4_AD.text = mDevice.NG4_AD
1660  txtNG5_AD.text = mDevice.NG5_AD
1670  txtNG6_AD.text = mDevice.NG6_AD

1680  txtGG1_Ad.text = mDevice.GG1_AD
1690  txtGG2_Ad.text = mDevice.GG2_AD
1700  txtGG3_Ad.text = mDevice.GG3_AD
1710  txtGG4_Ad.text = mDevice.GG4_AD
1720  txtGG5_Ad.text = mDevice.GG5_AD
1730  txtGG6_Ad.text = mDevice.GG6_AD




1740 End If

1750 End If

1760 Select Case mDeviceType.NumInputs
    Case 2, 3
1770  optInput1.Visible = True
1780  optInput2.Visible = True
      If mDeviceType.NumInputs = 3 Or mDevice.UseTamperAsInput Then ' Device.UseTamperAsInput
        optInput3.Visible = True
      Else
        optInput3.Visible = False
      End If
        
  

      'chkClearByReset.Visible = False
1790  optAssur1.Visible = True
1800  optAssur2.Visible = True
1810 Case 1
1820  optInput1.Visible = False
1830  optInput2.Visible = False
1840  optInput1.Value = True
1850  optAssur1.Visible = True
1860  optAssur2.Visible = False
      optInput3.Visible = False
1870 Case 0
1880  optInput1.Visible = False
1890  optInput2.Visible = False
1900  optAssur1.Visible = False
1910  optAssur2.Visible = False
      optInput3.Visible = False
1920  optInput1.Value = True
1930 End Select

1940 If cboDeviceType.text = COM_DEV_NAME Then
1950 chkClearByReset.Visible = False
1960 chkClearByReset.Value = 0
1970 cmdConfigureSerial.Visible = True
1980 cmdAutoEnroll.Visible = False And EnrollButtonVisible()
1990 Else
2000 chkClearByReset.Visible = True
2010 cmdAutoEnroll.Visible = True And EnrollButtonVisible()
2020 cmdConfigureSerial.Visible = False

2030 End If


2040 If mDeviceType.Portable Then
2050 cboDeviceMode.ListIndex = 1
2060 ElseIf mDeviceType.Fixed Then
2070 cboDeviceMode.ListIndex = 2
2080 Else
2090 cboDeviceMode.ListIndex = 0
2100 End If




SetNewDeviceType_Resume:
2110 On Error GoTo 0
2120 Exit Sub

SetNewDeviceType_Error:

2130 LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.SetNewDeviceType." & Erl
2140 Resume SetNewDeviceType_Resume


End Sub
Function FillActivePartitionList() As Long
        Dim part          As cPartition
        Dim li            As ListItem
        Dim Found         As Boolean
        Dim XML           As String

        '  Dim partitions    As Collection
        Dim HTTPRequest   As cHTTPRequest
        Dim ZoneInfo      As cZoneInfo

        'mDevice.ZoneID = 0

10      Set HTTPRequest = New cHTTPRequest
20      Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, mDevice.ZoneID)

        Dim AvailablePart As cPartition


30      lvPartitions.ListItems.Clear

40      If USE6080 Then

50        For Each part In ZoneInfo.Partitionlist


60          Set li = lvPartitions.ListItems.Add(, , part.PartitionID)
            ' need to get partition desc
70          On Error Resume Next
80          Set AvailablePart = Partitions(part.PartitionID & "")
90          If AvailablePart Is Nothing Then
100           li.SubItems(1) = "??"
110         Else
120           li.SubItems(1) = AvailablePart.Description
130         End If
            'li.SubItems(2) = IIf(part.IsLocation, "X", "")


140       Next

150       For Each li In lvPartitions.ListItems
160         li.Selected = 0
170       Next
180     End If

End Function
Sub ShowPanels()
  On Error GoTo ShowPanels_Error

  Select Case TabStrip.SelectedItem.Key

    Case "output"
      'If mInput2 Then
      Select Case InputNumber
        Case 3
          fraOutput_B.Visible = True
          fraOutput_A.Visible = False
          fraOutput.Visible = False
        Case 2
          fraOutput_A.Visible = True
          fraOutput.Visible = False
          fraOutput_B.Visible = False
        Case Else
          fraOutput.Visible = True
          fraOutput_A.Visible = False
          fraOutput_B.Visible = False
      End Select

      fraTx.Visible = False
      fraTimes.Visible = False
      fraassur.Visible = False
      fraPartitions.Visible = False

    Case "assure"
      fraassur.Visible = True
      fraTx.Visible = False
      fraOutput.Visible = False
      fraOutput_A.Visible = False
      fraOutput_B.Visible = False
      fraTimes.Visible = False
      fraPartitions.Visible = False

    Case "times"
      fraTimes.Visible = True
      fraTx.Visible = False
      fraOutput.Visible = False
      fraOutput_A.Visible = False
      fraOutput_B.Visible = False
      fraassur.Visible = False
      fraPartitions.Visible = False

      Select Case InputNumber
          ' TODO
        Case 3
          lblActive.Caption = "Select the time Alarm for input 3 is disabled"
          fraDisable_B.Visible = True
          fraDisable_A.Visible = False
          fraDisable.Visible = False
        Case 2
          lblActive.Caption = "Select the time Alarm for input 2 is disabled"
          fraDisable_A.Visible = True
          fraDisable_B.Visible = False
          fraDisable.Visible = False
        Case Else
          lblActive.Caption = "Select the time Alarm for input 1 is disabled"
          fraDisable.Visible = True
          fraDisable_B.Visible = False
          fraDisable_A.Visible = False
      End Select

    Case "partitions"
      If Save() Then

        fraPartitions.Visible = True
        fraTimes.Visible = False
        fraTx.Visible = False
        fraOutput.Visible = False
        fraOutput_A.Visible = False
        fraOutput_B.Visible = False
        fraassur.Visible = False
        
          FillAvailPartitionList  ' MUST CALL THIS FIRST
          FillActivePartitionList
        
      End If

    Case Else
      Select Case InputNumber
        Case 3
          txtAnnounce3.Visible = True
          txtAnnounce2.Visible = False
          txtAnnounce.Visible = False
          
          fraInput_B.Visible = True
          fraInput_A.Visible = False
          fraInput.Visible = False
        Case 2


          txtAnnounce2.Visible = True
          txtAnnounce3.Visible = False
          txtAnnounce.Visible = False

          fraInput_A.Visible = True
          fraInput.Visible = False
          fraInput_B.Visible = False
        
        Case Else
          txtAnnounce.Visible = True
          txtAnnounce2.Visible = False
          txtAnnounce3.Visible = False

          
          fraInput.Visible = True
          fraInput_A.Visible = False
          fraInput_B.Visible = False
      End Select

      fraTx.Visible = True
      fraOutput.Visible = False
      fraOutput_A.Visible = False
      fraOutput_B.Visible = False
      fraTimes.Visible = False
      fraassur.Visible = False
      fraPartitions.Visible = False

  End Select
  Select Case InputNumber
    Case 3
      lblAnnounce.Caption = "Announce 3"
    Case 2
      lblAnnounce.Caption = "Announce 2"
    Case Else

      lblAnnounce.Caption = "Announce"
  End Select

ShowPanels_Resume:
  On Error GoTo 0
  Exit Sub

ShowPanels_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTransmitter.ShowPanels." & Erl
  Resume ShowPanels_Resume


End Sub

Private Sub optInput3_Click()
  InputNumber = 3
  ShowPanels
  If Not bycode Then
    '  Display
  End If
End Sub

Private Sub TabStrip_Click()
  ShowPanels
End Sub


Private Sub tmrEnroller_Timer()
  CheckForNewDevice
End Sub
Function CheckForNewDevice()
  Dim xmlmessage    As String

10 If regws Is Nothing Then
20  DisableAutoEnroll
30 Else
    '  Debug.Print "Waiting for Device Reset"
40  If regws.StatusCode <> 1 Then
50    Debug.Print "Enroller Status " & regws.StatusCode
60  End If
70  If regws.HasMessages Then
80    Debug.Print "Enroller Has Device Reset"
90    xmlmessage = regws.GetNextMessage
100   DisableAutoEnroll
      Debug.Print xmlmessage
110   Set RegDevice = New cRegDevice
120   RegDevice.ParseXML xmlmessage
      
130   LoadRegDevice

140 End If
150 End If
End Function

Private Sub tmrSearch_Timer()
  tmrSearch.Enabled = False
  If USE6080 Then
    FillAvailPartitionList
  End If
End Sub

Private Sub txtAnnounce_GotFocus()
  On Error Resume Next
  ClearAlerts
  SelAll txtAnnounce
End Sub

Private Sub txtAnnounce_LostFocus()
'  If Input3 Then
'    mDevice.Announce_B = Trim(txtAnnounce3.text)
'  ElseIf Input2 Then
'    mDevice.Announce_A = Trim(txtAnnounce2.text)
'  Else
    mDevice.Announce = Trim(txtAnnounce.text)
'  End If
End Sub

Private Sub txtAnnounce2_GotFocus()
  On Error Resume Next
  ClearAlerts
  
  SelAll txtAnnounce2
End Sub

Private Sub txtAnnounce2_LostFocus()
'  If Input3 Then
'    mDevice.Announce_B = Trim(txtAnnounce3.text)
'  ElseIf Input2 Then
    mDevice.Announce_A = Trim(txtAnnounce2.text)
'  Else

End Sub

Private Sub txtAnnounce3_GotFocus()
  On Error Resume Next
  ClearAlerts
  
  SelAll txtAnnounce3
End Sub

Private Sub txtAnnounce3_LostFocus()

    mDevice.Announce_B = Trim(txtAnnounce3.text)

End Sub

Private Sub txtClearingDevice_GotFocus()
  SelAll txtClearingDevice
End Sub

Private Sub txtClearingDevice_KeyPress(KeyAscii As Integer)
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

Private Sub txtCustom_Click()
  SelAll txtCustom
End Sub

Private Sub txtDisableEnd_A_Change()
  ClearAlerts
  Dim t             As Integer
  t = Val(txtDisableEnd_A.text)
  If t < 12 Then
    If t = 0 Then
      lblEndHr_A.Caption = "MidNight"
    Else
      lblEndHr_A.Caption = t & " AM"
    End If
  ElseIf t = 12 Then
    lblEndHr_A.Caption = t & " PM"
  Else
    lblEndHr_A.Caption = t - 12 & " PM"
  End If

End Sub

Private Sub txtDisableEnd_A_GotFocus()
  SelAll txtDisableEnd_A
End Sub

Private Sub txtDisableEnd_A_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtDisableEnd_A.text) + 1
      txtDisableEnd_A.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtDisableEnd_A.text) - 1
      txtDisableEnd_A.text = Max(newval, 0)

    Case Else

      KeyAscii = KeyProcMax(txtDisableEnd_A, KeyAscii, False, 0, 2, 23)
  End Select

End Sub

Private Sub txtDisableEnd_A_LostFocus()
  mDevice.DisableEnd_A = Val(txtDisableEnd_A.text)
End Sub

Private Sub txtDisableEnd_B_Change()
  ClearAlerts
  Dim t             As Integer
  t = Val(txtDisableEnd_B.text)
  If t < 12 Then
    If t = 0 Then
      lblendHr_B.Caption = "MidNight"
    Else
      lblendHr_B.Caption = t & " AM"
    End If
  ElseIf t = 12 Then
    lblendHr_B.Caption = t & " PM"
  Else
    lblendHr_B.Caption = t - 12 & " PM"
  End If

End Sub

Private Sub txtDisableEnd_B_GotFocus()
  SelAll txtDisableEnd_B
End Sub

Private Sub txtDisableEnd_B_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtDisableEnd_B.text) + 1
      txtDisableEnd_B.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtDisableEnd_B.text) - 1
      txtDisableEnd_B.text = Max(newval, 0)

    Case Else

      KeyAscii = KeyProcMax(txtDisableEnd_B, KeyAscii, False, 0, 2, 23)
  End Select

End Sub

Private Sub txtDisableEnd_B_LostFocus()
  mDevice.DisableEnd_B = Val(txtDisableEnd_B.text)
End Sub

Private Sub txtDisableEnd_Change()
  ClearAlerts
  Dim t             As Integer
  t = Val(txtDisableEnd.text)
  If t < 12 Then
    If t = 0 Then
      lblEndHr.Caption = "MidNight"
    Else
      lblEndHr.Caption = t & " AM"
    End If
  ElseIf t = 12 Then
    lblEndHr.Caption = t & " PM"
  Else
    lblEndHr.Caption = t - 12 & " PM"
  End If

End Sub

Private Sub txtDisableEnd_GotFocus()
  SelAll txtDisableEnd
End Sub

Private Sub txtDisableEnd_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtDisableEnd.text) + 1
      txtDisableEnd.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtDisableEnd.text) - 1
      txtDisableEnd.text = Max(newval, 0)

    Case Else

      KeyAscii = KeyProcMax(txtDisableEnd, KeyAscii, False, 0, 2, 23)
  End Select
End Sub

Private Sub txtDisableEnd_LostFocus()
  'If Input2 Then
  '  mDevice.DisableEnd_A = Val(txtDisableEnd.Text)
  'Else
  mDevice.DisableEnd = Val(txtDisableEnd.text)
  'End If

End Sub

Private Sub txtDisableStart_A_Change()

  Dim t             As Integer
  t = Val(txtDisableStart_A.text)
  If t < 12 Then
    If t = 0 Then
      lblStartHr_A.Caption = "MidNight"
    Else
      lblStartHr_A.Caption = t & " AM"
    End If
  ElseIf t = 12 Then
    lblStartHr_A.Caption = t & " PM"
  Else
    lblStartHr_A.Caption = t - 12 & " PM"
  End If

End Sub

Private Sub txtDisableStart_A_GotFocus()
  ClearAlerts
  SelAll txtDisableStart_A

End Sub

Private Sub txtDisableStart_A_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtDisableStart_A.text) + 1
      txtDisableStart_A.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtDisableStart_A.text) - 1
      txtDisableStart_A.text = Max(newval, 0)

    Case Else
      KeyAscii = KeyProcMax(txtDisableStart_A, KeyAscii, False, 0, 2, 23)
  End Select

End Sub

Private Sub txtDisableStart_A_LostFocus()
  mDevice.DisableStart_A = Val(txtDisableStart_A.text)
End Sub

Private Sub txtDisableStart_B_Change()
  Dim t             As Integer
  t = Val(txtDisableStart_B.text)
  If t < 12 Then
    If t = 0 Then
      lblStartHr_B.Caption = "MidNight"
    Else
      lblStartHr_B.Caption = t & " AM"
    End If
  ElseIf t = 12 Then
    lblStartHr_B.Caption = t & " PM"
  Else
    lblStartHr_B.Caption = t - 12 & " PM"
  End If

End Sub

Private Sub txtDisableStart_B_GotFocus()
  ClearAlerts
  SelAll txtDisableStart_B
End Sub

Private Sub txtDisableStart_B_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtDisableStart_B.text) + 1
      txtDisableStart_B.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtDisableStart_B.text) - 1
      txtDisableStart_B.text = Max(newval, 0)

    Case Else
      KeyAscii = KeyProcMax(txtDisableStart_B, KeyAscii, False, 0, 2, 23)
  End Select

End Sub

Private Sub txtDisableStart_B_LostFocus()
  mDevice.DisableStart_B = Val(txtDisableStart_B.text)
End Sub

Private Sub txtDisableStart_Change()

  Dim t             As Integer
  t = Val(txtDisableStart.text)
  If t < 12 Then
    If t = 0 Then
      lblStartHr.Caption = "MidNight"
    Else
      lblStartHr.Caption = t & " AM"
    End If
  ElseIf t = 12 Then
    lblStartHr.Caption = t & " PM"
  Else
    lblStartHr.Caption = t - 12 & " PM"
  End If
End Sub

Private Sub txtDisableStart_GotFocus()
  ClearAlerts
  SelAll txtDisableStart
End Sub

Private Sub txtDisableStart_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  Select Case KeyAscii
    Case vbKeyAdd, 43
      KeyAscii = 0
      newval = Val(txtDisableStart.text) + 1
      txtDisableStart.text = Min(newval, 23)
    Case vbKeySubtract, 45
      KeyAscii = 0
      newval = Val(txtDisableStart.text) - 1
      txtDisableStart.text = Max(newval, 0)

    Case Else
      KeyAscii = KeyProcMax(txtDisableStart, KeyAscii, False, 0, 2, 23)
  End Select


End Sub

Private Sub txtDisableStart_LostFocus()



  mDevice.DisableStart = Val(txtDisableStart.text)


End Sub

Private Sub txtGG1_Ad_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG1_Ad, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG1D_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG1D, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG2_Ad_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG2_Ad, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG2d_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG2d, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG3_Ad_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG3_Ad, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG3d_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG3d, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG4_Ad_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG4_Ad, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG4d_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG4d, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG5_Ad_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG5_Ad, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG5d_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG5d, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG6_Ad_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG6_Ad, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtGG6d_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtGG6d, KeyAscii, False, 0, 3, 999)
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

Private Sub txtPause_A_GotFocus()
  ClearAlerts
  SelAll txtPause_A

End Sub

Private Sub txtPause_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtPause_A, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtPause_A_LostFocus()
  mDevice.Pause_A = Val(txtPause_A.text)
End Sub

Private Sub txtPause_GotFocus()
  ClearAlerts
  SelAll txtPause
End Sub

Private Sub txtPause_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtPause, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtPause_LostFocus()
  mDevice.Pause = Val(txtPause.text)
End Sub

Private Sub txtRepeats_A_GotFocus()
  ClearAlerts
  SelAll txtRepeats_A
End Sub

Private Sub txtRepeats_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtRepeats_A, KeyAscii, False, 0, 2, 10)
End Sub

Private Sub txtRepeats_A_LostFocus()
  mDevice.Repeats_A = Val(txtRepeats_A.text)
End Sub

Private Sub txtRepeats_GotFocus()
  ClearAlerts
  SelAll txtRepeats
End Sub

Private Sub txtRepeats_KeyPress(KeyAscii As Integer)

  KeyAscii = KeyProcMax(txtRepeats, KeyAscii, False, 0, 2, 10)
End Sub

Private Sub txtRepeats_LostFocus()
  mDevice.Repeats = Val(txtRepeats.text)

End Sub

Private Sub txtSearchBox_Change()
  tmrSearch.Enabled = True
End Sub

Private Sub txtSerial_Change()
  On Error Resume Next
  txtSerial.ToolTipText = "TX ID: " & Val("&h" & Right$(txtSerial.text, 6))
End Sub

Private Sub txtSerial_GotFocus()
  
  

  ClearAlerts
  SelAll txtSerial
  SerialEditByFactory

End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
  'KeyAscii = KeyProcAlpha(KeyAscii)
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

Private Sub txtSerial_LostFocus()
  mDevice.Serial = txtSerial.text
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
  Set mDevice = New cESDevice
End Sub

Public Property Get Device() As cESDevice

  Set Device = mDevice

End Property

Public Property Set Device(Value As cESDevice)

  Set mDevice = Value

End Property

Private Sub txtSerial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
 SerialEditByFactory

End Sub
