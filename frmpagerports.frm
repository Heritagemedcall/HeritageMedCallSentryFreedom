VERSION 5.00
Begin VB.Form frmOutputServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output Server"
   ClientHeight    =   3315
   ClientLeft      =   720
   ClientTop       =   7425
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.TextBox txtSerialNumber 
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
         Height          =   345
         Left            =   195
         MaxLength       =   8
         TabIndex        =   82
         Top             =   1605
         Width           =   1350
      End
      Begin VB.Frame fraProto7 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   30
         TabIndex        =   7
         Top             =   3120
         Width           =   2955
         Begin VB.ComboBox cboMarquisCode 
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
            ItemData        =   "frmPagerPorts.frx":0000
            Left            =   1410
            List            =   "frmPagerPorts.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label lblMarquis 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message Type"
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
            Left            =   60
            TabIndex        =   57
            Top             =   75
            Width           =   1245
         End
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
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2280
         Width           =   510
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
         TabIndex        =   49
         Top             =   1785
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
         TabIndex        =   50
         Top             =   2370
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
         Left            =   180
         MaxLength       =   50
         TabIndex        =   2
         Top             =   435
         Width           =   2670
      End
      Begin VB.ComboBox cboProtocol 
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
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2340
      End
      Begin VB.Frame fraProto6 
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
         Height          =   2895
         Left            =   3000
         TabIndex        =   30
         Tag             =   "Dialogic"
         Top             =   90
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CheckBox chkKeepOnPaging 
            Caption         =   "All"
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
            Left            =   3765
            TabIndex        =   54
            Top             =   2475
            Width           =   600
         End
         Begin VB.ComboBox cboDivaLines 
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
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   450
            Width           =   3615
         End
         Begin VB.ComboBox cboAckDigit 
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
            Left            =   2745
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   2430
            Width           =   885
         End
         Begin VB.TextBox txtTimeout 
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
            Left            =   1185
            MaxLength       =   3
            TabIndex        =   51
            Top             =   2430
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox txtTag 
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
            Left            =   750
            MaxLength       =   70
            TabIndex        =   38
            Top             =   1200
            Width           =   3615
         End
         Begin VB.TextBox txtPhone 
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
            Left            =   750
            MaxLength       =   15
            TabIndex        =   36
            Top             =   840
            Width           =   1665
         End
         Begin VB.TextBox txtRedialDelay 
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
            Left            =   2910
            MaxLength       =   2
            TabIndex        =   48
            Top             =   2025
            Width           =   510
         End
         Begin VB.TextBox txtMsgRepeats 
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
            TabIndex        =   42
            Top             =   1620
            Width           =   510
         End
         Begin VB.TextBox txtRedials 
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
            Left            =   1185
            MaxLength       =   2
            TabIndex        =   46
            Top             =   2025
            Width           =   510
         End
         Begin VB.TextBox txtMsgSpacing 
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
            Left            =   3915
            MaxLength       =   2
            TabIndex        =   44
            Top             =   1620
            Width           =   510
         End
         Begin VB.TextBox txtMsgDelay 
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
            Left            =   1185
            MaxLength       =   2
            TabIndex        =   40
            Top             =   1620
            Width           =   510
         End
         Begin VB.ComboBox cboDevices 
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
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   450
            Width           =   3615
         End
         Begin VB.ComboBox cboVoices 
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
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   90
            Width           =   3615
         End
         Begin VB.Label lblAckDigit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACK Digit"
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
            Left            =   1830
            TabIndex        =   52
            Top             =   2490
            Width           =   825
         End
         Begin VB.Label z12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(default)"
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
            Left            =   2490
            TabIndex        =   56
            Top             =   900
            Width           =   720
         End
         Begin VB.Label z10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Timeout"
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
            Left            =   420
            TabIndex        =   55
            Top             =   2490
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tag"
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
            Left            =   345
            TabIndex        =   37
            Top             =   1268
            Width           =   345
         End
         Begin VB.Label z1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ph #"
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
            Left            =   270
            TabIndex        =   35
            Top             =   908
            Width           =   420
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Redial Delay"
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
            Left            =   1740
            TabIndex        =   47
            Top             =   2085
            Width           =   1095
         End
         Begin VB.Label z6 
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
            Left            =   1785
            TabIndex        =   41
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label z9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Redials"
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
            Left            =   465
            TabIndex        =   45
            Top             =   2085
            Width           =   645
         End
         Begin VB.Label z8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spacing"
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
            TabIndex        =   43
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label z7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Msg Delay"
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
            Left            =   210
            TabIndex        =   39
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label lblDev 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Device"
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
            Left            =   75
            TabIndex        =   33
            Top             =   510
            Width           =   615
         End
         Begin VB.Label lblv 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voice"
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
            Left            =   195
            TabIndex        =   31
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame fraProto4 
         BorderStyle     =   0  'None
         Height          =   2790
         Left            =   3090
         TabIndex        =   10
         Tag             =   "ON TRAK"
         Top             =   180
         Width           =   4470
         Begin VB.ComboBox cboRelay5 
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
            TabIndex        =   67
            Top             =   570
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay6 
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
            TabIndex        =   68
            Top             =   930
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay7 
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
            TabIndex        =   69
            Top             =   1260
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay8 
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
            TabIndex        =   70
            Top             =   1620
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay4 
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
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1620
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay3 
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
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1260
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay2 
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
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   930
            Width           =   1545
         End
         Begin VB.ComboBox cboRelay1 
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
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   570
            Width           =   1545
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Left            =   2670
            TabIndex        =   75
            Top             =   630
            Width           =   120
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Left            =   2670
            TabIndex        =   74
            Top             =   990
            Width           =   120
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6"
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
            Left            =   2670
            TabIndex        =   73
            Top             =   1350
            Width           =   120
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "7"
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
            Left            =   2670
            TabIndex        =   72
            Top             =   1710
            Width           =   120
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relay 3"
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
            Left            =   0
            TabIndex        =   62
            Top             =   1710
            Width           =   660
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relay 2"
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
            Left            =   0
            TabIndex        =   61
            Top             =   1350
            Width           =   660
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relay 1"
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
            Left            =   0
            TabIndex        =   60
            Top             =   990
            Width           =   660
         End
         Begin VB.Label lblp1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relay 0"
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
            Left            =   0
            TabIndex        =   59
            Top             =   630
            Width           =   660
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relays"
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
            Left            =   2250
            TabIndex        =   58
            Top             =   210
            Width           =   585
         End
      End
      Begin VB.Frame fraProto1 
         BorderStyle     =   0  'None
         Height          =   2850
         Left            =   3090
         TabIndex        =   16
         Tag             =   "Serial IO"
         Top             =   135
         Width           =   4095
         Begin VB.TextBox txtLF 
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
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   29
            Top             =   2340
            Width           =   510
         End
         Begin VB.ComboBox cboBaud 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   390
            Width           =   2340
         End
         Begin VB.ComboBox cboBits 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1170
            Width           =   2340
         End
         Begin VB.ComboBox cboStop 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1560
            Width           =   2340
         End
         Begin VB.ComboBox cboFlow 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1950
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox cboPort 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   2340
         End
         Begin VB.ComboBox cboParity 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   780
            Width           =   2340
         End
         Begin VB.Label lblLF 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LF"
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
            Left            =   1260
            TabIndex        =   76
            Top             =   2400
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Comm Port"
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
            Left            =   540
            TabIndex        =   18
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Data bits"
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
            Left            =   675
            TabIndex        =   23
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Stop bits"
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
            TabIndex        =   25
            Top             =   1650
            Width           =   765
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Flow Control"
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
            Left            =   390
            TabIndex        =   27
            Top             =   2025
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Bits per second"
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
            TabIndex        =   19
            Top             =   495
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Parity"
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
            Left            =   960
            TabIndex        =   21
            Top             =   885
            Width           =   495
         End
      End
      Begin VB.Frame fraProto2 
         BorderStyle     =   0  'None
         Height          =   2370
         Left            =   3090
         TabIndex        =   11
         Tag             =   "PA System TTS"
         Top             =   315
         Width           =   3930
         Begin VB.CheckBox chkRepeatFirst 
            Alignment       =   1  'Right Justify
            Caption         =   "Repeat First Announce"
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
            Left            =   75
            TabIndex        =   15
            Top             =   1140
            Width           =   2565
         End
         Begin VB.CheckBox chkKeyPA 
            Alignment       =   1  'Right Justify
            Caption         =   "Key PA System"
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
            Left            =   60
            TabIndex        =   14
            Top             =   675
            Width           =   2580
         End
         Begin VB.ComboBox cboSoundDevice 
            Enabled         =   0   'False
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
            Left            =   1455
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   180
            Width           =   2340
         End
         Begin VB.Label lblSoundDevice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Sound Device"
            Enabled         =   0   'False
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
            Left            =   75
            TabIndex        =   12
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.Frame fraproto3 
         BorderStyle     =   0  'None
         Height          =   2370
         Left            =   3060
         TabIndex        =   9
         Tag             =   "Email"
         Top             =   300
         Width           =   4275
      End
      Begin VB.Frame fraProto8 
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   3000
         TabIndex        =   77
         Top             =   120
         Width           =   2595
         Begin VB.TextBox txtHostPort 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            MaxLength       =   5
            TabIndex        =   79
            Top             =   660
            Width           =   750
         End
         Begin VB.TextBox txtIP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            MaxLength       =   15
            TabIndex        =   78
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port"
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
            Left            =   390
            TabIndex        =   81
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP"
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
            Left            =   570
            TabIndex        =   80
            Top             =   300
            Width           =   195
         End
      End
      Begin VB.Label lblPause 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pause (Sec.)"
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
         Left            =   195
         TabIndex        =   5
         Top             =   2295
         Width           =   1110
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
         Left            =   165
         TabIndex        =   1
         Top             =   135
         Width           =   975
      End
      Begin VB.Label lblProtocol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Protocol"
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
         TabIndex        =   3
         Top             =   795
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOutputServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ServerID As Long


'NO, per output device Configuration.KeepOnPaging = Val(ReadSetting("Configuration", "KeepOnPaging", "0")) And 1

' need a preamble for dialup monitoring
' current string length for tag = 70
' absolute max string length for tag = 100

' PCA           : no setup frame
' serial port   : proto1
' TTS w/ relay  : proto2
' email         : proto3

' voice dialer  : proto6



Private Sub cboProtocol_Click()

  txtPause.Visible = True
  lblPause.Visible = True
  txtSerialNumber.Visible = False

  Select Case GetComboItemData(cboProtocol)
  
    Case PROTOCOL_MOBILE
    

      fraProto1.Visible = False
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False
      txtPause.Visible = False
      lblPause.Visible = False
      cboDivaLines.Visible = False
      cboDevices.Visible = False
  
  
    Case PROTOCOL_TAP_IP
      fraProto8.Visible = True
      fraProto1.Visible = False
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      
      
  
    Case PROTOCOL_REMOTE
      txtPause.Visible = False
      lblPause.Visible = False
      cboDivaLines.Visible = False
      cboDevices.Visible = False
      fraProto1.Visible = False
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False
      
      
    
  
    Case PROTOCOL_DIALOGIC
      fraProto6.Visible = True

      cboDivaLines.Visible = True
      cboDevices.Visible = False
      fraProto1.Visible = False
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False
    
    Case PROTOCOL_DIALER
      fraProto6.Visible = True
      cboDevices.Visible = True
      cboDivaLines.Visible = False
      
      fraProto1.Visible = False
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False
      
    Case PROTOCOL_PCA
      fraProto2.Visible = False
      fraProto1.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False

    Case PROTOCOL_TAP, PROTOCOL_TAP2, PROTOCOL_COMP1, PROTOCOL_COMP2, PROTOCOL_CENTRAL
      fraProto1.Visible = True
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False
    Case PROTOCOL_SDACT2
    
      fraProto1.Visible = True
      txtSerialNumber.Visible = True
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False


    Case PROTOCOL_TTS
      fraProto2.Visible = True
      fraProto1.Visible = False
      fraProto4.Visible = False
      fraproto3.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False

    Case PROTOCOL_EMAIL
      fraproto3.Visible = True
      fraProto1.Visible = False
      fraProto2.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False

    Case PROTOCOL_ONTRAK
    
      fraProto4.Visible = True
      fraProto1.Visible = False
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False

    'Case PROTOCOL_WEB
    'Case PROTOCOL_MARQUIS
    
    Case Else  'PROTOCOL_NONE
      fraProto1.Visible = True
      fraProto2.Visible = False
      fraproto3.Visible = False
      fraProto4.Visible = False
      fraProto6.Visible = False
      fraProto7.Visible = False
      fraProto8.Visible = False
  End Select
End Sub

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If ValidateForm() Then
    If Save() Then
      PreviousForm
      Unload Me
    Else
      messagebox Me, "Save Error", App.Title, vbInformation
    End If
  End If

End Sub
Function ValidateForm() As Boolean

  If Len(Trim(txtDescription.text)) = 0 Then
    messagebox Me, "Please Fill in a Description for This Device", App.Title, vbInformation
    Exit Function
  End If
  If cboProtocol.ListIndex < 0 Then
    Exit Function
  End If

  If cboProtocol.ItemData(cboProtocol.ListIndex) = PROTOCOL_SDACT2 Then
    txtSerialNumber.text = UCase$(Right$("00000000" & txtSerialNumber.text, 8))
    If left(txtSerialNumber.text, 1) = "0" Then
      messagebox Me, "SerialNumber Cannot Begin with 0", App.Title, vbCritical
      Exit Function
    ElseIf left(txtSerialNumber.text, 1) = "D" Then
      messagebox Me, "SerialNumber Cannot Begin with D", App.Title, vbCritical
      Exit Function
    ElseIf left(txtSerialNumber.text, 2) = "B2" Then
      messagebox Me, "SerialNumber Cannot Begin with B2", App.Title, vbCritical
      Exit Function
    End If

  End If

  If cboProtocol.ItemData(cboProtocol.ListIndex) = PROTOCOL_DIALER Then
    If cboVoices.ListIndex = -1 Or cboDevices.ListIndex = -1 Then
      Exit Function
    End If
  End If

  ValidateForm = True
End Function
Function Save() As Boolean
  Dim rs As Recordset
  Dim protocol As Integer
  Dim Action As Integer

  protocol = GetComboItemData(cboProtocol)

  Save = True

  Set rs = New ADODB.Recordset
  Select Case protocol
    Case PROTOCOL_PCA
      rs.Open "SELECT * FROM pagerdevices WHERE protocolid = " & protocol, conn, gCursorType, gLockType
    Case PROTOCOL_DIALER
      rs.Open "SELECT * FROM pagerdevices WHERE protocolid = " & protocol, conn, gCursorType, gLockType
    Case PROTOCOL_SDACT2
        rs.Open "SELECT * FROM pagerdevices WHERE protocolid = " & protocol, conn, gCursorType, gLockType
    Case Else
      rs.Open "SELECT * FROM pagerdevices WHERE ID = " & ServerID, conn, gCursorType, gLockType
  End Select
  If rs.EOF Then
    rs.addnew
    Action = 1  ' add
  Else
    Action = 2  ' edit
  End If

  rs("Description") = Trim(txtDescription.text)
  rs("protocolid") = protocol
  rs("Pause") = Val(txtPause.text)
  rs("Twice") = 0
  
  
  
  ' new 10/16/06
  rs("DialerVoice") = ""
  rs("DialerModem") = 0
  rs("DialerPhone") = ""
  rs("DialerTag") = ""
  rs("DialerMsgDelay") = 0
  rs("DialerMsgRepeats") = 0
  rs("DialerMsgSpacing") = 0
  rs("DialerRedials") = 0
  rs("DialerRedialDelay") = 0
  rs("DialerAckDigit") = 0
  rs("MarquisCode") = Max(0, cboMarquisCode.ListIndex)
  'rs("DialerAckDigit") = cboAckDigit.ItemData(cboAckDigit.ListIndex)
  rs("Relay1") = 0
  rs("Relay2") = 0
  rs("Relay3") = 0
  rs("Relay4") = 0
  rs("Relay5") = 0
  rs("Relay6") = 0
  rs("Relay7") = 0
  rs("Relay8") = 0
  
  rs("LF") = Format(Val(txtLF.text), "0")
  
  'new   2017-08-07
   
  
  Select Case protocol
    
    Case PROTOCOL_SDACT2
    
      
    
      rs("KeepPaging") = chkKeepOnPaging.Value And 1
      rs("protocolid") = protocol
      rs("Description") = "SDACT2"
      
      rs("Port") = GetComboItemData(cboPort)
      rs("BaudRate") = GetComboItemData(cboBaud)
      rs("Parity") = GetParityString(GetComboItemData(cboParity))
      rs("Bits") = Val(cboBits.text)
      rs("Flow") = GetComboItemData(cboFlow)
      rs("Twice") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""
      rs("AudioDevice") = txtSerialNumber.text
      
      
'
'     'rs("Pause") = 0
'      Rs("Twice") = 0
'      ' new 10/16/06
'      Rs("DialerVoice") = cboVoices.text
'      Rs("DialerModem") = cboDevices.ItemData(cboDevices.ListIndex)
'      Rs("DialerPhone") = Trim(txtPhone.text)
'      Rs("DialerTag") = Trim(txtTag.text)
'      Rs("DialerMsgDelay") = Val(txtMsgDelay.text)
'      Rs("DialerMsgRepeats") = Val(txtMsgRepeats.text)
'      Rs("DialerMsgSpacing") = Val(txtMsgSpacing.text)
'      Rs("DialerRedials") = Val(txtRedials.text)
'      Rs("DialerRedialDelay") = Val(txtRedialDelay.text)
'      Rs("DialerAckDigit") = cboAckDigit.ItemData(cboAckDigit.ListIndex)
    
    
    
    
    Case PROTOCOL_DIALER
      rs("KeepPaging") = chkKeepOnPaging.Value And 1
      rs("protocolid") = protocol
      rs("Description") = "DIALER"
      rs("Port") = 0
      rs("BaudRate") = 0
      rs("Parity") = 0
      rs("Bits") = 0
      rs("Flow") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""
      'rs("Pause") = 0
      rs("Twice") = 0
      ' new 10/16/06
      rs("DialerVoice") = cboVoices.text
      rs("DialerModem") = cboDevices.ItemData(cboDevices.ListIndex)
      rs("DialerPhone") = Trim(txtPhone.text)
      rs("DialerTag") = Trim(txtTag.text)
      rs("DialerMsgDelay") = Val(txtMsgDelay.text)
      rs("DialerMsgRepeats") = Val(txtMsgRepeats.text)
      rs("DialerMsgSpacing") = Val(txtMsgSpacing.text)
      rs("DialerRedials") = Val(txtRedials.text)
      rs("DialerRedialDelay") = Val(txtRedialDelay.text)
      rs("DialerAckDigit") = cboAckDigit.ItemData(cboAckDigit.ListIndex)

    Case PROTOCOL_DIALOGIC
      rs("KeepPaging") = chkKeepOnPaging.Value And 1
      rs("protocolid") = protocol
      rs("Description") = "DIALOGIC"
      rs("Port") = 0
      rs("BaudRate") = 0
      rs("Parity") = 0
      rs("Bits") = 0
      rs("Flow") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""
      'rs("Pause") = 0
      rs("Twice") = 0
      ' new 10/16/06
      rs("DialerVoice") = cboVoices.text
      rs("DialerModem") = cboDivaLines.ItemData(cboDivaLines.ListIndex)
      rs("DialerPhone") = Trim(txtPhone.text)
      rs("DialerTag") = Trim(txtTag.text)
      rs("DialerMsgDelay") = Val(txtMsgDelay.text)
      rs("DialerMsgRepeats") = Val(txtMsgRepeats.text)
      rs("DialerMsgSpacing") = Val(txtMsgSpacing.text)
      rs("DialerRedials") = Val(txtRedials.text)
      rs("DialerRedialDelay") = Val(txtRedialDelay.text)
      rs("DialerAckDigit") = cboAckDigit.ItemData(cboAckDigit.ListIndex)


    Case PROTOCOL_PCA
      rs("Description") = "PCA"
      rs("Port") = 0
      rs("BaudRate") = 0
      rs("Parity") = 0
      rs("Bits") = 0
      rs("Flow") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""
      rs("Pause") = Val(txtPause.text)
      rs("Twice") = 0
    
    
    
    
    Case PROTOCOL_TTS
      rs("Port") = 0
      rs("BaudRate") = 0
      rs("Parity") = 0
      rs("Bits") = 0
      rs("Flow") = 0
      rs("KeyPA") = chkKeyPA.Value
      rs("Twice") = chkRepeatFirst.Value
      rs("AudioDevice") = cboSoundDevice.text
      
      
    Case PROTOCOL_ONTRAK
      rs("Description") = "RELAY"
      rs("Port") = 0
      rs("BaudRate") = 0
      rs("Parity") = 0
      rs("Bits") = 0
      rs("Flow") = 0
      rs("Twice") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""
      rs("Relay1") = Max(0, cboRelay1.ListIndex)
      rs("Relay2") = Max(0, cboRelay2.ListIndex)
      rs("Relay3") = Max(0, cboRelay3.ListIndex)
      rs("Relay4") = Max(0, cboRelay4.ListIndex)
      rs("Relay5") = Max(0, cboRelay5.ListIndex)
      rs("Relay6") = Max(0, cboRelay6.ListIndex)
      rs("Relay7") = Max(0, cboRelay7.ListIndex)
      rs("Relay8") = Max(0, cboRelay8.ListIndex)
    
    Case PROTOCOL_TAP_IP
      rs("DialerPhone") = Trim(Me.txtIP.text)
      rs("Port") = Max(1, Val(txtHostPort.text))
      rs("BaudRate") = 9600
      rs("Parity") = "N"
      rs("Bits") = 8
      rs("Flow") = 0
      rs("Twice") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""
    
    Case Else
      rs("Port") = GetComboItemData(cboPort)
      rs("BaudRate") = GetComboItemData(cboBaud)
      rs("Parity") = GetParityString(GetComboItemData(cboParity))
      rs("Bits") = Val(cboBits.text)
      rs("Flow") = GetComboItemData(cboFlow)
      rs("Twice") = 0
      rs("KeyPA") = 0
      rs("AudioDevice") = ""


  End Select
  
  
  rs("Deleted") = 0
  rs("IncludePhone") = 0
  rs("Pin") = ""
  rs.Update
  If ServerID = 0 Then
    rs.MoveLast
  End If
  ServerID = rs("ID")
  rs.Close
  Set rs = Nothing
  
  
  ' need to add to devices table
  
  
  
  
  
  SetPriorityChannels
  
  ChangePageDevice ServerID, Action


End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    
  
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select

End Sub

Private Sub Form_Load()
  
  
  SetFrames
  
  FillCombos
  
  ResetForm
  
  ResetActivityTime
  
End Sub
Private Sub SetFrames()
  fraEnabler.BackColor = Me.BackColor
  fraProto1.BackColor = Me.BackColor
  fraProto2.BackColor = Me.BackColor
  fraproto3.BackColor = Me.BackColor
  fraProto4.BackColor = Me.BackColor
  fraProto6.BackColor = Me.BackColor

End Sub

Sub FillCombos()

  Dim j As Integer

  cboProtocol.Clear
  AddToCombo cboProtocol, PROTOCOL_NONE_TEXT, PROTOCOL_NONE
  
  AddToCombo cboProtocol, PROTOCOL_TAP_TEXT, PROTOCOL_TAP
  AddToCombo cboProtocol, PROTOCOL_TAP2_TEXT, PROTOCOL_TAP2
  AddToCombo cboProtocol, PROTOCOL_TAP_IP_TEXT, PROTOCOL_TAP_IP
  AddToCombo cboProtocol, PROTOCOL_COMP1_TEXT, PROTOCOL_COMP1
  AddToCombo cboProtocol, PROTOCOL_COMP2_TEXT, PROTOCOL_COMP2
  AddToCombo cboProtocol, PROTOCOL_TTS_TEXT, PROTOCOL_TTS
  AddToCombo cboProtocol, PROTOCOL_EMAIL_TEXT, PROTOCOL_EMAIL
  AddToCombo cboProtocol, PROTOCOL_DIALER_TEXT, PROTOCOL_DIALER
  AddToCombo cboProtocol, PROTOCOL_CENTRAL_TEXT, PROTOCOL_CENTRAL
  AddToCombo cboProtocol, PROTOCOL_ONTRAK_TEXT, PROTOCOL_ONTRAK
  AddToCombo cboProtocol, PROTOCOL_DIALOGIC_TEXT, PROTOCOL_DIALOGIC
  AddToCombo cboProtocol, PROTOCOL_REMOTE_TEXT, PROTOCOL_REMOTE
  AddToCombo cboProtocol, PROTOCOL_MOBILE_TEXT, PROTOCOL_MOBILE
  AddToCombo cboProtocol, PROTOCOL_PCA_TEXT, PROTOCOL_PCA
  AddToCombo cboProtocol, PROTOCOL_SDACT2_TEXT, PROTOCOL_SDACT2
  cboProtocol.ListIndex = 0


  cboPort.Clear
  
  AddToCombo cboPort, "None", 0
  For j = 1 To 256
    AddToCombo cboPort, "COM " & j, j
  Next
  cboPort.ListIndex = 0

  cboBits.Clear
  For j = 4 To 8
    AddToCombo cboBits, j, j
  Next
  cboBits.ListIndex = cboBits.listcount - 1

  cboAckDigit.Clear
  AddToCombo cboAckDigit, "None", 0
  AddToCombo cboAckDigit, "0", TAPI_DTMF_0
  AddToCombo cboAckDigit, "1", TAPI_DTMF_1
  AddToCombo cboAckDigit, "2", TAPI_DTMF_2
  AddToCombo cboAckDigit, "3", TAPI_DTMF_3
  AddToCombo cboAckDigit, "4", TAPI_DTMF_4
  AddToCombo cboAckDigit, "5", TAPI_DTMF_5
  AddToCombo cboAckDigit, "6", TAPI_DTMF_6
  AddToCombo cboAckDigit, "7", TAPI_DTMF_7
  AddToCombo cboAckDigit, "8", TAPI_DTMF_8
  AddToCombo cboAckDigit, "9", TAPI_DTMF_9
  AddToCombo cboAckDigit, "*", TAPI_DTMF_STAR
  AddToCombo cboAckDigit, "#", TAPI_DTMF_POUND
  cboAckDigit.ListIndex = 0



  cboBaud.Clear
  'AddToCombo cboBaud, "75", 75
  'AddToCombo cboBaud, "110", 110
  'AddToCombo cboBaud, "150", 150
  AddToCombo cboBaud, "300", 300
  AddToCombo cboBaud, "600", 600
  AddToCombo cboBaud, "1200", 1200
  AddToCombo cboBaud, "2400", 2400
  AddToCombo cboBaud, "4800", 4800
  AddToCombo cboBaud, "7200", 7200
  AddToCombo cboBaud, "9600", 9600
  AddToCombo cboBaud, "14400", 14400
  AddToCombo cboBaud, "19200", 19200
  AddToCombo cboBaud, "38400", 38400
  AddToCombo cboBaud, "57600", 57600
  AddToCombo cboBaud, "115200", 115200
  AddToCombo cboBaud, "128000", 128000

  cboBaud.ListIndex = 6

  cboParity.Clear
  AddToCombo cboParity, "Even", 0
  AddToCombo cboParity, "Odd", 1
  AddToCombo cboParity, "None", 2
  AddToCombo cboParity, "Mark", 3
  AddToCombo cboParity, "Space", 4

  cboParity.ListIndex = 2

  cboStop.Clear
  AddToCombo cboStop, 1, 10
  AddToCombo cboStop, 1.5, 15
  AddToCombo cboStop, 2, 20
  cboStop.ListIndex = 0

  cboFlow.Clear
  AddToCombo cboFlow, "None", 0
  AddToCombo cboFlow, "Hardware", 1
  AddToCombo cboFlow, "Xon/Xoff", 2

  cboFlow.ListIndex = 0

  cboSoundDevice.Clear
  AddToCombo cboSoundDevice, "Default", 0
  cboSoundDevice.ListIndex = 0

  cboMarquisCode.ListIndex = 0

  
  ' for ontrak relay  08/09
  cboRelay1.Clear
  cboRelay2.Clear
  cboRelay3.Clear
  cboRelay4.Clear
  cboRelay5.Clear
  cboRelay6.Clear
  cboRelay7.Clear
  cboRelay8.Clear
  
  
  
  AddToCombo cboRelay1, "PA Mic", 0
  AddToCombo cboRelay1, "One Shot", 1
  AddToCombo cboRelay1, "Flashing", 2
  AddToCombo cboRelay1, "Latching", 3
  
  AddToCombo cboRelay2, "PA Mic", 0
  AddToCombo cboRelay2, "One Shot", 1
  AddToCombo cboRelay2, "Flashing", 2
  AddToCombo cboRelay2, "Latching", 3
  
  AddToCombo cboRelay3, "PA Mic", 0
  AddToCombo cboRelay3, "One Shot", 1
  AddToCombo cboRelay3, "Flashing", 2
  AddToCombo cboRelay3, "Latching", 3
  
  AddToCombo cboRelay4, "PA Mic", 0
  AddToCombo cboRelay4, "One Shot", 1
  AddToCombo cboRelay4, "Flashing", 2
  AddToCombo cboRelay4, "Latching", 3
  
  AddToCombo cboRelay5, "PA Mic", 0
  AddToCombo cboRelay5, "One Shot", 1
  AddToCombo cboRelay5, "Flashing", 2
  AddToCombo cboRelay5, "Latching", 3
  
  AddToCombo cboRelay6, "PA Mic", 0
  AddToCombo cboRelay6, "One Shot", 1
  AddToCombo cboRelay6, "Flashing", 2
  AddToCombo cboRelay6, "Latching", 3
  
  AddToCombo cboRelay7, "PA Mic", 0
  AddToCombo cboRelay7, "One Shot", 1
  AddToCombo cboRelay7, "Flashing", 2
  AddToCombo cboRelay7, "Latching", 3
  
  AddToCombo cboRelay8, "PA Mic", 0
  AddToCombo cboRelay8, "One Shot", 1
  AddToCombo cboRelay8, "Flashing", 2
  AddToCombo cboRelay8, "Latching", 3
  
  
  
  cboRelay1.ListIndex = 0
  cboRelay2.ListIndex = 0
  cboRelay3.ListIndex = 0
  cboRelay4.ListIndex = 0
  cboRelay5.ListIndex = 0
  cboRelay6.ListIndex = 0
  cboRelay7.ListIndex = 0
  cboRelay8.ListIndex = 0
  
  
  
  
  ' new 10/16/06
  ' DIALER
  
  SpecialLog "Call EnumerateLines"
  EnumerateLines
  SpecialLog "Call EnumerateDivaLines"
  EnumerateDivaLines
  SpecialLog "Call EnumerateVoices"
  EnumerateVoices
  SpecialLog "End of Fill Combos"

End Sub
Sub EnumerateVoices()
  Dim spvoice  As spvoice
  Dim token   As ISpeechObjectToken
  Dim Voices  As Collection
  
  cboVoices.Clear
  
  Set spvoice = New spvoice
  Set Voices = New Collection
  If Not spvoice Is Nothing Then
    For Each token In spvoice.GetVoices
      Voices.Add token
    Next
  End If
  Set spvoice = Nothing
  
  For Each token In Voices
    cboVoices.AddItem token.GetDescription()
    'cboVoices.ItemData(cboVoices.ListIndex) = token.ID
  Next
  If cboVoices.listcount > 0 Then
    cboVoices.ListIndex = 0
  End If
  
End Sub
Sub EnumerateDivaLines()
'  If Not DivaSytem Is Nothing Then
'    ' get number of lines 1 to 24
'
'
'  End If
  
  
  'Dim DIVALine  As cDIVALine
  ' maybe get number of channels available
  
  'For Each provider In TAPIProviders
  Dim j As Long
  For j = 1 To 24
    cboDivaLines.AddItem "Dialogic Line " & j
    cboDivaLines.ItemData(cboDivaLines.NewIndex) = j
  Next
  If cboDivaLines.listcount > 0 Then
    cboDivaLines.ListIndex = 0
  End If
  
  
  
End Sub

Sub EnumerateLines()

        Dim TapiLine  As CTAPILine
        Dim j         As Long
        Dim s         As String
        Dim provider  As cTAPIProvider

        Dim TAPIProviders As Collection
        
        
10      cboDevices.Clear
        
20      Set TAPIProviders = New Collection


30      Set TapiLine = New CTAPILine
        
          
        
40      If TapiLine.Create <> 0 Then
50        For j = 0 To TapiLine.numLines - 1
60          SpecialLog "EnumerateLines TapiLine.numLines J " & j & "  " & TapiLine.numLines
70          TapiLine.CurrentLineID = j
80          If TapiLine.NegotiatedAPIVersion > 0 Then
              SpecialLog "EnumerateLines TapiLine.NegotiatedAPIVersion > 0 " & j
90            If TapiLine.LineSupportsVoiceCalls Then
                SpecialLog "EnumerateLines LineSupportsVoiceCalls " & j
                ' change "modem" to Telephony ??)
100             If InStr(1, TapiLine.ProviderInfo, "Modem", vbTextCompare) > 0 Then ' Or left$(TapiLine.ProviderInfo, 7) = "SIPTAPI" Then
                'If InStr(1, TapiLine.ProviderInfo, "Telephony", vbTextCompare) > 0 Then
110               Set provider = New cTAPIProvider
120               provider.ProviderInfo = TapiLine.ProviderInfo
130               provider.LineName = TapiLine.LineName
140               provider.PermanentLineID = TapiLine.PermanentLineID
150               provider.ID = j
160               TAPIProviders.Add provider
170             End If
180           End If
190         End If
200       Next
210     End If
        SpecialLog "EnumerateLines Call TapiLine.Finalize "
220     TapiLine.Finalize "Enumerate Lines Ouput Servers"
230     Set TapiLine = Nothing

240     For Each provider In TAPIProviders
250       cboDevices.AddItem provider.LineName
260       cboDevices.ItemData(cboDevices.NewIndex) = provider.PermanentLineID
270     Next
280     If cboDevices.listcount > 0 Then
290       cboDevices.ListIndex = 0
300     End If
        
        
End Sub

Sub Fill()
  Dim rs                 As Recordset
  Dim protocol           As Integer
  Dim j                  As Integer

  ResetForm
  Set rs = ConnExecute("SELECT * FROM pagerdevices WHERE ID = " & ServerID)
  chkKeyPA.Value = 0

  If Not rs.EOF Then

    txtDescription = rs("description") & ""
    protocol = rs("protocolid")
    txtPause.text = Val(rs("Pause") & "")
    chkKeyPA.Value = IIf(rs("keypa") = 1, 1, 0)
    chkRepeatFirst.Value = IIf(rs("twice") = 1, 1, 0)
    cboProtocol.ListIndex = Max(0, CboGetIndexByItemData(cboProtocol, protocol))
    cboSoundDevice.ListIndex = 0
    cboMarquisCode.ListIndex = Max(0, Val(rs("MarquisCode" & "")))
    cboRelay1.ListIndex = Max(0, Val(rs("Relay1" & "")))
    cboRelay2.ListIndex = Max(0, Val(rs("Relay2" & "")))
    cboRelay3.ListIndex = Max(0, Val(rs("Relay3" & "")))
    cboRelay4.ListIndex = Max(0, Val(rs("Relay4" & "")))
    cboRelay5.ListIndex = Max(0, Val(rs("Relay5" & "")))
    cboRelay6.ListIndex = Max(0, Val(rs("Relay6" & "")))
    cboRelay7.ListIndex = Max(0, Val(rs("Relay7" & "")))
    cboRelay8.ListIndex = Max(0, Val(rs("Relay8" & "")))

    If protocol = PROTOCOL_TTS Then
      ' nothing to do
    ElseIf protocol = PROTOCOL_SDACT2 Then
      For j = cboProtocol.listcount - 1 To 0 Step -1
        If cboProtocol.ItemData(j) = Val(rs("protocolid") & "") Then
          Exit For
        End If
      Next
      cboProtocol.ListIndex = j
      txtSerialNumber.text = rs("audiodevice") & ""
      cboPort.ListIndex = Max(0, CboGetIndexByItemData(cboPort, rs("port")))
      cboBaud.ListIndex = Max(0, CboGetIndexByItemData(cboBaud, rs("baudrate")))
      cboParity.ListIndex = Max(0, CboGetIndexByItemData(cboParity, GetParityID(rs("parity"))))
      cboBits.ListIndex = Max(0, CboGetIndexByItemData(cboBits, rs("bits")))
      cboStop.ListIndex = Max(0, CboGetIndexByItemData(cboStop, GetComboByText(cboStop, rs("stopbits") & "")))
      cboFlow.ListIndex = Max(0, CboGetIndexByItemData(cboFlow, rs("flow")))
      txtLF.text = Format(Val(rs("LF") & ""), "0")



    ElseIf protocol = PROTOCOL_DIALER Then
      ' new 10/16/06
      For j = cboVoices.listcount - 1 To 0 Step -1
        If 0 = StrComp(cboVoices.list(j), rs("DialerVoice") & "", vbTextCompare) Then
          Exit For
        End If
      Next
      cboVoices.ListIndex = j

      For j = cboDevices.listcount - 1 To 0 Step -1
        If cboDevices.ItemData(j) = Val(rs("DialerModem") & "") Then
          Exit For
        End If
      Next
      cboDevices.ListIndex = j

      txtPhone.text = rs("DialerPhone") & ""
      txtTag.text = rs("DialerTag") & ""
      txtMsgDelay.text = Format(Val(rs("DialerMsgDelay") & ""), "0")
      txtMsgRepeats.text = Format(Val(rs("DialerMsgRepeats") & ""), "0")
      txtMsgSpacing.text = Format(Val(rs("DialerMsgSpacing") & ""), "0")
      txtRedials.text = Format(Val(rs("DialerRedials") & ""), "0")
      txtRedialDelay.text = Format(Val(rs("DialerRedialDelay") & ""), "0")
      cboAckDigit.ListIndex = Max(0, CboGetIndexByItemData(cboAckDigit, rs("DialerAckDigit")))

      chkKeepOnPaging.Value = Val(rs("KeepPaging") & "") And 1


    ElseIf protocol = PROTOCOL_DIALOGIC Then
      ' new 10/16/06
      For j = cboVoices.listcount - 1 To 0 Step -1
        If 0 = StrComp(cboVoices.list(j), rs("DialerVoice") & "", vbTextCompare) Then
          Exit For
        End If
      Next
      cboVoices.ListIndex = j



      For j = cboDivaLines.listcount - 1 To 0 Step -1
        If cboDivaLines.ItemData(j) = Val(rs("DialerModem") & "") Then
          Exit For
        End If
      Next
      cboDivaLines.ListIndex = j



      txtPhone.text = rs("DialerPhone") & ""
      txtTag.text = rs("DialerTag") & ""
      txtMsgDelay.text = Format(Val(rs("DialerMsgDelay") & ""), "0")
      txtMsgRepeats.text = Format(Val(rs("DialerMsgRepeats") & ""), "0")
      txtMsgSpacing.text = Format(Val(rs("DialerMsgSpacing") & ""), "0")
      txtRedials.text = Format(Val(rs("DialerRedials") & ""), "0")
      txtRedialDelay.text = Format(Val(rs("DialerRedialDelay") & ""), "0")
      cboAckDigit.ListIndex = Max(0, CboGetIndexByItemData(cboAckDigit, rs("DialerAckDigit")))

      chkKeepOnPaging.Value = Val(rs("KeepPaging") & "") And 1

    ElseIf protocol = PROTOCOL_TAP_IP Then
      txtHostPort.text = rs("port") & ""
      txtIP.text = rs("DialerPhone") & ""


    Else                       ' includes new Central monitor
      cboPort.ListIndex = Max(0, CboGetIndexByItemData(cboPort, rs("port")))
      cboBaud.ListIndex = Max(0, CboGetIndexByItemData(cboBaud, rs("baudrate")))
      cboParity.ListIndex = Max(0, CboGetIndexByItemData(cboParity, GetParityID(rs("parity"))))
      cboBits.ListIndex = Max(0, CboGetIndexByItemData(cboBits, rs("bits")))
      cboStop.ListIndex = Max(0, CboGetIndexByItemData(cboStop, GetComboByText(cboStop, rs("stopbits") & "")))
      cboFlow.ListIndex = Max(0, CboGetIndexByItemData(cboFlow, rs("flow")))
      txtLF.text = Format(Val(rs("LF") & ""), "0")


    End If
  End If

  txtSerialNumber.Visible = (cboProtocol.ItemData(cboProtocol.ListIndex) = PROTOCOL_SDACT2)

  rs.Close
  Set rs = Nothing


End Sub

Sub ResetForm()
  txtDescription.text = ""
  txtPause.text = "0"
  cboProtocol.ListIndex = 0
  cboPort.ListIndex = 0
  cboBaud.ListIndex = 6
  cboBits.ListIndex = cboBits.listcount - 1
  cboParity.ListIndex = 2
  cboStop.ListIndex = 0
  cboFlow.ListIndex = 0
  cboSoundDevice.ListIndex = 0
  cboAckDigit.ListIndex = 0
  chkKeyPA.Value = 0
  chkRepeatFirst.Value = 0
  cboMarquisCode.ListIndex = 0
  chkKeepOnPaging.Value = 0


' DIALER 10/15/06
  If cboVoices.listcount > 0 Then
    cboVoices.ListIndex = 0
  End If
  If cboDevices.listcount > 0 Then
    cboDevices.ListIndex = 0
  End If
  txtPhone.text = ""
  txtTag.text = ""
  txtMsgDelay.text = "0"
  txtMsgRepeats.text = "0"
  txtMsgSpacing.text = "0"
  txtRedials.text = "0"
  txtRedialDelay.text = "0"

  txtLF.text = "0"



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

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub

Private Sub txtDescription_GotFocus()
  SelAll txtDescription
End Sub

Private Sub txtHostPort_GotFocus()
  SelAll txtHostPort
End Sub

Private Sub txtHostPort_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtHostPort, KeyAscii, False, 0, 5, 65000)

End Sub

Private Sub txtIP_GotFocus()
  SelAll txtIP
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8 ' backspace
    Case 48 To 57
    Case 46
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtLF_GotFocus()
 SelAll txtLF
End Sub

Private Sub txtLF_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtLF, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtMsgDelay_GotFocus()
  SelAll txtMsgDelay
End Sub

Private Sub txtMsgDelay_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMsgDelay, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtMsgRepeats_GotFocus()
  SelAll txtMsgRepeats
End Sub

Private Sub txtMsgRepeats_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMsgRepeats, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtMsgSpacing_GotFocus()
  SelAll txtMsgSpacing
End Sub

Private Sub txtMsgSpacing_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtMsgSpacing, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtPause_GotFocus()
  SelAll txtPause
End Sub

Private Sub txtPause_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtPause, KeyAscii, False, 0, 2, 60)
End Sub

Private Sub txtPhone_GotFocus()
  SelAll txtPhone
End Sub

Private Sub txtRedialDelay_GotFocus()
  SelAll txtRedialDelay
End Sub

Private Sub txtRedialDelay_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtRedialDelay, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtRedials_GotFocus()
  SelAll txtRedials
End Sub

Private Sub txtRedials_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtRedials, KeyAscii, False, 0, 2, 99)
End Sub

Private Sub txtSerialNumber_Change()
  On Error Resume Next
  txtSerialNumber.ToolTipText = "TX ID: " & Val("&h" & Right$(txtSerialNumber.text, 8))
End Sub

Private Sub txtSerialNumber_GotFocus()
 
  SelAll txtSerialNumber
End Sub

Private Sub txtSerialNumber_KeyPress(KeyAscii As Integer)
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

Private Sub txtTag_GotFocus()
  SelAll txtTag
End Sub

Private Sub txtTimeout_GotFocus()
  SelAll txtTimeout
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtTimeout, KeyAscii, False, 0, 3, 999)
End Sub
