VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmExceptionReport 
   Caption         =   "Edit Exception Report"
   ClientHeight    =   14460
   ClientLeft      =   4680
   ClientTop       =   2475
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   14460
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   14505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
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
         TabIndex        =   41
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
         Left            =   7725
         TabIndex        =   39
         Top             =   1695
         Width           =   1175
      End
      Begin VB.Frame fraFileFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   48
         Top             =   6105
         Width           =   7425
         Begin VB.CheckBox chkSaveAsFile 
            Caption         =   "Save As File"
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
            Left            =   390
            TabIndex        =   17
            Top             =   330
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox chkSendAsEmail 
            Caption         =   "Send As Email"
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
            Left            =   390
            TabIndex        =   50
            Top             =   330
            Width           =   2295
         End
         Begin VB.Frame fraFileType 
            BorderStyle     =   0  'None
            Caption         =   "FileType"
            Height          =   1605
            Left            =   2940
            TabIndex        =   18
            Top             =   270
            Width           =   3945
            Begin VB.OptionButton optTabDelimited 
               Caption         =   "Tab Delimited Table"
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
               Left            =   120
               TabIndex        =   19
               Top             =   390
               Value           =   -1  'True
               Width           =   3285
            End
            Begin VB.OptionButton optHTML 
               Caption         =   "HTML Document"
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
               Left            =   120
               TabIndex        =   21
               Top             =   1110
               Width           =   3405
            End
            Begin VB.OptionButton optTabDelimitedNoHeader 
               Caption         =   "Tab Delimited Table / NO Headers"
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
               Left            =   120
               TabIndex        =   20
               Top             =   750
               Width           =   3585
            End
            Begin VB.Label lblFileFormat 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "File Format:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   90
               Width           =   1005
            End
         End
         Begin VB.TextBox txtFolder 
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
            Left            =   150
            MaxLength       =   250
            TabIndex        =   23
            Top             =   2040
            Visible         =   0   'False
            Width           =   6285
         End
         Begin VB.CommandButton cmdGetfolder 
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
            Left            =   6480
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmExceptionReport.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2040
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblFolderRemote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder"
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
            Left            =   150
            TabIndex        =   22
            Top             =   1800
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.Frame fraEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   47
         Top             =   3315
         Width           =   7425
         Begin VB.TextBox txtEmailSubject 
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
            Left            =   240
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1035
            Width           =   3795
         End
         Begin VB.TextBox txtEmailRecipient 
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
            Left            =   240
            LinkTimeout     =   150
            MaxLength       =   255
            TabIndex        =   14
            Top             =   375
            Width           =   6900
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Line"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   780
            Width           =   1080
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Recipient(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.Frame fraCriteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   46
         Top             =   8895
         Width           =   7425
         Begin VB.ListBox lstEventTypes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            Left            =   60
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   26
            Top             =   360
            Width           =   5325
         End
         Begin VB.Label lblRooms 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Device Type"
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
            TabIndex        =   25
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label lblAllRooms 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2460
            MouseIcon       =   "frmExceptionReport.frx":052A
            TabIndex        =   27
            Top             =   60
            Width           =   345
         End
         Begin VB.Label lblNone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3000
            TabIndex        =   28
            Top             =   60
            Width           =   585
         End
      End
      Begin VB.Frame fraSendSchedule 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Rooms"
         Height          =   2640
         Left            =   60
         TabIndex        =   43
         Top             =   11685
         Width           =   7440
         Begin VB.Frame fraTimes 
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   210
            TabIndex        =   45
            Top             =   1350
            Width           =   2325
            Begin VB.TextBox txtAssurStart 
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
               Height          =   300
               Left            =   510
               MaxLength       =   2
               TabIndex        =   32
               ToolTipText     =   "0 is Midnight"
               Top             =   30
               Width           =   585
            End
            Begin VB.TextBox txtAssurEnd 
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
               Height          =   300
               Left            =   510
               MaxLength       =   2
               TabIndex        =   35
               ToolTipText     =   "0 is Midnight"
               Top             =   360
               Width           =   585
            End
            Begin VB.Label lblAssurStart 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start"
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
               TabIndex        =   31
               Top             =   90
               Width           =   420
            End
            Begin VB.Label lblAssurStop 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "End"
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
               Left            =   105
               TabIndex        =   34
               Top             =   405
               Width           =   345
            End
            Begin VB.Label lblEndHr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "12 PM"
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
               Left            =   1140
               TabIndex        =   36
               Top             =   405
               Width           =   555
            End
            Begin VB.Label lblStartHr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "12 PM"
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
               Left            =   1140
               TabIndex        =   33
               Top             =   90
               Width           =   555
            End
         End
         Begin VB.ComboBox cboPeriod 
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
            TabIndex        =   30
            Top             =   450
            Width           =   2325
         End
         Begin VB.Frame fraDays 
            BorderStyle     =   0  'None
            Height          =   2145
            Left            =   2790
            TabIndex        =   37
            Top             =   150
            Width           =   2385
            Begin VB.ListBox lstDOW 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1635
               Left            =   30
               Style           =   1  'Checkbox
               TabIndex        =   38
               Top             =   270
               Width           =   2235
            End
            Begin VB.Label lblHeader 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
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
               TabIndex        =   44
               Top             =   0
               Width           =   435
            End
         End
         Begin VB.Label lblPeriod 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Period"
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
            TabIndex        =   29
            Top             =   180
            Width           =   1020
         End
      End
      Begin VB.Frame fraGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   40
         Top             =   270
         Width           =   7605
         Begin VB.TextBox txtTime 
            Alignment       =   1  'Right Justify
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
            Left            =   5820
            MaxLength       =   5
            TabIndex        =   60
            Text            =   "1"
            ToolTipText     =   "Time to Clear Alarm (seconds)"
            Top             =   780
            Width           =   615
         End
         Begin VB.OptionButton optReportType 
            Caption         =   "Exception"
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
            Index           =   0
            Left            =   4080
            TabIndex        =   59
            Top             =   720
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton optReportType 
            Caption         =   "Device Inventory"
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
            Index           =   2
            Left            =   4080
            TabIndex        =   58
            Top             =   1470
            Width           =   1935
         End
         Begin VB.CommandButton cmdSendNow 
            Caption         =   "Send  Now"
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
            Left            =   5130
            TabIndex        =   57
            Top             =   1980
            Width           =   1150
         End
         Begin VB.OptionButton optReportType 
            Caption         =   "Alarm Count"
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
            Index           =   1
            Left            =   4080
            TabIndex        =   6
            Top             =   1095
            Width           =   1725
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "View Now"
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
            Left            =   3030
            TabIndex        =   11
            Top             =   1980
            Width           =   1150
         End
         Begin VB.TextBox txtDateTo 
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
            Left            =   1650
            TabIndex        =   10
            Top             =   1980
            Width           =   1245
         End
         Begin VB.TextBox txtdateFrom 
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
            Left            =   360
            TabIndex        =   8
            Top             =   1980
            Width           =   1245
         End
         Begin VB.TextBox txtReportName 
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
            Left            =   240
            MaxLength       =   20
            TabIndex        =   3
            Top             =   375
            Width           =   3795
         End
         Begin VB.TextBox txtComment 
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
            Left            =   240
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1035
            Width           =   3795
         End
         Begin VB.CheckBox chkDisabled 
            Caption         =   "Manual/Disabled"
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
            Left            =   4110
            TabIndex        =   5
            Top             =   360
            Width           =   2145
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "> Minutes"
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
            Left            =   6480
            TabIndex        =   61
            Top             =   840
            Width           =   840
         End
         Begin VB.Label lblDateError 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Error"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000CC&
            Height          =   195
            Left            =   3015
            TabIndex        =   56
            Top             =   2040
            Width           =   420
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
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
            Left            =   1710
            TabIndex        =   9
            Top             =   1740
            Width           =   810
         End
         Begin VB.Label lblStart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
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
            Left            =   360
            TabIndex        =   7
            Top             =   1740
            Width           =   885
         End
         Begin VB.Label lblRPTName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Report Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   780
            Width           =   780
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2970
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   5239
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "General"
               Key             =   "main"
               Object.Tag             =   "main"
               Object.ToolTipText     =   "General Settings"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Criteria"
               Key             =   "criteria"
               Object.Tag             =   "criteria"
               Object.ToolTipText     =   "Report Criteria"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Send Schedule"
               Key             =   "sendschedule"
               Object.Tag             =   "sendschedule"
               Object.ToolTipText     =   "When to Send Report"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Output Options"
               Key             =   "outputoptions"
               Object.Tag             =   "outputoptions"
               Object.ToolTipText     =   "Output Options"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Email Settings"
               Key             =   "email"
               Object.ToolTipText     =   "Configure Email Settings"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Settings"
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
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Format and Options"
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
         TabIndex        =   55
         Top             =   5910
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send Schedule"
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
         TabIndex        =   54
         Top             =   11490
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   7500
         TabIndex        =   53
         Top             =   12420
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Criteria"
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
         TabIndex        =   52
         Top             =   8700
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General"
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
         Left            =   8340
         TabIndex        =   51
         Top             =   6180
         Visible         =   0   'False
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmExceptionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mReportID   As Long
Public Report       As cExceptionReport

Public Sub Fill()
'FillEvents
  FillLists


  Dim j             As Integer
  Dim Room          As cRoom
  Dim dataWrapper   As cDataWrapper
  Dim Rs            As ADODB.Recordset
  Dim SQL           As String

  
  

  SQL = "SELECT * FROM ExceptionReports WHERE reportid = " & ReportID

  Set Rs = ConnExecute(SQL)

  If Not Rs.EOF Then
    Report.Parse Rs

  Else
    Set Report = New cExceptionReport
  End If


  
  ListBoxClearSelections lstEventTypes




  ' general
  chkDisabled.Value = IIf(Report.Disabled, 1, 0)  '
  txtReportName.text = Report.ReportName
  txtComment.text = Report.Comment

  ' file format
  chkSaveAsFile.Value = 1  ' IIf(Report.SaveAsFile, 1, 0)
  chkSendAsEmail.Value = IIf(Report.SendAsEmail, 1, 0)
  txtFolder.text = Report.DestFolder

  Select Case Report.FileFormat
  Case AUTOREPORTFORMAT_TAB_NOHEADER
    optTabDelimitedNoHeader.Value = True
  Case AUTOREPORTFORMAT_HTML
    optHTML.Value = True
  Case Else  ' AUTOREPORTFORMAT_TAB
    optTabDelimited.Value = True
  End Select


  ' email tab
  txtEmailRecipient.text = Report.recipient
  txtEmailSubject.text = Report.Subject

  ' send schedule

  For j = cboPeriod.listcount - 1 To 1 Step -1
    If cboPeriod.ItemData(j) = Report.DayPeriod Then
      Exit For
    End If
  Next
  cboPeriod.ListIndex = j


  For j = 6 To 0 Step -1
    If 2 ^ j And Report.DAYS Then
      lstDOW.Selected(j) = True
    Else
      lstDOW.Selected(j) = False
    End If
  Next

  txtAssurStart.text = Report.DayPartStart
  txtAssurEnd.text = Report.DayPartEnd


  ' criteria

  For j = lstEventTypes.listcount - 1 To 0 Step -1
    lstEventTypes.Selected(j) = False
    For Each dataWrapper In Report.DevTypes

      If dataWrapper.LongValue = lstEventTypes.ItemData(j) Then
        If dataWrapper.LongValue <> 0 Then
          lstEventTypes.Selected(j) = True
        End If
      End If
    Next
  Next

  txtTime.text = Report.ResponseTime

  Select Case Report.ReportType
  Case RPT_INV
    optReportType(2).Value = True
  Case RPT_COUNT
    optReportType(1).Value = True
  Case Else  ' RPT_EXCEPTION
    optReportType(0).Value = True
  End Select

  updatescreen


End Sub
'Sub FillEvents()
'  Dim rs            As ADODB.Recordset
'  Dim SQl           As String
'  Dim Count         As Long
'  lstEventTypes.Clear
'
'  '  Set rs = ConnExecute("SELECT * FROM Devicetypes")
'  '  Do Until rs.EOF
'  '    AddToListBox lstEventTypes, rs("description") & "", rs("id")
'  '
'  '    rs.MoveNext
'  '  Loop
'  '  rs.Close
'
'  'AddToListBox lstEventType, "All", 0
'
'  'AddToListBox lstEventType, , 0
'  'AddToListBox lstEventType, "Low Battery", EVT_BATTERY_FAIL
'  'AddToListBox lstEventType, "Trouble", EVT_CHECKIN_FAIL
'  'AddToListBox lstEventType, "Tamper", EVT_TAMPER
'  'AddToCombo cboEventType, "Comm Error", EVT_COMM_TIMEOUT
'  'lstEventTypes.ListIndex = 0
'
'
'End Sub
Sub ResetForm()
'  lblDateError.Caption = ""
'  txtdateFrom.text = Format(Now, "mm/dd/yy")
'  txtDateTo.text = Format(Now, "mm/dd/yy")

End Sub


Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cboPeriod_Click()
  updatescreen
End Sub

Private Sub chkSaveAsFile_Click()
  updatescreen
End Sub

Private Sub chkSendAsEmail_Click()
    updatescreen
End Sub

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Save
End Sub
Function Save() As Boolean

  Dim j             As Integer
  Dim Bitfield      As Integer  ' for DOW
  Dim Room          As cRoom
  Dim SQL           As String
  Dim dataWrapper   As cDataWrapper
  Dim Rs            As ADODB.Recordset

  ResetActivityTime

  If mReportID = 0 Then

    'Set Report = New cExceptionReport
  End If

  ' general
  Report.Disabled = chkDisabled.Value
  Report.ReportName = Trim$(txtReportName.text)
  Report.Comment = Trim$(txtComment.text)

  ' file format
  Report.SaveAsFile = 1  ' IIf(chkSaveAsFile.Value <> 0, 1, 0)
  Report.SendAsEmail = IIf(chkSendAsEmail.Value <> 0, 1, 0)
  Select Case True
  Case optTabDelimitedNoHeader.Value
    Report.FileFormat = AUTOREPORTFORMAT_TAB_NOHEADER
  Case optHTML.Value
    Report.FileFormat = AUTOREPORTFORMAT_HTML
  Case Else  ' optTabDelimited.Value
    Report.FileFormat = AUTOREPORTFORMAT_TAB
  End Select
  Report.DestFolder = Trim$(txtFolder.text)

  ' email tab
  Report.recipient = Trim$(txtEmailRecipient.text)
  Report.Subject = Trim$(txtEmailSubject.text)

  ' criteria
  If Val(txtTime.text) < 1 Then
    txtTime.text = 1
  End If
  Report.ResponseTime = Val(txtTime.text)

  If optReportType(2).Value Then
    Report.ReportType = RPT_INV
  ElseIf optReportType(1).Value Then
    Report.ReportType = RPT_COUNT
  Else
    Report.ReportType = RPT_EXCEPTION
  End If


  ' get just the events selected
  Set Report.DevTypes = New Collection

'  If (Report.ReportType = RPT_INV) Then
'
'  For j = 0 To lstDeviceTypes.listcount - 1
'    If lstDeviceTypes.Selected(j) Then
'      Set dataWrapper = New cDataWrapper
'      dataWrapper.LongValue = lstDeviceTypes.ItemData(j)
'      Report.DevTypes.Add dataWrapper
'    End If
'  Next
'
'
'  Else
'
  For j = 0 To lstEventTypes.listcount - 1
    If lstEventTypes.Selected(j) Then
      Set dataWrapper = New cDataWrapper
      dataWrapper.LongValue = lstEventTypes.ItemData(j)
      Report.DevTypes.Add dataWrapper
    End If
  Next
'  End If

  Dim HasSecondShift As Boolean
  Dim HasThirdShift As Boolean

  If Configuration.EndFirst = Configuration.EndNight Then   ' no second or third shift' regardless of third shift ending
    HasSecondShift = False
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = True
  End If


  ' send schedule

  Report.DayPeriod = GetComboItemData(cboPeriod)

  Select Case Report.DayPeriod
  
  Case AUTOREPORT_SHIFT1
    If HasThirdShift Then
      Report.DayPartStart = Configuration.EndThird
      Report.DayPartEnd = Configuration.EndFirst
    ElseIf HasSecondShift Then
      Report.DayPartStart = Configuration.EndNight
      Report.DayPartEnd = Configuration.EndFirst
    Else
      Report.DayPartStart = 0
      Report.DayPartEnd = 24
    End If
  Case AUTOREPORT_SHIFT2
    Report.DayPartStart = Configuration.EndFirst
    Report.DayPartEnd = Configuration.EndNight
  Case AUTOREPORT_SHIFT3
    Report.DayPartStart = Configuration.EndNight
    Report.DayPartEnd = Configuration.EndThird
  Case AUTOREPORT_DAILY
    Report.DayPartStart = Val(txtAssurStart.text)  ' if equal, then cutoff time is the time
    Report.DayPartEnd = Val(txtAssurEnd.text)
  Case Else
      Report.DayPartStart = 0
      Report.DayPartEnd = 24
  End Select




  For j = 0 To lstDOW.listcount - 1
    If lstDOW.Selected(j) Then
      Bitfield = Bitfield Or 2 ^ j
    End If
  Next
  If Bitfield = 0 Then Bitfield = 1
  Report.DAYS = Bitfield



  SQL = "SELECT * FROM ExceptionReports WHERE ReportID = " & mReportID
  Set Rs = New ADODB.Recordset
  Rs.Open SQL, conn, gCursorType, gLockType
  If Rs.EOF Then
    Rs.addnew
  End If
  Report.UpdateData Rs  ' calls routine to update data in recordset
  Rs.Update



  Report.ReportID = Rs("reportid")
  mReportID = Report.ReportID
  Rs.Close
  Set Rs = Nothing

  Save = Err.Number = 0
  LoadAutoExReports



End Function

Private Sub cmdTestReport_Click()
  RunExceptionReportNow


End Sub


Sub RunExceptionReportNow()
  If Save() Then

  End If

End Sub


Private Sub cmdView_Click()
  ResetActivityTime
  If ValidateDates() Then
    ViewExceptionReport
  End If
End Sub

Sub ViewExceptionReport()
  Dim Criteria      As String
  Dim StartDate     As String
  Dim EndDate       As String

  If Save() Then
    If Report.ReportID <> 0 Then
      Criteria = CStr(Report.ReportID)
      StartDate = txtdateFrom.text
      EndDate = txtDateTo.text

      Call ExceptionReportView(Report.ReportType, Criteria, StartDate, EndDate)
    End If
  End If

End Sub


Private Sub cmdSendNow_Click()
 Dim rpt As cExceptionAutoReport
  ResetActivityTime
 If Save() Then
    If Report.ReportID <> 0 Then
      For Each rpt In gAutoExReports
        If rpt.ReportID = Report.ReportID Then
          rpt.NextReportDue = Now
        End If
      Next
    End If
  End If
End Sub

Private Sub Form_Initialize()
  Set Report = New cExceptionReport
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
  SetControls
  ' FillLists
  ShowPanel TabStrip.SelectedItem.Key
  updatescreen
  cmdSendNow.Enabled = MASTER
  Call ValidateDates
End Sub
Sub FillLists()
  Dim j             As Integer



  cboPeriod.Clear
  AddToCombo cboPeriod, "Daily", AUTOREPORT_DAILY
  AddToCombo cboPeriod, "1st Shift", AUTOREPORT_SHIFT1
  AddToCombo cboPeriod, "2nd Shift", AUTOREPORT_SHIFT2
  AddToCombo cboPeriod, "3rd Shift", AUTOREPORT_SHIFT3
  AddToCombo cboPeriod, "Weekly", AUTOREPORT_WEEKLY
  AddToCombo cboPeriod, "Monthly", AUTOREPORT_MONTHLY
  cboPeriod.ListIndex = 0



  lstDOW.Clear
  For j = 1 To 7
    AddToListBox lstDOW, Format(j, "dddd"), j - 1
  Next
  lstDOW.ListIndex = 0



  ' todo
  FillDevices

End Sub


'Sub FillDeviceTypes()
'
'  Dim d             As cDeviceType
'  Dim i             As Integer
'  Dim rs            As ADODB.Recordset
'  'Dim DeviceTypes   As Collection
'
'  Dim SQl           As String
'
'  Dim Model         As String
'
'
'  'Set DeviceTypes = New Collection
'
'  ' get devices actually used
'
'
'  lstDeviceTypes.Clear
'
'  SQl = "SELECT DISTINCT model FROM Devices ORDER BY model"  ' only models used
'  Set rs = ConnExecute(SQl)
'  Do While Not rs.EOF
'
'    Model = rs("Model") & ""
'
'    For i = 0 To MAX_ESDEVICETYPES
'      On Error Resume Next
'      If 0 = StrComp(ESDeviceType(i).Model, Model, vbTextCompare) Then
'        AddToListBox lstDeviceTypes, ESDeviceType(i).desc & " (" & ESDeviceType(i).Model & ")", ESDeviceType(i).CLSPTI
'        Exit For
'      End If
'    Next
'    rs.MoveNext
'  Loop
'
'  Set rs = Nothing
'
'
'
'
'
'  If lstDeviceTypes.listcount > 0 Then
'    lstDeviceTypes.ListIndex = 0
'  End If
'
'
'
'
'End Sub
Sub FillDevices()
  Dim d             As cDeviceType
  Dim i             As Integer
  Dim Rs            As ADODB.Recordset
  Dim DeviceTypes   As Collection

  Dim SQL           As String

  Dim Model         As String


  Set DeviceTypes = New Collection

  ' get devices actually used


  lstEventTypes.Clear

  SQL = "SELECT DISTINCT model FROM Devices ORDER BY model"  ' only models used
  Set Rs = ConnExecute(SQL)
  Do While Not Rs.EOF

    Model = Rs("Model") & ""

    For i = 0 To MAX_ESDEVICETYPES
      On Error Resume Next
      If 0 = StrComp(ESDeviceType(i).Model, Model, vbTextCompare) Then
        AddToListBox lstEventTypes, ESDeviceType(i).desc & " (" & ESDeviceType(i).Model & ")", ESDeviceType(i).CLSPTI
        Exit For
      End If
    Next



    Rs.MoveNext
  Loop

  Set Rs = Nothing





  If lstEventTypes.listcount > 0 Then
    lstEventTypes.ListIndex = 0
  End If


End Sub

Sub ShowPanel(ByVal Key As String)
  Select Case LCase(Key)
  Case "email"
    fraEmail.Visible = True
    fraCriteria.Visible = False
    fraSendSchedule.Visible = False
    fraGeneral.Visible = False
    fraFileFormat.Visible = False

  Case "sendschedule"
    fraSendSchedule.Visible = True
    fraCriteria.Visible = False
    fraEmail.Visible = False
    fraGeneral.Visible = False
    fraFileFormat.Visible = False

  Case "outputoptions"
    fraFileFormat.Visible = True
    fraCriteria.Visible = False
    fraEmail.Visible = False
    fraGeneral.Visible = False
    fraSendSchedule.Visible = False

  Case "criteria"

    fraCriteria.Visible = True
    fraEmail.Visible = False
    fraFileFormat.Visible = False
    fraSendSchedule.Visible = False
    fraGeneral.Visible = False


  Case Else  ' general
    fraGeneral.Visible = True
    fraCriteria.Visible = False
    fraEmail.Visible = False
    fraFileFormat.Visible = False
    fraSendSchedule.Visible = False
  End Select
End Sub

Sub SetControls()
  Dim f             As Control

  For Each f In Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor

    End If
  Next


  fraGeneral.left = TabStrip.ClientLeft
  fraGeneral.top = TabStrip.ClientTop
  fraGeneral.Height = TabStrip.ClientHeight
  fraGeneral.Width = TabStrip.ClientWidth

  fraCriteria.left = TabStrip.ClientLeft
  fraCriteria.top = TabStrip.ClientTop
  fraCriteria.Height = TabStrip.ClientHeight
  fraCriteria.Width = TabStrip.ClientWidth

  fraEmail.left = TabStrip.ClientLeft
  fraEmail.top = TabStrip.ClientTop
  fraEmail.Height = TabStrip.ClientHeight
  fraEmail.Width = TabStrip.ClientWidth

  fraSendSchedule.left = TabStrip.ClientLeft
  fraSendSchedule.top = TabStrip.ClientTop
  fraSendSchedule.Height = TabStrip.ClientHeight
  fraSendSchedule.Width = TabStrip.ClientWidth

  fraFileFormat.left = TabStrip.ClientLeft
  fraFileFormat.top = TabStrip.ClientTop
  fraFileFormat.Height = TabStrip.ClientHeight
  fraFileFormat.Width = TabStrip.ClientWidth

  txtdateFrom.text = Format(Now, "mm/dd/yy")
  txtDateTo.text = Format(Now, "mm/dd/yy")

End Sub
Sub updatescreen()

  If chkSendAsEmail.Value = 0 Then
    cmdSendNow.ToolTipText = "Save to " & App.Path & "\AutoReports\"
  Else
    cmdSendNow.ToolTipText = "Email to " & txtEmailRecipient.text
  End If
    

  Select Case GetComboItemData(cboPeriod)
    Case AUTOREPORT_DAILY

      fraTimes.Visible = True
      fraDays.Visible = True

    Case AUTOREPORT_SHIFT1, AUTOREPORT_SHIFT2, AUTOREPORT_SHIFT3
      fraTimes.Visible = False
      fraDays.Visible = True
    Case Else
      fraTimes.Visible = False
      fraDays.Visible = False
  End Select
  
  If Me.optReportType(2).Value Then
    fraTimes.Visible = False
  End If



End Sub
Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub

Private Sub lblAllRooms_Click()
  Dim j             As Long
  For j = 0 To lstEventTypes.listcount - 1
    lstEventTypes.Selected(j) = True
  Next
End Sub

Private Sub lblNone_Click()
  Dim j             As Long
  For j = 0 To lstEventTypes.listcount - 1
    lstEventTypes.Selected(j) = False
  Next
End Sub

Private Sub lstEventTypes_Click()
    ResetActivityTime
End Sub

Private Sub optReportType_Click(index As Integer)
  
    ResetActivityTime
  Select Case index
  
  Case 2
    lstEventTypes.Visible = True
    txtTime.Enabled = False
    lblRooms.Caption = "Device Types"
    txtdateFrom.Visible = False
    txtDateTo.Visible = False
    lblStart.Visible = False
    lblEnd.Visible = False
    fraTimes.Visible = False
    
  Case 1
    lstEventTypes.Visible = True
    txtTime.Enabled = False
    lblRooms.Caption = "Device Types"
    txtdateFrom.Visible = True
    txtDateTo.Visible = True
    lblStart.Visible = True
    lblEnd.Visible = True
    fraTimes.Visible = True
  
  
  Case Else  ' RPT_EXCEPTION
    lstEventTypes.Visible = True
    txtTime.Enabled = True
    lblRooms.Caption = "Device Types"
    txtdateFrom.Visible = True
    txtDateTo.Visible = True
    lblStart.Visible = True
    lblEnd.Visible = True
    fraTimes.Visible = True
    
  End Select
  
  updatescreen
  
End Sub

Private Sub TabStrip_Click()
  ShowPanel TabStrip.SelectedItem.Key
End Sub

Private Sub txtAssurEnd_Change()
  lblEndHr.Caption = ConvertHourToAMPM(Val(txtAssurEnd.text))
End Sub

Private Sub txtAssurEnd_GotFocus()
  SelAll txtAssurEnd
End Sub

Private Sub txtAssurEnd_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
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

Private Sub txtAssurStart_Change()
  lblStartHr.Caption = ConvertHourToAMPM(Val(txtAssurStart.text))
End Sub

Private Sub txtAssurStart_GotFocus()
  SelAll txtAssurStart
End Sub

Private Sub txtAssurStart_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
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

Private Sub txtComment_GotFocus()
  SelAll txtComment

End Sub

Private Sub txtdateFrom_Change()
  Dim s             As String
  s = Trim(txtdateFrom.text)
  If IsDate(s) Then
    lblStart.ForeColor = &H80000012
  Else
    lblStart.ForeColor = vbRed
  End If

End Sub

Private Sub txtdateFrom_GotFocus()
  SelAll txtdateFrom

End Sub

Private Sub txtdateFrom_KeyPress(KeyAscii As Integer)
  Dim newval        As Date
  Select Case KeyAscii
  Case vbKeyAdd, 43
    KeyAscii = 0
    If IsDate(txtdateFrom.text) Then
      newval = DateAdd("d", 1, txtdateFrom.text)
      txtdateFrom.text = Format(newval, "mm/dd/yy")
    End If
  Case vbKeySubtract, 45
    KeyAscii = 0
    If IsDate(txtdateFrom.text) Then
      newval = DateAdd("d", -1, txtdateFrom.text)
      txtdateFrom.text = Format(newval, "mm/dd/yy")
    End If

  Case Else
  End Select
End Sub

Private Sub txtdateFrom_LostFocus()
  ValidateDates
End Sub

Private Sub txtDateTo_Change()
  Dim s             As String
  s = Trim(txtDateTo.text)
  If IsDate(s) Then
    lblEnd.ForeColor = &H80000012
  Else
    lblEnd.ForeColor = vbRed
  End If
End Sub

Private Sub txtDateTo_GotFocus()
  SelAll txtDateTo

End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
  Dim newval        As Date
  Select Case KeyAscii
  Case vbKeyAdd, 43
    KeyAscii = 0
    If IsDate(txtDateTo.text) Then
      newval = DateAdd("d", 1, txtDateTo.text)
      txtDateTo.text = Format(newval, "mm/dd/yy")
    End If
  Case vbKeySubtract, 45
    KeyAscii = 0
    If IsDate(txtDateTo.text) Then
      newval = DateAdd("d", -1, txtDateTo.text)
      txtDateTo.text = Format(newval, "mm/dd/yy")
    End If

  Case Else
  End Select

End Sub

Private Sub txtDateTo_LostFocus()
  ValidateDates

End Sub
Function ValidateDates() As Boolean
  Dim s             As String

  lblDateError.Caption = ""
  s = Trim(txtdateFrom.text)
  If IsDate(s) Then
    txtdateFrom.text = Format(s, "mm/dd/yy")
    s = Trim(txtDateTo.text)
    If IsDate(s) Then
      txtDateTo.text = Format(s, "mm/dd/yy")
      ValidateDates = True
    Else
      lblDateError.Caption = "End Date Invalid"
    End If
  Else
    lblDateError.Caption = "Start Date Invalid"
  End If

End Function


Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  fraEnabler.BackColor = Me.BackColor
  SetParent fraEnabler.hwnd, hwnd
End Sub
Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

Public Property Get ReportID() As Long

  ReportID = mReportID

End Property

Public Property Let ReportID(ByVal ReportID As Long)

  mReportID = ReportID

End Property

Private Sub txtEmailRecipient_GotFocus()
  SelAll txtEmailRecipient
End Sub

Private Sub txtEmailSubject_GotFocus()
  SelAll txtEmailSubject
End Sub

Private Sub txtFolder_GotFocus()
  SelAll txtFolder
End Sub

Private Sub txtReportName_GotFocus()
  SelAll txtReportName
End Sub

Private Sub txtTime_GotFocus()
  SelAll txtTime
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
  Dim newval        As Integer
  '  Select Case KeyAscii
  '    Case vbKeyAdd, 43
  '      KeyAscii = 0
  '      newval = Val(txtTime.text)
  '      txtTime.text = Min(newval, 99999)
  '    Case vbKeySubtract, 45
  '      KeyAscii = 0
  '      newval = Val(txtDisableStart_A.text) - 1
  '      txtDisableStart_A.text = Max(newval, 0)
  '
  '    Case Else
  KeyAscii = KeyProcMax(txtTime, KeyAscii, False, 0, 5, 99999)
  '  End Select
End Sub

Private Sub txtTime_LostFocus()
  If Val(txtTime.text) < 1 Then
    txtTime.text = 1
  End If

End Sub
