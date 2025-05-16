VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmAutoReportEdit 
   Caption         =   "Edit Auto Report"
   ClientHeight    =   15120
   ClientLeft      =   10065
   ClientTop       =   2115
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   15120
   ScaleWidth      =   9720
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   14505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
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
         TabIndex        =   55
         Top             =   1695
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
         TabIndex        =   56
         Top             =   2280
         Width           =   1175
      End
      Begin VB.Frame fraGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   35
         Top             =   270
         Width           =   7365
         Begin VB.CommandButton cmdTestReport 
            Caption         =   "Send Now"
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
            Left            =   4170
            TabIndex        =   41
            Top             =   810
            Width           =   1150
         End
         Begin VB.CheckBox chkDisabled 
            Caption         =   "Disabled"
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
            Left            =   4170
            TabIndex        =   37
            Top             =   360
            Width           =   1725
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
            TabIndex        =   40
            Top             =   1035
            Width           =   3795
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
            TabIndex        =   38
            Top             =   375
            Width           =   3795
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
            TabIndex        =   39
            Top             =   780
            Width           =   780
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
            TabIndex        =   36
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.Frame fraSendSchedule 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Rooms"
         Height          =   2640
         Left            =   60
         TabIndex        =   9
         Top             =   11685
         Width           =   8460
         Begin VB.Frame fraDays 
            BorderStyle     =   0  'None
            Height          =   2145
            Left            =   2790
            TabIndex        =   19
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
               TabIndex        =   21
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
               TabIndex        =   20
               Top             =   0
               Width           =   435
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
            TabIndex        =   11
            Top             =   450
            Width           =   2325
         End
         Begin VB.Frame fraTimes 
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   210
            TabIndex        =   12
            Top             =   870
            Width           =   2325
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
               TabIndex        =   17
               ToolTipText     =   "0 is Midnight"
               Top             =   360
               Width           =   585
            End
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
               TabIndex        =   14
               ToolTipText     =   "0 is Midnight"
               Top             =   30
               Width           =   585
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
               TabIndex        =   15
               Top             =   90
               Width           =   555
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
               TabIndex        =   18
               Top             =   405
               Width           =   555
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
               TabIndex        =   16
               Top             =   405
               Width           =   345
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
               TabIndex        =   13
               Top             =   90
               Width           =   420
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
            TabIndex        =   10
            Top             =   180
            Width           =   1020
         End
      End
      Begin VB.Frame fraCriteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   43
         Top             =   8895
         Width           =   8925
         Begin VB.ListBox lstEvents 
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
            Left            =   2280
            Style           =   1  'Checkbox
            TabIndex        =   52
            Top             =   330
            Width           =   2085
         End
         Begin VB.ListBox lstRooms 
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
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   51
            Top             =   330
            Width           =   2085
         End
         Begin VB.ComboBox cboSort 
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
            Left            =   4590
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   330
            Width           =   2205
         End
         Begin VB.Label lblNoEvents 
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
            Left            =   3660
            TabIndex        =   49
            Top             =   60
            Width           =   465
         End
         Begin VB.Label lblAllEvents 
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
            Left            =   3120
            TabIndex        =   48
            Top             =   60
            Width           =   225
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
            Left            =   1500
            TabIndex        =   46
            Top             =   60
            Width           =   465
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
            Left            =   960
            MouseIcon       =   "frmAutoReportEdit.frx":0000
            TabIndex        =   45
            Top             =   60
            Width           =   225
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sort"
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
            Left            =   4590
            TabIndex        =   50
            Top             =   60
            Width           =   360
         End
         Begin VB.Label lblRooms 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rooms"
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
            TabIndex        =   44
            Top             =   60
            Width           =   585
         End
         Begin VB.Label lblEvents 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Events"
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
            Left            =   2280
            TabIndex        =   47
            Top             =   60
            Width           =   600
         End
      End
      Begin VB.Frame fraEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   4
         Top             =   3315
         Width           =   7425
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
            MaxLength       =   255
            TabIndex        =   6
            Top             =   375
            Width           =   6900
         End
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
            TabIndex        =   8
            Top             =   1035
            Width           =   3795
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
            TabIndex        =   5
            Top             =   120
            Width           =   1035
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
            TabIndex        =   7
            Top             =   780
            Width           =   1080
         End
      End
      Begin VB.Frame fraFileFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   60
         TabIndex        =   23
         Top             =   6105
         Width           =   7425
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
            Picture         =   "frmAutoReportEdit.frx":0152
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2040
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
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
            TabIndex        =   32
            Top             =   2040
            Visible         =   0   'False
            Width           =   6285
         End
         Begin VB.Frame fraFileType 
            BorderStyle     =   0  'None
            Caption         =   "FileType"
            Height          =   1605
            Left            =   2940
            TabIndex        =   26
            Top             =   270
            Width           =   3945
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
               TabIndex        =   29
               Top             =   750
               Width           =   3585
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
               TabIndex        =   30
               Top             =   1110
               Width           =   3405
            End
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
               TabIndex        =   28
               Top             =   390
               Value           =   -1  'True
               Width           =   3285
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
               TabIndex        =   27
               Top             =   90
               Width           =   1005
            End
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
            TabIndex        =   25
            Top             =   330
            Width           =   2295
         End
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
            TabIndex        =   24
            Top             =   300
            Visible         =   0   'False
            Width           =   2295
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
            TabIndex        =   31
            Top             =   1800
            Visible         =   0   'False
            Width           =   540
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
         TabIndex        =   34
         Top             =   6180
         Visible         =   0   'False
         Width           =   675
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
         TabIndex        =   42
         Top             =   8700
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   54
         Top             =   12420
         Visible         =   0   'False
         Width           =   420
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
         TabIndex        =   3
         Top             =   11490
         Visible         =   0   'False
         Width           =   1305
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
         TabIndex        =   22
         Top             =   5910
         Visible         =   0   'False
         Width           =   2025
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
         TabIndex        =   2
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAutoReportEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mReportID As Long
Public Report  As cAutoReport

Sub SelectAllRooms()
  Dim j As Long
  For j = 0 To lstRooms.listcount - 1
    lstRooms.Selected(j) = True

  Next
  lstRooms.ListIndex = Min(0, lstRooms.listcount - 1)
End Sub

Sub DeSelectAllRooms()
  Dim j As Long
  For j = 0 To lstRooms.listcount - 1
    lstRooms.Selected(j) = False

  Next
  lstRooms.ListIndex = Min(0, lstRooms.listcount - 1)
End Sub
Sub DeSelectAllEvents()
  Dim j As Long
  For j = 0 To lstEvents.listcount - 1
    lstEvents.Selected(j) = False

  Next
  lstEvents.ListIndex = Min(0, lstEvents.listcount - 1)
End Sub
Sub SelectAllEvents()
  Dim j As Long
  For j = 0 To lstEvents.listcount - 1
    lstEvents.Selected(j) = True

  Next
  lstEvents.ListIndex = Min(0, lstEvents.listcount - 1)


End Sub
Function Save() As Boolean

  Dim j               As Integer
  Dim Bitfield        As Integer  ' for DOW
  Dim Room            As cRoom
  Dim SQL             As String
  Dim dataWrapper     As cDataWrapper
  Dim rs              As ADODB.Recordset
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


  ' send schedule

  Dim HasSecondShift As Boolean
  Dim HasThirdShift As Boolean

  If Configuration.EndFirst = Configuration.EndNight Then
    HasSecondShift = False
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight = Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = False
  ElseIf Configuration.EndFirst <> Configuration.EndNight And Configuration.EndNight <> Configuration.EndThird Then
    HasSecondShift = True
    HasThirdShift = True
  Else
    HasSecondShift = False
    HasThirdShift = False
  End If


  Report.DayPeriod = GetComboItemData(cboPeriod)
  
  Select Case Report.DayPeriod
    Case AUTOREPORT_SHIFT1
      If HasThirdShift Then
      Report.DayPartStart = Configuration.EndThird
      Report.DayPartEnd = Configuration.EndFirst
      
      Else
      Report.DayPartStart = Configuration.EndNight
      Report.DayPartEnd = Configuration.EndFirst
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
    
  End Select
  
  



  For j = 0 To lstDOW.listcount - 1
    If lstDOW.Selected(j) Then
      Bitfield = Bitfield Or 2 ^ j
    End If
  Next
  If Bitfield = 0 Then Bitfield = 1
  Report.DAYS = Bitfield

  ' criteria
  ' get just the rooms selected
  Set Report.Rooms = New Collection
  For j = 0 To lstRooms.listcount - 1
    If lstRooms.Selected(j) Then
      Set Room = New cRoom
      Room.Room = lstRooms.list(j)
      Room.RoomID = lstRooms.ItemData(j)
      Report.Rooms.Add Room
    End If
  Next

  ' get just the events selected
  Set Report.Events = New Collection

  For j = 0 To lstEvents.listcount - 1
    If lstEvents.Selected(j) Then
      Set dataWrapper = New cDataWrapper
      dataWrapper.LongValue = lstEvents.ItemData(j)
      Report.Events.Add dataWrapper
    End If
  Next

  ' and the sort order
  Report.SortOrder = GetComboItemData(cboSort)


  SQL = "SELECT * FROM autoreports WHERE ReportID = " & Report.ReportID
  Set rs = New ADODB.Recordset
  rs.Open SQL, conn, gCursorType, gLockType
  If rs.EOF Then
    rs.addnew
  End If
  Report.UpdateData rs
  rs.Update



  Report.ReportID = rs("reportid")
  mReportID = Report.ReportID
  rs.Close
  Set rs = Nothing

  LoadAutoReports
  
  'SQL = "INSERT Into MyTable(ColumnName) values(" & Value & ")" & vbCrLf & " SELECT @@IDENTITY"

End Function

Private Sub cboPeriod_Click()
  updatescreen
End Sub

Private Sub cmdTestReport_Click()
  ResetActivityTime
  Dim rc As Boolean
 ' Debug.Assert 0
  
  rc = Report.due()
  Report.DoReport
End Sub

'Private Sub cboTime_Click()
'  updatescreen
'End Sub

Private Sub Form_Initialize()
  Set Report = New cAutoReport
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
  SetControls
  FillLists
  ShowPanel TabStrip.SelectedItem.Key
  updatescreen
End Sub
Sub updatescreen()

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




End Sub
Sub FillLists()

  Dim j As Integer

  cboPeriod.Clear
  AddToCombo cboPeriod, "Daily", AUTOREPORT_DAILY
  AddToCombo cboPeriod, "1st Shift", AUTOREPORT_SHIFT1
  AddToCombo cboPeriod, "2nd Shift", AUTOREPORT_SHIFT2
  AddToCombo cboPeriod, "3nd Shift", AUTOREPORT_SHIFT3
  AddToCombo cboPeriod, "Weekly", AUTOREPORT_WEEKLY
  AddToCombo cboPeriod, "Monthly", AUTOREPORT_MONTHLY
  cboPeriod.ListIndex = 0

  lstDOW.Clear
  For j = 1 To 7
    AddToListBox lstDOW, Format(j, "dddd"), j - 1
  Next
  lstDOW.ListIndex = 0

  cboSort.Clear
  AddToCombo cboSort, "Room", AUTOREPORT_SORT_ROOM
  AddToCombo cboSort, "Elapsed Time (Desc)", AUTOREPORT_SORT_ELAPSED
  AddToCombo cboSort, "Chronological", AUTOREPORT_SORT_CHRONO
  cboSort.ListIndex = 0

  ' Rooms List
  RefreshRoomlist
  ' Events List
  RefreshEventList


  'cboTime.Clear
  '  For j = 1 To 24
  '    cboTime.AddItem Format(j, "00") & ":00" & IIf(j = 12, " (noon)", IIf(j = 24, " (midnight)", ""))
  '    cboTime.ItemData(cboTime.NewIndex) = j * 100
  '  Next
  '  cboTime.ListIndex = 0


  '  lstDOM.Clear
  '  For j = 1 To 28
  '    lstDOM.AddItem Format(j)
  '    lstDOM.ItemData(lstDOM.NewIndex) = j
  '  Next
  '  lstDOM.ListIndex = 0

End Sub
Sub RefreshSortList()

End Sub
Sub RefreshRoomlist()
  Dim SQL As String
  Dim rs As Recordset
  lstRooms.Clear

  AddToListBox lstRooms, "<Unassigned>", 0

  SQL = "Select room, roomid from rooms order by room"
  Set rs = ConnExecute(SQL)
  Do Until rs.EOF
    AddToListBox lstRooms, rs("room") & "", rs("roomid")
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing



End Sub
Sub RefreshEventList()
  lstEvents.Clear

  AddToListBox lstEvents, "Alarms", EVT_EMERGENCY
  AddToListBox lstEvents, "Alerts", EVT_ALERT
  AddToListBox lstEvents, "Low Battery", EVT_BATTERY_FAIL
  AddToListBox lstEvents, "Trouble", EVT_CHECKIN_FAIL
  AddToListBox lstEvents, "Tamper", EVT_TAMPER
  AddToListBox lstEvents, "External", EVT_EXTERN
  AddToListBox lstEvents, "Line Loss", EVT_LINELOSS
  'AddToCombo cboEventType, "Comm Error", EVT_COMM_TIMEOUT
  lstEvents.ListIndex = 0

End Sub

Sub SetControls()
  Dim f As Control

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

Public Sub Fill()
  Dim j             As Integer
  Dim Room          As cRoom
  Dim dataWrapper   As cDataWrapper
  Dim rs            As ADODB.Recordset
  Dim SQL           As String




  SQL = "SELECT * FROM AutoReports WHERE reportid = " & Report.ReportID

  Set rs = ConnExecute(SQL)

  If Not rs.EOF Then
    Report.Parse rs
    
  Else
    Set Report = New cAutoReport
  End If


  ListBoxClearSelections lstRooms
  ListBoxClearSelections lstEvents


  ' general
  chkDisabled.Value = IIf(Report.Disabled, 1, 0) '
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
  For j = lstRooms.listcount - 1 To 0 Step -1
    For Each Room In Report.Rooms
      If Room.RoomID = lstRooms.ItemData(j) Then
        lstRooms.Selected(j) = True
      End If
    Next
  Next

  For j = lstEvents.listcount - 1 To 0 Step -1
    For Each dataWrapper In Report.Events
      If dataWrapper.LongValue = lstEvents.ItemData(j) Then
        lstEvents.Selected(j) = True
      End If
    Next
  Next

  For j = cboSort.listcount - 1 To 1 Step -1
    If cboSort.ItemData(j) = Report.SortOrder Then
      Exit For
    End If
  Next
  cboSort.ListIndex = j


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

Private Sub Form_Terminate()
  Set Report = Nothing
End Sub

Private Sub lblAllEvents_Click()
  SelectAllEvents
End Sub

Private Sub lblAllRooms_Click()
  SelectAllRooms
End Sub


Private Sub lblNoEvents_Click()
  DeSelectAllEvents
End Sub

Private Sub lblNone_Click()
  DeSelectAllRooms

End Sub

Private Sub optDaily_Click()
  updatescreen
End Sub

Private Sub optMonthly_Click()
  updatescreen
End Sub

Private Sub TabStrip_Click()
  ShowPanel TabStrip.SelectedItem.Key
End Sub


Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  ResetActivityTime
  Save

End Sub

Private Sub txtAssurEnd_Change()
  lblEndHr.Caption = ConvertHourToAMPM(Val(txtAssurEnd.text))
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

Private Sub txtAssurStart_Change()
  lblStartHr.Caption = ConvertHourToAMPM(Val(txtAssurStart.text))
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

Public Property Get ReportID() As Long
  ReportID = Report.ReportID
End Property

Public Property Let ReportID(ByVal Value As Long)
  Set Report = New cAutoReport
  Report.ReportID = Value
End Property

