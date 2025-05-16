VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmOutputMask 
   Caption         =   "OutputMask"
   ClientHeight    =   3210
   ClientLeft      =   4770
   ClientTop       =   7950
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
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
         TabIndex        =   30
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
         TabIndex        =   31
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Frame fraOutput 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2475
         Left            =   45
         TabIndex        =   2
         Top             =   450
         Width           =   7455
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
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   810
            Width           =   1875
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
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   465
            Width           =   1875
         End
         Begin VB.TextBox txtEscalate 
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
            Left            =   4635
            MaxLength       =   4
            TabIndex        =   22
            ToolTipText     =   "Escalation Timer in Minutes"
            Top             =   1192
            Width           =   675
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
            Left            =   450
            TabIndex        =   27
            Top             =   2055
            Visible         =   0   'False
            Width           =   1995
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
            Left            =   2235
            MaxLength       =   2
            TabIndex        =   19
            Top             =   1185
            Width           =   510
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
            Left            =   2235
            MaxLength       =   3
            TabIndex        =   25
            Top             =   1545
            Width           =   510
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
            TabIndex        =   13
            Top             =   780
            Width           =   1875
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
            TabIndex        =   7
            Top             =   435
            Width           =   1875
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
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   435
            Width           =   1875
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
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   780
            Width           =   1875
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
            Left            =   3000
            TabIndex        =   26
            Top             =   1560
            Width           =   2205
         End
         Begin VB.TextBox txtBattEsc 
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
            Left            =   4635
            MaxLength       =   4
            TabIndex        =   23
            ToolTipText     =   "Escalation Timer in Minutes"
            Top             =   1192
            Width           =   675
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Third Shift"
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
            Left            =   5550
            TabIndex        =   5
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esc"
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
            TabIndex        =   16
            Top             =   855
            Width           =   330
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grp"
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
            TabIndex        =   10
            Top             =   510
            Width           =   315
         End
         Begin VB.Label lblBattEsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Escalation Timer"
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
            Left            =   3090
            TabIndex        =   20
            ToolTipText     =   "Escalation Timer in Minutes"
            Top             =   1245
            Width           =   1425
         End
         Begin VB.Label lblTroubleEsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Escalation Timer"
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
            Left            =   3090
            TabIndex        =   21
            ToolTipText     =   "Escalation Timer in Minutes"
            Top             =   1245
            Width           =   1425
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
            Left            =   360
            TabIndex        =   24
            Top             =   1605
            Width           =   1740
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
            Left            =   1380
            TabIndex        =   18
            Top             =   1245
            Width           =   720
         End
         Begin VB.Label lblog2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esc"
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
            TabIndex        =   12
            Top             =   825
            Width           =   330
         End
         Begin VB.Label lblOG1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grp"
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
            TabIndex        =   6
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grp"
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
            Left            =   2340
            TabIndex        =   8
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esc"
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
            Left            =   2340
            TabIndex        =   14
            Top             =   825
            Width           =   330
         End
         Begin VB.Label lblDayShift 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frist Shift"
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
            Left            =   945
            TabIndex        =   3
            Top             =   150
            Width           =   825
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Second Shift"
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
            TabIndex        =   4
            Top             =   150
            Width           =   1110
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   3015
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5318
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Trouble"
               Key             =   "trouble"
               Object.ToolTipText     =   "Trouble Window Settings"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Battery"
               Key             =   "battery"
               Object.ToolTipText     =   "Low Battery Settings"
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
         TabIndex        =   28
         Top             =   1395
         Visible         =   0   'False
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmOutputMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BatteryOutput   As cOutputMask
Private TroubleOutput   As cOutputMask
'Private TamperOutput    As cOutputMask



Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
 
End Sub

Private Sub cmdOK_Click()
  DoDataExchange
  If SaveAll() Then
    Fill
    'PreviousForm
    'Unload Me
  End If
  
End Sub
Function SaveAll() As Boolean
          Dim rs As Recordset
          Dim Count As Long
          Dim SQl As String
          Dim troubleesc As Long
          Dim batteryesc As Long
10        If InIDE Then
              'Stop
20        End If

30        On Error GoTo SaveAll_Error

40        troubleesc = Val(txtEscalate.text)
50        batteryesc = Val(txtBattEsc.text)
60        If batteryesc <= 0 Then batteryesc = 3
70        If troubleesc <= 0 Then troubleesc = 3

80        If 0 Then 'TroubleOutput.OG1 <= 0 Then
90            'SQl = "Delete FROM ScreenMasks WHERE Screen = " & SCREEN_TROUBLE
100       Else

110           Set rs = ConnExecute("SELECT COUNT(Screen) FROM ScreenMasks WHERE Screen = " & SCREEN_TROUBLE)
120           Count = rs(0)
130           rs.Close
140           If Count = 0 Then

150               SQl = "Insert Into ScreenMasks (Screen , OG1, OG2, OG3, NG1, NG2, NG3, Repeats, RepeatUntil, SendCancel, Pause, ScreenName," & _
                        "OG4, OG5, OG6, NG4, NG5, NG6, OG1D, OG2D, OG3D, OG4D, OG5D, OG6D, NG1D, NG2D, NG3D, NG4D, NG5D, NG6D, GG1, GG2, GG3, GG4, GG5, GG6, GG1D, GG2D, GG3D, GG4D, GG5D, GG6D, GG1_A, GG2_A, GG3_A, GG4_A, GG5_A, GG6_A, GG1_AD, GG2_AD, GG3_AD, GG4_AD, GG5_AD, GG6_AD)"

160               SQl = SQl & " Values ("

170               SQl = SQl & SCREEN_TROUBLE & "," & TroubleOutput.OG1 & "," & TroubleOutput.OG2 & "," & TroubleOutput.OG3 & "," & TroubleOutput.NG1 & "," & _
                        TroubleOutput.NG2 & "," & TroubleOutput.NG3 & "," & TroubleOutput.Repeats & "," & TroubleOutput.RepeatUntil & "," & TroubleOutput.SendCancel & "," & TroubleOutput.Pause & "," & "'Trouble'" & "," & _
                        TroubleOutput.OG4 & "," & TroubleOutput.OG5 & "," & TroubleOutput.OG6 & "," & TroubleOutput.NG4 & "," & TroubleOutput.NG5 & "," & TroubleOutput.NG6 & "," & _
                        TroubleOutput.OG1D & "," & TroubleOutput.OG2D & "," & TroubleOutput.OG3D & "," & TroubleOutput.OG4D & "," & TroubleOutput.OG5D & "," & TroubleOutput.OG6D & "," & _
                        TroubleOutput.NG1D & "," & TroubleOutput.NG2D & "," & TroubleOutput.NG3D & "," & TroubleOutput.NG4D & "," & TroubleOutput.NG5D & "," & TroubleOutput.NG6D & "," & _
                        TroubleOutput.GG1 & "," & TroubleOutput.GG2 & "," & TroubleOutput.GG3 & "," & TroubleOutput.GG4 & "," & TroubleOutput.GG5 & "," & TroubleOutput.GG6 & "," & _
                        TroubleOutput.GG1D & "," & TroubleOutput.GG2D & "," & TroubleOutput.GG3D & "," & TroubleOutput.GG4D & "," & TroubleOutput.GG5D & "," & TroubleOutput.GG6D & "," & _
                      0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & _
                      0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ")"
180           Else

190               SQl = "UPDATE ScreenMasks SET " & _
                      " OG1 = " & TroubleOutput.OG1 & _
                        ", OG2 = " & TroubleOutput.OG2 & _
                        ", OG3 = " & TroubleOutput.OG3 & _
                        ", OG4 = " & TroubleOutput.OG4 & _
                        ", OG5 = " & TroubleOutput.OG5 & _
                        ", OG6 = " & TroubleOutput.OG6 & _
                        ", NG1 = " & TroubleOutput.NG1 & _
                        ", NG2 = " & TroubleOutput.NG2 & _
                        ", NG3 = " & TroubleOutput.NG3 & _
                        ", NG4 = " & TroubleOutput.NG4 & _
                        ", NG5 = " & TroubleOutput.NG5 & _
                        ", NG6 = " & TroubleOutput.NG6 & _
                        ", OG1d = " & troubleesc & ", OG2d = " & 0 & ", OG3d = " & 0 & ", OG4d = " & 0 & ", OG5d = " & 0 & ", OG6d = " & 0 & _
                        ", NG1d = " & troubleesc & ", NG2d = " & 0 & ", NG3d = " & 0 & ", NG4d = " & 0 & ", NG5d = " & 0 & ", NG6d = " & 0 & _
                        ", GG1 =  " & TroubleOutput.GG1 & ", GG2 = " & TroubleOutput.GG2 & ", GG3 = 0, GG4 =  0 , GG5 =  0 , GG6 =  0" & _
                        ", GG1d = 0, GG2d =  0 , GG3d =  0 , GG4d =  0 , GG5d =  0 , GG6d =  0" & _
                        ", GG1_a = 0, GG2_a = 0, GG3_a = 0 , GG4_a = 0 , GG5_a = 0 , GG6_a = 0" & _
                        ", GG1_ad = 0, GG2_ad = 0 , GG3_ad = 0, GG4_ad = 0, GG5_ad = 0, GG6_ad = 0" & _
                        ", Repeats = " & TroubleOutput.Repeats & _
                        ", RepeatUntil = " & TroubleOutput.RepeatUntil & _
                        ", Pause = " & TroubleOutput.Pause & _
                        ", SendCancel = " & TroubleOutput.SendCancel & _
                        ", ScreenName = 'Trouble' WHERE Screen = " & SCREEN_TROUBLE

200           End If
210       End If
220       ConnExecute SQl

230       If 0 Then 'BatteryOutput.OG1 <= 0 Then
240           'SQl = "Delete  FROM ScreenMasks WHERE Screen = " & SCREEN_BATTERY
250       Else

260           Set rs = ConnExecute("SELECT COUNT(Screen) FROM ScreenMasks WHERE Screen = " & SCREEN_BATTERY)
270           Count = rs(0)
280           rs.Close
290           If Count = 0 Then

300               SQl = "Insert Into ScreenMasks (Screen , OG1, OG2, OG3, NG1, NG2, NG3, Repeats, RepeatUntil, SendCancel, Pause, ScreenName, OG4, OG5, OG6, NG4, NG5, NG6, OG1D, OG2D, OG3D, OG4D, OG5D, OG6D, NG1D, NG2D, NG3D, NG4D, NG5D, NG6D, GG1, GG2, GG3, GG4, GG5, GG6, GG1D, GG2D, GG3D, GG4D, GG5D, GG6D, GG1_A, GG2_A, GG3_A, GG4_A, GG5_A, GG6_A, GG1_AD, GG2_AD, GG3_AD, GG4_AD, GG5_AD, GG6_AD)"

310               SQl = SQl & " Values ("
320               SQl = SQl & SCREEN_BATTERY & "," & BatteryOutput.OG1 & "," & BatteryOutput.OG2 & "," & BatteryOutput.OG3 & "," & BatteryOutput.NG1 & "," & _
                        BatteryOutput.NG2 & "," & BatteryOutput.NG3 & "," & BatteryOutput.Repeats & "," & BatteryOutput.RepeatUntil & "," & BatteryOutput.SendCancel & "," & BatteryOutput.Pause & "," & "'Battery'" & "," & _
                        BatteryOutput.OG4 & "," & BatteryOutput.OG5 & "," & BatteryOutput.OG6 & "," & BatteryOutput.NG4 & "," & BatteryOutput.NG5 & "," & BatteryOutput.NG6 & "," & _
                        BatteryOutput.OG1D & "," & BatteryOutput.OG2D & "," & BatteryOutput.OG3D & "," & BatteryOutput.OG4D & "," & BatteryOutput.OG5D & "," & BatteryOutput.OG6D & "," & _
                        BatteryOutput.NG1D & "," & BatteryOutput.NG2D & "," & BatteryOutput.NG3D & "," & BatteryOutput.NG4D & "," & BatteryOutput.NG5D & "," & BatteryOutput.NG6D & "," & _
                        BatteryOutput.GG1 & "," & BatteryOutput.GG2 & "," & BatteryOutput.GG3 & "," & BatteryOutput.GG4 & "," & BatteryOutput.GG5 & "," & BatteryOutput.GG6 & "," & _
                        BatteryOutput.GG1D & "," & BatteryOutput.GG2D & "," & BatteryOutput.GG3D & "," & BatteryOutput.GG4D & "," & BatteryOutput.GG5D & "," & BatteryOutput.GG6D & "," & _
                      0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & _
                      0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ")"


330           Else

340               SQl = "UPDATE ScreenMasks SET " & _
                      " OG1 = " & BatteryOutput.OG1 & _
                        ", OG2 = " & BatteryOutput.OG2 & _
                        ", OG3 = " & BatteryOutput.OG3 & _
                        ", OG4 = " & BatteryOutput.OG4 & _
                        ", OG5 = " & BatteryOutput.OG5 & _
                        ", OG6 = " & BatteryOutput.OG6 & _
                        ", NG1 = " & BatteryOutput.NG1 & _
                        ", NG2 = " & BatteryOutput.NG2 & _
                        ", NG3 = " & BatteryOutput.NG3 & _
                        ", NG4 = " & BatteryOutput.NG4 & _
                        ", NG5 = " & BatteryOutput.NG5 & _
                        ", NG6 = " & BatteryOutput.NG6 & _
                        ", OG1d = " & batteryesc & ", OG2d = " & 0 & ", OG3d = " & 0 & ", OG4d = " & 0 & ", OG5d = " & 0 & ", OG6d = " & 0 & _
                        ", NG1d = " & batteryesc & ", NG2d = " & 0 & ", NG3d = " & 0 & ", NG4d = " & 0 & ", NG5d = " & 0 & ", NG6d = " & 0 & _
                        ", GG1 =  " & BatteryOutput.GG1 & ", GG2 = " & BatteryOutput.GG2 & ", GG3 = 0, GG4 =  0 , GG5 =  0 , GG6 =  0" & _
                        ", GG1d = 0, GG2d =  0 , GG3d =  0 , GG4d =  0 , GG5d =  0 , GG6d =  0" & _
                        ", GG1_a = 0, GG2_a = 0, GG3_a = 0 , GG4_a = 0 , GG5_a = 0 , GG6_a = 0" & _
                        ", GG1_ad = 0, GG2_ad = 0 , GG3_ad = 0, GG4_ad = 0, GG5_ad = 0, GG6_ad = 0" & _
                        ", Repeats = " & BatteryOutput.Repeats & _
                        ", RepeatUntil = " & BatteryOutput.RepeatUntil & _
                        ", Pause = " & BatteryOutput.Pause & _
                        ", SendCancel = " & BatteryOutput.SendCancel & _
                        ", ScreenName = 'Battery' WHERE Screen = " & SCREEN_BATTERY
350           End If
360       End If
370       ConnExecute SQl

375       SaveAll = True

SaveAll_Resume:
380       Set rs = Nothing
390       On Error GoTo 0
400       Exit Function


SaveAll_Error:
410       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmOutputMask.SaveAll." & Erl
420       Resume SaveAll_Resume

End Function

Private Sub Form_Load()
  
  SetControls
  LoadSettings
  LoadOutputs
  ResetActivityTime
End Sub
Sub SetControls()
  fraEnabler.BackColor = Me.BackColor
  
  fraOutput.left = TabStrip.ClientLeft
  fraOutput.top = TabStrip.ClientTop
  fraOutput.Height = TabStrip.ClientHeight
  fraOutput.BackColor = Me.BackColor


End Sub
Sub LoadSettings()
   On Error GoTo LoadSettings_Error

  Set TroubleOutput = New cOutputMask
  'Set TamperOutput = New cOutputMask
  Set BatteryOutput = New cOutputMask
  
  
  
  Dim rs As Recordset
  Set rs = ConnExecute("SELECT * FROM ScreenMasks")
  Do Until rs.EOF
    Select Case rs("Screen")
      Case SCREEN_TROUBLE
        TroubleOutput.OG1 = Val("" & rs("og1"))
        TroubleOutput.OG2 = Val("" & rs("og2"))
        TroubleOutput.OG3 = Val("" & rs("og3"))
        TroubleOutput.OG4 = Val("" & rs("og4"))
        TroubleOutput.OG5 = Val("" & rs("og5"))
        TroubleOutput.OG6 = Val("" & rs("og6"))
  
        TroubleOutput.NG1 = Val("" & rs("ng1"))
        TroubleOutput.NG2 = Val("" & rs("ng2"))
        TroubleOutput.NG3 = Val("" & rs("ng3"))
        TroubleOutput.NG4 = Val("" & rs("ng4"))
        TroubleOutput.NG5 = Val("" & rs("ng5"))
        TroubleOutput.NG6 = Val("" & rs("ng6"))


        TroubleOutput.GG1 = Val("" & rs("gg1"))
        TroubleOutput.GG2 = Val("" & rs("gg2"))
        TroubleOutput.GG3 = Val("" & rs("gg3"))
        TroubleOutput.GG4 = Val("" & rs("gg4"))
        TroubleOutput.GG5 = Val("" & rs("gg5"))
        TroubleOutput.GG6 = Val("" & rs("gg6"))


        TroubleOutput.OG1D = Val("" & rs("og1d"))
        TroubleOutput.NG1D = Val("" & rs("ng1d"))
        TroubleOutput.GG1D = Val("" & rs("gg1d"))
  
        TroubleOutput.Repeats = Val("" & rs("Repeats"))
        TroubleOutput.RepeatUntil = Val("" & rs("RepeatUntil"))
        TroubleOutput.Pause = Val("" & rs("Pause"))
        TroubleOutput.SendCancel = Val("" & rs("SendCancel"))
        TroubleOutput.ScreenName = "" & rs("ScreenName")
  

      Case SCREEN_BATTERY
        BatteryOutput.OG1 = Val("" & rs("og1"))
        BatteryOutput.OG2 = Val("" & rs("og2"))
        BatteryOutput.OG3 = Val("" & rs("og3"))
        BatteryOutput.OG4 = Val("" & rs("og4"))
        BatteryOutput.OG5 = Val("" & rs("og5"))
        BatteryOutput.OG6 = Val("" & rs("og6"))
  
        BatteryOutput.NG1 = Val("" & rs("ng1"))
        BatteryOutput.NG2 = Val("" & rs("ng2"))
        BatteryOutput.NG3 = Val("" & rs("ng3"))
        BatteryOutput.NG4 = Val("" & rs("ng4"))
        BatteryOutput.NG5 = Val("" & rs("ng5"))
        BatteryOutput.NG6 = Val("" & rs("ng6"))
  
        BatteryOutput.GG1 = Val("" & rs("gg1"))
        BatteryOutput.GG2 = Val("" & rs("gg2"))
        BatteryOutput.GG3 = Val("" & rs("gg3"))
        BatteryOutput.GG4 = Val("" & rs("gg4"))
        BatteryOutput.GG5 = Val("" & rs("gg5"))
        BatteryOutput.GG6 = Val("" & rs("gg6"))
  
  
  
        BatteryOutput.OG1D = Val("" & rs("og1d"))
        BatteryOutput.NG1D = Val("" & rs("ng1d"))
        BatteryOutput.GG1D = Val("" & rs("gg1d"))
  
  
        BatteryOutput.Repeats = Val("" & rs("Repeats"))
        BatteryOutput.RepeatUntil = Val("" & rs("RepeatUntil"))
        BatteryOutput.Pause = Val("" & rs("Pause"))
        BatteryOutput.SendCancel = Val("" & rs("SendCancel"))
        BatteryOutput.ScreenName = "" & rs("ScreenName")
    End Select
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  
  

LoadSettings_Resume:
   On Error GoTo 0
   Exit Sub

LoadSettings_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmOutputMask.LoadSettings." & Erl
  Resume LoadSettings_Resume

  
End Sub
Sub LoadOutputs()
  Dim rs As Recordset
  cboGroup1.Clear
  cboGroup2.Clear
  cboGroupN1.Clear
  cboGroupN2.Clear
  cboGroupG1.Clear
  cboGroupG2.Clear
  
  
  Set rs = ConnExecute("SELECT * FROM pagergroups ORDER BY Description")
  AddToCombo cboGroup1, "< none > ", 0
  AddToCombo cboGroup2, "< none > ", 0
  AddToCombo cboGroupN1, "< none > ", 0
  AddToCombo cboGroupN2, "< none > ", 0
  AddToCombo cboGroupG1, "< none > ", 0
  AddToCombo cboGroupG2, "< none > ", 0

  Do Until rs.EOF
    AddToCombo cboGroup1, rs("description") & "", rs("groupID")
    AddToCombo cboGroup2, rs("description") & "", rs("groupID")
    AddToCombo cboGroupN1, rs("description") & "", rs("groupID")
    AddToCombo cboGroupN2, rs("description") & "", rs("groupID")
    AddToCombo cboGroupG1, rs("description") & "", rs("groupID")
    AddToCombo cboGroupG2, rs("description") & "", rs("groupID")
    
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing

    cboGroup1.ListIndex = 0
    cboGroup2.ListIndex = 0
    cboGroupN1.ListIndex = 0
    cboGroupN2.ListIndex = 0
    cboGroupG1.ListIndex = 0
    cboGroupG2.ListIndex = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHost
End Sub

Private Sub TabStrip_BeforeClick(Cancel As Integer)
  
  DoDataExchange
  
End Sub
Sub DoDataExchange()
  Select Case LCase(TabStrip.SelectedItem.Key)
    Case "trouble"
        TroubleOutput.OG1 = GetComboItemData(cboGroup1)
        TroubleOutput.OG2 = GetComboItemData(cboGroup2)
        TroubleOutput.OG3 = 0
        TroubleOutput.OG4 = 0
        TroubleOutput.OG5 = 0
        TroubleOutput.OG6 = 0
        
        
        
        TroubleOutput.NG1 = GetComboItemData(cboGroupN1)
        TroubleOutput.NG2 = GetComboItemData(cboGroupN2)
        TroubleOutput.NG3 = 0
        TroubleOutput.NG4 = 0
        TroubleOutput.NG5 = 0
        TroubleOutput.NG6 = 0
        
        TroubleOutput.GG1 = GetComboItemData(cboGroupG1)
        TroubleOutput.GG2 = GetComboItemData(cboGroupG2)
        TroubleOutput.GG3 = 0
        TroubleOutput.GG4 = 0
        TroubleOutput.GG5 = 0
        TroubleOutput.GG6 = 0
        
        
        
        TroubleOutput.OG1D = Val(txtEscalate.text)
        
        TroubleOutput.Repeats = Val(txtRepeats.text)
        TroubleOutput.RepeatUntil = chkRepeatUntil.Value
        TroubleOutput.Pause = Val(txtPause.text)
        TroubleOutput.SendCancel = chkSendCancel.Value
        TroubleOutput.ScreenName = "Trouble"
      
'    Case "tamper"
'        TamperOutput.OG1 = GetComboItemData(cboGroup1)
'        TamperOutput.OG2 = GetComboItemData(cboGroup2)
'        TamperOutput.OG3 = 0
'        TamperOutput.NG1 = GetComboItemData(cboGroupN1)
'        TamperOutput.NG2 = GetComboItemData(cboGroupN2)
'        TamperOutput.NG3 = 0
'        TamperOutput.Repeats = Val(txtRepeats.text)
'        TamperOutput.RepeatUntil = chkRepeatUntil.value
'        TamperOutput.Pause = Val(txtPause.text)
'        TamperOutput.SendCancel = chkSendCancel.value
'        TamperOutput.ScreenName = "Tamper"
    
    
    Case "battery"
    
        BatteryOutput.OG1 = GetComboItemData(cboGroup1)
        BatteryOutput.OG2 = GetComboItemData(cboGroup2)
        BatteryOutput.OG3 = 0
        BatteryOutput.OG4 = 0
        BatteryOutput.OG5 = 0
        BatteryOutput.OG6 = 0
        
        BatteryOutput.NG1 = GetComboItemData(cboGroupN1)
        BatteryOutput.NG2 = GetComboItemData(cboGroupN2)
        BatteryOutput.NG3 = 0
        BatteryOutput.NG4 = 0
        BatteryOutput.NG5 = 0
        BatteryOutput.NG6 = 0
        
        BatteryOutput.GG1 = GetComboItemData(cboGroupG1)
        BatteryOutput.GG2 = GetComboItemData(cboGroupG2)
        BatteryOutput.GG3 = 0
        BatteryOutput.GG4 = 0
        BatteryOutput.GG5 = 0
        BatteryOutput.GG6 = 0
        
        
        
        BatteryOutput.OG1D = Val(txtBattEsc.text)
        
        BatteryOutput.Repeats = Val(txtRepeats.text)
        BatteryOutput.RepeatUntil = chkRepeatUntil.Value
        BatteryOutput.Pause = Val(txtPause.text)
        BatteryOutput.SendCancel = chkSendCancel.Value
        BatteryOutput.ScreenName = "Battery"
    
    
  End Select

End Sub

Private Sub TabStrip_Click()
  Fill
End Sub
Sub Fill()
  
  txtEscalate.text = TroubleOutput.OG1D
  txtBattEsc.text = BatteryOutput.OG1D
        
          
  Select Case LCase(TabStrip.SelectedItem.Key)
    Case "trouble"
        cboGroup1.ListIndex = Max(0, CboGetIndexByItemData(cboGroup1, TroubleOutput.OG1))
        cboGroup2.ListIndex = Max(0, CboGetIndexByItemData(cboGroup2, TroubleOutput.OG2))
        cboGroupN1.ListIndex = Max(0, CboGetIndexByItemData(cboGroupN1, TroubleOutput.NG1))
        cboGroupN2.ListIndex = Max(0, CboGetIndexByItemData(cboGroupN2, TroubleOutput.NG2))
        cboGroupG1.ListIndex = Max(0, CboGetIndexByItemData(cboGroupG1, TroubleOutput.GG1))
        cboGroupG2.ListIndex = Max(0, CboGetIndexByItemData(cboGroupG2, TroubleOutput.GG2))
        
        txtRepeats.text = TroubleOutput.Repeats
        txtPause.text = TroubleOutput.Pause
        chkRepeatUntil.Value = IIf(TroubleOutput.RepeatUntil = 1, 1, 0)
        chkSendCancel.Value = IIf(TroubleOutput.SendCancel = 1, 1, 0)
        lblTroubleEsc.Visible = True
        txtEscalate.Visible = True
        lblBattEsc.Visible = False
        txtBattEsc.Visible = False
'    Case "tamper"
'        cboGroup1.ListIndex = CboGetIndexByItemData(cboGroup1, TamperOutput.OG1)
'        cboGroup2.ListIndex = CboGetIndexByItemData(cboGroup2, TamperOutput.OG2)
'        cboGroupN1.ListIndex = CboGetIndexByItemData(cboGroupN1, TamperOutput.NG1)
'        cboGroupN2.ListIndex = CboGetIndexByItemData(cboGroupN2, TamperOutput.NG2)
'        txtRepeats.text = TamperOutput.Repeats
'        txtPause.text = TamperOutput.Pause
'        chkRepeatUntil.value = IIf(TamperOutput.RepeatUntil = 1, 1, 0)
'        chkSendCancel.value = IIf(TamperOutput.SendCancel = 1, 1, 0)
    
    
    Case "battery"
        cboGroup1.ListIndex = Max(0, CboGetIndexByItemData(cboGroup1, BatteryOutput.OG1))
        cboGroup2.ListIndex = Max(0, CboGetIndexByItemData(cboGroup2, BatteryOutput.OG2))
        cboGroupN1.ListIndex = Max(0, CboGetIndexByItemData(cboGroupN1, BatteryOutput.NG1))
        cboGroupN2.ListIndex = Max(0, CboGetIndexByItemData(cboGroupN2, BatteryOutput.NG2))
        cboGroupG1.ListIndex = Max(0, CboGetIndexByItemData(cboGroupG1, BatteryOutput.GG1))
        cboGroupG2.ListIndex = Max(0, CboGetIndexByItemData(cboGroupG2, BatteryOutput.GG2))
        
        
        txtRepeats.text = BatteryOutput.Repeats
        txtPause.text = BatteryOutput.Pause
        chkRepeatUntil.Value = IIf(BatteryOutput.RepeatUntil = 1, 1, 0)
        chkSendCancel.Value = IIf(BatteryOutput.SendCancel = 1, 1, 0)
        lblBattEsc.Visible = True
        txtBattEsc.Visible = True
        lblTroubleEsc.Visible = False
        txtEscalate.Visible = False
  
  
  End Select

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

Private Sub txtBattEsc_KeyPress(KeyAscii As Integer)
   KeyAscii = KeyProcMax(txtBattEsc, KeyAscii, False, 0, 4, 9999)
End Sub

Private Sub txtEscalate_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtEscalate, KeyAscii, False, 0, 4, 9999)
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
