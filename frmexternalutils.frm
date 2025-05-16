VERSION 5.00
Begin VB.Form frmExternalUtils 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   15
   ClientTop       =   5910
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   9555
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9300
      Begin VB.CommandButton cmdGetFile3 
         Height          =   330
         Left            =   7530
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmExternalUtils.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdGetFile2 
         Height          =   330
         Left            =   7530
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmExternalUtils.frx":052A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1860
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdGetFile1 
         Height          =   330
         Left            =   7530
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmExternalUtils.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdRun3 
         Caption         =   "Run"
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
         Left            =   180
         TabIndex        =   15
         Top             =   2340
         Width           =   600
      End
      Begin VB.TextBox txtFilename3 
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
         Left            =   1290
         TabIndex        =   19
         Top             =   2850
         Width           =   6225
      End
      Begin VB.TextBox txtParam3 
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
         TabIndex        =   17
         Top             =   2490
         Width           =   2655
      End
      Begin VB.TextBox txtUtilName3 
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
         TabIndex        =   16
         Top             =   2490
         Width           =   2655
      End
      Begin VB.CommandButton cmdRun2 
         Caption         =   "Run"
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
         Left            =   180
         TabIndex        =   9
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox txtFilename2 
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
         Left            =   1290
         TabIndex        =   13
         Top             =   1830
         Width           =   6225
      End
      Begin VB.TextBox txtParam2 
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
         TabIndex        =   11
         Top             =   1470
         Width           =   2655
      End
      Begin VB.TextBox txtUtilName2 
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
         TabIndex        =   10
         Top             =   1470
         Width           =   2655
      End
      Begin VB.CommandButton cmdRun1 
         Caption         =   "Run"
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
         Left            =   180
         TabIndex        =   3
         Top             =   330
         Width           =   600
      End
      Begin VB.TextBox txtFilename1 
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
         Left            =   1290
         TabIndex        =   7
         Top             =   840
         Width           =   6225
      End
      Begin VB.TextBox txtParam1 
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
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtUtilName1 
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
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdApply 
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
         Left            =   8055
         TabIndex        =   21
         Top             =   1815
         Width           =   1175
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
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
         Left            =   8055
         TabIndex        =   22
         Top             =   2430
         Width           =   1175
      End
      Begin VB.Label lblFile3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
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
         Left            =   930
         TabIndex        =   18
         Top             =   2910
         Width           =   315
      End
      Begin VB.Label lblFile2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
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
         Left            =   930
         TabIndex        =   12
         Top             =   1860
         Width           =   315
      End
      Begin VB.Label lblFile1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
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
         Left            =   930
         TabIndex        =   6
         Top             =   900
         Width           =   315
      End
      Begin VB.Label lblParam1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters"
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
         Left            =   3660
         TabIndex        =   2
         Top             =   210
         Width           =   960
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utlilty"
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
         Left            =   930
         TabIndex        =   1
         Top             =   210
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmExternalUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaller As String

Private Sub cmdApply_Click()
  ResetActivityTime
  If Save() Then
    PreviousForm
    Unload Me
  End If

End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub
Private Function Save() As Boolean
  WriteSetting "ExtApps", "Util1", Trim(txtUtilName1.text)
  WriteSetting "ExtApps", "App1", Trim(txtFilename1.text)
  WriteSetting "ExtApps", "Params1", Trim(txtParam1.text)
  WriteSetting "ExtApps", "Util2", Trim(txtUtilName2.text)
  WriteSetting "ExtApps", "App2", Trim(txtFilename2.text)
  WriteSetting "ExtApps", "Params2", Trim(txtParam2.text)
  WriteSetting "ExtApps", "Util3", Trim(txtUtilName3.text)
  WriteSetting "ExtApps", "App3", Trim(txtFilename3.text)
  WriteSetting "ExtApps", "Params3", Trim(txtParam3.text)

End Function

Private Sub cmdGetFile1_Click()
  Save
  GetExternalApp "1"
End Sub

Private Sub cmdGetFile2_Click()
  Save
  GetExternalApp "2"
End Sub

Private Sub cmdGetFile3_Click()
  Save
  GetExternalApp "3"
End Sub

Private Sub cmdRun1_Click()

  If Len(Trim(txtFilename1.text)) > 0 Then
    Runapp txtFilename1.text, Me.txtParam1.text
  End If
  
End Sub
Private Sub Runapp(ByVal exename As String, ByVal params As String)
  
  ResetActivityTime
  exename = Trim(exename)
  params = Trim(params)
  On Error Resume Next
  Shell exename & " " & params, vbNormalFocus
  

End Sub


Private Sub cmdRun2_Click()
  
  If Len(Trim(txtFilename2.text)) > 0 Then
    Runapp txtFilename2.text, Me.txtParam2.text
  End If

End Sub

Private Sub cmdRun3_Click()
  If Len(Trim(txtFilename3.text)) > 0 Then
    Runapp txtFilename3.text, Me.txtParam3.text
  End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      SendKeys "{tab}"
  End Select


End Sub

Private Sub Form_Load()
  ResetActivityTime
  fraEnabler.BackColor = Me.BackColor
  ArrangeControls
End Sub
Public Sub Fill()
    
  txtUtilName1.text = ReadSetting("ExtApps", "Util1", "")
  txtFilename1.text = ReadSetting("ExtApps", "App1", "")
  txtParam1.text = ReadSetting("ExtApps", "Params1", "")
  
  txtUtilName2.text = ReadSetting("ExtApps", "Util2", "")
  txtFilename2.text = ReadSetting("ExtApps", "App2", "")
  txtParam2.text = ReadSetting("ExtApps", "Params2", "")
  
  txtUtilName3.text = ReadSetting("ExtApps", "Util3", "")
  txtFilename3.text = ReadSetting("ExtApps", "App3", "")
  txtParam3.text = ReadSetting("ExtApps", "Params3", "")

End Sub
Public Sub Host(ByVal hwnd As Long)
  
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  SetParent fraEnabler.hwnd, hwnd
  fraEnabler.BackColor = Me.BackColor
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
  Caller = ""
End Sub
  
Private Sub ArrangeControls()
  Select Case gUser.LEvel
    Case LEVEL_FACTORY
      'all available
      
    Case LEVEL_ADMIN
      txtUtilName1.Locked = True
      txtUtilName2.Locked = True
      txtUtilName3.Locked = True
      txtParam1.Visible = False
      txtParam2.Visible = False
      txtParam3.Visible = False
      lblParam1.Visible = False
      lblFile1.Visible = False
      lblFile2.Visible = False
      lblFile3.Visible = False
      txtFilename1.Visible = False
      txtFilename2.Visible = False
      txtFilename3.Visible = False
      cmdGetFile1.Visible = False
      cmdGetFile2.Visible = False
      cmdGetFile3.Visible = False
      cmdApply.Enabled = False
    
    Case LEVEL_SUPERVISOR
      txtUtilName1.Locked = True
      txtUtilName2.Locked = True
      txtUtilName3.Locked = True
      
      txtParam1.Visible = False
      txtParam2.Visible = False
      txtParam3.Visible = False
      
      lblParam1.Visible = False
      lblFile1.Visible = False
      lblFile2.Visible = False
      lblFile3.Visible = False
      
      txtFilename1.Visible = False
      txtFilename2.Visible = False
      txtFilename3.Visible = False
      
      cmdGetFile1.Visible = False
      cmdGetFile2.Visible = False
      cmdGetFile3.Visible = False
      cmdApply.Enabled = False
    
    Case Else ' "BASIC USER"
      cmdRun1.Enabled = False
      cmdRun2.Enabled = False
      cmdRun3.Enabled = False
            
      lblParam1.Visible = False
      
      txtUtilName1.Locked = True
      txtUtilName2.Locked = True
      txtUtilName3.Locked = True
      
      lblFile1.Visible = False
      lblFile2.Visible = False
      lblFile3.Visible = False
      
      txtFilename1.Visible = False
      txtFilename2.Visible = False
      txtFilename3.Visible = False
      
      txtParam1.Visible = False
      txtParam2.Visible = False
      txtParam3.Visible = False
      
      cmdGetFile1.Visible = False
      cmdGetFile2.Visible = False
      cmdGetFile3.Visible = False
      cmdApply.Enabled = False
  End Select
  
End Sub

Public Property Get Caller() As String
  Caller = mCaller
End Property

Public Property Let Caller(ByVal Caller As String)
  mCaller = Caller
End Property
