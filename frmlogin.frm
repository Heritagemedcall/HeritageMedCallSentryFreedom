VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Login"
   ClientHeight    =   4935
   ClientLeft      =   4410
   ClientTop       =   4515
   ClientWidth     =   5025
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   450
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1290
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd4 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   450
      Picture         =   "frmLogin.frx":1172
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd7 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   450
      Picture         =   "frmLogin.frx":22A2
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3030
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmdStar 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   450
      Picture         =   "frmLogin.frx":33A8
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3900
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd2 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   1485
      Picture         =   "frmLogin.frx":44D8
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1290
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd5 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   1485
      Picture         =   "frmLogin.frx":55A6
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd8 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   1485
      Picture         =   "frmLogin.frx":66C4
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3030
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd0 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   1485
      Picture         =   "frmLogin.frx":77CE
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3900
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd3 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   2520
      Picture         =   "frmLogin.frx":8906
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1290
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd6 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   2520
      Picture         =   "frmLogin.frx":99DC
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmd9 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   2520
      Picture         =   "frmLogin.frx":AB18
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3030
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmdPound 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   2520
      Picture         =   "frmLogin.frx":BC64
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3900
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3780
      Picture         =   "frmLogin.frx":CD7E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1845
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3780
      Picture         =   "frmLogin.frx":DECC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3045
      Width           =   990
   End
   Begin VB.TextBox txtLogin 
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
      IMEMode         =   3  'DISABLE
      Left            =   510
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   735
      Width           =   2985
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2445
      TabIndex        =   16
      Top             =   4965
      Width           =   135
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Login"
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
      Left            =   1350
      TabIndex        =   0
      Top             =   255
      Width           =   1125
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Quit As Boolean
Public ShowCountdown As Boolean
Private mCountDown As Long

Private keymatrix(0 To 3, 0 To 4) As String


Sub DoKeypress(ByVal Col As Integer, ByVal row As Integer)
  Dim KeyCode As String
  ClickSound
  KeyCode = keymatrix(Col, row)

  Select Case KeyCode
    Case "*"
      txtLogin.text = ""
    Case "#"
      If Len(txtLogin.text) > 0 Then
        txtLogin.text = left(txtLogin.text, Len(txtLogin.text) - 1)
        txtLogin.SelStart = Len(txtLogin.text)
      End If
    Case Else
      If Len(KeyCode) Then
        SendKeypress2Window txtLogin.hwnd, KeyCode
      End If
  End Select

End Sub
Private Sub ClickSound()
'Make a click sound for button press
End Sub

Private Sub cmd0_Click()
  DoKeypress 2, 4
  SetFocusTo txtLogin

End Sub

Private Sub cmd1_Click()
  DoKeypress 1, 1
  SetFocusTo txtLogin

End Sub

Private Sub cmd2_Click()
  DoKeypress 2, 1
  SetFocusTo txtLogin

End Sub

Private Sub cmd3_Click()
  DoKeypress 3, 1
  SetFocusTo txtLogin

End Sub

Private Sub cmd4_Click()
  DoKeypress 1, 2
  SetFocusTo txtLogin

End Sub

Private Sub cmd5_Click()
  DoKeypress 2, 2
  SetFocusTo txtLogin

End Sub

Private Sub cmd6_Click()
  DoKeypress 3, 2
  SetFocusTo txtLogin

End Sub

Private Sub cmd7_Click()
  DoKeypress 1, 3
  SetFocusTo txtLogin

End Sub

Private Sub cmd8_Click()
  DoKeypress 2, 3
  SetFocusTo txtLogin

End Sub

Private Sub cmd9_Click()
  DoKeypress 3, 3
  SetFocusTo txtLogin

End Sub

Private Sub cmdPound_Click()
  DoKeypress 3, 4
  SetFocusTo txtLogin

End Sub

Private Sub cmdStar_Click()
  DoKeypress 1, 4
  SetFocusTo txtLogin

End Sub

Private Sub cmdCancel_Click()

  ' Quit Program
  
  Dim user As cUser
  Set user = GetUser(Trim(txtLogin.text))
  If user.Level >= LEVEL_SUPERVISOR Then  ' factory and admin ? add supervisor??
    If MASTER Then
      If Configuration.WatchdogType > 0 Then
        If MsgBox("Shutting Down Will Disable Watchdog!" & vbCrLf & "  Proceed With Shut Down?", vbCritical Or vbQuestion Or vbYesNo, "SHUT DOWN") = vbNo Then
          Exit Sub
        End If
        SetWatchdog 0
        Set BerkshireWD = New cBerkshire


      End If
      PingMonitor True, "facilityid=" & Configuration.MonitorFacilityID & "&eventcode=" & EVENT_FACILITY_SHUTDOWN
    End If
    
    
    frmTimer.StopTimer
    
    QuitSession
    PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_STOP, 0
    LoggedIn = False
    Me.Visible = False
    
    Win32.SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    
    Quit = True
    DestroyObjects
    Unload frmMain
    Unload frmTimer
    
    End
''      Unload frmLogin
''

''      QuitSession
''      Unload Me
    
    
    'Quit = True
    'Me.Hide
  Else
    Beep
    lblLogin.Caption = "Supervisor Password Required." & vbCrLf & "Please ReEnter."
    SetFocusTo txtLogin
  End If

End Sub


''    If frmLogin.Quit Then
''      Unload frmLogin
''      frmTimer.StopTimer
''      Unload frmTimer
''      PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_STOP, 0
''      LoggedIn = False
''      QuitSession
''      Unload Me
''
''
''    Else
''      Me.Enabled = True
''      Unload frmLogin
''      UpdateScreenElements
''      txtLogin.text = ""
''
''      PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_LOGIN, 0
''      LoggedIn = True
''    End If



Private Sub cmdOK_Click()
  DoOK
End Sub
Sub DoOK()
  If ProcessLogin(txtLogin.text) Then
    Me.Visible = False
    Win32.SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    frmMain.Enabled = True
    frmMain.UpdateScreenElements
    frmMain.txtLogin.text = ""

    PostEvent Nothing, Nothing, Nothing, EVT_SYSTEM_LOGIN, 0
    LoggedIn = True
    
    Unload Me
    
  Else
    Beep
    lblLogin.Caption = "Login Failure. Please Reenter."
    SetFocusTo txtLogin
  End If
  frmLogin.Caption = "System Login"
End Sub

Private Sub Form_Initialize()
' pushbutton keys on screen

  keymatrix(1, 1) = "1"
  keymatrix(2, 1) = "2"
  keymatrix(3, 1) = "3"

  keymatrix(1, 2) = "4"
  keymatrix(2, 2) = "5"
  keymatrix(3, 2) = "6"

  keymatrix(1, 3) = "7"
  keymatrix(2, 3) = "8"
  keymatrix(3, 3) = "9"

  keymatrix(1, 4) = "*"  ' clear
  keymatrix(2, 4) = "0"
  keymatrix(3, 4) = "#"  ' back

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  Select Case KeyAscii
'    Case vbKeyReturn
'      KeyAscii = 0
'      SendKeys "{tab}"
'  End Select



End Sub

Private Sub Form_Load()
  
  
  'Win32.SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
  
  Win32.SetWindowPos Me.hwnd, Win32.HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE
  
  CenterFormOnForm Me, Nothing
  
  
  Me.Caption = IIf(MASTER, "", "REMOTE CONSOLE ") & "System Login"
  lblWarning.Visible = False
  Call GetLicensing
  If gRegistered Then
    ' OK
  ElseIf gSentinel.Expired Then
    lblWarning.Caption = "Evaluation Period Expired"
    Me.Height = Me.Height + lblWarning.Height + 120
    lblWarning.Visible = True
  Else
    lblWarning.Caption = "Evaluation Period Ends in " & gDaysLeft & " Days"
    Me.Height = Me.Height + lblWarning.Height + 120
    lblWarning.Visible = True
  End If


End Sub

Private Sub txtLogin_GotFocus()
'  SelAll txtLogin
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyReturn
      KeyAscii = 0
      DoOK

  End Select

End Sub



Public Property Let CountDown(ByVal Value As Long)
  If ShowCountdown Then
    Me.Caption = "System Login (" & Value & ")"
  Else
   If Me.Caption <> "System Login" Then
    Me.Caption = "System Login"
   End If
  End If
   
   
  

End Property
