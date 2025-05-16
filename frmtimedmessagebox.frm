VERSION 5.00
Begin VB.Form frmTimedMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   1875
   ClientLeft      =   7305
   ClientTop       =   6615
   ClientWidth     =   3150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   1020
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "No"
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
      Left            =   1590
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
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
      Left            =   450
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Timer timerDismiss 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   30
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Height          =   720
      Left            =   810
      TabIndex        =   3
      Top             =   360
      Width           =   2115
   End
   Begin VB.Image imgIcon 
      Height          =   525
      Left            =   195
      Picture         =   "frmTimedMessageBox.frx":0000
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "frmTimedMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DefTimeOut = 30  ' seconds

Public response As Integer
Private AutoCancel As Integer

Public Property Let prompt(RHS As String)
  lblMessage.Caption = RHS
End Property
Public Property Let Title(RHS As String)
  Me.Caption = RHS
End Property
Public Property Let Timeout(RHS As Integer)
  If RHS <= 0 Then
    timerDismiss.Enabled = False
  Else
    timerDismiss.interval = Min(RHS * 1000, 32760)
    timerDismiss.Enabled = True
  End If

End Property

Public Property Let Buttons(ByVal Value As Long)

  Dim ButtonSet As Integer
  ButtonSet = Value And &HF

  Select Case ButtonSet
    Case vbOKCancel
      cmdYes.Visible = True
      cmdNo.Visible = True
      cmdOK.Visible = False
      AutoCancel = vbCancel
    Case vbAbortRetryIgnore
      cmdYes.Visible = True
      cmdNo.Visible = True
      'cmdCancel.Visible = True
      cmdOK.Visible = False
      AutoCancel = vbIgnore

    Case vbYesNoCancel
      cmdYes.Visible = True
      cmdNo.Visible = True
      'cmdCancel.Visible = True
      cmdOK.Visible = False
      AutoCancel = vbCancel

    Case vbYesNo
      cmdYes.Visible = True
      cmdNo.Visible = True
      cmdOK.Visible = False
      AutoCancel = vbNo

    Case vbRetryCancel

      AutoCancel = vbCancel
    Case Else
      cmdOK.Visible = True
  End Select

  If Value And vbYesNo Then
  End If

  If Value And 1 Then
    cmdYes.Visible = True
    cmdNo.Visible = True
    AutoCancel = vbNo
  End If

  If Value And vbQuestion Then
    LoadCustomRes imgIcon, vbQuestion, "MSGBOX"
  End If
  If Value And vbInformation Then
    LoadCustomRes imgIcon, vbInformation, "MSGBOX"
  End If
  If Value And vbExclamation Then
    LoadCustomRes imgIcon, vbExclamation, "MSGBOX"
  End If
  If Value And vbCritical Then
    LoadCustomRes imgIcon, vbCritical, "MSGBOX"
  End If


End Property


Private Sub cmdNo_Click()
  response = vbNo
  Unload Me

End Sub

Private Sub cmdOK_Click()
  response = vbOK
  Unload Me
End Sub

Private Sub cmdYes_Click()
  response = vbYes
  Unload Me

End Sub

Private Sub Form_Load()
ResetActivityTime
End Sub
Sub RestTimer()
  response = AutoCancel
  Unload Me

End Sub

Private Sub LoadCustomRes(pic As Object, ID As Integer, ResType As String)
  Dim a()     As Byte
  Dim hfile   As Integer
  On Error Resume Next

  a = LoadResData(ID, ResType)
  hfile = FreeFile
  Open "~~tmpres.tmp" For Binary Access Write As #hfile
  Put #hfile, , a
  Close #hfile
  pic.Picture = LoadPicture("~~tmpres.tmp")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  timerDismiss.Enabled = False
End Sub

Private Sub timerDismiss_Timer()
  response = AutoCancel
  Unload Me
End Sub
