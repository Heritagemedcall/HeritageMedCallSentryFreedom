VERSION 5.00
Begin VB.Form frmAnnounce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Announce"
   ClientHeight    =   3660
   ClientLeft      =   465
   ClientTop       =   3795
   ClientWidth     =   10680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   315
      TabIndex        =   0
      Top             =   165
      Width           =   9060
      Begin VB.CommandButton cmdReminderSetup 
         Caption         =   "Reminder Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4320
         TabIndex        =   15
         Top             =   2220
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdStaff 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2625
         TabIndex        =   14
         Top             =   2220
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdReminders 
         Caption         =   "Edit Reminders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   930
         TabIndex        =   13
         Top             =   2220
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cboMessages 
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
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cboMessages"
         Top             =   540
         Width           =   7260
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   315
         Left            =   8520
         Picture         =   "frmAnnounce.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   345
      End
      Begin VB.CommandButton cmdAnnounceTime 
         Caption         =   "Announce Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7650
         TabIndex        =   11
         Top             =   1710
         Width           =   1275
      End
      Begin VB.ComboBox cboGroup 
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
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1545
         Width           =   2850
      End
      Begin VB.ComboBox cbopager 
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
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1095
         Width           =   2850
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7650
         TabIndex        =   10
         Top             =   1125
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7650
         TabIndex        =   12
         Top             =   2370
         Width           =   1275
      End
      Begin VB.TextBox txtMessage 
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
         Left            =   900
         MaxLength       =   200
         TabIndex        =   5
         Top             =   525
         Visible         =   0   'False
         Width           =   6990
      End
      Begin VB.Label lblMe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send an Announcement"
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
         Left            =   300
         TabIndex        =   1
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label lblGroup 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   330
         TabIndex        =   8
         Top             =   1590
         Width           =   525
      End
      Begin VB.Label lblPager 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pager"
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
         TabIndex        =   6
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label lblMessage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
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
         Left            =   90
         TabIndex        =   2
         Top             =   555
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmAnnounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReturnValue As Long

Public Sub Focus()
  On Error Resume Next
  

End Sub

Private Sub cboMessages_GotFocus()
  cboMessages.SelStart = 0
  cboMessages.SelLength = Len(cboMessages.text)
End Sub

Private Sub cboMessages_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyReturn

      Announce
  End Select

End Sub

Private Sub cboMessages_KeyUp(KeyCode As Integer, Shift As Integer)
  'AutoSel cboMessages, KeyCode

End Sub

Private Sub cmdAnnounceTime_Click()
  AnnounceTime
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
  DeleteCannedMessage cboMessages.text

End Sub

Private Sub cmdOK_Click()

  Announce

End Sub
Sub Announce()
'Dim pg As cPageItem
  Dim text As String
  On Error Resume Next
  'Set pg = New cPageItem
  'pg.Message = Trim(txtMessage.text)
  text = Trim(cboMessages.text)
  If cbopager.ListIndex > 0 Then
    If MASTER Then
      SendToPager text, GetComboItemData(cbopager), 0, "", "", PAGER_NORMAL, left$(text, 19), 0, 0
    Else
      ClientSendToPager text, GetComboItemData(cbopager), 0, "", 0
    End If
  End If
  If cboGroup.ListIndex > 0 Then
    If MASTER Then
      SendToGroup text, GetComboItemData(cboGroup), "", "", PAGER_NORMAL, left$(text, 19), 0, 0
    Else
      ClientSendToGroup text, GetComboItemData(cboGroup), "", 0
    End If
  End If

  AppendToCannedMessages text
  

End Sub

Function DeleteCannedMessage(ByVal text As String) As Long
  Dim SQL As String
  SQL = "DELETE FROM CannedMessages WHERE Message = " & q(text)
  ConnExecute SQL
  If Not MASTER Then
    ClientNotify "CannedMessages"
  End If

  FillMessages

End Function
Function AppendToCannedMessages(ByVal text As String) As Long
  Dim j As Integer
  Dim SQL As String
  text = Trim(left(text, 80))
  If Len(text) > 0 Then

    For j = cboMessages.listcount - 1 To 0 Step -1
      If 0 = StrComp(cboMessages.list(j), cboMessages.text) Then
        Exit For
      End If
    Next
    If j = -1 Then
      cboMessages.AddItem text
      SQL = "INSERT INTO cannedmessages (Message) Values (" & q(text) & ")"
      ConnExecute SQL
      If Not MASTER Then
        ClientNotify "CannedMessages"
      End If

    End If
  End If
End Function

Private Sub cmdReminders_Click()
  ShowReminders
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdReminderSetup_Click()
 ShowReminderConfig
 
End Sub

Private Sub cmdStaff_Click()
  ShowStaff 0, 0, 1, "Announce"
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
  SetControls
  Fill
  FillMessages
End Sub
Sub FillMessages()
  Dim SQL As String
  Dim Rs  As Recordset
  
  cboMessages.Clear
  SQL = "Select * FROM CannedMessages ORDER BY Message"
  Set Rs = ConnExecute(SQL)
  Do Until Rs.EOF
    cboMessages.AddItem Rs("Message") & ""
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing
End Sub
Sub SetControls()
  
  If NO_REMINDERS Then
    cmdReminders.Visible = False
    cmdStaff.Visible = False
    cmdReminderSetup.Visible = False
  Else
    cmdReminders.Visible = gUser.LEvel > LEVEL_USER
    cmdStaff.Visible = gUser.LEvel > LEVEL_USER
    cmdReminderSetup.Visible = gUser.LEvel > LEVEL_SUPERVISOR
  End If
  fraEnabler.BackColor = BackColor

End Sub
Public Sub Fill()
  Dim Rs As Recordset
  cbopager.Clear
  cboGroup.Clear

  Set Rs = ConnExecute("SELECT * FROM pagergroups ORDER BY groupname")
  AddToCombo cboGroup, "< none > ", 0
  Do Until Rs.EOF
    AddToCombo cboGroup, Rs("description") & "", Rs("groupID")
    Rs.MoveNext
  Loop
  Rs.Close

  Set Rs = ConnExecute("SELECT * FROM pagers order by description")
  AddToCombo cbopager, "< none > ", 0
  Do Until Rs.EOF
    AddToCombo cbopager, Rs("Description") & "", Rs("pagerid")
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing

  cbopager.ListIndex = 0
  cboGroup.ListIndex = 0


End Sub


Private Sub Form_Unload(Cancel As Integer)
  UnHost
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

Sub AnnounceTime()



  Dim pg As cPageItem
  Dim TimeString  As String
  
  
  On Error Resume Next
  'Set pg = New cPageItem
  'pg.Message = Trim(txtMessage.text)
  TimeString = Format(Now, gTimeFormatString)    ' "hh:nn AM/PM")
  If cbopager.ListIndex > 0 Then
    If MASTER Then
      SendToPager TimeString, GetComboItemData(cbopager), 0, "", "", PAGER_NORMAL, left$(TimeString, 19), 0, 0
    Else
      ClientSendToPager TimeString, GetComboItemData(cbopager), 0, "", 0
    End If
  End If
  If cboGroup.ListIndex > 0 Then
    If MASTER Then
      SendToGroup TimeString, GetComboItemData(cboGroup), "", "", PAGER_NORMAL, left$(TimeString, 19), 0, 0
    Else
      ClientSendToGroup TimeString, GetComboItemData(cboGroup), "", 0
    End If
  End If

End Sub



Private Sub txtMessage_GotFocus()
  SelAll txtMessage
End Sub

Private Sub txtMessage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'  If Button = vbRightButton Then
'    If txtMessage.text > "" Then
'      menuchoice = -1
'
'      PopupMenu mnuContext
'      Select Case menuchoice
'        Case MNU_DELETE
'          messagebox Me, "You chose Delete", App.Title, vbInformation
'      End Select
'    End If
'  End If
End Sub
