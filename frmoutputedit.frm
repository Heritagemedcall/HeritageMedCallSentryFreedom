VERSION 5.00
Begin VB.Form frmOutputEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output Edit"
   ClientHeight    =   3360
   ClientLeft      =   17655
   ClientTop       =   7575
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   8970
      Begin VB.Frame fraProtoOnTrak 
         BorderStyle     =   0  'None
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
         Left            =   2865
         TabIndex        =   19
         Top             =   1740
         Width           =   3375
         Begin VB.ComboBox cboRelayNum 
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
            Left            =   1605
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   60
            Width           =   1755
         End
         Begin VB.Label lblRelayNum 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relay Number"
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
            Left            =   0
            TabIndex        =   21
            Top             =   105
            Width           =   1500
         End
      End
      Begin VB.Frame fraProto7 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2820
         TabIndex        =   16
         Top             =   1740
         Width           =   3120
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
            ItemData        =   "frmOutputEdit.frx":0000
            Left            =   1665
            List            =   "frmOutputEdit.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label lblMarquis 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marquis Code"
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
            Left            =   75
            TabIndex        =   18
            Top             =   90
            Width           =   1995
         End
      End
      Begin VB.CheckBox chkNoName 
         Alignment       =   1  'Right Justify
         Caption         =   "No Resident Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2805
         TabIndex        =   22
         Top             =   2175
         Width           =   2490
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6150
         TabIndex        =   15
         Top             =   555
         Width           =   855
      End
      Begin VB.CheckBox chkIncludePhone 
         Alignment       =   1  'Right Justify
         Caption         =   "Include Phone #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   330
         TabIndex        =   14
         Top             =   2175
         Width           =   2130
      End
      Begin VB.TextBox txtIdentifier 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   6
         Top             =   570
         Width           =   3975
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   2
         Top             =   165
         Width           =   2670
      End
      Begin VB.TextBox txtDefaultMessage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2595
         Visible         =   0   'False
         Width           =   2670
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
         TabIndex        =   13
         Top             =   2370
         Width           =   1155
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
         TabIndex        =   12
         Top             =   1785
         Width           =   1155
      End
      Begin VB.ComboBox cbodevices 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1350
         Width           =   2850
      End
      Begin VB.TextBox txtPin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.CheckBox chkNoRepeats 
         Alignment       =   1  'Right Justify
         Caption         =   "No Repeats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   720
         TabIndex        =   11
         Top             =   1740
         Width           =   1725
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   855
         TabIndex        =   1
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label lblIdentifier 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address/ID/Ph#"
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
         Left            =   375
         TabIndex        =   5
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label lblDefMessage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Message"
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
         Left            =   270
         TabIndex        =   3
         Top             =   2610
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label lblOutputDevice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output Server"
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
         Left            =   615
         TabIndex        =   9
         Top             =   1395
         Width           =   1425
      End
      Begin VB.Label lblPin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin"
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
         Left            =   1725
         TabIndex        =   7
         Top             =   990
         Visible         =   0   'False
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmOutputEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PagerID      As Long
Public Serial       As String
Private protocol    As Long

Private Sub cbodevices_Click()

' get device protocol by text




  Dim pd As cPageDevice
  For Each pd In gPageDevices
    If 0 = StrComp(pd.Description, cbodevices.text, vbTextCompare) Then
      protocol = pd.ProtocolID
      Exit For
    End If
  Next
  'For Each oserver In gservers
  'Next
  chkNoRepeats.Visible = True
  chkNoName.Visible = True
  chkIncludePhone.Visible = True

  lblPin.Caption = "Pin"
  lblPin.Visible = False
  txtPin.Visible = False
  
  Select Case protocol

  Case PROTOCOL_MOBILE
    cmdFind.Visible = False
    fraProto7.Visible = False
    fraProtoOnTrak.Visible = False
    chkNoRepeats.Visible = False
    lblPin.Caption = "Mobile Group"
    lblPin.Visible = True
    txtPin.Visible = True

  Case PROTOCOL_REMOTE
    cmdFind.Visible = False
    fraProto7.Visible = False
    fraProtoOnTrak.Visible = False
    chkNoRepeats.Visible = False
    
    chkNoName.Visible = False
    chkIncludePhone.Visible = False
  Case PROTOCOL_PCA
    cmdFind.Visible = True
    fraProto7.Visible = False
    fraProtoOnTrak.Visible = False
    ' show pca lookup
    ' change text headings
  Case PROTOCOL_DIALER
    cmdFind.Visible = False
    fraProto7.Visible = False
    fraProtoOnTrak.Visible = False
    chkNoRepeats.Value = 1

'  Case PROTOCOL_APOLLO
'    fraProtoOnTrak.Visible = False
'    cmdFind.Visible = False
'    fraProto7.Visible = True


  Case PROTOCOL_MARQUIS
    fraProtoOnTrak.Visible = False
    cmdFind.Visible = False
    fraProto7.Visible = True

  Case PROTOCOL_ONTRAK
    fraProtoOnTrak.Visible = True
    cmdFind.Visible = False
    fraProto7.Visible = False
  Case PROTOCOL_TTS
    fraProtoOnTrak.Visible = True
    fraProto7.Visible = False
    cmdFind.Visible = False

  Case Else
    
    'fraProto7.Visible = False
    fraProto7.Visible = True
    fraProtoOnTrak.Visible = False
    cmdFind.Visible = False
  End Select

End Sub

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdFind_Click()
  Dim text As String
  If Save() Then
    If PagerID <> 0 Then

      FindPCAs txtIdentifier.text, PagerID
      Fill
    Else
      Beep
    End If
  Else
    Beep
  End If


End Sub



Private Sub cmdOK_Click()
  Dim t As String
  ResetActivityTime
  t = Trim(txtDescription.text)
  If Len(t) = 0 Then
    messagebox Me, "Please Fill in a Description", App.Title, vbInformation
  Else
    If Save() Then
      PreviousForm
      Unload Me
    Else
      messagebox Me, "Error Saving", App.Title, vbInformation
    End If
  End If
End Sub
Function Save() As Boolean
        Dim Rs            As Recordset
        Dim PagerID       As Long
10      On Error GoTo Save_Error

20      If Validate() Then

30        Save = True



          'Set rs = New ADODB.Recordset
          'If Me.PagerID = 0 Then
          'rs.Open "SELECT max( pagerid ) as maxpager FROM pagers", conn, gCursorType, gLockType
          'If Not rs.EOF Then
          '  PagerID = rs("maxpager") + 1
          'End If


40        Set Rs = New ADODB.Recordset

          'If Me.PagerID = 0 Then
50        Rs.Open "SELECT * FROM pagers WHERE pagerID = " & Me.PagerID, conn, gCursorType, gLockType
60        If Rs.EOF Then
70          Rs.addnew
            'Serial = "AA" & right$("000000" &
80        End If

90        Rs("Description") = Trim(txtDescription.text)
100       Rs("DeviceID") = GetComboItemData(cbodevices)
110       Rs("identifier") = Trim(txtIdentifier.text)
120       Rs("DefaultMessage") = Trim(txtDefaultMessage.text)
130       Rs("NoRepeats") = chkNoRepeats.Value

140       If protocol = PROTOCOL_REMOTE Then
150         Rs("NoRepeats") = 1
160       End If

170       Rs("IncludePhone") = chkIncludePhone.Value
180       Rs("NoName") = chkNoName.Value
190       Rs("PIN") = Trim$(txtPin.text)
200       Rs("marquiscode") = Max(0, cboMarquisCode.ListIndex)
210       Rs("relaynum") = Max(0, cboRelayNum.ListIndex)
220       Rs.Update
230       If Me.PagerID = 0 Then
240         Rs.MoveLast
250       End If
260       Me.PagerID = Rs("pagerID")
270       Rs.Close
280     End If


Save_Resume:
290     On Error Resume Next
300     Rs.Close
310     Set Rs = Nothing

320     On Error GoTo 0
330     Exit Function

Save_Error:

340     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmOutputEdit.Save." & Erl
350     Save = False
360     Resume Save_Resume

End Function
Private Function Validate() As Boolean
  txtDescription.text = Trim(txtDescription.text)
  Validate = (Len(txtDescription.text) > 0)
End Function

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
  FillCombos
  ResetForm

End Sub

Sub FillCombos()
  Dim Rs As Recordset

  cbodevices.Clear
  AddToCombo cbodevices, "None", 0
  Set Rs = ConnExecute("SELECT Description, ID FROM PagerDevices ORDER BY Description")
  Do Until Rs.EOF
    AddToCombo cbodevices, Rs("Description") & "", Rs("ID")
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing
  
  cboMarquisCode.Clear
  cboMarquisCode.AddItem "-None-" ' 0
  cboMarquisCode.AddItem "Normal" ' 1
  cboMarquisCode.AddItem "Emergency" '2
  cboMarquisCode.AddItem "Help" '3
  cboMarquisCode.AddItem "Help Lav" '4
  cboMarquisCode.AddItem "Info" '5
  cboMarquisCode.AddItem "Apollo" ' 6
  
  cboMarquisCode.ListIndex = 0
  
'  cboOTMode.Clear
'  cboOTMode.AddItem "Disabled"
'  cboOTMode.AddItem "Trigger"
'  cboOTMode.AddItem "Flashing"
'  cboOTMode.AddItem "Steady-on"
'  cboOTMode.ListIndex = 0
  
  cboRelayNum.Clear
  cboRelayNum.AddItem "-Not Specified-"
 
  cboRelayNum.AddItem "0"
  cboRelayNum.AddItem "1"
  cboRelayNum.AddItem "2"
  cboRelayNum.AddItem "3"
  cboRelayNum.AddItem "4"
  cboRelayNum.AddItem "5"
  cboRelayNum.AddItem "6"
  cboRelayNum.AddItem "7"
  
  cboRelayNum.ListIndex = 0
  
  

End Sub


Sub Fill()
  Dim Rs As Recordset
  Dim index As Long

  Set Rs = ConnExecute("SELECT * FROM Pagers WHERE Pagerid = " & PagerID)
  If Rs.EOF Then
    ResetForm
  Else
    cbodevices.ListIndex = Max(0, CboGetIndexByItemData(cbodevices, Rs("DeviceID")))
    
    txtDescription.text = Rs("Description") & ""
  
    'If Protocol <> PROTOCOL_PCA Then
      txtIdentifier.text = Rs("Identifier") & ""
    'Else
        'txtIdentifier.text =  Val("" & rs("RelayNum"))
    'End If
    
    txtDefaultMessage.text = Rs("DefaultMessage") & ""
    
    chkNoRepeats.Value = IIf(Rs("NoRepeats") = 1, 1, 0)
    chkIncludePhone.Value = IIf(Rs("IncludePhone") = 1, 1, 0)
    chkNoName.Value = IIf(Rs("NoName") = 1, 1, 0)
    
    txtPin.text = Rs("Pin") & ""
    
    index = Val("" & Rs("MarquisCode"))
    
    If index > -1 And index < cboMarquisCode.listcount Then
       cboMarquisCode.ListIndex = index
    Else
       cboMarquisCode.ListIndex = 0
    End If
    
    index = Val("" & Rs("RelayNum"))
    
    
    
    If index > -1 And index < cboRelayNum.listcount Then
       cboRelayNum.ListIndex = index
    Else
       cboRelayNum.ListIndex = 0
    End If

    index = Val("" & Rs("RelayNum"))


  End If
  Rs.Close


End Sub

Sub ResetForm()
  txtDescription.text = ""
  txtIdentifier.text = ""
  txtDefaultMessage.text = ""
  cboMarquisCode.ListIndex = 0
  cboRelayNum.ListIndex = 0
  cbodevices.ListIndex = 0
  
  chkNoRepeats.Value = 0
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
  ResetActivityTime
  UnHost
End Sub



Private Sub txtDefaultMessage_GotFocus()
  SelAll txtDefaultMessage

End Sub

Private Sub txtDescription_GotFocus()
  SelAll txtDescription
End Sub

Private Sub txtIdentifier_GotFocus()
  SelAll txtIdentifier

End Sub

Private Sub txtPin_GotFocus()
  SelAll txtPin

End Sub
