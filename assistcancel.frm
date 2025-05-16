VERSION 5.00
Begin VB.Form frmAssistCancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff Assist Acknowledge and Disposition"
   ClientHeight    =   4185
   ClientLeft      =   7605
   ClientTop       =   5805
   ClientWidth     =   9450
   Icon            =   "AssistCancel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDisp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   540
      MaxLength       =   50
      TabIndex        =   1
      Top             =   750
      Width           =   6345
   End
   Begin VB.ListBox lstDisp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   540
      TabIndex        =   3
      Top             =   1290
      Width           =   6480
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Yes"
      Enabled         =   0   'False
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
      Left            =   7665
      TabIndex        =   2
      Top             =   915
      Width           =   1175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "No"
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
      Left            =   7665
      TabIndex        =   4
      Top             =   1575
      Width           =   1175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Acknowledging an Assistance Call Will Terminate the Call"
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
      TabIndex        =   5
      Top             =   135
      Width           =   4905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disposition"
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
      TabIndex        =   0
      Top             =   465
      Width           =   945
   End
End
Attribute VB_Name = "frmAssistCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Disposition       As String
' only add if returned



Private Sub cmdApply_Click()
  Dim DispText           As String

  SyncList

  DispText = Trim$(txtDisp.text)
  If Len(DispText) Then
    'write changes to DB
    ' return DispText
    Disposition = DispText
    Unload Me
  Else
    messagebox Me, "Please Supply a Disposition", "Ack Staff Assist", vbCritical Or vbOKOnly
  End If


End Sub

Sub SyncList()
  Dim j                  As Long
  Dim DispText           As String
  Dim newitem

  DispText = Trim$(txtDisp.text)
  If Len(DispText) Then
    For j = lstDisp.listcount - 1 To 0 Step -1
      If 0 = StrComp(lstDisp.list(j), DispText, vbTextCompare) Then
        lstDisp.ListIndex = j
        lstDisp.Selected(j) = True

        Exit For
      End If
    Next
    If j = -1 Then

      lstDisp.AddItem (DispText)
      newitem = lstDisp.NewIndex
      lstDisp.Selected(newitem) = True
      lstDisp.ListIndex = newitem
    End If
  End If


End Sub

Private Sub cmdExit_Click()
  Disposition = ""
  Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> 1 Then
    Disposition = ""
  End If
End Sub

Private Sub lstDisp_DblClick()
  txtDisp.text = Me.lstDisp.text
End Sub

Private Sub txtdisp_Change()
  'Me.Caption = txtDisp.text
  Dim disp               As String
  disp = Trim$(txtDisp.text)
  If Len(disp) Then
    Me.cmdApply.Enabled = True
  Else
    Me.cmdApply.Enabled = False
  End If
End Sub

Private Sub txtDisp_KeyPress(KeyAscii As Integer)
  Dim j                  As Long
  Dim DispText           As String
  DispText = Trim(txtDisp.text)
  If Len(DispText) Then
    If KeyAscii = vbKeyReturn Then
      SyncList
    End If


  End If
End Sub

Private Sub Form_Load()

  Dim SQL                As String
  Dim rs                 As ADODB.Recordset

  SQL = "SELECT text FROM dispositions"

  lstDisp.Clear
  Set rs = ConnExecute(SQL)

  Do Until rs.EOF
    lstDisp.AddItem rs("text") & ""
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing


  '  Me.lstDisp.AddItem "Row 1"
  '  Me.lstDisp.AddItem "Row 2"
  '  Me.lstDisp.AddItem "Row 3"
  '
  '  Me.lstDisp.AddItem "Row 4"
  '  Me.lstDisp.AddItem "Row 5"
  '  Me.lstDisp.AddItem "Row 6"
  '
  '
  '  Me.lstDisp.AddItem "Row 7"
  '  Me.lstDisp.AddItem "Row 8"
  '  Me.lstDisp.AddItem "Row 9"
  '
  '  Me.lstDisp.AddItem "Row 10"
  '  Me.lstDisp.AddItem "Row 11"
  '  Me.lstDisp.AddItem "Row 12"
  '




End Sub

Private Sub txtDisp_LostFocus()
  Dim j                  As Long
  Dim DispText           As String
  DispText = Trim(txtDisp.text)
  If Len(DispText) Then

    SyncList


  End If

End Sub
