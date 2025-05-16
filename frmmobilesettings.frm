VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmMobileSettings 
   Caption         =   "Mobile Settings"
   ClientHeight    =   3180
   ClientLeft      =   5670
   ClientTop       =   9870
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   9090
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         TabIndex        =   4
         Top             =   825
         Width           =   1175
      End
      Begin VB.TextBox txtPhrase 
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
         Left            =   570
         MaxLength       =   50
         TabIndex        =   2
         Top             =   345
         Width           =   6000
      End
      Begin MSComctlLib.ListView lvPhrases 
         Height          =   2220
         Left            =   555
         TabIndex        =   3
         Top             =   720
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   3916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Phrase"
            Object.Width           =   10583
         EndProperty
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Disposition Phrases"
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
         Left            =   555
         TabIndex        =   1
         Top             =   105
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmMobileSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub GetPhrases()
 Dim rs As ADODB.Recordset
 Dim li As ListItem
 
 
 lvPhrases.ListItems.Clear
 Set rs = ConnExecute("SELECT * FROM Dispositions ORDER BY ID DESC")
 Do Until rs.EOF
  Set li = lvPhrases.ListItems.Add(, rs("ID") & "s", rs("Text") & "")
  rs.MoveNext
 Loop
 rs.Close
 Set rs = Nothing
 cmdDelete.Enabled = False
End Sub


Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub


Private Sub cmdDelete_Click()
  Dim li As ListItem
  For Each li In Me.lvPhrases.ListItems
    If li.Checked Then
      ConnExecute "DELETE FROM Dispositions WHERE ID = " & Val(li.Key)
    End If
  Next
  GetPhrases
  
End Sub

Private Sub cmdOK_Click()
  ResetActivityTime
  Dim Phrase As String
  Phrase = Trim$(txtPhrase.Text)
  If Len(Phrase) > 0 Then
    ConnExecute "INSERT INTO Dispositions (Text) values (" & q(Phrase) & ")"
    txtPhrase.Text = ""
  End If
  GetPhrases

End Sub
Public Sub Host(ByVal hwnd As Long)
  fraEnabler.BackColor = Me.BackColor
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Sub Fill()
  GetPhrases
End Sub


Private Sub lvPhrases_Click()
  Dim li As ListItem
  Me.cmdDelete.Enabled = False
  For Each li In Me.lvPhrases.ListItems
    If li.Checked Then
      Me.cmdDelete.Enabled = True
      Exit For
    End If
  Next
End Sub
