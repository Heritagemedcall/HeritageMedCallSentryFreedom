VERSION 5.00
Begin VB.Form frmOutputGroup 
   Caption         =   "Output Group"
   ClientHeight    =   3315
   ClientLeft      =   9960
   ClientTop       =   7395
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Frame fraMain 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2040
         Left            =   120
         TabIndex        =   5
         Top             =   870
         Width           =   7320
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Del"
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
            Left            =   3345
            TabIndex        =   10
            Top             =   1125
            Width           =   615
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
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
            Left            =   3345
            TabIndex        =   9
            Top             =   450
            Width           =   615
         End
         Begin VB.ListBox lstAssigned 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            Left            =   4065
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   345
            Width           =   3120
         End
         Begin VB.ListBox lstAvail 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   135
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   345
            Width           =   3120
         End
         Begin VB.Label lblMembers 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Member Outputs"
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
            Left            =   4155
            TabIndex        =   7
            Top             =   75
            Width           =   1980
         End
         Begin VB.Label lblAvailable 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available Outputs"
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
            Left            =   825
            TabIndex        =   6
            Top             =   60
            Width           =   1515
         End
      End
      Begin VB.TextBox txtNotes 
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
         Left            =   1995
         MaxLength       =   50
         TabIndex        =   4
         Top             =   435
         Width           =   2670
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1995
         MaxLength       =   50
         TabIndex        =   2
         Top             =   75
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
         Left            =   7680
         TabIndex        =   14
         Top             =   2370
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
         Left            =   7680
         TabIndex        =   13
         Top             =   1785
         Width           =   1175
      End
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
         Left            =   7680
         TabIndex        =   12
         Top             =   30
         Visible         =   0   'False
         Width           =   1175
      End
      Begin VB.Label lblNotes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         TabIndex        =   3
         Top             =   480
         Width           =   510
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Description"
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
         TabIndex        =   1
         Top             =   150
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmOutputGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' With PCAs, each PCA is an output.
'

Public GroupID As Long

Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Height = ENABLER_HEIGHT
  SetParent fraEnabler.hwnd, hwnd
  fraEnabler.BackColor = Me.BackColor
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub


Private Sub cmdAdd_Click()
  ResetActivityTime
  If GroupID = 0 Then
    If DoSave() Then
      AddToGroup
    Else
      Beep
    End If
  Else
    AddToGroup
  End If
End Sub
Sub AddToGroup()
  Dim PagerID As Long
  Dim j       As Integer
  Dim SQl     As String
  
  PagerID = GetListBoxItemData(lstAvail)
  
  If PagerID <> 0 Then
    For j = lstAssigned.listcount - 1 To 0 Step -1
      If lstAssigned.ItemData(j) = PagerID Then
        Exit For
      End If
    Next
    If j < 0 Then ' we can't find it as assigned
      SQl = "INSERT INTO grouppager (groupid,pagerid) values (" & Join(Array(GroupID, PagerID), ",") & ")"
      ConnExecute SQl
      FillAssigned
    End If
  End If
  

End Sub

Sub RemoveFromGroup()
  Dim PagerID As Long
  Dim SQl As String
  
  PagerID = GetListBoxItemData(lstAssigned)
  If PagerID <> 0 Then
    SQl = "DELETE FROM grouppager WHERE GroupID = " & GroupID & " AND pagerid = " & PagerID
    ConnExecute SQl
    FillAssigned
  End If

End Sub
Private Sub cmdCancel_Click()
  
  PreviousForm
  Unload Me
  
End Sub


Private Sub cmdOK_Click()
  ResetActivityTime
  If DoSave() Then
  End If
End Sub
Function DoSave() As Boolean
  Dim t As String
  Dim Success As Boolean
  t = Trim(txtDescription.text)
  If Len(t) = 0 Then
    messagebox Me, "Please Fill in the Description for This Group", App.Title, vbInformation
  Else
    Success = Save()
    If Success Then
      
    Else
      messagebox Me, "Save Error", App.Title, vbInformation
    End If
  End If
  DoSave = Success
End Function

Function Save() As Boolean
  Save = True
  Dim rs As Recordset
  
  Save = True
  
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM pagergroups WHERE groupID = " & GroupID, conn, gCursorType, gLockType
  
  
  If rs.EOF Then
    rs.addnew
  End If
    
  rs("Description") = Trim(txtDescription.text)
  rs("notes") = Trim(txtNotes.text)
  rs.Update
  If GroupID = 0 Then
    rs.MoveLast
  End If
  GroupID = rs("GroupID")
  rs.Close
    

End Function
Sub Fill()
  Dim rs As Recordset
  ResetForm
  Set rs = ConnExecute("SELECT * FROM Pagergroups WHERE GROUPID = " & GroupID)
  If Not rs.EOF Then
    txtDescription.text = rs("Description") & ""
    txtNotes.text = rs("Notes") & ""
  End If
  rs.Close
  FillAvailable
  FillAssigned
End Sub
Sub ResetForm()
  txtDescription.text = ""
  txtNotes.text = ""
End Sub


Private Sub cmdRemove_Click()
  ResetActivityTime
  If lstAssigned.ListIndex > -1 Then
    RemoveFromGroup
  End If
End Sub

Sub FillAvailable()
  Dim rs As Recordset
  lstAvail.Clear
  Set rs = ConnExecute("SELECT * FROM Pagers")
  Do Until rs.EOF
    AddToListBox lstAvail, rs("Description") & "", rs("pagerid")
    rs.MoveNext
  Loop
  rs.Close
  
End Sub

Sub FillAssigned()
  Dim rs As Recordset
  Dim SQl As String
  lstAssigned.Clear
  
  
  SQl = "SELECT Pagers.Description, Pagers.PagerID, GroupPager.GroupID " & _
        " FROM Pagers INNER JOIN GroupPager ON Pagers.PagerID = GroupPager.PagerID " & _
        " WHERE GroupPager.GroupID = " & GroupID
        
  Set rs = ConnExecute(SQl)
  
  Do Until rs.EOF
    AddToListBox lstAssigned, rs("Description") & "", rs("pagerid")
    rs.MoveNext
  Loop
  rs.Close

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
  ArrangeControls
End Sub
Sub ArrangeControls()
  framain.BackColor = Me.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ResetActivityTime
  UnHost
End Sub

Private Sub txtDescription_GotFocus()
  SelAll txtDescription
End Sub

Private Sub txtNotes_GotFocus()
  SelAll txtNotes
End Sub
