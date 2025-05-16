VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPublicEvents 
   Caption         =   "Public Events"
   ClientHeight    =   3270
   ClientLeft      =   13500
   ClientTop       =   4365
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   9825
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9585
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
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
         Left            =   8385
         TabIndex        =   5
         Top             =   1785
         Visible         =   0   'False
         Width           =   1175
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   8385
         TabIndex        =   4
         Top             =   2370
         Width           =   1175
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   8385
         TabIndex        =   3
         Top             =   30
         Width           =   1175
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         Left            =   8385
         TabIndex        =   2
         Top             =   615
         Width           =   1175
      End
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
         Left            =   8385
         TabIndex        =   1
         Top             =   1200
         Width           =   1175
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   5265
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPublicEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    EditPublicEvent 0, 0
End Sub




Private Sub cmdDelete_Click()
  Dim ID As Long
  If Not lvMain.SelectedItem Is Nothing Then
     ID = Val(lvMain.SelectedItem.Key) ' 0 'ownerid doesn't care here
     DeleteReminder ID
     Fill
  Else
    Beep
  End If
End Sub

Private Sub cmdEdit_Click()
  If Not lvMain.SelectedItem Is Nothing Then
    EditPublicEvent Val(lvMain.SelectedItem.Key), 0 'ownerid doesn't care here
  Else
    Beep
  End If
End Sub

Private Sub cmdExit_Click()

  PreviousForm
  Unload Me
  
End Sub

Private Sub Form_Load()
ResetActivityTime
   SetControls
End Sub
Public Sub Fill()
  Dim rs              As ADODB.Recordset
  Dim li              As ListItem
  Dim Reminder        As cReminder
  Dim c               As Collection
  Dim Bothnames       As Boolean
  
  Set c = New Collection
  
  lvMain.ListItems.Clear
  Set rs = ConnExecute("SELECT reminders.* , staff.namelast, staff.namefirst FROM Reminders LEFT JOIN Staff ON Reminders.OwnerID = Staff.StaffID  WHERE Reminders.IsPublic = 1 ORDER BY Reminders.Description")
  Do Until rs.EOF
    DoEvents
    Set Reminder = New cReminder
    Reminder.Parse rs
    c.Add Reminder
    
    
    Bothnames = Len(rs("namelast") & "") > 0 And Len(rs("namefirst") & "") > 0
    
    
    Set li = lvMain.ListItems.Add(, Reminder.reminderid & "s", Reminder.ReminderName)
    ' set an expired flag
    li.SubItems(1) = LCase$(Reminder.FrequencyToString())
    li.SubItems(2) = Reminder.ScheduleToString()
    li.SubItems(3) = Reminder.TimeOfDayToString()
    li.SubItems(4) = rs("namelast") & IIf(Bothnames, ", ", "") & rs("namefirst")
    
    ' show the date?
    
    
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing




End Sub
Sub SetControls()
  Dim f As Control

  For Each f In Controls
    If TypeOf f Is Frame Then
      f.BackColor = Me.BackColor

    End If
  Next
  lvMain.ColumnHeaders.Clear
  lvMain.ColumnHeaders.Add , , "Name", 3000
  lvMain.ColumnHeaders.Add , , "F", 400
  lvMain.ColumnHeaders.Add , , "Days", 1500
  lvMain.ColumnHeaders.Add , , "Time", 1200
  lvMain.ColumnHeaders.Add , , "Owner", 1200
  
End Sub
Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub
Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub
