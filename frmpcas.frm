VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPCAs 
   Caption         =   "PCA List"
   ClientHeight    =   3120
   ClientLeft      =   255
   ClientTop       =   2295
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   9195
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         Left            =   7725
         TabIndex        =   5
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
         Left            =   7725
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
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
         Left            =   7725
         TabIndex        =   3
         Top             =   615
         Visible         =   0   'False
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
         Left            =   7725
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   1175
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
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
         TabIndex        =   1
         Top             =   1785
         Width           =   1175
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2985
         Left            =   15
         TabIndex        =   6
         Top             =   0
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgLst"
         SmallIcons      =   "imgLst"
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
            Text            =   "a"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPCAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PagerID        As Long
Public Serial         As String


Sub Apply()
  Dim DeviceID  As Long
  Dim SQl       As String
  Dim newserial As String
  
  If lvMain.SelectedItem Is Nothing Then
    ' nada
    Beep
  Else
    DeviceID = Val(lvMain.SelectedItem.Key)
    Dim d As cESDevice
    'Set d = Devices.Devices(DeviceID)
    newserial = lvMain.SelectedItem.text
    'If Not d Is Nothing Then
      SQl = "UPDATE pagers SET identifier = " & q(newserial) & " WHERE pagerID = " & PagerID
      ConnExecute SQl
      PreviousForm
      Unload Me
    'End If
      
  End If
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

Public Sub Fill()
  
  ShowTransmitters
  
End Sub

Private Sub cmdAdd_Click()
  EditTransmitter 0
End Sub

Public Sub ShowTransmitters()
    
  Configurelvmain
  RefreshDevices
End Sub


Sub Configurelvmain()
  Dim ch As ColumnHeader
  
  If lvMain.ColumnHeaders.Count < 6 Then
  
  
  lvMain.ColumnHeaders.Clear
  lvMain.Sorted = True
  Set ch = lvMain.ColumnHeaders.Add(, "S", "Serial", 1100)
  Set ch = lvMain.ColumnHeaders.Add(, "M", "Model", 1200)
  Set ch = lvMain.ColumnHeaders.Add(, "Desc", "Desc", 1440)
  Set ch = lvMain.ColumnHeaders.Add(, "Room", " ", 1440)
  Set ch = lvMain.ColumnHeaders.Add(, "Bldg", " ", 1350)
  Set ch = lvMain.ColumnHeaders.Add(, "Assur", " ", 700)
  
  
  End If
End Sub


Sub RefreshDevices()
  DisableButtons

  Dim SQl        As String
  Dim rs         As Recordset
  Dim li         As ListItem
  Dim index      As Long

  Dim SortedDevices As Collection
  Set SortedDevices = New Collection

  lvMain.ListItems.Clear
  lvMain.Sorted = True
  lvMain.SortKey = 0
  lvMain.SortOrder = lvwAscending

  SQl = " SELECT Devices.* FROM Devices WHERE Devices.model = 'EN3954' ORDER BY Devices.Serial"
  Set rs = ConnExecute(SQl)

  Do Until rs.EOF
    'For j = 1 To Devices.Devices.Count
    '  Set d = Devices.Devices(j)
    '  If d.Serial = rs("Serial") Then
    '    Exit For
    '  End If
    '  Set d = Nothing
    'Next
    'If Not d Is Nothing Then
      
      Set li = lvMain.ListItems.Add(, rs("DeviceID") & "B", Right("00000000" & rs("Serial"), 8))
      li.SubItems(1) = rs("Model") & ""
      li.SubItems(2) = rs("announce") & ""
      li.SubItems(3) = " "
      li.SubItems(4) = " "
      '   End If
'      If d.NumInputs > 1 Then
'        assure = IIf(d.UseAssur Or d.UseAssur2 = 1, "Y", "N") & " " & IIf(d.UseAssur_A Or d.UseAssur2_A = 1, "Y", "N")
'      Else
'        assure = IIf(d.UseAssur Or d.UseAssur2 = 1, "Y", "N")
'      End If
      li.SubItems(5) = " "
      
      If rs("serial") = Serial Then
        index = li.index
      End If
    'End If
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  
  If index > 0 And index <= lvMain.ListItems.Count Then
    lvMain.ListItems(index).Selected = True
    lvMain.ListItems(index).EnsureVisible
  End If
  EnableButtons

End Sub


Private Sub cmdApply_Click()
  Apply
End Sub


Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub Form_Activate()
 ' RefreshDevices
End Sub

Private Sub Form_Load()
ResetActivityTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub

Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  
  If lvMain.SortKey = ColumnHeader.index - 1 Then
    If lvMain.SortOrder = lvwAscending Then
      lvMain.SortOrder = lvwDescending
    Else
      lvMain.SortOrder = lvwAscending
    End If
  Else
    lvMain.SortOrder = lvwAscending
  End If
  lvMain.SortKey = ColumnHeader.index - 1

End Sub


Sub DisableButtons()
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdApply.Enabled = False
  cmdExit.Enabled = False
End Sub
Sub EnableButtons()
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  
End Sub

