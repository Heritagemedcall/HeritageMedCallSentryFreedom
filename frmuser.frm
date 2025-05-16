VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "User"
   ClientHeight    =   3210
   ClientLeft      =   1650
   ClientTop       =   4080
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   9525
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Frame framPermissions 
         BorderStyle     =   0  'None
         Height          =   1290
         Left            =   5145
         TabIndex        =   14
         Top             =   765
         Width           =   2340
         Begin VB.CheckBox chkDeleteTrans 
            Caption         =   "Delete Transmitters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   825
            Width           =   2055
         End
         Begin VB.CheckBox chkDeleteRooms 
            Caption         =   "Delete Rooms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   465
            Width           =   1815
         End
         Begin VB.CheckBox chkDeleteRes 
            Caption         =   "Delete Residents"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   105
            Width           =   1815
         End
      End
      Begin VB.OptionButton optFactory 
         Caption         =   "Factory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2550
         TabIndex        =   11
         Top             =   1770
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.OptionButton optuser 
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2550
         TabIndex        =   10
         Top             =   1440
         Width           =   2490
      End
      Begin VB.OptionButton optsupervisor 
         Caption         =   "Admin 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2550
         TabIndex        =   9
         Top             =   1110
         Width           =   2490
      End
      Begin VB.OptionButton optAdmin 
         Caption         =   "Admin 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2550
         TabIndex        =   8
         Top             =   780
         Width           =   2490
      End
      Begin VB.TextBox txtpwd2 
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
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1890
         Width           =   1755
      End
      Begin VB.TextBox txtPwd1 
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
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1215
         Width           =   1755
      End
      Begin VB.TextBox txtUser 
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
         Left            =   180
         MaxLength       =   15
         TabIndex        =   2
         Top             =   525
         Width           =   1755
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
         Left            =   7665
         TabIndex        =   13
         Top             =   2475
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
         Left            =   7665
         TabIndex        =   12
         Top             =   1890
         Width           =   1175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permissions"
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
         Left            =   5160
         TabIndex        =   18
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access Level"
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
         Left            =   2550
         TabIndex        =   7
         Top             =   495
         Width           =   1155
      End
      Begin VB.Label lblpwd2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confrim Login"
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
         Left            =   180
         TabIndex        =   5
         Top             =   1590
         Width           =   1170
      End
      Begin VB.Label lblpwd1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Left            =   180
         TabIndex        =   3
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserID As Long
Private UserPermissions As cUserPermissions
Public Sub Fill()
  
  Dim Rs As Recordset
  
  Set UserPermissions = New cUserPermissions
  Set Rs = ConnExecute("select * from users where userid = " & UserID)
  If Not Rs.EOF Then
    txtUser.text = Rs("username") & ""
    txtPwd1.text = Rs("Password") & ""
    txtpwd2.text = txtPwd1.text
    Select Case Rs("Level")
      Case LEVEL_FACTORY
        Me.optFactory.Value = True
      Case LEVEL_ADMIN
        Me.optAdmin.Value = True
      Case LEVEL_SUPERVISOR
        optsupervisor.Value = True
      Case Else ' just a user
        optuser.Value = True
    End Select
    UserPermissions.ParseUserPermissions (Val(Rs("permissions") & ""))
    
  Else
    txtUser.text = ""
    txtPwd1.text = ""
    txtpwd2.text = ""
    optuser.Value = True
    UserID = 0
  End If
  Rs.Close
  Set Rs = Nothing
  
  chkDeleteTrans.Value = UserPermissions.CanDeleteTransmitters And 1
  chkDeleteRooms.Value = UserPermissions.CanDeleteRooms And 1
  chkDeleteRes.Value = UserPermissions.CanDeleteResidents And 1
  
  txtUser.Locked = (UserID <> 0)

End Sub


Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If Save() Then
    PreviousForm
    Unload Me
  End If

End Sub

Private Sub Form_Initialize()
  Set UserPermissions = New cUserPermissions
End Sub

Private Sub Form_Load()
ResetActivityTime
  SetControls
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

Function Save() As Boolean
  Dim LEvel As Long
    

  txtUser.text = Trim(txtUser.text)
  txtPwd1.text = Trim(txtPwd1.text)
  txtpwd2.text = Trim(txtpwd2.text)

  
  If optFactory.Value Then
    LEvel = LEVEL_FACTORY
  ElseIf optAdmin.Value Then
    LEvel = LEVEL_ADMIN
  ElseIf optsupervisor.Value Then
    LEvel = LEVEL_SUPERVISOR
  Else
    LEvel = LEVEL_USER
  End If
  
  UserPermissions.SetUserPermissions chkDeleteTrans.Value, chkDeleteRooms.Value, chkDeleteRes.Value
  
  If txtUser.text = "" Then
    messagebox Me, "Username Cannot be Blank.", App.Title, vbInformation
    Exit Function
  End If
  ' compare passwords
  If txtPwd1.text <> txtpwd2.text Then
    
    messagebox Me, "Passwords Must Match.", App.Title, vbInformation
    Exit Function
  End If

  If Len(txtPwd1.text) < 4 Then
    messagebox Me, "Logon Must be at least 4 characters", App.Title, vbInformation
    Exit Function
  End If

  If UserID = 0 Then
    ' check for dupe name
    If DupeUser(txtUser.text) Then
      messagebox Me, "Username Already Exists", App.Title, vbInformation
      Exit Function
    ElseIf DupeLogin(txtPwd1.text, UserID) Then
      messagebox Me, "Login Already Exists", App.Title, vbInformation
      Exit Function
    Else
      UserID = AddUser(txtUser.text, txtPwd1.text, LEvel, UserPermissions.UnParseUserPermissions)
      Save = UserID <> 0
    End If
  Else
    If DupeLogin(txtPwd1.text, UserID) Then
      messagebox Me, "Login Already Exists", App.Title, vbInformation
      Exit Function
    Else
      Save = UpdateUser(txtPwd1.text, UserID, LEvel, UserPermissions.UnParseUserPermissions)
    End If
  End If
  
  
  
  SyncApacheUsers
End Function
Function AddUser(ByVal Username As String, ByVal Password As String, ByVal LEvel As Long, ByVal Permissions As Long) As Long
  Dim Rs As Recordset
  Set Rs = New ADODB.Recordset
  Rs.Open "SELECT * FROM users WHERE userID = 0 ", conn, gCursorType, gLockType
  Rs.addnew
  Rs("username") = Username
  Rs("Password") = Password
  Rs("level") = LEvel
  Rs("Permissions") = Permissions
  Rs.Update
  Rs.MoveLast
  AddUser = Rs("userid")
  Rs.Close
  Set Rs = Nothing

End Function
Function UpdateUser(ByVal Password As String, ByVal UserID As Long, ByVal LEvel As Long, ByVal Permissions As Long) As Long
  
  Dim Rs As Recordset
  
  Set Rs = New ADODB.Recordset
  Rs.Open "SELECT * FROM users WHERE userID = " & UserID, conn, gCursorType, gLockType
  Rs("Password") = Password
  Rs("level") = LEvel
  Rs("Permissions") = Permissions
  Rs.Update
  Rs.MoveLast
  UpdateUser = Rs("userid")
  Rs.Close
  Set Rs = Nothing

End Function

Function DupeUser(ByVal Username As String) As Boolean
  Dim SQL As String
  Dim Rs As Recordset

  SQL = "select count(*) from users where username = " & q(Username)
  Set Rs = ConnExecute(SQL)
  DupeUser = Rs(0) > 0
  Rs.Close
End Function
Function DupeLogin(ByVal Pwd As String, ByVal ID As Long) As Boolean
  Dim SQL     As String
  Dim Rs      As Recordset
  Dim NewID   As Long
  
  DupeLogin = True
  
  SQL = "select userid from users where password = " & q(Pwd)
  Set Rs = ConnExecute(SQL)
  If Not Rs.EOF Then
    NewID = Rs("UserID")

  End If
  Rs.Close
  Set Rs = Nothing
      
    If (NewID <> 0) Then
      DupeLogin = (NewID <> ID)
    Else
      DupeLogin = False
    End If
    
  
End Function


Private Sub SetControls()
  Dim c As Control
  For Each c In Controls
    If TypeOf c Is Frame Then
      c.BackColor = Me.BackColor
    End If
  Next


  

  fraEnabler.BackColor = Me.BackColor

  fraEnabler.left = 0
  fraEnabler.top = 0

End Sub

