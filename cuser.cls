VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public UserID       As Long
Public Username     As String
Public Password     As String
Private mLevel      As Long
Public ConsoleID    As String
Public LoggedOn     As Boolean
Public LastSeen     As Date
Public Session      As Long

Public UserPermissions As cUserPermissions


'Global Const LEVEL_FACTORY = 256
'Global Const LEVEL_ADMIN = 128
'Global Const LEVEL_SUPERVISOR = 32
'Global Const LEVEL_USER = 1

Public Function ParseHtPassword(ByVal UserString As String) As Boolean
  
  Dim UserPass() As String
  On Error Resume Next
  UserPass = Split(UserString, ":")
  Username = UserPass(0)
  ParseHtPassword = Username > 0

End Function


Public Property Get LevelString() As String
  If Username = "" Then
    LevelString = ""
  Else
    LevelString = GetLevelString(Level)
  End If
End Property

Public Property Get CanEditInfo() As Boolean
  Select Case Level
    Case LEVEL_FACTORY, LEVEL_ADMIN, LEVEL_SUPERVISOR
      CanEditInfo = True
    Case Else
      CanEditInfo = False
  End Select
  
End Property

Private Sub Class_Initialize()
  Set UserPermissions = New cUserPermissions
  LastSeen = Now
End Sub

Public Property Get Level() As Long

  Level = mLevel

End Property

Public Property Let Level(ByVal value As Long)
  Dim AppPath As String
  mLevel = value
  On Error Resume Next

  AppPath = App.Path
  If Right$(AppPath, 1) <> "\" Then
    AppPath = AppPath & "\"
  End If

  If mLevel >= LEVEL_FACTORY Then
    App.HelpFile = AppPath & "Help\HelpFactory.chm"

  ElseIf mLevel >= LEVEL_ADMIN Then
    App.HelpFile = AppPath & "Help\HelpAdmin2.chm"

  ElseIf mLevel >= LEVEL_SUPERVISOR Then
    App.HelpFile = AppPath & "Help\HelpAdmin1.chm"

  Else
    App.HelpFile = AppPath & "Help\HelpUser.chm"
  End If

End Property



