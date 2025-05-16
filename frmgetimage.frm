VERSION 5.00
Begin VB.Form frmGetImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Image"
   ClientHeight    =   3945
   ClientLeft      =   1095
   ClientTop       =   6540
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3500
      Left            =   30
      TabIndex        =   0
      Top             =   180
      Width           =   9150
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
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
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
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
         Top             =   1200
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
         TabIndex        =   5
         Top             =   1785
         Width           =   1175
      End
      Begin VB.DirListBox lstFolders 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   2520
      End
      Begin VB.DriveListBox lstDrives 
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
         Left            =   60
         TabIndex        =   3
         Top             =   2820
         Width           =   2520
      End
      Begin VB.FileListBox lstFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   2655
         TabIndex        =   2
         Top             =   135
         Width           =   2700
      End
      Begin VB.Image imgPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   5625
         Stretch         =   -1  'True
         Top             =   150
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmGetImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ResidentID As Long
Public filename   As String

Private LastDrive As String

Public Sub Fill()

End Sub

Function FullPath() As String
  Dim path As String
  path = lstFolders.path
  If Right(path, 1) <> "\" Then
    path = path & "\"
  End If
  FullPath = path & lstFiles.filename
End Function

Function Save() As Boolean

  Dim rs  As Recordset
  Dim filename As String
  filename = FullPath
  
  
  
    
  
  'ResidentID = 8
  If ResidentID <> 0 Then
    On Error Resume Next
    Set rs = New ADODB.Recordset '
    rs.Open "SELECT * FROM Residents WHERE ResidentID = " & ResidentID, conn, gCursorType, gLockType
    If Not rs.EOF Then
      
      If Len(filename) Then
        rs("Imagepath") = filename
        SaveImageToDB filename, rs("imagedata")
      Else
        rs("Imagepath") = ""
        rs("imagedata") = Null
      End If
      rs.Update
        
    End If
    rs.Close
    Save = Err.Number = 0
  End If
End Function
Function Remove() As Boolean
  Dim SQl As String
  If ResidentID <> 0 Then
    On Error Resume Next
    Dim rs  As Recordset
    Set rs = New ADODB.Recordset '
    SQl = "SELECT * FROM Residents WHERE ResidentID = " & ResidentID
    rs.Open SQl, conn, gCursorType, gLockType
    If Not rs.EOF Then
      rs("Imagepath") = ""
      rs("imagedata") = Null
      rs.Update
    End If
    rs.Close
    Set rs = Nothing

    Remove = Err.Number = 0
  End If

End Function

Private Sub cmdApply_Click()
  ResetActivityTime
  If Save() Then
      PreviousForm
      Unload Me
  End If
End Sub

Private Sub cmdExit_Click()
  ResetActivityTime
    PreviousForm
    Unload Me

End Sub

Private Sub cmdRemove_Click()
  ResetActivityTime
  If Remove() Then
    PreviousForm
    Unload Me
  
  End If
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
  Connect
  fraEnabler.BackColor = Me.BackColor
  lstFiles.Pattern = "*.jpg"
  lstFiles.path = lstFolders.path
  LastDrive = lstDrives.Drive
End Sub


Private Sub lstFiles_Click()
  ShowPicture
End Sub
Sub ShowPicture()
  On Error Resume Next
  If Len(lstFiles.filename) Then
    imgPic.Picture = LoadPicture(FullPath)
  Else
    imgPic.Picture = LoadPicture("")
  End If

End Sub


Private Sub lstFolders_Change()
  Static Busy As Boolean

  On Error Resume Next

  If Not Busy Then
  
    Busy = True
    ' Change the current directory
    ChDir lstFolders.path

    If Err.Number = 0 Then
      lstFiles.path = lstFolders.path
      lstDrives.Drive = left(lstFolders.path, 2)
    Else
      Err.Clear
    End If

    Busy = False
    ShowPicture
  End If
  
End Sub

Private Sub lstDrives_Change()

  Dim Retry As Boolean
  On Error Resume Next

  Retry = True
  Do While Retry
    Retry = False
    lstFiles.path = lstFolders.path
    lstFolders.path = lstDrives.Drive
    ' If and error occurs
    Select Case Err.Number
      Case 68  ' Not accessable Error
        If vbRetry = messagebox(Me, lstDrives.Drive & " is not accessible", App.Title, vbRetryCancel + vbCritical) Then
          Retry = True
        Else  'Switch to previous known drive
          lstDrives.Drive = LastDrive
        End If
      Case 0
        ' done
      Case Else
        ' Ooops
         messagebox Me, "Unexpected File Access Error " & Err.Number & " : " & Err.Description, App.Title, vbInformation
        lstDrives.Drive = LastDrive
    End Select
  Loop
  ShowPicture
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

