VERSION 5.00
Begin VB.Form frmTemperatureDevice 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   6000
   ClientTop       =   5535
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   9135
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.CommandButton cmdResetSetPoint 
         Caption         =   "Reset"
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
         Left            =   7740
         TabIndex        =   18
         Top             =   660
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cboMode_A 
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
         Left            =   4440
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox cboMode 
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
         Left            =   1980
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtHigh_A 
         Alignment       =   1  'Right Justify
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
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   13
         ToolTipText     =   "Number of characters to skip from beginning"
         Top             =   1515
         Width           =   1170
      End
      Begin VB.TextBox txtLow_A 
         Alignment       =   1  'Right Justify
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
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   11
         ToolTipText     =   "Number of characters to skip from beginning"
         Top             =   1140
         Width           =   1170
      End
      Begin VB.TextBox txtHigh 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   6
         ToolTipText     =   "Number of characters to skip from beginning"
         Top             =   1515
         Width           =   1170
      End
      Begin VB.TextBox txtLow 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "Number of characters to skip from beginning"
         Top             =   1140
         Width           =   1170
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
         TabIndex        =   17
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
         Left            =   7725
         TabIndex        =   16
         Top             =   1785
         Width           =   1175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1723 Temperature Device Settings"
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
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
         Left            =   1365
         TabIndex        =   7
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
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
         Left            =   3885
         TabIndex        =   14
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "External Temperature"
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
         Left            =   4050
         TabIndex        =   9
         Top             =   720
         Width           =   1830
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Temperature"
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
         Left            =   1575
         TabIndex        =   2
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
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
         Left            =   4005
         TabIndex        =   10
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
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
         Left            =   1485
         TabIndex        =   3
         Top             =   1200
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmTemperatureDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDevice As cESDevice


Private Sub Form_Unload(Cancel As Integer)

  Set mDevice = Nothing

End Sub



Public Sub Fill()
  If Not (device Is Nothing) Then
      txtLow.text = Format(device.LowSet, "0.000")
      txtLow_A.text = Format(device.LowSet_A, "0.000")
      txtHigh.text = Format(device.HiSet, "0.000")
      txtHigh_A.text = Format(device.HiSet_A, "0.000")
      cboMode.ListIndex = Max(0, Min(2, device.EnableTemperature))
      cboMode_A.ListIndex = Max(0, Min(2, device.EnableTemperature_A))
      

  End If
End Sub

Private Sub cmdOK_Click()
  Dim rc As Boolean
  If cboMode.ListIndex < 0 Then
    cboMode.ListIndex = 0
  End If
  If cboMode_A.ListIndex < 0 Then
    cboMode_A.ListIndex = 0
  End If
  rc = Save()
  
End Sub
Sub SetControls()
  cmdResetSetPoint.Visible = False
  fraEnabler.BackColor = Me.BackColor
  FillCombos



End Sub


Function Save() As Boolean
  

  Dim rs    As Recordset
  Dim SQl   As String
  Dim Count As Long



  On Error GoTo Save_Error

  SQl = "SELECT count(serial) as devcount FROM devices WHERE serial = " & q(device.Serial)
  Set rs = ConnExecute(SQl)
  Count = rs(0)
  rs.Close
  Set rs = Nothing

  If Count > 0 Then
    device.LowSet = Val(txtLow.text)
    device.LowSet_A = Val(txtLow_A.text)
    device.HiSet = Val(txtHigh.text)
    device.HiSet_A = Val(txtHigh_A.text)
    device.EnableTemperature = Max(0, cboMode.ListIndex)
    device.EnableTemperature_A = Max(0, cboMode_A.ListIndex)
    
      If MASTER Then
          If SaveTemperatureDevice(device, gUser.UserName) Then
            Devices.RefreshBySerial device.Serial
            Save = True
          End If
      Else
          Save = (RemoteSaveTemperatureDevice(device) = 0) ' 0 is no errors
          ' check this!!!! Only works with master
          Devices.RefreshBySerial device.Serial
          
      End If
  End If


Save_Resume:
   On Error GoTo 0
   Exit Function

Save_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmTemperatureDevice.Save." & Erl
  Resume Save_Resume

End Function

Private Sub cmdResetSetPoint_Click()
  If MASTER Then
    If Not (device Is Nothing) Then
      device.ReachedSet = 0
      device.ReachedSet_A = 0
    End If
  End If
End Sub

Private Sub Form_Load()
ResetActivityTime
 SetControls
End Sub


Sub FillCombos()
  txtLow.text = 0
  txtLow_A.text = 0
  txtHigh.text = 0
  txtHigh_A.text = 0

  cboMode.Clear
  cboMode_A.Clear
  AddToCombo cboMode, "Disabled", 0
  AddToCombo cboMode, "On Rise", 1
  AddToCombo cboMode, "On Fall", 2
  
  AddToCombo cboMode_A, "Disabled", 0
  AddToCombo cboMode_A, "On Rise", 1
  AddToCombo cboMode_A, "On Fall", 2
  
  cboMode.ListIndex = 0
  cboMode_A.ListIndex = 0
  
End Sub

Private Sub txtHigh_A_GotFocus()
  SelAll txtHigh_A
End Sub

Private Sub txtHigh_A_LostFocus()
  txtHigh_A.text = Format(Val(txtHigh_A.text), "0.000")
End Sub

Private Sub txtHigh_GotFocus()
  SelAll txtHigh
End Sub

Private Sub txtHigh_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtHigh, KeyAscii, True, 3, 8, 999)
End Sub

Private Sub txtHigh_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtHigh_A, KeyAscii, True, 3, 8, 999)
End Sub

Private Sub txtHigh_LostFocus()
  txtHigh.text = Format(Val(txtHigh.text), "0.000")
End Sub

Private Sub txtLow_A_GotFocus()
  SelAll txtLow_A
End Sub

Private Sub txtLow_A_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtLow_A, KeyAscii, True, 3, 8, 999)
End Sub

Private Sub txtLow_A_LostFocus()
  txtLow_A.text = Format(Val(txtLow_A.text), "0.000")
End Sub

Private Sub txtLow_GotFocus()
  SelAll txtLow
End Sub

Private Sub txtLow_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtLow, KeyAscii, True, 3, 8, 999)
  
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

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub



Public Property Get device() As cESDevice

  Set device = mDevice

End Property

Public Property Set device(device As cESDevice)

  Set mDevice = device

End Property

Private Sub txtLow_LostFocus()
  txtLow.text = Format(Val(txtLow.text), "0.000")
End Sub
