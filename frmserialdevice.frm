VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSerialDevice 
   Caption         =   "Serial Input Setup"
   ClientHeight    =   3090
   ClientLeft      =   5505
   ClientTop       =   10440
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Frame fraGeneral 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   15
         TabIndex        =   15
         Top             =   330
         Width           =   7605
         Begin VB.CheckBox chkPET 
            Caption         =   "Verbose"
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
            Left            =   5220
            TabIndex        =   27
            Top             =   465
            Width           =   1485
         End
         Begin VB.CheckBox chkSerialTapProtocol 
            Caption         =   "TAP Protocol"
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
            Left            =   3720
            TabIndex        =   26
            Top             =   465
            Width           =   1485
         End
         Begin VB.TextBox txtPream 
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
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   28
            ToolTipText     =   "Number of characters to skip from beginning"
            Top             =   810
            Width           =   5970
         End
         Begin VB.TextBox txtExcludeWords 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1560
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   32
            ToolTipText     =   "Separate Words with Spaces, Case sensitive"
            Top             =   1845
            Width           =   5970
         End
         Begin VB.TextBox txtEOLChar 
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
            Left            =   6150
            MaxLength       =   3
            TabIndex        =   21
            ToolTipText     =   "ASCII code of End-of-Line character"
            Top             =   105
            Width           =   630
         End
         Begin VB.TextBox txtIncludeWords 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1560
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   30
            ToolTipText     =   "Separate Words with Spaces, Case sensitive"
            Top             =   1170
            Width           =   5970
         End
         Begin VB.TextBox txtAutoClear 
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
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   23
            ToolTipText     =   "Enter 1 to 999. Zero Disables"
            Top             =   465
            Width           =   630
         End
         Begin VB.TextBox txtSkip 
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
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   17
            ToolTipText     =   "Number of characters to skip from beginning"
            Top             =   105
            Width           =   630
         End
         Begin VB.TextBox txtLength 
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
            Left            =   4050
            MaxLength       =   3
            TabIndex        =   19
            ToolTipText     =   "Number of characters to include in output"
            Top             =   105
            Width           =   630
         End
         Begin VB.Label lblPream 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preamble"
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
            Left            =   615
            TabIndex        =   25
            Top             =   870
            Width           =   795
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disqualifiers"
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
            Left            =   360
            TabIndex        =   31
            Top             =   1920
            Width           =   1050
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special EOL"
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
            Left            =   4890
            TabIndex        =   20
            Top             =   165
            Width           =   1065
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minutes"
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
            Left            =   2340
            TabIndex        =   24
            Top             =   525
            Width           =   675
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Auto Clear"
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
            Left            =   510
            TabIndex        =   22
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lbl3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trigger Words"
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
            Left            =   195
            TabIndex        =   29
            Top             =   1245
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message Start"
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
            TabIndex        =   16
            Top             =   165
            Width           =   1230
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message Length"
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
            Left            =   2445
            TabIndex        =   18
            Top             =   165
            Width           =   1410
         End
      End
      Begin VB.Frame fraPort 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2445
         Left            =   60
         TabIndex        =   2
         Top             =   405
         Width           =   5490
         Begin VB.ComboBox cboBaud 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   465
            Width           =   2340
         End
         Begin VB.ComboBox cboBits 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1245
            Width           =   2340
         End
         Begin VB.ComboBox cboStop 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1635
            Width           =   2340
         End
         Begin VB.ComboBox cboFlow 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2025
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox cboPort 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   75
            Width           =   2340
         End
         Begin VB.ComboBox cboParity 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   855
            Width           =   2340
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Comm Port"
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
            Left            =   600
            TabIndex        =   3
            Top             =   150
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Data bits"
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
            Left            =   735
            TabIndex        =   9
            Top             =   1305
            Width           =   780
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Stop bits"
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
            Left            =   750
            TabIndex        =   11
            Top             =   1695
            Width           =   765
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Flow Control"
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
            Left            =   450
            TabIndex        =   13
            Top             =   2085
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Bits per second"
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
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Parity"
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
            Left            =   1020
            TabIndex        =   7
            Top             =   915
            Width           =   495
         End
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
         TabIndex        =   33
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
         TabIndex        =   34
         Top             =   2370
         Width           =   1175
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2970
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   5239
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "General"
               Object.Tag             =   "general"
               Object.ToolTipText     =   "General Settings"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Port Setup"
               Key             =   "port"
               Object.ToolTipText     =   "Port and data format setup"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSerialDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Serial As String
Private rs As Recordset
Private mDevice As cESDevice

Const MAXSERIALMESSAGELENGTH = 80 ' was 40 ' changed per Jerry 2/22/13

Private Sub Form_Unload(Cancel As Integer)

  Set mDevice = Nothing

End Sub


Public Sub Fill()
  'Dim rs As Recordset
  Dim protocol As Integer
  ResetForm

  'Set rs = connexecute("SELECT * FROM devices WHERE serial = " & Q(Serial))

  If Not device Is Nothing Then
 
    Debug.Print device.Serial
    txtSkip.text = device.SerialSkip
    txtLength.text = device.SerialMessageLen
    txtAutoClear.text = device.SerialAutoClear
    txtIncludeWords.text = device.SerialInclude
    txtExcludeWords.text = device.SerialExclude
    cboPort.ListIndex = Max(0, CboGetIndexByItemData(cboPort, device.SerialPort))
    cboBaud.ListIndex = Max(0, CboGetIndexByItemData(cboBaud, device.SerialBaud))
    cboParity.ListIndex = Max(0, CboGetIndexByItemData(cboParity, GetParityID(device.SerialParity)))
    cboBits.ListIndex = Max(0, CboGetIndexByItemData(cboBits, device.Serialbits))
    cboStop.ListIndex = GetComboByText(cboStop, device.SerialStopbits)
    If cboStop.ListIndex < 0 Then
      cboStop.ListIndex = 0
    End If
    chkSerialTapProtocol.Value = IIf(device.SerialTapProtocol = 0, 0, 1)
    chkPET.Value = IIf(device.SerialTapProtocol = 2, 1, 0)
    chkPET.Enabled = chkSerialTapProtocol.Value = 1
    cboFlow.ListIndex = Max(0, CboGetIndexByItemData(cboFlow, device.SerialFlow))
    txtEOLChar.text = device.SerialEOLChar
    txtPream.text = device.SerialPreamble
    
  End If



End Sub
Sub ResetForm()
    
  txtSkip.text = "0"
  txtLength.text = "0"
  txtAutoClear.text = "0"
  txtIncludeWords.text = ""
  txtExcludeWords.text = ""
  cboPort.ListIndex = 0
  cboBaud.ListIndex = 6
  cboBits.ListIndex = cboBits.listcount - 1
  cboParity.ListIndex = 2
  cboStop.ListIndex = 0
  cboFlow.ListIndex = 0
  txtEOLChar.text = 0
  
  chkPET.Value = 0
  chkSerialTapProtocol.Value = 0
  
  
End Sub

Function Save() As Boolean
  
  
        Dim rs    As Recordset
        Dim SQl   As String
        Dim Count As Long
  
  
        'Set rs = New ADODB.Recordset
        'rs.Open "SELECT count(serial) FROM devices WHERE serial = " & Q(Serial), conn, gCursorType, gLockType
  
10       On Error GoTo Save_Error

20      SQl = "SELECT count(serial) as devcount FROM devices WHERE serial = " & q(device.Serial)
30      Set rs = ConnExecute(SQl)
40      Count = rs(0)
50      rs.Close
60      Set rs = Nothing
  
70      If Count > 0 Then
            If chkSerialTapProtocol.Value = 1 Then
              device.SerialTapProtocol = chkPET.Value + 1
            Else
              device.SerialTapProtocol = 0
            End If
           
80          device.SerialSkip = Min(Val(txtSkip.text), 250)
90          device.SerialMessageLen = Min(Val(txtLength.text), MAXSERIALMESSAGELENGTH)
100         device.SerialAutoClear = Min(Val(txtAutoClear.text), 999)
110         device.SerialInclude = Trim(txtIncludeWords.text)
120         device.SerialExclude = Trim(txtExcludeWords.text)
130         device.SerialPort = GetComboItemData(cboPort)
140         device.SerialBaud = GetComboItemData(cboBaud)
150         device.SerialParity = GetParityString(GetComboItemData(Me.cboParity))
160         device.Serialbits = Val(cboBits.text)
170         device.SerialFlow = GetComboItemData(cboFlow)
180         device.SerialStopbits = cboStop.text
190         device.SerialSettings = device.SerialBaud & device.SerialParity & device.Serialbits & device.SerialStopbits
200         device.SerialEOLChar = Min(255, Val(txtEOLChar.text))
210         device.SerialPreamble = left(txtPream.text, 100)
220         If MASTER Then
230             If SaveSerialDevice(device, gUser.UserName) Then
240               Devices.RefreshBySerial device.Serial
250               SetupSerialDevice device
260               Save = True
270             End If
280         Else
290             Save = (RemoteSaveSerialDevice(device) = 0) ' 0 is no errors
                'Devices.RefreshBySerial Device.Serial
                'SetupSerialDevice Device
300         End If
310     End If
      '  Set rs = New ADODB.Recordset
      '  rs.Open "SELECT * FROM devices WHERE serial = " & Q(Serial), conn, gCursorType, gLockType
      '  If Not rs.EOF Then
      '      Device.SerialSkip = Min(Val(txtSkip.text), 250)
      '      Device.SerialMessageLen = Min(Val(txtLength.text), 40)
      '      Device.SerialAutoClear = Min(Val(txtAutoClear.text), 999)
      '      Device.SerialInclude = Trim(txtIncludeWords.text)
      '      Device.SerialExclude = Trim(txtExcludeWords.text)
      '      Device.SerialPort = GetComboItemData(cboPort)
      '      Device.serialbaud = GetComboItemData(cboBaud)
      '      Device.SerialParity = GetParityString(GetComboItemData(Me.cboParity))
      '      Device.Serialbits = Val(cboBits.text)
      '      Device.SerialFlow = GetComboItemData(cboFlow)
      '      Device.SerialStopbits = cboStop.text
      '      Device.SerialSettings = rs("SerialBaud") & rs("SerialParity") & rs("SerialBits") & rs("SerialStopBits")
      '      Device.SerialEOLChar = Min(255, Val(txtEOLChar.text))
      '      Device.SerialPreamble = left(txtPream.text, 100)
      '
      '
      '      rs("SerialSkip") = Device.SerialSkip
      '      rs("SerialMessageLen") = Device.SerialMessageLen
      '      rs("SerialAutoClear") = Device.SerialAutoClear
      '      rs("SerialInclude") = left(Device.SerialInclude, 255)
      '      rs("SerialExclude") = left(Device.SerialExclude, 255)
      '      rs("SerialPort") = Device.SerialPort
      '      rs("SerialBaud") = Device.serialbaud
      '      rs("SerialParity") = Device.SerialParity
      '      rs("SerialBits") = Device.Serialbits
      '      rs("SerialFlow") = Device.SerialFlow
      '      rs("SerialStopBits") = Device.SerialStopbits
      '      rs("SerialSettings") = Device.SerialSettings
      '      rs("SerialEOLChar") = Device.SerialEOLChar
      '      rs("SerialPreamble") = Device.SerialPreamble
      '      rs.Update
      '      Devices.RefreshBySerial Device.Serial
      '      SetupSerialDevice Device
      '   End If
      '   rs.Close
      '   Set rs = Nothing

Save_Resume:
320      On Error GoTo 0
330      Exit Function

Save_Error:

340     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmSerialDevice.Save." & Erl
350     Resume Save_Resume

End Function

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

Private Sub chkSerialTapProtocol_Click()
  chkPET.Enabled = chkSerialTapProtocol.Value = 1
  
End Sub

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Save
End Sub

Private Sub Form_Load()
ResetActivityTime
  SetControls
  ShowPanel TabStrip.SelectedItem.Key
End Sub
Sub SetControls()
  fraEnabler.BackColor = Me.BackColor
  
  fraGeneral.left = TabStrip.ClientLeft
  fraGeneral.top = TabStrip.ClientTop
  fraGeneral.Height = TabStrip.ClientHeight
  fraGeneral.Width = TabStrip.ClientWidth
  
  fraGeneral.BackColor = Me.BackColor
  fraPort.left = TabStrip.ClientLeft
  fraPort.top = TabStrip.ClientTop
  fraPort.Height = TabStrip.ClientHeight
  fraPort.Width = TabStrip.ClientWidth
  fraPort.BackColor = Me.BackColor
  FillCombos



End Sub
Sub FillCombos()

  Dim j As Integer

  cboPort.Clear
  AddToCombo cboPort, "None", 0
  For j = 1 To 256
    AddToCombo cboPort, "COM " & j, j
  Next
  cboPort.ListIndex = 0

  cboBits.Clear
  For j = 4 To 8
    AddToCombo cboBits, j, j
  Next
  cboBits.ListIndex = cboBits.listcount - 1



  cboBaud.Clear
  'AddToCombo cboBaud, "75", 75
  'AddToCombo cboBaud, "110", 110
  'AddToCombo cboBaud, "150", 150
  AddToCombo cboBaud, "300", 300
  AddToCombo cboBaud, "600", 600
  AddToCombo cboBaud, "1200", 1200
  AddToCombo cboBaud, "2400", 2400
  AddToCombo cboBaud, "4800", 4800
  AddToCombo cboBaud, "7200", 7200
  AddToCombo cboBaud, "9600", 9600
  AddToCombo cboBaud, "14400", 14400
  AddToCombo cboBaud, "19200", 19200
  AddToCombo cboBaud, "38400", 38400
  AddToCombo cboBaud, "57600", 57600
  AddToCombo cboBaud, "115200", 115200
  AddToCombo cboBaud, "128000", 128000

  cboBaud.ListIndex = 6

  cboParity.Clear
  AddToCombo cboParity, "Even", 0
  AddToCombo cboParity, "Odd", 1
  AddToCombo cboParity, "None", 2
  AddToCombo cboParity, "Mark", 3
  AddToCombo cboParity, "Space", 4

  cboParity.ListIndex = 2

  cboStop.Clear
  AddToCombo cboStop, 1, 10
  AddToCombo cboStop, 1.5, 15
  AddToCombo cboStop, 2, 20
  cboStop.ListIndex = 0

  cboFlow.Clear
  AddToCombo cboFlow, "None", 0
  AddToCombo cboFlow, "Hardware", 1
  AddToCombo cboFlow, "Xon/Xoff", 2

  cboFlow.ListIndex = 0


End Sub

Sub ShowPanel(ByVal Key As String)
  Select Case LCase(Key)
    Case "port"
      fraPort.Visible = True
      fraGeneral.Visible = False
    Case Else ' general
      fraGeneral.Visible = True
      fraPort.Visible = False
  End Select
  
End Sub

Private Sub TabStrip_Click()
  ShowPanel TabStrip.SelectedItem.Key
End Sub

Private Sub txtAutoClear_GotFocus()
  SelAll txtAutoClear
End Sub

Private Sub txtAutoClear_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtAutoClear, KeyAscii, False, 0, 3, 999)
End Sub

Private Sub txtEOLChar_GotFocus()
  SelAll txtEOLChar
End Sub

Private Sub txtEOLChar_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtEOLChar, KeyAscii, False, 0, 3, 255)
End Sub

Private Sub txtLength_GotFocus()
  SelAll txtLength
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtLength, KeyAscii, False, 0, 3, 255)
End Sub

Private Sub txtSkip_GotFocus()
  SelAll txtSkip
End Sub

Private Sub txtSkip_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtSkip, KeyAscii, False, 0, 3, 255)
End Sub

Public Property Get device() As cESDevice

  Set device = mDevice

End Property

Public Property Set device(Value As cESDevice)

  Set mDevice = Value

End Property
