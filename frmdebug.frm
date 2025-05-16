VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug"
   ClientHeight    =   3375
   ClientLeft      =   1380
   ClientTop       =   10560
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   9240
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3060
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.CheckBox chkExtendFactory 
         Caption         =   "Extended Factory Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4260
         TabIndex        =   18
         ToolTipText     =   "Check to Log TAP Pager Protocol"
         Top             =   1650
         Width           =   2685
      End
      Begin VB.CheckBox chkLogTAP 
         Caption         =   "LOG TAP Pager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4275
         TabIndex        =   9
         ToolTipText     =   "Check to Log TAP Pager Protocol"
         Top             =   1305
         Width           =   2670
      End
      Begin VB.Frame fra6080export 
         BorderStyle     =   0  'None
         Caption         =   "6080 Export"
         Height          =   795
         Left            =   960
         TabIndex        =   11
         Top             =   1950
         Width           =   5925
         Begin VB.CommandButton cmdConfig 
            Caption         =   "Configuration"
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
            Left            =   4230
            TabIndex        =   16
            Top             =   420
            Width           =   1350
         End
         Begin VB.CommandButton cmdExportZones 
            Caption         =   "Zones"
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
            Left            =   0
            TabIndex        =   13
            Top             =   420
            Width           =   1350
         End
         Begin VB.CommandButton cmdExportPartitions 
            Caption         =   "Partitons"
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
            Left            =   1410
            TabIndex        =   14
            Top             =   420
            Width           =   1350
         End
         Begin VB.CommandButton cmdExportSoftPoints 
            Caption         =   "Soft Points"
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
            Left            =   2820
            TabIndex        =   15
            Top             =   420
            Width           =   1350
         End
         Begin VB.Label lbl6080Exp 
            AutoSize        =   -1  'True
            Caption         =   "6080 Export"
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
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdApplyBatt 
         Caption         =   "Set"
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
         Left            =   5880
         TabIndex        =   8
         Top             =   930
         Width           =   1175
      End
      Begin VB.TextBox txtLowBattDelay 
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
         Height          =   300
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Minutes to Verify Low Battery"
         Top             =   930
         Width           =   585
      End
      Begin VB.CheckBox chkNoDataErrorLog 
         Caption         =   "No Data Error Log"
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
         Left            =   4260
         TabIndex        =   5
         Top             =   510
         Width           =   2685
      End
      Begin VB.CheckBox chkNoStrayData 
         Caption         =   "No Stray Data"
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
         Left            =   4260
         TabIndex        =   4
         Top             =   180
         Width           =   2685
      End
      Begin VB.CheckBox chkShowPacketData 
         Caption         =   "Show Packet Data *"
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
         Left            =   780
         TabIndex        =   3
         Top             =   810
         Width           =   2685
      End
      Begin VB.CheckBox chkTAPIData 
         Caption         =   "Show TAPI Data *"
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
         Left            =   780
         TabIndex        =   2
         Top             =   480
         Width           =   2685
      End
      Begin VB.CheckBox chkLocationData 
         Caption         =   "Show Location Data *"
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
         Left            =   780
         TabIndex        =   1
         Top             =   150
         Width           =   2685
      End
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
         TabIndex        =   17
         Top             =   2370
         Width           =   1175
      End
      Begin VB.Label lblLowBattDelay 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low Batt Delay"
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
         Left            =   3660
         TabIndex        =   6
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label lblasterisk 
         BackStyle       =   0  'Transparent
         Caption         =   "* Clears on Reboot"
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
         Left            =   990
         TabIndex        =   10
         Top             =   1230
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents HTTPRequest As cHTTPRequest
Attribute HTTPRequest.VB_VarHelpID = -1
Dim ActiveRequest As String


Private Sub chkExtendFactory_Click()
  gExtendFactory = IIf(Me.chkExtendFactory.Value = 1, True, False)
  WriteSetting "Debug", "ExtendFactory", gExtendFactory
End Sub

Private Sub chkLocationData_Click()
  gShowLocationData = IIf(chkLocationData.Value = 1, True, False)
End Sub

Private Sub chkLogTAP_Click()
  gLogTAP = IIf(chkLogTAP.Value = 1, True, False)
  WriteSetting "Debug", "LogTap", gLogTAP
End Sub

Private Sub chkNoDataErrorLog_Click()
  gNoDataErrorLog = IIf(chkNoDataErrorLog.Value = 1, True, False)
  WriteSetting "Debug", "NoDataErrorLog", gNoDataErrorLog

End Sub

Private Sub chkNoStrayData_Click()
  gNoStrayData = IIf(chkNoStrayData.Value = 1, True, False)
  WriteSetting "Debug", "NoStrayData", gNoStrayData
End Sub

Private Sub chkShowPacketData_Click()
  gShowPacketData = IIf(chkShowPacketData.Value = 1, True, False)
End Sub

Private Sub chkTAPIData_Click()
  gShowTAPIData = IIf(chkTAPIData.Value = 1, True, False)
End Sub

Private Sub cmdApplyBatt_Click()
  
  gLoBattDelay = Val(Me.txtLowBattDelay.text)
  WriteSetting "Configuration", "LowBattDelay", gLoBattDelay
  
End Sub

Private Sub cmdConfig_Click()
  cmdConfig.Caption = "Busy"
  ExportConfiguration
  cmdConfig.Caption = "Configuration"

End Sub

Sub ExportConfiguration()
        Dim HTTPRequest        As cHTTPRequest
        Dim hfile              As Long
        Dim exportfilename     As String
        Dim Configdata()         As Byte



10      On Error GoTo ExportConfiguration_Error

20      fraEnabler.Enabled = False

30      exportfilename = App.path & "\" & Format$(Now, "yyyy-mm-ddThh-nn-ss") & ".cfg"

40      Set HTTPRequest = New cHTTPRequest

50      Call HTTPRequest.Backup6080(exportfilename, GetHTTP & "://" & IP1, USER1, PW1)

ExportConfiguration_Resume:
120     On Error Resume Next
'130     Close hfile
140     Set HTTPRequest = Nothing
150     fraEnabler.Enabled = True

160     On Error GoTo 0
170     Exit Sub

ExportConfiguration_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at frmDebug.ExportConfiguration." & Erl
190     Resume ExportConfiguration_Resume
End Sub

Private Sub cmdExit_Click()
 PreviousForm
 Unload Me
End Sub

Private Sub cmdExportPartitions_Click()
  cmdExportPartitions.Caption = "Busy"
  ExportPartitions
  cmdExportPartitions.Caption = "Partitions"
End Sub

Private Sub cmdExportSoftPoints_Click()
  cmdExportSoftPoints.Caption = "Busy"
  ExportSoftPoints
  cmdExportSoftPoints.Caption = "Soft Points"
End Sub

Private Sub cmdExportZones_Click()
  cmdExportZones.Caption = "Busy"
  
  ExportZones
  cmdExportZones.Caption = "Zones"
End Sub

Sub ExportZones()
  
  Dim rc As Long
  Dim XML   As String
  
  fraEnabler.Enabled = False
  
    
  ActiveRequest = "ZoneList"
  Set HTTPRequest = New cHTTPRequest
    
    Call HTTPRequest.GetZoneList(GetHTTP & "://" & IP1, USER1, PW1)
  Do Until HTTPRequest.Ready
    DoEvents
  Loop

    Select Case HTTPRequest.StatusCode
    Case 200, 201
      rc = 1
    Case Else
      rc = 0
  End Select
  If rc Then
    XML = HTTPRequest.XML
    If Len(XML) Then
      LogXML XML, "Export-Zones"
    End If
  End If
      
  fraEnabler.Enabled = True

  Set HTTPRequest = Nothing
  
End Sub

Sub ExportPartitions()
  
  Dim rc As Long
  Dim XML   As String
  
  fraEnabler.Enabled = False
  
    
  ActiveRequest = "PartitionList"
  Set HTTPRequest = New cHTTPRequest
    
  Call HTTPRequest.GetPartitionList(GetHTTP & "://" & IP1, USER1, PW1)
  Do Until HTTPRequest.Ready
    DoEvents
  Loop

    Select Case HTTPRequest.StatusCode
    Case 200, 201
      rc = 1
    Case Else
      rc = 0
  End Select
  If rc Then
    XML = HTTPRequest.XML
    If Len(XML) Then
      LogXML XML, "Export-Partitions"
    End If
  End If
      
  fraEnabler.Enabled = True

  Set HTTPRequest = Nothing
  
  End Sub

Sub ExportSoftPoints()
  
  Dim rc As Long
  Dim XML   As String
  
  fraEnabler.Enabled = False
  
    
  ActiveRequest = "SoftPointList"
  Set HTTPRequest = New cHTTPRequest
    
  Call HTTPRequest.GetSoftPointList(GetHTTP & "://" & IP1, USER1, PW1)
  Do Until HTTPRequest.Ready
   DoEvents
  Loop

    Select Case HTTPRequest.StatusCode
    Case 200, 201
      rc = 1
    Case Else
      rc = 0
  End Select
  If rc Then
    XML = HTTPRequest.XML
    If Len(XML) Then
      LogXML XML, "Export-SoftPoints"
    End If
  End If
      
  Set HTTPRequest = Nothing
  fraEnabler.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
  Fill
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
  fraEnabler.BackColor = Me.BackColor
End Sub

Public Sub Fill()

  If USE6080 Then
    lbl6080Exp.Visible = True
    cmdExportZones.Visible = True
    cmdExportPartitions.Visible = True
    cmdExportSoftPoints.Visible = True
    cmdConfig.Visible = True
  Else
    lbl6080Exp.Visible = False
    cmdExportZones.Visible = False
    cmdExportPartitions.Visible = False
    cmdExportSoftPoints.Visible = False
    cmdConfig.Visible = False
  End If
  chkLogTAP.Value = IIf(gLogTAP, 1, 0)
  chkNoDataErrorLog.Value = IIf(gNoDataErrorLog, 1, 0)
  chkNoStrayData.Value = IIf(gNoStrayData, 1, 0)
  chkLocationData.Value = IIf(gShowLocationData, 1, 0)
  chkTAPIData.Value = IIf(gShowTAPIData, 1, 0)
  chkShowPacketData = IIf(gShowPacketData, 1, 0)
  txtLowBattDelay.text = gLoBattDelay
  
  If gUser.LEvel >= LEVEL_FACTORY Then
    chkExtendFactory.Enabled = True
  Else
    chkExtendFactory.Enabled = False
  End If
  chkExtendFactory.Value = IIf(gExtendFactory, 1, 0)
  
  
End Sub

Private Sub HTTPrequest_Done()
  
  Select Case ActiveRequest
    Case "ZoneList"
    Case "PartitionList"
    Case "SoftPoints"
    Case Else
  End Select
    
  Me.fraEnabler.Enabled = True
End Sub

Private Sub txtLowBattDelay_GotFocus()
  SelAll txtLowBattDelay
End Sub

Private Sub txtLowBattDelay_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyProcMax(txtLowBattDelay, KeyAscii, False, 0, 3, 999)
End Sub
