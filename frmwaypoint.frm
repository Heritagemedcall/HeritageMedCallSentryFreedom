VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmWaypoint 
   Caption         =   "Waypoint"
   ClientHeight    =   3525
   ClientLeft      =   7080
   ClientTop       =   9930
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   9090
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8925
      Begin VB.Frame fraloc 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2475
         Left            =   75
         TabIndex        =   2
         Top             =   375
         Width           =   7260
         Begin VB.TextBox txtPCA 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2415
            MaxLength       =   8
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1650
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton cmdPCA 
            Caption         =   "Set PCA/Controller"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -1950
            TabIndex        =   22
            Top             =   1695
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6015
            TabIndex        =   26
            Top             =   1650
            Width           =   900
         End
         Begin VB.CommandButton cmdAutoLocate 
            Caption         =   "Locate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4890
            TabIndex        =   25
            Top             =   1650
            Width           =   900
         End
         Begin VB.TextBox txtSurveyDeviceID 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2415
            MaxLength       =   8
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton cmdSetSurveyDevice 
            Caption         =   "Set Survey Device"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -1950
            TabIndex        =   23
            Top             =   2085
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.TextBox txtSignal3 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6345
            MaxLength       =   2
            TabIndex        =   21
            Top             =   1170
            Width           =   390
         End
         Begin VB.TextBox txtSignal2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6345
            MaxLength       =   2
            TabIndex        =   18
            Top             =   810
            Width           =   390
         End
         Begin VB.TextBox txtRepeater3 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   20
            Top             =   1170
            Width           =   1035
         End
         Begin VB.TextBox txtRepeater2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   17
            Top             =   810
            Width           =   1035
         End
         Begin VB.TextBox txtRepeater1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   14
            Top             =   450
            Width           =   1035
         End
         Begin VB.TextBox txtSignal1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6345
            MaxLength       =   2
            TabIndex        =   15
            Top             =   450
            Width           =   390
         End
         Begin VB.TextBox txtDesc 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1125
            MaxLength       =   25
            TabIndex        =   4
            Top             =   150
            Width           =   2550
         End
         Begin VB.TextBox txtWing 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1125
            MaxLength       =   14
            TabIndex        =   10
            Top             =   1230
            Width           =   2550
         End
         Begin VB.TextBox txtFloor 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1125
            MaxLength       =   14
            TabIndex        =   8
            Top             =   870
            Width           =   2550
         End
         Begin VB.TextBox txtBldg 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1125
            MaxLength       =   14
            TabIndex        =   6
            Top             =   510
            Width           =   2550
         End
         Begin VB.Label lblPCA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Survey PCA"
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
            Left            =   1080
            TabIndex        =   34
            Top             =   1718
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Survey Transmitter"
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
            Left            =   540
            TabIndex        =   32
            Top             =   2108
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label lblID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   3780
            TabIndex        =   31
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   75
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Left            =   4710
            TabIndex        =   19
            Top             =   1275
            Width           =   120
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Left            =   4710
            TabIndex        =   16
            Top             =   915
            Width           =   120
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Left            =   4710
            TabIndex        =   13
            Top             =   540
            Width           =   120
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Signal"
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
            Left            =   6285
            TabIndex        =   12
            Top             =   180
            Width           =   540
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repeater"
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
            Left            =   4980
            TabIndex        =   11
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wing"
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
            TabIndex        =   9
            Top             =   1305
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Floor"
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
            Left            =   570
            TabIndex        =   7
            Top             =   900
            Width           =   435
         End
         Begin VB.Label lblDesc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Left            =   30
            TabIndex        =   3
            Top             =   225
            Width           =   975
         End
         Begin VB.Label lblBld 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Building"
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
            Left            =   315
            TabIndex        =   5
            Top             =   540
            Width           =   690
         End
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
         Left            =   7665
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   585
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
         TabIndex        =   29
         Top             =   1755
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
         Left            =   7665
         TabIndex        =   30
         Top             =   2340
         Width           =   1175
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   3015
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5318
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Location"
               Key             =   "loc"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Repeaters"
               Key             =   "data"
               Object.ToolTipText     =   "Assurance Settings"
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
      Begin VB.Label lblDecimal1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11310
         TabIndex        =   28
         Top             =   1395
         Visible         =   0   'False
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmWaypoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Private mLocating As Boolean

Public Sub ProcessPacket(packet As cESPacket)

' at this point if configuration.noncs is true and first hop was NC, then it should not get here

' From ModES.BatchSurvey

  Dim rcx As Long
  Dim Serial As String
  ' this is were we handle events from the system
  Debug.Print "frmWaypint ProcessPacket "; packet.Serial
  
  If SurveyDevice Is Nothing Then Exit Sub

  

  If gDirectedNetwork Then
    If packet.IsLocatorPacket Then
      Serial = packet.LocatedSerial
    Else
      Serial = packet.Serial
    End If
  Else
    Serial = packet.Serial
  End If

  If SurveyDevice.Serial = Serial Then
    If gDirectedNetwork Then
      If Configuration.OnlyLocators Then ' don't need these anymore
        If packet.IsLocatorPacket Then
          SurveyDevice.AddLocater packet
        End If
      Else
        If packet.IsLocatorPacket Then  ' don't need these anymore
          SurveyDevice.AddLocater packet
        ElseIf packet.Alarm0 = 1 Then   ' button pushed
          SurveyDevice.AddLocater packet
        End If
      End If
    ElseIf packet.Alarm0 = 1 Then       ' button pushed non DN
      SurveyDevice.AddLocater packet  ' adds to firsthops of surveydevice
    End If
  End If

  If SurveyDevice.PCASerial = packet.Serial Then   ' only if PCA
    If packet.CMD = &H11 Then
      ' skip it
    Else
      
      rcx = SurveyDevice.ResponseCode(packet, Configuration.surveymode)
      Select Case rcx
        Case SURVEY_RC0  ' nothing... supervisory
          Debug.Print "nothing... supervisory"
        
        Case SURVEY_RC1  ' OK
          Debug.Print "SurveyDevice.ProcessLocations"
          SurveyDevice.ProcessLocations
          ' update screen elements
          txtRepeater1.text = SurveyDevice.Location1
          txtRepeater2.text = SurveyDevice.Location2
          txtRepeater3.text = SurveyDevice.Location3
          txtSignal1.text = SurveyDevice.Signal1
          txtSignal2.text = SurveyDevice.Signal2
          txtSignal3.text = SurveyDevice.Signal3

          dbgloc "Assigning Repeaters and Levels: " & Me.txtDesc.text & vbCrLf

          
          Locating = False
        Case SURVEY_RC2  ' cancel this one
          Debug.Print "Waypoint Update Cancelled"
          dbgloc "Waypoint Update Cancelled: " & Me.txtDesc.text & vbCrLf
          Locating = False
        Case SURVEY_RC3  ' quit
          Debug.Print "Quitting Waypoint Updates"
          dbgloc "Quitting Waypoint Updates: " & Me.txtDesc.text & vbCrLf
          Locating = False
        Case Else
          ' just blow thru

      End Select
    End If
  End If
End Sub





Private Sub cmdAutoLocate_Click()
  If mLocating Then
    Locating = False
  Else
    AutoLocate
  End If

End Sub
Private Sub cmdLocate_Click()
  If mLocating Then
    Locating = False
  Else
    AutoLocate
  End If
End Sub
Sub AutoLocate()

  If Configuration.surveymode = PCA_MODE Then
    Outbounds.AddMessage Configuration.SurveyPCA, MSGTYPE_TWOWAYNID, "", 0
  End If
  dbgloc "FrmWaypoints.Autolocate Count: " & 1  ' how many to do

  Set WayPointForm = Me  ' is this form!

  Locating = True
  SurveyEnabled = True
  NextWayPoint

End Sub

Sub NextWayPoint()
  Dim ID As Long
  Dim NewMessage            As cWirelessMessage
  Dim PagerID As Long

  Sleep 1

  Set SurveyDevice = New cESSurveyDevice  'SurveyDevice is form level object

  If Configuration.surveymode = PCA_MODE Then  'PCA mode

    ' set up serials
    SurveyDevice.Serial = Configuration.SurveyDevice
    SurveyDevice.PCASerial = Configuration.SurveyPCA

    ' make the prompt that shows on the PCA

    SurveyDevice.RequestSurvey Me.BuildPrompt(), 2

    Set NewMessage = New cWirelessMessage
    NewMessage.MessageData = SurveyDevice.MessageString
    NewMessage.TimeStamp = DateAdd("s", -1, Now)  ' send right away, no delay
    NewMessage.NeedsAck = SurveyDevice.RequireACK  ' (Val("&h" & MID(mMessageData, 17, 2)) And (BIT_7)) = (BIT_7)
    NewMessage.SequenceID = SurveyDevice.Sequence  ' Val("&h" & MID(mMessageData, 21, 4))

    'Send it to the PCA
    Outbounds.AddPreparedMessage NewMessage
    Set NewMessage = Nothing

    'dbg "Sent PCA Message, Get " & s & vbCrLf
  ElseIf Configuration.surveymode = EN1221_MODE Then
  
    PagerID = Configuration.SurveyPager
  
    If PagerID <> 0 Then
      SurveyDevice.PagerID = PagerID
      SurveyDevice.Serial = Configuration.SurveyDevice  ' use just one device
      SurveyDevice.PCASerial = Configuration.SurveyDevice
      ' create outbound message for ouput "Pager"
      SurveyDevice.RequestSurvey Me.BuildPrompt(), 2, True  ' only OK, cancel
      SendToPager SurveyDevice.MessageString, SurveyDevice.PagerID, 0, "", "", PAGER_NORMAL, "", 0, 0
    End If
  Else  ' TWO Button Pendant Mode
    
    PagerID = Configuration.SurveyPager
    'PagerID = Configuration.SurveyPager
    If PagerID <> 0 Then
      SurveyDevice.PagerID = PagerID
      SurveyDevice.Serial = Configuration.SurveyDevice  ' use just one device
      SurveyDevice.PCASerial = Configuration.SurveyDevice
      ' create outbound message for ouput "Pager"
      SurveyDevice.RequestSurvey Me.BuildPrompt(), 2, True  ' only OK, cancel
      SendToPager SurveyDevice.MessageString, SurveyDevice.PagerID, 0, "", "", PAGER_NORMAL, "", 0, 0
    End If

  End If
  
  If SurveyDevice Is Nothing Then ' just a sanity check
    Locating = False
  End If
  
End Sub

Public Function BuildPrompt() As String
  BuildPrompt = Trim(txtDesc.text) & " " & Trim(txtBldg.text) & " " & Trim(Me.txtFloor.text) & " " & Trim(Me.txtWing.text)
End Function


Private Sub cmdCancel_Click()
  SurveyEnabled = False
  PreviousForm
  Unload Me
End Sub

Private Sub cmdClear_Click()
  txtRepeater1.text = ""
  txtRepeater2.text = ""
  txtRepeater3.text = ""
  txtSignal1.text = "0"
  txtSignal2.text = "0"
  txtSignal3.text = "0"

End Sub

Private Sub cmdNew_Click()
  ID = 0
  Fill
End Sub

Private Sub cmdOK_Click()
  If Save() Then
    PreviousForm
    Unload Me
  End If

End Sub

Private Function Save() As Boolean


  Dim Rs As Recordset

  If ID = 0 Then

    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * FROM waypoints WHERE ID = 0 ", conn, gCursorType, gLockType
    Rs.addnew
    Rs("description") = Trim(txtDesc.text)
    Rs("building") = Trim(txtBldg.text)
    Rs("floor") = Trim(txtFloor.text)
    Rs("wing") = Trim(txtWing.text)
    Rs("repeater1") = UCase(Trim(txtRepeater1.text))
    Rs("repeater2") = UCase(Trim(txtRepeater2.text))
    Rs("repeater3") = UCase(Trim(txtRepeater3.text))
    Rs("Signal1") = Val(txtSignal1.text)
    Rs("Signal2") = Val(txtSignal2.text)
    Rs("Signal3") = Val(txtSignal3.text)


    Rs.Update
    Rs.MoveLast
    ID = Rs("id")
    Rs.Close
    Set Rs = Nothing
  Else

    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * FROM waypoints WHERE ID = " & ID, conn, gCursorType, gLockType
    Rs("description") = Trim(txtDesc.text)
    Rs("building") = Trim(txtBldg.text)
    Rs("floor") = Trim(txtFloor.text)
    Rs("wing") = Trim(txtWing.text)
    Rs("repeater1") = UCase(Trim(txtRepeater1.text))
    Rs("repeater2") = UCase(Trim(txtRepeater2.text))
    Rs("repeater3") = UCase(Trim(txtRepeater3.text))
    Rs("Signal1") = Val(txtSignal1.text)
    Rs("Signal2") = Val(txtSignal2.text)
    Rs("Signal3") = Val(txtSignal3.text)
    Rs.Update
    Rs.MoveLast
    ID = Rs("id")
    Rs.Close
    Set Rs = Nothing

  End If

  Dim w As cWayPoint
  Dim j As Integer
  For j = Waypoints.Count To 1 Step -1
    Set w = Waypoints.waypoint(j)
    If w.ID = ID Then
      Exit For
    End If
  Next
  If j = 0 Then
    Set w = New cWayPoint
    w.ID = ID
  End If
  w.Description = Trim(txtDesc.text)
  w.Building = Trim(txtBldg.text)
  w.Floor = Trim(txtFloor.text)
  w.Wing = Trim(txtWing.text)
  w.Repeater1 = Trim(txtRepeater1.text)
  w.Repeater2 = Trim(txtRepeater2.text)
  w.Repeater3 = Trim(txtRepeater3.text)
  w.Signal1 = Val(txtSignal1.text)
  w.Signal2 = Val(txtSignal2.text)
  w.Signal3 = Val(txtSignal3.text)



End Function



Private Sub cmdPCA_Click()
  'SetDevices
End Sub

Private Sub cmdSetSurveyDevice_Click()
  'SetDevices
End Sub


Private Sub Form_Load()
  ResetActivityTime
  SetControls
  Set WayPointForm = Me

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
  Dim Rs As ADODB.Recordset
  Dim SQL As String
  Dim NewID As Long

  SQL = "SELECT * FROM Waypoints WHERE ID = " & ID
  Set Rs = ConnExecute(SQL)

  If Rs.EOF Then

    txtDesc.text = ""
    txtBldg.text = ""
    txtFloor.text = ""
    txtWing.text = ""
    txtRepeater1.text = ""
    txtRepeater2.text = ""
    txtRepeater3.text = ""
    txtSignal1.text = 0
    txtSignal2.text = 0
    txtSignal3.text = 0
  Else
    txtDesc.text = Rs("description") & ""
    txtBldg.text = Rs("building") & ""
    txtFloor.text = Rs("floor") & ""
    txtWing.text = Rs("wing") & ""
    txtRepeater1.text = Rs("repeater1") & ""
    txtRepeater2.text = Rs("repeater2") & ""
    txtRepeater3.text = Rs("repeater3") & ""
    txtSignal1.text = Rs("Signal1")
    txtSignal2.text = Rs("Signal2")
    txtSignal3.text = Rs("Signal3")
    NewID = Rs("ID")
  End If
  Rs.Close
  Set Rs = Nothing
  ID = NewID
  If ID <> 0 Then
    lblID.Caption = ID
  Else
    lblID.Caption = ""
  End If
  txtSurveyDeviceID.text = Configuration.SurveyDevice
  txtPCA.text = Configuration.SurveyPCA
  cmdAutoLocate.Caption = IIf(SurveyEnabled, "Cancel", "Locate")

End Sub
Sub SetControls()
  fraEnabler.BackColor = Me.BackColor
  fraloc.BackColor = Me.BackColor
  txtPCA.text = Configuration.SurveyPCA
  txtSurveyDeviceID.text = Configuration.SurveyDevice
  If Configuration.surveymode = PCA_MODE Then
    txtPCA.Visible = True
    lblPCA.Visible = True
  Else
    txtPCA.Visible = False
    lblPCA.Visible = False
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set WayPointForm = Nothing
  Set SurveyDevice = Nothing
  SurveyEnabled = False
  Locating = False

End Sub

Private Sub txtPCA_GotFocus()
  SelAll txtPCA
End Sub

Private Sub txtPCA_KeyPress(KeyAscii As Integer)
' handle hex data only
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
  'KeyAscii = KeyProcHex(txtSerial, KeyAscii, False, 0, 8)

End Sub

Private Sub txtRepeater1_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtRepeater1_LostFocus()
  txtRepeater1.text = UCase(txtRepeater1.text)
End Sub

Private Sub txtRepeater2_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtRepeater2_LostFocus()
  txtRepeater2.text = UCase(txtRepeater2.text)
End Sub

Private Sub txtRepeater3_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtRepeater3_LostFocus()
  txtRepeater3.text = UCase(txtRepeater3.text)
End Sub

Private Sub txtSignal1_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)

    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtSignal2_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)

    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select

End Sub

Private Sub txtSignal3_KeyPress(KeyAscii As Integer)
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)

    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select

End Sub

Private Sub txtSurveyDeviceID_GotFocus()
  SelAll txtSurveyDeviceID
End Sub

Private Sub txtSurveyDeviceID_KeyPress(KeyAscii As Integer)
' handle hex data only
  KeyAscii = ToUpper(KeyAscii)
  Select Case Chr(KeyAscii)
    Case "A" To "F"
    Case "1" To "9"
    Case "0"
    Case Chr(8)
    Case Else
      KeyAscii = 0
  End Select
  'KeyAscii = KeyProcHex(txtSerial, KeyAscii, False, 0, 8)
End Sub

Public Property Get Locating() As Boolean
  Locating = mLocating
End Property

Public Property Let Locating(ByVal Locating As Boolean)
  mLocating = Locating
  cmdAutoLocate.Caption = IIf(mLocating, "Cancel", "Locate")
  If mLocating = False Then
    SurveyEnabled = False
  End If
End Property
