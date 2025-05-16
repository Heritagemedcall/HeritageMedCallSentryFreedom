VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmUpgrade 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   750
   ClientTop       =   5250
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   9060
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
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
         Height          =   525
         Left            =   7680
         TabIndex        =   2
         Top             =   2460
         Width           =   1175
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Upgrade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7680
         TabIndex        =   1
         Top             =   720
         Width           =   1175
      End
      Begin MSComctlLib.ListView lvNonUpgrade 
         Height          =   2760
         Left            =   2535
         TabIndex        =   3
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4868
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serial"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Model"
            Object.Width           =   1834
         EndProperty
      End
      Begin MSComctlLib.ListView lvFailed 
         Height          =   2760
         Left            =   5070
         TabIndex        =   4
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4868
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serial"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Model"
            Object.Width           =   1834
         EndProperty
      End
      Begin MSComctlLib.ListView lvUpgraded 
         Height          =   2760
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4868
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serial"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Model"
            Object.Width           =   1834
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Failed Upgrading"
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
         Left            =   5130
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Not Upgraded"
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
         Left            =   2595
         TabIndex        =   6
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Upgraded"
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
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UpgradeDevices As Collection
' WE can only upgrade certain devicetypes
'EN1210
'EN1210'EOL
'EN1210W
'EN1212
'EN1215
'EN1215'EOL
'EN1215W
'EN1215W'EOL
'EN1216
'EN1223D
'EN1223S
'EN1224
'EN1224_ON
'EN1233D
'EN1233S
'EN1234D
'EN1235D
'EN1235DF
'EN1235S
'EN1235SF
'EN1236D
'EN1238D
'EN1240
'EN1242
'EN1244
'EN1247
'EN1249
'EN1252
'EN1260
'EN1261
'EN1262
'EN1265
'EN1941
'EN5040
'EN5000

Private Sub cmdCancel_Click()
  PreviousForm
  Unload Me
End Sub

'"EN1210","EN1210EOL","EN1210W","EN1212","EN1215","EN1215EOL","EN1215W","EN1215WEOL","EN1216","EN1223D","EN1223S","EN1224","EN1224_ON","EN1233D","EN1233S","EN1234D","EN1235D","EN1235DF","EN1235S","EN1235SF","EN1236D","EN1238D","EN1240","EN1242","EN1244","EN1247","EN1249","EN1252","EN1260","EN1261","EN1262","EN1265","EN1941"


Private Sub cmdNew_Click()
  cmdNew.Enabled = False
  cmdCancel.Enabled = False
  DoUpgrade
  cmdCancel.Enabled = True
  cmdNew.Enabled = True
End Sub


Sub DoUpgrade()
  ' make sure connected to ACG
  ' if not, notify and Bail
  Dim AllowedTypes  As Collection
  Dim AllowedType   As cSimpleDevice
  Dim ATArray       As Variant
  Dim simpledevice  As cSimpleDevice

  Dim Model         As String
  Dim SQL           As String
  Dim j             As Long
  Dim ZoneID        As Long

  Dim HTTPRequest   As cHTTPRequest
  Dim rc            As Long
  Dim ZoneInfoList  As cZoneInfoList
  Dim ZoneInfo      As cZoneInfo
  Set ZoneInfoList = New cZoneInfoList

  Set HTTPRequest = New cHTTPRequest
  Call HTTPRequest.GetZoneList(GetHTTP & "://" & IP1, USER1, PW1)
  Do Until HTTPRequest.Ready
    DoEvents
  Loop
  Select Case HTTPRequest.StatusCode
    Case 200, 201
    Case Else
  End Select
  If Len(HTTPRequest.XML) Then
    rc = ZoneInfoList.LoadXML(HTTPRequest.XML)
  End If
  Set HTTPRequest = Nothing
  If rc Then

    Set UpgradeDevices = New Collection

    ATArray = Array("EN1210", "EN1210EOL", "EN1210-60", "EN1210W", "EN1210W-60", "EN1210-240", "EN1212", "EN1215", _
                    "EN1215EOL", "EN1215W", "EN1215WEOL", "EN1216", "EN1221-60", "EN1221S-60", "EN1223D", _
                    "EN1223S", "ES1223S-60", "EN1224", "EN1224_ON", "EN1233D", "EN1233S", _
                    "EN1234D", "EN1235D", "EN1235DF", "EN1235S", "EN1235SF", "EN1252", _
                    "EN1236D", "EN1238D", "EN1240", "EN1242", "EN1244", "EN1247", "EN1249", _
                     "EN1260", "EN1261", "EN1262", "EN1265", "EN1941", "EN1941-60", "EN1941XS", "EN5040", "EN5000")

    Set AllowedTypes = New Collection
    Dim t           As String
    Dim tkey        As String

    For j = LBound(ATArray) To UBound(ATArray)
      t = ATArray(j)
      tkey = Right$(ATArray(j), Len(ATArray(j)) - 2)
      AllowedTypes.Add t, tkey  ' chop off the EN
    Next

    Dim ESDeviceType As ESDeviceTypeType
    Dim rs          As ADODB.Recordset
    Set rs = ConnExecute("SELECT * FROM Devices where deleted = 0")
    Do Until rs.EOF
      Set simpledevice = New cSimpleDevice
      simpledevice.ID = rs("DeviceID")
      simpledevice.ACGID = rs("IDM")
      simpledevice.Serial = rs("Serial") & ""
      simpledevice.Model = rs("Model") & ""
      simpledevice.AltModel = Model


      If simpledevice.Model = "EN5000" Then
        simpledevice.Model = "EN5040"
      End If
      If simpledevice.Model = "EN5040" Then
        simpledevice.IsRepeater = True
      End If

      If 0 = StrComp(left$(simpledevice.AltModel, 2), "ES", vbTextCompare) Then
        Model = simpledevice.AltModel
        Mid$(Model, 1, 2) = "EN"
        simpledevice.AltModel = Model
      End If

      ESDeviceType = GetDeviceTypeByModel(simpledevice.Model)

      If InStr(simpledevice.Model, "5040") Then
        simpledevice.MID = simpledevice.MID
      End If

      simpledevice.MID = ESDeviceType.MID
      simpledevice.PTI = ESDeviceType.PTI
      simpledevice.CLS = ESDeviceType.CLS
      'simpledevice.Checkin6080 = ESDeviceType.Checkin


      simpledevice.IsPortable = ESDeviceType.Portable
      simpledevice.IsRef = ESDeviceType.Fixed

      'simpledevice.Checkin6080 = 20 * 180 '20x  (3) minutes standard
      simpledevice.Checkin6080 = Max(MIN_CHECKIN, ESDeviceType.Checkin * 1 * 60)
      UpgradeDevices.Add simpledevice, simpledevice.Serial & ""

      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    On Error Resume Next
    For j = 1 To UpgradeDevices.Count
      Set simpledevice = UpgradeDevices(j)
      'split into upgradable and non-upgradable
      Set AllowedType = Nothing
      Set AllowedType = AllowedTypes(simpledevice.AltModel)
      If AllowedType Is Nothing Then
        simpledevice.UpGradeStatus = 2  ' non upgrade
        'NonUpgradeDevices.Add UpgradeDevices(j)   ' add it to our list of non upgraded units so we can report it
        'UpgradeDevices.Remove j                   ' remove it from our good list
      End If
    Next

    Dim hfile       As Long


    For j = 1 To UpgradeDevices.Count
      Set simpledevice = UpgradeDevices(j)
      ZoneID = ZoneInfoList.ScanforSerial(simpledevice.Serial)  ' get ZoneID if it's in our list
      If ZoneID Then
        ' it's registered already , just update the database
        'update the database w/ ZoneID
        ' get the device
        If InIDE Then
          hfile = FreeFile
          Open App.Path & "\upgrade.Log" For Append As #hfile
        End If
        Set HTTPRequest = New cHTTPRequest
        Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, ZoneID)
        If Not ZoneInfo Is Nothing Then
          If InIDE Then
            Print #hfile, simpledevice.Serial; " Already Registered, Update Database"
          End If
          SQL = "UPDATE devices SET IDM = " & ZoneID & ", " & "IDL = " & ZoneInfo.IDL & "  WHERE serial = '" & simpledevice.Serial & "'"
          ConnExecute SQL
          ' Don't change any settings
          simpledevice.UpGradeStatus = 1
        Else
          If InIDE Then
            Print #hfile, simpledevice.Serial; " Can't get info"
          End If
          simpledevice.UpGradeStatus = 3
        End If
        Set ZoneInfo = Nothing
        Set HTTPRequest = Nothing
        If InIDE Then
          Close hfile
        End If
      Else
        ' try and register it
        'Sleep 200
        'If simpledevice.IDL = 1 Then
        '          Stop
        'End If
        If InIDE Then
          hfile = FreeFile
          Open App.Path & "\upgrade.Log" For Append As #hfile
          Print #hfile, simpledevice.Serial; " Attempting Upgrade"
          Close hfile
        End If

        ZoneID = UpgradeDevice(simpledevice)

        If ZoneID Then
          If InIDE Then
            hfile = FreeFile
            Open App.Path & "\upgrade.Log" For Append As #hfile
            Print #hfile, simpledevice.Serial; " Got ZoneID "; ZoneID
            Close hfile
          End If
          Set HTTPRequest = New cHTTPRequest
          Set ZoneInfo = HTTPRequest.GetSingleZoneInfo(GetHTTP & "://" & IP1, USER1, PW1, ZoneID)
          If Not ZoneInfo Is Nothing Then

            SQL = "UPDATE devices SET IDM = " & ZoneID & ", " & "IDL = " & ZoneInfo.IDL & " WHERE serial = '" & simpledevice.Serial & "'"
            ConnExecute SQL
            simpledevice.UpGradeStatus = 1
          Else
            simpledevice.UpGradeStatus = 3

          End If
          Set ZoneInfo = Nothing
          Set HTTPRequest = Nothing

        Else
          ' Failed Upgrades
          'FailedUpgradeDevices.Add simpledevice, simpledevice.Serial & ""
          'UpgradeDevices.Remove j
          simpledevice.UpGradeStatus = 3
          If InIDE Then
            hfile = FreeFile
            Open App.Path & "\upgrade.Log" For Append As #hfile
            Print #hfile, simpledevice.Serial; " Failed Upgrade"
            Close hfile
          End If

        End If
      End If
      DoEvents
    Next

  Else      'couldn't get ZoneInfolist

  End If
  UpdateResults
End Sub
Sub UpdateResults()
  Dim li            As ListItem
  Dim simpledevice  As cSimpleDevice
  lvUpgraded.ListItems.Clear
  lvNonUpgrade.ListItems.Clear
  lvFailed.ListItems.Clear
  For Each simpledevice In UpgradeDevices
    Select Case simpledevice.UpGradeStatus
      Case 1

        Set li = lvUpgraded.ListItems.Add(, , simpledevice.Serial)
        li.SubItems(1) = simpledevice.Model
      Case 2
        Set li = lvNonUpgrade.ListItems.Add(, , simpledevice.Serial)
        li.SubItems(1) = simpledevice.Model

      Case 3
        Set li = lvFailed.ListItems.Add(, , simpledevice.Serial)
        li.SubItems(1) = simpledevice.Model

      Case Else
    End Select
  Next


End Sub
Sub Fill()
  ' nothing to do
End Sub
Private Sub Form_Load()
ResetActivityTime
  Set UpgradeDevices = New Collection
  fraEnabler.BackColor = Me.BackColor
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
fraEnabler.Enabled = True
End Sub

Public Sub Host(ByVal hwnd As Long)

  SetParent fraEnabler.hwnd, hwnd
End Sub

Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

Private Sub tmrSearch_Timer()

End Sub

Private Sub Label_Click()
  Debug.Print "Label_Click()"
End Sub

Private Sub lvNonUpgrade_Click()
 Debug.Print "lvNonUpgrade_Click()"
End Sub

Private Sub lvUpgraded_Click()
  Debug.Print "lvUpgraded_Click()"
End Sub

Private Sub lvUpgraded_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Debug.Print "lvUpgraded_ColumnClick"
End Sub
