VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmTimer 
   BorderStyle     =   0  'None
   ClientHeight    =   1740
   ClientLeft      =   5760
   ClientTop       =   4440
   ClientWidth     =   5775
   Icon            =   "frmTimer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin Heritage_Freedom2.AsyncReader AsyncReader1 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
   End
   Begin MSWinsockLib.Winsock wsGeneric 
      Index           =   0
      Left            =   1860
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrReminders 
      Interval        =   1000
      Left            =   5220
      Top             =   330
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4200
      Top             =   300
   End
   Begin MSWinsockLib.Winsock WinsockHost 
      Index           =   0
      Left            =   1890
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1005
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   210
      Top             =   300
   End
   Begin MSWinsockLib.Winsock WinsockClient 
      Left            =   2670
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WinsockClientInterraction 
      Left            =   3210
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Winsock Array"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   900
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminders"
      Height          =   195
      Left            =   4950
      TabIndex        =   4
      Top             =   90
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Timer "
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "One Second"
      Height          =   195
      Left            =   3990
      TabIndex        =   2
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      Height          =   195
      Left            =   2910
      TabIndex        =   1
      Top             =   60
      Width           =   390
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listen Array"
      Height          =   195
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   60
      Width           =   825
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents ft As cFastClock
Attribute ft.VB_VarHelpID = -1

Private Stopped       As Boolean
Const MAX_CLIENT_CONNECTIONS = 24  '


Public Function GetGenericWinsock() As Winsock
        Dim j As Long
        Dim NextUp As Long
        Dim ws As Winsock
        Set ws = wsGeneric(0)
        
10      On Error Resume Next
20      For j = wsGeneric.UBound To 1 Step -1
30        If wsGeneric(j) Is Nothing Then
40          Load wsGeneric(j)
50          Exit For
60        End If
70        If wsGeneric(j).tag = "" Then
            ' it free, use it
80          Exit For
90        End If
100     Next
110     If j = 0 Then
120       NextUp = wsGeneric.UBound + 1
130       Load wsGeneric(NextUp)
140       Set ws = wsGeneric(NextUp)
150       ws.tag = "in use"
160       Set GetGenericWinsock = ws
170       Set ws = Nothing
180     Else
190       wsGeneric(j).tag = "in use"
200       Set GetGenericWinsock = wsGeneric(j)
210     End If

  If Err.Number Then
    Debug.Print Err.Number, Err.Description, Erl
  End If

End Function

Private Sub Form_Load()

  Dim MAC                As String
  Dim MasterIP           As String
  Dim Adapter            As cAdapter


  If InIDE Then
    Timer1.interval = 20
  End If
  If MASTER Then


'    MAC = Trim$(ReadSetting("Configuration", "MAC", ""))
'
'    If Len(MAC) = 0 Then
'      ' get first adapter MAC
'      If Adapters.Adapters.Count Then
'        Set Adapter = Adapters.Adapters(1)
'        MAC = Adapter.MacAddress
'        MasterIP = Adapter.Address
'        WriteSetting "Configuration", "MAC", MAC
'      Else
'        MAC = ""
'        MasterIP = "0.0.0.0"
'        WriteSetting "Configuration", "MAC", MAC
'      End If
'    Else  ' we have a mac, is it any good?
'      For Each Adapter In Adapters
'        If 0 = StrComp(Adapter.MacAddress, MAC, vbTextCompare) Then
'          MasterIP = Adapter.Address
'          MAC = Adapter.MacAddress
'          Exit For
'        End If
'        If Len(MasterIP) = 0 Then
'          If Adapters.Adapters.Count Then
'            Set Adapter = Adapters.Adapters(1)
'            MAC = Adapter.MacAddress
'            MasterIP = Adapter.Address
'            WriteSetting "Configuration", "MAC", MAC
'          Else
'            MAC = ""
'            MasterIP = "0.0.0.0"
'            WriteSetting "Configuration", "MAC", MAC
'          End If
'        Else
'          WriteSetting "Configuration", "IP", MasterIP
'          WriteSetting "Configuration", "MAC", MAC
'        End If
'      Next
'
'
'
'      Debug.Print WinsockHost(0).LocalIP
'      WinsockHost(0).Bind "2500", MasterIP

    'End If
    'On Error Resume Next
    'ConsoleID = Replace(MAC, ":", "")
    On Error GoTo 0

  End If
End Sub

Public Sub StopTimer()
  Stopped = True
  If InIDE Then
    Timer1.Enabled = False
  Else
    If Not ft Is Nothing Then
      ft.StopIt
    End If
  End If
End Sub
Public Sub StartTimer()
  If InIDE Then
    Timer1.Enabled = True
  Else
    If ft Is Nothing Then
      Set ft = New cFastClock
      ft.interval = 20  ' twenty ms
    End If
    ft.RunIt
  End If
  Stopped = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'MsgBox "FRMTIMER.UNLOAD"
  
  Dim unloader As Long
  StopTimer

  Sleep 100
  unloader = Win32.timeGetTime() + 100
  Do Until unloader < Win32.timeGetTime()
    DoEvents
  Loop

  If Not InIDE Then
    Set ft = Nothing
  End If
  unloader = Win32.timeGetTime() + 100
  Do Until unloader < Win32.timeGetTime()
    DoEvents
  Loop

End Sub

Private Sub ft_TimerEvent()
  If Not Stopped Then
    Call HeartBeat
  End If
End Sub

Private Sub Timer1_Timer()
  Static t As Long
  Dim Elapsed As Long
  
'  Elapsed = Win32.timeGetTime - t
  'If Elapsed > 200 Then
    'dbg "Timer1_timer delayed " & Elapsed
  'End If
 ' t = Win32.timeGetTime
  'dbg "Timer1_timer"
  Call HeartBeat

End Sub

Private Sub Timer2_Timer()
  'This timer is for long-running things that may block normal execution
  Call AuxLoop
  
End Sub

Private Sub Timer3_Timer()

End Sub

Private Sub tmrReminders_Timer()
  If MASTER Then
    RemindersUpdate ' once second interval
  End If
End Sub

Private Sub WinsockClient_Close()
  dbg "WinsockClient.Close"
  WinsockClient.Close

End Sub

Private Sub WinsockClient_Connect()
  HostConnection.Socket_Connect
  'frmMain.PacketToggle
End Sub

Private Sub WinsockClient_ConnectionRequest(ByVal requestID As Long)
' ignore... we are the client, not the server!
End Sub

Private Sub WinsockClient_DataArrival(ByVal bytesTotal As Long)
  HostConnection.Socket_DataArrival bytesTotal
  'frmMain.PacketToggle
End Sub

Private Sub WinsockClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' oops
' log error
  HostConnection.Socket_Close
End Sub

Private Sub WinsockClient_SendComplete()
  HostConnection.Socket_SendComplete
  'frmMain.PacketToggle
End Sub

Private Sub WinsockClient_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  HostConnection.Socket_SendProgress bytesSent, bytesRemaining
  'frmMain.PacketToggle
End Sub

Private Sub WinsockClientInterraction_Close()
  WinsockClientInterraction.Close
End Sub

Private Sub WinsockClientInterraction_Connect()
  HostInterraction.Socket_Connect
End Sub

Private Sub WinsockClientInterraction_ConnectionRequest(ByVal requestID As Long)
  ' ignore
End Sub

Private Sub WinsockClientInterraction_DataArrival(ByVal bytesTotal As Long)
  HostInterraction.Socket_DataArrival bytesTotal
End Sub

Private Sub WinsockClientInterraction_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'HostInterraction.WinsockHostInterraction_Error.Close
  HostInterraction.Socket_Close
End Sub

Private Sub WinsockClientInterraction_SendComplete()
  HostInterraction.Socket_SendComplete
  'frmMain.PacketToggle

End Sub

Private Sub WinsockClientInterraction_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  HostInterraction.Socket_SendProgress bytesSent, bytesRemaining
  'frmMain.PacketToggle
End Sub

Private Sub WinsockHost_Close(index As Integer)
  dbg "WinsockHost(" & index & ").Close"
  WinsockHost(index).Close
End Sub


Private Sub WinsockHost_ConnectionRequest(index As Integer, ByVal requestID As Long)
  Dim freeindex As Integer
  Dim j         As Integer
  Dim ClientConnection As cClientConnection
  Debug.Print "WinsockHost_ConnectionRequest ", index, requestID

  If index = 0 Then   ' 0 is the listener
    ' Accept the connection.
    For j = 1 To ClientConnections.Count
      Set ClientConnection = ClientConnections(j)
      'dbg " ClientConnection " & j & " = " & IIf(ClientConnection.Closed, "Closed", "In Use")
      If ClientConnection.Closed Then
        Exit For
      End If
    Next

    If j <= ClientConnections.Count Then
      'reuse it
      'dbg " ClientConnection " & j & " being reused"
      ClientConnection.LocalPort = 0
      ClientConnection.Accept requestID
    'ElseIf j >= ClientConnections.count  Then
    
    ElseIf (j > MAX_CLIENT_CONNECTIONS) Then '8 Then
      Set ClientConnection = Nothing
    
    Else
      'dbg " ClientConnection (NEW)"
      Set ClientConnection = New cClientConnection
      ClientConnections.Add ClientConnection
      Load WinsockHost(ClientConnections.Count)
      Set ClientConnection.Socket = WinsockHost(ClientConnections.Count)
      ClientConnection.LocalPort = 0
      ClientConnection.Accept requestID
    End If
    
  
  
  End If

End Sub

'Public Function GetFreeConnectionIndex() As Integer
'  Dim index As Integer
'  Dim Count As Integer
'  Count = WinsockHost.UBound
'  GetFreeConnectionIndex = 0
'
'  ' Try and find a winsock which has been created but is not in use
'  ' any longer - this should help keep memory usage lower.
'  For index = 1 To Count
'    If GetFreeConnectionIndex = 0 Then
'      Select Case WinsockHost(index).State
'        Case sckClosed
'          GetFreeConnectionIndex = index
'          Exit For
'        Case sckError
'          WinsockHost(index).Close
'          GetFreeConnectionIndex = index
'          Exit For
'      End Select
'    End If
'  Next
'
'  ' If there isn't one free create a new one.
'  If GetFreeConnectionIndex = 0 Then
'    Count = Count + 1
'    Load WinsockHost(Count)
'    GetFreeConnectionIndex = Count
'  End If
'End Function

' WinsockHost is the listener

Private Sub WinsockHost_DataArrival(index As Integer, ByVal bytesTotal As Long)
  ClientConnections(index).Socket_DataArrival bytesTotal
End Sub

Private Sub WinsockHost_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  On Error Resume Next
  If index > 0 Then
    'dbg "frmtimer.WinsockHost_Error index = " & index
    WinsockHost(index).Close
    ClientConnections(index).Socket_Error Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay
    'ClientConnections(index).Close
  End If
End Sub
Private Sub WinsockHost_SendComplete(index As Integer)
'  ClientConnections(index).Socket_SendComplete
End Sub
Private Sub WinsockHost_Connect(index As Integer)
  ClientConnections(index).Socket_Connect
End Sub

Private Sub WinsockHost_SendProgress(index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'ClientConnections(index).Socket_SendProgress bytesSent, bytesRemaining
End Sub

Private Sub ws_pager_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

