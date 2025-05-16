VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmAdapters 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   1995
   ClientTop       =   9150
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   9945
   Begin VB.Frame fraEnabler 
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9660
      Begin VB.CommandButton cmdRefresh 
         Height          =   585
         Left            =   8700
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAdapters.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   675
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
         Left            =   8475
         TabIndex        =   7
         Top             =   2370
         Width           =   1170
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
         Left            =   8475
         TabIndex        =   6
         Top             =   1785
         Width           =   1170
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2415
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   4260
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "IP"
            Text            =   "IP Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "D"
            Text            =   "Adapter Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   4560
         TabIndex        =   3
         Top             =   150
         Width           =   585
      End
      Begin VB.Label lblMAC 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   1660
         TabIndex        =   2
         Top             =   150
         Width           =   585
      End
      Begin VB.Label lblCurrentIP 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   1
         Top             =   150
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private quitting    As Boolean
Private LastIndex   As Long

Public Sub Fill()
        Dim Adapter As cAdapter
        Dim li As ListItem
        
        
        Dim CurrentIP As String
        Dim CurrentAdapter As String
        Dim CurrentMAC As String
        Dim LocalIP As String
        
10      LocalIP = frmTimer.WinsockHost(0).LocalIP
        

        
20      Adapters.RefreshAdapters
        
        
30      lvMain.ListItems.Clear
40      For Each Adapter In Adapters
        
50         Set li = lvMain.ListItems.Add(, , Adapter.DhcpIPAddress)
60         li.SubItems(1) = Adapter.MacAddress
70         li.SubItems(2) = Adapter.Description
           
80         If 0 = StrComp(Adapter.DhcpIPAddress, LocalIP, vbTextCompare) Then
90          CurrentIP = Adapter.DhcpIPAddress
100         CurrentAdapter = Adapter.Description
110         CurrentMAC = Adapter.MacAddress
120         li.Selected = True
130        End If
           

140     Next
150     If lvMain.ListItems.Count = 0 Then
160        cmdApply.Enabled = False
170        Set li = lvMain.ListItems.Add(, , "0.0.0.0")
180        li.SubItems(1) = " "
190        li.SubItems(2) = "No Network Adapters Configured"
200        cmdApply.Enabled = False
210     Else
220       cmdApply.Enabled = True
230     End If
          
          
240    lblCurrentIP.Caption = CurrentIP
250    lblMAC.Caption = CurrentMAC
260    lblDesc.Caption = CurrentAdapter
            
          
          
End Sub


Sub Apply()
  Dim li                 As ListItem
  Dim Adapter            As cAdapter
  If lvMain.SelectedItem Is Nothing Then
    Beep
  Else
    On Error Resume Next

      
    Set li = lvMain.SelectedItem
    Set Adapter = Adapters.GetAdapterByMAC(li.SubItems(1))
    
    BootLog "Change Get WinsockHost IP " & frmTimer.WinsockHost(0).LocalIP
    
    BootLog "Change Adapter Desc " & Adapter.Description & " #" & Err.Number
    If Not (Adapter Is Nothing) Then
      
      If Adapter.DhcpIPAddress <> "0.0.0.0" Then
        WriteSetting "Adapter", "MasterIP", Adapter.DhcpIPAddress
        BootLog "Change MasterIP " & Adapter.DhcpIPAddress & " #" & Err.Number
        WriteSetting "Adapter", "MAC", Adapter.MacAddress
        BootLog "Change Adapter MAC " & Adapter.MacAddress & " #" & Err.Number
        WriteSetting "Adapter", "AdapterName", Adapter.Description
        BootLog "Change AdapterName " & Adapter.Description & " #" & Err.Number
        WriteSetting "Adapter", "ServiceName", Adapter.ServiceName
        BootLog "Change ServiceName " & Adapter.ServiceName & " #" & Err.Number
        frmTimer.WinsockHost(0).Close
        BootLog "Change WinsockHost Closed"
        frmTimer.WinsockHost(0).Bind Configuration.HostPort, Adapter.DhcpIPAddress
        BootLog "Change WinsockHost Bind " & Configuration.HostPort & " " & Adapter.DhcpIPAddress & " #" & Err.Number
        frmTimer.WinsockHost(0).Listen
        BootLog "Change WinsockHost Listening " & " #" & Err.Number
      End If
    End If

  End If

  Fill

End Sub

Public Sub Host(ByVal hwnd As Long)
  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = 9660
  fraEnabler.Height = ENABLER_HEIGHT

  SetParent fraEnabler.hwnd, hwnd
End Sub
Public Sub UnHost()
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdRefresh_Click()
    ResetActivityTime
  Fill
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  
  UnHost
End Sub
Private Sub Form_Load()
  ResetActivityTime
  Configurelvmain
  Fill
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    LastIndex = Item.index
  End If
  'ResetRemoteRefreshCounter
End Sub

Private Sub cmdApply_Click()
  ResetActivityTime
  Apply
End Sub
Sub Configurelvmain()
  Dim ch As ColumnHeader
  
  
  
  
  lvMain.ColumnHeaders.Clear
  lvMain.Sorted = False
  Set ch = lvMain.ColumnHeaders.Add(, "IP", "IP Address", 1600)
  Set ch = lvMain.ColumnHeaders.Add(, "MAC", "MAC", 2000)
  Set ch = lvMain.ColumnHeaders.Add(, "D", "Adpater Name", 4500)
  
  Me.lblCurrentIP.left = 60
  lblMAC.left = 1600 + 60
  lblDesc.left = 1600 + 2000 + 60
  
  
End Sub
