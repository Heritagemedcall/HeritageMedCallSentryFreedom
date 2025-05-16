VERSION 5.00
Begin VB.Form frmPhantom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stress Test"
   ClientHeight    =   1500
   ClientLeft      =   3345
   ClientTop       =   3210
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Height          =   495
      Left            =   780
      TabIndex        =   0
      Top             =   450
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1260
      Top             =   1290
   End
End
Attribute VB_Name = "frmPhantom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ProcessESPacket i6080.ConvertToPacket(i6080.GetNextMessage)

'c6080.PushData(ByVal XML As String)
Private Status As Long
Private XML As String

Private Running As Boolean

Private Sub cmdStartStop_Click()
  If Running Then
    StopPhantom
  Else
    StartPhantom
  End If
  If Running Then
    cmdStartStop.Caption = "Stop"
  Else
    cmdStartStop.Caption = "Start"
  End If
End Sub

Sub StartPhantom()
  Timer1.Enabled = True
  Running = True
End Sub
Sub StopPhantom()
  Timer1.Enabled = False
  Running = False
End Sub


Private Sub Form_Load()
  Centerform Me
  Fill
  
End Sub

Sub Fill()
'  Dim Device As cESDevice
'
'  lstDeviceList.Clear
  
'  For Each Device In Devices.Devices
'    lstDeviceList.AddItem Device.Serial
'    lstDeviceList.ItemData(lstDeviceList.NewIndex) = Device.DecimalSerial
'  Next

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
  Status = Status + 1
  
  Me.Caption = "Seconds " & Status
  XML = ""
  Select Case Status
    Case 1 ' alarm
      XML = "<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.Zone/alarmState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>4019267</MetaSourceLocalID><MetaTime>2014-10-01T15:52:41.396Z</MetaTime><MetaPriority>1</MetaPriority></MetadataHeader><EventData><Info><ID>61</ID><Type>EN1223S Pendant</Type><Description>B23D5443</Description><PartitionList></PartitionList><SCI>ALARM 1</SCI><SCICode>1</SCICode><CurrState>00000100</CurrState></Info><ValueState><IntrusionAlarm>Intrusion</IntrusionAlarm></ValueState></EventData></AreaControlEvent>"
    Case 2  ' alarm clear
      XML = "<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.Zone/alarmState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>4019267</MetaSourceLocalID><MetaTime>2014-10-01T15:52:45.397Z</MetaTime><MetaPriority>4</MetaPriority></MetadataHeader><EventData><Info><ID>61</ID><Type>EN1223S Pendant</Type><Description>B23D5443</Description><PartitionList></PartitionList><SCI>ALARM 1 CLR</SCI><SCICode>2</SCICode><CurrState>00000000</CurrState></Info><ValueState><IntrusionAlarm>OK</IntrusionAlarm></ValueState></EventData></AreaControlEvent>"
    Case 3  ' pendant reset
      XML = "<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.Zone/troubleState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>4019267</MetaSourceLocalID><MetaTime>2014-10-01T15:52:55.116Z</MetaTime><MetaPriority>6</MetaPriority></MetadataHeader><EventData><Info><ID>61</ID><Type>EN1223S Pendant</Type><Description>B23D5443</Description><PartitionList><Partition><PartitionID>1</PartitionID><Description>South hall</Description></Partition></PartitionList><SCI>RESET</SCI><SCICode>21</SCICode><CurrState>00000008</CurrState></Info><ValueState><IntrusionTrouble>Trouble</IntrusionTrouble></ValueState></EventData></AreaControlEvent>"
    Case 4, 12, 22, 38
      XML = "<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.Zone/alarmState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>1061482</MetaSourceLocalID><MetaTime>2014-10-02T18:54:49.751Z</MetaTime><MetaPriority>2</MetaPriority></MetadataHeader><EventData><Info><ID>46</ID><Type>EN1210W Door/Window</Type><Description>B210326A</Description><PartitionList><Partition><PartitionID>1</PartitionID><Description>South hall</Description></Partition></PartitionList><SCI>TAMPER</SCI><SCICode>11</SCICode><CurrState>00000020</CurrState></Info><ValueState><IntrusionAlarm>Tamper</IntrusionAlarm></ValueState></EventData></AreaControlEvent>"
    Case 8, 16, 28, 42
      XML = "<AreaControlEvent><MetadataHeader><MetaVersion>1.0</MetaVersion><MetaID>psiaalliance.org/AreaControl.Zone/alarmState</MetaID><MetaSourceID>{207DF398-5015-9281-50FF-4BC6056EC896}</MetaSourceID><MetaSourceLocalID>1061482</MetaSourceLocalID><MetaTime>2014-10-02T18:54:51.841Z</MetaTime><MetaPriority>4</MetaPriority></MetadataHeader><EventData><Info><ID>46</ID><Type>EN1210W Door/Window</Type><Description>B210326A</Description><PartitionList><Partition><PartitionID>1</PartitionID><Description>South hall</Description></Partition></PartitionList><SCI>TAMPER CLR</SCI><SCICode>12</SCICode><CurrState>00000000</CurrState></Info><ValueState><IntrusionAlarm>OK</IntrusionAlarm></ValueState></EventData></AreaControlEvent>"
    Case Is > 64
      Status = 0
    Case Else
      
      
  End Select
  
  If Len(XML) Then
    If USE6080 Then
      i6080.PushData (XML)
    End If
  End If
End Sub
