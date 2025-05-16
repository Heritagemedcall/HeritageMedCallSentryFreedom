VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZoneInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'To Get List
'http://192.168.60.80/PSIA/AreaControl/PartitionMembers/Zones/ZoneInfoList

'<ZoneInfo>
' <ID>1</ID>
' <DeviceID>3351057</DeviceID>
' <PTI>0</PTI>
' <Type>Base Receiver</Type>
' <MID>0</MID>
' <Locatable>false</Locatable>
' <Description>base receiver</Description>
' <IsRef>false</IsRef>
' <SyncWindow>300</SyncWindow>
' <SyncTimeout>0</SyncTimeout>
' <MessageExpirationTime>0</MessageExpirationTime>
' <CheckInTime>180</CheckInTime>
' <SupervisionWindow>3600</SupervisionWindow>
' <PartitionList></PartitionList>
' <IsAlarm>false</IsAlarm>
' <IsMissing>false</IsMissing>
' <IsSPDevice>false</IsSPDevice>
'</ZoneInfo>

Private mXML As String
Public Validated    As Long

Public ID           As Long       ' the 6080 assigned ID (sequence)
Public DeviceID     As Long
Public PTI          As Long
Public TypeName     As String
Public MID          As Long
Private mLocatable  As String

Public Description  As String
Private mIsRef      As String     ' fixed reference point

Public SyncWindow   As Long       ' for two-way devices
Public SyncTimeOut  As Long
Public MessageExpirationTime As Long


Public CheckInTime  As Long       ' how often device checks in
Private mSupervisionWindow As Long  ' how long to wait before declaring missing

Public PartitionListString As String


Private mIsAlarm    As String     ' Is there an alarm pending
Private mIsMissing  As String
Private mIsSPDevice As String     ' Is Soft Pointer Device for setting up Soft Points


' string to long
Public IsLocatable   As Long
Public IsFixedDevice As Long
Public Alarmed       As Long
Public Missing       As Long
Public IsSoftPointer As Long



Public Partitionlist As Collection

Public currstate As String
'<CurrState>00000000</CurrState>




Public Property Get IDL() As Long
  If IsLocatable Then
    IDL = 1
  ElseIf IsFixedDevice Then
    IDL = 2
  ElseIf IsSoftPointer Then
    IDL = 3
  Else
    IDL = 0
  End If
End Property

Public Function ParseNode(Node As IXMLDOMNode) As Long
  Dim childnode     As IXMLDOMNode
  Dim childnodes    As IXMLDOMNodeList
  On Error Resume Next
  For Each childnode In Node.childnodes
    Select Case childnode.baseName
      Case "ID"
        ID = Val(childnode.text)
      Case "DeviceID"
        DeviceID = Val(childnode.text)
      Case "PTI"
        PTI = Val(childnode.text)
      Case "Type"
        TypeName = childnode.text
      Case "MID"
        MID = Val(childnode.text)
      Case "Locatable"
        locatable = childnode.text
      Case "Description"
        Description = XMLDecode(Trim$(childnode.text))
      Case "IsRef"
        IsRef = childnode.text
      Case "SyncWindow"
        SyncWindow = Val(childnode.text)
      Case "SyncTimeout"
        SyncTimeOut = Val(childnode.text)
      Case "MessageExpirationTime"
        MessageExpirationTime = Val(childnode.text)
      Case "CheckInTime"
        CheckInTime = Val(childnode.text)
      Case "SupervisionWindow"
        SupervisionWindow = Val(childnode.text)
      Case "IsAlarm"
        IsAlarm = childnode.text
      Case "IsMissing"
        IsMissing = childnode.text
      Case "IsSPDevice"
        IsSPDevice = childnode.text
      Case "PartitionList"
        PartitionListString = childnode.text
        ParsePartitionListString childnode.childnodes
        
      Case "CurrState"
        currstate = childnode.text
        If TypeName <> "" Then
          
        End If
        
    End Select
  Next
  ParseNode = (Err.Number = 0)
End Function
Public Function ParsePartitionListString(PartitionListNodes As IXMLDOMNodeList) As Long
'<PartitionList>
' <Partition>
'   <PartitionID>13</PartitionID>
' </Partition>
' <Partition>
'   <PartitionID>23</PartitionID>
' </Partition>
' <Partition>
'  <PartitionID>8</PartitionID>
' </Partition>
'</PartitionList>
  
  Dim pnode As IXMLDOMNode
  Dim partition As cPartition
  
  For Each pnode In PartitionListNodes
    Set partition = New cPartition
    Partitionlist.Add partition
    Select Case pnode.baseName
      Case "Partition"
        partition.PartitionID = Val(pnode.text)
    End Select
  Next
  
  
End Function

Public Function ParseXML(ByVal XML As String) As Long
  Dim doc As DOMDocument60
  Dim Node As IXMLDOMNode
  
  mXML = XML
  
  Set doc = New DOMDocument60
  If doc.LoadXML(XML) Then
    Set Node = doc.firstChild
    ParseNode Node
    ParseXML = 1
  Else
    ParseXML = 0
  End If
  
  
  
End Function

Public Property Get locatable() As String
  locatable = mLocatable
End Property

Public Property Let locatable(ByVal locatable As String)
  IsLocatable = (LCase$(locatable) = "true") And 1
  mLocatable = locatable
End Property

Public Property Get IsRef() As String
  IsRef = mIsRef
End Property

Public Property Let IsRef(ByVal IsRef As String)
  IsFixedDevice = (LCase$(IsRef) = "true") And 1
  mIsRef = IsRef
End Property

Public Property Get IsAlarm() As String
  IsAlarm = mIsAlarm
End Property

Public Property Let IsAlarm(ByVal IsAlarm As String)
  Alarmed = (LCase$(IsAlarm) = "true") And 1
  mIsAlarm = IsAlarm
End Property

Public Property Get IsMissing() As String
  IsMissing = mIsMissing
End Property

Public Property Let IsMissing(ByVal IsMissing As String)
  Missing = (LCase$(IsMissing) = "true") And 1
  mIsMissing = IsMissing
End Property

Public Property Get IsSPDevice() As String
  IsSPDevice = mIsSPDevice
End Property

Public Property Let IsSPDevice(ByVal IsSPDevice As String)
  IsSoftPointer = (LCase$(IsSPDevice) = "true") And 1
  mIsSPDevice = IsSPDevice
End Property

Public Property Get HexID() As String

  HexID = Hex$(DeviceID)

End Property


Public Property Get SupervisionWindow() As Long

  SupervisionWindow = mSupervisionWindow

End Property

Public Property Let SupervisionWindow(ByVal SupervisionWindow As Long)

  mSupervisionWindow = SupervisionWindow

End Property

Private Sub Class_Initialize()
  Set Partitionlist = New Collection
End Sub

Public Property Get XML() As String

  XML = mXML

End Property

Private Property Let XML(ByVal XML As String)

  mXML = XML

End Property
