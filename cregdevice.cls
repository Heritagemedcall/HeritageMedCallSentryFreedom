VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'<Registration>
'  <MetadataHeader>
'    <MetaVersion>1.0</MetaVersion>
'    <MetaID>inovonics.com/ACG.System/registration</MetaID>
'    <MetaSourceID>{207DF398-5015-9281-4F2C-71CE056CEA4C}</MetaSourceID>
'    <MetaSourceLocalID>1061482</MetaSourceLocalID>
'    <MetaTime>2013-04-18T20:19:12.16807Z</MetaTime>
'    <MetaPriority>6</MetaPriority>
'  </MetadataHeader>
'  <EventData>
'    <DeviceID>1061482</DeviceID>
'    <DeviceType>EN1210W Door/Window</DeviceType>
'    <IsRegistered>false</IsRegistered>
'    <MID>178</MID>
'    <PTI>3</PTI>
'  </EventData>
'</Registration>


Public i6080ID      As String  '<MetaSourceID>{207DF398-5015-9281-4F2C-71CE056CEA4C}</MetaSourceID>
Public MFG          As String  '<MetaID>inovonics.com/ACG.System/registration</MetaID>

Public ID           As Long  ' Seqence # ?
Public DeviceID     As Long  ' device serial number
Private mDeviceType As String
Public IsRegistered As Boolean
Public MID          As Long
Public PTI          As Long



Public Priority     As Long
Public EventTime    As Date
Public Property Get CLSPTI() As Long
  CLSPTI = CLS * 256& + PTI
End Property
Public Property Get CLS() As Long
  
  Select Case left$(mDeviceType, 6)
    Case "EN5040"
      CLS = &H41
    Case "EN5000"
      CLS = &H41
    Case "EN3954"
      CLS = &H39
    Case "ES1723"
      CLS = &H3C
    Case "EN6080"
      CLS = &H1
    Case Else
      CLS = &H3E
  End Select
End Property

Public Property Get FullHexSerial() As String
  FullHexSerial = Right$("00" & Hex$(MID), 2) & Right$("000000" & Hex$(DeviceID), 6)

End Property

Public Function ParseXML(ByVal XML As String) As Long
  Dim doc           As DOMDocument60

  Dim RootNode      As IXMLDOMNode
  Dim RootName      As String
  Dim Node          As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList
  Dim childnode     As IXMLDOMNode

  Set doc = New DOMDocument60
  If doc.LoadXML(XML) Then
    Debug.Print XML
    
    Set RootNode = doc.firstChild
    If 0 = StrComp(RootNode.baseName, "Registration", vbTextCompare) Then
      If RootNode.childnodes.Length > 1 Then
        Set Node = RootNode.childnodes(1)
        If 0 = StrComp(Node.baseName, "EventData", vbTextCompare) Then
          For Each childnode In Node.childnodes
            Select Case childnode.baseName
              Case "DeviceID"
                DeviceID = Val(childnode.text)
              Case "DeviceType"
                mDeviceType = childnode.text
              Case "IsRegistered"
                IsRegistered = CBool(childnode.text)
              Case "MID"
                MID = Val(childnode.text) ' 178 dec = B2 Hex
              Case "PTI"
                If 0 = StrComp(left$(mDeviceType, 6), "EN5040", vbTextCompare) Then
                  PTI = &H1
                Else
                  PTI = Val(childnode.text) ' 1941 XS is 12 decimal hex 0x0C
                  
                End If
            End Select
          Next
          ParseXML = 1
        End If
      End If
    End If
    Set doc = Nothing

  End If

End Function

' read only
Public Property Get DeviceType() As String
  DeviceType = mDeviceType
End Property

