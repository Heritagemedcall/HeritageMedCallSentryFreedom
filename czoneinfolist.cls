VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZoneInfoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public doc          As DOMDocument60
Public RootNode     As IXMLDOMNode
Public RootName     As String


Private Node        As IXMLDOMNode
Private NodeList    As IXMLDOMNodeList

Public ZoneList     As Collection

Public Function ScanforSerial(ByVal Serial As String) As Long
  Dim ZoneInfo As cZoneInfo
  Dim SerialNum As Long
  ' Scans ZoneInfoList for Transmitter/Zone returns nothing if not found
  SerialNum = Val("&h" & Right$(Serial, 6))

  
  For Each ZoneInfo In ZoneList
    If ZoneInfo.DeviceID = SerialNum Then
      ScanforSerial = ZoneInfo.ID
      Exit For
    End If
  Next
    
End Function

Public Function LoadXML(ByVal XML As String) As Long
  Dim rc            As Boolean
  ' returns 1 on success
  On Error Resume Next
  Set doc = Nothing
  Set doc = New DOMDocument60
  doc.Async = False
  rc = doc.LoadXML(XML)

  If rc Then
    Set RootNode = doc.firstChild
    If RootNode Is Nothing Then
      rc = False
    Else
      RootName = RootNode.baseName
      ParseZoneList
    End If
  
  Else
  
    LogXML XML, "Error-ZoneList"
  
  End If
  LoadXML = rc

End Function

Public Function ZoneInfoByHexSerial(ByVal HexSerial As String) As cZoneInfo
  On Error Resume Next
  Set ZoneInfoByHexSerial = ZoneList(HexSerial & "")
End Function

Public Function ParseZoneList()
  Dim Node          As IXMLDOMNode
  Dim subnode       As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList
  Dim ZoneInfo      As cZoneInfo

  Set ZoneList = New Collection

  Set Node = doc.selectSingleNode("ZoneInfoList")
  If Not Node Is Nothing Then
    For Each subnode In Node.childnodes
      If subnode.baseName = "ZoneInfo" Then
        Set ZoneInfo = New cZoneInfo
        DoEvents
        If ZoneInfo.ParseNode(subnode) Then
          ZoneList.Add ZoneInfo, Right$("00" & Hex$(ZoneInfo.MID), 2) & Right$("000000" & ZoneInfo.HexID, 6)
        End If
      End If
    Next
  End If


End Function

Private Sub Class_Initialize()
  Set ZoneList = New Collection
End Sub

Private Sub Class_Terminate()
  Set ZoneList = Nothing
End Sub
