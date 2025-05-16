VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSoftPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' stream softpoints
'/PSIA/Metadata/stream?SoftPoint=true

' Delete Soft Point
'/PSIA/AreaControl/SoftPointList/n
'DELETE

' update softpoint partition
' /PSIA/AreaControl/SoftPointList/n n = softpoint ID
'<SoftPoint>
'<PartitionID>m</PartitionID>
'<SoftPoint>

'Get softpoint list
'/PSIA/AreaControl/SoftPointList
'<SoftPointList>
'<SoftPoint>
'<ID>2</ID>
'<PartitionID>2</PartitionID>
'<Receivers>
'<DeviceID1>0</DeviceID1>
'<Rssi1>0</Rssi1>
'<DeviceID2>0</DeviceID2>
'<Rssi2>0</Rssi2>
'<DeviceID3>6845869</DeviceID3>
'<Rssi3>3</Rssi3>
'<DeviceID4>6117604</DeviceID4>
'<Rssi4>12</Rssi4>
'<DeviceID5>6845861</DeviceID5>
'<Rssi5>3</Rssi5>
'<DeviceID6>6845927</DeviceID6>
'<Rssi6>14</Rssi6>
'<DeviceID7>6845931</DeviceID7>
'<Rssi7>12</Rssi7>
'<DeviceID8>6587694</DeviceID8>
'<Rssi8>63</Rssi8>
'</Receivers>
'</SoftPoint>
'<SoftPoint>
'<ID>3</ID>
'<PartitionID>2</PartitionID>
'<Receivers>
'<DeviceID1>0</DeviceID1>
'<Rssi1>0</Rssi1>
'<DeviceID2>0</DeviceID2>
'<Rssi2>0</Rssi2>
'<DeviceID3>6845869</DeviceID3>
'<Rssi3>3</Rssi3>
'PROPRIETARY DOCUMENT
'03-08-2013 IG1002Av1.5 © Inovonics, 2013 – www.inovonics.com 60
'<DeviceID4>6117604</DeviceID4>
'<Rssi4>12</Rssi4>
'<DeviceID5>6845861</DeviceID5>
'<Rssi5>3</Rssi5>
'<DeviceID6>6845927</DeviceID6>
'<Rssi6>14</Rssi6>
'<DeviceID7>6845931</DeviceID7>
'<Rssi7>12</Rssi7>
'<DeviceID8>6587694</DeviceID8>
'<Rssi8>63</Rssi8>
'</Receivers>
'</SoftPoint>
'</SoftPointList>


Public PartitionID  As Long  'we assign this
Public PartitionDesc As String

Public Num          As Long

Public ID           As Long  ' 6080 assigns this

' these all come from the 6080

Public DeviceID1    As Long  ' decimal serialnumber
Public Rssi1        As Long  '
Public DeviceID2    As Long  ' decimal serialnumber
Public Rssi2        As Long  '
Public DeviceID3    As Long  ' decimal serialnumber
Public Rssi3        As Long  '
Public DeviceID4    As Long  ' decimal serialnumber
Public Rssi4        As Long  '
Public DeviceID5    As Long  ' decimal serialnumber
Public Rssi5        As Long  '
Public DeviceID6    As Long  ' decimal serialnumber
Public Rssi6        As Long  '
Public DeviceID7    As Long  ' decimal serialnumber
Public Rssi7        As Long  '
Public DeviceID8    As Long  ' decimal serialnumber
Public Rssi8        As Long  '

Public Status       As Long



Public TimeString   As String

'<ResponseStatus version="1.0" xmlns:urn="psialliance-org">
'<requestURL>/PSIA/AreaControl/SoftPointList/SoftPoint</requestURL>
'<statusCode>1</statusCode>
'<statusString>201 Created</statusString>
'<id>3</id>
'</ResponseStatus>

'<SoftPointEvent>
'<MetadataHeader>
'<MetaVersion>1.0</MetaVersion>
'<MetaID>Inovonics.com/ACG.System/softpointing</MetaID>
'<MetaSourceID>{C15768C8-E695-4315-A06E-CF49E1409654}</MetaSourceID>
'<MetaSourceLocalID>6475870</MetaSourceLocalID>
'<MetaTime>2012-06-21T23:18:56.8969Z</MetaTime>
'<MetaPriority>6</MetaPriority>
'</MetadataHeader>
'<EventData>
'<Transmitter>
'<ID>6</ID>
'<DeviceID>6853237</DeviceID>
'<DeviceType>EN1223S Pendant</DeviceType>
'<Description>Device 6853237</Description>
'<MID>178</MID>
'</Transmitter>
'<Receivers>
'<DeviceID1>6587694</DeviceID1>
'<Rssi1>65</Rssi1>
'<DeviceID2>6117604</DeviceID2>
'<Rssi2>9</Rssi2>
'<DeviceID3>6845931</DeviceID3>
'<Rssi3>7</Rssi3>
'<DeviceID4>6845938</DeviceID4>
'<Rssi4>34</Rssi4>
'<DeviceID5>6587694</DeviceID5>
'<Rssi5>23</Rssi5>
'<DeviceID6>6845927</DeviceID6>
'<Rssi6></Rssi6>
'<DeviceID7>6845873</DeviceID7>
'<Rssi7>9</Rssi7>
'<DeviceID8>6845927</DeviceID8>
'<Rssi8>13</Rssi8>
'</Receivers>
'</EventData>
'</SoftPointEvent>

Public Function GetUpdateXML() As String



End Function

Public Function GetRegistrationXML() As String
  Dim XML           As String
  ' use with unregisetered soft point
  '/PSIA/AreaControl/SoftPointList
  ' POST
  ' response is:
  '<ResponseStatus version="1.0" xmlns:urn="psialliance-org">
  ' <requestURL>/PSIA/AreaControl/SoftPointList/SoftPoint</requestURL>
  ' <statusCode>1</statusCode>
  ' <statusString>201 Created</statusString>
  ' <id>3</id>
  '</ResponseStatus>
  
  
  
  XML = "<SoftPoint>"
  XML = XML & "<PartitionID>" & PartitionID & "</PartitionID>"
  XML = XML & "<Receivers>"

  XML = XML & "<DeviceID1>" & DeviceID1 & "</DeviceID1>"
  XML = XML & "<Rssi1>" & Rssi1 & "</Rssi1>"

  XML = XML & "<DeviceID2>" & DeviceID2 & "</DeviceID2>"
  XML = XML & "<Rssi2>" & Rssi2 & "</Rssi2>"

  XML = XML & "<DeviceID3>" & DeviceID3 & "</DeviceID3>"
  XML = XML & "<Rssi3>" & Rssi3 & "</Rssi3>"

  XML = XML & "<DeviceID4>" & DeviceID4 & "</DeviceID4>"
  XML = XML & "<Rssi4>" & Rssi4 & "</Rssi4>"

  XML = XML & "<DeviceID5>" & DeviceID5 & "</DeviceID5>"
  XML = XML & "<Rssi5>" & Rssi5 & "</Rssi5>"

  XML = XML & "<DeviceID6>" & DeviceID6 & "</DeviceID6>"
  XML = XML & "<Rssi6>" & Rssi6 & "</Rssi6>"

  XML = XML & "<DeviceID7>" & DeviceID7 & "</DeviceID7>"
  XML = XML & "<Rssi7>" & Rssi7 & "</Rssi7>"

  XML = XML & "<DeviceID8>" & DeviceID8 & "</DeviceID8>"
  XML = XML & "<Rssi8>" & Rssi8 & "</Rssi8>"

  XML = XML & "</Receivers></SoftPoint>"

  GetRegistrationXML = XML

  '<DeviceID1>6587694</DeviceID1>
  '<Rssi1>65</Rssi1>
  '<DeviceID2>6117604</DeviceID2>
  '<Rssi2>9</Rssi2>
  '<DeviceID3>6845931</DeviceID3>
  '<Rssi3>7</Rssi3>
  '<DeviceID4>6845938</DeviceID4>
  '<Rssi4>34</Rssi4>
  '<DeviceID5>6587694</DeviceID5>
  '<Rssi5>23</Rssi5>
  '<DeviceID6>6845927</DeviceID6>
  '<Rssi6></Rssi6>
  '<DeviceID7>6845873</DeviceID7>
  '<Rssi7>9</Rssi7>
  '<DeviceID8>6845927</DeviceID8>
  '<Rssi8>13</Rssi8></Receivers>
  '</Receivers>
  '</SoftPoint>





End Function


Public Function ParseEventXML(ByVal XML As String) As Boolean
  Dim doc           As DOMDocument60
  Dim Node          As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList
  Set doc = New DOMDocument60
  If doc.LoadXML(XML) Then
    TimeString = Format(Now, "HH:NN:SS")
      
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/PartitionID")
    'PartitionID = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID1")
    DeviceID1 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi1")
    Rssi1 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID2")
    DeviceID2 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi2")
    Rssi2 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID3")
    DeviceID3 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi3")
    Rssi3 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID4")
    DeviceID4 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi4")
    Rssi4 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID5")
    DeviceID5 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi5")
    Rssi5 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID6")
    DeviceID6 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi6")
    Rssi6 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID7")
    DeviceID7 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi7")
    Rssi7 = Val(Node.text)

    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/DeviceID8")
    DeviceID8 = Val(Node.text)
    Set Node = doc.selectSingleNode("SoftPointEvent/EventData/Receivers/Rssi8")
    Rssi8 = Val(Node.text)

    Status = 1
    ParseEventXML = True
  End If





End Function

'<SoftPoint>
'<PartitionID>23</PartitionID>
'<Receivers>
'<DeviceID1>6587694</DeviceID1>
'<Rssi1>65</Rssi1>
'<DeviceID2>6117604</DeviceID2>
'<Rssi2>9</Rssi2>
'<DeviceID3>6845931</DeviceID3>
'<Rssi3>7</Rssi3>
'<DeviceID4>6845938</DeviceID4>
'<Rssi4>34</Rssi4>
'<DeviceID5>6587694</DeviceID5>
'<Rssi5>23</Rssi5>
'<DeviceID6>6845927</DeviceID6>
'<Rssi6></Rssi6>
'<DeviceID7>6845873</DeviceID7>
'<Rssi7>9</Rssi7>
'<DeviceID8>6845927</DeviceID8>
'<Rssi8>13</Rssi8></Receivers>
'</Receivers>
'</SoftPoint>


Public Function ParseXML(ByVal XML As String) As Boolean
  Dim Node          As IXMLDOMNode
  Dim subnode       As IXMLDOMNode
  Dim ssnode        As IXMLDOMNode
  Dim NodeList      As IXMLDOMNodeList

  Dim doc           As DOMDocument60
  Set doc = New DOMDocument60



  If doc.LoadXML(XML) Then
    Set Node = doc.firstChild
    If Not Node Is Nothing Then
      For Each subnode In Node.childnodes
        Select Case subnode.baseName
          Case "ID"
            ID = Val(subnode.text)
          Case "PartitionID"
            PartitionID = Val(subnode.text)
          Case "Receivers"
            For Each ssnode In subnode.childnodes
              Select Case ssnode.baseName
                Case "DeviceID1"
                  DeviceID1 = Val(ssnode.text)
                Case "Rssi1"
                  Rssi1 = Val(ssnode.text)
                Case "DeviceID2"
                  DeviceID2 = Val(ssnode.text)
                Case "Rssi2"
                  Rssi2 = Val(ssnode.text)
                Case "DeviceID3"
                  DeviceID3 = Val(ssnode.text)
                Case "Rssi3"
                  Rssi3 = Val(ssnode.text)
                Case "DeviceID4"
                  DeviceID4 = Val(ssnode.text)
                Case "Rssi4"
                  Rssi4 = Val(ssnode.text)
                Case "DeviceID5"
                  DeviceID5 = Val(ssnode.text)
                Case "Rssi5"
                  Rssi5 = Val(ssnode.text)
                Case "DeviceID6"
                  DeviceID6 = Val(ssnode.text)
                Case "Rssi6"
                  Rssi6 = Val(ssnode.text)
                Case "DeviceID7"
                  DeviceID7 = Val(ssnode.text)
                Case "Rssi7"
                  Rssi7 = Val(ssnode.text)
                Case "DeviceID8"
                  DeviceID8 = Val(ssnode.text)
                Case "Rssi8"
                  Rssi8 = Val(ssnode.text)
              End Select
            Next
        End Select
      Next

      ParseXML = True
    End If

  End If

End Function
