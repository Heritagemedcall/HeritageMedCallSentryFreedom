VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAdapters        As Collection

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_NO_MORE_ITEMS = 259&

Private Const REG_OPTION_NON_VOLATILE = 0

Private m_SelectedSection As Long

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_SZ = 1
' Get the key information for this key and
' its subkeys.

Dim doc                  As DOMDocument60

Dim currentNode          As IXMLDOMNode


'Public Function GetNetworkCardNames() As Long  ' count
'  Dim root               As IXMLDOMNode
'  Dim cn                 As IXMLDOMNode
'  Dim cnn                As IXMLDOMNode
'  Dim cnnn               As IXMLDOMNode
'
'  Dim Adapter            As cAdapter
'  'Dim interface          As cInterface
'  Dim docadapters        As DOMDocument60
'
'  Set docadapters = New DOMDocument60
'  Set doc = docadapters
'
'
'  doc.LoadXML "<Root></Root>"
'  Set currentNode = doc.childnodes(0)
'  Set root = doc.childnodes(0)
'  Call GetKeyInfo(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkCards", root)
'
'  Set mAdapters = Nothing
'  Set mAdapters = New Collection
'  For Each cn In root.childnodes
'    If cn.baseName = "Adapter" Then
'      Set Adapter = New cAdapter
'
'      For Each cnn In cn.childnodes
'        Select Case cnn.baseName
'          Case "ID"
'            Adapter.ID = Val(cnn.text & "")
'          Case "ServiceName"
'            Adapter.ServiceName = cnn.text
'          Case "Description"
'            Adapter.Description = cnn.text
'        End Select
'      Next
'      If Len(Adapter.ServiceName) > 0 Then
'        mAdapters.Add Adapter, Adapter.ServiceName
'      End If
'    End If
'  Next

'
'  '  <Root>
'  '    <Adapter>
'  '      <ID>13</ID>
'  '      <ServiceName>{8D030C62-7BEF-4A2A-AE35-E8361EA8FB80}</ServiceName>
'  '      <Description>Dialogic Diva UM-Analog-8 PCIe v1</Description>
'  '    </Adapter>
'  '    <Adapter>
'  '      <ID>14</ID>
'  '      <ServiceName>{2D98E2B5-4263-4001-A0E7-7C57D20F78AA}</ServiceName>
'  '      <Description>USB to Ethernet Adapter</Description>
'  '    </Adapter>
'  '    <Adapter>
'  '      <ID>8</ID>
'  '      <ServiceName>{0650B14C-468C-4AD2-A902-B0EE107591C0}</ServiceName>
'  '      <Description>Realtek PCIe GBE Family Controller</Description>
'  '    </Adapter>
'  '  </Root>
'
'
'  'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkCards
'End Function




Public Function GetAdapterByName(ByVal Description As String) As cAdapter



  Dim Adapter            As cAdapter
  For Each Adapter In mAdapters
    If 0 = StrComp(Adapter.Description, Description, vbTextCompare) Then
      Set GetAdapterByName = Adapter
      Exit For
    End If
  Next


End Function

Public Function GetAdapterByMAC(ByVal MAC As String) As cAdapter
  Dim Adapter            As cAdapter
  
  If InStr(1, MAC, ":") Then
    MAC = Replace(MAC, ":", "")
  ElseIf InStr(1, MAC, "-") Then
    MAC = Replace(MAC, "-", "")
  ElseIf InStr(1, MAC, " ") Then
    MAC = Replace(MAC, " ", "")
  End If
  
  For Each Adapter In mAdapters
    If 0 = StrComp(Adapter.MacAddress, MAC, vbTextCompare) Then
      Set GetAdapterByMAC = Adapter
      Exit For
    End If
  Next


End Function

Public Function GetAdapterByServiceName(ByVal ServiceName As String) As cAdapter
 
  Dim Adapter            As cAdapter
  For Each Adapter In mAdapters
    If 0 = StrComp(Adapter.ServiceName, ServiceName, vbTextCompare) Then
      Set GetAdapterByServiceName = Adapter
      Exit For
    End If
  Next
 
End Function

Public Function GetAdapterByIP(ByVal IP As String) As cAdapter
  Dim Adapter            As cAdapter
  For Each Adapter In mAdapters
    If 0 = StrComp(Adapter.DhcpIPAddress, IP, vbTextCompare) Then
      Set GetAdapterByIP = Adapter
      Exit For
    End If
  Next


End Function

Public Property Get Count() As Long
  Count = mAdapters.Count
End Property
Public Function RefreshAdapters() As Long

        Dim root               As IXMLDOMNode
        Dim cn                 As IXMLDOMNode
        Dim cnn                As IXMLDOMNode
        Dim cnnn               As IXMLDOMNode
        Dim Interfaces         As Collection
        Dim Interface          As cInterface

        Dim docadapters        As DOMDocument60

        'GetNetworkCardNames


  On Error GoTo RefreshAdapters_Error

10      Set mAdapters = GetAdapters()

        '  Set docadapters = New DOMDocument60
        '  Set doc = docadapters
        '
        '
        '  doc.LoadXML "<Root></Root>"
        '  Set currentNode = doc.childnodes(0)
        '  Set root = doc.childnodes(0)
        '
        '  Call GetKeyInfo(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\services\Tcpip\Parameters\adapters", root)
        '
        '  Set Interfaces = New Collection
        '
        '  For Each cn In root.childnodes
        '    If 0 = StrComp(cn.baseName, "Adapter", vbTextCompare) Then
        '      Set Interface = New cInterface
        '      For Each cnn In cn.childnodes
        '        If 0 = StrComp(cnn.baseName, "ID", vbTextCompare) Then
        '          Interface.ID = cnn.text
        '          Interfaces.Add Interface
        '          Exit For
        '        End If
        '      Next
        '
        '    End If
        '  Next
        '
        Dim docinterfaces      As DOMDocument60

20      Set docinterfaces = New DOMDocument60
30      Set doc = docinterfaces
40      doc.LoadXML "<Root></Root>"
50      Set root = doc.childnodes(0)

60      Call GetKeyInfo(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\services\Tcpip\Parameters\interfaces", root)

        Dim Adapter            As cAdapter

70      For Each Adapter In mAdapters
          'Interface.TCPAckFrequency = -1
          'Interface.TCPNoDelay = -1
80        For Each cn In root.childnodes
90          If 0 = StrComp(cn.baseName, "Adapter") Then

100           For Each cnn In cn.childnodes

110             Select Case LCase$(cnn.baseName)
                  Case "id"
120                 If 0 = StrComp(cnn.text, Adapter.ServiceName, vbTextCompare) Then
130                   Adapter.Found = True
140                   For Each cnnn In cn.childnodes
150                     Select Case cnnn.baseName

                          Case "TCPNoDelay"
160                         Adapter.TCPNoDelay = Val(cnnn.text)
170                       Case "TCPAckFrequency"
180                         Adapter.TCPAckFrequency = Val(cnnn.text)
190                       Case "IPAddress"  ' if
200                         Adapter.IpAddress = (cnnn.text)
210                       Case "DhcpIPAddress"  ' if
220                         Adapter.DhcpIPAddress = (cnnn.text)

230                     End Select


240                   Next  ' each cnnn
250                 End If
260             End Select

270           Next  ' each cnn

280         End If
290       Next
300     Next



        ' leave this in in case we want to automatically set the TCPAck Freq and Delay

        Dim WSH                As Object
        Dim RegKey             As String
        Dim rc                 As Long



'310     Set WSH = CreateObject("WScript.Shell")
'
'320     For Each Adapter In mAdapters
'
'330       If Adapter.TCPAckFrequency <> 1 Then
'340         RegKey = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & Adapter.ServiceName & "\TCPAckFrequency"
'350         rc = WSH.RegWrite(RegKey, "1", "REG_DWORD")
'360         If rc = 0 Then
'              'LogToFile "Updated Interface " & Adapter.Description & " TCPAckFrequency to 1"
'370         Else
'              'LogToFile "Update to Interface " & Adapter.Description & " TCPAckFrequency FAIL"
'380         End If
'390       Else
'            'LogToFile "Interface " & interface.ID & " TCPAckFrequency OK"
'400       End If
'
'410       If Adapter.TCPNoDelay <> 1 Then
'420         RegKey = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & Adapter.ServiceName & "\TCPNoDelay"
'430         rc = WSH.RegWrite(RegKey, "1", "REG_DWORD")
'440         If rc = 0 Then
'              'LogToFile "Updated Interface " & Adapter.Description & " TCPNoDelay to 1"
'450         Else
'              'LogToFile "Update to Interface " & Adapter.Description & " TCPNoDelay FAIL"
'460         End If
'470       Else
'            'LogToFile "Interface " & interface.ID & " TCPNoDelay OK"
'480       End If
'
'          '
'490     Next
'
'
'
'500     Set WSH = Nothing


RefreshAdapters_Resume:

  On Error GoTo 0
  Exit Function

RefreshAdapters_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAdapters.RefreshAdapters." & Erl
  Resume RefreshAdapters_Resume

End Function



Private Function GetKeyInfo(ByVal section As Long, ByVal key_name As String, Parent As IXMLDOMNode) As String
  Dim subkeys            As Collection
  Dim subkey_values      As Collection
  Dim subkey_num         As Integer
  Dim subkey_name        As String
  Dim subkey_value       As String
  Dim Length             As Long
  Dim hKey               As Long
  Dim txt                As String
  Dim subkey_txt         As String
  Dim value_num          As Long
  Dim value_name_len     As Long
  Dim value_name         As String
  Dim Reserved           As Long
  Dim value_type         As Long
  Dim value_string       As String
  Dim value_data(1 To 1024) As Byte
  Dim value_data_len     As Long
  Dim i                  As Integer

  Dim element            As IXMLDOMElement

  Dim element2           As IXMLDOMElement

  Set subkeys = New Collection
  Set subkey_values = New Collection

  ' Open the key.
  If RegOpenKeyEx(section, key_name, 0&, KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS, hKey) <> ERROR_SUCCESS Then
    MsgBox "Error opening key. Check Permissions"
    Exit Function
  End If

  ' Enumerate the key's values.
  value_num = 0



  Do

    value_name_len = 1024
    value_name = Space$(value_name_len)
    value_data_len = 1024

    If RegEnumValue(hKey, value_num, value_name, value_name_len, 0, value_type, value_data(1), value_data_len) <> ERROR_SUCCESS Then Exit Do

    value_name = left$(value_name, value_name_len)

    value_string = ""

    Select Case value_type
      Case REG_BINARY

      Case REG_DWORD
        value_string = "&H" & Format$(Hex$(value_data(4)), "00") & Format$(Hex$(value_data(3)), "00") & Format$(Hex$(value_data(2)), "00") & Format$(Hex$(value_data(1)), "00")
        'Debug.Print "REG_DWORD "; value_string
      Case REG_DWORD_BIG_ENDIAN

      Case REG_DWORD_LITTLE_ENDIAN

      Case REG_EXPAND_SZ

      Case REG_FULL_RESOURCE_DESCRIPTOR

      Case REG_LINK

      Case REG_MULTI_SZ

        For i = 1 To value_data_len - 1
          If value_data(i) <> 0 Then
            value_string = value_string & Chr$(value_data(i))
          End If
        Next i
        Debug.Print "REG_MULTI_SZ "; value_string


      Case REG_NONE

      Case REG_RESOURCE_LIST

      Case REG_RESOURCE_REQUIREMENTS_LIST

      Case REG_SZ

        For i = 1 To value_data_len - 1
          value_string = value_string & Chr$(value_data(i))
        Next i

        Debug.Print "REG_SZ "; value_string
    End Select

    Set element = doc.createElement(value_name)
    element.text = value_string
    Parent.appendChild element

    value_num = value_num + 1
  Loop

  ' Enumerate the subkeys.
  subkey_num = 0
  Do
    ' Enumerate subkeys until we get an error.
    Length = 256
    subkey_name = Space$(Length)
    If RegEnumKey(hKey, subkey_num, subkey_name, Length) <> ERROR_SUCCESS Then Exit Do
    subkey_num = subkey_num + 1
    subkey_name = left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)





    Set element = doc.createElement("Adapter")
    'element.Text = subkey_name
    Set element2 = doc.createElement("ID")

    Parent.appendChild element

    element2.text = subkey_name
    element.appendChild element2

    subkeys.Add subkey_name

    ' Get the subkey's value.
    Length = 256
    subkey_value = Space$(Length)
    If RegQueryValue(hKey, subkey_name, subkey_value, Length) <> ERROR_SUCCESS Then
      subkey_values.Add "Error"
    Else
      ' Remove the trailing null character.
      subkey_value = left$(subkey_value, Length - 1)
      subkey_values.Add subkey_value
    End If

    subkey_txt = GetKeyInfo(section, key_name & "\" & subkeys(subkey_num), element)

  Loop

  ' Close the key.
  If RegCloseKey(hKey) <> ERROR_SUCCESS Then
    MsgBox "Error closing key."
  End If

  GetKeyInfo = txt
End Function



Public Function RefreshAdaptersOld() As Long
  ' doesn't work if adapter is unplugged

'  Dim WMI As Object, NICs As Object, NIC As Object
'  Dim adapter            As cAdapter
'
'  On Error Resume Next
'
'  Set Adapters = New Collection
'
'  Set WMI = GetObject("winmgmts:\\.\root\cimv2")
'  Set NICs = WMI.ExecQuery("SELECT * FROM " & "Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
'
'  For Each NIC In NICs
'
'    Err.Clear
'    Set adapter = New cAdapter
'    adapter.Address = NIC.IpAddress(0)
'    adapter.MacAddress = NIC.MacAddress
'    adapter.Description = NIC.Description
'    If Err.Number = 0 Then
'      Adapters.Add adapter
'    End If
'
'  Next
'
'  RefreshAdapters = Adapters.Count


End Function



Private Sub Class_Initialize()
  Set mAdapters = New Collection
End Sub

Private Sub Class_Terminate()
  
  Set mAdapters = Nothing
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mAdapters.[_NewEnum]
End Function

Public Function Item(ByVal index As Variant) As cAdapter
Attribute Item.VB_UserMemId = 0
  Set Item = mAdapters.Item(index)
End Function


Public Property Get Adapters() As Collection
' make this go away
  Set Adapters = mAdapters

End Property

