Attribute VB_Name = "modIPHLPapi"


Option Explicit

' Declarations needed for GetAdaptersInfo & GetIfTable
Private Const MIB_IF_TYPE_OTHER                   As Long = 1
Private Const MIB_IF_TYPE_ETHERNET                As Long = 6
Private Const MIB_IF_TYPE_TOKENRING               As Long = 9
Private Const MIB_IF_TYPE_FDDI                    As Long = 15
Private Const MIB_IF_TYPE_PPP                     As Long = 23
Private Const MIB_IF_TYPE_LOOPBACK                As Long = 24
Private Const MIB_IF_TYPE_SLIP                    As Long = 28

Private Const MIB_IF_ADMIN_STATUS_UP              As Long = 1
Private Const MIB_IF_ADMIN_STATUS_DOWN            As Long = 2
Private Const MIB_IF_ADMIN_STATUS_TESTING         As Long = 3

Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL  As Long = 0
Private Const MIB_IF_OPER_STATUS_UNREACHABLE      As Long = 1
Private Const MIB_IF_OPER_STATUS_DISCONNECTED     As Long = 2
Private Const MIB_IF_OPER_STATUS_CONNECTING       As Long = 3
Private Const MIB_IF_OPER_STATUS_CONNECTED        As Long = 4
Private Const MIB_IF_OPER_STATUS_OPERATIONAL      As Long = 5

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
Private Const DEFAULT_MINIMUM_ENTITIES            As Long = 32
Private Const MAX_HOSTNAME_LEN                    As Long = 128
Private Const MAX_DOMAIN_NAME_LEN                 As Long = 128
Private Const MAX_SCOPE_ID_LEN                    As Long = 256

Private Const MAXLEN_IFDESCR                      As Long = 256
Private Const MAX_INTERFACE_NAME_LEN              As Long = MAXLEN_IFDESCR * 2
Private Const MAXLEN_PHYSADDR                     As Long = 8

Private Type TIME_t
  aTime As Long
End Type

Private Type IP_ADDRESS_STRING
  IPadrString     As String * 16
End Type

Private Type IP_ADDR_STRING
  AdrNext         As Long
  IpAddress       As IP_ADDRESS_STRING
  IpMask          As IP_ADDRESS_STRING
  NTEcontext      As Long
End Type


Private Type IP_ADAPTER_INFO
  Next As Long
  ComboIndex As Long
  AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
  Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
  MACadrLength        As Long
  MacAddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
  AdapterIndex        As Long
  AdapterType         As Long  ' MSDN Docs say "UInt", but is 4 bytes
  DhcpEnabled         As Long  ' MSDN Docs say "UInt", but is 4 bytes
  CurrentIpAddress    As Long
  IpAddressList       As IP_ADDR_STRING
  GatewayList         As IP_ADDR_STRING
  DhcpServer          As IP_ADDR_STRING
  HaveWins            As Long  ' MSDN Docs say "Bool", but is 4 bytes
  PrimaryWinsServer   As IP_ADDR_STRING
  SecondaryWinsServer As IP_ADDR_STRING
  LeaseObtained       As TIME_t
  LeaseExpires        As TIME_t
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Private Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" (ByRef pdwNumIf As Long) As Long
Private Declare Function GetIfEntry Lib "iphlpapi.dll" (ByRef pIfRow As Any) As Long
Private Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long



Public Function GetAdapters() As Collection

  Dim AdapInfo           As IP_ADAPTER_INFO
  Dim bufLen             As Long
  Dim sts                As Long
  Dim retStr             As String
  Dim numStructs         As Integer
  Dim i                  As Integer
  Dim IPinfoBuf()        As Byte
  Dim srcPtr             As Long
  Dim CurrentIpAddress   As Long
  Dim CurrentIpAddressString As String

  Dim Adapters           As Collection
  Dim Adapter            As cAdapter
  Set Adapters = New Collection


  'Dim numcards As Long
  'Dim AdapterName2  As String


  ' Get size of buffer to allocate
  sts = GetAdaptersInfo(AdapInfo, bufLen)
  If (bufLen = 0) Then
    Exit Function
  End If
  numStructs = bufLen / Len(AdapInfo)
  retStr = numStructs & " Adapter(s):" & vbCrLf

  ' reserve byte buffer & get it filled with adapter information
  ' !!! Don't Redim AdapInfo array of IP_ADAPTER_INFO,
  ' !!! because VB doesn't allocate it contiguous (padding/alignment)

  ReDim IPinfoBuf(0 To bufLen - 1) As Byte
  sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)
  If (sts <> 0) Then Exit Function

  ' Copy IP_ADAPTER_INFO slices into UDT structure
  srcPtr = VarPtr(IPinfoBuf(0))
  For i = 0 To numStructs - 1
    If (srcPtr <> 0) Then

      CopyMemory AdapInfo, ByVal srcPtr, Len(AdapInfo)

      ' Extract Ethernet MAC address
      If (AdapInfo.AdapterType = MIB_IF_TYPE_ETHERNET) Then
        Set Adapter = New cAdapter
        Adapter.MacAddress = MAC2String(AdapInfo.MacAddress)
        Adapter.Description = sz2string(AdapInfo.Description)
        Adapter.ServiceName = sz2string(AdapInfo.AdapterName)  ' this is the GUID that links the adapter
        ' CurrentIpAddress = AdapInfo.CurrentIpAddress
        Adapter.DhcpIPAddress = sz2string(AdapInfo.IpAddressList.IpAddress.IPadrString)  ' this is dynamic info

        Adapters.Add Adapter, Adapter.ServiceName
        'if we don't get a CurrentIpAddressString then we need to get the stored value
        ' also if offline then it's 0.0.0.0


        'currentaddressstring = AdapInfo.IpAddressList
        'retStr = retStr & vbCrLf & "[" & i & "] " & sz2string(AdapInfo.Description) & vbCrLf & vbTab & MAC2String(AdapInfo.MACaddress) & vbCrLf
        'Exit For
      End If

      srcPtr = AdapInfo.Next
    End If
  Next i

  ' Return list of MAC address(es)
  'GetMACs_AdaptInfo = retStr
  Set GetAdapters = Adapters
  Set Adapters = Nothing
End Function


' Convert a byte array containing a MAC address to a hex string
Private Function MAC2String(AdrArray() As Byte) As String
  Dim HexString   As String
  Dim HexByte     As String
  Dim i           As Integer

  For i = 0 To 5
    If (i > UBound(AdrArray)) Then
      HexByte = "00"
    Else
      HexByte = Right("00" & Hex$(AdrArray(i)), 2)
    End If
    HexString = HexString & HexByte
  Next i

  MAC2String = HexString

End Function

Private Function sz2string(ByVal szStr As String) As String
  Dim Ptr As Long
  Ptr = InStr(1, szStr, Chr$(0)) - 1
  Ptr = Max(0, Ptr)
  sz2string = left$(szStr, Ptr)
  
End Function


