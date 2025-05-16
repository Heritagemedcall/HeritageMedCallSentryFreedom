VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DialogicSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Enum Diva
  MAX_CARDS = 3
  MAX_CHANNELS_percard = 8
  MAX_CHANNELS = MAX_CARDS * MAX_CHANNELS_percard
End Enum


Private Divasys       As DIVASDKLib.DivaSystem
Private DivaInst      As DIVASDKLib.DivaInstance
Private DivaDevice()  As DIVASDKLib.DivaDevice

Private mReserved(0 To Diva.MAX_CHANNELS) As Boolean

Private bStarted As Boolean
Private Channels As Long
'Private AllCalls(0 To DIVA.MAX_CHANNELS) As long
'Private AllCalls(0 To DIVA.MAX_CHANNELS) As SingleCallHandler

Private mInstalled(0 To Diva.MAX_CHANNELS) As Long

Private mDivaCalls(0 To Diva.MAX_CHANNELS) As cDivaCall

Private mDeviceCount    As Long
Private mTotalChannels  As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
 
Public Sub UpdateClock()
        Dim j As Long
10      On Error GoTo UpdateClock_Error

20      For j = 1 To Diva.MAX_CHANNELS
30        If mInstalled(j) Then
40          If Not mReserved(j) Then  ' reserved are emergency lines and are handled there
50            mDivaCalls(j).UpdateClock
60          End If
70        End If
80      Next

UpdateClock_Resume:

90      On Error GoTo 0
100     Exit Sub

UpdateClock_Error:

110     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.UpdateClock." & Erl
120     Resume UpdateClock_Resume
End Sub


Public Function GetCallHandlerByID(ByVal ID As Long) As cDivaCall
 

  If Installed(ID) Then
    Set GetCallHandlerByID = mDivaCalls(ID)
  End If
End Function



Public Property Get Installed(ByVal index As Long) As Boolean


10      On Error GoTo Installed_Error

20      If index > 0 And index <= Diva.MAX_CHANNELS Then
30        Installed = mInstalled(index)
40      Else
50        Installed = False
60      End If

Installed_Resume:

70      On Error GoTo 0
80      Exit Property

Installed_Error:

90      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.Installed." & Erl
100     Resume Installed_Resume

End Property


Public Property Get Reserved(ByVal index As Long) As Boolean

10      On Error GoTo Reserved_Error

20      If index > 0 And index <= Diva.MAX_CHANNELS Then
30        Reserved = mReserved(index)
40      Else
50        Reserved = True
60      End If

Reserved_Resume:

70      On Error GoTo 0
80      Exit Property

Reserved_Error:

90      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.Reserved." & Erl
100     Resume Reserved_Resume

End Property
Public Property Let Reserved(ByVal index As Long, ByVal value As Boolean)

10      On Error GoTo Reserved_Error

20      Debug.Print "Reserving Dialogic Channel " & index & " = " & value
30      If index > 0 And index <= Diva.MAX_CHANNELS Then

40        mReserved(index) = value
50      End If

Reserved_Resume:

60      On Error GoTo 0
70      Exit Property

Reserved_Error:

80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.Reserved." & Erl
90      Resume Reserved_Resume

End Property

Public Property Get DeviceCount() As Long
 
' how many cards installed
  DeviceCount = mDeviceCount
End Property

Public Property Get TotalChannels() As Long
 
  TotalChannels = mTotalChannels
End Property

Sub Init()

        Dim Key As Long
        Dim rc As Long
        Dim i As Long
        Dim j As Long
        Dim index As Long
        Dim CallObject As DIVASDKLib.DivaCall
        Dim SDKCode    As String

10      On Error GoTo Init_Error
        
        'dbg "Checking for Diva SDK registration 10"

'20      RC = Win32.RegOpenKey(HKEY_CLASSES_ROOT, "DivaSDK.DivaSystem", Key)

21        rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, "DivaSDK.DivaSystem", 0, KEY_ALL_ACCESS, Key)
        
          'dbg "Checked for Diva SDK registration 21 returned " & RC
'          code = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZX", "")

30      If rc = 1 Then
          'dbg "Diva SDK Not found 30"
          NODIVA = True
40
50        Exit Sub  ' SDK not installed
60      End If


        RegCloseKey Key

        'dbg "Diva SDK Registered 60"

70      If NODIVA Then
          SpecialLog "DialogicSystem NODIVA Reg Key DivaSDK.DivaSystem"
80        Exit Sub
90      End If

100     On Error Resume Next

        'dbg "Initializing Diva SDK 100"

110     Set Divasys = CreateObject("DivaSDK.DivaSystem")   'New DivaSystem  'DIVASDKLib.DivaSystem

120     If Err.Number Then
          SpecialLog "DialogicSystem INIT Diva SDK Error " & Err.Number & " " & Err.Description
130       Exit Sub
140     End If

150     If Not (Divasys Is Nothing) Then
          'dbg "Diva SDK Initialized 150"
160       On Error GoTo 0
170       mDeviceCount = Min(Divasys.numdevices, MAX_CARDS)

          'dbg "Diva Devices Installed 170 Count: " & mDeviceCount
          

180       If (mDeviceCount > 0) Then
190         ReDim DivaDevice(0 To MAX_CARDS)
200         mTotalChannels = Divasys.TotalChannels

           'dbg "Diva Channels Installed " & mTotalChannels

210         If DivaInst Is Nothing Then
220           Set DivaInst = Divasys.CreateInstance(False, TotalChannels, 7, 1024)
              SpecialLog " Divasys.CreateInstance " & TotalChannels

230         End If

240         For j = 1 To mDeviceCount
    
              'dbg " Loop Devices 240 Divasys.GetDevice(" & j & ")"


250           Set DivaDevice(j) = Divasys.GetDevice(j)
260           For i = 1 To DivaDevice(j).Channels
                
                'dbg " Loop Channels 260 DivaDevice(" & j & ").Channels"


270             index = i + ((j - 1) * 8)
280             mInstalled(index) = 1
                dbg " Loop Channels/Create Call 280 DivaInst.CreateCall"
290             Set CallObject = DivaInst.CreateCall()
300             CallObject.device = j
310             CallObject.Channel = i
320             CallObject.AsyncMode = True
330             CallObject.SignalEvents = True
340             Set mDivaCalls(index) = New cDivaCall
350             mDivaCalls(index).SetDivaLine CallObject
360             mDivaCalls(index).device = j
370             mDivaCalls(index).Channel = i
380             mDivaCalls(index).index = index

390           Next
400         Next

            'dbg "Done Allocating cards and lines 400"

410       Else
            SpecialLog "Diva SDK Not Initialized 410"
420         Set Divasys = Nothing
430       End If

440     End If


Init_Resume:

450     On Error GoTo 0
460     Exit Sub

Init_Error:

470     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.Init." & Erl
480     Resume Init_Resume

End Sub
Function GetNextAvailable() As Long
        Dim j As Long
        Dim device As DIVASDKLib.DivaDevice
        Dim LineStatus As Long
        Dim DivaCall As cDivaCall

10      On Error GoTo GetNextAvailable_Error

20      For j = 1 To MAX_CHANNELS
30        If mInstalled(j) Then
40          If (Not mReserved(j)) Then
50            Set DivaCall = mDivaCalls(j)
60            Set device = DivaDevice(mDivaCalls(j).device)
70            LineStatus = device.AnalogLineStatus(mDivaCalls(j).Channel)
80            If LineStatus = DIVASDKLib.AnalogLineIdle Then

90              If Not (DivaCall.Busy) Then
100               GetNextAvailable = j
110               Exit For
120             End If  ' if not busy
130           End If  ' if idle
140         End If  ' if not reserved
150       End If  ' if installed
160     Next


GetNextAvailable_Resume:

170     On Error GoTo 0
180     Exit Function

GetNextAvailable_Error:

190     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.GetNextAvailable." & Erl
200     Resume GetNextAvailable_Resume

End Function


Function GetAnyAvailable() As Long
        Dim j As Long
        Dim device As DIVASDKLib.DivaDevice
        Dim LineStatus As Long
        Dim DivaCall As cDivaCall

10      On Error GoTo GetAnyAvailable_Error

20      For j = 1 To MAX_CHANNELS
30        If mInstalled(j) Then
40          ''''If (Not mReserved(j)) Then
50            Set DivaCall = mDivaCalls(j)
60            Set device = DivaDevice(mDivaCalls(j).device)
70            LineStatus = device.AnalogLineStatus(mDivaCalls(j).Channel)
80            If LineStatus = DIVASDKLib.AnalogLineIdle Then

90              If Not (DivaCall.Busy) Then
100               GetAnyAvailable = j
110               Exit For
120             End If  ' if not busy
130           End If  ' if idle
140         ''''End If  ' if not reserved
150       End If  ' if installed
160     Next


GetAnyAvailable_Resume:

170     On Error GoTo 0
180     Exit Function

GetAnyAvailable_Error:

190     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at DialogicSystem.GetAnyAvailable." & Erl
200     Resume GetAnyAvailable_Resume

End Function


Private Sub Class_Initialize()
  Init
End Sub

Private Sub Class_Terminate()
  Set Divasys = Nothing
End Sub
'Private Function ReadRegString(Key As HKEY_Type, SubKey As String, sName As String, Optional sDefault As String = "") As String
'  Dim hKey            As Long
'  Dim retval          As Long
'  Dim sBuffer         As String
'  Dim slength         As Long
'  Dim DataType        As Long
'  Dim Ptr As Long
'
'  Const BUFFERSIZE = 255
'
'  retval = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)
'
'  If retval <> ERROR_NONE Then
'    ReadRegString = sDefault
'    Exit Function
'  End If
'
'  sBuffer = String$(BUFFERSIZE, vbNullChar)
'  slength = BUFFERSIZE
'  retval = RegQueryValueEx(hKey, sName, 0, DataType, ByVal sBuffer, slength)
'
'  If retval = ERROR_NONE Then
'    If DataType = REG_SZ Or DataType = REG_EXPAND_SZ Then
'      Ptr = InStr(slength - 1, sBuffer, vbNullChar, vbBinaryCompare)
'      If Ptr > 0 Then
'        Ptr = Ptr - 1
'      End If
'      sBuffer = left$(sBuffer, Ptr)
'
'      ReadRegString = sBuffer
'    Else
'      ReadRegString = sDefault
'    End If
'  Else
'    ReadRegString = sDefault
'  End If
'
'  retval = RegCloseKey(hKey)
'
'End Function
'
