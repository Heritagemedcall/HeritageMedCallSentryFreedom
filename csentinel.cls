VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSentinel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

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

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String) As Long

Private Declare Function RegCreatekey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    lpcbName As Long, _
    lpReserved As Long, _
    ByVal lpClass As String, _
    lpcbClass As Long, _
    lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Byte, _
    lpcbData As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Main Registry Keys
'------------------------------------------------------------------------



Private mEvaluationID       As String
Private mMachineID          As String
Private mRegistrationID     As String
Private mEvaluationID2      As String

Private mFileMask           As String

Private mRegistryKey        As String
Private mRegistrySubKey     As String

Private mTrialRuns          As Long
Private mTrialDays          As Long

Private mRememberMID        As Boolean
Private mUseMacAddress      As Boolean

Private mPassword           As String

Private mRental             As Boolean
Private mRentalNoticeDays   As Long

Private mUserLevelSecurity  As Boolean

Private mSentinel           As SEN10L2.Registration

Public Function DoRegistration(ByVal CustomerCode As String, ByVal RegistrationCode As String) As Long
  'CheckRegistration
  ' new for 11/21/06, adding RESET to Clear all reg info

  Dim encrypter          As clsDES
  Dim Value              As Double
  Dim ReturnValue        As String
  Dim SaveRegCode        As String
  Dim TheDate            As Long

  DoRegistration = 0
  SaveRegCode = RegistrationCode

  If Len(CustomerCode) > 0 And Len(RegistrationCode) > 0 Then
    Set encrypter = New clsDES
    ReturnValue = encrypter.DecryptString(RegistrationCode, CustomerCode, True)
    Value = Val(ReturnValue)
    Value = Fix(Value)
    ' get the value, and if it's non-zero, we're good to go, so write the encrypted string to the registry
    If MASTER Then
      If Value > 0 Then
        DoRegistration = Value
        Value = WriteRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZX", SaveRegCode)
      ElseIf left(ReturnValue, 5) = "RESET" Then
        If Len(ReturnValue) > 6 Then
          TheDate = Val(MID(ReturnValue, 7))
          If Abs(TheDate - DateValue(Now)) < 3 Then  ' allow 3 days on either side for reset, otherwise, don't allow
            Value = WriteRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZX", "0")
            DeleteRegistration
            Set mSentinel = Nothing
            Init
          End If
        End If
      End If
    Else                       ' is not master
      If Value = -1 Then
        ' this is a remote lic
        DoRegistration = Value
        Value = WriteRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZR", SaveRegCode)
      ElseIf left(ReturnValue, 5) = "RESET" Then
        If Len(ReturnValue) > 6 Then
          TheDate = Val(MID(ReturnValue, 7))
          If Abs(TheDate - DateValue(Now)) < 3 Then  ' allow 3 days on either side for reset, otherwise, don't allow
            Value = WriteRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZR", "0")
            DeleteRegistration
            Set mSentinel = Nothing
            Init
          End If
        End If
      End If
    End If
  Else
    Beep
  End If                       '  Len(CustomerCode)


End Function


'Public Function DoRegistration(ByVal CustomerCode As String, ByVal RegistrationCode As String) As Long
''CheckRegistration
'  Dim encrypter As clsDES
'  Dim Value As Double
'  Dim SaveRegCode As String
'  DoRegistration = 0
'  SaveRegCode = RegistrationCode
'  If Len(CustomerCode) And Len(RegistrationCode) Then
'    Set encrypter = New clsDES
'    Value = Val(encrypter.DecryptString(RegistrationCode, CustomerCode, True))
'    Value = Fix(Value)
'    ' get the value, and if it's non-zero, we're good to go, so write the encrypted string to the registry
'    If Value > 0 Then
'      DoRegistration = Value
'      Value = WriteRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZX", SaveRegCode)
'    End If
'  Else
'    Beep
'  End If
'
'End Function
Public Function GetDeviceCount() As Long
  Dim code               As String

  Dim Value              As Double
  Dim encrypter          As clsDES
  Set encrypter = New clsDES

  If MASTER Then
    code = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZX", "")
  Else
    code = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\" & RegistryKey & "\" & RegistrySubKey, "ZR", "")
  End If
  Value = Val(encrypter.DecryptString(code, mMachineID, True))

  GetDeviceCount = Fix(Value)

End Function

Private Sub Class_Initialize()
  Init
End Sub
Private Sub Init()
  Set mSentinel = New SEN10L2.Registration
  Filemask = "HERITAGE"
  Filemask = G_REG_FILEMASK
  
  RegistryKey = "HeritageMedcall"  'App.CompanyName   ' default
  RegistryKey = G_REG_KEY   'App.CompanyName   ' default
  
  RegistrySubKey = "Freedom2"  'App.Title      ' default
  RegistrySubKey = G_REG_SUB_KEY
  'RememberMID = True ' lock it to this machine.

  TrialRuns = 30
  TrialDays = 15
  RentalNoticeDays = 15

  'mSentinel.LogFile = App.Path & "\serror.log"

  mSentinel.TrialDays = CStr(TrialDays)
  mSentinel.TrialRuns = CStr(TrialRuns)

  mSentinel.RememberMID = RememberMID
  mSentinel.UseMacAddress = UseMacAddress

  mSentinel.ApplicationRental = Rental
  mSentinel.RentalNoticeDays = RentalNoticeDays

  mSentinel.Filemask = Filemask

  mSentinel.RegistrySubKey = RegistryKey & "\" & RegistrySubKey

  mSentinel.UserLevelSecurity = UserLevelSecurity
  CheckRegistration

End Sub



Public Function CheckRegistration() As Long
  Dim rc As Long

  rc = mSentinel.CheckRegistration()
  mMachineID = mSentinel.MachineID
  mRegistrationID = mSentinel.RegistrationID
  mEvaluationID = mSentinel.EvaluationID
  mEvaluationID2 = mSentinel.EvaluationID2

  CheckRegistration = rc

End Function


Private Sub Class_Terminate()
  Set mSentinel = Nothing
End Sub

Public Property Get DaysLeft() As Long
  DaysLeft = Val(mSentinel.DaysLeft)
End Property

Public Function DeleteRegistration() As Boolean
  Dim rc As Long
  rc = mSentinel.CheckRegistration()

  mMachineID = mSentinel.MachineID
  
  mRegistrationID = mSentinel.RegistrationID
  mEvaluationID = mSentinel.EvaluationID
  mEvaluationID2 = mSentinel.EvaluationID2

  DeleteRegistration = mSentinel.DeleteRegistration(Password)

End Function

Public Property Get Evaluating() As Boolean
  Evaluating = (2 And GetStatus()) = 2
End Property

Public Property Get EvaluationID() As String
  EvaluationID = mEvaluationID
End Property

Public Property Let EvaluationID(ByVal Value As String)
  mEvaluationID = Value
End Property

Public Property Get EvaluationID2() As String
  EvaluationID2 = mEvaluationID2
End Property

Public Property Let EvaluationID2(ByVal Value As String)
  mEvaluationID2 = Value
End Property

Public Property Get ExpirationDate() As Date
  If IsDate(mSentinel.RentalExpirationDate) Then
    ExpirationDate = CDate(mSentinel.RentalExpirationDate)
  End If
End Property

'Public Property Get Expired() As Boolean
'    Expired = (GetStatus() = 0)
'End Property

Public Property Get Filemask() As String
  Filemask = mFileMask
End Property

Public Property Let Filemask(ByVal Value As String)
  mFileMask = Value
End Property

Public Property Get FirstRunDate() As Date
  If IsDate(mSentinel.FirstRunDate) Then
    FirstRunDate = CDate(mSentinel.FirstRunDate)
  End If
End Property

Public Function GetLastError() As String
  GetLastError = mSentinel.Error
End Function

Public Function GetStatus() As Long
  GetStatus = mSentinel.CheckRegistration()
End Function

Public Property Get IsRegistered() As Boolean
  IsRegistered = (1 And GetStatus()) = 1
End Property

Public Property Get LastRunDate() As Date
  If IsDate(mSentinel.LastRunDate) Then
    LastRunDate = CDate(mSentinel.LastRunDate)
  End If
End Property

Public Property Get MachineID() As String
  MachineID = mSentinel.MachineID
End Property

Public Property Let MachineID(ByVal Value As String)
  mMachineID = Value
End Property

Public Property Get NeverExpires() As Boolean
  NeverExpires = mSentinel.NO_EXPIRATION = "99999"
End Property

Public Property Get Password() As String
  Password = left$(mSentinel.MachineID, 5) & LCase$(left$(Filemask, 5)) & Format$(CStr(Date), "yyyymmdd")
End Property

Public Function Register(ByVal RegistrationID As String) As Boolean
  Register = mSentinel.Register(RegistrationID)
End Function

Public Function RegisterNoEvalPeriod() As Boolean
  Dim rc As Long
  rc = mSentinel.CheckRegistration()


  mMachineID = mSentinel.MachineID
  mRegistrationID = mSentinel.RegistrationID
  mEvaluationID = mSentinel.EvaluationID
  mEvaluationID2 = mSentinel.EvaluationID2
  RegisterNoEvalPeriod = mSentinel.Register("", Password)

End Function

Public Property Get RegistrationID() As String
  RegistrationID = mRegistrationID
End Property

Public Property Let RegistrationID(ByVal Value As String)
  mRegistrationID = Value
End Property

Public Property Get RegistryKey() As String
  RegistryKey = mRegistryKey
End Property

Public Property Let RegistryKey(ByVal Value As String)
  mRegistryKey = Value
End Property

Public Property Get RegistrySubKey() As String
  RegistrySubKey = mRegistrySubKey
End Property

Public Property Let RegistrySubKey(ByVal Value As String)
  mRegistrySubKey = Value
End Property

Public Property Get RememberMID() As Boolean
  RememberMID = mSentinel.RememberMID
End Property

Public Property Let RememberMID(ByVal Value As Boolean)
  mSentinel.RememberMID = Value
End Property

Public Property Get Rental() As Boolean
  Rental = mSentinel.ApplicationRental
End Property

Public Property Let Rental(ByVal Value As Boolean)
  mSentinel.ApplicationRental = Value
End Property

Public Property Get RentalExpirationDate() As Date
  If IsDate(mSentinel.RentalExpirationDate) Then
    RentalExpirationDate = CDate(mSentinel.RentalExpirationDate)
  End If
End Property

Public Property Get RentalNoticeDays() As Long
  RentalNoticeDays = Val(mSentinel.RentalNoticeDays)
End Property

Public Property Let RentalNoticeDays(ByVal Value As Long)
  mSentinel.RentalNoticeDays = CStr(Value)
End Property

Public Property Get RunsLeft() As Long
  RunsLeft = Val(mSentinel.RunsLeft)
End Property

Property Get STATUS_ERROR() As Integer
  STATUS_ERROR = -1
End Property

Property Get STATUS_EVALUATING() As Integer
  STATUS_EVALUATING = 2
End Property

Property Get STATUS_EXPIRED() As Integer
  STATUS_EXPIRED = 0
End Property

Property Get STATUS_INVALID_RID() As Integer
  STATUS_INVALID_RID = 3
End Property

Property Get STATUS_REGISTERED() As Integer
  STATUS_REGISTERED = 1
End Property

Public Property Get TotalDays() As Long
  TotalDays = Val(mSentinel.TotalDays)
End Property

Public Property Get TotalRuns() As Long
  TotalRuns = Val(mSentinel.TotalRuns)
End Property

Public Property Get TrialDays() As Long
  TrialDays = Val(mSentinel.TrialDays)
End Property

Public Property Let TrialDays(ByVal Value As Long)
  mSentinel.TrialDays = CStr(Value)
End Property

Public Property Get TrialRuns() As Long
  TrialRuns = Val(mSentinel.TrialRuns)
End Property

Public Property Let TrialRuns(ByVal Value As Long)
  mSentinel.TrialRuns = CStr(Value)
End Property

Public Property Get UseMacAddress() As Boolean
  UseMacAddress = mSentinel.UseMacAddress
End Property

Public Property Let UseMacAddress(ByVal Value As Boolean)
  mSentinel.UseMacAddress = Value
End Property

Public Property Get UserLevelSecurity() As Boolean
  UserLevelSecurity = mSentinel.UserLevelSecurity
End Property

Public Property Let UserLevelSecurity(ByVal Value As Boolean)
  mSentinel.UserLevelSecurity = Value
End Property

Public Function ValidateEID(ByVal ID As String) As Boolean
  ValidateEID = mSentinel.ValidateEvaluationID(ID)
End Function

Public Function ValidateMID(ByVal ID As String) As Boolean
  ValidateMID = mSentinel.ValidateMachineID(ID)
End Function

Public Function ValidateRID(ByVal ID As String) As Boolean
  ValidateRID = mSentinel.ValidateRegistrationID(ID)
End Function

Public Property Get Version() As String
  Version = mSentinel.Version
End Property

Public Function DaysBeforeExpire() As String
  Dim DaysRemaining As String
  DaysRemaining = DateDiff("d", Now, GetExpirationDate)
  DaysBeforeExpire = DaysRemaining
End Function

'  On Error Resume Next
'  Value = ReadRegString(HKEY_LOCAL_MACHINE, REGKEY, REGSUBKEY, "")

'  keyfound = WriteRegString(HKEY_LOCAL_MACHINE, REGKEY, REGSUBKEY, Value)

'Public Function SetExpirationDate(ByVal ExpirationDate As Date) As Boolean
'  Dim Value As String
'  Value = encrypter.EncryptString(Format(ExpirationDate, "mm-dd-yyyy"), Sentinel.MachineID, True)
'  SetExpirationDate = WriteRegString(HKEY_LOCAL_MACHINE, Sentinel.RegistryKey, Sentinel.RegistrySubKey, Value)
'
'End Function

Public Function Expired() As Boolean
' check for rollback

  If mSentinel.LastRunDate > DateAdd("d", 2, Now) Then
    Expired = True
  ElseIf Evaluating = True Then
    Expired = False
  ElseIf DaysRemaining < 0 Then
    Expired = True
  End If
End Function
Private Function DaysRemaining() As Long
'DaysRemaining = DateDiff("d", Now, GetExpirationDate)
  DaysRemaining = mSentinel.DaysLeft
End Function

Private Function GetExpirationDate() As Date
  Dim Value As String

  GetStatus

End Function

Public Function WriteRegString(Key As HKEY_Type, SubKey As String, sName As String, sData As String) As Boolean
  Dim hKey        As Long
  Dim retval      As Long
  Dim deposition  As Long
  Dim secattr     As SECURITY_ATTRIBUTES

  secattr.nLength = Len(secattr)
  secattr.lpSecurityDescriptor = 0
  secattr.bInheritHandle = 1

  retval = RegCreateKeyEx(Key, SubKey, ByVal 0&, "", ByVal 0&, KEY_WRITE, secattr, hKey, deposition)
  If retval <> ERROR_NONE Then
    WriteRegString = False
    Exit Function
  End If

  retval = RegSetValueEx(hKey, sName, 0, REG_SZ, ByVal sData, Len(sData))

  If retval <> ERROR_NONE Then
    WriteRegString = False
    Exit Function
  End If

  retval = RegCloseKey(hKey)

  WriteRegString = True

End Function

Public Function ReadRegString(Key As HKEY_Type, SubKey As String, sName As String, Optional sDefault As String = "") As String
  Dim hKey            As Long
  Dim retval          As Long
  Dim sBuffer         As String
  Dim slength         As Long
  Dim DataType        As Long
  Dim Ptr As Long
  
  Const BUFFERSIZE = 255

  retval = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)

  If retval <> ERROR_NONE Then
    ReadRegString = sDefault
    Exit Function
  End If

  sBuffer = String$(BUFFERSIZE, vbNullChar)
  slength = BUFFERSIZE
  retval = RegQueryValueEx(hKey, sName, 0, DataType, ByVal sBuffer, slength)

  If retval = ERROR_NONE Then
    If DataType = REG_SZ Or DataType = REG_EXPAND_SZ Then
      Ptr = InStr(slength - 1, sBuffer, vbNullChar, vbBinaryCompare)
      If Ptr > 0 Then
        Ptr = Ptr - 1
      End If
      sBuffer = left$(sBuffer, Ptr)
      
      ReadRegString = sBuffer
    Else
      ReadRegString = sDefault
    End If
  Else
    ReadRegString = sDefault
  End If

  retval = RegCloseKey(hKey)

End Function



