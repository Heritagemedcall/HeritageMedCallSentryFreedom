VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cESDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const TEMPERATURE_NONE = 0
Const TEMPERATURE_RISE = 1
Const TEMPERATURE_FALL = 2

' Temperature things for alarms
' << From DB

Public Configurationstring As String

Private mZoneID     As Long  ' 6080 system
Public Zone         As cZoneInfo
Public Validated    As Long

Public IsSPDevice   As Long  ' 0 or 1
Public IsRef        As Long  ' 0 or 1
Public Checkin6080  As Long  ' in seconds  3x60x custom checkin time

Public IsMissing    As Long

Public EnableTemperature As Integer      ' 0 disabled, 1 on Rise, 2 on Fall
Public EnableTemperature_A As Integer


Public LowSet       As Single
Public LowSet_A     As Single

Public HiSet        As Single
Public HiSet_A      As Single

'>> END From DB

Public ReachedSet   As Integer           ' dynamic
Public ReachedSet_A As Integer
Public ReachedSet_B As Integer

'Private mIsTemperatureDev     As Integer ' 1 = true, 0 = false , from devicetype

Private mTempAlarm  As Integer
Private mTempAlarm_A As Integer


'Public LastTemperature       As Single   ' mebbe not needed
'Public LastTemperature_A     As Single   ' mebbe not needed

Private mTemperature As Single
Private mTemperature_A As Single



' How it works:

' if EnableTemperature is on Rise (1)
' must reach LowSet to set ReachedSet = 1
' Alarm is generated when ReachedSet = 1 AND Temperature0 >= HiSet


' if EnableTemperature is on Fall (2)
' must reach HighSet0 to set ReachedSet = 1
' Alarm is generated when ReachedSet = 1 AND Temperature <= LowSet



' end temperature variables


Private RecentRSS   As Collection

' Ignore will toss all packets for this device
' does not show up as stray, nor in alarms
Public Ignored      As Integer

Private LastLowBatt As Date
Private LowBattPending As Boolean

' new for more debugging of repeaters

Public NID          As Integer
Private mFirstHop   As String
Private mLastLevel  As Integer
Public LastMargin   As Integer
Public LastConfigResponse As Date
Public LastSeen     As Date

Public Layer        As Integer
Private mJamCount   As Long


' singletons
Public NumInputs    As Long       '  one or two (or three?) alarms

Public IgnoreTamper As Integer
Public UseTamperAsInput As Integer


Public Serial       As String
Public Model        As String

Public CheckinFail  As Long

Public MIDPTI       As Long
Public CLSPTI       As Long

Public MID          As Long
Public CLS          As Long
Private mPTI        As Long

Public locKW        As String
Public RoomFlags    As Long



Private mDecimalSerial As Long


Public IsLatching   As Integer  ' 0 or 1 ' not used
Public IsPortable   As Integer  ' 0 or 1

Public DeviceID     As Long
Public Description  As String

Public LastTamper   As Date        ' date/time of last restore
Public LastTamperRestore As Date  ' date/time of last restore
Private mLastSupervise As Date  ' date/time of last Supervise
Public LastTest     As Date      ' date/time of last test

Public LastLocationText As String

Public NoCheckin    As Integer

Public IsLocator    As Integer
Public TimeStamp    As Long       ' from TimeGetTime

Private mTamper     As Integer         ' 0 or 1
Private mDead       As Integer

Public NoTamper     As Integer        ' ingnores tamper bit

Private mBattery    As Integer  ' 0 or 1
Public STAT         As Long
Public LineLoss     As Integer        ' 0 or 1
Private mIDM        As Long
Private mIDL        As Long           ' 6080 <none>, Mobile, Fixed Device, Sorft Pointer

Private mAutoClear  As Long         ' in seconds for all devicetypes, not just serial devices
' Serial Autoclear overrides this!!!

' One For Each Input

Public Announce     As String
Public Announce_A   As String
Public Announce_B   As String

Public LastAlarm    As Date       ' date/time of last alarm
Public LastAlarm_A  As Date         ' date/time of last alarm
Public LastAlarm_B  As Date         ' date/time of last alarm

Private mLastRestore As Date  ' date/time of last restore
Private MLastRestore_A As Date  ' date/time of last restore
Private mLastRestore_B As Date   ' date/time of last restore

Public alarm        As Integer      ' 0 or 1
Public Alarm_A      As Integer        ' 0 or 1
Public Alarm_B      As Integer

Public Alert        As Integer      ' 0 or 1
Public Alert_A      As Integer        ' 0 or 1
Public Alert_B      As Integer        ' 0 or 1

Public Acknowleged  As Integer  ' 0 or 1
Public Acknowleged_A As Integer  ' 0 or 1
Public Acknowleged_B As Integer  ' 0 or 1

Public Silenced     As Integer         ' 0 or 1
Public Silenced_A   As Integer  ' 0 or 1
Public Silenced_B   As Integer  ' 0 or 1

Public DisableStart As Integer  ' start of disabled period (0 to 23 hrs)
Public DisableEnd   As Integer  ' end of disabled period   (0 to 23 hrs)

Public DisableStart_A As Integer  ' start of disabled period (0 to 23 hrs)
Public DisableEnd_A As Integer  ' end of disabled period   (0 to 23 hrs)

Public DisableStart_B As Integer  ' start of disabled period (0 to 23 hrs)
Public DisableEnd_B As Integer  ' end of disabled period   (0 to 23 hrs)


Public DisableStart2 As Integer  ' start of disabled period (0 to 23 hrs)
Public DisableEnd2  As Integer  ' end of disabled period   (0 to 23 hrs)

Public DisableStart2_A As Integer  ' start of disabled period (0 to 23 hrs)
Public DisableEnd2_A As Integer  ' end of disabled period   (0 to 23 hrs)

Public DisableStart2_B As Integer  ' start of disabled period (0 to 23 hrs)
Public DisableEnd2_B As Integer  ' end of disabled period   (0 to 23 hrs)



Public Pause        As Long   ' seconds between paging
Public Pause_A      As Long     ' seconds between paging
Public Pause_B      As Long     ' seconds between paging

Public IsAway       As Integer
Public IsAway_A     As Integer
Public IsAway_B     As Integer

Public ClearByReset As Integer  ' 0 or 1


' paging repeats (repeatuntil <> 0 then repeat until cancelled)
Public Repeats      As Long
Public RepeatUntil  As Long

Public Repeats_A    As Long
Public RepeatUntil_A As Long

Public Repeats_B    As Long
Public RepeatUntil_B As Long



Public Owner        As String
Public ResidentID   As Long  ' if 0 then unassigned to a resident. Resident's roomID is vectored off resident record!
Public RoomID       As Long  ' if 0 then not assigned to a room

'Public Owner_A      As String  ' not used
'Public ResidentID_A As Long  ' not used ' if 0 then unassigned to a resident. Resident's roomID is vectored off resident record!
'Public RoomID_A     As Long  ' not used ' if 0 then not assigned to a room


Public Phone        As String
Public Notes        As String

'Public Phone_A      As String
'Public Notes_A      As String


Public AssurBit     As Integer ' Set to 1 for next assur scan
Public AssurInput   As Integer

Public AssurBit_A   As Integer  ' not used
Public AssurBit_B   As Integer  ' not used


Public AssurSecure  As Integer
Public AssurSecure_A As Integer
Public AssurSecure_B As Integer

Public AlarmMask    As Integer
Public AlarmMask2   As Integer

Public AlarmMask_A  As Integer
Public AlarmMask2_A As Integer

Public AlarmMask_B  As Integer
Public AlarmMask2_B As Integer


Public UseAssur     As Integer
Public UseAssur2    As Integer

Public UseAssur_A   As Integer  ' not used
Public UseAssur2_A  As Integer  ' not used

Public UseAssur_B   As Integer  ' not used
Public UseAssur2_B  As Integer  ' not used


Public Custom       As String  ' Custom device-type Name

'primary button


Public OG1          As Long  ' these are groups for primary button
Public OG2          As Long
Public OG3          As Long
Public OG4          As Long
Public OG5          As Long
Public OG6          As Long

Public NG1          As Long
Public NG2          As Long
Public NG3          As Long
Public NG4          As Long
Public NG5          As Long
Public NG6          As Long

Public GG1          As Long
Public GG2          As Long
Public GG3          As Long
Public GG4          As Long
Public GG5          As Long
Public GG6          As Long


Public OG1D         As Long  ' these are delays for primary button
Public OG2D         As Long
Public OG3D         As Long
Public OG4D         As Long
Public OG5D         As Long
Public OG6D         As Long


Public NG1D         As Long
Public NG2D         As Long
Public NG3D         As Long
Public NG4D         As Long
Public NG5D         As Long
Public NG6D         As Long


Public GG1D         As Long
Public GG2D         As Long
Public GG3D         As Long
Public GG4D         As Long
Public GG5D         As Long
Public GG6D         As Long


' button 2

' groups
Public OG1_A        As Long
Public OG2_A        As Long
Public OG3_A        As Long
Public OG4_A        As Long
Public OG5_A        As Long
Public OG6_A        As Long


Public NG1_A        As Long
Public NG2_A        As Long
Public NG3_A        As Long
Public NG4_A        As Long
Public NG5_A        As Long
Public NG6_A        As Long

Public GG1_A        As Long
Public GG2_A        As Long
Public GG3_A        As Long
Public GG4_A        As Long
Public GG5_A        As Long
Public GG6_A        As Long


' delays

Public OG1_AD       As Long
Public OG2_AD       As Long
Public OG3_AD       As Long
Public OG4_AD       As Long
Public OG5_AD       As Long
Public OG6_AD       As Long


Public NG1_AD       As Long
Public NG2_AD       As Long
Public NG3_AD       As Long
Public NG4_AD       As Long
Public NG5_AD       As Long
Public NG6_AD       As Long

Public GG1_AD       As Long
Public GG2_AD       As Long
Public GG3_AD       As Long
Public GG4_AD       As Long
Public GG5_AD       As Long
Public GG6_AD       As Long


' button 3 (or tamper as alarm)

' groups
Public OG1_B        As Long
Public OG2_B        As Long
Public OG3_B        As Long
Public OG4_B        As Long
Public OG5_B        As Long
Public OG6_B        As Long


Public NG1_B        As Long
Public NG2_B        As Long
Public NG3_B        As Long
Public NG4_B        As Long
Public NG5_B        As Long
Public NG6_B        As Long

Public GG1_B        As Long
Public GG2_B        As Long
Public GG3_B        As Long
Public GG4_B        As Long
Public GG5_B        As Long
Public GG6_B        As Long


' delays

Public OG1_BD       As Long
Public OG2_BD       As Long
Public OG3_BD       As Long
Public OG4_BD       As Long
Public OG5_BD       As Long
Public OG6_BD       As Long


Public NG1_BD       As Long
Public NG2_BD       As Long
Public NG3_BD       As Long
Public NG4_BD       As Long
Public NG5_BD       As Long
Public NG6_BD       As Long

Public GG1_BD       As Long
Public GG2_BD       As Long
Public GG3_BD       As Long
Public GG4_BD       As Long
Public GG5_BD       As Long
Public GG6_BD       As Long





Public SendCancel   As Integer  ' 0 or 1
Public SendCancel_A As Integer  ' 0 or 1
Public SendCancel_B As Integer  ' 0 or 1

Private mRoom       As String
'Public Room_A       As String  ' not used

Public NameFirst    As String
'Public NameFirst_A  As String  ' not used
Public NameLast     As String
'Public NameLast_A   As String  ' not used

Public Building     As String

Private mSupervisePeriod As Long  ' in minutes

' for external serial device
Public SerialTapProtocol As Integer
Public SerialSkip   As Long
Public SerialMessageLen As Long
Public SerialAutoClear As Integer
Public SerialPort   As Integer
Public SerialBaud   As Long
Public Serialbits   As Integer
Public SerialParity As String
Public SerialStopbits As String
Public SerialInclude As String
Public SerialExclude As String
Public SerialFlow   As Integer
Public SerialSettings As String
Public SerialEOLChar As Integer
Public SerialPreamble As String

Public SerialAlarmID As Long

Public Function GetLocKW(ByVal searchstring As String, ByVal LocKWords As String) As String
  Dim j                  As Long
  Dim LocKWArray()       As String

  searchstring = Trim$(searchstring)
  LocKWords = Trim$(LocKWords)

  If Len(searchstring) = 0 Or Len(LocKWords) = 0 Then Exit Function


  LocKWArray = Split(LocKWords, ",")
  For j = LBound(LocKWArray) To UBound(LocKWArray)
    LocKWArray(j) = Trim$(LocKWArray(j))
    If Len(LocKWArray(j)) > 0 Then
      If 0 = StrComp(searchstring, Trim$(LocKWArray(j)), vbTextCompare) Then
        GetLocKW = LocKWArray(j)
        Exit Function
      End If
    End If
  Next





End Function

Public Function IsInLocKW(ByVal searchstring As String, ByVal LocKWords As String) As Boolean
  Dim j                  As Long
  Dim LocKWArray()       As String

  searchstring = Trim$(searchstring)
  LocKWords = Trim$(LocKWords)
  IsInLocKW = False

  If Len(searchstring) = 0 Or Len(LocKWords) = 0 Then Exit Function


  LocKWArray = Split(LocKWords, ",")
  For j = LBound(LocKWArray) To UBound(LocKWArray)
    LocKWArray(j) = Trim$(LocKWArray(j))
    If Len(LocKWArray(j)) > 0 Then
      If 0 = StrComp(searchstring, Trim$(LocKWArray(j)), vbTextCompare) Then
        IsInLocKW = True
        Exit Function
      End If
    End If
  Next





End Function



Public Property Get TemperatureAlarm() As Integer
  TemperatureAlarm = mTempAlarm

End Property
Public Property Get TemperatureAlarm_a() As Integer
  TemperatureAlarm_a = mTempAlarm_A

End Property

Public Property Get Temperature() As Single

  Temperature = mTemperature

End Property

Public Property Let Temperature(ByVal Value As Single)

  mTemperature = Value
  
  'EnableTemperature = 1 'use rise for testing
  
  Debug.Print "Reached Setpoint " & ReachedSet
  
  Select Case EnableTemperature
  
  
  Case TEMPERATURE_RISE
    If ReachedSet = 0 Then        ' then we're waitng for temp to pass the low set
      mTempAlarm = 0
      If (Value <= LowSet) Then
        ReachedSet = 1
        
      Else
        'nada                     ' keep waiting
      End If
    Else                          ' we're live
      If (Value >= HiSet) Then
        mTempAlarm = 1            ' alarm condition
        Debug.Print "Alarm condition = True"
      Else
        Debug.Print "Alarm condition = False"
        mTempAlarm = 0            ' alarm condition
      End If
    End If
  
  
  
  Case TEMPERATURE_FALL
    If ReachedSet = 0 Then        ' then we're waitng for temp to pass the low set
      mTempAlarm = 0
      If (Value >= HiSet) Then
        ReachedSet = 1
      Else
        'nada                     ' just bail
      End If
    Else                          ' we're live
      If (Value <= LowSet) Then
        mTempAlarm = 1            ' alarm condition
      Else
        mTempAlarm = 0            ' alarm condition
      End If
    End If
    
  Case Else  ' TEMPERATURE_NONE do nothing
    ' nada
  End Select

End Property

Public Property Get Temperature_A() As Single

  Temperature_A = mTemperature_A

End Property

Public Property Let Temperature_A(ByVal Value As Single)

  mTemperature_A = Value
  
  
  
  Select Case EnableTemperature_A
  
  Case TEMPERATURE_RISE
    If ReachedSet_A = 0 Then      ' then we're waitng for temp to pass the low set
      mTempAlarm_A = 0
      If (Value <= LowSet_A) Then
        ReachedSet_A = 1
      Else
        'nada                     ' keep waiting
      End If
    Else                          ' we're live
      If (Value >= HiSet_A) Then
        mTempAlarm_A = 1          ' alarm condition
      Else
        mTempAlarm_A = 0          ' alarm condition
      End If
    End If
  
  
  
  Case TEMPERATURE_FALL
    If ReachedSet_A = 0 Then      ' then we're waitng for temp to pass the low set
      mTempAlarm_A = 0
      If (Value >= HiSet_A) Then
        ReachedSet_A = 1
      Else
        'nada                     ' just bail
      End If
    Else                          ' we're live
      If (Value <= LowSet_A) Then
        mTempAlarm_A = 1          ' alarm condition
      Else
        mTempAlarm_A = 0          ' alarm condition
      End If
    End If
    
  Case Else  ' TEMPERATURE_NONE do nothing
    ' nada
  End Select


End Property

Public Function GetZoneInfo() As Long

  If 0 Then ' never runs
'    If CBool(USE6080) And CBool(MASTER) Then
'      ' get single zone
'
'      Set Zone = ZoneInfoList.ZoneInfoByHexSerial(Me.Serial)
'      If Zone Is Nothing Then
'        ZoneID = 0
'      Else
'        ZoneID = Zone.ID
'      End If
'    End If
  End If
End Function


Public Function ToXML() As String
  Dim XML As String ' xml for base device
  XML = XML & taggit("serial", Serial)
  XML = XML & taggit("alarmmask", CStr(AlarmMask))
  XML = XML & taggit("alarmmask_a", CStr(AlarmMask_A))
  XML = XML & taggit("alarmmask2", CStr(AlarmMask2))
  XML = XML & taggit("alarmmask2_a", CStr(AlarmMask2_A))
  XML = XML & taggit("alarmmask2_b", CStr(AlarmMask2_B))
  XML = XML & taggit("alert", CStr(Alert))
  XML = XML & taggit("alert_a", CStr(Alert_A))
  XML = XML & taggit("alert_b", CStr(Alert_B))
  ' termperature devices
  XML = XML & taggit("lowset", CStr(LowSet))
  XML = XML & taggit("lowset_a", CStr(LowSet_A))
  XML = XML & taggit("hiset", CStr(HiSet))
  XML = XML & taggit("hiset_a", CStr(HiSet_A))
  XML = XML & taggit("enabletemperature", CStr(EnableTemperature))
  XML = XML & taggit("enabletemperature_a", CStr(EnableTemperature_A))
  XML = XML & taggit("temperature", CStr(Temperature))
  XML = XML & taggit("temperature_a", CStr(Temperature_A))


  XML = XML & taggit("announce", XMLEncode(Announce))
  XML = XML & taggit("announce_a", XMLEncode(Announce_A))
  XML = XML & taggit("announce_b", XMLEncode(Announce_B))
  XML = XML & taggit("assurbit", CStr(AssurBit))
  XML = XML & taggit("assurbit_a", CStr(AssurBit_A))
  XML = XML & taggit("assurinput", CStr(AssurInput))
  XML = XML & taggit("assursecure", CStr(AssurSecure))
  XML = XML & taggit("building", XMLEncode(Building))
  XML = XML & taggit("clearbyreset", CStr(ClearByReset))

  XML = XML & taggit("description", XMLEncode(Description))

  XML = XML & taggit("deviceid", CStr(DeviceID))

  XML = XML & taggit("disableend", CStr(DisableEnd))
  XML = XML & taggit("disableend_a", CStr(DisableEnd_A))
  XML = XML & taggit("disableend2", CStr(DisableEnd2))
  XML = XML & taggit("disableend2_a", CStr(DisableEnd2_A))
  XML = XML & taggit("disablestart", CStr(DisableStart))
  XML = XML & taggit("disablestart_a", CStr(DisableStart_A))
  XML = XML & taggit("disablestart2", CStr(DisableStart2))
  XML = XML & taggit("disablestart2_a", CStr(DisableStart2_A))

  XML = XML & taggit("isaway", CStr(IsAway))
  XML = XML & taggit("isaway_a", CStr(IsAway_A))
  XML = XML & taggit("isaway_b", CStr(IsAway_B))
  XML = XML & taggit("islatching", CStr(IsLatching))
  XML = XML & taggit("islocator", CStr(IsLocator))
  XML = XML & taggit("isportable", CStr(IsPortable))
  XML = XML & taggit("midpti", CStr(MIDPTI))
  XML = XML & taggit("clspti", CStr(CLSPTI))
  XML = XML & taggit("model", XMLEncode(Model))
  XML = XML & taggit("custom", XMLEncode(Custom))

  ' day shift
  ' button 1 groups
  XML = XML & taggit("og1", CStr(OG1))
  XML = XML & taggit("og2", CStr(OG2))
  XML = XML & taggit("og3", CStr(OG3))
  XML = XML & taggit("og4", CStr(OG4))
  XML = XML & taggit("og5", CStr(OG5))
  XML = XML & taggit("og6", CStr(OG6))

  ' button 1 delay
  XML = XML & taggit("og1d", CStr(OG1D))
  XML = XML & taggit("og2d", CStr(OG2D))
  XML = XML & taggit("og3d", CStr(OG3D))
  XML = XML & taggit("og4d", CStr(OG4D))
  XML = XML & taggit("og5d", CStr(OG5D))
  XML = XML & taggit("og6d", CStr(OG6D))

  ' button 2 groups
  XML = XML & taggit("og1_a", CStr(OG1_A))
  XML = XML & taggit("og2_a", CStr(OG2_A))
  XML = XML & taggit("og3_a", CStr(OG3_A))
  XML = XML & taggit("og4_a", CStr(OG4_A))
  XML = XML & taggit("og5_a", CStr(OG5_A))
  XML = XML & taggit("og6_a", CStr(OG6_A))

  ' button 2 delay
  XML = XML & taggit("og1_ad", CStr(OG1_AD))
  XML = XML & taggit("og2_ad", CStr(OG2_AD))
  XML = XML & taggit("og3_ad", CStr(OG3_AD))
  XML = XML & taggit("og4_ad", CStr(OG4_AD))
  XML = XML & taggit("og5_ad", CStr(OG5_AD))
  XML = XML & taggit("og6_ad", CStr(OG6_AD))
    
  ' button 3 groups
    
  XML = XML & taggit("og1_b", CStr(OG1_B))
  XML = XML & taggit("og2_b", CStr(OG2_B))
  XML = XML & taggit("og3_b", CStr(OG3_B))
  XML = XML & taggit("og4_b", CStr(OG4_B))
  XML = XML & taggit("og5_b", CStr(OG5_B))
  XML = XML & taggit("og6_b", CStr(OG6_B))

  ' button 3 delay
  XML = XML & taggit("og1_bd", CStr(OG1_BD))
  XML = XML & taggit("og2_bd", CStr(OG2_BD))
  XML = XML & taggit("og3_bd", CStr(OG3_BD))
  XML = XML & taggit("og4_bd", CStr(OG4_BD))
  XML = XML & taggit("og5_bd", CStr(OG5_BD))
  XML = XML & taggit("og6_bd", CStr(OG6_BD))
    
    
    
  ' night shift
  
  XML = XML & taggit("ng1", CStr(NG1))
  XML = XML & taggit("ng2", CStr(NG2))
  XML = XML & taggit("ng3", CStr(NG3))
  XML = XML & taggit("ng4", CStr(NG4))
  XML = XML & taggit("ng5", CStr(NG5))
  XML = XML & taggit("ng6", CStr(NG6))

  XML = XML & taggit("ng1d", CStr(NG1D))
  XML = XML & taggit("ng2d", CStr(NG2D))
  XML = XML & taggit("ng3d", CStr(NG3D))
  XML = XML & taggit("ng4d", CStr(NG4D))
  XML = XML & taggit("ng5d", CStr(NG5D))
  XML = XML & taggit("ng6d", CStr(NG6D))


  XML = XML & taggit("ng1_a", CStr(NG1_A))
  XML = XML & taggit("ng2_a", CStr(NG2_A))
  XML = XML & taggit("ng3_a", CStr(NG3_A))
  XML = XML & taggit("ng4_a", CStr(NG4_A))
  XML = XML & taggit("ng5_a", CStr(NG5_A))
  XML = XML & taggit("ng6_a", CStr(NG6_A))

  XML = XML & taggit("ng1_ad", CStr(NG1_AD))
  XML = XML & taggit("ng2_ad", CStr(NG2_AD))
  XML = XML & taggit("ng3_ad", CStr(NG3_AD))
  XML = XML & taggit("ng4_ad", CStr(NG4_AD))
  XML = XML & taggit("ng5_ad", CStr(NG5_AD))
  XML = XML & taggit("ng6_ad", CStr(NG6_AD))


  XML = XML & taggit("ng1_b", CStr(NG1_B))
  XML = XML & taggit("ng2_b", CStr(NG2_B))
  XML = XML & taggit("ng3_b", CStr(NG3_B))
  XML = XML & taggit("ng4_b", CStr(NG4_B))
  XML = XML & taggit("ng5_b", CStr(NG5_B))
  XML = XML & taggit("ng6_b", CStr(NG6_B))

  XML = XML & taggit("ng1_bd", CStr(NG1_BD))
  XML = XML & taggit("ng2_bd", CStr(NG2_BD))
  XML = XML & taggit("ng3_bd", CStr(NG3_BD))
  XML = XML & taggit("ng4_bd", CStr(NG4_BD))
  XML = XML & taggit("ng5_bd", CStr(NG5_BD))
  XML = XML & taggit("ng6_bd", CStr(NG6_BD))



  ' third shift
  'button 1 groups
  XML = XML & taggit("gg1", CStr(GG1))
  XML = XML & taggit("gg2", CStr(GG2))
  XML = XML & taggit("gg3", CStr(GG3))
  XML = XML & taggit("gg4", CStr(GG4))
  XML = XML & taggit("gg5", CStr(GG5))
  XML = XML & taggit("gg6", CStr(GG6))

  'button 1 delay
  XML = XML & taggit("gg1d", CStr(GG1D))
  XML = XML & taggit("gg2d", CStr(GG2D))
  XML = XML & taggit("gg3d", CStr(GG3D))
  XML = XML & taggit("gg4d", CStr(GG4D))
  XML = XML & taggit("gg5d", CStr(GG5D))
  XML = XML & taggit("gg6d", CStr(GG6D))
  
  'button 2 groups
  XML = XML & taggit("gg1_a", CStr(GG1_A))
  XML = XML & taggit("gg2_a", CStr(GG2_A))
  XML = XML & taggit("gg3_a", CStr(GG3_A))
  XML = XML & taggit("gg4_a", CStr(GG4_A))
  XML = XML & taggit("gg5_a", CStr(GG5_A))
  XML = XML & taggit("gg6_a", CStr(GG6_A))

  'button 2 delay
  XML = XML & taggit("gg1_ad", CStr(GG1_AD))
  XML = XML & taggit("gg2_ad", CStr(GG2_AD))
  XML = XML & taggit("gg3_ad", CStr(GG3_AD))
  XML = XML & taggit("gg4_ad", CStr(GG4_AD))
  XML = XML & taggit("gg5_ad", CStr(GG5_AD))
  XML = XML & taggit("gg6_ad", CStr(GG6_AD))


  'button 3 groups
  XML = XML & taggit("gg1_b", CStr(GG1_B))
  XML = XML & taggit("gg2_b", CStr(GG2_B))
  XML = XML & taggit("gg3_b", CStr(GG3_B))
  XML = XML & taggit("gg4_b", CStr(GG4_B))
  XML = XML & taggit("gg5_b", CStr(GG5_B))
  XML = XML & taggit("gg6_b", CStr(GG6_B))

  'button 3 delay
  XML = XML & taggit("gg1_bd", CStr(GG1_BD))
  XML = XML & taggit("gg2_bd", CStr(GG2_BD))
  XML = XML & taggit("gg3_bd", CStr(GG3_BD))
  XML = XML & taggit("gg4_bd", CStr(GG4_BD))
  XML = XML & taggit("gg5_bd", CStr(GG5_BD))
  XML = XML & taggit("gg6_bd", CStr(GG6_BD))

  XML = XML & taggit("usetamperasinput", CStr(UseTamperAsInput))

  XML = XML & taggit("notamper", CStr(NoTamper))
  
  XML = XML & taggit("ignoretamper", CStr(IgnoreTamper))
  
  XML = XML & taggit("owner", XMLEncode(Owner))
  XML = XML & taggit("owner_a", "")
  XML = XML & taggit("pause", CStr(Pause))
  XML = XML & taggit("pause_a", CStr(Pause_A))
  XML = XML & taggit("pause_b", CStr(Pause_B))
  XML = XML & taggit("repeats", CStr(Repeats))
  XML = XML & taggit("repeats_a", CStr(Repeats_A))
  XML = XML & taggit("repeats_b", CStr(Repeats_B))
  XML = XML & taggit("repeatuntil", CStr(RepeatUntil))
  XML = XML & taggit("repeatuntil_a", CStr(RepeatUntil_A))
  XML = XML & taggit("repeatuntil_b", CStr(RepeatUntil_B))
  XML = XML & taggit("residentid", CStr(ResidentID))
  XML = XML & taggit("residentid_a", "")
  XML = XML & taggit("room", XMLEncode(Room))
  XML = XML & taggit("room_a", "")
  XML = XML & taggit("roomid", CStr(RoomID))
  XML = XML & taggit("roomid_a", "")
  XML = XML & taggit("sendcancel", CStr(SendCancel))
  XML = XML & taggit("sendcancel_a", CStr(SendCancel_A))
  XML = XML & taggit("sendcancel_b", CStr(SendCancel_B))
  XML = XML & taggit("superviseperiod", CStr(SupervisePeriod))
  XML = XML & taggit("useassur", CStr(UseAssur))
  XML = XML & taggit("useassur_a", CStr(UseAssur_A))
  XML = XML & taggit("useassur_b", CStr(UseAssur_B))
  XML = XML & taggit("useassur2", CStr(UseAssur2))
  XML = XML & taggit("useassur2_a", CStr(UseAssur2_A))
  XML = XML & taggit("useassur2_b", CStr(UseAssur2_B))
                  
  XML = XML & taggit("tamperasinput", CStr(UseTamperAsInput))
  
  XML = XML & taggit("ignored", CStr(Ignored))
  
  XML = XML & taggit("configurationstring", XMLEncode(Configurationstring))

  ToXML = XML '.ToString

End Function
Public Function SerialToXML() As String
  ' xml for serial extensions
  Dim XML As String
  XML = taggit("deviceid", CStr(DeviceID))
  XML = XML & taggit("serial", Serial)
  XML = XML & taggit("serialtapprotocol", CStr(SerialTapProtocol))
  XML = XML & taggit("serialskip", CStr(SerialSkip))
  XML = XML & taggit("serialmessagelen", CStr(SerialMessageLen))
  XML = XML & taggit("serialautoclear", CStr(SerialAutoClear))
  XML = XML & taggit("serialinclude", XMLEncode(SerialInclude))
  XML = XML & taggit("serialexclude", XMLEncode(SerialExclude))
  XML = XML & taggit("serialport", CStr(SerialPort))
  XML = XML & taggit("serialbaud", CStr(SerialBaud))
  XML = XML & taggit("serialparity", XMLEncode(SerialParity))
  XML = XML & taggit("serialbits", CStr(Serialbits))
  XML = XML & taggit("serialflow", CStr(SerialFlow))
  XML = XML & taggit("serialstopbits", SerialStopbits)
  XML = XML & taggit("serialsettings", SerialSettings)
  XML = XML & taggit("serialeolchar", CStr(SerialEOLChar))
  XML = XML & taggit("serialpreamble", XMLEncode(SerialPreamble))

  XML = XML & taggit("lowset", XMLEncode(LowSet))
  XML = XML & taggit("lowset_a", XMLEncode(LowSet_A))
  XML = XML & taggit("hiset", XMLEncode(HiSet))
  XML = XML & taggit("hiset_a", XMLEncode(HiSet_A))
  
  XML = XML & taggit("enabletemperature", XMLEncode(EnableTemperature))
  XML = XML & taggit("enabletemperature_a", XMLEncode(EnableTemperature_A))


  SerialToXML = XML '.ToString

End Function

'Public Property Get HasSettled(ByVal EventTime As Date, Window As Long) As Boolean
'   HasSettled = DateDiff("s", LastAlarm, EventTime) > Window
'End Property

Public Property Get IsPCA() As Integer       ' 0 or 1
  IsPCA = IIf((Model = PCA_DEV_NAME), 1, 0)
End Property


Public Property Get IsSerialDevice() As Integer       ' 0 or 1
  IsSerialDevice = IIf((Model = COM_DEV_NAME), 1, 0)
End Property



Public Property Get IsLate() As Boolean
  
  
  If (IsAway And IsPortable) Then
    Trace Serial & " is Away and Portable"
    'Debug.Print Serial & " is Away and Portable"
  Else
    'Debug.Print mSupervisePeriod
    'Debug.Print "cEsdevice.Islate LastSupevise, Datediff (min) " & Serial & " " & LastSupervise & ", " & DateDiff("n", LastSupervise, Now)
    If Ignored Then
      IsLate = False
    ElseIf NoCheckin Then
      IsLate = False
    Else
      If USE6080 Then
        IsLate = IsMissing
        If IsMissing Then
          IsLate = IsMissing
        End If
      Else
        If DateDiff("n", LastSupervise, Now) > mSupervisePeriod Then
          'Debug.Assert 0
          Debug.Print "DateDiff " & DateDiff("n", LastSupervise, Now), mSupervisePeriod
        Else
          Debug.Print "DateDiff " & DateDiff("n", LastSupervise, Now), mSupervisePeriod
        End If
        IsLate = DateDiff("n", LastSupervise, Now) > mSupervisePeriod
      End If
      
    End If
    'OutputDebugString "Supervise " & LastSupervise & " " & DateDiff("n", LastSupervise, Now) & " " & mSupervisePeriod & vbCrLf
  End If
End Property

Public Property Let SupervisePeriod(ByVal Value As Long)
  mSupervisePeriod = Value
End Property
Public Property Get SupervisePeriod() As Long
  SupervisePeriod = mSupervisePeriod
End Property


Public Function LastFirst() As String
  If Len(NameLast) > 0 Then
    If Len(NameFirst) > 0 Then
      LastFirst = NameLast & ", " & NameFirst
    Else
      LastFirst = NameLast
    End If
  Else
    LastFirst = NameFirst
  End If
End Function
'Public Function LastFirst_A() As String
'
'  If Len(NameLast_A) > 0 Then
'    If Len(NameFirst_A) > 0 Then
'      LastFirst_A = NameLast_A & ", " & NameFirst_A
'    Else
'      LastFirst_A = NameLast_A
'    End If
'  Else
'    LastFirst_A = NameFirst_A
'  End If
'
'End Function

Public Function FetchRoom() As String
  Dim SQL As String
  Dim rs As Recordset
  Room = ""
  'Room_A = ""

  Building = ""
  If RoomID <> 0 Then

    SQL = "SELECT Room, away , lockw , flags FROM Rooms where ROOMID = " & RoomID
    Set rs = ConnExecute(SQL)
    If Not rs.EOF Then
      Room = rs("Room") & ""
      Building = ""  'rs("Building") & ""
      IsAway = IsAway Or IIf(rs("away") = 1, 1, 0)
      locKW = rs("lockw") & ""
      RoomFlags = Val(rs("flags") & "")
    End If
    rs.Close

    Set rs = Nothing
  End If

End Function
Public Function FetchResident() As String

  Dim SQL As String
  Dim rs As Recordset

  NameLast = ""
  NameFirst = ""
  'NameLast_A = ""
  'NameFirst_A = ""

  If ResidentID <> 0 Then

    SQL = "SELECT namelast,namefirst,away FROM residents WHERE ResidentID = " & ResidentID
    Set rs = ConnExecute(SQL)
    If Not rs.EOF Then
      NameLast = rs("namelast") & ""
      NameFirst = rs("namefirst") & ""
      IsAway = IsAway Or IIf(rs("away") = 1, 1, 0)
      
    End If
    rs.Close
    Set rs = Nothing
  End If

End Function

Public Property Get isDisabled() As Boolean


'  Dim SecStart As Long
'  Dim SecEnd As Long
'  Dim SecondsSinceMidnight As Long


  Dim DayStart As Date
  Dim DayEnd As Date
  Dim CurrentTime As Date


  If DisableStart = DisableEnd Then
    isDisabled = False
  Else

    CurrentTime = Now

    DayStart = Format(Now, "mm/dd/yyyy")
    DayEnd = DayStart

    DayStart = DateAdd("h", DisableStart, DayStart)
    DayEnd = DateAdd("h", DisableEnd, DayEnd)
    If (DayEnd < DayStart) Then
      DayEnd = DateAdd("d", 1, DayEnd)
    End If
    If (CurrentTime >= DayStart) And (CurrentTime < DayEnd) Then
      isDisabled = True
    End If
  End If



'  SecStart = DisableStart * SecondsPerHour
'  SecEnd = DisableEnd * SecondsPerHour
'
'  SecondsSinceMidnight = Timer()  ' seconds since midnight
'
'  If DisableStart = DisableEnd Then
'    isDisabled = False
'  ElseIf DisableStart < DisableEnd Then  ' normal 8:00 - 17:00 ' between
'    '  isDisabled = (SecondsSinceMidnight >= DisableStart * 3600) And (SecondsSinceMidnight < DisableEnd * 3600)
'    isDisabled = (SecondsSinceMidnight >= SecStart) And (SecondsSinceMidnight < SecEnd)
'  Else  ' flipped 17:00 - 8:00 ' after start OR before end
'    isDisabled = (SecondsSinceMidnight >= SecStart) Or (SecondsSinceMidnight < SecEnd)
'  End If

End Property

Public Property Get isDisabled_A() As Boolean


'  Dim SecStart As Long
'  Dim SecEnd As Long
'  Dim SecondsSinceMidnight As Long

  Dim DayStart As Date
  Dim DayEnd As Date
  Dim CurrentTime As Date


  If DisableStart_A = DisableEnd_A Then
    isDisabled_A = False
  Else

    CurrentTime = Now

    DayStart = Format(Now, "mm/dd/yyyy")
    DayEnd = DayStart

    DayStart = DateAdd("h", DisableStart_A, DayStart)
    DayEnd = DateAdd("h", DisableEnd_A, DayEnd)
    If (DayEnd < DayStart) Then
      DayEnd = DateAdd("d", 1, DayEnd)
    End If
    If (CurrentTime >= DayStart) And (CurrentTime < DayEnd) Then
      isDisabled_A = True
    End If
  End If



  '  SecStart = DisableStart_A * SecondsPerHour
  '  SecEnd = DisableEnd_A * SecondsPerHour
  '
  '  SecondsSinceMidnight = Timer()  ' seconds since midnight
  '
  '  If DisableStart_A = DisableEnd_A Then
  '    isDisabled_A = False
  '  ElseIf DisableStart_A < DisableEnd_A Then  ' normal 8:00 - 17:00 ' between
  '    isDisabled_A = (SecondsSinceMidnight >= SecStart) And (SecondsSinceMidnight < SecEnd)
  '  Else  ' flipped 17:00 - 8:00 ' after start OR before end
  '
  '
  '
  '    If (SecondsSinceMidnight >= SecStart) Then
  '       SecEnd = SecEnd + (12 * SecondsPerHour)
  '    End If
  '
  '    isDisabled_A = (SecondsSinceMidnight >= SecStart) Or (SecondsSinceMidnight < SecEnd)
  '  End If

End Property

Public Property Get isDisabled_B() As Boolean


'  Dim SecStart As Long
'  Dim SecEnd As Long
'  Dim SecondsSinceMidnight As Long

  Dim DayStart As Date
  Dim DayEnd As Date
  Dim CurrentTime As Date

  If DisableStart_B = DisableEnd_A Then
    isDisabled_B = False
  Else

    CurrentTime = Now

    DayStart = Format(Now, "mm/dd/yyyy")
    DayEnd = DayStart

    DayStart = DateAdd("h", DisableStart_B, DayStart)
    DayEnd = DateAdd("h", DisableEnd_B, DayEnd)
    If (DayEnd < DayStart) Then
      DayEnd = DateAdd("d", 1, DayEnd)
    End If
    If (CurrentTime >= DayStart) And (CurrentTime < DayEnd) Then
      isDisabled_B = True
    End If
  End If



  '  SecStart = DisableStart_B * SecondsPerHour
  '  SecEnd = DisableEnd_B * SecondsPerHour
  '
  '  SecondsSinceMidnight = Timer()  ' seconds since midnight
  '
  '  If DisableStart_B = DisableEnd_B Then
  '    isDisabled_B = False
  '  ElseIf DisableStart_B < DisableEnd_B Then  ' normal 8:00 - 17:00 ' between
  '    isDisabled_B = (SecondsSinceMidnight >= SecStart) And (SecondsSinceMidnight < SecEnd)
  '  Else  ' flipped 17:00 - 8:00 ' after start OR before end
  '    isDisabled_B = (SecondsSinceMidnight >= SecStart) Or (SecondsSinceMidnight < SecEnd)
  '  End If

End Property


Public Function Parse(rs As Recordset) As Long
  Dim j As Long
  

  On Error GoTo Parse_Error

  Clear
  IDM = Val(rs("IDM") & "") ' 6080 zone id as applied also sets ZoneID
  
  IDL = Val(rs("IDL") & "") ' Mode: Mob'bile , Fix'd , SP (soft pointer)
  
  DeviceID = rs("DeviceID")
  NumInputs = Max(1, Val("" & rs("NumInputs")))
  IsPortable = IIf(rs("IsPortable") = 1, 1, 0)
  ClearByReset = IIf(rs("ClearByReset") = 1, 1, 0)
  AssurInput = Min(NumInputs, Val("" & rs("assurinput")))  ' limit to inputs supported

  Configurationstring = (rs("Configurationstring") & "")

  ' first shift
  
  
  
  OG1 = Val(rs("OG1") & "")
  OG2 = Val(rs("OG2") & "")
  OG3 = Val(rs("OG3") & "")
  OG4 = Val(rs("OG4") & "")
  OG5 = Val(rs("OG5") & "")
  OG6 = Val(rs("OG6") & "")

  OG1D = Val(rs("OG1d") & "")
  OG2D = Val(rs("OG2d") & "")
  OG3D = Val(rs("OG3d") & "")
  OG4D = Val(rs("OG4d") & "")
  OG5D = Val(rs("OG5d") & "")
  OG6D = Val(rs("OG6d") & "")

  OG1_A = Val(rs("OG1_A") & "")
  OG2_A = Val(rs("OG2_A") & "")
  OG3_A = Val(rs("OG3_A") & "")
  OG4_A = Val(rs("OG4_A") & "")
  OG5_A = Val(rs("OG5_A") & "")
  OG6_A = Val(rs("OG6_A") & "")
  
  OG1_AD = Val(rs("OG1_Ad") & "")
  OG2_AD = Val(rs("OG2_Ad") & "")
  OG3_AD = Val(rs("OG3_Ad") & "")
  OG4_AD = Val(rs("OG4_Ad") & "")
  OG5_AD = Val(rs("OG5_Ad") & "")
  OG6_AD = Val(rs("OG6_Ad") & "")
  
  OG1_B = Val(rs("OG1_B") & "")
  OG2_B = Val(rs("OG2_B") & "")
  OG3_B = Val(rs("OG3_B") & "")
  OG4_B = Val(rs("OG4_B") & "")
  OG5_B = Val(rs("OG5_B") & "")
  OG6_B = Val(rs("OG6_B") & "")
  
  OG1_BD = Val(rs("OG1_Bd") & "")
  OG2_BD = Val(rs("OG2_Bd") & "")
  OG3_BD = Val(rs("OG3_Bd") & "")
  OG4_BD = Val(rs("OG4_Bd") & "")
  OG5_BD = Val(rs("OG5_Bd") & "")
  OG6_BD = Val(rs("OG6_Bd") & "")
  
  
  
  ' second (night shift)
  
  NG1 = Val(rs("NG1") & "")
  NG2 = Val(rs("NG2") & "")
  NG3 = Val(rs("NG3") & "")
  NG4 = Val(rs("NG4") & "")
  NG5 = Val(rs("NG5") & "")
  NG6 = Val(rs("NG6") & "")

  NG1D = Val(rs("NG1d") & "")
  NG2D = Val(rs("NG2d") & "")
  NG3D = Val(rs("NG3d") & "")
  NG4D = Val(rs("NG4d") & "")
  NG5D = Val(rs("NG5d") & "")
  NG6D = Val(rs("NG6d") & "")

  NG1_A = Val(rs("NG1_A") & "")
  NG2_A = Val(rs("NG2_A") & "")
  NG3_A = Val(rs("NG3_A") & "")
  NG4_A = Val(rs("NG4_A") & "")
  NG5_A = Val(rs("NG5_A") & "")
  NG6_A = Val(rs("NG6_A") & "")

  NG1_AD = Val(rs("NG1_Ad") & "")
  NG2_AD = Val(rs("NG2_Ad") & "")
  NG3_AD = Val(rs("NG3_Ad") & "")
  NG4_AD = Val(rs("NG4_Ad") & "")
  NG5_AD = Val(rs("NG5_Ad") & "")
  NG6_AD = Val(rs("NG6_Ad") & "")
  
  
  NG1_B = Val(rs("NG1_B") & "")
  NG2_B = Val(rs("NG2_B") & "")
  NG3_B = Val(rs("NG3_B") & "")
  NG4_B = Val(rs("NG4_B") & "")
  NG5_B = Val(rs("NG5_B") & "")
  NG6_B = Val(rs("NG6_B") & "")

  NG1_BD = Val(rs("NG1_Bd") & "")
  NG2_BD = Val(rs("NG2_Bd") & "")
  NG3_BD = Val(rs("NG3_Bd") & "")
  NG4_BD = Val(rs("NG4_Bd") & "")
  NG5_BD = Val(rs("NG5_Bd") & "")
  NG6_BD = Val(rs("NG6_Bd") & "")
  
  
  
' new for third shift

  GG1 = Val(rs("GG1") & "") ' INPUT 1 group
  GG2 = Val(rs("GG2") & "")
  GG3 = Val(rs("GG3") & "")
  GG4 = Val(rs("GG4") & "")
  GG5 = Val(rs("GG5") & "")
  GG6 = Val(rs("GG6") & "")

  GG1D = Val(rs("GG1d") & "") ' escalation time
  GG2D = Val(rs("GG2d") & "")
  GG3D = Val(rs("GG3d") & "")
  GG4D = Val(rs("GG4d") & "")
  GG5D = Val(rs("GG5d") & "")
  GG6D = Val(rs("GG6d") & "")

  GG1_A = Val(rs("GG1_A") & "") ' INPUT 2
  GG2_A = Val(rs("GG2_A") & "")
  GG3_A = Val(rs("GG3_A") & "")
  GG4_A = Val(rs("GG4_A") & "")
  GG5_A = Val(rs("GG5_A") & "")
  GG6_A = Val(rs("GG6_A") & "")


  GG1_AD = Val(rs("GG1_AD") & "")
  GG2_AD = Val(rs("GG2_AD") & "")
  GG3_AD = Val(rs("GG3_AD") & "")
  GG4_AD = Val(rs("GG4_AD") & "")
  GG5_AD = Val(rs("GG5_AD") & "")
  GG6_AD = Val(rs("GG6_AD") & "")


  GG1_B = Val(rs("GG1_B") & "") ' INPUT 3
  GG2_B = Val(rs("GG2_B") & "")
  GG3_B = Val(rs("GG3_B") & "")
  GG4_B = Val(rs("GG4_B") & "")
  GG5_B = Val(rs("GG5_B") & "")
  GG6_B = Val(rs("GG6_B") & "")


  GG1_BD = Val(rs("GG1_BD") & "")
  GG2_BD = Val(rs("GG2_BD") & "")
  GG3_BD = Val(rs("GG3_BD") & "")
  GG4_BD = Val(rs("GG4_BD") & "")
  GG5_BD = Val(rs("GG5_BD") & "")
  GG6_BD = Val(rs("GG6_BD") & "")



  DisableStart = Val(rs("disablestart") & "")
  DisableEnd = Val(rs("disableEnd") & "")
  Repeats = Val(rs("repeats") & "")
  Pause = Val(rs("pause") & "")
  RepeatUntil = Val(rs("RepeatUntil") & "")
  SendCancel = IIf(rs("sendcancel") = 1, 1, 0)
  Serial = rs("serial") & ""
  Model = rs("model") & ""
  UseAssur = IIf(rs("UseAssur") = 1, 1, 0)
  UseAssur2 = IIf(rs("UseAssur2") = 1, 1, 0)
  AssurSecure = IIf(rs("VacationSuper") = 1, 1, 0)
  AlarmMask = Min(2, Val(rs("AlarmMask") & ""))
  Announce = rs("Announce") & ""

  IgnoreTamper = IIf(rs("ignoretamper") = 1, 1, 0)
  
  ResidentID = IIf(IsNull(rs("ResidentID")), 0, rs("ResidentID"))  ' assigned to resident, 0 if not assigned to resident
  RoomID = IIf(IsNull(rs("RoomID")), 0, rs("RoomID"))  ' assigned to room/location rather than resident else 0

  DisableStart_A = Val(rs("disablestart_A") & "")
  DisableEnd_A = Val(rs("disableEnd_A") & "")
  Repeats_A = Val(rs("repeats_A") & "")
  Pause_A = Val(rs("pause_A") & "")
  RepeatUntil_A = Val(rs("RepeatUntil_A") & "")
  SendCancel_A = IIf(rs("sendcancel_A") = 1, 1, 0)
  UseAssur_A = IIf(rs("UseAssur_A") = 1, 1, 0)
  UseAssur2_A = IIf(rs("UseAssur2_A") = 1, 1, 0)
  AssurSecure_A = IIf(rs("VacationSuper_A") = 1, 1, 0)
  AlarmMask_A = IIf(rs("AlarmMask_A") = 1, 1, 0)
  Announce_A = rs("Announce_A") & ""


  DisableStart_B = Val(rs("disablestart_B") & "")
  DisableEnd_B = Val(rs("disableEnd_B") & "")
  Repeats_B = Val(rs("repeats_B") & "")
  Pause_B = Val(rs("pause_B") & "")
  RepeatUntil_B = Val(rs("RepeatUntil_B") & "")
  SendCancel_B = IIf(rs("sendcancel_B") = 1, 1, 0)
  UseAssur_B = IIf(rs("UseAssur_B") = 1, 1, 0)
  UseAssur2_B = IIf(rs("UseAssur2_B") = 1, 1, 0)
  AssurSecure_B = IIf(rs("VacationSuper_B") = 1, 1, 0)
  AlarmMask_B = IIf(rs("AlarmMask_B") = 1, 1, 0)
  Announce_B = rs("Announce_B") & ""


  UseTamperAsInput = rs("UseTamperAsInput") & ""

  ' temperature devices
  
  EnableTemperature = Max(0, Min(2, Val(rs("EnableTemp") & "")))
  EnableTemperature_A = Max(0, Min(2, Val(rs("EnableTemp_a") & "")))
  LowSet = Val(rs("lowset") & "")
  LowSet_A = Val(rs("lowset_a") & "")
  HiSet = Val(rs("hiset") & "")
  HiSet_A = Val(rs("hiset_a") & "")

  SerialTapProtocol = Val(rs("SerialTapProtocol") & "")
  SerialSkip = rs("SerialSkip")
  SerialMessageLen = rs("SerialMessageLen")
  SerialAutoClear = rs("SerialAutoClear")
  SerialPort = rs("SerialPort")
  SerialBaud = rs("SerialBaud")
  Serialbits = rs("SerialBits")
  SerialParity = rs("SerialParity") & ""
  SerialStopbits = rs("SerialStopBits") & ""
  SerialInclude = rs("SerialInclude") & ""
  SerialExclude = rs("SerialExclude") & ""
  SerialFlow = rs("SerialFlow")
  SerialEOLChar = rs("SerialEOLChar")
  If SerialBaud <> 0 Then
    SerialSettings = SerialBaud & SerialParity & Serialbits & SerialStopbits
  End If
  SerialPreamble = rs("SerialPreamble") & ""

  Ignored = Val(rs("ignored") & "")

  Custom = rs("custom") & ""

  'locKW = rs("lockw") & ""
  
  For j = 0 To MAX_ESDEVICETYPES
    If (0 = StrComp(ESDeviceType(j).Model, Model, vbTextCompare)) Then
      MIDPTI = ESDeviceType(j).MIDPTI
      CLSPTI = ESDeviceType(j).CLSPTI
      
      
      SupervisePeriod = ESDeviceType(j).Checkin
      If SupervisePeriod = 0 Then
        SupervisePeriod = 100
      End If
      IsLatching = ESDeviceType(j).Latching
      Description = ESDeviceType(j).desc
      IsPortable = ESDeviceType(j).Portable
      
      'AutoClear = ESDeviceType(j).AutoClear
      NoCheckin = ESDeviceType(j).NoCheckin
      AutoClear = ESDeviceType(j).AutoClear
      'AutoClear = GetAutoClear(Model)
      Exit For
    End If
  Next
  
  ' these are by device type, not per device
  'MIDPTI = GetMIDPTI(Model)
  'CLSPTI = GetCLSPTI(Model)

  

  'SupervisePeriod = GetSupervisePeriod(Model)
  'AutoClear = GetAutoClear(Model)
  'NoCheckin = GetNoCheckIn(Model)
'  Me.FetchResident
'  Me.FetchRoom
  Parse = True

Parse_Resume:
  

  On Error GoTo 0
  Exit Function

Parse_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cESDevice.Parse." & Erl
  Parse = False
  Resume Parse_Resume

End Function

Sub Refresh()
  Dim rs As Recordset
  Dim SQL As String

  If Len(Serial) > 0 Then
    SQL = "SELECT * FROM Devices WHERE Serial = " & q(Serial)
    Set rs = ConnExecute(SQL)
    If Not rs.EOF Then
      
      Me.Parse rs
      Me.FetchResident
      Me.FetchRoom
    End If
    rs.Close
    Set rs = Nothing
  End If
End Sub



Sub Clear()
  ResidentID = 0
  'ResidentID_A = 0
  RoomID = 0
  'RoomID_A = 0
  NameLast = ""
  NameFirst = ""
  'NameLast_A = ""
  'NameFirst_A = ""

  STAT = 0
  LastAlarm = 0
  LastRestore = 0
  LastAlarm_A = 0
  LastAlarm_B = 0
  LastRestore_A = 0
  LastRestore_B = 0

  ReachedSet = 0
  ReachedSet_A = 0
  ReachedSet_B = 0

  NoCheckin = 0

  LastSupervise = Now
  LastTest = 0

  SerialTapProtocol = 0
  SerialSkip = 0
  SerialMessageLen = 0
  SerialAutoClear = 0
  SerialPort = 0
  SerialBaud = 9600
  Serialbits = 8
  SerialParity = "N"
  SerialStopbits = "1"
  SerialInclude = ""
  SerialExclude = ""
  SerialFlow = 0
  SerialEOLChar = 0
  SerialSettings = ""
  SerialPreamble = ""

  Ignored = 0

  locKW = ""

End Sub
' add pending message
' Outbounds.AddMessage Device.serial, MSGTYPE_REPEATERNID, "", 0



Private Sub Class_Initialize()
  'Set RecentHopsRSS = New Collection
  
  NID = -1
  Layer = -1
  SupervisePeriod = 20   ' default
  AssurInput = 1        ' default
  LastSupervise = Now
  Set Zone = New cZoneInfo
End Sub

Public Property Get Assur() As Integer
  Assur = IIf(UseAssur = 1 Or UseAssur2 = 1, 1, 0)
End Property

Public Property Get Assur_A() As Integer
  Assur_A = IIf(UseAssur_A = 1 Or UseAssur2_A = 1, 1, 0)
End Property

Public Property Get LastRestore() As Date
  LastRestore = mLastRestore
End Property

Public Property Let LastRestore(ByVal LastRestore As Date)
  mLastRestore = LastRestore
End Property

Public Property Get LastRestore_A() As Date
  LastRestore_A = MLastRestore_A
End Property

Public Property Let LastRestore_A(ByVal LastRestore_A As Date)
  MLastRestore_A = LastRestore_A
End Property


Public Property Get JamCount() As Long

  JamCount = mJamCount

End Property

Public Sub IncrementJam(ByVal Value As Byte)
  
  If Value = 0 Then
    If mJamCount > 0 Then
      mJamCount = mJamCount - 1
    End If
  Else
    If mJamCount < 10000 Then
      mJamCount = mJamCount + 1
    End If
  End If

End Sub

Public Property Get Battery() As Integer
  Battery = mBattery
End Property
Public Property Get BatteryDelayTimeout() As Boolean
  Dim TimedOut As Boolean
  Dim Diff As Double

  If LowBattPending Then
    Diff = DateDiff("n", Now, LastLowBatt)
    TimedOut = Diff <= 0
    'dbg Me.Serial & " Low Battery Delay Times Out In " & Diff & " Minutes"
    If TimedOut Then
      'dbg Me.Serial & " Low Battery Ready for Trouble "
      LowBattPending = False
      BatteryDelayTimeout = True
    End If
  End If
End Property
Public Property Let Battery(ByVal Value As Integer)
' delay displaying intermittent low - battery alarms for "swingers"
  If mBattery = 0 Then
    If (Value = 1) Then
      If (gLoBattDelay > 0) Then
        LastLowBatt = DateAdd("n", gLoBattDelay, Now)  ' should'a had a checkin in the interim
      Else
        LastLowBatt = Now
      End If
      LowBattPending = True  ' pending on transition from 0 to 1
      'dbg Me.Serial & " Low Battery Pending, Ready at " & LastLowBatt

    Else
      LowBattPending = False  ' reset if value still 0
      'dbg Me.Serial & " Battery Good "
    End If


  Else
    If Value = 0 Then
      LowBattPending = False  ' reset if value has transitioned to 0
      'dbg Me.Serial & " Battery Cleared"
    End If
  End If
  mBattery = Value

End Property

Public Property Get Tamper() As Integer
  Tamper = mTamper
End Property

Public Property Let Tamper(ByVal Value As Integer)
  mTamper = Value
  Exit Property

  '  If mTamper = 0 Then
  '    If (Value = 1) Then
  '      If (gLoBattDelay > 0) Then
  '        LastLowBatt = DateAdd("n", gLoBattDelay, Now) ' should'a had a checkin in the interim
  '        Debug.Print "Event expires on " & LastLowBatt
  '
  '      Else
  '        LastLowBatt = Now
  '      End If
  '      LowBattPending = True ' pending on transition from 0 to 1
  '    Else
  '      LowBattPending = False ' reset if value still 0
  '    End If
  '  Else
  '    If Value = 0 Then
  '      LowBattPending = False ' reset if value has transitioned to 0
  '    End If
  '  End If
  '
  '
  '

End Property

Public Property Get FirstHop() As String

  FirstHop = mFirstHop

End Property

Public Property Let FirstHop(ByVal FirstHop As String)
  mFirstHop = FirstHop
  
End Property

Public Property Get LastLevel() As Integer

  LastLevel = mLastLevel

End Property

Public Property Let LastLevel(ByVal LastLevel As Integer)
  mLastLevel = LastLevel
End Property

Public Property Get AutoClear() As Long

  AutoClear = mAutoClear

End Property

Public Property Let AutoClear(ByVal AutoClear As Long)

  mAutoClear = AutoClear

End Property


Public Property Get IsTemperatureDev() As Integer
  ' read only
  IsTemperatureDev = IsTemperatureDevice(CLSPTI)

End Property

'Public Property Let IsTemperatureDev(ByVal IsTemperatureDev As Integer)
'
'  mIsTemperatureDev = IsTemperatureDev
'
'End Property

Public Property Get DecimalSerial() As Long
  DecimalSerial = Val("&h" & Right$(Serial, 6))
End Property


Public Property Get Dead() As Integer

  Dead = mDead

End Property

Public Property Let Dead(ByVal Value As Integer)
  
  IsMissing = Value
  
  mDead = Value

End Property

''*********** IDM is alias for ZONE ID of 6080 system
Public Property Get IDM() As Long
  IDM = ZoneID
End Property

Public Property Let IDM(ByVal Value As Long)
   ZoneID = Value
End Property

Public Property Get IDL() As Long
  'IDL is the mode for the device
  
  IDL = mIDL

End Property

Public Property Let IDL(ByVal IDL As Long)

  mIDL = IDL

End Property

Public Property Get ZoneID() As Long

  ZoneID = mZoneID

End Property

Public Property Let ZoneID(ByVal ZoneID As Long)
 ' Debug.Assert ZoneID
  'Debug.Print "Zone ID " & ZoneID
  If ZoneID = 0 Then
  '  Debug.Assert 0
  End If
  
  mZoneID = ZoneID

End Property

Public Property Get LastRestore_B() As Date

  LastRestore_B = mLastRestore_B

End Property

Public Property Let LastRestore_B(ByVal LastRestore_B As Date)

  mLastRestore_B = LastRestore_B

End Property

Public Property Get Room() As String

  Room = mRoom

End Property

Public Property Let Room(ByVal Room As String)

  mRoom = Room

End Property

Public Property Get PTI() As Long

  PTI = mPTI

End Property

Public Property Let PTI(ByVal PTI As Long)

  mPTI = PTI

End Property

Public Property Get LastSupervise() As Date

  LastSupervise = mLastSupervise

End Property

Public Property Let LastSupervise(ByVal LastSupervise As Date)

  mLastSupervise = LastSupervise

End Property
