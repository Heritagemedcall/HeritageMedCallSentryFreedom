Attribute VB_Name = "modDevices"

Option Explicit

Public Function ResetRunningDevices(ByVal Model As String) As Long
  Dim Device        As cESDevice
  Dim ESDev         As ESDeviceTypeType
  Dim SQl           As String
  Dim j             As Long


  For j = 1 To MAX_ESDEVICETYPES
    ESDev = ESDeviceType(j)
    If 0 = StrComp(ESDev.Model, Model, vbTextCompare) Then
      Exit For  ' found it, we're done
    End If
  Next


  If j <= MAX_ESDEVICETYPES Then  ' we haven't gone over the deep end

    ' I think it already saved the updates

    '    ' Main tab
    SQl = "UPDATE Devices SET "
    'SQL = SQL & " Description = " & q(ESDev.Desc) ' there is no description in devices, only custom
    SQl = SQl & " Custom = " & q(ESDev.desc)
    'SQL = SQL & " Checkin = " & ESDev.Checkin
    SQl = SQl & " ,ClearByReset = " & ESDev.ClearByReset
    SQl = SQl & ", IgnoreTamper = " & ESDev.IgnoreTamper

    ' tab 1
    SQl = SQl & ", Announce = " & q(ESDev.Announce)
    SQl = SQl & ", Repeats = " & ESDev.Repeats
    SQl = SQl & ", Pause = " & ESDev.Pause
    SQl = SQl & ", repeatuntil = " & ESDev.RepeatUntil
    SQl = SQl & ", SendCancel = " & ESDev.SendCancel
    
    ' tab 2

    SQl = SQl & ", Announce_A = " & q(ESDev.Announce2)
    SQl = SQl & ", Repeats_A = " & ESDev.Repeats_A
    SQl = SQl & ", Pause_A = " & ESDev.Pause_A
    SQl = SQl & ", repeatuntil_A = " & ESDev.RepeatUntil_A
    SQl = SQl & ", SendCancel_A = " & ESDev.SendCancel_A

    SQl = SQl & ", OG1 = " & ESDev.OG1
    SQl = SQl & ", OG2 = " & ESDev.OG2
    SQl = SQl & ", OG3 = " & ESDev.OG3
    SQl = SQl & ", OG4 = " & ESDev.OG4
    SQl = SQl & ", OG5 = " & ESDev.OG5
    SQl = SQl & ", OG6 = " & ESDev.OG6

    SQl = SQl & ", OG1d = " & ESDev.OG1D
    SQl = SQl & ", OG2d = " & ESDev.OG2D
    SQl = SQl & ", OG3d = " & ESDev.OG3D
    SQl = SQl & ", OG4d = " & ESDev.OG4D
    SQl = SQl & ", OG5d = " & ESDev.OG5D
    SQl = SQl & ", OG6d = " & ESDev.OG6D

    SQl = SQl & ", OG1_A = " & ESDev.OG1_A
    SQl = SQl & ", OG2_A = " & ESDev.OG2_A
    SQl = SQl & ", OG3_A = " & ESDev.OG3_A
    SQl = SQl & ", OG4_A = " & ESDev.OG4_A
    SQl = SQl & ", OG5_A = " & ESDev.OG5_A
    SQl = SQl & ", OG6_A = " & ESDev.OG6_A

    SQl = SQl & ", OG1_AD = " & ESDev.OG1_AD
    SQl = SQl & ", OG2_AD = " & ESDev.OG2_AD
    SQl = SQl & ", OG3_AD = " & ESDev.OG3_AD
    SQl = SQl & ", OG4_AD = " & ESDev.OG4_AD
    SQl = SQl & ", OG5_AD = " & ESDev.OG5_AD
    SQl = SQl & ", OG6_AD = " & ESDev.OG6_AD


    SQl = SQl & ", NG1 = " & ESDev.NG1
    SQl = SQl & ", NG2 = " & ESDev.NG2
    SQl = SQl & ", NG3 = " & ESDev.NG3
    SQl = SQl & ", NG4 = " & ESDev.NG4
    SQl = SQl & ", NG5 = " & ESDev.NG5
    SQl = SQl & ", NG6 = " & ESDev.NG6


    SQl = SQl & ", NG1_A = " & ESDev.NG1_A
    SQl = SQl & ", NG2_A = " & ESDev.NG2_A
    SQl = SQl & ", NG3_A = " & ESDev.NG3_A
    SQl = SQl & ", NG4_A = " & ESDev.NG4_A
    SQl = SQl & ", NG5_A = " & ESDev.NG5_A
    SQl = SQl & ", NG6_A = " & ESDev.NG6_A


    SQl = SQl & ", NG1d = " & ESDev.NG1D
    SQl = SQl & ", NG2d = " & ESDev.NG2D
    SQl = SQl & ", NG3d = " & ESDev.NG3D
    SQl = SQl & ", NG4d = " & ESDev.NG4D
    SQl = SQl & ", NG5d = " & ESDev.NG5D
    SQl = SQl & ", NG6d = " & ESDev.NG6D


    SQl = SQl & ", NG1_AD = " & ESDev.NG1_AD
    SQl = SQl & ", NG2_AD = " & ESDev.NG2_AD
    SQl = SQl & ", NG3_AD = " & ESDev.NG3_AD
    SQl = SQl & ", NG4_AD = " & ESDev.NG4_AD
    SQl = SQl & ", NG5_AD = " & ESDev.NG5_AD
    SQl = SQl & ", NG6_AD = " & ESDev.NG6_AD

    SQl = SQl & ", gG1 = " & ESDev.GG1
    SQl = SQl & ", gG2 = " & ESDev.GG2
    SQl = SQl & ", gG3 = " & ESDev.GG3
    SQl = SQl & ", gG4 = " & ESDev.GG4
    SQl = SQl & ", gG5 = " & ESDev.GG5
    SQl = SQl & ", gG6 = " & ESDev.GG6


    SQl = SQl & ", gG1_A = " & ESDev.GG1_A
    SQl = SQl & ", gG2_A = " & ESDev.GG2_A
    SQl = SQl & ", gG3_A = " & ESDev.GG3_A
    SQl = SQl & ", gG4_A = " & ESDev.GG4_A
    SQl = SQl & ", gG5_A = " & ESDev.GG5_A
    SQl = SQl & ", gG6_A = " & ESDev.GG6_A


    SQl = SQl & ", gG1d = " & ESDev.GG1D
    SQl = SQl & ", gG2d = " & ESDev.GG2D
    SQl = SQl & ", gG3d = " & ESDev.GG3D
    SQl = SQl & ", gG4d = " & ESDev.GG4D
    SQl = SQl & ", gG5d = " & ESDev.GG5D
    SQl = SQl & ", gG6d = " & ESDev.GG6D


    SQl = SQl & ", gG1_AD = " & ESDev.GG1_AD
    SQl = SQl & ", gG2_AD = " & ESDev.GG2_AD
    SQl = SQl & ", gG3_AD = " & ESDev.GG3_AD
    SQl = SQl & ", gG4_AD = " & ESDev.GG4_AD
    SQl = SQl & ", gG5_AD = " & ESDev.GG5_AD
    SQl = SQl & ", gG6_AD = " & ESDev.GG6_AD


    SQl = SQl & " WHERE model = " & q(Model)

    ConnExecute SQl  ' update all devices of this model



    Dim counter     As Long

    If MASTER Then
      If USE6080 Then
        ' update all checkin times for this devicetype in loop below
        ' get list of devices that match this devicetypoe
      End If
    End If

    For Each Device In Devices.Devices
      counter = counter + 1
      If 0 = (counter Mod 100) Then
        DoEvents
      End If


      If 0 = StrComp(Device.Model, Model, vbTextCompare) Then

        If MASTER Then
          If USE6080 Then
            ' update all checkin times for this devicetype



          End If
        End If

        ' main tab
        Device.AutoClear = ESDev.AutoClear
        Device.Description = ESDev.desc
        Device.Custom = ESDev.desc
        Device.SupervisePeriod = ESDev.Checkin
        Device.IgnoreTamper = ESDev.IgnoreTamper

        'Input1 tab
        Device.Announce = ESDev.Announce
        Device.Repeats = ESDev.Repeats  ' # of repeats
        Device.ClearByReset = ESDev.ClearByReset
        Device.Pause = ESDev.Pause
        Device.SendCancel = ESDev.SendCancel    ' Send cancel
        Device.RepeatUntil = ESDev.RepeatUntil  ' until reset ?

        'Input2 tab
        Device.Announce_A = ESDev.Announce2
        Device.Repeats_A = ESDev.Repeats_A
        Device.Pause_A = ESDev.Pause_A
        Device.SendCancel_A = ESDev.SendCancel_A    ' Send cancel
        Device.RepeatUntil_A = ESDev.RepeatUntil_A  ' until reset ?


        Device.OG1 = ESDev.OG1
        Device.OG2 = ESDev.OG2
        Device.OG3 = ESDev.OG3
        Device.OG4 = ESDev.OG4
        Device.OG5 = ESDev.OG5
        Device.OG6 = ESDev.OG6


        Device.OG1D = ESDev.OG1D
        Device.OG2D = ESDev.OG2D
        Device.OG3D = ESDev.OG3D
        Device.OG4D = ESDev.OG4D
        Device.OG5D = ESDev.OG5D
        Device.OG6D = ESDev.OG6D


        Device.OG1_A = ESDev.OG1_A
        Device.OG2_A = ESDev.OG2_A
        Device.OG3_A = ESDev.OG3_A
        Device.OG4_A = ESDev.OG4_A
        Device.OG5_A = ESDev.OG5_A
        Device.OG6_A = ESDev.OG6_A

        Device.OG1_AD = ESDev.OG1_AD
        Device.OG2_AD = ESDev.OG2_AD
        Device.OG3_AD = ESDev.OG3_AD
        Device.OG4_AD = ESDev.OG4_AD
        Device.OG5_AD = ESDev.OG5_AD
        Device.OG6_AD = ESDev.OG6_AD

        Device.NG1 = ESDev.NG1
        Device.NG2 = ESDev.NG2
        Device.NG3 = ESDev.NG3
        Device.NG4 = ESDev.NG4
        Device.NG5 = ESDev.NG5
        Device.NG6 = ESDev.NG6


        Device.NG1D = ESDev.NG1D
        Device.NG2D = ESDev.NG2D
        Device.NG3D = ESDev.NG3D
        Device.NG4D = ESDev.NG4D
        Device.NG5D = ESDev.NG5D
        Device.NG6D = ESDev.NG6D

        Device.NG1_A = ESDev.NG1_A
        Device.NG2_A = ESDev.NG2_A
        Device.NG3_A = ESDev.NG3_A
        Device.NG4_A = ESDev.NG4_A
        Device.NG5_A = ESDev.NG5_A
        Device.NG6_A = ESDev.NG6_A

        Device.NG1_AD = ESDev.NG1_AD
        Device.NG2_AD = ESDev.NG2_AD
        Device.NG3_AD = ESDev.NG3_AD
        Device.NG4_AD = ESDev.NG4_AD
        Device.NG5_AD = ESDev.NG5_AD
        Device.NG6_AD = ESDev.NG6_AD


        Device.GG1 = ESDev.GG1
        Device.GG2 = ESDev.GG2
        Device.GG3 = ESDev.GG3
        Device.GG4 = ESDev.GG4
        Device.GG5 = ESDev.GG5
        Device.GG6 = ESDev.GG6


        Device.GG1D = ESDev.GG1D
        Device.GG2D = ESDev.GG2D
        Device.GG3D = ESDev.GG3D
        Device.GG4D = ESDev.GG4D
        Device.GG5D = ESDev.GG5D
        Device.GG6D = ESDev.GG6D

        Device.GG1_A = ESDev.GG1_A
        Device.GG2_A = ESDev.GG2_A
        Device.GG3_A = ESDev.GG3_A
        Device.GG4_A = ESDev.GG4_A
        Device.GG5_A = ESDev.GG5_A
        Device.GG6_A = ESDev.GG6_A

        Device.GG1_AD = ESDev.GG1_AD
        Device.GG2_AD = ESDev.GG2_AD
        Device.GG3_AD = ESDev.GG3_AD
        Device.GG4_AD = ESDev.GG4_AD
        Device.GG5_AD = ESDev.GG5_AD
        Device.GG6_AD = ESDev.GG6_AD




      End If
    Next

  End If
End Function

Public Function ResetDevice(rs As ADODB.Recordset, ByVal FastMode As Boolean) As Long
        Dim rst           As Recordset
        Dim d             As cESDevice
        Dim j             As Long
        Dim t             As Long

        Dim Serial        As String
        Dim r             As cResident
        Dim Room          As cRoom
        
10      t = Win32.timeGetTime


20      On Error GoTo ResetDevice_Error

        'Debug.Print " Resetting Device " & Serial

30      Serial = rs("Serial") & ""


40      Set d = Devices.Device(Serial)
50      If d Is Nothing Then
60        Set d = New cESDevice
70        d.Serial = Serial
80        Devices.AddDevice d
90        d.SupervisePeriod = gSupervisePeriod  ' default
          ' log add device
100     End If
        
        
110     d.Parse rs  ' parse also sets checkin timeout (supervise period)
        ' not needed
        '  d.FetchResident
        '  d.FetchRoom
        ' not needed

120     If d.IsSerialDevice() Then
130       If MASTER Then
140         SetupSerialDevice d
150       End If
160     End If


170     d.LastSupervise = Now

        'Debug.Print "ResetDevice 210 " & Serial & "  " & Win32.timeGetTime - t
180     Set r = Residents.Resident(d.ResidentID)
190     Set Room = Rooms.Room(d.RoomID)

      
200       If Not r Is Nothing Then
210         d.NameLast = r.NameLast
220         d.NameFirst = r.NameFirst
            
230         d.IsAway = IIf(r.Vacation, 1, 0)
240       End If
250       If Not Room Is Nothing Then
260         d.Room = Room.Room
270         d.IsAway = d.IsAway Or IIf(Room.Away = 1, 1, 0)
280       End If
        

        


290     Set r = Nothing
300     Set Room = Nothing

' REMOVED SINCE WE DON'T USE THESE

'310     If d.ResidentID_A <> 0 Then
'320       Set r = Residents.Resident(d.ResidentID_A)
'330       If Not r Is Nothing Then
'340         d.IsAway_A = IIf(r.Vacation, 1, 0)
'350       End If
'
'          '360 Set rs = ConnExecute("SELECT away FROM Residents WHERE ResidentID = " & d.ResidentID_A)
'          '370 If Not rs.EOF Then
'          '380 d.IsAway_A = IIf(rs("away") = 1, 1, 0)
'          '390 End If
'          '400 rs.Close
'360     ElseIf d.RoomID_A <> 0 Then
'
'
'370       Set Room = Rooms.Room(d.RoomID_A)
'380       If Not Room Is Nothing Then
'390         d.IsAway_A = d.IsAway_A Or IIf(Room.Away = 1, 1, 0)
'400       End If
'          '420 Set rs = ConnExecute("SELECT away FROM Rooms WHERE RoomID = " & d.RoomID_A)
'          '430 If Not rs.EOF Then
'          '440 d.IsAway_A = d.IsAway_A Or IIf(rs("away") = 1, 1, 0)
'          '450 End If
'          '460 rs.Close
'
'
'410     End If



        'Debug.Print "ResetDevice 470 " & Serial & "  " & Win32.timeGetTime - t
        ' log reset device


ResetDevice_Resume:
420     Set rst = Nothing

        'Debug.Print "ResetDevice Done " & Serial & "  " & Win32.timeGetTime - t

430     On Error GoTo 0
440     Exit Function

ResetDevice_Error:

450     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDevices.ResetDevice." & Erl
460     Resume ResetDevice_Resume

        '  If FastMode Then
        '
        'x130:
        '    For j = 0 To MAX_ESDEVICETYPES
        '
        '      If (0 = StrComp(d.Model, ESDeviceType(j).Model, vbTextCompare)) Then
        '        d.SupervisePeriod = ESDeviceType(j).Checkin
        '        If d.SupervisePeriod = 0 Then
        '          d.SupervisePeriod = 100
        '        End If
        '        d.LastSupervise = Now
        '        d.IsLatching = ESDeviceType(j).Latching
        '        d.Description = ESDeviceType(j).Desc
        '        d.IsPortable = ESDeviceType(j).Portable
        '        Exit For
        '      End If
        '    Next
        '
        'x140:
        '
        '  Else
        '
        '140 Set rst = ConnExecute("SELECT * FROM Devicetypes WHERE model = " & q(d.Model))
        '150 If Not rst.EOF Then
        '160   d.SupervisePeriod = Val(rst("checkin") & "")
        '      If d.SupervisePeriod = 0 Then
        '        d.SupervisePeriod = 100
        '      End If
        '      d.LastSupervise = Now
        '170   d.IsLatching = IIf(rst("islatching") = 1, 1, 0)
        '180   d.Description = rst("description") & ""
        '190   d.IsPortable = IIf(rst("isportable") = 1, 1, 0)
        '
        '200 End If
        '210 rst.Close
        '
        '  End If



End Function
Function SetupSerialDevice(d As cESDevice) As Long
  Dim j      As Integer
  Dim si     As cSerialInput
  
  RemoveSerialDevice d.Serial
  

  Set si = New cSerialInput
  si.Serial = d.Serial
  si.Port = d.SerialPort
  si.Baud = d.SerialBaud
  si.Parity = d.SerialParity
  si.DataBits = d.Serialbits
  si.Stopbits = d.SerialStopbits
  si.SerialTapProtocol = d.SerialTapProtocol
  si.Skip = d.SerialSkip
  si.PhraseLength = d.SerialMessageLen
  si.SetWords (d.SerialInclude)
  si.SetExclude (d.SerialExclude)
  si.EOLChar = d.SerialEOLChar
  si.Settings = "baud=" & si.Baud & " parity=" & si.Parity & " data=" & si.DataBits & " stop=" & si.Stopbits
  
  If MASTER Then
    If si.Port <> 0 Then
      si.CloseComm
      If si.start() = 0 Then
        SerialIns.Add si
      End If
    End If
  End If
End Function
Public Function RemoveSerialDevice(ByVal Serial As String) As Boolean
  Dim j As Integer
  For j = SerialIns.count To 1 Step -1
    If SerialIns(j).Serial = Serial Then
      SerialIns.Remove j
    End If
  Next
End Function


Public Function GetAnnounce(ByVal Model As String) As String
        Dim rs    As Recordset

10       On Error GoTo GetAnnounce_Error

20      Set rs = ConnExecute("SELECT Announce FROM DeviceTypes WHERE model = " & q(Model))
30      If Not rs.EOF Then
40        GetAnnounce = rs(0).Value & ""
50      End If
60      rs.Close


GetAnnounce_Resume:
70       On Error GoTo 0
80       Exit Function

GetAnnounce_Error:

90      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDevices.GetAnnounce." & Erl
100     Resume GetAnnounce_Resume


End Function

Public Function GetIsPortable(ByVal Model As String) As Integer
        Dim rs    As Recordset
        Dim Value As Integer
  
10       On Error GoTo GetIsPortable_Error

20      Set rs = ConnExecute("SELECT isportable FROM DeviceTypes WHERE model = " & q(Model))
30      If Not rs.EOF Then
40        Value = rs(0).Value
50        GetIsPortable = IIf(Value = 1, 0, 1)
60      End If
70      rs.Close

GetIsPortable_Resume:
80       On Error GoTo 0
90       Exit Function

GetIsPortable_Error:

100     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDevices.GetIsPortable." & Erl
110     Resume GetIsPortable_Resume

  
End Function
