Attribute VB_Name = "modPaging"
Option Explicit
'uses global variable: gShift      As Integer  ' 0 ,1 (or maybe up to 2 for Three shifts)

Public gPageDevices As New Collection
Public gPageRequests As New cPageItems
Public gPages As New Collection

''Public gPageOutputs    As New Collection




Global Const PAGER_NORMAL = 0
Global Const PAGER_CANCEL = 1
Global Const PAGER_SET_MARQUIS = 2
Global Const PAGER_CLEAR_MARQUIS = 4
Global Const PAGER_SET_RELAY = 8
Global Const PAGER_CLEAR_RELAY = 16
Global Const PAGER_SET_APOLLO = 32
Global Const PAGER_CLEAR_APOLLO = 64




Global Const MARQUIS_NONE = 0
Global Const MARQUIS_NORMAL = 1
Global Const MARQUIS_EMERGENCY = 2
Global Const MARQUIS_HELP = 3
Global Const MARQUIS_HELP_LAV = 4
Global Const MARQUIS_INFO = 5
Global Const MARQUIS_APOLLO = 6





'Need Global delay for page repeats
'Need per output device interpage delay

' Each Output Channel gets it's own que

' PageRequest is made
' shift night groups into daygroups if night shift

' Poll Loop (Polls PageRequest)

' if TimeToPage then
' PageRequest adds page to que for each Output Channel in group
' Note: Each output channel handles it's own intermessage delay

' if Escalate then
' PageRequest shifts groups left (OG1 = OG2, OG2 = OG3) at Escalate


'If alarm/lowbatt etc is reset or timesout then
' Need stop on reset
' End Poll Loop


' If repeatcount = 0 then
' escalate if OG1 is non ZERO
' Start repeatcount over again

' if RepeatUntil <> 0 then repeat forever

' If Output Channel has no repeats then
' Either Escalate or Delete


' If acknowledged, or restored, then
' search all PageRequest Ques for page and remove

Public gLastPageSerial As Long


Public Function GetPageDeviceByID(ByVal DeviceID As Long) As cPageDevice
  Dim pd As cPageDevice
  
  For Each pd In gPageDevices
    If pd.DeviceID = DeviceID Then
      Set GetPageDeviceByID = pd
      Exit For
    End If
  Next
  Set pd = Nothing

End Function

Sub UpdatePageLocation(ByVal ID As String, ByVal Location As String)
    Dim j As Integer
    Dim PageItem As cPageItem
    
    For j = 1 To gPageRequests.Count
      Set PageItem = gPageRequests.PageItem(j)
      PageItem.locationtext = Location
    Next

End Sub




Sub AddPageRequest(alarm As cAlarm, ByVal EventType As Long, Optional ByVal OutputGroupID As Long = 0)
        Dim PageItem    As cPageItem
        Dim d           As cESDevice
        Dim j           As Integer
        Dim Serial      As String
        Dim inputnum    As Long
        Dim Announce    As String
        Dim OutputMask  As cOutputMask
        Dim RoomText    As String

10      On Error GoTo AddPageRequest_Error

20      If alarm Is Nothing Then Exit Sub
30      Serial = alarm.Serial

40      inputnum = alarm.inputnum
        
50      For j = gPageRequests.Count To 1 Step -1
60        Set PageItem = gPageRequests.PageItem(j)
70        If PageItem.AlarmID = alarm.ID Then
            If EventType = PageItem.EventType Then
            ' may need to chage to this (2018-07-18)
            ' If EventType = PageItem.EventType and pageitem.InputNum = alarm.InputNumThen
                'dbg "Duplicate Alarm " & Alarm.ID
80
              Exit For
            End If
90        End If
100     Next


110     If j = 0 Then  ' we need to add it

          'dbg "Adding Alarm " & Alarm.ID

120       Set PageItem = New cPageItem
130       PageItem.AlarmID = alarm.ID
140       PageItem.Serial = alarm.Serial
150       PageItem.inputnum = alarm.inputnum
160       PageItem.EventType = EventType
          


170       Set d = Devices.Device(Serial)

180       If Not d Is Nothing Then
190         PageItem.Preamble = alarm.Preamble
200         PageItem.RoomText = alarm.RoomText
210         PageItem.ResidentText = alarm.ResidentText
220         PageItem.locationtext = alarm.locationtext
230         PageItem.Phone = alarm.Phone

240         Select Case EventType

              Case EVT_BATTERY_FAIL
250             Set OutputMask = GetOutputMask(SCREEN_BATTERY, "Battery Trouble")
260           Case EVT_CHECKIN_FAIL
270             Set OutputMask = GetOutputMask(SCREEN_TROUBLE, "Checkin Trouble")
280           Case EVT_COMM_TIMEOUT
290             Set OutputMask = GetOutputMask(SCREEN_TROUBLE, "System Trouble")
300           Case EVT_GENERAL_TROUBLE
310             Set OutputMask = GetOutputMask(SCREEN_TROUBLE, "General Trouble")
320           Case EVT_TAMPER
330             Set OutputMask = GetOutputMask(SCREEN_TROUBLE, "Device Tamper")
340           Case EVT_LINELOSS
350             Set OutputMask = GetOutputMask(SCREEN_TROUBLE, "Device NO AC")
360           Case Else
                
370             Set OutputMask = New cOutputMask
380             If inputnum = 3 Then
                
390               OutputMask.RepeatTwice = True
400               OutputMask.Announce = alarm.Announce
410               OutputMask.SendCancel = d.SendCancel_B
                  ' could be simplified with Alarm output groups
420               OutputMask.OG1 = d.OG1_B
430               OutputMask.OG2 = d.OG2_B
440               OutputMask.OG3 = d.OG3_B
450               OutputMask.OG4 = d.OG4_B
460               OutputMask.OG5 = d.OG5_B
470               OutputMask.OG6 = d.OG6_B

480               OutputMask.OG1D = d.OG1_BD
490               OutputMask.OG2D = d.OG2_BD
500               OutputMask.OG3D = d.OG3_BD
510               OutputMask.OG4D = d.OG4_BD
520               OutputMask.OG5D = d.OG5_BD
530               OutputMask.OG6D = d.OG6_BD


540               OutputMask.NG1 = d.NG1_B
550               OutputMask.NG2 = d.NG2_B
560               OutputMask.NG3 = d.NG3_B
570               OutputMask.NG4 = d.NG4_B
580               OutputMask.NG5 = d.NG5_B
590               OutputMask.NG6 = d.NG6_B

600               OutputMask.NG1D = d.NG1_BD
610               OutputMask.NG2D = d.NG2_BD
620               OutputMask.NG3D = d.NG3_BD
630               OutputMask.NG4D = d.NG4_BD
640               OutputMask.NG5D = d.NG5_BD
650               OutputMask.NG6D = d.NG6_BD


660               OutputMask.GG1 = d.GG1_B
670               OutputMask.GG2 = d.GG2_B
680               OutputMask.GG3 = d.GG3_B
690               OutputMask.GG4 = d.GG4_B
700               OutputMask.GG5 = d.GG5_B
710               OutputMask.GG6 = d.GG6_B

720               OutputMask.GG1D = d.GG1_BD
730               OutputMask.GG2D = d.GG2_BD
740               OutputMask.GG3D = d.GG3_BD
750               OutputMask.GG4D = d.GG4_BD
760               OutputMask.GG5D = d.GG5_BD
770               OutputMask.GG6D = d.GG6_BD




780               OutputMask.Pause = d.Pause_B  ' should be same as regular?
790               OutputMask.Repeats = d.Repeats_B
800               OutputMask.RepeatUntil = d.RepeatUntil_B
                
                
                
810             ElseIf inputnum = 2 Then
820               OutputMask.RepeatTwice = True
830               OutputMask.Announce = alarm.Announce
840               OutputMask.SendCancel = d.SendCancel_A
                  ' could be simplified with Alarm output groups
850               OutputMask.OG1 = d.OG1_A
860               OutputMask.OG2 = d.OG2_A
870               OutputMask.OG3 = d.OG3_A
880               OutputMask.OG4 = d.OG4_A
890               OutputMask.OG5 = d.OG5_A
900               OutputMask.OG6 = d.OG6_A

910               OutputMask.OG1D = d.OG1_AD
920               OutputMask.OG2D = d.OG2_AD
930               OutputMask.OG3D = d.OG3_AD
940               OutputMask.OG4D = d.OG4_AD
950               OutputMask.OG5D = d.OG5_AD
960               OutputMask.OG6D = d.OG6_AD


970               OutputMask.NG1 = d.NG1_A
980               OutputMask.NG2 = d.NG2_A
990               OutputMask.NG3 = d.NG3_A
1000              OutputMask.NG4 = d.NG4_A
1010              OutputMask.NG5 = d.NG5_A
1020              OutputMask.NG6 = d.NG6_A

1030              OutputMask.NG1D = d.NG1_AD
1040              OutputMask.NG2D = d.NG2_AD
1050              OutputMask.NG3D = d.NG3_AD
1060              OutputMask.NG4D = d.NG4_AD
1070              OutputMask.NG5D = d.NG5_AD
1080              OutputMask.NG6D = d.NG6_AD


1090              OutputMask.GG1 = d.GG1_A
1100              OutputMask.GG2 = d.GG2_A
1110              OutputMask.GG3 = d.GG3_A
1120              OutputMask.GG4 = d.GG4_A
1130              OutputMask.GG5 = d.GG5_A
1140              OutputMask.GG6 = d.GG6_A

1150              OutputMask.GG1D = d.GG1_AD
1160              OutputMask.GG2D = d.GG2_AD
1170              OutputMask.GG3D = d.GG3_AD
1180              OutputMask.GG4D = d.GG4_AD
1190              OutputMask.GG5D = d.GG5_AD
1200              OutputMask.GG6D = d.GG6_AD




1210              OutputMask.Pause = d.Pause_A  ' should be same as regular?
1220              OutputMask.Repeats = d.Repeats_A
1230              OutputMask.RepeatUntil = d.RepeatUntil_A
                  'pageitem.CancelText = Alarm.CancelText ' not used for now
1240            Else  ' input # 1
1250              OutputMask.RepeatTwice = True
1260              OutputMask.Announce = alarm.Announce
1270              OutputMask.SendCancel = d.SendCancel

1280              OutputMask.OG1 = d.OG1
1290              OutputMask.OG2 = d.OG2
1300              OutputMask.OG3 = d.OG3
1310              OutputMask.OG4 = d.OG4
1320              OutputMask.OG5 = d.OG5
1330              OutputMask.OG6 = d.OG6

1340              OutputMask.OG1D = d.OG1D
1350              OutputMask.OG2D = d.OG2D
1360              OutputMask.OG3D = d.OG3D
1370              OutputMask.OG4D = d.OG4D
1380              OutputMask.OG5D = d.OG5D
1390              OutputMask.OG6D = d.OG6D


1400              OutputMask.NG1 = d.NG1
1410              OutputMask.NG2 = d.NG2
1420              OutputMask.NG3 = d.NG3
1430              OutputMask.NG4 = d.NG4
1440              OutputMask.NG5 = d.NG5
1450              OutputMask.NG6 = d.NG6


1460              OutputMask.NG1D = d.NG1D
1470              OutputMask.NG2D = d.NG2D
1480              OutputMask.NG3D = d.NG3D
1490              OutputMask.NG4D = d.NG4D
1500              OutputMask.NG5D = d.NG5D
1510              OutputMask.NG6D = d.NG6D


1520              OutputMask.GG1 = d.GG1
1530              OutputMask.GG2 = d.GG2
1540              OutputMask.GG3 = d.GG3
1550              OutputMask.GG4 = d.GG4
1560              OutputMask.GG5 = d.GG5
1570              OutputMask.GG6 = d.GG6

1580              OutputMask.GG1D = d.GG1D
1590              OutputMask.GG2D = d.GG2D
1600              OutputMask.GG3D = d.GG3D
1610              OutputMask.GG4D = d.GG4D
1620              OutputMask.GG5D = d.GG5D
1630              OutputMask.GG6D = d.GG6D



1640              OutputMask.Pause = d.Pause
1650              OutputMask.Repeats = d.Repeats
1660              OutputMask.RepeatUntil = d.RepeatUntil

                  'pageitem.CancelText = Alarm.CancelText ' not used for now
1670            End If
1680        End Select
            
1690        If Not OutputMask Is Nothing Then
1700          PageItem.RepeatTwice = OutputMask.RepeatTwice
1710          PageItem.Announce = OutputMask.Announce
1720          PageItem.SendCancel = OutputMask.SendCancel
              PageItem.inputnum = inputnum

1730          PageItem.OG1 = OutputMask.OG1
1740          PageItem.OG2 = OutputMask.OG2
1750          PageItem.OG3 = OutputMask.OG3
1760          PageItem.OG4 = OutputMask.OG4
1770          PageItem.OG5 = OutputMask.OG5
1780          PageItem.OG6 = OutputMask.OG6

1790          PageItem.OG1D = OutputMask.OG1D
1800          PageItem.OG2D = OutputMask.OG2D
1810          PageItem.OG3D = OutputMask.OG3D
1820          PageItem.OG4D = OutputMask.OG4D
1830          PageItem.OG5D = OutputMask.OG5D
1840          PageItem.OG6D = OutputMask.OG6D



1850          PageItem.NG1 = OutputMask.NG1
1860          PageItem.NG2 = OutputMask.NG2
1870          PageItem.NG3 = OutputMask.NG3
1880          PageItem.NG4 = OutputMask.NG4
1890          PageItem.NG5 = OutputMask.NG5
1900          PageItem.NG6 = OutputMask.NG6


1910          PageItem.NG1D = OutputMask.NG1D
1920          PageItem.NG2D = OutputMask.NG2D
1930          PageItem.NG3D = OutputMask.NG3D
1940          PageItem.NG4D = OutputMask.NG4D
1950          PageItem.NG5D = OutputMask.NG5D
1960          PageItem.NG6D = OutputMask.NG6D


1970          PageItem.GG1 = OutputMask.GG1
1980          PageItem.GG2 = OutputMask.GG2
1990          PageItem.GG3 = OutputMask.GG3
2000          PageItem.GG4 = OutputMask.GG4
2010          PageItem.GG5 = OutputMask.GG5
2020          PageItem.GG6 = OutputMask.GG6


2030          PageItem.GG1D = OutputMask.GG1D
2040          PageItem.GG2D = OutputMask.GG2D
2050          PageItem.GG3D = OutputMask.GG3D
2060          PageItem.GG4D = OutputMask.GG4D
2070          PageItem.GG5D = OutputMask.GG5D
2080          PageItem.GG6D = OutputMask.GG6D



2090          PageItem.Pause = OutputMask.Pause
2100          PageItem.Repeats = OutputMask.Repeats
2110          PageItem.RepeatUntil = OutputMask.RepeatUntil

              If OutputGroupID <> 0 Then
                PageItem.OutputGroupID = OutputGroupID
              End If
2120          PageItem.Init
2130          gPageRequests.AddPageItem PageItem
              'dbg "Adding pageitem in ModPaging.AddPageRequest line 785 " & OutputMask.Verbose()
2140        End If  'Not OutputMask Is Nothing Then
2150      End If  'Not d Is Nothing
2160    End If  'j = 0


AddPageRequest_Resume:
2170    On Error GoTo 0
2180    Exit Sub

AddPageRequest_Error:

2190    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.AddPageRequest." & Erl
2200    Resume AddPageRequest_Resume


End Sub

Public Sub AddAssistanceRequest(ByVal GroupID As Long)

'
'  On Error GoTo AddAssistanceRequest_Error
'  ' Get Group Settings
'  Dim SQL As String
'  Dim rsGroup As ADODB.Recordset
'  Dim PageItem As cPageItem
'
'  SQL = "SELECT * FROM Groups WHERE Groupid = " & GroupID
'  Set rsGroup = ConnExecute(SQL)
'  Do While Not rsGroup.EOF
'
'      ' for j = 1 to grouppagers.
'
'
'        Set PageItem = New cPageItem
'
''        pageitem.Pause = OutputMask.Pause
''
''        pageitem.Repeats = OutputMask.Repeats
''        pageitem.RepeatUntil = OutputMask.RepeatUntil
''        pageitem.Init
''        gPageRequests.AddPageItem pageitem
'
'      Exit Do
'  Loop
'  rsGroup.Close
'  Set rsGroup = Nothing
'
'
'
'
'
'AddAssistanceRequest_Resume:
'
'  On Error GoTo 0
'  Exit Sub
'
'AddAssistanceRequest_Error:
'
'  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.AddAssistanceRequest." & Erl
'  Resume AddAssistanceRequest_Resume
End Sub


'Public Sub SendEndofEventPage(pageitem As cPageItem, ByVal EventType As Long)
'
'  Dim d           As cESDevice
'  Dim GroupID     As Long
'
'  dbg "modpaging.SendEndofEventPage NOT USED???"
'  Exit Sub
'
'
'
'  If Not pageitem Is Nothing Then
'    Set d = Devices.Device(pageitem.Serial)
'    If Not d Is Nothing Then
''      If GetCurrentShift = 1 Then ' night shift
''        GroupID = pageitem.LastGroup
''      Else
''        GroupID = pageitem.LastGroup
''      End If
'
'      'SendToGroup pageitem.Message & " Cancel", GroupID, "", ""
'      SendToGroup pageitem.Message, GroupID, "", "", PAGER_CANCEL, pageitem.MarquisMessage
'
'      'SendToGroup pageitem.Message & " Cancel", GroupID, "", ""  ' needed to shorten up to fit PCA
'    End If
'  End If
'End Sub
Function RemovePageRequest(ByVal AlarmID As Long) As cPageItem  ' ByVal Serial As String, ByVal EventType As Long, ByVal InputNum As Long) As cPageItem
      '' This is where we weill remove marquis messages and ONTRAK relay settings

        Dim j As Integer
        Dim PageItem As cPageItem
'        Dim Group As Variant


10      On Error GoTo RemovePageRequest_Error

20      For j = gPageRequests.Count To 1 Step -1
30        Set PageItem = gPageRequests.PageItem(j)
40        If PageItem.AlarmID = AlarmID Then  'If (pageitem.Serial = Serial) And (pageitem.EventType = EventType) And (pageitem.InputNum = InputNum) Then

50          Set RemovePageRequest = PageItem
60          gPageRequests.Remove j
70          Exit For
80        End If
90      Next


RemovePageRequest_Resume:
100     On Error GoTo 0
110     Exit Function

RemovePageRequest_Error:

120     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.RemovePageRequest." & Erl
130     Resume RemovePageRequest_Resume


End Function



Public Sub CheckPageRequests()
      ' called from polling clock once a second

        Dim PageItem  As cPageItem
        Dim j         As Integer
        Dim GroupID   As Long


10      On Error GoTo CheckPageRequests_Error
        'On Error GoTo 0
        ' not sure why this is a count-down

20      For j = gPageRequests.Count To 1 Step -1
30        Set PageItem = gPageRequests.PageItem(j)
40        PageItem.Send
50      Next

CheckPageRequests_Resume:
60      On Error GoTo 0
70      Exit Sub

CheckPageRequests_Error:

80      LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.CheckPageRequests." & Erl
90      Resume CheckPageRequests_Resume

End Sub


'Public Sub SendPageItemToGroup(pageitem As cPageItem, ByVal GroupID As Long)
'   'no longer Called from CheckPageRequests
'  ' ??? resurrected from DEAD 7/13/2018 ???
'
'   ' groupid is not needed anymore, now we need it ???
'
'        Dim rs As Recordset
'        Dim PagerID As Long
'
'10      On Error GoTo SendPageItemToGroup_Error
'
'20      'dbg "modpaging.SendPageItemToGroup pageitem.NumPages, groupid: " & pageitem.NumPages & " ," & GroupID
'
'30      Set rs = ConnExecute("SELECT * FROM GroupPager WHERE groupID = " & GroupID)
'
'40      Do Until rs.EOF
'50        PagerID = Val("0" & rs("pagerID"))
'60        If PagerID <> 0 Then
'
'70          If pageitem.RepeatTwice Then                        ' false if already sent twice to TTS for this message
'80            If PagerDeviceRepeatTwice(PagerID) Then
'90              pageitem.RepeatTwice = False                    ' don't send twice anymore
'100             SendToPager pageitem.message, PagerID, 1, pageitem.Phone, pageitem.RoomText, PAGER_NORMAL, pageitem.MarquisMessage  ' send it again in case they didn't hear it
'110           End If
'120         End If
'
'130         If PagerDeviceNoRepeats(PagerID) Then
'
'140           Debug.Print "PagerDeviceNoRepeats pagerid: "; PagerID
'150           If pageitem.SentOnce Then
'160             dbg "Already Sent " & pageitem.NumPages
'170           Else
'180             dbg "First Send"
'190             SendToPager pageitem.message, PagerID, 0, pageitem.Phone, pageitem.RoomText, PAGER_NORMAL, pageitem.MarquisMessage       ' always send it if match
'200           End If
'
'210         Else
'220           SendToPager pageitem.message, PagerID, 0, pageitem.Phone, pageitem.RoomText, PAGER_NORMAL, pageitem.MarquisMessage       ' always send it if match
'230         End If
'
'240       End If
'250       rs.MoveNext
'260     Loop
'270     pageitem.SentOnce = True
'280     rs.Close
'
'SendPageItemToGroup_Resume:
'290     On Error GoTo 0
'300     Exit Sub
'
'SendPageItemToGroup_Error:
'
'310     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.SendPageItemToGroup." & Erl
'320     Resume SendPageItemToGroup_Resume
'
'
'End Sub




'Function PagerDeviceRepeatTwice(ByVal PagerID As Long) As Boolean
'  Dim rs As Recordset
'  Dim sql As String
'
'  sql = " SELECT  PagerDevices.Twice FROM Pagers INNER JOIN PagerDevices ON Pagers.DeviceID = PagerDevices.ID WHERE Pagers.PagerID = " & PagerID
'  Set rs = connexecute(sql)
'  If Not rs.EOF Then
'    PagerDeviceRepeatTwice = IIf(rs("Twice") = 1, True, False)
'  End If
'  rs.Close
'  Set rs = Nothing
'End Function


'Function PagerDeviceNoRepeats(ByVal PagerID As Long) As Boolean
'  Dim rs As Recordset
'  Dim sql As String
'
'  sql = " SELECT  NoRepeats FROM Pagers WHERE Pagers.PagerID = " & PagerID
'  Set rs = connexecute(sql)
'  If Not rs.EOF Then
'    PagerDeviceNoRepeats = IIf(rs("NoRepeats") = 1, True, False)
'  End If
'  rs.Close
'  Set rs = Nothing
'
'End Function







Public Sub SendToGroup(ByVal message As String, ByVal GroupID As Long, ByVal Phone As String, ByVal RoomText As String, _
          ByVal Mode As Integer, ByVal MaquisMessage As String, ByVal PagerID As Long, ByVal InputNumber As Long)
         ' gets all pagers assigned to this gorup and sends a message
         ' needs a little modification maybe
         
          'dbg "Modpaging SendToGroup May need this for announce?"
          
          
          Dim rs As Recordset
10        On Error GoTo SendToGroup_Error

20        Set rs = ConnExecute("SELECT * FROM GroupPager WHERE groupID = " & GroupID)

30        Do Until rs.EOF
40          SendToPager message, rs("pagerID"), 0, Phone, RoomText, Mode, MaquisMessage, PagerID, InputNumber
50          rs.MoveNext
60        Loop
70        rs.Close

SendToGroup_Resume:
80         On Error GoTo 0
90         Exit Sub

SendToGroup_Error:

100       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.SendToGroup." & Erl
110       Resume SendToGroup_Resume



End Sub

Public Sub SendToPager(ByVal message As String, ByVal PagerID As Long, ByVal NoWait As Integer, ByVal Phone As String, ByVal RoomText As String, _
        ByVal Mode As Integer, ByVal MarquisMessage As String, ByVal AlarmID As Long, ByVal InputNumber As Long)

        Dim PageItem          As New cPageItem
        Dim pageDevice          As cPageDevice
        Dim rs          As Recordset
        Dim DeviceID    As Long
        Dim Address     As String
        Dim NoRepeats   As Integer
        Dim Sendphone   As Integer
        Dim MarquisCode As Integer
        Dim MarquisChar As String
        Dim RelayID     As Integer
        Dim NoName      As Long
        Dim NamedGroup   As String
        
        
        'Debug.Print " modPaging.SendToPager " & message
        

10      On Error GoTo SendToPager_Error

        'dbg "modpaging.sendtopager message, pagerid: " & Message & " ," & PagerID
        'dbg "modpaging.sendtopager Fix Me, get from the output, not database"
20      Set rs = ConnExecute("SELECT * FROM pagers WHERE pagerid = " & PagerID)

30      If Not rs.EOF Then
40        DeviceID = rs("deviceID")
50        Address = rs("identifier") & ""
60        NoName = rs("noname")
70        NoRepeats = IIf(rs("NoRepeats") = 1, 1, 0)
80        Sendphone = IIf(rs("IncludePhone") = 1, 1, 0)

90        RelayID = Val(rs("relaynum") & "")
          NamedGroup = rs("Pin") & ""
          
100       MarquisCode = Val(rs("MarquisCode") & "")

110       If MarquisCode = MARQUIS_APOLLO Then
120         NoRepeats = 1
130         Sendphone = 0
140         MarquisChar = ""
150         If Mode = PAGER_NORMAL Then
160           Mode = PAGER_SET_APOLLO
170         End If

180       ElseIf (MarquisCode > MARQUIS_NONE) And (MarquisCode <> MARQUIS_APOLLO) Then
190         If Mode = PAGER_NORMAL Then
200           Mode = PAGER_SET_MARQUIS
210         End If
            
220         NoRepeats = 1
230         Sendphone = 0
240         MarquisChar = MarquiCode2MarquiChar(MarquisCode)
250       End If
260     End If
270     rs.Close
280     Set rs = Nothing

290     For Each pageDevice In gPageDevices
300       If pageDevice.DeviceID = DeviceID Then
            
310         Set PageItem = New cPageItem

            'PageItem.AssistRequest = AssistRequest

            
320         PageItem.AlarmID = AlarmID
330         PageItem.RelayID = RelayID
            'PageItem.OutputGroupID = GroupID ' output groupID of originating group
            
340         PageItem.IsCancel = (Mode = PAGER_CANCEL)
            PageItem.NamedGroup = NamedGroup
            PageItem.inputnum = InputNumber
            
            
350         Select Case Mode

              Case PAGER_SET_APOLLO
360             If (MarquisCode = MARQUIS_APOLLO) Then
                  'message = left$(Trim$(MarquisMessage), 19)
370               PageItem.message = MarquisMessage & MarquisChar
380               PageItem.Address = Address
390               PageItem.NoWait = NoWait
400               PageItem.RoomText = ""
410               pageDevice.AddPage PageItem
                  'Trace "APOLLO ON: " & pageItem.Message, True
420             End If
430           Case PAGER_CLEAR_APOLLO
440             If (MarquisCode = MARQUIS_APOLLO) Then
                  'message = left$(Trim$(MarquisMessage), 19)
450               PageItem.message = "RESET " & MarquisMessage
460               PageItem.Address = Address
470               PageItem.NoWait = NoWait
480               PageItem.RoomText = ""
490               pageDevice.AddPage PageItem
                  'Trace "APOLLO OFF: " & pageItem.Message, True
500             End If


510           Case PAGER_SET_MARQUIS
520             If (MarquisCode > 0) Then
530               message = left$(Trim$(MarquisMessage), 19)
540               PageItem.message = message & MarquisChar
550               PageItem.Address = Address
560               PageItem.NoWait = NoWait
570               PageItem.RoomText = ""
580               pageDevice.AddPage PageItem
                  'Trace "MARQUIS  ON: " & pageItem.Message, True
590             End If
600           Case PAGER_CLEAR_MARQUIS
610             If (MarquisCode > 0) Then
620               message = left$(Trim$(MarquisMessage), 19)
630               PageItem.message = "RESET " & message & MarquisChar
640               PageItem.Address = Address
650               PageItem.NoWait = NoWait
660               PageItem.RoomText = ""
670               pageDevice.AddPage PageItem
                  'Trace "MARQUIS OFF: " & pageItem.Message, True

              

680             End If
690           Case PAGER_SET_RELAY
                ' not implemented yet
700           Case PAGER_CLEAR_RELAY
                ' not implemented yet
                
                
710           Case PAGER_CANCEL
720             If (MarquisCode = 0) Then
730               If Sendphone Then
740                 PageItem.message = message & " Cancel " & IIf(Len(Phone) > 0, " Phone: " & Phone, "")
750               Else
760                 PageItem.message = message & " Cancel "
770               End If
780               PageItem.Address = Address
790               PageItem.NoWait = NoWait
800               PageItem.RoomText = RoomText
810               pageDevice.AddPage PageItem

820             End If

830           Case Else
840             If (MarquisCode = 0) Then
850               If Sendphone Then
860                 PageItem.message = message & IIf(Len(Phone) > 0, " Phone: " & Phone, "")
870               Else
880                 PageItem.message = message
890               End If
900               PageItem.RelayID = RelayID
910               PageItem.Address = Address
920               PageItem.NoWait = NoWait
930               PageItem.RoomText = RoomText
                  
940               pageDevice.AddPage PageItem

950             End If


960         End Select

970         Exit For
980       End If
990     Next


SendToPager_Resume:
1000    On Error GoTo 0
1010    Exit Sub

SendToPager_Error:

1020    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.SendToPager." & Erl
1030    Resume SendToPager_Resume




End Sub


Public Function GetSystemTimeString() As String
  GetSystemTimeString = "The Time is " & ConvertDatetoTTS(Format(Now, gTimeFormatString))   ' AM/PM"))
End Function

Public Function ConvertDatetoTTS(ByVal text As String) As String

  Dim Parts() As String
  Dim hrPart  As String
  Dim MinPart As String
  Dim ampm    As String


10         On Error GoTo ConvertDatetoTTS_Error

20        If Year(text) <= 1900 Then  ' time only
30          text = Format(text, "h nn AM/PM")
40          Parts = Split(text, " ")
50          hrPart = Val(Parts(0))
60          MinPart = Val(Parts(1))
70          ampm = Parts(2)
80          If Val(hrPart) > 12 Then
90            hrPart = Val(hrPart) - 12
100         End If
110         Select Case Val(MinPart)
           Case 0
120             MinPart = ""
130           Case 1 To 9
140             MinPart = "O " & MinPart
150           Case Else  ' no change
160         End Select
170         ConvertDatetoTTS = hrPart & " " & MinPart & " " & IIf(ampm = "PM", "P M ", "A M ")
180       Else
190         ConvertDatetoTTS = text
200       End If

ConvertDatetoTTS_Resume:
210        On Error GoTo 0
220        Exit Function

ConvertDatetoTTS_Error:

230       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.ConvertDatetoTTS." & Erl
240       Resume ConvertDatetoTTS_Resume

End Function

Public Function ChangePageDevice(ByVal DeviceID As Long, ByVal Action As Integer) As Long
        Dim pd As cPageDevice
        Dim rs As Recordset
        Dim j   As Integer

10      On Error GoTo ChangePageDevice_Error

20      Select Case Action
          Case 1  ' add
30          For j = 1 To gPageDevices.Count
40            Set pd = gPageDevices(j)
50            If pd.DeviceID = DeviceID Then
60              pd.CloseConnection
70              gPageDevices.Remove j
                '        Set pd = Nothing
80              Exit For
90            End If
100         Next

110         Set rs = ConnExecute("Select * from PagerDevices where id = " & DeviceID)
120         If Not rs.EOF Then
130           Set pd = New cPageDevice
140           pd.DeviceID = rs("ID")
150           gPageDevices.Add pd
160           pd.ProtocolID = Val(rs("protocolid") & "")
170           pd.AudioDevice = rs("AudioDevice") & ""
180           pd.BaudRate = rs("baudrate") & ""
190           pd.BITS = rs("bits") & ""
200           pd.Description = rs("Description") & ""
210           pd.Parity = rs("parity") & ""
220           pd.Port = Val(rs("Port") & "")
230           pd.Settings = rs("settings") & ""
240           pd.Pause = rs("Pause")
              pd.LFeeds = Val(rs("lf") & "")
              ' MARQUIS
250           pd.MarquisControlCode = Max(0, Val(rs("MarquisCode") & ""))




              ' ONTRAK RELAY
260           pd.Relay1 = Max(0, Val(rs("Relay1") & ""))
270           pd.Relay2 = Max(0, Val(rs("Relay2") & ""))
280           pd.Relay3 = Max(0, Val(rs("Relay3") & ""))
290           pd.Relay4 = Max(0, Val(rs("Relay4") & ""))
300           pd.Relay5 = Max(0, Val(rs("Relay5") & ""))
310           pd.Relay6 = Max(0, Val(rs("Relay6") & ""))
320           pd.Relay7 = Max(0, Val(rs("Relay7") & ""))
330           pd.Relay8 = Max(0, Val(rs("Relay8") & ""))

              ' TTS
340           pd.PASystemKey = IIf(rs("KeyPA") = 1, 1, 0)
350           pd.PARepeatTwice = IIf(rs("Twice") = 1, 1, 0)
              ' DIALER
360           pd.DialerModem = rs("DialerModem")
370           pd.DialerMsgDelay = rs("DialerMsgDelay")
380           pd.DialerMsgRepeats = rs("DialerMsgRepeats")
390           pd.DialerMsgSpacing = rs("DialerMsgSpacing")
400           pd.DialerPhone = rs("DialerPhone") & ""
410           pd.DialerRedialDelay = rs("DialerRedialDelay")
420           pd.DialerRedials = rs("DialerRedials")
430           pd.DialerTag = rs("DialerTag") & ""
440           pd.DialerVoice = rs("DialerVoice") & ""
          

              '355           'pd.DialerTerminateDigit = Val(rs("DialerTerminateDigit") & "")
450           pd.DialerTerminateDigit = Val(rs("DialerAckDigit") & "")
              pd.KeepPaging = Val(rs("KeepPaging") & "") And 1
  
460           pd.OpenConnection

470           pd.Checked = True
480         End If
490         rs.Close
500         Set rs = Nothing


510       Case 3  ' remove
520         For j = 1 To gPageDevices.Count
530           Set pd = gPageDevices(j)
540           If pd.DeviceID = DeviceID Then
550             pd.CloseConnection
560             gPageDevices.Remove j
570             Set pd = Nothing
580             Exit For
590           End If
600         Next


610       Case 2  ' edit
620         For j = 1 To gPageDevices.Count
630           Set pd = gPageDevices(j)
640           If pd.DeviceID = DeviceID Then
650             pd.CloseConnection
660             gPageDevices.Remove j
670             Set pd = Nothing
680             Exit For
690           End If
700         Next

710         Set rs = ConnExecute("Select * from PagerDevices where id = " & DeviceID)
720         If Not rs.EOF Then
730           Set pd = New cPageDevice
740           pd.DeviceID = rs("ID")
750           gPageDevices.Add pd
760           pd.ProtocolID = Val(rs("protocolid") & "")
770           pd.AudioDevice = rs("AudioDevice") & ""
780           pd.BaudRate = rs("baudrate") & ""
790           pd.BITS = rs("bits") & ""
800           pd.Description = rs("Description") & ""
810           pd.Parity = rs("parity") & ""
820           pd.Port = Val(rs("Port") & "")
830           pd.Settings = rs("settings") & ""

              'pd.OpenConnection ??

              ' TTS
840           pd.PASystemKey = IIf(rs("keypa") = 1, 1, 0)
850           pd.PARepeatTwice = IIf(rs("Twice") = 1, 1, 0)
860           pd.Pause = rs("Pause")
              pd.LFeeds = Val(rs("lf") & "")
              ' DIALER
870           pd.DialerModem = rs("DialerModem")  ' long, unique device integer
880           pd.DialerMsgDelay = rs("DialerMsgDelay")
890           pd.DialerMsgRepeats = rs("DialerMsgRepeats")
900           pd.DialerMsgSpacing = rs("DialerMsgSpacing")
910           pd.DialerPhone = rs("DialerPhone") & ""
920           pd.DialerRedialDelay = rs("DialerRedialDelay")
930           pd.DialerRedials = rs("DialerRedials")
940           pd.DialerTag = rs("DialerTag") & ""
950           pd.DialerVoice = rs("DialerVoice") & ""
960           pd.DialerTerminateDigit = Val(rs("DialerAckDigit") & "")
              pd.KeepPaging = Val(rs("KeepPaging") & "") And 1

              ' MARQUIS
970           pd.MarquisControlCode = Max(0, Val(rs("MarquisCode") & ""))

              ' ONTRAK RELAY
980           pd.Relay1 = Max(0, Val(rs("Relay1") & ""))
990           pd.Relay2 = Max(0, Val(rs("Relay2") & ""))
1000          pd.Relay3 = Max(0, Val(rs("Relay3") & ""))
1010          pd.Relay4 = Max(0, Val(rs("Relay4") & ""))
1020          pd.Relay5 = Max(0, Val(rs("RelaY5") & ""))
1030          pd.Relay6 = Max(0, Val(rs("Relay6") & ""))
1040          pd.Relay7 = Max(0, Val(rs("Relay7") & ""))
1050          pd.Relay8 = Max(0, Val(rs("Relay8") & ""))

1060          pd.OpenConnection
1070          pd.Checked = True
1080        End If
1090        rs.Close

1100    End Select

ChangePageDevice_Resume:
1110    Set rs = Nothing
1120    On Error GoTo 0
1130    Exit Function

ChangePageDevice_Error:

1140    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.ChangePageDevice." & Erl
1150    Resume ChangePageDevice_Resume


End Function

Function InitPageDevices() As Long
        Dim rs                 As Recordset
        Dim j                  As Integer
        Dim pd                 As cPageDevice
        Dim Found              As Boolean
        Dim protocol           As Long
        Dim NoSubSystem        As Boolean
        ' we need to see if it exists
        ' if it doesn't then
        ' add it

10      On Error GoTo InitPageDevices_Error
        ' dbg "InitPageDevices Init Dialogic System 10"

        ' dbg "InitPageDevices Init Dialogic System Done 20"
        ' Set DialogicSystem = New DialogicSystem ' moved to where protocol is looked at

20      Set rs = ConnExecute("Select * from PagerDevices")
30      Do Until rs.EOF
40        protocol = Val(rs("protocolid") & "")
50        dbg "protocol " & protocol

60        On Error Resume Next  ' added 2/13/18

70        NoSubSystem = False
80        If protocol = PROTOCOL_DIALOGIC Then
90          If DialogicSystem Is Nothing Then
100           Set DialogicSystem = New DialogicSystem
110         End If
120         NoSubSystem = NODIVA Or (DialogicSystem.TotalChannels = 0)
130         If NoSubSystem Then
140           SpecialLog "NO Diva Subsystem"
150         End If
160         If Err.Number Then
170           If InIDE Then
180             MsgBox "Error " & Err.Number & " (" & Err.Description & ") at modPaging.InitPageDevices." & Erl
190             NoSubSystem = True
200             Err.Clear
              End If
210         End If
220       End If


230       On Error GoTo InitPageDevices_Error

240       If NoSubSystem Then
            ' skip it, no dialogic
250         NoSubSystem = NoSubSystem


260       Else
270         Found = False
280         For j = gPageDevices.Count To 1 Step -1
290           Set pd = gPageDevices(j)
300           If pd.DeviceID = rs("ID") Then
310             Found = True         ' marked as still in the system
320             Exit For
330           End If
340         Next
350         If Not Found Then
360           Set pd = New cPageDevice
370           pd.DeviceID = rs("ID")
380           gPageDevices.Add pd
390         End If
400         pd.ProtocolID = Val(rs("protocolid") & "")
410         pd.AudioDevice = rs("AudioDevice") & ""
420         pd.BaudRate = rs("baudrate") & ""
430         pd.BITS = rs("bits") & ""
440         pd.Description = rs("Description") & ""
450         pd.Parity = rs("parity") & ""
460         pd.Port = Val(rs("Port") & "")
            If pd.Port = 8 Or pd.Port = 9 Then
             ' Debug.Assert 0
            End If

470         pd.Settings = rs("settings") & ""
480         pd.Pause = rs("Pause")
490         pd.PASystemKey = IIf(rs("keypa") = 1, 1, 0)
500         pd.PARepeatTwice = IIf(rs("twice") = 1, 1, 0)
            'If pd.ProtocolID = PROTOCOL_REMOTE Then
            '  pd.PARepeatTwice = 0

510         pd.LFeeds = Val(rs("lf") & "")
            ' MARQUIS
520         pd.MarquisControlCode = Max(0, Val(rs("MarquisCode") & ""))


            ' ONTRAK RELAY
530         pd.Relay1 = Max(0, Val(rs("Relay1") & ""))
540         pd.Relay2 = Max(0, Val(rs("Relay2") & ""))
550         pd.Relay3 = Max(0, Val(rs("Relay3") & ""))
560         pd.Relay4 = Max(0, Val(rs("Relay4") & ""))
570         pd.Relay5 = Max(0, Val(rs("Relay5") & ""))
580         pd.Relay6 = Max(0, Val(rs("Relay6") & ""))
590         pd.Relay7 = Max(0, Val(rs("Relay7") & ""))
600         pd.Relay8 = Max(0, Val(rs("Relay8") & ""))


            ' DIALER  &      'DIALOGIC
            Dim Channel        As Long

610         Channel = rs("DialerModem")
620         If pd.ProtocolID = PROTOCOL_DIALOGIC Then
630           If pd.DialerModem <> Channel Then  ' unreserve old channel
640             DialogicSystem.Reserved(pd.DialerModem) = False
650           End If
660           DialogicSystem.Reserved(Channel) = True  ' reserve channel

670         End If

680         pd.DialerModem = rs("DialerModem")  ' long, unique device integer
690         pd.DialerMsgDelay = rs("DialerMsgDelay")
700         pd.DialerMsgRepeats = rs("DialerMsgRepeats")
710         pd.DialerMsgSpacing = rs("DialerMsgSpacing")
720         pd.DialerPhone = rs("DialerPhone") & ""
730         pd.DialerRedialDelay = rs("DialerRedialDelay")
740         pd.DialerRedials = rs("DialerRedials")
750         pd.DialerTag = rs("DialerTag") & ""
760         pd.DialerVoice = rs("DialerVoice") & ""
770         pd.DialerTerminateDigit = Val(rs("DialerAckDigit") & "")
            pd.KeepPaging = Val(rs("KeepPaging") & "") And 1 ' fix 2018-11-14

780         If MASTER Then
790           pd.OpenConnection
800         End If
810         pd.Checked = True
820       End If
830       rs.MoveNext
840     Loop
850     rs.Close
860     Set rs = Nothing

870     For j = gPageDevices.Count To 1 Step -1
880       Set pd = gPageDevices(j)
890       If pd.Checked = False Then
900         If MASTER Then
910           pd.CloseConnection
920         End If
930         gPageDevices.Remove j
940       End If
950     Next

960     For j = gPageDevices.Count To 1 Step -1
970       Set pd = gPageDevices(j)
980       pd.Checked = False
990     Next

InitPageDevices_Resume:
1000    On Error GoTo 0
1010    Exit Function

InitPageDevices_Error:

1020    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPaging.InitPageDevices." & Erl
1030    Resume InitPageDevices_Resume


End Function




Public Function MarquiCode2MarquiChar(ByVal MarquisCode As Long) As String
' returns a one-char code from the list below

'  Const none = 0
'  Const Normal = 1
'  Const emergency = 2
'  Const help = 3
'  Const help_lav = 4
'  Const info = 5
'  Const Apollo = 6

'Public Const Marquis_NONE = 0
'Public Const Marquis_NORMAL = 1
'Public Const Marquis_EMERGENCY = 2
'Public Const Marquis_HELP = 3
'Public Const Marquis_HELP_LAV = 4
'Public Const Marquis_INFO = 5
'Public Const Marquis_APOLLO = 6


  Select Case MarquisCode
    Case MARQUIS_NONE
      MarquiCode2MarquiChar = "*"
    Case MARQUIS_NORMAL
      MarquiCode2MarquiChar = "$"
    Case MARQUIS_EMERGENCY
      MarquiCode2MarquiChar = "+"
    Case MARQUIS_HELP
      MarquiCode2MarquiChar = "="
    Case MARQUIS_HELP_LAV
      MarquiCode2MarquiChar = "#"
    Case MARQUIS_INFO
      MarquiCode2MarquiChar = "*"
    Case MARQUIS_APOLLO
      MarquiCode2MarquiChar = ""
    Case Else
      MarquiCode2MarquiChar = "*"
      
  End Select
End Function

'Public Sub SendClearMarquis(pageitem As cPageItem)
'
'  SendToGroup pageitem.Message, GroupID, "", "", CLEAR_MARQUIS
'
'End Sub

'Public Sub SendClearRelay(pageitem As cPageItem)
'  SendToGroup pageitem.Message, GroupID, "", "", CLEAR_RELAY
'End Sub


'Public Sub SendCancelToGroup()
'  SendCancelToGroup pageitem.Message & " Cancel", GroupID, "", ""
'End Sub
