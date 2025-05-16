VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPageItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' new 7/10/1018
' for all alarms
Public OutputGroupID      As Long
Public OutputGroupName    As String

Public AssistRequest       As Long


' These are mobile-only specific
Public NamedGroup         As String
Public NamedGroupID       As String


Public ActiveGroup       As cPagerGroup
Private mCurrentShift    As Integer

' flow control
Public starttime         As Date  ' when this was created
Public EscalateTime      As Date  ' what time we need to escalate
Public EscalationLevel   As Long
Public Escalated         As Boolean
Public LastPageTime      As Date  ' when last page was sent
'Public NextPageTime     As Date ' when we can send again (LastPageTime +  pause)

Public NoWait            As Long
Public SentOnce          As Boolean

Public Announce          As String
' we'll add this to the collection of pageitems
' call to it every second from main timing loop

' Has the repeat every n seconds elapsed ?
' if so then
' walk each outputdevice and see if it
' a) is ready
' b) has more pages to go
' c) other special handling

' for use/set in cPageDevice
Public SendComplete      As Long
Public TimeAdded         As Date
Public TimeSent          As Date


' For ONTRACK/ADU200


Private mRelayID         As Long  ' not sure how this is appropriate here, other than a pass thru


' Transmitter settings
Public Serial            As String
Public inputnum          As Long
Public EventType         As Long
Private mAlarmID         As Long  ' to track this thru the process




Public locationtext      As String
Public Phone             As String



' as requested from transmitter
' output mask from Troubles and Batts also set these
Public Repeats           As Long  ' how many repeats are asked for
Public Pause As Long           ' seconds between pages 'Public RepeatEvery            As Long  ' seconds between pages
Public RepeatUntil       As Long  ' repeat until cancel
Private mSendCancel      As Boolean  ' to send cancel or not
Public SendCount         As Long  ' how many times sent

Public EscSendCount      As Long  ' how many times sent

'Public Protocol         As Long

' OUTPUT (pager)

Private mAddress         As String  ' pager address
Public NoRepeats         As Integer  ' overrides repeats on a pager by pager

' not used here ?
'Public Sendphone        As Boolean  ' whether or not to send resident Phone #
'Public DeviceID         As Long     ' output server/protocol
'Public PIN              As String
'Public DefaultMessage   As String
'Public MarquisCode      As String



Public IsCancel          As Boolean

' management
Private mGroupList       As Collection


Private mResidentText    As String

' at the server level ?
Public Preamble          As String
Private mRepeatTwice     As Integer  ' flag to send page twice to PCA, stored in outputdevice's record


Private mRoomText        As String
Private mCancelText      As String


Private mOG1             As Long
Private mOG2             As Long
Private mOG3             As Long
Private mOG4             As Long
Private mOG5             As Long
Private mOG6             As Long

Private mOG1D            As Long
Private mOG2D            As Long
Private mOG3D            As Long
Private mOG4D            As Long
Private mOG5D            As Long
Private mOG6D            As Long


Private mNG1             As Long
Private mNG2             As Long
Private mNG3             As Long
Private mNG4             As Long
Private mNG5             As Long
Private mNG6             As Long


Private mNG1D            As Long
Private mNG2D            As Long
Private mNG3D            As Long
Private mNG4D            As Long
Private mNG5D            As Long
Private mNG6D            As Long

Private mGG1             As Long
Private mGG2             As Long
Private mGG3             As Long
Private mGG4             As Long
Private mGG5             As Long
Private mGG6             As Long

Private mGG1D            As Long
Private mGG2D            As Long
Private mGG3D            As Long
Private mGG4D            As Long
Private mGG5D            As Long
Private mGG6D            As Long

Private mGroup0          As cPagerGroup
Private mGroup1          As cPagerGroup
Private mGroup2          As cPagerGroup
Private mGroup3          As cPagerGroup
Private mGroup4          As cPagerGroup
Private mGroup5          As cPagerGroup

Private mMessage         As String
'Private mMarquisMsgs As Collection
Private mNoNameMessage   As String
Private InitDone         As Boolean



Public Function ProcessDialerACK()
    
    Dim pager As cPager
    
    For Each pager In ActiveGroup.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        ' we want to stop all other pages
        If pager.ProtocolID = PROTOCOL_DIALER Then
          pager.NoRepeats = 1
        End If
        
      End If
    Next
    
    Exit Function
    
    For Each pager In mGroup0.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        pager.NoRepeats = 1
        dbg "setting pager.NoRepeats = 1 "
      End If
    Next
    
    For Each pager In mGroup1.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        pager.NoRepeats = 1
        dbg "setting pager.NoRepeats = 1 "
      End If
    
    Next

    For Each pager In mGroup2.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        pager.NoRepeats = 1
        dbg "setting pager.NoRepeats = 1 "
      End If
    
    Next
    
    For Each pager In mGroup3.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        pager.NoRepeats = 1
        dbg "setting pager.NoRepeats = 1 "
      End If
    
    Next

    For Each pager In mGroup4.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        pager.NoRepeats = 1
        dbg "setting pager.NoRepeats = 1 "
      End If
    
    Next

    For Each pager In mGroup5.Pagers
      If pager.ProtocolID = PROTOCOL_DIALER Then
        pager.NoRepeats = 1
        dbg "setting pager.NoRepeats = 1 "
      End If
    
    Next



End Function




Public Function Init()
  On Error Resume Next
  
  Set mGroup0.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
  Set ActiveGroup = mGroup0
  ActiveGroup.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
  
End Function

Public Function GetGroupDuration(ByVal EscalationLevel As Integer, ByVal Shift As Integer) As Long
  ' returns minutes of duration of

  Dim delay As Long
  
  On Error Resume Next
  
  
  Shift = Shift
  Select Case Shift
    Case SHIFT_NIGHT
      Select Case EscalationLevel
        Case 0
          delay = NG1D
        Case 1
          delay = NG2D
        Case 2
          delay = NG3D
        Case 3
          delay = NG4D
        Case 4
          delay = NG5D
        Case 5
          delay = NG6D
      End Select
      
    Case SHIFT_GRAVE
      Select Case EscalationLevel
        Case 0
          delay = GG1D
        Case 1
          delay = GG2D
        Case 2
          delay = GG3D
        Case 3
          delay = GG4D
        Case 4
          delay = GG5D
        Case 5
          delay = GG6D
      End Select
      
      
    Case Else ' SHIFT_DAY
      Select Case EscalationLevel
        Case 0
          delay = OG1D
        Case 1
          delay = OG2D
        Case 2
          delay = OG3D
        Case 3
          delay = OG4D
        Case 4
          delay = OG5D
        Case 5
          delay = OG6D
      End Select

  End Select
  If delay <= 0 Then
    delay = 3
  End If
  GetGroupDuration = delay

End Function

Public Function Send() As Long
        Dim pager         As cPager
        Dim NextGroupID   As Long
        ' SendToPager adds a pageitem to output server's que
10      On Error GoTo Send_Error

'20      On Error GoTo 0


30      If Ready Then  ' ready is true if sendcount = 0 or ALARM intermessage delay has expired
40        For Each pager In ActiveGroup.Pagers
50          If pager.SendCount = 0 Then
60            If pager.Twice Then
70              pager.Twice = False  ' nowait = 1
80              SendToPager pager.message, pager.PagerID, 1, pager.Phone, pager.RoomText, PAGER_NORMAL, pager.MarquisMessage, AlarmID, pager.inputnum
90            End If
100           SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_NORMAL, pager.MarquisMessage, AlarmID, pager.inputnum
110           pager.SendCount = pager.SendCount + 1
              'SendCount = SendCount + 1  ' how many times have we sent this
120         Else
130           If NoRepeats Then
                ' done
140           ElseIf RepeatUntil Then
                'dbg "Repeat until cleared  ***********"
150             If pager.NoRepeats Then
                  ' dbg "Device CANNOT Repeat ***********"
160             Else
170               SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_NORMAL, pager.MarquisMessage, AlarmID, pager.inputnum
180               pager.SendCount = pager.SendCount + 1
190             End If
                'SendCount = Min(1000000, SendCount + 1) ' how many times have we sent this
200           ElseIf pager.SendCount > Repeats Then
                'dbg "ALL REPEATS DONE"
                'EscalationLevel = 1
                ' don't send
210           Else
220             If pager.NoRepeats Then
                  'dbg "Device CANNOT Repeat ***********"
230             Else
240               SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_NORMAL, pager.MarquisMessage, AlarmID, pager.inputnum
250               pager.SendCount = pager.SendCount + 1
260             End If
270           End If
280         End If

290       Next  'pager In activegroup
300       SendCount = Min(1000000, SendCount + 1)  ' how many times have we sent this total
310     End If  'if ready

320     If ActiveGroup.Expired Then
330       NextGroupID = GetPagerGroupID(EscalationLevel + 1, CurrentShift)
340       If NextGroupID <> 0 Then  ' if it's zero then we're done
            'escalate
350         If EscalationLevel < 6 Then
360           EscalationLevel = EscalationLevel + 1
370           dbgTAPI "cPageItem.Send Escalating " & EscalationLevel
380         Else
390           Exit Function  ' Done, no more levels
400         End If
410         Select Case EscalationLevel
              Case 0
420             Set mGroup0.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
430             mGroup0.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
440             Set ActiveGroup = mGroup0
               
450           Case 1
460             Set mGroup1.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
470             mGroup1.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
480             Set ActiveGroup = mGroup1
490             SendCount = 0
500           Case 2
510             Set mGroup2.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
520             mGroup2.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
530             Set ActiveGroup = mGroup2
540             SendCount = 0
550           Case 3
560             Set mGroup3.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
570             mGroup3.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
580             Set ActiveGroup = mGroup3
590             SendCount = 0
600           Case 4
610             Set mGroup4.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
620             mGroup4.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
630             Set ActiveGroup = mGroup4
640             SendCount = 0
650           Case 5
660             Set mGroup5.Pagers = GetEscalationGroupPagers(EscalationLevel, CurrentShift)
670             mGroup5.Expires = DateAdd("n", GetGroupDuration(EscalationLevel, CurrentShift), Now)
680             Set ActiveGroup = mGroup5
690             SendCount = 0
700           Case Else
710             Exit Function
720         End Select
730       Else
740         Exit Function  ' Done, no levels available
750       End If
760     End If


Send_Resume:

770     On Error GoTo 0
780     Exit Function

Send_Error:

790     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageItem.Send." & Erl
800     Resume Send_Resume

End Function

Function GetPagerGroupID(ByVal LEvel As Integer, ByVal Shift As Integer) As Long
  Dim ID As Long
  On Error Resume Next
  
  Select Case Shift
    Case SHIFT_NIGHT
      Select Case LEvel
        Case 0
          ID = NG1
        Case 1
          ID = NG2
        Case 2
          ID = NG3
        Case 3
          ID = NG4
        Case 4
          ID = NG5
        Case 5
          ID = NG6
        Case Else
          ID = 0
      End Select
    Case SHIFT_GRAVE
      Select Case LEvel
        Case 0
          ID = GG1
        Case 1
          ID = GG2
        Case 2
          ID = GG3
        Case 3
          ID = GG4
        Case 4
          ID = GG5
        Case 5
          ID = GG6
        Case Else
          ID = 0
      End Select
    
    
    Case Else  ' SHIFT_DAY
      Select Case LEvel
        Case 0
          ID = OG1
        Case 1
          ID = OG2
        Case 2
          ID = OG3
        Case 3
          ID = OG4
        Case 4
          ID = OG5
        Case 5
          ID = OG6
        Case Else
          ID = 0
          
      End Select

  End Select
  GetPagerGroupID = ID

End Function



Public Sub SendClearMarquis(ByVal LEvel As Integer)
' called by class_terminate
' once for each escalation level

  Dim Group As cPagerGroup
  Dim pager As cPager

  Select Case LEvel

    Case 5
      Set Group = mGroup5
    Case 4
      Set Group = mGroup4
    Case 3
      Set Group = mGroup3
    Case 2
      Set Group = mGroup2
    Case 1
      Set Group = mGroup1
    Case Else  ' case 0
      Set Group = mGroup0
  End Select

  If Group Is Nothing Then Exit Sub


  For Each pager In Group.Pagers
    If pager.IsMarquis Then
      SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_CLEAR_MARQUIS, pager.MarquisMessage, AlarmID, pager.inputnum

    End If
  Next


End Sub
Public Sub SendClearApollo(ByVal LEvel As Integer)
  Dim Group As cPagerGroup
  Dim pager As cPager

  Select Case LEvel

    Case 5
      Set Group = mGroup5
    Case 4
      Set Group = mGroup4
    Case 3
      Set Group = mGroup3
    Case 2
      Set Group = mGroup2
    Case 1
      Set Group = mGroup1
    Case Else  ' case 0
      Set Group = mGroup0
  End Select

  If Group Is Nothing Then Exit Sub


  For Each pager In Group.Pagers
    If pager.IsApollo Then
      SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_CLEAR_APOLLO, pager.MarquisMessage, AlarmID, pager.inputnum
    End If
  Next

End Sub
Public Sub SendClearRemotes(ByVal LEvel As Integer)
  modEvents.ClearRemotes AlarmID
End Sub

Public Sub SendClearRelay(ByVal LEvel As Integer)
  ' called by class_terminate

  Dim Group As cPagerGroup
  Dim pager As cPager
  
'  Select Case Level
'    Case 1
'      Set Group = mGroup1
'    Case Else
'      Set Group = mGroup0
'  End Select
  

  Select Case LEvel
    Case 0
      Set Group = mGroup0
    Case 1
      Set Group = mGroup1
    Case 2
      Set Group = mGroup2
    Case 3
      Set Group = mGroup3
    Case 4
      Set Group = mGroup4
    Case 5
      Set Group = mGroup5
    Case Else
      Exit Sub
  End Select
  
  If Group Is Nothing Then Exit Sub
  

  For Each pager In Group.Pagers
    If pager.ProtocolID = PROTOCOL_ONTRAK Then
      SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_CANCEL, pager.MarquisMessage, AlarmID, pager.inputnum
    End If
  Next


End Sub

Public Sub SendCancels(ByVal LEvel As Integer)

       ' called by class_terminate

        Dim Group As cPagerGroup  ' Collection
        Dim pager As cPager
        
10      On Error GoTo SendCancels_Error

20        Select Case LEvel
          Case 0
30          Set Group = mGroup0
40        Case 1
50          Set Group = mGroup1
60        Case 2
70          Set Group = mGroup2
80        Case 3
90          Set Group = mGroup3
100       Case 4
110         Set Group = mGroup4
120       Case 5
130         Set Group = mGroup5
140       Case Else
150         Exit Sub
160     End Select
      '
      '  Set Group = ActiveGroup '  cPagerGroup
        
170     If Group Is Nothing Then Exit Sub
        
        
180     For Each pager In Group.Pagers
190       If pager.IsMarquis Or pager.NoCancel Or pager.IsApollo Then ' marquis and apollo are handled when exiting pageitem
            'dbg "pager does not support cancel"
200       ElseIf pager.ProtocolID = PROTOCOL_REMOTE Then
              'dbg "pager does not support cancel"
210       Else
          
220         SendToPager pager.message, pager.PagerID, 0, pager.Phone, pager.RoomText, PAGER_CANCEL, pager.MarquisMessage, AlarmID, pager.inputnum
230       End If
240     Next

SendCancels_Resume:

250     On Error GoTo 0
260     Exit Sub

SendCancels_Error:

270     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageItem.SendCancels." & Erl
280     Resume SendCancels_Resume

End Sub



Function GetEscalationGroupPagers(ByVal LEvel As Integer, ByVal Shift As Integer) As Collection    ' 0 to 5
  
  
  
  Dim Group As Collection
  Dim GroupID As Long
  Dim rsGroup As ADODB.Recordset
  Dim rsPager As ADODB.Recordset
  Dim pager   As cPager
  Dim PagerID As Long
  Dim SQL     As String

  '----------------- gets pagers for group
  '----------------- do for selected escalation group
  '----------------- we only get these as needed


  On Error Resume Next
  

  Set Group = New Collection
  
  Select Case Shift
  
      Case SHIFT_GRAVE
      Select Case LEvel
        Case 5
          GroupID = mGG6
        Case 4
          GroupID = mGG5
        Case 3
          GroupID = mGG4
        Case 2
          GroupID = mGG3
        Case 1
          GroupID = mGG2
        Case Else  ' 0
          GroupID = mGG1
      End Select
  
  
    Case SHIFT_NIGHT
      Select Case LEvel
        Case 5
          GroupID = mNG6
        Case 4
          GroupID = mNG5
        Case 3
          GroupID = mOG4
        Case 2
          GroupID = mNG3
        Case 1
          GroupID = mNG2
        Case Else  ' 0
          GroupID = mNG1
      End Select

    
    
    Case Else  ' day shift
      Select Case LEvel
        Case 5
          GroupID = mOG6
        Case 4
          GroupID = mOG5
        Case 3
          GroupID = mOG4
        Case 2
          GroupID = mOG3
        Case 1
          GroupID = mOG2
        Case Else  ' 0
          GroupID = mOG1
      End Select
  End Select

  If GroupID <> 0 Then
    Set rsGroup = ConnExecute("SELECT * FROM GroupPager WHERE groupID = " & GroupID)
    Do Until rsGroup.EOF

      PagerID = rsGroup("pagerid")
      SQL = "SELECT Pagers.*, PagerDevices.KeyPA, PagerDevices.Twice, PagerDevices.ProtocolID FROM Pagers INNER JOIN PagerDevices ON Pagers.DeviceID = PagerDevices.ID   WHERE  pagerid = " & PagerID

      Set rsPager = ConnExecute(SQL)
      If Not rsPager.EOF Then
        Set pager = New cPager
        pager.PagerID = PagerID
        pager.AlarmID = AlarmID  ' class level ID for tracking thru system
        pager.Phone = Phone
        pager.RoomText = RoomText
        
        pager.MarquisCode = Val(rsPager("MarquisCode") & "")
        pager.MarquisMessage = MarquisMessage(pager.MarquisCode)
        pager.DefaultMessage = rsPager("DefaultMessage") & ""
        pager.DeviceID = rsPager("deviceID")
        pager.Description = rsPager("Description") & ""
        pager.Address = rsPager("identifier") & ""
        pager.Identifier = rsPager("identifier") & ""
        pager.Twice = rsPager("twice")
        pager.KeyPA = rsPager("KeyPA")
        
        pager.NoRepeats = IIf(rsPager("NoRepeats") = 1, 1, 0)
        pager.NoName = IIf(rsPager("noname") = 1, 1, 0)
        If pager.NoName = 1 Then
          pager.message = NoNameMessage
        Else
          pager.message = message
        End If
        
        pager.Sendphone = IIf(rsPager("IncludePhone") = 1, 1, 0)
        pager.PIN = rsPager("pin") & ""
        pager.ProtocolID = rsPager("ProtocolID")
        pager.NoCancel = IIf(pager.ProtocolID = PROTOCOL_CENTRAL, 1, 0)
        
        
        pager.RelayNum = Val(rsPager("RelayNum") & "")

        Group.Add pager
      End If
      rsPager.Close
      rsGroup.MoveNext
    Loop  ' next pager
    rsGroup.Close
  End If



  Set rsPager = Nothing
  Set rsGroup = Nothing
  Set GetEscalationGroupPagers = Group

End Function

Private Sub Class_Terminate()

        Dim j             As Integer
        'Send cancels, clear marquis, and reset relays for each escalation level

10      If EscalationLevel > 0 Then
20        RemoveDupeCancels
30      End If

40      For j = EscalationLevel To 0 Step -1
50        If SendCancel Then
60          SendCancels j
70        End If
80        SendClearMarquis j
          SendClearApollo j
90        SendClearRelay j
          

100     Next

110     Set mGroup5 = Nothing
120     Set mGroup4 = Nothing
130     Set mGroup3 = Nothing
140     Set mGroup2 = Nothing
150     Set mGroup1 = Nothing
160     Set mGroup0 = Nothing

170     Set mGroupList = Nothing
'180     Set mMarquisMsgs = Nothing
End Sub

Sub RemoveDupeCancels()
        Dim Group              As cPagerGroup  ' Collection
        Dim GroupPager         As cPager


        Dim pager              As cPager
        Dim Pagers             As Collection
        Dim samepagers         As Collection
        Dim j                  As Long
        Dim p                  As Long

        ' theory:
        ' start with last escalation group
        ' if it's not a pager have in the list, add it to the list
        ' if we have it in the list, remove it from the group

10      On Error GoTo RemoveDupeCancels_Error

20      If EscalationLevel = 0 Then Exit Sub


30      Set Pagers = New Collection
40      Set samepagers = New Collection

50      For j = EscalationLevel To 0 Step -1
60        Select Case j
            Case 0
70            Set Group = mGroup0
80          Case 1
90            Set Group = mGroup1
100         Case 2
110           Set Group = mGroup2
120         Case 3
130           Set Group = mGroup3
140         Case 4
150           Set Group = mGroup4
160         Case 5
170           Set Group = mGroup5
180         Case Else
              'nada
190       End Select

          Dim sp               As Long
          Dim PagerID          As String
          Dim genobj           As cGenericObject
          Dim tempgenobj       As cGenericObject
200       For p = Group.Pagers.Count To 1 Step -1
210         Set GroupPager = Group.Pagers(p)
220         For sp = 1 To samepagers.Count

230           Set tempgenobj = samepagers(sp)
240           If (tempgenobj.IntVal = GroupPager.PagerID) Then
250             Group.Pagers.Remove p
260             Exit For
270           End If

              'End If
280         Next
290         If sp > samepagers.Count Then
300           Set genobj = New cGenericObject
310           genobj.IntVal = GroupPager.PagerID
320           samepagers.Add genobj
330         End If
340       Next

350     Next


RemoveDupeCancels_Resume:

360     On Error GoTo 0
370     Exit Sub

RemoveDupeCancels_Error:

380     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cPageItem.RemoveDupeCancels." & Erl
390     Resume RemoveDupeCancels_Resume

End Sub

Public Property Let message(ByVal Value As String)
  mMessage = Value
End Property
Public Property Let NoNameMessage(ByVal Value As String)
  mNoNameMessage = Value
End Property

Public Property Get NoNameMessage() As String
  
  Dim LocationPhrase As String
  On Error Resume Next
  
  If mNoNameMessage = "" Then
    Preamble = Trim(Preamble)
    Announce = Trim(Announce)
    RoomText = Trim(RoomText)
    ResidentText = Trim(ResidentText)
    locationtext = Trim(locationtext)
    
    
    If (Configuration.locationtext = 1) And (Len(locationtext) > 0) Then
        LocationPhrase = "Location: "
    End If
    
    Dim Temp As String
    Temp = Trim$(IIf(Len(Preamble) > 0, Preamble & " ", " ") & Announce & " " & IIf(Len(RoomText) > 0, "Room: " & RoomText, " ") & " " & LocationPhrase & locationtext)
    Do While InStr(Temp, "  ") > 0
      Temp = Replace(Temp, "  ", " ")
    Loop
    NoNameMessage = Temp
  Else
    NoNameMessage = mNoNameMessage
  End If
End Property


Public Property Get message() As String
 
  Dim LocationPhrase As String
 
  If mMessage = "" Then
    Preamble = Trim(Preamble)
    Announce = Trim(Announce)
    RoomText = Trim(RoomText)
    ResidentText = Trim(ResidentText)
    locationtext = Trim(locationtext)
    
    If (Configuration.locationtext = 1) And (Len(locationtext) > 0) Then
        LocationPhrase = "Location: "
    End If
    
    message = Trim$(IIf(Len(Preamble) > 0, Preamble & " ", " ") & Announce & " " & IIf(Len(RoomText) > 0, "Room: " & RoomText, " ") & " " & ResidentText & " " & LocationPhrase & locationtext)
  Else
    message = mMessage
  End If
End Property

Public Function MarquisMessage(ByVal MarquisCode As Long) As String

    Dim s As String
    If (Len(RoomText) > 0) Then
      s = Announce & " " & RoomText
    ElseIf (Len(ResidentText) > 0) Then
      s = Announce & " " & ResidentText
    Else
      s = Announce
    End If
    

    s = Replace(s, "*", " ")
    s = Replace(s, "$", " ")
    s = Replace(s, "+", " ")
    s = Replace(s, "=", " ")
    s = Replace(s, "#", " ")
    
    Do While InStr(s, "  ")
      s = Replace(s, "  ", " ")
    Loop
    

    If MarquisCode = MARQUIS_APOLLO Then
      s = Trim$(s)
    Else
      s = left$(Trim$(s), 19)
    End If
    
    MarquisMessage = s

    
End Function

Public Function Ready() As Boolean
  Dim IsReady As Boolean
  
  On Error Resume Next
  
  If SendCount = 0 Then '  Then ' send it right away
    'dbg "cpageitem.ready (NOT sent yet)"
    LastPageTime = Now
    IsReady = True
  Else
    IsReady = (DateDiff("s", Now, DateAdd("s", Pause, LastPageTime)) <= 0)
    'Debug.Print "cPageitem.ready Next Send in " & DateDiff("s", Now, DateAdd("s", Pause, LastPageTime))
    If IsReady Then
      LastPageTime = Now
    End If
  End If
  'Debug.Print "cPagItem.Ready "; IsReady
  Ready = IsReady
End Function

Public Property Get OG1() As Long
  OG1 = mOG1
End Property

Public Property Let OG1(ByVal OG1 As Long)
  mOG1 = OG1
End Property

Public Property Get OG2() As Long
  OG2 = mOG2
End Property

Public Property Let OG2(ByVal OG2 As Long)
  mOG2 = OG2
End Property

Public Property Get OG3() As Long
  OG3 = mOG3
End Property
Public Property Let OG3(ByVal OG3 As Long)
  mOG3 = OG3
End Property

Public Property Get NG1() As Long
  NG1 = mNG1
End Property

Public Property Let NG1(ByVal NG1 As Long)
  mNG1 = NG1
End Property

Public Property Get NG2() As Long
  NG2 = mNG2
End Property

Public Property Let NG2(ByVal NG2 As Long)
  mNG2 = NG2
End Property

Public Property Get NG3() As Long
  NG3 = mNG3
End Property

Public Property Let NG3(ByVal NG3 As Long)
  mNG3 = NG3
End Property

Private Sub Class_Initialize()
    
  
  Set mGroup0 = New cPagerGroup
  Set mGroup1 = New cPagerGroup
  Set mGroup2 = New cPagerGroup
  Set mGroup3 = New cPagerGroup
  Set mGroup4 = New cPagerGroup
  Set mGroup5 = New cPagerGroup
  
  CurrentShift = GetCurrentShift() ' set once for the event.... the current shift owns this
  
  starttime = Now
  EscalateTime = DateAdd("s", Configuration.EscTimer, starttime)
  
  'Set mMarquisMsgs = New Collection
  Set mGroupList = New Collection
  
  mRepeatTwice = 1
End Sub

Public Property Get CancelText() As String
  CancelText = mCancelText
End Property

Public Property Let CancelText(ByVal CancelText As String)
  mCancelText = CancelText
End Property

Public Property Get RepeatTwice() As Integer
  'read only
  RepeatTwice = mRepeatTwice
End Property

Public Property Let RepeatTwice(ByVal Value As Integer)
  mRepeatTwice = Value
End Property

Public Property Get RoomText() As String
  RoomText = mRoomText
End Property

Public Property Let RoomText(ByVal Value As String)
 ' If value = "" Then Stop
  mRoomText = Value
End Property


'Public Property Get MarquisMsgs() As Collection
'  Set MarquisMsgs = mMarquisMsgs
'End Property
'
'Public Property Set MarquisMsgs(MarquisMsgs As Collection)
'  Set mMarquisMsgs = MarquisMsgs
'End Property

Public Property Get GroupList() As Collection
  Set GroupList = mGroupList
End Property

Public Property Set GroupList(GroupList As Collection)
  Set mGroupList = GroupList
End Property

Public Property Get Group3() As cPagerGroup
  Set Group3 = mGroup3
End Property

Public Property Set Group3(Group3 As cPagerGroup)
  Set mGroup3 = Group3
End Property

Public Property Get Group4() As cPagerGroup
  Set Group4 = mGroup4
End Property

Public Property Set Group4(Group4 As cPagerGroup)
  Set mGroup4 = Group4
End Property

Public Property Get Group5() As cPagerGroup
  Set Group5 = mGroup5
End Property

Public Property Set Group5(Group5 As cPagerGroup)
  Set mGroup5 = Group5
End Property

'Public Property Get DayGroup3() As Collection
'
'  Set DayGroup3 = mDayGroup3
'
'End Property
'
'Public Property Set DayGroup3(DayGroup3 As Collection)
'
'  Set mDayGroup3 = DayGroup3
'
'End Property
'
'Public Property Get NiteGroup3() As Collection
'
'  Set NiteGroup3 = mNiteGroup3
'
'End Property
'
'Public Property Set NiteGroup3(NiteGroup3 As Collection)
'
'  Set mNiteGroup3 = NiteGroup3
'
'End Property

'Public Property Get DayGroup4() As Collection
'
'  Set DayGroup4 = mDayGroup4
'
'End Property
'
'Public Property Set DayGroup4(DayGroup4 As Collection)
'
'  Set mDayGroup4 = DayGroup4
'
'End Property
'
'Public Property Get NiteGroup4() As Collection
'
'  Set NiteGroup4 = mNiteGroup4
'
'End Property
'
'Public Property Set NiteGroup4(NiteGroup4 As Collection)
'
'  Set mNiteGroup4 = NiteGroup4
'
'End Property
'
'Public Property Get DayGroup5() As Collection
'
'  Set DayGroup5 = mDayGroup5
'
'End Property
'
'Public Property Set DayGroup5(DayGroup5 As Collection)
'
'  Set mDayGroup5 = DayGroup5
'
'End Property
'
'Public Property Get NiteGroup5() As Collection
'
'  Set NiteGroup5 = mNiteGroup5
'
'End Property

'Public Property Set NiteGroup5(NiteGroup5 As Collection)
'
'  Set mNiteGroup5 = NiteGroup5
'
'End Property

Public Property Get OG4() As Long

  OG4 = mOG4

End Property

Public Property Let OG4(ByVal OG4 As Long)

  mOG4 = OG4

End Property

Public Property Get OG5() As Long

  OG5 = mOG5

End Property

Public Property Let OG5(ByVal OG5 As Long)

  mOG5 = OG5

End Property

Public Property Get OG6() As Long

  OG6 = mOG6

End Property

Public Property Let OG6(ByVal OG6 As Long)

  mOG6 = OG6

End Property

Public Property Get NG4() As Long

  NG4 = mNG4

End Property

Public Property Let NG4(ByVal NG4 As Long)

  mNG4 = NG4

End Property

Public Property Get NG5() As Long

  NG5 = mNG5

End Property

Public Property Let NG5(ByVal NG5 As Long)

  mNG5 = NG5

End Property

Public Property Get NG6() As Long

  NG6 = mNG6

End Property

Public Property Let NG6(ByVal NG6 As Long)

  mNG6 = NG6

End Property

Public Property Get RelayID() As Long
  RelayID = mRelayID
End Property

Public Property Let RelayID(ByVal RelayID As Long)
  mRelayID = RelayID
End Property

Public Property Get AlarmID() As Long
  AlarmID = mAlarmID
End Property

Public Property Let AlarmID(ByVal AlarmID As Long)
  mAlarmID = AlarmID
End Property

Public Property Get Address() As String
  Address = mAddress
End Property

Public Property Let Address(ByVal Value As String)
  mAddress = Value
End Property

Public Property Get ResidentText() As String
  ResidentText = mResidentText
End Property

Public Property Let ResidentText(ByVal ResidentText As String)
  mResidentText = ResidentText
End Property

Public Property Get OG1D() As Long
  OG1D = mOG1D
End Property

Public Property Let OG1D(ByVal OG1D As Long)
  mOG1D = OG1D
End Property

Public Property Get OG2D() As Long
  OG2D = mOG2D
End Property

Public Property Let OG2D(ByVal OG2D As Long)
  mOG2D = OG2D
End Property

Public Property Get OG3D() As Long
  OG3D = mOG3D
End Property

Public Property Let OG3D(ByVal OG3D As Long)

  mOG3D = OG3D

End Property

Public Property Get OG4D() As Long

  OG4D = mOG4D

End Property

Public Property Let OG4D(ByVal OG4D As Long)

  mOG4D = OG4D

End Property

Public Property Get OG5D() As Long

  OG5D = mOG5D

End Property

Public Property Let OG5D(ByVal OG5D As Long)

  mOG5D = OG5D

End Property

Public Property Get OG6D() As Long

  OG6D = mOG6D

End Property

Public Property Let OG6D(ByVal OG6D As Long)

  mOG6D = OG6D

End Property

Public Property Get NG1D() As Long

  NG1D = mNG1D

End Property

Public Property Let NG1D(ByVal NG1D As Long)

  mNG1D = NG1D

End Property

Public Property Get NG2D() As Long

  NG2D = mNG2D

End Property

Public Property Let NG2D(ByVal NG2D As Long)

  mNG2D = NG2D

End Property

Public Property Get NG3D() As Long

  NG3D = mNG3D

End Property

Public Property Let NG3D(ByVal NG3D As Long)

  mNG3D = NG3D

End Property

Public Property Get NG4D() As Long

  NG4D = mNG4D

End Property

Public Property Let NG4D(ByVal NG4D As Long)

  mNG4D = NG4D

End Property

Public Property Get NG5D() As Long

  NG5D = mNG5D

End Property

Public Property Let NG5D(ByVal NG5D As Long)

  mNG5D = NG5D

End Property

Public Property Get NG6D() As Long

  NG6D = mNG6D

End Property

Public Property Let NG6D(ByVal NG6D As Long)

  mNG6D = NG6D

End Property

Public Property Get GG1() As Long

  GG1 = mGG1

End Property

Public Property Let GG1(ByVal GG1 As Long)

  mGG1 = GG1

End Property

Public Property Get GG2() As Long

  GG2 = mGG2

End Property

Public Property Let GG2(ByVal GG2 As Long)

  mGG2 = GG2

End Property

Public Property Get GG3() As Long

  GG3 = mGG3

End Property

Public Property Let GG3(ByVal GG3 As Long)

  mGG3 = GG3

End Property

Public Property Get GG4() As Long

  GG4 = mGG4

End Property

Public Property Let GG4(ByVal GG4 As Long)

  mGG4 = GG4

End Property

Public Property Get GG5() As Long

  GG5 = mGG5

End Property

Public Property Let GG5(ByVal GG5 As Long)

  mGG5 = GG5

End Property

Public Property Get GG6() As Long

  GG6 = mGG6

End Property

Public Property Let GG6(ByVal GG6 As Long)

  mGG6 = GG6

End Property

Public Property Get GG1D() As Long

  GG1D = mGG1D

End Property

Public Property Let GG1D(ByVal GG1D As Long)

  mGG1D = GG1D

End Property

Public Property Get GG2D() As Long

  GG2D = mGG2D

End Property

Public Property Let GG2D(ByVal GG2D As Long)

  mGG2D = GG2D

End Property

Public Property Get GG3D() As Long

  GG3D = mGG3D

End Property

Public Property Let GG3D(ByVal GG3D As Long)

  mGG3D = GG3D

End Property

Public Property Get GG4D() As Long

  GG4D = mGG4D

End Property

Public Property Let GG4D(ByVal GG4D As Long)

  mGG4D = GG4D

End Property

Public Property Get GG5D() As Long

  GG5D = mGG5D

End Property

Public Property Let GG5D(ByVal GG5D As Long)

  mGG5D = GG5D

End Property

Public Property Get GG6D() As Long

  GG6D = mGG6D

End Property

Public Property Let GG6D(ByVal GG6D As Long)

  mGG6D = GG6D

End Property

Public Property Get CurrentShift() As Integer

  CurrentShift = mCurrentShift

End Property

Public Property Let CurrentShift(ByVal CurrentShift As Integer)

  mCurrentShift = CurrentShift

End Property

Public Property Get SendCancel() As Boolean

    SendCancel = mSendCancel

End Property

Public Property Let SendCancel(ByVal SendCancel As Boolean)

    mSendCancel = SendCancel

End Property
