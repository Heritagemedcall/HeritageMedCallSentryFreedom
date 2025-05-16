VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWaypoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWaypoints As Collection


Public Property Get Count() As Long
  Count = mWaypoints.Count
End Property

Public Property Get waypoint(ByVal index As Integer) As cWayPoint
  If index > 0 And index <= mWaypoints.Count Then
    Set waypoint = mWaypoints(index)
  Else
    Set waypoint = New cWayPoint
  End If
End Property

Private Sub Class_Initialize()
  Set mWaypoints = New Collection
End Sub

Public Sub ClearChecked()
  Dim w As cWayPoint
  For Each w In mWaypoints
    w.Checked = False
  Next
End Sub

Public Function RemoveByID(ByVal ID As Long)
  
  On Error Resume Next
  mWaypoints.Remove "x" & ID
  
  
'  Dim j As Integer
'
'
'
'  For j = 1 To mWaypoints.Count
'    If mWaypoints(j).ID = ID Then
'      mWaypoints.Remove j
'      Exit For
'    End If
'  Next

End Function
Public Function Exists(ByVal ID As Long) As Boolean
  Dim j As Integer
  Dim waypoint As cWayPoint
  
  On Error Resume Next
  Set waypoint = mWaypoints.Item("x" & ID)
  
  If Not waypoint Is Nothing Then
    waypoint.Checked = True
    Exists = True
  End If
      
'  For j = 1 To mWaypoints.Count
'
'    If mWaypoints(j).ID = ID Then
'      mWaypoints(j).Checked = True
'      Exists = True
'      Exit For
'    End If
'  Next

End Function
Public Sub RemoveDeadWood()
  Dim j As Integer
  
  For j = mWaypoints.Count To 1 Step -1
    DoEvents
    If (Not mWaypoints(j).Checked) Then
      mWaypoints.Remove j
    End If
  Next

  ClearChecked
End Sub
Function AddWayPoint(rs As ADODB.Recordset) As Boolean
  Dim waypoint As cWayPoint
  Set waypoint = New cWayPoint

  waypoint.ID = rs("ID")
  waypoint.InternalID = "x" & waypoint.ID
  waypoint.Description = rs("Description") & ""
  waypoint.Building = rs("building") & ""
  waypoint.Floor = rs("floor") & ""
  waypoint.Wing = rs("wing") & ""
  waypoint.Repeater1 = rs("repeater1") & ""
  waypoint.Repeater2 = rs("repeater2") & ""
  waypoint.Repeater3 = rs("repeater3") & ""
  waypoint.Signal1 = Val(rs("Signal1") & "")
  waypoint.Signal2 = Val(rs("Signal2") & "")
  waypoint.Signal3 = Val(rs("Signal3") & "")
  waypoint.Checked = True
  mWaypoints.Add waypoint, waypoint.InternalID

End Function

Function Locate(Alarm As cAlarm)
      ' how this is supposed to work:

      ' see if there are any trifectas
      ' if 1, we're done
      ' if > 1 then narrow list if possible, Done
      ' if 0 then widen signal range
      ' if still 0 then try 3 any order
      ' if 1 we're done
      ' if > 1 then narrow list if possible, Done
      ' if still 0 then try 2 any order
      ' if 1 we're done
      ' if > 1 then narrow list if possible, Done
      ' if still 0 then use ???

      ' 12/13/05 conversation w/JT
      ' in case of ties, two approaches (could be more)
      ' closest sum of signals
      ' least maximum deviation

10      On Error GoTo Locate_Error

20      dbgloc "Locating" & vbCrLf

        Dim waypoint    As cWayPoint
        Dim pool        As Collection  ' a collection of candidates

        Dim w           As cWayPoint
        Dim Sequence    As String  ' 1,2 or 2,1 or 3,1 etc

        Dim d           As Double  ' Deviation
30      d = 5

        Dim Signal1     As Double  ' signal levels
        Dim Signal2     As Double
        Dim Signal3     As Double

        Dim r1          As String  ' Serials of repeaters
        Dim r2          As String
        Dim r3          As String

        Dim j           As Integer

        Dim s           As String

40      Signal1 = Alarm.Signal1
50      Signal2 = Alarm.Signal2
60      Signal3 = Alarm.Signal3

        ' for testing : Alarm.Repeater2 = "01142000"

70      r1 = Alarm.Repeater1
80      r2 = Alarm.Repeater2
90      r3 = Alarm.Repeater3
        
        Debug.Print "Locating"
        
        Debug.Print r1, Signal1
        Debug.Print r2, Signal2
        Debug.Print r3, Signal3
        Debug.Print ""

100     Set pool = New Collection

110     If mWaypoints.Count > 0 Then
120       For Each waypoint In mWaypoints
            'first see if we have any Trifecta Winners
130         If waypoint.Repeater1 = r1 And waypoint.Repeater2 = r2 And waypoint.Repeater3 = r3 Then
140           pool.Add waypoint
150         End If
160       Next
170       dbgloc "TRI.Count " & pool.Count

180       If pool.Count = 1 Then  ' only one winner...
190         dbgloc "Only one Trifecta"
200         Set w = pool(1)
210         Alarm.locationtext = w.Description
            ' done!
220       ElseIf pool.Count > 0 Then
            ' cull any that deviate more than Deviation of D (initally 5)
230         Set pool = LevelMatch(pool, Signal1, Signal2, Signal3, d)
240         If pool.Count = 1 Then
250           dbgloc "Level Matched Pool to 1"
260           Set w = pool(1)
270           Alarm.locationtext = w.Description
280         ElseIf pool.Count > 1 Then
290           dbgloc "Level Matched Pool to " & pool.Count & vbCrLf
300           Set pool = NarrowMatch(pool, Signal1, Signal2, Signal3, d)
310           If pool.Count >= 1 Then  ' pick the first one (or only one)
320             dbgloc "Narrowed Pool to " & pool.Count
330             Set w = pool(1)
340             Alarm.locationtext = w.Description
350           ElseIf pool.Count = 1 Then
360             dbgloc "Narrowed Pool to 1"
370             Set w = pool(1)
380             Alarm.locationtext = w.Description
390           Else
400             dbgloc "Narrowed Pool to: Shouldn't happen." & pool.Count & vbCrLf
410           End If
420         Else  ' need to widen match
430           Set pool = WidenMatch(pool, Signal1, Signal2, Signal3, d)
440           dbgloc "Widen Matched Pool to " & pool.Count
450           If pool.Count > 0 Then
460             Set w = pool(1)
470             Alarm.locationtext = w.Description
480           End If
490         End If
500       Else  ' No Trifecta, but see if we picked the top 3 horses in any order
510         dbgloc "No Trifectas"
520         Set pool = MisOrderedRepeaters(mWaypoints, Alarm)  ' returns pool of 3 any-order matches
530         If pool.Count = 1 Then
540           dbgloc "Picked Top 3 Out-of-Order Count 1"
550           Set w = pool(1)
560           Alarm.locationtext = w.Description
570         ElseIf pool.Count > 0 Then  ' could try and narrow
580           dbgloc "Picked Top 3 Out-of-Order Count " & pool.Count
590           Set pool = FixOrder(pool, Alarm)
600           Set pool = LevelMatch(pool, Signal1, Signal2, Signal3, d)
610           If pool.Count = 1 Then
620             Set w = pool(1)
630             Alarm.locationtext = w.Description
640           ElseIf pool.Count > 1 Then
650             Set pool = NarrowMatch(pool, Signal1, Signal2, Signal3, d)
660             If pool.Count > 0 Then
670               Set w = pool(1)
680               Alarm.locationtext = w.Description
690             End If
700           Else
710             Set pool = WidenMatch(pool, Signal1, Signal2, Signal3, d)
720             dbgloc "Widen Matched Pool to " & pool.Count
730             If pool.Count > 0 Then
740               Set w = pool(1)
750               Alarm.locationtext = w.Description
760             End If
770           End If
780         Else  ' pool count = 0
              ' need to see if we have 2 of 3 match
              ' need 1,2 or 2,1 order, maybe 1,3 ? 3,1 ? 2,3 ? 3,2 ?
790           dbgloc "NO Top 3 " & pool.Count
800           For j = 1 To 6
810             Select Case j
                  Case 1
820                 Sequence = "1,2"
830               Case 2
840                 Sequence = "2,1"
850               Case 3
860                 Sequence = "1,3"
870               Case 4
880                 Sequence = "3,1"
890               Case 5
900                 Sequence = "2,3"
910               Case 6
920                 Sequence = "3,2"
930             End Select

940             Set pool = Pick2Repeaters(mWaypoints, Alarm, Sequence)
950             dbgloc Sequence & " Pair Count " & pool.Count
960             If pool.Count = 1 Then
970               dbgloc Sequence & "One Pair Match"
980               Set w = pool(1)
990               Alarm.locationtext = w.Description
1000              Exit For
1010            ElseIf pool.Count > 0 Then  'try and narrow
1020              Set pool = LevelMatch2(pool, Signal1, Signal2, d)
1030              If pool.Count = 1 Then
1040                dbgloc Sequence & " Best Level Pair Count 1"
1050                Set w = pool(1)
1060                Alarm.locationtext = w.Description
1070                Exit For
1080              ElseIf pool.Count > 1 Then
1090                dbgloc Sequence & " Level Pair Count " & pool.Count
1100                Set pool = NarrowMatch2(pool, Signal1, Signal2, d)
1110                If pool.Count > 0 Then
1120                  dbgloc Sequence & " First Narrowed Pair  Count " & pool.Count
1130                  Set w = pool(1)
1140                  Alarm.locationtext = w.Description
1150                  Exit For
1160                End If
1170              End If
1180            End If
1190          Next
1200          If w Is Nothing Then
1210            If mWaypoints.Count > 0 Then
                  ' get Repeatername from waypoint

1220              Alarm.locationtext = GetRepeaterName(Alarm.Repeater1)
1230            Else
1240              Set w = New cWayPoint
1250              Alarm.locationtext = GetRepeaterName(Alarm.Repeater1)
1260            End If  ' mWaypoints.Count > 0
1270          End If  ' pool count = 0
1280        End If
1290      End If
1300    End If

1310    If Not w Is Nothing Then

          'dbgloc "" & vbCrLf
1320      s = left("DEVICE" & "         ", 10) & " "
1330      dbgloc s & Alarm.Repeater1 & " " & Format(Alarm.Signal1, "0") & " " & Alarm.Repeater2 & " " & Format(Alarm.Signal2, "0") & " " & Alarm.Repeater3 & " " & Format(Alarm.Signal3, "0")
1340      s = left(w.Description & "         ", 10) & " "
1350      dbgloc s & w.Repeater1 & " " & Format(w.Signal1, "0") & " " & w.Repeater2 & " " & Format(w.Signal2, "0") & " " & w.Repeater3 & " " & Format(w.Signal3, "0")
          'dbgloc "" & vbCrLf
1360    Else
1370      dbgloc left("DEVICE" & "         ", 10) & " " & r1 & " " & Format(Signal1, "0") & " " & r2 & " " & Format(Signal2, "0") & " " & r3 & " " & Format(Signal3, "0")
1380    End If

1390    If Len(Alarm.locationtext) > 0 Then
1400      If Alarm.Guid <> 0 Then
1410        UpdateAnnounceAndLocationText Alarm
1420      End If
1430    End If

Locate_Resume:
1440    On Error GoTo 0
1450    Exit Function

Locate_Error:

1460    LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cWaypoints.Locate." & Erl
1470    Resume Locate_Resume


End Function

Private Function FixOrder(pool As Collection, Alarm As cAlarm) As Collection
  Dim NewPool As Collection
  Dim w       As cWayPoint
  Dim NewW    As cWayPoint
  Dim r1 As String
  Dim r2 As String
  Dim r3 As String
  
  r1 = Alarm.Repeater1
  r2 = Alarm.Repeater2
  r3 = Alarm.Repeater3
  
  Set NewPool = New Collection
  For Each w In pool
    Set NewW = New cWayPoint
    If w.Repeater1 = r1 Then  ' 2 and 3 are swapped
      NewW.Serial = w.Serial
      NewW.Description = w.Description
      NewW.Floor = w.Floor
      NewW.Building = w.Building
      NewW.Repeater1 = r1
      NewW.Signal1 = w.Signal1
      NewW.Repeater2 = r3     ' 2 and 3 are swapped
      NewW.Signal2 = w.Signal3
      NewW.Repeater3 = r2
      NewW.Signal3 = w.Signal2
      NewPool.Add NewW
    ElseIf w.Repeater2 = r2 Then '1 and 3 are swapped
      NewW.Serial = w.Serial
      NewW.Description = w.Description
      NewW.Floor = w.Floor
      NewW.Building = w.Building
      NewW.Repeater2 = r2
      NewW.Signal2 = w.Signal2
      NewW.Repeater1 = r3     '1 and 3 are swapped
      NewW.Signal1 = w.Signal3
      NewW.Repeater3 = r1
      NewW.Signal3 = w.Signal1
      NewPool.Add NewW
    Else                      '1 and 2 are swapped
      NewW.Serial = w.Serial
      NewW.Description = w.Description
      NewW.Floor = w.Floor
      NewW.Building = w.Building
      NewW.Repeater3 = r3
      NewW.Signal3 = w.Signal3
      NewW.Repeater1 = r2     '1 and 2 are swapped
      NewW.Signal1 = w.Signal2
      NewW.Repeater2 = r1
      NewW.Signal2 = w.Signal1
      NewPool.Add NewW
    End If
  Next
  Set pool = Nothing
  Set FixOrder = NewPool
  
End Function

'Private Function BestRepeater(Waypoints As Collection, Alarm As cAlarm) As Collection
'  Dim ar1 As String
'  Dim ar2 As String
'  Dim ar3 As String
'
'  Dim wr1 As String
'
'
'  Dim pool      As Collection
'  Dim pool2     As Collection
'  Dim waypoint  As cWayPoint
'
'  Dim d         As Double
'
'  ar1 = Alarm.Repeater1
'  ar2 = Alarm.Repeater2
'  ar3 = Alarm.Repeater3
'
'  Set pool = New Collection
'  For Each waypoint In Waypoints
'    ' maybe the're out of order
'    wr1 = waypoint.Repeater1
'    If ar1 = wr1 Then  ' 1
'      pool.Add waypoint
'    End If
'  Next
'
'  If pool.count > 0 Then
'    Set pool2 = New Collection
'    ' rank repeater
'    For d = 0 To 50
'      For Each waypoint In pool
'        If InRange(waypoint.Signal1, Alarm.Signal1, d) Then
'          pool2.Add waypoint
'        End If
'      Next
'      If pool2.count > 0 Then
'        Set pool = pool2
'        Exit For
'      End If
'    Next
'  End If
'  Set BestRepeater = pool
'
'End Function


Private Function Pick2Repeaters(Waypoints As Collection, Alarm As cAlarm, ByVal Sequence As String) As Collection
  Dim ar1 As String
  Dim ar2 As String
  Dim ar3 As String


  Dim pool As Collection
  Dim waypoint As cWayPoint
  Dim wp As cWayPoint


  ' returns a pool of waypoints where the first two repeaters are matched to the alarm data, third repeater is a wildcard

  ar1 = Alarm.Repeater1
  ar2 = Alarm.Repeater2
  ar3 = Alarm.Repeater3
  
  dbgloc "Sequence:  " & Sequence & vbCrLf
  dbgloc "Repeaters: " & ar1 & " " & ar2 & " " & ar3 & vbCrLf
  

  Set pool = New Collection
  Sequence = LCase(Sequence)

  For Each waypoint In Waypoints
  
    'dbgloc "Waypoint : " & waypoint.Repeater1 & " " & waypoint.Repeater2 & " " & waypoint.Repeater3 & " " & waypoint.Description & vbCrLf
    
    Select Case Sequence
      Case "1,2"
        If ar1 = waypoint.Repeater1 And ar2 = waypoint.Repeater2 Then  ' 1 and 2 ' natural
          Set wp = New cWayPoint
          wp.Repeater1 = waypoint.Repeater1
          wp.Repeater2 = waypoint.Repeater2
          wp.Signal1 = waypoint.Signal1
          wp.Signal2 = waypoint.Signal2
          wp.Building = waypoint.Building
          wp.Description = waypoint.Description
          wp.Floor = waypoint.Floor
          pool.Add wp
        End If
      Case "2,1"
        If ar1 = waypoint.Repeater2 And ar2 = waypoint.Repeater1 Then  ' 2 and 1
          Set wp = New cWayPoint
          wp.Repeater1 = waypoint.Repeater2
          wp.Repeater2 = waypoint.Repeater1
          wp.Signal1 = waypoint.Signal2
          wp.Signal2 = waypoint.Signal1
          wp.Building = waypoint.Building
          wp.Description = waypoint.Description
          wp.Floor = waypoint.Floor
          pool.Add wp
        End If
      
      Case "1,3"
        If ar1 = waypoint.Repeater1 And ar2 = waypoint.Repeater3 Then  ' 1 and 3
          Set wp = New cWayPoint
          wp.Repeater1 = waypoint.Repeater1
          wp.Repeater2 = waypoint.Repeater3
          wp.Signal1 = waypoint.Signal1
          wp.Signal2 = waypoint.Signal3
          wp.Building = waypoint.Building
          wp.Description = waypoint.Description
          wp.Floor = waypoint.Floor
          pool.Add wp
        End If
      Case "3,1"
        If ar1 = waypoint.Repeater3 And ar2 = waypoint.Repeater1 Then  ' 3 and 1
          Set wp = New cWayPoint
          wp.Repeater1 = waypoint.Repeater3
          wp.Repeater2 = waypoint.Repeater1
          wp.Signal1 = waypoint.Signal3
          wp.Signal2 = waypoint.Signal1
          wp.Building = waypoint.Building
          wp.Description = waypoint.Description
          wp.Floor = waypoint.Floor
          pool.Add wp
        End If
      Case "2,3"
        If ar1 = waypoint.Repeater2 And ar2 = waypoint.Repeater3 Then  ' 2 and 3
          Set wp = New cWayPoint
          wp.Repeater1 = waypoint.Repeater2
          wp.Repeater2 = waypoint.Repeater3
          wp.Signal1 = waypoint.Signal2
          wp.Signal2 = waypoint.Signal3
          wp.Building = waypoint.Building
          wp.Description = waypoint.Description
          wp.Floor = waypoint.Floor
          pool.Add wp
        End If
      
      Case "3,2"
        If ar1 = waypoint.Repeater3 And ar2 = waypoint.Repeater2 Then  ' 3 and 2
          Set wp = New cWayPoint
          wp.Repeater1 = waypoint.Repeater3
          wp.Repeater2 = waypoint.Repeater2
          wp.Signal1 = waypoint.Signal3
          wp.Signal2 = waypoint.Signal2
          wp.Building = waypoint.Building
          wp.Description = waypoint.Description
          wp.Floor = waypoint.Floor
          pool.Add wp
        End If
    
    
    End Select

  Next

  Set Pick2Repeaters = pool



End Function
'Private Function Best2of3Repeaters(Waypoints As Collection, Alarm As cAlarm) As Collection
'  Dim ar1 As String
'  Dim ar2 As String
'  Dim ar3 As String
'
'  Dim wr1 As String
'  Dim wr2 As String
'  Dim wr3 As String
'
'  Dim pool As Collection
'  Dim waypoint As cWayPoint
'
'  ar1 = Alarm.Repeater1
'  ar2 = Alarm.Repeater2
'  ar3 = Alarm.Repeater3
'
'  Set pool = New Collection
'  For Each waypoint In Waypoints
'    ' maybe the're out of order
'    wr1 = waypoint.Repeater1
'    wr2 = waypoint.Repeater2
'    wr3 = waypoint.Repeater3
'
'    If 0 Then
'      ' nada just for pretty
'    ElseIf ar1 = wr1 And ar2 = wr2 Then  ' 1 and 2
'      pool.Add waypoint
'    ElseIf ar1 = wr1 And ar2 = wr3 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr2 And ar2 = wr1 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr2 And ar2 = wr3 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr3 And ar2 = wr1 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr3 And ar2 = wr2 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr1 And ar3 = wr2 Then  ' 1 and 3
'      pool.Add waypoint
'    ElseIf ar1 = wr1 And ar3 = wr3 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr2 And ar3 = wr1 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr2 And ar3 = wr3 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr3 And ar3 = wr1 Then
'      pool.Add waypoint
'    ElseIf ar1 = wr3 And ar3 = wr2 Then
'      pool.Add waypoint
'    ElseIf ar2 = wr1 And ar3 = wr2 Then  ' 2 and 3
'      pool.Add waypoint
'    ElseIf ar2 = wr1 And ar3 = wr3 Then
'      pool.Add waypoint
'    ElseIf ar2 = wr2 And ar3 = wr1 Then
'      pool.Add waypoint
'    ElseIf ar2 = wr2 And ar3 = wr3 Then
'      pool.Add waypoint
'    ElseIf ar2 = wr3 And ar3 = wr1 Then
'      pool.Add waypoint
'    ElseIf ar2 = wr3 And ar3 = wr2 Then
'      pool.Add waypoint
'    End If
'  Next
'
'  Set Best2of3Repeaters = pool
'
'End Function
Private Function MisOrderedRepeaters(Waypoints As Collection, Alarm As cAlarm) As Collection
  Dim ar1 As String
  Dim ar2 As String
  Dim ar3 As String

  Dim wr1 As String
  Dim wr2 As String
  Dim wr3 As String

  Dim pool As Collection
  Dim waypoint As cWayPoint

  ar1 = Alarm.Repeater1
  ar2 = Alarm.Repeater2
  ar3 = Alarm.Repeater3

  Set pool = New Collection
  For Each waypoint In Waypoints
    ' maybe the're out of order
    wr1 = waypoint.Repeater1
    wr2 = waypoint.Repeater2
    wr3 = waypoint.Repeater3


    If ar2 = wr1 Then
      If ar1 = wr2 Then
        If ar3 = wr3 Then    '2,1,3
          pool.Add waypoint
        End If
      ElseIf ar1 = wr3 Then
        If ar3 = wr2 Then    '2,3,1
          pool.Add waypoint
        End If
      End If
    ElseIf ar3 = wr1 Then
      If ar1 = wr2 Then
        If ar2 = wr3 Then    '3,1,2
          pool.Add waypoint
        End If
      ElseIf ar1 = wr3 Then  '3,2,1
        If ar2 = wr2 Then
          pool.Add waypoint
        End If
      End If
    ElseIf ar1 = wr1 Then
      If ar2 = wr3 Then
        If ar3 = wr2 Then    '1,3,2
         pool.Add waypoint
        End If
      Else
        ' NADA               ' we don't do 1,2,3
      End If
    End If
  Next

  Set MisOrderedRepeaters = pool

End Function

Private Function LevelMatch(Waypoints As Collection, ByVal Signal1 As Double, ByVal Signal2 As Double, ByVal Signal3 As Double, ByVal d As Double) As Collection
  Dim waypoint As cWayPoint
  Dim Results As Collection

  ' this is assuming that there were no matches... must be within +/- D level for each of 3 levels

  Set Results = New Collection
  For Each waypoint In Waypoints
    If InRange(waypoint.Signal1, Signal1, d) And InRange(waypoint.Signal2, Signal2, d) And InRange(waypoint.Signal3, Signal3, d) Then
      Results.Add waypoint
    End If
  Next
  If Results.Count > 0 Then
    Set LevelMatch = Results
  Else
    Set LevelMatch = Waypoints  ' couldn't do any better
  End If

End Function
Private Function LevelMatch2(Waypoints As Collection, ByVal Signal1 As Double, ByVal Signal2 As Double, ByVal d As Double) As Collection
  Dim waypoint As cWayPoint
  Dim Results As Collection
  
  ' this is assuming that there were no matches... must be within +/- D level for each of 3 levels

  Set Results = New Collection
  For Each waypoint In Waypoints
      If InRange(waypoint.Signal1, Signal1, d) And InRange(waypoint.Signal2, Signal2, d) Then
      Results.Add waypoint
    End If
  Next
  If Results.Count > 0 Then
    Set LevelMatch2 = Results
  Else
    Set LevelMatch2 = Waypoints  ' couldn't do any better
  End If

End Function


Private Function WidenMatch(Waypoints As Collection, ByVal Signal1 As Double, ByVal Signal2 As Double, ByVal Signal3 As Double, ByVal d As Double) As Collection
  Dim waypoint    As cWayPoint
  Dim Results     As Collection

  ' this is assuming that there were no matches... see if we can find some by loosening the deviation

  Set Results = New Collection
  For d = d + 1 To d + 10
    For Each waypoint In Waypoints
      If InRange(waypoint.Signal1, Signal1, d) And InRange(waypoint.Signal2, Signal2, d) And InRange(waypoint.Signal3, Signal3, d) Then
        Results.Add waypoint
      End If
    Next
    If Results.Count > 0 Then  ' there were some at this level, we're done
      Exit For
    End If
  Next
  If Results.Count > 0 Then
    Set WidenMatch = Results  ' could be none or more
  Else
    Set WidenMatch = Waypoints  ' couldn't do any better
  End If
End Function
Private Function NarrowMatch(Waypoints As Collection, ByVal Signal1 As Double, ByVal Signal2 As Double, ByVal Signal3 As Double, ByVal d As Double) As Collection
  Dim waypoint      As cWayPoint
  Dim pool          As Collection
  Dim Results       As Collection

  ' initally it's all matching waypoints passed in

  Set pool = Waypoints
  Set Results = New Collection
  For d = d - 1 To 0 Step -1

    For Each waypoint In pool
      If InRange(waypoint.Signal1, Signal1, d) And InRange(waypoint.Signal2, Signal2, d) And InRange(waypoint.Signal3, Signal3, d) Then
        Results.Add waypoint
      End If
    Next
    If Results.Count = 0 Then  ' nothing fit, so go with our current Pool of candidates
      Set Results = pool
      Exit For
    End If
    Set pool = Results
    Set Results = New Collection
  Next

  If Results.Count > 0 Then
    Set NarrowMatch = Results
  Else
    Set NarrowMatch = Waypoints
  End If


End Function

Private Function NarrowMatch2(Waypoints As Collection, ByVal Signal1 As Double, ByVal Signal2 As Double, ByVal d As Double) As Collection
  Dim waypoint      As cWayPoint
  Dim pool          As Collection
  Dim Results       As Collection

  Set pool = Waypoints
  Set Results = New Collection
  For d = d - 1 To 0 Step -1

    For Each waypoint In pool
       If InRange(waypoint.Signal1, Signal1, d) And InRange(waypoint.Signal2, Signal2, d) Then
        Results.Add waypoint
      End If
    Next
    If Results.Count = 0 Then  ' nothing fit, so go with our current Pool of candidates
      Set Results = pool
      Exit For
    End If
    Set pool = Results
    Set Results = New Collection
  Next

  If Results.Count > 0 Then
    Set NarrowMatch2 = Results
  Else
    Set NarrowMatch2 = Waypoints
  End If


End Function

'Private Function AverageMatch(Waypoints As Collection, ByVal Average As Double, ByVal MaxD As Double) As Collection
'  Dim waypoint As cWayPoint
'  Dim Results As Collection
'  Dim D As Double
'  '
'  Set Results = New Collection
'  '
'  '  For D = 0 To MaxD
'  '    For Each waypoint In Waypoints
'  '      If InRange(waypoint.Average, Average, D) Then
'  '        Results.Add waypoint
'  '      End If
'  '    Next
'  '    If Results.Count > 0 Then  ' close enuf
'  '      Exit For
'  '    End If
'  '  Next
'  '
'  Set AverageMatch = Results
'  '
'End Function

Private Function InRange(ByVal a As Double, ByVal b As Double, ByVal limit As Double) As Boolean
  InRange = Abs(a - b) <= limit
End Function
