VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlarmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID                As Long
Public FC1               As Long
Public FC2               As Long
Public IDM               As Long
Public IDL               As Long
Public Serial            As String
Public EventDate         As Date
Public alarm             As Long
Public Tamper            As Long
Public IsLocator         As Long
Public Battery           As Long
Public LOCIDM            As Long
Public LOCIDL            As Long
Public ResidentID        As Long
Public RoomID            As Long
Public EventType         As Long
Public AlarmID           As Long

Public CLSPTI            As Long

Public Resident          As String
Public Room              As String

Public NameFirst         As String
Public NameLast          As String


Public Location          As String
Public Phone             As String
Public Notes             As String

Public User              As String
Public DeviceType        As String

Public Extra             As String

Public Count             As Long
Public Count2            As Long

Public Average           As Double
Public Average2          As Double

Public History           As Collection

Private mAnnounce        As String

Private mACKTime         As Date
Private mACKElapsed      As Long
Private mACKCancelTime   As Date

Private mCancelTime      As Date
Private mElapsedTime     As Long  ' seconds

' new with mobile 2019-02-10
Private mResponseTime    As Date
Private mResponseElapsed As Long  ' seconds
Private mFinalizeTime    As Date
Private mFinalizeElapsed As Long  ' seconds


Private mFinalizeText    As String
Private mResponder       As String '? same as use??


Public Function GetResidentInfo(ByVal ResidentID As Long) As String
        Dim rs As Recordset
10       On Error GoTo GetResidentInfo_Error

20      Set rs = ConnExecute("SELECT NameLast, NameFirst,phone,info FROM residents WHERE residentid = " & ResidentID)
30      If Not rs.EOF Then
40        Resident = ConvertLastFirst(rs("NameLast") & "", rs("NameFirst") & "")
          Notes = rs("info") & ""
          Phone = rs("phone") & ""
50      End If
60      rs.Close
70      Set rs = Nothing

GetResidentInfo_Resume:
80       On Error GoTo 0
90       Exit Function

GetResidentInfo_Error:

100     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modDBLib.GetResidentInfo." & Erl
110     Resume GetResidentInfo_Resume


End Function



Public Property Get ToString() As String
  Dim s As String
  Select Case EventType
     Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
     
       s = Announce
     Case EVT_ASSISTANCE
       s = Announce
     Case Else
      s = ""
  End Select
    
  ToString = Join(Array(DQ(Serial), DQ(Room), DQ(Resident), DQ(EventName), DQ(s), DQ(AlarmTimeToString), DQ(CancelTimeToString), DQ(ElapsedTimeToString), DQ(User)), ",")
  
  
End Property
Public Property Get ToCSVString() As String
  Dim s As String
  Select Case EventType
     Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
       s = Announce
     Case EVT_ASSISTANCE
       s = Announce
       
     Case Else
      s = ""
  End Select
    
  ToCSVString = Join(Array(DQ(Serial), DQ(Room), DQ(Resident), DQ(EventName), DQ(s), DQ(AlarmTimeToString), DQ(CancelTimeToString), DQ(ElapsedTimeToString), DQ(User)), ",")
  
End Property
Public Property Get ToTABBEDString() As String
  Dim s As String
  Select Case EventType
     Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
       s = Announce
     Case EVT_ASSISTANCE
       s = Announce
    Case EVT_EXTERN
       s = Announce
     Case Else
      s = ""
  End Select
    
  ToTABBEDString = Join(Array(Room, Resident, EventName, s, AlarmTimeToString, ElapsedTimeToString, Responder, FinalizeTimeToString, FinalizeText), vbTab)
  
End Property


Public Property Get AlarmTimeToString() As String
  AlarmTimeToString = Format(EventDate, "mm/dd/yy " & gTimeFormatString)  'hh:nnA/P")
  
End Property
Public Property Get ACKTimeToString() As String
  If ACKTime <= 0 Then
    ACKTimeToString = " "
  Else
    ACKTimeToString = Format(ACKTime, "mm/dd " & gTimeFormatString)  ' hh:nn A/P")
  End If
End Property
Public Property Get ACKElapsedTimeToString() As String
  If ACKElapsed > -1 Then
    ACKElapsedTimeToString = SecondsToTimeString(ACKElapsed)
  Else
   ACKElapsedTimeToString = " "
  End If
End Property

Public Property Get ResponseElapsedTimeToString() As String
  If ResponseElapsed > -1 Then
    ResponseElapsedTimeToString = SecondsToTimeString(ResponseElapsed)
  Else
   ResponseElapsedTimeToString = " "
  End If
End Property

Public Property Get FinalizeElapsedTimeToString() As String
  If FinalizeElapsed > -1 Then
    FinalizeElapsedTimeToString = SecondsToTimeString(FinalizeElapsed)
  Else
   FinalizeElapsedTimeToString = " "
  End If
End Property



Public Property Get ResponseTimeToString() As String

  If mResponseTime <= 0 Then
    ResponseTimeToString = SecondsToTimeString(mResponseTime)
  Else
   ResponseTimeToString = " "
  End If


End Property

Public Property Get FinalizeTimeToString() As String

  If mFinalizeElapsed > 0 Then
    FinalizeTimeToString = SecondsToTimeString(mFinalizeElapsed)
  Else
   FinalizeTimeToString = " "
  End If

End Property






Public Property Get CancelTimeToString() As String
  If CancelTime <= 0 Then
    CancelTimeToString = " "
  Else
    CancelTimeToString = Format(CancelTime, "mm/dd " & gTimeFormatString)  ' hh:nn A/P")
  End If
End Property

Public Property Get AverageTimeToString() As String

  If Average > 0 Then
    'AverageTimeToString = Format(DateAdd("s", Average, 0), "hh:nn:ss")
    
    AverageTimeToString = SecondsToTimeString(Average)
    
  Else
   AverageTimeToString = "N/A"
  End If

End Property

Public Property Get AverageTimeToString2() As String

  If Average > 0 Then
    'AverageTimeToString = Format(DateAdd("s", Average, 0), "hh:nn:ss")
    AverageTimeToString2 = SecondsToTimeString(Average2)
  Else
   AverageTimeToString2 = "N/A"
  End If

End Property





Public Property Get ElapsedTimeToString() As String
  If ElapsedTime > -1 Then
    'ElapsedTimeToString = Format(DateAdd("s", ElapsedTime, 0), "hh:nn:ss")
    ElapsedTimeToString = SecondsToTimeString(ElapsedTime)
  Else
   ElapsedTimeToString = " "
  End If
End Property

Public Property Get EventName()
  EventName = GetEventName(EventType)
End Property
Public Property Get ACKTime() As Date
  ACKTime = mACKTime
End Property
Public Property Let ACKElapsed(ByVal Value As Long)
  mACKElapsed = Value
End Property


Public Property Get ACKElapsed() As Long
  ACKElapsed = mACKElapsed
End Property
  
Public Property Get CancelTime() As Date
  CancelTime = mCancelTime
End Property
Public Property Let ElapsedTime(ByVal Value As Long)
  mElapsedTime = Value
End Property

Public Property Get ElapsedTime() As Long
  ElapsedTime = mElapsedTime
End Property
Public Property Get AlarmTime() As Date
  AlarmTime = EventDate
End Property

'ID | Room | Resident | EventType | AlarmTime | CancelTime | ElapsedTime | SoftwareOperator

Public Sub Parse(rs As Recordset)
  ID = rs("ID")  ' this record's ID
  Serial = rs("Serial") & ""
  EventDate = rs("EventDate")
  alarm = rs("Alarm")
  Tamper = rs("Tamper")
  Battery = rs("Battery")
  'ResidentID = rs("residentID")
  'RoomID = rs("RoomID")
  Room = rs("Room") & ""
  
  Announce = IIf((Len(Announce) = 0), rs("Announce") & "", "")
  EventType = rs("EventType")
  User = rs("Username") & ""
  AlarmID = rs("AlarmID")  ' parent alarm






End Sub

Public Function GetHistoryNew(ByVal ID As Long, rs As Recordset) As Long

      ' assist from monthly report
      'Global Const EVT_ASSISTANCE = 74
      'Global Const EVT_ASSISTANCE_RESPOND = 75
      'Global Const EVT_ASSISTANCE_FINALIZE = 76
      'Global Const EVT_ASSISTANCE_ACK = 77
      'Global Const EVT_ASSISTANCE_RESTORE = 78
      'Global Const EVT_ASSISTANCE_END = 79
10      On Error GoTo GetHistoryNew_Error

20      Do While Not rs.EOF

          'Debug.Print "ID, Alarmid "; ID, rs("Alarmid")

30        If ID = rs("Alarmid") Then

            '      If gIsJET Then
            '        Select Case rs("EventType")
            '          Case EVT_EMERGENCY_FINALIZE, EVT_ALERT_FINALIZE, EVT_ASSISTANCE_FINALIZE
            '             FinalizeText = rs("info") & ""
            '             FinalizeElapsed = DateDiff("s", EventDate, mCancelTime)
            '
            '
            '
            '
            '          Case EVT_EMERGENCY_RESPOND, EVT_ALERT_RESPOND, EVT_ASSISTANCE_RESPOND
            '             Responder = rs("username") & ""
            '
            '
            '          Case EVT_EMERGENCY_END, EVT_ALERT_END, EVT_ALERT_END, EVT_BATTERY_RESTORE, EVT_CHECKIN, EVT_TAMPER_RESTORE, EVT_EXTERN_END, EVT_EXTERN_TROUBLE_RESTORE, EVT_VACATION, EVT_VACATION_RETURN
            '            mCancelTime = rs("eventdate")
            '            If CLng(mCancelTime) <> 0 Then
            '              mElapsedTime = DateDiff("s", EventDate, mCancelTime)
            '            Else
            '              mElapsedTime = -1  ' no end time recorded
            '            End If
            '
            '          Case EVT_ALERT_ACK, EVT_EMERGENCY_ACK
            '            mACKTime = rs("eventdate")
            '            mACKCancelTime = mACKTime
            '            If CLng(mACKTime) <> 0 Then
            '              mACKElapsed = DateDiff("s", EventDate, mACKTime)
            '            Else
            '              mACKElapsed = -1  ' no end time recorded
            '            End If
            '            User = rs("username") & ""
            '        End Select
            '      Else
            
            Dim EvType As Long
40          EvType = Val(rs("EventType")) & ""
            
50          Select Case EvType

              Case EVT_EMERGENCY_FINALIZE, EVT_ALERT_FINALIZE, EVT_ASSISTANCE_FINALIZE, EVT_ASSISTANCE_ACK
                
60              Debug.Print "EVT_EMERGENCY_FINALIZE " & rs("eventdate")
70              If FinalizeText = "" Then
80                FinalizeText = rs("info") & ""
90              End If
                
100             If CDbl(mCancelTime) <= 0 Then
                  
110               FinalizeElapsed = -1
120             Else
130               If mCancelTime <= 0 Then
140                 FinalizeElapsed = -1
150               Else
160                 FinalizeElapsed = DateDiff("s", mCancelTime, rs("Eventdate"))
170               End If

180             End If

190           Case EVT_EMERGENCY_RESPOND, EVT_ALERT_RESPOND, EVT_ASSISTANCE_RESPOND
200             Responder = rs("username") & ""

210             Debug.Print "EVT_EMERGENCY_RESPOND " & rs("eventdate") & " " & Responder

220           Case EVT_EMERGENCY_RESTORE, EVT_ALERT_RESTORE, EVT_ASSISTANCE_RESTORE, EVT_ASSISTANCE_END  ' moved these here for advanced reporting
230             Debug.Print "EVT_EMERGENCY_RESTORE " & rs("eventdate")
240             If FinalizeText = "" Then
250               FinalizeText = rs("info") & ""
260             End If

270             mCancelTime = rs("eventdate")
280             If CLng(mCancelTime) <> 0 Then
290               mElapsedTime = DateDiff("s", EventDate, mCancelTime)
300             Else
310               mElapsedTime = -1  ' no end time recorded
320             End If

              
330           Case EVT_EMERGENCY_END, EVT_ALERT_END, EVT_ALERT_END, EVT_BATTERY_RESTORE, EVT_CHECKIN, EVT_TAMPER_RESTORE, EVT_EXTERN_END, EVT_EXTERN_TROUBLE_RESTORE, EVT_VACATION, EVT_VACATION_RETURN
340             Debug.Print "EVT_EMERGENCY_END " & rs("eventdate")
350             mCancelTime = rs("eventdate")
                
360             If CLng(mCancelTime) <> 0 Then
370               mElapsedTime = DateDiff("s", EventDate, mCancelTime)
380             Else
390               mElapsedTime = -1  ' no end time recorded
400             End If
              
              
              'Case EVT_EMERGENCY_END, EVT_ALERT_END, EVT_ALERT_END, EVT_BATTERY_RESTORE, EVT_CHECKIN, EVT_TAMPER_RESTORE, EVT_EXTERN_END, EVT_EXTERN_TROUBLE_RESTORE, EVT_VACATION, EVT_VACATION_RETURN
410           Case EVT_BATTERY_RESTORE, EVT_CHECKIN, EVT_TAMPER_RESTORE, EVT_EXTERN_END, EVT_EXTERN_TROUBLE_RESTORE, EVT_VACATION, EVT_VACATION_RETURN
420             mCancelTime = rs("eventdate")
430             If CLng(mCancelTime) <> 0 Then
440               mElapsedTime = DateDiff("s", EventDate, mCancelTime)
450             Else
460               mElapsedTime = -1  ' no end time recorded
470             End If
              
480           Case EVT_ALERT_ACK, EVT_EMERGENCY_ACK, EVT_EXTERN_ACK
                
490             mACKTime = rs("eventdate")
500             mACKCancelTime = mACKTime
510             If CLng(mACKTime) <> 0 Then
520               mACKElapsed = DateDiff("s", EventDate, mACKTime)
530             Else
540               mACKElapsed = -1   ' no end time recorded
550             End If
560             User = rs("username") & ""
570             If Len(Responder) = 0 Then
                  'Responder = User
580             End If
                
590         End Select

            'End If
600         rs.MoveNext
610       Else
620         If ID > rs("Alarmid") Then
630           rs.MoveNext
640         Else
650           Exit Do
660         End If
670       End If
680     Loop


GetHistoryNew_Resume:

690     On Error GoTo 0
700     Exit Function

GetHistoryNew_Error:

710     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarmHistory.GetHistoryNew." & Erl
720     Resume GetHistoryNew_Resume

End Function


'Public Sub GetHistory(conn As Connection)
'  Dim rs            As Recordset
'  Dim SQL           As String
'  Dim AlarmHistory  As cAlarmHistory
'  Dim CancelTime    As Date
'
'  SQL = "SELECT ID, serial, eventdate, alarm, tamper, battery,residentid, roomid, announce, eventtype, username , alarmid  FROM alarms WHERE AlarmID = " & ID
'  Set rs = ConnExecute(SQL)
'  Do Until rs.EOF
'    Set AlarmHistory = New cAlarmHistory
'    AlarmHistory.Parse rs
'    Select Case AlarmHistory.EventType
'    Case EVT_EMERGENCY_END, EVT_ALERT_END, EVT_ALERT_END, EVT_BATTERY_RESTORE, EVT_CHECKIN, EVT_TAMPER_RESTORE, EVT_EXTERN_END, EVT_EXTERN_TROUBLE_RESTORE, EVT_VACATION, EVT_VACATION_RETURN
'      mCancelTime = AlarmHistory.EventDate
'      If CLng(AlarmHistory.EventDate) <> 0 Then
'        'mCancelTime = AlarmHistory.EventDate
'        mElapsedTime = DateDiff("s", EventDate, AlarmHistory.EventDate)
'      Else
'        mElapsedTime = -1  ' no end time recorded
'      End If
'    Case EVT_ALERT_ACK, EVT_EMERGENCY_ACK
'      mACKTime = AlarmHistory.EventDate
'      mACKCancelTime = AlarmHistory.EventDate
'      'If gElapsedEqACK = 1 Then
'      '  mCancelTime = AlarmHistory.EventDate
'      'End If
'
'      If mACKTime <> 0 Then
'        mACKElapsed = DateDiff("s", EventDate, AlarmHistory.EventDate)
'      Else
'        mACKElapsed = -1  ' no end time recorded
'      End If
'      User = rs("username") & ""
'    End Select
'    '    End If
'    rs.MoveNext
'  Loop
'  rs.Close
'  Set rs = Nothing
'End Sub

Private Sub Class_Initialize()
  Set History = New Collection
End Sub


Public Property Get Announce() As String
  Announce = mAnnounce
End Property

Public Property Let Announce(ByVal Announce As String)
  mAnnounce = Announce
End Property


Public Property Get ResponseTime() As Date

  ResponseTime = mResponseTime

End Property

Public Property Let ResponseTime(ByVal ResponseTime As Date)

  mResponseTime = ResponseTime

End Property

Public Property Get ResponseElapsed() As Long

  ResponseElapsed = mResponseElapsed

End Property

Public Property Let ResponseElapsed(ByVal ResponseElapsed As Long)

  mResponseElapsed = ResponseElapsed

End Property


Public Property Get FinalizeTime() As Date

  FinalizeTime = mFinalizeTime

End Property

Public Property Let FinalizeTime(ByVal FinalizeTime As Date)

  mFinalizeTime = FinalizeTime

End Property

Public Property Get FinalizeElapsed() As Long

  FinalizeElapsed = mFinalizeElapsed

End Property

Public Property Let FinalizeElapsed(ByVal FinalizeElapsed As Long)

  mFinalizeElapsed = FinalizeElapsed

End Property

Public Property Get FinalizeText() As String

  FinalizeText = mFinalizeText

End Property

Public Property Let FinalizeText(ByVal FinalizeText As String)

  mFinalizeText = FinalizeText

End Property

Public Property Get Responder() As String

  Responder = mResponder

End Property

Public Property Let Responder(ByVal Responder As String)

  mResponder = Responder

End Property
