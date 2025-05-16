VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExceptionReportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mNumEntries As Long
Private mSumElapsed As Long


Private rows             As Collection
Public Dest              As Integer

Private hfile            As Integer
Const DEST_SCREEN = 0
Const DEST_PRINTER = 1
Const DEST_FILE = 2

Public lv                As ListView

Public Title             As String

Private ReportDate       As Date
Private ReportType       As String
Private alarmtypes       As String
Private mReportType      As Integer
Private mStartDate       As Date
Private mEndDate         As Date
Private mCriteria        As String

Const COL_INV_SERIAL = 0
Const COL_INV_MODEL = 1
Const COL_INV_LAST = 2
Const COL_INV_FIRST = 4.25
Const COL_INV_ROOM = 6.75



Const COL_SERIAL = 0
Const COL_ROOM = 0.75
Const COL_RESIDENT = 1.5
Const COL_EVENT = 3.5
Const COL_PHONE = 3.5
Const COL_ANNOUNCE = 4.5
Const COL_NOTES = 5
Const COL_ALARMTIME = 8.5
Const COL_ELAPSEDTIME = 9.5
Const COL_OPERATOR = 9.75

Const COL_DEVICETYPE = 0
Const COL_NUM_CALLS = 3.5


'Const COL_SERIAL = 0
'Const COL_ROOM = 0.75
'Const COL_RESIDENT = 1.5
'Const COL_EVENT = 3.5
'Const COL_PHONE = 3.5
'Const COL_ANNOUNCE = 4.5
'Const COL_NOTES = 5
'Const COL_ALARMTIME = 8.5
'Const COL_ELAPSEDTIME = 9.5
'Const COL_OPERATOR = 9.75

Private Function PrintAverage()
  Dim li                 As ListItem

  Dim Average            As String

  Select Case mReportType
    Case RPT_EXCEPTION         '  RPT_COUNT ??
      If mNumEntries > 0 Then
        Average = mSumElapsed / mNumEntries
      End If
      If Val(Average) > 0 Then
        Average = Format(DateAdd("s", CDbl(Average), 0), "hh:nn:ss")
      Else
        Average = "N/A"
      End If


      Select Case Dest
        Case DEST_SCREEN
          If Not lv Is Nothing Then
            Set li = lv.ListItems.Add(, , "")
            li.SubItems(1) = "Average"
            li.SubItems(6) = Average
          End If
        Case DEST_PRINTER

          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Average"
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, Average
          Printer.Print " "

        Case DEST_FILE
          If hfile <> 0 Then
            Print #hfile, Join(Array("Average", "", "", "", "", "", "", "", "", Average, ""), vbTab)
          End If
      End Select
  End Select

End Function

Public Function FillAdvanced(ByVal ReportType As Integer, ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date)
10      mReportType = ReportType
20      mStartDate = StartDate
30      mEndDate = EndDate
40      mCriteria = Criteria
        Dim SQL                As String
50      Set rows = New Collection

60      ReportDate = StartDate


70      Select Case ReportType

          Case RPT_INV
80          Title = "Inventory Report for " & Format(Now, "mm/dd/yy")
90          FindInv Criteria
100         PrintReport


110       Case RPT_EXCEPTION
120         Title = "Exception Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
130         FindException Criteria, StartDate, EndDate  ' kinda of like
140         PrintReport

150       Case RPT_COUNT
160         Title = "Event Count Report " & Format(StartDate, "mm/dd/yy") & " to " & Format(EndDate, "mm/dd/yy")
170         FindCount Criteria, StartDate, EndDate
180         PrintReport
190       Case Else
            ' no data
200         PrintReport
210     End Select



End Function
Function FindCount(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL               As String
        Dim rsEvents          As ADODB.Recordset
        Dim rsReport          As ADODB.Recordset
        'Dim AlarmHistory  As cAlarmHistory
        Dim WhereClause       As String
        Dim ExceptionLimit    As Double
        Dim devdesc           As String
        Dim params()          As String
        Dim counter           As Long
        Dim models            As String
        Dim ExceptionReport   As cExceptionReport
        
        Dim ESDeviceType      As ESDeviceTypeType
        
        Dim CLSPTI            As Long
        Dim Model             As String
        
        Dim CountData         As cCountData
        Dim modelList         As Collection
        
        'Dim Dev               As cESDevice
        
        Dim j As Long
        
        Dim CountTotal        As cCountData
        Set CountTotal = New cCountData
        
10      Set modelList = New Collection
        

20      params = Split(Criteria, ",")

30      Set ExceptionReport = New cExceptionReport

40      SQL = "SELECT * FROM ExceptionReports WHERE ReportID = " & Val(params(0))

50      Set rsReport = ConnExecute(SQL)

60      If Not (rsReport.EOF) Then
70        ExceptionReport.Parse rsReport
          
80        models = ExceptionReport.DevString
90        ExceptionLimit = ExceptionReport.ResponseTime   ' not used
100       EndDate = DateAdd("d", 1, EndDate)


110      SQL = " SELECT Rooms.Room, Alarms.id, Alarms.serial,alarms.eventdate, Alarms.alarm, Alarms.tamper,alarms.battery, " & _
                " alarms.residentid, alarms.roomid, alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst ,Rooms.Room , Devices.Model  " & _
                " FROM Alarms " & _
                " LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID " & _
                " LEFT JOIN Devices ON Alarms.Serial = Devices.Serial " & _
                " LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
                " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) "

120       Set rsEvents = ConnExecute(SQL)
          
125       rsEvents.Filter = "eventtype = " & EVT_EMERGENCY & " OR eventtype = " & EVT_ALERT & " Or EventType = " & EVT_EXTERN


130       Do Until rsEvents.EOF
            
140         'Set Dev = Nothing
150         'Set Dev = Devices.device(rsEvents("Serial") & "")
160         If 1 Then
170           'devdesc = Dev.Description
              Model = rsEvents("Model") & ""
              
              ESDeviceType = GetESDeviceTypeByModel(Model)
              CLSPTI = ESDeviceType.CLSPTI
                
                 

              
'180           CLSPTI = Val(GetCLSPTI(Model))  ' 45571 for door/window
190           If CLSPTI Then
200             If InStr(1, models, CStr(CLSPTI), vbTextCompare) Then
                              
210               counter = counter + 1
220               If counter Mod nDoEvents = 0 Then  ' just yield every n hits
230                 DoEvents
240               End If
                  
250               Set CountData = Nothing
260               On Error Resume Next
270               Set CountData = modelList(Model)
280               On Error GoTo 0
290               If CountData Is Nothing Then
300                 Set CountData = New cCountData
310                 CountData.Model = Model
320                 CountData.Description = ESDeviceType.desc
330                 CountData.CLSPTI = ESDeviceType.CLSPTI
340                 CountData.Increment
350                 modelList.Add CountData, Model
                    
360               Else
370                 CountData.Increment
380               End If

390             End If
400           End If
410         End If
420         rsEvents.MoveNext

430       Loop
440       rsEvents.Close
450       Set rsEvents = Nothing
460     End If
470     rsReport.Close
480     Set rsReport = Nothing

490     For Each CountData In modelList
500       rows.Add CountData
          CountTotal.Count = CountTotal.Count + CountData.Count
510     Next
        CountTotal.Description = "Total"
        rows.Add CountTotal

End Function

Function FindInv(ByVal Criteria As String)
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim rsReport           As ADODB.Recordset
        Dim params()           As String
        Dim InventoryReport    As cExceptionReport
        Dim DeviceTypes        As String
        Dim models             As String
        Dim modelList()        As String
        Dim row                As cAlarmHistory


        Dim PTIList()          As String
        Dim PTIString          As String

        Dim RequestedModelString As String

        Dim DEVList()          As ModelPTI

        Dim j                  As Long
        Dim i                  As Long

        Dim PTIs               As String

        Dim rsdev              As ADODB.Recordset


10      On Error GoTo FindInv_Error
20      Set rows = New Collection

30      ReDim DEVList(0)
40      DEVList(0).CLSPTI = 0
50      SQL = "SELECT DISTINCT model FROM Devices"
60      Set rs = ConnExecute(SQL)
        ' we now should have a list of models actually used

        Dim ESDeviceType As ESDeviceTypeType

70      Do Until rs.EOF
80        i = i + 1
90        ReDim Preserve DEVList(i) As ModelPTI
95        ESDeviceType = GetESDeviceTypeByModel(rs("model") & "")
100       DEVList(i).Model = ESDeviceType.Model
110       DEVList(i).CLSPTI = ESDeviceType.CLSPTI
            
'    For i = 0 To MAX_ESDEVICETYPES
'      On Error Resume Next
'      If 0 = StrComp(ESDeviceType(i).Model, Model, vbTextCompare) Then
'        AddToListBox lstEventTypes, ESDeviceType(i).desc & " (" & ESDeviceType(i).Model & ")", ESDeviceType(i).CLSPTI
'        Exit For
'      End If
'    Next
  
            
120       rs.MoveNext
130     Loop


140     rs.Close
150     Set rs = Nothing


160     Set InventoryReport = New cExceptionReport
170     params = Split(Criteria, ",")
180     SQL = "SELECT * FROM ExceptionReports WHERE ReportID = " & Val(params(0))
190     Set rsReport = ConnExecute(SQL)
200     If Not (rsReport.EOF) Then

210       InventoryReport.Parse rsReport

220       PTIList() = Split(InventoryReport.DevString, vbTab)

230       ReDim RequestedModels(0)
          Dim n                As Long

240       For j = 1 To UBound(DEVList)
250         For i = LBound(PTIList) To UBound(PTIList)
260           If DEVList(j).CLSPTI = Val(PTIList(i)) Then
270             ReDim Preserve RequestedModels(n)
280             RequestedModels(n) = q(DEVList(j).Model)
290             n = n + 1
300             Exit For
310           End If
320         Next
330       Next

340       RequestedModelString = Join(RequestedModels, ",")



350       SQL = "SELECT Devices.Serial, devices.model, Rooms.Room, Residents.namelast, residents.namefirst FROM "
360       SQL = SQL & " ( Devices  LEFT JOIN residents ON Devices.residentid =  Residents.ResidentID )"
370       SQL = SQL & " LEFT JOIN Rooms ON devices.roomid = Rooms.roomid "
          SQL = SQL & " WHERE devices.model in ( " & RequestedModelString & ")"
380       SQL = SQL & " ORDER BY Rooms.room, devices.model"

390       Set rs = ConnExecute(SQL)
          '    Dim hfile As Long
          '    hfile = FreeFile





400       'Open App.path & "\testdata.txt" For Output As hfile
410       'Print #hfile, "Serial,Model,Last,First,Room"
420       Do Until rs.EOF
430         Set row = New cAlarmHistory
440         row.Serial = rs("serial") & ""
450         'row.DeviceType = rs("model") & ""
            
            ESDeviceType = GetESDeviceTypeByModel(rs("model") & "")
            row.DeviceType = ESDeviceType.desc & " (" & ESDeviceType.Model & ")"

460         row.NameFirst = rs("namefirst") & ""
470         row.NameLast = rs("namelast") & ""
480         row.Room = rs("room") & ""
490         rows.Add row             'Print #hfile, rs("serial") & "," & rs("Model") & "," & rs("namelast") & "," & rs("namefirst") & "," & rs("room")
500         rs.MoveNext
510       Loop

530       rs.Close


540     End If
550     rsReport.Close

FindInv_Resume:

560     Set rs = Nothing
570     Set rsReport = Nothing

580     On Error GoTo 0
590     Exit Function

FindInv_Error:

600     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionReportData.FindInv." & Erl
610     Resume FindInv_Resume

End Function


Function FindException(ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date) As Long
        Dim SQL                As String
        Dim rs                 As ADODB.Recordset
        Dim rsReport           As ADODB.Recordset
        Dim AlarmHistory       As cAlarmHistory
        Dim WhereClause        As String
        Dim ExceptionLimit     As Double
        'Dim Devices()     As String
        Dim params()           As String
        Dim counter            As Long
        Dim CLSPTIs             As String
        Dim rsd                As ADODB.Recordset
        Dim devdesc            As String

        Dim Dev                As cESDevice

        Dim CLSPTI             As Long

        Dim ExceptionReport    As cExceptionReport


        Dim SQL2               As String
        Dim rs2                As ADODB.Recordset
        Dim Whereclause2       As String
        Dim rc                 As Long



10      On Error GoTo FindException_Error

20      params = Split(Criteria, ",")

30      Set ExceptionReport = New cExceptionReport

40      SQL = "SELECT * FROM ExceptionReports WHERE ReportID = " & Val(params(0))

50      Set rsReport = ConnExecute(SQL)

60      If Not (rsReport.EOF) Then
70        ExceptionReport.Parse rsReport
          'ExceptionLimit = ExceptionReport.ResponseTime
80        EndDate = DateAdd("d", 1, EndDate)
90        CLSPTIs = ExceptionReport.DevString
100       ExceptionLimit = ExceptionReport.ResponseTime * 60



          Dim alarmtypes       As String



110       alarmtypes = EVT_EMERGENCY & "," & EVT_ALERT & "," & EVT_EXTERN

          '110       SQl = " SELECT Rooms.Room, Alarms.id, Alarms.serial,alarms.eventdate, Alarms.alarm, Alarms.tamper,alarms.battery, " & _
           '                " alarms.residentid, alarms.roomid, alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst " & _
           '                " FROM (Alarms LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
           '                " WHERE (alarms.eventtype in (" & alarmtypes & ")) AND ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
           '                " ORDER BY Alarms.EventDate "

120       SQL = " SELECT Rooms.Room, Alarms.id, Alarms.serial,alarms.eventdate, Alarms.alarm, Alarms.tamper,alarms.battery, " & _
              " alarms.residentid, alarms.roomid, alarms.announce, alarms.eventtype, alarms.username, alarms.alarmid , Residents.NameLast, Residents.NameFirst ,Rooms.Room , Devices.Model  " & _
              " FROM Alarms " & _
              " LEFT JOIN Residents ON Alarms.ResidentID = Residents.ResidentID " & _
              " LEFT JOIN Devices ON Alarms.Serial = Devices.Serial " & _
              " LEFT JOIN Rooms ON Alarms.roomID = Rooms.RoomID " & _
              " WHERE ((alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") AND (alarms.AlarmID = 0)) " & _
              " ORDER BY Alarms.id "


130       Set rs = ConnExecute(SQL)

140       rs.Filter = "eventtype = " & EVT_EMERGENCY & " OR eventtype = " & EVT_ALERT & " Or EventType = " & EVT_EXTERN


150       Whereclause2 = "  WHERE (alarms.EventDate >= " & DateDelim & StartDate & DateDelim & ") And (alarms.EventDate < " & DateDelim & EndDate & DateDelim & ") "

160       SQL2 = " SELECT Alarms.ID, Alarms.Serial, Alarms_1.AlarmID as alarmid, Alarms_1.EventDate as EventDate, Alarms_1.EventType as EventType, Alarms_1.info as info, Alarms_1.UserName as UserName " & _
               " FROM Alarms INNER JOIN Alarms AS Alarms_1 ON Alarms.ID = Alarms_1.AlarmID " & _
                 Whereclause2 & _
               " ORDER BY Alarms.ID; "

170       Set rs2 = ConnExecute(SQL2)

180       Do Until rs.EOF

190
200
            ' at this point should have the device from the database

210         If 1 Then
220           'devdesc = Dev.Description
              
230           CLSPTI = Val(GetCLSPTI(rs("Model") & ""))   ' 45571 for door/window
240           If CLSPTI Then
250             If InStr(1, CLSPTIs, CStr(CLSPTI), vbTextCompare) Then
260               counter = counter + 1
270               If counter > 200 Then  ' just yield every n hits: DoEventYield is 25 by defaukt
280                 counter = 0
290                 DoEvents
300               End If

                  If CLSPTI <> &HEEEE Then

                  'Set AlarmHistory = GetAlarmHistory(rs)
310               Set AlarmHistory = New cAlarmHistory

320               AlarmHistory.ID = rs("ID")  ' this record's ID
330               AlarmHistory.Serial = rs("Serial") & ""
340               AlarmHistory.EventDate = rs("EventDate")
350               AlarmHistory.alarm = rs("Alarm")
360               AlarmHistory.Tamper = rs("Tamper")
370               AlarmHistory.Battery = rs("Battery")
380               AlarmHistory.Announce = rs("Announce") & ""
390               AlarmHistory.EventType = rs("EventType")
400               AlarmHistory.User = rs("Username") & ""
410               AlarmHistory.AlarmID = rs("AlarmID")  ' parent alarm

420               If Len(rs("NameLast") & "") Then
430                 If Len(rs("NameFirst") & "") Then
440                   AlarmHistory.Resident = rs("NameLast") & ", " & rs("NameFirst")
450                 Else
460                   AlarmHistory.Resident = rs("NameLast") & ""
470                 End If
480               ElseIf Len(rs("NameFirst") & "") Then
490                 AlarmHistory.Resident = rs("NameFirst") & ""
500               End If
510               AlarmHistory.Room = rs("room") & ""

520               rc = AlarmHistory.GetHistoryNew(AlarmHistory.ID, rs2)

530               If AlarmHistory.ElapsedTime > ExceptionLimit Then
540                 rows.Add AlarmHistory
550               End If
                  End If
560             End If
570           End If
580         End If
590         rs.MoveNext
600       Loop
610       rs.Close
620       rs2.Close
630     End If
640     rsReport.Close
650     Set rsReport = Nothing

FindException_Resume:

660     Set rs = Nothing
670     Set rs2 = Nothing
680     Set rsReport = Nothing
690     On Error GoTo 0
700     Exit Function

FindException_Error:

710     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionReportData.FindException." & Erl
720     Resume FindException_Resume

End Function



Sub Fill(ByVal ReportID As String)
'  Select Case reportID
'    Case "currentshift"
'      ReportType = "S"
'      ShiftReport
'      PrintReport
'    Case "today"
'      ReportType = "D"
'      DailyReport
'      PrintReport
'    Case "thisweek"
'      ReportType = "W"
'      WeeklyReport
'      PrintReport
'    Case "thismonth"
'      ReportType = "M"
'      MonthlyReport
'      PrintReport
'    Case Else
'      ReportType = ""
'      DailyReport
'      PrintReport
'  End Select
End Sub

Sub PrintReport()
        Dim AlarmHistory  As Object ' late bound
        Dim filename      As String
        
10      On Error GoTo PrintReport_Error

20      Select Case Dest
          Case DEST_SCREEN  ' if To Screen ' Just add to listview
30          lv.ListItems.Clear
40          lv.Refresh
50        Case DEST_FILE  ' if To File ' one continuous file
60          hfile = FreeFile
70          filename = Replace(Title, "/", "-")
80          Open Configuration.ReportPath & CleanFileName(filename) & ".txt" For Output As hfile

90        Case DEST_PRINTER  ' if To Printer
100         Printer.Orientation = vbPRORLandscape
110         Printer.Print " ";
120         Printer.ScaleMode = vbInches
130         Printer.FontSize = 10
140         Printer.FontBold = False

150     End Select

160     PrintBanner
170     PrintRowHeader
180     If Dest = DEST_SCREEN Then
190        Win32.LockWindowUpdate lv.hwnd
200     End If

210       For Each AlarmHistory In rows

230       Select Case Dest
            Case DEST_SCREEN  ' if To Screen ' Just add to listview
240         Case DEST_PRINTER  ' if To Printer
              ' check pagelength
250           If (Printer.ScaleHeight - Printer.CurrentY) < 0.5 Then  '  half inch bottom margin
260             PrintFooter
270             FormFeed Printer
280             PrintRowHeader
290           End If
300         Case DEST_FILE  ' if To File ' one continuous file
310       End Select

320       PrintRow AlarmHistory
330     Next

340     Select Case Dest
          Case DEST_SCREEN  ' if To Screen ' Just add to listview
           PrintAverage
350        Win32.LockWindowUpdate 0

          
360       Case DEST_PRINTER  ' if To Printer
           PrintAverage
370         PrintFooter
380         Printer.EndDoc
390         Printer.Orientation = vbPRORPortrait
400       Case DEST_FILE  ' if To File ' one continuous file
            PrintAverage
410         If hfile <> 0 Then
420           Close hfile
430           hfile = 0
440         End If
450     End Select

PrintReport_Resume:
460     On Error GoTo 0
470     Exit Sub

PrintReport_Error:

480     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cExceptionReportData.PrintReport." & Erl
490     Resume PrintReport_Resume


End Sub

Sub PrintCountRow(CountData As cCountData)

End Sub

Sub PrintRow(AlarmHistory As Object)
' ID Room Resident Eventtype Alarmtime Canceltime ElapsedTime Username
' tab delimited
  Dim li As ListItem
  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
      If Not lv Is Nothing Then
        Select Case mReportType
          Case RPT_INV
            'Serial, Device Type, Resident First Name, Resident Last Name and Room Description
            
            Set li = lv.ListItems.Add(, , AlarmHistory.Serial)
            li.SubItems(1) = AlarmHistory.DeviceType
            li.SubItems(2) = AlarmHistory.NameLast
            li.SubItems(3) = AlarmHistory.NameFirst
            li.SubItems(4) = AlarmHistory.Room
            
            
          Case RPT_EXCEPTION
            
            Set li = lv.ListItems.Add(, , AlarmHistory.Serial)

            li.SubItems(1) = AlarmHistory.Room
            li.SubItems(2) = AlarmHistory.Resident
            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT ' change 2/15/2012
                li.SubItems(3) = AlarmHistory.Announce
              Case Else
                li.SubItems(3) = AlarmHistory.EventName
            End Select
            
            'AlarmHistory.Announce
            'li.SubItems(3) = AlarmHistory.EventName
            li.SubItems(4) = AlarmHistory.AlarmTimeToString
            
            
            
            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                If gElapsedEqACK = USE_ELAPSED_ACK Then

                  If AlarmHistory.ACKElapsed > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                  End If
                Else
                  If AlarmHistory.ElapsedTime > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                  End If
                End If
            End Select         'Case AlarmHistory.Eventtype
            
            
            If gElapsedEqACK = USE_ELAPSED_ACK Then
              If AlarmHistory.ACKTime > 0 Then
                li.SubItems(5) = AlarmHistory.ACKTimeToString
                li.SubItems(6) = AlarmHistory.ACKElapsedTimeToString
                
              Else  ' try and use end of alarm time
                li.SubItems(5) = AlarmHistory.CancelTimeToString
                li.SubItems(6) = AlarmHistory.ElapsedTimeToString
              End If
            Else
              li.SubItems(5) = AlarmHistory.CancelTimeToString
              li.SubItems(6) = AlarmHistory.ElapsedTimeToString
            End If
            li.SubItems(7) = AlarmHistory.User
          
          
          Case RPT_COUNT

            Set li = lv.ListItems.Add(, , AlarmHistory.CLSPTI)

            li.SubItems(1) = AlarmHistory.Model
            li.SubItems(2) = AlarmHistory.Description
            li.SubItems(3) = AlarmHistory.Count
            'li.SubItems(4) = AlarmHistory.AveragePerDay
          Case Else

            Set li = lv.ListItems.Add(, , AlarmHistory.Serial)
            li.SubItems(1) = AlarmHistory.Room
            li.SubItems(2) = AlarmHistory.Resident
            li.SubItems(3) = AlarmHistory.EventName
            
            li.SubItems(4) = AlarmHistory.AlarmTimeToString

            If gElapsedEqACK = USE_ELAPSED_ACK Then
              If AlarmHistory.ACKTime > 0 Then
                li.SubItems(5) = AlarmHistory.ACKTimeToString
                li.SubItems(6) = AlarmHistory.ACKElapsedTimeToString
              Else  ' try and use end of alarm time
                li.SubItems(5) = AlarmHistory.CancelTimeToString
                li.SubItems(6) = AlarmHistory.ElapsedTimeToString
              End If
            Else
              li.SubItems(5) = AlarmHistory.CancelTimeToString
              li.SubItems(6) = AlarmHistory.ElapsedTimeToString
            End If
            '            li.SubItems(5) = AlarmHistory.CancelTimeToString
            '            li.SubItems(6) = AlarmHistory.ElapsedTimeToString
            li.SubItems(7) = AlarmHistory.User

            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                If gElapsedEqACK = USE_ELAPSED_ACK Then

                  If AlarmHistory.ACKElapsed > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                  End If
                Else
                  If AlarmHistory.ElapsedTime > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                  End If
                End If
            End Select         'Case AlarmHistory.Eventtype


        End Select
      End If
    Case DEST_PRINTER  ' if To Printer:
      Select Case mReportType
        Case RPT_INV
          'Serial, Device Type, Resident First Name, Resident Last Name and Room Description
          PrintLJ Printer, COL_INV_SERIAL, Printer.CurrentY, AlarmHistory.Serial & " "
          PrintLJ Printer, COL_INV_MODEL, Printer.CurrentY, AlarmHistory.DeviceType & " "
          PrintLJ Printer, COL_INV_LAST, Printer.CurrentY, AlarmHistory.NameLast & " "
          PrintLJ Printer, COL_INV_FIRST, Printer.CurrentY, AlarmHistory.NameFirst & " "
          PrintLJ Printer, COL_INV_ROOM, Printer.CurrentY, AlarmHistory.Room & " "
          Printer.Print " "
        Case RPT_EXCEPTION
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial & " "  ' need apostrophy for excel to like it
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, AlarmHistory.Announce
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString
          If gElapsedEqACK = USE_ELAPSED_ACK Then
            If AlarmHistory.ACKTime > 0 Then
              PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ACKElapsedTimeToString
            Else  ' try and use end of alarm time
              PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
            End If
          Else
            PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
          End If
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
          Printer.Print " "
        
            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                If gElapsedEqACK = USE_ELAPSED_ACK Then

                  If AlarmHistory.ACKElapsed > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                  End If
                Else
                  If AlarmHistory.ElapsedTime > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                  End If
                End If
            End Select         'Case AlarmHistory.Eventtype
        
        
        Case RPT_COUNT
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Model & " "  ' need apostrophy for excel to like it
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.Description
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.Count
          Printer.Print " "
        Case Else
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, AlarmHistory.Serial
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, AlarmHistory.Room
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, AlarmHistory.Resident
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, AlarmHistory.EventName
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, AlarmHistory.Announce
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, AlarmHistory.AlarmTimeToString
          If gElapsedEqACK = 1 Then
            If AlarmHistory.ACKTime > 0 Then

              PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ACKElapsedTimeToString
            Else  ' try and use end of alarm time

              PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
            End If
          Else

            PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, AlarmHistory.ElapsedTimeToString
          End If
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, AlarmHistory.User
          Printer.Print " "
            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                If gElapsedEqACK = USE_ELAPSED_ACK Then

                  If AlarmHistory.ACKElapsed > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                  End If
                Else
                  If AlarmHistory.ElapsedTime > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                  End If
                End If
            End Select         'Case AlarmHistory.Eventtype


      End Select

    Case DEST_FILE  ' if To File
      If hfile <> 0 Then
        ' need to modify this !
        Select Case mReportType
          Case RPT_INV
            'Serial, Device Type, Resident First Name, Resident Last Name and Room Description
             Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.DeviceType, AlarmHistory.NameLast, AlarmHistory.NameFirst, AlarmHistory.Room), vbTab)
          Case RPT_EXCEPTION
            'Print #hFile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.user), vbTab)
            Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, _
                                     AlarmHistory.ACKTimeToString, AlarmHistory.ACKElapsedTimeToString, _
                                     AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.User), vbTab)
          
            Select Case AlarmHistory.EventType
              Case EVT_EMERGENCY, EVT_ALERT, EVT_EXTERN
                If gElapsedEqACK = USE_ELAPSED_ACK Then

                  If AlarmHistory.ACKElapsed > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ACKElapsed
                  End If
                Else
                  If AlarmHistory.ElapsedTime > 0 Then
                    mNumEntries = mNumEntries + 1
                    mSumElapsed = mSumElapsed + AlarmHistory.ElapsedTime
                  End If
                End If
            End Select         'Case AlarmHistory.Eventtype
          
          
          Case RPT_COUNT
            
            ' not used Set li = lv.ListItems.Add(, , AlarmHistory.CLSPTI)
            Print #hfile, Join(Array(AlarmHistory.Model, AlarmHistory.Description, AlarmHistory.Count), vbTab)
          Case Else

            Print #hfile, Join(Array(AlarmHistory.Serial, AlarmHistory.Room, AlarmHistory.Resident, AlarmHistory.EventName, AlarmHistory.Announce, AlarmHistory.AlarmTimeToString, _
                                     AlarmHistory.ACKTimeToString, AlarmHistory.ACKElapsedTimeToString, _
                                     AlarmHistory.CancelTimeToString, AlarmHistory.ElapsedTimeToString, AlarmHistory.User), vbTab)
        
        End Select

      End If
  End Select

End Sub

Sub PrintRowHeader()
' ID Room Resident Eventtype Alarmtime Canceltime ElapsedTime Username
' tab delimited

  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
      ConfigureListView


    Case DEST_PRINTER  ' if To Printer:
      Select Case mReportType
        Case RPT_INV
        'Serial, Device Type, Resident Last Name, Resident First Name and Room Description
          PrintLJ Printer, COL_INV_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_INV_MODEL, Printer.CurrentY, "Model"
          PrintLJ Printer, COL_INV_LAST, Printer.CurrentY, "Last"
          PrintLJ Printer, COL_INV_FIRST, Printer.CurrentY, "First"
          PrintLJ Printer, COL_INV_ROOM, Printer.CurrentY, "Room"
          Printer.Print ""
        
        Case RPT_EXCEPTION
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, "Announce"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Alarm Time"
          'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, "Cancel Time"
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "Elapsed Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print ""
          HR Printer
        Case RPT_COUNT
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Model"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Description"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Count"
          Printer.Print ""
          HR Printer
        
        Case Else
          PrintLJ Printer, COL_SERIAL, Printer.CurrentY, "Serial"
          PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
          PrintLJ Printer, COL_RESIDENT, Printer.CurrentY, "Resident"
          PrintLJ Printer, COL_EVENT, Printer.CurrentY, "Event"
          PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, "Announce"
          PrintRJ Printer, COL_ALARMTIME, Printer.CurrentY, "Alarm Time"
          'PrintRJ Printer, COL_CANCELTIME, Printer.CurrentY, "Cancel Time"
          PrintRJ Printer, COL_ELAPSEDTIME, Printer.CurrentY, "Elapsed Time"
          PrintLJ Printer, COL_OPERATOR, Printer.CurrentY, "Operator"
          Printer.Print " "
          HR Printer

      End Select
    Case DEST_FILE  ' if To File
      If hfile <> 0 Then
        
            
        Select Case mReportType
          Case RPT_INV
            Print #hfile, Join(Array("Serial", "Model", "Last", "First", "Room"), vbTab)
          Case RPT_EXCEPTION
            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
                                     "ACK Time", "ACK Elapsed Time", _
                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)
          Case RPT_COUNT
            
            Print #hfile, Join(Array("Device", "Description", "Count"), vbTab)
          Case Else
            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
                                     "ACK Time", "ACK Elapsed Time", _
                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)
        End Select
      End If
    
'    Case DEST_HTML
'      If hfile <> 0 Then
'        Select Case mReportType
'          Case RPT_EXCEPTION
'            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
'                                     "ACK Time", "ACK Elapsed Time", _
'                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)
'          Case RPT_COUNT
'
'            Print #hfile, Join(Array("Device", "Description", "Count"), vbTab)
'          Case Else
'            Print #hfile, Join(Array("Serial", "Room", "Resident", "Event", "Announce", "Alarm Time", _
'                                     "ACK Time", "ACK Elapsed Time", _
'                                     "Cancel Time", "Elapsed Time", "Operator"), vbTab)
'        End Select
'      End If
'
    
    
  End Select


End Sub
Sub ConfigureListView()

'Global Const RPT_ROOM = 1
'Global Const RPT_RES = 2
'Global Const RPT_DEVICE = 3
'Global Const RPT_EVENT = 4
'Global Const RPT_ASSUR = 5
'Global Const RPT_DEVHIST = 6     ' add/remove devices
'Global Const RPT_RESHIST = 7     ' add/remove residents
'Global Const RPT_EXCEPTION = 8   ' exception (alarm ack/clear time)
'Global Const RPT_COUNT = 9       ' number of alarms per device type

  If Not lv Is Nothing Then
    Select Case mReportType
      Case RPT_INV
        ' Serial, Device Type, Resident First Name, Resident Last Name and Room Description
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "Serial", 1200
        lv.ColumnHeaders.Add , , "Model", 1400
        lv.ColumnHeaders.Add , , "Last", 1900
        lv.ColumnHeaders.Add , , "First", 1900
        lv.ColumnHeaders.Add , , "Room", 1900
      
      Case RPT_EXCEPTION
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 0
        lv.ColumnHeaders.Add , , "Room", 1200
        lv.ColumnHeaders.Add , , "Resident", 1600
        lv.ColumnHeaders.Add , , "Event", 1300
        lv.ColumnHeaders.Add , , "Alarm", 1500
        lv.ColumnHeaders.Add , , "Cancel", 1300
        lv.ColumnHeaders.Add , , "Elapsed", 850, lvwColumnRight
        lv.ColumnHeaders.Add , , "Operator", 900
      
      Case RPT_COUNT
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "CLSPTI", 0
        lv.ColumnHeaders.Add , , "Model", 1200
        lv.ColumnHeaders.Add , , "Descr", 3600
        lv.ColumnHeaders.Add , , "Count", 1300, lvwColumnRight
        

      Case Else
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "ID", 0
        lv.ColumnHeaders.Add , , "Room", 1200
        lv.ColumnHeaders.Add , , "Resident", 1600
        lv.ColumnHeaders.Add , , "Event", 800
        lv.ColumnHeaders.Add , , "Alarm", 1500
        lv.ColumnHeaders.Add , , "Cancel", 1300
        lv.ColumnHeaders.Add , , "Elapsed", 850
        lv.ColumnHeaders.Add , , "Operator", 900

    End Select
    lv.Refresh
  End If
End Sub

Sub PrintBanner()
' date and time of report
  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
    Case DEST_PRINTER  ' if To Printer:


      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.ScaleMode = vbInches

      Printer.Print ""
      PrintCentered Printer, Printer.CurrentY, Title
      PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "") ' hh:nn am/pm")

      Printer.Print ""
      Printer.Print ""


    Case DEST_FILE  ' if To File

  End Select
End Sub
Sub PrintFooter()
' copyright etc
  Select Case Dest
    Case DEST_SCREEN  ' if To Screen:
    Case DEST_PRINTER  ' if To Printer:
      Dim text As String

      text = App.LegalCopyright & " " & App.ProductName & " (Ver. 1." & Format(App.Revision, "000") & ")"
      PrintCJ Printer, Printer.ScaleWidth / 2, Printer.ScaleHeight - Printer.TextHeight("A") * 1.5, text

    Case DEST_FILE  ' if To File

  End Select


End Sub

'
'Function GetAlarmHistory(rs As Recordset) As cAlarmHistory
'  Dim History As cAlarmHistory
'  Set History = New cAlarmHistory
'
'  History.Parse rs
'  If History.ResidentID <> 0 Then
'    History.GetResidentInfo History.ResidentID
'  End If
'  History.Room = GetRoomName(History.RoomID)
'  History.GetHistory conn
'
'  Set GetAlarmHistory = History
'End Function


Private Sub Class_Initialize()

  Set rows = New Collection
  alarmtypes = Join(Array(EVT_EMERGENCY, EVT_ALERT, EVT_TAMPER, EVT_BATTERY_FAIL, EVT_CHECKIN_FAIL, EVT_COMM_TIMEOUT), ",")
End Sub

Private Sub Class_Terminate()
  Set rows = Nothing
End Sub






