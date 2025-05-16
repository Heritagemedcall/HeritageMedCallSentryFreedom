Attribute VB_Name = "modAssur"
Option Explicit

Global gAssurStartTime    As Date
Global gAssurEndTime      As Date
Global gAssurStartTime2   As Date
Global gAssurEndTime2     As Date
Global AssureVacationDevices As Collection

Global InAssurPeriod      As Boolean

Global gAssurAutoSend     As Boolean

Global gExportStyle           As Long
Global gExportStyleNoHeaders  As Long

Global Const STYLE_FILETYPE_TAB = 0
Global Const STYLE_FILETYPE_TAB_NOHEADERS = 1
Global Const STYLE_FILETYPE_HTML = 2


Public Function IsAssurActive()
  If MASTER Then
    IsAssurActive = ((Configuration.AssurStart - Configuration.AssurEnd) <> 0) Or ((Configuration.AssurStart2 - Configuration.AssurEnd2) <> 0)
  Else
    IsAssurActive = RemoteIsAssurActive()
  End If

End Function

Function AutoSendAssur(Items As Collection)
        Dim AssurText As String
        Dim VacText As String
        Dim AssurFilename As String
        Dim AssurVacFilename As String

        'SortVacationItems AssureVacationDevices

        Dim VacationItems As Collection
        Dim j As Long
        Dim InClause As String
        Dim SQl As String
        Dim rs As ADODB.Recordset
        Dim AssurListSerial() As String
        '
        Dim Item As cESDevice
        '
10      On Error GoTo AutoSendAssur_Error

20      Set VacationItems = New Collection
30      If Not AssureVacationDevices Is Nothing Then
40        If AssureVacationDevices.Count Then
50          ReDim AssurListSerial(1 To AssureVacationDevices.Count)

60          For j = 1 To AssureVacationDevices.Count
70            AssurListSerial(j) = "'" & AssureVacationDevices(j).Serial & "'"
80          Next
90          InClause = Join(AssurListSerial, ",")


100         SQl = "SELECT Devices.Serial, Devices.DeviceID, Residents.NameLast, Rooms.Room, Residents.NameFirst,  Residents.phone, Residents.RoomID, Rooms.RoomID, Devices.RoomID, Devices.ResidentID " & _
                " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
                " WHERE Devices.residentid <> 0 AND (Devices.Serial In (" & InClause & ")) " & _
                " ORDER BY Residents.NameLast, Residents.NameFirst, Rooms.Room; "

110         Set rs = ConnExecute(SQl)
120         Do Until rs.EOF
130           Set Item = New cESDevice
140           Item.Serial = rs("serial") & ""
150           Item.NameLast = rs("namelast") & ""
160           Item.NameFirst = rs("namefirst") & ""
170           Item.Phone = rs("phone") & ""
180           Item.Room = rs("room") & ""
190           VacationItems.Add Item
200           rs.MoveNext
210         Loop
220         rs.Close


230         SQl = "SELECT Devices.Serial, Devices.DeviceID, Residents.NameLast, Rooms.Room, Residents.NameFirst,  Residents.phone, Residents.RoomID, Rooms.RoomID, Devices.RoomID, Devices.ResidentID " & _
                " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
                " WHERE Devices.residentid = 0 AND (Devices.Serial In (" & InClause & ")) " & _
                " ORDER BY Rooms.Room; "

240         Set rs = ConnExecute(SQl)
250         Do Until rs.EOF
260           Set Item = New cESDevice
270           Item.Serial = rs("serial") & ""
280           Item.NameLast = rs("namelast") & ""
290           Item.NameFirst = rs("namefirst") & ""
300           Item.Phone = rs("phone") & ""
310           Item.Room = rs("room") & ""
320           VacationItems.Add Item
330           rs.MoveNext
340         Loop
350         rs.Close

360         Set rs = Nothing
370       End If
380     End If

390     Set AssureVacationDevices = VacationItems

400     Select Case Configuration.AssurFileFormat

        Case STYLE_FILETYPE_HTML
410       AssurText = GenTable_HTML(Items)
420       VacText = GenVacTable_HTML()
430     Case STYLE_FILETYPE_TAB_NOHEADERS
440       AssurText = GenTable_TSV(Items, True)
450       VacText = GenVacTable_TSV(True)
460     Case Else  'STYLE_FILETYPE_TAB
470       AssurText = GenTable_TSV(Items, False)
480       VacText = GenVacTable_TSV(False)
490     End Select

500     AssurFilename = SaveAssur(AssurText)
510     AssurVacFilename = SaveVac(VacText)
520     If Configuration.AssurSendAsEmail Then
530       EmailAssur AssurFilename & ";" & AssurVacFilename
540     End If





AutoSendAssur_Resume:

550     On Error GoTo 0
560     Exit Function

AutoSendAssur_Error:

570     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.AutoSendAssur." & Erl
580     Resume AutoSendAssur_Resume

End Function

Function SaveAssur(ByVal text As String) As String

  Dim hfile As Integer
  Dim filename As String
  Dim NeededFolder As String

  NeededFolder = App.path
  If Right$(NeededFolder, 1) <> "\" Then
    NeededFolder = NeededFolder & "\"
  End If
  NeededFolder = NeededFolder & "Assur"

  If Not (DirExists(NeededFolder)) Then
    MkDir NeededFolder
  End If
  hfile = FreeFile
  filename = NeededFolder & "\AssurRpt-" & Format$(Now, "YYYYMMDD-HH") & IIf(Configuration.AssurFileFormat = 2, ".html", ".txt")
  If FileExists(filename) Then
    On Error Resume Next
    Kill filename
    On Error GoTo 0
  End If
  Open filename For Output As #hfile
  Print #hfile, text
  Close hfile
  SaveAssur = filename


End Function



Function SaveVac(ByVal text As String) As String

  Dim hfile As Integer
  Dim filename As String
  Dim NeededFolder As String

  NeededFolder = App.path
  If Right$(NeededFolder, 1) <> "\" Then
    NeededFolder = NeededFolder & "\"
  End If
  NeededFolder = NeededFolder & "Assur"

  If Not (DirExists(NeededFolder)) Then
    MkDir NeededFolder
  End If
  hfile = FreeFile
  filename = NeededFolder & "\AssurVac-" & Format$(Now, "YYYYMMDD-HH") & IIf(Configuration.AssurFileFormat = 2, ".html", ".txt")
  If FileExists(filename) Then
    On Error Resume Next
    Kill filename
    On Error GoTo 0
  End If
  Open filename For Output As #hfile
  Print #hfile, text
  Close hfile
  SaveVac = filename


End Function


Function EmailAssur(ByVal Filenames As String)


        Dim mapi    As Object
        'Dim mapi As SENTRYMAIL.MAPITransport

        Dim SMPTMAILER As SMTPMailer.SendMail

        Dim Subject As String
        Dim message As String

10      message = "Check-in Report Attached: " & Filenames



20      On Error Resume Next
30      If (Configuration.UseSMTP = MAIL_SMTP) Then
40        If gSMTPMailer Is Nothing Then
'50          Set gSMTPMailer = New SendMail
'            Set SMPTMAILER = New SMTPMailer.SendMail
'            Set gSMTPMailer = gSMTPMailer
50          Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
60        End If

70        If Err.Number Then
80          Set gSMTPMailer = Nothing
'90          Set gSMTPMailer = New SendMail
90           Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
100         Err.Clear
110       End If

120       If gSMTPMailer Is Nothing Then
130         LogProgramError "Could not create SMTPMailer Object in CpageDevice.SENDMAPI2." & Erl
            
140       Else
150         Call gSMTPMailer.Send("", "", Configuration.AssurEmailRecipient, Configuration.AssurEmailSubject, message, Filenames)
160         If Err.Number Then
170           Err.Clear
180           Set gSMTPMailer = Nothing
'190           Set gSMTPMailer = New SendMail
190           Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
200           Call gSMTPMailer.Send("", "", Configuration.AssurEmailRecipient, Configuration.AssurEmailSubject, message, Filenames)
210         End If


220       End If

230     Else
240       Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
250       If mapi Is Nothing Then
260         LogProgramError "Could not create SENTRYMAIL Object in modAssurEmailAssur." & Erl
270       Else
280         Call mapi.SendWithAttachments("", "", Configuration.AssurEmailRecipient, Configuration.AssurEmailSubject, message, Filenames)
290       End If

300     End If


        '// Username, Password, Address, Subject,Body, AttachmentsList ' Attachemnet list is a semicolon ";" delimited list of file attachments

310     Set mapi = Nothing



End Function


Function GenTable_TSV(Items As Collection, ByVal noheaders As Boolean) As String
  Dim Item    As cAssureItem
  Dim text    As String
  Dim row     As String

  If (False = noheaders) Then
    text = Join(Array("Resident", "Phone", "Room", "Device"), vbTab) & vbCrLf
  End If

  For Each Item In Items
    text = text + Join(Array(Item.NameFull, Item.Phone, Item.Room, Item.Serial), vbTab) & vbCrLf
  Next
  GenTable_TSV = text



End Function

Function GenVacTable_TSV(ByVal noheaders As Boolean) As String
  Dim Item    As cESDevice
  Dim text    As String
  Dim row     As String
  Dim NameFull As String

  If (False = noheaders) Then
    text = Join(Array("Resident", "Phone", "Room", "Device"), vbTab) & vbCrLf
  End If

  For Each Item In AssureVacationDevices
    NameFull = ConvertLastFirst(Item.NameLast, Item.NameFirst)
    text = text + Join(Array(NameFull, Item.Phone, Item.Room, Item.Serial), vbTab) & vbCrLf
  Next
  GenVacTable_TSV = text



End Function
Function GenVacTable_HTML() As String
  Dim Item    As cESDevice
  Dim text    As String
  Dim row     As String
  Dim odd     As Boolean
  Dim NameFull As String
  

  text = text + "<html>"
  text = text + "<head>"
  text = text + "<style type=""text/css"">"
  text = text + "body {width:900px; font-family:arial,verdana,sans-serif;}"
  text = text + "table.main {width:900px;font-size:1.0em;}"
  text = text + "tr.header td {background-color: #ADD8E6; color: black; margin:0px; padding:2px; font-weight:bold;}"
  text = text + "tr.even td {background-color: #FAFAD2; color: black; margin:0px; padding:2px}"
  text = text + "tr.odd td {background-color: white; color: black; margin:0px; padding:2px}"
  text = text + "h1 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.3em}"
  text = text + "h2 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.0em;}"
  text = text + "p.complete {background-color: white; color:gray;margin:5px;text-align:left;font-size:0.9em;}"
  
  text = text + "</style>"
  text = text + "</head>"
  text = text + "<body>"
  
  text = text + "<h1>Vacation Report</h1>"
  text = text + "<h1>" & HTMLEncode(Configuration.Facility) & "</h1>"
  text = text + "<h2>" & Format$(Now, "mmmm dd, yy hh:nn AMPM") & "</h2>"
  text = text + "<br />"
  text = text + "<table class='main'>" & vbCrLf
  text = text + "<tr class='header'>" & vbCrLf
  text = text + "<td>Resident</td> <td>Phone</td> <td>Room</td> <td>Device</td>" & vbCrLf
  text = text + "</tr>" & vbCrLf

  For Each Item In AssureVacationDevices
    NameFull = ConvertLastFirst(Item.NameLast, Item.NameFirst)
    text = text + "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
    text = text + "<td>" & HTMLEncode(NameFull) & "</td><td>" & HTMLEncode(Item.Phone) & "</td><td>" & HTMLEncode(Item.Room) & "</td><td>" & Item.Serial & "</td>" & vbCrLf
    text = text + "</tr>" & vbCrLf
    odd = Not odd
  Next
  text = text + "</table>" & vbCrLf
  text = text + "<p class='complete'>Report Complete</p>"
  
  text = text + "</body>"
  text = text + "</html>"

  GenVacTable_HTML = text



End Function



Function GenTable_HTML(Items As Collection) As String
  Dim Item    As cAssureItem
  Dim text    As String
  Dim row     As String
  Dim odd As Boolean


  text = text + "<html>"
  text = text + "<head>"
  text = text + "<style type=""text/css"">"
  text = text + "body {width:900px; font-family:arial,verdana,sans-serif;}"
  text = text + "table.main {width:900px;font-size:1.0em;}"
  text = text + "tr.header td {background-color: #ADD8E6; color: black; margin:0px; padding:2px; font-weight:bold;}"
  text = text + "tr.even td {background-color: #FAFAD2; color: black; margin:0px; padding:2px}"
  text = text + "tr.odd td {background-color: white; color: black; margin:0px; padding:2px}"
  text = text + "h1 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.3em}"
  text = text + "h2 {background-color: white; color: black;margin:5px;text-align:left;font-size:1.0em;}"
  text = text + "p.complete {background-color: white; color:gray;margin:5px;text-align:left;font-size:0.9em;}"
  
  text = text + "</style>"
  text = text + "</head>"
  text = text + "<body>"
  
  text = text + "<h1>Check-in Report</h1>"
  text = text + "<h1>" & HTMLEncode(Configuration.Facility) & "</h1>"
  text = text + "<h2>" & Format$(Now, "mmmm dd, yy hh:nn AMPM") & "</h2>"
  text = text + "<br />"
  text = text + "<table class='main'>" & vbCrLf
  text = text + "<tr class='header'>" & vbCrLf
  text = text + "<td>Resident</td> <td>Phone</td> <td>Room</td> <td>Device</td>" & vbCrLf
  text = text + "</tr>" & vbCrLf

  For Each Item In Items

    text = text + "<tr " & IIf(odd, "class='odd'", "class='even'") & ">" & vbCrLf
    text = text + "<td>" & HTMLEncode(Item.NameFull) & "</td><td>" & HTMLEncode(Item.Phone) & "</td><td>" & HTMLEncode(Item.Room) & "</td><td>" & HTMLEncode(Item.Serial) & "</td>" & vbCrLf
    text = text + "</tr>" & vbCrLf
    odd = Not odd
  Next
  text = text + "</table>" & vbCrLf
  text = text + "<p class='complete'>Report Complete</p>"
  
  text = text + "</body>"
  text = text + "</html>"

  GenTable_HTML = text



End Function



'Public Sub Send_Log_Files(sDate As String, sSendfiles As String)
'
'    Dim objMsg          As Object
'
'    Set objMsg = CreateObject("CDO.Message")
'    objMsg.From = "IRLS Log Files"
'    objMsg.To = "email@hotmail.com"
'    objMsg.Subject = "IRLS Log files for " & sDate
'    objMsg.TextBody = "Please find attached the IRLS log files for " & sDate
'    objMsg.addattachment "c:\log1.log"
'    objMsg.addattachment "c:\log2.log"
'    objMsg.Send
'    Set objMsg = Nothing
'
'End Sub

Sub CancelAssur()
' clears Assurance bit for all devices


  Dim d     As cESDevice
  Dim start As Long

10        On Error GoTo CancelAssur_Error

20        'start = Win32.timeGetTime()
30        Assurs.Clear

40        For Each d In Devices.Devices
50          d.AssurBit = 0
60        Next

65        Set AssureVacationDevices = New Collection



CancelAssur_Resume:
70        On Error GoTo 0
80        Exit Sub

CancelAssur_Error:

90        LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.CancelAssur." & Erl
100       Resume CancelAssur_Resume

End Sub

Sub CheckAssur()

10        On Error GoTo CheckAssur_Error

20        If Configuration.AssurStart <> Configuration.AssurEnd Then
        
30          DoAssure1
40        End If

50        If Configuration.AssurStart2 <> Configuration.AssurEnd2 Then
60          DoAssure2
70        End If

CheckAssur_Resume:
80        On Error GoTo 0
90        Exit Sub

CheckAssur_Error:

100       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.CheckAssur." & Erl
110       Resume CheckAssur_Resume

End Sub
Sub DoAssure1()

  Dim HrsDiff   As Long



10        On Error GoTo DoAssure1_Error
          'If Configuration.AssurStart < Configuration.AssurEnd Then
            
          Dim tempstart As Date
          Dim tempend As Date
'          tempstart = DateAdd("h", Configuration.AssurStart, Format(Now, "mm-dd-yyyy"))
'          tempend = DateAdd("h", Configuration.AssurEnd, Format(Now, "mm-dd-yyyy"))
'
'          If tempstart < tempend Then
'              gAssurStartTime = tempstart
'              gAssurEndTime = tempend
'          Else
'              gAssurStartTime = tempstart
'              gAssurEndTime = DateAdd("h", 24, tempend)
'
'          End If


20        If gAssurStartTime = 0 Then  ' done once
30          InitAssurePeriod
40        End If

50        If DateDiff("n", gAssurStartTime, Now) >= 0 Then
60          BeginAssure 1
            
70          gAssurStartTime = DateAdd("h", 24, gAssurStartTime)
80        End If

90        If DateDiff("n", gAssurEndTime, Now) >= 0 Then
100         EndAssure
110         If Configuration.AssurStart > Configuration.AssurEnd Then
120           HrsDiff = 24 + (Configuration.AssurEnd - Configuration.AssurStart)
130         Else
140           HrsDiff = Configuration.AssurEnd - Configuration.AssurStart
150         End If
160         gAssurEndTime = DateAdd("h", HrsDiff, gAssurStartTime)
170       End If

DoAssure1_Resume:
180       On Error GoTo 0
190       Exit Sub

DoAssure1_Error:

200       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.DoAssure1." & Erl
210       Resume DoAssure1_Resume


End Sub

Sub DoAssure2()
  Dim HrsDiff   As Long

10           On Error GoTo DoAssure2_Error

20          If gAssurStartTime2 = 0 Then  ' done once
30            InitAssurePeriod2
40          End If

50          If DateDiff("n", gAssurStartTime2, Now) >= 0 Then
60            BeginAssure 2
70            gAssurStartTime2 = DateAdd("h", 24, gAssurStartTime2)
80          End If

90          If DateDiff("n", gAssurEndTime2, Now) >= 0 Then
100           EndAssure
110           If Configuration.AssurStart2 > Configuration.AssurEnd2 Then
120             HrsDiff = 24 + (Configuration.AssurEnd2 - Configuration.AssurStart2)
130           Else
140             HrsDiff = Configuration.AssurEnd2 - Configuration.AssurStart2
150           End If
160           gAssurEndTime2 = DateAdd("h", HrsDiff, gAssurStartTime2)
170         End If


DoAssure2_Resume:
180          On Error GoTo 0
190          Exit Sub

DoAssure2_Error:

200         LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.DoAssure2." & Erl
210         Resume DoAssure2_Resume


End Sub


Sub InitAssurePeriod()

  Dim HrsDiff   As Long

10        On Error GoTo InitAssurePeriod_Error

20        If gAssurStartTime = 0 Then
30          gAssurStartTime = CDate(Format(Now, "mm/dd/yyyy") & " " & Configuration.AssurStart & ":00")
40        Else
50          gAssurStartTime = CDate(Format(DateAdd("h", 24, gAssurStartTime), "mm/dd/yyyy") & " " & Configuration.AssurStart & ":00")
60        End If




70        If Configuration.AssurStart > Configuration.AssurEnd Then
80          HrsDiff = 24 + (Configuration.AssurEnd - Configuration.AssurStart)
90        Else
100         HrsDiff = Configuration.AssurEnd - Configuration.AssurStart
110       End If
120       gAssurEndTime = DateAdd("h", HrsDiff, gAssurStartTime)
130       If DateDiff("n", gAssurEndTime, Now) > 0 Then
140         InitAssurePeriod
            
150       End If

InitAssurePeriod_Resume:
160       On Error GoTo 0
170       Exit Sub

InitAssurePeriod_Error:

180       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.InitAssurePeriod." & Erl
190       Resume InitAssurePeriod_Resume


End Sub

Sub InitAssurePeriod2()

  Dim HrsDiff   As Long

10        On Error GoTo InitAssurePeriod2_Error

20        If gAssurStartTime2 = 0 Then
30          gAssurStartTime2 = CDate(Format(Now, "mm/dd/yyyy") & " " & Configuration.AssurStart2 & ":00")
40        Else
50          gAssurStartTime2 = CDate(Format(DateAdd("h", 24, gAssurStartTime2), "mm/dd/yyyy") & " " & Configuration.AssurStart2 & ":00")
60        End If

70        If Configuration.AssurStart2 > Configuration.AssurEnd2 Then
80          HrsDiff = 24 + (Configuration.AssurEnd2 - Configuration.AssurStart2)
90        Else
100         HrsDiff = Configuration.AssurEnd2 - Configuration.AssurStart2
110       End If
120       gAssurEndTime2 = DateAdd("h", HrsDiff, gAssurStartTime2)

130       If DateDiff("n", gAssurEndTime2, Now) > 0 Then
140         InitAssurePeriod2
            
150       End If

          

InitAssurePeriod2_Resume:
160       On Error GoTo 0
170       Exit Sub

InitAssurePeriod2_Error:

180       LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.InitAssurePeriod2." & Erl
190       Resume InitAssurePeriod2_Resume


End Sub

Sub BeginAssure(ByVal period As Integer)

  Dim start              As Long
  Dim d                  As cESDevice
  Dim DayNum             As Integer
  Dim rs                 As Recordset
  Dim AssureDevices      As Collection
  Dim SQl                As String
  Dim j                  As Integer
  Dim ctr                As Long


  On Error GoTo BeginAssure_Error

  'If InAssurPeriod Then Exit Sub

  Debug.Print "Begin Assure"
  InAssurPeriod = True

  CancelAssur
  DayNum = Weekday(Now, firstdayofweek:=vbSunday)

  Set AssureDevices = New Collection
  Set AssureVacationDevices = New Collection

  start = Win32.timeGetTime()

  ' if resident/room away... add to vacation list, don't flag for assurance
  ' only these are candidates for assurance

'  Select Case period
'    Case 2
'      SQl = "SELECT * FROM Devices WHERE UseAssur2 = 1"
'    Case Else
'      SQl = "SELECT * FROM Devices WHERE UseAssur = 1"
'  End Select
'
'  Debug.Assert 0               ' kill off devices that alerady checked in (happens on reboot)



  For j = 1 To Devices.Count
    ctr = ctr + 1
    If ctr > 200 Then          ' doevents every 200 devices
      ctr = 0
      DoEvents
    End If
    Set d = Devices.Item(j)
    Select Case period
      Case 2
        If d.UseAssur2 = 1 Then
          AssureDevices.Add d
        End If
      Case Else
        If d.UseAssur = 1 Then
          AssureDevices.Add d
        End If
    End Select
  Next





  Debug.Print "******* vacations **********"

  ' do resident assurance
  SQl = "SELECT assurdays, away ,ResidentID FROM Residents WHERE ResidentID <> 0 AND Assurdays >= 2 order by namelast, namefirst"
  Set rs = ConnExecute(SQl)



  Do Until rs.EOF

    If rs("Away") Then
      Debug.Print "ResidentID : Away " & rs("residentID") & " : " & rs("Away")
    End If
    For Each d In AssureDevices

      If d.ResidentID = rs("residentID") Then
        If rs("Away") <> 0 Then
          AssureVacationDevices.Add d
        Else
          If IsAssurday(DayNum, rs("Assurdays")) Then
            d.AssurBit = 1
          End If
        End If

      End If
    Next
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing

  Debug.Print "Room Check-in "

  ' do room assurance
  SQl = "SELECT assurdays, away, Room, RoomID FROM Rooms WHERE  Assurdays >= 2 order by Room "
  Set rs = ConnExecute(SQl)

  Dim Dev                As cESDevice

  'Debug.Assert 0

  Do Until rs.EOF

    If rs("Away") <> 0 Then
      Debug.Print "RoomID : Away " & rs("RoomID") & " : " & rs("Away")
    End If

    For Each d In AssureDevices
      If d.RoomID = rs("roomID") Then
        If rs("Away") <> 0 Then  ' it's away if non-zero
          For Each Dev In AssureVacationDevices
            If Dev.Serial = d.Serial Then
              Exit For         ' bail if already on vacation via resident
            End If
          Next
          If d.AssurBit = 0 Then
            AssureVacationDevices.Add d
          End If
        Else
          If IsAssurday(DayNum, rs("Assurdays")) Then
            d.AssurBit = 1     ' set the assure bit



          End If
        End If

      End If
    Next
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing


  
  Select Case period
    Case 2
      SQl = "SELECT * FROM alarms WHERE EventType = " & EVT_ASSUR_CHECKIN & _
            " AND eventdate >= " & DateDelim & gAssurStartTime2 & DateDelim
    Case Else
      SQl = "SELECT * FROM alarms WHERE EventType = " & EVT_ASSUR_CHECKIN & _
            " AND eventdate >= " & DateDelim & gAssurStartTime & DateDelim
  End Select

  Set rs = ConnExecute(SQl)
  Dim counter As Long
  Do While Not rs.EOF
    counter = counter + 1
    If counter > 200 Then
      counter = 0
      DoEvents
    End If
    
    For Each d In AssureDevices
      If StrComp(d.Serial, rs("serial") & "", vbTextCompare) = 0 Then
        Debug.Print d.Serial & " Already Checked in modAssur.BeginAssure"
        d.AssurBit = 0
        Exit For
      End If
    Next
    rs.MoveNext
  Loop

  rs.Close
  Set rs = Nothing

  PostEvent Nothing, Nothing, Nothing, EVT_ASSUR_START, 0



  Set AssureDevices = Nothing

  Debug.Print "Time to Set all Assure Bits: "; Win32.timeGetTime() - start


BeginAssure_Resume:
  On Error GoTo 0
  Exit Sub

BeginAssure_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.BeginAssure." & Erl
  Resume BeginAssure_Resume


End Sub


Sub EndAssure()

        Dim d     As cESDevice
        Dim start As Long
        Dim counter As Long

10      On Error GoTo EndAssure_Error
20      InAssurPeriod = False

30      Debug.Print "End Assure"

40      start = Win32.timeGetTime()

50      For Each d In Devices.Devices
60        counter = counter + 1
70        If counter > 200 Then
80          counter = 0
90          DoEvents
100       End If
          
110       If d.AssurBit = 1 Then
120         If d.AssurInput = 2 Then
130           PostEvent d, Nothing, Nothing, EVT_ASSUR_FAIL, 2
140         Else
150           PostEvent d, Nothing, Nothing, EVT_ASSUR_FAIL, 1
160         End If
170       End If
180     Next

190     PostEvent Nothing, Nothing, Nothing, EVT_ASSUR_END, 0
200     Debug.Print "Time to Process all Assure Bits: "; Win32.timeGetTime() - start

EndAssure_Resume:
210     On Error GoTo 0
220     Exit Sub

EndAssure_Error:

230     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.EndAssure." & Erl
240     Resume EndAssure_Resume


End Sub

Sub SetDevicesAwayByResident(ByVal ResidentID As Long, ByVal Away As Integer)
        Dim d As cESDevice
10      On Error GoTo SetDevicesAwayByResident_Error

20      For Each d In Devices.Devices
30        If d.ResidentID = ResidentID Then
40          d.IsAway = Away
            
            'd.Alarm = 0 ' added 9/8/2011 to clear device alarm status, ready for next trigger
            'd.Alarm_A = 0
            
50          d.Dead = 0
60          If Away = 0 Then
70            d.LastSupervise = Now
80          End If
90        End If
100     Next

SetDevicesAwayByResident_Resume:
110     On Error GoTo 0
120     Exit Sub

SetDevicesAwayByResident_Error:

130     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.SetDevicesAwayByResident." & Erl
140     Resume SetDevicesAwayByResident_Resume

End Sub

Sub SetDevicesAwayByRoom(ByVal RoomID As Long, ByVal Away As Integer)

        Dim d As cESDevice
        ' not sure if we want to do this
10      On Error GoTo SetDevicesAwayByRoom_Error

20      For Each d In Devices.Devices
30        If d.RoomID = RoomID Then
40          d.IsAway = Away
            'd.Alarm = 0 ' added 9/8/2011 to clear device alarm status, ready for next trigger
            'd.Alarm_A = 0
50          d.Dead = 0
60          If Away = 0 Then
70            d.LastSupervise = Now
80          End If

90        End If
100     Next

SetDevicesAwayByRoom_Resume:
110     On Error GoTo 0
120     Exit Sub

SetDevicesAwayByRoom_Error:

130     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAssur.SetDevicesAwayByRoom." & Erl
140     Resume SetDevicesAwayByRoom_Resume


End Sub
Public Function IsAssurday(ByVal DOW As Integer, ByVal mask As Long) As Boolean
'Sunday is bit 1 (value = 2), Monday is Bit 2, ' bit 0 is reserved
  If DOW > 0 And DOW < 8 Then  ' only 1 thru 7 ' bit 0 is reserved
    IsAssurday = ((2 ^ DOW) And mask) <> 0
  End If
End Function
Public Function GetAssurDaysFromValue(ByVal Value As Long) As String
      'index (bit) 1 thru 7
      'index 1 is monday
      'index 2 is tues etc
        Dim AssurString As String
10      AssurString = String(7, "_")
        ' ReturnValue is either 1 or 0 (on or off)
20      If Value And 2 ^ 1 Then
30        Mid(AssurString, 7, 1) = "S"
40      End If
50      If Value And 2 ^ 2 Then
60        Mid(AssurString, 1, 1) = "M"
70      End If
80      If Value And 2 ^ 3 Then
90        Mid(AssurString, 2, 1) = "T"
100     End If
110     If Value And 2 ^ 4 Then
120       Mid(AssurString, 3, 1) = "W"
130     End If
140     If Value And 2 ^ 5 Then
150       Mid(AssurString, 4, 1) = "T"
160     End If
170     If Value And 2 ^ 6 Then
180       Mid(AssurString, 5, 1) = "F"
190     End If
200     If Value And 2 ^ 7 Then
210       Mid(AssurString, 6, 1) = "S"
220     End If
230     GetAssurDaysFromValue = AssurString

End Function

