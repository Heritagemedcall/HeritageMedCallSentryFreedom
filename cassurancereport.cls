VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAssuranceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim X As Double
Dim y As Double

Public Function AssurPrintList(lv As ListView, ByVal Title As String) As Boolean

  If Printer Is Nothing Then Exit Function

  Dim li As ListItem

  Dim inchesperpixelx As Double
  Dim inchesperpixely As Double

  Dim counter As Long

  On Error GoTo PrintError

  X = 0
  y = 0

  inchesperpixelx = Printer.TwipsPerPixelX / 1440
  inchesperpixely = Printer.TwipsPerPixelY / 1440


  Printer.Print ""
  AssurBanner Title
  AssurHeader

  For Each li In lv.ListItems
    If Printer.CurrentX >= Printer.ScaleHeight - 0.5 Then  ' allow 1/2" margin
      AssurBanner Title
      AssurHeader
    End If
    counter = counter + 1
    If counter > 200 Then
      counter = 0
      DoEvents
    End If
    AssurRow left(li.Key, 8)
  Next
  Printer.EndDoc
  VacationPrintList


  AssurPrintList = Err.Number = 0

PrintError:

End Function

Private Sub AssurBanner(ByVal Title As String)
  Printer.FontSize = 12

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, Title & " List"
  Printer.Print " "
  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, "Printed " & Format(Now, "dddd, mmm dd, yyyy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "")   ' hh:nn AM/PM")
  Printer.Print " "
  Printer.Print " "

End Sub

Private Sub AssurHeader()

  PrintLJ Printer, 0.5, Printer.CurrentY, "ID"
  PrintLJ Printer, 1.5, Printer.CurrentY, "Name"
  PrintLJ Printer, 4, Printer.CurrentY, "Room"
  PrintLJ Printer, 5, Printer.CurrentY, "Phone"
  'PrintLJ Printer, 5, Printer.CurrentY, "Time"
  'PrintLJ Printer, 6, Printer.CurrentY, "Silenced"
  Printer.Print ""
  HR Printer
  Printer.Print ""

End Sub

Private Sub AssurRow(ByVal Key As String)


        Dim j   As Long

        Dim rs        As Recordset
        Dim rsres     As Recordset
        Dim rsRoom    As Recordset
        Dim Serial    As String
        Dim name      As String
        Dim Room      As String
        Dim Model     As String
        Dim Phone     As String
        Dim EventTime As String

        'Dim Device    As cESDevice

        'Dim ResRoomID As Long  ' resident's room if resident
        'Dim RoomID    As Long  ' room if not resident (facility device)

        Dim Alarm     As cAlarm

        Dim SQl       As String


10      Serial = ""
20      name = ""
30      Room = ""
40      Model = ""
50      Phone = ""

60      EventTime = ""


70      For j = 1 To Assurs.alarms.Count
80        name = ""
90        Room = ""
100       Phone = ""
110       DoEvents
          
120       Set Alarm = Assurs.alarms(j)
130       If Alarm.Serial = Key Then
            'Set Device = Devices.Device(key)
            'If Not Device Is Nothing Then
              'ResidentID = device.ResidentID
140           Serial = Alarm.Serial
              ' get device
150           SQl = " SELECT DeviceID, Serial, model, ResidentID, RoomID " & _
                  " FROM Devices WHERE serial =" & q(Alarm.Serial)
160           Set rs = ConnExecute(SQl)
170           If Not rs.EOF Then
                'Serial = rs("serial") & ""
180             Set rsres = ConnExecute("SELECT * FROM Residents WHERE ResidentID = " & rs("ResidentID"))
190             If Not rsres.EOF Then
200               If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
210                 name = ""
220               Else
230                 name = ConvertLastFirst(rsres("namelast") & "", rsres("namefirst") & "")
240               End If
                  
250               Phone = rsres("Phone") & ""  ' phone
260             End If
270           End If
280           rsres.Close

290           Set rsRoom = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & rs("RoomID"))
300           If Not rsRoom.EOF Then
310             Room = rsRoom("Room") & ""
320           End If
330           rsRoom.Close
           ' End If
340         rs.Close
          
350         If gTimeFormat = 1 Then
360           EventTime = Format(Alarm.DateTime, "mm/dd hh:nn")
370         Else
380           EventTime = Format(Alarm.DateTime, "mm/dd hh:nn am/pm")
390         End If
            

400         Exit For
410       End If
420     Next


430     PrintLJ Printer, 0.5, Printer.CurrentY, Serial
440     PrintLJ Printer, 1.5, Printer.CurrentY, name
450     PrintLJ Printer, 4, Printer.CurrentY, Room
460     PrintLJ Printer, 5, Printer.CurrentY, Phone
470     Printer.Print ""

End Sub


 Private Sub VacationPrintList()

  Dim d       As cESDevice

  VacationBanner
  VacationHeader

  Dim VacationItems As Collection
  Dim j             As Long
  Dim InClause      As String
  Dim SQl           As String
  Dim rs            As ADODB.Recordset
  Dim AssurListSerial() As String
  
  

  If Not AssureVacationDevices Is Nothing Then

    If AssureVacationDevices.Count Then
      ReDim AssurListSerial(1 To AssureVacationDevices.Count)
      

      For j = 1 To AssureVacationDevices.Count
        AssurListSerial(j) = "'" & AssureVacationDevices(j).Serial & "'"
      Next
      InClause = Join(AssurListSerial, ",")


      SQl = "SELECT Devices.Serial, Devices.DeviceID, Residents.NameLast, Rooms.Room, Residents.NameFirst,  Residents.phone, Residents.RoomID, Rooms.RoomID, Devices.RoomID, Devices.ResidentID " & _
          " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
          " WHERE Devices.residentid <> 0 AND (Devices.Serial In (" & InClause & ")) " & _
          " ORDER BY Residents.NameLast, Residents.NameFirst, Rooms.Room; "

      Set rs = ConnExecute(SQl)
      Do Until rs.EOF

        PrintLJ Printer, 0.5, Printer.CurrentY, rs("serial")
        PrintLJ Printer, 1.5, Printer.CurrentY, ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
        PrintLJ Printer, 4, Printer.CurrentY, rs("Room") & ""
        PrintLJ Printer, 5, Printer.CurrentY, rs("phone") & ""
        Printer.Print ""
        If (Printer.ScaleHeight - Printer.CurrentY) < 0.75 Then
          Printer.NewPage
          VacationHeader
        End If
        rs.MoveNext
      Loop
      rs.Close


      SQl = "SELECT Devices.Serial, Devices.DeviceID, Residents.NameLast, Rooms.Room, Residents.NameFirst,  Residents.phone, Residents.RoomID, Rooms.RoomID, Devices.RoomID, Devices.ResidentID " & _
          " FROM (Devices LEFT JOIN Residents ON Devices.ResidentID = Residents.ResidentID) LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
          " WHERE Devices.residentid = 0 AND (Devices.Serial In (" & InClause & ")) " & _
          " ORDER BY Rooms.Room; "

      Set rs = ConnExecute(SQl)
      Do Until rs.EOF

        PrintLJ Printer, 0.5, Printer.CurrentY, rs("serial")
        PrintLJ Printer, 1.5, Printer.CurrentY, ConvertLastFirst(rs("namelast") & "", rs("namefirst") & "")
        PrintLJ Printer, 4, Printer.CurrentY, rs("Room") & ""
        PrintLJ Printer, 5, Printer.CurrentY, rs("phone") & ""
        Printer.Print ""
        If (Printer.ScaleHeight - Printer.CurrentY) < 0.75 Then
          Printer.NewPage
          VacationHeader
        End If
        rs.MoveNext
      Loop
      rs.Close


      '      For Each d In AssureVacationDevices
      '        VacationRow d.Serial
      '        If (Printer.ScaleHeight - Printer.CurrentY) < 0.75 Then
      '          Printer.NewPage
      '          VacationHeader
      '
      '
      '        End If
      '      Next

    End If
  End If
  Set rs = Nothing
  Printer.EndDoc
  ' Set AssureVacationDevices = New Collection
End Sub


Private Sub VacationRow(ByVal Key As String)

  Dim rs        As Recordset
  Dim rsres     As Recordset
  Dim rsRoom    As Recordset
  Dim Serial    As String
  Dim name      As String
  Dim Room      As String
  Dim Model     As String
  Dim Phone     As String
  Dim device    As cESDevice
  Dim SQl       As String


  Serial = ""
  name = ""
  Room = ""
  Model = ""
  Phone = ""

  Set device = Devices.device(Key)
  If Not device Is Nothing Then
   
    SQl = " SELECT DeviceID, Serial, model, ResidentID, RoomID" & _
        " FROM Devices WHERE serial =" & q(Key)
    Set rs = ConnExecute(SQl)
    If Not rs.EOF Then
      Serial = rs("serial") & ""
      Set rsres = ConnExecute("SELECT * FROM Residents WHERE ResidentID = " & rs("ResidentID"))
      If Not rsres.EOF Then
        name = ConvertLastFirst(rsres("namelast") & "", rsres("namefirst") & "")
        Phone = rsres("phone") & ""
      End If
      rsres.Close

      Set rsRoom = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & rs("RoomID"))
      If Not rsRoom.EOF Then
        Room = rsRoom("Room") & ""
      End If
      rsRoom.Close
    End If
    rs.Close

    PrintLJ Printer, 0.5, Printer.CurrentY, Serial
    PrintLJ Printer, 1.5, Printer.CurrentY, name
    PrintLJ Printer, 4, Printer.CurrentY, Room
    PrintLJ Printer, 5, Printer.CurrentY, Phone
    Printer.Print ""

  End If

End Sub
Private Sub VacationBanner()
  Printer.FontSize = 12

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, "Vacation List"
  Printer.Print " "
  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, "Printed " & Format(Now, "dddd, mmm dd, yyyy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "")   ' hh:nn AM/PM")
  Printer.Print " "
  Printer.Print " "

End Sub

Private Sub VacationHeader()



  PrintLJ Printer, 0.5, Printer.CurrentY, "ID"
  PrintLJ Printer, 1.5, Printer.CurrentY, "Name"
  PrintLJ Printer, 4, Printer.CurrentY, "Room"
  PrintLJ Printer, 5, Printer.CurrentY, "Phone"
  'PrintLJ Printer, 5, Printer.CurrentY, "Time"
  'PrintLJ Printer, 6, Printer.CurrentY, "Silenced"
  Printer.Print ""
  HR Printer
  Printer.Print ""

End Sub




