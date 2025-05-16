VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBatteryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim X As Double
Dim Y As Double
Dim mAlarms As cAlarms

Sub PrintList(lv As ListView, ByVal Title As String)

  Dim li As ListItem

  Dim inchesperpixelx As Double
  Dim inchesperpixely As Double

  If Title <> "" Then
    Set mAlarms = Alerts
  Else
    Set mAlarms = alarms
  End If
  
  X = 0
  Y = 0

  inchesperpixelx = Printer.TwipsPerPixelX / 1440
  inchesperpixely = Printer.TwipsPerPixelY / 1440


  Printer.Print ""
  banner Title
  Header

  For Each li In lv.ListItems
    DoEvents
    If Printer.CurrentX >= Printer.ScaleHeight - 0.5 Then  ' allow 1/2" margin
      Footer
      FormFeed Printer
      banner Title
      Header
    End If
    row li
  Next
  Footer
  Fini Printer
End Sub
Sub banner(ByVal Title As String)
  Printer.FontSize = 12

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, Title & " List"
  PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "")   ' hh:nn a/p")
  
  Printer.Print " "
  Printer.Print " "

End Sub

Sub Header()

  PrintLJ Printer, Printer.ScaleLeft, Printer.CurrentY, "ID"
  PrintLJ Printer, 0.5, Printer.CurrentY, "Name"
  PrintLJ Printer, 2, Printer.CurrentY, "Room"
  PrintLJ Printer, 3.5, Printer.CurrentY, "Location"
  PrintLJ Printer, 5, Printer.CurrentY, "Time"
  PrintLJ Printer, 6, Printer.CurrentY, "Silenced"
  PrintLJ Printer, 7, Printer.CurrentY, "Ack"
  Printer.Print ""
  HR Printer
  Printer.Print ""

End Sub

Sub row(li As ListItem)

  Dim Key As String
  Dim j   As Long

  Dim rs        As Recordset
  Dim rsres     As Recordset
  Dim rsRoom    As Recordset

  Dim ResidentID As Long

  Dim Serial    As String
  Dim Name      As String
  Dim room      As String
  Dim Location  As String
  Dim Model     As String
  Dim EventTime As String
  Dim Silenced  As String
  Dim ACK       As String

  Dim Device    As cESDevice
  Dim Alarm     As cAlarm
  Dim SQL       As String




  Key = li.Key

  For j = 1 To mAlarms.alarms.Count

    Serial = ""
    Name = ""
    room = ""
    Model = ""
    EventTime = ""
    Silenced = ""
    ACK = ""
    
    Set Alarm = mAlarms.alarms(j)
    If Alarm.Serial = left(Key, 8) Then
      Set Device = Devices.Device(Alarm.Serial)
      If Not Device Is Nothing Then
        ResidentID = Device.ResidentID
        SQL = " SELECT DeviceID, Serial, IDM, IDL, DeviceTypes.model, ResidentID, RoomID" & _
            " FROM Devices LEFT JOIN DeviceTypes ON Devices.model = DeviceTypes.model " & _
            " WHERE Devices.serial =" & Q(Device.Serial)
        Set rs = ConnExecute(SQL)
        If Not rs.EOF Then
          Serial = rs("serial") & ""

          Set rsres = ConnExecute("SELECT * FROM Residents WHERE ResidentID = " & rs("ResidentID"))
          If Not rsres.EOF Then
            Name = ConvertLastFirst(rsres("namelast") & "", rsres("namefirst") & "")
            Set rsRoom = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & rsres("RoomID"))
            If Not rsRoom.EOF Then
              room = rsRoom("Room") & ""
            End If
            rsRoom.Close
          Else
            Set rsRoom = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & rs("RoomID"))
            If Not rsRoom.EOF Then
              room = rsRoom("Room") & ""
            End If
            rsRoom.Close
          End If

          rsres.Close
        End If
        rs.Close



        SQL = " SELECT Rooms.Room, Rooms.Locator, Rooms.Building FROM Devices LEFT JOIN Rooms ON Devices.RoomID = Rooms.RoomID " & _
            " WHERE  Devices.serial =" & Q(Alarm.Serial)
        Set rs = ConnExecute(SQL)
        If Not rs.EOF Then
          room = rs("Room") & ""
        End If
        rs.Close

        EventTime = Format(Alarm.DateTime, gTimeFormatString)    '"hh:nn a/p")
        Silenced = IIf(Alarm.SilenceTime = 0, "", Format(Alarm.SilenceTime, gTimeFormatString))    '"hh:nn a/p"))
        ACK = IIf(Alarm.ACKTime = 0, "", Format(Alarm.ACKTime, gTimeFormatString))    '"hh:nn a/p"))

      End If
      Exit For
    End If
    
  Next


  PrintLJ Printer, Printer.ScaleLeft, Printer.CurrentY, Serial
  PrintLJ Printer, 0.5, Printer.CurrentY, Name
  PrintLJ Printer, 2, Printer.CurrentY, room
  PrintLJ Printer, 3.5, Printer.CurrentY, Location
  PrintLJ Printer, 5, Printer.CurrentY, EventTime
  PrintLJ Printer, 6, Printer.CurrentY, Silenced
  PrintLJ Printer, 6, Printer.CurrentY, ACK
  Printer.Print ""
  DoEvents
End Sub

Sub Footer()
  Dim oldfont As String
  Dim oldsize As Double
  Dim oldbold As Boolean
  
  oldfont = Printer.FontName
  oldsize = Printer.FontSize
  oldbold = Printer.FontBold
  Printer.FontSize = 8
  Printer.FontBold = False
  
  PrintCentered Printer, Printer.ScaleHeight - Printer.TextHeight("A"), "Sentry Wireless V " & App.Major & "." & App.Minor & "." & App.Revision
  
  
  Printer.FontName = oldfont
  Printer.FontSize = oldsize
  Printer.FontBold = oldbold
  
  
End Sub



