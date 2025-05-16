VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWaypointReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Waypoint ID (Description) - Repeater 1 serial - Repeater 1 room - 1 signal - Repeater 2 serial - Repeater 2 room - 2 signal - Repeater 3 serial - Repeater 3 room - 3 signal
 

Const COL_WAYPOINT = 0

Const COL_REPEATER1 = 2.1
Const COL_ROOM1 = COL_REPEATER1 + 0.8
Const COL_SIGNAL1 = 4.6 ' RJ

Const COL_REPEATER2 = 4.7
Const COL_ROOM2 = COL_REPEATER2 + 0.8
Const COL_SIGNAL2 = 7.3 ' RJ


Const COL_REPEATER3 = 7.4
Const COL_ROOM3 = COL_REPEATER3 + 0.8
Const COL_SIGNAL3 = 10 ' RJ

Private Title    As String


Sub PrintReport()
        Dim waypoint  As cWayPoint
        'Dim Filename      As String
        Dim j             As Integer
10      On Error GoTo PrintReport_Error

20      Printer.Orientation = vbPRORLandscape

30      Printer.Print " ";
40      Printer.ScaleMode = vbInches
50      Printer.FontSize = 10
60      Printer.FontBold = False


70      PrintBanner
80      PrintRowHeader
90      For j = 1 To Waypoints.count
        Set waypoint = Waypoints.waypoint(j)
          
          
        
        
100       DoEvents
          ' check pagelength
110       If (Printer.ScaleHeight - Printer.CurrentY) < 0.5 Then  '  half inch bottom margin
120         PrintFooter
130         FormFeed Printer
140         PrintRowHeader
150       End If
160       PrintRow waypoint
170     Next


190     PrintFooter
200     Printer.EndDoc
210     Printer.Orientation = vbPRORPortrait


PrintReport_Resume:
230     On Error GoTo 0
240     Exit Sub

PrintReport_Error:

250     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cWaypointReport.PrintReport." & Erl
260     Resume PrintReport_Resume


End Sub


Sub PrintRow(waypoint As cWayPoint)
  Dim i As Integer
  Dim dev As cESDevice

  For i = 1 To Devices.count
    Set dev = Devices.Devices(i)
    If dev.Serial = waypoint.Repeater1 Then
      waypoint.Repeater1Desc = GetRoomByID(dev.RoomID)
      Exit For
    End If
  Next
  DoEvents

  For i = 1 To Devices.count
    Set dev = Devices.Devices(i)
    If dev.Serial = waypoint.Repeater2 Then
      waypoint.Repeater2Desc = GetRoomByID(dev.RoomID)
      Exit For
    End If
  Next
  DoEvents

  For i = 1 To Devices.count
    Set dev = Devices.Devices(i)
    If dev.Serial = waypoint.Repeater3 Then
      waypoint.Repeater3Desc = GetRoomByID(dev.RoomID)
      Exit For
    End If
  Next
  DoEvents


  PrintLJ Printer, COL_WAYPOINT, Printer.CurrentY, waypoint.Description
  
  PrintLJ Printer, COL_REPEATER1, Printer.CurrentY, waypoint.Repeater1
  PrintLJ Printer, COL_ROOM1, Printer.CurrentY, waypoint.Repeater1Desc
  PrintRJ Printer, COL_SIGNAL1, Printer.CurrentY, waypoint.Signal1
  
  PrintLJ Printer, COL_REPEATER2, Printer.CurrentY, waypoint.Repeater2
  PrintLJ Printer, COL_ROOM2, Printer.CurrentY, waypoint.Repeater2Desc
  PrintRJ Printer, COL_SIGNAL2, Printer.CurrentY, waypoint.Signal2
  
  PrintLJ Printer, COL_REPEATER3, Printer.CurrentY, waypoint.Repeater3
  PrintLJ Printer, COL_ROOM3, Printer.CurrentY, waypoint.Repeater3Desc
  PrintRJ Printer, COL_SIGNAL3, Printer.CurrentY, waypoint.Signal3
  
  Printer.Print " "


End Sub

Sub PrintFooter()
' copyright etc
  Dim text As String
  text = App.LegalCopyright & " " & App.ProductName & " (Ver. 1." & Format(App.Revision, "000") & ")"
  PrintCJ Printer, Printer.ScaleWidth / 2, Printer.ScaleHeight - Printer.TextHeight("A") * 1.5, text


End Sub
Sub PrintBanner()
' date and time of report

  Title = "Waypoint Report"

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print ""
  PrintCentered Printer, Printer.CurrentY, Title
  PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "") ' hh:nn am/pm")

  Printer.Print ""
  Printer.Print ""


End Sub
Sub PrintRowHeader()
  
  PrintLJ Printer, COL_WAYPOINT, Printer.CurrentY, "Waypoint"
  PrintLJ Printer, COL_REPEATER1, Printer.CurrentY, "Repeater"
  PrintLJ Printer, COL_ROOM1, Printer.CurrentY, "Location"
  PrintRJ Printer, COL_SIGNAL1, Printer.CurrentY, "Level"
  
  PrintLJ Printer, COL_REPEATER2, Printer.CurrentY, "Repeater"
  PrintLJ Printer, COL_ROOM2, Printer.CurrentY, "Location"
  PrintRJ Printer, COL_SIGNAL2, Printer.CurrentY, "Level"
  
  
  PrintLJ Printer, COL_REPEATER3, Printer.CurrentY, "Repeater"
  PrintLJ Printer, COL_ROOM3, Printer.CurrentY, "Location"
  PrintRJ Printer, COL_SIGNAL3, Printer.CurrentY, "Level"

  Printer.Print " "
  HRSegment Printer, COL_WAYPOINT, COL_REPEATER1 - 0.1
  HRSegment Printer, COL_REPEATER1, COL_SIGNAL1
  HRSegment Printer, COL_REPEATER2, COL_SIGNAL2
  HRSegment Printer, COL_REPEATER3, COL_SIGNAL3


End Sub
