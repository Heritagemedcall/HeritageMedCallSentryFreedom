VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlarmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private X As Double
Private Y As Double
Private mAlarms As cAlarms
Private inchesperpixelx As Double
Private inchesperpixely As Double

Const COL_LEFT = 0
Const COL_NAME = 0.7 ' was 1"
Const COL_ROOM = 2.4 ' was 2.5
Const col_location = 3.5
Const COL_TIME = 5
Const COL_ANNOUNCE = 5.75
Const col_Ack = 7



Sub PrintList(lv As ListView, alarms As cAlarms, ByVal Title As String)

        Dim li As ListItem

10       On Error GoTo PrintList_Error

20      Set mAlarms = alarms

30      X = 0
40      Y = 0

50      inchesperpixelx = Printer.TwipsPerPixelX / 1440
60      inchesperpixely = Printer.TwipsPerPixelY / 1440


70      Printer.Print ""
80      banner Title
90      Header

100     For Each li In lv.ListItems
110       DoEvents
120       If Printer.CurrentX >= Printer.ScaleHeight - 0.5 Then  ' allow 1/2" margin
130         Footer
140         FormFeed Printer
150         banner Title
160         Header
170       End If
180       Row li
190     Next
200     Footer
210     Fini Printer

PrintList_Resume:
220      On Error GoTo 0
230      Exit Sub

PrintList_Error:

240     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarmReport.PrintList." & Erl
250     Resume PrintList_Resume

End Sub
Sub banner(ByVal Title As String)
  Printer.FontSize = 12

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, Title & " List"
  PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "") ' hh:nn am/pm")

  Printer.Print " "
  Printer.Print " "

End Sub

Sub Header()

  PrintLJ Printer, Printer.ScaleLeft, Printer.CurrentY, "ID"
  PrintLJ Printer, COL_NAME, Printer.CurrentY, "Name"
  PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
  PrintLJ Printer, col_location, Printer.CurrentY, "Location"
  PrintLJ Printer, COL_TIME, Printer.CurrentY, "Time"
  PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, "Announce"
  PrintLJ Printer, col_Ack, Printer.CurrentY, "Ack"
  Printer.Print ""
  HR Printer
  Printer.Print ""

End Sub

Sub Row(li As ListItem)

        Dim key As Long
        Dim serial    As String
        Dim oldfont As String
        Dim oldsize As Double
        Dim oldbold As Boolean



10      On Error GoTo Row_Error

11      oldfont = Printer.FontName
12      oldsize = Printer.FontSize
13      oldbold = Printer.FontBold

20      Printer.FontSize = 9

30      serial = left(li.key, 8)
40      PrintLJ Printer, Printer.ScaleLeft, Printer.CurrentY, li.text
50      PrintLJ Printer, COL_NAME, Printer.CurrentY, li.SubItems(1)
60      PrintLJ Printer, COL_ROOM, Printer.CurrentY, li.SubItems(2)
70      PrintLJ Printer, col_location, Printer.CurrentY, li.SubItems(3)
80      PrintLJ Printer, COL_TIME, Printer.CurrentY, li.SubItems(4)
90      PrintLJ Printer, COL_ANNOUNCE, Printer.CurrentY, li.SubItems(5)
100     PrintLJ Printer, col_Ack, Printer.CurrentY, li.SubItems(6)
110     Printer.Print ""
120     DoEvents

130     Printer.FontName = oldfont
140     Printer.FontSize = oldsize
150     Printer.FontBold = oldbold

Row_Resume:
160     On Error GoTo 0
170     Exit Sub

Row_Error:

180     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cAlarmReport.Row." & Erl
190     Resume Row_Resume


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



