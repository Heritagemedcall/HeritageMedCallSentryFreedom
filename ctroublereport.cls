VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTroubleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim X As Double
Dim Y As Double
Dim mAlarms As cAlarms
Dim Columns As Integer

Const COL_LEFT = 0
Const COL_NAME = 1
Const COL_ROOM = 2.5
'Const COL_LOCATION = 3.5 ' OR DEVICE
Const COL_DEVICE = 3.5 ' OR DEVICE
Const COL_TIME = 5
Const COL_SILENCED = 5.75
Const COL_TYPE = 7

'Const col_Announce = 5.75
'Const col_Ack = 7



Public Sub PrintList(lv As ListView, alarms As cAlarms, ByVal Title As String)

        Dim li As ListItem

        Dim inchesperpixelx As Double
        Dim inchesperpixely As Double

        'On Error GoTo 0

10      On Error GoTo PrintList_Error

20      Set mAlarms = alarms
  
30      Columns = lv.ColumnHeaders.count
40      X = 0
50      Y = 0

60      inchesperpixelx = Printer.TwipsPerPixelX / 1440
70      inchesperpixely = Printer.TwipsPerPixelY / 1440


80      Printer.Print ""
90      banner Title
100     Header

110     For Each li In lv.ListItems
120       DoEvents
130       If Printer.CurrentX >= Printer.ScaleHeight - 0.5 Then  ' allow 1/2" margin
140         Footer
150         FormFeed Printer
160         banner Title
170         Header
180       End If
190       Row li, lv.ColumnHeaders.count
200     Next
210     Footer

PrintList_Resume:

220     Fini Printer
230      On Error GoTo 0
240      Exit Sub

PrintList_Error:

250     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cTroubleReport.PrintList." & Erl
260     Resume PrintList_Resume

End Sub
Private Sub banner(ByVal Title As String)
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

  PrintLJ Printer, 0, Printer.CurrentY, "ID"
  PrintLJ Printer, COL_NAME, Printer.CurrentY, "Name"
  PrintLJ Printer, COL_ROOM, Printer.CurrentY, "Room"
  PrintLJ Printer, COL_DEVICE, Printer.CurrentY, "Device" ' model
  PrintLJ Printer, COL_TIME, Printer.CurrentY, "Time"
  PrintLJ Printer, COL_SILENCED, Printer.CurrentY, "Silenced"
  PrintLJ Printer, COL_TYPE, Printer.CurrentY, "Type"
  
  Printer.Print ""
  HR Printer
  Printer.Print ""

End Sub

Private Sub Row(li As ListItem, ByVal NumCols As Integer)

  
  
  
  
  
  PrintLJ Printer, 0, Printer.CurrentY, li.text
  PrintLJ Printer, COL_NAME, Printer.CurrentY, li.SubItems(1)
  PrintLJ Printer, COL_ROOM, Printer.CurrentY, li.SubItems(2)
  PrintLJ Printer, COL_DEVICE, Printer.CurrentY, li.SubItems(3)
  PrintLJ Printer, COL_TIME, Printer.CurrentY, li.SubItems(4)
  PrintLJ Printer, COL_SILENCED, Printer.CurrentY, li.SubItems(5)
  If NumCols > 6 Then
    PrintLJ Printer, COL_TYPE, Printer.CurrentY, li.SubItems(6)
  End If
  
  Printer.Print ""
  DoEvents
End Sub

Private Sub Footer()
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



