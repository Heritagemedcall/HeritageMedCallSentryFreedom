VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInfoPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private inchesperpixelx As Double
Private inchesperpixely As Double
Private res As cResident
Public Function PrintInfo(ByVal ID As Long)
  Dim Title As String
  
  Dim SQl As String
  Dim Room As String
  
  Set res = New cResident
  res.Fetch ID
  res.GetTransmitters
  res.Room = GetResidentRooms(ID)

  inchesperpixelx = Printer.TwipsPerPixelX / 1440
  inchesperpixely = Printer.TwipsPerPixelY / 1440
  
  If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
    Title = "Resident/Room Info"
  Else
    Title = "Resident Info for " & ConvertLastFirst(res.NameLast, res.NameFirst)
  End If
  Printer.Print ""
  banner Title
  PrintData
  Footer
  Fini Printer

  
End Function
Sub PrintData()
  Dim y As Double
  'Printer.FontSize = 10
  
  y = Printer.CurrentY
  PrintLJ Printer, 0.1, y, "Resident:"
  If gUser.LEvel <= LEVEL_USER And Configuration.HideHIPPANames <> 0 Then
    PrintLJ Printer, 1, y, "N/A"
  Else
    PrintLJ Printer, 1, y, ConvertLastFirst(res.NameLast, res.NameFirst)
  End If
  
  Printer.Print
  y = Printer.CurrentY
  
  PrintLJ Printer, 0.1, y, "Room:"
  PrintLJ Printer, 1, y, res.Room
  
  Printer.Print
  y = Printer.CurrentY
  
  PrintLJ Printer, 0.1, y, "Phone:"
  PrintLJ Printer, 1, y, res.Phone
    
  Printer.Print
  y = Printer.CurrentY
  
  PrintLJ Printer, 0.1, y, "Notes:"
  printWR Printer, 1, 8, y, res.info
  
  'drawtext printer.hDC ,res.Info,len(res.Info),rect,
  
  
  y = Printer.CurrentY
  'PrintLJ Printer, 0.1, Y, "Hello:"
  
  
  
  'Printer.Print
  'Y = Printer.CurrentY
  
  
  
 ' PrintLJ Printer, 0.1, Y, "Status:"
 ' PrintLJ Printer, 1, Y, IIf(res.Vacation = 1, "On Vacation", "Present")
  
  
  

End Sub
Sub printWR(device As Object, ByVal LeftMarg As Double, ByVal RightMarg As Double, ByVal top As Double, ByVal text As String)
  Dim RECT As RECT
  RECT.left = LeftMarg / inchesperpixelx
  RECT.Right = RightMarg / inchesperpixelx
  RECT.top = top / inchesperpixely
  RECT.Bottom = top / inchesperpixely
  DrawText Printer.hDC, text, Len(text), RECT, DT_WORDBREAK Or Win32.DT_CALCRECT
  DrawText Printer.hDC, text, Len(text), RECT, DT_WORDBREAK
  Printer.CurrentY = RECT.Bottom * inchesperpixely

  

End Sub


Sub banner(ByVal Title As String)


  Printer.FontSize = 12

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, Title
  PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "") ' hh:nn am/pm")

  Printer.Print " "
  Printer.Print " "


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

