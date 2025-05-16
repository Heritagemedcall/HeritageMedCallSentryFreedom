Attribute VB_Name = "modPrint"
Option Explicit
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" _
    (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long

Sub PrintLineFeed(Device As Object)
  Device.Print "" & vbCrLf

End Sub

Sub PrintLJ(Device As Object, X, y, ByVal text As String)

  Device.CurrentX = X
  Device.CurrentY = y
  Device.Print text;

End Sub
Sub PrintRJ(Device As Object, X, y, ByVal text As String)

  Device.CurrentX = X - Device.TextWidth(text)
  Device.CurrentY = y
  Device.Print text;

End Sub

Sub PrintCJ(Device As Object, X, y, ByVal text As String)

  Device.CurrentX = X - (Device.TextWidth(text) \ 2)
  Device.CurrentY = y
  Device.Print text;

End Sub

Sub PrintCentered(Device As Object, y, ByVal text As String)

  Device.CurrentX = Device.ScaleWidth / 2 - Device.TextWidth(text) / 2
  Device.CurrentY = y
  Device.Print text;
End Sub
Sub HR(Device As Object)
  Dim X As Double
  Dim y As Double

  X = Device.CurrentX
  y = Device.CurrentY

  Device.Line (Device.ScaleLeft, Device.CurrentY)-(Device.ScaleWidth, Device.CurrentY)

  Device.CurrentX = X
  Device.CurrentY = y

End Sub


Sub HRSegment(Device As Object, ByVal start As Double, ByVal finish As Double)
  Dim X As Double
  Dim y As Double

  X = Device.CurrentX
  y = Device.CurrentY

  Device.Line (start, Device.CurrentY)-(finish, Device.CurrentY)

  Device.CurrentX = X
  Device.CurrentY = y

End Sub


Sub FormFeed(Device As Object)
  If Device Is Printer Then
    Device.NewPage
  End If
End Sub

Sub Fini(Device As Object)
  If Device Is Printer Then
    Device.EndDoc
  End If

End Sub

Public Function PrintScreen(ByVal frm As Form) As Boolean
        Static Busy
10      On Error GoTo PrintScreen_Error

20      If Busy Then Exit Function
30      Busy = True

        Dim fw As Long, fh As Long
        Dim printerpixelsx  As Long
        Dim ratio     As Double
        Dim pixelsX   As Long
        Dim RC As Long
        Dim screenHDC As Long

        Dim FormLeft As Double
        Dim FormTop As Double

        Dim printerwidth As Double
        Dim printerheight As Double

        Dim formwidth As Double
        Dim formheight As Double

        Dim ratiox As Double
        Dim ratioy As Double
        Dim aspect As Double
        Dim TitlebarHeight As Double
        Dim MemDC As cMemDC


        If Printer Is Nothing Then Exit Function
        



40      frm.Refresh
50      DoEvents

60      Printer.ScaleMode = vbTwips
70      printerwidth = Printer.Width / Printer.TwipsPerPixelX  ' pixels
80      printerheight = Printer.Height / Printer.TwipsPerPixelY  ' pixels

90      formwidth = frm.Width / Screen.TwipsPerPixelX  ' pixels
100     formheight = frm.Height / Screen.TwipsPerPixelY  ' pixels

110     FormLeft = frm.left / Screen.TwipsPerPixelX
120     FormTop = frm.top / Screen.TwipsPerPixelY

130     ratiox = (Screen.TwipsPerPixelX / Printer.TwipsPerPixelX)
140     ratioy = (Screen.TwipsPerPixelY / Printer.TwipsPerPixelY)
150     aspect = ratiox / ratioy

160     ratiox = Format(ratiox, "0.0")

170     Printer.Print " "
180     frm.Refresh
190     DoEvents

200     Do While formwidth * ratiox >= printerwidth * 0.9
210       DoEvents
220       ratiox = ratiox - 0.1
230     Loop

240     Set MemDC = New cMemDC
250     MemDC.CreateFromWidthHeight formwidth * ratiox, formheight * ratiox  ' * aspect

260     screenHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)

270     RC = Win32.StretchBlt(MemDC.hDC, 0, 0, MemDC.Width, MemDC.Height, screenHDC, FormLeft, FormTop, formwidth, formheight, SRCCOPY)
280     DoEvents
290     DeleteDC screenHDC

300     RC = Win32.BitBlt(Printer.hDC, 0, 0, MemDC.Width, MemDC.Height, MemDC.hDC, 0, 0, SRCCOPY)
310     DoEvents
320     Printer.EndDoc
330     DoEvents

340     Set MemDC = Nothing
350     If Printer.ScaleMode <> vbTwips Then
360       Printer.ScaleMode = vbTwips
370     End If

PrintScreen_Resume:

380     Busy = False
390     On Error GoTo 0
400     Exit Function

PrintScreen_Error:

410     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modPrint.PrintScreen." & Erl
420     Resume PrintScreen_Resume


End Function



