VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Option Explicit

Public Partial          As Boolean

Private index           As Long
Private StartIndex      As Long
Private EndIndex        As Long

Private X               As Double
Private Y               As Double

Private inchesperpixelx As Double
Private inchesperpixely As Double


Public Function PrintList(lv As ListView)
        Dim li As ListItem
        Dim Title As String


10       On Error GoTo PrintList_Error

20      If Partial Then
30        StartIndex = lv.GetFirstVisible.index
40        EndIndex = Min(ListViewGetVisibleCount(lv) + StartIndex, lv.ListItems.count)
50        Title = "Partial Device List"
60      Else
70        StartIndex = 1
80        EndIndex = lv.ListItems.count
90        Title = "Device List"
100     End If
110       X = 0
120       Y = 0
  
130     If StartIndex > 0 Then

140       inchesperpixelx = Printer.TwipsPerPixelX / 1440
150       inchesperpixely = Printer.TwipsPerPixelY / 1440


160       Printer.Print " "
170       Printer.CurrentX = 0
180       Printer.CurrentY = 0
190       banner Title
200       Header
    
210       For index = StartIndex To EndIndex
220         Set li = lv.ListItems(index)
            'Debug.Print lv.ListItems(Index).Text
230         DoEvents
240         If Printer.CurrentX >= Printer.ScaleHeight - 0.5 Then  ' allow 1/2" margin
250           Footer
260           FormFeed Printer
270           banner Title
280           Header
290         End If
300         Row li
310       Next
320       Footer
330       Fini Printer
340     End If


PrintList_Resume:
350      On Error GoTo 0
360      Exit Function

PrintList_Error:

        'LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cDeviceList.PrintList." & Erl
370     Resume PrintList_Resume


End Function


Sub Row(li As ListItem)

  Dim key As Long
  Dim Name      As String
  Dim Room      As String

  key = Val(li.key)

  Name = li.text
  Room = li.SubItems(1)
  PrintLJ Printer, 0.5, Printer.CurrentY, Name
  PrintLJ Printer, 2, Printer.CurrentY, Room
  
  Printer.Print ""
  DoEvents
End Sub

Sub banner(ByVal Title As String)
  Printer.FontSize = 12

  Printer.CurrentX = 0
  Printer.CurrentY = 0
  Printer.ScaleMode = vbInches

  Printer.Print " "
  PrintCentered Printer, Printer.CurrentY, Title
  PrintRJ Printer, Printer.ScaleWidth, Printer.CurrentY, Format(Now, "mm/dd/yy " & gTimeFormatString) & IIf(gTimeFormat = 1, " hrs", "")   ' hh:nn a/p")
  
  Printer.Print " "
  Printer.Print " "

End Sub

Sub Header()
  Printer.Print ""
  PrintLJ Printer, 0.5, Printer.CurrentY, "ID"
  PrintLJ Printer, 2, Printer.CurrentY, "Assignment"
  Printer.Print ""
  HR Printer
  

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
  PrintCentered Printer, Printer.ScaleHeight - Printer.TextHeight("A"), PRODUCT_NAME & " (Ver " & App.Major & Format(App.Revision, "000") & ")"
  Printer.FontName = oldfont
  Printer.FontSize = oldsize
  Printer.FontBold = oldbold
End Sub





