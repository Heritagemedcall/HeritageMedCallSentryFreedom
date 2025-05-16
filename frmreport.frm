VERSION 5.00
Begin VB.Form frmPrintPreview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Preview"
   ClientHeight    =   10935
   ClientLeft      =   8055
   ClientTop       =   2820
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picReport 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   11265
      Left            =   -30
      ScaleHeight     =   11205
      ScaleWidth      =   12750
      TabIndex        =   0
      Top             =   -15
      Visible         =   0   'False
      Width           =   12810
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PrintPages As Collection

Private rsalarms      As Recordset




Const Col0 = 0.25
Const col1 = 1
Const col2 = 2
Const col3 = 3
Const Col4 = 4
Const Col5 = 5
Const Col6 = 6
Const Col7 = 7
Const Col8 = 8

Sub DoReport(ByVal DateRange As String, alarmtypes As String)  ' comma delimited params

  Dim SQl As String

  SQl = "SELECT * FROM alarms WHERE alarmID = 0"  'primary alarms only
  Set rsalarms = ConnExecute(SQl)
  picReport.ScaleMode = vbInches

  picReport.Print ""
  PrintHeaders

  Do Until rsalarms.EOF
    PrintRow rsalarms
    PrintSubEvents (rsalarms("ID"))
    picReport.Print ""
    rsalarms.MoveNext
  Loop
  rsalarms.Close
  Set rsalarms = Nothing




End Sub
Sub PrintHeaders()
  picReport.FontBold = True
  picReport.CurrentX = Col0
  picReport.Print "Device ID";

  picReport.CurrentX = col1
  picReport.Print "Event Type";


  picReport.CurrentX = col2
  picReport.Print "Time";

  picReport.CurrentX = col3
  picReport.Print "Resident";

  picReport.CurrentX = Col4
  picReport.Print "Room";

  picReport.Print "";
  picReport.CurrentX = Col5
  picReport.Print "";
  picReport.CurrentX = Col6
  picReport.Print "";
  picReport.CurrentX = Col7
  picReport.Print "";
  picReport.CurrentX = Col8
  picReport.Print "";

  picReport.Print ""
  picReport.Line (Col0, picReport.CurrentY)-(Col8, picReport.CurrentY)
  picReport.Print ""

End Sub

Sub PrintSubEvents(ByVal AlarmID As Long)
  Dim rs As Recordset
  Set rs = ConnExecute("SELECT * FROM alarms WHERE AlarmID = " & AlarmID & " ORDER BY ID ")
  Do Until rs.EOF
    PrintSubRow rs
    rs.MoveNext
  Loop

  rs.Close



End Sub
Sub PrintSubRow(rs As Recordset)

  picReport.CurrentX = Col0

  picReport.CurrentX = col1
  picReport.Print GetEventTypeName(rs("eventtype"));

  picReport.CurrentX = col2
  picReport.Print Format(rs("EventDate"), "mm/dd " & gTimeFormatString);  ';

  picReport.CurrentX = col3
  picReport.Print ;

  picReport.CurrentX = Col4
  picReport.Print "";

  picReport.CurrentX = Col5
  picReport.Print "";

  picReport.CurrentX = Col6
  picReport.Print "";

  picReport.CurrentX = Col7
  picReport.Print "";

  picReport.CurrentX = Col8
  picReport.Print "";

  picReport.Print ""


End Sub
Sub PrintRow(rs As Recordset)
  Dim rsinfo  As Recordset
  Dim ResID   As Long
  Dim RoomID  As Long
  Dim name    As String
  Dim Room    As String

  ResID = Val(0 & rs("ResidentID"))

  Set rsinfo = ConnExecute("SELECT * FROM Residents WHERE ResidentID = " & ResID)
  If Not rsinfo.EOF Then
    RoomID = Val(0 & rsinfo("RoomID"))
    name = ConvertLastFirst(rsinfo("namelast") & "", rsinfo("namefirst") & "")
  End If
  rsinfo.Close
  
  'TODO: get room info from resident info

  If RoomID = 0 Then
    RoomID = Val(0 & rs("RoomID"))
  End If

  Set rsinfo = ConnExecute("SELECT * FROM Rooms WHERE RoomID = " & RoomID)
  If Not rsinfo.EOF Then

    Room = rsinfo("Room") & ""
  End If
  rsinfo.Close

  picReport.CurrentX = Col0
  picReport.Print rs("serial") & "";
  picReport.CurrentX = col1
  picReport.Print GetEventTypeName(rs("eventtype"));

  picReport.CurrentX = col2
  picReport.Print Format(rs("EventDate"), "mm/dd " & gTimeFormatString);  ' hh:nn");

  picReport.CurrentX = col3
  picReport.Print name;

  picReport.CurrentX = Col4
  picReport.Print Room;

  picReport.CurrentX = Col5
  picReport.Print "";

  picReport.CurrentX = Col6
  picReport.Print "";

  picReport.CurrentX = Col7
  picReport.Print "";

  picReport.CurrentX = Col8
  picReport.Print "";
  picReport.Print ""
End Sub

Sub LineFeed(ByVal hDC As Long, X, Y)

  Dim RECT As Win32.RECT
  Dim text As String
  text = vbCrLf
  RECT.top = Y
  RECT.left = X
  DrawText hDC, text, Len(text), RECT, DT_NOPREFIX Or Win32.DT_LEFT


End Sub

Sub PrintLJ(ByVal hDC As Long, X, Y, ByVal text As String)
  Dim RECT As Win32.RECT
  RECT.top = Y
  RECT.left = X
  DrawText hDC, text, Len(text), RECT, DT_NOPREFIX Or Win32.DT_LEFT Or DT_SINGLELINE

End Sub
Sub PrintRJ(ByVal hDC As Long, X, Y, ByVal text As String)

  Dim RECT As Win32.RECT
  RECT.top = Y
  RECT.left = X
  DrawText hDC, text, Len(text), RECT, DT_NOPREFIX Or Win32.DT_RIGHT Or DT_SINGLELINE


End Sub

Sub PrintCJ(ByVal hDC As Long, X, Y, ByVal text As String)

  Dim RECT As Win32.RECT
  RECT.top = Y
  RECT.left = X
  DrawText hDC, text, Len(text), RECT, DT_NOPREFIX Or Win32.DT_CENTER Or DT_SINGLELINE

End Sub

Sub PrintCentered(ByVal hDC As Long, Y, ByVal text As String)

  Dim RECT As Win32.RECT
  RECT.top = Y
  RECT.left = GetDeviceCaps(hDC, HORZRES) / 2
  DrawText hDC, text, Len(text), RECT, DT_NOPREFIX Or Win32.DT_CENTER Or DT_SINGLELINE
End Sub


Sub FormFeed()
End Sub

Sub Fini()

End Sub

Private Sub Form_Load()
  SetTranslucent Me.hwnd, 230
  Me.left = 1200
  Me.top = 10
  Connect
  Me.Visible = True
  picReport.Visible = True
  Fill
  ResetActivityTime
End Sub
Public Sub Fill()
  picReport.CLS
  DoReport "", ""
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    picReport.left = Me.ScaleLeft
    picReport.top = Me.ScaleTop
    picReport.Width = Me.ScaleWidth
    picReport.Height = Me.ScaleHeight
  End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  If Not rsalarms Is Nothing Then
    rsalarms.Close
    Set rsalarms = Nothing
  End If
End Sub

