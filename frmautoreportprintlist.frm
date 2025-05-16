VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmAutoReportPrintList 
   Caption         =   "Auto Report Print List"
   ClientHeight    =   11220
   ClientLeft      =   2190
   ClientTop       =   4755
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   748
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   757
   Begin VB.Frame fraEnabler 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Rooms"
      Height          =   10845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10425
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7950
         TabIndex        =   13
         Top             =   390
         Width           =   1175
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7935
         TabIndex        =   2
         Top             =   1725
         Width           =   1175
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7935
         TabIndex        =   1
         Top             =   2310
         Width           =   1175
      End
      Begin VB.Frame fraPrinter 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   30
         TabIndex        =   3
         Top             =   3120
         Width           =   9165
         Begin VB.TextBox txtFilePath 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   120
            Width           =   5535
         End
         Begin VB.TextBox txtCurrentPrinter 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   450
            Width           =   4455
         End
         Begin MSComctlLib.ListView lvPrinters 
            Height          =   1635
            Left            =   330
            TabIndex        =   6
            Top             =   780
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   2884
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label lblCurrPath 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Folder:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   14
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label lblCurrPrn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Printer:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   12
            Top             =   450
            Width           =   1305
         End
      End
      Begin VB.Frame fraFolder 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   30
         TabIndex        =   4
         Top             =   5880
         Width           =   9135
         Begin VB.DirListBox lstFolders 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3990
         End
         Begin VB.DriveListBox lstDrives 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   10
            Top             =   2130
            Width           =   3990
         End
         Begin VB.CommandButton cmdNewFolder 
            Caption         =   "New Folder"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            TabIndex        =   9
            Top             =   465
            Width           =   1500
         End
         Begin VB.TextBox txtNewFolder 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4110
            TabIndex        =   8
            Top             =   45
            Width           =   3135
         End
      End
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   2985
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   5265
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Printer"
               Key             =   "printer"
               Object.ToolTipText     =   "Choose Printer"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Output Folder"
               Key             =   "folder"
               Object.Tag             =   "folder"
               Object.ToolTipText     =   "Destination Folder"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAutoReportPrintList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastDrive As String

Function PrintAutoReportsList() As Boolean
  Dim hfile As Integer
  Dim html As String
  Dim AutoReports As Collection: Set AutoReports = New Collection
  Dim SQl As String
  Dim rs As ADODB.Recordset
  Dim Report As cAutoReport
  Dim AlarmEvent As cDataWrapper
  
  Const LeftMargin = 0.5
  Const Indent = 0.5
  Const col1 = 1#
  Const col1_5 = 1.5
  Const col2 = 2#
  Const col3 = 3#
  
  SQl = "SELECT * FROM AutoReports ORDER BY reportname"

  Set AutoReports = New Collection
  Set rs = ConnExecute(SQl)
  Do Until rs.EOF
    Set Report = New cAutoReport
    Report.Parse rs
    AutoReports.Add Report
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing

  
  hfile = FreeFile
  Open Configuration.AutoReportsListFolder & "\AutoReportsList.html" For Output As #hfile

  html = "<html>"
  html = html & "<head>"
  html = html + "<style type=""text/css"">"
  html = html + "body {width:900px; font-family:arial,verdana,sans-serif;}"
  html = html + "table.main {width:900px;font-size:1em;table-layout:fixed;}"
  html = html + "table.sub {width:400px;font-size:1em;table-layout:fixed;}"
  html = html + "tr.header td {background-color: #ADD8E6; color: black; margin:0px; padding:0px; font-weight:bold;}"
  html = html + "td.title {color: black; margin:0px; padding:0px;font-size:1em; font-weight:bold;width:150px;}"
  html = html + "td.data {color: black; margin:0px; padding:0px;font-size:1em; font-weight:bold;width:100px;}"
  html = html + "tr.even td {background-color: #FAFAD2; color: black; margin:0px; padding:0px;}"
  html = html + "tr.odd td {background-color: white; color: black; margin:0px; padding:0px;}"
  html = html + "h1 {background-color: white; color: black;margin:5px;text-align:left;font-size:0.8em;}"
  html = html + "h2 {background-color: white; color: black;margin:5px;text-align:left;font-size:0.8em;}"
  html = html + "p.complete {background-color: white; color:gray;margin:5px;text-align:left;font-size:0.8em;}"
  
  html = html + "</style>"

  html = html & "</head>" & vbCrLf
  html = html & "<body>" & vbCrLf
  
  html = html & "<h1>Automatic Reports</h1><br/>" & vbCrLf
  
  html = html & "<table class='main'> " & vbCrLf
  html = html & "<tr>" & HTMLTD(Now) & "</tr>" & vbCrLf
  html = html & "</table><br/><hr><br/>" & vbCrLf
  
  For Each Report In AutoReports
    DoEvents
    html = html & "<table class='main'>" & vbCrLf
    html = html & "<tr>" & HTMLTD("Report Name:", "class='title' ") & HTMLTD(Report.ReportName) & "</tr>" & vbCrLf
    html = html & "<tr>" & HTMLTD("Comment:", "class='title' ") & HTMLTD(Report.Comment) & "</tr>" & vbCrLf
    'html = html & "<tr>" & HTMLTD("<hr>", "colspan='2' ") & "</tr>" & vbCrLf
    
    html = html & "<tr>" & HTMLTD("Schedule:", "class='title' ") & HTMLTD(PeriodToString(Report.DayPeriod)) & "</tr>" & vbCrLf
    
    Select Case Report.DayPeriod
      Case AUTOREPORT_DAILY:
        html = html & "<tr>" & vbCrLf
        html = html & "<td> </td><td>" & vbCrLf
        html = html & "<table class='sub'>" & vbCrLf
        html = html & "<tr>" & HTMLTD("Start Time:", "class='title'") & HTMLTD(DaypartToString(Report.DayPartStart)) & "</tr>" & vbCrLf
        html = html & "<tr>" & HTMLTD("End Time:", "class='title' ") & HTMLTD(DaypartToString(Report.DayPartEnd)) & "</tr>" & vbCrLf
        html = html & "<tr>" & HTMLTD("Days:", "class='title' ") & HTMLTD(DaysToString(Report.DAYS)) & "</tr>" & vbCrLf
        html = html & "</table>" & vbCrLf
        html = html & "</td>" & vbCrLf
        html = html & "</tr>" & vbCrLf
    End Select
   ' html = html & "<tr>" & HTMLTD("<hr>", "colspan='2' ") & "</tr>" & vbCrLf
    html = html & "<tr>" & HTMLTD("Events", "class='title'") & HTMLTD("") & "</tr>" & vbCrLf
    
    
    For Each AlarmEvent In Report.Events
        html = html & "<tr>"
        html = html & "<td> </td><td>" & vbCrLf
        html = html & "<table class='sub'>" & vbCrLf
        html = html & "<tr>" & HTMLTD("Event:", "class='title'") & HTMLTD(EventToString(AlarmEvent.LongValue)) & "</tr>" & vbCrLf
        html = html & "</table>" & vbCrLf
        html = html & "</td>"
        html = html & "</tr>" & vbCrLf
    Next
    'html = html & "<tr>" & HTMLTD("<hr>", "colspan='2' ") & "</tr>" & vbCrLf
    html = html & "<tr>" & HTMLTD("Send as Email:", "class='title'") & HTMLTD(IIf(Report.SendAsEmail, "Yes", "No")) & "</tr>" & vbCrLf
    If Report.SendAsEmail Then
        html = html & "<tr>"
        html = html & "<td> </td><td>" & vbCrLf
        html = html & "<table class='sub'>" & vbCrLf
        html = html & "<tr>" & HTMLTD("Email Recipient:", "class='title'") & HTMLTD(Report.recipient) & "</tr>" & vbCrLf
        html = html & "<tr>" & HTMLTD("Email Subject:", "class='title'") & HTMLTD(Report.Subject) & "</tr>" & vbCrLf
        html = html & "</table>" & vbCrLf
        html = html & "</td>"
        html = html & "</tr>" & vbCrLf
      
    End If
    'html = html & "<tr>" & HTMLTD("<hr>", "colspan='2' ") & "</tr>" & vbCrLf
    html = html & "<tr>" & HTMLTD("File Format:", "class='title'") & HTMLTD(FileFormatToString(Report.FileFormat)) & "</tr>" & vbCrLf
    'html = html & "<tr>" & HTMLTD("<hr>", "colspan='2' ") & "</tr>" & vbCrLf
    html = html & "</table>" & vbCrLf
   html = html & "<br /><hr><br /><br />" & vbCrLf
  Next


  html = html & "</body>" & vbCrLf
  html = html & "</html>" & vbCrLf

  Print #hfile, html
  Close hfile
  

  Dim j As Integer

  Dim p As Printer
  Dim Oldprinter As Printer
  Set Oldprinter = Printer

  Dim OldPrinterName As String: OldPrinterName = Printer.DeviceName

  Dim textht As Single
  
  Dim oldy As Single

  For j = 0 To Printers.Count - 1
    If 0 = StrComp(Printers(j).DeviceName, Configuration.AutoReportsListPrinter, vbTextCompare) Then
      Set Printer = Printers(j)
      Exit For
    End If
  Next


  ' header
  Printer.ScaleMode = vbInches
  Printer.FontSize = 10
  textht = Printer.TextHeight("T")
  modPrint.PrintLineFeed Printer
  Printer.FontBold = True
  oldy = Printer.CurrentY
  modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Automatic Reports "
  Debug.Print Printer.CurrentY - oldy, textht
  Printer.FontBold = False
  Printer.CurrentY = Printer.CurrentY + textht * 1.25
  
  
  modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, Now
  Printer.CurrentY = Printer.CurrentY + textht * 1.25
  Printer.Print
  modPrint.HR Printer
  
  For Each Report In AutoReports
    Printer.Print
    DoEvents
    If (Printer.CurrentY / Printer.ScaleHeight) > 0.75 Then
      Printer.CurrentY = Printer.CurrentY + textht * 1.25
      modPrint.PrintCentered Printer, Printer.CurrentY, "-continued on page " & Printer.Page + 1 & "-"
      modPrint.FormFeed Printer
      Printer.CurrentY = Printer.CurrentY + textht * 1.25
      modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Page " & Printer.Page
      Printer.Print
      Printer.Print
      modPrint.HR Printer
      Printer.Print
      
      
    End If
    Printer.FontBold = True
    modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Report Name: "
    Printer.FontBold = False
    modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, Report.ReportName
    Printer.CurrentY = Printer.CurrentY + textht * 1.25
    Printer.FontBold = True
    
    modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Comment: "
    Printer.FontBold = False
    modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, Report.Comment
    Printer.CurrentY = Printer.CurrentY + textht * 1.25
    Printer.FontBold = True
    modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Schedule: "
    Printer.FontBold = False
    modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, PeriodToString(Report.DayPeriod)
    Printer.CurrentY = Printer.CurrentY + textht * 1.25
    
    Select Case Report.DayPeriod
      Case AUTOREPORT_DAILY:
        Printer.FontBold = True
        modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "Start Time: "
        Printer.FontBold = False
        modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, DaypartToString(Report.DayPartStart)
        Printer.CurrentY = Printer.CurrentY + textht * 1.25
        Printer.FontBold = True
        modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "End Time: "
        Printer.FontBold = False
        modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, DaypartToString(Report.DayPartEnd)
        Printer.CurrentY = Printer.CurrentY + textht * 1.25
        Printer.FontBold = True
        modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "Days: "
        Printer.FontBold = False
        modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, DaysToString(Report.DAYS)
        Printer.CurrentY = Printer.CurrentY + textht * 1.25
    End Select
   
    Printer.FontBold = True
    modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Events "
    Printer.FontBold = False
    Printer.CurrentY = Printer.CurrentY + textht * 1.25
    For Each AlarmEvent In Report.Events
          Printer.FontBold = True
          modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "Event: "
          Printer.FontBold = False
          modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, EventToString(AlarmEvent.LongValue)
          Printer.CurrentY = Printer.CurrentY + textht * 1.25
          
    Next
    
    Printer.FontBold = True
    modPrint.PrintLJ Printer, LeftMargin, Printer.CurrentY, "Send as Email: "
    Printer.FontBold = False
       modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, IIf(Report.SendAsEmail, "Yes", "No")
    Printer.CurrentY = Printer.CurrentY + textht * 1.25
    If Report.SendAsEmail Then
      Printer.FontBold = True
      modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "Email Recipient: "
      Printer.FontBold = False
        modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, Report.recipient
      Printer.CurrentY = Printer.CurrentY + textht * 1.25
      Printer.FontBold = True
      modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "Email Subject: "
      Printer.FontBold = False
        modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, Report.Subject
      Printer.CurrentY = Printer.CurrentY + textht * 1.25
    End If
      Printer.FontBold = True
      modPrint.PrintLJ Printer, LeftMargin + Indent, Printer.CurrentY, "File Format: "
      Printer.FontBold = False
      modPrint.PrintLJ Printer, Printer.CurrentX, Printer.CurrentY, FileFormatToString(Report.FileFormat)
      Printer.Print
      Printer.Print
      modPrint.HR Printer

      
  Next
      Printer.Print
      modPrint.PrintCentered Printer, Printer.CurrentY, "End of Report"
  
  Printer.EndDoc

  If Len(OldPrinterName) > 0 Then
  For j = 0 To Printers.Count - 1
    If 0 = StrComp(Printers(j).DeviceName, OldPrinterName, vbTextCompare) Then
      Set Printer = Printers(j)
      Exit For
    End If
  Next
  End If

End Function
'Function DaypartToString(ByVal DayPart As Long) As String
'  If DayPart = 0 Then
'    DaypartToString = "12 Midnight"
'  ElseIf DayPart = 12 Then
'    DaypartToString = "12 Noon"
'  ElseIf DayPart > 12 Then
'    DaypartToString = DayPart - 12 & " PM"
'  Else
'    DaypartToString = DayPart & " AM"
'  End If
'
'
'
'
'End Function
'
'
'Function DaysToString(ByVal DAYS As Long) As String
'
'
'  Const DayList = "SMTWTFS"
'
'  Dim j       As Long
'  Dim result  As String
'
'  For j = 0 To 6
'    If DAYS And (2 ^ j) Then
'      result = result & mid$(DayList, j + 1, 1)
'    Else
'      result = result & "_"
'    End If
'  Next
'  DaysToString = result
'
'
'
'End Function

'Public Function EventToString(ByVal EventID As Long) As String
'  Select Case EventID
'    Case EVT_EMERGENCY
'      EventToString = "Alarms"
'    Case EVT_ALERT
'      EventToString = "Alerts"
'    Case EVT_BATTERY_FAIL
'      EventToString = "Low Battery"
'    Case EVT_CHECKIN_FAIL
'      EventToString = "Trouble"
'    Case EVT_TAMPER
'      EventToString = "Tamper"
'    Case EVT_EXTERN
'      EventToString = "External"
'    Case Else
'      EventToString = "Unspecified"
'  End Select
'End Function

'Function FileFormatToString(ByVal Format As Long) As String
'  Select Case Format
'    Case AUTOREPORTFORMAT_TAB_NOHEADER
'      FileFormatToString = "Tab Delimited, No Headers"
'    Case AUTOREPORTFORMAT_HTML
'      FileFormatToString = "HTML"
'    Case Else '      AUTOREPORTFORMAT_TAB
'      FileFormatToString = "Tab Delimited"
'    End Select
'
'End Function
'
'Function PeriodToString(ByVal period As Long) As String
'  Select Case period
'    Case AUTOREPORT_DAILY: PeriodToString = "Daily"
'    Case AUTOREPORT_SHIFT1: PeriodToString = "Day Shift"
'    Case AUTOREPORT_SHIFT2: PeriodToString = "Night Shift"
'    Case AUTOREPORT_WEEKLY: PeriodToString = "Weekly"
'    Case AUTOREPORT_MONTHLY: PeriodToString = "Monthly"
'    Case Else: PeriodToString = "Unspecified"
'  End Select
'End Function


Private Sub cmdApply_Click()
  ResetActivityTime
  Save
  
End Sub
Sub Save()
  Configuration.AutoReportsListFolder = lstFolders.path
  
  If Not (lvPrinters.SelectedItem Is Nothing) Then
    Configuration.AutoReportsListPrinter = lvPrinters.SelectedItem.text
    WriteSetting "Configuration", "AutoReportsListPrinter", Configuration.AutoReportsListPrinter
  End If
  
  WriteSetting "Configuration", "AutoReportsListFolder", Configuration.AutoReportsListFolder
  Fill
  txtCurrentPrinter.text = Configuration.AutoReportsListPrinter
End Sub

Private Sub cmdExit_Click()
  PreviousForm
  Unload Me
End Sub

Private Sub cmdNewFolder_Click()
  Dim folder As String
  Dim path As String
  
  path = lstFolders.path
  If Right$(path, 1) <> "\" Then
    path = path & "\"
  End If
  
  folder = Trim$(txtNewFolder.text)
  Do While (left$(folder, 1) = "\") And (Len(folder) > 0)
    folder = Right$(folder, Len(folder) - 1)
  Loop
  
  If Len(folder) > 0 Then
    On Error Resume Next
    MkDir path & folder
 
  End If
  lstFolders.Refresh
End Sub

Private Sub cmdPrint_Click()
  ResetActivityTime
  If Printer Is Nothing Then Exit Sub
  
  DisableButtons
  PrintAutoReportsList
  EnableButtons
  
End Sub
Sub DisableButtons()
  cmdApply.Enabled = False
  cmdExit.Enabled = False
  cmdPrint.Enabled = False
End Sub
Sub EnableButtons()
  cmdApply.Enabled = True
  cmdExit.Enabled = True
  cmdPrint.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHost
End Sub

Public Sub Host(ByVal hwnd As Long)

  fraEnabler.left = 0
  fraEnabler.top = 0
  fraEnabler.Width = ENABLER_WIDTH
  fraEnabler.Height = ENABLER_HEIGHT
  SetParent fraEnabler.hwnd, hwnd
  fraEnabler.BackColor = Me.BackColor
End Sub
Public Sub UnHost()
  'PreviousForm
  SetParent fraEnabler.hwnd, Me.hwnd
End Sub

Private Sub Form_Load()
  ResetActivityTime
  Me.ScaleMode = vbPixels
  ConfigureLists
  ArrangeControls
  FillPrinters
  FillFolders
End Sub
Private Sub TabStrip_Click()
  SetTabs
End Sub
Sub ConfigureLists()
  lvPrinters.ColumnHeaders.Clear
  lvPrinters.ColumnHeaders.Add , , "Printer", 250 * Screen.TwipsPerPixelX
  
End Sub

Public Sub Fill()
  'txtCurrentPrinter.text = Printer.DeviceName
  FillPrinters
  FillFolders
  txtFilePath.text = lstFolders.path
End Sub
Sub FillFolders()
  On Error Resume Next
  lstDrives.Drive = Configuration.AutoReportsListFolder
  lstFolders.path = Configuration.AutoReportsListFolder
  
End Sub

Sub SetTabs()
  Select Case TabStrip.SelectedItem.Key
    Case "printer"
      fraPrinter.Visible = True
      fraFolder.Visible = False
      
    
    Case Else
      fraFolder.Visible = True
      fraPrinter.Visible = False
      

  End Select

End Sub
Sub ArrangeControls()

  fraEnabler.BackColor = Me.BackColor

  fraFolder.left = TabStrip.ClientLeft
  fraFolder.top = TabStrip.ClientTop
  fraFolder.Height = TabStrip.ClientHeight
  fraFolder.Width = TabStrip.ClientWidth
  fraFolder.BackColor = Me.BackColor

  fraPrinter.left = TabStrip.ClientLeft
  fraPrinter.top = TabStrip.ClientTop
  fraPrinter.Height = TabStrip.ClientHeight
  fraPrinter.Width = TabStrip.ClientWidth
  fraPrinter.BackColor = Me.BackColor


  SetTabs


End Sub

Sub FillPrinters()
  Dim p As Printer
  Dim li As ListItem
  Dim index As Long
  Dim j As Integer
  
  Dim CurrentPrinter As String
  Dim ActivePrinter  As String
  
  CurrentPrinter = Configuration.AutoReportsListPrinter
  
  lvPrinters.ListItems.Clear
  
  For j = 0 To Printers.Count - 1
    Set p = Printers(j)
    Set li = lvPrinters.ListItems.Add(, index & "s", p.DeviceName)
    If 0 = StrComp(CurrentPrinter, p.DeviceName, vbTextCompare) Then
      ActivePrinter = p.DeviceName
      li.Selected = True
    End If
    
    index = index + 1
  Next
    
  If ActivePrinter = "" Then ' printer is MIA
    CurrentPrinter = Printer.DeviceName
  Else
    CurrentPrinter = ActivePrinter ' OK printer
  End If
  
  txtCurrentPrinter.text = CurrentPrinter

  ShowSelectedPrinter CurrentPrinter

End Sub
Sub ShowSelectedPrinter(ByVal PrinterName As String)
  Dim li As ListItem
  
  For Each li In lvPrinters.ListItems
    If 0 = StrComp(PrinterName, li.text, vbTextCompare) Then
      li.Selected = True
      Exit For
    Else
      li.Selected = False
    End If
  Next
End Sub

Private Sub lstDrives_Change()
  Dim Retry As Boolean
  On Error Resume Next

  Retry = True
  Do While Retry
    Retry = False
    lstFolders.path = lstDrives.Drive
    ' If and error occurs
    Select Case Err.Number
      Case 68  ' Not accessable Error
        If vbRetry = messagebox(Me, lstDrives.Drive & " is not accessible", App.Title, vbRetryCancel + vbCritical) Then
          Retry = True
        Else  'Switch to previous known drive
          lstDrives.Drive = LastDrive
        End If
      Case 0
        ' done
      Case Else
        ' Ooops
         messagebox Me, "Unexpected File Access Error " & Err.Number & " : " & Err.Description, App.Title, vbInformation
        lstDrives.Drive = LastDrive
    End Select
  Loop



End Sub

Private Sub lstFolders_Change()
  Static Busy As Boolean

  On Error Resume Next

  If Not Busy Then
  
    Busy = True
    ' Change the current directory
    ChDir lstFolders.path

    If Err.Number = 0 Then
      lstDrives.Drive = left(lstFolders.path, 2)
    Else
      Err.Clear
    End If

    Busy = False

  End If

End Sub
