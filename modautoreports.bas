Attribute VB_Name = "modAutoreports"
Option Explicit

Global Const AUTOREPORT_DAILY = 0
Global Const AUTOREPORT_SHIFT1 = 1
Global Const AUTOREPORT_SHIFT2 = 2
Global Const AUTOREPORT_SHIFT3 = 3

Global Const AUTOREPORT_WEEKLY = 7
Global Const AUTOREPORT_MONTHLY = 30



Global Const AUTOREPORT_SORT_ROOM = 0
Global Const AUTOREPORT_SORT_ELAPSED = 1 ' in longest to shortest
Global Const AUTOREPORT_SORT_CHRONO = 2

Global Const AUTOREPORTFORMAT_TAB = 0
Global Const AUTOREPORTFORMAT_TAB_NOHEADER = 1 ' in longest to shortest
Global Const AUTOREPORTFORMAT_HTML = 2

Global gAutoReports As Collection
Global gAutoExReports As Collection


Sub LoadAutoExReports()

  Dim SQl       As String
  Dim Report    As cExceptionAutoReport
  Dim rs        As ADODB.Recordset

   On Error GoTo LoadAutoExReports_Error
   

10      If MASTER Then
20        SQl = "SELECT * FROM ExceptionReports WHERE Disabled <> 1"
30        Set gAutoExReports = New Collection
40        Set rs = ConnExecute(SQl)
50        Do Until rs.EOF
60          Set Report = New cExceptionAutoReport
70          Report.Parse rs
75          gAutoExReports.Add Report
80          rs.MoveNext
90        Loop
100       rs.Close
110       Set rs = Nothing

120     End If

LoadAutoExReports_Resume:
   On Error GoTo 0
   Exit Sub

LoadAutoExReports_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAutoReports.LoadAutoExReports." & Erl
  Resume LoadAutoExReports_Resume


End Sub




Sub LoadAutoReports()
  Dim SQl       As String
  Dim Report    As cAutoReport
  Dim rs        As ADODB.Recordset

   On Error GoTo LoadAutoReports_Error
   

10      If MASTER Then
20        SQl = "SELECT * FROM Autoreports WHERE Disabled <> 1"
30        Set gAutoReports = New Collection
40        Set rs = ConnExecute(SQl)
50        Do Until rs.EOF
60          Set Report = New cAutoReport
70          Report.Parse rs
75          gAutoReports.Add Report
80          rs.MoveNext
90        Loop
100       rs.Close
110       Set rs = Nothing

120     End If

LoadAutoReports_Resume:
   On Error GoTo 0
   Exit Sub

LoadAutoReports_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modAutoReports.LoadAutoReports." & Erl
  Resume LoadAutoReports_Resume


End Sub



Public Function CheckIfReportsDue() As Boolean
  Dim Report As cAutoReport
  Dim ExReport As cExceptionAutoReport

  If MASTER Then
    If Not gAutoReports Is Nothing Then
  
    For Each Report In gAutoReports
      If Report.due Then
        Report.DoReport
        Exit For ' do one at at time
      End If
    Next
    End If
    
    If Not gAutoExReports Is Nothing Then
    For Each ExReport In gAutoExReports ' do the exception reports
      If ExReport.due Then
        ExReport.DoReport
        Exit For ' do one at at time
      End If
    Next
    End If
    
  End If


End Function
Public Function DeleteAutoReport(ByVal ReportID As Long) As Long
  Dim Report As cAutoReport
  Dim j As Integer
  For j = 1 To gAutoReports.count
    Set Report = gAutoReports(j)
    If (Report.ReportID = ReportID) Then
      gAutoReports.Remove j
      Exit For
    End If
  Next
End Function

Public Function DeleteAutoExReport(ByVal ReportID As Long) As Long
  Dim Report As cExceptionAutoReport
  Dim j As Integer
  For j = 1 To gAutoExReports.count
    Set Report = gAutoExReports(j)
    If (Report.ReportID = ReportID) Then
      gAutoExReports.Remove j
      Exit For
    End If
  Next
End Function

