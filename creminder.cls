VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public reminderid       As Long

Public OwnerID          As Long  ' who owns this (either coordinator or resident)

Public IsPublic         As Integer ' 0 = no , 1 = yes

Public ReminderName     As String
Public ReminderMessage  As String

Public Coordinator      As String

Public LeadTime         As Long    ' minutes

Public Disabled         As Integer
Public Cancelled        As Integer

Public Recurring        As Integer

Public Frequency        As Long     ' one time, daily, weekly, monthly

Public DayString        As String   ' list of days
Public DAYS             As Collection

Public SpecificDay      As String

Public TimeofDay        As Long    ' time of day in HHMM (00 to 24, 0 to 59)

Public TimeHours        As Long
Public TimeMinutes      As Long



Private msubscribers    As Collection
Private mPagers         As Collection

Public FirstName        As String
Public LastName         As String

Public DOM              As Long
Public DOW              As Long

Public DeliveryPointID  As Long

Public DeliveryPoint    As cDeliveryPoint

' NOTE: need a SubscriberID <--> ReminderID table

Public NextReminderDue  As Date  ' dynamic value



Public Property Get TimeOfDayToString() As String
    Dim ampm As String
    Dim Clock As String
    
    
    If TimeHours = 0 Then
      If TimeMinutes = 0 Then
        TimeOfDayToString = "Midnight"
      Else
        TimeOfDayToString = "12" & ":" & Format$(TimeMinutes, "00") & " AM"
      End If
    
    ElseIf TimeHours < 12 Then

      
      TimeOfDayToString = TimeHours & ":" & Format$(TimeMinutes, "00") & " AM"
    ElseIf TimeHours = 12 Then
      If TimeMinutes = 0 Then
        TimeOfDayToString = "12 Noon"
      Else
        TimeOfDayToString = "12" & ":" & Format$(TimeMinutes, "00") & " PM"
      End If
    
    Else
      
    
    TimeOfDayToString = TimeHours - 12 & ":" & Format$(TimeMinutes, "00") & " PM"
    End If
    
    
    
  
End Property


Public Property Get DOWToString() As String
  Const DAYS = "SMTWTFS"
  
  Dim j       As Long
  Dim result  As String
  
  For j = 0 To 6
    If DOW And (2 ^ j) Then
      result = result & mid$(DAYS, j + 1, 1)
    Else
      result = result & "_"
    End If
  Next
  DOWToString = result

End Property

Public Property Get DOMToString() As String
    If (DOM > 0 And DOM <= 31) Then
      DOMToString = Format$(DOM, "0")
      Select Case Right$(DOMToString, 1)
        Case 1
          DOMToString = DOMToString & "st"
        Case 2
          DOMToString = DOMToString & "nd"
        Case 3
          DOMToString = DOMToString & "rd"
        Case 4, 5, 6, 7, 8, 9, 0
          DOMToString = DOMToString & "th"
        Case Else
        
      End Select
    Else
      DOMToString = "Error"
    End If
        
End Property


Public Property Get ScheduleToString() As String
  
  
  Select Case Frequency
    Case REMINDER_DATE
       ScheduleToString = Format$(SpecificDay, "mm/dd/yy")
    Case REMINDER_DAILY
       ScheduleToString = DOWToString()
    Case REMINDER_MONTHLY
       ScheduleToString = DOMToString()
    Case Else
       ScheduleToString = ""
  End Select
  
End Property

Public Property Get FrequencyToString() As String
  
  Select Case Frequency
    Case REMINDER_DATE
      FrequencyToString = "1"
    Case REMINDER_DAILY
      FrequencyToString = "D"
    Case REMINDER_MONTHLY
      FrequencyToString = "M"
    Case Else
      FrequencyToString = ""
  End Select
End Property
Public Sub ClearSubscribers()
  Set subscribers = New Collection
End Sub
Public Sub ClearPagers()
  Set Pagers = New Collection
End Sub
Public Function GetPagerDest()

End Function

Public Function GetDeliveryPoint() As cDeliveryPoint
  Dim rs                As ADODB.Recordset
  Dim SQL               As String

  Dim j                 As Integer
  Dim rows()            As String
  Dim cells()           As String
  Dim Status            As Long
  Dim DeliveryPoints    As String

  Set DeliveryPoint = New cDeliveryPoint
  If IsPublic Then
    SQL = "Select deliverypoints FROM Staff WHERE staffID = " & OwnerID
  Else
    SQL = "Select deliverypoints FROM residents WHERE residentID = " & OwnerID
  End If

  Set rs = ConnExecute(SQL)
  If Not rs.EOF Then
    DeliveryPoints = rs("deliverypoints") & ""
    On Error Resume Next
    rows = Split(DeliveryPoints, vbCrLf, 3)
    For j = LBound(rows) To UBound(rows)
      cells = Split(rows(j), vbTab)
      If j = DeliveryPointID Then  ' public
        DeliveryPoint.Status = Val(cells(0))
        DeliveryPoint.Address = Trim$(cells(1))
        DeliveryPoint.AddressType = Trim$(cells(2))
        Exit For
      End If
    Next
  End If
  rs.Close
  Set rs = Nothing
 Set GetDeliveryPoint = DeliveryPoint
End Function

'Public Function GetPublicDeliveryPoint() As cDeliveryPoint
'
'  Dim j As Integer
'  Dim rows()    As String
'  Dim cells()   As String
'  Dim Status As Long
'
'
'  Set mPublicDeliveryPoint = New cDeliveryPoint
'  On Error Resume Next
'  rows = Split(mDeliveryPoints, vbCrLf, 3)
'  For j = LBound(rows) To UBound(rows)
'    cells = Split(rows(j), vbTab)
'    Status = Val(cells(0))
'    If Status = DELIVERY_POINT_STATUS.BOTH_STATUS Or Status = DELIVERY_POINT_STATUS.PUBLIC_STATUS Then     ' public
'      mPublicDeliveryPoint.Status = Status
'      mPublicDeliveryPoint.Address = Trim$(cells(1))
'      mPublicDeliveryPoint.AddressType = Trim$(cells(2))
'      Exit For
'    End If
'  Next
'
'  Set GetPublicDeliveryPoint = mPublicDeliveryPoint
'
'End Function



Public Function Parse(rs As ADODB.Recordset) As Long
  reminderid = rs("reminderid")
  OwnerID = rs("OwnerID")
  IsPublic = rs("ispublic")
  ReminderName = rs("description") & ""
  ReminderMessage = rs("message") & ""
  Coordinator = rs("coordinator") & ""
  LeadTime = rs("LeadTime")
  Disabled = rs("disabled")
  Cancelled = rs("cancelled")
  Recurring = rs("Recurring")
  Frequency = rs("Frequency")
  DayString = rs("daysactive") & ""
  
  
  DOW = rs("DOW")
  DOM = rs("DOM")
  
  'Days = parsedaysstring
  
  SpecificDay = rs("specificday") & ""
  TimeofDay = rs("Timeofday")
  TimeHours = TimeofDay \ 100
  TimeMinutes = (TimeofDay - (TimeHours * 100))
  'sanity check
  If TimeHours > 23 Then TimeHours = 23
  If TimeHours < 0 Then TimeHours = 0
  If TimeMinutes > 59 Then TimeMinutes = 59
  If TimeMinutes < 0 Then TimeMinutes = 0
  
  
   
  DeliveryPointID = Val(rs("deliverypoint") & "")
   
   
   
End Function

Public Function UpdateData(rs As ADODB.Recordset)

  'rs("reminderid") = ReminderID
  
  rs("OwnerID") = OwnerID
  rs("ispublic") = IIf(IsPublic, 1, 0)
  rs("residentid") = 0
  rs("description") = ReminderName
  rs("message") = ReminderMessage
  rs("coordinator") = Coordinator
  rs("LeadTime") = LeadTime
  rs("disabled") = IIf(Disabled, 1, 0)
  rs("cancelled") = IIf(Cancelled, 1, 0)
  rs("Recurring") = IIf(Recurring, 1, 0)
  rs("Frequency") = Frequency
  rs("daysactive") = DayString
  
  rs("DOW") = DOW
  rs("DOM") = DOM
  
  'Days = parsedaysstring
  
  rs("specificday") = SpecificDay
  
  If TimeHours > 23 Then TimeHours = 23
  If TimeHours < 0 Then TimeHours = 0
  If TimeMinutes > 59 Then TimeMinutes = 59
  If TimeMinutes < 0 Then TimeMinutes = 0
  rs("Timeofday") = (TimeHours * 100) + TimeMinutes
  
  rs("deliverypoint") = DeliveryPointID


End Function


Private Sub Class_Initialize()
  Set DAYS = New Collection
  Set subscribers = New Collection
  Set Pagers = New Collection
  Set DeliveryPoint = New cDeliveryPoint
End Sub

Sub Finalize()
  Set DAYS = Nothing
  Set subscribers = Nothing

End Sub

Private Sub Class_Terminate()
  Finalize
  Set msubscribers = Nothing
End Sub

Public Property Get subscribers() As Collection
  Set subscribers = msubscribers
End Property

Public Property Set subscribers(subscribers As Collection)
  Set msubscribers = subscribers
End Property

Public Property Get Pagers() As Collection
  Set Pagers = mPagers
End Property

Public Property Set Pagers(Pagers As Collection)
  Set mPagers = Pagers
End Property


Public Property Get OwnerName() As String
  Dim Bothnames As String
  
  Bothnames = Len(LastName) > 0 And Len(FirstName) > 0
  OwnerName = LastName & IIf(Bothnames, ", ", "") & FirstName
  
  

End Property

