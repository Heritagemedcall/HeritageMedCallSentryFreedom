Attribute VB_Name = "modReminders"
Option Explicit

Global Const REMINDER_INACTIVE = 0
Global Const REMINDER_DATE = 1
Global Const REMINDER_DAILY = 2
Global Const REMINDER_MONTHLY = 4

Global Const NO_REMINDERS = False


Global gTTSEngine        As cSpeechEngine ' this is the global speech enging for converting TTP to files
Global TTSVoices         As Collection

Global gRemindersToSend     As Collection
Global EmailQue             As Collection
Global VoiceMailQue         As Collection

Global LeadTime As Long

Private ReminderTickCount As Long
Private NextReminderCheck As Date

Sub ReminderAnnouncement()

  Dim text      As String
  Dim PagerID   As Long
  Dim GroupID   As Long
  On Error Resume Next

  text = Trim(text)
  SendToPager text, PagerID, 0, "", "", PAGER_NORMAL, left$(text, 19), 0, 0
  SendToGroup text, GroupID, "", "", PAGER_NORMAL, left$(text, 19), 0, 0
End Sub




Sub CheckDialogics()
  If Not MASTER Then Exit Sub
  
  If Not (DialogicSystem Is Nothing) Then
    DialogicSystem.UpdateClock
  End If

End Sub

Sub ProcessVoicemail()
  ' call this from reminder message loop timer

  Dim Voicemail     As cVoiceMail
  Dim AvailableChannel As Long
  Dim callhandler   As cDivaCall

  If Not MASTER Then Exit Sub

  If DialogicSystem Is Nothing Then
    'LogProgramError "Dialogic System Not Loaded"
    Exit Sub
  End If

  ' get open voicemail dialer channel
  If VoiceMailQue.Count > 0 Then
    AvailableChannel = DialogicSystem.GetNextAvailable()
    dbgTAPI "ProcessVoicemail.AvailableChannel " & AvailableChannel
    If AvailableChannel > 0 Then
      If VoiceMailQue.Count > 0 Then


        Set Voicemail = VoiceMailQue(1)
        Debug.Print "DeQueing For " & Voicemail.Address & " on Channel " & AvailableChannel
        VoiceMailQue.Remove 1
        If (Len(Trim$(Voicemail.Address))) Then

          Set callhandler = DialogicSystem.GetCallHandlerByID(AvailableChannel)
          callhandler.Busy = True
          callhandler.MessageDelay = Configuration.ReminderMsgDelay
          callhandler.MessageRepeatDelay = Configuration.ReminderMsgSpacing
          callhandler.MessageRepeats = Configuration.ReminderMsgRepeats
          callhandler.RedialAttempts = Configuration.ReminderRedials
          callhandler.RedialWait = Configuration.ReminderRedialDelay

          callhandler.PhoneNumber = Voicemail.Address
          callhandler.message = Voicemail.message

          If Voicemail.AckDigit <> 0 Then
            callhandler.TerminateDigit = Configuration.ReminderAckDigit
          Else
            callhandler.TerminateDigit = 0
          End If
          'if callhandler.VoiceName <>
          'callhandler.VoiceName =
          callhandler.BeginCall  ' set and forget
          Set callhandler = Nothing

        End If
      End If
    End If
  End If


End Sub

Sub GetTTSVoices()
  Dim voice As spvoice
  Dim token As ISpeechObjectToken
  Dim ttsVoice As cTTSVoice

  Set voice = New spvoice
  Set TTSVoices = New Collection
  For Each token In voice.GetVoices
    Set ttsVoice = New cTTSVoice
    ttsVoice.VoiceName = token.GetDescription()
    ttsVoice.VoiceIndex = TTSVoices.Count  ' list is zero based
    TTSVoices.Add ttsVoice, ttsVoice.VoiceName
  Next

  Set token = Nothing
  Set ttsVoice = Nothing
  Set voice = Nothing

End Sub

Public Function TranslateDeliveryPointType(ByVal DeliveryPointType As Long) As String
  Const DELIVERY_POINT_NONE = "None"
  Const DELIVERY_POINT_PHONE = "Phone"
  Const DELIVERY_POINT_PHONE_ACK = "Phone ACK"
  Const DELIVERY_POINT_EMAIL = "Email"
  
  
  Select Case DeliveryPointType
    Case DELIVERY_POINT.EMAIL
      TranslateDeliveryPointType = DELIVERY_POINT_EMAIL
    Case DELIVERY_POINT.phone_ack
      TranslateDeliveryPointType = DELIVERY_POINT_PHONE_ACK
    Case DELIVERY_POINT.Phone
      TranslateDeliveryPointType = DELIVERY_POINT_PHONE
    Case Else
      TranslateDeliveryPointType = DELIVERY_POINT_NONE

  End Select


End Function

Function EmailReminder(ByVal Address As String, ByVal message As String)


        Dim mapi    As Object
        'Dim mapi As SENTRYMAIL.MAPITransport
        Dim Filenames As String

        Dim Subject As String
        Dim s       As String

        If Not MASTER Then Exit Function

          #If brookdale Then
10        s = "TechConnect"

          #ElseIf esco Then
20        s = "CareConnect"

          #Else
30        s = "Sentry"

          #End If


40      Subject = "Reminder from " & s



50      On Error Resume Next
60      If (Configuration.UseSMTP = MAIL_SMTP) Then
70        If gSMTPMailer Is Nothing Then
'80          Set gSMTPMailer = New SendMail
80           Set gSMTPMailer = CreateObject("smtpmailer.SendMail")
90        End If
100       If gSMTPMailer Is Nothing Then
110         LogProgramError "Could not create SMTPMailer Object in modReminders.EmailReminder." & Erl
120       Else
130         Call gSMTPMailer.Send("", "", Address, Subject, message, Filenames)
140       End If

150     Else
160       Set mapi = CreateObject("SENTRYMAIL.MAPITransport")
170       If mapi Is Nothing Then
180         LogProgramError "Could not create SENTRYMAIL Object in  modReminders.EmailReminder." & Erl
190       Else
200         Call mapi.SendWithAttachments("", "", Address, Subject, message, Filenames)
210       End If

220     End If


        '// Username, Password, Address, Subject,Body, AttachmentsList ' Attachemnet list is a semicolon ";" delimited list of file attachments

230     Set mapi = Nothing



End Function


Sub SendReminders()


' The reminders here are the ready reminders that need to be sent
' get distribution points
' build email que and voicemail que
  Dim j               As Long
  Dim Reminder        As cReminder
  Dim Subscriber      As cReminderSubscriber
  Dim DeliveryPoint   As cDeliveryPoint
  Dim message         As String

  

  Dim Voicemail       As cVoiceMail

  If Not MASTER Then Exit Sub

  Do While gRemindersToSend.Count
    Debug.Print "***********************"

    Set Reminder = gRemindersToSend(1)
    gRemindersToSend.Remove 1


    Set DeliveryPoint = Reminder.GetDeliveryPoint()
    Call FetchSubScribers(Reminder)
    If (1) Then

      ' send one to owner
      message = Reminder.ReminderMessage

      If Reminder.DeliveryPointID >= 0 Then
        Debug.Print "Send to Reminder Owner: " & message & " Via " & TranslateDeliveryPointType(DeliveryPoint.AddressType) & " to  " & DeliveryPoint.Address
        If DeliveryPoint.AddressType = DELIVERY_POINT.EMAIL Then
          If Len(DeliveryPoint.Address) Then  ' except if empty
            EmailReminder DeliveryPoint.Address, Reminder.ReminderMessage
          End If
        ElseIf DeliveryPoint.AddressType = DELIVERY_POINT.Phone Then
          If Len(DeliveryPoint.Address) Then  ' except if empty
            Set Voicemail = New cVoiceMail
            Voicemail.AckDigit = 0
            Voicemail.Address = DeliveryPoint.Address
            Voicemail.message = message
            VoiceMailQue.Add Voicemail
          End If
          ' send via dialogic
        ElseIf DeliveryPoint.AddressType = DELIVERY_POINT.phone_ack Then
          If Len(DeliveryPoint.Address) Then  ' except if empty
            Set Voicemail = New cVoiceMail
            Voicemail.AckDigit = Asc("0")
            Voicemail.Address = DeliveryPoint.Address
            Voicemail.message = message
            VoiceMailQue.Add Voicemail
          End If
        Else
          ' no send
        End If
      End If

      ' and then one to each subscriber, if they have a valid distribution point
      For Each Subscriber In Reminder.subscribers
        If Subscriber.ResidentID <> 0 Or Subscriber.StaffID <> 0 Then
          Set DeliveryPoint = Subscriber.GetPublicDeliveryPoint()
          If Not DeliveryPoint Is Nothing Then
            Debug.Print "Send to Subscriber " & Subscriber.NameAll & ": " & message & " Via " & TranslateDeliveryPointType(DeliveryPoint.AddressType) & " to  " & DeliveryPoint.Address
            If DeliveryPoint.AddressType = DELIVERY_POINT.EMAIL Then
              If Len(DeliveryPoint.Address) Then
                EmailReminder DeliveryPoint.Address, message
              End If
            ElseIf DeliveryPoint.AddressType = DELIVERY_POINT.Phone Then
              ' send via dialogic
              Set Voicemail = New cVoiceMail
              Voicemail.AckDigit = 0
              Voicemail.Address = DeliveryPoint.Address
              Voicemail.message = message
              VoiceMailQue.Add Voicemail

            ElseIf DeliveryPoint.AddressType = DELIVERY_POINT.phone_ack Then
              ' send via dialogic
              Set Voicemail = New cVoiceMail
              Voicemail.AckDigit = Asc("0")
              Voicemail.Address = DeliveryPoint.Address
              Voicemail.message = message
              VoiceMailQue.Add Voicemail
            Else
              ' no send
            End If
          End If
        ElseIf Subscriber.GroupID <> 0 Then
          ' send to system group
          SendToGroup message, Subscriber.GroupID, "", "", PAGER_NORMAL, left$(message, 19), 0, 0
        ElseIf Subscriber.PagerID <> 0 Then
          ' send to system pager (single output)
          SendToPager message, Subscriber.PagerID, 0, "", "", PAGER_NORMAL, left$(message, 19), 0, 0
        End If
      Next
    End If


  Loop

'  text = Trim(cboMessages.text)
'  If cbopager.ListIndex > 0 Then
'    If MASTER Then
'      SendToPager text, GetComboItemData(cbopager), 0, "", "", PAGER_NORMAL, left$(text, 19), 0
'    Else
'      ClientSendToPager text, GetComboItemData(cbopager), 0, ""
'    End If
'  End If
'  If cboGroup.ListIndex > 0 Then
'    If MASTER Then
'      SendToGroup text, GetComboItemData(cboGroup), "", "", PAGER_NORMAL, left$(text, 19)
'    Else
'      ClientSendToGroup text, GetComboItemData(cboGroup), ""
'    End If
'  End If



End Sub



Sub RemindersUpdate()
  Dim CurrentTime As Date
  
  LeadTime = 0  ' # of minutes of lead time' usually not set here, but thru a configuration setting

  ' Note Hour returns 0 from midnght, 0 to 23 hrs
  CurrentTime = Format$(Now, "dd/mm/yyyy hh:nn")

  If NextReminderCheck = 0 Then  ' this is at startup only
    NextReminderCheck = DateAdd("n", 2, CurrentTime)
    Exit Sub
  End If

  If DateDiff("n", CurrentTime, NextReminderCheck, vbSunday) = 0 Then
    NextReminderCheck = DateAdd("n", 1, CurrentTime)
    'Debug.Print "Time to process"
    CheckForReadyReminders
  ElseIf NextReminderCheck < CurrentTime Then
    NextReminderCheck = CurrentTime
  Else
    SendReminders
    ProcessVoicemail
    CheckDialogics
  End If

End Sub

Function CheckForReadyReminders() As Long ' should check either like once a minute or at midnight, set up all those that need processed
  Dim Rs        As ADODB.Recordset
  Dim SQL       As String
  Dim Reminder  As cReminder
  Dim ScheduledDate As Date
  Dim TheTime   As String
  Dim hrs       As Long
  Dim mins      As Long
  Dim DOW       As Long
  Dim DOM       As Long

  Dim LeadDate  As Date
  Dim Daymark   As String
  
  ' Lead Date is now + lead time
  LeadDate = Format$(DateAdd("n", LeadTime, Now), "mm/dd/yyyy hh:nn")

  DOW = Weekday(LeadDate)
  DOM = Day(LeadDate)

  If gRemindersToSend Is Nothing Then
    Set gRemindersToSend = New Collection
  End If

  ' create our time
  ' if the reminder is between now and ' allow for lead time

  'Debug.Print "****************"

  SQL = "SELECT * FROM reminders WHERE Cancelled = 0  AND specificday = '" & Format$(LeadDate, "mm/dd/yyyy") & "' AND frequency = " & REMINDER_DATE  ' lets do specific one-time - specific day reminders
  Set Rs = ConnExecute(SQL)
  Do Until Rs.EOF

    If IsDate(Rs("specificday")) Then  ' sanity check

      ' build date of reminder from parts
      hrs = Rs("timeofday") \ 100
      mins = Rs("timeofday") - (hrs * 100)
      ScheduledDate = CDate(Rs("specificday") & " " & hrs & ":" & mins)

      '      Debug.Print "Scheduled Date " & ScheduledDate
      '      Debug.Print "LeadDate Date " & LeadDate

      If 0 = DateDiff("n", ScheduledDate, LeadDate, vbSunday) Then
        Debug.Print "A reminder candidate " & ScheduledDate & "at " & Now
        Set Reminder = New cReminder
        Reminder.Parse Rs
        Reminder.GetDeliveryPoint
        gRemindersToSend.Add Reminder

      End If



    End If
    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing


    ' required for week day selection
  Daymark = "_______"
  Mid(Daymark, DOW, 1) = DOW


  SQL = "SELECT * FROM reminders WHERE Cancelled = 0 AND frequency = " & REMINDER_DAILY & " AND daysactive LIKE '" & Daymark & "'"  'These are daily
  Set Rs = ConnExecute(SQL)

  Do Until Rs.EOF
    hrs = Rs("timeofday") \ 100
    mins = Rs("timeofday") - (hrs * 100)
    ScheduledDate = CDate(Format(LeadDate, "mm/dd/yyyy") & " " & hrs & ":" & mins)
    If 0 = DateDiff("n", ScheduledDate, LeadDate, vbSunday) Then
      Debug.Print "A reminder candidate " & ScheduledDate & "at " & Now
      Set Reminder = New cReminder

      Reminder.Parse Rs
     gRemindersToSend.Add Reminder
    End If

    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing

  
  SQL = "SELECT * FROM reminders WHERE Cancelled = 0 AND frequency = " & REMINDER_MONTHLY & " AND DOM = " & DOM  'These are monthly

  Set Rs = ConnExecute(SQL)

  Do Until Rs.EOF
    hrs = Rs("timeofday") \ 100
    mins = Rs("timeofday") - (hrs * 100)
    ScheduledDate = CDate(Format(LeadDate, "mm/dd/yyyy") & " " & hrs & ":" & mins)
    If 0 = DateDiff("n", ScheduledDate, LeadDate, vbSunday) Then
      Debug.Print "A reminder candidate " & ScheduledDate & "at " & Now
      Set Reminder = New cReminder
      Reminder.Parse Rs
     gRemindersToSend.Add Reminder
    End If

    Rs.MoveNext
  Loop
  Rs.Close
  Set Rs = Nothing

End Function


Function DeleteReminder(reminderid As Long) As Long

  Dim Rs      As ADODB.Recordset
  Dim SQL     As String

10          On Error GoTo DeleteReminder_Error

20          SQL = "DELETE FROM ReminderSubscribers WHERE ReminderID = " & reminderid
30          ConnExecute SQL

40          SQL = "DELETE FROM Reminders WHERE ReminderID = " & reminderid
50          ConnExecute SQL


DeleteReminder_Resume:
60          On Error GoTo 0
70          Exit Function

DeleteReminder_Error:

80          LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modReminders.DeleteReminder." & Erl
90          Resume DeleteReminder_Resume


End Function

Public Function FetchSystemScribers(Reminder As cReminder) As Long
  Dim Rs      As ADODB.Recordset
  Dim SQL     As String
  Dim Key     As String
  Dim i       As Long
  Dim Subscriber As cReminderSubscriber




End Function

Public Function FetchSubScribers(Reminder As cReminder) As Long
        Dim Rs      As ADODB.Recordset
        Dim SQL     As String
        Dim Key     As String
        Dim i       As Long
        Dim Subscriber As cReminderSubscriber

10      On Error GoTo FetchSubScribers_Error

20      Set Reminder.subscribers = New Collection

30      SQL = "SELECT Namelast, NameFirst, Residents.ResidentID, DeliveryPoints,  0 as StaffID FROM Residents "
40      SQL = SQL & " INNER JOIN  ReminderSubscribers ON Residents.ResidentID = ReminderSubscribers.ResidentID "
50      SQL = SQL & " WHERE deleted = 0  AND  ReminderSubscribers.reminderid =  " & Reminder.reminderid
60      SQL = SQL & " UNION ALL "
70      SQL = SQL & " SELECT  Namelast, NameFirst, 0 as ResidentID, DeliveryPoints, Staff.StaffID FROM Staff "
80      SQL = SQL & " INNER JOIN  ReminderSubscribers ON Staff.staffID = ReminderSubscribers.subscriberID "
90      SQL = SQL & " WHERE deleted = 0 AND ReminderSubscribers.reminderid =  " & Reminder.reminderid
100     SQL = SQL & " ORDER BY NameLast, NameFirst"
110     Set Rs = ConnExecute(SQL)
120     Do Until Rs.EOF
130       i = i + 1  'DoEvents
140       Set Subscriber = New cReminderSubscriber
150       Subscriber.NameLast = Rs("Namelast") & ""
160       Subscriber.NameFirst = Rs("Namefirst") & ""
170       Subscriber.DeliveryPoints = Rs("DeliveryPoints") & ""

180       Subscriber.ResidentID = Rs("ResidentID")

190       Subscriber.StaffID = Rs("staffID")

          'Subscriber.Key = Subscriber.ResidentID & "-" & Subscriber.StaffID
200       Reminder.subscribers.Add Subscriber
210       Rs.MoveNext
220     Loop
230     Rs.Close





' groups

440     SQL = "SELECT  pagergroups.description, ReminderSubscribers.groupid  FROM pagergroups "
450     SQL = SQL & " INNER JOIN  ReminderSubscribers ON pagergroups.groupID = ReminderSubscribers.groupID "
460     SQL = SQL & " WHERE ReminderSubscribers.groupid <> 0 AND  ReminderSubscribers.ReminderID =  " & Reminder.reminderid
470     SQL = SQL & " ORDER BY description"
480     Set Rs = ConnExecute(SQL)
490     Do Until Rs.EOF
500       i = i + 1  'DoEvents
510       Set Subscriber = New cReminderSubscriber
520       Subscriber.NameLast = ""
530       Subscriber.NameFirst = ""
540       Subscriber.DeliveryPoints = ""
550       Subscriber.PagerName = Rs("description") & ""
560       Subscriber.ResidentID = 0

570       Subscriber.StaffID = 0

580       Subscriber.GroupID = Rs("groupid")
590       Subscriber.PagerID = 0

          'Subscriber.Key = Subscriber.ResidentID & "-" & Subscriber.StaffID
600       Reminder.subscribers.Add Subscriber
610       Rs.MoveNext
620     Loop
630     Rs.Close

' individual pagers

240     SQL = "SELECT pagers.description, ReminderSubscribers.pagerid  FROM pagers "
250     SQL = SQL & " INNER JOIN  ReminderSubscribers ON pagers.pagerID = ReminderSubscribers.pagerID "
260     SQL = SQL & " WHERE  ReminderSubscribers.pagerid <> 0 AND  ReminderSubscribers.reminderid =  " & Reminder.reminderid
270     SQL = SQL & " ORDER BY description"

280     Set Rs = ConnExecute(SQL)
290     Do Until Rs.EOF
300       i = i + 1  'DoEvents
310       Set Subscriber = New cReminderSubscriber
320       Subscriber.NameLast = ""
330       Subscriber.NameFirst = ""
340       Subscriber.DeliveryPoints = ""
350       Subscriber.PagerName = Rs("description") & ""
360       Subscriber.ResidentID = 0

370       Subscriber.StaffID = 0

380       Subscriber.GroupID = 0
390       Subscriber.PagerID = Rs("pagerid")

          'Subscriber.Key = Subscriber.ResidentID & "-" & Subscriber.StaffID
400       Reminder.subscribers.Add Subscriber
410       Rs.MoveNext
420     Loop
430     Rs.Close





        'Stop
        ' need to get pager info

FetchSubScribers_Resume:
640     Set Rs = Nothing
650     FetchSubScribers = i
660     On Error GoTo 0
670     Exit Function

FetchSubScribers_Error:

680     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modReminders.FetchSubScribers." & Erl
690     Resume FetchSubScribers_Resume

End Function
Public Function SaveReminder(Reminder As cReminder) As Long

  Dim Rs        As ADODB.Recordset
  Dim SQL       As String

  On Error GoTo SaveReminder_Error

  SQL = "SELECT * FROM Reminders WHERE ReminderID = " & Reminder.reminderid

  Set Rs = New ADODB.Recordset
  Rs.Open SQL, conn, gCursorType, gLockType

  If Rs.EOF Then  ' need to add record
    Rs.addnew
  End If
  Reminder.UpdateData Rs
  Rs.Update
  Reminder.reminderid = Rs("reminderid")
  Rs.Close

  SaveReminder = Reminder.reminderid

SaveReminder_Resume:
  Set Rs = Nothing
  On Error GoTo 0
  Exit Function

SaveReminder_Error:

  LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modReminders.SaveReminder." & Erl
  Resume SaveReminder_Resume

End Function
Public Function SaveSubscribers(subscribers As Collection, ByVal reminderid As Long) As Long
  Dim SQL         As String
  Dim ValueList   As String
  Dim Subscriber  As cReminderSubscriber


  If reminderid <> 0 Then

    SQL = "delete from remindersubscribers where ReminderID = " & reminderid
    ConnExecute SQL
    For Each Subscriber In subscribers
      
      ValueList = reminderid & "," & Subscriber.StaffID & "," & Subscriber.ResidentID & "," & Subscriber.PagerID & "," & Subscriber.GroupID
      
      SQL = "INSERT INTO remindersubscribers ( ReminderID, SubscriberID, ResidentID, PagerID, GroupID) Values (" & ValueList & ")"
      Debug.Print SQL
      ConnExecute SQL
    Next




  End If



End Function


Public Function GetReminder(ByVal reminderid As Long) As cReminder
  Dim Reminder As cReminder
  Dim Rs        As ADODB.Recordset
  Dim SQL       As String

10           On Error GoTo GetReminer_Error

20          Set Reminder = New cReminder
30          If reminderid <> 0 Then


40            SQL = "SELECT * FROM Reminders WHERE ReminderID = " & reminderid
50            Set Rs = ConnExecute(SQL)
60            If Not Rs.EOF Then
70              Reminder.Parse Rs
80            End If
90            Rs.Close





110         End If
120         Set GetReminder = Reminder

GetReminer_Resume:
125          Set Rs = Nothing
130          On Error GoTo 0
140          Exit Function

GetReminer_Error:

150         LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modReminders.GetReminder." & Erl
160         Resume GetReminer_Resume


End Function
