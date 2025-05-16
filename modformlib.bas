Attribute VB_Name = "modFormLib"
Option Explicit

Global hostedform As Form
Global HostedForms As New Collection
Global Const ENABLER_WIDTH = 15000
Global Const ENABLER_WIDTH_WIDE = 18000
Global Const ENABLER_HEIGHT = 3500


Function ShowMobileSettings()
  HostForm frmMobileSettings
  frmMobileSettings.Fill

End Function


Function ShowMobile()

  HostForm frmMobile
  frmMobile.Fill


End Function



Function HostForm(f As Form) As Long
  'ResetRemoteRefreshCounter
  If HostedForms.Count >= 10 Then ' max 10 levels deep
    HostedForms(1).UnHost
    HostedForms.Remove 1
  End If
  
  If HostedForms.Count > 0 Then
    HostedForms(HostedForms.Count).UnHost ' unhost last form, but don't remove it
  End If
  
  HostedForms.Add f
  
  Set hostedform = HostedForms(HostedForms.Count)
  
  frmMain.fraHost.Width = ENABLER_WIDTH
  frmMain.fraHost.Height = ENABLER_HEIGHT
  hostedform.Host frmMain.fraHost.hwnd
  frmMain.Assur = False

End Function
Function PreviousFormWithValue(ByVal ReturnValue As String) As Long
  'ResetRemoteRefreshCounter
  If HostedForms.Count > 0 Then
    HostedForms(HostedForms.Count).UnHost
    HostedForms.Remove HostedForms.Count
  End If
  If HostedForms.Count > 0 Then
    Set hostedform = HostedForms(HostedForms.Count)
    hostedform.Host frmMain.fraHost.hwnd
    hostedform.ReturnValue = ReturnValue
    'HostedForm.Fill
  End If
End Function


Function PreviousForm() As Long
  'ResetRemoteRefreshCounter
  Dim HostedCount As Integer
  
  HostedCount = HostedForms.Count
  If HostedCount > 0 Then
    HostedForms(HostedCount).UnHost
    HostedForms.Remove HostedCount
  End If
  HostedCount = HostedForms.Count
  If HostedCount > 0 Then
    Set hostedform = HostedForms(HostedCount)
    hostedform.Host frmMain.fraHost.hwnd
    hostedform.Fill
  End If
End Function
Function ClearHostedForms() As Long
  'ResetRemoteRefreshCounter
  If HostedForms.Count > 0 Then
    HostedForms(HostedForms.Count).UnHost
  End If
  Set HostedForms = New Collection
End Function
Function RemoveHostedForms() As Long
  Dim j As Long
  For j = HostedForms.Count To 1 Step -1
    HostedForms(j).UnHost
  Next
  Set HostedForms = New Collection

End Function


Sub ChangeHostAdapter()
  HostForm frmAdapters
  frmAdapters.Fill

End Sub

Sub CreatePartitonsFromRooms()
  HostForm frmRooms2Partitions
  frmRooms2Partitions.Fill
End Sub

Sub ExceptionReportView(ByVal ReportType As Long, ByVal Criteria As String, ByVal StartDate As String, ByVal EndDate As String)
  HostForm frmExceptionView
  Call frmExceptionView.AdvancedReport(ReportType, Criteria, StartDate, EndDate)

End Sub


Sub PrintAutoReportList()
  HostForm frmAutoReportPrintList
  frmAutoReportPrintList.Fill

End Sub
Sub ManageAvailablePartitions()
  HostForm frmPartitions
  frmPartitions.Fill
End Sub


Sub ShowPictures(ByVal ResidentID As Long)
  HostForm frmGetImage
  frmGetImage.ResidentID = ResidentID
  frmGetImage.Fill
End Sub
Sub ShowRooms(ByVal ResidentID As Long, ByVal TransmitterID As Long, ByVal RoomID As Long, ByVal Caller As String)
        
10      On Error Resume Next
        
20      HostForm frmRooms
30      frmRooms.RoomID = RoomID
40      frmRooms.ResidentID = ResidentID
50      frmRooms.TransmitterID = TransmitterID
60      frmRooms.Caller = Caller ' EX: "TX" = transmitter
70      frmRooms.Fill

80      If Err.Number Then
90        LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modFormLib.ShowRooms." & Erl
100     End If

End Sub


Sub ShowAnnouncementForm()
  HostForm frmAnnounce
  frmAnnounce.Fill
  frmAnnounce.Focus
End Sub
Sub ShowReminderConfig()
  HostForm frmReminderConfig
  frmReminderConfig.Fill
End Sub

Sub ShowStaff(ByVal StaffID As Long, ByVal reminderid As Long, ByVal IsPublic As Integer, ByVal Caller As String)
  HostForm frmStaff
  
  frmStaff.ResidentID = StaffID
  frmStaff.Caller = Caller
  frmStaff.IsPublic = IsPublic
  frmStaff.Fill

End Sub

Sub ShowResidents(ByVal ResidentID As Long, ByVal TransmitterID As Long, ByVal Caller As String)
  
  HostForm frmResidents
  frmResidents.ResidentID = ResidentID
  frmResidents.TransmitterID = TransmitterID
  frmResidents.Caller = Caller
  frmResidents.Fill

End Sub
Sub EditOutput(ByVal ID As Long)
  HostForm frmOutputEdit
  frmOutputEdit.PagerID = ID
  frmOutputEdit.Fill
  

End Sub
Sub EditSerialDevice(Device As cESDevice)
  HostForm frmSerialDevice
  Set frmSerialDevice.Device = Device
  frmSerialDevice.Fill


End Sub
Sub EditTemperatureDevice(Device As cESDevice)
  HostForm frmTemperatureDevice
  Set frmTemperatureDevice.Device = Device
  frmTemperatureDevice.Fill


End Sub

Sub ShowGroups(ByVal ResidentID As Long, ByVal RoomID As Long)
  HostForm frmOutputGroups
  frmOutputGroups.ResidentID = ResidentID
  frmOutputGroups.RoomID = RoomID
  frmOutputGroups.Fill

End Sub
Sub EditGroup(ByVal ID As Long)
  HostForm frmOutputGroup
  frmOutputGroup.GroupID = ID
  frmOutputGroup.Fill

End Sub
Sub ListAutoReports()
  HostForm frmAutoReports
  frmAutoReports.Fill
End Sub

Sub EditStaff(ByVal ID As Long)
  HostForm frmStaffEdit
  frmStaffEdit.StaffID = ID
  frmStaffEdit.Fill

End Sub



Sub EditResident(ByVal ID As Long)
  HostForm frmResident
  frmResident.ResidentID = ID
  frmResident.Fill
End Sub
Sub ShowTransmitters(ByVal ResidentID As Long, ByVal RoomID As Long)
  HostForm frmTransmitters
  frmTransmitters.ResidentID = ResidentID
  frmTransmitters.RoomID = RoomID
  frmTransmitters.Fill
  
End Sub
Sub FindPCAs(ByVal Serial As String, ByVal PagerID As Long)
  HostForm frmPCAs
  frmPCAs.PagerID = PagerID
  frmPCAs.Serial = Serial
  frmPCAs.Fill
  
End Sub
Public Sub EditExceptionReport(ByVal ID As Long)
  HostForm frmExceptionReport
  frmExceptionReport.ReportID = ID
  frmExceptionReport.Fill

End Sub


Sub EditAutoReport(ByVal ID As Long)
  HostForm frmAutoReportEdit
  frmAutoReportEdit.ReportID = ID
  frmAutoReportEdit.Fill

End Sub
Sub EditTransmitter(ID As Long)
  
  ' not sure why there's a ref here?
  Dim f As Form
  Set f = frmTransmitter
  'frmDevice
  HostForm f
  f.DeviceID = ID
  f.Fill
End Sub

Sub EditPublicEvent(ID As Long, OwnerID As Long)
  HostForm frmPublicEventEdit
  frmPublicEventEdit.IsPublic = 1
  frmPublicEventEdit.OwnerID = OwnerID
  frmPublicEventEdit.reminderid = ID
  frmPublicEventEdit.Fill
End Sub


Sub EditPrivateEvent(ID As Long, OwnerID As Long)
  HostForm frmPublicEventEdit
  frmPublicEventEdit.IsPublic = 0
  frmPublicEventEdit.OwnerID = OwnerID
  frmPublicEventEdit.reminderid = ID
  frmPublicEventEdit.Fill
End Sub



Sub EditRoom(ID As Long)
  HostForm frmRoomEdit
  frmRoomEdit.RoomID = ID
  frmRoomEdit.Fill

End Sub
Sub ShowReportMenu()
  HostForm frmReportChoices
End Sub

Sub AdvancedReport(ByVal ReportType As Integer, ByVal Criteria As String, ByVal StartDate As Date, ByVal EndDate As Date)
  
  HostForm frmReports
  Call frmReports.AdvancedReport(ReportType, Criteria, StartDate, EndDate)
End Sub


Sub BasicReport(ByVal ID As String)
  HostForm frmReports
  frmReports.ReportID = ID
  frmReports.Fill
End Sub

Sub ShowOutputs(ByVal ID As Long)
  HostForm frmOutputs
  frmOutputs.Fill

End Sub

Sub ShowOutputServers(ByVal ID As Long)
  HostForm frmOutputServers
  frmOutputServers.Fill

End Sub

Sub EditOutputServer(ByVal ID As Long)
  SpecialLog "Call HostForm frmOutputServer"
  HostForm frmOutputServer
  SpecialLog "Call frmOutputServer.ServerID = ID"
  frmOutputServer.ServerID = ID
  SpecialLog "Call frmOutputServer.Fill"
  frmOutputServer.Fill
  SpecialLog "Done frmOutputServer.Fill"
End Sub

Sub ShowConfigure1()
  HostForm frmConfigure
  'frmConfigure.Fill

End Sub
Sub DoRepeaters()
  HostForm frmRepeaters
  frmRepeaters.Fill
  
End Sub
Sub DoImports()
  HostForm frmImport
   
End Sub
Sub GetWaveFile(ByVal ID As Long)
  HostForm frmGetWav
  frmGetWav.ID = ID
  
End Sub
Sub ShowTransmitterTypes()
  HostForm frmTransmitterTypes
  frmTransmitterTypes.ClearForm
End Sub

Function tMsgBox(prompt As String, Buttons As VbMsgBoxStyle, Title As String, Timeout As Integer, Context As Form)
  Dim f As frmTimedMessageBox
  f.prompt = prompt
  f.Buttons = Buttons
  f.Title = Title
  f.Timeout = Timeout ' in seconds
  If Context Is Nothing Then
    f.Show vbModal
  Else
    f.Show vbModal, Context
  End If
  
  
End Function
Sub ShowBackupSettings()
  HostForm frmBackup
  frmBackup.Fill

End Sub
Sub ShowAssurSend()
  HostForm frmAssurSend
  frmAssurSend.Fill
End Sub

Sub ShowScreenMask()
  HostForm frmOutputMask
  frmOutputMask.Fill

End Sub
Sub ShowUsers()
  HostForm frmUsers
  frmUsers.Fill
End Sub

Sub ShowUpgrade()
  HostForm frmUpgrade
  
End Sub

Sub ShowFactorySettings()
  HostForm frmFactory
  frmFactory.Fill

End Sub

Sub ShowEmailSettings()
  HostForm frmSMTPSetup
  frmSMTPSetup.Fill
  
End Sub
Sub ShowOtherPrograms()
  HostForm frmExternalUtils
  frmExternalUtils.Fill

End Sub
Sub GetExternalApp(ByVal Source As String)
  HostForm frmExternalApp
  frmExternalApp.Source = Source
  frmExternalApp.Fill

End Sub
Sub EditSoftPoints()
  HostForm frmSoftPoints
  frmSoftPoints.Fill
End Sub


Sub EditUser(ByVal ID As Long)
  HostForm frmUser
  frmUser.UserID = ID
  frmUser.Fill
End Sub

Sub ShowWaypoints()
  
  HostForm frmWaypoints
  frmWaypoints.Fill

End Sub
Sub EditWaypoint(ByVal ID As Long)
  HostForm frmWaypoint
  frmWaypoint.ID = ID
  frmWaypoint.Fill
End Sub
Sub EditPCA(ByVal Serial As String, Optional ByVal Caller As String)
  HostForm frmPCAConfiguration
  frmPCAConfiguration.Serial = Serial
'  frmRooms.Caller = Caller
  frmPCAConfiguration.Fill

End Sub


Sub SetMainCaption()
  On Error Resume Next
  frmMain.UpdateScreenElements
 
End Sub
Sub GetFolder(ByVal Folder As String, ByVal Caller As String)
  HostForm frmGetFolder
  frmGetFolder.Caller = Caller
  frmGetFolder.Folder = Folder
  frmGetFolder.FolderRemote = False
  frmGetFolder.Fill

End Sub

Sub GetFolderremote(ByVal Folder As String, ByVal Caller As String)
  HostForm frmGetFolder
  frmGetFolder.Caller = Caller
  frmGetFolder.FolderRemote = True
  frmGetFolder.Folder = Folder
  frmGetFolder.Fill

End Sub
Sub ShowReminders()
  HostForm frmPublicEvents
  frmPublicEvents.Fill
  
End Sub
Sub ShowDukane()
  HostForm frmDukane
  frmDukane.Fill

End Sub
Sub ShowPush()
  HostForm frmPush
  frmPush.Fill
End Sub



Sub ShowDebugScreen()
  HostForm frmDebug
  frmDebug.Fill
End Sub
