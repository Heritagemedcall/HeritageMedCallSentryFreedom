VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRemoteAutoEnroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This is for asyncronous Remote Auto Enroll

Public RemoteEnrollEnabled As Boolean
Public RemoteEnrollReady   As Boolean

Private XML As String
Private mTimeout As Long

Private Const TIMEOUT_WINDOW = 30000 ' 30 seconds
Private mStarted As Boolean

Public Function Cancel() As String
  ' build xml cancel string
  
  
  
  Dim str As String
  
  str = "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & taggit("Cancel", "OK") & vbCrLf
  str = str & "</HMC>"
  
  Cancel = str
  RemoteEnrollReady = False
  mStarted = False
End Function
Public Property Get Ready() As Boolean
  Ready = RemoteEnrollReady
End Property


Public Property Get Timeout() As Boolean
  Timeout = (Win32.timeGetTime > mTimeout) And Started
End Property

Public Property Get TimeOutError() As String
  'TimeOutError = ReturnFailure("AutoEnroll")
  'build custom timeout xml
  Dim XML As String
    XML = "<?xml version=""1.0""?>" & vbCrLf
    XML = XML & "<HMC revision=""" & App.Revision & """>" & vbCrLf
    XML = XML & taggit("TimeOut", "CheckRemoteAutoEnroll") & vbCrLf
    XML = XML & "</HMC>"
    TimeOutError = XML
End Property

Public Function CheckTimeout() As Boolean
    
  If Win32.timeGetTime > mTimeout Then
    CheckTimeout = True
  End If
End Function

'    CheckTimeout = True
'    RemoteEnrollReady = True
'    xml = ReturnFailure("AutoEnroll")
'  End If
'
'End Function

Public Function start(ByVal Client As String) As Boolean
  
  RemoteEnrollReady = False
  RemoteEnrollEnabled = True
  XML = ""
  mTimeout = Win32.timeGetTime + TIMEOUT_WINDOW
  start = True
  Started = True

End Function

Public Function AutoEnroll(p As cESPacket) As Boolean
  XML = ReturnAutoEnroll("AutoEnroll", p.Serial, p.CLSPTI)
  RemoteEnrollEnabled = False
  RemoteEnrollReady = True
End Function
Function ReturnAutoEnroll(ByVal Action As String, ByVal Serial As String, ByVal CLSPTI As Long) As String
  
  Dim str As String
  
  str = "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & taggit("Success", Action) & vbCrLf
  str = str & taggit("Serial", Serial) & vbCrLf
  str = str & taggit("CLSPTI", CStr(CLSPTI)) & vbCrLf
  str = str & "</HMC>"
  
  ReturnAutoEnroll = str

End Function
Function ReturnAutoEnrollWaiting() As String
  
  Dim str As String '
  
  
  str = "<?xml version=""1.0""?>" & vbCrLf
  str = str & "<HMC revision=""" & App.Revision & """>" & vbCrLf
  str = str & taggit("Waiting", "OK") & vbCrLf
  str = str & "</HMC>"
  
  ReturnAutoEnrollWaiting = str

End Function


Public Function GetAutoEnrollResult() As String
  GetAutoEnrollResult = XML
End Function

Public Property Get Started() As Boolean
  Started = mStarted
End Property

Public Property Let Started(ByVal Started As Boolean)
  mStarted = Started
End Property
