VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPushProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LastError         As String
Public Que               As Collection
Private evt              As cPushEvent
Private Retrys           As Long
Private WithEvents HttpAsync As cHttpAsync
Attribute HttpAsync.VB_VarHelpID = -1
Public Busy              As Boolean

Private LogData          As Boolean

'Public AsyncControl As Control


'Public Sub SetAsyncControl(Async As Control)
'  Set AsyncControl = Async
'End Sub


Private Sub HttpAsync_Done()

  ' use with HTTP Object

  Debug.Print " >>>>>>>>>> HttpAsync_Done " & HttpAsync.Response

  LastError = HttpAsync.Status
  
  Select Case Val(LastError)
    Case 200 To 299
      'LogOutGoing "HTTP OK " & LastError
    Case 300 To 399
      LogOutGoing "HTTP Info " & LastError
      
    Case Else
      LogOutGoing "HTTP Error " & LastError
  End Select

  'If LastError = 0 Then
  Set evt = Nothing
  'End If

  If evt Is Nothing Then
    If Que.Count Then
      DeQueAndSend             ' try and process que
    Else
      Busy = False
    End If
  Else
    Busy = False               ' que is empty... done!
  End If

End Sub





'Public Sub ProcessDone(ByVal Status As Long, ByVal BytesRead As Long, ByVal Response As String)
'  LastError = Status
'
'  Set evt = Nothing
'
'
'  If evt Is Nothing Then
'    If Que.Count Then
'      DeQueAndSend               ' try and process que
'    Else
'      Busy = False
'    End If
'  Else
'    Busy = False               ' que is empty... done!
'  End If
'
'
'End Sub


Public Sub Send()

  'Debug.Print "cPushProcessor.Send QueCount = " & Que.Count & "  Busy = "; Busy

  If Not Busy Then
    If Que.Count Then
      DeQueAndSend
    End If
  End If
  On Error Resume Next
  LogData = CBool(ReadSetting("Push", "Log", "False"))

End Sub

Function DeQueAndSend()
  On Error Resume Next
  Dim Address            As String
  Dim QS                 As String

  Busy = True

  'Debug.Print "cPushProcessor.DeQueAndSend QueCount = " & Que.Count

  Address = GetPushAddress

  If Len(Address) Then         ' hopefully good address

    If InStr(1, Address, "https://", vbTextCompare) <> 1 Then
      If InStr(1, Address, "http://", vbTextCompare) <> 1 Then
        Address = "http://" & Address
      End If
    End If

    
    Set evt = Que.Item(1)
    Que.Remove 1
    evt.Retries = 0
    
    
    

    If Not (evt Is Nothing) Then
      If evt.Retries > -1 Then
        evt.Retries = evt.Retries - 1
        
        QS = evt.ToQueryString()

        If (LogData) Then
          
          LogOutGoing Right$("  " & evt.EventType, 2) & " DeQueAndSend " & evt.Serial & "  " & Address & QS
        End If

        HttpAsync.Send "GET", Address & QS, "", "", ""
        '        AsyncControl.GetData Address, QS
        Exit Function
      End If
    End If
  Else
    Que.Remove 1
  End If

End Function


'Private Sub Async_Done()
'
'  ''not used.
'
'  LastError = AsyncControl.Status
'
'  If LastError = 6 Or LastError = 0 Then
'    Set evt = Nothing
'  End If
'
'  If evt Is Nothing Then
'    If Que.Count Then
'      DeQueAndSend             ' try and process que
'    Else
'      Busy = False
'    End If
'  Else
'    Busy = False               ' que is empty... done!
'  End If
'
'End Sub

'Private Sub Connection_Done()
'
'  ' use with HTTP Object
'
'  Debug.Print " >>>>>>>>>> Connection_Done"
'
'  LastError = Connection.LastStatus & ""
'  'If LastError = 0 Then
'  Set evt = Nothing
'  'End If
'
'  If evt Is Nothing Then
'    If Que.Count Then
'      DeQueAndSend             ' try and process que
'    Else
'      Busy = False
'    End If
'  Else
'    Busy = False               ' que is empty... done!
'  End If
'End Sub


Public Function AddEvent(e As cPushEvent)

  Debug.Print e.EventType & " AddEvent " & e.Serial
  
  If Que.Count > 2000 Then
    Que.Remove 1
  End If
  Que.Add e
End Function

Public Function AddByID(ByVal ID As Long)
  Dim e                  As cPushEvent

  Set e = New cPushEvent
  e.CreateFromRowID ID
  If e.IsValid Then
    AddEvent e
  End If
  Set e = Nothing

End Function

Public Sub ClearQue()

  If Not Busy Then
    Busy = True
    Do While Que.Count
      Que.Remove 1
    Loop
    Busy = False
  End If

End Sub

Private Function GetRetries() As Long
  GetRetries = Val(ReadSetting("Push", "Retries", 0))
End Function


Private Function GetPushAddress() As String
  GetPushAddress = ReadSetting("Push", "URL", "")
End Function

Private Sub Class_Initialize()
  Set Que = New Collection
  Set HttpAsync = New cHttpAsync
End Sub

Private Sub Class_Terminate()
  Set HttpAsync = Nothing
  Set Que = Nothing
End Sub

Private Sub HttpAsync_ReadyStateChange(ByVal value As Long)

  If value < 1 Or value > 4 Then
    Debug.Assert 0
  End If

  Debug.Print "HttpAsync_ReadyStateChange " & value
End Sub

Sub LogOutGoing(ByVal s As String)
  Dim hfile              As Long
  Dim filename As String
  filename = App.Path & "\Push.Log"
  limitFileSize filename
  On Error Resume Next
  
  If (LogData) Then
    hfile = FreeFile
    Open filename For Append As hfile
    Print #hfile, Format$(Now, "hh:nn:ss") & " " & s
    Close hfile
  End If

End Sub
