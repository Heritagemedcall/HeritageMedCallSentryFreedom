VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Done()
Public Event ReadyStateChange(ByVal value As Long)

Public logging As Boolean

Private XMLHttpRequest   As MSXML2.XMLHTTP60
Private WithEvents OnReadyStateWrapper As ReadyStateHandler
Attribute OnReadyStateWrapper.VB_VarHelpID = -1
Public Status            As String
Public XML               As String
Public Response          As String



Public Function Send(ByVal Method As String, ByVal url As String, ByVal PostData As String, Optional ByVal Username As String = "", Optional ByVal Password As String = "") As Long
  Dim Auth As String

  XMLHttpRequest.OnReadyStateChange = OnReadyStateWrapper ' Assign the wrapper class object to onreadystatechange NOT Sure Why No SET for object.
  XMLHttpRequest.Open Method, url, True
  ' we need fresh meat:
'     auth = EncodeBase64(Username & ":" & Password)  ' 'Auth = "QWRtaW46qwrTAw4=" = "Admin:Admin"
'     Call XMLHttpRequest.setRequestHeader("Authorization", "Basic " & auth)
  
  Call XMLHttpRequest.setRequestHeader("If-Modified-Since", "Sat, 29 Oct 1994 19:43:31 GMT")
  
  If Len(PostData) Then
    XMLHttpRequest.Send PostData
  Else
    XMLHttpRequest.Send
  End If
End Function

Private Sub Class_Initialize()
  Set XMLHttpRequest = New MSXML2.XMLHTTP60
  Set OnReadyStateWrapper = New ReadyStateHandler
  

End Sub

Private Sub Class_Terminate()
  Set XMLHttpRequest = Nothing
  Set OnReadyStateWrapper = Nothing
End Sub

Private Sub OnReadyStateWrapper_OnReadyStateChangeEvent()

  RaiseEvent ReadyStateChange(XMLHttpRequest.READYSTATE)

  If XMLHttpRequest.READYSTATE = 4 Then
    Status = XMLHttpRequest.Status
    Response = XMLHttpRequest.ResponseText
    RaiseEvent Done
  End If

End Sub


