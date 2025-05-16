VERSION 5.00
Begin VB.UserControl AsyncReader 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ScaleHeight     =   1155
   ScaleWidth      =   1275
End
Attribute VB_Name = "AsyncReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Constant  Value Description
'vbAsyncStatusCodeError  0 An error occurred during the asynchronous download.
'vbAsyncStatusCodeFindingResource  1 AsyncRead is finding the resource specified in AsyncProperty.Status that holds the storage being downloaded.
'vbAsyncStatusCodeConnecting 2 AsyncRead is connecting to the resource specified in AsyncProperty.Status that holds the storage being downloaded.
'vbAsyncStatusCodeRedirecting  3 AsyncRead has been redirected to a different location specified in AsyncRead.Property.Status.
'vbAsyncStatusCodeBeginDownloadData  4 AsyncRead has begun receiving data for the storage specified in AsyncProperty.Status.
'vbAsyncStatusCodeDownloadingData  5 AsyncRead has received more data for the storage specified in AsyncProperty.Status.
'vbAsyncStatusCodeEndDownloadData  6 AsyncRead has finished receiving data for the storage specified in AsyncProperty.Status.
'vbAsyncStatusCodeUsingCashedCopy  10  AsyncRead is retrieving the requested storage from a cached copy. AsyncProperty.Status is empty.
'vbAsyncStatusCodeSendingRequest 11  AsyncRead is requesting the storage specified in AsyncProperty.Status.
'vbAsybcStatusCodeMIMETypeAvailable  13  The MIME type of the requested storage is specified in AsyncProperty.Status.
'vbAsyncStatusCodeCacheFileNameAvailable 14  The filename of the local file cache for requested storage is specified in AsyncProperty.Status.
'vbAsyncStatusCodeBeginSyncOperation 15  The AsyncRead will operate synchronously.
'vbAsyncstatusCodeEndSyncOperation 16  The AsyncRead has completed synchronous operation.
        
Public Event Done()

Public Response As String
Public Status As Long
Public ParentObject As cPushProcessor

Private mUrl As String
Public LogData As Boolean

Public Sub GetData(ByVal url As String, ByVal Querystring As String)

  mUrl = url & Querystring

  UserControl.AsyncRead mUrl, vbAsyncTypeByteArray, "Response"
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
  Dim BytesRead As Long
  
  Status = AsyncProp.StatusCode
  BytesRead = AsyncProp.BytesRead
  
  If Status = 0 Then
    LogPushOperation "Error " & BytesRead
  ElseIf Status = 6 Then
    LogPushOperation "Success " & BytesRead
  Else
    LogPushOperation "Other Result " & Status
  End If
    
  
  
  'Status = 0 if Error
  'Status = 6 if Done
  'Status = 3 is redirect
  

  
    
'  If AsyncProp.PropertyName = "Response" Then
'    If BytesRead > 0 Then
'      Response = StrConv(AsyncProp.value, vbUnicode)
'    End If
'  End If
  
  'UserControl.CancelAsyncRead "Response"
  
'  If Not ParentObject Is Nothing Then
'    ParentObject.ProcessDone Status, BytesRead, Response
'  Else
'    Debug.Assert 0
'  End If
  'RaiseEvent Done
  
End Sub
Private Sub UserControl_Initialize()
  ' nothing doin'
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
  Debug.Print "AsyncReadProgress AsyncProp.StatusCode " & AsyncProp.StatusCode
  ' AsyncProp.StatusCode = 15/16 if Synchronous
End Sub


Sub LogPushOperation(ByVal s As String)
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
