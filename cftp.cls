VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BUFFERSIZE = 255

Private mInternet           As Long
Private mConnection     As Long
'Private hfile           As Long
Private TransferType    As Long
Private Semantic        As Long

Private ErrorMessage    As String
Private mDirCol         As New cDirList

Public Event FileTransferProgress(CurrentBytes As Long, TotalBytes As Long)

Property Get Directory() As cDirList
  Set Directory = mDirCol
End Property

Private Sub ErrorOut(ByVal ErrCode As Long, ByRef Func As String)
  Dim rc As Long
  Dim Temp As Long
  Dim Buffer As String
  
  Buffer = String(255, vbNullChar)
  rc = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, GetModuleHandle("wininet.dll"), ErrCode, 0, Buffer, Len(Buffer), 0)
  ErrorMessage = Func & " error code: " & ErrCode & " Message: " & Buffer
  If (ErrCode = 12003) Then
    ' Extended error information was returned
    Buffer = String(2048, vbNullChar)
    rc = InternetGetLastResponseInfo(Temp, Buffer, Len(Buffer))
    
    ErrorMessage = Buffer
  End If
End Sub

Property Get GetLastErrorMessage() As String
  GetLastErrorMessage = ErrorMessage
End Property

Public Sub CloseConnection()
  If mConnection <> 0 Then
    InternetCloseHandle mConnection
  End If
  mConnection = 0
End Sub

Public Function SimpleFTPPutFile(LocalName As String, RemoteName As String) As Boolean
  ' uses FTPPutFile to write entire file at once
  
  If (FTPPutFile(mConnection, LocalName, RemoteName, TransferType, 0) = False) Then
    ErrorOut Err.LastDllError, "SimpleFtpPutFile"
    SimpleFTPPutFile = False
    Exit Function
  Else
    SimpleFTPPutFile = True
  End If
End Function

Public Function RenameFTPFile(Existing As String, NewName As String) As Boolean
  If (FtpRenameFile(mConnection, Existing, NewName) = False) Then
    ErrorOut Err.LastDllError, "RenameFTPFile"
    RenameFTPFile = False
    Exit Function
  Else
    RenameFTPFile = True
  End If
End Function

Public Function CreateFTPDirectory(Directory As String) As Boolean
  If (FtpCreateDirectory(mConnection, Directory) = False) Then
    ErrorOut Err.LastDllError, "CreateFTPDirectory"
    CreateFTPDirectory = False
    Exit Function
  Else
    CreateFTPDirectory = True
  End If
End Function

Public Function RemoveFTPDirectory(Directory As String) As Boolean
  If (FtpRemoveDirectory(mConnection, Directory) = False) Then
    ErrorOut Err.LastDllError, "RemoveFTPDirectory"
    RemoveFTPDirectory = False
    Exit Function
  Else
    RemoveFTPDirectory = True
  End If
End Function

Public Function DeleteFTPFile(Remote As String) As Boolean
  If (FtpDeleteFile(mConnection, Remote) = False) Then
    ErrorOut Err.LastDllError, "DeleteFTPFile"
    DeleteFTPFile = False
    Exit Function
  Else
    DeleteFTPFile = True
  End If
End Function

Public Function OpenConnection(Server As String, User As String, Password As String) As Boolean
  If mConnection <> 0 Then
    InternetCloseHandle mConnection
  End If
  mConnection = InternetConnect(mInternet, Server, INTERNET_INVALID_PORT_NUMBER, User, Password, INTERNET_SERVICE_FTP, Semantic, 0)
  If mConnection = 0 Then
    ErrorOut Err.LastDllError, "InternetConnect"
    OpenConnection = False
    Exit Function
  Else
    OpenConnection = True
  End If
End Function

Public Function FTPUploadFile(ByVal LocalFile As String, ByVal RemoteFile As String) As Boolean
  ' uses InternetWriteFile to write chunked transfer
  ' returns true on success
 
  Dim Data(BUFFERSIZE - 1) As Byte
  
  Dim Written       As Long
  Dim Size          As Long
  Dim Sum           As Long
  Dim Block         As Long
  Dim hfile        As Integer
  
  Dim hFTP         As Long
  
  hfile = FreeFile

  Sum = 0
  Block = 0
  LocalFile = Trim$(LocalFile)
  RemoteFile = Trim$(RemoteFile)

  If LocalFile <> "" And RemoteFile <> "" Then
    hFTP = FtpOpenFile(mConnection, RemoteFile, GENERIC_WRITE, TransferType, 0)
    If hFTP = 0 Then
      ErrorOut Err.LastDllError, "FtpOpenFile:PutFile"
      FTPUploadFile = False
      Exit Function
    End If

    Open LocalFile For Binary Access Read As #hfile
    Size = LOF(hfile)
    
    
    For Block = 1 To Size \ BUFFERSIZE
      Get #hfile, , Data
      If (InternetWriteFile(hFTP, Data(0), BUFFERSIZE, Written) = 0) Then
        FTPUploadFile = False
        ErrorOut Err.LastDllError, "InternetWriteFile"
        Exit Function
      End If
      DoEvents
      Sum = Sum + BUFFERSIZE

      RaiseEvent FileTransferProgress(Sum, Size)
    Next Block

    'check for leftovers
    If Size Mod BUFFERSIZE <> 0 Then
      Get #hfile, , Data
      If (InternetWriteFile(hFTP, Data(0), Size Mod BUFFERSIZE, Written) = 0) Then
        FTPUploadFile = False
        ErrorOut Err.LastDllError, "InternetWriteFile2"
        Exit Function
      End If
    End If

    Sum = Sum + (Size Mod BUFFERSIZE)
    Size = Sum
    RaiseEvent FileTransferProgress(Sum, Size)
    Close #hfile
    InternetCloseHandle (hFTP)
    FTPUploadFile = True
  End If
End Function

Public Function FTPDownloadFile(LocalFile As String, RemoteFile As String) As Boolean
  Dim Data(BUFFERSIZE - 1) As Byte  ' array of 100 elements 0 to 99
  
  Dim Written       As Long
  Dim Size          As Long
  Dim Sum           As Long
  Dim Block         As Long
  
  Dim hFileHandle   As Integer
  
  Dim hfile         As Long
  
  FTPDownloadFile = False

  Sum = 0
  Block = 0

  LocalFile = Trim(LocalFile)
  RemoteFile = Trim(RemoteFile)

  If LocalFile <> "" And RemoteFile <> "" Then
    Size = GetFTPFileSize(RemoteFile)
    If Size > 0 Then
      hfile = FtpOpenFile(mConnection, RemoteFile, GENERIC_READ, TransferType, 0)
      If hfile = 0 Then
        ErrorOut Err.LastDllError, "FtpOpenFile:GetFile"
        Exit Function
      End If

      hFileHandle = FreeFile
      Open LocalFile For Binary Access Write As #hFileHandle
      Seek #hFileHandle, 1
      Sum = 1
      For Block = 1 To Size \ BUFFERSIZE
        If (InternetReadFile(hfile, Data(0), BUFFERSIZE, Written) = 0) Then
          ErrorOut Err.LastDllError, "InternetReadFile"
          Close #hFileHandle
          Exit Function
        End If
        Put #hFileHandle, , Data
        DoEvents
        Sum = Sum + BUFFERSIZE
        RaiseEvent FileTransferProgress(Sum, Size)
      Next Block

      'Check for leftovers
      If Size Mod BUFFERSIZE <> 0 Then
        ReDim Data2((Size Mod BUFFERSIZE) - 1) As Byte
        If (InternetReadFile(hfile, Data2(0), Size Mod BUFFERSIZE, Written) = 0) Then
          ErrorOut Err.LastDllError, "InternetReadFile2"
          Close #hFileHandle
          Exit Function
        End If
      End If

      Put #hFileHandle, , Data2

      Sum = Sum + (Size Mod BUFFERSIZE)
      Size = Sum
      RaiseEvent FileTransferProgress(Sum, Size)
      Close #hFileHandle
      InternetCloseHandle (hfile)
      FTPDownloadFile = True
    End If
  End If
End Function

Public Function SimpleFTPGetFile(sLocal As String, sRemote As String) As Boolean
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
  If (FTPGetFile(mConnection, sRemote, sLocal, False, FILE_ATTRIBUTE_NORMAL, TransferType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    ErrorOut Err.LastDllError, "SimpleFtpGetFile"
    SimpleFTPGetFile = False
    Exit Function
  Else
    SimpleFTPGetFile = True
  End If
End Function

Public Function GetFTPDirectory() As String
  Dim Directory As String
  Directory = String(1024, vbNullChar)
  If (FtpGetCurrentDirectory(mConnection, Directory, 1024) = False) Then
    ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
    Exit Function
  Else
    GetFTPDirectory = left(Directory, InStr(1, Directory, String(1, vbNullChar), vbBinaryCompare) - 1)
  End If
End Function

Public Function SetFTPDirectory(ByVal DirName As String)
  If (FtpSetCurrentDirectory(mConnection, DirName) = False) Then
    ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
    SetFTPDirectory = False
    Exit Function
  Else
    SetFTPDirectory = True
  End If
End Function

Public Function GetFTPFileSize(ByVal filename As String) As Long
  Dim szDir As String
  Dim hFind As Long
  Dim nLastError As Long
  Dim pData As WIN32_FIND_DATA

  hFind = FtpFindFirstFile(mConnection, filename, pData, 0, 0)
  nLastError = Err.LastDllError
  If hFind = 0 Then
    If (nLastError = ERROR_NO_MORE_FILES) Then
      GetFTPFileSize = -1  ' File not found
    Else
      GetFTPFileSize = -2  ' Other error
      ErrorOut Err.LastDllError, "FtpFindFirstFile"
    End If
    Exit Function
  End If

  GetFTPFileSize = pData.nFileSizeLow
  InternetCloseHandle (hFind)
End Function

Public Function GetDirectoryListing(sFilter As String) As cDirList
  Dim szDir As String
  Dim hFind As Long
  Dim nLastError As Long
  Dim dError As Long
  Dim Ptr As Long
  Dim pData As WIN32_FIND_DATA
  Dim filename As String

  Set mDirCol = Nothing
  hFind = FtpFindFirstFile(mConnection, sFilter, pData, 0, 0)
  nLastError = Err.LastDllError
  If hFind = 0 Then
    If (nLastError <> ERROR_NO_MORE_FILES) Then
      ErrorOut Err.LastDllError, "FtpFindFirstFile"
    End If
    Exit Function
  End If

  dError = NO_ERROR
  Dim bRet As Boolean

  filename = left(pData.cFileName, InStr(1, pData.cFileName, String(1, vbNullChar), vbBinaryCompare) - 1)
  mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, filename
  Do
    pData.cFileName = String(MAX_PATH, 0)
    bRet = InternetFindNextFile(hFind, pData)
    If Not bRet Then
      dError = Err.LastDllError
      If dError = ERROR_NO_MORE_FILES Then
        Exit Do
      Else
        ErrorOut Err.LastDllError, "InternetFindNextFile"
        InternetCloseHandle (hFind)
        Exit Function
      End If
    Else
      filename = left(pData.cFileName, InStr(1, pData.cFileName, String(1, vbNullChar), vbBinaryCompare) - 1)
      mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, filename
    End If
  Loop

  Set GetDirectoryListing = mDirCol
  InternetCloseHandle (hFind)
End Function

Public Sub SetTransferASCII()
  TransferType = FTP_TRANSFER_TYPE_ASCII
End Sub

Public Sub SetTransferBinary()
  TransferType = FTP_TRANSFER_TYPE_BINARY
End Sub

Public Sub SetModeActive()
  Semantic = 0
End Sub

Public Sub SetModePassive()
  Semantic = INTERNET_FLAG_PASSIVE
End Sub

Private Sub Class_Initialize()
  mInternet = InternetOpen("FTP Uploader", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If mInternet = 0 Then
    ErrorOut Err.LastDllError, "InternetOpen"
  End If
  TransferType = FTP_TRANSFER_TYPE_ASCII
  TransferType = FTP_TRANSFER_TYPE_BINARY
  Semantic = 0
  mConnection = 0
End Sub

Private Sub Class_Terminate()
  InternetCloseHandle mInternet
End Sub

