VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDirList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As New Collection

Public Function Add(Attrib As Long, CreationTime As Date, LastAccessTime As Date, LastWriteTime As Date, FileSize As Long, Filename As String) As cDirItem
  Dim newItem As New cDirItem

  With newItem
    .Archive = (Attrib And FILE_ATTRIBUTE_ARCHIVE)
    .Compressed = (Attrib And FILE_ATTRIBUTE_COMPRESSED)
    .Directory = (Attrib And FILE_ATTRIBUTE_DIRECTORY)
    .Hidden = (Attrib And FILE_ATTRIBUTE_HIDDEN)
    .Normal = (Attrib And FILE_ATTRIBUTE_NORMAL)
    .Offline = (Attrib And FILE_ATTRIBUTE_OFFLINE)
    .ReadOnly = (Attrib And FILE_ATTRIBUTE_READONLY)
    .System = (Attrib And FILE_ATTRIBUTE_SYSTEM)
    .Temporary = (Attrib And FILE_ATTRIBUTE_TEMPORARY)
    .CreationTime = CreationTime
    .LastAccessTime = LastAccessTime
    .LastWriteTime = LastWriteTime
    .FileSize = FileSize
    .Filename = Filename
  End With
  mCol.Add newItem, Filename
End Function

Public Function Clear()
  Set mCol = New Collection
End Function

Public Function Item(Index As Variant) As cDirItem
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(Index)
End Function

Public Function Count() As Long
  Count = mCol.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Function

