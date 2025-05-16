VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RoomID       As Long
Public Room         As String
Public Description  As String
Public Building     As String
Public Assurdays    As Long
Public Away         As Long
Public Vacation     As Integer ' not used
Public Deleted      As Long
Public locKW        As String
Public Flags        As Long

Public Function Parse(rs As ADODB.Recordset) As Long
  
      RoomID = rs("RoomID")
      Room = rs("room") & ""
      'Description = rs("Description") & ""
      Building = rs("building") & ""
      Assurdays = rs("AssurDays")
      Away = rs("Away")
      'Vacation = 0
      locKW = Trim$(rs("lockw") & "")
      Deleted = Val(rs("Deleted") & "")
      Flags = Val(rs("Flags") & "")
  
End Function

