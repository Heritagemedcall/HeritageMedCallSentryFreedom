VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDuke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DukeSerial   As String
Public DukeID       As Long
Public DukeType     As Long
Public DukeStatus   As Long
Public DukeMessage  As String
Public DukePacket   As String

Public Enum DUKE_EVENT
  BAD = -1
  UNK = 0
  CALL_PLACE = 1
  CALL_CLEAR = 2
  CALL_CANCEL = 3
End Enum


Function ToESPacket() As cESPacket
  Dim packet As cESPacket
  Set packet = New cESPacket

  packet.DateTime = Now
  packet.Alarm = DukeStatus
  packet.PacketType = 1
  packet.SerialPacket = DukePacket

  packet.Serial = "D" & DukeType & Right$("000000" & DukeID, 6)
  
  packet.SetMIDClassPTI (&HD0 + DukeType), DukeType, &HDD

  Set ToESPacket = packet

  Set packet = Nothing


End Function


'Public Property Get AlarmStatus() As Long
'  Select Case DukeStatus
'    Case DUKE_EVENT.CALL_PLACE
'      AlarmStatus = 1
'    Case DUKE_EVENT.CALL_CANCEL
'      AlarmStatus = 0
'    Case DUKE_EVENT.CALL_CLEAR
'      AlarmStatus = 0
'    Case Else
'      AlarmStatus = 0
'  End Select
'End Property

Public Function ParsePacket(ByVal packet As String) As Long
  Dim Fields() As String

  Fields = Split(packet, vbTab, 4)
  If UBound(Fields) < 3 Then
    ParsePacket = -1  ' failure
  Else
    DukeSerial = Right$(Fields(0), 6)
    DukeID = Val(DukeSerial)
    DukeType = Val(Fields(1))
    DukeMessage = Trim(left$(Fields(2), 13))
    ParsePacket = 1
    ' DukeStatus = DukeMessageToStatus(DukeMessage)
  End If

End Function




'Public Function DukeMessageToStatus(ByVal Message As String) As Long
'  If InStr(1, Message, "CALL_PLACE", vbTextCompare) > 0 Then
'    DukeMessageToStatus = DUKE_EVENT.CALL_PLACE  ' alarm
'  ElseIf InStr(1, Message, "CALL_CLEAR", vbTextCompare) > 0 Then
'    DukeMessageToStatus = DUKE_EVENT.CALL_CLEAR  ' cleared
'  ElseIf InStr(1, Message, "CANCEL", vbTextCompare) > 0 Then
'    DukeMessageToStatus = DUKE_EVENT.CALL_CANCEL  ' Cancelled
'  Else
'    DukeMessageToStatus = UNK
'  End If
'
'
'End Function
'
'Public Function DukeStatusToMessage(ByVal Status As Long) As String
'  Select Case Status
'    Case DUKE_EVENT.BAD
'      DukeStatusToMessage = "ERROR"
'    Case DUKE_EVENT.CALL_CANCEL
'      DukeStatusToMessage = "CALL_CANCEL"
'    Case DUKE_EVENT.CALL_CLEAR
'      DukeStatusToMessage = "CALL_CLEAR"
'    Case DUKE_EVENT.CALL_PLACE
'      DukeStatusToMessage = "CALL_PLACE"
'    Case Else
'      DukeStatusToMessage = "UNK"
'  End Select
'
'
'End Function
