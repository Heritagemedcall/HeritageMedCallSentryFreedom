VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUserPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDeleteTransmitters As Long
Private mDeleteRooms As Long
Private mDeleteResidents As Long

Function ParseUserPermissions(ByVal PermissionValue As Long) As Long
  
mDeleteResidents = IIf(PermissionValue And DELETE_RESIDENTS, 1, 0)
mDeleteRooms = IIf(PermissionValue And DELETE_ROOMS, 1, 0)
mDeleteTransmitters = IIf(PermissionValue And DELETE_TRANSMITTERS, 1, 0)

End Function

Function UnParseUserPermissions() As Long
  Dim Value As Long
  Value = IIf(mDeleteResidents, DELETE_RESIDENTS, 0)
  Value = Value Or IIf(mDeleteRooms, DELETE_ROOMS, 0)
  Value = Value Or IIf(mDeleteTransmitters, DELETE_TRANSMITTERS, 0)

  UnParseUserPermissions = Value

End Function

Function SetUserPermissions(ByVal DeleteTransmitters As Long, ByVal DeleteRooms As Long, ByVal DeleteResidents As Long) As Long
  mDeleteResidents = IIf(DeleteResidents, 1, 0)
  mDeleteRooms = IIf(DeleteRooms, 1, 0)
  mDeleteTransmitters = IIf(DeleteTransmitters, 1, 0)

End Function

Public Property Get CanDeleteResidents() As Long
  CanDeleteResidents = mDeleteResidents
End Property

Public Property Get CanDeleteRooms() As Long
  CanDeleteRooms = mDeleteRooms
End Property

Public Property Get CanDeleteTransmitters() As Long
  CanDeleteTransmitters = mDeleteTransmitters
End Property


Public Property Let CanDeleteResidents(ByVal Value As Long)
  mDeleteResidents = IIf(Value And DELETE_RESIDENTS, 1, 0)
End Property

Public Property Let CanDeleteRooms(ByVal Value As Long)
  mDeleteResidents = IIf(Value And DELETE_ROOMS, 1, 0)
End Property

Public Property Let CanDeleteTransmitters(ByVal Value As Long)
  mDeleteTransmitters = IIf(Value And DELETE_TRANSMITTERS, 1, 0)
End Property

