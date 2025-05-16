VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Address           As String
Public MacAddress        As String
Public Description       As String
Public ID                As Long
Public ServiceName       As String
Public TCPNoDelay        As Long
Public TCPAckFrequency   As Long

Public IpAddress         As String  ' fixed address
Public DhcpIPAddress     As String      ' assigned address

Public Found             As Boolean
