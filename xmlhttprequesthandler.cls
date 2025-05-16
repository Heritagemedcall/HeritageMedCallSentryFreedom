VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLHTTPRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RequestType  As String

Public LastStatus   As READYSTATE

Public XMLHttpRequest As MSXML2.XMLHTTP60


Enum READYSTATE
  READYSTATE_UNINITIALIZED = 0    'Default initialization state
  READYSTATE_LOADING = 1          'Object is currently loading its properties.
  READYSTATE_LOADED = 2           'Object has been initialized
  READYSTATE_INTERACTIVE = 3      'Object is interactive, but not all of its data is available
  READYSTATE_COMPLETE = 4         'Object has received all of its data
End Enum


Public Event Done()


Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0

  '
  'Highlight the procedure name "OnReadyStateChange" by selecting it in the Code window.
  'From the Tools menu, click Procedure Attributes.
  'In the Procedure Attributes dialog, the Name combo box should show "OnReadyStateChange."
  'Click Advanced.
  'In Procedure ID, select "(Default)" from the available options.
  'Click OK.

  LastStatus = XMLHttpRequest.READYSTATE

  Select Case LastStatus
    Case READYSTATE_COMPLETE
      'Debug.Print "READYSTATE_COMPLETE"
      RaiseEvent Done
    Case READYSTATE_INTERACTIVE
      'Debug.Print "READYSTATE_INTERACTIVE"
    Case READYSTATE_LOADED
      'Debug.Print "READYSTATE_LOADED"
    Case READYSTATE_LOADING
      'Debug.Print "READYSTATE_LOADING"
    Case READYSTATE_UNINITIALIZED
      'Debug.Print "READYSTATE_UNINITIALIZED"
    Case Else
      'Debug.Print "READYSTATE_UNK"
  End Select


  'Debug.Print Form1.XMLHttpRequest.readyState
  'If XMLHttpRequest.READYSTATE = READYSTATE_COMPLETE Then

  'End If
End Sub


Private Sub Class_Terminate()
  If Not (XMLHttpRequest Is Nothing) Then

    XMLHttpRequest.abort
  End If
End Sub
