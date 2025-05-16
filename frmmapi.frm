VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMAPI 
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   11925
   ClientTop       =   11235
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   3600
   Begin MSMAPI.MAPIMessages Messages 
      Left            =   1215
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession Session 
      Left            =   270
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   -1  'True
   End
End
Attribute VB_Name = "frmMAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

