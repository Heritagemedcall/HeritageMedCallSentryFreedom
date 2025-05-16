VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   4410
   ClientTop       =   4695
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin MSComctlLib.ProgressBar progress 
         Height          =   195
         Left            =   1575
         TabIndex        =   5
         Top             =   2760
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Image imgCareConnect 
         Height          =   1620
         Left            =   1230
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4920
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3345
         TabIndex        =   2
         Top             =   2445
         Width           =   690
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   30
         TabIndex        =   1
         Top             =   3195
         Width           =   7005
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3248
         TabIndex        =   3
         Top             =   2115
         Width           =   885
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3233
         TabIndex        =   4
         Top             =   1755
         Width           =   915
      End
      Begin VB.Image imgTechConn 
         Appearance      =   0  'Flat
         Height          =   1695
         Left            =   1433
         Picture         =   "frmSplash.frx":BD52
         Stretch         =   -1  'True
         Top             =   30
         Width           =   4515
      End
      Begin VB.Image imgHeritage 
         Height          =   1620
         Left            =   1230
         Picture         =   "frmSplash.frx":F9E7
         Stretch         =   -1  'True
         Top             =   30
         Width           =   4920
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblProductName.Caption = PRODUCT_NAME
    Caption = IIf(MASTER, "HOST ", "REMOTE CONSOLE ") & PRODUCT_NAME
    #If brookdale Then
      imgHeritage.Visible = False
      imgCareConnect.Visible = False
      imgTechConn.Visible = True
     
      
        
    #ElseIf esco Then
      
      imgTechConn.Visible = False
      imgHeritage.Visible = False
      imgCareConnect.Visible = True
      
    #Else
      imgTechConn.Visible = False
      imgCareConnect.Visible = False
      imgHeritage.Visible = True
      
     
    #End If
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & IIf(MASTER, "", " Remote Console")
    
    lblCopyright.Caption = App.LegalCopyright
    lblWarning.Caption = " Warning: This computer program is protected by copyright law and international treaties.  Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe civil and criminal penalties, and will be prosecuted to the maximum extent possible under the law."
    
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

