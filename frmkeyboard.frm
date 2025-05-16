VERSION 5.00
Begin VB.Form frmKeyboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keyboard"
   ClientHeight    =   4155
   ClientLeft      =   6195
   ClientTop       =   5325
   ClientWidth     =   12135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Tab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   50
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   51
      Tag             =   "caps"
      Top             =   915
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   49
      Left            =   585
      Style           =   1  'Graphical
      TabIndex        =   50
      Tag             =   "caps"
      Top             =   2535
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   0
      Left            =   10335
      Style           =   1  'Graphical
      TabIndex        =   49
      Tag             =   "back"
      Top             =   2535
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "    |    \"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   48
      Left            =   10845
      Style           =   1  'Graphical
      TabIndex        =   48
      Tag             =   "backslash"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "    }   ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   47
      Left            =   10065
      Style           =   1  'Graphical
      TabIndex        =   47
      Tag             =   "rbracket"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "   {    ["
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   46
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   46
      Tag             =   "lbracket"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Caps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   45
      Left            =   585
      Style           =   1  'Graphical
      TabIndex        =   45
      Tag             =   "caps"
      Top             =   1725
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  +  ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   44
      Left            =   9795
      Style           =   1  'Graphical
      TabIndex        =   44
      Tag             =   "="
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  _  -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   43
      Left            =   9015
      Style           =   1  'Graphical
      TabIndex        =   43
      Tag             =   "-"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   42
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   42
      Tag             =   "enter"
      Top             =   1725
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   41
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "back"
      Top             =   105
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Space"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   40
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "space"
      Top             =   3360
      Width           =   5490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  ?   /"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   39
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "slash"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  >   ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   38
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "period"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  <   ,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   37
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "comma"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   36
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "m"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   35
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "n"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   34
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "b"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   33
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "v"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   32
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "c"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   31
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "x"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   30
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "z"
      Top             =   2535
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "   :   ;"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   29
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "semi"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   28
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "l"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   27
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "k"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   26
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "j"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   25
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "h"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   24
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "g"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   23
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "f"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   22
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "d"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   21
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "s"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   20
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "a"
      Top             =   1725
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   19
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "p"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   18
      Left            =   7725
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "o"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   17
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "i"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   16
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "u"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   15
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "y"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   14
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "t"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   13
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "r"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   12
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "e"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   11
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "w"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   10
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "q"
      Top             =   915
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "   )   0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   9
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "0"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "   (   9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   8
      Left            =   7455
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "9"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  *   8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   7
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "8"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  &&  7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   6
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "7"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  ^   6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   5
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "6"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   " %  5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   4
      Left            =   4335
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "5"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  $  4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   3
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "4"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  #  3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   2
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "3"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   " @  2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   1
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "   !   1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   0
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "1"
      Top             =   105
      Width           =   780
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CapsOn As Boolean


Private Sub cmdKey_Click(index As Integer)
  Select Case cmdKey(index).tag
    Case "1"
      If CapsOn Then
          
      Else
      
      End If
    Case "2"
    
  
  
  End Select
End Sub

Private Sub Command1_Click(index As Integer)
  ResetActivityTime
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ResetActivityTime
End Sub

Private Sub Form_Load()
  ResetActivityTime
  DrawKeyboard
End Sub


Sub DrawKeyboard()

End Sub
