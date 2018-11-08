VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parameters Help"
   ClientHeight    =   5175
   ClientLeft      =   3255
   ClientTop       =   1740
   ClientWidth     =   3210
   Icon            =   "ParametersHelp.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&OK, I've read this help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   4800
      Width           =   3230
   End
   Begin VB.Label Label4 
      Caption         =   """SplitterVB  -h""   ----> Show this help !!! :)))"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   14
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Fourth Option :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   """SplitterVB  -v""   ----> Give the version n°"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Give the version number of SplitterVB."
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   11
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Third Option :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "For Exemple : ""-m  c:\tmp\"""
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "To Unsplit a file.(001.CHK must be in path)"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "To Split a file in many files with defined size."
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   """-m  <pathofthefile>"""
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Second Option :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "For Exemple: ""-s 1400000 temp.zip c:\tmp\"""
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   """-s <sizeinBytes><filetosplit><pathofthefile>"""
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "First Option :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   $"ParametersHelp.frx":08CA
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "There are 4 different options :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form3
End
End Sub
