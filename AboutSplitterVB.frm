VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About me and my prog !!!!!!!"
   ClientHeight    =   2580
   ClientLeft      =   2355
   ClientTop       =   1830
   ClientWidth     =   4965
   Icon            =   "AboutSplitterVB.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "AboutSplitterVB.frx":08CA
   ScaleHeight     =   2580
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "!! Freeware !!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1810
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1830
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2580
      Left            =   0
      Picture         =   "AboutSplitterVB.frx":24A5
      ScaleHeight     =   2520
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   480
         Top             =   600
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"AboutSplitterVB.frx":4DEC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload About
End Sub

Private Sub Timer1_Timer()
If Command1.Caption = "!! Freeware !!" Then
Command1.Caption = ""
Else
Command1.Caption = "!! Freeware !!"
End If
End Sub
