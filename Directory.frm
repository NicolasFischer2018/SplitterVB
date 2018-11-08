VERSION 5.00
Object = "{C508455C-BC19-11D0-B028-444553540000}#2.0#0"; "ACTIVEZIPPER.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Directory..."
   ClientHeight    =   3135
   ClientLeft      =   3510
   ClientTop       =   2565
   ClientWidth     =   3300
   Icon            =   "Directory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin ActiveZipperOcx.ActiveZipper zipper 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   661
      _ExtentY        =   450
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3360
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.DriveListBox Lecteur 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin VB.DirListBox Dossier 
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim count, list

count = File1.ListCount
Do While count <> 0
zipper.SourceFile = (Dossier.path + "\" + File1.list(count - 1))
zipper.OutPutFile = Dossier.path + "\splitter.zip"
zipper.Compress
count = count - 1
Loop
MsgBox "Zip file create But broken...", vbOKOnly
Unload Form2
End Sub

Private Sub Command2_Click()
Unload Form2
End Sub

Private Sub Dossier_Change()
File1.path = Dossier.path
End Sub

Private Sub Form_Load()
Dossier.path = Lecteur.Drive
End Sub

Private Sub Lecteur_Change()
Dossier.path = Lecteur.Drive
End Sub
