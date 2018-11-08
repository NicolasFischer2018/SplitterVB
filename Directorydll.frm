VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Directory..."
   ClientHeight    =   3135
   ClientLeft      =   3510
   ClientTop       =   2565
   ClientWidth     =   5340
   Icon            =   "Directorydll.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   2820
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
   Begin VB.Label Label1 
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const vitess As Integer = 2000
Private Sub Command1_Click()
Dim count, count1, list
Dim ficfree1, free2, taille, rest, pos1, pos2
Dim len2, n
Dim len1(255) As Long
Dim string1 As String * vitess
Dim bit As Byte

Reset
count = File1.ListCount
count1 = File1.ListCount
ficfree1 = FreeFile
Open "c:\tmp3\res.dll" For Output As ficfree1
n = 0
Do While count <> 0
len1(n) = FileLen(Dossier.Path + "\" + File1.list(count - 1))
Print #ficfree1, File1.list(count - 1) + " " + Str(len1(n))
count = count - 1
n = n + 1
Loop
Close (ficfree1)
ficfree1 = FreeFile
len2 = FileLen("c:\tmp3\res.dll")
Open "c:\tmp3\res.dll" For Binary Access Write As ficfree1
Seek ficfree1, len2
pos1 = len2
pos2 = 0
count = count1
Do While count <> 0
free2 = FreeFile
Open Dossier.Path + "\" + File1.list(count - 1) For Binary Access Read As free2
taille = Int(LOF(free2) / vitess)
rest = LOF(free2) - taille
Do While Not LOF(free2)
DoEvents
Do While taille <> 0
Get pos2, free2, string1
Put pos1, ficfree1, string1
pos1 = pos1 + vitess
pos2 = pos2 + vitess
taille = taille - 1
Loop
Do While rest <> 0
Get pos2, free2, bit
Put pos1, ficfree1, bit
pos1 = pos1 + 1
pos2 = pos2 + 1
rest = rest - 1
Loop
pos2 = 0
'enter code here
Loop
Close (2)
Loop
Unload Form2
End Sub

Private Sub Command2_Click()
Unload Form2
End Sub

Private Sub Dossier_Change()
Dim count
File1.Path = Dossier.Path
count = File1.ListCount
Label1.Caption = "Nbr de fichier :" + Str(count)
End Sub

Private Sub Form_Load()
Dossier.Path = Lecteur.Drive
End Sub

Private Sub Lecteur_Change()
Dossier.Path = Lecteur.Drive
End Sub
