VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C508455C-BC19-11D0-B028-444553540000}#2.0#0"; "ACTIVEZIPPER.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Splitter VB (build 009) : New Features Version ..."
   ClientHeight    =   2280
   ClientLeft      =   2490
   ClientTop       =   2280
   ClientWidth     =   4695
   Icon            =   "splitterVB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4695
   Begin ActiveZipperOcx.ActiveZipper zip2 
      Left            =   1920
      Top             =   0
      _ExtentX        =   661
      _ExtentY        =   450
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2280
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Height          =   350
      Left            =   2700
      MaskColor       =   &H00FF0000&
      Picture         =   "splitterVB.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   290
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "splitterVB.frx":09C4
      Left            =   3120
      List            =   "splitterVB.frx":09D1
      TabIndex        =   10
      Text            =   "Choose Size"
      Top             =   310
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "! Desplit !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Text            =   "1400000"
      Top             =   310
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   0
      TabIndex        =   2
      Text            =   "c:\tmp4\"
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   0
      TabIndex        =   1
      Text            =   "First Choose Here !!!       ------->"
      Top             =   290
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "! Split !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Size :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   3120
      TabIndex        =   12
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "(Check for Specify)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   150
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Destination Directory or where to desplit your files (001.CHK must be present !!!)  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   650
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "File to Split :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnuoption 
      Caption         =   "&Option"
      Begin VB.Menu mnuzip 
         Caption         =   "&Zip + Split"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuunzip 
         Caption         =   "&Unsplit + Unzip"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnudesplit 
      Caption         =   "&Desplit From Floppy"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "                                    &AbOuT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const vitess As Long = 2000

Private Sub Check1_Click()
If Check1.Value = 1 Then
Combo1.Visible = False
Label3.Caption = "(In Bytes)"
Else
Label3.Caption = "(Check for Specify)"
Combo1.Visible = True
End If
End Sub

Private Sub Command1_Click()
Dim taille, pos1, pos2, sizeof As Long
Dim temp1 As Byte
Dim temp As String * vitess
Dim path, kenny, param1, param As String
Dim nbr, coco, n, fic, rest, free, divi, reste, k, l, m, percent

kenny = "normal"
nbr = 0
pos1 = 1
pos2 = 1
temp = 1

Command1.Caption = "En cours ...."
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Combo1.Enabled = False
Command3.Enabled = False
Label4.Caption = Form1.Caption

Reset

Open Text1.Text For Binary Access Read As 1
sizeof = LOF(1)
If Check1.Value = 1 Then
taille = Val(Text3.Text)
Else
If Combo1.ListIndex = -1 Then
MsgBox "Pas de taille: Veuillez en selectionnez une", vbOKOnly
GoTo endofsplit
End If
If Combo1.ListIndex = 0 Then
taille = 1400000
End If
If Combo1.ListIndex = 1 Then
taille = 2000000
End If
If Combo1.ListIndex = 2 Then
taille = 96000000
End If
End If

If sizeof < taille Then
MsgBox "Taille incorecte : Veuillez entrer une autre taille", vbOKOnly
GoTo endofsplit
End If

fic = LOF(1) / taille
fic = Int(fic) 'nbr de fichiers a creer
m = 100 / fic
rest = LOF(1) - (fic * taille) 'nbr d'octets pour le dernier fichier
path = Text2.Text + "001.CHK"
Open path For Output As 2
param = Text1.Text

Do While Left(Right(param, i), 1) <> "\"
param1 = Right(param, i)
i = i + 1
Loop

param1 = Trim(Str(fic + 1)) + " " + param1 + " 1"
Print #2, param1
Close 2

n = 0
Do While Not EOF(1)
DoEvents
Progress.Value = n

If nbr < 9 Then
path = Text2.Text + "00" + Trim(Str(nbr + 1))
End If
If nbr >= 9 Then
path = Text2.Text + "0" + Trim(Str(nbr + 1))
End If
If nbr >= 99 Then
path = Text2.Text + Trim(Str(nbr + 1))
End If

free = FreeFile
Open path For Binary Access Write As free
divi = Int(taille / vitess)
reste = taille - (divi * vitess)
percent = divi + reste
percent = (1 / percent) * 100
coco = 0
Do While (pos2 - 1) <> taille
DoEvents
Do While k <> divi
Get 1, pos1, temp
Put free, pos2, temp
pos1 = pos1 + vitess
pos2 = pos2 + vitess
k = k + 1
coco = coco + percent
Form1.Caption = "Work in Progress : " + Left(Str(coco), 5) + "%"
Loop
Do While l <> reste
Get 1, pos1, temp1
Put free, pos2, temp1
pos1 = pos1 + 1
pos2 = pos2 + 1
l = l + 1
coco = coco + percent
'Form1.Caption = "Work in Progress : " + Left(Str(coco), 5) + "%"
Loop
k = 0
l = 0
Loop
pos2 = 1
nbr = nbr + 1
'Form1.Caption = "Fichier " + Str(nbr) + " crée"
If nbr = fic Then
taille = rest
 End If
Close free
n = n + m
If n > 100 Then
n = 100
End If
Loop
Form1.Caption = Label4.Caption
Progress.Value = 100
Close
Kill path

MsgBox "Les Fichiers sont crées !", vbOKOnly
If Label4.Caption = "1" Then
End
End If
GoTo endofsplit

Exit Sub
endofsplit:
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Command2.Enabled = True
Command1.Caption = "! Split !"
Command1.Enabled = True
Combo1.Enabled = True
Command3.Enabled = True
End Sub
Private Sub Command2_Click()
Dim nbr, free, i, strlen, boucle, taille, n, reste, divi, k, l, m, percent, coco
Dim path, param, param1, param2 As String
Dim pos1, pos2 As Long
Dim temp As String * vitess
Dim temp1 As Byte

n = 0
nbr = 0
pos1 = 1
pos2 = 1

Command2.Caption = "En cours ...."
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Combo1.Enabled = False
Command3.Enabled = False
Label4.Caption = Form1.Caption
Reset

path = Text2.Text + "001.CHK"
Open path For Input As 2
Line Input #2, param
strlen = Len(param)

Do While Right(Left(param, n), 1) <> " "
param1 = Left(param, n)
n = n + 1
Loop
 
param = Right(param, strlen - n)

Do While Right(Left(param, n), 1) <> " "
param2 = Left(param, n)
n = n + 1
Loop

path = Text2.Text + param2
Open path For Binary Access Write As 1
boucle = Val(param1)
m = 100 / boucle
n = 0

Do While nbr <> boucle
DoEvents
Progress.Value = n
free = FreeFile

If nbr < 9 Then
path = Text2.Text + "00" + Trim(Str(nbr + 1))
End If
If nbr >= 9 Then
path = Text2.Text + "0" + Trim(Str(nbr + 1))
End If
If nbr >= 99 Then
path = Text2.Text + Trim(Str(nbr + 1))
End If

Open path For Binary Access Read As free
taille = LOF(free)
divi = Int(taille / vitess)
rest = taille - (divi * vitess)
percent = divi + rest
percent = ((1 / percent) * 100) / boucle
coco = 0
Do While (pos1 - 1) <> taille
DoEvents
Do While k <> divi
Get free, pos1, temp
Put 1, pos2, temp
pos1 = pos1 + vitess
pos2 = pos2 + vitess
k = k + 1
coco = coco + percent
'Form1.Caption = "Work in Progress : " + Left(Str(coco), 5) + "%"
Loop
Do While l <> rest
Get free, pos1, temp1
Put 1, pos2, temp1
pos1 = pos1 + 1
pos2 = pos2 + 1
l = l + 1
coco = coco + percent
'Form1.Caption = "Work in Progress : " + Left(Str(coco), 5) + "%"
Loop
Loop
nbr = nbr + 1
pos1 = 1
k = 0
l = 0
Close free
n = n + m
Loop
Close
Progress.Value = 100
Reset
Form1.Caption = Label4.Caption
If Label4.Caption <> "unzip" Then
MsgBox "Fichier " + param2 + " Crée", vbOKOnly
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Command2.Enabled = True
Command2.Caption = "! Desplit !"
Command1.Enabled = True
Combo1.Enabled = True
Command3.Enabled = True
Else
Label4.Caption = param2
End If
If Label4.Caption = "1" Then
End
End If
End Sub

Private Sub Command3_Click()
CD1.DialogTitle = "Choose File to Split :"
CD1.Filter = "All Files"
On Error GoTo kenny
CD1.ShowOpen
Text1.Text = CD1.FileName
Exit Sub
kenny:
End Sub

Private Sub Form_Load()
Dim argv, param1(4), tampon As String
Dim param As String * 255
Dim n, m, lenstr, advance

argv = Command
'argv = "-h"
LSet param = argv
m = 1
n = 1
advance = 0
lenstr = Len(argv) + 1

' Traitement des Parametres de la ligne de commande
If argv <> "" Then
Check1.Value = 1
Label4.Caption = "1"
Form1.WindowState = 1
Form1.Visible = False

' Decorticage des parametres de la ligne de commande
Do While advance <> lenstr
tampon = "start"
    Do While tampon <> " "
    tampon = Right(Left(param, m), 1)
    m = m + 1
    advance = advance + 1
    Loop
param1(n) = Trim(Left(param, m - 1))
param = Trim(Right(param, (Len(param) - (m - 1))))
n = n + 1
m = 1
Loop

'Interpretation des parametres
    'Pour le premier parametres !!!
    If Left(param1(1), 1) <> "-" Then
    GoTo error
    Else
    tampon = Right(param1(1), 1)
        If tampon = "s" Then
            If Val(param1(2)) <> 0 Then
            Text3.Text = param1(2)
            Else
            GoTo error
            End If
            If param1(4) <> "" Then
            Text2.Text = param1(4)
            Else
            GoTo error
            End If
            If param1(3) <> "" Then
            Text1.Text = param1(4) + param1(3)
            Else
            GoTo error
            End If
            Command1_Click
            End
        Else
            If tampon = "m" Then
            If param1(2) <> "" Then
            Text2.Text = param1(2)
            Else
            GoTo error
            End If
            Command2_Click
            End
            Else
                If tampon = "h" Then
                Form3.Show
                Else
                    If tampon = "v" Then
                    MsgBox Form1.Caption, vbOKOnly
                    End
                    Else
                    GoTo error
                    End If
                End If
            End If
        End If
    End If
End If
Exit Sub
error:
MsgBox "Mauvais Parametres !!", vbOKOnly
End
End Sub

Private Sub Form_Terminate()
Reset
End Sub

Private Sub mnuabout_Click()
About.Show
End Sub

Private Sub mnudesplit_Click()
Dim result, coco

result = MsgBox("Insert Disk1, Please", vbOKCancel)
MsgBox "Not implented Yet : Sorry :(", vbOKOnly
If result = 1 Then

'coco = fileop.CopyFile("a:\001.chk", Text2.Text + "001.chk")
End If
End Sub

Private Sub mnuunzip_Click()
Dim path As String
MsgBox "Not implented Yet : Sorry :(", vbOKOnly
'Label4.Caption = "unzip"
'Command2_Click

''Mettre effacement des fichiers 001, 002, etc...

'path = Text2.Text + Label4.Caption
'zip2.SourceFile = path
'zip2.OutPutFile = "c:\tmp4"
'zip2.Decompress
End Sub

Private Sub mnuzip_Click()
Dim success As Integer
MsgBox "Warning !!! : Bad Zip Generation", vbOKOnly
Load Form2
Form2.Show
End Sub

