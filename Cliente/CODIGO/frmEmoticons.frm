VERSION 5.00
Begin VB.Form frmEmoticons 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Emoticons"
   ClientHeight    =   2130
   ClientLeft      =   3840
   ClientTop       =   5220
   ClientWidth     =   4140
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   1200
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   22
      Left            =   240
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   21
      Left            =   720
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   20
      Left            =   1200
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   19
      Left            =   1680
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   18
      Left            =   2160
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   17
      Left            =   2640
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   16
      Left            =   3600
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   15
      Left            =   240
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   14
      Left            =   720
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   13
      Left            =   1200
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   12
      Left            =   1680
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   11
      Left            =   2160
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   10
      Left            =   2640
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   3600
      Top             =   600
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   3120
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   720
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   1680
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   2160
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   2640
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   3120
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   3120
      Top             =   600
      Width           =   315
   End
End
Attribute VB_Name = "frmEmoticons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Image1(22).Picture = LoadPicture(DirGraficos & "20554.bmp")
Image1(21).Picture = LoadPicture(DirGraficos & "20553.bmp")
Image1(20).Picture = LoadPicture(DirGraficos & "20552.bmp")
Image1(19).Picture = LoadPicture(DirGraficos & "20551.bmp")
Image1(18).Picture = LoadPicture(DirGraficos & "20532.bmp")
Image1(17).Picture = LoadPicture(DirGraficos & "20549.bmp")
Image1(16).Picture = LoadPicture(DirGraficos & "20548.bmp")
Image1(15).Picture = LoadPicture(DirGraficos & "20547.bmp")
Image1(14).Picture = LoadPicture(DirGraficos & "20546.bmp")
Image1(13).Picture = LoadPicture(DirGraficos & "20545.bmp")
Image1(12).Picture = LoadPicture(DirGraficos & "20544.bmp")
Image1(11).Picture = LoadPicture(DirGraficos & "20543.bmp")
Image1(10).Picture = LoadPicture(DirGraficos & "20542.bmp")
Image1(9).Picture = LoadPicture(DirGraficos & "20541.bmp")
Image1(8).Picture = LoadPicture(DirGraficos & "20540.bmp")
Image1(7).Picture = LoadPicture(DirGraficos & "20539.bmp")
Image1(6).Picture = LoadPicture(DirGraficos & "20538.bmp")
Image1(5).Picture = LoadPicture(DirGraficos & "20537.bmp")
Image1(4).Picture = LoadPicture(DirGraficos & "20536.bmp")
Image1(3).Picture = LoadPicture(DirGraficos & "20535.bmp")
Image1(2).Picture = LoadPicture(DirGraficos & "20534.bmp")
Image1(1).Picture = LoadPicture(DirGraficos & "20533.bmp")

End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 1
SendData (";" & ":S")

Case 2
SendData (";" & ":(")

Case 3
SendData (";" & ":CA")

Case 4
SendData (";" & ";)")

Case 5
SendData (";" & ":$")

Case 6
SendData (";" & ">.>")

Case 7
SendData (";" & "?")

Case 8
SendData (";" & "!")

Case 9
SendData (";" & "...")

Case 10
SendData (";" & "¬¬")

Case 11
SendData (";" & ":@")

Case 12
SendData (";" & "º_º")

Case 13
SendData (";" & "-_-")

Case 14
SendData (";" & ":3")

Case 15
SendData (";" & "^^")

Case 16
SendData (";" & ":D")

Case 17
SendData (";" & ":P")

Case 18
SendData (";" & "'_'")

Case 19
SendData (";" & ":O")

Case 20
SendData (";" & "xD")

Case 21
SendData (";" & ":'(")

Case 22
SendData (";" & ":)")

End Select

Unload Me

End Sub
