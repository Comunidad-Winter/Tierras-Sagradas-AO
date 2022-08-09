VERSION 5.00
Begin VB.Form frmMedallas 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   LinkTopic       =   "Form2"
   Picture         =   "frmMedallas.frx":0000
   ScaleHeight     =   4230
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   5
      Left            =   2200
      Top             =   3390
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   0
      Left            =   4320
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   2
      Left            =   120
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   3
      Left            =   4320
      Top             =   2160
      Width           =   2385
   End
   Begin VB.Image PrizeCmd 
      Height          =   660
      Index           =   4
      Left            =   120
      Top             =   2160
      Width           =   2385
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6330
      Top             =   15
      Width           =   495
   End
End
Attribute VB_Name = "frmMedallas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Champ.jpg")
PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Gema.jpg")
PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Puntos.jpg")
PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Sacris.jpg")
PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Slot.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Champ.jpg")
PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Gema.jpg")
PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Puntos.jpg")
PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Sacris.jpg")
PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_Slot.jpg")
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub PrizeCmd_Click(Index As Integer)
Call SendData("GEDS" & Index + 1)
Unload Me
End Sub
Private Sub PrizeCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_ChampI.jpg")
If Index = 2 Then PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_GemaI.jpg")
If Index = 3 Then PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_PuntosI.jpg")
If Index = 4 Then PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_SacrisI.jpg")
If Index = 5 Then PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_SlotI.jpg")
End Sub
Private Sub PrizeCmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then PrizeCmd(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_ChampA.jpg")
If Index = 2 Then PrizeCmd(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_GemaA.jpg")
If Index = 3 Then PrizeCmd(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_PuntosA.jpg")
If Index = 4 Then PrizeCmd(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_SacrisA.jpg")
If Index = 5 Then PrizeCmd(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Medallas_SlotA.jpg")
End Sub

