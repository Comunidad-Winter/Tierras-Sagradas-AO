VERSION 5.00
Begin VB.Form frmGods 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Dioses"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgGods 
      Height          =   1230
      Left            =   330
      Top             =   675
      Width           =   4665
   End
   Begin VB.Label lblOfrecidos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/10000"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2295
      Width           =   4935
   End
   Begin VB.Image imgAlmas 
      Height          =   300
      Left            =   150
      Picture         =   "frmGods.frx":0000
      Top             =   2280
      Width           =   5025
   End
   Begin VB.Label txtValor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   252
      Left            =   1800
      TabIndex        =   1
      Top             =   2715
      Width           =   1815
   End
   Begin VB.Image cmdSalir 
      Height          =   375
      Left            =   4800
      Top             =   120
      Width           =   375
   End
   Begin VB.Image cmdOfrecer 
      Height          =   525
      Left            =   330
      Top             =   3195
      Width           =   4665
   End
End
Attribute VB_Name = "frmGods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOfrecer_Click()
    If txtValor.Caption = 0 Then Unload Me: Exit Sub
    Call SendData("OFDIOZ" & txtValor.Caption)
    Unload Me
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
cmdOfrecer.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_OfrecerN.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOfrecer.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_OfrecerN.jpg")
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

    If KeyAscii = vbKeyBack And Len(txtValor.Caption) = 0 Then Exit Sub

    If KeyAscii = vbKeyBack And Len(txtValor.Caption) <> 0 Then
       txtValor.Caption = mid(txtValor.Caption, 1, Len(txtValor.Caption) - 1)
    Else
        txtValor.Caption = txtValor.Caption & Chr$(KeyAscii)  'convert to character
    End If

End Sub
Private Sub cmdOfrecer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOfrecer.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_OfrecerI.jpg")
End Sub
Private Sub cmdOfrecer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOfrecer.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_OfrecerA.jpg")
End Sub
