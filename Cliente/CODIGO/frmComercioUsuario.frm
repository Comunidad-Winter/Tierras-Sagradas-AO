VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNuevoComercio 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComercioUsuario.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Can 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4600
      TabIndex        =   10
      Text            =   "1"
      Top             =   6280
      Width           =   1545
   End
   Begin VB.ListBox Ofrecer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      IntegralHeight  =   0   'False
      Left            =   6720
      TabIndex        =   8
      Top             =   2040
      Width           =   2250
   End
   Begin VB.ListBox TusItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4125
      Left            =   3840
      TabIndex        =   7
      Top             =   2040
      Width           =   2250
   End
   Begin VB.ListBox Oferta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      IntegralHeight  =   0   'False
      ItemData        =   "frmComercioUsuario.frx":29182
      Left            =   600
      List            =   "frmComercioUsuario.frx":29184
      TabIndex        =   5
      Top             =   1980
      Width           =   2320
   End
   Begin VB.TextBox Texto 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      MaxLength       =   150
      TabIndex        =   3
      Text            =   "Escribi aca tu mensaje para el otro usuario y apreta la tecla 'Enter' o clickea en Enviar"
      Top             =   8430
      Width           =   6855
   End
   Begin VB.PictureBox picInv 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   550
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   830
      Width           =   540
   End
   Begin RichTextLib.RichTextBox Consola 
      Height          =   735
      Left            =   375
      TabIndex        =   2
      Top             =   7600
      Width           =   8750
      _ExtentX        =   15425
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmComercioUsuario.frx":29186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label conQuien 
      BackStyle       =   0  'Transparent
      Caption         =   "Con tu vieja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Command2 
      Height          =   390
      Left            =   6220
      Top             =   3240
      Width           =   405
   End
   Begin VB.Image Command1 
      Height          =   390
      Left            =   6220
      Top             =   4070
      Width           =   405
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   6285
      Width           =   1665
   End
   Begin VB.Image Command3 
      Height          =   435
      Left            =   6680
      Top             =   6630
      Width           =   2415
   End
   Begin VB.Image cmdAgregarOro 
      Height          =   435
      Left            =   3750
      Top             =   6630
      Width           =   2415
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Envia tu oferta de oro o items al otro usuario."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   7695
   End
   Begin VB.Image Res 
      Height          =   435
      Index           =   1
      Left            =   1800
      Top             =   6630
      Width           =   1215
   End
   Begin VB.Image Res 
      Height          =   435
      Index           =   2
      Left            =   530
      Top             =   6630
      Width           =   1215
   End
   Begin VB.Label lblOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   6280
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   7660
      Top             =   8380
      Width           =   1215
   End
   Begin VB.Label Image2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmNuevoComercio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarOro_Click()

If Not IsNumeric(Can.text) Then Exit Sub
If UserGLD < Can.text Then lblEstado.Caption = "No tienes suficiente oro.": Exit Sub

If Can.text > 0 And Can.text <= 999999999 Then
    Label4.Caption = PonerPuntos(Can.text)
    uOro = Can.text
End If

End Sub
Private Sub Command1_Click()
comAgregarOferta TusItems.ListIndex, Val(Can.text)
End Sub
Private Sub Command2_Click()
comQuitarOferta Ofrecer.ListIndex, Val(Can.text)
End Sub
Private Sub Command3_Click()
If uOro = "0" And Ofrecer.ListCount = 0 Then
    Mensaje.Escribir "Hace una oferta primero."
Exit Sub
End If
    'Desahablita
    Command3.Enabled = False
    Command3.Picture = General_Load_Interface_Picture("ComercioPJ_Ofrecer_Bloqueado.jpg")
    'Cambia lbl
    lblEstado.Caption = "Enviando Ofertas al otro usuario y esperando respuesta..."
comEnviarOferta
End Sub
Private Sub Form_Load()
comMensaje "Bienvenido al nuevo sistema de comercio de TSAO, elija los items y haga su oferta, responda la del otro usuario y listo!, para usar el chat escriba el mensaje y aprete ""ENVIAR"" o tan solo ""ENTER"" y listo.", 255, 255, 0, 0, 0
lblOro.Caption = "0"
Label4.Caption = "0"
uOro = 0
rOro = 0

Me.Picture = General_Load_Interface_Picture("ComercioPJ.jpg")
Image1.Picture = General_Load_Interface_Picture("ComercioPJ_Enviar.jpg")
Command3.Picture = General_Load_Interface_Picture("ComercioPJ_Ofrecer.jpg")
cmdAgregarOro.Picture = General_Load_Interface_Picture("ComercioPJ_Oro.jpg")
Command1.Picture = General_Load_Interface_Picture("ComercioPJ_Flecha_Derecha.jpg")
Command2.Picture = General_Load_Interface_Picture("ComercioPJ_Flecha_Izquierda.jpg")

Res(2).Picture = General_Load_Interface_Picture("ComercioPJ_Aceptar_Bloqueado.jpg")
Res(1).Picture = General_Load_Interface_Picture("ComercioPJ_Rechazar_Bloqueado.jpg")

cmdAgregarOro.Enabled = True
Command3.Enabled = True
Command2.Enabled = True
Command1.Enabled = True
Res(2).Enabled = False
Res(1).Enabled = False

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

End Sub
Private Sub cmdAgregarOro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAgregarOro.Enabled = True Then cmdAgregarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Oro_Mouse.jpg")
End Sub
Private Sub cmdAgregarOro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAgregarOro.Enabled = True Then cmdAgregarOro.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Oro_Apretado.jpg")
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Command3.Enabled = True Then Command3.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Ofrecer_Mouse.jpg")
End Sub
Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Command3.Enabled = True Then Command3.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Ofrecer_Apretado.jpg")
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Command1.Enabled = True Then Command1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Flecha_Derecha_Mouse.jpg")
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Command1.Enabled = True Then Command1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Flecha_Derecha_Apretado.jpg")
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Command2.Enabled = True Then Command2.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Flecha_Izquierda_Mouse.jpg")
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Command2.Enabled = True Then Command2.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Flecha_Izquierda_Apretado.jpg")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Enviar_Mouse.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Enviar_Apretado.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = General_Load_Interface_Picture("ComercioPJ_Enviar.jpg")
If Command3.Enabled = True Then Command3.Picture = General_Load_Interface_Picture("ComercioPJ_Ofrecer.jpg")
If cmdAgregarOro.Enabled = True Then cmdAgregarOro.Picture = General_Load_Interface_Picture("ComercioPJ_Oro.jpg")
If Command1.Enabled = True Then Command1.Picture = General_Load_Interface_Picture("ComercioPJ_Flecha_Derecha.jpg")
If Command2.Enabled = True Then Command2.Picture = General_Load_Interface_Picture("ComercioPJ_Flecha_Izquierda.jpg")

If Res(2).Enabled = True Then Res(2).Picture = General_Load_Interface_Picture("ComercioPJ_Aceptar.jpg")
If Res(1).Enabled = True Then Res(1).Picture = General_Load_Interface_Picture("ComercioPJ_Rechazar.jpg")
End Sub
Private Sub Image1_Click()
SendData "VHC" & Texto.text
comMensaje UserName & " >> " & Texto.text, 255, 255, 255, True, False, False
Texto.text = ""
End Sub
Private Sub Image2_Click()
SendData "TCM"
End Sub
Private Sub Oferta_Click()
comDibujarRec Oferta.ListIndex
End Sub
Private Sub Ofrecer_Click()
comDibujarOfe
End Sub
Private Sub Res_Click(Index As Integer)
        lblEstado.Caption = "Oferta del otro usuario aceptada."
        comRespuesta Index
End Sub
Private Sub Res_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then Res(2).Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Aceptar_Mouse.jpg")
    If Index = 1 Then Res(1).Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Rechazar_Mouse.jpg")
End Sub
Private Sub Res_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then Res(2).Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Aceptar_Apretado.jpg")
    If Index = 1 Then Res(1).Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Rechazar_Apretado.jpg")
End Sub
Private Sub Can_Change()
On Error GoTo errHandler
    If Val(Can.text) < 0 Then Can.text = 10000
Exit Sub
errHandler:
    Can.text = "1"
End Sub
Private Sub Can_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End If
End Sub
Private Sub Texto_Click()

If Texto.text = "Escribi aca tu mensaje para el otro usuario y apreta la tecla 'Enter' o clickea en Enviar" Then
Texto.text = ""
End If

End Sub
Private Sub Texto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
SendData "VHC" & Texto.text
comMensaje "" & UserName & "> " & Texto.text, 255, 0, 0
Texto.text = ""
End If
End Sub
Private Sub TusItems_Click()
comDibujarTusItems TusItems.ListIndex
End Sub

