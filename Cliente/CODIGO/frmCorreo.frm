VERSION 5.00
Begin VB.Form frmCorreo 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "HOLA"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   5475
      MaxLength       =   5
      TabIndex        =   12
      Text            =   "1"
      Top             =   5385
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.ListBox lstObjsEnviar 
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
      Height          =   1785
      Left            =   6315
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstObjs 
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
      Height          =   1785
      Left            =   2700
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtMensaje 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1830
      Left            =   2715
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1590
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox txtAsunto 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   2715
      MaxLength       =   20
      TabIndex        =   8
      Top             =   930
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox txtDestinatario 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   195
      MaxLength       =   15
      TabIndex        =   7
      Top             =   900
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.ListBox lstContactos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4080
      ItemData        =   "frmCorreo.frx":0000
      Left            =   180
      List            =   "frmCorreo.frx":0040
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.ListBox lstObjetos 
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
      Height          =   1785
      Left            =   2700
      TabIndex        =   5
      Top             =   3660
      Width           =   2500
   End
   Begin VB.TextBox lblMensaje 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   2715
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmCorreo.frx":008B
      Top             =   1440
      Width           =   6200
   End
   Begin VB.ListBox lstMails 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5895
      IntegralHeight  =   0   'False
      ItemData        =   "frmCorreo.frx":02BB
      Left            =   150
      List            =   "frmCorreo.frx":0319
      TabIndex        =   0
      Top             =   630
      Width           =   2415
   End
   Begin VB.Image cmdSalir2 
      Height          =   495
      Left            =   8520
      Top             =   120
      Width           =   465
   End
   Begin VB.Image cmdQui 
      Height          =   375
      Left            =   5580
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdAdd 
      Height          =   375
      Left            =   5580
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdRetirar 
      Height          =   375
      Left            =   2685
      Top             =   5595
      Width           =   2535
   End
   Begin VB.Image cmdNuevo 
      Height          =   495
      Left            =   4305
      Top             =   6150
      Width           =   3015
   End
   Begin VB.Image cmdSalir 
      Height          =   495
      Left            =   8520
      Top             =   120
      Width           =   495
   End
   Begin VB.Image cmdGuardar 
      Height          =   615
      Left            =   5340
      Top             =   5355
      Width           =   3615
   End
   Begin VB.Image cmdBorrar 
      Height          =   615
      Left            =   5340
      Top             =   4515
      Width           =   3615
   End
   Begin VB.Image cmdResponder 
      Height          =   615
      Left            =   5340
      Top             =   3660
      Width           =   3615
   End
   Begin VB.Label lblAsunto 
      BackStyle       =   0  'Transparent
      Caption         =   "Hola si dame oro plis"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "24/11/2808 16:48"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7050
      TabIndex        =   2
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label lblRemitente 
      BackStyle       =   0  'Transparent
      Caption         =   "SuperMiniGnomo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   645
      Width           =   1695
   End
   Begin VB.Image cmdSend 
      Height          =   735
      Left            =   1620
      Top             =   5925
      Visible         =   0   'False
      Width           =   6240
   End
End
Attribute VB_Name = "frmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
correosAgregarItem lstObjs.ListIndex, Val(txtAmount.text)
End Sub
Private Sub cmdQui_Click()
correosQuitarItem lstObjsEnviar.ListIndex, Val(txtAmount.text)
End Sub

Private Sub cmdRetirar_Click()
Call SendData("CZR" & lstMails.ListIndex + 1)
End Sub

Private Sub cmdSend_Click()

If Len(frmCorreo.txtDestinatario.text) < 3 Then
    Mensaje.Escribir ("El destinatario debe tener al minimo 3 letras.")
    Exit Sub
End If

If Len(frmCorreo.txtAsunto.text) < 10 And Len(frmCorreo.txtMensaje.text) < 10 Then
    Mensaje.Escribir ("El asunto y el mensaje deben tener un minimo de 10 letras.")
    Exit Sub
End If

correosEnviarItems
Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_N.jpg")
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_N.jpg")
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_N.jpg")
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_N.jpg")
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_N.jpg")
cmdAdd.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BDERECHA_N.jpg")
cmdQui.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BIZQUIERDA_N.jpg")
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_N.jpg")
End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_Main.jpg")
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_N.jpg")
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_N.jpg")
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_N.jpg")
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_N.jpg")
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_N.jpg")

lblRemitente.Caption = ""
lblRemitente.Visible = True
lblFecha.Caption = ""
lblFecha.Visible = True
lblMensaje.text = ""
lblMensaje.Visible = True
lblAsunto.Caption = ""
lblAsunto.Visible = True
lstMails.Visible = True
lstObjetos.Clear
lstObjetos.Visible = True
cmdResponder.Visible = True
cmdBorrar.Visible = True
cmdGuardar.Visible = True
cmdNuevo.Visible = True
cmdRetirar.Visible = True


cmdAdd.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BDERECHA_N.jpg")
cmdQui.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BIZQUIERDA_N.jpg")
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_N.jpg")

'Sacamos todo
cmdSend.Visible = False
cmdAdd.Visible = False
cmdQui.Visible = False
lstObjs.Visible = False
lstObjsEnviar.Visible = False
lstContactos.Visible = False
txtMensaje.Visible = False
txtAsunto.Visible = False
txtDestinatario.Visible = False
txtAmount.Visible = False

End Sub
Private Sub cmdBorrar_Click()
Call SendData("CZB" & lstMails.ListIndex + 1)
End Sub
Private Sub cmdNuevo_Click()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Main.jpg")

'Ponemos todo en false
lblAsunto.Visible = True
lblRemitente.Visible = False
lblFecha.Visible = False
lblMensaje.Visible = False
lstMails.Visible = False
lstObjetos.Visible = False
cmdResponder.Visible = False
cmdBorrar.Visible = False
cmdGuardar.Visible = False
cmdNuevo.Visible = False
cmdRetirar.Visible = False

'y en true
cmdSend.Visible = True
cmdAdd.Visible = True
cmdQui.Visible = True

lstObjs.Visible = True
lstObjsEnviar.Visible = True
lstContactos.Visible = True

txtMensaje.Visible = True
txtMensaje.text = ""

txtAsunto.Visible = True
txtAsunto.text = ""

txtDestinatario.Visible = True
txtDestinatario.text = ""

txtAmount.Visible = True
txtAmount.text = "1"


End Sub
Private Sub cmdResponder_Click()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Main.jpg")

'Ponemos todo en false
lblAsunto.Visible = True
lblRemitente.Visible = False
lblFecha.Visible = False
lblMensaje.Visible = False
lstMails.Visible = False
lstObjetos.Visible = False
cmdResponder.Visible = False
cmdBorrar.Visible = False
cmdGuardar.Visible = False
cmdNuevo.Visible = False
cmdRetirar.Visible = False

'y en true
cmdSend.Visible = True
cmdAdd.Visible = True
cmdQui.Visible = True

lstObjs.Visible = True
lstObjsEnviar.Visible = True
lstContactos.Visible = True

txtMensaje.Visible = True
txtMensaje.text = ""

txtAsunto.Visible = True
txtAsunto.text = "RE " & lstMails.List(lstMails.ListIndex)

txtDestinatario.Visible = True
txtDestinatario.text = lstMails.List(lstMails.ListIndex)

txtAmount.Visible = True
txtAmount.text = "1"


End Sub
Private Sub cmdSalir2_Click()
    correosCerrar
End Sub
Private Sub lstContactos_Click()
txtDestinatario.text = UCase$(lstContactos.List(lstContactos.ListIndex))
End Sub
Private Sub lstMails_Click()
If lstMails.ListIndex + 1 = 0 Then Exit Sub
If lstMails.ListIndex = CorreoListIndex Then Exit Sub

lblMensaje.text = ""
lblAsunto.Caption = ""
lblFecha.Caption = ""
lblRemitente.Caption = ""
lstObjetos.Clear

CorreoListIndex = lstMails.ListIndex
Call SendData("CZC" & lstMails.ListIndex + 1)

End Sub
Private Sub cmdResponder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_I.jpg")
End Sub
Private Sub cmdResponder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdResponder.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RespMsj_A.jpg")
End Sub
Private Sub cmdBorrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_I.jpg")
End Sub
Private Sub cmdBorrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_BorrarMsj_A.jpg")
End Sub
Private Sub cmdGuardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_I.jpg")
End Sub
Private Sub cmdGuardar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdGuardar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_GuardarMsj_A.jpg")
End Sub
Private Sub cmdRetirar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_I.jpg")
End Sub
Private Sub cmdRetirar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdRetirar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_RetirarObj_A.jpg")
End Sub
Private Sub cmdNuevo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_I.jpg")
End Sub
Private Sub cmdNuevo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNuevo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo2_NuevoMsj_A.jpg")
End Sub
Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BDERECHA_I.jpg")
End Sub
Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BDERECHA_A.jpg")
End Sub
Private Sub cmdQui_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdQui.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BIZQUIERDA_I.jpg")
End Sub
Private Sub cmdQui_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdQui.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_BIZQUIERDA_A.jpg")
End Sub
Private Sub cmdSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_I.jpg")
End Sub
Private Sub cmdSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSend.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Correo1_Enviar_A.jpg")
End Sub
