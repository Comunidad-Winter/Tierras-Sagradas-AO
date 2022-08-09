VERSION 5.00
Begin VB.Form frmViajar 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmViajar.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmViajar.frx":DF4D
      Top             =   2640
      Width           =   6060
   End
   Begin VB.Image Command5 
      Height          =   780
      Left            =   3400
      Picture         =   "frmViajar.frx":DFD3
      Top             =   720
      Width           =   1545
   End
   Begin VB.Image Command4 
      Height          =   780
      Left            =   1800
      Picture         =   "frmViajar.frx":135ED
      Top             =   720
      Width           =   1545
   End
   Begin VB.Image Image2 
      Height          =   780
      Left            =   5040
      Picture         =   "frmViajar.frx":187C7
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Image Command7 
      Height          =   780
      Left            =   5040
      Picture         =   "frmViajar.frx":1DB6E
      Top             =   720
      Width           =   1545
   End
   Begin VB.Image Command6 
      Height          =   780
      Left            =   3405
      Picture         =   "frmViajar.frx":22F6E
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6240
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Command3 
      Height          =   780
      Left            =   1800
      Picture         =   "frmViajar.frx":286CC
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Image Command2 
      Height          =   780
      Left            =   180
      Picture         =   "frmViajar.frx":2DF51
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Image Command1 
      Height          =   780
      Left            =   180
      Picture         =   "frmViajar.frx":335EA
      Top             =   720
      Width           =   1545
   End
End
Attribute VB_Name = "frmViajar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/VIAJAR TANARIS")
Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "La ciudad donde comienzan tus aventuras, para el sur (mapa 18) se encuentra la entrada a las catacumbas. Llendo para el Norte se llega a la ciudad de Thir. La laguna de Tanaris es un lugar de renuion de numerosos aventureros. Los negocios de esta ciudad vende el equipo mas basico para los novatos. El castillo 34 se encuentra al Sur de la ciudad."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisI.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "Esta ciudad se encuentra en el norte del mundo, posee varios negocios y es una buena zona de partida para los aventureros que quieren explorar el polo."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaI.JPG")
Me.Tag = "1"
End Sub
Private Sub Command2_Click()
Call SendData("/VIAJAR ANVILMAR")
Unload Me
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "La ciudad capital de la Alianza Imperial, esta gran ciudad se encuentra en el sur, en el mapa de abajo (Mapa 41) se encuentra el muelle desde donde cada dia parten barcos al peligroso desierto del sur o al castillo 33. Al norte se encuentra otra de las entradas a las Catacumbas. En esta ciudad se vende el mejor equipo disponible a la venta."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarI.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub
Private Sub Command5_Click()
    If MsgBox("Se requiere tener una barca para moverse por este mapa o podés no regresar con vida, ¿Viajar de todas formas?", vbYesNo) = vbYes Then
        Call SendData("/VIAJAR JHUMBEL")
    End If
    Unload Me
End Sub
Private Sub Command3_Click()
Call SendData("/VIAJAR KAHLIMDOR")
Unload Me
End Sub
Private Sub Command4_Click()
Call SendData("/VIAJAR THIR")
Unload Me
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "La terrible ciudad central de la Horda Infernal, se encuentra en el norte cerca de la zona de torneos y el polo, en el mapa de la derecha esta otra de las entradas a la catacumbas y bajando por el mar se llega a un peligroso Dungeon. En esta ciudad se vende el mejor equipo disponible a la venta."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorI.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub
Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "Este pequeño pueblo se encuentra en los bosques, al sur esta el bosque de los osos que es un buen lugar para conseguir pieles. Llendo al norte se encuentra el polo. Le clase de objetos que se venden aca son los mismo que Tanaris."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirI.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub
Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "Este pueblo esta en grupo de islas del mapa 69, es el mejor lugar para ir hacia la peligrosa dungeon del 70. Tiene unos pocos negocios, un cura y un banquero."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelI.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub

Private Sub Command6_Click()
Call SendData("/VIAJAR RUVENDEL")
Unload Me
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "Esta ciudad se encuentra en el medio del desierto del sur, cerca de la peligrosa Piramide de Inthak, posee vendedores de pociones, un cura, un banquero y algunos negocios pequeños."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakI.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub

Private Sub Command7_Click()
Call SendData("/VIAJAR INTHAK")
Unload Me
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = "Esta ciudad se encuentra en el mapa 26, al norte se encuentra un volcan y la entrada al dungeon infernal, al sur se encuentra la entrada a la isla y a la cueva de los osos. En esta ciudad hay varios tipos de negocios y un ring de pelea para los guerreros mas valientes."
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelI.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
Me.Tag = "1"
End Sub

Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
Me.Picture = General_Load_Interface_Picture("Viajar_Main.jpg")
Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Tag = "1" Then
    Me.Tag = "0"
    Command1.Picture = General_Load_Interface_Picture("Viajar_BTanarisN.JPG")
    Command2.Picture = General_Load_Interface_Picture("Viajar_BAnvilmarN.JPG")
    Command3.Picture = General_Load_Interface_Picture("Viajar_BKahlimdorN.JPG")
    Command4.Picture = General_Load_Interface_Picture("Viajar_BThirN.JPG")
    Command5.Picture = General_Load_Interface_Picture("Viajar_BJhumbelN.JPG")
    Command6.Picture = General_Load_Interface_Picture("Viajar_BRuvendelN.JPG")
    Command7.Picture = General_Load_Interface_Picture("Viajar_BInthakN.JPG")
    Image2.Picture = General_Load_Interface_Picture("Viajar_BHelkaN.JPG")
End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
Call SendData("/VIAJAR HELKA")
Unload Me
End Sub
