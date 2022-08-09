VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   1275
      TabIndex        =   4
      Top             =   600
      Width           =   3120
   End
   Begin VB.PictureBox headview 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2490
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   6540
      Width           =   495
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3450
      Width           =   2700
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00D1
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":00DB
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   2700
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00EE
      Left            =   8160
      List            =   "frmCrearPersonaje.frx":0101
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2055
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   4410
      Picture         =   "frmCrearPersonaje.frx":012E
      Top             =   585
      Width           =   330
   End
   Begin VB.Image Genero 
      Height          =   480
      Index           =   1
      Left            =   2790
      Top             =   1395
      Width           =   480
   End
   Begin VB.Image Genero 
      Height          =   480
      Index           =   0
      Left            =   2235
      Top             =   1395
      Width           =   480
   End
   Begin VB.Image Raza 
      Height          =   480
      Index           =   4
      Left            =   3810
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Raza 
      Height          =   480
      Index           =   3
      Left            =   3150
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Raza 
      Height          =   480
      Index           =   2
      Left            =   2520
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Raza 
      Height          =   480
      Index           =   1
      Left            =   1920
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Raza 
      Height          =   480
      Index           =   0
      Left            =   1275
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   11
      Left            =   3735
      Top             =   5595
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   10
      Left            =   2520
      Top             =   5595
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   9
      Left            =   1680
      Top             =   1.50000e5
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   8
      Left            =   1230
      Top             =   5595
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   7
      Left            =   4350
      Top             =   4605
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   6
      Left            =   3120
      Top             =   4605
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   5
      Left            =   1905
      Top             =   4605
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   4
      Left            =   660
      Top             =   4605
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   3
      Left            =   4350
      Top             =   3675
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   2
      Left            =   3120
      Top             =   3675
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   1
      Left            =   1905
      Top             =   3675
      Width           =   480
   End
   Begin VB.Image Clase 
      Height          =   480
      Index           =   0
      Left            =   675
      Top             =   3675
      Width           =   480
   End
   Begin VB.Image menoshead 
      Height          =   375
      Left            =   2160
      Top             =   6570
      Width           =   255
   End
   Begin VB.Image mashead 
      Height          =   375
      Left            =   3045
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image boton 
      Height          =   600
      Index           =   1
      Left            =   2880
      MouseIcon       =   "frmCrearPersonaje.frx":05C2
      MousePointer    =   99  'Custom
      Top             =   7125
      Width           =   2400
   End
   Begin VB.Image boton 
      Height          =   600
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmCrearPersonaje.frx":0714
      MousePointer    =   99  'Custom
      Top             =   7125
      Width           =   2400
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public SkillPoints As Byte

Function CheckData() As Boolean

If UserName = "" Then
    MsgBox "Asigne nombre a su personaje."
    Exit Function
End If

If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

CheckData = True


End Function
Private Sub boton_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
     Case 0
       
        UserName = txtNombre.text
       
       If Len(txtNombre.text) < 4 Then
        MsgBox "El nombre debe de tener mas de 4 caracteres!!"
        Exit Sub
    End If
     
If Len(txtNombre.text) >= 16 Then
    MsgBox "El nombre debe de tener menos de 15 caracteres!!"
    Exit Sub
End If

Dim AllCr As Long
Dim CantidadEsp As Byte
Dim thiscr As String

Do
    AllCr = AllCr + 1
    If AllCr > Len(UserName) Then Exit Do
    thiscr = mid(UserName, AllCr, 1)
    If InStr(1, " ", UCase(thiscr)) = 1 Then
           CantidadEsp = CantidadEsp + 1
    End If
Loop

If CantidadEsp > 1 Then
     Mensaje.Escribir "El nombre no puede tener mas de 1 espacio."
     Exit Sub
End If

        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
       
        UserRaza = lstRaza.List(lstRaza.ListIndex)
        UserSexo = lstGenero.List(lstGenero.ListIndex)
        UserClase = lstProfesion.List(lstProfesion.ListIndex)
        UserHogar = "Tanaris"
       
        'Barrin 3/10/03
        If CheckData() Then
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
   
    Me.MousePointer = 11
    
    EstadoLogin = CrearNuevoPj
  If Not frmMain.Socket1.Connected Then
        frmMain.Socket1.Connect
    End If
    
        PJClickeado = UserName
        Call Login
    End If
       
        
    Case 1

      Unload Me
End Select


End Sub
Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function
Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\CrearPJ_Main.jpg")

'Clases
Clase(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Mago_Normal.jpg")
Clase(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Guerrero_Normal.jpg")
Clase(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Paladin_Normal.jpg")
Clase(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Clerigo_Normal.jpg")
Clase(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Asesino_Normal.jpg")
Clase(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Bardo_Normal.jpg")
Clase(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Druida_Normal.jpg")
Clase(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Cazador_Normal.jpg")
Clase(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Ladron_Normal.jpg")
Clase(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Pirata_Normal.jpg")
Clase(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Artesano_Normal.jpg")
Clase(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Recolector_Normal.jpg")

'Razas
Raza(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Humano_Normal.jpg")
Raza(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Elfo_Normal.jpg")
Raza(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\ElfoOscuro_Normal.jpg")
Raza(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Enano_Normal.jpg")
Raza(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Gnomo_Normal.jpg")

'Genero
Genero(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Hombre_Normal.jpg")
Genero(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Mujer_Normal.jpg")

'Botones
boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\CrearPersonaje_Normal.jpg")
boton(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\VolverAtras_Normal.jpg")
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Nick_NoDisponible.jpg")


Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 1

End Sub
Private Sub MenosHead_Click()

Call Audio.PlayWave(SND_CLICK)

Actualea = Actualea - 1

If Actualea > MaxEleccion Then
Actualea = MaxEleccion

ElseIf Actualea < MinEleccion Then
Actualea = MinEleccion

End If

Dim SR As RECT
SR.bottom = 32
SR.Right = 32
SR.left = 0
SR.top = 0

Call engine.DrawGrhtoHdc(HeadData(Actualea).Head(3).GrhIndex, SR, headview, 8, 5)

End Sub
Private Sub MasHead_Click()

Call Audio.PlayWave(SND_CLICK)

Actualea = Actualea + 1

If Actualea > MaxEleccion Then
Actualea = MaxEleccion

ElseIf Actualea < MinEleccion Then
Actualea = MinEleccion

End If

Dim SR As RECT
SR.bottom = 32
SR.Right = 32
SR.left = 0
SR.top = 0

Call engine.DrawGrhtoHdc(HeadData(Actualea).Head(3).GrhIndex, SR, headview, 8, 5)

End Sub
Private Sub Raza_Click(Index As Integer)

If Index = 0 Then
    Raza(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Elfo_Normal.jpg")
    Raza(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\ElfoOscuro_Normal.jpg")
    Raza(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Enano_Normal.jpg")
    Raza(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Gnomo_Normal.jpg")
    
    Raza(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Humano_Apretado.jpg")
    lstRaza.text = "Humano"
ElseIf Index = 1 Then
    Raza(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Humano_Normal.jpg")
    Raza(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\ElfoOscuro_Normal.jpg")
    Raza(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Enano_Normal.jpg")
    Raza(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Gnomo_Normal.jpg")
    
    Raza(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Elfo_Apretado.jpg")
    lstRaza.text = "Elfo"
ElseIf Index = 2 Then
    Raza(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Humano_Normal.jpg")
    Raza(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Elfo_Normal.jpg")
    Raza(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Enano_Normal.jpg")
    Raza(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Gnomo_Normal.jpg")
    
    Raza(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\ElfoOscuro_Apretado.jpg")
    lstRaza.text = "Elfo Oscuro"
ElseIf Index = 3 Then
    Raza(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Humano_Normal.jpg")
    Raza(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Elfo_Normal.jpg")
    Raza(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\ElfoOscuro_Normal.jpg")
    Raza(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Gnomo_Normal.jpg")
    
    Raza(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Enano_Apretado.jpg")
    lstRaza.text = "Enano"
ElseIf Index = 4 Then
    Raza(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Humano_Normal.jpg")
    Raza(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Elfo_Normal.jpg")
    Raza(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\ElfoOscuro_Normal.jpg")
    Raza(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Enano_Normal.jpg")
    
    Raza(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Gnomo_Apretado.jpg")
    lstRaza.text = "Gnomo"
End If

Call DameOpciones

End Sub
Private Sub Genero_Click(Index As Integer)

If Index = 0 Then
   Genero(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Hombre_Apretado.jpg")
   Genero(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Mujer_Normal.jpg")
   lstGenero.text = "Hombre"
ElseIf Index = 1 Then
   Genero(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Hombre_Normal.jpg")
   Genero(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Mujer_Apretado.jpg")
   lstGenero.text = "Mujer"
End If

Call DameOpciones

End Sub
Private Sub Clase_Click(Index As Integer)

If Index = 0 Then
    Call LimpiarBotones
    Clase(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Mago_Apretado.jpg")
    lstProfesion.text = "Mago"
ElseIf Index = 1 Then
    Call LimpiarBotones
    Clase(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Guerrero_Apretado.jpg")
    lstProfesion.text = "Guerrero"
ElseIf Index = 2 Then
    Call LimpiarBotones
    Clase(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Paladin_Apretado.jpg")
    lstProfesion.text = "Paladin"
ElseIf Index = 3 Then
    Call LimpiarBotones
    Clase(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Clerigo_Apretado.jpg")
    lstProfesion.text = "Clerigo"
ElseIf Index = 4 Then
    Call LimpiarBotones
    Clase(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Asesino_Apretado.jpg")
    lstProfesion.text = "Asesino"
ElseIf Index = 5 Then
    Call LimpiarBotones
    Clase(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Bardo_Apretado.jpg")
    lstProfesion.text = "Bardo"
ElseIf Index = 6 Then
    Call LimpiarBotones
    Clase(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Druida_Apretado.jpg")
    lstProfesion.text = "Druida"
ElseIf Index = 7 Then
    Call LimpiarBotones
    Clase(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Cazador_Apretado.jpg")
    lstProfesion.text = "Cazador"
ElseIf Index = 8 Then
    Call LimpiarBotones
    Clase(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Ladron_Apretado.jpg")
    lstProfesion.text = "Ladron"
ElseIf Index = 9 Then
    Call LimpiarBotones
    Clase(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Pirata_Apretado.jpg")
    lstProfesion.text = "Pirata"
ElseIf Index = 10 Then
    Call LimpiarBotones
    Clase(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Artesano_Apretado.jpg")
    lstProfesion.text = "Artesano"
ElseIf Index = 11 Then
    Call LimpiarBotones
    Clase(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Recolector_Apretado.jpg")
    lstProfesion.text = "Recolector"
End If

End Sub
Private Sub LimpiarBotones()
Clase(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Mago_Normal.jpg")
Clase(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Guerrero_Normal.jpg")
Clase(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Paladin_Normal.jpg")
Clase(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Clerigo_Normal.jpg")
Clase(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Asesino_Normal.jpg")
Clase(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Bardo_Normal.jpg")
Clase(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Druida_Normal.jpg")
Clase(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Cazador_Normal.jpg")
Clase(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Ladron_Normal.jpg")
Clase(9).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Pirata_Normal.jpg")
Clase(10).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Artesano_Normal.jpg")
Clase(11).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Recolector_Normal.jpg")
End Sub
Private Sub txtNombre_Change()
txtNombre.text = LTrim(txtNombre.text)

       If Len(txtNombre.text) < 4 Or Len(txtNombre.text) >= 16 Or Right$(txtNombre.text, 1) = " " Or Not AsciiValidos(txtNombre.text) Then
            Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Nick_NoDisponible.jpg")
       Else
            Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\Nick_Disponible.jpg")
       End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\CrearPersonaje_Normal.jpg")
boton(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\VolverAtras_Normal.jpg")
End Sub
Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\CrearPersonaje_Iluminado.jpg")
If Index = 1 Then boton(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\VolverAtras_Iluminado.jpg")
End Sub
Private Sub boton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then boton(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\CrearPersonaje_Apretado.jpg")
If Index = 1 Then boton(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\InterfazCP\VolverAtras_Apretado.jpg")
End Sub
