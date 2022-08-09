VERSION 5.00
Begin VB.Form OpcionesNew 
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   Picture         =   "OpcionesNew.frx":0000
   ScaleHeight     =   6075
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Resetear PJ"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNivel 
      Caption         =   "Editar Nivel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.HScrollBar MP3Volume 
      Height          =   195
      LargeChange     =   50
      Left            =   1440
      Max             =   4000
      TabIndex        =   0
      Top             =   1760
      Value           =   1250
      Width           =   3855
   End
   Begin VB.Image chkEmojis 
      Height          =   180
      Left            =   2520
      Top             =   4600
      Width           =   195
   End
   Begin VB.Image chkMiniMap 
      Height          =   180
      Left            =   2300
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image imgSaveOpt 
      Height          =   420
      Left            =   1800
      Top             =   5200
      Width           =   2085
   End
   Begin VB.Image chkContacto 
      Height          =   180
      Left            =   1490
      Top             =   4040
      Width           =   195
   End
   Begin VB.Image chkModoVentana 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   4110
      Top             =   3550
      Width           =   195
   End
   Begin VB.Image chkMisionDiaria 
      Height          =   180
      Left            =   4230
      Top             =   3270
      Width           =   195
   End
   Begin VB.Image imgSalir 
      Height          =   255
      Left            =   5400
      Top             =   120
      Width           =   375
   End
   Begin VB.Image chkMuerte 
      Height          =   180
      Left            =   2185
      Top             =   4070
      Width           =   195
   End
   Begin VB.Image chkReflejos 
      Height          =   180
      Left            =   2340
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image chkContadores 
      Height          =   180
      Left            =   2900
      Top             =   3545
      Width           =   195
   End
   Begin VB.Image chkDesvanecimiento 
      Height          =   180
      Left            =   3180
      Top             =   3270
      Width           =   195
   End
   Begin VB.Image chkTransparencias 
      Height          =   180
      Left            =   3760
      Top             =   3010
      Width           =   195
   End
   Begin VB.Image chkNombres 
      Height          =   180
      Left            =   1372
      Top             =   2762
      Width           =   195
   End
   Begin VB.Image chkLetrasSuben 
      Height          =   180
      Left            =   3300
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image chkParticulas 
      Height          =   180
      Left            =   1555
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image chkSombras 
      Height          =   180
      Left            =   1312
      Top             =   2030
      Width           =   195
   End
   Begin VB.Image chkAuras 
      Height          =   180
      Left            =   1040
      Top             =   1780
      Width           =   195
   End
   Begin VB.Image chkMenu 
      Height          =   180
      Left            =   5190
      Top             =   3270
      Width           =   195
   End
   Begin VB.Image chkDobleClick 
      Height          =   180
      Left            =   4120
      Top             =   3040
      Width           =   195
   End
   Begin VB.Image chkInteractuar 
      Height          =   180
      Left            =   3300
      Top             =   2780
      Width           =   195
   End
   Begin VB.Image chkModoHabla 
      Height          =   180
      Left            =   4970
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image chkPrivados 
      Height          =   180
      Left            =   2480
      Top             =   3050
      Width           =   195
   End
   Begin VB.Image chkGlobales 
      Height          =   180
      Left            =   2470
      Top             =   2780
      Width           =   195
   End
   Begin VB.Image chkMsj 
      Height          =   180
      Left            =   3690
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image FPS 
      Height          =   180
      Index           =   0
      Left            =   1650
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image FPS 
      Height          =   180
      Index           =   1
      Left            =   2340
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image FPS 
      Height          =   180
      Index           =   3
      Left            =   4200
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image FPS 
      Height          =   180
      Index           =   2
      Left            =   3110
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image chkSound 
      Height          =   180
      Left            =   2300
      Top             =   2020
      Width           =   195
   End
   Begin VB.Image chkMusic 
      Height          =   180
      Left            =   1190
      Top             =   1780
      Width           =   195
   End
   Begin VB.Image ConfigMacros 
      Height          =   480
      Left            =   3615
      Top             =   1850
      Width           =   1395
   End
   Begin VB.Image ConfigTeclas 
      Height          =   480
      Left            =   750
      Top             =   1850
      Width           =   1395
   End
   Begin VB.Image imgRender 
      Height          =   630
      Left            =   3920
      Top             =   750
      Width           =   1770
   End
   Begin VB.Image imgControles 
      Height          =   630
      Left            =   2050
      Top             =   750
      Width           =   1770
   End
   Begin VB.Image imgJuego 
      Height          =   630
      Left            =   200
      Top             =   750
      Width           =   1770
   End
End
Attribute VB_Name = "OpcionesNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ConfigChanged As Boolean
Private MusicChanged As Boolean
Private form_Mov As clsFormMovementManager

Private Sub chkAuras_Click()
    If tmpConfiguracion.Auras = 1 Then
        tmpConfiguracion.Auras = 0
    Else
        tmpConfiguracion.Auras = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Auras, chkAuras)
    ConfigChanged = True
End Sub

Private Sub chkContacto_Click()
    If tmpConfiguracion.AnunciarContacto = 1 Then
        tmpConfiguracion.AnunciarContacto = 0
    Else
        tmpConfiguracion.AnunciarContacto = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.AnunciarContacto, chkContacto)
    ConfigChanged = True
End Sub
Private Sub chkContadores_Click()
    If tmpConfiguracion.Contador = 1 Then
        tmpConfiguracion.Contador = 0
    Else
        tmpConfiguracion.Contador = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Contador, chkContadores)
    ConfigChanged = True
End Sub
Private Sub chkDesvanecimiento_Click()
    If tmpConfiguracion.Desvanecimientos = 1 Then
        tmpConfiguracion.Desvanecimientos = 0
    Else
        tmpConfiguracion.Desvanecimientos = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Desvanecimientos, chkDesvanecimiento)
    ConfigChanged = True
End Sub
Private Sub chkDobleClick_Click()
    If tmpConfiguracion.DobleClick = 1 Then
        tmpConfiguracion.DobleClick = 0
    Else
        tmpConfiguracion.DobleClick = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.DobleClick, chkDobleClick)
    ConfigChanged = True
End Sub

Private Sub chkEmojis_Click()
    If tmpConfiguracion.VerEmoticons = 1 Then
        tmpConfiguracion.VerEmoticons = 0
    Else
        tmpConfiguracion.VerEmoticons = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.VerEmoticons, chkEmojis)
    ConfigChanged = True
End Sub
Private Sub chkGlobales_Click()
    If tmpConfiguracion.Desactivar_Globales = 1 Then
        tmpConfiguracion.Desactivar_Globales = 0
    Else
        tmpConfiguracion.Desactivar_Globales = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Desactivar_Globales, chkGlobales)
    ConfigChanged = True
End Sub
Private Sub chkInteractuar_Click()
    If tmpConfiguracion.Interactuar = 1 Then
        tmpConfiguracion.Interactuar = 0
    Else
        tmpConfiguracion.Interactuar = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Interactuar, chkInteractuar)
    ConfigChanged = True
End Sub
Private Sub chkLetrasSuben_Click()
    If tmpConfiguracion.Letras_Suben = 1 Then
        tmpConfiguracion.Letras_Suben = 0
    Else
        tmpConfiguracion.Letras_Suben = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Letras_Suben, chkLetrasSuben)
    ConfigChanged = True
End Sub
Private Sub chkMenu_Click()
    If tmpConfiguracion.MenuDesplegable = 1 Then
        tmpConfiguracion.MenuDesplegable = 0
    Else
        tmpConfiguracion.MenuDesplegable = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.MenuDesplegable, chkMenu)
    ConfigChanged = True
End Sub

Private Sub chkMiniMap_Click()
    If tmpConfiguracion.VerMiniMapa = 1 Then
        tmpConfiguracion.VerMiniMapa = 0
    Else
        tmpConfiguracion.VerMiniMapa = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.VerMiniMapa, chkMiniMap)
    ConfigChanged = True
End Sub

Private Sub chkMisionDiaria_Click()
    If tmpConfiguracion.MisionDiaria = 1 Then
        tmpConfiguracion.MisionDiaria = 0
    Else
        tmpConfiguracion.MisionDiaria = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.MisionDiaria, chkMisionDiaria)
    ConfigChanged = True
End Sub
Private Sub chkModoHabla_Click()
    If tmpConfiguracion.HablaNumerico = 1 Then
        tmpConfiguracion.HablaNumerico = 0
    Else
        tmpConfiguracion.HablaNumerico = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.HablaNumerico, chkModoHabla)
    ConfigChanged = True
End Sub

Private Sub chkModoVentana_Click()
    If tmpConfiguracion.MoverPantalla = 1 Then
        tmpConfiguracion.MoverPantalla = 0
    Else
        tmpConfiguracion.MoverPantalla = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.MoverPantalla, chkModoVentana)
    ConfigChanged = True
End Sub

Private Sub chkMsj_Click()
    If tmpConfiguracion.Mensajes = 1 Then
        tmpConfiguracion.Mensajes = 0
    Else
        tmpConfiguracion.Mensajes = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Mensajes, chkMsj)
    ConfigChanged = True
End Sub
Private Sub chkMuerte_Click()
    If tmpConfiguracion.CartelMuerte = 1 Then
        tmpConfiguracion.CartelMuerte = 0
    Else
        tmpConfiguracion.CartelMuerte = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.CartelMuerte, chkMuerte)
    ConfigChanged = True
End Sub
Private Sub chkMusic_Click()
    If tmpConfiguracion.Music = 1 Then
        tmpConfiguracion.Music = 0
    Else
        tmpConfiguracion.Music = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Music, chkMusic)
    MusicChanged = True
    ConfigChanged = True
End Sub
Private Sub chkNombres_Click()
    If tmpConfiguracion.Nombres = 1 Then
        tmpConfiguracion.Nombres = 0
    Else
        tmpConfiguracion.Nombres = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Nombres, chkNombres)
    ConfigChanged = True
End Sub
Private Sub chkParticulas_Click()
    If tmpConfiguracion.Particulas = 1 Then
        tmpConfiguracion.Particulas = 0
    Else
        tmpConfiguracion.Particulas = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Particulas, chkParticulas)
    ConfigChanged = True
End Sub
Private Sub chkPrivados_Click()
    If tmpConfiguracion.Desactivar_Privados = 1 Then
        tmpConfiguracion.Desactivar_Privados = 0
    Else
        tmpConfiguracion.Desactivar_Privados = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Desactivar_Privados, chkPrivados)
    ConfigChanged = True
End Sub
Private Sub chkReflejos_Click()
    If tmpConfiguracion.ReflejosAgua = 1 Then
        tmpConfiguracion.ReflejosAgua = 0
    Else
        tmpConfiguracion.ReflejosAgua = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.ReflejosAgua, chkReflejos)
    ConfigChanged = True
End Sub
Private Sub chkSombras_Click()
    If tmpConfiguracion.Sombras = 1 Then
        tmpConfiguracion.Sombras = 0
    Else
        tmpConfiguracion.Sombras = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Sombras, chkSombras)
    ConfigChanged = True
End Sub
Private Sub chkSound_Click()
    If tmpConfiguracion.Sound = 1 Then
        tmpConfiguracion.Sound = 0
    Else
        tmpConfiguracion.Sound = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Sound, chkSound)
    ConfigChanged = True
End Sub
Private Sub chkTransparencias_Click()
    If tmpConfiguracion.Transparencias = 1 Then
        tmpConfiguracion.Transparencias = 0
    Else
        tmpConfiguracion.Transparencias = 1
    End If
    
    Call AplicarTick(tmpConfiguracion.Transparencias, chkTransparencias)
    ConfigChanged = True
End Sub

Private Sub cmdNivel_Click()
 Call SendData("/EDITLVL")
End Sub

Private Sub cmdReset_Click()
    If MsgBox("Se reiniciarán todos los stats de tu personaje, ¿Desea realizar un reset?", vbYesNo) = vbYes Then
        Call SendData("/RRPERSONAJE")
    End If
End Sub

Private Sub ConfigMacros_Click()
frmMakro.Show , frmMain
End Sub
Private Sub ConfigTeclas_Click()
Call frmTeclas.Show(vbModeless, frmMain)
End Sub
Private Sub Form_Load()

Set form_Mov = New clsFormMovementManager
form_Mov.Initialize OpcionesNew

ConfigChanged = False
MusicChanged = False
Call LoadTempOptions
Call CargarMain("JUEGO")

imgSaveOpt.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Guardar_N.jpg")
imgJuego.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Juego_N.jpg")
imgRender.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Render_N.jpg")
imgControles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Controles_N.jpg")

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSaveOpt.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Guardar_N.jpg")
    imgJuego.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Juego_N.jpg")
    imgRender.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Render_N.jpg")
    imgControles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Controles_N.jpg")
    ConfigTeclas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Teclas_N.jpg")
    ConfigMacros.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Macro_N.jpg")
End Sub

Private Sub FPS_Click(Index As Integer)
    Dim i As Long
    For i = 1 To 3
        FPS(i).Picture = Nothing
    Next
    
    If Index = 0 Then
        tmpConfiguracion.FPS = 18
    ElseIf Index = 1 Then
        tmpConfiguracion.FPS = 32
    ElseIf Index = 2 Then
        tmpConfiguracion.FPS = 65
    ElseIf Index = 3 Then
        tmpConfiguracion.FPS = 0
    End If
    
    If tmpConfiguracion.FPS = 18 Then
        FPS(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
    ElseIf tmpConfiguracion.FPS = 32 Then
        FPS(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
    ElseIf tmpConfiguracion.FPS = 65 Then
        FPS(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
    ElseIf tmpConfiguracion.FPS = 0 Then
        FPS(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
    End If
    
    ConfigChanged = True
End Sub
Private Sub imgSalir_Click()
If ConfigChanged = True Then
    If MsgBox("La configuracion ha cambiado, ¿Salir sin guardar?", vbYesNo) = vbYes Then
        Unload Me
    End If
Else
    Unload Me
End If
End Sub
Private Sub imgControles_Click()
    Call CargarMain("CONTROLES")
End Sub
Private Sub imgJuego_Click()
    Call CargarMain("JUEGO")
End Sub
Private Sub imgRender_Click()
    Call CargarMain("RENDER")
End Sub
Private Sub imgRender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRender.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Render_A.jpg")
End Sub
Private Sub imgRender_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRender.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Render_I.jpg")
End Sub
Private Sub imgJuego_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgJuego.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Juego_A.jpg")
End Sub
Private Sub imgJuego_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgJuego.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Juego_I.jpg")
End Sub
Private Sub imgControles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgControles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Controles_A.jpg")
End Sub
Private Sub imgControles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgControles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Controles_I.jpg")
End Sub
Private Sub ConfigTeclas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConfigTeclas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Teclas_A.jpg")
End Sub
Private Sub ConfigTeclas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConfigTeclas.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Teclas_I.jpg")
End Sub
Private Sub ConfigMacros_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConfigMacros.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Macro_A.jpg")
End Sub
Private Sub ConfigMacros_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ConfigMacros.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Macro_I.jpg")
End Sub
Private Sub AplicarTick(Activate As Byte, Aux As Image)
    
    If Activate = 1 Then
        Aux.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
    Else
        Aux.Picture = Nothing
    End If

End Sub
Private Sub CargarMain(Main As String)

Dim i As Long

Select Case UCase$(Main)
    Case "JUEGO"
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_Juego.jpg")
        chkMusic.Visible = True
        Call AplicarTick(tmpConfiguracion.Music, chkMusic)
        chkSound.Visible = True
        Call AplicarTick(tmpConfiguracion.Sound, chkSound)
        chkMsj.Visible = True
        Call AplicarTick(tmpConfiguracion.Mensajes, chkMsj)
        chkGlobales.Visible = True
        Call AplicarTick(tmpConfiguracion.Desactivar_Globales, chkGlobales)
        chkPrivados.Visible = True
        Call AplicarTick(tmpConfiguracion.Desactivar_Privados, chkPrivados)
        chkMisionDiaria.Visible = True
        Call AplicarTick(tmpConfiguracion.MisionDiaria, chkMisionDiaria)
        chkModoVentana.Visible = True
        Call AplicarTick(tmpConfiguracion.MoverPantalla, chkModoVentana)
        chkContacto.Visible = True
        Call AplicarTick(tmpConfiguracion.AnunciarContacto, chkContacto)
        
        MP3Volume.Visible = True
        MP3Volume.Value = Configuracion.MP3Volume
        
        cmdNivel.Visible = True
        cmdReset.Visible = True
                
        'Reiniciamos todos los ticks
            For i = 0 To 3
                FPS(i).Visible = True
                FPS(i).Picture = Nothing
            Next
            
            'Seleccionamos solo el que realmente estamos utilizando.
            If tmpConfiguracion.FPS = 18 Then
                FPS(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
            ElseIf tmpConfiguracion.FPS = 32 Then
                FPS(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
            ElseIf tmpConfiguracion.FPS = 65 Then
                FPS(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
            ElseIf tmpConfiguracion.FPS = 0 Then
                FPS(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Tick2.jpg")
            End If
            
            
        ConfigTeclas.Visible = False
        ConfigMacros.Visible = False
        chkModoHabla.Visible = False
        chkInteractuar.Visible = False
        chkDobleClick.Visible = False
        chkMenu.Visible = False
        chkAuras.Visible = False
        chkSombras.Visible = False
        chkParticulas.Visible = False
        chkLetrasSuben.Visible = False
        chkNombres.Visible = False
        chkTransparencias.Visible = False
        chkDesvanecimiento.Visible = False
        chkContadores.Visible = False
        chkReflejos.Visible = False
        chkMuerte.Visible = False
        chkMiniMap.Visible = False
        chkEmojis.Visible = False
        
    Case "CONTROLES"
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_Controles.jpg")
        ConfigTeclas.Visible = True
        ConfigMacros.Visible = True
        chkModoHabla.Visible = True
        Call AplicarTick(tmpConfiguracion.HablaNumerico, chkModoHabla)
        chkInteractuar.Visible = True
        Call AplicarTick(tmpConfiguracion.Interactuar, chkInteractuar)
        chkDobleClick.Visible = True
        Call AplicarTick(tmpConfiguracion.DobleClick, chkDobleClick)
        chkMenu.Visible = True
        Call AplicarTick(tmpConfiguracion.MenuDesplegable, chkMenu)
        
        
        MP3Volume.Visible = False
        chkMusic.Visible = False
        chkSound.Visible = False
        chkMsj.Visible = False
        chkGlobales.Visible = False
        chkPrivados.Visible = False
        For i = 0 To 3
            FPS(i).Visible = False
        Next
        chkMisionDiaria.Visible = False
        chkModoVentana.Visible = False
        chkContacto.Visible = False
        chkAuras.Visible = False
        chkSombras.Visible = False
        chkParticulas.Visible = False
        chkLetrasSuben.Visible = False
        chkNombres.Visible = False
        chkTransparencias.Visible = False
        chkDesvanecimiento.Visible = False
        chkContadores.Visible = False
        chkReflejos.Visible = False
        chkMuerte.Visible = False
        chkMiniMap.Visible = False
        chkEmojis.Visible = False
        cmdNivel.Visible = False
        cmdReset.Visible = False
        
    Case "RENDER"
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opciones_Render.jpg")
        chkAuras.Visible = True
        Call AplicarTick(tmpConfiguracion.Auras, chkAuras)
        chkSombras.Visible = True
        Call AplicarTick(tmpConfiguracion.Sombras, chkSombras)
        chkParticulas.Visible = True
        Call AplicarTick(tmpConfiguracion.Particulas, chkParticulas)
        chkLetrasSuben.Visible = True
        Call AplicarTick(tmpConfiguracion.Letras_Suben, chkLetrasSuben)
        chkNombres.Visible = True
        Call AplicarTick(tmpConfiguracion.Nombres, chkNombres)
        chkTransparencias.Visible = True
        Call AplicarTick(tmpConfiguracion.Transparencias, chkTransparencias)
        chkDesvanecimiento.Visible = True
        Call AplicarTick(tmpConfiguracion.Desvanecimientos, chkDesvanecimiento)
        chkContadores.Visible = True
        Call AplicarTick(tmpConfiguracion.Contador, chkContadores)
        chkReflejos.Visible = True
        Call AplicarTick(tmpConfiguracion.ReflejosAgua, chkReflejos)
        chkMuerte.Visible = True
        Call AplicarTick(tmpConfiguracion.CartelMuerte, chkMuerte)
        chkMiniMap.Visible = True
        Call AplicarTick(tmpConfiguracion.VerMiniMapa, chkMiniMap)
        chkEmojis.Visible = True
        Call AplicarTick(tmpConfiguracion.VerEmoticons, chkEmojis)
        
        MP3Volume.Visible = False
        chkMusic.Visible = False
        chkSound.Visible = False
        chkMsj.Visible = False
        chkGlobales.Visible = False
        chkPrivados.Visible = False
        For i = 0 To 3
            FPS(i).Visible = False
        Next
        chkMisionDiaria.Visible = False
        chkModoVentana.Visible = False
        chkContacto.Visible = False
        ConfigTeclas.Visible = False
        ConfigMacros.Visible = False
        chkModoHabla.Visible = False
        chkInteractuar.Visible = False
        chkDobleClick.Visible = False
        chkMenu.Visible = False
        cmdNivel.Visible = False
        cmdReset.Visible = False
End Select
        
End Sub
Public Sub LoadOptions()

    Dim l_file As clsIniReader
    Set l_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\UserConfig.ini"
    
    Configuracion.Music = l_file.GetValue("OPTIONS", "Config_Music")
    Configuracion.Sound = l_file.GetValue("OPTIONS", "Config_Sound")
    Configuracion.FPS = l_file.GetValue("OPTIONS", "Config_FPS")
    Configuracion.Mensajes = l_file.GetValue("OPTIONS", "Config_Mensajes")
    Configuracion.Desactivar_Globales = l_file.GetValue("OPTIONS", "Config_Globales")
    Configuracion.Desactivar_Privados = l_file.GetValue("OPTIONS", "Config_Privados")
    Configuracion.MisionDiaria = l_file.GetValue("OPTIONS", "Config_MisionDiaria")
    Configuracion.MoverPantalla = l_file.GetValue("OPTIONS", "Config_ModoVentana")
    Configuracion.AnunciarContacto = l_file.GetValue("OPTIONS", "Config_Contactos")
    Configuracion.MP3Volume = l_file.GetValue("OPTIONS", "Config_MP3Volume")
    
    Configuracion.HablaNumerico = l_file.GetValue("OPTIONS", "Config_HablaNumerico")
    Configuracion.Interactuar = l_file.GetValue("OPTIONS", "Config_InteractuarDobleClick")
    Configuracion.DobleClick = l_file.GetValue("OPTIONS", "Config_DobleClick")
    Configuracion.MenuDesplegable = l_file.GetValue("OPTIONS", "Config_MenuDesplegable")
    
    Configuracion.Auras = l_file.GetValue("OPTIONS", "Config_Auras")
    Configuracion.Sombras = l_file.GetValue("OPTIONS", "Config_Sombras")
    Configuracion.Particulas = l_file.GetValue("OPTIONS", "Config_Particulas")
    Configuracion.Letras_Suben = l_file.GetValue("OPTIONS", "Config_LetrasSuben")
    Configuracion.Nombres = l_file.GetValue("OPTIONS", "Config_Nombres")
    Configuracion.Transparencias = l_file.GetValue("OPTIONS", "Config_Transparencias")
    Configuracion.Desvanecimientos = l_file.GetValue("OPTIONS", "Config_Desvanecimiento")
    Configuracion.Contador = l_file.GetValue("OPTIONS", "Config_Contadores")
    Configuracion.ReflejosAgua = l_file.GetValue("OPTIONS", "Config_ReflejosAgua")
    Configuracion.CartelMuerte = l_file.GetValue("OPTIONS", "Config_CartelMuerte")
    Configuracion.VerMiniMapa = l_file.GetValue("OPTIONS", "Config_MiniMap")
    Configuracion.VerEmoticons = l_file.GetValue("OPTIONS", "Config_Emoticons")
End Sub
Private Sub LoadTempOptions()
    'Guardamos el actual, en un temporal para trabajarlo dentro del formulario
    tmpConfiguracion = Configuracion
End Sub
Private Sub SaveOptions()

    Dim l_file As clsIniReader
    Set l_file = New clsIniReader
    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\UserConfig.ini"
    
    Configuracion = tmpConfiguracion
    
    l_file.ChangeValue "OPTIONS", "Config_Music", Configuracion.Music
    l_file.ChangeValue "OPTIONS", "Config_Sound", Configuracion.Sound
    l_file.ChangeValue "OPTIONS", "Config_FPS", Configuracion.FPS
    l_file.ChangeValue "OPTIONS", "Config_Mensajes", Configuracion.Mensajes
    l_file.ChangeValue "OPTIONS", "Config_Globales", Configuracion.Desactivar_Globales
    l_file.ChangeValue "OPTIONS", "Config_Privados", Configuracion.Desactivar_Privados
    l_file.ChangeValue "OPTIONS", "Config_MisionDiaria", Configuracion.MisionDiaria
    l_file.ChangeValue "OPTIONS", "Config_ModoVentana", Configuracion.MoverPantalla
    l_file.ChangeValue "OPTIONS", "Config_Contactos", Configuracion.AnunciarContacto
    l_file.ChangeValue "OPTIONS", "Config_MP3Volume", Configuracion.MP3Volume
    
    l_file.ChangeValue "OPTIONS", "Config_HablaNumerico", Configuracion.HablaNumerico
    l_file.ChangeValue "OPTIONS", "Config_InteractuarDobleClick", Configuracion.Interactuar
    l_file.ChangeValue "OPTIONS", "Config_DobleClick", Configuracion.DobleClick
    l_file.ChangeValue "OPTIONS", "Config_MenuDesplegable", Configuracion.MenuDesplegable
    
    l_file.ChangeValue "OPTIONS", "Config_Auras", Configuracion.Auras
    l_file.ChangeValue "OPTIONS", "Config_Sombras", Configuracion.Sombras
    l_file.ChangeValue "OPTIONS", "Config_Particulas", Configuracion.Particulas
    l_file.ChangeValue "OPTIONS", "Config_LetrasSuben", Configuracion.Letras_Suben
    l_file.ChangeValue "OPTIONS", "Config_Nombres", Configuracion.Nombres
    l_file.ChangeValue "OPTIONS", "Config_Transparencias", Configuracion.Transparencias
    l_file.ChangeValue "OPTIONS", "Config_Desvanecimiento", Configuracion.Desvanecimientos
    l_file.ChangeValue "OPTIONS", "Config_Contadores", Configuracion.Contador
    l_file.ChangeValue "OPTIONS", "Config_ReflejosAgua", Configuracion.ReflejosAgua
    l_file.ChangeValue "OPTIONS", "Config_CartelMuerte", Configuracion.CartelMuerte
    l_file.ChangeValue "OPTIONS", "Config_MiniMap", Configuracion.VerMiniMapa
    l_file.ChangeValue "OPTIONS", "Config_Emoticons", Configuracion.VerEmoticons
    
    If Configuracion.Music = 0 And MusicChanged = True Then
        Audio.MP3_Stop
        Audio.MP3_Destroy
    End If
    
    If Configuracion.Sound = 0 Then
        Audio.StopMidi
        Audio.StopWave
    End If
    
    l_file.DumpFile App.Path & "\Data\INIT\UserConfig.ini"

End Sub
Private Sub imgSaveOpt_Click()
    SaveOptions
    Unload Me
End Sub
Private Sub imgSaveOpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSaveOpt.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Guardar_A.jpg")
End Sub
Private Sub imgSaveOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSaveOpt.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Opc_Guardar_I.jpg")
End Sub

Private Sub MP3Volume_Change()
    tmpConfiguracion.MP3Volume = MP3Volume.Value
    ConfigChanged = True
End Sub
