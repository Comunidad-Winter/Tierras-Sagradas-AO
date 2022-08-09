Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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
'La Plata - Pcia, Buenos Aires - Republica Argentina.

'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Movement_Speed As Single

' Api mouse_event
Public Declare Sub mouse_event Lib "user32" _
                        (ByVal dwFlags As Long, _
                        ByVal dX As Long, _
                        ByVal dy As Long, _
                        ByVal cButtons As Long, _
                        ByVal dwExtraInfo As Long)
  
' Constantes para la función mouse_event
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Public UserIpAdress As String
Public InventorySlots As Byte
Public AramSeconds As Integer

Public Type tQuests
    Nombre As String
    Info As String
    Tipo As Byte
    puntos As Long
    Oro As Long
    Dificultad As String
    NivelMinimo As Byte
    Mapas As String
    PosiblesDrops As String
    Usuarios As Integer
    NPCs As Byte
    NumNPC(1 To 3) As Integer
    CantNPC(1 To 3) As Long
    IndexOBJ As Long
    CantOBJ As Long
End Type
Public InfoQuests() As tQuests

Public Const ScreenWidth As Long = 800
Public Const ScreenHeight As Long = 600
Public Const DegreeToRadian As Single = 0.0174532925 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

Public ConteoH As Integer, ConteoW As Integer, TransparenciaCont As Byte
Public datCM() As Byte
Public RetiraObj As Byte
Public RetiraOro As Byte
Public CorreoListIndex As Integer
Public PantallaCompleta As Boolean
Public TieneColorMapa As Boolean

'Renderizacion conectar-cuenta
Public CrearAura(0 To 9) As Boolean
Public MostrarTodo(0 To 9) As Boolean
Public PersonajeX As Integer
Public BodyHeadY As Integer
Public ArmaEscudoY As Integer
Public NickX As Integer
Public NickY As Integer
Public EntrarX As Integer
Public EntrarY As Integer
Public ButtonLogin As String
Public ButtonCC As String
Public ButtonRC As String
Public ButtonVW As String
Public ButtonCP As String
Public ButtonCPass As String
Public ButtonSalir As String
Public ButtonEntrarPJ As String
Public ButtonBorrarPJ As String
Public RenderConnect As Boolean
Public RenderAccount As Boolean

'/CHORI
Public UserMinHPCHORI As Integer
Public UserMaxHPCHORI As Integer
Public UserMinMANCHORI As Integer
Public UserMaxMANCHORI As Integer
Public NickCHORI As String

'ChatContacts
Public NickContacto(1 To 5) As String
Public RecibioMensaje(1 To 5) As Boolean
Public ChatEnUso(1 To 5) As Boolean
Public MouseBarraChat(1 To 5) As Boolean
Public ChatForm(1 To 5) As New frmChatForm
Public VentanitaMostrar(1 To 5) As Byte

'menu
Public nombreotro As String

'CS
Public LiderDueño As String
Public ClanDueño As String
Public HoraComienzo As String
Public ClanesInscriptos As Byte
Public Inscripto As Byte

Public CantidadCanjeYegua As Long
Public RangoPRIV(1 To 6) As String
Public EsStatusCOLOR(0 To 9)

'Casas
Public DueñoKsa As String
Public Preciox As Long
Public Fechix As String

'Mouse
Public MouseOK As Boolean
Public MouseItem As Integer
Public MouseRendOK As Boolean
Public OfMouse As Boolean
Public ButtonIN As Boolean
Public PUEDO As Boolean

'Main
Public UserBOVItem As Long
Public NickPJ As String
Public DibujadoContinuoInv As Boolean
Public DyDActivado As Boolean

'S.O.S
Public UserPrivilegios As Byte
Type tMensajesSos
    Tipo As String
    Autor As String
    Contenido As String
End Type
Public MensajesSOS(1 To 120) As tMensajesSos
Public EsUsuario As String
Public MensajesNumber As Integer
Public TieneParaResponder As Boolean
Public Stopped As Byte

'Denuncias - nuevo
Type tDenuncias
    Tipo As String
    Autor As String
    Contenido As String
    YP As String
    ID As String
    Nick As String
    UltimoLogeo As String
    PrimerDenuncia As String
    UltimaDenuncia As String
    Estado As String
End Type
Public Denuncias(1 To 50) As tDenuncias
Public DenunciasNumber As Integer

'Cuenta Regresiva
Public Cuenta As Boolean
Public Conteo As Long
Public Tiempo As Byte

'Connect - Account
Public Aurix_Angle As Single
Public ClickeoTextCuenta As Boolean
Public ClickeoTextPassw As Boolean
Public TextBoxCuenta As String
Public TextBoxPassw As String
Public TextBoxPasswR As String
Public Const MapConnect = 998
Public Const MapCuent = 999
Public BarritaTextConnect As Byte

Public Type Account_Charge
Nombre As String
Head As Integer
Body As Integer
Shield As Integer
Weapon As Integer
Casco As Integer
Index As Integer
Level As Integer
Clase As String
Existe As Boolean
Raza As String
Muerto As Integer
End Type
Public CargarPJ(0 To 8) As Account_Charge

'Teclas
Public Const NUMBINDS = 22
Public BindKeys(1 To NUMBINDS) As tBindedKey
Public Type tBindedKey
    KeyCode As Integer
    Name As String
End Type

'Consola
Public UserConsola As Boolean

'Carteles
Public CartelInvisibilidad As Integer

'Opciones
Public Type User_Config
    Music As Byte
    Sound As Byte
    FPS As Byte
    Mensajes As Byte
    Desactivar_Globales As Byte
    Desactivar_Privados As Byte
    MisionDiaria As Byte
    MoverPantalla As Byte
    AnunciarContacto As Byte
    MP3Volume As Long
    
    HablaNumerico As Byte
    Interactuar As Byte
    DobleClick As Byte
    MenuDesplegable As Byte
    
    Auras As Byte
    Sombras As Byte
    Particulas As Byte
    Letras_Suben As Byte
    Nombres As Byte
    Transparencias As Byte
    Desvanecimientos As Byte
    Contador As Byte
    ReflejosAgua As Byte
    CartelMuerte As Byte
    VerMiniMapa As Byte
    VerEmoticons As Byte
End Type

Public Configuracion As User_Config
Public tmpConfiguracion As User_Config
Public Interfaces() As String
Public CodigoRecibido As String
Public Nombredelmapaxx As String
Public CantidadDePersonajes As Byte

'd/d
Public AllowDrag As Byte
Public RemDragX As Integer
Public RemDragY As Integer
Public MouseOverMap As Integer

'Botones
Public Enum eButtonStates
    BNormal = 1
    Iluminado = 2
    Apretado = 3
    Bloqueado = 4
End Enum

Public form_Moviment As clsFormMovementManager

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_SNAPSHOT As Byte = 44 ' PrintScreen virtual keycode
Public Const PS_TheForm As Integer = 0
Public Const PS_TheScreen As Byte = 1

Public CustomKeys As New clsCustomKeys
    
Public mode As Boolean
Public temp_rgb(3) As Long
Public LuzGrh(3) As Long
Public AlphaY As Byte

Public MinEleccion As Integer, MaxEleccion As Integer, Actualea As Integer

Public nombrecuent As String
Public passcuent As String
Public Texto            As New clsDX8Font

Public MagMin As Integer
Public MagMax As Integer
Public MagMina As Integer
Public MagMaxa As Integer
Public MagMinb As Integer
Public MagMaxb As Integer
Public MagMinc As Integer
Public MagMaxc As Integer
Public MagMind As Integer
Public MagMaxd As Integer
Public CascMin As Integer
Public CascMax As Integer
Public EscuMin As Integer
Public EscuMax As Integer
Public ArmaMin As Integer
Public ArmaMax As Integer
Public ArmorMin As Integer
Public ArmorMax As Integer
Public HerrMin As Integer
Public HerrMax As Integer

'juanjo
Public PJClickeado As String

Public rcvName As String
Public rcvHead As Integer
Public rcvBody As Integer
Public rcvShield As Integer
Public rcvWeapon As Integer
Public rcvCasco As Integer
Public rcvIndex As Integer
Public rcvCrimi As Boolean
Public rcvBaned As Integer
Public rcvLevel As Integer
Public rcvClase As String
Public rcvMuerto As Integer
Public rcvRaza As String

Public PJSAmount As Integer
'/juanjo

Public HDSerial As String

Public Type TextoDesv
    Text As String
    LifeTime As Long
    StartTime As Long
    Desvanecimiento As Byte
    Existe As Boolean
    Sube As Byte
    X As Integer
    Y As Integer
    color As Long
    Tiempito As Boolean
End Type
Public TextDesv As TextoDesv

'Objetos públicos
Public Dialogos As New cDialogos
Public Audio As New clsAudio
Public Light As New clsLight
Public Inventario As New clsGrapchicalInventory
Public SurfaceDB As clsSurfaceManDynDX8   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 52) As tColor

Public Type tAuras
    GrhIndex As Integer
    r As Byte
    g As Byte
    b As Byte
    Giratoria As Byte
    Offset As Byte
    RojoF As Byte
    AzulF As Byte
    VerdeF As Byte
End Type

Public AurasPJ() As tAuras

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public currentMidi As Long
Public CurrentMP3 As Long

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 0
Public Const tMagia = 0
Public Const tUs = 0

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public VersionC As String

Public UsaMacro As Boolean
Public CnTd As Byte
Public SecuenciaMacroHechizos As Byte
Public UserBancoOro As Long
Public UserBancoOroPropio As Long

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
Public UserBancoInventoryB(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 35
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 20

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521

Public Enum Skills
     Suerte = 1
     Magia = 2
     Robar = 3
     Tacticas = 4
     Armas = 5
     meditar = 6
     Apuñalar = 7
     Ocultarse = 8
     Supervivencia = 9
     Talar = 10
     Comerciar = 11
     Defensa = 12
     Pesca = 13
     Mineria = 14
     Carpinteria = 15
     Herreria = 16
     Liderazgo = 17 ' NOTA: Solia decir "Curacion"
     Domar = 18
     Proyectiles = 19
     Wresterling = 20
     Navegacion = 21
     DefensaMagica = 22
End Enum

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public IsSeguroC As Boolean
Public Const MENSAJE_SEGURO_RESU_ON As String = ">>SEGURO DE RESURRECCION ACTIVADO<<"
Public Const MENSAJE_SEGURO_RESU_OFF As String = ">>SEGURO DE RESURRECCION DESACTIVADO<<"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "
'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    
End Type

Type tEstadisticasUsu
    Clase As String
    Email As String
    Advertencias As Byte
    DuelosGanados As Integer
    DuelosPerdidos As Integer
    CopasDeOro As Integer
    CopasDePlata As Integer
    CopasDeBronce As Integer
    QuestCompletadas As Integer
    CiudadanosMatados As Integer
    CriminalesMatados As Integer
    NPCSMATADOS As Integer
    Jerarquia As String
    Restantes As String
    Alineacion As Byte
    GuerrasGanadas As Integer
    CvcsGanados As Integer
    MVPMatados As Integer
    PuntosTorneo As String
    Hogar As String
    Genero As String
    Nivel As Byte
    Bonif1 As String
    Bonif2 As String
    Bonif3 As String
    Nombre As String
    TipoQuest As Byte
    DescQuest As String
    PremioOro As Long
    PremioPuntis As Integer
    CantidadNPCs As Byte
    YaMatados As Byte
    TorneosParticipados As Integer
    MaximasRondas As Integer
    Eventos As Integer
    ParejasGanadas As Integer
    ParejasPerdidas As Integer
    GuerrasPerdidas As Integer
    NeutralesMatados As Integer
    MuertesUsuario As Integer
    Raza As String
    UserReputacion As Byte
    PuntosDonador As Long
End Type

Type tEstadisticasFrm
    Nivel As Integer
    Faccion As Byte
End Type

Public Nombres As Boolean
Public MixedKey As Long

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserGLD As Long
Public UserReputacione As Long
Public UserLvl As Integer
Public UserPuntosTorneo As Long
Public UserCanAttack As Integer
Public UserCanAttackMagia As Integer
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserEstadisticas As tEstadisticasUsu
Public formuEstadisticas As tEstadisticasFrm
Public UserDescansar As Boolean
Public pausa As Boolean
Public ISItem As Boolean
Public UserParalizado As Boolean
Public TiempoParalizado As Byte
Public UserNavegando As Boolean
Public UserHogar As String
Public UserStatus As Byte
Public nameprivado As String

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String

Public Const NUMSKILLS As Byte = 21
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 16
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Integer
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Integer
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Musica As Boolean
Public Sound As Boolean
Public TimerPing(1 To 2) As Long
Public EnvioFPS As Long

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

'Barrin 30/9/03
Public UserPuedeRefrescar As Boolean

Public UsingSkill As Integer

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearAccount = 4
    LoginAccount = 5
    BorrarPj = 6
    RecuPW = 7
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
'    FXMEDITARCHICO = 4
'    FXMEDITARMEDIANO = 5
'    FXMEDITARGRANDE = 6
'    FXMEDITARXGRANDE = 16
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    TRANSFO = 35
End Enum


'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

''RichTextBox Transparente''
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&
''[END]''

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type



    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Public Const WM_SETTEXT = &HC
    Public Const WM_GETTEXT = &HD
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const EM_SETREADONLY = &HCF
    

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera
