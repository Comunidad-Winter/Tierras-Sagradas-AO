Attribute VB_Name = "Declaraciones"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public ObjSlot1 As Byte
Public ObjSlot2 As Byte

Public BOnlines As Integer
Public aDos As New clsAntiDoss

'Portales de dioses
Public AlmasNecesarias As Long

'Misiones diarias perrito
Public MisionesDiarias() As tMDiarias
Public ArrayExp(1 To 59) As Long

Public Type tMDiarias
    Nombre As String
    Info As String
    Dia As Byte
    Tipo As Byte
    Cantidad As Long
    QuestNumber As Byte
    NumNPC As Integer
    PTS As Integer
    pOro As Long
    pObjetoIndex As Integer
    pObjetoAmount As Integer
    pPuntos As Long
    ConteoUser As Long
    NumeroMision As Byte
    Completada As Byte
End Type

'Mensaje automatico
Public MinutitosMensaje As Integer
Public MensajeAutomatico As Boolean
Public TextoMensajeAutomatico As String
Public TiempoMensajeAutomatico As Integer

Public TanaTelep As WorldPos

'Guerras
Public HayGuerraAnvil As Boolean
Public HayGuerraKhalim As Boolean
Public Minus As Byte

Public CofresAzar() As AzarCofres

Public Type AzarCofres
    'Cofres random
    CantObjs As Integer
    ObjIndex() As Integer
    ObjAmount() As Integer
    ObjProbability() As Integer
End Type
    
'Donaciones
Public DonationList() As tDonaciones
Public Type tDonaciones
    ObjName As String
    ObjValor As Integer
    NumObjs As Byte
    Body As Integer
    Arma As Integer
    Escudo As Integer
    Casco As Integer
    Aura As Byte
    GrhIndex As Integer
    Tuniquita As String
    Tuniquita2 As String
    Tuniquita3 As String
    Tuniquita4 As String
End Type

Public ReyGuerraIndex As Integer
Public IndexReyAncalagon As Integer
Public UsuarioRevisado As String
Public RevisandoUsuario As Boolean
Public FragsJerarquia(1 To 4) As Integer
Public NPCInvocaciones(1 To 9) As Integer
Public ActivarTimerRejas As Boolean

'Multiplicaciones
Public MultiplicadorOro As Long
Public MultiplicadorExp As Long
Public MultiplicadorDrop As Long

'S.O.S
Type tMensajesSos
    Tipo As String
    Autor As String
    Contenido As String
End Type
Public MensajesSOS(1 To 50) As tMensajesSos
Public MensajesNumber As Integer

'/FORTALEZA
Public ArieteUno As Integer
Public ArieteDos As Integer
Public ArieteTres As Integer
Public RejaNorte As Integer
Public RejaCentral As Integer
Public RejaSur As Integer
Public RejaNorteAtacada As Boolean
Public RejaCentralAtacada As Boolean
Public RejaSurAtacada As Boolean
'/FORTALEZA

Public MinutosGuerras As Byte
Public HayGuerra As Boolean
Public GuardianesHordas As Byte
Public GuardianesAlianza As Byte

Public CronologiaParticipantes(1 To 64) As String
Public CronologiaParticipantesList(1 To 64) As String

'Duelos
Public EspectadoresEnArena1 As Byte
Public EspectadoresEnArena2 As Byte
Public EspectadoresEnArena3 As Byte
Public EspectadoresEnArena4 As Byte
Public NombreDueleando(1 To 8) As String

'Canjes
Public PremiosList() As tPremiosCanjes
Public Type tPremiosCanjes
    ObjName As String
    ObjIndexP As Integer
    ObjRequiere As Integer
    ObjDescripcion As String
    ObjMaxAt As Byte
    ObjMinAt As Byte
    ObjMindef As Byte
    ObjMaxdef As Byte
    ObjMinAtMag As Byte
    ObjMaxAtMag As Byte
    ObjMinDefMag As Byte
    ObjMaxDefMag As Byte
End Type

Public PuntosPremios As Long
Public puntixasd As Integer
Public UserInfo As String

Public opcion(1 To 5) As Boolean, Opciones(1 To 5) As String, HayEncuesta As Boolean, Encuesta As String, Votos(1 To 5) As Integer, LvlEncuesta As Byte

'#Fer - Subastas.
Public Hay_Subasta As Boolean
Public itemsubasta As Byte
Public orosubasta As Long
Public cantsubasta As Integer
Public OroOfrecido As Long
Public OroOfrecidox As Long
Public UltimoOfertador As String
Public MinutinSubasta As Byte
Public Subastador As String
Public objetosubastado As Obj

'Torneos
Public TiroCuentaDM As Boolean
Public CuentaTorneo As Integer
Public UsuariosEnTorneo As Integer
Public Hay_Torneo As Boolean
Public TModalidad As String
Public TNivelMinimo As Byte
Public CParticipantes As Byte
Public CuentaAutomatico As Integer

'Portal
Public MapaPortal As Byte
Public YPortal As Byte
Public XPortal As Byte

'Rey Ancalagon
Public MinutosRey As Byte
Public ReyON As Byte
Public MurioDragon As Byte
Public GuardiasRey As Byte

Public UltimaParty As Long
Public TiempoDuelo(1 To 4) As Byte
Public PremiosCastis As Byte
Public VerPrivados As Boolean
Public VerClanes As Boolean

Public InvocoBicho As Boolean
Public SegundosInvo As Byte
Public Const mapainvo = 93
Public Const mapainvoX1 = 29
Public Const mapainvoY1 = 26
Public Const mapainvoX2 = 26
Public Const mapainvoY2 = 29
Public Const mapainvoX3 = 32
Public Const mapainvoY3 = 29
Public Const mapainvoX4 = 29
Public Const mapainvoY4 = 32

' TODO: Y ESTO ? LO CONOCE GD ?
Public Nombre1 As String
Public Nombre2 As String
Public Const FX_TELEPORT_INDEX As Integer = 1

Public CvcFunciona As Boolean
Public MinutosPoder As Integer

Public PasoHD As Boolean
Public HDSerialIndex As String

'Global
Public ChatGlobal As Boolean

Public PJEnCuenta As String
Public PJEnCuentaB As String
Public totalAccounts As Long
Public totalPjs As Long

''
' Modulo de declaraciones. Aca hay de todo.
'

Public MixedKey As Long
Public ServerIp As String
Public CrcSubKey As String

Public CuentaRegresiva As Long
Public MapaCont As Byte

Public CastilloNorte As String
Public CastilloSur As String
Public CastilloEste As String
Public CastilloOeste As String
Public Fortaleza As String

Public GranPoder As Integer

Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14
Public Const FXAPUÑALAR = 54
Public Const FXE1 = 78
Public Const FXE2 = 79
Public Const FXE3 = 80
Public Const FXE4 = 81
Public Const FXE5 = 82
Public Const FXE6 = 83
Public Const FXE7 = 84
Public Const FXE8 = 85
Public Const FXE9 = 86
Public Const FXE10 = 87
Public Const FXE11 = 88
Public Const FXE12 = 89
Public Const FXE13 = 90
Public Const FXE14 = 91
Public Const FXE15 = 92
Public Const FXE16 = 93
Public Const FXE17 = 94
Public Const FXE18 = 95
Public Const FXE19 = 96
Public Const FXE20 = 97

Public Const iFragataFantasmal = 87

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Enum PlayerType
    User = 0
    UserSupport = 1
    EventManager = 2
    TournamentManager = 3
    Coordination = 4
    Development = 5
    Administrator = 6
End Enum

Public Const LimiteNewbie As Byte = 9

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

'Barrin 3/10/03
Public Const TIEMPO_INICIOMEDITAR As Integer = 1000

Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 1053
Public Const LAUDMAGICO As Integer = 696

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 43
    FXMEDITARCIUDA = 44
    FXMEDITARCRIMI = 42
    FXMEDITARTRANSFO = 16
    FXNOBLE = 45
    FXNUEVATPNEUTRAL = 103
    FXNUEVATPALIANZA = 104
    FXNUEVATPHORDA = 105
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

'TODO : Reemplazar por un enum
Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"
Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
    uOnlyUsuario = 5
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
    uTeleporta = 5
    uBurbuja = 6
    uInvocaMascota = 7
End Enum

Public Const DRAGON As Integer = 6
Public Const MAXUSERHECHIZOS As Byte = 20


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAXREP As Long = 99999999
Public Const MAXORO As Long = 999999999
Public Const MAXEXP As Long = 999999999

Public Const MAXATRIBUTOS As Byte = 35
Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const Leña As Integer = 58


Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000
Public Const MAXCHARSX As Integer = 550

Public Const HACHA_LEÑADOR As Integer = 127
Public Const PIQUETE_MINERO As Integer = 187

Public Const DAGA As Integer = 15
Public Const FOGATA_APAG As Integer = 136
Public Const FOGATA As Integer = 63
Public Const ORO_MINA As Integer = 194
Public Const PLATA_MINA As Integer = 193
Public Const HIERRO_MINA As Integer = 192
Public Const MARTILLO_HERRERO As Integer = 389
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const ObjArboles As Integer = 4
Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Integer = 138

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Timbero = 7
    Guardiascaos = 8
    Renunciar = 9
    ReyCastillo = 10
    Quest = 11
    Viajero = 12
    Ciudadania = 13
    Inscribe = 14
    ShowCasas = 15
    Arenas = 16
    QuestNoble = 17
    NpcDioses = 18
    cirujano = 19
    QuintaJera = 21
    BoveClan = 22
    Correos = 23
    EntregaCajas = 24
End Enum

Public Const MIN_APUÑALAR As Byte = 10

Public Const MapCastilloN = 33
Public Const MapCastilloS = 31
Public Const MapCastilloE = 34
Public Const MapCastilloO = 32

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 22

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 18

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 5


''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = 777

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuertoN As Integer = 8
Public Const iCabezaMuertoN As Integer = 500
Public Const iCuerpoMuertoH As Integer = 205
Public Const iCabezaMuertoH As Integer = 511
Public Const iCuerpoMuertoA As Integer = 206
Public Const iCabezaMuertoA As Integer = 512


Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    Armas = 5
    Meditar = 6
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
    Liderazgo = 17
    Domar = 18
    Proyectiles = 19
    Wresterling = 20
    Navegacion = 21
    DefensaMagica = 22
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_TRANSF As Byte = 150
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 6

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42
Public Const SND_BEBER As Byte = 46

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 35

' CATEGORIAS PRINCIPALES
Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otcASCO = 17
    otHerramientas = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otAriete = 36
    otGemaOctarina = 37
    otMapaTesoro = 38
    otGemaSagrada = 39
    otGemaNegra = 40
    otCristales = 41
    otContenedor = 42
    otCajasDios = 43
    otFragmento = 45
    otPocionResu = 46
    otCofreAzar = 47
    otCualquiera = 1000
End Enum

'Texto
Public Const FONTTYPE_GANAR As String = "~240~240~50~1~0"
Public Const FONTTYPE_CONSOLA As String = "~0~128~128~0~0"
Public Const FONTTYPE_GLOBAL As String = "~0~128~128~0~0"
Public Const FONTTYPE_UDP As String = "~255~0~0~0~0"
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_TORNEIN As String = "~225~249~158~1~0"
Public Const FONTTYPE_ORO As String = "~225~222~119~1~0"
Public Const FONTTYPE_OROX As String = "~225~222~119~1~0"
Public Const FONTTYPE_TSUBASTA As String = "~255~255~255~0~0"
Public Const FONTTYPE_TDSUBASTA As String = "~255~255~255~1~0"
Public Const FONTTYPE_SUBASTA As String = "~48~128~255~0~0"
Public Const FONTTYPE_NPCS As String = "~86~87~89~0~0"
Public Const FONTTYPE_ATNPC As String = "~114~0~4~1~0"
Public Const FONTTYPE_GANAORO As String = "~145~9~179~0~0"
Public Const FONTTYPE_DIOSES As String = "~100~0~255~0~0"
Public Const FONTTYPE_DIOSESI As String = "~100~0~255~0~1"
Public Const FONTTYPE_DIOSESN As String = "~100~0~255~1~0"
Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_INFOITALIC As String = "~65~190~156~0~1"
Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~255~255~255~0~1"
Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
Public Const FONTTYPE_FORTA As String = "~177~153~57~1~1"
Public Const FONTTYPE_CASTI As String = "~255~255~100~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~0~64~128~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~140~0~0~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~0~64~128~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~140~0~0~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"
Public Const FONTTYPE_ADVERTENCIAS As String = "~128~0~0~1~1"
Public Const FONTTYPE_AMARILLON As String = "~255~255~0~1~0"
Public Const FONTTYPE_EXPEN As String = "~236~186~107~1~0"
Public Const FONTTYPE_GRISN As String = "~130~130~130~1~0"
Public Const FONTTYPE_DAREXP As String = "~255~255~0~1~0"
Public Const FONTTYPE_ROJO As String = "~255~0~0~0~0"
Public Const FONTTYPE_GLOBALUSUARIO As String = "~173~170~255~0~0"
Public Const FONTTYPE_GLOBALNOBLE As String = "~255~255~0~0~0"
Public Const FONTTYPE_GLOBALGM As String = "~0~255~128~0~0"

'Colores en negrita
Public Const FONTTYPE_BORDON As String = "~128~0~0~1~0"
Public Const FONTTYPE_VERDEN As String = "~0~255~0~1~0"
Public Const FONTTYPE_OLIVE As String = "~107~142~35~1~0"
Public Const FONTTYPE_ROJON As String = "~255~0~0~1~0"
Public Const FONTTYPE_AZULN As String = "~0~0~255~1~0"
Public Const FONTTYPE_VIOLETAN As String = "~128~0~128~1~0"
Public Const FONTTYPE_CELESTEN As String = "~128~255~255~1~0"
Public Const FONTTYPE_DON As String = "~255~0~0~0~1"
Public Const FONTTYPE_AZULC As String = "~0~64~128~1~0"

'Colores en cursiva & negrita
Public Const FONTTYPE_BORDOCN As String = "~128~0~0~1~1"
Public Const FONTTYPE_VERDECN As String = "~0~255~0~1~1"
Public Const FONTTYPE_ROJOCN As String = "~255~0~0~1~1"
Public Const FONTTYPE_AZULCN As String = "~0~0~255~1~1"
Public Const FONTTYPE_VIOLETACN As String = "~128~0~128~1~1"
Public Const FONTTYPE_CELESTECN As String = "~128~255~255~1~1"
Public Const FONTTYPE_GRISCN As String = "~130~130~130~1~1"

'Colores en cursiva
Public Const FONTTYPE_BLANCOC As String = "~255~255~255~0~1"
Public Const FONTTYPE_BORDOC As String = "~128~0~0~0~1"
Public Const FONTTYPE_VERDEC As String = "~0~255~0~0~1"
Public Const FONTTYPE_ROJOC As String = "~255~0~0~0~1"
Public Const FONTTYPE_VIOLETAC As String = "~128~0~128~0~1"
Public Const FONTTYPE_CELESTEC As String = "~128~255~255~0~1"
Public Const FONTTYPE_GRISC As String = "~130~130~130~0~1"
Public Const FONTTYPE_VERDEL As String = "~0~185~0~1~0"



'Estadisticas
Public Const STAT_MAXELV As Byte = 70
Public Const STAT_MAXSTA As Integer = 30000
Public Const STAT_MAXMAN As Integer = 30000
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99


' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type tHechizo
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    ExclusivoClase As String
    ExclusivoClasedos As String
    ProhibidoClase As String
    
    Particle_Speed As Single
   Particle_Index  As Integer
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Resis As Byte
    
    Tipo As TipoHechizo
    
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    ActivaNobleza As Byte
    BacuNecesario As Byte
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    MaxDef1   As Integer
    MinDef1   As Integer
    RemueveInvisibilidadParcial As Byte
    CuartaJerarquia As Byte
    PortalMap As Byte
    PortalX As Byte
    PortalY As Byte
    Telepo As Byte
    
    Invoca As Byte
    NumNPC As Integer
    Cant As Integer
    
    Materializa As Byte
    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
    ProbTirar As Byte
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    ExObject(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
End Type

Public Type tPartyData
    pIndex As Integer
    RemXP As Double 'La exp. en el server se cuenta con Doubles
    TargetUser As Integer 'Para las invitaciones
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Public Type Char
    AuraA As Integer
    AuraW As Integer
    AuraE As Integer
    AuraR As Integer
    AuraC As Integer
    Account As String
    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
        
    FX As Integer
    loops As Integer
    
    Particula As Byte 'by juanjo
    
    Heading As eHeading
End Type

'Tipos de objetos
Public Type ObjData
    Name As String 'Nombre del obj
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    Apuñala As Byte
    Aura As Byte
    
    HechizoIndex As Integer
    DosManos As Byte
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    proyectil As Integer
    Municion As Integer
    
    Inmoviliza As Byte
    
    Crucial As Byte
    AntiLimpieza As Byte
    Intransferible As Byte
    ItemDios As Byte
    CristalesMax As Integer
    CristalesMin As Integer
    Dios As String
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    lvl As Byte
    
    DañoMagicoMin As Integer
    DañoMagicoMax As Integer
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    PuertaDoble As Byte
    Porton As Byte
    RejaForta As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    TipoCofre As Byte
    
    RazaEnana As Byte
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    Piedras As Byte
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    Texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As String
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
    
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type

Public Type BancoInventarioB
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type
'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

'Estadisticas de los usuarios
Public Type UserStats
    ParejasGanadas As Long
    ParejasPerdidas As Long
    DuelosGanados As Long
    DuelosPerdidos As Long
    MuertesUser As Long
    TrofOro As Byte
    TrofBronce As Byte
    TrofPlata As Byte
    MedOro As Byte
    TorneosParticipados As Integer
    MaximasRondas As Long
    PuntosTorneo As Long
    PuntosDonacion As Long
    TSPoints As Long
    Reputacione As Long
    GLD As Long 'Dinero
    Banco As Long
    MET As Integer
    
    MaxHP As Integer
    MinHP As Integer
    
    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    def As Integer
    Exp As Double
    ELV As Long
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    CriminalesMatados As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
End Type

'Duelos
Public dMap        As String
Public dUser       As String
Public dMoney      As Long
Public dIndex      As String
Public ArenaOcupada(1 To 4) As Boolean
'Duelos
    
Public Type UserMithStatus
EsStatus As Byte
EligioStatus As Byte
End Type

Public Type UserConsejos
PertAlCons As Byte
PertAlConsCaos As Byte
LiderConsejo As Byte
LiderConsejoCaos As Byte
End Type

Public HayPareja As Boolean

Public Desafio As Desafio

Public Type Desafio
Primero As Integer
Segundo As Integer
tPrimero As Integer
tSegundo As Integer
tTercero As Integer
tCuarto As Integer
End Type

Public Pareja As Pareja
 
Public Type Pareja
    Jugador(1 To 4) As Integer
End Type

Public Estrella As sRanking
 
Public Type sRanking
    TOPDuelos(1 To 3) As String
    TOPParejas(1 To 3) As String
    TOPFrags(1 To 3) As String
    TOPTorneos(1 To 3) As String
    TOPRondas(1 To 3) As String
    TOPEvents(1 To 3) As String
End Type

Public Type UserCofres
    'Dioses
    Item(1 To 4) As String
    Cant As String
End Type

'Flags
Public Type UserFlags
    TimeDeads As Byte
    CantDeads As Byte
    EnAram As Boolean
    AramRojo As Boolean
    AramAzul As Boolean
    AramDeads As Integer
    AramSeconds As Integer
    AntiAFK As Long
    TieneRanking As Boolean
    PosRanking As Byte
    NombreAmigo(1 To 20) As String
    AmigoDesadmitido(1 To 20) As Byte
    PuedeEntrarCVC As Boolean
    UltimoMatado As String
    NumCorreos As Byte
    NueCorreos(1 To 30) As String
    Correo(1 To 30) As String
    PuedeRetirarOro As Byte
    PuedeRetirarObj As Byte
    OroQueOferto As Long
    MandoDesafioA As Integer
    TieneDesafioDe As Integer
    TiempoOnlineHoy As Long
    TiempoParaCofres As Long
    Voto As Boolean
    VotoPorLaOpcion(1 To 5) As Integer
    EnGuerra As Byte
    JerarquiaDios As Byte
    SirvienteDeDios As String
    AlmasContenidas As Long
    AlmasOfrecidas As Long
    CuentaBancaria As String
    SolicitudDe As Integer
    MandoSolicitudA As Integer
    Pareja As String
    TorneoUsers As Byte
    Automatico As Boolean
    EspectadorArena1 As Byte
    EspectadorArena2 As Byte
    EspectadorArena3 As Byte
    EspectadorArena4 As Byte
    CaballerodelDragon As Byte
    targetBot As Byte
    DondeTiroMap As Integer
    DondeTiroX As Integer
    DondeTiroY As Integer
    TiroPortalL As Byte
    CvcsGanados As Integer
    GuerrasGanadas As Integer
    GuerrasPerdidas As Integer
    MVPMatados As Integer
    ApuestaOro As Long
    MascotinIndex As Integer
    InvocoMascota As Byte
    PartyIndex As Long
    PartySolicitud As Byte
    PartyLider As Byte
    TeniaElDon As Byte
    Stopped As Byte
    DefensaBurbu             As Integer
    IntervaloBurbu           As Integer
    SubeManaG As Byte
    SubeVidaG As Byte
    ActivoGema As Byte
    GemaActivada As String
    TimeGema As Byte
    ConsultaEnviada As Boolean
    NumeroConsulta As Long
    EnCvc As Boolean
    CvcBlue As Byte
    CvcRed As Byte
    CastiBlue As Byte
    CastiRed As Byte
    SeguroCVC As Boolean
    SuPareja As Integer
    EsperaPareja As Boolean
    tSuPareja As Integer
    tEsperaPareja As Boolean
    EnPareja As Boolean
    LeMandaronDuelo As Boolean
    UltimoEnMandarDuelo As String
    EnDuelo As Boolean
    DueliandoContra As String
    EnQueArena As Byte
    GranPoder As Byte
    Desenterrando As Byte
    QuestCompletadas As Integer
    EsPremium As Byte
    VencePremium As String
    EsNoble As Byte
    estado As Byte
    MuereQuest(1 To 3) As Long
    Questeando As Byte
    UserNumQuest As Byte
    TimeRevivir As Integer
    PJerarquia As Byte
    SJerarquia As Byte
    TJerarquia As Byte
    CJerarquia As Byte
    CJerarquiaC As Byte
    Desafio As Integer
    EnDesafio As Integer
    rondas As Integer
    llegolvl50 As Byte
    Llegolvlmax As Byte
    EleDeTierra As Byte
    EleDeFuego As Byte
    EleDeAgua As Byte
    DeseoRecibirMSJ As Byte
    Emoticons As Byte
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    ModoCombate As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    MapaAnterior As Byte
    XAnterior As Byte
    YAnterior As Byte
    NumTorneo As Byte
    EnTorneo As Byte
    MapaAnterior_dos As Byte
    XAnterior_dos As Byte
    YAnterior_dos As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    Navegando As Byte
    Montando As Byte
    Transformado As Byte
    Seguro As Boolean
    SeguroResu As Boolean
    SeguroClan As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    NpcInv As Integer
    
    Ban As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    
    Privilegios As PlayerType
    EsRolesMaster As Boolean
    
    LastCrimMatado As String
    LastCiudMatado As String
    LastNeutrMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    PertAlCons As Byte
    PertAlConsCaos As Byte
    
    Silenciado As Byte
    
    Mimetizado As Byte
    
    CentinelaOK As Boolean 'Centinela
End Type

Public Type UserCounters
    TiempoElemental As Byte
    InmoManopla As Byte
    TransportePremium As Byte
    TransporteCastillos(31 To 35) As Byte
    Seguimiento As Intervalos
    SegundosParaRevivir As Byte
    CreoTeleport As Boolean
    TimeTeleport As Integer
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Invisibilidad As Integer
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    Pasos As Integer
    ConquistandoCS As Byte
     TimeComandos As Byte
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    'Barrin 3/10/03
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    'Barrin
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
End Type

Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Double
    CiudadanosMatados As Double
    NeutralesMatados As Double
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
End Type

Public Type cFlagComer
    cObj(20)    As Obj
    cComercia   As Boolean
    cQuien      As Integer
    cOfrecio    As Boolean
    cRecivio    As Boolean
    cRespuesta  As Byte
    cOro As Long
End Type

Public Type cFlagCorreo
    cObj(20)    As Obj
    cComercia   As Boolean
    cQuien      As Integer
    cOfrecio    As Boolean
    cRecivio    As Boolean
    cRespuesta  As Byte
    cOro As Long
End Type

'Tipo de los Usuarios
Public Type User
    cComercio As cFlagComer
    cCorreo As cFlagCorreo
    EnCvc As Boolean
    ViejaPos As WorldPos
    NickMascota As String
    Name As String
    ID As Long
    clave2 As Long
    clave As String
    
    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    modName As String
    Password As String
    
    Bon1 As String
    Bon2 As String
    Bon3 As String
    
    UltimoLogeo As String
    PrimeraDenuncia As String
    UltimaDenuncia As String
    
    Char As Char 'Define la apariencia
    CharMimetizado As Char
    OrigChar As Char
    
    Desc As String ' Descripcion
    DescRM As String
    
    clase As String
    Raza As String
    Genero As String
    email As String
    Hogar As String
    
    InventorySlots As Byte
        
    Invent As Inventario
    
    Pos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    RDBuffer As String 'Buffer roto
    
    CommandsBuffer As New CColaArray
    ColaSalida As New Collection
    SockPuedoEnviar As Boolean
    
    '[KEVIN]
    BancoInvent As BancoInventario
    BancoInventB As BancoInventarioB
    '[/KEVIN]
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer
    
    Stats As UserStats
    flags As UserFlags
    Misiones As tMDiarias
    CofreDios As UserCofres
    ConsejoInfo As UserConsejos
    StatusMith As UserMithStatus
    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long
    
    Faccion As tFacciones
    
    PrevCheckSum As Long
    PacketNumber As Long
    RandKey As Long
    
    ip As String
    hd As String
    
    Accounted As String
    AccountedPass As String
    UserPremiumMap As Long

    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito
    
    AreasInfo As AreaInfo
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    UsuariosMatados As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    AfectaRelampago As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    
    LanzaFlecha As Byte
    '[KEVIN]
    'DeQuest As Byte
    
    'ExpDada As Long
    ExpCount As Long '[ALEJO]
    '[/KEVIN]
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    AtacaAPJ As Integer
    AtacaANPC As Integer
    AIAlineacion As e_Alineacion
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
End Type
' New type for holding the pathfinding info


Public Type npc
    Name As String
    Char As Char 'Define como se vera
    Desc As String
    DescExtra As String

    NPCtype As eNPCType
    Numero As Integer
    MVP As Byte

    level As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte

    Pos As WorldPos 'Posicion
    Orig As WorldPos

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Inflacion As Long

    GiveEXP As Long
    GiveGLD As Long
    GivePTS As Long
    GiveGLDMin As Long
    GiveGLDMax As Long
    
    Cristales As Byte
    CristalesPequesMin As Byte
    CristalesPequesMax As Byte
    CristalesMedianosMin As Byte
    CristalesMedianosMax As Byte
    CristalesGrandesMin As Byte
    CristalesGrandesMax As Byte
    CristalesEpicosMin As Byte
    CristalesEpicosMax As Byte
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    DueñoMascota As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    base_light(0 To 3) As Integer 'Indica si el tile tiene luz propia.
    light_value(0 To 3) As Long
    particle_group_index As Integer
    Blocked As Byte
    Graphic(1 To 4) As Integer
    userindex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    trigger As eTrigger
    BotIndex As Long
    Cofre As Integer
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    r As Byte
    g As Byte
    b As Byte
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public BackUp As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Const ENDL As String * 2 = vbCrLf
Public Const ENDC As String * 1 = vbNullChar
Public recordusuarios As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public PuedeCrearPersonajes As Integer
Public ServerSoloGMs As Integer

Public EnPausa As Boolean


'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist() As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To STAT_MAXELV) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public BanIps As New Collection
'*********************************************************

Public Tanaris As WorldPos
Public Prision As WorldPos
Public Libertad As WorldPos

Public ControllerIPs(1 To 600) As IPControl

Public Type IPControl
    NumIP As String
    CantConnect As Long
    CantPackets As Long
End Type

Public SonidosMapas As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum
