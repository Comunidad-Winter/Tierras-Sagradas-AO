Attribute VB_Name = "Mod_BOTS"
Option Explicit
 
'Defensa del bot jeje
Private Const IA_MINDEF  As Integer = 10
Private Const IA_MAXDEF  As Integer = 12
 
 Public Const MAX_BOTS   As Byte = 10
 
'Charindex reservado.
Private Const IA_CHAR    As Integer = (MAXCHARSX - MAX_BOTS)
 
'Datos del char
 
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
 
'ATENCION : Acá van los números de objetos!!!
 
Private Const IA_HEAD    As Integer = 107
Private Const IA_BODY    As Integer = 859
 
 
'ATENCION : Acá van los números de objetos!!!
 
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
 
'Cantidad de hechizos que lanza
 
Private Const IA_M_SPELL As Byte = 3
Private Const IA_NUMCHAT As Byte = 5
 
'Constantes de intervalos.
 
Private Const IA_SINT   As Integer = 800    'Intervalo entre hechizo-hechizo.
Private Const IA_SREMO  As Integer = 500    'Intervalo remo.
Private Const IA_MOVINT As Integer = 240    'Intervalo caminta.
Private Const IA_USEOBJ As Integer = 250    'Intervalo usar potas.
Private Const IA_HITINT As Integer = 200    'Intervalo para golpe
Private Const IA_PROINT As Integer = 700    'Intervalo de flecha
Private Const IA_TALKIN As Integer = 400   'Intervalo de hablAR :P
 
'Probabilidades de que te pegue
 
Private Const IA_CASTEO As Byte = 77
 
Private Const IA_PROBEV As Byte = 160
Private Const IA_PROBEX As Byte = 220
 
Private Const IA_SLOTS  As Byte = 20
 
Type ia_Interval
     SpellCount         As Byte         'Intervalo para tirar hechizos.
     UseItemCount       As Byte         'Intervalo para usar pociones.
     MoveCharCount      As Byte         'Intervalo para mover el char.
     ParalizisCount     As Byte         'Intervalo para removerse.
     HitCount           As Byte         'Intervalo para pegar golpesito.
     ArrowCount         As Byte         'Intervalo para flechas
     ChatCount          As Byte         'INtervalo para hablar XD
End Type
 
Type ia_Spells
     DamageMin          As Byte         'Minimo daño que hace.
     DamageMax          As Byte         'Maximo daño que hace.
     spellIndex         As Byte         'Lo usamos para el fx.
End Type
 
Enum eIASupportActions
     SRemover = 1                       'Remueve.
     SCurar = 2                         'Cura.
End Enum
 
Enum eIAClase
     Clerigo = 1                        'Bot Clero
     Mago = 2                           'Bot Mago
     Cazador = 3                        'Bot kza
End Enum
 
Enum eIAactions
     ePegar = 1                          'accion pegar.
     eMagia = 2                          'atacar con hechizo
End Enum
 
Enum eIAMoviments
     SeguirVictima = 1                   'Si seguia la victima
     MoverRandom = 2                     'Random moviment :P
End Enum
 
Type BOT
     GrupoID            As Integer
     EsCriminal         As Boolean
     Pos                As WorldPos     'Posicion en el mundo.
     maxVida            As Integer      'Maxima vida.
     Vida               As Integer      'Vida del bot.
     clase              As eIAClase     'Clases de bot.
     Tag                As String       'Tag del bot.
     Mana               As Integer      'Mana del bot.
     maxMana            As Integer      'Maxima mana
     Char               As Char         'Apariencia.
     AreasInfo As AreaInfo
     Invocado           As Boolean      'Si existe.
     Paralizado         As Boolean      'Si está inmo.
     Intervalos         As ia_Interval  'Intervalos de acciones.
     Viajante           As Boolean      'Bot Viajante :P
     ViajanteUser       As Integer      'Usuario que atacó al viajante.
     UltimaAccion       As eIAactions   'ULTIMA ACCION/ATAQUE.
     UltimoMovimiento   As eIAMoviments 'ULTIMO MOVIMIENTO
     Navegando          As Boolean      'Navegando?
     ViajanteAntes      As WorldPos     'Pos cuando un viajante ataca un usuario.
     Inv(1 To IA_SLOTS) As Obj          'Inventario del bot.
     UltimaIdaObjeto    As Boolean      'Ultimo movimiento fue buscar objs?
End Type
 
Public ia_Bot(1 To MAX_BOTS)           As BOT
Public ia_spell(1 To IA_M_SPELL)       As ia_Spells
Public ia_Chats(1 To IA_NUMCHAT)      As String
 
'Cantidad de bots invocados.
Public NumInvocados                    As Byte
 
Function ia_CascoByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve el casco/gorro según la clase del bot
 
Select Case ia_Bot(BotIndex).clase
 
       Case eIAClase.Clerigo        'Bot clero
            ia_CascoByClase = 131   'Completo : P
       
       Case eIAClase.Mago           'Bot mago.
            ia_CascoByClase = 868   'Vara
           
       Case eIAClase.Cazador        'Bot kza
            ia_CascoByClase = 405   'de plata
       
End Select
 
End Function
 
Function ia_EscudoByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve el escudo según la clase del bot
 
Select Case ia_Bot(BotIndex).clase
 
       Case eIAClase.Clerigo        'Bot clero
            ia_EscudoByClase = 130  'De plata.
       
       Case eIAClase.Mago           'Bot mago.
            ia_EscudoByClase = -1   'Nada
           
       Case eIAClase.Cazador        'bot kaza
            ia_EscudoByClase = 404  'escudo d tortu
       
End Select
 
End Function
 
Function ia_ArmaByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve el arma según la clase del bot
 
Select Case ia_Bot(BotIndex).clase
 
       Case eIAClase.Clerigo        'Bot clero
            ia_ArmaByClase = 129    'Dos filos : P
       
       Case eIAClase.Mago           'Bot mago.
            ia_ArmaByClase = 945    'Vara
           
       Case eIAClase.Cazador        'bot cazador
            ia_ArmaByClase = 665    'arko de kza
       
End Select
 
End Function
 
Function ia_VidaByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve la vida según la clase.
 
Select Case ia_Bot(BotIndex).clase
       Case eIAClase.Clerigo        '<Clerigo.
            'Vida random. (de clerigo 41)
            ia_VidaByClase = 21 + (RandomNumber(8, 10) * 49)
       
       Case eIAClase.Mago           '<Mago
            'Vida random (de mago 39)
            ia_VidaByClase = RandomNumber(360, 380)
           
       Case eIAClase.Cazador        '<Kza
            'Vida random de cazador humano 42
            ia_VidaByClase = 21 + (RandomNumber(8, 11) * 49)
           
End Select
 
End Function
 
Function ia_ManaByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve maná según la clase.
 
Select Case ia_Bot(BotIndex).clase
       Case eIAClase.Clerigo        '<Clerigo.
            'Mana de clero 41 : P
            ia_ManaByClase = 1480
       
       Case eIAClase.Mago           '<Mago
            'Mana de mago 39 : P
            ia_ManaByClase = 1954
           
       Case eIAClase.Cazador        'caza sin mana
            ia_ManaByClase = 0
           
End Select
 
End Function
 
Function ia_CalcularGolpe(ByVal VictimIndex As Integer) As Integer
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Calcula el golpe (daño) q hace el bot al user.
 
Dim ParteCuerpo     As Integer
Dim DañoAbsorvido   As Integer
 
ParteCuerpo = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
 
'Si pega en la cabeza.
If ParteCuerpo = PartesCuerpo.bCabeza Then
   'Si tiene casco baja el golpe
       If UserList(VictimIndex).Invent.CascoEqpObjIndex <> 0 Then
          DañoAbsorvido = RandomNumber(ObjData(UserList(VictimIndex).Invent.CascoEqpObjIndex).MinDef, ObjData(UserList(VictimIndex).Invent.CascoEqpObjIndex).MaxDef)
       End If
Else
    'Se fija por la armadura.
       If UserList(VictimIndex).Invent.ArmourEqpObjIndex <> 0 Then
          DañoAbsorvido = RandomNumber(ObjData(UserList(VictimIndex).Invent.ArmourEqpObjIndex).MinDef, ObjData(UserList(VictimIndex).Invent.ArmourEqpObjIndex).MaxDef)
       End If
End If
       
'DEVUELVE.
ia_CalcularGolpe = (RandomNumber(150, 180) - DañoAbsorvido)
       
End Function
 
Function ia_AciertaGolpe(ByVal VictimIndex As Integer) As Boolean
' @note         :  Evasión del usuario, esto le faltan unos retoques.
 
Dim tempEvasion     As Long
Dim tempEvasionEsc  As Long
Dim tempResultado   As Long
 
'Evasión del usuario.
tempEvasion = PoderEvasion(VictimIndex)
 
'Evasión del usuario con escudos.
'Tiene escudo?
If UserList(VictimIndex).Invent.EscudoEqpObjIndex <> 0 Then
    tempEvasionEsc = PoderEvasionEscudo(VictimIndex)
    tempEvasionEsc = tempEvasion + tempEvasionEsc
Else
    tempEvasionEsc = 0
End If
 
'Acierta?
tempResultado = MaximoInt(10, MinimoInt(90, 50 + (IA_PROBEX - tempEvasion) * 0.4))
 
'Random.
ia_AciertaGolpe = (RandomNumber(1, 100) <= tempResultado)
 
End Function
 
Function ia_PuedeMeele(ByRef PosBot As WorldPos, ByRef PosVictim As WorldPos, ByRef NewHeading As eHeading) As Boolean
' @note         :  Se fija si está al lado, y guarda el heading.
 
With PosVictim
   
    'Mirando hacia la derecha lo tiene ?
    If PosBot.X + 1 = .X Then
       ia_PuedeMeele = (.Y = PosBot.Y)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.EAST
       End If
       
       Exit Function
    End If
   
    'mirando hacia izq?
    If PosBot.X - 1 = .X Then
       ia_PuedeMeele = (.Y = PosBot.Y)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.WEST
       End If
       
       Exit Function
    End If
   
    'mirando arriba
    If PosBot.Y - 1 = .Y Then
       ia_PuedeMeele = (.X = PosBot.X)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.NORTH
       End If
       
       Exit Function
    End If
   
    'Abajo.
    If PosBot.Y + 1 = .Y Then
       ia_PuedeMeele = (PosBot.X = .X)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.SOUTH
       End If
       
       Exit Function
    End If
   
End With
 
End Function
 
Sub ia_CreateChar(ByVal ProximoBot As Byte)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Crea el char.
 
Dim PackageToSend   As String
 
With ia_Bot(ProximoBot).Char
 
    .Body = ObjData(IA_BODY).Ropaje
    .Head = IA_HEAD
   
    'Siempre tienen arma.
    .WeaponAnim = ObjData(ia_ArmaByClase(ProximoBot)).WeaponAnim
   
    'Escudo no, me fijo si tienen..
    If ia_EscudoByClase(ProximoBot) <> -1 Then
        .ShieldAnim = ObjData(ia_EscudoByClase(ProximoBot)).ShieldAnim
    End If
   
    'Casco si..
    .CascoAnim = ObjData(ia_CascoByClase(ProximoBot)).CascoAnim
   
    'Precalculado : P
    .CharIndex = IA_CHAR + ProximoBot
   
    'Preparo el paquete de datos.
   
            Dim tmp_Color As Byte
           
            If ia_Bot(ProximoBot).EsCriminal Then
               tmp_Color = 1
            Else
               tmp_Color = 3
            End If
   
    Call SendData(SendTarget.ToMap, ia_Bot(ProximoBot).Char.CharIndex, ia_Bot(ProximoBot).Pos.Map, "CC" & .Body & "," & .Head & "," & eHeading.SOUTH & "," & .CharIndex & "," & ia_Bot(ProximoBot).Pos.X & "," & ia_Bot(ProximoBot).Pos.Y & "," & .WeaponAnim & "," & .ShieldAnim & " ," & 0 & "," & 0 & "," & .CascoAnim & "," & ia_Bot(ProximoBot).Tag & "," & tmp_Color & "," & 0)

   
End With
 
End Sub
Public Function ia_Spawn(ByVal iaClase As eIAClase, ByVal mapa, ByVal X, ByVal Y, ByRef BotTag As String, ByVal Viajante As Boolean, ByVal esPk As Boolean, ByVal g_ID As Integer) As Integer
 
Dim ProximoBot  As Byte
Dim PackageSend As String
 
ProximoBot = IA_GetNextSlot
Debug.Print ProximoBot
 
If Not ProximoBot <> 0 Then Exit Function
 
With ia_Bot(ProximoBot)
   
    .Invocado = True
   
    .clase = iaClase
   
    .GrupoID = g_ID
   
    .Mana = ia_ManaByClase(ProximoBot)
    .Vida = ia_VidaByClase(ProximoBot)
    .maxMana = .Mana
    .maxVida = .Vida
   
    .EsCriminal = esPk
   
    'Si es "viajante"..
    .Viajante = Viajante
   
    .Tag = BotTag
   
    .Paralizado = False
   
    'Seteo la posición.
    .Pos.Map = mapa
    .Pos.X = X
    .Pos.Y = Y
   
    'Creo el char.
    ia_CreateChar ProximoBot
   
    'Primer action ! : D
    ia_Action ProximoBot
   
    'PackageSend = PrepareMessageChatOverHead("VeNGan PutOs xD!", .Char.CharIndex, vbCyan)
    'ia_SendToBotArea ProximoBot, PackageSend
   
    .Intervalos.SpellCount = 100
   
    NumInvocados = NumInvocados + 1
   
    MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = ProximoBot
   
    'devuelvo el id del bot
    ia_Spawn = ProximoBot
   
End With
 
End Function
 
Public Sub ia_Spells()
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
'Un poco hardcodeado pero bueno :D
 
'Hechizo 1 : descarga.
ia_spell(1).DamageMax = 120
ia_spell(1).DamageMax = 177
ia_spell(1).spellIndex = 23
 
'Hechizo 2 : apoca
 
ia_spell(2).DamageMin = 190
ia_spell(2).DamageMax = 220
ia_spell(2).spellIndex = 25
 
'Paralizar.
ia_spell(3).DamageMax = 0
ia_spell(3).DamageMin = 0
ia_spell(3).spellIndex = 9
 
ia_Chats(1) = "LKA JÑKKLJA KÑL JAÑLKÑLKAJ LK AJ"
ia_Chats(2) = "NEGRO"
ia_Chats(3) = "CHAU CHE"
ia_Chats(4) = "NANANANA TE ISE PAPIYAA"
ia_Chats(5) = "HIJO DE PUTA"
 
End Sub
 
Sub ia_RandomMoveChar(ByVal BotIndex As Byte, ByVal siguiendoIndex As Integer, ByRef HError As Boolean)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
With ia_Bot(BotIndex)
 
    Dim nRandom     As Byte
   
    '25% De probabilidades de moverse a
    'cualquiera de las cuatro direcciones.
   
    nRandom = RandomNumber(1, 4)
   
    Select Case nRandom
   
           Case 1
           
                If ia_LegalPos(.Pos.X + 1, .Pos.Y, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
               
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                .Pos.X = .Pos.X + 1
                Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
           Case 2
           
                If ia_LegalPos(.Pos.X - 1, .Pos.Y, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                .Pos.X = .Pos.X - 1
                Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
           Case 3
           
                If ia_LegalPos(.Pos.X, .Pos.Y + 1, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                .Pos.Y = .Pos.Y + 1
                Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
           Case 4
           
                If ia_LegalPos(.Pos.X, .Pos.Y - 1, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
               
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                .Pos.Y = .Pos.Y - 1
                Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
   
    End Select
 
End With
 
End Sub
 
Sub ia_CargarRutas(ByRef MAPFILE As String, ByVal MapIndex As Integer)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @modificated  :  Carga las rutas de un mapa : D
 
Dim loopX   As Long
Dim LoopY   As Long
Dim tmpVal  As eHeading
 
For loopX = 1 To 100
    For LoopY = 1 To 100
       
        tmpVal = val(GetVar(MAPFILE, CStr(loopX) & "," & CStr(LoopY), "Direccion"))
       
        If tmpVal <> 0 Then
          ' MapData(MapIndex, loopX, loopY).Rutas(1) = tmpVal
        End If
       
    Next LoopY
Next loopX
 
End Sub
 
Function ia_LegalPos(ByVal X As Byte, ByVal Y As Byte, ByVal BotIndex As Byte, Optional ByVal siguiendoUser As Integer = 0) As Boolean
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @modificated  :  Esta función ya no trabaja con la pos del npc si no que ahora usa los parámetros.
 
ia_LegalPos = False
 
With MapData(ia_Bot(BotIndex).Pos.Map, X, Y)
 
     '¿Es un mapa valido?
    If (ia_Bot(BotIndex).Pos.Map <= 0 Or ia_Bot(BotIndex).Pos.Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then Exit Function
 
     'Tile bloqueado?
     If .Blocked <> 0 Then Exit Function
   
     'Hay un usuario?
     If .userindex > 0 Then
        'Si no es un adminInvisible entonces nos vamos.
        If UserList(.userindex).flags.AdminInvisible <> 1 Then Exit Function
    End If
 
    'Hay un NPC?
    If .NpcIndex <> 0 Then Exit Function
     
    'Hay un bot?
    If .BotIndex <> 0 Then Exit Function
   
    'Siguiendo Index?
    If siguiendoUser <> 0 Then
        'Válido para evitar el rango Y pero no su eje X.
        If Abs(Y - UserList(siguiendoUser).Pos.Y) > RANGO_VISION_Y Then Exit Function
   
        If Abs(X - UserList(siguiendoUser).Pos.X) > RANGO_VISION_X Then Exit Function
    End If
   
     ia_LegalPos = True
   
End With
 
End Function
 
Sub ia_SearchPath(ByVal BotIndex As Byte, ByRef tPos As WorldPos, ByRef findHeading As eHeading)
 
' @designer     :  maTih.-
' @date         :  2012/03/13
' @                Buscá una ruta y devuelve un puntero con el heading.
 
findHeading = FindDirection(ia_Bot(BotIndex).Pos, tPos)
 
End Sub
 
Sub ia_MoveToHeading(ByVal BotIndex As Byte, ByVal toHeading As eHeading, ByRef FoundErr As Boolean)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Mueve el char del npc hacia una posición.
 
FoundErr = True
 
Select Case toHeading
 
       Case eHeading.NORTH  '<Move norte.
            'No legal pos.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y - 1, BotIndex) Then Exit Sub
           
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.Map, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
            'Set la nueva posición
            ia_Bot(BotIndex).Pos.Y = ia_Bot(BotIndex).Pos.Y - 1
            Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
       Case eHeading.EAST   '<Move este.
            'Si hay posición inválida no se peude mover.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X + 1, ia_Bot(BotIndex).Pos.Y, BotIndex) Then Exit Sub
           
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.Map, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
           
            'Set la nueva posición
            ia_Bot(BotIndex).Pos.X = ia_Bot(BotIndex).Pos.X + 1
            Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
       Case eHeading.SOUTH  '<Move sur.
            'Si hay posición inválida no se peude mover.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y + 1, BotIndex) Then Exit Sub
           
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.Map, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
           
            'Set la nueva posición
            ia_Bot(BotIndex).Pos.Y = ia_Bot(BotIndex).Pos.Y + 1
            Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
       Case eHeading.WEST   '<Move oeste.
            'Si hay posición inválida no se peude mover.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X - 1, ia_Bot(BotIndex).Pos.Y, BotIndex) Then Exit Sub
           
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.Map, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
           
            'Set la nueva posición
            ia_Bot(BotIndex).Pos.X = ia_Bot(BotIndex).Pos.X - 1
            Call SendToUserAreaButindexBOT(BotIndex, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
           
End Select
 
FoundErr = False
 
End Sub
 
 
Sub ia_MoveViajante(ByVal BotIndex As Byte, ByVal Direccion As eHeading)
' @note         :  Move el viajante hacia una posición
 
Dim HabiaAgua As Boolean
 
With ia_Bot(BotIndex)
 
     'Hacia donde se mueve..
     Select Case Direccion
           
            Case eHeading.NORTH     'Norte.
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.Y = .Pos.Y - 1
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
                 
            Case eHeading.EAST      'Este.
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.X = .Pos.X + 1
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
           
            Case eHeading.SOUTH     'Sur.
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.Y = .Pos.Y + 1
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
                 
            Case eHeading.WEST      'Oeste.
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.X = .Pos.X - 1
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
     End Select
     
     HabiaAgua = HayAgua(.Pos.Map, .Pos.X, .Pos.Y)
     
     If HabiaAgua Then
        'Si hay agua cambio el cuerpo.
        Call ia_MoveToHeading(.Char.CharIndex, Direccion, False)
        .Navegando = True
     Else
        'No habia agua, y... estaba navegando?
        If .Navegando Then
           'cambio el body y demas.
           Call ia_MoveToHeading(.Char.CharIndex, Direccion, False)
           .Navegando = False
        End If
    End If
   
     'Actualizamso
     Call ModAreas.SendToAreaByPos(.Pos.Map, .Pos.X, .Pos.Y, "+" & .Char.CharIndex & "," & .Pos.X & "," & .Pos.Y)
       
End With
 
End Sub
 
Function ia_HeadingToMolestNpc(ByVal NpcIndex As Integer) As eHeading
' @note         :  Devuelve un heading para un npc que está molestando el paso.
 
Dim nPos    As WorldPos
 
nPos = Npclist(NpcIndex).Pos
 
With MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
 
     'Pos legal hacia arriba?
     If LegalPosNPC(nPos.Map, nPos.X, nPos.Y - 1, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.NORTH
        End If
     End If
     
     'Pos legal hacia abajo?
     If LegalPosNPC(nPos.Map, nPos.X, nPos.Y + 1, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.SOUTH
        End If
     End If
     
     'Pos legal hacia la izquierda?
     If LegalPosNPC(nPos.Map, nPos.X - 1, nPos.Y, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.WEST
        End If
     End If
     
     'Pos legal hacia la derecha?
     If LegalPosNPC(nPos.Map, nPos.X + 1, nPos.Y, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.EAST
        End If
     End If
     
End With
 
End Function
 
Function ia_Objetos(ByVal BotIndex As Byte) As WorldPos
' @note         :  Busca objetos valiosos en el area.
 
Dim loopX   As Long
Dim LoopY   As Long
Dim BotPos  As WorldPos
 
BotPos = ia_Bot(BotIndex).Pos
 
'********************************
 
'borro esto ya que no libero esta parte : p
 
'********************************
 
ia_Objetos.Map = 0
 
End Function
 
Function ia_SlotInventario(ByVal BotIndex As Byte) As Byte
' @note         :  Busca un slot libre.
 
Dim loopX   As Long
 
For loopX = 1 To IA_SLOTS
    With ia_Bot(BotIndex).Inv(loopX)
         'No hay objeto.
         If Not .ObjIndex <> 0 Then
            ia_SlotInventario = CByte(loopX)
            Exit Function
         End If
    End With
Next loopX
 
ia_SlotInventario = 0
 
End Function
 
Sub ia_ActionViajante(ByVal BotIndex As Byte)
' @note         :  Acciones de los bots que viajan hacia mapas.
 
Dim RutaDir     As eHeading
Dim molestNpc   As Integer
Dim ObjetoPos   As WorldPos
 
With ia_Bot(BotIndex)
 
     'Está paralizado?
     If .Paralizado Then
        'Puede tirar hechizos.
        If .Intervalos.SpellCount = 0 Then
           'se remueve
           Call SendData(SendTarget.ToPCArea, .Char.CharIndex, 0, "N|" & vbCyan & "AN HOAX VORP" & CStr(.Char.CharIndex))
           .Paralizado = False
           .Intervalos.SpellCount = (IA_SINT / 30)
        End If
     End If
       
     'Se puede mover?
     If Not .Intervalos.MoveCharCount = 0 Then Exit Sub
       
     .Intervalos.MoveCharCount = (IA_MOVINT / 50)
     
     'Tiene una ruta?
     RutaDir = ia_HayRuta(.Pos)
   
     'Ve un objeto valioso?
     ObjetoPos = ia_Objetos(BotIndex)
     
     If ObjetoPos.Map <> 0 Then
        'Lo va a buscar, pero antes , setea su vieja pos.
        If Not .UltimaIdaObjeto Then
            .ViajanteAntes = .Pos
        End If
       
        ia_SearchPath BotIndex, ObjetoPos, RutaDir
        .UltimaIdaObjeto = True
     End If
     
     'No hay ruta?
     If Not RutaDir <> 0 Then
        'habia atacado un usuario? si es así volvemos a la pos.
        ia_SearchPath BotIndex, .ViajanteAntes, RutaDir
     End If
     
     If RutaDir <> 0 Then
       
        'Hacia donde mueve?
        Select Case RutaDir
               
               Case eHeading.NORTH      '<Mueve norte.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).NpcIndex
                   
                    #If Barcos <> 0 Then
                        If molestNpc <> 0 Then
                            Call SendData(SendTarget.ToPCArea, .Char.CharIndex, 0, "N|" & vbWhite & "¡Maldita criatura, obstruyes mi paso!" & CStr(.Char.CharIndex))
                            Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                        End If
                    #End If
                   
               Case eHeading.SOUTH      '<Mueve sur.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.Map, .Pos.X, .Pos.Y + 1).NpcIndex
                   
                    If molestNpc <> 0 Then
                        Call SendData(SendTarget.ToPCArea, .Char.CharIndex, 0, "N|" & vbWhite & "¡Maldita criatura, obstruyes mi paso!" & CStr(.Char.CharIndex))
                       'muevo el npc
                       Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
                       
               Case eHeading.EAST       '<Mueve este.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.Map, .Pos.X + 1, .Pos.Y).NpcIndex
                   
                    If molestNpc <> 0 Then
                       Call SendData(SendTarget.ToPCArea, .Char.CharIndex, 0, "N|" & vbWhite & "¡Maldita criatura, obstruyes mi paso!" & CStr(.Char.CharIndex))
                       'muevo el npc
                       Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
                   
               Case eHeading.WEST       '<Mueve oeste.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.Map, .Pos.X - 1, .Pos.Y).NpcIndex
                   
                    If molestNpc <> 0 Then
                       Call SendData(SendTarget.ToPCArea, .Char.CharIndex, 0, "N|" & vbWhite & "¡Maldita criatura, obstruyes mi paso!" & CStr(.Char.CharIndex))
                       'Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
        End Select
       
        'Move:p
        ia_MoveViajante BotIndex, RutaDir
        'Set el heading.
        .Char.Heading = RutaDir
     End If
     
 
     
     'Encontramos una salida? - translados.
     If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map <> 0 Then
        'Mapa válido?
        If MapaValido(MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map) Then
            'Asignamos nuevas posiciones, borramos el char anterior.
            Call EraseUserChar(SendTarget.ToMap, 0, 0, .Char.CharIndex)
            'Pos del npc.
            .Pos.Map = MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map
           
            'Por si no tiene heading.
            If Not .Char.Heading <> 0 Then .Char.Heading = eHeading.SOUTH
           
            'Nueva X?
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.X <> 0 Then
                .Pos.X = MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.X
            End If
           
            'Nueva Y?
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Y <> 0 Then
                .Pos.Y = MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Y
            End If
           
             MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
            'Creamos.
           
            Dim tmp_Color As Byte
           
            'preparo el color del nick
            If .EsCriminal Then
               tmp_Color = 1
            Else
               tmp_Color = 3
            End If
           
            Call SendData(SendTarget.ToMap, .Char.CharIndex, .Pos.Map, "CC" & .Char.Body & "," & .Char.Head & "," & .Char.Heading & "," & .Char.CharIndex & "," & .Pos.X & "," & .Pos.Y & "," & .Char.WeaponAnim & "," & .Char.ShieldAnim & "," & 0 & "," & 0 & "," & .Char.CascoAnim & "," & .Tag & "," & tmp_Color & "," & 0)
        End If
     End If
     
End With
 
End Sub
 
Function ia_HayRuta(ByRef InPos As WorldPos) As eHeading
' @note         :  Devuelve la dircción de la ruta en una pos.
 
With MapData(InPos.Map, InPos.X, InPos.Y)
     
     'ia_HayRuta = .Rutas(1)
     
End With
 
End Function
 
Sub ia_SupportOthers(ByVal BotIndex As Byte, ByRef Supported As Boolean)
' @note         :  Un bot supportea otro.
 
Dim botIndexToSupport   As Byte
Dim supportAction       As eIASupportActions
 
'Si no tiene intervalo..
If ia_Bot(BotIndex).Intervalos.SpellCount <> 0 Then Exit Sub
 
'Busca un bot a ayudar.
botIndexToSupport = ia_GetSupportBot(BotIndex, supportAction)
 
'No encontró, no supportea..
If Not botIndexToSupport <> 0 Then Supported = False: Exit Sub
 
'Que acción debe realizar?
Select Case supportAction
 
       Case eIASupportActions.SCurar        '<Cura un compañero
            'Lanza graves.
            'Crea fx.
            Call SendData(SendTarget.ToMap, ia_Bot(botIndexToSupport).Char.CharIndex, ia_Bot(botIndexToSupport).Pos.Map, "CFX" & ia_Bot(botIndexToSupport).Char.CharIndex & "," & Hechizos(5).FXgrh & "," & Hechizos(5).loops)
           
            'Cartel.
            Call SendData(SendTarget.ToPCArea, botIndexToSupport, ia_Bot(BotIndex).Pos.Map, "N|" & vbCyan & "°" & Hechizos(5).PalabrasMagicas & "°" & ia_Bot(botIndexToSupport).Char.CharIndex)
           
            'Suma un random de vida.
            ia_Bot(botIndexToSupport).Vida = ia_Bot(botIndexToSupport).maxVida + RandomNumber(55, 77)
           
            'PARA QUE NO PASE LA VIDA MAXIMA
            If ia_Bot(botIndexToSupport).Vida > ia_Bot(botIndexToSupport).maxVida Then ia_Bot(botIndexToSupport).Vida = ia_Bot(botIndexToSupport).maxVida
       
            Supported = True
       
      Case eIASupportActions.SRemover       '<Remueve paralizis.
            'Crea el fx, el remo no tiene fx.
            'ia_sendtobotarea botindextosupport
           
            'Paralizis count.
            If ia_Bot(botIndexToSupport).Intervalos.ParalizisCount > 6 Then Exit Sub
           
            'Cartel
            Call SendData(SendTarget.ToPCArea, botIndexToSupport, ia_Bot(BotIndex).Pos.Map, "N|" & vbCyan & "°" & Hechizos(10).PalabrasMagicas & "°" & ia_Bot(botIndexToSupport).Char.CharIndex)
           
            'Saca el flag
            ia_Bot(botIndexToSupport).Paralizado = False
           
            Supported = True
           
End Select
 
End Sub
 
Function ia_BotEnArea(ByVal BotIndex As Byte, ByVal otherBotIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
 
Dim BotIndexPos As WorldPos
 
BotIndexPos = ia_Bot(BotIndex).Pos
 
Dim loopX   As Long
Dim LoopY   As Long
 
For LoopY = BotIndexPos.Y - MinYBorder + 1 To BotIndexPos.Y + MinYBorder - 1
        For loopX = BotIndexPos.X - MinXBorder + 1 To BotIndexPos.X + MinXBorder - 1
            'hay un bot
            If MapData(BotIndexPos.Map, loopX, LoopY).BotIndex = otherBotIndex Then
                ia_BotEnArea = True
                Exit Function
            End If
       
        Next loopX
Next LoopY
 
ia_BotEnArea = False
 
End Function
 
Function ia_GetSupportBot(ByVal BotIndex As Byte, ByRef SAction As eIASupportActions) As Byte
' @note         :  Busca un bot a ayudar.
 
Dim loopX   As Long
 
For loopX = 1 To MAX_BOTS
   
    'Si no es mi BotIndex
    If loopX <> BotIndex Then
       
       'Está invocado?
       If ia_Bot(loopX).Invocado Then
          'Está en el area?
          If ia_BotEnArea(BotIndex, loopX) Then
             'Está paralizado/tiene poca vida?
             If ia_Bot(loopX).Vida <> ia_Bot(loopX).maxVida Or ia_Bot(loopX).Paralizado Then
                'Encontrado.
                ia_GetSupportBot = CByte(loopX)
                'Devuelve la acción.
                SAction = IIf(ia_Bot(loopX).Vida <> ia_Bot(loopX).maxVida, eIASupportActions.SCurar, eIASupportActions.SRemover)
                Exit Function
             End If
          End If
       End If
       
    End If
   
Next loopX
 
ia_GetSupportBot = 0
End Function
 
Sub ia_Action(ByVal BotIndex As Byte)
 
On Error GoTo Errhandler
' @note         :  Acciones de los bots.
 
Dim pIndex      As Integer
Dim sRandom     As Integer
Dim rMan        As Integer
Dim FoundErr    As Boolean
Dim moveHeading As eHeading
Dim AyudoBot    As Boolean
 
If EnPausa Then Exit Sub
 
With ia_Bot(BotIndex)
 
    'Es un bot viajante?
    If .Viajante Then
          'Mientras no esté contra ningún pibe
          If Not .ViajanteUser <> 0 Then
             ia_CheckInts BotIndex
             ia_ActionViajante BotIndex
             Exit Sub
          End If
    End If
   
    'si no lo ataco nadie  busca un target
    If (.ViajanteUser = 0) Then
        pIndex = ia_FindTarget(.Pos, .EsCriminal)
    Else
        pIndex = .ViajanteUser
    End If
   
    'No hay usuario.
    If pIndex <= 0 Then Exit Sub
 
    'Contadores de intervalo.
    ia_CheckInts BotIndex
   
    'EL bot boquea XD
    If Not .Intervalos.ChatCount <> 0 Then
       .Intervalos.ChatCount = (IA_TALKIN / 40)
       
       'Envia msj random
       Call SendData(SendTarget.ToMap, .Char.CharIndex, 0, "N|" & vbRed & "°" & ia_Chats(RandomNumber(1, 5)) & "°" & .Char.CharIndex)
       .Intervalos.SpellCount = (IA_SINT / 100)
    End If
   
    'Si se puede mover AND no está inmo se mueve al azar.
    If .Intervalos.MoveCharCount = 0 And .Paralizado = False Then
       
        'Tiene target?
        If pIndex <> 0 Then
           'busco un path.
           ia_SearchPath BotIndex, UserList(pIndex).Pos, moveHeading
        End If
       
        'Es clero?
        If Not .clase <> eIAClase.Clerigo Then
           'Si tiene la vida llena lo persigue.
           If .Vida = .maxVida Then
              ia_MoveToHeading BotIndex, moveHeading, FoundErr
           Else
            'Si no , se mueve al azar.
              ia_RandomMoveChar BotIndex, pIndex, FoundErr
           End If
         End If
                   
         'Es mago?
        If .clase = eIAClase.Mago Or .clase = eIAClase.Cazador Then
           'Si no tiene la vida llena se mueve al azar.
           If Not .Vida = .maxVida Then
              ia_RandomMoveChar BotIndex, pIndex, FoundErr
           Else
              'Tiene la vida llena, que fue el ultimo movimiento?
              'Siguio la victima?
              If .UltimoMovimiento = eIAMoviments.SeguirVictima Then
                 'Mueve random.
                 ia_RandomMoveChar BotIndex, pIndex, FoundErr
                 'Seteo.
                 .UltimoMovimiento = eIAMoviments.MoverRandom
              Else
                 'Se movió al azar, sigue su victima.
                 ia_MoveToHeading BotIndex, moveHeading, FoundErr
                 'Seteo el nuevo flag.
                 .UltimoMovimiento = eIAMoviments.SeguirVictima
             End If
        End If
       End If
       
       'se movio.
        If Not FoundErr Then
          'Se movió, guardo el BotIndex.
          MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
         
          'NEW--------
          'Checkeo si es una posición válida.
   
          'Actualizamos.
          Call SendToUserAreaButindex(BotIndex, "+" & .Char.CharIndex & "," & .Pos.X & "," & .Pos.Y)
         
          .Intervalos.MoveCharCount = (IA_MOVINT / 40)
        End If
       
    End If
   
   
    'STATS..
   
        'Prioriza la vida ante todo
       
        If .Vida < .maxVida Then
           
            'Checkeo el intervalo.
            If .Intervalos.UseItemCount > 0 Then Exit Sub
           
            'Recupera 20 cada 200 ms.
            .Vida = .Vida + 20
           
            If .Vida > .maxVida Then .Vida = .maxVida
           
            'Uso la poción, seteo el interval
            .Intervalos.UseItemCount = (IA_USEOBJ / 40)
           
            Exit Sub
        End If
       
        'Si tenia la vida llena usa azules.
       
        If .Mana < .maxMana Then
       
            'Checkeo el intervalo.
           
            If .Intervalos.UseItemCount = 0 Then
           
                Dim recuperoMana    As Long
               
                'Recupera un % de la mana.
                If .clase <> eIAClase.Mago Then
                    recuperoMana = Porcentaje(.maxMana, 5)
                Else
                    recuperoMana = Porcentaje(.maxMana, 3)
                End If
               
                'aumento el mana
                .Mana = .Mana + recuperoMana
           
                'controlo el limite
                If .Mana > .maxMana Then .Mana = .maxMana
           
            'seteo el int
            .Intervalos.UseItemCount = (IA_USEOBJ / 40)
 
            End If
           
            'Hacer una constante después, con esto hacemos un random
            'Para que azulee y combee a la ves.
            If RandomNumber(1, 4) < 4 Then Exit Sub
        End If
   
    'Bueno si está acá es por que tenia la vida y mana llenas.
     
    'Es cazador??
    If .clase = eIAClase.Cazador Then
       'Intervalo permite?
       If Not .Intervalos.ArrowCount = 0 Then Exit Sub
       'Kza manqea XD - 25% de prob fallar
       If RandomNumber(1, 100) > 65 Then Exit Sub
       'Probabilidad de evadir.
       If Not RandomNumber(1, 100) <= MaximoInt(10, MinimoInt(90, 50 + ((220 - PoderEvasion(pIndex)) * 0.4))) Then
          'Atacó y falló!!
          Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te lanzó un flechazo pero falló!~69~190~156")
          'setea intervalo
          .Intervalos.ArrowCount = (IA_PROINT / 25)
          Exit Sub
       End If
       
       Dim ArrowDamage  As Integer  '<DañoBase.
       Dim ArmourIndex  As Integer  '<ArmaduraObjIndex
       Dim HelmetIndex  As Integer  '<CascoObjIndex
       
       ArrowDamage = RandomNumber(185, 225)
       
       'Restamos si tiene armadura.
       ArmourIndex = UserList(pIndex).Invent.ArmourEqpObjIndex
       HelmetIndex = UserList(pIndex).Invent.CascoEqpObjIndex
       
       'Pega en cabeza?
       If RandomNumber(1, 6) = 6 Then
          'Absorve.
          If HelmetIndex <> 0 Then
             ArrowDamage = ArrowDamage - RandomNumber(ObjData(HelmetIndex).MinDef, ObjData(HelmetIndex).MaxDef)
          End If
       Else
          'Armadura absorce.
          If ArmourIndex <> 0 Then
             ArrowDamage = ArrowDamage - RandomNumber(ObjData(ArmourIndex).MinDef, ObjData(ArmourIndex).MaxDef)
          End If
       End If
       
       'crea fx.
       'SendData SendTarget.ToPCArea, pIndex, mod_DunkanProtocol.Send_CreateArrow(.Char.CharIndex, UserList(pIndex).Char.CharIndex, ObjData(553).GrhIndex)
       
       'crea daño
       'Call mod_DunkanGeneral.Enviar_DañoAUsuario(pIndex, ArrowDamage)
       
       'Sacude un flechazo.
       UserList(pIndex).Stats.MinHP = UserList(pIndex).Stats.MinHP - ArrowDamage
       
       Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te ha pegado un flechazo por " & ArrowDamage & ".~128~0~0~1")
       
       'Muere?
       If UserList(pIndex).Stats.MinHP <= 0 Then
          UserDie pIndex
          Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te ha matado!~255~255~0~1")
       End If
       
       'Intervalo
       .Intervalos.ArrowCount = (IA_PROINT / 20)
       
       'client update
       SendUserHP pIndex
       Exit Sub
    End If
   
    'Puede castear?
    'Si el usuario no tiene la vida llena ataca
    Dim tmpHP   As Long
   
    tmpHP = (UserList(pIndex).Stats.MinHP)
   
    'obtengo el % de vida del user
    tmpHP = (tmpHP * 100) / (UserList(pIndex).Stats.MaxHP)
   
    If .Intervalos.SpellCount = 0 Then
   
    'Es clérigo y puede pegar??
    If (.clase = eIAClase.Clerigo) And .Intervalos.HitCount = 0 And Not .UltimaAccion = eIAactions.ePegar Then
       'Está al alcance de la víctima para un gole meele?
       Dim newBotHeading   As eHeading
       
       If ia_PuedeMeele(.Pos, UserList(pIndex).Pos, newBotHeading) Then
            'Acierta el golpe?
            If ia_AciertaGolpe(pIndex) Then
               'Antes que nada cambiamos el heading, si es válido.
               If newBotHeading <> 0 And newBotHeading <> .Char.Heading Then
                  'ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_ChangeHeadingChar(.Char.CharIndex, newBotHeading)
               End If
               
               'Calcula el golpe
               Dim GolpeVal     As Integer
               GolpeVal = ia_CalcularGolpe(pIndex)
               
               'Resta hp.
               UserList(pIndex).Stats.MinHP = UserList(pIndex).Stats.MinHP - GolpeVal
               
               'crea el fx de la sangre.
               'SendData SendTarget.ToPCArea, pIndex, PrepareMessageCreateFX(UserList(pIndex).Char.CharIndex, FXSANGRE, 5)
               
               'Avisa.
               Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te ha pegado por " & CStr(GolpeVal) & ".~125~0~0~1")
               
               'Setea flag.
               .UltimaAccion = eIAactions.ePegar
               
               'Muere?
               If UserList(pIndex).Stats.MinHP <= 0 Then
                  Call UserDie(pIndex)
               End If
               
               'update hp.
               SendUserHP pIndex
               
               'Intervalo de golpe.
               .Intervalos.HitCount = (IA_HITINT / 40)
               'Intervalo de hechizo.
               .Intervalos.SpellCount = (IA_SINT / 40)
               'Intervalo de golpe+pociones.
               .Intervalos.UseItemCount = (IA_USEOBJ / 60)
               Exit Sub
            End If
        End If
    End If
   
       'Feo, aunque digamos que solo hace apoca desc remo
       'Así que va a andar bien.
       
       'Si la mana es < a 300 [gasto del remo] no hacemos nada.
       
       If .Mana < 300 Then Exit Sub
       
       'Si está paralizado AND el usuario no tiene poka vida prioriza removerse.
       
        If .Paralizado And tmpHP > 60 Then
           
            'Intervalo de remo :@
            If .Intervalos.ParalizisCount <> 0 Then Exit Sub
           
            'Palabras mágicas.
            Call SendData(SendTarget.ToPCArea, BotIndex, ia_Bot(BotIndex).Pos.Map, "N|" & vbCyan & "°" & Hechizos(10).PalabrasMagicas & "°" & ia_Bot(BotIndex).Char.CharIndex)
           
            .Paralizado = False
           
            'Agrego esto por que si no tirarle inmo era al pedo
            'Seguia caminando practicamente :PP
           
            .Intervalos.ParalizisCount = (IA_SREMO / 10)
           
            'Se removió entonces salimos del sub y seteamos el intervalo
           
            .Intervalos.SpellCount = (IA_SINT / 40)
           
            Exit Sub
           
        End If
       
        'No está paralizado entonces castea un hechizo random.
       
        'chances de pegar
        If RandomNumber(1, 100) > IA_CASTEO Then Exit Sub
       
        sRandom = RandomNumber(1, IA_M_SPELL)
       
        'Ayuda otros bots si es que hay
        If NumInvocados <> 1 Then
           ia_SupportOthers BotIndex, AyudoBot
           
           'ayudo ya al bot?
           If AyudoBot Then
              'SETEA INTERVALO
              .Intervalos.SpellCount = (IA_SINT / 40)
              Exit Sub
           End If
           
        End If
           
        'Si el usuario ya estaba paralizado AND el random es paralizar, entonces buscamos de nuevo
        If UserList(pIndex).flags.Paralizado = 1 And sRandom = 3 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
       
        'Si soy mago y el usuario es mago también no paraliza.
        If UCase$(UserList(pIndex).clase) = "MAGO" And .clase = eIAClase.Mago Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
       
        'Si el usuario tiene menos del 75% de vida juega al ataque.
       
        If tmpHP < 75 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
       
        'Si no llega con la mana del hechizo AND la del otro
        'tampoco entonces no hacemos nada
       
        If sRandom = 1 Then
           
            'Si no llega a la mana del spell 1 (descarga)
            'No hacemos nada ya que tampoco llega
            'al apocalipsis.
           
            rMan = Hechizos(ia_spell(1).spellIndex).ManaRequerido
           
            If rMan > .Mana Then Exit Sub
           
        ElseIf sRandom = 2 Then
       
            rMan = Hechizos(ia_spell(2).spellIndex).ManaRequerido
               
            'Pero si es spell 2 (apoca) AND llegamos
            'con la mana para descarga, entonces
            'Seteamos sRandom como 1 y casteamos
            'descarga.
           
            If rMan > .Mana Then
               
                'Modifico la formula y hago un random
                'Dado a que una ves que queda en -1000 de mana
                'Nunca más tira apoca y castea puras descargas.
               
                If .Mana > 460 And RandomNumber(1, 100) < 30 Then
                    sRandom = 1
                Else
                    Exit Sub
                End If
            End If
       End If
       
        rMan = Hechizos(ia_spell(sRandom).spellIndex).ManaRequerido
       
        'Descontamos la maná y seteamos el intervalo.
        .Mana = .Mana - rMan
        
        'Set última action.
        .UltimaAccion = eIAactions.eMagia
       
        .Intervalos.SpellCount = (IA_SINT / 20) 'Se chekea cada 40 ms.
       
        'Creamos el fx y le descontamos la vida al usuario.
        Call SendData(SendTarget.ToMap, BotIndex, ia_Bot(BotIndex).Pos.Map, "N|" & vbCyan & "°" & Hechizos(ia_spell(sRandom).spellIndex).PalabrasMagicas & "°" & ia_Bot(BotIndex).Char.CharIndex)
        Call SendData(SendTarget.ToMap, UserList(pIndex).Char.CharIndex, UserList(pIndex).Pos.Map, "CFX" & UserList(pIndex).Char.CharIndex & "," & Hechizos(ia_spell(sRandom).spellIndex).FXgrh & "," & Hechizos(ia_spell(sRandom).spellIndex).loops)
       
        'Paralizar?
        If sRandom = 3 Then
           'Paralizado : P
           UserList(pIndex).flags.Paralizado = 1
           UserList(pIndex).Counters.Paralisis = IntervaloParalizado
           Call SendData(SendTarget.toindex, pIndex, 0, "PARADOK")
           Call SendData(SendTarget.toindex, pIndex, 0, "PU" & UserList(pIndex).Pos.X & "," & UserList(pIndex).Pos.Y)
           
           Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te ha paralizado.~69~190~156")
        End If
       
        'Random damage :D
       
        sRandom = RandomNumber(ia_spell(sRandom).DamageMin, ia_spell(sRandom).DamageMax)
       
        'Al daño le restamos , si el usuario tiene, defensa mágica.
        If UserList(pIndex).Invent.HerramientaEqpObjIndex <> 0 Then
           sRandom = sRandom - RandomNumber(ObjData(UserList(pIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(pIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
       
        'NO numeros negativos.
        If sRandom < 0 Then sRandom = 0
       
        'Quitamos daño.
        UserList(pIndex).Stats.MinHP = UserList(pIndex).Stats.MinHP - sRandom
           
        If sRandom <> 0 Then
            'AVISO AL USUARIO DE ESTO
            Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te ha quitado " & sRandom & " puntos de vida.~128~0~0~1")
        End If
       
        'Check si muere.
        If UserList(pIndex).Stats.MinHP <= 0 Then
             UserDie pIndex
             
            'Era viajante y mató el usuario?, resteo el ui
             If Not pIndex <> .ViajanteUser Then
                .ViajanteUser = 0
             End If
             
             'aviso que murio.
             Call SendData(SendTarget.toindex, pIndex, 0, "N|" & .Tag & " te ha matado!~255~255~0~1")
        End If
       
        'Actualizamos el cliente.
       
        SendUserStats pIndex
       
    End If
End With
 
Exit Sub
 
Errhandler:
 
End Sub
 
Sub ia_EnviarChar(ByVal userindex As Integer, ByVal BotIndex As Byte)
' @                Envia el char del bot a un usuario (sistema de areas!!)
 
    With ia_Bot(BotIndex).Char
            Dim tmp_Color As Byte
           
            If ia_Bot(BotIndex).EsCriminal Then
               tmp_Color = 1
            Else
               tmp_Color = 3
            End If
           
            Call SendData(SendTarget.ToMap, ia_Bot(BotIndex).Char.CharIndex, ia_Bot(BotIndex).Pos.Map, "CC" & .Body & "," & .Head & "," & eHeading.SOUTH & "," & .CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y & "," & .WeaponAnim & "," & .ShieldAnim & " ," & 0 & "," & 0 & "," & .CascoAnim & "," & ia_Bot(BotIndex).Tag & "," & tmp_Color & "," & 0)
    End With
 
End Sub
Sub CalcularArea(ByVal BotIndex As Integer)

    Dim TempInt As Long
        TempInt = ia_Bot(BotIndex).Pos.X \ 9
        ia_Bot(BotIndex).AreasInfo.AreaReciveX = AreasRecive(TempInt)
        ia_Bot(BotIndex).AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = ia_Bot(BotIndex).Pos.Y \ 9
        ia_Bot(BotIndex).AreasInfo.AreaReciveY = AreasRecive(TempInt)
        ia_Bot(BotIndex).AreasInfo.AreaPerteneceY = 2 ^ TempInt

End Sub
Sub ia_UserDamage(ByVal spell As Byte, ByVal BotIndex As Byte, ByVal userindex As Integer, Optional ByVal is_RuneArea As Boolean = False)
 
Dim rMan     As Integer
Dim Damage   As Integer
 
'Checkeo que el hechizo no sea 0.
If Not spell <> 0 Then Exit Sub
 
With UserList(userindex)
 
    rMan = Hechizos(spell).ManaRequerido
   
    'Llega con la mana?
    If rMan > .Stats.MinMAN Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|18")
        Exit Sub
    End If
   
    If Hechizos(spell).Inmoviliza Or Hechizos(spell).Paraliza Then
        
        If ia_Bot(BotIndex).Paralizado = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "N|" & ia_Bot(BotIndex).Tag & " ya se encuentra paralizado.~69~190~156")
            Exit Sub
        End If
       
        'Le pongo el flag en verdadero.
        ia_Bot(BotIndex).Paralizado = True
       
        'Mensaje informando.
        Call SendData(SendTarget.toindex, userindex, 0, "N|Has paralizado a " & ia_Bot(BotIndex).Tag & "~69~190~156")
       
        'Creo la animacion sobre el char.
        Call SendData(SendTarget.ToMap, ia_Bot(BotIndex).Char.CharIndex, ia_Bot(BotIndex).Pos.Map, "CFX" & ia_Bot(BotIndex).Char.CharIndex & "," & Hechizos(spell).FXgrh & "," & Hechizos(spell).loops)
       
        'SpellWorlds.
        DecirPalabrasMagicas Hechizos(spell).PalabrasMagicas, userindex
       
        'Quito mana y energia
        .Stats.MinMAN = .Stats.MinMAN - rMan
       
        'le doy intervalo
       
        ia_Bot(BotIndex).Intervalos.ParalizisCount = (IA_SREMO / 10)
       
        SendUserMP userindex
       
        Exit Sub
    End If
   
    'Era un Viajante
   
    Damage = RandomNumber(Hechizos(spell).MinHP, Hechizos(spell).MaxHP)
    Damage = Damage + Porcentaje(Damage, 3 * .Stats.ELV)
   
   If Not Damage <> 0 Then Exit Sub
    'Quitamos vida
    If Hechizos(spell).StaffAffected Then
       If UCase$(UserList(userindex).clase) = "MAGO" Then
          If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
             Damage = (Damage * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
          Else
             Damage = Damage * 0.7 'Baja damage a 70% del original
          End If
        End If
     End If
   
    Damage = Damage - RandomNumber(52, 60)
    ia_Bot(BotIndex).Vida = ia_Bot(BotIndex).Vida - Damage
   
    'No está paralizado.
    If Not ia_Bot(BotIndex).Paralizado Then
        'Le pegaron, se cagó todo y se mueve random.
        Dim keepMoving  As Boolean
   
        ia_RandomMoveChar BotIndex, userindex, keepMoving
   
        'No hubo error, por ende se movió.
        If Not keepMoving Then
           'Guardo la nueva pos.
           MapData(ia_Bot(BotIndex).Pos.Map, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = BotIndex
       
           'Actualizo el area del bot.
           'Call ModAreas.SendToAreaByPos(ia_Bot(BotIndex).Pos.Map, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y, "+" & ia_Bot(BotIndex).Char.CharIndex & "," & ia_Bot(BotIndex).Pos.X & "," & ia_Bot(BotIndex).Pos.Y)
       
           'Intervalo de caminata.
           ia_Bot(BotIndex).Intervalos.MoveCharCount = (IA_MOVINT / 40)
        End If
       
    End If
   
    'Aviso al usuario.
    Call SendData(SendTarget.toindex, userindex, 0, "N|Le has quitado " & CStr(Damage) & " puntos de vida a " & ia_Bot(BotIndex).Tag & ".~255~0~0")
   
    'Tiro las spell worlds
    DecirPalabrasMagicas Hechizos(spell).PalabrasMagicas, userindex
   
    'Creo el fx.
    Call SendData(SendTarget.ToMap, ia_Bot(BotIndex).Char.CharIndex, ia_Bot(BotIndex).Pos.Map, "CFX" & ia_Bot(BotIndex).Char.CharIndex & "," & Hechizos(spell).FXgrh & "," & Hechizos(spell).loops)
   
    'saco mana y energia y actualizo el cliente
    .Stats.MinMAN = .Stats.MinMAN - rMan
       
    SendUserMP userindex
   
    If ia_Bot(BotIndex).Vida <= 0 Then
        'Murió?
        ia_EraseChar BotIndex, True
        Call SendData(SendTarget.toindex, userindex, 0, "N|Has matado a " & ia_Bot(BotIndex).Tag & "~69~190~156")
    End If
   
End With
 
End Sub
 
Sub ia_DamageHit(ByVal BotIndex As Byte, ByVal userindex As Integer)

Dim nDamage      As Integer
 
'Calculo el daño.
nDamage = CalcularDaño(userindex)
 
'Resto la defensa del bot.
nDamage = nDamage - (RandomNumber(IA_MINDEF, IA_MAXDEF))
 
'Aviso al usuario.
Call SendData(SendTarget.toindex, userindex, 0, "N|Le has pegado en el torso a " & ia_Bot(BotIndex).Tag & " por " & nDamage & "~255~0~0~1")
 
'Creo daño :)
'ia_SendToBotArea BotIndex, mod_DunkanProtocol.Send_CreateDamage(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y, nDamage)
 
'Resto vida.
ia_Bot(BotIndex).Vida = ia_Bot(BotIndex).Vida - nDamage
 
'seteo el flag.
'UserList(UserIndex).AtacoViajante = BotIndex
 
'Murio?
If ia_Bot(BotIndex).Vida <= 0 Then
 
    'Era viajante?
    If ia_Bot(BotIndex).Viajante Then
       'Reset el flag.
       'UserList(UserIndex).AtacoViajante = 0
    End If
   
    ia_EraseChar BotIndex, True
   
End If
 
End Sub
 
Sub ia_SendToBotArea(ByVal BotIndex As Byte, ByVal PackData As String)
' @                Envia paquetes al area de un bot.
 
'Nueva versión del sub, más simple y diría que más práctica : P
 
With ia_Bot(BotIndex)
    'd3 ao, borro esto : p
   
    'con esto tenemos algo simple, cuando mandamos el send
    'tobotarea, nos devuelve un array con los ui y el ping de cada
    'uno, y flush_ping tiene el promedio :), despues solo nos
    'queda comprobar si el usuario puede flushbuffear los datos
    'y enviamos, sacrificamos memoria pero ganamos MUCHA conexión.
   
    'Dim flush_Ping      As Integer
    'Dim arr_PingUsers() As Integer
   
    'Call modSendData.SendToAreaByPos(.Pos.map, .Pos.X, .Pos.Y, PackData, .GrupoID, flush_Ping)
   
    'Do While flush_Ping <> 0
    '    If can_Update_Ping(arr_PingUsers(flush_Ping)) Then
    '       Call flusH_buffer_to_base_Ping(arr_PingUsers(flush_Ping), flush_Ping, .GrupoID)
    '    End If
       
    '    flush_Ping = flush_Ping - 1
       
    'Loop
   
    Call ModAreas.SendToAreaByPos(.Pos.Map, .Pos.X, .Pos.Y, PackData)
End With
 
End Sub
 
Sub ia_TirarInventario(ByVal BotIndex As Byte)
' @note         :  Pincha el inventario de un bot.
 
Dim loopX   As Long
Dim iObjs() As Integer
Dim iObj    As Obj
Dim tmpPos  As WorldPos
 
'Arma array de objetos
ia_ArrayObjetos iObjs, BotIndex
 
For loopX = 1 To UBound(iObjs())
 
    'Crea el objeto.
    iObj.ObjIndex = iObjs(loopX)
 
    'Si el objIndex es >= 36 and <=30  , son pociones
    If iObjs(loopX) >= 36 And iObjs(loopX) <= 39 Then
       iObj.Amount = RandomNumber(1000, 1200)
    Else
       'No eran pociones, son flechas?
       If Not iObjs(loopX) <> 553 Then
          iObj.Amount = RandomNumber(500, 900)
       Else
          iObj.Amount = 1
       End If
    End If
   
    'Si eran pociones azules y el bot era caza..
    If iObj.Amount = 37 And ia_Bot(BotIndex).clase = eIAClase.Cazador Then iObj.Amount = 0
   
    'si hay objIndex.
    If iObj.ObjIndex Then
        'Busca un tile libre.
        Call Tilelibre(ia_Bot(BotIndex).Pos, tmpPos, iObj)
   
        'Si encontró (raro que no encuentre)
        If tmpPos.X <> 0 And tmpPos.Y <> 0 Then
           'Crea el objeto
           'MakeObj iObj, tmpPos.Map, tmpPos.X, tmpPos.Y
        End If
    End If
   
Next loopX
 
'Ya tiro los objetos de su equipo, ahora , si era viajante, tira los que lukeo, si es que tiene.
If ia_Bot(BotIndex).Viajante Then
   For loopX = 1 To IA_SLOTS
       With ia_Bot(BotIndex).Inv(loopX)
           
            iObj.ObjIndex = .ObjIndex
            iObj.Amount = .Amount
           
            Call Tilelibre(ia_Bot(BotIndex).Pos, tmpPos, iObj)
           
            'Si encontró posición.
            If tmpPos.X <> 0 And tmpPos.Y <> 0 Then
               'MakeObj iObj, tmpPos.Map, tmpPos.X, tmpPos.Y
            End If
       End With
   Next loopX
End If
 
End Sub
 
Sub ia_ArrayObjetos(ByRef arrayObjs() As Integer, ByVal BotIndex As Byte)
' @note         :  Arma un array de objetos.
 
'Set primeras dimensiones. (potas,arma y casco)
 
ReDim arrayObjs(1 To 4) As Integer
 
'Pociones.
arrayObjs(1) = 38
arrayObjs(2) = 37
 
'Arma
arrayObjs(3) = ia_ArmaByClase(BotIndex)
 
'Casco
arrayObjs(4) = ia_CascoByClase(BotIndex)
 
'Si no es mago, tiene escudo y dopas.
If ia_Bot(BotIndex).clase <> eIAClase.Mago Then
   'redim
   ReDim Preserve arrayObjs(1 To 7) As Integer
   arrayObjs(5) = ia_EscudoByClase(BotIndex)
   arrayObjs(6) = 36
   arrayObjs(7) = 39
End If
 
'Si es caza, tira flechas.
'No sabemos el ultimo elemento que tenemos!! no jugarsela a tirar 5.
 
If ia_Bot(BotIndex).clase = eIAClase.Cazador Then
   ReDim Preserve arrayObjs(1 To UBound(arrayObjs()) + 1) As Integer
   arrayObjs(UBound(arrayObjs())) = 553
End If
 
End Sub
Sub ia_EraseChar(ByVal BotIndex As Byte, Optional ByVal killedbyUSER As Boolean = False)
' @note         :  Borra el char y los datos del bot.
 
With ia_Bot(BotIndex)
    Call SendToUserAreaButindexBOT(BotIndex, "BP" & .Char.CharIndex)
   
    'Borro el botIndex
    MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
   
    Dim dummyPos    As WorldPos
   
    .ViajanteAntes = dummyPos
   
    'Mató un usuario? pincha inventario
    If killedbyUSER Then
       ia_TirarInventario BotIndex
    End If
   
    'Reset char,
    With .Char
         .Body = 0
         .CascoAnim = 0
         .FX = 0
         .loops = 0
         .Head = 0
         .Heading = 0
         .ShieldAnim = 0
         .WeaponAnim = 0
    End With
   
    'Reset STATS
    .Vida = 0
    .Mana = 0
   
    'Reset pos.
    With .Pos
         .Map = 0
         .X = 0
         .Y = 0
    End With
   
    'Reset flags.
    .Invocado = False
    .Paralizado = False
   
    'Reset intervalos.
    With .Intervalos
         .MoveCharCount = 0
         .SpellCount = 0
         .UseItemCount = 0
         .ParalizisCount = 0
    End With
   
    'Reset viajante flag.
    .Viajante = False
    .ViajanteUser = 0
   
    'Resta el contador
    NumInvocados = NumInvocados - 1
   
End With
 
End Sub
 
Sub ia_CheckInts(ByVal BotIndex As Byte)
 
With ia_Bot(BotIndex).Intervalos
     
    If .ArrowCount > 0 Then .ArrowCount = .ArrowCount - 1
    If .MoveCharCount > 0 Then .MoveCharCount = .MoveCharCount - 1
    If .SpellCount > 0 Then .SpellCount = .SpellCount - 1
    If .UseItemCount > 0 Then .UseItemCount = .UseItemCount - 1
    If .ParalizisCount > 0 Then .ParalizisCount = .ParalizisCount - 1
    If .HitCount > 0 Then .HitCount = .HitCount - 1
    If .ChatCount > 0 Then .ChatCount = .ChatCount - 1
   
End With
 
End Sub
Function ia_FindTarget(Pos As WorldPos, Optional ByVal esPk As Boolean = False) As Integer
' @note         :  Busca alguien a quien pegar
 
Dim loopX       As Long         '< Bucle del tileX.
Dim LoopY       As Long         '< Bucle del tileY.
Dim tmpIndex    As Integer
 
For LoopY = Pos.Y - (MinYBorder + 1) To Pos.Y + (MinYBorder - 1)
        For loopX = Pos.X - (MinXBorder + 1) To Pos.X + (MinXBorder - 1)
            'Hay usuario?
            If MapData(Pos.Map, loopX, LoopY).userindex > 0 Then
               'No está muerto
               If UserList(MapData(Pos.Map, loopX, LoopY).userindex).flags.Muerto = 0 Then
                  'Es ciuda el bot y el usuario?
                  If Not esPk Then
                     'el bot no es pk.
                     ia_FindTarget = MapData(Pos.Map, loopX, LoopY).userindex
                  Else
                     tmpIndex = MapData(Pos.Map, loopX, LoopY).userindex
                     If Not esPk And Criminal(tmpIndex) Then
                         ia_FindTarget = tmpIndex
                     Else
                        If esPk And Not Criminal(tmpIndex) Then
                           ia_FindTarget = tmpIndex
                        End If
                    End If
                 End If
                  Exit Function
               End If
            End If
        Next loopX
Next LoopY
 
ia_FindTarget = 0
End Function
 
Function IA_GetNextSlot() As Byte
' @ Devuelve un slot para bots.
 
Dim loopX   As Long
 
For loopX = 1 To MAX_BOTS
    If Not ia_Bot(loopX).Invocado Then
       IA_GetNextSlot = CByte(loopX)
       Exit Function
    End If
Next loopX
 
IA_GetNextSlot = 0
 
End Function
