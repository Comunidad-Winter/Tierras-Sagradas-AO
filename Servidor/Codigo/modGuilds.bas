Attribute VB_Name = "modGuilds"
Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const ORDENARLISTADECLANES = True
'True si se envia la lista ordenada por alineacion

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Public Guilds()         As clsClan
'array global de guilds, se indexa por userlist().guildindex

Public Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Public Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Public GMsEscuchando As New Collection
Public UsuariosEnCvcClan1 As Long
Public UsuariosEnCvcClan2 As Long
Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum
'alineaciones permitidas

Public forminfo As String

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    Guerra = -1
    PAZ = 0
    Aliados = 1
End Enum


'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()

Dim CantClanes  As String
Dim i           As Integer
Dim TempStr     As String
Dim Alin        As ALINEACION_GUILD
    'Dim Leer As clsIniReader
'Set Leer = LeerClan
    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
    If IsNumeric(CantClanes) Then
        CANTIDADDECLANES = CInt(CantClanes)
    Else
        CANTIDADDECLANES = 0
    End If
    
    For i = 1 To CANTIDADDECLANES
        Set Guilds(i) = New clsClan
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
        Call Guilds(i).Inicializar(TempStr, i, Alin)
    Next i
    
End Sub

Public Sub LoadGuildsClanes()

Dim CantClanes  As String
Dim ii           As Integer
Dim TempStrr     As String

    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
    If IsNumeric(CantClanes) Then
        CANTIDADDECLANES = CInt(CantClanes)
    Else
        CANTIDADDECLANES = 0
    End If
    
    For ii = 1 To CANTIDADDECLANES
        TempStrr = GetVar(GUILDINFOFILE, "GUILD" & ii, "GUILDNAME")
        Call Guilds(ii).InicializarNombresClanes(TempStrr)
    Next ii
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal userindex As Integer, ByVal GuildIndex As Integer) As Boolean
Dim NuevoL  As Boolean
Dim NuevaA  As Boolean
Dim News    As String

    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(userindex, GuildIndex) Then
        Call Guilds(GuildIndex).ConectarMiembro(userindex)
        UserList(userindex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    End If

End Function


Public Function m_ValidarPermanencia(ByVal userindex As Integer, ByVal SumaAntifaccion As Boolean, ByRef CambioAlineacion As Boolean, ByRef CambioLider As Boolean) As Boolean
Dim GuildIndex  As Integer
Dim ML          As String
Dim M           As String
Dim UI          As Integer
Dim Sale        As Boolean
Dim i           As Integer

    m_ValidarPermanencia = True
    GuildIndex = UserList(userindex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(userindex, GuildIndex) Then
    
        m_ValidarPermanencia = False
        If SumaAntifaccion Then Guilds(GuildIndex).PuntosAntifaccion = Guilds(GuildIndex).PuntosAntifaccion + 1
        
        CambioAlineacion = (m_EsGuildFounder(UserList(userindex).Name, GuildIndex) Or Guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION)
        
        If CambioAlineacion Then
            'aca tenemos un problema, el fundador acaba de cambiar el rumbo del clan o nos zarpamos de antifacciones
            'Tenemos que resetear el lider, revisar si el lider permanece y si no asignarle liderazgo al fundador

            Call Guilds(GuildIndex).CambiarAlineacion(ALINEACION_NEUTRO)
            Guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            'para la nueva alineacion, hay que revisar a todos los Pjs!

            'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
            'internamente al iterador en el proceso
            CambioLider = False
            i = 1
            ML = Guilds(GuildIndex).GetMemberList(",")
            M = ReadField(i, ML, Asc(","))
            While M <> vbNullString

                'vamos a violar un poco de capas..
                UI = NameIndex(M)
                If UI > 0 Then
                    Sale = Not m_EstadoPermiteEntrar(UI, GuildIndex)
                Else
                    Sale = Not m_EstadoPermiteEntrarChar(M, GuildIndex)
                End If

                If Sale Then
                    If m_EsGuildFounder(M, GuildIndex) Then 'hay que sacarlo de las armadas
                        If UI > 0 Then
                            UserList(UI).Faccion.FuerzasCaos = 0
                            UserList(UI).Faccion.ArmadaReal = 0
                            UserList(UI).Faccion.Reenlistadas = 1
                        Else
                            If FileExist(CharPath & M & ".chr") Then
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "EjercitoCaos", 0)
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "ArmadaReal", 0)
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "Reenlistadas", 1)
                                Call WriteVar(CharPath & M & ".chr", "FLAGS", "CJerarquia", 0)
                            End If
                        End If
                        m_ValidarPermanencia = True
                    Else    'sale si no es guildfounder
                        If m_EsGuildLeader(M, GuildIndex) Then
                            'pierde el liderazgo
                            CambioLider = True
                            Call Guilds(GuildIndex).SetLeader(Guilds(GuildIndex).Fundador)
                        End If

                        Call m_EcharMiembroDeClan(-1, M)
                    End If
                End If
                i = i + 1
                M = ReadField(i, ML, Asc(","))
            Wend
        Else
            'no se va el fundador, el peor caso es que se vaya el lider
            Call m_EcharMiembroDeClan(-1, UserList(userindex).Name)   'y lo echamos
        End If
    End If
    

End Function

Public Sub m_DesconectarMiembroDelClan(ByVal userindex As Integer, ByVal GuildIndex As Integer)
    If UserList(userindex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call Guilds(GuildIndex).DesConectarMiembro(userindex)
End Sub

Public Function m_EsGuildSubLeader1(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildSubLeader1 = (UCase$(PJ) = UCase$(Trim$(Guilds(GuildIndex).GetSubLider1)))
End Function
Public Function m_EsGuildSubLeader2(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildSubLeader2 = (UCase$(PJ) = UCase$(Trim$(Guilds(GuildIndex).GetSubLider2)))
End Function

Public Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(Guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(Guilds(GuildIndex).Fundador)))
End Function


'Public Function GetLeader(ByVal GuildIndex As Integer) As String
'    GetLeader = vbNullString
'
'    If GuildIndex <= 0 Then Exit Function
'    GetLeader = Guilds(GuildIndex).GetLeader()
'End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del clan de Expulsado
Dim userindex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeClan = 0

    userindex = NameIndex(Expulsado)
    If userindex > 0 Then
        'pj online
        GI = UserList(userindex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
            If m_EsGuildSubLeader1(Expulsado, GI) Then
             Call WriteVar(GUILDINFOFILE, "GUILD" & GI, "SubLider1", "Fermin")
            ElseIf m_EsGuildSubLeader2(Expulsado, GI) Then
             Call WriteVar(GUILDINFOFILE, "GUILD" & GI, "SubLider2", "Fermin")
            End If
            
            If UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 120 Then Exit Function
            
            If m_EsGuildLeader(Expulsado, GI) Then Guilds(GI).SetLeader (Guilds(GI).Fundador)
                Call Guilds(GI).DesConectarMiembro(userindex)
                Call Guilds(GI).ExpulsarMiembro(Expulsado)
                UserList(userindex).GuildIndex = 0
                Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetGuildIndexFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
            If m_EsGuildSubLeader1(Expulsado, GI) Then
             Call WriteVar(GUILDINFOFILE, "GUILD" & GI, "SubLider1", "Fermin")
            ElseIf m_EsGuildSubLeader2(Expulsado, GI) Then
             Call WriteVar(GUILDINFOFILE, "GUILD" & GI, "SubLider2", "Fermin")
            End If
                If m_EsGuildLeader(Expulsado, GI) Then Guilds(GI).SetLeader (Guilds(GI).Fundador)
                Call Guilds(GI).ExpulsarMiembro(Expulsado)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If

End Function

Public Sub ActualizarWebSite(ByVal userindex As Integer, ByRef Web As String)
Dim GI As Integer

    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then Exit Sub
    
    Call Guilds(GI).SetURL(Web)
    
End Sub


Public Sub ActualizarCodexYDesc(ByRef Datos As String, ByVal GuildIndex As Integer)
Dim CantCodex       As Integer
Dim i               As Integer

    If GuildIndex = 0 Then Exit Sub
    Call Guilds(GuildIndex).SetDesc(ReadField(1, Datos, Asc("¬")))
    CantCodex = CInt(ReadField(2, Datos, Asc("¬")))
    For i = 1 To CantCodex
        Call Guilds(GuildIndex).SetCodex(i, ReadField(2 + i, Datos, Asc("¬")))
    Next i
    For i = CantCodex + 1 To CANTIDADMAXIMACODEX
        Call Guilds(GuildIndex).SetCodex(i, vbNullString)
    Next i

End Sub

Public Sub ActualizarNoticias(ByVal userindex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(userindex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then Exit Sub
    
    Call Guilds(GI).SetGuildNews(Datos)
        
End Sub

Public Function CrearNuevoClan(ByRef GuildInfo As String, ByVal FundadorIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
Dim GuildName       As String
Dim Descripcion     As String
Dim URL             As String
Dim codex()         As String
Dim CantCodex       As Integer
Dim i               As Integer
Dim DummyString     As String

    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    GuildName = Trim$(ReadField(2, GuildInfo, Asc("¬")))

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
        Exit Function
    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function
    End If

    Descripcion = ReadField(1, GuildInfo, Asc("¬"))
    URL = ReadField(3, GuildInfo, Asc("¬"))
    CantCodex = CInt(ReadField(4, GuildInfo, Asc("¬")))

    If CantCodex > 0 Then
        ReDim codex(1 To CantCodex) As String
        For i = 1 To CantCodex
            codex(i) = ReadField(4 + i, GuildInfo, Asc("¬"))
        Next i
    End If

    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(Guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

        'constructor custom de la clase clan
        Set Guilds(CANTIDADDECLANES) = New clsClan
        Call Guilds(CANTIDADDECLANES).Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
        
        'Damos de alta al clan como nuevo inicializando sus archivos
        Call Guilds(CANTIDADDECLANES).InicializarNuevoClan(UserList(FundadorIndex).Name)
        
        'seteamos codex y descripcion
        For i = 1 To CantCodex
            Call Guilds(CANTIDADDECLANES).SetCodex(i, codex(i))
        Next i
        Call Guilds(CANTIDADDECLANES).SetDesc(Descripcion)
        Call Guilds(CANTIDADDECLANES).SetGuildNews("Clan creado con alineación : " & Alineacion2String(Alineacion))
        Call Guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).Name)
        Call Guilds(CANTIDADDECLANES).SetURL(URL)
        
        '"conectamos" al nuevo miembro a la lista de la clase
        Call Guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).Name)
        Call Guilds(CANTIDADDECLANES).ConectarMiembro(FundadorIndex)
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call WarpUserChar(FundadorIndex, UserList(FundadorIndex).Pos.Map, UserList(FundadorIndex).Pos.X, UserList(FundadorIndex).Pos.Y, False)
        
        For i = 1 To CANTIDADDECLANES - 1
            Call Guilds(i).ProcesarFundacionDeOtroClan
        Next i
    Else
        refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
        Exit Function
    End If
    
    CrearNuevoClan = True
    Call QuitarObjetos(939, 1, FundadorIndex)
    Call QuitarObjetos(1048, 1, FundadorIndex)

End Function

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios = PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    m_PuedeSalirDeClan = UCase$(Guilds(GuildIndex).GetLeader) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal userindex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean

    PuedeFundarUnClan = False
    If UserList(userindex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    If UserList(userindex).Stats.ELV < 50 Or UserList(userindex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
        refError = "Para fundar un clan debes ser nivel 50 y tener 100 en liderazgo."
        Exit Function
    End If
    
    If Not TieneObjetos(939, 1, userindex) Then
            refError = "Necesitas un Amuleto de Lider para fundar un clan."
            Exit Function
    End If
    
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(userindex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro de la armada."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            If Not Ciudadano(userindex) Then
                refError = "Para fundar un clan de ciudadanos debes ser ciudadano."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Not Criminal(userindex) Then
                refError = "Para fundar un clan de criminales debes ser criminal."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(userindex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legión oscura"
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            If Not Neutral(userindex) Then
                refError = "Para fundar un clan neutro no debes pertenecer a ninguna facción."
                Exit Function
            End If
    End Select
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
Dim Promedio    As Long
Dim ELV         As Integer
Dim f           As Byte

    m_EstadoPermiteEntrarChar = False
    
    Personaje = Replace(Personaje, "\", vbNullString)
    Personaje = Replace(Personaje, "/", vbNullString)
    Personaje = Replace(Personaje, ".", vbNullString)
    
    If FileExist(CharPath & Personaje & ".chr") Then
        Promedio = CByte(GetVar(CharPath & Personaje & ".chr", "STATUS", "EsStatus"))
        Select Case Guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_ARMADA
                If Promedio = 1 Or Promedio = 3 Or Promedio = 5 Then
                    m_EstadoPermiteEntrarChar = True
                End If
            Case ALINEACION_GUILD.ALINEACION_CIUDA
                If Promedio = 1 Or Promedio = 3 Or Promedio = 5 Then
                    m_EstadoPermiteEntrarChar = True
                End If
            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
                If Promedio = 2 Or Promedio = 4 Or Promedio = 6 Then
                    m_EstadoPermiteEntrarChar = True
                End If
            Case ALINEACION_GUILD.ALINEACION_NEUTRO
                If Promedio = 0 Then
                    m_EstadoPermiteEntrarChar = True
                End If
            Case ALINEACION_GUILD.ALINEACION_LEGION
                If Promedio = 2 Or Promedio = 4 Or Promedio = 6 Then
                    m_EstadoPermiteEntrarChar = True
                End If
            Case Else
                m_EstadoPermiteEntrarChar = True
        End Select
    End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal userindex As Integer, ByVal GuildIndex As Integer) As Boolean
    Select Case Guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Ciudadano(userindex) And _
                    IIf(UserList(userindex).Stats.ELV >= 25, UserList(userindex).Faccion.ArmadaReal <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = Criminal(userindex) And _
                    IIf(UserList(userindex).Stats.ELV >= 25, UserList(userindex).Faccion.FuerzasCaos <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = Neutral(userindex)
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Ciudadano(userindex)
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = Criminal(userindex)
        Case Else   'game masters
            m_EstadoPermiteEntrar = True
    End Select
End Function


Public Function String2Alineacion(ByRef s As String) As ALINEACION_GUILD
    Select Case s
        Case "Neutro"
            String2Alineacion = ALINEACION_NEUTRO
        Case "Legión oscura"
            String2Alineacion = ALINEACION_LEGION
        Case "Armada Real"
            String2Alineacion = ALINEACION_ARMADA
        Case "Game Masters"
            String2Alineacion = ALINEACION_MASTER
        Case "Legal"
            String2Alineacion = ALINEACION_CIUDA
        Case "Criminal"
            String2Alineacion = ALINEACION_CRIMINAL
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutro"
        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Legión oscura"
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Armada Real"
        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            Alineacion2String = "Legal"
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.Aliados
            Relacion2String = "A"
        Case RELACIONES_GUILD.Guerra
            Relacion2String = "G"
        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"
        Case RELACIONES_GUILD.Aliados
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal s As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(s))
        Case vbNullString, "P"
            String2Relacion = PAZ
        Case "G"
            String2Relacion = Guerra
        Case "A"
            String2Relacion = Aliados
        Case Else
            String2Relacion = PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

'old function by morgo

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))

    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        GuildNameValido = False
        Exit Function
    End If
    
Next i

GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
Dim i   As Integer

YaExiste = False
GuildName = UCase$(GuildName)

For i = 1 To CANTIDADDECLANES
    YaExiste = (UCase$(Guilds(i).GuildName) = GuildName)
    If YaExiste Then Exit Function
Next i



End Function

Public Function v_AbrirElecciones(ByVal userindex As Integer, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(userindex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GuildIndex) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas"
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call Guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal userindex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_UsuarioVota = False
    GuildIndex = UserList(userindex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If

    If Not Guilds(GuildIndex).EleccionesAbiertas Then
        refError = "No hay elecciones abiertas en tu clan."
        Exit Function
    End If
    
    If InStr(1, Guilds(GuildIndex).GetMemberList(","), Votado, vbTextCompare) <= 0 Then
        refError = Votado & " no pertenece al clan"
        Exit Function
    End If

    If Guilds(GuildIndex).YaVoto(UserList(userindex).Name) Then
        refError = "Ya has votado, no puedes cambiar tu voto"
        Exit Function
    End If
    
    Call Guilds(GuildIndex).ContabilizarVoto(UserList(userindex).Name, Votado)
    v_UsuarioVota = True

End Function

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de clase
Dim i       As Integer
Dim Temps   As String
    PlayerName = Replace(PlayerName, "\", vbNullString)
    PlayerName = Replace(PlayerName, "/", vbNullString)
    PlayerName = Replace(PlayerName, ".", vbNullString)
    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0
    End If
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
'me da el indice del guildname
Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(Guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal userindex As Integer, ByVal GuildIndex As Integer) As String
Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = Guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> userindex And (UserList(i).flags.Privilegios < PlayerType.EventManager Or UserList(userindex).flags.Privilegios >= PlayerType.EventManager) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
            i = Guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function
Public Function SendFriendList(ByVal userindex As Integer, Optional UsuarioDeslogeo As String = "(NADIE)") As String
Dim tStr As String
Dim tInt As Integer

    tStr = 20 & ","
    For tInt = 1 To 20
        If UserList(userindex).flags.NombreAmigo(tInt) = "" Or UserList(userindex).flags.NombreAmigo(tInt) = " " Or UserList(userindex).flags.NombreAmigo(tInt) = "(NADIE)" Then UserList(userindex).flags.NombreAmigo(tInt) = "(Nadie)"
        If NameIndex(UserList(userindex).flags.NombreAmigo(tInt)) = 0 Or UCase$(UserList(userindex).flags.NombreAmigo(tInt)) = UCase$(UsuarioDeslogeo) Then
                tStr = tStr & UserList(userindex).flags.NombreAmigo(tInt) & "(OFF),"
        Else
                tStr = tStr & UserList(userindex).flags.NombreAmigo(tInt) & "(ON),"
        End If
    Next tInt
    SendFriendList = tStr
End Function
Public Function SendGuildsList(ByVal userindex As Integer) As String
Dim tStr As String
Dim tInt As Integer

    tStr = CANTIDADDECLANES & ","
    For tInt = 1 To CANTIDADDECLANES
        tStr = tStr & Guilds(tInt).GuildName & "-" & Guilds(tInt).Alineacion & "-" & Guilds(tInt).NivelClan & ","
    Next tInt
    SendGuildsList = tStr
End Function
Public Function SendGuildDetails(ByRef GuildName As String) As String
Dim tStr    As String
Dim GI      As Integer
Dim i       As Integer

    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Function
    
    tStr = Guilds(GI).NivelClan & "¬"
    tStr = tStr & Guilds(GI).Alineacion & "¬"
    tStr = tStr & Guilds(GI).GetReputacion & "¬"
    tStr = tStr & Guilds(GI).Fundador & "¬"
    tStr = tStr & Guilds(GI).GetFechaFundacion & "¬"
    tStr = tStr & Guilds(GI).GetLeader & "¬"
    tStr = tStr & Guilds(GI).GetSubLider1 & "¬"
    tStr = tStr & Guilds(GI).GetSubLider2 & "¬"
    tStr = tStr & Guilds(GI).GetURL & "¬"
    tStr = tStr & CStr(Guilds(GI).CantidadDeMiembros) & "¬"
    For i = 1 To CANTIDADMAXIMACODEX
        tStr = tStr & Guilds(GI).GetCodex(i) & "¬"
    Next i
    tStr = tStr & Guilds(GI).GetDesc & "¬"
    tStr = tStr & Guilds(GI).GuildName
    
    SendGuildDetails = tStr
End Function
Public Function SendGuildUserInfo(ByVal userindex As Integer) As String
Dim tStr    As String
Dim tStrx    As String
Dim tInt    As Integer
Dim CantAsp As Integer
Dim GI      As Integer
Dim i       As Integer

SendGuildUserInfo = vbNullString


If UserList(userindex).GuildIndex <= 0 Then Exit Function
    
GI = UserList(userindex).GuildIndex
    
If m_EsGuildLeader(UserList(userindex).Name, GI) Then Exit Function
If m_EsGuildSubLeader1(UserList(userindex).Name, GI) Then Exit Function
If m_EsGuildSubLeader2(UserList(userindex).Name, GI) Then Exit Function
    
CastilloNorte = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte")
CastilloSur = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur")
CastilloEste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste")
CastilloOeste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste")
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If

   'If Not m_EsGuildSubLeader(UserList(UserIndex).name, GI) Then Exit Function
    '<-------Lista de guilds ---------->
    
    UserInfo = Guilds(GI).PuntosClan & "¬"
    UserInfo = UserInfo & Guilds(GI).NivelClan & "¬"
    UserInfo = UserInfo & Guilds(GI).GetLeader & "¬"
    UserInfo = UserInfo & Guilds(GI).GetSubLider1 & "¬"
    UserInfo = UserInfo & Guilds(GI).GetSubLider2 & "¬"
    UserInfo = UserInfo & CastilloNorte & "¬"
    UserInfo = UserInfo & CastilloSur & "¬"
    UserInfo = UserInfo & CastilloOeste & "¬"
    UserInfo = UserInfo & CastilloEste & "¬"
    UserInfo = UserInfo & Guilds(GI).GetReputacion & "¬"
    UserInfo = UserInfo & Guilds(GI).Fundador & "¬"
    
    UserInfo = UserInfo & CANTIDADDECLANES & "¬"
    
    For tInt = 1 To CANTIDADDECLANES
        UserInfo = UserInfo & Guilds(tInt).GuildName & "-" & Guilds(tInt).NivelClan & "-" & Guilds(tInt).Alineacion & "¬"
    Next tInt
    
    '<-------Lista de miembros ---------->
    UserInfo = UserInfo & Guilds(GI).CantidadDeMiembros & "¬"
    UserInfo = UserInfo & Guilds(GI).GetMemberList("¬", False, True) & "¬"
    
    SendGuildUserInfo = UserInfo
   

End Function
Public Function SendGuildLeaderInfo(ByVal userindex As Integer) As String
Dim tStr    As String
Dim tStrx    As String
Dim tInt    As Integer
Dim CantAsp As Integer
Dim GI      As Integer
Dim i       As Integer

    SendGuildLeaderInfo = vbNullString
    GI = UserList(userindex).GuildIndex
    
CastilloNorte = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte")
CastilloSur = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur")
CastilloEste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste")
CastilloOeste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste")
Fortaleza = GetVar(IniPath & "configuracion.ini", "CASTILLO", "Fortaleza")
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If

    If Not m_EsGuildLeader(UserList(userindex).Name, GI) And Not m_EsGuildSubLeader1(UserList(userindex).Name, GI) And Not m_EsGuildSubLeader2(UserList(userindex).Name, GI) Then Exit Function

    '<-------Lista de guilds ---------->
    
    tStr = Guilds(GI).PuntosClan & "¬"
    tStr = tStr & Guilds(GI).NivelClan & "¬"
    tStr = tStr & Guilds(GI).GetLeader & "¬"
    tStr = tStr & Guilds(GI).GetSubLider1 & "¬"
    tStr = tStr & Guilds(GI).GetSubLider2 & "¬"
    For i = 1 To CANTIDADMAXIMACODEX
        tStr = tStr & Guilds(GI).GetCodex(i) & "¬"
    Next i
    tStr = tStr & CastilloNorte & "¬"
    tStr = tStr & CastilloSur & "¬"
    tStr = tStr & CastilloOeste & "¬"
    tStr = tStr & CastilloEste & "¬"
    tStr = tStr & Guilds(GI).GetDesc & "¬"
    tStr = tStr & Guilds(GI).GetURL & "¬"
    tStr = tStr & Guilds(GI).CVCG & "¬"
    tStr = tStr & Guilds(GI).CVCP & "¬"
    tStr = tStr & Guilds(GI).CASTIS & "¬"
    tStr = tStr & Guilds(GI).GetReputacion & "¬"
    
    tStr = tStr & CANTIDADDECLANES & "¬"
    
    For tInt = 1 To CANTIDADDECLANES
        tStr = tStr & Guilds(tInt).GuildName & "$" & Guilds(tInt).Alineacion & "$" & Guilds(tInt).NivelClan & "¬"
    Next tInt
    
    '<-------Lista de miembros ---------->
    tStr = tStr & Guilds(GI).CantidadDeMiembros & "¬"
    tStr = tStr & Guilds(GI).GetMemberList("¬", True) & "¬"
    
    '<------- Guild News -------->
    tStr = tStr & Replace(Guilds(GI).GetGuildNews, vbCrLf, "º") & "¬"
    
    '<------- Solicitudes ------->
    CantAsp = Guilds(GI).CantidadAspirantes()
    tStr = tStr & CantAsp & "¬"
    If CantAsp > 0 Then
        tStr = tStr & Guilds(GI).GetAspirantes("¬") & "¬"
    End If
    
    SendGuildLeaderInfo = tStr
   

End Function
Public Function SendGuildSubLeaderInfo(ByVal userindex As Integer) As String
Dim tStr    As String
Dim tInt    As Integer
Dim CantAsp As Integer
Dim GI      As Integer

    SendGuildSubLeaderInfo = vbNullString
    GI = UserList(userindex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    'If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Function
   If Not m_EsGuildSubLeader1(UserList(userindex).Name, GI) Then Exit Function
    '<-------Lista de guilds ---------->
    tStr = CANTIDADDECLANES & "¬"
    
    For tInt = 1 To CANTIDADDECLANES
        tStr = tStr & Guilds(tInt).GuildName & "¬"
    Next tInt
    
    '<-------Lista de miembros ---------->
    tStr = tStr & Guilds(GI).CantidadDeMiembros & "¬"
    tStr = tStr & Guilds(GI).GetMemberList("¬", True) & "¬"
    
    '<------- Guild News -------->
    tStr = tStr & Replace(Guilds(GI).GetGuildNews, vbCrLf, "º") & "¬"
    
    '<------- Solicitudes ------->
    CantAsp = Guilds(GI).CantidadAspirantes()
    tStr = tStr & CantAsp & "¬"
    If CantAsp > 0 Then
        tStr = tStr & Guilds(GI).GetAspirantes("¬") & "¬"
    End If

    SendGuildSubLeaderInfo = tStr

End Function


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = Guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = Guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = Guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)
    End If
End Function
Public Function r_DeclararGuerra(ByVal userindex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildGuerra)
    
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan"
        Exit Function
    End If

    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, Guerra)
    Call Guilds(GIG).SetRelacion(GI, Guerra)

    r_DeclararGuerra = GIG

End Function


Public Function r_AceptarPropuestaDePaz(ByVal userindex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDePaz = 0
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If Guilds(GI).GetRelacion(GIG) <> Guerra Then
        refError = "No estás en guerra con ese clan"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar"
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, PAZ)
    Call Guilds(GIG).SetRelacion(GI, PAZ)
    
    r_AceptarPropuestaDePaz = GIG

End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal userindex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(userindex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, Aliados) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function
    End If
    
    Call Guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call Guilds(GIG).SetGuildNews(Guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & Guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal userindex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(userindex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function
    End If
    
    Call Guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call Guilds(GIG).SetGuildNews(Guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & Guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function


Public Function r_AceptarPropuestaDeAlianza(ByVal userindex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If Guilds(GI).GetRelacion(GIG) <> PAZ Then
        refError = "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, Aliados) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, Aliados)
    Call Guilds(GIG).SetRelacion(GI, Aliados)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function


Public Function r_ClanGeneraPropuesta(ByVal userindex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
Dim OtroClanGI      As Integer
Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan"
        Exit Function
    End If
    
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe!"
        Exit Function
    End If
    
    If Guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = PAZ Then
        If Guilds(GI).GetRelacion(OtroClanGI) <> Guerra Then
            refError = "No estás en guerra con " & OtroClan
            Exit Function
        End If
    ElseIf Tipo = Guerra Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = Aliados Then
        If Guilds(GI).GetRelacion(OtroClanGI) <> PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function
        End If
    End If
    
    Call Guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal userindex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
Dim OtroClanGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not Guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada"
        Exit Function
    End If
    
    r_VerPropuesta = Guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal userindex As Integer, ByVal Tipo As RELACIONES_GUILD) As String
Dim GI  As Integer
Dim i   As Integer


    GI = UserList(userindex).GuildIndex
    If GI > 0 And GI <= CANTIDADDECLANES Then
        i = Guilds(GI).Iterador_ProximaPropuesta(Tipo)
        While i > 0
            r_ListaDePropuestas = r_ListaDePropuestas & Guilds(i).GuildName & ","
            i = Guilds(GI).Iterador_ProximaPropuesta(Tipo)
        Wend
        If Len(r_ListaDePropuestas) > 0 Then
            r_ListaDePropuestas = Left$(r_ListaDePropuestas, Len(r_ListaDePropuestas) - 1)
        End If
    End If

End Function

Public Function r_CantidadDePropuestas(ByVal userindex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
Dim GI As Integer
    GI = UserList(userindex).GuildIndex
    If GI > 0 And GI <= CANTIDADDECLANES Then
        r_CantidadDePropuestas = Guilds(GI).CantidadPropuestas(Tipo)
    End If
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal Guild As Integer, ByRef Detalles As String)
    Aspirante = Replace(Aspirante, "\", "")
    Aspirante = Replace(Aspirante, "/", "")
    Aspirante = Replace(Aspirante, ".", "")
    Call Guilds(Guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    Aspirante = Replace(Aspirante, "\", "")
    Aspirante = Replace(Aspirante, "/", "")
    Aspirante = Replace(Aspirante, ".", "")
    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal userindex As Integer, ByRef Nombre As String, ByRef motivo As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim UI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If

    NroAspirante = Guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan"
        Exit Function
    End If

    Call Guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & Guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal userindex As Integer, ByRef Nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) Or Not m_EsGuildSubLeader1(UserList(userindex).Name, GI) Or Not m_EsGuildSubLeader2(UserList(userindex).Name, GI) Then
        Exit Function
    End If
    
    NroAspirante = Guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = Guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function
Public Function a_NuevoAspirante(ByVal userindex As Integer, ByRef Clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(userindex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
        Exit Function
    End If
    
    If EsNewbie(userindex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(Clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe! Avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(userindex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineación " & Alineacion2String(Guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If

    If Guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "ASPIRANTEA")

    If ViejoSolicitado <> vbNullString Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = Guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(userindex).Name)
            If ViejoNroAspirante > 0 Then
                Call Guilds(ViejoGuildINdex).RetirarAspirante(UserList(userindex).Name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call Guilds(NuevoGuildIndex).NuevoAspirante(UserList(userindex).Name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal userindex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userindex).Name, GI) And Not m_EsGuildSubLeader1(UserList(userindex).Name, GI) And Not m_EsGuildSubLeader2(UserList(userindex).Name, GI) Then
        refError = "No eres el líder o SubLider de tu clan"
        Exit Function
    End If
    
    NroAspirante = Guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El personaje no es aspirante al clan"
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If UserList(AspiranteUI).GuildIndex > 0 Then
            refError = Aspirante & " ya esta en un clan."
            Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
        
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(Guilds(GI).Alineacion)
            Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
    
        If FileExist(CharPath & UCase$(Aspirante) & ".chr", vbNormal) = True Then
            If GetVar(CharPath & Aspirante & ".chr", "GUILD", "GUILDINDEX") > 0 Then
                refError = Aspirante & " ya esta en un clan."
                Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
                Exit Function
            End If
        
            If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
                refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(Guilds(GI).Alineacion)
                Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
               Exit Function
            End If
        End If
        
    End If
    'el pj es aspirante al clan y puede entrar
    
    Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call Guilds(GI).AceptarNuevoMiembro(Aspirante)

    a_AceptarAspirante = True

End Function


