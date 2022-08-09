Attribute VB_Name = "General"
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

'Global ANpc As Long
'Global Anpc_host As Long

Option Explicit

Global LeerNPCs As New clsIniReader
Global LeerNPCsHostiles As New clsIniReader

Sub DarCuerpoDesnudo(ByVal userindex As Integer, Optional ByVal Mimetizado As Boolean = False)

Select Case UCase$(UserList(userindex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 21
                    Else
                        UserList(userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 39
                    Else
                        UserList(userindex).Char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 32
                    Else
                        UserList(userindex).Char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 40
                    Else
                        UserList(userindex).Char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 53
                    Else
                        UserList(userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 60
                    Else
                        UserList(userindex).Char.Body = 60
                    End If
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 53
                    Else
                        UserList(userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 60
                    Else
                        UserList(userindex).Char.Body = 60
                    End If
      End Select
    Case Else
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 21
                    Else
                        UserList(userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 39
                    Else
                        UserList(userindex).Char.Body = 39
                    End If
      End Select
    
End Select

UserList(userindex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Sub EnviarSpawnList(ByVal userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(SendTarget.toindex, userindex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
    #If UsarQueSocket = 0 Then

        Obj.AddressFamily = AF_INET
        Obj.Protocol = IPPROTO_IP
        Obj.SocketType = SOCK_STREAM
        Obj.Binary = False
        Obj.Blocking = False
        Obj.BufferSize = 1024
        Obj.LocalPort = Port
        Obj.backlog = 5
        Obj.listen

    #End If
End Sub
Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Set AodefConv = New AoDefenderConverter

Call BanIpCargar
Call CargarPremiosList
Call CargarDonaciones
Call CargarQuests
Call Tesoros
Call Mod_BOTS.ia_Spells
Call CleanWorld_Initialize

Call LoadBalance
Call CargarIntervalos

CastilloNorte = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte")
CastilloSur = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur")
CastilloEste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste")
CastilloOeste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste")
Fortaleza = GetVar(IniPath & "configuracion.ini", "CASTILLO", "Fortaleza")

Prision.Map = 78
Libertad.Map = 28
Prision.X = 50
Prision.Y = 50
Libertad.X = 50
Libertad.Y = 50

TanaTelep.Map = 28
TanaTelep.X = 54
TanaTelep.Y = 36

HayAram = False
InvocoBicho = False
EspectadoresEnArena1 = 0
EspectadoresEnArena2 = 0
EspectadoresEnArena3 = 0
EspectadoresEnArena4 = 0

TModalidad = "0"
PuntosPremios = 0
 
MinutosRey = 20
PremiosCastis = 60
ChatGlobal = False
GRANK_Setup
CargarMisiones

    MensajeAutomatico = True
    TextoMensajeAutomatico = "¡Registrate en el foro para participar de los grandes eventos de Tierras Sagradas, ingresá en www.tierras-sagradas.com!~215~215~0~1~0"
    TiempoMensajeAutomatico = 10
    MinutitosMensaje = 0

LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Guilds(1 To MAX_GUILDS) As clsClan


IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"


Dim i As Byte
For i = 1 To STAT_MAXELV
    If val(i * 3) < 100 Then
        LevelSkill(i).LevelValue = val(i * 3)
    Else
        LevelSkill(i).LevelValue = 100
    End If
Next i

RejaNorte = 10000
RejaCentral = 10000
RejaSur = 10000
GranPoder = 0
AlmasNecesarias = GetVar(App.Path & "\Dioses\" & "Configuracion.ini", "INIT", "AlmasNecesarias")

ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Artesano"
ListaClases(12) = "Recolector"

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "DefensaMagica"

frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.caption = frmMain.caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).caption = "Cargando datos iniciales.."
frmCargando.Image1.Width = 0
Call LoadGuildsDB


Call CargarSpawnList
Call CargarForbidenWords
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿

MaxUsers = 0
BOnlines = 0
Call LoadSini
Call CargaApuestas
'Call LoadRanking

'*************************************************
Call CargaNpcsDat
'*************************************************

'Call LoadOBJData
Call LoadOBJData
Call CargarHechizos
Call CargarCofresRandom
    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero

If BootDelBackUp Then
    frmCargando.Label1(2).caption = "Cargando BackUp"
    Call CargarBackUp
Else
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo

MultiplicadorExp = GetVar(IniPath & "Server.ini", "INIT", "MultiplicadordeExp")
MultiplicadorOro = GetVar(IniPath & "Server.ini", "INIT", "MultiplicadordeOro")
MultiplicadorDrop = GetVar(IniPath & "Server.ini", "INIT", "MultiplicadordeDrop")

FragsJerarquia(1) = GetVar(IniPath & "Facciones.ini", "Jerarquias", "Primera")
FragsJerarquia(2) = GetVar(IniPath & "Facciones.ini", "Jerarquias", "Segunda")
FragsJerarquia(3) = GetVar(IniPath & "Facciones.ini", "Jerarquias", "Tercera")
FragsJerarquia(4) = GetVar(IniPath & "Facciones.ini", "Jerarquias", "Cuarta")


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim loopC As Integer

For loopC = 1 To 9
    NPCInvocaciones(loopC) = val(GetVar(IniPath & "configuracion.ini", "INVOCACIONES", "Npc" & loopC))
Next loopC

'Resetea las conexiones de los usuarios
For loopC = 1 To MaxUsers
    UserList(loopC).ConnID = -1
    UserList(loopC).ConnIDValida = False
Next loopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿


Call CargarExperiencia

Unload frmCargando


'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #n

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(file, FileType) <> ""
End Function

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = mid$(Text, LastPos + 1)
End If

End Function
Public Function Tilde(data As String) As String
 
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function
Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function
Sub MostrarNumUsers()
Dim TuPrimaTuerta As Integer
TuPrimaTuerta = NumUsers + BOnlines
Call SendData(ToAll, 0, 0, "ON" & TuPrimaTuerta)
frmMain.CantUsuarios.caption = "Numero de usuarios jugando: " & TuPrimaTuerta

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function
Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

Dim loopC As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For loopC = 1 To MaxUsers
    Call CloseSocket(loopC)
Next

ReDim UserList(1 To MaxUsers)

For loopC = 1 To MaxUsers
    UserList(loopC).ConnID = -1
    UserList(loopC).ConnIDValida = False
Next loopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

'Log it
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " servidor reiniciado."
Close #n

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub
Public Function TieneItemDiosEquipado(ByVal userindex As Integer) As Boolean
    
 Dim i As Long
For i = 1 To UserList(userindex).InventorySlots
  If UserList(userindex).Invent.Object(i).ObjIndex > 0 Then
    If ObjData(UserList(userindex).Invent.Object(i).ObjIndex).ItemDios = 1 And UserList(userindex).Invent.Object(i).Equipped = 1 Then
            TieneItemDiosEquipado = True
        Exit Function
    End If
  End If
Next i

TieneItemDiosEquipado = False
    
End Function
Public Function Intemperie(ByVal userindex As Integer) As Boolean
    
    If MapInfo(UserList(userindex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger <> 1 And _
           MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger <> 2 And _
           MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function
Public Sub TiempoInvocacion(ByVal userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal userindex As Integer)

Dim modifi As Integer

If UserList(userindex).Counters.Frio < IntervaloFrio Then
  UserList(userindex).Counters.Frio = UserList(userindex).Counters.Frio + 1
Else
  If MapInfo(UserList(userindex).Pos.Map).Terreno = Nieve Then
    Call SendData(SendTarget.toindex, userindex, 0, "||677")
    modifi = Porcentaje(UserList(userindex).Stats.MaxHP, 5)
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - modifi
    If UserList(userindex).Stats.MinHP < 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||678")
            UserList(userindex).Stats.MinHP = 0
              If userindex = GranPoder Then
                Call OtorgarGranPoder(0)
              End If
            Call UserDie(userindex)
    End If
    Call SendData(SendTarget.toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
  Else
    modifi = Porcentaje(UserList(userindex).Stats.MaxSta, 5)
    Call QuitarSta(userindex, modifi)
    Call SendData(SendTarget.toindex, userindex, 0, "ASS" & UserList(userindex).Stats.MinSta)
  End If
  
  UserList(userindex).Counters.Frio = 0
  Call SendUserHP(userindex)
  
End If

End Sub

Public Sub EfectoMimetismo(ByVal userindex As Integer)

If UserList(userindex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(userindex).Counters.Mimetismo = UserList(userindex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(SendTarget.toindex, userindex, 0, "||679")
    
    UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        
    
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
    Call ChangeUserChar(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
End If
            
End Sub
Public Sub EfectoInvisibilidad(ByVal userindex As Integer)

If UserList(userindex).Counters.Invisibilidad > 0 Then
     UserList(userindex).Counters.Invisibilidad = UserList(userindex).Counters.Invisibilidad - 1
     Call SendData(SendTarget.toindex, userindex, 0, "INVI" & UserList(userindex).Counters.Invisibilidad)
Else
     UserList(userindex).Counters.Invisibilidad = 0
     UserList(userindex).flags.Invisible = 0
     If UserList(userindex).flags.Oculto = 0 Then
         Call SendData(SendTarget.toindex, userindex, 0, "||195")
         Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
         Call SendData(SendTarget.toindex, userindex, 0, "INVI0")
     End If
End If
End Sub
Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub
Public Sub EfectoParalisisUser(ByVal userindex As Integer)

If UserList(userindex).Counters.Paralisis > 0 Then
    UserList(userindex).Counters.Paralisis = UserList(userindex).Counters.Paralisis - 1
Else
    UserList(userindex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
End If

End Sub

Public Sub RecStamina(userindex As Integer, Intervalo As Integer)

'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub
   
If UserList(userindex).flags.Desnudo = 1 Then Exit Sub
If UserList(userindex).Stats.MinAGU = 0 Then Exit Sub
If UserList(userindex).Stats.MinHam = 0 Then Exit Sub

Dim massta As Integer
If UserList(userindex).Stats.MinSta < UserList(userindex).Stats.MaxSta Then
   If UserList(userindex).Counters.STACounter < Intervalo Then
       UserList(userindex).Counters.STACounter = UserList(userindex).Counters.STACounter + 1
   Else
       UserList(userindex).Counters.STACounter = 0
       massta = RandomNumber(80, 130)
       UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta + massta
       SendUserST (userindex)
       If UserList(userindex).Stats.MinSta > UserList(userindex).Stats.MaxSta Then
            UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(userindex As Integer)
Dim n As Integer

If UserList(userindex).Counters.Veneno < IntervaloVeneno Then
  UserList(userindex).Counters.Veneno = UserList(userindex).Counters.Veneno + 1
Else
  Call SendData(SendTarget.toindex, userindex, 0, "||680")
  UserList(userindex).Counters.Veneno = 0
  n = RandomNumber(1, 5)
  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - n
  If UserList(userindex).Stats.MinHP < 1 Then Call UserDie(userindex)
   If userindex = GranPoder And UserList(userindex).Stats.MinHP <= 0 Then
            Call OtorgarGranPoder(0)
    End If
  Call SendData(SendTarget.toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
  Call SendUserHP(userindex)
End If

End Sub

Public Sub DuracionPociones(userindex As Integer)

'Controla la duracion de las pociones
If UserList(userindex).flags.DuracionEfecto > 0 Then
   UserList(userindex).flags.DuracionEfecto = UserList(userindex).flags.DuracionEfecto - 1
   If UserList(userindex).flags.DuracionEfecto = 0 Then
        UserList(userindex).flags.TomoPocion = False
        UserList(userindex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(userindex).Stats.UserAtributos(loopX) = UserList(userindex).Stats.UserAtributosBackUP(loopX)
        Next
        
        SendUserAgilidad (userindex)
        SendUserFuerza (userindex)
   End If
End If

End Sub

Public Sub HambreYSed(userindex As Integer)

If UserList(userindex).flags.Privilegios = PlayerType.User Then
    'Sed
    If UserList(userindex).Stats.MinAGU > 0 Then
        If UserList(userindex).Counters.AGUACounter < IntervaloSed Then
              UserList(userindex).Counters.AGUACounter = UserList(userindex).Counters.AGUACounter + 1
        Else
              UserList(userindex).Counters.AGUACounter = 0
              UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU - 10
                                
              If UserList(userindex).Stats.MinAGU <= 0 Then
                   UserList(userindex).Stats.MinAGU = 0
                   UserList(userindex).flags.Sed = 1
              End If
              
              Call EnviarHambreYsed(userindex)
        End If
    End If
    
    'hambre
    If UserList(userindex).Stats.MinHam > 0 Then
       If UserList(userindex).Counters.COMCounter < IntervaloHambre Then
            UserList(userindex).Counters.COMCounter = UserList(userindex).Counters.COMCounter + 1
       Else
            UserList(userindex).Counters.COMCounter = 0
            UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam - 10
            If UserList(userindex).Stats.MinHam <= 0 Then
                   UserList(userindex).Stats.MinHam = 0
                   UserList(userindex).flags.Hambre = 1
            End If
            
            Call EnviarHambreYsed(userindex)
        End If
    End If
End If


End Sub
Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
Call LeerNPCs.Initialize(npcfile)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Call LeerNPCsHostiles.Initialize(npcfile)

End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub
Sub PasarSegundo()
    
    If CuentaRegresiva > 0 Then
        If CuentaRegresiva > 1 Then
            SendData SendTarget.ToMap, 0, MapaCont, "||455@" & CuentaRegresiva - 1
        Else
            SendData SendTarget.ToMap, 0, MapaCont, "||682"
            
            If UCase$(TModalidad) = "DM" Then
                TiroCuentaDM = True
            End If
        End If
        
        Dim ToL As Byte
        ToL = CuentaRegresiva - 1
        
        CuentaRegresiva = ToL
        SendData SendTarget.ToMap, 0, MapaCont, "CU" & ToL
    End If

    If SegundosInvo > 0 Then
      SegundosInvo = SegundosInvo - 1
      
        If SegundosInvo = 0 Then
          Dim invpos As WorldPos
          invpos.Map = mapainvo
          invpos.X = 29
          invpos.Y = 29
              
          Dim temandoelnpc As Byte
          temandoelnpc = RandomNumber(1, 100)
            
          If temandoelnpc >= 8 And temandoelnpc <= 75 Then
              Call SpawnNpc(NPCInvocaciones(RandomNumber(1, 5)), invpos, True, False)
          ElseIf temandoelnpc < 8 Then
              Call SpawnNpc(NPCInvocaciones(9), invpos, True, False)
          Else
              Call SpawnNpc(NPCInvocaciones(RandomNumber(6, 8)), invpos, True, False)
          End If
              
        End If
    End If
    
    If CuentaTorneo > 0 Then
        If CuentaTorneo > 1 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||683@" & CuentaTorneo - 1)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, "||684")
            Hay_Torneo = True
            UsuariosEnTorneo = 0
        End If
        
        CuentaTorneo = CuentaTorneo - 1
    End If
    
    If CuentaAutomatico > 0 Then
        If CuentaAutomatico > 1 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||683@" & CuentaAutomatico - 1)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, "||684")
            Torneo_Activo = True
        End If
        
        CuentaAutomatico = CuentaAutomatico - 1
    End If


Dim i As Integer
    For i = 1 To LastUser
    
                'Cerrar usuario
                If UserList(i).Counters.Saliendo Then
                    UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                    
                    If UserList(i).Counters.Salir <= 0 Then
                        Call SendData(SendTarget.toindex, i, 0, "||681")
                        Call SendData(SendTarget.toindex, i, 0, "FINOK")
                        
                        Call CloseSocket(i)
                    End If
                End If
    
If UserList(i).flags.UserLogged And UserList(i).ConnID <> -1 Then

    If UserList(i).Counters.InmoManopla > 0 Then
        UserList(i).Counters.InmoManopla = UserList(i).Counters.InmoManopla - 1
    End If
    
    'Nuevo invi
    If UserList(i).flags.AdminInvisible <> 1 And UserList(i).flags.Invisible = 1 Then Call EfectoInvisibilidad(i)
    
    'Nuevo paralisis
    If UserList(i).flags.Paralizado = 1 Then Call EfectoParalisisUser(i)

        If UserList(i).Counters.TimeComandos > 0 Then
            UserList(i).Counters.TimeComandos = UserList(i).Counters.TimeComandos - 1
            
        If UserList(i).flags.Privilegios >= PlayerType.UserSupport Then UserList(i).Counters.TimeComandos = 0
    End If
        
'ARAM
If UserList(i).flags.EnAram Then
    If UserList(i).flags.AramSeconds > 0 Then
        UserList(i).flags.AramSeconds = UserList(i).flags.AramSeconds - 1
        Call SendData(SendTarget.toindex, i, 0, "ARAM" & UserList(i).flags.AramSeconds)
        
        If UserList(i).flags.AramSeconds = 0 Then
            Call Aram_RevivirUsuario(i)
        End If
    End If
End If

If UserList(i).flags.TimeDeads > 0 Then
    UserList(i).flags.TimeDeads = UserList(i).flags.TimeDeads - 1
    
    If UserList(i).flags.TimeDeads = 0 Then
        UserList(i).flags.CantDeads = 0
    End If
End If
        
        
'TRANSPORTE CASTILLOS.
If UserList(i).Counters.TransporteCastillos(35) > 0 Then
    UserList(i).Counters.TransporteCastillos(35) = UserList(i).Counters.TransporteCastillos(35) - 1

    If UserList(i).Counters.TransporteCastillos(35) = 0 Then
            Call WarpUserChar(i, 81, RandomNumber(46, 52), RandomNumber(35, 41), True)
            Call SendData(SendTarget.toindex, i, 0, "||651@" & UserList(i).Name)
    End If
ElseIf UserList(i).Counters.TransporteCastillos(33) > 0 Then
    UserList(i).Counters.TransporteCastillos(33) = UserList(i).Counters.TransporteCastillos(33) - 1

    If UserList(i).Counters.TransporteCastillos(33) = 0 Then
            Call WarpUserChar(i, 33, 32, 63, True)
            Call SendData(SendTarget.toindex, i, 0, "||651@" & UserList(i).Name)
    End If
ElseIf UserList(i).Counters.TransporteCastillos(31) > 0 Then
    UserList(i).Counters.TransporteCastillos(31) = UserList(i).Counters.TransporteCastillos(31) - 1

    If UserList(i).Counters.TransporteCastillos(31) = 0 Then
            Call WarpUserChar(i, 31, 32, 63, True)
            Call SendData(SendTarget.toindex, i, 0, "||651@" & UserList(i).Name)
    End If
ElseIf UserList(i).Counters.TransporteCastillos(32) > 0 Then
    UserList(i).Counters.TransporteCastillos(32) = UserList(i).Counters.TransporteCastillos(32) - 1

    If UserList(i).Counters.TransporteCastillos(32) = 0 Then
            Call WarpUserChar(i, 32, 32, 63, True)
            Call SendData(SendTarget.toindex, i, 0, "||651@" & UserList(i).Name)
    End If
ElseIf UserList(i).Counters.TransporteCastillos(34) > 0 Then
    UserList(i).Counters.TransporteCastillos(34) = UserList(i).Counters.TransporteCastillos(34) - 1

    If UserList(i).Counters.TransporteCastillos(34) = 0 Then
            Call WarpUserChar(i, 34, 32, 63, True)
            Call SendData(SendTarget.toindex, i, 0, "||651@" & UserList(i).Name)
    End If
End If
'TRANSPORTE CASTILLOS

'Transporte premium

If UserList(i).Counters.TransportePremium > 0 Then
    UserList(i).Counters.TransportePremium = UserList(i).Counters.TransportePremium - 1
    
        If UserList(i).Counters.TransportePremium = 0 Then
            If UserList(i).UserPremiumMap = 0 Then
                    Call WarpUserChar(i, 1, 54, 36, True)
                    Call SendData(toindex, i, 0, "||348")
            ElseIf UserList(i).UserPremiumMap <> 0 Then
                
                If UserList(i).UserPremiumMap = 22 Then
                    Call WarpUserChar(i, 22, 60, 50, True)
                ElseIf UserList(i).UserPremiumMap = 90 Then
                    Call WarpUserChar(i, 90, 52, 56, True)
                ElseIf UserList(i).UserPremiumMap = 17 Then
                    Call WarpUserChar(i, 17, 55, 66, True)
                ElseIf UserList(i).UserPremiumMap = 93 Then
                    Call WarpUserChar(i, 93, 40, 22, True)
                End If
                
                Call SendData(toindex, i, 0, "||651@" & UserList(i).Name)
            End If
            
        End If
End If
      
If TesoroContando = True Then
    If UserList(i).flags.Desenterrando = 1 Then
            TiempoTesoro = TiempoTesoro - 1
            
            If UCase$(UserList(i).clase) = "LADRON" And UserList(i).Stats.ELV >= 45 Then TiempoTesoro = 0
           
        If TiempoTesoro = 0 Then
            Dim recompensitah As Byte
            Dim recompensitahoro As Long
            recompensitah = RandomNumber(1, 20)
            recompensitahoro = RandomNumber(GetVar(App.Path & "\Dat\" & "Tesoros.dat", "EXTRAS", "MinOro"), GetVar(App.Path & "\Dat\" & "Tesoros.dat", "EXTRAS", "MaxOro"))
            
            If recompensitah > 12 Then
                Call SendData(SendTarget.toindex, i, 0, "||57@10")
                Call AgregarPuntos(i, 10)
            End If
            
            If recompensitah > 5 Then
                Dim RecomTesoro As Obj, CantiObjs As Byte, RandomObjs As Byte
                CantiObjs = GetVar(App.Path & "\Dat\" & "Tesoros.dat", "" & UCase$(UserList(i).clase) & "", "CantObj")
                RandomObjs = RandomNumber(1, CantiObjs)
                RecomTesoro.Amount = 1
                RecomTesoro.ObjIndex = GetVar(App.Path & "\Dat\" & "Tesoros.dat", "" & UCase$(UserList(i).clase) & "", "Obj" & RandomObjs)
                
                If Not MeterItemEnInventario(i, RecomTesoro) Then
                    Call TirarItemAlPiso(UserList(i).Pos, RecomTesoro)
                End If
            End If
            
            If recompensitah > 18 Then
                Dim sacritesoro As Obj
                sacritesoro.ObjIndex = 936
                sacritesoro.Amount = 1

                If Not MeterItemEnInventario(i, sacritesoro) Then
                    Call TirarItemAlPiso(UserList(i).Pos, sacritesoro)
                End If
            End If
            
            UserList(i).Stats.GLD = UserList(i).Stats.GLD + recompensitahoro
            Call SendData(SendTarget.toindex, i, 0, "||63@" & PonerPuntos(recompensitahoro))
                        
            SendUserGLD (i)
            TesoroContando = False
            UserList(i).flags.Desenterrando = 0
            Call CofreAbierto
            Call Tesoros
        End If
            
    End If
End If

If UserList(i).Counters.CreoTeleport = True Then
            UserList(i).Counters.TimeTeleport = UserList(i).Counters.TimeTeleport + 1
            
    Dim mapa As Byte
    Dim X As Byte
    Dim Y As Byte
    
    mapa = UserList(i).flags.DondeTiroMap
    X = UserList(i).flags.DondeTiroX
    Y = UserList(i).flags.DondeTiroY

            If UserList(i).Counters.TimeTeleport = 8 Then
                Call EraseObj(ToMap, 0, UserList(i).flags.DondeTiroMap, MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).OBJInfo.Amount, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY)
                Dim ET As Obj
                ET.Amount = 1
                ET.ObjIndex = 378

                If MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).userindex > 0 Then
                    Call WarpUserChar(MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).userindex, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY + 1, True)
                End If

                Call MakeObj(ToMap, 0, UserList(i).flags.DondeTiroMap, ET, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY)
                MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).TileExit.Map = MapaPortal
                MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).TileExit.X = XPortal
                MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).TileExit.Y = YPortal
                
                If MapaPortal = 27 Then
                    Call SendData(toindex, i, 0, "||685")
                ElseIf MapaPortal = 29 Then
                    Call SendData(toindex, i, 0, "||686")
                End If
                
            ElseIf UserList(i).Counters.TimeTeleport >= 21 Then
                UserList(i).flags.TiroPortalL = 0
                UserList(i).Counters.TimeTeleport = 0
                UserList(i).Counters.CreoTeleport = False
                Call EraseObj(ToMap, 0, UserList(i).flags.DondeTiroMap, MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).OBJInfo.Amount, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY)
                MapData(mapa, X, Y).TileExit.Map = 0
                MapData(mapa, X, Y).TileExit.X = 0
                MapData(mapa, X, Y).TileExit.Y = 0
                UserList(i).flags.DondeTiroMap = 0
                UserList(i).flags.DondeTiroX = 0
                UserList(i).flags.DondeTiroY = 0
    End If
End If
            
    If UserList(i).flags.ActivoGema Then
            UserList(i).flags.TimeGema = UserList(i).flags.TimeGema - 1
            
        If UserList(i).flags.TimeGema <= 0 Then
                UserList(i).flags.ActivoGema = 0
                UserList(i).flags.GemaActivada = ""
                UserList(i).flags.TimeGema = 0
                SendData SendTarget.toindex, i, 0, "||88"
        End If
        
            'Gema roja aumenta vida
            If UserList(i).flags.GemaActivada = "Roja" Then
            
            UserList(i).flags.SubeVidaG = UserList(i).flags.SubeVidaG + 1
            
                If UserList(i).flags.SubeVidaG = 15 And UserList(i).flags.Muerto = 0 And UserList(i).Stats.MinHP < UserList(i).Stats.MaxHP Then
                    UserList(i).Stats.MinHP = UserList(i).Stats.MinHP + RandomNumber(15, 20)
                    If UserList(i).Stats.MinHP > UserList(i).Stats.MaxHP Then UserList(i).Stats.MinHP = UserList(i).Stats.MaxHP
                    
                    SendUserHP (i)
                    UserList(i).flags.SubeVidaG = 0
                End If
            
            End If
            
            'Gema azul aumenta mana
            If UserList(i).flags.GemaActivada = "Azul" Then
            
            UserList(i).flags.SubeManaG = UserList(i).flags.SubeManaG + 1
            
                If UserList(i).flags.SubeManaG = 15 And UserList(i).flags.Muerto = 0 And UserList(i).Stats.MinMAN < UserList(i).Stats.MaxMAN Then
                    UserList(i).Stats.MinMAN = UserList(i).Stats.MinMAN + RandomNumber(80, 150)
                    If UserList(i).Stats.MinMAN > UserList(i).Stats.MaxMAN Then UserList(i).Stats.MinMAN = UserList(i).Stats.MaxMAN
                    
                      SendUserMP (i)
                      UserList(i).flags.SubeManaG = 0
                 End If
            End If
        
    End If
    
    If UCase$(UserList(i).clase) = "DRUIDA" And UserList(i).flags.EleDeAgua = 1 Then
        Dim EAG As Long
        For EAG = 1 To MAXMASCOTAS
          If UserList(i).MascotasIndex(EAG) > 0 Then
            If Npclist(UserList(i).MascotasIndex(EAG)).Numero = ELEMENTALAGUA And UserList(i).Stats.MinHP < UserList(i).Stats.MaxHP Then
                UserList(i).Counters.TiempoElemental = UserList(i).Counters.TiempoElemental + 1
                If UserList(i).Counters.TiempoElemental = 3 Then
                    If InMapBounds(Npclist(UserList(i).MascotasIndex(EAG)).Pos.Map, Npclist(UserList(i).MascotasIndex(EAG)).Pos.X, Npclist(UserList(i).MascotasIndex(EAG)).Pos.Y) Then Call NpcLanzaSpellSobreUser(UserList(i).MascotasIndex(EAG), i, 59)
                    UserList(i).Counters.TiempoElemental = 0
                End If
            End If
          End If
        Next EAG
    End If
            
    If UserList(i).flags.IntervaloBurbu > 1 Then
          UserList(i).flags.IntervaloBurbu = UserList(i).flags.IntervaloBurbu - 1
    ElseIf UserList(i).flags.IntervaloBurbu = 1 Then
          UserList(i).flags.IntervaloBurbu = 0
          UserList(i).flags.DefensaBurbu = 0
          SendData SendTarget.toindex, i, 0, "||81"
    End If

        If UserList(i).flags.Muerto Then
            If UserList(i).flags.TimeRevivir > 0 Then
                UserList(i).flags.TimeRevivir = UserList(i).flags.TimeRevivir - 1
            End If
        End If
        
        If UserList(i).Counters.SegundosParaRevivir > 0 Then
            UserList(i).Counters.SegundosParaRevivir = UserList(i).Counters.SegundosParaRevivir - 1
        
            If UserList(i).Counters.SegundosParaRevivir = 0 Then
                Call RevivirUsuario(i)
            End If
        End If
        
End If
        
    Next i
    
End Sub
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")
    Call SendData(SendTarget.ToAll, 0, 0, "||687")
    
    Dim i As Integer
    'Guardamos los personajes
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, "||688")
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")

    haciendoBK = False
End Sub
Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String
 
Cifra = str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function
Sub LimpiarMundoEntero()
Call SendData(SendTarget.ToAll, 0, 0, "||689")

Dim i As Long

'Borramos todo a la pija ya que lo pidio el gm.
For i = 1 To MAX_OBJS_CLEAR
    'Si el indice tiene un objeto
    If tClearWorld(i).Map <> 0 And tClearWorld(i).X <> 0 And tClearWorld(i).Y <> 0 Then
        With tClearWorld(i)
            If MapData(.Map, .X, .Y).OBJInfo.ObjIndex > 0 Then
                Call EraseObj(ToMap, 0, .Map, 10000, .Map, .X, .Y)
                .Map = 0
                .X = 0
                .Y = 0
                .Tiempo = 0
                .ObjIndex = 0
            End If
        End With
    End If
Next i
 
Call SendData(SendTarget.ToAll, 0, 0, "||690")
End Sub
Sub LimpiarMapa(ByVal MapaActual As Long)

Dim i As Long

'Borramos todo a la pija ya que lo pidio el gm.
For i = 1 To MAX_OBJS_CLEAR
    'Si el indice tiene un objeto
    If tClearWorld(i).Map <> 0 And tClearWorld(i).X <> 0 And tClearWorld(i).Y <> 0 Then
        With tClearWorld(i)
            If MapData(.Map, .X, .Y).OBJInfo.ObjIndex > 0 And .Map = MapaActual Then
                Call EraseObj(ToMap, 0, .Map, 10000, .Map, .X, .Y)
                .Map = 0
                .X = 0
                .Y = 0
                .Tiempo = 0
                .ObjIndex = 0
            End If
        End With
    End If
Next i

End Sub
Public Function codex() As String
    Dim i As Long
    codex = ""
        For i = 1 To 15
            If RandomNumber(1, 2) = 1 Then 'Mayuscula
                codex = codex & Chr(RandomNumber(65, 90))
            Else 'Minuscula
                codex = codex & Chr(RandomNumber(97, 122))
            End If
        Next i
End Function
Sub ControlarPortalLum(ByVal userindex As Integer)
  
    If UserList(userindex).Counters.CreoTeleport = True Then
        Call EraseObj(ToMap, 0, UserList(userindex).flags.DondeTiroMap, MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).OBJInfo.Amount, UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY)
        MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).TileExit.Map = 0
        MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).TileExit.X = 0
        MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).TileExit.Y = 0
        UserList(userindex).flags.DondeTiroMap = 0
        UserList(userindex).flags.DondeTiroX = 0
        UserList(userindex).flags.DondeTiroY = 0
    End If
    
End Sub
Public Sub FriendConnect(userindex As Integer, NombreAmigo As String)

Dim forsitoh As Integer
Dim tStr As String
    For forsitoh = 1 To 20
    
    If UCase$(UserList(userindex).flags.NombreAmigo(forsitoh)) = UCase$(NombreAmigo) Then
        If NombreAmigo = "" Or NombreAmigo = " " Then Exit Sub
        Call SendData(SendTarget.toindex, userindex, 0, "KFM" & NombreAmigo)
        tStr = SendFriendList(userindex)
        Call SendData(SendTarget.toindex, userindex, 0, "LDM" & SendFriendList(userindex))
     Exit Sub
    End If
Next forsitoh

End Sub
Public Sub FriendDisconnect(userindex As Integer, NombreAmigo As String)

Dim forsitoh As Integer
Dim tStr As String
 For forsitoh = 1 To 20
    If UCase$(UserList(userindex).flags.NombreAmigo(forsitoh)) = UCase$(NombreAmigo) Then
        If NombreAmigo = "" Or NombreAmigo = " " Then Exit Sub
            Call SendData(SendTarget.toindex, userindex, 0, "DFM" & NombreAmigo)
            tStr = SendFriendList(userindex, NombreAmigo)
            Call SendData(SendTarget.toindex, userindex, 0, "LDM" & SendFriendList(userindex, NombreAmigo))
     Exit Sub
    End If
Next forsitoh

End Sub
Public Function SendNobleList1(ByVal userindex As Integer) As String
Dim tStr1 As String
Dim tIntx As Integer


    tStr1 = GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "ItemsRequeridos") & ","
For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "ItemsRequeridos")
    tStr1 = tStr1 & ObjData(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "Obj" & tIntx), 45))).Name & ","
    tStr1 = tStr1 & val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "Obj" & tIntx), 45)) & ","
Next tIntx


    SendNobleList1 = tStr1
End Function
Public Function SendNobleList2(ByVal userindex As Integer) As String
Dim tStr2 As String
Dim tIntx As Integer


    tStr2 = GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "ItemsRequeridos") & ","
For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "ItemsRequeridos")
    tStr2 = tStr2 & ObjData(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "Obj" & tIntx), 45))).Name & ","
    tStr2 = tStr2 & val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "Obj" & tIntx), 45)) & ","
Next tIntx


    SendNobleList2 = tStr2
End Function
Public Function SendNobleList3(ByVal userindex As Integer) As String
Dim tStr3 As String
Dim tIntx As Integer


    tStr3 = GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "ItemsRequeridos") & ","
For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "ItemsRequeridos")
    tStr3 = tStr3 & ObjData(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "Obj" & tIntx), 45))).Name & ","
    tStr3 = tStr3 & val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "Obj" & tIntx), 45)) & ","
Next tIntx


    SendNobleList3 = tStr3
End Function
Public Function SendNobleList4(ByVal userindex As Integer) As String
Dim tStr4 As String
Dim tIntx As Integer


    tStr4 = GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "ItemsRequeridos") & ","
For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "ItemsRequeridos")
    tStr4 = tStr4 & ObjData(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "Obj" & tIntx), 45))).Name & ","
    tStr4 = tStr4 & val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "Obj" & tIntx), 45)) & ","
Next tIntx


    SendNobleList4 = tStr4
End Function
Public Function SendTorneoList(ByVal userindex As Integer) As String
Dim tStr As String
Dim tIntx As Integer
 
    tStr = UsuariosEnTorneo & ","
    For tIntx = 1 To LastUser
      If UserList(tIntx).flags.NumTorneo > 0 Then
        CronologiaParticipantesList(UserList(tIntx).flags.NumTorneo) = UserList(tIntx).Name
      End If
    Next tIntx
    
    For tIntx = 1 To UsuariosEnTorneo
        tStr = tStr & CronologiaParticipantesList(tIntx) & ","
    Next tIntx
    
    SendTorneoList = tStr
    
End Function
Public Sub RevisarDuelo(ByVal userindex As Integer)

        Dim uDuelo1     As Integer
        Dim uDuelo2     As Integer
        Dim especterr As Long

     If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 1 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        
        If MisionesDiarias(UserList(uDuelo2).Misiones.NumeroMision).Tipo = 3 Then
            If UserList(uDuelo2).Misiones.ConteoUser < MisionesDiarias(UserList(uDuelo2).Misiones.NumeroMision).Cantidad Then UserList(uDuelo2).Misiones.ConteoUser = UserList(uDuelo2).Misiones.ConteoUser + 1
        End If
        
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||691@1@" & UserList(uDuelo2).Name & "@" & UserList(uDuelo1).Name & "@" & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo1).flags.ApuestaOro) / 2))
        Call LogDuelos("Arena 1>> " & UserList(uDuelo2).Name & " venció en duelo a " & UserList(uDuelo1).Name & " por " & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo2).flags.ApuestaOro) / 2)) & " monedas de oro.")
        UserList(uDuelo2).Stats.GLD = UserList(uDuelo2).Stats.GLD + (val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo2).flags.ApuestaOro) / 2)
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        UserList(uDuelo2).Stats.DuelosGanados = UserList(uDuelo2).Stats.DuelosGanados + 1
        UserList(uDuelo1).Stats.DuelosPerdidos = UserList(uDuelo1).Stats.DuelosPerdidos + 1
        Call GRANK_User_Check(Duels, UserList(uDuelo2).Name, UserList(uDuelo2).Stats.DuelosGanados)
        TiempoDuelo(1) = 0
        NombreDueleando(1) = ""
        NombreDueleando(2) = ""
        ArenaOcupada(1) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena1 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena1 = 0
                EspectadoresEnArena1 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
     If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 2 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador

        If UserList(uDuelo2).Misiones.Tipo = 3 Then
            If UserList(uDuelo2).Misiones.ConteoUser < UserList(uDuelo2).Misiones.Cantidad Then UserList(uDuelo2).Misiones.ConteoUser = UserList(uDuelo2).Misiones.ConteoUser + 1
        End If
        
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||691@2@" & UserList(uDuelo2).Name & "@" & UserList(uDuelo1).Name & "@" & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo1).flags.ApuestaOro) / 2))
        Call LogDuelos("Arena 2>> " & UserList(uDuelo2).Name & " venció en duelo a " & UserList(uDuelo1).Name & " por " & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo2).flags.ApuestaOro) / 2)) & " monedas de oro.")
        UserList(uDuelo2).Stats.GLD = UserList(uDuelo2).Stats.GLD + (val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo2).flags.ApuestaOro) / 2)
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        UserList(uDuelo2).Stats.DuelosGanados = UserList(uDuelo2).Stats.DuelosGanados + 1
        UserList(uDuelo1).Stats.DuelosPerdidos = UserList(uDuelo1).Stats.DuelosPerdidos + 1
        Call GRANK_User_Check(Duels, UserList(uDuelo2).Name, UserList(uDuelo2).Stats.DuelosGanados)
        TiempoDuelo(2) = 0
        NombreDueleando(3) = ""
        NombreDueleando(4) = ""
        ArenaOcupada(2) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena2 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena2 = 0
                EspectadoresEnArena2 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
     If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 3 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        
        If MisionesDiarias(UserList(uDuelo2).Misiones.NumeroMision).Tipo = 3 Then
            If UserList(uDuelo2).Misiones.ConteoUser < MisionesDiarias(UserList(uDuelo2).Misiones.NumeroMision).Cantidad Then UserList(uDuelo2).Misiones.ConteoUser = UserList(uDuelo2).Misiones.ConteoUser + 1
        End If
        
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||691@3@" & UserList(uDuelo2).Name & "@" & UserList(uDuelo1).Name & "@" & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo1).flags.ApuestaOro) / 2))
        Call LogDuelos("Arena 3>> " & UserList(uDuelo2).Name & " venció en duelo a " & UserList(uDuelo1).Name & " por " & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo1).flags.ApuestaOro) / 2)) & " monedas de oro.")
        UserList(uDuelo2).Stats.GLD = UserList(uDuelo2).Stats.GLD + (val(dMoney) + val(dMoney) / 2)
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        UserList(uDuelo2).Stats.DuelosGanados = UserList(uDuelo2).Stats.DuelosGanados + 1
        UserList(uDuelo1).Stats.DuelosPerdidos = UserList(uDuelo1).Stats.DuelosPerdidos + 1
        Call GRANK_User_Check(Duels, UserList(uDuelo2).Name, UserList(uDuelo2).Stats.DuelosGanados)
        TiempoDuelo(3) = 0
        NombreDueleando(5) = ""
        NombreDueleando(6) = ""
        ArenaOcupada(3) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena3 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena3 = 0
                EspectadoresEnArena3 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
     If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 4 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        
        If MisionesDiarias(UserList(uDuelo2).Misiones.NumeroMision).Tipo = 3 Then
            If UserList(uDuelo2).Misiones.ConteoUser < MisionesDiarias(UserList(uDuelo2).Misiones.NumeroMision).Cantidad Then UserList(uDuelo2).Misiones.ConteoUser = UserList(uDuelo2).Misiones.ConteoUser + 1
        End If
        
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||691@4@" & UserList(uDuelo2).Name & "@" & UserList(uDuelo1).Name & "@" & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo1).flags.ApuestaOro) / 2))
        Call LogDuelos("Arena 4>> " & UserList(uDuelo2).Name & " venció en duelo a " & UserList(uDuelo1).Name & " por " & PonerPuntos((val(UserList(uDuelo1).flags.ApuestaOro) + val(UserList(uDuelo1).flags.ApuestaOro) / 2)) & " monedas de oro.")
        UserList(uDuelo2).Stats.GLD = UserList(uDuelo2).Stats.GLD + (val(dMoney) + val(dMoney) / 2)
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        UserList(uDuelo2).Stats.DuelosGanados = UserList(uDuelo2).Stats.DuelosGanados + 1
        UserList(uDuelo1).Stats.DuelosPerdidos = UserList(uDuelo1).Stats.DuelosPerdidos + 1
        Call GRANK_User_Check(Duels, UserList(uDuelo2).Name, UserList(uDuelo2).Stats.DuelosGanados)
        TiempoDuelo(4) = 0
        NombreDueleando(7) = ""
        NombreDueleando(8) = ""
        ArenaOcupada(4) = False
                
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena4 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena4 = 0
                EspectadoresEnArena4 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
    
End Sub
Public Sub SalirDuelo(ByVal userindex As Integer)

        Dim uDuelo1     As Integer
        Dim uDuelo2     As Integer
        Dim especterr As Long

      If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 1 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||692@1@" & UserList(uDuelo1).Name
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        TiempoDuelo(1) = 0
        NombreDueleando(1) = ""
        NombreDueleando(2) = ""
        ArenaOcupada(1) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena1 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena1 = 0
                EspectadoresEnArena1 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
      If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 2 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||692@2@" & UserList(uDuelo1).Name
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        TiempoDuelo(2) = 0
        NombreDueleando(3) = ""
        NombreDueleando(4) = ""
        ArenaOcupada(2) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena2 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena2 = 0
                EspectadoresEnArena2 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
      If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 3 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||692@3@" & UserList(uDuelo1).Name
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        TiempoDuelo(3) = 0
        NombreDueleando(5) = ""
        NombreDueleando(6) = ""
        ArenaOcupada(3) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena3 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena3 = 0
                EspectadoresEnArena3 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
      If UserList(userindex).flags.EnDuelo = True And UserList(userindex).flags.EnQueArena = 4 Then
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        UserList(uDuelo1).flags.EnQueArena = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        UserList(uDuelo2).flags.EnQueArena = 0
        'Set Usuario Ganador
        'Set Todo
        SendData SendTarget.ToAll, userindex, 0, "||692@4@" & UserList(uDuelo1).Name
        WarpUserChar uDuelo1, UserList(uDuelo1).flags.MapaAnterior, UserList(uDuelo1).flags.XAnterior, UserList(uDuelo1).flags.YAnterior, True
        WarpUserChar uDuelo2, UserList(uDuelo2).flags.MapaAnterior, UserList(uDuelo2).flags.XAnterior, UserList(uDuelo2).flags.YAnterior, True
        TiempoDuelo(4) = 0
        NombreDueleando(7) = ""
        NombreDueleando(8) = ""
        ArenaOcupada(4) = False
        
        For especterr = 1 To LastUser
            If UserList(especterr).flags.EspectadorArena4 = 1 Then
                WarpUserChar especterr, UserList(especterr).flags.MapaAnterior, UserList(especterr).flags.XAnterior, UserList(especterr).flags.YAnterior, True
                UserList(especterr).flags.EspectadorArena4 = 0
                EspectadoresEnArena4 = 0
            End If
        Next especterr
        
        SendUserGLD (uDuelo1)
        SendUserGLD (uDuelo2)
        
    End If
    
End Sub
Public Function DamePos(ByRef original_Pos As WorldPos) As WorldPos
 
'
' @ Devuelve un tile libre.
 
Dim iRange      As Long
Dim iX          As Long
Dim iY          As Long
Dim now_Index   As Integer
Dim no_User     As Boolean
Dim not_Pos     As WorldPos
 
not_Pos = original_Pos
DamePos.Map = original_Pos.Map
 
With original_Pos
     For iRange = 1 To 5
         For iX = (.X - iRange) To (.X + iRange)
             For iY = (.Y - iRange) To (.Y + iRange)
                 
                 now_Index = MapData(.Map, iX, iY).userindex
                 
                 'No hay n usuario
                 If (now_Index = 0) And MapData(.Map, iX, iY).Blocked = 0 Then
                    DamePos.X = iX
                    DamePos.Y = iY
                    no_User = True
                 End If
                 
                 'No hay usuario, revisa npc
                 If (no_User = True) Then
                    now_Index = MapData(.Map, iX, iY).NpcIndex
                   
                    'No hay un npc.
                    If (now_Index = 0) Then
                       DamePos.X = iX
                       DamePos.Y = iY
                       Exit Function
                    Else
                       no_User = False
                    End If
                 End If
 
             Next iY
         Next iX
     Next iRange
End With
 
'Llega acá, devuelve la posición original.
DamePos = not_Pos
 
End Function
Public Sub Desmontar(userindex As Integer)
    UserList(userindex).flags.Montando = 0
    UserList(userindex).flags.InvocoMascota = 0
        UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userindex)
        End If
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
 
Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "USM" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).flags.Montando)
Call SendData(toindex, userindex, 0, "EQUIT")
End Sub
Public Sub CambiarNickMascota(userindex As Integer, NickNuevo As String)
UserList(userindex).NickMascota = NickNuevo
End Sub
Function ZonaCura(ByVal userindex As Integer) As Boolean
' Autor: Joan Calderón - SaturoS.
'Codigo: Sacerdotes automaticos.
Dim X As Integer, Y As Integer
For Y = UserList(userindex).Pos.Y - MinYBorder + 1 To UserList(userindex).Pos.Y + MinYBorder - 1
        For X = UserList(userindex).Pos.X - MinXBorder + 1 To UserList(userindex).Pos.X + MinXBorder - 1
       
            If MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex).NPCtype = 1 Then
                    If Distancia(UserList(userindex).Pos, Npclist(MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex).Pos) < 20 Then
                        ZonaCura = True
                        Exit Function
                    End If
                End If
            End If
           
        Next X
Next Y
ZonaCura = False
End Function
Sub AutoCuraUser(ByVal userindex As Integer)
' Autor: Joan Calderón - SaturoS.
'Codigo: Sacerdotes automaticos.

If EstaEnRing(userindex) Then
    Call SendData(toindex, userindex, 0, "||395")
    Exit Sub
End If

If UserList(userindex).flags.Muerto = 1 Then
    Call RevivirUsuario(userindex)
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
    Call SendData(toindex, userindex, 0, "||693")
    Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "TW20") ' este es el sonido cuando cura o resucita al personaje
    
    Call SendUserHP(userindex)
    Call SendUserST(userindex)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 65 & "," & 0)
End If
 
If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    Call SendData(toindex, userindex, 0, "||694")
    Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "TW20") ' este es el sonido de cuando resucita o cura al personaje.
    Call SendUserHP(userindex)
End If
 
If UserList(userindex).flags.Envenenado = 1 Then UserList(userindex).flags.Envenenado = 0
 
 
End Sub
Public Sub SwapObjects(ByVal userindex As Integer)
Dim tmpUserObj As UserOBJ
 
    If ObjSlot1 > UserList(userindex).InventorySlots Or ObjSlot2 > UserList(userindex).InventorySlots Then Exit Sub
 
    With UserList(userindex)
               
        'Cambiamos si alguno es una herramienta
        If .Invent.HerramientaEqpSlot = ObjSlot1 Then
            .Invent.HerramientaEqpSlot = ObjSlot2
        ElseIf .Invent.HerramientaEqpSlot = ObjSlot2 Then
            .Invent.HerramientaEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un armor
        If .Invent.ArmourEqpSlot = ObjSlot1 Then
            .Invent.ArmourEqpSlot = ObjSlot2
        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
            .Invent.ArmourEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un barco
        If .Invent.BarcoSlot = ObjSlot1 Then
            .Invent.BarcoSlot = ObjSlot2
        ElseIf .Invent.BarcoSlot = ObjSlot2 Then
            .Invent.BarcoSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un casco
        If .Invent.CascoEqpSlot = ObjSlot1 Then
            .Invent.CascoEqpSlot = ObjSlot2
        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
            .Invent.CascoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un escudo
        If .Invent.EscudoEqpSlot = ObjSlot1 Then
            .Invent.EscudoEqpSlot = ObjSlot2
        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
            .Invent.EscudoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es munición
        If .Invent.MunicionEqpSlot = ObjSlot1 Then
            .Invent.MunicionEqpSlot = ObjSlot2
        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
            .Invent.MunicionEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un arma
        If .Invent.WeaponEqpSlot = ObjSlot1 Then
            .Invent.WeaponEqpSlot = ObjSlot2
        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
            .Invent.WeaponEqpSlot = ObjSlot1
        End If
       
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, userindex, ObjSlot1)
        Call UpdateUserInv(False, userindex, ObjSlot2)

        Call Desequipar(userindex, ObjSlot1)
        Call Desequipar(userindex, ObjSlot2)
    End With
End Sub
Function EsHorda(ByVal index As Integer) As Boolean

If UserList(index).StatusMith.EsStatus = 2 Or UserList(index).StatusMith.EsStatus = 4 Or UserList(index).StatusMith.EsStatus = 6 Then
    EsHorda = True
    Exit Function
Else
    EsHorda = False
End If


End Function
Function EsAlianza(ByVal index As Integer) As Boolean

If UserList(index).StatusMith.EsStatus = 1 Or UserList(index).StatusMith.EsStatus = 3 Or UserList(index).StatusMith.EsStatus = 5 Then
    EsAlianza = True
    Exit Function
Else
    EsAlianza = False
End If


End Function
Public Sub RevivirMapaUser(ByVal userindex As Integer)
Dim RevivirMap As Integer
For RevivirMap = 1 To LastUser
If UserList(RevivirMap).Pos.Map = UserList(userindex).Pos.Map Then
If UserList(RevivirMap).flags.Muerto = 1 Then
Call RevivirUsuario(RevivirMap)
Call DarCuerpoDesnudo(RevivirMap)
End If
End If
Next RevivirMap
Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "||695")
End Sub
Public Sub RevivirMapa(ByVal userindex As Integer, rData As String)
Dim tStr As String
Dim RevivirMap As Integer
For RevivirMap = 1 To LastUser
If UserList(RevivirMap).Pos.Map = rData Then
tStr = UserList(RevivirMap).Pos.Map
If UserList(RevivirMap).flags.Muerto = 1 Then
Call RevivirUsuario(RevivirMap)
Call DarCuerpoDesnudo(RevivirMap)
End If
End If
Next RevivirMap
Call SendData(SendTarget.ToMap, 0, tStr, "||696")
End Sub
Public Sub PortalesDeDioses(Dios As String)

Select Case UCase$(Dios)
    Case "TARRASKE"
                MapData(176, 49, 18).TileExit.Map = 181
                MapData(176, 49, 18).TileExit.X = 51
                MapData(176, 49, 18).TileExit.Y = 65
                
                MapData(176, 50, 18).TileExit.Map = 181
                MapData(176, 50, 18).TileExit.X = 51
                MapData(176, 50, 18).TileExit.Y = 65
                
                MapData(176, 51, 18).TileExit.Map = 181
                MapData(176, 51, 18).TileExit.X = 51
                MapData(176, 51, 18).TileExit.Y = 65
                
                MapData(176, 52, 18).TileExit.Map = 181
                MapData(176, 52, 18).TileExit.X = 51
                MapData(176, 52, 18).TileExit.Y = 65
                
                MapData(176, 53, 18).TileExit.Map = 181
                MapData(176, 53, 18).TileExit.X = 51
                MapData(176, 53, 18).TileExit.Y = 65
                
                MapData(176, 54, 18).TileExit.Map = 181
                MapData(176, 54, 18).TileExit.X = 51
                MapData(176, 54, 18).TileExit.Y = 65
                
                MapData(176, 55, 18).TileExit.Map = 181
                MapData(176, 55, 18).TileExit.X = 51
                MapData(176, 55, 18).TileExit.Y = 65
        
        
    Case "MIFRIT"
                MapData(177, 46, 23).TileExit.Map = 180
                MapData(177, 46, 23).TileExit.X = 50
                MapData(177, 46, 23).TileExit.Y = 61
                
                MapData(177, 47, 23).TileExit.Map = 180
                MapData(177, 47, 23).TileExit.X = 50
                MapData(177, 47, 23).TileExit.Y = 61
                
                MapData(177, 48, 23).TileExit.Map = 180
                MapData(177, 48, 23).TileExit.X = 50
                MapData(177, 48, 23).TileExit.Y = 61
                
                MapData(177, 49, 23).TileExit.Map = 180
                MapData(177, 49, 23).TileExit.X = 50
                MapData(177, 49, 23).TileExit.Y = 61
                
                MapData(177, 50, 23).TileExit.Map = 180
                MapData(177, 50, 23).TileExit.X = 50
                MapData(177, 50, 23).TileExit.Y = 61
                
                MapData(177, 51, 23).TileExit.Map = 180
                MapData(177, 51, 23).TileExit.X = 50
                MapData(177, 51, 23).TileExit.Y = 61
                
                MapData(177, 52, 23).TileExit.Map = 180
                MapData(177, 52, 23).TileExit.X = 50
                MapData(177, 52, 23).TileExit.Y = 61
        
    Case "POSEIDON"
                MapData(159, 49, 50).TileExit.Map = 160
                MapData(159, 49, 50).TileExit.X = 50
                MapData(159, 49, 50).TileExit.Y = 65
                
                MapData(159, 50, 50).TileExit.Map = 160
                MapData(159, 50, 50).TileExit.X = 50
                MapData(159, 50, 50).TileExit.Y = 65
                
                MapData(159, 51, 50).TileExit.Map = 160
                MapData(159, 51, 50).TileExit.X = 50
                MapData(159, 51, 50).TileExit.Y = 65
                
                MapData(159, 52, 50).TileExit.Map = 160
                MapData(159, 52, 50).TileExit.X = 50
                MapData(159, 52, 50).TileExit.Y = 65
                
                MapData(159, 53, 50).TileExit.Map = 160
                MapData(159, 53, 50).TileExit.X = 50
                MapData(159, 53, 50).TileExit.Y = 65
                
                MapData(159, 54, 50).TileExit.Map = 160
                MapData(159, 54, 50).TileExit.X = 50
                MapData(159, 54, 50).TileExit.Y = 65
                
                MapData(159, 55, 50).TileExit.Map = 160
                MapData(159, 55, 50).TileExit.X = 50
                MapData(159, 55, 50).TileExit.Y = 65
        
    Case "EREBROS"
                MapData(171, 48, 36).TileExit.Map = 170
                MapData(171, 48, 36).TileExit.X = 50
                MapData(171, 48, 36).TileExit.Y = 86

                MapData(171, 49, 36).TileExit.Map = 170
                MapData(171, 49, 36).TileExit.X = 50
                MapData(171, 49, 36).TileExit.Y = 86
                
                MapData(171, 50, 36).TileExit.Map = 170
                MapData(171, 50, 36).TileExit.X = 50
                MapData(171, 50, 36).TileExit.Y = 86
                
                MapData(171, 51, 36).TileExit.Map = 170
                MapData(171, 51, 36).TileExit.X = 50
                MapData(171, 51, 36).TileExit.Y = 86
                
                MapData(171, 52, 36).TileExit.Map = 170
                MapData(171, 52, 36).TileExit.X = 50
                MapData(171, 52, 36).TileExit.Y = 86
                
                MapData(171, 53, 36).TileExit.Map = 170
                MapData(171, 53, 36).TileExit.X = 50
                MapData(171, 53, 36).TileExit.Y = 86
                
                MapData(171, 54, 36).TileExit.Map = 170
                MapData(171, 54, 36).TileExit.X = 50
                MapData(171, 54, 36).TileExit.Y = 86
        
    Case "CERRARTPS"
    Dim kk As Integer
        For kk = 49 To 55
                MapData(176, kk, 18).TileExit.Map = 0
                MapData(176, kk, 18).TileExit.X = 0
                MapData(176, kk, 18).TileExit.Y = 0
        Next kk
        
        For kk = 46 To 52
                MapData(177, kk, 23).TileExit.Map = 0
                MapData(177, kk, 23).TileExit.X = 0
                MapData(177, kk, 23).TileExit.Y = 0
        Next kk
        
        For kk = 49 To 55
                MapData(159, kk, 50).TileExit.Map = 0
                MapData(159, kk, 50).TileExit.X = 0
                MapData(159, kk, 50).TileExit.Y = 0
        Next kk
        
        For kk = 48 To 54
                MapData(171, kk, 36).TileExit.Map = 0
                MapData(171, kk, 36).TileExit.X = 0
                MapData(171, kk, 36).TileExit.Y = 0
        Next kk
        
    End Select
End Sub
Public Function Generar(Upper As Integer, _
                              Optional Lower As Integer = 1, _
                              Optional Cantidad As Integer = 1) As Variant
  
    On Error GoTo Error_Function
    ' verifica que el valor del máximo no sea inferior a la _
      cantidad de números que se generarán
    If Cantidad > ((Upper + 1) - (Lower - 1)) Then
        Exit Function
    End If
      
    Dim X           As Integer
    Dim n           As Integer
    Dim arrNums()   As Variant ' array temporal con los números
    Dim colNumbers  As New Collection
      
    ReDim arrNums(Cantidad - 1)
    With colNumbers
        For X = Lower To Upper
            .Add X
        Next X
        For X = 0 To Cantidad - 1
            ' devuelve el número aleatorio
            n = RandomFERNumber(0, colNumbers.Count + 1)
            arrNums(X) = colNumbers(n)
            colNumbers.Remove n
        Next X
    End With
    Set colNumbers = Nothing
    ' devuelve los números a la función
    Generar = arrNums
Exit Function
Error_Function:
    Generar = ""
End Function
  
' genera el valor aleatorio
''''''''''''''''''''''''''''
Public Function RandomFERNumber(Upper As Integer, Lower As Integer) As Integer
    'Generates a Random Number BETWEEN the LOWER and UPPER values
    Randomize
    RandomFERNumber = Int((Upper - Lower + 1) * Rnd + Lower)
End Function
Public Function MapaEspecial(ByVal userindex As Integer)
    With UserList(userindex).Pos
        MapaEspecial = .Map = 101 Or .Map = 54 Or .Map = 8 Or .Map = 72 Or .Map = 78 Or .Map = 141 Or .Map = 100 Or .Map = 99 Or .Map = 18 Or .Map = 19
    End With
End Function

