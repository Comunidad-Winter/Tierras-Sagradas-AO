Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.9.0.2
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio, Jonatan Ezequiel Salguero
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
On Error Resume Next

Dim DaExp As Integer
Dim DaPT As Integer

DaExp = CInt(UserList(VictimIndex).Stats.ELV)

    'Lo mata
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||60@" & UserList(VictimIndex).Name & "@" & DaExp)
    Call SendData(SendTarget.toindex, VictimIndex, 0, "||61@" & UserList(AttackerIndex).Name)

If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then

            If VictimIndex = GranPoder Then
                Call OtorgarGranPoder(AttackerIndex)
                UserList(VictimIndex).flags.GranPoder = 0
                SendUserVariant (VictimIndex)
            End If

If UserList(AttackerIndex).Pos.Map <> 71 And UserList(AttackerIndex).Pos.Map <> 109 And UserList(AttackerIndex).Pos.Map <> 108 And UserList(AttackerIndex).Pos.Map <> 100 And UserList(AttackerIndex).Pos.Map <> 106 And UserList(AttackerIndex).Pos.Map <> 108 And UCase$(UserList(AttackerIndex).flags.UltimoMatado) <> UCase$(UserList(VictimIndex).Name) Then
    If UserList(AttackerIndex).Stats.UsuariosMatados < 32000 Then _
        UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1
        UserList(AttackerIndex).flags.UltimoMatado = UserList(AttackerIndex).Name
        Call GRANK_User_Check(Kills, UserList(AttackerIndex).Name, UserList(AttackerIndex).Stats.UsuariosMatados)
    End If
End If
    
'desafio
If Desafio.Primero = AttackerIndex And Desafio.Segundo = VictimIndex Then
    Call SendData(SendTarget.ToAll, 0, 0, "||62@" & UserList(AttackerIndex).Name & "@" & UserList(VictimIndex).Name)
    UserList(AttackerIndex).flags.rondas = UserList(AttackerIndex).flags.rondas + 1
    
    Call LogDesafios("" & UserList(AttackerIndex).Name & " derroto a " & UserList(VictimIndex).Name & " y lleva " & UserList(AttackerIndex).flags.rondas & " rondas.")
    
    If UserList(AttackerIndex).flags.rondas > UserList(AttackerIndex).Stats.MaximasRondas Then
        UserList(AttackerIndex).Stats.MaximasRondas = UserList(AttackerIndex).flags.rondas
    End If
    
    Call GRANK_User_Check(Rounds, UserList(AttackerIndex).Name, UserList(AttackerIndex).Stats.MaximasRondas)
    UserList(AttackerIndex).Stats.GLD = UserList(AttackerIndex).Stats.GLD + 30000
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||63@30.000")
    UserList(AttackerIndex).Stats.MinHP = UserList(AttackerIndex).Stats.MaxHP
    UserList(AttackerIndex).Stats.MinMAN = UserList(AttackerIndex).Stats.MaxMAN
    
        SendUserHP (AttackerIndex)
        SendUserMP (AttackerIndex)
        SendUserGLD (AttackerIndex)
    
        If UCase$(UserList(VictimIndex).Genero) = "MUJER" Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, VictimIndex, UserList(VictimIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, VictimIndex, UserList(VictimIndex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
        End If
    
    Call WarpUserChar(Desafio.Segundo, 1, 54, 36, True) 'Poner el mapa en donde salen
    UserList(VictimIndex).flags.EnDesafio = 0
    UserList(VictimIndex).flags.Desafio = 0
    Desafio.Segundo = 0
    
    If UserList(AttackerIndex).flags.rondas = 5 Or UserList(AttackerIndex).flags.rondas = 10 Or UserList(AttackerIndex).flags.rondas = 15 Or UserList(AttackerIndex).flags.rondas = 20 Or UserList(AttackerIndex).flags.rondas = 25 Or UserList(AttackerIndex).flags.rondas = 30 Or UserList(AttackerIndex).flags.rondas = 35 Or UserList(AttackerIndex).flags.rondas = 40 Or UserList(AttackerIndex).flags.rondas = 45 Or UserList(AttackerIndex).flags.rondas = 50 Or UserList(AttackerIndex).flags.rondas = 55 Or UserList(AttackerIndex).flags.rondas = 60 Or UserList(AttackerIndex).flags.rondas = 65 Or UserList(AttackerIndex).flags.rondas = 70 Or UserList(AttackerIndex).flags.rondas = 75 Or UserList(AttackerIndex).flags.rondas = 80 Or UserList(AttackerIndex).flags.rondas = 85 Or UserList(AttackerIndex).flags.rondas = 90 Or UserList(AttackerIndex).flags.rondas = 95 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||64@" & UserList(AttackerIndex).Name & "@" & UserList(AttackerIndex).flags.rondas)
    ElseIf UserList(AttackerIndex).flags.rondas >= 100 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||64@" & UserList(AttackerIndex).Name & "@" & UserList(AttackerIndex).flags.rondas)
    End If
    
    If MisionesDiarias(UserList(AttackerIndex).Misiones.NumeroMision).Tipo = 4 Then
        UserList(AttackerIndex).Misiones.ConteoUser = UserList(AttackerIndex).flags.rondas
    End If
    
End If
 
If Desafio.Primero = VictimIndex And AttackerIndex = Desafio.Segundo Then
    Call SendData(SendTarget.ToAll, 0, 0, "||65@" & UserList(AttackerIndex).Name & "@" & UserList(VictimIndex).Name)
    Call WarpUserChar(AttackerIndex, 1, 54, 36, True)
    Call WarpUserChar(Desafio.Primero, 1, 54, 36, True)
    UserList(AttackerIndex).Stats.GLD = UserList(AttackerIndex).Stats.GLD + 100000
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||63@100.000")
    SendUserGLD (AttackerIndex)
    UserList(AttackerIndex).flags.Desafio = 0
    UserList(VictimIndex).flags.EnDesafio = 0
    UserList(VictimIndex).flags.rondas = 0
    Desafio.Segundo = 0
    Desafio.Primero = 0
End If

If UserList(AttackerIndex).flags.UserNumQuest <> 0 And QuestsList(UserList(AttackerIndex).flags.UserNumQuest).Tipo = 2 Then
    Call modQuests.RestarUser(AttackerIndex, VictimIndex)
End If


    UserList(AttackerIndex).Stats.Reputacione = UserList(AttackerIndex).Stats.Reputacione + 10
    UserList(VictimIndex).Stats.Reputacione = UserList(VictimIndex).Stats.Reputacione - 10

'Log
Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)

End Sub


Sub RevivirUsuario(ByVal userindex As Integer)

UserList(userindex).flags.Muerto = 0
UserList(userindex).Stats.MinHP = 35

If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
End If

Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 65 & "," & 0)
Call DarCuerpoDesnudo(userindex)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendUserHP(userindex)
Call SendUserMP(userindex)

End Sub
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, _
                    ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    UserList(userindex).Char.Body = Body
    UserList(userindex).Char.Head = Head
    UserList(userindex).Char.Heading = Heading
    UserList(userindex).Char.WeaponAnim = Arma
    UserList(userindex).Char.ShieldAnim = Escudo
    UserList(userindex).Char.CascoAnim = Casco
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "CP" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & Casco)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & Casco)
    End If
End Sub
Sub EnviarSubirNivel(ByVal userindex As Integer, ByVal Puntos As Integer)
    Call SendData(SendTarget.toindex, userindex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal userindex As Integer)
    Dim i As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
       cad = cad & UserList(userindex).Stats.UserSkills(i) & ","
    Next i
    
    SendData SendTarget.toindex, userindex, 0, "SKILLS" & cad$
End Sub
Sub EnviarAtrib(ByVal userindex As Integer)
Dim i As Integer
Dim cad As String
For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(userindex).Stats.UserAtributos(i) & ","
Next
Call SendData(SendTarget.toindex, userindex, 0, "ATR" & cad)
End Sub
Sub InfoQuest(ByVal userindex As Integer)
    Call SendData(SendTarget.toindex, userindex, 0, "INK" & UserList(userindex).flags.UserNumQuest & "," & UserList(userindex).flags.MuereQuest(1) & "," & UserList(userindex).flags.MuereQuest(2) & "," & UserList(userindex).flags.MuereQuest(3))
End Sub
Public Sub EnviarMiniEstadisticas(ByVal userindex As Integer)

Dim TotalAdvert As Byte
Dim Emailxsd As String
Dim JerarquiaNum As String
Dim NeedSiguiente As String
Dim Alineacion As Byte
TotalAdvert = val(GetVar(CharPath & UserList(userindex).Name & ".chr", "Advertencias", "Number"))
Emailxsd = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "" & UserList(userindex).Accounted & "", "Mail")

If UserList(userindex).StatusMith.EsStatus = 0 Then
JerarquiaNum = "None"
NeedSiguiente = "None"
Alineacion = 0
End If

If UserList(userindex).flags.PJerarquia = 0 And UserList(userindex).flags.SJerarquia = 0 And UserList(userindex).flags.TJerarquia = 0 And UserList(userindex).flags.CJerarquia = 0 Then
JerarquiaNum = "None"
ElseIf UserList(userindex).flags.PJerarquia = 1 And UserList(userindex).flags.SJerarquia = 0 And UserList(userindex).flags.TJerarquia = 0 And UserList(userindex).flags.CJerarquia = 0 Then
JerarquiaNum = "1 de 4"
ElseIf UserList(userindex).flags.PJerarquia = 0 And UserList(userindex).flags.SJerarquia = 1 And UserList(userindex).flags.TJerarquia = 0 And UserList(userindex).flags.CJerarquia = 0 Then
JerarquiaNum = "2 de 4"
ElseIf UserList(userindex).flags.PJerarquia = 0 And UserList(userindex).flags.SJerarquia = 0 And UserList(userindex).flags.TJerarquia = 1 And UserList(userindex).flags.CJerarquia = 0 Then
JerarquiaNum = "3 de 4"
ElseIf UserList(userindex).flags.PJerarquia = 0 And UserList(userindex).flags.SJerarquia = 0 And UserList(userindex).flags.TJerarquia = 0 And UserList(userindex).flags.CJerarquia = 1 Then
JerarquiaNum = "4 de 4"
End If

If UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Then
If JerarquiaNum = "None" Then
NeedSiguiente = FragsJerarquia(1) - UserList(userindex).Faccion.CriminalesMatados
ElseIf JerarquiaNum = "1 de 4" Then
NeedSiguiente = FragsJerarquia(2) - UserList(userindex).Faccion.CriminalesMatados
ElseIf JerarquiaNum = "2 de 4" Then
NeedSiguiente = FragsJerarquia(3) - UserList(userindex).Faccion.CriminalesMatados
ElseIf JerarquiaNum = "3 de 4" Then
NeedSiguiente = FragsJerarquia(4) - UserList(userindex).Faccion.CriminalesMatados
ElseIf JerarquiaNum = "4 de 4" Then
NeedSiguiente = "None"
End If
Alineacion = 2
End If

If UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Then
If JerarquiaNum = "None" Then
NeedSiguiente = FragsJerarquia(1) - UserList(userindex).Faccion.CiudadanosMatados
ElseIf JerarquiaNum = "1 de 4" Then
NeedSiguiente = FragsJerarquia(2) - UserList(userindex).Faccion.CiudadanosMatados
ElseIf JerarquiaNum = "2 de 4" Then
NeedSiguiente = FragsJerarquia(3) - UserList(userindex).Faccion.CiudadanosMatados
ElseIf JerarquiaNum = "3 de 4" Then
NeedSiguiente = FragsJerarquia(4) - UserList(userindex).Faccion.CiudadanosMatados
ElseIf JerarquiaNum = "4 de 4" Then
NeedSiguiente = "None"
End If
Alineacion = 1
End If

'############### SETEAMOS QUESTS ####################
Dim TipoQuest As Byte
Dim DescQuest As String
Dim PremioOro As Long
Dim PremioPuntis As Integer
Dim CantidadNPCs As Byte
Dim YaMatados As Byte

'If UserList(UserIndex).flags.UserNumQuest = 0 Then
'    TipoQuest = 0
'    DescQuest = 0
'    PremioOro = 0
'    PremioPuntis = 0
'    CantidadNPCs = 0
'    YaMatados = 0
'Else
 '   TipoQuest = GetVar(DatPath & "quests.dat", "Quest" & UserList(UserIndex).flags.UserNumQuest, "Tipo")
 '   DescQuest = GetVar(DatPath & "quests.dat", "Quest" & UserList(UserIndex).flags.UserNumQuest, "Name")
 '   PremioOro = GetVar(DatPath & "quests.dat", "Quest" & UserList(UserIndex).flags.UserNumQuest, "Oro")
'    PremioPuntis = GetVar(DatPath & "quests.dat", "Quest" & UserList(UserIndex).flags.UserNumQuest, "Premio")
'    CantidadNPCs = GetVar(DatPath & "quests.dat", "Quest" & UserList(UserIndex).flags.UserNumQuest, "Cant")
'    'YaMatados = UserList(UserIndex).flags.MuereQuest
'End If
'############### SETEAMOS QUESTS ####################


With UserList(userindex)
Call SendData(SendTarget.toindex, userindex, 0, "KIGF" & .clase & "," & Emailxsd & "," & TotalAdvert & "," & .Stats.DuelosGanados & "," & .Stats.DuelosPerdidos & "," & .Stats.TrofOro & "," & .Stats.TrofPlata & "," & .Stats.TrofBronce & "," & .flags.QuestCompletadas & "," & .Faccion.CiudadanosMatados & "," & .Faccion.CriminalesMatados & "," & .Stats.NPCsMuertos & "," & JerarquiaNum & "," & NeedSiguiente & "," & Alineacion & "," & UserList(userindex).flags.GuerrasGanadas & "," & UserList(userindex).flags.CvcsGanados & "," & UserList(userindex).flags.MVPMatados & "," & UserList(userindex).Stats.PuntosTorneo & "," & UserList(userindex).Hogar & "," & .Genero & "," & .Stats.ELV & "," & .Bon1 & "," & .Bon2 & "," & .Bon3 & "," & .Name & "," & TipoQuest & "," & DescQuest & "," & PremioOro & "," & PremioPuntis & "," & CantidadNPCs & "," & YaMatados & "," & .Stats.TorneosParticipados & "," & .Stats.MaximasRondas & "," & .Stats.MedOro & "," & .Stats.ParejasGanadas & "," & .Stats.ParejasPerdidas & "," _
 & .flags.GuerrasPerdidas & "," & .Faccion.NeutralesMatados & "," & .Stats.MuertesUser & "," & .Raza & "," & .Stats.Reputacione & "," & .Stats.PuntosDonacion)
End With

End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, userindex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(userindex).Char.CharIndex) = 0
    
    If UserList(userindex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén en el mismo mapa
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "BP" & UserList(userindex).Char.CharIndex)
        Call QuitarUser(userindex, UserList(userindex).Pos.Map)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BP" & UserList(userindex).Char.CharIndex)
    End If
    
    MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = 0
    UserList(userindex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
Dim CharIndex As Integer
 
If InMapBounds(Map, X, Y) Then
'If needed make a new character in list
If UserList(userindex).Char.CharIndex = 0 Then
CharIndex = NextOpenCharIndex
UserList(userindex).Char.CharIndex = CharIndex
CharList(CharIndex) = userindex
End If
 
'Place character on map
MapData(Map, X, Y).userindex = userindex
 
'Send make character command to clients
Dim klan As String
If UserList(userindex).GuildIndex > 0 Then
klan = Guilds(UserList(userindex).GuildIndex).GuildName
End If
 
Dim bCr As Byte
Dim SendPrivilegios As Byte
 
bCr = Criminal(userindex)

    SendUserVariant (userindex)
    SendUserAura (userindex)
    'SendUserRank (userindex)
        
    If klan <> "" Then
        If sndRoute = SendTarget.toindex Then
        
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).Name & " <" & klan & ">" & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
            Else
                Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).Name & " <" & klan & ">" & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
            End If
        
        ElseIf sndRoute = SendTarget.ToMap Then
            Call AgregarUser(userindex, UserList(userindex).Pos.Map)
            Call CheckUpdateNeededUser(userindex, USER_NUEVO)
        End If
    Else 'if tiene clan
        If sndRoute = SendTarget.toindex Then
        
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                Call SendData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).Name & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
            Else
                Call SendData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).Name & "," & UserList(userindex).StatusMith.EsStatus & "," & "," & UserList(userindex).flags.Privilegios)
            End If
            
        ElseIf sndRoute = SendTarget.ToMap Then
            Call AgregarUser(userindex, UserList(userindex).Pos.Map)
            Call CheckUpdateNeededUser(userindex, USER_NUEVO)
        End If
    End If 'if clan
    
End If
Exit Sub
 
hayerror:
LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
'Resume Next
Call CloseSocket(userindex)
End Sub
Sub CheckUserLevel(ByVal userindex As Integer)

On Error GoTo Errhandler

Dim PTS As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim WasNewbie As Boolean

Dim VidaVeinte As Integer
Dim VidaMaxima As Integer
Dim VidaMinima As Integer
Dim RandomVeinteMin As Byte
Dim RandomVeinteMax As Byte
Dim DefineRandomMin As Byte
Dim DefineRandomMax As Byte
Dim RandomFinal As Byte

'¿Alcanzo el maximo nivel?
If UserList(userindex).Stats.ELV = 60 Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.ELU = 0
    Exit Sub
End If

If UserList(userindex).Stats.ELV >= 50 Then

    Do While (UserList(userindex).Stats.Exp >= UserList(userindex).Stats.ELU) And (UserList(userindex).Stats.ELV < 60)
    
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_NIVEL)
        Call SendData(SendTarget.toindex, userindex, 0, "||67")
    
           
        UserList(userindex).Stats.ELV = UserList(userindex).Stats.ELV + 1
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 58 & "," & 0)
        
        UserList(userindex).Stats.Exp = 0
        If UserList(userindex).Stats.ELV < 60 Then
            UserList(userindex).Stats.ELU = ArrayExp(UserList(userindex).Stats.ELV)
        Else
            UserList(userindex).Stats.ELU = 0
        End If
        
        Call EnviarSkills(userindex)
        Call EnviarSubirNivel(userindex, PTS)


        If UserList(userindex).Stats.ELV = 53 Then
            '  Dim vidasumadita As Byte
            '  vidasumadita = RandomNumber(2, 4)
            '  Call SendData(SendTarget.toindex, userindex, 0, "||68@53@" & vidasumadita)
            '  UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + vidasumadita
            '  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
            Call SendData(SendTarget.toindex, userindex, 0, "99" & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion1") & "," & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion2"))
        ElseIf UserList(userindex).Stats.ELV = 56 Then
            '  vidasumadita = RandomNumber(2, 4)
            '  Call SendData(SendTarget.toindex, userindex, 0, "||68@56@" & vidasumadita)
            '  UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + vidasumadita
            '  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
            Call SendData(SendTarget.toindex, userindex, 0, "99" & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion1") & "," & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion2"))
        ElseIf UserList(userindex).Stats.ELV = 60 Then
                If UserList(userindex).flags.Llegolvlmax = 0 Then
                    UserList(userindex).Stats.Exp = 0
                    UserList(userindex).Stats.ELU = 0
                    Call SendData(SendTarget.ToAll, userindex, 0, "||70@" & UserList(userindex).Name)
                    Call SendData(SendTarget.toindex, userindex, 0, "||57@200")
                    Call AgregarPuntos(userindex, 200)
                    UserList(userindex).flags.Llegolvlmax = 1
                    SendUserEXP userindex
                End If
                
                Call SendData(SendTarget.toindex, userindex, 0, "99" & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel3Opcion1") & "," & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel3Opcion2"))
        End If
  
        SendUserEXP userindex
        SendUserLVL userindex
        
    Loop
    
Exit Sub
End If


WasNewbie = EsNewbie(userindex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(userindex).Stats.Exp >= UserList(userindex).Stats.ELU
    
    'Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_NIVEL)
    Call SendData(SendTarget.toindex, userindex, 0, "||67")
    
    If UserList(userindex).Stats.ELV = 1 Then
        PTS = 10
    Else
        PTS = 10
    End If
    
    UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts + PTS
       
    UserList(userindex).Stats.ELV = UserList(userindex).Stats.ELV + 1
        
     If UserList(userindex).Stats.ELV = 20 Then
        Dim j As Integer
        If Not TieneHechizo(8, userindex) Then
            'Buscamos un slot vacio
            For j = 1 To MAXUSERHECHIZOS
                If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
            Next j
                
            If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
                Exit Sub
            Else
                UserList(userindex).Stats.UserHechizos(j) = 8
                Call UpdateUserHechizos(False, userindex, CByte(j))
            End If
        End If
    End If
    
    UserList(userindex).Stats.Exp = 0
    
If UserList(userindex).Stats.ELV = 50 Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.ELU = 999999999
    If UserList(userindex).flags.llegolvl50 = 0 Then
    Call SendData(SendTarget.ToAll, userindex, 0, "||69@" & UserList(userindex).Name & "@50")
    Call SendData(SendTarget.toindex, userindex, 0, "||57@50")
    Call AgregarPuntos(userindex, 50)
    UserList(userindex).flags.llegolvl50 = 1
    End If
End If
    
    
    If UserList(userindex).Stats.ELV = 10 Then
        Call QuitarNewbieObj(userindex)
    End If

        'Seteamos la experiencia para el próximo nivel
        UserList(userindex).Stats.ELU = ArrayExp(UserList(userindex).Stats.ELV)

    Dim AumentoHP As Integer
    Select Case UCase$(UserList(userindex).clase)
        Case "GUERRERO"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 555
                    VidaMaxima = 565
                    RandomVeinteMin = 10
                    RandomVeinteMax = 11
                    VidaVeinte = 225
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case "ELFO"
                    VidaMinima = 520
                    VidaMaxima = 530
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 190
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case "ELFO OSCURO"
                    VidaMinima = 525
                    VidaMaxima = 535
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 195
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case "GNOMO"
                    VidaMinima = 495
                    VidaMaxima = 505
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 205
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 10
                Case "ENANO"
                    VidaMinima = 575
                    VidaMaxima = 585
                    RandomVeinteMin = 11
                    RandomVeinteMax = 12
                    VidaVeinte = 245
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case Else
                    AumentoHP = RandomNumber(8, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "CAZADOR"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 485
                    VidaMaxima = 495
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 195
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 10
                Case "ELFO"
                    VidaMinima = 465
                    VidaMaxima = 475
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 195
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 9
                Case "ELFO OSCURO"
                    VidaMinima = 475
                    VidaMaxima = 485
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 205
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 9
                Case "GNOMO"
                    VidaMinima = 445
                    VidaMaxima = 455
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 185
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 9
                Case "ENANO"
                    VidaMinima = 525
                    VidaMaxima = 535
                    RandomVeinteMin = 10
                    RandomVeinteMax = 11
                    VidaVeinte = 225
                    DefineRandomMin = 10
                    DefineRandomMax = 11
                    RandomFinal = 10
            End Select

            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "PALADIN"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 505
                    VidaMaxima = 515
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 205
                    DefineRandomMin = 10
                    DefineRandomMax = 11
                    RandomFinal = 10
                    
                Case "ELFO"
                    VidaMinima = 485
                    VidaMaxima = 495
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 195
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 10
                    
                Case "ELFO OSCURO"
                    VidaMinima = 490
                    VidaMaxima = 500
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 200
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 10
                    
                Case "GNOMO"
                    VidaMinima = 475
                    VidaMaxima = 485
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 195
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 10
                    
                Case "ENANO"
                    VidaMinima = 515
                    VidaMaxima = 525
                    RandomVeinteMin = 10
                    RandomVeinteMax = 11
                    VidaVeinte = 215
                    DefineRandomMin = 10
                    DefineRandomMax = 11
                    RandomFinal = 10
                Case Else
                    AumentoHP = RandomNumber(9, 10)
                End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) - 1
            AumentoSTA = AumentoSTDef
            
        Case "MAGO" 'QUEDOOOOOOOOOOOADOOOOOO
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 385
                    VidaMaxima = 395
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 155
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "ELFO"
                    VidaMinima = 365
                    VidaMaxima = 375
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 155
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 7
                Case "GNOMO"
                    VidaMinima = 335
                    VidaMaxima = 345
                    RandomVeinteMin = 6
                    RandomVeinteMax = 6
                    VidaVeinte = 135
                    DefineRandomMin = 6
                    DefineRandomMax = 7
                    RandomFinal = 7
                Case "ELFO OSCURO"
                    VidaMinima = 370
                    VidaMaxima = 380
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 160
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 7
                Case "ENANO"
                    VidaMinima = 395
                    VidaMaxima = 405
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 155
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case Else
                    AumentoHP = RandomNumber(7, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1
            If UCase$(UserList(userindex).Raza) = "ELFO OSCURO" Then
                AumentoMANA = 2.9 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            ElseIf UCase$(UserList(userindex).Raza) = "HUMANO" Then
                AumentoMANA = 2.95 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            Else
                AumentoMANA = 2.85 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            End If
            AumentoSTA = AumentoSTMago
        
        Case "CLERIGO"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 445
                    VidaMaxima = 455
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 9
                Case "ELFO"
                    VidaMinima = 415
                    VidaMaxima = 425
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case "ELFO OSCURO"
                    VidaMinima = 420
                    VidaMaxima = 430
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 180
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case "GNOMO"
                    VidaMinima = 390
                    VidaMaxima = 400
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 170
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "ENANO"
                    VidaMinima = 460
                    VidaMaxima = 470
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 190
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 9
                Case Else
                    AumentoHP = RandomNumber(8, 8)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 1.9 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "DRUIDA"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 415
                    VidaMaxima = 425
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case "ELFO"
                    VidaMinima = 395
                    VidaMaxima = 405
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 165
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "ELFO OSCURO"
                    VidaMinima = 405
                    VidaMaxima = 415
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "GNOMO"
                    VidaMinima = 370
                    VidaMaxima = 380
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 170
                    DefineRandomMin = 6
                    DefineRandomMax = 7
                    RandomFinal = 7
                Case "ENANO"
                    VidaMinima = 430
                    VidaMaxima = 440
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 190
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case Else
                    AumentoHP = RandomNumber(6, 8)
            End Select
            
            AumentoHIT = 2
            If UCase$(UserList(userindex).Raza) = "GNOMO" Then
                AumentoMANA = 1.98 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            Else
                AumentoMANA = 1.93 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            End If
            AumentoSTA = AumentoSTDef
        
        Case "ASESINO"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 415
                    VidaMaxima = 425
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case "ELFO"
                    VidaMinima = 395
                    VidaMaxima = 405
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 165
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "ELFO OSCURO"
                    VidaMinima = 400
                    VidaMaxima = 410
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 170
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "GNOMO"
                    VidaMinima = 375
                    VidaMaxima = 385
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 6
                    DefineRandomMax = 7
                    RandomFinal = 7
                Case "ENANO"
                    VidaMinima = 425
                    VidaMaxima = 435
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 185
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case Else
                    AumentoHP = RandomNumber(7, 8)
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "BARDO"
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 445
                    VidaMaxima = 455
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 9
                Case "ELFO"
                    VidaMinima = 415
                    VidaMaxima = 425
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 175
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case "ELFO OSCURO"
                    VidaMinima = 420
                    VidaMaxima = 430
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 180
                    DefineRandomMin = 8
                    DefineRandomMax = 9
                    RandomFinal = 8
                Case "GNOMO"
                    VidaMinima = 390
                    VidaMaxima = 400
                    RandomVeinteMin = 7
                    RandomVeinteMax = 8
                    VidaVeinte = 170
                    DefineRandomMin = 7
                    DefineRandomMax = 8
                    RandomFinal = 8
                Case "ENANO"
                    VidaMinima = 460
                    VidaMaxima = 470
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 190
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 9
                Case Else
                    AumentoHP = RandomNumber(8, 8)
            End Select
            
            AumentoHIT = 2
            If UCase$(UserList(userindex).Raza) = "HUMANO" Then
                AumentoMANA = 1.94 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            ElseIf UCase$(UserList(userindex).Raza) = "GNOMO" Then
                AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            Else
                AumentoMANA = 1.98 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            End If
            AumentoSTA = AumentoSTDef
    
        Case Else
            Select Case UCase$(UserList(userindex).Raza)
                Case "HUMANO"
                    VidaMinima = 555
                    VidaMaxima = 565
                    RandomVeinteMin = 10
                    RandomVeinteMax = 11
                    VidaVeinte = 225
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case "ELFO"
                    VidaMinima = 520
                    VidaMaxima = 530
                    RandomVeinteMin = 8
                    RandomVeinteMax = 9
                    VidaVeinte = 190
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case "ELFO OSCURO"
                    VidaMinima = 525
                    VidaMaxima = 535
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 195
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case "GNOMO"
                    VidaMinima = 495
                    VidaMaxima = 505
                    RandomVeinteMin = 9
                    RandomVeinteMax = 10
                    VidaVeinte = 205
                    DefineRandomMin = 9
                    DefineRandomMax = 10
                    RandomFinal = 10
                Case "ENANO"
                    VidaMinima = 575
                    VidaMaxima = 585
                    RandomVeinteMin = 11
                    RandomVeinteMax = 12
                    VidaVeinte = 245
                    DefineRandomMin = 11
                    DefineRandomMax = 12
                    RandomFinal = 11
                Case Else
                    AumentoHP = RandomNumber(8, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
    End Select
    
    If UserList(userindex).Stats.ELV <= 20 Then
        If UserList(userindex).Stats.MaxHP + (21 - UserList(userindex).Stats.ELV) * RandomVeinteMin < VidaVeinte Then
            AumentoHP = RandomVeinteMax
        Else
            AumentoHP = RandomVeinteMin
        End If
    ElseIf UserList(userindex).Stats.ELV < 30 Then
        AumentoHP = RandomNumber(DefineRandomMin, DefineRandomMax)
    Else
        AumentoHP = RandomFinal
    End If
    
    'Actualizamos HitPoints
    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + AumentoHP
        
    'VIDAS
    'Actualizamos Stamina
    UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MaxSta + AumentoSTA
    If UserList(userindex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(userindex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(userindex).Stats.MaxMAN = UserList(userindex).Stats.MaxMAN + AumentoMANA
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxMAN > STAT_MAXMAN Then _
            UserList(userindex).Stats.MaxMAN = STAT_MAXMAN
    Else
        If UserList(userindex).Stats.MaxMAN > 9999 Then _
            UserList(userindex).Stats.MaxMAN = 9999
    End If
    
    'Actualizamos Golpe Máximo
    UserList(userindex).Stats.MaxHIT = UserList(userindex).Stats.MaxHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(userindex).Stats.MinHIT = UserList(userindex).Stats.MinHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then SendData SendTarget.toindex, userindex, 0, "||71@" & AumentoHP
    If AumentoSTA > 0 Then SendData SendTarget.toindex, userindex, 0, "||72@" & AumentoSTA
    If AumentoMANA > 0 Then SendData SendTarget.toindex, userindex, 0, "||73@" & AumentoMANA
    If AumentoHIT > 0 Then
        SendData SendTarget.toindex, userindex, 0, "||74@" & AumentoHIT
        SendData SendTarget.toindex, userindex, 0, "||75@" & AumentoHIT
    End If
    
    With UserList(userindex)
    
        If .Stats.ELV < 20 Then
            SendData SendTarget.toindex, userindex, 0, "||76@" & .clase & "@" & .Raza & "@" & VidaMinima & "@" & VidaMaxima
        ElseIf .Stats.ELV <= 29 Then
            SendData SendTarget.toindex, userindex, 0, "||77@" & .Stats.MaxHP + (30 - .Stats.ELV) * DefineRandomMin + (20 * RandomFinal) & "@" & .Stats.MaxHP + (30 - .Stats.ELV) * DefineRandomMax + (20 * RandomFinal)
        Else
            SendData SendTarget.toindex, userindex, 0, "||78@" & .Stats.MaxHP + (50 - .Stats.ELV) * RandomFinal
        End If
                    
        Dim conso As String
        conso = .Char.CharIndex
        'Call SendData(SendTarget.toindex, UserIndex, .Pos.Map, "N|" & vbCyan & "°" & "Vida +" & AumentoHP & " Mana +" & AumentoMANA & " Golpe +" & AumentoHIT & "." & "°" & conso)
        Call SendData(SendTarget.ToPCArea, userindex, .Pos.Map, "CFF" & .Char.CharIndex & "," & 58 & "," & 0)
        Call LogDesarrollo(Date & " " & .Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
        .Stats.MinHP = .Stats.MaxHP
    End With
        
        Call EnviarSkills(userindex)
        Call EnviarSubirNivel(userindex, PTS)
    
        SendUserLVL userindex
        SendUserHP userindex
        SendUserMP userindex
        SendUserST userindex
Loop
'End If

Exit Sub
Errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal userindex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(userindex).flags.Navegando = 1

End Function

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As Byte)

Dim nPos As WorldPos
Dim UserMIndex As Integer
Dim BlokeVacio As Byte 'Ahorramos comprobacion :)
    BlokeVacio = 0 'asegurarse¿?¿? xD Bromilla :P
    
    nPos = UserList(userindex).Pos
    Call HeadtoPos(nHeading, nPos)
    
UserMIndex = MapData(nPos.Map, nPos.X, nPos.Y).userindex
    'Mover Casper
    If UserMIndex > 0 Then
        If UserList(UserMIndex).flags.Muerto = 1 Then
            Select Case nHeading
            Case 1 'Norte
                If LegalPos(nPos.Map, nPos.X, nPos.Y - 1) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X, nPos.Y - 1)
                    BlokeVacio = 1
                End If
            Case 2 'Este
                If LegalPos(nPos.Map, nPos.X + 1, nPos.Y) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X + 1, nPos.Y)
                    BlokeVacio = 1
                End If
            Case 3 'Sur
                If LegalPos(nPos.Map, nPos.X, nPos.Y + 1) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X, nPos.Y + 1)
                    BlokeVacio = 1
                End If
            Case 4 'Oeste
                If LegalPos(nPos.Map, nPos.X - 1, nPos.Y) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X - 1, nPos.Y)
                    BlokeVacio = 1
                End If
            End Select
           
            If BlokeVacio = 0 Then
                If LegalPos(nPos.Map, nPos.X, nPos.Y + 1) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X, nPos.Y + 1)
                ElseIf LegalPos(nPos.Map, nPos.X, nPos.Y - 1) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X, nPos.Y - 1)
                ElseIf LegalPos(nPos.Map, nPos.X + 1, nPos.Y) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X + 1, nPos.Y)
                ElseIf LegalPos(nPos.Map, nPos.X - 1, nPos.Y) Then
                    Call WarpUserChar(UserMIndex, nPos.Map, nPos.X - 1, nPos.Y + 1)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
                    Call SendData(SendTarget.toindex, UserMIndex, 0, "PU" & UserList(UserMIndex).Pos.X & "," & UserList(UserMIndex).Pos.Y)
                    Exit Sub
                End If
            End If
        End If
    End If
    'Fin Mover Casper
    
    If LegalPos(UserList(userindex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(userindex)) Then
        If MapInfo(UserList(userindex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...
            Call SendToUserAreaButindex(userindex, "+" & UserList(userindex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
        End If
        
        'Update map and user pos
        MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = 0
        UserList(userindex).Pos = nPos
        UserList(userindex).Char.Heading = nHeading
        MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = userindex
        
      If UserList(userindex).Pos.Map = mapainvo And InvocoBicho = False Then
         If MapData(mapainvo, mapainvoX1, mapainvoY1).userindex > 0 And MapData(mapainvo, mapainvoX2, mapainvoY2).userindex > 0 And MapData(mapainvo, mapainvoX3, mapainvoY3).userindex > 0 And MapData(mapainvo, mapainvoX4, mapainvoY4).userindex > 0 Then
            Call SendData(SendTarget.ToMap, 0, mapainvo, "PCF" & 82 & "," & 29 & "," & 29 & "," & 30)
            SegundosInvo = 2
            InvocoBicho = True
        End If
       End If
        
        If ZonaCura(userindex) Then Call AutoCuraUser(userindex)
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(userindex, nHeading)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
    End If
    
    If UserList(userindex).Counters.Trabajando Then _
        UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1
End Sub

Sub ChangeUserInv(userindex As Integer, Slot As Byte, Object As UserOBJ)

    UserList(userindex).Invent.Object(Slot) = Object
    
    If Object.ObjIndex > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef & "," _
        & ObjData(Object.ObjIndex).Valor \ 3)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")
    End If

End Sub


Function NextOpenCharIndex() As Integer
'Modificada por el oso para codificar los MP1234,2,1 en 2 bytes
'para lograrlo, el charindex no puede tener su bit numero 6 (desde 0) en 1
'y tampoco puede ser un charindex que tenga el bit 0 en 1.

On Local Error GoTo hayerror

Dim loopC As Integer
    
    loopC = 1
    
    While loopC < MAXCHARS
        If CharList(loopC) = 0 And Not ((loopC And &HFFC0&) = 64) Then
            NextOpenCharIndex = loopC
            NumChars = NumChars + 1
            If loopC > LastChar Then LastChar = loopC
            Exit Function
        Else
            loopC = loopC + 1
        End If
    Wend

Exit Function
hayerror:
LogError ("NextOpenCharIndex: num: " & Err.Number & " desc: " & Err.Description)

End Function
Function NextOpenUser() As Integer
    Dim loopC As Long
   
    For loopC = 1 To MaxUsers + 1
        If loopC > MaxUsers Then Exit For
        If (UserList(loopC).ConnID = -1 And UserList(loopC).flags.UserLogged = False) Then Exit For
    Next loopC
   
    NextOpenUser = loopC
End Function
 
Sub SendUserHitBox(ByVal userindex As Integer)
Dim lagaminarma As Integer
Dim lagamaxarma As Integer
Dim lagaminarmor As Integer
Dim lagamaxarmor As Integer
Dim lagaminescu As Integer
Dim lagamaxescu As Integer
Dim lagamincasc As Integer
Dim lagamaxcasc As Integer
Dim lagaminherr As Integer
Dim lagamaxherr As Integer
 
Dim llagamindef As Integer
Dim llagamaxdef As Integer
 
Dim llagamindefa As Integer
Dim llagamaxdefa As Integer
 
Dim llagamindefb As Integer
Dim llagamaxdefb As Integer
 
Dim llagamindefc As Integer
Dim llagamaxdefc As Integer
 
Dim llagamindefd As Integer
Dim llagamaxdefd As Integer
 
If UserList(userindex).Invent.WeaponEqpSlot > 0 Then
lagaminarma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT
lagamaxarma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT
llagamindef = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DefensaMagicaMin
llagamaxdef = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DefensaMagicaMax
Else
lagaminarma = "0"
lagamaxarma = "0"
llagamindef = "0"
llagamaxdef = "0"
End If
If UserList(userindex).Invent.ArmourEqpSlot > 0 Then
lagaminarmor = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef
lagamaxarmor = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef
llagamindefa = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin
llagamaxdefa = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax
Else
lagaminarmor = "0"
lagamaxarmor = "0"
llagamindefa = "0"
llagamaxdefa = "0"
End If
If UserList(userindex).Invent.EscudoEqpSlot > 0 Then
lagaminescu = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MinDef
lagamaxescu = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MaxDef
llagamindefb = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).DefensaMagicaMin
llagamaxdefb = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).DefensaMagicaMax
Else
lagaminescu = "0"
lagamaxescu = "0"
llagamindefb = "0"
llagamaxdefb = "0"
End If
If UserList(userindex).Invent.CascoEqpSlot > 0 Then
lagamincasc = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef
lagamaxcasc = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef
llagamindefc = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin
llagamaxdefc = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax
Else
lagamincasc = "0"
lagamaxcasc = "0"
llagamindefc = "0"
llagamaxdefc = "0"
End If
If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
llagamindefd = ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin
llagamaxdefd = ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax
Else
lagaminherr = "0"
lagamaxherr = "0"
llagamindefd = "0"
llagamaxdefd = "0"
End If
Call SendData(toindex, userindex, 0, "ANM" & lagaminarma & "," & lagamaxarma & "," & lagaminarmor & "," & lagamaxarmor & "," & lagaminescu & "," & lagamaxescu & "," & lagamincasc & "," & lagamaxcasc & "," & lagaminherr & "," & lagamaxherr & "," & llagamindef & "," & llagamaxdef & "," & llagamindefa & "," & llagamaxdefa & "," & llagamindefb & "," & llagamaxdefb & "," & llagamindefc & "," & llagamaxdefc & "," & llagamindefd & "," & llagamaxdefd)
End Sub
Sub EnviarPuntos(ByVal userindex As Integer)
 Call SendData(SendTarget.toindex, userindex, 0, "PNT" & UserList(userindex).Stats.PuntosTorneo)
End Sub
Sub EnviarHambreYsed(ByVal userindex As Integer)
    Call SendData(SendTarget.toindex, userindex, 0, "EHYS" & UserList(userindex).Stats.MaxAGU & "," & UserList(userindex).Stats.MinAGU & "," & UserList(userindex).Stats.MaxHam & "," & UserList(userindex).Stats.MinHam)
End Sub
Sub SendUserStatux(ByVal userindex As Integer)
On Error Resume Next
Dim Info As String
 
Call SendData(SendTarget.toindex, userindex, 0, "EZT" & "," & UserList(userindex).StatusMith.EsStatus)
 
Info = "PX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).Name
Call SendData(ToMap, userindex, UserList(userindex).Pos.Map, (Info))
 
End Sub
Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
Dim GuildI As Integer


    Call SendData(SendTarget.toindex, sendIndex, 0, "||855@" & UserList(userindex).Name)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||856@" & UserList(userindex).clase)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||857@" & UserList(userindex).Raza)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||858@" & UserList(userindex).Genero)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||859@" & UserList(userindex).Stats.Reputacione)
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "||860@" & UserList(userindex).Stats.ELV & "@" & UserList(userindex).Stats.Exp & "@" & UserList(userindex).Stats.ELU)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||861@" & UserList(userindex).Stats.MinSta & "@" & UserList(userindex).Stats.MaxSta)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||862@" & UserList(userindex).Stats.MinHP & "@" & UserList(userindex).Stats.MaxHP)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||863@" & UserList(userindex).Stats.MinMAN & "@" & UserList(userindex).Stats.MaxMAN)
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "||864@" & UserList(userindex).Stats.MinHIT & "@" & UserList(userindex).Stats.MaxHIT)
        
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||865@" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef & "@" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef)
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||865@0@0")
    End If
    
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||866@" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef & "@" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef)
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||866@0@0")
    End If
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "||867@" & UserList(userindex).Stats.GLD & "@" & UserList(userindex).Pos.X & "@" & UserList(userindex).Pos.Y & "@" & UserList(userindex).Pos.Map)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||868@" & UserList(userindex).Stats.Banco)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||869@" & UserList(userindex).Faccion.CiudadanosMatados & "@" & UserList(userindex).Faccion.CriminalesMatados & "@" & UserList(userindex).Faccion.NeutralesMatados)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||870@" & UserList(userindex).Stats.NPCsMuertos)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||871@" & UserList(userindex).Stats.PuntosTorneo)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||872@" & UserList(userindex).Stats.PuntosDonacion)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||873@" & UserList(userindex).flags.TiempoOnlineHoy)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||874@" & UCase$(UserList(userindex).Accounted))
    Call SendData(SendTarget.toindex, sendIndex, 0, "||875@1@" & UserList(userindex).Bon1)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||875@2@" & UserList(userindex).Bon2)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||875@3@" & UserList(userindex).Bon3)

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
 
With UserList(userindex)
Call SendData(SendTarget.toindex, sendIndex, 0, "N|Pj: " & .Name & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "N|CiudadanosMatados: " & .Faccion.CiudadanosMatados & "CriminalesMatados: " & .Faccion.CriminalesMatados & "NeutralesMatados: " & .Faccion.NeutralesMatados & "UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "N|NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "N|Clase: " & .clase & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "N|Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With
 
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "N|Pj: " & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, sendIndex, 0, "N|Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.toindex, sendIndex, 0, "N|Ban: " & Ban & FONTTYPE_INFO)
        If Ban = "1" Then
            Call SendData(SendTarget.toindex, sendIndex, 0, "N|Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "N|El pj no existe: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub
Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next

    Dim j As Long
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|" & UserList(userindex).Name & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "N| Tiene " & UserList(userindex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.toindex, sendIndex, 0, "N| Objeto " & j & " " & ObjData(UserList(userindex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(userindex).Invent.Object(j).Amount & FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "N|" & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, sendIndex, 0, "N| Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call SendData(SendTarget.toindex, sendIndex, 0, "N| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
            End If
        Next j
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||189@" & CharName)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.toindex, sendIndex, 0, "N|" & UserList(userindex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(SendTarget.toindex, sendIndex, 0, "N| " & SkillsNames(j) & " = " & UserList(userindex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
Call SendData(SendTarget.toindex, sendIndex, 0, "N| SkillLibres:" & UserList(userindex).Stats.SkillPts & FONTTYPE_INFO)
End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim loopC As Integer
  
loopC = 1
  
Do Until UserList(loopC).ConnID = SocketId

    loopC = loopC + 1
    
    If loopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = loopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim loopC As Integer
  
loopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(loopC).Name) = Nombre

    loopC = loopC + 1
    
    If loopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = loopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(SendTarget.toindex, Npclist(NpcIndex).MaestroUser, 0, "||876@" & UserList(userindex).Name)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal userindex As Integer)

'Fortaleza - rey.
If Npclist(NpcIndex).Pos.Map = 81 And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 20000 And Npclist(NpcIndex).Stats.MinHP < 30000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||79@" & Guilds(UserList(userindex).GuildIndex).GuildName)

If Npclist(NpcIndex).Pos.Map = MapCastilloN And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 20000 And Npclist(NpcIndex).Stats.MinHP <> 30000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloNorte), 0, "||80@" & MapCastilloN & "@" & Guilds(UserList(userindex).GuildIndex).GuildName)
If Npclist(NpcIndex).Pos.Map = MapCastilloS And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 20000 And Npclist(NpcIndex).Stats.MinHP <> 30000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloSur), 0, "||80@" & MapCastilloS & "@" & Guilds(UserList(userindex).GuildIndex).GuildName)
If Npclist(NpcIndex).Pos.Map = MapCastilloE And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 20000 And Npclist(NpcIndex).Stats.MinHP <> 30000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloEste), 0, "||80@" & MapCastilloE & "@" & Guilds(UserList(userindex).GuildIndex).GuildName)
If Npclist(NpcIndex).Pos.Map = MapCastilloO And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 20000 And Npclist(NpcIndex).Stats.MinHP <> 30000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloOeste), 0, "||80@" & MapCastilloO & "@" & Guilds(UserList(userindex).GuildIndex).GuildName)

'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(userindex).Name
 
 If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(userindex, Npclist(NpcIndex).MaestroUser)
 
'Si atacaste mascota, te las picas de ciuda =D - Mithrandir
If EsMascotaCiudadano(NpcIndex, userindex) Then
Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
Npclist(NpcIndex).Hostile = 1
Else
'Reputacion
If Npclist(NpcIndex).Stats.Alineacion = 0 Then
End If
 
'hacemos que el npc se defienda
Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
Npclist(NpcIndex).Hostile = 1
End If
 
End Sub
Function PuedeApuñalar(ByVal userindex As Integer) As Boolean

If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UCase$(UserList(userindex).clase) = "ASESINO") And _
  (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal userindex As Integer, ByVal Skill As Integer)

If UserList(userindex).flags.Hambre = 0 And _
   UserList(userindex).flags.Sed = 0 Then
    Dim Aumenta As Integer
    Dim Prob As Integer
    
    If UserList(userindex).Stats.ELV <= 3 Then
        Prob = 25
    ElseIf UserList(userindex).Stats.ELV > 3 _
        And UserList(userindex).Stats.ELV < 6 Then
        Prob = 35
    ElseIf UserList(userindex).Stats.ELV >= 6 _
        And UserList(userindex).Stats.ELV < 10 Then
        Prob = 40
    ElseIf UserList(userindex).Stats.ELV >= 10 _
        And UserList(userindex).Stats.ELV < 20 Then
        Prob = 45
    Else
        Prob = 50
    End If
    
    Aumenta = RandomNumber(1, Prob)
    
    Dim lvl As Integer
    lvl = UserList(userindex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(userindex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
        UserList(userindex).Stats.UserSkills(Skill) = UserList(userindex).Stats.UserSkills(Skill) + 1
        'Call SendData(SendTarget.toindex, Userindex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(Userindex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
        
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + 50
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
        
        'Call SendData(SendTarget.toindex, Userindex, 0, "||¡Has ganado 50 puntos de experiencia!" & FONTTYPE_FIGHT)
        Call CheckUserLevel(userindex)
    End If
End If

End Sub

Sub UserDie(ByVal userindex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UCase$(UserList(userindex).Genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "QDL" & UserList(userindex).Char.CharIndex)
    
       If Criminal(userindex) Or Ciudadano(userindex) Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbRed & "°" & "¡Aaaahhhh!" & "°" & str(UserList(userindex).Char.CharIndex))
        Else
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbRed & "°" & "¡Aaaahhhh!" & "°" & str(UserList(userindex).Char.CharIndex))
        End If
        
    UserList(userindex).Stats.MinHP = 0
    UserList(userindex).Stats.MinSta = 0
    UserList(userindex).flags.AtacadoPorNpc = 0
    UserList(userindex).flags.AtacadoPorUser = 0
    UserList(userindex).flags.Envenenado = 0
    UserList(userindex).flags.Muerto = 1
    UserList(userindex).flags.TimeRevivir = 20
    
    Call SendUserHitBox(userindex)
    Dim aN As Integer
    
    aN = UserList(userindex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""
    End If
    
    '<<< Montura >>>
    If UserList(userindex).flags.Montando = 1 Then
     Call Desmontar(userindex)
    End If
    
    If UserList(userindex).flags.InvocoMascota = 1 Then
     Call QuitarNPC(UserList(userindex).flags.MascotinIndex)
     UserList(userindex).flags.InvocoMascota = 0
    End If
    
    '<<<< Paralisis >>>>
    If UserList(userindex).flags.Paralizado = 1 Then
        UserList(userindex).flags.Paralizado = 0
        Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
    End If
    
    '<<<< Meditando >>>>
    If UserList(userindex).flags.Meditando Then
        UserList(userindex).flags.Meditando = False
        Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
    End If
       
    '<<<<< Seg Resu >>>>>
    If UserList(userindex).flags.SeguroResu = False Then
    Call SendData(SendTarget.toindex, userindex, 0, "SEGONR")
        UserList(userindex).flags.SeguroResu = True
    End If
    
    '<<<< Invisible >>>>
    If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).flags.Invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
    End If
    
    If TriggerZonaPelea(userindex, userindex) <> TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(userindex) Or Criminal(userindex) Then
            Call TirarTodo(userindex)
        Else
            If EsNewbie(userindex) Then Call TirarTodosLosItemsNoNewbies(userindex)
        End If
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpSlot)
    End If
    'desequipar municiones
    If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
    End If
    
    If UserList(userindex).flags.IntervaloBurbu > 1 Then
        UserList(userindex).flags.IntervaloBurbu = 0
        UserList(userindex).flags.DefensaBurbu = 0
        SendData SendTarget.toindex, userindex, 0, "||81"
    End If
    
    If UserList(userindex).flags.EnAram Then
        Call Aram_ContarMuerte(userindex)
    End If
    
    'desequipar escudo
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
    End If
    
    If UserList(userindex).flags.Automatico = True Then
        Call Rondas_UsuarioMuere(userindex)
    End If
    
    ' << Restauramos el mimetismo
    If UserList(userindex).flags.Mimetizado = 1 Then
        UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
        UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
        UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
        UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
        UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        UserList(userindex).Counters.Mimetismo = 0
        UserList(userindex).flags.Mimetizado = 0
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(userindex).flags.Navegando = 0 Then
        If UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Or UserList(userindex).StatusMith.EsStatus = 5 Then
           UserList(userindex).Char.Body = iCuerpoMuertoA
           UserList(userindex).Char.Head = iCabezaMuertoA
        ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Or UserList(userindex).StatusMith.EsStatus = 6 Then
           UserList(userindex).Char.Body = iCuerpoMuertoH
           UserList(userindex).Char.Head = iCabezaMuertoH
        Else
           UserList(userindex).Char.Body = iCuerpoMuertoN
           UserList(userindex).Char.Head = iCabezaMuertoN
        End If
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.CascoAnim = NingunCasco
    Else
        UserList(userindex).Char.Body = iFragataFantasmal ';)
    End If
        
    If UserList(userindex).flags.EnDuelo Then Call RevisarDuelo(userindex)
        
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> casted - pareja 2vs2
    If UserList(userindex).flags.EnPareja And HayPareja Then
        If UserList(Pareja.Jugador(1)).flags.EnPareja = True And UserList(Pareja.Jugador(2)).flags.EnPareja = True And UserList(Pareja.Jugador(1)).flags.Muerto = 1 And UserList(Pareja.Jugador(2)).flags.Muerto = 1 Then
        
            Dim i As Long
            
            HayPareja = False
            UserList(Pareja.Jugador(1)).Stats.ParejasPerdidas = UserList(Pareja.Jugador(1)).Stats.ParejasPerdidas + 1
            UserList(Pareja.Jugador(2)).Stats.ParejasPerdidas = UserList(Pareja.Jugador(2)).Stats.ParejasPerdidas + 1
            
            UserList(Pareja.Jugador(3)).Stats.ParejasGanadas = UserList(Pareja.Jugador(3)).Stats.ParejasGanadas + 1
            UserList(Pareja.Jugador(4)).Stats.ParejasGanadas = UserList(Pareja.Jugador(4)).Stats.ParejasGanadas + 1
            
            If MisionesDiarias(UserList(Pareja.Jugador(3)).Misiones.NumeroMision).Tipo = 7 Then
                If UserList(Pareja.Jugador(3)).Misiones.ConteoUser < MisionesDiarias(UserList(Pareja.Jugador(3)).Misiones.NumeroMision).Cantidad Then UserList(Pareja.Jugador(3)).Misiones.ConteoUser = UserList(Pareja.Jugador(3)).Misiones.ConteoUser + 1
            End If
            
            If MisionesDiarias(UserList(Pareja.Jugador(4)).Misiones.NumeroMision).Tipo = 7 Then
                If UserList(Pareja.Jugador(4)).Misiones.ConteoUser < MisionesDiarias(UserList(Pareja.Jugador(4)).Misiones.NumeroMision).Cantidad Then UserList(Pareja.Jugador(4)).Misiones.ConteoUser = UserList(Pareja.Jugador(4)).Misiones.ConteoUser + 1
            End If
            
            Call GRANK_User_Check(Couple, UserList(Pareja.Jugador(3)).Name, UserList(Pareja.Jugador(3)).Stats.ParejasGanadas)
            Call GRANK_User_Check(Couple, UserList(Pareja.Jugador(4)).Name, UserList(Pareja.Jugador(4)).Stats.ParejasGanadas)
            UserList(Pareja.Jugador(3)).Stats.GLD = UserList(Pareja.Jugador(3)).Stats.GLD + 200000
            UserList(Pareja.Jugador(4)).Stats.GLD = UserList(Pareja.Jugador(4)).Stats.GLD + 200000
            Call SendData(SendTarget.toindex, Pareja.Jugador(3), 0, "||63@200.000")
            Call SendData(SendTarget.toindex, Pareja.Jugador(4), 0, "||63@200.000")
            SendUserGLD (Pareja.Jugador(3))
            SendUserGLD (Pareja.Jugador(4))
            
            Call SendData(SendTarget.ToAll, 0, 0, "||82@" & UserList(Pareja.Jugador(3)).Name & "@" & UserList(Pareja.Jugador(4)).Name)
            
            For i = 1 To 4
                Call WarpUserChar(Pareja.Jugador(i), UserList(Pareja.Jugador(i)).flags.MapaAnterior, UserList(Pareja.Jugador(i)).flags.XAnterior, UserList(Pareja.Jugador(i)).flags.YAnterior)
                UserList(Pareja.Jugador(i)).flags.EnPareja = False
                UserList(Pareja.Jugador(i)).flags.EsperaPareja = False
                UserList(Pareja.Jugador(i)).flags.SuPareja = 0
                Pareja.Jugador(i) = 0
            Next i
            
        End If
       
    If UserList(Pareja.Jugador(3)).flags.EnPareja = True And UserList(Pareja.Jugador(4)).flags.EnPareja = True And UserList(Pareja.Jugador(3)).flags.Muerto = 1 And UserList(Pareja.Jugador(4)).flags.Muerto = 1 Then
            HayPareja = False
            UserList(Pareja.Jugador(3)).Stats.ParejasPerdidas = UserList(Pareja.Jugador(3)).Stats.ParejasPerdidas + 1
            UserList(Pareja.Jugador(4)).Stats.ParejasPerdidas = UserList(Pareja.Jugador(4)).Stats.ParejasPerdidas + 1
            
            UserList(Pareja.Jugador(1)).Stats.ParejasGanadas = UserList(Pareja.Jugador(1)).Stats.ParejasGanadas + 1
            UserList(Pareja.Jugador(2)).Stats.ParejasGanadas = UserList(Pareja.Jugador(2)).Stats.ParejasGanadas + 1
            
            If MisionesDiarias(UserList(Pareja.Jugador(1)).Misiones.NumeroMision).Tipo = 7 Then
                If UserList(Pareja.Jugador(1)).Misiones.ConteoUser < MisionesDiarias(UserList(Pareja.Jugador(1)).Misiones.NumeroMision).Cantidad Then UserList(Pareja.Jugador(1)).Misiones.ConteoUser = UserList(Pareja.Jugador(1)).Misiones.ConteoUser + 1
            End If
            
            If MisionesDiarias(UserList(Pareja.Jugador(2)).Misiones.NumeroMision).Tipo = 7 Then
                If UserList(Pareja.Jugador(2)).Misiones.ConteoUser < MisionesDiarias(UserList(Pareja.Jugador(2)).Misiones.NumeroMision).Cantidad Then UserList(Pareja.Jugador(2)).Misiones.ConteoUser = UserList(Pareja.Jugador(2)).Misiones.ConteoUser + 1
            End If
            
            Call GRANK_User_Check(Couple, UserList(Pareja.Jugador(1)).Name, UserList(Pareja.Jugador(1)).Stats.ParejasGanadas)
            Call GRANK_User_Check(Couple, UserList(Pareja.Jugador(2)).Name, UserList(Pareja.Jugador(2)).Stats.ParejasGanadas)
            UserList(Pareja.Jugador(1)).Stats.GLD = UserList(Pareja.Jugador(1)).Stats.GLD + 200000
            UserList(Pareja.Jugador(2)).Stats.GLD = UserList(Pareja.Jugador(2)).Stats.GLD + 200000
            Call SendData(SendTarget.toindex, Pareja.Jugador(1), 0, "||63@200.000")
            Call SendData(SendTarget.toindex, Pareja.Jugador(2), 0, "||63@200.000")
            SendUserGLD (Pareja.Jugador(1))
            SendUserGLD (Pareja.Jugador(2))
            Call SendData(SendTarget.ToAll, 0, 0, "||82@" & UserList(Pareja.Jugador(1)).Name & "@" & UserList(Pareja.Jugador(2)).Name)
            
            For i = 1 To 4
                Call WarpUserChar(Pareja.Jugador(i), UserList(Pareja.Jugador(i)).flags.MapaAnterior, UserList(Pareja.Jugador(i)).flags.XAnterior, UserList(Pareja.Jugador(i)).flags.YAnterior)
                UserList(Pareja.Jugador(i)).flags.EnPareja = False
                UserList(Pareja.Jugador(i)).flags.EsperaPareja = False
                UserList(Pareja.Jugador(i)).flags.SuPareja = 0
                Pareja.Jugador(i) = 0
            Next i
        End If
End If


    If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
        UserList(userindex).Stats.MuertesUser = UserList(userindex).Stats.MuertesUser + 1
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                    UserList(userindex).MascotasIndex(i) = 0
                    UserList(userindex).MascotasType(i) = 0
               End If
        End If
    Next i
    
    UserList(userindex).NroMacotas = 0
    
    If MapInfo(UserList(userindex).Pos.Map).Pk = True And UserList(userindex).Pos.Map <> 100 And UserList(userindex).Pos.Map <> 71 And UserList(userindex).Pos.Map <> 106 And UserList(userindex).Pos.Map <> 109 And UserList(userindex).Pos.Map <> 108 Then
        Call SendData(SendTarget.toindex, userindex, 0, "MUERT")
    End If
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, val(userindex), UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    Call SendUserHP(userindex)
    Call SendUserST(userindex)
    
 If UserList(userindex).EnCvc Then
 
     Dim CvcsGanados As Long
     Dim CvcsPerdidos As Long
     Dim TuHermana As Long
     
                With UserList(userindex)
                    If Guilds(.GuildIndex).GuildName = Nombre1 Then
                                modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 - 1
                                If modGuilds.UsuariosEnCvcClan1 = 0 Then
                                    Call SendData(SendTarget.ToAll, userindex, 0, "||85@" & Nombre2 & "@" & Nombre1)
                                        
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                                                                                            
                                    CvcsGanados = Guilds(GuildIndex(Nombre2)).CVCG
                                    CvcsPerdidos = Guilds(GuildIndex(Nombre1)).CVCP
                                    CvcsGanados = CvcsGanados + 1
                                    CvcsPerdidos = CvcsPerdidos + 1
                                    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Nombre2), "CVCG", CvcsGanados)
                                    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Nombre1), "CVCP", CvcsPerdidos)
                                        
                                            Dim choto As Long
                                           choto = Guilds(GuildIndex(Nombre2)).GetReputacion
                                           choto = choto + 75
                                           Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Nombre2), "REPU", choto)
                                           
                                           
                                Call GRANK_User_Check(GuildVSGuild, Nombre2, Guilds(GuildIndex(Nombre2)).CVCG)
                                Call GRANK_User_Check(GuildReputation, Nombre2, Guilds(GuildIndex(Nombre2)).GetReputacion)
    
                                End If
                            End If
                      
                
                    If Guilds(.GuildIndex).GuildName = Nombre2 Then
                        If .EnCvc = True Then
                                modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 - 1
                                If modGuilds.UsuariosEnCvcClan2 = 0 Then
                                    Call SendData(SendTarget.ToAll, userindex, 0, "||85@" & Nombre1 & "@" & Nombre2)
                                    
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                    
                                    CvcsGanados = Guilds(GuildIndex(Nombre1)).CVCG
                                    CvcsPerdidos = Guilds(GuildIndex(Nombre2)).CVCP
                                        CvcsGanados = CvcsGanados + 1
                                        CvcsPerdidos = CvcsPerdidos + 1
                                        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Nombre1), "CVCG", CvcsGanados)
                                        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Nombre2), "CVCP", CvcsPerdidos)
                                    
                                    
                                           choto = Guilds(GuildIndex(Nombre1)).GetReputacion
                                           choto = choto + 75
                                           Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & GuildIndex(Nombre2), "REPU", choto)
                                           
                                    
                                    Call GRANK_User_Check(GuildVSGuild, Nombre1, Guilds(GuildIndex(Nombre1)).CVCG)
                                    Call GRANK_User_Check(GuildReputation, Nombre1, Guilds(GuildIndex(Nombre1)).GetReputacion)
                                End If
                            End If
                        End If
                End With
            'Next ijaji
    End If
        
    
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    UserList(Atacante).flags.CantDeads = UserList(Atacante).flags.CantDeads + 1
    If UserList(Atacante).flags.TimeDeads > 0 And UserList(Atacante).flags.CantDeads = 2 Then
        Call SendData(SendTarget.ToPCArea, Atacante, UserList(Atacante).Pos.Map, "TW666")
        UserList(Atacante).flags.TimeDeads = 7
    ElseIf UserList(Atacante).flags.TimeDeads > 0 And UserList(Atacante).flags.CantDeads = 3 Then
        Call SendData(SendTarget.ToPCArea, Atacante, UserList(Atacante).Pos.Map, "TW667")
        UserList(Atacante).flags.TimeDeads = 7
    ElseIf UserList(Atacante).flags.TimeDeads > 0 And UserList(Atacante).flags.CantDeads >= 4 Then
        Call SendData(SendTarget.ToPCArea, Atacante, UserList(Atacante).Pos.Map, "TW668")
        UserList(Atacante).flags.TimeDeads = 7
    Else
        UserList(Atacante).flags.TimeDeads = 7
    End If

    If EsNewbie(Muerto) Then Exit Sub
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    If UserList(Atacante).Pos.Map = 72 Or UserList(Atacante).Pos.Map = 110 Or UserList(Atacante).Pos.Map = 99 Or UserList(Atacante).Pos.Map = 100 Or UserList(Atacante).Pos.Map = 18 Or UserList(Atacante).Pos.Map = 118 Or UserList(Atacante).Pos.Map = 101 Then Exit Sub
    
    If MisionesDiarias(UserList(Atacante).Misiones.NumeroMision).Tipo = 2 Then
        If UserList(Atacante).Misiones.ConteoUser < MisionesDiarias(UserList(Atacante).Misiones.NumeroMision).Cantidad Then UserList(Atacante).Misiones.ConteoUser = UserList(Atacante).Misiones.ConteoUser + 1
    End If
    
    If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).Name
            If UserList(Atacante).Faccion.CriminalesMatados < 65000 Then _
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
        End If
    End If

    If EsAlianza(Muerto) Then
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).Name
            If UserList(Atacante).Faccion.CiudadanosMatados < 65000 Then _
                UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
        End If
    End If
        
    If Neutral(Muerto) Then
        If UserList(Atacante).flags.LastNeutrMatado <> UserList(Muerto).Name Then _
            UserList(Atacante).flags.LastNeutrMatado = UserList(Muerto).Name
        If UserList(Atacante).Faccion.NeutralesMatados < 65000 Then _
            UserList(Atacante).Faccion.NeutralesMatados = UserList(Atacante).Faccion.NeutralesMatados + 1
    End If


End Sub
Sub TilelibreCristales(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj)

Dim Notfound As Boolean
Dim loopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
Dim ChotoX As Byte
    hayobj = False
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
        
        If loopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        ChotoX = RandomNumber(0, 4)
        
        For tY = Pos.Y - loopC - ChotoX To Pos.Y + loopC + ChotoX
            For tX = Pos.X - loopC - ChotoX To Pos.X + loopC + ChotoX
            
                If LegalPos(nPos.Map, tX, tY) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0)
                  '  If Not hayobj Then _
                   '     hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = Pos.X + loopC + ChotoX
                        tY = Pos.Y + loopC + ChotoX
                    End If
                End If
            
            Next tX
        Next tY
        loopC = loopC + 1
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub
Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim loopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
        
        If loopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - loopC To Pos.Y + loopC
            For tX = Pos.X - loopC To Pos.X + loopC
            
                If LegalPos(nPos.Map, tX, tY) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = Pos.X + loopC
                        tY = Pos.Y + loopC
                    End If
                End If
            
            Next tX
        Next tY
        
        loopC = loopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

                    If (Map = 31 Or Map = 32 Or Map = 33 Or Map = 34 Or Map = 81) And UserList(userindex).flags.Privilegios = PlayerType.User Then
                         If Not UserList(userindex).GuildIndex <> 0 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||86")
                             Exit Sub
                         End If
                    End If
                    
                    Dim loopC As Long
                    Dim UsersEnCastillo As Byte
                    UsersEnCastillo = 0
                    
                    'No pueden entrar si son más de 6.
                    If Map = 31 Then
                            For loopC = 1 To LastUser
                                If UserList(loopC).GuildIndex > 0 Then
                                    If UserList(loopC).Pos.Map = 31 And UCase$(Guilds(UserList(loopC).GuildIndex).GuildName) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) Then
                                        UsersEnCastillo = UsersEnCastillo + 1
                                    End If
                                End If
                            Next loopC
                    ElseIf Map = 32 Then
                            For loopC = 1 To LastUser
                                If UserList(loopC).GuildIndex > 0 Then
                                    If UserList(loopC).Pos.Map = 32 And UCase$(Guilds(UserList(loopC).GuildIndex).GuildName) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) Then
                                        UsersEnCastillo = UsersEnCastillo + 1
                                    End If
                                End If
                            Next loopC
                    ElseIf Map = 33 Then
                            For loopC = 1 To LastUser
                                If UserList(loopC).GuildIndex > 0 Then
                                    If UserList(loopC).Pos.Map = 33 And UCase$(Guilds(UserList(loopC).GuildIndex).GuildName) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) Then
                                        UsersEnCastillo = UsersEnCastillo + 1
                                    End If
                                End If
                            Next loopC
                    ElseIf Map = 34 Then
                            For loopC = 1 To LastUser
                                If UserList(loopC).GuildIndex > 0 Then
                                    If UserList(loopC).Pos.Map = 34 And UCase$(Guilds(UserList(loopC).GuildIndex).GuildName) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) Then
                                        UsersEnCastillo = UsersEnCastillo + 1
                                    End If
                                End If
                            Next loopC
                    End If
                            
                    If UsersEnCastillo >= 6 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||87")
                        Exit Sub
                    End If

    'Quitar el dialogo
    Call SendToUserArea(userindex, "QDL" & UserList(userindex).Char.CharIndex)
    Call SendData(SendTarget.toindex, userindex, UserList(userindex).Pos.Map, "QTDL")
    
    OldMap = UserList(userindex).Pos.Map
    OldX = UserList(userindex).Pos.X
    OldY = UserList(userindex).Pos.Y
    
    Call EraseUserChar(SendTarget.ToMap, 0, OldMap, userindex)
        
    If OldMap <> Map Then
        Call SendData(SendTarget.toindex, userindex, 0, "CM" & Map & "," & MapInfo(Map).r & "," & MapInfo(Map).g & "," & MapInfo(Map).b)
        Call SendData(SendTarget.toindex, userindex, 0, "XM" & MapInfo(Map).Music)
        Call SendData(SendTarget.toindex, userindex, 0, "N~" & MapInfo(Map).Name)
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    If MapData(Map, X, Y).userindex <> 0 Then
        UserList(userindex).Pos.X = X
        UserList(userindex).Pos.Y = Y
        UserList(userindex).Pos.Map = Map
        UserList(userindex).Pos = DamePos(UserList(userindex).Pos)
    Else
        UserList(userindex).Pos.X = X
        UserList(userindex).Pos.Y = Y
        UserList(userindex).Pos.Map = Map
    End If
    
    If userindex = GranPoder And (MapInfo(UserList(userindex).Pos.Map).Pk = False Or UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 81) Then
        Call OtorgarGranPoder(0)
        UserList(userindex).flags.GranPoder = 0
    End If
    
    If MapInfo(UserList(userindex).Pos.Map).Pk = False And UserList(userindex).flags.TimeGema > 0 Then
        UserList(userindex).flags.ActivoGema = 0
        UserList(userindex).flags.GemaActivada = ""
        UserList(userindex).flags.TimeGema = 0
        SendData SendTarget.toindex, userindex, 0, "||88"
    End If
    
    If UserList(userindex).flags.InvocoMascota = 1 Then
     Call QuitarNPC(UserList(userindex).flags.MascotinIndex)
     UserList(userindex).flags.InvocoMascota = 0
    End If
    
If UserList(userindex).GuildIndex > 0 And (UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 81) Then
    If UserList(userindex).Pos.Map = MapCastilloN And (UCase$(CastilloNorte) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiRed = 0
        UserList(userindex).flags.CastiBlue = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloN And (UCase$(CastilloNorte) <> UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiBlue = 0
        UserList(userindex).flags.CastiRed = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloS And (UCase$(CastilloSur) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiRed = 0
        UserList(userindex).flags.CastiBlue = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloS And (UCase$(CastilloSur) <> UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiBlue = 0
        UserList(userindex).flags.CastiRed = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloE And (UCase$(CastilloEste) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiRed = 0
        UserList(userindex).flags.CastiBlue = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloE And (UCase$(CastilloEste) <> UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiBlue = 0
        UserList(userindex).flags.CastiRed = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloO And (UCase$(CastilloOeste) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiRed = 0
        UserList(userindex).flags.CastiBlue = 1
    ElseIf UserList(userindex).Pos.Map = MapCastilloO And (UCase$(CastilloOeste) <> UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiBlue = 0
        UserList(userindex).flags.CastiRed = 1
    ElseIf UserList(userindex).Pos.Map = 81 And (UCase$(Fortaleza) = UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiRed = 0
        UserList(userindex).flags.CastiBlue = 1
    ElseIf UserList(userindex).Pos.Map = 81 And (UCase$(Fortaleza) <> UCase$(Guilds(UserList(userindex).GuildIndex).GuildName)) Then
        UserList(userindex).flags.CastiBlue = 0
        UserList(userindex).flags.CastiRed = 1
    End If
End If

If UserList(userindex).Pos.Map <> 31 And UserList(userindex).Pos.Map <> 32 And UserList(userindex).Pos.Map <> 33 And UserList(userindex).Pos.Map <> 34 And UserList(userindex).Pos.Map <> 81 Then
    If UserList(userindex).flags.CastiBlue = 1 Or UserList(userindex).flags.CastiRed = 1 Then
        UserList(userindex).flags.CastiBlue = 0
        UserList(userindex).flags.CastiRed = 0
    End If
End If

    Call MakeUserChar(SendTarget.ToMap, 0, Map, userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
    Call SendData(SendTarget.toindex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
    
        'Seguis invisible al pasar de mapa
    If (UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1) And (Not UserList(userindex).flags.AdminInvisible = 1) Then
        Call SendToUserArea(userindex, "NOVER" & UserList(userindex).Char.CharIndex & ",1")
    End If
    If UserList(userindex).flags.AdminInvisible = 1 Then Call SendToUserArea(userindex, "NOVER" & UserList(userindex).Char.CharIndex & ",1")
    
    If FX And UserList(userindex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_WARP)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & ",0")
    End If
    
    Call WarpMascotas(userindex)
End Sub

Sub UpdateUserMap(ByVal userindex As Integer)

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

'EnviarNoche UserIndex

On Error GoTo 0

Map = UserList(userindex).Pos.Map

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).userindex > 0 And userindex <> MapData(Map, X, Y).userindex Then
            Call MakeUserChar(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).userindex, Map, X, Y)
                If UserList(MapData(Map, X, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).userindex).flags.Oculto = 1 Then Call SendData(SendTarget.toindex, userindex, 0, "NOVER" & UserList(MapData(Map, X, Y).userindex).Char.CharIndex & ",1")
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType <> eOBJType.otArboles Then
                Call MakeObj(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                          Call Bloquear(SendTarget.toindex, userindex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                          Call Bloquear(SendTarget.toindex, userindex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                          
                          If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).PuertaDoble = 1 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X + 2, Y, MapData(Map, X + 2, Y).Blocked)
                          ElseIf ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Porton = 1 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X + 2, Y, MapData(Map, X + 2, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X - 2, Y, MapData(Map, X - 2, Y).Blocked)
                          ElseIf ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).RejaForta = 1 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X + 2, Y, MapData(Map, X + 2, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, Map, X - 2, Y, MapData(Map, X - 2, Y).Blocked)
                          End If
                          
                End If
            End If
        End If
        
    Next X
Next Y

End Sub
Sub WarpMascotas(ByVal userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(userindex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(userindex).MascotasIndex(i))
                UserList(userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||89")
        UserList(userindex).flags.EleDeAgua = 0
        UserList(userindex).flags.EleDeFuego = 0
        UserList(userindex).flags.EleDeTierra = 0
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(userindex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(userindex).Pos, False, PetRespawn(i))
            UserList(userindex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(userindex).MascotasIndex(i) = 0 Then
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
                If UserList(userindex).NroMacotas > 0 Then UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
            Npclist(UserList(userindex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(userindex).MascotasIndex(i)).Target = 0
            Npclist(UserList(userindex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    UserList(userindex).NroMacotas = NroPets

End Sub
Sub RepararMascotas(ByVal userindex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(userindex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(userindex).NroMacotas Then UserList(userindex).NroMacotas = 0

End Sub
Sub Cerrar_Usuario(ByVal userindex As Integer, Optional ByVal Tiempo As Integer = -1)

If UserList(userindex).flags.Stopped Then Exit Sub

    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(userindex).flags.Transformado = 1 Then
            Call DarCuerpoDesnudo(userindex)
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            UserList(userindex).flags.Transformado = 0
        Exit Sub
    End If
    
    If UserList(userindex).flags.UserLogged And Not UserList(userindex).Counters.Saliendo Then
        UserList(userindex).Counters.Saliendo = True
        UserList(userindex).Counters.Salir = IIf(UserList(userindex).flags.Privilegios > PlayerType.User, 0, Tiempo)
        
        
        Call SendData(SendTarget.toindex, userindex, 0, "||90@" & UserList(userindex).Counters.Salir)
    End If
    
    
    Call SalirDuelo(userindex)
    
    If UserList(userindex).flags.estado = 1 Then
        UserList(userindex).flags.estado = 0
    End If
    
If UserList(userindex).Pos.Map = 72 Then 'cambiar el 12 por numero de mapa
    If UserList(userindex).flags.EnDesafio = 1 And Desafio.Primero = userindex And Desafio.Segundo <> 0 Then 'cambiar el 12 por numero de mapa
        Call WarpUserChar(Desafio.Primero, 1, 54, 36, True) 'mapa donde lleva al creador del desafio
        Call WarpUserChar(Desafio.Segundo, 1, 54, 37, True) 'mapa donde llevar al retador
        UserList(Desafio.Primero).flags.EnDesafio = 0
        UserList(Desafio.Primero).flags.rondas = 0
        UserList(Desafio.Segundo).flags.Desafio = 0
        Call SendData(SendTarget.ToAll, 0, 0, "||91@" & UserList(Desafio.Primero).Name)
        Desafio.Primero = 0
        Desafio.Segundo = 0
    ElseIf UserList(userindex).flags.EnDesafio = 1 And Desafio.Primero = userindex And Desafio.Segundo = 0 Then 'cambiar el 12 por numero de mapa Then
        Call WarpUserChar(Desafio.Primero, 1, 54, 36, True) 'mapa donde lleva al creador del desafio
        UserList(Desafio.Primero).flags.EnDesafio = 0
        UserList(Desafio.Primero).flags.rondas = 0
        Call SendData(SendTarget.ToAll, 0, 0, "||91@" & UserList(Desafio.Primero).Name)
        Desafio.Primero = 0
        Desafio.Segundo = 0
    Else
        If UserList(userindex).flags.Desafio = 1 And Desafio.Segundo = userindex And Desafio.Primero <> 0 Then 'cambiar el 12 por numero de mapa Then
            Call WarpUserChar(Desafio.Segundo, 1, 54, 36, True) 'mapa donde lleva al retador
            UserList(Desafio.Segundo).flags.Desafio = 0
            Call SendData(SendTarget.ToAll, 0, 0, "||92@" & UserList(Desafio.Segundo).Name)
            Desafio.Segundo = 0
        Exit Sub
        End If
    Exit Sub
   End If
 Exit Sub
End If

   If UserList(userindex).flags.EnCvc = True Then
        UserList(userindex).flags.EnCvc = False
        Call WarpUserChar(userindex, 1, 54, 36, True)
    End If
            
If UserList(userindex).Pos.Map = 54 Then 'mapa de pareja
    If Pareja.Jugador(1) > 0 And Pareja.Jugador(2) > 0 And Pareja.Jugador(3) > 0 And Pareja.Jugador(4) > 0 And UserList(userindex).flags.EnPareja = True Then
        Call WarpUserChar(Pareja.Jugador(1), UserList(Pareja.Jugador(1)).flags.MapaAnterior, UserList(Pareja.Jugador(1)).flags.XAnterior, UserList(Pareja.Jugador(1)).flags.YAnterior)
        Call WarpUserChar(Pareja.Jugador(2), UserList(Pareja.Jugador(2)).flags.MapaAnterior, UserList(Pareja.Jugador(2)).flags.XAnterior, UserList(Pareja.Jugador(2)).flags.YAnterior)
        Call WarpUserChar(Pareja.Jugador(3), UserList(Pareja.Jugador(3)).flags.MapaAnterior, UserList(Pareja.Jugador(3)).flags.XAnterior, UserList(Pareja.Jugador(3)).flags.YAnterior)
        Call WarpUserChar(Pareja.Jugador(4), UserList(Pareja.Jugador(4)).flags.MapaAnterior, UserList(Pareja.Jugador(4)).flags.XAnterior, UserList(Pareja.Jugador(4)).flags.YAnterior)
        UserList(Pareja.Jugador(1)).flags.EnPareja = False
        UserList(Pareja.Jugador(1)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(1)).flags.SuPareja = 0
        UserList(Pareja.Jugador(2)).flags.EnPareja = False
        UserList(Pareja.Jugador(2)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(2)).flags.SuPareja = 0
        UserList(Pareja.Jugador(3)).flags.EnPareja = False
        UserList(Pareja.Jugador(3)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(3)).flags.SuPareja = 0
        UserList(Pareja.Jugador(4)).flags.EnPareja = False
        UserList(Pareja.Jugador(4)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(4)).flags.SuPareja = 0
        HayPareja = False
        Call SendData(SendTarget.ToAll, 0, 0, "||93")
    Exit Sub
  End If
End If

If userindex = Desafio.tPrimero Or userindex = Desafio.tSegundo Then
        Call WarpUserChar(Desafio.tPrimero, UserList(Desafio.tPrimero).flags.MapaAnterior, UserList(Desafio.tPrimero).flags.XAnterior, UserList(Desafio.tPrimero).flags.YAnterior)
        Call WarpUserChar(Desafio.tSegundo, UserList(Desafio.tSegundo).flags.MapaAnterior, UserList(Desafio.tSegundo).flags.XAnterior, UserList(Desafio.tSegundo).flags.YAnterior)
        UserList(Desafio.tPrimero).flags.tEsperaPareja = False
        UserList(Desafio.tPrimero).flags.tSuPareja = 0
        UserList(Desafio.tSegundo).flags.tEsperaPareja = False
        UserList(Desafio.tSegundo).flags.tSuPareja = 0
        Desafio.tPrimero = 0
        Desafio.tSegundo = 0
        
      If Desafio.tTercero <> 0 And Desafio.tCuarto <> 0 Then
        Call WarpUserChar(Desafio.tTercero, UserList(Desafio.tTercero).flags.MapaAnterior, UserList(Desafio.tTercero).flags.XAnterior, UserList(Desafio.tTercero).flags.YAnterior)
        Call WarpUserChar(Desafio.tCuarto, UserList(Desafio.tCuarto).flags.MapaAnterior, UserList(Desafio.tCuarto).flags.XAnterior, UserList(Desafio.tCuarto).flags.YAnterior)
        UserList(Desafio.tTercero).flags.tEsperaPareja = False
        UserList(Desafio.tTercero).flags.tSuPareja = 0
        UserList(Desafio.tCuarto).flags.tEsperaPareja = False
        UserList(Desafio.tCuarto).flags.tSuPareja = 0
        Desafio.tTercero = 0
        Desafio.tCuarto = 0
      End If
        
        Call SendData(SendTarget.ToAll, 0, 0, "||94")
        Call SendData(SendTarget.ToAll, 0, 0, "||95@" & UserList(Desafio.tPrimero).Name & "@" & UserList(Desafio.tSegundo).Name)
ElseIf userindex = Desafio.tTercero Or userindex = Desafio.tCuarto Then

        Call WarpUserChar(Desafio.tTercero, UserList(Desafio.tTercero).flags.MapaAnterior, UserList(Desafio.tTercero).flags.XAnterior, UserList(Desafio.tTercero).flags.YAnterior)
        Call WarpUserChar(Desafio.tCuarto, UserList(Desafio.tCuarto).flags.MapaAnterior, UserList(Desafio.tCuarto).flags.XAnterior, UserList(Desafio.tCuarto).flags.YAnterior)
        UserList(Desafio.tTercero).flags.tEsperaPareja = False
        UserList(Desafio.tTercero).flags.tSuPareja = 0
        UserList(Desafio.tCuarto).flags.tEsperaPareja = False
        UserList(Desafio.tCuarto).flags.tSuPareja = 0
        Desafio.tTercero = 0
        Desafio.tCuarto = 0
        
        Call SendData(SendTarget.ToAll, 0, 0, "||95@" & UserList(Desafio.tTercero).Name & "@" & UserList(Desafio.tCuarto).Name)
    Exit Sub
End If

End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).Name

If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If

End Sub
Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)

If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Pj Inexistente" & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Estadisticas de: " & Nombre & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu") & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta") & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD") & FONTTYPE_INFO)
End If
Exit Sub

End Sub
Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(SendTarget.toindex, sendIndex, 0, "N|" & CharName & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "N| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call SendData(SendTarget.toindex, sendIndex, 0, "||189@" & CharName)
End If

End Sub
Sub LlevarUsuarios()
Dim ijaji As Integer
For ijaji = 1 To LastUser
If UserList(ijaji).Pos.Map = 8 And UserList(ijaji).EnCvc = True Then
    UserList(ijaji).flags.PuedeEntrarCVC = False
    UserList(ijaji).flags.CvcBlue = 0
    UserList(ijaji).flags.CvcRed = 0
    Call WarpUserChar(ijaji, UserList(ijaji).ViejaPos.Map, UserList(ijaji).ViejaPos.X, UserList(ijaji).ViejaPos.Y, True)
    UserList(ijaji).EnCvc = False
End If
Next ijaji
End Sub
