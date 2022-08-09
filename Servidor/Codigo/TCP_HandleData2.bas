Attribute VB_Name = "TCP_HandleData2"
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

Public Sub HandleData_2(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)

CastilloNorte = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte")
CastilloSur = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur")
CastilloEste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste")
CastilloOeste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste")
Fortaleza = GetVar(IniPath & "configuracion.ini", "CASTILLO", "Fortaleza")

Dim loopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim T() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rData)
    
    Case "/TORNEO"
    
        If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
           
        UserList(userindex).Counters.TimeComandos = 5
    
            If UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or MapInfo(UserList(userindex).Pos.Map).Pk = True Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||291")
                Exit Sub
            End If
            
            If CuentaAutomatico > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||161")
                Exit Sub
            End If
            
            If Torneo_Activo = True Then
                Call Torneos_Entra(userindex)
                Exit Sub
            End If
            
            If Hay_Torneo = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||292")
                Exit Sub
            End If
       
            If CuentaTorneo > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||161")
                Exit Sub
            End If
           
            If TModalidad = "5" Then
                Call SendData(SendTarget.toindex, userindex, 0, "||292")
                Exit Sub
            End If
           
            If UserList(userindex).Stats.ELV < TNivelMinimo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||293@" & TNivelMinimo)
                Exit Sub
            End If
           
            If CParticipantes = UsuariosEnTorneo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||294@" & UsuariosEnTorneo)
                Exit Sub
            End If
            
            If TieneItemDiosEquipado(userindex) = True Then
                Call SendData(toindex, userindex, 0, "||295")
             Exit Sub
            End If
           
            If UserList(userindex).flags.EnTorneo = 0 Then
           
                Call SendData(SendTarget.toindex, userindex, 0, "||296")
                UserList(userindex).flags.EnTorneo = 1
                UsuariosEnTorneo = UsuariosEnTorneo + 1
                UserList(userindex).flags.NumTorneo = UsuariosEnTorneo
                UserList(userindex).Stats.TorneosParticipados = UserList(userindex).Stats.TorneosParticipados + 1
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||297")
              Exit Sub
            End If
        Exit Sub
           
        Case "/PARTICIPANTES"
            If Hay_Torneo = True Then
                tStr = ""
                
                For loopC = 1 To LastUser
                    'If UserList(LoopC).name <> "" Then
                    If UserList(loopC).flags.NumTorneo > 0 And UserList(loopC).Name <> "" Then
                      If UserList(loopC).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Then 'MAPA TORNEO ----
                        CronologiaParticipantes(UserList(loopC).flags.NumTorneo) = " " & UserList(loopC).flags.NumTorneo & ". " & UCase$(UserList(loopC).Name) & " [" & UserList(loopC).clase & "]"
                        'tStr = tStr & "[" & UserList(LoopC).flags.NumTorneo & "] " & UCase$(UserList(LoopC).name) & "[" & UserList(LoopC).clase & "], "
                      ElseIf loopC <= 0 Then
                        CronologiaParticipantes(UserList(loopC).flags.NumTorneo) = " " & UserList(loopC).flags.NumTorneo & ". " & UCase$(UserList(loopC).Name) & " [OFFLINE]"
                        'tStr = tStr & "[" & UserList(LoopC).flags.NumTorneo & "]" & UCase$(UserList(LoopC).name) & "[OFFLINE], "
                      Else
                        CronologiaParticipantes(UserList(loopC).flags.NumTorneo) = " " & UserList(loopC).flags.NumTorneo & ". " & UCase$(UserList(loopC).Name) & " [ELIMINADO]"
                        'tStr = tStr & "[" & UserList(LoopC).flags.NumTorneo & "]" & UCase$(UserList(LoopC).name) & "[ELIMINADO], "
                      End If
                    End If
                Next loopC
                
                For loopC = 1 To UsuariosEnTorneo
                    tStr = tStr & CronologiaParticipantes(loopC)
                Next loopC
                
                Call SendData(SendTarget.toindex, userindex, 0, "||298@" & tStr)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||299")
            End If
        Exit Sub
    
    Case "/HORDA"

        If UserList(userindex).StatusMith.EsStatus <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||300")
        Exit Sub
        End If

        If UserList(userindex).StatusMith.EligioStatus = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||301")
        Exit Sub
        End If

        If UserList(userindex).GuildIndex > 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||302")
        Exit Sub
        End If

        If UserList(userindex).Stats.ELV < 10 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||303")
        Exit Sub
        End If

        If MapInfo(UserList(userindex).Pos.Map).Pk = True Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 18 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||291")
            Exit Sub
        End If

        Call VolverCriminal(userindex)
Exit Sub

Case "/ALIANZA"

    If UserList(userindex).StatusMith.EsStatus <> 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||300")
    Exit Sub
    End If

    If UserList(userindex).StatusMith.EligioStatus = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||301")
    Exit Sub
    End If
    
    If UserList(userindex).GuildIndex > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||302")
    Exit Sub
    End If

    If MapInfo(UserList(userindex).Pos.Map).Pk = True Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 18 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||291")
        Exit Sub
    End If
    
    If UserList(userindex).Stats.ELV < 10 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||303")
    Exit Sub
    End If
    
    Call VolverCiudadano(userindex)
Exit Sub

        Case "/QUEST"
                    If UserList(userindex).flags.Muerto = 1 Then
                              Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                              Exit Sub
                    End If
                    If UserList(userindex).flags.TargetNPC = 0 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                          Exit Sub
                    End If
                    If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                              Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                              Exit Sub
                    End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Quest Then
            Call SendData(SendTarget.toindex, userindex, 0, "DAMEQUEST")
            Else
            Call SendData(SendTarget.toindex, userindex, 0, "||9")
            End If
        Exit Sub
        
        Case "/NOQUEST"
            If UserList(userindex).flags.UserNumQuest = 0 And UserList(userindex).flags.Questeando = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||304")
                Exit Sub
            End If
            
            modQuests.ResetQuest (userindex)
            Call SendData(SendTarget.toindex, userindex, 0, "||305")
        Exit Sub

        Case "/INFOSUB"
           If Hay_Subasta = False Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||308")
                Exit Sub
            End If
           
            Call SendData(SendTarget.toindex, userindex, 0, "||309@" & Subastador)
            Call SendData(SendTarget.toindex, userindex, 0, "||310@" & ObjData(objetosubastado.ObjIndex).Name)
            Call SendData(SendTarget.toindex, userindex, 0, "||311@" & cantsubasta)
           
            If NameIndex(UltimoOfertador) = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||312@" & orosubasta)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||313@" & PonerPuntos(OroOfrecido) & "@" & UltimoOfertador)
            End If
        Exit Sub
 
        Case "/SUBASTAR"
           If Hay_Subasta = True Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||314")
                Exit Sub
            End If
       
            SendData SendTarget.toindex, userindex, 0, "INITSUB"
        Exit Sub


        Case "/EMOTICONS"
            If UserList(userindex).flags.Emoticons = 1 Then
                 UserList(userindex).flags.Emoticons = 0
                Exit Sub
            Else
                UserList(userindex).flags.Emoticons = 1
            End If
        Exit Sub
        
     Case "/MSJ"
            If UserList(userindex).flags.DeseoRecibirMSJ = 1 Then
                UserList(userindex).flags.DeseoRecibirMSJ = 0
                Call SendData(SendTarget.toindex, userindex, 0, "||315")
            Else
                UserList(userindex).flags.DeseoRecibirMSJ = 1
                Call SendData(SendTarget.toindex, userindex, 0, "||316")
            End If
        Exit Sub
        
        Case "/CIUDADANIA"
                If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||9")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||10")
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> Ciudadania Then Exit Sub
            
            If UserList(userindex).StatusMith.EsStatus <> 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||317")
                Exit Sub
            End If
            
            If UserList(userindex).Pos.Map = 130 Then
            If UserList(userindex).Hogar = "Inthak" Then Exit Sub
            UserList(userindex).Hogar = "Inthak"
            End If
            
            If UserList(userindex).Pos.Map = 25 Then
            If UserList(userindex).Hogar = "Thir" Then Exit Sub
            UserList(userindex).Hogar = "Thir"
            End If
            
            If UserList(userindex).Pos.Map = 25 Then
            If UserList(userindex).Hogar = "Ruvendel" Then Exit Sub
            UserList(userindex).Hogar = "Ruvendel"
            End If
            
            Call SendData(SendTarget.toindex, userindex, 0, "||318@" & UserList(userindex).Hogar)
        Exit Sub
        
        Case "/MONTAR"
        
        If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
        If UserList(userindex).flags.TargetUser <> 0 Then Exit Sub
        
            If UserList(userindex).flags.Transformado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||319")
            Exit Sub
            End If
        
        If (Npclist(UserList(userindex).flags.TargetNPC).Numero = 156 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 157 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 158 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 181 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 182) And (Npclist(UserList(userindex).flags.TargetNPC).DueñoMascota = userindex) Then
        If UserList(userindex).flags.Montando = 0 Then
            UserList(userindex).Char.Head = 0
            If UserList(userindex).flags.Muerto = 0 Then
              If Npclist(UserList(userindex).flags.TargetNPC).Numero = 156 Then
                UserList(userindex).Char.Body = 331 'Body montura dorada
              ElseIf Npclist(UserList(userindex).flags.TargetNPC).Numero = 157 Then
                UserList(userindex).Char.Body = 330 'Body montura roja
              ElseIf Npclist(UserList(userindex).flags.TargetNPC).Numero = 158 Then
                UserList(userindex).Char.Body = 352 'Body montura oscura
              ElseIf Npclist(UserList(userindex).flags.TargetNPC).Numero = 181 Then
                UserList(userindex).Char.Body = 358 'Body caballo imperial
              ElseIf Npclist(UserList(userindex).flags.TargetNPC).Numero = 182 Then
                UserList(userindex).Char.Body = 359 'Body caballo infernal
              End If
            Else
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
            End If
            UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
            UserList(userindex).Char.ShieldAnim = NingunEscudo
            UserList(userindex).Char.WeaponAnim = NingunArma
            UserList(userindex).Char.CascoAnim = UserList(userindex).Char.CascoAnim
            UserList(userindex).flags.Montando = 1
            UserList(userindex).flags.InvocoMascota = 0
        End If
         
         
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "USM" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).flags.Montando)
            Call QuitarNPC(UserList(userindex).flags.TargetNPC)
            Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            Call SendData(toindex, userindex, 0, "EQUIT")
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||320")
        End If
        Exit Sub
       
        Case "/DESMONTAR"
            If UserList(userindex).flags.Montando = 1 Then
                Call Desmontar(userindex)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||321")
            End If
        Exit Sub
       
        Case "/QUITARMASCOTA"
        
        If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
        If UserList(userindex).flags.TargetUser <> 0 Then Exit Sub
        
        If (Npclist(UserList(userindex).flags.TargetNPC).Numero = 156 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 157 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 158 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 181 Or Npclist(UserList(userindex).flags.TargetNPC).Numero = 182) And (Npclist(UserList(userindex).flags.TargetNPC).DueñoMascota = userindex) Then
            Call QuitarNPC(UserList(userindex).flags.TargetNPC)
            UserList(userindex).flags.InvocoMascota = 0
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||320")
        End If
        Exit Sub
        
        Case "/PING"
            Call SendData(SendTarget.toindex, userindex, 0, "HOLASOYUNCIRUJA")
        Exit Sub
        
        Case "/BOTIX"
            Call ia_Spawn(eIAClase.Mago, 1, 50, 50, "Mago <TSAO>", False, True, 0)
        Exit Sub
        
        Case "/BOTIX2"
            Call ia_Spawn(eIAClase.Clerigo, 1, 51, 50, "Clerigo <TSAO>", False, True, 0)
        Exit Sub
        
        Case "/GUERRA"
        
            If HayGuerra = False Then Call SendData(toindex, userindex, 0, "||322"): Exit Sub
            If UserList(userindex).flags.PJerarquia = 0 And UserList(userindex).flags.SJerarquia = 0 And UserList(userindex).flags.TJerarquia = 0 And UserList(userindex).flags.CJerarquia = 0 Then Call SendData(toindex, userindex, 0, "||324"): Exit Sub
            If MapInfo(UserList(userindex).Pos.Map).Pk = True Then Call SendData(SendTarget.toindex, userindex, 0, "||323"): Exit Sub
            If UserList(userindex).flags.EnGuerra = 1 Then Exit Sub
            
            UserList(userindex).flags.EnGuerra = 1
            
            If HayGuerraKhalim = True Then
                If EsAlianza(userindex) Then
                    Call WarpUserChar(userindex, 1, 21, 30, True)
                ElseIf EsHorda(userindex) Then
                    Call WarpUserChar(userindex, 27, 50, 78, True)
                End If
            ElseIf HayGuerraAnvil = True Then
                If EsAlianza(userindex) Then
                    Call WarpUserChar(userindex, 29, 46, 68, True)
                ElseIf EsHorda(userindex) Then
                    Call WarpUserChar(userindex, 41, 50, 13, True)
                End If
            End If
            
            Call SendData(SendTarget.toindex, userindex, 0, "||325")
        Exit Sub
        
        Case "/CIRUJIA"
        If UserList(userindex).flags.Muerto = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||3")
                        Exit Sub
                    
                    '¿El target es un NPC valido?
                    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
                        If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||158")
                            Exit Sub
                        End If
                        End If
                        
        If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.cirujano Then
            SendData SendTarget.toindex, userindex, 0, "CIRUJA" & UserList(userindex).Raza & "," & UserList(userindex).Genero
        End If
        Exit Sub
        
        Case "/ADVERTENCIAS"
            Dim Cant As Byte
            Cant = val(GetVar(CharPath & UserList(userindex).Name & ".chr", "PENAS", "Cant"))
             
            Dim Pena As String
         
            If Cant = 0 Then
            SendData SendTarget.toindex, userindex, 0, "||326"
            Exit Sub
            End If
             
             
            Dim p As Integer
            For p = 1 To Cant
            Pena = GetVar(CharPath & UserList(userindex).Name & ".chr", "PENAS", "P" & p)
            SendData SendTarget.toindex, userindex, 0, "||327@" & p & "@" & Pena
            Next p
         
        Exit Sub
        
        Case "/VOTAR"
            If HayEncuesta = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||328")
             Exit Sub
            ElseIf UserList(userindex).Stats.ELV < LvlEncuesta Then
                Call SendData(SendTarget.toindex, userindex, 0, "||329@" & LvlEncuesta)
             Exit Sub
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "VOT" & Opciones(1) & "," & Opciones(2) & "," & Opciones(3) & "," & Opciones(4) & "," & Opciones(5) & "," & Encuesta)
             Exit Sub
            End If
        Exit Sub
        
        Case "/RESULTADOS"
            Dim Total As Integer
            
            If HayEncuesta = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||328")
            Exit Sub
            End If
            
            Total = Votos(1) + Votos(2) + Votos(3) + Votos(4) + Votos(5)
            
            If opcion(1) = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||330@" & Opciones(1) & "@" & Votos(1) & "@" & (Votos(1) * 100) / Total)
            End If
            
            If opcion(2) = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||330@" & Opciones(2) & "@" & Votos(2) & "@" & (Votos(2) * 100) / Total)
            End If
            
            If opcion(3) = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||330@" & Opciones(3) & "@" & Votos(3) & "@" & (Votos(3) * 100) / Total)
            End If
            
            If opcion(4) = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||330@" & Opciones(4) & "@" & Votos(4) & "@" & (Votos(4) * 100) / Total)
            End If
            
            If opcion(5) = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||330@" & Opciones(5) & "@" & Votos(5) & "@" & (Votos(5) * 100) / Total)
            End If
            
            Call SendData(SendTarget.toindex, userindex, 0, "||331@" & Total)
        Exit Sub
        
        
        Case "/CLAN"
        Dim GuildIndex As Integer
        
            If Not UserList(userindex).GuildIndex >= 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||120")
             Exit Sub
            End If
            
            Dim Usersclan As Integer
            For Usersclan = 1 To LastUser
            
        If UserList(Usersclan).GuildIndex >= 1 Then
         If Guilds(UserList(Usersclan).GuildIndex).GuildName = Guilds(UserList(userindex).GuildIndex).GuildName Then
            
            GuildIndex = UserList(Usersclan).GuildIndex
            
            If m_EsGuildLeader(UserList(Usersclan).Name, GuildIndex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "N|" & UserList(Usersclan).Name & "- Mapa:" & UserList(Usersclan).Pos.Map & "(" & MapInfo(UserList(Usersclan).Pos.Map).Name & ")" & " X:" & UserList(Usersclan).Pos.X & " Y:" & UserList(Usersclan).Pos.Y & "~90~90~100~0~0")
            ElseIf m_EsGuildSubLeader1(UserList(Usersclan).Name, GuildIndex) Or m_EsGuildSubLeader2(UserList(Usersclan).Name, GuildIndex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "N|" & UserList(Usersclan).Name & "- Mapa:" & UserList(Usersclan).Pos.Map & "(" & MapInfo(UserList(Usersclan).Pos.Map).Name & ")" & " X:" & UserList(Usersclan).Pos.X & " Y:" & UserList(Usersclan).Pos.Y & "~0~98~110")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "N|" & UserList(Usersclan).Name & "- Mapa:" & UserList(Usersclan).Pos.Map & "(" & MapInfo(UserList(Usersclan).Pos.Map).Name & ")" & " X:" & UserList(Usersclan).Pos.X & " Y:" & UserList(Usersclan).Pos.Y & "~255~0~0")
            End If
            
         End If
        End If
        
        Next Usersclan
        Exit Sub
        
        Case "/BONIFICADORES"
        
            If UCase$(UserList(userindex).clase) = "BARDO" Or UCase$(UserList(userindex).clase) = "CLERIGO" Or UCase$(UserList(userindex).clase) = "GUERRERO" Or UCase$(UserList(userindex).clase) = "CAZADOR" Or UCase$(UserList(userindex).clase) = "PALADIN" Or UCase$(UserList(userindex).clase) = "ASESINO" Or UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "DRUIDA" Then
                If UserList(userindex).Stats.ELV >= 20 And UserList(userindex).Bon1 = "Ninguno/No elegido." Then
                  Call SendData(SendTarget.toindex, userindex, 0, "99" & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel1Opcion1") & "," & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel1Opcion2"))
                  Exit Sub
                End If
                
                If UserList(userindex).Stats.ELV >= 40 And UserList(userindex).Bon2 = "Ninguno/No elegido." Then
                  Call SendData(SendTarget.toindex, userindex, 0, "99" & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion1") & "," & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion2"))
                  Exit Sub
                End If
                
                If UserList(userindex).Stats.ELV >= 60 And UserList(userindex).Bon3 = "Ninguno/No elegido." Then
                  Call SendData(SendTarget.toindex, userindex, 0, "99" & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel3Opcion1") & "," & GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel3Opcion2"))
                  Exit Sub
                End If
            End If
        
        Exit Sub
        
        Case "/EDITLVL"
            If UserList(userindex).Stats.ELV < 50 Then
                For loopC = 1 To (50 - UserList(userindex).Stats.ELV)
                    UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
                    Call CheckUserLevel(userindex)
                Next loopC
            End If
        Exit Sub
        
        Case "/RRPERSONAJE"
            UserList(userindex).Stats.MET = 1
            Dim MiInt As Long
            MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)
            
            UserList(userindex).Stats.MaxHP = 21
            UserList(userindex).Stats.MinHP = 21
            
            UserList(userindex).Stats.FIT = 1
            
            
            MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
            If MiInt = 1 Then MiInt = 2
            
            UserList(userindex).Stats.MaxSta = 20 * MiInt
            UserList(userindex).Stats.MinSta = 20 * MiInt
            
            
            UserList(userindex).Stats.MaxAGU = 100
            UserList(userindex).Stats.MinAGU = 100
            
            UserList(userindex).Stats.MaxHam = 100
            UserList(userindex).Stats.MinHam = 100
            
            UserList(userindex).InventorySlots = 21
            
            
            '<-----------------MANA----------------------->
            If UCase$(UserList(userindex).clase) = "MAGO" Or UCase$(UserList(userindex).clase) = "NIGROMANTE" Then
                MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 3
                UserList(userindex).Stats.MaxMAN = 100 + MiInt
                UserList(userindex).Stats.MinMAN = 100 + MiInt
            ElseIf UCase$(UserList(userindex).clase) = "CLERIGO" Or UCase$(UserList(userindex).clase) = "DRUIDA" _
                Or UCase$(UserList(userindex).clase) = "BARDO" Or UCase$(UserList(userindex).clase) = "ASESINO" Then
                    MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 4
                    UserList(userindex).Stats.MaxMAN = 50
                    UserList(userindex).Stats.MinMAN = 50
            Else
                UserList(userindex).Stats.MaxMAN = 0
                UserList(userindex).Stats.MinMAN = 0
            End If
            
            UserList(userindex).Stats.MaxHIT = 2
            UserList(userindex).Stats.MinHIT = 1
            
            UserList(userindex).Stats.Exp = 0
            UserList(userindex).Stats.ELU = 300
            UserList(userindex).Stats.ELV = 1
                        
            SendUserStats (userindex)
        Exit Sub
        
        Case "/PARTICIPAR"
                Call Aram_Ingresar(userindex)
        Exit Sub
        
        Case "/ONLINE"
        Dim a As Integer
        
                    'No se envia más la lista completa de usuarios
                    n = 0
                    For loopC = 1 To LastUser
                        If UserList(loopC).Name <> "" Then
                            n = n + 1
                        End If
                    Next loopC
                     
        Call SendData(SendTarget.toindex, userindex, 0, "||332@" & n + BOnlines & "@" & recordusuarios)
         
        a = 0
                    For loopC = 1 To LastUser
                    
                        If UserList(userindex).flags.Privilegios > PlayerType.User Then
                           If UserList(loopC).Name <> "" Then
                            Dim txdasd As String
                            txdasd = txdasd & "" & UserList(loopC).Name & ", "
                           End If
                        End If
                    
                        If UserList(loopC).Name <> "" Then
                            If UserList(loopC).StatusMith.EsStatus = 1 Or UserList(loopC).StatusMith.EsStatus = 3 Then
                            a = a + 1
                        End If
                        End If
                    Next loopC
                    
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||333@" & txdasd)
                    End If
         
        Call SendData(SendTarget.toindex, userindex, 0, "||334@" & a)
         
         
        a = 0
                    For loopC = 1 To LastUser
                        If UserList(loopC).Name <> "" Then
                            If UserList(loopC).StatusMith.EsStatus = 2 Or UserList(loopC).StatusMith.EsStatus = 4 Then
                            a = a + 1
                        End If
                        End If
                    Next loopC
         
        Call SendData(SendTarget.toindex, userindex, 0, "||335@" & a)
         
        a = 0
                    For loopC = 1 To LastUser
                        If UserList(loopC).Name <> "" Then
                            If UserList(loopC).StatusMith.EsStatus = 0 Or UserList(loopC).StatusMith.EsStatus = 20 Then
                            a = a + 1
                        End If
                        End If
                    Next loopC
         
        Call SendData(SendTarget.toindex, userindex, 0, "||336@" & a + BOnlines)
        Exit Sub

Case "/CERRARCLAN"

        If modGuilds.m_EsGuildSubLeader1(UserList(userindex).Name, UserList(userindex).GuildIndex) = True Or modGuilds.m_EsGuildSubLeader2(UserList(userindex).Name, UserList(userindex).GuildIndex) = True Then
            tInt = m_EcharMiembroDeClan(userindex, UserList(userindex).Name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||337@" & UserList(userindex).Name)
                Call SendData(SendTarget.toindex, userindex, 0, "||338")
            End If
        Exit Sub
    End If
    
    If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||323")
    Exit Sub
    End If
                
If Not UserList(userindex).GuildIndex >= 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||120")
Exit Sub
End If
 
If Not modGuilds.m_EsGuildLeader(UserList(userindex).Name, UserList(userindex).GuildIndex) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||339")
Exit Sub
End If
 
If Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros > 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||340")
Exit Sub
End If
 
Call SendData(SendTarget.ToAll, 0, 0, "||341@" & Guilds(UserList(userindex).GuildIndex).GuildName)
Call WriteVar(CharPath & Guilds(UserList(userindex).GuildIndex).Fundador & ".chr", "GUILD", "GUILDINDEX", "0")
Call WriteVar(CharPath & Guilds(UserList(userindex).GuildIndex).Fundador & ".chr", "GUILD", "AspiranteA", "0")
Call WriteVar(CharPath & Guilds(UserList(userindex).GuildIndex).Fundador & ".chr", "GUILD", "Miembro", "0")
Call Kill(App.Path & "\guilds\" & Guilds(UserList(userindex).GuildIndex).GuildName & "-members.mem")
Call Kill(App.Path & "\guilds\" & Guilds(UserList(userindex).GuildIndex).GuildName & "-solicitudes.sol")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Founder", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "GuildName", "cerrado" & UserList(userindex).GuildIndex)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Date", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Antifaccion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Alineacion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex1", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex2", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex3", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex4", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex5", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex6", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex7", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex8", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Desc", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "GuildNews", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "URL", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CVCP", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CVCG", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CASTIS", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "REPU", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", vbNullString)
'Call Guilds(UserList(UserIndex).GuildIndex).DesConectarMiembro(UserIndex)
UserList(userindex).GuildIndex = 0
Call CloseSocket(userindex)
Exit Sub

    Case "/REGRESAR"
            
        If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
        
        If UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||342")
            Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 18 Or UserList(userindex).Pos.Map = 166 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||343")
            Exit Sub
        End If
        
        If UserList(userindex).Stats.ELV < 10 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||344")
            Exit Sub
        End If
           
        If MapInfo(UserList(userindex).Pos.Map).Pk = False Then
            Call SendData(SendTarget.toindex, userindex, 0, "||345")
            Exit Sub
        End If
        
        If UserList(userindex).flags.EsPremium = 1 And UserList(userindex).flags.Muerto = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||346")
            UserList(userindex).UserPremiumMap = 0
            UserList(userindex).Counters.TransportePremium = 5
            Exit Sub
        End If
        
        If (UserList(userindex).Pos.Map = 121 Or UserList(userindex).Pos.Map = 94 Or UserList(userindex).Pos.Map = 95 Or UserList(userindex).Pos.Map = 19 Or UserList(userindex).Pos.Map = 10 Or UserList(userindex).Pos.Map = 36 Or UserList(userindex).Pos.Map = 43) Then
            If TieneObjetos(1270, 1, userindex) Then
                Call QuitarObjetos(1270, 1, userindex)
            Else
                 Call SendData(SendTarget.toindex, userindex, 0, "||347")
             Exit Sub
            End If
        End If
           
           
                If UserList(userindex).flags.Muerto = 0 Then _
                    Call UserDie(userindex)
                    
                    If userindex = GranPoder Then
                        GranPoder = 0
                        UserList(userindex).flags.GranPoder = 0
                        SendUserVariant (userindex)
                        Call OtorgarGranPoder(0)
                    End If
                   
                       Call WarpUserChar(userindex, 1, 49, 35, True)
                       Call SendData(toindex, userindex, 0, "||348")
            Exit Sub

        Case "/SALIR"
            
            If UserList(userindex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||349")
                Exit Sub
            End If
            
            If UserList(userindex).Pos.Map = 141 Or UserList(userindex).Pos.Map = 120 Or UserList(userindex).Pos.Map = 121 Or UserList(userindex).Pos.Map = 94 Or UserList(userindex).Pos.Map = 95 Or UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 81 Then Exit Sub
            
            If userindex = GranPoder Then
                GranPoder = 0
                UserList(userindex).flags.GranPoder = 0
                UserList(userindex).flags.TeniaElDon = 1
                SendUserVariant (userindex)
                Call OtorgarGranPoder(0)
            End If
            
            If UserList(userindex).flags.EnDuelo = True Then Call SalirDuelo(userindex)
            
            'casted - pareja 2vs2
            If UserList(userindex).Pos.Map = 72 Then 'cambiar el 12 por numero de mapa
                If UserList(userindex).flags.EnDesafio = 1 And Desafio.Primero = userindex And Desafio.Segundo <> 0 Then 'cambiar el 12 por numero de mapa
                    Call WarpUserChar(Desafio.Primero, 1, 49, 35, True) 'mapa donde lleva al creador del desafio
                    Call WarpUserChar(Desafio.Segundo, 1, 49, 36, True) 'mapa donde llevar al retador
                    UserList(Desafio.Primero).flags.EnDesafio = 0
                    UserList(Desafio.Primero).flags.rondas = 0
                    UserList(Desafio.Segundo).flags.Desafio = 0
                    Call SendData(SendTarget.ToAll, 0, 0, "||350@" & UserList(Desafio.Primero).Name)
                    Desafio.Primero = 0
                    Desafio.Segundo = 0
                ElseIf UserList(userindex).flags.EnDesafio = 1 And Desafio.Primero = userindex And Desafio.Segundo = 0 Then 'cambiar el 12 por numero de mapa Then
                    Call WarpUserChar(Desafio.Primero, 1, 49, 37, True)
                    UserList(Desafio.Primero).flags.EnDesafio = 0
                    UserList(Desafio.Primero).flags.rondas = 0
                    Call SendData(SendTarget.ToAll, 0, 0, "||350@" & UserList(Desafio.Primero).Name)
                    Desafio.Primero = 0
                Else
                    If UserList(userindex).flags.Desafio = 1 And Desafio.Segundo = userindex And Desafio.Primero <> 0 Then 'cambiar el 12 por numero de mapa Then
                    Call WarpUserChar(Desafio.Segundo, 1, 49, 34, True) 'mapa donde lleva al retador
                    UserList(Desafio.Segundo).flags.Desafio = 0
                    Call SendData(SendTarget.ToAll, 0, 0, "||351@" & UserList(Desafio.Segundo).Name)
                    Desafio.Segundo = 0
                    Exit Sub
                End If
              Exit Sub
            End If
          Exit Sub
        End If


            If UserList(userindex).flags.Privilegios = PlayerType.User Then
                If UserList(userindex).Pos.Map = 94 Or UserList(userindex).Pos.Map = 95 Or UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 18 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 101 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||317")
                 Exit Sub
                End If
            End If
            
            If UserList(userindex).flags.InvocoMascota = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||352")
            Exit Sub
            End If
             
            If UserList(userindex).flags.Montando = 1 Then
                Call Desmontar(userindex)
            End If
            
            If UserList(userindex).flags.Transformado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||353")
               Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).OrigChar.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
               UserList(userindex).flags.Transformado = 0
           End If
        
            ''mato los comercios seguros
            If UserList(userindex).cComercio.cComercia = True Then
               comCancelar userindex
            End If
            
            Call Cerrar_Usuario(userindex)
            Exit Sub
            
            
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(userindex, UserList(userindex).Name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||338")
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||354@" & UserList(userindex).Name)
                Call WarpUserChar(tInt, UserList(tInt).Pos.Map, UserList(tInt).Pos.X, UserList(tInt).Pos.Y, True)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||120")
            End If
            Exit Sub
            
            Case "/RENUNCIAR"

            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||9")
                  Exit Sub
            End If
            
            If UserList(userindex).GuildIndex > 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||302")
            Exit Sub
            End If
            
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||10")
                      Exit Sub
            End If
            
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> Renunciar Then Exit Sub

                If UserList(userindex).StatusMith.EsStatus = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||324")
                Else
                    UserList(userindex).StatusMith.EsStatus = 0
                    UserList(userindex).StatusMith.EligioStatus = 1
                    UserList(userindex).Faccion.CiudadanosMatados = 0
                    UserList(userindex).Faccion.CriminalesMatados = 0
                    UserList(userindex).Faccion.NeutralesMatados = 0
                    UserList(userindex).Stats.UsuariosMatados = 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||355")
                    Call SendUserStatux(userindex)
                    If UserList(userindex).Faccion.ArmadaReal = 1 Then
                    Call ExpulsarFaccionReal(userindex)
                    ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
                    Call ExpulsarFaccionCaos(userindex)
                End If
                End If
        Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||3")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||9")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||10")
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
            
        Case "/NOBLE"
        
            If Not TieneObjetos(1073, 1, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||356")
                Exit Sub
            ElseIf Not TieneObjetos(1074, 1, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||356")
                Exit Sub
            ElseIf Not TieneObjetos(1075, 1, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||356")
                Exit Sub
            ElseIf Not TieneObjetos(1076, 1, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||356")
                Exit Sub
            ElseIf Not TieneObjetos(1077, 1, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||356")
                Exit Sub
            End If
           
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
           
            If UserList(userindex).flags.EsNoble = 1 Then Exit Sub
           
            UserList(userindex).flags.EsNoble = 1
            Call QuitarObjetos(1073, 1, userindex)
            Call QuitarObjetos(1074, 1, userindex)
            Call QuitarObjetos(1075, 1, userindex)
            Call QuitarObjetos(1076, 1, userindex)
            Call QuitarObjetos(1077, 1, userindex)
            
        Dim j As Integer
    If Not TieneHechizo(46, userindex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
            
        If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||181")
        Else
            UserList(userindex).Stats.UserHechizos(j) = 46
            Call UpdateUserHechizos(False, userindex, CByte(j))
        End If
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||182")
    End If

        UserList(userindex).flags.EsNoble = 1
        Call SendData(SendTarget.toindex, userindex, 0, "||357@" & UserList(userindex).clase)
        Call SendData(SendTarget.ToAll, 0, 0, "||358@" & UserList(userindex).Name)
        SendUserVariant (userindex)
    Exit Sub
            
    Case "/DESENTERRAR"
     
       If (UserList(userindex).Pos.Map <> MapaTesoroMap) Or (UserList(userindex).Pos.X <> MapaTesoroX) Or (UserList(userindex).Pos.Y <> MapaTesoroY) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||359")
        Exit Sub
        End If
       
      If Not TieneObjetos(LlaveTesoro, 1, userindex) Then 'objeto
        Call SendData(SendTarget.toindex, userindex, 0, "||360")
        Exit Sub
        End If
        
        Call QuitarObjetos(LlaveTesoro, 1, userindex)
        UserList(userindex).flags.Desenterrando = 1
        TesoroContando = True
        TiempoTesoro = 30
        Call MakeObj(SendTarget.ToMap, 0, MapaTesoroMap, ObjetoT, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
        Call SendData(SendTarget.toindex, userindex, 0, "||361")
    Exit Sub
             
 Case "/PODER"
        If GranPoder = 0 Then OtorgarGranPoder (0)
        
        If GranPoder > 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||362@" & UserList(GranPoder).Name & "@" & UserList(GranPoder).Pos.Map)
        End If
        
        If GranPoder = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||363")
        End If
    
    Exit Sub
    
     Case "/SICV"
    
    If Not UserList(userindex).GuildIndex >= 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||120")
    Exit Sub
    End If
    
    If UserList(userindex).Pos.Map = 141 Then Exit Sub
    
    If CvcFunciona = True Then
        SendData SendTarget.toindex, userindex, 0, "||364"
    Exit Sub
    End If
    
    Nombre1 = Guilds(UserList(userindex).GuildIndex).GuildName
    Nombre2 = Guilds(UserList(userindex).GuildIndex).ClanPideDesafio

If Nombre1 = Nombre2 Then Exit Sub

        Dim je As Integer
        Dim pra As Long
        Dim j3 As Integer
        'Dim a As Long
        Dim b As Long
        Dim dam As Long
        Dim dam2 As Long
        
        a = 0
        b = 0
        
        If Guilds(UserList(userindex).GuildIndex).TieneParaDesafiar = True Then
        For dam = 1 To LastUser
        If UserList(dam).GuildIndex > 0 Then
            If Guilds(UserList(dam).GuildIndex).GuildName = Nombre1 And UserList(dam).flags.SeguroCVC = True And UserList(dam).Counters.Pena = 0 And UserList(dam).Pos.Map <> 72 And UserList(dam).Pos.Map <> 78 And UserList(dam).Pos.Map <> 54 And UserList(dam).Pos.Map <> 100 And UserList(dam).Pos.Map <> 99 And UserList(dam).Pos.Map <> 101 And UserList(dam).Pos.Map <> 18 And UserList(dam).Pos.Map <> 19 And UserList(dam).Pos.Map <> 141 And TieneItemDiosEquipado(dam) = False Then
                UserList(dam).flags.PuedeEntrarCVC = True
                a = a + 1
            End If
        End If
        
        If UserList(dam).GuildIndex > 0 Then
            If Guilds(UserList(dam).GuildIndex).GuildName = Nombre2 And UserList(dam).flags.SeguroCVC = True And UserList(dam).Counters.Pena = 0 And UserList(dam).Pos.Map <> 72 And UserList(dam).Pos.Map <> 78 And UserList(dam).Pos.Map <> 54 And UserList(dam).Pos.Map <> 100 And UserList(dam).Pos.Map <> 99 And UserList(dam).Pos.Map <> 101 And UserList(dam).Pos.Map <> 18 And UserList(dam).Pos.Map <> 19 And UserList(dam).Pos.Map <> 141 And TieneItemDiosEquipado(dam) = False Then
                UserList(dam).flags.PuedeEntrarCVC = True
                b = b + 1
            End If
        End If
        Next dam
        
        If a < 1 Then
           SendData SendTarget.toindex, userindex, 0, "||365"
        Exit Sub
        End If
        
        If b < 1 Then
           SendData SendTarget.toindex, userindex, 0, "||366"
        Exit Sub
        End If
        
        For dam2 = 1 To LastUser
        If UserList(dam2).GuildIndex > 0 Then
        If Guilds(UserList(dam2).GuildIndex).GuildName = Nombre1 Then
           If modGuilds.m_EsGuildLeader(UserList(dam2).Name, UserList(dam2).GuildIndex) Then
             If UserList(dam2).Stats.GLD > 100000 Then
                    UserList(dam2).Stats.GLD = UserList(dam2).Stats.GLD - 100000
                    Call SendUserGLD(dam2)
                Else
                    SendData SendTarget.toindex, userindex, 0, "||215@100.000"
                Exit Sub
              End If
             End If
            End If
        End If
        
   If UserList(dam2).GuildIndex > 0 Then
                If Guilds(UserList(dam2).GuildIndex).GuildName = Nombre2 Then
            If modGuilds.m_EsGuildLeader(UserList(dam2).Name, UserList(dam2).GuildIndex) Then
                If UserList(dam2).Stats.GLD > 100000 Then
                    UserList(dam2).Stats.GLD = UserList(dam2).Stats.GLD - 100000
                    Call SendUserGLD(dam2)
                Else
                    SendData SendTarget.toindex, userindex, 0, "||367"
                Exit Sub
              End If
             End If
            End If
         End If
        Next dam2
        
        modGuilds.UsuariosEnCvcClan2 = 0
        modGuilds.UsuariosEnCvcClan1 = 0
        
        SendData SendTarget.ToAll, userindex, 0, "||368@" & Guilds(UserList(userindex).GuildIndex).ClanPideDesafio & "@" & Guilds(UserList(userindex).GuildIndex).GuildName
        CvcFunciona = True
           For i = 1 To LastUser
            If UserList(i).GuildIndex <> 0 Then
            If UserList(i).flags.SeguroCVC = True Then
            If Guilds(UserList(i).GuildIndex).GuildName = Nombre1 And UserList(i).flags.PuedeEntrarCVC And modGuilds.UsuariosEnCvcClan1 < b Then
   '''''''''''         'Si viene el clan n°1
                If UserList(i).flags.Muerto = 1 Then
                    Call RevivirUsuario(i)
                    UserList(i).Stats.MinHP = UserList(i).Stats.MaxHP
                    SendUserHP (i)
                End If
   
                modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 + 1
                UserList(i).ViejaPos.Map = UserList(i).Pos.Map
                UserList(i).ViejaPos.X = UserList(i).Pos.X
                UserList(i).ViejaPos.Y = UserList(i).Pos.Y
                WarpUserChar i, 8, RandomNumber(37, 48), RandomNumber(70, 77), True
                UserList(i).EnCvc = True
                UserList(i).flags.CvcBlue = 1
                Debug.Print Nombre1 & " entra con " & modGuilds.UsuariosEnCvcClan1 & " usuarios al cvc"
            End If
            If Guilds(UserList(i).GuildIndex).GuildName = Nombre2 And UserList(i).flags.PuedeEntrarCVC And modGuilds.UsuariosEnCvcClan2 < a Then
 '''''''''''''''           'Si tambien viene el 2°
                If UserList(i).flags.Muerto = 1 Then
                    Call RevivirUsuario(i)
                    UserList(i).Stats.MinHP = UserList(i).Stats.MaxHP
                    SendUserHP (i)
                End If
 
                modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 + 1
                UserList(i).ViejaPos.Map = UserList(i).Pos.Map
                UserList(i).ViejaPos.X = UserList(i).Pos.X
                UserList(i).ViejaPos.Y = UserList(i).Pos.Y
                WarpUserChar i, 8, RandomNumber(75, 86), RandomNumber(35, 45), True
                UserList(i).EnCvc = True
                UserList(i).flags.CvcRed = 1
              Debug.Print Nombre2 & " entra con " & modGuilds.UsuariosEnCvcClan2 & " usuarios al cvc"
            End If
        End If
        End If

        Next i
        
        Guilds(UserList(userindex).GuildIndex).TieneParaDesafiar = False
        Guilds(UserList(userindex).GuildIndex).ClanPideDesafio = ""
        
        

        
        Else
        SendData SendTarget.toindex, userindex, 0, "||369"
        Exit Sub
        End If
    Exit Sub
        
    Case "/NCVC"
            UserList(userindex).flags.SeguroCVC = False
            SendData SendTarget.toindex, userindex, 0, "||370"
            Call SendData(SendTarget.toindex, userindex, 0, "SEGCVCOFF")
    Exit Sub
    
    Case "/SCVC"
            UserList(userindex).flags.SeguroCVC = True
            SendData SendTarget.toindex, userindex, 0, "||371"
            Call SendData(SendTarget.toindex, userindex, 0, "SEGCVCON")
    Exit Sub
    
    
        '########## SISTEMA DE PARTY - FER ###########
        Case "/NUEVAPARTY"
        
             If UserList(userindex).flags.PartyIndex <> 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||372")
               Exit Sub
             End If
            
            UltimaParty = UltimaParty + 1
            UserList(userindex).flags.PartyIndex = UltimaParty
            UserList(userindex).flags.PartyLider = 1
            Call SendData(SendTarget.toindex, userindex, 0, "||373")
        Exit Sub
        
        Case "/PARTY"
            If UserList(userindex).flags.TargetUser > 0 Then
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||374")
                    Exit Sub
                End If
        
                If UserList(userindex).flags.TargetUser = userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||375")
                    Exit Sub
                End If
                
                If UserList(userindex).flags.PartyIndex = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||376")
                    Exit Sub
                End If
                
                If UserList(userindex).flags.PartyLider = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||377")
                    Exit Sub
                End If
                
                If Distancia(UserList(UserList(userindex).flags.TargetUser).Pos, UserList(userindex).Pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||10")
                    Exit Sub
                End If
                
                If UserList(UserList(userindex).flags.TargetUser).flags.PartySolicitud = UserList(userindex).flags.PartyIndex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||378")
                  Exit Sub
                End If
                
                Dim cantmiemb As Byte
                For i = 1 To LastUser
                 If UserList(i).flags.PartyIndex = UserList(userindex).flags.PartyIndex Then
                    cantmiemb = cantmiemb + 1
                 End If
                Next i
                
                If cantmiemb >= 5 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||379")
                  Exit Sub
                End If
                                
                If UserList(UserList(userindex).flags.TargetUser).PartyIndex <> 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||380")
                  Exit Sub
                End If

                UserList(UserList(userindex).flags.TargetUser).flags.PartySolicitud = UserList(userindex).flags.PartyIndex
                Call SendData(SendTarget.toindex, UserList(userindex).flags.TargetUser, 0, "||381@" & UserList(userindex).Name)
                Call SendData(SendTarget.toindex, userindex, 0, "||382@" & UserList(UserList(userindex).flags.TargetUser).Name)
                
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||9")
            End If
        Exit Sub
        
        Case "/ACEPTAR"
        
         If UserList(userindex).flags.PartySolicitud = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||383")
            Exit Sub
         End If
         
         If UserList(userindex).flags.PartyIndex <> 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||384")
            Exit Sub
         End If
         
                For i = 1 To LastUser
                 If UserList(i).flags.PartyIndex = UserList(userindex).flags.PartySolicitud Then
                    cantmiemb = cantmiemb + 1
                 End If
                Next i
                
                If cantmiemb >= 5 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||385")
                  Exit Sub
                End If
         
            UserList(userindex).flags.PartyIndex = UserList(userindex).flags.PartySolicitud
            UserList(userindex).flags.PartyLider = 0
        
            Call SendData(SendTarget.ToPartyArea, userindex, 0, "||386@" & UserList(userindex).Name)
        
        Exit Sub
        
        Case "/CANCELAR"
        
            If UserList(userindex).flags.PartySolicitud <> 0 Then
               UserList(userindex).flags.PartySolicitud = 0
               Call SendData(SendTarget.toindex, userindex, 0, "||387")
            Else
                   Call SendData(SendTarget.toindex, userindex, 0, "||383")
               Exit Sub
            End If
         
            If UserList(userindex).flags.PartyIndex <> 0 Then
            
               Call SendData(SendTarget.ToPartyArea, userindex, 0, "||388@" & UserList(userindex).Name)
            
               UserList(userindex).flags.PartyIndex = 0
               UserList(userindex).flags.PartyLider = 0
               UserList(userindex).flags.PartySolicitud = 0
               
            Else
                   Call SendData(SendTarget.toindex, userindex, 0, "||389")
               Exit Sub
            End If
        
        Exit Sub
        
        Case "/FINPARTY"
         If UserList(userindex).flags.PartyIndex = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||389")
            Exit Sub
         ElseIf UserList(userindex).flags.PartyLider = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||377")
            Exit Sub
         End If
         
         Call SendData(SendTarget.ToPartyArea, userindex, 0, "||390")
         
         For i = 1 To LastUser
            If UserList(i).flags.PartyIndex = UserList(userindex).flags.PartyIndex Or UserList(i).flags.PartySolicitud = UserList(userindex).flags.PartyIndex Then
                UserList(i).flags.PartyIndex = 0
                UserList(i).flags.PartyLider = 0
                UserList(i).flags.PartySolicitud = 0
            End If
         Next i
         
         UserList(userindex).flags.PartyIndex = 0
         UserList(userindex).flags.PartyLider = 0
         UserList(userindex).flags.PartySolicitud = 0
        Exit Sub
        
        Case "/PINFO"
            If UserList(userindex).flags.PartyIndex = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||377")
             Exit Sub
            End If
            
            Call SendData(SendTarget.toindex, userindex, 0, "||391")
            
            For i = 1 To LastUser
               If UserList(i).flags.PartyIndex = UserList(userindex).flags.PartyIndex Then
                   If UserList(i).Name <> "" Then
                       Call SendData(SendTarget.toindex, userindex, 0, "||392@" & UserList(i).Name & "@" & UserList(i).Stats.ELV & "@" & UserList(i).clase)
                   End If
                End If
            Next i
         
        Exit Sub
        
        '########## SISTEMA DE PARTY - FER ###########
    
    
        Case "/MEDITAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            If UserList(userindex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||4")
                Exit Sub
            End If
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
                Call SendData(SendTarget.toindex, userindex, 0, "||393")
                Call SendUserMP(val(userindex))
                Exit Sub
            End If
            
            
            If Not UserList(userindex).flags.Meditando Then
               Call SendData(SendTarget.toindex, userindex, 0, "||394")
               Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
            Else
               Call SendData(SendTarget.toindex, userindex, 0, "||205")
               Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
            End If
           UserList(userindex).flags.Meditando = Not UserList(userindex).flags.Meditando
           
           If UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN Then Exit Sub
           
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(userindex).flags.Meditando Then
                UserList(userindex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                
                UserList(userindex).Char.loops = LoopAdEternum
                If UserList(userindex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARCHICO
                ElseIf UserList(userindex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(userindex).Stats.ELV < 50 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARGRANDE
                ElseIf UserList(userindex).Stats.ELV <= 59 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARXGRANDE
                ElseIf UserList(userindex).Stats.ELV >= 60 Then
                  If UserList(userindex).StatusMith.EsStatus = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXNUEVATPNEUTRAL & "," & LoopAdEternum)
                  ElseIf UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXNUEVATPALIANZA & "," & LoopAdEternum)
                  ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXNUEVATPHORDA & "," & LoopAdEternum)
                  End If
                End If
                
                If UserList(userindex).flags.Transformado = 1 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARTRANSFO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARTRANSFO
                End If
            Else
                UserList(userindex).Counters.bPuedeMeditar = False
                
                UserList(userindex).Char.FX = 0
                UserList(userindex).Char.loops = 0
                Call SendData(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||9")
               Exit Sub
           End If
           
            If EstaEnRing(userindex) Then
                Call SendData(toindex, userindex, 0, "||395")
                Exit Sub
            End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||11")
               Exit Sub
           End If
           Call RevivirUsuario(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "||396")
           Exit Sub
          
           Case "/DEMONIO"
           
           If UserList(userindex).flags.Navegando = 1 Or UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||397")
           Exit Sub
           End If
           
           If UserList(userindex).flags.Transformado = 1 Then
               Call DarCuerpoDesnudo(userindex)
               Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_TRANSF)
               UserList(userindex).flags.Transformado = 0
          
           ElseIf UserList(userindex).flags.CJerarquia = 1 And Criminal(userindex) Then
                UserList(userindex).Char.Head = 0
                UserList(userindex).Char.Body = 289
               Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, 289, 0, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_TRANSF)
                UserList(userindex).flags.Transformado = 1
          
           Exit Sub
           End If
           
            Case "/ANGEL"
            
            If UserList(userindex).flags.Navegando = 1 Or UserList(userindex).flags.Muerto = 1 Then
             Call SendData(SendTarget.toindex, userindex, 0, "||397")
            Exit Sub
            End If
           
           
            If UserList(userindex).flags.Transformado = 1 Then
               Call DarCuerpoDesnudo(userindex)
               Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_TRANSF)
               UserList(userindex).flags.Transformado = 0
               
           ElseIf UserList(userindex).flags.CJerarquia = 1 And Not Criminal(userindex) Then
            UserList(userindex).Char.Head = 0
            UserList(userindex).Char.Body = 288
           Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, 288, 0, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
           Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
           Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_TRANSF)
           UserList(userindex).flags.Transformado = 1
           
           Exit Sub
           End If
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||9")
               Exit Sub
           End If
           
            If EstaEnRing(userindex) Then
                Call SendData(toindex, userindex, 0, "||395")
                Exit Sub
            End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||12")
               Exit Sub
           End If
            If UserList(userindex).flags.Envenenado = True Then
                UserList(userindex).flags.Envenenado = False
            End If
           UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
           Call SendUserHP(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "||398")
           Exit Sub

Case "/ABANDONAR"

If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
    UserList(userindex).Counters.TimeComandos = 5
    
    If UserList(userindex).flags.EnDuelo Then
        Call SalirDuelo(userindex)
    End If
    
    If UserList(userindex).flags.EnGuerra = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||399")
        UserList(userindex).flags.EnGuerra = 0
        Call WarpUserChar(userindex, 1, 54, 36, True)
    End If
 
    If UserList(userindex).Pos.Map = 72 Then
    If UserList(userindex).flags.EnDesafio = 1 And Desafio.Primero = userindex Then  'cambiar el 12 por numero de mapa
        Call WarpUserChar(Desafio.Primero, 1, 54, 36, True) 'mapa donde lleva al creador del desafio
        UserList(Desafio.Primero).flags.EnDesafio = 0
        UserList(Desafio.Primero).flags.rondas = 0
        Call SendData(SendTarget.ToAll, 0, 0, "||400")
        Call SendData(SendTarget.ToAll, 0, 0, "||351@" & UserList(Desafio.Primero).Name)
        Desafio.Primero = 0
        
        If Desafio.Segundo <> 0 Then
            Call WarpUserChar(Desafio.Segundo, 1, 54, 37, True) 'mapa donde llevar al retador
            UserList(Desafio.Segundo).flags.Desafio = 0
            Desafio.Segundo = 0
        End If
     Exit Sub
    End If
     
    If UserList(userindex).flags.Desafio = 1 And Desafio.Segundo = userindex Then
        Call WarpUserChar(Desafio.Segundo, 1, 54, 36, True)
        UserList(Desafio.Segundo).flags.Desafio = 0
        Call SendData(SendTarget.ToAll, 0, 0, "||351@" & UserList(Desafio.Segundo).Name)
        Desafio.Segundo = 0
     Exit Sub
    End If
End If

If userindex = Desafio.tPrimero Or userindex = Desafio.tSegundo Then 'mapa de pareja
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
        
        Call SendData(SendTarget.ToAll, 0, 0, "||401")
        Call SendData(SendTarget.ToAll, 0, 0, "||402@" & UserList(Desafio.tPrimero).Name & "@" & UserList(Desafio.tSegundo).Name)
ElseIf userindex = Desafio.tTercero Or userindex = Desafio.tCuarto Then

        Call WarpUserChar(Desafio.tTercero, UserList(Desafio.tTercero).flags.MapaAnterior, UserList(Desafio.tTercero).flags.XAnterior, UserList(Desafio.tTercero).flags.YAnterior)
        Call WarpUserChar(Desafio.tCuarto, UserList(Desafio.tCuarto).flags.MapaAnterior, UserList(Desafio.tCuarto).flags.XAnterior, UserList(Desafio.tCuarto).flags.YAnterior)
        UserList(Desafio.tTercero).flags.tEsperaPareja = False
        UserList(Desafio.tTercero).flags.tSuPareja = 0
        UserList(Desafio.tCuarto).flags.tEsperaPareja = False
        UserList(Desafio.tCuarto).flags.tSuPareja = 0
        Desafio.tTercero = 0
        Desafio.tCuarto = 0
        
        Call SendData(SendTarget.ToAll, 0, 0, "||402@" & UserList(Desafio.tTercero).Name & "@" & UserList(Desafio.tCuarto).Name)
    Exit Sub
End If

If UserList(userindex).Pos.Map = 54 Then
    If Pareja.Jugador(1) > 0 And Pareja.Jugador(2) > 0 And UserList(userindex).flags.EnPareja = True And Pareja.Jugador(3) = 0 And Pareja.Jugador(4) = 0 Then
            Call WarpUserChar(Pareja.Jugador(1), UserList(Pareja.Jugador(1)).flags.MapaAnterior, UserList(Pareja.Jugador(1)).flags.XAnterior, UserList(Pareja.Jugador(1)).flags.YAnterior, True)
            Call WarpUserChar(Pareja.Jugador(2), UserList(Pareja.Jugador(2)).flags.MapaAnterior, UserList(Pareja.Jugador(2)).flags.XAnterior, UserList(Pareja.Jugador(2)).flags.YAnterior, True)
            Call SendData(SendTarget.ToAll, 0, 0, "||403@" & UserList(Pareja.Jugador(1)).Name & "@" & UserList(Pareja.Jugador(2)).Name)
            UserList(Pareja.Jugador(1)).flags.EnPareja = False
            UserList(Pareja.Jugador(1)).flags.EsperaPareja = False
            UserList(Pareja.Jugador(1)).flags.SuPareja = 0
            UserList(Pareja.Jugador(2)).flags.EnPareja = False
            UserList(Pareja.Jugador(2)).flags.EsperaPareja = False
            UserList(Pareja.Jugador(2)).flags.SuPareja = 0
            Pareja.Jugador(1) = 0
            Pareja.Jugador(2) = 0
            HayPareja = False
            Exit Sub
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||239")
            Exit Sub
        End If
End If

If UserList(userindex).Pos.Map = 101 And UserList(userindex).flags.EspectadorArena1 = 1 Then

    WarpUserChar userindex, UserList(userindex).flags.MapaAnterior, UserList(userindex).flags.XAnterior, UserList(userindex).flags.YAnterior, True
    UserList(userindex).flags.EspectadorArena1 = 0
    EspectadoresEnArena1 = EspectadoresEnArena1 - 1
    
ElseIf UserList(userindex).Pos.Map = 101 And UserList(userindex).flags.EspectadorArena2 = 1 Then

    WarpUserChar userindex, UserList(userindex).flags.MapaAnterior, UserList(userindex).flags.XAnterior, UserList(userindex).flags.YAnterior, True
    UserList(userindex).flags.EspectadorArena2 = 0
    EspectadoresEnArena2 = EspectadoresEnArena2 - 1
    
ElseIf UserList(userindex).Pos.Map = 101 And UserList(userindex).flags.EspectadorArena3 = 1 Then

    WarpUserChar userindex, UserList(userindex).flags.MapaAnterior, UserList(userindex).flags.XAnterior, UserList(userindex).flags.YAnterior, True
    UserList(userindex).flags.EspectadorArena3 = 0
    EspectadoresEnArena3 = EspectadoresEnArena3 - 1
    
ElseIf UserList(userindex).Pos.Map = 101 And UserList(userindex).flags.EspectadorArena4 = 1 Then

    WarpUserChar userindex, UserList(userindex).flags.MapaAnterior, UserList(userindex).flags.XAnterior, UserList(userindex).flags.YAnterior, True
    UserList(userindex).flags.EspectadorArena4 = 0
    EspectadoresEnArena4 = EspectadoresEnArena4 - 1
    
End If

Exit Sub

    Case "/DESAFIO"
    
    If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
    UserList(userindex).Counters.TimeComandos = 5
                
        If UserList(userindex).EnCvc = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||239")
        Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 141 Then Exit Sub
                
        If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||323")
            Exit Sub
        End If
                
                If UserList(userindex).Pos.Map = 101 Then Exit Sub
                If UserList(userindex).Pos.Map = 100 Then Exit Sub
                If UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 118 Then Exit Sub
           
            If UserList(userindex).flags.Muerto = 1 Then
              Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            
            If UserList(userindex).Stats.GLD < 120000 Then
              Call SendData(SendTarget.toindex, userindex, 0, "||215@120.000")
                Exit Sub
            End If
                    
            If UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 120 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||291")
                Exit Sub
            End If

            
        If TieneItemDiosEquipado(userindex) = True Then
            Call SendData(toindex, userindex, 0, "||404")
            Exit Sub
        End If
                   
            If Desafio.Primero <> 0 Then 'mapa de desafio
            Call SendData(SendTarget.toindex, userindex, 0, "||405")
                Exit Sub
            End If
                    
            If Desafio.Primero <> 0 And Desafio.Segundo <> 0 Then 'mapa de desafio
                Call SendData(SendTarget.toindex, userindex, 0, "||406")
            Exit Sub
            End If
            
            If UserList(userindex).Stats.ELV < 50 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||538")
                Exit Sub
            End If
          
            Call SendData(SendTarget.ToAll, 0, 0, "||407@" & UserList(userindex).Name & "@" & UserList(userindex).clase & "@" & UserList(userindex).Stats.ELV)
            
            Call WarpUserChar(userindex, 72, 52, 32, True) 'Mapa y posicion del mapa de desafio
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 120000
            UserList(userindex).flags.EnDesafio = 1
            Call SendUserGLD(userindex) 'enviamos todo
            Desafio.Primero = userindex
           
        Exit Sub
       
        Case "/DESAFIAR"
        
        If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
        UserList(userindex).Counters.TimeComandos = 5
        
        If UserList(userindex).EnCvc = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||239")
        Exit Sub
        End If
        
        If TieneItemDiosEquipado(userindex) = True Then
            Call SendData(toindex, userindex, 0, "||404")
            Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 141 Then Exit Sub
        
        If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||323")
                Exit Sub
            End If
        
            If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||3")
            Exit Sub
            End If
            
            If UserList(userindex).Stats.GLD < 30000 Then
              Call SendData(SendTarget.toindex, userindex, 0, "||215@30.000")
                Exit Sub
            End If
            
            If UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||291")
                Exit Sub
            End If
           
            If Desafio.Primero = 0 Then 'mapa de desafio
                Call SendData(SendTarget.toindex, userindex, 0, "||408")
                Exit Sub
            End If
                
            If Desafio.Primero <> 0 And Desafio.Segundo <> 0 Then 'mapa de desafio
            Call SendData(SendTarget.toindex, userindex, 0, "||409")
            Exit Sub
            End If
           
            Call SendData(SendTarget.ToAll, 0, 0, "||410@" & UserList(userindex).Name)
            Call SendData(SendTarget.toindex, Desafio.Primero, 0, "||411@" & UserList(userindex).Name & "@" & UserList(userindex).clase & "@" & UserList(userindex).Stats.ELV)
            Call WarpUserChar(userindex, 72, 52, 48, True) 'mapa y pos del desafio
            Call WarpUserChar(Desafio.Primero, 72, 52, 32, True)
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 30000
            UserList(userindex).flags.Desafio = 1
            'ATENCION ACA si usas mod twister ponele este signo ' al call senduserstatsbox de aca abajo si usas 0.11.5 poenele 'call enviaroro(userindex) y sacale al senduserstatbox el ''
            'Call EnviarOro(UserIndex) 'Esto para mod twist o cualquier mod que haya reducido los paquetes
            Call SendUserGLD(userindex) 'enviamos todo
            Desafio.Segundo = userindex
           
            Exit Sub
             
        Case "/EST"
            Call SendUserStatsTxt(userindex, userindex)
        Exit Sub
            
        Case "/SEG"
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEGON")
            End If
            UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            UserList(userindex).flags.SeguroClan = Not UserList(userindex).flags.SeguroClan
        Exit Sub
            
            Case "/SEGR"
            If UserList(userindex).flags.SeguroResu = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "SEGOFR")
                UserList(userindex).flags.SeguroResu = False
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEGONR")
                UserList(userindex).flags.SeguroResu = True
            End If
           ' UserList(UserIndex).flags.SeguroResu = Not UserList(UserIndex).flags.SeguroResu
            Exit Sub
        '[/Alejo]
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            
            If UserList(userindex).flags.Comerciando Then Exit Sub
            
            
            If UserList(userindex).flags.Privilegios = PlayerType.Coordination Then Exit Sub
            
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(userindex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||13")
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(userindex)
            '[Alejo]
            ElseIf UserList(userindex).flags.TargetUser > 0 Then
                comManda userindex
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||9")
            End If
        Exit Sub
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 5 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||13")
                    Exit Sub
                End If
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(userindex)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||9")
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||9")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||158")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(userindex)
           Else
                  Call EnlistarCaos(userindex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||9")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||10")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||9")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||12")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(userindex)
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(userindex)
           End If
           Exit Sub
    End Select
 
   If UCase$(Left$(rData, 5)) = "/CVC " Then
 
        Dim Ret         As String
        Dim Retsub         As String
        Dim Que         As String
        Dim Usuarios    As Integer
        Dim ja          As Integer
        Dim pre         As Long
        Dim h           As Integer
        Dim pret        As String
        Dim pretSub        As String
        Dim ClanName    As String
        
            ClanName = Right$(rData, Len(rData) - 5)
            
            
            
        If Not UserList(userindex).GuildIndex >= 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||120")
        Exit Sub
        End If
        
        If ClanName = Guilds(UserList(userindex).GuildIndex).GuildName Then Exit Sub
    
            If CvcFunciona = True Then
                SendData SendTarget.toindex, userindex, 0, "||364"
            Exit Sub
            End If
            
            Usuarios = 0
          
            For ja = 1 To LastUser
                If UserList(ja).GuildIndex > 0 Then
                    If Guilds(UserList(userindex).GuildIndex).GuildName = Guilds(UserList(ja).GuildIndex).GuildName Then
                    If UserList(ja).flags.SeguroCVC = True And UserList(ja).flags.Muerto = 0 And UserList(ja).Counters.Pena = 0 And UserList(ja).Pos.Map <> 71 And UserList(ja).Pos.Map <> 100 And UserList(ja).Pos.Map <> 108 And UserList(ja).Pos.Map <> 109 And UserList(ja).Pos.Map <> 106 Then
                        Usuarios = Usuarios + 1
                    End If
                    End If
                End If
            Next ja
            
            If Usuarios < 1 Then
                SendData SendTarget.toindex, userindex, 0, "||365"
            Exit Sub
            End If
            
            rData = Right$(rData, Len(rData) - 5)
            If UserList(userindex).GuildIndex <> 0 Then
            
            Ret = SendGuildLeaderInfo(userindex)
            Retsub = SendGuildSubLeaderInfo(userindex)
        
           
            If Ret = vbNullString And Retsub = vbNullString Then
            SendData SendTarget.toindex, userindex, 0, "||412"
                Exit Sub
            Else
           
           
           
            For h = 1 To LastUser
             If UserList(h).GuildIndex <> 0 Then
           
                If LCase(Guilds(UserList(h).GuildIndex).GuildName) = LCase(ClanName) Then
                    pret = SendGuildLeaderInfo(h)
                    
                If pret = vbNullString Then
                Else
                    SendData SendTarget.toindex, h, 0, "||413@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@" & Usuarios
                    SendData SendTarget.toindex, userindex, 0, "||414@" & Guilds(UserList(h).GuildIndex).GuildName
                End If
                
                Guilds(UserList(h).GuildIndex).TieneParaDesafiar = True
                Guilds(UserList(h).GuildIndex).ClanPideDesafio = Guilds(UserList(userindex).GuildIndex).GuildName
                Else
                
                End If
        End If
            Next h
            'CVC
            Exit Sub
        End If
    End If
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
      rData = Right$(rData, Len(rData) - 6)
          If rData <> " " And rData <> "" Then
            Call SendData(SendTarget.ToPartyArea, userindex, 0, "||415@" & UserList(userindex).Name & "@" & rData)
          End If
        Exit Sub
    End If
   
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(userindex, tInt)
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 4)) = "/IR " Then
        rData = Right$(rData, Len(rData) - 4)
    
        With UserList(userindex)
        
            If .EnCvc = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||239")
            Exit Sub
            End If
        
                If UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 120 Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 8 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||416")
                    Exit Sub
                End If
    
                
                If UserList(userindex).flags.Paralizado = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||417")
                    Exit Sub
                End If
        
                If UCase$(rData) = "22" Or UCase$(rData) = "90" Or UCase$(rData) = "17" Or UCase$(rData) = "INVOCACIONES" Then
                    If UserList(userindex).flags.EsPremium = 0 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||418")
                        Exit Sub
                    End If
                    
                        If UCase$(rData) = "INVOCACIONES" Then
                            UserList(userindex).UserPremiumMap = 93
                        Else
                            UserList(userindex).UserPremiumMap = val(rData)
                        End If
                    
                    Call SendData(SendTarget.toindex, userindex, 0, "||419")
                    UserList(userindex).Counters.TransportePremium = 2
                    
                    Exit Sub
                End If
        
        
        If Not UserList(userindex).GuildIndex >= 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||120")
            Exit Sub
        End If
        
        If UCase$(rData) = "FORTALEZA" Or UCase$(rData) = "35" Then
                If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(Fortaleza) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||420")
                    Exit Sub
                End If
                
                If UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||421")
                    Exit Sub
                End If
                
                If userindex = GranPoder Then
                    Call OtorgarGranPoder(0)
                    UserList(userindex).flags.GranPoder = 0
                    SendUserVariant (userindex)
                End If
                
                Call SendData(SendTarget.toindex, userindex, 0, "||419")
                UserList(userindex).Counters.TransporteCastillos(35) = 2
            Exit Sub
        End If
        If UCase$(rData) = "33" Then
        
                If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloNorte) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||420")
                    Exit Sub
                End If
                
            Call SendData(SendTarget.toindex, userindex, 0, "||419")
            UserList(userindex).Counters.TransporteCastillos(33) = 2
            mapa = MapCastilloN
        End If
        If UCase$(rData) = "31" Then
                If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloSur) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||420")
                    Exit Sub
                End If
        
            Call SendData(SendTarget.toindex, userindex, 0, "||419")
            UserList(userindex).Counters.TransporteCastillos(31) = 2
            mapa = MapCastilloS
        End If
        If UCase$(rData) = "34" Then
                If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloEste) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||420")
                    Exit Sub
                End If
        
            Call SendData(SendTarget.toindex, userindex, 0, "||419")
            UserList(userindex).Counters.TransporteCastillos(34) = 2
            mapa = MapCastilloE
        End If
        
        If UCase$(rData) = "32" Then
        
                If UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) <> UCase$(CastilloOeste) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||420")
                    Exit Sub
                End If
                
            Call SendData(SendTarget.toindex, userindex, 0, "||419")
            UserList(userindex).Counters.TransporteCastillos(32) = 2
            mapa = MapCastilloO
        End If
       
        If mapa = 0 Then Exit Sub
                If userindex = GranPoder Then
                    Call OtorgarGranPoder(0)
                    UserList(userindex).flags.GranPoder = 0
                    SendUserVariant (userindex)
                End If
        Exit Sub
        End With

      Exit Sub
    End If
    
'HORA
If UCase$(Left$(rData, 5)) = "/HORA" Then
    rData = Right$(rData, Len(rData) - 5)
    
  If UserList(userindex).flags.Privilegios > PlayerType.UserSupport Then
    Call SendData(SendTarget.ToAll, 0, 0, "||853@" & Time & "@" & Date)
  Else
    Call SendData(SendTarget.toindex, userindex, 0, "||853@" & Time & "@" & Date)
  End If
  
    Exit Sub
End If

    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(userindex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(userindex).Name, "Mensaje a Gms:" & rData, False)
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||429@" & UserList(userindex).Name & "@" & rData)
        End If
        Exit Sub
    End If
    
    
    Select Case UCase(Left(rData, 5))
        Case "/_BUG "
            n = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As n
            Print #n,
            Print #n,
            Print #n, "########################################################################"
            Print #n, "########################################################################"
            Print #n, "Usuario:" & UserList(userindex).Name & "  Fecha:" & Date & "    Hora:" & Time
            Print #n, "########################################################################"
            Print #n, "BUG:"
            Print #n, Right$(rData, Len(rData) - 5)
            Print #n, "########################################################################"
            Print #n, "########################################################################"
            Print #n,
            Print #n,
            Close #n
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 10))
 
        Case "/CASTILLOS"
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Norte@" & CastilloNorte)
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Sur@" & CastilloSur)
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Este@" & CastilloEste)
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Oeste@" & CastilloOeste)
        Call SendData(SendTarget.toindex, userindex, 0, "||431@" & Fortaleza)
        
      If UserList(userindex).GuildIndex > 0 Then
        If Guilds(UserList(userindex).GuildIndex).GuildName = CastilloNorte And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloSur And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloEste And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloOeste And Guilds(UserList(userindex).GuildIndex).GuildName = Fortaleza Then
            Call SendData(SendTarget.toindex, userindex, 0, "||432@" & PremiosCastis)
        End If
      End If
        
        Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/NICK "
            rData = Right$(rData, Len(rData) - 6)
            tIndex = NameIndex(rData)
            
            If (IsAdministrator(rData)) Or (IsDevelopment(rData)) Or (IsCoordination(rData)) Or (IsTournamentManager(rData)) Or (IsEventManager(rData)) Or (IsUserSupport(rData)) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||433")
            Exit Sub
            End If
            
            If tIndex <= 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||434@" & UCase$(rData))
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||435@" & UCase$(rData))
            End If
        Exit Sub
        Case "/DESC "
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||436")
                Exit Sub
            End If
            UserList(userindex).Desc = Trim$(rData)
            Call SendData(SendTarget.toindex, userindex, 0, "||437")
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(userindex, rData, tStr) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||438@" & tStr)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||439")
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/NOADV " Then
        Name = Right$(rData, Len(rData) - 7)
            If Name = "" Then Exit Sub
       
            Name = Replace(Name, "\", "")
            Name = Replace(Name, "/", "")
       
            tIndex = NameIndex(rData)
     
        If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
     
        If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
        SendData SendTarget.toindex, userindex, 0, "||440"
        Exit Sub
        End If
     
        Call WriteVar(CharPath & Name & ".chr", "Penas", "CANT", 0)
     
        Dim Penas As Integer
        For Penas = 1 To 5
        Call WriteVar(CharPath & Name & ".chr", "Penas", "P" & Penas, "0")
        Next Penas
     
        Call SendData(SendTarget.toindex, userindex, 0, "||441")
     
     
        Exit Sub
        End If
 
If UCase$(Left$(rData, 9)) = "/LIBERAR " Then
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||442")
    Else
        Call WarpUserChar(tIndex, 1, 54, 36, True)
        UserList(tIndex).Counters.Pena = 0
        Call SendData(SendTarget.ToAdmins, 0, 0, "||443@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        Call SendData(SendTarget.toindex, tIndex, 0, "||444")
    Exit Sub
  End If
End If
    
If UCase$(Left$(rData, 7)) = "/PENAS " Then
        Name = Right$(rData, Len(rData) - 7)
        If Name = "" Then Exit Sub
        
        If UserList(userindex).flags.Privilegios < PlayerType.UserSupport Then Exit Sub
   
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
   
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||445")
            Else
                While tInt > 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||327@" & tInt & "@" & GetVar(CharPath & Name & ".chr", "PENAS", "P" & tInt))
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||189@" & Name)
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PAREJA "
        rData = Right$(rData, Len(rData) - 8)
        
        tIndex = NameIndex(rData)
        
        If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
        UserList(userindex).Counters.TimeComandos = 5
               
        If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
        End If
        
        If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 141 Then Exit Sub
        
        If UserList(userindex).Stats.GLD < 120000 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||215@120.000")
        Exit Sub
        End If
        
        If UserList(tIndex).Stats.GLD < 120000 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||446")
        Exit Sub
        End If
        
        If UserList(userindex).EnCvc = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||239")
        Exit Sub
        End If
        
        If tIndex = userindex Then Exit Sub
        
        'Si la pareja no puede participar
        If UserList(tIndex).EnCvc = True Or UserList(tIndex).flags.Muerto = 1 Or TieneItemDiosEquipado(tIndex) = True Or UserList(tIndex).cComercio.cComercia = True Or UserList(tIndex).Pos.Map = 101 Or UserList(tIndex).Pos.Map = 54 Or UserList(tIndex).Pos.Map = 8 Or UserList(tIndex).Pos.Map = 72 Or UserList(tIndex).Pos.Map = 78 Or UserList(tIndex).Pos.Map = 100 Or UserList(tIndex).Pos.Map = 110 Or UserList(tIndex).Pos.Map = 120 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||447")
        Exit Sub
        End If
        
        If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||323")
            Exit Sub
        End If
       
        If UserList(userindex).flags.Muerto = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||3")
        Exit Sub
        End If
        
        If TieneItemDiosEquipado(userindex) = True Then
            Call SendData(toindex, userindex, 0, "||404")
            Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 120 Or UserList(userindex).Pos.Map = 110 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||291")
        Exit Sub
        End If
        
        If Pareja.Jugador(3) > 0 And Pareja.Jugador(4) > 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||406")
        Exit Sub
        End If
       
        If UserList(userindex).clase = UserList(tIndex).clase Then
            Call SendData(SendTarget.toindex, userindex, 0, "||448")
        Exit Sub
        End If
       
        If Pareja.Jugador(1) = 0 And Pareja.Jugador(2) = 0 Then 'mapa de duelos 2vs2
        UserList(tIndex).flags.EsperaPareja = True
        UserList(userindex).flags.SuPareja = tIndex
       
            If UserList(userindex).flags.EsperaPareja = False Then
                Call SendData(SendTarget.toindex, tIndex, 0, "||449@" & UserList(userindex).Name)
            End If
       
            If UserList(tIndex).flags.SuPareja = userindex Then
            Pareja.Jugador(1) = userindex
            Pareja.Jugador(2) = tIndex
            UserList(Pareja.Jugador(1)).flags.EnPareja = True
            UserList(Pareja.Jugador(2)).flags.EnPareja = True
            
            'Guardamos posiciones
            UserList(Pareja.Jugador(1)).flags.MapaAnterior = UserList(Pareja.Jugador(1)).Pos.Map
            UserList(Pareja.Jugador(1)).flags.XAnterior = UserList(Pareja.Jugador(1)).Pos.X
            UserList(Pareja.Jugador(1)).flags.YAnterior = UserList(Pareja.Jugador(1)).Pos.Y
            
            UserList(Pareja.Jugador(2)).flags.MapaAnterior = UserList(Pareja.Jugador(2)).Pos.Map
            UserList(Pareja.Jugador(2)).flags.XAnterior = UserList(Pareja.Jugador(2)).Pos.X
            UserList(Pareja.Jugador(2)).flags.YAnterior = UserList(Pareja.Jugador(2)).Pos.Y
            
            
            Call WarpUserChar(Pareja.Jugador(1), 54, 41, 55) 'mapa 2vs2, posicion jugador numero 1
            Call WarpUserChar(Pareja.Jugador(2), 54, 43, 57) 'mapa 2vs2, posicion jugador numero 2
            UserList(Pareja.Jugador(1)).Stats.GLD = UserList(Pareja.Jugador(1)).Stats.GLD - 120000
            UserList(Pareja.Jugador(2)).Stats.GLD = UserList(Pareja.Jugador(2)).Stats.GLD - 120000
            SendUserGLD (Pareja.Jugador(1))
            SendUserGLD (Pareja.Jugador(2))
            Call SendData(SendTarget.ToAll, 0, 0, "||450@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
            End If
       
        Exit Sub
        End If
       
        If Pareja.Jugador(1) > 0 And Pareja.Jugador(2) > 0 Then 'mapa de duelos 2vs2
        UserList(tIndex).flags.EsperaPareja = True
        UserList(userindex).flags.SuPareja = tIndex
 
            If UserList(userindex).flags.EsperaPareja = False Then
                Call SendData(SendTarget.toindex, tIndex, 0, "||449@" & UserList(userindex).Name)
            End If
 
            If UserList(tIndex).flags.SuPareja = userindex Then
                Pareja.Jugador(3) = userindex
                Pareja.Jugador(4) = tIndex
                UserList(Pareja.Jugador(3)).flags.EnPareja = True
                UserList(Pareja.Jugador(4)).flags.EnPareja = True
                
                'Guardamos posiciones
                UserList(Pareja.Jugador(3)).flags.MapaAnterior = UserList(Pareja.Jugador(3)).Pos.Map
                UserList(Pareja.Jugador(3)).flags.XAnterior = UserList(Pareja.Jugador(3)).Pos.X
                UserList(Pareja.Jugador(3)).flags.YAnterior = UserList(Pareja.Jugador(3)).Pos.Y
                
                UserList(Pareja.Jugador(4)).flags.MapaAnterior = UserList(Pareja.Jugador(4)).Pos.Map
                UserList(Pareja.Jugador(4)).flags.XAnterior = UserList(Pareja.Jugador(4)).Pos.X
                UserList(Pareja.Jugador(4)).flags.YAnterior = UserList(Pareja.Jugador(4)).Pos.Y
                
                Call WarpUserChar(Pareja.Jugador(1), 54, 41, 55) 'mapa 2vs2, posicion jugador numero 1
                Call WarpUserChar(Pareja.Jugador(2), 54, 43, 57) 'mapa 2vs2, posicion jugador numero 2
                Call WarpUserChar(Pareja.Jugador(3), 54, 60, 40) 'mapa 2vs2, posicion jugador numero 3
                Call WarpUserChar(Pareja.Jugador(4), 54, 62, 42) 'mapa 2vs2, posicion jugador numero 4
                UserList(Pareja.Jugador(3)).Stats.GLD = UserList(Pareja.Jugador(3)).Stats.GLD - 120000
                UserList(Pareja.Jugador(4)).Stats.GLD = UserList(Pareja.Jugador(4)).Stats.GLD - 120000
                SendUserGLD (Pareja.Jugador(3))
                SendUserGLD (Pareja.Jugador(4))
                Call SendData(SendTarget.ToAll, 0, 0, "||451@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
                HayPareja = True
            End If
       
        Exit Sub
        End If
    End Select
        
    Procesado = False
End Sub

