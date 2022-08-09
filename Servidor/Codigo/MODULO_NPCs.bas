Attribute VB_Name = "NPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit
Public Const ReyNpcN As Integer = 910
Public Const ReyNpcF As Integer = 615

Sub QuitarMascota(ByVal userindex As Integer, ByVal NpcIndex As Integer)

Dim i As Integer
UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(userindex).MascotasIndex(i) = NpcIndex Then
     UserList(userindex).MascotasIndex(i) = 0
     UserList(userindex).MascotasType(i) = 0
     Exit For
  End If
Next i

If Npclist(NpcIndex).Numero = ELEMENTALAGUA Then UserList(userindex).flags.EleDeAgua = 0
If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Then UserList(userindex).flags.EleDeFuego = 0

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal userindex As Integer)
  On Error Resume Next

   Dim MiNPC As npc
   MiNPC = Npclist(NpcIndex)
      
   If userindex <> 0 And Npclist(NpcIndex).NPCtype = eNPCType.ReyCastillo Then
        Call MuereRey(userindex, NpcIndex)
        Exit Sub
    End If
    
    If NpcIndex = 963 Or NpcIndex = 964 Then
        Call Aram_KillTower(userindex)
    End If
    
    If NpcIndex = ArieteUno Then
        Call SendData(SendTarget.ToMap, 0, 81, "||46")
        Call QuitarNPC(ArieteUno)
        ArieteUno = 0
        RejaSurAtacada = False
    Exit Sub
    End If
    
    If NpcIndex = ArieteDos Then
        Call SendData(SendTarget.ToMap, 0, 81, "||47")
        Call QuitarNPC(ArieteDos)
        ArieteDos = 0
        RejaCentralAtacada = False
    Exit Sub
    End If
    
    If NpcIndex = ArieteTres Then
        Call SendData(SendTarget.ToMap, 0, 81, "||48")
        Call QuitarNPC(ArieteTres)
        ArieteTres = 0
        RejaNorteAtacada = False
    Exit Sub
    End If
    
If MiNPC.Numero = 951 And UserList(userindex).Pos.Map = 141 Then
 If TieneObjetos(1073, 1, userindex) = True And TieneObjetos(1074, 1, userindex) = True And TieneObjetos(1075, 1, userindex) = True And TieneObjetos(1076, 1, userindex) = True Then
    Dim Fer As Obj
    Fer.ObjIndex = 1077
    Fer.Amount = 1
    Call MeterItemEnInventario(userindex, Fer)
    Call SendData(SendTarget.toindex, userindex, 0, "||49")
 End If
 
    Dim i As Long
   
    For i = 1 To LastUser
        If UserList(i).Pos.Map = 141 Then
            Call WarpUserChar(i, 26, RandomNumber(56, 66), RandomNumber(36, 39))
        End If
    Next i
    
     Call QuitarNPC(NpcIndex)
    Exit Sub
End If
    
    'rey orco - apocalipsis
    If MiNPC.MVP = 1 Then
        SendData SendTarget.ToAll, 0, 0, "||697@" & UserList(userindex).Name & "@" & MiNPC.Name
        UserList(userindex).flags.MVPMatados = UserList(userindex).flags.MVPMatados + 1
    End If

    'Mascota
    Dim asd As Integer
    If MiNPC.Numero = 156 Or MiNPC.Numero = 157 Or MiNPC.Numero = 158 Or MiNPC.Numero = 181 Or MiNPC.Numero = 182 Then
        asd = Npclist(NpcIndex).DueñoMascota
        UserList(asd).flags.InvocoMascota = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
   
    
   If userindex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & MiNPC.flags.Snd3)
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        
        'El user que lo mato tiene mascotas?
        If UserList(userindex).NroMacotas > 0 Then
            Dim T As Integer
            For T = 1 To MAXMASCOTAS
                  If UserList(userindex).MascotasIndex(T) > 0 Then
                      If Npclist(UserList(userindex).MascotasIndex(T)).TargetNPC = NpcIndex Then
                              Call FollowAmo(UserList(userindex).MascotasIndex(T))
                      End If
                  End If
            Next T
        End If
        
        '[/KEVIN]
        Call SendData(SendTarget.toindex, userindex, 0, "||50")
        
        If UserList(userindex).Stats.NPCsMuertos < 32000 Then _
            UserList(userindex).Stats.NPCsMuertos = UserList(userindex).Stats.NPCsMuertos + 1
        
        Call CheckUserLevel(userindex)
    End If
    
'Gana horda
If NpcIndex = ReyGuerraIndex And HayGuerraAnvil = True Then
    Call SendData(SendTarget.ToAll, 0, 0, "||51")

Dim loopC As Integer
For loopC = 1 To LastUser
    If Criminal(loopC) And UserList(loopC).Pos.Map = 29 Then
        UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + 300000
        UserList(loopC).flags.GuerrasGanadas = UserList(loopC).flags.GuerrasGanadas + 1
        Call SendData(SendTarget.toindex, loopC, 0, "||63@300.000")
    End If
    
    If EsAlianza(loopC) And UserList(loopC).Pos.Map = 29 Then
        UserList(loopC).flags.GuerrasPerdidas = UserList(loopC).flags.GuerrasPerdidas + 1
    End If
Next loopC

If Criminal(userindex) Then
    Call AgregarPuntos(userindex, 30)
    Call SendData(SendTarget.toindex, userindex, 0, "||53")
End If

HayGuerra = False
HayGuerraAnvil = False
Minus = 0
MapInfo(29).Pk = False

'Gana alianza
ElseIf NpcIndex = ReyGuerraIndex And HayGuerraKhalim = True Then
    Call SendData(SendTarget.ToAll, 0, 0, "||52")
    
For loopC = 1 To LastUser
    If EsAlianza(loopC) And UserList(loopC).Pos.Map = 27 Then
        UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + 300000
        UserList(loopC).flags.GuerrasGanadas = UserList(loopC).flags.GuerrasGanadas + 1
        Call SendData(SendTarget.toindex, loopC, 0, "||63@300.000")
    End If
    
    If Criminal(loopC) And UserList(loopC).Pos.Map = 27 Then
        UserList(loopC).flags.GuerrasPerdidas = UserList(loopC).flags.GuerrasPerdidas + 1
    End If
Next loopC

    If EsAlianza(userindex) Then
        Call AgregarPuntos(userindex, 30)
        Call SendData(SendTarget.toindex, userindex, 0, "||53")
    End If

HayGuerra = False
HayGuerraKhalim = False
Minus = 0
MapInfo(27).Pk = False

End If
    
   'Quitamos el npc
   Call QuitarNPC(NpcIndex)
   
   If MiNPC.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MiNPC, userindex)
        Call NPCDarPuntos(MiNPC, userindex)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC, userindex)
   End If
   
If UserList(userindex).Pos.Map = 95 Then

    If MiNPC.Numero = 938 Then
            If GuardiasRey < 4 Then
                GuardiasRey = GuardiasRey + 1
            End If
           
            If GuardiasRey = 4 Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "||699")
                Npclist(IndexReyAncalagon).Char.AuraA = 0
                Call MakeNPCChar(SendTarget.ToMap, 0, 0, IndexReyAncalagon, Npclist(IndexReyAncalagon).Pos.Map, Npclist(IndexReyAncalagon).Pos.X, Npclist(IndexReyAncalagon).Pos.Y)
            End If
        Exit Sub
    End If

    If MiNPC.Numero = 937 Then
       Dim PosicionD As WorldPos
       Dim DracoPowa As Integer
       DracoPowa = 936
    
            PosicionD.Map = MiNPC.Pos.Map
            PosicionD.X = MiNPC.Pos.X
            PosicionD.Y = MiNPC.Pos.Y
            Call SpawnNpc(DracoPowa, PosicionD, True, False)
        Exit Sub
      End If
     
   If MiNPC.Numero = 936 Then
        ReyON = 0
        IndexReyAncalagon = 0
        GuardiasRey = 0
        Call SendData(ToAll, 0, 0, "||54")
        Call AgregarPuntos(userindex, 25)
        Call SendData(SendTarget.toindex, userindex, 0, "||57@25")
     Exit Sub
   End If
End If
   
   
    If MiNPC.Numero = 958 Then
        If GuardianesAlianza < 4 Then
            GuardianesAlianza = GuardianesAlianza + 1
        End If
     Exit Sub
    ElseIf MiNPC.Numero = 959 Then
        If GuardianesHordas < 4 Then
            GuardianesHordas = GuardianesHordas + 1
        End If
     Exit Sub
    End If
   
   If MiNPC.Cristales = 1 Then
    Call NPC_TIRAR_CRISTALES(MiNPC, userindex)
   End If
   
    If UserList(userindex).Pos.Map = mapainvo Then
     InvocoBicho = False
    End If

'------------------------------
'5TA JERARQUIA
'------------------------------
If UserList(userindex).Faccion.RecompensasReal = 4 Or UserList(userindex).Faccion.RecompensasCaos = 4 Then
  Dim Objetito As Obj
  Objetito.Amount = 1
  
  If MiNPC.Numero = 566 Then
  
    Objetito.ObjIndex = 1221
    If Not MeterItemEnInventario(userindex, Objetito) Then
        Call TirarItemAlPiso(UserList(userindex).Pos, Objetito)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||700@" & ObjData(Objetito.ObjIndex).Name)
    
  ElseIf MiNPC.Numero = 542 Then
  
    Objetito.ObjIndex = 1220
    If Not MeterItemEnInventario(userindex, Objetito) Then
        Call TirarItemAlPiso(UserList(userindex).Pos, Objetito)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||700@" & ObjData(Objetito.ObjIndex).Name)
    
  ElseIf MiNPC.Numero = 564 Then
  
    Objetito.ObjIndex = 1222
    If Not MeterItemEnInventario(userindex, Objetito) Then
        Call TirarItemAlPiso(UserList(userindex).Pos, Objetito)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||700@" & ObjData(Objetito.ObjIndex).Name)
    
  ElseIf MiNPC.Numero = 949 Then
  
    Objetito.ObjIndex = 1224
    If Not MeterItemEnInventario(userindex, Objetito) Then
        Call TirarItemAlPiso(UserList(userindex).Pos, Objetito)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||700@" & ObjData(Objetito.ObjIndex).Name)
    
  ElseIf MiNPC.Numero = 911 Then
  
    Objetito.ObjIndex = 1223
    If Not MeterItemEnInventario(userindex, Objetito) Then
        Call TirarItemAlPiso(UserList(userindex).Pos, Objetito)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||700@" & ObjData(Objetito.ObjIndex).Name)
    
  End If
End If
'------------------------------
'5TA JERARQUIA
'------------------------------

  If MiNPC.Numero = 92 Then
        asd = Npclist(NpcIndex).MaestroUser
        UserList(asd).flags.EleDeAgua = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
    
    If MiNPC.Numero = 93 Then
        asd = Npclist(NpcIndex).MaestroUser
        UserList(asd).flags.EleDeFuego = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
    
    If MiNPC.Numero = 94 Then
        asd = Npclist(NpcIndex).MaestroUser
        UserList(asd).flags.EleDeTierra = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
    
    If MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Tipo = 1 And MiNPC.Numero = MisionesDiarias(UserList(userindex).Misiones.NumeroMision).NumNPC Then
        If UserList(userindex).Misiones.ConteoUser < MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad Then UserList(userindex).Misiones.ConteoUser = UserList(userindex).Misiones.ConteoUser + 1
    End If
    
    If UserList(userindex).flags.UserNumQuest > 0 Then Call modQuests.RestarNPC(userindex, MiNPC.Numero)

   'ReSpawn o no
   Call ReSpawnNpc(MiNPC)
   
   
Exit Sub

'Errhandler:
    'Call LogError("Error en MuereNpc")
    
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AfectaRelampago = 0
        .AguaValida = 0
        .AttackedBy = ""
        .Attacking = 0
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .Invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .UseAINow = False
        .AtacaAPJ = 0
        .AtacaANPC = 0
        .AIAlineacion = e_Alineacion.ninguna
    End With
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Char.Body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.Heading = 0
Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = ""
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones: Npclist(NpcIndex).Expresiones(j) = "": Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Attackable = 0
    Npclist(NpcIndex).CanAttack = 0
    Npclist(NpcIndex).Comercia = 0
    Npclist(NpcIndex).GiveEXP = 0
    Npclist(NpcIndex).GiveGLD = 0
    Npclist(NpcIndex).GivePTS = 0
    Npclist(NpcIndex).GiveGLDMin = 0
    Npclist(NpcIndex).GiveGLDMax = 0
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).Inflacion = 0
    Npclist(NpcIndex).InvReSpawn = 0
    Npclist(NpcIndex).level = 0
    
    If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)
    
    Npclist(NpcIndex).MaestroUser = 0
    Npclist(NpcIndex).MaestroNpc = 0
    
    Npclist(NpcIndex).Mascotas = 0
    Npclist(NpcIndex).Movement = 0
    Npclist(NpcIndex).Name = "NPC SIN INICIAR"
    Npclist(NpcIndex).NPCtype = 0
    Npclist(NpcIndex).Numero = 0
    Npclist(NpcIndex).Orig.Map = 0
    Npclist(NpcIndex).Orig.X = 0
    Npclist(NpcIndex).Orig.Y = 0
    Npclist(NpcIndex).PoderAtaque = 0
    Npclist(NpcIndex).PoderEvasion = 0
    Npclist(NpcIndex).Pos.Map = 0
    Npclist(NpcIndex).Pos.X = 0
    Npclist(NpcIndex).Pos.Y = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNPC = 0
    Npclist(NpcIndex).TipoItems = 0
    Npclist(NpcIndex).Veneno = 0
    Npclist(NpcIndex).Desc = ""
    
    
    Dim j As Integer
    For j = 1 To Npclist(NpcIndex).NroSpells
        Npclist(NpcIndex).Spells(j) = 0
    Next j
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo Errhandler

If NpcIndex = 0 Then Exit Sub

    Npclist(NpcIndex).flags.NPCActive = False
    
    If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
        Call EraseNPCChar(SendTarget.ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex)
    End If
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If

Exit Sub

Errhandler:
    'Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos) As Boolean
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y) Then
        TestSpawnTrigger = _
        MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And _
        MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And _
        MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex = 0 Then Exit Sub
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.Map = mapa 'mapa
        altpos.Map = mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(1, 100)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(1, 100)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 Then altpos.X = newpos.X
            If newpos.Y <> 0 Then altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
            
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.Map = newpos.Map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    Map = altpos.Map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).Pos.Map = Map
                    Npclist(nIndex).Pos.X = X
                    Npclist(nIndex).Pos.Y = Y
                    Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, X, Y)
                    Exit Sub
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.Map = newpos.Map
                        Npclist(nIndex).Pos.X = newpos.X
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(SendTarget.ToMap, 0, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        Map = newpos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
    End If
    
    If Npclist(nIndex).Numero <> 156 And Npclist(nIndex).Numero <> 157 And Npclist(nIndex).Numero <> 158 And Npclist(nIndex).Numero <> 181 And Npclist(nIndex).Numero <> 182 Then
      Npclist(nIndex).DueñoMascota = 0
    End If
    
    'Crea el NPC
    Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, X, Y)

End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(Map, X, Y).NpcIndex = NpcIndex
    
    If sndRoute = SendTarget.ToMap Then
        Call ArgegarNpc(NpcIndex)
        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y & "," & Npclist(NpcIndex).Char.WeaponAnim & "," & Npclist(NpcIndex).Char.ShieldAnim & ",,," & Npclist(NpcIndex).Char.CascoAnim & ",,,," & Npclist(NpcIndex).Char.AuraA & "," & Npclist(NpcIndex).Numero)
    End If

End Sub

Sub ChangeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)

If NpcIndex > 0 Then
    Npclist(NpcIndex).Char.Body = Body
    Npclist(NpcIndex).Char.Head = Head
    Npclist(NpcIndex).Char.Heading = Heading
    If sndRoute = SendTarget.ToMap Then
        Call SendToNpcArea(NpcIndex, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
    End If
End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los cliente
If sndRoute = SendTarget.ToMap Then
    Call SendToNpcArea(NpcIndex, "DP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Pos.X & "," & Npclist(NpcIndex).Pos.Y)
Else
    Call SendData(sndRoute, sndIndex, sndMap, "DP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Pos.X & "," & Npclist(NpcIndex).Pos.Y)
End If

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If Npclist(NpcIndex).Numero = 621 Or Npclist(NpcIndex).Numero = 620 Then Exit Sub
    
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
        ' es una posicion legal
        If LegalPos(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida = 1) Then
        
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            
            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).Char.Heading = nHeading
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        End If
Else ' No es mascota
        ' Controlamos que la posicion sea legal, los npc que
        ' no son mascotas tienen mas restricciones de movimiento.
        If LegalPosNPC(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then
            
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            
            '[Alejo-18-5]
            'server
            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).Char.Heading = nHeading
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        Else
            If Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                Npclist(NpcIndex).PFINFO.PathLenght = 0
            End If
        
        End If
    End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo Errhandler

Dim loopC As Integer
  
For loopC = 1 To MAXNPCS + 1
    If loopC > MAXNPCS Then Exit For
    If Not Npclist(loopC).flags.NPCActive Then Exit For
Next loopC
  
NextOpenNPC = loopC


Exit Function
Errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal userindex As Integer)

Dim n As Integer
n = RandomNumber(1, 100)
If n < 30 Then
    UserList(userindex).flags.Envenenado = 1
    Call SendData(SendTarget.toindex, userindex, 0, "||55")
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'Crea un NPC del tipo Npcindex

Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim it As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

it = 0

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

Do While Not PosicionValida
        
        Call ClosestLegalPos(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If Npclist(nIndex).flags.TierraInvalida Then
            If LegalPos(newpos.Map, newpos.X, newpos.Y, True) Then _
                PosicionValida = True
        Else
            If LegalPos(newpos.Map, newpos.X, newpos.Y, False) Or LegalPos(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) Then _
                PosicionValida = True
        End If
        
        If PosicionValida Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = newpos.Map
            Npclist(nIndex).Pos.X = newpos.X
            Npclist(nIndex).Pos.Y = newpos.Y
        Else
            newpos.X = 0
            newpos.Y = 0
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = 0
            Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
Loop

'asignamos las nuevas coordenas
Map = newpos.Map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y

If Npclist(nIndex).Numero <> 156 And Npclist(nIndex).Numero <> 157 And Npclist(nIndex).Numero <> 158 And Npclist(nIndex).Numero <> 181 And Npclist(nIndex).Numero <> 182 Then
    Npclist(nIndex).DueñoMascota = 0
End If

Npclist(nIndex).Char.AuraA = 0

'Crea el NPC
Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, X, Y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, Map, "TW" & SND_WARP)
    Call SendData(SendTarget.ToNPCArea, nIndex, Map, "CFX" & Npclist(nIndex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
End If

SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)

If MiNPC.flags.Respawn = 0 Then Exit Sub
If MiNPC.Pos.Map = 18 Or MiNPC.Pos.Map = 180 Or MiNPC.Pos.Map = 181 Or MiNPC.Pos.Map = 160 Or MiNPC.Pos.Map = 170 Or MiNPC.Pos.Map = 100 Or MiNPC.Pos.Map = 54 Or MiNPC.Pos.Map = 72 Or MiNPC.Pos.Map = 110 Or MiNPC.Pos.Map = 8 Or MiNPC.Pos.Map = 118 Or MiNPC.Pos.Map = 99 Or MiNPC.Pos.Map = 120 Or MiNPC.Pos.Map = 93 Then Exit Sub

If (MiNPC.flags.Respawn = 1) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).Pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function
Sub NPCTirarOro(MiNPC As npc, userindex As Integer)

'SI EL NPC TIENE ORO LO TIRAMOS
If MiNPC.GiveGLDMin > 0 And MiNPC.GiveGLDMax > 0 Then
    Dim cantidaddeoro As Long
    Dim LuzyRiqueza As Long
    cantidaddeoro = RandomNumber(MiNPC.GiveGLDMin * MultiplicadorOro, MiNPC.GiveGLDMax * MultiplicadorOro)
    
    LuzyRiqueza = cantidaddeoro / 4
    
If UserList(userindex).flags.GemaActivada = "Naranja" Then
    cantidaddeoro = cantidaddeoro * 1.5
End If
    
If UserList(userindex).flags.estado = 1 Then
    cantidaddeoro = cantidaddeoro * 1.5
End If

If UserList(userindex).Invent.ArmourEqpObjIndex = 1049 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1050 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1456 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1497 Then
    cantidaddeoro = cantidaddeoro + LuzyRiqueza
End If
    
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + cantidaddeoro
    Call SendData(SendTarget.toindex, userindex, 0, "||56@" & PonerPuntos(cantidaddeoro))
    SendUserGLD (userindex)
End If

End Sub
Sub NPCDarPuntos(MiNPC As npc, ByVal userindex As Integer)

If MiNPC.GivePTS > 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||57@" & MiNPC.GivePTS)
    Call AgregarPuntos(userindex, MiNPC.GivePTS)
End If

End Sub

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer los NPCS se deberá usar la
'nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

Dim NpcIndex As Integer
Dim npcfile As String
Dim Leer As clsIniReader

If NpcNumber > 499 Then
        'NpcFile = DatPath & "NPCs-HOSTILES.dat"
        Set Leer = LeerNPCsHostiles
Else
        'NpcFile = DatPath & "NPCs.dat"
        Set Leer = LeerNPCs
End If

NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = Leer.GetValue("NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "EscudoAnim"))
Npclist(NpcIndex).Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "ArmaAnim"))
Npclist(NpcIndex).Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))

Npclist(NpcIndex).Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))

'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

Npclist(NpcIndex).Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
Npclist(NpcIndex).GivePTS = val(Leer.GetValue("NPC" & NpcNumber, "GivePTS"))
Npclist(NpcIndex).GiveGLDMin = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLDMin"))
Npclist(NpcIndex).GiveGLDMax = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLDMax"))

'Cristales
Npclist(NpcIndex).Cristales = val(Leer.GetValue("NPC" & NpcNumber, "Cristales"))
Npclist(NpcIndex).CristalesPequesMin = val(Leer.GetValue("NPC" & NpcNumber, "CristalesPequesMin"))
Npclist(NpcIndex).CristalesPequesMax = val(Leer.GetValue("NPC" & NpcNumber, "CristalesPequesMax"))
Npclist(NpcIndex).CristalesMedianosMin = val(Leer.GetValue("NPC" & NpcNumber, "CristalesMedianosMin"))
Npclist(NpcIndex).CristalesMedianosMax = val(Leer.GetValue("NPC" & NpcNumber, "CristalesMedianosMax"))
Npclist(NpcIndex).CristalesGrandesMin = val(Leer.GetValue("NPC" & NpcNumber, "CristalesGrandesMin"))
Npclist(NpcIndex).CristalesGrandesMax = val(Leer.GetValue("NPC" & NpcNumber, "CristalesGrandesMax"))
Npclist(NpcIndex).CristalesEpicosMin = val(Leer.GetValue("NPC" & NpcNumber, "CristalesEpicosMin"))
Npclist(NpcIndex).CristalesEpicosMax = val(Leer.GetValue("NPC" & NpcNumber, "CristalesEpicosMax"))
'Npclist(NpcIndex).MVP = Leer.GetValue("NPC" & NpcNumber, "MVP")


Npclist(NpcIndex).PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))

Npclist(NpcIndex).InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))


Npclist(NpcIndex).Stats.MaxHP = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))


Dim loopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
For loopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & loopC)
    Npclist(NpcIndex).Invent.Object(loopC).ProbTirar = val(ReadField(3, ln, 45))
    Npclist(NpcIndex).Invent.Object(loopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(loopC).Amount = val(ReadField(2, ln, 45))
Next loopC

Npclist(NpcIndex).flags.LanzaFlecha = val(Leer.GetValue("NPC" & NpcNumber, "LanzaFlecha"))

Npclist(NpcIndex).flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For loopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(loopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & loopC))
Next loopC


If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
    Npclist(NpcIndex).NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For loopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(loopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & loopC)
        Npclist(NpcIndex).Criaturas(loopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & loopC)
    Next loopC
End If


Npclist(NpcIndex).Inflacion = val(Leer.GetValue("NPC" & NpcNumber, "Inflacion"))

Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
Npclist(NpcIndex).flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
Npclist(NpcIndex).flags.AfectaRelampago = val(Leer.GetValue("NPC" & NpcNumber, "AfectaRelampago"))
Npclist(NpcIndex).flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
If aux = "" Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For loopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(loopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & loopC)
    Next loopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC = NpcIndex

End Function


Sub EnviarListaCriaturas(ByVal userindex As Integer, ByVal NpcIndex)
  Dim SD As String
  Dim k As Integer
  SD = SD & Npclist(NpcIndex).NroCriaturas & ","
  For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
  Next k
  SD = "LSTCRI" & SD
  Call SendData(SendTarget.toindex, userindex, 0, SD)
End Sub


Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = ""
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = TipoAI.SigueAmo 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNPC = 0

End Sub
Public Sub MuereRey(ByVal userindex As Integer, NpcIndex As Integer)
Dim lachotadelalana As Long
Dim reNpcPos As WorldPos
Dim reNpcIndex As Integer
Dim Castillo As Integer
Dim CastillosConquistados As Long
Castillo = 0
If UserList(userindex).Pos.Map = MapCastilloN Then Castillo = 1
If UserList(userindex).Pos.Map = MapCastilloS Then Castillo = 2
If UserList(userindex).Pos.Map = MapCastilloE Then Castillo = 3
If UserList(userindex).Pos.Map = MapCastilloO Then Castillo = 4
If UserList(userindex).Pos.Map = 81 Then Castillo = 5
If Castillo = 0 Then Exit Sub
 

reNpcPos.Map = Npclist(NpcIndex).Pos.Map
reNpcPos.X = Npclist(NpcIndex).Pos.X
reNpcPos.Y = Npclist(NpcIndex).Pos.Y
reNpcIndex = NpcIndex
 
If Castillo = 1 Then
   CastilloNorte = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(ToAll, 0, 0, "||58@" & (Guilds(UserList(userindex).GuildIndex).GuildName) & "@" & MapCastilloN)
   
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte", Guilds(UserList(userindex).GuildIndex).GuildName)
   CastilloNorte = Guilds(UserList(userindex).GuildIndex).GuildName
   Dim pija As Long
   pija = Guilds(UserList(userindex).GuildIndex).CASTIS + 1
   Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CASTIS", pija)
   Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 2 Then
   CastilloSur = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(ToAll, 0, 0, "||58@" & (Guilds(UserList(userindex).GuildIndex).GuildName) & "@" & MapCastilloS)
   
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur", Guilds(UserList(userindex).GuildIndex).GuildName)
   CastilloSur = Guilds(UserList(userindex).GuildIndex).GuildName
   pija = Guilds(UserList(userindex).GuildIndex).CASTIS + 1
   Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CASTIS", pija)
   Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 3 Then
   CastilloEste = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(ToAll, 0, 0, "||58@" & (Guilds(UserList(userindex).GuildIndex).GuildName) & "@" & MapCastilloE)
   
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste", Guilds(UserList(userindex).GuildIndex).GuildName)
   CastilloEste = Guilds(UserList(userindex).GuildIndex).GuildName
   pija = Guilds(UserList(userindex).GuildIndex).CASTIS + 1
   Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CASTIS", pija)
   Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 4 Then
   CastilloOeste = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(ToAll, 0, 0, "||58@" & (Guilds(UserList(userindex).GuildIndex).GuildName) & "@" & MapCastilloO)
   
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste", Guilds(UserList(userindex).GuildIndex).GuildName)
   CastilloOeste = Guilds(UserList(userindex).GuildIndex).GuildName
   pija = Guilds(UserList(userindex).GuildIndex).CASTIS + 1
   Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CASTIS", pija)
   Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 5 Then
   Fortaleza = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(ToAll, 0, 0, "||59@" & (Guilds(UserList(userindex).GuildIndex).GuildName))
   
   '#Reparamos rejas
         MapData(81, 49, 84).OBJInfo.ObjIndex = 1471
         Call ModAreas.SendToAreaByPos(81, 49, 84, "HO" & ObjData(1471).GrhIndex & "," & 49 & "," & 84)
                        'Bloquea
                        MapData(81, 49, 84).Blocked = 1
                        MapData(81, 49 - 1, 84).Blocked = 1
                        MapData(81, 49 - 2, 84).Blocked = 1
                        MapData(81, 49 + 1, 84).Blocked = 1
                        MapData(81, 49 + 2, 84).Blocked = 1
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49, 84, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 - 1, 84, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 - 2, 84, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 + 1, 84, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 + 2, 84, 1)
                        
                        
         MapData(81, 49, 48).OBJInfo.ObjIndex = 1471
         Call ModAreas.SendToAreaByPos(81, 49, 48, "HO" & ObjData(1471).GrhIndex & "," & 49 & "," & 48)
         
                        'Desbloquea
                        MapData(81, 49, 48).Blocked = 1
                        MapData(81, 49 - 1, 48).Blocked = 1
                        MapData(81, 49 - 2, 48).Blocked = 1
                        MapData(81, 49 + 1, 48).Blocked = 1
                        MapData(81, 49 + 2, 48).Blocked = 1
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49, 48, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 - 1, 48, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 - 2, 48, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 + 1, 48, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 + 2, 48, 1)
                        
                        
         MapData(81, 49, 68).OBJInfo.ObjIndex = 1471
         Call ModAreas.SendToAreaByPos(81, 49, 68, "HO" & ObjData(1471).GrhIndex & "," & 49 & "," & 68)
         
                        'Desbloquea
                        MapData(81, 49, 68).Blocked = 1
                        MapData(81, 49 - 1, 68).Blocked = 1
                        MapData(81, 49 - 2, 68).Blocked = 1
                        MapData(81, 49 + 1, 68).Blocked = 1
                        MapData(81, 49 + 2, 68).Blocked = 1
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49, 68, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 - 1, 68, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 - 2, 68, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 + 1, 68, 1)
                        Call Bloquear(SendTarget.ToMap, 0, 81, 81, 49 + 2, 68, 1)
   '#Reparamos rejas
   
   RejaSur = 10000
   RejaNorte = 10000
   RejaCentral = 10000
   RejaCentralAtacada = False
   RejaSurAtacada = False
   RejaNorteAtacada = False
   
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "Fortaleza", Guilds(UserList(userindex).GuildIndex).GuildName)
   pija = Guilds(UserList(userindex).GuildIndex).CASTIS + 1
   Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "CASTIS", pija)
   Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
End If

For lachotadelalana = 1 To LastUser
   If UserList(lachotadelalana).Pos.Map = MapCastilloE Or UserList(lachotadelalana).Pos.Map = MapCastilloN Or UserList(lachotadelalana).Pos.Map = MapCastilloS Or UserList(lachotadelalana).Pos.Map = MapCastilloO Or UserList(lachotadelalana).Pos.Map = 81 Then
    Call WarpUserChar(lachotadelalana, UserList(lachotadelalana).Pos.Map, UserList(lachotadelalana).Pos.X, UserList(lachotadelalana).Pos.Y, False)
   End If
Next lachotadelalana

Call QuitarNPC(NpcIndex)

If Castillo = 5 Then
    Call SpawnNpc(ReyNpcF, reNpcPos, True, False)
Else
    Call SpawnNpc(ReyNpcN, reNpcPos, True, False)
End If

   pija = Guilds(UserList(userindex).GuildIndex).GetReputacion + 75
   Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "REPU", pija)
   
   Call GRANK_User_Check(Castles, Guilds(UserList(userindex).GuildIndex).GuildName, Guilds(UserList(userindex).GuildIndex).CASTIS)
   Call GRANK_User_Check(GuildReputation, Guilds(UserList(userindex).GuildIndex).GuildName, Guilds(UserList(userindex).GuildIndex).GetReputacion)

End Sub

