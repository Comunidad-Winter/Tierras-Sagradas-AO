Attribute VB_Name = "modHechizos"
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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal userindex As Integer, ByVal spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Or UserList(userindex).flags.Privilegios > PlayerType.User Then Exit Sub

If UserList(userindex).GuildIndex <> 0 Then
    If Npclist(NpcIndex).Numero = 614 And UserList(userindex).Pos.Map = 81 And UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(Fortaleza) Then Exit Sub
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloN And UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(CastilloNorte) Then Exit Sub
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloS And UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(CastilloSur) Then Exit Sub
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloE And UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(CastilloEste) Then Exit Sub
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloO And UCase$(Guilds(UserList(userindex).GuildIndex).GuildName) = UCase$(CastilloOeste) Then Exit Sub
End If

Npclist(NpcIndex).CanAttack = 0
Dim Daño As Integer

If Hechizos(spell).SubeHP = 1 Then

    Daño = RandomNumber(Hechizos(spell).MinHP, Hechizos(spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Hechizos(spell).WAV)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(spell).FXgrh & "," & Hechizos(spell).loops)

    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + Daño
    If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call SendData(SendTarget.ToPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "N|" & vbCyan & "°" & Hechizos(spell).PalabrasMagicas & "°" & Npclist(NpcIndex).Char.CharIndex)
    Call SendData(SendTarget.toindex, userindex, 0, "||148@" & Npclist(NpcIndex).Name & "@" & Daño)
    Call SendUserHP(userindex)

ElseIf Hechizos(spell).SubeHP = 2 Then
    
    If UserList(userindex).flags.Privilegios = PlayerType.User Then
    
        Daño = RandomNumber(Hechizos(spell).MinHP, Hechizos(spell).MaxHP)
        
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
         If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        
        If Daño < 0 Then Daño = 0
        
        If Npclist(NpcIndex).Numero = 93 Then Daño = RandomNumber(30, 50)
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Hechizos(spell).WAV)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(spell).FXgrh & "," & Hechizos(spell).loops)
    
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - Daño
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbGreen & "°" & Hechizos(spell).PalabrasMagicas & "°" & str$(Npclist(NpcIndex).Char.CharIndex))
        
        Call SendData(SendTarget.toindex, userindex, 0, "||830@" & Npclist(NpcIndex).Name & "@" & Daño)
        Call SendUserHP(val(userindex))
        
        'Muere
        If UserList(userindex).Stats.MinHP < 1 Then
            UserList(userindex).Stats.MinHP = 0

        If userindex = GranPoder Then
        GranPoder = 0
        UserList(userindex).flags.GranPoder = 0
        Call OtorgarGranPoder(0)
        SendUserVariant (userindex)
    End If
            Call UserDie(userindex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(userindex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(userindex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(spell).Paraliza = 1 Then
     If UserList(userindex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Hechizos(spell).WAV)
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(spell).FXgrh & "," & Hechizos(spell).loops)
            
            Dim ProbabilidadInmunidad As Byte
            ProbabilidadInmunidad = RandomNumber(1, 8)
            If ProbabilidadInmunidad = 6 Then
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 1556 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||831")
                    Exit Sub
                End If
            End If
            
            
            UserList(userindex).flags.Paralizado = 1
          UserList(userindex).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
            Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
     End If
End If

If Npclist(NpcIndex).flags.LanzaFlecha = 1 Then
    Call SendData(SendTarget.ToMap, 0, Npclist(NpcIndex).Pos.Map, "FLECHI" & Npclist(NpcIndex).Char.CharIndex & "," & UserList(userindex).Char.CharIndex & "," & 753)
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If Npclist(TargetNPC).Numero = 621 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0

Dim Daño As Integer

If Hechizos(spell).SubeHP = 2 Then
    
        Daño = RandomNumber(Hechizos(spell).MinHP, Hechizos(spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "TW" & Hechizos(spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "CFX" & Npclist(TargetNPC).Char.CharIndex & "," & Hechizos(spell).FXgrh & "," & Hechizos(spell).loops)
        
    If Npclist(NpcIndex).MaestroUser > 0 Then
      If GranPoder = Npclist(NpcIndex).MaestroUser Then
        Call SendData(SendTarget.ToPCArea, Npclist(NpcIndex).MaestroUser, UserList(Npclist(NpcIndex).MaestroUser).Pos.Map, "N|" & vbGreen & "°" & Hechizos(spell).PalabrasMagicas & "°" & str$(Npclist(NpcIndex).Char.CharIndex))
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - Daño * 2
        Call SendData(SendTarget.ToPCArea, Npclist(NpcIndex).MaestroUser, Npclist(Npclist(NpcIndex).MaestroUser).Pos.Map, "N|" & vbYellow & "°-" & Daño * 2 & "°" & str(Npclist(TargetNPC).Char.CharIndex))
        Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, Round(Daño, 0))
      Else
        Call SendData(SendTarget.ToPCArea, Npclist(NpcIndex).MaestroUser, UserList(Npclist(NpcIndex).MaestroUser).Pos.Map, "N|" & vbGreen & "°" & Hechizos(spell).PalabrasMagicas & "°" & str$(Npclist(NpcIndex).Char.CharIndex))
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - Daño
        Call SendData(SendTarget.ToPCArea, Npclist(NpcIndex).MaestroUser, Npclist(Npclist(NpcIndex).MaestroUser).Pos.Map, "N|" & vbYellow & "°-" & Daño & "°" & str(Npclist(TargetNPC).Char.CharIndex))
        Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, Round(Daño, 0))
      End If
    Else
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - Daño
    End If
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal userindex As Integer) As Boolean

On Error GoTo Errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal userindex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, userindex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||181")
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||832@" & Hechizos(hIndex).Nombre)
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, CByte(Slot), 1)
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||182")
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal userindex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(userindex).Char.CharIndex
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbCyan & "°" & s & "°" & ind)
    Exit Sub
End Sub

Function PuedeLanzar(ByVal userindex As Integer, ByVal HechizoIndex As Integer) As Boolean

If (Len(Hechizos(HechizoIndex).ExclusivoClase) > 0 And Hechizos(HechizoIndex).ExclusivoClase <> UCase$(UserList(userindex).clase)) And (Len(Hechizos(HechizoIndex).ExclusivoClasedos) > 0 And Hechizos(HechizoIndex).ExclusivoClasedos <> UCase$(UserList(userindex).clase)) Then
Call SendData(SendTarget.toindex, userindex, 0, "||833")
PuedeLanzar = False
Exit Function
End If

If UserList(userindex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(userindex).flags.TargetMap
    wp2.X = UserList(userindex).flags.TargetX
    wp2.Y = UserList(userindex).flags.TargetY
        
    If UserList(userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(userindex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||17")
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||834")
            PuedeLanzar = False
        End If
    Else
            Call SendData(SendTarget.toindex, userindex, 0, "||18")
            PuedeLanzar = False
    End If
Else
   Call SendData(SendTarget.toindex, userindex, 0, "||3")
   PuedeLanzar = False
End If

If MapData(wp2.Map, wp2.X, wp2.Y).userindex = 0 And MapData(wp2.Map, wp2.X, wp2.Y + 1).userindex = 0 And MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex = 0 And MapData(wp2.Map, wp2.X, wp2.Y + 1).NpcIndex = 0 And Hechizos(HechizoIndex).RemueveInvisibilidadParcial = 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||19")
End If

End Function
Sub HechizoTerrenoEstado(ByVal userindex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(userindex).flags.TargetX
    PosCasteadaY = UserList(userindex).flags.TargetY
    PosCasteadaM = UserList(userindex).flags.TargetMap
    
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).userindex > 0 Then
                        'hay un user
                        If MapData(PosCasteadaM, TempX, TempY).userindex <> userindex And UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.AdminInvisible = 0 Then
                            UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Invisible = 0
                            UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Counters.Invisibilidad = 0
                            Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Char.CharIndex & ",0")
                            Call SendData(SendTarget.toindex, MapData(PosCasteadaM, TempX, TempY).userindex, 0, "INVI0")
                            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    End If

End Sub
Sub InvocarMascota(ByVal userindex As Integer, ByRef b As Boolean)
 
Dim hechi As Integer, Fer As Integer
Dim TargetPosicion As WorldPos
hechi = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
 
TargetPosicion.Map = UserList(userindex).flags.TargetMap
TargetPosicion.X = UserList(userindex).flags.TargetX
TargetPosicion.Y = UserList(userindex).flags.TargetY
 
    If Hechizos(hechi).NumNPC = 156 Or Hechizos(hechi).NumNPC = 157 Or Hechizos(hechi).NumNPC = 158 Or Hechizos(hechi).NumNPC = 181 Or Hechizos(hechi).NumNPC = 182 Then
        If UserList(userindex).flags.InvocoMascota = 1 Then
            Call QuitarNPC(UserList(userindex).flags.MascotinIndex)
            UserList(userindex).flags.InvocoMascota = 0
        Exit Sub
        End If
    End If
    
        If UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 151 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Or UserList(userindex).Pos.Map = 141 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||20")
        Exit Sub
        End If
        
        If MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).userindex > 0 Or MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).NpcIndex > 0 Or MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).OBJInfo.ObjIndex > 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||20")
        Exit Sub
        End If
        
        Fer = SpawnNpc(Hechizos(hechi).NumNPC, TargetPosicion, True, False)
        If Hechizos(hechi).NumNPC = 156 Or Hechizos(hechi).NumNPC = 157 Or Hechizos(hechi).NumNPC = 158 Or Hechizos(hechi).NumNPC = 181 Or Hechizos(hechi).NumNPC = 182 Then
        If UserList(userindex).flags.Montando = 1 Then
            Call Desmontar(userindex)
        End If
            UserList(userindex).flags.InvocoMascota = 1
        End If
       
        If Fer > 0 Then
            UserList(userindex).flags.MascotinIndex = Fer
            Call DoFollow(Fer, UserList(userindex).Name)
            Npclist(Fer).DueñoMascota = userindex
        End If
 
 
Call InfoHechizo(userindex)
b = True
 
End Sub
Sub HechizoInvocacion(ByVal userindex As Integer, ByRef b As Boolean)

If UserList(userindex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If UserList(userindex).Pos.Map <> 71 And UserList(userindex).Pos.Map <> 100 And UserList(userindex).Pos.Map <> 107 And UserList(userindex).Pos.Map <> 118 And UserList(userindex).Pos.Map <> 109 And UserList(userindex).Pos.Map <> 120 And UserList(userindex).Pos.Map <> 106 And UserList(userindex).Pos.Map <> 108 And UserList(userindex).Pos.Map <> 110 Then
    If MapInfo(UserList(userindex).Pos.Map).Pk = False Or MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call SendData(SendTarget.toindex, userindex, 0, "||21")
        Exit Sub
    End If
End If

Dim h As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(userindex).flags.TargetMap
TargetPos.X = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    If Hechizos(h).NumNPC = 94 Then
    If UserList(userindex).flags.EleDeTierra = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||22")
    Exit Sub
    End If
    End If
    If Hechizos(h).NumNPC = 92 Then
    If UserList(userindex).flags.EleDeAgua = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||23")
    Exit Sub
    End If
    End If
    If Hechizos(h).NumNPC = 93 Then
    If UserList(userindex).flags.EleDeFuego = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||24")
    Exit Sub
    End If
    End If
For j = 1 To Hechizos(h).Cant
    
    If UserList(userindex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(h).NumNPC, TargetPos, True, False)
        If Hechizos(h).NumNPC = 92 Then
            UserList(userindex).flags.EleDeAgua = 1
            UserList(userindex).Counters.TiempoElemental = 0
        End If
        If Hechizos(h).NumNPC = 93 Then
            UserList(userindex).flags.EleDeFuego = 1
        End If
        If Hechizos(h).NumNPC = 94 Then
            UserList(userindex).flags.EleDeTierra = 1
        End If
        If ind > 0 Then
            UserList(userindex).NroMacotas = UserList(userindex).NroMacotas + 1
            
            index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(index) = ind
            UserList(userindex).MascotasType(index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = userindex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            Npclist(ind).GivePTS = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(userindex)
b = True


End Sub
 
Sub HechizoBurbujaDefensiva(ByVal U As Integer, h As Integer, ByRef b As Boolean)
 
Dim T As WorldPos
 
T.Map = UserList(U).flags.TargetMap
T.X = UserList(U).flags.TargetX
T.Y = UserList(U).flags.TargetY
 
If MapData(T.Map, T.X, T.Y).NpcIndex > 0 Then Exit Sub
 
If UserList(U).flags.TargetUser <= 0 Then
    SendData SendTarget.toindex, U, 0, "||25"
Exit Sub
End If
 
b = True
SendData SendTarget.toindex, U, 0, "TW" & 255
UserList(U).flags.IntervaloBurbu = 61 '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ACA ponen la cant. de tiempo que quieren que defienda
UserList(U).flags.DefensaBurbu = RandomNumber(Hechizos(h).MinDef1, Hechizos(h).MaxDef1)
 
Call InfoHechizo(U)
 
End Sub
Sub HandleHechizoTerreno(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
            Call HechizoInvocacion(userindex, b)
        Case TipoHechizo.uTeleporta
            Call HechizoTelep(userindex, b)
        Case TipoHechizo.uBurbuja
            Call HechizoBurbujaDefensiva(userindex, uh, b)
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(userindex, b)
        Case TipoHechizo.uInvocaMascota
            Call InvocarMascota(userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    
 If UserList(userindex).flags.Privilegios = PlayerType.User Then
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    
    Call SendUserMP(userindex)
    Call SendUserST(userindex)
 End If
End If


End Sub

Sub HandleHechizoUsuario(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(userindex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    
If UserList(userindex).flags.Privilegios = PlayerType.User Then
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
End If
    
    Call SendUserMP(userindex)
    Call SendUserST(userindex)
    UserList(userindex).flags.TargetUser = 0
End If
End Sub

Sub HandleHechizoNPC(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(userindex).flags.TargetNPC, uh, b, userindex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(userindex).flags.TargetNPC, userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).flags.TargetNPC = 0
    
    If UserList(userindex).flags.Privilegios = PlayerType.User Then
        UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
        If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
        UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
        If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    End If

    Call SendUserMP(userindex)
    Call SendUserST(userindex)
End If

End Sub


Sub LanzarHechizo(index As Integer, userindex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(userindex).Stats.UserHechizos(index)

        If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||26")
            Exit Sub
        End If
        
    If Hechizos(uh).CuartaJerarquia = 1 Then
        If UserList(userindex).flags.CJerarquia = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||28")
            Exit Sub
        End If
    
        If UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 18 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 111 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Or UserList(userindex).Pos.Map = 166 Or UserList(userindex).Pos.Map = 164 Or UserList(userindex).Pos.Map = 162 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||27")
            Exit Sub
        End If
    End If

If PuedeLanzar(userindex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case TargetType.uUsuarios
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||15")
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||25")
            End If
        Case TargetType.uNPC
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||15")
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||29")
            End If
        Case TargetType.uUsuariosYnpc
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||15")
                End If
        'Está el bot?
        ElseIf MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).BotIndex <> 0 Then
           'Checkeo que esté invocado.
           If ia_Bot(MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).BotIndex).Invocado Then
              'compruebo que este en mi grupo
              'If ia_Bot(bot_Index).GrupoID = UserList(UserIndex).Group_User.Grupo_ID Then
                 ia_UserDamage uh, MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).BotIndex, userindex, False
              'End If
           End If
        ElseIf MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY + 1).BotIndex <> 0 Then
           'Checkeo que esté invocado.
           If ia_Bot(MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY + 1).BotIndex).Invocado Then
              'compruebo que este en mi grupo
              'If ia_Bot(bot_Index).GrupoID = UserList(UserIndex).Group_User.Grupo_ID Then
                 ia_UserDamage uh, MapData(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY + 1).BotIndex, userindex, False
              'End If
           End If
            ElseIf UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||15")
                End If
            End If
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(userindex, uh)
        Case TargetType.uOnlyUsuario
            'Verifica si el objetivo es el user.
            If UserList(userindex).flags.TargetUser = userindex Then
                Call HandleHechizoUsuario(userindex, uh)
            Else
            'Si no es tira mensaje de error
                Call SendData(SendTarget.toindex, userindex, 0, "||30")
            End If
    End Select
    
End If

If UserList(userindex).Counters.Trabajando Then _
    UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal userindex As Integer, ByRef b As Boolean)



Dim h As Integer, TU As Integer
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
TU = UserList(userindex).flags.TargetUser

    If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||26")
        Exit Sub
    End If

If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then

    If userindex = TU Then
     Call SendData(SendTarget.toindex, userindex, 0, "||31")
    Exit Sub
    End If
   
    If UserList(userindex).flags.EspectadorArena1 = 1 Or UserList(userindex).flags.EspectadorArena2 = 1 Or UserList(userindex).flags.EspectadorArena3 = 1 Or UserList(userindex).flags.EspectadorArena4 = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||142")
        Exit Sub
    End If
   

     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(userindex, TU) Then Exit Sub
            
            If userindex <> TU Then
                Call UsuarioAtacadoPorUsuario(userindex, TU)
            End If
            
            Call InfoHechizo(userindex)
            b = True
            
            If UserList(TU).Invent.HerramientaEqpObjIndex = 1556 Then
                Dim ProbabilidadInmunidad As Byte
                ProbabilidadInmunidad = RandomNumber(1, 10)
                
                    If ProbabilidadInmunidad = 6 Then
                            Call SendData(SendTarget.toindex, TU, 0, "||831")
                            Call SendData(SendTarget.toindex, userindex, 0, "||835")
                        Exit Sub
                    End If
            End If
            
                UserList(TU).flags.Paralizado = 1
                UserList(TU).Counters.Paralisis = IntervaloParalizado
            
                Call SendData(SendTarget.toindex, TU, 0, "PARADOK")
                Call SendData(SendTarget.toindex, TU, 0, "PU" & UserList(TU).Pos.X & "," & UserList(TU).Pos.Y)
            
        Exit Sub
    End If
End If

If Hechizos(h).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
        
        If UserList(userindex).Counters.InmoManopla > 0 Then Exit Sub
    
        If EsHorda(TU) And EsAlianza(userindex) And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE Then
            Call SendData(SendTarget.toindex, userindex, 0, "||146")
         Exit Sub
        End If
        
        If EsHorda(userindex) And EsAlianza(TU) And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE Then
            Call SendData(SendTarget.toindex, userindex, 0, "||146")
         Exit Sub
        End If
        
            If UCase$(TModalidad) = "DM" And TU <> userindex And UserList(userindex).Pos.Map = 100 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||836")
               Exit Sub
             End If
    
        If UserList(userindex).flags.EspectadorArena1 = 1 Or UserList(userindex).flags.EspectadorArena2 = 1 Or UserList(userindex).flags.EspectadorArena3 = 1 Or UserList(userindex).flags.EspectadorArena4 = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||142")
            Exit Sub
        ElseIf UserList(TU).flags.EspectadorArena1 = 1 Or UserList(TU).flags.EspectadorArena2 = 1 Or UserList(TU).flags.EspectadorArena3 = 1 Or UserList(TU).flags.EspectadorArena4 = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||143")
            Exit Sub
        End If
     
        UserList(TU).flags.Paralizado = 0
        Call SendData(SendTarget.toindex, TU, 0, "PARADOK")
        
        Call InfoHechizo(userindex)
        b = True
    End If
End If



If Hechizos(h).Invisibilidad = 1 Then
   
    If UserList(TU).flags.Muerto = 1 Then
            If TU = userindex Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||144")
            End If
        b = False
        Exit Sub
    End If
    
If UserList(userindex).GuildIndex > 0 Then
    If UserList(TU).GuildIndex <> UserList(userindex).GuildIndex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||837")
    Exit Sub
    End If
    ElseIf TU <> userindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||837")
    Exit Sub
    End If
    
    If UserList(TU).Pos.Map = 100 Or UserList(TU).Pos.Map = 18 Or UserList(TU).Pos.Map = 166 Or UserList(TU).Pos.Map = 99 Or UserList(TU).Pos.Map = 118 Or UserList(TU).Pos.Map = 101 Or UserList(TU).Pos.Map = 54 Or UserList(TU).Pos.Map = 8 Or UserList(TU).Pos.Map = 72 Or UserList(TU).Pos.Map = 121 Or UserList(TU).Pos.Map = 94 Or UserList(TU).Pos.Map = 95 Or UserList(TU).Pos.Map = 31 Or UserList(TU).Pos.Map = 32 Or UserList(TU).Pos.Map = 33 Or UserList(TU).Pos.Map = 34 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||838")
    Exit Sub
    End If
    
    If MapData(UserList(TU).Pos.Map, UserList(TU).Pos.X, UserList(TU).Pos.Y).trigger = 6 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||838")
    Exit Sub
    End If
    
    If UserList(TU).flags.Invisible = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||839")
    Exit Sub
    End If
    
    UserList(TU).flags.Invisible = 1
    UserList(TU).Counters.Invisibilidad = 25
    Call SendData(SendTarget.ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
    Call InfoHechizo(userindex)
    b = True
End If

If Hechizos(h).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(userindex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Privilegios >= PlayerType.UserSupport Then
        Exit Sub
    End If
    
    If UserList(userindex).flags.Mimetizado = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||840")
        Exit Sub
    End If
    
    'copio el char original al mimetizado
    
    With UserList(userindex)
        .CharMimetizado.Body = .Char.Body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.Body = UserList(TU).Char.Body
        .Char.Head = UserList(TU).Char.Head
        .Char.CascoAnim = UserList(TU).Char.CascoAnim
        .Char.ShieldAnim = UserList(TU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(TU).Char.WeaponAnim
    
        Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(userindex)
   b = True
End If


If Hechizos(h).Envenena = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Maldicion = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(userindex)
        b = True
End If
                        
If Hechizos(h).Revivir = 1 Then

        If UserList(TU).flags.SeguroResu = True And UserList(TU).flags.Muerto = 1 Then
                Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "||842@" & UserList(userindex).Name)
                Call SendData(SendTarget.toindex, userindex, 0, "||841")
            b = False
            Exit Sub
        End If
        
    If UserList(TU).flags.TimeRevivir > 0 Then
        SendData SendTarget.toindex, userindex, 0, "||843@" & UserList(TU).flags.TimeRevivir
        Exit Sub
    End If
    
    If UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Then
        SendData SendTarget.toindex, userindex, 0, "||838"
        Exit Sub
    End If
    
    If EsAlianza(userindex) And EsHorda(TU) Then
        SendData SendTarget.toindex, userindex, 0, "||844"
        Exit Sub
    End If
    
    If EsHorda(userindex) And EsAlianza(TU) Then
        SendData SendTarget.toindex, userindex, 0, "||844"
        Exit Sub
    End If
    
    If UserList(TU).flags.Muerto = 1 Then
        b = True
        Call InfoHechizo(userindex)
        
        If UCase$(UserList(userindex).clase) = "CLERIGO" Then
            Call RevivirUsuario(TU)
            UserList(TU).Stats.MinHP = UserList(TU).Stats.MaxHP
            SendUserHP (TU)
            SendData SendTarget.toindex, TU, 0, "||749@" & UserList(userindex).Name
        Else
            UserList(TU).Counters.SegundosParaRevivir = 10
            SendData SendTarget.toindex, TU, 0, "||845"
        End If
        
        UserList(userindex).Stats.MinHP = 10
        SendUserHP (userindex)
    Else
        b = False
    End If

End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal userindex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||846")
    Exit Sub
End If

If Hechizos(hIndex).Paraliza = 1 Or Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Or Npclist(NpcIndex).Numero = ELEMENTALAGUA Then
        Call SendData(SendTarget.toindex, userindex, 0, "||846")
        Exit Sub
    End If
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||847")
        Exit Sub
   End If
        
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||847")
        Exit Sub
   End If
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If
    
If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Paralizado = 1
    Npclist(NpcIndex).flags.Inmovilizado = 0
    Npclist(NpcIndex).Contadores.Paralisis = 500
    b = True
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||848")
End If
 
If Hechizos(hIndex).Inmoviliza = 1 Then
 If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
    Npclist(NpcIndex).flags.Inmovilizado = 1
    Npclist(NpcIndex).flags.Paralizado = 1
    Npclist(NpcIndex).Contadores.Paralisis = 500
    Call InfoHechizo(userindex)
    b = True
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||848")
 End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal userindex As Integer, ByRef b As Boolean)
On Error Resume Next

Dim Daño As Long
Dim HechiCritico As Byte
HechiCritico = RandomNumber(1, 5)

    If Hechizos(hIndex).BacuNecesario = 1 And ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|Necesitas un báculo o cetro de mago para lanzar este hechizo.~255~255~0~1~0")
        Exit Sub
    End If

'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * 70)
    
    Call InfoHechizo(userindex)
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + Daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call SendData(SendTarget.toindex, userindex, 0, "||849@" & Daño)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then

    If Hechizos(hIndex).Nombre = "Relampago" Then
        If Npclist(NpcIndex).flags.AfectaRelampago = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||848")
        Exit Sub
        End If
    End If
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||847")
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * 70)
    
    Call CheckPets(NpcIndex, userindex, True)

If HechiCritico = 1 Or HechiCritico = 5 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||138")
    Daño = Daño * 2
End If

    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(userindex).clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                Daño = (Daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                Daño = Daño * 0.7 'Baja daño a 70% del original
            End If
        ElseIf UCase$(UserList(userindex).clase) = "DRUIDA" Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus = 0 Then
                    Daño = (Daño * 73) / 100
                Else
                    Daño = (Daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 65)) / 100
                End If
        End If
    End If

If Npclist(NpcIndex).MaestroUser > 0 Then
    Daño = Daño * 1.5
End If

Daño = Daño * 1.4

If userindex = GranPoder Then Daño = Daño * 1.8


    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DañoMagicoMin > 0 Then
        Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DañoMagicoMin, ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DañoMagicoMax)
    End If
    
    If Hechizos(hIndex).Nombre = "Big Bang Kame Hame Ha" Then
        Daño = Npclist(NpcIndex).Stats.MaxHP
    End If

    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbYellow & "°-" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))
    
    Call InfoHechizo(userindex)
    b = True
    Call NpcAtacado(NpcIndex, userindex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    If Hechizos(hIndex).Nombre = "Relampago" And UserList(userindex).Stats.ELV < 50 Then
        Call CalcularDarExp(userindex, NpcIndex, Daño)
    ElseIf Hechizos(hIndex).Nombre <> "Relampago" Then
        Call CalcularDarExp(userindex, NpcIndex, Daño)
    Else
        'nada
    End If
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño
    'SendData SendTarget.toindex, UserIndex, 0, "||Has lanzado " & Hechizos(hIndex).Nombre & " sobre la criatura." & FONTTYPE_ROJON
    
    SendData SendTarget.toindex, userindex, 0, "||850@" & Daño

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, userindex)
    End If
End If

If UserList(userindex).flags.Privilegios > PlayerType.UserSupport Then Call LogGMss(UserList(userindex).Name, "Tiró el hechizo " & Hechizos(hIndex).Nombre & " al npc " & Npclist(NpcIndex).Name & "", False)

End Sub

Sub InfoHechizo(ByVal userindex As Integer)


    Dim h As Integer
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, userindex)
    
    If UserList(userindex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(UserList(userindex).flags.TargetUser).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, UserList(userindex).Pos.Map, "TW" & Hechizos(h).WAV)
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).Pos.Map, "CFX" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, UserList(userindex).Pos.Map, "TW" & Hechizos(h).WAV)
    End If
    
    If UserList(userindex).flags.TargetUser > 0 Then
        If userindex <> UserList(userindex).flags.TargetUser Then
            Call SendData(SendTarget.toindex, userindex, 0, "N|" & Hechizos(h).HechizeroMsg & " " & UserList(UserList(userindex).flags.TargetUser).Name & "~255~0~0~1")
            Call SendData(SendTarget.toindex, UserList(userindex).flags.TargetUser, 0, "N|" & UserList(userindex).Name & " " & Hechizos(h).TargetMsg & "~255~0~0~1")
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "N|" & Hechizos(h).PropioMsg & "~255~0~0~1")
        End If
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|" & Hechizos(h).HechizeroMsg & " la criatura.~255~0~0~1")
    End If

End Sub

Sub HechizoPropUsuario(ByVal userindex As Integer, ByRef b As Boolean)

Dim h As Integer
Dim Daño As Integer
Dim tempChr As Integer
    
    
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
tempChr = UserList(userindex).flags.TargetUser
      
    If UserList(userindex).flags.EspectadorArena1 = 1 Or UserList(userindex).flags.EspectadorArena2 = 1 Or UserList(userindex).flags.EspectadorArena3 = 1 Or UserList(userindex).flags.EspectadorArena4 = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||142")
        Exit Sub
    End If
    
    If UserList(tempChr).flags.EspectadorArena1 = 1 Or UserList(tempChr).flags.EspectadorArena2 = 1 Or UserList(tempChr).flags.EspectadorArena3 = 1 Or UserList(tempChr).flags.EspectadorArena4 = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||143")
        Exit Sub
    End If
    
    If Hechizos(h).BacuNecesario = 1 And ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "N|Necesitas un báculo o cetro de mago para lanzar este hechizo.~255~255~0~1~0")
        Exit Sub
    End If
      
If UserList(tempChr).flags.Muerto = 1 And Hechizos(h).Revivir = 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||144")
    Exit Sub
End If

If Hechizos(h).ActivaNobleza = 1 Then
 If UserList(userindex).flags.EsNoble = 0 Then Exit Sub
 
    If UserList(userindex).flags.estado = 0 Then
        UserList(userindex).flags.estado = 1
        
        If UserList(userindex).flags.Navegando = 0 Then
            UserList(userindex).OrigChar.CascoAnim = UserList(userindex).Char.CascoAnim
            UserList(userindex).Char.CascoAnim = 32
            Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        End If
        
            SendUserVariant (userindex)
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXNOBLE & ",")
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 61 & "," & 0)
        Exit Sub
    Else
            UserList(userindex).flags.estado = 0
            
            If UserList(userindex).flags.Navegando = 0 Then
                Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).OrigChar.CascoAnim)
            End If
            
            SendUserVariant (userindex)
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXNOBLE & ",")
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 61 & "," & 0)
        Exit Sub
    End If
End If

' <-------- Agilidad ---------->
If Hechizos(h).SubeAgilidad = 1 Then
    
    Call InfoHechizo(userindex)
    Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 7000
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + Daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > 35 Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = 35
    SendUserAgilidad (tempChr)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    SendUserAgilidad (tempChr)
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(h).SubeFuerza = 1 Then
    Call InfoHechizo(userindex)
    Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + Daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > 35 Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = 35
    
    UserList(tempChr).flags.TomoPocion = True
    SendUserFuerza (tempChr)
    b = True
    
ElseIf Hechizos(h).SubeFuerza = 2 Then

    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    SendUserFuerza (tempChr)
    b = True
    
End If

'Salud
If Hechizos(h).SubeHP = 1 Then
    
    Daño = RandomNumber(30, 50)
    
    If UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP Then
        Call SendData(SendTarget.toindex, userindex, 0, "||145")
        Exit Sub
    End If
    
    If EsAlianza(tempChr) And EsHorda(userindex) And TriggerZonaPelea(tempChr, userindex) <> TRIGGER6_PERMITE Then
        Call SendData(SendTarget.toindex, userindex, 0, "||146")
        Exit Sub
    End If
    
    If EsHorda(tempChr) And EsAlianza(userindex) And TriggerZonaPelea(tempChr, userindex) <> TRIGGER6_PERMITE Then
        Call SendData(SendTarget.toindex, userindex, 0, "||146")
        Exit Sub
    End If
    
    Call InfoHechizo(userindex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + Daño
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    Call SendUserHP(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||147@" & Daño & "@" & UserList(tempChr).Name)
        Call SendData(SendTarget.toindex, tempChr, 0, "||148@" & UserList(userindex).Name & "@" & Daño)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||149@" & Daño)
    End If
    
    b = True
ElseIf Hechizos(h).SubeHP = 2 Then

    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub 'Esto cubre todos los ifs de mierda que borramos acá abajo
    
    If userindex = tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||31")
        Exit Sub
    End If

    Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * 70)
    
    If UserList(tempChr).flags.GemaActivada = "Celeste" Then
        Daño = Daño - (Daño * 10 / 100)
    End If

    'BONIFICADORES - Daño mágico/Defensa Mágica.
    If UserList(tempChr).Bon2 = "Aumenta en 3 puntos tu resistencia magica." And UserList(tempChr).Bon3 = "Aumenta en 3 puntos tu resistencia magica." Then
      Daño = Round(Daño - (Daño * 3 / 100)) 'menos 3%
    ElseIf UserList(tempChr).Bon3 = "Aumenta en 4 puntos tu resistencia magica." Or UserList(tempChr).Bon2 = "Aumenta en 4 puntos tu resistencia magica." Or UserList(tempChr).Bon2 = "Aumenta en 3 puntos tu resistencia magica." Or UserList(tempChr).Bon3 = "Aumenta en 3 puntos tu resistencia magica." Then
      Daño = Round(Daño - (Daño * 1.5 / 100)) 'menos 1.5%
    End If
    
    If UserList(userindex).Bon2 = "Aumenta tu daño magico." And UserList(userindex).Bon3 = "Aumenta tu daño magico." Then
        Daño = Round(Daño + (Daño * 3 / 100)) 'mas 3%
    ElseIf UserList(userindex).Bon2 = "Aumenta tu daño magico." Or UserList(userindex).Bon3 = "Aumenta tu daño magico." Then
        Daño = Round(Daño + (Daño * 1.5 / 100)) 'mas 1.5%
    End If
    'BONIFICADORES - Daño mágico/Defensa Mágica.
    
    
    If Hechizos(h).StaffAffected Then
        If UCase$(UserList(userindex).clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                Daño = (Daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                Daño = Daño * 0.7 'Baja daño a 70% del original
            End If
        ElseIf UCase$(UserList(userindex).clase) = "DRUIDA" Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus = 0 Then
                    Daño = (Daño * 73) / 100
                Else
                    Daño = (Daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 65)) / 100
                End If
        End If
    End If
    
    If UserList(userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        Daño = Round(Daño + (Daño * 4 / 100))
    End If
    
    'BALANCEO: Subimos/bajamos daño mágico del atacante.
    Daño = Round(Daño + (Daño * ModificarAtaqueMagico(UserList(userindex).clase) / 100))
    
    '/: Subimos/bajamos la defensa mágica del usuario que recibe
        Daño = Round(Daño - (Daño * ModificarDefensaMagica(UserList(tempChr).clase) / 100))
    'BALANCEO:
    
    If GranPoder = userindex Then Daño = Daño * 1.5
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
     'Armaduras antimagia
     If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
        Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DañoMagicoMin > 0 Then
        Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DañoMagicoMin, ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DañoMagicoMax)
    End If
    
    If Daño < 0 Then Daño = 0
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
    
    Call SendData(SendTarget.toindex, userindex, 0, "||150@" & Daño & "@" & UserList(tempChr).Name)
    Call SendData(SendTarget.toindex, tempChr, 0, "||151@" & UserList(userindex).Name & "@" & Daño)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & &HFFFF& & "°" & "- " & Daño & "" & "°" & str(UserList(tempChr).Char.CharIndex))
    
    Call AllMascotasAtacanUser(userindex, tempChr)
    Call AllMascotasAtacanUser(tempChr, userindex)
    
    Call SendUserMP(userindex)
    Call SendUserHP(tempChr)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
    
       If tempChr = GranPoder Then
            Call OtorgarGranPoder(userindex)
            UserList(tempChr).flags.GranPoder = 0
            SendUserVariant (tempChr)
        End If
        
        Call ContarMuerte(tempChr, userindex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, userindex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If


End Sub
Sub HechizoTelep(userindex As Integer, b As Boolean)
    
    Dim TU As Integer
    Dim h As Integer
    Dim i As Integer
    
    Dim PosTIROTELEPORT As WorldPos
    
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    If Hechizos(h).Nombre = "Teletransportar a Kahlimdor" Or Hechizos(h).Nombre = "Teletransportar a Anvilmar" Or Hechizos(h).Nombre = "Teleport a la Ciudad" Then
        If UserList(userindex).flags.TiroPortalL = 1 Then
            Call SendData(toindex, userindex, 0, "||32")
            Exit Sub
        End If
    End If
    
If UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 78 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||27")
Exit Sub
End If
    
If MapInfo(UserList(userindex).Pos.Map).Pk = False Then
    Call SendData(SendTarget.toindex, userindex, 0, "||33")
    Exit Sub
End If
    
If UserList(userindex).EnCvc = True Then Exit Sub
If UserList(userindex).flags.EnDesafio = 1 Then Exit Sub
If UserList(userindex).flags.EnDuelo = 1 Then Exit Sub
If UserList(userindex).flags.EnPareja = 1 Then Exit Sub

If UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 31 Then
SendData SendTarget.toindex, userindex, 0, "||34"
Exit Sub
End If
    
    If HayAgua(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||33")
                Exit Sub
            End If
            

    If MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).NpcIndex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX - 1, UserList(userindex).flags.TargetY).NpcIndex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY - 1).NpcIndex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX - 1, UserList(userindex).flags.TargetY - 1).NpcIndex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX + 1, UserList(userindex).flags.TargetY).NpcIndex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY + 1).NpcIndex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX + 1, UserList(userindex).flags.TargetY + 1).NpcIndex > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||33")
     Exit Sub
    ElseIf MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).userindex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX - 1, UserList(userindex).flags.TargetY).userindex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY - 1).userindex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX - 1, UserList(userindex).flags.TargetY - 1).userindex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX + 1, UserList(userindex).flags.TargetY).userindex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY + 1).userindex > 0 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX + 1, UserList(userindex).flags.TargetY + 1).userindex > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||33")
     Exit Sub
    ElseIf MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Blocked = 1 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX - 1, UserList(userindex).flags.TargetY).Blocked = 1 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY - 1).Blocked = 1 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX - 1, UserList(userindex).flags.TargetY - 1).Blocked = 1 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX + 1, UserList(userindex).flags.TargetY).Blocked = 1 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY + 1).Blocked = 1 Or MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX + 1, UserList(userindex).flags.TargetY + 1).Blocked = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||33")
     Exit Sub
    End If

            
    
    If Hechizos(h).Telepo = 1 Then
      
      MapaPortal = Hechizos(h).PortalMap
      XPortal = Hechizos(h).PortalX
      YPortal = Hechizos(h).PortalY
        
      
        PosTIROTELEPORT.X = UserList(userindex).flags.TargetX
        PosTIROTELEPORT.Y = UserList(userindex).flags.TargetY
        PosTIROTELEPORT.Map = UserList(userindex).flags.TargetMap
      
        UserList(userindex).flags.DondeTiroMap = PosTIROTELEPORT.Map
        UserList(userindex).flags.DondeTiroX = PosTIROTELEPORT.X
        UserList(userindex).flags.DondeTiroY = PosTIROTELEPORT.Y
      
        If MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).OBJInfo.ObjIndex Then
            Exit Sub
        End If
      
        If MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.Map Then
            Exit Sub
        End If
      
        If MapData(UserList(userindex).Pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Blocked Then
            Exit Sub
        End If
      
      
        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 0
      
        Call MakeObj(ToMap, userindex, UserList(userindex).Pos.Map, ET, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY)
        b = True
        
        UserList(userindex).Stats.MinMAN = 0
        SendUserMP (userindex)
                          
        UserList(userindex).Counters.TimeTeleport = 0
        UserList(userindex).Counters.CreoTeleport = True
        UserList(userindex).flags.TiroPortalL = 1
    End If
    
    Call SendData(toindex, userindex, 0, "||36")
    Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "PCF" & 28 & "," & UserList(userindex).flags.TargetX & "," & UserList(userindex).flags.TargetY & "," & 170)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Hechizos(h).WAV)
    Call InfoHechizo(userindex)
End Sub
Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(userindex, Slot, UserList(userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userindex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(userindex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(userindex, LoopC, UserList(userindex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(userindex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(userindex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(SendTarget.toindex, userindex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(SendTarget.toindex, userindex, 0, "SHS" & Slot & "," & "0" & "," & "(Nada)")

End If


End Sub


Public Sub DesplazarHechizo(ByVal userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||37")
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(SendTarget.toindex, userindex, 0, "||37")
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, userindex, CualHechizo)

End Sub
