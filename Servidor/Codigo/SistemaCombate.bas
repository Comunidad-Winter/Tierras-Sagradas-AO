Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18
Function ModificadorEvasion(ByVal clase As String) As Single

Select Case UCase$(clase)
    Case "GUERRERO"
        ModificadorEvasion = 0.8
    Case "CAZADOR"
        ModificadorEvasion = 0.9
    Case "PALADIN"
        ModificadorEvasion = 0.8
    Case "ASESINO"
        ModificadorEvasion = 0.9
    Case "LADRON"
        ModificadorEvasion = 1
    Case "BARDO"
        ModificadorEvasion = 0.9
    Case "CLERIGO"
        ModificadorEvasion = 0.8
    Case "MAGO"
        ModificadorEvasion = 0.1
    Case "DRUIDA"
        ModificadorEvasion = 0.75
    Case Else
        ModificadorEvasion = 0.8
End Select
End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As String) As Single
Select Case UCase$(clase)
    Case "GUERRERO"
        ModificadorPoderAtaqueArmas = 1
    Case "CAZADOR"
        ModificadorPoderAtaqueArmas = 0.8
    Case "PALADIN"
        ModificadorPoderAtaqueArmas = 0.93
    Case "ASESINO"
        ModificadorPoderAtaqueArmas = 0.9
    Case "LADRON"
        ModificadorPoderAtaqueArmas = 0.75
    Case "CLERIGO"
        ModificadorPoderAtaqueArmas = 0.85
    Case "BARDO"
        ModificadorPoderAtaqueArmas = 0.75
    Case "DRUIDA"
        ModificadorPoderAtaqueArmas = 0.72
    Case "ARTESANO"
        ModificadorPoderAtaqueArmas = 0.6
    Case "RECOLECTOR"
        ModificadorPoderAtaqueArmas = 0.6
    Case Else
        ModificadorPoderAtaqueArmas = 0.6
End Select
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As String) As Single
Select Case UCase$(clase)
    Case "GUERRERO"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "CAZADOR"
        ModificadorPoderAtaqueProyectiles = 1
    Case "PALADIN"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "ASESINO"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "LADRON"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "CLERIGO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "BARDO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "DRUIDA"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "MAGO"
        ModificadorPoderAtaqueProyectiles = 0.5
    Case "RECOLECTOR"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "ARTESANO"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case Else
        ModificadorPoderAtaqueProyectiles = 0.5
End Select
End Function

Function ModicadorDañoClaseArmas(ByVal clase As String) As Single
Select Case UCase$(clase)
    Case "GUERRERO"
        ModicadorDañoClaseArmas = 1.1
    Case "CAZADOR"
        ModicadorDañoClaseArmas = 0.9
    Case "PALADIN"
        ModicadorDañoClaseArmas = 0.925
    Case "ASESINO"
        ModicadorDañoClaseArmas = 0.9
    Case "LADRON"
        ModicadorDañoClaseArmas = 0.8
    Case "CLERIGO"
        ModicadorDañoClaseArmas = 0.8
    Case "BARDO"
        ModicadorDañoClaseArmas = 0.78
    Case "DRUIDA"
        ModicadorDañoClaseArmas = 0.665
    Case "RECOLECTOR"
        ModicadorDañoClaseArmas = 0.8
    Case "ARTESANO"
        ModicadorDañoClaseArmas = 0.8
    Case Else
        ModicadorDañoClaseArmas = 0.5
End Select
End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As String) As Single
Select Case UCase$(clase)
    Case "GUERRERO"
        ModicadorDañoClaseProyectiles = 0.96
    Case "CAZADOR"
        ModicadorDañoClaseProyectiles = 1.01
    Case "PALADIN"
        ModicadorDañoClaseProyectiles = 0.7
    Case "ASESINO"
        ModicadorDañoClaseProyectiles = 0.5
    Case "LADRON"
        ModicadorDañoClaseProyectiles = 0.8
    Case "CLERIGO"
        ModicadorDañoClaseProyectiles = 0.5
    Case "BARDO"
        ModicadorDañoClaseProyectiles = 0.8
    Case "DRUIDA"
        ModicadorDañoClaseProyectiles = 0.7
    Case "RECOLECTOR"
        ModicadorDañoClaseProyectiles = 0.6
    Case "ARTESANO"
        ModicadorDañoClaseProyectiles = 0.7
    Case Else
        ModicadorDañoClaseProyectiles = 0.5
End Select
End Function

Function ModEvasionDeEscudoClase(ByVal clase As String) As Single

Select Case UCase$(clase)
Case "GUERRERO"
        ModEvasionDeEscudoClase = 1
    Case "CAZADOR"
        ModEvasionDeEscudoClase = 0.8
    Case "PALADIN"
        ModEvasionDeEscudoClase = 0.9
    Case "ASESINO"
        ModEvasionDeEscudoClase = 0.8
    Case "LADRON"
        ModEvasionDeEscudoClase = 0.7
    Case "CLERIGO"
        ModEvasionDeEscudoClase = 0.8
    Case "BARDO"
        ModEvasionDeEscudoClase = 0.85
    Case "DRUIDA"
        ModEvasionDeEscudoClase = 0.75
    Case "RECOLECTOR"
        ModEvasionDeEscudoClase = 0.7
    Case "ARTESANO"
        ModEvasionDeEscudoClase = 0.7
    Case Else
        ModEvasionDeEscudoClase = 0.6
End Select

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MinimoInt = b
    Else: MinimoInt = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MaximoInt = a
    Else: MaximoInt = b
End If
End Function
 
Function PoderEvasionEscudo(ByVal userindex As Integer) As Long
 
If UserList(userindex).Invent.EscudoEqpObjIndex = 0 Then
PoderEvasionEscudo = ((UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 30) * _
ModEvasionDeEscudoClase(UserList(userindex).clase)) / 2
Else
PoderEvasionEscudo = ((UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MaxDef + 30) * _
ModEvasionDeEscudoClase(UserList(userindex).clase)) / 2
End If
 
End Function

Function PoderEvasion(ByVal userindex As Integer) As Long
    Dim lTemp As Long
     With UserList(userindex)
       lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
          ModificadorEvasion(.clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(.Stats.ELV - 12, 0)))
    End With
End Function



'Function PoderEvasion(ByVal UserIndex As Integer) As Long
'Dim PoderEvasionTemp As Long

'If UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < 31 Then
'    PoderEvasionTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) * _
'    ModificadorEvasion(UserList(UserIndex).Clase))
'ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < 61 Then
'        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) + _
'        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
'        ModificadorEvasion(UserList(UserIndex).Clase))
'ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < 91 Then
'        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) + _
'        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
'        ModificadorEvasion(UserList(UserIndex).Clase))
'Else
'        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) + _
'        (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
'        ModificadorEvasion(UserList(UserIndex).Clase))
'End If
'PoderEvasion = (PoderEvasionTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))
'
'End Function
'
'
'



Function PoderAtaqueArma(ByVal userindex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userindex).Stats.UserSkills(eSkill.Armas) < 31 Then
    PoderAtaqueTemp = (UserList(userindex).Stats.UserSkills(eSkill.Armas) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Armas) + _
    UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Armas) + _
    (2 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).clase))
Else
   PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Armas) + _
   (3 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(userindex).clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal userindex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(userindex).clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) + _
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(userindex).clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) + _
        (2 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(userindex).clase))
Else
       PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) + _
      (3 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(userindex).clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))

End Function

Function PoderAtaqueWresterling(ByVal userindex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userindex).Stats.UserSkills(eSkill.Wresterling) < 31 Then
    PoderAtaqueTemp = (UserList(userindex).Stats.UserSkills(eSkill.Wresterling) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) < 61 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Wresterling) + _
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(userindex).clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) < 91 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Wresterling) + _
        (2 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(userindex).clase))
Else
       PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Wresterling) + _
       (3 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(userindex).clase))
End If

PoderAtaqueWresterling = (PoderAtaqueTemp + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(userindex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(userindex)
    Else
        PoderAtaque = PoderAtaqueArma(userindex)
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWresterling(userindex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(userindex, Proyectiles)
       Else
            Call SubirSkill(userindex, Armas)
       End If
    Else
        Call SubirSkill(userindex, Wresterling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(userindex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(userindex)

SkillTacticas = UserList(userindex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(userindex).Stats.UserSkills(eSkill.Defensa)

'Esta usando un escudo ???
If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
    If Not NpcImpacto Then
        If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_ESCUDO)
                Call SendData(SendTarget.toindex, userindex, 0, "7")
                Call SubirSkill(userindex, Defensa)
            End If
        End If
    End If
End If
End Function


Public Function CalcularDaño(ByVal userindex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim DañoMaxArma As Long

''sacar esto si no queremos q la matadracos mate el dragon si o si
Dim matodragon As Boolean
matodragon = False


If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
       
        'Usa la mata dragones?
        If UserList(userindex).Invent.WeaponEqpObjIndex = 1053 And Npclist(NpcIndex).NPCtype = DRAGON Then ' Usa la matadragones?
          If UserList(userindex).flags.UserNumQuest = 0 Then
                ModifClase = ModicadorDañoClaseArmas(UserList(userindex).clase)
                DañoArma = RandomNumber(220, 225)
                DañoMaxArma = 350
          Else
                ModifClase = ModicadorDañoClaseArmas(UserList(userindex).clase)
                DañoArma = 1
                DañoMaxArma = 1
          End If
        Else ' daño comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(userindex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(userindex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(userindex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If UserList(userindex).Invent.WeaponEqpObjIndex = 1053 Then
            ModifClase = ModicadorDañoClaseArmas(UserList(userindex).clase)
                DañoArma = 1 ' Si usa la espada matadragones daño es 1
            DañoMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(userindex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(userindex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(userindex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If
Else
    CalcularDaño = CInt(UserList(userindex).Stats.MaxHIT / 5)
    Exit Function
End If


DañoUsuario = RandomNumber(UserList(userindex).Stats.MinHIT, UserList(userindex).Stats.MaxHIT)

''sacar esto si no queremos q la matadracos mate el dragon si o si
If matodragon Then
    CalcularDaño = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
Else
    CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase)
End If

End Function

Public Sub UserDañoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
Dim Daño As Long
Dim GolpeCritico As Byte
GolpeCritico = RandomNumber(1, 5)

If PuedeAtacarNPC(userindex, NpcIndex) = False Then Exit Sub

Daño = CalcularDaño(userindex, NpcIndex)

'esta navegando? si es asi le sumamos el daño del barco
If UserList(userindex).flags.Navegando = 1 Then _
        Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(userindex).Invent.BarcoObjIndex).MaxHIT)

Daño = Daño - Npclist(NpcIndex).Stats.def

If Npclist(NpcIndex).MaestroUser > 0 Then
    Daño = Daño * 1.5
End If

If Daño < 0 Then Daño = 0

If GolpeCritico = 1 Or GolpeCritico = 4 Then
     
     If GranPoder = userindex Then Daño = Daño * 1.8
     
    Call CalcularDarExp(userindex, NpcIndex, Round(Daño * 2, 0))
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbYellow & "°-" & Daño * 2 & "°" & str(Npclist(NpcIndex).Char.CharIndex))
    Call SendData(SendTarget.toindex, userindex, 0, "||138")
    Call SendData(SendTarget.toindex, userindex, 0, "U2" & Round(Daño * 2, 0))
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Round(Daño * 2, 0)
    
    If Npclist(NpcIndex).Stats.MinHP > 0 Then
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(userindex) Then
           Call DoApuñalar(userindex, NpcIndex, 0, Daño * 2)
           Call SubirSkill(userindex, Apuñalar)
        End If
    End If
 
Else
 
     If GranPoder = userindex Then Daño = Daño * 1.8
        Call CalcularDarExp(userindex, NpcIndex, Daño)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbYellow & "°-" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))
        Call SendData(SendTarget.toindex, userindex, 0, "U2" & Daño)
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño

If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apuñalar por la espalda al enemigo
    If PuedeApuñalar(userindex) Then
       Call DoApuñalar(userindex, NpcIndex, 0, Daño)
       Call SubirSkill(userindex, Apuñalar)
    End If
End If

End If


Call CheckPets(NpcIndex, userindex, True)

 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        
        ' Si era un Dragon perdemos la espada matadragones
        If Npclist(NpcIndex).NPCtype = DRAGON Then
            'Si tiene equipada la matadracos se la sacamos
            If UserList(userindex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                Call QuitarObjetos(EspadaMataDragonesIndex, 1, userindex)
            End If
            If Npclist(NpcIndex).Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(userindex).Name & " mató un dragón")
        End If
        
        
        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
        
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If UserList(userindex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex Then Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0
                Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
            End If
        Next j
        
        Call MuereNpc(NpcIndex, userindex)
End If

End Sub


Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal userindex As Integer)

Dim Daño As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
Dim antdaño As Integer, defbarco As Integer
Dim Obj As ObjData



Daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = Daño


If UserList(userindex).flags.Navegando = 1 Then
    Obj = ObjData(UserList(userindex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If


Lugar = RandomNumber(1, 6)


Select Case Lugar
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(userindex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           Daño = Daño - absorbido
           If Daño < 1 Then Daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           Daño = Daño - absorbido
           If Daño < 1 Then Daño = 1
        End If
End Select

If Daño > 149 Then
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & &HFFFF& & "°" & "- " & Daño & "" & "°" & str(UserList(userindex).Char.CharIndex))
Else
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & &HFFFF& & "°" & "- " & Daño & "" & "°" & str(UserList(userindex).Char.CharIndex))
End If

Call SendData(SendTarget.toindex, userindex, 0, "N2" & Lugar & "," & Daño)

If UserList(userindex).flags.Privilegios = PlayerType.User Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - Daño

Call SendUserHP(userindex)

'Muere el usuario
If UserList(userindex).Stats.MinHP <= 0 Then

    Call SendData(SendTarget.toindex, userindex, 0, "6") ' Le informamos que ha muerto ;)
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = ""
        End If
    End If
    
      If userindex = GranPoder Then
        GranPoder = 0
        Call OtorgarGranPoder(0)
        UserList(userindex).flags.GranPoder = 0
        SendUserVariant (userindex)
    End If
    
    Call UserDie(userindex)

End If

End Sub
Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal userindex As Integer, Optional ByVal CheckElementales As Boolean = True)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(j) > 0 Then
       If UserList(userindex).MascotasIndex(j) <> NpcIndex Then
        If CheckElementales And Npclist(UserList(userindex).MascotasIndex(j)).Numero <> ELEMENTALAGUA Then
            If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex
            'Npclist(UserList(UserIndex).MascotasIndex(j)).flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
            Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
        End If
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal userindex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(userindex).MascotasIndex(j))
    End If
Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean

If UserList(userindex).flags.AdminInvisible = 1 Then Exit Function
If UserList(userindex).flags.Privilegios <> PlayerType.User Then Exit Function

If UserList(userindex).GuildIndex > 0 Then
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloN And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloNorte Then Exit Function
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloS And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloSur Then Exit Function
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloE And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloEste Then Exit Function
    If Npclist(NpcIndex).Numero = 620 And UserList(userindex).Pos.Map = MapCastilloO And Guilds(UserList(userindex).GuildIndex).GuildName = CastilloOeste Then Exit Function
End If

' El npc puede atacar ???
If Npclist(NpcIndex).CanAttack = 1 Then
    NpcAtacaUser = True
    Call CheckPets(NpcIndex, userindex, False)

    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = userindex

    If UserList(userindex).flags.AtacadoPorNpc = 0 And _
       UserList(userindex).flags.AtacadoPorUser = 0 Then UserList(userindex).flags.AtacadoPorNpc = NpcIndex
Else
    NpcAtacaUser = False
    Exit Function
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd1)

If NpcImpacto(NpcIndex, userindex) Then
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(userindex).flags.Meditando = False Then
        If UserList(userindex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    End If
    
    Call NpcDaño(NpcIndex, userindex)
    Call SendData(SendTarget.toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
    '¿Puede envenenar?
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(userindex)
Else
    Call SendData(SendTarget.toindex, userindex, 0, "N1")
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbRed & "°¡Fallo!" & "°" & str(UserList(userindex).Char.CharIndex))
End If



'-----Tal vez suba los skills------
Call SubirSkill(userindex, Tacticas)

'Controla el nivel del usuario
Call CheckUserLevel(userindex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long, dif As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)


End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
Dim Daño As Integer
Dim ANpc As npc, DNpc As npc
ANpc = Npclist(Atacante)

Daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - Daño

If Npclist(Victima).Stats.MinHP < 1 Then
        
        If Npclist(Atacante).flags.AttackedBy <> "" Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If
        
        Call FollowAmo(Atacante)
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
End If

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

' El npc puede atacar ???
'If Npclist(Atacante).CanAttack = 1 Then
       'Npclist(Atacante).CanAttack = 0
        'If cambiarMOvimiento Then
        '    Npclist(Victima).TargetNPC = Atacante
        '    Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        'End If
'Else
'    Exit Sub
'End If

If Npclist(Atacante).flags.Snd1 > 0 Then Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).Pos.Map, "TW" & Npclist(Atacante).flags.Snd1)

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & Npclist(Victima).flags.Snd2)
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_IMPACTO2)
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).Pos.Map, "TW" & SND_IMPACTO)
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_IMPACTO)
    End If
    Call NpcDañoNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).Pos.Map, "TW" & SND_SWING)
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_SWING)
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)

If Distancia(UserList(userindex).Pos, Npclist(NpcIndex).Pos) > MAXDISTANCIAARCO Then
   Call SendData(SendTarget.toindex, userindex, 0, "||139")
   Exit Sub
End If

If PuedeAtacarNPC(userindex, NpcIndex) = False Then Exit Sub

Call NpcAtacado(NpcIndex, userindex)

If UserImpactoNpc(userindex, NpcIndex) Then
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    Else
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_IMPACTO2)
    End If
    
    If UserList(userindex).Invent.MunicionEqpObjIndex And ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "FLECHI" & UserList(userindex).Char.CharIndex & "," & Npclist(NpcIndex).Char.CharIndex & "," & ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).GrhIndex)
    End If
    
    Call UserDañoNpc(userindex, NpcIndex)
   
Else
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_SWING)
    Call SendData(SendTarget.toindex, userindex, 0, "U1")
    
    If UserList(userindex).Invent.MunicionEqpObjIndex And ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "FLECHI" & UserList(userindex).Char.CharIndex & "," & Npclist(NpcIndex).Char.CharIndex & "," & ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).GrhIndex & "," & 1)
    End If
    
End If

End Sub

Public Sub UsuarioAtaca(ByVal userindex As Integer)

On Error GoTo Errhandler

'If UserList(UserIndex).flags.PuedeAtacar = 1 Then
    'Quitamos stamina
    If UserList(userindex).Stats.MinSta >= 10 Then
        Call QuitarSta(userindex, RandomNumber(1, 10))
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||17")
        Exit Sub
    End If
    
    'UserList(UserIndex).flags.PuedeAtacar = 0
    
    Dim AttackPos As WorldPos
    AttackPos = UserList(userindex).Pos
    Call HeadtoPos(UserList(userindex).Char.Heading, AttackPos)
    
    'Exit if not legal
    If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_SWING)
        Exit Sub
    End If
    
    Dim index As Integer
    index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex
        
    'Look for user
    If index > 0 Then
        If UserList(userindex).flags.GemaActivada = "Violeta" Then
                UserEnvenena userindex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex
            End If
        Call UsuarioAtacaUsuario(userindex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex)
        Call SendUserData(userindex)
        Call SendUserData(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex)
        Exit Sub
    End If
    
    'Look for NPC
    If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
    
        If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then

            Call UsuarioAtacaNpc(userindex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)
            
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||140")
            Exit Sub
        End If
        
        Exit Sub
    End If
    
        'Está el bot?
        Dim bot_Index   As Byte
       
        bot_Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).BotIndex
       
        If bot_Index <> 0 Then
           'Checkeo que esté invocado.
           If ia_Bot(bot_Index).Invocado Then
              'compruebo que este en mi grupo
              'If ia_Bot(bot_Index).GrupoID = UserList(UserIndex).Group_User.Grupo_ID Then
                 ia_DamageHit bot_Index, userindex
              'End If
           End If
        End If
    
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_SWING)
    Call SendUserData(userindex)

If UserList(userindex).Counters.Trabajando Then _
    UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1
    
If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
    
Errhandler:
    'Call LogError("Error en UsuarioAtaca: " & Err.Description)

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)

Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
If Arma > 0 Then
    proyectil = ObjData(Arma).proyectil = 1
Else
    proyectil = False
End If

'Calculamos el poder de evasion...
'BONIFICADORES - Evasión:
If UserList(VictimaIndex).Bon1 = "Aumenta tu evasion." And UserList(VictimaIndex).Bon2 = "Aumenta tu evasion." Then
UserPoderEvasion = PoderEvasion(VictimaIndex) + 0.04
ElseIf UserList(VictimaIndex).Bon1 = "Aumenta tu evasion." And UserList(VictimaIndex).Bon3 = "Aumenta tu evasion." Then
UserPoderEvasion = PoderEvasion(VictimaIndex) + 0.04
ElseIf UserList(VictimaIndex).Bon1 = "Aumenta tu evasion." Then
UserPoderEvasion = PoderEvasion(VictimaIndex) + 0.02
ElseIf UserList(VictimaIndex).Bon2 = "Aumenta tu evasion." Then
UserPoderEvasion = PoderEvasion(VictimaIndex) + 0.02
ElseIf UserList(VictimaIndex).Bon3 = "Aumenta tu evasion." Then
UserPoderEvasion = PoderEvasion(VictimaIndex) + 0.02
Else
UserPoderEvasion = PoderEvasion(VictimaIndex)
End If
'BONIFICADORES - Evasión:

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then

'BONIFICADORES - Bloquear con Escudos:
If UserList(VictimaIndex).Bon1 = "Aumenta tu posibilidad de bloquear con escudos." And UserList(VictimaIndex).Bon2 = "Aumenta tu posibilidad de bloquear con escudos." Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex) + 0.08
ElseIf UserList(VictimaIndex).Bon1 = "Aumenta tu posibilidad de bloquear con escudos." Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex) + 0.04
ElseIf UserList(VictimaIndex).Bon2 = "Aumenta tu posibilidad de bloquear con escudos." Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex) + 0.04
ElseIf UserList(VictimaIndex).Bon3 = "Aumenta tu posibilidad de bloquear con escudos." Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex) + 0.04
End If
'BONIFICADORES - Bloquear con Escudos:


   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
'BONIFICADORES - Ataque con flechas:
     If UserList(AtacanteIndex).Bon2 = "Aumenta tu posibilidad de pegar con flechas." Or UserList(AtacanteIndex).Bon3 = "Aumenta tu posibilidad de pegar con flechas." Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex) + 0.04
     Else
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
     End If
'BONIFICADORES - Ataque con flechas:
    Else
'BONIFICADORES - Ataque con armas:
    If UserList(AtacanteIndex).Bon1 = "Aumenta tu posibilidad de pegar con armas." And UserList(AtacanteIndex).Bon2 = "Aumenta tu posibilidad de pegar con armas." Then
        PoderAtaque = PoderAtaqueArma(AtacanteIndex) + 0.08
    ElseIf UserList(AtacanteIndex).Bon1 = "Aumenta tu posibilidad de pegar con armas." Then
        PoderAtaque = PoderAtaqueArma(AtacanteIndex) + 0.04
    ElseIf UserList(AtacanteIndex).Bon2 = "Aumenta tu posibilidad de pegar con armas." Then
        PoderAtaque = PoderAtaqueArma(AtacanteIndex) + 0.04
    Else
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
    End If
'BONIFICADORES - Ataque con armas:
    End If
    
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWresterling(AtacanteIndex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If

UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 And UCase$(UserList(VictimaIndex).clase) <> "MAGO" Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_ESCUDO)
              Call SendData(SendTarget.toindex, AtacanteIndex, 0, "8")
              Call SendData(SendTarget.toindex, VictimaIndex, 0, "7")
              Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
    
If UsuarioImpacto Then
   If Arma > 0 Then
           If Not proyectil Then
                  Call SubirSkill(AtacanteIndex, Armas)
           Else
                  Call SubirSkill(AtacanteIndex, Proyectiles)
           End If
   Else
        Call SubirSkill(AtacanteIndex, Wresterling)
   End If
   
            'Arco de 4ta jerarquia paraliza
            If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex = 1219 Then
                Dim ProbParalizar As Byte
                ProbParalizar = RandomNumber(1, 12)
                
                If ProbParalizar = 7 Then
                    If UserList(VictimaIndex).flags.Paralizado = 0 Then
                        Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "TW" & Hechizos(9).WAV)
                        Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & Hechizos(9).FXgrh & "," & Hechizos(9).loops)
                       
                       
                        UserList(VictimaIndex).flags.Paralizado = 1
                        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
                        Call SendData(SendTarget.toindex, VictimaIndex, 0, "PARADOK")
                        Call SendData(SendTarget.toindex, VictimaIndex, 0, "PU" & UserList(VictimaIndex).Pos.X & "," & UserList(VictimaIndex).Pos.Y)
                        Call SendData(SendTarget.toindex, VictimaIndex, 0, "||141@" & UserList(AtacanteIndex).Name)
                    End If
                End If
            End If
   
End If

'SE APUÑALA SIEMPRE.
If UsuarioImpacto = False Then
    If UserList(AtacanteIndex).Char.Heading = UserList(VictimaIndex).Char.Heading And UCase$(UserList(AtacanteIndex).clase) = "ASESINO" Then
        UsuarioImpacto = True
    End If
End If


End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

If UserList(AtacanteIndex).flags.EspectadorArena1 = 1 Or UserList(AtacanteIndex).flags.EspectadorArena2 = 1 Or UserList(AtacanteIndex).flags.EspectadorArena3 = 1 Or UserList(AtacanteIndex).flags.EspectadorArena4 = 1 Then
        Call SendData(SendTarget.toindex, AtacanteIndex, 0, "||142")
    Exit Sub
End If

If UserList(VictimaIndex).flags.EspectadorArena1 = 1 Or UserList(VictimaIndex).flags.EspectadorArena2 = 1 Or UserList(VictimaIndex).flags.EspectadorArena3 = 1 Or UserList(VictimaIndex).flags.EspectadorArena4 = 1 Then
        Call SendData(SendTarget.toindex, AtacanteIndex, 0, "||143")
    Exit Sub
End If

If Distancia(UserList(AtacanteIndex).Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
   Call SendData(SendTarget.toindex, AtacanteIndex, 0, "||139")
   Exit Sub
End If


Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(VictimaIndex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
Else
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_SWING)
    Call SendData(SendTarget.toindex, AtacanteIndex, 0, "U1")
    Call SendData(SendTarget.toindex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).Name)
    
    If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil Then
        Call SendData(SendTarget.ToMap, 0, UserList(AtacanteIndex).Pos.Map, "FLECHI" & UserList(AtacanteIndex).Char.CharIndex & "," & UserList(VictimaIndex).Char.CharIndex & "," & ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).GrhIndex & "," & 1)
    End If
    
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "N|" & vbRed & "°¡Fallo!" & "°" & str(UserList(VictimaIndex).Char.CharIndex))
End If

If UCase$(UserList(AtacanteIndex).clase) = "LADRON" Then Call Desarmar(AtacanteIndex, VictimaIndex)

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
On Error Resume Next

Dim Daño As Long, antdaño As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer

Dim Obj As ObjData

Daño = CalcularDaño(AtacanteIndex)
antdaño = Daño

Call UserEnvenena(AtacanteIndex, VictimaIndex)

If UserList(AtacanteIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     Daño = Daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(VictimaIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim Resist As Byte
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    Resist = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo
End If

Lugar = RandomNumber(1, 6)

Select Case Lugar
  
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco - Resist
           Daño = Daño - absorbido
           Daño = Daño + 20
           If Daño < 0 Then Daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco - Resist
           Daño = Daño - absorbido
           If Daño < 0 Then Daño = 1
        End If
End Select

If UserList(VictimaIndex).flags.GemaActivada = "Verde" Then
Daño = Daño - (Daño * 10 / 100 + RandomNumber(0, 4))
End If

If UserList(VictimaIndex).flags.IntervaloBurbu > 1 Then
Daño = Daño - UserList(VictimaIndex).flags.DefensaBurbu
End If

'Bonificador - BARDO:
If UserList(AtacanteIndex).Bon3 = "Aumenta levemente tu daño con armas." Then
    Daño = Daño + (Daño * 4 / 100)
End If
'Bonificador - BARDO:

'BALANCEO

    'Subimos/bajamos el ataque fisico del atacante
    Daño = Round(Daño + (Daño * ModificarAtaqueFisico(UserList(AtacanteIndex).clase) / 100))
    
    '/: Subimos/bajamos la defensa fisica del que recibe el ataque
    Daño = Round(Daño - (Daño * ModificarDefensaFisica(UserList(VictimaIndex).clase) / 100))
        
'BALANCEO

If AtacanteIndex = GranPoder Then Daño = Daño * 1.5


'SI tiene una manopla tratamos de inmear al enemigo.
    If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Inmoviliza = 1 Then
        Dim RandomManopla As Byte
        RandomManopla = RandomNumber(1, 100)
        If Lugar = PartesCuerpo.bCabeza Then RandomManopla = 1
        If RandomManopla <= 12 Then
            Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "TW" & 16)
            Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & 8 & "," & 0)
            
            UserList(VictimaIndex).Counters.InmoManopla = 2
            UserList(VictimaIndex).flags.Paralizado = 1
            UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.toindex, VictimaIndex, 0, "PARADOK")
            Call SendData(SendTarget.toindex, VictimaIndex, 0, "PU" & UserList(VictimaIndex).Pos.X & "," & UserList(VictimaIndex).Pos.Y)
            Call SendData(SendTarget.toindex, VictimaIndex, 0, "||896@" & UserList(AtacanteIndex).Name)
        End If
    End If
    

If Daño < 0 Then Daño = 0

Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "N|" & vbYellow & "°" & "- " & Daño & "" & "°" & str(UserList(VictimaIndex).Char.CharIndex))
Call SendData(SendTarget.toindex, AtacanteIndex, 0, "N5" & Lugar & "," & Daño & "," & UserList(VictimaIndex).Name)
Call SendData(SendTarget.toindex, VictimaIndex, 0, "N4" & Lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)

UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño

If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
        'Si usa un arma quizas suba "Combate con armas"
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call SubirSkill(AtacanteIndex, Armas)
        Else
        'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, Wresterling)
        End If
        
    If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil Then
     Call SendData(SendTarget.ToMap, 0, UserList(AtacanteIndex).Pos.Map, "FLECHI" & UserList(AtacanteIndex).Char.CharIndex & "," & UserList(VictimaIndex).Char.CharIndex & "," & ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).GrhIndex)
    End If
    
    
        Call SubirSkill(AtacanteIndex, Tacticas)
        
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, Daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
        End If
End If


If UserList(VictimaIndex).Stats.MinHP <= 0 Then
    
    Call ContarMuerte(VictimaIndex, AtacanteIndex)
    
    ' Para que las mascotas no sigan intentando luchar y
    ' comiencen a seguir al amo
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
    Next j
    
    Call ActStats(VictimaIndex, AtacanteIndex)
    Call UserDie(VictimaIndex)
Else
    'Está vivo - Actualizamos el HP
    Call SendUserHP(VictimaIndex)
End If

'Controla el nivel del usuario
Call CheckUserLevel(AtacanteIndex)

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal Victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
        If Npclist(UserList(Maestro).MascotasIndex(iCount)).Numero = ELEMENTALFUEGO Then
                Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(Victim).Name
                Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
                Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    End If
Next iCount

End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

Dim T As eTrigger6

If UserList(VictimIndex).flags.Muerto = 1 Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||154"
    PuedeAtacar = False
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= PlayerType.UserSupport Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then SendData SendTarget.toindex, AttackerIndex, 0, "||155"
    PuedeAtacar = False
    Exit Function
End If

If (UserList(AttackerIndex).flags.Invisible = 1 Or UserList(AttackerIndex).flags.Oculto = 1) And UserList(AttackerIndex).flags.Privilegios <= PlayerType.UserSupport Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||156")
    PuedeAtacar = False
Exit Function
End If

If UserList(AttackerIndex).flags.EnAram Then
    If (UserList(AttackerIndex).flags.AramAzul And UserList(VictimIndex).flags.AramAzul) Or (UserList(AttackerIndex).flags.AramRojo And UserList(VictimIndex).flags.AramRojo) Then
        Call SendData(SendTarget.toindex, AttackerIndex, 0, "||903")
    End If
End If

If UserList(VictimIndex).Pos.X > UserList(AttackerIndex).Pos.X And (UserList(VictimIndex).Pos.X - UserList(AttackerIndex).Pos.X) > 7 Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||158")
    PuedeAtacar = False
    Exit Function
ElseIf UserList(VictimIndex).Pos.X < UserList(AttackerIndex).Pos.X And (UserList(AttackerIndex).Pos.X - UserList(VictimIndex).Pos.X) > 7 Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||158")
    PuedeAtacar = False
    Exit Function
ElseIf UserList(VictimIndex).Pos.Y > UserList(AttackerIndex).Pos.Y And (UserList(VictimIndex).Pos.Y - UserList(AttackerIndex).Pos.Y) > 7 Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||158")
    PuedeAtacar = False
    Exit Function
ElseIf UserList(VictimIndex).Pos.Y < UserList(AttackerIndex).Pos.Y And (UserList(AttackerIndex).Pos.Y - UserList(VictimIndex).Pos.Y) > 7 Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||158")
    PuedeAtacar = False
    Exit Function
End If

Dim klan As String
    If UserList(VictimIndex).GuildIndex > 0 And UserList(AttackerIndex).GuildIndex > 0 Then
              klan = Guilds(UserList(AttackerIndex).GuildIndex).GuildName
             
             If UserList(AttackerIndex).flags.SeguroClan = True Then
              If UCase$(Guilds(UserList(VictimIndex).GuildIndex).GuildName) = UCase$(Guilds(UserList(AttackerIndex).GuildIndex).GuildName) Then
                  Call SendData(SendTarget.toindex, AttackerIndex, 0, "||159")
                  Exit Function
              End If
            End If
              
              If UCase$(Guilds(UserList(VictimIndex).GuildIndex).GuildName) = UCase$(Guilds(UserList(AttackerIndex).GuildIndex).GuildName) And UserList(AttackerIndex).Pos.Map = 8 Then
                  Call SendData(SendTarget.toindex, AttackerIndex, 0, "||159")
                  Exit Function
              End If
    End If
     
            If UserList(AttackerIndex).flags.PartyIndex <> 0 Then
                If UserList(VictimIndex).flags.PartyIndex = UserList(AttackerIndex).flags.PartyIndex Then
                        Call SendData(SendTarget.toindex, AttackerIndex, 0, "||160")
                    Exit Function
                End If
            End If

            If UserList(AttackerIndex).Pos.Map = 118 And CuentaRegresiva > 0 Then
                Call SendData(SendTarget.toindex, AttackerIndex, 0, "||161")
                Exit Function
            End If

            If (UserList(AttackerIndex).Pos.Map = 100 Or UserList(AttackerIndex).Pos.Map = 99) And Hay_Torneo = True Then
             If UCase$(TModalidad) = "DM" And TiroCuentaDM = False Then
                 Call SendData(SendTarget.toindex, AttackerIndex, 0, "||162")
               Exit Function
             End If
            
             If UsuarioPelea(1) <> AttackerIndex And UsuarioPelea(2) <> AttackerIndex And UsuarioPelea(3) <> AttackerIndex And UsuarioPelea(4) <> AttackerIndex And UsuarioPelea(5) <> AttackerIndex And UsuarioPelea(6) <> AttackerIndex And UsuarioPelea(7) <> AttackerIndex And UsuarioPelea(8) <> AttackerIndex Then
              If TModalidad = "1" Or TModalidad = "2" Or TModalidad = "3" Or TModalidad = "4" Then
                  Call SendData(SendTarget.toindex, AttackerIndex, 0, "||162")
                 Exit Function
              End If
             End If
            End If

T = TriggerZonaPelea(AttackerIndex, VictimIndex)

If T = TRIGGER6_PERMITE Then
    PuedeAtacar = True
    Exit Function
ElseIf T = TRIGGER6_PROHIBE Then
    PuedeAtacar = False
    Exit Function
End If


If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||163")
    PuedeAtacar = False
    Exit Function
End If

If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
    MapData(UserList(AttackerIndex).Pos.Map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||164")
    PuedeAtacar = False
    Exit Function
End If

If UserList(AttackerIndex).flags.Privilegios = PlayerType.UserSupport Then
    PuedeAtacar = False
    Exit Function
End If

'Se asegura que la victima no es un GM
If UserList(VictimIndex).flags.Privilegios >= PlayerType.UserSupport Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||155"
    PuedeAtacar = False
    Exit Function
End If

If UserList(AttackerIndex).flags.Muerto = 1 Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||3"
    PuedeAtacar = False
    Exit Function
End If

If EsAlianza(VictimIndex) And EsAlianza(AttackerIndex) And UserList(AttackerIndex).Pos.Map <> MapCastilloS And UserList(AttackerIndex).Pos.Map <> MapCastilloN And UserList(AttackerIndex).Pos.Map <> MapCastilloE And UserList(AttackerIndex).Pos.Map <> MapCastilloO And UserList(AttackerIndex).Pos.Map <> 109 And UserList(AttackerIndex).Pos.Map <> 108 And UserList(AttackerIndex).Pos.Map <> 106 And UserList(AttackerIndex).Pos.Map <> 71 And UserList(AttackerIndex).Pos.Map <> 100 And UserList(AttackerIndex).Pos.Map <> 107 And UserList(AttackerIndex).Pos.Map <> 109 And UserList(AttackerIndex).Pos.Map <> 110 And UserList(AttackerIndex).Pos.Map <> 106 And UserList(AttackerIndex).Pos.Map <> 81 Then
        Call SendData(SendTarget.toindex, AttackerIndex, 0, "||165")
    Exit Function
End If
       
If EsHorda(VictimIndex) And EsHorda(AttackerIndex) And UserList(AttackerIndex).Pos.Map <> MapCastilloS And UserList(AttackerIndex).Pos.Map <> MapCastilloN And UserList(AttackerIndex).Pos.Map <> MapCastilloE And UserList(AttackerIndex).Pos.Map <> MapCastilloO And UserList(AttackerIndex).Pos.Map <> 109 And UserList(AttackerIndex).Pos.Map <> 108 And UserList(AttackerIndex).Pos.Map <> 106 And UserList(AttackerIndex).Pos.Map <> 71 And UserList(AttackerIndex).Pos.Map <> 100 And UserList(AttackerIndex).Pos.Map <> 107 And UserList(AttackerIndex).Pos.Map <> 109 And UserList(AttackerIndex).Pos.Map <> 110 And UserList(AttackerIndex).Pos.Map <> 106 And UserList(AttackerIndex).Pos.Map <> 81 Then
        Call SendData(SendTarget.toindex, AttackerIndex, 0, "||166")
    Exit Function
End If
   

PuedeAtacar = True

End Function


Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean

If UserList(AttackerIndex).flags.Muerto = 1 Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||3"
    PuedeAtacarNPC = False
    Exit Function
End If

If UserList(AttackerIndex).flags.Privilegios > PlayerType.User And UserList(AttackerIndex).flags.Privilegios <= PlayerType.TournamentManager Then
    PuedeAtacarNPC = False
    Exit Function
End If

If (Npclist(NpcIndex).Numero = 617 Or Npclist(NpcIndex).Numero = 948) And EsHorda(AttackerIndex) Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||167")
    PuedeAtacarNPC = False
  Exit Function
End If

If (Npclist(NpcIndex).Numero = 618 Or Npclist(NpcIndex).Numero = 947) And EsAlianza(AttackerIndex) Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||167")
    PuedeAtacarNPC = False
  Exit Function
End If

If (Npclist(NpcIndex).Numero = 963) And UserList(AttackerIndex).flags.AramRojo Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||897")
    PuedeAtacarNPC = False
  Exit Function
End If

If (Npclist(NpcIndex).Numero = 964) And UserList(AttackerIndex).flags.AramAzul Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||897")
    PuedeAtacarNPC = False
  Exit Function
End If

If Npclist(NpcIndex).Pos.Map = 95 And Npclist(NpcIndex).Numero = 937 And GuardiasRey <= 3 Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||168")
    PuedeAtacarNPC = False
  Exit Function
End If

If Npclist(NpcIndex).NPCtype = ReyCastillo Or Npclist(NpcIndex).Numero = 615 Then
    If (Npclist(NpcIndex).Pos.Map = MapCastilloN Or Npclist(NpcIndex).Pos.Map = MapCastilloS Or Npclist(NpcIndex).Pos.Map = MapCastilloE Or Npclist(NpcIndex).Pos.Map = MapCastilloO Or Npclist(NpcIndex).Pos.Map = 81) Then
            Dim castiact As String
            If Npclist(NpcIndex).Pos.Map = MapCastilloN Then castiact = CastilloNorte
            If Npclist(NpcIndex).Pos.Map = MapCastilloS Then castiact = CastilloSur
            If Npclist(NpcIndex).Pos.Map = MapCastilloE Then castiact = CastilloEste
            If Npclist(NpcIndex).Pos.Map = MapCastilloO Then castiact = CastilloOeste
            If Npclist(NpcIndex).Pos.Map = 81 Then castiact = Fortaleza
            
                If Not UserList(AttackerIndex).GuildIndex <> 0 Then
                    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||120")
                    PuedeAtacarNPC = False
                 Exit Function
                End If
            
                If UCase$(Guilds(UserList(AttackerIndex).GuildIndex).GuildName) = UCase$(castiact) Then
                    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||169")
                    PuedeAtacarNPC = False
                    Exit Function
                End If
                
                
                If UserList(AttackerIndex).Pos.Map = 81 Then
                    If Guilds(UserList(AttackerIndex).GuildIndex).GuildName <> CastilloNorte Or Guilds(UserList(AttackerIndex).GuildIndex).GuildName <> CastilloSur Or Guilds(UserList(AttackerIndex).GuildIndex).GuildName <> CastilloEste Or Guilds(UserList(AttackerIndex).GuildIndex).GuildName <> CastilloOeste Then
                      Call SendData(SendTarget.toindex, AttackerIndex, 0, "||125")
                      PuedeAtacarNPC = False
                      Exit Function
                     End If
                End If
    End If
End If


PuedeAtacarNPC = True

End Function


'[KEVIN]
'
'[Alejo]
'Modifique un poco el sistema de exp por golpe, ahora
'son 2/3 de la exp mientras esta vivo, el resto se
'obtiene al matarlo.
'Ahora además
Sub CalcularDarExp(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)

If UserList(userindex).Stats.ELV >= 60 Then Exit Sub

Dim ExpSinMorir As Long
Dim ExpaDar As Long
Dim TotalNpcVida As Long
Dim YeguitaGorda As Long

If ElDaño <= 0 Then ElDaño = 0
TotalNpcVida = Npclist(NpcIndex).Stats.MaxHP

If ElDaño > Npclist(NpcIndex).Stats.MinHP Then ElDaño = Npclist(NpcIndex).Stats.MinHP


ExpaDar = ((Npclist(NpcIndex).GiveEXP / TotalNpcVida) * ElDaño) * MultiplicadorExp
If ExpaDar <= 0 Then Exit Sub

If ExpaDar > 0 Then
    If UserList(userindex).flags.GemaActivada = "Plateada" Then
        ExpaDar = val(ExpaDar) + (val(ExpaDar) / 2)
    End If
    
    'Oscuridad
    If UserList(userindex).Invent.ArmourEqpObjIndex = 1051 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1052 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1455 Or UserList(userindex).Invent.ArmourEqpObjIndex = 1496 Then
        ExpaDar = val(ExpaDar) + (val(ExpaDar) / 4)
    End If
    

    If UserList(userindex).flags.PartyIndex > 0 Then
        'DAMOS EXP A LA PARTY
          Dim losusersdelaparty As Long
          For losusersdelaparty = 1 To LastUser
            If (UserList(losusersdelaparty).Stats.ELV < 50) And (losusersdelaparty <> userindex) Then
              If (UserList(losusersdelaparty).flags.PartyIndex = UserList(userindex).flags.PartyIndex) And (UserList(losusersdelaparty).Pos.Map = UserList(userindex).Pos.Map) Then
                UserList(losusersdelaparty).Stats.Exp = UserList(losusersdelaparty).Stats.Exp + (val(ExpaDar) * 10 / 100)
               If UserList(losusersdelaparty).Stats.Exp > MAXEXP Then _
                   UserList(losusersdelaparty).Stats.Exp = MAXEXP
                    Call SendData(SendTarget.toindex, losusersdelaparty, 0, "||170@" & PonerPuntos(val(ExpaDar)))
                    SendUserEXP (losusersdelaparty)
                Call CheckUserLevel(losusersdelaparty)
              End If
             End If
          Next losusersdelaparty
      'DAMOS EXP A LA PARTY
    End If
    
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + val(ExpaDar)
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
        Call SendData(SendTarget.toindex, userindex, 0, "||170@" & PonerPuntos(val(ExpaDar)))
    
    Call CheckUserLevel(userindex)
    Call SendUserEXP(userindex)
End If

'[/KEVIN]
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

If Origen > 0 And Destino > 0 And Origen <= UBound(UserList) And Destino <= UBound(UserList) Then
    If MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger = eTrigger.ZONAPELEA Or _
        MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger = eTrigger.ZONAPELEA Then
        If (MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger) Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If
Else
    TriggerZonaPelea = TRIGGER6_AUSENTE
End If

End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
    End If
    
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            
            If num < 60 Then
                UserList(VictimaIndex).flags.Envenenado = 1
                Call SendData(SendTarget.toindex, VictimaIndex, 0, "||171@" & UserList(AtacanteIndex).Name)
                Call SendData(SendTarget.toindex, AtacanteIndex, 0, "||172@" & UserList(VictimaIndex).Name)
            End If
        End If
    End If
End If

End Sub

