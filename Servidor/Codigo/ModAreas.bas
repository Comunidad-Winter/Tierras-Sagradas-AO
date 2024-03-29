Attribute VB_Name = "ModAreas"
'Argentum Online 0.11.20
'Copyright (C) 2002 M�rquez Pablo Ignacio, Jonatan Ezequiel Salguero
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

' Modulo de envio por areas compatible con la versi�n 9.10.x ... By DuNga

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    MinX As Integer '-!!!
    MinY As Integer '-!!!
    
    AreaID As Long
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type

Public Const USER_NUEVO As Byte = 255

'Cuidado:
' ���LAS AREAS EST�N HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte

Public AreasRecive(12) As Integer
'Private AreasEnvia(12) As Integer

Public ConnGroups() As ConnGroup

Public Sub InitAreas()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim loopX As Long
    Dim CurArea As Byte

' Setup areas...
    For loopC = 0 To 11
        AreasRecive(loopC) = (2 ^ loopC) Or IIf(loopC <> 0, 2 ^ (loopC - 1), 0) Or IIf(loopC <> 11, 2 ^ (loopC + 1), 0)
'        AreasEnvia(LoopC) = 2 ^ (LoopC + 1)
    Next loopC
    
    For loopC = 1 To 100
        PosToArea(loopC) = loopC \ 9
    Next loopC
    
    For loopC = 1 To 100
        For loopX = 1 To 100
            'Usamos 121 IDs de area para saber si pasasamos de area "m�s r�pido"
            AreasInfo(loopC, loopX) = (loopC \ 9 + 1) * (loopX \ 9 + 1)
        Next loopX
    Next loopC

'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For loopC = 1 To NumMaps
        ConnGroups(loopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour))
        
        If ConnGroups(loopC).OptValue = 0 Then ConnGroups(loopC).OptValue = 1
        ReDim ConnGroups(loopC).UserEntrys(1 To ConnGroups(loopC).OptValue) As Long
    Next loopC
End Sub

Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la funci�n de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim loopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
        
        For loopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(loopC).OptValue) \ 2))
            
            ConnGroups(loopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, tCurDay & "-" & tCurHour))
            If ConnGroups(loopC).OptValue = 0 Then ConnGroups(loopC).OptValue = 1
            If ConnGroups(loopC).OptValue >= MapInfo(loopC).NumUsers Then ReDim Preserve ConnGroups(loopC).UserEntrys(1 To ConnGroups(loopC).OptValue) As Long
        Next loopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal userindex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la funci�n clave del sistema de areas... Es llamada al mover un user
'**************************************************************
    If UserList(userindex).AreasInfo.AreaID = AreasInfo(UserList(userindex).Pos.X, UserList(userindex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long, Map As Long
    
    With UserList(userindex)
        
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        
        Map = UserList(userindex).Pos.Map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call SendData(SendTarget.toindex, userindex, 0, "CA" & Chr$(.Pos.X) & Chr$(.Pos.Y))
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                If MapData(Map, X, Y).userindex Then
                    TempInt = MapData(Map, X, Y).userindex
                    
                    If userindex <> TempInt Then
                        Call MakeUserChar(SendTarget.toindex, userindex, 0, CInt(TempInt), Map, X, Y)
                        Call MakeUserChar(SendTarget.toindex, CInt(TempInt), 0, userindex, .Pos.Map, .Pos.X, .Pos.Y)
                        
                        'Si el user estaba invisible le avisamos al nuevo cliente de eso

                            If UserList(TempInt).flags.Invisible Or UserList(TempInt).flags.Oculto Then
                                 Call EnviarDatosASlot(userindex, "NOVER" & UserList(TempInt).Char.CharIndex & ",1" & ENDC)
                            End If
                            
                            If UserList(userindex).flags.Invisible Or UserList(userindex).flags.Oculto Then
                                 Call EnviarDatosASlot(TempInt, "NOVER" & UserList(userindex).Char.CharIndex & ",1" & ENDC)
                            End If
                            
                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(SendTarget.toindex, userindex, 0, userindex, Map, X, Y)
                    End If
                
                End If
                
                ' << Bots >>
                Dim botI As Integer
               
                botI = MapData(Map, X, Y).BotIndex
               
                If (botI <> 0) Then
                    If (ia_Bot(botI).Invocado = True) Then
                        Call Mod_BOTS.ia_EnviarChar(userindex, botI)
                    End If
                End If
                
                '<<< Npc >>>
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                 End If
                 
                 'PARTICULAS FER###
                 If MapData(Map, X, Y).particle_group_index Then
                    Call SendData(SendTarget.toindex, userindex, 0, "PCF" & MapData(Map, X, Y).particle_group_index & "," & X & "," & Y & "," & 0)
                 End If
                 
                 If MapData(Map, X, Y).range_light Then
                    Call SendData(SendTarget.toindex, userindex, 0, "PCL" & X & "," & Y & "," & MapData(Map, X, Y).range_light & "," & MapData(Map, X, Y).rgb_light(1) & "," & MapData(Map, X, Y).rgb_light(2) & "," & MapData(Map, X, Y).rgb_light(3))
                 End If
                 
                '<<< Item >>>
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    TempInt = MapData(Map, X, Y).OBJInfo.ObjIndex
                    'If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        'If (TempInt = 378) Then
                       '     Call SendData(SendTarget.toindex, userindex, 0, "PCF" & 101 & "," & X & "," & Y & "-1")
                        'Else
                            Call SendData(SendTarget.toindex, userindex, 0, "HO" & ObjData(TempInt).GrhIndex & "," & X & "," & Y)
                        'End If
                        
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                            
                          If ObjData(TempInt).PuertaDoble = 1 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 2, Y, MapData(Map, X + 2, Y).Blocked)
                          ElseIf ObjData(TempInt).Porton = 1 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 2, Y, MapData(Map, X + 2, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X - 2, Y, MapData(Map, X - 2, Y).Blocked)
                          ElseIf ObjData(TempInt).RejaForta = 1 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 2, Y, MapData(Map, X + 2, Y).Blocked)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X - 2, Y, MapData(Map, X - 2, Y).Blocked)
                          End If
                          
                          If TempInt = 1472 Or TempInt = 1470 Then
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X, Y, 0)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X - 1, Y, 0)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 1, Y, 0)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X + 2, Y, 0)
                            Call Bloquear(SendTarget.toindex, userindex, 0, CInt(Map), X - 2, Y, 0)
                          End If
                            
                        End If
                    End If
                'End If
            
            Next Y
        Next X
            
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' Se llama cuando se mueve un Npc
'**************************************************************
On Error Resume Next
    
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long
    
    With Npclist(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100

        
        'Actualizamos!!!
        If MapInfo(.Pos.Map).NumUsers <> 0 Then
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(.Pos.Map, X, Y).userindex Then _
                        Call MakeNPCChar(SendTarget.toindex, MapData(.Pos.Map, X, Y).userindex, 0, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
                Next Y
            Next X
        End If
            
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub QuitarUser(ByVal userindex As Integer, ByVal Map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    Dim loopC As Long
    
    'Saco del viejo mapa
    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
    TempVal = ConnGroups(Map).CountEntrys
    
    For loopC = 1 To TempVal + 1
        If ConnGroups(Map).UserEntrys(loopC) = userindex Then Exit For
    Next loopC
    
    For loopC = loopC To TempVal
        ConnGroups(Map).UserEntrys(loopC) = ConnGroups(Map).UserEntrys(loopC + 1)
    Next loopC
    
    If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
    End If
End Sub

Public Sub AgregarUser(ByVal userindex As Integer, ByVal Map As Integer, Optional ByVal EsNuevo As Boolean = True)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    
    If EsNuevo Then
        If Not MapaValido(Map) Then Exit Sub
        'Update map and connection groups data
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
        TempVal = ConnGroups(Map).CountEntrys
        
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
        End If
        
        ConnGroups(Map).UserEntrys(TempVal) = userindex
    End If

    'Update user
    UserList(userindex).AreasInfo.AreaID = 0
    
    UserList(userindex).AreasInfo.AreaPerteneceX = 0
    UserList(userindex).AreasInfo.AreaPerteneceY = 0
    UserList(userindex).AreasInfo.AreaReciveX = 0
    UserList(userindex).AreasInfo.AreaReciveY = 0
End Sub

Public Sub ArgegarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Npclist(NpcIndex).AreasInfo.AreaID = 0
    
    Npclist(NpcIndex).AreasInfo.AreaPerteneceX = 0
    Npclist(NpcIndex).AreasInfo.AreaPerteneceY = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveX = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveY = 0
End Sub

Public Sub SendToUserArea(ByVal userindex As Integer, ByVal sdData As String, Optional Encriptar As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim TempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(userindex).Pos.Map
    AreaX = UserList(userindex).AreasInfo.AreaPerteneceX
    AreaY = UserList(userindex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then Exit Sub
    sdData = sdData & ENDC
    
    For loopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopC)
        
        If UserList(TempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(TempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(TempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
        End If
    Next loopC
End Sub
Public Sub SendToUserAreaButindexBOT(ByVal userindex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' ESTA SOLO SE USA PARA ENVIAR MPs asi que se puede encriptar desde aca :)
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
            
    sdData = sdData & ENDC

    
    Map = ia_Bot(userindex).Pos.Map
    Call CalcularArea(userindex)
    AreaX = ia_Bot(userindex).AreasInfo.AreaPerteneceX
    AreaY = ia_Bot(userindex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub

    For loopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopC)
            
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                Call EnviarDatosASlot(TempIndex, sdData)
            End If
        End If
    Next loopC
End Sub
Public Sub SendToUserAreaButindex(ByVal userindex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' ESTA SOLO SE USA PARA ENVIAR MPs asi que se puede encriptar desde aca :)
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    sdData = sdData & ENDC
    
    Map = UserList(userindex).Pos.Map
    AreaX = UserList(userindex).AreasInfo.AreaPerteneceX
    AreaY = UserList(userindex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then Exit Sub

    For loopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopC)
            
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If TempIndex <> userindex Then
                    If UserList(TempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(TempIndex, sdData)
                    End If
                End If
            End If
        End If
    Next loopC
End Sub

Public Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = Npclist(NpcIndex).Pos.Map
    AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    
    sdData = sdData & ENDC
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopC)
        
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
        End If
    Next loopC
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
    
    AreaX = 2 ^ (AreaX \ 9)
    AreaY = 2 ^ (AreaY \ 9)
    
    sdData = sdData & ENDC
    
    If Not MapaValido(Map) Then Exit Sub

    For loopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopC)
            
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
        End If
    Next loopC
End Sub
