Attribute VB_Name = "modQuests"
Option Explicit
Public QuestsList() As tQuests
Public Type tQuests
    Name As String
    Tipo As Byte
    Usuarios As Integer
    Puntos As Integer
    Oro As Long
    NivelMinimo As Byte
    NPCs As Byte
    NumNPC() As Integer
    CantNPC() As Integer
    
    NumOBJ As Integer
    CantOBJ As Integer
End Type
Public Sub CargarQuests()
        Dim p As Integer, LoopC As Integer, LoopD
        p = val(GetVar(App.Path & "\Dat\QUESTS.dat", "INIT", "Num"))
   
        ReDim QuestsList(p) As tQuests
       
           
        For LoopC = 1 To p
            QuestsList(LoopC).Name = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Nombre")
            QuestsList(LoopC).Tipo = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Tipo")
            QuestsList(LoopC).Puntos = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Puntos")
            QuestsList(LoopC).Oro = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Oro")
            
            QuestsList(LoopC).NumOBJ = ReadField(1, GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "OBJ"), Asc("-"))
            QuestsList(LoopC).CantOBJ = ReadField(2, GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "OBJ"), Asc("-"))
            
            If QuestsList(LoopC).Tipo = 1 Then
                QuestsList(LoopC).NPCs = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "NPCs")
                
                    'Seteamos el numero y la cantidad de npcs..
                    ReDim QuestsList(LoopC).NumNPC(QuestsList(LoopC).NPCs) As Integer
                    ReDim QuestsList(LoopC).CantNPC(QuestsList(LoopC).NPCs) As Integer
                    
                    For LoopD = 1 To QuestsList(LoopC).NPCs
                        QuestsList(LoopC).NumNPC(LoopD) = ReadField(1, GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Npc" & LoopD), Asc("-"))
                        QuestsList(LoopC).CantNPC(LoopD) = ReadField(2, GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Npc" & LoopD), Asc("-"))
                    Next LoopD
            ElseIf QuestsList(LoopC).Tipo = 2 Then
                QuestsList(LoopC).Usuarios = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "Usuarios")
            End If
            
            QuestsList(LoopC).NivelMinimo = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & LoopC, "NivelMinimo")
                
        Next LoopC
End Sub
Public Sub RestarNPC(ByVal userindex As Integer, ByVal KillNPC As Integer)

    Dim NroQuest As Byte, i As Long, CompletoQuest As Boolean
    NroQuest = UserList(userindex).flags.UserNumQuest

    If QuestsList(NroQuest).Tipo = 1 Then
        
        CompletoQuest = True
        For i = 1 To QuestsList(NroQuest).NPCs
            If KillNPC = QuestsList(NroQuest).NumNPC(i) Then
                UserList(userindex).flags.MuereQuest(i) = UserList(userindex).flags.MuereQuest(i) + 1
            End If
            
            If (CompletoQuest = True) And (UserList(userindex).flags.MuereQuest(i) < QuestsList(NroQuest).CantNPC(i)) Then
                CompletoQuest = False
            End If
        Next i
            
       If CompletoQuest = True Then
            'Mision diaria
            If MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Tipo = 6 And MisionesDiarias(UserList(userindex).Misiones.NumeroMision).QuestNumber = UserList(userindex).flags.UserNumQuest Then
                If UserList(userindex).Misiones.ConteoUser < MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad Then UserList(userindex).Misiones.ConteoUser = UserList(userindex).Misiones.ConteoUser + 1
            End If
    
            If UserList(userindex).flags.estado = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||66")
                    Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(QuestsList(NroQuest).Oro * 2))
                    Call SendData(SendTarget.toindex, userindex, 0, "||57@" & QuestsList(NroQuest).Puntos * 2)
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + (val(QuestsList(NroQuest).Oro) * 2)
                    Call AgregarPuntos(userindex, (val(QuestsList(NroQuest).Puntos) * 2))
                    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + val(QuestsList(NroQuest).Puntos)
                    
                    modQuests.ResetQuest (userindex)
                    UserList(userindex).flags.QuestCompletadas = UserList(userindex).flags.QuestCompletadas + 1
            
            ElseIf UserList(userindex).flags.estado = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||66")
                    Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(QuestsList(NroQuest).Oro))
                    Call SendData(SendTarget.toindex, userindex, 0, "||57@" & QuestsList(NroQuest).Puntos)
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + (val(QuestsList(NroQuest).Oro))
                    Call AgregarPuntos(userindex, (val(QuestsList(NroQuest).Puntos)))
                    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + val(QuestsList(NroQuest).Puntos)
                    
                    modQuests.ResetQuest (userindex)
                    UserList(userindex).flags.QuestCompletadas = UserList(userindex).flags.QuestCompletadas + 1
            End If
            
            Call SendData(SendTarget.toindex, userindex, UserList(userindex).Pos.Map, "[Q]")
            SendUserGLD (userindex)
            
            If QuestsList(NroQuest).NumOBJ > 0 Then
                Dim OBJPremio As Obj
                OBJPremio.ObjIndex = QuestsList(NroQuest).NumOBJ
                OBJPremio.Amount = QuestsList(NroQuest).CantOBJ
                        
                If Not MeterItemEnInventario(userindex, OBJPremio) Then
                    Call TirarItemAlPiso(UserList(userindex).Pos, OBJPremio)
                End If
            End If
    End If
  End If

End Sub
Public Sub RestarUser(ByVal userindex As Integer, ByVal VictimIndex As Integer)

    Dim NroQuest As Byte
    NroQuest = UserList(userindex).flags.UserNumQuest

    If QuestsList(NroQuest).Tipo = 2 Then
        If UserList(userindex).flags.Questeando = 1 And TriggerZonaPelea(userindex, VictimIndex) <> TRIGGER6_PERMITE Then
            UserList(userindex).flags.MuereQuest(1) = UserList(userindex).flags.MuereQuest(1) + 1
        End If
         
        If UserList(userindex).flags.MuereQuest(1) = QuestsList(NroQuest).Usuarios Then
            'Mision diaria
            If MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Tipo = 6 And MisionesDiarias(UserList(userindex).Misiones.NumeroMision).QuestNumber = UserList(userindex).flags.UserNumQuest Then
                If UserList(userindex).Misiones.ConteoUser < MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad Then UserList(userindex).Misiones.ConteoUser = UserList(userindex).Misiones.ConteoUser + 1
            End If
    
            If UserList(userindex).flags.estado = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||66")
                    Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(QuestsList(NroQuest).Oro * 2))
                    Call SendData(SendTarget.toindex, userindex, 0, "||57@" & QuestsList(NroQuest).Puntos * 2)
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + (val(QuestsList(NroQuest).Oro) * 2)
                    Call AgregarPuntos(userindex, (val(QuestsList(NroQuest).Puntos) * 2))
                    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + val(QuestsList(NroQuest).Puntos)
                    
                    modQuests.ResetQuest (userindex)
                    UserList(userindex).flags.QuestCompletadas = UserList(userindex).flags.QuestCompletadas + 1
            
            ElseIf UserList(userindex).flags.estado = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||66")
                    Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(QuestsList(NroQuest).Oro))
                    Call SendData(SendTarget.toindex, userindex, 0, "||57@" & QuestsList(NroQuest).Puntos * 2)
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + (val(QuestsList(NroQuest).Oro))
                    Call AgregarPuntos(userindex, (val(QuestsList(NroQuest).Puntos)))
                    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + val(QuestsList(NroQuest).Puntos)
                    
                    modQuests.ResetQuest (userindex)
                    UserList(userindex).flags.QuestCompletadas = UserList(userindex).flags.QuestCompletadas + 1
            End If
            
            SendUserGLD (userindex)
            
            If QuestsList(NroQuest).NumOBJ > 0 Then
                Dim OBJPremio As Obj
                OBJPremio.ObjIndex = QuestsList(NroQuest).NumOBJ
                OBJPremio.Amount = QuestsList(NroQuest).CantOBJ
                        
                If Not MeterItemEnInventario(userindex, OBJPremio) Then
                    Call TirarItemAlPiso(UserList(userindex).Pos, OBJPremio)
                End If
            End If
        End If
    End If


End Sub
Public Sub ResetQuest(ByVal userindex As Integer)

        Dim g As Long

    For g = 1 To 3
        UserList(userindex).flags.MuereQuest(g) = 0
    Next g
        
        UserList(userindex).flags.Questeando = 0
        UserList(userindex).flags.UserNumQuest = 0
End Sub

Public Sub CargarMisiones()
    Dim n As Integer, LoopC As Integer, loopX As Integer
    n = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "INIT", "NumMisiones"))
    
    ReDim MisionesDiarias(n) As tMDiarias
    
    For LoopC = 1 To n
        MisionesDiarias(LoopC).Nombre = GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "Nombre")
        MisionesDiarias(LoopC).Info = GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "Info")
        MisionesDiarias(LoopC).Tipo = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "Tipo"))
        MisionesDiarias(LoopC).Cantidad = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "Cantidad"))
        MisionesDiarias(LoopC).NumNPC = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "NumNPC"))
        MisionesDiarias(LoopC).QuestNumber = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "QuestNumber"))
        MisionesDiarias(LoopC).PTS = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "PTS"))
        
        MisionesDiarias(LoopC).pOro = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "Oro"))
        MisionesDiarias(LoopC).pPuntos = val(GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "Puntos"))
        MisionesDiarias(LoopC).pObjetoIndex = ReadField(1, GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "OBJ"), Asc("-"))
        MisionesDiarias(LoopC).pObjetoAmount = ReadField(2, GetVar(App.Path & "\Dat\MisionesDiarias.dat", "MISION" & LoopC, "OBJ"), Asc("-"))
    Next LoopC
    
End Sub
Public Sub VerificarMisionDiaria(userindex As Integer)

    If UserList(userindex).Misiones.ConteoUser >= MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad Then
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pOro
        SendUserGLD (userindex)
        Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pOro))
        UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pPuntos
        Call SendData(SendTarget.toindex, userindex, 0, "||57@" & PonerPuntos(MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pPuntos))
        
        If MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pObjetoAmount > 0 Then
            Dim eOBJ As Obj
            eOBJ.ObjIndex = MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pObjetoIndex
            eOBJ.Amount = MisionesDiarias(UserList(userindex).Misiones.NumeroMision).pObjetoAmount
            
            If Not MeterItemEnInventario(userindex, eOBJ) Then
                Call TirarItemAlPiso(UserList(userindex).Pos, eOBJ)
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "||879")
        End If
        
        UserList(userindex).Misiones.Completada = 1
    End If
        
End Sub
