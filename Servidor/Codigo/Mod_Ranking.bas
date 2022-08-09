Attribute VB_Name = "Mod_Ranking"

Option Explicit

'@ Ranking general (contiene estadísticas que no se reinician,
'  por lo tanto duran lo que dure la versión)

'Límite del ranking
Private Const kGRANK_LIMIT As Integer = 10

'Tipos de ranking
Public Enum ENU_GRANK_Mode
    LowerBound
    
    Duels
    Kills
    Rounds
    Couple
    Tournaments
    Events
    GuildVSGuild
    GuildReputation
    Castles
    
    'Semanal
    Reputation
    
    UpperBound
End Enum

'Individuo perteneciente al ranking
Private Type TYP_GRANK_Slot
    Name As String
    Value As Long
End Type

'Rankings
Private m_rank(ENU_GRANK_Mode.LowerBound + 1 To ENU_GRANK_Mode.UpperBound - 1, 1 To kGRANK_LIMIT) As TYP_GRANK_Slot
'Carga el archivo del ranking
Public Sub GRANK_Setup()

Dim l_file As clsIniReader
Dim l_name As String
Dim l_line As String

Dim i As Long
Dim j As Long

    Set l_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Ranking.txt"
        
    '@ load all ranks
    For i = ENU_GRANK_Mode.LowerBound + 1 To ENU_GRANK_Mode.UpperBound - 1
        '@ get current rank's name
        l_name = GRANK_Name_Get(i)
        
        For j = 1 To kGRANK_LIMIT
            '@ get line
            l_line = l_file.GetValue(l_name, j)
            
            '@ store
            m_rank(i, j).Name = ReadField(1, l_line, Asc("-"))
            m_rank(i, j).Value = ReadField(2, l_line, Asc("-"))
            
            If j <= 3 Then
                Call SetStars(i, j)
            End If
            
        Next
    
    Next

End Sub
Private Sub SetStars(ByVal mode_ As ENU_GRANK_Mode, ByVal Pos As Integer)

    If mode_ = Duels Then
        If m_rank(mode_, Pos).Value > 0 Then Estrella.TOPDuelos(Pos) = m_rank(mode_, Pos).Name
    ElseIf mode_ = Couple Then
        If m_rank(mode_, Pos).Value > 0 Then Estrella.TOPParejas(Pos) = m_rank(mode_, Pos).Name
    ElseIf mode_ = Tournaments Then
        If m_rank(mode_, Pos).Value > 0 Then Estrella.TOPTorneos(Pos) = m_rank(mode_, Pos).Name
    ElseIf mode_ = Events Then
        If m_rank(mode_, Pos).Value > 0 Then Estrella.TOPEvents(Pos) = m_rank(mode_, Pos).Name
    ElseIf mode_ = Kills Then
        If m_rank(mode_, Pos).Value > 0 Then Estrella.TOPFrags(Pos) = m_rank(mode_, Pos).Name
    ElseIf mode_ = Rounds Then
        If m_rank(mode_, Pos).Value > 0 Then Estrella.TOPRondas(Pos) = m_rank(mode_, Pos).Name
    End If
    
End Sub
'Guarda el ranking en el archivo
Public Sub GRANK_Dump(ByVal mode_ As ENU_GRANK_Mode)

On Error Resume Next

Dim i As Long
Dim l_name As String
Dim l_line As String
    
    '@ get current rank's name
    l_name = GRANK_Name_Get(mode_)

    For i = 1 To kGRANK_LIMIT
    
        Call WriteVar(App.Path & "\Ranking.txt", l_name, i, m_rank(mode_, i).Name & "-" & m_rank(mode_, i).Value)
        
        If i <= 3 Then
            Call SetStars(mode_, i)
        End If
    Next

End Sub
'Comprueba si un usuario entró/subió en el ranking
Public Sub GRANK_User_Check(ByVal mode_ As ENU_GRANK_Mode, ByRef name_ As String, ByVal value_ As Long)

Dim l_slot As Long
Dim i As Long
    
    If IsAdministrator(name_) Or IsDevelopment(name_) Or IsCoordination(name_) Or IsTournamentManager(name_) Or IsEventManager(name_) Or IsUserSupport(name_) Then Exit Sub

    '@ find user
    l_slot = GRANK_User_In(mode_, name_)

    '@ he ain't here
    If l_slot = -1 Then l_slot = kGRANK_LIMIT
    
    '@ check slots in reverse
    For i = l_slot To 1 Step -1
        If m_rank(mode_, i).Value > value_ Then Exit For
    Next
    
    '@ he's not capable of entering the ranking
    If (i = l_slot) Then
        If (m_rank(mode_, i).Name <> name_) Then Exit Sub
    End If
    
    '@ our real slot is i+1, so we fix it
    i = i + 1
    
    '@ move list down
    For i = l_slot To i + 1 Step -1
    
        '@ copy
        m_rank(mode_, i) = m_rank(mode_, i - 1)
    
    Next
    
    '@ update data
    m_rank(mode_, i).Name = name_
    m_rank(mode_, i).Value = value_
    
    '@ save data
    GRANK_Dump mode_
    
End Sub 'Devuelve el slot del ranking en el que se encuentra el usuario, -1 si no se encuentra
Private Function GRANK_User_In(ByVal mode_ As ENU_GRANK_Mode, ByRef name_ As String) As Long

Dim i As Long

    For i = 1 To kGRANK_LIMIT
    
        '@ found them
        If UCase$(m_rank(mode_, i).Name) = UCase$(name_) Then
        
            '@ return, leave
            GRANK_User_In = i
            Exit Function
        
        End If
    
    Next
    
    '@ not found
    GRANK_User_In = -1

End Function
Public Sub Info_Rank(ByVal Rank As ENU_GRANK_Mode, ByVal userindex As Integer)

    
        Dim tStr As String, i As Long
        tStr = ""
            For i = 1 To 10
                tStr = tStr & m_rank(Rank, i).Name & "-" & m_rank(Rank, i).Value & ","
            Next i
            
            If Rank = Reputation Then
                tStr = tStr & UserList(userindex).Stats.Reputacione
            End If
            
            Call SendData(SendTarget.toindex, userindex, 0, "MTOP" & tStr)

End Sub

'Devuelve el nombre de un ranking con el que está ubicado en el archivo
Private Function GRANK_Name_Get(ByVal rank_ As ENU_GRANK_Mode) As String

    Select Case rank_
    
        Case ENU_GRANK_Mode.Duels
            GRANK_Name_Get = "DUELS"
            
        Case ENU_GRANK_Mode.Kills
            GRANK_Name_Get = "KILLS"
            
        Case ENU_GRANK_Mode.Reputation
            GRANK_Name_Get = "REPUTATION"
            
        Case ENU_GRANK_Mode.Tournaments
            GRANK_Name_Get = "TOURNAMENTS"
            
        Case ENU_GRANK_Mode.Events
            GRANK_Name_Get = "EVENTS"
            
        Case ENU_GRANK_Mode.Couple
            GRANK_Name_Get = "COUPLE"
            
        Case ENU_GRANK_Mode.Rounds
            GRANK_Name_Get = "ROUNDS"
            
        Case ENU_GRANK_Mode.GuildVSGuild
            GRANK_Name_Get = "CVC"
            
        Case ENU_GRANK_Mode.Castles
            GRANK_Name_Get = "CASTLES"
            
        Case ENU_GRANK_Mode.GuildReputation
            GRANK_Name_Get = "GUILDREPUTATION"
    
    End Select
    
End Function
Public Sub SRANK_Gives()

Dim mode_ As ENU_GRANK_Mode, TopIndex As Integer, TopName As String
Dim NumCorreos As Byte
Dim NueCorreos As String
Dim NTCR As String
Dim CorreoTemporal As String
Dim iMoC As Long

Dim TempActual As Long

mode_ = Reputation

TopIndex = NameIndex(m_rank(mode_, 1).Name)
TopName = m_rank(mode_, 1).Name
      
If FileExist(CharPath & TopName & ".chr") = True Then
    'Usuario ganador: TOP 1
    
      If TopIndex <> 0 Then
        UserList(TopIndex).flags.NumCorreos = UserList(TopIndex).flags.NumCorreos + 1
        UserList(TopIndex).flags.Correo(UserList(TopIndex).flags.NumCorreos) = "Servidor$Recibiste un objeto$El ranking semanal fue finalizado, recibes estos objetos por haber terminado en 1ra posición.$" & Date & "$1549-3-Cofre Común,0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),"
        UserList(TopIndex).flags.NueCorreos(UserList(TopIndex).flags.NumCorreos) = 1
        Call SendData(SendTarget.toindex, TopIndex, 0, "||631")
        
        UserList(TopIndex).Stats.GLD = UserList(TopIndex).Stats.GLD + 500000
        UserList(TopIndex).Stats.PuntosTorneo = UserList(TopIndex).Stats.PuntosTorneo + 50
        
      Else
            NumCorreos = GetVar(CharPath & TopName & ".chr", "CORREO", "NUMCORREOS")
            NueCorreos = GetVar(CharPath & TopName & ".chr", "CORREO", "NUECORREOS")
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "CORREONUM" & NumCorreos + 1, "Servidor$Recibiste un objeto$El ranking semanal fue finalizado, recibes estos objetos por haber terminado en 1ra posición.$" & Date & "1549-3-Cofre Común,0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),")
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "NUMCORREOS", NumCorreos + 1)
            
            For iMoC = 1 To 30
                CorreoTemporal = ReadField(iMoC, NueCorreos, Asc(","))
                If iMoC = NumCorreos + 1 Then
                    NTCR = NTCR & iMoC & "-1,"
                Else
                    NTCR = NTCR & iMoC & "-" & ReadField(2, CorreoTemporal, Asc("-")) & ","
                End If
            Next iMoC
            
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "NUECORREOS", NTCR)
            
            'ORO
            TempActual = GetVar(CharPath & TopName & ".chr", "STATS", "GLD")
            Call WriteVar(CharPath & TopName & ".chr", "STATS", "GLD", TempActual + 500000)
            'PUNTOS
            TempActual = GetVar(CharPath & TopName & ".chr", "STATS", "PUNTOSTORNEO")
            Call WriteVar(CharPath & TopName & ".chr", "STATS", "PUNTOSTORNEO", TempActual + 50)
        End If
End If
        

TopIndex = NameIndex(m_rank(mode_, 2).Name)
TopName = m_rank(mode_, 2).Name
      

If FileExist(CharPath & TopName & ".chr") = True Then
    'Usuario ganador: TOP 2
    
      If TopIndex <> 0 Then
        UserList(TopIndex).flags.NumCorreos = UserList(TopIndex).flags.NumCorreos + 1
        UserList(TopIndex).flags.Correo(UserList(TopIndex).flags.NumCorreos) = "Servidor$Recibiste un objeto$El ranking semanal fue finalizado, recibes estos objetos por haber terminado en 2da posición.$" & Date & "$1549-2-Cofre Común,0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),"
        UserList(TopIndex).flags.NueCorreos(UserList(TopIndex).flags.NumCorreos) = 1
        Call SendData(SendTarget.toindex, TopIndex, 0, "||631")
        
        UserList(TopIndex).Stats.GLD = UserList(TopIndex).Stats.GLD + 300000
        UserList(TopIndex).Stats.PuntosTorneo = UserList(TopIndex).Stats.PuntosTorneo + 35
        
      Else
            NumCorreos = GetVar(CharPath & TopName & ".chr", "CORREO", "NUMCORREOS")
            NueCorreos = GetVar(CharPath & TopName & ".chr", "CORREO", "NUECORREOS")
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "CORREONUM" & NumCorreos + 1, "Servidor$Recibiste un objeto$El ranking semanal fue finalizado, recibes estos objetos por haber terminado en 2da posición.$" & Date & "$1549-2-Cofre Común,0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),")
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "NUMCORREOS", NumCorreos + 1)
            
            For iMoC = 1 To 30
                CorreoTemporal = ReadField(iMoC, NueCorreos, Asc(","))
                If iMoC = NumCorreos + 1 Then
                    NTCR = NTCR & iMoC & "-1,"
                Else
                    NTCR = NTCR & iMoC & "-" & ReadField(2, CorreoTemporal, Asc("-")) & ","
                End If
            Next iMoC
            
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "NUECORREOS", NTCR)
            
            'ORO
            TempActual = GetVar(CharPath & TopName & ".chr", "STATS", "GLD")
            Call WriteVar(CharPath & TopName & ".chr", "STATS", "GLD", TempActual + 300000)
            'PUNTOS
            TempActual = GetVar(CharPath & TopName & ".chr", "STATS", "PUNTOSTORNEO")
            Call WriteVar(CharPath & TopName & ".chr", "STATS", "PUNTOSTORNEO", TempActual + 35)
        End If
End If


TopIndex = NameIndex(m_rank(mode_, 3).Name)
TopName = m_rank(mode_, 3).Name

If FileExist(CharPath & TopName & ".chr") = True Then
    'Usuario ganador: TOP 3
    
      If TopIndex <> 0 Then
        UserList(TopIndex).flags.NumCorreos = UserList(TopIndex).flags.NumCorreos + 1
        UserList(TopIndex).flags.Correo(UserList(TopIndex).flags.NumCorreos) = "Servidor$Recibiste un objeto$El ranking semanal fue finalizado, recibes estos objetos por haber terminado en 3ra posición.$" & Date & "$1549-1-Cofre Común,0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),"
        UserList(TopIndex).flags.NueCorreos(UserList(TopIndex).flags.NumCorreos) = 1
        Call SendData(SendTarget.toindex, TopIndex, 0, "||631")
        
        UserList(TopIndex).Stats.GLD = UserList(TopIndex).Stats.GLD + 150000
        UserList(TopIndex).Stats.PuntosTorneo = UserList(TopIndex).Stats.PuntosTorneo + 20
        
      Else
            NumCorreos = GetVar(CharPath & TopName & ".chr", "CORREO", "NUMCORREOS")
            NueCorreos = GetVar(CharPath & TopName & ".chr", "CORREO", "NUECORREOS")
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "CORREONUM" & NumCorreos + 1, "Servidor$Recibiste un objeto$El ranking semanal fue finalizado, recibes estos objetos por haber terminado en 3ra posición.$" & Date & "$1549-1-Cofre Común,0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),")
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "NUMCORREOS", NumCorreos + 1)
            
            For iMoC = 1 To 30
                CorreoTemporal = ReadField(iMoC, NueCorreos, Asc(","))
                If iMoC = NumCorreos + 1 Then
                    NTCR = NTCR & iMoC & "-1,"
                Else
                    NTCR = NTCR & iMoC & "-" & ReadField(2, CorreoTemporal, Asc("-")) & ","
                End If
            Next iMoC
            
            Call WriteVar(CharPath & TopName & ".chr", "CORREO", "NUECORREOS", NTCR)
            
            'ORO
            TempActual = GetVar(CharPath & TopName & ".chr", "STATS", "GLD")
            Call WriteVar(CharPath & TopName & ".chr", "STATS", "GLD", TempActual + 150000)
            'PUNTOS
            TempActual = GetVar(CharPath & TopName & ".chr", "STATS", "PUNTOSTORNEO")
            Call WriteVar(CharPath & TopName & ".chr", "STATS", "PUNTOSTORNEO", TempActual + 20)
        End If
End If

End Sub
Public Sub ResetReputation()

Dim mode_ As ENU_GRANK_Mode, i As Integer
mode_ = Reputation

haciendoBK = True
Call SendData(SendTarget.ToAll, 0, 0, "BKW")
Call SendData(SendTarget.ToAll, 0, 0, "||880")

    For i = 1 To LastUser
        If UserList(i).flags.UserLogged = True Then UserList(i).Stats.Reputacione = 0
    Next
    
    Dim DirChar As String
    DirChar = Dir(App.Path & "\Charfile\")
     
    Do While DirChar <> ""
      If InStr(1, LCase$(DirChar), ".chr") > 0 Then
        Call WriteVar(CharPath & DirChar & ".chr", "STATS", "Reputacione", "0")
      End If
      DirChar = Dir
    Loop
    
    For i = 1 To 10
        m_rank(mode_, i).Name = "N/A"
        m_rank(mode_, i).Value = 0
    Next

'@ save data
GRANK_Dump mode_

Call SendData(SendTarget.ToAll, 0, 0, "BKW")
haciendoBK = False
Call SendData(SendTarget.ToAll, 0, 0, "||881")

End Sub
