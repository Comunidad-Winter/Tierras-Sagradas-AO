Attribute VB_Name = "TorneosAutomaticos"
Option Explicit
Public Torneo_Activo As Boolean
Public Torneo_Esperando As Boolean
Private Torneo_Rondas As Integer
Private Torneo_Luchadores() As Integer
 
Private Const mapatorneo As Integer = 100
' esquinas superior isquierda del ring
Private Const esquina1x As Integer = 41
Private Const esquina1y As Integer = 42
' esquina inferior derecha del ring
Private Const esquina2x As Integer = 60
Private Const esquina2y As Integer = 57
' Donde esperan
Private Const esperax As Integer = 27
Private Const esperay As Integer = 43
' Mapa desconecta
Private Const mapa_fuera As Integer = 1
Private Const fueraesperay As Integer = 45
Private Const fueraesperax As Integer = 55
 ' estas son las pocisiones de las 2 esquinas de la zona de espera, en su mapa tienen que tener en la misma posicion las 2 esquinas.
Private Const X1 As Integer = 23
Private Const X2 As Integer = 31
Private Const Y1 As Integer = 37
Private Const Y2 As Integer = 58
Sub Torneoauto_Cancela()
On Error GoTo errorh:
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.ToAll, 0, 0, "||1")
    Dim i As Integer
     For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuera
                    FuturePos.X = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                      UserList(Torneo_Luchadores(i)).flags.Automatico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_Cancela()
On Error GoTo errorh
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.ToAll, 0, 0, "||2")
    Dim i As Integer
    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuera
                    FuturePos.X = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                    UserList(Torneo_Luchadores(i)).flags.Automatico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_UsuarioMuere(ByVal userindex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
On Error GoTo rondas_usuariomuere_errorh
        Dim i As Integer, Pos As Integer, j As Integer
        Dim combate As Integer, LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
If (Not Torneo_Activo) Then
                Exit Sub
            ElseIf (Torneo_Activo And Torneo_Esperando) Then
                For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                    If (Torneo_Luchadores(i) = userindex) Then
                        Torneo_Luchadores(i) = -1
                        Call WarpUserChar(userindex, mapa_fuera, fueraesperay, fueraesperax, True)
                         UserList(userindex).flags.Automatico = False
                        Exit Sub
                    End If
                Next i
                Exit Sub
            End If
 
        For Pos = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(Pos) = userindex) Then Exit For
        Next Pos
 
        ' si no lo ha encontrado
        If (Torneo_Luchadores(Pos) <> userindex) Then Exit Sub
       
 '  Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.
 
If UserList(userindex).Pos.X >= X1 And UserList(userindex).Pos.X <= X2 And UserList(userindex).Pos.Y >= Y1 And UserList(userindex).Pos.Y <= Y2 Then
Call SendData(SendTarget.ToAll, 0, 0, "||478@" & UserList(userindex).Name)
Call WarpUserChar(userindex, mapa_fuera, fueraesperax, fueraesperay, True)
UserList(userindex).flags.Automatico = False
Torneo_Luchadores(Pos) = -1
Exit Sub
End If
 
        combate = 1 + (Pos - 1) \ 2
 
        'ponemos li1 y li2 (luchador index) de los que combatian
        LI1 = 2 * (combate - 1) + 1
        LI2 = LI1 + 1
 
        'se informa a la gente
        If (Real) Then
                Call SendData(SendTarget.ToAll, 0, 0, "||479@" & UserList(userindex).Name)
        Else
                Call SendData(SendTarget.ToAll, 0, 0, "||478@" & UserList(userindex).Name)
        End If
 
        'se le teleporta fuera si murio
        If (Real) Then
                Call WarpUserChar(userindex, mapa_fuera, fueraesperax, fueraesperay, True)
                 UserList(userindex).flags.Automatico = False
        ElseIf (Not CambioMapa) Then
             
                 Call WarpUserChar(userindex, mapa_fuera, fueraesperax, fueraesperay, True)
                  UserList(userindex).flags.Automatico = False
        End If
 
        'se le borra de la lista y se mueve el segundo a li1
        If (Torneo_Luchadores(LI1) = userindex) Then
                Torneo_Luchadores(LI1) = Torneo_Luchadores(LI2) 'cambiamos slot
                Torneo_Luchadores(LI2) = -1
        Else
                Torneo_Luchadores(LI2) = -1
        End If
 
    'si es la ultima ronda
    If (Torneo_Rondas = 1) Then
    
        Call SendData(SendTarget.ToAll, 0, 0, "||480@" & UserList(Torneo_Luchadores(LI1)).Name)
        Call SendData(SendTarget.ToAll, 0, 0, "||481")
        
        Dim medallaoro As Obj
        medallaoro.Amount = 1
        medallaoro.ObjIndex = 1025
        
        If Not MeterItemEnInventario(Torneo_Luchadores(LI1), medallaoro) Then
            Call TirarItemAlPiso(UserList(Torneo_Luchadores(LI1)).Pos, medallaoro)
        End If
    
        UserList(Torneo_Luchadores(LI1)).Stats.MedOro = UserList(Torneo_Luchadores(LI1)).Stats.MedOro + 1
        UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + 100
        Call WarpUserChar(Torneo_Luchadores(LI1), mapa_fuera, fueraesperax, fueraesperay, True)
    
        Call GRANK_User_Check(Events, UserList(Torneo_Luchadores(LI1)).Name, UserList(Torneo_Luchadores(LI1)).Stats.MedOro)
        Call WriteVar(CharPath & UserList(Torneo_Luchadores(LI1)).Name & ".chr", "STATS", "MedOro", UserList(Torneo_Luchadores(LI1)).Stats.MedOro)
        
        Call SendData(toindex, Torneo_Luchadores(LI1), 0, "||57@50")
        Call AgregarPuntos(Torneo_Luchadores(LI1), 50)
        Call WriteVar(CharPath & UserList(Torneo_Luchadores(LI1)).Name & ".chr", "STATS", "PuntosTorneo", UserList(Torneo_Luchadores(LI1)).Stats.PuntosTorneo)
    
        UserList(Torneo_Luchadores(LI1)).flags.Automatico = False
        Torneo_Activo = False
        Exit Sub
    Else
        'a su compañero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Torneo_Luchadores(LI1), 100, esperax, esperay, True)
    End If
 
               
        'si es el ultimo combate de la ronda
        If (2 ^ Torneo_Rondas = 2 * combate) Then
 
                Call SendData(SendTarget.ToAll, 0, 0, "||482")
                Torneo_Rondas = Torneo_Rondas - 1
 
        'antes de llamar a la proxima ronda hay q copiar a los tipos
        For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
        Next i
ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
        Call Rondas_Combate(1)
        Exit Sub
        End If
 
        'vamos al siguiente combate
        Call Rondas_Combate(combate + 1)
rondas_usuariomuere_errorh:
 
End Sub
 
 
 
Sub Rondas_UsuarioDesconecta(ByVal userindex As Integer)
On Error GoTo errorh
Call SendData(SendTarget.ToAll, 0, 0, "||478@" & UserList(userindex).Name)
Call Rondas_UsuarioMuere(userindex, False, False)
errorh:
End Sub
Sub Rondas_UsuarioCambiamapa(ByVal userindex As Integer)
On Error GoTo errorh
        Call Rondas_UsuarioMuere(userindex, False, True)
errorh:
End Sub
 
Sub Torneos_Inicia(ByVal userindex As Integer, ByVal rondas As Integer)
On Error GoTo errorh
        If (Torneo_Activo) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||484")
                Exit Sub
        End If
        
        Call SendData(SendTarget.ToAll, 0, 0, "||485@" & UserList(userindex).Name & "@" & val(2 ^ rondas))
        CuentaAutomatico = 10
        Call SendData(SendTarget.ToAll, 0, 0, "TW48")
       
        Torneo_Rondas = rondas
        Torneo_Esperando = True
 
        ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
        Dim i As Integer
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                Torneo_Luchadores(i) = -1
        Next i
errorh:
End Sub
 
 
 
Sub Torneos_Entra(ByVal userindex As Integer)
On Error GoTo errorh
        Dim i As Integer
       
        If (Not Torneo_Activo) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||486")
                Exit Sub
        End If
        
            If UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or MapInfo(UserList(userindex).Pos.Map).Pk = True Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 118 Or UserList(userindex).Pos.Map = 120 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||291")
                Exit Sub
            End If
            
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||3")
          Exit Sub
        End If
        
        If (Not Torneo_Esperando) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||487")
                Exit Sub
        End If
        
        If UserList(userindex).flags.Muerto = 1 Then Exit Sub
       
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) = userindex) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||488")
                        Exit Sub
                End If
        Next i
 
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        If (Torneo_Luchadores(i) = -1) Then
                Torneo_Luchadores(i) = userindex
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = 100
                    FuturePos.X = RandomNumber(23, 32): FuturePos.Y = RandomNumber(37, 58)
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                   
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                 UserList(Torneo_Luchadores(i)).flags.Automatico = True
                 
                Call SendData(SendTarget.toindex, userindex, 0, "||489")
               
                Call SendData(SendTarget.ToAll, 0, 0, "||490@" & UserList(userindex).Name)
                If (i = UBound(Torneo_Luchadores)) Then
                Call SendData(SendTarget.ToAll, 0, 0, "||491")
                Torneo_Esperando = False
                Call Rondas_Combate(1)
     
                End If
                  Exit Sub
        End If
        Next i
errorh:
End Sub
 
 
Sub Rondas_Combate(combate As Integer)
On Error GoTo errorh
Dim UI1 As Integer, UI2 As Integer
    UI1 = Torneo_Luchadores(2 * (combate - 1) + 1)
    UI2 = Torneo_Luchadores(2 * combate)
   
    If (UI2 = -1) Then
        UI2 = Torneo_Luchadores(2 * (combate - 1) + 1)
        UI1 = Torneo_Luchadores(2 * combate)
    End If
   
    If (UI1 = -1) Then
        Call SendData(SendTarget.ToAll, 0, 0, "||492")
        If (Torneo_Rondas = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.ToAll, 0, 0, "||493@" & UserList(UI2).Name)
                UserList(UI2).flags.Automatico = False
                ' dale_recompensa()
                Torneo_Activo = False
                Exit Sub
            End If
            Call SendData(SendTarget.ToAll, 0, 0, "||494")
            Exit Sub
        End If
        If (UI2 <> -1) Then _
            Call SendData(SendTarget.ToAll, 0, 0, "||495@" & UserList(UI2).Name)
   
        If (2 ^ Torneo_Rondas = 2 * combate) Then
            Call SendData(SendTarget.ToAll, 0, 0, "||496")
            Torneo_Rondas = Torneo_Rondas - 1
            'antes de llamar a la proxima ronda hay q copiar a los tipos
            Dim i As Integer, j As Integer
            For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
            Next i
            ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
            Call Rondas_Combate(1)
            Exit Sub
        End If
        Call Rondas_Combate(combate + 1)
        Exit Sub
    End If
 
    UserList(UI1).Stats.MinHP = UserList(UI1).Stats.MaxHP
    UserList(UI2).Stats.MinHP = UserList(UI2).Stats.MaxHP
    UserList(UI1).Stats.MinMAN = UserList(UI1).Stats.MaxMAN
    UserList(UI2).Stats.MinMAN = UserList(UI2).Stats.MaxMAN
    SendUserHP (UI1)
    SendUserMP (UI1)
    
    SendUserHP (UI2)
    SendUserMP (UI2)
    
    Call SendData(SendTarget.ToAll, 0, 0, "||497@" & UserList(UI1).Name & "@" & UserList(UI2).Name)
 
    Call WarpUserChar(UI1, mapatorneo, esquina1x, esquina1y, True)
    Call WarpUserChar(UI2, mapatorneo, esquina2x, esquina2y, True)
errorh:
End Sub

