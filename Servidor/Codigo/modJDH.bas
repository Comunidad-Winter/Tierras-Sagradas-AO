Attribute VB_Name = "modJDH"
Option Explicit
 
'***************
'AUTOR: Toyz - Luciano
'FECHA: 14/12/16 - 07:30
'***************
Private Const Tiempo_Cancelamiento As Integer = 180
Private Const Cofre_Abierto As Integer = 10 'Número de cofre abierto.
Private Const Cofre_Cerrado As Integer = 11 'Número de cofre cerrado.
 
Private Type tUsuario
    ID As Integer
    Posicion As WorldPos
    X As Byte
    Y As Byte
End Type
 
Private Type tCofres
    Objetos(1 To 6) As Obj
    X As Byte
    Y As Byte
    Abierto As Boolean
End Type
 
Private Type tJDH
    Activo As Boolean
    Usuarios(1 To 1) As tUsuario
    Cofres(1 To 9) As tCofres
    Conteo As Integer
    Cupos As Byte
    mapa As Integer
    Premio As Long
    PremioPTS As Long
    Inscripcion As Long
    Total As Byte
    Restantes As Byte
End Type
 
Private JDH As tJDH
 
Public Sub Carga_JDH()
    Dim LoopC As Long
    Dim loopX As Long
    Dim LoopZ As Long
    Dim DataCofre As Obj
 
    DataCofre.Amount = 1
    DataCofre.ObjIndex = Cofre_Cerrado
 
    With JDH
        .Cupos = UBound(.Usuarios())
        .mapa = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "EVENTO", "Mapa")
        For LoopC = 1 To .Cupos
            .Usuarios(LoopC).X = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "USUARIO#" & LoopC, "X")
            .Usuarios(LoopC).Y = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "USUARIO#" & LoopC, "Y")
        Next LoopC
        For loopX = 1 To UBound(.Cofres())
            .Cofres(loopX).X = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "COFRE#" & loopX, "X")
            .Cofres(loopX).Y = GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "COFRE#" & loopX, "Y")
            MakeObj ToMap, 0, .mapa, DataCofre, .mapa, .Cofres(loopX).X, .Cofres(loopX).Y
            MapData(.mapa, .Cofres(loopX).X, .Cofres(loopX).Y).Blocked = 1
            MapData(.mapa, .Cofres(loopX).X, .Cofres(loopX).Y).Cofre = loopX
            Bloquear ToMap, 0, .mapa, .mapa, .Cofres(loopX).X, .Cofres(loopX).Y, 1
            For LoopZ = 1 To UBound(.Cofres(loopX).Objetos())
                .Cofres(loopX).Objetos(LoopZ).ObjIndex = CByte(ReadField(1, (GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "COFRE#" & loopX, "OBJETO#" & LoopZ)), 45))
                .Cofres(loopX).Objetos(LoopZ).Amount = CByte(ReadField(2, (GetVar(App.Path & "\Dat\JuegosDelHambre.dat", "COFRE#" & loopX, "OBJETO#" & LoopZ)), 45))
            Next LoopZ
        Next loopX
    End With
End Sub
 
Public Sub Armar_JDH(ByVal ID As Integer, ByVal Premio As Long, ByVal PremioPTS As Long, ByVal Inscripcion As Long)
    With JDH
        If .Activo = True Then
            Call SendData(SendTarget.toindex, ID, 0, "||884")
            Exit Sub
        End If
        
        .Inscripcion = Inscripcion
        .Premio = Premio
        .PremioPTS = PremioPTS
        .Total = .Cupos
        .Restantes = .Total
        .Activo = True
        .Conteo = Tiempo_Cancelamiento
        Call SendData(SendTarget.ToAll, 0, 0, "||885@" & .Cupos & "@" & .Inscripcion)
    End With
End Sub
 
Public Sub Entrar_JDH(ByVal ID As Integer)
    Dim ID_JDH As Byte
    With JDH
        If Puede_Entrar(ID) = False Then Exit Sub
        Call SendData(SendTarget.toindex, ID, 0, "||886")
        UserList(ID).Stats.GLD = UserList(ID).Stats.GLD - .Inscripcion
        .Cupos = .Cupos - 1
        ID_JDH = JDH_ID
        UserList(ID).flags.EnJDH = ID_JDH
        .Usuarios(ID_JDH).ID = ID
        .Usuarios(ID_JDH).Posicion = UserList(ID).Pos
        Save_Inventory (ID) '//Salvamos el inventario
        WarpUserChar ID, .mapa, .Usuarios(ID_JDH).X, .Usuarios(ID_JDH).Y, False
        SendUserGLD ID
        UserList(ID).flags.NotMove = True
        Call SendData(SendTarget.toindex, ID, 0, "STOPD" & UserList(ID).flags.NotMove)
        If .Cupos = 0 Then
            Call SendData(SendTarget.toindex, ID, 0, "||887")
            .Conteo = 10
            HayJDH = True
            frmMain.JDH.Enabled = True
        End If
    End With
End Sub
 
Private Function JDH_ID() As Byte
    Dim LoopC As Long
    With JDH
        For LoopC = 1 To .Total
            If .Usuarios(LoopC).ID = 0 Then
                JDH_ID = LoopC
                Exit Function
            End If
        Next LoopC
    End With
End Function
 
Private Function Puede_Entrar(ByVal ID As Integer) As Boolean
    Puede_Entrar = False
    If UserList(ID).flags.Muerto > 0 Then
        Call SendData(SendTarget.toindex, ID, 0, "||3")
        Exit Function
    End If
    If UserList(ID).flags.EnJDH > 0 Then
        Call SendData(SendTarget.toindex, ID, 0, "||97")
        Exit Function
    End If
    If JDH.Activo = False Then
        Call SendData(SendTarget.toindex, ID, 0, "||882")
        Exit Function
    End If
    If JDH.Cupos = 0 Then
        Call SendData(SendTarget.toindex, ID, 0, "||883")
        Exit Function
    End If
    If UserList(ID).Stats.GLD < JDH.Inscripcion Then
        Call SendData(SendTarget.toindex, ID, 0, "||663")
        Exit Function
    End If
    If Not UserList(ID).Pos.Map = 1 Then
        Call SendData(SendTarget.toindex, ID, 0, "||323")
        Exit Function
    End If
    Puede_Entrar = True
End Function
 
Public Sub Contar_JDH()
    Dim LoopC As Long
    Dim loopX As Long
    With JDH
        If .Conteo = 0 Then
            .Conteo = -1
            If .Activo = True Then
                For LoopC = 1 To .Total
                        UserList(.Usuarios(LoopC).ID).flags.NotMove = False
                        Call SendData(SendTarget.toindex, .Usuarios(LoopC).ID, 0, "STOPD" & UserList(.Usuarios(LoopC).ID).flags.NotMove)
                Next LoopC
                If .Cupos = 0 Then
                    Call SendData(SendTarget.ToMap, JDH.mapa, 0, "N|Juegos del Hambre> ¡YA!" & FONTTYPE_ORO)
                    frmMain.JDH.Enabled = False
                Else
                    Call SendData(SendTarget.ToMap, JDH.mapa, 0, "||890")
                    Cancelar_JDH
                End If
            End If
        End If
     
        If .Conteo > 0 Then
            If .Cupos = 0 Then _
                Call SendData(SendTarget.ToMap, JDH.mapa, 0, "N|Juegos del Hambre> " & .Conteo & FONTTYPE_INFO)
            .Conteo = .Conteo - 1
        End If
    End With
End Sub
 
Private Function ID_Usuario() As Byte
    Dim LoopC As Long
    For LoopC = 1 To JDH.Total
            If JDH.Usuarios(LoopC).ID > 0 Then
                ID_Usuario = LoopC
                Exit For
            End If
    Next LoopC
End Function
 
Public Sub Muere_JDH(ByVal ID As Integer)
    Dim ID_JDH As Byte
    ID_JDH = UserList(ID).flags.EnJDH
    If ID_JDH = 0 Then Exit Sub
    UserList(ID).flags.EnJDH = 0
    With JDH
        .Restantes = .Restantes - 1
        If .Restantes > 1 Then Call SendData(SendTarget.toindex, ID, 0, "||889@" & .Restantes)
        Call SendData(SendTarget.toindex, ID, 0, "||888")
        ReLoad_Inventory (ID)
        WarpUserChar ID, .Usuarios(ID_JDH).Posicion.Map, .Usuarios(ID_JDH).Posicion.X, .Usuarios(ID_JDH).Posicion.Y, False
        .Usuarios(ID_JDH).ID = 0
        If .Restantes <= 1 Then Finalizar
    End With
End Sub
Private Sub Finalizar()
    Dim LoopC As Long
    Dim Dame_ID As Byte
    Dim ID As Integer
    With JDH
        Dame_ID = ID_Usuario
        ID = .Usuarios(Dame_ID).ID
        Call SendData(SendTarget.ToAll, 0, 0, "||891@" & UserList(ID).Name)
        UserList(ID).Stats.GLD = UserList(ID).Stats.GLD + .Premio
        UserList(ID).Stats.PuntosTorneo = UserList(ID).Stats.PuntosTorneo + .PremioPTS
        SendUserGLD ID
        UserList(ID).flags.EnJDH = 0
        .Premio = 0
        .PremioPTS = 0
        ReLoad_Inventory (ID)
        WarpUserChar ID, .Usuarios(Dame_ID).Posicion.Map, .Usuarios(Dame_ID).Posicion.X, .Usuarios(Dame_ID).Posicion.Y, False
        Limpiar
        HayJDH = False
    End With
End Sub
Public Sub Cancelar_JDH()
    Dim LoopC As Long
    With JDH
        If .Activo = False Then Exit Sub
        For LoopC = 1 To .Total
            If .Usuarios(LoopC).ID > 0 Then
                ReLoad_Inventory (ID)
                WarpUserChar .Usuarios(LoopC).ID, .Usuarios(LoopC).Posicion.Map, .Usuarios(LoopC).Posicion.X, .Usuarios(LoopC).Posicion.Y, False
                UserList(.Usuarios(LoopC).ID).flags.EnJDH = 0
                UserList(.Usuarios(LoopC).ID).Stats.GLD = UserList(.Usuarios(LoopC).ID).Stats.GLD + .Inscripcion
                Call SendData(SendTarget.toindex, .Usuarios(LoopC).ID, 0, "||892")
                SendUserGLD .Usuarios(LoopC).ID
                HayJDH = False
            End If
        Next LoopC
    End With
    Limpiar
End Sub
 
Public Sub Desconexion_JDH(ByVal ID As Integer)
    If UserList(ID).flags.EnJDH = 0 Then Exit Sub
    With JDH
        TirarTodosLosItems ID
        ReLoad_Inventory (ID)
        WarpUserChar ID, .Usuarios(UserList(ID).flags.EnJDH).Posicion.Map, .Usuarios(UserList(ID).flags.EnJDH).Posicion.X, .Usuarios(UserList(ID).flags.EnJDH).Posicion.Y, True
        .Usuarios(UserList(ID).flags.EnJDH).ID = 0
        UserList(ID).flags.EnJDH = 0
        .Cupos = .Cupos + 1
    End With
End Sub
Private Sub Limpiar()
    Dim LoopC As Long
    With JDH
        .Activo = False
        .Conteo = -1
        .Cupos = UBound(.Usuarios())
        .Inscripcion = 0
        .Premio = 0
        .PremioPTS = 0
        .Restantes = 0
        .Total = 0
        For LoopC = 1 To .Total
            .Usuarios(LoopC).ID = 0
        Next LoopC
        ReCargar_Cofres
    End With
End Sub
Public Sub Clickea_Cofre(ByRef Pos As WorldPos)
    Dim ID As Byte
    Dim DataCofre As Obj
    Dim LoopC As Long
    Dim n_Pos As WorldPos
 
    DataCofre.Amount = 1
    DataCofre.ObjIndex = Cofre_Abierto
    ID = MapData(Pos.Map, Pos.X, Pos.Y).Cofre
 
    With JDH
        If ID = 0 Then Exit Sub
        If .Cupos > 0 Then Exit Sub
        If .Activo = False Then Exit Sub
        If .Cofres(ID).Abierto = True Then Exit Sub
        If .Conteo <> -1 Then Exit Sub
     
        .Cofres(ID).Abierto = True
     
        EraseObj ToMap, 0, .mapa, MapData(Pos.Map, Pos.X, Pos.Y).OBJInfo.Amount, Pos.Map, Pos.X, Pos.Y
        MakeObj ToMap, 0, .mapa, DataCofre, .mapa, .Cofres(ID).X, .Cofres(ID).Y
     
        For LoopC = 1 To UBound(.Cofres(ID).Objetos())
            Tilelibre Pos, n_Pos, .Cofres(ID).Objetos(LoopC)
            MakeObj ToMap, 0, .mapa, .Cofres(ID).Objetos(LoopC), .mapa, n_Pos.X, n_Pos.Y
        Next LoopC
    End With
End Sub
 
Private Sub ReCargar_Cofres()
    Dim DataCofre As Obj
    Dim LoopC As Long
 
    DataCofre.Amount = 1
    DataCofre.ObjIndex = Cofre_Cerrado
 
    With JDH
        For LoopC = 1 To UBound(.Cofres())
            .Cofres(LoopC).Abierto = False
            EraseObj ToMap, 0, .mapa, DataCofre.Amount, .mapa, .Cofres(LoopC).X, .Cofres(LoopC).Y
            MakeObj ToMap, 0, .mapa, DataCofre, .mapa, .Cofres(LoopC).X, .Cofres(LoopC).Y
        Next LoopC
    End With
    
    Call LimpiarMundoEntero
End Sub
Private Sub Save_Inventory(ByVal ID As Integer)
    
    '//Guardamos todo el inventario actual del usuario
    Dim LoopC As Long
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        UserList(ID).Invent.ExObject(LoopC).ObjIndex = UserList(ID).Invent.Object(LoopC).ObjIndex
        UserList(ID).Invent.ExObject(LoopC).Amount = UserList(ID).Invent.Object(LoopC).Amount
    Next LoopC
    
    '//Lo desnudamos
        Call LimpiarInventario(ID)
        Call DarCuerpoDesnudo(ID)
        
        UserList(ID).Invent.ArmourEqpSlot = 0
        UserList(ID).Invent.ArmourEqpObjIndex = 0
        
        UserList(ID).Invent.WeaponEqpObjIndex = 0
        UserList(ID).Invent.WeaponEqpSlot = 0
        UserList(ID).Char.CascoAnim = 0
        UserList(ID).Char.WeaponAnim = 0
        UserList(ID).Char.ShieldAnim = 0
        
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(ID).Pos.Map, val(ID), UserList(ID).Char.Body, UserList(ID).Char.Head, UserList(ID).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
        Call UpdateUserInv(True, ID, 0)
    
End Sub
Private Sub ReLoad_Inventory(ByVal ID As Integer)

        '//Lo desnudamos
        If UserList(userindex).flags.Muerto = 0 Then Call DarCuerpoDesnudo(ID)
        
        UserList(ID).Invent.ArmourEqpSlot = 0
        UserList(ID).Invent.ArmourEqpObjIndex = 0
        
        UserList(ID).Invent.WeaponEqpObjIndex = 0
        UserList(ID).Invent.WeaponEqpSlot = 0
        UserList(ID).Char.CascoAnim = 0
        UserList(ID).Char.WeaponAnim = 0
        UserList(ID).Char.ShieldAnim = 0
        
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(ID).Pos.Map, val(ID), UserList(ID).Char.Body, UserList(ID).Char.Head, UserList(ID).Char.Heading, NingunArma, NingunEscudo, NingunCasco)

    '//Devolvemos el inventario que tenía antes de ingresar
    Dim LoopC As Long
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        UserList(ID).Invent.Object(LoopC).ObjIndex = UserList(ID).Invent.ExObject(LoopC).ObjIndex
        UserList(ID).Invent.Object(LoopC).Amount = UserList(ID).Invent.ExObject(LoopC).Amount
    Next LoopC
    
    '//Ahora le reiniciamos el salvado.
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        UserList(ID).Invent.ExObject(LoopC).ObjIndex = 0
        UserList(ID).Invent.ExObject(LoopC).Amount = 0
        UserList(ID).Invent.ExObject(LoopC).Equipped = 0
    Next LoopC

    Call UpdateUserInv(True, ID, 0)

End Sub
