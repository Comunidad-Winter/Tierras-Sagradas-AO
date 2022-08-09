Attribute VB_Name = "Mod_AntiCheat"
Option Explicit
 
 
Public Type Intervalos
 
     Poteo As Long
 
     Golpe As Integer
 
     Casteo As Integer
     
     Click As Integer
     
     Flechas As Integer

     Trabajar As Integer
 
End Type

Private Type tIntervals
    Golpe As Integer
    Flechas As Integer
    LanzarHechizo As Integer
    PoteoU As Integer
    PoteoClick As Integer
    Work As Integer
End Type

Public setIntervals As tIntervals
 
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub CargarIntervalos()

Dim l_file As clsIniReader

    Set l_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Dat\Intervalos.ini"
    
    setIntervals.Golpe = l_file.GetValue("INTERVALOS", "Golpe")
    setIntervals.Flechas = l_file.GetValue("INTERVALOS", "Flechas")
    setIntervals.LanzarHechizo = l_file.GetValue("INTERVALOS", "LanzarHechizo")
    setIntervals.PoteoU = l_file.GetValue("INTERVALOS", "PoteoU")
    setIntervals.PoteoClick = l_file.GetValue("INTERVALOS", "PoteoClick")
    setIntervals.Work = l_file.GetValue("INTERVALOS", "Work")

End Sub
Public Sub RestoTiempo(ByVal userindex As Integer)
 
 
     '// Miqueas150
 
        '// Vamos restando tiempo a os intervalos para poder ejecutarlos :v
 
 
     With UserList(userindex).Counters
 
 
             If .Seguimiento.Golpe > 0 Then '// Restamos al intervalo "Golpe" para poder pegar
 
                        .Seguimiento.Golpe = .Seguimiento.Golpe - 1
 
                End If
 
 
                If .Seguimiento.Casteo > 0 Then '// Restamos al intervalo "Casteo" para poder pegar
 
                     .Seguimiento.Casteo = .Seguimiento.Casteo - 1
 
             End If
             
                If .Seguimiento.Flechas > 0 Then '// Restamos al intervalo "Flechas" para poder pegar
 
                     .Seguimiento.Flechas = .Seguimiento.Flechas - 1
 
             End If
             
                If .Seguimiento.Poteo > 0 Then '// Restamos al intervalo "Potear" para poder potear
 
                     .Seguimiento.Poteo = .Seguimiento.Poteo - 1
 
             End If
             
                If .Seguimiento.Click > 0 Then '// Restamos al intervalo "Clickear" para poder potear
 
                     .Seguimiento.Click = .Seguimiento.Click - 1
 
             End If
             
             
                If .Seguimiento.Trabajar > 0 Then '// Restamos al intervalo "Trabajar" para poder potear
 
                     .Seguimiento.Trabajar = .Seguimiento.Trabajar - 1
 
             End If
 
     End With
 
 
End Sub
 

Public Sub SetIntervalos(ByVal userindex As Integer)
 
 
     '// Miqueas150
 
        '// Seteamos las Variables a 0
 
 
     With UserList(userindex).Counters
 
 
             .Seguimiento.Casteo = 0
 
             .Seguimiento.Golpe = 0
             
             .Seguimiento.Flechas = 0
             
             .Seguimiento.Poteo = 0
             
             .Seguimiento.Click = 0
             
             .Seguimiento.Trabajar = 0
 
     End With
 
 
End Sub
 
 
Public Function PuedoCasteoHechizo(ByVal userindex As Integer) As Boolean
 
 
     '// Miqueas
 
        '// Controlamos que pueda Tirar Hechizos
 
 
     With UserList(userindex).Counters
 
 
             If .Seguimiento.Casteo > 0 Then
 
                     PuedoCasteoHechizo = False
 
 
                     Exit Function
 
 
             End If
 
       
 
             PuedoCasteoHechizo = True
 
       
 
             '// 21 * 40 = 840 Mseg entre casteo y casteo
 
                .Seguimiento.Casteo = setIntervals.LanzarHechizo
 
 
        End With
 
 
End Function
Public Function PuedoTrabajar(ByVal userindex As Integer) As Boolean
 
 
     '// Miqueas
 
        '// Controlamos que pueda Tirar Hechizos
 
 
     With UserList(userindex).Counters
 
 
             If .Seguimiento.Trabajar > 0 Then
 
                     PuedoTrabajar = False
 
 
                     Exit Function
 
 
             End If
 
       
 
             PuedoTrabajar = True
 
       
 
             '// 21 * 40 = 840 Mseg entre casteo y casteo
 
                .Seguimiento.Trabajar = setIntervals.Work
 
 
        End With
 
 
End Function
Public Function PuedoFlechear(ByVal userindex As Integer) As Boolean
 
 
        '// Miqueas
 
     '// Controlamos que pueda Pegar
 
 
        With UserList(userindex).Counters
 
 
                If .Seguimiento.Flechas > 0 Then
 
                        PuedoFlechear = False
 
 
                        Exit Function
 
 
                End If
 
             
 
                PuedoFlechear = True
 
             
 
                '// 28*40 = 1120 Mseg Entre golpe y golpe
 
             .Seguimiento.Golpe = 25
             .Seguimiento.Flechas = setIntervals.Flechas
 
     End With
 
 
End Function
 Public Function PuedoClickear(ByVal userindex As Integer) As Boolean
 
 
        '// Miqueas
 
     '// Controlamos que pueda Pegar
 
 
        With UserList(userindex).Counters
 
 
                If .Seguimiento.Click > 0 Then
 
                        PuedoClickear = False
 
 
                        Exit Function
 
 
                End If
 
             
 
                PuedoClickear = True
 
             
 
                '// 28*40 = 1120 Mseg Entre golpe y golpe
             .Seguimiento.Poteo = setIntervals.PoteoU
             .Seguimiento.Click = setIntervals.PoteoClick
     End With
 
 
End Function
 Public Function PuedoPotear(ByVal userindex As Integer) As Boolean
 
 
        '// Miqueas
 
     '// Controlamos que pueda Pegar
 
 
        With UserList(userindex).Counters
 
 
                If .Seguimiento.Poteo > 0 Then
 
                        PuedoPotear = False
 
 
                        Exit Function
 
 
                End If
 
             
 
                PuedoPotear = True
 
             
 
                '// 28*40 = 1120 Mseg Entre golpe y golpe
 
             .Seguimiento.Poteo = setIntervals.PoteoU
 
 
     End With
 
 
End Function
Public Function PuedoPegar(ByVal userindex As Integer) As Boolean
 
 
        '// Miqueas
 
     '// Controlamos que pueda Pegar
 
 
        With UserList(userindex).Counters
 
 
                If .Seguimiento.Golpe > 0 Then
 
                        PuedoPegar = False
 
 
                        Exit Function
 
 
                End If
 
             
 
                PuedoPegar = True
 
             
 
                '// 28*40 = 1120 Mseg Entre golpe y golpe
 
             .Seguimiento.Golpe = setIntervals.Golpe
             .Seguimiento.Flechas = 17
 
 
     End With
 
 
End Function
Private Sub BanAntiCheat(ByVal userindex As String)
 
 
        '***************************************************
 
     '// Autor: Miqueas
 
        '// 23/11/13
 
     '// No implementado
 
        '// ¿Hace falta una explicacion de lo que hace ?
 
     '// Bueno si, Banea al usuario, Bane codigo original funcion de baneo x ip
 
        '***************************************************
 
 
     Dim tUser    As Integer
 
     Dim cantPenas As Byte
 
 
 
     Const Reason  As String = "Uso de programas externos"
 
 
     tUser = userindex
 
 
     With UserList(tUser)
 
 
 
             '// Msj para escracharlo
 
                'Call SendData(SendTarget.ToAll, 0, "Sistema de AntiCheat> " & " ha baneado a " & .name & ": BAN POR " & LCase$(Reason) & "." & FONTTYPE_SERVER))
 
       
 
                '// Ponemos el flag de ban a 1
 
             .flags.Ban = 1
 
     
 
             '// Ponemos el flag de ban a 1
 
                Call WriteVar(CharPath & .Name & ".chr", "FLAGS", "Ban", "1")
 
           
 
                '// Ponemos la pena
 
             cantPenas = val(GetVar(CharPath & .Name & ".chr", "PENAS", "Cant"))
 
         
 
             '// Sumamos la pena
 
                Call WriteVar(CharPath & .Name & ".chr", "PENAS", "Cant", cantPenas + 1)
 
           
 
                '// Aplicamos por que se lo Baneo
 
             Call WriteVar(CharPath & .Name & ".chr", "PENAS", "P" & cantPenas + 1, "By - Anti Cheat" & ": BAN POR " & LCase$(Reason) & " " & Date$ & " " & Time$)
 
     
 
             Call CloseSocket(tUser)
 
         
 
     End With
 
 
End Sub
Public Sub InicializarContadores(userindex As Integer)
    UserList(userindex).Counters.TransportePremium = 0
    UserList(userindex).Counters.TransporteCastillos(31) = 0
    UserList(userindex).Counters.TransporteCastillos(32) = 0
    UserList(userindex).Counters.TransporteCastillos(33) = 0
    UserList(userindex).Counters.TransporteCastillos(34) = 0
    UserList(userindex).Counters.TransporteCastillos(35) = 0
    UserList(userindex).Counters.TimeComandos = 0
    UserList(userindex).Counters.InmoManopla = 0
    UserList(userindex).Counters.Saliendo = False
    UserList(userindex).Counters.Salir = 0
    UserList(userindex).Counters.IdleCount = 0
End Sub
