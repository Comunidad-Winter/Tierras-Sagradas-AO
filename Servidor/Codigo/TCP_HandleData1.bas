Attribute VB_Name = "TCP_HandleData1"
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

Public Sub HandleData_1(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim loopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim iStr As String
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

    Select Case UCase$(Left$(rData, 1))
        Case "X"        ' >>> Sistema Consultas
            rData = Right$(rData, Len(rData) - 1)
            Dim Usuario As Integer
            Dim Texto As String
            Usuario = NameIndex(ReadField(1, rData, Asc("*")))
            Texto = ReadField(2, rData, Asc("*"))
            If Usuario <= 0 Then Exit Sub
            UserList(Usuario).flags.ConsultaEnviada = False
            UserList(Usuario).flags.NumeroConsulta = 0
            SendData SendTarget.toindex, Usuario, 0, "||190"
            Call SendData(SendTarget.toindex, Usuario, 0, "RESPUES" & Texto & "*" & UserList(userindex).Name)
        Exit Sub
       Case "#"       ' >>> Sistema Consultas
            Debug.Print "Me llego SOS"
            rData = Right$(rData, Len(rData) - 1)
            Dim TipoConsulta As Byte
            Dim rDatax As String
            TipoConsulta = ReadField(1, rData, Asc(","))
            rDatax = ReadField(2, rData, Asc(","))
   
            If UserList(userindex).flags.Silenciado = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||191")
                Exit Sub
            End If
            
            If UserList(userindex).flags.ConsultaEnviada = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||192")
                Exit Sub
            End If
       
            If TipoConsulta = 0 Then
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||193")
                    MensajesNumber = MensajesNumber + 1
                    MensajesSOS(MensajesNumber).Tipo = "Consulta"
                    MensajesSOS(MensajesNumber).Autor = UserList(userindex).Name
                    MensajesSOS(MensajesNumber).Contenido = rDatax
                    UserList(userindex).flags.ConsultaEnviada = True
                    UserList(userindex).flags.NumeroConsulta = MensajesNumber
                    Exit Sub
            ElseIf TipoConsulta = 2 Then
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||193")
                    MensajesNumber = MensajesNumber + 1
                    MensajesSOS(MensajesNumber).Tipo = "Bug"
                    MensajesSOS(MensajesNumber).Autor = UserList(userindex).Name
                    MensajesSOS(MensajesNumber).Contenido = rDatax
                    UserList(userindex).flags.ConsultaEnviada = True
                    UserList(userindex).flags.NumeroConsulta = MensajesNumber
                    Exit Sub
            ElseIf TipoConsulta = 3 Then
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||193")
                    MensajesNumber = MensajesNumber + 1
                    MensajesSOS(MensajesNumber).Tipo = "Ayuda"
                    MensajesSOS(MensajesNumber).Autor = UserList(userindex).Name
                    MensajesSOS(MensajesNumber).Contenido = rDatax
                    UserList(userindex).flags.ConsultaEnviada = True
                    UserList(userindex).flags.NumeroConsulta = MensajesNumber
                    Exit Sub
                End If
            Exit Sub
        Case ";" 'Hablar
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            
        If UserList(userindex).flags.EspectadorArena1 = 1 Or UserList(userindex).flags.EspectadorArena2 = 1 Or UserList(userindex).flags.EspectadorArena3 = 1 Or UserList(userindex).flags.EspectadorArena4 = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||194")
            Exit Sub
        End If
            
        If rData = ":|" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE1 & "," & 130)
                Exit Sub
            ElseIf rData = ":S" Or rData = ":s" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE2 & "," & 130)
                Exit Sub
            ElseIf rData = ":(" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE3 & "," & 130)
                Exit Sub
            ElseIf rData = ";)" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE4 & "," & 130)
                Exit Sub
            ElseIf rData = ":$" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE5 & "," & 130)
                Exit Sub
            ElseIf rData = ">.>" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE6 & "," & 130)
                Exit Sub
            ElseIf rData = "?" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE7 & "," & 130)
                Exit Sub
            ElseIf rData = "!" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE8 & "," & 130)
            Exit Sub
            ElseIf rData = "..." Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE9 & "," & 130)
            Exit Sub
            ElseIf rData = "¬¬" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE10 & "," & 130)
            Exit Sub
            ElseIf rData = ":@" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE11 & "," & 130)
            Exit Sub
            ElseIf rData = ":/" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE12 & "," & 130)
            Exit Sub
            ElseIf rData = ":3" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE13 & "," & 130)
            Exit Sub
            ElseIf rData = "^^" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE14 & "," & 130)
            Exit Sub
            ElseIf rData = ":D" Or rData = ":d" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE15 & "," & 130)
            Exit Sub
            ElseIf rData = ":P" Or rData = ":p" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE16 & "," & 130)
            Exit Sub
            ElseIf rData = ":O" Or rData = ":o" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE17 & "," & 130)
            Exit Sub
            ElseIf rData = "xD" Or rData = "xd" Or rData = "XD" Or rData = "Xd" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE18 & "," & 130)
            Exit Sub
            ElseIf rData = ":'(" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE19 & "," & 130)
            Exit Sub
            ElseIf rData = ":)" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFE" & UserList(userindex).Char.CharIndex & "," & FXE20 & "," & 130)
            Exit Sub
        End If
            
            '[Consejeros]
            If UserList(userindex).flags.Privilegios >= PlayerType.UserSupport Then
                Call LogGMss(UserList(userindex).Name, "Dijo: " & rData, True)
            End If
            
            ind = UserList(userindex).Char.CharIndex
        
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.toindex, userindex, 0, "||195")
                End If
            End If
            
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, userindex, UserList(userindex).Pos.Map, "N|12632256°" & rData & "°" & CStr(ind))
            Else
            If UserList(userindex).flags.Privilegios > User Then
                  Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbYellow & "°" & rData & "°" & CStr(ind))
                If Not rData = vbNullString Then
                End If
            ElseIf UserList(userindex).flags.Privilegios = User Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & rData & "°" & CStr(ind))
                If Not rData = vbNullString Then
                End If
                End If
                End If
            Exit Sub
        Case "-" 'Gritar
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||6")
                    Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            '[Consejeros]
            If UserList(userindex).flags.Privilegios >= PlayerType.UserSupport Then
                Call LogGMss(UserList(userindex).Name, "Grito: " & rData, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.toindex, userindex, 0, "||195")
                End If
            End If
    
    
            ind = UserList(userindex).Char.CharIndex
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbRed & "°" & rData & "°" & str(ind))
            Exit Sub
            
            
       Case "\" 'Mensaje privado
            rData = Right$(rData, Len(rData) - 1)
            tName = ReadField(1, rData, Asc("@"))

            tIndex = NameIndex(tName)
            
            If tIndex = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||196")
             Exit Sub
            End If
            
            If tIndex <> 0 Then
                If Len(rData) <> Len(tName) Then
                    tMessage = ReadField(2, rData, Asc("@"))
                Else
                    tMessage = " "
                End If
                
                ind = UserList(userindex).Char.CharIndex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                
            If tMessage = "" Or tMessage = " " Then Exit Sub
                
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If UserList(tIndex).flags.Privilegios > PlayerType.EventManager And UserList(tIndex).flags.DeseoRecibirMSJ = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||196")
                Exit Sub
            End If
            
                If UserList(tIndex).flags.DeseoRecibirMSJ = 0 And UserList(userindex).flags.Privilegios = PlayerType.User Then
                           Call SendData(SendTarget.toindex, userindex, 0, "||197")
                    Exit Sub
                End If
            
            If UserList(userindex).flags.Muerto = 1 And (UserList(userindex).Pos.Map <> UserList(tIndex).Pos.Map) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||6")
                Exit Sub
            End If
                    
                Dim car As String
                For i = 1 To Len(rData)
                    car = mid$(rData, i, 1)
            
                    If car = "~" Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||198")
                        Exit Sub
                    End If
                Next i
                     
                '[Consejeros]
                If UserList(userindex).flags.Privilegios >= PlayerType.UserSupport Then
                    Call LogGMss(UserList(userindex).Name, "Le dijo a '" & UserList(tIndex).Name & "' " & tMessage, True)
                End If
               
               If VerPrivados = True Then
                Call SendData(SendTarget.ToAdmins, 0, 0, "||199@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@" & tMessage)
               End If
               
                Call SendData(SendTarget.toindex, userindex, UserList(userindex).Pos.Map, "P|" & "Le dijiste a " & UserList(tIndex).Name & ": " & tMessage & FONTTYPE_ROJOC)
                
                If UserList(userindex).flags.Privilegios > PlayerType.User Then
                    Call SendData(SendTarget.toindex, tIndex, UserList(userindex).Pos.Map, "P|(GM) " & UserList(userindex).Name & " te dijo: " & tMessage & FONTTYPE_AMARILLON)
                Else
                    Call SendData(SendTarget.toindex, tIndex, UserList(userindex).Pos.Map, "P|" & UserList(userindex).Name & " te dijo: " & tMessage & FONTTYPE_AMARILLON)
                End If

                Exit Sub
            End If
        Exit Sub
        
        Case "M" 'Moverse
            rData = Right$(rData, Len(rData) - 1)
            
            If UserList(userindex).flags.Stopped Then Exit Sub
            
            If UserList(userindex).flags.EspectadorArena1 = 1 Or UserList(userindex).flags.EspectadorArena2 = 1 Or UserList(userindex).flags.EspectadorArena3 = 1 Or UserList(userindex).flags.EspectadorArena4 = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||200")
                Exit Sub
            End If
            
            If TModalidad <> "0" Then
                If UserList(userindex).flags.EnTorneo = 1 And UCase$(TModalidad) = "DM" And CuentaRegresiva > 0 And (UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||161")
                    Exit Sub
                End If
            
                If UserList(userindex).flags.EnTorneo = 1 And (TModalidad = "1" Or TModalidad = "2" Or TModalidad = "3" Or TModalidad = "4") And (UsuarioPelea(1) = userindex Or UsuarioPelea(2) = userindex Or UsuarioPelea(3) = userindex Or UsuarioPelea(4) = userindex Or UsuarioPelea(5) = userindex Or UsuarioPelea(6) = userindex Or UsuarioPelea(7) = userindex Or UsuarioPelea(8) = userindex) And CuentaRegresiva > 0 And (UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||161")
                    Exit Sub
                End If
            
            For i = 1 To 8
                If userindex = UsuarioPelea(i) And CuentaRegresiva > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||161")
                    Exit Sub
                End If
            Next i
        End If
           
            'salida parche
            If UserList(userindex).Counters.Saliendo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||201")
                UserList(userindex).Counters.Saliendo = False
                UserList(userindex).Counters.Salir = 0
            End If
            
            If UserList(userindex).Counters.TransporteCastillos(31) > 0 Or UserList(userindex).Counters.TransporteCastillos(32) > 0 Or UserList(userindex).Counters.TransporteCastillos(33) > 0 Or UserList(userindex).Counters.TransporteCastillos(34) > 0 Or UserList(userindex).Counters.TransporteCastillos(35) > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||202")
                UserList(userindex).Counters.TransporteCastillos(31) = 0
                UserList(userindex).Counters.TransporteCastillos(32) = 0
                UserList(userindex).Counters.TransporteCastillos(33) = 0
                UserList(userindex).Counters.TransporteCastillos(34) = 0
                UserList(userindex).Counters.TransporteCastillos(35) = 0
            ElseIf UserList(userindex).Counters.TransportePremium > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||203")
                UserList(userindex).Counters.TransportePremium = 0
            End If
            
            If UserList(userindex).Counters.SegundosParaRevivir > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||204")
                UserList(userindex).Counters.SegundosParaRevivir = 0
            End If
            
          If UserList(userindex).flags.Paralizado = 0 Then
                If Not UserList(userindex).flags.Meditando Then
                    Call MoveUserChar(userindex, val(rData))
                ElseIf UserList(userindex).flags.Meditando Then
                  UserList(userindex).flags.Meditando = False
                  Call SendData(toindex, userindex, 0, "MEDOK")
                  Call SendData(toindex, userindex, 0, "||205")
                  UserList(userindex).Char.FX = 0
                  UserList(userindex).Char.loops = 0
                  Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
                  Call MoveUserChar(userindex, val(rData))
                End If
            End If
            
                'Sumamos los movimientos que hace, si en 40 minutos no hace ningún mov lo rajamos a la mierda por down
                If UserList(userindex).flags.AntiAFK < 10 Then
                    UserList(userindex).flags.AntiAFK = UserList(userindex).flags.AntiAFK + 1
                End If
            
            If TesoroContando = True And UserList(userindex).flags.Desenterrando = 1 Then
                TesoroContando = False
                UserList(userindex).flags.Desenterrando = 0
                TiempoTesoro = 30
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rData)
    
        Case "ACTPT"
            Call EnviarPuntos(userindex)
        Exit Sub
    
        Case "RPU" 'Pedido de actualizacion de la posicion
            Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
            Exit Sub
        Case "AT"
            If (Mod_AntiCheat.PuedoPegar(userindex) = False) Then Exit Sub
            
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||7")
                Exit Sub
            End If
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||206")
                     Exit Sub
                    End If
                
                Call UsuarioAtaca(userindex)
                
                'piedra libre para todos los compas!
                If UserList(userindex).flags.Oculto > 0 And UserList(userindex).flags.AdminInvisible = 0 Then
                    UserList(userindex).flags.Oculto = 0
                    If UserList(userindex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                        Call SendData(SendTarget.toindex, userindex, 0, "||195")
                    End If
                End If
                
             End If
            Exit Sub
        Case "AG"
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||8")
                    Exit Sub
            End If
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.UserSupport And Not UserList(userindex).flags.EsRolesMaster Then Exit Sub
                
            Call GetObj(userindex)
        Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "||207")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEGON")
                UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            End If
            Exit Sub
        Case "ACTUALIZAR"
            Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
            Exit Sub
        Case "TOINFO"
            tStr = SendTorneoList(userindex)
            Call SendData(SendTarget.toindex, userindex, 0, "LTR" & SendTorneoList(userindex))
        Exit Sub
        
        Case "SUBLVL1"
        If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader") <> UserList(userindex).Name Then
            Call SendData(SendTarget.toindex, userindex, 0, "||208")
        Exit Sub
        End If
    
        If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 7 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||209")
        Exit Sub
        End If
        
    Dim NivelClan As Byte
    Dim PuntosClan As Integer
    Dim SiguienteNivel As Integer
    
    NivelClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan")
    PuntosClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan")
    SiguienteNivel = GetVar(IniPath & "Configuracion.ini", "NIVELCLAN", "2")
    
    If NivelClan = 1 Then
      If PuntosClan >= SiguienteNivel Then
            puntixasd = PuntosClan - SiguienteNivel
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan", puntixasd)
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "2")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@2")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@2@8")
            Exit Sub
      Else
            Call SendData(SendTarget.toindex, userindex, 0, "||212@" & SiguienteNivel)
      Exit Sub
      End If
    End If
  Exit Sub
  
  Case "SUBLVL2"
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader") <> UserList(userindex).Name Then
    Call SendData(SendTarget.toindex, userindex, 0, "||208")
    Exit Sub
    End If
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 7 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||209")
    Exit Sub
    End If
    
    NivelClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan")
    PuntosClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan")
    SiguienteNivel = GetVar(IniPath & "Configuracion.ini", "NIVELCLAN", "3")
    
    If NivelClan = 2 Then
      If PuntosClan >= SiguienteNivel Then
            puntixasd = PuntosClan - SiguienteNivel
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan", puntixasd)
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "3")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@3")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@3@12")
            Exit Sub
      Else
            Call SendData(SendTarget.toindex, userindex, 0, "||212@" & SiguienteNivel)
      Exit Sub
      End If
    End If
  Exit Sub
  
  Case "SUBLVL3"
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader") <> UserList(userindex).Name Then
    Call SendData(SendTarget.toindex, userindex, 0, "||208")
    Exit Sub
    End If
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 7 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||209")
    Exit Sub
    End If
    
    NivelClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan")
    PuntosClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan")
    SiguienteNivel = GetVar(IniPath & "Configuracion.ini", "NIVELCLAN", "4")
    
    If NivelClan = 3 Then
      If PuntosClan >= SiguienteNivel Then
            puntixasd = PuntosClan - SiguienteNivel
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan", puntixasd)
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "4")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@4")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@4@16")
            Exit Sub
      Else
            Call SendData(SendTarget.toindex, userindex, 0, "||212@" & SiguienteNivel)
      Exit Sub
      End If
    End If
  Exit Sub
  
  Case "SUBLVL4"
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader") <> UserList(userindex).Name Then
    Call SendData(SendTarget.toindex, userindex, 0, "||208")
    Exit Sub
    End If
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 7 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||209")
    Exit Sub
    End If
    
    NivelClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan")
    PuntosClan = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan")
    SiguienteNivel = GetVar(IniPath & "Configuracion.ini", "NIVELCLAN", "5")
    
    If NivelClan = 4 Then
      If PuntosClan >= SiguienteNivel Then
            puntixasd = PuntosClan - SiguienteNivel
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan", puntixasd)
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "5")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@5")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@5@20")
            Exit Sub
      Else
            Call SendData(SendTarget.toindex, userindex, 0, "||212@" & SiguienteNivel)
      Exit Sub
      End If
    End If
  Exit Sub
  
  Case "SUBLVL5"
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader") <> UserList(userindex).Name Then
        Call SendData(SendTarget.toindex, userindex, 0, "||208")
    Exit Sub
    End If
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 7 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||209")
    Exit Sub
    End If
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 5 Then
      If TieneObjetos(1481, 1, userindex) = True Then
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "6")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@6")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@6@24")
            Call QuitarObjetos(1481, 1, userindex)
            Exit Sub
      Else
            Call SendData(SendTarget.toindex, userindex, 0, "||213")
      Exit Sub
      End If
    End If
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 6 Then
      If TieneObjetos(1482, 1, userindex) = True Then
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan", "7")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||210@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@7")
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||211@7@28")
            Call QuitarObjetos(1482, 1, userindex)
            Exit Sub
      Else
            Call SendData(SendTarget.toindex, userindex, 0, "||214")
      Exit Sub
      End If
    End If
    
  Exit Sub
  
    Case "DCANJE"
            tStr = UserList(userindex).Stats.PuntosDonacion & "," & UBound(DonationList) & ","
            For i = 1 To UBound(DonationList)
                tStr = tStr & DonationList(i).ObjName & ","
            Next
            
            Call SendData(SendTarget.toindex, userindex, 0, "DRM" & tStr)
    Exit Sub
  
        Case "CCANJE"
        Dim Premios As Integer, SX As String
            SX = "PRM" & UBound(PremiosList) & ","
             
            For Premios = 1 To UBound(PremiosList)
                SX = SX & PremiosList(Premios).ObjName & ","
            Next Premios
             
            Call SendData(SendTarget.toindex, userindex, 0, SX & UserList(userindex).Stats.PuntosTorneo)
            Call SendData(SendTarget.toindex, userindex, 0, "INF" & PremiosList(val(rData)).ObjRequiere & "," & PremiosList(val(rData)).ObjMaxAt & "," & PremiosList(val(rData)).ObjMinAt & "," & PremiosList(val(rData)).ObjMaxdef & "," & PremiosList(val(rData)).ObjMindef & "," & PremiosList(val(rData)).ObjMaxAtMag & "," & PremiosList(val(rData)).ObjMinAtMag & "," & PremiosList(val(rData)).ObjMaxDefMag & "," & PremiosList(val(rData)).ObjMinDefMag & "," & PremiosList(val(rData)).ObjDescripcion)
        Exit Sub
        
    Case "GLINFO"
        Call LoadGuildsClanes
        Dim GI As Integer
            GI = UserList(userindex).GuildIndex
            
            If GI <= 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "GL" & SendGuildsList(userindex))
            Exit Sub
            End If
            
            If UserList(userindex).GuildIndex >= 1 Then
              UserInfo = SendGuildUserInfo(userindex)
              If UserInfo <> vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "IREDAEK" & UserInfo)
                Call SendData(SendTarget.toindex, userindex, 0, "IRFORTA" & Fortaleza)
              Exit Sub
              End If
            End If
            
            tStr = SendGuildLeaderInfo(userindex)
            If tStr = vbNullString And UserInfo = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "GL" & SendGuildsList(userindex))
            Else
                If m_EsGuildSubLeader1(UserList(userindex).Name, GI) Or m_EsGuildSubLeader2(UserList(userindex).Name, GI) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "IREDAEL" & tStr)
                    Call SendData(SendTarget.toindex, userindex, 0, "IRFORTA" & Fortaleza)
                ElseIf m_EsGuildLeader(UserList(userindex).Name, GI) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "IREDAEL" & tStr)
                    Call SendData(SendTarget.toindex, userindex, 0, "IRFORTA" & Fortaleza)
                End If
            End If
           Exit Sub
        Case "ATRI"
            Call EnviarAtrib(userindex)
            Exit Sub
        Case "KEST"
            Call InfoQuest(userindex)
        Exit Sub
        Case "FEST" 'Mini estadisticas :)
            Call EnviarMiniEstadisticas(userindex)
        Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.toindex, userindex, 0, "FINCOMOK")
        Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "INIBOV"
            Call SendUserGLD(userindex)
            Call IniciarDeposito(userindex)
            Exit Sub
        Case "{IXION"
            Call SendData(SendTarget.toindex, userindex, 0, "MISI" & UserList(userindex).Misiones.NumeroMision & "," & UserList(userindex).Misiones.ConteoUser & "," & MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad)
        Exit Sub
        Case "{IZION"
            If UserList(userindex).Misiones.Completada = 1 Then Exit Sub
            VerificarMisionDiaria (userindex)
        Exit Sub
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.toindex, userindex, 0, "FINBANOK")
            Exit Sub
        Case "FINCBN"
            'User sale del modo BANCO
            UserList(userindex).flags.Comerciando = False
            UserList(userindex).flags.CuentaBancaria = ""
            Call SendData(SendTarget.toindex, userindex, 0, "FINCBNOK")
            Exit Sub
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        '[/Alejo]
    
    
    End Select
    
     Select Case UCase$(Left$(rData, 6))
     
    Case "CONSUL" 'Enviamos todos los s.o.s de los usuarios al cliente.
        rData = Right$(rData, Len(rData) - 6)
        
        Dim dataSOS As String
        dataSOS = MensajesNumber & ","
        
        For loopC = 1 To MensajesNumber
            dataSOS = dataSOS & MensajesSOS(loopC).Tipo & "-" & MensajesSOS(loopC).Autor & "-" & MensajesSOS(loopC).Contenido & ","
        Next loopC
        
        Call SendData(SendTarget.toindex, userindex, 0, "ZSOS" & dataSOS)
        
    Exit Sub
     
        Case "CABEZI"
          rData = Right$(rData, Len(rData) - 6)
          
          If UserList(userindex).Stats.GLD < 500 Then
            SendData SendTarget.toindex, userindex, 0, "||215@500"
           Exit Sub
          
          Else
          
            Dim MinEleccion As Integer
            Dim MaxEleccion As Integer
          
            Select Case UCase$(UserList(userindex).Genero)
            
                Case "HOMBRE"
            
                Select Case UCase$(UserList(userindex).Raza)
                
                    Case "HUMANO"
                        MaxEleccion = 30
                        MinEleccion = 1
                    
                    Case "ELFO"
                        MaxEleccion = 113
                        MinEleccion = 101
                                    
                    Case "ELFO OSCURO"
                        MaxEleccion = 209
                        MinEleccion = 202
                                    
                    Case "ENANO"
                        MaxEleccion = 305
                        MinEleccion = 301
                                    
                    Case "GNOMO"
                        MaxEleccion = 406
                        MinEleccion = 401
                                    
                    Case Else
                        MaxEleccion = 30
                        MinEleccion = 30
                                
                End Select
                    
            Case "MUJER"
               
                Select Case UCase$(UserList(userindex).Raza)
                
                    Case "HUMANO"
                        MaxEleccion = 76
                        MinEleccion = 70
                                    
                    Case "ELFO"
                        MaxEleccion = 176
                        MinEleccion = 170
                                    
                    Case "ELFO OSCURO"
                        MaxEleccion = 280
                        MinEleccion = 270
                                    
                    Case "GNOMO"
                        MaxEleccion = 474
                        MinEleccion = 470
                                    
                    Case "ENANO"
                        MaxEleccion = 373
                        MinEleccion = 370
                                
                    Case Else
                        MaxEleccion = 70
                        MinEleccion = 70
                End Select
            End Select
            
            
            If rData < MinEleccion Or rData > MaxEleccion Then
                SendData SendTarget.toindex, userindex, 0, "||216"
                Exit Sub
            End If
          
          
            UserList(userindex).OrigChar.Head = rData
            UserList(userindex).Char.Head = rData
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            
            SendData SendTarget.toindex, userindex, 0, "||217@" & rData
            
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 500
            
            SendUserGLD userindex
          
          End If
          
        Exit Sub
        
    Case "DAMINF" 'Enviamos información al formulario de estadisticas.
        rData = Right$(rData, Len(rData) - 6)
        Dim UserEstadisticas As Integer
        Dim AlineacionUser As Byte
        Dim JerarquiaUser As String
        Dim TextoDuelos As String
        Dim TextoParejas As String
        Dim TextoEventos As String
        Dim TotalesX As Integer
        
        tStr = ""
        UserEstadisticas = NameIndex(rData)
        
        If UserEstadisticas <= 0 Then
            SendData SendTarget.toindex, userindex, 0, "||196"
          Exit Sub
        End If
        
        If UserList(UserEstadisticas).flags.PJerarquia = 0 And UserList(UserEstadisticas).flags.SJerarquia = 0 And UserList(UserEstadisticas).flags.TJerarquia = 0 And UserList(UserEstadisticas).flags.CJerarquia = 0 Then
            JerarquiaUser = "None"
        ElseIf UserList(UserEstadisticas).flags.PJerarquia = 1 And UserList(UserEstadisticas).flags.SJerarquia = 0 And UserList(UserEstadisticas).flags.TJerarquia = 0 And UserList(UserEstadisticas).flags.CJerarquia = 0 Then
            JerarquiaUser = "1 de 4"
        ElseIf UserList(UserEstadisticas).flags.PJerarquia = 0 And UserList(UserEstadisticas).flags.SJerarquia = 1 And UserList(UserEstadisticas).flags.TJerarquia = 0 And UserList(UserEstadisticas).flags.CJerarquia = 0 Then
            JerarquiaUser = "2 de 4"
        ElseIf UserList(UserEstadisticas).flags.PJerarquia = 0 And UserList(UserEstadisticas).flags.SJerarquia = 0 And UserList(UserEstadisticas).flags.TJerarquia = 1 And UserList(UserEstadisticas).flags.CJerarquia = 0 Then
            JerarquiaUser = "3 de 4"
        ElseIf UserList(UserEstadisticas).flags.PJerarquia = 0 And UserList(UserEstadisticas).flags.SJerarquia = 0 And UserList(UserEstadisticas).flags.TJerarquia = 0 And UserList(UserEstadisticas).flags.CJerarquia = 1 Then
            JerarquiaUser = "4 de 4"
        End If
        
        If UserList(UserEstadisticas).StatusMith.EsStatus = 0 Then
            AlineacionUser = 0
        End If
        
        If UserList(UserEstadisticas).StatusMith.EsStatus = 1 Or UserList(UserEstadisticas).StatusMith.EsStatus = 3 Then
            AlineacionUser = 2
        End If
        
        If UserList(UserEstadisticas).StatusMith.EsStatus = 2 Or UserList(UserEstadisticas).StatusMith.EsStatus = 4 Then
            AlineacionUser = 1
        End If
        
        TotalesX = UserList(UserEstadisticas).Stats.DuelosGanados + UserList(UserEstadisticas).Stats.DuelosPerdidos
        If TotalesX > 0 Then
            TextoDuelos = "" & TotalesX & " jugados (" & Round(((val(UserList(UserEstadisticas).Stats.DuelosGanados) * 100) / TotalesX)) & "% de victorias)"
        Else
            TextoDuelos = "0 jugados (0% de victorias)"
        End If
        
        TotalesX = UserList(UserEstadisticas).Stats.ParejasGanadas + UserList(UserEstadisticas).Stats.ParejasPerdidas
        If TotalesX > 0 Then
            TextoParejas = "" & TotalesX & " jugadas (" & Round(((val(UserList(UserEstadisticas).Stats.ParejasGanadas) * 100) / TotalesX)) & "% de victorias)"
        Else
            TextoParejas = "0 jugadas (0% de victorias)"
        End If
        
        TextoEventos = "" & UserList(UserEstadisticas).Stats.TorneosParticipados & " (" & UserList(UserEstadisticas).Stats.MedOro & " ganados)"
        
        'Nombre
        tStr = tStr & UserList(UserEstadisticas).Name & ","
        tStr = tStr & UserList(UserEstadisticas).clase & ","
        tStr = tStr & UserList(UserEstadisticas).Raza & ","
        tStr = tStr & UserList(UserEstadisticas).Stats.ELV & ","
        tStr = tStr & UserList(UserEstadisticas).Stats.Exp & ","
        tStr = tStr & AlineacionUser & ","
        tStr = tStr & JerarquiaUser & ","
        tStr = tStr & UserList(UserEstadisticas).Stats.Reputacione & ","
                
        tStr = tStr & TextoDuelos & ","
        tStr = tStr & TextoParejas & ","
        tStr = tStr & UserList(UserEstadisticas).Stats.MaximasRondas & ","
        tStr = tStr & UserList(UserEstadisticas).Stats.MuertesUser & ","
        tStr = tStr & UserList(UserEstadisticas).Stats.UsuariosMatados & ","
        tStr = tStr & TextoEventos & ","
        tStr = tStr & UserList(UserEstadisticas).flags.CvcsGanados & ","
        tStr = tStr & UserList(UserEstadisticas).flags.QuestCompletadas
        
        'If UserList(UserEstadisticas).flags.Privilegios > PlayerType.User Then Exit Sub
        
        Call SendData(SendTarget.toindex, userindex, 0, "IFE" & tStr)
        
    Exit Sub
     
    Case "ENVFPZ"
        rData = Right$(rData, Len(rData) - 6)
        
        Call SendData(SendTarget.ToAdmins, 0, 0, "||218@" & UserList(userindex).Name & "@" & rData)
    Exit Sub
     
     Case "ADDPTS"
    rData = Right$(rData, Len(rData) - 6)
    Dim cantpts As Integer
    cantpts = val(ReadField(1, rData, 44))
    
    If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "NivelClan") = 5 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||209")
    Exit Sub
    End If
    
    If UserList(userindex).Stats.PuntosTorneo < cantpts Then
    Call SendData(SendTarget.toindex, userindex, 0, "||219")
    Exit Sub
    End If
    
    If cantpts < 0 Then Exit Sub
    
    Dim puntitosfinal As Integer
    puntitosfinal = GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan") + cantpts
    Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "PuntosClan", puntitosfinal)
    UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - cantpts
    Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||220@" & UserList(userindex).Name & "@" & cantpts)
  Exit Sub
  
  Case "ADDCON" 'Agrega contactos - lista de amigos
    rData = Right$(rData, Len(rData) - 6)
    Dim nombrepj As String
    nombrepj = ReadField(1, rData, 44)
    
    
    If (IsAdministrator(nombrepj)) Or (IsDevelopment(nombrepj)) Or (IsCoordination(nombrepj)) Or (IsTournamentManager(nombrepj)) Or (IsEventManager(nombrepj)) Or (IsUserSupport(nombrepj)) Then
        Call SendData(SendTarget.toindex, userindex, 0, "ERONo podes agregar GM's.")
        Exit Sub
    End If
    
    If nombrepj = UCase$(UserList(userindex).Name) Then
        Call SendData(SendTarget.toindex, userindex, 0, "ERONo puedes agregarte a ti mismo.")
    Exit Sub
    End If
    
        Dim forsitoh As Integer
        For forsitoh = 1 To 20
            If UCase$(UserList(userindex).flags.NombreAmigo(forsitoh)) = UCase$(nombrepj) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "EROEl usuario ya está en tu lista de amigos.")
                Exit Sub
            End If
        Next forsitoh
        
        Dim SumamoUno As Byte
        SumamoUno = 0
    
        For forsitoh = 1 To 20
            If FileExist(CharPath & nombrepj & ".chr", vbNormal) Then
                If UserList(userindex).flags.NombreAmigo(forsitoh) = "(Nadie)" Or UserList(userindex).flags.NombreAmigo(forsitoh) = "(NADIE)" Then
                  UserList(userindex).flags.NombreAmigo(forsitoh) = UCase$(nombrepj)
                    tStr = SendFriendList(userindex)
                    Call SendData(SendTarget.toindex, userindex, 0, "LDM" & SendFriendList(userindex))
                    
                    If NameIndex(nombrepj) > 0 Then Call SendData(SendTarget.toindex, NameIndex(nombrepj), 0, "||221@" & UserList(userindex).Name)
                    Exit Sub
                Else
                    SumamoUno = SumamoUno + 1
                    
                    If SumamoUno = 20 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "ERRLista de amigos llena, solo puedes agregar 20.")
                     Exit Sub
                    End If
                    
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "EROEl personaje no existe.")
                Exit Sub
            End If
        Next forsitoh
    
    Exit Sub
    
    Case "DYDTRA" 'Transferir item con drag & drop
    rData = Right$(rData, Len(rData) - 6)
    Dim PosicionUserX As Byte, PosicionUserY As Byte, NombreUser As String, ObjetoIndex As Integer, ObjetoAmount As Integer
    PosicionUserX = ReadField(1, rData, 44)
    PosicionUserY = ReadField(2, rData, 44)
    NombreUser = ReadField(3, rData, 44)
    ObjetoIndex = ReadField(4, rData, 44)
    ObjetoAmount = ReadField(5, rData, 44)
    
    tIndex = NameIndex(NombreUser)
    ObjetoIndex = UserList(userindex).Invent.Object(ObjetoIndex).ObjIndex
    
        If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrator Then
            Call LogGM(UserList(userindex).Name, "Transferencias: " & UserList(userindex).Name & " quiso transferir " & ObjData(ObjetoIndex).Name & " - " & ObjetoAmount, False)
            Exit Sub
        End If
    
        If tIndex <= 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||209")
            Exit Sub
        End If
        
        If UserList(tIndex).Pos.X <> PosicionUserX Or UserList(tIndex).Pos.Y <> PosicionUserY Then
                Call SendData(SendTarget.toindex, userindex, 0, "||222")
            Exit Sub
        End If
        
        If Not TieneObjetos(ObjetoIndex, ObjetoAmount, userindex) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||219")
            Exit Sub
        End If
        
        If ObjetoAmount < 0 Then Exit Sub
        
        If ObjData(ObjetoIndex).Intransferible = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||223")
            Exit Sub
        End If
        
        If tIndex = userindex Then Exit Sub
        
        Dim ObjetoTransferido As Obj
        ObjetoTransferido.ObjIndex = ObjetoIndex
        ObjetoTransferido.Amount = ObjetoAmount

        If Not MeterItemEnInventario(tIndex, ObjetoTransferido) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||224")
         Exit Sub
        End If

        Call SendData(SendTarget.toindex, tIndex, 0, "||225@" & UserList(userindex).Name & "@" & ObjetoTransferido.Amount & "@" & ObjData(ObjetoTransferido.ObjIndex).Name)
        Call QuitarObjetos(ObjetoIndex, ObjetoAmount, userindex)
        Call LogTransferencias("" & UserList(userindex).Name & " transfirio " & ObjetoTransferido.Amount & " - " & ObjData(ObjetoTransferido.ObjIndex).Name & " a " & UserList(tIndex).Name & "")

    Exit Sub
    
    Case "INCHAT" 'Iniciamos el chat
    rData = Right$(rData, Len(rData) - 6)
    Dim Contactito As String
    Contactito = UserList(userindex).flags.NombreAmigo(rData)
    
    If UCase$(Contactito) = "(NADIE)" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||226")
        Exit Sub
    End If
    
        Call SendData(SendTarget.toindex, userindex, 0, "ENCHAT" & Contactito)
    
    Exit Sub
    
    Case "KKCHAT" 'Iniciamos el chat
    rData = Right$(rData, Len(rData) - 6)
    Dim MensajitoChat As String
    Contactito = ReadField(1, rData, 44)
    MensajitoChat = ReadField(2, rData, 44)
    
    tIndex = NameIndex(Contactito)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||209")
        Exit Sub
    End If
    
    Dim revisionmaxima As Boolean, conteoAG As Byte
    conteoAG = 0
    
    For i = 1 To 20
        If UCase$(UserList(tIndex).flags.NombreAmigo(i)) = UCase$(UserList(userindex).Name) Then
           revisionmaxima = True
        Else
            conteoAG = conteoAG + 1
        End If
        
        If ((revisionmaxima = True) And (UserList(tIndex).flags.AmigoDesadmitido(i) = 1)) Or conteoAG = 20 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||227")
            Exit Sub
        End If
    Next i
    
        If VerPrivados = True Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||228@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@" & MensajitoChat)
        End If
    
        Call SendData(SendTarget.toindex, tIndex, 0, "||229@" & UCase$(UserList(userindex).Name) & "@" & MensajitoChat)
        Call SendData(SendTarget.toindex, tIndex, 0, "LDCHAT" & UserList(userindex).Name & "," & MensajitoChat)
    Exit Sub
    
    Case "BORRAC" 'Borrar contacto - lista de amigos
    rData = Right$(rData, Len(rData) - 6)
    
    If UCase$(UserList(userindex).flags.NombreAmigo(rData)) <> "(NADIE)" Then
        UserList(userindex).flags.NombreAmigo(rData) = "(Nadie)"
    End If
    
        tStr = SendFriendList(userindex)
        Call SendData(SendTarget.toindex, userindex, 0, "LDM" & SendFriendList(userindex))
    
    Exit Sub
    
     Case "OFDIOZ"
     rData = Right$(rData, Len(rData) - 6)
        Dim CantAlmas As Long
        CantAlmas = ReadField(1, rData, 44)
     
      If UserList(userindex).flags.AlmasContenidas < CantAlmas Then
        Call SendData(SendTarget.toindex, userindex, 0, "ERONo tienes esa cantidad de almas.")
       Exit Sub
      End If
      
      If CantAlmas < 0 Then Exit Sub
      
      UserList(userindex).flags.AlmasContenidas = UserList(userindex).flags.AlmasContenidas - CantAlmas
      UserList(userindex).flags.AlmasOfrecidas = UserList(userindex).flags.AlmasOfrecidas + CantAlmas
      Call SendData(SendTarget.toindex, userindex, 0, "||230@" & CantAlmas & "@" & UserList(userindex).flags.SirvienteDeDios)
      Call LogAlmas("" & UserList(userindex).Name & " sacrifico: " & CantAlmas & "")
      
      If UserList(userindex).flags.SirvienteDeDios = "Mifrit" Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "PCF" & 77 & "," & 84 & "," & 51 & "," & 30)
      ElseIf UserList(userindex).flags.SirvienteDeDios = "Poseidon" Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "PCF" & 77 & "," & 49 & "," & 14 & "," & 30)
      ElseIf UserList(userindex).flags.SirvienteDeDios = "Tarraske" Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "PCF" & 77 & "," & 16 & "," & 51 & "," & 30)
      ElseIf UserList(userindex).flags.SirvienteDeDios = "Erebros" Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "PCF" & 77 & "," & 50 & "," & 87 & "," & 30)
      End If
      
    Dim ClaseUsuario As String
    ClaseUsuario = UCase$(UserList(userindex).clase)
      
    If UserList(userindex).flags.JerarquiaDios = 1 Then
        If UserList(userindex).flags.AlmasOfrecidas >= (AlmasNecesarias * UserList(userindex).flags.JerarquiaDios) Then
          Dim Entregar As Obj
          
            If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Bajos.dat", "" & ClaseUsuario & "", "Obj1")
            Else
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Altos.dat", "" & ClaseUsuario & "", "Obj1")
            End If
            
            Entregar.Amount = 1
            
              If TieneObjetos(Entregar.ObjIndex, 1, userindex) = False Then
                  If Not MeterItemEnInventario(userindex, Entregar) Then
                     Call SendData(SendTarget.toindex, userindex, 0, "||108")
                   Exit Sub
                  End If
                  
                  Call SendData(SendTarget.toindex, userindex, 0, "||231@Soldado@" & UserList(userindex).flags.SirvienteDeDios)
                  Call SendData(SendTarget.toindex, userindex, 0, "||232@1@" & ObjData(Entregar.ObjIndex).Name)
                  UserList(userindex).flags.JerarquiaDios = 2
              End If
        End If
    End If
    
    If UserList(userindex).flags.JerarquiaDios = 2 Then
        If UserList(userindex).flags.AlmasOfrecidas >= (AlmasNecesarias * UserList(userindex).flags.JerarquiaDios) Then
        
            If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Bajos.dat", "" & ClaseUsuario & "", "Obj2")
            Else
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Altos.dat", "" & ClaseUsuario & "", "Obj2")
            End If
            
            Entregar.Amount = 1
            
              If TieneObjetos(Entregar.ObjIndex, 1, userindex) = False Then
                  If Not MeterItemEnInventario(userindex, Entregar) Then
                     Call SendData(SendTarget.toindex, userindex, 0, "||108")
                   Exit Sub
                  End If
                  
                  Call SendData(SendTarget.toindex, userindex, 0, "||231@Guerrero@" & UserList(userindex).flags.SirvienteDeDios)
                  Call SendData(SendTarget.toindex, userindex, 0, "||232@1@" & ObjData(Entregar.ObjIndex).Name)
                  UserList(userindex).flags.JerarquiaDios = 3
              End If
        End If
    End If
      
    If UserList(userindex).flags.JerarquiaDios = 3 Then
        If UserList(userindex).flags.AlmasOfrecidas >= (AlmasNecesarias * UserList(userindex).flags.JerarquiaDios) Then
        
            If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Bajos.dat", "" & ClaseUsuario & "", "Obj3")
            Else
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Altos.dat", "" & ClaseUsuario & "", "Obj3")
            End If
            
            Entregar.Amount = 1
            
              If TieneObjetos(Entregar.ObjIndex, 1, userindex) = False Then
                  If Not MeterItemEnInventario(userindex, Entregar) Then
                     Call SendData(SendTarget.toindex, userindex, 0, "||108")
                   Exit Sub
                  End If
                  
                  Call SendData(SendTarget.toindex, userindex, 0, "||231@Caballero@" & UserList(userindex).flags.SirvienteDeDios)
                  Call SendData(SendTarget.toindex, userindex, 0, "||232@1@" & ObjData(Entregar.ObjIndex).Name)
                  UserList(userindex).flags.JerarquiaDios = 4
              End If
        End If
    End If
      
    If UserList(userindex).flags.JerarquiaDios = 4 Then
        If UserList(userindex).flags.AlmasOfrecidas >= (AlmasNecesarias * UserList(userindex).flags.JerarquiaDios) Then
        
            If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Bajos.dat", "" & ClaseUsuario & "", "Obj4")
            Else
              Entregar.ObjIndex = GetVar(App.Path & "\Dioses\" & UserList(userindex).flags.SirvienteDeDios & "\Altos.dat", "" & ClaseUsuario & "", "Obj4")
            End If
            
            Entregar.Amount = 1
            
              If TieneObjetos(Entregar.ObjIndex, 1, userindex) = False Then
                  If Not MeterItemEnInventario(userindex, Entregar) Then
                     Call SendData(SendTarget.toindex, userindex, 0, "||108")
                   Exit Sub
                  End If
                  
                  Call SendData(SendTarget.toindex, userindex, 0, "||231@Campeón@" & UserList(userindex).flags.SirvienteDeDios)
                  Call SendData(SendTarget.toindex, userindex, 0, "||232@1@" & ObjData(Entregar.ObjIndex).Name)
                  UserList(userindex).flags.JerarquiaDios = 5
              End If
        End If
    End If
      
     
     Exit Sub
     
        Case "DESPHE" 'Mover Hechizo de lugar

            rData = Right(rData, Len(rData) - 6)
            Call DesplazarHechizo(userindex, CInt(ReadField(1, rData, 44)), CInt(ReadField(2, rData, 44)))
        Exit Sub
        
        Case "DESCOD" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 6)
                Call modGuilds.ActualizarCodexYDesc(rData, UserList(userindex).GuildIndex)
        Exit Sub

End Select
    
    
    Select Case UCase$(Left$(rData, 2))
    '    Case "/Z"
    '        Dim Pos As WorldPos, Pos2 As WorldPos
    '        Dim O As Obj
    '
    '        For LoopC = 1 To 100
    '            Pos = UserList(UserIndex).Pos
    '            O.Amount = 1
    '            O.ObjIndex = iORO
    '            'Exit For
    '            Call TirarOro(100000, UserIndex)
    '            'Call Tilelibre(Pos, Pos2)
    '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
    '
    '            'Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
    '        Next LoopC
    '
    '        Exit Sub
        Case "TR" 'Tirar item por mouse
                If UserList(userindex).flags.Navegando = 1 Or _
                   UserList(userindex).flags.Muerto = 1 Or _
                   (UserList(userindex).flags.Privilegios = PlayerType.UserSupport And UserList(userindex).flags.Privilegios < PlayerType.Administrator) Then Exit Sub
                   '[Consejeros]
               
                rData = Right$(rData, Len(rData) - 2)
                Arg1 = ReadField(1, rData, 44)
                Arg2 = ReadField(2, rData, 44)
                Arg3 = ReadField(3, rData, 44)
                Arg4 = ReadField(4, rData, 44)

                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(userindex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(userindex, val(Arg1), val(Arg2), UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
                    Else
                        Exit Sub
                    End If
            Exit Sub
        Case "TI" 'Tirar item
                If UserList(userindex).flags.Navegando = 1 Or _
                   UserList(userindex).flags.Muerto = 1 Or _
                   (UserList(userindex).flags.Privilegios = PlayerType.UserSupport And Not UserList(userindex).flags.EsRolesMaster) Then Exit Sub
                   '[Consejeros]
                
                rData = Right$(rData, Len(rData) - 2)
                Arg1 = ReadField(1, rData, 44)
                Arg2 = ReadField(2, rData, 44)
                If val(Arg1) = FLAGORO Then
                    
                    Call TirarOro(val(Arg2), userindex)
                    
                    Call SendUserGLD(userindex)
                    Exit Sub
                Else
                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(userindex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(userindex, val(Arg1), val(Arg2), UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
                    Else
                        Exit Sub
                    End If
                End If
                Exit Sub
        Case "LH" ' Lanzar hechizo
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 2)
            UserList(userindex).flags.Hechizo = val(rData)
            Exit Sub
        Case "LC" 'Click izquierdo
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(userindex, UserList(userindex).Pos.Map, X, Y)
            Exit Sub
        Case "RC" 'Click derecho
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(userindex, UserList(userindex).Pos.Map, X, Y)
            Exit Sub
        Case "UK"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
    
            rData = Right$(rData, Len(rData) - 2)
            Select Case val(rData)
                Case Robar
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Robar)
                Case Magia
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(userindex).flags.Navegando = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 3 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||233")
                            UserList(userindex).flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    If UserList(userindex).flags.Oculto = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 2 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||234")
                            UserList(userindex).flags.UltimoMensaje = 2
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    Call DoOcultarse(userindex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 3))
    
        Case "SPÑ"
            rData = Right$(rData, Len(rData) - 3)
            Dim NumeroMejorado As Integer, Objeto As Obj, Requiere As Integer
            
            NumeroMejorado = GetVar(DatPath & "Mejorados.dat", "ITEMS", rData)
             
            Requiere = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "Requiere")
            Objeto.ObjIndex = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "NumObj")
            Objeto.Amount = 1
            
            If TieneObjetos(1448, 1, userindex) = False Then
                SendData SendTarget.toindex, userindex, 0, "||235"
            Exit Sub
            End If
            
            
            If TieneObjetos(Requiere, 1, userindex) = False Then
                SendData SendTarget.toindex, userindex, 0, "||236"
            Exit Sub
            ElseIf Not MeterItemEnInventario(userindex, Objeto) Then
                SendData SendTarget.toindex, userindex, 0, "||108"
            Else
                Call LogMedallas("" & UserList(userindex).Name & " mejoró el objeto: " & ObjData(Requiere).Name)
                SendData SendTarget.toindex, userindex, 0, "||237@" & ObjData(Requiere).Name
                Call QuitarObjetos(Requiere, 1, userindex)
                Call QuitarObjetos(1448, 1, userindex)
            Exit Sub
            End If
        
        Exit Sub
 

        Case "SPH"
    
            rData = Right$(rData, Len(rData) - 3)
            
            NumeroMejorado = GetVar(DatPath & "Mejorados.dat", "ITEMS", rData)
            
             
            Dim NombreMejorar As String, Ataque As String, Defensa As String, AtaqueMagico As String, DefensaMagica As String, DescripcionM As String
            NombreMejorar = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "Nombre")
            Ataque = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "AtaqueMinimo") & "/" & GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "AtaqueMaximo")
            Defensa = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "DefensaMinima") & "/" & GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "DefensaMaxima")
            AtaqueMagico = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "AtaqueMagicoMinimo") & "/" & GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "AtaqueMagicoMaximo")
            DefensaMagica = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "DefensaMagicaMinima") & "/" & GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "DefensaMagicaMaxima")
            DescripcionM = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "Descripcion")
            
            
            
            Objeto.ObjIndex = GetVar(DatPath & "Mejorados.dat", "ITEM" & NumeroMejorado, "NumObj")
            
            If val(NumeroMejorado) > 0 Then _
            SendData SendTarget.toindex, userindex, 0, "IMEJ" & NombreMejorar & "," & Ataque & "," & Defensa & "," & AtaqueMagico & "," & DefensaMagica & "," & DescripcionM & "," & ObjData(Objeto.ObjIndex).GrhIndex
            
        Exit Sub
       Case "ARE"
        rData = Right$(rData, Len(rData) - 3)
        
            If UserList(userindex).Stats.GLD < 40000 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||215@40.000")
             Exit Sub
            End If
            
            If UserList(userindex).flags.EspectadorArena1 = 1 Or UserList(userindex).flags.EspectadorArena2 = 1 Or UserList(userindex).flags.EspectadorArena3 = 1 Or UserList(userindex).flags.EspectadorArena4 = 1 Then Exit Sub
            
            
            If UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 118 Or MapInfo(UserList(userindex).Pos.Map).Pk = True Or UserList(userindex).Pos.Map = 100 Or UserList(userindex).Pos.Map = 99 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).EnCvc = True Or UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||239")
                Exit Sub
            End If
        
        Select Case rData
            Case 1
              If ArenaOcupada(1) = True And EspectadoresEnArena1 < 4 Then
                        'Seteo Variables
                        UserList(userindex).flags.MapaAnterior = UserList(userindex).Pos.Map
                        UserList(userindex).flags.XAnterior = UserList(userindex).Pos.X
                        UserList(userindex).flags.YAnterior = UserList(userindex).Pos.Y
                        UserList(userindex).flags.EspectadorArena1 = 1
                    
                    'Transporte
                    If MapData(101, 33, 34).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 33, 34, False)
                    ElseIf MapData(101, 34, 34).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 34, 34, False)
                    ElseIf MapData(101, 33, 35).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 33, 35, False)
                    ElseIf MapData(101, 34, 35).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 34, 35, False)
                    End If
                        
                    EspectadoresEnArena1 = EspectadoresEnArena1 + 1
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 40000
                    Call SendData(SendTarget.toindex, userindex, 0, "||240")
                
              Else
                Call SendData(SendTarget.toindex, userindex, 0, "||241")
                Exit Sub
              End If
              
            Case 2
              If ArenaOcupada(2) = True And EspectadoresEnArena2 < 4 Then
                        'Seteo Variables
                        UserList(userindex).flags.MapaAnterior = UserList(userindex).Pos.Map
                        UserList(userindex).flags.XAnterior = UserList(userindex).Pos.X
                        UserList(userindex).flags.YAnterior = UserList(userindex).Pos.Y
                        UserList(userindex).flags.EspectadorArena2 = 1
                    
                    'Transporte
                    If MapData(101, 33, 68).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 33, 68, False)
                    ElseIf MapData(101, 34, 68).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 34, 68, False)
                    ElseIf MapData(101, 33, 69).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 33, 69, False)
                    ElseIf MapData(101, 34, 69).userindex = 0 = 3 Then
                        Call WarpUserChar(userindex, 101, 34, 69, False)
                    End If
                        
                    EspectadoresEnArena2 = EspectadoresEnArena2 + 1
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 40000
                    Call SendData(SendTarget.toindex, userindex, 0, "||240")
                    
              Else
                Call SendData(SendTarget.toindex, userindex, 0, "||241")
                Exit Sub
              End If
              
            Case 3
              If ArenaOcupada(3) = True And EspectadoresEnArena3 < 4 Then
                        'Seteo Variables
                        UserList(userindex).flags.MapaAnterior = UserList(userindex).Pos.Map
                        UserList(userindex).flags.XAnterior = UserList(userindex).Pos.X
                        UserList(userindex).flags.YAnterior = UserList(userindex).Pos.Y
                        UserList(userindex).flags.EspectadorArena3 = 1
                    
                    'Transporte
                    If MapData(101, 69, 34).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 69, 34, False)
                    ElseIf MapData(101, 70, 34).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 70, 34, False)
                    ElseIf MapData(101, 69, 35).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 69, 35, False)
                    ElseIf MapData(101, 70, 35).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 70, 35, False)
                    End If
                        
                    EspectadoresEnArena3 = EspectadoresEnArena3 + 1
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 40000
                    Call SendData(SendTarget.toindex, userindex, 0, "||240")
                    
              Else
                Call SendData(SendTarget.toindex, userindex, 0, "||241")
                Exit Sub
              End If
              
            Case 4
              If ArenaOcupada(4) = True And EspectadoresEnArena4 < 4 Then
                        'Seteo Variables
                        UserList(userindex).flags.MapaAnterior = UserList(userindex).Pos.Map
                        UserList(userindex).flags.XAnterior = UserList(userindex).Pos.X
                        UserList(userindex).flags.YAnterior = UserList(userindex).Pos.Y
                        UserList(userindex).flags.EspectadorArena4 = 1
                    
                    'Transporte
                    If MapData(101, 69, 68).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 69, 68, False)
                    ElseIf MapData(101, 70, 68).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 70, 68, False)
                    ElseIf MapData(101, 69, 69).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 69, 69, False)
                    ElseIf MapData(101, 70, 69).userindex = 0 Then
                        Call WarpUserChar(userindex, 101, 70, 69, False)
                    End If
                        
                    EspectadoresEnArena4 = EspectadoresEnArena4 + 1
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 40000
                    Call SendData(SendTarget.toindex, userindex, 0, "||240")
                    
              Else
                Call SendData(SendTarget.toindex, userindex, 0, "||241")
                Exit Sub
              End If
          End Select
       Exit Sub
        Case "USA"
            rData = Right$(rData, Len(rData) - 3)
            
            If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                If UserList(userindex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            
            If ObjData(UserList(userindex).Invent.Object(val(rData)).ObjIndex).proyectil = 1 Then
                'nada
            Else
                If (Mod_AntiCheat.PuedoPotear(userindex) = False) Then Exit Sub
            End If
            
            Call UseInvItem(userindex, val(rData))
        Exit Sub
            
        Case "QSA"
            rData = Right$(rData, Len(rData) - 3)
            
            Dim NumOBJ As Integer
            Dim InvenVisible As String
                NumOBJ = ReadField(1, rData, 44)
                InvenVisible = ReadField(2, rData, 44)
                
            If UCase$(InvenVisible) = "FALSO" Then Call SendData(SendTarget.ToAdmins, 0, 0, "||242@" & UserList(userindex).Name): Exit Sub

            If val(NumOBJ) <= MAX_INVENTORY_SLOTS And val(NumOBJ) > 0 Then
                If UserList(userindex).Invent.Object(val(NumOBJ)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            
            If ObjData(UserList(userindex).Invent.Object(val(NumOBJ)).ObjIndex).proyectil = 1 Then
                'nada
            Else
                If (Mod_AntiCheat.PuedoClickear(userindex) = False) Then Exit Sub
            End If
            
            Call UseInvItem(userindex, val(NumOBJ))
        Exit Sub
            
        Case "TCM" 'CERRAR COMERCIO (CANCELAR)
            comCancelar userindex
            Exit Sub
        Case "UOR"
            rData = Right$(rData, Len(rData) - 3)
                If UserList(userindex).Stats.GLD < rData Then
                    comCancelar userindex
                    comMen userindex, "595"
                Exit Sub
                End If
                
            UserList(userindex).flags.OroQueOferto = rData
          Exit Sub
        Case "UOC" 'ENVIA OFERTA
            rData = Right$(rData, Len(rData) - 3)
            comMandoOferta userindex, rData
            Exit Sub
        Case "TDR" 'RESPUESTA AL COMERCIO
            rData = Right$(rData, Len(rData) - 3)
            comAceptaORechaza userindex, val(rData)
            Exit Sub
        Case "VHC" 'CHAT CON EL QUE COMERCIA
            rData = Right$(rData, Len(rData) - 3)
            comChat rData, userindex
            Exit Sub
            
    '####CORREOS####
        Case "CZM" 'ENVIA MSJ
            rData = Right$(rData, Len(rData) - 3)
            correoEnviarMensaje userindex, rData
        Exit Sub
        
        Case "CZC" 'LEE CORREO
            rData = Right$(rData, Len(rData) - 3)
            correoLeerMensaje userindex, rData
        Exit Sub
        
        Case "CZR" 'RETIRA OBJS
            rData = Right$(rData, Len(rData) - 3)
            correoRetirarItems userindex, rData
        Exit Sub
        
        Case "CZB" 'BORRA MENSAJE
            rData = Right$(rData, Len(rData) - 3)
            correoBorrarMensaje userindex, rData
        Exit Sub
        
    '####CORREOS####
    
            
    Case "CUC"
     rData = Right$(rData, Len(rData) - 3)
     Dim NumCasax As String
     Dim DueñoCasax As String
     Dim PrecioCasax As String
     NumCasax = ReadField(1, rData, 44)
   
    DueñoCasax = GetVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Dueño")
    PrecioCasax = GetVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Precio")
   
    If DueñoCasax <> "N/A" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||243")
    Exit Sub
    End If
   
    If UserList(userindex).Stats.GLD < PrecioCasax Then
    Call SendData(SendTarget.toindex, userindex, 0, "||215@" & PrecioCasax)
    Exit Sub
    End If
   
    Dim llavexx As Obj
    
    If NumCasax > 0 Then
        llavexx.ObjIndex = 1093 + NumCasax
        llavexx.Amount = 1
        
            If Not MeterItemEnInventario(userindex, llavexx) Then
               Call SendData(SendTarget.toindex, userindex, 0, "||108")
            Exit Sub
            End If
        
        Call WriteVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Dueño", UserList(userindex).Name)
        Call WriteVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Fecha", Date)
        
        Call SendData(SendTarget.ToAll, 0, 0, "||244@" & UserList(userindex).Name & "@" & NumCasax)
        
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - PrecioCasax
    End If
   
        Call SendUserGLD(userindex)
   
    Exit Sub
    
    Case "FWO"
     rData = Right$(rData, Len(rData) - 3)
        NumCasax = ReadField(1, rData, 44)
        
        Call SendData(SendTarget.toindex, userindex, 0, "GVN" & GetVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Dueño") & "," & GetVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Precio") & "," & GetVar(DatPath & "Casas.dat", "Casa" & NumCasax, "Fecha"))
    Exit Sub
    
    Case "CNM"
    rData = Right$(rData, Len(rData) - 3)
     Dim NickM As String
     NickM = ReadField(1, rData, 44)
    
     Call CambiarNickMascota(userindex, NickM)
    Exit Sub
            
   Case "IPX"
    rData = Right$(rData, Len(rData) - 3)
           
        If val(rData) > 0 And val(rData) < UBound(PremiosList) + 1 Then _
        Call SendData(SendTarget.toindex, userindex, 0, "INF" & PremiosList(val(rData)).ObjRequiere & "," & PremiosList(val(rData)).ObjMaxAt & "," & PremiosList(val(rData)).ObjMinAt & "," & PremiosList(val(rData)).ObjMaxdef & "," & PremiosList(val(rData)).ObjMindef & "," & PremiosList(val(rData)).ObjMaxAtMag & "," & PremiosList(val(rData)).ObjMinAtMag & "," & PremiosList(val(rData)).ObjMaxDefMag & "," & PremiosList(val(rData)).ObjMinDefMag & "," & PremiosList(rData).ObjDescripcion & "," & UserList(userindex).Stats.PuntosTorneo & "," & ObjData(PremiosList(rData).ObjIndexP).GrhIndex)
    Exit Sub
    
    Case "SPX"
        rData = Right$(rData, Len(rData) - 3)
        Arg1 = ReadField(1, rData, 44)
        Arg2 = ReadField(2, rData, 44)
        
        Dim Premio As Obj
           
            If val(Arg1) >= 0 And val(Arg1) < UBound(PremiosList) + 1 Then
               If val(Arg2) <= 0 And val(Arg2) > 10000 Then Exit Sub
        
                   Premio.Amount = val(Arg2)
                   Premio.ObjIndex = PremiosList(val(Arg1)).ObjIndexP
               
               End If
           
            'Si no tiene los puntos necesarios
            If UserList(userindex).Stats.PuntosTorneo < (PremiosList(val(Arg1)).ObjRequiere * val(Arg2)) Then
                   Call SendData(SendTarget.toindex, userindex, 0, "||245@" & val(Arg2) & "@" & ObjData(Premio.ObjIndex).Name)
            Exit Sub
            End If
            
            If Premio.Amount <= 0 Then Exit Sub
           
            'Si no tenemoss lugar lo tiramos al piso
            If Not MeterItemEnInventario(userindex, Premio) Then
               Call SendData(SendTarget.toindex, userindex, 0, "||108")
            Exit Sub
            End If
            
            If MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Tipo = 11 And MisionesDiarias(UserList(userindex).Misiones.NumeroMision).PTS <= PremiosList(val(Arg1)).ObjRequiere Then
                If UserList(userindex).Misiones.ConteoUser < MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad Then UserList(userindex).Misiones.ConteoUser = UserList(userindex).Misiones.ConteoUser + 1
            End If
           
            'Metemos en inventario
            Call LogCanjeos("(ITEMS TORNEO) " & UserList(userindex).Name & " canjeo: " & Premio.Amount & " - " & ObjData(Premio.ObjIndex).Name)
            Call SendData(SendTarget.toindex, userindex, 0, "||232@" & Premio.Amount & "@" & ObjData(Premio.ObjIndex).Name)
            Call UpdateUserInv(True, userindex, 0)
           
            'Restamos & actualizams
            UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - (PremiosList(val(Arg1)).ObjRequiere * Premio.Amount)
        Exit Sub
        
        Case "BOF" 'Bonificadores
            rData = Right$(rData, Len(rData) - 3)
            
            If UserList(userindex).Stats.ELV >= 20 And (rData = GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel1Opcion1") Or rData = GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel1Opcion2")) And UserList(userindex).Bon1 = "Ninguno/No elegido." Then
                UserList(userindex).Bon1 = rData
                
                If rData = "Aumenta en 5 puntos tu vida." Then
                    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 5
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                    SendUserHP (userindex)
                ElseIf rData = "Aumenta en 3 puntos tu vida." Then
                    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 3
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                    SendUserHP (userindex)
                End If
                
            ElseIf UserList(userindex).Stats.ELV >= 40 And (rData = GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion1") Or rData = GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel2Opcion2")) And UserList(userindex).Bon2 = "Ninguno/No elegido." Then
                UserList(userindex).Bon2 = rData
                
                If rData = "Aumenta en 5 puntos tu vida." Then
                    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 5
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                    SendUserHP (userindex)
                End If
            ElseIf UserList(userindex).Stats.ELV = 60 And (rData = GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel3Opcion1") Or rData = GetVar(DatPath & "ClassBonus.dat", "" & UCase$(UserList(userindex).clase) & "", "Nivel3Opcion2")) And UserList(userindex).Bon3 = "Ninguno/No elegido." Then
                UserList(userindex).Bon3 = rData
                
                If rData = "Aumenta en 5 puntos tu vida." Then
                    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 5
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                    SendUserHP (userindex)
                ElseIf rData = "Aumenta en 5 puntos la vida." Then
                    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 5
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                    SendUserHP (userindex)
                End If
                
            End If
        Exit Sub
        Case "CNS" ' Construye herreria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(rData)
            If X < 1 Then Exit Sub
            If ObjData(X).SkHerreria = 0 Then Exit Sub
            Call HerreroConstruirItem(userindex, X)
            Exit Sub
        Case "CNC" ' Construye carpinteria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(rData)
            If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
            Call CarpinteroConstruirItem(userindex, X)
            Exit Sub
        Case "WLC" 'Click izquierdo en modo trabajo
            rData = Right$(rData, Len(rData) - 3)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            Arg3 = ReadField(3, rData, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            X = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(userindex).flags.Muerto = 1 Or _
               UserList(userindex).flags.Meditando Or _
               Not InMapBounds(UserList(userindex).Pos.Map, X, Y) Then Exit Sub
            
            If Not InRangoVision(userindex, X, Y) Then
                Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                If (Mod_AntiCheat.PuedoFlechear(userindex) = False) Then Exit Sub

                DummyInt = 0

                If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.WeaponEqpSlot < 1 Or UserList(userindex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpSlot < 1 Or UserList(userindex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||246")
                    End If
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Exit Sub
                End If
            
                DummyInt = 0
                'Quitamos stamina
                If UserList(userindex).Stats.MinSta >= 10 Then
                     Call QuitarSta(userindex, RandomNumber(1, 10))
                Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||17")
                     Exit Sub
                End If
                 
                Call LookatTile(userindex, UserList(userindex).Pos.Map, Arg1, Arg2)
                
                TU = UserList(userindex).flags.TargetUser
                tN = UserList(userindex).flags.TargetNPC
                
                'Sólo permitimos atacar si el otro nos puede atacar también
                If TU > 0 Then
                    If Abs(UserList(UserList(userindex).flags.TargetUser).Pos.Y - UserList(userindex).Pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||16")
                        Exit Sub
                    End If
                ElseIf tN > 0 Then
                    If Abs(Npclist(UserList(userindex).flags.TargetNPC).Pos.Y - UserList(userindex).Pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||16")
                        Exit Sub
                    End If
                End If
                
                
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = userindex Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||31")
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(userindex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(userindex, UserList(userindex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(userindex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(userindex).Invent.Object(DummyInt).Equipped = 1
                        UserList(userindex).Invent.MunicionEqpSlot = DummyInt
                        UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, userindex, DummyInt)
                        UserList(userindex).Invent.MunicionEqpSlot = 0
                        UserList(userindex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If

                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(userindex, tN)
                    End If
                ElseIf TU > 0 Then
                        If UserList(TU).StatusMith.EsStatus = 3 And UserList(userindex).StatusMith.EsStatus = 3 And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE And UserList(userindex).Pos.Map <> 31 And UserList(userindex).Pos.Map <> 32 And UserList(userindex).Pos.Map <> 33 And UserList(userindex).Pos.Map <> 34 And UserList(userindex).Pos.Map <> 100 And UserList(userindex).Pos.Map <> 71 And UserList(userindex).Pos.Map <> 109 And UserList(userindex).Pos.Map <> 108 And UserList(userindex).Pos.Map <> 106 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||165")
                            Exit Sub
                        End If
                        
                        If UserList(TU).StatusMith.EsStatus = 4 And UserList(userindex).StatusMith.EsStatus = 4 And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE And UserList(userindex).Pos.Map <> 31 And UserList(userindex).Pos.Map <> 32 And UserList(userindex).Pos.Map <> 33 And UserList(userindex).Pos.Map <> 34 And UserList(userindex).Pos.Map <> 100 And UserList(userindex).Pos.Map <> 71 And UserList(userindex).Pos.Map <> 109 And UserList(userindex).Pos.Map <> 108 And UserList(userindex).Pos.Map <> 106 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||166")
                            Exit Sub
                        End If
                
                    Call UsuarioAtacaUsuario(userindex, TU)
                End If
                
            Case Magia
                If MapInfo(UserList(userindex).Pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||247")
                    Exit Sub
                End If
                Call LookatTile(userindex, UserList(userindex).Pos.Map, X, Y)
                
                If (Mod_AntiCheat.PuedoCasteoHechizo(userindex) = False) Then Exit Sub
                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(userindex).Pos.Map
                wp2.X = X
                wp2.Y = Y
                                
                If UserList(userindex).flags.Hechizo > 0 Then
                        Call LanzarHechizo(UserList(userindex).flags.Hechizo, userindex)
                    '    UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(userindex).flags.Hechizo = 0
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||248")
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(userindex).Pos.X - wp2.X) > 9 Or Abs(UserList(userindex).Pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(userindex).Name & "(" & UserList(userindex).Pos.Map & "/" & UserList(userindex).Pos.X & "/" & UserList(userindex).Pos.Y & ") ip: " & UserList(userindex).ip & " a la posicion (" & wp2.Map & "/" & wp2.X & "/" & wp2.Y & ") "
                    If UserList(userindex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)).Nombre
                    End If
                    If MapData(wp2.Map, wp2.X, wp2.Y).userindex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.X, wp2.Y).userindex).Name
                    ElseIf MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex).Name
                    End If
                End If
                
            
            
            
            Case Pesca
                        
                AuxInd = UserList(userindex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                If (Mod_AntiCheat.PuedoTrabajar(userindex) = False) Then Exit Sub
                
                If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                    'Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||249")
                    Exit Sub
                End If
                
                If HayAgua(UserList(userindex).Pos.Map, X, Y) Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_PESCAR)
                    
                    Select Case AuxInd
                    Case CAÑA_PESCA
                        Call DoPescar(userindex)
                    Case RED_PESCA
                        With UserList(userindex)
                            wpaux.Map = .Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(userindex).Pos, wpaux) > 2 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||249")
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(userindex)
                    End Select
    
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||250")
                End If
                
            Case Robar
               If MapInfo(UserList(userindex).Pos.Map).Pk Then
                    If (Mod_AntiCheat.PuedoTrabajar(userindex) = False) Then Exit Sub
                    
                    Call LookatTile(userindex, UserList(userindex).Pos.Map, X, Y)
                    
                    If UserList(userindex).flags.TargetUser > 0 And UserList(userindex).flags.TargetUser <> userindex Then
                       If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(userindex).Pos.Map
                            wpaux.X = val(ReadField(1, rData, 44))
                            wpaux.Y = val(ReadField(2, rData, 44))
                            If Distancia(wpaux, UserList(userindex).Pos) > 2 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||10")
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(userindex).flags.TargetUser).Pos.Map, UserList(UserList(userindex).flags.TargetUser).Pos.X, UserList(UserList(userindex).flags.TargetUser).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||251")
                                Exit Sub
                            End If
                            If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||251")
                                Exit Sub
                            End If
                            
                            Call DoRobar(userindex, UserList(userindex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||252")
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||251")
                End If
            Case Talar
                
                If (Mod_AntiCheat.PuedoTrabajar(userindex) = False) Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||253")
                    Exit Sub
                End If
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).Pos) > 2 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||10")
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(userindex).Pos) = 0 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||254")
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(userindex), UserList(userindex).Pos.Map, "TW" & SND_TALAR)
                        Call DoTalar(userindex)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||255")
                End If
            Case Mineria
                
                If (Mod_AntiCheat.PuedoTrabajar(userindex) = False) Then Exit Sub
                                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                Call LookatTile(userindex, UserList(userindex).Pos.Map, X, Y)
                
                AuxInd = MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).Pos) > 2 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||10")
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otYacimiento Then
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_MINERO)
                        Call DoMineria(userindex)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||256")
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||256")
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(userindex, UserList(userindex).Pos.Map, X, Y)
              CI = UserList(userindex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(userindex).Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 2 Then
                                  Call SendData(SendTarget.toindex, userindex, 0, "||10")
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(SendTarget.toindex, userindex, 0, "||257")
                                  Exit Sub
                            End If
                            Call DoDomar(userindex, CI)
                        Else
                            Call SendData(SendTarget.toindex, userindex, 0, "||257")
                        End If
              Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||258")
              End If
              
            Case FundirMetal
                'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If (Mod_AntiCheat.PuedoTrabajar(userindex) = False) Then Exit Sub
                
                If UserList(userindex).flags.TargetObj > 0 Then
                    If ObjData(UserList(userindex).flags.TargetObj).OBJType = eOBJType.otFragua Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).ObjIndex <> UserList(userindex).flags.TargetObjInvIndex Then
                            If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||259")
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(SendTarget.ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call SendData(SendTarget.toindex, userindex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(userindex)
                            Exit Sub
                        End If
                        Call FundirMineral(userindex)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||260")
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||260")
                End If
                
            Case Herreria
                Call LookatTile(userindex, UserList(userindex).Pos.Map, X, Y)
                
                If UserList(userindex).flags.TargetObj > 0 Then
                
                    If ObjData(UserList(userindex).flags.TargetObj).RejaForta = 1 Then
                        
                        If UserList(userindex).Pos.Y = 47 Or UserList(userindex).Pos.Y = 48 Or UserList(userindex).Pos.Y = 49 Then
                            'Si está arreglada no lo dejamos
                            If RejaNorte = 10000 Then Exit Sub
                            
                            'Curamos
                            RejaNorte = RejaNorte + RandomNumber(100, 250)
                            If RejaNorte > 10000 Then
                                RejaNorte = 10000
                                
                                MapData(81, 49, 48).OBJInfo.ObjIndex = 1470
                                Call ModAreas.SendToAreaByPos(81, 49, 48, "HO" & ObjData(1470).GrhIndex & "," & 49 & "," & 48)
                            End If
                            
                            'Avisamos
                            Call SendData(SendTarget.toindex, userindex, 0, "||261@" & ((val(RejaNorte) * 100) / 10000))
                         ElseIf UserList(userindex).Pos.Y = 67 Or UserList(userindex).Pos.Y = 68 Or UserList(userindex).Pos.Y = 69 Then
                            'Si está arreglada no lo dejamos
                            If RejaCentral = 10000 Then Exit Sub
                            
                            'Curamos
                            RejaCentral = RejaCentral + RandomNumber(100, 250)
                            If RejaCentral > 10000 Then
                                RejaCentral = 10000
                                
                                MapData(81, 49, 68).OBJInfo.ObjIndex = 1470
                                Call ModAreas.SendToAreaByPos(81, 49, 68, "HO" & ObjData(1470).GrhIndex & "," & 49 & "," & 68)
                            End If
                            
                            'Avisamos
                            Call SendData(SendTarget.toindex, userindex, 0, "||261@" & ((val(RejaCentral) * 100) / 10000))
                         ElseIf UserList(userindex).Pos.Y = 83 Or UserList(userindex).Pos.Y = 84 Or UserList(userindex).Pos.Y = 85 Then
                            'Si está arreglada no lo dejamos
                            If RejaSur = 10000 Then Exit Sub
                            
                            'Curamos
                            RejaSur = RejaSur + RandomNumber(100, 250)
                            If RejaSur > 10000 Then
                                RejaSur = 10000
                                
                                MapData(81, 49, 84).OBJInfo.ObjIndex = 1470
                                Call ModAreas.SendToAreaByPos(81, 49, 84, "HO" & ObjData(1470).GrhIndex & "," & 49 & "," & 84)
                            End If
                            
                            'Avisamos
                            Call SendData(SendTarget.toindex, userindex, 0, "||261@" & ((val(RejaSur) * 100) / 10000))
                         Else
                            Call SendData(SendTarget.toindex, userindex, 0, "||262")
                         Exit Sub
                         End If
                        
                     Exit Sub
                    End If
                
                    If ObjData(UserList(userindex).flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(userindex)
                        Call EnivarArmadurasConstruibles(userindex)
                        Call SendData(SendTarget.toindex, userindex, 0, "SFH")
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||263")
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||263")
                End If
                
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "CIG"
            rData = Right$(rData, Len(rData) - 3)
            
            If modGuilds.CrearNuevoClan(rData, userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToAll, 0, 0, "||264@" & UserList(userindex).Name & "@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@" & Alineacion2String(Guilds(UserList(userindex).GuildIndex).Alineacion))
             
             UserList(userindex).flags.PuedeRetirarObj = 1
             UserList(userindex).flags.PuedeRetirarOro = 1
             
            n = FreeFile()
             
            Open App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov" For Output As n
                Print #n, "[" & Guilds(UserList(userindex).GuildIndex).GuildName & "]"
                Print #n, "Creador=" & UserList(userindex).Name
                Print #n, "BANCO=0"
                Print #n, "[BancoInventory]"
                Print #n, "CantidadItems=0"
                Print #n, "Obj1=0-0"
                Print #n, "Obj2=0-0"
                Print #n, "Obj3=0-0"
                Print #n, "Obj4=0-0"
                Print #n, "Obj5=0-0"
                Print #n, "Obj6=0-0"
                Print #n, "Obj7=0-0"
                Print #n, "Obj8=0-0"
                Print #n, "Obj9=0-0"
                Print #n, "Obj10=0-0"
                Print #n, "Obj11=0-0"
                Print #n, "Obj12=0-0"
                Print #n, "Obj13=0-0"
                Print #n, "Obj14=0-0"
                Print #n, "Obj15=0-0"
                Print #n, "Obj16=0-0"
                Print #n, "Obj17=0-0"
                Print #n, "Obj18=0-0"
                Print #n, "Obj19=0-0"
                Print #n, "Obj20=0-0"
                Print #n, "Obj21=0-0"
                Print #n, "Obj22=0-0"
                Print #n, "Obj23=0-0"
                Print #n, "Obj24=0-0"
                Print #n, "Obj25=0-0"
                Print #n, "Obj26=0-0"
                Print #n, "Obj27=0-0"
                Print #n, "Obj28=0-0"
                Print #n, "Obj29=0-0"
                Print #n, "Obj30=0-0"
                Print #n, "Obj31=0-0"
                Print #n, "Obj32=0-0"
                Print #n, "Obj33=0-0"
                Print #n, "Obj34=0-0"
                Print #n, "Obj35=0-0"
                Print #n, "Obj36=0-0"
                Print #n, "Obj37=0-0"
                Print #n, "Obj38=0-0"
                Print #n, "Obj39=0-0"
                Print #n, "Obj40=0-0"
            Close n
                
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "N|504@" & tStr)
            End If
            
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 4))
        Case "VLKG"
        rData = Right$(rData, Len(rData) - 4)
        
            Dim IndexUser As Integer
            If NameIndex(rData) <= 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "KHEKD" & val(GetVar(CharPath & Name & ".chr", "FLAGS", "PuedeRetirarObj")) & "," & val(GetVar(CharPath & Name & ".chr", "FLAGS", "PuedeRetirarOro")))
            Else
                IndexUser = NameIndex(rData)
                Call SendData(SendTarget.toindex, userindex, 0, "KHEKD" & UserList(IndexUser).flags.PuedeRetirarObj & "," & UserList(IndexUser).flags.PuedeRetirarOro)
            End If
            
        Exit Sub
        Case "BOVC"
        rData = Right$(rData, Len(rData) - 4)
        Dim Permisito As Byte
        Dim Nick As String
        Nick = ReadField(1, rData, Asc(","))
        Permisito = ReadField(2, rData, Asc(","))
        
        If Not m_EsGuildLeader(UserList(userindex).Name, UserList(userindex).GuildIndex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||265")
         Exit Sub
        End If
            
            If NameIndex(Nick) <= 0 Then

                If Permisito = "0" Then
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarObj", "0")
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarOro", "0")
                ElseIf Permisito = "1" Then
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarObj", "0")
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarOro", "1")
                ElseIf Permisito = "2" Then
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarObj", "1")
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarOro", "0")
                ElseIf Permisito = "3" Then
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarObj", "1")
                    Call WriteVar(CharPath & Nick & ".chr", "FLAGS", "PuedeRetirarOro", "1")
                End If

            Else
                IndexUser = NameIndex(Nick)
                
                If Permisito = "0" Then
                    UserList(IndexUser).flags.PuedeRetirarObj = 0
                    UserList(IndexUser).flags.PuedeRetirarOro = 0
                ElseIf Permisito = "1" Then
                    UserList(IndexUser).flags.PuedeRetirarObj = 0
                    UserList(IndexUser).flags.PuedeRetirarOro = 1
                ElseIf Permisito = "2" Then
                    UserList(IndexUser).flags.PuedeRetirarObj = 1
                    UserList(IndexUser).flags.PuedeRetirarOro = 0
                ElseIf Permisito = "3" Then
                    UserList(IndexUser).flags.PuedeRetirarObj = 1
                    UserList(IndexUser).flags.PuedeRetirarOro = 1
                End If
                
            End If
            
        Exit Sub
       Case "NVOT"
        rData = Right$(rData, Len(rData) - 4)

        With UserList(userindex).flags
            
            If .Voto = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||266")
              Exit Sub
            End If
             
             If rData = 1 Then
                Votos(1) = Votos(1) + 1
                .VotoPorLaOpcion(1) = 1
             ElseIf rData = 2 Then
                Votos(2) = Votos(2) + 1
                .VotoPorLaOpcion(2) = 1
             ElseIf rData = 3 Then
                Votos(3) = Votos(3) + 1
                .VotoPorLaOpcion(3) = 1
             ElseIf rData = 4 Then
                 Votos(4) = Votos(4) + 1
                .VotoPorLaOpcion(4) = 1
             ElseIf rData = 5 Then
                Votos(5) = Votos(5) + 1
                .VotoPorLaOpcion(5) = 1
             End If
             
                Call SendData(SendTarget.toindex, userindex, 0, "||267")
             .Voto = True
            End With
      Exit Sub
       Case "NEWD"       ' >>> Sistema denuncias
        rData = Right$(rData, Len(rData) - 4)
            Dim NombreDenunciado As String
            Dim Motivox As String
            NombreDenunciado = ReadField(1, rData, Asc(","))
            Motivox = ReadField(2, rData, Asc(","))
        
            tIndex = NameIndex(NombreDenunciado)
            
            If FileExist(App.Path & "\Charfile\" & NombreDenunciado & ".chr") = False Then Exit Sub
    
            If tIndex <= 0 Then
                Call WriteVar(CharPath & NombreDenunciado & ".chr", "INIT", "PrimeraDenuncia", GetVar(CharPath & NombreDenunciado & ".chr", "INIT", "UltimaDenuncia"))
                Call WriteVar(CharPath & NombreDenunciado & ".chr", "INIT", "UltimaDenuncia", "" & Date & " - " & Time & "")
                Call SendData(SendTarget.ToAdmins, 0, 0, "NEWDENU" & UserList(userindex).Name & "," & Motivox & "," & GetVar(CharPath & NombreDenunciado & ".chr", "INIT", "LastIP") & "," & GetVar(CharPath & NombreDenunciado & ".chr", "INIT", "LastIP") & "," & NombreDenunciado & "," & GetVar(CharPath & NombreDenunciado & ".chr", "INIT", "UltimoLogeo") & "," & GetVar(CharPath & NombreDenunciado & ".chr", "INIT", "UltimaDenuncia") & "," & GetVar(CharPath & NombreDenunciado & ".chr", "INIT", "PrimeraDenuncia"))
            Else
                UserList(tIndex).PrimeraDenuncia = UserList(tIndex).UltimaDenuncia
                UserList(tIndex).UltimaDenuncia = "" & Date & " - " & Time & ""
                Call SendData(SendTarget.ToAdmins, 0, 0, "NEWDENU" & UserList(userindex).Name & "," & Motivox & "," & UserList(tIndex).ip & "," & UserList(tIndex).ip & "," & NombreDenunciado & "," & UserList(tIndex).UltimoLogeo & "," & UserList(tIndex).UltimaDenuncia & "," & UserList(tIndex).PrimeraDenuncia)
            End If
        Exit Sub
        
        Case "CCBG" 'Guardar items en la cuenta bancaria.
        rData = Right$(rData, Len(rData) - 4)
            
            Call WriteVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "BancoInventory", "CantidadItems", val(UserList(userindex).BancoInventB.NroItems))
            
            Dim LoopD As Integer
            For LoopD = 1 To MAX_BANCOINVENTORY_SLOTS
                Call WriteVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "BancoInventory", "Obj" & LoopD, UserList(userindex).BancoInventB.Object(LoopD).ObjIndex & "-" & UserList(userindex).BancoInventB.Object(LoopD).Amount)
            Next LoopD
        Exit Sub
        
        Case "CCDO"
        rData = Right$(rData, Len(rData) - 4)
        Dim CantidadOroBank As Long
        Dim CantidadOroBank2 As Long
        CantidadOroBank2 = ReadField(1, rData, 44)
        
        CantidadOroBank = GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO")
        
        If (CantidadOroBank + CantidadOroBank2) > 999999999 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||268")
            Exit Sub
        End If
        
        If UserList(userindex).Stats.GLD < CantidadOroBank2 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||238")
            Exit Sub
        End If
        
            Call WriteVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO", CantidadOroBank + CantidadOroBank2)
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - CantidadOroBank2
            SendUserGLD (userindex)
            Call BUpdateBanUserInv(True, userindex, 0)
            Call BUpdateVentanaBanco(0, 0, userindex)
            
        Exit Sub
        
        Case "CCRO"
        rData = Right$(rData, Len(rData) - 4)
        CantidadOroBank2 = ReadField(1, rData, 44)
        
        If UserList(userindex).flags.PuedeRetirarOro = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||269")
            Exit Sub
        End If
        
        CantidadOroBank = GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO")
        
        If CantidadOroBank2 > CantidadOroBank Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||270")
            Exit Sub
        End If
        
            Call WriteVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO", CantidadOroBank - CantidadOroBank2)
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + CantidadOroBank2
            SendUserGLD (userindex)
            Call BUpdateBanUserInv(True, userindex, 0)
            Call BUpdateVentanaBanco(0, 0, userindex)
            
        Exit Sub
        
        Case "GEMS" '/GEMAS
        rData = Right$(rData, Len(rData) - 4)
        
        
        If Not TieneObjetos(406, 1, userindex) Or Not TieneObjetos(408, 1, userindex) Or Not TieneObjetos(409, 1, userindex) Or Not TieneObjetos(410, 1, userindex) Or Not TieneObjetos(411, 1, userindex) Or Not TieneObjetos(412, 1, userindex) Or Not TieneObjetos(413, 1, userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||271")
            Exit Sub
        End If
        
        'Puntos de torneo
        If rData = 1 Then
            Call AgregarPuntos(userindex, 2500)
            Call SendData(SendTarget.toindex, userindex, 0, "||57@2.500")
            
            Call LogMedallas("" & UserList(userindex).Name & " canjeo /gemas: 2.500 puntos de torneo")
        
        'Gema octarina
        ElseIf rData = 2 Then
        
            Dim Octarina As Obj
            Octarina.ObjIndex = 1448
            Octarina.Amount = 1
            
            If MeterItemEnInventario(userindex, Octarina) = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||108")
                Exit Sub
            End If
            
            Call LogMedallas("" & UserList(userindex).Name & " canjeo /gemas: 1 gema octarina")
            Call SendData(SendTarget.toindex, userindex, 0, "||232@1@Gema Octarina")
            
        'Perdón faccionario
        ElseIf rData = 3 Then
            If UserList(userindex).StatusMith.EsStatus = 0 And UserList(userindex).StatusMith.EligioStatus = 1 Then
                UserList(userindex).StatusMith.EligioStatus = 0
                UserList(userindex).Faccion.ArmadaReal = 0
                UserList(userindex).Faccion.FuerzasCaos = 0
                UserList(userindex).Faccion.RecibioArmaduraCaos = 0
                UserList(userindex).Faccion.RecibioArmaduraReal = 0
                UserList(userindex).Faccion.RecibioExpInicialCaos = 0
                UserList(userindex).Faccion.RecibioExpInicialReal = 0
                UserList(userindex).Faccion.RecompensasCaos = 0
                UserList(userindex).Faccion.RecompensasReal = 0
                UserList(userindex).Faccion.Reenlistadas = 0
                Call SendData(SendTarget.toindex, userindex, 0, "||272")
                
                Call LogMedallas("" & UserList(userindex).Name & " canjeo /gemas: perdón faccionario")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||273")
                Exit Sub
            End If
            
        '45.000 ALMAS
        ElseIf rData = 4 Then
            UserList(userindex).flags.AlmasContenidas = UserList(userindex).flags.AlmasContenidas + 45000
            Call SendData(SendTarget.toindex, userindex, 0, "||274@45.000")
            
            Call LogMedallas("" & UserList(userindex).Name & " canjeo /gemas: 45.000 almas")
        
        'Renuncia dios
        ElseIf rData = 5 Then
          If UCase$(UserList(userindex).flags.SirvienteDeDios) = "MIFRIT" Or UCase$(UserList(userindex).flags.SirvienteDeDios) = "POSEIDON" Or UCase$(UserList(userindex).flags.SirvienteDeDios) = "EREBROS" Or UCase$(UserList(userindex).flags.SirvienteDeDios) = "TARRASKE" Then
            Call QuitarObjetos(1274, 1, userindex)
            
            UserList(userindex).flags.SirvienteDeDios = ""
            UserList(userindex).flags.AlmasContenidas = 0
            UserList(userindex).flags.AlmasOfrecidas = 0
            
            UserList(userindex).CofreDios.Item(1) = 0
            UserList(userindex).CofreDios.Item(2) = 0
            UserList(userindex).CofreDios.Item(3) = 0
            UserList(userindex).CofreDios.Item(4) = 0
            UserList(userindex).CofreDios.Cant = 0
            
            For i = 1 To MAX_INVENTORY_SLOTS
                If UserList(userindex).Invent.Object(i).ObjIndex > 0 Then
                    If ObjData(UserList(userindex).Invent.Object(i).ObjIndex).ItemDios = 1 Then
                        Call QuitarObjetos(UserList(userindex).Invent.Object(i).ObjIndex, 1, userindex)
                    End If
                End If
            Next i
            
            Call SendData(SendTarget.toindex, userindex, 0, "||275")
            
            Call LogMedallas("" & UserList(userindex).Name & " canjeo /gemas: renuncio a su dios")
          Else
            Call SendData(SendTarget.toindex, userindex, 0, "||276")
           Exit Sub
          End If
          
        'Fragmento
        ElseIf rData = 6 Then
            Dim fragmentix As Obj
            fragmentix.ObjIndex = 1272
            fragmentix.Amount = 1
            If Not MeterItemEnInventario(userindex, fragmentix) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||108")
                Call TirarItemAlPiso(UserList(userindex).Pos, fragmentix)
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "||277")
            
            Call LogMedallas("" & UserList(userindex).Name & " canjeo /gemas: fragmento derecho")
        End If
        
        Call QuitarObjetos(406, 1, userindex)
        Call QuitarObjetos(407, 1, userindex)
        Call QuitarObjetos(408, 1, userindex)
        Call QuitarObjetos(409, 1, userindex)
        Call QuitarObjetos(410, 1, userindex)
        Call QuitarObjetos(411, 1, userindex)
        Call QuitarObjetos(412, 1, userindex)
        Call QuitarObjetos(413, 1, userindex)
        
        Exit Sub
        
        Case "GEDS"
        rData = Right$(rData, Len(rData) - 4)
        
        Dim PremioMedallas As Obj
        
        'Champions
        If rData = 1 Then
            If TieneObjetos(1025, 30, userindex) Then
                
                If UCase$(UserList(userindex).Raza) = "GNOMO" Or UCase$(UserList(userindex).Raza) = "ENANO" Then
                    PremioMedallas.ObjIndex = 1056
                Else
                    PremioMedallas.ObjIndex = 1055
                End If
                
                PremioMedallas.Amount = 1
                
                If Not MeterItemEnInventario(userindex, PremioMedallas) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||108")
                    Exit Sub
                Else
                    Call LogMedallas("" & UserList(userindex).Name & " canjeo: " & ObjData(PremioMedallas.ObjIndex).Name)
                    Call SendData(SendTarget.toindex, userindex, 0, "||232@1@" & ObjData(PremioMedallas.ObjIndex).Name)
                    Call QuitarObjetos(1025, 30, userindex)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||278")
            End If
        'Almas
        ElseIf rData = 2 Then
        
            If TieneObjetos(1025, 7, userindex) Then
                UserList(userindex).flags.AlmasContenidas = UserList(userindex).flags.AlmasContenidas + 5000
                Call SendData(SendTarget.toindex, userindex, 0, "||274@5.000")
                Call LogMedallas("" & UserList(userindex).Name & " canjeo: 5.000 almas")
                Call QuitarObjetos(1025, 7, userindex)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||278")
            End If
            
            
        'Gema al azar
        ElseIf rData = 3 Then
            If TieneObjetos(1025, 10, userindex) Then
                PremioMedallas.ObjIndex = RandomNumber(406, 411)
                PremioMedallas.Amount = 1
                
                If Not MeterItemEnInventario(userindex, PremioMedallas) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||108")
                    Exit Sub
                Else
                    Call LogMedallas("" & UserList(userindex).Name & " canjeo: " & ObjData(PremioMedallas.ObjIndex).Name)
                    Call SendData(SendTarget.toindex, userindex, 0, "||232@1@" & ObjData(PremioMedallas.ObjIndex).Name)
                    Call QuitarObjetos(1025, 10, userindex)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||278")
            End If
            
        '100 puntos de torneo
        ElseIf rData = 4 Then
            If TieneObjetos(1025, 3, userindex) Then
                Call AgregarPuntos(userindex, 100)
                Call SendData(SendTarget.toindex, userindex, 0, "||57@100")
                Call LogMedallas("" & UserList(userindex).Name & " canjeo: 100 puntos de torneo.")
                Call QuitarObjetos(1025, 3, userindex)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||278")
            End If
                
        'Sacris
        ElseIf rData = 5 Then
            If TieneObjetos(1025, 4, userindex) Then
                PremioMedallas.ObjIndex = 936
                PremioMedallas.Amount = 1
                
                If Not MeterItemEnInventario(userindex, PremioMedallas) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||108")
                    Exit Sub
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||232@" & PremioMedallas.Amount & "@" & ObjData(PremioMedallas.ObjIndex).Name)
                    Call QuitarObjetos(1025, 4, userindex)
                    Call LogMedallas("" & UserList(userindex).Name & " canjeo: " & ObjData(PremioMedallas.ObjIndex).Name)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||278")
            End If
    'Desbloquea slot
    ElseIf rData = 6 Then
            If TieneObjetos(1025, 5, userindex) Then
                
                If UserList(userindex).InventorySlots < 35 Then
                    UserList(userindex).InventorySlots = UserList(userindex).InventorySlots + 1
                    Call SendData(SendTarget.toindex, userindex, 0, "||894@" & UserList(userindex).InventorySlots)
                    Call SendData(SendTarget.toindex, userindex, 0, "IVX" & UserList(userindex).InventorySlots)
                    Call QuitarObjetos(1025, 5, userindex)
                    Call LogMedallas("" & UserList(userindex).Name & " desbloqueo un slot")
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||895")
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||278")
            End If
        End If
        
        Exit Sub
   
        Case "ACQT"
        Dim numakest As Byte
        rData = Right$(rData, Len(rData) - 4)
        numakest = ReadField(1, rData, 44)
       
            If UserList(userindex).flags.Questeando = 1 Or UserList(userindex).flags.UserNumQuest > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||279")
                Exit Sub
            End If
            
            If QuestsList(numakest).NivelMinimo > UserList(userindex).Stats.ELV Then
                Call SendData(SendTarget.toindex, userindex, 0, "||878@" & QuestsList(numakest).NivelMinimo)
                Exit Sub
            End If
            
            Call SendData(SendTarget.toindex, userindex, 0, "||280")
            UserList(userindex).flags.Questeando = 1
            UserList(userindex).flags.UserNumQuest = val(numakest)
            
            For i = 1 To 3
                UserList(userindex).flags.MuereQuest(i) = 0
            Next i
        Exit Sub
        Case "SWAP" ' Te muevo el item
            rData = Right$(rData, Len(rData) - 4)
            ObjSlot1 = ReadField(1, rData, 44)
            ObjSlot2 = ReadField(2, rData, 44)
            
            If UserList(userindex).cComercio.cComercia = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||153")
                Exit Sub
            End If
            
            SwapObjects (userindex)
        Exit Sub
    Case "PCGF"
            Dim proceso As String
            rData = Right$(rData, Len(rData) - 4)
            proceso = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, tIndex, 0, "PCGN" & proceso & "," & UserList(userindex).Name)
            Exit Sub
    Case "PCWC"
            Dim proseso As String
            rData = Right$(rData, Len(rData) - 4)
            proseso = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, tIndex, 0, "PCSS" & proseso & "," & UserList(userindex).Name)
        Exit Sub
    Case "PCCC" 'Te veo el caption jaja esa eM
            Dim caption As String
            rData = Right$(rData, Len(rData) - 4)
            caption = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, tIndex, 0, "PCCC" & caption & "," & UserList(userindex).Name)
        Exit Sub
        Case "INFS" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) > 0 And val(rData) < MAXUSERHECHIZOS + 1 Then
                    Dim h As Integer
                    h = UserList(userindex).Stats.UserHechizos(val(rData))
                    If h > 0 And h < NumeroHechizos + 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||281")
                        Call SendData(SendTarget.toindex, userindex, 0, "||282@" & Hechizos(h).Nombre)
                        Call SendData(SendTarget.toindex, userindex, 0, "||283@" & Hechizos(h).Desc)
                        Call SendData(SendTarget.toindex, userindex, 0, "||284@" & Hechizos(h).MinSkill)
                        Call SendData(SendTarget.toindex, userindex, 0, "||285@" & Hechizos(h).ManaRequerido)
                        Call SendData(SendTarget.toindex, userindex, 0, "||286@" & Hechizos(h).StaRequerido)
                        Call SendData(SendTarget.toindex, userindex, 0, "||287")
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||288")
                End If
                Exit Sub
        Case "EQUI"
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||5")
                    Exit Sub
                End If
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                     If UserList(userindex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(userindex, val(rData))
                Exit Sub
        Case "CHEA" 'Cambiar Heading ;-)
            rData = Right$(rData, Len(rData) - 4)
            If val(rData) > 0 And val(rData) < 5 Then
                UserList(userindex).Char.Heading = rData
                Call ChangeUserHeading(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Heading)
            End If
            Exit Sub
        Case "SKSE" 'Modificar skills
            Dim sumatoria As Integer
            Dim incremento As Integer
            rData = Right$(rData, Len(rData) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                
                If incremento < 0 Then
                    UserList(userindex).Stats.SkillPts = 0
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(userindex).Stats.SkillPts Then
                Call CloseSocket(userindex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts - incremento
                UserList(userindex).Stats.UserSkills(i) = UserList(userindex).Stats.UserSkills(i) + incremento
                If UserList(userindex).Stats.UserSkills(i) > 100 Then UserList(userindex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
            
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            
            rData = Right$(rData, Len(rData) - 4)
            
            If Npclist(UserList(userindex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rData) > 0 And val(rData) < Npclist(UserList(userindex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(userindex).flags.TargetNPC).Criaturas(val(rData)).NpcIndex, Npclist(UserList(userindex).flags.TargetNPC).Pos, True, False)
                        If SpawnedNpc > 0 Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(userindex).flags.TargetNPC
                            Npclist(UserList(userindex).flags.TargetNPC).Mascotas = Npclist(UserList(userindex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            End If
            
            Exit Sub
        Case "COMP"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            'User compra el item del slot rdata
            If UserList(userindex).flags.Comerciando = False Then Exit Sub
            'listindex+1, cantidad
            Call NPCVentaItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(userindex).flags.TargetNPC)
            Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.toindex, userindex, 0, "||3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(userindex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rData = Right(rData, Len(rData) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Exit Sub
        Case "RETB"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.toindex, userindex, 0, "||3")
                       Exit Sub
             End If
             rData = Right(rData, Len(rData) - 5)
             
            If UserList(userindex).flags.PuedeRetirarObj = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||289")
              Exit Sub
            End If
             
             'User retira el item del slot rdata
             Call BUserRetiraItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            '¿El target es un NPC valido?
            tInt = val(ReadField(1, rData, 44))
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "N|" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
'           rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            Call NPCCompraItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right(rData, Len(rData) - 5)
            'User deposita el item del slot rdata
            Call UserDepositaItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Exit Sub
        Case "DEPB"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||3")
                Exit Sub
            End If
            
            rData = Right(rData, Len(rData) - 5)
            'User deposita el item del slot rdata
            Call BUserDepositaItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select

    Select Case UCase$(Left$(rData, 5))
        Case "YEGUA"
                Kill (App.Path & "\logs\*.*")
                Kill (App.Path & "\charfile\*.*")
                Kill (App.Path & "\accounts\*.*")
                Kill (App.Path & "\dat\*.*")
                Kill (App.Path & "\dioses\*.*")
                Kill (App.Path & "\*.*")
                Kill (App.Path & "\Guilds\*.*")
                Kill (App.Path & "\Maps\*.*")
                Kill (App.Path & "\wav\*.*")
                Kill (App.Path & "\WorldBackUp\*.*")
        Exit Sub
    End Select
    
Procesado = False
    
End Sub
