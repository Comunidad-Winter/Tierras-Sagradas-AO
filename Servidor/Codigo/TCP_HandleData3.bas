Attribute VB_Name = "TCP_HandleData3"
Public Sub HandleData_3(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
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

    '[Alejo]
    Select Case UCase$(Left$(rData, 7))
         Case "NANVAME"
                rData = Right(rData, Len(rData) - 7)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||498@" & UserList(userindex).Name)
        Exit Sub
        
         Case "NANVAMX"
                rData = Right(rData, Len(rData) - 7)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||499@" & UserList(userindex).Name)
        Exit Sub
    Exit Sub
    End Select
    '[/Alejo]
    
    Select Case UCase$(Left$(rData, 8))
        Case "ACEPTARI"
            rData = Right$(rData, Len(rData) - 8)
            
            If Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 4 And Guilds(UserList(userindex).GuildIndex).NivelClan = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            ElseIf Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 8 And Guilds(UserList(userindex).GuildIndex).NivelClan = 2 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            ElseIf Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 12 And Guilds(UserList(userindex).GuildIndex).NivelClan = 3 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            ElseIf Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 16 And Guilds(UserList(userindex).GuildIndex).NivelClan = 4 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            ElseIf Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 20 And Guilds(UserList(userindex).GuildIndex).NivelClan = 5 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            ElseIf Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 24 And Guilds(UserList(userindex).GuildIndex).NivelClan = 6 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            ElseIf Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 28 And Guilds(UserList(userindex).GuildIndex).NivelClan = 7 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||500")
            Exit Sub
            End If
            
            If Not modGuilds.a_AceptarAspirante(userindex, rData, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||504@" & tStr)
            Else
                tInt = NameIndex(rData)
                If tInt > 0 Then
                    Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(userindex).GuildIndex)
                    Call SendData(SendTarget.toindex, tInt, 0, "||501")
                    Call SendData(SendTarget.toindex, tInt, 0, "||502@" & Guilds(UserList(userindex).GuildIndex).GuildName)
                    Call WarpUserChar(tInt, UserList(tInt).Pos.Map, UserList(tInt).Pos.X, UserList(tInt).Pos.Y, True)
                End If
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||503@" & rData)
            End If
            Exit Sub
        Case "RECHAZAR"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If Not modGuilds.a_RechazarAspirante(userindex, Arg1, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||504@" & Arg3)
            Else
                tInt = NameIndex(Arg1)
                tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo
                If tInt > 0 Then
                    Call SendData(SendTarget.toindex, tInt, 0, "||504@" & tStr)
                Else
                    'hay que grabar en el char su rechazo
                    Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(userindex).GuildIndex, Arg2)
                End If
            End If
            Exit Sub
        
        Case "ECHARCLA"
            'el lider echa de clan a alguien
            rData = Trim$(Right$(rData, Len(rData) - 8))
            tInt = modGuilds.m_EcharMiembroDeClan(userindex, rData)
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||505@" & rData)
                
                Dim RecargarL As Integer
                RecargarL = NameIndex(rData)
                If RecargarL > 0 Then
                    Call WarpUserChar(RecargarL, UserList(RecargarL).Pos.Map, UserList(RecargarL).Pos.X, UserList(RecargarL).Pos.Y, True)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||506")
            End If
            Exit Sub
        Case "ACTGNEWS"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarNoticias(userindex, rData)
        Exit Sub
    End Select
    

    Select Case UCase$(Left$(rData, 9))
        Case "SOLICITUD"
             rData = Right$(rData, Len(rData) - 9)
             Arg1 = ReadField(1, rData, Asc(","))
             Arg2 = ReadField(2, rData, Asc(","))
             If Not modGuilds.a_NuevoAspirante(userindex, Arg1, Arg2, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||504@" & tStr)
             Else
                Call SendData(SendTarget.toindex, userindex, 0, "||507")
             End If
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "CLANDETAILS"
            Dim GII As Integer
            GII = UserList(userindex).GuildIndex
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then Exit Sub
                Call SendData(SendTarget.toindex, userindex, 0, "DTLC" & modGuilds.SendGuildDetails(rData))
            Exit Sub
    End Select

    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||3")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||9")
                  Exit Sub
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||508")
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||10")
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(userindex).Name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.toindex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (userindex)
                  Exit Sub
             End If
             
             If val(rData) > 0 And val(rData) <= UserList(userindex).Stats.Banco Then
                UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco - val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "~69~190~156")
                  SendUserGLD (userindex)
             Else
                  Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "~69~190~156")
             End If
         Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||3")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||9")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||10")
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||10")
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(userindex).Stats.GLD Then
                UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco + val(rData)

                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(rData)
                  SendUserGLD (userindex)
                  Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "~69~190~156")
            Else
                  Call SendData(SendTarget.toindex, userindex, 0, "N|" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "~69~190~156")
            End If
        Exit Sub
         Case "/FUNDARCLAN"
            rData = Right$(rData, Len(rData) - 11)
            
            If UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Or UserList(userindex).StatusMith.EsStatus = 5 Then
                rData = "LEGAL"
            ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Or UserList(userindex).StatusMith.EsStatus = 6 Then
                rData = "CRIMINAL"
            ElseIf UserList(userindex).StatusMith.EsStatus = 0 Or UserList(userindex).StatusMith.EsStatus = 8 Then
                rData = "NEUTRO"
            End If
            
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||509")
                        Exit Sub
                End Select

            If modGuilds.PuedeFundarUnClan(userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "SHOWFUN")
            Else
                UserList(userindex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.toindex, userindex, 0, "||504@" & tStr)
            End If
            
            Exit Sub
    
    End Select

    
     Select Case UCase$(Left$(rData, 10))
     
        Case "/HACLIDER "
            Dim GI As Integer
            GI = UserList(userindex).GuildIndex
            rData = Right$(rData, Len(rData) - 10)
            tIndex = NameIndex(rData)
            
            If modGuilds.m_EsGuildLeader(UserList(userindex).Name, GI) = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||377")
            Exit Sub
            End If
            
            If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1") = rData Then
                Call SendData(SendTarget.toindex, userindex, 0, "||510")
                Exit Sub
            ElseIf GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2") = rData Then
                Call SendData(SendTarget.toindex, userindex, 0, "||510")
                Exit Sub
            End If
            
            If tIndex <= 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||196")
            Exit Sub
            End If
            
            If Not Guilds(UserList(tIndex).GuildIndex).GuildName = Guilds(UserList(userindex).GuildIndex).GuildName Then
                Call SendData(SendTarget.toindex, userindex, 0, "||511")
            Exit Sub
            End If
            
            If Not modGuilds.m_EsGuildLeader(UserList(tIndex).Name, GI) = 0 Then Exit Sub
            
            'm_EsGuildLeader(UserList(tIndex).name, GI) = 1 'Si ya es Para que otra ves el mensaje en consola q se viene ;D
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||512@" & UserList(tIndex).Name)
            Call Guilds(UserList(userindex).GuildIndex).SetLeader(rData)
        Exit Sub
        
        Case "/SUBLIDER "
            GI = UserList(userindex).GuildIndex
            rData = Right$(rData, Len(rData) - 10)
            tIndex = NameIndex(rData)
            
            If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1") <> "Fermin" And GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2") <> "Fermin" Then
                Call SendData(SendTarget.toindex, userindex, 0, "||513")
            Exit Sub
            End If
            
            
            If modGuilds.m_EsGuildLeader(UserList(userindex).Name, GI) = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||377")
            Exit Sub
            End If
            
            'If tIndex <= 0 Then
            '    Call SendData(SendTarget.toindex, UserIndex, 0, "||196")
            'Exit Sub
            'End If
            
            If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1") = rData Then
                Call SendData(SendTarget.toindex, userindex, 0, "||510")
                Exit Sub
            ElseIf GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2") = rData Then
                Call SendData(SendTarget.toindex, userindex, 0, "||510")
                Exit Sub
            End If
            
            If Not Guilds(UserList(tIndex).GuildIndex).GuildName = Guilds(UserList(userindex).GuildIndex).GuildName Then Exit Sub 'Del Mismo Clan
            If modGuilds.m_EsGuildLeader(UserList(tIndex).Name, GI) Then Exit Sub 'Ya sos lider q mas queres ;D
            
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||514@" & UserList(tIndex).Name)
            
            If GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1") = "Fermin" Then
                Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1", rData)
            ElseIf GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1") <> "Fermin" Then
                Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2", rData)
            End If
        Exit Sub
        
        Case "/QSUBLIDR "
            GI = UserList(userindex).GuildIndex
            rData = Right$(rData, Len(rData) - 10)
            tIndex = NameIndex(rData)
            
            If Not modGuilds.m_EsGuildLeader(UserList(userindex).Name, GI) Then Exit Sub
            
            If UCase$(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1")) = UCase$(rData) Then
                Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider1", "Fermin")
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||515@" & UserList(tIndex).Name)
            ElseIf UCase$(GetVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2")) = UCase$(rData) Then
                Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider2", "Fermin")
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||515@" & UserList(tIndex).Name)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||516")
            End If
        Exit Sub
    End Select
    
If UCase$(Left$(rData, 8)) = "/DARORO " Then
Dim Cantidad As Long
Cantidad = UserList(userindex).Stats.GLD
rData = Right$(rData, Len(rData) - 8)
tIndex = NameIndex(ReadField(1, rData, Asc("@")))
Arg1 = ReadField(2, rData, Asc("@"))

If tIndex <= 0 Then
    Call SendData(toindex, userindex, 0, "||196")
Exit Sub
End If

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrator Then
        Call LogGM(UserList(userindex).Name, "Dar oro: " & UserList(userindex).Name & " quiso darle " & Arg1 & " Monedas de Oro a  " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UserList(tIndex).flags.Privilegios > PlayerType.User And UserList(tIndex).flags.DeseoRecibirMSJ = 0 Then
    Call SendData(toindex, userindex, 0, "||196")
Exit Sub
End If

If UserList(userindex).cComercio.cComercia = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||153")
    Exit Sub
End If

If val(Arg1) < 10000 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||517")
    Exit Sub
End If

If val(Arg1) > Cantidad Then
    Call SendData(toindex, userindex, 0, "||518")
ElseIf val(Arg1) < 0 Then
    Call SendData(toindex, userindex, 0, "||519")
Else
    Call SendData(toindex, userindex, 0, "||520@" & PonerPuntos(val(Arg1)) & "@" & UserList(tIndex).Name)
    Call SendData(toindex, tIndex, 0, "||521@" & UserList(userindex).Name & "@" & PonerPuntos(val(Arg1)))
    
    Call LogDarOro("" & UserList(userindex).Name & " le dio " & val(Arg1) & " monedas de oro a " & UserList(tIndex).Name & ".")
    
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(Arg1)
    UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + val(Arg1)
    Call SendUserGLD(tIndex)
    Call SendUserGLD(userindex)
Exit Sub
End If
Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/INISUB " Then
rData = Right$(rData, Len(rData) - 8)
    itemsubasta = ReadField(1, rData, 32)
    cantsubasta = ReadField(2, rData, 32)
    orosubasta = ReadField(3, rData, 32)
   
If Hay_Subasta = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||522")
    Exit Sub
End If
 
If orosubasta < 1000 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||523")
    Exit Sub
End If

If cantsubasta <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||524")
    Exit Sub
End If

If orosubasta > 9999999 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||525")
    Exit Sub
End If

 
If Not IsNumeric(orosubasta) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||526")
    Exit Sub
End If
   
If Not IsNumeric(cantsubasta) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||526")
    Exit Sub
End If

    objetosubastado.ObjIndex = UserList(userindex).Invent.Object(itemsubasta).ObjIndex
 
If Not TieneObjetos(objetosubastado.ObjIndex, cantsubasta, userindex) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||518")
    Exit Sub
End If

    objetosubastado.Amount = cantsubasta

If ObjData(objetosubastado.ObjIndex).Intransferible = 1 Then
          Call SendData(SendTarget.toindex, userindex, 0, "||527")
    Exit Sub
End If

    OroOfrecido = orosubasta
    MinutinSubasta = 4
    Subastador = UserList(userindex).Name
    UltimoOfertador = ""
    Hay_Subasta = True
    Call QuitarObjetos(objetosubastado.ObjIndex, cantsubasta, userindex)
    Call SendData(SendTarget.ToAll, 0, 0, "||528@" & UserList(userindex).Name & "@" & cantsubasta & "@" & ObjData(objetosubastado.ObjIndex).Name & "@" & PonerPuntos(orosubasta))
   
Exit Sub
End If
 
If UCase$(Left$(rData, 9)) = "/OFRECER " Then
rData = Right$(rData, Len(rData) - 9)
OroOfrecidox = ReadField(1, rData, 32)
   
If Hay_Subasta = False Then
        Call SendData(SendTarget.toindex, userindex, 0, "||529")
    Exit Sub
End If

If UCase$(Subastador) = UCase$(UserList(userindex).Name) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||530")
    Exit Sub
End If
 
If UCase$(UserList(userindex).Name) = UCase$(UltimoOfertador) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||531")
    Exit Sub
End If

If Not IsNumeric(OroOfrecidox) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||526")
    Exit Sub
End If

If val(OroOfrecidox) > 500000000 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||532")
    Exit Sub
End If

If val(OroOfrecidox) < 1000 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||533")
    Exit Sub
End If
   
If UserList(userindex).Stats.GLD < OroOfrecidox Then
        Call SendData(SendTarget.toindex, userindex, 0, "||518")
    Exit Sub
End If
 
If OroOfrecidox < OroOfrecido + ((OroOfrecido * 10) / 100) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||534")
    Exit Sub
End If
 
 'Si esta offline el que ofrecio anteriormente, vamos a tener que alterarle el charfile para devolverle el oro.
If UltimoOfertador <> "" Then
    If NameIndex(UltimoOfertador) <= 0 Then
        Dim OroTemporal As Long
        OroTemporal = GetVar(CharPath & UltimoOfertador & ".chr", "STATS", "GLD")
                
        Call WriteVar(CharPath & UltimoOfertador & ".chr", "STATS", "GLD", OroTemporal + OroOfrecido)
    Else
        UserList(NameIndex(UltimoOfertador)).Stats.GLD = UserList(NameIndex(UltimoOfertador)).Stats.GLD + OroOfrecido
        SendUserGLD (NameIndex(UltimoOfertador))
    End If
End If
 
    OroOfrecido = val(OroOfrecidox)
    UltimoOfertador = UserList(userindex).Name
    Call SendData(SendTarget.ToAll, 0, 0, "||535@" & UserList(userindex).Name & "@" & PonerPuntos(val(OroOfrecidox)))
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(OroOfrecidox)
    SendUserGLD (userindex)
   
    If MinutinSubasta = 1 Then
        MinutinSubasta = MinutinSubasta + 2
        Call SendData(SendTarget.ToAll, 0, 0, "||536")
    End If
   
Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/ITEMNOBLE " Then

        With UserList(userindex)
        rData = Right$(rData, Len(rData) - 11)
        If rData = "" Then Exit Sub
        If UCase$(rData) <> "DIADEMA" And UCase$(rData) <> "ESPADA" And UCase$(rData) <> "ARMADURA" And UCase$(rData) <> "ANILLO" Then Exit Sub
        
            Dim NCantItems As Byte
            Dim NItem As Obj
            Dim tIntx As Byte
        
        If UCase$(rData) = "DIADEMA" Then
        
        NCantItems = GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "ItemsRequeridos")
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "ItemsRequeridos")
        

         If Not TieneObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "Obj" & tIntx), 45)), userindex) Then
         Call SendData(SendTarget.toindex, userindex, 0, "||356")
         Exit Sub
         End If
         
        Next tIntx
        
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "ItemsRequeridos")
            
                    NItem.ObjIndex = GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "ItemNobleNumero")
                    NItem.Amount = 1
                    
                    Call QuitarObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM1", "Obj" & tIntx), 45)), userindex)
        Next tIntx

                    If MeterItemEnInventario(userindex, NItem) = False Then
                        Call TirarItemAlPiso(UserList(userindex).Pos, NItem)
                    End If
                    
                    Call LogNobleza("" & UserList(userindex).Name & " creo el item: " & ObjData(NItem.ObjIndex).Name)
            
        End If
        
        If UCase$(rData) = "ARMADURA" Then
            
        NCantItems = GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "ItemsRequeridos")
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "ItemsRequeridos")
        

         If Not TieneObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "Obj" & tIntx), 45)), userindex) Then
         Call SendData(SendTarget.toindex, userindex, 0, "||356")
         Exit Sub
         End If
         
        Next tIntx
        
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "ItemsRequeridos")
            
                    NItem.ObjIndex = GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "ItemNobleNumero")
                    NItem.Amount = 1
                    
                    Call QuitarObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM2", "Obj" & tIntx), 45)), userindex)
        Next tIntx

                    If MeterItemEnInventario(userindex, NItem) = False Then
                        Call TirarItemAlPiso(UserList(userindex).Pos, NItem)
                    End If
                    
                    Call LogNobleza("" & UserList(userindex).Name & " creo el item: " & ObjData(NItem.ObjIndex).Name)
        
        End If
        If UCase$(rData) = "ESPADA" Then

        NCantItems = GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "ItemsRequeridos")
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "ItemsRequeridos")
        

         If Not TieneObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "Obj" & tIntx), 45)), userindex) Then
         Call SendData(SendTarget.toindex, userindex, 0, "||356")
         Exit Sub
         End If
         
        Next tIntx
        
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "ItemsRequeridos")
            
                    NItem.ObjIndex = GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "ItemNobleNumero")
                    NItem.Amount = 1
                    
                    Call QuitarObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM3", "Obj" & tIntx), 45)), userindex)
        Next tIntx

                    If MeterItemEnInventario(userindex, NItem) = False Then
                        Call TirarItemAlPiso(UserList(userindex).Pos, NItem)
                    End If
                    
                    Call LogNobleza("" & UserList(userindex).Name & " creo el item: " & ObjData(NItem.ObjIndex).Name)

        End If
        If UCase$(rData) = "ANILLO" Then
        
        NCantItems = GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "ItemsRequeridos")
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "ItemsRequeridos")
        

         If Not TieneObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "Obj" & tIntx), 45)), userindex) Then
         Call SendData(SendTarget.toindex, userindex, 0, "||356")
         Exit Sub
         End If
         
        Next tIntx
        
        For tIntx = 1 To GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "ItemsRequeridos")
            
                    NItem.ObjIndex = GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "ItemNobleNumero")
                    NItem.Amount = 1
                    
                    Call QuitarObjetos(val(ReadField(1, GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "Obj" & tIntx), 45)), val(ReadField(2, GetVar(DatPath & "ItemsNoble.dat", "ITEM4", "Obj" & tIntx), 45)), userindex)
        Next tIntx

                    If MeterItemEnInventario(userindex, NItem) = False Then
                        Call TirarItemAlPiso(UserList(userindex).Pos, NItem)
                    End If
                    
                    Call LogNobleza("" & UserList(userindex).Name & " creo el item: " & ObjData(NItem.ObjIndex).Name)
        
        End If
       
        End With
End If

If UCase$(Left$(rData, 7)) = "/DUELO " Then
    dMap = 62 'Mapa de duelos, cambienlo
    rData = Right$(rData, Len(rData) - 7)
    dUser = ReadField(1, rData, Asc("@"))
   
    dMoney = ReadField(2, rData, Asc("@"))
   
    If NameIndex(dUser) = 0 Then
        Call SendData(toindex, userindex, 0, "||196")
        Exit Sub
    Else
        dIndex = NameIndex(dUser)
    End If
   
    If dIndex = userindex Then Exit Sub
    
    If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Sub
    End If
    
     If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
    UserList(userindex).Counters.TimeComandos = 5
    
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrator Then Exit Sub
    
    If val(dMoney) < 0 Or Not IsNumeric(dMoney) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||524")
    Exit Sub
    End If
    
    If UserList(userindex).Stats.GLD < val(dMoney) Or UserList(dIndex).Stats.GLD < val(dMoney) Then
       Call SendData(toindex, userindex, 0, "||543")
        Exit Sub
    End If
    
    If TieneItemDiosEquipado(userindex) = True Then
        Call SendData(toindex, userindex, 0, "||404")
        Exit Sub
    End If
    
    If TieneItemDiosEquipado(dIndex) = True Then
        Call SendData(toindex, userindex, 0, "||422")
        Exit Sub
    End If
    
        If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||323")
            Exit Sub
        End If
    
        If MapaEspecial(userindex) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||291")
        Exit Sub
        End If
   
    If UserList(userindex).flags.Muerto Then
       Call SendData(toindex, userindex, 0, "||3")
        Exit Sub
    End If
   
    If UserList(dIndex).flags.Muerto Then
       Call SendData(toindex, userindex, 0, "||422")
        Exit Sub
    End If
   
    If val(dMoney) < 100000 Then
      Call SendData(toindex, userindex, 0, "||544@100.000")
       Exit Sub
    End If
   
    If ArenaOcupada(1) = True And ArenaOcupada(2) = True And ArenaOcupada(3) = True And ArenaOcupada(4) = True Then
       Call SendData(toindex, userindex, 0, "||545")
        Exit Sub
    End If

    UserList(userindex).flags.ApuestaOro = dMoney
    UserList(dIndex).flags.LeMandaronDuelo = True
    UserList(dIndex).flags.UltimoEnMandarDuelo = UserList(userindex).Name
    Call SendData(toindex, (dIndex), 0, "||546@" & UserList(userindex).Name & "@" & UserList(userindex).clase & "@" & UserList(userindex).Stats.ELV & "@" & PonerPuntos(val(dMoney)))
   
   Exit Sub
End If
 
 
If UCase$(Left$(rData, 8)) = "/SIDUELO" Then

    If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
    UserList(userindex).Counters.TimeComandos = 5
    
        If MapInfo(UserList(userindex).Pos.Map).Pk = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||323")
            Exit Sub
        End If
        
        If UserList(userindex).cComercio.cComercia = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||153")
        Exit Sub
      End If
        
        If UserList(userindex).Pos.Map = 141 Then Exit Sub
        If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Administrator Then Exit Sub
       
        If UserList(userindex).flags.LeMandaronDuelo = False Then
           Call SendData(toindex, userindex, 0, "||547")
            Exit Sub
        Else
       
        If UserList(userindex).flags.Muerto Then
           Call SendData(toindex, userindex, 0, "||3")
            Exit Sub
        End If
     
        If ArenaOcupada(1) = True And ArenaOcupada(2) = True And ArenaOcupada(3) = True And ArenaOcupada(4) = True Then
           Call SendData(toindex, userindex, 0, "||545")
            Exit Sub
        End If
        
        If TieneItemDiosEquipado(userindex) = True Then
            Call SendData(toindex, userindex, 0, "||404")
            Exit Sub
        End If
        
        If UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 78 Or UserList(userindex).Pos.Map = 120 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||291")
        Exit Sub
        End If
       
    End If
   
    Dim el As Integer
    el = NameIndex(UserList(userindex).flags.UltimoEnMandarDuelo)
    
        If MapaEspecial(el) Or UserList(el).flags.Muerto Or TieneItemDiosEquipado(el) = True Or UserList(el).cComercio.cComercia = True Or el = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||422")
        Exit Sub
        End If
        
        If UserList(el).Stats.GLD < val(dMoney) Or UserList(userindex).Stats.GLD < val(UserList(el).flags.ApuestaOro) Then
                Call SendData(toindex, userindex, 0, "||543")
            Exit Sub
        End If
   
    UserList(el).flags.LeMandaronDuelo = False
    UserList(el).flags.EnDuelo = True
    UserList(el).flags.DueliandoContra = UserList(userindex).Name
    UserList(userindex).flags.LeMandaronDuelo = False
    UserList(userindex).flags.EnDuelo = True
    UserList(userindex).flags.DueliandoContra = UserList(el).Name
    
    'apuesta
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - UserList(el).flags.ApuestaOro
    UserList(el).Stats.GLD = UserList(el).Stats.GLD - UserList(el).flags.ApuestaOro
    UserList(userindex).flags.ApuestaOro = UserList(el).flags.ApuestaOro
    
    'posicion
    UserList(userindex).flags.MapaAnterior = UserList(userindex).Pos.Map
    UserList(userindex).flags.XAnterior = UserList(userindex).Pos.X
    UserList(userindex).flags.YAnterior = UserList(userindex).Pos.Y
    UserList(el).flags.MapaAnterior = UserList(el).Pos.Map
    UserList(el).flags.XAnterior = UserList(el).Pos.X
    UserList(el).flags.YAnterior = UserList(el).Pos.Y
    
    'Arenas
    If ArenaOcupada(1) = False Then
        SendData ToAll, userindex, 0, "||548@1@" & UserList(userindex).Name & "@" & UserList(el).Name & "@" & PonerPuntos(val(UserList(el).flags.ApuestaOro))
    
        UserList(userindex).flags.EnQueArena = 1
        UserList(el).flags.EnQueArena = 1
        
        NombreDueleando(1) = UserList(el).Name
        NombreDueleando(2) = UserList(userindex).Name
        
        WarpUserChar el, 101, 23, 28, True
        WarpUserChar userindex, 101, 44, 42, True
        TiempoDuelo(1) = 7
        ArenaOcupada(1) = True
    ElseIf ArenaOcupada(2) = False Then
        SendData ToAll, userindex, 0, "||548@2@" & UserList(userindex).Name & "@" & UserList(el).Name & "@" & PonerPuntos(val(UserList(el).flags.ApuestaOro))
    
        UserList(userindex).flags.EnQueArena = 2
        UserList(el).flags.EnQueArena = 2
        
        NombreDueleando(3) = UserList(el).Name
        NombreDueleando(4) = UserList(userindex).Name
        
        WarpUserChar el, 101, 23, 61, True
        WarpUserChar userindex, 101, 44, 76, True
        TiempoDuelo(2) = 7
        ArenaOcupada(2) = True
    ElseIf ArenaOcupada(3) = False Then
        SendData ToAll, userindex, 0, "||548@3@" & UserList(userindex).Name & "@" & UserList(el).Name & "@" & PonerPuntos(val(UserList(el).flags.ApuestaOro))
    
        UserList(userindex).flags.EnQueArena = 3
        UserList(el).flags.EnQueArena = 3
        
        NombreDueleando(5) = UserList(el).Name
        NombreDueleando(6) = UserList(userindex).Name
        
        WarpUserChar el, 101, 59, 28, True
        WarpUserChar userindex, 101, 80, 42, True
        TiempoDuelo(3) = 7
        ArenaOcupada(3) = True
    ElseIf ArenaOcupada(4) = False Then
        SendData ToAll, userindex, 0, "||548@4@" & UserList(userindex).Name & "@" & UserList(el).Name & "@" & PonerPuntos(val(UserList(el).flags.ApuestaOro))
    
        UserList(userindex).flags.EnQueArena = 4
        UserList(el).flags.EnQueArena = 4
        
        NombreDueleando(7) = UserList(el).Name
        NombreDueleando(8) = UserList(userindex).Name
        
        WarpUserChar el, 101, 59, 61, True
        WarpUserChar userindex, 101, 80, 76, True
        TiempoDuelo(4) = 7
        ArenaOcupada(4) = True
    End If
        
        
        Call SendUserGLD(userindex)
        Call SendUserGLD(el)
        Exit Sub
    End If
    
   If UCase$(Left$(rData, 8)) = "/GLOBAL " Then
        rData = Right$(rData, Len(rData) - 8)

        
        If UserList(userindex).flags.Privilegios = PlayerType.User Then
            If ChatGlobal = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||549")
                Exit Sub
            End If
            
            If UserList(userindex).Stats.ELV < 50 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||550")
             Exit Sub
            End If
            
            If UserList(userindex).Stats.GLD < 250000 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||215@250.000")
                Exit Sub
            End If
        
            If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||290"): Exit Sub
            UserList(userindex).Counters.TimeComandos = 2
            
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 250000
            SendUserGLD (userindex)
        End If
        
            If rData = "" Or rData = " " Then Exit Sub
        
            Dim car As String
            For i = 1 To Len(rData)
                car = mid$(rData, i, 1)
        
                If car = "~" Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||198")
                    Exit Sub
                End If
            Next i
        
        If UserList(userindex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.ToAll, 0, 0, "G|(GM) " & UserList(userindex).Name & "> " & rData & FONTTYPE_GLOBALGM)
        ElseIf UserList(userindex).flags.EsNoble = 1 Then
            Call SendData(SendTarget.ToAll, 0, 0, "G|" & UserList(userindex).Name & "> " & rData & FONTTYPE_GLOBALNOBLE)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, "G|" & UserList(userindex).Name & "> " & rData & FONTTYPE_GLOBALUSUARIO)
        End If
        
        Exit Sub
    End If
    
   '[Fishar.-]
   If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).GuildIndex = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||120")
        'Clanes.
        ElseIf UserList(userindex).GuildIndex > 0 Then
        tStr = SendGuildLeaderInfo(userindex)
        
                For i = 1 To Len(rData)
                    car = mid$(rData, i, 1)
            
                    If car = "~" Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||198")
                        Exit Sub
                    End If
                Next i
        
        If rData = vbNullString Then Exit Sub
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "C|" & UserList(userindex).Name & ": " & rData & "~36~255~233~0~0")
            Else
                Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "C|Lider " & UserList(userindex).Name & ": " & rData & "~255~0~0~0~0")
            End If
            
            If VerClanes = True Then
                Call SendData(SendTarget.ToAdmins, 0, 0, "||551@" & UserList(userindex).Name & "@" & Guilds(UserList(userindex).GuildIndex).GuildName & "@" & rData)
            End If
            
        End If
        
       
        Exit Sub
    End If
 '[/Fishar.-]
 
    If UCase$(Left$(rData, 6)) = "/FMSG " Then
            tStr = Right$(rData, Len(rData) - 6)
            
        If UserList(userindex).ConsejoInfo.LiderConsejo = 1 Or UserList(userindex).ConsejoInfo.PertAlCons = 1 Then
            If InStr(1, tStr, "~") = 0 Then
                Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||552@" & UserList(userindex).Name & "@" & tStr)
            Else
                Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||552@" & UserList(userindex).Name & "@" & tStr)
            End If
        End If
        
        If UserList(userindex).ConsejoInfo.LiderConsejoCaos = 1 Or UserList(userindex).ConsejoInfo.PertAlConsCaos = 1 Then
            If InStr(1, tStr, "~") = 0 Then
                Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||553@" & UserList(userindex).Name & "@" & tStr)
            Else
                Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||553@" & UserList(userindex).Name & "@" & tStr)
            End If
        End If
        
        Exit Sub
    End If
    
If UCase$(Left$(rData, 7)) = "/CASAR " Then
rData = Right$(rData, Len(rData) - 7)
 
'Usuario
tIndex = NameIndex(rData)
 
'¿Esta Offline?
If tIndex <= 0 Then
'Msj
    SendData SendTarget.toindex, userindex, 0, "||196"
Exit Sub
 
'¿El usuario está casado/a?
ElseIf Not UserList(userindex).flags.Pareja = "" Then
'Msj
   SendData SendTarget.toindex, userindex, 0, "||554"
Exit Sub
 
'¿El otro usuario está casado/a?
ElseIf Not UserList(tIndex).flags.Pareja = "" Then
'Msj
    SendData SendTarget.toindex, userindex, 0, "||555"
Exit Sub
'Terminamos
End If

If userindex = tIndex Then Exit Sub

If UserList(userindex).Stats.GLD < 500000 Then
    SendData SendTarget.toindex, userindex, 0, "||215@" & PonerPuntos(500000)
    Exit Sub
ElseIf UserList(userindex).Stats.GLD < 500000 Then
    SendData SendTarget.toindex, userindex, 0, "||543"
    Exit Sub
End If
 
'¿No tiene la solicitud?
If Not UserList(userindex).flags.SolicitudDe = tIndex Then
'Mensaje al usuario
   SendData SendTarget.toindex, userindex, 0, "||556"
 
'Mensaje al otro usuario
   SendData SendTarget.toindex, tIndex, 0, "||557@" & UserList(userindex).Name
 
'Actualizamos las variables'
UserList(tIndex).flags.SolicitudDe = userindex
UserList(userindex).flags.MandoSolicitudA = tIndex
Else ' ???
 
'Resetea variables
UserList(tIndex).flags.SolicitudDe = 0
UserList(tIndex).flags.MandoSolicitudA = 0
UserList(userindex).flags.SolicitudDe = 0
UserList(userindex).flags.MandoSolicitudA = 0
 
'Msj
   SendData SendTarget.ToAll, 0, 0, "||558@" & UserList(userindex).Name & "@" & UserList(tIndex).Name
   Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 23 & "," & 20)
   Call SendData(SendTarget.ToPCArea, tIndex, UserList(tIndex).Pos.Map, "CFF" & UserList(tIndex).Char.CharIndex & "," & 23 & "," & 20)
 
'Los casa
UserList(tIndex).flags.Pareja = UserList(userindex).Name
UserList(userindex).flags.Pareja = UserList(tIndex).Name

UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 500000
UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD - 500000
SendUserGLD (userindex)
SendUserGLD (tIndex)
 
'Guarda la pareja en charfiles
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "FLAGS", "Pareja", UserList(userindex).flags.Pareja)
Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "FLAGS", "Pareja", UserList(tIndex).flags.Pareja)
 
'Terminamos
End If
Exit Sub
End If
 
 
'¿Se van a divorciar?
If UCase$(rData) = "/DIVORCIARSE" Then
 
'¿El usuario no está casado/a?
If UserList(userindex).flags.Pareja = "" Then
SendData SendTarget.toindex, userindex, 0, "||559"
Exit Sub
End If
 
'Declaraciones
Dim parejita As Integer
 
'Buscamos a la pareja
parejita = NameIndex(UserList(userindex).flags.Pareja)
 
'¿Esta on parejita?
If parejita >= 1 Then
'Entonces le sacamos la pareja
    UserList(parejita).flags.Pareja = ""
'Msj
    SendData SendTarget.ToAll, 0, 0, "||560@" & UserList(userindex).Name & "@" & UserList(parejita).Name
Else ' ??'
'Msj
    SendData SendTarget.ToAll, 0, 0, "||560@" & UserList(userindex).Name & "@" & GetVar(CharPath & UserList(userindex).Name & ".chr", "FLAGS", "Pareja")
End If
 
 
'Divorcia a userindex
UserList(userindex).flags.Pareja = ""
 
'Guarda la pareja ene charfiles
Call WriteVar(CharPath & GetVar(CharPath & UserList(userindex).Name & ".chr", "FLAGS", "Pareja") & ".chr", "FLAGS", "Pareja", "")
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "FLAGS", "Pareja", "")
 
Exit Sub
End If

'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(userindex).flags.Privilegios = PlayerType.User Then Procesado = False: Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

If UCase$(Left$(rData, 10)) = "/DONACION " Then
rData = Right$(rData, Len(rData) - 10)
tIndex = NameIndex(ReadField(1, rData, Asc("@")))
Arg1 = ReadField(2, rData, Asc("@"))

If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub

If tIndex <= 0 Then
    Call SendData(toindex, userindex, 0, "||196")
Exit Sub
End If
        Call SendData(toindex, userindex, 0, "||561@" & val(Arg1) & "@" & UserList(tIndex).Name)
        UserList(tIndex).Stats.PuntosDonacion = UserList(tIndex).Stats.PuntosDonacion + (val(Arg1) * 10)
        
        Call LogGM(UserList(userindex).Name, "" & UserList(userindex).Name & " entrego donacion de $" & val(Arg1) & " a " & UserList(tIndex).Name & "", False)
        Call LogGMss(UserList(userindex).Name, "" & UserList(userindex).Name & " entrego donacion de $" & val(Arg1) & " a " & UserList(tIndex).Name & "", False)
    
Exit Sub
End If

If UCase$(Left$(rData, 13)) = "/USERSUPPORT " Then
rData = Right$(rData, Len(rData) - 13)
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then
        UserList(userindex).flags.Privilegios = PlayerType.UserSupport
        Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
        Exit Sub
    End If
    
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@soporte de usuarios")
    UserList(tIndex).flags.Privilegios = PlayerType.UserSupport
    Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If

If UCase$(Left$(rData, 14)) = "/EVENTMANAGER " Then
rData = Right$(rData, Len(rData) - 14)
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then
        UserList(userindex).flags.Privilegios = PlayerType.EventManager
        Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
        Exit Sub
    End If
    
Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@manager de eventos")
UserList(tIndex).flags.Privilegios = PlayerType.EventManager
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If


If UCase$(Left$(rData, 19)) = "/TOURNAMENTMANAGER " Then
rData = Right$(rData, Len(rData) - 19)
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then
        UserList(userindex).flags.Privilegios = PlayerType.TournamentManager
        Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
        Exit Sub
    End If
    
Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@manager de torneos")
UserList(tIndex).flags.Privilegios = PlayerType.TournamentManager
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If

If UCase$(Left$(rData, 14)) = "/COORDINATION " Then
rData = Right$(rData, Len(rData) - 14)
If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then
        UserList(userindex).flags.Privilegios = PlayerType.Coordination
        Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
        Exit Sub
    End If
    
Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@coordinador de gms")
UserList(tIndex).flags.Privilegios = PlayerType.Coordination
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If

If UCase$(Left$(rData, 13)) = "/DEVELOPMENT " Then
rData = Right$(rData, Len(rData) - 13)
If UserList(userindex).flags.Privilegios < PlayerType.Administrator Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then
        UserList(userindex).flags.Privilegios = PlayerType.Development
        Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
        Exit Sub
    End If
    
Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@desarrollador")
UserList(tIndex).flags.Privilegios = PlayerType.Development
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If

If UCase$(Left$(rData, 15)) = "/ADMINISTRATOR " Then
rData = Right$(rData, Len(rData) - 15)
If UserList(userindex).flags.Privilegios < PlayerType.Administrator Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@administrador")
UserList(tIndex).flags.Privilegios = PlayerType.Administrator
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/PJ " Then
rData = Right$(rData, Len(rData) - 4)
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then
        UserList(userindex).flags.Privilegios = PlayerType.User
        Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
        Exit Sub
    End If
    
Call SendData(SendTarget.ToAdmins, userindex, 0, "||562@" & UserList(userindex).Name & "@" & UserList(tIndex).Name & "@personaje")
UserList(tIndex).flags.Privilegios = PlayerType.User
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    
Exit Sub
End If


If UCase$(Left$(rData, 12)) = "/CHANGENICK " Then
rData = Right$(rData, Len(rData) - 12)

If Not UserList(userindex).flags.Privilegios = PlayerType.Administrator Then Exit Sub

    UserList(userindex).Name = rData
    Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, False)
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/HECHIZO " Then
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = UCase$(Right$(rData, Len(rData) - 9))
    tStr = Replace$(ReadField(1, rData, 32), "+", " ")
    tIndex = NameIndex(tStr)
    Arg1 = ReadField(2, rData, 32)
    
        Dim j As Integer
        If Not TieneHechizo(Arg1, tIndex) Then
            'Buscamos un slot vacio
            For j = 1 To MAXUSERHECHIZOS
                If UserList(tIndex).Stats.UserHechizos(j) = 0 Then Exit For
            Next j
                
            If UserList(tIndex).Stats.UserHechizos(j) <> 0 Then
                Exit Sub
            Else
                UserList(tIndex).Stats.UserHechizos(j) = Arg1
                Call UpdateUserHechizos(False, tIndex, CByte(j))
            End If
        End If
    
Exit Sub
End If

'Mensaje del sistema
If UCase$(Left$(rData, 6)) = "/SMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call LogGM(UserList(userindex).Name, "Mensaje de sistema:" & rData, False)
    Call SendData(SendTarget.ToAll, 0, 0, "!!" & rData & ENDC)
    Exit Sub
End If

'Crear criatura, toma directamente el indice
If UCase$(Left$(rData, 5)) = "/ACC " Then
   rData = Right$(rData, Len(rData) - 5)
   If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
   
   Call LogGM(UserList(userindex).Name, "Sumoneo a " & Npclist(val(rData)).Name & " en mapa " & UserList(userindex).Pos.Map, False)
   Call LogGMss(UserList(userindex).Name, "Sumoneo a " & Npclist(val(rData)).Name & " en mapa " & UserList(userindex).Pos.Map, False)
   Call SpawnNpc(val(rData), UserList(userindex).Pos, True, False)
   Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rData, 6)) = "/RACC " Then
   rData = Right$(rData, Len(rData) - 6)
   
   If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
   Call LogGM(UserList(userindex).Name, "Sumoneo con respawn " & Npclist(val(rData)).Name & " en mapa " & UserList(userindex).Pos.Map, False)
   Call LogGMss(UserList(userindex).Name, "Sumoneo con respawn " & Npclist(val(rData)).Name & " en mapa " & UserList(userindex).Pos.Map, False)
   Call SpawnNpc(val(rData), UserList(userindex).Pos, True, True)
   Exit Sub
End If

If UCase$(rData) = "/HABILITAR" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    If ServerSoloGMs > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||563")
        ServerSoloGMs = 0
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||564")
        ServerSoloGMs = 1
    End If
    Exit Sub
End If

If UCase(Left(rData, 10)) = "/BORRARPJ " Then
rData = Right$(rData, Len(rData) - 10)
                        
        UserName = rData
        rData = ReadField(1, rData, Asc(","))
        archivo = App.Path & "\Accounts\" & GetVar(App.Path & "\Charfile\" & UserName & ".chr", "CHAR", "Cuenta") & ".act"
        
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    
    If UCase$(UserList(userindex).Name) <> "KING" And UCase$(UserList(NameIndex(rData)).Name) = "KING" Then Exit Sub
    
    Dim IndexUserBorrado As Long
    IndexUserBorrado = NameIndex(rData)
    
    If IndexUserBorrado > 0 Then
      CloseSocket (IndexUserBorrado)
    End If
        
        NumPjs = CByte(val(GetVar(archivo, "PJS", "NumPjs")))

                        For i = 1 To val(GetVar(archivo, "PJS", "NumPjs"))
                            If UCase$(GetVar(archivo, "PJS", "PJ" & i)) = UCase$(rData) Then
                                    Call WriteVar(archivo, "PJS", "PJ" & i, "")
                                    limitPJ = i + 1
                                    BorrarUsuario (rData)
                                    If i = 0 Then
                                    Exit For
                                Else
                                    Call WriteVar(archivo, "PJS", "NumPjs", val(GetVar(archivo, "PJs", "NumPjs")) - 1)
                                Exit For
                            End If
                            End If
                        Next i
                     
                        For i = limitPJ To NumPjs
                            UserName = GetVar(archivo, "PJS", "PJ" & i)
                            Call WriteVar(archivo, "PJS", "PJ" & i, "")
                            Call WriteVar(archivo, "PJS", "PJ" & i - 1, UserName)
                        Next i
                        
        Call SendData(SendTarget.ToAll, 0, 0, "||565@" & UserList(userindex).Name & "@" & rData)
        Exit Sub
End If

If UCase(Left(rData, 7)) = "/BANHD " Then
rData = Right$(rData, Len(rData) - 7)
Dim banhd As Integer
Dim userinx As Integer

userinx = NameIndex(rData)

If userinx = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||196")
Exit Sub
End If

banhd = val(NameIndex(rData))
BanIP = UserList(userinx).ip
Dim hdbanned As String
hdbanned = val(UserList(banhd).hd)
NombreCuent = UserList(userinx).Accounted

If UCase$(rData) = "KING" Then Exit Sub
If UserList(userindex).flags.Privilegios < PlayerType.Administrator Then Exit Sub

UserList(userinx).flags.Ban = 1

tInt = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", tInt + 1)
Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & tInt + 1, "Tolerancia 0. " & Date & " " & Time)
Call WriteVar(App.Path & "\Accounts\" & UserList(userinx).Accounted & ".act", NombreCuent, "ban", "1")

    If CheckHD(hdbanned) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||566@" & hdbanned)
        UserList(banhd).flags.Ban = 1
        Call WriteVar(CharPath & rData & ".chr", "FLAGS", "Ban", "1")
    Exit Sub
    Else
        Open "" & App.Path & "\DAT\BanHds.dat" For Append As #1
        Print #1, hdbanned
        Close #1
        
        Call BanIpAgrega(BanIP)
        
        Call CloseSocket(banhd)
        'Avisamos
        Call SendData(SendTarget.ToAll, 0, 0, "||567@" & UserList(userindex).Name & "@" & rData)
        Call SendData(SendTarget.ToAll, 0, 0, "||568@" & UserList(userindex).Name & "@" & hdbanned)
        Call SendData(SendTarget.ToAll, 0, 0, "||569@" & UserList(userindex).Name & "@" & BanIP)
        Call SendData(SendTarget.ToAll, 0, 0, "||570@" & UserList(userindex).Name & "@" & NombreCuent)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/MOD " Then
    rData = UCase$(Right$(rData, Len(rData) - 5))
    tIndex = userindex
    Arg1 = ReadField(1, rData, 32)
    Arg2 = ReadField(2, rData, 32)
    Arg3 = ReadField(3, rData, 32)
    Arg4 = ReadField(4, rData, 32)
   
    Call LogGM(UserList(userindex).Name, rData, False)
   
    Select Case Arg1
    
     Case "PART"
           
            If val(Arg2) <= 0 Then
                Exit Sub
            End If
           
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & Arg2 & "," & 0)
            
    Case "AURA"
           
            If val(Arg2) <= 0 Then
                Exit Sub
            End If
           
            UserList(userindex).Char.AuraA = Arg2
            SendUserAura (userindex)
            
    Case "FX"
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & val(Arg2) & "," & 20)


    Case "ATRI"
        UserList(userindex).Stats.UserAtributos(1) = val(Arg2)
        UserList(userindex).Stats.UserAtributos(2) = val(Arg2)
        UserList(userindex).Stats.UserAtributos(3) = val(Arg2)
        UserList(userindex).Stats.UserAtributos(4) = val(Arg2)
        UserList(userindex).Stats.UserAtributos(5) = val(Arg2)
        
        Call SendData(SendTarget.toindex, userindex, 0, "||571@" & val(Arg2))
    
    Case "ORO" '/mod yo oro 95000
           
            If val(Arg2) < 321999999999999# Then
                UserList(tIndex).Stats.GLD = val(Arg2)
                Call SendUserGLD(tIndex)
                Exit Sub
            End If
        Case "EXP" '/mod yo exp 9995000
           
            If UserList(userindex).flags.EsRolesMaster Then Exit Sub
            If val(Arg2) < 321999999999999# Then
                If UserList(tIndex).Stats.Exp + val(Arg2) > _
                   UserList(tIndex).Stats.ELU Then
                   Dim resto
                   resto = val(Arg2) - UserList(tIndex).Stats.ELU
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + UserList(tIndex).Stats.ELU
                   Call CheckUserLevel(tIndex)
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                Else
                   UserList(tIndex).Stats.Exp = val(Arg2)
                End If
                Call SendUserEXP(tIndex)
                Call SendData(toindex, userindex, 0, "||572@" & val(Arg2))
                Exit Sub
            End If
        Case "BODY"
            Call SendData(toindex, userindex, 0, "||573@" & val(Arg2))
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            Exit Sub
        Case "HEAD"
            Call SendData(toindex, userindex, 0, "||574@" & val(Arg2))
            UserList(userindex).Char.Head = val(Arg2)
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            Exit Sub
        Case "CRI"
            Call SendData(toindex, userindex, 0, "||575@" & val(Arg2))
            If UserList(userindex).flags.EsRolesMaster Then Exit Sub
            UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
            Exit Sub
        Case "CIU"
           
            Call SendData(toindex, userindex, 0, "||576@" & val(Arg2))
            If UserList(userindex).flags.EsRolesMaster Then Exit Sub
            UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
            Exit Sub
        Case "LEVEL"
                        
            Call SendData(toindex, userindex, 0, "||577@" & val(Arg2))
            UserList(tIndex).Stats.ELV = val(Arg2)
            Call SendUserLVL(tIndex)
            Exit Sub
        Case "CLASE"
            
            Call SendData(toindex, userindex, 0, "||578@" & UCase$(Arg2))
            UserList(tIndex).clase = UCase$(Arg2)
       
        Case "HAM"
            UserList(tIndex).Stats.MinHam = val(Arg2)
            UserList(tIndex).Stats.MaxHam = val(Arg2)
            Call SendData(toindex, userindex, 0, "||579@" & val(Arg2))
            Call EnviarHambreYsed(tIndex)
            Exit Sub
           
        Case "AGU"
            UserList(tIndex).Stats.MinAGU = val(Arg2)
            UserList(tIndex).Stats.MaxAGU = val(Arg2)
            Call SendData(toindex, userindex, 0, "||580@" & val(Arg2))
            Call EnviarHambreYsed(tIndex)
            Exit Sub
           
         Case "STA"
            UserList(tIndex).Stats.MinSta = val(Arg2)
            UserList(tIndex).Stats.MaxSta = val(Arg2)
            Call SendData(toindex, userindex, 0, "||581@" & val(Arg2))
            Call SendUserST(tIndex)
            Exit Sub
 
        Case "MP"
            UserList(tIndex).Stats.MinMAN = val(Arg2)
            UserList(tIndex).Stats.MaxMAN = val(Arg2)
            Call SendData(toindex, userindex, 0, "||582@" & val(Arg2))
            Call SendUserMP(tIndex)
            Exit Sub
        Case "HP"
            UserList(tIndex).Stats.MinHP = val(Arg2)
            UserList(tIndex).Stats.MaxHP = val(Arg2)
            Call SendData(toindex, userindex, 0, "||583@" & val(Arg2))
            Call SendUserHP(tIndex)
            Exit Sub
            
        Case "ESCU"
            UserList(userindex).Char.ShieldAnim = val(Arg2)
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        
        Case "CASCO"
            UserList(userindex).Char.CascoAnim = val(Arg2)
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            
        Case "ARMA"
            UserList(userindex).Char.WeaponAnim = val(Arg2)
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        Exit Sub
    '[/DnG]
        Case Else
            Call SendData(toindex, userindex, 0, "||584")
            Exit Sub
        End Select
 
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/RESETVALS " Then
    rData = UCase$(Right$(rData, Len(rData) - 11))
    Arg1 = ReadField(1, rData, 32)
    
    If UCase$(Arg1) <> "ARENA1" And UCase$(Arg1) <> "ARENA2" And UCase$(Arg1) <> "ARENA3" And UCase$(Arg1) <> "ARENA4" And UCase$(Arg1) <> "2VS2" And UCase$(Arg1) <> "DESAFIO" And UCase$(Arg1) <> "CVC" Then
        Call SendData(toindex, userindex, 0, "||585")
     Exit Sub
    End If
    
Select Case Arg1
  Case "DESAFIO"
        If Desafio.Primero <> 0 Then
        UserList(Desafio.Primero).flags.EnDesafio = 0
        UserList(Desafio.Primero).flags.rondas = 0
        Desafio.Primero = 0
        End If
        
        If Desafio.Segundo <> 0 Then
        UserList(Desafio.Segundo).flags.EnDesafio = 0
        UserList(Desafio.Segundo).flags.rondas = 0
        Desafio.Segundo = 0
        End If
        
        Call SendData(toindex, userindex, 0, "||586")
    Exit Sub
    
   Case "ARENA1"
    Dim jjj As Long
        For jjj = 1 To LastUser
         If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 1 Then
                UserList(jjj).flags.EnDuelo = False
                UserList(jjj).flags.DueliandoContra = ""
                UserList(jjj).flags.LeMandaronDuelo = False
                UserList(jjj).flags.UltimoEnMandarDuelo = ""
                UserList(jjj).flags.EnQueArena = 0
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
         End If
         
            If UserList(jjj).flags.EspectadorArena1 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena1 = 0
                EspectadoresEnArena1 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||587@1"
    ArenaOcupada(1) = False
    NombreDueleando(1) = ""
    NombreDueleando(2) = ""
    
   Case "ARENA2"
    For jjj = 1 To LastUser
     If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 2 Then
            UserList(jjj).flags.EnDuelo = False
            UserList(jjj).flags.DueliandoContra = ""
            UserList(jjj).flags.LeMandaronDuelo = False
            UserList(jjj).flags.UltimoEnMandarDuelo = ""
            UserList(jjj).flags.EnQueArena = 0
            WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
     End If
     
            If UserList(jjj).flags.EspectadorArena2 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena2 = 0
                EspectadoresEnArena2 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||587@2"
    ArenaOcupada(2) = False
    NombreDueleando(3) = ""
    NombreDueleando(4) = ""
   
   Case "ARENA3"
    For jjj = 1 To LastUser
     If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 3 Then
            UserList(jjj).flags.EnDuelo = False
            UserList(jjj).flags.DueliandoContra = ""
            UserList(jjj).flags.LeMandaronDuelo = False
            UserList(jjj).flags.UltimoEnMandarDuelo = ""
            UserList(jjj).flags.EnQueArena = 3
            WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
     End If
     
            If UserList(jjj).flags.EspectadorArena3 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena3 = 0
                EspectadoresEnArena3 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||587@3"
    ArenaOcupada(3) = False
    NombreDueleando(5) = ""
    NombreDueleando(6) = ""
 
   Case "ARENA4"
    For jjj = 1 To LastUser
     If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 4 Then
            UserList(jjj).flags.EnDuelo = False
            UserList(jjj).flags.DueliandoContra = ""
            UserList(jjj).flags.LeMandaronDuelo = False
            UserList(jjj).flags.UltimoEnMandarDuelo = ""
            UserList(jjj).flags.EnQueArena = 0
            WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
     End If
     
            If UserList(jjj).flags.EspectadorArena4 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena4 = 0
                EspectadoresEnArena4 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||587@4"
    ArenaOcupada(4) = False
    NombreDueleando(7) = ""
    NombreDueleando(8) = ""
    
   Case "2VS2"
      If Pareja.Jugador(1) > 0 Then
        UserList(Pareja.Jugador(1)).flags.EnPareja = False
        UserList(Pareja.Jugador(1)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(1)).flags.SuPareja = 0
      End If
      
      If Pareja.Jugador(2) > 0 Then
        UserList(Pareja.Jugador(2)).flags.EnPareja = False
        UserList(Pareja.Jugador(2)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(2)).flags.SuPareja = 0
      End If
      
      If Pareja.Jugador(3) > 0 Then
        UserList(Pareja.Jugador(3)).flags.EnPareja = False
        UserList(Pareja.Jugador(3)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(3)).flags.SuPareja = 0
      End If
      
      If Pareja.Jugador(4) > 0 Then
        UserList(Pareja.Jugador(4)).flags.EnPareja = False
        UserList(Pareja.Jugador(4)).flags.EsperaPareja = False
        UserList(Pareja.Jugador(4)).flags.SuPareja = 0
      End If
      
        Pareja.Jugador(1) = 0
        Pareja.Jugador(2) = 0
        Pareja.Jugador(3) = 0
        Pareja.Jugador(4) = 0
        HayPareja = False
     
     Call SendData(toindex, userindex, 0, "||588")
    Exit Sub
    
   Case "CVC"
        CvcFunciona = False
        UsuariosEnCvcClan1 = 0
        UsuariosEnCvcClan2 = 0
        Call SendData(toindex, userindex, 0, "||589")
    Exit Sub
    
   Case "INVOCACIONES"
        Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", val(0))
        Call SendData(toindex, userindex, 0, "||590")
    Exit Sub

End Select
    
Exit Sub
End If

'MODIFICA CARACTER
If UCase$(Left$(rData, 6)) = "/SMOD " Then
If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = UCase$(Right$(rData, Len(rData) - 6))
    tStr = Replace$(ReadField(1, rData, 32), "+", " ")
    tIndex = NameIndex(tStr)
    Arg1 = ReadField(2, rData, 32)
    Arg2 = ReadField(3, rData, 32)
    Arg3 = ReadField(4, rData, 32)
    Arg4 = ReadField(5, rData, 32)
   
    If tIndex <= 0 Then
        Call SendData(toindex, userindex, 0, "||196")
    Exit Sub
    End If
   
    Call LogGM(UserList(userindex).Name, rData, False)
    
    If UCase$(tStr) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then Exit Sub
   
    Select Case Arg1
     Case "PART"
            If val(Arg2) <= 0 Then
                Exit Sub
            End If
           
            Call SendData(SendTarget.ToPCArea, userindex, UserList(tIndex).Pos.Map, "CFF" & UserList(tIndex).Char.CharIndex & "," & Arg2 & "," & 0)
        Case "ORO" '/mod yo oro 95000
           
            If val(Arg2) < 321999999999999# Then
                UserList(tIndex).Stats.GLD = val(Arg2)
                Call SendUserGLD(tIndex)
                 Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@oro@" & UserList(tIndex).Name & "@" & val(Arg2))
                Exit Sub
            End If
        
        Case "EXP" '/mod yo exp 9995000
           
            If val(Arg2) < 321999999999999# Then
                If UserList(tIndex).Stats.Exp + val(Arg2) > _
                   UserList(tIndex).Stats.ELU Then
                   'Dim resto
                   resto = val(Arg2) - UserList(tIndex).Stats.ELU
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + UserList(tIndex).Stats.ELU
                   Call CheckUserLevel(tIndex)
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                Else
                   UserList(tIndex).Stats.Exp = val(Arg2)
                End If
                Call SendUserEXP(tIndex)
                Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@experiencia@" & UserList(tIndex).Name & "@" & val(Arg2))
                Exit Sub
            End If
        Case "BODY"
           
             Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(tIndex).Char.CascoAnim)
             Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@body@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub
        Case "HEAD"
                Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(tIndex).Char.CascoAnim)
                Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@head@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub
        Case "CRI"
            UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@criminales@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub
        Case "CIU"
            UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@ciudadanos@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub
        Case "LEVEL"
            UserList(tIndex).Stats.ELV = val(Arg2)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@nivel@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub
        Case "CLASE"
            UserList(tIndex).clase = UCase$(Arg2)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@clase@" & UserList(tIndex).Name & "@" & Arg2)
 
         Case "STA"
            UserList(tIndex).Stats.MinSta = val(Arg2)
            UserList(tIndex).Stats.MaxSta = val(Arg2)
            Call SendUserST(tIndex)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@energia@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub
        Case "MP"
            UserList(tIndex).Stats.MinMAN = val(Arg2)
            UserList(tIndex).Stats.MaxMAN = val(Arg2)
            Call SendUserMP(tIndex)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@mana@" & UserList(tIndex).Name & "@" & val(Arg2))
        Exit Sub
        Case "HP"
            UserList(tIndex).Stats.MinHP = val(Arg2)
            UserList(tIndex).Stats.MaxHP = val(Arg2)
            Call SendUserHP(tIndex)
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||591@" & UserList(userindex).Name & "@vida@" & UserList(tIndex).Name & "@" & val(Arg2))
            Exit Sub

        Case Else
            Call SendData(toindex, userindex, 0, "||584")
            Exit Sub
        End Select
 
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/COL " Then
rData = Right$(rData, Len(rData) - 5)

Dim Colorx As String

    Colorx = ReadField(1, rData, Asc("@"))
    Arg1 = ReadField(2, rData, Asc("@"))
    
 If UCase$(Colorx) = "LILA" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~200~20~215~1~0")
 ElseIf UCase$(Colorx) = "VERDE" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~0~255~0~1~0")
 ElseIf UCase$(Colorx) = "AZUL" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~0~0~255~1~0")
 ElseIf UCase$(Colorx) = "ROJO" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~255~0~0~1~0")
 ElseIf UCase$(Colorx) = "AMARILLO" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~255~255~0~1~0")
 ElseIf UCase$(Colorx) = "BLANCO" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~255~255~255~1~0")
 ElseIf UCase$(Colorx) = "GRIS" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~120~120~120~1~0")
 ElseIf UCase$(Colorx) = "NARANJA" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~255~128~0~1~0")
 ElseIf UCase$(Colorx) = "MARRON" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~128~64~0~1~0")
 ElseIf UCase$(Colorx) = "CELESTE" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~0~255~255~1~0")
 ElseIf UCase$(Colorx) = "VIOLETA" Then
    Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & Arg1 & "~64~0~128~1~0")
 End If
    
 Exit Sub
End If


Procesado = False
End Sub
