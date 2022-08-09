Attribute VB_Name = "TCP_HandleData4"
Option Explicit

Public Sub HandleData_4(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
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


Select Case UCase$(Left$(rData, 3))
  Case "DPX"
    rData = Right$(rData, Len(rData) - 3)
    Arg2 = ReadField(1, rData, 44)
    
    Dim tItems As Long
    Dim IndexObj As Obj
    Dim NameObj As String

        If val(Arg2) > 0 And val(Arg2) <= UBound(DonationList) Then
                NameObj = ""
                For tItems = 1 To DonationList(val(Arg2)).NumObjs
                
                  If DonationList(val(Arg2)).Tuniquita <> "0" And tItems = 1 Then
                        IndexObj.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        NameObj = NameObj & DonationList(val(Arg2)).Tuniquita & " -" & IndexObj.Amount & ","
                  ElseIf DonationList(val(Arg2)).Tuniquita2 <> "0" And tItems = 2 Then
                        IndexObj.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        NameObj = NameObj & DonationList(val(Arg2)).Tuniquita2 & " -" & IndexObj.Amount & ","
                  ElseIf DonationList(val(Arg2)).Tuniquita3 <> "0" And tItems = 3 Then
                        IndexObj.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        NameObj = NameObj & DonationList(val(Arg2)).Tuniquita3 & " -" & IndexObj.Amount & ","
                  ElseIf DonationList(val(Arg2)).Tuniquita4 <> "0" And tItems = 4 Then
                        IndexObj.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        NameObj = NameObj & DonationList(val(Arg2)).Tuniquita4 & " -" & IndexObj.Amount & ","
                  Else
                        IndexObj.ObjIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        IndexObj.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        
                        If IndexObj.ObjIndex = 9999 Then
                            NameObj = NameObj & "Puntos de Torneo -" & IndexObj.Amount & ","
                        ElseIf IndexObj.ObjIndex = 9998 Then
                            NameObj = NameObj & "Montura de Dragón Rojo -" & IndexObj.Amount & ","
                        ElseIf IndexObj.ObjIndex = 9997 Then
                            NameObj = NameObj & "Montura de Dragón Dorado -" & IndexObj.Amount & ","
                        ElseIf IndexObj.ObjIndex = 9996 Then
                            NameObj = NameObj & "Pack Premium - (" & IndexObj.Amount & " Mes),"
                        Else
                            NameObj = NameObj & ObjData(IndexObj.ObjIndex).Name & "-" & IndexObj.Amount & ","
                        End If
                    End If
                Next tItems
                
                Dim tBody As Integer
                Dim tHead As Integer
                Dim tWeapon As Integer
                Dim tShield As Integer
                Dim tCasco As Integer
                Dim tGrhIndex As Integer
                
                With DonationList(val(Arg2))
                    If .Body > 0 Then
                        tBody = .Body
                        tHead = UserList(userindex).Char.Head
                        
                            
                            If .Arma > 0 Then
                                tWeapon = .Arma
                            Else
                                tWeapon = UserList(userindex).Char.WeaponAnim
                            End If
                            
                            If .Escudo > 0 Then
                                tShield = .Escudo
                            Else
                                tShield = UserList(userindex).Char.ShieldAnim
                            End If
                            
                            If .Casco > 0 Then
                                tCasco = .Casco
                            Else
                                tCasco = UserList(userindex).Char.CascoAnim
                            End If
                    Else
                        tGrhIndex = DonationList(val(Arg2)).GrhIndex
                    End If
                End With
                
            Call SendData(SendTarget.toindex, userindex, 0, "DNF" & tBody & "," & tHead & "," & tWeapon & "," & tShield & "," & tCasco & "," & DonationList(val(Arg2)).Aura & "," & tGrhIndex & "," & DonationList(val(Arg2)).ObjValor & "," & DonationList(val(Arg2)).NumObjs & "," & NameObj)
           End If
    Exit Sub
    
    Case "DRX"
        rData = Right$(rData, Len(rData) - 3)
        Arg2 = ReadField(1, rData, 44)
             'i no tiene los puntos necesarios
             
    Dim tPremio As Obj
    Dim rIndex As Integer
    Dim j As Long
    Dim Alturinha As String
    
            If UCase$(UserList(userindex).Raza) = "GNOMO" Or UCase$(UserList(userindex).Raza) = "ENANO" Then
                Alturinha = "Bajos"
            Else
                Alturinha = "Altos"
            End If
            
                 If UserList(userindex).Stats.PuntosDonacion < DonationList(val(Arg2)).ObjValor Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||632")
                        Exit Sub
                 End If
                 
            For tItems = 1 To DonationList(val(Arg2)).NumObjs
                If DonationList(val(Arg2)).Tuniquita <> "0" And tItems = 1 Then
                    rIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Tunica" & Alturinha), 45))
                ElseIf DonationList(val(Arg2)).Tuniquita2 <> "0" And tItems = 2 Then
                    rIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Tunica2" & Alturinha), 45))
                ElseIf DonationList(val(Arg2)).Tuniquita3 <> "0" And tItems = 3 Then
                    rIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Tunica3" & Alturinha), 45))
                ElseIf DonationList(val(Arg2)).Tuniquita4 <> "0" And tItems = 4 Then
                    rIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Tunica4" & Alturinha), 45))
                Else
                    rIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                End If
                     
                        If rIndex < 9996 Then
                           tPremio.ObjIndex = rIndex
                           tPremio.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                      
                           If Not MeterItemEnInventario(userindex, tPremio) Then
                               Call TirarItemAlPiso(UserList(userindex).Pos, tPremio)
                           End If
                           
                           Call LogCanjeos("" & UserList(userindex).Name & " canjeo: " & tPremio.Amount & " - " & ObjData(tPremio.ObjIndex).Name)
                        
                     ElseIf rIndex = 9996 Then
                        Dim tempDia As Byte, tempMes As Byte, tempAño As Integer
                            Dim tempFecha As String
                            
                        If UserList(userindex).flags.EsPremium = 0 Then
                            tempDia = ReadField(1, Date, Asc("/"))
                            tempMes = ReadField(2, Date, Asc("/"))
                            tempAño = ReadField(3, Date, Asc("/"))
                        Else
                            tempDia = ReadField(1, UserList(userindex).flags.VencePremium, Asc("/"))
                            tempMes = ReadField(2, UserList(userindex).flags.VencePremium, Asc("/"))
                            tempAño = ReadField(3, UserList(userindex).flags.VencePremium, Asc("/"))
                        End If
                        
                            If (tempMes < 12) And (tempDia <= 28) Then
                                tempFecha = "" & tempDia & "/" & tempMes + 1 & "/" & tempAño
                            ElseIf (tempMes < 11) And (tempDia > 28) Then
                                tempFecha = "1/" & tempMes + 2 & "/" & tempAño
                            ElseIf (tempMes = 12) And (tempDia <= 28) Then
                                tempFecha = "" & tempDia & "/1/" & tempAño + 1
                            ElseIf (tempDia > 28) Then
                                If (tempMes = 11) Then tempFecha = "1/1/" & tempAño + 1
                                If (tempMes = 12) Then tempFecha = "1/2/" & tempAño + 1
                            End If
                     
                        UserList(userindex).flags.EsPremium = 1
                        UserList(userindex).flags.VencePremium = tempFecha
                        Call SendData(SendTarget.toindex, userindex, 0, "||893@" & tempFecha)
                     ElseIf rIndex = 9998 Then
                       If Not TieneHechizo(52, userindex) Then
                           'Buscamos un slot vacio
                           For j = 1 To MAXUSERHECHIZOS
                               If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
                           Next j
                               
                           If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
                               Exit Sub
                           Else
                               UserList(userindex).Stats.UserHechizos(j) = 52
                               Call UpdateUserHechizos(False, userindex, CByte(j))
                           End If
                       End If
                       
                       Call SendData(SendTarget.toindex, userindex, 0, "||133")
                    ElseIf rIndex = 9997 Then
                       If Not TieneHechizo(51, userindex) Then
                           'Buscamos un slot vacio
                           For j = 1 To MAXUSERHECHIZOS
                               If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
                           Next j
                               
                           If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
                               Exit Sub
                           Else
                               UserList(userindex).Stats.UserHechizos(j) = 51
                               Call UpdateUserHechizos(False, userindex, CByte(j))
                           End If
                       End If
                  
                  Call SendData(SendTarget.toindex, userindex, 0, "||133")
                        ElseIf rIndex = 9999 Then
                          Call AgregarPuntos(userindex, val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45)))
                          Call SendData(SendTarget.toindex, userindex, 0, "||57@" & val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45)))
                        End If
                       
                      Next tItems
                      
                      'Metemos en inventario
                     Call UpdateUserInv(True, userindex, 0)
                    
                     'Restamos & actualizams
                     UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - DonationList(val(Arg2)).ObjValor
    Exit Sub
    
End Select

Select Case UCase$(Left$(rData, 6))

    Case "DOWNSI"
    rData = Right$(rData, Len(rData) - 6)
    
        tIndex = NameIndex(rData)
    
        If tIndex > 0 Then
            If UserList(userindex).flags.Hechizo = 0 Then Exit Sub
            If (Mod_AntiCheat.PuedoCasteoHechizo(userindex) = False) Then Exit Sub
            
            UserList(userindex).flags.TargetUser = tIndex
            Call LanzarHechizo(UserList(userindex).flags.Hechizo, userindex)
            SendUserData (userindex)
            SendUserData (tIndex)
        Else
            Exit Sub
        End If

    Exit Sub
    
    Case "RANKIN"
        rData = Right$(rData, Len(rData) - 6)
        Arg1 = ReadField(1, rData, 44)
        
        Select Case Arg1
            Case 0
                Call Info_Rank(Kills, userindex)
            Case 1
                Call Info_Rank(Duels, userindex)
            Case 2
                Call Info_Rank(Rounds, userindex)
            Case 3
                Call Info_Rank(Couple, userindex)
            Case 4
                Call Info_Rank(Tournaments, userindex)
            Case 5
                Call Info_Rank(Events, userindex)
            Case 6
                Call Info_Rank(GuildVSGuild, userindex)
            Case 7
                Call Info_Rank(Castles, userindex)
            Case 8
                Call Info_Rank(GuildReputation, userindex)
            Case 9
                Call Info_Rank(Reputation, userindex)
        End Select
    Exit Sub

End Select


Procesado = False
    
End Sub
