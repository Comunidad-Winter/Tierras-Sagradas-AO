Attribute VB_Name = "TCP_HandleData4"
Option Explicit

Public Sub HandleData_4(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


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


Select Case UCase$(Left$(rData, 3))
  Case "DPX"
    rData = Right$(rData, Len(rData) - 3)
    Arg2 = ReadField(1, rData, 44)
    
    Dim tItems As Long
    Dim IndexObj As obj
    Dim NameObj As String

        If val(Arg2) > 0 And val(Arg2) <= UBound(DonationList) Then
                NameObj = ""
                For tItems = 1 To DonationList(val(Arg2)).NumObjs
                
                        IndexObj.ObjIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        IndexObj.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        
                        If IndexObj.ObjIndex = 9999 Then
                            NameObj = NameObj & "Puntos de Torneo -" & IndexObj.Amount & ","
                        ElseIf IndexObj.ObjIndex = 9998 Then
                            NameObj = NameObj & "Montura de Drag�n Rojo -" & IndexObj.Amount & ","
                        ElseIf IndexObj.ObjIndex = 9997 Then
                            NameObj = NameObj & "Montura de Drag�n Dorado -" & IndexObj.Amount & ","
                        ElseIf IndexObj.ObjIndex = 9996 Then
                            NameObj = NameObj & "Pack Premium - (" & IndexObj.Amount & " Mes),"
                        ElseIf IndexObj.ObjIndex = 9995 Then
                            NameObj = NameObj & "Skin de " & ObjData(IndexObj.Amount).Name & ","
                        Else
                            NameObj = NameObj & ObjData(IndexObj.ObjIndex).Name & "-" & IndexObj.Amount & ","
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
                    
                        If (UCase$(UserList(userindex).Raza) = "GNOMO" Or UCase$(UserList(userindex).Raza) = "ENANO") Then
                            tBody = .BodyB
                        Else
                            tBody = .Body
                        End If
                        
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
                            
                            tGrhIndex = DonationList(val(Arg2)).GrhIndex
                    Else
                        tGrhIndex = DonationList(val(Arg2)).GrhIndex
                    End If
                End With
                
            Call SendData(SendTarget.toindex, userindex, 0, "DNF" & tBody & "," & tHead & "," & tWeapon & "," & tShield & "," & tCasco & "," & DonationList(val(Arg2)).Aura & "," & tGrhIndex & "," & DonationList(val(Arg2)).ObjValor & "," & DonationList(val(Arg2)).Desc & "," & DonationList(val(Arg2)).NumObjs & "," & NameObj)
           End If
    Exit Sub
    
    Case "DRX"
        rData = Right$(rData, Len(rData) - 3)
        Arg2 = ReadField(1, rData, 44)
             'i no tiene los puntos necesarios
             
    Dim tPremio As obj
    Dim rIndex As Integer
    Dim j As Long
            
                 If UserList(userindex).Stats.PuntosDonacion < DonationList(val(Arg2)).ObjValor Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||632")
                        Exit Sub
                 End If
                 
            For tItems = 1 To DonationList(val(Arg2)).NumObjs
            
                rIndex = val(ReadField(1, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
            
                    If rIndex < 9995 Then
                    
                        tPremio.ObjIndex = rIndex
                        tPremio.Amount = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                      
                        If Not MeterItemEnInventario(userindex, tPremio) Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||639")
                        Exit Sub
                        End If
                        
                        Call SendData(SendTarget.toindex, userindex, 0, "||232@" & tPremio.Amount & "@" & ObjData(tPremio.ObjIndex).Name)
                           
                        Call LogCanjeos("" & UserList(userindex).Name & " canjeo: " & tPremio.Amount & " - " & ObjData(tPremio.ObjIndex).Name)
                     ElseIf rIndex = 9995 Then
                        Dim tmpRObj, tmpNGraf As Integer
                        tmpRObj = val(ReadField(2, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        tmpNGraf = val(ReadField(3, GetVar(DatPath & "ItemsDonaciones.dat", "ITEM" & val(Arg2), "Obj" & tItems), 45))
                        
                        With UserList(userindex)
                            If .cantSkins > UBound(.Skin) Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||959")
                                Exit Sub
                            End If
                            
                            For i = 1 To .cantSkins
                                If .Skin(i).numObj = tmpRObj Then
                                    Call SendData(SendTarget.toindex, userindex, 0, "||960")
                                Exit Sub
                                End If
                            Next i
                                                            
                            .cantSkins = .cantSkins + 1
                            .Skin(.cantSkins).numObj = tmpRObj
                            .Skin(.cantSkins).newGraf = tmpNGraf
                            
                            Call SendData(SendTarget.toindex, userindex, 0, "||961@" & ObjData(tmpRObj).Name)
                        End With
                                                
                        Call LogCanjeos("" & UserList(userindex).Name & " canjeo skin " & tmpRObj & " - " & tmpNGraf)
                     ElseIf rIndex = 9996 Then
                        Dim tempDia As Byte, tempMes As Byte, tempA�o As Integer
                            Dim tempFecha As String
                            
                        If UserList(userindex).flags.EsPremium = 0 Then
                            tempDia = ReadField(2, Date, Asc("/"))
                            tempMes = ReadField(1, Date, Asc("/"))
                            tempA�o = ReadField(3, Date, Asc("/"))
                        Else
                            tempDia = ReadField(1, UserList(userindex).flags.VencePremium, Asc("/"))
                            tempMes = ReadField(2, UserList(userindex).flags.VencePremium, Asc("/"))
                            tempA�o = ReadField(3, UserList(userindex).flags.VencePremium, Asc("/"))
                        End If
                        
                            If (tempMes < 12) And (tempDia <= 28) Then
                                tempFecha = "" & tempDia & "/" & tempMes + 1 & "/" & tempA�o
                            ElseIf (tempMes < 11) And (tempDia > 28) Then
                                tempFecha = "1/" & tempMes + 2 & "/" & tempA�o
                            ElseIf (tempMes = 12) And (tempDia <= 28) Then
                                tempFecha = "" & tempDia & "/1/" & tempA�o + 1
                            ElseIf (tempDia > 28) Then
                                If (tempMes = 11) Then tempFecha = "1/1/" & tempA�o + 1
                                If (tempMes = 12) Then tempFecha = "1/2/" & tempA�o + 1
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
        
        'Call SendData(SendTarget.toindex, UserIndex, 0, "EROSistema deshabilitado temporalmente."): Exit Sub
        
        Dim Rank As Integer
        
        If UCase(Arg1) = "DUELOS" Then Rank = eRanking.TOPDuelos
        If UCase(Arg1) = "PAREJAS" Then Rank = eRanking.TOPParejas
        If UCase(Arg1) = "RONDAS" Then Rank = eRanking.TOPRondas
        If UCase(Arg1) = "REPUTACION" Then Rank = eRanking.TOPReputacion
        If UCase(Arg1) = "TORNEOS" Then Rank = eRanking.TOPTorneos
        If UCase(Arg1) = "CVCS" Then Rank = eRanking.TOPCVCS
        If UCase(Arg1) = "CASTILLOS" Then Rank = eRanking.TOPCastillos
        If UCase(Arg1) = "REPUCLANES" Then Rank = eRanking.TOPRepuClanes
        If UCase(Arg1) = "FRAGS" Then Rank = eRanking.TOPFrags
        
        tStr = ""
            For i = 1 To 10
                If (Ranking(Rank).Nombre(i) <> "") Then
                    tStr = tStr & Ranking(Rank).Nombre(i) & "-" & Ranking(Rank).Value(i) & ","
                Else
                    tStr = tStr & "N/A-0,"
                End If
            Next i
                
                Call SendData(SendTarget.toindex, userindex, 0, "MTOP" & tStr)
    Exit Sub

End Select


Procesado = False
    
End Sub
