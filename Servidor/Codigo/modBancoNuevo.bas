Attribute VB_Name = "modBancoNuevo"
Sub BIniciarDeposito(ByVal userindex As Integer)
On Error GoTo Errhandler

'Hacemos un Update del inventario del usuario
Call BUpdateBanUserInv(True, userindex, 0)
'Atcualizamos el dinero
Call SendUserGLD(userindex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData SendTarget.toindex, userindex, 0, "INITCBANK"

UserList(userindex).flags.Comerciando = True

Errhandler:

End Sub

Sub BSendBanObj(userindex As Integer, Slot As Byte, Object As UserOBJ)


UserList(userindex).BancoInventB.Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(SendTarget.toindex, userindex, 0, "SBG" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).OBJType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef & "," & GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO") & "," & UserList(userindex).Stats.GLD)

Else

    Call SendData(SendTarget.toindex, userindex, 0, "SBG" & Slot & "," & "0" & "," & "(Nada)" & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & GetVar(App.Path & "\guilds\Bancos\" & Guilds(UserList(userindex).GuildIndex).GuildName & ".bov", "" & Guilds(UserList(userindex).GuildIndex).GuildName & "", "BANCO") & "," & UserList(userindex).Stats.GLD)

End If


End Sub

Sub BUpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).BancoInventB.Object(Slot).ObjIndex > 0 Then
        Call BSendBanObj(userindex, Slot, UserList(userindex).BancoInventB.Object(Slot))
    Else
        Call BSendBanObj(userindex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).BancoInventB.Object(LoopC).ObjIndex > 0 Then
            Call BSendBanObj(userindex, LoopC, UserList(userindex).BancoInventB.Object(LoopC))
        Else
            
            Call BSendBanObj(userindex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub BUserRetiraItem(ByVal userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo Errhandler


If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub

If Cantidad < 1 Then Exit Sub

Call SendUserGLD(userindex)

   
       If UserList(userindex).BancoInventB.Object(i).Amount > 0 Then
            If Cantidad > UserList(userindex).BancoInventB.Object(i).Amount Then Cantidad = UserList(userindex).BancoInventB.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call BUserReciveObj(userindex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos el banco
            Call BUpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana de comercio
            Call BUpdateVentanaBanco(i, 0, userindex)
       End If



Errhandler:

End Sub

Sub BUserReciveObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer


If UserList(userindex).BancoInventB.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = UserList(userindex).BancoInventB.Object(ObjIndex).ObjIndex


'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = obji And _
   UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > UserList(userindex).InventorySlots Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > UserList(userindex).InventorySlots Then
        Slot = 1
        Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > UserList(userindex).InventorySlots Then
                Call SendData(SendTarget.toindex, userindex, 0, "||108")
                Exit Sub
            End If
        Loop
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If

'Mete el obj en el slot
If UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(userindex).Invent.Object(Slot).ObjIndex = obji
    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + Cantidad
    
    Call BQuitarBancoInvItem(userindex, CByte(ObjIndex), Cantidad)
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||108")
End If


End Sub

Sub BQuitarBancoInvItem(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = UserList(userindex).BancoInventB.Object(Slot).ObjIndex

    'Quita un Obj

       UserList(userindex).BancoInventB.Object(Slot).Amount = UserList(userindex).BancoInventB.Object(Slot).Amount - Cantidad
        
        If UserList(userindex).BancoInventB.Object(Slot).Amount <= 0 Then
            UserList(userindex).BancoInventB.NroItems = UserList(userindex).BancoInventB.NroItems - 1
            UserList(userindex).BancoInventB.Object(Slot).ObjIndex = 0
            UserList(userindex).BancoInventB.Object(Slot).Amount = 0
        End If

    
    
End Sub

Sub BUpdateVentanaBanco(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal userindex As Integer)
 
 Call SendData(SendTarget.toindex, userindex, 0, "BANCOBK" & Slot & "," & NpcInv)
 
End Sub

Sub BUserDepositaItem(ByVal userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo Errhandler

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub

'El usuario deposita un item
Call SendUserGLD(userindex)
   
If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call BUserDejaObj(userindex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos el inventario del banco
            Call BUpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana del banco
            
            Call BUpdateVentanaBanco(Item, 1, userindex)
            
End If

Errhandler:

End Sub

Sub BUserDejaObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub

If Cantidad < 1 Then Exit Sub

obji = UserList(userindex).Invent.Object(ObjIndex).ObjIndex

If ObjData(UserList(userindex).Invent.Object(ObjIndex).ObjIndex).Intransferible = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||185")
Exit Sub
End If

'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(userindex).BancoInventB.Object(Slot).ObjIndex = obji And _
         UserList(userindex).BancoInventB.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
        
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(userindex).BancoInventB.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(SendTarget.toindex, userindex, 0, "||186")
                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(userindex).BancoInventB.NroItems = UserList(userindex).BancoInventB.NroItems + 1
        
        
End If

If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(userindex).BancoInventB.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(userindex).BancoInventB.Object(Slot).ObjIndex = obji
        UserList(userindex).BancoInventB.Object(Slot).Amount = UserList(userindex).BancoInventB.Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||186")
    End If

Else
    Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
End If

End Sub

