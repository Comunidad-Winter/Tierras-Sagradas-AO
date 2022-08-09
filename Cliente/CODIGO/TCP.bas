Attribute VB_Name = "Mod_TCP"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer


Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean

Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim charindex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    
    Dim tStr As String
    Dim tstr2 As String
    
    If left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)
    If left$(sData, 4) = "ARAM" Then AramSeconds = Right$(sData, Len(sData) - 4)
    
        Debug.Print "Recibido: " & sData
    
    Select Case sData
        Case "MUERT"
          If Configuracion.CartelMuerte = 1 Then
            frmMuertito.Show , frmMain
          End If
        Exit Sub
    
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
            AlphaY = 130
            ISItem = True
            mode = True
            logged = True
            EngineRun = True
            UserDescansar = False
            Nombres = True
            IsSeguroC = True
            
            If IsSeguroC = True Then
              frmMain.PicSeg.Visible = True
            Else
              frmMain.PicSeg.Visible = False
            End If
            
            If ISItem = True Then
             frmMain.PicItemSeg.Visible = True
            Else
             frmMain.PicItemSeg.Visible = False
            End If
            
           If frmCrearPersonaje.Visible Then
                Unload frmCrearPersonaje
                Unload frmConnect
                frmMain.Show
            End If
            Call SetConnected

            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.RemoveAllDialogs
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
        Exit Sub
        Case "FINOK" ' Graceful exit ;))
            #If UsarWrench = 1 Then
                        frmMain.Socket1.Disconnect
            #Else
                        If frmMain.Winsock1.State <> sckClosed Then _
                            frmMain.Winsock1.Close
            #End If
            frmMain.Visible = False
            logged = False
            UserParalizado = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            charlist(UserCharIndex).color = 0
            
            Call CambiarConectar("CONECTAR")
            
            Call Audio.StopWave
            bFogata = False
            SkillPoints = 0
            Call Dialogos.RemoveAllDialogs
            For i = 1 To LastChar
                charlist(i).invisible = False
            Next i
            
            bK = 0
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                frmConnect.MousePointer = 1
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        Case "FINCBNOK"          ' >>>>> Finaliza Cuenta Bancaria :: FINCBNOK
            frmNuevoBancoObj.List1(0).Clear
            frmNuevoBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmNuevoBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANKO"
            frmBanco.Show , frmMain
        Exit Sub
        Case "INITSUB"           ' >>>>> Inicia Subasta :: #Fer
            i = 1
            frmSubastar.ItemList.Clear
           
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmSubastar.ItemList.AddItem Inventario.ItemName(i)
                Else
                        frmSubastar.ItemList.AddItem "Nada"
                End If
                i = i + 1
            Loop
            frmSubastar.Show , frmMain
        Exit Sub
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim ii As Integer
            ii = 1
            Do While ii <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(ii) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(ii)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
        Exit Sub
        Case "INITCBANK"           ' >>>>> Inicia cuenta bancaria.
            ii = 1
            Do While ii <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(ii) <> 0 Then
                        frmNuevoBancoObj.List1(1).AddItem "" & Inventario.ItemName(ii) & " - " & Inventario.Amount(ii) & ""
                Else
                        frmNuevoBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventoryB)
                If UserBancoInventoryB(i).OBJIndex <> 0 Then
                        frmNuevoBancoObj.List1(0).AddItem "" & UserBancoInventoryB(i).Name & " - " & UserBancoInventoryB(i).Amount & ""
                Else
                        frmNuevoBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            
                frmNuevoBancoObj.OroBove.Text = PonerPuntos(UserBancoOro)
                frmNuevoBancoObj.MiOro.Text = PonerPuntos(UserBancoOroPropio)
            
            Comerciando = True
            Unload frmNuevoBanco
            frmNuevoBancoObj.Show , frmMain
        Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
       Case "BORROK"
            
        Mensaje.Escribir "El personaje ha sido borrado."
        #If UsarWrench = 1 Then
                    frmMain.Socket1.Disconnect
        #Else
                    If frmMain.Winsock1.State <> sckClosed Then _
                        frmMain.Winsock1.Close
        #End If
        
            Call CambiarConectar("CONECTAR")

        Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "HOLASOYUNCIRUJA"
            TimerPing(2) = GetTickCount()
            Dim cuantolagtengoxd As String
            
            If TimerPing(2) - TimerPing(1) > 0 And TimerPing(2) - TimerPing(1) < 100 Then
            cuantolagtengoxd = "0 Lag"
            ElseIf TimerPing(2) - TimerPing(1) > 100 And TimerPing(2) - TimerPing(1) < 200 Then
            cuantolagtengoxd = "Bajo"
            ElseIf TimerPing(2) - TimerPing(1) > 200 And TimerPing(2) - TimerPing(1) < 400 Then
            cuantolagtengoxd = "Medio"
            ElseIf TimerPing(2) - TimerPing(1) > 400 And TimerPing(2) - TimerPing(1) < 900 Then
            cuantolagtengoxd = "Alto"
            ElseIf TimerPing(2) - TimerPing(1) > 900 Then
            cuantolagtengoxd = "Injugable"
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, "<<Recibido: el ping es de " & TimerPing(2) - TimerPing(1) & " Mili-Segundos (" & (TimerPing(2) - TimerPing(1)) / 1000 & " Seg) LAG: " & cuantolagtengoxd, 0, 255, 0, True, False, False)
        Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
            Case "SEGONR" ' <--- Activa el seguro de resi
Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, False)
Exit Sub
Case "SEGOFR" ' <--- Desactiva el seguro de resu
Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, False)
Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            IsSeguroC = False
            frmMain.PicSeg.Visible = True
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            IsSeguroC = False
            frmMain.PicSeg.Visible = False
            Exit Sub
    End Select

Select Case left(sData, 1)
        Case "+"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)


            charindex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))

    With charlist(charindex)
            
            
        For i = 1 To 3
            If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                .FxIndex(i) = 0
                .Fx(i).Loops = 0
            End If
        Next i
            
        ' Play steps sounds if the user is not an admin of any kind
        If .priv = 0 Then
            Call DoPasosFx(charindex)
        End If
    End With

            Call engine.Char_Move_by_Pos(charindex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            
            charindex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))
            
    With charlist(charindex)
    
        For i = 1 To 3
            If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                .FxIndex(i) = 0
                .Fx(i).Loops = 0
            End If
        Next i
    
        ' Play steps sounds if the user is not an admin of any kind
        If .priv = 0 Then
            Call DoPasosFx(charindex)
        End If
    End With
    
            Call engine.Char_Move_by_Pos(charindex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
    
    End Select

    Select Case left$(sData, 2)
    
        Case "99"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            frmBonificadores.lblBeneficio(0) = ReadField(1, Rdata, 44)
            frmBonificadores.lblBeneficio(1) = ReadField(2, Rdata, 44)
            frmBonificadores.Show , frmMain
        Exit Sub
        
        Case "CU"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim CunT As Byte
            CunT = Val(Rdata)
            If CunT = 0 Then
            Cuenta = False
            Tiempo = 45
            Conteo = 27850
            ConteoH = GrhData(Conteo).pixelHeight
            ConteoW = GrhData(Conteo).pixelWidth
            TransparenciaCont = 220
            frmMain.Timer1.Enabled = True
            ElseIf CunT < 6 Then
            Conteo = 27850 + CunT
            ConteoH = GrhData(Conteo).pixelHeight
            ConteoW = GrhData(Conteo).pixelWidth
            TransparenciaCont = 220
            frmMain.Timer1.Enabled = True
            Cuenta = True
            If CunT = 0 Then Cuenta = False And Tiempo = 45
            Else
            Cuenta = False
            End If
        Exit Sub
        
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            
            Call DibujarMinimap
            
            If FileExist(App.Path & "\Data\MAPAS\" & "Mapa" & UserMap & ".map", vbNormal) Then
                Open App.Path & "\Data\MAPAS\" & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
                
                day_r_old = Val(ReadField(2, Rdata, 44))
                day_g_old = Val(ReadField(3, Rdata, 44))
                day_b_old = Val(ReadField(4, Rdata, 44))
                base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
                
                If day_r_old > 0 Or day_g_old > 0 Or day_b_old > 0 Then
                    TieneColorMapa = True
                Else
                    TieneColorMapa = False
                End If
                
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call UnloadAllForms
                End
            End If
            Exit Sub
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).charindex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            Call DibujarPuntoMinimap
            frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
       Case "RT"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
           
            If Rdata <> vbNullString Then
               Call RenderGM.Create(Rdata)
            End If
        Case "||"                 ' >>>>> Nuevo dialogo :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim IDText As Long
            Dim DatoAdd(1 To 8) As String 'Datos adicionales
            
            IDText = Val(ReadField(1, Rdata, Asc("@")))
            tStr = Messages(IDText).Text
            
            'Reemplazo los datos adicionales
            For i = 1 To 8
                DatoAdd(i) = ReadField(1 + i, Rdata, Asc("@"))
                If DatoAdd(i) = vbNullString Then Exit For
            
                tStr = Replace(tStr, "%" & i, DatoAdd(i))
            Next i
                
            'Tiramos el texto a la consola
            AddtoRichTextBox frmMain.RecTxt, tStr, FontTypes(Messages(IDText).font).r, FontTypes(Messages(IDText).font).g, FontTypes(Messages(IDText).font).b, FontTypes(Messages(IDText).font).bold, FontTypes(Messages(IDText).font).italic
        Exit Sub
        Case "N|"                 ' >>>>> Dialogo de Usuarios y NPCs ::    N|
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CreateDialog ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
                If Configuracion.Mensajes = 1 Then AddtoRichTextBox frmMain.RecTxt, ReadField(2, Rdata, 176), 255, 255, 255, 0, 0
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

        Exit Sub
        Case "P|"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            If Configuracion.Desactivar_Privados = 0 Then
                AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                AddtoRichTextBox frmMain.PrivatesConsole, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            End If
        Exit Sub
        Case "C|"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            AddtoRichTextBox frmMain.ClanConsole, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
        Exit Sub
        Case "G|"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            If Configuracion.Desactivar_Globales = 0 Then
                AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                AddtoRichTextBox frmMain.GlobalConsole, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            End If
        Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                Mensaje.Escribir Rdata
            End If
            Exit Sub
            
        Case "ON"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            frmMain.ONLINES.Caption = Rdata
        Exit Sub
        
        Case "LK" ' >>>>> newbie
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(1, Rdata, 44)
            charlist(charindex).esNW = Val(ReadField(2, Rdata, 44))
        Case "XC"              ' >>>>> Nombres :: XC - Actualizamos todo a un solito paquete
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).color = Val(ReadField(2, Rdata, 44))
        Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
            Call DibujarPuntoMinimap
            frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
        Exit Sub
        Case "CC" ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
             
               
            'charlist(CharIndex).FxIndex = Val(ReadField(9, Rdata, 44))
            'charlist(CharIndex).Fx.Loops = Val(ReadField(10, Rdata, 44))
            charlist(charindex).Nombre = ReadField(12, Rdata, 44)
            charlist(charindex).EsStatus = Val(ReadField(13, Rdata, 44))
            charlist(charindex).priv = Val(ReadField(14, Rdata, 44))
            charlist(charindex).NPCAura = Val(ReadField(15, Rdata, 44))
            charlist(charindex).NPCNumber = Val(ReadField(16, Rdata, 44))
            
            Call InitGrh(charlist(charindex).NPCAuraG, AurasPJ(charlist(charindex).NPCAura).GrhIndex)
            charlist(charindex).NPCAuraAngle = 0
            
            If Val(ReadField(2, Rdata, 44)) = 500 Or Val(ReadField(2, Rdata, 44)) = 501 Or Val(ReadField(2, Rdata, 44)) = 511 Or Val(ReadField(2, Rdata, 44)) = 511 Or Val(ReadField(2, Rdata, 44)) = 512 Then
                charlist(charindex).Muerto = True
            End If
            
            'Guardamos
            Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
            Call RefreshAllChars
        Exit Sub
        Case "NF"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(1, Rdata, 44)
            charlist(charindex).SinEnlistarHorda = Val(ReadField(2, Rdata, 44))
        Exit Sub
         Case "LP"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(1, Rdata, 44)
            charlist(charindex).SinEnlistarAlianza = Val(ReadField(2, Rdata, 44))
        Exit Sub
        Case "PX"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(1, Rdata, 44)
            charlist(charindex).EsStatus = Val(ReadField(2, Rdata, 44))
            charlist(charindex).Nombre = ReadField(3, Rdata, 44)
        Exit Sub
        Case "DP"             ' >>>>> Borrar un NPC :: DP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(1, Rdata, 44)
            Call EraseChar(charindex)
            Call Dialogos.RemoveDialog(charindex)
            Call RefreshAllChars
        Exit Sub
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.RemoveDialog(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
    With charlist(charindex)
            
        For i = 1 To 3
            If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                .FxIndex(i) = 0
                .Fx(i).Loops = 0
            End If
        Next i
            
        ' Play steps sounds if the user is not an admin of any kind
        If .priv = 0 Then
            Call DoPasosFx(charindex)
        End If
    End With
            
            Call engine.Char_Move_by_Pos(charindex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
            
        Case "|H"    '>>>> Cambiar Heading Personaje :: |H
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            charlist(charindex).Heading = Val(ReadField(2, Rdata, 44))
        Exit Sub
        
        Case "|B"    '>>>> Cambiar Body Personaje :: |B
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            charlist(charindex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
        Exit Sub
        
        Case "|C"    '>>>> Cambiar Casco Personaje :: |C
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            charlist(charindex).Casco = CascoAnimData(Val(ReadField(2, Rdata, 44)))
        Exit Sub
        
        Case "|E"    '>>>> Cambiar Escudo Personaje :: |E
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            charlist(charindex).Escudo = ShieldAnimData(Val(ReadField(2, Rdata, 44)))
        Exit Sub
        
        Case "|W"    '>>>> Cambiar Arma Personaje :: |W
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            charlist(charindex).Arma = WeaponAnimData(Val(ReadField(2, Rdata, 44)))
        Exit Sub
        
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            engine.RemoveCharAparence Val(ReadField(1, Rdata, 44)), Val(ReadField(3, Rdata, 44)), Val(ReadField(2, Rdata, 44)), _
            Val(ReadField(3, Rdata, 44)), Val(ReadField(4, Rdata, 44)), Val(ReadField(5, Rdata, 44)), _
            Val(ReadField(6, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
            Val(ReadField(8, Rdata, 44))
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "N~"           ' >>>>> Nombre del Mapa
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Nombredelmapaxx = Rdata
        Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                        If Sound = True Then Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                    Else
                        If Sound = True Then Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            
            Exit Sub
        Case "XM"           ' >>>>> Play un MP3 :: XM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CurrentMP3 = Val(ReadField(1, Rdata, 45))
            
            
                If CurrentMP3 <> 0 Then
                    Audio.MP3_Play CurrentMP3
                End If
            
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW

                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")

            Exit Sub
        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case left$(sData, 3)
    
    Case "MAR"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        
        frmDuelos.Jugador1 = ReadField(1, Rdata, 44)
        frmDuelos.Jugador2 = ReadField(2, Rdata, 44)
        frmDuelos.Jugador3 = ReadField(3, Rdata, 44)
        frmDuelos.Jugador4 = ReadField(4, Rdata, 44)
        frmDuelos.Jugador5 = ReadField(5, Rdata, 44)
        frmDuelos.Jugador6 = ReadField(6, Rdata, 44)
        frmDuelos.Jugador7 = ReadField(7, Rdata, 44)
        frmDuelos.Jugador8 = ReadField(8, Rdata, 44)
        frmDuelos.Show
    Exit Sub
    
    Case "ICO" 'INICIO DE COMERCIO, SISTEMA NUEVO BY GHINZUL
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        comIniciar Rdata
    Exit Sub
    
    Case "IOR" 'RECIVO LA OFERTA (EN ORO)
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        rOro = Rdata
    Exit Sub
    
    Case "ICI" 'RECIVO LA OFERTA
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        comReciviOferta Rdata
    Exit Sub
    
    Case "VCC" 'CERRAR COMERCIO PUES
        comCerrar
    Exit Sub
    
    Case "MEC" 'MENSAJE EN CONSOLA
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        comMensaje ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
    Exit Sub
    
    '#####CORREOS####
    Case "IDO"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        correosIniciar Rdata
    Exit Sub
    
    Case "IFO"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        correosIniciarForm Rdata
    Exit Sub
    
    Case "IAO"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        correosListaAmigos Rdata
    Exit Sub
    
    Case "ILO"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        correosCargarMensaje Rdata
    Exit Sub
    '#####CORREOS####
    
    Case "NVG"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        charindex = Val(ReadField(1, Rdata, 44))
        charlist(charindex).Navegando = Val(ReadField(2, Rdata, 44))
    Exit Sub
        
    Case "MFC"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        frmCasas.Show , frmMain
    Exit Sub
    
    Case "TAL"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Dim mslink As String
        mslink = ReadField(1, Rdata, 44)
        
    If MsgBox("Los administradores del juego quieren que veas un link de una pagina web (" & mslink & "). ¿Deseas abrirla?", vbYesNo) = vbYes Then
        OpenBrowser "" & mslink & "", 0
    End If
        
    Exit Sub
 
    Case "GVN"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Dim ksitax As String
        Dim prsitox As Long
        Dim fchitax As String
        ksitax = ReadField(1, Rdata, 44)
        prsitox = ReadField(2, Rdata, 44)
        fchitax = ReadField(3, Rdata, 44)
    
        DueñoKsa = ksitax
        Preciox = prsitox
        Fechix = fchitax
        
        If DueñoKsa = "N/A" Then
         DueñoKsa = "DISPONIBLE"
         frmCasas.Command1.Enabled = True
        Else
         frmCasas.Command1.Enabled = False
        End If
       
        frmCasas.lblDueño.Caption = "DUEÑO: " & DueñoKsa
        frmCasas.lblPrecio.Caption = "PRECIO: " & PonerPuntos(prsitox)
        frmCasas.lblFecha.Caption = "FECHA: " & Fechix
    Exit Sub
        Case "USM"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).Montando = Val(ReadField(2, Rdata, 44))
        Exit Sub
        Case "LTR"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            Call frmTorneoManager.PonerListaTorneo(Rdata)
        Exit Sub
        Case "TSU"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            Call frmTorneoUsuarios.PonerListaTorneo(Rdata)
        Exit Sub
        Case "TSD"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            
            If Rdata = "0" Then
                frmTorneoUsuarios.Label1.Caption = "No se esta organizando ningun torneo actualmente. Podés organizar uno por 400.000 monedas de oro"
                frmTorneoUsuarios.Image4.Visible = True
            Else
                frmTorneoUsuarios.Label1.Caption = "" & Rdata & " esta organizando un Deathmatch para 16 participantes. El precio de inscripcion es de 200.000 monedas de oro y el nivel minimo es 20."
                frmTorneoUsuarios.Image4.Visible = False
            End If
            
        Exit Sub
        Case "8G1"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            frmNobleza.lstReq(0).Clear
            'frmGMAyudando.List1.Clear
            Dim noj As Integer
                For noj = 1 To Val(ReadField(1, Rdata, Asc(",")))
                    frmNobleza.lstReq(0).AddItem ReadField(2 * noj, ReadField(1, Rdata, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, Rdata, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "8G2"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            frmNobleza.lstReq(1).Clear
            'frmGMAyudando.List1.Clear
                For noj = 1 To Val(ReadField(1, Rdata, Asc(",")))
                    frmNobleza.lstReq(1).AddItem ReadField(2 * noj, ReadField(1, Rdata, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, Rdata, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "8G3"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            frmNobleza.lstReq(2).Clear
            'frmGMAyudando.List1.Clear
                For noj = 1 To Val(ReadField(1, Rdata, Asc(",")))
                    frmNobleza.lstReq(2).AddItem ReadField(2 * noj, ReadField(1, Rdata, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, Rdata, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "8G4"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            frmNobleza.lstReq(3).Clear
            'frmGMAyudando.List1.Clear
                For noj = 1 To Val(ReadField(1, Rdata, Asc(",")))
                    frmNobleza.lstReq(3).AddItem ReadField(2 * noj, ReadField(1, Rdata, Asc("%")), Asc(",")) & " (" & Val(ReadField((2 * noj) + 1, ReadField(1, Rdata, Asc("%")), Asc(","))) & ")"
                Next noj
            frmNobleza.Show , frmMain
        Exit Sub
        Case "LDM" 'carga lista de amigos
            Rdata = Right(Rdata, Len(Rdata) - 3)
            Call frmMain.PonerListaAmigos(Rdata)
        Exit Sub
        Case "KFM" 'conecta amigo
        Rdata = Right(Rdata, Len(Rdata) - 3)
            If Configuracion.AnunciarContacto = 1 Then AddtoRichTextBox frmMain.RecTxt, "" & UCase$(Rdata) & " se ha conectado.", 0, 255, 0, True, False, False
        Exit Sub
        
        Case "DFM" 'desconecta amigo
        Rdata = Right(Rdata, Len(Rdata) - 3)
            If Configuracion.AnunciarContacto = 1 Then AddtoRichTextBox frmMain.RecTxt, "" & UCase$(Rdata) & " se ha desconectado.", 255, 0, 0, True, False, False
        Exit Sub
        
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.RemoveDialog(Val(Rdata))
            Exit Sub
        Case "CFF"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField$(1, Rdata, 44))
            charlist(charindex).particle_count = Val(ReadField$(2, Rdata, 44))
           
            Call General_Char_Particle_Create(charlist(charindex).particle_count, charindex)
            Call RefreshAllChars
        Exit Sub
        Case "PCF"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Dim Particulita As Byte
            Dim Tiempito As Byte
            Dim equiss As Byte
            Dim equiyy As Byte
            Particulita = Val(ReadField$(1, Rdata, 44))
            equiss = Val(ReadField$(2, Rdata, 44))
            equiyy = Val(ReadField$(3, Rdata, 44))
            Tiempito = Val(ReadField$(4, Rdata, 44))
           
            Call General_Particle_Create(Particulita, equiss, equiyy, Tiempito)
        Exit Sub
      Case "CTC"                  ' >>>> Crear particula en char.
            Dim char_index      As Integer
            Dim other_CharIndex As Integer
            Dim particle_Index  As Integer
            Dim particle_Speed  As Single
           
            'Corto la data.
            Rdata = Right$(Rdata, Len(Rdata) - 3)
           
            'Busco el char.
            char_index = Val(ReadField(1, Rdata, 44))
            other_CharIndex = Val(ReadField(2, Rdata, 44))
           
            'Datos de la partícula.
            particle_Index = Val(ReadField(3, Rdata, 44))
            particle_Speed = CSng(ReadField(4, Rdata, 44))
           
            'Si los chars son válidos y la partícula también.
            If (char_index <> 0) And (other_CharIndex <> 0) Then
               If (particle_Index <> 0) And (particle_Index <= UBound(StreamData())) Then
                  Call engine.Create_Particle(char_index, other_CharIndex, particle_Index, particle_Speed)
               End If
            End If
       Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField(1, Rdata, 44))
            Call SetCharacterFx(charindex, Val(ReadField(2, Rdata, 44)), Val(ReadField(3, Rdata, 44)))
        Exit Sub
       Case "CFE"                  ' >>>>> Mostrar Emoticones :: CFE
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).EmoticonIndex = Val(ReadField(2, Rdata, 44))
            charlist(charindex).EmoticonLoops = Val(ReadField(3, Rdata, 44))
            Call SetCharacterEmoticon(charindex, charlist(charindex).EmoticonIndex, charlist(charindex).EmoticonLoops)
        Exit Sub
        Case "ANM"
       
        Rdata = Right$(Rdata, Len(Rdata) - 3)
            ArmaMin = Val(ReadField(1, Rdata, 44))
            ArmaMax = Val(ReadField(2, Rdata, 44))
            ArmorMin = Val(ReadField(3, Rdata, 44))
            ArmorMax = Val(ReadField(4, Rdata, 44))
            EscuMin = Val(ReadField(5, Rdata, 44))
            EscuMax = Val(ReadField(6, Rdata, 44))
            CascMin = Val(ReadField(7, Rdata, 44))
            CascMax = Val(ReadField(8, Rdata, 44))
            HerrMin = Val(ReadField(9, Rdata, 44))
            HerrMax = Val(ReadField(10, Rdata, 44))
            MagMin = Val(ReadField(11, Rdata, 44))
            MagMax = Val(ReadField(12, Rdata, 44))
            MagMina = Val(ReadField(13, Rdata, 44))
            MagMaxa = Val(ReadField(14, Rdata, 44))
            MagMinb = Val(ReadField(15, Rdata, 44))
            MagMaxb = Val(ReadField(16, Rdata, 44))
            MagMinc = Val(ReadField(17, Rdata, 44))
            MagMaxc = Val(ReadField(18, Rdata, 44))
            MagMind = Val(ReadField(19, Rdata, 44))
            MagMaxd = Val(ReadField(20, Rdata, 44))
 
        With frmMain
                .Arma.Caption = ArmaMin & "/" & ArmaMax
                .Defensa.Caption = ArmorMin + EscuMin + CascMin + HerrMin & "/" & ArmorMax + EscuMax + CascMax + HerrMax
                .DefMag.Caption = MagMin + MagMina + MagMinb + MagMinc + MagMind & "/" & MagMax + MagMaxa + MagMaxb + MagMaxc + MagMaxd
        End With
        
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg N, n2
            frmMSG.Show , frmMain
        Exit Sub
            
        Case "EZT"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserStatus = Val(ReadField(1, Rdata, 44))
        Exit Sub
        
        Case "LDG"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserPrivilegios = Val(Rdata)
            If UserPrivilegios = 0 Then
                frmMain.GMSOS.Visible = False
                frmMain.GMTORNEO.Visible = False
                frmMain.GMPANEL.Visible = False
                frmMain.Command1.Visible = False
            Else
                frmMain.GMSOS.Visible = True
                frmMain.GMTORNEO.Visible = True
                frmMain.GMPANEL.Visible = True
                frmMain.Command1.Visible = True
            End If
        Exit Sub
        
        Case "IVX"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
                       
            InventorySlots = Rdata
        Exit Sub
        
        Case "CHX"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHPCHORI = Val(ReadField(1, Rdata, 44))
            UserMinHPCHORI = Val(ReadField(2, Rdata, 44))
            UserMaxMANCHORI = Val(ReadField(3, Rdata, 44))
            UserMinMANCHORI = Val(ReadField(4, Rdata, 44))
            NickCHORI = ReadField(5, Rdata, 44)
           
            If Form1.Visible = False Then
                Form1.Show
            End If
            
            Form1.Shape1.Width = (((UserMinHPCHORI / 100) / (UserMaxHPCHORI / 100)) * 1695)
            Form1.Shape2.Width = (((UserMinMANCHORI / 100) / (UserMaxMANCHORI / 100)) * 1695)
            Form1.Label1.Caption = UserMinHPCHORI & "/" & UserMaxHPCHORI
            Form1.Label2.Caption = UserMinMANCHORI & "/" & UserMaxMANCHORI
            Form1.Caption = NickCHORI
           
        Exit Sub
        Case "VOT"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Dim Vot(1 To 5) As String, Votacion As String
        Vot(1) = ReadField(1, Rdata, 44)
        Vot(2) = ReadField(2, Rdata, 44)
        Vot(3) = ReadField(3, Rdata, 44)
        Vot(4) = ReadField(4, Rdata, 44)
        Vot(5) = ReadField(5, Rdata, 44)
        Votacion = ReadField(6, Rdata, 44)
        
            With frmVotacionUser
            
                For i = 1 To 5
                    If Vot(i) = "" Then
                        .Votos(i - 1).Enabled = False
                        .Votos(i - 1).Caption = "N/A"
                    Else
                        .Votos(i - 1).Enabled = True
                        .Votos(i - 1).Caption = Vot(i)
                    End If
                Next i
                
                .Label1.Caption = Votacion
                .Show , frmMain
            
            End With
        Exit Sub
        
        Case "WEN"
            FrmNewPoll.Show , frmMain
        Exit Sub
        
     
         Case "BYE"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
            With frmVotacionUser
                .Label1.Caption = ""
                
                For i = 1 To 5
                    .Votos(i - 1).Enabled = False
                    .Votos(i - 1).Caption = "N/A"
                Next i
            End With
        Exit Sub
        
        Case "IFE"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            frmEstadisticasUsuario.lblNombre.Caption = ReadField(1, Rdata, 44)
            frmEstadisticasUsuario.lblClase.Caption = ReadField(2, Rdata, 44)
            frmEstadisticasUsuario.lblRaza.Caption = ReadField(3, Rdata, 44)
            frmEstadisticasUsuario.lblNivel = ReadField(4, Rdata, 44)
            frmEstadisticasUsuario.lblExp.Caption = ReadField(5, Rdata, 44)
            frmEstadisticasUsuario.lblFaccion = ReadField(6, Rdata, 44)
            frmEstadisticasUsuario.lblJerarquia.Caption = ReadField(7, Rdata, 44)
            frmEstadisticasUsuario.lblReputacion.Caption = ReadField(8, Rdata, 44)
            frmEstadisticasUsuario.lblDuelos.Caption = ReadField(9, Rdata, 44)
            frmEstadisticasUsuario.lblParejas.Caption = ReadField(10, Rdata, 44)
            frmEstadisticasUsuario.lblRondas.Caption = ReadField(11, Rdata, 44)
            frmEstadisticasUsuario.lblMuertes.Caption = ReadField(12, Rdata, 44)
            frmEstadisticasUsuario.lblUsuariosMatados.Caption = ReadField(13, Rdata, 44)
            frmEstadisticasUsuario.lblEventos.Caption = ReadField(14, Rdata, 44)
            frmEstadisticasUsuario.lblCVCS.Caption = ReadField(15, Rdata, 44)
            frmEstadisticasUsuario.lblQuests.Caption = ReadField(16, Rdata, 44)

            frmEstadisticasUsuario.Show , frmMain

        Exit Sub
        
        Case "AU|"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Dim Armadura As Integer
            Dim Weapon As Integer
            Dim EscudoA As Integer
            Dim Ring As Integer
            Dim CascoA As Integer
            
            charindex = ReadField(1, Rdata, 44)
            Armadura = Val(ReadField(2, Rdata, 44))
            Weapon = Val(ReadField(3, Rdata, 44))
            EscudoA = Val(ReadField(4, Rdata, 44))
            Ring = Val(ReadField(5, Rdata, 44))
            CascoA = Val(ReadField(6, Rdata, 44))
            
            
            If Armadura > 0 Then
                charlist(charindex).Aura_IndexA = Armadura
                
                If AurasPJ(charlist(charindex).Aura_IndexA).RojoF > 0 Then
                    charlist(charindex).AuraAntiguoR = AurasPJ(charlist(charindex).Aura_IndexA).RojoF
                    charlist(charindex).AuraAntiguoG = AurasPJ(charlist(charindex).Aura_IndexA).VerdeF
                    charlist(charindex).AuraAntiguoB = AurasPJ(charlist(charindex).Aura_IndexA).AzulF
                        
                    charlist(charindex).AuraQueremosLlegarR = AurasPJ(charlist(charindex).Aura_IndexA).r
                    charlist(charindex).AuraQueremosLlegarG = AurasPJ(charlist(charindex).Aura_IndexA).g
                    charlist(charindex).AuraQueremosLlegarB = AurasPJ(charlist(charindex).Aura_IndexA).b
                    
                    charlist(charindex).AuraProximoR = AurasPJ(charlist(charindex).Aura_IndexA).RojoF
                    charlist(charindex).AuraProximoG = AurasPJ(charlist(charindex).Aura_IndexA).VerdeF
                    charlist(charindex).AuraProximoB = AurasPJ(charlist(charindex).Aura_IndexA).AzulF
                    charlist(charindex).AuraLlegoAlColor = True
                    
                End If
                
                Call InitGrh(charlist(charindex).AuraA, AurasPJ(charlist(charindex).Aura_IndexA).GrhIndex)
                charlist(charindex).Aura_AngleA = 0
            Else
                charlist(charindex).Aura_IndexA = 0
            End If
            
            If Ring > 0 Then
                charlist(charindex).Aura_IndexR = Ring
                Call InitGrh(charlist(charindex).AuraR, AurasPJ(charlist(charindex).Aura_IndexR).GrhIndex)
                charlist(charindex).Aura_AngleR = 0
            Else
                charlist(charindex).Aura_IndexR = 0
            End If
            
            
            If Weapon > 0 Then
                charlist(charindex).Aura_IndexW = Weapon
                Call InitGrh(charlist(charindex).AuraW, AurasPJ(charlist(charindex).Aura_IndexW).GrhIndex)
                charlist(charindex).Aura_AngleW = 0
            Else
                charlist(charindex).Aura_IndexW = 0
            End If
                
                
            If CascoA > 0 Then
                charlist(charindex).Aura_IndexC = CascoA
                Call InitGrh(charlist(charindex).AuraC, AurasPJ(charlist(charindex).Aura_IndexC).GrhIndex)
                charlist(charindex).Aura_AngleC = 0
            Else
                charlist(charindex).Aura_IndexC = 0
            End If
                
            If EscudoA > 0 Then
                charlist(charindex).Aura_IndexE = EscudoA
                Call InitGrh(charlist(charindex).AuraE, AurasPJ(charlist(charindex).Aura_IndexE).GrhIndex)
                charlist(charindex).Aura_AngleE = 0
            Else
                charlist(charindex).Aura_IndexE = 0
            End If
        Exit Sub
        
        Case "[Q]"
            With TextDesv
                .color = D3DColorXRGB(255, 255, 0)
                .StartTime = GetTickCount()
                .Text = "¡QUEST COMPLETADA, FELICIDADES!"
                .LifeTime = 800
                .Sube = 18
                .Desvanecimiento = 20
                .Tiempito = False
                .Existe = True
                .X = 285
                .Y = 170
                
                Call General_Char_Particle_Create(58, UserCharIndex)
            End With
        Exit Sub
        
        Case "[H]" 'Actualiza vida
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserMaxHP = Val(ReadField(1, Rdata, 44))
             UserMinHP = Val(ReadField(2, Rdata, 44))
             
            'Seteamos el shape y el label de vida
            frmMain.HPShp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 95)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
                
            'Seteamos que el usuario murió
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        Exit Sub
        
        Case "[M]" 'Actualiza mana
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserMaxMAN = Val(ReadField(1, Rdata, 44))
             UserMinMAN = Val(ReadField(2, Rdata, 44))
             
             'Seteamos el shape y el label de vida
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 95)
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            Else
                frmMain.MANShp.Width = 0
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            End If
        Exit Sub
        
        Case "[S]" 'Actualiza stamina
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserMaxSTA = Val(ReadField(1, Rdata, 44))
             UserMinSTA = Val(ReadField(2, Rdata, 44))
             
             'Seteamos el shape y label de energia
             frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 95)
             frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        Exit Sub
        
        Case "[G]" 'Actualiza oro
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserGLD = Val(ReadField(1, Rdata, 44))
            
             frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
        Exit Sub
        
        Case "|S1" 'Actualiza slot
             Rdata = Right$(Rdata, Len(Rdata) - 3)
            
             Call Inventario.ActualizarSlotCant(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44))
        Exit Sub
        
        Case "|S2" 'Actualiza slot
             Rdata = Right$(Rdata, Len(Rdata) - 3)
            
             Call Inventario.ActualizarSlotEquipped(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44))
        Exit Sub
        
        Case "[L]" 'Actualiza nivel
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserLvl = Val(ReadField(1, Rdata, 44))
            
            'Seteamos el label de nivel
            If UserLvl >= 50 Then
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbYellow
            Else
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbRed
            End If
        Exit Sub
        
        Case "[E]" 'Actualizar experiencia
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserPasarNivel = Val(ReadField(1, Rdata, 44))
             UserExp = Val(ReadField(2, Rdata, 44))
            
            'Seteamos ancho de barra, label y experiencia, todo junto.
            If UserPasarNivel > 0 Then
                frmMain.ExpBar.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 195)
                frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
                frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            Else
                frmMain.ExpBar.Width = 195
                frmMain.exp.Caption = "0/0"
                frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"
            End If
             
        Exit Sub
        
        Case "[B]" 'Actualizar oro de la boveda
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             UserBOVItem = Val(ReadField(1, Rdata, 44))
             
             'Seteo label de boveda
             frmBanco.Text1.Caption = PonerPuntos(UserBOVItem)
        Exit Sub
        
        Case "[N]" 'Actualiza el nombre
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             NickPJ = ReadField(1, Rdata, 44)
             
             'label de nombre
             frmMain.Label8.Caption = NickPJ
             UserName = NickPJ
        Exit Sub
        
        Case "[F]" 'Actualiza fuerza
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             frmMain.Fuerza.Caption = ReadField(1, Rdata, 44)
        Exit Sub
        
        Case "[A]" 'Actualiza agilidad
             Rdata = Right$(Rdata, Len(Rdata) - 3)
             frmMain.Agilidad.Caption = ReadField(1, Rdata, 44)
        Exit Sub
        
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            Call Inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44))
        Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
        Exit Sub
        
        Case "SBG"                 ' >>>>> Actualiza Cuenta Bancaria
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventoryB(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventoryB(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventoryB(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventoryB(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventoryB(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventoryB(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventoryB(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventoryB(slot).Def = Val(ReadField(9, Rdata, 44))
            UserBancoOro = Val(ReadField(10, Rdata, 44))
            UserBancoOroPropio = Val(ReadField(11, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventoryB(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventoryB(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventoryB(slot).Name
            End If
            
        Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "INK"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            frmEstadisticas.InformarQuests (Rdata)
        Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show , frmMain
        Exit Sub
        
        Case "DRM"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            
            frmMercadoTS.lstPacks.Clear
            frmMercadoTS.lblTSPoints.Caption = PonerPuntos(ReadField(1, Rdata, 44))
            For i = 1 To ReadField(2, Rdata, 44)
                frmMercadoTS.lstPacks.AddItem ReadField(2 + i, Rdata, 44)
            Next i
            
            frmMercadoTS.Show , frmMain
        Exit Sub
        
        Case "DNF"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            Dim ContentDonation As String
            Dim TempDonation As String
            
            picDonation.Body = ReadField(1, Rdata, 44)
            picDonation.Head = ReadField(2, Rdata, 44)
            picDonation.Weapon = ReadField(3, Rdata, 44)
            picDonation.Shield = ReadField(4, Rdata, 44)
            picDonation.Casco = ReadField(5, Rdata, 44)
            picDonation.Aura = ReadField(6, Rdata, 44)
            picDonation.GrhIndex = ReadField(7, Rdata, 44)
            
            If picDonation.Aura > 0 Then
                Call InitGrh(picDonation.AuraA, AurasPJ(picDonation.Aura).GrhIndex)
                picDonation.Aura_Angle = 0
            End If
            
            With frmMercadoTS
                    ContentDonation = ""
                    For i = 1 To ReadField(9, Rdata, 44)
                        TempDonation = ReadField(9 + i, Rdata, 44)
                        ContentDonation = ContentDonation & "• " & ReadField(1, TempDonation, Asc("-")) & " - " & ReadField(2, TempDonation, Asc("-")) & vbCrLf & ""
                    Next i
                    
                    .lblContent.Caption = ContentDonation
                    .lblPrice = PonerPuntos(ReadField(8, Rdata, 44))
            End With
        Exit Sub
        
        Case "PRM"
                Rdata = Right(Rdata, Len(Rdata) - 3)
               
                frmCanjes.ListaPremios.Clear
                For i = 1 To Val(ReadField(1, Rdata, 44))
                    frmCanjes.ListaPremios.AddItem ReadField(i + 1, Rdata, 44)
                Next i
               
                frmCanjes.Show , frmMain
        Exit Sub
               
            Case "INF"
                Rdata = Right(Rdata, Len(Rdata) - 3)
            With frmCanjes
                    .Requiere.Caption = ReadField(1, Rdata, 44)
                    .lAtaque.Caption = ReadField(3, Rdata, 44) & "/" & ReadField(2, Rdata, 44)
                    .lDef.Caption = ReadField(5, Rdata, 44) & "/" & ReadField(4, Rdata, 44)
                    .lAM.Caption = ReadField(7, Rdata, 44) & "/" & ReadField(6, Rdata, 44)
                    .lDM.Caption = ReadField(9, Rdata, 44) & "/" & ReadField(8, Rdata, 44)
                    .lDescripcion.Text = ReadField(10, Rdata, 44)
                    .lPuntos.Caption = ReadField(11, Rdata, 44)
                    
            CantidadCanjeYegua = ReadField(1, Rdata, 44)
            
                        If .Requiere.Caption = "0" Then
            .Requiere.Caption = "N/A"
            End If
                        If .lAtaque.Caption = "0/0" Then
            .lAtaque.Caption = "N/A"
            End If
                        If .lDef.Caption = "0/0" Then
            .lDef.Caption = "N/A"
            End If
                        If .lAM.Caption = "0/0" Then
            .lAM.Caption = "N/A"
            End If
                        If .lDM.Caption = "0/0" Then
            .lDM.Caption = "N/A"
            End If

            Dim Grhpremios As Integer
            Grhpremios = ReadField(12, Rdata, 44)
                Dim SR As RECT, DR As RECT
                
                SR.left = 0
                SR.top = 0
                SR.Right = 32
                SR.bottom = 32
                
                DR.left = 0
                DR.top = 0
                DR.Right = 32
                DR.bottom = 32
                
                Call engine.DrawGrhtoHdc(Grhpremios, SR, frmCanjes.Picture1)
            End With
        Exit Sub
            
        Case "ERO"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
        
            Mensaje.Label1 = Rdata
            Mensaje.Show
        
        Exit Sub
        
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmConnect.MousePointer = 1
            'frmPasswdSinPadrinos.MousePointer = 1
            
            'If Not frmCrearPersonaje.Visible Then
            '    frmMain.Socket1.Disconnect
            'End If
            
            Mensaje.Escribir Rdata
        Exit Sub
    End Select
    
    
    Select Case left$(sData, 4)
    
        Case "|TSF"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        
            frmNewTiendaTS.lblTSPoints = Val(ReadField(1, Rdata, Asc(",")))
            frmNewTiendaTS.Show , frmMain
        Exit Sub
    
        Case "MISI"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call frmMisionesDiarias.ParseQuest(Rdata)
        Exit Sub
        
        Case "MISJ"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            If Configuracion.MisionDiaria = 1 Then Call frmMisionesDiarias.ParseQuest(Rdata)
        Exit Sub
        
        Case "MTOP"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        
            Call frmRanking.MostrarRanking(Rdata)
        Exit Sub
    
        Case "RANK"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        charindex = ReadField(1, Rdata, 44)
        
        charlist(charindex).TieneRanking = ReadField(2, Rdata, 44)
        charlist(charindex).PosRanking = ReadField(3, Rdata, 44)
        
        Exit Sub
        
        Case "ZSOS"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        MensajesNumber = ReadField(1, Rdata, Asc(","))
        
        Dim SOSTemporal As String
        frmGmPanelSOS.UserSOSList.Clear
        SOSTemporal = ""
        
            For i = 1 To MensajesNumber
                SOSTemporal = ReadField(1 + i, Rdata, Asc(","))
                MensajesSOS(i).Tipo = ReadField(1, SOSTemporal, Asc("-"))
                MensajesSOS(i).Autor = ReadField(2, SOSTemporal, Asc("-"))
                MensajesSOS(i).Contenido = ReadField(3, SOSTemporal, Asc("-"))
                frmGmPanelSOS.UserSOSList.AddItem "[" & MensajesSOS(i).Tipo & "] - " & MensajesSOS(i).Autor
                frmGmPanelSOS.UserSOSList.Refresh
            Next i
        
        Exit Sub
    
        Case "ARIE"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = ReadField(1, Rdata, 44)
            charlist(charindex).Ariete = True
        Exit Sub
    
        Case "MJOR"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            
            FrmMejorar.ListaMejorados.AddItem Rdata
            
            
            If UCase$(Rdata) = "SIN ITEMS MEJORABLES" Then
                FrmMejorar.ListaMejorados.Enabled = False
            Else
                FrmMejorar.ListaMejorados.Enabled = True
            End If
            
            FrmMejorar.Show , frmMain
        Exit Sub
            
        Case "IMEJ"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            
            With FrmMejorar
            
            .Nombre.Caption = ReadField(1, Rdata, 44)
            .Ataque.Caption = ReadField(2, Rdata, 44)
            .Defensa.Caption = ReadField(3, Rdata, 44)
            .AtaqueMagico.Caption = ReadField(4, Rdata, 44)
            .DefensaMagica.Caption = ReadField(5, Rdata, 44)
            .Desc.Text = ReadField(6, Rdata, 44)
            
            SR.bottom = 32
            SR.Right = 32
            SR.left = 0
            SR.top = 0
            
                Dim GrhMejorar As Integer
                    GrhMejorar = ReadField(7, Rdata, 44)
                    Call engine.DrawGrhtoHdc(GrhMejorar, SR, .Item)
            End With
            
        Exit Sub
        Case "GODS"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        Dim AlmasOfrecidas As Long
        Dim AlmasNecesarias As Long
        Dim SirvienteDe As String
        AlmasOfrecidas = Val(ReadField(1, Rdata, 44))
        AlmasNecesarias = Val(ReadField(2, Rdata, 44))
        SirvienteDe = ReadField(3, Rdata, 44)
        
        frmGods.lblOfrecidos = "" & AlmasOfrecidas & "/" & AlmasNecesarias & ""
        frmGods.imgAlmas.Width = (((AlmasOfrecidas / 100) / (AlmasNecesarias / 100)) * 335)
        
        frmGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Main.jpg")
        
        If UCase$(SirvienteDe) = "MIFRIT" Then
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios2.jpg")
        ElseIf UCase$(SirvienteDe) = "TERRASKE" Then
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios4.jpg")
        ElseIf UCase$(SirvienteDe) = "EREBROS" Then
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios1.jpg")
        ElseIf UCase$(SirvienteDe) = "POSEIDON" Then
         frmGods.imgGods.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Almas_Dios3.jpg")
        End If
               
        frmGods.Show , frmMain
         
        Exit Sub
        
        Case "PCCC" ' CHOTS | Poner Captions en frm
            Dim Caption As String
            Dim Nomvre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Caption = ReadField(1, Rdata, 44)
            Nomvre = ReadField(2, Rdata, 44)
            Call frmProcesos.Show
            frmProcesos.Procesos.Visible = True
            frmProcesos.Captions.Visible = False
            frmProcesos.Command1.Enabled = False
            frmProcesos.Command2.Enabled = True
            frmProcesos.Captions.AddItem Caption
        Case "PCCP" ' CHOTS | Listar Captions
            frmProcesos.Captions.Clear
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call frmProcesos.Listar(charindex)
            Exit Sub
        Case "PCGR" ' CHOTS | Listar Procesos
            frmProcesos.Procesos.Clear
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call enumProc(charindex)
        Exit Sub
        Case "PCSC" ' CHOTS | Listar Prosesos
            frmProcesos.Procesos.Clear
            frmProcesos.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call PROC(charindex)
        Exit Sub
        Case "PCGN" ' CHOTS | Poner Procesos en frm
            Dim Proceso As String
            Dim Nombre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Proceso = ReadField(1, Rdata, 44)
            Nombre = ReadField(2, Rdata, 44)
            frmProcesos.Procesos.AddItem Proceso
            frmProcesos.Caption = Nombre
            frmProcesos.txtUrl.Text = Nombre
        Case "PCSS" ' CHOTS | Poner Prosesos en frm
            Dim Proseso As String
            Dim Nonbre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Proseso = ReadField(1, Rdata, 44)
            Nonbre = ReadField(2, Rdata, 44)
            frmProcesos.Procesos.AddItem Proseso
    Case "MENU"
        If Configuracion.MenuDesplegable = 0 Then Exit Sub
                Dim esgm As Byte
                Rdata = Right$(Rdata, Len(Rdata) - 4)
                nombreotro = ReadField(1, Rdata, 44)
                esgm = ReadField(2, Rdata, 44)
                    If esgm > 0 Then
                        frmMenuGM.Show , frmMain
                    Else
                        frmMenu.Show , frmMain
                    End If
                Exit Sub
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "DTLC"
            Rdata = Right(Rdata, Len(Rdata) - 4)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
        Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 40)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 40)
            frmMain.AGUABAR.Caption = UserMinAGU & "%"
            frmMain.COMIDABAR.Caption = UserMinHAM & "%"
            Exit Sub
        Case "KIGF" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .Clase = ReadField(1, Rdata, 44)
                .Email = ReadField(2, Rdata, 44)
                .Advertencias = Val(ReadField(3, Rdata, 44))
                .DuelosGanados = Val(ReadField(4, Rdata, 44))
                .DuelosPerdidos = Val(ReadField(5, Rdata, 44))
                .CopasDeOro = Val(ReadField(6, Rdata, 44))
                .CopasDePlata = Val(ReadField(7, Rdata, 44))
                .CopasDeBronce = Val(ReadField(8, Rdata, 44))
                .QuestCompletadas = Val(ReadField(9, Rdata, 44))
                .CiudadanosMatados = Val(ReadField(10, Rdata, 44))
                .CriminalesMatados = Val(ReadField(11, Rdata, 44))
                .NPCSMATADOS = Val(ReadField(12, Rdata, 44))
                .Jerarquia = ReadField(13, Rdata, 44)
                .Restantes = ReadField(14, Rdata, 44)
                .Alineacion = ReadField(15, Rdata, 44)
                .GuerrasGanadas = ReadField(16, Rdata, 44)
                .CvcsGanados = ReadField(17, Rdata, 44)
                .MVPMatados = ReadField(18, Rdata, 44)
                .PuntosTorneo = PonerPuntos(ReadField(19, Rdata, 44))
                .Hogar = ReadField(20, Rdata, 44)
                .Genero = ReadField(21, Rdata, 44)
                .Nivel = ReadField(22, Rdata, 44)
                .Bonif1 = ReadField(23, Rdata, 44)
                .Bonif2 = ReadField(24, Rdata, 44)
                .Bonif3 = ReadField(25, Rdata, 44)
                .Nombre = ReadField(26, Rdata, 44)
                .TipoQuest = ReadField(27, Rdata, 44)
                .DescQuest = ReadField(28, Rdata, 44)
                .PremioOro = ReadField(29, Rdata, 44)
                .PremioPuntis = ReadField(30, Rdata, 44)
                .CantidadNPCs = ReadField(31, Rdata, 44)
                .YaMatados = ReadField(32, Rdata, 44)
                .TorneosParticipados = ReadField(33, Rdata, 44)
                .MaximasRondas = ReadField(34, Rdata, 44)
                .Eventos = ReadField(35, Rdata, 44)
                .ParejasGanadas = ReadField(36, Rdata, 44)
                .ParejasPerdidas = ReadField(37, Rdata, 44)
                .GuerrasPerdidas = ReadField(38, Rdata, 44)
                .NeutralesMatados = ReadField(39, Rdata, 44)
                .MuertesUsuario = ReadField(40, Rdata, 44)
                .Raza = ReadField(41, Rdata, 44)
                .UserReputacion = ReadField(42, Rdata, 44)
                .PuntosDonador = ReadField(43, Rdata, 44)
            End With
                frmEstadisticas.Iniciar_Labels
                frmEstadisticas.Show , frmMain
            Exit Sub
            
        Case "SUNX"
            frmNoesNW.Show , frmMain
        Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case left$(sData, 5)
        Case UCase$(Chr$(110)) & mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)

            Exit Sub
        Case "KHEKD"
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        
            RetiraObj = ReadField(1, Rdata, Asc(","))
            RetiraOro = ReadField(2, Rdata, Asc(","))
        
        Exit Sub
        Case "ZMOTD"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.Text = Rdata
            Exit Sub
       Case "INIAC"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            
            Call CambiarConectar("CUENTA")
            
        'frmCuent.SetFocus
        Exit Sub
        Case "STOPD"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            Stopped = ReadField(1, Rdata, 44)
        Exit Sub
        Case "CODEH"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CodigoRecibido = ReadField(1, Rdata, 44)
        Exit Sub
        Case "DLPSJ"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CantidadDePersonajes = ReadField(1, Rdata, 44)
        Exit Sub
        Case "ADDPJ"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
           
            rcvName = ReadField(1, Rdata, 44)
            rcvIndex = ReadField(2, Rdata, 44)
            rcvHead = ReadField(3, Rdata, 44)
            rcvBody = ReadField(4, Rdata, 44)
            rcvWeapon = ReadField(5, Rdata, 44)
            rcvShield = ReadField(6, Rdata, 44)
            rcvCasco = ReadField(7, Rdata, 44)
            rcvLevel = ReadField(8, Rdata, 44)
            rcvClase = ReadField(9, Rdata, 44)
            rcvMuerto = ReadField(10, Rdata, 44)
            rcvRaza = ReadField(11, Rdata, 44)
                      
            CargarPJ(rcvIndex - 1).Nombre = rcvName
            CargarPJ(rcvIndex - 1).Body = rcvBody
            CargarPJ(rcvIndex - 1).Head = rcvHead
            CargarPJ(rcvIndex - 1).Casco = rcvCasco
            CargarPJ(rcvIndex - 1).Shield = rcvShield
            CargarPJ(rcvIndex - 1).Weapon = rcvWeapon
            CargarPJ(rcvIndex - 1).Level = rcvLevel
            CargarPJ(rcvIndex - 1).Existe = True
            CargarPJ(rcvIndex - 1).Raza = rcvRaza
            CargarPJ(rcvIndex - 1).Muerto = rcvMuerto
        Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select

    Select Case left(sData, 6)
    
        Case "FLECHI" 'flecha a char
         Rdata = Right$(Rdata, Len(Rdata) - 6)
            engine.Crear_Flecha Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44)), Val(ReadField(3, Rdata, 44)), 0, Val(ReadField(4, Rdata, 44))
        Exit Sub
    
    'Acá abre la ventana al primer usuario (el que va a enviar el mensaje)
      Case "ENCHAT"
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        
        For i = 1 To 5
            If UCase$(NickContacto(i)) = UCase$(Rdata) Then
                Mensaje.Escribir "Ya tienes una ventana de chat abierta con este usuario."
             Exit Sub
            End If
        
            If ChatEnUso(i) = False Then
                NickContacto(i) = UCase$(Rdata)
                ChatEnUso(i) = True
                VentanitaMostrar(i) = 2
                
                ChatForm(i).Caption = Rdata
                ChatForm(i).lblName = Rdata
                ChatForm(i).Show , frmMain
                Exit Sub
            End If
        Next i
        
      Exit Sub
      
      'Acá la ventana al segundo usuario (el que recibe)
      Case "LDCHAT"
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        Dim Remitente As String, Mensajitox As String
        Remitente = ReadField(1, Rdata, 44)
        Mensajitox = ReadField(2, Rdata, 44)
        
        For i = 1 To 5
            If UCase$(NickContacto(i)) = UCase$(Remitente) Then
                    AddtoRichTextBox ChatForm(i).rtbChat, "" & Remitente & " dice: " & Mensajitox & "", 255, 0, 0, True
                    RecibioMensaje(i) = True
                Exit Sub
            End If
        
            If ChatEnUso(i) = False Then
                NickContacto(i) = UCase$(Remitente)
                ChatEnUso(i) = True
                RecibioMensaje(i) = True
                
                ChatForm(i).Caption = Remitente
                ChatForm(i).lblName = Remitente
                AddtoRichTextBox ChatForm(i).rtbChat, "" & Remitente & " dice: " & Mensajitox & "", 255, 0, 0, True
                Exit Sub
            End If
        Next i
        
      Exit Sub
    
      Case "CIRUJA"
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        Dim Raza As String, Genero As String
        Raza = ReadField(1, Rdata, 44)
        Genero = ReadField(2, Rdata, 44)
        FrmCirujia.Show , frmMain
        Call FrmCirujia.ParseHead(Raza, Genero)
    Exit Sub
        Case "AXELPT"
            frmMenuMascota.Show , frmMain
        Exit Sub
      Case "GENPAS" 'GENERAR PASSWORD PARA RECUPERAR CUENTA [Dylan.-]
        Rdata = Right$(Rdata, Len(Rdata) - 6)
            Dim PassGenerada As String
            PassGenerada = Rdata
            'frmMensaje.Show
            MsgBox "Su nueva contraseña es: " & PassGenerada & ". Asegúrate de cambiar la contraseña antes de entrar en un personaje, de lo contrario no podrás acceder a tus personajes."
            Unload frmRecuperar
        Exit Sub
        Case "PEDPRE" 'ENVIO DE PREGUNTA SECRETA [Dylan.-]
        Rdata = Right$(Rdata, Len(Rdata) - 6)
            If frmCambiarPass.Visible = True Then
            frmCambiarPass.pregunta.Caption = Rdata
            Exit Sub
            End If
            If frmRecuperar.Visible = True Then
            frmRecuperar.Height = 4980
            frmRecuperar.txtMail.Locked = True
            frmRecuperar.txtNombre.Locked = True
            frmRecuperar.txtPregunta.Visible = True
            frmRecuperar.txtRespuesta.Visible = True
            frmRecuperar.txtRespuesta.SetFocus
            frmRecuperar.Recuperar.Visible = True
            frmRecuperar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Recuperar2Fin.jpg")
            frmRecuperar.Siguiente.Visible = False
            frmRecuperar.Cancelar.Visible = False
            frmRecuperar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Recuperar2.jpg")
            frmRecuperar.txtPregunta.Caption = Rdata
        Exit Sub
        End If
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case left$(sData, 7)
          Case "RESPUES"         ' >>> Sistema Consultas - Fishar.-
            Rdata = Right(Rdata, Len(Rdata) - 7)
            TieneParaResponder = True
            frmMensaje.msg.Text = ReadField(1, Rdata, Asc("*")) & vbCrLf & "Respondido por: " & ReadField(2, Rdata, Asc("*"))
        Case "NEWDENU"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            DenunciasNumber = DenunciasNumber + 1
            Denuncias(DenunciasNumber).Autor = ReadField(1, Rdata, Asc(","))
            Denuncias(DenunciasNumber).Contenido = ReadField(2, Rdata, Asc(","))
            Denuncias(DenunciasNumber).ID = ReadField(3, Rdata, Asc(","))
            Denuncias(DenunciasNumber).YP = ReadField(4, Rdata, Asc(","))
            Denuncias(DenunciasNumber).Nick = ReadField(5, Rdata, Asc(","))
            Denuncias(DenunciasNumber).UltimoLogeo = ReadField(6, Rdata, Asc(","))
            Denuncias(DenunciasNumber).UltimaDenuncia = ReadField(7, Rdata, Asc(","))
            Denuncias(DenunciasNumber).PrimerDenuncia = ReadField(8, Rdata, Asc(","))
            Denuncias(DenunciasNumber).Estado = "NO LEIDO"
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "IREDAEL"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmClanes.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "IREDAEK"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmClanesUsuario.ParseUserInfo(Rdata)
            Exit Sub
        Case "IRFORTA"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            frmClanes.lblCastillos(4).Caption = Rdata
            frmClanesUsuario.txtCastillo(4).Text = Rdata
        Exit Sub
        Case "CLANDETSUB"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseSubGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "ENVFPS" 'envia fps del usuario
           Call SendData("ENVFPZ" & EnvioFPS)
        Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            If UserParalizado = False Then
                UserParalizado = True
                TiempoParalizado = 25
            ElseIf UserParalizado = True Then
                UserParalizado = False
                TiempoParalizado = 0
            End If
        Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                            frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                ii = 1
                Do While ii <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(ii).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    ii = ii + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        Case "BANCOBK"           ' Banco OK :: BANCOBK
            If frmNuevoBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                            frmNuevoBancoObj.List1(1).AddItem "" & Inventario.ItemName(i) & " - " & Inventario.Amount(i) & ""
                    Else
                            frmNuevoBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                ii = 1
                Do While ii <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventoryB(ii).OBJIndex <> 0 Then
                            frmNuevoBancoObj.List1(0).AddItem "" & UserBancoInventoryB(ii).Name & " - " & UserBancoInventoryB(ii).Amount & ""
                    Else
                            frmNuevoBancoObj.List1(0).AddItem "Nada"
                    End If
                    ii = ii + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                frmNuevoBancoObj.OroBove.Text = PonerPuntos(UserBancoOro)
                frmNuevoBancoObj.MiOro.Text = PonerPuntos(UserBancoOroPropio)
                
                'If ReadField(2, Rdata, 44) = "0" Then
                '        frmNuevoBancoObj.List1(0).ListIndex = frmNuevoBancoObj.LastIndex1
                'Else
                '        frmNuevoBancoObj.List1(1).ListIndex = frmNuevoBancoObj.LastIndex2
                'End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "TRAVELS"
          frmViajar.Show , frmMain
        Exit Sub
        Case "ABPANEL"
            frmPanelGm.Show vbModal, frmMain
            Exit Sub
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For i = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(i)
                Next i
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            End If
            Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(left$(Rdata, 9))
        Case "DAMEQUEST"
                Call frmQuestInfo.CargarList
        Exit Sub
    End Select
    
    ';Call LogCustom("Unhandled data: " & Rdata)
    
End Sub

Sub SendData(ByVal sdData As String)

    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

    Dim AuxCmd As String
    AuxCmd = UCase$(left$(sdData, 5))
    
    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If

#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub

Sub Login()
    If EstadoLogin = Normal Then
        SendData ("KERD22" & Val(HDSerial))
        SendData ("OOLOGI" & PJClickeado & "," & nombrecuent & "," & CodigoRecibido)
    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("KERD22" & Val(HDSerial))
        SendData ("NLOGIN" & UserName & "," & UserRaza & "," & UserSexo & "," & UserClase & "," & UserHogar _
                & "," & nombrecuent _
                & "," & Actualea)
     ElseIf EstadoLogin = CrearAccount Then
         With frmCuentas
            SendData ("NACCNT" & .Cuenta & "," & .Pass & "," & .Mail & "," & frmPasswdSinPadrinos.Text1 & "," & frmPasswdSinPadrinos.Text2)
        End With
 
    ElseIf EstadoLogin = BorrarPj Then
        SendData ("TBRP" & PJClickeado & "," & nombrecuent & "," & CodigoRecibido)
    ElseIf EstadoLogin = LoginAccount Then
        SendData ("KERD22" & Val(HDSerial))
        SendData ("ALOGIN" & nombrecuent & "," & UserPassword & "," & VersionC)
    End If
End Sub
