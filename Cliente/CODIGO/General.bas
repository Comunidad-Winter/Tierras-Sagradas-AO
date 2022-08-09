Attribute VB_Name = "Mod_General"
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


Public bK As Long
Public bRK As Long

Public bFogata As Boolean

Public lFrameTimer As Long
Public Function DirPath(ByVal Path As String) As String
'•Parra: Nuevo Engine v2.0
    Select Case Path
        Case "Graficos"
            DirPath = App.Path & "\Data\GRAFICOS\"
            Exit Function
        
        Case "Sound"
            DirPath = App.Path & "\Data\SOUNDS\WAV\"
            Exit Function
        
        Case "Midi"
            DirPath = App.Path & "\Data\SOUNDS\MIDI\"
            Exit Function
        
        Case "Maps"
            DirPath = App.Path & "\Data\MAPAS\"
            Exit Function
    End Select
End Function

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\Data\" & "GRAFICOS" & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\Data\SOUNDS\" & "WAV" & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\Data\SOUNDS\" & "MIDI" & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\Data\" & "MAPAS" & "\"
End Function

Public Function SumaDigitos(ByVal Numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (Numero Mod 10)
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal Numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (Numero Mod 10) - 1
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function Complex(ByVal Numero As Integer) As Integer
    If Numero Mod 2 <> 0 Then
        Complex = Numero * SumaDigitos(Numero)
    Else
        Complex = Numero * SumaDigitosMenos(Numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal Numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(Numero)
    AuxInteger2 = SumaDigitosMenos(Numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim Arch As String
    
    Arch = App.Path & "\Data\INIT\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(Arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(Arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(Arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(Arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(Arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub
Public Sub CargarQuests()
        Dim p As Integer, loopc As Integer, LoopD
        Dim l_file As clsIniReader

        Set l_file = New clsIniReader
    
    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\QUESTS.dat"
        
        
        p = l_file.GetValue("INIT", "Num")
   
        ReDim InfoQuests(p) As tQuests
       
           
        For loopc = 1 To p
            InfoQuests(loopc).Nombre = l_file.GetValue("Quest" & loopc, "Nombre")
            InfoQuests(loopc).Tipo = l_file.GetValue("Quest" & loopc, "Tipo")
            InfoQuests(loopc).Info = l_file.GetValue("Quest" & loopc, "Info")
            InfoQuests(loopc).puntos = l_file.GetValue("Quest" & loopc, "Puntos")
            InfoQuests(loopc).Oro = l_file.GetValue("Quest" & loopc, "Oro")
            
            InfoQuests(loopc).Dificultad = l_file.GetValue("Quest" & loopc, "Dificultad")
            InfoQuests(loopc).NivelMinimo = l_file.GetValue("Quest" & loopc, "NivelMinimo")
            InfoQuests(loopc).Mapas = l_file.GetValue("Quest" & loopc, "Mapas")
            InfoQuests(loopc).PosiblesDrops = l_file.GetValue("Quest" & loopc, "PosiblesDrops")

                InfoQuests(loopc).NPCs = l_file.GetValue("Quest" & loopc, "NPCs")
                    
                    For LoopD = 1 To InfoQuests(loopc).NPCs
                        InfoQuests(loopc).NumNPC(LoopD) = ReadField(1, l_file.GetValue("Quest" & loopc, "Npc" & LoopD), Asc("-"))
                        InfoQuests(loopc).CantNPC(LoopD) = ReadField(2, l_file.GetValue("Quest" & loopc, "Npc" & LoopD), Asc("-"))
                    Next LoopD
            
            InfoQuests(loopc).IndexOBJ = ReadField(1, l_file.GetValue("Quest" & loopc, "OBJ"), Asc("-"))
            InfoQuests(loopc).CantOBJ = ReadField(2, l_file.GetValue("Quest" & loopc, "OBJ"), Asc("-"))
                
        Next loopc
End Sub
Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.Path & "\Data\INIT\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 47 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    

'NW
ColoresPJ(22).r = 255
ColoresPJ(22).g = 255
ColoresPJ(22).b = 202
'NW
'Poder
ColoresPJ(20).r = 225
ColoresPJ(20).g = 225
ColoresPJ(20).b = 225
'Poder
'Horda sin enlistar
ColoresPJ(47).r = 227
ColoresPJ(47).g = 141
ColoresPJ(47).b = 150
'Horda sin enlistar
'Alianza sin enlistar
ColoresPJ(46).r = 132
ColoresPJ(46).g = 193
ColoresPJ(46).b = 225
'Alianza sin enlistar
'Horda Enlistado
ColoresPJ(50).r = 255
ColoresPJ(50).g = 0
ColoresPJ(50).b = 0
'Horda Enlistado
'Alianza Enlistado
ColoresPJ(49).r = 0
ColoresPJ(49).g = 128
ColoresPJ(49).b = 255
'Alianza Enlistado
'Neutral
ColoresPJ(48).r = 125
ColoresPJ(48).g = 125
ColoresPJ(48).b = 125
'Neutral

'CONCILIO ALIANZA
ColoresPJ(51).r = 16
ColoresPJ(51).g = 38
ColoresPJ(51).b = 96

'CONCILIO HORDA
ColoresPJ(52).r = 69
ColoresPJ(52).g = 13
ColoresPJ(52).b = 14
End Sub
Sub CargarAuras()
    Dim archivoC As String, CantAuras As Integer
    archivoC = App.Path & "\Data\INIT\Auras.dat"
    CantAuras = GetVar(App.Path & "\Data\INIT\Auras.dat", "INIT", "NumAuras")
    
    If Not FileExist(archivoC, vbArchive) Then
        Call MsgBox("ERROR: no se ha podido cargar las auras. Falta el archivo auras.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    ReDim AurasPJ(CantAuras) As tAuras
    

    Dim XX As Long
    For XX = 1 To CantAuras
        AurasPJ(XX).GrhIndex = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "GrhIndex")
        AurasPJ(XX).r = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Rojo")
        AurasPJ(XX).g = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Verde")
        AurasPJ(XX).b = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Azul")
        AurasPJ(XX).Offset = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Offset")
        AurasPJ(XX).Giratoria = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "Giratoria")
        AurasPJ(XX).RojoF = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "RojoF")
        AurasPJ(XX).AzulF = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "AzulF")
        AurasPJ(XX).VerdeF = GetVar(App.Path & "\Data\INIT\Auras.dat", "AURA" & XX, "VerdeF")
    Next XX
    
End Sub
Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim Arch As String
    
    Arch = App.Path & "\Data\INIT\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(Arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(Arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(Arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(Arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(Arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

If UserConsola = 0 Then
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
    
        'RichTextBox.Refresh
    End With
    End If
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).charindex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
   
    If checkemail And UserEmail = "" Then
        Mensaje.Escribir "Dirección de email invalida"
        Exit Function
    End If
   
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
           Mensaje.Escribir "Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido."
            Exit Function
        End If
    Next loopc
   
    If nombrecuent = "" Then
        Mensaje.Escribir "Ingrese un nombre de cuenta."
        Exit Function
    End If
   
    If UserPassword = "" Then
        Mensaje.Escribir "Ingrese un password."
        Exit Function
    End If
    If Len(nombrecuent) > 30 Then
        Mensaje.Escribir "La cuenta debe tener menos de 30 letras."
        Exit Function
    End If
   
    For loopc = 1 To Len(nombrecuent)
        CharAscii = Asc(mid$(nombrecuent, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            Mensaje.Escribir "Cuenta inválida. El caractér " & Chr$(CharAscii) & " no está permitido."
            Exit Function
        End If
    Next loopc
   
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form

    For Each mifrm In Forms
        Unload mifrm

    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function
Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
 'Set Connected
    Connected = True
    
     Unload frmConnect
     
If UserLvl > 50 Then
    frmMain.LvlLbl.ForeColor = vbYellow
Else
    frmMain.LvlLbl.ForeColor = vbRed
End If

frmMain.Visible = True
Call DibujarPuntoMinimap
Call DibujarMinimap

Call AgregarParticulasyLuces(UserMap)

If TieneColorMapa = False Then
    day_r_old = 215
    day_g_old = 215
    day_b_old = 215
    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
End If
    
End Sub
Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    
    If Stopped = 1 Then Exit Sub
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
            If UserNavegando = False And HayAgua(UserPos.X, UserPos.Y - 1) = True Then LegalOk = False
            If UserNavegando = True And HayAgua(UserPos.X, UserPos.Y - 1) = False Then LegalOk = False
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
            If UserNavegando = False And HayAgua(UserPos.X + 1, UserPos.Y) = True Then LegalOk = False
            If UserNavegando = True And HayAgua(UserPos.X + 1, UserPos.Y) = False Then LegalOk = False
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
            If UserNavegando = False And HayAgua(UserPos.X, UserPos.Y + 1) = True Then LegalOk = False
            If UserNavegando = True And HayAgua(UserPos.X, UserPos.Y + 1) = False Then LegalOk = False
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
            If UserNavegando = False And HayAgua(UserPos.X - 1, UserPos.Y) = True Then LegalOk = False
            If UserNavegando = True And HayAgua(UserPos.X - 1, UserPos.Y) = False Then LegalOk = False
    End Select
    
   If LegalOk Then
            Call SendData("M" & Direccion)
            DibujarPuntoMinimap
            engine.Char_Move_by_Head UserCharIndex, Direccion
            MoveScreen Direccion
            UserMeditar = False
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            charlist(UserCharIndex).Heading = Direccion
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub
Sub CheckKeys()
Static LastMovement As Long
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
        
    If UserParalizado Then
    
        If GetTickCount() - LastMovement > 135 Then
                LastMovement = GetTickCount()
        Else
                Exit Sub
        End If
    
            If GetKeyState(BindKeys(14).KeyCode) < 0 Then
                If BindKeys(14).KeyCode <> 38 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 1 Then
                        Call SendData("CHEA" & 1)
                        charlist(UserCharIndex).Heading = 1
                        Exit Sub
                    End If
            End If
       
            'Move Right
            If GetKeyState(BindKeys(17).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If BindKeys(17).KeyCode <> 39 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 2 Then
                        Call SendData("CHEA" & 2)
                        charlist(UserCharIndex).Heading = 2
                        Exit Sub
                    End If
            End If
       
            'Move down
            If GetKeyState(BindKeys(15).KeyCode) < 0 Then
                If BindKeys(15).KeyCode <> 40 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 3 Then
                        Call SendData("CHEA" & 3)
                        charlist(UserCharIndex).Heading = 3
                        Exit Sub
                    End If
            End If
       
            'Move left
            If GetKeyState(BindKeys(16).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If BindKeys(16).KeyCode <> 37 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                    If charlist(UserCharIndex).Heading <> 4 Then
                        Call SendData("CHEA" & 4)
                        charlist(UserCharIndex).Heading = 4
                        Exit Sub
                    End If
            End If
            
        Exit Sub
    End If
    
        If GetTickCount() - LastMovement > 56 Then
                LastMovement = GetTickCount()
        Else
                Exit Sub
        End If
   
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
            'Move Up
            If GetKeyState(BindKeys(14).KeyCode) < 0 Then
                
                If BindKeys(14).KeyCode <> 38 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                
                Call MoveTo(NORTH)
                Call DibujarPuntoMinimap
                frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
       
            'Move Right
            If GetKeyState(BindKeys(17).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If BindKeys(17).KeyCode <> 39 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                
                Call MoveTo(EAST)
                Call DibujarPuntoMinimap
                frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
       
            'Move down
            If GetKeyState(BindKeys(15).KeyCode) < 0 Then
                If BindKeys(15).KeyCode <> 40 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                
                Call MoveTo(SOUTH)
                Call DibujarPuntoMinimap
                frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
       
            'Move left
            If GetKeyState(BindKeys(16).KeyCode) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If BindKeys(16).KeyCode <> 37 And frmMain.SendTxt.Visible = True Then Exit Sub
                
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                
                Call MoveTo(WEST)
                Call DibujarPuntoMinimap
                frmMain.Coord.Caption = Nombredelmapaxx & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
    End If
End Sub


'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
    
        Case E_Heading.EAST
            X = 1
    
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    Dim TempLng As Byte
    Dim TempByte1 As Byte
    Dim TempByte2 As Byte
    Dim TempByte3 As Byte
    Dim i As Byte

    'By Lorwik - www.rincondelao.com.ar
    engine.Particle_Group_Remove_All
    Light.Light_Remove_All
    handle = FreeFile()
    
    Open DirPath("Maps") & "Mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            For i = 0 To 3
                MapData(X, Y).light_value(i) = False
            Next i
            Get handle, , ByFlags
            MapData(X, Y).luz = 0
            MapData(X, Y).particle_group = 0
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            If ByFlags And 32 Then
               Get handle, , tempint
                'By Lorwik - www.rincondelao.com.ar
                MapData(X, Y).particle_group_index = General_Particle_Create(tempint, X, Y, -1)
            End If
            
            'Erase NPCs
            If MapData(X, Y).charindex > 0 Then
                Call EraseChar(MapData(X, Y).charindex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
    
    Close handle
    
    If Map = 999 And frmConnect.Visible = True Then
        'Call General_Particle_Create(6, 50, 49, -1)
        'Call General_Particle_Create(6, 41, 42, -1)
        'Call General_Particle_Create(6, 59, 42, -1)
        'Call General_Particle_Create(6, 41, 56, -1)
        'Call General_Particle_Create(6, 59, 56, -1)
    End If
    
    If Map = 998 And frmConnect.Visible = True Then
        'Call General_Particle_Create(6, 46, 61, -1)
        'Call General_Particle_Create(6, 54, 61, -1)
        'Call General_Particle_Create(6, 46, 50, -1)
        'Call General_Particle_Create(6, 54, 50, -1)
        Call General_Particle_Create(3, 38, 41, -1)
        Call General_Particle_Create(3, 62, 41, -1)
        
        Call General_Particle_Create(76, 56, 27)
        'Call General_Particle_Create(61, 48, 27)
        'Call General_Particle_Create(61, 51, 27)
        'Call General_Particle_Create(61, 54, 27)
        
        'Call General_Particle_Create(6, 46, 34, -1)
        'Call General_Particle_Create(6, 47, 25, -1)
        'Call General_Particle_Create(6, 53, 25, -1)
        'Call General_Particle_Create(6, 42, 25, -1)
        'Call General_Particle_Create(6, 36, 16, -1)
        'Call General_Particle_Create(6, 43, 16, -1)
        'Call General_Particle_Create(6, 58, 25, -1)
        'Call General_Particle_Create(6, 57, 16, -1)
        'Call General_Particle_Create(6, 64, 16, -1)
        'Light.Create_Light_To_Map 46, 35, 3, 255, 255, 255
        'Light.Create_Light_To_Map 54, 35, 3, 255, 255, 255
        'Light.Create_Light_To_Map 54, 43, 3, 255, 255, 255
        'Light.Create_Light_To_Map 46, 43, 3, 255, 255, 255
        'Light.Create_Light_To_Map 53, 26, 3, 100, 100, 100
        'Light.Create_Light_To_Map 47, 26, 3, 100, 100, 100
        'Light.Create_Light_To_Map 58, 26, 3, 100, 100, 100
        'Light.Create_Light_To_Map 42, 26, 3, 100, 100, 100
      Exit Sub
    End If

    If frmConnect.Visible = True Then Exit Sub
    
    Call AgregarParticulasyLuces(Map)
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    
    Call DibujarPuntoMinimap
    Call DibujarMinimap
                
End Sub
'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.Path & "\Data\INIT\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
On Error GoTo errorH
    Dim f As String
    Dim c As Integer
    Dim i As Long
    
    f = App.Path & "\Data\INIT\sinfo.dat"
    c = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For i = 1 To c
        ServersLst(i).Desc = GetVar(f, "S" & i, "Desc")
        ServersLst(i).Ip = Trim$(GetVar(f, "S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(GetVar(f, "S" & i, "P2"))
        ServersLst(i).Puerto = CInt(GetVar(f, "S" & i, "PJ"))
    Next i
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    End
End Sub
Public Function CurServerIp() As String

 CurServerIp = "45.235.98.111"
 'CurServerIp = "127.0.0.1"

End Function
Public Function CurServerPort() As Integer

        CurServerPort = "7200"

End Function
Sub Main()

On Error Resume Next

Dim strIconPath As String
Call frmCargando.ProgresoBarra(0)

strIconPath = App.Path & "\Data\icono.ico"
frmCargando.Icon = LoadPicture(strIconPath)
frmConnect.Icon = LoadPicture(strIconPath)
frmMain.Icon = LoadPicture(strIconPath)

HDSerial = GetDriveSerialNumber


    Dim i As Integer

    Call WriteClientVer

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    
    ChDrive App.Path
    ChDir App.Path
    
    Set Light = New clsLight
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8
        
    frmCargando.Show
     
    If MsgBox("¿Desea ejecutar el juego en modo ventana?", vbYesNo, "Tierras Sagradas AO") = vbNo Then
        Call Resolucion.SetResolucion
        PantallaCompleta = True
    End If
    
    Dim lc As Integer
    Dim LACONCHA As String
    lc = 0
    For lc = 1 To NUMBINDS
        LACONCHA = GetVar(App.Path & "\Data\INIT\" & "Teclas.tsao", "TECLAS", Str(lc))
        BindKeys(lc).KeyCode = Val(ReadField(1, LACONCHA, 44))
        BindKeys(lc).Name = ReadField(2, LACONCHA, 44)
    Next lc
    
    ClickeoTextCuenta = True
    TextBoxCuenta = ""
    TextBoxPassw = ""
    TextBoxPasswR = ""
    VersionC = GetVar(App.Path & "\Data\INIT\versiones.ini", "VERSION", "V")
    
    frmCargando.Refresh
    
    frmMain.Socket1.Startup
    Call InicializarNombres
    Call frmCargando.ProgresoBarra(10)
    UserMap = 1
    LoadGrhData
    CargarParticulas
    CargarCabezas
    CargarCascos
    CargarCuerpos
    Call frmCargando.ProgresoBarra(30)
    
    Call CargarParticulas
    CargarFxs
    Call engine.Engine_Init
    Call frmCargando.ProgresoBarra(50)
    

    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call frmCargando.ProgresoBarra(70)
    
    modTextos.InitFonts
    modTextos.LoadText
    Call frmCargando.ProgresoBarra(80)
    
    Call CargarColores
    Call CargarQuests
    Call CargarAuras
    Call General_Load_Interfaces
    Call OpcionesNew.LoadOptions
    Call frmCargando.ProgresoBarra(90)

    RangoPRIV(1) = "<Staff TSAO>"
    RangoPRIV(2) = "<Staff TSAO>"
    RangoPRIV(3) = "<Staff TSAO>"
    RangoPRIV(4) = "<Coordination>"
    RangoPRIV(5) = "<Development>"
    RangoPRIV(6) = "<Administrator>"
    
    EsStatusCOLOR(0) = D3DColorXRGB(ColoresPJ(48).r, ColoresPJ(48).g, ColoresPJ(48).b)
    EsStatusCOLOR(1) = D3DColorXRGB(ColoresPJ(46).r, ColoresPJ(46).g, ColoresPJ(46).b)
    EsStatusCOLOR(2) = D3DColorXRGB(ColoresPJ(47).r, ColoresPJ(47).g, ColoresPJ(47).b)
    EsStatusCOLOR(3) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
    EsStatusCOLOR(4) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
    EsStatusCOLOR(5) = D3DColorXRGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b)
    EsStatusCOLOR(6) = D3DColorXRGB(ColoresPJ(52).r, ColoresPJ(52).g, ColoresPJ(52).b)
    EsStatusCOLOR(8) = D3DColorXRGB(ColoresPJ(22).r, ColoresPJ(22).g, ColoresPJ(22).b)
    
    Call frmCargando.ProgresoBarra(100)
    
    Unload frmCargando
    
    'Inicializamos el sonido
    Call Audio.Initialize(frmMain.hWnd, App.Path & "\Data\SOUNDS\" & "WAV" & "\", App.Path & "\Data\SOUNDS\" & "MIDI" & "\")
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)


    frmPres.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Pres" & RandomNumber(1, 4) & ".jpg")
    frmPres.Timer1.Enabled = True
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    Call CambiarConectar("CONECTAR")

    frmMain.InvEqu.Picture = General_Load_Interface_Picture("Centronuevoinventario.jpg")
    frmMain.Picture = General_Load_Interface_Picture("Principal.jpg")
    
    Sound = Configuracion.Sound
    Musica = Configuracion.Music
    
    Dim IntroMusic As Byte
    IntroMusic = RandomNumber(1, 3)
    If IntroMusic = 1 Then
        Audio.MP3_Play "70"
    ElseIf IntroMusic = 2 Then
        Audio.MP3_Play "73"
    Else
        Audio.MP3_Play "140"
    End If

    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    Dialogos.font = frmMain.font
    
Engine_Set_TileBuffer 9
engine.Start
    
Exit Sub
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & err.HelpContext & " Desc:" & err.Description & " Fuente:" & err.Source
    End
End Sub
Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, Value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
                MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub
Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Pirata"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Recolector"
    ListaClases(12) = "Artesano"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.DefensaMagica) = "Defensa Magica"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub

Public Sub LogError(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Data\errores.log" For Append As #nfile
Print #nfile, Desc
Close #nfile
End Sub

Public Sub LogCustom(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Data\custom.log" For Append As #nfile
Print #nfile, Now & " " & Desc
Close #nfile
End Sub
Sub DameOpciones()
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)

Case "Hombre"

Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)

Case "Humano"
Actualea = 1
MaxEleccion = 30
MinEleccion = 1

Case "Elfo"
Actualea = 101
MaxEleccion = 113
MinEleccion = 101
                
Case "Elfo Oscuro"
Actualea = 202
MaxEleccion = 209
MinEleccion = 202
                
Case "Enano"
Actualea = 301
MaxEleccion = 305
MinEleccion = 301
                
Case "Gnomo"
Actualea = 401
MaxEleccion = 406
MinEleccion = 401
                
Case Else
Actualea = 30
MaxEleccion = 30
MinEleccion = 30
                
End Select
        
Case "Mujer"
   
Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)

Case "Humano"
Actualea = 70
MaxEleccion = 76
MinEleccion = 70
                
Case "Elfo"
Actualea = 170
MaxEleccion = 176
MinEleccion = 170
                
Case "Elfo Oscuro"
Actualea = 270
MaxEleccion = 280
MinEleccion = 270
                
Case "Gnomo"
Actualea = 470
MaxEleccion = 474
MinEleccion = 470
                
Case "Enano"
Actualea = 370
MaxEleccion = 373
MinEleccion = 370
            
Case Else
Actualea = 70
MaxEleccion = 70
MinEleccion = 70
                
End Select

End Select

Dim SR As RECT
SR.bottom = 32
SR.Right = 32
SR.left = 0
SR.top = 0
 
frmCrearPersonaje.headview.Cls
Call engine.DrawGrhtoHdc(HeadData(Actualea).Head(3).GrhIndex, SR, frmCrearPersonaje.headview, 8, 5)
 
End Sub
Public Sub DibujarPuntoMinimap()
    
With frmMain
.Puntito.left = UserPos.X - 2
.Puntito.top = UserPos.Y - 3
End With
    
End Sub
Public Sub DibujarMinimap()

If Configuracion.VerMiniMapa = 1 Then
    If FileExist(App.Path & "\Data\GRAFICOS\MiniMap\Mapa" & UserMap & ".bmp", vbNormal) Then
        frmMain.Minimap.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\MiniMap\Mapa" & UserMap & ".bmp")
        frmMain.Minimap.Refresh
    Else
        frmMain.Minimap.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\MiniMap\Nada.bmp")
    End If
End If

End Sub

Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String
 
Cifra = Str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", 0)
End Function
Public Function General_Load_Interface_Picture(ByVal PicName As String) As IPicture

On Error GoTo err
'vars
Dim GUIFolder As String
GUIFolder = App.Path & "\Data\GRAFICOS\Principal\"

'vemos si existe la interfas sino cargamos la default
If FileExist(GUIFolder & PicName, vbNormal) Then 'existe la cargamos
    Set General_Load_Interface_Picture = LoadPicture(GUIFolder & PicName)
    'Dest.Picture = LoadPicture(GUIFolder & PicName)
Else 'usamos la default
    Set General_Load_Interface_Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\" & PicName)
    'Dest.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\" & PicName)
End If

Exit Function

'error
err:
LogError "Error al cargar la imagen " & PicName & ", la imagen no se encontro."

End Function
Public Sub General_Load_Interfaces()

Dim N As Integer
Dim i As Integer

N = Val(GetVar(App.Path & "\Data\INIT\Interfaz.dat", "MAIN", "Interfaces"))

ReDim Interfaces(1 To N) As String

For i = 1 To N
    Interfaces(i) = GetVar(App.Path & "\Data\INIT\Interfaz.dat", "INTERFACES", "N" & i)
Next i

End Sub
Public Sub TirarItemMouse()
    Dim tX As Byte
    Dim tY As Byte
    Dim CantidadGG As String
    OfMouse = True
    Call ConvertCPtoTP(frmMain.MouseX, frmMain.MouseY, tX, tY)
    
    Dim Namepos As String, NameReal As String
    
If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
    
    If MapData(tX, tY).charindex > 0 Then
    
        'If MapData(tX, tY).charindex = charindex Then Exit Sub
        
            If charlist(MapData(tX, tY).charindex).NPCNumber = 36 Then
                'depositar
                If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                    SendData ("DEPO" & "," & Inventario.SelectedItem & "," & 1)
                ElseIf Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    CantidadGG = InputBox("Ingresa la cantidad de " & Inventario.ItemName(Inventario.SelectedItem) & " que quieras DEPOSITAR (0 para cancelar):", "¿Cantidad?", "0")
                    If Not IsNumeric(CantidadGG) Then Exit Sub
                    If CantidadGG = 0 Or CantidadGG > 10000 Then Exit Sub
                    SendData ("DEPO" & "," & Inventario.SelectedItem & "," & CantidadGG)
                End If
              Exit Sub
            End If
        
        
            Namepos = InStr(charlist(MapData(tX, tY).charindex).Nombre, "<")
            If Namepos = 0 Then Namepos = Len(charlist(MapData(tX, tY).charindex).Nombre) + 2
            NameReal = left$(charlist(MapData(tX, tY).charindex).Nombre, Namepos - 2)
    
        'transferir
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            If MsgBox("¿Transferirle 1 " & Inventario.ItemName(Inventario.SelectedItem) & " al usuario " & charlist(MapData(tX, tY).charindex).Nombre & "?", vbYesNo, "Confirmacion") = vbYes Then
                Call SendData("DYDTRA" & tX & "," & tY & "," & NameReal & "," & Inventario.SelectedItem & "," & 1)
            End If
        ElseIf Inventario.Amount(Inventario.SelectedItem) > 1 Then
            CantidadGG = InputBox("Ingresa la cantidad de " & Inventario.ItemName(Inventario.SelectedItem) & " que quieras TRANSFERIR a " & charlist(MapData(tX, tY).charindex).Nombre & " (0 para cancelar):", "¿Cantidad?", "0")
            If Not IsNumeric(CantidadGG) Then Exit Sub
            If CantidadGG = 0 Or CantidadGG > 10000 Then Exit Sub
            Call SendData("DYDTRA" & tX & "," & tY & "," & NameReal & "," & Inventario.SelectedItem & "," & CantidadGG)
        End If
        
           MouseRendOK = False
      Exit Sub
        
    Else
        'tirar
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call SendData("TR" & Inventario.SelectedItem & "," & 1 & "," & tX & "," & tY)
        ElseIf Inventario.Amount(Inventario.SelectedItem) > 1 Then
                CantidadGG = InputBox("Ingresa la cantidad de " & Inventario.ItemName(Inventario.SelectedItem) & " que quieras TIRAR (0 para cancelar):", "¿Cantidad?", "0")
                If Not IsNumeric(CantidadGG) Then Exit Sub
                If CantidadGG = 0 Or CantidadGG > 10000 Then Exit Sub
                Call SendData("TR" & Inventario.SelectedItem & "," & CantidadGG & "," & tX & "," & tY)
        End If
    End If
End If

End Sub
Public Sub CambiarConectar(Tuvieja As String)

frmConnect.Visible = True

If UCase$(Tuvieja) = "CONECTAR" Then
          With frmConnect
                .imgName.Visible = True
                .imgPass.Visible = True
                .imgConectar.Visible = True
                .imgAnti.Visible = True
                .imgCrearCuenta.Visible = True
                .imgRecuperarCuenta.Visible = True
                .imgWeb.Visible = True
                
                .imgCambiarPass.Visible = False
                .imgCrearPersonaje.Visible = False
                .imgSalir4.Visible = False
            
                Dim i As Long
                For i = 0 To 8
                .PJ(i).Visible = False
                Next i
          End With
          
        For i = 0 To 8
                CargarPJ(i).Nombre = 0
                CargarPJ(i).Body = 0
                CargarPJ(i).Head = 0
                CargarPJ(i).Casco = 0
                CargarPJ(i).Shield = 0
                CargarPJ(i).Weapon = 0
                CargarPJ(i).Level = 0
                CargarPJ(i).Existe = False
                CargarPJ(i).Raza = 0
                CargarPJ(i).Muerto = 0
        Next i
          
          RenderAccount = False
          RenderConnect = True

     If FileExist(App.Path & "\Data\MAPAS\" & "Mapa" & MapConnect & ".map", vbNormal) Then
        Call SwitchMap(MapConnect)
         day_r_old = 210
         day_g_old = 100
         day_b_old = 100
         base_light = ARGB(day_r_old, day_g_old, day_b_old, 220)
    Else
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        Call UnloadAllForms
    End If
    
ElseIf UCase$(Tuvieja) = "CUENTA" Then

          With frmConnect
                .imgName.Visible = False
                .imgPass.Visible = False
                .imgConectar.Visible = False
                .imgAnti.Visible = False
                .imgCrearCuenta.Visible = False
                .imgRecuperarCuenta.Visible = False
                .imgWeb.Visible = False
                
                .imgCambiarPass.Visible = True
                .imgCrearPersonaje.Visible = True
                .imgSalir4.Visible = True
            
                For i = 0 To 8
                .PJ(i).Visible = True
                Next i
          End With
          
            RenderConnect = False
            RenderAccount = True
            
            Audio.StopWave
            
             If FileExist(App.Path & "\Data\MAPAS\" & "Mapa" & MapCuent & ".map", vbNormal) Then
                    day_r_old = 220
                    day_g_old = 220
                    day_b_old = 220
                    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
                            
                    Call InitGrh(AurixPJ, 27601)
                    
                    Call SwitchMap(MapCuent)
            Else
                MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
                Call UnloadAllForms
            End If
End If



End Sub
Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
'***************************************************
'Author: Nahuel Casas (Zagen)
'Last Modify Date: 07/12/2009
' 07/12/2009: Zagen - Convertì las funciones, en formulas mas fàciles de modificar.
'***************************************************
    On Error Resume Next
          Dim fso As Object, Drv As Object, DriveSerial As Long
         
          'Creamos el objeto FileSystemObject.
          Set fso = CreateObject("Scripting.FileSystemObject")
         
          'Asignamos el driver principal.
          If DriveLetter <> "" Then
              Set Drv = fso.GetDrive(DriveLetter)
          Else
              Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
          End If
     
          With Drv
              If .IsReady Then
                  DriveSerial = Abs(.SerialNumber)
              Else    '"Si el driver no està como para empezar ..."
                  DriveSerial = -1
              End If
          End With
         
          'Borramos y limpiamos.
          Set Drv = Nothing
          Set fso = Nothing
    'Seteamos :)
    GetDriveSerialNumber = DriveSerial
         
End Function
Public Sub DarColorCambiante(ByVal charindex As Long)

With charlist(charindex)
    Select Case .EsStatus
        Case 0 '125 125 125
            .AntiguoR = 125
            .AntiguoG = 125
            .AntiguoB = 125
            
        Case 1
            .AntiguoR = 132
            .AntiguoG = 193
            .AntiguoB = 225
            
        Case 2 '227 141 150
            .AntiguoR = 227
            .AntiguoG = 141
            .AntiguoB = 150
            
        Case 3 '0 128 255
            .AntiguoR = 0
            .AntiguoG = 128
            .AntiguoB = 255
        Case 4 '255 0 0
            .AntiguoR = 255
            .AntiguoG = 0
            .AntiguoB = 0
        
        Case 5
            .AntiguoR = 16
            .AntiguoG = 38
            .AntiguoB = 96
            
        Case 6 '69 13 14
            .AntiguoR = 69
            .AntiguoG = 13
            .AntiguoB = 14
            
    End Select
    
            If .color = 40 Then
                'Le damos el color original directamente
                If .ProximoR = 0 And .ProximoG = 0 And .ProximoB = 0 Then
                    .ProximoR = 255
                    .ProximoG = 255
                    .ProximoB = 0
                End If
            
                    'Si ya supero el máximo le damos directamente l color y empezamos a darle paso al azul.
                    If .ProximoR >= 255 And .ProximoG >= 255 And .LlegoAlColor = False Then
                        .ProximoR = 255
                        .ProximoG = 255
                        .ProximoB = 0
                        .LlegoAlColor = True
                    End If
                
                'Empezamos a darle color amarillo
                If ((.ProximoR < 255) Or (.ProximoG < 255)) And .LlegoAlColor = False Then
                    .ProximoR = .ProximoR + 1
                    .ProximoG = .ProximoG + 1
                    .ProximoB = .ProximoB - 1
                                       
                    If .ProximoR >= 255 Then .ProximoR = 255
                    If .ProximoG >= 255 Then .ProximoG = 255
                    If .ProximoB < 0 Then .ProximoB = 0
                    
                'Si ya llego al amarillo, empezamos a darle el color gris.
                ElseIf .LlegoAlColor = True Then
                
                    .ProximoR = .ProximoR - 1
                    .ProximoG = .ProximoG - 1
                    .ProximoB = .ProximoB + 1
                    
                    If .ProximoR <= .AntiguoR Then .ProximoR = .AntiguoR
                    If .ProximoG <= .AntiguoG Then .ProximoG = .AntiguoG
                    If .ProximoB >= .AntiguoB Then .ProximoB = .AntiguoB
                    
                    'Ya llegamos al gris, vamos a darle paso al amarillo
                    If .ProximoR = .AntiguoR And .ProximoG = .AntiguoG And .ProximoB = .AntiguoB Then
                        .LlegoAlColor = False
                    End If
                
                End If
            ElseIf .color = 42 Then
                'Le damos el color original directamente
                If .ProximoR = 0 And .ProximoG = 0 And .ProximoB = 0 Then
                    .ProximoR = 255
                    .ProximoG = 255
                    .ProximoB = 255
                End If
            
                    'Si ya supero el máximo le damos directamente l color y empezamos a darle paso al azul.
                    If .ProximoR >= 255 And .ProximoG >= 255 And .ProximoB >= 255 And .LlegoAlColor = False Then
                        .ProximoR = 255
                        .ProximoG = 255
                        .ProximoB = 255
                        .LlegoAlColor = True
                    End If
                
                'Empezamos a darle color amarillo
                If ((.ProximoR < 255) Or (.ProximoG < 255) Or (.ProximoB < 255)) And .LlegoAlColor = False Then
                    .ProximoR = .ProximoR + 1
                    .ProximoG = .ProximoG + 1
                    .ProximoB = .ProximoB + 1
                                       
                    If .ProximoR >= 255 Then .ProximoR = 255
                    If .ProximoG >= 255 Then .ProximoG = 255
                    If .ProximoB >= 255 Then .ProximoB = 255
                    
                'Si ya llego al amarillo, empezamos a darle el color gris.
                ElseIf .LlegoAlColor = True Then
                
                    .ProximoR = .ProximoR - 1
                    .ProximoG = .ProximoG - 1
                    .ProximoB = .ProximoB - 1
                    
                    If .ProximoR <= .AntiguoR Then .ProximoR = .AntiguoR
                    If .ProximoG <= .AntiguoG Then .ProximoG = .AntiguoG
                    If .ProximoB <= .AntiguoB Then .ProximoB = .AntiguoB
                    
                    'Ya llegamos al gris, vamos a darle paso al amarillo
                    If .ProximoR = .AntiguoR And .ProximoG = .AntiguoG And .ProximoB = .AntiguoB Then
                        .LlegoAlColor = False
                    End If
                
                End If
                
            ElseIf .color = 41 Then
                'Le damos el color original directamente
                If .ProximoR = 0 And .ProximoG = 0 And .ProximoB = 0 Then
                    .ProximoR = 95
                    .ProximoG = 45
                    .ProximoB = 95
                End If
                
                'Empezamos a darle color amarillo
                If .LlegoAlColor = False Then
                    'ROJO
                    If (.ProximoR < 95) Then
                        .ProximoR = .ProximoR + 1
                        If .ProximoR >= 95 Then .ProximoR = 95
                    End If
                    If (.ProximoR > 95) Then
                        .ProximoR = .ProximoR - 1
                        If .ProximoR <= 95 Then .ProximoR = 95
                    End If
                    'ROJO
                    
                    'VERDE
                    If (.ProximoG < 45) Then
                        .ProximoG = .ProximoG + 1
                        If .ProximoG >= 45 Then .ProximoG = 45
                    End If
                    If (.ProximoG > 45) Then
                        .ProximoG = .ProximoG - 1
                        If .ProximoG <= 45 Then .ProximoG = 45
                    End If
                    'VERDE
                    
                    'AZUL
                    If (.ProximoB < 95) Then
                        .ProximoB = .ProximoB + 1
                        If .ProximoB >= 95 Then .ProximoB = 95
                    End If
                    If (.ProximoB > 95) Then
                        .ProximoB = .ProximoB - 1
                        If .ProximoB >= 95 Then .ProximoB = 95
                    End If
                    'AZUL
                    
                    'Si ya supero el máximo le damos directamente l color y empezamos a darle paso al azul.
                    If .ProximoR = 95 And .ProximoG = 45 And .ProximoB = 95 And .LlegoAlColor = False Then
                        .ProximoR = 95
                        .ProximoG = 45
                        .ProximoB = 95
                        .LlegoAlColor = True
                    End If
                    
                'Si ya llego al amarillo, empezamos a darle el color gris.
                ElseIf .LlegoAlColor = True Then
                
                    'ROJO
                    If (.ProximoR < .AntiguoR) Then
                        .ProximoR = .ProximoR + 1
                        If .ProximoR >= .AntiguoR Then .ProximoR = .AntiguoR
                    End If
                    If (.ProximoR > .AntiguoR) Then
                        .ProximoR = .ProximoR - 1
                        If .ProximoR <= .AntiguoR Then .ProximoR = .AntiguoR
                    End If
                    'ROJO
                    
                    'VERDE
                    If (.ProximoG < .AntiguoG) Then
                        .ProximoG = .ProximoG + 1
                        If .ProximoG >= .AntiguoG Then .ProximoG = .AntiguoG
                    End If
                    If (.ProximoG > .AntiguoG) Then
                        .ProximoG = .ProximoG - 1
                        If .ProximoG <= .AntiguoG Then .ProximoG = .AntiguoG
                    End If
                    'VERDE
                    
                    'AZUL
                    If (.ProximoB < .AntiguoB) Then
                        .ProximoB = .ProximoB + 1
                        If .ProximoB >= .AntiguoB Then .ProximoB = .AntiguoB
                    End If
                    If (.ProximoB > .AntiguoB) Then
                        .ProximoB = .ProximoB - 1
                        If .ProximoB >= .AntiguoB Then .ProximoB = .AntiguoB
                    End If
                    'AZUL
                    
                    'Ya llegamos al gris, vamos a darle paso al amarillo
                    If .ProximoR = .AntiguoR And .ProximoG = .AntiguoG And .ProximoB = .AntiguoB Then
                        .LlegoAlColor = False
                    End If
                
                End If
                
            End If
    
End With

End Sub
Private Function LeerInt(ByVal Ruta As String) As Integer
Dim f As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function
