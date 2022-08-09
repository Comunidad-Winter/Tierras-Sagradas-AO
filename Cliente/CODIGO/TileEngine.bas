Attribute VB_Name = "Mod_TileEngine"
Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

Public indexList(0 To 5) As Integer
Public ibQuad As DxVBLibA.Direct3DIndexBuffer8
Public vbQuadIdx As DxVBLibA.Direct3DVertexBuffer8

Public lastTime As Long 'Para controlar la velocidad
Public fpsLastCheck As Long

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

Public AurixPJ As Grh
Public Octarina As Grh


Public Type tDonation
    Nombre As String
    Head As Integer
    Body As Integer
    GrhIndex As Integer
    Shield As Integer
    Weapon As Integer
    Casco As Integer
    Aura As Integer
    AuraA As Grh
    Aura_Angle As Single
End Type
Public picDonation As tDonation

'Apariencia del personaje
Public Type Char
    NPCNumber As Integer
    NPCAura As Byte
    NPCAuraG As Grh
    NPCAuraAngle As Single
    Pixel_X As Byte
    Pixel_Y As Byte
    Aura_IndexA As Integer
    AuraA As Grh
    Aura_AngleA As Single
    Aura_IndexW As Integer
    AuraW As Grh
    Aura_AngleW As Single
    Aura_IndexR As Integer
    AuraR As Grh
    Aura_AngleR As Single
    Aura_IndexE As Integer
    AuraE As Grh
    Aura_AngleE As Single
    Aura_IndexC As Integer
    AuraC As Grh
    Aura_AngleC As Single
    Navegando As Byte
    Montando As Byte
    Conquistandou As Byte
    esNW As Byte
    SinEnlistarHorda As Byte
    SinEnlistarAlianza As Byte
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    TieneRanking As Boolean
    PosRanking As Byte
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    Fx(1 To 3) As Grh
    FxIndex(1 To 3) As Integer
    EsStatus As Byte
    Ariete As Boolean
    
    'Color don/noble
    AntiguoR As Integer
    AntiguoG As Integer
    AntiguoB As Integer
    ProximoR As Integer
    ProximoB As Integer
    ProximoG As Integer
    LlegoAlColor As Boolean
    
    'Color auras
    AuraAntiguoR As Byte
    AuraAntiguoB As Byte
    AuraAntiguoG As Byte
    AuraQueremosLlegarR As Byte
    AuraQueremosLlegarG As Byte
    AuraQueremosLlegarB As Byte
    AuraProximoR As Byte
    AuraProximoG As Byte
    AuraProximoB As Byte
    AuraLlegoAlColor As Boolean
    
    Emoticon As Grh
    EmoticonIndex As Integer
    EmoticonLoops As Integer
    
    particle_count As Integer
    particle_group() As Long
    
    Criminal As Byte
    color As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    NPCMuerto(1 To 50) As Boolean
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    priv As Byte
    TransparenciaBody As Byte
    Llegoalatransp As Boolean
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group_index As Integer
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NpcDesv(1 To 50) As Byte
    
    light_value(3) As Long
    base_light(3) As Long
    
    luz As Integer
    color(3) As Long
    
    particle_group As Integer
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim EngineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long


Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type size
    CX As Long
    CY As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long
Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
End Sub
Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .EsStatus = 0
        .active = 0
        .Criminal = 0
        .FxIndex(1) = 0
        .FxIndex(2) = 0
        .FxIndex(3) = 0
        .invisible = False
        .Moving = 0
        .Muerto = False
        .Nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub
Sub EraseChar(ByVal CharIndex As Integer)

    On Error Resume Next

    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************

    charlist(CharIndex).active = 0

    'Update lastchar
    If CharIndex = LastChar Then

        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If

    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    Call ResetCharInfo(CharIndex)

    'Update NumChars
    NumChars = NumChars - 1

End Sub
Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .Muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

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
    
    
    If LegalPos(UserPos.X + X, UserPos.Y + Y) Then
        'Fill temp pos
        tX = UserPos.X + X
        tY = UserPos.Y + Y
        
        'Check to see if its out of bounds
        If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
            Exit Sub
        Else
            'Start moving... MainLoop does the rest
            AddtoUserPos.X = X
            UserPos.X = tX
            AddtoUserPos.Y = Y
            UserPos.Y = tY
            UserMoving = 1
            
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                    MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                    MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        End If
    End If
End Sub
Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
        '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        If charlist(MapData(X, Y).CharIndex).Muerto = False Then Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    LegalPos = True
End Function
Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************

    With charlist(CharIndex)
    
        If Fx = 0 Then
            Dim i As Long
            For i = 1 To 3
                .FxIndex(i) = 0
            Next i
            Exit Sub
        End If
            
    
        Dim slotLibre As Byte
        slotLibre = BuscarSlotLibreFx(CharIndex)
        .FxIndex(slotLibre) = Fx
       
        If .FxIndex(slotLibre) > 0 Then
            Call InitGrh(.Fx(slotLibre), FxData(Fx).Animacion)
       
            .Fx((slotLibre)).Loops = Loops
        End If
    End With
End Sub
Public Function BuscarSlotLibreFx(ByVal CharIndex As Integer) As Byte
    Dim i As Integer
    For i = 1 To 3
        If charlist(CharIndex).FxIndex(i) = 0 Then
            BuscarSlotLibreFx = i
            Exit Function
        End If
    Next i
End Function
Public Sub SetCharacterEmoticon(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .EmoticonIndex = Fx
        
        If .EmoticonIndex > 0 Then
            Call InitGrh(.Emoticon, FxData(Fx).Animacion)
            .EmoticonLoops = Loops
        End If
    End With
End Sub
Function EsArbol(ByVal GhrNumber As Long) As Boolean
EsArbol = (GhrNumber = 7222 Or _
    GhrNumber = 7223 Or _
    GhrNumber = 7224 Or _
    GhrNumber = 7225 Or _
    GhrNumber = 7226 Or _
    GhrNumber = 7000 Or _
    GhrNumber = 7001 Or _
    GhrNumber = 7002 Or _
    GhrNumber = 22077 Or _
    GhrNumber = 22078 Or _
    GhrNumber = 22079 Or _
    GhrNumber = 22080 Or _
    GhrNumber = 22081 Or _
    GhrNumber = 22082 Or _
    GhrNumber = 22083 Or _
    GhrNumber = 22084 Or _
    GhrNumber = 22085 Or _
    GhrNumber = 22086 Or _
    GhrNumber = 8489 Or _
    GhrNumber = 8483)
End Function

