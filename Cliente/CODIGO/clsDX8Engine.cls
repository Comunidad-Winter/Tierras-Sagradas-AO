VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'particulas
Dim base_tile_size As Integer
Private bTechoAB As Byte

Dim temp_verts(3)       As TLVERTEX

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
 
Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Private Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    ID As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
   
    Speed As Single
    life_counter As Long
End Type
 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
' / Viaje a chars.
    Bind_To_Char    As Integer
    Bind_Speed      As Single
    Now_Viaje_X     As Integer
    Now_Viaje_Y     As Integer
    active As Boolean
    ID As Long
    map_x As Integer
    map_y As Integer
    char_index As Long
 
    frame_counter As Single
    frame_speed As Single
   
    stream_type As Byte
 
    particle_stream() As Particle
    particle_count As Long
   
    grh_index_list() As Long
    grh_index_count As Long
   
    alpha_blend As Boolean
   
    alive_counter As Long
    never_die As Boolean
   
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
   
    'Added by Juan Martín Sotuyo Dodero
    Speed As Single
    life_counter As Long
   
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
 
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
'/particulas

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const COLOR_KEY As Long = &HFF000000

Dim font_count As Long
Dim font_last As Long

Private font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Dim Texture As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Public timerTicksPerFrame As Double
Private particletimer As Single
Private EngineBaseSpeed As Single
Private lFrameTimer As Long
Private lframelimiter As Long
Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop As Integer
Private MainViewLeft As Integer
Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer
Private Const GrhFogata As Integer = 1521

Private Type Flecha
    X As Single
    Y As Single
    xb As Single
    yb As Single
    angle As Single
    Grh As Grh
    Rotacion As Byte
    EnUso As Byte
End Type
 
Private Const MaxFlecha As Byte = 15
Private Flechas_list(MaxFlecha) As Flecha

Dim bump_map_supported As Boolean
Private Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time    As Currency

    Static end_time   As Currency

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
Public Sub Text_Render(ByVal font As D3DXFont, ByVal text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal Shadow As Boolean = False, Optional ByVal Nombre As Byte, Optional ByVal rojo As Byte, Optional ByVal verde As Byte, Optional ByVal azul As Byte)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + Height
    TextRect.Right = left + Width
        
    If Shadow Then
        ShadowRect.top = top - 1
        ShadowRect.left = left - 2
        ShadowRect.bottom = (top + Height) - 1
        ShadowRect.Right = (left + Width) - 2
    End If
    
DirectD3D8.DrawText font, color, text, TextRect, format
End Sub
Public Sub Text_Render_Inventario(ByVal font As D3DXFont, ByVal text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal Shadow As Boolean = False)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
' editado por: juanjo
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + Height
    TextRect.Right = left + Width
        
    If Shadow Then
        ShadowRect.top = top - 1
        ShadowRect.left = left - 2
        ShadowRect.bottom = (top + Height) - 1
        ShadowRect.Right = (left + Width) - 2
       'DirectD3D8.DrawText font, &HFF000000, Text, ShadowRect, format
    'DirectD3D8.DrawText font, D3DColorXRGB(0, 255, 0), Text, ShadowRect, format
    End If
    
    DirectD3D8.DrawText font, color, text, TextRect, format
    
End Sub
Public Sub Text_Render_ext(ByVal text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, Optional ByVal Shadow As Boolean = False, Optional ByVal Center As Boolean = False)
    If Center = True Then
        Call Text_Render_Inventario(font_list(1), text, top, left, Width, Height, color, fa_center, Shadow)
    Else
        Call Text_Render_Inventario(font_list(1), text, top, left, Width, Height, color, DT_TOP Or DT_LEFT, Shadow)
    End If
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = DirectD3D8.CreateFont(DirectDevice, font_desc.hFont)
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
  MakeVector.X = X
  MakeVector.Y = Y
  MakeVector.Z = Z
End Function
Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set SurfaceDB = New clsSurfaceManDynDX8

    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate()
    Set DirectD3D8 = New D3DX8
    
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

        With D3DWindow
            .Windowed = True
            .SwapEffect = D3DSWAPEFFECT_COPY
            .BackBufferFormat = DispMode.format
            .BackBufferWidth = 800
            .BackBufferHeight = 600
            .EnableAutoDepthStencil = 1
            .AutoDepthStencilFormat = D3DFMT_D16
            .hDeviceWindow = frmMain.renderer.hWnd
        End With
        
    DispMode.format = D3DFMT_X8R8G8B8

    If DirectD3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then

        Dim Caps8 As D3DCAPS8

        DirectD3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
        Else
            DispMode.format = DispModeBK.format
        End If

    Else
        DispMode.format = DispModeBK.format
    End If
    
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    Texto.Engine_Init_FontSettings
    Texto.Engine_Init_FontTextures
                                                            
    With DirectDevice
        
        '//Transformed and lit vertices dont need lighting
        '   so we disable it...
        .SetRenderState D3DRS_LIGHTING, False
        
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        'Particles
        .SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    End With
    
    Call SurfaceDB.Init(DirectD3D8, DirectDevice, 90)
    
    Call InitTileEngine
    
    Font_Create "Tahoma", 8, True, 0
    
    bRunning = True

End Sub

Public Sub InitTileEngine()
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************

    IniPath = App.Path & "\Init\"
    
    TilePixelWidth = 32
    TilePixelHeight = 32
    WindowTileHeight = Round(frmMain.renderer.Height / 32, 0)
    WindowTileWidth = Round(frmMain.renderer.Width / 32, 0)
    
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    
    EngineBaseSpeed = 0.0185
    
    'Set FPS value to 60 for startup
    FPS = 101
    FramesPerSecCounter = 101
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    'Index Buffer. Dunkan
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
    
    Set ibQuad = DirectDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
    
    Set vbQuadIdx = DirectDevice.CreateVertexBuffer(Len(temp_verts(0)) * 4, 0, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1, D3DPOOL_MANAGED)
    
End Sub

Public Sub Engine_Deinit()
Erase MapData
Erase charlist

Set DirectDevice = Nothing
Set DirectD3D = Nothing
Set DirectX = Nothing
End
End Sub
Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
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
Public Sub Draw_Conteo(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer, DestWidth As Integer, DestHeight As Integer, Alpha As Byte)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorARGB(Alpha, 255, 255, 255)
    rgb_list(1) = D3DColorARGB(Alpha, 255, 255, 255)
    rgb_list(2) = D3DColorARGB(Alpha, 255, 255, 255)
    rgb_list(3) = D3DColorARGB(Alpha, 255, 255, 255)
    
        If GrhData(grh_index).TileWidth <> 1 Then
            X = X - Int(GrhData(grh_index).TileWidth * (32 \ 2)) + 32 \ 2
        End If
 
        If GrhData(grh_index).TileHeight <> 1 Then
            Y = Y - Int(GrhData(grh_index).TileHeight * 32) + 32
        End If
    
    Device_Box_Textured_Render_Advance grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, DestWidth, DestHeight
        
End Sub
Public Sub Draw_GrhIndex(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
If frmConnect.Visible = True Then
    rgb_list(0) = D3DColorARGB(220, 255, 255, 255)
    rgb_list(1) = D3DColorARGB(220, 255, 255, 255)
    rgb_list(2) = D3DColorARGB(220, 255, 255, 255)
    rgb_list(3) = D3DColorARGB(220, 255, 255, 255)
Else
    rgb_list(0) = D3DColorARGB(255, 255, 255, 255)
    rgb_list(1) = D3DColorARGB(255, 255, 255, 255)
    rgb_list(2) = D3DColorARGB(255, 255, 255, 255)
    rgb_list(3) = D3DColorARGB(255, 255, 255, 255)
End If
    
    Device_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
        
End Sub
Public Sub Draw_GrhInv(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, True
End Sub
Private Sub Draw_Aura(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Aura_Index = 1, Optional ByVal UsuarioIndex As Integer)
 
Dim CurrentGrhIndex As Integer
Dim Light(3) As Long
 
If Grh.GrhIndex = 0 Then Exit Sub
 
If Animate Then
If Grh.Started = 1 Then
Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
 
If Grh.Loops <> -1 Then
If Grh.Loops > 0 Then
Grh.Loops = Grh.Loops - 1
Else
Grh.Started = 0
End If
End If
End If
End If
End If
 
'Figure out what frame to draw (always 1 if not animated)
CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
'Center Grh over X,Y pos
If Center Then
 
If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
End If
 
If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
End If
 
End If

Dim ColorRojo As Byte
Dim ColorVerde As Byte
Dim ColorAzul As Byte


If AurasPJ(Aura_Index).RojoF > 0 And UsuarioIndex > 0 Then

    With charlist(UsuarioIndex)
               If .AuraLlegoAlColor = False Then
                    'Empezamos a darle colores - rojo
                    If (.AuraQueremosLlegarR < .AuraAntiguoR) Then
                        .AuraProximoR = .AuraProximoR - 1
                        If .AuraProximoR <= .AuraQueremosLlegarR Then .AuraProximoR = .AuraQueremosLlegarR
                    ElseIf (.AuraQueremosLlegarR > .AuraAntiguoR) And .AuraLlegoAlColor = False Then
                        .AuraProximoR = .AuraProximoR + 1
                        If .AuraProximoR >= .AuraQueremosLlegarR Then .AuraProximoR = .AuraQueremosLlegarR
                    End If
                        
                    'VERDE
                    If (.AuraQueremosLlegarG < .AuraAntiguoG) Then
                        .AuraProximoG = .AuraProximoG - 1
                        If .AuraProximoG <= .AuraQueremosLlegarG Then .AuraProximoG = .AuraQueremosLlegarG
                    ElseIf (.AuraQueremosLlegarG > .AuraAntiguoG) Then
                        .AuraProximoG = .AuraProximoG + 1
                        If .AuraProximoG >= .AuraQueremosLlegarG Then .AuraProximoG = .AuraQueremosLlegarG
                    End If
                    
                    'AZUL
                    If (.AuraQueremosLlegarB < .AuraAntiguoB) Then
                        .AuraProximoB = .AuraProximoB - 1
                        If .AuraProximoB <= .AuraQueremosLlegarB Then .AuraProximoB = .AuraQueremosLlegarB
                    ElseIf (.AuraQueremosLlegarB > .AuraAntiguoB) Then
                        .AuraProximoB = .AuraProximoB + 1
                        If .AuraProximoB >= .AuraQueremosLlegarB Then .AuraProximoB = .AuraQueremosLlegarB
                    End If
                    
                    If .AuraProximoR = .AuraQueremosLlegarR And .AuraProximoG = .AuraQueremosLlegarG And .AuraProximoB = .AuraQueremosLlegarB Then
                        .AuraLlegoAlColor = True
                    End If
                    
                'Si ya llego al color, empezamos a darle el antiguo.
                ElseIf .AuraLlegoAlColor = True Then
                    'ROJO
                    If (.AuraProximoR < .AuraAntiguoR) Then
                        .AuraProximoR = .AuraProximoR + 1
                        If .AuraProximoR <= .AuraProximoR Then .AuraProximoR = .AuraAntiguoR
                    ElseIf (.AuraProximoR > .AuraAntiguoR) Then
                        .AuraProximoR = .AuraProximoR - 1
                        If .AuraProximoR >= .AuraProximoR Then .AuraProximoR = .AuraAntiguoR
                    End If
                        
                    'VERDE
                    If (.AuraProximoG < .AuraAntiguoG) Then
                        .AuraProximoG = .AuraProximoG + 1
                        If .AuraProximoG >= .AuraProximoG Then .AuraProximoG = .AuraAntiguoG
                    ElseIf (.AuraProximoG > .AuraAntiguoG) Then
                        .AuraProximoG = .AuraProximoG - 1
                        If .AuraProximoG <= .AuraProximoG Then .AuraProximoG = .AuraAntiguoG
                    End If
                    
                    'AZUL
                    If (.AuraProximoB < .AuraAntiguoB) Then
                        .AuraProximoB = .AuraProximoB + 1
                        If .AuraProximoB >= .AuraProximoB Then .AuraProximoB = .AuraAntiguoB
                    ElseIf (.AuraProximoB > .AuraAntiguoB) Then
                        .AuraProximoB = .AuraProximoB - 1
                        If .AuraProximoB <= .AuraProximoB Then .AuraProximoB = .AuraAntiguoB
                    End If
                    
                    If .AuraProximoR = .AuraAntiguoR And .AuraProximoG = .AuraAntiguoG And .AuraProximoB = .AuraAntiguoB Then
                        .AuraLlegoAlColor = False
                    End If
                End If
                    
                ColorRojo = .AuraProximoR
                ColorVerde = .AuraProximoG
                ColorAzul = .AuraProximoB
            End With
    
Else
    ColorRojo = AurasPJ(Aura_Index).r
    ColorVerde = AurasPJ(Aura_Index).g
    ColorAzul = AurasPJ(Aura_Index).b
End If

If Aura_Index = 98 Then
    Light(0) = D3DColorXRGB(255, 255, 255)
    Light(1) = Light(0)
    Light(2) = Light(0)
    Light(3) = Light(0)
Else
    Light(0) = D3DColorXRGB(ColorRojo, ColorVerde, ColorAzul)
    Light(1) = Light(0)
    Light(2) = Light(0)
    Light(3) = Light(0)
End If
 
Device_Box_Textured_Render CurrentGrhIndex, _
X, Y, _
GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
Light(), GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
Alpha _
, angle
 
End Sub
Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False, Optional alphaa As Byte = 255)
On Error Resume Next
    Dim CurrentGrhIndex As Integer
    
    If Grh.GrhIndex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
        
            If Not Invert_y Then
                Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
            Else
                Grh.FrameCounter = Grh.FrameCounter
            End If
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If
 
 Device_Box_Textured_Render CurrentGrhIndex, _
        X, Y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        MapData(map_x, map_y).light_value, _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        Alpha _
        , angle, 0, Invert_x, Invert_y, alphaa
'exits:
 
End Sub
Public Sub Render()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************

Static re As RECT
    re.left = 0
    re.top = 0
    re.bottom = frmMain.renderer.ScaleHeight
    re.Right = frmMain.renderer.ScaleWidth

DirectDevice.BeginScene
DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    
    ShowNextFrame
    
    If Configuracion.FPS = 65 And FPS > 65 Then
        frmMain.FPSMain.Caption = "65"
    Else
        frmMain.FPSMain.Caption = FPS
    End If
    
    EnvioFPS = FPS
    
    Call Dialogos.Render
    
    If GetTickCount >= (fpsLastCheck + 1000) Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        fpsLastCheck = GetTickCount
        lFrameTimer = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If


    lframelimiter = GetTickCount
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * EngineBaseSpeed
    particletimer = timerElapsedTime * 0.05
    
DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
DirectDevice.EndScene
DirectDevice.Present re, ByVal 0, 0, ByVal 0

If DibujadoContinuoInv = True And MouseRendOK = False And DyDActivado = True Then engine.DrawInv
    
End Sub
Sub ShowNextFrame()
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
End Sub
Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim ColorTechos(3) As Long
    
    
    'ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - Engine_Get_TileBuffer
    maxY = screenmaxY + Engine_Get_TileBuffer
    minX = screenminX - Engine_Get_TileBuffer
    maxX = screenmaxX + Engine_Get_TileBuffer
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1

    '<----- Layer 1, 2 ----->
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
        
            If Map_InBounds(X, Y) Then
                'Layer 1
                    Call Draw_Grh(MapData(X, Y).Graphic(1), _
                            (ScreenX - 1) * 32 + PixelOffsetX, _
                            (ScreenY - 1) * 32 + PixelOffsetY, _
                            0, 1, , X, Y)
                            
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, , X, Y)
                End If
            End If
            
            ScreenX = ScreenX + 1
        Next X
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
   Dim j As Integer
    Dim angle As Single
    Dim xb As Integer
    Dim yb As Integer
   
        For j = 0 To MaxFlecha
            If Flechas_list(j).EnUso = 1 Then
               
                'Update the position
                angle = DegreeToRadian * Engine_GetAngle(Flechas_list(j).X, Flechas_list(j).Y, Flechas_list(j).xb, Flechas_list(j).yb)
                Flechas_list(j).X = (Flechas_list(j).X + (Sin(angle) * timerElapsedTime * 0.63))
                Flechas_list(j).Y = (Flechas_list(j).Y - (Cos(angle) * timerElapsedTime * 0.63))
               
                'Update the rotation
                If Flechas_list(j).Rotacion > 0 Then
                    Flechas_list(j).angle = Flechas_list(j).angle + (Flechas_list(j).Rotacion * timerElapsedTime * 0.1)
                    Do While Flechas_list(j).angle > 360
                        Flechas_list(j).angle = Flechas_list(j).angle - 360
                    Loop
                End If
 
                'Draw if within range
                xb = ((-minX - 1) * 32) + Flechas_list(j).X + PixelOffsetX - 250 '((10 - 32) * 32)
                yb = ((-minY - 1) * 32) + Flechas_list(j).Y + PixelOffsetY - 250 '((10 - 32) * 32)
                If yb >= -32 Then
                    If yb <= (ScreenHeight + 32) Then
                        If xb >= -32 Then
                            If xb <= (ScreenWidth + 32) Then
                                    If (Flechas_list(j).X > Flechas_list(j).xb) And (Flechas_list(j).Y > Flechas_list(j).yb) Then
                                        Call Draw_Grh(Flechas_list(j).Grh, xb, yb, 0, 1, , X, Y, 23.5)
                                    ElseIf (Flechas_list(j).X > Flechas_list(j).xb) And (Flechas_list(j).Y < Flechas_list(j).yb) Then
                                        Call Draw_Grh(Flechas_list(j).Grh, xb, yb, 0, 1, , X, Y, 110)
                                    ElseIf (Flechas_list(j).X < Flechas_list(j).xb) And (Flechas_list(j).Y < Flechas_list(j).yb) Then
                                        Call Draw_Grh(Flechas_list(j).Grh, xb, yb, 0, 1, , X, Y, 1.3)
                                    ElseIf (Flechas_list(j).X = Flechas_list(j).xb) And (Flechas_list(j).Y < Flechas_list(j).yb) Then
                                        Call Draw_Grh(Flechas_list(j).Grh, xb, yb, 0, 1, , X, Y, 1.92)
                                    ElseIf (Flechas_list(j).X = Flechas_list(j).xb) And (Flechas_list(j).Y > Flechas_list(j).yb) Then
                                        Call Draw_Grh(Flechas_list(j).Grh, xb, yb, 0, 1, , X, Y, 200)
                                    Else
                                        Call Draw_Grh(Flechas_list(j).Grh, xb, yb, 0, 1, , X, Y, 0)
                                End If
                           
                            End If
                        End If
                    End If
                End If
               
            End If
        Next j
       
        For j = 0 To MaxFlecha
            If Flechas_list(j).Grh.GrhIndex Then
                If Abs(Flechas_list(j).X - Flechas_list(j).xb) <= 20 Then
                    If Abs(Flechas_list(j).Y - Flechas_list(j).yb) <= 20 Then
                        Flechas_list(j).EnUso = 0
                    End If
                End If
            End If
        Next j
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer
    For Y = minY To maxY
        ScreenX = minXOffset - Engine_Get_TileBuffer
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            If Map_InBounds(X, Y) Then
                With MapData(X, Y)
                    'Object Layer
                    If .ObjGrh.GrhIndex <> 0 Then
                         If EsArbol(.ObjGrh.GrhIndex) And Configuracion.Transparencias = 1 And (Y > (UserPos.Y - 2) And Y < (UserPos.Y + 7)) And (X > (UserPos.X - 4) And X < (UserPos.X + 4)) Then
                                  Call Draw_Grh(.ObjGrh, _
                                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y, , , , 120)
                         Else
                                Call Draw_Grh(.ObjGrh, _
                                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
                         End If
                    End If
                    
                    'Char layer
                    If .charindex <> 0 Then
                        Call Char_Render(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
                    End If
                    
                    'Layer 3
                    If .Graphic(3).GrhIndex <> 0 Then
                        If EsArbol(.Graphic(3).GrhIndex) And Configuracion.Transparencias = 1 And (Y > (UserPos.Y - 2) And Y < (UserPos.Y + 7)) And (X > (UserPos.X - 4) And X < (UserPos.X + 4)) Then
                            Call Draw_Grh(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y, , , , 120)
                        Else
                            Call Draw_Grh(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
                        End If
                    End If
                End With
            End If
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
     If MouseRendOK = True Then
            Draw_GrhIndex MouseItem, frmMain.MouseRenderX, frmMain.MouseRenderY
     End If
     
     
    If Configuracion.Particulas = 1 Then
        ScreenY = minYOffset - Engine_Get_TileBuffer
        For Y = minY To maxY
            ScreenX = minXOffset - Engine_Get_TileBuffer
            For X = minX To maxX
                    With MapData(X, Y)
                        '***********************************************
                            If .particle_group > 0 Then
                                Particle_Group_Render .particle_group, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY
                            End If
                    End With
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    
    If Configuracion.Desvanecimientos = 1 Then
        If bTecho Then
            If bTechoAB > 0 Then
                bTechoAB = bTechoAB - 0.4 * 15
                If bTechoAB < 10 Then bTechoAB = 0
            End If
        Else
            If bTechoAB < 255 Then
                bTechoAB = bTechoAB + 0.4 * 15
                If bTechoAB > 245 Then bTechoAB = 255
            End If
        End If
    Else
        bTechoAB = 255
    End If
   
    If bTechoAB > 0 Then
    '<----- Layer 4 ----->
        ScreenY = minYOffset - Engine_Get_TileBuffer
        For Y = minY To maxY
            ScreenX = minXOffset - Engine_Get_TileBuffer
            For X = minX To maxX
                If Map_InBounds(X, Y) Then
                    'Layer 4 **********************************
                    If MapData(X, Y).Graphic(4).GrhIndex Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), _
                            ScreenX * 32 + PixelOffsetX, _
                            ScreenY * 32 + PixelOffsetY, _
                            1, 1, , X, Y, , , , bTechoAB)
                    End If
                    '**********************************
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    If Cuenta = True Then
        ConteoH = ConteoH + (GrhData(Conteo).pixelHeight / 96)
        ConteoW = ConteoW + (GrhData(Conteo).pixelWidth / 96)
        Call Draw_Conteo(Conteo, 270, 220, ConteoW, ConteoH, TransparenciaCont)
    End If
    
    If Tiempo > 1 Then
    
        ConteoH = ConteoH + (GrhData(Conteo).pixelHeight / 96)
        ConteoW = ConteoW + (GrhData(Conteo).pixelWidth / 96)
    
        Call Draw_Conteo(Conteo, 250, 330, ConteoW, ConteoH, TransparenciaCont)
        Tiempo = Tiempo - 1
    End If


    If TextDesv.Existe = True Then
        
      If TextDesv.LifeTime >= 292 Then
        If TextDesv.Sube > 0 Then _
        TextDesv.Sube = TextDesv.Sube - 1
        If TextDesv.Sube > 0 Then
        TextDesv.Y = TextDesv.Y - TextDesv.Sube / 1.2
        TextDesv.Desvanecimiento = TextDesv.Desvanecimiento + 12
        End If
      End If
            
                If (GetTickCount() - TextDesv.StartTime) >= TextDesv.LifeTime Then
                    TextDesv.Tiempito = True
                End If
                
                If TextDesv.Tiempito = True Then
                    If TextDesv.Desvanecimiento >= 10 Then TextDesv.Desvanecimiento = TextDesv.Desvanecimiento - 10
                    
                    If TextDesv.Desvanecimiento <= 9 Then
                        TextDesv.Existe = False
                    End If
                End If
                
    
                Call Texto.Engine_Text_Draw(TextDesv.X, TextDesv.Y + 2, TextDesv.text, TextDesv.color, TextDesv.Desvanecimiento, True)
    End If
    
        If AramSeconds > 0 Then
            Call Texto.Engine_Text_Draw(270, 250, "Resucitarás en " & AramSeconds & " segundos!", D3DColorXRGB(255, 255, 255), 170, DT_CENTER)
        End If

            'CHAT CONTACTOS
            If ChatEnUso(1) = True Then
                If MouseBarraChat(1) = True Then
                    Draw_FilledBox 103, 394, 82, 17, D3DColorARGB(90, 208, 106, 106), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(1) = 1 Then
                    Draw_FilledBox 103, 394, 82, 17, D3DColorARGB(90, 115, 100, 13), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(1) = 2 Then
                    Draw_FilledBox 103, 394, 82, 17, D3DColorARGB(90, 32, 21, 49), D3DColorARGB(170, 35, 28, 25)
                Else
                    Draw_FilledBox 103, 394, 82, 17, D3DColorARGB(90, 100, 100, 100), D3DColorARGB(170, 35, 28, 25)
                End If
                
                Call Texto.Engine_Text_Draw(143, 395, UCase$(NickContacto(1)), D3DColorARGB(255, 255, 255, 255), 200, DT_CENTER)
            End If
            
            If ChatEnUso(2) = True Then
                If MouseBarraChat(2) = True Then
                    Draw_FilledBox 188, 394, 82, 17, D3DColorARGB(90, 208, 106, 106), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(2) = 1 Then
                    Draw_FilledBox 188, 394, 82, 17, D3DColorARGB(90, 115, 100, 13), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(2) = 2 Then
                    Draw_FilledBox 188, 394, 82, 17, D3DColorARGB(90, 32, 21, 49), D3DColorARGB(170, 35, 28, 25)
                Else
                    Draw_FilledBox 188, 394, 82, 17, D3DColorARGB(90, 100, 100, 100), D3DColorARGB(170, 35, 28, 25)
                End If
                
                Call Texto.Engine_Text_Draw(228, 395, UCase$(NickContacto(2)), D3DColorARGB(255, 255, 255, 255), 200, DT_CENTER)
            End If
            
            If ChatEnUso(3) = True Then
                If MouseBarraChat(3) = True Then
                    Draw_FilledBox 273, 394, 82, 17, D3DColorARGB(90, 208, 106, 106), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(3) = 1 Then
                    Draw_FilledBox 273, 394, 82, 17, D3DColorARGB(90, 115, 100, 13), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(3) = 2 Then
                    Draw_FilledBox 273, 394, 82, 17, D3DColorARGB(90, 32, 21, 49), D3DColorARGB(170, 35, 28, 25)
                Else
                    Draw_FilledBox 273, 394, 82, 17, D3DColorARGB(90, 100, 100, 100), D3DColorARGB(170, 35, 28, 25)
                End If
                
                Call Texto.Engine_Text_Draw(313, 395, UCase$(NickContacto(3)), D3DColorARGB(255, 255, 255, 255), 200, DT_CENTER)
            End If
            
            If ChatEnUso(4) = True Then
                If MouseBarraChat(4) = True Then
                    Draw_FilledBox 358, 394, 82, 17, D3DColorARGB(90, 208, 106, 106), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(4) = 1 Then
                    Draw_FilledBox 358, 394, 82, 17, D3DColorARGB(90, 115, 100, 13), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(4) = 2 Then
                    Draw_FilledBox 358, 394, 82, 17, D3DColorARGB(90, 32, 21, 49), D3DColorARGB(170, 35, 28, 25)
                Else
                    Draw_FilledBox 358, 394, 82, 17, D3DColorARGB(90, 100, 100, 100), D3DColorARGB(170, 35, 28, 25)
                End If
                
                Call Texto.Engine_Text_Draw(398, 395, UCase$(NickContacto(4)), D3DColorARGB(255, 255, 255, 255), 200, DT_CENTER)
            End If
            
            If ChatEnUso(5) = True Then
                If MouseBarraChat(5) = True Then
                    Draw_FilledBox 443, 394, 82, 17, D3DColorARGB(90, 208, 106, 106), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(5) = 1 Then
                    Draw_FilledBox 443, 394, 82, 17, D3DColorARGB(90, 115, 100, 13), D3DColorARGB(170, 35, 28, 25)
                ElseIf VentanitaMostrar(5) = 2 Then
                    Draw_FilledBox 443, 394, 82, 17, D3DColorARGB(90, 32, 21, 49), D3DColorARGB(170, 35, 28, 25)
                Else
                    Draw_FilledBox 443, 394, 82, 17, D3DColorARGB(90, 100, 100, 100), D3DColorARGB(170, 35, 28, 25)
                End If
                
                Call Texto.Engine_Text_Draw(483, 395, UCase$(NickContacto(5)), D3DColorARGB(255, 255, 255, 255), 200, DT_CENTER)
            End If
    
End Sub
Public Sub DrawDonations()
    Static re As RECT
        re.left = 0
        re.top = 0
        re.bottom = 129
        re.Right = 129
   
DirectDevice.BeginScene
DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0

            If picDonation.Aura > 0 Then
                    If 1 = AurasPJ(picDonation.Aura).Giratoria Then
                            picDonation.Aura_Angle = picDonation.Aura_Angle + 0.004
                            If picDonation.Aura_Angle >= 180 Then picDonation.Aura_Angle = 0
                    End If
                       
                    'If AurasPJ(picDonation.Aura).offset > 0 Then
                    '    Call Draw_Aura(picDonation.AuraA, 47 + BodyData(picDonation.Body).HeadOffset.X, BodyData(picDonation.Body).HeadOffset.Y - AurasPJ(picDonation.Aura).offset, 48, 52, True, , , picDonation.Aura_Angle, picDonation.Aura)
                    'Else
                        Call Draw_Aura(picDonation.AuraA, 47 + BodyData(picDonation.Body).HeadOffset.X, 125 + BodyData(picDonation.Body).HeadOffset.Y, 1, 1, True, , , picDonation.Aura_Angle, picDonation.Aura)
                    'End If
            End If
        
        If picDonation.Body <> 0 Then
            Call Draw_Grh(BodyData(picDonation.Body).Walk(3), 47, 55, 1, 1)
            Call Draw_Grh(HeadData(picDonation.Head).Head(3), 47 + BodyData(picDonation.Body).HeadOffset.X, 55 + BodyData(picDonation.Body).HeadOffset.Y, 1, 1)
            
            If picDonation.Casco <> 0 Then
                Call Draw_Grh(CascoAnimData(picDonation.Casco).Head(3), 47 + BodyData(picDonation.Body).HeadOffset.X, 55 + BodyData(picDonation.Body).HeadOffset.Y, 1, 1)
            End If
            
            If picDonation.Weapon <> 0 Then
                Call Draw_Grh(WeaponAnimData(picDonation.Weapon).WeaponWalk(3), 47 + BodyData(picDonation.Body).HeadOffset.X, 90 + BodyData(picDonation.Body).HeadOffset.Y, 1, 1)
            End If
            
            If picDonation.Shield <> 0 Then
                Call Draw_Grh(ShieldAnimData(picDonation.Shield).ShieldWalk(3), 49 + BodyData(picDonation.Body).HeadOffset.X, 90 + BodyData(picDonation.Body).HeadOffset.Y, 1, 1)
            End If
            
            Call Texto.Engine_Text_Draw(62, 84, frmMain.Label8.Caption, D3DColorARGB(255, 255, 255, 0), 255, DT_CENTER)
        Else
            Call Draw_GrhIndex(picDonation.GrhIndex, 47, 55)
        End If

DirectDevice.EndScene
DirectDevice.Present re, ByVal 0, frmMercadoTS.picPack.hWnd, ByVal 0

End Sub
Public Sub ConectarRenderizado()

Static re As RECT
    re.left = 0
    re.top = 0
    re.bottom = 600
    re.Right = 800
   
DirectDevice.BeginScene
DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0


If RenderConnect = True Then
        
        Dim X As Long
        Dim Y As Long
        
    Call Draw_GrhIndex(27807, 0, 0)
    

'particulas
   For Y = 1 To 19
       For X = 1 To 25
            With MapData(37 + X, 23 + Y)
                '***********************************************
                If .particle_group > 0 Then
                    Particle_Group_Render .particle_group, (X - 1) * 32, (Y - 1) * 32
                End If
            End With
       Next X
   Next Y

'Boton Ingresar
If ButtonLogin = "Normal" Then Call Draw_GrhIndex(27809, 326, 339)
If ButtonLogin = "Iluminado" Then Call Draw_GrhIndex(27810, 326, 339)
If ButtonLogin = "Apretado" Then Call Draw_GrhIndex(27811, 326, 339)

'Boton crear cuenta
If ButtonCC = "Normal" Then Call Draw_GrhIndex(27816, 227, 503)
If ButtonCC = "Iluminado" Then Call Draw_GrhIndex(27817, 227, 503)
If ButtonCC = "Apretado" Then Call Draw_GrhIndex(27818, 227, 503)

'Boton recuperar cuenta
If ButtonRC = "Normal" Then Call Draw_GrhIndex(27819, 350, 503)
If ButtonRC = "Iluminado" Then Call Draw_GrhIndex(27820, 350, 503)
If ButtonRC = "Apretado" Then Call Draw_GrhIndex(27821, 350, 503)

'Boton VerWeb
If ButtonVW = "Normal" Then Call Draw_GrhIndex(27828, 473, 503)
If ButtonVW = "Iluminado" Then Call Draw_GrhIndex(27829, 473, 503)
If ButtonVW = "Apretado" Then Call Draw_GrhIndex(27830, 473, 503)

If ClickeoTextCuenta = True Then
    If (BarritaTextConnect / 2) = Int((BarritaTextConnect / 2)) Then
        Draw_FilledBox 415 + (Len(TextBoxCuenta) * 3.4), 232, 1, 11, D3DColorARGB(100, 255, 255, 0), D3DColorARGB(100, 255, 255, 0)
    End If
ElseIf ClickeoTextPassw = True Then
    If (BarritaTextConnect / 2) = Int((BarritaTextConnect / 2)) Then
        Draw_FilledBox 415 + (Len(TextBoxPassw) * 3.4), 303, 1, 11, D3DColorARGB(100, 255, 255, 0), D3DColorARGB(100, 255, 255, 0)
    End If
End If

If ClickeoTextCuenta = True Then
Text_Render font_list(1), TextBoxCuenta, 232, 330, 169, 21, D3DColorARGB(200, 164, 161, 69), DT_CENTER, True
Else
Text_Render font_list(1), TextBoxCuenta, 232, 330, 169, 21, D3DColorARGB(200, 164, 161, 160), DT_CENTER, True
End If

If ClickeoTextPassw = True Then
Text_Render font_list(1), TextBoxPasswR, 303, 330, 169, 21, D3DColorARGB(200, 164, 161, 69), DT_CENTER, True
Else
Text_Render font_list(1), TextBoxPasswR, 303, 330, 169, 21, D3DColorARGB(200, 164, 161, 160), DT_CENTER, True
End If

ElseIf RenderAccount = True Then

    Call Draw_GrhIndex(27808, 0, 0)
        
'particulas
   For Y = 1 To 19
       For X = 1 To 25
            With MapData(37 + X, 40 + Y)
                '***********************************************
                If .particle_group > 0 Then
               
                Particle_Group_Render .particle_group, (X - 1) * 32, (Y - 1) * 32
               
                End If
            End With
       Next X
   Next Y
   
   'Call Draw_GrhIndex(27807, 75, 15)

Dim i As Long

    For i = 0 To 8
      If CantidadDePersonajes >= i Then
        If CargarPJ(i).Existe = True Then
        
            If CargarPJ(i).Muerto = 1 Then
                CargarPJ(i).Body = 8
                CargarPJ(i).Head = 500
            End If
            
        'Cargamos las coordenadas de dibujo
        If i = 0 Then 'PJ 1
         PersonajeX = 388
         BodyHeadY = 300
         ArmaEscudoY = 336 '36+
         NickX = 403 '15+
         NickY = 337 '37
        ElseIf i = 1 Then 'PJ 2
         PersonajeX = 388
         BodyHeadY = 467
         ArmaEscudoY = 503
         NickX = 403
         NickY = 504
        ElseIf i = 2 Then 'PJ 3
         PersonajeX = 568
         BodyHeadY = 432
         ArmaEscudoY = 469
         NickX = 583
         NickY = 470
        ElseIf i = 3 Then 'PJ 4
         PersonajeX = 648
         BodyHeadY = 300
         ArmaEscudoY = 336
         NickX = 663
         NickY = 337
        ElseIf i = 4 Then 'PJ 5
         PersonajeX = 568
         BodyHeadY = 169
         ArmaEscudoY = 205
         NickX = 583
         NickY = 206
        ElseIf i = 5 Then 'PJ 6
         PersonajeX = 388
         BodyHeadY = 135
         ArmaEscudoY = 171
         NickX = 403
         NickY = 172
        ElseIf i = 6 Then 'PJ 7
         PersonajeX = 210
         BodyHeadY = 169
         ArmaEscudoY = 205
         NickX = 225
         NickY = 206
        ElseIf i = 7 Then 'PJ 8
         PersonajeX = 128
         BodyHeadY = 300
         ArmaEscudoY = 336
         NickX = 143
         NickY = 337
        ElseIf i = 8 Then 'PJ 9
         PersonajeX = 210
         BodyHeadY = 432
         ArmaEscudoY = 469
         NickX = 225
         NickY = 470
        End If
            
            'Dibujamos el personaje
            Call Draw_Grh(BodyData(CargarPJ(i).Body).Walk(3), PersonajeX, BodyHeadY, 1, 1)
            Call Draw_Grh(HeadData(CargarPJ(i).Head).Head(3), PersonajeX + BodyData(CargarPJ(i).Body).HeadOffset.X, BodyHeadY + BodyData(CargarPJ(i).Body).HeadOffset.Y, 1, 1)
            
            If CargarPJ(i).Casco <> 0 Then
                Call Draw_Grh(CascoAnimData(CargarPJ(i).Casco).Head(3), PersonajeX + BodyData(CargarPJ(i).Body).HeadOffset.X + 1, BodyHeadY + BodyData(CargarPJ(i).Body).HeadOffset.Y, 1, 1)
            End If
            
            If CargarPJ(i).Weapon <> 0 Then
                Call Draw_Grh(WeaponAnimData(CargarPJ(i).Weapon).WeaponWalk(3), PersonajeX + BodyData(CargarPJ(i).Body).HeadOffset.X, ArmaEscudoY + BodyData(CargarPJ(i).Body).HeadOffset.Y, 1, 1)
            End If
            
            If CargarPJ(i).Shield <> 0 Then
                Call Draw_Grh(ShieldAnimData(CargarPJ(i).Shield).ShieldWalk(3), (PersonajeX + 2) + BodyData(CargarPJ(i).Body).HeadOffset.X, ArmaEscudoY + BodyData(CargarPJ(i).Body).HeadOffset.Y, 1, 1)
            End If
            
            'Dibujamos el nick
            Call Texto.Engine_Text_Draw(NickX, NickY, CargarPJ(i).Nombre, D3DColorARGB(255, 255, 255, 255), 255, DT_CENTER)

            Call Texto.Engine_Text_Draw(NickX, (NickY + 13), "Nivel: " & CargarPJ(i).Level, D3DColorARGB(255, 255, 255, 255), 255, DT_CENTER)
            
        'Al pasar el mouse por arriba
         If CrearAura(i) = True Then
            Aurix_Angle = Aurix_Angle + 0.025
            If Aurix_Angle >= 180 Then Aurix_Angle = 0
            
            Call Draw_Aura(AurixPJ, (NickX - 15), (NickY - 5), 1, 1, True, 1, 1, Aurix_Angle, 43)
        End If
            
        If MostrarTodo(i) = True Then
            Aurix_Angle = Aurix_Angle + 0.025
            If Aurix_Angle >= 180 Then Aurix_Angle = 0
            
            Call Draw_Aura(AurixPJ, (NickX - 15), (NickY - 5), 1, 1, True, 1, 1, Aurix_Angle, 8)
            
            Call Texto.Engine_Text_Draw(NickX, NickY, CargarPJ(i).Nombre, D3DColorARGB(255, 255, 255, 0), 255, DT_CENTER)
            
            Call Texto.Engine_Text_Draw(NickX, (NickY + 13), "Nivel: " & CargarPJ(i).Level, D3DColorARGB(255, 255, 255, 0), 255, DT_CENTER)
            
            If ButtonEntrarPJ = "Normal" Then Call Draw_GrhIndex(27836, EntrarX, EntrarY)
            If ButtonEntrarPJ = "Iluminado" Then Call Draw_GrhIndex(27834, EntrarX, EntrarY)
            If ButtonEntrarPJ = "Apretado" Then Call Draw_GrhIndex(27835, EntrarX, EntrarY)
            
            If ButtonBorrarPJ = "Normal" Then Call Draw_GrhIndex(27848, EntrarX + 68, EntrarY)
            If ButtonBorrarPJ = "Iluminado" Then Call Draw_GrhIndex(27846, EntrarX + 68, EntrarY)
            If ButtonBorrarPJ = "Apretado" Then Call Draw_GrhIndex(27847, EntrarX + 68, EntrarY)
            
        End If
            
      End If
     End If
    Next i


    If ButtonCP = "Normal" Then Call Draw_GrhIndex(27839, 291, 381)
    If ButtonCP = "Iluminado" Then Call Draw_GrhIndex(27837, 291, 381)
    If ButtonCP = "Apretado" Then Call Draw_GrhIndex(27838, 291, 381)

    If ButtonCPass = "Normal" Then Call Draw_GrhIndex(27842, 618, 513)
    If ButtonCPass = "Iluminado" Then Call Draw_GrhIndex(27840, 618, 513)
    If ButtonCPass = "Apretado" Then Call Draw_GrhIndex(27841, 618, 513)

    If ButtonSalir = "Normal" Then Call Draw_GrhIndex(27833, 681, 90)
    If ButtonSalir = "Iluminado" Then Call Draw_GrhIndex(27831, 681, 90)
    If ButtonSalir = "Apretado" Then Call Draw_GrhIndex(27832, 681, 90)

End If

    If GetTickCount >= (fpsLastCheck + 1000) Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        fpsLastCheck = GetTickCount
        lFrameTimer = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If

    lframelimiter = GetTickCount
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * EngineBaseSpeed
    particletimer = timerElapsedTime * 0.05
    
DirectDevice.EndScene
DirectDevice.Present re, ByVal 0, frmConnect.Picture1.hWnd, ByVal 0
    
'Text_Render font_list(1), "FPS: " & FPS, 400, 300, 169, 21, D3DColorARGB(200, 255, 255, 0), DT_CENTER, True

End Sub
Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Function Geometry_Create_TLVertex2(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
'mz
Geometry_Create_TLVertex2.X = X
Geometry_Create_TLVertex2.Y = Y
Geometry_Create_TLVertex2.Z = Z
Geometry_Create_TLVertex2.rhw = rhw
Geometry_Create_TLVertex2.color = color
Geometry_Create_TLVertex2.Specular = Specular
Geometry_Create_TLVertex2.tu1 = tu1
Geometry_Create_TLVertex2.tv1 = tv1
Geometry_Create_TLVertex2.tu2 = tu2
Geometry_Create_TLVertex2.tv2 = tv2
End Function
Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, _
                                ByRef dest As RECT, _
                                ByRef src As RECT, _
                                ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, _
                                Optional ByRef Textures_Height As Long, _
                                Optional ByVal angle As Single, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False)

    '**************************************************************
    'Author: Aaron Perkins
    'Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/17/2002
    '
    ' * v1      * v3
    ' |\        |
    ' |  \      |
    ' |    \    |
    ' |      \  |
    ' |        \|
    ' * v0      * v2
    '**************************************************************
    Dim x_center    As Single

    Dim y_center    As Single

    Dim radius      As Single

    Dim x_Cor       As Single

    Dim y_Cor       As Single

    Dim left_point  As Single

    Dim right_point As Single

    Dim Temp        As Single
    
    Dim auxr As RECT
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left) / 2
        y_center = dest.top + (dest.bottom - dest.top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        Temp = (dest.Right - x_center) / radius
        right_point = Atn(Temp / Sqr(-Temp * Temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    auxr = src
 
    If angle < 0 Then
        src.left = auxr.Right
        src.Right = auxr.left
    End If
    
    If Invert_y Then
        src.top = auxr.bottom
        src.bottom = auxr.top
    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / Textures_Width, src.top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
   
    'If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = base_light
        If (light_value(1) = 0) Then light_value(1) = base_light
        If (light_value(2) = 0) Then light_value(2) = base_light
        If (light_value(3) = 0) Then light_value(3) = base_light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    DirectDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub
Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, _
                                      ByVal dest_x As Integer, _
                                      ByVal dest_y As Integer, _
                                      ByVal src_width As Integer, _
                                      ByVal src_height As Integer, _
                                      ByRef rgb_list() As Long, _
                                      ByVal src_x As Integer, _
                                      ByVal src_y As Integer, _
                                      Optional ByVal alpha_blend As Boolean, _
                                      Optional ByVal angle As Single, Optional ByVal Shadow As Byte = 0, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False, Optional alphaa As Byte = 255)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    'If Not char_current_blind Then
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
        
    If alphaa <> 255 Then
        Dim Aux As D3DCOLORVALUE
        ARGBtoD3DCOLORVALUE RGB(day_r_old, day_g_old, day_b_old), Aux
        Dim i As Long
        For i = 0 To 3
        light_value(i) = D3DColorARGB(alphaa, Aux.r, Aux.g, Aux.b)
        Next i
    End If
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .left = dest_x
        .Right = dest_x + src_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle, Invert_x, Invert_y
    
    'Set Textures
    DirectDevice.SetTexture 0, d3dTextures.Texture
    
   If Shadow Then
        temp_verts(1).X = temp_verts(1).X + src_width / 2
        temp_verts(1).Y = temp_verts(1).Y - src_height / 2
       
        temp_verts(3).X = temp_verts(3).X + src_width
        temp_verts(3).Y = temp_verts(3).Y - src_width
    End If
    
    If alpha_blend Then
        'Set Rendering for alphablending
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Texture
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub
Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
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
End Sub
Private Sub Char_Render(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
On Error Resume Next
    Dim moved As Boolean
    Dim Pos As Integer
    Dim Line As String
    Dim color As Long
   
With charlist(charindex)
    Dim i As Long
    
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
Select Case .priv
    Case 0
            If .color = 40 Or .color = 41 Or .color = 42 Then
                Call DarColorCambiante(charindex)
                color = D3DColorXRGB(.ProximoR, .ProximoG, .ProximoB)
            ElseIf .color = 49 Then
                color = D3DColorXRGB(0, 0, 255)
            ElseIf .color = 50 Then
                color = D3DColorXRGB(255, 0, 0)
            Else
                color = EsStatusCOLOR(.EsStatus)
            End If
    Case Else 'el resto
        color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
End Select

            If Configuracion.Contador = 1 Then
                If UserParalizado = True And CartelInvisibilidad > 0 Then
                    Call Draw_GrhIndex(23610, 10, 5)
                    Draw_FilledBox 13, 40, 25, 6, D3DColorARGB(150, 125, 125, 125), D3DColorARGB(150, 125, 125, 125)
                    Draw_FilledBox 13, 40, TiempoParalizado, 6, D3DColorARGB(200, 255, 255, 0), D3DColorARGB(1, 200, 200, 200)
                    Call Draw_GrhIndex(23611, 45, 5)
                    Draw_FilledBox 48, 40, 25, 6, D3DColorARGB(150, 125, 125, 125), D3DColorARGB(150, 125, 125, 125)
                    Draw_FilledBox 48, 40, CartelInvisibilidad, 6, D3DColorARGB(200, 255, 255, 0), D3DColorARGB(1, 200, 200, 200)
                ElseIf CartelInvisibilidad > 0 Then
                    Call Draw_GrhIndex(23611, 10, 5)
                    Draw_FilledBox 13, 40, 25, 6, D3DColorARGB(150, 125, 125, 125), D3DColorARGB(150, 125, 125, 125)
                    Draw_FilledBox 13, 40, CartelInvisibilidad, 6, D3DColorARGB(200, 255, 255, 0), D3DColorARGB(1, 200, 200, 200)
                ElseIf UserParalizado Then
                    Call Draw_GrhIndex(23610, 10, 5)
                    Draw_FilledBox 13, 40, 25, 6, D3DColorARGB(150, 125, 125, 125), D3DColorARGB(150, 125, 125, 125)
                    Draw_FilledBox 13, 40, TiempoParalizado, 6, D3DColorARGB(200, 255, 255, 0), D3DColorARGB(1, 200, 200, 200)
                End If
            End If
            
    Movement_Speed = 0.5
 
If Not .invisible Then
        'REFLEJOS
        If HayAgua(.Pos.X, .Pos.Y + 1) And Configuracion.ReflejosAgua = 1 Then
            If .Body.Walk(.Heading).GrhIndex Then _
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY + 40, 1, 1, , X, Y, 0, False, True, 100)
                        
            If UserEstado = 0 Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX - .Body.HeadOffset.X, PixelOffsetY - .Body.HeadOffset.Y + 11, 1, 0, , X, Y, 0, False, True, 100)
            Else
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX - .Body.HeadOffset.X, PixelOffsetY - .Body.HeadOffset.Y + 24, 1, 0, , X, Y, 0, False, True, 100)
            End If
                        
            If .Casco.Head(.Heading).GrhIndex Then _
                Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX - .Body.HeadOffset.X + 1, PixelOffsetY - .Body.HeadOffset.Y + 11, 1, 1, , X, Y, , False, True, 100)
                        
            If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + 40, 1, 1, , X, Y, 0, False, True, 100)
                            
            If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY + 40, 1, 1, , X, Y, 0, False, True, 100)
        End If
        
        'AURAS
        If Configuracion.Auras = 1 Then
            If .NPCNumber = 0 Then
                If .Aura_IndexA > 0 And .Navegando = 0 And .Montando = 0 Then
                    If 1 = AurasPJ(.Aura_IndexA).Giratoria Then
                    .Aura_AngleA = .Aura_AngleA + 0.004
                    If .Aura_AngleA >= 180 Then .Aura_AngleA = 0
                    End If
                       
                    If AurasPJ(.Aura_IndexA).Offset > 0 Then
                        Call Draw_Aura(.AuraA, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72 - AurasPJ(.Aura_IndexA).Offset, 48, 52, True, X, Y, .Aura_AngleA, .Aura_IndexA, charindex)
                    Else
                        Call Draw_Aura(.AuraA, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72, 48, 52, True, X, Y, .Aura_AngleA, .Aura_IndexA, charindex)
                    End If
                    
                End If
                
                If .Aura_IndexW > 0 And .Navegando = 0 And .Montando = 0 Then
                    If 1 = AurasPJ(.Aura_IndexW).Giratoria Then
                    .Aura_AngleW = .Aura_AngleW + 0.004
                    If .Aura_AngleW >= 180 Then .Aura_AngleW = 0
                    End If
                       
                    If AurasPJ(.Aura_IndexW).Offset > 0 Then
                        Call Draw_Aura(.AuraW, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72 - AurasPJ(.Aura_IndexW).Offset, 48, 52, True, X, Y, .Aura_AngleW, .Aura_IndexW)
                    Else
                        Call Draw_Aura(.AuraW, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72, 48, 52, True, X, Y, .Aura_AngleW, .Aura_IndexW)
                    End If
                End If
                
                If .Aura_IndexE > 0 And .Navegando = 0 And .Montando = 0 Then
                    If 1 = AurasPJ(.Aura_IndexE).Giratoria Then
                    .Aura_AngleE = .Aura_AngleE + 0.004
                    If .Aura_AngleE >= 180 Then .Aura_AngleE = 0
                    End If
                       
                    If AurasPJ(.Aura_IndexE).Offset > 0 Then
                        Call Draw_Aura(.AuraE, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72 - AurasPJ(.Aura_IndexE).Offset, 48, 52, True, X, Y, .Aura_AngleE, .Aura_IndexE)
                    Else
                        Call Draw_Aura(.AuraE, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72, 48, 52, True, X, Y, .Aura_AngleE, .Aura_IndexE)
                    End If
                End If
                
                If .Aura_IndexR > 0 And .Navegando = 0 And .Montando = 0 Then
                    If 1 = AurasPJ(.Aura_IndexR).Giratoria Then
                    .Aura_AngleR = .Aura_AngleR + 0.004
                    If .Aura_AngleR >= 180 Then .Aura_AngleR = 0
                    End If
                        
                    If AurasPJ(.Aura_IndexR).Offset > 0 Then
                        Call Draw_Aura(.AuraR, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72 - AurasPJ(.Aura_IndexR).Offset, 48, 52, True, X, Y, .Aura_AngleR, .Aura_IndexR)
                    Else
                        Call Draw_Aura(.AuraR, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72, 48, 52, True, X, Y, .Aura_AngleR, .Aura_IndexR)
                    End If
                End If
                
                If .Aura_IndexC > 0 And .Navegando = 0 And .Montando = 0 Then
                    If 1 = AurasPJ(.Aura_IndexC).Giratoria Then
                    .Aura_AngleC = .Aura_AngleC + 0.004
                    If .Aura_AngleC >= 180 Then .Aura_AngleC = 0
                    End If
                       
                    If AurasPJ(.Aura_IndexC).Offset > 0 Then
                        Call Draw_Aura(.AuraR, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72 - AurasPJ(.Aura_IndexC).Offset, 48, 52, True, X, Y, .Aura_AngleC, .Aura_IndexC)
                    Else
                        Call Draw_Aura(.AuraR, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72, 48, 52, True, X, Y, .Aura_AngleC, .Aura_IndexC)
                    End If
                End If
             End If
        End If
            '//AURAS
            
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then
            
            If Configuracion.Sombras = 1 Then Call Draw_Grh_Sombra(.Body.Walk(.Heading), PixelOffsetX - 6, PixelOffsetY, 1, 1, , X, Y, , 1)
                
                
              'Transparencia, muerto¿?
              If .Muerto = True And (Configuracion.Transparencias = 1) Then
                    If .Llegoalatransp = False And .TransparenciaBody < 100 Then
                        .TransparenciaBody = .TransparenciaBody + 1
                        If .TransparenciaBody = 100 Then .Llegoalatransp = True
                    ElseIf .Llegoalatransp = True Then
                        .TransparenciaBody = .TransparenciaBody - 1
                        If .TransparenciaBody = 0 Then .Llegoalatransp = False
                    End If
            
                    Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y, , , , .TransparenciaBody + 45)
              Else
              
                'de espaldas, primero el escudo y arriba el body.
                If .Heading = 1 Or .Heading = 2 Then
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y)
                    
                    Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                    
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y)
                Else
                    Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                    
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y)
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y)
                End If
                
              End If
            End If
                    
        'SOMBRAS
        If .Head.Head(.Heading).GrhIndex Then
            
              If Configuracion.Sombras = 1 Then Call Draw_Grh_Sombra(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.Y - 8, 1, 0, , X, Y, , 1)

            If .Heading = 1 Then
              If .Muerto = True And (Configuracion.Transparencias = 1) Then
                Call Draw_Grh(.Head.Head(.Heading), (PixelOffsetX + .Body.HeadOffset.X) - 1, PixelOffsetY + .Body.HeadOffset.Y + 1, 1, 0, , X, Y, , , , .TransparenciaBody + 45)
              Else
                Call Draw_Grh(.Head.Head(.Heading), (PixelOffsetX + .Body.HeadOffset.X) - 1, PixelOffsetY + .Body.HeadOffset.Y + 1, 1, 0, , X, Y)
              End If
            Else
               If .Muerto = True And (Configuracion.Transparencias = 1) Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 1, 1, 0, , X, Y, , , , .TransparenciaBody + 45)
              Else
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 1, 1, 0, , X, Y)
              End If
            End If
                                
            If .Casco.Head(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.Y, 1, 1, , X, Y)
            End If
                            
                    'Draw name over head
                    If Nombres And Configuracion.Nombres = 1 Then
                        If Len(.Nombre) > 0 Then
                            Pos = InStr(.Nombre, "<")
                            If Pos = 0 Then Pos = Len(.Nombre) + 2

                        'Nick
                        Line = left$(.Nombre, Pos - 2)
                        
                      If .priv <> 0 Then
                        Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, RangoPRIV(.priv), color, 170, DT_CENTER)
                      End If
                                                
                      If .priv = 0 Then
                        If HayAgua(.Pos.X, .Pos.Y + 1) Or .Muerto Then
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 30, Line, color, 80, DT_CENTER)
                        Else
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 30, Line, color, , DT_CENTER)
                        End If
                      Else
                        Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 30, Line, color, 170, DT_CENTER)
                      End If
                        
                        Dim CantLetras As String
                        
                        'ESTRELLA RANKING
                        CantLetras = Len(Line)
                        If .TieneRanking = True Then
                            If .PosRanking = 1 Then
                                Call Draw_GrhIndex(30450, PixelOffsetX - (CantLetras * 3), PixelOffsetY + 30)
                            ElseIf .PosRanking = 2 Then
                                Call Draw_GrhIndex(30451, PixelOffsetX - (CantLetras * 3), PixelOffsetY + 30)
                            ElseIf .PosRanking = 3 Then
                                Call Draw_GrhIndex(30452, PixelOffsetX - (CantLetras * 3), PixelOffsetY + 30)
                            End If
                        End If
                        
                        If .color = 41 Then
                            Call Draw_GrhIndex(31135, PixelOffsetX - (CantLetras * 4), PixelOffsetY + 28)
                        End If
                        
                                                
                        'AGREGAMOS EL NICK AL FRM DENUNCIAS
                        Dim aDenuncias As Integer, aEncontro As Boolean
                        aEncontro = False
                        For aDenuncias = 0 To frmGM.lstVistos.ListCount - 1
                            If frmGM.lstVistos.List(aDenuncias) = Line Then
                                aEncontro = True
                            End If
                        Next aDenuncias
                        
                        If aEncontro = False And Line <> "" And Line <> " " Then
                          frmGM.lstVistos.AddItem Line
                        End If
                           
                                    'Clan
                                    If .priv = 0 Then
                                        Line = mid$(.Nombre, Pos)
                                                
                                        If HayAgua(.Pos.X, .Pos.Y + 1) Or .Muerto Then
                                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, Line, color, 80, DT_CENTER)
                                        Else
                                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, Line, color, , DT_CENTER)
                                        End If
                                    End If
                    End If
                End If
            Else
                Pos = InStr(.Nombre, "<")
                If Pos = 0 Then Pos = Len(.Nombre) + 2
                Line = left$(.Nombre, Pos - 2)
                
                If UCase$(frmMain.Label8.Caption) = UCase$(Line) Or UserPrivilegios >= 1 Then
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 30, Line, color, 80, DT_CENTER)
                            
                    If .priv = 0 Then
                        Line = mid$(.Nombre, Pos)
                        Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, Line, color, 80, DT_CENTER)
                    End If
                    
                    If .Llegoalatransp = False And .TransparenciaBody < 100 Then
                        .TransparenciaBody = .TransparenciaBody + 1
                        If .TransparenciaBody = 100 Then .Llegoalatransp = True
                    ElseIf .Llegoalatransp = True Then
                        .TransparenciaBody = .TransparenciaBody - 1
                        If .TransparenciaBody = 0 Then .Llegoalatransp = False
                    End If
                    
                    'Cuerpo/arma/escudo
                    If .Heading = 1 Or .Heading = 2 Then
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                            Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y, , , , .TransparenciaBody)
                        
                        Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y, , , , .TransparenciaBody)
                        
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y, , , , .TransparenciaBody)
                    Else
                        Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y, , , , .TransparenciaBody)
                        
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y, , , , .TransparenciaBody)
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                            Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 38, 1, 1, , X, Y, , , , .TransparenciaBody)
                    End If
                    
                    'Head
                    If .Heading = 1 Then
                        Call Draw_Grh(.Head.Head(.Heading), (PixelOffsetX + .Body.HeadOffset.X) - 1, PixelOffsetY + .Body.HeadOffset.Y + 1, 1, 0, , X, Y, , , , .TransparenciaBody)
                    Else
                        Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 1, 1, 0, , X, Y, , , , .TransparenciaBody)
                    End If
                    
                    'Casco
                    If .Casco.Head(.Heading).GrhIndex Then _
                            Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.Y, 1, 1, , X, Y, , , , .TransparenciaBody)
                    
                    
                    If .priv <> 0 Then
                        Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, RangoPRIV(.priv), color, 80, DT_CENTER)
                    End If
                End If
                
            End If 'If not visible
           
        Movement_Speed = 1
           
        ''Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X - 168, PixelOffsetY + .Body.HeadOffset.Y, charindex)

       'particulas para meditar
       If Configuracion.Particulas = 1 Then
            If .particle_count > 0 Then
                For i = 1 To .particle_count
                    If .particle_group(i) > 0 Then
                        Particle_Group_Render .particle_group(i), PixelOffsetX, PixelOffsetY
                    End If
                Next i
            End If
       End If
        
            If .NPCAura > 0 Then
                If 1 = AurasPJ(.NPCAura).Giratoria Then
                    .NPCAuraAngle = .NPCAuraAngle + 0.004
                    If .NPCAuraAngle >= 180 Then .NPCAuraAngle = 0
                End If
                   
                If AurasPJ(.NPCAura).Offset > 0 Then
                    Call Draw_Aura(.NPCAuraG, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72 - AurasPJ(.NPCAura).Offset, 48, 52, True, X, Y, .NPCAuraAngle, .NPCAura)
                Else
                    Call Draw_Aura(.NPCAuraG, PixelOffsetX + .Body.HeadOffset.X, .Body.HeadOffset.Y + PixelOffsetY + 72, 48, 52, True, X, Y, .NPCAuraAngle, .NPCAura)
                End If
            End If
        
        If Configuracion.VerEmoticons = 1 Then
            If .EmoticonIndex <> 0 Then
                Call Draw_Grh(.Emoticon, PixelOffsetX + FxData(.EmoticonIndex).OffsetX, PixelOffsetY + FxData(.EmoticonIndex).OffsetY, 1, 1, False, , , , , , 150)
            
                If .EmoticonLoops > 0 Then .EmoticonLoops = .EmoticonLoops - 1
                
                If .EmoticonLoops = 0 Then
                    .EmoticonIndex = 0
                    .EmoticonLoops = 0
                End If
            End If
        End If
        
        'Draw FX
        For i = 1 To 3
            If .FxIndex(i) <> 0 Then
                Call Draw_Grh(.Fx(i), PixelOffsetX + FxData(.FxIndex(i)).OffsetX, PixelOffsetY + FxData(.FxIndex(i)).OffsetY, 1, 1, False, , , , , , 150)
                    
                    If .Fx(i).Started = 0 Then _
                        .FxIndex(i) = 0
            End If
        Next i
    End With
    
End Sub

Public Sub Start()

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 3) As Integer
Dim loopc As Integer
    
DoEvents
vertList(0).X = 0: vertList(0).Y = 0
vertList(1).X = 800: vertList(1).Y = 0
vertList(2).X = 0: vertList(2).Y = 600
vertList(3).X = 800: vertList(3).Y = 600
vertList(0).rhw = 1
vertList(1).rhw = 1
vertList(2).rhw = 1
vertList(3).rhw = 1
vertList(0).color = D3DColorXRGB(255, 255, 255)
vertList(1).color = D3DColorXRGB(255, 255, 255)
vertList(2).color = D3DColorXRGB(255, 255, 255)
vertList(3).color = D3DColorXRGB(255, 255, 255)
vertList(0).tu = 0
vertList(0).tv = 0
vertList(1).tu = 1
vertList(1).tv = 0
vertList(2).tu = 0
vertList(2).tv = 1
vertList(3).tu = 1
vertList(3).tv = 1
On Error Resume Next ' putos timers!!!!!

Do While prgRun
    
    If frmConnect.Visible Then
        engine.ConectarRenderizado
    End If

        'TODO : Porque el pausado de 20 ms???
            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmNuevoComercio.Visible And Not frmBancoObj.Visible Then
                CheckKeys
                lastTime = GetTickCount
            End If
            
        If Configuracion.FPS = 18 Then '18 FPS
            While (GetTickCount - lFrameTimer) \ 56 < FramesPerSecCounter
                Sleep 5
            Wend
        ElseIf Configuracion.FPS = 32 Then '32 FPS
            While (GetTickCount - lFrameTimer) \ 33 < FramesPerSecCounter
                Sleep 5
            Wend
        ElseIf Configuracion.FPS = 65 Then '65 FPS
            While (GetTickCount - lFrameTimer) \ 15 < FramesPerSecCounter
                Sleep 5
            Wend
        Else 'FPS LIBRES
            While (GetTickCount - lFrameTimer) < FramesPerSecCounter
                Sleep 5
            Wend
        End If
    
    If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
        engine.Render
        engine.DrawInv
    Else
        Sleep 10&
    End If
    DoEvents
    
            esttick = GetTickCount
            For loopc = 1 To UBound(timers)
                timers(loopc) = timers(loopc) + (esttick - ulttick)
                'Timer de trabajo
                If timers(1) >= tUs Then
                    timers(1) = 0
                    NoPuedeUsar = False
                End If
                'timer de attaque (77)
                If timers(2) >= tAt Then
                    timers(2) = 0
                    UserCanAttack = 1
                    UserPuedeRefrescar = True
                End If
                If timers(3) >= tMagia Then
                    timers(3) = 0
                    UserCanAttackMagia = 1
                    UserPuedeRefrescar = True
                End If
            Next loopc
            ulttick = GetTickCount
            
     DoEvents
Loop

    engine.Engine_Deinit

    EngineRun = False
    frmCargando.Show

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    Audio.MP3_Destroy
    
End
End Sub

Public Sub SetCharacterFx(ByVal charindex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************

    With charlist(charindex)
    
        If Fx = 0 Then
            Dim i As Long
            For i = 1 To 3
                .FxIndex(i) = 0
            Next i
            Exit Sub
        End If
            
    
        Dim slotLibre As Byte
        slotLibre = BuscarSlotLibreFx(charindex)
        .FxIndex(slotLibre) = Fx
       
        If .FxIndex(slotLibre) > 0 Then
            Call InitGrh(.Fx(slotLibre), FxData(Fx).Animacion)
       
            .Fx((slotLibre)).Loops = Loops
        End If
    End With
End Sub

Public Sub Char_Move_by_Head(ByVal charindex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(charindex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
    If LegalPos(nX, nY) Then
            MapData(nX, nY).charindex = charindex
            .Pos.X = nX
            .Pos.Y = nY
            MapData(X, Y).charindex = 0
            
            .MoveOffsetX = -1 * (32 * addx)
            .MoveOffsetY = -1 * (32 * addy)
            
            .Moving = 1
            .Heading = nHeading
            
            .scrollDirectionX = addx
            .scrollDirectionY = addy
        
        If UserEstado <> 1 Then Call DoPasosFx(charindex)
        
        'areas viejos
        If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
            Call EraseChar(charindex)
        End If
    End If
    End With
End Sub

Public Sub Char_Move_by_Pos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(charindex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).charindex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
    If LegalPos(nX, nY) Then
        
        MapData(nX, nY).charindex = charindex
        
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        Dim i As Long
        For i = 1 To 3
            If .FxIndex(i) = 4 Or .FxIndex(i) = 5 Or .FxIndex(i) = 6 Or .FxIndex(i) = 42 Or .FxIndex(i) = 43 Or .FxIndex(i) = 44 Or .FxIndex(i) = 45 Or .FxIndex(i) = 16 Or .FxIndex(i) = 103 Or .FxIndex(i) = 104 Or .FxIndex(i) = 105 Then    'If it's meditating, we remove the FX
                .FxIndex(i) = 0
                .Fx(i).Loops = 0
            End If
        Next i
    
    If Not EstaPCarea(charindex) Then
    Call Dialogos.RemoveDialog(charindex)
    'Call Hits.RemoveHit(CharIndex)
    End If
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)
    End If
  End If
    
    End With
End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean
    With charlist(charindex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function
Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, f
    DirectD3D8.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function
Public Function RemoveCharAparence(ByVal charindex As Integer, ByVal Muerto As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As E_Heading, _
                        ByVal tempint As Integer, ByVal tempweap As Integer, ByVal tempesc As Integer, _
                        ByVal Fx As Integer, ByVal Loops As Integer)
                        
            charindex = charindex
            
            If Head = 500 Then
                charlist(charindex).Muerto = True
            ElseIf Head = 501 Then
                charlist(charindex).Muerto = True
            ElseIf Head = 502 Then
                charlist(charindex).Muerto = True
            ElseIf Head = 511 Then
                charlist(charindex).Muerto = True
            ElseIf Head = 512 Then
                charlist(charindex).Muerto = True
            Else
                charlist(charindex).Muerto = False
            End If
                
            charlist(charindex).Body = BodyData(Body)
            charlist(charindex).Head = HeadData(Head)
            charlist(charindex).Heading = Heading
            
            tempint = tempint
            If tempint <> 0 Then charlist(charindex).Arma = WeaponAnimData(tempint)
            tempint = tempweap
            If tempint <> 0 Then charlist(charindex).Escudo = ShieldAnimData(tempint)
            tempint = tempesc
            If tempint <> 0 Then charlist(charindex).Casco = CascoAnimData(tempint)

End Function
Public Function ChangeBodyAparence(ByVal charindex As Integer, ByVal Body As Integer)
                        
            charindex = charindex
            charlist(charindex).Body = BodyData(Body)
            
End Function
Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)

    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    
    With box_rect
        .bottom = Y + Height - 1
        .left = X
        .Right = X + Width - 1
        .top = Y
    End With
    
    With Outline
        .bottom = Y + Height
        .left = X - 1
        .Right = X + Width
        .top = Y - 1
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub
Public Sub DrawInv()
    Static re As RECT
    re.left = 0
    re.top = 0
    re.bottom = 167
    re.Right = 232

    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    DirectDevice.BeginScene
        Inventario.DrawInventory
    DirectDevice.EndScene
    DirectDevice.Present re, ByVal 0, frmMain.picInv.hWnd, ByVal 0
End Sub
Public Sub GrhRenderToHdc(ByVal grh_index As Long, desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False, Optional color As String = &H0&)

    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    
    If grh_index <= 0 Then Exit Sub

    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If


        If Extract_File(graphics, App.Path & "\Data\GRAFICOS\", GrhData(grh_index).FileNum & ".png", App.Path & "\Data\GRAFICOS\") Then
            file_path = App.Path & "\Data\GRAFICOS\" & GrhData(grh_index).FileNum & ".png"
        End If
        
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
            
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
        If transparent = False Then
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        Else
            MaskDC = CreateCompatibleDC(desthDC)
            
            PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
            Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
            
            'Render tranparently
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
            
            Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
            DeleteDC MaskDC
        End If
        
        Call DeleteObject(SelectObject(hdcsrc, PrevObj))
        DeleteDC hdcsrc

Kill file_path

    Exit Sub
End Sub
Public Sub DrawGrhtoHdc(ByVal Grh As Integer, _
                        ByRef rec As RECT, _
                        ByRef pic As PictureBox, _
                        Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0, Optional ByVal r As Byte = 255, Optional ByVal g As Byte = 255, Optional ByVal b As Byte = 255)
 
        '*****************************************************************
       'Draws a Grh's portion to the given area of any Device Context
       '*****************************************************************
 
        Dim Light(3) As Long
        Light(0) = D3DColorXRGB(r, g, b)
        Light(1) = Light(0)
        Light(2) = Light(0)
        Light(3) = Light(0)
        
        If Grh = 0 Then pic.Picture = Nothing: Exit Sub
       
        'Clear the inventory window
       DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
       DirectDevice.BeginScene
        
            Device_Box_Textured_Render Grh, _
                0, 0, _
                GrhData(Grh).pixelWidth, GrhData(Grh).pixelHeight, _
                Light, _
                GrhData(Grh).sX, GrhData(Grh).sY
       
       DirectDevice.EndScene
       DirectDevice.Present rec, ByVal 0, pic.hWnd, ByVal 0
       
End Sub
Private Sub Draw_Grh_Sombra(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Shadow As Byte = 0)
 
    Dim CurrentGrhIndex As Integer
    If Grh.GrhIndex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
               
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
   
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
 
        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If
 
    Dim shadowRgb(3) As Long
    shadowRgb(0) = 1677721600
    shadowRgb(1) = 1677721600
    shadowRgb(2) = 1677721600
    shadowRgb(3) = 1677721600
 
    Device_Box_Textured_Render CurrentGrhIndex, _
        X + 5, Y + 8, _
        (GrhData(CurrentGrhIndex).pixelWidth / 1.4), (GrhData(CurrentGrhIndex).pixelHeight / 1.6), _
        shadowRgb(), _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        Alpha, angle, 1
       
End Sub

'----------------------------------------PARTICULAS---------------------------------
Private Sub Convert_Heading_to_Direction(ByVal Heading As Long, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addy As Long
    Dim addx As Long
   
    'Figure out which way to move
    Select Case Heading
        Case 1
            addy = -1
   
        Case 2
            addx = 1
   
        Case 3
            addy = 1
           
        Case 4
            addx = -1
   
    End Select
   
    direction_x = direction_x + addx
    direction_y = direction_y + addy
End Sub
 
 
Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
   
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
    End If
End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
   
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
   
    Particle_Group_Remove_All = True
End Function
 
Public Function Particle_Group_Find(ByVal ID As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until particle_group_list(loopc).ID = ID
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new particle effect
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    MapData(map_x, map_y).particle_group = particle_group_index
End Sub
 
Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional ByVal Bind_X As Integer = -1, Optional ByVal Bind_Y As Integer = -1)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/24/2003
'
'**************************************************************
 
    If no_move = False Then
                If temp_particle.alive_counter = 0 Then
                    InitGrh temp_particle.Grh, grh_index, alpha_blend
                    temp_particle.X = RandomNumber(X1, X2)
                    temp_particle.Y = RandomNumber(Y1, Y2)
                    temp_particle.vector_x = RandomNumber(vecx1, vecx2)
                    temp_particle.vector_y = RandomNumber(vecy1, vecy2)
                    temp_particle.angle = angle
                    temp_particle.alive_counter = RandomNumber(life1, life2)
                    temp_particle.friction = fric
                Else
                    'Continue old particle
                    'Do gravity
                    If gravity = True Then
                        temp_particle.vector_y = temp_particle.vector_y + grav_strength
                        If temp_particle.Y > 0 Then
                            'bounce
                            temp_particle.vector_y = bounce_strength
                        End If
                    End If
                    'Do rotation
                   If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                               
                    If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
                End If
 
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
   
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
   
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.GrhIndex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
 
    'Draw it
    If temp_particle.Grh.GrhIndex Then
       
        ' / Bind to CharIndex's.
        If (Bind_X <> -1) Then
           If (Bind_Y <> -1) Then
                Grh_Render temp_particle.Grh, temp_particle.X + Bind_X, temp_particle.Y + Bind_Y, rgb_list(), True, True, alpha_blend, False
              Exit Sub
           End If
        End If
       
        'Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, True, True, alpha_blend, , , temp_particle.angle
                Grh_Render temp_particle.Grh, _
                temp_particle.X + screen_x, _
                temp_particle.Y + screen_y, _
                rgb_list(), 1, False, True _
                , True
    End If
End Sub
 
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
   
    'Set colors
    If UserMinHP = 0 And frmConnect.Visible = False Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
       
    If particle_group_list(particle_group_index).alive_counter Then
   
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
        
        ' / Bind To CharIndex's.
        Dim tmp_X As Integer, tmp_Y As Integer
       
        If Get_Pos(particle_group_index) Then
            tmp_X = particle_group_list(particle_group_index).Now_Viaje_X
            tmp_Y = particle_group_list(particle_group_index).Now_Viaje_Y
        Else
            tmp_X = screen_x
            tmp_Y = screen_y
        End If
   
   
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
       
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey, tmp_X, tmp_Y
                           
        Next loopc
       
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
   
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub
 
Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero ([URL='mailto:juansotuyo@hotmail.com']juansotuyo@hotmail.com[/URL])
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type
    End If
End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function
 
Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
   
            Particle_Group_Map_Pos_Set = True
        End If
    End If
End Function
Public Function Particle_Group_Move(ByVal particle_group_index As Long, ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim nX As Integer
    Dim nY As Integer
   
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Particle_Group_Move = False
        Exit Function
    End If
   
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
   
        map_x = particle_group_list(particle_group_index).map_x
        map_y = particle_group_list(particle_group_index).map_y
       
        nX = map_x
        nY = map_y
       
        Convert_Heading_to_Direction Heading, nX, nY
       
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
            'Move it
            particle_group_list(particle_group_index).map_x = nX
            particle_group_list(particle_group_index).map_y = nY
           
            Particle_Group_Move = True
        End If
    End If
End Function
 
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next

    Dim Temp As particle_group
    Dim i As Integer
   
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    Exit For
                End If
            Next i
        End If
    End If
   
    particle_group_list(particle_group_index) = Temp
           
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub
 
Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                               
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).active = True
   
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    'color
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
 
    'handle
    particle_group_list(particle_group_index).ID = ID
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Sub
 
Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
   
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
   
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
       
    End If
 
End Function
 
Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************
 
Dim i As Integer
 
For i = 1 To charlist(char_index).particle_count
    If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).particle_group(i)
        Exit Function
    End If
Next i
 
Char_Particle_Group_Find = -1
 
End Function
 
Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
   
    Char_Particle_Group_Next_Open = loopc
 
Exit Function
 
ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1
 
End Function
 
Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
   
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If
 
End Function
 
Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
   
    If Char_Check(char_index) Then
        For i = 1 To charlist(char_index).particle_count
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
    End If
   
End Function
 
Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
 
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
 
Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero ([URL='mailto:juansotuyo@hotmail.com']juansotuyo@hotmail.com[/URL])
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
   
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.FrameCounter = 1
            'End If
        End If
    End If
   
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
   
    'Center Grh over X, Y pos
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
   
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
    End If
   
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        Width, Height, alpha_blend ', grh.angle
End Sub
Private Function Char_Check(ByVal char_index As Integer) As Boolean
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
   
End Function
 
Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False, Optional ByVal KillAnim As Boolean = 0)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
On Error Resume Next
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
   
    If Grh.GrhIndex = 0 Then Exit Sub
       
    'Animation
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed * Movement_Speed
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If Grh.Loops <> -1 Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
 
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    'If Not Grh_Check(Grh.grhindex) Then Exit Sub
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
       
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
   
    If v_centered Then
        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * 32) + 32
        End If
    End If
   
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend
 
End Sub
 Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, _
                            ByRef MaskDC As Long, _
                            ByVal src_x As Integer, _
                            ByVal src_y As Integer, _
                            ByVal src_width As Integer, _
                            ByVal src_height As Integer)

    Dim X          As Integer

    Dim Y          As Integer

    Dim TransColor As Long
   
    'ColorKey = hex(COLOR_KEY)
   
    'Check if it has an alpha component
    'If Len(ColorKey) > 6 Then
    'get rid of alpha
    '    ColorKey = "&H" & Right$(ColorKey, 6)
    'End If
    'piluex prueba
    'TransColor = Val(ColorKey)
    TransColor = &H0
 
    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For X = src_x To src_width + src_x

            If GetPixel(hdcsrc, X, Y) = TransColor Then
                SetPixel MaskDC, X, Y, vbWhite
                SetPixel hdcsrc, X, Y, vbBlack
            Else
                SetPixel MaskDC, X, Y, vbBlack
            End If

        Next X
    Next Y

End Sub

'---------------------------------------FIN PARTICULAS-------------------------------
Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
 
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function
 
End Function
Public Sub Crear_Flecha(ByVal AI As Integer, ByVal CI As Integer, ByVal Grh As Integer, ByVal r As Byte, fallo As Byte)
       
    Dim i As Integer
   
    For i = 1 To MaxFlecha
          If Flechas_list(i).EnUso = 0 Then Exit For
    Next
   
    If i = 0 Then i = 1
   
   
   
    Dim addfalloX As Byte
    Dim addFalloY As Integer
   
    If fallo <> 0 Then
   
    addfalloX = RandomNumber(5, 25)
    addFalloY = RandomNumber(5, 25)
   
    Else
   
    addFalloY = -35
    'addfalloX = 16
   
    End If
   
    Flechas_list(i).xb = charlist(CI).Pos.X * 32 + charlist(CI).MoveOffsetX + addfalloX
    Flechas_list(i).yb = charlist(CI).Pos.Y * 32 + charlist(CI).MoveOffsetY + addFalloY
    Flechas_list(i).Rotacion = r
    Flechas_list(i).X = charlist(AI).Pos.X * 32 + charlist(AI).MoveOffsetX
    Flechas_list(i).Y = charlist(AI).Pos.Y * 32 + charlist(AI).MoveOffsetY '+ addFalloY
    Flechas_list(i).EnUso = 1
    InitGrh Flechas_list(i).Grh, Grh
   
    'Flechas_list(i).Angle = Engine_GetAngle(Flechas_list(i).xb, Flechas_list(i).Y, Flechas_list(i).xb, Flechas_list(i).yb)
   
End Sub
Public Function Create_Particle(ByVal charindex As Integer, ByVal victimCharIndex As Integer, ByVal particle_Index As Integer, ByVal Bind_Speed As Single, Optional ByVal particle_life As Integer = -1)
 
'
' @ Crea particulas sobre char.
 
With StreamData(particle_Index)
     
     Dim tmp_List(3)    As Long
     Dim tmp_Loop       As Long
     
     For tmp_Loop = 0 To 3
         With .colortint(tmp_Loop)
              tmp_List(tmp_Loop) = RGB(.r, .g, .b)
         End With
     Next tmp_Loop
     
    Create_Particle = Particle_Group_Create(charlist(charindex).Pos.X, charlist(charindex).Pos.Y, .grh_list, tmp_List(), .NumOfParticles, particle_Index, _
    .AlphaBlend, IIf(particle_life = -1, .life_counter, particle_life), .Speed, , .X1, .Y1, .angle, _
    .vecx1, .vecx2, .vecy1, .vecy2, _
    .life1, .life2, .friction, .spin_speedL, _
    .gravity, .grav_strength, .bounce_strength, .X2, _
    .Y2, .XMove, .move_x1, .move_x2, .move_y1, _
    .move_y2, .YMove, .spin_speedH, .spin)
     
     'Set the binding info.
     If (Create_Particle <> 0) Then
        particle_group_list(Create_Particle).Bind_To_Char = victimCharIndex
        particle_group_list(Create_Particle).Now_Viaje_X = charlist(charindex).Pixel_X
        particle_group_list(Create_Particle).Now_Viaje_Y = charlist(charindex).Pixel_Y
        particle_group_list(Create_Particle).Bind_Speed = Bind_Speed
     End If
     
End With
 
 
End Function
 
Public Function Get_Pos(ByVal particle_Index As Integer) As Boolean
 
'
' @ Devuelve la posición de una partícula.
 
On Error GoTo errHandle
 
With particle_group_list(particle_Index)
     
     Dim C_Angle    As Single
     Dim Target_X   As Integer
     Dim Target_Y   As Integer
     
     'Si viaja a un char.
     If (.Bind_To_Char <> 0) Then
        'Actualiza la posición de destino.
        Target_X = charlist(.Bind_To_Char).Pixel_X
        Target_Y = charlist(.Bind_To_Char).Pixel_Y
       
        'Calcula el ángulo.
         C_Angle = Engine_GetAngle(.Now_Viaje_X, .Now_Viaje_Y, Target_X, Target_Y)
   
        'Actualiza la posición.
        .Now_Viaje_X = (.Now_Viaje_X + Sin(C_Angle * DegreeToRadian) * .Bind_Speed)
        .Now_Viaje_Y = (.Now_Viaje_Y - Cos(C_Angle * DegreeToRadian) * .Bind_Speed)
       
        'Llegó al target?
        If (.Now_Viaje_X = Target_X) And (.Now_Viaje_Y = Target_Y) Then
           
           ' / Reset binding info.
           .Now_Viaje_X = 0
           .Now_Viaje_Y = 0
           .Bind_Speed = 0
           .Bind_To_Char = 0
           
           Call Particle_Group_Remove(particle_Index)
        End If
       
        Get_Pos = True
       
    Else
       
        Get_Pos = False
       
    End If
   
Exit Function
 
errHandle:
 
Get_Pos = False
End With
 
End Function

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef color As D3DCOLORVALUE)
'*********************************************************
'****** Coded by Dunkan ([email=emanuel.m@dunkancorp.com]emanuel.m@dunkancorp.com[/email]) *******
'*********************************************************
Dim dest(3) As Byte
CopyMemory dest(0), ARGB, 4
color.A = dest(3)
color.r = dest(2)
color.g = dest(1)
color.b = dest(0)
End Function

