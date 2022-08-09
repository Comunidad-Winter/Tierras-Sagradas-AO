VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (destination As Any, _
                                       source As Any, _
                                       ByVal length As Long)
 
'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
 
        Width As Long
        Height As Long
 
End Type
 
Private Type POINTAPI
 
        X As Long
        Y As Long
 
End Type
 
Private Type CharVA
 
        Vertex(0 To 3) As TLVERTEX
 
End Type
 
Private Type VFH
 
        BitmapWidth As Long         'Size of the bitmap itself
        BitmapHeight As Long
        CellWidth As Long           'Size of the cells (area for each character)
        CellHeight As Long
        BaseCharOffset As Byte      'The character we start from
        CharWidth(0 To 255) As Byte 'The actual factual width of each character
        CharVA(0 To 255) As CharVA
 
End Type
 
Private Type CustomFont
 
        HeaderInfo As VFH           'Holds the header information
        Texture As Direct3DTexture8 'Holds the texture of the text
        RowPitch As Integer         'Number of characters per row
        RowFactor As Single         'Percentage of the texture width each character takes
        ColFactor As Single         'Percentage of the texture height each character takes
        CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
        TextureSize As POINTAPI     'Size of the texture
 
End Type
 
'public Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cfonts(1 To 2) As CustomFont  ' _Default2 As CustomFont
Sub Engine_Text_Draw(ByVal left As Long, _
                     ByVal top As Long, _
                     ByVal text As String, _
                     ByVal color As Long, _
                     Optional ByVal Alpha As Byte = 255, _
                     Optional ByVal Center As Boolean = False)
 
        Dim aux As D3DCOLORVALUE
 
        If Alpha <> 255 Then
                ARGBtoD3DCOLORVALUE color, aux
                color = D3DColorARGB(Alpha, aux.r, aux.g, aux.b)
        End If
   
        Engine_Render_Text cfonts(1), text, left, top, color, Center
 
End Sub
 
Private Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef color As D3DCOLORVALUE)
 
        Dim dest(3) As Byte
 
        CopyMemory dest(0), ARGB, 4
        color.A = dest(3)
        color.r = dest(2)
        color.g = dest(1)
        color.b = dest(0)
 
End Function
 
Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, _
                               ByVal text As String, _
                               ByVal X As Long, _
                               ByVal Y As Long, _
                               ByVal color As Long, _
                               Optional ByVal Center As Boolean = False)
 
        Dim TempVA(0 To 3) As TLVERTEX
 
        Dim tempstr()      As String
 
        Dim Count          As Integer
 
        Dim ascii()        As Byte
 
        Dim i              As Long
 
        Dim j              As Long
 
        Dim TempColor      As Long
 
        Dim ResetColor     As Byte
 
        Dim YOffset        As Single
 
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DirectDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
        'Check if we have the device
        If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
        'Check for valid text to render
        If LenB(text) = 0 Then Exit Sub
   
        'Get the text into arrays (split by vbCrLf)
        tempstr = Split(text, vbCrLf)
   
        'Set the temp color (or else the first character has no color)
        TempColor = color
 
        'Set the texture
        DirectDevice.SetTexture 0, UseFont.Texture
   
        If Center Then
                X = X - Engine_GetTextWidth(cfonts(1), text) * 0.5
        End If
   
        'Loop through each line if there are line breaks (vbCrLf)
        For i = 0 To UBound(tempstr)
 
                If Len(tempstr(i)) > 0 Then
                        YOffset = i * UseFont.CharHeight
                        Count = 0
       
                        'Convert the characters to the ascii value
                        ascii() = StrConv(tempstr(i), vbFromUnicode)
       
                        'Loop through the characters
                        For j = 1 To Len(tempstr(i))
   
                                'Copy from the cached vertex array to the temp vertex array
                                CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), 32 * 4
 
                                'Set up the verticies
                                TempVA(0).X = X + Count
                                TempVA(0).Y = Y + YOffset
                       
                                TempVA(1).X = TempVA(1).X + X + Count
                                TempVA(1).Y = TempVA(0).Y
 
                                TempVA(2).X = TempVA(0).X
                                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
 
                                TempVA(3).X = TempVA(1).X
                                TempVA(3).Y = TempVA(2).Y
                       
                                'Set the colors
                                TempVA(0).color = TempColor
                                TempVA(1).color = TempColor
                                TempVA(2).color = TempColor
                                TempVA(3).color = TempColor
                       
                                'Draw the verticies
                                DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                     
                                'Shift over the the position to render the next character
                                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
 
                                'Check to reset the color
                                If ResetColor Then
                                        ResetColor = 0
                                        TempColor = color
                                End If
               
                        Next j
           
                End If
 
        Next i
   
End Sub
 
Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, _
                                     ByVal text As String) As Integer
 
        Dim i As Long
 
        'Make sure we have text
        If LenB(text) = 0 Then Exit Function
   
        'Loop through the text
        For i = 1 To Len(text)
       
                'Add up the stored character widths
                Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(text, i, 1)))
       
        Next i
 
End Function
 
Sub Engine_Init_FontTextures()
 
        On Error GoTo eDebug:
 
        Dim TexInfo As D3DXIMAGE_INFO_A
 
        'Check if we have the device
        If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
        '*** Default font ***
   
        'Set the texture
        Set cfonts(1).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.Path & "\Data\INIT\texdefault.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
   
        'Store the size of the texture
        cfonts(1).TextureSize.X = TexInfo.Width
        cfonts(1).TextureSize.Y = TexInfo.Height
   
        Exit Sub
 
eDebug:
 
        If err.Number = "-2005529767" Then
                MsgBox "Error en la textura utilizada de DirectX 8", vbCritical
        End If
 
        End
 
End Sub
 
Sub Engine_Init_FontSettings()
 
        '*********************************************************
        '****** Coded by Dunkan ([email=emanuel.m@dunkancorp.com]emanuel.m@dunkancorp.com[/email]) *******
        '*********************************************************
        Dim FileNum  As Byte
 
        Dim LoopChar As Long
 
        Dim Row      As Single
 
        Dim u        As Single
 
        Dim v        As Single
 
        '*** Default font ***
 
        'Load the header information
        FileNum = FreeFile
        Open App.Path & "\Data\INIT\texdefault.dat" For Binary As #FileNum
        Get #FileNum, , cfonts(1).HeaderInfo
        Close #FileNum
   
        'Calculate some common values
        cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
        cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
        cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
        cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
   
        'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
        For LoopChar = 0 To 255
       
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
                u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
                v = Row * cfonts(1).RowFactor
 
                'Set the verticies
                With cfonts(1).HeaderInfo.CharVA(LoopChar)
                        .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
                        .Vertex(0).rhw = 1
                        .Vertex(0).tu = u
                        .Vertex(0).tv = v
                        .Vertex(0).X = 0
                        .Vertex(0).Y = 0
                        .Vertex(0).Z = 0
           
                        .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
                        .Vertex(1).rhw = 1
                        .Vertex(1).tu = u + cfonts(1).ColFactor
                        .Vertex(1).tv = v
                        .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
                        .Vertex(1).Y = 0
                        .Vertex(1).Z = 0
           
                        .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
                        .Vertex(2).rhw = 1
                        .Vertex(2).tu = u
                        .Vertex(2).tv = v + cfonts(1).RowFactor
                        .Vertex(2).X = 0
                        .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
                        .Vertex(2).Z = 0
           
                        .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
                        .Vertex(3).rhw = 1
                        .Vertex(3).tu = u + cfonts(1).ColFactor
                        .Vertex(3).tv = v + cfonts(1).RowFactor
                        .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
                        .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
                        .Vertex(3).Z = 0
                End With
       
        Next LoopChar
 
End Sub
 

