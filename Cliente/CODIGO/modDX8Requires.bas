Attribute VB_Name = "modDX8Requires"
Option Explicit

Public vertList(3) As TLVERTEX

Public Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

'DX8 Objects
Public DirectX As DirectX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8
Public DirectD3D8 As D3DX8

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX2
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu1 As Single
    tv1 As Single
    tu2 As Single
    tv2 As Single
End Type

Public Const PI As Single = 3.14159265358979
Public base_light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte

'JOJOJO
Public engine As New clsDX8Engine
'JOJOJO
Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal A As Long) As Long
        
    Dim c As Long
        
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = A * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    ARGB = c

End Function
Function Map_InBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************

      If (X < XMinMapSize) Or (X > XMaxMapSize) Or (Y < YMinMapSize) Or (Y > YMaxMapSize) Then
            Map_InBounds = False

            Exit Function

      End If
    
      Map_InBounds = True
End Function
Public Sub Engine_Set_TileBuffer(ByVal setTileBufferSize As Single)
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    TileBufferSize = setTileBufferSize
    
End Sub
Public Function Engine_Get_TileBuffer() As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    Engine_Get_TileBuffer = TileBufferSize
    
End Function
