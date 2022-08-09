VERSION 5.00
Begin VB.Form Drag 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Drag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long

Dim bytRegion() As Byte
Dim nBytes As Long

Public Sub MakeRegionCtrl(ByRef frm As Control, ByVal TrnsColor As Long)
frm.BorderStyle = 0
Dim ScaleSize As Long
Dim Width, Height As Long
Dim rgnMain As Long
Dim X, Y As Long
Dim rgnPixel As Long
Dim RGBColor As Long
Dim dcMain As Long
Dim bmpMain As Long
ScaleSize = frm.ScaleMode
frm.ScaleMode = 3
Width = frm.ScaleX(frm.Picture.Width, vbHimetric, vbPixels)
Height = frm.ScaleY(frm.Picture.Height, vbHimetric, vbPixels)
frm.Width = Width * Screen.TwipsPerPixelX
frm.Height = Height * Screen.TwipsPerPixelY
rgnMain = CreateRectRgn(0, 0, Width, Height)
dcMain = CreateCompatibleDC(frm.hDC)
bmpMain = SelectObject(dcMain, frm.Picture.handle)
For Y = 0 To Height
    For X = 0 To Width
        RGBColor = GetPixel(dcMain, X, Y)
        If RGBColor = TrnsColor Then
            rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1)
            CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
            DeleteObject rgnPixel
        End If
    Next X
Next Y
SelectObject dcMain, bmpMain
DeleteDC dcMain
DeleteObject bmpMain
If rgnMain <> 0 Then
 nBytes = GetRegionData(rgnMain, 0, ByVal 0&)
    If nBytes > 0 Then
        ReDim bytRegion(0 To nBytes - 1)
        nBytes = GetRegionData(rgnMain, nBytes, bytRegion(0))
    End If
    SetWindowRgn frm.hWnd, rgnMain, True
    'CenterForm Me
End If
frm.ScaleMode = 3 'ScaleSize
End Sub '''

Private Sub MakeRegion(ByRef frm As Form, ByVal TrnsColor As Long)
frm.BorderStyle = 0
Dim ScaleSize As Long
Dim Width, Height As Long
Dim rgnMain As Long
Dim X, Y As Long
Dim rgnPixel As Long
Dim RGBColor As Long
Dim dcMain As Long
Dim bmpMain As Long
ScaleSize = frm.ScaleMode
frm.ScaleMode = 3
Width = frm.ScaleX(frm.Picture.Width, vbHimetric, vbPixels)
Height = frm.ScaleY(frm.Picture.Height, vbHimetric, vbPixels)
frm.Width = Width * Screen.TwipsPerPixelX
frm.Height = Height * Screen.TwipsPerPixelY
rgnMain = CreateRectRgn(0, 0, Width, Height)
dcMain = CreateCompatibleDC(frm.hDC)
bmpMain = SelectObject(dcMain, frm.Picture.handle)
For Y = 0 To Height
    For X = 0 To Width
        RGBColor = GetPixel(dcMain, X, Y)
        If RGBColor = TrnsColor Then
            rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1)
            CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
            DeleteObject rgnPixel
        End If
    Next X
Next Y
SelectObject dcMain, bmpMain
DeleteDC dcMain
DeleteObject bmpMain
If rgnMain <> 0 Then
 nBytes = GetRegionData(rgnMain, 0, ByVal 0&)
    If nBytes > 0 Then
        ReDim bytRegion(0 To nBytes - 1)
        nBytes = GetRegionData(rgnMain, nBytes, bytRegion(0))
    End If
    SetWindowRgn frm.hWnd, rgnMain, True
    'CenterForm Me
End If
frm.ScaleMode = 3 'ScaleSize
End Sub

Public Sub SetPicture(Filename As String, ClrTransparent As Long)
If InStr(1, Filename, "0.bmp") > 0 Then Exit Sub
Set Me.Picture = LoadPicture(Filename)
MakeRegion Me, ClrTransparent
End Sub ''

Private Sub Form_Load()
'Drag.SetPicture DirGraficos & GrafNum & ".bmp", 0

End Sub
'LaX = -7
'LaY = 181
''Call Form_MouseDown(vbLeftButton, 0, 1, 1)
'If bNoResChange = True Then
'    If RemDragX = 0 Then
'        Drag.Left = (frmMain.picInv.Left + (LaX * 15)) + 8100  '* 4
'    Else
'        Drag.Left = RemDragX
'    End If
'    If RemDragY = 0 Then
'        Drag.Top = (frmMain.picInv.Top + (LaY * 15)) + 2100  'Y * 4
'    Else
'        Drag.Top = RemDragY
'    End If
'Else
'    If RemDragX = 0 Then
'        Drag.Left = (frmMain.picInv.Left + (LaX * 15)) + 9700  '* 4
'    Else
'        Drag.Left = RemDragX
'    End If
'    If RemDragY = 0 Then
'        Drag.Top = (frmMain.picInv.Top + (LaY * 15)) + 3200  'Y * 4
'    Else
'        Drag.Top = RemDragY
'    End If
'End If
'End Sub
'Public Sub Repocisionar()
'LaX = -7
'LaY = 181
'If bNoResChange = True Then
'    If RemDragX = 0 Then
'        Drag.Left = (frmMain.picInv.Left + (LaX * 15)) + 8100  '* 4
'    Else
'        Drag.Left = RemDragX
'    End If
'    If RemDragY = 0 Then
'        Drag.Top = (frmMain.picInv.Top + (LaY * 15)) + 2100  'Y * 4
'    Else
'        Drag.Top = RemDragY
'    End If
'Else
'    If RemDragX = 0 Then
'        Drag.Left = (frmMain.picInv.Left + (LaX * 15)) + 9700  '* 4
'    Else
'        Drag.Left = RemDragX
'    End If
'    If RemDragY = 0 Then
'        Drag.Top = (frmMain.picInv.Top + (LaY * 15)) + 3200  'Y * 4
'    Else
'        Drag.Top = RemDragY
'    End If
'End If
'End Sub
'Public Sub Dibujar()
'Drag.SetPicture "C:\Archivos de programa\Tierras Perdidas 2.9.2\Data\GRAFICOS\" & GrafNum & ".bmp", 0
'End Sub'''''

'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'AllowDrag = 1
''Me.Left = (X * 15) + 8285
''Me.Top = (Y * 15) + 4790
'FormDrag Me
' '           Me.Left = Me.Left + (X * 15) ' (frmMain.picInv.Left + (X * 15)) '* 4
' '           Me.Top = Me.Top + (Y * 15) ' (frmMain.picInv.Top + (Y * 15)) 'Y * 4
'End Sub''

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If AllowDrag = 1 Then
'    If Button = 2 Then
'    MsgBox "a"
'    End If
'    'FormDrag Me
'    'Me.Left = 8282 + (X * 15) '(X * 14) + 8285
'    'Me.Top = 4790 + (X * 15) '(Y * 14) + 4790
'End If
''MsgBox X
''If X >= 13 And X <= 15 Then
''    If Y >= 14 And Y <= 15 Then
''        Call frmMain.TirarItem
''        Me.Visible = False
''        'MsgBox "drag"
''    End If
''End If
''AllowDrag = 0
''Unload Me
'End Sub''

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AllowDrag = 0
'Repocisionar
End Sub

Private Sub Form_OLECompleteDrag(Effect As Long)
MsgBox "a"
End Sub
