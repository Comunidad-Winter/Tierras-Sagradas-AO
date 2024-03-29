Attribute VB_Name = "Carteles"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer


Sub InitCartel(Ley As String, Grh As Long)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim i As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    i = 0
    Call DarFormato(Leyenda, i, k, anti)
    i = 0
    Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
        
       i = i + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To i)
Else
    Exit Sub
End If
End Sub


Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(k) = mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti)
End If
End Function

Sub DibujarCartel()
 
        If Not Cartel Then Exit Sub
 
        Dim X As Integer, Y As Integer, j As Long, desp As Integer
 
        X = XPosCartel + 25
        Y = YPosCartel + 60
   
        Call engine.Draw_GrhIndex(textura, XPosCartel, YPosCartel)
 
        For j = 0 To UBound(LeyendaFormateada)
                Texto.Engine_Text_Draw X, Y + desp, LeyendaFormateada(j), -1
                desp = desp + (frmMain.font.size) + 5
        Next
 
End Sub
Public Sub realizarSuma(ByVal charindex As Integer)

        Dim tmpChar, tmpName As String
        Dim i As Long, pos As Integer
        
With charlist(charindex)
        pos = InStr(.Nombre, "<")
        If pos = 0 Then pos = Len(.Nombre) + 2
        tmpName = left$(.Nombre, pos - 2)
        
        .sumatoriaEstrella = 0
        For i = 1 To Len(tmpName)
            tmpChar = Asc(mid(tmpName, i, 1))
            If tmpChar >= 65 And tmpChar <= 90 Then
                .sumatoriaEstrella = .sumatoriaEstrella + 4.2
            Else
                .sumatoriaEstrella = .sumatoriaEstrella + 3.8
            End If
        Next i
End With

End Sub
