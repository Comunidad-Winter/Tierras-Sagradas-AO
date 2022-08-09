VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creación de un Clan"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmGuildFoundation.frx":0000
      Left            =   440
      List            =   "frmGuildFoundation.frx":000D
      TabIndex        =   2
      Text            =   "ELEGI ALINEACION"
      Top             =   2180
      Width           =   3260
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   440
      TabIndex        =   1
      Top             =   2750
      Width           =   3250
   End
   Begin VB.TextBox txtClanName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   440
      TabIndex        =   0
      Top             =   1660
      Width           =   3250
   End
   Begin VB.Image BCancelar 
      Height          =   600
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image BSiguiente 
      Height          =   600
      Left            =   2160
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
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
Private Const InterfaceName As String = "GuildFundation"

Private Sub BCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BCancelar.Picture = ChangeButtonState(Apretado, "BCancelar")
End Sub

Private Sub BCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BCancelar.Tag = "0" Then
    Call ChangeButtonsNormal
    BCancelar.Picture = ChangeButtonState(Iluminado, "BCancelar")
    BCancelar.Tag = "1"
End If
End Sub

Private Sub BCancelar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub

Private Sub BSiguiente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BSiguiente.Picture = ChangeButtonState(Apretado, "BSiguiente")
End Sub

Private Sub BSiguiente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BSiguiente.Tag = "0" Then
    Call ChangeButtonsNormal
    BSiguiente.Picture = ChangeButtonState(Iluminado, "BSiguiente")
    BSiguiente.Tag = "1"
End If
End Sub

Private Sub BSiguiente_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub

Private Sub BSiguiente_Click()

If txtClanName = "" Then Exit Sub
If Combo1.ListIndex < 0 Then Exit Sub

If Len(txtClanName.text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        Mensaje.Escribir "Nombre invalido."
        Exit Sub
    End If
Else
    Mensaje.Escribir "Nombre demasiado extenso."
    Exit Sub
End If

ClanName = txtClanName
Site = Text2
Unload Me
frmGuildDetails.Show , Me
End Sub

Private Sub BCancelar_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
Me.SetFocus
End Sub

Private Sub Form_Load()

Me.Picture = General_Load_Interface_Picture("GuildFundation_Main.jpg")

If Len(txtClanName.text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        Mensaje.Escribir "Nombre invalido."
        Exit Sub
    End If
Else
        Mensaje.Escribir "Nombre demasiado extenso."
        Exit Sub
End If

ChangeButtonsNormal

End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

BCancelar.Picture = ChangeButtonState(BNormal, "BCancelar")
BSiguiente.Picture = ChangeButtonState(BNormal, "BSiguiente")


Dim j
For Each j In Me
    j.Tag = "0"
Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub
