VERSION 5.00
Begin VB.Form frmGuildDetails 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6855
   ClientLeft      =   2535
   ClientTop       =   2100
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1480
      Left            =   620
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   320
      Width           =   5665
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   5640
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   5220
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   4820
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   4410
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   4010
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   3180
      Width           =   5680
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   5680
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   1
      Left            =   4680
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   1905
   End
End
Attribute VB_Name = "frmGuildDetails"
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

Option Explicit
Private Const InterfaceName As String = "GuildDetails"


Private Sub Form_Load()

Me.Picture = General_Load_Interface_Picture("GuildDetails_Main.jpg")

ChangeButtonsNormal

End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index

Case 0
    Unload Me
Case 1
    Dim fdesc$
    fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
'    If Not AsciiValidos(fdesc$) Then
'        MsgBox "La descripcion contiene caracteres invalidos"
'        Exit Sub
'    End If
    
    Dim k As Integer
    Dim Cont As Integer
    Cont = 0
    For k = 0 To txtCodex1.UBound
'        If Not AsciiValidos(txtCodex1(k)) Then
'            MsgBox "El codex tiene invalidos"
'            Exit Sub
'        End If
        If Len(txtCodex1(k).text) > 0 Then Cont = Cont + 1
    Next k
    If Cont < 4 Then
            MsgBox "Debes definir al menos cuatro mandamientos."
            Exit Sub
    End If
    
    Dim chunk$
    
    If CreandoClan Then
        chunk$ = "CIG" & fdesc$
        chunk$ = chunk$ & "¬" & ClanName & "¬" & Site & "¬" & Cont
    Else
        chunk$ = "DESCOD" & fdesc$ & "¬" & Cont
    End If
    
    
    
    For k = 0 To txtCodex1.UBound
        chunk$ = chunk$ & "¬" & txtCodex1(k)
    Next k
    
    
    Call SendData(chunk$)
    
    CreandoClan = False
    
    Unload Me
    
End Select



End Sub
Private Sub Form_Deactivate()

'If Not frmGuildLeader.Visible Then
'    Me.SetFocus
'Else
'    'Unload Me
'End If
'

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FormDrag Me
End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

Image1(1).Picture = ChangeButtonState(BNormal, "BAceptar")
Image1(0).Picture = ChangeButtonState(BNormal, "BCancelar")

Image1(0).Tag = "0"
Image1(1).Tag = "0"

Me.Tag = "0"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Image1(0).Picture = ChangeButtonState(Apretado, "BCancelar")
If Index = 1 Then Image1(1).Picture = ChangeButtonState(Apretado, "BAceptar")
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 And Image1(0).Tag = "0" Then
    Call ChangeButtonsNormal
    Image1(0).Picture = ChangeButtonState(Iluminado, "BCancelar")
    Image1(0).Tag = "1"
End If

If Index = 1 And Image1(1).Tag = "0" Then
    Call ChangeButtonsNormal
    Image1(1).Picture = ChangeButtonState(Iluminado, "BAceptar")
    Image1(1).Tag = "1"
End If
End Sub
