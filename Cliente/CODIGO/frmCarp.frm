VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCarp.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "1"
      Top             =   3600
      Width           =   855
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   2120
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4230
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2640
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   360
      Top             =   2520
      Width           =   1545
   End
End
Attribute VB_Name = "frmCarp"
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

Private Sub Image2_Click()

On Error Resume Next

'If Not IsNumeric(Text1) Then
'    Mensaje.Escribir "Introduce un valor numerico"
'    Exit Sub
'End If

Call SendData("CNC" & ObjCarpintero(lstArmas.ListIndex))

'Unload Me

End Sub

Private Sub Image1_Click()

Unload Me

End Sub

Private Sub Form_Deactivate()

'Me.SetFocus
End Sub
