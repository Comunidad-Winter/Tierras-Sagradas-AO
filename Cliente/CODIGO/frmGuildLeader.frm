VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administraci�n del Clan"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
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
   ScaleHeight     =   6510
   ScaleWidth      =   5880
   Begin VB.CommandButton Command9 
      Caption         =   "Propuestas de alianzas"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5460
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5970
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Propuestas de paz"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4950
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar URL de la web del clan"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3930
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command10 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":069A
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":07EC
         Left            =   120
         List            =   "frmGuildLeader.frx":07EE
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":07F0
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0942
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0A94
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0BE6
         Left            =   120
         List            =   "frmGuildLeader.frx":0BE8
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
      Begin VB.CommandButton cmdElecciones 
         Caption         =   "Abrir elecciones"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0BEA
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1850
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0D3C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1170
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   840
         ItemData        =   "frmGuildLeader.frx":0E8E
         Left            =   120
         List            =   "frmGuildLeader.frx":0E90
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1620
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdElecciones_Click()
    Call SendData("ABREELEC")
    Unload Me
End Sub

Private Sub Command1_Click()

frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.listIndex))

'Unload Me

End Sub

Private Sub Command10_Click()
frmGuildBrief.EsLeader = False
Call SendData("CLANDETAILS" & guildslist.List(guildslist.listIndex))
End Sub

Private Sub Command2_Click()


frmCharInfo.frmmiembros = True
Call SendData("1HRINFO<" & members.List(members.listIndex))

'Unload Me
End Sub

Private Sub Command3_Click()

Dim k$

k$ = Replace(txtguildnews, vbCrLf, "�")

Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command4_Click()

frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & guildslist.List(guildslist.listIndex))

'Unload Me

End Sub

Private Sub Command5_Click()

Call frmGuildDetails.Show(vbModal, frmGuildLeader)

'Unload Me

End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
'Unload Me
End Sub

Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub
Private Sub Command9_Click()
Call SendData("ENVALPRO")
End Sub


Private Sub Command8_Click()
Unload Me
frmMain.SetFocus
End Sub


Public Sub ParseLeaderInfo(ByVal Data As String)

If Me.Visible Then Exit Sub

Dim r%, T%

r% = Val(ReadField(1, Data, Asc("�")))

For T% = 1 To r%
    guildslist.AddItem ReadField(1 + T%, Data, Asc("�"))
Next T%

r% = Val(ReadField(T% + 1, Data, Asc("�")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

For k% = 1 To r%
    members.AddItem ReadField(T% + 1 + k%, Data, Asc("�"))
Next k%

txtguildnews = Replace(ReadField(T% + k% + 1, Data, Asc("�")), "�", vbCrLf)

T% = T% + k% + 2

r% = Val(ReadField(T%, Data, Asc("�")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(T% + k%, Data, Asc("�"))
Next k%
Dim I, o, TotalItems As Long

TotalItems = frmGuildLeader.guildslist.ListCount
    
    For I = 0 To TotalItems
    For o = 0 To TotalItems
        If frmGuildLeader.guildslist.List(I) = "cerrado" & o Then
            frmGuildLeader.guildslist.RemoveItem (I)
        End If
    Next
Next
Me.Show , frmMain

End Sub

Public Sub ParseSubLeaderInfo(ByVal Data As String)

If Me.Visible Then Exit Sub

Dim r%, T%

r% = Val(ReadField(1, Data, Asc("�")))

For T% = 1 To r%
    guildslist.AddItem ReadField(1 + T%, Data, Asc("�"))
Next T%

r% = Val(ReadField(T% + 1, Data, Asc("�")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

For k% = 1 To r%
    members.AddItem ReadField(T% + 1 + k%, Data, Asc("�"))
Next k%

txtguildnews = Replace(ReadField(T% + k% + 1, Data, Asc("�")), "�", vbCrLf)

T% = T% + k% + 2

r% = Val(ReadField(T%, Data, Asc("�")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(T% + k%, Data, Asc("�"))
Next k%
Dim I, o, TotalItems As Long

TotalItems = frmGuildLeader.guildslist.ListCount
    
    For I = 0 To TotalItems
    For o = 0 To TotalItems
        If frmGuildLeader.guildslist.List(I) = "cerrado" & o Then
            frmGuildLeader.guildslist.RemoveItem (I)
        End If
    Next
Next
txtnews.Visible = False
cmdElecciones.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command9.Visible = False
Frame1.Left = 2880
Frame1.Top = 0
Frame2.Visible = False
Command8.Top = 2280
Command8.Left = 1440
frmGuildLeader.Height = 3300
Command10.Visible = True
Command4.Visible = False
frmGuildBrief.EsLeader = False
Me.Show , frmMain

End Sub


Private Sub Form_Deactivate()
'Me.SetFocus
End Sub
