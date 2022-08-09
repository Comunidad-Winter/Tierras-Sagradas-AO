VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8445
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
   ScaleHeight     =   8895
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Desc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1560
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   6565
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   7920
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   18
      Top             =   5760
      Width           =   6855
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   17
      Top             =   5454
      Width           =   6855
   End
   Begin VB.Image BSolicitarIngreso 
      Height          =   525
      Left            =   2220
      MousePointer    =   99  'Custom
      Top             =   8250
      Width           =   4050
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      Caption         =   "dddddddddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   2450
      Width           =   6975
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      Caption         =   "dddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "ddddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   878
      Width           =   6975
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "dddddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   3075
      Width           =   6975
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      Caption         =   "ddddddddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   2745
      Width           =   6975
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "ddddddddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2130
      Width           =   7215
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      Caption         =   "dddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1815
      Width           =   6975
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      Caption         =   "ddddddddddddddddddddddddddddddddddddd"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1515
      Width           =   6975
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Omnia Somnia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   180
      Width           =   3735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   5200
      Width           =   6855
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   4920
      Width           =   6855
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   4637
      Width           =   6855
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   4343
      Width           =   6855
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   4072
      Width           =   6855
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   3773
      Width           =   6855
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EsLeader As Boolean

Public Sub ParseGuildInfo(ByVal Buffer As String)

Enemigos.Caption = ReadField(1, Buffer, Asc("¬"))
eleccion.Caption = ReadField(2, Buffer, Asc("¬"))
Oro.Caption = PonerPuntos(ReadField(3, Buffer, Asc("¬")))
fundador.Caption = ReadField(4, Buffer, Asc("¬"))
creacion.Caption = ReadField(5, Buffer, Asc("¬"))
lider.Caption = ReadField(6, Buffer, Asc("¬"))

If ReadField(2, Buffer, Asc("¬")) = 2 Then
    eleccion.ForeColor = &H80&
    eleccion = "HORDA INFERNAL"
ElseIf ReadField(2, Buffer, Asc("¬")) = 4 Then
    eleccion.ForeColor = &HC00000
    eleccion = "ALIANZA IMPERIAL"
ElseIf ReadField(2, Buffer, Asc("¬")) = 3 Then
    eleccion.ForeColor = &H404040
    eleccion = "NEUTRAL"
End If

If ReadField(7, Buffer, Asc("¬")) = "Fermin" And ReadField(8, Buffer, Asc("¬")) = "Fermin" Then
Aliados.Caption = "-"
ElseIf ReadField(7, Buffer, Asc("¬")) <> "Fermin" And ReadField(8, Buffer, Asc("¬")) = "Fermin" Then
Aliados.Caption = ReadField(7, Buffer, Asc("¬"))
ElseIf ReadField(7, Buffer, Asc("¬")) <> "Fermin" And ReadField(8, Buffer, Asc("¬")) <> "Fermin" Then
Aliados.Caption = "" & ReadField(7, Buffer, Asc("¬")) & " y " & ReadField(8, Buffer, Asc("¬")) & ""
ElseIf ReadField(7, Buffer, Asc("¬")) = "Fermin" And ReadField(8, Buffer, Asc("¬")) <> "Fermin" Then
Aliados.Caption = ReadField(8, Buffer, Asc("¬"))
End If

web.Caption = ReadField(9, Buffer, Asc("¬"))
Miembros.Caption = ReadField(10, Buffer, Asc("¬"))

Dim T As Long

For T = 1 To 8
    Codex(T - 1).Caption = ReadField(10 + T, Buffer, Asc("¬"))
Next T

Dim des As String

des = ReadField(19, Buffer, Asc("¬"))
nombre.Caption = ReadField(20, Buffer, Asc("¬"))
Desc.text = Replace(des, "º", vbCrLf)
Me.Show vbModal, frmMain

End Sub
Public Sub ParseSubGuildInfo(ByVal Buffer As String)

Enemigos.Caption = ReadField(1, Buffer, Asc("¬"))
eleccion.Caption = ReadField(2, Buffer, Asc("¬"))
Oro.Caption = ReadField(3, Buffer, Asc("¬"))
fundador.Caption = ReadField(4, Buffer, Asc("¬"))
creacion.Caption = ReadField(5, Buffer, Asc("¬"))
lider.Caption = ReadField(6, Buffer, Asc("¬"))

If ReadField(7, Buffer, Asc("¬")) = "Fermin" And ReadField(8, Buffer, Asc("¬")) = "Fermin" Then
Aliados.Caption = "-"
ElseIf ReadField(7, Buffer, Asc("¬")) <> "Fermin" And ReadField(8, Buffer, Asc("¬")) = "Fermin" Then
Aliados.Caption = ReadField(7, Buffer, Asc("¬"))
ElseIf ReadField(7, Buffer, Asc("¬")) <> "Fermin" And ReadField(8, Buffer, Asc("¬")) <> "Fermin" Then
Aliados.Caption = "" & ReadField(7, Buffer, Asc("¬")) & " y " & ReadField(8, Buffer, Asc("¬")) & ""
ElseIf ReadField(7, Buffer, Asc("¬")) = "Fermin" And ReadField(8, Buffer, Asc("¬")) <> "Fermin" Then
Aliados.Caption = ReadField(8, Buffer, Asc("¬"))
End If

web.Caption = ReadField(9, Buffer, Asc("¬"))
Miembros.Caption = ReadField(10, Buffer, Asc("¬"))

Dim T As Long

For T = 1 To 8
    Codex(T - 1).Caption = ReadField(10 + T, Buffer, Asc("¬"))
Next T

Dim des As String

des = ReadField(19, Buffer, Asc("¬"))
nombre.Caption = ReadField(20, Buffer, Asc("¬"))
Desc.text = Replace(des, "º", vbCrLf)
Me.Show , frmMain

End Sub
Private Sub BCerrar_Click()
Unload Me
End Sub

Private Sub BSolicitarIngreso_Click()
Dim f$

f$ = "SOLICITUD" & nombre
f$ = f$ & "," & Replace(Replace("holi", ",", ";"), vbCrLf, "º")

Call SendData(f$)

Unload Me
End Sub
Private Sub BSolicitarIngreso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BSolicitarIngreso.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildBrief_BSolicitarIngreso_I.jpg")
End Sub
Private Sub BSolicitarIngreso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BSolicitarIngreso.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildBrief_BSolicitarIngreso_A.jpg")
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildBrief_Main.jpg")
BSolicitarIngreso.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildBrief_BSolicitarIngreso_N.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BSolicitarIngreso.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\GuildBrief_BSolicitarIngreso_N.jpg")
End Sub
Private Sub Image1_Click()
Unload Me
End Sub

