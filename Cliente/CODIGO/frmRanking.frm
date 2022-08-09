VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmRanking 
   BorderStyle     =   0  'None
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   LinkTopic       =   "Form2"
   ScaleHeight     =   5385
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Informacion 
      Height          =   855
      Left            =   600
      TabIndex        =   23
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      _Version        =   393217
      BackColor       =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmRanking.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Corbel"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Reputacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label ChangeRank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   720
      Width           =   375
   End
   Begin VB.Label NameRank 
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "Century751 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   680
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   19
      Top             =   3390
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   18
      Top             =   3150
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   16
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   15
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   14
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   13
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   12
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   11
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Puntaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   10
      Top             =   1215
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terremoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   1185
      Width           =   975
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   8
      Left            =   2415
      Top             =   4700
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   7
      Left            =   1440
      Top             =   4700
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   6
      Left            =   465
      Top             =   4700
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   5
      Left            =   2400
      Top             =   4280
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   4
      Left            =   1440
      Top             =   4280
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   3
      Left            =   465
      Top             =   4280
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   2
      Left            =   2415
      Top             =   3850
      Width           =   840
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   1
      Left            =   1440
      Top             =   3850
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3240
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Botones 
      Height          =   285
      Index           =   0
      Left            =   465
      Top             =   3850
      Width           =   840
   End
End
Attribute VB_Name = "frmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RankingActual As Byte '1=GENERAL ; 2=SEMANAL

Public Sub MostrarRanking(Rdata As String)
    
Dim NickTemporal As String
Dim j As Long
'TOP10.ListItems.Clear
NickTemporal = ""

For j = 1 To 10
    NickTemporal = ReadField(j, Rdata, Asc(","))
    
    Nombre(j).Caption = ReadField(1, NickTemporal, Asc("-"))
    Puntaje(j).Caption = ReadField(2, NickTemporal, Asc("-"))
Next j

If RankingActual = 2 Then
    Reputacion.Caption = PonerPuntos(ReadField(11, Rdata, Asc(",")))
End If
    
End Sub
Private Sub ChangeRank_Click()
    Dim i As Long

    If RankingActual = 1 Then 'GENERAL
        For i = 0 To 8
            Botones(i).Visible = False
        Next
        
        Dim text As String
        text = "Este ranking se reiniciará todos los lunes a las 5:00 a.m, la reputación de todos los usuarios volverá a 0. Los usuarios que salgan en el top 3 cada semana obtendrán grandes premios!"
        With Informacion
            .SelStart = 0
            .SelLength = 0
            
            .SelBold = False
            .SelItalic = False
            
            .SelColor = RGB(255, 255, 255)
            
            .SelText = IIf(False, text, text & vbCrLf)
        End With
    
        Reputacion.Visible = True
        Informacion.Visible = True
        NameRank.Caption = "SEMANAL"
        RankingActual = 2
        Call SendData("RANKIN" & 9)
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_Main_Repu.jpg")
    Else
        For i = 0 To 8
            Botones(i).Visible = True
        Next
        
        Dim j As Long
        For j = 1 To 10
            Nombre(j).Caption = ""
            Puntaje(j).Caption = ""
        Next j
        
        Reputacion.Visible = False
        Informacion.Visible = False
        NameRank.Caption = "GENERAL"
        RankingActual = 1
        Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_Main.jpg")
    End If

End Sub
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

    Dim j As Long
    For j = 1 To 10
        Nombre(j).Caption = ""
        Puntaje(j).Caption = ""
    Next j
    
    Reputacion.Visible = False
    Informacion.Visible = False
    RankingActual = 1

    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_Main.jpg")
    Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsN.jpg")
    Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosN.jpg")
    Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasN.jpg")
    Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasN.jpg")
    Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoN.jpg")
    Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_EventosN.jpg")
    Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcN.jpg")
    Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosN.jpg")
    Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesN.jpg")
End Sub
Private Sub Image2_Click()
    Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsN.jpg")
    Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosN.jpg")
    Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasN.jpg")
    Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasN.jpg")
    Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoN.jpg")
    Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_EventosN.jpg")
    Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcN.jpg")
    Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosN.jpg")
    Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesN.jpg")
End Sub
Private Sub Botones_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsI.jpg")
    If Index = 1 Then Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosI.jpg")
    If Index = 2 Then Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasI.jpg")
    If Index = 3 Then Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasI.jpg")
    If Index = 4 Then Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoI.jpg")
    If Index = 5 Then Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_EventosI.jpg")
    If Index = 6 Then Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcI.jpg")
    If Index = 7 Then Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosI.jpg")
    If Index = 8 Then Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesI.jpg")
End Sub
Private Sub Botones_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Index = 0 Then Botones(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_FragsA.jpg")
    If Index = 1 Then Botones(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_DuelosA.jpg")
    If Index = 2 Then Botones(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RondasA.jpg")
    If Index = 3 Then Botones(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_ParejasA.jpg")
    If Index = 4 Then Botones(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_TorneoA.jpg")
    If Index = 5 Then Botones(5).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_EventosA.jpg")
    If Index = 6 Then Botones(6).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CvcA.jpg")
    If Index = 7 Then Botones(7).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_CastillosA.jpg")
    If Index = 8 Then Botones(8).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Ranking_RepuClanesA.jpg")
    
    
End Sub
Private Sub Botones_Click(Index As Integer)
    Call SendData("RANKIN" & Index)
End Sub
