VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClanesUsuario 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCastillo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   4
      Left            =   3360
      TabIndex        =   15
      Top             =   3120
      Width           =   1790
   End
   Begin MSComctlLib.ListView lstGuildsList 
      Height          =   2700
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nivel"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Faccion"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.TextBox CantiPuntos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Text            =   "0"
      Top             =   3030
      Width           =   1185
   End
   Begin VB.TextBox txtCastillo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1790
   End
   Begin VB.TextBox txtCastillo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   1790
   End
   Begin VB.TextBox txtCastillo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   1790
   End
   Begin VB.TextBox txtCastillo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   3
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1790
   End
   Begin MSComctlLib.ListView lstMembers 
      Height          =   2700
      Left            =   3360
      TabIndex        =   13
      Top             =   3840
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   4763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ubicacion"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.Image imgCastillos 
      Height          =   255
      Index           =   0
      Left            =   5450
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6000
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSublider2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "mago alianzita"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1890
      Width           =   1695
   End
   Begin VB.Label lblFundador 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "superminignomo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1001
      Width           =   1935
   End
   Begin VB.Label lblPuntos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "15.000"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2465
      Width           =   2458
   End
   Begin VB.Image imgAddPuntos 
      Height          =   300
      Left            =   1720
      Top             =   3015
      Width           =   1400
   End
   Begin VB.Image imgAbandonar 
      Height          =   375
      Left            =   3625
      Top             =   6650
      Width           =   2415
   End
   Begin VB.Image imgCastillos 
      Height          =   255
      Index           =   1
      Left            =   5440
      Picture         =   "f.frx":0000
      Top             =   1190
      Width           =   855
   End
   Begin VB.Image imgCastillos 
      Height          =   255
      Index           =   2
      Left            =   5450
      Picture         =   "f.frx":2A90
      Top             =   1670
      Width           =   855
   End
   Begin VB.Image imgCastillos 
      Height          =   255
      Index           =   3
      Left            =   5445
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image imgCastillos 
      Height          =   255
      Index           =   4
      Left            =   5450
      Top             =   2630
      Width           =   855
   End
   Begin VB.Image imgMasInfoDelClan 
      Height          =   375
      Left            =   480
      Top             =   6650
      Width           =   2415
   End
   Begin VB.Label lblNumMembers 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2360
      TabIndex        =   8
      Top             =   2185
      Width           =   615
   End
   Begin VB.Label lblReputacion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "362.487"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2727
      Width           =   1215
   End
   Begin VB.Label lblNivel 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2185
      Width           =   615
   End
   Begin VB.Label lblSubLider 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "pepemago"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1570
      Width           =   1575
   End
   Begin VB.Label lblLider 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "pepe el clero agitador"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1270
      Width           =   2295
   End
End
Attribute VB_Name = "frmClanesUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ParseUserInfo(ByVal Data As String)

lstMembers.ListItems.Clear
lstGuildsList.ListItems.Clear

lblPuntos.Caption = ReadField(1, Data, Asc("¬"))
lblNivel.Caption = ReadField(2, Data, Asc("¬"))
lblLider.Caption = ReadField(3, Data, Asc("¬"))

If ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
lblSubLider.Caption = ReadField(4, Data, Asc("¬"))
lblSublider2.Caption = "-"
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
lblSubLider.Caption = ReadField(5, Data, Asc("¬"))
lblSublider2.Caption = "-"
ElseIf ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
lblSubLider.Caption = "" & ReadField(4, Data, Asc("¬")) & ""
lblSublider2.Caption = "" & ReadField(5, Data, Asc("¬")) & ""
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
lblSubLider.Caption = "-"
lblSublider2.Caption = "-"
End If


txtCastillo(0).Text = ReadField(6, Data, Asc("¬"))
txtCastillo(1).Text = ReadField(7, Data, Asc("¬"))
txtCastillo(2).Text = ReadField(8, Data, Asc("¬"))
txtCastillo(3).Text = ReadField(9, Data, Asc("¬"))
lblReputacion.Caption = PonerPuntos(ReadField(10, Data, Asc("¬")))
lblFundador.Caption = ReadField(11, Data, Asc("¬"))

Dim r%, T%

r% = Val(ReadField(12, Data, Asc("¬")))

Dim ClanTemporal As String
Dim NombreClan As String
Dim FaccionClan As String
Dim NivelClan As Byte
Dim IndexK As Integer
ClanTemporal = ""
IndexK = 1

For T% = 1 To r%
    ClanTemporal = ReadField(12 + T%, Data, Asc("¬"))
    NombreClan = ReadField(1, ClanTemporal, Asc("-"))
    
    If UCase$(NombreClan) <> UCase$("cerrado" & T% & "") Then
        FaccionClan = ReadField(3, ClanTemporal, Asc("-"))
        NivelClan = ReadField(2, ClanTemporal, Asc("-"))
        
        lstGuildsList.ListItems.Add IndexK, , NombreClan
        lstGuildsList.ListItems(IndexK).ListSubItems.Add , , NivelClan
        
        If FaccionClan = 3 Then
            lstGuildsList.ListItems(IndexK).ListSubItems.Add , , "NEUTRAL"
        ElseIf FaccionClan = 4 Or FaccionClan = 5 Then
            lstGuildsList.ListItems(IndexK).ListSubItems.Add , , "ALIANZA"
        ElseIf FaccionClan = 2 Or FaccionClan = 3 Then
            lstGuildsList.ListItems(IndexK).ListSubItems.Add , , "HORDA"
        End If

        IndexK = IndexK + 1
    End If
Next T%

r% = Val(ReadField(11 + T% + 1, Data, Asc("¬")))
lblNumMembers.Caption = r%

Dim k%

Dim MiembroTemporal As String

For k% = 1 To r%
    MiembroTemporal = ReadField(11 + T% + 1 + k%, Data, Asc("¬"))
    
    lstMembers.ListItems.Add , , ReadField(1, MiembroTemporal, Asc("$"))
    lstMembers.ListItems(k%).ListSubItems.Add , , ReadField(2, MiembroTemporal, Asc("$"))
    lstMembers.ListItems(k%).ListSubItems.Add , , ReadField(3, MiembroTemporal, Asc("$"))
Next k%


Call Aplicar_Transparencia(Me.hWnd, 240)
Call Aplicar_Transparencia(lstMembers.hWnd, 100)
Call Aplicar_Transparencia(lstGuildsList.hWnd, 100)

Me.Show , frmMain

End Sub
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Main.jpg")
imgCastillos(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_ABANDONARCLAN_N.jpg")
imgAddPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_EnviarPuntos_N.jpg")
'imgSalir.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_OK_N.jpg")
imgMasInfoDelClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_INFO_N.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCastillos(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgCastillos(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_N.jpg")
imgAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_ABANDONARCLAN_N.jpg")
imgAddPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_EnviarPuntos_N.jpg")
imgMasInfoDelClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_INFO_N.jpg")
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub imgSalir_Click()
Unload Me
End Sub
Private Sub imgMasInfoDelClan_Click()
   If lstGuildsList.SelectedItem.Index <= 0 Then Exit Sub
Call SendData("CLANDETAILS" & lstGuildsList.ListItems.Item(lstGuildsList.SelectedItem.Index).Text)
End Sub
Private Sub imgAbandonar_Click()
Call SendData("/SALIRCLAN")
Unload Me
End Sub
Private Sub ImgCastillos_Click(Index As Integer)
If Index = 0 Then Call SendData("/IR 35")
If Index = 1 Then Call SendData("/IR 33")
If Index = 2 Then Call SendData("/IR 31")
If Index = 3 Then Call SendData("/IR 34")
If Index = 4 Then Call SendData("/IR 32")
End Sub
Private Sub imgAddPuntos_Click()
If Not IsNumeric(CantiPuntos.Text) Then Exit Sub

Call SendData("ADDPTS" & CantiPuntos.Text)
Call SendData("GLINFO")
End Sub
Private Sub imgMasInfoDelClan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMasInfoDelClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_INFO_I.jpg")
End Sub
Private Sub imgMasInfoDelClan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMasInfoDelClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_INFO_A.jpg")
End Sub
Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSalir.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_OK_I.jpg")
End Sub
Private Sub imgSalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSalir.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_OK_A.jpg")
End Sub
Private Sub imgAddPuntos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAddPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_EnviarPuntos_I.jpg")
End Sub
Private Sub imgAddPuntos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAddPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_EnviarPuntos_A.jpg")
End Sub
Private Sub imgAbandonar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_ABANDONARCLAN_I.jpg")
End Sub
Private Sub imgAbandonar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_ABANDONARCLAN_A.jpg")
End Sub
Private Sub ImgCastillos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then imgCastillos(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_I.jpg")
If Index = 1 Then imgCastillos(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_I.jpg")
If Index = 2 Then imgCastillos(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_I.jpg")
If Index = 3 Then imgCastillos(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_I.jpg")
If Index = 4 Then imgCastillos(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_I.jpg")
End Sub
Private Sub ImgCastillos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then imgCastillos(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_A.jpg")
If Index = 1 Then imgCastillos(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_A.jpg")
If Index = 2 Then imgCastillos(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_A.jpg")
If Index = 3 Then imgCastillos(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_A.jpg")
If Index = 4 Then imgCastillos(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_Member_Panel_Ir1_A.jpg")
End Sub
