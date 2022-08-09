VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClanes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1075
      Left            =   240
      TabIndex        =   20
      Text            =   "Text9"
      Top             =   3885
      Width           =   9700
   End
   Begin VB.TextBox txtWeb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   2760
      TabIndex        =   19
      Text            =   "Text10"
      Top             =   3255
      Width           =   4520
   End
   Begin VB.TextBox txtNoticias 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   270
      TabIndex        =   18
      Text            =   "Text11"
      Top             =   5370
      Width           =   9585
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   5160
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2580
      Width           =   4650
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2220
      Width           =   4650
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1860
      Width           =   4650
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1515
      Width           =   4650
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   270
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2580
      Width           =   4750
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2220
      Width           =   4750
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1860
      Width           =   4750
   End
   Begin VB.TextBox txtCodex 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1515
      Width           =   4750
   End
   Begin MSComctlLib.ListView Members 
      Height          =   5775
      Left            =   330
      TabIndex        =   26
      Top             =   1275
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reputacion"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Raza"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Clase"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Oro"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Alianzas matados"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Neutrales matados"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Hordas matados"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Ultima vez online"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSComctlLib.ListView lstGuildList 
      Height          =   3062
      Left            =   360
      TabIndex        =   27
      Top             =   3165
      Width           =   9393
      _ExtentX        =   16563
      _ExtentY        =   5398
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
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Faccion"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView lstSolicitudes 
      Height          =   2775
      Left            =   225
      TabIndex        =   28
      Top             =   4365
      Width           =   7915
      _ExtentX        =   13970
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nivel"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reputacion"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Raza"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Clase"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Oro"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Alianzas matados"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Neutrales matados"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Hordas matados"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Ultima vez online"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label lblCastillos 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   5295
      TabIndex        =   25
      Top             =   3660
      Width           =   2625
   End
   Begin VB.Image imgCastillo 
      Height          =   375
      Index           =   4
      Left            =   8355
      Picture         =   "$.frx":0000
      Top             =   3540
      Width           =   1575
   End
   Begin VB.Label lblCastles 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCVCLosses 
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCVCWins 
      BackStyle       =   0  'Transparent
      Caption         =   "178"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgVerDetalles 
      Height          =   495
      Left            =   435
      Top             =   6405
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.Image imgPuntos 
      Height          =   345
      Left            =   3585
      Picture         =   "$.frx":2EEB
      Top             =   1545
      Width           =   1425
   End
   Begin VB.Image imgSubirLvl 
      Height          =   360
      Left            =   3870
      Picture         =   "$.frx":60BA
      Top             =   1905
      Width           =   1140
   End
   Begin VB.Label lblPuntos 
      BackStyle       =   0  'Transparent
      Caption         =   "1.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image imgCerrarClan 
      Height          =   735
      Left            =   8370
      Top             =   6630
      Width           =   4575
   End
   Begin VB.Image imgInfo 
      Height          =   375
      Left            =   240
      Top             =   6630
      Width           =   8055
   End
   Begin VB.Image imgModo 
      Height          =   495
      Index           =   3
      Left            =   7680
      Top             =   645
      Width           =   2205
   End
   Begin VB.Image imgModo 
      Height          =   495
      Index           =   2
      Left            =   5235
      Top             =   645
      Width           =   2205
   End
   Begin VB.Image imgModo 
      Height          =   495
      Index           =   1
      Left            =   2685
      Top             =   645
      Width           =   2205
   End
   Begin VB.Image imgModo 
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   645
      Width           =   2205
   End
   Begin VB.Label lblRep 
      BackStyle       =   0  'Transparent
      Caption         =   "-5.837"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   3630
      Width           =   855
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1830
      TabIndex        =   8
      Top             =   1890
      Width           =   375
   End
   Begin VB.Label lblMaxMiembros 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2250
      Width           =   375
   End
   Begin VB.Label lblMiembros 
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2610
      Width           =   495
   End
   Begin VB.Label lblLider 
      BackStyle       =   0  'Transparent
      Caption         =   "Magoxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2970
      Width           =   2175
   End
   Begin VB.Label lblSubLideres 
      BackStyle       =   0  'Transparent
      Caption         =   "CleroPowa y xxMagoAlianzaxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   3315
      Width           =   3255
   End
   Begin VB.Label lblCastillos 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5295
      TabIndex        =   3
      Top             =   1740
      Width           =   2865
   End
   Begin VB.Label lblCastillos 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5295
      TabIndex        =   2
      Top             =   2220
      Width           =   2870
   End
   Begin VB.Label lblCastillos 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5295
      TabIndex        =   1
      Top             =   2700
      Width           =   2870
   End
   Begin VB.Label lblCastillos 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5295
      TabIndex        =   0
      Top             =   3180
      Width           =   2865
   End
   Begin VB.Image imgCastillo 
      Height          =   375
      Index           =   0
      Left            =   8355
      Picture         =   "$.frx":9C07
      Top             =   1620
      Width           =   1575
   End
   Begin VB.Image imgCastillo 
      Height          =   375
      Index           =   1
      Left            =   8355
      Picture         =   "$.frx":CAF2
      Top             =   2100
      Width           =   1575
   End
   Begin VB.Image imgCastillo 
      Height          =   375
      Index           =   2
      Left            =   8355
      Picture         =   "$.frx":F9DD
      Top             =   2580
      Width           =   1575
   End
   Begin VB.Image imgCastillo 
      Height          =   375
      Index           =   3
      Left            =   8355
      Picture         =   "$.frx":128C8
      Top             =   3060
      Width           =   1575
   End
   Begin VB.Image imgAceptar 
      Height          =   855
      Left            =   8220
      Top             =   4350
      Width           =   1695
   End
   Begin VB.Image imgBorrar 
      Height          =   855
      Left            =   8220
      Top             =   5310
      Width           =   1695
   End
   Begin VB.Image imgRechazar 
      Height          =   855
      Left            =   8220
      Top             =   6270
      Width           =   1695
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   9600
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu mnuLider 
      Caption         =   "Lider"
      Visible         =   0   'False
      Begin VB.Menu mnuExpulsar 
         Caption         =   "Expulsar"
      End
      Begin VB.Menu mnuBoveda 
         Caption         =   "Permisos de Boveda"
         Begin VB.Menu mnuDep 
            Caption         =   "Solo depositar"
         End
         Begin VB.Menu mnuObjs 
            Caption         =   "Permitir retirar objetos"
         End
         Begin VB.Menu mnuGld 
            Caption         =   "Permitir retirar oro"
         End
         Begin VB.Menu mnuFull 
            Caption         =   "Permitir retirar objetos y oro"
         End
      End
      Begin VB.Menu mnus 
         Caption         =   "Sub Liderazgo"
         Begin VB.Menu mnusubhacer 
            Caption         =   "Hacer Sub Lider"
         End
         Begin VB.Menu mnusubsacar 
            Caption         =   "Sacar Sub Lider"
         End
      End
      Begin VB.Menu mnudolider 
         Caption         =   "Pasar Liderazgo"
      End
      Begin VB.Menu mnucontacts 
         Caption         =   "Agregar a contactos"
      End
      Begin VB.Menu mnumensaje 
         Caption         =   "Enviar un mensaje"
      End
   End
End
Attribute VB_Name = "frmClanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ParseLeaderInfo(ByVal Data As String)

Members.ListItems.Clear
lstSolicitudes.ListItems.Clear

lblPuntos.Caption = ReadField(1, Data, Asc("¬"))
lblNivel.Caption = ReadField(2, Data, Asc("¬"))
lblLider.Caption = ReadField(3, Data, Asc("¬"))

If ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
lblSubLideres.Caption = ReadField(4, Data, Asc("¬"))
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
lblSubLideres.Caption = ReadField(5, Data, Asc("¬"))
ElseIf ReadField(4, Data, Asc("¬")) <> "Fermin" And ReadField(5, Data, Asc("¬")) <> "Fermin" Then
lblSubLideres.Caption = "" & ReadField(4, Data, Asc("¬")) & " y " & ReadField(5, Data, Asc("¬")) & ""
ElseIf ReadField(4, Data, Asc("¬")) = "Fermin" And ReadField(5, Data, Asc("¬")) = "Fermin" Then
lblSubLideres.Caption = "-"
End If

Dim tX As Long

For tX = 1 To 8
    txtCodex(tX - 1).Text = ReadField(5 + tX, Data, Asc("¬"))
Next tX

lblCastillos(0).Caption = ReadField(14, Data, Asc("¬"))
lblCastillos(1).Caption = ReadField(15, Data, Asc("¬"))
lblCastillos(2).Caption = ReadField(16, Data, Asc("¬"))
lblCastillos(3).Caption = ReadField(17, Data, Asc("¬"))
'lblCastillos(4).Caption = ReadField(18, Data, Asc("¬"))

txtWeb.Text = ReadField(19, Data, Asc("¬"))
lblCVCWins.Caption = ReadField(20, Data, Asc("¬"))
lblCVCLosses.Caption = ReadField(21, Data, Asc("¬"))
lblCastles.Caption = ReadField(22, Data, Asc("¬"))
lblRep.Caption = PonerPuntos(ReadField(23, Data, Asc("¬")))

Dim des As String
des = ReadField(18, Data, Asc("¬"))
txtInfo.Text = Replace(des, "º", vbCrLf)

Dim r%, T%

r% = Val(ReadField(24, Data, Asc("¬")))

Dim ClanTemporal As String
Dim NombreClan As String
Dim FaccionClan As String
Dim NivelClan As Byte
Dim IndexK As Integer

lstGuildList.ListItems.Clear
ClanTemporal = ""
IndexK = 1

For T% = 1 To r%
    ClanTemporal = ReadField(24 + T%, Data, Asc("¬"))
    NombreClan = ReadField(1, ClanTemporal, Asc("$"))
    
    If UCase$(NombreClan) <> UCase$("cerrado" & T% & "") Then
        FaccionClan = ReadField(2, ClanTemporal, Asc("$"))
        NivelClan = ReadField(3, ClanTemporal, Asc("$"))
        
        lstGuildList.ListItems.Add IndexK, , NombreClan
        
        If FaccionClan = 3 Then
            lstGuildList.ListItems(IndexK).ListSubItems.Add , , "NEUTRAL"
        ElseIf FaccionClan = 4 Or FaccionClan = 5 Then
            lstGuildList.ListItems(IndexK).ListSubItems.Add , , "ALIANZA"
        ElseIf FaccionClan = 2 Or FaccionClan = 3 Then
            lstGuildList.ListItems(IndexK).ListSubItems.Add , , "HORDA"
        End If
        
        lstGuildList.ListItems(IndexK).ListSubItems.Add , , NivelClan
        IndexK = IndexK + 1
    End If
Next T%

r% = Val(ReadField(23 + T% + 1, Data, Asc("¬")))
lblMiembros.Caption = r%

If lblNivel.Caption = "1" Then
lblMaxMiembros.Caption = "4"
ElseIf lblNivel.Caption = "2" Then
lblMaxMiembros.Caption = "8"
ElseIf lblNivel.Caption = "3" Then
lblMaxMiembros.Caption = "12"
ElseIf lblNivel.Caption = "4" Then
lblMaxMiembros.Caption = "16"
ElseIf lblNivel.Caption = "5" Then
lblMaxMiembros.Caption = "20"
ElseIf lblNivel.Caption = "6" Then
lblMaxMiembros.Caption = "24"
ElseIf lblNivel.Caption = "7" Then
lblMaxMiembros.Caption = "28"
End If

Dim k%
Dim MiembroTemporal As String

MiembroTemporal = ""

For k% = 1 To r%
    MiembroTemporal = ReadField(23 + T% + 1 + k%, Data, Asc("¬"))
    
    Members.ListItems.Add , , ReadField(1, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).bold = True
    
    'Le damos el color al nick
    If ReadField(1, MiembroTemporal, Asc("$")) = ReadField(3, Data, Asc("¬")) Then
        Members.ListItems(k%).ForeColor = vbRed
    ElseIf ReadField(1, MiembroTemporal, Asc("$")) = ReadField(4, Data, Asc("¬")) Or ReadField(1, MiembroTemporal, Asc("$")) = ReadField(5, Data, Asc("¬")) Then
        Members.ListItems(k%).ForeColor = vbBlue
    ElseIf ReadField(2, MiembroTemporal, Asc("$")) = "ON" Then
        Members.ListItems(k%).ForeColor = vbGreen
    Else
        Members.ListItems(k%).ForeColor = &HC0C0C0
    End If
    
    Members.ListItems(k%).ListSubItems.Add , , ReadField(2, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(3, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(4, MiembroTemporal, Asc("$"))
    
    If ReadField(4, MiembroTemporal, Asc("$")) < 0 Then
        Members.ListItems(k%).ListSubItems.Item(3).ForeColor = vbRed
    End If
    
    Members.ListItems(k%).ListSubItems.Add , , ReadField(5, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(6, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(7, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(8, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(9, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(10, MiembroTemporal, Asc("$"))
    Members.ListItems(k%).ListSubItems.Add , , ReadField(11, MiembroTemporal, Asc("$"))
    
Next k%

txtNoticias = Replace(ReadField(24 + T% + k%, Data, Asc("¬")), "º", vbCrLf)

T% = 24 + T% + k% + 1

r% = Val(ReadField(T%, Data, Asc("¬")))

For k% = 1 To r%
    MiembroTemporal = ReadField(T% + k%, Data, Asc("¬"))
    
    lstSolicitudes.ListItems.Add , , ReadField(1, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).bold = True
    
    'Le damos el color al nick
    If ReadField(1, MiembroTemporal, Asc("$")) = ReadField(3, Data, Asc("¬")) Then
        lstSolicitudes.ListItems(k%).ForeColor = vbRed
    ElseIf ReadField(1, MiembroTemporal, Asc("$")) = ReadField(4, Data, Asc("¬")) Or ReadField(1, MiembroTemporal, Asc("$")) = ReadField(5, Data, Asc("¬")) Then
        lstSolicitudes.ListItems(k%).ForeColor = vbBlue
    ElseIf ReadField(2, MiembroTemporal, Asc("$")) = "ON" Then
        lstSolicitudes.ListItems(k%).ForeColor = vbGreen
    Else
        lstSolicitudes.ListItems(k%).ForeColor = &HC0C0C0
    End If
    
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(2, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(3, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(4, MiembroTemporal, Asc("$"))
    
    If ReadField(4, MiembroTemporal, Asc("$")) < 0 Then
        lstSolicitudes.ListItems(k%).ListSubItems.Item(3).ForeColor = vbRed
    End If
    
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(5, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(6, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(7, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(8, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(9, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(10, MiembroTemporal, Asc("$"))
    lstSolicitudes.ListItems(k%).ListSubItems.Add , , ReadField(11, MiembroTemporal, Asc("$"))
Next k%

Call Aplicar_Transparencia(Me.hWnd, 240)
Call Aplicar_Transparencia(Members.hWnd, 100)
Call Aplicar_Transparencia(lstGuildList.hWnd, 100)

Me.Show , frmMain

End Sub
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Main.jpg")
imgModo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_InformacionGeneral_N.jpg")
imgModo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AdministrarMiembros_N.jpg")
imgModo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_EditarClan_N.jpg")
imgModo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Extra_N.jpg")
imgSubirLvl.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Main_SubirNivel_N.jpg")
imgCastillo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_A.jpg")
imgPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AgregarPuntos_N.jpg")
imgAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_ACEPTARSOLICITUD_N.jpg")
imgBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_BORRARSOLICITUD_N.jpg")
imgRechazar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_RECHAZARSOLICITUD_N.jpg")
imgInfo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_ActualizarInformacion_N.jpg")
imgCerrarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_CerrarClan_N.jpg")
imgVerDetalles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_4_VerDetalles_N.jpg")
 
'1
imgAceptar.Visible = True
imgRechazar.Visible = True
imgCastillo(0).Visible = True
imgCastillo(1).Visible = True
imgCastillo(2).Visible = True
imgCastillo(3).Visible = True
imgSubirLvl.Visible = True
imgPuntos.Visible = True
imgBorrar.Visible = True
lblNivel.Visible = True
lblPuntos.Visible = True
lblLider.Visible = True
lblCastillos(0).Visible = True
lblCastillos(1).Visible = True
lblCastillos(2).Visible = True
lblCastillos(3).Visible = True
lblCastillos(4).Visible = True
lblSubLideres.Visible = True
lstSolicitudes.Visible = True
'2
Members.Visible = False
'3
txtCodex(0).Visible = False
txtCodex(1).Visible = False
txtCodex(2).Visible = False
txtCodex(3).Visible = False
txtCodex(4).Visible = False
txtCodex(5).Visible = False
txtCodex(6).Visible = False
txtCodex(7).Visible = False
txtInfo.Visible = False
txtWeb.Visible = False
txtNoticias.Visible = False
imgCerrarClan.Visible = False
'4
imgInfo.Visible = False
lblCVCWins.Visible = False
lblCVCLosses.Visible = False
lstGuildList.Visible = False
imgVerDetalles.Visible = False

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Volvemos las imagenes a la normalidad
imgModo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_InformacionGeneral_N.jpg")
imgModo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AdministrarMiembros_N.jpg")
imgModo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_EditarClan_N.jpg")
imgModo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Extra_N.jpg")
imgSubirLvl.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Main_SubirNivel_N.jpg")
imgCastillo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgCastillo(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_N.jpg")
imgPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AgregarPuntos_N.jpg")
imgAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_ACEPTARSOLICITUD_N.jpg")
imgBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_BORRARSOLICITUD_N.jpg")
imgRechazar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_RECHAZARSOLICITUD_N.jpg")

End Sub
Private Sub ImgModo_Click(Index As Integer)

If Index = 0 Then 'información general
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Main.jpg")
'1
imgAceptar.Visible = True
imgRechazar.Visible = True
imgCastillo(0).Visible = True
imgCastillo(1).Visible = True
imgCastillo(2).Visible = True
imgCastillo(3).Visible = True
imgCastillo(4).Visible = True
imgSubirLvl.Visible = True
imgPuntos.Visible = True
imgBorrar.Visible = True
lblNivel.Visible = True
lblPuntos.Visible = True
lblLider.Visible = True
lblCastillos(0).Visible = True
lblCastillos(1).Visible = True
lblCastillos(2).Visible = True
lblCastillos(3).Visible = True
lblCastillos(4).Visible = True
lblMaxMiembros.Visible = True
lblMiembros.Visible = True
lblSubLideres.Visible = True
lstSolicitudes.Visible = True
lblRep.Visible = True
'2
Members.Visible = False
'3
txtCodex(1).Visible = False
txtCodex(2).Visible = False
txtCodex(3).Visible = False
txtCodex(4).Visible = False
txtCodex(5).Visible = False
txtCodex(6).Visible = False
txtCodex(7).Visible = False
txtCodex(0).Visible = False
txtInfo.Visible = False
txtWeb.Visible = False
txtNoticias.Visible = False
imgCerrarClan.Visible = False
'4
imgInfo.Visible = False
lblCVCWins.Visible = False
lblCVCLosses.Visible = False
lblCastles.Visible = False
lstGuildList.Visible = False
imgVerDetalles.Visible = False

ElseIf Index = 1 Then
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_2_Main.jpg")
'1
imgAceptar.Visible = False
imgRechazar.Visible = False
imgCastillo(0).Visible = False
imgCastillo(1).Visible = False
imgCastillo(2).Visible = False
imgCastillo(3).Visible = False
imgCastillo(4).Visible = False
imgSubirLvl.Visible = False
imgPuntos.Visible = False
imgBorrar.Visible = False
lblNivel.Visible = False
lblPuntos.Visible = False
lblLider.Visible = False
lblCastillos(0).Visible = False
lblCastillos(1).Visible = False
lblCastillos(2).Visible = False
lblCastillos(3).Visible = False
lblCastillos(4).Visible = False
lblMaxMiembros.Visible = False
lblMiembros.Visible = False
lblSubLideres.Visible = False
lstSolicitudes.Visible = False
lblRep.Visible = False
'2
Members.Visible = True
'3
txtCodex(1).Visible = False
txtCodex(2).Visible = False
txtCodex(3).Visible = False
txtCodex(4).Visible = False
txtCodex(5).Visible = False
txtCodex(6).Visible = False
txtCodex(7).Visible = False
txtCodex(0).Visible = False
txtInfo.Visible = False
txtWeb.Visible = False
txtNoticias.Visible = False
imgCerrarClan.Visible = False
'4
imgInfo.Visible = False
lblCVCWins.Visible = False
lblCVCLosses.Visible = False
lblCastles.Visible = False
lstGuildList.Visible = False
imgVerDetalles.Visible = False

ElseIf Index = 2 Then
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_Main.jpg")
'1
imgAceptar.Visible = False
imgRechazar.Visible = False
imgCastillo(0).Visible = False
imgCastillo(1).Visible = False
imgCastillo(2).Visible = False
imgCastillo(3).Visible = False
imgCastillo(4).Visible = False
imgSubirLvl.Visible = False
imgPuntos.Visible = False
imgBorrar.Visible = False
lblNivel.Visible = False
lblPuntos.Visible = False
lblLider.Visible = False
lblCastillos(0).Visible = False
lblCastillos(1).Visible = False
lblCastillos(2).Visible = False
lblCastillos(3).Visible = False
lblCastillos(4).Visible = False
lblMaxMiembros.Visible = False
lblMiembros.Visible = False
lblSubLideres.Visible = False
lstSolicitudes.Visible = False
lblRep.Visible = False
'2
Members.Visible = False
'3
txtCodex(1).Visible = True
txtCodex(2).Visible = True
txtCodex(3).Visible = True
txtCodex(4).Visible = True
txtCodex(5).Visible = True
txtCodex(6).Visible = True
txtCodex(7).Visible = True
txtCodex(0).Visible = True
txtInfo.Visible = True
txtWeb.Visible = True
txtNoticias.Visible = True
imgCerrarClan.Visible = True
'4
imgInfo.Visible = True
lblCVCWins.Visible = False
lblCVCLosses.Visible = False
lblCastles.Visible = False
lstGuildList.Visible = False
imgVerDetalles.Visible = False

ElseIf Index = 3 Then
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_4_Main.jpg")
'1
imgAceptar.Visible = False
imgRechazar.Visible = False
imgCastillo(0).Visible = False
imgCastillo(1).Visible = False
imgCastillo(2).Visible = False
imgCastillo(3).Visible = False
imgCastillo(4).Visible = False
imgSubirLvl.Visible = False
imgPuntos.Visible = False
imgBorrar.Visible = False
lblNivel.Visible = False
lblPuntos.Visible = False
lblLider.Visible = False
lblCastillos(0).Visible = False
lblCastillos(1).Visible = False
lblCastillos(2).Visible = False
lblCastillos(3).Visible = False
lblCastillos(4).Visible = False
lblMaxMiembros.Visible = False
lblMiembros.Visible = False
lblSubLideres.Visible = False
lstSolicitudes.Visible = False
lblRep.Visible = False
'2
Members.Visible = False
'3
txtCodex(1).Visible = False
txtCodex(2).Visible = False
txtCodex(3).Visible = False
txtCodex(4).Visible = False
txtCodex(5).Visible = False
txtCodex(6).Visible = False
txtCodex(7).Visible = False
txtCodex(0).Visible = False
txtInfo.Visible = False
txtWeb.Visible = False
txtNoticias.Visible = False
imgCerrarClan.Visible = False
'4
imgInfo.Visible = False
lblCVCWins.Visible = True
lblCVCLosses.Visible = True
lstGuildList.Visible = True
imgVerDetalles.Visible = True
lblCastles.Visible = True
End If


End Sub
Private Sub imgAceptar_Click()
If lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).Text = "" Then
Else
Call SendData("ACEPTARI" & lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).Text)
Call SendData("GLINFO")
End If
End Sub
Private Sub imgBorrar_Click()
    
Call SendData("RECHAZAR" & lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).Text & "," & Replace(Replace("asd", ",", " "), vbCrLf, " "))
lstSolicitudes.ListItems.Remove lstSolicitudes.SelectedItem.Index
    
Call SendData("GLINFO")
End Sub
Private Sub imgRechazar_Click()

Call SendData("RECHAZAR" & lstSolicitudes.ListItems.Item(lstSolicitudes.SelectedItem.Index).Text & "," & Replace(Replace("asd", ",", " "), vbCrLf, " "))
lstSolicitudes.ListItems.Remove lstSolicitudes.SelectedItem.Index
    
Call SendData("GLINFO")
End Sub
Private Sub ImgCastillo_Click(Index As Integer)
If Index = 0 Then Call SendData("/IR 33")
If Index = 1 Then Call SendData("/IR 31")
If Index = 2 Then Call SendData("/IR 34")
If Index = 3 Then Call SendData("/IR 32")
If Index = 4 Then Call SendData("/IR 35")
End Sub
Private Sub mnuExpulsar_Click()
Call SendData("ECHARCLA" & Members.ListItems.Item(Members.SelectedItem.Index).Text)
Call SendData("GLINFO")
End Sub
Private Sub mnuSubHacer_Click()
Call SendData("/SUBLIDER " & Members.ListItems.Item(Members.SelectedItem.Index).Text)
Call SendData("GLINFO")
End Sub
Private Sub mnuSubSacar_Click()
Call SendData("/QSUBLIDR " & Members.ListItems.Item(Members.SelectedItem.Index).Text)
Call SendData("GLINFO")
End Sub
Private Sub mnuDoLider_Click()
If MsgBox("¿Está seguro que desea pasarle el liderazgo a " & Members.ListItems.Item(Members.SelectedItem.Index).Text & "?", vbYesNo) = vbYes Then Call SendData("/HACLIDER " & Members.ListItems.Item(Members.SelectedItem.Index).Text)
Unload Me
End Sub
Private Sub mnucontacts_Click()
Call SendData("ADDCON" & Members.ListItems.Item(Members.SelectedItem.Index).Text)
End Sub
Private Sub mnumensaje_Click()
TheUser = Members.ListItems.Item(Members.SelectedItem.Index).Text
'frmMain.ModoHabla.Text = "\"
'frmMain.mnuNormal.Checked = False
'frmMain.mnuprivado.Checked = True
'frmMain.mnudenunciar.Checked = False
'frmMain.mnuClanes.Checked = False
'frmMain.mnuParty.Checked = False
'frmMain.mnuGritar.Checked = False
'frmMain.mnuGlo.Checked = False
'frmMain.mnuFaccion.Checked = False
End Sub
Private Sub imgCerrarClan_Click()
Dim numerito As Integer
Dim numeritox As Integer
numerito = RandomNumber(1000, 9999)
numeritox = InputBox("Escriba el siguiente numero: " & numerito & "", "Cerrar Clan")

If numeritox = numerito Then
Call SendData("/CERRARCLAN")
Unload Me
Else
Mensaje.Escribir "Escribi bien el codigo."
End If
End Sub
Private Sub imgVerDetalles_Click()
If lstGuildList.SelectedItem.Index <= 0 Then Exit Sub
Call SendData("CLANDETAILS" & lstGuildList.ListItems.Item(lstGuildList.SelectedItem.Index).Text)
End Sub
Private Sub imgPuntos_Click()
Dim CantiPuntos As String
CantiPuntos = InputBox("Cantidad de puntos a Agregar:", "Agregar Puntos")
If Not IsNumeric(CantiPuntos) Then Exit Sub
If CantiPuntos = 0 Then Exit Sub

Call SendData("ADDPTS" & CantiPuntos)
Call SendData("GLINFO")
End Sub
Private Sub imgSubirLvl_Click()

If lblNivel.Caption = "7" Then
Mensaje.Escribir "El clan es nivel máximo"
Exit Sub
End If

If lblNivel < 5 Then
    Call SendData("SUBLVL" & lblNivel)
Else
    Call SendData("SUBLVL5")
End If

Call SendData("GLINFO")
End Sub
Private Sub imgSalir_Click()
Unload Me
frmMain.SetFocus
End Sub
Private Sub ImgModo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then imgModo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_InformacionGeneral_I.jpg")
If Index = 1 Then imgModo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AdministrarMiembros_I.jpg")
If Index = 2 Then imgModo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_EditarClan_I.jpg")
If Index = 3 Then imgModo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Extra_I.jpg")
End Sub
Private Sub ImgModo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then imgModo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_InformacionGeneral_A.jpg")
If Index = 1 Then imgModo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AdministrarMiembros_A.jpg")
If Index = 2 Then imgModo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_EditarClan_A.jpg")
If Index = 3 Then imgModo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Extra_A.jpg")
End Sub
Private Sub ImgCastillo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then imgCastillo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_I.jpg")
If Index = 1 Then imgCastillo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_I.jpg")
If Index = 2 Then imgCastillo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_I.jpg")
If Index = 3 Then imgCastillo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_I.jpg")
If Index = 4 Then imgCastillo(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_I.jpg")
End Sub
Private Sub ImgCastillo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then imgCastillo(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_A.jpg")
If Index = 1 Then imgCastillo(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_A.jpg")
If Index = 2 Then imgCastillo(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_A.jpg")
If Index = 3 Then imgCastillo(3).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_A.jpg")
If Index = 4 Then imgCastillo(4).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_IR1_A.jpg")
End Sub
Private Sub imgSubirLvl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSubirLvl.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Main_SubirNivel_I.jpg")
End Sub
Private Sub imgSubirLvl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSubirLvl.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_Main_SubirNivel_A.jpg")
End Sub
Private Sub imgPuntos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AgregarPuntos_I.jpg")
End Sub
Private Sub imgPuntos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPuntos.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_AgregarPuntos_A.jpg")
End Sub
Private Sub imgAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_ACEPTARSOLICITUD_i.jpg")
End Sub
Private Sub imgAceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAceptar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_ACEPTARSOLICITUD_A.jpg")
End Sub
Private Sub imgBorrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_BORRARSOLICITUD_I.jpg")
End Sub
Private Sub imgBorrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBorrar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_BORRARSOLICITUD_A.jpg")
End Sub
Private Sub imgRechazar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRechazar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_RECHAZARSOLICITUD_I.jpg")
End Sub
Private Sub imgRechazar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRechazar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_1_RECHAZARSOLICITUD_A.jpg")
End Sub
Private Sub imgInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgInfo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_ActualizarInformacion_I.jpg")
End Sub
Private Sub imgInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgInfo.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_ActualizarInformacion_A.jpg")
End Sub
Private Sub imgCerrarClan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCerrarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_CERRARCLAN_I.jpg")
End Sub
Private Sub imgCerrarClan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCerrarClan.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_3_CERRARCLAN_A.jpg")
End Sub
Private Sub imgVerDetalles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgVerDetalles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_4_VerDetalles_I.jpg")
End Sub
Private Sub imgVerDetalles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgVerDetalles.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Guild_4_VerDetalles_A.jpg")
End Sub
Private Sub members_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    Call SendData("VLKG" & Members.ListItems.Item(Members.SelectedItem.Index).Text)
    
    If RetiraObj = 1 And RetiraOro = 1 Then
        mnuGld.Checked = False
        mnuObjs.Checked = False
        mnuDep.Checked = False
        mnuFull.Checked = True
    ElseIf RetiraObj = 0 And RetiraOro = 1 Then
        mnuGld.Checked = True
        mnuObjs.Checked = False
        mnuDep.Checked = False
        mnuFull.Checked = False
    ElseIf RetiraObj = 1 And RetiraOro = 0 Then
        mnuObjs.Checked = True
        mnuGld.Checked = False
        mnuDep.Checked = False
        mnuFull.Checked = False
    ElseIf RetiraObj = 0 And RetiraOro = 0 Then
        mnuDep.Checked = True
        mnuObjs.Checked = False
        mnuGld.Checked = False
        mnuFull.Checked = False
    End If
    
    PopUpMenu mnuLider
End If

End Sub
Private Sub mnuDep_Click()
Dim NumeriL As Byte
NumeriL = 0
Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).Text & "," & NumeriL)
End Sub
Private Sub mnuGld_Click()
NumeriL = 1
Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).Text & "," & NumeriL)
End Sub
Private Sub mnuObjs_Click()
NumeriL = 2
Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).Text & "," & NumeriL)
End Sub
Private Sub mnuFull_Click()
NumeriL = 3
Call SendData("BOVC" & Members.ListItems.Item(Members.SelectedItem.Index).Text & "," & NumeriL)
End Sub
