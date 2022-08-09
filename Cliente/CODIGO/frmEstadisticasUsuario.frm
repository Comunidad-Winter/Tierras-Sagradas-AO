VERSION 5.00
Begin VB.Form frmEstadisticasUsuario 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   LinkTopic       =   "Form2"
   Picture         =   "frmEstadisticasUsuario.frx":0000
   ScaleHeight     =   4230
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblParejas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8421 jugados (67% victorias)"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   1230
      Width           =   1935
   End
   Begin VB.Label lblDuelos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "244 jugados (15% victorias)"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   830
      Width           =   1935
   End
   Begin VB.Label lblRondas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   1615
      Width           =   855
   End
   Begin VB.Label lblUsuariosMatados 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4.811"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   2370
      Width           =   615
   End
   Begin VB.Label lblMuertes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   2000
      Width           =   1335
   End
   Begin VB.Label lblQuests 
      BackStyle       =   0  'Transparent
      Caption         =   "124"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6140
      TabIndex        =   10
      Top             =   3555
      Width           =   375
   End
   Begin VB.Label lblCVCS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10.654"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   3150
      Width           =   2175
   End
   Begin VB.Label lblEventos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "22.377"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lblReputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "20.000"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   3530
      Width           =   1335
   End
   Begin VB.Label lblJerarquia 
      BackStyle       =   0  'Transparent
      Caption         =   "4ta Jerarquia"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3150
      Width           =   1335
   End
   Begin VB.Label lblFaccion 
      BackStyle       =   0  'Transparent
      Caption         =   "HORDA"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2790
      Width           =   1815
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999999999/999999999"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2370
      Width           =   2295
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50 + 20"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblRaza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblClase 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mago"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   765
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   6480
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmEstadisticasUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

lblNivel.Caption = formuEstadisticas.Nivel


If formuEstadisticas.Faccion = 1 Then
    lblFaccion.ForeColor = &H80&
    lblFaccion.Caption = "HORDA INFERNAL"
ElseIf formuEstadisticas.Faccion = 2 Then
    lblFaccion.ForeColor = &HC00000
    lblFaccion.Caption = "ALIANZA IMPERIAL"
ElseIf formuEstadisticas.Faccion = 0 Then
    lblFaccion.ForeColor = &H404040
    lblFaccion.Caption = "NEUTRAL"
End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

