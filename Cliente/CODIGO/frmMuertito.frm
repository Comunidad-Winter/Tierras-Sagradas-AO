VERSION 5.00
Begin VB.Form frmMuertito 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Volver a la ciudad"
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar como fantasma"
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   930
      Left            =   720
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Has sido asesinado..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmMuertito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
Call SendData("/REGRESAR")
Unload Me
End Sub
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\MUERTE.jpg")
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
End Sub
