VERSION 5.00
Begin VB.Form FrmNewPoll 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Encuestas"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Crear Encuesta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox NivelMinimo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "Nivel Minimo"
      Top             =   3000
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NADA"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2535
      Width           =   855
   End
   Begin VB.TextBox Opcion5 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "N/A"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NADA"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   2055
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NADA"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1575
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "NADA"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1095
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NADA"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   615
      Width           =   855
   End
   Begin VB.TextBox Opcion2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Opcion 2"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Opcion3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "N/A"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Opcion4 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "N/A"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Opcion1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Opcion 1"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Encuesta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Titulo de la encuesta"
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FrmNewPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Opcion1.text = "N/A"
End Sub

Private Sub Command2_Click()
Opcion2.text = "N/A"
End Sub

Private Sub Command3_Click()
Opcion3.text = "N/A"
End Sub

Private Sub Command4_Click()
Opcion4.text = "N/A"
End Sub

Private Sub Command5_Click()
Opcion5.text = "N/A"
End Sub
Private Sub Command6_Click()
    'Creamos la encuesta'
    Call SendData("/ENCUESTA " & Encuesta.text & "@" & Opcion1.text & "@" & Opcion2.text & "@" & Opcion3.text & "@" & Opcion4.text & "@" & Opcion5.text & "@" & NivelMinimo.text)
    Unload Me
    'Creamos la encuesta'
End Sub
Private Sub Command7_Click()
    Unload Me ' cerramos
End Sub

