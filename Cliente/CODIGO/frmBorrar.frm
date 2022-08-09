VERSION 5.00
Begin VB.Form frmCambiarPass 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   15
   ClientTop       =   -30
   ClientWidth     =   3360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBorrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox repnewpass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox newpass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox passant 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox respuesta 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   3015
   End
   Begin VB.Label Pregunta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pregunta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Repita la Nueva Pass:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Pass:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pass Actual:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pregunta Secreta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta Secreta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   675
      Width           =   1650
   End
End
Attribute VB_Name = "frmCambiarPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************************
'****************************************CAMBIO DE PASS BY LWK*****************************************************
'******************************************************************************************************************

Private Sub Command1_Click()
If Len(newpass.Text) < 6 Then
    MsgBox "El password de la cuenta debe de tener mas de 6 caracteres.", vbCritical
    Exit Sub
End If

If newpass <> repnewpass Then
    MsgBox "Las passwords que tipeo no coinciden"
    Exit Sub
End If

If respuesta.Text = " " Then
MsgBox "No se ha detectado ninguna respuesta secreta"
Exit Sub
End If


Call SendData("REPASS" & nombrecuent & "," & pregunta.Caption & "," & respuesta & "," & passant & "," & newpass & "," & repnewpass)

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmCambiarPass.Caption = "Cambio de Password cuenta " & nombrecuent
End Sub
