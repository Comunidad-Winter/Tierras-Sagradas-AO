VERSION 5.00
Begin VB.Form frmMenuMascota 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SUBIRSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "CAMBIAR NOMBRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "TRANSFERIR MASCOTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "VENDER MASCOTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "QUITAR MASCOTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[SALIR]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   1785
   End
End
Attribute VB_Name = "frmMenuMascota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'If Configuracion.Alpha_Interfaz_Transparencia > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim nName As String
Select Case Index
    Case 0
        nName = InputBox("Elegi un nuevo nombre para tu mascota", "Cambio de nombre")
        If Len(nName) > 15 Then
            Mensaje.Escribir "El tamaño maximo del nombre es 15 letras"
            Exit Sub
        End If
        If nName = "" Or Len(nName) > 18 Or IsNumeric(nName) Or Not AsciiValidos(nName) Then
            Mensaje.Escribir "Nombre invalido"
            Exit Sub
        End If
        Call SendData("CNM" & nName)
    Case 5
        Call SendData("/MONTAR")
    Case 1
     '   nName = InputBox("Escribi el nombre del otro usuario", "Transferencia de mascota")
     '   If Len(nName) > 15 Then
     '       Mensaje.Escribir "El tamaño maximo del nombre es 15 letras"
     '       Exit Sub
     '   End If
     '   Call SendData("/MONT " & nName)
    Case 3
        Call SendData("/QUITARMASCOTA")
    Case 4
        Unload Me
    Case 2
     '   nName = InputBox("Escribi el nombre del otro usuario", "Venta de mascota")
     '
     '   If Len(nName) > 15 Then
     '       Mensaje.Escribir "El tamaño maximo del nombre es 15 letras"
     '       Exit Sub
     '   End If
     '
     '   Dim Gold As Long
     '   Gold = InputBox("Escribi la cantidad de oro oro a la que deseas vendersela")
     '
     '   If Gold > 0 And Name <> "" Then
      '      Call SendData("138" & Name & "," & Gold)
     '   Else
     '       If Gold > 100000000 Then
     '           Mensaje.Escribir "Maximo 100.000.000 de oro."
     '       Else
      '          Mensaje.Escribir "Campos invalidos."
      '      End If
      '  End If
End Select
Unload Me
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim loopc As Integer
For loopc = 0 To 5
    Label1(loopc).ForeColor = &HE0E0E0
Next loopc
Label1(Index).ForeColor = &HFFFF&
End Sub

