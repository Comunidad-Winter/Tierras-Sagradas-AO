VERSION 5.00
Begin VB.Form frmGM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ayuda GM"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMotivo 
      Height          =   1815
      ItemData        =   "frmGM.frx":0000
      Left            =   1905
      List            =   "frmGM.frx":001C
      TabIndex        =   9
      Top             =   3090
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lstVistos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   1905
      TabIndex        =   10
      Text            =   "Escribi o selecciona de la lista"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.OptionButton Op 
      Caption         =   "Denunciar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Op 
      Caption         =   "Consulta regular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver Manual"
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtMotivo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   53
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   4920
      Width           =   4215
   End
   Begin VB.OptionButton Op 
      Caption         =   "Reportar Bug"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   1
      Top             =   1125
      Width           =   1335
   End
   Begin VB.OptionButton Op 
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGM.frx":008D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGM.frx":01AD
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4095
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'If GetIndex = 0 Then
If Op(0).value = False And Op(1).value = False And Op(2).value = False And Op(3).value = False Then
    Mensaje.Escribir "Seleccionar una opción."
    Exit Sub
End If
If Len(txtMotivo) > 250 Then
    Mensaje.Escribir "Máximo 250 caracteres."
    Exit Sub
End If

If Op(0).value = True And Len(txtMotivo) < 20 Then
    Mensaje.Escribir "Tamaño mínimo 20 caracteres, no se permiten mensajes del tipo 'GM Veni pls' y parecidos."
Exit Sub
End If

If InStr(1, txtMotivo, ",") Then
    MsgBox "Imposible mandar un mensaje GM con el signo ',' (Coma) adentro del mensaje, edita el mensaje y volve a enviarlo.", vbCritical, "Error #85"
    Exit Sub
End If

If DarIndiceElegido = -1 Then
    Mensaje.Escribir "Selecciona una opción."
    Exit Sub
End If

If Op(2).value = True And Len(txtMotivo) < 50 Then
    Mensaje.Escribir "Escribi un minimo de 50 caracteres explicando el bug."
    Exit Sub
ElseIf Op(2).value Then
        Call SendData("#" & DarIndiceElegido & "," & txtMotivo)
        Debug.Print "Mande SOS"
    Mensaje.Escribir "Gracias por reportar."
End If

If Op(1).value = True Then
    Mensaje.Escribir "Utiliza esto solo de urgencia, por ejemplo denunciar chiters o anti faccion."
    Call SendData("NEWD" & txtNombre.text & "," & lstMotivo.List(lstMotivo.ListIndex))
    'Call SendData("/DENUNCIAR " & txtMotivo)
End If

If Op(0).value = True Then
    Mensaje.Escribir "¡Tu consulta a sido enviada!"
    Call SendData("#" & DarIndiceElegido & "," & txtMotivo)
    'Debug.Print "Mande SOS"
End If

Unload Me

End Sub
Private Function DarIndiceElegido() As Integer

Dim i As Integer

For i = 0 To 3
    If Op(i).value = True Then
        DarIndiceElegido = i
        Exit Function
    End If
Next i

DarIndiceElegido = -1

End Function

Private Sub Command2_Click()
'frmAyuda.Show vbModeless, frmMain
End Sub

Private Sub Command3_Click()
'shellexecute( "http://tpao.com.ar/upload/newthread.php?do=newthread&f=20", 4
End Sub
Private Sub lstVistos_Click()
If lstVistos.ListIndex > -1 And lstVistos.List(lstVistos.ListIndex) <> "" Then
    txtNombre.text = lstVistos.List(lstVistos.ListIndex)
End If
End Sub

Private Sub Op_Click(Index As Integer)
If Index = 1 Then
    txtMotivo.Visible = False
    Label2.Visible = False
    
    lstVistos.Visible = True
    txtNombre.Visible = True
    lstMotivo.Visible = True
Else
    txtMotivo.Visible = True
    Label2.Visible = True
    
    lstVistos.Visible = False
    txtNombre.Visible = False
    lstMotivo.Visible = False
End If

End Sub

Private Sub txtMotivo_Change()
If Len(txtMotivo) = 250 Then
    Mensaje.Escribir "Tamaño maximo 250 caracteres."
End If
End Sub

Private Sub txtMotivo_Click()
If txtMotivo.text = "Tamaño máximo 250 caracteres." Then txtMotivo.text = ""
'If GetIndex = 2 Then Mensaje.Escribir "Utiliza esto para acusar posibles cheaters o actos de corrupción."
'If GetIndex = 3 Then Mensaje.Escribir "Su mensaje sera almacenado en nuestra base de datos."
'If GetIndex = 4 Then Mensaje.Escribir "Describa bien el error o bug y el programador lo tratara de solucionar lo antes posible."
'If GetIndex = 5 Then Mensaje.Escribir "No utilice esto para pedir ser del Staff."
End Sub

'Public Function GetIndex() As Byte
'If Consulta(1).value = True Then
'    GetIndex = 1
'    Exit Function
'End If
'If Consulta(2).value = True Then
'    GetIndex = 2
'    Exit Function
'End If
'If Consulta(3).value = True Then
'    GetIndex = 3
'    Exit Function
'End If
'If Consulta(4).value = True Then
'    GetIndex = 4
'    Exit Function
'End If
'If Consulta(5).value = True Then
'    GetIndex = 5
'    Exit Function
'End If
'GetIndex = 0
'End Function
