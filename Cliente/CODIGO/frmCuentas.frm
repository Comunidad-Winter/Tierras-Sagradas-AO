VERSION 5.00
Begin VB.Form frmCuentas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Crear nueva cuenta:"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   200
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   200
      Left            =   480
      TabIndex        =   5
      Top             =   3880
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.TextBox Mail2 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   350
      TabIndex        =   4
      Top             =   3060
      Width           =   3060
   End
   Begin VB.TextBox Mail 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   340
      Left            =   350
      TabIndex        =   3
      Top             =   2370
      Width           =   3060
   End
   Begin VB.TextBox RePass 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   350
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1700
      Width           =   3060
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   350
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1050
      Width           =   3060
   End
   Begin VB.TextBox Cuenta 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   340
      Left            =   350
      MaxLength       =   18
      TabIndex        =   0
      Top             =   430
      Width           =   3060
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Image Cancel 
      Height          =   525
      Left            =   300
      Top             =   4850
      Width           =   1260
   End
   Begin VB.Image Siguiente 
      Height          =   525
      Left            =   2160
      Top             =   4850
      Width           =   1260
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const InterfaceName As String = "CrearCuenta"

Private Sub Image1_Click()

'Call OpenBrowser("http://" & Cliente.ForumURL & "/showthread.php?t=12659", 4)

End Sub

Private Sub Siguiente_Click()


If Pass <> RePass Then
    Mensaje.Escribir "Las passwords no coinciden"
    Exit Sub
End If

If Mail <> Mail2 Then
    Mensaje.Escribir "Los emails no coinciden"
    Exit Sub
End If

If Cuenta = "" Or Len(Cuenta) > 18 Or IsNumeric(Cuenta) Or Not AsciiValidos(Cuenta) Then
    Mensaje.Escribir "Nombre invalido"
    Exit Sub
End If

If InStr(1, Mail, "@", vbTextCompare) = 0 Or InStr(1, Mail2, "@", vbTextCompare) = 0 Then
    Mensaje.Escribir "Dirección de mail invalida."
    Exit Sub
End If

If InStr(1, Pass, "ñ") > 0 Or InStr(1, Pass, "Ñ") > 0 Then
        Mensaje.Escribir "No puedes utilizar la letra ñ en la contraseña."
    Exit Sub
End If

If Check1.Value = vbChecked Then
'   Call OpenBrowser("http://" & Cliente.ForumURL & "/register.php?", 4)
End If

If Check2.Value = vbUnchecked Then
   'Call OpenBrowser("http://" & Cliente.ForumURL & "/showthread.php?t=12659", 4)
   Mensaje.Escribir "Debes leer y aceptar el reglamento."
   Exit Sub
End If

frmPasswdSinPadrinos.Show , frmConnect
EstadoLogin = CrearAccount

End Sub

Private Sub Cancel_Click()
ShowTutorial = False
'Unload Me
frmMain.Socket1.Disconnect
Unload Me
ShowTutorial = False
End Sub

Private Sub Form_Load()

frmConnect.MousePointer = vbNormal

Me.Picture = General_Load_Interface_Picture(InterfaceName & "_Main.jpg")

ChangeButtonsNormal

MsgBox "Recorda poner un nombre de cuenta distinto del de tus personajes para mas seguridad."

End Sub


Private Sub Siguiente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Siguiente.Picture = ChangeButtonState(Apretado, "BCrear")
End Sub

Private Sub Siguiente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Siguiente.Tag = "0" Then
    Call ChangeButtonsNormal
    Siguiente.Picture = ChangeButtonState(Iluminado, "BCrear")
    Siguiente.Tag = "1"
End If
End Sub

Private Sub Siguiente_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub

Private Sub Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cancel.Picture = ChangeButtonState(Apretado, "BAtras")
End Sub

Private Sub Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Cancel.Tag = "0" Then
    Call ChangeButtonsNormal
    Cancel.Picture = ChangeButtonState(Iluminado, "BAtras")
    Cancel.Tag = "1"
End If
End Sub

Private Sub Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

Cancel.Picture = ChangeButtonState(BNormal, "BAtras")
Siguiente.Picture = ChangeButtonState(BNormal, "BCrear")

Dim j
For Each j In Me
    j.Tag = "0"
Next

Me.Tag = "0"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub
