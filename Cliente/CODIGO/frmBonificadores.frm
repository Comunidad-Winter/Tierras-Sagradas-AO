VERSION 5.00
Begin VB.Form frmBonificadores 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4920
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBeneficio 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   600
      Index           =   1
      Left            =   1030
      TabIndex        =   1
      Top             =   1430
      Width           =   4200
   End
   Begin VB.Label lblBeneficio 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Index           =   0
      Left            =   1030
      TabIndex        =   0
      Top             =   600
      Width           =   4200
   End
   Begin VB.Image Bonificacion 
      Height          =   690
      Index           =   1
      Left            =   200
      Top             =   1400
      Width           =   720
   End
   Begin VB.Image Bonificacion 
      Height          =   690
      Index           =   0
      Left            =   200
      Top             =   550
      Width           =   720
   End
End
Attribute VB_Name = "frmBonificadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const InterfaceName As String = "Bonificadores"
Private Sub Bonificacion_Click(Index As Integer)

If MsgBox("¿Elegir este bonificador para tu clase?", vbYesNo) = vbYes Then
 
 If Index = 0 Then
   Call SendData("BOF" & lblBeneficio(0).Caption)
 ElseIf Index = 1 Then
   Call SendData("BOF" & lblBeneficio(1).Caption)
 End If
 
    Unload Me
    Exit Sub
End If

End Sub
Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function
Private Sub ChangeButtonsNormal()

Bonificacion(0).Picture = ChangeButtonState(BNormal, "BArriva")
Bonificacion(1).Picture = ChangeButtonState(BNormal, "BAbajo")

Dim j
For Each j In Me
    j.Tag = "0"
Next

Me.Tag = "0"

End Sub
Private Sub Bonificacion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Bonificacion(0).Picture = ChangeButtonState(Apretado, "BArriva")
If Index = 1 Then Bonificacion(1).Picture = ChangeButtonState(Apretado, "BAbajo")
End Sub
Private Sub Bonificacion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
If Bonificacion(0).Tag = "0" Then
    Call ChangeButtonsNormal
    Bonificacion(0).Picture = ChangeButtonState(Iluminado, "BArriva")
    Bonificacion(0).Tag = "1"
End If
End If

If Index = 1 Then
If Bonificacion(1).Tag = "0" Then
    Call ChangeButtonsNormal
    Bonificacion(1).Picture = ChangeButtonState(Iluminado, "BAbajo")
    Bonificacion(1).Tag = "1"
End If
End If

End Sub

Private Sub Bonificacion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeButtonsNormal
End Sub

Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = General_Load_Interface_Picture(InterfaceName & "_Main.jpg")

ChangeButtonsNormal

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Tag = "0" Then
    Call ChangeButtonsNormal
    Me.Tag = "1"
End If

End Sub
Private Sub Image1_Click()
Unload Me
End Sub
