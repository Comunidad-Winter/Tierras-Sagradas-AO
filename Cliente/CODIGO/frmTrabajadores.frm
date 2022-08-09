VERSION 5.00
Begin VB.Form frmNobleza 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Noble"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTrabajadores.frx":0000
   ScaleHeight     =   7380
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3910
      Picture         =   "frmTrabajadores.frx":17C31
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   4100
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1170
      Picture         =   "frmTrabajadores.frx":18475
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   4100
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3925
      Picture         =   "frmTrabajadores.frx":18CB9
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   570
      Width           =   480
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   1450
      Index           =   3
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":194FD
      Left            =   3010
      List            =   "frmTrabajadores.frx":1953D
      TabIndex        =   4
      Top             =   4750
      Width           =   2340
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   1450
      Index           =   2
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":19588
      Left            =   260
      List            =   "frmTrabajadores.frx":195C8
      TabIndex        =   3
      Top             =   4750
      Width           =   2340
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   1450
      Index           =   1
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":19613
      Left            =   3020
      List            =   "frmTrabajadores.frx":19653
      TabIndex        =   2
      Top             =   1200
      Width           =   2340
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1180
      Picture         =   "frmTrabajadores.frx":1969E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   570
      Width           =   480
   End
   Begin VB.ListBox lstReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   1450
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmTrabajadores.frx":19EE2
      Left            =   260
      List            =   "frmTrabajadores.frx":19F22
      TabIndex        =   0
      Top             =   1200
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5280
      Top             =   0
      Width           =   375
   End
   Begin VB.Image cmdConstruir 
      Height          =   435
      Index           =   3
      Left            =   3010
      Top             =   6300
      Width           =   2340
   End
   Begin VB.Image cmdConstruir 
      Height          =   435
      Index           =   2
      Left            =   260
      Top             =   6300
      Width           =   2340
   End
   Begin VB.Image cmdConstruir 
      Height          =   435
      Index           =   1
      Left            =   3020
      Top             =   2760
      Width           =   2340
   End
   Begin VB.Image cmdConstruir 
      Height          =   435
      Index           =   0
      Left            =   260
      Top             =   2760
      Width           =   2340
   End
End
Attribute VB_Name = "frmNobleza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const InterfaceName As String = "Nobleza"
Private Sub cmdConstruir_Click(Index As Integer)

'SendData ClientPacketID.NobleConstruirItem & SeparatorASCII & (Index + 1)

If Index = 0 Then Call SendData("/ITEMNOBLE DIADEMA")
If Index = 1 Then Call SendData("/ITEMNOBLE ARMADURA")
If Index = 2 Then Call SendData("/ITEMNOBLE ESPADA")
If Index = 3 Then Call SendData("/ITEMNOBLE ANILLO")

Unload Me

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

'If Configuracion.Alpha_Interfaz_Transparencia > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia

Me.Picture = General_Load_Interface_Picture(InterfaceName & "_Main.jpg")

ChangeButtonsNormal

lstReq(0).Clear
lstReq(1).Clear
lstReq(2).Clear
lstReq(3).Clear

End Sub

Private Sub cmdConstruir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Index = 0 Then cmdConstruir(0).Picture = ChangeButtonState(Apretado, "BConstruir")
If Index = 1 Then cmdConstruir(1).Picture = ChangeButtonState(Apretado, "BConstruir2")
If Index = 2 Then cmdConstruir(2).Picture = ChangeButtonState(Apretado, "BConstruir4")
If Index = 3 Then cmdConstruir(3).Picture = ChangeButtonState(Apretado, "BConstruir3")

End Sub

Private Sub cmdConstruir_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
If cmdConstruir(0).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(0).Picture = ChangeButtonState(Iluminado, "BConstruir")
    cmdConstruir(0).Tag = "1"
End If
End If

If Index = 1 Then
If cmdConstruir(1).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(1).Picture = ChangeButtonState(Iluminado, "BConstruir2")
    cmdConstruir(1).Tag = "1"
End If
End If

If Index = 2 Then
If cmdConstruir(2).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(2).Picture = ChangeButtonState(Iluminado, "BConstruir4")
    cmdConstruir(2).Tag = "1"
End If
End If

If Index = 3 Then
If cmdConstruir(3).Tag = "0" Then
    Call ChangeButtonsNormal
    cmdConstruir(3).Picture = ChangeButtonState(Iluminado, "BConstruir3")
    cmdConstruir(3).Tag = "1"
End If
End If

End Sub

Private Sub cmdConstruir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function

Private Sub ChangeButtonsNormal()

cmdConstruir(0).Picture = ChangeButtonState(BNormal, "BConstruir")
cmdConstruir(1).Picture = ChangeButtonState(BNormal, "BConstruir2")
cmdConstruir(2).Picture = ChangeButtonState(BNormal, "BConstruir4")
cmdConstruir(3).Picture = ChangeButtonState(BNormal, "BConstruir3")

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

Private Sub Image1_Click()

Unload Me
 
End Sub
