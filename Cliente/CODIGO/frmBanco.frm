VERSION 5.00
Begin VB.Form frmBanco 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   505
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   2
      Left            =   315
      MousePointer    =   99  'Custom
      Top             =   2205
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   1
      Left            =   315
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   0
      Left            =   315
      MousePointer    =   4  'Icon
      Top             =   915
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const InterfaceName As String = "Banco"
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
Me.Picture = General_Load_Interface_Picture(InterfaceName & "_Main.jpg")
ChangeButtonsNormal
End Sub
Private Sub Image1_Click(Index As Integer)

Dim strInput As String
Dim cantidad As Long

Select Case Index
Case 0
    Call SendData("INIBOV")
    Unload Me
Case 1
Do
strInput = InputBox("Ingresa la cantidad a depositar", "Depositar", "0")
If StrPtr(cantidad) = 0 Then Exit Sub
If Not IsNumeric(strInput) Then Exit Sub

Loop While Not IsNumeric(strInput)

cantidad = strInput

Call SendData("/DEPOSITAR " & cantidad & "")
Case 2
Do
strInput = InputBox("Ingresa la cantidad a retirar", "Retirar", "0")
If StrPtr(cantidad) = 0 Then Exit Sub
If Not IsNumeric(strInput) Then Exit Sub

Loop While Not IsNumeric(strInput)

cantidad = strInput

Call SendData("/RETIRAR " & cantidad & "")
End Select

End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Index = 0 Then Image1(0).Picture = ChangeButtonState(Apretado, "BBoveda")
If Index = 1 Then Image1(1).Picture = ChangeButtonState(Apretado, "BDepositar")
If Index = 2 Then Image1(2).Picture = ChangeButtonState(Apretado, "BRetirar")

End Sub
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
If Image1(0).Tag = "0" Then
    Call ChangeButtonsNormal
    Image1(0).Picture = ChangeButtonState(Iluminado, "BBoveda")
    Image1(0).Tag = "1"
End If
End If

If Index = 1 Then
If Image1(1).Tag = "0" Then
    Call ChangeButtonsNormal
    Image1(1).Picture = ChangeButtonState(Iluminado, "BDepositar")
    Image1(1).Tag = "1"
End If
End If

If Index = 2 Then
If Image1(2).Tag = "0" Then
    Call ChangeButtonsNormal
    Image1(2).Picture = ChangeButtonState(Iluminado, "BRetirar")
    Image1(2).Tag = "1"
End If
End If

End Sub
Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub
Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")
If Estado = Apretado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "A.jpg")

End Function
Private Sub ChangeButtonsNormal()

Image1(0).Picture = ChangeButtonState(BNormal, "BBoveda")
Image1(1).Picture = ChangeButtonState(BNormal, "BDepositar")
Image1(2).Picture = ChangeButtonState(BNormal, "BRetirar")

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
Private Sub Label1_Click()
Unload Me
End Sub
