VERSION 5.00
Begin VB.Form frmNoesNW 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Elegí tu alineación"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image3 
      Height          =   2100
      Left            =   420
      MousePointer    =   99  'Custom
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   2100
      Left            =   2520
      MousePointer    =   99  'Custom
      Top             =   2185
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   4600
      MousePointer    =   99  'Custom
      Top             =   2185
      Width           =   1455
   End
End
Attribute VB_Name = "frmNoesNW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const InterfaceName As String = "Enlistar"

Private Sub Form_Load()

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = General_Load_Interface_Picture("Enlistar_Main.jpg")
ChangeButtonsNormal

End Sub

Private Sub Image1_Click()

If MsgBox("¿Deseas pertenecer a la Horda?", vbYesNo) = vbYes Then
    SendData ("/HORDA")
End If

Unload Me

End Sub

Private Sub Image2_Click()

Unload Me

End Sub

Private Sub Image3_Click()

If MsgBox("¿Deseas pertenecer a la Alianza?", vbYesNo) = vbYes Then
    SendData ("/ALIANZA")
End If

Unload Me

End Sub
Private Function ChangeButtonState(ByVal Estado As eButtonStates, ByVal Name As String) As IPicture

If Estado = BNormal Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "N.jpg")
If Estado = Iluminado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "I.jpg")
If Estado = Bloqueado Then Set ChangeButtonState = General_Load_Interface_Picture(InterfaceName & "_" & Name & "B.jpg")

End Function

Private Sub ChangeButtonsNormal()

Image1.Picture = ChangeButtonState(BNormal, "BHorda")
Image2.Picture = ChangeButtonState(BNormal, "BNeutral")
Image3.Picture = ChangeButtonState(BNormal, "BAlianza")

Dim j

For Each j In Me
    j.Tag = "0"
Next

Me.Tag = "0"

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
Image1.Picture = ChangeButtonState(Apretado, "BHorda")

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Image1.Tag = "0" Then
    Call ChangeButtonsNormal
    Image1.Picture = ChangeButtonState(Iluminado, "BHorda")
    Image1.Tag = "1"
End If



End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
Image2.Picture = ChangeButtonState(Apretado, "BNeutral")

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Image2.Tag = "0" Then
    Call ChangeButtonsNormal
    Image2.Picture = ChangeButtonState(Iluminado, "BNeutral")
    Image2.Tag = "1"
End If



End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
Image3.Picture = ChangeButtonState(Apretado, "BAlianza")

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Image3.Tag = "0" Then
    Call ChangeButtonsNormal
    Image3.Picture = ChangeButtonState(Iluminado, "BAlianza")
    Image3.Tag = "1"
End If



End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ChangeButtonsNormal

End Sub

