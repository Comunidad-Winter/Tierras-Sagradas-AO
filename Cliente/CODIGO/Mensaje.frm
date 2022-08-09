VERSION 5.00
Begin VB.Form Mensaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   465
      Left            =   1440
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Mensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Escribir(ByVal text As String)
Label1.Caption = text

    If frmMain.Visible = True Then
        Me.Show , frmMain
    ElseIf frmConnect.Visible = True Then
        Me.Show , frmConnect
    Else
        Me.Show
    End If
    
    
If Len(text) > 75 Then
    If Len(text) < 120 Then
        Label1.FontSize = 9
    Else
        Label1.FontSize = 8
    End If
Else
    Label1.FontSize = 12
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Mensaje.jpg")
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\OKM.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\OKM.jpg")
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\OKMApretado.jpg")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\OKMEncima.jpg")
End Sub

