VERSION 5.00
Begin VB.Form FrmMejorar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   Picture         =   "FrmMejorar.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListaMejorados 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2870
      IntegralHeight  =   0   'False
      ItemData        =   "FrmMejorar.frx":12256
      Left            =   160
      List            =   "FrmMejorar.frx":12258
      TabIndex        =   3
      Top             =   1050
      Width           =   2760
   End
   Begin VB.PictureBox Item 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   430
      Left            =   4090
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1050
      Width           =   435
   End
   Begin VB.TextBox Desc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   990
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1620
      Width           =   3340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   4080
      Top             =   3480
      Width           =   3360
   End
   Begin VB.Label AtaqueMagico 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6660
      TabIndex        =   7
      Top             =   3100
      Width           =   735
   End
   Begin VB.Label DefensaMagica 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6660
      TabIndex        =   6
      Top             =   2820
      Width           =   735
   End
   Begin VB.Label Defensa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4710
      TabIndex        =   5
      Top             =   3100
      Width           =   735
   End
   Begin VB.Label Ataque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4710
      TabIndex        =   4
      Top             =   2820
      Width           =   735
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   1050
      Width           =   2775
   End
End
Attribute VB_Name = "FrmMejorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
If ListaMejorados.Enabled = False Or ListaMejorados.text = "" Or UCase$(ListaMejorados.text) = "SIN ITEMS MEJORABLES" Then
MsgBox "No hay ningún objeto para mejorar."
Exit Sub
End If
SendData "SPÑ" & ListaMejorados.text
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()

End Sub

Private Sub ListaMejorados_Click()
Call SendData("SPH" & ListaMejorados.text)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\Mejorar_Mejorar_I.jpg")
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\Mejorar_Mejorar_A.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\Mejorar_Mejorar_N.jpg")
End Sub

Private Sub Form_Load()
    Image1.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\Mejorar_Mejorar_N.jpg")
    Me.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\Mejorar_Main.jpg")
End Sub
