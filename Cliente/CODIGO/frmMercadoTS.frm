VERSION 5.00
Begin VB.Form frmMercadoTS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMercadoTS.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1155
      ScaleHeight     =   1890
      ScaleWidth      =   1905
      TabIndex        =   4
      Top             =   3090
      Width           =   1935
   End
   Begin VB.ListBox lstPacks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2340
      IntegralHeight  =   0   'False
      Left            =   480
      TabIndex        =   2
      Top             =   615
      Width           =   5775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   3720
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4185
      TabIndex        =   3
      Top             =   4455
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   6360
      Top             =   120
      Width           =   255
   End
   Begin VB.Image lblPurchase 
      Height          =   525
      Left            =   2100
      Picture         =   "frmMercadoTS.frx":12247
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label lblContent 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1170
      Left            =   3360
      TabIndex        =   1
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblTSPoints 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   4185
      TabIndex        =   0
      Top             =   4770
      Width           =   915
   End
End
Attribute VB_Name = "frmMercadoTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private EleccionDonacion As String
Private ItemYaSeleccionado As Byte
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Call SendData("DPX" & 1)

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Donacion_Main.jpg")
lblPurchase.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\DonacionCanjear_N.jpg")
End Sub
Private Sub lblPurchase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPurchase.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\DonacionCanjear_I.jpg")
End Sub
Private Sub lblPurchase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPurchase.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\DonacionCanjear_A.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPurchase.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\DonacionCanjear_N.jpg")
End Sub
Private Sub Image1_Click()
    Timer1.Enabled = False
    Unload Me
End Sub
Private Sub lblPurchase_Click()

        If MsgBox("¿Estás seguro que deseas canjear " & lstPacks.text & "?", vbYesNo) = vbYes Then
            Call SendData("DRX" & lstPacks.ListIndex + 1)
        End If
    
End Sub
Private Sub lstPacks_Click()
    Call SendData("DPX" & lstPacks.ListIndex + 1)
End Sub

Private Sub Timer1_Timer()
    Call engine.DrawDonations
End Sub
