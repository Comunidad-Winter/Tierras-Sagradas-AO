VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmMisionesDiarias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Misiones Diarias"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000001&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3120
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   5
      Top             =   2160
      Width           =   555
   End
   Begin RichTextLib.RichTextBox txtInformacion 
      Height          =   2535
      Left            =   575
      TabIndex        =   4
      Top             =   1260
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMisionesDiarias.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblComplete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1410
      Width           =   3495
   End
   Begin VB.Image BarraCompletada 
      Height          =   285
      Left            =   2690
      Picture         =   "frmMisionesDiarias.frx":0096
      Top             =   1400
      Width           =   3450
   End
   Begin VB.Label lblPTS 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   2423
      Width           =   975
   End
   Begin VB.Label lblGLD 
      BackStyle       =   0  'Transparent
      Caption         =   "100.000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4730
      TabIndex        =   1
      Top             =   2615
      Width           =   975
   End
   Begin VB.Image cmdExit 
      Height          =   495
      Left            =   5880
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdReclamar 
      Height          =   630
      Left            =   2760
      Top             =   3000
      Width           =   3270
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asesinar 10 usuarios."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "frmMisionesDiarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ParseQuest(ByVal Buffer As String)

    Me.Show , frmMain

Dim l_file As clsIniReader

    Set l_file = New clsIniReader
    l_file.Initialize App.Path & "\Data\INIT\MisionesDiarias.txt"

    Dim NroQuest As Long
    NroQuest = ReadField(1, Buffer, Asc(","))
    
    lblNombre.Caption = l_file.GetValue("MISION" & NroQuest, "Nombre")
        
        Dim text As String
        text = l_file.GetValue("MISION" & NroQuest, "Descripcion")
        With txtInformacion
            .text = ""
            .SelStart = 0
            .SelLength = 0
            
            .SelBold = False
            .SelItalic = False
            
            .SelColor = RGB(255, 255, 255)
            
            .SelText = IIf(False, text, text & vbCrLf)
        End With
        
        lblPTS.Caption = PonerPuntos(l_file.GetValue("MISION" & NroQuest, "Puntos"))
        lblGLD.Caption = PonerPuntos(l_file.GetValue("MISION" & NroQuest, "Oro"))
        
        If l_file.GetValue("MISION" & NroQuest, "OBJ") > 0 Then
            Picture1.Visible = True
            
            Dim XR As RECT
            Dim tObjIndex As Integer
            tObjIndex = l_file.GetValue("MISION" & NroQuest, "OBJ")
            XR.left = 0
            XR.top = 0
            XR.Right = 32
            XR.bottom = 32
            
            Picture1.Refresh
            Call engine.DrawGrhtoHdc(tObjIndex, XR, Picture1)
        Else
            Picture1.Visible = False
        End If

    Dim Progreso As Long
    Dim Necesario As Long
        Progreso = ReadField(2, Buffer, Asc(","))
        Necesario = ReadField(3, Buffer, Asc(","))
        
        BarraCompletada.Width = (((Progreso / 100) / (Necesario / 100)) * 3450)
        lblComplete.Caption = "" & PonerPuntos(Progreso) & " / " & PonerPuntos(Necesario)
        
        If Progreso < Necesario Then
            cmdReclamar.Enabled = False
        Else
            cmdReclamar.Enabled = True
        End If
    
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
    Set form_Moviment = New clsFormMovementManager
    form_Moviment.Initialize Me
    
    Me.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\MisionesDiarias_Main.jpg")
    cmdReclamar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\MisionesDiarias_botonN.jpg")
End Sub
Private Sub cmdReclamar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdReclamar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\MisionesDiarias_botonI.jpg")
End Sub
Private Sub cmdReclamar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdReclamar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\MisionesDiarias_botonA.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdReclamar.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\MisionesDiarias_botonN.jpg")
End Sub
Private Sub cmdReclamar_Click()
    Call SendData("{IZION")
    Unload Me
End Sub
