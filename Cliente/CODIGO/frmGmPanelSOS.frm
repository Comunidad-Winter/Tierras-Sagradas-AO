VERSION 5.00
Begin VB.Form frmGmPanelSOS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SOS/Panel GM"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Modo 
      Caption         =   "SOS"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   57
      Top             =   120
      Width           =   1450
   End
   Begin VB.CommandButton Modo 
      Caption         =   "TRABAJADORES"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   56
      Top             =   120
      Width           =   1450
   End
   Begin VB.CommandButton Modo 
      Caption         =   "DENUNCIADOS"
      Height          =   495
      Index           =   4
      Left            =   3120
      TabIndex        =   55
      Top             =   120
      Width           =   1450
   End
   Begin VB.CommandButton Modo 
      Caption         =   "COMANDOS"
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   54
      Top             =   120
      Width           =   1450
   End
   Begin VB.CommandButton Modo 
      Caption         =   "EXTRAS"
      Height          =   495
      Index           =   3
      Left            =   6000
      TabIndex        =   53
      Top             =   120
      Width           =   1450
   End
   Begin VB.Frame Frame 
      Height          =   5775
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   7455
      Begin VB.ListBox lstDenuncias 
         Height          =   4740
         ItemData        =   "frmGmPanelSOS.frx":0000
         Left            =   120
         List            =   "frmGmPanelSOS.frx":0007
         TabIndex        =   51
         Top             =   480
         Width           =   2415
      End
      Begin VB.Frame frameDenunciados 
         Caption         =   "Informacion del usuario denunciado"
         Height          =   3255
         Left            =   2640
         TabIndex        =   42
         Top             =   240
         Width           =   4695
         Begin VB.ListBox lstDenunciantes 
            Height          =   1620
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   4455
         End
         Begin VB.Label lblSospecha 
            Caption         =   "Sospechoso:"
            Height          =   255
            Left            =   3000
            TabIndex        =   44
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblID 
            Caption         =   "ID: "
            Height          =   255
            Left            =   1560
            TabIndex        =   49
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblIP 
            Caption         =   "IP: "
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblUltimaDenuncia 
            Caption         =   "Ultima denuncia:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label lblUltimoLogeo 
            Caption         =   "Ultimo logeo: "
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   4455
         End
         Begin VB.Label Label13 
            Caption         =   "Denunciado por:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblPrimerDenuncia 
            Caption         =   "Primer Denuncia: "
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   4455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Acciones"
         Height          =   2175
         Left            =   2640
         TabIndex        =   35
         Top             =   3480
         Width           =   4695
         Begin VB.CommandButton cmdDCheat 
            Caption         =   "/CHEAT"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmdDChori 
            Caption         =   "/CHORI"
            Height          =   375
            Left            =   2520
            TabIndex        =   40
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmdDEspiar 
            Caption         =   "/ESPIAR"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton cmdDStop 
            Caption         =   "/STOP"
            Height          =   375
            Left            =   2520
            TabIndex        =   38
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton cmdDAmurar 
            Caption         =   "BANEAR CUENTA CON MOTIVO ""Uso de cheats"""
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   4335
         End
         Begin VB.CommandButton cmdDDelete 
            Caption         =   "BORRAR DENUNCIA"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   1680
            Width           =   4335
         End
      End
      Begin VB.CommandButton cmdActualizarDenuncias 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   1560
         TabIndex        =   34
         Top             =   5280
         Width           =   975
      End
      Begin VB.CheckBox chShowOff 
         Caption         =   "Solo mostrar onlines"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Nombre - Denuncias - Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame 
      Height          =   5775
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7455
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2400
         Width           =   3220
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1600
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   3220
      End
      Begin VB.ListBox UserSOSList 
         Height          =   4155
         ItemData        =   "frmGmPanelSOS.frx":0022
         Left            =   120
         List            =   "frmGmPanelSOS.frx":0029
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Responder"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Castigar a los NWs"
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   4440
         Width           =   5415
         Begin VB.TextBox UserNICK 
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
         Begin VB.ListBox Punishment 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            ItemData        =   "frmGmPanelSOS.frx":003A
            Left            =   120
            List            =   "frmGmPanelSOS.frx":004D
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TheTime 
            Height          =   285
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   6
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmGmPanelSOS.frx":007C
            Left            =   3240
            List            =   "frmGmPanelSOS.frx":008F
            TabIndex        =   5
            Text            =   "Mal SOS"
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Aplicar Pena"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   5175
         End
         Begin VB.Label Label3 
            Caption         =   "Nick/Cuenta:"
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Pena:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Tiempo:"
            Height          =   255
            Left            =   4680
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar Lista de SOS"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   4080
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje del Usuario:"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Respuesta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   2160
         Width           =   3255
      End
   End
   Begin VB.Frame Frame 
      Height          =   5775
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton cmdTrabajo 
         Caption         =   "Devolver a ..."
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   31
         Top             =   2640
         Width           =   3375
      End
      Begin VB.CommandButton cmdTrabajo 
         Caption         =   "Explotar a ..."
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   24
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CommandButton cmdTrabajo 
         Caption         =   "Ir a ..."
         Height          =   495
         Index           =   2
         Left            =   3960
         TabIndex        =   23
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CommandButton cmdTrabajo 
         Caption         =   "Traer a ..."
         Height          =   495
         Index           =   1
         Left            =   3960
         TabIndex        =   22
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton cmdTrabajo 
         Caption         =   "Realizar chequeo de inasistencia a ..."
         Height          =   495
         Index           =   0
         Left            =   3960
         TabIndex        =   21
         Top             =   480
         Width           =   3375
      End
      Begin VB.ListBox lstTrabajo 
         Height          =   5130
         ItemData        =   "frmGmPanelSOS.frx":00D5
         Left            =   120
         List            =   "frmGmPanelSOS.frx":00DC
         TabIndex        =   19
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblTrabajo 
         Caption         =   "Usuarios Trabajando:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame 
      Height          =   5775
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir Foro en Hoguera :p"
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   3120
         Width           =   6975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir TH de Donantes"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Width           =   6975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir Foro Staff"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   6975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir Foro Denuncias"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   6975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir Foro"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame 
      Height          =   5775
      Index           =   2
      Left            =   120
      TabIndex        =   58
      Top             =   720
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame Frame3 
         Caption         =   "Torneos"
         Height          =   855
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   7215
         Begin VB.ListBox lstTorneoTipo 
            Height          =   450
            ItemData        =   "frmGmPanelSOS.frx":00F7
            Left            =   960
            List            =   "frmGmPanelSOS.frx":0101
            TabIndex        =   69
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox lstTorneoNum 
            Height          =   255
            ItemData        =   "frmGmPanelSOS.frx":0115
            Left            =   2929
            List            =   "frmGmPanelSOS.frx":0128
            TabIndex        =   68
            Top             =   360
            Width           =   735
         End
         Begin VB.ListBox lstTorneoModo 
            Height          =   255
            ItemData        =   "frmGmPanelSOS.frx":0140
            Left            =   5880
            List            =   "frmGmPanelSOS.frx":0150
            TabIndex        =   67
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Organizar"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "para"
            Height          =   255
            Left            =   2520
            TabIndex        =   71
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "participantes con modalidad"
            Height          =   255
            Left            =   3720
            TabIndex        =   70
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "LMSG"
         Height          =   855
         Left            =   120
         TabIndex        =   64
         Top             =   1200
         Width           =   7215
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   120
            TabIndex        =   65
            Text            =   "Text3"
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Bodys"
         Height          =   975
         Left            =   120
         TabIndex        =   59
         Top             =   2160
         Width           =   1215
         Begin VB.CommandButton cmdchangehead 
            Caption         =   "+"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   63
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdchangehead 
            Caption         =   "-"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtheadnumc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   61
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdchangehead 
            Caption         =   "Poner"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   975
         End
      End
   End
   Begin VB.Menu menus 
      Caption         =   "Mnus"
      Visible         =   0   'False
      Begin VB.Menu mIr 
         Caption         =   "Ir"
      End
      Begin VB.Menu mSum 
         Caption         =   "Sum"
      End
      Begin VB.Menu mDV 
         Caption         =   "Devolver"
      End
      Begin VB.Menu mIrInvi 
         Caption         =   "Ir Invisible"
      End
      Begin VB.Menu mDel 
         Caption         =   "Borrar Consulta"
      End
   End
End
Attribute VB_Name = "frmGmPanelSOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UserNombre As String

Private Sub cmdActualizarDenuncias_Click()
Dim ocx As Long
lstDenuncias.Clear
For ocx = 1 To DenunciasNumber
    lstDenuncias.AddItem "" & Denuncias(ocx).Nick & " - " & Denuncias(ocx).Contenido & " - " & Denuncias(ocx).Estado & ""
Next ocx
End Sub

Private Sub cmdchangehead_Click(Index As Integer)

If Index = 0 Then txtheadnumc.text = txtheadnumc.text - 1
If Index = 1 Then txtheadnumc.text = txtheadnumc.text + 1
If Index = 2 Then Call SendData("/MOD BODY " & txtheadnumc.text)

End Sub

Private Sub cmdDAmurar_Click()
Call SendData("/BANACC " & Denuncias(lstDenuncias.ListIndex + 1).Nick)
End Sub

Private Sub cmdDCheat_Click()
Call SendData("/CHEAT " & Denuncias(lstDenuncias.ListIndex + 1).Nick)
End Sub

Private Sub cmdDChori_Click()
Call SendData("/CHORI " & Denuncias(lstDenuncias.ListIndex + 1).Nick)
End Sub

Private Sub cmdDDelete_Click()
If lstDenuncias.ListIndex < 0 Then Exit Sub
lstDenuncias.RemoveItem (lstDenuncias.ListIndex)

Dim i As Long
For i = lstDenuncias.ListIndex + 1 To 499
    Denuncias(i).Autor = Denuncias(i + 1).Autor
    Denuncias(i).Contenido = Denuncias(i + 1).Contenido
    Denuncias(i).ID = Denuncias(i + 1).ID
    Denuncias(i).YP = Denuncias(i + 1).YP
    Denuncias(i).Nick = Denuncias(i + 1).Nick
    Denuncias(i).UltimoLogeo = Denuncias(i + 1).UltimoLogeo
    Denuncias(i).UltimaDenuncia = Denuncias(i + 1).UltimaDenuncia
    Denuncias(i).PrimerDenuncia = Denuncias(i + 1).PrimerDenuncia
    Denuncias(i).Estado = Denuncias(i + 1).Estado
Next i

DenunciasNumber = DenunciasNumber - 1
End Sub

Private Sub cmdDEspiar_Click()
Call SendData("/ESPIAR " & Denuncias(lstDenuncias.ListIndex + 1).Nick)
End Sub

Private Sub cmdDStop_Click()
Call SendData("/STOP " & Denuncias(lstDenuncias.ListIndex + 1).Nick)
End Sub

Private Sub Command1_Click()
If Text2 = "" Then
    Mensaje.Escribir "Escribí un mensaje para el usuario."
    Exit Sub
End If

If UserSOSList.ListIndex > -1 And UserSOSList.List(UserSOSList.ListIndex) <> "" Then
Call SendData("X" & EsUsuario & "*" & Text2.text)
EsUsuario = ""
If UserSOSList.ListIndex < 0 Then Exit Sub
Call SendData("SOSDONE" & MensajesSOS(UserSOSList.ListIndex + 1).Autor & "," & UserSOSList.ListIndex + 1)
End If

End Sub

Private Sub Command2_Click()

If Punishment.ListIndex = 4 Then
    Call SendData(Punishment.List(Punishment.ListIndex) & " " & UserNICK)
    Exit Sub
End If

If Punishment.ListIndex = 2 Then
    Call SendData(Punishment.List(Punishment.ListIndex) & " " & UserNICK & "@" & Combo1.text & "@" & TheTime)
    Exit Sub
End If

If Punishment.ListIndex = 1 Then
    Call SendData(Punishment.List(Punishment.ListIndex) & " " & UserNICK & "@" & TheTime)
    Exit Sub
End If

Call SendData(Punishment.List(Punishment.ListIndex) & " " & UserNICK & "@" & Combo1.text)

End Sub

Private Sub Command3_Click()
UserSOSList.Clear
Text1 = ""
Text2 = ""
UserNICK = ""
Call SendData("CONSUL")
End Sub

Private Sub Command5_Click(Index As Integer)

If Index = 0 Then
    OpenBrowser "http://www.tierras-perdidas.com/f/index.php?", 4
ElseIf Index = 1 Then
    OpenBrowser "http://www.tierras-perdidas.com/f/forumdisplay.php?f=20", 4
ElseIf Index = 2 Then
    OpenBrowser "http://www.tierras-perdidas.com/f/forumdisplay.php?f=10", 4
ElseIf Index = 3 Then
    OpenBrowser "http://www.tierras-perdidas.com/f/showthread.php?t=23741", 4
ElseIf Index = 4 Then
    OpenBrowser "http://www.tierras-perdidas.com/f/forumdisplay.php?f=7", 4
End If

End Sub

Private Sub Form_Load()
If UserPrivilegios = 0 Then Unload Me
'Punishment.ListIndex = 0

UserSOSList.Clear
Call SendData("CONSUL")

lstDenuncias.Clear
For ocx = 1 To DenunciasNumber
    lstDenuncias.AddItem "" & Denuncias(ocx).Nick & " - " & Denuncias(ocx).Contenido & " - " & Denuncias(ocx).Estado & ""
Next ocx

End Sub
Private Sub lstDenuncias_Click()
If lstDenuncias.ListIndex > -1 And lstDenuncias.List(lstDenuncias.ListIndex) <> "" Then
    lstDenunciantes.Clear
    lblIP.Caption = "IP: " & Denuncias(lstDenuncias.ListIndex + 1).YP
    lblID.Caption = "ID: " & Denuncias(lstDenuncias.ListIndex + 1).ID
    lblPrimerDenuncia.Caption = "Primer Denuncia: " & Denuncias(lstDenuncias.ListIndex + 1).PrimerDenuncia
    lblUltimaDenuncia.Caption = "Ultima Denuncia: " & Denuncias(lstDenuncias.ListIndex + 1).UltimaDenuncia
    lblUltimoLogeo.Caption = "Ultimo Logeo: " & Denuncias(lstDenuncias.ListIndex + 1).UltimoLogeo
    lstDenunciantes.AddItem Denuncias(lstDenuncias.ListIndex + 1).Autor
    Denuncias(lstDenuncias.ListIndex + 1).Estado = "LEIDO"
Else
    lblIP.Caption = ""
    lblID.Caption = ""
    lblPrimerDenuncia.Caption = ""
    lblUltimaDenuncia.Caption = ""
    lblUltimoLogeo.Caption = ""
    lstDenunciantes.Clear
End If
End Sub

Private Sub mDel_Click()
If UserSOSList.ListIndex < 0 Then Exit Sub
Call SendData("SOSDONE" & MensajesSOS(UserSOSList.ListIndex + 1).Autor & "," & UserSOSList.ListIndex + 1)
End Sub

Private Sub mDV_Click()
Call SendData("/DV " & UserNICK.text)
End Sub

Private Sub mIr_Click()
Call SendData("/IRA " & UserNICK.text)
End Sub

Private Sub mIrInvi_Click()
Call SendData("/IRCERCA " & UserNICK.text)
End Sub

Private Sub Modo_Click(Index As Integer)

Frame(0).Visible = False
Frame(1).Visible = False
Frame(2).Visible = False
Frame(3).Visible = False
Frame(4).Visible = False

Frame(Index).Visible = True

If Index = 0 Then
ElseIf Index = 1 Then
    Frame(Index).Visible = True
ElseIf Index = 2 Then
End If

End Sub
Private Sub mSum_Click()
Call SendData("/SUM " & UserNICK.text)
End Sub
Private Sub UserSOSList_Click()
If UserSOSList.ListIndex > -1 And UserSOSList.List(UserSOSList.ListIndex) <> "" Then
    Text1.text = MensajesSOS(UserSOSList.ListIndex + 1).Contenido
    UserNICK.text = MensajesSOS(UserSOSList.ListIndex + 1).Autor
    EsUsuario = MensajesSOS(UserSOSList.ListIndex + 1).Autor
Else
    UserNICK.text = ""
    Text1 = ""
    'SendDataClientPacketID.ConsultasDelete
End If
End Sub
Private Sub UserSOSList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopUpMenu menus
End Sub
