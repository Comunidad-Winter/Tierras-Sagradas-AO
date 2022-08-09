VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Quests Info"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picOBJ 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   4410
      Width           =   495
   End
   Begin VB.PictureBox picNPC 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   3360
      Picture         =   "frmQuestInfo.frx":15EFE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox picNPC 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2760
      Picture         =   "frmQuestInfo.frx":16B40
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   3500
      Width           =   495
   End
   Begin MSComctlLib.ListView lstQuests 
      Height          =   1100
      Left            =   350
      TabIndex        =   7
      Top             =   570
      Width           =   5275
      _ExtentX        =   9313
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483643
      BackColor       =   -2147483641
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dificultad"
         Object.Width           =   1765
      EndProperty
   End
   Begin VB.PictureBox picNPC 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2160
      Picture         =   "frmQuestInfo.frx":17782
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   3500
      Width           =   495
   End
   Begin VB.Label lblCantOBJ 
      BackStyle       =   0  'Transparent
      Caption         =   "x5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblPuntos 
      BackStyle       =   0  'Transparent
      Caption         =   "1.200"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2020
      TabIndex        =   13
      Top             =   4635
      Width           =   1815
   End
   Begin VB.Label lblOro 
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   850
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblCantNPC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   11
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCantNPC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCantNPC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblDrops 
      BackStyle       =   0  'Transparent
      Caption         =   "Gema Violeta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1755
      TabIndex        =   4
      Top             =   2580
      Width           =   3855
   End
   Begin VB.Label lblMapa 
      BackStyle       =   0  'Transparent
      Caption         =   "30 - 133"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2780
      Width           =   1815
   End
   Begin VB.Label lblNPCs 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2390
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1620
      TabIndex        =   1
      Top             =   2190
      Width           =   1815
   End
   Begin VB.Label lblDificultad 
      BackStyle       =   0  'Transparent
      Caption         =   "Fácil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   5520
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   1680
      Top             =   5160
      Width           =   2655
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private form_Mov As clsFormMovementManager
Private NroQuest As Integer
Public Sub CargarList()
    
    Dim j As Long
    
    lstQuests.ListItems.Clear

    For j = 1 To UBound(InfoQuests)
        lstQuests.ListItems.Add j, , InfoQuests(j).Nombre
        
        lstQuests.ListItems(j).ListSubItems.Add , , InfoQuests(j).Dificultad
        lstQuests.ListItems(j).ListSubItems.Item(1).bold = True
    Next j
    
    Me.Show , frmMain

End Sub
Private Sub Image1_Click()
Call SendData("ACQT" & NroQuest)
Unload Me
End Sub
Private Sub Image2_Click()
Unload Me
End Sub
Private Sub Form_Load()

Set form_Mov = New clsFormMovementManager
form_Mov.Initialize frmQuestInfo

lblDificultad.Visible = False
lblNivel.Visible = False
lblNPCs.Visible = False
lblMapa.Visible = False
lblDrops.Visible = False
lblOro.Visible = False
lblPuntos.Visible = False

picOBJ.Visible = False
lblCantOBJ.Visible = False

Dim g As Long
For g = 0 To 2
    picNPC(g).Visible = False
    lblCantNPC(g).Visible = False
    
    picNPC(g).top = 3500
    lblCantNPC(g).top = 3970
Next g

Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\MisionMain_N.jpg")

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\MisionMain.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\MisionMain_N.jpg")
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\MisionMain_A.jpg")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\MisionMain_I.jpg")
End Sub
Private Sub lstQuests_ItemClick(ByVal Item As MSComctlLib.ListItem)

lblDificultad.Visible = True
lblNivel.Visible = True
lblNPCs.Visible = True
lblMapa.Visible = True
lblDrops.Visible = True
lblOro.Visible = True
lblPuntos.Visible = True

    lblDificultad.Caption = InfoQuests(lstQuests.SelectedItem.Index).Dificultad
    lblNivel.Caption = InfoQuests(lstQuests.SelectedItem.Index).NivelMinimo
    lblNPCs.Caption = InfoQuests(lstQuests.SelectedItem.Index).CantNPC(1) + InfoQuests(lstQuests.SelectedItem.Index).CantNPC(2) + InfoQuests(lstQuests.SelectedItem.Index).CantNPC(3)
    lblMapa.Caption = InfoQuests(lstQuests.SelectedItem.Index).Mapas
    lblDrops.Caption = InfoQuests(lstQuests.SelectedItem.Index).PosiblesDrops
    lblOro.Caption = PonerPuntos(InfoQuests(lstQuests.SelectedItem.Index).Oro)
    lblPuntos.Caption = PonerPuntos(InfoQuests(lstQuests.SelectedItem.Index).puntos)
    
    'invisibles
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    
    If InfoQuests(lstQuests.SelectedItem.Index).NPCs = 1 Then
        picNPC(0).Visible = True
        lblCantNPC(0).Visible = True
        picNPC(0).left = 2760
        lblCantNPC(0).left = 2760
        
        picNPC(0).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(lstQuests.SelectedItem.Index).NumNPC(1) & ".jpg")
        lblCantNPC(0).Caption = "x" & InfoQuests(lstQuests.SelectedItem.Index).CantNPC(1)
    ElseIf InfoQuests(lstQuests.SelectedItem.Index).NPCs = 2 Then
        picNPC(0).Visible = True
        lblCantNPC(0).Visible = True
        picNPC(1).Visible = True
        lblCantNPC(1).Visible = True
        
        picNPC(0).left = 2400
        lblCantNPC(0).left = 2400
        picNPC(1).left = 3120
        lblCantNPC(1).left = 3120
        
        picNPC(0).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(lstQuests.SelectedItem.Index).NumNPC(1) & ".jpg")
        lblCantNPC(0).Caption = "x" & InfoQuests(lstQuests.SelectedItem.Index).CantNPC(1)
        
        picNPC(1).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(lstQuests.SelectedItem.Index).NumNPC(2) & ".jpg")
        lblCantNPC(1).Caption = "x" & InfoQuests(lstQuests.SelectedItem.Index).CantNPC(2)
    ElseIf InfoQuests(lstQuests.SelectedItem.Index).NPCs = 3 Then
        picNPC(0).Visible = True
        lblCantNPC(0).Visible = True
        picNPC(1).Visible = True
        lblCantNPC(1).Visible = True
        picNPC(2).Visible = True
        lblCantNPC(2).Visible = True
        
        picNPC(0).left = 2160
        lblCantNPC(0).left = 2160
        picNPC(1).left = 2760
        lblCantNPC(1).left = 2760
        picNPC(2).left = 3360
        lblCantNPC(2).left = 3360
        
        picNPC(0).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(lstQuests.SelectedItem.Index).NumNPC(1) & ".jpg")
        lblCantNPC(0).Caption = "x" & InfoQuests(lstQuests.SelectedItem.Index).CantNPC(1)
        
        picNPC(1).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(lstQuests.SelectedItem.Index).NumNPC(2) & ".jpg")
        lblCantNPC(1).Caption = "x" & InfoQuests(lstQuests.SelectedItem.Index).CantNPC(2)
        
        picNPC(2).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(lstQuests.SelectedItem.Index).NumNPC(3) & ".jpg")
        lblCantNPC(2).Caption = "x" & InfoQuests(lstQuests.SelectedItem.Index).CantNPC(3)
    End If
    
    If InfoQuests(lstQuests.SelectedItem.Index).IndexOBJ > 0 Then
        Dim SR As RECT
        SR.left = 0
        SR.top = 0
        SR.Right = 32
        SR.bottom = 32
        
        picOBJ.Visible = True
        lblCantOBJ.Visible = True
    
        picOBJ.Refresh
        Call engine.DrawGrhtoHdc(InfoQuests(lstQuests.SelectedItem.Index).IndexOBJ, SR, picOBJ)
        lblCantOBJ.Caption = PonerPuntos(InfoQuests(lstQuests.SelectedItem.Index).CantOBJ)
    Else
        picOBJ.Visible = False
    End If
    
    NroQuest = lstQuests.SelectedItem.Index
End Sub
