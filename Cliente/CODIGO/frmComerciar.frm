VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   1155
      TabIndex        =   8
      Text            =   "1"
      Top             =   6615
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000001&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   285
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   930
      Width           =   555
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3930
      Index           =   1
      Left            =   3675
      TabIndex        =   1
      Top             =   1995
      Width           =   2490
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3930
      Index           =   0
      ItemData        =   "frmComerciar.frx":0000
      Left            =   420
      List            =   "frmComerciar.frx":0002
      TabIndex        =   0
      Top             =   1995
      Width           =   2490
   End
   Begin VB.Image imgQuit 
      Height          =   375
      Left            =   6120
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click derecho para cerrar la ventana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   6900
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   2
      Left            =   3675
      MouseIcon       =   "frmComerciar.frx":0004
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6660
      Width           =   2490
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   3675
      MouseIcon       =   "frmComerciar.frx":0156
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6180
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   420
      MouseIcon       =   "frmComerciar.frx":02A8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6180
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   3
      Left            =   4155
      TabIndex        =   7
      Top             =   1335
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   4
      Left            =   4155
      TabIndex        =   6
      Top             =   1080
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Index           =   2
      Left            =   915
      TabIndex        =   5
      Top             =   1380
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Index           =   1
      Left            =   915
      TabIndex        =   4
      Top             =   1110
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   855
      Width           =   2835
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************
'*****************************
'*****      Samke       ******
'*****************************
'**************************************************
'**************************************************
'*****      SoHnsalxixon_u2@hotmail.com      ******
'**************************************************
'**************************************************

Private Todo As Byte
Private m_Interval As Integer
Private m_Number As Integer
Private m_Increment As Integer
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub cantidad_Change()
    If Val(cantidad.text) < 1 Then
        cantidad.text = 1
    End If
    
    If Val(cantidad.text) > MAX_INVENTORY_OBJS Then
        cantidad.text = 1
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
'If Configuracion.Alpha_Interfaz_Activar > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia

Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\comerciar.jpg")
m_Number = 1
m_Interval = 30
Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_N.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
    Image1(1).Tag = 1
End If
Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_N.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Call SendData("FINBAN")
    Unload Me
End If
End Sub

Private Sub Image1_Click(Index As Integer)


Select Case Index
    Case 0
If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        If UserGLD >= NPCInventory(List1(0).ListIndex + 1).Valor * Val(cantidad) Then
                SendData ("COMP" & "," & List1(0).ListIndex + 1 & "," & cantidad.text)
                
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If

    List1(0).Clear
    List1(1).Clear
        
   Case 1
If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub
        LastIndex2 = List1(1).ListIndex
        If Not Inventario.Equipped(List1(1).ListIndex + 1) Then
            SendData ("VEND" & "," & List1(1).ListIndex + 1 & "," & cantidad.text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
        
    List1(0).Clear
    List1(1).Clear
        
    Case 2
        If List1(Todo).ListIndex >= 0 Then
            If Todo = 1 Then
                cantidad = Label1(2).Caption
            Else
                cantidad = Label1(2).Caption
            End If
        End If
                
End Select

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_I.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_I.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
                Image1(0).Tag = 1
        End If
    Case 2
        Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_I.jpg")
        'If Image1(2).Tag = 1 Then
        '        Image1(2).Picture = general_load_interface_picture("Botónokapretado.jpg")
        '        Image1(2).Tag = 0
        'End If
        
End Select
End Sub

Private Sub cantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgQuit_Click()
    Call SendData("FINBAN")
    Unload Me
End Sub

Private Sub List1_Click(Index As Integer)
Dim SR As RECT, DR As RECT, GrhIndex As Integer

SR.left = 0
SR.top = 0
SR.Right = 32
SR.bottom = 32

DR.left = 0
DR.top = 0
DR.Right = 32
DR.bottom = 32

Todo = Index

Select Case Index
    Case 0
        Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).Name
        Label1(1).Caption = PonerPuntos(NPCInventory(List1(0).ListIndex + 1).Valor)
        Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).Amount
        GrhIndex = NPCInventory(List1(0).ListIndex + 1).GrhIndex
        Select Case NPCInventory(List1(0).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe: " & NPCInventory(List1(0).ListIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe: " & NPCInventory(List1(0).ListIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & NPCInventory(List1(0).ListIndex + 1).Def
                Label1(4).Visible = True
            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & NPCInventory(List1(0).ListIndex + 1).Def
                Label1(4).Visible = True
            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & NPCInventory(List1(0).ListIndex + 1).Def
                Label1(4).Visible = True
        End Select
        Call engine.DrawGrhtoHdc(GrhIndex, SR, Picture1)
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
        Label1(1).Caption = Inventario.Valor(List1(1).ListIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).ListIndex + 1)
        GrhIndex = Inventario.GrhIndex(List1(1).ListIndex + 1)
        Select Case Inventario.OBJType(List1(1).ListIndex + 1)
            Case 2
                Label1(3).Caption = "Max Golpe: " & Inventario.MaxHit(List1(1).ListIndex + 1)
                Label1(4).Caption = "Min Golpe: " & Inventario.MinHit(List1(1).ListIndex + 1)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(List1(1).ListIndex + 1)
                Label1(4).Visible = True
            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(List1(1).ListIndex + 1)
                Label1(4).Visible = True
            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa: " & Inventario.Def(List1(1).ListIndex + 1)
                Label1(4).Visible = True
        End Select
        Call engine.DrawGrhtoHdc(Inventario.GrhIndex(List1(1).ListIndex + 1), SR, Picture1)
End Select

'Call engine.GrhRenderToHdc(GrhIndex, Picture1.hDC, 0, 0)

End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub tmrNumber_Timer()

Const MIN_NUMBER = 1
Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If

    cantidad.text = format$(m_Number)
    
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_A.jpg")
If Index = 1 Then Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_A.jpg")
If Index = 2 Then Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_A.jpg")
End Sub
Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Image1(0).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Comprar_N.jpg")
If Index = 1 Then Image1(1).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Vender_N.jpg")
If Index = 2 Then Image1(2).Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Comerciar_Todo_N.jpg")
End Sub


