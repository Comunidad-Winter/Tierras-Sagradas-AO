VERSION 5.00
Begin VB.Form frmTeclas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Teclas"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   21
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   5340
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   20
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5340
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   19
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3180
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   18
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4620
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4620
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   12
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "F"
      Top             =   300
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usar las teclas 0 a 7 del teclado numerico de la derecha para cambiar modo de habla."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   36
      Top             =   6120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   300
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1740
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2460
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3180
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3900
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   6
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   7
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1740
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   10
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2460
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   11
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3180
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar teclas"
      Height          =   315
      Index           =   0
      Left            =   5400
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Modo predeterminado"
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5400
      TabIndex        =   4
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   300
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   14
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   15
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1740
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   16
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2460
      Width           =   2415
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro de Resurrecion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   22
      Left            =   5760
      TabIndex        =   48
      Top             =   5640
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   21
      Left            =   2760
      TabIndex        =   47
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro de Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   20
      Left            =   120
      TabIndex        =   45
      Top             =   5040
      Width           =   1470
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desactivar Musica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   5400
      TabIndex        =   43
      Top             =   2880
      Width           =   1665
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar Emoticons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   2760
      TabIndex        =   41
      Top             =   4320
      Width           =   1665
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   39
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atacar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   585
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tirar objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usar objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   32
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Equipar objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   31
      Top             =   2880
      Width           =   1320
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar / Desactivar Seguro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   30
      Top             =   3600
      Width           =   2445
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar / Ocultar Nicknames"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   2760
      TabIndex        =   29
      Top             =   3600
      Width           =   2520
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro de Resurrecion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   5400
      TabIndex        =   28
      Top             =   3600
      Width           =   2085
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2760
      TabIndex        =   27
      Top             =   720
      Width           =   570
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar posición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   2760
      TabIndex        =   26
      Top             =   1440
      Width           =   1680
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultarse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   2760
      TabIndex        =   25
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo combate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   2760
      TabIndex        =   24
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia arriba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   5400
      TabIndex        =   23
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia abajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   5400
      TabIndex        =   22
      Top             =   720
      Width           =   1785
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia la izquierda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   5400
      TabIndex        =   21
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia la derecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   5400
      TabIndex        =   20
      Top             =   2160
      Width           =   2220
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar screenshot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   2760
      TabIndex        =   19
      Top             =   0
      Width           =   1635
   End
End
Attribute VB_Name = "frmTeclas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TempVars(0 To NUMBINDS) As Integer

Function AlreadyBinded(KeyCode As Integer) As Boolean

Dim i As Integer

If (KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12) Or (KeyCode = 44) Or (KeyCode = 106) Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad7) Then
    AlreadyBinded = True
    Exit Function
End If

For i = 1 To NUMBINDS
    If (TempVars(i - 1) = KeyCode) Then
        AlreadyBinded = True
        Exit Function
    End If
Next i

End Function

Private Sub Check1_Click()
'UseNumPad = Check1.value
'Call WriteVar(App.Path & "\Data\INIT\Config.ini", "Opciones", "UsarNumpad", Val(UseNumPad))
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim i As Integer
Dim lc As Integer
Dim bCambio As Boolean
Dim Resultado As VbMsgBoxResult

Select Case Index
    
    Case 0
        Call GuardaConfigEnVariables
        For lc = 1 To NUMBINDS
            Call WriteVar(App.Path & "\Data\INIT\Teclas.tsao", "TECLAS", str(lc), str(BindKeys(lc).KeyCode) & "," & BindKeys(lc).Name)
        Next lc
    Case 1
        Call LoadDefaultBinds
        Call CargaConfigEnForm
    Case 2
    
        For i = 1 To NUMBINDS
            If TempVars(i - 1) <> BindKeys(i).KeyCode Then
                bCambio = True
                Exit For
            End If
        Next
        
        If bCambio Then
            Resultado = MsgBox("Realizo cambios en la configuración ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")
            If Resultado = vbYes Then
                Call GuardaConfigEnVariables
                For lc = 1 To NUMBINDS
                    Call WriteVar(App.Path & "\Data\INIT\Teclas.tsao", "TECLAS", str(lc), str(BindKeys(lc).KeyCode) & "," & BindKeys(lc).Name)
                Next lc
            End If
        End If
        
        If Resultado <> vbCancel Then Unload Me

End Select

End Sub

Private Sub txConfig_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim Name As String
Name = txConfig(Index).text

If KeyCode > 0 Then
    
    If AlreadyBinded(KeyCode) Then
        Beep
        txConfig(Index).ForeColor = vbRed
        Exit Sub
    Else
        txConfig(Index).ForeColor = vbBlack
    End If
    
    If KeyCode = vbKeyShift Then
        Name = "Shift"
    ElseIf KeyCode = vbKeyLeft Then
        Name = "Flecha Izquierda"
    ElseIf KeyCode = vbKeyRight Then
        Name = "Flecha Derecha"
    ElseIf KeyCode = vbKeyDown Then
        Name = "Flecha Abajo"
    ElseIf KeyCode = vbKeyUp Then
        Name = "Flecha Arriba"
    ElseIf KeyCode = vbKeyControl Then
        Name = "Control"
    ElseIf KeyCode = 18 Then
        Name = "Alt"
    ElseIf KeyCode = vbKeyPageDown Then
        Name = "Page Down"
    ElseIf KeyCode = vbKeyPageUp Then
        Name = "Page Up"
    ElseIf KeyCode = vbKeySeparator Then 'Enter teclado numerico
        Name = "Intro"
    ElseIf KeyCode = vbKeySpace Then
        Name = "Barra Espaciadora"
    ElseIf KeyCode = vbKeyDelete Then
        Name = "Delete"
    ElseIf KeyCode = vbKeyEnd Then
        Name = "Fin"
    ElseIf KeyCode = vbKeyHome Then
        Name = "Inicio"
    ElseIf KeyCode = vbKeyInsert Then
        Name = "Insert"
    ElseIf KeyCode = 192 Then
        Name = "Ñ"
    Else
        Name = Chr(KeyCode)
    End If
    
    TempVars(Index) = KeyCode
    txConfig(Index).text = Name

End If

End Sub
Private Sub GuardaConfigEnVariables()

Dim i As Integer

For i = 1 To NUMBINDS
    BindKeys(i).Name = txConfig(i - 1).text
    BindKeys(i).KeyCode = TempVars(i - 1)
Next

End Sub
Private Sub CargaConfigEnForm()

Dim i As Integer

For i = 1 To NUMBINDS
    txConfig(i - 1).text = BindKeys(i).Name
    TempVars(i - 1) = BindKeys(i).KeyCode
Next

End Sub
Private Sub Form_Load()
Call CargaConfigEnForm
'Check1.value = UseNumPad
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim i As Integer
Dim bCambio As Boolean
Dim Resultado As VbMsgBoxResult

For i = 1 To NUMBINDS
    If TempVars(i - 1) <> BindKeys(i).KeyCode Then
        bCambio = True
        Exit For
    End If
Next

If bCambio Then
    Resultado = MsgBox("Realizo cambios en la configuración ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")
    If Resultado = vbYes Then Call GuardaConfigEnVariables
End If

If Resultado = vbCancel Then Cancel = 1

End Sub

Sub LoadDefaultBinds()

Dim Arch, LACONCHA As String, lc As Integer
Arch = App.Path & "\Data\INIT\" & "Teclas.tsao"

For lc = 1 To NUMBINDS
    LACONCHA = GetVar(App.Path & "\Data\INIT\" & "Teclas.tsao", "DEFAULTS", str(lc))
    BindKeys(lc).KeyCode = Val(ReadField(1, LACONCHA, 44))
    BindKeys(lc).Name = ReadField(2, LACONCHA, 44)
Next lc


End Sub
