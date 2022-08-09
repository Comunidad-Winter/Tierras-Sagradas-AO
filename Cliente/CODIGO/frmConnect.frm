VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "frmConnect.frx":000C
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Image img_BorrarPJ 
         Height          =   300
         Left            =   11160
         Top             =   12000
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image img_EntrarPJ 
         Height          =   300
         Left            =   10080
         Top             =   12000
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   0
         Left            =   5580
         Top             =   4080
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   1
         Left            =   5580
         Top             =   6600
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   2
         Left            =   8265
         Top             =   6090
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   3
         Left            =   9480
         Top             =   4095
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   4
         Left            =   8280
         Top             =   2130
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   5
         Left            =   5595
         Top             =   1620
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   6
         Left            =   2880
         Top             =   2130
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   7
         Left            =   1680
         Top             =   4080
         Width           =   975
      End
      Begin VB.Image PJ 
         Height          =   1335
         Index           =   8
         Left            =   2880
         Top             =   6090
         Width           =   975
      End
      Begin VB.Image imgCambiarPass 
         Height          =   495
         Left            =   9240
         Top             =   7680
         Width           =   2340
      End
      Begin VB.Image imgBorrarCuenta 
         Height          =   495
         Left            =   9360
         Top             =   8880
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Image imgSalir4 
         Height          =   615
         Left            =   10200
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Image imgCrearPersonaje 
         Height          =   495
         Left            =   4350
         Top             =   5700
         Width           =   3375
      End
      Begin VB.Image imgName 
         Height          =   375
         Left            =   3855
         Top             =   3390
         Width           =   4620
      End
      Begin VB.Image imgPass 
         Height          =   375
         Left            =   3855
         Top             =   4440
         Width           =   4620
      End
      Begin VB.Image imgAnti 
         Height          =   375
         Left            =   7710
         Top             =   5040
         Width           =   495
      End
      Begin VB.Image imgWeb 
         Height          =   675
         Left            =   7080
         Top             =   7560
         Width           =   1755
      End
      Begin VB.Image imgRecuperarCuenta 
         Height          =   675
         Left            =   5250
         Top             =   7560
         Width           =   1740
      End
      Begin VB.Image imgCrearCuenta 
         Height          =   675
         Left            =   3405
         Top             =   7560
         Width           =   1785
      End
      Begin VB.Image imgConectar 
         Height          =   630
         Left            =   4890
         Top             =   5085
         Width           =   2520
      End
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   240
      Top             =   8190
      Width           =   2535
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WS_EX_APPWINDOW               As Long = &H40000
Private Const GWL_EXSTYLE                   As Long = (-20)
Private Const SW_HIDE                       As Long = 0
Private Const SW_SHOW                       As Long = 5
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Dim i As Long
Dim PJApretado As Byte
Dim BorrarRandom As String
Dim ElRandom As String

Public CMouseX As Integer
Public CMouseY As Integer

Private m_bActivated As Boolean
Private Sub Form_Activate()
    If Not m_bActivated Then
        m_bActivated = True
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
        Call ShowWindow(hWnd, SW_HIDE)
        Call ShowWindow(hWnd, SW_SHOW)
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

If RenderConnect = True Then

    If KeyAscii = vbKeyTab And ClickeoTextCuenta = True Then
            ClickeoTextCuenta = False
            ClickeoTextPassw = True
        Exit Sub
    ElseIf KeyAscii = vbKeyTab And ClickeoTextPassw = True Then
            ClickeoTextCuenta = True
            ClickeoTextPassw = False
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
            nombrecuent = TextBoxCuenta
            passcuent = TextBoxPassw
    
            If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                DoEvents
            End If
           
           
            'update user info
            nombrecuent = TextBoxCuenta
            UserPassword = TextBoxPassw
           
            If CheckUserData(False) = True Then
                EstadoLogin = LoginAccount
                frmConnect.MousePointer = 99
                frmMain.Socket1.HostAddress = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
            End If
    Exit Sub
    End If
    
    
    If ClickeoTextCuenta = True Then
    
        If KeyAscii = vbKeyBack And Len(TextBoxCuenta) = 0 Then Exit Sub
    
        If KeyAscii = vbKeyBack And Len(TextBoxCuenta) <> 0 Then
            TextBoxCuenta = mid(TextBoxCuenta, 1, Len(TextBoxCuenta) - 1)
        Else
            If Len(TextBoxCuenta) >= 15 Then Exit Sub
            TextBoxCuenta = TextBoxCuenta & Chr$(KeyAscii)  'convert to character
        End If
        
    ElseIf ClickeoTextPassw = True Then
    
        If KeyAscii = vbKeyBack And Len(TextBoxPassw) = 0 Then Exit Sub
    
        If KeyAscii = vbKeyBack And Len(TextBoxPassw) <> 0 Then
            TextBoxPassw = mid(TextBoxPassw, 1, Len(TextBoxPassw) - 1)
            TextBoxPasswR = mid(TextBoxPasswR, 1, Len(TextBoxPasswR) - 1)
        Else
            If Len(TextBoxPassw) >= 15 Then Exit Sub
            TextBoxPassw = TextBoxPassw & Chr$(KeyAscii)  'convert to character
            TextBoxPasswR = TextBoxPasswR & "*"
        End If
        
    End If
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        Call UnloadAllForms
End If

End Sub
Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]

 BarritaTextConnect = 0
 
 ButtonLogin = "Normal"
 ButtonCC = "Normal"
 ButtonRC = "Normal"
 ButtonVW = "Normal"
ButtonCP = "Normal"
ButtonCPass = "Normal"
ButtonSalir = "Normal"
ButtonEntrarPJ = "Normal"
ButtonBorrarPJ = "Normal"

End Sub

Private Sub imgAnti_Click()
frmKeypad.Show , frmConnect
End Sub

Private Sub imgConectar_Click()
        nombrecuent = TextBoxCuenta
        passcuent = TextBoxPassw

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
       
       
        'update user info
        nombrecuent = TextBoxCuenta
        UserPassword = TextBoxPassw
       
        If CheckUserData(False) = True Then
            EstadoLogin = LoginAccount
            frmConnect.MousePointer = 99
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
        End If
End Sub
Private Sub imgCrearCuenta_Click()
       EstadoLogin = CrearAccount
       If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
End Sub
Private Sub imgName_DblClick()
    TextBoxCuenta = ""
End Sub
Private Sub imgPass_DblClick()
     TextBoxPassw = ""
     TextBoxPasswR = ""
End Sub
Private Sub imgName_Click()
     ClickeoTextCuenta = True
     ClickeoTextPassw = False
End Sub
Private Sub imgPass_Click()
     ClickeoTextCuenta = False
     ClickeoTextPassw = True
End Sub

Private Sub imgRecuperarCuenta_Click()
        EstadoLogin = RecuPW

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
            
        frmRecuperar.Visible = True
End Sub
Private Sub imgWeb_Click()
    End
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If RenderConnect = True Then
 ButtonLogin = "Normal"
 ButtonCC = "Normal"
 ButtonRC = "Normal"
 ButtonVW = "Normal"
End If
 
    CMouseX = X
    CMouseY = Y
 
If RenderAccount = True Then
    For i = 0 To 9
    If MostrarTodo(i) = True Then
        MostrarTodo(i) = False
        img_EntrarPJ.Visible = False
        img_BorrarPJ.Visible = False
        img_EntrarPJ.top = 800
        img_BorrarPJ.top = 800
    End If
    
    If CrearAura(i) = True Then
     CrearAura(i) = False
    End If
    Next i
    
    ButtonCP = "Normal"
    ButtonCPass = "Normal"
    ButtonSalir = "Normal"
    ButtonEntrarPJ = "Normal"
    ButtonBorrarPJ = "Normal"
End If

End Sub
Private Sub Timer1_Timer()

BarritaTextConnect = BarritaTextConnect + 1

If BarritaTextConnect > 240 Then
BarritaTextConnect = 0
End If

End Sub
Private Sub imgConectar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonLogin = "Iluminado"
End Sub
Private Sub imgConectar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonLogin = "Apretado"
End Sub
Private Sub imgConectar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonLogin = "Normal"
End Sub
Private Sub imgCrearCuenta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCC = "Iluminado"
End Sub
Private Sub imgCrearCuenta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCC = "Apretado"
End Sub
Private Sub imgCrearCuenta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCC = "Normal"
End Sub
Private Sub imgRecuperarCuenta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonRC = "Iluminado"
End Sub
Private Sub imgRecuperarCuenta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonRC = "Apretado"
End Sub
Private Sub imgRecuperarCuenta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonRC = "Normal"
End Sub
Private Sub imgWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonVW = "Iluminado"
End Sub
Private Sub imgWeb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonVW = "Apretado"
End Sub
Private Sub imgWeb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonVW = "Normal"
End Sub
Private Sub img_BorrarPJ_Click()
      BorrarRandom = RandomNumber(1000, 9999)
      ElRandom = InputBox("Esta accion no podra ser revertida, para confirmar ingrse el codigo " & BorrarRandom & " para borrar su personaje.", "Borrar Personaje")
        
      If BorrarRandom = ElRandom Then Call SendData("TBRP" & CargarPJ(PJApretado).Nombre & "," & nombrecuent & "," & CodigoRecibido)
End Sub
Private Sub img_EntrarPJ_Click()
    SendData ("OOLOGI" & CargarPJ(PJApretado).Nombre & "," & nombrecuent & "," & CodigoRecibido)
End Sub
Private Sub PJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

For i = 0 To 9
If CrearAura(i) = True Then
 CrearAura(i) = False
End If
Next i

If MostrarTodo(Index) = True Then
 CrearAura(Index) = False
 ButtonEntrarPJ = "Normal"
 ButtonBorrarPJ = "Normal"
Else
 CrearAura(Index) = True
    img_EntrarPJ.Visible = False
    img_BorrarPJ.Visible = False
        img_EntrarPJ.top = 800
        img_BorrarPJ.top = 800
End If

End Sub
Private Sub PJ_Click(Index As Integer)

For i = 0 To 9
If MostrarTodo(i) = True Then
    MostrarTodo(i) = False
End If

If CrearAura(i) = True Then
 CrearAura(i) = False
End If
Next i

If Index = 0 Then
    EntrarX = 362 '-28
    EntrarY = 260 '-7 o 40¿?
ElseIf Index = 1 Then
    EntrarX = 362
    EntrarY = 427
ElseIf Index = 2 Then
    EntrarX = 542
    EntrarY = 392
ElseIf Index = 3 Then
    EntrarX = 622
    EntrarY = 260
ElseIf Index = 4 Then
    EntrarX = 542
    EntrarY = 129
ElseIf Index = 5 Then
    EntrarX = 362
    EntrarY = 95
ElseIf Index = 6 Then
    EntrarX = 184
    EntrarY = 129
ElseIf Index = 7 Then
    EntrarX = 102
    EntrarY = 260
ElseIf Index = 8 Then
    EntrarX = 184
    EntrarY = 392
End If
    
img_EntrarPJ.top = EntrarY
img_EntrarPJ.left = EntrarX
img_BorrarPJ.top = EntrarY
img_BorrarPJ.left = EntrarX + 68

If CargarPJ(Index).Existe = True Then
    MostrarTodo(Index) = True
    img_EntrarPJ.Visible = True
    img_BorrarPJ.Visible = True
    PJApretado = Index
End If

End Sub
Private Sub imgCrearPersonaje_Click()
    If CargarPJ(8).Existe = True Then
        Mensaje.Escribir "No puedes crear más personajes."
    Else
        Call Audio.PlayWave("click.wav")
    
        EstadoLogin = Dados
        'frmCuent.Visible = False
        frmCrearPersonaje.Show , frmConnect
        Audio.StopWave
        'frmCrearPersonaje.MousePointer = 11
    End If
End Sub
Private Sub imgSalir4_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup

    Call CambiarConectar("CONECTAR")

    frmConnect.Visible = True
End Sub
Private Sub imgCambiarPass_Click()

        Call Audio.PlayWave("click.wav")
        Dim anteriorpw As String
        Dim nuevapw As String
        Dim renuevapw As String
        
        anteriorpw = InputBox("Ingrese su actual contraseña:", "Cambiar Password")
        nuevapw = InputBox("Ingrese su nueva contraseña:", "Cambiar Password")
        renuevapw = InputBox("Repita su nueva contraseña:", "Cambiar Password")
        
        If nuevapw <> renuevapw Then
            Mensaje.Escribir "Las passwords que tipeo no coinciden"
            Exit Sub
        End If
        
        If Len(nuevapw) > 15 Then
            Mensaje.Escribir "La password no puede superar los 15 caracteres"
            Exit Sub
        End If
        
        If InStr(1, nuevapw, "ñ") > 0 Or InStr(1, nuevapw, "Ñ") > 0 Then
                Mensaje.Escribir "No puedes utilizar la letra ñ en la contraseña."
            Exit Sub
        End If
        
        SendData ("REPASS" & nombrecuent & "," & anteriorpw & "," & nuevapw & "," & renuevapw)
End Sub
Private Sub imgCrearPersonaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCP = "Iluminado"
End Sub
Private Sub imgCrearPersonaje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCP = "Apretado"
End Sub
Private Sub imgCrearPersonaje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCP = "Normal"
End Sub
Private Sub imgSalir4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonSalir = "Iluminado"
End Sub
Private Sub imgSalir4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonSalir = "Apretado"
End Sub
Private Sub imgSalir4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonSalir = "Normal"
End Sub
Private Sub imgCambiarPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPass = "Iluminado"
End Sub
Private Sub imgCambiarPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPass = "Apretado"
End Sub
Private Sub imgCambiarPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonCPass = "Normal"
End Sub
Private Sub img_EntrarPJ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonEntrarPJ = "Iluminado"
End Sub
Private Sub img_EntrarPJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonEntrarPJ = "Apretado"
End Sub
Private Sub img_EntrarPJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonEntrarPJ = "Normal"
End Sub
Private Sub img_BorrarPJ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonBorrarPJ = "Iluminado"
End Sub
Private Sub img_BorrarPJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonBorrarPJ = "Apretado"
End Sub
Private Sub img_BorrarPJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonBorrarPJ = "Normal"
End Sub

