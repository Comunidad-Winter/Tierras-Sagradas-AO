VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Monotype Corsiva"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   7935
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerFotos 
      Interval        =   5000
      Left            =   10200
      Top             =   360
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   360
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1680
      Top             =   960
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   960
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   960
      ScaleHeight     =   3900
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   1920
      Width           =   6255
      Begin VB.Label TextoFoto 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   55
         TabIndex        =   2
         Top             =   3250
         Width           =   6135
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   960
      ScaleHeight     =   3900
      ScaleWidth      =   6255
      TabIndex        =   3
      Top             =   8000
      Width           =   6255
      Begin VB.Label TextoFoto2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3450
         Width           =   6255
      End
   End
   Begin VB.Image Librerias 
      Height          =   405
      Left            =   120
      Picture         =   "Form1.frx":23070
      Top             =   290
      Width           =   1350
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   11300
      Top             =   320
      Width           =   255
   End
   Begin VB.Label lEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   750
      TabIndex        =   0
      Top             =   6840
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   11600
      Top             =   320
      Width           =   255
   End
   Begin VB.Image imgJugar 
      Height          =   675
      Left            =   9405
      Picture         =   "Form1.frx":283E7
      Top             =   6600
      Width           =   1995
   End
   Begin VB.Image imgLauncher 
      Height          =   240
      Left            =   765
      Picture         =   "Form1.frx":2DC2B
      Top             =   6855
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private form_Moviment As clsFormMovementManager

' Esta otra API es para lanzar el ejecutable en modo administrador, usando el verbo indocumentado “runas” como parámetro lpOperation.
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" _
( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer, Poronga As Long, Chota As Long, CantFotos As Byte
Private Sub Form_Load()
    imgLauncher.Width = 0
    Me.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher.gif")
    imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_B.jpg")
    imgJugar.Enabled = False
    
    CantFotos = GetVar(App.Path & "\Data\INIT\versiones.ini", "TEXTOS", "Cantidad")
    
    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
   
    MakeFormTransparent Me, vbRed
    
        If GetVar(App.Path & "\Data\INIT\versiones.ini", "VERSION", "PRIMERAVEZ") = 0 Then
                ShellExecute Me.hwnd, "runas", App.Path & "\AdminLibSetup.exe", "", App.Path, vbNormalFocus
                lEstado.Caption = "Es la primera vez que ejecutas TSAO, estamos ejecutando el instalador de librerías, espera.."
                Timer2.Enabled = True
        Else
            Timer1.Enabled = True
        End If
        
Dim FotoSiguiente As Byte
FotoSiguiente = RandomNumber(1, CantFotos)

Picture1.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Foto_" & FotoSiguiente & ".jpg")

Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me
    
End Sub

Private Sub Image1_Click()
    If MsgBox("¿Desea cerrar el launcher?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
       End
    End If
End Sub
Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, data
    Close #F
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgJugar.Enabled = True Then
        imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_N.jpg")
    ElseIf imgJugar.Enabled = False Then
        imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_B.jpg")
    End If
    
    Librerias.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Librerias_L.jpg")
End Sub

Private Sub Image2_Click()
Me.WindowState = vbMinimized
End Sub
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", 0)
End Function
Private Sub imgJugar_Click()
    If MsgBox("¿Desea ejecutar el juego?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
       ShellExecute Me.hwnd, "", App.Path & "\Tierras Sagradas.exe", "", App.Path, vbNormalFocus
       End
    End If
End Sub
Private Sub imgJugar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_A.jpg")
End Sub
Private Sub imgJugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_I.jpg")
End Sub
Private Sub Librerias_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Librerias.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Librerias_LA.jpg")
End Sub
Private Sub Librerias_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Librerias.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Librerias_LI.jpg")
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
    Select Case State
        Case icError
            lEstado.Caption = "Error en la conexión, descarga abortada."
            bDone = True
            dError = True
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long
            
            FileSize = Inet1.GetHeader("Content-length")
            'ProgressBar1.Max = FileSize
            Poronga = FileSize
            Chota = 0
            
            lEstado.Caption = "Iniciando descarga.."
            
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    Chota = Chota + Len(vtData) * 2
                    imgLauncher.Width = (((Chota / 100) / (Poronga / 100)) * 7935)
                   'lEstado.Caption = "Descargando parche: " & (imgLauncher.Width + Len(vtData) * 2) / 1000000 & " MBs de " & (FileSize / 1000000) & " MBs"""
                    lEstado.Caption = "Descargando parche: [" & CLng((Chota * 100) / Poronga) & "% completado.]"

                    DoEvents
                Loop
            Close #1
            
            lEstado.Caption = "Descarga finalizada"
            'LSize.Caption = FileSize & "bytes"
            
            bDone = True
    End Select
End Sub

Private Sub Librerias_Click()
    MsgBox "En segundos se te ejecutrá el insalador de librerías, descarga todas las librerias, selecciona los 3 puntitos (...) y hace click en 'Registrar todas las Librerias' y podrás jugar sin problemas!"
    ShellExecute Me.hwnd, "", App.Path & "\AdminLibSetup.exe", "", App.Path, vbNormalFocus
End Sub

Private Sub Timer1_Timer()
 On Error GoTo Error:
    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    lEstado.Caption = "Buscando Actualizaciones.."
    
    iX = Inet1.OpenURL("https://www.tierras-sagradas.com/AU/Actualizaciones.txt") 'Host
    tX = LeerInt(App.Path & "\Data\INIT\Update.tsao")
    DifX = iX - tX
    
    If Not (DifX <= 0) Then
        If MsgBox("Se encontro una actualizacion, se recomienda abrir el launcher en modo Administrador para descargarla correctamente. ¿Desea continuar de todos modos?", vbYesNo, "Tierras Sagradas AO") = vbYes Then
                lEstado.Caption = "Actualizacion encontrada, comenzando a descargar.."
            
                    Inet1.AccessType = icUseDefault
                    dNum = iX
                    
                If DifX > 1 Then
                    Inet1.URL = "https://www.tierras-sagradas.com/AU/ParcheCompleto.zip"
                    Directory = App.Path & "\Data\ParcheCompleto.zip"
                Else
                    Inet1.URL = "https://www.tierras-sagradas.com/AU/Parche" & dNum & ".zip" 'Host
                    Directory = App.Path & "\Data\Parche" & dNum & ".zip"
                End If
                    
                    bDone = False
                    dError = False
                        
                    Form1.Inet1.Execute , "GET"
                    
                    Do While bDone = False
                    DoEvents
                    Loop
                    
                    If dError Then Exit Sub
                    
                    Unzip Directory, App.Path & "\"
                    Kill Directory
                
                Dim NuevaVersion As String
                
                NuevaVersion = Inet2.OpenURL("https://www.tierras-sagradas.com/AU/Version.txt")
                Call WriteVar(App.Path & "\Data\INIT\versiones.ini", "VERSION", "V", NuevaVersion)
        Else
            End
        End If
    End If
     
    Call GuardarInt(App.Path & "\Data\INIT\Update.tsao", iX)
    
    imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_N.jpg")
    imgJugar.Enabled = True
    
    imgLauncher.Width = 7935
    lEstado.Caption = "Cliente actualizado correctamente."
    
    Timer1.Enabled = False
    
Error:
    imgJugar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Jugar_N.jpg")
    imgJugar.Enabled = True
    
    imgLauncher.Width = 7935
    lEstado.Caption = "Cliente actualizado correctamente."
    
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
       
       Call WriteVar(App.Path & "\Data\INIT\versiones.ini", "VERSION", "PRIMERAVEZ", "1")
       
       Timer1.Enabled = True
       Timer2.Enabled = False
End Sub
Private Sub TimerFotos_Timer()

Dim i As Single
Dim FotoSiguiente As Byte
Dim FotoAnterior As Byte
FotoSiguiente = RandomNumber(1, CantFotos)

If FotoAnterior = FotoSiguiente Then FotoSiguiente = RandomNumber(1, CantFotos)
FotoAnterior = FotoSiguiente

On Error Resume Next
Picture2.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Launcher_Foto_" & FotoSiguiente & ".jpg")
For i = 1 To Picture2.ScaleWidth Step 12
Picture1.PaintPicture Picture2.Picture, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 0, 0, i, Picture2.ScaleHeight, &HCC0020
DoEvents
Next

End Sub
