VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Tierras Sagradas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imgProgress 
      Height          =   585
      Left            =   2355
      Top             =   7245
      Width           =   7515
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'MawenAO frmCargando -www.mawenao.net
'Todos los derechos reservados
'agush
'www.gs-zone.org
 
Option Explicit
 
Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 336
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
Private Const WS_EX_APPWINDOW               As Long = &H40000
Private Const GWL_EXSTYLE                   As Long = (-20)
Private Const SW_HIDE                       As Long = 0
Private Const SW_SHOW                       As Long = 5
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
 
Private m_bActivated As Boolean
 
Private Sub Form_Activate()
    If Not m_bActivated Then
        m_bActivated = True
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
        Call ShowWindow(hWnd, SW_HIDE)
        Call ShowWindow(hWnd, SW_SHOW)
    End If
End Sub
Private Sub Form_Load()
 Me.Picture = LoadPicture(DirGraficos & "cargando.jpg")
'Me.Icon = LoadPicture(App.path & "\Data\GRAFICOS\Icono.ico")
imgProgress.Picture = LoadPicture(DirGraficos & "cargando_barra.jpg")
End Sub
Public Sub ProgresoBarra(ByVal porc As Long)

    If porc = 0 Then
        imgProgress.Width = 0
        Exit Sub
    End If
        
    
    Dim num As Long, i As Long
    num = (501 * porc) / 100
    
    If num >= 501 Then
        imgProgress.Width = 501
        Exit Sub
    End If
        
    For i = imgProgress.Width To num
        imgProgress.Width = imgProgress.Width + 1
    Next

End Sub
