VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tierras Sagradas"
   ClientHeight    =   7650
   ClientLeft      =   16605
   ClientTop       =   11955
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleMode       =   0  'User
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
   Begin VB.Timer Rejas 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   3000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Status"
      Height          =   615
      Left            =   3240
      TabIndex        =   12
      Top             =   0
      Width           =   3855
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "Online"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tierras Sagradas AO"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hacer un World Save"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar PJS"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   6975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   1530
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   6975
   End
   Begin VB.Timer LimpiezaTimer 
      Interval        =   60000
      Left            =   7800
      Top             =   3360
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7800
      Top             =   3000
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   3000
   End
   Begin VB.Timer CmdExec 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   3360
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   8520
      Top             =   3360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BroadCast"
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6975
      Begin VB.CommandButton Command5 
         Caption         =   "Enviar mensaje en SMSG a todos los GM."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3120
         Width           =   6495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Enviar mensaje en SMSG a todos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   6495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Enviar mensaje en RMSG a todos los GM."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   6495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enviar mensaje en RMSG a todos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   6495
      End
      Begin VB.TextBox BroadMsg 
         Height          =   915
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Consola:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Record de usuarios online:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   300
      Width           =   3060
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios online: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   2805
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Acciones"
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
      Begin VB.Menu cmdResetSockets 
         Caption         =   "Reiniciar Sockets"
      End
      Begin VB.Menu cmdBanIps 
         Caption         =   "Ban All IPs"
      End
      Begin VB.Menu cmdUnbanIps 
         Caption         =   "Unban All IPs"
      End
      Begin VB.Menu cmdBanID 
         Caption         =   "Ban ID"
      End
   End
   Begin VB.Menu cmdInterv 
      Caption         =   "Intervalos"
   End
   Begin VB.Menu cmdUsers 
      Caption         =   "Usuarios"
   End
   Begin VB.Menu cmdAbrir 
      Caption         =   "Abrir.."
      Begin VB.Menu cmdDats 
         Caption         =   "Dats"
         Begin VB.Menu cmdObj 
            Caption         =   "Objetos"
         End
         Begin VB.Menu cmdHechiz 
            Caption         =   "Hechizos"
         End
         Begin VB.Menu cmdNpcs 
            Caption         =   "NPC"
         End
         Begin VB.Menu cmdNpcsH 
            Caption         =   "Bichos"
         End
         Begin VB.Menu cmdPremios 
            Caption         =   "Premios"
         End
         Begin VB.Menu cmdQuests 
            Caption         =   "Quests"
         End
      End
      Begin VB.Menu cmdServIni 
         Caption         =   "Server.INI"
      End
   End
   Begin VB.Menu cmdRecargar 
      Caption         =   "Recargar.."
      Begin VB.Menu cmdROBJ 
         Caption         =   "Recargar Objetos"
      End
      Begin VB.Menu cmdRNPC 
         Caption         =   "Recargar NPCs"
      End
      Begin VB.Menu cmdRHechiz 
         Caption         =   "Recargar Hechizos"
      End
      Begin VB.Menu cmdRServini 
         Caption         =   "Recargar Server.ini"
      End
   End
   Begin VB.Menu cmdConfig 
      Caption         =   "Configuracion"
   End
   Begin VB.Menu cmdAcercaDe 
      Caption         =   "Acerca de.."
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG7
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
Dim iUserIndex As Integer

For iUserIndex = 1 To MaxUsers
   
   'Conexion activa? y es un usuario loggeado?
If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged = True Then
    If UserList(iUserIndex).flags.UserNumQuest = 0 Then
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
            Call SendData(SendTarget.toindex, iUserIndex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
            
            'mato los comercios seguros
            If UserList(iUserIndex).cComercio.cComercia = True Then
               comCancelar iUserIndex
            End If
            
            Call Cerrar_Usuario(iUserIndex)
        End If
    End If
  End If
  
Next iUserIndex

End Sub



Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo

'Nos ahorramos el timer de mierda de 3000ms
Static Timeruno As Byte
Timeruno = Timeruno + 1
    If Timeruno = 3 Then
        Dim i As Integer
    
        For i = 1 To MaxUsers
            If UserList(i).flags.UserLogged Then _
                If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
        Next i
        
        Timeruno = 0
    End If
'Nos ahorramos el timer de mierda de 3000ms

'Nos ahorramos el timer npcataca de 2000ms
Dim npc As Integer
Static Timerdos As Byte

Timerdos = Timerdos + 1
    If Timerdos = IntervaloNpcPuedeAtacar Then
        For npc = 1 To LastNPC
            Npclist(npc).CanAttack = 1
        Next npc
        
        Timerdos = 0
    End If
'Nos ahorramos el timer npcataca de 2000ms

Exit Sub

errhand:
Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number & " - " & Timeruno & " - " & Timerdos)
End Sub
Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub
Private Sub CmdExec_Timer()

    Dim i    As Integer

    Static n As Long

    On Error Resume Next ':(((
    n = n + 1

    For i = 1 To MaxUsers

        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            If Not UserList(i).CommandsBuffer.IsEmpty Then
                Call HandleData(i, UserList(i).CommandsBuffer.Pop)
            End If

            If n >= 10 Then
                If UserList(i).ColaSalida.Count > 0 Then ' And UserList(i).SockPuedoEnviar Then
                    #If UsarQueSocket = 1 Then
                        Call IntentarEnviarDatosEncolados(i)
                    #End If
                End If
            End If
        End If

    Next i

    If n >= 10 Then
        n = 0
    End If

    Exit Sub

hayerror:

End Sub
Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub
Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
    Call LimpiaWsApi(frmMain.hWnd)
#ElseIf UsarQueSocket = 0 Then
    Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
    Serv.Detener
#End If

Call DescargaNpcsDat

Dim loopC As Integer

For loopC = 1 To MaxUsers
    If UserList(loopC).ConnID <> -1 Then Call CloseSocket(loopC)
Next

'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server cerrado."
Close #n

End

Set SonidosMapas = Nothing

End Sub
Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim iNpcIndex As Integer

On Error Resume Next

    Dim loopX   As Long
    For loopX = 1 To MAX_BOTS
        If ia_Bot(loopX).Invocado Then ia_Action loopX
    Next loopX

 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iUserIndex = 1 To MaxUsers
   'Conexion activa?
   If UserList(iUserIndex).ConnID <> -1 Then
      '¿User valido?

      If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then

         Call Mod_AntiCheat.RestoTiempo(iUserIndex)
         Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.X, UserList(iUserIndex).Pos.Y)
          
         If UserList(iUserIndex).flags.Muerto = 0 Then
               If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
               If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoVeneno(iUserIndex)
               If UserList(iUserIndex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                
               Call DuracionPociones(iUserIndex)
               Call HambreYSed(iUserIndex)
                
                If (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                    If UserList(iUserIndex).Stats.MinSta < UserList(iUserIndex).Stats.MaxSta Then Call RecStamina(iUserIndex, StaminaIntervaloSinDescansar)
                End If
               
               If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
       End If 'Muerto
     Else 'no esta logeado?
        If UserList(iUserIndex).flags.Stopped = 1 Then Exit Sub
        
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iUserIndex).Counters.IdleCount = 0
              Call Cerrar_Usuario(iUserIndex)
              Call CloseSocket(iUserIndex)
        End If
        
     End If 'UserLogged

   End If

   Next iUserIndex

Exit Sub
'hayerror:
'LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
'[/Alejo]
  'DoEvents
End Sub
Private Sub LimpiezaTimer_Timer()

Dim i As Long

'TIEMPO ONLINE
For i = 1 To LastUser
    UserList(i).flags.TiempoOnlineHoy = UserList(i).flags.TiempoOnlineHoy + 1
    UserList(i).flags.TiempoParaCofres = UserList(i).flags.TiempoParaCofres + 1

    'Cofres cada 40min
    If UserList(i).flags.TiempoParaCofres = 40 Then
        UserList(i).flags.TiempoParaCofres = 0
        
        If UserList(i).flags.AntiAFK < 10 And UserList(i).flags.Paralizado = 0 And UserList(i).flags.Stopped = 0 Then
            Call SendData(SendTarget.toindex, i, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
            Call Cerrar_Usuario(i)
        Else
            Dim Cofrez As Obj
            Cofrez.ObjIndex = 1549
            Cofrez.Amount = 2
        
            If Not MeterItemEnInventario(i, Cofrez) Then
                Call TirarItemAlPiso(UserList(i).Pos, Cofrez)
            End If
            Call SendData(SendTarget.toindex, i, 0, "||459")
        End If
        
        UserList(i).flags.AntiAFK = 0
    End If
Next i

        If GranPoder = 0 Then
            OtorgarGranPoder (0)
        Else
         If UserList(GranPoder).flags.Muerto = 1 Or UserList(GranPoder).flags.Privilegios > PlayerType.User Or MapInfo(UserList(GranPoder).Pos.Map).Pk = False Or UserList(GranPoder).Pos.Map = 78 Or UserList(GranPoder).Pos.Map = 101 Or UserList(GranPoder).Pos.Map = 18 Or UserList(GranPoder).Pos.Map = 54 Or UserList(GranPoder).Pos.Map = 8 Or UserList(GranPoder).Pos.Map = 72 Or UserList(GranPoder).Pos.Map = 100 Then
           OtorgarGranPoder (0)
         End If
        End If
        
'############################SUBASTAS#################################
If MinutinSubasta > 0 And Hay_Subasta = True Then
MinutinSubasta = MinutinSubasta - 1
 
    If MinutinSubasta >= 1 And MinutinSubasta < 3 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||468@" & MinutinSubasta)
    ElseIf MinutinSubasta = 0 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||468@0")
 
    Hay_Subasta = False
    
    If UltimoOfertador = "" Then
    
        'Si esta offline el que subasto, le devolvemos el item vía correo.
        If NameIndex(Subastador) <= 0 Then
            Dim NumCorreos As Byte
            Dim NueCorreos As String
            Dim NTCR As String
            Dim CorreoTemporal As String
            Dim iMoC As Long
            
            NumCorreos = GetVar(CharPath & Subastador & ".chr", "CORREO", "NUMCORREOS")
            NueCorreos = GetVar(CharPath & Subastador & ".chr", "CORREO", "NUECORREOS")
            Call WriteVar(CharPath & Subastador & ".chr", "CORREO", "CORREONUM" & NumCorreos + 1, "Servidor$Recibiste un objeto$La subasta finalizo sin ninguna oferta y te devolvimos el objeto que subastaste.$" & Date & "$" & objetosubastado.ObjIndex & "-" & objetosubastado.Amount & "-" & ObjData(objetosubastado.ObjIndex).Name & ",0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),")
            Call WriteVar(CharPath & Subastador & ".chr", "CORREO", "NUMCORREOS", NumCorreos + 1)
            
            For iMoC = 1 To 30
                CorreoTemporal = ReadField(iMoC, NueCorreos, Asc(","))
                If iMoC = NumCorreos + 1 Then
                    NTCR = NTCR & iMoC & "-1,"
                Else
                    NTCR = NTCR & iMoC & "-" & ReadField(2, CorreoTemporal, Asc("-")) & ","
                End If
            Next iMoC
            
            Call WriteVar(CharPath & Subastador & ".chr", "CORREO", "NUECORREOS", NTCR)
        Else
            Call MeterItemEnInventario(NameIndex(Subastador), objetosubastado)
        End If
        
        Call SendData(SendTarget.ToAll, 0, 0, "||469")
        Subastador = ""
    Else
        
        'Si esta offline el que oferto, alteramos el charfile con un correo que tenga el objeto subastado.
        If NameIndex(UltimoOfertador) <= 0 Then
            NumCorreos = GetVar(CharPath & UltimoOfertador & ".chr", "CORREO", "NUMCORREOS")
            NueCorreos = GetVar(CharPath & UltimoOfertador & ".chr", "CORREO", "NUECORREOS")
            Call WriteVar(CharPath & UltimoOfertador & ".chr", "CORREO", "CORREONUM" & NumCorreos + 1, "Servidor$Recibiste un objeto$La subasta finalizo y recibiste el objeto que compraste.$" & Date & "$" & objetosubastado.ObjIndex & "-" & objetosubastado.Amount & "-" & ObjData(objetosubastado.ObjIndex).Name & ",0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),0-0-(Nada),")
            Call WriteVar(CharPath & UltimoOfertador & ".chr", "CORREO", "NUMCORREOS", NumCorreos + 1)

            For iMoC = 1 To 30
                CorreoTemporal = ReadField(iMoC, NueCorreos, Asc(","))
                If iMoC = NumCorreos + 1 Then
                    NTCR = NTCR & iMoC & "-1,"
                Else
                    NTCR = NTCR & iMoC & "-" & ReadField(2, CorreoTemporal, Asc("-")) & ","
                End If
            Next iMoC
            
            Call WriteVar(CharPath & UltimoOfertador & ".chr", "CORREO", "NUECORREOS", NTCR)
        Else
            Call MeterItemEnInventario(NameIndex(UltimoOfertador), objetosubastado)
        End If
        
        'Si esta offline el que subasto le alteramos el charfile con el nuevo oro
        If NameIndex(Subastador) <= 0 Then
            Dim OroTemporal As Long
            OroTemporal = GetVar(CharPath & Subastador & ".chr", "STATS", "GLD")
            
            Call WriteVar(CharPath & Subastador & ".chr", "STATS", "GLD", OroTemporal + OroOfrecido)
            
        Else
            UserList(NameIndex(Subastador)).Stats.GLD = UserList(NameIndex(Subastador)).Stats.GLD + OroOfrecido
            SendUserGLD (NameIndex(Subastador))
        End If
        
        Call SendData(SendTarget.ToAll, 0, 0, "||470@" & UltimoOfertador & "@" & PonerPuntos(OroOfrecido))
        Subastador = ""
        UltimoOfertador = ""
    End If
     
    End If
End If
'############################SUBASTAS#################################
        
        
'###MENSAJE AUTOMATICO
If Len(TextoMensajeAutomatico) > 1 Then
    MinutitosMensaje = MinutitosMensaje + 1
    
    If MinutitosMensaje = TiempoMensajeAutomatico Then
        Call SendData(SendTarget.ToAll, 0, 0, "N|" & TextoMensajeAutomatico)
        MinutitosMensaje = 0
    End If
End If
'###MENSAJE AUTOMATICO
        
        
'#########################REY####################################
If ReyON = 0 Then
    MinutosRey = MinutosRey + 1
    
        'Posiciones Guardias
        Dim Guardia1 As WorldPos
        Dim Guardia2 As WorldPos
        Dim Guardia3 As WorldPos
        Dim Guardia4 As WorldPos
       
        Dim Guardia As Integer
        Guardia = 938
       
        Guardia1.Map = 95
        Guardia1.X = 50
        Guardia1.Y = 17
       
        Guardia2.Map = 95
        Guardia2.X = 49
        Guardia2.Y = 18
       
        Guardia3.Map = 95
        Guardia3.X = 51
        Guardia3.Y = 18
     
        Guardia4.Map = 95
        Guardia4.X = 50
        Guardia4.Y = 19
        '/Posiciones Guardias
     
    Dim PosicionR As WorldPos
        Dim Rey As Integer
        Rey = 937
     
        PosicionR.Map = 95
        PosicionR.X = 50
        PosicionR.Y = 18
       
       
        If MinutosRey = 60 Then
            Call SendData(ToAll, 0, 0, "||471")
            IndexReyAncalagon = SpawnNpc(Rey, PosicionR, True, False)
            Npclist(IndexReyAncalagon).Char.AuraA = 3
            Call MakeNPCChar(SendTarget.ToMap, 0, 0, IndexReyAncalagon, Npclist(IndexReyAncalagon).Pos.Map, Npclist(IndexReyAncalagon).Pos.X, Npclist(IndexReyAncalagon).Pos.Y)
            MinutosRey = 0
            GuardiasRey = 0
            Call SpawnNpc(Guardia, Guardia1, True, False)
            Call SpawnNpc(Guardia, Guardia2, True, False)
            Call SpawnNpc(Guardia, Guardia3, True, False)
            Call SpawnNpc(Guardia, Guardia4, True, False)
            ReyON = 1
        End If
End If

'#############################REY########################################

'############### TIEMPO DUELOS  ##############################

If TiempoDuelo(1) > 0 Then
TiempoDuelo(1) = TiempoDuelo(1) - 1

If TiempoDuelo(1) = 0 Then

    Dim jjj As Long
        For jjj = 1 To LastUser
         If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 1 Then
                UserList(jjj).flags.EnDuelo = False
                UserList(jjj).flags.DueliandoContra = ""
                UserList(jjj).flags.LeMandaronDuelo = False
                UserList(jjj).flags.UltimoEnMandarDuelo = ""
                UserList(jjj).flags.EnQueArena = 0
                UserList(jjj).Stats.GLD = UserList(jjj).Stats.GLD + UserList(jjj).flags.ApuestaOro
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
         End If
         
            If UserList(jjj).flags.EspectadorArena1 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena1 = 0
                EspectadoresEnArena1 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||472@1"
    ArenaOcupada(1) = False
    NombreDueleando(1) = ""
    NombreDueleando(2) = ""
    
    End If
End If

If TiempoDuelo(2) > 0 Then
TiempoDuelo(2) = TiempoDuelo(2) - 1

If TiempoDuelo(2) = 0 Then

    For jjj = 1 To LastUser
     If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 2 Then
            UserList(jjj).flags.EnDuelo = False
            UserList(jjj).flags.DueliandoContra = ""
            UserList(jjj).flags.LeMandaronDuelo = False
            UserList(jjj).flags.UltimoEnMandarDuelo = ""
            UserList(jjj).flags.EnQueArena = 0
            UserList(jjj).Stats.GLD = UserList(jjj).Stats.GLD + UserList(jjj).flags.ApuestaOro
            WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
     End If
     
            If UserList(jjj).flags.EspectadorArena2 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena2 = 0
                EspectadoresEnArena2 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||472@2"
    ArenaOcupada(2) = False
    NombreDueleando(3) = ""
    NombreDueleando(4) = ""
    
    End If
End If

If TiempoDuelo(3) > 0 Then
TiempoDuelo(3) = TiempoDuelo(3) - 1

If TiempoDuelo(3) = 0 Then

    For jjj = 1 To LastUser
     If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 3 Then
            UserList(jjj).flags.EnDuelo = False
            UserList(jjj).flags.DueliandoContra = ""
            UserList(jjj).flags.LeMandaronDuelo = False
            UserList(jjj).flags.UltimoEnMandarDuelo = ""
            UserList(jjj).flags.EnQueArena = 3
            UserList(jjj).Stats.GLD = UserList(jjj).Stats.GLD + UserList(jjj).flags.ApuestaOro
            WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
     End If
     
            If UserList(jjj).flags.EspectadorArena3 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena3 = 0
                EspectadoresEnArena3 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||472@3"
    ArenaOcupada(3) = False
    NombreDueleando(5) = ""
    NombreDueleando(6) = ""
    
    End If
End If

If TiempoDuelo(4) > 0 Then
TiempoDuelo(4) = TiempoDuelo(4) - 1

If TiempoDuelo(4) = 0 Then

    For jjj = 1 To LastUser
     If UserList(jjj).flags.EnDuelo = True And UserList(jjj).flags.EnQueArena = 4 Then
            UserList(jjj).flags.EnDuelo = False
            UserList(jjj).flags.DueliandoContra = ""
            UserList(jjj).flags.LeMandaronDuelo = False
            UserList(jjj).flags.UltimoEnMandarDuelo = ""
            UserList(jjj).flags.EnQueArena = 0
            UserList(jjj).Stats.GLD = UserList(jjj).Stats.GLD + UserList(jjj).flags.ApuestaOro
            WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
     End If
     
            If UserList(jjj).flags.EspectadorArena4 = 1 Then
                WarpUserChar jjj, UserList(jjj).flags.MapaAnterior, UserList(jjj).flags.XAnterior, UserList(jjj).flags.YAnterior, True
                UserList(jjj).flags.EspectadorArena4 = 0
                EspectadoresEnArena4 = 0
            End If
    Next jjj
    
    SendData SendTarget.ToAll, 0, 0, "||472@4"
    ArenaOcupada(4) = False
    NombreDueleando(7) = ""
    NombreDueleando(8) = ""
    
    End If
End If

'##################################LIMPIEZA Y PREMIOS CASTILLOS############################
PremiosCastis = PremiosCastis - 1

'Restamos un minuto a los objetos tirados.
CleanWorld_Clear
'Borramos los objetos al finalizar los 10 minutos.

If PremiosCastis = 0 Then
    If (NumUsers + BOnlines) >= 10 Then
        Call DarPremioCastillos
    End If
    
    Call GuardarUsuarios
    PremiosCastis = 60
End If
'##################################LIMPIEZA Y PREMIOS CASTILLOS############################


'##################################AUTO SAVE######################################
'fired every minute
Static MinutosLatsClean As Long
Static MinsSocketReset As Long
Static MinsPjesSave As Long
Static MinutosNumUsersCheck As Long

Dim num As Long

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call ModAreas.AreasOptimizacion
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

#If UsarQueSocket = 1 Then
' ok la cosa es asi, este cacho de codigo es para
' evitar los problemas de socket. a menos que estes
' seguro de lo que estas haciendo, te recomiendo
' que lo dejes tal cual está.
' alejo.
MinsSocketReset = MinsSocketReset + 1
' cada 1 minutos hacer el checkeo
If MinsSocketReset >= 5 Then
    MinsSocketReset = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then
                If UserList(i).flags.Stopped = 0 Then
                    Call Cerrar_Usuario(i)
                    Call CloseSocket(i)
                End If
        End If
    Next i
    'Call ReloadSokcet
    
    Call LogCriticEvent("NumUsers: " & NumUsers & " WSAPISock2Usr: " & WSAPISock2Usr.Count)
End If
#End If

Call PurgarPenas
Call CheckIdleUser
'##################################AUTO SAVE######################################


End Sub

Private Sub mnuCerrar_Click()


If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub
Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub
Private Sub Rejas_Timer()
    Dim Daño As Byte
    Dim PosicionAriete As WorldPos
    Dim DevolucionAriete As WorldPos
    Dim Arietaso As Obj
    Arietaso.ObjIndex = 1469
    Arietaso.Amount = 1
    
    If RejaCentralAtacada = False And RejaNorteAtacada = False And RejaSurAtacada = False Then Me.Enabled = False
    
    '/REJA SUR
    If RejaSurAtacada = True Then
    
        Daño = RandomNumber(150, 250)
    
    
        If Daño > RejaSur Then
        
            'Cambiamos la reja.
             MapData(81, 49, 84).OBJInfo.ObjIndex = 1472
             Call ModAreas.SendToAreaByPos(81, 49, 84, "HO" & ObjData(1472).GrhIndex & "," & 49 & "," & 84)
             
                            'Desbloquea
                            MapData(81, 49, 84).Blocked = 0
                            MapData(81, 49 - 1, 84).Blocked = 0
                            MapData(81, 49 - 2, 84).Blocked = 0
                            MapData(81, 49 + 1, 84).Blocked = 0
                            MapData(81, 49 + 2, 84).Blocked = 0
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49, 84, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 - 1, 84, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 - 2, 84, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 + 1, 84, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 + 2, 84, 0)
            
            'Avisamos y devolvemos el ariete.
            Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||473@sur")
                DevolucionAriete.Map = Npclist(ArieteUno).Pos.Map
                DevolucionAriete.X = Npclist(ArieteUno).Pos.X
                DevolucionAriete.Y = Npclist(ArieteUno).Pos.Y
                
            Call QuitarNPC(ArieteUno)
            Call TirarItemAlPiso(DevolucionAriete, Arietaso)
            
            RejaSur = 0
            RejaSurAtacada = False
    
        Else
            'Restamos.
            RejaSur = RejaSur - Daño
            
            'Avisamos.
            If RejaSur > 2000 Then
                Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||474@sur")
            Else
                Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||475@sur")
            End If
            
        End If
    
    End If
    
    
    '/REJA NORTE
    If RejaNorteAtacada = True Then
    
        Daño = RandomNumber(150, 250)
    
    
        If Daño > RejaNorte Then
        
            'Cambiamos la reja.
             MapData(81, 49, 48).OBJInfo.ObjIndex = 1472
             Call ModAreas.SendToAreaByPos(81, 49, 48, "HO" & ObjData(1472).GrhIndex & "," & 49 & "," & 48)
             
                            'Desbloquea
                            MapData(81, 49, 48).Blocked = 0
                            MapData(81, 49 - 1, 48).Blocked = 0
                            MapData(81, 49 - 2, 48).Blocked = 0
                            MapData(81, 49 + 1, 48).Blocked = 0
                            MapData(81, 49 + 2, 48).Blocked = 0
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49, 48, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 - 1, 48, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 - 2, 48, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 + 1, 48, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 + 2, 48, 0)
            
            'Avisamos y devolvemos el ariete.
            Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||473@norte")
                DevolucionAriete.Map = Npclist(ArieteTres).Pos.Map
                DevolucionAriete.X = Npclist(ArieteTres).Pos.X
                DevolucionAriete.Y = Npclist(ArieteTres).Pos.Y
                
            Call QuitarNPC(ArieteTres)
            Call TirarItemAlPiso(DevolucionAriete, Arietaso)
            
            RejaNorte = 0
            RejaNorteAtacada = False
    
        Else
            'Restamos.
            RejaNorte = RejaNorte - Daño
            
            'Avisamos.
            If RejaNorte > 2000 Then
                Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||474@norte")
            Else
                Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||475@norte")
            End If
            
        End If
    
    End If
    
    
    If RejaCentralAtacada = True Then
    
        Daño = RandomNumber(150, 250)
    
    
        If Daño > RejaCentral Then
        
            'Cambiamos la reja.
             MapData(81, 49, 68).OBJInfo.ObjIndex = 1472
             Call ModAreas.SendToAreaByPos(81, 49, 68, "HO" & ObjData(1472).GrhIndex & "," & 49 & "," & 68)
             
                            'Desbloquea
                            MapData(81, 49, 68).Blocked = 0
                            MapData(81, 49 - 1, 68).Blocked = 0
                            MapData(81, 49 - 2, 68).Blocked = 0
                            MapData(81, 49 + 1, 68).Blocked = 0
                            MapData(81, 49 + 2, 68).Blocked = 0
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49, 68, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 - 1, 68, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 - 2, 68, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 + 1, 68, 0)
                            Call Bloquear(SendTarget.ToAll, 0, 0, 81, 49 + 2, 68, 0)
            
            'Avisamos y devolvemos el ariete.
            Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||473@central")
                DevolucionAriete.Map = Npclist(ArieteDos).Pos.Map
                DevolucionAriete.X = Npclist(ArieteDos).Pos.X
                DevolucionAriete.Y = Npclist(ArieteDos).Pos.Y
                
            Call QuitarNPC(ArieteDos)
            Call TirarItemAlPiso(DevolucionAriete, Arietaso)
            
            RejaCentral = 0
            RejaCentralAtacada = False
    
        Else
            'Restamos.
            RejaCentral = RejaCentral - Daño
            
            'Avisamos.
            If RejaCentral > 2000 Then
                Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||474@central")
            Else
                Call SendData(SendTarget.ToDiosesYclan, GuildIndex(Fortaleza), 0, "||475@central")
            End If
            
        End If
    
    End If
End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
                ''ia comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                     'Usamos AI si hay algun user en el mapa
                     If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                     End If
                     mapa = Npclist(NpcIndex).Pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                  End If
                          End If
                     End If
                     
                End If
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        TCPServ.SetDato ID, NewIndex

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
Dim T() As String
Dim loopC As Long
Dim RD As String
On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If

RD = StrConv(Datos, vbUnicode)

UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD

T = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(T) > 0 Then
    UserList(MiDato).RDBuffer = T(UBound(T))
    
    For loopC = 0 To UBound(T) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If T(loopC) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(T(loopC)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then
                Call HandleData(MiDato, T(loopC))
              Else
                Exit Sub
              End If
        End If
    Next loopC
End If
Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAcercaDe_Click()
MsgBox "Tierras Sagradas."
End Sub
Private Sub cmdBanID_Click()

Dim hdbanned As String
hdbanned = InputBox("ID del Ciruja:", "Ban ID")

If CheckHD(hdbanned) Then
MsgBox "" & hdbanned & ", ya se encuentra en la lista de HD's baneados."
Exit Sub
Else
Open "" & App.Path & "\DAT\BanHds.dat" For Append As #1
Print #1, hdbanned
Close #1
MsgBox "ID: " & hdbanned & " agregada a la lista de HD's baneadas."
End If

End Sub
Private Sub cmdConfig_Click()
frmServidor.Show , frmMain
End Sub
Private Sub cmdHechiz_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Dat/Hechizos.dat", "", "", 1)
End Sub
Private Sub cmdInterv_Click()
FrmInterv.Show
End Sub

Private Sub cmdNpcs_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Dat/NPCs.dat", "", "", 1)
End Sub

Private Sub cmdNpcsH_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Dat/NPCs-Hostiles.dat", "", "", 1)
End Sub
Private Sub cmdObj_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Dat/Obj.dat", "", "", 1)
End Sub
Private Sub cmdPremios_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Dat/Premios.dat", "", "", 1)
End Sub
Private Sub cmdQuests_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Dat/Quests.dat", "", "", 1)
End Sub
Private Sub cmdRHechiz_Click()
Call CargarHechizos
End Sub
Private Sub cmdRNPC_Click()
Call CargaNpcsDat
End Sub
Private Sub cmdROBJ_Click()
Call LoadOBJData
End Sub
Private Sub cmdRServini_Click()
Call LoadSini
End Sub
Private Sub cmdServIni_Click()
Call ShellExecute(frmMain.hWnd, "open", App.Path & "/Server.ini", "", "", 1)
End Sub

Private Sub cmdUnbanIps_Click()
Dim i As Long, n As Long

Dim sENtrada As String

sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distición de mayusculas minusculas para desbanear a todos los personajes", "UnBan", "hola")
If sENtrada = "estoy DE acuerdo" Then
    
    n = BanIps.Count
    For i = 1 To BanIps.Count
        BanIps.Remove 1
    Next i
    
    MsgBox "Se han habilitado " & n & " ipes"
End If

End Sub

Private Sub cmdUsers_Click()
Form1.Show , frmMain
End Sub

Private Sub Command1_Click()
Me.MousePointer = 11
Call GuardarUsuarios
Me.MousePointer = 0
MsgBox "Grabado de personajes OK!"
End Sub
Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, 0, "||476@" & BroadMsg.Text)
End Sub

Private Sub Command3_Click()
Call SendData(SendTarget.ToAdmins, 0, 0, "||477@" & BroadMsg.Text)
End Sub

Private Sub Command4_Click()
Call SendData(SendTarget.ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Private Sub Command5_Click()
Call SendData(SendTarget.ToAdmins, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Private Sub Command6_Click()
    MsgBox "Utiliza /GRABAR ingame o GRABAR PJS!"
End Sub
