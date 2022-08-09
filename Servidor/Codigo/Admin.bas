Attribute VB_Name = "Admin"
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
'Argentum Online is based on Baronsoft's VB6 Online RPG
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

Public Type tMotd
    Texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public ReiniciarServer As Long

Public tInicioServer As Long

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNpcPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloAutoReiniciar As Long   'segundos

Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public DeNoche As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

Public Type TCPESStats
    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date
End Type

Public TCPESStats As TCPESStats
Public Function ValidarLoginMSG(ByVal n As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(n)
AuxInteger2 = SDM(n)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function
Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long

Call SendData(SendTarget.ToAll, 0, 0, "||656")

Dim j As Integer, k As Integer

For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

For loopX = 1 To NumMaps
    'DoEvents
    If MapInfo(loopX).BackUp = 1 Then
        Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
    End If
Next loopX

If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
        Call BackUPnPc(loopX)
    End If
Next

Call SendData(SendTarget.ToAll, 0, 0, "||657")

End Sub

Public Sub PurgarPenas()
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                Call SendData(SendTarget.toindex, i, 0, "||658@" & UserList(i).Counters.Pena)
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(SendTarget.toindex, i, 0, "||444")
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal userindex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(userindex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(userindex, Prision.Map, Prision.X, Prision.Y, True)
        
        Call SendData(SendTarget.toindex, userindex, 0, "||659@" & Minutos)
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function CuentaExiste(ByVal Cuenta As String) As Boolean
 
CuentaExiste = FileExist(App.Path & "\Accounts\" & UCase$(Cuenta) & ".act", vbNormal)
 
End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")

'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function
Public Function CheckHD(ByVal hd As String) As Boolean
'***************************************************
'Author: Nahuel Casas (Zagen)
'Last Modify Date: 07/12/2009
' 07/12/2009: Zagen - Agregè la funcion de agregar los digitos de un Serial Baneado.
'***************************************************
Open App.Path & "\DAT\BanHds.dat" For Input As #1
Dim Linea As String, Total As String

Do Until EOF(1)
Line Input #1, Linea
Total = Total + Linea + vbCrLf
Loop
Close #1

Dim Ret As String
If InStr(1, Total, hd) Then
CheckHD = True
End If
End Function
Public Sub BanIpAgrega(ByVal ip As String)
BanIps.Add ip

Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIps.Count And Dale
    Dale = (BanIps.Item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim n As Long

n = BanIpBuscar(ip)
If n > 0 Then
    BanIps.Remove n
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim LoopC As Long

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

ArchN = FreeFile()
Open ArchivoBanIp For Output As #ArchN

For LoopC = 1 To BanIps.Count
    Print #ArchN, BanIps.Item(LoopC)
Next LoopC

Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

Do While BanIps.Count > 0
    BanIps.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoBanIp For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    BanIps.Add Tmp
Loop

Close #ArchN

End Sub
Public Function UserDarPrivilegioLevel(ByVal Name As String) As Long
    If IsAdministrator(Name) Then
        UserDarPrivilegioLevel = 6
    ElseIf IsDevelopment(Name) Then
        UserDarPrivilegioLevel = 5
    ElseIf IsCoordination(Name) Then
        UserDarPrivilegioLevel = 4
    ElseIf IsTournamentManager(Name) Then
        UserDarPrivilegioLevel = 3
    ElseIf IsEventManager(Name) Then
        UserDarPrivilegioLevel = 2
    ElseIf IsUserSupport(Name) Then
        UserDarPrivilegioLevel = 1
    Else
        UserDarPrivilegioLevel = 0
    End If
End Function

