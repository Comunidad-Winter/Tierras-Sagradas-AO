Attribute VB_Name = "TCP"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio, Jonatan Ezequiel Salguero
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

'RUTAS DE ENVIO DE DATOS
Public Enum SendTarget
    toindex = 0         'Envia a un solo User
    ToAll = 1           'A todos los Users
    ToMap = 2           'Todos los Usuarios en el mapa
    ToPCArea = 3        'Todos los Users en el area de un user determinado
    ToNone = 4          'Ninguno
    ToAllButIndex = 5   'Todos menos el index
    ToMapButIndex = 6   'Todos en el mapa menos el indice
    ToGM = 7
    ToNPCArea = 8       'Todos los Users en el area de un user determinado
    ToGuildMembers = 9
    ToAdmins = 10
    ToPCAreaButIndex = 11
    ToAdminsAreaButConsejeros = 12
    ToDiosesYclan = 13
    ToConsejo = 14
    ToClanArea = 15
    ToConsejoCaos = 16
    ToRolesMasters = 17
    ToDeadArea = 18
    ToCiudadanos = 19
    ToCriminales = 20
    ToPartyArea = 21
    ToReal = 22
    ToCaos = 23
    ToCiudadanosYRMs = 24
    ToCriminalesYRMs = 25
    ToRealYRMs = 26
    ToCaosYRMs = 27
End Enum


#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpoYCabeza(ByRef UserBody As Integer, ByRef UserHead As Integer, ByVal Raza As String, ByVal Gen As String)
'TODO: Poner las heads en arrays, así se acceden por índices
'y no hay problemas de discontinuidad de los índices.
'También se debe usar enums para raza y sexo
Select Case Gen
   Case "Hombre"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 30)
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(1, 13) + 100
                If UserHead = 113 Then UserHead = 201       'Un índice no es continuo.... :S muy feo
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(1, 8) + 201
                UserBody = 3
            Case "Enano"
                UserHead = RandomNumber(1, 5) + 300
                UserBody = 52
            Case "Gnomo"
                UserHead = RandomNumber(1, 6) + 400
                UserBody = 52
            Case Else
                UserHead = 1
                UserBody = 1
        End Select
   Case "Mujer"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 7) + 69
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(1, 7) + 169
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(1, 11) + 269
                UserBody = 3
            Case "Gnomo"
                UserHead = RandomNumber(1, 5) + 469
                UserBody = 52
            Case "Enano"
                UserHead = RandomNumber(1, 3) + 369
                UserBody = 52
            Case Else
                UserHead = 70
                UserBody = 1
        End Select
End Select

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal userindex As Integer) As Boolean

Dim loopC As Integer

For loopC = 1 To NUMSKILLS
    If UserList(userindex).Stats.UserSkills(loopC) < 0 Then
        Exit Function
        If UserList(userindex).Stats.UserSkills(loopC) > 100 Then UserList(userindex).Stats.UserSkills(loopC) = 100
    End If
Next loopC

ValidateSkills = True
    
End Function
Public Function IsYourChr(ByVal Account As String, ByVal PJ As String)
 
Dim i As Integer
Dim NumPjs As Integer
Dim ChrToView As String
 
 
 
NumPjs = GetVar(App.Path & "\Accounts\" & Account & ".act", "PJS", "NumPjs")
 
IsYourChr = False
 
For i = 1 To NumPjs
    ChrToView = GetVar(App.Path & "\Accounts\" & Account & ".act", "PJS", "PJ" & i)
    If ChrToView = PJ Then IsYourChr = True
Next i
 
End Function
 
Sub ConnectAccount(ByVal userindex As Integer, Name As String, Password As String)
 
Dim i As Integer
Dim Pjjj As String
Dim NumPjs As Integer
Dim ArchivodeUser As String
Dim Pos() As String
Dim Oro() As Long
Dim Nivel() As String
Dim PuntosdeCanje() As Integer
Dim OroBanco() As Byte
Dim cosa As Integer
 
 
If Password <> GetVar(App.Path & "\Accounts\" & Name & ".act", Name, "password") Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERRPassword incorrecto.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If GetVar(App.Path & "\Accounts\" & Name & ".act", Name, "Ban") = 1 Then
    Dim MotivitoTemporal As String
    MotivitoTemporal = GetVar(App.Path & "\Accounts\" & Name & ".act", Name, "Motivo")
        Call SendData(SendTarget.toindex, userindex, 0, "ERRTu cuenta se encuentra actualmente baneada por: " & ReadField(2, MotivitoTemporal, Asc(",")) & " con motivo: " & ReadField(1, MotivitoTemporal, Asc(", ")) & ".")
    Call CloseSocket(userindex)
    Exit Sub
End If
 
UserList(userindex).Accounted = Name
UserList(userindex).AccountedPass = Password
 
NumPjs = GetVar(App.Path & "\Accounts\" & Name & ".act", "PJS", "NumPjs")
 
If TienePjs(Name) = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "INIAC" & Name & "," & NumPjs + 1)
Else
    Call SendData(SendTarget.toindex, userindex, 0, "INIAC0")
End If
ArchivodeUser = App.Path & "\charfile\"
For i = 1 To NumPjs
    Pjjj = GetVar(App.Path & "\Accounts\" & Name & ".act", "PJS", "PJ" & i)
    If Pjjj = "" Then Exit Sub
    Call LoadUserAccount(Pjjj & ".chr")
    Call SendData(SendTarget.toindex, userindex, 0, "ADDPJ" & Pjjj & "," & i & "," & PJEnCuenta)
Next i

Call SendData(SendTarget.toindex, userindex, 0, "DLPSJ" & NumPjs)
Call SendData(SendTarget.toindex, userindex, 0, "CODEH" & GetVar(App.Path & "\Accounts\" & Name & ".act", "SEGURIDAD", "CodeX"))

End Sub
Sub ChrToAccount(ByVal Accounted As String, tName As String)
 
Dim NumPjs As Integer
Dim n As Integer
 
NumPjs = GetVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "NumPjs")
 
If NumPjs = 1 And GetVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs) = "" Then
    Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "NumPjs", NumPjs)
    Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs, tName)
    Exit Sub
End If
 
NumPjs = NumPjs + 1
 
Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "NumPjs", NumPjs)
Call WriteVar(App.Path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs, tName)
 
 
End Sub
Sub CreateAccount(ByVal Account As String, Password As String, Mail As String, pregunta As String, Respuesta As String, userindex As Integer)
 
On Error GoTo Errhandler
 
If FileExist(App.Path & "\Accounts\" & Account & ".act", vbNormal) = True Then
Call SendData(SendTarget.toindex, userindex, 0, "ERREl nombre de la cuenta ya está siendo utilizado por otro usuario.")
    Exit Sub
End If
 
Dim n As Integer
Dim i As Integer
 
 
n = FreeFile()
 
Open App.Path & "\Accounts\" & Account & ".act" For Output As n
    Print #n, "[" & Account & "]"
    Print #n, "password=" & Password
    Print #n, "mail=" & Mail
    Print #n, "Pregunta=" & pregunta
    Print #n, "Respuesta=" & Respuesta
    Print #n, "Ban=0"
    Print #n, "BANCO=0"
    Print #n, "CambioPW=2"
    Print #n, "AntiguaPW=RecienCreada"
    Print #n, "[SEGURIDAD]"
    Print #n, "CodeX=" & codex
    Print #n, "[PJS]"
    Print #n, "NumPjs=0"
    Print #n, "PJ1="
    Print #n, "PJ2="
    Print #n, "PJ3="
    Print #n, "PJ4="
    Print #n, "PJ5="
    Print #n, "PJ6="
    Print #n, "PJ7="
    Print #n, "PJ8="
    Print #n, "PJ9="
    Print #n, "[BancoInventory]"
    Print #n, "CantidadItems=0"
    Print #n, "Obj1=0-0"
    Print #n, "Obj2=0-0"
    Print #n, "Obj3=0-0"
    Print #n, "Obj4=0-0"
    Print #n, "Obj5=0-0"
    Print #n, "Obj6=0-0"
    Print #n, "Obj7=0-0"
    Print #n, "Obj8=0-0"
    Print #n, "Obj9=0-0"
    Print #n, "Obj10=0-0"
    Print #n, "Obj11=0-0"
    Print #n, "Obj12=0-0"
    Print #n, "Obj13=0-0"
    Print #n, "Obj14=0-0"
    Print #n, "Obj15=0-0"
    Print #n, "Obj16=0-0"
    Print #n, "Obj17=0-0"
    Print #n, "Obj18=0-0"
    Print #n, "Obj19=0-0"
    Print #n, "Obj20=0-0"
    Print #n, "Obj21=0-0"
    Print #n, "Obj22=0-0"
    Print #n, "Obj23=0-0"
    Print #n, "Obj24=0-0"
    Print #n, "Obj25=0-0"
    Print #n, "Obj26=0-0"
    Print #n, "Obj27=0-0"
    Print #n, "Obj28=0-0"
    Print #n, "Obj29=0-0"
    Print #n, "Obj30=0-0"
    Print #n, "Obj31=0-0"
    Print #n, "Obj32=0-0"
    Print #n, "Obj33=0-0"
    Print #n, "Obj34=0-0"
    Print #n, "Obj35=0-0"
    Print #n, "Obj36=0-0"
    Print #n, "Obj37=0-0"
    Print #n, "Obj38=0-0"
    Print #n, "Obj39=0-0"
    Print #n, "Obj40=0-0"
    Print #n, "[AMIGOS]"
    Print #n, "A1=(Nadie)"
    Print #n, "AD1=0"
    Print #n, "A2=(Nadie)"
    Print #n, "AD2=0"
    Print #n, "A3=(Nadie)"
    Print #n, "AD3=0"
    Print #n, "A4=(Nadie)"
    Print #n, "AD4=0"
    Print #n, "A5=(Nadie)"
    Print #n, "AD5=0"
    Print #n, "A6=(Nadie)"
    Print #n, "AD6=0"
    Print #n, "A7=(Nadie)"
    Print #n, "AD7=0"
    Print #n, "A8=(Nadie)"
    Print #n, "AD8=0"
    Print #n, "A9=(Nadie)"
    Print #n, "AD9=0"
    Print #n, "A10=(Nadie)"
    Print #n, "AD10=0"
    Print #n, "A11=(Nadie)"
    Print #n, "AD11=0"
    Print #n, "A12=(Nadie)"
    Print #n, "AD12=0"
    Print #n, "A13=(Nadie)"
    Print #n, "AD13=0"
    Print #n, "A14=(Nadie)"
    Print #n, "AD14=0"
    Print #n, "A15=(Nadie)"
    Print #n, "AD15=0"
    Print #n, "A16=(Nadie)"
    Print #n, "AD16=0"
    Print #n, "A17=(Nadie)"
    Print #n, "AD17=0"
    Print #n, "A18=(Nadie)"
    Print #n, "AD18=0"
    Print #n, "A19=(Nadie)"
    Print #n, "AD19=0"
    Print #n, "A20=(Nadie)"
    Print #n, "AD20=0"
Close n
 
DoEvents
 
Call SendData(SendTarget.toindex, userindex, 0, "ERR¡Cuenta creada con exito!.")
Call CloseSocket(userindex)
 
Exit Sub
 
Errhandler:
 
Call LogError("NewAccount - Error = " & Err.Number & " - Descripción = " & Err.Description)
 
End Sub
 
Public Function TienePjs(ByVal Account As String) As Boolean
 
Dim frstPj As String
 
frstPj = GetVar(App.Path & "\Accounts\" & Account & ".act", "PJS", "PJ0")
 
If frstPj <> "" Then
    TienePjs = True
Else
    TienePjs = False
End If
 
End Function
'Barrin 3/3/03
'Agregué PadrinoName y Padrino password como opcionales, que se les da un valor siempre y cuando el servidor esté usando el sistema
Sub ConnectNewUser(userindex As Integer, Name As String, UserRaza As String, UserSexo As String, UserClase As String, Hogar As String, _
                     Cuenta As String, Head As String)

If Not AsciiValidos(Name) Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERRNombre invalido.")
    Exit Sub
End If

Dim loopC As Integer
Dim totalskpts As Long

'¿Existe el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

UserList(userindex).flags.Muerto = 0
UserList(userindex).flags.Escondido = 0

UserList(userindex).flags.estado = 0
UserList(userindex).flags.EsNoble = 0
UserList(userindex).flags.CaballerodelDragon = 0

UserList(userindex).Name = Name
UserList(userindex).clase = UserClase
UserList(userindex).Raza = UserRaza
UserList(userindex).Genero = UserSexo
UserList(userindex).Hogar = "Tanaris"
UserList(userindex).Password = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "SEGURIDAD", "CodeX")
UserList(userindex).Stats.Reputacione = 0
UserList(userindex).StatusMith.EsStatus = 8
UserList(userindex).Bon1 = "Ninguno/No elegido."
UserList(userindex).Bon2 = "Ninguno/No elegido."
UserList(userindex).Bon3 = "Ninguno/No elegido."
UserList(userindex).flags.Stopped = 0
UserList(userindex).flags.AlmasContenidas = 0
UserList(userindex).flags.SirvienteDeDios = "N/A"
UserList(userindex).flags.JerarquiaDios = 0

'Misiones diarias
Call OtorgarMisionDiaria(userindex)

UserList(userindex).Stats.UserAtributos(1) = 18
UserList(userindex).Stats.UserAtributos(2) = 18
UserList(userindex).Stats.UserAtributos(3) = 18
UserList(userindex).Stats.UserAtributos(4) = 18
UserList(userindex).Stats.UserAtributos(5) = 18
UserList(userindex).flags.DeseoRecibirMSJ = 1

Select Case UCase$(UserRaza)
    Case "HUMANO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 3
    Case "ELFO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) - 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 2
    Case "ELFO OSCURO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 1
    Case "ENANO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) - 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) - 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 1
    Case "GNOMO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) - 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) - 1
End Select


With UserList(userindex).Stats
Dim Mimamamemima As Byte
 
For Mimamamemima = 1 To NUMSKILLS
    .UserSkills(Mimamamemima) = 100
Next Mimamamemima
 
End With

UserList(userindex).Char.Heading = eHeading.SOUTH

Call DarCuerpoYCabeza(UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Raza, UserList(userindex).Genero)
UserList(userindex).Char.Head = Head
UserList(userindex).OrigChar = UserList(userindex).Char
   
UserList(userindex).Char.WeaponAnim = 12
UserList(userindex).Char.ShieldAnim = NingunEscudo
UserList(userindex).Char.CascoAnim = NingunCasco

UserList(userindex).Stats.MET = 1
Dim MiInt As Long
MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(userindex).Stats.MaxHP = 21
UserList(userindex).Stats.MinHP = 21

UserList(userindex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userindex).Stats.MaxSta = 20 * MiInt
UserList(userindex).Stats.MinSta = 20 * MiInt


UserList(userindex).Stats.MaxAGU = 100
UserList(userindex).Stats.MinAGU = 100

UserList(userindex).Stats.MaxHam = 100
UserList(userindex).Stats.MinHam = 100

UserList(userindex).InventorySlots = 21


'<-----------------MANA----------------------->
If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "NIGROMANTE" Then
    MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 3
    UserList(userindex).Stats.MaxMAN = 100 + MiInt
    UserList(userindex).Stats.MinMAN = 100 + MiInt
ElseIf UCase$(UserClase) = "CLERIGO" Or UCase$(UserClase) = "DRUIDA" _
    Or UCase$(UserClase) = "BARDO" Or UCase$(UserClase) = "ASESINO" Then
        MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 4
        UserList(userindex).Stats.MaxMAN = 50
        UserList(userindex).Stats.MinMAN = 50
Else
    UserList(userindex).Stats.MaxMAN = 0
    UserList(userindex).Stats.MinMAN = 0
End If


If UCase$(UserClase) = "MAGO" Then
    UserList(userindex).Stats.UserHechizos(1) = 65
    UserList(userindex).Stats.UserHechizos(2) = 8
    UserList(userindex).Stats.UserHechizos(3) = 14
    UserList(userindex).Stats.UserHechizos(4) = 43
    UserList(userindex).Stats.UserHechizos(5) = 18
    UserList(userindex).Stats.UserHechizos(6) = 20
    UserList(userindex).Stats.UserHechizos(7) = 11
    UserList(userindex).Stats.UserHechizos(8) = 12
    UserList(userindex).Stats.UserHechizos(9) = 5
    UserList(userindex).Stats.UserHechizos(10) = 15
    UserList(userindex).Stats.UserHechizos(11) = 23
    UserList(userindex).Stats.UserHechizos(12) = 47
    UserList(userindex).Stats.UserHechizos(13) = 33
    UserList(userindex).Stats.UserHechizos(14) = 35
    UserList(userindex).Stats.UserHechizos(15) = 25
    UserList(userindex).Stats.UserHechizos(16) = 9
    UserList(userindex).Stats.UserHechizos(17) = 10
ElseIf UCase$(UserClase) = "CLERIGO" Then
    UserList(userindex).Stats.UserHechizos(1) = 65
    UserList(userindex).Stats.UserHechizos(2) = 8
    UserList(userindex).Stats.UserHechizos(3) = 14
    UserList(userindex).Stats.UserHechizos(4) = 43
    UserList(userindex).Stats.UserHechizos(5) = 18
    UserList(userindex).Stats.UserHechizos(6) = 20
    UserList(userindex).Stats.UserHechizos(7) = 11
    UserList(userindex).Stats.UserHechizos(8) = 12
    UserList(userindex).Stats.UserHechizos(9) = 5
    UserList(userindex).Stats.UserHechizos(10) = 15
    UserList(userindex).Stats.UserHechizos(11) = 23
    UserList(userindex).Stats.UserHechizos(12) = 29
    UserList(userindex).Stats.UserHechizos(13) = 9
    UserList(userindex).Stats.UserHechizos(14) = 10
ElseIf UCase$(UserClase) = "BARDO" Then
    UserList(userindex).Stats.UserHechizos(1) = 65
    UserList(userindex).Stats.UserHechizos(2) = 8
    UserList(userindex).Stats.UserHechizos(3) = 14
    UserList(userindex).Stats.UserHechizos(4) = 43
    UserList(userindex).Stats.UserHechizos(5) = 18
    UserList(userindex).Stats.UserHechizos(6) = 20
    UserList(userindex).Stats.UserHechizos(7) = 11
    UserList(userindex).Stats.UserHechizos(8) = 12
    UserList(userindex).Stats.UserHechizos(9) = 5
    UserList(userindex).Stats.UserHechizos(10) = 15
    UserList(userindex).Stats.UserHechizos(11) = 23
    UserList(userindex).Stats.UserHechizos(12) = 4
    UserList(userindex).Stats.UserHechizos(13) = 9
    UserList(userindex).Stats.UserHechizos(14) = 10
ElseIf UCase$(UserClase) = "DRUIDA" Then
    UserList(userindex).Stats.UserHechizos(1) = 65
    UserList(userindex).Stats.UserHechizos(2) = 8
    UserList(userindex).Stats.UserHechizos(3) = 14
    UserList(userindex).Stats.UserHechizos(4) = 43
    UserList(userindex).Stats.UserHechizos(5) = 18
    UserList(userindex).Stats.UserHechizos(6) = 20
    UserList(userindex).Stats.UserHechizos(7) = 11
    UserList(userindex).Stats.UserHechizos(8) = 12
    UserList(userindex).Stats.UserHechizos(9) = 5
    UserList(userindex).Stats.UserHechizos(10) = 15
    UserList(userindex).Stats.UserHechizos(11) = 23
    UserList(userindex).Stats.UserHechizos(12) = 26
    UserList(userindex).Stats.UserHechizos(13) = 27
    UserList(userindex).Stats.UserHechizos(14) = 4
    UserList(userindex).Stats.UserHechizos(15) = 9
    UserList(userindex).Stats.UserHechizos(16) = 10
ElseIf UCase$(UserClase) = "ASESINO" Then
    UserList(userindex).Stats.UserHechizos(1) = 65
    UserList(userindex).Stats.UserHechizos(2) = 8
    UserList(userindex).Stats.UserHechizos(3) = 14
    UserList(userindex).Stats.UserHechizos(4) = 43
    UserList(userindex).Stats.UserHechizos(5) = 18
    UserList(userindex).Stats.UserHechizos(6) = 20
    UserList(userindex).Stats.UserHechizos(7) = 11
    UserList(userindex).Stats.UserHechizos(8) = 12
    UserList(userindex).Stats.UserHechizos(9) = 5
    UserList(userindex).Stats.UserHechizos(10) = 15
    UserList(userindex).Stats.UserHechizos(11) = 23
    UserList(userindex).Stats.UserHechizos(12) = 22
    UserList(userindex).Stats.UserHechizos(13) = 16
    UserList(userindex).Stats.UserHechizos(14) = 24
    UserList(userindex).Stats.UserHechizos(15) = 10
ElseIf UCase$(UserClase) = "PALADIN" Then
    UserList(userindex).Stats.UserHechizos(1) = 65
    UserList(userindex).Stats.UserHechizos(2) = 8
    UserList(userindex).Stats.UserHechizos(3) = 14
    UserList(userindex).Stats.UserHechizos(4) = 43
    UserList(userindex).Stats.UserHechizos(5) = 18
    UserList(userindex).Stats.UserHechizos(6) = 20
    UserList(userindex).Stats.UserHechizos(7) = 11
    UserList(userindex).Stats.UserHechizos(8) = 12
    UserList(userindex).Stats.UserHechizos(9) = 5
    UserList(userindex).Stats.UserHechizos(10) = 15
    UserList(userindex).Stats.UserHechizos(11) = 23
    UserList(userindex).Stats.UserHechizos(12) = 22
    UserList(userindex).Stats.UserHechizos(13) = 21
    UserList(userindex).Stats.UserHechizos(14) = 24
    UserList(userindex).Stats.UserHechizos(15) = 10
End If

UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1

UserList(userindex).Stats.GLD = 0

UserList(userindex).Stats.Exp = 0
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(userindex).Invent.NroItems = 7

UserList(userindex).Invent.Object(1).ObjIndex = 467
UserList(userindex).Invent.Object(1).Amount = 100

UserList(userindex).Invent.Object(2).ObjIndex = 468
UserList(userindex).Invent.Object(2).Amount = 100

UserList(userindex).Invent.Object(3).ObjIndex = 460
UserList(userindex).Invent.Object(3).Amount = 1
UserList(userindex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case "Humano"
        UserList(userindex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(userindex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(userindex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(userindex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(userindex).Invent.Object(4).ObjIndex = 466
End Select

UserList(userindex).Invent.Object(4).Amount = 1
UserList(userindex).Invent.Object(4).Equipped = 1

UserList(userindex).Invent.ArmourEqpSlot = 4
UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(4).ObjIndex

UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(3).ObjIndex
UserList(userindex).Invent.WeaponEqpSlot = 3

UserList(userindex).Invent.Object(5).ObjIndex = 461
UserList(userindex).Invent.Object(5).Amount = 150
UserList(userindex).Invent.Object(6).ObjIndex = 462
UserList(userindex).Invent.Object(6).Amount = 150
UserList(userindex).Invent.Object(7).ObjIndex = 1491
UserList(userindex).Invent.Object(7).Amount = 150

For loopC = 1 To 30
    UserList(userindex).flags.Correo(loopC) = 0
    UserList(userindex).flags.NueCorreos(loopC) = 0
Next loopC

UserList(userindex).flags.NumCorreos = 0

UserList(userindex).Char.Account = Cuenta
UserList(userindex).Accounted = Cuenta

Dim ln As String

UserList(userindex).BancoInvent.NroItems = CInt(GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "CantidadItems"))
For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = (GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "Obj" & loopC))
    UserList(userindex).BancoInvent.Object(loopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(userindex).BancoInvent.Object(loopC).Amount = CInt(ReadField(2, ln, 45))
Next loopC

UserList(userindex).Stats.Banco = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "" & UserList(userindex).Accounted & "", "BANCO")
Call SaveUser(userindex, CharPath & UCase$(Name) & ".chr")

Call ChrToAccount(Cuenta, Name)

totalPjs = val(GetVar(IniPath & "Server.ini", "INIT", "PJS"))
totalPjs = totalPjs + 1

Call WriteVar(IniPath & "Server.ini", "INIT", "PJS", str(totalPjs))
  
'Open User
Call ConnectUser(userindex, Name, Cuenta, UserList(userindex).Password)
  
End Sub
Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)
Dim loopC As Integer

On Error GoTo Errhandler

If UserList(userindex).flags.Stopped Then Exit Sub

    Call aDos.RestarConexion(UserList(userindex).ip)
    
    If userindex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    If UserList(userindex).flags.Automatico = True Then
        Call Rondas_UsuarioDesconecta(userindex)
    End If
      
    If UserList(userindex).ConnID <> -1 Then
        Call CloseSocketSL(userindex)
    End If
        
    'mato los comercios seguros
    If UserList(userindex).cComercio.cComercia = True Then
       comCancelar userindex
    End If
    
    UserList(userindex).flags.CuentaBancaria = ""
    
    If UserList(userindex).GuildIndex > 0 Then
     Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "||96@" & UserList(userindex).Name)
    End If
    
    If UserList(userindex).flags.Montando = 1 Then
        Call Desmontar(userindex)
    End If

    If UserList(userindex).flags.InvocoMascota = 1 Then
     Call QuitarNPC(UserList(userindex).flags.MascotinIndex)
     UserList(userindex).flags.InvocoMascota = 0
    End If

If UserList(userindex).flags.PartyIndex <> 0 And UserList(userindex).flags.PartyLider = 0 Then
    Call SendData(SendTarget.ToPartyArea, userindex, 0, "||97@" & UserList(userindex).Name)
            UserList(userindex).flags.PartyIndex = 0
            UserList(userindex).flags.PartyLider = 0
            UserList(userindex).flags.PartySolicitud = 0
ElseIf UserList(userindex).flags.PartyIndex <> 0 And UserList(userindex).flags.PartyLider = 1 Then
    Call SendData(SendTarget.ToPartyArea, userindex, 0, "||98")
            
    Dim atodosxdjeje As Long
    For atodosxdjeje = 1 To LastUser
            If UserList(atodosxdjeje).flags.PartyIndex = UserList(userindex).flags.PartyIndex Or UserList(atodosxdjeje).flags.PartySolicitud = UserList(userindex).flags.PartyIndex Then
                UserList(atodosxdjeje).flags.PartyIndex = 0
                UserList(atodosxdjeje).flags.PartyLider = 0
                UserList(atodosxdjeje).flags.PartySolicitud = 0
            End If
    Next atodosxdjeje
    
        UserList(userindex).flags.PartyIndex = 0
        UserList(userindex).flags.PartyLider = 0
        UserList(userindex).flags.PartySolicitud = 0
    
End If

If UserList(userindex).Counters.CreoTeleport = True Then
    Dim mapa As Byte
    Dim X As Byte
    Dim Y As Byte
    
    mapa = UserList(userindex).flags.DondeTiroMap
    X = UserList(userindex).flags.DondeTiroX
    Y = UserList(userindex).flags.DondeTiroY
    
    UserList(userindex).flags.TiroPortalL = 0
    UserList(userindex).Counters.TimeTeleport = 0
    UserList(userindex).Counters.CreoTeleport = False
    Call EraseObj(ToMap, 0, UserList(userindex).flags.DondeTiroMap, MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).OBJInfo.Amount, UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY)
    MapData(mapa, X, Y).TileExit.Map = 0
    MapData(mapa, X, Y).TileExit.X = 0
    MapData(mapa, X, Y).TileExit.Y = 0
    UserList(userindex).flags.DondeTiroMap = 0
    UserList(userindex).flags.DondeTiroX = 0
    UserList(userindex).flags.DondeTiroY = 0
End If

    
    If UserList(userindex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
            
            If (NumUsers + BOnlines) < 10 And GranPoder > 0 Then
                Call SendData(SendTarget.toindex, GranPoder, 0, "||701")
                UserList(GranPoder).flags.GranPoder = 0
                SendUserVariant (userindex)
                GranPoder = 0
            End If
                
        Call CloseUser(userindex)
    Else
        Call ResetUserSlot(userindex)
    End If
    
    UserList(userindex).ConnID = -1
    UserList(userindex).ConnIDValida = False
    
Exit Sub

Errhandler:
    UserList(userindex).ConnID = -1
    UserList(userindex).ConnIDValida = False
    Call ResetUserSlot(userindex)

#If UsarQueSocket = 1 Then
    If UserList(userindex).ConnID <> -1 Then
    Call ControlarPortalLum(userindex)
    UserList(userindex).flags.TiroPortalL = 0
    UserList(userindex).Counters.TimeTeleport = 0
    UserList(userindex).Counters.CreoTeleport = False
    Call CloseSocketSL(userindex)
    End If
#End If

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & userindex)
End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal userindex As Integer)

#If UsarQueSocket = 1 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call BorraSlotSock(UserList(userindex).ConnID)
    Call WSApiCloseSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#End If
End Sub

Public Function EnviarDatosASlot(ByVal userindex As Integer, Datos As String) As Long

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    Ret = WsApiEnviar(userindex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(userindex)
        Call Cerrar_Usuario(userindex)
    End If
    EnviarDatosASlot = Ret
    Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************

    Dim Encolar As Boolean
    Encolar = False
    
    EnviarDatosASlot = 0
    
    If UserList(userindex).ColaSalida.Count <= 0 Then
        If frmMain.Socket2(userindex).Write(Datos, Len(Datos)) < 0 Then
            If frmMain.Socket2(userindex).LastError = WSAEWOULDBLOCK Then
                UserList(userindex).SockPuedoEnviar = False
                Encolar = True
            Else
                Call Cerrar_Usuario(userindex)
            End If
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(userindex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 2 Then '**********************************************

Dim Encolar As Boolean
Dim Ret As Long
    
    Encolar = False
    
    '//
    '// Valores de retorno:
    '//                     0: Todo OK
    '//                     1: WSAEWOULDBLOCK
    '//                     2: Error critico
    '//
    If UserList(userindex).ColaSalida.Count <= 0 Then
        Ret = frmMain.Serv.Enviar(UserList(userindex).ConnID, Datos, Len(Datos))
        If Ret = 1 Then
            Encolar = True
        ElseIf Ret = 2 Then
            Call CloseSocketSL(userindex)
            Call Cerrar_Usuario(userindex)
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(userindex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 3 Then
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(userindex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(userindex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(userindex, True)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & userindex & "/" & UserList(userindex).ConnID & "/" & Datos)
#End If '**********************************************

End Function

Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)

On Error Resume Next

Dim loopC As Integer
Dim X As Integer
Dim Y As Integer

Select Case sndRoute

    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, X, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    
    Case SendTarget.toindex
        If UserList(sndIndex).ConnID <> -1 Then
            Call EnviarDatosASlot(sndIndex, sndData)
            Exit Sub
        End If


    Case SendTarget.ToNone
        Exit Sub
        
        
    Case SendTarget.ToAdmins
        For loopC = 1 To LastUser
            If UserList(loopC).ConnID <> -1 Then
                If UserList(loopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(loopC, sndData)
               End If
            End If
        Next loopC
        Exit Sub
        
    Case SendTarget.ToAll
        For loopC = 1 To LastUser
            If UserList(loopC).ConnID <> -1 Then
                If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) And (loopC <> sndIndex) Then
                If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case SendTarget.ToMap
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If UserList(loopC).flags.UserLogged Then
                    If UserList(loopC).Pos.Map = sndMap Then
                        Call EnviarDatosASlot(loopC, sndData)
                    End If
                End If
            End If
        Next loopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) And loopC <> sndIndex Then
                If UserList(loopC).Pos.Map = sndMap Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
            
    Case SendTarget.ToGuildMembers
        
        loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While loopC > 0
            If (UserList(loopC).ConnID <> -1) Then
                Call EnviarDatosASlot(loopC, sndData)
            End If
            loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub


    Case SendTarget.ToDeadArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).userindex > 0 Then
                        If UserList(MapData(sndMap, X, Y).userindex).flags.Muerto = 1 Or UCase$(UserList(MapData(sndMap, X, Y).userindex).clase) = "CLERIGO" Or UserList(MapData(sndMap, X, Y).userindex).flags.Privilegios >= 1 Or MapInfo(UserList(MapData(sndMap, X, Y).userindex).Pos.Map).Pk = False Then
                           If UserList(MapData(sndMap, X, Y).userindex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).userindex, sndData)
                           End If
                        End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).userindex > 0) And (MapData(sndMap, X, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
       
    Case SendTarget.ToClanArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).userindex > 0) Then
                        If UserList(MapData(sndMap, X, Y).userindex).ConnID <> -1 Then
                            If UserList(sndIndex).GuildIndex > 0 And UserList(MapData(sndMap, X, Y).userindex).GuildIndex = UserList(sndIndex).GuildIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).userindex, sndData)
                            End If
                        End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub



    Case SendTarget.ToPartyArea
        For loopC = 1 To LastUser
            If UserList(loopC).ConnID <> -1 Then
                If UserList(loopC).flags.PartyIndex > 0 And UserList(loopC).flags.PartyIndex = UserList(sndIndex).flags.PartyIndex Then
                    Call EnviarDatosASlot(loopC, sndData)
               End If
            End If
        Next loopC
    Exit Sub
        
    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).userindex > 0) And (MapData(sndMap, X, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).userindex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, X, Y).userindex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).userindex, sndData)
                            End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).Pos.Y - MinYBorder + 1 To Npclist(sndIndex).Pos.Y + MinYBorder - 1
            For X = Npclist(sndIndex).Pos.X - MinXBorder + 1 To Npclist(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, X, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    Case SendTarget.ToDiosesYclan
        loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While loopC > 0
            If (UserList(loopC).ConnID <> -1) Then
                Call EnviarDatosASlot(loopC, sndData)
            End If
            loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        loopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While loopC > 0
            If (UserList(loopC).ConnID <> -1) Then
                Call EnviarDatosASlot(loopC, sndData)
            End If
            loopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub
        
        
Case SendTarget.ToConsejo
For loopC = 1 To LastUser
If (UserList(loopC).ConnID <> -1) Then
If UserList(loopC).ConsejoInfo.PertAlCons > 0 Then
Call EnviarDatosASlot(loopC, sndData)
End If
End If
Next loopC
Exit Sub
Case SendTarget.ToConsejoCaos
For loopC = 1 To LastUser
If (UserList(loopC).ConnID <> -1) Then
If UserList(loopC).ConsejoInfo.PertAlConsCaos > 0 Then
Call EnviarDatosASlot(loopC, sndData)
End If
End If
Next loopC
Exit Sub
    Case SendTarget.ToRolesMasters
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If UserList(loopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case SendTarget.ToCiudadanos
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If Not Criminal(loopC) Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case SendTarget.ToCriminales
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If Criminal(loopC) Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case SendTarget.ToReal
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If UserList(loopC).Faccion.ArmadaReal = 1 Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case SendTarget.ToCaos
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If UserList(loopC).Faccion.FuerzasCaos = 1 Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
        
    Case ToCiudadanosYRMs
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If Not Criminal(loopC) Or UserList(loopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case ToCriminalesYRMs
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If Criminal(loopC) Or UserList(loopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case ToRealYRMs
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If UserList(loopC).Faccion.ArmadaReal = 1 Or UserList(loopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
    
    Case ToCaosYRMs
        For loopC = 1 To LastUser
            If (UserList(loopC).ConnID <> -1) Then
                If UserList(loopC).Faccion.FuerzasCaos = 1 Or UserList(loopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
            End If
        Next loopC
        Exit Sub
End Select

End Sub
Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.Map, X, Y).userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function

Function ValidateChr(ByVal userindex As Integer) As Boolean

If UserList(userindex).Char.Body = 0 Then UserList(userindex).Char.Body = 21

ValidateChr = UserList(userindex).Char.Head <> 0 _
                And UserList(userindex).Char.Body <> 0 _
                And ValidateSkills(userindex)

End Function
Sub ConnectUser(ByVal userindex As Integer, Name As String, Cuenta As String, CodexRecibido As String)
Dim n As Integer
Dim tStr As String

'Reseteamos los FLAGS
UserList(userindex).flags.Escondido = 0
UserList(userindex).flags.TargetNPC = 0
UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
UserList(userindex).flags.TargetObj = 0
UserList(userindex).flags.TargetUser = 0
UserList(userindex).Char.FX = 0

'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(userindex)
    Exit Sub
End If

Dim Numx As Integer
Dim Numxx As Integer
Dim Nick As String
Numx = val(GetVar(IniPath & "HD.ini", "INIT", "GMS"))
 
For Numxx = 1 To Numx
    Nick = UCase$(GetVar(IniPath & "HD.ini", "GM" & Numxx, "Nombre"))
   
    If UCase$(Name) = UCase$(Nick) Then
        If (val(UserList(userindex).hd) <> val(GetVar(IniPath & "HD.ini", "GM" & Numxx, "HD1"))) And (val(UserList(userindex).hd) <> val(GetVar(IniPath & "HD.ini", "GM" & Numxx, "HD2"))) Then
            Call SendData(SendTarget.toindex, userindex, 0, "ERONo podes logear este personaje.")
            Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
            Call CloseSocket(userindex)
          Exit Sub
        End If
    End If
Next Numxx

'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(userindex, UserList(userindex).ip) = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
        Call SendData(SendTarget.toindex, userindex, 0, "ERONo es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

If GetVar(App.Path & "\Accounts\" & Cuenta & ".act", Cuenta, "CambioPW") = 0 Or GetVar(App.Path & "\Accounts\" & Cuenta & ".act", Cuenta, "CambioPW") = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "ERRDebido a que un ex-staff puede tener acceso a tu antigua contraseña, deberas cambiarla obligatoriamente. Hacemos esto por tu seguridad, tratamos de evitar que tu cuenta sea robada.")
        Call CloseSocket(userindex)
    Exit Sub
End If

'¿Existe el personaje?
If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERREl personaje no existe.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Es el passwd valido?
Dim TempCode As String
TempCode = GetVar(CharPath & Name & ".chr", "INIT", "Password")
If UCase$(TempCode) <> UCase$(CodexRecibido) Then
    Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
    Call SendData(SendTarget.toindex, userindex, 0, "EROPassword incorrecto.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(userindex, Name) Then
    If UserList(NameIndex(Name)).Counters.Saliendo Then
        Call SendData(SendTarget.toindex, userindex, 0, "ERREl usuario está deslogeando.")
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "ERRPerdon, un usuario con el mismo nombre se ha logeado, intente de nuevo en 5 minutos.")
    End If
    Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
    Call CloseSocket(userindex)
    Exit Sub
End If

'No dejamos logear un personaje de la misma cuenta ni de casualidad
Dim j As Long
Dim SeteoNumPjs As Byte
Dim SeteoNickName As String
SeteoNumPjs = GetVar(App.Path & "\Accounts\" & Cuenta & ".act", "PJS", "NumPjs")
    
If SeteoNumPjs >= 1 Then
    For j = 1 To SeteoNumPjs
        SeteoNickName = GetVar(App.Path & "\Accounts\" & Cuenta & ".act", "PJS", "PJ" & j)
            If NameIndex(SeteoNickName) > 0 And UCase$(SeteoNickName) <> UCase$(Name) Then
                Call SendData(SendTarget.toindex, userindex, 0, "EROPerdon, un usuario de la misma cuenta está conectado, intente de nuevo en 5 minutos.")
                Call SendData(SendTarget.toindex, userindex, 0, "MEJUI")
                CloseSocket userindex
            Exit Sub
            End If
    Next j
End If
'No dejamos logear un personaje de la misma cuenta ni de casualidad

'Cargamos el personaje
Dim Leer As New clsIniReader
Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")

'Cargamos los datos del personaje
Call LoadUserInit(userindex, Leer)

Call LoadUserStats(userindex, Leer)
Call LoadUserStatus(userindex, Leer)

If Not ValidateChr(userindex) Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERRError en el personaje.")
    Call CloseSocket(userindex)
    Exit Sub
End If

Set Leer = Nothing

If UserList(userindex).Invent.EscudoEqpSlot = 0 Then UserList(userindex).Char.ShieldAnim = NingunEscudo
If UserList(userindex).Invent.CascoEqpSlot = 0 Then UserList(userindex).Char.CascoAnim = NingunCasco
If UserList(userindex).Invent.WeaponEqpSlot = 0 Then UserList(userindex).Char.WeaponAnim = NingunArma


Call UpdateUserInv(True, userindex, 0)
Call UpdateUserHechizos(True, userindex, 0)

If UserList(userindex).flags.Navegando = 1 Then
     UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.BarcoObjIndex).Ropaje
     UserList(userindex).Char.Head = 0
     UserList(userindex).Char.WeaponAnim = NingunArma
     UserList(userindex).Char.ShieldAnim = NingunEscudo
     UserList(userindex).Char.CascoAnim = NingunCasco
End If

If UserList(userindex).flags.Paralizado Then
        Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
End If

'Posicion de comienzo
If UserList(userindex).Pos.Map = 0 Then

    UserList(userindex).Pos = Tanaris

    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex <> 0 Then
        UserList(userindex).Pos = DamePos(UserList(userindex).Pos)
    End If
    
Else

    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex <> 0 Then
        UserList(userindex).Pos = DamePos(UserList(userindex).Pos)
    End If
End If

If Not MapaValido(UserList(userindex).Pos.Map) Then
    Call SendData(SendTarget.toindex, userindex, 0, "ERREL PJ se encuenta en un mapa invalido.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'Nombre de sistema
UserList(userindex).Name = Name
Call Mod_AntiCheat.SetIntervalos(userindex)
Call Mod_AntiCheat.InicializarContadores(userindex)

UserList(userindex).showName = True 'Por default los nombres son visibles

'Aca parcheamos el rollback de clanes
If UserList(userindex).GuildIndex > 0 Then
    If UserList(userindex).GuildIndex > GetVar(App.Path & "\guilds\guildsinfo.inf", "INIT", "NroGuilds") Then
        UserList(userindex).GuildIndex = 0
        Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "GUILDINDEX", "0")
        Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "AspiranteA", "0")
        Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "Miembro", "0")
    Else
        Dim ListMembers As Byte
        Dim EstaEnClan As Boolean
        ListMembers = GetVar(App.Path & "\guilds\" & Guilds(UserList(userindex).GuildIndex).GuildName & "-members.mem", "INIT", "NroMembers")
        EstaEnClan = False
        
        Dim loopC As Long
        For loopC = 1 To ListMembers
            If UCase$(GetVar(App.Path & "\guilds\" & Guilds(UserList(userindex).GuildIndex).GuildName & "-members.mem", "Members", "Member" & loopC)) = UCase$(UserList(userindex).Name) Then
                EstaEnClan = True
            End If
        Next loopC
        
        If EstaEnClan = False Then
            UserList(userindex).GuildIndex = 0
            Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "GUILDINDEX", "0")
            Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "AspiranteA", "0")
            Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "GUILD", "Miembro", "0")
        End If
    End If
End If

'Info
Call SendData(SendTarget.toindex, userindex, 0, "IU" & userindex) 'Enviamos el User index
Call SendData(SendTarget.toindex, userindex, 0, "CM" & UserList(userindex).Pos.Map & "," & MapInfo(UserList(userindex).Pos.Map).r & "," & MapInfo(UserList(userindex).Pos.Map).g & "," & MapInfo(UserList(userindex).Pos.Map).b)
Call SendData(SendTarget.toindex, userindex, 0, "XM" & MapInfo(UserList(userindex).Pos.Map).Music)
Call SendData(SendTarget.toindex, userindex, 0, "N~" & MapInfo(UserList(userindex).Pos.Map).Name)

'Vemos que clase de user es (se lo usa para setear los privilegios alcrear el PJ)
If IsAdministrator(Name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Administrator
    Call LogGM(UserList(userindex).Name, "Se conecto con ip:" & UserList(userindex).ip, False)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||702@administrador " & UserList(userindex).Name)
ElseIf IsDevelopment(Name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Development
    Call LogGM(UserList(userindex).Name, "Se conecto con ip:" & UserList(userindex).ip, False)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||702@desarrollador " & UserList(userindex).Name)
ElseIf IsCoordination(Name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Coordination
    Call LogGM(UserList(userindex).Name, "Se conecto con ip:" & UserList(userindex).ip, False)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||702@coordinador " & UserList(userindex).Name)
ElseIf IsTournamentManager(Name) Then
    UserList(userindex).flags.Privilegios = PlayerType.TournamentManager
    Call LogGM(UserList(userindex).Name, "Se conecto con ip:" & UserList(userindex).ip, False)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||702@manager de torneos " & UserList(userindex).Name)
ElseIf IsEventManager(Name) Then
    UserList(userindex).flags.Privilegios = PlayerType.EventManager
    Call LogGM(UserList(userindex).Name, "Se conecto con ip:" & UserList(userindex).ip, False)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||702@manager de eventos " & UserList(userindex).Name)
ElseIf IsUserSupport(Name) Then
    UserList(userindex).flags.Privilegios = PlayerType.UserSupport
    Call LogGM(UserList(userindex).Name, "Se conecto con ip:" & UserList(userindex).ip, False)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||702@soporte de usuarios " & UserList(userindex).Name)
Else
    UserList(userindex).flags.Privilegios = PlayerType.User
End If

'Bug de mierda (ni idea porque sucederá pero lo solucionamos a lo negro)
If UserList(userindex).Invent.EscudoEqpSlot > 0 Then
    If UserList(userindex).Invent.Object(UserList(userindex).Invent.EscudoEqpSlot).Equipped = 0 Then
        UserList(userindex).Invent.EscudoEqpSlot = 0
    End If
End If

Call SendData(SendTarget.toindex, userindex, 0, "LDG" & UserList(userindex).flags.Privilegios)

'Crea  el personaje del usuario
Call MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
Call SendData(SendTarget.toindex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
''[/el oso]

Call EnviarHambreYsed(userindex)
Call SendUserStatux(userindex)

If haciendoBK Or EnPausa Then
    Call SendData(SendTarget.toindex, userindex, 0, "BKW")
End If

'Actualiza el Num de usuarios
'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
NumUsers = NumUsers + 1
UserList(userindex).flags.UserLogged = True

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "Logged", "1")

MapInfo(UserList(userindex).Pos.Map).NumUsers = MapInfo(UserList(userindex).Pos.Map).NumUsers + 1

If UserList(userindex).Stats.SkillPts > 0 Then
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, UserList(userindex).Stats.SkillPts)
End If

If (NumUsers + BOnlines) > recordusuarios Then
    Call SendData(SendTarget.ToAll, 0, 0, "||703@" & NumUsers + BOnlines)
    recordusuarios = NumUsers + BOnlines
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    frmMain.Label1.caption = "Record de Usuarios Online: " & recordusuarios & ""
End If

If UserList(userindex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasType(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(UserList(userindex).MascotasType(i), UserList(userindex).Pos, True, True)
            
            If UserList(userindex).MascotasIndex(i) > 0 Then
                Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
                Call FollowAmo(UserList(userindex).MascotasIndex(i))
            Else
                UserList(userindex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If UserList(userindex).flags.Navegando = 1 Then Call SendData(SendTarget.toindex, userindex, 0, "NAVEG")

UserList(userindex).flags.Seguro = True
UserList(userindex).flags.SeguroResu = True
UserList(userindex).flags.SeguroClan = True

UserList(userindex).flags.ConsultaEnviada = False
UserList(userindex).flags.NumeroConsulta = 0

UserList(userindex).flags.PartyIndex = 0
UserList(userindex).flags.PartyLider = 0
UserList(userindex).flags.PartySolicitud = 0

Call SendData(SendTarget.toindex, userindex, 0, "RPT" & UserList(userindex).Stats.Reputacione)
Call SendData(SendTarget.toindex, userindex, 0, "SEGCVCON")
UserList(userindex).flags.SeguroCVC = True
If ServerSoloGMs > 0 Then
    If UserList(userindex).flags.Privilegios < ServerSoloGMs Then
        Call SendData(SendTarget.toindex, userindex, 0, "ERRServidor restringido a administradores de jerarquia mayor o igual a: " & ServerSoloGMs & ". Por favor intente en unos momentos.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

If UserList(userindex).GuildIndex > 0 Then
    'welcome to the show baby...
    If Not modGuilds.m_ConectarMiembroAClan(userindex, UserList(userindex).GuildIndex) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||704")
    End If
Else
    UserList(userindex).GuildIndex = 0
End If

tStr = SendFriendList(userindex)
Call SendData(SendTarget.toindex, userindex, 0, "LDM" & SendFriendList(userindex))

Dim atodosxdjeje As Integer
For atodosxdjeje = 1 To LastUser
    Call FriendConnect(atodosxdjeje, UserList(userindex).Name)
Next atodosxdjeje

Call SendData(SendTarget.toindex, userindex, 0, "LOGGED")
Dim Tienemsj As String
If UserList(userindex).flags.DeseoRecibirMSJ = 1 Then
Tienemsj = "activados"
End If
If UserList(userindex).flags.DeseoRecibirMSJ = 0 Then
Tienemsj = "desactivados"
End If

Call SendUserHitBox(userindex)

Dim ggizipls As Long
For ggizipls = 1 To UserList(userindex).InventorySlots
    If UserList(userindex).Invent.Object(ggizipls).Equipped = 1 Then
      If ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura > 0 And ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).OBJType = eOBJType.otArmadura Then
        UserList(userindex).Char.AuraA = ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura
      ElseIf ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura > 0 And ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).OBJType = eOBJType.otWeapon Then
        UserList(userindex).Char.AuraW = ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura
      ElseIf ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura > 0 And ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).OBJType = eOBJType.otESCUDO Then
        UserList(userindex).Char.AuraE = ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura
      ElseIf ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura > 0 And ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).OBJType = eOBJType.otcASCO Then
        UserList(userindex).Char.AuraC = ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura
      ElseIf ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura > 0 And ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).OBJType = eOBJType.otHerramientas Then
        UserList(userindex).Char.AuraR = ObjData(UserList(userindex).Invent.Object(ggizipls).ObjIndex).Aura
      End If
    End If
Next ggizipls

Call SendUserStats(userindex)
Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)

UserList(userindex).Password = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "SEGURIDAD", "CodeX")
Call MostrarNumUsers

If UserList(userindex).Pos.Map = 94 Or UserList(userindex).Pos.Map = 95 Or UserList(userindex).Pos.Map = 31 Or UserList(userindex).Pos.Map = 32 Or UserList(userindex).Pos.Map = 33 Or UserList(userindex).Pos.Map = 34 Or UserList(userindex).Pos.Map = 72 Or UserList(userindex).Pos.Map = 110 Or UserList(userindex).Pos.Map = 54 Or UserList(userindex).Pos.Map = 8 Or UserList(userindex).Pos.Map = 101 Or UserList(userindex).Pos.Map = 166 Then
    Call SendData(SendTarget.toindex, userindex, 0, "!!Deslogeaste en un lugar sin salida y por ese motivo has sido llevado a tanaris.")
    Call WarpUserChar(userindex, 1, 54, 36, True)
End If

Call SendData(SendTarget.toindex, userindex, 0, "STOPD" & UserList(userindex).flags.Stopped)

Call GRANK_User_Check(Duels, UserList(userindex).Name, UserList(userindex).Stats.DuelosGanados)
Call GRANK_User_Check(Kills, UserList(userindex).Name, UserList(userindex).Stats.UsuariosMatados)
Call GRANK_User_Check(Tournaments, UserList(userindex).Name, UserList(userindex).Stats.TrofOro)
Call GRANK_User_Check(Events, UserList(userindex).Name, UserList(userindex).Stats.MedOro)
Call GRANK_User_Check(Couple, UserList(userindex).Name, UserList(userindex).Stats.ParejasGanadas)

Call GRANK_User_Check(Reputation, UserList(userindex).Name, UserList(userindex).Stats.Reputacione)

Call WarpUserChar(userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, True)

Call SendData(SendTarget.toindex, userindex, 0, "||705")
Call SendData(SendTarget.toindex, userindex, 0, "||706@" & val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant")))
Call SendData(SendTarget.toindex, userindex, 0, "||707")
Call SendData(SendTarget.toindex, userindex, 0, "||708")
Call SendData(SendTarget.toindex, userindex, 0, "||709@" & UserList(userindex).Name)

If Tienemsj = "activados" Then
    Call SendData(SendTarget.toindex, userindex, 0, "||710")
ElseIf Tienemsj = "desactivados" Then
    Call SendData(SendTarget.toindex, userindex, 0, "||711")
End If

If UserList(userindex).GuildIndex > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Norte@" & CastilloNorte)
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Sur@" & CastilloSur)
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Este@" & CastilloEste)
        Call SendData(SendTarget.toindex, userindex, 0, "||430@Oeste@" & CastilloOeste)
        Call SendData(SendTarget.toindex, userindex, 0, "||431@" & Fortaleza)
        Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "||713@" & UserList(userindex).Name)
End If

UserList(userindex).flags.TiempoOnlineHoy = 0
UserList(userindex).flags.TiempoParaCofres = 0
Call SendData(SendTarget.toindex, userindex, 0, "IVX" & UserList(userindex).InventorySlots)

'Seteamos en falso antes de arrancar a revisar si tiene una estrella.
'UserList(userindex).flags.PosRanking = 0
'UserList(userindex).flags.TieneRanking = False
'For ggizipls = 1 To 3
'    If (UCase$(UserList(userindex).Name) = UCase$(Estrella.TOPTorneos(ggizipls))) Or (UCase$(UserList(userindex).Name) = UCase$(Estrella.TOPDuelos(ggizipls))) Or (UCase$(UserList(userindex).Name) = UCase$(Estrella.TOPFrags(ggizipls))) Or (UCase$(UserList(userindex).Name) = UCase$(Estrella.TOPRondas(ggizipls))) Or (UCase$(UserList(userindex).Name) = UCase$(Estrella.TOPParejas(ggizipls))) Or (UCase$(UserList(userindex).Name) = UCase$(Estrella.TOPEvents(ggizipls))) Then
'        UserList(userindex).flags.TieneRanking = True
'        If UserList(userindex).flags.PosRanking = 0 Then UserList(userindex).flags.PosRanking = ggizipls
'        SendUserRank (userindex)
' '       Exit For
'   End If
'Next ggizipls

'Le damos una misión por día.
Dim FechaActual As String, tDiaAct As Byte, tMesAct As Byte, tAñoAct As Integer
FechaActual = ReadField(1, Date, Asc("/"))
If UserList(userindex).Misiones.Dia <> FechaActual Then
    Call OtorgarMisionDiaria(userindex)
Else
    If UserList(userindex).Misiones.Completada = 0 Then Call SendData(SendTarget.toindex, userindex, 0, "MISJ" & UserList(userindex).Misiones.NumeroMision & "," & UserList(userindex).Misiones.ConteoUser & "," & MisionesDiarias(UserList(userindex).Misiones.NumeroMision).Cantidad)
End If

'Vence premium¿?
If UserList(userindex).flags.EsPremium = 1 Then
    tDiaAct = ReadField(1, Date, Asc("/"))
    tMesAct = ReadField(2, Date, Asc("/"))
    tAñoAct = ReadField(3, Date, Asc("/"))
    If (tAñoAct > ReadField(3, UserList(userindex).flags.VencePremium, Asc("/")) Or tMesAct > ReadField(2, UserList(userindex).flags.VencePremium, Asc("/"))) Or (tAñoAct = ReadField(3, UserList(userindex).flags.VencePremium, Asc("/")) And tMesAct = ReadField(2, UserList(userindex).flags.VencePremium, Asc("/")) And tDiaAct >= ReadField(3, UserList(userindex).flags.VencePremium, Asc("/"))) Then
        If TieneObjetos(1553, 1, userindex) Then Call QuitarObjetos(1553, 1, userindex)
        If TieneObjetos(1554, 1, userindex) Then Call QuitarObjetos(1554, 1, userindex)
        UserList(userindex).flags.EsPremium = 0
        SendUserVariant (userindex)
    End If
End If

'Cambiamos el head en caso de que la tenga bug.
            Dim MinEleccion As Integer
            Dim MaxEleccion As Integer
          
If UserList(userindex).flags.Muerto = 0 And UserList(userindex).flags.Navegando = 0 Then
            Select Case UCase$(UserList(userindex).Genero)
            
                Case "HOMBRE"
            
                Select Case UCase$(UserList(userindex).Raza)
                
                    Case "HUMANO"
                        MaxEleccion = 30
                        MinEleccion = 1
                    
                    Case "ELFO"
                        MaxEleccion = 112
                        MinEleccion = 101
                                    
                    Case "ELFO OSCURO"
                        MaxEleccion = 209
                        MinEleccion = 202
                                    
                    Case "ENANO"
                        MaxEleccion = 305
                        MinEleccion = 301
                                    
                    Case "GNOMO"
                        MaxEleccion = 406
                        MinEleccion = 401
                                    
                    Case Else
                        MaxEleccion = 30
                        MinEleccion = 30
                                
                End Select
                    
            Case "MUJER"
               
                Select Case UCase$(UserList(userindex).Raza)
                
                    Case "HUMANO"
                        MaxEleccion = 76
                        MinEleccion = 70
                                    
                    Case "ELFO"
                        MaxEleccion = 176
                        MinEleccion = 170
                                    
                    Case "ELFO OSCURO"
                        MaxEleccion = 280
                        MinEleccion = 270
                                    
                    Case "GNOMO"
                        MaxEleccion = 474
                        MinEleccion = 470
                                    
                    Case "ENANO"
                        MaxEleccion = 373
                        MinEleccion = 370
                                
                    Case Else
                        MaxEleccion = 70
                        MinEleccion = 70
                End Select
            End Select
            
            
            If UserList(userindex).Char.Head < MinEleccion Or UserList(userindex).Char.Head > MaxEleccion Then
                UserList(userindex).Char.Head = MinEleccion
                UserList(userindex).OrigChar.Head = MinEleccion
                Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                Exit Sub
            End If
End If
'Cambiamos la cabeza en caso de que la tenga bug.



End Sub
Sub ResetFacciones(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .NeutralesMatados = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
    End With
End Sub

Sub ResetContadores(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .COMCounter = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0

        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex)
        .Name = ""
        .modName = ""
        .Desc = ""
        .DescRM = ""
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = ""
        .RDBuffer = ""
        .clase = ""
        .email = ""
        .Genero = ""
        .Hogar = ""
        .Raza = ""

        .RandKey = 0
        .PrevCheckSum = 0
        .PacketNumber = 0
        
        With .Stats
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .FIT = 0
            .SkillPts = 0
        End With
      
With .Faccion
.CriminalesMatados = 0
.NeutralesMatados = 0
.CiudadanosMatados = 0
End With
    End With
End Sub
Sub ResetGuildInfo(ByVal userindex As Integer)
    If UserList(userindex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(userindex, UserList(userindex).GuildIndex)
    End If
    UserList(userindex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************
    With UserList(userindex).flags
        .Comerciando = False
        .CuentaBancaria = ""
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = ""
        .Hambre = 0
        .Sed = 0
        .ModoCombate = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = PlayerType.User
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .Hechizo = 0
        .PertAlCons = 0
        .PertAlConsCaos = 0
        .Silenciado = 0
        .CentinelaOK = False
    End With
    With UserList(userindex).ConsejoInfo
.PertAlCons = 0
.PertAlConsCaos = 0
End With
End Sub

Sub ResetUserSpells(ByVal userindex As Integer)
    Dim loopC As Long
    For loopC = 1 To MAXUSERHECHIZOS
        UserList(userindex).Stats.UserHechizos(loopC) = 0
    Next loopC
End Sub

Sub ResetUserPets(ByVal userindex As Integer)
    Dim loopC As Long
    
    UserList(userindex).NroMacotas = 0
        
    For loopC = 1 To MAXMASCOTAS
        UserList(userindex).MascotasIndex(loopC) = 0
        UserList(userindex).MascotasType(loopC) = 0
    Next loopC
End Sub

Sub ResetUserBanco(ByVal userindex As Integer)
    Dim loopC As Long
    
    For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(userindex).BancoInvent.Object(loopC).Amount = 0
          UserList(userindex).BancoInvent.Object(loopC).Equipped = 0
          UserList(userindex).BancoInvent.Object(loopC).ObjIndex = 0
    Next loopC
    
    UserList(userindex).BancoInvent.NroItems = 0
End Sub
Sub ResetUserSlot(ByVal userindex As Integer)

Dim UsrTMP As User

Set UserList(userindex).CommandsBuffer = Nothing


Set UserList(userindex).ColaSalida = Nothing
UserList(userindex).SockPuedoEnviar = False
UserList(userindex).ConnIDValida = False
UserList(userindex).ConnID = -1

Call ResetFacciones(userindex)
Call ResetContadores(userindex)
Call ResetCharInfo(userindex)
Call ResetBasicUserInfo(userindex)
Call ResetGuildInfo(userindex)
Call ResetUserFlags(userindex)
Call LimpiarInventario(userindex)
Call ResetUserSpells(userindex)
Call ResetUserPets(userindex)
Call ResetUserBanco(userindex)

UserList(userindex) = UsrTMP

End Sub


Sub CloseUser(ByVal userindex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo Errhandler

Dim n As Integer
Dim X As Integer
Dim Y As Integer
Dim loopC As Integer
Dim Map As Integer
Dim Name As String
Dim Raza As String
Dim clase As String
Dim i As Integer

Dim aN As Integer

aN = UserList(userindex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
UserList(userindex).flags.AtacadoPorNpc = 0

Map = UserList(userindex).Pos.Map
X = UserList(userindex).Pos.X
Y = UserList(userindex).Pos.Y
Name = UCase$(UserList(userindex).Name)
Raza = UserList(userindex).Raza
clase = UserList(userindex).clase

UserList(userindex).Char.FX = 0
UserList(userindex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
'quitamos la particula
UserList(userindex).Char.Particula = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFF" & UserList(userindex).Char.CharIndex & "," & 0)

UserList(userindex).flags.UserLogged = False
UserList(userindex).Counters.Saliendo = False

UserList(userindex).UltimoLogeo = Date

If userindex = GranPoder Then
    GranPoder = 0
    Call OtorgarGranPoder(0)
End If

If UserList(userindex).flags.Voto = True Then
    If UserList(userindex).flags.VotoPorLaOpcion(1) = 1 Then
        Votos(1) = Votos(1) - 1
    ElseIf UserList(userindex).flags.VotoPorLaOpcion(2) = 1 Then
        Votos(2) = Votos(2) - 1
    ElseIf UserList(userindex).flags.VotoPorLaOpcion(3) = 1 Then
        Votos(3) = Votos(3) - 1
    ElseIf UserList(userindex).flags.VotoPorLaOpcion(4) = 1 Then
        Votos(4) = Votos(4) - 1
    ElseIf UserList(userindex).flags.VotoPorLaOpcion(5) = 1 Then
        Votos(5) = Votos(5) - 1
    End If
End If

'Le devolvemos el body y head originales
If UserList(userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(userindex)

' Grabamos el personaje del usuario
Call SaveUser(userindex, CharPath & Name & ".chr")
'usado para borrar Pjs
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "Logged", "0")

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToMapButIndex, userindex, Map, "QDL" & UserList(userindex).Char.CharIndex)
End If

'Borrar el personaje
If UserList(userindex).Char.CharIndex > 0 Then
    Call EraseUserChar(SendTarget.ToMap, userindex, Map, userindex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
    End If
Next i

    Dim atodosxdjeje As Integer
    For atodosxdjeje = 1 To LastUser
        Call FriendDisconnect(atodosxdjeje, Name)
    Next atodosxdjeje

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If UserList(userindex).flags.ConsultaEnviada = True Then
        For n = UserList(userindex).flags.NumeroConsulta To MensajesNumber
            If MensajesNumber >= UserList(userindex).flags.NumeroConsulta Then
                    MensajesSOS(n).Autor = MensajesSOS(n + 1).Autor
                    MensajesSOS(n).Tipo = MensajesSOS(n + 1).Tipo
                    MensajesSOS(n).Contenido = MensajesSOS(n + 1).Contenido
                
                MensajesSOS(n + 1).Autor = ""
                MensajesSOS(n + 1).Tipo = ""
                MensajesSOS(n + 1).Contenido = ""
            End If
        Next n
        
            MensajesNumber = MensajesNumber - 1
            
            Dim dataSOS As String
            dataSOS = MensajesNumber & ","
            
            For loopC = 1 To MensajesNumber
                dataSOS = dataSOS & MensajesSOS(loopC).Tipo & "-" & MensajesSOS(loopC).Autor & "-" & MensajesSOS(loopC).Contenido & ","
            Next loopC
            
            Call SendData(SendTarget.ToAdmins, 0, 0, "ZSOS" & dataSOS)
            UserList(userindex).flags.ConsultaEnviada = False
            UserList(userindex).flags.NumeroConsulta = 0
End If

  If UserList(userindex).flags.EnCvc = True Then
                UserList(userindex).flags.EnCvc = False
                WarpUserChar userindex, 1, 50, 50, True
            End If
            
            If UserList(userindex).EnCvc Then
            'Dim ijaji As Integer
            'For ijaji = 1 To LastUser
                With UserList(userindex)
                    If Guilds(.GuildIndex).GuildName = Nombre1 Then
                        If .EnCvc = True Then
                                modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 - 1
                                UserList(userindex).EnCvc = False
                                If modGuilds.UsuariosEnCvcClan1 = 0 And CvcFunciona = True Then
                                    Call SendData(SendTarget.ToAll, userindex, 0, "||85@" & Nombre2 & "@" & Nombre1)
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                End If
                         End If
                     End If
                     
                    If Guilds(.GuildIndex).GuildName = Nombre2 Then
                        If .EnCvc = True Then
                                modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 - 1
                                UserList(userindex).EnCvc = False
                                If modGuilds.UsuariosEnCvcClan2 = 0 And CvcFunciona = True Then
                                    Call SendData(SendTarget.ToAll, userindex, 0, "||85@" & Nombre1 & "@" & Nombre2)
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                End If
                        End If
                    End If
                End With
            'Next ijaji
    End If

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

Call ResetUserSlot(userindex)

Call MostrarNumUsers

Exit Sub

Errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)


End Sub


Sub HandleData(ByVal userindex As Integer, ByVal rData As String)

'
' ATENCION: Cambios importantes en HandleData.
' =========
'
'           La funcion se encuentra dividida en 2,
'           una parte controla los comandos que
'           empiezan con "/" y la otra los comanos
'           que no. (Basado en la idea de Barrin)
'

'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
On Error GoTo ErrorHandler:
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'
'Ah, no me queres hacer caso ? Entonces
'atenete a las consecuencias!!
'
    Dim CadenaOriginal As String
    
    Dim loopC As Integer
    Dim nPos As WorldPos
    Dim tStr As String
    Dim iStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim tIndex As Integer
    Dim tName As String
    Dim tMessage As String
    Dim AuxInd As Integer
    Dim Arg1 As String
    Dim Arg2 As String
    Dim Arg3 As String
    Dim Arg4 As String
    Dim Ver As String
    Dim encpass As String
    Dim Pass As String
    Dim mapa As Integer
    Dim Name As String
    Dim ind
    Dim n As Integer
    Dim wpaux As WorldPos
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim T() As String
    Dim i As Integer
    
    Dim sndData As String
    Dim ClientChecksum As String
    Dim ServerSideChecksum As Long
    Dim IdleCountBackup As Long

    
    '¿Tiene un indece valido?
    If userindex <= 0 Then
        Call CloseSocket(userindex)
        Exit Sub
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    IdleCountBackup = UserList(userindex).Counters.IdleCount
    UserList(userindex).Counters.IdleCount = 0
   
If Not UserList(userindex).flags.UserLogged Then

    Select Case Left$(rData, 6)
         
        'Declaraciones
        Dim CuentaName As String
        Dim mailName As String
        Dim preguntapedida As String
        Dim Respuesta As String
        
        Case "REECUH" 'Segunda parte de recuperar cuenta.
        rData = Right$(rData, Len(rData) - 6)
        CuentaName = ReadField(1, rData, Asc(","))
        Respuesta = ReadField(2, rData, Asc(","))
        
        If Respuesta <> GetVar(App.Path & "\Accounts\" & CuentaName & ".act", CuentaName, "Respuesta") Then
                Call SendData(SendTarget.toindex, userindex, 0, "ERRLa respuesta es incorrecta.")
                CloseSocket (userindex)
            Exit Sub
        Else
            Dim PasswordGen As Integer
            PasswordGen = RandomNumber(100, 999)
            Call SendData(SendTarget.toindex, userindex, 0, "EROHas recuperado la cuenta, utiliza la contraseña " & PasswordGen & " para poder logearte.")
            Call WriteVar(App.Path & "\Accounts\" & CuentaName & ".act", CuentaName, "Password", PasswordGen)
            Exit Sub
        End If
      
       
        Case "RECCUU" 'primera parte de recuperar cuenta.
        rData = Right$(rData, Len(rData) - 6)
        CuentaName = ReadField(1, rData, Asc(","))
        mailName = ReadField(2, rData, Asc(","))
        preguntapedida = GetVar(App.Path & "\Accounts\" & CuentaName & ".act", CuentaName, "Pregunta")
        
            If Not CuentaExiste(CuentaName) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRLa cuenta no existe.")
                    CloseSocket (userindex)
            Exit Sub
            End If
            
        If mailName <> GetVar(App.Path & "\Accounts\" & CuentaName & ".act", CuentaName, "mail") Then
            Call SendData(SendTarget.toindex, userindex, 0, "ERRMail invalido.")
            CloseSocket (userindex)
            Exit Sub
        End If

        Call SendData(toindex, userindex, 0, "PEDPRE" & preguntapedida)
        
        Exit Sub
                           
        Case "REPASS"
                        ' - Cambio de Pass
                        rData = Right$(rData, Len(rData) - 6)
                        Dim PassName As String
                        Dim PassVieja As String
                        Dim PassNueva As String
                        Dim RePassNueva As String
                       
                        PassName = ReadField(1, rData, Asc(","))
                        PassVieja = ReadField(2, rData, Asc(","))
                        PassNueva = ReadField(3, rData, Asc(","))
                        RePassNueva = ReadField(4, rData, Asc(","))
                        
                        If PassNueva = PassVieja Then
                            Call SendData(SendTarget.toindex, userindex, 0, "ERONo puedes volver a utilizar la misma contraseña.")
                        Exit Sub
                        End If
                       
                        If Len(PassNueva) <= 3 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "EROLa contraseña debe tener un minimo de 3 caracteres.")
                        Exit Sub
                        End If
                       
                        If PassVieja <> GetVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "password") Then
                            Call SendData(SendTarget.toindex, userindex, 0, "EROLa Password actual que nos proporciono, no coincide con la del registro.")
                        Exit Sub
                        End If
                        
                        If (UCase$(GetVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "AntiguaPW")) = UCase$(PassNueva)) Or (UCase$(GetVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "AntiguaPW2")) = UCase$(PassNueva)) Then
                            Call SendData(SendTarget.toindex, userindex, 0, "EROEsa password es la que tenias anteriormente, no puedes volver a reutilizarla (esto es por tu seguridad).")
                        Exit Sub
                        End If
                        
                        If GetVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "CambioPW") = 0 Then
                            Call WriteVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "CambioPW", "1")
                            Call WriteVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "AntiguaPW", PassVieja)
                        ElseIf GetVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "CambioPW") = 1 Then
                            Call WriteVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "CambioPW", "2")
                            Call WriteVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "AntiguaPW2", PassVieja)
                        End If
                        
                        Call WriteVar(App.Path & "\Accounts\" & PassName & ".act", PassName, "Password", PassNueva)
                        Call SendData(SendTarget.toindex, userindex, 0, "EROLa password de su cuenta fue cambiada con exito. Ahora para logear debera de utilizar la nueva.")
                        
                        Call LogPassw("[CUENTA: " & PassName & "] Clave anterior: " & PassVieja & " - Clave nueva: " & PassNueva & "")
                       
                Exit Sub
        
        Case "OOLOGI"
                rData = Right$(rData, Len(rData) - 6)
            Dim Personaje As String
            Dim Acc As String
            
            Personaje = ReadField(1, rData, Asc(","))
            Acc = ReadField(2, rData, Asc(","))
               
               If Not PersonajeExiste(Personaje) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERREl personaje no existe.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
                
                If Not BANCheck(Personaje) Then
                    Call ConnectUser(userindex, Personaje, Acc, ReadField(3, rData, Asc(",")))
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRSe te ha prohibido la entrada a Tierras Sagradas AO debido a tu mal comportamiento.")
                End If
            Exit Sub
                
        
        Case "ALOGIN"
                rData = Right$(rData, Len(rData) - 6)

                'If ReadField(3, rData, 44) <> GetVar(IniPath & "Server.ini", "INIT", "ClientVersion") Then
                '    Call SendData(SendTarget.toindex, UserIndex, 0, "ERRTu cliente está desactualizado, deberás ejecutar el launcher y esperar que se descargue la nueva actualización para volver a ingresar.")
                '    Call CloseSocket(UserIndex, True)
                 ''   Exit Sub
                'End If
                
                            If PasoHD = False Then
                                UserList(userindex).hd = rData
                                Call WriteVar(App.Path & "\Charfile\" & UserList(userindex).Name & ".CHR", "INIT", "LastHD", rData)
                                Call SendData(SendTarget.toindex, userindex, 0, "ERRTu PC se encuentra bajo Tolerancia 0.")
                                Debug.Print ">>>>CLIENTE IP: " & UserList(userindex).ip & " - TIENE TOLERANCIA 0, QUISO ENTRAR, PERO NO PUDO ;)"
                                Call CloseSocket(userindex)
                             Exit Sub
                            End If
                            
                            UserList(userindex).hd = HDSerialIndex
                
                If Not AsciiValidos(ReadField(1, rData, 44)) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRNombre invalido.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
               
                If Not CuentaExiste(ReadField(1, rData, 44)) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRLa cuenta no existe.")
                    Call CloseSocket(userindex, True)
                    Exit Sub
                End If
                
                Call ConnectAccount(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44))
                Exit Sub
                
        
         Case "TIRDAD"
 
                UserList(userindex).Stats.UserAtributos(1) = 18
                UserList(userindex).Stats.UserAtributos(2) = 18
                UserList(userindex).Stats.UserAtributos(3) = 18
                UserList(userindex).Stats.UserAtributos(4) = 18
                UserList(userindex).Stats.UserAtributos(5) = 18
                
              Call SendData(SendTarget.toindex, userindex, 0, "DADOS" & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion))
                
             
        Exit Sub

        
        Case "NACCNT"
            
                rData = Right$(rData, Len(rData) - 6)
                
                Dim NCuenta As String
                Dim Passw As String
                Dim Mail As String
                Dim pregunta As String
 
                'cuentas
                NCuenta = ReadField(1, rData, Asc(","))
                Passw = ReadField(2, rData, Asc(","))
                Mail = ReadField(3, rData, Asc(","))
                pregunta = ReadField(4, rData, Asc(","))
                Respuesta = ReadField(5, rData, Asc(","))
                
               Call CreateAccount(NCuenta, Passw, Mail, pregunta, Respuesta, userindex)

                
            Exit Sub
            
        Case "KERD22"
        rData = Right$(rData, Len(rData) - 6)
        
        Dim HDSerial As String
        Dim IPPublica As String
        HDSerial = ReadField(1, rData, Asc(","))
        IPPublica = ReadField(2, rData, Asc(","))
        
            If CheckHD(HDSerial) Then
                PasoHD = False
            Exit Sub
            Else
               HDSerialIndex = val(HDSerial)
                PasoHD = True
            End If
            
            'UserList(userindex).ip = IPPublica
        Exit Sub

        Case "THCJXD"

                rData = Right$(rData, Len(rData) - 6)
                
                            If PasoHD = False Then
                                UserList(userindex).hd = rData
                                Call WriteVar(App.Path & "\Charfile\" & UserList(userindex).Name & ".CHR", "INIT", "LastHD", rData)
                                Call SendData(SendTarget.toindex, userindex, 0, "ERRTu PC se encuentra bajo Tolerancia 0.")
                                Debug.Print ">>>>CLIENTE IP: " & UserList(userindex).ip & " - TIENE TOLERANCIA 0, QUISO ENTRAR, PERO NO PUDO ;)"
                                Call CloseSocket(userindex)
                             Exit Sub
                            End If
                            UserList(userindex).hd = HDSerialIndex
                
                Dim passwd As String
                
                Personaje = ReadField(1, rData, Asc(","))
                Acc = ReadField(2, rData, Asc(","))

                tName = ReadField(1, rData, 44)
                    
                    If Not AsciiValidos(Personaje) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "ERRNombre invalido.")
                        Call CloseSocket(userindex, True)
                        Exit Sub
                    End If
                    
                    If Not PersonajeExiste(Personaje) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "ERREl personaje no existe.")
                        Call CloseSocket(userindex, True)
                        Exit Sub
                    End If
                    
                    If Not BANCheck(Personaje) Then
                    
                    If Acc <> UserList(userindex).Accounted Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRError al conectar, intente de nuevo.")
                    Call CloseSocket(userindex, True)
                        Exit Sub
                    End If
                    
                        Call ConnectUser(userindex, Personaje, Acc, ReadField(3, rData, Asc(",")))
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "ERRSe te ha prohibido la entrada a Tierras Sagradas debido a tu mal comportamiento. Consulta a un administrador para saber el motivo de la prohibición.")
                    End If
                Exit Sub

        
        Case "NLOGIN"

                    If PasoHD = False Then
                        Call SendData(SendTarget.toindex, userindex, 0, "ERRTu PC se encuentra bajo Tolerancia 0")
                      Debug.Print ">>>>CLIENTE IP: " & UserList(userindex).ip & " - TIENE TOLERANCIA 0, QUISO ENTRAR, PERO NO PUDO ;)"
                        Call CloseSocket(userindex)
                     Exit Sub
                    End If
                    
                    UserList(userindex).hd = HDSerialIndex

                If PuedeCrearPersonajes = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRLa creacion de personajes en este servidor se ha deshabilitado.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                If ServerSoloGMs <> 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                rData = Right$(rData, Len(rData) - 6)
                    
                    Call ConnectNewUser(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), _
                                        ReadField(6, rData, 44), ReadField(7, rData, 44))
                
                Exit Sub
        End Select
    
    Select Case Left$(rData, 4)
    
     Case "TBRP" ' <<< borra personajes
                        'On Error GoTo ExitErr1:
                        'LwK - borrado de pj
                    rData = Right$(rData, Len(rData) - 4)
                        Dim UserName As String
                        Dim limitPJ As Byte
                        Dim NumPjs As Byte
                        Dim archivo As String
                        Dim passcuent As String
                        
                        UserName = UCase$(ReadField(2, rData, Asc(",")))
                        passcuent = ReadField(3, rData, Asc(","))
                        rData = ReadField(1, rData, Asc(","))
                        archivo = App.Path & "\Accounts\" & UserName & ".act"
                        NumPjs = CByte(val(GetVar(archivo, "PJS", "NumPjs")))
                        
                        If UCase$(passcuent) <> UCase$(GetVar(CharPath & UCase$(rData) & ".chr", "INIT", "Password")) Then
                            Call SendData(SendTarget.toindex, userindex, 0, "ERRPassword incorrecto.")
                            Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
                            Call CloseSocket(userindex)
                        Exit Sub
                        End If
                        
                        If (IsAdministrator(rData)) Or (IsDevelopment(rData)) Or (IsCoordination(rData)) Or (IsTournamentManager(rData)) Or (IsEventManager(rData)) Or (IsUserSupport(rData)) Then
                            Call SendData(SendTarget.toindex, userindex, 0, "ERONo podes borrar gms.")
                            Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
                            Call CloseSocket(userindex)
                          Exit Sub
                        End If
                        
                        If val(GetVar(CharPath & UCase$(rData) & ".chr", "STATS", "ELV")) >= 50 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "ERRNo podes borrar usuarios nivel 50 o superior.")
                            Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
                            Call CloseSocket(userindex)
                        Exit Sub
                        End If
                        
                        If val(GetVar(CharPath & UCase$(rData) & ".chr", "GUILD", "GUILDINDEX")) > 0 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "ERRNo podes borrar usuarios que estén dentro de un clan.")
                            Call SendData(SendTarget.toindex, userindex, 0, "FINOK")
                            Call CloseSocket(userindex)
                        Exit Sub
                        End If
                        
                       
                        For i = 1 To val(GetVar(archivo, "PJS", "NumPjs"))
                            If UCase$(GetVar(archivo, "PJS", "PJ" & i)) = UCase$(rData) Then
                            Call WriteVar(archivo, "PJS", "PJ" & i, "")
                                limitPJ = i + 1
                                BorrarUsuario (rData)
                                If i = 0 Then
                                Exit For
                                Else
                                Call WriteVar(archivo, "PJS", "NumPjs", val(GetVar(archivo, "PJs", "NumPjs")) - 1)
                               
                                Exit For
                            End If
                            End If
                        Next i
                     
                        For i = limitPJ To NumPjs
                            UserName = GetVar(archivo, "PJS", "PJ" & i)
                            Call WriteVar(archivo, "PJS", "PJ" & i, "")
                            Call WriteVar(archivo, "PJS", "PJ" & i - 1, UserName)
                        Next i
                       
        
                Call SendData(SendTarget.toindex, userindex, 0, "EROPersonaje borrado con exito.")
                Call CloseSocket(userindex)
            Exit Sub
            'End If
    End Select

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Si no esta logeado y envia un comando diferente a los
    'de arriba cerramos la conexion.
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call CloseSocket(userindex)
    Exit Sub
End If ' if not user logged


Dim Procesado As Boolean

' bien ahora solo procesamos los comandos que NO empiezan
' con "/".
If Left$(rData, 1) <> "/" Then
    
    Call HandleData_4(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    Call HandleData_1(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    Call HandleData_2(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    Call HandleData_3(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    
' bien hasta aca fueron los comandos que NO empezaban con
' "/". Ahora adiviná que sigue :)
Else
    
    Call HandleData_1(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    Call HandleData_2(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    Call HandleData_3(userindex, rData, Procesado)
    If Procesado Then Exit Sub
    
     If UserList(userindex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.toindex, userindex, 0, "||714")
     End If
End If ' "/"

If UserList(userindex).flags.Privilegios = PlayerType.User Then
    UserList(userindex).Counters.IdleCount = IdleCountBackup
End If

'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(userindex).flags.Privilegios = PlayerType.User Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

'<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<

If UCase$(Left$(rData, 10)) = "/ENCUESTA " Then
rData = Right$(rData, Len(rData) - 10)
Dim Encuestap As String, Op(1 To 5) As String, Nencuesta As String
Encuestap = ReadField(1, rData, Asc("@")) ' Encuesta
Op(1) = ReadField(2, rData, Asc("@")) ' Opcion 1
Op(2) = ReadField(3, rData, Asc("@")) ' Opcion 2
Op(3) = ReadField(4, rData, Asc("@")) ' Opcion 3
Op(4) = ReadField(5, rData, Asc("@")) ' Opcion 4
Op(5) = ReadField(6, rData, Asc("@")) ' Opcion 5
Nencuesta = ReadField(7, rData, Asc("@")) ' Nivel ENCUESTA

If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub

'Reseteamos a la fuerza'
Encuesta = ""
HayEncuesta = False
Opciones(1) = ""
opcion(1) = False
Opciones(2) = ""
opcion(2) = False
Opciones(3) = ""
opcion(3) = False
Opciones(4) = ""
opcion(4) = False
Opciones(5) = ""
opcion(5) = False
Call SendData(SendTarget.ToAll, 0, 0, "BYE")
'Reseteamos a la fuerza'

If UCase$(Op(1)) = "N/A" Or UCase$(Op(2)) = "N/A" Then Exit Sub

If Not IsNumeric(Nencuesta) Then ' Si no son números, es igual a 1
    LvlEncuesta = 1
Else ' o si no, que se ponga lo que se puso en /NEWPOLL
    LvlEncuesta = Nencuesta
End If

'Estos si o si'
Encuesta = Encuestap
HayEncuesta = True
Opciones(1) = Op(1)
opcion(1) = True
Opciones(2) = Op(2)
opcion(2) = True
'Estos si o si

'Aca empezamos a preguntar si se activa o no'
If Not UCase$(Op(3)) = "N/A" Then
Opciones(3) = Op(3)
opcion(3) = True
End If

If Not UCase$(Op(4)) = "N/A" Then
Opciones(4) = Op(4)
opcion(4) = True
End If

If Not UCase$(Op(5)) = "N/A" Then
Opciones(5) = Op(5)
opcion(5) = True
End If
'Aca terminamos de preguntar si se activa o no'

'Manda el mensaje a la consola'
Call SendData(SendTarget.ToAll, 0, 0, "||715@" & Encuesta)
'Manda el mensaje a la consola'
Exit Sub
End If





If UCase$(rData) = "/NEWPOLL" Then
If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
    Call SendData(SendTarget.toindex, userindex, 0, "WEN")
Exit Sub
End If

If UCase$(rData) = "/ENDPOLL" Then

    If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
    
    If HayEncuesta = False Then
        Call SendData(SendTarget.toindex, userindex, 0, "||716")
    Exit Sub
    End If
    
    Call SendData(SendTarget.ToAll, 0, 0, "||717@" & UserList(userindex).Name)
    
    Encuesta = ""
    HayEncuesta = False
    Opciones(1) = ""
    opcion(1) = False
    Opciones(2) = ""
    opcion(2) = False
    Opciones(3) = ""
    opcion(3) = False
    Opciones(4) = ""
    opcion(4) = False
    Opciones(5) = ""
    opcion(5) = False
    
    Votos(1) = 0
    Votos(2) = 0
    Votos(3) = 0
    Votos(4) = 0
    Votos(5) = 0
    
    For i = 1 To LastUser
        UserList(i).flags.Voto = False
        UserList(i).flags.VotoPorLaOpcion(1) = 0
        UserList(i).flags.VotoPorLaOpcion(2) = 0
        UserList(i).flags.VotoPorLaOpcion(3) = 0
        UserList(i).flags.VotoPorLaOpcion(4) = 0
        UserList(i).flags.VotoPorLaOpcion(5) = 0
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, "BYE")
Exit Sub
End If

'/IRCERCA
If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
    Dim indiceUserDestino As Integer
    rData = Right$(rData, Len(rData) - 9) 'obtiene el nombre del usuario
    tIndex = NameIndex(rData)
    
    'Si es dios o Admins no podemos salvo que nosotros también lo seamos
    If UserList(userindex).flags.Privilegios < PlayerType.UserSupport Then Exit Sub
    
    If tIndex <= 0 Then 'existe el usuario destino?
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If

    For tInt = 2 To 5 'esto for sirve ir cambiando la distancia destino
        For i = UserList(tIndex).Pos.X - tInt To UserList(tIndex).Pos.X + tInt
            For DummyInt = UserList(tIndex).Pos.Y - tInt To UserList(tIndex).Pos.Y + tInt
                If (i >= UserList(tIndex).Pos.X - tInt And i <= UserList(tIndex).Pos.X + tInt) And (DummyInt = UserList(tIndex).Pos.Y - tInt Or DummyInt = UserList(tIndex).Pos.Y + tInt) Then
                    If MapData(UserList(tIndex).Pos.Map, i, DummyInt).userindex = 0 And LegalPos(UserList(tIndex).Pos.Map, i, DummyInt) Then
                        Call WarpUserChar(userindex, UserList(tIndex).Pos.Map, i, DummyInt, True)
                        Exit Sub
                    End If
                ElseIf (DummyInt >= UserList(tIndex).Pos.Y - tInt And DummyInt <= UserList(tIndex).Pos.Y + tInt) And (i = UserList(tIndex).Pos.X - tInt Or i = UserList(tIndex).Pos.X + tInt) Then
                    If MapData(UserList(tIndex).Pos.Map, i, DummyInt).userindex = 0 And LegalPos(UserList(tIndex).Pos.Map, i, DummyInt) Then
                        Call WarpUserChar(userindex, UserList(tIndex).Pos.Map, i, DummyInt, True)
                        Exit Sub
                    End If
                End If
            Next DummyInt
        Next i
    Next tInt
    
    Exit Sub
End If

If UCase$(rData) = "/FINALIZAR" Then
 
 If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
 
If Hay_Torneo = True Then
UsuariosEnTorneo = 0
 
Dim tornein As Long
For tornein = 1 To LastUser
If UserList(tornein).flags.EnTorneo = 1 Then
UserList(tornein).flags.EnTorneo = 0
End If
If UserList(tornein).flags.NumTorneo = 1 Then
UserList(tornein).flags.NumTorneo = 0
End If
Next tornein
 
Call SendData(SendTarget.ToAll, 0, 0, "||718")
 
Hay_Torneo = False
TModalidad = "0"
PuntosPremios = 0
End If
 
  Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/LINK " Then
    rData = Right$(rData, Len(rData) - 6)
    tStr = ReadField(1, rData, Asc("@")) '¿A QUIEN?
    Name = ReadField(2, rData, Asc("@")) 'LINK
    
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    If UCase$(UserList(userindex).Name) <> "KING" And UCase$(tStr) = "KING" Then Exit Sub
    
If UCase$(tStr) = "TODOS" Then
    SendData SendTarget.ToAll, 0, 0, "TAL" & Name
ElseIf UCase$(tStr) = "GMS" Then
    SendData SendTarget.ToAdmins, 0, 0, "TAL" & Name
Else
    tIndex = NameIndex(tStr)
  If tIndex <= 0 Then Exit Sub
  
SendData SendTarget.toindex, tIndex, 0, "TAL" & Name
End If

Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/STOP " Then
rData = Right$(rData, Len(rData) - 6)
tIndex = NameIndex(rData)
    
    If UCase$(UserList(tIndex).Name) = "KING" Then Exit Sub

    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||719")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Stopped = 1 Then
        UserList(tIndex).flags.Stopped = 0
        Call SendData(SendTarget.toindex, userindex, 0, "||720")
    Else
        UserList(tIndex).flags.Stopped = 1
        Call SendData(SendTarget.toindex, userindex, 0, "||721")
    End If
    
    Call SendData(SendTarget.toindex, tIndex, 0, "STOPD" & UserList(tIndex).flags.Stopped)
    
    
    Call LogGM(UserList(userindex).Name, "/STOP " & UserList(tIndex).Name, False)
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/STOPOFF " Then
rData = Right$(rData, Len(rData) - 9)
    
If UCase$(rData) = "KING" Then Exit Sub

    If GetVar(CharPath & rData & ".chr", "FLAGS", "STOP") = 1 Then
      Call WriteVar(CharPath & rData & ".chr", "FLAGS", "STOP", 0)
      Call SendData(SendTarget.toindex, userindex, 0, "||720")
    Else
      Call WriteVar(CharPath & rData & ".chr", "FLAGS", "STOP", 1)
      Call SendData(SendTarget.toindex, userindex, 0, "||721")
    End If
    
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/PELEAR" Then
    rData = Right$(rData, Len(rData) - 7)
    
If Hay_Torneo = False Then
    Call SendData(SendTarget.toindex, userindex, 0, "||722")
Exit Sub
End If

If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
   
If UserList(userindex).flags.TargetUser = userindex Then
 Call ResetearPeleas
 Call SendData(SendTarget.toindex, userindex, 0, "||723")
Exit Sub
End If
 
If UserList(userindex).flags.TargetUser = 0 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||9")
Exit Sub
End If
 
   
  If val(TModalidad) = 1 Then
    Call Pelear1vs1(userindex)
  ElseIf val(TModalidad) = 2 Then
   Call Pelear2vs2(userindex)
  ElseIf val(TModalidad) = 3 Then
   Call Pelear3vs3(userindex)
  ElseIf val(TModalidad) = 4 Then
   Call Pelear4vs4(userindex)
  Else
   Call SendData(SendTarget.toindex, userindex, 0, "||722")
   Exit Sub
  End If
   
    Exit Sub
End If

If UCase(Left(rData, 14)) = "/DESCALIFICAR " Then
rData = Right$(rData, Len(rData) - 14)
Dim des As String
des = NameIndex(rData)

    If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub

    If UserList(des).flags.EnTorneo = 1 Then
            UserList(des).flags.EnTorneo = 0
            
            For i = 1 To LastUser
                If UserList(i).flags.NumTorneo > UserList(des).flags.NumTorneo Then
                    UserList(i).flags.NumTorneo = UserList(i).flags.NumTorneo - 1
                End If
            Next i
            
            UserList(des).flags.NumTorneo = 0
            UsuariosEnTorneo = UsuariosEnTorneo - 1
            
            SendData SendTarget.ToAll, 0, 0, "||724@" & UserList(des).Name
            Call WarpUserChar(des, 1, 50, 50)
    Else
            SendData SendTarget.toindex, 0, 0, "||725"
    End If

Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/PLATA " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
Dim trofeosplata As Obj
trofeosplata.ObjIndex = 896
trofeosplata.Amount = 1

 If Not tIndex > 0 Then Exit Sub
 
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then Exit Sub

    If Not MeterItemEnInventario(tIndex, trofeosplata) Then
        Call TirarItemAlPiso(UserList(tIndex).Pos, trofeosplata)
    End If
    
    Call SendData(SendTarget.ToAll, userindex, 0, "||728@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
    UserList(tIndex).Stats.TrofPlata = UserList(tIndex).Stats.TrofPlata + 1
    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + 100
    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "TrofPlata", UserList(tIndex).Stats.TrofPlata)
    
    Call SendData(ToAll, userindex, 0, "||729@" & UserList(tIndex).Name & "@" & UserList(tIndex).Stats.TrofPlata)
    
If PuntosPremios <> 0 Then
    Call SendData(toindex, tIndex, 0, "||57@" & PuntosPremios)
    Call AgregarPuntos(tIndex, PuntosPremios)
    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "PuntosTorneo", UserList(tIndex).Stats.PuntosTorneo)
End If

    Call LogGM(UserList(userindex).Name, "Le entregó un trofeo de Plata a " & UserList(tIndex).Name, False)
Exit Sub
End If
 
If UCase$(Left$(rData, 8)) = "/BRONCE " Then
    rData = Right$(rData, Len(rData) - 8)
    tIndex = NameIndex(rData)
    
    Dim trofeosbronce As Obj
    trofeosbronce.ObjIndex = 897
    trofeosbronce.Amount = 1

    If Not tIndex > 0 Then Exit Sub
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then Exit Sub
 
    If Not MeterItemEnInventario(tIndex, trofeosbronce) Then
    Call TirarItemAlPiso(UserList(tIndex).Pos, trofeosbronce)
    End If
    
    Call SendData(SendTarget.ToAll, userindex, 0, "||726@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
    UserList(tIndex).Stats.TrofBronce = UserList(tIndex).Stats.TrofBronce + 1
    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + 50
    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "TrofBronce", UserList(tIndex).Stats.TrofBronce)
    Call SendData(ToAll, userindex, 0, "||727@" & UserList(tIndex).Name & "@" & UserList(tIndex).Stats.TrofBronce)

If PuntosPremios <> 0 Then
    Call SendData(toindex, tIndex, 0, "||57@" & PuntosPremios)
    Call AgregarPuntos(tIndex, PuntosPremios)
    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "PuntosTorneo", UserList(tIndex).Stats.PuntosTorneo)
End If
    
    Call LogGM(UserList(userindex).Name, "Le entregó un trofeo de Bronce a " & UserList(tIndex).Name, False)
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/MEDALLA " Then
    rData = Right$(rData, Len(rData) - 9)
    tIndex = NameIndex(rData)
Dim medallaoro As Obj
medallaoro.Amount = 1
medallaoro.ObjIndex = 1025

 If Not tIndex > 0 Then Exit Sub
 If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
 If UserList(tIndex).flags.Privilegios > PlayerType.User Then Exit Sub

If TModalidad <> 5 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||722")
Exit Sub
End If

    If Not MeterItemEnInventario(tIndex, medallaoro) Then
        Call TirarItemAlPiso(UserList(tIndex).Pos, medallaoro)
    End If
    
    Call SendData(ToAll, userindex, 0, "||732@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
    UserList(tIndex).Stats.MedOro = UserList(tIndex).Stats.MedOro + 1
    Call GRANK_User_Check(Events, UserList(tIndex).Name, UserList(tIndex).Stats.MedOro)
    UserList(tIndex).Stats.Reputacione = UserList(tIndex).Stats.Reputacione + PuntosPremios

    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "MedOro", UserList(tIndex).Stats.MedOro)
    
    Call SendData(ToAll, userindex, 0, "||733@" & UserList(tIndex).Name & "@" & UserList(tIndex).Stats.MedOro)
    Call SendData(toindex, tIndex, 0, "||57@" & PuntosPremios)
    Call AgregarPuntos(tIndex, PuntosPremios)
    
    Call LogGM(UserList(userindex).Name, "Le entregó una medalla de Oro a " & UserList(tIndex).Name, False)
    
    If MisionesDiarias(UserList(tIndex).Misiones.NumeroMision).Tipo = 8 Then
        If UserList(tIndex).Misiones.ConteoUser < MisionesDiarias(UserList(tIndex).Misiones.NumeroMision).Cantidad Then UserList(tIndex).Misiones.ConteoUser = UserList(tIndex).Misiones.ConteoUser + 1
    End If
    
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/MVP " Then
rData = Right$(rData, Len(rData) - 5)


If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub


SendData SendTarget.ToAll, 0, 0, "||734" & UserList(userindex).Pos.Map

Dim npc As Integer, PosNpc As WorldPos

Select Case rData

Case 1
npc = 942

Call SpawnNpc(npc, UserList(userindex).Pos, False, False)

Case 2
npc = 616

Call SpawnNpc(npc, UserList(userindex).Pos, False, False)

End Select

Exit Sub
End If
 
If UCase$(Left$(rData, 5)) = "/ORO " Then
    rData = Right$(rData, Len(rData) - 5)
    tIndex = NameIndex(rData)
Dim trofeosoro As Obj
trofeosoro.Amount = 1
trofeosoro.ObjIndex = 895

 If Not tIndex > 0 Then Exit Sub
 If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
 If UserList(tIndex).flags.Privilegios > PlayerType.User Then Exit Sub

    If Not MeterItemEnInventario(tIndex, trofeosoro) Then
        Call TirarItemAlPiso(UserList(tIndex).Pos, trofeosoro)
    End If
    
    Call SendData(ToAll, userindex, 0, "||730@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
    UserList(tIndex).Stats.TrofOro = UserList(tIndex).Stats.TrofOro + 1
    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "TrofOro", UserList(tIndex).Stats.TrofOro)
    Call SendData(ToAll, userindex, 0, "||731@" & UserList(tIndex).Name & "@" & UserList(tIndex).Stats.TrofOro)
    
If PuntosPremios <> 0 Then
    Call SendData(toindex, tIndex, 0, "||57@" & PuntosPremios)
    Call AgregarPuntos(tIndex, PuntosPremios)
    Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "PuntosTorneo", UserList(tIndex).Stats.PuntosTorneo)
End If

    UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + 250
    Call LogGM(UserList(userindex).Name, "Le entregó un trofeo de Oro a " & UserList(tIndex).Name, False)
    Call GRANK_User_Check(Tournaments, UserList(userindex).Name, UserList(userindex).Stats.TrofOro)
    
    If MisionesDiarias(UserList(tIndex).Misiones.NumeroMision).Tipo = 8 Then
        If UserList(tIndex).Misiones.ConteoUser < MisionesDiarias(UserList(tIndex).Misiones.NumeroMision).Cantidad Then UserList(tIndex).Misiones.ConteoUser = UserList(tIndex).Misiones.ConteoUser + 1
    End If
    
Exit Sub
End If

'¿Donde esta?
If UCase$(Left$(rData, 7)) = "/DONDE " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    If UserList(tIndex).flags.Privilegios >= PlayerType.EventManager Then Exit Sub
        Call SendData(SendTarget.toindex, userindex, 0, "||735@" & UserList(tIndex).Name & "@" & UserList(tIndex).Pos.Map & "@" & UserList(tIndex).Pos.X & "@" & UserList(tIndex).Pos.Y)
        Call LogGM(UserList(userindex).Name, "/Donde " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/NENE " Then
    rData = Right$(rData, Len(rData) - 6)

    If MapaValido(val(rData)) Then
        Dim NpcIndex As Integer
            Dim ContS As String


            ContS = ""
        For NpcIndex = 1 To LastNPC

        '¿esta vivo?
        If Npclist(NpcIndex).flags.NPCActive _
                And Npclist(NpcIndex).Pos.Map = val(rData) _
                    And Npclist(NpcIndex).Hostile = 1 And _
                        Npclist(NpcIndex).Stats.Alineacion = 2 Then
                       ContS = ContS & Npclist(NpcIndex).Name & ", "

        End If

        Next NpcIndex
                If ContS <> "" Then
                    ContS = Left(ContS, Len(ContS) - 2)
                Else
                    ContS = "No hay NPCS"
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||736@" & ContS)
                Call LogGM(UserList(userindex).Name, "Numero enemigos en mapa " & rData, False)
    End If
    Exit Sub
End If

    If UCase$(rData) = "/RESMAP" Then
        rData = Right$(rData, Len(rData) - 7)
        Dim RevivirMap As Integer
        
        If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
        
        For RevivirMap = 1 To LastUser
                If UserList(RevivirMap).Pos.Map = UserList(userindex).Pos.Map Then
                    tStr = UserList(RevivirMap).Pos.Map
                    
                        If UserList(RevivirMap).flags.Muerto = 1 Then
                            Call RevivirUsuario(RevivirMap)
                            Call DarCuerpoDesnudo(RevivirMap)
                         End If
                End If
        Next RevivirMap
        
            Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "||696")
         Exit Sub
        End If
    
    If UCase$(rData) = "/TELEPLOC" Then
        Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
        Call LogGM(UserList(userindex).Name, "/TELEPLOC a x:" & UserList(userindex).flags.TargetX & " Y:" & UserList(userindex).flags.TargetY & " Map:" & UserList(userindex).Pos.Map, False)
        Exit Sub
    End If

    If UCase$(Left$(rData, 7)) = "/CHEAT " Then
        rData = Right$(rData, Len(rData) - 7)
        tIndex = NameIndex(rData)
        
            If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
        
            If tIndex <= 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||196")
            Else
            
                If UCase$(UserList(tIndex).Name) = "KING" Then Exit Sub
            
                Call SendData(SendTarget.toindex, tIndex, 0, "PCCP" & userindex)
                Call SendData(SendTarget.toindex, tIndex, 0, "PCGR" & userindex)
            
                Call LogGM(UserList(userindex).Name, "/CHEAT a " & rData, False)
            End If
        Exit Sub
    End If

'Teleportar
If UCase$(Left$(rData, 7)) = "/TELEP " Then
    rData = Right$(rData, Len(rData) - 7)
    mapa = val(ReadField(2, rData, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rData, 32)
    If Name = "" Then Exit Sub
    
    If UCase$(Name) <> "YO" Then
        If UserList(userindex).flags.Privilegios <= PlayerType.EventManager Then
            Exit Sub
        End If
        tIndex = NameIndex(Name)
    Else
        tIndex = userindex
    End If
    X = val(ReadField(3, rData, 32))
    Y = val(ReadField(4, rData, 32))
    If Not InMapBounds(mapa, X, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(Name) = "KING" Then Exit Sub
    
    Call WarpUserChar(tIndex, mapa, X, Y, True)
    Call SendData(SendTarget.toindex, tIndex, 0, "||651@" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, False)
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
    rData = Right$(rData, Len(rData) - 11)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    If UserList(tIndex).flags.Silenciado = 0 Then
        UserList(tIndex).flags.Silenciado = 1
        Call SendData(SendTarget.toindex, userindex, 0, "||737")
        Call SendData(SendTarget.toindex, tIndex, 0, "!!ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en mas. utilice /GM para contactar un administrador.")
        Call LogGM(UserList(userindex).Name, "/silenciar " & UserList(tIndex).Name, False)
    Else
        UserList(tIndex).flags.Silenciado = 0
        Call SendData(SendTarget.toindex, userindex, 0, "||738")
        Call LogGM(UserList(userindex).Name, "/DESsilenciar " & UserList(tIndex).Name, False)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "SOSDONE" Then
    rData = Right$(rData, Len(rData) - 7)
    Dim NombreSOS As String
    Dim NumeroSOS As String
    NombreSOS = ReadField(1, rData, Asc(","))
    NumeroSOS = ReadField(2, rData, Asc(","))
    
    For n = NumeroSOS To MensajesNumber
        If MensajesNumber >= NumeroSOS Then
            'If n + 1 < MensajesNumber Then
                MensajesSOS(n).Autor = MensajesSOS(n + 1).Autor
                MensajesSOS(n).Tipo = MensajesSOS(n + 1).Tipo
                MensajesSOS(n).Contenido = MensajesSOS(n + 1).Contenido
            'End If
            
            MensajesSOS(n + 1).Autor = ""
            MensajesSOS(n + 1).Tipo = ""
            MensajesSOS(n + 1).Contenido = ""
        End If
    Next n
    
        MensajesNumber = MensajesNumber - 1
        
        Dim dataSOS As String
        dataSOS = MensajesNumber & ","
        
        For loopC = 1 To MensajesNumber
            dataSOS = dataSOS & MensajesSOS(loopC).Tipo & "-" & MensajesSOS(loopC).Autor & "-" & MensajesSOS(loopC).Contenido & ","
        Next loopC
        
        Call SendData(SendTarget.ToAdmins, 0, 0, "ZSOS" & dataSOS)
    
    
    If NameIndex(NombreSOS) = 0 Then Exit Sub
    UserList(NameIndex(NombreSOS)).flags.ConsultaEnviada = False
    UserList(NameIndex(NombreSOS)).flags.NumeroConsulta = 0
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/CONT " Then
    rData = val(Right$(rData, Len(rData) - 6))
    If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
    If rData < 0 Or rData >= 61 Then Exit Sub
    MapaCont = UserList(userindex).Pos.Map
    If rData = 0 Or rData = "" Or rData = " " Then
        Call SendData(SendTarget.ToMap, 0, MapaCont, "||739")
        CuentaRegresiva = 0
        Exit Sub
    Else
    
    Call SendData(SendTarget.ToMap, 0, MapaCont, "||740@" & rData)
    SendData SendTarget.ToMap, 0, MapaCont, "CU" & rData
    CuentaRegresiva = rData
    End If
    Exit Sub
End If

'IR A
If UCase$(Left$(rData, 5)) = "/IRA " Then
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    'Si es dios o Admins no podemos salvo que nosotros también lo seamos
    If UserList(userindex).flags.Privilegios = PlayerType.User Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    

    Call WarpUserChar(userindex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y + 1, True)
    
    If UserList(userindex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.toindex, tIndex, 0, "||877@" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y, False)
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rData) = "/INVISIBLE" Then
    Call DoAdminInvisible(userindex)
    Call LogGM(UserList(userindex).Name, "/INVISIBLE", False)
    Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/ESPIAR " Then
    rData = Right$(rData, Len(rData) - 8)
    If UserList(userindex).flags.Privilegios < PlayerType.UserSupport Then Exit Sub
   
    tIndex = NameIndex(rData)
   
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
   
    If UserList(userindex).flags.AdminInvisible = 0 Then Call DoAdminInvisible(userindex)
    Call WarpUserChar(userindex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X - 2, UserList(tIndex).Pos.Y - 2, True)
   
    'If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.toindex, tIndex, 0, "||" & UserList(UserIndex).name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "/ESPIAR " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y, False)
    Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/ADVERTIR " Then
            If UserList(userindex).flags.Privilegios = User Then Exit Sub
       
            Dim TotalAdvert As Integer
            rData = UCase$(Right$(rData, Len(rData) - 10))
            Name = ReadField(1, rData, Asc("@"))
            tStr = ReadField(2, rData, Asc("@"))
                tIndex = NameIndex(Name)
 
    If Name = "" Or tStr = "" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||741")
        Exit Sub
    End If
 
    If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
        SendData SendTarget.toindex, userindex, 0, "||440"
      Exit Sub
    End If
               
        If UserList(userindex).flags.Privilegios < PlayerType.UserSupport Then Exit Sub
               
            TotalAdvert = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
       
            TotalAdvert = TotalAdvert + 1
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "CANT", TotalAdvert)
       
       
            If UserList(userindex).flags.EnTorneo = 1 Then
                UserList(userindex).flags.EnTorneo = 0
                
                For i = 1 To LastUser
                    If UserList(i).flags.NumTorneo > UserList(userindex).flags.NumTorneo Then
                        UserList(i).flags.NumTorneo = UserList(i).flags.NumTorneo - 1
                    End If
                Next i
                
                UserList(userindex).flags.NumTorneo = 0
                UsuariosEnTorneo = UsuariosEnTorneo - 1
                
                SendData SendTarget.ToAll, userindex, 0, "||724@" & UserList(userindex).Name
            End If
            
            
                Call SendData(SendTarget.ToAll, 0, 0, "||742@" & UserList(userindex).Name & "@" & Name)
                Call SendData(SendTarget.toindex, tIndex, 0, "||743@" & UserList(userindex).Name & "@" & tStr & "@" & TotalAdvert)
       
            If FileExist(CharPath & Name & ".chr", vbNormal) Then
                tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt)
                Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt, "Advertido por: " & UserList(userindex).Name & " | Motivo: " & LCase$(tStr) & ". | Fecha: " & Date & " " & Time & ".")
            End If
 
            'Encarcelamos el total de advertencias x 5 Ejemplo 2 Advertencias, lo encarcela por 10 Minutos.
            If Not val(TotalAdvert) >= 5 Then
                If tIndex >= 1 Then
                            Call Encarcelar(tIndex, TotalAdvert * 5, UserList(userindex).Name)
                Else
                    Call WriteVar(CharPath & Name & ".chr", "INIT", "POSITION", Prision.Map & "-" & Prision.X & "-" & Prision.Y)
                End If
            End If
 
            'Si llego al Maximo de Advertencias?
            If val(TotalAdvert) >= 5 Then
                Call SendData(SendTarget.ToAdmins, 0, 0, "||744@" & Name)
                    tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
                    Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, "El Servidor te ha Baneado Automaticamente. El Motivo es: Acumulacion de Advertencias. " & Date & " " & Time)
               
                'Desconectamos al usuario
                If Not tIndex <= 0 Then Call CloseSocket(tIndex)
           
                'Baneamos ^^
                Call WriteVar(CharPath & Name & ".chr", "FLAGS", "Ban", "1")
            End If
        Exit Sub
    End If

If UCase$(Left$(rData, 8)) = "/CARCEL " Then
    '/carcel nick@motivo@<tiempo>
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right$(rData, Len(rData) - 8)
    
    Name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or Name = "" Or tStr = "" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||745")
        Exit Sub
    End If
    i = val(ReadField(3, rData, Asc("@")))
    
    tIndex = NameIndex(Name)
    
    'If UCase$(Name) = "REEVES" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||442")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then Exit Sub
    
    If i > 60 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||746")
        Exit Sub
    End If
    
    Name = Replace(Name, "\", "")
    Name = Replace(Name, "/", "")
    
    If FileExist(CharPath & Name & ".chr", vbNormal) Then
        tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).Name) & ": CARCEL " & i & "m, MOTIVO: " & LCase$(tStr) & " " & Date & " " & Time)
    End If
    
    Call Encarcelar(tIndex, i, UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, " encarcelo a " & Name, False)
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/RMATA" Then

    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = UserList(userindex).flags.TargetNPC
    If tIndex > 0 Then
        Dim MiNPC As npc
        MiNPC = Npclist(tIndex)
        Call QuitarNPC(tIndex)
        Call ReSpawnNpc(MiNPC)
        
    'SERES
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||9")
    End If
    
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/SOBJ " Then

    rData = Right$(rData, Len(rData) - 6)
    For i = 1 To UBound(ObjData)
        If InStr(1, Tilde(ObjData(i).Name), Tilde(rData)) Then
            Call SendData(toindex, userindex, 0, "||748@" & ObjData(i).Name & "@" & i)
            n = n + 1
        End If
    Next
    If n = 0 Then
        Call SendData(toindex, userindex, 0, "||747")
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/EDIT " Then
    Call LogGM(UserList(userindex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub

Dim tesuboelnivel As Byte
For tesuboelnivel = 1 To 60
        UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.ELU
        Call CheckUserLevel(tIndex)
Next tesuboelnivel

    Exit Sub
End If


If UCase$(Left$(rData, 6)) = "/INFO " Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        SendUserStatsTxtOFF userindex, rData
    Else
        SendUserStatsTxt userindex, tIndex
    End If

    Exit Sub
End If
    
'INV DEL USER
If UCase$(Left$(rData, 5)) = "/INV " Then
    Call LogGM(UserList(userindex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        SendUserInvTxtFromChar userindex, rData
    Else
        SendUserInvTxt userindex, tIndex
    End If

    Exit Sub
End If

'INV DEL USER
If UCase$(Left$(rData, 5)) = "/BOV " Then
    Call LogGM(UserList(userindex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        SendUserBovedaTxtFromChar userindex, rData
    Else
        SendUserBovedaTxt userindex, tIndex
    End If

    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
    rData = Right$(rData, Len(rData) - 9)
    Name = rData
    If UCase$(Name) <> "YO" Then
        tIndex = NameIndex(Name)
    Else
        tIndex = userindex
    End If
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UserList(userindex).flags.Privilegios < PlayerType.EventManager Then Exit Sub
    
    UserList(tIndex).flags.Muerto = 0
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(tIndex)
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).Pos.Map, val(tIndex), UserList(tIndex).Char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(tIndex).Char.CascoAnim)
    Call SendUserHP(val(tIndex))
    Call SendData(SendTarget.toindex, tIndex, 0, "||749@" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Resucito a " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UCase$(rData) = "/ONLINEGM" Then
        For loopC = 1 To LastUser
            'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
            If (UserList(loopC).Name <> "") And UserList(loopC).flags.Privilegios > PlayerType.User And (UserList(loopC).flags.Privilegios < PlayerType.EventManager Or UserList(userindex).flags.Privilegios >= PlayerType.EventManager) Then
                tStr = tStr & UserList(loopC).Name & ", "
            End If
        Next loopC
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toindex, userindex, 0, "N|" & tStr & "~69~190~156")
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "N|No hay gms online~69~190~156")
        End If
        Exit Sub
End If

'Barrin 30/9/03
If UCase$(rData) = "/ONLINEMAP" Then
    For loopC = 1 To LastUser
        If UserList(loopC).Name <> "" And UserList(loopC).Pos.Map = UserList(userindex).Pos.Map And (UserList(loopC).flags.Privilegios < PlayerType.EventManager Or UserList(userindex).flags.Privilegios >= PlayerType.EventManager) Then
            tStr = tStr & UserList(loopC).Name & ", "
        End If
    Next loopC
    If Len(tStr) > 2 Then _
        tStr = Left$(tStr, Len(tStr) - 2)
    Call SendData(SendTarget.toindex, userindex, 0, "||750@" & tStr)
    Exit Sub
End If

'Echar usuario
If UCase$(Left$(rData, 7)) = "/ECHAR " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||442")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(SendTarget.toindex, userindex, 0, "||751")
        Exit Sub
    End If
        
    Call SendData(SendTarget.ToAll, 0, 0, "||752@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
    Call CloseSocket(tIndex)
    Call LogGM(UserList(userindex).Name, "Echo a " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/KILL " Then
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    rData = Right$(rData, Len(rData) - 6)
    tIndex = NameIndex(rData)
    
    If UCase$(UserList(userindex).Name) <> "KING" And UCase$(UserList(tIndex).Name) = "KING" Then Exit Sub
    
    If tIndex > 0 Then
        Call UserDie(tIndex)
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "||753@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        Call LogGM(UserList(userindex).Name, " ejecuto a " & UserList(tIndex).Name, False)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
    End If
Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/DAMETODO " Then
    rData = Right$(rData, Len(rData) - 10)
    tIndex = NameIndex(rData)
    
    If Not UserList(userindex).flags.Privilegios >= PlayerType.Coordination Then Exit Sub
    
    If UCase$(UserList(userindex).Name) <> "KING" And UCase$(UserList(tIndex).Name) = "KING" Then Exit Sub
    
    If tIndex > 0 Then
        Call DameTodo(tIndex)
        Call SendData(SendTarget.toindex, tIndex, 0, "||754@" & UserList(userindex).Name)
        Call SendData(SendTarget.ToAdmins, 0, 0, "||755@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        Call LogGM(UserList(userindex).Name, " le tiro todos los items a " & UserList(tIndex).Name, False)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
    End If
Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/DAMEBANCO " Then
    rData = Right$(rData, Len(rData) - 11)
    tIndex = NameIndex(rData)
    
    If Not UserList(userindex).flags.Privilegios >= PlayerType.Coordination Then Exit Sub
    
    If UCase$(UserList(userindex).Name) <> "KING" And UCase$(UserList(tIndex).Name) = "KING" Then Exit Sub
    
    If tIndex > 0 Then
        Call DameBanco(tIndex)
        Call SendData(SendTarget.toindex, tIndex, 0, "||756@" & UserList(userindex).Name)
        Call SendData(SendTarget.ToAdmins, 0, 0, "||757@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        Call LogGM(UserList(userindex).Name, " le tiro todos los items de la boveda a " & UserList(tIndex).Name, False)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
    End If
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/BAN " Then
 If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    rData = Right$(rData, Len(rData) - 5)
    tStr = ReadField(1, rData, Asc("@")) ' NICK
    tIndex = NameIndex(tStr)
    Name = ReadField(2, rData, Asc("@")) ' MOTIVO
    
    If UCase$(tStr) = "KING" Then Exit Sub

    If Name = "" Or tStr = "" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||758")
    Exit Sub
    End If
    
    If tIndex <= 0 Then
        If FileExist(CharPath & tStr & ".chr", vbNormal) Then
            
            If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
                Call SendData(SendTarget.toindex, userindex, 0, "||759")
                Exit Sub
            End If
            
            Call LogBanFromName(tStr, userindex, Name)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).Name) & ": BAN POR " & LCase$(Name) & " " & Date & " " & Time)

            Call LogGM(UserList(userindex).Name, "BAN a " & tStr, False)
            
            Call SendData(SendTarget.ToAll, 0, 0, "||760@" & UserList(userindex).Name & "@" & tStr)
            
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||189@" & tStr)
        End If
    Else
    
        Call LogBan(tIndex, userindex, Name)
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(userindex).Name, "BAN a " & UserList(tIndex).Name, False)
        

        Call SendData(SendTarget.ToAll, 0, 0, "||752@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        Call SendData(SendTarget.ToAll, 0, 0, "||760@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        
        'ponemos el flag de ban a 1
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).Name) & ": BAN POR " & LCase$(Name) & " " & Date & " " & Time)
        
        Call CloseSocket(tIndex)
    End If

    Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/BANACC " Then
    If UserList(userindex).flags.Privilegios <= PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    Dim Motivox As String
    tStr = ReadField(1, rData, Asc("@")) ' NICK
    Motivox = "" & ReadField(2, rData, Asc("@")) & "," & UserList(userindex).Name
    tIndex = NameIndex(tStr)
    
    If UCase$(tStr) = "KING" Then Exit Sub

   Dim NombreCuent As String
   
   If tIndex <> 0 Then
    NombreCuent = UserList(tIndex).Accounted
    Call SendData(SendTarget.ToAll, 0, 0, "||752@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
    Call CloseSocket(tIndex)
  Else
    NombreCuent = GetVar(CharPath & rData & ".chr", "CHAR", "Cuenta")
  End If
  
   Call WriteVar(App.Path & "\Accounts\" & NombreCuent & ".act", NombreCuent, "ban", "1")
   Call WriteVar(App.Path & "\Accounts\" & NombreCuent & ".act", NombreCuent, "Motivo", Motivox)
    
    'ponemos el flag de ban a 1
    Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            
    'ponemos la pena
    tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).Name) & ": BAN POR " & LCase$(Name) & " " & Date & " " & Time)

    Call LogGM(UserList(userindex).Name, "BAN a " & tStr, False)
   
   Call SendData(SendTarget.ToAll, 0, 0, "||760@" & UserList(userindex).Name & "@" & tStr)
   Call SendData(SendTarget.ToAll, 0, 0, "||761@" & UserList(userindex).Name & "@" & NombreCuent)
   Call LogGM(UserList(userindex).Name, "Baneo la cuenta y el personaje: " & NombreCuent & " - " & tStr & "", False)
   Exit Sub
End If


If UCase$(Left$(rData, 7)) = "/UNBAN " Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 7)
    
    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")
    
    If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||189@" & rData)
        Exit Sub
    End If
    
    Call UnBan(rData)
    
    'penas
    i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
    Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
    Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(userindex).Name) & ": UNBAN. " & Date & " " & Time)
    
    Call LogGM(UserList(userindex).Name, "/UNBAN a " & rData, False)
    Call SendData(SendTarget.ToAll, 0, 0, "||762@" & UserList(userindex).Name & "@" & rData)

    Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/UNBANACC " Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 10)
    
    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")
    
    If Not FileExist(App.Path & "\Charfile\" & rData & ".chr", vbNormal) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||189@" & rData)
        Exit Sub
    End If
    
    Dim Cuent As String
    Cuent = GetVar(CharPath & rData & ".chr", "CHAR", "Cuenta")
    
    Call WriteVar(App.Path & "\Accounts\" & Cuent & ".act", Cuent, "ban", "0")
    
    Call LogGM(UserList(userindex).Name, "/UNBANACC a " & rData, False)
    
    Call SendData(SendTarget.ToAll, 0, 0, "||763@" & UserList(userindex).Name & "@" & Cuent)
    

    Exit Sub
End If


If UCase$(Left$(rData, 10)) = "/DOTORNEO " Then
    rData = Right$(rData, Len(rData) - 10)
    TModalidad = ReadField(1, rData, Asc("@"))
    
    If ReadField(1, rData, Asc("@")) = vbNullString Then
       Call SendData(SendTarget.toindex, userindex, 0, "||764")
    Exit Sub
    End If
    
If Hay_Torneo = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||765")
 Exit Sub
End If
    
If UCase$(TModalidad) = "DM" Then
     CParticipantes = ReadField(2, rData, Asc("@"))
     TNivelMinimo = ReadField(3, rData, Asc("@"))
     
    If TNivelMinimo > 60 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||766")
      Exit Sub
    End If
   
    If TNivelMinimo < 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||766")
      Exit Sub
    End If
   
    If CParticipantes < 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||524")
      Exit Sub
    End If
     
    Call SendData(SendTarget.ToAll, 0, 0, "||767@" & UserList(userindex).Name & "@TODOS CONTRA TODOS@" & CParticipantes & "@" & TNivelMinimo)
    CuentaTorneo = 10
    UsuariosEnTorneo = 0
    Hay_Torneo = True
    TiroCuentaDM = False

    Call LogTorneos("" & UserList(userindex).Name & " - Modalidad: " & UCase$(TModalidad) & " - Nivel Minimo: " & TNivelMinimo & " - Participantes: " & CParticipantes)
 Exit Sub
End If
   
    If TModalidad <> 5 Then
     CParticipantes = ReadField(2, rData, Asc("@"))
     TNivelMinimo = ReadField(3, rData, Asc("@"))
    Else
     CParticipantes = ReadField(2, rData, Asc("@"))
     TNivelMinimo = 1
    End If
   
    If TNivelMinimo > 70 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||766")
      Exit Sub
    End If
   
    If TNivelMinimo < 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||766")
      Exit Sub
    End If
   
    If CParticipantes < 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||524")
      Exit Sub
    End If
   
If Hay_Torneo = False Then
If TModalidad = "1" Or UCase$(TModalidad) = "1VS1" Then
Call SendData(SendTarget.ToAll, 0, 0, "||767@" & UserList(userindex).Name & "@1 VS 1@" & CParticipantes & "@" & TNivelMinimo)
CuentaTorneo = 10
UsuariosEnTorneo = 0
Hay_Torneo = True
ElseIf TModalidad = "2" Or UCase$(TModalidad) = "2VS2" Then
Call SendData(SendTarget.ToAll, 0, 0, "||767@" & UserList(userindex).Name & "@2 VS 2@" & CParticipantes & "@" & TNivelMinimo)
CuentaTorneo = 10
UsuariosEnTorneo = 0
Hay_Torneo = True
ElseIf TModalidad = "3" Or UCase$(TModalidad) = "3VS3" Then
Call SendData(SendTarget.ToAll, 0, 0, "||767@" & UserList(userindex).Name & "@3 VS 3@" & CParticipantes & "@" & TNivelMinimo)
CuentaTorneo = 10
UsuariosEnTorneo = 0
Hay_Torneo = True
ElseIf TModalidad = "4" Or UCase$(TModalidad) = "4VS4" Then
Call SendData(SendTarget.ToAll, 0, 0, "||767@" & UserList(userindex).Name & "@4 VS 4@" & CParticipantes & "@" & TNivelMinimo)
CuentaTorneo = 10
UsuariosEnTorneo = 0
Hay_Torneo = True
ElseIf TModalidad = "5" Then
Call SendData(SendTarget.ToAll, 0, 0, "||768@" & UserList(userindex).Name & "@" & CParticipantes)
PuntosPremios = val(CParticipantes)
UsuariosEnTorneo = 0
Hay_Torneo = True
End If
Else
Call SendData(SendTarget.toindex, userindex, 0, "||765")
Exit Sub
End If
 
For tornein = 1 To LastUser
    If UserList(tornein).flags.EnTorneo = 1 Then
        UserList(tornein).flags.EnTorneo = 0
    End If
     
    If UserList(tornein).flags.NumTorneo > 0 Then
        UserList(tornein).flags.NumTorneo = 0
    End If
Next tornein

Call LogTorneos("" & UserList(userindex).Name & " - Modalidad: " & UCase$(TModalidad) & " - Nivel Minimo: " & TNivelMinimo & " - Participantes: " & CParticipantes)
 
Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/DV " Then
    rData = Right$(rData, Len(rData) - 4)
    tIndex = NameIndex(rData)
   
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then Exit Sub
 
    If (UserList(tIndex).Pos.Map = 100 Or UserList(tIndex).Pos.Map = 99) And UserList(tIndex).flags.MapaAnterior_dos <> 0 Then
        Call WarpUserChar(tIndex, UserList(tIndex).flags.MapaAnterior_dos, UserList(tIndex).flags.XAnterior_dos, UserList(tIndex).flags.YAnterior_dos, True)
    Exit Sub
    End If
 
    If UserList(tIndex).flags.MapaAnterior <> 0 Then
        Call WarpUserChar(tIndex, UserList(tIndex).flags.MapaAnterior, UserList(tIndex).flags.XAnterior, UserList(tIndex).flags.YAnterior, True)
    Else
        Call WarpUserChar(tIndex, 1, RandomNumber(52, 56), RandomNumber(36, 38), True)
    End If
 
Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/TSUM " Then
    rData = Right$(rData, Len(rData) - 6)
    Dim tsum1 As Byte
    Dim tsum2 As Byte
    tsum1 = ReadField(1, rData, Asc("@"))
    tsum2 = ReadField(2, rData, Asc("@"))
   
Dim usuariosvv As Integer
For usuariosvv = 1 To LastUser
    If UserList(usuariosvv).flags.NumTorneo >= val(tsum1) And UserList(usuariosvv).flags.NumTorneo <= val(tsum2) Then
            
        If UserList(usuariosvv).Pos.Map = 54 Or UserList(usuariosvv).Pos.Map = 8 Or UserList(usuariosvv).Pos.Map = 72 Or UserList(usuariosvv).Pos.Map = 110 Or UserList(usuariosvv).Pos.Map = 111 Or UserList(usuariosvv).Pos.Map = 78 Or UserList(usuariosvv).Pos.Map = 101 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||769")
        Else
            Call WarpUserChar(usuariosvv, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y + 1, False)
        End If
    End If
Next usuariosvv
 
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/PREMIAR " Then
rData = Right$(rData, Len(rData) - 9)
tIndex = NameIndex(ReadField(1, rData, 32))
Arg1 = ReadField(2, rData, 32)

If tIndex <= 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||196")
Exit Sub
End If

If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
 
    If val(Arg1) < 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||524")
    Else
        Call AgregarPuntos(tIndex, val(Arg1))
        Call SendData(SendTarget.toindex, userindex, 0, "||770@" & val(Arg1) & "@" & UserList(tIndex).Name)
        Call SendData(SendTarget.toindex, tIndex, 0, "||771@" & UserList(userindex).Name & "@" & val(Arg1))
        Call WriteVar(CharPath & UserList(tIndex).Name & ".chr", "STATS", "PuntosTorneo", UserList(tIndex).Stats.PuntosTorneo)
        Call LogGM(UserList(userindex).Name, "/Darpun " & val(Arg1) & " " & UserList(tIndex).Name, False)
    End If
    
Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/HOME " Then
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" And UCase$(UserList(userindex).Name) <> "KING" Then Exit Sub
    
                    If tIndex = GranPoder Then
                        GranPoder = 0
                        UserList(tIndex).flags.GranPoder = 0
                        SendUserVariant (tIndex)
                        Call OtorgarGranPoder(0)
                    End If
                   
                       Call WarpUserChar(tIndex, 1, 54, 36, True)
                       Call SendData(toindex, userindex, 0, "||772")
                       Call SendData(toindex, tIndex, 0, "||773")
    
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/EXP " Then
    rData = Right$(rData, Len(rData) - 5)
    
If Not IsNumeric(rData) Then Exit Sub
If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    
    MultiplicadorExp = rData
    Call WriteVar(IniPath & "Server.ini", "INIT", "MultiplicadordeExp", rData)
    Call SendData(SendTarget.ToAll, 0, 0, "||774@" & rData)
    
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/GLD " Then
rData = Right$(rData, Len(rData) - 5)
    
If Not IsNumeric(rData) Then Exit Sub
If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    
    MultiplicadorOro = rData
    Call WriteVar(IniPath & "Server.ini", "INIT", "MultiplicadordeOro", rData)
    Call SendData(SendTarget.ToAll, 0, 0, "||775@" & rData)
    
Exit Sub
End If


If UCase$(Left$(rData, 6)) = "/DROP " Then
rData = Right$(rData, Len(rData) - 6)
    
If Not IsNumeric(rData) Then Exit Sub
If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    
    MultiplicadorDrop = rData
    Call WriteVar(IniPath & "Server.ini", "INIT", "MultiplicadordeDrop", rData)
    Call SendData(SendTarget.ToAll, 0, 0, "||776@" & rData)
    
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/ATORNEO " Then
rData = Right$(rData, Len(rData) - 9)
Dim Torneos As Integer
Torneos = CInt(rData)

If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub

If (Torneos > 0 And Torneos < 6) Then
    Call Torneos_Inicia(userindex, Torneos)
End If

Call LogTorneos("" & UserList(userindex).Name & " - /ATORNEO " & Torneos)

Exit Sub
End If

If UCase(rData) = "/CANCELARTORNEO" Then
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    Call Rondas_Cancela
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/CHORI " Then
    rData = Right$(rData, Len(rData) - 7)
    Name = ReadField(1, rData, 32)
   
    If NameIndex(Name) <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
    Exit Sub
    End If
   
    If RevisandoUsuario = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||777")
    Exit Sub
    End If
   
    RevisandoUsuario = True
    UsuarioRevisado = NameIndex(Name)
    Call SendData(SendTarget.toindex, userindex, 0, "CHX" & UserList(UsuarioRevisado).Stats.MaxHP & "," & UserList(UsuarioRevisado).Stats.MinHP & "," & UserList(UsuarioRevisado).Stats.MaxMAN & "," & UserList(UsuarioRevisado).Stats.MinMAN & "," & UserList(UsuarioRevisado).Name)
   
    Call LogGM(UserList(userindex).Name, "/CHORI " & Name, False)
Exit Sub
End If

'Summon
If UCase$(Left$(rData, 5)) = "/SUM " Then
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
    If UCase$(rData) = "KING" Then Exit Sub
    
    If UserList(tIndex).Pos.Map = 101 Or UserList(tIndex).Pos.Map = 72 Or UserList(tIndex).Pos.Map = 54 Or UserList(tIndex).Pos.Map = 78 Or UserList(tIndex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||769")
        Exit Sub
    End If
    
    UserList(tIndex).flags.MapaAnterior = UserList(tIndex).Pos.Map
    UserList(tIndex).flags.XAnterior = UserList(tIndex).Pos.X
    UserList(tIndex).flags.YAnterior = UserList(tIndex).Pos.Y
    
    Call SendData(SendTarget.toindex, tIndex, 0, "||778@" & UserList(userindex).Name)
    Call WarpUserChar(tIndex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y + 1, True)
    
    Call LogGM(UserList(userindex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(userindex).Pos.Map & " X:" & UserList(userindex).Pos.X & " Y:" & UserList(userindex).Pos.Y, False)
    Exit Sub
End If

'Crear criatura
If UCase$(Left$(rData, 3)) = "/CC" Then
   Call EnviarSpawnList(userindex)
   Exit Sub
End If

If UCase$(rData) = "/VERPRIVADOS" Then

If UserList(userindex).flags.Privilegios >= PlayerType.Coordination Then
    If VerPrivados = False Then
        VerPrivados = True
        Call SendData(SendTarget.ToAdmins, 0, 0, "||779")
    Else
        VerPrivados = False
        Call SendData(SendTarget.ToAdmins, 0, 0, "||780")
    End If
End If

Exit Sub
End If

If UCase$(rData) = "/VERCLANES" Then

If UserList(userindex).flags.Privilegios >= PlayerType.Coordination Then
    If VerClanes = False Then
        VerClanes = True
        Call SendData(SendTarget.ToAdmins, 0, 0, "||781")
    Else
        VerClanes = False
        Call SendData(SendTarget.ToAdmins, 0, 0, "||782")
    End If
End If

Exit Sub
End If

If UCase$(rData) = "/FINCHORI" Then
    RevisandoUsuario = False
    Call SendData(SendTarget.toindex, userindex, 0, "||783")
Exit Sub
End If

If UCase$(rData) = "/RESETBONIS" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 11)
    If UserList(userindex).flags.TargetUser = 0 Then Exit Sub
    
    UserList(userindex).Bon1 = "Ninguno/No elegido."
    UserList(userindex).Bon2 = "Ninguno/No elegido."
    UserList(userindex).Bon3 = "Ninguno/No elegido."
    
    Call SendData(SendTarget.toindex, userindex, 0, "N|Se han reseteado los bonificadores del usuario seleccionado.~255~255~255~1")
    
    Exit Sub
End If

'Resetea el inventario
If UCase$(rData) = "/RESETINV" Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
    Call ResetNpcInv(UserList(userindex).flags.TargetNPC)
    Call LogGM(UserList(userindex).Name, "/RESETINV " & Npclist(UserList(userindex).flags.TargetNPC).Name, False)
    Exit Sub
End If

'/Clean
If UCase$(rData) = "/LIMPIAR" Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call LimpiarMundoEntero
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/RMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rData, False)
    If rData <> "" Then
        Call SendData(SendTarget.ToAll, 0, 0, "N|" & UserList(userindex).Name & "> " & rData & FONTTYPE_CELESTEN & ENDC)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/TFER " Then
    rData = Right$(rData, Len(rData) - 6)
    
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    If Not IsNumeric(rData) Then Exit Sub
        
        BOnlines = val(rData)
        Call MostrarNumUsers
        
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/LMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    tStr = ReadField(1, rData, Asc("@"))
    Name = ReadField(2, rData, Asc("@"))
    
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    If tStr = "" Then MensajeAutomatico = False
    If Not IsNumeric(Name) Then Exit Sub
    
    MensajeAutomatico = True
    TextoMensajeAutomatico = tStr
    TiempoMensajeAutomatico = Name
    MinutitosMensaje = 0
        
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/FPS " Then
    rData = Right$(rData, Len(rData) - 5)
    tIndex = NameIndex(rData)
    Call LogGM(UserList(userindex).Name, "/FPS A:" & UserList(tIndex).Name, False)
    
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||196")
        Exit Sub
    End If
    
   Call SendData(SendTarget.toindex, tIndex, 0, "ENVFPS")
    
Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/GO " Then
    rData = Right$(rData, Len(rData) - 4)
    mapa = val(ReadField(1, rData, 32))
    
    If Not FileExist(App.Path & "\Maps\Mapa" & mapa & ".map") Then Exit Sub
    
    Call LogGM(UserList(userindex).Name, "/GO " & mapa, False)
    
    Call WarpUserChar(userindex, mapa, 50, 50, True)
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/LMAP " Then
    rData = Right$(rData, Len(rData) - 6)
    mapa = val(ReadField(1, rData, 32))
    
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    If Not FileExist(App.Path & "\Maps\Mapa" & mapa & ".map") Then Exit Sub
    
    Call LogGM(UserList(userindex).Name, "/LMAP " & mapa, False)
    
        Call LimpiarMapa(mapa)
        Call SendData(SendTarget.toindex, userindex, 0, "||784@" & mapa)
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rData, 9)) = "/NICK2IP " Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    tIndex = NameIndex(UCase$(rData))
    Call LogGM(UserList(userindex).Name, "NICK2IP Solicito la IP de " & rData, False)
    If tIndex > 0 Then
        If (UserList(userindex).flags.Privilegios > PlayerType.User And UserList(tIndex).flags.Privilegios = PlayerType.User) Or (UserList(userindex).flags.Privilegios >= PlayerType.EventManager) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||786@" & rData & "@" & UserList(tIndex).ip)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||785")
        End If
    Else
       Call SendData(SendTarget.toindex, userindex, 0, "||196")
    End If
    Exit Sub
End If
 
'Ip del nick
If UCase$(Left$(rData, 9)) = "/IP2NICK " Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)

    If InStr(rData, ".") < 1 Then
        tInt = NameIndex(rData)
        If tInt < 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||196")
            Exit Sub
        End If
        rData = UserList(tInt).ip
    End If
    tStr = vbNullString
    Call LogGM(UserList(userindex).Name, "IP2NICK Solicito los Nicks de IP " & rData, False)
    For loopC = 1 To LastUser
        If UserList(loopC).ip = rData And UserList(loopC).Name <> "" And UserList(loopC).flags.UserLogged Then
            If (UserList(userindex).flags.Privilegios > PlayerType.User And UserList(loopC).flags.Privilegios = PlayerType.User) Or (UserList(userindex).flags.Privilegios >= PlayerType.EventManager) Then
                tStr = tStr & UserList(loopC).Name & ", "
            End If
        End If
    Next loopC
    
    Call SendData(SendTarget.toindex, userindex, 0, "||787@" & rData & "@" & tStr)
    Exit Sub
End If


'Crear Teleport
If UCase(Left(rData, 4)) = "/CT " Then
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    '/ct mapa_dest x_dest y_dest
    rData = Right(rData, Len(rData) - 4)
    Call LogGM(UserList(userindex).Name, "/CT: " & rData, False)
    mapa = ReadField(1, rData, 32)
    X = ReadField(2, rData, 32)
    Y = ReadField(3, rData, 32)
    
    If MapaValido(mapa) = False Or InMapBounds(mapa, X, Y) = False Then
        Exit Sub
    End If
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).userindex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).NpcIndex > 0 Then
        Exit Sub
    End If
    
    If MapData(mapa, X, Y).OBJInfo.ObjIndex > 0 Then
        Call SendData(SendTarget.toindex, userindex, mapa, "||788")
        Exit Sub
    End If
    
    Dim ET As Obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call MakeObj(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, ET, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1)
    
    MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).TileExit.Map = mapa
    MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).TileExit.X = X
    MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y - 1).TileExit.Y = Y
    Call SendData(SendTarget.toindex, userindex, 0, "||789@" & mapa & "@" & X & "@" & Y)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||790@" & UserList(userindex).Name & "@" & mapa & "@" & X & "@" & Y)
    
    Exit Sub
End If

'Destruir Teleport
'toma el ultimo click
If UCase(Left(rData, 3)) = "/DT" Then
    '/dt
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    Call LogGM(UserList(userindex).Name, "/DT", False)
    
    mapa = UserList(userindex).flags.TargetMap
    X = UserList(userindex).flags.TargetX
    Y = UserList(userindex).flags.TargetY
    
    If ObjData(MapData(mapa, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTeleport And _
        MapData(mapa, X, Y).TileExit.Map > 0 Then
        Call EraseObj(SendTarget.ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
        Call EraseObj(SendTarget.ToMap, 0, MapData(mapa, X, Y).TileExit.Map, 1, MapData(mapa, X, Y).TileExit.Map, MapData(mapa, X, Y).TileExit.X, MapData(mapa, X, Y).TileExit.Y)
        MapData(mapa, X, Y).TileExit.Map = 0
        MapData(mapa, X, Y).TileExit.X = 0
        MapData(mapa, X, Y).TileExit.Y = 0
    End If
    Call SendData(SendTarget.toindex, userindex, 0, "||791")
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SETDESC " Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    DummyInt = UserList(userindex).flags.TargetUser
    If DummyInt > 0 Then
        UserList(DummyInt).DescRM = rData
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||9")
    End If
    Exit Sub
    
End If

Select Case UCase$(Left$(rData, 8))
    Case "/TALKAS "
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
            'Asegurarse haya un NPC seleccionado
            If UserList(userindex).flags.TargetNPC > 0 Then
                tStr = Right$(rData, Len(rData) - 8)
                
                Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).Pos.Map, "N|" & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||9")
            End If
        Exit Sub
End Select


'[Barrin 30-11-03]
'Quita todos los objetos del area
If UCase$(rData) = "/MASSDEST" Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    For Y = UserList(userindex).Pos.Y - MinYBorder + 1 To UserList(userindex).Pos.Y + MinYBorder - 1
            For X = UserList(userindex).Pos.X - MinXBorder + 1 To UserList(userindex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex > 0 Then _
                    If ItemNoEsDeMapa(MapData(UserList(userindex).Pos.Map, X, Y).OBJInfo.ObjIndex) Then Call EraseObj(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, 10000, UserList(userindex).Pos.Map, X, Y)
            Next X
    Next Y
    Call LogGM(UserList(userindex).Name, "/MASSDEST", False)
    Exit Sub
End If
'[/Barrin 30-11-03]


If UCase$(Left$(rData, 12)) = "/ACEPTCONSE " Then
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
rData = Right$(rData, Len(rData) - 12)
tIndex = NameIndex(rData)
If tIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||196")
Else
Call SendData(SendTarget.ToAll, 0, 0, "||792@" & rData)
UserList(tIndex).ConsejoInfo.PertAlCons = 1
UserList(tIndex).ConsejoInfo.LiderConsejo = 1
UserList(tIndex).StatusMith.EsStatus = 5
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
End If
Exit Sub
End If

If UCase$(Left$(rData, 16)) = "/ACEPTCONSECAOS " Then
If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
rData = Right$(rData, Len(rData) - 16)
tIndex = NameIndex(rData)
If tIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||196")
Else
Call SendData(SendTarget.ToAll, 0, 0, "||793@" & rData)
UserList(tIndex).ConsejoInfo.PertAlConsCaos = 1
UserList(tIndex).ConsejoInfo.LiderConsejoCaos = 1
UserList(tIndex).StatusMith.EsStatus = 6
Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
End If
Exit Sub
End If

    If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
        rData = Right$(rData, Len(rData) - 11)
        tIndex = NameIndex(rData)
        
        If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
        
            If tIndex <= 0 Then
                If FileExist(CharPath & rData & ".chr") Then
                    Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||189@" & rData)
                End If
            Else
                If UserList(tIndex).ConsejoInfo.PertAlCons > 0 Then
                    UserList(tIndex).ConsejoInfo.PertAlCons = 0
                    UserList(tIndex).ConsejoInfo.LiderConsejo = 0
                    UserList(tIndex).StatusMith.EsStatus = 3
                    Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
                    Call SendData(SendTarget.ToAll, 0, 0, "||794@" & rData)
                End If
             
                If UserList(tIndex).ConsejoInfo.PertAlConsCaos > 0 Then
                    UserList(tIndex).StatusMith.EsStatus = 4
                    UserList(tIndex).ConsejoInfo.LiderConsejoCaos = 0
                    UserList(tIndex).ConsejoInfo.PertAlConsCaos = 0
                    Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
                    Call SendData(SendTarget.ToAll, 0, 0, "||795@" & rData)
                End If
            End If
        Exit Sub
    End If
    '[/yb]

If UCase$(Left$(rData, 8)) = "/TRIGGER" Then
    Call LogGM(UserList(userindex).Name, rData, False)
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    
    rData = Trim(Right(rData, Len(rData) - 8))
    mapa = UserList(userindex).Pos.Map
    X = UserList(userindex).Pos.X
    Y = UserList(userindex).Pos.Y
    If rData <> "" Then
        tInt = MapData(mapa, X, Y).trigger
        MapData(mapa, X, Y).trigger = val(rData)
    End If
    Call SendData(SendTarget.toindex, userindex, 0, "||796@" & MapData(mapa, X, Y).trigger & "@" & mapa & "@" & X & "@" & Y)
    Exit Sub
End If

'Ban x IP
If UCase(Left(rData, 7)) = "/BANIP " Then
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Dim BanIP As String, XNick As Boolean
    
    rData = Right$(rData, Len(rData) - 7)
    tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")
    'busca primero la ip del nick
    tIndex = NameIndex(tStr)
    
    If UCase$(tStr) = "KING" Then Exit Sub
    
    If tIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(userindex).Name, "/BanIP " & rData, False)
        BanIP = tStr
    Else
        XNick = True
        Call LogGM(UserList(userindex).Name, "/BanIP " & UserList(tIndex).Name & " - " & UserList(tIndex).ip, False)
        BanIP = UserList(tIndex).ip
    End If
    
    rData = Right$(rData, Len(rData) - Len(tStr))
    
    If BanIpBuscar(BanIP) > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||797")
        Exit Sub
    End If
    
    Call BanIpAgrega(BanIP)
    
    If XNick = True Then
        Call LogBan(tIndex, userindex, "Ban por IP desde Nick por " & rData)
        
        Call SendData(SendTarget.ToAdmins, 0, 0, "||752@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        Call SendData(SendTarget.ToAdmins, 0, 0, "||798@" & UserList(userindex).Name & "@" & UserList(tIndex).Name)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(userindex).Name, "Echo a " & UserList(tIndex).Name, False)
        Call LogGM(UserList(userindex).Name, "BAN a " & UserList(tIndex).Name, False)
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If

'Desbanea una IP
If UCase(Left(rData, 9)) = "/UNBANIP " Then
    
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    
    rData = Right(rData, Len(rData) - 9)
    Call LogGM(UserList(userindex).Name, "/UNBANIP " & rData, False)
    
    If BanIpQuita(rData) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||799@" & rData)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||800")
    End If
    
    Exit Sub
End If

'Crear Item
If UCase(Left(rData, 11)) = "/HACERITEM " Then
    ' [GS] Soporte de /haceritem ¬¬
    If UCase(Left(rData, 11)) = "/HACERITEM " Then
        rData = Right$(rData, Len(rData) - 11)
    Else
        rData = Right$(rData, Len(rData) - 4)
    End If
    
    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
   
    Call LogGM(UserList(userindex).Name, "/HACERITEM: " & rData & "@" & tInt, False)
    Call LogGMss(UserList(userindex).Name, "/HACERITEM: " & rData & "@" & tInt, False)
   
    ' [GS] CI y cantidades
    tInt = val(ReadField(2, rData, Asc("@")))
    If tInt < 1 Or tInt > 10000 Then tInt = 1
    rData = val(ReadField(1, rData, Asc("@")))
    ' [/GS]
   
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If (val(rData) < 1) Or (val(rData) > NumObjDatas) Then
        Exit Sub
    End If
   
    Dim Objeto As Obj
    Objeto.Amount = tInt ' [GS] Cantidad
    Objeto.ObjIndex = val(rData)
   
    Call MakeObj(ToMap, 0, UserList(userindex).Pos.Map, Objeto, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
   
    Call SendData(SendTarget.toindex, userindex, 0, "||801@" & Objeto.ObjIndex & "@" & ObjData(Objeto.ObjIndex).Name & "@" & Objeto.Amount)
    Call SendData(SendTarget.ToAdmins, userindex, 0, "||802@" & UserList(userindex).Name & "@" & Objeto.ObjIndex & "@" & ObjData(Objeto.ObjIndex).Name & "@" & Objeto.Amount)
    Exit Sub
End If

'Global:
If UCase$(Left$(rData, 9)) = "/NOGLOBAL" Then
rData = Right$(rData, Len(rData) - 9)

    If UserList(userindex).flags.Privilegios > PlayerType.User And UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub

    If ChatGlobal = False Then
    Call SendData(SendTarget.ToAll, 0, 0, "||803")
    ChatGlobal = True
    ElseIf ChatGlobal = True Then
    Call SendData(SendTarget.ToAll, 0, 0, "||804")
    ChatGlobal = False
    End If

Exit Sub
End If

'Destruir
If UCase$(Left$(rData, 5)) = "/DEST" Then
    Call LogGM(UserList(userindex).Name, "/DEST", False)
    Call LogGMss(UserList(userindex).Name, "/DEST", False)
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    rData = Right$(rData, Len(rData) - 5)
    Call EraseObj(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, 10000, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
    Exit Sub
End If

'Bloquear
If UCase$(Left$(rData, 5)) = "/BLOQ" Then
    Call LogGM(UserList(userindex).Name, "/BLOQ", False)
    Call LogGMss(UserList(userindex).Name, "/BLOQ", False)
    rData = Right$(rData, Len(rData) - 5)
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Blocked = 0 Then
        MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Blocked = 1
        Call Bloquear(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, 1)
    Else
        MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Blocked = 0
        Call Bloquear(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y, 0)
    End If
    Exit Sub
End If

'Quitar NPC
If UCase$(rData) = "/MATA" Then
    rData = Right$(rData, Len(rData) - 5)
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
    Call QuitarNPC(UserList(userindex).flags.TargetNPC)
    Call LogGM(UserList(userindex).Name, "/MATA " & Npclist(UserList(userindex).flags.TargetNPC).Name, False)
    Call LogGMss(UserList(userindex).Name, "/MATA " & Npclist(UserList(userindex).flags.TargetNPC).Name, False)
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rData) = "/MASSKILL" Then
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    For Y = UserList(userindex).Pos.Y - MinYBorder + 1 To UserList(userindex).Pos.Y + MinYBorder - 1
            For X = UserList(userindex).Pos.X - MinXBorder + 1 To UserList(userindex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex)
            Next X
    Next Y
    Call LogGM(UserList(userindex).Name, "/MASSKILL", False)
    Exit Sub
End If

If UCase$(rData) = "/SALEREY" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    If ReyON = 0 Then
    
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
       
        Call SendData(ToAll, 0, 0, "||805")
        Call SpawnNpc(Guardia, Guardia1, True, False)
        Call SpawnNpc(Guardia, Guardia2, True, False)
        Call SpawnNpc(Guardia, Guardia3, True, False)
        Call SpawnNpc(Guardia, Guardia4, True, False)
        IndexReyAncalagon = SpawnNpc(Rey, PosicionR, True, False)
        Npclist(IndexReyAncalagon).Char.AuraA = 3
        Call MakeNPCChar(SendTarget.ToMap, 0, 0, IndexReyAncalagon, Npclist(IndexReyAncalagon).Pos.Map, Npclist(IndexReyAncalagon).Pos.X, Npclist(IndexReyAncalagon).Pos.Y)
        MinutosRey = 0
        GuardiasRey = 0
        ReyON = 1
    End If
    Exit Sub
End If

If UCase$(rData) = "/INIARAM" Then
    Call Aram_Inscripciones
    Exit Sub
End If

If UCase$(rData) = "/CANCELARAM" Then
    Call Aram_Cancelar
    Exit Sub
End If

'Apagamos
If UCase$(rData) = "/OFF" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Administrator Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(userindex).Name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
End If

If UCase$(Left$(rData, 12)) = "/GUARDARMAPA" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    Call GrabarMapa(UserList(userindex).Pos.Map, App.Path & "\Maps\Mapa" & UserList(userindex).Pos.Map)
    Exit Sub
End If

If UCase$(Left$(rData, 12)) = "/MODMAPINFO " Then
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    rData = Right(rData, Len(rData) - 12)
    Select Case UCase(ReadField(1, rData, 32))
    Case "PK"
        tStr = ReadField(2, rData, 32)
        If tStr <> "" Then
            MapInfo(UserList(userindex).Pos.Map).Pk = IIf(tStr = "0", True, False)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).Pos.Map & ".dat", "Mapa" & UserList(userindex).Pos.Map, "Pk", tStr)
        End If
    End Select
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/GRABAR" Then
    If UserList(userindex).flags.Privilegios < PlayerType.TournamentManager Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    Call GuardarUsuarios
    Exit Sub
End If

If UCase$(rData) = "/ECHARTODOSPJS" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    Call EcharPjsNoPrivilegiados
    Exit Sub
End If

If UCase$(rData) = "/RELOADSINI" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call LogGM(UserList(userindex).Name, rData, False)
    Call LoadSini
    Exit Sub
End If

If UCase$(rData) = "/LOADPREMIOS" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call SendData(SendTarget.ToAll, userindex, 0, "||806@PREMIOS")
    Call CargarPremiosList
    Call CargarDonaciones
    Call SendData(SendTarget.ToAll, userindex, 0, "||807@PREMIOS")
    Exit Sub
End If

If UCase$(rData) = "/LOADOBJ" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call SendData(SendTarget.ToAll, userindex, 0, "||806@OBJETOS")
    Call LoadOBJData
    Call CargarCofresRandom
    Call SendData(SendTarget.ToAll, userindex, 0, "||807@OBJETOS")
    Exit Sub
End If

If UCase$(rData) = "/LOADBALANCE" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call SendData(SendTarget.ToAll, userindex, 0, "||806@BALANCE")
    Call LoadBalance
    Call CargarIntervalos
    Call SendData(SendTarget.ToAll, userindex, 0, "||807@BALANCE")
    Exit Sub
End If
 
If UCase$(rData) = "/LOADHECHIZOS" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call SendData(SendTarget.ToAll, userindex, 0, "||806@HECHIZOS")
    Call CargarHechizos
    Call SendData(SendTarget.ToAll, userindex, 0, "||807@HECHIZOS")
    Exit Sub
End If
 
If UCase$(rData) = "/LOADNPCS" Then
    If UserList(userindex).flags.Privilegios < PlayerType.Coordination Then Exit Sub
    Call SendData(SendTarget.ToAll, userindex, 0, "||806@NPCS")
    Call DescargaNpcsDat
    Call CargaNpcsDat
    Call SendData(SendTarget.ToAll, userindex, 0, "||807@NPCS")
    Exit Sub
End If

Call SendData(SendTarget.toindex, userindex, 0, "||714")

Exit Sub

ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(userindex).Name & "UI:" & userindex & " N: " & Err.Number & " D: " & Err.Description)
 'Resume
 'Call CloseSocket(UserIndex)
 'Call Cerrar_Usuario(UserIndex)
 
 

End Sub

Sub ReloadSokcet()
On Error GoTo Errhandler
#If UsarQueSocket = 1 Then
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
Errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal userindex As Integer)

Call SendData(SendTarget.toindex, userindex, 0, "NOC" & IIf(DeNoche And (MapInfo(UserList(userindex).Pos.Map).Zona = Campo Or MapInfo(UserList(userindex).Pos.Map).Zona = Ciudad), "1", "0"))
Call SendData(SendTarget.toindex, userindex, 0, "NOC" & IIf(DeNoche, "1", "0"))

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim loopC As Long

For loopC = 1 To LastUser
    If UserList(loopC).flags.UserLogged And UserList(loopC).ConnID >= 0 And UserList(loopC).ConnIDValida Then
        If UserList(loopC).flags.Privilegios < PlayerType.UserSupport Then
            Call CloseSocket(loopC)
        End If
    End If
Next loopC

End Sub

