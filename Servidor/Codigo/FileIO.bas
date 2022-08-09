Attribute VB_Name = "ES"
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
Sub LoadUserAccount(ByVal PJinit As String)
On Error Resume Next
PJEnCuenta = GetVar(CharPath & "\" & PJinit, "INIT", "Head") & "," & GetVar(CharPath & "\" & PJinit, "INIT", "Body") & "," & _
GetVar(CharPath & "\" & PJinit, "INIT", "Arma") & "," & GetVar(CharPath & "\" & PJinit, "INIT", "Escudo") & "," & GetVar(CharPath & "\" & PJinit, "INIT", "Casco") & "," & GetVar(CharPath & "\" & PJinit, "STATS", "ELV") & "," & GetVar(CharPath & "\" & PJinit, "INIT", "Clase") & "," & GetVar(CharPath & "\" & PJinit, "FLAGS", "Muerto") & "," & GetVar(CharPath & "\" & PJinit, "INIT", "Raza")
End Sub
Function EstaEnRing(ByVal userindex As Integer) As Boolean

    Dim X As Byte
    Dim Y As Byte
    
        For X = 74 To 89
            For Y = 19 To 37
                If UserList(userindex).Pos.Map = 26 And UserList(userindex).Pos.X = X And UserList(userindex).Pos.Y = Y Then
                    EstaEnRing = True
                    Exit Function
                End If
            Next Y
        Next X

    EstaEnRing = False
End Function
Public Sub CargarSpawnList()
    Dim n As Integer, loopC As Integer
    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador
    For loopC = 1 To n
        SpawnList(loopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopC))
        SpawnList(loopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopC)
    Next loopC
    
End Sub
Function IsAdministrator(ByVal Name As String) As Boolean

    If UCase$(Name) = "KING" Then
        IsAdministrator = True
        Exit Function
    End If
    
    IsAdministrator = False
End Function
Function IsCoordination(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
 
NumWizs = val(GetVar(IniPath & "Staff.ini", "STAFF", "Coordinations"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Staff.ini", "Coordinations", "Coordination" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        IsCoordination = True
        Exit Function
    End If
Next WizNum
IsCoordination = False
End Function
Function IsDevelopment(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
 
NumWizs = val(GetVar(IniPath & "Staff.ini", "STAFF", "Developments"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Staff.ini", "Developments", "Development" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        IsDevelopment = True
        Exit Function
    End If
Next WizNum
IsDevelopment = False
End Function
Function IsTournamentManager(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
 
NumWizs = val(GetVar(IniPath & "Staff.ini", "STAFF", "TournamentManagers"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Staff.ini", "TournamentManagers", "TournamentManager" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        IsTournamentManager = True
        Exit Function
    End If
Next WizNum
IsTournamentManager = False
End Function
Function IsEventManager(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
 
NumWizs = val(GetVar(IniPath & "Staff.ini", "STAFF", "EventManagers"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Staff.ini", "EventManagers", "EventManager" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        IsEventManager = True
        Exit Function
    End If
Next WizNum
IsEventManager = False
End Function
Function IsUserSupport(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
 
NumWizs = val(GetVar(IniPath & "Staff.ini", "STAFF", "UserSupports"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Staff.ini", "UserSupports", "UserSupport" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        IsUserSupport = True
        Exit Function
    End If
Next WizNum
IsUserSupport = False
 
End Function
Public Sub CargarPremiosList()
        Dim p As Integer, loopC As Integer
        p = val(GetVar(App.Path & "\Dat\Premios.dat", "INIT", "NumPremios"))
   
        ReDim PremiosList(p) As tPremiosCanjes
       
           
        For loopC = 1 To p
            PremiosList(loopC).ObjName = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "Nombre")
            PremiosList(loopC).ObjIndexP = val(GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "NumObj"))
            PremiosList(loopC).ObjRequiere = val(GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "Requiere"))
            PremiosList(loopC).ObjMaxAt = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "AtaqueMaximo")
            PremiosList(loopC).ObjMinAt = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "AtaqueMinimo")
            PremiosList(loopC).ObjMindef = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "DefensaMinima")
            PremiosList(loopC).ObjMaxdef = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "DefensaMaxima")
            PremiosList(loopC).ObjMinAtMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "AtaqueMagicoMinimo")
            PremiosList(loopC).ObjMaxAtMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "AtaqueMagicoMaximo")
            PremiosList(loopC).ObjMinDefMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "DefensaMagicaMinima")
            PremiosList(loopC).ObjMaxDefMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "DefensaMagicaMaxima")
            PremiosList(loopC).ObjDescripcion = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopC, "Descripcion")
        Next loopC
End Sub
Public Sub CargarCofresRandom()
    Dim n As Integer, loopC As Integer, loopX As Integer
    n = val(GetVar(App.Path & "\Dat\Cofres.dat", "INIT", "NumCofres"))
    
    ReDim CofresAzar(n) As AzarCofres
    
    For loopC = 1 To n
        CofresAzar(loopC).CantObjs = val(GetVar(App.Path & "\Dat\Cofres.dat", "COFRE" & loopC, "CantObjs"))
        
        ReDim CofresAzar(loopC).ObjIndex(CofresAzar(loopC).CantObjs) As Integer
        ReDim CofresAzar(loopC).ObjAmount(CofresAzar(loopC).CantObjs) As Integer
        ReDim CofresAzar(loopC).ObjProbability(CofresAzar(loopC).CantObjs) As Integer
        
        For loopX = 1 To CofresAzar(loopC).CantObjs
            CofresAzar(loopC).ObjIndex(loopX) = val(ReadField(1, GetVar(App.Path & "\Dat\Cofres.dat", "COFRE" & loopC, "Obj" & loopX), 45))
            CofresAzar(loopC).ObjAmount(loopX) = val(ReadField(2, GetVar(App.Path & "\Dat\Cofres.dat", "COFRE" & loopC, "Obj" & loopX), 45))
            CofresAzar(loopC).ObjProbability(loopX) = val(ReadField(3, GetVar(App.Path & "\Dat\Cofres.dat", "COFRE" & loopC, "Obj" & loopX), 45))
        Next loopX
    Next loopC
    
End Sub
Public Sub CargarDonaciones()
    Dim p As Integer, loopC As Integer
    
        'Cargamos las 'novedades'
        p = val(GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "INIT", "NumPremios"))
        ReDim DonationList(p) As tDonaciones
       
           
        For loopC = 1 To p
            DonationList(loopC).ObjName = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Nombre")
            DonationList(loopC).ObjValor = val(GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Valor"))
            DonationList(loopC).NumObjs = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "NumObjetos")
            DonationList(loopC).Body = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Body")
            DonationList(loopC).Arma = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Arma")
            DonationList(loopC).Escudo = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Escudo")
            DonationList(loopC).Casco = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Casco")
            DonationList(loopC).Aura = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Aura")
            
            DonationList(loopC).GrhIndex = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "GrhIndex")
            
            DonationList(loopC).Tuniquita = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Tuniquita")
            DonationList(loopC).Tuniquita2 = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Tuniquita2")
            DonationList(loopC).Tuniquita3 = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Tuniquita3")
            DonationList(loopC).Tuniquita4 = GetVar(App.Path & "\Dat\ItemsDonaciones.dat", "ITEM" & loopC, "Tuniquita4")
        Next loopC
        
        
End Sub
Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
Dim n As Integer, cad As String, Tam As Long
n = FreeFile(1)
Open Name For Input As #n
Tam = 0
Do While Not EOF(n)
    Tam = Tam + 1
    Line Input #n, cad
Loop
Close n
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim n As Integer, i As Integer
n = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #n

For i = 1 To UBound(ForbidenNames)
    Line Input #n, ForbidenNames(i)
Next i

Close n

End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
Hechizos(Hechizo).Particle_Speed = CSng(val(Leer.GetValue("Hechizo" & Hechizo, "ParticulaVelocidad")))
Hechizos(Hechizo).Particle_Index = CInt(val(Leer.GetValue("Hechizo" & Hechizo, "ParticulaIndex")))
    
    Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    If Hechizos(Hechizo).Tipo = 6 Then
       
          Hechizos(Hechizo).MaxDef1 = Leer.GetValue("Hechizo" & Hechizo, "MaxDefensa")
          Hechizos(Hechizo).MinDef1 = Leer.GetValue("Hechizo" & Hechizo, "MinDefensa")
       
    End If
    Hechizos(Hechizo).WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).ActivaNobleza = val(Leer.GetValue("Hechizo" & Hechizo, "ActivaNobleza"))
    Hechizos(Hechizo).BacuNecesario = val(Leer.GetValue("Hechizo" & Hechizo, "Staff"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    Hechizos(Hechizo).ExclusivoClase = UCase$(Leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase"))
    Hechizos(Hechizo).ExclusivoClasedos = UCase$(Leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase2"))
    Hechizos(Hechizo).ProhibidoClase = UCase$(Leer.GetValue("Hechizo" & Hechizo, "ProhibidoClase"))
    
    Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNPC = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    Hechizos(Hechizo).CuartaJerarquia = val(Leer.GetValue("Hechizo" & Hechizo, "CuartaJerarquia"))
    
    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).PortalMap = val(Leer.GetValue("Hechizo" & Hechizo, "PortalMap"))
    Hechizos(Hechizo).PortalX = val(Leer.GetValue("Hechizo" & Hechizo, "PortalX"))
    Hechizos(Hechizo).PortalY = val(Leer.GetValue("Hechizo" & Hechizo, "PortalY"))
    Hechizos(Hechizo).Telepo = val(Leer.GetValue("Hechizo" & Hechizo, "Telepo"))
    Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
    
    Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
Next Hechizo

Set Leer = Nothing
Exit Sub

Errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub
Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim TempInt As Integer
    Dim loopC As Long
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
            
    Put FreeFileMap, , MapInfo(Map).MapVersion
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(Map, X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).Graphic(2) Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).Graphic(3) Then ByFlags = ByFlags Or 4
                If MapData(Map, X, Y).Graphic(4) Then ByFlags = ByFlags Or 8
                If MapData(Map, X, Y).trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(Map, X, Y).Graphic(1)
                
                For loopC = 2 To 4
                    If MapData(Map, X, Y).Graphic(loopC) Then _
                        Put FreeFileMap, , MapData(Map, X, Y).Graphic(loopC)
                Next loopC
                
                If MapData(Map, X, Y).trigger Then _
                    Put FreeFileMap, , CInt(MapData(Map, X, Y).trigger)
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                   If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        MapData(Map, X, Y).OBJInfo.ObjIndex = 0
                        MapData(Map, X, Y).OBJInfo.Amount = 0
                    End If
                End If
    
                If MapData(Map, X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).NpcIndex Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(Map, X, Y).TileExit.Map Then
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.X
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Y
                End If
                
                If MapData(Map, X, Y).NpcIndex Then _
                    Put FreeFileInf, , Npclist(MapData(Map, X, Y).NpcIndex).Numero
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
                End If
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    'write .dat file
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Name", MapInfo(Map).Name)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "MusicNum", MapInfo(Map).Music)
    Call WriteVar(MAPFILE & ".dat", "mapa" & Map, "MagiaSinefecto", MapInfo(Map).MagiaSinEfecto)

    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Terreno", MapInfo(Map).Terreno)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Zona", MapInfo(Map).Zona)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Restringir", MapInfo(Map).Restringir)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "BackUp", str(MapInfo(Map).BackUp))

    If MapInfo(Map).Pk Then
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "0")
    Else
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "1")
    End If

End Sub
Sub LoadArmasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadObjCarpintero()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To n) As Integer

For lc = 1 To n
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo Errhandler

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
    ObjData(Object).Aura = val(Leer.GetValue("OBJ" & Object, "CreaAura"))
    
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otcASCO
            ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).DosManos = val(Leer.GetValue("OBJ" & Object, "DosManos"))
            ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otHerramientas
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
    End Select
    
    ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).TipoCofre = val(Leer.GetValue("OBJ" & Object, "TipoCofre"))
    
    ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    ObjData(Object).lvl = val(Leer.GetValue("OBJ" & Object, "LVL"))
     
    ObjData(Object).DañoMagicoMax = val(Leer.GetValue("OBJ" & Object, "DañoMagicoMax"))
    ObjData(Object).DañoMagicoMin = val(Leer.GetValue("OBJ" & Object, "DañoMagicoMax"))
    
    ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    
    ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    
    ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    ObjData(Object).Inmoviliza = val(Leer.GetValue("OBJ" & Object, "Inmoviliza"))
    ObjData(Object).AntiLimpieza = val(Leer.GetValue("OBJ" & Object, "AntiLimpieza"))
    ObjData(Object).CristalesMax = val(Leer.GetValue("OBJ" & Object, "CristalesMax"))
    ObjData(Object).CristalesMin = val(Leer.GetValue("OBJ" & Object, "CristalesMin"))
    ObjData(Object).Intransferible = val(Leer.GetValue("OBJ" & Object, "Intransferible"))
    ObjData(Object).ItemDios = val(Leer.GetValue("OBJ" & Object, "ItemDios"))
    ObjData(Object).Dios = Leer.GetValue("OBJ" & Object, "Dios")
    
    ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    ObjData(Object).PuertaDoble = val(Leer.GetValue("OBJ" & Object, "PuertaDoble"))
    ObjData(Object).Porton = val(Leer.GetValue("OBJ" & Object, "Porton"))
    ObjData(Object).RejaForta = val(Leer.GetValue("OBJ" & Object, "RejaForta"))
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).Texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Leer.GetValue("OBJ" & Object, "CP" & i)
    Next i
    
    ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then
        ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
        ObjData(Object).Piedras = val(Leer.GetValue("OBJ" & Object, "Piedras"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
Next Object

Set Leer = Nothing

Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserStats(ByVal userindex As Integer, ByRef UserFile As clsIniReader)

Dim loopC As Integer


For loopC = 1 To NUMATRIBUTOS
  UserList(userindex).Stats.UserAtributos(loopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & loopC))
  UserList(userindex).Stats.UserAtributosBackUP(loopC) = UserList(userindex).Stats.UserAtributos(loopC)
Next loopC

For loopC = 1 To NUMSKILLS
  UserList(userindex).Stats.UserSkills(loopC) = CInt(UserFile.GetValue("SKILLS", "SK" & loopC))
Next loopC

For loopC = 1 To MAXUSERHECHIZOS
  UserList(userindex).Stats.UserHechizos(loopC) = CInt(UserFile.GetValue("Hechizos", "H" & loopC))
Next loopC

UserList(userindex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
UserList(userindex).Stats.PuntosTorneo = CLng(UserFile.GetValue("STATS", "PuntosTorneo"))
UserList(userindex).Stats.PuntosDonacion = CLng(UserFile.GetValue("STATS", "PuntosDonacion"))
UserList(userindex).Stats.TSPoints = CLng(UserFile.GetValue("STATS", "TSPoints"))
UserList(userindex).Stats.ParejasGanadas = CLng(UserFile.GetValue("STATS", "ParejasGanadas"))
UserList(userindex).Stats.ParejasPerdidas = CLng(UserFile.GetValue("STATS", "ParejasPerdidas"))
UserList(userindex).Stats.MuertesUser = CLng(UserFile.GetValue("STATS", "MuertesUser"))
UserList(userindex).Stats.DuelosGanados = CLng(UserFile.GetValue("STATS", "DuelosGanados"))
UserList(userindex).Stats.DuelosPerdidos = CLng(UserFile.GetValue("STATS", "DuelosPerdidos"))
UserList(userindex).Stats.TrofOro = CLng(UserFile.GetValue("STATS", "TrofOro"))
UserList(userindex).Stats.TorneosParticipados = CLng(UserFile.GetValue("STATS", "TorneosParticipados"))
UserList(userindex).Stats.TrofPlata = CLng(UserFile.GetValue("STATS", "TrofPlata"))
UserList(userindex).Stats.TrofBronce = CLng(UserFile.GetValue("STATS", "TrofBronce"))
UserList(userindex).Stats.MedOro = CLng(UserFile.GetValue("STATS", "MedOro"))

UserList(userindex).Stats.MET = CInt(UserFile.GetValue("STATS", "MET"))
UserList(userindex).Stats.MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
UserList(userindex).Stats.MinHP = CInt(UserFile.GetValue("STATS", "MinHP"))

UserList(userindex).Stats.Reputacione = CLng(UserFile.GetValue("STATS", "Reputacione"))
UserList(userindex).Stats.FIT = CInt(UserFile.GetValue("STATS", "FIT"))
UserList(userindex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
UserList(userindex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

UserList(userindex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
UserList(userindex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

UserList(userindex).Stats.MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
UserList(userindex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))

UserList(userindex).Stats.MaxAGU = CInt(UserFile.GetValue("STATS", "MaxAGU"))
UserList(userindex).Stats.MinAGU = CInt(UserFile.GetValue("STATS", "MinAGU"))

UserList(userindex).Stats.MaxHam = CInt(UserFile.GetValue("STATS", "MaxHAM"))
UserList(userindex).Stats.MinHam = CInt(UserFile.GetValue("STATS", "MinHAM"))

UserList(userindex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

UserList(userindex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
UserList(userindex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
UserList(userindex).Stats.ELV = CLng(UserFile.GetValue("STATS", "ELV"))

UserList(userindex).Stats.UsuariosMatados = CInt(UserFile.GetValue("MUERTES", "UserMuertes"))
UserList(userindex).Stats.CriminalesMatados = CInt(UserFile.GetValue("MUERTES", "CrimMuertes"))
UserList(userindex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

UserList(userindex).ConsejoInfo.PertAlCons = CByte(UserFile.GetValue("CONSEJO", "PERTENECE"))
UserList(userindex).ConsejoInfo.LiderConsejo = CByte(UserFile.GetValue("CONSEJO", "LIDERCONSEJO"))
UserList(userindex).ConsejoInfo.PertAlConsCaos = CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS"))
UserList(userindex).ConsejoInfo.LiderConsejoCaos = CByte(UserFile.GetValue("CONSEJO", "LIDERCONSEJOCAOS"))


UserList(userindex).CofreDios.Cant = val(UserFile.GetValue("COFREDIOS", "Cant"))
For loopC = 1 To 4
    UserList(userindex).CofreDios.Item(loopC) = val(UserFile.GetValue("COFREDIOS", "Item" & loopC))
Next loopC


Dim NCTemporal As String

For loopC = 1 To 30
    UserList(userindex).flags.Correo(loopC) = UserFile.GetValue("CORREO", "CORREONUM" & loopC)
    NCTemporal = ReadField(loopC, UserFile.GetValue("CORREO", "NUECORREOS"), Asc(","))
    UserList(userindex).flags.NueCorreos(loopC) = ReadField(2, NCTemporal, Asc("-"))
Next loopC
    
    UserList(userindex).flags.NumCorreos = UserFile.GetValue("CORREO", "NUMCORREOS")

End Sub
Sub LoadUserStatus(ByVal userindex As Integer, ByRef UserFile As clsIniReader)
UserList(userindex).StatusMith.EsStatus = CByte(UserFile.GetValue("STATUS", "EsStatus"))
UserList(userindex).StatusMith.EligioStatus = CByte(UserFile.GetValue("STATUS", "Eligio"))
End Sub
Sub LoadUserInit(ByVal userindex As Integer, ByRef UserFile As clsIniReader)

Dim loopC As Long
Dim ln As String

UserList(userindex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
UserList(userindex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
UserList(userindex).Faccion.CiudadanosMatados = CDbl(UserFile.GetValue("FACCIONES", "CiudMatados"))
UserList(userindex).Faccion.CriminalesMatados = CDbl(UserFile.GetValue("FACCIONES", "CrimMatados"))
UserList(userindex).Faccion.NeutralesMatados = CDbl(UserFile.GetValue("FACCIONES", "NeutrMatados"))
UserList(userindex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
UserList(userindex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
UserList(userindex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
UserList(userindex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
UserList(userindex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
UserList(userindex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
UserList(userindex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))

UserList(userindex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
UserList(userindex).flags.DeseoRecibirMSJ = CByte(UserFile.GetValue("FLAGS", "MsjPrivado"))
UserList(userindex).flags.Emoticons = CByte(UserFile.GetValue("FLAGS", "Emoticons"))
UserList(userindex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

UserList(userindex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
UserList(userindex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
UserList(userindex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
UserList(userindex).flags.Pareja = CStr(UserFile.GetValue("FLAGS", "Pareja"))
UserList(userindex).flags.Llegolvlmax = CByte(UserFile.GetValue("FLAGS", "Llegolvlmax"))
UserList(userindex).flags.llegolvl50 = CByte(UserFile.GetValue("FLAGS", "Llegolvl50"))

UserList(userindex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
UserList(userindex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
If UserList(userindex).flags.Paralizado = 1 Then
    UserList(userindex).Counters.Paralisis = IntervaloParalizado
End If
UserList(userindex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
UserList(userindex).flags.Montando = CByte(UserFile.GetValue("FLAGS", "Montando"))
UserList(userindex).flags.estado = CByte(UserFile.GetValue("FLAGS", "Estado"))
UserList(userindex).flags.EsNoble = CByte(UserFile.GetValue("FLAGS", "EsNoble"))
UserList(userindex).flags.EsPremium = CByte(UserFile.GetValue("PREMIUM", "EsPremium"))
UserList(userindex).flags.VencePremium = UserFile.GetValue("PREMIUM", "Vencimiento")
UserList(userindex).flags.CaballerodelDragon = CByte(UserFile.GetValue("FLAGS", "CaballerodelDragon"))
UserList(userindex).flags.PJerarquia = CByte(UserFile.GetValue("FLAGS", "PJerarquia"))
UserList(userindex).flags.SJerarquia = CByte(UserFile.GetValue("FLAGS", "SJerarquia"))
UserList(userindex).flags.TJerarquia = CByte(UserFile.GetValue("FLAGS", "TJerarquia"))
UserList(userindex).flags.CJerarquia = CByte(UserFile.GetValue("FLAGS", "CJerarquia"))
UserList(userindex).flags.CJerarquiaC = CByte(UserFile.GetValue("FLAGS", "CJerarquiaC"))
UserList(userindex).flags.Transformado = CByte(UserFile.GetValue("FLAGS", "Transformado"))
UserList(userindex).flags.Questeando = CByte(UserFile.GetValue("FLAGS", "Questeando"))
UserList(userindex).flags.MuereQuest(1) = CByte(UserFile.GetValue("FLAGS", "MuereQuest1"))
UserList(userindex).flags.MuereQuest(2) = CByte(UserFile.GetValue("FLAGS", "MuereQuest2"))
UserList(userindex).flags.MuereQuest(3) = CByte(UserFile.GetValue("FLAGS", "MuereQuest3"))
UserList(userindex).flags.UserNumQuest = CByte(UserFile.GetValue("FLAGS", "UserNumQuest"))
UserList(userindex).flags.QuestCompletadas = CLng(UserFile.GetValue("FLAGS", "QuestCompletadas"))
UserList(userindex).flags.GuerrasGanadas = CLng(UserFile.GetValue("FLAGS", "GuerrasGanadas"))
UserList(userindex).flags.GuerrasPerdidas = CLng(UserFile.GetValue("FLAGS", "GuerrasPerdidas"))
UserList(userindex).flags.MVPMatados = CLng(UserFile.GetValue("FLAGS", "MVPMatados"))
UserList(userindex).flags.CvcsGanados = UserFile.GetValue("FLAGS", "CvcsGanados")
UserList(userindex).flags.AlmasContenidas = CLng(UserFile.GetValue("FLAGS", "AlmasContenidas"))
UserList(userindex).flags.AlmasOfrecidas = CLng(UserFile.GetValue("FLAGS", "AlmasOfrecidas"))
UserList(userindex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
UserList(userindex).flags.Stopped = CByte(UserFile.GetValue("FLAGS", "STOP"))
UserList(userindex).flags.JerarquiaDios = CByte(UserFile.GetValue("FLAGS", "JerarquiaDios"))
UserList(userindex).flags.SirvienteDeDios = UserFile.GetValue("FLAGS", "SirvienteDeDios")

UserList(userindex).flags.PuedeRetirarObj = CByte(UserFile.GetValue("FLAGS", "PuedeRetirarObj"))
UserList(userindex).flags.PuedeRetirarOro = CByte(UserFile.GetValue("FLAGS", "PuedeRetirarOro"))

UserList(userindex).email = UserFile.GetValue("CONTACTO", "Email")
UserList(userindex).Char.Account = UserFile.GetValue("CHAR", "Cuenta")

UserList(userindex).InventorySlots = UserFile.GetValue("CHAR", "InventorySlots")

UserList(userindex).Genero = UserFile.GetValue("INIT", "Genero")
UserList(userindex).clase = UserFile.GetValue("INIT", "Clase")
UserList(userindex).NickMascota = UserFile.GetValue("INIT", "NickMascota")
UserList(userindex).Raza = UserFile.GetValue("INIT", "Raza")
UserList(userindex).Hogar = UserFile.GetValue("INIT", "Hogar")
UserList(userindex).Char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))

UserList(userindex).Password = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "SEGURIDAD", "CodeX")

UserList(userindex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
UserList(userindex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
UserList(userindex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
UserList(userindex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
UserList(userindex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
UserList(userindex).OrigChar.Heading = eHeading.SOUTH


UserList(userindex).Misiones.Dia = UserFile.GetValue("MISIONES", "Dia")
UserList(userindex).Misiones.NumeroMision = UserFile.GetValue("MISIONES", "NumeroMision")
UserList(userindex).Misiones.ConteoUser = UserFile.GetValue("MISIONES", "ConteoUser")
UserList(userindex).Misiones.Completada = UserFile.GetValue("MISIONES", "Completada")


If UserList(userindex).flags.Muerto = 0 Then
    UserList(userindex).Char = UserList(userindex).OrigChar
Else

 If UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Or UserList(userindex).StatusMith.EsStatus = 5 Then
    UserList(userindex).Char.Body = iCuerpoMuertoA
    UserList(userindex).Char.Head = iCabezaMuertoA
 ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Or UserList(userindex).StatusMith.EsStatus = 6 Then
    UserList(userindex).Char.Body = iCuerpoMuertoH
    UserList(userindex).Char.Head = iCabezaMuertoH
 Else
    UserList(userindex).Char.Body = iCuerpoMuertoN
    UserList(userindex).Char.Head = iCabezaMuertoN
 End If
 
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.CascoAnim = NingunCasco
End If

UserList(userindex).Bon1 = UserFile.GetValue("INIT", "Bon1")
UserList(userindex).Bon2 = UserFile.GetValue("INIT", "Bon2")
UserList(userindex).Bon3 = UserFile.GetValue("INIT", "Bon3")

UserList(userindex).UltimoLogeo = UserFile.GetValue("INIT", "UltimoLogeo")
UserList(userindex).UltimaDenuncia = UserFile.GetValue("INIT", "UltimaDenuncia")
UserList(userindex).PrimeraDenuncia = UserFile.GetValue("INIT", "PrimeraDenuncia")

UserList(userindex).Desc = UserFile.GetValue("INIT", "Desc")


UserList(userindex).Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
UserList(userindex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
UserList(userindex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

UserList(userindex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(userindex).BancoInvent.NroItems = CInt(GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = (GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "Obj" & loopC))
    UserList(userindex).BancoInvent.Object(loopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(userindex).BancoInvent.Object(loopC).Amount = CInt(ReadField(2, ln, 45))
Next loopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For loopC = 1 To MAX_INVENTORY_SLOTS
    ln = UserFile.GetValue("Inventory", "Obj" & loopC)
    UserList(userindex).Invent.Object(loopC).ObjIndex = CInt(ReadField(1, ln, 45))
    
        If CInt(ReadField(2, ln, 45)) <= 0 Then
            UserList(userindex).Invent.Object(loopC).Amount = 1
        Else
            UserList(userindex).Invent.Object(loopC).Amount = CInt(ReadField(2, ln, 45))
        End If
        
    UserList(userindex).Invent.Object(loopC).Equipped = CByte(ReadField(3, ln, 45))
Next loopC

'Obtiene el indice-objeto del arma
UserList(userindex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
If UserList(userindex).Invent.WeaponEqpSlot > 0 Then
    UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(userindex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
If UserList(userindex).Invent.ArmourEqpSlot > 0 Then
    UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.ArmourEqpSlot).ObjIndex
    UserList(userindex).flags.Desnudo = 0
Else
    UserList(userindex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(userindex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
If UserList(userindex).Invent.EscudoEqpSlot > 0 Then
    UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(userindex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
If UserList(userindex).Invent.CascoEqpSlot > 0 Then
    UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(userindex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
If UserList(userindex).Invent.BarcoSlot > 0 Then
    UserList(userindex).Invent.BarcoObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(userindex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
If UserList(userindex).Invent.MunicionEqpSlot > 0 Then
    UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto herramienta
UserList(userindex).Invent.HerramientaEqpSlot = CInt(UserFile.GetValue("Inventory", "HerramientaSlot"))
If UserList(userindex).Invent.HerramientaEqpSlot > 0 Then
    UserList(userindex).Invent.HerramientaEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.HerramientaEqpSlot).ObjIndex
End If

'Obtiene la lista de amigos
    For loopC = 1 To 20
        UserList(userindex).flags.NombreAmigo(loopC) = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "AMIGOS", "A" & loopC)
        UserList(userindex).flags.AmigoDesadmitido(loopC) = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "AMIGOS", "AD" & loopC)
    Next loopC
    
UserList(userindex).Stats.Banco = GetVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "" & UserList(userindex).Accounted & "", "BANCO")

UserList(userindex).NroMacotas = 0

ln = UserFile.GetValue("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(userindex).GuildIndex = CInt(ln)
Else
    UserList(userindex).GuildIndex = 0
End If

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
        Else
            tFileName = App.Path & MapPath & "Mapa" & Map
        End If
        
        Call CargarMapa(Map, tFileName)
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
        
        frmCargando.Label1(2).caption = "Cargando Mapas (" & Map & "/" & NumMaps & ")"
        frmCargando.Label2.caption = "[" & Round(frmCargando.Image1.Width / 60) & "%]"
        frmCargando.Image1.Width = frmCargando.Image1.Width + (5985 / NumMaps)
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub
Public Sub CargarMapa(ByVal Map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer
      
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(Map).MapVersion
    Get #FreeFileMap, , MiCabecera
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    
    'inf Header
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.dat file
            Get FreeFileMap, , ByFlags

            If ByFlags And 1 Then
                MapData(Map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(Map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempInt
                MapData(Map, X, Y).trigger = TempInt
            End If
            
            If ByFlags And 32 Then Get FreeFileMap, , MapData(Map, X, Y).particle_group_index
            
            'www.rincondelao.com.ar
            If ByFlags And 64 Then
                Get FreeFileMap, , MapData(Map, X, Y).base_light(0)
                Get FreeFileMap, , MapData(Map, X, Y).base_light(1)
                Get FreeFileMap, , MapData(Map, X, Y).base_light(2)
                Get FreeFileMap, , MapData(Map, X, Y).base_light(3)
                
                If MapData(Map, X, Y).base_light(0) Then _
                    Get FreeFileMap, , MapData(Map, X, Y).light_value(0)
                
                If MapData(Map, X, Y).base_light(1) Then _
                    Get FreeFileMap, , MapData(Map, X, Y).light_value(1)
                'www.rincondelao.com.ar
                If MapData(Map, X, Y).base_light(2) Then _
                    Get FreeFileMap, , MapData(Map, X, Y).light_value(2)
                
                If MapData(Map, X, Y).base_light(3) Then _
                    Get FreeFileMap, , MapData(Map, X, Y).light_value(3)
            End If
            
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.X
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(Map, X, Y).NpcIndex
                
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else
                        npcfile = DatPath & "NPCs.dat"
                    End If

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                    Else
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    End If
                            
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                            
                    Call MakeNPCChar(SendTarget.ToMap, 0, 0, MapData(Map, X, Y).NpcIndex, 1, 1, 1)
                End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
            End If
        Next X
    Next Y
    
    
    Close FreeFileMap
    Close FreeFileInf
    
    MapInfo(Map).Name = GetVar(MAPFl & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(MAPFl & ".dat", "Mapa" & Map, "MusicNum")
    MapInfo(Map).MagiaSinEfecto = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "MagiaSinEfecto"))
    MapInfo(Map).NoEncriptarMP = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "NoEncriptarMP"))
    
    If val(GetVar(MAPFl & ".dat", "Mapa" & Map, "R")) = 0 Then
        MapInfo(Map).r = 200
    Else
        MapInfo(Map).r = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "R"))
    End If
    
    If val(GetVar(MAPFl & ".dat", "Mapa" & Map, "G")) = 0 Then
        MapInfo(Map).g = 200
    Else
        MapInfo(Map).g = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "G"))
    End If
    
    If val(GetVar(MAPFl & ".dat", "Mapa" & Map, "B")) = 0 Then
        MapInfo(Map).b = 200
    Else
        MapInfo(Map).b = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "B"))
    End If
    
    If val(GetVar(MAPFl & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(MAPFl & ".dat", "Mapa" & Map, "Terreno")
    MapInfo(Map).Zona = GetVar(MAPFl & ".dat", "Mapa" & Map, "Zona")
    MapInfo(Map).Restringir = GetVar(MAPFl & ".dat", "Mapa" & Map, "Restringir")
    MapInfo(Map).BackUp = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "BACKUP"))
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & "." & Err.Description)
End Sub

Sub LoadSini()

Dim Temporal As Long
Dim Temporal1 As Long
Dim loopC As Integer

Call SendData(SendTarget.ToAll, 0, 0, "||665")

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

'Misc
CrcSubKey = val(GetVar(IniPath & "Server.ini", "INIT", "CrcSubKey"))

ServerIp = GetVar(IniPath & "Server.ini", "INIT", "ServerIp")
Temporal = InStr(1, ServerIp, ".")
Temporal1 = (mid(ServerIp, 1, Temporal - 1) And &H7F) * 16777216
ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + mid(ServerIp, 1, Temporal - 1) * 65536
ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + mid(ServerIp, 1, Temporal - 1) * 256
ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))

MixedKey = (Temporal1 + ServerIp) Xor &H65F64B42

Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

IntervaloNpcPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = IntervaloNpcPuedeAtacar

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

frmMain.CmdExec.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

IntervaloAutoReiniciar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloAutoReiniciar"))
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

Tanaris.Map = 1
Tanaris.X = 50
Tanaris.Y = 50

Call SendData(SendTarget.ToAll, 0, 0, "||666")

End Sub
Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub
Sub SaveUser(ByVal userindex As Integer, ByVal UserFile As String)
On Error GoTo Errhandler

Dim OldUserHead As Long

'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(userindex).clase = "" Or UserList(userindex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(userindex).Name)
    Exit Sub
End If


If UserList(userindex).flags.Mimetizado = 1 Then
    UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
End If



If FileExist(UserFile, vbNormal) Then
       If UserList(userindex).flags.Muerto = 1 Then
        OldUserHead = UserList(userindex).Char.Head
        UserList(userindex).Char.Head = CStr(GetVar(UserFile, "INIT", "Head"))
       End If
'       Kill UserFile
End If

Dim loopC As Integer

Call WriteVar(UserFile, "FLAGS", "MuereQuest1", CStr(UserList(userindex).flags.MuereQuest(1)))
Call WriteVar(UserFile, "FLAGS", "MuereQuest2", CStr(UserList(userindex).flags.MuereQuest(2)))
Call WriteVar(UserFile, "FLAGS", "MuereQuest3", CStr(UserList(userindex).flags.MuereQuest(3)))

Call WriteVar(UserFile, "FLAGS", "Questeando", CStr(UserList(userindex).flags.Questeando))
Call WriteVar(UserFile, "FLAGS", "UserNumQuest", CStr(UserList(userindex).flags.UserNumQuest))
Call WriteVar(UserFile, "FLAGS", "GuerrasGanadas", CStr(UserList(userindex).flags.GuerrasGanadas))
Call WriteVar(UserFile, "FLAGS", "GuerrasPerdidas", CStr(UserList(userindex).flags.GuerrasPerdidas))
Call WriteVar(UserFile, "FLAGS", "MVPMatados", CStr(UserList(userindex).flags.MVPMatados))
Call WriteVar(UserFile, "FLAGS", "QuestCompletadas", CStr(UserList(userindex).flags.QuestCompletadas))
Call WriteVar(UserFile, "FLAGS", "Muerto", CStr(UserList(userindex).flags.Muerto))
Call WriteVar(UserFile, "FLAGS", "CvcsGanados", UserList(userindex).flags.CvcsGanados)
Call WriteVar(UserFile, "FLAGS", "AlmasContenidas", CStr(UserList(userindex).flags.AlmasContenidas))
Call WriteVar(UserFile, "FLAGS", "AlmasOfrecidas", CStr(UserList(userindex).flags.AlmasOfrecidas))
Call WriteVar(UserFile, "FLAGS", "STOP", CStr(UserList(userindex).flags.Stopped))
Call WriteVar(UserFile, "FLAGS", "MsjPrivado", CStr(UserList(userindex).flags.DeseoRecibirMSJ))
Call WriteVar(UserFile, "FLAGS", "Emoticons", CStr(UserList(userindex).flags.Emoticons))
Call WriteVar(UserFile, "FLAGS", "Llegolvlmax", CStr(UserList(userindex).flags.Llegolvlmax))
Call WriteVar(UserFile, "FLAGS", "Llegolvl50", CStr(UserList(userindex).flags.llegolvl50))
Call WriteVar(UserFile, "FLAGS", "Escondido", CStr(UserList(userindex).flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "Estado", CStr(UserList(userindex).flags.estado))
Call WriteVar(UserFile, "FLAGS", "EsNoble", CStr(UserList(userindex).flags.EsNoble))
Call WriteVar(UserFile, "PREMIUM", "EsPremium", CStr(UserList(userindex).flags.EsPremium))
Call WriteVar(UserFile, "PREMIUM", "Vencimiento", CStr(UserList(userindex).flags.VencePremium))
Call WriteVar(UserFile, "FLAGS", "CaballerodelDragon", CStr(UserList(userindex).flags.CaballerodelDragon))
Call WriteVar(UserFile, "FLAGS", "Hambre", CStr(UserList(userindex).flags.Hambre))
Call WriteVar(UserFile, "FLAGS", "Sed", CStr(UserList(userindex).flags.Sed))
Call WriteVar(UserFile, "FLAGS", "Desnudo", CStr(UserList(userindex).flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Pareja", CStr(UserList(userindex).flags.Pareja))
Call WriteVar(UserFile, "FLAGS", "Ban", CStr(UserList(userindex).flags.Ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", CStr(UserList(userindex).flags.Navegando))
Call WriteVar(UserFile, "FLAGS", "Montando", CStr(UserList(userindex).flags.Montando))
Call WriteVar(UserFile, "FLAGS", "PJerarquia", CStr(UserList(userindex).flags.PJerarquia))
Call WriteVar(UserFile, "FLAGS", "SJerarquia", CStr(UserList(userindex).flags.SJerarquia))
Call WriteVar(UserFile, "FLAGS", "TJerarquia", CStr(UserList(userindex).flags.TJerarquia))
Call WriteVar(UserFile, "FLAGS", "CJerarquia", CStr(UserList(userindex).flags.CJerarquia))
Call WriteVar(UserFile, "FLAGS", "CJerarquiaC", CStr(UserList(userindex).flags.CJerarquiaC))
Call WriteVar(UserFile, "FLAGS", "Transformado", CStr(UserList(userindex).flags.Transformado))
Call WriteVar(UserFile, "FLAGS", "SirvienteDeDios", UserList(userindex).flags.SirvienteDeDios)
Call WriteVar(UserFile, "FLAGS", "JerarquiaDios", CStr(UserList(userindex).flags.JerarquiaDios))

Call WriteVar(UserFile, "FLAGS", "PuedeRetirarObj", CStr(UserList(userindex).flags.PuedeRetirarObj))
Call WriteVar(UserFile, "FLAGS", "PuedeRetirarOro", CStr(UserList(userindex).flags.PuedeRetirarOro))

Call WriteVar(UserFile, "FLAGS", "Envenenado", CStr(UserList(userindex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", CStr(UserList(userindex).flags.Paralizado))

Call WriteVar(UserFile, "CONSEJO", "PERTENECE", CStr(UserList(userindex).ConsejoInfo.PertAlCons))
Call WriteVar(UserFile, "CONSEJO", "LIDERCONSEJO", CStr(UserList(userindex).ConsejoInfo.LiderConsejo))
Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", CStr(UserList(userindex).ConsejoInfo.PertAlConsCaos))
Call WriteVar(UserFile, "CONSEJO", "LIDERCONSEJOCAOS", CStr(UserList(userindex).ConsejoInfo.LiderConsejoCaos))


Call WriteVar(UserFile, "COUNTERS", "Pena", CStr(UserList(userindex).Counters.Pena))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", CStr(UserList(userindex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", CStr(UserList(userindex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", CStr(UserList(userindex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", CStr(UserList(userindex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "NeutrMatados", CStr(UserList(userindex).Faccion.NeutralesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", CStr(UserList(userindex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", CStr(UserList(userindex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", CStr(UserList(userindex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", CStr(UserList(userindex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", CStr(UserList(userindex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", CStr(UserList(userindex).Faccion.RecompensasReal))
Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", CStr(UserList(userindex).Faccion.Reenlistadas))

'¿Fueron modificados los atributos del usuario?
If Not UserList(userindex).flags.TomoPocion Then
    For loopC = 1 To UBound(UserList(userindex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & loopC, CStr(UserList(userindex).Stats.UserAtributos(loopC)))
    Next
Else
    For loopC = 1 To UBound(UserList(userindex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & loopC, CStr(UserList(userindex).Stats.UserAtributosBackUP(loopC)))
    Next
End If

For loopC = 1 To UBound(UserList(userindex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & loopC, CStr(UserList(userindex).Stats.UserSkills(loopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(userindex).email)

Call WriteVar(UserFile, "MISIONES", "Dia", UserList(userindex).Misiones.Dia)
Call WriteVar(UserFile, "MISIONES", "NumeroMision", UserList(userindex).Misiones.NumeroMision)
Call WriteVar(UserFile, "MISIONES", "ConteoUser", UserList(userindex).Misiones.ConteoUser)
Call WriteVar(UserFile, "MISIONES", "Completada", UserList(userindex).Misiones.Completada)

Call WriteVar(UserFile, "INIT", "Bon1", UserList(userindex).Bon1)
Call WriteVar(UserFile, "INIT", "Bon2", UserList(userindex).Bon2)
Call WriteVar(UserFile, "INIT", "Bon3", UserList(userindex).Bon3)

Call WriteVar(UserFile, "INIT", "UltimoLogeo", UserList(userindex).UltimoLogeo)
Call WriteVar(UserFile, "INIT", "UltimaDenuncia", UserList(userindex).UltimaDenuncia)
Call WriteVar(UserFile, "INIT", "PrimeraDenuncia", UserList(userindex).PrimeraDenuncia)

Call WriteVar(UserFile, "INIT", "Genero", UserList(userindex).Genero)
Call WriteVar(UserFile, "INIT", "Raza", UserList(userindex).Raza)
Call WriteVar(UserFile, "INIT", "Hogar", UserList(userindex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", UserList(userindex).clase)
Call WriteVar(UserFile, "INIT", "Desc", UserList(userindex).Desc)
Call WriteVar(UserFile, "INIT", "NickMascota", UserList(userindex).NickMascota)
Call WriteVar(UserFile, "INIT", "Password", UserList(userindex).Password)

Call WriteVar(UserFile, "INIT", "Heading", CStr(UserList(userindex).Char.Heading))
Call WriteVar(UserFile, "INIT", "Head", CStr(UserList(userindex).OrigChar.Head))

If UserList(userindex).flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", CStr(UserList(userindex).Char.Body))
End If

Call WriteVar(UserFile, "INIT", "Arma", CStr(UserList(userindex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", CStr(UserList(userindex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", CStr(UserList(userindex).Char.CascoAnim))

Call WriteVar(UserFile, "INIT", "LastIP", UserList(userindex).ip)
Call WriteVar(UserFile, "INIT", "Position", UserList(userindex).Pos.Map & "-" & UserList(userindex).Pos.X & "-" & UserList(userindex).Pos.Y)
Call WriteVar(UserFile, "INIT", "LastHD", UserList(userindex).hd)
Call WriteVar(UserFile, "CHAR", "Cuenta", UserList(userindex).Accounted)
Call WriteVar(UserFile, "CHAR", "InventorySlots", UserList(userindex).InventorySlots)


Call WriteVar(UserFile, "STATS", "GLD", CStr(UserList(userindex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "PuntosTorneo", CStr(UserList(userindex).Stats.PuntosTorneo))
Call WriteVar(UserFile, "STATS", "PuntosDonacion", CStr(UserList(userindex).Stats.PuntosDonacion))
Call WriteVar(UserFile, "STATS", "TSPoints", CStr(UserList(userindex).Stats.TSPoints))
Call WriteVar(UserFile, "STATS", "ParejasGanadas", CStr(UserList(userindex).Stats.ParejasGanadas))
Call WriteVar(UserFile, "STATS", "ParejasPerdidas", CStr(UserList(userindex).Stats.ParejasPerdidas))
Call WriteVar(UserFile, "STATS", "MuertesUser", CStr(UserList(userindex).Stats.MuertesUser))
Call WriteVar(UserFile, "STATS", "DuelosGanados", CStr(UserList(userindex).Stats.DuelosGanados))
Call WriteVar(UserFile, "STATS", "DuelosPerdidos", CStr(UserList(userindex).Stats.DuelosPerdidos))
Call WriteVar(UserFile, "STATS", "TorneosParticipados", CStr(UserList(userindex).Stats.TorneosParticipados))
Call WriteVar(UserFile, "STATS", "TrofOro", CStr(UserList(userindex).Stats.TrofOro))
Call WriteVar(UserFile, "STATS", "TrofPlata", CStr(UserList(userindex).Stats.TrofPlata))
Call WriteVar(UserFile, "STATS", "TrofBronce", CStr(UserList(userindex).Stats.TrofBronce))
Call WriteVar(UserFile, "STATS", "MedOro", CStr(UserList(userindex).Stats.MedOro))
Call WriteVar(UserFile, "STATUS", "EsStatus", CStr(UserList(userindex).StatusMith.EsStatus))
Call WriteVar(UserFile, "STATUS", "Eligio", CStr(UserList(userindex).StatusMith.EligioStatus))

Call WriteVar(UserFile, "STATS", "BANCO", UserList(userindex).Stats.Banco)

Call WriteVar(UserFile, "STATS", "MET", CStr(UserList(userindex).Stats.MET))
Call WriteVar(UserFile, "STATS", "MaxHP", CStr(UserList(userindex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", CStr(UserList(userindex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "Reputacione", CLng(UserList(userindex).Stats.Reputacione))
Call WriteVar(UserFile, "STATS", "FIT", CStr(UserList(userindex).Stats.FIT))
Call WriteVar(UserFile, "STATS", "MaxSTA", CStr(UserList(userindex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", CStr(UserList(userindex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", CStr(UserList(userindex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", CStr(UserList(userindex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", CStr(UserList(userindex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", CStr(UserList(userindex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", CStr(UserList(userindex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", CStr(UserList(userindex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", CStr(UserList(userindex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", CStr(UserList(userindex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", CStr(UserList(userindex).Stats.SkillPts))
  
Call WriteVar(UserFile, "STATS", "EXP", CStr(UserList(userindex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", CStr(UserList(userindex).Stats.ELV))

Call WriteVar(UserFile, "COFREDIOS", "Cant", val(UserList(userindex).CofreDios.Cant))
For loopC = 1 To 4
    Call WriteVar(UserFile, "COFREDIOS", "Item" & loopC, val(UserList(userindex).CofreDios.Item(loopC)))
Next loopC


Dim NCStr As String
For loopC = 1 To 30
    Call WriteVar(UserFile, "CORREO", "CORREONUM" & loopC, UserList(userindex).flags.Correo(loopC))
    NCStr = NCStr & loopC & "-" & UserList(userindex).flags.NueCorreos(loopC) & ","
Next loopC
    
    Call WriteVar(UserFile, "CORREO", "NUMCORREOS", UserList(userindex).flags.NumCorreos)
    Call WriteVar(UserFile, "CORREO", "NUECORREOS", NCStr)



Call WriteVar(UserFile, "STATS", "ELU", CStr(UserList(userindex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", CStr(UserList(userindex).Stats.UsuariosMatados))
Call WriteVar(UserFile, "MUERTES", "CrimMuertes", CStr(UserList(userindex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", CStr(UserList(userindex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "CantidadItems", val(UserList(userindex).BancoInvent.NroItems))
Dim LoopD As Integer
For LoopD = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "Obj" & LoopD, UserList(userindex).BancoInvent.Object(LoopD).ObjIndex & "-" & UserList(userindex).BancoInvent.Object(LoopD).Amount)
Next LoopD
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(userindex).Invent.NroItems))

For loopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & loopC, UserList(userindex).Invent.Object(loopC).ObjIndex & "-" & UserList(userindex).Invent.Object(loopC).Amount & "-" & UserList(userindex).Invent.Object(loopC).Equipped)
Next

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", str(UserList(userindex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", str(UserList(userindex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", str(UserList(userindex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", str(UserList(userindex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", str(UserList(userindex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", str(UserList(userindex).Invent.MunicionEqpSlot))
Call WriteVar(UserFile, "Inventory", "HerramientaSlot", str(UserList(userindex).Invent.HerramientaEqpSlot))

Dim cad As String

For loopC = 1 To MAXUSERHECHIZOS
    cad = UserList(userindex).Stats.UserHechizos(loopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & loopC, cad)
Next

Dim NroMascotas As Long
NroMascotas = UserList(userindex).NroMacotas

For loopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(userindex).MascotasIndex(loopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(userindex).MascotasIndex(loopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(userindex).MascotasType(loopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & loopC, cad)
    End If

Next


For loopC = 1 To 20
    Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "AMIGOS", "A" & loopC, UserList(userindex).flags.NombreAmigo(loopC))
    Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "AMIGOS", "AD" & loopC, UserList(userindex).flags.AmigoDesadmitido(loopC))
Next loopC

Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "" & UserList(userindex).Accounted & "", "BANCO", UserList(userindex).Stats.Banco)
Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", str(NroMascotas))

'Devuelve el head de muerto
If UserList(userindex).flags.Muerto = 1 Then
 If UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Or UserList(userindex).StatusMith.EsStatus = 5 Then
    UserList(userindex).Char.Head = iCabezaMuertoA
 ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Or UserList(userindex).StatusMith.EsStatus = 6 Then
    UserList(userindex).Char.Head = iCabezaMuertoH
 Else
    UserList(userindex).Char.Head = iCabezaMuertoN
 End If
End If

Call WriteVar(UserFile, "GUILD", "GUILDINDEX", UserList(userindex).GuildIndex)

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub
Sub SaveUserOpcional(ByVal userindex As Integer, ByVal UserFile As String)
On Error GoTo Errhandler

Dim OldUserHead As Long
Dim user_file As clsIniReader

    Set user_file = New clsIniReader
    
    '@ load file
    user_file.Initialize UserFile


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(userindex).clase = "" Or UserList(userindex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(userindex).Name)
    Exit Sub
End If


If UserList(userindex).flags.Mimetizado = 1 Then
    UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
End If



If FileExist(UserFile, vbNormal) Then
       If UserList(userindex).flags.Muerto = 1 Then
        OldUserHead = UserList(userindex).Char.Head
        UserList(userindex).Char.Head = CStr(GetVar(UserFile, "INIT", "Head"))
       End If
'       Kill UserFile
End If

Dim loopC As Integer

user_file.ChangeValue "FLAGS", "MuereQuest1", CStr(UserList(userindex).flags.MuereQuest(1))
user_file.ChangeValue "FLAGS", "MuereQuest2", CStr(UserList(userindex).flags.MuereQuest(2))
user_file.ChangeValue "FLAGS", "MuereQuest3", CStr(UserList(userindex).flags.MuereQuest(3))

user_file.ChangeValue "FLAGS", "Questeando", CStr(UserList(userindex).flags.Questeando)
user_file.ChangeValue "FLAGS", "UserNumQuest", CStr(UserList(userindex).flags.UserNumQuest)
user_file.ChangeValue "FLAGS", "GuerrasGanadas", CStr(UserList(userindex).flags.GuerrasGanadas)
user_file.ChangeValue "FLAGS", "GuerrasPerdidas", CStr(UserList(userindex).flags.GuerrasPerdidas)
user_file.ChangeValue "FLAGS", "MVPMatados", CStr(UserList(userindex).flags.MVPMatados)
user_file.ChangeValue "FLAGS", "QuestCompletadas", CStr(UserList(userindex).flags.QuestCompletadas)
user_file.ChangeValue "FLAGS", "Muerto", CStr(UserList(userindex).flags.Muerto)
user_file.ChangeValue "FLAGS", "CvcsGanados", UserList(userindex).flags.CvcsGanados
user_file.ChangeValue "FLAGS", "AlmasContenidas", CStr(UserList(userindex).flags.AlmasContenidas)
user_file.ChangeValue "FLAGS", "AlmasOfrecidas", CStr(UserList(userindex).flags.AlmasOfrecidas)
user_file.ChangeValue "FLAGS", "STOP", CStr(UserList(userindex).flags.Stopped)
user_file.ChangeValue "FLAGS", "MsjPrivado", CStr(UserList(userindex).flags.DeseoRecibirMSJ)
user_file.ChangeValue "FLAGS", "Emoticons", CStr(UserList(userindex).flags.Emoticons)
user_file.ChangeValue "FLAGS", "Llegolvlmax", CStr(UserList(userindex).flags.Llegolvlmax)
user_file.ChangeValue "FLAGS", "Llegolvl50", CStr(UserList(userindex).flags.llegolvl50)
user_file.ChangeValue "FLAGS", "Escondido", CStr(UserList(userindex).flags.Escondido)
user_file.ChangeValue "FLAGS", "Estado", CStr(UserList(userindex).flags.estado)
user_file.ChangeValue "FLAGS", "EsNoble", CStr(UserList(userindex).flags.EsNoble)
user_file.ChangeValue "PREMIUM", "EsPremium", CStr(UserList(userindex).flags.EsPremium)
user_file.ChangeValue "PREMIUM", "Vencimiento", CStr(UserList(userindex).flags.VencePremium)
user_file.ChangeValue "FLAGS", "CaballerodelDragon", CStr(UserList(userindex).flags.CaballerodelDragon)
user_file.ChangeValue "FLAGS", "Hambre", CStr(UserList(userindex).flags.Hambre)
user_file.ChangeValue "FLAGS", "Sed", CStr(UserList(userindex).flags.Sed)
user_file.ChangeValue "FLAGS", "Desnudo", CStr(UserList(userindex).flags.Desnudo)
user_file.ChangeValue "FLAGS", "Pareja", CStr(UserList(userindex).flags.Pareja)
user_file.ChangeValue "FLAGS", "Ban", CStr(UserList(userindex).flags.Ban)
user_file.ChangeValue "FLAGS", "Navegando", CStr(UserList(userindex).flags.Navegando)
user_file.ChangeValue "FLAGS", "Montando", CStr(UserList(userindex).flags.Montando)
user_file.ChangeValue "FLAGS", "PJerarquia", CStr(UserList(userindex).flags.PJerarquia)
user_file.ChangeValue "FLAGS", "SJerarquia", CStr(UserList(userindex).flags.SJerarquia)
user_file.ChangeValue "FLAGS", "TJerarquia", CStr(UserList(userindex).flags.TJerarquia)
user_file.ChangeValue "FLAGS", "CJerarquia", CStr(UserList(userindex).flags.CJerarquia)
user_file.ChangeValue "FLAGS", "CJerarquiaC", CStr(UserList(userindex).flags.CJerarquiaC)
user_file.ChangeValue "FLAGS", "Transformado", CStr(UserList(userindex).flags.Transformado)
user_file.ChangeValue "FLAGS", "SirvienteDeDios", UserList(userindex).flags.SirvienteDeDios
user_file.ChangeValue "FLAGS", "JerarquiaDios", CStr(UserList(userindex).flags.JerarquiaDios)

user_file.ChangeValue "FLAGS", "PuedeRetirarObj", CStr(UserList(userindex).flags.PuedeRetirarObj)
user_file.ChangeValue "FLAGS", "PuedeRetirarOro", CStr(UserList(userindex).flags.PuedeRetirarOro)

user_file.ChangeValue "FLAGS", "Envenenado", CStr(UserList(userindex).flags.Envenenado)
user_file.ChangeValue "FLAGS", "Paralizado", CStr(UserList(userindex).flags.Paralizado)

user_file.ChangeValue "CONSEJO", "PERTENECE", CStr(UserList(userindex).ConsejoInfo.PertAlCons)
user_file.ChangeValue "CONSEJO", "LIDERCONSEJO", CStr(UserList(userindex).ConsejoInfo.LiderConsejo)
user_file.ChangeValue "CONSEJO", "PERTENECECAOS", CStr(UserList(userindex).ConsejoInfo.PertAlConsCaos)
user_file.ChangeValue "CONSEJO", "LIDERCONSEJOCAOS", CStr(UserList(userindex).ConsejoInfo.LiderConsejoCaos)


user_file.ChangeValue "COUNTERS", "Pena", CStr(UserList(userindex).Counters.Pena)

user_file.ChangeValue "FACCIONES", "EjercitoReal", CStr(UserList(userindex).Faccion.ArmadaReal)
user_file.ChangeValue "FACCIONES", "EjercitoCaos", CStr(UserList(userindex).Faccion.FuerzasCaos)
user_file.ChangeValue "FACCIONES", "CiudMatados", CStr(UserList(userindex).Faccion.CiudadanosMatados)
user_file.ChangeValue "FACCIONES", "CrimMatados", CStr(UserList(userindex).Faccion.CriminalesMatados)
user_file.ChangeValue "FACCIONES", "NeutrMatados", CStr(UserList(userindex).Faccion.NeutralesMatados)
user_file.ChangeValue "FACCIONES", "rArCaos", CStr(UserList(userindex).Faccion.RecibioArmaduraCaos)
user_file.ChangeValue "FACCIONES", "rArReal", CStr(UserList(userindex).Faccion.RecibioArmaduraReal)
user_file.ChangeValue "FACCIONES", "rExCaos", CStr(UserList(userindex).Faccion.RecibioExpInicialCaos)
user_file.ChangeValue "FACCIONES", "rExReal", CStr(UserList(userindex).Faccion.RecibioExpInicialReal)
user_file.ChangeValue "FACCIONES", "recCaos", CStr(UserList(userindex).Faccion.RecompensasCaos)
user_file.ChangeValue "FACCIONES", "recReal", CStr(UserList(userindex).Faccion.RecompensasReal)
user_file.ChangeValue "FACCIONES", "Reenlistadas", CStr(UserList(userindex).Faccion.Reenlistadas)

'¿Fueron modificados los atributos del usuario?
If Not UserList(userindex).flags.TomoPocion Then
    For loopC = 1 To UBound(UserList(userindex).Stats.UserAtributos)
        user_file.ChangeValue "ATRIBUTOS", "AT" & loopC, CStr(UserList(userindex).Stats.UserAtributos(loopC))
    Next
Else
    For loopC = 1 To UBound(UserList(userindex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        user_file.ChangeValue "ATRIBUTOS", "AT" & loopC, CStr(UserList(userindex).Stats.UserAtributosBackUP(loopC))
    Next
End If

For loopC = 1 To UBound(UserList(userindex).Stats.UserSkills)
    user_file.ChangeValue "SKILLS", "SK" & loopC, CStr(UserList(userindex).Stats.UserSkills(loopC))
Next


user_file.ChangeValue "CONTACTO", "Email", UserList(userindex).email

user_file.ChangeValue "MISIONES", "Dia", UserList(userindex).Misiones.Dia
user_file.ChangeValue "MISIONES", "NumeroMision", UserList(userindex).Misiones.NumeroMision
user_file.ChangeValue "MISIONES", "ConteoUser", UserList(userindex).Misiones.ConteoUser
user_file.ChangeValue "MISIONES", "Completada", UserList(userindex).Misiones.Completada

user_file.ChangeValue "INIT", "Bon1", UserList(userindex).Bon1
user_file.ChangeValue "INIT", "Bon2", UserList(userindex).Bon2
user_file.ChangeValue "INIT", "Bon3", UserList(userindex).Bon3

user_file.ChangeValue "INIT", "UltimoLogeo", UserList(userindex).UltimoLogeo
user_file.ChangeValue "INIT", "UltimaDenuncia", UserList(userindex).UltimaDenuncia
user_file.ChangeValue "INIT", "PrimeraDenuncia", UserList(userindex).PrimeraDenuncia

user_file.ChangeValue "INIT", "Genero", UserList(userindex).Genero
user_file.ChangeValue "INIT", "Raza", UserList(userindex).Raza
user_file.ChangeValue "INIT", "Hogar", UserList(userindex).Hogar
user_file.ChangeValue "INIT", "Clase", UserList(userindex).clase
user_file.ChangeValue "INIT", "Desc", UserList(userindex).Desc
user_file.ChangeValue "INIT", "NickMascota", UserList(userindex).NickMascota
user_file.ChangeValue "INIT", "Password", UserList(userindex).Password

user_file.ChangeValue "INIT", "Heading", CStr(UserList(userindex).Char.Heading)
user_file.ChangeValue "INIT", "Head", CStr(UserList(userindex).OrigChar.Head)

If UserList(userindex).flags.Muerto = 0 Then
    user_file.ChangeValue "INIT", "Body", CStr(UserList(userindex).Char.Body)
End If

user_file.ChangeValue "INIT", "Arma", CStr(UserList(userindex).Char.WeaponAnim)
user_file.ChangeValue "INIT", "Escudo", CStr(UserList(userindex).Char.ShieldAnim)
user_file.ChangeValue "INIT", "Casco", CStr(UserList(userindex).Char.CascoAnim)

user_file.ChangeValue "INIT", "LastIP", UserList(userindex).ip
user_file.ChangeValue "INIT", "Position", UserList(userindex).Pos.Map & "-" & UserList(userindex).Pos.X & "-" & UserList(userindex).Pos.Y
user_file.ChangeValue "INIT", "LastHD", UserList(userindex).hd
user_file.ChangeValue "CHAR", "Cuenta", UserList(userindex).Accounted
user_file.ChangeValue "CHAR", "InventorySlots", UserList(userindex).InventorySlots


user_file.ChangeValue "STATS", "GLD", CStr(UserList(userindex).Stats.GLD)
user_file.ChangeValue "STATS", "PuntosTorneo", CStr(UserList(userindex).Stats.PuntosTorneo)
user_file.ChangeValue "STATS", "PuntosDonacion", CStr(UserList(userindex).Stats.PuntosDonacion)
user_file.ChangeValue "STATS", "TSPoints", CStr(UserList(userindex).Stats.TSPoints)
user_file.ChangeValue "STATS", "ParejasGanadas", CStr(UserList(userindex).Stats.ParejasGanadas)
user_file.ChangeValue "STATS", "ParejasPerdidas", CStr(UserList(userindex).Stats.ParejasPerdidas)
user_file.ChangeValue "STATS", "MuertesUser", CStr(UserList(userindex).Stats.MuertesUser)
user_file.ChangeValue "STATS", "DuelosGanados", CStr(UserList(userindex).Stats.DuelosGanados)
user_file.ChangeValue "STATS", "DuelosPerdidos", CStr(UserList(userindex).Stats.DuelosPerdidos)
user_file.ChangeValue "STATS", "TorneosParticipados", CStr(UserList(userindex).Stats.TorneosParticipados)
user_file.ChangeValue "STATS", "TrofOro", CStr(UserList(userindex).Stats.TrofOro)
user_file.ChangeValue "STATS", "TrofPlata", CStr(UserList(userindex).Stats.TrofPlata)
user_file.ChangeValue "STATS", "TrofBronce", CStr(UserList(userindex).Stats.TrofBronce)
user_file.ChangeValue "STATS", "MedOro", CStr(UserList(userindex).Stats.MedOro)
user_file.ChangeValue "STATUS", "EsStatus", CStr(UserList(userindex).StatusMith.EsStatus)
user_file.ChangeValue "STATUS", "Eligio", CStr(UserList(userindex).StatusMith.EligioStatus)

user_file.ChangeValue "STATS", "BANCO", UserList(userindex).Stats.Banco

user_file.ChangeValue "STATS", "MET", CStr(UserList(userindex).Stats.MET)
user_file.ChangeValue "STATS", "MaxHP", CStr(UserList(userindex).Stats.MaxHP)
user_file.ChangeValue "STATS", "MinHP", CStr(UserList(userindex).Stats.MinHP)

user_file.ChangeValue "STATS", "Reputacione", CLng(UserList(userindex).Stats.Reputacione)
user_file.ChangeValue "STATS", "FIT", CStr(UserList(userindex).Stats.FIT)
user_file.ChangeValue "STATS", "MaxSTA", CStr(UserList(userindex).Stats.MaxSta)
user_file.ChangeValue "STATS", "MinSTA", CStr(UserList(userindex).Stats.MinSta)

user_file.ChangeValue "STATS", "MaxMAN", CStr(UserList(userindex).Stats.MaxMAN)
user_file.ChangeValue "STATS", "MinMAN", CStr(UserList(userindex).Stats.MinMAN)

user_file.ChangeValue "STATS", "MaxHIT", CStr(UserList(userindex).Stats.MaxHIT)
user_file.ChangeValue "STATS", "MinHIT", CStr(UserList(userindex).Stats.MinHIT)

user_file.ChangeValue "STATS", "MaxAGU", CStr(UserList(userindex).Stats.MaxAGU)
user_file.ChangeValue "STATS", "MinAGU", CStr(UserList(userindex).Stats.MinAGU)

user_file.ChangeValue "STATS", "MaxHAM", CStr(UserList(userindex).Stats.MaxHam)
user_file.ChangeValue "STATS", "MinHAM", CStr(UserList(userindex).Stats.MinHam)

user_file.ChangeValue "STATS", "SkillPtsLibres", CStr(UserList(userindex).Stats.SkillPts)
  
user_file.ChangeValue "STATS", "EXP", CStr(UserList(userindex).Stats.Exp)
user_file.ChangeValue "STATS", "ELV", CStr(UserList(userindex).Stats.ELV)

user_file.ChangeValue "COFREDIOS", "Cant", val(UserList(userindex).CofreDios.Cant)
For loopC = 1 To 4
    user_file.ChangeValue "COFREDIOS", "Item" & loopC, val(UserList(userindex).CofreDios.Item(loopC))
Next loopC


Dim NCStr As String
For loopC = 1 To 30
    user_file.ChangeValue "CORREO", "CORREONUM" & loopC, UserList(userindex).flags.Correo(loopC)
    NCStr = NCStr & loopC & "-" & UserList(userindex).flags.NueCorreos(loopC) & ","
Next loopC
    
    user_file.ChangeValue "CORREO", "NUMCORREOS", UserList(userindex).flags.NumCorreos
    user_file.ChangeValue "CORREO", "NUECORREOS", NCStr



user_file.ChangeValue "STATS", "ELU", CStr(UserList(userindex).Stats.ELU)
user_file.ChangeValue "MUERTES", "UserMuertes", CStr(UserList(userindex).Stats.UsuariosMatados)
user_file.ChangeValue "MUERTES", "CrimMuertes", CStr(UserList(userindex).Stats.CriminalesMatados)
user_file.ChangeValue "MUERTES", "NpcsMuertes", CStr(UserList(userindex).Stats.NPCsMuertos)
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "CantidadItems", val(UserList(userindex).BancoInvent.NroItems))
Dim LoopD As Integer
For LoopD = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "BancoInventory", "Obj" & LoopD, UserList(userindex).BancoInvent.Object(LoopD).ObjIndex & "-" & UserList(userindex).BancoInvent.Object(LoopD).Amount)
Next LoopD
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
user_file.ChangeValue "Inventory", "CantidadItems", val(UserList(userindex).Invent.NroItems)

For loopC = 1 To MAX_INVENTORY_SLOTS
    user_file.ChangeValue "Inventory", "Obj" & loopC, UserList(userindex).Invent.Object(loopC).ObjIndex & "-" & UserList(userindex).Invent.Object(loopC).Amount & "-" & UserList(userindex).Invent.Object(loopC).Equipped
Next

user_file.ChangeValue "Inventory", "WeaponEqpSlot", str(UserList(userindex).Invent.WeaponEqpSlot)
user_file.ChangeValue "Inventory", "ArmourEqpSlot", str(UserList(userindex).Invent.ArmourEqpSlot)
user_file.ChangeValue "Inventory", "CascoEqpSlot", str(UserList(userindex).Invent.CascoEqpSlot)
user_file.ChangeValue "Inventory", "EscudoEqpSlot", str(UserList(userindex).Invent.EscudoEqpSlot)
user_file.ChangeValue "Inventory", "BarcoSlot", str(UserList(userindex).Invent.BarcoSlot)
user_file.ChangeValue "Inventory", "MunicionSlot", str(UserList(userindex).Invent.MunicionEqpSlot)
user_file.ChangeValue "Inventory", "HerramientaSlot", str(UserList(userindex).Invent.HerramientaEqpSlot)

Dim cad As String

For loopC = 1 To MAXUSERHECHIZOS
    cad = UserList(userindex).Stats.UserHechizos(loopC)
    user_file.ChangeValue "HECHIZOS", "H" & loopC, cad
Next

Dim NroMascotas As Long
NroMascotas = UserList(userindex).NroMacotas

For loopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(userindex).MascotasIndex(loopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(userindex).MascotasIndex(loopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(userindex).MascotasType(loopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        user_file.ChangeValue "MASCOTAS", "MAS" & loopC, cad
    End If

Next


For loopC = 1 To 20
    Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "AMIGOS", "A" & loopC, UserList(userindex).flags.NombreAmigo(loopC))
    Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "AMIGOS", "AD" & loopC, UserList(userindex).flags.AmigoDesadmitido(loopC))
Next loopC

Call WriteVar(App.Path & "\Accounts\" & UserList(userindex).Accounted & ".act", "" & UserList(userindex).Accounted & "", "BANCO", UserList(userindex).Stats.Banco)
user_file.ChangeValue "MASCOTAS", "NroMascotas", str(NroMascotas)

'Devuelve el head de muerto
If UserList(userindex).flags.Muerto = 1 Then
 If UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Or UserList(userindex).StatusMith.EsStatus = 5 Then
    UserList(userindex).Char.Head = iCabezaMuertoA
 ElseIf UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Or UserList(userindex).StatusMith.EsStatus = 6 Then
    UserList(userindex).Char.Head = iCabezaMuertoH
 Else
    UserList(userindex).Char.Head = iCabezaMuertoN
 End If
End If

user_file.ChangeValue "GUILD", "GUILDINDEX", UserList(userindex).GuildIndex

user_file.DumpFile UserFile

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub

'Newbie - O no eligio
Function Neutral(ByVal userindex As Integer) As Boolean
Neutral = UserList(userindex).StatusMith.EsStatus = 0
End Function
'Ciudadano
Function Ciudadano(ByVal userindex As Integer) As Boolean
Ciudadano = UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3 Or UserList(userindex).StatusMith.EsStatus = 5
End Function
'Criminal
Function Criminal(ByVal userindex As Integer) As Boolean
Criminal = UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4 Or UserList(userindex).StatusMith.EsStatus = 6
End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim loopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

If NpcNumero > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "MVP", Npclist(NpcIndex).MVP)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GivePTS", val(Npclist(NpcIndex).GivePTS))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLDMin", val(Npclist(NpcIndex).GiveGLDMin))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLDMax", val(Npclist(NpcIndex).GiveGLDMax))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))

'Cristales
Call WriteVar(npcfile, "NPC" & NpcNumero, "Cristales", val(Npclist(NpcIndex).Cristales))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesPequesMin", val(Npclist(NpcIndex).CristalesPequesMin))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesPequesMax", val(Npclist(NpcIndex).CristalesPequesMax))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesMedianosMin", val(Npclist(NpcIndex).CristalesMedianosMin))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesMedianosMax", val(Npclist(NpcIndex).CristalesMedianosMax))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesGrandesMin", val(Npclist(NpcIndex).CristalesGrandesMin))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesGrandesMax", val(Npclist(NpcIndex).CristalesGrandesMax))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesEpicosMin", val(Npclist(NpcIndex).CristalesEpicosMin))
Call WriteVar(npcfile, "NPC" & NpcNumero, "CristalesEpicosMax", val(Npclist(NpcIndex).CristalesEpicosMax))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For loopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & loopC, Npclist(NpcIndex).Invent.Object(loopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(loopC).Amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

Dim npcfile As String

If NpcNumber > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).MVP = GetVar(npcfile, "NPC" & NpcNumber, "MVP")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Char.ShieldAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "EscudoAnim"))
Npclist(NpcIndex).Char.WeaponAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "ArmaAnim"))
Npclist(NpcIndex).Char.CascoAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "CascoAnim"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
Npclist(NpcIndex).GivePTS = val(GetVar(npcfile, "NPC" & NpcNumber, "GivePTS"))
Npclist(NpcIndex).GiveGLDMin = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLDMin"))
Npclist(NpcIndex).GiveGLDMax = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLDMax"))

'Crsitales
Npclist(NpcIndex).Cristales = val(GetVar(npcfile, "NPC" & NpcNumber, "Cristales"))
Npclist(NpcIndex).CristalesPequesMin = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesPequesMin"))
Npclist(NpcIndex).CristalesPequesMax = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesPequesMax"))
Npclist(NpcIndex).CristalesMedianosMin = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesMedianosMin"))
Npclist(NpcIndex).CristalesMedianosMax = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesMedianosMax"))
Npclist(NpcIndex).CristalesGrandesMin = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesGrandesMin"))
Npclist(NpcIndex).CristalesGrandesMax = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesGrandesMax"))
Npclist(NpcIndex).CristalesEpicosMin = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesEpicosMin"))
Npclist(NpcIndex).CristalesEpicosMax = val(GetVar(npcfile, "NPC" & NpcNumber, "CristalesEpicosMax"))



Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))


Dim loopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For loopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & loopC)
        Npclist(NpcIndex).Invent.Object(loopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(loopC).Amount = val(ReadField(2, ln, 45))
       
    Next loopC
Else
    For loopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(loopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(loopC).Amount = 0
    Next loopC
End If

Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))


Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal userindex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(userindex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal userindex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(userindex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
Public Sub DarPremioCastillos()
On Error GoTo handler
Dim loopC As Integer
    For loopC = 1 To LastUser
        If UserList(loopC).GuildIndex <> 0 Then
            If Guilds(UserList(loopC).GuildIndex).GuildName = CastilloNorte Then
                UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + 200000
                Call SendData(SendTarget.toindex, (loopC), 0, "||667@200.000@33")
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            If Guilds(UserList(loopC).GuildIndex).GuildName = CastilloSur Then
                UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + 200000
                Call SendData(SendTarget.toindex, (loopC), 0, "||667@200.000@31")
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            If Guilds(UserList(loopC).GuildIndex).GuildName = CastilloEste Then
                UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + 200000
                Call SendData(SendTarget.toindex, (loopC), 0, "||667@200.000@34")
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            If Guilds(UserList(loopC).GuildIndex).GuildName = CastilloOeste Then
                UserList(loopC).Stats.GLD = UserList(loopC).Stats.GLD + 200000
                Call SendData(SendTarget.toindex, (loopC), 0, "||667@200.000@32")
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            
            If Guilds(UserList(loopC).GuildIndex).GuildName = CastilloNorte And Guilds(UserList(loopC).GuildIndex).GuildName = CastilloSur And Guilds(UserList(loopC).GuildIndex).GuildName = CastilloEste And Guilds(UserList(loopC).GuildIndex).GuildName = CastilloOeste Then
                Call AgregarPuntos(loopC, 20)
                Call SendData(SendTarget.toindex, (loopC), 0, "||668")
                Call SendData(SendTarget.toindex, (loopC), 0, "||57@20")
            End If
            
            If UCase$(Guilds(UserList(loopC).GuildIndex).GuildName) = UCase(Fortaleza) Then
                Call AgregarPuntos(loopC, 20)
                Call SendData(SendTarget.toindex, (loopC), 0, "||57@20")
                Call SendData(SendTarget.toindex, (loopC), 0, "||669")
            End If
            
        End If
    Next loopC
    
Call SendData(SendTarget.ToAll, 0, 0, "||670")
    
Exit Sub
handler:
Call LogError("Error en DarPremioCastillos.")
End Sub
