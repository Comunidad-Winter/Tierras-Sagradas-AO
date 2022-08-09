Attribute VB_Name = "modLogs"
Option Explicit

Dim i As Long
Dim LogVacio As Integer
Public Sub LogDrops(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Dropeos")
    'Logs.Dropeos = Logs.Dropeos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogNobleza(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Nobleza")
    'Logs.Nobleza = Logs.Nobleza & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogDuelos(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Duelos")
    'Logs.Duelos = Logs.Duelos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogDarOro(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\DarOro")
   ' Logs.DarOro = Logs.DarOro & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogDesafios(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Desafios")
   ' Logs.Desafios = Logs.Desafios & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogTransferencias(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Transferencias")
   ' Logs.Transferencias = Logs.Transferencias & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogTorneos(Texto As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\CrearTorneos.log" For Append Shared As #nfile
    Print #nfile, "" & Date & " " & Time & " " & Texto & ""
Close #nfile

Exit Sub
Errhandler:
   ' Logs.AgarrarItems = Logs.AgarrarItems & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogAgarrarItems(Texto As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\AgarraItems.log" For Append Shared As #nfile
    Print #nfile, "" & Date & " " & Time & " " & Texto & ""
Close #nfile

Exit Sub
Errhandler:
   ' Logs.AgarrarItems = Logs.AgarrarItems & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogPassw(Texto As String)
Call GuardarLogs("" & Texto & "", "\Turbios\Passwords")
End Sub
Public Sub LogAlmas(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Almas")
   ' Logs.Almas = Logs.Almas & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogCorreos(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\CorreosEnviados")
   ' Logs.EnviarCorreos = Logs.EnviarCorreos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogRCorreos(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\CorreosRetirados")
   ' Logs.RetirarCorreos = Logs.RetirarCorreos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogTirarItems(Texto As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\TirarItems.log" For Append Shared As #nfile
    Print #nfile, "" & Date & " " & Time & " " & Texto & ""
Close #nfile

Exit Sub
Errhandler:
   ' Logs.TirarItems = Logs.TirarItems & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogComercios(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Comercios")
   ' Logs.Comercios = Logs.Comercios & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogDepositos(Texto As String)

On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\Turbios\Depositos.log" For Append Shared As #nfile
    Print #nfile, "" & Date & " " & Time & " " & Texto & ""
Close #nfile

Exit Sub

Errhandler:
   ' Logs.Depositos = Logs.Depositos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogCanjeos(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Canjeos")
   ' Logs.Canjeos = Logs.Canjeos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogMedallas(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Medallas")
   ' Logs.Medallas = Logs.Medallas & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogAsesinato(Texto As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\Turbios\Asesinatos")
   ' Logs.Asesinatos = Logs.Asesinatos & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub logVentaCasa(ByVal Texto As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\VentaCasas")
   ' Logs.VentaCasas = Logs.VentaCasas & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogHackAttemp(Texto As String)
   ' Call GuardarLogs("" & Date & " " & Time & " " & Texto & "", "\HackAttemp")
   ' Logs.HackAttemp = Logs.HackAttemp & Date & " " & Time & " " & Texto & vbCrLf
End Sub
Public Sub LogCriticEvent(Desc As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & Desc & "", "CriticEvent")
   ' Logs.CriticEvent = Logs.CriticEvent & Date & " " & Time & " " & Desc & vbCrLf
End Sub
Public Sub LogError(Desc As String)
    Call GuardarLogs("" & Date & " " & Time & " " & Desc & "", "\Errores")
   ' Logs.Errores = Logs.Errores & Date & " " & Time & " " & Desc & vbCrLf
End Sub
Public Sub LogTarea(Desc As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & Desc & "", "haciendo")
   ' Logs.Tarea = Logs.Tarea & Date & " " & Time & " " & Desc & vbCrLf
End Sub
Public Sub LogDesarrollo(ByVal str As String)
    'Call GuardarLogs("" & Date & " " & Time & " " & str & "", "Desarrollo")
   ' Logs.Desarrollo = Logs.Desarrollo & Date & " " & Time & " " & str & vbCrLf
End Sub
Public Sub LogGMss(Nombre As String, Texto As String, Consejero As Boolean)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\WorldBackUp\" & Date & ".log" For Append Shared As #nfile
    Print #nfile, "" & Date & " " & Time & " " & Nombre & " - " & Texto & ""
Close #nfile

Exit Sub

Errhandler:

End Sub
Public Sub LogGM(Nombre As String, Texto As String, Consejero As Boolean)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Texto
Close #nfile

Exit Sub

Errhandler:

End Sub
Public Sub GuardarLogs(Texto As String, ArchivoTextual As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\" & ArchivoTextual & ".log" For Append Shared As #nfile
    Print #nfile, Texto
Close #nfile

Exit Sub

Errhandler:

End Sub
