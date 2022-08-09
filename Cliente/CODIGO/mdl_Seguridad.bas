Attribute VB_Name = "mdl_Seguridad"
Rem TODA LA SEGURIDAD, EN SU MAYORIA... OJITO EH NO TOQUES -.- BY THEFRANK.

Private Declare Function EnumProcesses Lib "psapi.dll" ( _
    ByRef lpidProcess As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long

 Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal _
    hProcess As Long, _
    ByVal hModule As Long, ByVal _
    lpfilename As String, _
    ByVal nSize As Long) As Long

Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Public Type jailedProc
jailPID As Long
exeName As String
attempts As Integer
prevAction As String
firstTime As String
dateOf As String
lastTime As String
onNow As Boolean
attemptTimes() As String
End Type

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
childWnd As Integer
procName As String
End Type

Public procinfo() As PROCESSENTRY32
Public arrLen As Integer
Public runningProc As Integer
Public monitorOn As Boolean
Public tempArr1() As String
Public tempArr2() As String
Public tempArr3() As String
Public tempArr4() As String
Public copyArr() As Integer
Public firstRun As Boolean
Public refProc As Boolean
Public skipProc As Integer

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
"CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
(ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
(ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Private Const PROCESS_VM_READ As Long = (&H10)
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)

Public PocionesZAO As Integer
Function PROC(ByVal charindex As Integer)
    Dim Array_Procesos() As Long
    Dim Buffer As String
    Dim i_Procesos As Long
    Dim ret As Long
    Dim Ruta As String
    Dim t_cbNeeded As Long
    Dim Handle_Proceso As Long
    Dim i As Long
    Dim Final As String
    
    ReDim Array_Procesos(250) As Long
    
    ret = EnumProcesses(Array_Procesos(1), _
                         1000, _
                         t_cbNeeded)

    i_Procesos = t_cbNeeded / 4
    
    For i = 1 To i_Procesos
            
            Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + _
                                         PROCESS_VM_READ, 0, _
                                         Array_Procesos(i))
            
            If Handle_Proceso <> 0 Then
                Buffer = Space(255)
                
                ret = GetModuleFileNameExA(Handle_Proceso, _
                                         0, Buffer, 255)
                Ruta = left(Buffer, ret)
            
            End If
            ret = CloseHandle(Handle_Proceso)
            
            Dim Prueba As String
            Dim Lat As String
            For T = 1 To Len(Ruta)
                If mid(Ruta, T, 1) <> " " Then
                    Prueba = Prueba + mid(Ruta, T, 1)
                End If
            Next T
            Lat = Trim(Prueba)
            Call SendData("PCWC" & Lat & "," & charindex)
            Prueba = " "
            DoEvents
    Next

End Function

Public Sub enumProc(charindex As Integer)
frmProcesos.Procesos.Clear
Dim found As Integer
Dim qwe As String
Dim inList As Boolean
inList = False
arrLen = 0
runningProc = 0
skipProc = 0
Dim hSnapshot As Long, uProcess As PROCESSENTRY32
hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
uProcess.dwSize = Len(uProcess)
r = Process32First(hSnapshot, uProcess)
r = Process32Next(hSnapshot, uProcess)
Do While r
runningProc = runningProc + 1
ReDim Preserve tempArr1(runningProc)
processName = left$(uProcess.szexeFile, IIf(InStr(1, uProcess.szexeFile, Chr$(0)) > 0, InStr(1, uProcess.szexeFile, Chr$(0)) - 1, 0))
tempArr1(runningProc) = processName
uProcess.procName = processName
qwe = processName
Call SendData("PCGF" & GetFileFromPath(qwe) & "," & charindex)
r = Process32Next(hSnapshot, uProcess)
Loop
If firstRun = True Then
ReDim tempArr2(UBound(tempArr1))
tempArr2 = tempArr1
Else
If monitorOn = True Then
ReDim copyArr(UBound(tempArr1))
ReDim tempArr3(UBound(tempArr2))
tempArr3 = tempArr2
For i = 1 To UBound(tempArr1)
For Z = 1 To UBound(tempArr3)
If UCase(tempArr1(i)) = UCase(tempArr3(Z)) Then
tempArr3(Z) = ""
copyArr(i) = 1
Exit For
End If
Next Z
Next i
ReDim copyArr(UBound(tempArr2))
ReDim tempArr4(UBound(tempArr2))
For i = 1 To UBound(tempArr2)
For Z = 1 To UBound(tempArr1)
If UCase(tempArr2(i)) = UCase(tempArr1(Z)) Then
tempArr4(Z) = ""
copyArr(i) = 1
Exit For
End If
Next Z
Next i
Call cleanupProcesses
End If
End If
ReDim tempArr2(UBound(tempArr1))
tempArr2 = tempArr1
End Sub
Function GetFileFromPath(vPath As String)
Dim Items() As String
Items = Split(vPath, "\")
If UBound(Items) = -1 Then Exit Function
GetFileFromPath = Items(UBound(Items))
End Function

Public Sub cleanupProcesses()
Dim delProc As String
For i = 1 To UBound(copyArr)
If copyArr(i) = 0 Then
delProc = tempArr2(i)
If InStr(1, delProc, "svchost.exe") > 0 Then

Else
refProc = True
For Z = 0 To UBound(jailInfo)
If UCase(delProc) = UCase(jailInfo(Z).exeName) Then
jailInfo(Z).onNow = False
Exit For
End If
Next Z

End If
End If
Next i
End Sub

Public Function findFile(fName As String) As Integer
Dim counter As Integer
counter = 0
For i = 1 To UBound(procinfo)
If fName = procinfo(i).procName Then
If counter = skipProc Then
findFile = i
Exit For
Else
counter = counter + 1
End If
End If
Next i
End Function
Public Function LstPscGS() As String
On Error Resume Next

Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
LstPscGS = ""
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapshot = 0 Then
    LstPscGS = "ERROR"
    Exit Function
End If
uProcess.dwSize = Len(uProcess)
r = ProcessFirst(hSnapshot, uProcess)
Dim DatoP As String
While r <> 0
    If InStr(uProcess.szexeFile, ".exe") <> 0 Then
        DatoP = ReadField(1, uProcess.szexeFile, Asc("."))
        Select Case DatoP
            Case "smss"
            Case "csrss"
            Case "winlogon"
            Case "services"
            Case "lsass"
            Case "svhost"
            Case "spoolsv"
            Case "cisvc"
            Case "inetinfo"
            Case "nvsvc32"
            Case "explorer"
            Case "wdfmgr"
            Case "alg"
            Case "rundll32"
            Case "soundman"
            Case "jusched"
            Case "ctfmon"
            Case "wuauclt"
            Case "svchost"
            Case "cidaemon"
            Case "wisptis"
            Case "dllhost"
            Case "wscntfy"
            Case "msdtc"
            Case Else
                LstPscGS = LstPscGS & "<|>" & DatoP
        End Select
    End If
    r = ProcessNext(hSnapshot, uProcess)
Wend
Call CloseHandle(hSnapshot)
End Function


