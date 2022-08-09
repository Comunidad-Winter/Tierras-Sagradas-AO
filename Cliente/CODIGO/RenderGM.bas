Attribute VB_Name = "RenderGM"
Option Explicit
 
' @ Designed by maTih.-
 
'Cantidad máxima de mensajes.
Const MAX_MSG           As Byte = 7
Const INVALID_SLOT      As Byte = 0
 
Type RMSGRender
     color                      As Long         'Que color lleva.
     text                       As String       'Texto..
     GMName                     As String       'Gm quien envia msg.
     Duracion                   As Long         'Lo que dura en pantalla.
     GTickStart                 As Long         'GetTickCount de cuando inicio.
     MSGActivated               As Boolean      'Para saber si hay mensaje
End Type
 
Public RenderMSG(1 To MAX_MSG) As RMSGRender
Private LastRMSG               As Byte
 
Function GetGMName(ByRef sText As String) As String
 
' @ Devuelve el nick del gm.
 
Dim inPos       As Integer
 
inPos = InStr(1, sText, ",")
 
If Not inPos <> 0 Then Exit Function
 
GetGMName = left$(sText, inPos - 1)
 
End Function
 
Function GetText(ByRef sText As String) As String
 
' @ Devuelve el texto quitando el nick del gm.
 
Dim inPos   As Integer
 
inPos = InStr(1, sText, ",")
 
If Not inPos <> 0 Then Exit Function
 
GetText = mid$(sText, inPos + 1)
 
End Function
 
Function GetDuracion(ByRef sText As String) As Long
 
' @ Devuelve la duracion en el render del texto.
 
GetDuracion = 5000 + (100 * (Len(sText)))
 
End Function
 
Function GetColor(ByRef GMName As String) As Long
 
' @ Devuelve el color del texto según el nick.
 
Dim gIndexChar  As Integer
 
gIndexChar = GetIndexChar(GMName)
 
If Not gIndexChar <> 0 Then Exit Function
 
With ColoresPJ(charlist(gIndexChar).priv)
 
    GetColor = D3DColorARGB(255, .r, .g, .b)
 
End With
 
End Function
 
Function GetIndexChar(ByRef GMName As String) As Integer
 
' @ Devuelve el CharIndex de GMName.
 
Dim loopX   As Long
Dim nowChar As String
 
For loopX = 1 To 10000    'MAX_CHARS
 
    nowChar = UCase$(charlist(loopX).Nombre)
   
    If nowChar = UCase$(GMName) Then Exit For
   
Next loopX
 
GetIndexChar = loopX
 
End Function
 
Function GetNextSlot() As Byte
 
' @ Busca un slot.
 
Dim loopX   As Long
 
For loopX = 1 To MAX_MSG
    If Not RenderMSG(loopX).MSGActivated Then
           GetNextSlot = CByte(loopX)
           Exit Function
    End If
Next loopX
 
GetNextSlot = INVALID_SLOT
 
End Function
 
Function CompactArray() As Boolean
 
' @ Corre los mensajes y deja el 1 libre.
 
On Error GoTo errHandler
 
Dim tmpDatas(1 To MAX_MSG) As RMSGRender
Dim loopX                  As Long
 
For loopX = 1 To MAX_MSG
    tmpDatas(loopX) = RenderMSG(loopX)
Next loopX
 
For loopX = 1 To 4
    tmpDatas(loopX + 1) = tmpDatas(loopX)
Next loopX
 
ClearMessage 1
 
CompactArray = (RenderMSG(1).GMName = vbNullString)
 
Exit Function
 
errHandler:
 
CompactArray = False
 
End Function
 
Function HayRmsg() As Boolean
 
' @ Devuelve si hay un mensaje para renderizar
 
     HayRmsg = (LastRMSG <> 0)
 
End Function
 
Sub RenderMessage()
 
' @ Renderiza los mensajes
 
Dim loopX   As Long
 
For loopX = 1 To MAX_MSG
   
    With RenderMSG(loopX)
         
         'Si hay mensaje.
         If .MSGActivated Then
            'Si todavia no llegó a su fin.
            If Not (GetTickCount - .GTickStart) > .Duracion Then
                'Render.
                Call Texto.Engine_Text_Draw(5, 5 + (loopX * 10), .GMName & "> " & .text, .color)
            Else
                ClearMessage loopX
            End If
        End If
             
    End With
   
Next loopX
 
End Sub
 
Sub Create(ByRef Rdata As String)
 
' @ Crea un mensaje.
 
Dim inTmpSlot       As Byte
Dim compactOk       As Boolean
 
'Busco un slot para el mensaje
inTmpSlot = GetNextSlot
 
'Si no encontré un mensaje entonces corro el array.
If Not inTmpSlot <> INVALID_SLOT Then
    compactOk = CompactArray
   
    'Si todo fue bien entonces uso el primer slot.
    If compactOk Then
       inTmpSlot = 1
    End If
 
    If Not compactOk Then Exit Sub
End If
 
With RenderMSG(inTmpSlot)
    .text = GetText(Rdata)
    .Duracion = GetDuracion(RenderMSG(inTmpSlot).text)
    .GMName = GetGMName(Rdata)
    .color = GetColor(RenderMSG(inTmpSlot).GMName)
    .GTickStart = GetTickCount
    .MSGActivated = True
End With
 
LastRMSG = LastRMSG + 1
 
End Sub
 
Sub ClearMessage(ByVal RSlot As Byte)
 
' @ Borra el msg
 
With RenderMSG(RSlot)
     .color = 0
     .GMName = vbNullString
     .text = vbNullString
     .Duracion = 0
     .GTickStart = 0
     .MSGActivated = False
End With
 
If Not (LastRMSG - 1) < 0 Then LastRMSG = LastRMSG - 1
 
End Sub
