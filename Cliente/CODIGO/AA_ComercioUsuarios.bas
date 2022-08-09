Attribute VB_Name = "AA_ComercioUsuarios"
'Nuevo sistema de comercio by GhinZuL
Option Explicit
Private Type cVal
    iIndex          As Long
    iCantidad       As Long
    iNombre         As String
    iOfrece         As Integer
    iGrhIndex       As Integer
End Type
Public rOro As Long
Public uOro As Long
Private cOferta(20) As cVal
Private cItem(20)   As cVal
Public cNombre      As String
Private cTempRead   As String
Private cTempRead2  As String
Private iCom        As Long
Private cOferto     As Boolean
Private cRecivi     As Boolean
Public Sub comIniciar(Rdata As String)
cNombre = ReadField(1, Rdata, Asc("$"))
cTempRead = ReadField(2, Rdata, Asc("$"))
    For iCom = 1 To 20
        With cItem(iCom)
            cTempRead2 = ReadField(iCom, cTempRead, Asc(","))
            .iIndex = ReadField(1, cTempRead2, Asc("-"))
            .iCantidad = ReadField(2, cTempRead2, Asc("-"))
            .iNombre = ReadField(3, cTempRead2, Asc("-"))
        End With
    Next iCom
frmNuevoComercio.conQuien.Caption = "Comerciando con: " & cNombre
comCarga
frmNuevoComercio.Show , frmMain
End Sub
Private Sub comCarga()
    With frmNuevoComercio
        .TusItems.Clear
        .Ofrecer.Clear
        .Oferta.Clear
            For iCom = 1 To 20
                If cItem(iCom).iCantidad = 0 Then
                    .TusItems.AddItem "(Nada) [0]"
                Else
                    .TusItems.AddItem cItem(iCom).iNombre & "-[" & cItem(iCom).iCantidad & "]"
                End If
                If cItem(iCom).iOfrece > 0 Then .Ofrecer.AddItem cItem(iCom).iNombre & "-[" & cItem(iCom).iOfrece & "]"
                
                If cOferta(iCom).iCantidad > 0 Then
                .Oferta.AddItem cOferta(iCom).iNombre & "-[" & cOferta(iCom).iCantidad & "]"
                .lblEstado.Caption = "Ofertas del otro usuario recibidas Aceptar para terminar la transaccion luego de enviar tu oferta o Rechazar para recibir otra oferta."
                End If
                
            Next iCom
            
            .lblOro.Caption = PonerPuntos(rOro)
    End With
End Sub
Public Sub comCerrar()
    With frmNuevoComercio
        .TusItems.Clear
        .Oferta.Clear
        .Ofrecer.Clear
        Dim cCei As Long
            For cCei = 1 To 20
                With cOferta(cCei)
                    .iCantidad = 0
                    .iIndex = 0
                    .iNombre = ""
                    .iOfrece = 0
                End With
                With cItem(cCei)
                    .iCantidad = 0
                    .iIndex = 0
                    .iNombre = ""
                    .iOfrece = 0
                End With
            Next cCei
        cRecivi = False
        cOferto = False
        cNombre = ""
        Unload frmNuevoComercio
    End With
End Sub
Public Sub comAgregarOferta(Index As Integer, Cant As Integer)
If cOferto Then Exit Sub
Index = Index + 1
    If cItem(Index).iCantidad < 1 Then Exit Sub
    If cItem(Index).iCantidad < Cant Then Cant = cItem(Index).iCantidad
cItem(Index).iOfrece = cItem(Index).iOfrece + Cant
cItem(Index).iCantidad = cItem(Index).iCantidad - Cant
comCarga
frmNuevoComercio.TusItems.ListIndex = Index - 1
End Sub
Public Sub comQuitarOferta(Index As Integer, Cant As Integer)
If cOferto Then Exit Sub
    If frmNuevoComercio.Ofrecer.text = "" Then Exit Sub
Dim cFo As Long
    For cFo = 1 To 20
        If cItem(cFo).iNombre = ReadField(1, frmNuevoComercio.Ofrecer.text, Asc("-")) Then
                If Cant > cItem(cFo).iOfrece Then Cant = cItem(cFo).iOfrece
            cItem(cFo).iOfrece = cItem(cFo).iOfrece - Cant
            cItem(cFo).iCantidad = cItem(cFo).iCantidad + Cant
            Exit For
        End If
    Next cFo
comCarga
frmNuevoComercio.TusItems.ListIndex = cFo - 1
End Sub
Public Sub comEnviarOferta()
        If cOferto = True Then Exit Sub
    cOferto = True
    Dim cTempPa As String
        For iCom = 1 To 20
            cTempPa = cTempPa & iCom & "-" & cItem(iCom).iOfrece & ","
        Next iCom
        
    SendData "UOR" & uOro
    SendData "UOC" & cTempPa
    
    'Bloqueamos todo
    frmNuevoComercio.Can.Enabled = False
    frmNuevoComercio.cmdAgregarOro.Enabled = False
    frmNuevoComercio.Command3.Enabled = False
    frmNuevoComercio.Command2.Enabled = False
    frmNuevoComercio.Command1.Enabled = False
    frmNuevoComercio.Command3.Picture = LoadPicture(App.Path & "\Data\Graficos\Principal\ComercioPJ_Ofrecer_Bloqueado.jpg")
    
End Sub
Public Sub comReciviOferta(Rdata As String)
If rOro <> 0 Then frmNuevoComercio.lblOro.Caption = PonerPuntos(rOro)

frmNuevoComercio.Res(1).Enabled = True
frmNuevoComercio.Res(2).Enabled = True

frmNuevoComercio.Res(2).Picture = General_Load_Interface_Picture("ComercioPJ_Aceptar.jpg")
frmNuevoComercio.Res(1).Picture = General_Load_Interface_Picture("ComercioPJ_Rechazar.jpg")

    For iCom = 1 To 20
        With cOferta(iCom)
            cTempRead = ReadField(iCom, Rdata, Asc(","))
            .iGrhIndex = ReadField(1, cTempRead, Asc("-"))
            .iCantidad = ReadField(2, cTempRead, Asc("-"))
            .iNombre = ReadField(3, cTempRead, Asc("-"))
        End With
    Next iCom
cRecivi = True
comCarga
End Sub
Public Sub comMensaje(Texto As String, Optional red As Integer = -1, Optional Green As Integer, Optional Blue As Integer, Optional bold As Boolean = False, Optional italic As Boolean = False, Optional Pete As Boolean = False)
    With frmNuevoComercio.Consola
        .SelStart = Len(.text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
            If Not red = -1 Then .SelColor = RGB(red, Green, Blue)
        .SelText = IIf(False, Texto, Texto & vbCrLf)
    End With
End Sub
Public Sub comRespuesta(Index As Integer)
If cRecivi = False Then
    frmNuevoComercio.lblEstado = "Debes esperar a que el otro usuario te envie una Oferta."
    Exit Sub
End If
SendData "TDR" & Index
End Sub
Public Sub comDibujarTusItems(Index As Integer)
Dim cSR As RECT, cDR As RECT
cSR.left = 0
cSR.top = 0
cSR.Right = 32
cSR.bottom = 32

cDR.left = 0
cDR.top = 0
cDR.Right = 32
cDR.bottom = 32

    Call engine.DrawGrhtoHdc(Inventario.GrhIndex(Index + 1), cSR, frmNuevoComercio.picInv)
End Sub
Public Sub comDibujarOfe()
If frmNuevoComercio.Ofrecer.text = "" Then Exit Sub
Dim cSR As RECT, cDR As RECT
cSR.left = 0
cSR.top = 0
cSR.Right = 32
cSR.bottom = 32

cDR.left = 0
cDR.top = 0
cDR.Right = 32
cDR.bottom = 32
Dim cFo As Long
    For cFo = 1 To 20
        If cItem(cFo).iNombre = ReadField(1, frmNuevoComercio.Ofrecer.text, Asc("-")) Then
                Call engine.DrawGrhtoHdc(Inventario.GrhIndex(cFo), cSR, frmNuevoComercio.picInv)
            Exit For
        End If
    Next cFo
End Sub
Public Sub comDibujarRec(Index As Integer)
If cRecivi = False Then Exit Sub
Dim cSR As RECT, cDR As RECT
cSR.left = 0
cSR.top = 0
cSR.Right = 32
cSR.bottom = 32

cDR.left = 0
cDR.top = 0
cDR.Right = 32
cDR.bottom = 32

    Call engine.DrawGrhtoHdc(cOferta(Index + 1).iGrhIndex, cSR, frmNuevoComercio.picInv)
End Sub


