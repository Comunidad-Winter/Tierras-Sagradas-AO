Attribute VB_Name = "AA_Correos"
Option Explicit
Private Type cVal
    iIndex          As Long
    iCantidad       As Long
    iNombre         As String
    iOfrece         As Integer
    iGrhIndex       As Integer
End Type
Private cItem(20)   As cVal
Private cRetirar(20)   As cVal
Private cTempRead   As String
Private cTempRead2  As String

Private iCorr        As Long
Private cOferto     As Boolean
Private cRecivi     As Boolean
Public Sub correosIniciar(Rdata As String)
cNombre = ReadField(1, Rdata, Asc("$"))
cTempRead = ReadField(2, Rdata, Asc("$"))
    For iCorr = 1 To 20
        With cItem(iCorr)
            cTempRead2 = ReadField(iCorr, cTempRead, Asc(","))
            .iIndex = ReadField(1, cTempRead2, Asc("-"))
            .iCantidad = ReadField(2, cTempRead2, Asc("-"))
            .iNombre = ReadField(3, cTempRead2, Asc("-"))
        End With
    Next iCorr

correosCarga
End Sub
Public Sub correosIniciarForm(Rdata As String)

Dim i As Long
frmCorreo.lstMails.Clear

For i = 1 To 30
    frmCorreo.lstMails.AddItem ReadField(i, Rdata, Asc(","))
Next i

If frmCorreo.Visible = False Then
    frmCorreo.Show , frmMain
Else
    frmCorreo.lstMails.ListIndex = CorreoListIndex
End If

End Sub
Public Sub correosListaAmigos(Rdata As String)

Dim i As Long
frmCorreo.lstContactos.Clear

For i = 1 To 20
    If UCase$(ReadField(i, Rdata, Asc(","))) <> "(NADIE)" Then
        frmCorreo.lstContactos.AddItem ReadField(i, Rdata, Asc(","))
    End If
Next i

End Sub
Private Sub correosCarga()
    With frmCorreo
        .lstObjetos.Clear
        .lstObjs.Clear
        .lstObjsEnviar.Clear
        
            For iCorr = 1 To 20
                If cItem(iCorr).iCantidad = 0 Then
                    .lstObjs.AddItem "Nada - 0"
                Else
                    .lstObjs.AddItem cItem(iCorr).iNombre & " - " & cItem(iCorr).iCantidad & ""
                End If
                If cItem(iCorr).iOfrece > 0 Then .lstObjsEnviar.AddItem cItem(iCorr).iNombre & " - " & cItem(iCorr).iOfrece & ""
            Next iCorr
        End With
End Sub
Public Sub correosAgregarItem(Index As Integer, Cant As Integer)
Index = Index + 1
    If cItem(Index).iCantidad < 1 Then Exit Sub
    If cItem(Index).iCantidad < Cant Then Cant = cItem(Index).iCantidad
    cItem(Index).iOfrece = cItem(Index).iOfrece + Cant
    cItem(Index).iCantidad = cItem(Index).iCantidad - Cant
    
correosCarga

frmCorreo.lstObjs.ListIndex = Index - 1
End Sub
Public Sub correosQuitarItem(Index As Integer, Cant As Integer)
If frmCorreo.lstObjsEnviar.text = "" Then Exit Sub

Dim cFo As Long
    For cFo = 1 To 20
        If "" & UCase$(cItem(cFo).iNombre) & " " = UCase$(ReadField(1, frmCorreo.lstObjsEnviar.text, Asc("-"))) Or UCase$(cItem(cFo).iNombre) = UCase$(ReadField(1, frmCorreo.lstObjsEnviar.text, Asc("-"))) Then
                If Cant > cItem(cFo).iOfrece Then Cant = cItem(cFo).iOfrece
            cItem(cFo).iOfrece = cItem(cFo).iOfrece - Cant
            cItem(cFo).iCantidad = cItem(cFo).iCantidad + Cant
            Exit For
        End If
    Next cFo
    
'    frmCorreo.lstObjsEnviar.ListIndex = cFo - 1

correosCarga
End Sub
Public Sub correosEnviarItems()
Dim cTempPa As String
    For iCorr = 1 To 20
        cTempPa = cTempPa & iCorr & "-" & cItem(iCorr).iOfrece & ","
    Next iCorr
    
    SendData "CZM" & frmCorreo.txtDestinatario.text & "$" & frmCorreo.txtAsunto.text & "$" & frmCorreo.txtMensaje.text & "$" & cTempPa
    correosCerrar
End Sub
Public Sub correosCargarMensaje(Rdata As String)

frmCorreo.lblAsunto.Caption = ""
frmCorreo.lblMensaje.text = ""
frmCorreo.lstObjetos.Clear
frmCorreo.lblFecha.Caption = ""
frmCorreo.lblRemitente.Caption = ""

frmCorreo.lblRemitente.Caption = ReadField(1, Rdata, Asc("$"))
frmCorreo.lblAsunto.Caption = ReadField(2, Rdata, Asc("$"))
frmCorreo.lblMensaje.text = ReadField(3, Rdata, Asc("$"))
frmCorreo.lblFecha.Caption = ReadField(4, Rdata, Asc("$"))

Dim cDatPalOtro As String, ComienzoPaLeer As Integer

cDatPalOtro = cDatPalOtro & ReadField(1, Rdata, Asc("$")) & "$" & ReadField(2, Rdata, Asc("$")) & "$" & ReadField(3, Rdata, Asc("$")) & "$" & ReadField(4, Rdata, Asc("$")) & "$"
ComienzoPaLeer = Len(cDatPalOtro)

    For iCorr = 1 To 20
        With cRetirar(iCorr)
            If iCorr = 1 Then
                cTempRead = ReadField(1, mid(Rdata, ComienzoPaLeer), Asc(","))
            Else
                cTempRead = ReadField(iCorr, Rdata, Asc(","))
            End If
            
            .iGrhIndex = ReadField(1, cTempRead, Asc("-"))
            .iCantidad = ReadField(2, cTempRead, Asc("-"))
            .iNombre = ReadField(3, cTempRead, Asc("-"))
            
            If .iNombre <> "(Nada)" Then
                frmCorreo.lstObjetos.AddItem "" & .iNombre & " - " & .iCantidad & ""
            End If
        End With
    Next iCorr
    
    
End Sub
Public Sub correosCerrar()
    With frmCorreo
        .lstObjetos.Clear
        .lstObjs.Clear
        .lstObjsEnviar.Clear
        
        Dim cCei As Long
            For cCei = 1 To 20
                With cItem(cCei)
                    .iCantidad = 0
                    .iIndex = 0
                    .iNombre = ""
                    .iOfrece = 0
                End With
            Next cCei

        Unload frmCorreo
    End With
End Sub
