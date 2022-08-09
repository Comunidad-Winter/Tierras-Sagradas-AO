Attribute VB_Name = "modTextos"
Option Explicit

Private Type tFont
    r As Byte
    g As Byte
    b As Byte
    bold As Boolean
    italic As Boolean
End Type

Public Enum FontTypeNames
FONTTYPE_GANAR = 1
FONTTYPE_CONSOLA = 2
FONTTYPE_GLOBAL = 3
FONTTYPE_UDP = 4
FONTTYPE_TALK = 5
FONTTYPE_TORNEIN = 6
FONTTYPE_ORO = 7
FONTTYPE_OROX = 8
FONTTYPE_TSUBASTA = 9
FONTTYPE_TDSUBASTA = 10
FONTTYPE_SUBASTA = 11
FONTTYPE_NPCS = 12
FONTTYPE_NPCSX = 13
FONTTYPE_ATNPC = 14
FONTTYPE_GANAORO = 15
FONTTYPE_DIOSES = 16
FONTTYPE_DIOSESI = 17
FONTTYPE_DIOSESN = 18
FONTTYPE_FIGHT = 19
FONTTYPE_WARNING = 20
FONTTYPE_INFO = 21
FONTTYPE_INFOBOLD = 22
FONTTYPE_INFOITALIC = 23
FONTTYPE_EJECUCION = 24
FONTTYPE_PARTY = 25
FONTTYPE_VENENO = 26
FONTTYPE_GUILD = 27
FONTTYPE_SERVER = 28
FONTTYPE_FORTA = 29
FONTTYPE_CASTI = 30
FONTTYPE_GUILDMSG = 31
FONTTYPE_CONSEJO = 32
FONTTYPE_CONSEJOCAOS = 33
FONTTYPE_CONSEJOVesA = 34
FONTTYPE_CONSEJOCAOSVesA = 35
FONTTYPE_CENTINELA = 36
FONTTYPE_ADVERTENCIAS = 37
FONTTYPE_AMARILLON = 38
FONTTYPE_EXPEN = 39
FONTTYPE_GRISN = 40
FONTTYPE_DAREXP = 41
FONTTYPE_ROJO = 42
FONTTYPE_GLOBALUSUARIO = 43
FONTTYPE_GLOBALNOBLE = 44
FONTTYPE_GLOBALGM = 45

'Colores Comunes
FONTTYPE_BLANCO = 46
FONTTYPE_BORDO = 47
FONTTYPE_VERDE = 48
FONTTYPE_AZUL = 49
FONTTYPE_VIOLETA = 50
FONTTYPE_AMARILLO = 51
FONTTYPE_CELESTE = 52
FONTTYPE_GRIS = 53

'Colores en negrita
FONTTYPE_BLANCON = 54
FONTTYPE_BORDON = 55
FONTTYPE_VERDEN = 56
FONTTYPE_OLIVE = 57
FONTTYPE_ROJON = 58
FONTTYPE_AZULN = 59
FONTTYPE_VIOLETAN = 60
FONTTYPE_CELESTEN = 61
FONTTYPE_DON = 62
FONTTYPE_AZULC = 63

'Colores en cursiva & negrita
FONTTYPE_BLANCOCN = 64
FONTTYPE_BORDOCN = 65
FONTTYPE_VERDECN = 66
FONTTYPE_ROJOCN = 67
FONTTYPE_AZULCN = 68
FONTTYPE_VIOLETACN = 69
FONTTYPE_CELESTECN = 70
FONTTYPE_GRISCN = 71

'Colores en cursiva
FONTTYPE_BLANCOC = 72
FONTTYPE_BORDOC = 73
FONTTYPE_VERDEC = 74
FONTTYPE_ROJOC = 75
FONTTYPE_VIOLETAC = 76
FONTTYPE_CELESTEC = 77
FONTTYPE_GRISC = 78
FONTTYPE_VERDEL = 79
FONTTYPE_CHAT = 80
FONTTYPE_REJA = 81
FONTTYPE_NARANJA = 82
FONTTYPE_CONTEO = 83
FONTTYPE_YA = 84
FONTTYPE_NEWTORNEO = 85
FONTTYPE_NARANJAN = 86
FONTTYPE_TROFEOS1 = 87
FONTTYPE_TROFEOS2 = 88
FONTTYPE_TROFEOS3 = 89
End Enum

Private Type tMessages
    text As String
    font As FontTypeNames
End Type
 
Public Messages() As tMessages
Public FontTypes(89) As tFont
Public Sub InitFonts()

    With FontTypes(FontTypeNames.FONTTYPE_GANAR)
        .r = 240
        .g = 240
        .b = 50
        .bold = True
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSOLA)
        .g = 128
        .b = 128
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
        .g = 128
        .b = 128
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_UDP)
        .r = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .r = 255
        .g = 255
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TORNEIN)
        .r = 225
        .g = 249
        .b = 158
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ORO)
        .r = 225
        .g = 222
        .b = 119
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_OROX)
        .r = 225
        .g = 222
        .b = 119
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TSUBASTA)
        .r = 255
        .g = 255
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TDSUBASTA)
        .r = 255
        .g = 255
        .b = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_SUBASTA)
        .r = 48
        .g = 128
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NPCS)
        .r = 86
        .g = 87
        .b = 89
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NPCSX)
        .r = 255
        .g = 83
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ATNPC)
        .r = 114
        .g = 0
        .b = 4
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GANAORO)
        .r = 145
        .g = 9
        .b = 179
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOSES)
        .r = 100
        .g = 0
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOSESI)
        .r = 100
        .g = 0
        .b = 255
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOSESN)
        .r = 100
        .g = 0
        .b = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .r = 255
        .g = 0
        .b = 0
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .r = 32
        .g = 51
        .b = 223
        .bold = True
        .italic = True
    End With
    
    '~69~190~156
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .r = 69
        .g = 190
        .b = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .r = 69
        .g = 190
        .b = 156
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOITALIC)
        .r = 69
        .g = 190
        .b = 156
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .r = 130
        .g = 130
        .b = 130
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .r = 255
        .g = 255
        .b = 255
        .italic = True
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).g = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .r = 255
        .g = 255
        .b = 255
        .bold = True
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).g = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_FORTA)
        .r = 177
        .g = 153
        .b = 57
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CASTI)
        .r = 255
        .g = 255
        .b = 100
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .r = 228
        .g = 199
        .b = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .g = 64
        .b = 128
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .r = 140
        .g = 0
        .b = 0
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .g = 64
        .b = 128
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .r = 140
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .g = 170
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ADVERTENCIAS)
        .r = 128
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_AMARILLON)
        .r = 255
        .g = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EXPEN)
        .r = 236
        .g = 186
        .b = 107
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GRISN)
        .r = 130
        .g = 130
        .b = 130
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DAREXP)
        .r = 255
        .g = 255
        .bold = True
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_ROJO).r = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBALUSUARIO)
        .r = 173
        .g = 170
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBALNOBLE)
        .r = 255
        .g = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBALGM)
        .g = 255
        .b = 128
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BLANCO)
        .r = 255
        .g = 255
        .b = 255
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_BORDO).r = 128
    FontTypes(FontTypeNames.FONTTYPE_VERDE).g = 255
    FontTypes(FontTypeNames.FONTTYPE_AZUL).b = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_VIOLETA)
        .r = 128
        .b = 128
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_AMARILLO)
        .r = 255
        .g = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CELESTE)
        .r = 128
        .g = 255
        .b = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GRIS)
        .r = 130
        .g = 130
        .b = 130
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BLANCON)
        .r = 255
        .g = 255
        .b = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BORDON)
        .r = 128
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VERDEN)
        .g = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_OLIVE)
        .r = 107
        .g = 142
        .b = 35
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ROJON)
        .r = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_AZULN)
        .b = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VIOLETAN)
        .r = 128
        .b = 128
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CELESTEN)
        .r = 128
        .g = 255
        .b = 255
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DON)
        .r = 255
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_AZULC)
        .g = 64
        .b = 128
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BLANCOCN)
        .r = 255
        .g = 255
        .b = 255
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BORDOCN)
        .r = 128
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VERDECN)
        .g = 255
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ROJOCN)
        .r = 255
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_AZULCN)
        .b = 255
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VIOLETACN)
        .r = 128
        .b = 128
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CELESTECN)
        .r = 128
        .g = 255
        .b = 255
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GRISCN)
        .r = 130
        .g = 130
        .b = 130
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BLANCOC)
        .r = 255
        .g = 255
        .b = 255
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_BORDOC)
        .r = 128
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VERDEC)
        .g = 255
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ROJOC)
        .r = 255
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VIOLETAC)
        .r = 128
        .b = 128
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CELESTEC)
        .r = 128
        .g = 255
        .b = 255
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GRISC)
        .r = 130
        .g = 130
        .b = 130
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VERDEL)
        .g = 128
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CHAT)
        .r = 200
        .g = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_REJA)
        .r = 177
        .g = 153
        .b = 57
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NARANJA)
        .r = 255
        .g = 128
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONTEO)
        .r = 230
        .g = 180
        .b = 10
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_YA)
        .r = 220
        .g = 83
        .b = 14
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NEWTORNEO)
        .r = 250
        .g = 210
        .b = 140
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NARANJAN)
        .r = 255
        .g = 128
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TROFEOS1)
        .r = 233
        .g = 192
        .bold = True
    End With

    With FontTypes(FontTypeNames.FONTTYPE_TROFEOS2)
        .r = 196
        .g = 196
        .b = 196
        .bold = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TROFEOS3)
        .r = 255
        .g = 128
        .b = 128
        .bold = True
    End With

End Sub
Sub LoadText()
 
Dim l_file As clsIniReader
Dim i As Long, amountMessages As Long

    Set l_file = New clsIniReader
    
    '@ load file
    l_file.Initialize App.Path & "\Data\INIT\Textos.tsao"
    
'@@ Leemos un archivo donde están todos los mensajes almacenados.
amountMessages = l_file.GetValue("TEXTOS", "Cant")
ReDim Messages(amountMessages) As tMessages
 
For i = 1 To amountMessages
    Messages(i).text = l_file.GetValue("TEXTO" & i, "Mensaje")
    Messages(i).font = l_file.GetValue("TEXTO" & i, "Font")
Next i
 
End Sub
