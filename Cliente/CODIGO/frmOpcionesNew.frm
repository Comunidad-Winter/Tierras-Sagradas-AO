VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Configuracion"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Rendimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   7
      Left            =   8640
      TabIndex        =   70
      Top             =   3960
      Width           =   6135
      Begin MSComctlLib.Slider slPerformance 
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Max             =   4
         SelStart        =   2
         Value           =   2
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Nivel Maximo: Se mostraran todos los efectos. Recomendado para Pentium 4 o AMD Athlon con 1GB  de ram o superior."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   5160
         Width           =   5895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   $"frmOpcionesNew.frx":0000
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   78
         Top             =   4320
         Width           =   5895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   $"frmOpcionesNew.frx":00E0
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   77
         Top             =   3480
         Width           =   5895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   $"frmOpcionesNew.frx":018A
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   76
         Top             =   2880
         Width           =   5895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   $"frmOpcionesNew.frx":0214
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   75
         Top             =   2040
         Width           =   5895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   $"frmOpcionesNew.frx":02B7
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Nivel Actual:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   73
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Level 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Medio"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   72
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Informacion"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   8
      Left            =   -4800
      TabIndex        =   80
      Top             =   5400
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "Abrir Reproductor de Radio AO"
         Height          =   375
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   1800
         Width           =   5655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ir al Sitio Oficial"
         Height          =   375
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   83
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ir al Foro"
         Height          =   375
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   840
         Width           =   5655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ver el Mini Manual"
         Height          =   375
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   81
         Top             =   1320
         Width           =   5655
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Extras"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   6
      Left            =   6480
      TabIndex        =   63
      Top             =   2640
      Width           =   6135
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "No mover pantalla principal con el mouse si el juego esta en modo ventana."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   97
         Top             =   3840
         Width           =   5655
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00000000&
         Caption         =   "Mini Mapa"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   84
         Top             =   2760
         Width           =   5655
         Begin VB.CheckBox chkMiniMapa 
            BackColor       =   &H00000000&
            Caption         =   "Usar Mini Mapa"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   86
            Top             =   180
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkMiniMapaPos 
            BackColor       =   &H00000000&
            Caption         =   "No mostrar posicion (PC Viejas)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   85
            Top             =   540
            Value           =   1  'Checked
            Width           =   2900
         End
      End
      Begin VB.CheckBox chkUseEmoticons 
         BackColor       =   &H00000000&
         Caption         =   "Reemplazar texto por Emoticon (Ejemplo "":P"" o ""XD"")"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   1200
         Value           =   1  'Checked
         Width           =   5775
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00000000&
         Caption         =   "Formato de las Screen Shots"
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   66
         Top             =   1680
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "JPG (Menor peso, peor calidad)"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "BMP (Pesadas, alta calidad)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.CheckBox chkDeath 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Cartel de muerte."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkShowNicks 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Nicks"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Configuracion de Controles y Macro Interno"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   5
      Left            =   2880
      TabIndex        =   56
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox chkNumPad 
         BackColor       =   &H00000000&
         Caption         =   "No usar teclado numerico para cambiar modo de Habla."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Top             =   3840
         Width           =   5655
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         Caption         =   "Acciones con el mouse"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   240
         TabIndex        =   59
         Top             =   1800
         Width           =   5655
         Begin VB.CheckBox chkDesplegarMenu 
            BackColor       =   &H00000000&
            Caption         =   "Al hacer click derecho/doble click sobre un usuario desplegar menu contextual."
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   62
            Top             =   1200
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkUsarDoble 
            BackColor       =   &H00000000&
            Caption         =   "Abrir puertas/comerciar/interactuar con NPCs con doble click."
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   61
            Top             =   240
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkDerechoAsDoble 
            BackColor       =   &H00000000&
            Caption         =   "Usar Click derecho como doble click."
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   60
            Top             =   720
            Value           =   1  'Checked
            Width           =   5295
         End
      End
      Begin VB.CommandButton cmdConfigurarMacro 
         Caption         =   "Configurar Macro Interno"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   1080
         Width           =   5655
      End
      Begin VB.CommandButton cmdConfigTeclas 
         Caption         =   "Configurar Teclas"
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Chat"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   4
      Left            =   2640
      TabIndex        =   48
      Top             =   5880
      Width           =   6135
      Begin VB.CheckBox chkContactSignsOut 
         BackColor       =   &H00000000&
         Caption         =   "Avisar cuando un contacto deslogea"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkContactSignsIn 
         BackColor       =   &H00000000&
         Caption         =   "Avisar cuando un contacto inicia secion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   1560
         Value           =   1  'Checked
         Width           =   5175
      End
      Begin VB.CheckBox chkUseJugandoTP 
         BackColor       =   &H00000000&
         Caption         =   "Usar Mensaje ""Jugando TSAO"" en MSN"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Value           =   2  'Grayed
         Width           =   4335
      End
      Begin VB.CheckBox chkUseSoundAlert 
         BackColor       =   &H00000000&
         Caption         =   "Usar alertas de sonido"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   400
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00000000&
         Caption         =   "Extra"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   49
         Top             =   2760
         Width           =   5655
         Begin VB.TextBox txtFontSize 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3720
            TabIndex        =   50
            Text            =   "11"
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000007&
            Caption         =   "Tamaño de fuente de la consola de Chat:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   3135
         End
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de la Consola"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   3
      Left            =   2880
      TabIndex        =   43
      Top             =   5520
      Width           =   6135
      Begin VB.CheckBox chkNoPrivates 
         BackColor       =   &H00000000&
         Caption         =   "Desactivar Privados"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox chkNoGlobal 
         BackColor       =   &H00000000&
         Caption         =   "Desactivar Globales"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox chkUseOldFont 
         BackColor       =   &H00000000&
         Caption         =   "Usar fuente vieja en consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox ChkShowChat 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar mensajes de los usuarios en consola."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   5775
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Interfaz"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -120
      TabIndex        =   36
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Interfaz"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   2
      Left            =   3000
      TabIndex        =   35
      Top             =   5520
      Width           =   6135
      Begin VB.FileListBox InterfazFiles 
         Height          =   1065
         Left            =   720
         TabIndex        =   90
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         Caption         =   "Opciones de Skins"
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   5655
         Begin VB.CommandButton Command2 
            Caption         =   "Borrar Interfaz"
            Height          =   375
            Left            =   2880
            TabIndex        =   89
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar Interfaz"
            Height          =   375
            Left            =   240
            TabIndex        =   88
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ListBox SkinList 
            Height          =   1230
            ItemData        =   "frmOpcionesNew.frx":037A
            Left            =   240
            List            =   "frmOpcionesNew.frx":0384
            TabIndex        =   42
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   0
         TabIndex        =   37
         Top             =   2040
         Width           =   5655
         Begin VB.CheckBox chkTransparencia 
            BackColor       =   &H00000000&
            Caption         =   "Usar Transparencia"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   0
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin MSComctlLib.Slider slAlphaLevel 
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Min             =   40
            Max             =   255
            SelStart        =   180
            TickFrequency   =   10
            Value           =   180
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            Caption         =   "Transparencia de la Interfaz:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Video"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   1
      Left            =   1680
      TabIndex        =   28
      Top             =   5760
      Width           =   6135
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOpcionesNew.frx":03AF
         Left            =   1320
         List            =   "frmOpcionesNew.frx":03BF
         TabIndex        =   105
         Text            =   "65 FPS"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkTextoSube 
         BackColor       =   &H00000000&
         Caption         =   "Texto suben sobre la cabeza"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   103
         Top             =   600
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkContadores 
         BackColor       =   &H00000000&
         Caption         =   "Contadores de tiempo de invisibilidad, paralisis, etc."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   102
         Top             =   840
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkAuras 
         BackColor       =   &H00000000&
         Caption         =   "Sistema de Auras"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   101
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkSangre 
         BackColor       =   &H00000000&
         Caption         =   "Efecto simple de sangre"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   100
         Top             =   1200
         Width           =   5655
      End
      Begin VB.CheckBox chkParticulas 
         BackColor       =   &H00000000&
         Caption         =   "Efecto de particulas"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -2280
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   91
         Top             =   -120
         Value           =   1  'Checked
         Width           =   5655
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Sombras y Reflejos"
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   5655
         Begin VB.CheckBox chkReflejos 
            BackColor       =   &H00000000&
            Caption         =   "Reflejos en el agua"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   960
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkSombrasNPCs 
            BackColor       =   &H00000000&
            Caption         =   "Sombras en NPCs"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   600
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.CheckBox chkSombrasPJs 
            BackColor       =   &H00000000&
            Caption         =   "Sombras en personajes"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Value           =   1  'Checked
            Width           =   5295
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00000000&
         Caption         =   "Transparencia"
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   29
         Top             =   4200
         Width           =   5655
         Begin VB.CheckBox ckTranspPJs 
            BackColor       =   &H00000000&
            Caption         =   "Transparencia de PJs Muertos"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.CheckBox ckArboleTechos 
            BackColor       =   &H00000000&
            Caption         =   "Transparencia de Arboles y Techos"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.CheckBox ckDiayNoche 
            BackColor       =   &H00000000&
            Caption         =   "Mostrar efectos de Mañana/Dia/Tarde/Noche"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   2  'Grayed
            Width           =   3735
         End
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar FPS a"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   104
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.Frame OptionCategory 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de Sonido"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   0
      Left            =   -2400
      TabIndex        =   12
      Top             =   2160
      Width           =   6135
      Begin VB.HScrollBar scrFXVolume 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   94
         Top             =   2040
         Value           =   100
         Width           =   5895
      End
      Begin VB.HScrollBar scrMIDIVolume 
         Height          =   255
         Left            =   120
         Max             =   51
         TabIndex        =   92
         Top             =   960
         Value           =   50
         Width           =   5895
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Reproducir un MIDI"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton cmdMidiStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   2760
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdMidiPlay 
            Caption         =   "Play"
            Height          =   375
            Left            =   1440
            TabIndex        =   26
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdMidiPrevious 
            Caption         =   ">"
            Height          =   375
            Left            =   3840
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdMidiNext 
            Caption         =   "<"
            Height          =   375
            Left            =   1320
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   23
            Text            =   "MIDI"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdOpenMP3 
         Caption         =   "Abrir Reproductor MP3"
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.HScrollBar scrMP3Volume 
         Height          =   255
         Left            =   120
         Max             =   2500
         TabIndex        =   19
         Top             =   3240
         Value           =   2500
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.CommandButton cmdActivateMP3 
         Caption         =   "Activar MP3"
         Height          =   345
         Left            =   120
         MouseIcon       =   "frmOpcionesNew.frx":03E7
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2520
         Width           =   5895
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Sistema de FX"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   5520
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton chkSoundSystem 
            BackColor       =   &H00000000&
            Caption         =   "Sistema viejo (Recomendado para PCs viejas"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   170
            Value           =   -1  'True
            Width           =   5715
         End
         Begin VB.OptionButton chkSoundSystem 
            BackColor       =   &H00000000&
            Caption         =   "Sistema nuevo (Recomendado para PCs nuevas)"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   460
            Width           =   5595
         End
      End
      Begin VB.CommandButton cmdActivateMusic 
         Caption         =   "Activar MIDI"
         Height          =   345
         Left            =   120
         MouseIcon       =   "frmOpcionesNew.frx":0539
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   360
         Width           =   5895
      End
      Begin VB.CommandButton cmdActivateFX 
         Caption         =   "Activar FX"
         Height          =   345
         Left            =   120
         MouseIcon       =   "frmOpcionesNew.frx":068B
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1440
         Width           =   5895
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Volumen FX"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Volumen MIDI"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Volumen MP3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton ExitNoSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CommandButton SaveChanges 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Informacion/Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Rendimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Extras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Controles/Macros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Chat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Consola"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Interfaz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Video"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton TipoOpcion 
         Caption         =   "Sonido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ConfigChanged As Byte
Private FpsChanged As Byte
Private Sub Check1_Click()
Configuracion.General_NoMoveScreen = Check1.Value
ConfigChanged = 1
End Sub

Private Sub chkAuras_Click()
ConfigChanged = 1
Configuracion.Video_Toggle_Aura = chkAuras.Value
End Sub

Private Sub chkContactSignsIn_Click()
Configuracion.Chat_Contact_SignsIn = chkContactSignsIn.Value
ConfigChanged = 1
End Sub

Private Sub chkParticulas_Click()
ConfigChanged = 1
Configuracion.Video_Toggle_Particulas = chkParticulas.Value
End Sub

Private Sub chkReflejos_Click()
ConfigChanged = 1
Configuracion.Video_Toggle_Reflejos = chkReflejos.Value
End Sub

Private Sub chkSombrasNPCs_Click()
ConfigChanged = 1
Configuracion.Video_Toggle_Sombras_NPC = chkSombrasNPCs.Value
End Sub

Private Sub chkSombrasPJs_Click()
ConfigChanged = 1
Configuracion.Video_Toggle_Sombras = chkSombrasPJs.Value
End Sub

Private Sub cmdActivateMP3_Click()
    If Configuracion.Sonido_Musica = 1 Then
            cmdActivateMP3.Caption = "Activar MP3"
            Configuracion.Sonido_MP3 = 0
            Audio.StopMidi
        Else
            cmdActivateMP3.Caption = "Desactivar MP3"
            Configuracion.Sonido_MP3 = 1
            Audio.MP3_Play CurrentMP3
        End If

        ConfigChanged = 1
End Sub

Private Sub cmdConfigurarMacro_Click()
frmMakro.Show , frmMain
End Sub
Private Sub chkContactSignsOut_Click()
Configuracion.Chat_Contact_SignsOut = chkContactSignsOut.Value
ConfigChanged = 1
End Sub

Private Sub chkCounters_Click()
'Configuracion.Consola_Show_Counters = chkCounters.value
'ConfigChanged = 1
End Sub

Private Sub chkDeath_Click()
Configuracion.General_Mostrar_Cartel_Muerte = chkDeath.Value
ConfigChanged = 1
End Sub

Private Sub chkDerechoAsDoble_Click()
Configuracion.MouseActions_RClick = chkDerechoAsDoble.Value
ConfigChanged = 1
End Sub

Private Sub chkDesplegarMenu_Click()
Configuracion.MouseActions_Activate = chkDesplegarMenu.Value
ConfigChanged = 1
End Sub

Private Sub chkDialogoNuevoEfecto_Click()
'Configuracion.General_New_Dialog_Effect = chkDialogoNuevoEfecto.value
'ConfigChanged = 1
End Sub
Private Sub chkLimitFPS_Click()
ConfigChanged = 1
End Sub
Private Sub chkMiniMapa_Click()
ConfigChanged = 1
Configuracion.MiniMap_Activate = chkMiniMapa.Value

If Configuracion.MiniMap_Activate = 0 Then
    frmMain.Minimap.Visible = False
Else
    frmMain.Minimap.Visible = True
End If

    Call DibujarPuntoMinimap
    Call DibujarMinimap
If chkMiniMapa.Value = 1 And Not SuficientePC And ConfigChanged = 1 Then MsgBox "Recomendable DESACTIVAR en esta Maquina", vbOKOnly, "Aviso"
End Sub

Private Sub chkMiniMapaPos_Click()
ConfigChanged = 1
Configuracion.MiniMap_Show_Position = chkMiniMapaPos.Value
If chkMiniMapaPos.Value = 1 And Not SuficientePC And ConfigChanged = 1 Then MsgBox "Recomendable DESACTIVAR en esta Maquina", vbOKOnly, "Aviso"
End Sub
Private Sub chkNoGlobal_Click()
Configuracion.Consola_Globales_DeActivate = chkNoGlobal.Value
ConfigChanged = 1
End Sub
Private Sub chkNoPrivates_Click()
Configuracion.Consola_Privados_DeActivate = chkNoPrivates.Value
ConfigChanged = 1
End Sub
Private Sub chkNumPad_Click()
Configuracion.Consola_Use_Num_Pad = chkNumPad.Value
ConfigChanged = 1
End Sub

Private Sub ChkShowChat_Click()
Configuracion.Consola_Show_User_Messages = ChkShowChat.Value
ConfigChanged = 1
End Sub

Private Sub chkShowNicks_Click()
Configuracion.General_Show_Nicks = chkShowNicks.Value
Nombres = chkShowNicks
ConfigChanged = 1
End Sub

Private Sub chkSoundSystem_Click(Index As Integer)
If chkSoundSystem(0).Value = True Then
    Configuracion.Sonido_Sistema_FXs = 0
Else
    Configuracion.Sonido_Sistema_FXs = 1
End If
ConfigChanged = 1
End Sub

Private Sub chkTransparencia_Click()
slAlphaLevel.Enabled = chkTransparencia.Value
Configuracion.Alpha_Interfaz_Activar = chkTransparencia.Value
Configuracion.Alpha_Interfaz_Transparencia = Configuracion.Alpha_Interfaz_Transparencia * Configuracion.Alpha_Interfaz_Activar
ConfigChanged = 1
End Sub

Private Sub chkUsarDoble_Click()
Configuracion.MouseActions_DClick = chkUsarDoble.Value
ConfigChanged = 1
End Sub

Private Sub chkUseEmoticons_Click()
Configuracion.General_Emoticons_Reeplace = chkUseEmoticons.Value
ConfigChanged = 1
End Sub

Private Sub chkUseJugandoTP_Click()
'jajajajajajaa
End Sub

Private Sub chkUseOldEffect_Click()
'Configuracion.General_Not_Use_New_Nick_Effect = chkUseOldEffect.value
'ConfigChanged = 1
End Sub

Private Sub chkUseOldFont_Click()
Configuracion.Consola_Not_Use_New_Font = chkUseOldFont.Value
ConfigChanged = 1
End Sub

Private Sub chkUseOldNickFont_Click()
' Configuracion.General_Not_Use_New_Nick_Font = chkUseOldNickFont.value
'ConfigChanged = 1
End Sub

Private Sub chkUseSoundAlert_Click()
Configuracion.Chat_Use_Sound_Alert = chkUseSoundAlert.Value
ConfigChanged = 1
End Sub

Private Sub ckArboleTechos_Click()
ConfigChanged = 1
Configuracion.Alpha_Usar_Transparencias_Objetos = ckArboleTechos.Value
If ckArboleTechos.Value = 1 And Not SuficientePC And ConfigChanged = 1 Then MsgBox "Recomendable DESACTIVAR en esta Maquina", vbOKOnly, "Aviso"
End Sub
Private Sub ckDiayNoche_Click()
ConfigChanged = 1
Configuracion.Alpha_Use_Dia_Noche = ckDiayNoche.Value
End Sub

Private Sub ckTranspPJs_Click()
ConfigChanged = 1
Configuracion.Alpha_Usar_Transparencias_PJs = ckTranspPJs.Value
If ckTranspPJs.Value = 1 And Not SuficientePC And ConfigChanged = 1 Then MsgBox "Recomendable DESACTIVAR en esta Maquina", vbOKOnly, "Aviso"
End Sub
Private Sub cmdActivateMusic_Click()


On Error Resume Next

        If Configuracion.Sonido_Musica = 1 Then
            cmdActivateMusic.Caption = "Activar MIDI"
            Configuracion.Sonido_Musica = 0
            Sound = False
            Me.scrMIDIVolume.Enabled = False
            Audio.StopMidi
        Else
            cmdActivateMusic.Caption = "Desactivar MIDI"
            Configuracion.Sonido_Musica = 1
            Me.scrMIDIVolume.Enabled = True
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
            Sound = True
        End If

ConfigChanged = 1

End Sub
Private Sub cmdActivateFX_Click()
        If Configuracion.Sonido_Fx = 1 Then
            Call Audio.StopWave
            cmdActivateFX.Caption = "Activar FX"
            Configuracion.Sonido_Fx = 0
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
            Me.scrFXVolume.Enabled = False
            ClientSetup.bNoSound = 0
            Sound = False
        Else
            cmdActivateFX.Caption = "Desactivar FX"
            Configuracion.Sonido_Fx = 1
            Me.scrFXVolume.Enabled = True
            Me.scrFXVolume.Value = Audio.SoundVolume
            ClientSetup.bNoSound = 1
            Sound = True
        End If
    
ConfigChanged = 1
End Sub
Private Sub cmdConfigTeclas_Click()
Call frmTeclas.Show(vbModeless, frmMain)
End Sub

Private Sub ExitNoSave_Click()
If ConfigChanged = 1 Then
    If MsgBox("La configuracion ha cambiado, ¿Salir sin guardar?", vbYesNo) = vbYes Then
        
        'ReloadVals
        Unload Me
    End If
Else
    Unload Me
End If

ConfigChanged = 0
End Sub
Private Sub Form_Load()
ConfigChanged = 0
LoadOptions
Dim i As Integer
For i = 0 To 8
    OptionCategory(i).left = 2520
    OptionCategory(i).top = 0
Next i

SkinList.Clear
For i = 1 To UBound(Interfaces)
    SkinList.AddItem Interfaces(i)
Next i
SkinList.ListIndex = Configuracion.Interfaz_Skin
End Sub

Private Sub SaveChanges_Click()

If Configuracion.MiniMap_Show_Position = 1 Then
frmMain.Puntito.Visible = False
Else
frmMain.Puntito.Visible = True
End If

If Combo1.text = "65 FPS" Then
    Configuracion.General_Limit_FPS = 65
ElseIf Combo1.text = "18 FPS" Then
    Configuracion.General_Limit_FPS = 18
ElseIf Combo1.text = "32 FPS" Then
    Configuracion.General_Limit_FPS = 32
ElseIf Combo1.text = "FPS LIBRES" Then
    Configuracion.General_Limit_FPS = 0
End If

SaveVals
Unload Me

End Sub

Private Sub slAlphaLevel_Click()
Configuracion.Alpha_Interfaz_Transparencia = slAlphaLevel.Value
ConfigChanged = 1
End Sub
Private Sub TipoOpcion_Click(Index As Integer)
Dim i As Integer
For i = 0 To 8
    If i <> Index Then OptionCategory(i).Visible = False
    If i <> Index Then TipoOpcion(i).BackColor = &H8000000F
Next i
OptionCategory(Index).Visible = True
TipoOpcion(Index).BackColor = &H8000000D
End Sub

Private Sub slPerformance_Change()
If Not SuficientePC And slPerformance.Value > 2 Then MsgBox "La Memoria RAM de tu PC es demasiado baja se recomiendan al menos 512 MB de RAM para una configuracion mejor a Media/Baja", vbOKOnly, "Advertencia"

Select Case slPerformance.Value
    Case 0
        Configuracion.Alpha_Interfaz_Transparencia = 255
        Configuracion.Alpha_Usar_Transparencias_PJs = 0
        Configuracion.Alpha_Usar_Transparencias_Objetos = 0
        Configuracion.Alpha_Interfaz_Activar = 0
        Configuracion.Sonido_Sistema_FXs = 0
        Configuracion.MiniMap_Activate = 0
        Configuracion.MiniMap_Show_Position = 0
        Configuracion.Consola_Show_Counters = 0
        Configuracion.General_Engine_Speed = 0
        Configuracion.Video_Toggle_Reflejos = 0
        Configuracion.Video_Particulas_Intensidad = 0
        Configuracion.Video_Toggle_Aura = 0
        Configuracion.Video_Toggle_Particulas = 0
        Configuracion.Video_Toggle_Reflejos_Movimiento = 0
        Configuracion.Video_Toggle_Sombras = 0
        Configuracion.Video_Toggle_Sombras_NPC = 0
        Level = "Minimo"
    Case 1
        Configuracion.Alpha_Interfaz_Transparencia = 180
        Configuracion.Alpha_Interfaz_Activar = 1
        Configuracion.Sonido_Sistema_FXs = 0
        Configuracion.MiniMap_Activate = 0
        Configuracion.MiniMap_Show_Position = 0
        Configuracion.Consola_Show_Counters = 0
        Configuracion.Alpha_Usar_Transparencias_PJs = 0
        Configuracion.Alpha_Usar_Transparencias_Objetos = 0
        Configuracion.General_Engine_Speed = 0
        Configuracion.Video_Toggle_Reflejos = 0
        Configuracion.Video_Particulas_Intensidad = 100
        Configuracion.Video_Toggle_Aura = 0
        Configuracion.Video_Toggle_Particulas = 0
        Configuracion.Video_Toggle_Reflejos_Movimiento = 1
        Configuracion.Video_Toggle_Sombras = 0
        Configuracion.Video_Toggle_Sombras_NPC = 0
        Level = "Bajo"
    Case 2
        Configuracion.Alpha_Interfaz_Transparencia = 180
        Configuracion.Alpha_Interfaz_Activar = 1
        Configuracion.Sonido_Sistema_FXs = 0
        Configuracion.MiniMap_Activate = 1
        Configuracion.MiniMap_Show_Position = 0
        Configuracion.Consola_Show_Counters = 1
        Configuracion.Alpha_Usar_Transparencias_PJs = 1
        Configuracion.Alpha_Usar_Transparencias_Objetos = 0
        Configuracion.General_Engine_Speed = 0
        Configuracion.Video_Toggle_Reflejos = 1
        Configuracion.Video_Particulas_Intensidad = 100
        Configuracion.Video_Toggle_Aura = 1
        Configuracion.Video_Toggle_Particulas = 1
        Configuracion.Video_Toggle_Reflejos_Movimiento = 1
        Configuracion.Video_Toggle_Sombras = 1
        Configuracion.Video_Toggle_Sombras_NPC = 1
        Level = "Medio"
    Case 3
        Configuracion.Alpha_Interfaz_Transparencia = 180
        Configuracion.Alpha_Interfaz_Activar = 1
        Configuracion.Sonido_Sistema_FXs = 0
        Configuracion.MiniMap_Activate = 1
        Configuracion.MiniMap_Show_Position = 1
        Configuracion.Consola_Show_Counters = 1
        Configuracion.Alpha_Usar_Transparencias_PJs = 1
        Configuracion.Alpha_Usar_Transparencias_Objetos = 1
        Configuracion.General_Engine_Speed = 0
        Configuracion.Video_Toggle_Reflejos = 1
        Configuracion.Video_Particulas_Intensidad = 100
        Configuracion.Video_Toggle_Aura = 1
        Configuracion.Video_Toggle_Particulas = 1
        Configuracion.Video_Toggle_Reflejos_Movimiento = 1
        Configuracion.Video_Toggle_Sombras = 1
        Configuracion.Video_Toggle_Sombras_NPC = 1
        Level = "Alto"
    Case 4
        Configuracion.Alpha_Interfaz_Transparencia = 180
        Configuracion.Alpha_Interfaz_Activar = 1
        Configuracion.Sonido_Sistema_FXs = 0
        Configuracion.MiniMap_Activate = 1
        Configuracion.MiniMap_Show_Position = 1
        Configuracion.Consola_Show_Counters = 1
        Configuracion.Alpha_Usar_Transparencias_PJs = 1
        Configuracion.Alpha_Usar_Transparencias_Objetos = 1
        Configuracion.General_Engine_Speed = 3
        Configuracion.Video_Toggle_Reflejos = 1
        Configuracion.Video_Particulas_Intensidad = 100
        Configuracion.Video_Toggle_Aura = 1
        Configuracion.Video_Toggle_Particulas = 1
        Configuracion.Video_Toggle_Reflejos_Movimiento = 1
        Configuracion.Video_Toggle_Sombras = 1
        Configuracion.Video_Toggle_Sombras_NPC = 1
        Level = "Maximo"
End Select
Configuracion.Performance_Level = slPerformance.Value
End Sub
Private Sub scrFXVolume_Change()
Audio.SoundVolume = scrFXVolume.Value
End Sub
Private Sub scrFXVolume_Scroll()
Audio.SoundVolume = scrFXVolume.Value
End Sub
Private Sub scrMIDIVolume_Change()
Audio.MusicVolume = scrMIDIVolume.Value
End Sub
Private Sub scrMIDIVolume_Scroll()
Audio.MusicVolume = scrMIDIVolume.Value
End Sub
Private Sub scrMP3Volume_Change()
Configuracion.Sonido_MP3_Volumen = scrMP3Volume.Value
'frmMain.mp4.Volume = Configuracion.Sonido_MP3_Volumen
ConfigChanged = 1
End Sub

Private Sub cmdMidiNext_Click()
If Not IsNumeric(Text1) Then
    Mensaje.Escribir "Ingresa un numero de MIDI"
    Exit Sub
End If
Text1.text = Text1.text - 1
'CurMidi = Val(Text1) & ".mid"
'LoopMidi = 1
'Call CargarMIDI(DirMidi & CurMidi)
'Call Play_Midi
'Call Audio.PlayMIDI(CurMidi, LoopMidi)
End Sub

Private Sub cmdMidiPlay_Click()
If Not IsNumeric(Text1) Then
    Mensaje.Escribir "Ingresa un numero de MIDI"
    Exit Sub
End If
'CurMidi = Val(Text1) & ".mid"
'LoopMidi = 1
'Call CargarMIDI(DirMidi & CurMidi)
'Call Play_Midi
'Call Audio.PlayMIDI(CurMidi, LoopMidi)
End Sub

Private Sub cmdMidiPrevious_Click()
If Not IsNumeric(Text1) Then
    Mensaje.Escribir "Ingresa un numero de MIDI"
    Exit Sub
End If
Text1.text = Text1.text + 1
'CurMidi = Val(Text1) & ".mid"
'LoopMidi = 1
'Call CargarMIDI(DirMidi & CurMidi)
'Call Play_Midi
'Call Audio.PlayMIDI(CurMidi, LoopMidi)
End Sub
Private Sub cmdMidiStop_Click()
'Audio.StopMidi
End Sub
Sub ReloadVals()
With Configuracion
    .Alpha_Usar_Transparencias_Objetos = tmpConfiguracion.Alpha_Usar_Transparencias_Objetos
    .Alpha_Usar_Transparencias_PJs = tmpConfiguracion.Alpha_Usar_Transparencias_PJs
    .Alpha_Interfaz_Transparencia = tmpConfiguracion.Alpha_Interfaz_Transparencia
    .Alpha_Interfaz_Activar = tmpConfiguracion.Alpha_Interfaz_Activar
    .Performance_Level = tmpConfiguracion.Performance_Level
    .Interfaz_Skin = tmpConfiguracion.Interfaz_Skin
    .Sonido_Musica = tmpConfiguracion.Sonido_Musica
    .Sonido_Fx = tmpConfiguracion.Sonido_Fx
    .Sonido_MP3 = tmpConfiguracion.Sonido_MP3
    .Sonido_MP3_Volumen = tmpConfiguracion.Sonido_MP3_Volumen
    .Sonido_Sistema_FXs = tmpConfiguracion.Sonido_Sistema_FXs
    .General_Mostrar_Cartel_Muerte = tmpConfiguracion.General_Mostrar_Cartel_Muerte
    .General_Emoticons_Reeplace = tmpConfiguracion.General_Emoticons_Reeplace
    .General_NoMoveScreen = tmpConfiguracion.General_NoMoveScreen
    .General_Not_Use_New_Nick_Effect = tmpConfiguracion.General_Not_Use_New_Nick_Effect
    .General_Not_Use_New_Nick_Font = tmpConfiguracion.General_Not_Use_New_Nick_Font
    .General_Show_Nicks = tmpConfiguracion.General_Show_Nicks
    .General_Limit_FPS = tmpConfiguracion.General_Limit_FPS
    .General_New_Dialog_Effect = tmpConfiguracion.General_New_Dialog_Effect
    .General_ScreenShots_Format = tmpConfiguracion.General_ScreenShots_Format
    .Consola_Globales_DeActivate = tmpConfiguracion.Consola_Globales_DeActivate
    .Consola_Privados_DeActivate = tmpConfiguracion.Consola_Privados_DeActivate
    .Consola_Show_Counters = tmpConfiguracion.Consola_Show_Counters
    .Consola_Show_User_Messages = tmpConfiguracion.Consola_Show_User_Messages
    .Consola_Not_Use_New_Font = tmpConfiguracion.Consola_Not_Use_New_Font
    .MiniMap_Activate = tmpConfiguracion.MiniMap_Activate
    .MiniMap_Show_Position = tmpConfiguracion.MiniMap_Show_Position
    .MouseActions_Activate = tmpConfiguracion.MouseActions_Activate
    .MouseActions_RClick = tmpConfiguracion.MouseActions_RClick
    .MouseActions_DClick = tmpConfiguracion.MouseActions_DClick
    .Chat_Contact_SignsIn = tmpConfiguracion.Chat_Contact_SignsIn
    .Chat_Contact_SignsOut = tmpConfiguracion.Chat_Contact_SignsOut
    .Chat_Use_Sound_Alert = tmpConfiguracion.Chat_Use_Sound_Alert
    .Chat_Font_Size = tmpConfiguracion.Chat_Font_Size
    .Chat_ShowJugandoTP = tmpConfiguracion.Chat_ShowJugandoTP
    .Consola_Use_Num_Pad = tmpConfiguracion.Consola_Use_Num_Pad
    .Alpha_Use_Dia_Noche = tmpConfiguracion.Alpha_Use_Dia_Noche
    .Video_Particulas_Intensidad = tmpConfiguracion.Video_Particulas_Intensidad
    .Video_Toggle_Aura = tmpConfiguracion.Video_Toggle_Aura
    .Video_Toggle_Particulas = tmpConfiguracion.Video_Toggle_Particulas
    .Video_Toggle_Reflejos = tmpConfiguracion.Video_Toggle_Reflejos
    .Video_Toggle_Reflejos_Movimiento = tmpConfiguracion.Video_Toggle_Reflejos_Movimiento
    .Video_Toggle_Sombras = tmpConfiguracion.Video_Toggle_Sombras
    .Video_Toggle_Sombras_NPC = tmpConfiguracion.Video_Toggle_Sombras_NPC
End With
End Sub
Public Sub SaveVals()
With tmpConfiguracion
    .Alpha_Usar_Transparencias_Objetos = Configuracion.Alpha_Usar_Transparencias_Objetos
    .Alpha_Usar_Transparencias_PJs = Configuracion.Alpha_Usar_Transparencias_PJs
    .Alpha_Interfaz_Transparencia = Configuracion.Alpha_Interfaz_Transparencia
    .Alpha_Interfaz_Activar = Configuracion.Alpha_Interfaz_Activar
    .Performance_Level = Configuracion.Performance_Level
    .Interfaz_Skin = Configuracion.Interfaz_Skin
    .Sonido_Musica = Configuracion.Sonido_Musica
    .Sonido_Fx = Configuracion.Sonido_Fx
    .Sonido_MP3 = Configuracion.Sonido_MP3
    .Sonido_MP3_Volumen = Configuracion.Sonido_MP3_Volumen
    .Sonido_Sistema_FXs = Configuracion.Sonido_Sistema_FXs
    .General_Mostrar_Cartel_Muerte = Configuracion.General_Mostrar_Cartel_Muerte
    .General_Emoticons_Reeplace = Configuracion.General_Emoticons_Reeplace
    .General_NoMoveScreen = Configuracion.General_NoMoveScreen
    .General_Not_Use_New_Nick_Effect = Configuracion.General_Not_Use_New_Nick_Effect
    .General_Not_Use_New_Nick_Font = Configuracion.General_Not_Use_New_Nick_Font
    .General_Show_Nicks = Configuracion.General_Show_Nicks
    .General_ScreenShots_Format = Configuracion.General_ScreenShots_Format
    '.General_Limit_FPS = chkLimitFPS.value
    .General_New_Dialog_Effect = Configuracion.General_New_Dialog_Effect
    .Consola_Globales_DeActivate = Configuracion.Consola_Globales_DeActivate
    .Consola_Privados_DeActivate = Configuracion.Consola_Privados_DeActivate
    .Consola_Show_Counters = Configuracion.Consola_Show_Counters
    .Consola_Show_User_Messages = Configuracion.Consola_Show_User_Messages
    .Consola_Not_Use_New_Font = Configuracion.Consola_Not_Use_New_Font
    .MiniMap_Activate = Configuracion.MiniMap_Activate
    .MiniMap_Show_Position = Configuracion.MiniMap_Show_Position
    .MouseActions_Activate = Configuracion.MouseActions_Activate
    .MouseActions_RClick = Configuracion.MouseActions_RClick
    .MouseActions_DClick = Configuracion.MouseActions_DClick
    .Chat_Contact_SignsIn = Configuracion.Chat_Contact_SignsIn
    .Chat_Contact_SignsOut = Configuracion.Chat_Contact_SignsOut
    .Chat_Use_Sound_Alert = Configuracion.Chat_Use_Sound_Alert
    .Chat_Font_Size = Configuracion.Chat_Font_Size
    .Chat_ShowJugandoTP = Configuracion.Chat_ShowJugandoTP
    .Consola_Use_Num_Pad = Configuracion.Consola_Use_Num_Pad
    .Alpha_Use_Dia_Noche = Configuracion.Alpha_Use_Dia_Noche
    .Video_Particulas_Intensidad = Configuracion.Video_Particulas_Intensidad
    .Video_Toggle_Aura = Configuracion.Video_Toggle_Aura
    .Video_Toggle_Particulas = Configuracion.Video_Toggle_Particulas
    .Video_Toggle_Reflejos = Configuracion.Video_Toggle_Reflejos
    .Video_Toggle_Reflejos_Movimiento = Configuracion.Video_Toggle_Reflejos_Movimiento
    .Video_Toggle_Sombras = Configuracion.Video_Toggle_Sombras
    .Video_Toggle_Sombras_NPC = Configuracion.Video_Toggle_Sombras_NPC
End With
With Configuracion
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Usar_Transparencias_Objetos", .Alpha_Usar_Transparencias_Objetos)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Usar_Transparencias_PJs", .Alpha_Usar_Transparencias_PJs)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Interfaz_Transparencia", .Alpha_Interfaz_Transparencia)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Interfaz_Activar", .Alpha_Interfaz_Activar)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Performance_Level", .Performance_Level)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Interfaz_Skin", .Interfaz_Skin)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_Musica", .Sonido_Musica)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_Fx", .Sonido_Fx)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_MP3", .Sonido_MP3)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_MP3_Volumen", .Sonido_MP3_Volumen)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_Sistema_FXs", .Sonido_Sistema_FXs)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Mostrar_Cartel_Muerte", .General_Mostrar_Cartel_Muerte)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Emoticons_Reeplace", .General_Emoticons_Reeplace)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_NoMoveScreen", .General_NoMoveScreen)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Not_Use_New_Nick_Effect", .General_Not_Use_New_Nick_Effect)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Not_Use_New_Nick_Font", .General_Not_Use_New_Nick_Font)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Show_Nicks", .General_Show_Nicks)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Limit_FPS", .General_Limit_FPS)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_New_Dialog_Effect", .General_New_Dialog_Effect)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_ScreenShots_Format", .General_ScreenShots_Format)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Globales_DeActivate", .Consola_Globales_DeActivate)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Privados_DeActivate", .Consola_Privados_DeActivate)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Show_Counters", .Consola_Show_Counters)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Show_User_Messages", .Consola_Show_User_Messages)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Not_Use_New_Font", .Consola_Not_Use_New_Font)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Use_Num_Pad", .Consola_Use_Num_Pad)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MiniMap_Activate", .MiniMap_Activate)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MiniMap_Show_Position", .MiniMap_Show_Position)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MouseActions_Activate", .MouseActions_Activate)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MouseActions_RClick", .MouseActions_RClick)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MouseActions_DClick", .MouseActions_DClick)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Contact_SignsIn", .Chat_Contact_SignsIn)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Contact_SignsOut", .Chat_Contact_SignsOut)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Use_Sound_Alert", .Chat_Use_Sound_Alert)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Font_Size", .Chat_Font_Size)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Use_Dia_Noche", .Alpha_Use_Dia_Noche)
    
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Particulas_Intensidad", .Video_Particulas_Intensidad)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Aura", .Video_Toggle_Aura)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Particulas", .Video_Toggle_Particulas)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Reflejos", .Video_Toggle_Reflejos)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Reflejos_Movimiento", .Video_Toggle_Reflejos_Movimiento)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Sombras", .Video_Toggle_Sombras)
    Call WriteVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Sombras_NPC", .Video_Toggle_Sombras_NPC)
End With
'Mensaje.Escribir "Configuracion guardada."
End Sub
Public Sub LoadVals()
With Configuracion
    .Alpha_Usar_Transparencias_Objetos = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Usar_Transparencias_Objetos"))
    .Alpha_Usar_Transparencias_PJs = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Usar_Transparencias_PJs"))
    .Alpha_Interfaz_Transparencia = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Interfaz_Transparencia"))
    .Alpha_Interfaz_Activar = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Interfaz_Activar"))
    If .FirstRuneada = 0 Then
        .Performance_Level = 2
    Else
        .Performance_Level = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Performance_Level"))
    End If
    .Interfaz_Skin = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Interfaz_Skin"))
    .Sonido_Musica = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_Musica"))
    .Sonido_Fx = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_Fx"))
    .Sonido_MP3 = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_MP3"))
    .Sonido_MP3_Volumen = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_MP3_Volumen"))
    .Sonido_Sistema_FXs = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Sonido_Sistema_FXs"))
    .General_Mostrar_Cartel_Muerte = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Mostrar_Cartel_Muerte"))
    .General_Emoticons_Reeplace = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Emoticons_Reeplace"))
    .General_NoMoveScreen = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_NoMoveScreen"))
    .General_Not_Use_New_Nick_Effect = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Not_Use_New_Nick_Effect"))
    .General_Not_Use_New_Nick_Font = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Not_Use_New_Nick_Font"))
    .General_Limit_FPS = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Limit_FPS"))
    .General_New_Dialog_Effect = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_New_Dialog_Effect"))
    .General_Show_Nicks = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_Show_Nicks"))
    .General_ScreenShots_Format = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "General_ScreenShots_Format"))
    .Consola_Globales_DeActivate = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Globales_DeActivate"))
    .Consola_Privados_DeActivate = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Privados_DeActivate"))
    .Consola_Show_Counters = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Show_Counters"))
    .Consola_Show_User_Messages = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Show_User_Messages"))
    .Consola_Not_Use_New_Font = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Not_Use_New_Font"))
    .Consola_Use_Num_Pad = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Consola_Use_Num_Pad"))
    .MiniMap_Activate = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MiniMap_Activate"))
    .MiniMap_Show_Position = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MiniMap_Show_Position"))
    .MouseActions_Activate = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MouseActions_Activate"))
    .MouseActions_RClick = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MouseActions_RClick"))
    .MouseActions_DClick = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "MouseActions_DClick"))
    .Chat_Contact_SignsIn = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Contact_SignsIn"))
    .Chat_Contact_SignsOut = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Contact_SignsOut"))
    .Chat_Use_Sound_Alert = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Use_Sound_Alert"))
    .Chat_Font_Size = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_Font_Size"))
    .Chat_ShowJugandoTP = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Chat_ShowJugandoTP"))
    .Alpha_Use_Dia_Noche = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Alpha_Use_Dia_Noche"))
    .Video_Particulas_Intensidad = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Particulas_Intensidad"))
    .Video_Toggle_Aura = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Aura"))
    .Video_Toggle_Particulas = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Particulas"))
    .Video_Toggle_Reflejos = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Reflejos"))
    .Video_Toggle_Reflejos_Movimiento = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Reflejos_Movimiento"))
    .Video_Toggle_Sombras = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Sombras"))
    .Video_Toggle_Sombras_NPC = Val(GetVar(App.Path & "\Data\INIT\UserOptions.ini", "Opciones", "Video_Toggle_Sombras_NPC"))
End With
With tmpConfiguracion
    .Alpha_Usar_Transparencias_Objetos = Configuracion.Alpha_Usar_Transparencias_Objetos
    .Alpha_Usar_Transparencias_PJs = Configuracion.Alpha_Usar_Transparencias_PJs
    .Alpha_Interfaz_Transparencia = Configuracion.Alpha_Interfaz_Transparencia
    .Alpha_Interfaz_Activar = Configuracion.Alpha_Interfaz_Activar
    .Performance_Level = Configuracion.Performance_Level
    .Interfaz_Skin = Configuracion.Interfaz_Skin
    .Sonido_Musica = Configuracion.Sonido_Musica
    .Sonido_Fx = Configuracion.Sonido_Fx
    .Sonido_MP3 = Configuracion.Sonido_MP3
    .Sonido_MP3_Volumen = Configuracion.Sonido_MP3_Volumen
    .Sonido_Sistema_FXs = Configuracion.Sonido_Sistema_FXs
    .General_Mostrar_Cartel_Muerte = Configuracion.General_Mostrar_Cartel_Muerte
    .General_Emoticons_Reeplace = Configuracion.General_Emoticons_Reeplace
    .General_NoMoveScreen = Configuracion.General_NoMoveScreen
    .General_Not_Use_New_Nick_Effect = Configuracion.General_Not_Use_New_Nick_Effect
    .General_Not_Use_New_Nick_Font = Configuracion.General_Not_Use_New_Nick_Font
    .General_Show_Nicks = Configuracion.General_Show_Nicks
    .General_Limit_FPS = Configuracion.General_Limit_FPS
    .General_New_Dialog_Effect = Configuracion.General_New_Dialog_Effect
    .General_ScreenShots_Format = Configuracion.General_ScreenShots_Format
    .Consola_Globales_DeActivate = Configuracion.Consola_Globales_DeActivate
    .Consola_Privados_DeActivate = Configuracion.Consola_Privados_DeActivate
    .Consola_Show_Counters = Configuracion.Consola_Show_Counters
    .Consola_Show_User_Messages = Configuracion.Consola_Show_User_Messages
    .Consola_Not_Use_New_Font = Configuracion.Consola_Not_Use_New_Font
    .MiniMap_Activate = Configuracion.MiniMap_Activate
    .MiniMap_Show_Position = Configuracion.MiniMap_Show_Position
    .MouseActions_Activate = Configuracion.MouseActions_Activate
    .MouseActions_RClick = Configuracion.MouseActions_RClick
    .MouseActions_DClick = Configuracion.MouseActions_DClick
    .Chat_Contact_SignsIn = Configuracion.Chat_Contact_SignsIn
    .Chat_Contact_SignsOut = Configuracion.Chat_Contact_SignsOut
    .Chat_Use_Sound_Alert = Configuracion.Chat_Use_Sound_Alert
    .Chat_Font_Size = Configuracion.Chat_Font_Size
    .Chat_ShowJugandoTP = Configuracion.Chat_ShowJugandoTP
    .Consola_Use_Num_Pad = Configuracion.Consola_Use_Num_Pad
    .Alpha_Use_Dia_Noche = Configuracion.Alpha_Use_Dia_Noche
    .Video_Particulas_Intensidad = Configuracion.Video_Particulas_Intensidad
    .Video_Toggle_Aura = Configuracion.Video_Toggle_Aura
    .Video_Toggle_Particulas = Configuracion.Video_Toggle_Particulas
    .Video_Toggle_Reflejos = Configuracion.Video_Toggle_Reflejos
    .Video_Toggle_Reflejos_Movimiento = Configuracion.Video_Toggle_Reflejos_Movimiento
    .Video_Toggle_Sombras = Configuracion.Video_Toggle_Sombras
    .Video_Toggle_Sombras_NPC = Configuracion.Video_Toggle_Sombras_NPC
End With

End Sub
Sub LoadOptions()

On Error Resume Next

If Configuracion.General_Limit_FPS = "0" Then
    Me.Combo1.text = "FPS LIBRES"
Else
    Me.Combo1.text = "" & Configuracion.General_Limit_FPS & " FPS"
End If

Me.chkAuras = Configuracion.Video_Toggle_Aura
Me.chkParticulas = Configuracion.Video_Toggle_Particulas
Me.chkReflejos = Configuracion.Video_Toggle_Reflejos
Me.chkSombrasNPCs = Configuracion.Video_Toggle_Sombras_NPC
Me.chkSombrasPJs = Configuracion.Video_Toggle_Sombras

Me.chkContactSignsIn = Configuracion.Chat_Contact_SignsIn
ConfigChanged = 0
Me.chkContactSignsOut = Configuracion.Chat_Contact_SignsOut
ConfigChanged = 0
Me.chkDeath = Configuracion.General_Mostrar_Cartel_Muerte
ConfigChanged = 0
Me.chkDerechoAsDoble = Configuracion.MouseActions_RClick
ConfigChanged = 0
Me.chkNumPad = Configuracion.Consola_Use_Num_Pad
ConfigChanged = 0
Me.chkDesplegarMenu = Configuracion.MouseActions_Activate
ConfigChanged = 0
Me.chkMiniMapa = Configuracion.MiniMap_Activate
ConfigChanged = 0
Me.chkMiniMapaPos = Configuracion.MiniMap_Show_Position
ConfigChanged = 0
Me.chkNoGlobal = Configuracion.Consola_Globales_DeActivate
ConfigChanged = 0
Me.chkNoPrivates = Configuracion.Consola_Privados_DeActivate
ConfigChanged = 0
Me.ChkShowChat = Configuracion.Consola_Show_User_Messages
ConfigChanged = 0
Me.chkShowNicks = Configuracion.General_Show_Nicks
ConfigChanged = 0
Me.scrFXVolume = Audio.SoundVolume
ConfigChanged = 0
Me.scrMIDIVolume = Audio.MusicVolume
ConfigChanged = 0
If Configuracion.Sonido_Sistema_FXs = 0 Then
    Me.chkSoundSystem(0).Value = True
    Me.chkSoundSystem(1).Value = False
Else
    Me.chkSoundSystem(0).Value = False
    Me.chkSoundSystem(1).Value = True
End If
If Configuracion.General_ScreenShots_Format = 1 Then
    Option2.Value = False
    Option1.Value = True
Else
    Option2.Value = True
    Option1.Value = False
End If
' = Configuracion.Sonido_Sistema_FXs
Me.chkTransparencia = Configuracion.Alpha_Interfaz_Activar
ConfigChanged = 0
Me.chkUsarDoble = Configuracion.MouseActions_DClick
ConfigChanged = 0
'Me.chkLimitFPS.value = Configuracion.General_Limit_FPS
'ConfigChanged = 0
Me.chkUseEmoticons = Configuracion.General_Emoticons_Reeplace
ConfigChanged = 0
Me.Check1 = Configuracion.General_NoMoveScreen
'Me.chkUseJugandoTP
ConfigChanged = 0
'Me.chkUseOldEffect = Configuracion.General_Not_Use_New_Nick_Effect
ConfigChanged = 0
Me.chkUseOldFont = Configuracion.Consola_Not_Use_New_Font
ConfigChanged = 0
Me.chkUseSoundAlert = Configuracion.Chat_Use_Sound_Alert
ConfigChanged = 0
If Configuracion.Sonido_Musica = 1 Then
    Me.cmdActivateMusic.Caption = "Desactivar MIDI"
Else
    Me.cmdActivateMusic.Caption = "Activar MIDI"
End If
If Configuracion.Sonido_MP3 = 0 Then
    Me.cmdActivateMP3.Caption = "Activar MP3"
Else
    Me.cmdActivateMP3.Caption = "Desactivar MP3"
End If
If Configuracion.Sonido_Fx = 0 Then
    Me.cmdActivateFX.Caption = "Activar FX"
Else
    Me.cmdActivateFX.Caption = "Desactivar FX"
End If
Me.ckArboleTechos = Configuracion.Alpha_Usar_Transparencias_Objetos
ConfigChanged = 0
Me.ckDiayNoche = Configuracion.Alpha_Use_Dia_Noche
ConfigChanged = 0
Me.ckTranspPJs = Configuracion.Alpha_Usar_Transparencias_PJs
ConfigChanged = 0
Me.slAlphaLevel = Configuracion.Alpha_Interfaz_Transparencia
ConfigChanged = 0
Me.slPerformance = Configuracion.Performance_Level
ConfigChanged = 0
Me.scrMP3Volume = Configuracion.Sonido_MP3_Volumen
ConfigChanged = 0
Me.txtFontSize = Configuracion.Chat_Font_Size
If Configuracion.Performance_Level = 0 Then Level = "Minimo"
If Configuracion.Performance_Level = 1 Then Level = "Bajo"
If Configuracion.Performance_Level = 2 Then Level = "Medio"
If Configuracion.Performance_Level = 3 Then Level = "Alto"
If Configuracion.Performance_Level = 4 Then Level = "Maximo"
ConfigChanged = 0
End Sub
Private Sub SkinList_Click()
If SkinList.List(SkinList.ListIndex) = "" Then Exit Sub
Configuracion.Interfaz_Skin = SkinList.ListIndex
With frmMain
    'Dim I As Integer
    'For I = 0 To 3
    '    .kkkkkkkkkkkkk(I).Enabled = True
    'Next I
    
    .InvEqu.Picture = General_Load_Interface_Picture("Centronuevoinventario.jpg")

   ' .DespInv(0).Visible = True
   ' .DespInv(1).Visible = True
    .picInv.Visible = True
    .ItemName.Visible = True

    .hlst.Visible = False
    .cmdInfo.Visible = False
    .CmdLanzar.Visible = False
    
    .cmdMoverHechi(0).Visible = False
    .cmdMoverHechi(1).Visible = False
            
    .Picture = General_Load_Interface_Picture("Principal.jpg")
End With
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim location As String
Dim InfoInterfaz As String
Dim InterfazName As String

location = InputBox("Pone la carpeta donde esta la interfaz, ejemplo: 'C:\Descargas\Interfas Azul'", "Directorio de la Interfaz")
InterfazFiles.Path = location & "\"
InfoInterfaz = location & "\Interfaz.txt"

If Not FileExist(InfoInterfaz, vbNormal) Then
    MsgBox "La interfaz no contiene el archivo nesesario 'Interfaz.txt' o la direccion a la interfaz '" & location & "' es incorrecta."
    Exit Sub
End If

InterfazName = GetVar(InfoInterfaz, "Info", "N")

If InterfazName = "" Then
    MsgBox "La interfaz no tiene el archivo 'Interfaz.txt' correctamente echo."
    Exit Sub
End If

Dim i As Integer
    
Dim r As Byte

Dim Pos As Byte

r = 0

For i = 0 To SkinList.ListCount
    If SkinList.List(i) = InterfazName Then
        If MsgBox("¡Ya existe una interfaz con ese nombre!, ¿Sobreescribir?", vbYesNo) = vbNo Then
            r = 0
            Exit Sub
        Else
            r = 1
            Pos = i
        End If
    End If
Next i

If i > SkinList.ListCount Then

For i = 0 To SkinList.ListCount
    If SkinList.List(i) = "" Then
        r = 1
        Pos = i
        Exit For
    End If
Next i

End If

Call MkDir(App.Path & "\Data\GRAFICOS\" & InterfazName)

For i = 0 To InterfazFiles.ListCount
    Call FileCopy(location & "\" & InterfazFiles.List(i), App.Path & "\Data\GRAFICOS\" & InterfazName & "\" & InterfazFiles.List(i))
Next i

If r = 0 Then
    Call WriteVar(App.Path & "\Data\INIT\Interfaz.dat", "INTERFACES", "N" & UBound(Interfaces) + 1, InterfazName)
    Call WriteVar(App.Path & "\Data\INIT\Interfaz.dat", "MAIN", "Interfaces", UBound(Interfaces) + 1)
    ReDim Preserve Interfaces(UBound(Interfaces) + 1) As String
    Interfaces(UBound(Interfaces)) = InterfazName
    SkinList.AddItem InterfazName
Else
    Call WriteVar(App.Path & "\Data\INIT\Interfaz.dat", "INTERFACES", "N" & Pos + 1, InterfazName)
    Interfaces(Pos + 1) = InterfazName
    SkinList.List(Pos) = InterfazName
End If

MsgBox "¡Interfaz instalada correctamente!"

'display it
Configuracion.Interfaz_Skin = SkinList.ListIndex
With frmMain
    
    .InvEqu.Picture = General_Load_Interface_Picture("Centronuevoinventario.jpg")

    '.DespInv(0).Visible = True
    '.DespInv(1).Visible = True
    .picInv.Visible = True
    .ItemName.Visible = True

    .hlst.Visible = False
    .cmdInfo.Visible = False
    .CmdLanzar.Visible = False
    
    .cmdMoverHechi(0).Visible = False
    .cmdMoverHechi(1).Visible = False

    .Picture = General_Load_Interface_Picture("Principal.jpg")
End With

End Sub
Private Sub Command2_Click()
On Error Resume Next
If SkinList.ListIndex = -1 Then Exit Sub

If UBound(Interfaces) = SkinList.ListIndex + 1 Then
    Call WriteVar(App.Path & "\Data\INIT\Interfaz.dat", "MAIN", "Interfaces", UBound(Interfaces) - 1)
End If

Call WriteVar(App.Path & "\Data\INIT\Interfaz.dat", "INTERFACES", "N" & SkinList.ListIndex + 1, "")

Interfaces(SkinList.ListIndex + 1) = ""
SkinList.List(SkinList.ListIndex) = ""
Configuracion.Interfaz_Skin = 0

    
If MsgBox("¿Queres ademas borrar la interfaz de tu disco duro?", vbYesNo) = vbYes Then
    Call RmDir(App.Path & "\Data\GRAFICOS\" & SkinList.List(SkinList.ListIndex))
End If

MsgBox "Interfaz correctamente desinstalada."
End Sub
