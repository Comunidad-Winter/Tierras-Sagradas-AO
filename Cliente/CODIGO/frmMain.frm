VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "Cswsk32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   -330
   ClientTop       =   1140
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6120
      Left            =   0
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   534
      TabIndex        =   39
      Top             =   1860
      Width           =   8010
      Begin SocketWrenchCtrl.Socket Socket1 
         Left            =   0
         Top             =   0
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         AutoResolve     =   0   'False
         Backlog         =   1
         Binary          =   0   'False
         Blocking        =   0   'False
         Broadcast       =   0   'False
         BufferSize      =   2048
         HostAddress     =   ""
         HostFile        =   ""
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   0
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   0
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   999999
         Type            =   1
         Urgent          =   0   'False
      End
   End
   Begin VB.PictureBox Minimap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   10230
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   38
      Top             =   6360
      Width           =   1500
      Begin VB.Shape Puntito 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H000000FF&
         Height          =   90
         Left            =   480
         Top             =   480
         Width           =   90
      End
   End
   Begin VB.Timer SegurodeItems 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3720
      Top             =   2040
   End
   Begin VB.ListBox ChatContacts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1320
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":26DAA
      Left            =   5940
      List            =   "frmMain.frx":26DEA
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   30
      Width           =   2055
   End
   Begin VB.Timer WorkMacro 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   4200
      Top             =   2040
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   4680
      Top             =   2040
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1320
      Left            =   30
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   30
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   2328
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":26EDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3131
      Left            =   3240
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox PrivatesConsole 
      Height          =   1320
      Left            =   30
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes Privados"
      Top             =   30
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   2328
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":26F5A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1485
      Visible         =   0   'False
      Width           =   7560
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   8700
      ScaleHeight     =   174
      ScaleMode       =   0  'User
      ScaleWidth      =   174
      TabIndex        =   6
      Top             =   2325
      Width           =   2610
   End
   Begin RichTextLib.RichTextBox ClanConsole 
      Height          =   1320
      Left            =   30
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes de Clan"
      Top             =   30
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   2328
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":26FE0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   45
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1485
      Visible         =   0   'False
      Width           =   7560
   End
   Begin RichTextLib.RichTextBox GlobalConsole 
      Height          =   1320
      Left            =   30
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes Globales"
      Top             =   30
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   2328
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":27066
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2385
      IntegralHeight  =   0   'False
      Left            =   8775
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2475
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Image imgMenuGral 
      Height          =   450
      Left            =   9240
      Top             =   5580
      Width           =   1575
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Left            =   7680
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image imgRanking 
      Height          =   420
      Left            =   9720
      Top             =   1380
      Width           =   615
   End
   Begin VB.Image DyD 
      Height          =   315
      Left            =   8115
      Top             =   5070
      Width           =   315
   End
   Begin VB.Label rep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   9240
      TabIndex        =   37
      Top             =   780
      Width           =   480
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,00,00)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8400
      TabIndex        =   36
      Top             =   8355
      Width           =   3195
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   7365
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label ItemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   35
      Top             =   5055
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   9720
      Top             =   5040
      Width           =   180
   End
   Begin VB.Image ImgSkills 
      Height          =   375
      Left            =   9600
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image ImgEstadisticas 
      Height          =   495
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Image ImgCastillos 
      Height          =   375
      Left            =   9480
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Image ImgPremios 
      Height          =   375
      Left            =   9480
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image ImgEquipo 
      Height          =   495
      Left            =   9600
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "587.654.333"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   150
      Left            =   10440
      TabIndex        =   27
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11640
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8700
      TabIndex        =   24
      Top             =   585
      Width           =   210
   End
   Begin VB.Label ONLINES 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   705
      TabIndex        =   23
      Top             =   8205
      Width           =   255
   End
   Begin VB.Label FPSMain 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   705
      TabIndex        =   22
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Defensa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "40/40"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4260
      TabIndex        =   21
      Top             =   8355
      Width           =   855
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "15/15"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      TabIndex        =   20
      Top             =   8355
      Width           =   855
   End
   Begin VB.Label DefMag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10/10"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   19
      Top             =   8355
      Width           =   855
   End
   Begin VB.Label GMSOS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOS"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   8700
      Width           =   975
   End
   Begin VB.Label GMTORNEO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TORNEO"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10680
      TabIndex        =   11
      Top             =   8700
      Width           =   975
   End
   Begin VB.Label GMPANEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CASTI GM"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   8700
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel M�ximo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   11640
      TabIndex        =   9
      Top             =   75000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   11520
      Top             =   0
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11760
      Top             =   0
      Width           =   270
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6705
      TabIndex        =   7
      Top             =   8355
      Width           =   495
   End
   Begin VB.Image CmdLanzar 
      Height          =   600
      Left            =   8040
      MouseIcon       =   "frmMain.frx":270EC
      MousePointer    =   99  'Custom
      Top             =   4905
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10800
      MouseIcon       =   "frmMain.frx":2723E
      MousePointer    =   99  'Custom
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11490
      MouseIcon       =   "frmMain.frx":27390
      MousePointer    =   99  'Custom
      Top             =   3705
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   11490
      MouseIcon       =   "frmMain.frx":274E2
      MousePointer    =   99  'Custom
      Top             =   3330
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   10080
      MouseIcon       =   "frmMain.frx":27634
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1800
      Width           =   1875
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8040
      MouseIcon       =   "frmMain.frx":27786
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WENDEL"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   8115
      TabIndex        =   2
      Top             =   90
      Width           =   3825
   End
   Begin VB.Image ExpBar 
      Appearance      =   0  'Flat
      Height          =   75
      Left            =   8025
      Picture         =   "frmMain.frx":278D8
      Top             =   1275
      Width           =   3960
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   1095
      Index           =   3
      Left            =   0
      TabIndex        =   30
      Top             =   7920
      Width           =   8535
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   1815
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label ManaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   6675
      Width           =   1815
   End
   Begin VB.Image MPShp 
      Height          =   180
      Left            =   8160
      Picture         =   "frmMain.frx":2DA75
      Top             =   6675
      Width           =   1770
   End
   Begin VB.Label StaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   135
      Left            =   8160
      TabIndex        =   15
      Top             =   7020
      Width           =   1800
   End
   Begin VB.Label AGUABAR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   18
      Top             =   7350
      Width           =   1800
   End
   Begin VB.Image AguaSP 
      Height          =   180
      Left            =   8160
      Picture         =   "frmMain.frx":301C6
      Top             =   7335
      Width           =   1770
   End
   Begin VB.Image SPShp 
      Height          =   180
      Left            =   8160
      Picture         =   "frmMain.frx":3294C
      Top             =   7005
      Width           =   1770
   End
   Begin VB.Label COMIDABAR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   7680
      Width           =   1800
   End
   Begin VB.Image COMIDASp 
      Height          =   180
      Left            =   8160
      Picture         =   "frmMain.frx":350B2
      Top             =   7665
      Width           =   1770
   End
   Begin VB.Label exp 
      BackStyle       =   0  'Transparent
      Caption         =   "999999/999999"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   9240
      TabIndex        =   26
      Top             =   990
      Width           =   2100
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8160
      TabIndex        =   14
      Top             =   6345
      Width           =   1800
   End
   Begin VB.Image HpSHP 
      Height          =   180
      Left            =   8160
      Picture         =   "frmMain.frx":37811
      Top             =   6345
      Width           =   1770
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   8355
      Width           =   495
   End
   Begin VB.Image InvEqu 
      Height          =   3690
      Left            =   8025
      Top             =   1845
      Width           =   3960
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   9015
      Index           =   0
      Left            =   8040
      TabIndex        =   29
      Top             =   0
      Width           =   3975
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu cmdAmigos 
      Caption         =   "Amigos"
      Visible         =   0   'False
      Begin VB.Menu cmdAddC 
         Caption         =   "Agregar a un contacto nuevo.."
      End
      Begin VB.Menu cmdIniciarChat 
         Caption         =   "Iniciar chat"
      End
      Begin VB.Menu cmdChat 
         Caption         =   "Sacar/Restaurar admision a contacto seleccionado"
      End
      Begin VB.Menu cmdBorrarC 
         Caption         =   "Eliminar contacto seleccionado"
      End
   End
   Begin VB.Menu mnuVista 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuHabla 
         Caption         =   "Modo de Habla"
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu mnuGritar 
            Caption         =   "Gritar"
         End
         Begin VB.Menu mnuClanes 
            Caption         =   "Clanes"
         End
         Begin VB.Menu mnuGlo 
            Caption         =   "Global"
         End
         Begin VB.Menu mnuParty 
            Caption         =   "Party"
         End
         Begin VB.Menu mnuFaccion 
            Caption         =   "Faccion"
         End
         Begin VB.Menu mnudenunciar 
            Caption         =   "Denunciar"
         End
         Begin VB.Menu mnuprivado 
            Caption         =   "Privado"
         End
      End
      Begin VB.Menu cmdConsole 
         Caption         =   "Activar/desactivar consola"
      End
      Begin VB.Menu cmdContacts 
         Caption         =   "Mostrar/ocultar contactos"
      End
      Begin VB.Menu cmdConsoleClan 
         Caption         =   "Mostrar/ocultar consola clan"
      End
      Begin VB.Menu cmdGlobal 
         Caption         =   "Mostrar/ocultar consola globales"
      End
      Begin VB.Menu cmdPrivados 
         Caption         =   "Mostrar/ocultar consola privados"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private form_Moviment As clsFormMovementManager
Public InvMouseBoton As Long

Public InvMouseLanzar As Long

Public InvMousePantalla As Long

Public MouseRenderX As Long
Public MouseRenderY As Long

Private TiempoActual As Long
Private Contador As Integer
Dim UserStartX, UserStartY
Dim L As Integer

Public MouseX As Long
Public MouseY As Long

Private minimapMouseX As Long
Private minimapMouseY As Long

Public ActualSecond As Long
Public lastSecond As Long
Public tX As Byte
Public tY As Byte
Private clicX As Long
Private clicY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public MouseXInv As Single
Public MouseYInv As Single

Public IsPlaying As Byte
Dim cantDet As Integer

Dim endEvent As Long
Dim PuedeMacrear As Boolean
Public Sub PonerListaAmigos(ByVal rData As String)

Dim j As Integer, k As Integer
For j = 0 To ChatContacts.ListCount - 1
    Me.ChatContacts.RemoveItem 0
Next j
k = CInt(ReadField(1, rData, 44))

For j = 1 To k
    ChatContacts.AddItem ReadField(1 + j, rData, 44)
Next j

If k < 20 Then
    For j = k + 1 To 20
        ChatContacts.AddItem "(NADIE)(OFF)"
    Next j
End If

ChatContacts.Refresh

End Sub

Private Sub ChatContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
PopupMenu cmdAmigos
End If

End Sub

Private Sub cmdConsoleClan_Click()
If ClanConsole.Visible = False Then 'ta invisible la activamos
    RecTxt.Visible = False
    PrivatesConsole.Visible = False
    GlobalConsole.Visible = False
    ClanConsole.Visible = True
Else
    RecTxt.Visible = True
    PrivatesConsole.Visible = False
    GlobalConsole.Visible = False
    ClanConsole.Visible = False
End If
End Sub

Private Sub cmdAddC_Click()
On Error Resume Next
Dim Name As String
Name = InputBox("�Nombre?")

If Name = "" Or IsNumeric(Name) Or Len(Name) > 15 Then
    Mensaje.Escribir "Nombre invalido."
    Exit Sub
End If

Call SendData("ADDCON" & Name)

End Sub
Private Sub cmdIniciarChat_Click()
On Error Resume Next
Dim i As Long, axx As Byte
axx = 0

    For i = 1 To 5
        If ChatEnUso(i) = True Then
        axx = axx + 1
        End If
        
        If axx = 5 Then
        Mensaje.Escribir "No podes abrir m�s de 5 ventanas de chat al mismo tiempo."
        Exit Sub
        End If
    Next i

Call SendData("INCHAT" & ChatContacts.ListIndex + 1)

End Sub
Private Sub cmdBorrarC_Click()
If ChatContacts.ListIndex < 0 Then
    Mensaje.Escribir "Selecciona un contacto con click primero."
    Exit Sub
End If

If ChatContacts.text = "(NADIE)(OFF)" Then Exit Sub
If MsgBox("�Eliminar?", vbYesNo) = vbYes Then Call SendData("BORRAC" & ChatContacts.ListIndex + 1)
End Sub
Private Sub CmdLanzar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
InvMouseLanzar = Button
If Not GetAsyncKeyState(Button) < 0 Then InvMouseLanzar = 0
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.ListIndex + 1)

Select Case Index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select

End Sub
Private Sub cmdConsole_Click()
If UserConsola = False Then
    Call AddtoRichTextBox(frmMain.RecTxt, ">>Consola General Desactivada.", 255, 255, 255, True, False, False)
    UserConsola = True
Else
    UserConsola = False
    Call AddtoRichTextBox(frmMain.RecTxt, ">>Consola General Activada.", 255, 255, 255, True, False, False)
End If
End Sub
Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, False)
        Exit Sub
    End If
    TrainingMacro.Interval = 550
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, False)
End Sub

Public Sub DesactivarMacroHechizos()
        TrainingMacro.Enabled = False
        SecuenciaMacroHechizos = 0
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, False)
End Sub

Private Sub Command1_Click()
    LoadGrhData
    MsgBox "ta"
    
End Sub

Private Sub DyD_Click()

If DyDActivado = False Then
    DyD.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\D&D2_N.jpg")
    DyDActivado = True
ElseIf DyDActivado = True Then
    DyD.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\D&D_N.jpg")
    DyDActivado = False
        MouseRendOK = False
        DibujadoContinuoInv = False
        ButtonIN = False
        PUEDO = False
        MouseOK = False
        MouseItem = 0
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
InvMousePantalla = Button
If Not GetAsyncKeyState(Button) < 0 Then InvMousePantalla = 0

    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub FPS_Timer()

    If RecibioMensaje(1) = True Then
            If VentanitaMostrar(1) = 0 Then
              VentanitaMostrar(1) = 1
            ElseIf VentanitaMostrar(1) = 1 Then
              VentanitaMostrar(1) = 0
            End If
    ElseIf RecibioMensaje(2) = True Then
            If VentanitaMostrar(2) = 0 Then
              VentanitaMostrar(2) = 1
            ElseIf VentanitaMostrar(2) = 1 Then
              VentanitaMostrar(2) = 0
            End If
    ElseIf RecibioMensaje(3) = True Then
            If VentanitaMostrar(3) = 0 Then
              VentanitaMostrar(3) = 1
            ElseIf VentanitaMostrar(3) = 1 Then
              VentanitaMostrar(3) = 0
            End If
    ElseIf RecibioMensaje(4) = True Then
            If VentanitaMostrar(4) = 0 Then
              VentanitaMostrar(4) = 1
            ElseIf VentanitaMostrar(4) = 1 Then
              VentanitaMostrar(4) = 0
            End If
    ElseIf RecibioMensaje(5) = True Then
            If VentanitaMostrar(5) = 0 Then
              VentanitaMostrar(5) = 1
            ElseIf VentanitaMostrar(5) = 1 Then
              VentanitaMostrar(5) = 0
            End If
    End If
    
    If logged And Not frmMain.Visible Then
        Unload frmAccount
        frmMain.Show
    End If
    
    If TiempoParalizado > 0 Then TiempoParalizado = TiempoParalizado - 1
    
    Static resetFlag As Byte
    resetFlag = resetFlag + 1
    
    If resetFlag = 120 Then
        toyMacreado = False
        resetFlag = 0
    End If
    
End Sub

Private Sub cmdGlobal_Click()
If GlobalConsole.Visible = False Then 'ta invisible la activamos
    RecTxt.Visible = False
    PrivatesConsole.Visible = False
    GlobalConsole.Visible = True
    ClanConsole.Visible = False
Else
    RecTxt.Visible = True
    PrivatesConsole.Visible = False
    GlobalConsole.Visible = False
    ClanConsole.Visible = False
End If
End Sub

Private Sub GMSOS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmGmPanelSOS.Show , frmMain
GMSOS.ForeColor = &HC0C0&
GMSOS.BackColor = &HFF&
End Sub

Private Sub GMSOS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
GMSOS.ForeColor = &HFFFF&
GMSOS.BackColor = &H80&
End Sub
Private Sub GMTORNEO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmTorneoManager.List1.Clear
Call SendData("TOINFO")
GMTORNEO.ForeColor = &HC0C0&
GMTORNEO.BackColor = &HFF&
End Sub
Private Sub GMTORNEO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
GMTORNEO.ForeColor = &HFFFF&
GMTORNEO.BackColor = &H80&
End Sub
Private Sub GMPANEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SendData("/TELEP YO 104 51 51")
GMPANEL.ForeColor = &HC0C0&
GMPANEL.BackColor = &HFF&
End Sub
Private Sub GMPANEL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
GMPANEL.ForeColor = &HFFFF&
GMPANEL.BackColor = &H80&
End Sub

Private Sub Image3_Click()

If TieneParaResponder = False Then
    frmGM.Show , frmMain
Else
    frmMensaje.Show , frmMain
    TieneParaResponder = False
End If

End Sub
Private Sub imgMenu_Click()
    PopupMenu mnuVista
End Sub

Private Sub imgMenuGral_Click()
    AddtoRichTextBox frmMain.RecTxt, "Este es el bot�n del men� general, haciendo doble doble click podr�s visualizar el formulario.", 255, 255, 255, False, False, False
End Sub
Private Sub imgMenuGral_dblClick()
    frmMenuGral.Show , frmMain
End Sub
Private Sub imgRanking_Click()
    AddtoRichTextBox frmMain.RecTxt, "Este es el bot�n del ranking, haciendo doble doble click podr�s visualizar el formulario.", 255, 255, 255, False, False, False
End Sub
Private Sub imgRanking_dblClick()
    frmRanking.Show , frmMain
End Sub
Private Sub kkkkkkkkkkkkk_Click(Index As Integer)

If Index = 0 Then Call Inventario.MoverItemInv(0)
If Index = 1 Then Call Inventario.MoverItemInv(1)
If Index = 2 Then Call Inventario.MoverItemInv(2)
If Index = 3 Then Call Inventario.MoverItemInv(3)

End Sub
Private Sub Label2_Click()
Call SendData("CCANJE")
Call SendData("ACTPT")
End Sub

Private Sub Image4_Click()
    frmSalir.Show vbModal, frmMain
End Sub

Private Sub Image5_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub lblMapaName_Click()
AddtoRichTextBox frmMain.RecTxt, "Este es el nombre del mapa en cual te encuentras.", 255, 255, 255, False, False, False
End Sub

Private Sub Label8_Click()
        SendData "FEST"
End Sub
Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub Minimap_Click()
    If MouseBoton = vbRightButton Then
        Call SendData("/TELEP YO " & UserMap & " " & minimapMouseX & " " & minimapMouseY)
    End If
End Sub
Private Sub Minimap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    minimapMouseX = X
    minimapMouseY = Y
End Sub
Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
End Sub
Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub
Private Sub mnuTirar_Click()
    Call TirarItem
End Sub
Private Sub mnuUsar_Click()
    Call UsarItem
End Sub
Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicaci�n en el mapa.", 255, 255, 255, False, False, False
End Sub

Private Sub cmdContacts_Click()
If ChatContacts.Visible = True Then 'si es visible ocultamos
    ChatContacts.Visible = False
    RecTxt.Width = 532
    ClanConsole.Width = 532
    GlobalConsole.Width = 532
    PrivatesConsole.Width = 532
    mnunormal_Click
Else
    ChatContacts.Visible = True
    RecTxt.Width = 394
    ClanConsole.Width = 394
    GlobalConsole.Width = 394
    PrivatesConsole.Width = 394
End If
End Sub
Private Sub cmdPrivados_Click()
If PrivatesConsole.Visible = False Then 'ta invisible la activamos
    RecTxt.Visible = False
    PrivatesConsole.Visible = True
    GlobalConsole.Visible = False
    ClanConsole.Visible = False
Else
    RecTxt.Visible = True
    PrivatesConsole.Visible = False
    GlobalConsole.Visible = False
    ClanConsole.Visible = False
End If
End Sub
Private Sub TirarItem()
If ISItem = True Then
Call AddtoRichTextBox(frmMain.RecTxt, "Desactiva el seguro de items primero con la tecla '*'", 255, 0, 0, False, False, False)

Exit Sub
Else
 If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
     If Inventario.Amount(Inventario.SelectedItem) = 1 Then
       SendData "TI" & Inventario.SelectedItem & "," & 1
                Else
              If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End If
End Sub
Private Sub AgarrarItem()
    SendData "AG"
End Sub
Private Sub UsarItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & Inventario.SelectedItem
End Sub
Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Configuracion.MoverPantalla = 0 Then Exit Sub
If PantallaCompleta = True Then Exit Sub
L = 1
UserStartX = X
UserStartY = Y
End Sub
Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If L = 1 Then
        frmMain.left = frmMain.left + (X - UserStartX)
        frmMain.top = frmMain.top + (Y - UserStartY)
        Drag.left = Drag.left + (X - UserStartX)
        Drag.top = Drag.top + (Y - UserStartY)
        RemDragX = Drag.left
        RemDragY = Drag.top
End If

End Sub
Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
L = 0
End Sub



''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''
Private Sub TrainingMacro_Timer()

    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    'Select Case SecuenciaMacroHechizos
        'Case 0
            If hlst.List(hlst.ListIndex) <> "(Nada)" And UserCanAttackMagia = 1 Then
                
                If UsingSkill = 0 Then
                    Call SendData("LH" & hlst.ListIndex + 1)
                    UsingSkill = Magia
                    frmMain.MousePointer = 2
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Else
                    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    UsingSkill = 0
                    UserCanAttackMagia = 0
                    frmMain.MousePointer = vbDefault
                End If
            
            End If
            'SecuenciaMacroHechizos = 1
        'Case 1
            'Call ConvertCPtoTP(renderer.left, renderer.top, MouseX, MouseY, tx, tY)
            'If (UsingSkill = Magia Or UsingSkill = Proyectiles) Then Exit Sub
            'SendData "WLC" & tX & "," & tY & "," & UsingSkill
            'UsingSkill = 0
            'SecuenciaMacroHechizos = 0
        'Case Else
            'DesactivarMacroHechizos
    'End Select
    
End Sub
Private Sub cmdLanzar_Click()

    On Error GoTo errh:

    If hlst.List(hlst.ListIndex) <> "(Nada)" And UserCanAttackMagia = 1 Then
        Static sActMX, sAntMX, sActMY, sAntMY As Integer
        sActMX = antMX
        sAntMX = tempMX
        sActMY = antMY
        sAntMY = tempMY
        
        If (sAntMX = antMX And sActMX = tempMX And sAntMY = tempMY And sAntMY = antMY) Or (antMX = tempMX And antMY = tempMY) Or ((tmpMouse(1).X - tmpMouse(2).X > 100) Or (tmpMouse(2).X - tmpMouse(1).X > 100)) Or ((tmpMouse(1).Y - tmpMouse(2).Y > 100) Or (tmpMouse(2).Y - tmpMouse(1).Y > 100)) Then
            cantDet = cantDet + 1
        End If
        
        If cantDet >= 4 Then
            Call SendData("TENGOMACROS")
            toyMacreado = True
            cantDet = 0
        End If
        
        If (toyMacreado) And ((tmpMouse(1).X - tmpMouse(2).X > 50) Or (tmpMouse(2).X - tmpMouse(1).X > 50)) Or ((tmpMouse(1).Y - tmpMouse(2).Y > 50) Or (tmpMouse(2).Y - tmpMouse(1).Y > 50)) Then Exit Sub
    
        UsingSkill = Magia
        
        If Configuracion.Cursor = 16 Then
            frmMain.MousePointer = 2
        Else
            frmMain.MousePointer = 99
            frmMain.MouseIcon = LoadPicture(App.Path & "\Data\GRAFICOS\Cursores\Cursor" & Configuracion.Cursor & ".ico")
        End If
        
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
    
        Call SendData("LH" & hlst.ListIndex + 1)
        UserCanAttackMagia = 0
        UsaMacro = True
    End If
    
    Exit Sub

errh:
    If hlst.List(hlst.ListIndex) <> "(Nada)" And UserCanAttackMagia = 1 Then
        If antMX = tempMX And antMY = tempMY Then
            cantDet = cantDet + 1
        End If
        
        If cantDet >= 4 Then
            Call SendData("TENGOMACROS")
            cantDet = 0
        End If
        
        UsingSkill = Magia
        frmMain.MousePointer = 2
        
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
    
        Call SendData("LH" & hlst.ListIndex + 1)
        UserCanAttackMagia = 0
        UsaMacro = True
    End If

    
End Sub
Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    antMX = tempMX
    antMY = tempMY

    tempMX = X
    tempMY = Y
    
    UsaMacro = False
    CnTd = 0
End Sub
Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.ListIndex + 1)
End Sub
Private Sub ImgSkills_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain
End Sub
Private Sub Form_Click()

    If Cartel Then Cartel = False
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

        If MouseShift <> 1 Then
            If MouseBoton = vbRightButton Then
                If Configuracion.DobleClick = 1 Then
                    'MsgBox "Ese boton no tiene uso."
                    SendData "LC" & tX & "," & tY
                    SendData "RC" & tX & "," & tY
                End If
                
                    If ChatEnUso(1) = True And (MouseX >= 103 And MouseX <= 185) And (MouseY > 394 And MouseY < 408) Then
                        ChatEnUso(1) = False
                    ElseIf ChatEnUso(2) = True And (MouseX >= 188 And MouseX <= 270) And (MouseY > 394 And MouseY < 408) Then
                        ChatEnUso(2) = False
                    ElseIf ChatEnUso(3) = True And (MouseX >= 273 And MouseX <= 350) And (MouseY > 394 And MouseY < 408) Then
                        ChatEnUso(3) = False
                    ElseIf ChatEnUso(4) = True And (MouseX >= 358 And MouseX <= 440) And (MouseY > 394 And MouseY < 408) Then
                        ChatEnUso(4) = False
                    ElseIf ChatEnUso(5) = True And (MouseX >= 443 And MouseX <= 525) And (MouseY > 394 And MouseY < 408) Then
                        ChatEnUso(5) = False
                    End If
            End If
        
            If MouseBoton <> vbRightButton Then
                If ChatEnUso(1) = True And (MouseX >= 103 And MouseX <= 185) And (MouseY > 394 And MouseY < 408) Then
                    ChatForm(1).lblName = UCase$(NickContacto(1))
                    VentanitaMostrar(1) = 2
                    RecibioMensaje(1) = False
                    ChatForm(1).Show , frmMain
                ElseIf ChatEnUso(2) = True And (MouseX >= 188 And MouseX <= 270) And (MouseY > 394 And MouseY < 408) Then
                    ChatForm(2).lblName = UCase$(NickContacto(2))
                    VentanitaMostrar(2) = 2
                    RecibioMensaje(2) = False
                    ChatForm(2).Show , frmMain
                ElseIf ChatEnUso(3) = True And (MouseX >= 273 And MouseX <= 350) And (MouseY > 394 And MouseY < 408) Then
                    ChatForm(3).lblName = UCase$(NickContacto(3))
                    VentanitaMostrar(3) = 2
                    RecibioMensaje(3) = False
                    ChatForm(3).Show , frmMain
                ElseIf ChatEnUso(4) = True And (MouseX >= 358 And MouseX <= 440) And (MouseY > 394 And MouseY < 408) Then
                    ChatForm(4).lblName = UCase$(NickContacto(4))
                    VentanitaMostrar(4) = 2
                    RecibioMensaje(4) = False
                    ChatForm(4).Show , frmMain
                ElseIf ChatEnUso(5) = True And (MouseX >= 443 And MouseX <= 525) And (MouseY > 394 And MouseY < 408) Then
                    ChatForm(5).lblName = UCase$(NickContacto(5))
                    VentanitaMostrar(5) = 2
                    RecibioMensaje(5) = False
                    ChatForm(5).Show , frmMain
                End If
                
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If TrainingMacro.Enabled Then DesactivarMacroHechizos
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    UsingSkill = 0
                End If
                
            End If
        ElseIf (MouseShift) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
    
End Sub
Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandler

If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
    
Dim i As Integer, A As Byte
    
    For i = 1 To NUMBINDS
        If KeyCode = BindKeys(i).KeyCode Then
            A = 1
        End If
    Next i
    
    If KeyCode = vbKeyEscape Then
        frmSalir.Show vbModal, frmMain
        Exit Sub
    End If
    
    
    Select Case KeyCode
            Case vbKeyF9:
                frmMakro.Show , frmMain
            
                Case vbKey1:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla0"))
                    
                Case vbKey2:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla1"))
                    
                Case vbKey3:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla2"))
                    
                Case vbKey4:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla3"))
                    
                Case vbKey5:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla4"))
                    
                Case vbKey6:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla5"))
                    
                Case vbKey7:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla6"))
                    
                Case vbKey8:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla7"))
                    
                Case vbKey9:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla8"))
                    
                Case vbKey0:
                    enviarMacro (GetVar(App.Path & "\Data\INIT\Macro.tsao", "Macro", "Tecla9"))
    End Select

        
        Select Case KeyCode
        
            Case vbKeyNumpad0 To vbKeyNumpad8:
                If Configuracion.HablaNumerico = 0 Then Exit Sub
                If SendTxt.Visible = False And frmBancoObj.Visible = False And frmCantidad.Visible = False And frmNuevoComercio.Visible = False Then
                    Call TalkMode(KeyCode - 96)
               End If
                
                Case BindKeys(19).KeyCode
                    frmEmoticons.Show , frmMain
                
                Case BindKeys(20).KeyCode
                    If Musica = False Then
                        Musica = True
                        Audio.MP3_Play CurrentMP3
                    Else
                        Musica = False
                        Audio.MP3_Stop
                        Audio.MP3_Destroy
                    End If
                
                Case BindKeys(2).KeyCode:
                    Call AgarrarItem
                
                Case BindKeys(12).KeyCode
          
                Case BindKeys(5).KeyCode
                    Call EquiparItem
                    
                Case BindKeys(18).KeyCode:
                    Call frmMapa.Show(vbModeless, frmMain)
                
                Case BindKeys(7).KeyCode
                    Nombres = Not Nombres
                
                Case BindKeys(9).KeyCode
                    Call SendData("UK" & Robar)
                            
                Case BindKeys(11).KeyCode
                    Call SendData("UK" & Ocultarse)
                
                Case BindKeys(3).KeyCode
                    Call TirarItem
                
                Case BindKeys(4).KeyCode
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                
                Case BindKeys(10).KeyCode
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
                    
                Case BindKeys(21).KeyCode
                    If ISItem = True Then
                        ISItem = False
                        Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO DE ITEMS DESACTIVADO<<", 255, 0, 0, True, False, False)
                        SegurodeItems.Enabled = True
                    Else
                        ISItem = True
                        Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO DE ITEMS ACTIVADO<<", 0, 255, 0, True, False, False)
                        SegurodeItems.Enabled = False
                    End If
                
                Case BindKeys(13).KeyCode
                DoEvents
                Call keybd_event(VK_SNAPSHOT, PS_TheScreen, 0, 0)
                DoEvents
                For i = 1 To 1000
                       If Not FileExist(App.Path & "\Data\SCREENSHOTS\Screen" & i & ".jpg", vbNormal) Then Exit For
                Next
                
                SavePicture Clipboard.GetData, App.Path & "\Data\SCREENSHOTS\Screen" & i & ".jpg"
                Call AddtoRichTextBox(frmMain.RecTxt, "�La Scren Shot fue guardada en la carpeta Screen Shots del Cliente con el nombre Screen" & i & ".jpg !", 0, 0, 255, False, True, False)
                
                Case BindKeys(8).KeyCode
                    Call SendData("/SEGR")
                    
                Case vbKeyF7
                 If frmMain.WorkMacro.Enabled = True Then
                    frmMain.WorkMacro.Enabled = False
                Else
                    frmMain.WorkMacro.Enabled = True
                End If
                
                Case BindKeys(6).KeyCode
                    Call SendData("/SEG")
            End Select
        Else
 
        End If
    
    Select Case KeyCode
        Case BindKeys(14).KeyCode
            quitKey (KeyCode)
        Case BindKeys(15).KeyCode
            quitKey (KeyCode)
        Case BindKeys(16).KeyCode
            quitKey (KeyCode)
        Case BindKeys(17).KeyCode
            quitKey (KeyCode)
            
        Case vbKeyF5
            'Call OpcionesNew.Show(vbModeless, frmMain)
            frmMenuGral.Show , frmMain
        
        Case vbKeyF6
            Call SendData("/MEDITAR")
        
        Case vbKeyF7:
                If TrainingMacro.Enabled Then
                    DesactivarMacroHechizos
                Else
                    ActivarMacroHechizos
                End If
            
        Case BindKeys(1).KeyCode
            If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
            End If
        
        Case vbKeyReturn:
             For i = 1 To 5
                If ChatForm(i).Visible = True Then Exit Sub
             Next i
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                SendTxt.SetFocus
                End If
    End Select
    
    Exit Sub
    
errHandler:
     MsgBox ": " & err.Number & " : " & err.Description, vbOKOnly, "Error"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case BindKeys(14).KeyCode
            AddKey (KeyCode)
        Case BindKeys(15).KeyCode
            AddKey (KeyCode)
        Case BindKeys(16).KeyCode
            AddKey (KeyCode)
        Case BindKeys(17).KeyCode
            AddKey (KeyCode)
    End Select

End Sub
Private Function indexKey(key As Integer) As Byte

    Dim i As Long, tmpI As Byte
    tmpI = 0
    For i = 1 To 4
        If dKeys(i) = key Then
            tmpI = i
        End If
    Next i
    
    indexKey = tmpI

End Function
Private Sub quitKey(key As Integer)
    Dim i As Long, tmpI As Byte
    tmpI = indexKey(key)
    
    If (tmpI = 4) Then dKeys(4) = 0: Exit Sub
    If (tmpI = 0) Then Exit Sub
    
    For i = tmpI To 3
        dKeys(i) = dKeys(i + 1)
    Next i
    
End Sub
Private Sub AddKey(key As Integer)
    Dim i As Long, tmpI As Byte
    tmpI = indexKey(key)
    
    If (tmpI > 1) Then quitKey (key)
    If (tmpI = 1) Then Exit Sub
    
    dKeys(4) = dKeys(3)
    dKeys(3) = dKeys(2)
    dKeys(2) = dKeys(1)
    dKeys(1) = key
    
End Sub
Private Sub Form_Load()

    frmMain.ExpBar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Principal_expBar.jpg")
    
    Picture = General_Load_Interface_Picture("principal.jpg")
    InvEqu.Picture = General_Load_Interface_Picture("Centronuevoinventario.jpg")
    
            ImgEstadisticas.Enabled = False
            ImgEquipo.Enabled = False
            ImgSkills.Enabled = False
            ImgCastillos.Enabled = False
            ImgPremios.Enabled = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
            DyDActivado = False
            
            Detectar RecTxt.hWnd, Me.hWnd
            Detectar GlobalConsole.hWnd, Me.hWnd
            Detectar PrivatesConsole.hWnd, Me.hWnd
            Detectar ClanConsole.hWnd, Me.hWnd
        
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y

If DyDActivado = False Then DyD.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\D&D_N.jpg")
    
End Sub
Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub
Private Sub Label4_Click()

    Call Audio.PlayWave("click.wav")

    InvEqu.Picture = General_Load_Interface_Picture("Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    DyD.Visible = True
    
    ItemName.Visible = True
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
End Sub

Private Sub Label7_Click()
    
    Call Audio.PlayWave("click.wav")

    InvEqu.Picture = General_Load_Interface_Picture("Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True

    DyD.Visible = False
    ItemName.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub
Private Sub picInv_DblClick()
    If InvMouseBoton = 0 Then Exit Sub
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    If (Inventario.SelectedItem > 0) And (Not NoPuedeUsarClick) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "QSA" & Inventario.SelectedItem & "," & frmMain.picInv.Visible
     InvMouseBoton = 0
End Sub
Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InvMouseBoton = Button
    If Not GetAsyncKeyState(Button) < 0 Then InvMouseBoton = 0
 
 If DyDActivado = True Then
    'Inventario.InvSelectedItem = Inventario.ClickItem(X, Y)
    DibujadoContinuoInv = True
 End If

End Sub
Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not DyDActivado Then Exit Sub
 
MouseXInv = X
MouseYInv = Y

If MouseRendOK = True Then MouseRendOK = False
 
End Sub
Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave("click.wav")
    
  If DyDActivado = True Then
    DibujadoContinuoInv = False
  End If
  
End Sub
Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And (Not frmMakro.Visible) And (Not frmMenuGral.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmCantidad.Visible) And frmMain.WindowState = vbMaximized And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
      If ImgEquipo.Enabled = False Then
        hlst.SetFocus
      End If
    End If
End Sub
Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - imped� se inserten caract�res no imprimibles
'**************************************************************
    If Len(SendTxt.text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.text)
            CharAscii = Asc(mid$(SendTxt.text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.text = tempstr
        End If
        
        stxtbuffer = SendTxt.text
    End If
End Sub
Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
    
        If left$(stxtbuffer, 1) = "/" Then
            If UCase(left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
                    j = (Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    stxtbuffer = "/PASSWD " & j
            ElseIf UCase$(stxtbuffer) = "/GM" Then
                If TieneParaResponder = False Then
                 frmGM.Show , frmMain
                Else
                 frmMensaje.Show , frmMain
                 TieneParaResponder = False
                End If
                
                stxtbuffer = ""
                SendTxt.text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/GEMAS" Then
                frmMenuGral.Show , frmMain
                frmMenuGral.mostrarGem_Med (True)
                stxtbuffer = ""
                SendTxt.text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/MEDALLAS" Then
                frmMenuGral.Show , frmMain
                frmMenuGral.mostrarGem_Med (True)
                stxtbuffer = ""
                SendTxt.text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/DONE" Then
                frmMenuGral.Show , frmMain
                Call SendData("DCANJE")
                stxtbuffer = ""
                SendTxt.text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            End If
                
            
            Call SendData(stxtbuffer)
            
     Else
        Call SendData(modoHabla & stxtbuffer, True)
    End If

        stxtbuffer = ""
        SendTxt.text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.text
    End If
End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''

Private Sub Socket1_Connect()

On Error Resume Next
   
  If EstadoLogin = E_MODO.CrearNuevoPj Then
    Call Login
 
  ElseIf EstadoLogin = E_MODO.Normal Then
    Call Login
 
  ElseIf EstadoLogin = E_MODO.Dados Then
    frmCrearPersonaje.Show , frmConnect
   
  ElseIf EstadoLogin = E_MODO.CrearAccount Then
    frmConnect.mostrarNuevaCuenta (True)
      
  ElseIf EstadoLogin = E_MODO.RecuPW Then
    frmConnect.mostrarRecuperarCuenta (True)
   
  ElseIf EstadoLogin = E_MODO.LoginAccount Then
    Call Login
   
  ElseIf EstadoLogin = E_MODO.BorrarPj Then
    Call Login
  End If
End Sub
Private Sub Socket1_Disconnect()
    Dim i As Long

    logged = False
    Connected = False
    AoDefResult = 0
   
    Socket1.Cleanup
   
    frmConnect.MousePointer = vbNormal
    frmCrearPersonaje.Visible = False
    Call frmConnect.mostrarConectar(True, False, True)
   
    frmMain.Visible = False
    pausa = False
    UserMeditar = False
 
    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
   
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i
 
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
 
    SkillPoints = 0
    Alocados = 0
End Sub
 
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Mensaje.Escribir "Por favor espere, intentando completar conexion."
        Exit Sub
    End If
    
    Mensaje.Escribir "Conexi�n rechazada por el Servidor"
    
    frmConnect.MousePointer = 1
    Response = 0
    frmMain.Socket1.Disconnect
End Sub
 

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, dataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub WorkMacro_Timer()

'If Me.ItemName.Caption = "Hacha de Le�ador" Or Me.ItemName.Caption = "Piquete de Minero" Or Me.ItemName.Caption = "Ca�a de Pescar" Then
'    SendData "KLQ" & Inventario.SelectedItem
'    SendData "WLC" & tX & "," & tY & "," & UsingSkill
'Else
'    AddtoRichTextBox frmMain.RecTxt, "Este item no puede ser utilizado con el macro de trabajo!", 255, 255, 255
'    frmMain.WorkMacro.Enabled = False
'    AddtoRichTextBox frmMain.RecTxt, "Macro de trabajo desactivado.", 255, 255, 255
'    Exit Sub
'End If

End Sub
Private Sub mnunormal_Click()

modoHabla = ";"
mnuNormal.Checked = True
mnuprivado.Checked = False
mnudenunciar.Checked = False
mnuClanes.Checked = False
mnuParty.Checked = False
mnuGritar.Checked = False
mnuGlo.Checked = False
mnuFaccion.Checked = False
'mnuchat.Checked = False
'ClanConsole.Visible = False
'GlobalConsole.Visible = False
'PrivatesConsole.Visible = False
RecTxt.Visible = True
End Sub

Private Sub mnuparty_Click()
modoHabla = "/pmsg "
'hablar.Caption = "P"
mnuNormal.Checked = False
mnuprivado.Checked = False
mnudenunciar.Checked = False
mnuClanes.Checked = False
mnuParty.Checked = True
mnuGritar.Checked = False
mnuGlo.Checked = False
mnuFaccion.Checked = False
'mnuchat.Checked = False
End Sub
Private Sub mnuprivado_Click()
Dim Usuario As String
TheUser = InputBox("Escriba el nombre del destinatario del mensaje", "Mensaje Privado")
modoHabla = "\" & TheUser & "@"
mnuNormal.Checked = False
mnuprivado.Checked = True
mnudenunciar.Checked = False
mnuClanes.Checked = False
mnuParty.Checked = False
mnuGritar.Checked = False
mnuGlo.Checked = False
mnuFaccion.Checked = False
'mnuchat.Checked = False
'hablar.Caption = "P"
End Sub

Private Sub mnuchat_Click()
'MSNImage.Visible = False
'RecTxt.Visible = False
''ChatBox.Visible = True
'If ChatContacts.Visible = True Then
'    ChatContacts.Visible = False
'    RecTxt.Width = 520
'Else
'    ChatContacts.Visible = True
'    RecTxt.Width = 392
'End If
'modohabla.Text = "*" & ChatingWith & ","
'mnuNormal.Checked = False
'mnuprivado.Checked = False
'mnudenunciar.Checked = False
'mnuClanes.Checked = False
'mnuParty.Checked = False
'mnuGritar.Checked = False
'mnuGlo.Checked = False
'mnuFaccion.Checked = False
'Contact(ChatingWith).Name = ""
'mnuchat.Checked = True
End Sub

Private Sub mnuclanes_Click()
modoHabla = "/cmsg "
'hablar.Caption = "C"
mnuNormal.Checked = False
mnuprivado.Checked = False
mnudenunciar.Checked = False
mnuClanes.Checked = True
mnuParty.Checked = False
mnuGritar.Checked = False
mnuGlo.Checked = False
mnuFaccion.Checked = False
'mnuchat.Checked = False
End Sub

Private Sub mnudenunciar_Click()
modoHabla = "/gmsg "
'hablar.Caption = "D"
mnuNormal.Checked = False
mnuprivado.Checked = False
mnudenunciar.Checked = True
mnuClanes.Checked = False
mnuParty.Checked = False
mnuGritar.Checked = False
mnuGlo.Checked = False
mnuFaccion.Checked = False
'mnuchat.Checked = False
End Sub

Private Sub mnuFaccion_Click()
modoHabla = "/FMSG "
mnuNormal.Checked = False
mnuprivado.Checked = False
mnudenunciar.Checked = False
mnuClanes.Checked = False
mnuParty.Checked = False
mnuGritar.Checked = False
mnuGlo.Checked = False
mnuFaccion.Checked = True
'mnuchat.Checked = False
End Sub

Private Sub mnuGlo_Click()
modoHabla = "/GLOBAL "
mnuNormal.Checked = False
mnuprivado.Checked = False
mnudenunciar.Checked = False
mnuClanes.Checked = False
mnuParty.Checked = False
mnuGritar.Checked = False
mnuGlo.Checked = True
mnuFaccion.Checked = False
'mnuchat.Checked = False
End Sub

Private Sub mnugritar_Click()
modoHabla = "-"
'hablar.Caption = "G"
mnuNormal.Checked = False
mnuprivado.Checked = False
mnudenunciar.Checked = False
mnuClanes.Checked = False
mnuParty.Checked = False
mnuGritar.Checked = True
mnuGlo.Checked = False
mnuFaccion.Checked = False
'mnuchat.Checked = False
End Sub
Sub TalkMode(ByVal Modo As Integer)
Select Case Modo
    Case 0:
        mnunormal_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Normal", 255, 255, 255, False, False, False)
    Case 1:
        mnugritar_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Gritar", 255, 255, 255, False, False, False)
    Case 2:
        mnuclanes_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Clan", 255, 255, 255, False, False, False)
    Case 3:
        mnuGlo_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Global", 255, 255, 255, False, False, False)
    Case 4:
        mnuparty_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Party", 255, 255, 255, False, False, False)
    Case 5:
        mnuFaccion_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Faccion", 255, 255, 255, False, False, False)
    Case 6:
        mnudenunciar_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Denunciar", 255, 255, 255, False, False, False)
    Case 7:
        mnuprivado_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Privado", 255, 255, 255, False, False, False)
End Select
End Sub
Private Sub SegurodeItems_Timer()
Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO DE ITEMS ACTIVADO<<", 0, 255, 0, True, False, False)
SegurodeItems.Enabled = False
ISItem = True
End Sub
Private Sub renderer_Click()
Call Form_Click
End Sub

Private Sub renderer_DblClick()
Call Form_DblClick
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub
Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseX = X
    MouseY = Y

    MouseRenderX = X
    MouseRenderY = Y
    
    If DyDActivado = True And MouseOK = True Then
        MouseRendOK = True
        MouseOK = False
    End If
    
If (MouseX >= 103 And MouseX <= 185) And (MouseY > 394 And MouseY < 408) Then
    MouseBarraChat(1) = True
Else
    MouseBarraChat(1) = False
End If

If (MouseX >= 188 And MouseX <= 270) And (MouseY > 394 And MouseY < 408) Then
    MouseBarraChat(2) = True
Else
    MouseBarraChat(2) = False
End If

If (MouseX >= 273 And MouseX <= 350) And (MouseY > 394 And MouseY < 408) Then
    MouseBarraChat(3) = True
Else
    MouseBarraChat(3) = False
End If

If (MouseX >= 358 And MouseX <= 440) And (MouseY > 394 And MouseY < 408) Then
    MouseBarraChat(4) = True
Else
    MouseBarraChat(4) = False
End If

If (MouseX >= 443 And MouseX <= 525) And (MouseY > 394 And MouseY < 408) Then
    MouseBarraChat(5) = True
Else
    MouseBarraChat(5) = False
End If
    
End Sub
Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
    
    If DyDActivado = True And MouseRendOK = True Then
        MouseRendOK = False
        TirarItemMouse
        ButtonIN = False
        PUEDO = False
        MouseOK = False
        MouseItem = 0
    End If
    
    ButtonIN = False
    MouseRendOK = False
    MouseOK = False
    MouseItem = 0
    
End Sub
