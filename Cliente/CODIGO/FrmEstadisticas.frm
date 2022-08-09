VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   7395
   ClientLeft      =   4215
   ClientTop       =   1635
   ClientWidth     =   5625
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOBJ 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picNPC 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1560
      Picture         =   "FrmEstadisticas.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   47
      Top             =   3255
      Width           =   495
   End
   Begin VB.PictureBox picNPC 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2520
      Picture         =   "FrmEstadisticas.frx":0F4C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   46
      Top             =   3255
      Width           =   495
   End
   Begin VB.PictureBox picNPC 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   3480
      Picture         =   "FrmEstadisticas.frx":1B8E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   45
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtQuestDescription 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1575
      Left            =   285
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblCantOBJ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x5"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   54
      Top             =   4455
      Width           =   495
   End
   Begin VB.Label lblPTS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2.500"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   52
      Top             =   4545
      Width           =   615
   End
   Begin VB.Label lblORO 
      BackStyle       =   0  'Transparent
      Caption         =   "10.000.000"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   51
      Top             =   4350
      Width           =   975
   End
   Begin VB.Label lblCantNPC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "01/10"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   50
      Top             =   3750
      Width           =   495
   End
   Begin VB.Label lblCantNPC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1/5"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   49
      Top             =   3750
      Width           =   495
   End
   Begin VB.Label lblCantNPC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/10"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   48
      Top             =   3750
      Width           =   495
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   20
      Left            =   4440
      TabIndex        =   44
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5160
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   25
      Left            =   2160
      TabIndex        =   43
      Top             =   5535
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   23
      Left            =   2400
      TabIndex        =   42
      Top             =   4860
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   24
      Left            =   2880
      TabIndex        =   41
      Top             =   5205
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   22
      Left            =   1320
      TabIndex        =   40
      Top             =   4515
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   21
      Left            =   2160
      TabIndex        =   39
      Top             =   4185
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   2400
      TabIndex        =   38
      Top             =   2640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   2520
      TabIndex        =   37
      Top             =   2370
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   2520
      TabIndex        =   36
      Top             =   2085
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   16
      Left            =   1440
      TabIndex        =   35
      Top             =   1770
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "1723"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image imgHoja 
      Height          =   450
      Index           =   2
      Left            =   4920
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHoja 
      Height          =   405
      Index           =   1
      Left            =   4920
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   32
      Top             =   4860
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   31
      Top             =   4500
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   2235
      TabIndex        =   30
      Top             =   5865
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   28
      Top             =   5175
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   27
      Top             =   4845
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   26
      Top             =   4515
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   25
      Top             =   4185
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   24
      Top             =   3090
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   21
      Top             =   2070
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   20
      Top             =   1755
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCounters 
      BackStyle       =   0  'Transparent
      Caption         =   "172"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image imgQuestAbandonar 
      Height          =   540
      Left            =   330
      Top             =   5400
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Label lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   5
      Left            =   4170
      TabIndex        =   17
      Top             =   5130
      Width           =   615
   End
   Begin VB.Label lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   1770
      TabIndex        =   16
      Top             =   5130
      Width           =   615
   End
   Begin VB.Label lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   4845
      TabIndex        =   15
      Top             =   4755
      Width           =   615
   End
   Begin VB.Label lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   2745
      TabIndex        =   14
      Top             =   4755
      Width           =   495
   End
   Begin VB.Label lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Label lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   6000
      Width           =   4575
   End
   Begin VB.Label lblBonificadores 
      BackStyle       =   0  'Transparent
      Caption         =   "No elegido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   11
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Label lblAtri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   975
      TabIndex        =   10
      Top             =   4755
      Width           =   600
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Prueba"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   675
      Width           =   3375
   End
   Begin VB.Label lblPuntosDonador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3600
      MouseIcon       =   "FrmEstadisticas.frx":27D0
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label lblPuntosTorneo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   900
      MouseIcon       =   "FrmEstadisticas.frx":2ADA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label lblMail 
      BackStyle       =   0  'Transparent
      Caption         =   "a@a.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2850
      Width           =   2535
   End
   Begin VB.Label lblHogar 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanaris"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2475
      Width           =   2175
   End
   Begin VB.Label lblGenero 
      BackStyle       =   0  'Transparent
      Caption         =   "Hombre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1395
      Width           =   1935
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "Mago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1350
      TabIndex        =   3
      Top             =   1755
      Width           =   2295
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo Oscuro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1245
      TabIndex        =   2
      Top             =   1050
      Width           =   2295
   End
   Begin VB.Label lblReputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "104.250"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgExtras 
      Height          =   630
      Left            =   3690
      Picture         =   "FrmEstadisticas.frx":2DE4
      Top             =   6555
      Width           =   1440
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "50 + 10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   2145
      Width           =   735
   End
   Begin VB.Image imgQuests 
      Height          =   630
      Left            =   2100
      Picture         =   "FrmEstadisticas.frx":7BE0
      Top             =   6555
      Width           =   1440
   End
   Begin VB.Image imgGeneral 
      Height          =   630
      Left            =   525
      Picture         =   "FrmEstadisticas.frx":CBFE
      Top             =   6555
      Width           =   1440
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Form_Load()
Set form_Moviment = New clsFormMovementManager
form_Moviment.Initialize Me

Me.Picture = General_Load_Interface_Picture("Estadisticas_1_Main.jpg")
imgGeneral.Picture = General_Load_Interface_Picture("Estadisticas_1_Principal_N.jpg")
'imgHabilidades.Picture = General_Load_Interface_Picture("Estadisticas_1_Habilidades_N.jpg")
imgQuests.Picture = General_Load_Interface_Picture("Estadisticas_1_Quest_N.jpg")
imgExtras.Picture = General_Load_Interface_Picture("Estadisticas_1_Extras_N.jpg")

'Ponemos visible toda la primera pagina.
lblNombre.Visible = True
lblLvl.Visible = True
lblRaza.Visible = True
lblGenero.Visible = True
lblHogar.Visible = True
lblMail.Visible = True
lblClase.Visible = True
lblReputacion.Visible = True
lblPuntosTorneo.Visible = True
lblPuntosDonador.Visible = True
lblBonificadores(1).Visible = True
lblBonificadores(2).Visible = True
lblBonificadores(3).Visible = True
'picInv.Visible = True
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = True
Next

'Vaciamos toda la pagina "extra" y "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

    lblPTS.Visible = False
    lblOro.Visible = False
    txtQuestDescription.Visible = False
    imgQuestAbandonar.Visible = False
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    picOBJ.Visible = False
    lblCantOBJ.Visible = False

Call Iniciar_Labels
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Volvemos las imagenes a la normalidad
imgGeneral.Picture = General_Load_Interface_Picture("Estadisticas_1_Principal_N.jpg")
'imgHabilidades.Picture = General_Load_Interface_Picture("Estadisticas_1_Habilidades_N.jpg")
imgQuests.Picture = General_Load_Interface_Picture("Estadisticas_1_Quest_N.jpg")
imgExtras.Picture = General_Load_Interface_Picture("Estadisticas_1_Extras_N.jpg")
imgQuestAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_2_AbandonarQuest_N.jpg")

End Sub
Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills

For i = 1 To NUMATRIBUTOS
    lblAtri(i).Caption = UserAtributos(i)
Next

'### PRIMERA HOJA
lblNombre.Caption = UserEstadisticas.Nombre

lblLvl.Caption = UserEstadisticas.Nivel

lblRaza.Caption = UserEstadisticas.Raza
lblClase.Caption = UserEstadisticas.Clase
lblGenero.Caption = UserEstadisticas.Genero
lblHogar.Caption = UserEstadisticas.Hogar
lblMail.Caption = UserEstadisticas.Email
lblPuntosTorneo.Caption = UserEstadisticas.PuntosTorneo
lblPuntosDonador.Caption = PonerPuntos(UserEstadisticas.PuntosDonador)
lblBonificadores(1).Caption = UserEstadisticas.Bonif1
lblBonificadores(2).Caption = UserEstadisticas.Bonif2
lblBonificadores(3).Caption = UserEstadisticas.Bonif3
lblReputacion.Caption = UserEstadisticas.UserReputacion

'### EXTRAS - HOJA 1
lblCounters(0).Caption = UserEstadisticas.TorneosParticipados
lblCounters(1).Caption = "0"
lblCounters(2).Caption = "0"
lblCounters(3).Caption = UserEstadisticas.CopasDeOro
lblCounters(4).Caption = UserEstadisticas.CopasDePlata
lblCounters(5).Caption = UserEstadisticas.CopasDeBronce
lblCounters(6).Caption = UserEstadisticas.Eventos
lblCounters(7).Caption = UserEstadisticas.DuelosGanados
lblCounters(8).Caption = UserEstadisticas.DuelosGanados + UserEstadisticas.DuelosPerdidos
lblCounters(9).Caption = UserEstadisticas.ParejasGanadas
lblCounters(10).Caption = UserEstadisticas.ParejasGanadas + UserEstadisticas.ParejasPerdidas
lblCounters(11).Caption = UserEstadisticas.CvcsGanados
lblCounters(12).Caption = UserEstadisticas.MaximasRondas
lblCounters(13).Caption = UserEstadisticas.GuerrasGanadas
lblCounters(14).Caption = UserEstadisticas.GuerrasGanadas + UserEstadisticas.GuerrasPerdidas

'### EXTRAS - HOJA 2
If UserEstadisticas.Alineacion = 1 Then
    lblCounters(15).ForeColor = &H80&
    lblCounters(15) = "HORDA INFERNAL"
ElseIf UserEstadisticas.Alineacion = 2 Then
    lblCounters(15).ForeColor = &HC00000
    lblCounters(15) = "ALIANZA IMPERIAL"
ElseIf UserEstadisticas.Alineacion = 0 Then
    lblCounters(15).ForeColor = &H404040
    lblCounters(15) = "NEUTRAL"
End If

lblCounters(16).Caption = UserEstadisticas.Jerarquia
lblCounters(17).Caption = UserEstadisticas.CiudadanosMatados
lblCounters(18).Caption = UserEstadisticas.NeutralesMatados
lblCounters(19).Caption = UserEstadisticas.CriminalesMatados
lblCounters(20).Caption = UserEstadisticas.Restantes
lblCounters(21).Caption = UserEstadisticas.NPCSMATADOS
lblCounters(22).Caption = UserEstadisticas.MuertesUsuario
lblCounters(23).Caption = UserEstadisticas.QuestCompletadas
lblCounters(24).Caption = "0"
lblCounters(25).Caption = UserEstadisticas.MVPMatados

'### HOJA QUEST
imgQuestAbandonar.Picture = General_Load_Interface_Picture("Estadisticas_2_AbandonarQuest_N.jpg")
 
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub imgGeneral_Click()
Me.Picture = General_Load_Interface_Picture("Estadisticas_1_Main.jpg")

'Ponemos visible toda la primera pagina.
lblNombre.Visible = True
lblLvl.Visible = True
lblRaza.Visible = True
lblClase.Visible = True
lblGenero.Visible = True
lblHogar.Visible = True
lblMail.Visible = True
lblReputacion.Visible = True
lblPuntosTorneo.Visible = True
lblPuntosDonador.Visible = True
lblBonificadores(1).Visible = True
lblBonificadores(2).Visible = True
lblBonificadores(3).Visible = True
'picInv.Visible = True
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = True
Next

'Vaciamos toda la pagina "extra" y "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

'Vaciamos la pagina "quest"
    lblPTS.Visible = False
    lblOro.Visible = False
    txtQuestDescription.Visible = False
    imgQuestAbandonar.Visible = False
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    picOBJ.Visible = False
    lblCantOBJ.Visible = False
End Sub
Public Sub InformarQuests(ByVal Data As String)

    Dim NroQuest As Byte
    Dim MuerteNPC(1 To 3) As Byte
    NroQuest = ReadField(1, Data, 44)
    
    MuerteNPC(1) = ReadField(2, Data, 44)
    MuerteNPC(2) = ReadField(3, Data, 44)
    MuerteNPC(3) = ReadField(4, Data, 44)
    
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    
    If NroQuest = 0 Then
        txtQuestDescription.text = "No estás haciendo ninguna quest."
    Exit Sub
    End If

    txtQuestDescription.text = "Nombre: " & InfoQuests(NroQuest).Nombre & vbCrLf & "Información: " & InfoQuests(NroQuest).Info & vbCrLf & "Dificultad: " & InfoQuests(NroQuest).Dificultad & vbCrLf
    
    'invisibles
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    
    If InfoQuests(NroQuest).NPCs = 1 Then
        picNPC(0).Visible = True
        lblCantNPC(0).Visible = True
        picNPC(0).left = 168
        lblCantNPC(0).left = 168
        
        picNPC(0).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(NroQuest).NumNPC(1) & ".jpg")
        lblCantNPC(0).Caption = MuerteNPC(1) & "/" & InfoQuests(NroQuest).CantNPC(1)
    ElseIf InfoQuests(NroQuest).NPCs = 2 Then
        picNPC(0).Visible = True
        lblCantNPC(0).Visible = True
        picNPC(1).Visible = True
        lblCantNPC(1).Visible = True
        
        picNPC(0).left = 144
        lblCantNPC(0).left = 144
        picNPC(1).left = 200
        lblCantNPC(1).left = 200
        
        picNPC(0).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(NroQuest).NumNPC(1) & ".jpg")
        lblCantNPC(0).Caption = MuerteNPC(1) & "/" & InfoQuests(NroQuest).CantNPC(1)
        
        picNPC(1).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(NroQuest).NumNPC(2) & ".jpg")
        lblCantNPC(1).Caption = MuerteNPC(2) & "/" & InfoQuests(NroQuest).CantNPC(2)
    ElseIf InfoQuests(NroQuest).NPCs = 3 Then
        picNPC(0).Visible = True
        lblCantNPC(0).Visible = True
        picNPC(1).Visible = True
        lblCantNPC(1).Visible = True
        picNPC(2).Visible = True
        lblCantNPC(2).Visible = True
        
        picNPC(0).left = 104
        lblCantNPC(0).left = 104
        picNPC(1).left = 168
        lblCantNPC(1).left = 168
        picNPC(2).left = 232
        lblCantNPC(2).left = 232
        
        picNPC(0).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(NroQuest).NumNPC(1) & ".jpg")
        lblCantNPC(0).Caption = MuerteNPC(1) & "/" & InfoQuests(NroQuest).CantNPC(1)
        
        picNPC(1).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(NroQuest).NumNPC(2) & ".jpg")
        lblCantNPC(1).Caption = MuerteNPC(2) & "/" & InfoQuests(NroQuest).CantNPC(2)
        
        picNPC(2).Picture = LoadPicture(App.Path & "\Data\INIT\Miniaturas\" & InfoQuests(NroQuest).NumNPC(3) & ".jpg")
        lblCantNPC(2).Caption = MuerteNPC(3) & "/" & InfoQuests(NroQuest).CantNPC(3)
    End If
    
    imgQuestAbandonar.Visible = True
    lblPTS.Caption = PonerPuntos(InfoQuests(NroQuest).puntos)
    lblOro.Caption = PonerPuntos(InfoQuests(NroQuest).Oro)
    lblPTS.Visible = True
    lblOro.Visible = True
    
    If InfoQuests(NroQuest).IndexOBJ > 0 Then
        Dim SR As RECT
        SR.left = 0
        SR.top = 0
        SR.Right = 32
        SR.bottom = 32
        
        picOBJ.Visible = True
        lblCantOBJ.Visible = True
    
        picOBJ.Refresh
        Call engine.DrawGrhtoHdc(InfoQuests(NroQuest).IndexOBJ, SR, picOBJ)
        lblCantOBJ.Caption = "x" & InfoQuests(NroQuest).CantOBJ
    End If
End Sub
Private Sub imgQuests_Click()

    SendData "KEST"
Me.Picture = General_Load_Interface_Picture("Estadisticas_2_Main.jpg")

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblClase.Visible = False
lblGenero.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

txtQuestDescription.Visible = True

'Vaciamos toda la pagina "extra" y "habilidades"
For i = 0 To 25
lblCounters(i).Visible = False
Next
imgHoja(1).Visible = False
imgHoja(2).Visible = False

End Sub
Private Sub imgExtras_Click()
Me.Picture = General_Load_Interface_Picture("Estadisticas_4_1_Main.jpg")
imgHoja(2).Visible = True
imgHoja(1).Visible = False

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblGenero.Visible = False
lblClase.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Mostramos primera hoja de la pagina "extra" y vaciamos "habilidades"
For i = 0 To 14
lblCounters(i).Visible = True
Next
For i = 15 To 25
lblCounters(i).Visible = False
Next

'Vaciamos la pagina "quest"
    lblPTS.Visible = False
    lblOro.Visible = False
    txtQuestDescription.Visible = False
    imgQuestAbandonar.Visible = False
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    picOBJ.Visible = False
    lblCantOBJ.Visible = False
End Sub
Private Sub imgHoja_Click(Index As Integer)
If Index = 1 Then
Me.Picture = General_Load_Interface_Picture("Estadisticas_4_1_Main.jpg")
imgHoja(2).Visible = True
imgHoja(1).Visible = False

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblGenero.Visible = False
lblClase.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblReputacion.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Mostramos primera hoja de la pagina "extra" y vaciamos "habilidades"
For i = 0 To 14
lblCounters(i).Visible = True
Next
For i = 15 To 25
lblCounters(i).Visible = False
Next

'Vaciamos la pagina "quest"
    lblPTS.Visible = False
    lblOro.Visible = False
    txtQuestDescription.Visible = False
    imgQuestAbandonar.Visible = False
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    picOBJ.Visible = False
    lblCantOBJ.Visible = False

ElseIf Index = 2 Then
Me.Picture = General_Load_Interface_Picture("Estadisticas_4_2_Main.jpg")
imgHoja(2).Visible = False
imgHoja(1).Visible = True

'Vaciamos toda la primera pagina.
lblNombre.Visible = False
lblLvl.Visible = False
lblRaza.Visible = False
lblGenero.Visible = False
lblClase.Visible = False
lblHogar.Visible = False
lblMail.Visible = False
lblPuntosTorneo.Visible = False
lblPuntosDonador.Visible = False
lblBonificadores(1).Visible = False
lblBonificadores(2).Visible = False
lblBonificadores(3).Visible = False
For i = 1 To NUMATRIBUTOS
    lblAtri(i).Visible = False
Next

'Mostramos primera hoja de la pagina "extra" y vaciamos "habilidades"
For i = 0 To 14
lblCounters(i).Visible = False
Next
For i = 15 To 25
lblCounters(i).Visible = True
Next

'Vaciamos la pagina "quest"
    lblPTS.Visible = False
    lblOro.Visible = False
    txtQuestDescription.Visible = False
    imgQuestAbandonar.Visible = False
    picNPC(0).Visible = False
    lblCantNPC(0).Visible = False
    picNPC(1).Visible = False
    lblCantNPC(1).Visible = False
    picNPC(2).Visible = False
    lblCantNPC(2).Visible = False
    picOBJ.Visible = False
    lblCantOBJ.Visible = False
End If

End Sub
Private Sub imgQuestAbandonar_Click()
Call SendData("/NOQUEST")
End Sub
Private Sub imgGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGeneral.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Principal_I.jpg")
End Sub
Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGeneral.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Principal_A.jpg")
End Sub
Private Sub imgQuests_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuests.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Quest_I.jpg")
End Sub
Private Sub imgQuests_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuests.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Quest_A.jpg")
End Sub
Private Sub imgExtras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExtras.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Extras_I.jpg")
End Sub
Private Sub imgExtras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExtras.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_1_Extras_A.jpg")
End Sub
Private Sub imgQuestAbandonar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuestAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_2_AbandonarQuest_I.jpg")
End Sub
Private Sub imgQuestAbandonar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgQuestAbandonar.Picture = LoadPicture(App.Path & "\Data\GRAFICOS\Principal\Estadisticas_2_AbandonarQuest_A.jpg")
End Sub
Private Sub lblPuntosDonador_Click()
frmMercadoTS.Timer1.Enabled = True
Call SendData("DCANJE")
End Sub
Private Sub lblPuntosTorneo_Click()
Call SendData("CCANJE")
End Sub
