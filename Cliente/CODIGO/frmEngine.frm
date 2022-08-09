VERSION 5.00
Begin VB.Form frmEngine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Engine TSAO"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Crear Aura"
      Height          =   975
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "Crear Aura"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox Aurix 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Text            =   "23"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Numero del Aura:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Crearla sobre el Personaje"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Particulas"
      Height          =   2415
      Left            =   3720
      TabIndex        =   15
      Top             =   240
      Width           =   3255
      Begin VB.TextBox Time 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Namber 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Text            =   "0"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Crearla en el Mapa"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Crear Particula"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Tiempo:"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Numero:"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crear Luces"
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox Range 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Text            =   "3"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Blue 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Text            =   "255"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Green 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "255"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox red 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Text            =   "255"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear Luz"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1250
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Range:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Blue:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Geen:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Red:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox erre 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "255"
      Top             =   480
      Width           =   420
   End
   Begin VB.TextBox ge 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "255"
      Top             =   480
      Width           =   420
   End
   Begin VB.TextBox be 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "255"
      Top             =   480
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   490
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Luz del Render"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    day_r_old = erre.text
    day_g_old = ge.text
    day_b_old = be.text
    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
End Sub
Private Sub Command2_Click()
Light.Create_Light_To_Map UserPos.X, UserPos.Y, Range.text, red.text, Green.text, Blue.text
End Sub
Private Sub Command3_Click()

If Check1.value = 0 And Check2.value = 0 Then MsgBox "Elegí si la queres sobre el mapa o sobre el personaje down."

If Check1.value = 1 Then
    Call General_Particle_Create(Namber.text, UserPos.X, UserPos.Y, Time.text)
End If

If Check2.value = 1 Then
    Call SendData("/MOD PART " & Aurix.text)
End If


End Sub
Private Sub Command4_Click()
    Call SendData("/MOD AURA " & Aurix.text)
End Sub
