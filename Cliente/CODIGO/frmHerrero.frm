VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   4500
   ClientLeft      =   5655
   ClientTop       =   3465
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHerrero.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1695
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmHerrero.frx":6137
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   280
      TabIndex        =   0
      Top             =   265
      Width           =   2670
   End
   Begin VB.ListBox lstArmaduras 
      Height          =   2205
      Left            =   280
      TabIndex        =   1
      Top             =   265
      Width           =   2670
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":6150
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   240
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":6D3D
      Top             =   3240
      Width           =   2715
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   1680
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":75F0
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   240
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":7C77
      Top             =   2520
      Width           =   1485
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub Image2_Click()

On Error Resume Next

If lstArmas.Visible Then
 Call SendData("CNS" & ArmasHerrero(lstArmas.ListIndex))
Else
 Call SendData("CNS" & ArmadurasHerrero(lstArmaduras.ListIndex))
End If
End Sub
Private Sub Image3_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub
Private Sub Image4_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub
