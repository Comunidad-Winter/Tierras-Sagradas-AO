VERSION 5.00
Begin VB.Form frmCasas 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Propiedades"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   3045
      TabIndex        =   9
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informacion de la casa"
      Height          =   2895
      Left            =   3190
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Comprar"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2295
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1905
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1305
         Width           =   3375
      End
      Begin VB.TextBox text 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   615
         Width           =   3375
      End
      Begin VB.Label lblPrecio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   1600
         Width           =   3255
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblDueño 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dueño"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.ListBox ListaCasas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   0
      Top             =   125
      Width           =   2775
   End
End
Attribute VB_Name = "frmCasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NumeroKSA As Byte
Private Sub Command1_Click()
If ListaCasas.text = "Casa Número 1" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 1
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 2" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 2
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 3" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 3
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 4" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 4
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 5" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 5
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 6" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 6
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 7" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 7
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 8" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 8
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 9" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 9
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 10" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 10
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 11" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 11
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 12" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 12
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 13" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 13
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 14" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 14
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 15" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 15
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 16" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 16
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 17" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 17
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 18" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 18
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 19" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 19
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 20" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 20
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 21" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 21
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 22" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 22
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If


If ListaCasas.text = "Casa Número 23" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 23
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If


If ListaCasas.text = "Casa Número 24" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 24
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If


If ListaCasas.text = "Casa Número 25" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 25
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If


If ListaCasas.text = "Casa Número 26" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 26
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If


If ListaCasas.text = "Casa Número 27" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 27
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If

If ListaCasas.text = "Casa Número 28" Then
If MsgBox("¿Está seguro que desea comprar está casa?", vbYesNo) = vbYes Then
NumeroKSA = 28
Call SendData("CUC" & NumeroKSA)
Unload Me
End If
End If


End Sub
Private Sub Form_Load()
ListaCasas.AddItem "Casa Número 1"
ListaCasas.AddItem "Casa Número 2"
ListaCasas.AddItem "Casa Número 3"
ListaCasas.AddItem "Casa Número 4"
ListaCasas.AddItem "Casa Número 5"
ListaCasas.AddItem "Casa Número 6"
ListaCasas.AddItem "Casa Número 7"
ListaCasas.AddItem "Casa Número 8"
ListaCasas.AddItem "Casa Número 9"
ListaCasas.AddItem "Casa Número 10"
ListaCasas.AddItem "Casa Número 11"
ListaCasas.AddItem "Casa Número 12"
ListaCasas.AddItem "Casa Número 13"
ListaCasas.AddItem "Casa Número 14"
ListaCasas.AddItem "Casa Número 15"
ListaCasas.AddItem "Casa Número 16"
ListaCasas.AddItem "Casa Número 17"
ListaCasas.AddItem "Casa Número 18"
ListaCasas.AddItem "Casa Número 19"
ListaCasas.AddItem "Casa Número 20"
ListaCasas.AddItem "Casa Número 21"
ListaCasas.AddItem "Casa Número 22"
ListaCasas.AddItem "Casa Número 23"
ListaCasas.AddItem "Casa Número 24"
ListaCasas.AddItem "Casa Número 25"
ListaCasas.AddItem "Casa Número 26"
ListaCasas.AddItem "Casa Número 27"
ListaCasas.AddItem "Casa Número 28"
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub ListaCasas_Click()
If ListaCasas.text = "Casa Número 1" Then
NumeroKSA = 1
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 2" Then
NumeroKSA = 2
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 3" Then
NumeroKSA = 3
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 4" Then
NumeroKSA = 4
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 5" Then
NumeroKSA = 5
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 6" Then
NumeroKSA = 6
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 7" Then
NumeroKSA = 7
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 8" Then
NumeroKSA = 8
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 9" Then
NumeroKSA = 9
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 10" Then
NumeroKSA = 10
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 11" Then
NumeroKSA = 11
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 12" Then
NumeroKSA = 12
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 13" Then
NumeroKSA = 13
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 14" Then
NumeroKSA = 14
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 15" Then
NumeroKSA = 15
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 16" Then
NumeroKSA = 16
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 17" Then
NumeroKSA = 17
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 18" Then
NumeroKSA = 18
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 19" Then
NumeroKSA = 19
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 20" Then
NumeroKSA = 20
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 21" Then
NumeroKSA = 21
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 20" Then
NumeroKSA = 20
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 22" Then
NumeroKSA = 22
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 23" Then
NumeroKSA = 23
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 24" Then
NumeroKSA = 24
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 25" Then
NumeroKSA = 25
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 26" Then
NumeroKSA = 26
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 27" Then
NumeroKSA = 27
Call SendData("FWO" & NumeroKSA)
End If

If ListaCasas.text = "Casa Número 28" Then
NumeroKSA = 28
Call SendData("FWO" & NumeroKSA)
End If

End Sub
